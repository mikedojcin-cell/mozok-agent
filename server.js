const express = require('express');
const https = require('https');
const path = require('path');
const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const SUPABASE_URL = 'https://kwyycykglqrokqsbuiny.supabase.co';
const SUPABASE_KEY = process.env.SUPABASE_KEY;
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imt3eXljeWtnbHFyb2txc2J1aW55Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg2Nzk5MTksImV4cCI6MjA5NDI1NTkxOX0.pw2ej0U3xgd-i-1Xub_jWq7iI5RZC7N4f43jFSc9DfM';
const META_APP_ID = process.env.META_APP_ID;
const META_APP_SECRET = process.env.META_APP_SCRET
const META_REDIRECT = 'https://app.mozok.co/auth/meta/callback';
const BASE_URL = 'https://app.mozok.co';

// ── Send pacing state (in-memory, single-instance) ──────────────────────────
// Root cause of the "dozens of identical-template emails in a row, ~30-45s
// apart" pattern flagged in Mozok_Cold_Email_Sequence_v2.md: nothing server-
// side ever enforced a minimum gap between sends or a hard daily cap. The
// daily_send_goal setting existed but was never actually checked before
// sending. Fixed below — every send now waits out a randomized minimum gap
// and is blocked once today's cap is hit, regardless of how fast the caller
// (CRM loop, double-click, etc.) fires requests.
let _lastSendAtMs = 0;
const _sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));



// ─── AUTH MIDDLEWARE ────────────────────────────────────────────────────────────

async function requireAuth(req, res, next) {
  const auth = req.headers['authorization'];
  if (!auth || !auth.startsWith('Bearer ')) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  const token = auth.replace('Bearer ', '');
  try {
    const result = await new Promise((resolve, reject) => {
      const r = https.request({
        hostname: new URL(SUPABASE_URL).hostname,
        path: '/auth/v1/user',
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${token}`,
          'apikey': SUPABASE_ANON_KEY
        }
      }, res => {
        let d = '';
        res.on('data', c => d += c);
        res.on('end', () => { try { resolve({ status: res.statusCode, data: JSON.parse(d) }); } catch(e) { reject(e); } });
      });
      r.on('error', reject);
      r.end();
    });
    if (result.status === 200 && result.data.id) {
      req.user = result.data;
      next();
    } else {
      res.status(401).json({ error: 'Invalid or expired token' });
    }
  } catch(e) {
    res.status(401).json({ error: 'Auth check failed' });
  }
}

// Protect API routes (tracking + contact form stay public)
app.use('/api/pipeline', requireAuth);
app.use('/api/graph', requireAuth);
app.use('/api/meta', requireAuth);
app.use('/api/generate-post', requireAuth);

// ─── HELPERS ────────────────────────────────────────────────────────────────────

async function supabase(method, endpoint, body) {
  return new Promise((resolve, reject) => {
    const data = body ? JSON.stringify(body) : null;
    const url = new URL(SUPABASE_URL + endpoint);
    const req = https.request({
      hostname: url.hostname,
      path: url.pathname + url.search,
      method,
      headers: {
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`,
        'Content-Type': 'application/json',
        'Prefer': method === 'POST' ? 'return=representation' : '',
        ...(data ? { 'Content-Length': Buffer.byteLength(data) } : {})
      }
    }, res => {
      let d = '';
      res.on('data', c => d += c);
      res.on('end', () => { try { resolve(JSON.parse(d)); } catch(e) { resolve([]); } });
    });
    req.on('error', reject);
    if (data) req.write(data);
    req.end();
  });
}

async function getToken(tenantId, clientId, clientSecret) {
  return new Promise((resolve, reject) => {
    const body = `grant_type=client_credentials&client_id=${clientId}&client_secret=${encodeURIComponent(clientSecret)}&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default`;
    const req = https.request({
      hostname: 'login.microsoftonline.com',
      path: `/${tenantId}/oauth2/v2.0/token`,
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': body.length }
    }, res => {
      let d = '';
      res.on('data', c => d += c);
      res.on('end', () => { try { resolve(JSON.parse(d)); } catch(e) { reject(e); } });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

async function graphCall(token, method, urlPath, body, extraHeaders) {
  return new Promise((resolve, reject) => {
    const data = body ? JSON.stringify(body) : null;
    const req = https.request({
      hostname: 'graph.microsoft.com',
      path: urlPath,
      method,
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        ...(extraHeaders || {}),
        ...(data ? { 'Content-Length': Buffer.byteLength(data) } : {})
      }
    }, res => {
      let d = '';
      res.on('data', c => d += c);
      res.on('end', () => {
        try { resolve({ status: res.statusCode, data: d ? JSON.parse(d) : {} }); }
        catch(e) { resolve({ status: res.statusCode, data: d }); }
      });
    });
    req.on('error', reject);
    if (data) req.write(data);
    req.end();
  });
}

async function metaGet(url) {
  return new Promise((resolve, reject) => {
    https.get(url, res => {
      let d = '';
      res.on('data', c => d += c);
      res.on('end', () => { try { resolve(JSON.parse(d)); } catch(e) { resolve({}); } });
    }).on('error', reject);
  });
}

// ─── EMAIL CLICK TRACKING (public) ──────────────────────────────────────────────

app.get('/track/click/:contactId', async (req, res) => {
  const { contactId } = req.params;
  const redirectUrl = req.query.url || 'https://mozok.co';
  try {
    await supabase('POST', '/rest/v1/email_events', {
      contact_id: contactId,
      event_type: 'click',
      metadata: redirectUrl
    });
    let contacts = await supabase('GET', `/rest/v1/contacts?id=eq.${contactId}&select=click_count`);
    const currentCount = contacts[0]?.click_count || 0;
    await supabase('PATCH', `/rest/v1/contacts?id=eq.${contactId}`, {
      click_count: currentCount + 1,
      last_clicked_at: new Date().toISOString()
    });
  } catch(e) {
    console.error('Track click error:', e.message);
  }
  res.redirect(redirectUrl);
});

// ─── CLEAN VISIT REDIRECT (public) ──────────────────────────────────────────────

app.get('/visit/:contactId', async (req, res) => {
  const { contactId } = req.params;
  const redirectUrl = 'https://mozok.co';
  try {
    await supabase('POST', '/rest/v1/email_events', {
      contact_id: contactId,
      event_type: 'click',
      metadata: redirectUrl
    });
    const contacts = await supabase('GET', `/rest/v1/contacts?id=eq.${contactId}&select=click_count`);
    const currentCount = contacts[0]?.click_count || 0;
    await supabase('PATCH', `/rest/v1/contacts?id=eq.${contactId}`, {
      click_count: currentCount + 1,
      last_clicked_at: new Date().toISOString()
    });
  } catch(e) {
    console.error('Visit track error:', e.message);
  }
  res.redirect(redirectUrl);
});

// ─── EMAIL OPEN TRACKING (public) ───────────────────────────────────────────────

const PIXEL = Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64');

// NOTE: GoogleImageProxy, YahooMailProxy, "MS Exchange" and "Microsoft Office" were
// REMOVED from this list on 2026-07-25. Those strings are what Gmail/Yahoo/Outlook's
// own image-proxy infrastructure sends when a REAL HUMAN opens the email — proxying
// images through the provider's own servers (to hide the recipient's IP/UA) is now
// standard privacy behavior at Google, Yahoo, and Microsoft. Blocking those patterns
// was silently discarding real opens from anyone on Gmail or Outlook/O365 — i.e. most
// of a B2B prospect list. Left in place: dedicated corporate security-scanner and
// generic-HTTP-library signatures, which are unambiguous automated fetches.
const BOT_UA_PATTERNS = ['Synapse','msnbot','bingbot','curl','python-requests','okhttp','Go-http-client','Googlebot','Google-Read-Aloud','Cloudflare','Barracuda','Proofpoint','Mimecast','IronPort','SpamAssassin','Symantec','Trend Micro','Wget','libwww','Jakarta','Java/','Apache-HttpClient'];
function isBotUA(ua){if(!ua||ua.trim()==='')return true;const l=ua.toLowerCase();return BOT_UA_PATTERNS.some(p=>l.includes(p.toLowerCase()));}
app.get('/track/open/:contactId', async (req, res) => {
  const { contactId } = req.params;
  const ua = req.headers['user-agent'] || '';
  const isBot = isBotUA(ua);
  try {
    await supabase('POST', '/rest/v1/email_events', { contact_id: contactId, event_type: isBot ? 'bot_open' : 'open', metadata: ua });
    if (!isBot) {
      const contacts = await supabase('GET', `/rest/v1/contacts?id=eq.${contactId}&select=open_count`);
      const currentCount = contacts[0]?.open_count || 0;
      await supabase('PATCH', `/rest/v1/contacts?id=eq.${contactId}`, { open_count: currentCount + 1, last_opened_at: new Date().toISOString() });
    }
  } catch(e) { console.error('Track open error:', e.message); }
  const PIXEL = Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64');
  res.writeHead(200, { 'Content-Type': 'image/gif', 'Content-Length': PIXEL.length, 'Cache-Control': 'no-store, no-cache, must-revalidate, private', 'Pragma': 'no-cache' });
  res.end(PIXEL);
});

// ─── EMAIL PIPELINE API (protected) ─────────────────────────────────────────────

app.get('/api/pipeline', async (req, res) => {
  try {
    const contacts = await supabase('GET', '/rest/v1/contacts?email=neq.&status=neq.REMOVED&status=neq.TEST&order=created_at.desc&limit=500&select=id,firstname,lastname,email,company,status,email1_sent_at,email2_sent_at,email3_sent_at,last_opened_at,open_count,click_count,last_clicked_at');
    const now = Date.now();
    const pipeline = { ready:[], email1_sent:[], email2_due:[], email2_sent:[], email3_due:[], email3_sent:[], clicked:[], bounced:[] };
    if (!Array.isArray(contacts)) return res.json({ error: contacts?.message || 'Supabase error', pipeline: {}, total: 0, batch_stats: [], sent_today: 0 });

    for (const c of contacts) {
      const hasClicked=(c.click_count||0)>0;
      const e1sent=c.email1_sent_at?new Date(c.email1_sent_at):null;
      const e2sent=c.email2_sent_at?new Date(c.email2_sent_at):null;
      const e1days=e1sent?(now-e1sent)/86400000:null;
      const e2days=e2sent?(now-e2sent)/86400000:null;
      c.human_opened=(c.open_count||0)>0;
      if(c.status==='BOUNCED')pipeline.bounced.push(c);
      else if(hasClicked)pipeline.clicked.push(c);
      else if(c.status==='EMAIL_3_SENT')pipeline.email3_sent.push(c);
      else if(c.status==='EMAIL_2_SENT'&&e2days>=7)pipeline.email3_due.push(c);
      else if(c.status==='EMAIL_2_SENT')pipeline.email2_sent.push(c);
      else if(c.status==='EMAIL_1_SENT'&&e1days>=7)pipeline.email2_due.push(c);
      else if(c.status==='EMAIL_1_SENT')pipeline.email1_sent.push(c);
else if (c.human_opened) pipeline.email1_sent.push(c);
            else pipeline.ready.push(c);
    }
    const batchMap={};
    for(const c of contacts){if(!c.batch_id)continue;if(!batchMap[c.batch_id])batchMap[c.batch_id]={batch_id:c.batch_id,count:0,date:c.batch_id.split('_')[0]};batchMap[c.batch_id].count++;}
    const batch_stats=Object.values(batchMap).sort((a,b)=>b.date.localeCompare(a.date)).slice(0,14);
    const todayStr=new Date().toISOString().slice(0,10);
    const sent_today=contacts.filter(c=>(c.email1_sent_at||'').startsWith(todayStr)||(c.email2_sent_at||'').startsWith(todayStr)||(c.email3_sent_at||'').startsWith(todayStr)).length;
    res.json({ pipeline, total:contacts.length, batch_stats, sent_today });
  } catch(e){res.json({error:e.message});}
});

// ─── TRACKING TEST SEND (protected) ─────────────────────────────────────────────
// Lets you send a real tracked test email to any address you control (e.g. a
// Gmail inbox and an Outlook/O365 inbox) without touching real prospect data,
// so you can verify the open-tracking pixel actually fires per mail provider.
// Test contacts are tagged status:'TEST' and excluded from /api/pipeline above.
// Added 2026-07-25 — see public/test-send.html for the UI.
//
// Credentials come from env vars, NOT a form field. First attempt had the user
// retype tenantId/clientId/clientSecret into a browser form on every visit —
// Chrome's password manager kept hijacking those fields and autofilling real
// saved account passwords into them (ignores autocomplete="off" by design).
// Since these values are static per Microsoft app registration anyway, they
// belong server-side, same pattern as SUPABASE_KEY/ANTHROPIC_KEY below. Set
// TEST_SEND_TENANT_ID, TEST_SEND_CLIENT_ID, TEST_SEND_CLIENT_SECRET,
// TEST_SEND_USER_EMAIL in Render's Environment tab (same values you use in
// the main app's Connect form) — fixed 2026-07-25.

app.post('/api/test-email', requireAuth, async (req, res) => {
  const { to } = req.body;
  const tenantId = process.env.TEST_SEND_TENANT_ID;
  const clientId = process.env.TEST_SEND_CLIENT_ID;
  const clientSecret = process.env.TEST_SEND_CLIENT_SECRET;
  const userEmail = process.env.TEST_SEND_USER_EMAIL;
  if (!tenantId || !clientId || !clientSecret || !userEmail) {
    return res.json({ error: 'Test-send credentials not configured. Set TEST_SEND_TENANT_ID, TEST_SEND_CLIENT_ID, TEST_SEND_CLIENT_SECRET, and TEST_SEND_USER_EMAIL in Render → Environment (same values as the main app Connect form).' });
  }
  if (!to) {
    return res.json({ error: 'Missing test recipient email' });
  }
  try {
    const created = await supabase('POST', '/rest/v1/contacts', {
      email: to,
      firstname: 'Test',
      lastname: 'Send',
      company: 'Test',
      status: 'TEST',
      created_at: new Date().toISOString()
    });
    const contactId = created[0]?.id;
    if (!contactId) return res.json({ error: 'Could not create test contact: ' + JSON.stringify(created) });

    const tokenRes = await getToken(tenantId, clientId, clientSecret);
    if (!tokenRes.access_token) return res.json({ error: '[' + (tokenRes.error || 'token_error') + '] ' + (tokenRes.error_description || 'Token failed') });
    const token = tokenRes.access_token;

    const trackingPixel = `\n\n<img src="${BASE_URL}/track/open/${contactId}" width="1" height="1" style="display:none" />`;
    const htmlBody = `<p>This is a Mozok tracking test email, sent ${new Date().toLocaleString()}.</p><p>Open this from the actual inbox (not a preview pane) to test the tracking pixel.</p>` + trackingPixel;

    const msgBody = {
      message: {
        subject: `Mozok tracking test — ${new Date().toLocaleTimeString()}`,
        body: { contentType: 'html', content: htmlBody },
        toRecipients: [{ emailAddress: { address: to } }],
        from: { emailAddress: { address: userEmail } }
      },
      saveToSentItems: true
    };
    const r = await graphCall(token, 'POST', `/v1.0/users/${userEmail}/sendMail`, msgBody);
    if (r.status === 202) {
      await supabase('PATCH', `/rest/v1/contacts?id=eq.${contactId}`, { email1_sent_at: new Date().toISOString() }).catch(()=>{});
      return res.json({ success: true, contactId, to });
    }
    return res.json({ error: (r.data?.error?.code ? '[' + r.data.error.code + '] ' : '') + (r.data?.error?.message || `Send failed (${r.status})`) });
  } catch (e) {
    return res.json({ error: e.message });
  }
});

app.get('/api/test-status/:contactId', requireAuth, async (req, res) => {
  const { contactId } = req.params;
  try {
    const [contacts, events] = await Promise.all([
      supabase('GET', `/rest/v1/contacts?id=eq.${contactId}&select=email,open_count,click_count,last_opened_at,last_clicked_at,status`),
      supabase('GET', `/rest/v1/email_events?contact_id=eq.${contactId}&order=created_at.desc&limit=20&select=event_type,metadata,created_at`)
    ]);
    res.json({ contact: contacts[0] || null, events: Array.isArray(events) ? events : [] });
  } catch (e) {
    res.json({ error: e.message });
  }
});

// ─── GRAPH / SUPABASE API (protected) ───────────────────────────────────────────

app.post('/api/graph', async (req, res) => {
  const { action } = req.body;
  // mozok-crm.html's send buttons don't have a Connect form and never had
  // any way to supply Microsoft credentials — they were calling a
  // nonexistent route (/api/action) so this never got exercised until that
  // was fixed. Fall back to the same server-side credentials the test-send
  // route already uses (mirrors the values hardcoded in index.html's
  // Connect form) whenever the caller doesn't supply its own.
  const tenantId = req.body.tenantId || process.env.TEST_SEND_TENANT_ID;
  const clientId = req.body.clientId || process.env.TEST_SEND_CLIENT_ID;
  const clientSecret = req.body.clientSecret || process.env.TEST_SEND_CLIENT_SECRET;
  const userEmail = req.body.userEmail || process.env.TEST_SEND_USER_EMAIL || 'mike@mozok.co';

  if (action === 'getContacts') {
    const { status, offset = 0 } = req.body;  // ← offset added for pagination
    try {
      const data = await supabase('GET', `/rest/v1/contacts?status=eq.${status}&email=neq.&order=created_at.desc&limit=100&offset=${offset}`);
      return res.json({ results: Array.isArray(data) ? data : [], total: Array.isArray(data) ? data.length : 0 });
    } catch(e) { return res.json({ error: e.message }); }
  }

  if (action === 'updateContact') {
    const { contactId, props } = req.body;
    try {
      await supabase('PATCH', `/rest/v1/contacts?id=eq.${contactId}`, props);
      return res.json({ success: true });
    } catch(e) { return res.json({ error: e.message }); }
  }

  if (action === 'addContacts') {
    const { newContacts } = req.body;
    try {
      const result = await supabase('POST', '/rest/v1/contacts', newContacts);
      return res.json({ success: true, added: Array.isArray(result) ? result.length : 0 });
    } catch(e) { return res.json({ error: e.message }); }
  }

  if (action === 'getStats') {
    try {
      const statuses = ['NEW', 'EMAIL_1_SENT', 'EMAIL_2_SENT', 'EMAIL_3_SENT'];
      const stats = {};
      for (const s of statuses) {
        const data = await supabase('GET', `/rest/v1/contacts?status=eq.${s}&select=id`);
        stats[s] = Array.isArray(data) ? data.length : 0;
      }
      return res.json(stats);
    } catch(e) { return res.json({ error: e.message }); }
  }

  if (action === 'sendEmail') {
    let { to, body: emailBody } = req.body;
    const { contactId } = req.body;
    const emailNum = req.body.emailNum || 1;
    if (!contactId) console.warn(`[sendEmail] WARNING: sending to ${to} with no contactId — this email will have NO open/click tracking pixel and cannot be attributed to a contact.`);

    // ── Server-side idempotency guard ─────────────────────────────────────
    // Root cause of duplicate sends: this endpoint used to trust the client's
    // stale in-memory contact list to decide who was "ready". Two overlapping
    // batch clicks (or a retry after a slow response) would both compute the
    // same "ready" set and both call sendEmail for the same contact+emailNum.
    // We now re-check the CURRENT DB state right before sending, and refuse
    // to send again if this emailNum already has a sent timestamp.
    if (contactId) {
      const _guardRows = await supabase('GET', `/rest/v1/contacts?id=eq.${contactId}&select=email${emailNum}_sent_at`).catch(() => []);
      const _guard = (_guardRows || [])[0];
      if (_guard && _guard[`email${emailNum}_sent_at`]) {
        return res.json({ error: `Skipped: Email ${emailNum} was already sent to this contact at ${_guard[`email${emailNum}_sent_at`]}. Refusing duplicate send.`, duplicate: true });
      }
    }

    const [_sRows, _cRows] = await Promise.all([
      supabase('GET', '/rest/v1/campaign_settings?id=eq.1&select=email_subject_1,email_subject_2,email_subject_3,daily_send_goal'),
      contactId ? supabase('GET', `/rest/v1/contacts?id=eq.${contactId}&select=firstname,company,email`) : Promise.resolve([])
    ]);
    const _s = _sRows[0] || {}, _c = _cRows[0] || {};

    // ── Canonical email templates (server-authoritative) ──────────────────
    // Root cause of "it errored out" sending from mozok-crm.html: that page's
    // bulk-send button only ever posted { action, contactId, emailNum } — no
    // `to`, no `body`. This handler used to trust the caller to supply both,
    // so `emailBody.replace(...)` below threw on undefined for every single
    // contact in that batch. index.html's sendEmails() built subject/body
    // itself and always passed them, which is why sends from THAT page kept
    // working while mozok-crm.html's kept failing — two UIs, two different
    // behaviors, both silently drifting. Fixed by making the server the one
    // source of truth: it now derives subject/body/recipient itself whenever
    // the caller doesn't supply them, instead of assuming they're present.
    const _firstName = _c.firstname || 'there';
    const _company = _c.company || 'your business';
    const _templates = {
      1: {
        subject: `${_company} — a quick one`,
        body: `Hi ${_firstName},\n\nMost businesses I talk to know they should be posting consistently, but it never happens – no one owns it day to day.\n\nI run Mozok. I'll build ${_company} 3 sample social posts this week, free, using your actual business – no call, no pitch, just posts in your inbox by Friday so you can see if it's worth anything.\n\nWant them? Just reply "yes."\n\nMike\n248-800-3405`
      },
      2: {
        subject: `Re: ${_company} — a quick one`,
        body: `Hi ${_firstName},\n\nFollowing up in case this got buried. One client's page used to be a brochure - the occasional "we're hiring" post and not much else. We took it over and started posting what was actually happening: project milestones, team celebrations, staff anniversaries, the community work they were already doing. Within weeks, customers and suppliers were mentioning it to them unprompted, and job candidates said it was part of why they wanted to work there. Nothing changed about the business - people just finally saw it. Worth a quick look?\n\nIf social isn't a priority right now, no worries - just let me know and I'll stop following up.\n\nMike`
      },
      3: {
        subject: `One idea for ${_company}`,
        body: `Hi ${_firstName},\n\nLast thought, then I'll leave it – if you ever want a done-for-you posting schedule (3x/week, no contract, $497/mo), we're here: mozok.co\n\nEither way, wishing ${_company} a good rest of the year.\n\nMike`
      }
    };
    const _template = _templates[emailNum] || _templates[1];

    const subject = (_s[`email_subject_${emailNum}`] || req.body.subject || _template.subject)
      .replace(/\{\{firstname\}\}/gi, _c.firstname || '')
      .replace(/\{\{company\}\}/gi, _c.company || '')
      .trim();

    // If the caller didn't supply a body (mozok-crm.html's case) or a `to`
    // address, fall back to the canonical template / the contact's stored
    // email instead of crashing on undefined further down.
    if (!emailBody) emailBody = _template.body;
    if (!to && _c.email) to = _c.email;
    if (!to) return res.json({ error: 'No recipient email address available for this contact.' });

    // ── Hard daily send cap ──────────────────────────────────────────────
    // Cold-email deliverability data (see Mozok_Cold_Email_Sequence_v2.md)
    // recommends ~40-50 sends/day per mailbox. daily_send_goal defaults to
    // 100 and was never enforced — just a number the UI displayed. Enforce
    // it here, counting anything stamped as sent today across all 3 touches.
    const _dailyCap = _s.daily_send_goal || 45;
    const _todayStartIso = new Date(new Date().toISOString().slice(0, 10) + 'T00:00:00.000Z').toISOString();
    const _todayCountRows = await supabase(
      'GET',
      `/rest/v1/contacts?select=id&or=(email1_sent_at.gte.${_todayStartIso},email2_sent_at.gte.${_todayStartIso},email3_sent_at.gte.${_todayStartIso})`
    ).catch(() => null);
    if (Array.isArray(_todayCountRows) && _todayCountRows.length >= _dailyCap) {
      return res.json({ error: `Daily send cap reached (${_todayCountRows.length}/${_dailyCap} sent today). Sending more today risks spam flags — resume tomorrow, or raise daily_send_goal in Settings if you're sure.`, capReached: true });
    }

    // ── Minimum gap between sends ────────────────────────────────────────
    // Enforces real wall-clock pacing regardless of how fast the caller
    // loops. Randomized 20-40s so the pattern doesn't look like a bot firing
    // on a fixed interval, which is its own spam-filter signal.
    const _gapMs = 20000 + Math.floor(Math.random() * 20000);
    const _sinceLastMs = Date.now() - _lastSendAtMs;
    if (_sinceLastMs < _gapMs) {
      await _sleep(_gapMs - _sinceLastMs);
    }
    _lastSendAtMs = Date.now();

    try {
      const tokenRes = await getToken(tenantId, clientId, clientSecret);
      if (!tokenRes.access_token) return res.json({ error: '[' + (tokenRes.error||'token_error') + '] ' + (tokenRes.error_description || 'Token failed') });
      const token = tokenRes.access_token;

      const trackingPixel = contactId
        ? `\n\n<img src="${BASE_URL}/track/open/${contactId}" width="1" height="1" style="display:none" />`
        : '';

      const htmlBody = emailBody
        .replace(/\n/g, '<br>')
        .replace(
          /(https:\/\/app\.mozok\.co\/visit\/[^\s<]+)/g,
          '<a href="$1" style="color:#1D9E75;">mozok.co</a>'
        ) + trackingPixel;

      const msgBody = {
        message: {
          subject,
          body: { contentType: 'html', content: htmlBody },
          toRecipients: [{ emailAddress: { address: to } }],
          from: { emailAddress: { address: userEmail } }
        },
        saveToSentItems: true
      };
      const r = await graphCall(token, 'POST', `/v1.0/users/${userEmail}/sendMail`, msgBody);
      if(r.status===202){
        const batchDate=new Date().toISOString().slice(0,10);
        const batchId=req.body.batch_id||batchDate+'_1';
        let stampFailed = false;
        if(contactId){
          const stampBody = {[`email${emailNum}_sent_at`]:new Date().toISOString(),status:`EMAIL_${emailNum}_SENT`,batch_id:batchId};
          // The email is already sent at this point — a failed stamp must NOT
          // be silently swallowed, because that leaves the contact looking
          // "never emailed" and the next batch run will re-send to them.
          // Retry once, then surface it loudly if it still fails so the UI
          // can warn Mike instead of quietly re-spamming the contact later.
          try {
            await supabase('PATCH', `/rest/v1/contacts?id=eq.${contactId}`, stampBody);
          } catch (e1) {
            console.error('[sendEmail] Batch stamp FAILED, retrying:', e1.message);
            try {
              await supabase('PATCH', `/rest/v1/contacts?id=eq.${contactId}`, stampBody);
            } catch (e2) {
              stampFailed = true;
              console.error(`[sendEmail] Batch stamp FAILED TWICE for contact ${contactId} — email was sent but DB was not updated. This contact WILL look unsent and risks a duplicate send. Manual fix required.`, e2.message);
            }
          }
        }
        return res.json({success:true,status:202,batch_id:batchId, stampFailed});
      }
      return res.json({ error: (r.data?.error?.code ? '[' + r.data.error.code + '] ' : '') + (r.data?.error?.message || `Send failed (${r.status})`) });
    } catch(e) { return res.json({ error: e.message }); }
  }

  try {
    const tokenRes = await getToken(tenantId, clientId, clientSecret);
    if (!tokenRes.access_token) return res.json({ error: '[' + (tokenRes.error||'token_error') + '] ' + (tokenRes.error_description || 'Token failed') });
    const token = tokenRes.access_token;

    if (action === 'test') {
      const r = await graphCall(token, 'GET', `/v1.0/users/${userEmail}`);
      if (r.status === 200) return res.json({ success: true, name: r.data.displayName, email: r.data.mail || userEmail });
      return res.json({ error: r.data?.error?.message || 'Auth failed' });
    }

    if (action === 'getListId') {
      const r = await graphCall(token, 'GET', `/v1.0/users/${userEmail}/todo/lists`);
      const list = r.data?.value?.find(l => l.displayName === 'Tasks') || r.data?.value?.[0];
      return res.json({ listId: list?.id || '' });
    }

    if (action === 'createTask') {
      const { taskData } = req.body;
      const body = {
        title: taskData.title,
        importance: taskData.priority === 'high' ? 'high' : taskData.priority === 'low' ? 'low' : 'normal',
        body: { content: taskData.notes || '', contentType: 'text' },
        ...(taskData.due ? { dueDateTime: { dateTime: `${taskData.due}T09:00:00`, timeZone: 'America/Toronto' } } : {})
      };
      const r = await graphCall(token, 'POST', `/v1.0/users/${userEmail}/todo/lists/${taskData.listId}/tasks`, body);
      if (r.status === 201) return res.json({ success: true, id: r.data.id });
      return res.json({ error: r.data?.error?.message || 'Task creation failed' });
    }

    if (action === 'createEvent') {
      const { eventData } = req.body;
      const body = {
        subject: eventData.subject,
        body: { contentType: 'text', content: eventData.body || '' },
        start: { dateTime: eventData.start, timeZone: 'America/Toronto' },
        end: { dateTime: eventData.end, timeZone: 'America/Toronto' },
        categories: [eventData.category]
      };
      const r = await graphCall(token, 'POST', `/v1.0/users/${userEmail}/events`, body);
      if (r.status === 201) return res.json({ success: true });
      return res.json({ error: r.data?.error?.message || 'Event creation failed' });
    }

    return res.json({ error: 'Unknown action' });
  } catch (e) {
    return res.json({ error: e.message });
  }
});

// ─── META OAUTH ─────────────────────────────────────────────────────────────────

app.get('/auth/meta', (req, res) => {
  const { clientId } = req.query;
  const scope = 'pages_show_list,business_management';
  const url = `https://www.facebook.com/v19.0/dialog/oauth?client_id=${META_APP_ID}&redirect_uri=${encodeURIComponent(META_REDIRECT)}&scope=${scope}&state=${clientId}`;
  res.redirect(url);
});

app.get('/auth/meta/callback', async (req, res) => {
  const { code, state: clientId } = req.query;
  if (!code) return res.send('Error: no code returned from Meta');
  try {
    const tokenUrl = `https://graph.facebook.com/v19.0/oauth/access_token?client_id=${META_APP_ID}&redirect_uri=${encodeURIComponent(META_REDIRECT)}&client_secret=${META_APP_SECRET}&code=${code}`;
    const tokenRes = await metaGet(tokenUrl);
    if (!tokenRes.access_token) return res.send('Error getting access token: ' + JSON.stringify(tokenRes));
    await supabase('PATCH', `/rest/v1/clients?id=eq.${clientId}`, { meta_access_token: tokenRes.access_token });
    res.redirect('/dashboard.html?connected=true');
  } catch(e) {
    res.send('Error: ' + e.message);
  }
});

// ─── META INSIGHTS (protected) ──────────────────────────────────────────────────

app.get('/api/meta/pages/:clientId', async (req, res) => {
  const { clientId } = req.params;
  try {
    const rows = await supabase('GET', `/rest/v1/clients?id=eq.${clientId}&select=meta_access_token`);
    const token = rows[0]?.meta_access_token;
    if (!token) return res.json({ error: 'No Meta token – client needs to connect Facebook' });
    const pages = await metaGet(`https://graph.facebook.com/v19.0/me/accounts?access_token=${token}`);
    res.json({ pages: pages.data || [] });
  } catch(e) {
    res.json({ error: e.message });
  }
});

app.get('/api/meta/insights/:clientId', async (req, res) => {
  const { clientId } = req.params;
  try {
    const rows = await supabase('GET', `/rest/v1/clients?id=eq.${clientId}&select=meta_access_token`);
    const token = rows[0]?.meta_access_token;
    if (!token) return res.json({ error: 'No Meta token – client needs to connect Facebook' });
    const pages = await metaGet(`https://graph.facebook.com/v19.0/me/accounts?access_token=${token}`);
    if (!pages.data || !pages.data.length) return res.json({ error: 'No pages found' });
    const page = pages.data[0];
    const pageToken = page.access_token;
    const pageId = page.id;
    const [insights, posts] = await Promise.all([
      metaGet(`https://graph.facebook.com/v19.0/${pageId}/insights?metric=page_impressions,page_reach,page_engaged_users&period=month&access_token=${pageToken}`),
      metaGet(`https://graph.facebook.com/v19.0/${pageId}/posts?fields=message,created_time,insights.metric(post_impressions,post_engaged_users)&limit=10&access_token=${pageToken}`)
    ]);
    res.json({ page: { id: pageId, name: page.name }, insights: insights.data || [], posts: posts.data || [] });
  } catch(e) {
    res.json({ error: e.message });
  }
});

// ─── CONTACT FORM (public) ───────────────────────────────────────────────────────

app.post('/api/contact', async (req, res) => {
  const { name, email, message } = req.body;
  if (!name || !email || !message) return res.json({ error: 'Missing fields' });
  try {
    const tokenRes = await getToken(
      '26b2d778-d390-4c96-9c8b-96cf1bf43c5d',
      'c141d1d2-14a9-4a2a-b965-aca133a394f4',
      process.env.MS_CLIENT_SECRET || 'kcP8Q~bQ1w49F4JiVUE3HJCiOYWqiw0FFJMp0ayl'
    );
    if (!tokenRes.access_token) return res.json({ error: 'Auth failed' });
    const msgBody = {
      message: {
        subject: `New onboarding message from ${name}`,
        body: {
          contentType: 'html',
          content: `<div style="font-family:Arial,sans-serif;max-width:500px;"><h2 style="color:#1D9E75;">New message from Mozok onboarding</h2><p><strong>Name:</strong> ${name}</p><p><strong>Email:</strong> <a href="mailto:${email}">${email}</a></p><p><strong>Message:</strong></p><p style="background:#f5f5f5;padding:12px;border-radius:8px;">${message.replace(/\n/g,'<br>')}</p></div>`
        },
        toRecipients: [{ emailAddress: { address: 'info@mozok.co' } }],
        from: { emailAddress: { address: 'mike@mozok.co' } },
        replyTo: [{ emailAddress: { address: email, name } }]
      },
      saveToSentItems: true
    };
    const r = await graphCall(tokenRes.access_token, 'POST', '/v1.0/users/mike@mozok.co/sendMail', msgBody);
    if (r.status === 202) return res.json({ success: true });
    return res.json({ error: 'Send failed' });
  } catch(e) {
    return res.json({ error: e.message });
  }
});

// ─── CLAUDE POST GENERATOR (protected) ──────────────────────────────────────────

app.post('/api/generate-post', async (req, res) => {
  const { prompt } = req.body;
  if (!prompt) return res.json({ error: 'No prompt provided' });
  const ANTHROPIC_KEY = process.env.ANTHROPIC_KEY;
  if (!ANTHROPIC_KEY) return res.json({ error: 'ANTHROPIC_KEY not set in environment' });
  try {
    const data = JSON.stringify({
      model: 'claude-sonnet-4-5',
      max_tokens: 1000,
      messages: [{ role: 'user', content: prompt }]
    });
    const result = await new Promise((resolve, reject) => {
      const apiReq = https.request({
        hostname: 'api.anthropic.com',
        path: '/v1/messages',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': ANTHROPIC_KEY,
          'anthropic-version': '2023-06-01',
          'Content-Length': Buffer.byteLength(data)
        }
      }, r => {
        let d = '';
        r.on('data', c => d += c);
        r.on('end', () => { try { resolve(JSON.parse(d)); } catch(e) { reject(e); } });
      });
      apiReq.on('error', reject);
      apiReq.write(data);
      apiReq.end();
    });
    const text = result.content?.[0]?.text || '';
    if (!text) return res.json({ error: result.error?.message || 'No text returned' });
    res.json({ text });
  } catch(e) {
    res.json({ error: e.message });
  }
});

// Cheap, no-auth, no-DB endpoint for external uptime pings — hitting this keeps the
// Render free-tier dyno warm so tracking-pixel requests don't get dropped by the
// recipient's mail client while the server is cold-starting (see 2026-07-25 fix notes).
// MUST be registered before the '*' catch-all below, or the catch-all swallows it first.
app.get('/health', (req, res) => res.status(200).send('ok'));

// CAMPAIGN SETTINGS
// MUST be registered before the '*' catch-all below (same bug as /health above) —
// GET /api/campaign-settings was being silently swallowed by the catch-all and
// returning index.html's HTML instead of JSON. settings.html's fetch().json()
// would throw parsing '<script>...' as JSON, so the page always looked blank
// even after a successful save. Moved here 2026-07-25.
app.get('/api/campaign-settings', requireAuth, async (req, res) => {
  try{const rows=await supabase('GET','/rest/v1/campaign_settings?id=eq.1&select=*');res.json(rows[0]||{});}catch(e){res.json({error:e.message});}
});
app.post('/api/campaign-settings', requireAuth, async (req, res) => {
  const{target_location,industry,job_titles,company_size_min,company_size_max,daily_send_goal,email_subject_1,email_subject_2,email_subject_3}=req.body;
  // Fixed 2026-07-25: this used to POST a raw row with a hardcoded id:1 and no
  // real upsert (the 'Prefer' header passed here was silently dropped — the
  // supabase() helper only forwards method/endpoint/body). Second save onward
  // either errored on a duplicate id or silently created extra rows depending
  // on schema, and any field left out of the request body could get written
  // as null instead of being left alone. Now: PATCH the existing row if one
  // exists (partial update — untouched fields stay untouched), else create it.
  // The supabase() helper never throws on a rejected write — it resolves
  // whatever JSON body Supabase/PostgREST sent back, even for 400/403/etc, so
  // a bad column name or permission issue would silently report success:true
  // here. Now checking the actual response shape before claiming success.
  // 2026-07-25: campaign_settings.job_titles is a Postgres array column —
  // sending it as a delimited string (even an empty string) triggers "malformed
  // array literal" from PostgREST. Convert to a real array (or null when empty).
  // target_location, by contrast, IS plain text — the array-literal error we hit
  // earlier came specifically from job_titles being blank, not target_location.
  // Sending target_location as an array too (an earlier version of this fix did)
  // silently "succeeded" but stored the JSON-stringified array as literal text,
  // e.g. '["Ontario, Canada"]' with visible brackets/quotes — caught by testing
  // an actual save + reload, not by the error response, since PostgREST didn't
  // reject it. Confirmed via GET /api/campaign-settings: target_location comes
  // back as typeof 'string', job_titles as a real array.
  function toArrayOrNull(str, sep) {
    if (typeof str !== 'string' || !str.trim()) return null;
    const arr = str.split(sep).map(s => s.trim()).filter(Boolean);
    return arr.length ? arr : null;
  }
  function toTextOrNull(str) {
    if (typeof str !== 'string' || !str.trim()) return null;
    return str.trim();
  }
  try{
    const payload = {target_location: toTextOrNull(target_location), industry, job_titles: toArrayOrNull(job_titles, ','), company_size_min, company_size_max, daily_send_goal: daily_send_goal||100, email_subject_1, email_subject_2, email_subject_3, updated_at: new Date().toISOString()};
    const existing = await supabase('GET','/rest/v1/campaign_settings?id=eq.1&select=id');
    if (!Array.isArray(existing)) {
      return res.json({ error: 'Supabase read failed: ' + JSON.stringify(existing).slice(0, 300) });
    }
    let result;
    if (existing.length > 0) {
      result = await supabase('PATCH','/rest/v1/campaign_settings?id=eq.1', payload);
    } else {
      result = await supabase('POST','/rest/v1/campaign_settings', {id:1, ...payload});
    }
    // A successful PATCH/POST resolves to an array (rows) or [] (no
    // representation requested). A PostgREST error resolves to a plain object
    // with .code/.message/.hint instead.
    if (result && !Array.isArray(result) && (result.code || result.message)) {
      return res.json({ error: 'Supabase write failed: ' + (result.message || result.code) + (result.hint ? ' — ' + result.hint : '') });
    }
    const verify = await supabase('GET','/rest/v1/campaign_settings?id=eq.1&select=target_location,job_titles');
    res.json({success:true, saved: verify[0] || null});
  }catch(e){res.json({error:e.message});}
});

// ─── CATCH ALL ───────────────────────────────────────────────────────────────────

app.get('/login.html', (req, res) => res.sendFile(path.join(__dirname, 'public', 'login.html')));
app.get('/onboarding.html', (req, res) => res.sendFile(path.join(__dirname, 'public', 'onboarding.html')));
app.get('/dashboard.html', (req, res) => res.sendFile(path.join(__dirname, 'public', 'dashboard.html')));
app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

const PORT = process.env.PORT || 3000;
// BOUNCE CHECK
app.post('/api/bounce-check', requireAuth, async (req, res) => {
  const { tenantId, clientId, clientSecret, userEmail } = req.body;
  if(!tenantId||!clientId||!clientSecret||!userEmail)return res.json({error:'Missing credentials'});
  try {
    const tokenRes=await getToken(tenantId,clientId,clientSecret);
    if(!tokenRes.access_token)return res.json({error:'Token failed'});
    const token=tokenRes.access_token;
    const filter=encodeURIComponent("contains(subject,'Undeliverable') or contains(subject,'Delivery has failed') or contains(subject,'Mail Delivery Subsystem')");
    // NOTE (fixed 2026-08-01): contains() filters on message properties require the
    // ConsistencyLevel:eventual header + $count=true, or Graph returns a 400 that this
    // code was previously swallowing as a silent "checked:0" success. That's why bounce
    // detection has been reporting 0 checked even with real NDRs sitting in the inbox.
    const ndrRes=await graphCall(token,'GET',`/v1.0/users/${userEmail}/messages?$filter=${filter}&$top=50&$count=true&$select=subject,bodyPreview`,null,{'ConsistencyLevel':'eventual'});
    if(!ndrRes?.data?.value)return res.json({bounced:[],checked:0,debug_error:ndrRes?.data?.error?.message||('Graph returned status '+ndrRes?.status+' with no value array — likely a permissions or ConsistencyLevel issue, not "no bounces found."')});
    const bouncedEmails=[];
    for(const msg of ndrRes.data.value){const m=((msg.subject||'')+(msg.bodyPreview||'')).match(/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g)||[];m.forEach(e=>{if(!e.includes('postmaster')&&!e.includes('mailer-daemon')&&e.toLowerCase()!==userEmail.toLowerCase())bouncedEmails.push(e.toLowerCase());});}
    const unique=[...new Set(bouncedEmails)];
    let marked=0;
    for(const email of unique){const ex=await supabase('GET',`/rest/v1/contacts?email=eq.${encodeURIComponent(email)}&select=id,status`);if(ex.length>0&&ex[0].status!=='BOUNCED'){await supabase('PATCH',`/rest/v1/contacts?email=eq.${encodeURIComponent(email)}`,{status:'BOUNCED',bounced_at:new Date().toISOString()});marked++;}}
    res.json({bounced:unique,checked:ndrRes.data.value.length,marked});
  }catch(e){res.json({error:e.message});}
});

app.listen(PORT, () => console.log(`Mozok Agent running on port ${PORT}`));
