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

async function graphCall(token, method, urlPath, body) {
  return new Promise((resolve, reject) => {
    const data = body ? JSON.stringify(body) : null;
    const req = https.request({
      hostname: 'graph.microsoft.com',
      path: urlPath,
      method,
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
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
      else if(c.status==='EMAIL_2_SENT'&&e2days>=5)pipeline.email3_due.push(c);
      else if(c.status==='EMAIL_2_SENT')pipeline.email2_sent.push(c);
      else if(c.status==='EMAIL_1_SENT'&&e1days>=5)pipeline.email2_due.push(c);
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
  const { tenantId, clientId, clientSecret, userEmail, action } = req.body;

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
    const { to, body: emailBody, contactId } = req.body;
    const emailNum = req.body.emailNum || 1;
    if (!contactId) console.warn(`[sendEmail] WARNING: sending to ${to} with no contactId — this email will have NO open/click tracking pixel and cannot be attributed to a contact.`);
    const [_sRows, _cRows] = await Promise.all([
      supabase('GET', '/rest/v1/campaign_settings?id=eq.1&select=email_subject_1,email_subject_2,email_subject_3'),
      contactId ? supabase('GET', `/rest/v1/contacts?id=eq.${contactId}&select=firstname,company`) : Promise.resolve([])
    ]);
    const _s = _sRows[0] || {}, _c = _cRows[0] || {};
    const subject = (_s[`email_subject_${emailNum}`] || req.body.subject || 'Following up')
      .replace(/\{\{firstname\}\}/gi, _c.firstname || '')
      .replace(/\{\{company\}\}/gi, _c.company || '')
      .trim();
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
      if(r.status===202){const batchDate=new Date().toISOString().slice(0,10);const batchId=req.body.batch_id||batchDate+'_1';const emailNum=req.body.emailNum||1;if(contactId){await supabase('PATCH',`/rest/v1/contacts?id=eq.${contactId}`,{[`email${emailNum}_sent_at`]:new Date().toISOString(),status:`EMAIL_${emailNum}_SENT`,batch_id:batchId}).catch(e=>console.error('Batch stamp:',e.message));}return res.json({success:true,status:202,batch_id:batchId});}
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
    const ndrRes=await graphCall(token,'GET',`/v1.0/users/${userEmail}/messages?$filter=${filter}&$top=50&$select=subject,bodyPreview`,null);
    if(!ndrRes?.data?.value)return res.json({bounced:[],checked:0});
    const bouncedEmails=[];
    for(const msg of ndrRes.data.value){const m=((msg.subject||'')+(msg.bodyPreview||'')).match(/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g)||[];m.forEach(e=>{if(!e.includes('postmaster')&&!e.includes('mailer-daemon')&&e.toLowerCase()!==userEmail.toLowerCase())bouncedEmails.push(e.toLowerCase());});}
    const unique=[...new Set(bouncedEmails)];
    let marked=0;
    for(const email of unique){const ex=await supabase('GET',`/rest/v1/contacts?email=eq.${encodeURIComponent(email)}&select=id,status`);if(ex.length>0&&ex[0].status!=='BOUNCED'){await supabase('PATCH',`/rest/v1/contacts?email=eq.${encodeURIComponent(email)}`,{status:'BOUNCED',bounced_at:new Date().toISOString()});marked++;}}
    res.json({bounced:unique,checked:ndrRes.data.value.length,marked});
  }catch(e){res.json({error:e.message});}
});
// CAMPAIGN SETTINGS
app.get('/api/campaign-settings', requireAuth, async (req, res) => {
  try{const rows=await supabase('GET','/rest/v1/campaign_settings?id=eq.1&select=*');res.json(rows[0]||{});}catch(e){res.json({error:e.message});}
});
app.post('/api/campaign-settings', requireAuth, async (req, res) => {
  const{target_location,industry,job_titles,company_size_min,company_size_max,daily_send_goal,email_subject_1,email_subject_2,email_subject_3}=req.body;
  // The supabase() helper never throws on a rejected write — it resolves
  // whatever JSON body Supabase/PostgREST sent back, even for 400/403/etc, so
  // a bad column name or permission issue would silently report success:true
  // here. Now checking the actual response shape before claiming success.
  try{
    const payload = {target_location,industry,job_titles,company_size_min,company_size_max,daily_send_goal:daily_send_goal||100,email_subject_1,email_subject_2,email_subject_3,updated_at:new Date().toISOString()};
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

app.listen(PORT, () => console.log(`Mozok Agent running on port ${PORT}`));
