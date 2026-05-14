const express = require('express');
const https = require('https');
const path = require('path');
const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const SUPABASE_URL = 'https://kwyycykglqrokqsbuiny.supabase.co';
const SUPABASE_KEY = process.env.SUPABASE_KEY;
const META_APP_ID = process.env.META_APP_ID;
const META_APP_SECRET = process.env.META_APP_SECRET;
const META_REDIRECT = 'https://mozok-agent.onrender.com/auth/meta/callback';
const BASE_URL = 'https://mozok-agent.onrender.com';

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

// ─── EMAIL OPEN TRACKING ──────────────────────────────────────────────────────

// 1x1 transparent GIF
const PIXEL = Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64');

app.get('/track/open/:contactId', async (req, res) => {
  const { contactId } = req.params;
  try {
    // Log the open event
    await supabase('POST', '/rest/v1/email_events', {
      contact_id: contactId,
      event_type: 'open',
      metadata: req.headers['user-agent'] || ''
    });
    // Update contact open count and last opened
    await supabase('PATCH', `/rest/v1/contacts?id=eq.${contactId}`, {
      last_opened_at: new Date().toISOString(),
      open_count: 1 // Supabase will handle increment via RPC but this sets it
    });
    // Increment open count separately
    const contacts = await supabase('GET', `/rest/v1/contacts?id=eq.${contactId}&select=open_count`);
    const currentCount = contacts[0]?.open_count || 0;
    await supabase('PATCH', `/rest/v1/contacts?id=eq.${contactId}`, {
      open_count: currentCount + 1,
      last_opened_at: new Date().toISOString()
    });
  } catch(e) {
    console.error('Track open error:', e.message);
  }
  // Always return the pixel regardless
  res.writeHead(200, {
    'Content-Type': 'image/gif',
    'Content-Length': PIXEL.length,
    'Cache-Control': 'no-store, no-cache, must-revalidate, private',
    'Pragma': 'no-cache'
  });
  res.end(PIXEL);
});

// ─── EMAIL PIPELINE API ───────────────────────────────────────────────────────

app.get('/api/pipeline', async (req, res) => {
  try {
    const [contacts, events] = await Promise.all([
      supabase('GET', '/rest/v1/contacts?email=neq.&status=neq.REMOVED&order=created_at.desc&limit=500&select=id,firstname,lastname,email,company,status,email1_sent_at,email2_sent_at,email3_sent_at,last_opened_at,open_count'),
      supabase('GET', '/rest/v1/email_events?order=created_at.desc&limit=200')
    ]);

    const now = Date.now();
    const pipeline = {
      ready: [],
      email1_sent: [],
      email2_due: [],
      email2_sent: [],
      email3_due: [],
      email3_sent: [],
      opened: []
    };

    for (const c of contacts) {
      const hasOpened = c.open_count > 0;
      const e1sent = c.email1_sent_at ? new Date(c.email1_sent_at) : null;
      const e2sent = c.email2_sent_at ? new Date(c.email2_sent_at) : null;
      const e1days = e1sent ? (now - e1sent) / 86400000 : null;
      const e2days = e2sent ? (now - e2sent) / 86400000 : null;

      if (hasOpened) pipeline.opened.push(c);
      else if (c.status === 'EMAIL_3_SENT') pipeline.email3_sent.push(c);
      else if (c.status === 'EMAIL_2_SENT' && e2days >= 5) pipeline.email3_due.push(c);
      else if (c.status === 'EMAIL_2_SENT') pipeline.email2_sent.push(c);
      else if (c.status === 'EMAIL_1_SENT' && e1days >= 5) pipeline.email2_due.push(c);
      else if (c.status === 'EMAIL_1_SENT') pipeline.email1_sent.push(c);
      else pipeline.ready.push(c);
    }

    res.json({ pipeline, total: contacts.length });
  } catch(e) {
    res.json({ error: e.message });
  }
});

// ─── SUPABASE CONTACT ROUTES ──────────────────────────────────────────────────

app.post('/api/graph', async (req, res) => {
  const { tenantId, clientId, clientSecret, userEmail, action } = req.body;

  if (action === 'getContacts') {
    const { status } = req.body;
    try {
      const data = await supabase('GET', `/rest/v1/contacts?status=eq.${status}&email=neq.&order=created_at.desc&limit=100`);
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
    const { to, subject, body: emailBody, contactId } = req.body;
    try {
      const tokenRes = await getToken(tenantId, clientId, clientSecret);
      if (!tokenRes.access_token) return res.json({ error: tokenRes.error_description || 'Token failed' });
      const token = tokenRes.access_token;

      // Add tracking pixel to email body
      const trackingPixel = contactId
        ? `\n\n<img src="${BASE_URL}/track/open/${contactId}" width="1" height="1" style="display:none" />`
        : '';

      const htmlBody = emailBody.replace(/\n/g, '<br>') + trackingPixel;

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
      if (r.status === 202) return res.json({ success: true, status: 202 });
      return res.json({ error: r.data?.error?.message || `Send failed (${r.status})` });
    } catch(e) { return res.json({ error: e.message }); }
  }

  // ─── MICROSOFT GRAPH ────────────────────────────────────────────────────────
  try {
    const tokenRes = await getToken(tenantId, clientId, clientSecret);
    if (!tokenRes.access_token) return res.json({ error: tokenRes.error_description || 'Token failed' });
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

// ─── META OAUTH ───────────────────────────────────────────────────────────────

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
    await supabase('PATCH', `/rest/v1/clients?id=eq.${clientId}`, {
      meta_access_token: tokenRes.access_token
    });
    res.redirect('/dashboard.html?connected=true');
  } catch(e) {
    res.send('Error: ' + e.message);
  }
});

// ─── META INSIGHTS ────────────────────────────────────────────────────────────

app.get('/api/meta/pages/:clientId', async (req, res) => {
  const { clientId } = req.params;
  try {
    const rows = await supabase('GET', `/rest/v1/clients?id=eq.${clientId}&select=meta_access_token`);
    const token = rows[0]?.meta_access_token;
    if (!token) return res.json({ error: 'No Meta token — client needs to connect Facebook' });
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
    if (!token) return res.json({ error: 'No Meta token — client needs to connect Facebook' });
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

// ─── CLAUDE POST GENERATOR ────────────────────────────────────────────────────

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
        r.on('end', () => { 
          console.log('Anthropic response status:', r.statusCode);
          console.log('Anthropic response body:', d.substring(0, 500));
          try { resolve(JSON.parse(d)); } catch(e) { reject(e); } 
        });
      });
      apiReq.on('error', reject);
      apiReq.write(data);
      apiReq.end();
    });

    const text = result.content?.[0]?.text || '';
    if (!text) {
      const errMsg = result.error?.message || JSON.stringify(result);
      console.log('No text returned:', errMsg);
      return res.json({ error: errMsg });
    }
    res.json({ text });
  } catch(e) {
    console.log('Generate post error:', e.message);
    res.json({ error: e.message });
  }
});

// ─── CATCH ALL ────────────────────────────────────────────────────────────────

app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Mozok Agent running on port ${PORT}`));
