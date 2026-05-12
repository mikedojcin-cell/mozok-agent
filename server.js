const express = require('express');
const https = require('https');
const path = require('path');
const fs = require('fs');
const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const CONTACTS_FILE = path.join(__dirname, 'contacts.json');

function loadContacts() {
  try {
    return JSON.parse(fs.readFileSync(CONTACTS_FILE, 'utf8'));
  } catch(e) {
    return [];
  }
}

function saveContacts(contacts) {
  fs.writeFileSync(CONTACTS_FILE, JSON.stringify(contacts, null, 2));
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

app.post('/api/graph', async (req, res) => {
  const { tenantId, clientId, clientSecret, userEmail, action } = req.body;

  // Contact actions — no Microsoft token needed
  if (action === 'getContacts') {
    const { status } = req.body;
    const contacts = loadContacts().filter(c => c.status === status);
    return res.json({ results: contacts, total: contacts.length });
  }

  if (action === 'updateContact') {
    const { contactId, props } = req.body;
    const contacts = loadContacts();
    const idx = contacts.findIndex(c => c.id === contactId);
    if (idx >= 0) {
      contacts[idx] = { ...contacts[idx], ...props };
      saveContacts(contacts);
    }
    return res.json({ success: true });
  }

  if (action === 'addContacts') {
    const { newContacts } = req.body;
    const contacts = loadContacts();
    const existing = new Set(contacts.map(c => c.email));
    const toAdd = newContacts.filter(c => !existing.has(c.email));
    saveContacts([...contacts, ...toAdd]);
    return res.json({ success: true, added: toAdd.length });
  }

  // Microsoft actions
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

    if (action === 'sendEmail') {
      const { to, subject, body: emailBody } = req.body;
      const msgBody = {
        message: {
          subject,
          body: { contentType: 'text', content: emailBody },
          toRecipients: [{ emailAddress: { address: to } }],
          from: { emailAddress: { address: userEmail } }
        },
        saveToSentItems: true
      };
      const r = await graphCall(token, 'POST', `/v1.0/users/${userEmail}/sendMail`, msgBody);
      if (r.status === 202) return res.json({ success: true, status: 202 });
      return res.json({ error: r.data?.error?.message || `Send failed (${r.status})` });
    }

    return res.json({ error: 'Unknown action' });
  } catch (e) {
    return res.json({ error: e.message });
  }
});

app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Mozok Agent running on port ${PORT}`));
