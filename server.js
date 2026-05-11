const express = require('express');
const fetch = require('node-fetch');
const app = express();
const path = require('path');
app.use(express.static('public'));
app.use(express.json());
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

app.post('/api/graph', async (req, res) => {
  const { action, tenantId, clientId, clientSecret, userEmail, taskData, eventData } = req.body;
  try {
    const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ grant_type: 'client_credentials', client_id: clientId, client_secret: clientSecret, scope: 'https://graph.microsoft.com/.default' })
    });
    const tokenData = await tokenRes.json();
    if (!tokenData.access_token) return res.status(401).json({ error: tokenData.error_description || 'Auth failed' });
    const token = tokenData.access_token;
    const gh = { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };

    if (action === 'test') {
      const r = await fetch(`https://graph.microsoft.com/v1.0/users/${userEmail}`, { headers: gh });
      const d = await r.json();
      return d.error ? res.status(400).json({ error: d.error.message }) : res.json({ success: true, name: d.displayName, email: d.mail });
    }
    if (action === 'getListId') {
      const r = await fetch(`https://graph.microsoft.com/v1.0/users/${userEmail}/todo/lists`, { headers: gh });
      const d = await r.json();
      const def = d.value?.find(l => l.wellknownListName === 'defaultList') || d.value?.[0];
      return res.json({ listId: def?.id });
    }
    if (action === 'createTask') {
      const { listId, title, notes, due, priority } = taskData;
      const body = { title, importance: priority || 'normal', body: { contentType: 'text', content: notes || '' } };
      if (due) body.dueDateTime = { dateTime: due + 'T09:00:00', timeZone: 'America/Toronto' };
      const r = await fetch(`https://graph.microsoft.com/v1.0/users/${userEmail}/todo/lists/${listId}/tasks`, { method: 'POST', headers: gh, body: JSON.stringify(body) });
      const d = await r.json();
      return d.error ? res.status(400).json({ error: d.error.message }) : res.json({ success: true, id: d.id });
    }
    if (action === 'createEvent') {
      const { title, start, end, notes, category } = eventData;
      const r = await fetch(`https://graph.microsoft.com/v1.0/users/${userEmail}/events`, { method: 'POST', headers: gh, body: JSON.stringify({ subject: title, body: { contentType: 'text', content: notes || '' }, start: { dateTime: start, timeZone: 'America/Toronto' }, end: { dateTime: end, timeZone: 'America/Toronto' }, categories: [category] }) });
      const d = await r.json();
      return d.error ? res.status(400).json({ error: d.error.message }) : res.json({ success: true });
    }
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/', (req, res) => res.send('Mozok Agent API running'));
app.listen(process.env.PORT || 3000, () => console.log('Running'));
