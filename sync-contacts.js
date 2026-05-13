const https = require('https');

const APOLLO_KEY = process.env.APOLLO_KEY;
const SUPABASE_KEY = process.env.SUPABASE_KEY;
const SUPABASE_URL = 'https://your-project.supabase.co'; // ← You need to update this

function apolloRequest(body) {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(body);
    const options = {
      hostname: 'api.apollo.io',
      path: '/v1/mixed_people/search',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': 'no-cache',
        'X-Api-Key': APOLLO_KEY,
      }
    };
    const req = https.request(options, (res) => {
      let body = '';
      res.on('data', chunk => body += chunk);
      res.on('end', () => {
        try { resolve(JSON.parse(body)); }
        catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.write(data);
    req.end();
  });
}

function supabaseUpsert(contacts) {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(contacts);
    const url = new URL(`${SUPABASE_URL}/rest/v1/contacts`);
    const options = {
      hostname: url.hostname,
      path: url.pathname + '?on_conflict=apollo_id',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`,
        'Prefer': 'resolution=merge-duplicates'
      }
    };
    const req = https.request(options, (res) => {
      let body = '';
      res.on('data', chunk => body += chunk);
      res.on('end', () => {
        console.log(`Supabase response: ${res.statusCode}`);
        resolve(body);
      });
    });
    req.on('error', reject);
    req.write(data);
    req.end();
  });
}

async function syncContacts() {
  console.log('Starting Apollo → Supabase sync...');

  try {
    // Pull latest contacts from Apollo (distribution companies, Canada/US)
    const apolloData = await apolloRequest({
      q_organization_industry_tag_ids: ['5567cd4773696439b10b0000'], // Industrial/Manufacturing
      q_not_email_status: ['bounced'],
      person_titles: ['owner', 'president', 'ceo', 'purchasing manager', 'operations manager'],
      page: 1,
      per_page: 50
    });

    if (!apolloData.people || apolloData.people.length === 0) {
      console.log('No contacts returned from Apollo.');
      return;
    }

    console.log(`Fetched ${apolloData.people.length} contacts from Apollo.`);

    // Map Apollo fields to Supabase schema
    const contacts = apolloData.people.map(p => ({
      apollo_id: p.id,
      first_name: p.first_name || '',
      last_name: p.last_name || '',
      email: p.email || '',
      title: p.title || '',
      company: p.organization ? p.organization.name : '',
      phone: p.phone_numbers && p.phone_numbers[0] ? p.phone_numbers[0].raw_number : '',
      linkedin: p.linkedin_url || '',
      city: p.city || '',
      state: p.state || '',
      country: p.country || '',
      last_synced: new Date().toISOString()
    }));

    // Upsert into Supabase
    await supabaseUpsert(contacts);
    console.log(`✅ Synced ${contacts.length} contacts to Supabase.`);

  } catch (err) {
    console.error('❌ Sync failed:', err.message);
    process.exit(1);
  }
}

syncContacts();
