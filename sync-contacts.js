const https = require('https');

const APOLLO_KEY = process.env.APOLLO_KEY;
const SUPABASE_KEY = process.env.SUPABASE_KEY;
const SUPABASE_URL = 'https://kwyycykglqrokqsbuiny.supabase.co';

function apolloRequest(body) {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(body);
    const options = {
      hostname: 'api.apollo.io',
      path: '/v1/mixed_people/api_search',
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

function supabaseRequest(method, path, body) {
  return new Promise((resolve, reject) => {
    const data = body ? JSON.stringify(body) : null;
    const url = new URL(`${SUPABASE_URL}${path}`);
    const options = {
      hostname: url.hostname,
      path: url.pathname + url.search,
      method,
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`,
        'Prefer': method === 'POST' ? 'resolution=merge-duplicates' : 'return=representation'
      }
    };
    const req = https.request(options, (res) => {
      let body = '';
      res.on('data', chunk => body += chunk);
      res.on('end', () => {
        try { resolve({ status: res.statusCode, body: JSON.parse(body) }); }
        catch (e) { resolve({ status: res.statusCode, body }); }
      });
    });
    req.on('error', reject);
    if (data) req.write(data);
    req.end();
  });
}

// Location + page configs — cycles through these day by day
// Uses person_locations (where the person is) not organization_locations
const SEARCH_CONFIGS = [
  { locations: ['Windsor, Ontario, Canada'], titles: ['owner', 'president', 'ceo', 'founder', 'managing director'] },
  { locations: ['Windsor, Ontario, Canada'], titles: ['purchasing manager', 'operations manager', 'marketing manager', 'general manager'] },
  { locations: ['Detroit, Michigan, United States'], titles: ['owner', 'president', 'ceo', 'founder'] },
  { locations: ['Detroit, Michigan, United States'], titles: ['marketing manager', 'marketing director', 'operations manager'] },
  { locations: ['Ann Arbor, Michigan, United States'], titles: ['owner', 'president', 'ceo', 'founder'] },
  { locations: ['Troy, Michigan, United States', 'Sterling Heights, Michigan, United States'], titles: ['owner', 'president', 'ceo'] },
  { locations: ['Dearborn, Michigan, United States', 'Livonia, Michigan, United States'], titles: ['owner', 'president', 'ceo'] },
  { locations: ['Lansing, Michigan, United States'], titles: ['owner', 'president', 'ceo', 'founder'] },
  { locations: ['London, Ontario, Canada', 'Chatham, Ontario, Canada'], titles: ['owner', 'president', 'ceo', 'founder'] },
  { locations: ['Toledo, Ohio, United States'], titles: ['owner', 'president', 'ceo', 'founder'] },
];

async function getPageTracker() {
  const res = await supabaseRequest('GET', '/rest/v1/sync_state?key=eq.apollo_page&select=value');
  if (res.status === 200 && res.body && res.body.length > 0) {
    return JSON.parse(res.body[0].value);
  }
  return { config_index: 0, page: 1 };
}

async function savePageTracker(state) {
  await supabaseRequest('POST', '/rest/v1/sync_state?on_conflict=key', [
    { key: 'apollo_page', value: JSON.stringify(state) }
  ]);
}

async function syncContacts() {
  console.log('Starting Apollo → Supabase sync...');

  try {
    // Get current position
    let state = await getPageTracker();
    console.log(`Resuming from config_index=${state.config_index}, page=${state.page}`);

    const config = SEARCH_CONFIGS[state.config_index % SEARCH_CONFIGS.length];
    console.log(`Searching: ${config.locations.join(', ')} | Titles: ${config.titles.join(', ')} | Page: ${state.page}`);

    const apolloData = await apolloRequest({
      api_key: APOLLO_KEY,
      person_titles: config.titles,
      person_locations: config.locations,
      organization_num_employees_ranges: ['1,10', '11,50', '51,200'],
      contact_email_status: ['verified', 'likely to engage'],
      page: state.page,
      per_page: 100
    });

    if (!apolloData.people || apolloData.people.length === 0) {
      console.log(`Apollo response: ${JSON.stringify(apolloData).substring(0, 300)}`);
      console.log(`No contacts on page ${state.page} for this config. Moving to next location.`);
      // Advance to next config, reset page
      state = { config_index: (state.config_index + 1) % SEARCH_CONFIGS.length, page: 1 };
      await savePageTracker(state);
      console.log(`Next run will use config_index=${state.config_index}`);
      return;
    }

    console.log(`Fetched ${apolloData.people.length} contacts from Apollo.`);

    const contacts = apolloData.people.map(p => ({
      apollo_id: p.id,
      firstname: p.first_name || '',
      lastname: p.last_name || '',
      email: p.email || '',
      company: p.organization ? p.organization.name : '',
      phone: p.phone_numbers && p.phone_numbers[0] ? p.phone_numbers[0].raw_number : '',
      linkedin: p.linkedin_url || '',
      city: p.city || '',
      state: p.state || '',
      country: p.country || '',
      status: 'NEW',
      last_synced: new Date().toISOString()
    }));

    const result = await supabaseRequest('POST', '/rest/v1/contacts?on_conflict=apollo_id', contacts);
    console.log(`Supabase status: ${result.status}`);
    console.log(`Supabase body: ${JSON.stringify(result.body).substring(0, 500)}`);
    if (result.status === 200 || result.status === 201) {
      console.log(`✅ Synced ${contacts.length} contacts to Supabase.`);
    } else {
      console.log(`❌ Supabase insert failed - see body above for reason.`);
    }

    // Advance page for next run
    state.page += 1;
    await savePageTracker(state);
    console.log(`Next run will fetch page ${state.page} of the same config.`);

  } catch (err) {
    console.error('❌ Sync failed:', err.message);
    process.exit(1);
  }
}

syncContacts();
