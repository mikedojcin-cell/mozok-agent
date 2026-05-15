const https = require('https');

const APOLLO_KEY = process.env.APOLLO_KEY;
const SUPABASE_KEY = process.env.SUPABASE_KEY;
const SUPABASE_URL = 'https://kwyycykglqrokqsbuiny.supabase.co';

function apolloRequest(path, body) {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(body);
    const req = https.request({
      hostname: 'api.apollo.io',
      path,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': 'no-cache',
        'X-Api-Key': APOLLO_KEY,
      }
    }, (res) => {
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

async function enrichBatch(people) {
  const details = people.map(p => ({
    id: p.id,
    first_name: p.first_name,
    last_name: p.last_name || '',
    organization_name: p.organization ? p.organization.name : ''
  }));
  try {
    const result = await apolloRequest('/v1/people/bulk_match', {
      api_key: APOLLO_KEY,
      details,
      reveal_personal_emails: false
    });
    return result.matches || [];
  } catch(e) {
    console.log(`Enrich batch error: ${e.message}`);
    return [];
  }
}

function supabaseRequest(method, path, body) {
  return new Promise((resolve, reject) => {
    const data = body ? JSON.stringify(body) : null;
    const url = new URL(`${SUPABASE_URL}${path}`);
    const req = https.request({
      hostname: url.hostname,
      path: url.pathname + url.search,
      method,
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`,
        'Prefer': method === 'POST' ? 'resolution=merge-duplicates' : 'return=representation'
      }
    }, (res) => {
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

const OWNER_TITLES = ['owner', 'president', 'ceo', 'founder', 'managing director', 'principal'];
const MANAGER_TITLES = ['marketing manager', 'marketing director', 'general manager', 'operations manager', 'business development manager'];

const SEARCH_CONFIGS = [
  // ── CANADA ──────────────────────────────────────────────────────────────────
  { label: 'Ontario — Owners',          locations: ['Ontario, Canada'],               titles: OWNER_TITLES },
  { label: 'Ontario — Managers',        locations: ['Ontario, Canada'],               titles: MANAGER_TITLES },
  { label: 'Quebec — Owners',           locations: ['Quebec, Canada'],                titles: OWNER_TITLES },
  { label: 'British Columbia — Owners', locations: ['British Columbia, Canada'],      titles: OWNER_TITLES },
  { label: 'Alberta — Owners',          locations: ['Alberta, Canada'],               titles: OWNER_TITLES },
  { label: 'Manitoba — Owners',         locations: ['Manitoba, Canada'],              titles: OWNER_TITLES },
  { label: 'Saskatchewan — Owners',     locations: ['Saskatchewan, Canada'],          titles: OWNER_TITLES },
  { label: 'Nova Scotia — Owners',      locations: ['Nova Scotia, Canada'],           titles: OWNER_TITLES },
  // ── MICHIGAN ────────────────────────────────────────────────────────────────
  { label: 'Michigan — Owners',         locations: ['Michigan, United States'],       titles: OWNER_TITLES },
  { label: 'Michigan — Managers',       locations: ['Michigan, United States'],       titles: MANAGER_TITLES },
  // ── OHIO ─────────────────────────────────────────────────────────────────────
  { label: 'Ohio — Owners',             locations: ['Ohio, United States'],           titles: OWNER_TITLES },
  { label: 'Ohio — Managers',           locations: ['Ohio, United States'],           titles: MANAGER_TITLES },
  // ── PENNSYLVANIA ─────────────────────────────────────────────────────────────
  { label: 'Pennsylvania — Owners',     locations: ['Pennsylvania, United States'],   titles: OWNER_TITLES },
  { label: 'Pennsylvania — Managers',   locations: ['Pennsylvania, United States'],   titles: MANAGER_TITLES },
  // ── INDIANA ──────────────────────────────────────────────────────────────────
  { label: 'Indiana — Owners',          locations: ['Indiana, United States'],        titles: OWNER_TITLES },
  { label: 'Indiana — Managers',        locations: ['Indiana, United States'],        titles: MANAGER_TITLES },
  // ── ILLINOIS ─────────────────────────────────────────────────────────────────
  { label: 'Illinois — Owners',         locations: ['Illinois, United States'],       titles: OWNER_TITLES },
  { label: 'Illinois — Managers',       locations: ['Illinois, United States'],       titles: MANAGER_TITLES },
  // ── WISCONSIN ────────────────────────────────────────────────────────────────
  { label: 'Wisconsin — Owners',        locations: ['Wisconsin, United States'],      titles: OWNER_TITLES },
  { label: 'Wisconsin — Managers',      locations: ['Wisconsin, United States'],      titles: MANAGER_TITLES },
  // ── NEW YORK ─────────────────────────────────────────────────────────────────
  { label: 'New York — Owners',         locations: ['New York, United States'],       titles: OWNER_TITLES },
  { label: 'New York — Managers',       locations: ['New York, United States'],       titles: MANAGER_TITLES },
  // ── FLORIDA ──────────────────────────────────────────────────────────────────
  { label: 'Florida — Owners',          locations: ['Florida, United States'],        titles: OWNER_TITLES },
  { label: 'Florida — Managers',        locations: ['Florida, United States'],        titles: MANAGER_TITLES },
  // ── TEXAS ────────────────────────────────────────────────────────────────────
  { label: 'Texas — Owners',            locations: ['Texas, United States'],          titles: OWNER_TITLES },
  { label: 'Texas — Managers',          locations: ['Texas, United States'],          titles: MANAGER_TITLES },
  // ── GEORGIA ──────────────────────────────────────────────────────────────────
  { label: 'Georgia — Owners',          locations: ['Georgia, United States'],        titles: OWNER_TITLES },
  // ── NORTH CAROLINA ───────────────────────────────────────────────────────────
  { label: 'North Carolina — Owners',   locations: ['North Carolina, United States'], titles: OWNER_TITLES },
  // ── VIRGINIA ─────────────────────────────────────────────────────────────────
  { label: 'Virginia — Owners',         locations: ['Virginia, United States'],       titles: OWNER_TITLES },
  // ── MINNESOTA ────────────────────────────────────────────────────────────────
  { label: 'Minnesota — Owners',        locations: ['Minnesota, United States'],      titles: OWNER_TITLES },
  // ── MISSOURI ─────────────────────────────────────────────────────────────────
  { label: 'Missouri — Owners',         locations: ['Missouri, United States'],       titles: OWNER_TITLES },
  // ── TENNESSEE ────────────────────────────────────────────────────────────────
  { label: 'Tennessee — Owners',        locations: ['Tennessee, United States'],      titles: OWNER_TITLES },
  // ── COLORADO ─────────────────────────────────────────────────────────────────
  { label: 'Colorado — Owners',         locations: ['Colorado, United States'],       titles: OWNER_TITLES },
  // ── WASHINGTON ───────────────────────────────────────────────────────────────
  { label: 'Washington — Owners',       locations: ['Washington, United States'],     titles: OWNER_TITLES },
  // ── CALIFORNIA ───────────────────────────────────────────────────────────────
  { label: 'California — Owners',       locations: ['California, United States'],     titles: OWNER_TITLES },
  { label: 'California — Managers',     locations: ['California, United States'],     titles: MANAGER_TITLES },
];

async function getState() {
  const res = await supabaseRequest('GET', '/rest/v1/sync_state?key=eq.apollo_page&select=value');
  if (res.status === 200 && res.body && res.body.length > 0) {
    return JSON.parse(res.body[0].value);
  }
  return { config_index: 0, page: 1, exhausted: [] };
}

async function saveState(state) {
  await supabaseRequest('POST', '/rest/v1/sync_state?on_conflict=key', [
    { key: 'apollo_page', value: JSON.stringify(state) }
  ]);
}

function findNextConfig(state) {
  const exhausted = state.exhausted || [];
  const total = SEARCH_CONFIGS.length;
  let idx = state.config_index % total;
  // Find next non-exhausted config
  for (let i = 0; i < total; i++) {
    const candidate = (idx + i) % total;
    if (!exhausted.includes(candidate)) return candidate;
  }
  // All exhausted — reset and start over
  console.log('🔄 All configs exhausted — resetting and starting over!');
  return 0;
}

async function syncContacts() {
  console.log('Starting Apollo → Supabase sync...');

  try {
    let state = await getState();
    if (!state.exhausted) state.exhausted = [];

    const total = SEARCH_CONFIGS.length;
    const activeCount = total - state.exhausted.length;
    console.log(`Active configs: ${activeCount}/${total} | Exhausted: ${state.exhausted.length}`);

    // If all exhausted — reset
    if (state.exhausted.length >= total) {
      console.log('All configs exhausted — resetting all.');
      state.exhausted = [];
      state.config_index = 0;
      state.page = 1;
    }

    const config = SEARCH_CONFIGS[state.config_index % total];
    console.log(`Config ${state.config_index + 1}/${total}: ${config.label} | Page: ${state.page}`);

    const apolloData = await apolloRequest('/v1/mixed_people/api_search', {
      api_key: APOLLO_KEY,
      person_titles: config.titles,
      person_locations: config.locations,
      organization_num_employees_ranges: ['1,10', '11,50', '51,200'],
      contact_email_status: ['verified', 'likely to engage'],
      page: state.page,
      per_page: 100
    });

    if (!apolloData.people || apolloData.people.length === 0) {
      // Mark this config as exhausted
      console.log(`⚠ No contacts found — marking "${config.label}" as exhausted.`);
      if (!state.exhausted.includes(state.config_index % total)) {
        state.exhausted.push(state.config_index % total);
      }
      const nextIdx = findNextConfig({ ...state, config_index: (state.config_index + 1) % total });
      state.config_index = nextIdx;
      state.page = 1;
      await saveState(state);
      const next = SEARCH_CONFIGS[nextIdx];
      console.log(`Next run: ${next.label}`);
      return;
    }

    console.log(`Fetched ${apolloData.people.length} contacts. Enriching emails...`);

    const enriched = [];
    for (let i = 0; i < apolloData.people.length; i += 10) {
      const batch = apolloData.people.slice(i, i + 10);
      const matches = await enrichBatch(batch);
      enriched.push(...matches);
      console.log(`Enriched batch ${Math.floor(i/10)+1}: ${matches.length} matches`);
    }

    const emailMap = {};
    for (const m of enriched) {
      if (m.id && m.email) emailMap[m.id] = m.email;
    }

    const contacts = apolloData.people
      .filter(p => emailMap[p.id])
      .map(p => ({
        apollo_id: p.id,
        firstname: p.first_name || '',
        lastname: p.last_name || '',
        email: emailMap[p.id],
        company: p.organization ? p.organization.name : '',
        phone: p.phone_numbers && p.phone_numbers[0] ? p.phone_numbers[0].raw_number : '',
        linkedin: p.linkedin_url || '',
        city: p.city || '',
        state: p.state || '',
        country: p.country || '',
        status: 'NEW',
        last_synced: new Date().toISOString()
      }));

    console.log(`${contacts.length} contacts with verified emails.`);

    const result = await supabaseRequest('POST', '/rest/v1/contacts?on_conflict=apollo_id', contacts);
    console.log(`Supabase status: ${result.status}`);

    if (result.status === 200 || result.status === 201) {
      console.log(`✅ Synced ${contacts.length} contacts to Supabase.`);
    } else {
      console.log(`❌ Supabase error: ${JSON.stringify(result.body).substring(0, 300)}`);
    }

    // Move to next page or next config if low results
    if (apolloData.people.length < 50) {
      console.log(`Low results (${apolloData.people.length}) — moving to next config.`);
      if (!state.exhausted.includes(state.config_index % total)) {
        state.exhausted.push(state.config_index % total);
      }
      const nextIdx = findNextConfig({ ...state, config_index: (state.config_index + 1) % total });
      state.config_index = nextIdx;
      state.page = 1;
    } else {
      state.page += 1;
    }

    await saveState(state);
    const next = SEARCH_CONFIGS[state.config_index % total];
    console.log(`Next run: ${next.label}, page ${state.page}`);
    console.log(`Exhausted configs (${state.exhausted.length}): ${state.exhausted.map(i => SEARCH_CONFIGS[i].label).join(', ') || 'none'}`);

  } catch (err) {
    console.error('❌ Sync failed:', err.message);
    process.exit(1);
  }
}

syncContacts();
