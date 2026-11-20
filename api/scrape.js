// api/scrape.js
const axios = require('axios');
const cheerio = require('cheerio');
const XLSX = require('xlsx');

// ==================== CONSTANTS ====================
const TEXT_SEARCH_URL = "https://maps.googleapis.com/maps/api/place/textsearch/json";
const DETAILS_URL = "https://maps.googleapis.com/maps/api/place/details/json";

// ==================== EMAIL REGEX ====================
const EMAIL_REGEX = /\b[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,24}\b/g;
const EMAIL_REGEX_STRICT = /\b[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.(com|org|net|edu|gov|co|uk|ng|io|ai|me|us|ca|au|de|fr|es|it|br|in|jp|cn|ru|za|mx|nl|se|no|fi|dk|ch|at|be|pl|cz|gr|pt|ie|nz|sg|hk|my|ph|vn|th|id|tr|il|ae|sa|eg|ke|gh)\b/gi;
const OBFUSCATED_REGEXES = [
  /([a-zA-Z0-9._%+\-]+)\s*(?:\[at\]|\(at\)|\{at\}|\bat\b)\s*([a-zA-Z0-9.\-]+)\s*(?:\[dot\]|\(dot\)|\{dot\}|\bdot\b)\s*([a-zA-Z]{2,24})/gi,
  /([a-zA-Z0-9._%+\-]+)\s*@\s*([a-zA-Z0-9.\-]+)\s*\.\s*([a-zA-Z]{2,24})/g
];

// ==================== BLOCKLIST ====================
const EMAIL_BLOCKLIST = [
  'sentry.io', 'wixpress.com', 'example.com', 'godaddy.com', 'domain.com',
  'yourname@', 'info@domain', 'contact@domain', 'admin@domain',
  'support@domain', 'sales@domain', 'test@test', 'user@example',
  'noreply@', 'no-reply@', 'donotreply@', 'postmaster@',
  '.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp', '.css', '.js',
  '.ico', '.woff', '.ttf', '.eot', '.mp4', '.pdf',
  'data:image', 'javascript:', 'mailto:',
  'sentry', 'cloudflare', 'google-analytics', 'facebook.com', 'twitter.com',
  'instagram.com', 'youtube.com', 'linkedin.com', 'pinterest.com',
  'schema.org', 'w3.org', 'creativecommons.org',
  'jquery', 'bootstrap', 'font-awesome'
];

// ==================== CONTACT PATHS ====================
const CONTACT_PATHS = [
  'contact', 'contact-us', 'contactus', 'contact_',
  'about', 'about-us', 'aboutus', 'about_',
  'reach-us', 'reach', 'get-in-touch', 'touch', 'connect',
  'team', 'our-team', 'meet-the-team', 'staff', 'people',
  'support', 'help', 'customer-service', 'inquiry', 'enquiries', 'enquiry',
  'contact-page', 'kontakt', 'contacto', 'contactez', 'impressum',
  'privacy', 'legal', 'terms'
];

// ==================== EMAIL PRIORITY ====================
const EMAIL_PRIORITY = [
  { re: /^(info|hello|contact|enquiries|enquiry|sales|support|help)@/i, weight: 10 },
  { re: /^(admin|office|team|mail|email|reach|hi|hey)@/i, weight: 8 },
  { re: /^(founder|ceo|owner|manager|director|partnerships|booking|reservations)@/i, weight: 6 },
  { re: /^(careers|jobs|hr|pr)@/i, weight: 4 }
];

// ==================== HEADERS ====================
const HEADERS = {
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
  'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
  'Accept-Language': 'en-US,en;q=0.9',
  'Accept-Encoding': 'gzip, deflate, br',
  'Connection': 'keep-alive',
  'Upgrade-Insecure-Requests': '1',
  'Cache-Control': 'no-cache',
  'Pragma': 'no-cache'
};

// ==================== CONFIG ====================
const CONFIG = {
  FETCH_TIMEOUT_MS: 8000,
  MAX_CONTACT_PAGES: 8,
  MAX_TOTAL_PAGES_PER_SITE: 15,
  MAX_EMAILS_PER_SITE: 5,
  INTER_REQUEST_DELAY_MS: 30,
  DETAILS_DELAY_MS: 40,
  SITEMAP_CHECK: true,
  USE_GOOGLE_SEARCH_FALLBACK: true,
  GOOGLE_SEARCH_URL: 'https://www.google.com/search',
  CONCURRENCY_FOR_SITES: 4
};

// ==================== EMAIL HELPERS ====================
function normalizeEmail(raw) {
  if (!raw) return '';
  let e = raw
    .replace(/^mailto:/i, '')
    .replace(/^email:/i, '')
    .replace(/^e-mail:/i, '')
    .replace(/\?subject=.*$/i, '')
    .replace(/\?body=.*$/i, '')
    .replace(/^["'<(\[]+|["'>)\]]+$/g, '')
    .trim()
    .toLowerCase();
  e = e.replace(/[.,;:!?]+$/, '');
  return e;
}

function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;
  const lower = email.toLowerCase().trim();
  if (lower.length < 6 || lower.length > 254) return false;

  for (const bad of EMAIL_BLOCKLIST) {
    if (lower.includes(bad.toLowerCase())) return false;
  }

  const parts = lower.split('@');
  if (parts.length !== 2) return false;
  const [local, domain] = parts;
  if (!local || !domain) return false;
  if (local.length > 64) return false;
  if (!domain.includes('.')) return false;
  const tld = domain.split('.').pop();
  if (!tld || tld.length < 2 || !/^[a-z]+$/i.test(tld)) return false;
  if (domain.startsWith('.') || domain.endsWith('.')) return false;
  if (/^[0-9a-f]{16,}$/i.test(local)) return false;

  return true;
}

function scoreEmail(email) {
  let score = 0;
  for (const { re, weight } of EMAIL_PRIORITY) {
    if (re.test(email)) { score += weight; break; }
  }
  const local = email.split('@')[0];
  if (local.length <= 12) score += 2;
  if (!/\d{3,}/.test(local)) score += 1;
  return score;
}

// ==================== EMAIL EXTRACTION ====================
function deobfuscate(text) {
  return text
    .replace(/\s*\[at\]\s*/gi, '@')
    .replace(/\s*\(at\)\s*/gi, '@')
    .replace(/\s*\{at\}\s*/gi, '@')
    .replace(/\s*\[dot\]\s*/gi, '.')
    .replace(/\s*\(dot\)\s*/gi, '.')
    .replace(/\s*\{dot\}\s*/gi, '.');
}

function extractEmails(text) {
  if (!text || typeof text !== 'string') return [];
  const found = new Set();

  for (const m of text.match(EMAIL_REGEX_STRICT) || []) {
    const e = normalizeEmail(m);
    if (isValidEmail(e)) found.add(e);
  }
  for (const m of text.match(EMAIL_REGEX) || []) {
    const e = normalizeEmail(m);
    if (isValidEmail(e)) found.add(e);
  }
  const deob = deobfuscate(text);
  for (const re of OBFUSCATED_REGEXES) {
    let m;
    re.lastIndex = 0;
    while ((m = re.exec(deob)) !== null) {
      const candidate = normalizeEmail(`${m[1]}@${m[2]}.${m[3]}`);
      if (isValidEmail(candidate)) found.add(candidate);
    }
  }
  return Array.from(found);
}

// ==================== HTTP FETCH ====================
async function fetchHtml(url) {
  try {
    const response = await axios.get(url, {
      timeout: CONFIG.FETCH_TIMEOUT_MS,
      headers: HEADERS,
      maxRedirects: 5,
      validateStatus: s => s >= 200 && s < 400,
      maxContentLength: 5 * 1024 * 1024
    });
    const ct = (response.headers['content-type'] || '').toLowerCase();
    if (!ct.includes('html') && !ct.includes('xml') && !ct.includes('text')) {
      return null;
    }
    return { url: response.request?.res?.responseUrl || url, html: response.data };
  } catch (err) {
    return null;
  }
}

// ==================== PAGE PARSING ====================
function parseEmailsFromHtml(html, baseUrl) {
  const $ = cheerio.load(html);
  const emails = new Set();

  $('script[src], style, noscript, iframe, svg').remove();

  $('a[href^="mailto:" i]').each((_, el) => {
    const href = $(el).attr('href') || '';
    for (const e of extractEmails(href)) emails.add(e);
  });

  $('a[href]').each((_, el) => {
    const href = $(el).attr('href') || '';
    if (/mailto/i.test(href)) {
      for (const e of extractEmails(href)) emails.add(e);
    }
    if (/%40|mailto/i.test(href)) {
      try {
        for (const e of extractEmails(decodeURIComponent(href))) emails.add(e);
      } catch (_) {}
    }
  });

  const bodyText = $('body').text() + '\n' + ($('title').text() || '');
  for (const e of extractEmails(bodyText)) emails.add(e);

  const selectors = [
    '[class*="email" i]', '[id*="email" i]', '[class*="mail" i]', '[id*="mail" i]',
    '[class*="contact" i]', '[id*="contact" i]', '[itemprop="email"]',
    'footer', 'address', '.footer', '#footer', '[class*="footer" i]',
    '[class*="vcard" i]', '[class*="business" i]'
  ];
  for (const sel of selectors) {
    $(sel).each((_, el) => {
      const t = $(el).text();
      for (const e of extractEmails(t)) emails.add(e);
    });
  }

  $('script[type="application/ld+json"]').each((_, el) => {
    const raw = $(el).html() || '';
    for (const e of extractEmails(raw)) emails.add(e);
  });

  $('meta[content], [data-email], [data-mail]').each((_, el) => {
    const candidates = [
      $(el).attr('content'),
      $(el).attr('data-email'),
      $(el).attr('data-mail'),
      $(el).attr('data-contact')
    ].filter(Boolean);
    for (const c of candidates) {
      for (const e of extractEmails(c)) emails.add(e);
    }
  });

  $('script:not([src])').each((_, el) => {
    const raw = $(el).html() || '';
    const decoded = raw
      .replace(/\\u0040/g, '@')
      .replace(/&#64;/g, '@')
      .replace(/&commat;/g, '@')
      .replace(/%40/g, '@');
    for (const e of extractEmails(decoded)) {
      if (!/\.(png|jpg|jpeg|gif|svg|webp|css|js|ico)$/i.test(e.split('@')[1] || '')) {
        emails.add(e);
      }
    }
  });

  return Array.from(emails);
}

// ==================== LINK DISCOVERY ====================
function findContactLinks(html, baseUrl) {
  try {
    const $ = cheerio.load(html);
    const links = [];
    const seen = new Set();
    const baseHost = new URL(baseUrl).hostname.replace(/^www\./, '');

    $('a[href]').each((_, el) => {
      const href = $(el).attr('href');
      if (!href) return;
      const text = ($(el).text() || '').toLowerCase().trim();
      const hrefLower = href.toLowerCase();

      const isContact =
        CONTACT_PATHS.some(p => hrefLower.includes(p)) ||
        CONTACT_PATHS.some(p => text.includes(p)) ||
        /contact|about|reach|touch|team/i.test(text);

      if (!isContact) return;
      if (/^(mailto:|tel:|javascript:|#)/i.test(href)) return;

      try {
        const full = new URL(href, baseUrl).toString();
        const linkHost = new URL(full).hostname.replace(/^www\./, '');
        if (linkHost !== baseHost) return;
        if (!seen.has(full)) {
          seen.add(full);
          links.push(full);
        }
      } catch (_) {}
    });

    return links.slice(0, CONFIG.MAX_CONTACT_PAGES);
  } catch (_) {
    return [];
  }
}

async function fetchSitemapContactUrls(baseUrl) {
  if (!CONFIG.SITEMAP_CHECK) return [];
  try {
    const sitemapUrl = new URL('/sitemap.xml', baseUrl).toString();
    const res = await fetchHtml(sitemapUrl);
    if (!res) return [];
    const urls = [];
    const raw = typeof res.html === 'string' ? res.html : '';
    const urlMatches = raw.match(/<loc>([^<]+)<\/loc>/gi) || [];
    for (const m of urlMatches) {
      const u = m.replace(/<\/?loc>/gi, '').trim();
      if (CONTACT_PATHS.some(p => u.toLowerCase().includes(p))) urls.push(u);
    }
    return urls.slice(0, 3);
  } catch (_) {
    return [];
  }
}

// ==================== WEBSITE EMAIL FINDER ====================
async function findEmailsForWebsite(website) {
  const result = { emails: [], pagesTried: 0, emailSource: null };
  if (!website) return result;

  let startUrl;
  try {
    startUrl = new URL(website.startsWith('http') ? website : `http://${website}`).toString();
  } catch (_) {
    return result;
  }

  const visited = new Set();
  const queue = [startUrl];
  const allEmails = new Map();

  while (queue.length && result.pagesTried < CONFIG.MAX_TOTAL_PAGES_PER_SITE && allEmails.size < CONFIG.MAX_EMAILS_PER_SITE) {
    const url = queue.shift();
    if (!url || visited.has(url)) continue;
    visited.add(url);
    result.pagesTried++;

    const page = await fetchHtml(url);
    if (!page) continue;

    const pageEmails = parseEmailsFromHtml(page.html, page.url || url);
    for (const e of pageEmails) {
      if (!allEmails.has(e)) {
        allEmails.set(e, { score: scoreEmail(e), source: page.url || url });
      }
    }

    if (allEmails.size > 0) {
      const best = Math.max(...Array.from(allEmails.values()).map(v => v.score));
      if (best >= 10) break;
    }

    if (url === startUrl || visited.size === 1) {
      const contactLinks = findContactLinks(page.html, page.url || url);
      for (const l of contactLinks) {
        if (!visited.has(l) && !queue.includes(l)) queue.push(l);
      }
      if (allEmails.size === 0) {
        const sitemapLinks = await fetchSitemapContactUrls(page.url || url);
        for (const l of sitemapLinks) {
          if (!visited.has(l) && !queue.includes(l)) queue.push(l);
        }
      }
    }

    if (result.pagesTried < CONFIG.MAX_TOTAL_PAGES_PER_SITE) {
      await new Promise(r => setTimeout(r, CONFIG.INTER_REQUEST_DELAY_MS));
    }
  }

  const sorted = Array.from(allEmails.entries())
    .sort((a, b) => b[1].score - a[1].score)
    .slice(0, CONFIG.MAX_EMAILS_PER_SITE)
    .map(([email, meta]) => ({ email, ...meta }));

  result.emails = sorted.map(s => s.email);
  if (sorted.length) result.emailSource = sorted[0].source;
  return result;
}

// ==================== GOOGLE SEARCH FALLBACK ====================
async function googleSearchForWebsite(businessName, address) {
  if (!CONFIG.USE_GOOGLE_SEARCH_FALLBACK) return null;
  try {
    const q = encodeURIComponent(`${businessName} ${address} official website`);
    const res = await axios.get(`${CONFIG.GOOGLE_SEARCH_URL}?q=${q}`, {
      timeout: CONFIG.FETCH_TIMEOUT_MS,
      headers: HEADERS,
      maxRedirects: 3,
      validateStatus: () => true
    });
    const $ = cheerio.load(res.data);
    const links = [];
    $('a[href^="/url?q="]').each((_, el) => {
      const href = $(el).attr('href') || '';
      const m = href.match(/\/url\?q=([^&]+)/);
      if (m) {
        try {
          const u = decodeURIComponent(m[1]);
          if (!/google\.|youtube\.|facebook\.|instagram\.|twitter\.|linkedin\.|yelp\.|tripadvisor\.|wikipedia\./i.test(u)) {
            links.push(u);
          }
        } catch (_) {}
      }
    });
    return links[0] || null;
  } catch (_) {
    return null;
  }
}

// ==================== GOOGLE PLACES ====================
async function getPlaceDetails(placeId, apiKey) {
  const params = {
    place_id: placeId,
    fields: 'name,formatted_address,formatted_phone_number,international_phone_number,website,rating,user_ratings_total,url,types,business_status',
    key: apiKey
  };
  try {
    const response = await axios.get(DETAILS_URL, {
      params,
      timeout: 30000,
      validateStatus: () => true
    });
    const data = response.data || {};
    if (data.status && data.status !== 'OK') {
      console.log(`  ⚠️ details status=${data.status} for ${placeId}: ${data.error_message || ''}`);
    }
    return data.result || {};
  } catch (err) {
    console.log(`  ❌ Details threw for ${placeId}: ${err.message}`);
    return {};
  }
}

// ==================== COLLECT PLACES (with diagnostics) ====================
async function collectPlaces(query, location, apiKey, maxResults) {
  const tag = `[collectPlaces "${location}"]`;
  console.log(`${tag} START query="${query}" max=${maxResults}`);

  const results = [];
  const seen = new Set();
  const queryFormats = [
    `${query} in ${location}`,
    `${query} near ${location}`,
    `${query} ${location}`,
    `${query}+${location}`
  ];

  for (let fi = 0; fi < queryFormats.length; fi++) {
    if (results.length >= maxResults) break;
    const q = queryFormats[fi];
    console.log(`${tag} format ${fi + 1}/${queryFormats.length}: "${q}"`);

    let pageToken = null;
    let page = 0;

    do {
      if (page >= 3 || results.length >= maxResults) break;

      const params = { key: apiKey, query: q };
      if (pageToken) params.pagetoken = pageToken;

      try {
        const response = await axios.get(TEXT_SEARCH_URL, {
          params,
          timeout: 30000,
          validateStatus: () => true
        });

        const data = response.data || {};
        const status = data.status;
        const items = Array.isArray(data.results) ? data.results : [];

        // >>> THE CRITICAL LOG <<<
        console.log(
          `${tag} HTTP=${response.status} status=${status} ` +
          `items=${items.length} next=${Boolean(data.next_page_token)} ` +
          `err="${data.error_message || ''}"`
        );

        if (items.length === 0 && status === 'OK') {
          console.log(`${tag} ⚠️ status=OK but ZERO items. Body:`,
            JSON.stringify(data).slice(0, 600));
        }

        if (status === 'ZERO_RESULTS') {
          console.log(`${tag} (zero results for this format)`);
          break;
        }
        if (status !== 'OK') {
          console.log(`${tag} ❌ non-OK status=${status}`);
          if (status === 'REQUEST_DENIED' || status === 'OVER_QUERY_LIMIT') {
            console.log(`${tag} ⛔ Aborting: auth/quota issue.`);
            return results;
          }
          break;
        }

        let added = 0, skippedNoId = 0, skippedDup = 0;
        for (const item of items) {
          const pid = item.place_id;
          if (!pid) { skippedNoId++; continue; }
          if (seen.has(pid)) { skippedDup++; continue; }
          seen.add(pid);
          results.push(item);
          added++;
          if (results.length >= maxResults) break;
        }
        console.log(
          `${tag} added=${added} skippedNoId=${skippedNoId} ` +
          `skippedDup=${skippedDup} total=${results.length}`
        );

        pageToken = data.next_page_token || null;
        if (pageToken) await new Promise(r => setTimeout(r, 2000));
        page++;
      } catch (err) {
        console.log(`${tag} ❌ threw: ${err.message}`);
        if (err.response) {
          console.log(`     Body: ${JSON.stringify(err.response.data).slice(0, 300)}`);
        }
        break;
      }
    } while (pageToken && results.length < maxResults);
  }

  console.log(`${tag} FINAL total=${results.length}`);
  return results.slice(0, maxResults);
}

// ==================== MAIN HANDLER ====================
module.exports = async (req, res) => {
  const startedAt = Date.now();
  console.log('\n🚀 ===== NEW REQUEST =====');
  console.log(`📅 ${new Date().toISOString()}`);

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const API_KEY = process.env.GOOGLE_PLACES_API_KEY;
    if (!API_KEY) {
      return res.status(500).json({ error: 'GOOGLE_PLACES_API_KEY not configured' });
    }

    // -------- Parse + trim input --------
    const body = req.body || {};
    const query = typeof body.query === 'string' ? body.query.trim() : '';
    const location = typeof body.location === 'string' ? body.location.trim() : '';
    const subAreas = typeof body.subAreas === 'string' ? body.subAreas.trim() : '';
    const skipEmails = Boolean(body.skipEmails);
    const skipWebsiteLookup = Boolean(body.skipWebsiteLookup);
    const maxResultsNum = Math.min(Math.max(parseInt(body.maxResults, 10) || 100, 1), 2000);

    console.log('📋 Parameters:');
    console.log(`   query="${query}"`);
    console.log(`   location="${location}"`);
    console.log(`   subAreas="${subAreas}"`);
    console.log(`   maxResults=${maxResultsNum}`);
    console.log(`   skipEmails=${skipEmails}`);
    console.log(`   skipWebsiteLookup=${skipWebsiteLookup}`);

    if (!query || !location) {
      return res.status(400).json({
        error: 'query and location must be non-empty strings'
      });
    }

    // -------- Build location list --------
    const locations = [location];
    if (subAreas) {
      locations.push(...subAreas.split(',').map(s => s.trim()).filter(Boolean));
    }
    console.log(`📍 Locations (${locations.length}): ${locations.join(' | ')}`);

    const perLocationCap = Math.max(1, Math.ceil(maxResultsNum / locations.length));

    // -------- Collect places --------
    const allPlaces = [];
    const seenIds = new Set();
    for (const area of locations) {
      if (allPlaces.length >= maxResultsNum) break;
      const remaining = maxResultsNum - allPlaces.length;
      const cap = Math.min(perLocationCap, remaining);
      const found = await collectPlaces(query, area, API_KEY, cap);
      for (const p of found) {
        if (p.place_id && !seenIds.has(p.place_id)) {
          seenIds.add(p.place_id);
          allPlaces.push(p);
        }
      }
    }

    const finalPlaces = allPlaces.slice(0, maxResultsNum);

    console.log('\n🔎 DEBUG summary:');
    console.log(`   locations searched: ${locations.length}`);
    console.log(`   allPlaces collected: ${allPlaces.length}`);
    console.log(`   finalPlaces: ${finalPlaces.length}`);
    if (allPlaces.length > 0) {
      console.log(`   first place keys: ${Object.keys(allPlaces[0]).join(', ')}`);
      console.log(`   first place_id: ${allPlaces[0].place_id}`);
      console.log(`   first name: ${allPlaces[0].name}`);
    }

    if (finalPlaces.length === 0) {
      return res.status(200).json({
        success: true,
        total: 0,
        emailsFound: 0,
        emailPercentage: 0,
        data: [],
        excel: null,
        message: 'No businesses found. Check server logs — the [collectPlaces] lines show the raw Google status for each query format.',
        debug: {
          query,
          location,
          locationsSearched: locations,
          apiReached: true
        }
      });
    }

    // -------- Process each business --------
    const rows = new Array(finalPlaces.length);
    let emailsFound = 0;
    let completed = 0;

    const processOne = async (place, index) => {
      const details = await getPlaceDetails(place.place_id, API_KEY);
      await new Promise(r => setTimeout(r, CONFIG.DETAILS_DELAY_MS));

      let website = details.website || '';
      let email = '';
      let emailSource = '';

      if (!skipEmails) {
        if (!website && !skipWebsiteLookup) {
          const guessed = await googleSearchForWebsite(
            details.name || place.name,
            details.formatted_address || place.formatted_address || location
          );
          if (guessed) website = guessed;
        }

        if (website) {
          const found = await findEmailsForWebsite(website);
          if (found.emails.length) {
            email = found.emails[0];
            emailSource = found.emailSource;
            emailsFound++;
          }
        }
      }

      rows[index] = {
        name: details.name || place.name || '',
        category: (details.types || []).slice(0, 2).join(', '),
        phone: details.formatted_phone_number || '',
        website,
        email,
        emailSource,
        address: details.formatted_address || place.formatted_address || '',
        rating: details.rating || '',
        reviewCount: details.user_ratings_total || '',
        mapsLink: details.url || ''
      };

      completed++;
      if (completed % 10 === 0 || completed === finalPlaces.length) {
        console.log(`📊 Progress: ${completed}/${finalPlaces.length} | emails: ${emailsFound}`);
      }
    };

    for (let i = 0; i < finalPlaces.length; i += CONFIG.CONCURRENCY_FOR_SITES) {
      const batch = finalPlaces.slice(i, i + CONFIG.CONCURRENCY_FOR_SITES);
      await Promise.all(batch.map((p, j) => processOne(p, i + j)));
    }

    const validRows = rows.filter(Boolean);
    const totalWithEmail = validRows.filter(r => r.email).length;
    const emailPercentage = validRows.length
      ? Math.round((totalWithEmail / validRows.length) * 100)
      : 0;

    // -------- Build Excel --------
    const wb = XLSX.utils.book_new();
    const header = ['Name', 'Email', 'Phone', 'Website', 'Category', 'Address', 'Rating', 'Review Count', 'Google Maps Link', 'Email Source'];
    const wsData = [header, ...validRows.map(r => [
      r.name, r.email, r.phone, r.website, r.category, r.address,
      r.rating, r.reviewCount, r.mapsLink, r.emailSource
    ])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(wsData), 'All Leads');

    const emailRows = validRows.filter(r => r.email);
    if (emailRows.length) {
      const ews = XLSX.utils.aoa_to_sheet([
        ['Name', 'Email', 'Phone', 'Website', 'Category', 'Address', 'Email Source'],
        ...emailRows.map(r => [r.name, r.email, r.phone, r.website, r.category, r.address, r.emailSource])
      ]);
      XLSX.utils.book_append_sheet(wb, ews, 'Emails Only');
    }

    const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    console.log('\n✅ ===== COMPLETE =====');
    console.log(`   Total: ${validRows.length} | Emails: ${totalWithEmail} (${emailPercentage}%)`);
    console.log(`   Duration: ${((Date.now() - startedAt) / 1000).toFixed(1)}s`);

    return res.status(200).json({
      success: true,
      total: validRows.length,
      emailsFound: totalWithEmail,
      emailPercentage,
      data: validRows,
      excel: excelBuffer.toString('base64'),
      message: `Scraped ${validRows.length} businesses, found ${totalWithEmail} emails (${emailPercentage}%)`
    });

  } catch (error) {
    console.error('❌ SCRAPING ERROR:', error);
    return res.status(500).json({ error: error.message || 'Scraping failed' });
  }
};