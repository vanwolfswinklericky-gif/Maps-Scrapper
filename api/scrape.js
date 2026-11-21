// api/scrape.js
const axios = require('axios');
const cheerio = require('cheerio');
const XLSX = require('xlsx');

// ==================== NEW API CONSTANTS ====================
const TEXT_SEARCH_URL = 'https://places.googleapis.com/v1/places:searchText';
const DETAILS_URL_BASE = 'https://places.googleapis.com/v1/places/';

// ==================== EMAIL REGEX ====================
const EMAIL_REGEX = /\b[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,24}\b/g;
const EMAIL_REGEX_STRICT = /\b[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.(com|org|net|edu|gov|co|uk|ng|io|ai|me|us|ca|au|de|fr|es|it|br|in|jp|cn|ru|za|mx|nl|se|no|fi|dk|ch|at|be|pl|cz|gr|pt|ie|nz|sg|hk|my|ph|vn|th|id|tr|il|ae|sa|eg|ke|gh)\b/gi;

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
  CONCURRENCY_FOR_SITES: 4,
  // New API requires explicit field masks [citation:1][citation:5]
  TEXT_SEARCH_FIELDS: 'places.id,places.displayName,places.formattedAddress,places.websiteUri,places.nationalPhoneNumber,places.rating,places.userRatingCount,places.types,places.googleMapsUri',
  DETAILS_FIELDS: 'id,displayName,formattedAddress,websiteUri,nationalPhoneNumber,rating,userRatingCount,types,googleMapsUri'
};

// ==================== EMAIL HELPERS ====================
function normalizeEmail(raw) {
  if (!raw) return '';
  return raw
    .replace(/^mailto:/i, '').replace(/^email:/i, '').replace(/^e-mail:/i, '')
    .replace(/\?subject=.*$/i, '').replace(/\?body=.*$/i, '')
    .replace(/^["'<(\[]+|["'>)\]]+$/g, '')
    .trim().toLowerCase()
    .replace(/[.,;:!?]+$/, '');
}

function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;
  const lower = email.toLowerCase().trim();
  if (lower.length < 6 || lower.length > 254) return false;
  for (const bad of EMAIL_BLOCKLIST) if (lower.includes(bad.toLowerCase())) return false;
  const parts = lower.split('@');
  if (parts.length !== 2) return false;
  const [local, domain] = parts;
  if (!local || !domain || local.length > 64 || !domain.includes('.')) return false;
  const tld = domain.split('.').pop();
  if (!tld || tld.length < 2 || !/^[a-z]+$/i.test(tld)) return false;
  if (domain.startsWith('.') || domain.endsWith('.')) return false;
  if (/^[0-9a-f]{16,}$/i.test(local)) return false;
  return true;
}

function scoreEmail(email) {
  let score = 0;
  for (const { re, weight } of EMAIL_PRIORITY) if (re.test(email)) { score += weight; break; }
  const local = email.split('@')[0];
  if (local.length <= 12) score += 2;
  if (!/\d{3,}/.test(local)) score += 1;
  return score;
}

// ==================== EMAIL EXTRACTION ====================
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
    if (!ct.includes('html') && !ct.includes('xml') && !ct.includes('text')) return null;
    return { url: response.request?.res?.responseUrl || url, html: response.data };
  } catch (_) { return null; }
}

// ==================== PAGE PARSING ====================
function parseEmailsFromHtml(html, baseUrl) {
  const $ = cheerio.load(html);
  const emails = new Set();
  $('script[src], style, noscript, iframe, svg').remove();

  $('a[href^="mailto:" i]').each((_, el) => {
    for (const e of extractEmails($(el).attr('href') || '')) emails.add(e);
  });
  $('a[href]').each((_, el) => {
    const href = $(el).attr('href') || '';
    if (/mailto/i.test(href)) for (const e of extractEmails(href)) emails.add(e);
    if (/%40|mailto/i.test(href)) {
      try { for (const e of extractEmails(decodeURIComponent(href))) emails.add(e); } catch (_) {}
    }
  });
  for (const e of extractEmails($('body').text())) emails.add(e);
  for (const sel of ['[class*="email" i]', '[id*="email" i]', '[class*="contact" i]', 'footer', 'address']) {
    $(sel).each((_, el) => { for (const e of extractEmails($(el).text())) emails.add(e); });
  }
  $('script[type="application/ld+json"]').each((_, el) => {
    for (const e of extractEmails($(el).html() || '')) emails.add(e);
  });
  $('meta[content], [data-email], [data-mail]').each((_, el) => {
    for (const c of [$(el).attr('content'), $(el).attr('data-email'), $(el).attr('data-mail')].filter(Boolean)) {
      for (const e of extractEmails(c)) emails.add(e);
    }
  });
  $('script:not([src])').each((_, el) => {
    const decoded = ($(el).html() || '').replace(/\\u0040/g, '@').replace(/&#64;/g, '@').replace(/%40/g, '@');
    for (const e of extractEmails(decoded)) if (!/\.(png|jpg|css|js|ico)$/i.test(e.split('@')[1] || '')) emails.add(e);
  });
  return Array.from(emails);
}

// ==================== LINK DISCOVERY ====================
function findContactLinks(html, baseUrl) {
  try {
    const $ = cheerio.load(html);
    const links = [], seen = new Set();
    const baseHost = new URL(baseUrl).hostname.replace(/^www\./, '');
    $('a[href]').each((_, el) => {
      const href = $(el).attr('href');
      if (!href || /^(mailto:|tel:|javascript:|#)/i.test(href)) return;
      const text = ($(el).text() || '').toLowerCase();
      if (!CONTACT_PATHS.some(p => href.toLowerCase().includes(p) || text.includes(p))) return;
      try {
        const full = new URL(href, baseUrl).toString();
        if (new URL(full).hostname.replace(/^www\./, '') !== baseHost) return;
        if (!seen.has(full)) { seen.add(full); links.push(full); }
      } catch (_) {}
    });
    return links.slice(0, CONFIG.MAX_CONTACT_PAGES);
  } catch (_) { return []; }
}

async function fetchSitemapContactUrls(baseUrl) {
  if (!CONFIG.SITEMAP_CHECK) return [];
  try {
    const res = await fetchHtml(new URL('/sitemap.xml', baseUrl).toString());
    if (!res || typeof res.html !== 'string') return [];
    return (res.html.match(/<loc>([^<]+)<\/loc>/gi) || [])
      .map(m => m.replace(/<\/?loc>/gi, '').trim())
      .filter(u => CONTACT_PATHS.some(p => u.toLowerCase().includes(p)))
      .slice(0, 3);
  } catch (_) { return []; }
}

// ==================== WEBSITE EMAIL FINDER ====================
async function findEmailsForWebsite(website) {
  const result = { emails: [], pagesTried: 0, emailSource: null };
  if (!website) return result;
  let startUrl;
  try { startUrl = new URL(website.startsWith('http') ? website : `http://${website}`).toString(); }
  catch (_) { return result; }

  const visited = new Set(), queue = [startUrl], allEmails = new Map();
  while (queue.length && result.pagesTried < CONFIG.MAX_TOTAL_PAGES_PER_SITE && allEmails.size < CONFIG.MAX_EMAILS_PER_SITE) {
    const url = queue.shift();
    if (!url || visited.has(url)) continue;
    visited.add(url); result.pagesTried++;
    const page = await fetchHtml(url);
    if (!page) continue;
    for (const e of parseEmailsFromHtml(page.html, page.url || url)) {
      if (!allEmails.has(e)) allEmails.set(e, { score: scoreEmail(e), source: page.url || url });
    }
    if (allEmails.size > 0 && Math.max(...Array.from(allEmails.values()).map(v => v.score)) >= 10) break;
    if (url === startUrl || visited.size === 1) {
      for (const l of findContactLinks(page.html, page.url || url)) if (!visited.has(l) && !queue.includes(l)) queue.push(l);
      if (allEmails.size === 0) for (const l of await fetchSitemapContactUrls(page.url || url)) if (!visited.has(l) && !queue.includes(l)) queue.push(l);
    }
    if (result.pagesTried < CONFIG.MAX_TOTAL_PAGES_PER_SITE) await new Promise(r => setTimeout(r, CONFIG.INTER_REQUEST_DELAY_MS));
  }
  const sorted = Array.from(allEmails.entries()).sort((a, b) => b[1].score - a[1].score).slice(0, CONFIG.MAX_EMAILS_PER_SITE);
  result.emails = sorted.map(s => s[0]);
  if (sorted.length) result.emailSource = sorted[0][1].source;
  return result;
}

// ==================== GOOGLE SEARCH FALLBACK ====================
async function googleSearchForWebsite(businessName, address) {
  if (!CONFIG.USE_GOOGLE_SEARCH_FALLBACK) return null;
  try {
    const q = encodeURIComponent(`${businessName} ${address} official website`);
    const res = await axios.get(`${CONFIG.GOOGLE_SEARCH_URL}?q=${q}`, {
      timeout: CONFIG.FETCH_TIMEOUT_MS, headers: HEADERS, maxRedirects: 3, validateStatus: () => true
    });
    const $ = cheerio.load(res.data);
    for (const el of $('a[href^="/url?q="]').toArray()) {
      const m = ($(el).attr('href') || '').match(/\/url\?q=([^&]+)/);
      if (m) {
        try {
          const u = decodeURIComponent(m[1]);
          if (!/google\.|youtube\.|facebook\.|instagram\.|twitter\.|linkedin\.|yelp\.|tripadvisor\.|wikipedia\./i.test(u)) return u;
        } catch (_) {}
      }
    }
  } catch (_) {}
  return null;
}

// ==================== NEW PLACE DETAILS ====================
async function getPlaceDetails(placeId, apiKey) {
  try {
    const response = await axios.get(`${DETAILS_URL_BASE}${placeId}`, {
      headers: {
        'Content-Type': 'application/json',
        'X-Goog-Api-Key': apiKey,
        'X-Goog-FieldMask': CONFIG.DETAILS_FIELDS
      },
      timeout: 30000,
      validateStatus: () => true
    });
    const data = response.data || {};
    if (response.status !== 200) {
      console.log(`  ⚠️ Details HTTP ${response.status} for ${placeId}: ${data.error?.message || ''}`);
      return {};
    }
    return data;
  } catch (err) {
    console.log(`  ❌ Details threw for ${placeId}: ${err.message}`);
    return {};
  }
}

// ==================== NEW TEXT SEARCH ====================
async function collectPlaces(query, location, apiKey, maxResults) {
  const tag = `[collectPlaces "${location}"]`;
  console.log(`${tag} START query="${query}" max=${maxResults}`);

  const results = [], seen = new Set();
  const queryFormats = [`${query} in ${location}`, `${query} near ${location}`, `${query} ${location}`];

  for (let fi = 0; fi < queryFormats.length; fi++) {
    if (results.length >= maxResults) break;
    const textQuery = queryFormats[fi];
    console.log(`${tag} format ${fi + 1}: "${textQuery}"`);

    let pageToken = null, page = 0;
    do {
      if (page >= 3 || results.length >= maxResults) break;
      const body = { textQuery, pageSize: 20 };
      if (pageToken) body.pageToken = pageToken;

      try {
        const response = await axios.post(TEXT_SEARCH_URL, body, {
          headers: {
            'Content-Type': 'application/json',
            'X-Goog-Api-Key': apiKey,
            'X-Goog-FieldMask': CONFIG.TEXT_SEARCH_FIELDS
          },
          timeout: 30000,
          validateStatus: () => true
        });

        const data = response.data || {};
        const places = data.places || [];

        console.log(`${tag} HTTP=${response.status} places=${places.length} next=${Boolean(data.nextPageToken)}`);
        if (response.status !== 200) {
          console.log(`${tag} ❌ API error: ${data.error?.message || JSON.stringify(data).slice(0, 300)}`);
          if (response.status === 401 || response.status === 403) return results;
          break;
        }

        let added = 0;
        for (const item of places) {
          const pid = item.id;
          if (!pid || seen.has(pid)) continue;
          seen.add(pid);
          results.push(item);
          added++;
          if (results.length >= maxResults) break;
        }
        console.log(`${tag} added=${added} total=${results.length}`);

        pageToken = data.nextPageToken || null;
        if (pageToken) await new Promise(r => setTimeout(r, 2000));
        page++;
      } catch (err) {
        console.log(`${tag} ❌ threw: ${err.message}`);
        break;
      }
    } while (pageToken && results.length < maxResults);
  }
  console.log(`${tag} FINAL total=${results.length}`);
  return results.slice(0, maxResults);
}

// ==================== FIELD EXTRACTION HELPERS ====================
function getDisplayName(place) {
  return place.displayName?.text || place.displayName || '';
}
function getWebsite(place) {
  return place.websiteUri || '';
}
function getPhone(place) {
  return place.nationalPhoneNumber || place.internationalPhoneNumber || '';
}
function getAddress(place) {
  return place.formattedAddress || '';
}
function getMapsUri(place) {
  return place.googleMapsUri || '';
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
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const API_KEY = process.env.GOOGLE_PLACES_API_KEY;
    if (!API_KEY) return res.status(500).json({ error: 'GOOGLE_PLACES_API_KEY not configured' });

    const body = req.body || {};
    const query = typeof body.query === 'string' ? body.query.trim() : '';
    const location = typeof body.location === 'string' ? body.location.trim() : '';
    const subAreas = typeof body.subAreas === 'string' ? body.subAreas.trim() : '';
    const skipEmails = Boolean(body.skipEmails);
    const skipWebsiteLookup = Boolean(body.skipWebsiteLookup);
    const maxResultsNum = Math.min(Math.max(parseInt(body.maxResults, 10) || 100, 1), 2000);

    console.log(`📋 query="${query}" location="${location}" max=${maxResultsNum} skipEmails=${skipEmails}`);

    if (!query || !location) return res.status(400).json({ error: 'query and location must be non-empty strings' });

    const locations = [location];
    if (subAreas) locations.push(...subAreas.split(',').map(s => s.trim()).filter(Boolean));
    console.log(`📍 Locations (${locations.length}): ${locations.join(' | ')}`);

    // -------- Collect places using NEW API --------
    const allPlaces = [], seenIds = new Set();
    const perLocationCap = Math.max(1, Math.ceil(maxResultsNum / locations.length));

    for (const area of locations) {
      if (allPlaces.length >= maxResultsNum) break;
      const cap = Math.min(perLocationCap, maxResultsNum - allPlaces.length);
      const found = await collectPlaces(query, area, API_KEY, cap);
      for (const p of found) {
        if (p.id && !seenIds.has(p.id)) { seenIds.add(p.id); allPlaces.push(p); }
      }
    }

    const finalPlaces = allPlaces.slice(0, maxResultsNum);
    console.log(`🔎 DEBUG: locations=${locations.length} allPlaces=${allPlaces.length} finalPlaces=${finalPlaces.length}`);

    if (finalPlaces.length === 0) {
      return res.status(200).json({
        success: true, total: 0, emailsFound: 0, emailPercentage: 0, data: [], excel: null,
        message: 'No businesses found. Check logs for API status. Ensure Places API (New) is enabled and key has no API restrictions blocking it.',
        debug: { query, location, locationsSearched: locations }
      });
    }

    // -------- Process each business --------
    const rows = new Array(finalPlaces.length);
    let emailsFound = 0, completed = 0;

    const processOne = async (place, index) => {
      // Place Details (New) — note the different response structure [citation:2][citation:14]
      const details = await getPlaceDetails(place.id, API_KEY);
      await new Promise(r => setTimeout(r, CONFIG.DETAILS_DELAY_MS));

      let website = getWebsite(details) || getWebsite(place) || '';
      let email = '', emailSource = '';

      if (!skipEmails) {
        if (!website && !skipWebsiteLookup) {
          const guessed = await googleSearchForWebsite(
            getDisplayName(details) || getDisplayName(place),
            getAddress(details) || getAddress(place) || location
          );
          if (guessed) website = guessed;
        }
        if (website) {
          const found = await findEmailsForWebsite(website);
          if (found.emails.length) { email = found.emails[0]; emailSource = found.emailSource; emailsFound++; }
        }
      }

      rows[index] = {
        name: getDisplayName(details) || getDisplayName(place) || '',
        category: (details.types || place.types || []).slice(0, 2).join(', '),
        phone: getPhone(details) || getPhone(place) || '',
        website,
        email,
        emailSource,
        address: getAddress(details) || getAddress(place) || '',
        rating: details.rating || place.rating || '',
        reviewCount: details.userRatingCount || place.userRatingCount || '',
        mapsLink: getMapsUri(details) || getMapsUri(place) || ''
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
    const emailPercentage = validRows.length ? Math.round((totalWithEmail / validRows.length) * 100) : 0;

    // -------- Build Excel --------
    const wb = XLSX.utils.book_new();
    const header = ['Name', 'Email', 'Phone', 'Website', 'Category', 'Address', 'Rating', 'Review Count', 'Google Maps Link', 'Email Source'];
    const wsData = [header, ...validRows.map(r => [
      r.name, r.email, r.phone, r.website, r.category, r.address, r.rating, r.reviewCount, r.mapsLink, r.emailSource
    ])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(wsData), 'All Leads');

    const emailRows = validRows.filter(r => r.email);
    if (emailRows.length) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
        ['Name', 'Email', 'Phone', 'Website', 'Category', 'Address', 'Email Source'],
        ...emailRows.map(r => [r.name, r.email, r.phone, r.website, r.category, r.address, r.emailSource])
      ]), 'Emails Only');
    }

    const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    console.log(`\n✅ COMPLETE: ${validRows.length} businesses | ${totalWithEmail} emails (${emailPercentage}%) | ${((Date.now() - startedAt) / 1000).toFixed(1)}s`);

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