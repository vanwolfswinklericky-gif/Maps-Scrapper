// api/scrape.js
const axios = require('axios');
const cheerio = require('cheerio');
const XLSX = require('xlsx');

// ==================== CONSTANTS ====================
const TEXT_SEARCH_URL = "https://maps.googleapis.com/maps/api/place/textsearch/json";
const DETAILS_URL = "https://maps.googleapis.com/maps/api/place/details/json";

// ==================== EMAIL REGEX ====================
// Standard emails (TLD 2-24 chars, any letters)
const EMAIL_REGEX = /\b[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,24}\b/g;
// Strict version to prefer on first pass
const EMAIL_REGEX_STRICT = /\b[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.(com|org|net|edu|gov|co|uk|ng|io|ai|me|us|ca|au|de|fr|es|it|br|in|jp|cn|ru|za|mx|nl|se|no|fi|dk|ch|at|be|pl|cz|gr|pt|ie|nz|sg|hk|my|ph|vn|th|id|tr|il|ae|sa|eg|ke|gh)\b/gi;
// Obfuscated email patterns: "name [at] domain [dot] com"
const OBFUSCATED_REGEXES = [
  /([a-zA-Z0-9._%+\-]+)\s*(?:\[at\]|\(at\)|\{at\}|@\s*at\s*|\bat\b)\s*([a-zA-Z0-9.\-]+)\s*(?:\[dot\]|\(dot\)|\{dot\}|\bdot\b)\s*([a-zA-Z]{2,24})/gi,
  /([a-zA-Z0-9._%+\-]+)\s*@\s*([a-zA-Z0-9.\-]+)\s*\.\s*([a-zA-Z]{2,24})/g
];

// ==================== BLOCKLIST ====================
const EMAIL_BLOCKLIST = [
  // Third-party / placeholder
  'sentry.io', 'wixpress.com', 'example.com', 'godaddy.com', 'domain.com',
  'yourname@', 'info@domain', 'contact@domain', 'admin@domain',
  'support@domain', 'sales@domain', 'test@test', 'user@example',
  'noreply@', 'no-reply@', 'donotreply@', 'postmaster@',
  // Asset / protocol noise
  '.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp', '.css', '.js',
  '.ico', '.woff', '.ttf', '.eot', '.mp4', '.pdf',
  'data:image', 'javascript:', 'mailto:',
  // Common tracking / CDN
  'sentry', 'cloudflare', 'google-analytics', 'facebook.com', 'twitter.com',
  'instagram.com', 'youtube.com', 'linkedin.com', 'pinterest.com',
  'schema.org', 'w3.org', 'creativecommons.org',
  // libphonenumber, bootstrap, jquery markers that sneak in
  'jquery', 'bootstrap', 'font-awesome'
];

// ==================== CONTACT PATHS ====================
const CONTACT_PATHS = [
  'contact', 'contact-us', 'contactus', 'contact-us-1', 'contact_',
  'about', 'about-us', 'aboutus', 'about_',
  'reach-us', 'reach', 'get-in-touch', 'touch', 'connect',
  'team', 'our-team', 'meet-the-team', 'staff', 'people',
  'support', 'help', 'customer-service', 'inquiry', 'enquiries', 'enquiry',
  'contact-page', 'kontakt', 'contacto', 'contactez', 'impressum',
  'privacy', 'legal', 'terms'
];

// ==================== EMAIL PATTERNS (priority scoring) ====================
const EMAIL_PRIORITY = [
  { re: /^(info|hello|contact|enquiries|enquiry|sales|support|help)@/i, weight: 10 },
  { re: /^(admin|office|team|mail|email|reach|hi|hey)@/i, weight: 8 },
  { re: /^(founder|ceo|owner|manager|director|partnerships|booking|reservations)@/i, weight: 6 },
  { re: /^(careers|jobs|hr|pr)@/i, weight: 4 },
];

// ==================== HTTP HEADERS ====================
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
  MAX_TOTAL_PAGES_PER_SITE: 15,      // homepage + contact pages + sitemap pages
  MAX_EMAILS_PER_SITE: 5,            // keep top N, stop early if strong hit
  INTER_REQUEST_DELAY_MS: 30,
  DETAILS_DELAY_MS: 40,
  SITEMAP_CHECK: true,
  USE_GOOGLE_SEARCH_FALLBACK: true,  // find website if missing
  GOOGLE_SEARCH_URL: 'https://www.google.com/search',
  CONCURRENCY_FOR_SITES: 4           // parallel site processing
};

// ==================== HELPERS: EMAIL VALIDATION ====================

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
  // Drop trailing punctuation that regex can capture
  e = e.replace(/[.,;:!?]+$/, '');
  return e;
}

function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;
  const lower = email.toLowerCase().trim();
  if (lower.length < 6 || lower.length > 254) return false;

  // Blocklist
  for (const bad of EMAIL_BLOCKLIST) {
    if (lower.includes(bad.toLowerCase())) return false;
  }

  // Structural
  const parts = lower.split('@');
  if (parts.length !== 2) return false;
  const [local, domain] = parts;
  if (!local || !domain) return false;
  if (local.length > 64) return false;
  if (!domain.includes('.')) return false;
  const tld = domain.split('.').pop();
  if (!tld || tld.length < 2 || !/^[a-z]+$/i.test(tld)) return false;
  if (domain.startsWith('.') || domain.endsWith('.')) return false;

  // Reject things that look like version strings / hashes
  if (/^[0-9a-f]{16,}$/i.test(local)) return false;

  return true;
}

function scoreEmail(email) {
  let score = 0;
  for (const { re, weight } of EMAIL_PRIORITY) {
    if (re.test(email)) { score += weight; break; }
  }
  // Prefer shorter, cleaner local parts
  const local = email.split('@')[0];
  if (local.length <= 12) score += 2;
  if (!/\d{3,}/.test(local)) score += 1;  // avoid hash-y locals
  return score;
}

// ==================== EXTRACTION ====================

function deobfuscate(text) {
  // Normalize common obfuscations before extraction
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

  // Pass 1: strict
  for (const m of text.match(EMAIL_REGEX_STRICT) || []) {
    const e = normalizeEmail(m);
    if (isValidEmail(e)) found.add(e);
  }
  // Pass 2: general
  for (const m of text.match(EMAIL_REGEX) || []) {
    const e = normalizeEmail(m);
    if (isValidEmail(e)) found.add(e);
  }
  // Pass 3: obfuscated
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
      // Cap response size to avoid giant pages
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

  // Remove non-content nodes that commonly hold noise
  $('script[src], style, noscript, iframe, svg').remove();

  // 1. mailto links (strongest signal)
  $('a[href^="mailto:" i], a[href^="MAILTO:" i]').each((_, el) => {
    const href = $(el).attr('href') || '';
    for (const e of extractEmails(href)) emails.add(e);
  });

  // 2. mailto anywhere in href (case-insensitive, possibly encoded)
  $('a[href]').each((_, el) => {
    const href = $(el).attr('href') || '';
    if (/mailto/i.test(href)) {
      for (const e of extractEmails(href)) emails.add(e);
    }
    // also decode percent-encoded mailto
    if (/%40|mailto/i.test(href)) {
      try {
        for (const e of extractEmails(decodeURIComponent(href))) emails.add(e);
      } catch (_) {}
    }
  });

  // 3. Visible text
  const bodyText = $('body').text() + '\n' + ($('title').text() || '');
  for (const e of extractEmails(bodyText)) emails.add(e);

  // 4. Specific email/contact containers
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

  // 5. Inline JSON-LD (schema.org LocalBusiness / Organization / Person)
  $('script[type="application/ld+json"]').each((_, el) => {
    const raw = $(el).html() || '';
    for (const e of extractEmails(raw)) emails.add(e);
  });

  // 6. Meta / data attributes
  $('meta[content], [data-email], [data-mail], [href]').each((_, el) => {
    const candidates = [
      $(el).attr('content'),
      $(el).attr('data-email'),
      $(el).attr('data-mail'),
      $(el).attr('href'),
      $(el).attr('data-contact')
    ].filter(Boolean);
    for (const c of candidates) {
      for (const e of extractEmails(c)) emails.add(e);
    }
  });

  // 7. Encoded emails in remaining scripts (last resort, filtered)
  $('script:not([src])').each((_, el) => {
    const raw = $(el).html() || '';
    // decode \u0040 and HTML entities
    const decoded = raw
      .replace(/\\u0040/g, '@')
      .replace(/&#64;/g, '@')
      .replace(/&commat;/g, '@')
      .replace(/%40/g, '@');
    for (const e of extractEmails(decoded)) {
      // Extra filter for scripts: must look like a real address
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
      // Skip anchors, mailto, tel
      if (/^(mailto:|tel:|javascript:|#)/i.test(href)) return;

      try {
        const full = new URL(href, baseUrl).toString();
        // Same domain only
        const baseHost = new URL(baseUrl).hostname.replace(/^www\./, '');
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

function findSitemapUrl(html, baseUrl) {
  try {
    const $ = cheerio.load(html);
    // <link rel="sitemap" href="...">
    const linkSitemap = $('link[rel="sitemap"]').attr('href');
    if (linkSitemap) return new URL(linkSitemap, baseUrl).toString();
    // robots.txt often declares it, but let's just try the common default
    return new URL('/sitemap.xml', baseUrl).toString();
  } catch (_) {
    return null;
  }
}

async function fetchSitemapContactUrls(baseUrl) {
  if (!CONFIG.SITEMAP_CHECK) return [];
  try {
    const sitemapUrl = new URL('/sitemap.xml', baseUrl).toString();
    const res = await fetchHtml(sitemapUrl);
    if (!res) return [];
    const urls = [];
    // Sitemap might be XML or HTML-ish text; just grep URLs
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

  // Normalize URL
  let startUrl;
  try {
    startUrl = new URL(website.startsWith('http') ? website : `http://${website}`).toString();
  } catch (_) {
    return result;
  }

  const visited = new Set();
  const queue = [startUrl];
  const allEmails = new Map(); // email -> {score, source}

  while (queue.length && result.pagesTried < CONFIG.MAX_TOTAL_PAGES_PER_SITE && allEmails.size < CONFIG.MAX_EMAILS_PER_SITE) {
    const url = queue.shift();
    if (!url || visited.has(url)) continue;
    visited.add(url);
    result.pagesTried++;

    const page = await fetchHtml(url);
    if (!page) continue;

    // Extract emails from this page
    const pageEmails = parseEmailsFromHtml(page.html, page.url || url);
    for (const e of pageEmails) {
      if (!allEmails.has(e)) {
        allEmails.set(e, { score: scoreEmail(e), source: page.url || url });
      }
    }

    // If we already have strong hits, stop early
    if (allEmails.size > 0) {
      const best = Math.max(...Array.from(allEmails.values()).map(v => v.score));
      if (best >= 10) break;
    }

    // Discover more links to crawl
    if (url === startUrl || visited.size === 1) {
      const contactLinks = findContactLinks(page.html, page.url || url);
      for (const l of contactLinks) {
        if (!visited.has(l) && !queue.includes(l)) queue.push(l);
      }
      // Add sitemap contact URLs if nothing useful yet
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

  // Sort by score
  const sorted = Array.from(allEmails.entries())
    .sort((a, b) => b[1].score - a[1].score)
    .slice(0, CONFIG.MAX_EMAILS_PER_SITE)
    .map(([email, meta]) => ({ email, ...meta }));

  result.emails = sorted.map(s => s.email);
  if (sorted.length) {
    result.emailSource = sorted[0].source;
  }
  return result;
}

// ==================== GOOGLE SEARCH FALLBACK (find website when missing) ====================

async function googleSearchForWebsite(businessName, address) {
  if (!CONFIG.USE_GOOGLE_SEARCH_FALLBACK) return null;
  try {
    const q = encodeURIComponent(`${businessName} ${address} official website`);
    const res = await axios.get(`${CONFIG.GOOGLE_SEARCH_URL}?q=${q}`, {
      timeout: CONFIG.FETCH_TIMEOUT_MS,
      headers: HEADERS,
      maxRedirects: 3
    });
    const $ = cheerio.load(res.data);
    // Google blocks scrapers often; parse whatever we can
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
    const response = await axios.get(DETAILS_URL, { params, timeout: 30000 });
    return response.data.result || {};
  } catch (err) {
    console.error(`  ❌ Details failed for ${placeId}: ${err.message}`);
    return {};
  }
}

// ==================== SEARCH / PAGINATION ====================

async function collectPlaces(query, location, apiKey, maxResults) {
  const results = [];
  const seen = new Set();
  const queryFormats = [
    `${query} in ${location}`,
    `${query} near ${location}`,
    `${query} ${location}`
  ];

  for (const q of queryFormats) {
    if (results.length >= maxResults) break;
    let pageToken = null;
    let page = 0;
    const MAX_PAGES_PER_FORMAT = 3;

    do {
      if (page >= MAX_PAGES_PER_FORMAT) break;
      const params = { key: apiKey, query: q };
      if (pageToken) params.pagetoken = pageToken;

      try {
        const response = await axios.get(TEXT_SEARCH_URL, { params, timeout: 30000 });
        const data = response.data;
        const status = data.status;

        if (status !== 'OK' && status !== 'ZERO_RESULTS') {
          console.log(`  ❌ ${status} for "${q}": ${data.error_message || ''}`);
          break;
        }

        for (const item of data.results || []) {
          const pid = item.place_id;
          if (pid && !seen.has(pid)) {
            seen.add(pid);
            results.push(item);
            if (results.length >= maxResults) break;
          }
        }

        pageToken = data.next_page_token || null;
        // Google requires a short delay before the token is valid
        if (pageToken) await new Promise(r => setTimeout(r, 2000));
        page++;
      } catch (err) {
        console.log(`  ❌ Request failed: ${err.message}`);
        break;
      }
    } while (pageToken && results.length < maxResults);
  }

  return results.slice(0, maxResults);
}

// ==================== MAIN HANDLER ====================

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const API_KEY = process.env.GOOGLE_PLACES_API_KEY;
    if (!API_KEY) {
      return res.status(500).json({ error: 'GOOGLE_PLACES_API_KEY not configured' });
    }

    const {
      query, location, maxResults = 100, subAreas = '', skipEmails = false,
      skipWebsiteLookup = false // NEW: skip Google-search fallback
    } = req.body || {};

    if (!query || !location) {
      return res.status(400).json({ error: 'query and location are required' });
    }

    const cleanLocation = location.trim();
    const locations = [cleanLocation];
    if (subAreas && typeof subAreas === 'string') {
      locations.push(...subAreas.split(',').map(s => s.trim()).filter(Boolean));
    }
    const maxResultsNum = Math.min(parseInt(maxResults) || 100, 2000);
    const perLocationCap = Math.max(1, Math.ceil(maxResultsNum / locations.length));

    // -------- 1. Collect places --------
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

    if (finalPlaces.length === 0) {
      return res.status(200).json({
        success: true, total: 0, emailsFound: 0, emailPercentage: 0,
        data: [], excel: null, message: 'No businesses found.'
      });
    }

    // -------- 2. Fetch details + emails (with concurrency) --------
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
          // Try Google search fallback to find the site
          const guessed = await googleSearchForWebsite(
            details.name || place.name,
            details.formatted_address || place.formatted_address || cleanLocation
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
        website: website,
        email: email,
        emailSource: emailSource,
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

    // Run in batches of CONFIG.CONCURRENCY_FOR_SITES
    for (let i = 0; i < finalPlaces.length; i += CONFIG.CONCURRENCY_FOR_SITES) {
      const batch = finalPlaces.slice(i, i + CONFIG.CONCURRENCY_FOR_SITES);
      await Promise.all(batch.map((p, j) => processOne(p, i + j)));
    }

    const validRows = rows.filter(Boolean);
    const totalWithEmail = validRows.filter(r => r.email).length;
    const emailPercentage = validRows.length
      ? Math.round((totalWithEmail / validRows.length) * 100) : 0;

    // -------- 3. Build Excel --------
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
    console.error('SCRAPING ERROR:', error);
    return res.status(500).json({ error: error.message || 'Scraping failed' });
  }
};