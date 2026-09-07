// api/scrape.js
const axios = require('axios');
const cheerio = require('cheerio');
const XLSX = require('xlsx');

// Constants
const TEXT_SEARCH_URL = "https://maps.googleapis.com/maps/api/place/textsearch/json";
const DETAILS_URL = "https://maps.googleapis.com/maps/api/place/details/json";

// More aggressive email regex
const EMAIL_REGEX = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;
const EMAIL_REGEX_STRICT = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.(com|org|net|edu|gov|co|uk|ng|io|ai|me|us|ca|au|de|fr|es|it|br|in|jp|cn|ru|za|mx|nl|se|no|fi|dk|ch|at|be|pl|cz|gr|pt|ie|nz|sg|hk|my|ph|vn|th|id|tr|il|ae|sa|eg|ke|gh|za)/gi;

// Expanded blocklist - only block obvious junk
const EMAIL_BLOCKLIST = [
  'sentry.io', 'wixpress.com', 'example.com', 'godaddy.com', 
  'domain.com', 'yourname@', 'info@domain', 'contact@domain',
  'admin@domain', 'support@domain', 'sales@domain',
  '.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp', '.css', '.js',
  'data:image', 'javascript:', 'mailto:'
];

// Contact page keywords (expanded)
const CONTACT_PATHS = [
  'contact', 'about', 'about-us', 'contact-us', 'reach-us', 
  'get-in-touch', 'touch', 'connect', 'team', 'our-team',
  'support', 'help', 'customer-service', 'inquiry', 'enquiries',
  'reach', 'contact-page', 'contactus', 'aboutus'
];

// Common email patterns to check
const EMAIL_PATTERNS = [
  /info@/i, /contact@/i, /hello@/i, /support@/i, /sales@/i,
  /admin@/i, /careers@/i, /team@/i, /founder@/i, /ceo@/i,
  /owner@/i, /manager@/i, /reservation@/i, /bookings@/i
];

// Headers to mimic real browser
const HEADERS = {
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
  'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
  'Accept-Language': 'en-US,en;q=0.5',
  'Accept-Encoding': 'gzip, deflate, br',
  'Connection': 'keep-alive',
  'Upgrade-Insecure-Requests': '1'
};

// Helper: Check if email is valid and not blocked
function isValidEmail(email) {
  if (!email) return false;
  const lower = email.toLowerCase().trim();
  for (const bad of EMAIL_BLOCKLIST) {
    if (lower.includes(bad.toLowerCase())) return false;
  }
  if (!lower.includes('@')) return false;
  const parts = lower.split('@');
  if (parts.length !== 2) return false;
  const domain = parts[1];
  if (!domain || domain.length < 3 || !domain.includes('.')) return false;
  if (lower.length < 5 || lower.length > 100) return false;
  return true;
}

// Helper: Clean email
function cleanEmail(email) {
  if (!email) return '';
  let cleaned = email
    .replace(/^mailto:/i, '')
    .replace(/^email:/i, '')
    .replace(/^e-mail:/i, '')
    .replace(/\?subject=.*$/i, '')
    .replace(/\?body=.*$/i, '')
    .trim();
  cleaned = cleaned.replace(/^["']|["']$/g, '');
  return cleaned;
}

// Helper: Extract emails from text
function extractEmails(text) {
  if (!text) return [];
  const emails = new Set();
  let matches = text.match(EMAIL_REGEX_STRICT) || [];
  matches.forEach(m => {
    const cleaned = cleanEmail(m);
    if (isValidEmail(cleaned)) emails.add(cleaned.toLowerCase());
  });
  matches = text.match(EMAIL_REGEX) || [];
  matches.forEach(m => {
    const cleaned = cleanEmail(m);
    if (isValidEmail(cleaned)) emails.add(cleaned.toLowerCase());
  });
  return Array.from(emails);
}

// Helper: Find email on a page with multiple strategies
async function findEmailOnPage(url) {
  try {
    console.log(`      📄 Fetching: ${url}`);
    const response = await axios.get(url, {
      timeout: 10000,
      headers: HEADERS,
      maxRedirects: 5
    });
    if (response.status !== 200) {
      console.log(`      ⚠️ Status ${response.status} for ${url}`);
      return null;
    }
    const $ = cheerio.load(response.data);
    const emails = new Set();

    // Strategy 1: Check mailto links
    $('a[href^="mailto:"]').each((i, el) => {
      const href = $(el).attr('href');
      if (href) {
        const email = cleanEmail(href.replace('mailto:', ''));
        if (isValidEmail(email)) {
          emails.add(email.toLowerCase());
          console.log(`        📧 Found mailto email: ${email}`);
        }
      }
    });

    // Strategy 2: Check all links for email patterns
    $('a[href]').each((i, el) => {
      const href = $(el).attr('href');
      if (href) {
        const extracted = extractEmails(href);
        extracted.forEach(e => emails.add(e));
      }
    });

    // Strategy 3: Check visible text
    const bodyText = $('body').text();
    const extracted = extractEmails(bodyText);
    extracted.forEach(e => emails.add(e));

    // Strategy 4: Check specific elements
    const selectors = [
      'p', 'span', 'div', 'li', 'td', 'a',
      '.email', '.contact', '.info', '.address',
      '#email', '#contact', '#info',
      '[class*="email"]', '[class*="contact"]',
      '[id*="email"]', '[id*="contact"]'
    ];
    selectors.forEach(selector => {
      $(selector).each((i, el) => {
        const text = $(el).text();
        const extracted = extractEmails(text);
        extracted.forEach(e => emails.add(e));
      });
    });

    // Strategy 5: Check script and style tags
    $('script, style').each((i, el) => {
      const content = $(el).html() || '';
      const extracted = extractEmails(content);
      extracted.forEach(e => {
        if (e.includes('@') && !e.includes('{') && !e.includes('}')) {
          emails.add(e);
        }
      });
    });

    // Priority: Prefer business emails over generic
    const priorityEmails = [];
    const genericEmails = [];
    emails.forEach(email => {
      let isGeneric = false;
      for (const pattern of EMAIL_PATTERNS) {
        if (pattern.test(email)) {
          isGeneric = true;
          break;
        }
      }
      if (isGeneric) {
        genericEmails.push(email);
      } else {
        priorityEmails.push(email);
      }
    });

    const allEmails = [...priorityEmails, ...genericEmails];
    console.log(`      Found ${allEmails.length} total emails on page`);
    return allEmails.length > 0 ? allEmails[0] : null;
  } catch (error) {
    console.log(`      ❌ Error fetching ${url}: ${error.message}`);
    return null;
  }
}

// Helper: Find contact pages
async function findContactLinks(baseUrl) {
  try {
    console.log(`      🔍 Looking for contact pages on ${baseUrl}`);
    const response = await axios.get(baseUrl, {
      timeout: 10000,
      headers: HEADERS,
      maxRedirects: 5
    });
    if (response.status !== 200) {
      console.log(`      ⚠️ Status ${response.status} for ${baseUrl}`);
      return [];
    }
    const $ = cheerio.load(response.data);
    const links = [];
    const seenUrls = new Set();
    $('a[href]').each((i, el) => {
      const href = $(el).attr('href');
      if (!href) return;
      const hrefLower = href.toLowerCase();
      const text = $(el).text().toLowerCase();
      const isContact = CONTACT_PATHS.some(path => 
        hrefLower.includes(path) || text.includes(path)
      );
      if (isContact) {
        try {
          const fullUrl = new URL(href, baseUrl).toString();
          if (!seenUrls.has(fullUrl)) {
            seenUrls.add(fullUrl);
            links.push(fullUrl);
            console.log(`        Found contact page: ${fullUrl}`);
          }
        } catch (e) { /* Invalid URL, skip */ }
      }
    });
    console.log(`      Found ${links.length} contact pages`);
    return links.slice(0, 5);
  } catch (error) {
    console.log(`      ❌ Error finding contact pages: ${error.message}`);
    return [];
  }
}

// Helper: Find email for business
async function findEmailForBusiness(website) {
  if (!website) return '';
  console.log(`  📧 Searching for email on: ${website}`);
  let email = await findEmailOnPage(website);
  if (email) {
    console.log(`    ✅ Found email on main page: ${email}`);
    return email;
  }
  console.log(`    🔍 No email on main page, checking contact pages...`);
  const contactLinks = await findContactLinks(website);
  for (const link of contactLinks) {
    email = await findEmailOnPage(link);
    if (email) {
      console.log(`    ✅ Found email on contact page: ${email}`);
      return email;
    }
  }
  console.log(`    ❌ No email found anywhere on ${website}`);
  return '';
}

// Helper: Get place details
async function getPlaceDetails(placeId, apiKey) {
  const params = {
    place_id: placeId,
    fields: 'name,formatted_address,formatted_phone_number,international_phone_number,website,rating,user_ratings_total,url,types',
    key: apiKey
  };
  console.log(`  📋 Fetching details for place ID: ${placeId}`);
  const response = await axios.get(DETAILS_URL, { params, timeout: 30000 });
  const result = response.data.result || {};
  console.log(`    Name: ${result.name || 'Unknown'}`);
  console.log(`    Website: ${result.website || 'None'}`);
  return result;
}

// Helper: Generate expanded search locations
function generateSearchLocations(mainLocation, subAreas, maxLocations = 25) {
  const locations = [];
  
  // Add main location
  locations.push(mainLocation);
  
  // Add sub-areas
  if (subAreas && subAreas.trim()) {
    const areas = subAreas.split(',').map(s => s.trim()).filter(Boolean);
    locations.push(...areas);
  }
  
  // If we need more locations, add variations
  if (locations.length < maxLocations) {
    const variations = [
      `${mainLocation} area`,
      `${mainLocation} suburbs`,
      `${mainLocation} metro`,
      `Greater ${mainLocation}`
    ];
    for (const varLoc of variations) {
      if (locations.length >= maxLocations) break;
      if (!locations.includes(varLoc)) {
        locations.push(varLoc);
      }
    }
  }
  
  return locations.slice(0, maxLocations);
}

// Helper: Collect places with improved pagination
async function collectPlaces(query, location, apiKey, maxResults) {
  console.log(`\n📥 Starting search: "${query}" in "${location}"`);
  console.log(`   Max results: ${maxResults}`);
  
  const results = [];
  let pageToken = null;
  const seenIds = new Set();
  
  // Try different query formats
  const queryFormats = [
    `${query} in ${location}`,
    `${query} near ${location}`,
    `${query} ${location}`,
    `${query}+${location}`
  ];
  
  let currentFormatIndex = 0;
  let consecutiveFailures = 0;
  let totalAttempts = 0;
  
  while (results.length < maxResults && currentFormatIndex < queryFormats.length) {
    const params = { key: apiKey };
    
    if (pageToken) {
      params.pagetoken = pageToken;
      console.log(`  📄 Fetching next page (attempt ${consecutiveFailures + 1})...`);
    } else {
      params.query = queryFormats[currentFormatIndex];
      console.log(`  🔍 Trying format ${currentFormatIndex + 1}: "${params.query}"`);
      consecutiveFailures = 0;
    }
    
    totalAttempts++;
    
    try {
      console.log(`  📤 Sending request to Google Places API...`);
      console.log(`     URL: ${TEXT_SEARCH_URL}`);
      console.log(`     Params: ${JSON.stringify(params, null, 2)}`);
      
      const response = await axios.get(TEXT_SEARCH_URL, { params, timeout: 30000 });
      const data = response.data;
      const status = data.status;
      
      console.log(`  📥 Response status: ${status}`);
      
      if (status === 'OK') {
        const items = data.results || [];
        console.log(`  ✅ Found ${items.length} results on this page`);
        
        let newItems = 0;
        for (const item of items) {
          const pid = item.place_id;
          if (pid && !seenIds.has(pid)) {
            seenIds.add(pid);
            results.push(item);
            newItems++;
            if (results.length >= maxResults) break;
          }
        }
        console.log(`  📊 Added ${newItems} new businesses (total: ${results.length})`);
        
        // Handle pagination
        if (data.next_page_token && results.length < maxResults) {
          console.log(`  📄 Next page token received, waiting 3 seconds...`);
          // Wait at least 3 seconds for token to activate
          await new Promise(resolve => setTimeout(resolve, 3000));
          pageToken = data.next_page_token;
          consecutiveFailures = 0;
        } else {
          console.log(`  📄 No more pages for this format`);
          currentFormatIndex++;
          pageToken = null;
          consecutiveFailures = 0;
        }
        
      } else if (status === 'ZERO_RESULTS') {
        console.log(`  ⚠️ No results found for format: "${params.query}"`);
        currentFormatIndex++;
        pageToken = null;
        consecutiveFailures = 0;
      } else {
        console.log(`  ❌ Format "${params.query}" failed with status: ${status}`);
        console.log(`     Error message: ${data.error_message || 'No error message'}`);
        
        if (pageToken) {
          consecutiveFailures++;
          console.log(`  ⚠️ Page token failure #${consecutiveFailures}`);
          if (consecutiveFailures >= 2) {
            console.log(`  ⏭️ Skipping to next format after multiple failures`);
            currentFormatIndex++;
            pageToken = null;
            consecutiveFailures = 0;
          } else {
            console.log(`  ⏳ Waiting 5 seconds and retrying token...`);
            await new Promise(resolve => setTimeout(resolve, 5000));
          }
        } else {
          currentFormatIndex++;
          pageToken = null;
          consecutiveFailures = 0;
        }
      }
    } catch (error) {
      console.log(`  ❌ Request failed: ${error.message}`);
      if (error.response) {
        console.log(`     Response status: ${error.response.status}`);
        console.log(`     Response data: ${JSON.stringify(error.response.data, null, 2)}`);
      }
      if (pageToken) {
        consecutiveFailures++;
        if (consecutiveFailures >= 2) {
          currentFormatIndex++;
          pageToken = null;
          consecutiveFailures = 0;
        }
      } else {
        currentFormatIndex++;
        pageToken = null;
        consecutiveFailures = 0;
      }
    }
  }
  
  console.log(`\n📊 Search complete for "${location}"`);
  console.log(`   Total businesses found: ${results.length}`);
  console.log(`   Query formats tried: ${totalAttempts}`);
  
  return results.slice(0, maxResults);
}

// Main handler
module.exports = async (req, res) => {
  console.log('\n🚀 ===== NEW REQUEST =====');
  console.log(`📅 Time: ${new Date().toISOString()}`);
  
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  
  if (req.method === 'OPTIONS') {
    console.log('📡 OPTIONS request received');
    return res.status(200).end();
  }
  
  if (req.method !== 'POST') {
    console.log(`❌ Method not allowed: ${req.method}`);
    return res.status(405).json({ error: 'Method not allowed' });
  }
  
  try {
    const API_KEY = process.env.GOOGLE_PLACES_API_KEY;
    console.log(`🔑 API Key present: ${API_KEY ? 'Yes (starts with ' + API_KEY.substring(0, 8) + '...)' : 'No'}`);
    
    if (!API_KEY) {
      console.error('❌ API key not configured');
      return res.status(500).json({ 
        error: 'API key not configured. Please set GOOGLE_PLACES_API_KEY environment variable.' 
      });
    }
    
    const { 
      query, 
      location, 
      maxResults = 100, 
      subAreas = '', 
      skipEmails = false 
    } = req.body;
    
    console.log(`📋 Request parameters:`);
    console.log(`   Query: "${query}"`);
    console.log(`   Location: "${location}"`);
    console.log(`   Sub-Areas: "${subAreas}"`);
    console.log(`   Max Results: ${maxResults}`);
    console.log(`   Skip Emails: ${skipEmails}`);
    
    if (!query || !location) {
      console.error('❌ Missing required fields');
      return res.status(400).json({ 
        error: 'Missing required fields: query and location are required' 
      });
    }
    
    const cleanLocation = location.trim();
    console.log(`📍 Clean location: "${cleanLocation}"`);
    
    // Generate expanded search locations
    const searchLocations = generateSearchLocations(cleanLocation, subAreas, 25);
    console.log(`📍 Generated ${searchLocations.length} search locations`);
    console.log(`   ${searchLocations.join(', ')}`);
    
    const maxResultsNum = Math.min(parseInt(maxResults) || 100, 2000);
    console.log(`🎯 Target: ${maxResultsNum} businesses total`);
    console.log(`📧 Email lookup: ${skipEmails ? 'SKIPPED' : 'ENABLED (AGGRESSIVE)'}`);
    
    // Collect all places from multiple locations
    const allPlaces = [];
    const seenIds = new Set();
    const perLocationCap = Math.max(1, Math.ceil(maxResultsNum / searchLocations.length));
    
    for (const area of searchLocations) {
      if (allPlaces.length >= maxResultsNum) break;
      
      const remaining = maxResultsNum - allPlaces.length;
      const cap = Math.min(perLocationCap, remaining, 60);
      console.log(`\n🔍 Searching in "${area}" (cap: ${cap})...`);
      
      const found = await collectPlaces(query, area, API_KEY, cap);
      const newPlaces = found.filter(p => !seenIds.has(p.place_id));
      newPlaces.forEach(p => seenIds.add(p.place_id));
      allPlaces.push(...newPlaces);
      
      console.log(`   Found ${newPlaces.length} new businesses in "${area}"`);
      console.log(`   Running total: ${allPlaces.length}`);
    }
    
    const finalPlaces = allPlaces.slice(0, maxResultsNum);
    console.log(`\n✅ Collected ${finalPlaces.length} unique businesses total`);
    
    if (finalPlaces.length === 0) {
      console.warn('⚠️ No businesses found!');
      return res.status(200).json({
        success: true,
        total: 0,
        emailsFound: 0,
        emailPercentage: 0,
        data: [],
        excel: null,
        message: 'No businesses found. Try a different location or query.'
      });
    }
    
    // Process each place with email priority
    const rows = [];
    const total = finalPlaces.length;
    let emailsFound = 0;
    
    console.log(`\n📊 Processing ${total} businesses...`);
    console.log(`📧 Email lookup: ${skipEmails ? 'SKIPPED' : 'ENABLED'}`);
    
    for (let i = 0; i < total; i++) {
      const place = finalPlaces[i];
      console.log(`\n🔍 [${i+1}/${total}] Processing business...`);
      
      const details = await getPlaceDetails(place.place_id, API_KEY);
      await new Promise(resolve => setTimeout(resolve, 50));
      
      const website = details.website || '';
      let email = '';
      
      if (website && !skipEmails) {
        console.log(`  🌐 Website found: ${website}`);
        email = await findEmailForBusiness(website);
        if (email) {
          emailsFound++;
          console.log(`  ✅ Email found: ${email}`);
        } else {
          console.log(`  ❌ No email found for this business`);
        }
      } else if (!website) {
        console.log(`  ⚠️ No website found for this business`);
      } else if (skipEmails) {
        console.log(`  ⏭️ Email lookup skipped`);
      }
      
      rows.push({
        name: details.name || place.name || '',
        category: (details.types || []).slice(0, 2).join(', '),
        phone: details.formatted_phone_number || '',
        website: website,
        email: email || '',
        address: details.formatted_address || '',
        rating: details.rating || '',
        reviewCount: details.user_ratings_total || '',
        mapsLink: details.url || ''
      });
      
      // Log progress with email stats
      if ((i + 1) % 10 === 0 || i === total - 1) {
        console.log(`\n📊 Progress: ${i + 1}/${total}`);
        console.log(`   Emails found: ${emailsFound}`);
        console.log(`   Rate: ${Math.round((emailsFound / (i + 1)) * 100)}%`);
      }
    }
    
    // Create Excel file
    console.log('\n📁 Creating Excel file...');
    const wb = XLSX.utils.book_new();
    
    const wsData = [
      ['Name', 'Email', 'Phone', 'Website', 'Category', 'Address', 'Rating', 'Review Count', 'Google Maps Link'],
      ...rows.map(r => [r.name, r.email, r.phone, r.website, r.category, r.address, r.rating, r.reviewCount, r.mapsLink])
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, 'All Leads');
    
    const emailRows = rows.filter(r => r.email);
    if (emailRows.length > 0) {
      console.log(`📧 Creating emails-only sheet with ${emailRows.length} entries`);
      const emailData = [
        ['Name', 'Email', 'Phone', 'Website', 'Category', 'Address'],
        ...emailRows.map(r => [r.name, r.email, r.phone, r.website, r.category, r.address])
      ];
      const wsEmails = XLSX.utils.aoa_to_sheet(emailData);
      XLSX.utils.book_append_sheet(wb, wsEmails, 'Emails Only');
    }
    
    const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const excelBase64 = excelBuffer.toString('base64');
    
    const totalWithEmail = rows.filter(r => r.email).length;
    const emailPercentage = total > 0 ? Math.round((totalWithEmail / total) * 100) : 0;
    
    console.log('\n✅ ===== COMPLETE =====');
    console.log(`  Total businesses: ${rows.length}`);
    console.log(`  Emails found: ${totalWithEmail} (${emailPercentage}%)`);
    console.log(`  Businesses with websites: ${rows.filter(r => r.website).length}`);
    console.log(`  Excel file size: ${Math.round(excelBuffer.length / 1024)} KB`);
    console.log('=====================\n');
    
    res.status(200).json({
      success: true,
      total: rows.length,
      emailsFound: totalWithEmail,
      emailPercentage: emailPercentage,
      data: rows,
      excel: excelBase64,
      message: `✅ Scraped ${rows.length} businesses, found ${totalWithEmail} emails (${emailPercentage}%)`
    });
    
  } catch (error) {
    console.error('\n❌ ===== SCRAPING ERROR =====');
    console.error(`  Error: ${error.message}`);
    console.error(`  Stack: ${error.stack}`);
    console.error('=============================\n');
    
    res.status(500).json({ 
      error: error.message || 'An error occurred during scraping'
    });
  }
};