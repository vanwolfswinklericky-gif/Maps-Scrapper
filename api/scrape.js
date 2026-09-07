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
  /info@/i,
  /contact@/i,
  /hello@/i,
  /support@/i,
  /sales@/i,
  /admin@/i,
  /careers@/i,
  /team@/i,
  /founder@/i,
  /ceo@/i,
  /owner@/i,
  /manager@/i
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
  
  // Check blocklist
  for (const bad of EMAIL_BLOCKLIST) {
    if (lower.includes(bad.toLowerCase())) return false;
  }
  
  // Must have @ and a valid domain
  if (!lower.includes('@')) return false;
  
  const parts = lower.split('@');
  if (parts.length !== 2) return false;
  
  const domain = parts[1];
  if (!domain || domain.length < 3) return false;
  if (!domain.includes('.')) return false;
  
  // Check if it looks like a real email (not too long, not too short)
  if (lower.length < 5 || lower.length > 100) return false;
  
  return true;
}

// Helper: Clean email
function cleanEmail(email) {
  if (!email) return '';
  
  // Remove common prefixes
  let cleaned = email
    .replace(/^mailto:/i, '')
    .replace(/^email:/i, '')
    .replace(/^e-mail:/i, '')
    .replace(/\?subject=.*$/i, '')
    .replace(/\?body=.*$/i, '')
    .trim();
  
  // Remove quotes
  cleaned = cleaned.replace(/^["']|["']$/g, '');
  
  return cleaned;
}

// Helper: Extract emails from text
function extractEmails(text) {
  if (!text) return [];
  
  const emails = new Set();
  
  // Try strict regex first
  let matches = text.match(EMAIL_REGEX_STRICT) || [];
  matches.forEach(m => {
    const cleaned = cleanEmail(m);
    if (isValidEmail(cleaned)) {
      emails.add(cleaned.toLowerCase());
    }
  });
  
  // Try general regex
  matches = text.match(EMAIL_REGEX) || [];
  matches.forEach(m => {
    const cleaned = cleanEmail(m);
    if (isValidEmail(cleaned)) {
      emails.add(cleaned.toLowerCase());
    }
  });
  
  return Array.from(emails);
}

// Helper: Find email on a page with multiple strategies
async function findEmailOnPage(url) {
  try {
    const response = await axios.get(url, {
      timeout: 10000,
      headers: HEADERS,
      maxRedirects: 5
    });
    
    if (response.status !== 200) return null;
    
    const $ = cheerio.load(response.data);
    const emails = new Set();
    
    // Strategy 1: Check mailto links (most reliable)
    $('a[href^="mailto:"]').each((i, el) => {
      const href = $(el).attr('href');
      if (href) {
        const email = cleanEmail(href.replace('mailto:', ''));
        if (isValidEmail(email)) {
          emails.add(email.toLowerCase());
        }
      }
    });
    
    // Strategy 2: Check all links for email patterns
    $('a[href]').each((i, el) => {
      const href = $(el).attr('href');
      if (href) {
        // Check if href contains email
        const extracted = extractEmails(href);
        extracted.forEach(e => emails.add(e));
      }
    });
    
    // Strategy 3: Check visible text
    const bodyText = $('body').text();
    const extracted = extractEmails(bodyText);
    extracted.forEach(e => emails.add(e));
    
    // Strategy 4: Check specific elements that often contain emails
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
    
    // Strategy 5: Check script and style tags for email (sometimes encoded)
    $('script, style').each((i, el) => {
      const content = $(el).html() || '';
      const extracted = extractEmails(content);
      extracted.forEach(e => {
        // Only add if it looks like a real email (not encoded)
        if (e.includes('@') && !e.includes('{') && !e.includes('}')) {
          emails.add(e);
        }
      });
    });
    
    // Strategy 6: Check for email in data attributes
    $('[data-email], [data-contact], [data-info]').each((i, el) => {
      const dataEmail = $(el).attr('data-email') || $(el).attr('data-contact') || '';
      if (dataEmail) {
        const extracted = extractEmails(dataEmail);
        extracted.forEach(e => emails.add(e));
      }
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
      
      // Also check if it contains the business name (we don't have business name here)
      // So we'll just prioritize non-generic emails
      if (isGeneric) {
        genericEmails.push(email);
      } else {
        priorityEmails.push(email);
      }
    });
    
    // Return priority emails first, then generic
    const allEmails = [...priorityEmails, ...genericEmails];
    return allEmails.length > 0 ? allEmails[0] : null;
    
  } catch (error) {
    return null;
  }
}

// Helper: Find contact pages (expanded)
async function findContactLinks(baseUrl) {
  try {
    const response = await axios.get(baseUrl, {
      timeout: 10000,
      headers: HEADERS,
      maxRedirects: 5
    });
    
    if (response.status !== 200) return [];
    
    const $ = cheerio.load(response.data);
    const links = [];
    const seenUrls = new Set();
    
    // Find all links
    $('a[href]').each((i, el) => {
      const href = $(el).attr('href');
      if (!href) return;
      
      const hrefLower = href.toLowerCase();
      const text = $(el).text().toLowerCase();
      
      // Check if it's a contact-related link
      const isContact = CONTACT_PATHS.some(path => 
        hrefLower.includes(path) || text.includes(path)
      );
      
      if (isContact) {
        try {
          const fullUrl = new URL(href, baseUrl).toString();
          if (!seenUrls.has(fullUrl)) {
            seenUrls.add(fullUrl);
            links.push(fullUrl);
          }
        } catch (e) {
          // Invalid URL, skip
        }
      }
    });
    
    // Return up to 5 contact pages (more chances to find email)
    return links.slice(0, 5);
  } catch (error) {
    return [];
  }
}

// Helper: Find email for business with aggressive strategy
async function findEmailForBusiness(website) {
  if (!website) return '';
  
  console.log(`  📧 Searching for email on: ${website}`);
  
  // Try main page first
  let email = await findEmailOnPage(website);
  if (email) {
    console.log(`    ✅ Found email on main page: ${email}`);
    return email;
  }
  
  // Try contact pages
  console.log(`    🔍 Checking contact pages...`);
  const contactLinks = await findContactLinks(website);
  
  for (const link of contactLinks) {
    console.log(`      📄 Checking: ${link}`);
    email = await findEmailOnPage(link);
    if (email) {
      console.log(`    ✅ Found email on contact page: ${email}`);
      return email;
    }
  }
  
  console.log(`    ❌ No email found`);
  return '';
}

// Helper: Get place details
async function getPlaceDetails(placeId, apiKey) {
  const params = {
    place_id: placeId,
    fields: 'name,formatted_address,formatted_phone_number,international_phone_number,website,rating,user_ratings_total,url,types',
    key: apiKey
  };
  
  const response = await axios.get(DETAILS_URL, { params, timeout: 30000 });
  return response.data.result || {};
}

// Helper: Collect places via text search
async function collectPlaces(query, location, apiKey, maxResults) {
  const results = [];
  let pageToken = null;
  const seenIds = new Set();
  
  while (results.length < maxResults) {
    const params = { key: apiKey };
    if (pageToken) {
      params.pagetoken = pageToken;
    } else {
      params.query = `${query} in ${location}`;
    }
    
    const response = await axios.get(TEXT_SEARCH_URL, { params, timeout: 30000 });
    const data = response.data;
    const status = data.status;
    
    if (status !== 'OK' && status !== 'ZERO_RESULTS') {
      throw new Error(`Places API error: ${status} - ${data.error_message || ''}`);
    }
    
    for (const item of data.results || []) {
      const pid = item.place_id;
      if (pid && !seenIds.has(pid)) {
        seenIds.add(pid);
        results.push(item);
        if (results.length >= maxResults) break;
      }
    }
    
    pageToken = data.next_page_token;
    if (!pageToken) break;
    
    await new Promise(resolve => setTimeout(resolve, 2200));
  }
  
  return results.slice(0, maxResults);
}

// Main handler
module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }
  
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }
  
  try {
    const API_KEY = process.env.GOOGLE_PLACES_API_KEY;
    
    if (!API_KEY) {
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
    
    if (!query || !location) {
      return res.status(400).json({ 
        error: 'Missing required fields: query and location are required' 
      });
    }
    
    let areas = [location];
    if (subAreas && subAreas.trim()) {
      const extraAreas = subAreas.split(',').map(s => s.trim()).filter(Boolean);
      areas = [location, ...extraAreas];
    }
    
    const maxResultsNum = Math.min(parseInt(maxResults) || 100, 2000);
    
    console.log(`🔍 Searching for: ${query}`);
    console.log(`📍 Locations: ${areas.join(', ')}`);
    console.log(`📊 Target: ${maxResultsNum} businesses`);
    console.log(`📧 Email lookup: ${skipEmails ? 'SKIPPED' : 'ENABLED (AGGRESSIVE)'}`);
    
    // Collect all places
    const allPlaces = [];
    const seenIds = new Set();
    
    if (areas.length === 1) {
      const perAreaCap = Math.min(maxResultsNum, 60);
      
      console.log(`Searching in ${location}...`);
      const found = await collectPlaces(query, location, API_KEY, perAreaCap);
      const newPlaces = found.filter(p => !seenIds.has(p.place_id));
      newPlaces.forEach(p => seenIds.add(p.place_id));
      allPlaces.push(...newPlaces);
      
      // Try variations
      if (allPlaces.length < maxResultsNum) {
        const variations = [
          `${query} near ${location}`,
          `${query} in ${location} area`,
          `best ${query} ${location}`,
          `${query} ${location}`,
          `${query} services ${location}`,
          `${query} providers ${location}`
        ];
        
        for (const variation of variations) {
          if (allPlaces.length >= maxResultsNum) break;
          
          console.log(`  Trying variation: "${variation}"...`);
          const extraResults = await collectPlaces(variation, location, API_KEY, 60);
          const newExtra = extraResults.filter(p => !seenIds.has(p.place_id));
          newExtra.forEach(p => seenIds.add(p.place_id));
          allPlaces.push(...newExtra);
          
          if (allPlaces.length >= maxResultsNum) break;
        }
      }
    } else {
      const perAreaCap = Math.max(1, Math.ceil(maxResultsNum / areas.length));
      
      for (const area of areas) {
        if (allPlaces.length >= maxResultsNum) break;
        
        console.log(`Searching in ${area}...`);
        const found = await collectPlaces(query, area, API_KEY, perAreaCap);
        const newPlaces = found.filter(p => !seenIds.has(p.place_id));
        newPlaces.forEach(p => seenIds.add(p.place_id));
        allPlaces.push(...newPlaces);
        console.log(`  Found ${newPlaces.length} new places (total: ${allPlaces.length})`);
      }
    }
    
    const finalPlaces = allPlaces.slice(0, maxResultsNum);
    console.log(`✅ Collected ${finalPlaces.length} unique businesses`);
    
    // Process each place with email priority
    const rows = [];
    const total = finalPlaces.length;
    let emailsFound = 0;
    let websitesWithEmail = 0;
    
    for (let i = 0; i < total; i++) {
      const place = finalPlaces[i];
      const details = await getPlaceDetails(place.place_id, API_KEY);
      
      await new Promise(resolve => setTimeout(resolve, 50));
      
      const website = details.website || '';
      let email = '';
      
      if (website && !skipEmails) {
        email = await findEmailForBusiness(website);
        if (email) {
          emailsFound++;
          if (website) websitesWithEmail++;
        }
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
        console.log(`  Processed ${i + 1}/${total}... Emails found: ${emailsFound}`);
      }
    }
    
    // Create Excel file with email emphasis
    const wb = XLSX.utils.book_new();
    
    // Main sheet with all data
    const wsData = [
      ['Name', 'Email', 'Phone', 'Website', 'Category', 'Address', 'Rating', 'Review Count', 'Google Maps Link'],
      ...rows.map(r => [r.name, r.email, r.phone, r.website, r.category, r.address, r.rating, r.reviewCount, r.mapsLink])
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, 'All Leads');
    
    // Create separate sheet for emails only
    const emailRows = rows.filter(r => r.email);
    if (emailRows.length > 0) {
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
    
    console.log(`✅ COMPLETE!`);
    console.log(`  Total businesses: ${rows.length}`);
    console.log(`  Emails found: ${totalWithEmail} (${emailPercentage}%)`);
    console.log(`  Businesses with websites: ${rows.filter(r => r.website).length}`);
    
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
    console.error('❌ Scraping error:', error);
    res.status(500).json({ 
      error: error.message || 'An error occurred during scraping'
    });
  }
};