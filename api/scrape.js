// api/scrape.js
const axios = require('axios');
const cheerio = require('cheerio');
const XLSX = require('xlsx');

// Constants
const TEXT_SEARCH_URL = "https://maps.googleapis.com/maps/api/place/textsearch/json";
const DETAILS_URL = "https://maps.googleapis.com/maps/api/place/details/json";

const EMAIL_REGEX = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/;
const EMAIL_BLOCKLIST = ['sentry.io', 'wixpress.com', 'example.com', 'godaddy.com', '.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp', '@2x', 'yourname@', 'domain.com'];
const CONTACT_HINTS = ['contact', 'about', 'about-us', 'contact-us', 'reach-us'];

// Helper: Check if email is blocked
function isBlocked(email) {
  const lower = email.toLowerCase();
  return EMAIL_BLOCKLIST.some(bad => lower.includes(bad));
}

// Helper: Find email on a page
async function findEmailOnPage(url) {
  try {
    const response = await axios.get(url, {
      timeout: 8000,
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; LeadResearchBot/1.0)'
      }
    });
    
    if (response.status !== 200) return null;
    
    const $ = cheerio.load(response.data);
    
    // Check mailto links first
    const mailtoLinks = $('a[href^="mailto:"]');
    for (let i = 0; i < mailtoLinks.length; i++) {
      const href = $(mailtoLinks[i]).attr('href');
      if (href) {
        const candidate = href.split('mailto:')[1].split('?')[0].trim();
        if (candidate && !isBlocked(candidate)) {
          return candidate;
        }
      }
    }
    
    // Fallback: regex scan of visible text
    const text = $('body').text();
    const matches = text.match(EMAIL_REGEX);
    if (matches) {
      for (const match of matches) {
        if (!isBlocked(match)) {
          return match;
        }
      }
    }
    
    return null;
  } catch (error) {
    return null;
  }
}

// Helper: Find contact pages
async function findContactLinks(baseUrl) {
  try {
    const response = await axios.get(baseUrl, {
      timeout: 8000,
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; LeadResearchBot/1.0)'
      }
    });
    
    if (response.status !== 200) return [];
    
    const $ = cheerio.load(response.data);
    const links = [];
    
    $('a[href]').each((i, el) => {
      const href = $(el).attr('href');
      if (href) {
        const lowerHref = href.toLowerCase();
        if (CONTACT_HINTS.some(hint => lowerHref.includes(hint))) {
          try {
            const fullUrl = new URL(href, baseUrl).toString();
            if (!links.includes(fullUrl)) {
              links.push(fullUrl);
            }
          } catch (e) {
            // Invalid URL, skip
          }
        }
      }
    });
    
    return links.slice(0, 2);
  } catch (error) {
    return [];
  }
}

// Helper: Find email for business
async function findEmailForBusiness(website) {
  if (!website) return '';
  
  // Try main page
  let email = await findEmailOnPage(website);
  if (email) return email;
  
  // Try contact pages
  const contactLinks = await findContactLinks(website);
  for (const link of contactLinks) {
    email = await findEmailOnPage(link);
    if (email) return email;
  }
  
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
    
    // Wait for token to become valid
    await new Promise(resolve => setTimeout(resolve, 2200));
  }
  
  return results.slice(0, maxResults);
}

// Main handler - works for both Vercel and local dev
async function handler(req, res) {
  // Allow CORS for local development
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }
  
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }
  
  try {
    const { query, location, apiKey, maxResults = 100, extraAreas = '', skipEmails = false } = req.body;
    
    // Validation
    if (!query || !location || !apiKey) {
      return res.status(400).json({ error: 'Missing required fields: query, location, apiKey' });
    }
    
    const maxResultsNum = Math.min(parseInt(maxResults) || 100, 500);
    const areas = [location, ...extraAreas.split(',').map(s => s.trim()).filter(Boolean)];
    
    console.log(`Starting scrape: ${query} in ${areas.join(', ')}`);
    console.log(`Max results: ${maxResultsNum}, Skip emails: ${skipEmails}`);
    
    // Collect all places
    const allPlaces = [];
    const seenIds = new Set();
    const perAreaCap = Math.max(1, Math.ceil(maxResultsNum / areas.length));
    
    for (const area of areas) {
      if (allPlaces.length >= maxResultsNum) break;
      
      console.log(`Searching in ${area}...`);
      const found = await collectPlaces(query, area, apiKey, perAreaCap);
      const newPlaces = found.filter(p => !seenIds.has(p.place_id));
      newPlaces.forEach(p => seenIds.add(p.place_id));
      allPlaces.push(...newPlaces);
      console.log(`  Found ${newPlaces.length} new places (total: ${allPlaces.length})`);
    }
    
    const finalPlaces = allPlaces.slice(0, maxResultsNum);
    console.log(`Processing ${finalPlaces.length} businesses...`);
    
    // Process each place
    const rows = [];
    const total = finalPlaces.length;
    
    for (let i = 0; i < total; i++) {
      const place = finalPlaces[i];
      const details = await getPlaceDetails(place.place_id, apiKey);
      
      await new Promise(resolve => setTimeout(resolve, 50)); // Rate limiting
      
      const website = details.website || '';
      let email = '';
      
      if (website && !skipEmails) {
        email = await findEmailForBusiness(website);
      }
      
      rows.push({
        name: details.name || place.name || '',
        category: (details.types || []).slice(0, 2).join(', '),
        phone: details.formatted_phone_number || '',
        website: website,
        email: email,
        address: details.formatted_address || '',
        rating: details.rating || '',
        reviewCount: details.user_ratings_total || '',
        mapsLink: details.url || ''
      });
      
      // Log progress
      if ((i + 1) % 10 === 0 || i === total - 1) {
        console.log(`  Processed ${i + 1}/${total}...`);
      }
    }
    
    // Create Excel file
    const wb = XLSX.utils.book_new();
    const wsData = [
      ['Name', 'Category', 'Phone', 'Website', 'Email', 'Address', 'Rating', 'Review Count', 'Google Maps Link'],
      ...rows.map(r => [r.name, r.category, r.phone, r.website, r.email, r.address, r.rating, r.reviewCount, r.mapsLink])
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, 'Leads');
    
    const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const excelBase64 = excelBuffer.toString('base64');
    
    const withEmail = rows.filter(r => r.email).length;
    
    console.log(`✅ Complete! Found ${rows.length} businesses, ${withEmail} emails`);
    
    res.status(200).json({
      success: true,
      total: rows.length,
      emailsFound: withEmail,
      data: rows,
      excel: excelBase64,
      message: `Scraped ${rows.length} businesses, found ${withEmail} emails`
    });
    
  } catch (error) {
    console.error('Scraping error:', error);
    res.status(500).json({ 
      error: error.message || 'An error occurred during scraping',
      details: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
}

// For local development server
if (require.main === module) {
  const http = require('http');
  const url = require('url');
  
  const server = http.createServer(async (req, res) => {
    // Parse request body for POST
    if (req.method === 'POST' && req.url === '/api/scrape') {
      let body = '';
      req.on('data', chunk => { body += chunk; });
      req.on('end', () => {
        try {
          req.body = JSON.parse(body);
          handler(req, res);
        } catch (e) {
          res.statusCode = 400;
          res.end(JSON.stringify({ error: 'Invalid JSON' }));
        }
      });
    } else if (req.method === 'GET' && (req.url === '/' || req.url === '')) {
      // Serve index.html for local testing
      const fs = require('fs');
      const path = require('path');
      const filePath = path.join(__dirname, '../public/index.html');
      fs.readFile(filePath, 'utf8', (err, data) => {
        if (err) {
          res.statusCode = 404;
          res.end('Not found');
        } else {
          res.setHeader('Content-Type', 'text/html');
          res.end(data);
        }
      });
    } else {
      res.statusCode = 404;
      res.end('Not found');
    }
  });
  
  const PORT = process.env.PORT || 3000;
  server.listen(PORT, () => {
    console.log(`🚀 Server running at http://localhost:${PORT}`);
    console.log(`📝 Test the API: POST http://localhost:${PORT}/api/scrape`);
  });
} else {
  // Export for Vercel
  module.exports = handler;
}