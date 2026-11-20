# Maps Lead Scraper

## Deploying to Vercel

The scraper calls Google Places from the serverless function in `api/scrape.js`.
In the Vercel project settings, add the Google API key as an environment variable
named `GOOGLE_PLACES_API_KEY` for each environment you deploy to, then redeploy.
The key is not needed in the browser or in the site's frontend configuration.

The current implementation uses the Google Places API (Legacy) Text Search and
Place Details endpoints. Enable Places API (Legacy) and billing for the Google
Cloud project associated with the key. Restrict the key to the APIs it needs,
and do not commit the key to the repository.
