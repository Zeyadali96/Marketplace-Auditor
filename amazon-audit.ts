import { chromium } from 'playwright-extra';
import stealth from 'puppeteer-extra-plugin-stealth';
import * as cheerio from 'cheerio';
import { GoogleGenAI } from '@google/genai';
import { getSimilarity, cleanAndNormalizePrice } from './amazon-helpers.ts';

// Register standard browser stealth plugins on the extra chromium instance
try {
  chromium.use(stealth());
} catch (err: any) {
  console.warn('[AMAZON SETUP] Web stealth registrations:', err.message);
}

export const amazonLocalizationMap: Record<string, { 
  locale: string
  timezoneId: string
  city: string
  zip: string
  currency: string
  countryCode?: string
  deliverTo: string[]
  lat: number
  lon: number
  userAgent?: string
  acceptLanguage?: string
}> = {
  'amazon.co.uk': { 
    locale: 'en-GB', 
    timezoneId: 'Europe/London', 
    city: 'LND', 
    zip: 'SW1A 1AA', 
    currency: 'GBP', 
    countryCode: 'GB', 
    deliverTo: ['Deliver to', 'Livre à'], 
    lat: 51.5074, 
    lon: -0.1278,
    acceptLanguage: 'en-GB,en;q=0.9,en-US;q=0.8'
  },
  'amazon.de': { 
    locale: 'de-DE', 
    timezoneId: 'Europe/Berlin', 
    city: 'BER', 
    zip: '10117', 
    currency: 'EUR', 
    countryCode: 'DE', 
    deliverTo: ['Lieferung nach', 'Liefern an', 'Deliver to'], 
    lat: 52.5200, 
    lon: 13.4050,
    acceptLanguage: 'de-DE,de;q=0.9,en;q=0.8'
  },
  'amazon.fr': { 
    locale: 'fr-FR', 
    timezoneId: 'Europe/Paris', 
    city: 'PAR', 
    zip: '75001', 
    currency: 'EUR', 
    countryCode: 'FR', 
    deliverTo: ['Livrer à', 'Livraison à', 'Deliver to'], 
    lat: 48.8566, 
    lon: 2.3522,
    acceptLanguage: 'fr-FR,fr;q=0.9,en;q=0.8'
  },
  'amazon.it': { 
    locale: 'it-IT', 
    timezoneId: 'Europe/Rome', 
    city: 'ROM', 
    zip: '00118', 
    currency: 'EUR', 
    countryCode: 'IT', 
    deliverTo: ['Invia a', 'Consegna a', 'Deliver to'], 
    lat: 41.9028, 
    lon: 12.4964,
    acceptLanguage: 'it-IT,it;q=0.9,en;q=0.8'
  },
  'amazon.es': { 
    locale: 'es-ES', 
    timezoneId: 'Europe/Madrid', 
    city: 'MAD', 
    zip: '28001', 
    currency: 'EUR', 
    countryCode: 'ES', 
    deliverTo: ['Enviar a', 'Entrega en', 'Deliver to'], 
    lat: 40.4168, 
    lon: -3.7037,
    acceptLanguage: 'es-ES,es;q=0.9,en;q=0.8'
  },
  'amazon.nl': { 
    locale: 'nl-NL', 
    timezoneId: 'Europe/Amsterdam', 
    city: 'AMS', 
    zip: '1011 AB', 
    currency: 'EUR', 
    countryCode: 'NL', 
    deliverTo: ['Bezorgen in', 'Deliver to'], 
    lat: 52.3676, 
    lon: 4.9041,
    acceptLanguage: 'nl-NL,nl;q=0.9,en;q=0.8'
  },
  'amazon.pl': { 
    locale: 'pl-PL', 
    timezoneId: 'Europe/Warsaw', 
    city: 'WAW', 
    zip: '00-001', 
    currency: 'PLN', 
    countryCode: 'PL', 
    deliverTo: ['Dostawa do', 'Wyślij do', 'Deliver to'], 
    lat: 52.2297, 
    lon: 21.0122,
    acceptLanguage: 'pl-PL,pl;q=0.9,en;q=0.8'
  },
  'amazon.se': { 
    locale: 'sv-SE', 
    timezoneId: 'Europe/Stockholm', 
    city: 'STO', 
    zip: '111 20', 
    currency: 'SEK', 
    countryCode: 'SE', 
    deliverTo: ['Skicka till', 'Leverera till', 'Deliver to'], 
    lat: 59.3293, 
    lon: 18.0686,
    acceptLanguage: 'sv-SE,sv;q=0.9,en;q=0.8'
  },
  'amazon.com.be': { 
    locale: 'nl-BE', 
    timezoneId: 'Europe/Brussels', 
    city: 'BRU', 
    zip: '1000', 
    currency: 'EUR', 
    countryCode: 'BE', 
    deliverTo: ['Bezorgen in', 'Livrer à', 'Deliver to'], 
    lat: 50.8503, 
    lon: 4.3517,
    acceptLanguage: 'nl-BE,nl;q=0.9,fr;q=0.8,en;q=0.7'
  },
};

// Helper for Amazon Google Search Grounding Fallback
export async function tryAmazonViaGemini(asin: string, domain: string): Promise<any | null> {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.log('[AMAZON GEMINI] No GEMINI_API_KEY found, skipping.');
    return null;
  }

  try {
    const genai = new GoogleGenAI({
      apiKey,
      httpOptions: {
        headers: {
          'User-Agent': 'aistudio-build',
        }
      }
    });

    console.log(`[AMAZON GEMINI] Generating content with googleSearch grounding for ASIN: ${asin} on ${domain}`);

    const prompt = `You are a product data extraction assistant with access to Google Search.
TASK: Find the LIVE product page on ${domain} for ASIN "${asin}" and extract accurate data.
STEPS:
1. Use Google Search: search for site:${domain}/dp/${asin} or search for "Amazon ${asin} on ${domain}".
2. Open the product page result from ${domain} (NOT a cached/preview snippet).
3. Extract the following fields from the LIVE product page.
FIELD RULES:
- "price": The BUY BOX price — the actual price in the main Add-to-Cart section (EUR/GBP/USD depending on domain, do not include currency symbols in the price field itself, just the raw numerical string). NOT the crossed-out list/was price. Extract the exact numerical value (e.g. "14.99").
- "shipping": The STANDARD FREE delivery message only (look for "Gratis-Lieferung / Kostenlose Lieferung / Free Delivery"). Do NOT use Prime-only expedited delivery or "fastest delivery" messages. Extract just the date part (e.g. "Mittwoch, 11. Juni" or "Wednesday, June 11").
- "buyboxOwner": The seller shown under "Verkauf durch" or "Sold by" or similar. If it is Amazon itself, return exactly "Amazon". If third-party, return their exact store name. NEVER return "N/A" if it can be found. If the Buybox owner is Amazon, you must return "Amazon".
- "title": The full product title as shown on the page.
- "description": product description details or key features, first 500 characters.
- "images": list of primary image URLs.
- "bullets": The bullet-point feature list (array of strings).
- "variations": number of product variations (integer, e.g. 0).
- "hasAPlus": boolean indicating whether A+ Content / rich description is present on the page.

CRITICAL:
- Note: This product may have multiple variations (colors, sizes) with different prices. Extract the price ONLY for the variation matching ASIN ${asin}. Do not use a cached or snippet price if it belongs to a different variation.
- If the Buybox owner is Amazon, it must return "Amazon".
- If you cannot verify the variation's price, you must return "N/A" instead of guessing.
- Do NOT hallucinate or estimate values. Every field must be grounded in what you see on the page.
- If a value genuinely cannot be found, return "N/A".

Return ONLY this JSON in a \`\`\`json block:
{
  "title": "...",
  "price": "14.99",
  "shipping": "Mittwoch, 11. Juni",
  "buyboxOwner": "Amazon",
  "description": "...",
  "images": ["url1", "url2"],
  "bullets": ["bullet 1", "bullet 2"],
  "variations": 0,
  "hasAPlus": false
}`;

    const response = await genai.models.generateContent({
      model: 'gemini-3.5-flash',
      contents: prompt,
      config: {
        tools: [{ googleSearch: {} }],
        temperature: 0.1,
      }
    });

    const rawText = response.text?.trim() || '';
    console.log('[AMAZON GEMINI] Raw response length:', rawText.length);

    const jsonText = rawText.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();

    try {
      const parsed = JSON.parse(jsonText);
      if (parsed && parsed.title) {
        parsed._source = 'gemini-google-search';
        return parsed;
      }
    } catch (parseErr) {
      const jsonMatch = rawText.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        try {
          const extracted = JSON.parse(jsonMatch[0]);
          if (extracted && extracted.title) {
            extracted._source = 'gemini-google-search';
            return extracted;
          }
        } catch (_) {}
      }
      console.log('[AMAZON GEMINI] JSON parse failed:', parseErr);
    }
  } catch (e: any) {
    console.log('[AMAZON GEMINI] Strategy failed:', e.message);
  }

  return null;
}

// Helper to perform the comparison between Master Data and Scraped Live Data
export async function performAmazonAudit(master: any, live: any, domain?: string) {
  const result: any = {
    title: { master: master.title, live: live.title, similarity: getSimilarity(master.title, live.title), match: false },
    description: { 
      master: master.description, 
      live: live.description || (live.hasAPlus ? "A+ Content Detected (Standard description missing)" : ""), 
      similarity: 0, 
      match: false, 
      isAPlus: live.hasAPlus 
    },
    bullets: [],
    price: { master: master.price, live: live.price, match: false },
    shipping: { master: master.shipping, live: live.shipping, match: false, days: live.shippingDays },
    images: { master: master.images, live: live.images, match: false },
    variations: { match: live.variations > 1 }
  };

  if (live.hasAPlus && !live.description) {
    result.description.similarity = 1.0;
    result.description.match = true;
  } else {
    result.description.similarity = getSimilarity(master.description || "", live.description || "");
    if (result.description.similarity > 0.6 || live.hasAPlus) result.description.match = true;
  }

  if (result.title.similarity > 0.8) result.title.match = true;
  
  // Bullets match and alignment
  const masterBullets = Array.isArray(master.bullets) ? master.bullets.filter(Boolean) : [];
  const liveBullets = Array.isArray(live.bullets) ? live.bullets.filter(Boolean) : [];
  const bulletsResults: any[] = [];
  const matchedLiveIndices = new Set<number>();

  // 1. Pair up each master bullet with the best matching (highest similarity) live bullet
  masterBullets.forEach((mb: string, mIdx: number) => {
    let bestSim = 0;
    let bestLiveIndex = -1;
    
    liveBullets.forEach((lb: string, idx: number) => {
      if (matchedLiveIndices.has(idx)) return;
      const sim = getSimilarity(mb, lb);
      if (sim > bestSim) {
        bestSim = sim;
        bestLiveIndex = idx;
      }
    });

    if (bestLiveIndex !== -1 && bestSim > 0.3) {
      matchedLiveIndices.add(bestLiveIndex);
      bulletsResults.push({
        master: mb,
        live: liveBullets[bestLiveIndex],
        similarity: bestSim,
        match: bestSim > 0.7
      });
    } else {
      // It's a mismatch! Let's pair it with the corresponding live bullet by index if available and unmatched,
      // or the first available unmatched live bullet, instead of returning an empty string.
      let fallbackLive = "";
      if (mIdx < liveBullets.length && !matchedLiveIndices.has(mIdx)) {
        fallbackLive = liveBullets[mIdx];
        matchedLiveIndices.add(mIdx);
      } else {
        const unmatchedIdx = liveBullets.findIndex((_, idx) => !matchedLiveIndices.has(idx));
        if (unmatchedIdx !== -1) {
          fallbackLive = liveBullets[unmatchedIdx];
          matchedLiveIndices.add(unmatchedIdx);
        }
      }

      bulletsResults.push({
        master: mb,
        live: fallbackLive,
        similarity: fallbackLive ? getSimilarity(mb, fallbackLive) : 0,
        match: false
      });
    }
  });

  // 2. For remaining unmatched live bullets, we try to place them in master rows with empty live placeholders first
  liveBullets.forEach((lb: string, idx: number) => {
    if (matchedLiveIndices.has(idx)) return;
    
    const unfilledRow = bulletsResults.find(r => r.master && !r.live);
    if (unfilledRow) {
      unfilledRow.live = lb;
      unfilledRow.similarity = getSimilarity(unfilledRow.master, lb);
      unfilledRow.match = unfilledRow.similarity > 0.7;
      matchedLiveIndices.add(idx);
    } else {
      bulletsResults.push({
        master: "",
        live: lb,
        similarity: 0,
        match: false
      });
    }
  });

  result.bullets = bulletsResults;

  // Price match (fuzzy)
  const masterPriceNum = parseFloat(String(master.price || "").replace(/[^0-9.]/g, '')) || 0;
  const livePriceNum = parseFloat(String(live.price || "").replace(/[^0-9.]/g, '')) || 0;
  if (masterPriceNum > 0 && Math.abs(masterPriceNum - livePriceNum) < 1.0) result.price.match = true;

  if (live.images && live.images.length >= (master.images?.length || 1)) result.images.match = true;

  // Score calculation
  let scoreValue = 0;
  if (result.title.match) scoreValue += 30;
  if (result.description.match) scoreValue += 30;
  const bulletMatchCount = (result.bullets || []).filter((b: any) => b.match).length;
  scoreValue += Math.min(bulletMatchCount * 8, 40);
  result.score = scoreValue;

  return result;
}

// Main Amazon Audit Pipeline Function
export async function scrapperAmazon(url: string, domain: string, locConfig: any): Promise<{ content: string; browser: any }> {
  const proxyServer = process.env.PROXY_SERVER;
  const isDeployment = !!(process.env.RAILWAY_STATIC_URL || process.env.NODE_ENV === 'production');
  
  // Determine if this is a DE or FR domain (most restrictive geo-detection)
  const isStrictGeoRegion = domain === 'amazon.de' || domain === 'amazon.fr';
  
  const launchOptions: any = {
    args: [
      '--no-sandbox', 
      '--disable-setuid-sandbox', 
      '--disable-dev-shm-usage', 
      '--disable-gpu', 
      '--single-process',
      '--disable-blink-features=AutomationControlled',
      isStrictGeoRegion ? '--disable-web-resources' : '',
    ].filter(Boolean)
  };

  if (proxyServer) {
    launchOptions.proxy = {
      server: proxyServer,
      username: process.env.PROXY_USERNAME,
      password: process.env.PROXY_PASSWORD,
    };
  }

  const browser = await chromium.launch(launchOptions).catch(err => {
    console.error("AMAZON AUDIT FAILED TO LAUNCH CHROMIUM:", err);
    throw new Error(`Browser launch failed. Error: ${err.message}`);
  });

  try {
    // Enhanced geo-IP mapping with residential IPs for better bypass
    const geoIps: Record<string, string> = {
      'amazon.co.uk': '109.169.130.1',
      'amazon.de': '109.250.12.150',      // Enhanced: Specific German residential IP
      'amazon.fr': '194.254.129.220',     // Enhanced: Specific French residential IP
      'amazon.it': '2.224.10.50',
      'amazon.es': '80.58.61.250',
      'amazon.nl': '145.97.15.200',
      'amazon.pl': '193.0.96.12',
      'amazon.se': '155.4.200.100',
      'amazon.com.be': '193.190.198.1',
      'amazon.com': '12.203.111.99'
    };
    const regionIp = geoIps[domain] || '12.203.111.99';

    // Rotate user agents for better anti-bot bypass (especially important for DE/FR)
    const userAgents = [
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36',
      'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36',
      'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36',
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0',
      'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.3 Safari/605.1.15'
    ];
    const selectedUserAgent = userAgents[Math.floor(Math.random() * userAgents.length)];

    // Enhanced anti-bot headers with proper Accept-Language
    const context = await browser.newContext({
      userAgent: selectedUserAgent,
      viewport: { width: 1920, height: 1080 },
      locale: locConfig.locale,
      timezoneId: locConfig.timezoneId,
      permissions: ['geolocation'],
      geolocation: { 
        latitude: locConfig.lat || 40.7128, 
        longitude: locConfig.lon || -74.0060, 
        accuracy: 100 
      },
      extraHTTPHeaders: {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'Accept-Language': locConfig.acceptLanguage || `${locConfig.locale},en;q=0.9`,
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Charset': 'utf-8,iso-8859-1;q=0.9,*;q=0.8',
        'Upgrade-Insecure-Requests': '1',
        'Sec-Fetch-Site': 'none',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-User': '?1',
        'Sec-Fetch-Dest': 'document',
        'Sec-Ch-Ua': '"Not A(Brand";v="99", "Google Chrome";v="136", "Chromium";v="136"',
        'Sec-Ch-Ua-Mobile': '?0',
        'Sec-Ch-Ua-Platform': '"Windows"',
        'Pragma': 'no-cache',
        'Cache-Control': 'no-cache',
        'X-Forwarded-For': regionIp,
        'Client-IP': regionIp,
        'X-Real-IP': regionIp,
        'X-Client-IP': regionIp,
        'X-Original-IP': regionIp,
      },
      ignoreHTTPSErrors: true,
      bypassCSP: true,
    });

    // Enhanced cookie management for regional settings
    const sessionId = Math.floor(Math.random() * 9000000 + 1000000);
    const ubid = Math.floor(Math.random() * 9000000 + 1000000);
    const sessionToken = 'ST-' + Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15);
    
    const cookies = [
      { name: 'lc-main', value: locConfig.locale.replace('-', '_'), domain: `.${domain}`, path: '/' },
      { name: 'i18n-prefs', value: locConfig.currency, domain: `.${domain}`, path: '/' },
      // Enhanced: zip code cookie with proper format for each region
      { name: 'sp-cdn', value: `"${locConfig.city}:${locConfig.zip}"`, domain: `.${domain}`, path: '/' },
      { name: 'session-id', value: '123-' + sessionId + '-' + Math.floor(Math.random() * 9000000 + 1000000), domain: `.${domain}`, path: '/' },
      { name: 'ubid-main', value: '123-' + ubid + '-' + Math.floor(Math.random() * 9000000 + 1000000), domain: `.${domain}`, path: '/' },
      { name: 'session-token', value: sessionToken, domain: `.${domain}`, path: '/' },
      // Additional regional cookies for better geo-spoofing
      { name: 'x-main', value: locConfig.countryCode || 'US', domain: `.${domain}`, path: '/' },
      { name: 'country-id', value: locConfig.countryCode || 'US', domain: `.${domain}`, path: '/' },
    ];
    await context.addCookies(cookies);

    const page = await context.newPage();
    
    // Enhanced route aborting for better performance
    await page.route('**/*', (route) => {
      const rUrl = route.request().url();
      const resourceType = route.request().resourceType();
      if (['font', 'media'].includes(resourceType)) {
        return route.abort();
      }
      if (resourceType === 'stylesheet' && !rUrl.includes('amazon')) {
        return route.abort();
      }
      const blockPatterns = [
        'google-analytics', 'googletagmanager', 'doubleclick',
        'facebook.net', 'fbcdn', 'adsystem', 'advertising-api',
        'amazon-adsystem', 'fls-na.amazon', 'unagi.amazon',
        'completion.amazon', 'aax-', 'mads.'
      ];
      if (blockPatterns.some(p => rUrl.includes(p))) {
        return route.abort();
      }
      return route.continue();
    });

    console.log(`[scrapperAmazon] Navigating to ${url} with enhanced DE/FR geo-bypass`);
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });

    // Enhanced page settlement with better timing detection
    console.log(`[scrapperAmazon] Waiting for page to settle...`);
    await page.waitForLoadState('networkidle').catch(() => null);
    await page.waitForTimeout(Math.random() * 1000 + 2000); // 2-3 second random wait

    // Accept cookies first
    try {
      const cookieButtons = ['#sp-cc-accept', 'input[name="accept"]', '#cookie-accept', '#accept-cookies', '.a-button-inner input[data-action="accept-cookies"]'];
      for (const selector of cookieButtons) {
        try {
          if (await page.isVisible(selector)) {
            await page.click(selector).catch(() => null);
            await page.waitForTimeout(500);
            break;
          }
        } catch (_) {}
      }
    } catch (_) {}

    // Enhanced zip code injection with multiple strategies and better error handling
    console.log(`[scrapperAmazon] Starting enhanced zip code injection for ${domain}`);
    let zipCodeInjected = false;

    for (let attempt = 1; attempt <= 4; attempt++) {
      try {
        console.log(`[scrapperAmazon] Zip code injection attempt ${attempt}/4`);

        // STRATEGY 1: Click location popover and fill zip
        try {
          const locationSelectors = ['#nav-global-location-popover-link', '#glow-ingress-block', '#nav-main-ftr-location-slot', '#nav-location-main'];
          let locBtn = null;
          for (const sel of locationSelectors) {
            const el = await page.$(sel);
            if (el && await el.isVisible().catch(() => false)) {
              locBtn = el;
              break;
            }
          }

          if (locBtn) {
            await locBtn.click({ force: true, timeout: 5000 }).catch(() => null);
            await page.waitForTimeout(1500);

            const zipInputSelectors = ['#GLUXZipUpdateInput', '#GLUXZipUpdateInput_0', 'input[aria-label*="zip" i]', 'input[aria-label*="postcode" i]'];
            let zipInput = null;
            for (const sel of zipInputSelectors) {
              const el = await page.$(sel);
              if (el) {
                const visible = await el.isVisible().catch(() => false);
                if (visible) {
                  zipInput = el;
                  break;
                }
              }
            }

            if (zipInput) {
              await zipInput.click({ clickCount: 3, timeout: 3000 }).catch(() => null);
              await page.keyboard.press('Backspace').catch(() => null);
              await zipInput.fill('');
              await zipInput.type(locConfig.zip, { delay: 100 });
              await page.waitForTimeout(800);
              await page.keyboard.press('Enter').catch(() => null);
              await page.waitForTimeout(2500);

              const confirmBtn = await page.$('#GLUXConfirmClose input, button[name="glowDoneButton"], .a-popover-footer button');
              if (confirmBtn && await confirmBtn.isVisible().catch(() => false)) {
                await confirmBtn.click({ force: true, timeout: 3000 }).catch(() => null);
                await page.waitForTimeout(1500);
              }

              zipCodeInjected = true;
              console.log(`[scrapperAmazon] Zip code injected successfully on attempt ${attempt}`);
              break;
            }
          }
        } catch (stratErr: any) {
          console.log(`[scrapperAmazon] Strategy 1 failed: ${stratErr.message}`);
        }

        if (!zipCodeInjected && attempt < 4) {
          // If injection didn't work, reload and try again
          try {
            await page.reload({ waitUntil: 'domcontentloaded', timeout: 30000 }).catch(() => null);
            await page.waitForTimeout(2000);
          } catch (_) {}
        }
      } catch (err: any) {
        console.log(`[scrapperAmazon] Zip injection attempt ${attempt} failed:`, err.message);
      }
    }

    if (zipCodeInjected) {
      console.log(`[scrapperAmazon] Reloading page to apply new location...`);
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 }).catch(() => null);
      await page.waitForLoadState('networkidle').catch(() => null);
      await page.waitForTimeout(2500);
    } else {
      console.warn(`[scrapperAmazon] Could not inject zip code after multiple attempts`);
    }

    // Final verification that we're at the right location
    try {
      const locationSlot = await page.$('#nav-global-location-slot');
      if (locationSlot) {
        const locationText = await locationSlot.textContent().catch(() => '');
        console.log(`[scrapperAmazon] Current location slot text: ${locationText}`);
      }
    } catch (_) {}

    const content = await page.content();
    return { content, browser };

  } catch (err: any) {
    await browser.close().catch(() => null);
    throw err;
  }
}

// Main Amazon Audit Pipeline Function
export async function auditAmazon(asin: string, marketplace: string, masterData: any) {
  let browser;
  try {
    const domain = marketplace || 'amazon.com';
    const url = `https://www.${domain}/dp/${asin}`;

    const locConfig = amazonLocalizationMap[domain] || { locale: 'en-US', timezoneId: 'America/New_York', city: 'NYC', zip: '10001', currency: 'USD', deliverTo: ['Deliver to'], lat: 40.7128, lon: -74.0060 };

    const scraperResult = await scrapperAmazon(url, domain, locConfig);
    browser = scraperResult.browser;
    const content = scraperResult.content;
    const $ = cheerio.load(content);

    let amazonTitle = $('#productTitle').text().trim();
    if (!amazonTitle) {
      amazonTitle = $('#title').text().trim() || $('#productTitle_feature_div').text().trim() || $('#title_feature_div').text().trim();
    }
    
    if (!amazonTitle) {
      $('script[type="a-state"]').each((_, el) => {
        try {
           const dataAState = $(el).attr('data-a-state');
           const dataKey = dataAState ? JSON.parse(dataAState).key : '';
           if (dataKey === 'turbo-checkout-product-state' || dataKey === 'turbo-checkout-page-state') {
              const jsonText = $(el).text().trim();
              const turboData = JSON.parse(jsonText);
              if (turboData.lineItemInputs?.[0]?.productTitle) {
                amazonTitle = turboData.lineItemInputs[0].productTitle;
                return false;
              } else if (turboData.turboHeaderText) {
                amazonTitle = turboData.turboHeaderText.replace(/^.*?: /, '').trim();
                return false;
              }
           }
        } catch (e) { /* ignore parse error */ }
      });
    }
    if (!amazonTitle) amazonTitle = $('meta[name="title"]').attr('content')?.split(': Amazon')[0] || "";
    if (!amazonTitle) amazonTitle = $('h1').first().text().trim() || "";

    // --- 1. Price Extraction ---
    let buyBoxContext = $('#oneTimeBuyBox').length ? $('#oneTimeBuyBox') :
                        $('#newAccordionRow').length ? $('#newAccordionRow') :
                        $('#buyNewSection').length ? $('#buyNewSection') :
                        $('#apex_offerDisplay_desktop').length ? $('#apex_offerDisplay_desktop') :
                        $('#newUnifiedOfferDisplay').length ? $('#newUnifiedOfferDisplay') :
                        $('#corePrice_feature_div').length ? $('#corePrice_feature_div') :
                        $('#desktop_buybox').length ? $('#desktop_buybox') :
                        $('#rightCol').length ? $('#rightCol') :
                        $('body');

    let amazonPrice =
      buyBoxContext.find('.priceToPay .a-offscreen').first().text().trim() ||
      $('#apex_offerDisplay_desktop .priceToPay .a-offscreen').first().text().trim() ||
      $('#apex_offerDisplay_desktop .a-price .a-offscreen').first().text().trim() ||
      $('#newUnifiedOfferDisplay .priceToPay .a-offscreen').first().text().trim() ||
      $('#newUnifiedOfferDisplay .a-price .a-offscreen').first().text().trim() ||
      $('#corePriceDisplay_desktop_feature_div .priceToPay .a-offscreen').first().text().trim() ||
      $('#corePrice_desktop .priceToPay .a-offscreen').first().text().trim() ||
      $('#corePriceDisplay_desktop_feature_div .a-price .a-offscreen').first().text().trim() ||
      $('#corePrice_desktop .a-price .a-offscreen').first().text().trim() ||
      $('#corePrice_feature_div .priceToPay .a-offscreen').first().text().trim() ||
      $('#corePrice_feature_div .a-price .a-offscreen').first().text().trim() ||
      $('#buyNew_noncbb .a-price .a-offscreen').first().text().trim() ||
      $('#buyNewSection .a-price .a-offscreen').first().text().trim() ||
      $('#desktop_buybox .a-price .a-offscreen').first().text().trim() ||
      $('#price_inside_buybox').text().trim() ||
      $('.apex-core-price-identifier .a-offscreen').first().text().trim() ||
      $('#desktop_buybox .apexPriceToPay .a-offscreen').first().text().trim() ||
      $('#desktop_buybox .priceToPay .a-offscreen').first().text().trim() ||
      $('#rightCol .priceToPay .a-offscreen').first().text().trim() ||
      $('#rightCol .a-price .a-offscreen').first().text().trim() ||
      "";
    amazonPrice = cleanAndNormalizePrice(amazonPrice);

    let listPrice = $('.basisPrice .a-offscreen').text().trim() || "";
    listPrice = cleanAndNormalizePrice(listPrice);

    // Parse FREE Delivery Shipping Time
    let rawShippingTime = "";

    const freeDeliveryKeywords = [
      'free delivery',
      'free shipping',
      'gratis-lieferung',
      'kostenlose lieferung',
      'gratis versand',
      'kostenloser versand',
      'gratisversand',
      'livraison gratuite',
      'consegna gratuita',
      'spedizione gratuita',
      'entrega gratuita',
      'envío gratis',
      'envio gratis',
      'entrega gratis',
      'gratis bezorging',
      'gratis verzending',
      'darmowa dostawa',
      'bezpłatna dostawa',
      'bezplatna dostawa',
      'gratis leverans',
      'fri frakt'
    ];

    const primeExpeditedKeywords = [
      'prime member',
      'prime-mitglieder',
      'les membres prime',
      'i clienti prime',
      'los clientes prime',
      'prime-leden',
      'prime-medlemmar',
      'expedited',
      'schnellere lieferung',
      'fastest delivery',
      'fastest',
      'snelste bezorging',
      'livraison accélérée',
      'livraison plus rapide',
      'consegna più rapida',
      'entrega más rápida',
      'snabbare leverans',
      'szybsza dostawa',
      'order within',
      'bestellen sie innerhalb',
      'commandez dans',
      'ordina entro',
      'pide dentro',
      'bestel binnen',
      'zamów w ciągu',
      'stunden',
      'minuten',
      'hours',
      'minutes',
      'heures',
      'ore',
      'ore e',
      'horas y',
      'godzin',
      'minut',
      'timmar'
    ];

    const hasFreeDeliveryKeyword = (text: string) => {
      const t = text.toLowerCase();
      return freeDeliveryKeywords.some(kw => t.includes(kw));
    };

    const hasPrimeExpeditedKeyword = (text: string) => {
      const t = text.toLowerCase();
      return primeExpeditedKeywords.some(kw => t.includes(kw));
    };

    const primaryDelivery = $('#mir-layout-DELIVERY_BLOCK-slot-PRIMARY_DELIVERY_MESSAGE_LARGE').text().replace(/\s+/g, ' ').trim() ||
                            $('#deliveryBlockMessage').text().replace(/\s+/g, ' ').trim();
    const secondaryDelivery = $('#mir-layout-DELIVERY_BLOCK-slot-SECONDARY_DELIVERY_MESSAGE_LARGE').text().replace(/\s+/g, ' ').trim();

    if (primaryDelivery && hasFreeDeliveryKeyword(primaryDelivery) && !hasPrimeExpeditedKeyword(primaryDelivery)) {
      rawShippingTime = primaryDelivery;
    } else if (secondaryDelivery && hasFreeDeliveryKeyword(secondaryDelivery) && !hasPrimeExpeditedKeyword(secondaryDelivery)) {
      rawShippingTime = secondaryDelivery;
    } else if (primaryDelivery && hasFreeDeliveryKeyword(primaryDelivery)) {
      const sentences = primaryDelivery.split(/[.·•|]|\bor\b|\boder\b|\bou\b|\bo\b|\blub\b|\beller\b/i);
      const matched = sentences.find(s => hasFreeDeliveryKeyword(s) && !hasPrimeExpeditedKeyword(s));
      if (matched) {
        rawShippingTime = matched.trim();
      } else {
        const anyFree = sentences.find(s => hasFreeDeliveryKeyword(s));
        if (anyFree) {
          rawShippingTime = anyFree.trim();
        } else {
          rawShippingTime = primaryDelivery;
        }
      }
    } else if (secondaryDelivery && hasFreeDeliveryKeyword(secondaryDelivery)) {
      const sentences = secondaryDelivery.split(/[.·•|]|\bor\b|\boder\b|\bou\b|\bo\b|\blub\b|\beller\b/i);
      const matched = sentences.find(s => hasFreeDeliveryKeyword(s) && !hasPrimeExpeditedKeyword(s));
      if (matched) {
        rawShippingTime = matched.trim();
      } else {
        const anyFree = sentences.find(s => hasFreeDeliveryKeyword(s));
        if (anyFree) {
          rawShippingTime = anyFree.trim();
        } else {
          rawShippingTime = secondaryDelivery;
        }
      }
    } else {
      const deliveryBlock = $('#mir-layout-DELIVERY_BLOCK-slot-PRIMARY_DELIVERY_MESSAGE_LARGE, #mir-layout-DELIVERY_BLOCK, #deliveryBlockMessage');
      
      let foundLine = "";
      deliveryBlock.find('*').each((_, el) => {
        const txt = $(el).text().replace(/\s+/g, ' ').trim();
        if (txt && txt.length > 5 && txt.length < 150 && hasFreeDeliveryKeyword(txt) && !hasPrimeExpeditedKeyword(txt)) {
          foundLine = txt;
          return false;
        }
      });

      if (!foundLine) {
        const wholeText = deliveryBlock.text();
        const lines = wholeText.split(/[\n\r]+/).map(l => l.replace(/\s+/g, ' ').trim()).filter(l => l.length > 5);
        const bestLine = lines.find(l => hasFreeDeliveryKeyword(l) && !hasPrimeExpeditedKeyword(l));
        if (bestLine) {
          foundLine = bestLine;
        } else {
          const anyFreeLine = lines.find(l => hasFreeDeliveryKeyword(l));
          if (anyFreeLine) {
            foundLine = anyFreeLine;
          }
        }
      }

      if (foundLine) {
        rawShippingTime = foundLine;
      } else {
        const deliveryTimeAttr = deliveryBlock.find('span[data-csa-c-delivery-time]').attr('data-csa-c-delivery-time');
        if (deliveryTimeAttr) {
          rawShippingTime = deliveryTimeAttr;
        } else {
          rawShippingTime = primaryDelivery ||
                           deliveryBlock.find('.a-text-bold').first().text().trim() ||
                           deliveryBlock.text().replace(/\s+/g, ' ').trim() ||
                           "";
        }
      }
    }

    const cleanShippingText = (text: string): string => {
      if (!text) return "";
      const segments = text.split(/\s*[\.\!\?]+\s*/).map(s => s.trim()).filter(s => s.length > 0);
      const uniqueSegments: string[] = [];
      for (const seg of segments) {
        if (!uniqueSegments.some(us => us.toLowerCase() === seg.toLowerCase() || seg.toLowerCase().includes(us.toLowerCase()) || us.toLowerCase().includes(seg.toLowerCase()))) {
          uniqueSegments.push(seg);
        }
      }
      let cleaned = uniqueSegments.join('. ').trim();

      const prefixesToRemove = [
        /free\s+(?:delivery|shipping)\s*/gi,
        /gratis-lieferung\s*(?:am|von)?\s*/gi,
        /kostenlose\s+lieferung\s*(?:am|von)?\s*/gi,
        /gratis\s+versand\s*/gi,
        /kostenloser\s+versand\s*/gi,
        /gratisversand\s*/gi,
        /livraison\s+gratuite\s*(?:le)?\s*/gi,
        /consegna\s+gratuita\s*(?:il)?\s*/gi,
        /spedizione\s+gratuita\s*(?:il)?\s*/gi,
        /entrega\s+gratuita\s*/gi,
        /envío\s+gratis\s*(?:el)?\s*/gi,
        /envio\s+gratis\s*/gi,
        /entrega\s+gratis\s*/gi,
        /gratis\s+bezorging\s*(?:op)?\s*/gi,
        /gratis\s+verzending\s*/gi,
        /darmowa\s+dostawa\s*(?:w)?\s*/gi,
        /bezpłatna\s+dostawa\s*/gi,
        /bezplatna\s+dostawa\s*/gi,
        /gratis\s+leverans\s*/gi,
        /fri\s+frakt\s*/gi,
        /delivery\s*(?:on|by)?\s*/gi,
        /consegna\s*(?:entro)?\s*/gi,
        /entrega\s*(?:el)?\s*/gi,
        /dostawa\s*(?:w)?\s*/gi,
        /u\s+wyszukiwarki\s*/gi,
        /vandaag\s+bezorgd/gi,
        /aujourd'hui/gi
      ];
      for (const rx of prefixesToRemove) {
        cleaned = cleaned.replace(rx, '');
      }

      const suffixesToRemove = [
        /\s*przy\s+pierwszym\s+zamówieniu.*/gi,
        /\s*przy\s+pierwszym\s+zakupie.*/gi,
        /\s*avec\s+(?:votre\s+)?première\s+commande.*/gi,
        /\s*lors\s+de\s+(?:votre\s+)?première\s+commande.*/gi,
        /\s*bei\s+(?:der\s+)?ersten\s+bestellung.*/gi,
        /\s*per\s+il\s+primo\s+ordine.*/gi,
        /\s*en\s+tu\s+primer\s+pedido.*/gi,
        /\s*voor\s+(?:uw\s+)?eerste\s+bestelling.*/gi,
        /\s*on\s+your\s+first\s+order.*/gi,
        /\s*\(.*subscription.*\)/gi,
        /\s*\(.*prime member.*\)/gi,
        /\s*or\s+faster.*/gi,
        /\.?\s*Détails.*/gi,
        /\.?\s*Details.*/gi,
        /\.?\s*Informations?.*/gi,
        /\s+ze\s+szczegółami.*/gi,
        /\s*więcej\s+informacji.*/gi,
        /mehr\s+details.*/gi
      ];
      for (const rx of suffixesToRemove) {
        cleaned = cleaned.replace(rx, '');
      }

      cleaned = cleaned
        .replace(/^(?:le|la|el|los|on|op|am|op|przy|w|v|at|by|from|auf|de|di|d'|el)\s+/gi, '')
        .trim();

      cleaned = cleaned
        .replace(/^[\s,.;:or|]+|[\s,.;:or|]+$/gi, '')
        .replace(/\s+/g, ' ')
        .trim();

      return cleaned;
    };

    const cleanedShipping = cleanShippingText(rawShippingTime);
    if (cleanedShipping) {
      rawShippingTime = cleanedShipping;
    } else {
      rawShippingTime = rawShippingTime
        .replace(/^[\s,.;:or|]+|[\s,.;:or|]+$/gi, '')
        .replace(/\s+/g, ' ')
        .trim();
    }

    let amazonDesc = $('#productDescription').text().trim();
    if (!amazonDesc) amazonDesc = $('#feature-bullets').text().trim();
    
    let hasAPlus = !!($('#aplus').length || $('#aplus_feature_div').length || $('div[id*="aplus"]').length);

    // Enhanced Buy Box Owner Extraction with language-specific patterns for DE/FR
    let amazonBuyboxOwner = "";

    // Regional patterns for detecting Amazon seller vs third-party
    const amazonSellerPatterns = {
      'en': /sold by amazon|dispatched from and sold by amazon|fulfilled by amazon|amazon.com services/i,
      'de': /verkauf durch amazon|versandt und verkauft durch amazon|amazon eu sarl|amazon media eu|verfügbar bei und verkauft durch amazon/i,
      'fr': /vendu par amazon|expédié par amazon|amazon.fr|amazonfr/i,
      'it': /venduto da amazon|spedito e venduto da amazon|amazon media eu srl/i,
      'es': /vendido por amazon|enviado y vendido por amazon|amazon media eu/i,
      'nl': /verkocht door amazon|verzonden en verkocht door amazon/i,
      'pl': /sprzedawane przez amazon|wysłane i sprzedawane przez amazon/i,
      'sv': /säljs av amazon|skickas från och säljs av amazon/i,
    };

    // Determine language code from domain
    const langMap: Record<string, string> = {
      'amazon.de': 'de',
      'amazon.fr': 'fr',
      'amazon.it': 'it',
      'amazon.es': 'es',
      'amazon.nl': 'nl',
      'amazon.pl': 'pl',
      'amazon.se': 'sv',
      'amazon.co.uk': 'en',
      'amazon.com': 'en',
    };
    const lang = langMap[domain] || 'en';

    // Get the Buy Box context
    let mainMerchantEl: any = null;
    const merchantEls = $('#merchant-info');
    merchantEls.each((_, el) => {
      if ($(el).closest('#subscribeAndSaveAccordionRow, [class*="sns"], [id*="sns"]').length === 0) {
        mainMerchantEl = $(el);
        return false;
      }
    });
    if (!mainMerchantEl && merchantEls.length > 0) {
      mainMerchantEl = merchantEls.first();
    }
    const mainMerchantText = mainMerchantEl ? mainMerchantEl.text().replace(/\s+/g, ' ').trim() : "";

    // Enhanced language-specific selectors for regional versions
    let soldByTabular = "";
    const soldBySelectors = [
      'div[tabular-attribute-name*="Sold by" i]',
      'div[tabular-attribute-name*="Verkauf durch" i]',
      'div[tabular-attribute-name*="Vendido por" i]',
      'div[tabular-attribute-name*="Vendu par" i]',
      'div[tabular-attribute-name*="Venduto da" i]',
      'div[tabular-attribute-name*="Sprzedawca" i]',
      'div[tabular-attribute-name*="Säljs av" i]',
      'div[tabular-attribute-name*="Verkocht door" i]'
    ];

    for (const sel of soldBySelectors) {
      const els = buyBoxContext.find(`${sel} .tabular-buybox-text`);
      els.each((_, el) => {
        if ($(el).closest('#subscribeAndSaveAccordionRow, [class*="sns"], [id*="sns"]').length === 0) {
          soldByTabular = $(el).text().trim();
          return false;
        }
      });
      if (soldByTabular) break;
    }

    // Check if it's Amazon using language-specific patterns
    const amazonPattern = amazonSellerPatterns[lang as keyof typeof amazonSellerPatterns] || amazonSellerPatterns['en'];
    const isAmazonSeller = amazonPattern.test(mainMerchantText) || amazonPattern.test(soldByTabular);

    if (isAmazonSeller) {
      amazonBuyboxOwner = "Amazon";
    }

    if (!amazonBuyboxOwner) {
      let sellerTrigger = "";
      buyBoxContext.find('#sellerProfileTriggerId').each((_, el) => {
        if ($(el).closest('#subscribeAndSaveAccordionRow, [class*="sns"], [id*="sns"]').length === 0) {
          sellerTrigger = $(el).text().trim();
          return false;
        }
      });

      let merchantLink = "";
      buyBoxContext.find('#merchant-info a').each((_, el) => {
        if ($(el).closest('#subscribeAndSaveAccordionRow, [class*="sns"], [id*="sns"]').length === 0) {
          merchantLink = $(el).text().trim();
          return false;
        }
      });

      amazonBuyboxOwner = 
        soldByTabular ||
        buyBoxContext.find('div[offer-display-feature-name="desktop-merchant-info"] .offer-display-feature-text-message').first().text().trim() ||
        sellerTrigger ||
        merchantLink;
    }

    // Final fallback if still not found
    if (!amazonBuyboxOwner) {
      amazonBuyboxOwner = 
        $('div[tabular-attribute-name="Sold by"] .tabular-buybox-text').first().text().trim() ||
        $('div[tabular-attribute-name="Verkauf durch"] .tabular-buybox-text').first().text().trim() ||
        $('div[tabular-attribute-name="Vendu par"] .tabular-buybox-text').first().text().trim() ||
        $('div[tabular-attribute-name="Venduto da"] .tabular-buybox-text').first().text().trim() ||
        $('div[offer-display-feature-name="desktop-merchant-info"] .offer-display-feature-text-message').first().text().trim() ||
        $('#sellerProfileTriggerId').first().text().trim() ||
        $('#merchant-info a').first().text().trim();
    }

    // Clean up the extracted seller name
    if (!amazonBuyboxOwner && mainMerchantText.length > 0) {
      const sellerLink = mainMerchantEl ? (mainMerchantEl as any).find('a').first().text().trim() || $('#merchant-info a').first().text().trim() : "";
      if (sellerLink && !amazonPattern.test(sellerLink)) {
        amazonBuyboxOwner = sellerLink;
      } else {
        amazonBuyboxOwner = mainMerchantText
          .replace(/Dispatched from and sold by\s*/i, '')
          .replace(/Dispatched from Amazon\s*\.?\s*/i, '')
          .replace(/Sold by\s*/i, '')
          .replace(/Verkauf durch\s*/i, '')
          .replace(/Vendu par\s*/i, '')
          .replace(/Venduto da\s*/i, '')
          .replace(/Fulfilled by Amazon\s*\.?\s*/i, '')
          .replace(/\|.*$/s, '')
          .trim();
      }
    }

    // Final cleanup
    amazonBuyboxOwner = amazonBuyboxOwner
      .replace(/Sold by\s*:?\s*/gi, '')
      .replace(/Verkauf durch\s*:?\s*/gi, '')
      .replace(/Vendu par\s*:?\s*/gi, '')
      .replace(/Venduto da\s*:?\s*/gi, '')
      .replace(/Sprzedawca\s*:?\s*/gi, '')
      .replace(/Säljs av\s*:?\s*/gi, '')
      .replace(/Verkocht door\s*:?\s*/gi, '')
      .replace(/Dispatched from and sold by\s*/gi, '')
      .replace(/Dispatched from Amazon\.?\s*/gi, '')
      .replace(/Fulfilled by Amazon\.?\s*/gi, '')
      .replace(/\|.*$/s, '')
      .trim();

    // --- 7. Hardened Image Extraction ---
    const imageMap = new Map<string, string>();

    const getNormalizedInfo = (imageUrl: string | undefined) => {
      if (!imageUrl || typeof imageUrl !== "string" || imageUrl.length < 15) return null;
      let cleaned = imageUrl.split("?")[0].trim();
      if (cleaned.startsWith("//")) cleaned = "https:" + cleaned;
      cleaned = cleaned.replace(/^http:/, "https:");
      cleaned = cleaned.replace(/\._[a-zA-Z0-9,_-]+_\.?/g, ".");
      cleaned = cleaned.replace(/\.V[0-9]+_\.?/g, ".");
      cleaned = cleaned.replace(/\.(V|SS|SX|SY|AC|SR|SL|UL|CLa|SR|SS|SX|SY|UL|CLa)[0-9,s]+_\./g, ".");
      cleaned = cleaned.replace(/\.(V|SS|SX|SY|AC|SR|SL|UL|CLa|SR|SS|SX|SY|UL|CLa)[0-9,s]+\./g, ".");

      const idMatch = cleaned.match(/\/images\/(I|W|S|G)\/([^\.\/]+)/) || cleaned.match(/\/images\/([^\.\/]+)\./);
      if (!idMatch) return null;

      const baseId = ((idMatch.length > 2 ? idMatch[2] : idMatch[1]) as string).replace(/\.+$/, "");
      if (baseId.includes("SWCH") || cleaned.includes("_SW") || baseId.startsWith("ss_") || cleaned.includes("play-button")) return null;

      return { baseId, url: cleaned };
    };

    const mainHeroUrl = $("#landingImage").attr("data-old-hires") || $("#landingImage").attr("src");
    const heroInfo = getNormalizedInfo(mainHeroUrl);
    
    console.log("=== AMAZON IMAGE EXTRACTION DEBUG ===");
    console.log("Hero URL:", mainHeroUrl);
    console.log("Hero Info:", heroInfo);
    
    if (heroInfo) {
      imageMap.set(heroInfo.baseId, heroInfo.url);
      console.log("Hero added to map with baseId:", heroInfo.baseId);
    }

    let thumbCount = 0;
    $("#altImages li.imageThumbnail:not(.videoThumbnail) img, .imageThumbnail img, .altImages img").not("#landingImage").each((_, el) => {
      thumbCount++;
      const src = $(el).attr("data-old-hires") || $(el).attr("src");
      if (!src || src.includes("pixel.gif") || src.includes("play-button-overlay") || src.includes("transparent-pixel")) {
        return;
      }

      const rawSrcInfo = getNormalizedInfo($(el).attr("src"));
      if (heroInfo && rawSrcInfo && rawSrcInfo.baseId === heroInfo.baseId) {
        return;
      }

      const info = getNormalizedInfo(src);
      if (!info) return;
      if (imageMap.has(info.baseId)) return;
      
      imageMap.set(info.baseId, info.url);
    });

    let uniqueImages = Array.from(imageMap.values());

    const bulletSet = new Set<string>();
    const bulletSelectors = [
      '#feature-bullets ul li span.a-list-item',
      '#featurebullets_feature_div ul li span.a-list-item'
    ];

    $(bulletSelectors.join(', ')).each((_, el) => {
      const $el = $(el);
      const junkContainers = [
        '#customerReviews',
        '#reviews-medley-footer',
        '#cm-cr-dp-review-list',
        '.customer_review',
        '#fbt_x_cl_div',
        '#legal-disclaimer',
        '#ad-feedback-form-desktop-feature-bullets_secondary_view_div',
        '#reviews-image-gallery-container',
        '#social-proofing-faceout-feature-div',
        '#dp-ads-center-promo-pc_desktop_view_div',
        '.cr-widget-FocalReviews',
        '#detailBullets_feature_div',
        '#prodDetails',
        '#productDetails_feature_div',
        '#technicalSpecifications_feature_div',
        '#detail-bullets',
        '.detail-bullets',
        '.detail-bullet-list',
        '#productDetails_db_sections',
        '#averageCustomerReviews',
        '#averageCustomerReviews_feature_div',
        '#reviewsMedley',
        '#twister',
        '#twister-plus-inline-twister',
        '#inline-twister-row-all-options',
        '#variation_color_name',
        '#variation_size_name',
        '#variation_style_name',
        '#inline-twister-dim-values-container',
        '#tp-inline-twister-dim-values-container',
        '.twister-image-select',
        '#quickPromoBucketContent',
        '#similarities_feature_div',
        '#HLCXComparisonWidget_feature_div'
      ];
      if ($el.closest(junkContainers.join(', ')).length > 0) {
        return;
      }

      const $clone = $el.clone();
      $clone.find('script, style, .a-declarative, .a-popover-preload').remove();
      
      let text = '';
      const $span = $clone.find('span.a-list-item');
      if ($span.length > 0) {
        text = $span.first().text().trim();
      } else {
        text = $clone.text().trim();
      }
      text = text.replace(/^[•\-\*\s]+/, '').trim();

      const isJunk = (t: string) => {
        const lower = t.toLowerCase();
        const isBSR = 
          lower.includes('top 100') ||
          lower.includes('en más vendidos') ||
          lower.includes('en mas vendidos') ||
          lower.includes('best seller') ||
          lower.includes('bestseller') ||
          /nº\s*\d+/i.test(lower) ||
          /#\s*\d+/i.test(lower) ||
          /puesto\s*nº\s*\d+/i.test(lower) ||
          /ranking/i.test(lower) ||
          /clasificación/i.test(lower) ||
          /bestsellers/i.test(lower);

        return (
          lower.includes('window.ue') ||
          lower.includes('if(window.ue)') ||
          lower.includes('out of 5 stars') ||
          lower.includes('verified purchase') ||
          lower.includes('helpful report') ||
          lower.includes('reviewed in') ||
          lower.includes('not for sale to persons under') ||
          lower.includes('16 years of age') ||
          lower.includes('make sure this fits') ||
          lower.includes('geben sie ihr modell ein') ||
          lower.includes('sprawdź, czy pasuje') ||
          lower.includes('read more') ||
          lower.includes('customer reviews') ||
          lower.includes('by placing an order') ||
          lower.includes('declare that you are') ||
          lower.includes('used responsibly') ||
          t.length < 5 ||
          t.length > 1000 ||
          /^(\d+)\s+out of\s+5\s+stars/i.test(t) ||
          /Reviewed in the .* on \d+/.test(t) ||
          /^\d+ ratings?$/.test(t) ||
          isBSR
        );
      };

      if (text && !isJunk(text)) {
        bulletSet.add(text);
      }
    });

    let amazonBullets = Array.from(bulletSet);

    const variationsSet = new Set<string>();
    const optionSelectors = [
      '#twister li[data-asin]',
      '#inline-twister-row-all-options li[data-asin]',
      '#twister .a-button-toggle',
      '#twister .swatchAvailable',
      '#twister .swatchSelect',
      '#variation_color_name li',
      '#variation_size_name li',
      '#variation_style_name li',
      '#variation_pattern_name li',
      '#variation_flavor_name li',
      '#native_dropdown_selected_size_name option:not([value="-1"])',
      '#native_dropdown_selected_color_name option:not([value="-1"])',
      '#native_dropdown_selected_style_name option:not([value="-1"])',
      '.visualSelection .visual-selection-button',
      '.tp-inline-twister-dim-values-container li'
    ];

    $(optionSelectors.join(', ')).each((_, el) => {
      const $el = $(el);
      const text = $el.text().trim();
      const sAsin = $el.attr('data-asin');
      if (sAsin) {
        variationsSet.add(sAsin);
      } else if (text && text.length > 0 && !text.toLowerCase().includes('select')) {
        variationsSet.add(text);
      }
    });

    $('[data-totalvariationcount]').each((_, el) => {
      const countAttr = $(el).attr('data-totalvariationcount');
      if (countAttr) {
        const count = parseInt(countAttr);
        if (!isNaN(count) && count > 1) {
          for (let i = 0; i < Math.min(count, 10); i++) {
            variationsSet.add(`TwisterPlus-Var-${i}`);
          }
        }
      }
    });

    if (variationsSet.size <= 1) {
      const rows = $('#twister .a-row.variation-row, #twister .twister-selection-column, [id^="inline-twister-row-"]');
      rows.each((_, row) => {
        const $row = $(row);
        const options = $row.find('li, .a-button, option').filter((_, opt) => {
          const $opt = $(opt);
          const t = $opt.text().toLowerCase();
          return t.length > 0 && !t.includes('select');
        });
        
        if (options.length > 1) {
          options.each((_, opt) => {
            const val = $(opt).attr('data-asin') || $(opt).text().trim();
            if (val) variationsSet.add(val);
          });
        }
      });
    }

    let variationsCount = variationsSet.size;

    let shippingDays = "N/A";
    try {
      if (rawShippingTime) {
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const raw = rawShippingTime.toLowerCase();

        if (raw.includes('today') || raw.includes('vandaag') || raw.includes('aujourd') ||
            raw.includes('oggi') || raw.includes('heute')) {
          shippingDays = "0";
        } else if (raw.includes('tomorrow') || raw.includes('morgen') || raw.includes('demain') ||
                 raw.includes('domani') || raw.includes('jutro') || raw.includes('mañana') ||
                 raw.includes('imorgon')) {
          shippingDays = "1";
        } else if (raw.includes('overmorgen') || raw.includes('après-demain') || raw.includes('dopodomani')) {
          shippingDays = "2";
        } else {
          const rangeMatch = raw.match(/(\d+)\s*[-–]\s*(\d+)\s*(?:day|jour|tag|giorn|dí|di|dag|dn|dagar)/i);
          const singleMatch = raw.match(/(\d+)\s*(?:day|jour|tag|giorn|dí|di|dag|dn|dagar)/i);
          if (rangeMatch) {
            shippingDays = rangeMatch[2];
          } else if (singleMatch) {
            shippingDays = singleMatch[1];
          } else {
            let targetDate: Date | null = null;
            const monthNamesForRegex = [
              'jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec',
              'january', 'february', 'march', 'april', 'june', 'july', 'august', 'september', 'october', 'november', 'december',
              'januar', 'februar', 'märz', 'juni', 'juli', 'oktober', 'okt', 'dezember', 'dez',
              'janvier', 'janv', 'février', 'févr', 'avril', 'avr', 'juin', 'juillet', 'juil', 'août', 'aoû', 'septembre', 'octobre', 'novembre', 'décembre', 'déc',
              'enero', 'ene', 'febrero', 'marzo', 'abril', 'abr', 'mayo', 'may', 'junio', 'julio', 'agosto', 'ago', 'septiembre', 'octubre', 'noviembre', 'diciembre', 'dic',
              'gennaio', 'gen', 'febbraio', 'aprile', 'maggio', 'mag', 'giugno', 'giu', 'luglio', 'lug', 'settembre', 'set', 'ottobre', 'ott', 'dicembre',
              'styczeń', 'stycznia', 'sty', 'luty', 'lutego', 'lut', 'marzec', 'marca', 'kwiecień', 'kwietnia', 'kwi', 'maja', 'maj', 'mai', 'czerwiec', 'czerwca', 'cze', 'lipiec', 'lipca', 'sierpień', 'sierpnia', 'sie', 'wrzesień', 'września', 'wrz', 'październik', 'października', 'paź', 'listopad', 'listopada', 'lis', 'grudzień', 'grudnia', 'gru',
              'januari', 'augusti', 'maart', 'maa', 'mei', 'augustus'
            ];

            const getMonthIndex = (monthStr: string): number => {
              const m = monthStr.toLowerCase();
              const map: { [key: string]: number } = {
                'jan': 0, 'feb': 1, 'mar': 2, 'apr': 3, 'may': 4, 'jun': 5, 'jul': 6, 'aug': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dec': 11,
                'january': 0, 'february': 1, 'march': 2, 'april': 3, 'june': 5, 'july': 6, 'august': 7, 'september': 8, 'october': 9, 'november': 10, 'december': 11,
                'januar': 0, 'februar': 1, 'märz': 2, 'mär': 2, 'juni': 5, 'juli': 6, 'oktober': 9, 'okt': 9, 'dezember': 11, 'dez': 11,
                'janvier': 0, 'janv': 0, 'février': 1, 'févr': 1, 'mars': 2, 'avril': 3, 'avr': 3, 'juin': 5, 'juillet': 6, 'juil': 6, 'août': 7, 'aoû': 7, 'septembre': 8, 'octobre': 9, 'novembre': 10, 'décembre': 11, 'déc': 11,
                'enero': 0, 'ene': 0, 'febrero': 1, 'marzo': 2, 'abril': 3, 'abr': 3, 'mayo': 4, 'junio': 5, 'julio': 6, 'agosto': 7, 'ago': 7, 'septiembre': 8, 'octubre': 9, 'noviembre': 10, 'diciembre': 11, 'dic': 11,
                'gennaio': 0, 'gen': 0, 'febbraio': 1, 'aprile': 3, 'maggio': 4, 'mag': 4, 'giugno': 5, 'giu': 5, 'luglio': 6, 'lug': 6, 'settembre': 8, 'set': 8, 'ottobre': 9, 'ott': 9, 'dicembre': 11,
                'styczeń': 0, 'stycznia': 0, 'sty': 0, 'luty': 1, 'lutego': 1, 'lut': 1, 'marzec': 2, 'marca': 2, 'kwiecień': 3, 'kwietnia': 3, 'kwi': 3, 'maja': 4, 'maj': 4, 'mai': 4, 'czerwiec': 5, 'czerwca': 5, 'cze': 5, 'lipiec': 6, 'lipca': 6, 'sierpień': 7, 'sierpnia': 7, 'sie': 7, 'wrzesień': 8, 'września': 8, 'wrz': 8, 'październik': 9, 'października': 9, 'paź': 9, 'listopad': 10, 'listopada': 10, 'lis': 10, 'grudzień': 11, 'grudnia': 11, 'gru': 11,
                'januari': 0, 'augusti': 7,
                'maart': 2, 'maa': 2, 'mei': 4, 'augustus': 7
              };

              if (map[m] !== undefined) return map[m];
              for (const [key, val] of Object.entries(map)) {
                if (m.includes(key) || key.includes(m)) {
                  return val;
                }
              }
              return -1;
            };

            const monthRegexPattern = monthNamesForRegex.join('|');
            let day = -1;
            let monthStr = "";

            const rangeDateMatch = raw.match(new RegExp(`(?:\\d{1,2})\\s*[-–]\\s*(\\d{1,2})(?:\\.?\\s*(?:de|di|d')?\\s*)(${monthRegexPattern})`, 'i'));
            if (rangeDateMatch) {
              day = parseInt(rangeDateMatch[1]);
              monthStr = rangeDateMatch[2];
            } else {
              const match1 = raw.match(new RegExp(`(\\d{1,2})(?:\\.?\\s*(?:de|di|d')?\\s*)(${monthRegexPattern})`, 'i'));
              if (match1) {
                day = parseInt(match1[1]);
                monthStr = match1[2];
              } else {
                const match2 = raw.match(new RegExp(`(${monthRegexPattern})(?:\\s*(?:de|di)?\\s*)(\\d{1,2})`, 'i'));
                if (match2) {
                  day = parseInt(match2[2]);
                  monthStr = match2[1];
                }
              }
            }

            if (day !== -1 && monthStr) {
              const monthIndex = getMonthIndex(monthStr);
              if (monthIndex !== -1) {
                targetDate = new Date(today.getFullYear(), monthIndex, day);
                if (targetDate.getTime() < today.getTime() - (24 * 60 * 60 * 1000 * 2)) {
                  targetDate.setFullYear(today.getFullYear() + 1);
                }
              }
            }

            if (targetDate) {
              targetDate.setHours(0, 0, 0, 0);
              const diffTime = targetDate.getTime() - today.getTime();
              const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
              if (diffDays >= 0) shippingDays = diffDays.toString();
            }
          }
        }
      }
    } catch (e: any) {
      console.warn("Shipping days calculation failed:", e.message);
    }

    // Fallback: Check if scraping is incomplete
    const isScrapingIncomplete = 
      !amazonTitle || 
      amazonTitle.toLowerCase().includes('robot') || 
      amazonTitle.toLowerCase().includes('captcha') || 
      amazonTitle.toLowerCase().includes('unusual traffic') ||
      amazonTitle.toLowerCase().includes('automated access') ||
      amazonTitle.length < 3 ||
      !amazonPrice || 
      amazonPrice === 'N/A' || 
      shippingDays === 'N/A' ||
      !rawShippingTime || 
      rawShippingTime === 'N/A' ||
      !amazonBuyboxOwner || 
      amazonBuyboxOwner === 'N/A' || 
      amazonBuyboxOwner.trim() === '' ||
      amazonBullets.length === 0;

    if (isScrapingIncomplete) {
      console.log('[AMAZON] Scraped data is incomplete or empty. Trying Gemini Google Search Grounding fallback...');
      const geminiData = await tryAmazonViaGemini(asin, domain);
      if (geminiData && geminiData.title) {
        console.log('[AMAZON] Gemini Fallback Succeeded.');
        if (!amazonTitle || amazonTitle.toLowerCase().includes('robot') || amazonTitle.toLowerCase().includes('captcha') || amazonTitle.toLowerCase().includes('unusual traffic') || amazonTitle.length < 3) {
          amazonTitle = geminiData.title;
        }
        if (geminiData.price && (!amazonPrice || amazonPrice === 'N/A')) {
          amazonPrice = cleanAndNormalizePrice(geminiData.price);
        }
        if (geminiData.description && (!amazonDesc || amazonDesc.length < 5)) {
          amazonDesc = geminiData.description;
        }
        if (geminiData.buyboxOwner && (!amazonBuyboxOwner || amazonBuyboxOwner === 'N/A')) {
          amazonBuyboxOwner = geminiData.buyboxOwner;
        }
        if (geminiData.bullets && geminiData.bullets.length > 0 && amazonBullets.length === 0) {
          amazonBullets = geminiData.bullets;
        }
        if (geminiData.images && geminiData.images.length > 0 && uniqueImages.length === 0) {
          uniqueImages = geminiData.images;
        }
        if (geminiData.hasAPlus !== undefined && !hasAPlus) {
          hasAPlus = geminiData.hasAPlus;
        }
        if (geminiData.variations !== undefined && variationsCount <= 1) {
          variationsCount = typeof geminiData.variations === 'number' ? geminiData.variations : parseInt(geminiData.variations) || 0;
        }
        if (geminiData.shipping && (!rawShippingTime || rawShippingTime === 'N/A')) {
          rawShippingTime = geminiData.shipping;
          try {
            const raw = rawShippingTime.toLowerCase();
            if (raw.includes('today') || raw.includes('vandaag') || raw.includes('aujourd') ||
                raw.includes('oggi') || raw.includes('heute')) {
              shippingDays = "0";
            } else if (raw.includes('tomorrow') || raw.includes('morgen') || raw.includes('demain') ||
                       raw.includes('domani') || raw.includes('jutro') || raw.includes('mañana') ||
                       raw.includes('imorgon')) {
              shippingDays = "1";
            } else if (raw.includes('overmorgen') || raw.includes('après-demain') || raw.includes('dopodomani')) {
              shippingDays = "2";
            } else {
              const rangeMatch = raw.match(/(\d+)\s*[-–]\s*(\d+)/);
              const singleMatch = raw.match(/(\d+)/);
              if (rangeMatch) {
                shippingDays = rangeMatch[2];
              } else if (singleMatch) {
                shippingDays = singleMatch[1];
              }
            }
          } catch (shErr) {
            console.warn("Fallback shipping days parse failed:", shErr);
          }
        }
      }
    }

    const liveData = {
      title: amazonTitle,
      description: amazonDesc,
      bullets: amazonBullets,
      price: amazonPrice,
      listPrice: listPrice,
      currency: locConfig.currency,
      shipping: shippingDays !== "N/A" ? `${shippingDays} days` : rawShippingTime,
      shippingDays: shippingDays,
      rawShipping: rawShippingTime,
      variations: variationsCount,
      hasAPlus: hasAPlus,
      buyboxOwner: amazonBuyboxOwner,
      images: Array.from(new Set(uniqueImages))
    };

    const auditResult = await performAmazonAudit(masterData, liveData, domain);
    return { liveData, auditResult };

  } finally {
    if (browser) await browser.close();
  }
}
