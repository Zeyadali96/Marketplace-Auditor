import { chromium } from 'playwright';
import * as cheerio from 'cheerio';
import { GoogleGenAI } from '@google/genai';
import { getSimilarity, cleanAndNormalizePrice } from './helpers.ts';

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

    const prompt = `Perform a google search for "site:${domain}/dp/${asin}" or search for "Amazon ${asin} on ${domain}".
Locate the official product page on ${domain}.
Extract and return a single, exact JSON object with the following schema:
{
  "title": "exact full product title on Amazon",
  "price": "correct numerical price string e.g. 14.99 of the primary default offer",
  "shipping": "correct standard free delivery message / shipping time, e.g. 'Standard-Lieferung am Freitag, 22. Mai' or 'Wednesday, May 27' (do NOT use Prime expedited/fastest, just the standard free delivery message)",
  "description": "product description details or key features, first 500 characters",
  "images": ["image url 1", "image url 2"],
  "bullets": ["feature point 1", "feature point 2"],
  "buyboxOwner": "The seller name. If sold by Amazon, return 'Amazon'. If sold by a 3rd party, return the 3rd party seller name.",
  "variations": 0,
  "hasAPlus": false
}
Make sure all details (pricing, title, shipping, buyboxOwner) are fully grounded in search results. Ensure the return contains ONLY the raw JSON object. No conversational helper text, no markdown other than \`\`\`json.`;

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
export async function auditAmazon(asin: string, marketplace: string, masterData: any) {
  let browser;
  try {
    const domain = marketplace || 'amazon.com';
    const url = `https://www.${domain}/dp/${asin}`;
    
    const proxyServer = process.env.PROXY_SERVER;
    
    const launchOptions: any = {
      args: [
        '--no-sandbox', 
        '--disable-setuid-sandbox', 
        '--disable-dev-shm-usage', 
        '--disable-gpu', 
        '--single-process',
        '--disable-blink-features=AutomationControlled'
      ]
    };

    if (proxyServer) {
      launchOptions.proxy = {
        server: proxyServer,
        username: process.env.PROXY_USERNAME,
        password: process.env.PROXY_PASSWORD,
      };
    }

    browser = await chromium.launch(launchOptions).catch(err => {
      console.error("AMAZON AUDIT FAILED TO LAUNCH CHROMIUM:", err);
      throw new Error(`Browser launch failed. Error: ${err.message}`);
    });

    const amazonLocalizationMap: Record<string, { locale: string; timezoneId: string; city: string; zip: string; currency: string; countryCode?: string; deliverTo: string[] }> = {
      'amazon.co.uk': { locale: 'en-GB', timezoneId: 'Europe/London', city: 'LND', zip: 'SW1A 1AA', currency: 'GBP', countryCode: 'GB', deliverTo: ['Deliver to', 'Livre à'] },
      'amazon.de': { locale: 'de-DE', timezoneId: 'Europe/Berlin', city: 'BER', zip: '10117', currency: 'EUR', countryCode: 'DE', deliverTo: ['Lieferung nach', 'Liefern an', 'Deliver to'] },
      'amazon.fr': { locale: 'fr-FR', timezoneId: 'Europe/Paris', city: 'PAR', zip: '75001', currency: 'EUR', countryCode: 'FR', deliverTo: ['Livrer à', 'Livraison à', 'Deliver to'] },
      'amazon.it': { locale: 'it-IT', timezoneId: 'Europe/Rome', city: 'ROM', zip: '00118', currency: 'EUR', countryCode: 'IT', deliverTo: ['Invia a', 'Consegna a', 'Deliver to'] },
      'amazon.es': { locale: 'es-ES', timezoneId: 'Europe/Madrid', city: 'MAD', zip: '28001', currency: 'EUR', countryCode: 'ES', deliverTo: ['Enviar a', 'Entrega en', 'Deliver to'] },
      'amazon.nl': { locale: 'nl-NL', timezoneId: 'Europe/Amsterdam', city: 'AMS', zip: '1011 AB', currency: 'EUR', countryCode: 'NL', deliverTo: ['Bezorgen in', 'Deliver to'] },
      'amazon.pl': { locale: 'pl-PL', timezoneId: 'Europe/Warsaw', city: 'WAW', zip: '00-001', currency: 'PLN', countryCode: 'PL', deliverTo: ['Dostawa do', 'Wyślij do', 'Deliver to'] },
      'amazon.se': { locale: 'sv-SE', timezoneId: 'Europe/Stockholm', city: 'STO', zip: '111 20', currency: 'SEK', countryCode: 'SE', deliverTo: ['Skicka till', 'Leverera till', 'Deliver to'] },
      'amazon.com.be': { locale: 'nl-BE', timezoneId: 'Europe/Brussels', city: 'BRU', zip: '1000', currency: 'EUR', countryCode: 'BE', deliverTo: ['Bezorgen in', 'Livrer à', 'Deliver to'] },
    };

    const locConfig = amazonLocalizationMap[domain] || { locale: 'en-US', timezoneId: 'America/New_York', city: 'NYC', zip: '10001', currency: 'USD', deliverTo: ['Deliver to'] };

    const context = await browser.newContext({
      userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36',
      viewport: { width: 1920, height: 1080 },
      locale: locConfig.locale,
      timezoneId: locConfig.timezoneId,
      extraHTTPHeaders: {
        'Accept-Language': `${locConfig.locale},${locConfig.locale.split('-')[0]};q=0.9,en;q=0.8`,
        'Upgrade-Insecure-Requests': '1',
        'Sec-Fetch-Site': 'none',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-User': '?1',
        'Sec-Fetch-Dest': 'document'
      },
      ignoreHTTPSErrors: true
    });

    const cookies = [
      { name: 'lc-main', value: locConfig.locale.replace('-', '_'), domain: `.${domain}`, path: '/' },
      { name: 'i18n-prefs', value: locConfig.currency, domain: `.${domain}`, path: '/' },
      { name: 'sp-cdn', value: `"${locConfig.city}:${locConfig.zip}"`, domain: `.${domain}`, path: '/' },
      { name: 'session-id', value: '123-' + Math.floor(Math.random() * 9000000 + 1000000) + '-' + Math.floor(Math.random() * 9000000 + 1000000), domain: `.${domain}`, path: '/' },
      { name: 'ubid-main', value: '123-' + Math.floor(Math.random() * 9000000 + 1000000) + '-' + Math.floor(Math.random() * 9000000 + 1000000), domain: `.${domain}`, path: '/' },
      { name: 'session-token', value: 'ST-' + Math.random().toString(36).substring(2), domain: `.${domain}`, path: '/' }
    ];
    await context.addCookies(cookies);

    const page = await context.newPage();
    
    await page.route('**/*', (route) => {
      const rUrl = route.request().url();
      const resourceType = route.request().resourceType();
      if (['image', 'font', 'media'].includes(resourceType)) {
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

    try {
      console.log(`Auditing Amazon ${asin} on ${domain} (Target Zip: ${locConfig.zip})...`);
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
      
      try {
        const cookieButtons = ['#sp-cc-accept', 'input[name="accept"]', '#cookie-accept', '#accept-cookies', '.a-button-inner input[data-action="accept-cookies"]'];
        for (const selector of cookieButtons) {
          if (await page.isVisible(selector)) {
            await page.click(selector).catch(() => null);
            await page.waitForTimeout(500);
            break;
          }
        }
      } catch (err) { /* ignored */ }

      try {
        console.log(`UI Regional Unlock: Evaluating postcode injection for ${domain} (Target: ${locConfig.zip})`);
        
        for (let attempt = 1; attempt <= 3; attempt++) {
          const isRegionalLocked = await page.evaluate(({ zip }) => {
            const slot = document.querySelector('#nav-global-location-slot');
            if (!slot) return true;
            const text = slot.textContent || '';
            const zipFirstPart = zip.split(/[\s-]+/)[0];
            return !text.toLowerCase().includes(zip.toLowerCase()) && !text.toLowerCase().includes(zipFirstPart.toLowerCase());
          }, { zip: locConfig.zip });

          if (!isRegionalLocked) {
            console.log(`Location successfully verified as applied! Slot is set to: ${locConfig.zip}`);
            break;
          }

          console.log(`Location slot text doesn't contain ${locConfig.zip} (Attempt ${attempt}/3). Performing injection steps...`);
          
          const locBtn = await page.waitForSelector('#nav-global-location-slot, #glow-ingress-block, #nav-main-ftr-location-slot', { state: 'visible', timeout: 10000 }).catch(() => null);
          let popoverOpened = false;
          
          if (locBtn) {
            for (let clickAttempt = 1; clickAttempt <= 3; clickAttempt++) {
              console.log(`Clicking location button (attempt ${clickAttempt})...`);
              await locBtn.click({ force: true }).catch(() => null);
              await page.waitForTimeout(1500);
              const isVisible = await page.locator('.a-popover-modal, .a-popover, #GLUXZipUpdateInput, #GLUXZipUpdateInput_0, #GLUXCountryList').isVisible().catch(() => false);
              if (isVisible) {
                popoverOpened = true;
                break;
              }
            }
          }

          if (!popoverOpened) {
            console.warn('Popover never opened.');
            break;
          }

          const zipInputSelector = '#GLUXZipUpdateInput, #GLUXZipUpdateInput_0, input[aria-label*="zip" i], input[aria-label*="postcode" i], input[aria-label*="code" i], input[name="zipCode" i]';
          const countryListSelector = '#GLUXCountryList';
          
          await page.waitForSelector(`${zipInputSelector}, ${countryListSelector}`, { state: 'visible', timeout: 5000 }).catch(() => null);
          
          let zipInput = await page.$(zipInputSelector);
          let zipVisible = zipInput ? await zipInput.isVisible().catch(() => false) : false;
          
          if (!zipVisible) {
            console.log('Zip input not immediately visible, handling country select or list fallback...');
            const countryListVisible = await page.locator(countryListSelector).isVisible().catch(() => false);
            
            if (countryListVisible && locConfig.countryCode) {
              console.log(`Country dropdown detected. Attempting to select country: ${locConfig.countryCode}`);
              
              let selectedVal = null;
              try {
                const options = await page.$$eval(`${countryListSelector} option`, (opts: any[]) => 
                  opts.map(o => ({ value: o.value, text: o.textContent?.trim() || "" }))
                );
                const targetCode = locConfig.countryCode.toLowerCase();
                const matchByVal = options.find(o => o.value.toLowerCase() === targetCode);
                if (matchByVal) {
                  selectedVal = matchByVal.value;
                } else {
                  const countryNamesMap: Record<string, string[]> = {
                    'GB': ['united kingdom', 'royaume-uni', 'vereinigt', 'regno unito', 'reino unido', 'wielka brytania', 'storbritannien', 'england'],
                    'FR': ['france', 'frankreich', 'francia'],
                    'DE': ['germany', 'deutschland', 'allemagne', 'germania', 'niemcy', 'tyskland'],
                    'IT': ['italy', 'itálie', 'italie', 'italien', 'italia', 'włochy'],
                    'ES': ['spain', 'spanien', 'espagne', 'spagna', 'españa', 'hiszpania'],
                    'PL': ['poland', 'polen', 'pologne', 'polonia', 'polska'],
                    'NL': ['netherlands', 'niederlande', 'pays-bas', 'paesi bassi', 'países bajos', 'holandia', 'nederländerna', 'nederland'],
                    'SE': ['sweden', 'schweden', 'suède', 'svezia', 'suecia', 'szwecja', 'sverige']
                  };
                  const names = countryNamesMap[locConfig.countryCode] || [];
                  const matchByName = options.find(o => {
                    const txt = o.text.toLowerCase();
                    return names.some(n => txt.includes(n));
                  });
                  if (matchByName) {
                    selectedVal = matchByName.value;
                  }
                }
              } catch (e: any) {
                console.warn("Failed to read country options:", e.message);
              }
              
              if (selectedVal) {
                console.log(`Selecting country option: ${selectedVal}`);
                await page.selectOption(countryListSelector, { value: selectedVal }).catch(() => null);
              } else {
                console.log(`Fallback selecting country code: ${locConfig.countryCode}`);
                await page.selectOption(countryListSelector, { value: locConfig.countryCode }).catch(() => null);
              }
              
              await page.waitForTimeout(1000);
              const countryApplySelectors = [
                '#GLUXCountryListDropdown .a-button-input',
                'input[aria-labelledby="GLUXCountryList-announce"]',
                '#GLUXCountryList-announce ~ input',
                '.a-popover-footer input[type="submit"]',
                '.a-popover-footer .a-button-input',
                'button[name="glowDoneButton"]',
                '#a-popover-1 .a-button-input',
                '.a-popover-footer input'
              ];
              let countryApplied = false;
              for (const sel of countryApplySelectors) {
                try {
                  const btn = await page.$(sel);
                  if (btn && await btn.isVisible()) {
                    await btn.click({ force: true });
                    countryApplied = true;
                    console.log(`Country applied via: ${sel}`);
                    break;
                  }
                } catch (_) {}
              }
              if (!countryApplied) {
                await page.keyboard.press('Enter').catch(() => null);
              }
              await page.waitForTimeout(2500);
              await page.waitForLoadState('domcontentloaded', { timeout: 10000 }).catch(() => null);
              
              try {
                const dismissBtn = await page.$('button[name="glowDoneButton"], #GLUXConfirmClose input');
                if (dismissBtn && await dismissBtn.isVisible()) {
                  await dismissBtn.click({ force: true });
                  await page.waitForTimeout(1000);
                }
              } catch (_) {}

              console.log(`Reloading ${domain} after country selection...`);
              await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 }).catch(() => null);
            }
          }

          zipInput = await page.$(zipInputSelector);
          zipVisible = zipInput ? await zipInput.isVisible().catch(() => false) : false;

          if (zipInput && zipVisible) {
            console.log(`Entering zip code/postcode: "${locConfig.zip}"`);
            await zipInput.click({ clickCount: 3 }).catch(() => null);
            await page.keyboard.press('Backspace').catch(() => null);
            await zipInput.fill('');
            await zipInput.type(locConfig.zip, { delay: 50 });
            await page.waitForTimeout(500);
            
            const applySelectors = [
              '#GLUXZipUpdate input[type="submit"]',
              '#GLUXZipUpdate .a-button-input',
              '#GLUXZipUpdate > span > input',
              '#GLUXZipUpdate_Buttons input',
              '#GLUXZipUpdate_Buttons span.a-button-inner input',
              '#GLUXZipUpdate input',
              'input[aria-labelledby="GLUXZipUpdate-announce"]',
              '#GLUXZipUpdate-announce ~ input'
            ];
            let applied = false;
            for (const sel of applySelectors) {
              try {
                const btn = await page.$(sel);
                if (btn && await btn.isVisible()) {
                  await btn.click({ force: true });
                  applied = true;
                  console.log(`Apply clicked via: ${sel}`);
                  break;
                }
              } catch (_) {}
            }
            if (!applied) {
              await page.keyboard.press('Enter').catch(() => null);
            }
            
            await page.waitForTimeout(2000);
            const confirmSelectors = [
              '#GLUXConfirmClose input',
              '#GLUXConfirmClose .a-button-input',
              'input[data-action="GLUXConfirmResponse"]',
              '#GLUXConfirmClose-announce',
              'button[name="glowDoneButton"]',
              '.a-popover-footer .a-button-input',
              '.a-popover-footer input',
              'button:has-text("Done")',
              'button:has-text("Confirm")',
              'button:has-text("Continue")',
              'input[type="button"]:has-text("Done")',
              'span.a-button:has-text("Done") input',
              'span.a-button:has-text("Continue") input'
            ];
            for (const sel of confirmSelectors) {
              try {
                const btn = await page.$(sel);
                if (btn && await btn.isVisible()) {
                  await btn.click({ force: true });
                  console.log(`Confirmed done button via: ${sel}`);
                  break;
                }
              } catch (_) {}
            }
            await page.waitForTimeout(1500);
            console.log(`Reloading ${domain} after postcode updates to finalize location change...`);
            await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 }).catch(() => null);
          } else {
            console.warn('Zip input became unavailable or skipped after country select.');
          }
        }

        await page.waitForTimeout(800);
        const finalLocText = await page.evaluate(() => {
          const slot = document.querySelector('#nav-global-location-slot');
          return slot ? slot.textContent?.replace(/\s+/g, ' ').trim() : '';
        }).catch(() => '');
        console.log(`Location slot after injection: "${finalLocText}"`);
      } catch (err: any) {
        console.warn("Location UI injection skipped or failed:", err.message);
      }

      await page.waitForLoadState('domcontentloaded', { timeout: 10000 }).catch(() => null);

      try {
        const oneTimeSelectors = [
          '#newAccordionRow_header',
          '#newAccordionRow',
          '#oneTimeBuyBox_feature_div',
          '#buyNewSection_header',
          '#oneTimeBuyBox',
          '#buyBoxAccordion .a-accordion-row:not(#subscribeAndSaveAccordionRow) .a-accordion-header',
          '#buyBoxAccordion div[id*="oneTime"]',
          'div[id*="oneTimeBuyBox"]',
          '.oneTimeBuyBox',
          'input[id*="oneTime"]'
        ];
        for (const selector of oneTimeSelectors) {
          const btn = await page.$(selector);
          if (btn && await btn.isVisible()) {
            const isSelected = await page.evaluate((el: any) => {
              return el.classList.contains('a-accordion-row-active') || el.classList.contains('active') || !!el.querySelector('.a-accordion-row-active');
            }, btn).catch(() => false);
            
            if (!isSelected) {
              console.log(`Clicking One-Time Purchase header/accordion via: ${selector}`);
              await btn.click({ force: true });
              await page.waitForTimeout(1500);
            }
            break;
          }
        }
      } catch (err: any) {
        console.warn("Could not click One-Time Purchase accordion:", err.message);
      }
    } catch (e: any) {
      console.error("Navigation error:", e.message);
    }

    const content = await page.content();
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

    // --- 3. Buybox Owner Extraction ---
    let amazonBuyboxOwner = "";

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
    
    let soldByTabular = "";
    const soldByTabularEls = buyBoxContext.find('div[tabular-attribute-name*="Sold by" i] .tabular-buybox-text, div[tabular-attribute-name*="Verkauf durch" i] .tabular-buybox-text, div[tabular-attribute-name*="Vendido por" i] .tabular-buybox-text, div[tabular-attribute-name*="Vendu par" i] .tabular-buybox-text, div[tabular-attribute-name*="Sprzedawca" i] .tabular-buybox-text, div[tabular-attribute-name*="Säljs av" i] .tabular-buybox-text');
    soldByTabularEls.each((_, el) => {
      if ($(el).closest('#subscribeAndSaveAccordionRow, [class*="sns"], [id*="sns"]').length === 0) {
        soldByTabular = $(el).text().trim();
        return false;
      }
    });
    if (!soldByTabular && soldByTabularEls.length > 0) {
      soldByTabular = soldByTabularEls.first().text().trim();
    }

    let dispatchedByTabular = "";
    const dispatchedByTabularEls = buyBoxContext.find('div[tabular-attribute-name*="Dispatched" i] .tabular-buybox-text, div[tabular-attribute-name*="Verzonden" i] .tabular-buybox-text, div[tabular-attribute-name*="Shipped" i] .tabular-buybox-text, div[tabular-attribute-name*="Expédié" i] .tabular-buybox-text, div[tabular-attribute-name*="Invia" i] .tabular-buybox-text, div[tabular-attribute-name*="Wysyłka" i] .tabular-buybox-text, div[tabular-attribute-name*="Skickas från" i] .tabular-buybox-text');
    dispatchedByTabularEls.each((_, el) => {
      if ($(el).closest('#subscribeAndSaveAccordionRow, [class*="sns"], [id*="sns"]').length === 0) {
        dispatchedByTabular = $(el).text().trim();
        return false;
      }
    });
    if (!dispatchedByTabular && dispatchedByTabularEls.length > 0) {
      dispatchedByTabular = dispatchedByTabularEls.first().text().trim();
    }

    const isAmazonSeller = 
      /sold by amazon/i.test(mainMerchantText) ||
      /dispatched from and sold by amazon/i.test(mainMerchantText) ||
      /vendu par amazon/i.test(mainMerchantText) ||
      /verkauf durch amazon/i.test(mainMerchantText) ||
      /verzonden en verkocht door amazon/i.test(mainMerchantText) ||
      /spedito da e venduto da amazon/i.test(mainMerchantText) ||
      /vendido por amazon/i.test(mainMerchantText) ||
      /verkauft von amazon/i.test(mainMerchantText) ||
      /sprzedawane przez amazon/i.test(mainMerchantText) ||
      /säljs av amazon/i.test(mainMerchantText) ||
      /amazon/i.test(soldByTabular);

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
        buyBoxContext.find('div[tabular-attribute-name="Shipper / Seller"] .tabular-buybox-text').first().text().trim() ||
        buyBoxContext.find('div[offer-display-feature-name="desktop-merchant-info"] .offer-display-feature-text-message').first().text().trim() ||
        sellerTrigger ||
        merchantLink;
    }

    if (!amazonBuyboxOwner) {
      amazonBuyboxOwner = 
        $('div[tabular-attribute-name="Sold by"] .tabular-buybox-text').first().text().trim() ||
        $('div[tabular-attribute-name="Verkauf durch"] .tabular-buybox-text').first().text().trim() ||
        $('div[tabular-attribute-name="Shipper / Seller"] .tabular-buybox-text').first().text().trim() ||
        $('div[offer-display-feature-name="desktop-merchant-info"] .offer-display-feature-text-message').first().text().trim() ||
        $('#sellerProfileTriggerId').first().text().trim() ||
        $('#merchant-info a').first().text().trim();
    }

    if (!amazonBuyboxOwner && mainMerchantText.length > 0) {
      const sellerLink = mainMerchantEl ? (mainMerchantEl as any).find('a').first().text().trim() || $('#merchant-info a').first().text().trim() : "";
      if (sellerLink && sellerLink.toLowerCase() !== 'amazon') {
        amazonBuyboxOwner = sellerLink;
      } else {
        amazonBuyboxOwner = mainMerchantText
          .replace(/Dispatched from and sold by\s*/i, '')
          .replace(/Dispatched from Amazon\s*\.?\s*/i, '')
          .replace(/Sold by\s*/i, '')
          .replace(/Fulfilled by Amazon\s*\.?\s*/i, '')
          .replace(/\|.*$/s, '')
          .trim();
      }
    }

    amazonBuyboxOwner = amazonBuyboxOwner
      .replace(/Sold by\s*:?\s*/gi, '')
      .replace(/Venduto da\s*:?\s*/gi, '')
      .replace(/Verkauf durch\s*:?\s*/gi, '')
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
      '#feature-bullets ul li:not(:has(ul))',
      '#featurebullets_feature_div ul li:not(:has(ul))',
      '#feature-bullets-content li:not(:has(ul))',
      '[data-feature-name="product-facts"] .a-list-item',
      '.product-facts-title + .a-unordered-list li:not(:has(ul))',
      '#product-facts-grid li:not(:has(ul))',
      '#productFactsDesktopExpander .a-list-item',
      '#feature-bullets .a-list-item',
      '#featurebullets_feature_div .a-list-item',
      '#feature-bullets ul li',
      '#featurebullets_feature_div ul li',
      '#aboutThisItem ul.a-unordered-list.a-vertical li'
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

        const isTechnicalMeta =
          /asin\s*:/i.test(lower) ||
          /ean\s*:/i.test(lower) ||
          /isbn\s*:/i.test(lower) ||
          /fabricante\s*:/i.test(lower) ||
          /hersteller\s*:/i.test(lower) ||
          /manufacturer\s*:/i.test(lower) ||
          /brand\s*:/i.test(lower) ||
          /marca\s*:/i.test(lower) ||
          /referencia\s+del\s+fabricante/i.test(lower) ||
          /manufacturer\s+reference/i.test(lower) ||
          /part\s+number/i.test(lower) ||
          /número\s+de\s+modelo/i.test(lower) ||
          /model\s+number/i.test(lower) ||
          /dimensiones/i.test(lower) ||
          /dimensions/i.test(lower) ||
          /peso/i.test(lower) ||
          /weight/i.test(lower) ||
          /producto\s+en\s+amazon/i.test(lower) ||
          /product\s+since/i.test(lower) ||
          /disponible\s+desde/i.test(lower) ||
          /available\s+since/i.test(lower) ||
          /date\s+first\s+available/i.test(lower) ||
          /opiniones\s+de\s+los/i.test(lower) ||
          /customer\s+reviews/i.test(lower);

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
          isBSR ||
          isTechnicalMeta
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
      amazonTitle.length < 3 ||
      !amazonPrice || 
      amazonPrice === 'N/A' || 
      !amazonBuyboxOwner || 
      amazonBuyboxOwner === 'N/A' ||
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
