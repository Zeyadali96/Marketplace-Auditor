import express from 'express';
import { chromium } from 'playwright';
import { chromium as chromiumExtra } from 'playwright-extra';
import stealth from 'puppeteer-extra-plugin-stealth';
import * as cheerio from 'cheerio';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import axios from 'axios';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';
import stringSimilarity from 'string-similarity';
import { GoogleGenAI } from '@google/genai';

chromiumExtra.use(stealth());

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json({ limit: '50mb' }));

// Helper for string similarity
const getSimilarity = (str1: string, str2: string) => {
  if (!str1 || !str2) return 0;
  return stringSimilarity.compareTwoStrings(str1.toLowerCase(), str2.toLowerCase());
};

// 1. Proxy Image API
app.get('/api/proxy-image', async (req, res) => {
  const imageUrl = req.query.url as string;
  if (!imageUrl) return res.status(400).send('No URL provided');
  
  try {
    const response = await axios.get(imageUrl, {
      responseType: 'arraybuffer',
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
        'Referer': 'https://www.amazon.com/'
      }
    });
    const contentType = response.headers['content-type'];
    res.setHeader('Content-Type', contentType);
    res.send(response.data);
  } catch (error) {
    res.status(500).send('Error proxying image');
  }
});

// Helper to normalize price strings to dots
function cleanAndNormalizePrice(priceStr: string): string {
  if (!priceStr) return "";
  let s = priceStr.trim();
  // Remove space between numbers (e.g. "1 250,50" -> "1250,50")
  s = s.replace(/\s/g, '');
  
  if (s.includes('.') && s.includes(',')) {
    if (s.indexOf('.') > s.indexOf(',')) {
      s = s.replace(/,/g, '');
    } else {
      s = s.replace(/\./g, '').replace(/,/g, '.');
    }
  } else if (s.includes(',')) {
    const parts = s.split(',');
    const lastPart = parts[parts.length - 1].replace(/[^0-9]/g, '');
    if (lastPart.length === 2 || lastPart.length === 1) {
      s = s.replace(/,/g, '.');
    } else {
      s = s.replace(/,/g, '');
    }
  }
  // Select characters and decimals
  const match = s.match(/\d+(\.\d+)?/);
  return match ? match[0] : s.replace(/[^0-9.]/g, '');
}

// 2. Audit Amazon
app.post("/api/audit/amazon", async (req, res) => {
  let browser;
  try {
    const { asin, marketplace, masterData } = req.body;
    const domain = marketplace || 'amazon.com';
    const url = `https://www.${domain}/dp/${asin}`;
    
    const proxyServer = process.env.PROXY_SERVER;
    const proxyUsername = process.env.PROXY_USERNAME;
    const proxyPassword = process.env.PROXY_PASSWORD;
    
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
      const url = route.request().url();
      const resourceType = route.request().resourceType();
      // Block images, fonts, stylesheets, media, and known tracking/analytics domains
      if (['image', 'font', 'media'].includes(resourceType)) {
        return route.abort();
      }
      if (resourceType === 'stylesheet' && !url.includes('amazon')) {
        return route.abort();
      }
      const blockPatterns = [
        'google-analytics', 'googletagmanager', 'doubleclick',
        'facebook.net', 'fbcdn', 'adsystem', 'advertising-api',
        'amazon-adsystem', 'fls-na.amazon', 'unagi.amazon',
        'completion.amazon', 'aax-', 'mads.'
      ];
      if (blockPatterns.some(p => url.includes(p))) {
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
        const isRegionalLocked = await page.evaluate(({ zip }) => {
          const slot = document.querySelector('#nav-global-location-slot');
          if (!slot) return true;
          const text = slot.textContent || '';
          const zipFirstPart = zip.split(/[\s-]+/)[0];
          return !text.toLowerCase().includes(zip.toLowerCase()) && !text.toLowerCase().includes(zipFirstPart.toLowerCase());
        }, { zip: locConfig.zip });

        if (isRegionalLocked) {
          console.log(`UI Regional Unlock: Injecting ${locConfig.zip} for ${domain} (country: ${locConfig.countryCode})`);
          
          const locBtn = await page.waitForSelector('#nav-global-location-slot, #glow-ingress-block, #nav-main-ftr-location-slot', { state: 'visible', timeout: 10000 }).catch(() => null);
          let popoverOpened = false;
          
          if (locBtn) {
            for (let clickAttempt = 1; clickAttempt <= 3; clickAttempt++) {
              console.log(`Clicking location button (attempt ${clickAttempt})...`);
              await locBtn.click({ force: true }).catch(() => null);
              await page.waitForTimeout(1500);
              const isVisible = await page.locator('.a-popover-modal, .a-popover, #GLUXZipUpdateInput, #GLUXCountryList').isVisible().catch(() => false);
              if (isVisible) {
                popoverOpened = true;
                break;
              }
            }
          }

          if (popoverOpened) {
            const zipInputSelector = '#GLUXZipUpdateInput, #GLUXZipUpdateInput_0, input[aria-label*="zip"], input[aria-label*="postcode"], input[aria-label*="code"], input[name="zipCode"]';
            const countryListSelector = '#GLUXCountryList';
            
            // Wait for either Zip input or Country list dropdown
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
                  // Try to find by value match
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
                
                // Click Apply/Done/Go button for country selection
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
                
                // Wait for reload or DOM change
                await page.waitForTimeout(2500);
                await page.waitForLoadState('domcontentloaded', { timeout: 10000 }).catch(() => null);
                
                // Try to dismiss overlays if any
                try {
                  const dismissBtn = await page.$('button[name="glowDoneButton"], #GLUXConfirmClose input');
                  if (dismissBtn && await dismissBtn.isVisible()) {
                    await dismissBtn.click({ force: true });
                    await page.waitForTimeout(1000);
                  }
                } catch (_) {}

                // Re-open location popover for ZIP entry
                console.log('Re-opening location popover for ZIP entry...');
                let locBtn2 = null;
                for (let clickAttempt2 = 1; clickAttempt2 <= 3; clickAttempt2++) {
                  locBtn2 = await page.waitForSelector('#nav-global-location-slot, #glow-ingress-block, #nav-main-ftr-location-slot', { state: 'visible', timeout: 5000 }).catch(() => null);
                  if (locBtn2) {
                    await locBtn2.click({ force: true }).catch(() => null);
                    await page.waitForTimeout(1500);
                    const checkZip = await page.$(zipInputSelector);
                    if (checkZip && await checkZip.isVisible()) {
                      zipInput = checkZip;
                      zipVisible = true;
                      break;
                    }
                  }
                }
              }
            }

            if (zipInput || zipVisible) {
              if (!zipInput) {
                zipInput = await page.$(zipInputSelector);
              }
              if (zipInput) {
                console.log(`Entering zip code/postcode: "${locConfig.zip}"`);
                await zipInput.click({ clickCount: 3 }).catch(() => null);
                await page.keyboard.press('Backspace').catch(() => null);
                await zipInput.fill('');
                await zipInput.type(locConfig.zip, { delay: 50 });
                await page.waitForTimeout(500);
                
                // Click Apply button
                const applySelectors = [
                  '#GLUXZipUpdate input[type="submit"]',
                  '#GLUXZipUpdate .a-button-input',
                  '#GLUXZipUpdate > span > input',
                  '#GLUXZipUpdate_Buttons input',
                  '#GLUXZipUpdate_Buttons span.a-button-inner input',
                  '#GLUXZipUpdate input',
                  'input[aria-labelledby="GLUXZipUpdate-announce"]'
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
                
                // Click confirm/done if shown
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
                
                // Reload to apply the new delivery address
                console.log(`Reloading ${domain} after location injection to apply ZIP change...`);
                await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 }).catch(() => null);
              }
            } else {
              console.warn('Zip input never appeared or became visible.');
            }
          } else {
            console.warn('Popover never opened.');
          }
          
          // Verify location was set
          await page.waitForTimeout(800);
          const finalLocText = await page.evaluate(() => {
            const slot = document.querySelector('#nav-global-location-slot');
            return slot ? slot.textContent?.replace(/\s+/g, ' ').trim() : '';
          }).catch(() => '');
          console.log(`Location slot after injection: "${finalLocText}"`);
        }
      } catch (err: any) {
        console.warn("Location UI injection skipped or failed:", err.message);
      }

      await page.waitForLoadState('domcontentloaded', { timeout: 10000 }).catch(() => null);

      // Ensure "One-Time Purchase" option is selected if "Subscribe & Save" is default on the page
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

    // Extraction Logic
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
    // FIXED buyBoxContext — add the missing 3P FBA containers, in priority order
    let buyBoxContext = $('#oneTimeBuyBox').length ? $('#oneTimeBuyBox') :
                        $('#newAccordionRow').length ? $('#newAccordionRow') :
                        $('#buyNewSection').length ? $('#buyNewSection') :
                        $('#apex_offerDisplay_desktop').length ? $('#apex_offerDisplay_desktop') :
                        $('#newUnifiedOfferDisplay').length ? $('#newUnifiedOfferDisplay') :
                        $('#corePrice_feature_div').length ? $('#corePrice_feature_div') :
                        $('#desktop_buybox').length ? $('#desktop_buybox') :
                        $('#rightCol').length ? $('#rightCol') :
                        $('body');

    // FIXED amazonPrice — prioritise the buybox container, then specific 3P price IDs,
    // then scoped fallbacks. The LAST resort ($('body') scan) is intentionally removed
    // to prevent bleeding from "Frequently bought together" / carousel sections.
    let amazonPrice =
      // 1. Buybox-scoped priceToPay (works for both Amazon-sold and 3P sellers)
      buyBoxContext.find('.priceToPay .a-offscreen').first().text().trim() ||
      // 2. 3P FBA seller price containers (apex unified offer display)
      $('#apex_offerDisplay_desktop .priceToPay .a-offscreen').first().text().trim() ||
      $('#apex_offerDisplay_desktop .a-price .a-offscreen').first().text().trim() ||
      $('#newUnifiedOfferDisplay .priceToPay .a-offscreen').first().text().trim() ||
      $('#newUnifiedOfferDisplay .a-price .a-offscreen').first().text().trim() ||
      // 3. Amazon-sold price containers
      $('#corePriceDisplay_desktop_feature_div .priceToPay .a-offscreen').first().text().trim() ||
      $('#corePrice_desktop .priceToPay .a-offscreen').first().text().trim() ||
      $('#corePriceDisplay_desktop_feature_div .a-price .a-offscreen').first().text().trim() ||
      $('#corePrice_desktop .a-price .a-offscreen').first().text().trim() ||
      // 4. corePrice_feature_div (3P FBA on .co.uk / .de / .nl)
      $('#corePrice_feature_div .priceToPay .a-offscreen').first().text().trim() ||
      $('#corePrice_feature_div .a-price .a-offscreen').first().text().trim() ||
      // 5. Scoped buybox fallbacks (already present)
      $('#buyNew_noncbb .a-price .a-offscreen').first().text().trim() ||
      $('#buyNewSection .a-price .a-offscreen').first().text().trim() ||
      $('#desktop_buybox .a-price .a-offscreen').first().text().trim() ||
      $('#price_inside_buybox').text().trim() ||
      $('.apex-core-price-identifier .a-offscreen').first().text().trim() ||
      $('#desktop_buybox .apexPriceToPay .a-offscreen').first().text().trim() ||
      $('#desktop_buybox .priceToPay .a-offscreen').first().text().trim() ||
      // 6. rightCol scoped (safe because rightCol excludes product carousels)
      $('#rightCol .priceToPay .a-offscreen').first().text().trim() ||
      $('#rightCol .a-price .a-offscreen').first().text().trim() ||
      "";
    amazonPrice = cleanAndNormalizePrice(amazonPrice);

    let listPrice = $('.basisPrice .a-offscreen').text().trim() || "";
    listPrice = cleanAndNormalizePrice(listPrice);

    // Parse FREE Delivery Shipping Time (ignoring Prime expedited / fastest options)
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

    // Prioritize parsing the primary or secondary text if it directly contains "FREE delivery" and is NOT Prime expedited
    if (primaryDelivery && hasFreeDeliveryKeyword(primaryDelivery) && !hasPrimeExpeditedKeyword(primaryDelivery)) {
      rawShippingTime = primaryDelivery;
    } else if (secondaryDelivery && hasFreeDeliveryKeyword(secondaryDelivery) && !hasPrimeExpeditedKeyword(secondaryDelivery)) {
      rawShippingTime = secondaryDelivery;
    } else if (primaryDelivery && hasFreeDeliveryKeyword(primaryDelivery)) {
      // If the sentence contains both (e.g. fast expedited + standard free), split and find the FREE delivery portion
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
      // If not found in main slots directly, inspect the entire DELIVERY_BLOCK for any matching child node
      const deliveryBlock = $('#mir-layout-DELIVERY_BLOCK-slot-PRIMARY_DELIVERY_MESSAGE_LARGE, #mir-layout-DELIVERY_BLOCK, #deliveryBlockMessage');
      
      let foundLine = "";
      deliveryBlock.find('*').each((_, el) => {
        const txt = $(el).text().replace(/\s+/g, ' ').trim();
        if (txt && txt.length > 5 && txt.length < 150 && hasFreeDeliveryKeyword(txt) && !hasPrimeExpeditedKeyword(txt)) {
          foundLine = txt;
          return false; // break loop
        }
      });

      if (!foundLine) {
        // Try split whole delivery block text by newlines
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
        // Ultimate fallback to data-csa-c-delivery-time attribute or text bold or primary
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

    // Inlined helper function for precise and language-aware shipping text extraction and cleaning
    const cleanShippingText = (text: string): string => {
      if (!text) return "";
      
      // 1. Deduplicate sentences/phrases to handle DOM-level repetition (ideal for FR "Livraison GRATUITE..." repeats)
      const segments = text.split(/\s*[\.\!\?]+\s*/).map(s => s.trim()).filter(s => s.length > 0);
      const uniqueSegments: string[] = [];
      for (const seg of segments) {
        if (!uniqueSegments.some(us => us.toLowerCase() === seg.toLowerCase() || seg.toLowerCase().includes(us.toLowerCase()) || us.toLowerCase().includes(seg.toLowerCase()))) {
          uniqueSegments.push(seg);
        }
      }
      let cleaned = uniqueSegments.join('. ').trim();

      // 2. Remove localized free-delivery / marketing prefixes
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

      // 3. Remove localized suffixes / marketing qualifiers (such as first order qualifiers)
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

      // 4. Strip leading language particles and prepositions safely
      cleaned = cleaned
        .replace(/^(?:le|la|el|los|on|op|am|op|przy|w|v|at|by|from|auf|de|di|d'|el)\s+/gi, '')
        .trim();

      // 5. Trim residual punctuation and clean whitespace
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
      // Clean up trailing/leading garbage from punctuation if fallback
      rawShippingTime = rawShippingTime
        .replace(/^[\s,.;:or|]+|[\s,.;:or|]+$/gi, '')
        .replace(/\s+/g, ' ')
        .trim();
    }

    let amazonDesc = $('#productDescription').text().trim();
    if (!amazonDesc) amazonDesc = $('#feature-bullets').text().trim();
    
    const hasAPlus = !!($('#aplus').length || $('#aplus_feature_div').length || $('div[id*="aplus"]').length);

    // --- 3. Buybox Owner Extraction (Tiered Priority Resolver) ---
    let amazonBuyboxOwner = "";

    // Helper to find parent buybox container excluding subscribe & save accordion
    let mainMerchantEl: any = null;
    const merchantEls = $('#merchant-info');
    merchantEls.each((_, el) => {
      if ($(el).closest('#subscribeAndSaveAccordionRow, [class*="sns"], [id*="sns"]').length === 0) {
        mainMerchantEl = $(el);
        return false; // break
      }
    });
    if (!mainMerchantEl && merchantEls.length > 0) {
      mainMerchantEl = merchantEls.first();
    }
    const mainMerchantText = mainMerchantEl ? mainMerchantEl.text().replace(/\s+/g, ' ').trim() : "";
    
    // Check for tabular attributes in the buybox context, excluding SNS
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

    // Check if Amazon itself is the seller based on merchant info or tabular Sold by attributes
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

    // Priority 2: Scopes to buyBoxContext tabular displays or trigger links, excluding SNS
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

    // Priority 3: Fallback ONLY if scoping of buyBoxContext yielded nothing
    if (!amazonBuyboxOwner) {
      amazonBuyboxOwner = 
        $('div[tabular-attribute-name="Sold by"] .tabular-buybox-text').first().text().trim() ||
        $('div[tabular-attribute-name="Verkauf durch"] .tabular-buybox-text').first().text().trim() ||
        $('div[tabular-attribute-name="Shipper / Seller"] .tabular-buybox-text').first().text().trim() ||
        $('div[offer-display-feature-name="desktop-merchant-info"] .offer-display-feature-text-message').first().text().trim() ||
        $('#sellerProfileTriggerId').first().text().trim() ||
        $('#merchant-info a').first().text().trim();
    }

    // Priority 4: Parse original merchant text to sanitize the candidate or extract from text node
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

    // Clean residual prefixes from any extraction path
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

    const getNormalizedInfo = (url: string | undefined) => {
      if (!url || typeof url !== "string" || url.length < 15) return null;

      // 1. Clean query strings and normalize protocol/trim
      let cleaned = url.split("?")[0].trim();
      if (cleaned.startsWith("//")) cleaned = "https:" + cleaned;
      cleaned = cleaned.replace(/^http:/, "https:");

      // 2. Aggressively clean Amazon modifiers (e.g., ._AC_SX679_.)
      // These can appear in various formats. Stripping them reveals the base image URL.
      cleaned = cleaned.replace(/\._[a-zA-Z0-9,_-]+_\.?/g, ".");
      cleaned = cleaned.replace(/\.V[0-9]+_\.?/g, ".");
      cleaned = cleaned.replace(/\.(V|SS|SX|SY|AC|SR|SL|UL|CLa|SR|SS|SX|SY|UL|CLa)[0-9,s]+_\./g, ".");
      cleaned = cleaned.replace(/\.(V|SS|SX|SY|AC|SR|SL|UL|CLa|SR|SS|SX|SY|UL|CLa)[0-9,s]+\./g, ".");

      // 3. Identify the Amazon Image ID block for identity tracking
      const idMatch = cleaned.match(/\/images\/(I|W|S|G)\/([^\.\/]+)/) || cleaned.match(/\/images\/([^\.\/]+)\./);
      if (!idMatch) return null;

      const baseId = ((idMatch.length > 2 ? idMatch[2] : idMatch[1]) as string).replace(/\.+$/, "");
      
      // Exclude suspected Swatch images, generic icons or play buttons
      if (baseId.includes("SWCH") || cleaned.includes("_SW") || baseId.startsWith("ss_") || cleaned.includes("play-button")) return null;

      return { baseId, url: cleaned };
    };

    // 1. Always prioritize the main hero image first as part of insertion order
    const mainHeroUrl = $("#landingImage").attr("data-old-hires") || $("#landingImage").attr("src");
    const heroInfo = getNormalizedInfo(mainHeroUrl);
    
    console.log("=== AMAZON IMAGE EXTRACTION DEBUG ===");
    console.log("Hero URL:", mainHeroUrl);
    console.log("Hero Info:", heroInfo);
    
    if (heroInfo) {
      imageMap.set(heroInfo.baseId, heroInfo.url);
      console.log("Hero added to map with baseId:", heroInfo.baseId);
    }

    // 2. Gather thumbnails while ignoring duplicates of the hero or other images
    let thumbCount = 0;
    $("#altImages li.imageThumbnail:not(.videoThumbnail) img, .imageThumbnail img, .altImages img").not("#landingImage").each((_, el) => {
      thumbCount++;
      const elemId = $(el).attr("id");
      const src = $(el).attr("data-old-hires") || $(el).attr("src");
      
      console.log(`Thumbnail ${thumbCount} (id: ${elemId || 'none'}):`, src);
      
      if (!src || src.includes("pixel.gif") || src.includes("play-button-overlay") || src.includes("transparent-pixel")) {
        console.log("  -> SKIP: pixel/placeholder");
        return;
      }

      // Re-applying the surgical fix for hero duplication via raw src check
      const rawSrcInfo = getNormalizedInfo($(el).attr("src"));
      if (heroInfo && rawSrcInfo && rawSrcInfo.baseId === heroInfo.baseId) {
        console.log("  -> SKIP: matches hero baseId (raw src check)");
        return;
      }

      const info = getNormalizedInfo(src);
      if (!info) {
        console.log("  -> SKIP: null info");
        return;
      }
      
      console.log(`  -> baseId: ${info.baseId}`);
      
      if (imageMap.has(info.baseId)) {
        console.log("  -> SKIP: already in map");
        return;
      }
      
      console.log("  -> ADDED to map");
      imageMap.set(info.baseId, info.url);
    });

    const uniqueImages = Array.from(imageMap.values());
    console.log("\nFinal unique images count:", uniqueImages.length);
    console.log("Final baseIds:", Array.from(imageMap.keys()));
    console.log("=== END DEBUG ===\n");

    const bulletSet = new Set<string>();
    const bulletSelectors = [
      '#feature-bullets ul li:not(:has(ul))', // Only leaf li
      '#featurebullets_feature_div ul li:not(:has(ul))',
      '#feature-bullets-content li:not(:has(ul))',
      '[data-feature-name="product-facts"] .a-list-item',
      '.product-facts-title + .a-unordered-list li:not(:has(ul))',
      '#product-facts-grid li:not(:has(ul))',
      '#productFactsDesktopExpander .a-list-item'
    ];

    $(bulletSelectors.join(', ')).each((_, el) => {
      const $el = $(el);
      
      // 1. Stricter container exclusion to avoid reviews, ads, and legal sections
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
        '.a-expander-content.a-expander-partial-collapse-content'
      ];
      if ($el.closest(junkContainers.join(', ')).length > 0) {
        return;
      }

      // 2. Clone and remove scripts/styles
      const $clone = $el.clone();
      $clone.find('script, style, .a-declarative, .a-popover-preload').remove();
      
      let text = '';
      // Try to get the text from the specific span first, then the element itself
      const $span = $clone.find('span.a-list-item');
      if ($span.length > 0) {
        text = $span.first().text().trim();
      } else {
        text = $clone.text().trim();
      }
      
      // Remove leading bullets or markers if any
      text = text.replace(/^[•\-\*\s]+/, '').trim();

      const isJunk = (t: string) => {
        const lower = t.toLowerCase();
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
          /^\d+ ratings?$/.test(t)
        );
      };

      if (text && !isJunk(text)) {
        bulletSet.add(text);
      }
    });

    const amazonBullets = Array.from(bulletSet);

    const variationsSet = new Set<string>();
    // Detect variations from standard twister elements
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
      // Only count if it looks like a real choice
      const text = $el.text().trim();
      const asin = $el.attr('data-asin');
      
      if (asin) {
        variationsSet.add(asin);
      } else if (text && text.length > 0 && !text.toLowerCase().includes('select')) {
        variationsSet.add(text);
      }
    });

    // Check for "twister-plus" style using data attributes
    $('[data-totalvariationcount]').each((_, el) => {
      const countAttr = $(el).attr('data-totalvariationcount');
      if (countAttr) {
        const count = parseInt(countAttr);
        if (!isNaN(count) && count > 1) {
          // If a row says it has 40 variations, we definitely have variations
          // Add dummy entries to satisfy the check (variationsCount > 1)
          for (let i = 0; i < Math.min(count, 10); i++) {
            variationsSet.add(`TwisterPlus-Var-${i}`);
          }
        }
      }
    });

    // Fallback: Check if there are multiple visual rows that usually indicate variations
    if (variationsSet.size <= 1) {
      const rows = $('#twister .a-row.variation-row, #twister .twister-selection-column, [id^="inline-twister-row-"]');
      rows.each((_, row) => {
        const $row = $(row);
        // Only count if the row has multiple actual options
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

    const variationsCount = variationsSet.size;

    // Calculate shipping days difference
    let shippingDays = "N/A";
    try {
      if (rawShippingTime) {
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const raw = rawShippingTime.toLowerCase();

        // ── Shortcut: same-day signals ───────────────────────────────────────────
        if (raw.includes('today') || raw.includes('vandaag') || raw.includes('aujourd') ||
            raw.includes('oggi') || raw.includes('heute')) {
          shippingDays = "0";
        }
        // ── Shortcut: next-day signals ────────────────────────────────────────────
        else if (raw.includes('tomorrow') || raw.includes('morgen') || raw.includes('demain') ||
                 raw.includes('domani') || raw.includes('jutro') || raw.includes('mañana') ||
                 raw.includes('imorgon')) {
          shippingDays = "1";
        }
        // ── Shortcut: day-after-tomorrow ─────────────────────────────────────────
        else if (raw.includes('overmorgen') || raw.includes('après-demain') || raw.includes('dopodomani')) {
          shippingDays = "2";
        }
        else {
          // ── Date-based calculation ─────────────────
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
              // English
              'jan': 0, 'feb': 1, 'mar': 2, 'apr': 3, 'may': 4, 'jun': 5, 'jul': 6, 'aug': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dec': 11,
              'january': 0, 'february': 1, 'march': 2, 'april': 3, 'june': 5, 'july': 6, 'august': 7, 'september': 8, 'october': 9, 'november': 10, 'december': 11,
              // German
              'januar': 0, 'februar': 1, 'märz': 2, 'mär': 2, 'juni': 5, 'juli': 6, 'oktober': 9, 'okt': 9, 'dezember': 11, 'dez': 11,
              // French
              'janvier': 0, 'janv': 0, 'février': 1, 'févr': 1, 'mars': 2, 'avril': 3, 'avr': 3, 'juin': 5, 'juillet': 6, 'juil': 6, 'août': 7, 'aoû': 7, 'septembre': 8, 'octobre': 9, 'novembre': 10, 'décembre': 11, 'déc': 11,
              // Spanish
              'enero': 0, 'ene': 0, 'febrero': 1, 'marzo': 2, 'abril': 3, 'abr': 3, 'mayo': 4, 'junio': 5, 'julio': 6, 'agosto': 7, 'ago': 7, 'septiembre': 8, 'octubre': 9, 'noviembre': 10, 'diciembre': 11, 'dic': 11,
              // Italian
              'gennaio': 0, 'gen': 0, 'febbraio': 1, 'aprile': 3, 'maggio': 4, 'mag': 4, 'giugno': 5, 'giu': 5, 'luglio': 6, 'lug': 6, 'settembre': 8, 'set': 8, 'ottobre': 9, 'ott': 9, 'dicembre': 11,
              // Polish
              'styczeń': 0, 'stycznia': 0, 'sty': 0, 'luty': 1, 'lutego': 1, 'lut': 1, 'marzec': 2, 'marca': 2, 'kwiecień': 3, 'kwietnia': 3, 'kwi': 3, 'maja': 4, 'maj': 4, 'mai': 4, 'czerwiec': 5, 'czerwca': 5, 'cze': 5, 'lipiec': 6, 'lipca': 6, 'sierpień': 7, 'sierpnia': 7, 'sie': 7, 'wrzesień': 8, 'września': 8, 'wrz': 8, 'październik': 9, 'października': 9, 'paź': 9, 'listopad': 10, 'listopada': 10, 'lis': 10, 'grudzień': 11, 'grudnia': 11, 'gru': 11,
              // Swedish
              'januari': 0, 'augusti': 7,
              // Dutch
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

          const match1 = rawShippingTime.match(new RegExp(`(\\d{1,2})(?:\\.?\\s*(?:de|di|d')?\\s*)(${monthRegexPattern})`, 'i'));
          if (match1) {
            day = parseInt(match1[1]);
            monthStr = match1[2];
          } else {
            const match2 = rawShippingTime.match(new RegExp(`(${monthRegexPattern})(?:\\s*(?:de|di)?\\s*)(\\d{1,2})`, 'i'));
            if (match2) {
              day = parseInt(match2[2]);
              monthStr = match2[1];
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
    } catch (e: any) {
      console.warn("Shipping days calculation failed:", e.message);
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

    const auditResult = await performAudit(masterData, liveData, 'amazon', domain);
    res.json({ liveData, auditResult });

  } catch (error: any) {
    console.error("Amazon Audit Error:", error);
    res.status(500).json({ error: error.message });
  } finally {
    if (browser) await browser.close();
  }
});

// Helper for performAudit
async function performAudit(master: any, live: any, mode: string, domain?: string) {
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

  result.description.similarity = getSimilarity(master.description || "", live.description || "");

  if (result.title.similarity > 0.8) result.title.match = true;
  if (result.description.similarity > 0.6 || live.hasAPlus) result.description.match = true;
  
  // Bullets match
  if (master.bullets && Array.isArray(master.bullets)) {
    master.bullets.forEach((mb: string) => {
      let bestSim = 0;
      let bestLive = "";
      if (live.bullets && Array.isArray(live.bullets)) {
        live.bullets.forEach((lb: string) => {
          const sim = getSimilarity(mb, lb);
          if (sim > bestSim) {
            bestSim = sim;
            bestLive = lb;
          }
        });
      }
      result.bullets.push({ master: mb, live: bestLive, similarity: bestSim, match: bestSim > 0.7 });
    });
  }

  // Price match (fuzzy)
  const masterPriceNum = parseFloat(String(master.price || "").replace(/[^0-9.]/g, '')) || 0;
  const livePriceNum = parseFloat(String(live.price || "").replace(/[^0-9.]/g, '')) || 0;
  if (masterPriceNum > 0 && Math.abs(masterPriceNum - livePriceNum) < 1.0) result.price.match = true;

  if (live.images && live.images.length >= (master.images?.length || 1)) result.images.match = true;

  // Score calculation
  let scoreValue = 0;
  if (mode === 'amazon') {
    if (result.title.match) scoreValue += 30;
    if (result.description.match) scoreValue += 30;
    const bulletMatchCount = (result.bullets || []).filter((b: any) => b.match).length;
    scoreValue += Math.min(bulletMatchCount * 8, 40);
  } else {
    if (result.title.match) scoreValue += 50;
    if (result.description.match) scoreValue += 50;
  }
  result.score = scoreValue;

  return result;
}

function getScoreGrade(score: number): string {
  if (score > 70) return "excellent";
  if (score >= 50) return "acceptable";
  return "Needs improvement";
}

// --- Bol.com Helpers ---

// ── BOL STRATEGY 2: Gemini Google Search Grounding ───────────────────────────
// Gemini's googleSearch tool performs web search queries via Google's engine.
// Since Google's crawl engines are not blocked by Akamai, this completely bypasses WAF limitations.
// Requires GEMINI_API_KEY env var (already used by the app).
async function tryBolViaGemini(ean: string): Promise<any | null> {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.log('[BOL GEMINI] No GEMINI_API_KEY found, skipping.');
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

    console.log('[BOL GEMINI] Generating content with googleSearch grounding for EAN:', ean);

    const prompt = `Perform a google search for "bol.com product ${ean}" or search for "${ean}" directly on bol.com.
Locate the official product page on bol.com.
Extract and return a single, exact JSON object with the following schema:
{
  "title": "exact full product title on bol.com",
  "price": "correct numerical price string e.g. 14.99",
  "shipping": "correct shipping time/delivery message e.g. 'Morgen in huis' or 'Uiterlijk donderdag 22 mei'",
  "description": "product description details, first 500 characters",
  "images": ["image url 1", "image url 2"],
  "bullets": ["feature point 1", "feature point 2"],
  "productUrl": "the direct final product link on bol.com",
  "liveVariations": "variation options if any, else empty string"
}
Make sure all details (pricing, title, shipping) are fully grounded in search results. Ensure the return contains ONLY the raw JSON object. No conversational helper text, no markdown other than \`\`\`json.`;

    const response = await genai.models.generateContent({
      model: 'gemini-3.5-flash',
      contents: prompt,
      config: {
        tools: [{ googleSearch: {} }],
        temperature: 0.1,
      }
    });

    const rawText = response.text?.trim() || '';
    console.log('[BOL GEMINI] Raw response length:', rawText.length);

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
      console.log('[BOL GEMINI] JSON parse failed:', parseErr);
    }
  } catch (e: any) {
    console.log('[BOL GEMINI] Strategy failed:', e.message);
  }

  return null;
}

// ── BOL STRATEGY 3: Playwright stealth browser (hardened backup) ─────────────
async function goToProduct(page: any, searchTerm: string) {
  const searchUrl = `https://www.bol.com/nl/nl/s/?searchtext=${encodeURIComponent(searchTerm)}`;
  console.log(`[BOL BROWSER] Searching for: ${searchTerm}`);

  const isAkamaiChallenge = (c: string, t: string) => {
    const isBolTitle = t.toLowerCase() === 'bol' || t.toLowerCase() === 'bol.com';
    if (
      c.includes('js-accept-all-cookies') ||
      c.includes('consent-assign-all') ||
      c.includes('search-input') ||
      c.includes('lang="nl-NL"')
    ) return false;
    return (
      c.includes('sec-if-cpt-container') ||
      c.includes('Toegang tot deze pagina is geweigerd') ||
      c.includes('Access Denied') ||
      c.includes('Pardon Our Interruption') ||
      (isBolTitle && c.includes('<meta name="Pragma" content="no-cache">')) ||
      (isBolTitle && !c.includes('lang="nl-NL"'))
    );
  };

  const isHardBlocked = (c: string) => {
    const cLower = c.toLowerCase();
    return (
      (cLower.includes('ip adres') && cLower.includes('geblokkeerd')) ||
      cLower.includes('rustig aan speed racer') ||
      cLower.includes('human verification')
    );
  };

  const waitForAkamai = async () => {
    let content = await page.content().catch(() => '');
    let title = await page.title().catch(() => '');
    if (isHardBlocked(content)) {
      throw new Error('WAF_BLOCKED: Bol.com blocked the request. IP address is blocked by their anti-bot system.');
    }
    if (!isAkamaiChallenge(content, title)) return true;
    console.log('[BOL BROWSER] Akamai WAF challenge detected — waiting for auto-resolve...');
    try {
      await page.mouse.move(Math.random() * 400 + 100, Math.random() * 300 + 100);
      await page.waitForTimeout(300);
      await page.mouse.move(Math.random() * 600 + 200, Math.random() * 400 + 150);
      await page.mouse.wheel(0, Math.random() * 200 + 100);
      await page.waitForTimeout(300);
    } catch (_) {}
    try {
      await page.waitForFunction(() => {
        const t = document.title.toLowerCase();
        if (t !== 'bol' && t !== 'bol.com' && t !== '') return true;
        if (document.documentElement.outerHTML.includes('lang="nl-NL"')) return true;
        if (document.body && document.body.innerHTML.length > 20000) return true;
        return false;
      }, { timeout: 6_000, polling: 500 });
    } catch (_) {
      console.log('[BOL BROWSER] Akamai wait finished.');
    }
    content = await page.content().catch(() => '');
    title = await page.title().catch(() => '');
    if (isHardBlocked(content)) {
      throw new Error('WAF_BLOCKED: Bol.com blocked the request. IP address is blocked by their anti-bot system.');
    }
    if (isAkamaiChallenge(content, title)) {
      const snippet = content.replace(/\s+/g, ' ').substring(0, 150);
      throw new Error(`WAF_BLOCKED: Stuck on Akamai challenge. Snippet: ${snippet}`);
    }
    return true;
  };

  const handleCookieConsent = async () => {
    try {
      const consentSelectors = [
        'button#js-accept-all-cookies',
        '[data-test="consent-assign-all"]',
        '#onetrust-accept-btn-handler',
        'button[class*="accept"]',
        'button[id*="accept"]',
      ];
      let clicked = false;
      for (const sel of consentSelectors) {
        const btn = await page.$(sel).catch(() => null);
        if (btn && await btn.isVisible().catch(() => false)) {
          console.log(`[BOL BROWSER] Clicking consent: ${sel}`);
          await Promise.all([
            page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 5_000 }).catch(() => null),
            btn.click().catch(() => null)
          ]);
          clicked = true;
          break;
        }
      }
      if (!clicked) {
        const jsClicked = await page.evaluate(() => {
          const btns = [
            document.querySelector('button#js-accept-all-cookies'),
            document.querySelector('[data-test="consent-assign-all"]'),
            document.querySelector('button[data-test="consent-modal-accept"]')
          ];
          for (const b of btns) { if (b) { (b as HTMLElement).click(); return true; } }
          const buttons = Array.from(document.querySelectorAll('button, [role="button"]'));
          const target = buttons.find(b => {
            const t = (b.textContent || '').toLowerCase();
            return t.includes('akkoord') || t.includes('accepteer') || t.includes('accept') || t.includes('alle cookies');
          });
          if (target) { (target as HTMLElement).click(); return true; }
          return false;
        }).catch(() => false);
        if (jsClicked) {
          await page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 5_000 }).catch(() => null);
        }
      }
      await page.evaluate(() => {
        const overlay = document.querySelector('.consent-modal, #consent-modal, .cookie-consent');
        if (overlay) (overlay as HTMLElement).style.display = 'none';
        document.body.style.overflow = 'auto';
      }).catch(() => null);
      await page.waitForTimeout(800);
    } catch (_) {}
  };

  console.log('[BOL BROWSER] Step 1: Visiting homepage to warm Akamai session...');
  try {
    await page.goto('https://www.bol.com/nl/nl/', {
      waitUntil: 'domcontentloaded',
      timeout: 20_000
    });
    await page.waitForTimeout(Math.random() * 1200 + 600);
    await page.mouse.move(Math.random() * 400 + 200, Math.random() * 200 + 100);
    await page.mouse.wheel(0, Math.random() * 150 + 50);
    await page.waitForTimeout(Math.random() * 800 + 300);
  } catch (e) {
    console.log('[BOL BROWSER] Homepage warmup failed (non-fatal):', (e as Error).message);
  }

  await handleCookieConsent();

  {
    const hpContent = await page.content().catch(() => '');
    const hpTitle = await page.title().catch(() => '');
    if (isHardBlocked(hpContent)) {
      throw new Error('WAF_BLOCKED: Bol.com blocked the request. IP address is blocked by their anti-bot system.');
    }
    if (isAkamaiChallenge(hpContent, hpTitle)) {
      await waitForAkamai();
    }
  }

  console.log('[BOL BROWSER] Step 2: Navigating to search URL...');
  try {
    await page.goto(searchUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 });
  } catch (_) {}

  await handleCookieConsent();
  await page.waitForTimeout(1_200);

  try {
    await page.waitForFunction(() => {
      return (
        window.location.href.includes('/p/') ||
        !!document.querySelector('[data-test="product-item"]') ||
        !!document.querySelector('.product-item--row') ||
        !!document.querySelector('.ui-kit-card') ||
        document.body.innerText.toLowerCase().includes('0 resultaten') ||
        document.body.innerText.toLowerCase().includes('geen resultaten')
      );
    }, { timeout: 10_000, polling: 500 });
  } catch (_) {
    console.log('[BOL BROWSER] Timeout waiting for search results.');
  }

  await waitForAkamai();

  const titleCheck = await page.title().catch(() => '');
  if (titleCheck.toLowerCase() === 'bol' || titleCheck.toLowerCase() === 'bol.com') {
    await handleCookieConsent();
    await waitForAkamai();
  }

  const content = await page.content().catch(() => '');

  if (isHardBlocked(content)) {
    throw new Error('WAF_BLOCKED: Bol.com blocked the request. IP address is blocked by their anti-bot system.');
  }

  if (content.includes('geen resultaten gevonden') || content.includes('0 resultaten')) {
    throw new Error(`NO_RESULTS: Bol.com found no results for "${searchTerm}". Please verify the EAN/Search term.`);
  }

  if (!page.url().includes('/p/')) {
    const resultSelectors = [
      '[data-test="product-item"]',
      '.product-item--row',
      '.product-list',
      'li[class*="product"]',
      'div[class*="product-item"]',
      'a[href*="/p/"]',
      '.ui-kit-card'
    ];
    let foundResults = false;
    for (const sel of resultSelectors) {
      try {
        await page.waitForSelector(sel, { state: 'attached', timeout: 5_000 });
        foundResults = true;
        console.log(`[BOL BROWSER] Search results found: ${sel}`);
        break;
      } catch (_) {}
    }
    if (!foundResults) {
      await page.waitForTimeout(1_000);
      await page.mouse.wheel(0, 500);
      await page.waitForTimeout(1_000);
    }

    const productHref = await page.evaluate(() => {
      const titleLinkSelectors = [
        'a[data-test="product-title"]',
        'a.product-title',
        'a.product-item__title',
        '[data-test="product-item"] a[href*="/p/"]',
        '.product-item--row a[href*="/p/"]',
        'a[href*="/p/"]'
      ];
      for (const sel of titleLinkSelectors) {
        const el = document.querySelector(sel) as HTMLAnchorElement;
        if (el && el.href && el.href.includes('/p/') && !el.href.includes('/m/')) return el.href;
      }
      const allLinks = Array.from(document.querySelectorAll('a'));
      const productLink = allLinks.find(a => {
        const href = a.href;
        return href && href.includes('/p/') && !href.includes('/m/') && !href.includes('/s/') &&
               !href.includes('/c/') && !href.includes('/l/') && !href.includes('#');
      });
      return productLink ? productLink.href : null;
    }).catch(() => null);

    if (productHref) {
      console.log(`[BOL BROWSER] Navigating to product: ${productHref}`);
      const fullUrl = productHref.startsWith('http') ? productHref : 'https://www.bol.com' + productHref;
      await page.waitForTimeout(Math.random() * 600 + 300);
      await page.goto(fullUrl, { waitUntil: 'domcontentloaded', timeout: 20_000 }).catch(() => null);
      await page.waitForTimeout(1_000);
    } else {
      const debugTitle = await page.title().catch(() => '');
      const debugUrl = page.url();
      const debugContent = await page.content().catch(() => '');
      const snippet = debugContent.replace(/\s+/g, ' ').substring(0, 250);
      throw new Error(`No product link found on Bol.com. Title: "${debugTitle}", Snippet: ${snippet}`);
    }
  }
}

async function extractCatalogue(page: any) {
  let initialContent = await page.content().catch(() => '');
  let initialContentLower = initialContent.toLowerCase();
  if ((initialContentLower.includes('ip adres') && initialContentLower.includes('geblokkeerd')) || 
      initialContentLower.includes('rustig aan speed racer') ||
      initialContentLower.includes('human verification')) {
    throw new Error('WAF_BLOCKED: Bol.com blocked the request. IP address is blocked by their anti-bot system.');
  }

  await page.waitForLoadState('domcontentloaded', { timeout: 15_000 }).catch(() => null);
  await page.waitForTimeout(1_200);

  await page.mouse.wheel(0, 400);
  await page.waitForTimeout(1_200);

  return await page.evaluate(() => {
    let title = '';
    const el = document.querySelector('[data-test="title"]') || document.querySelector('h1.page-title') || document.querySelector('h1');
    title = el ? (el as HTMLElement).innerText.trim() : '';
    if (!title || title.length < 5) {
      title = document.title.split('|')[0].trim();
    }

    let description = '';
    const heading = Array.from(
      document.querySelectorAll('h2,h3,h4,b,strong,span')
    ).find(h =>
      (h.textContent ?? '').toLowerCase().includes('productbeschrijving') ||
      (h.textContent ?? '').toLowerCase().includes('product description')
    );

    if (heading) {
      const parent = heading.closest('section') ?? heading.parentElement;
      if (parent) {
        const clone = parent.cloneNode(true) as HTMLElement;
        clone.querySelectorAll('.js_description_read_more, [data-test="read-more"], .pdp-description__read-more, button, a.button--link')
          .forEach(el => el.remove());
        description = (clone.innerText ?? '')
          .replace(/Productbeschrijving|Product description/i, '')
          .trim()
          .replace(/toon meer|toon minder/gi, '')
          .trim();
      }
    }

    if (!description || description.length < 50) {
      const selectors = [
        '[data-test="description"]',
        '[data-test="product-description"]',
        '.js_product_description',
        '.product-description',
        '.product-description-content',
        'div[itemprop="description"]',
        '#descriptionBlock',
        'section#description',
        '.slot-product-description',
        '.pdp-description',
        '.manufacturer-info',
        '.product-info',
        '[data-test="product-info"]'
      ];
      const readMore = document.querySelector('.js_description_read_more, [data-test="read-more"], .pdp-description__read-more');
      if (readMore) (readMore as HTMLElement).click();

      const parts: string[] = [];
      selectors.forEach(sel => {
        const el = document.querySelector(sel);
        if (el) {
          const clone = el.cloneNode(true) as HTMLElement;
          clone.querySelectorAll('.js_description_read_more, [data-test="read-more"], .pdp-description__read-more, button, a.button--link')
            .forEach(b => b.remove());
          let txt = (clone.innerText ?? '').trim();
          if (txt.length > 20) {
            txt = txt
              .replace(/<br\s*\/?>/gi, '\n')
              .replace(/<\/?p>/gi, '\n')
              .replace(/<\/?div>/gi, '\n')
              .replace(/<\/?[^>]+>/g, ' ')
              .replace(/\s+/g, ' ')
              .trim()
              .replace(/toon meer|toon minder/gi, '')
              .trim();
            parts.push(txt);
          }
        }
      });
      if (parts.length) description = parts.join('\n\n');
    }

    // Price extraction
    let price = 'N/A';
    const pageHtml = document.documentElement.innerHTML;
    
    // 1. Try to find price in JSON-LD (Search for the "offers" block mentioned by user)
    const jsonLdMatch = pageHtml.match(/"offers"\s*:\s*\{[^}]*"@type"\s*:\s*"Offer"[^}]*"price"\s*:\s*"([\d.]+)"/) ||
                        pageHtml.match(/"offers"\s*:\s*\{[^}]*"@type"\s*:\s*"Offer"[^}]*"price"\s*:\s*([\d.]+)/);
    
    if (jsonLdMatch) {
      price = jsonLdMatch[1];
    } else {
      // Fallback: Other meta patterns if the direct "offers" block isn't exactly as expected
      const metaPriceMatch = pageHtml.match(/"price"\s*:\s*"([\d.]+)"\s*,\s*"priceCurrency"\s*:\s*"EUR"/) ||
                             pageHtml.match(/"price"\s*:\s*([\d.]+)\s*,\s*"priceCurrency"\s*:\s*"EUR"/);
      if (metaPriceMatch) price = metaPriceMatch[1];
    }

    // 2. Last resort: text-based extraction from the UI (Normal mentioned lines)
    if (price === 'N/A') {
      const allText = document.body.innerText;
      // Look for € symbol followed by digits, handling European comma decimals
      const euroMatch = allText.match(/€\s*([\d.]+,\d{2})/) || allText.match(/€\s*([\d.]+)/);
      if (euroMatch) {
        price = euroMatch[1].replace(',', '.').trim();
      }
    }

    let shipping = 'N/A';
    const shipMatch = pageHtml.match(/"deliveryDescription"\s*,\s*"([^"]+)"/);
    if (shipMatch) {
      shipping = shipMatch[1];
    } else {
      const shipSel = [
        '[data-test="delivery-message"]',
        '[data-test="delivery"]',
        'span[class*="delivery"]',
        'div[class*="shipping"]',
        '[class*="DeliveryInformation"]',
        'span[class*="Delivery"]',
        '.delivery-text',
        '[data-element-type="delivery"]',
        'span[itemprop="deliveryTime"]'
      ];
      for (const sel of shipSel) {
        const el = document.querySelector(sel);
        if (el) {
          const txt = (el as HTMLElement).innerText ?? (el as HTMLElement).textContent;
          if (txt && txt.trim().length) {
            shipping = txt.trim();
            break;
          }
        }
      }
      if (shipping === 'N/A') {
        const body = document.body.innerText;
        const m = body.match(/Uiterlijk\s+(.+?)(?:\s+in\s+huis|$)/i) ||
                  body.match(/Morgen\s+in\s+huis/i) ||
                  body.match(/Vandaag\s+.*?(?:in|om)/i) ||
                  body.match(/Bezorging:\s+(.+?)(?:\n|$)/i);
        if (m) shipping = m[0] ?? m[1] ?? 'N/A';
      }
    }

    const imgs: string[] = [];
    
    // 1. Always prioritize the main image
    const mainSel = [
      '[data-test="product-main-image"] img',
      '.js_main_product_image',
      '.pdp-main-image img',
      'img.js_main_product_image',
      '[data-test="pdp-main-image"] img'
    ];
    mainSel.forEach(sel => {
      const img = document.querySelector(sel) as HTMLImageElement | null;
      if (img && img.src && img.src.startsWith('http')) {
        imgs.push(img.src);
      } else if (img && img.getAttribute('data-src')) {
        const dsrc = img.getAttribute('data-src');
        if (dsrc && dsrc.startsWith('http')) imgs.push(dsrc);
      }
    });

    // 2. Extract thumbnails and other product images
    const allImages = Array.from(document.querySelectorAll('img'));
    allImages.forEach(img => {
      const alt = img.getAttribute('alt') || '';
      const src = img.src || img.getAttribute('data-src') || '';
      
      if (alt.includes('Afbeelding nummer')) {
        if (src && src.startsWith('http')) imgs.push(src);
      }
    });

    // 3. Fallback for thumbnails if the "Afbeelding nummer" pattern is missing
    const thumbSel = [
      '.js_product_media_items img',
      '.pdp-images img',
      '.js_image_container img',
      '.product-images__item img',
      '[data-test="pdp-thumbnails"] img'
    ];
    thumbSel.forEach(sel => {
      const thumbs = Array.from(document.querySelectorAll(sel)) as HTMLImageElement[];
      thumbs.forEach(i => {
        const src = i.src || i.getAttribute('data-src') || i.getAttribute('src');
        if (src && (src.includes('media.s-bol.com') || src.startsWith('http'))) {
          imgs.push(src);
        }
      });
    });

    const bulletSel = [
      '[data-test="product-features"] li',
      '.product-features li',
      '.js_product_features li',
      '.specs-list__item',
      '.product-specifications li'
    ];
    const bulletSet = new Set<string>();
    bulletSel.forEach(sel => {
      document.querySelectorAll(sel).forEach(li => {
        const txt = (li as HTMLElement).innerText.trim();
        if (txt.length > 3) bulletSet.add(txt);
      });
    });

    // Variations extraction
    let variationsData = '';
    const varItems = Array.from(document.querySelectorAll('div, label, a, span, button')).filter(el => {
      const t = (el.textContent || '').toLowerCase();
      return t.includes('kies je ');
    });
    
    if (varItems.length > 0) {
       // Look for the closest container and extract its text
       const container = varItems[0].closest('section, div.variant-container, div.product-variants, [data-test="variants"], [class*="variant"]') || varItems[0].parentElement?.parentElement;
       if (container) {
         variationsData = (container as HTMLElement).innerText.trim().replace(/\s+/g, ' ');
       } else {
         variationsData = "Various options found: " + varItems.map(v => v.textContent?.replace(/\s+/g, ' ').trim()).join(' | ');
       }
    } else {
       // Fallback checking for family properties in pageHtml
       const varMatch = pageHtml.match(/"productFamily"\s*:\s*\{"products"\s*:\s*\[(.*?)\]\s*\}/);
       if (varMatch) {
         try {
           const parsed = JSON.parse(`[${varMatch[1]}]`);
           variationsData = parsed.map((v: any) => `${v.title || v.name}`).join(' | ');
         } catch(e) {}
       }
    }

    return {
      title,
      description,
      price,
      shipping,
      images: Array.from(new Set(imgs)),
      bullets: Array.from(bulletSet),
      liveVariations: variationsData
    };
  });
}

function calculateBolShippingDays(rawShippingTime: string): string {
  if (!rawShippingTime) return "N/A";
  
  const text = rawShippingTime.toLowerCase();
  
  if (text.includes("vandaag")) return "0";
  if (text.includes("morgen")) return "1";
  if (text.includes("overmorgen")) return "2";
  
  const months = ['januari', 'februari', 'maart', 'april', 'mei', 'juni', 'juli', 'augustus', 'september', 'oktober', 'november', 'december'];
  const monthRegexText = months.join('|');
  const dateRegex = new RegExp(`(\\d{1,2})\\s*(${monthRegexText})`, 'i');
  
  const match = text.match(dateRegex);
  if (match) {
    const day = parseInt(match[1]);
    const month = months.indexOf(match[2].toLowerCase());
    if (month !== -1) {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      
      let targetDate = new Date(today.getFullYear(), month, day);
      if (targetDate.getTime() < today.getTime() - 1000 * 60 * 60 * 24 * 30) {
        targetDate = new Date(today.getFullYear() + 1, month, day);
      }
      
      const diffTime = targetDate.getTime() - today.getTime();
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      return diffDays >= 0 ? diffDays.toString() : "N/A";
    }
  }
  
  const days = ['zondag', 'maandag', 'dinsdag', 'woensdag', 'donderdag', 'vrijdag', 'zaterdag'];
  const dayRegex = new RegExp(`(${days.join('|')})`, 'i');
  const dayMatchText = text.match(dayRegex);
  if (dayMatchText) {
    const targetDay = days.indexOf(dayMatchText[1].toLowerCase());
    const today = new Date();
    today.setHours(0,0,0,0);
    const currentDay = today.getDay();
    let diffDays = targetDay - currentDay;
    if (diffDays <= 0) diffDays += 7;
    return diffDays.toString();
  }

  return "N/A";
}

// 3. Audit Bol.com
app.post("/api/audit/bol", async (req, res) => {
  let browser;
  try {
    const { ean, masterData } = req.body;
    if (!ean) throw new Error('Missing "ean" in request body');

    let data: any = null;
    let dataSource = 'browser';

    // ── Strategy 1: Gemini Google Search Grounding ─────────────────────────
    console.log('[BOL] Trying Strategy 1: Gemini Google Search Grounding...');
    data = await tryBolViaGemini(ean);
    if (data) {
      console.log('[BOL] Strategy 1 succeeded via Gemini.');
      dataSource = 'gemini';
    }

    // ── Strategy 2: Playwright stealth browser (hardened) ─────────────────────
    // Add delay before browser strategy to reduce rapid-fire requests
    if (!data) {
      console.log('[BOL] Adding 2-second delay before browser strategy to avoid WAF detection...');
      await new Promise(resolve => setTimeout(resolve, 2000));
      console.log('[BOL] Trying Strategy 2: Playwright stealth browser...');

      const launchOpts: any = {
        headless: false,
        args: [
          '--headless=new',
          '--no-sandbox',
          '--disable-setuid-sandbox',
          '--disable-blink-features=AutomationControlled',
          '--disable-dev-shm-usage',
          '--disable-gpu',
          '--window-size=1920,1080',
          '--disable-features=IsolateOrigins,site-per-process',
          '--disable-infobars',
          '--disable-extensions',
          '--disable-default-apps',
          '--no-first-run',
          '--no-default-browser-check',
          '--password-store=basic',
          '--use-mock-keychain',
          '--incognito'
        ]
      };

      const proxyServer = process.env.PROXY_SERVER;
      if (proxyServer) {
        launchOpts.proxy = {
          server: proxyServer,
          username: process.env.PROXY_USERNAME,
          password: process.env.PROXY_PASSWORD
        };
      }

      browser = await chromiumExtra.launch(launchOpts);

      const screenWidth = Math.floor(Math.random() * (1920 - 1366 + 1)) + 1366;
      const screenHeight = Math.floor(Math.random() * (1080 - 768 + 1)) + 768;

      const context = await browser.newContext({
        userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36',
        viewport: { width: screenWidth, height: screenHeight },
        screen: { width: screenWidth, height: screenHeight },
        locale: 'nl-NL',
        timezoneId: 'Europe/Amsterdam',
        colorScheme: 'light',
        extraHTTPHeaders: {
          'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
          'Accept-Language': 'nl-NL,nl;q=0.9,en-US;q=0.8,en;q=0.7',
          'Accept-Encoding': 'gzip, deflate, br, zstd',
          'sec-ch-ua': '"Chromium";v="136", "Google Chrome";v="136", "Not-A.Brand";v="99"',
          'sec-ch-ua-mobile': '?0',
          'sec-ch-ua-platform': '"Windows"',
          'sec-ch-ua-platform-version': '"15.0.0"',
          'sec-ch-ua-full-version-list': '"Chromium";v="136.0.7103.114", "Google Chrome";v="136.0.7103.114", "Not-A.Brand";v="99.0.0.0"',
          'upgrade-insecure-requests': '1',
          'sec-fetch-site': 'none',
          'sec-fetch-mode': 'navigate',
          'sec-fetch-user': '?1',
          'sec-fetch-dest': 'document',
          'cache-control': 'max-age=0',
          'DNT': '1'
        }
      });

      // Pre-inject OneTrust consent cookies to skip cookie consent redirect
      await context.addCookies([
        { name: 'consent_cookie', value: '1', domain: '.bol.com', path: '/', sameSite: 'Lax' },
        { name: 'accept_all_cookies', value: 'true', domain: '.bol.com', path: '/', sameSite: 'Lax' },
        {
          name: 'OptanonAlertBoxClosed',
          value: new Date().toISOString(),
          domain: '.bol.com',
          path: '/',
          sameSite: 'Lax'
        },
        {
          name: 'OptanonConsent',
          value: 'isIABGlobal=false&datestamp=' +
            encodeURIComponent(new Date().toUTCString()) +
            '&version=202209.1.0&hosts=&consentId=' +
            Math.random().toString(36).substring(2) +
            '&interactionCount=1&landingPath=NotLandingPage&groups=C0001%3A1%2CC0002%3A0%2CC0003%3A0%2CC0004%3A0&geolocation=NL%3BNH&AwaitingReconsent=false',
          domain: '.bol.com',
          path: '/',
          sameSite: 'Lax'
        }
      ]);

      const page = await context.newPage();

      await goToProduct(page, ean);
      data = await extractCatalogue(page);
      dataSource = 'browser';
    }

    const bolShippingDays = calculateBolShippingDays(data.shipping || '');

    const liveData = {
      title: data.title || '',
      price: data.price || 'N/A',
      description: data.description || '',
      images: data.images || [],
      url: data.productUrl || data.url || '',
      hasAPlus: false,
      shipping: bolShippingDays !== 'N/A' ? `${bolShippingDays} days` : (data.shipping || 'N/A'),
      shippingDays: bolShippingDays,
      rawShipping: data.shipping || '',
      variations: data.liveVariations && data.liveVariations.length > 5
        ? data.liveVariations.split('|').length || 1
        : 0,
      bullets: data.bullets || [],
      rawVariationsText: data.liveVariations || '',
      _dataSource: dataSource
    };

    const auditResult = await performAudit(masterData, liveData, 'bol');
    res.json({ liveData, auditResult });

  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (browser) await browser.close();
  }
});

// 4. Sheets APIs
app.post("/api/sheets/fetch", async (req, res) => {
  try {
    const { sheetId, mode, marketplace, sheetName: requestedSheetName } = req.body;
    const auth = new JWT({
      email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const doc = new GoogleSpreadsheet(sheetId, auth);
    await doc.loadInfo();

    const targetTabName = requestedSheetName || (mode === "bol" ? "Product data" : "Amazon Data");
    const sheet = doc.sheetsByTitle[targetTabName];
    if (!sheet) {
      throw new Error(`Sheet tab "${targetTabName}" not found in the spreadsheet.`);
    }

    // Amazon Data tab has a category group row in row 1;
    // the real column headers (ASIN, AMZ title DE, …) are in row 2.
    // Product data tab has normal headers in row 1.
    if (targetTabName === 'Amazon Data') {
      await sheet.loadHeaderRow(2);
    }

    const rows = await sheet.getRows();
    const data = rows.map(row => row.toObject());
    
    res.json({ data });
  } catch (error: any) {
    console.error("Fetch Sheet Error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/sheets/save-audit", async (req, res) => {
  try {
    const { sheetId, mode, identifier, auditResult, liveData, masterData } = req.body;
    
    if (!sheetId) throw new Error("Missing sheetId in request body");

    const auth = new JWT({
      email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const doc = new GoogleSpreadsheet(sheetId, auth);
    await doc.loadInfo();
    
    const targetSheetName = mode === "bol" ? "Bol QC Results" : "Amazon QC Results";
    // Case-insensitive search for the sheet
    const sheet = doc.sheetsByTitle[targetSheetName] || 
                  doc.sheetsByIndex.find(s => s.title.toLowerCase().trim() === targetSheetName.toLowerCase());
    
    if (!sheet) {
      const availableSheets = doc.sheetsByIndex.map(s => `"${s.title}"`).join(', ');
      throw new Error(`Sheet "${targetSheetName}" not found. Available sheets: ${availableSheets}`);
    }

    // Always override everything and start from row 2
    await sheet.clearRows();

    // Mapping according to instructions
    const bulletMatchCount = (auditResult && Array.isArray(auditResult.bullets)) ? auditResult.bullets.filter((b: any) => b.match).length : 0;
    const bulletMatchText = bulletMatchCount > 0 ? 'Yes' : 'No';

    const timestamp = new Date().toLocaleString();
    const resultRow: any = {
      'Date': timestamp,
      'date': timestamp,
      'EAN': identifier,
      'ean': identifier,
      'ASIN': identifier,
      'asin': identifier,
      'ASIN/EAN': identifier,
      'asin/ean': identifier,
      'Identifier': identifier,
      'identifier': identifier,
      'Mode': mode,
      'mode': mode,
      'Title Match': auditResult?.title?.match ? 'Yes' : 'No',
      'Description Match': liveData?.hasAPlus ? 'A+ content available' : (auditResult?.description?.match ? 'Yes' : 'No'),
      'Bullet Points Match': bulletMatchText,
      'Variation': auditResult?.variations?.match ? 'Yes' : 'No',
      'A+ Content': liveData?.hasAPlus ? 'Yes' : 'No',
      'Buybox Owner': liveData?.buyboxOwner || 'N/A',
      'Score': auditResult?.score ?? 0,
      'Score Grade': getScoreGrade(auditResult?.score ?? 0),
      'Price Live': liveData?.price || 'N/A',
      'price live': liveData?.price || 'N/A',
      'Price': liveData?.price || 'N/A',
      'price': liveData?.price || 'N/A',
      'Shipping Live (Days)': liveData?.shippingDays || 'N/A',
      'Shipping Time': liveData?.shippingDays || 'N/A',
      'shipping time': liveData?.shippingDays || 'N/A',
      'Notes': liveData?.hasAPlus ? 'A+ content available' : ''
    };

    const pfx = mode === 'bol' ? 'Bol' : 'Amazon';
    const shortPfx = mode === 'bol' ? 'BOL' : 'AMZ';

    // Images mapping - User wants "each relevant column based on number image"
    if (masterData && masterData.images && Array.isArray(masterData.images)) {
      masterData.images.slice(0, 10).forEach((url: string, i: number) => {
        const formula = url ? `=IMAGE("${url}")` : '';
        const index = i + 1;
        resultRow[`${pfx} Master Image ${index}`] = formula;
        resultRow[`${shortPfx} Master Image ${index}`] = formula;
        resultRow[`Master Image ${index}`] = formula;
        resultRow[`${shortPfx} IMG ${index}`] = formula;
        resultRow[`image ${mode} data ${index}`] = formula;
        resultRow[`Image ${pfx} Data ${index}`] = formula;
        
        resultRow[`Amazon Master Image ${index}`] = formula;
      });
    }

    if (liveData && liveData.images && Array.isArray(liveData.images)) {
      const sliceStart = mode === 'bol' ? 0 : 1;
      liveData.images.slice(sliceStart, sliceStart + 10).forEach((url: string, i: number) => {
        const formula = url ? `=IMAGE("${url}")` : '';
        const index = i + 1;
        resultRow[`${pfx} Live Image ${index}`] = formula;
        resultRow[`${shortPfx} Live Image ${index}`] = formula;
        resultRow[`Live Image ${index}`] = formula;
        resultRow[`${shortPfx} L IMG ${index}`] = formula;
        resultRow[`image ${mode} live ${index}`] = formula;
        resultRow[`Image ${pfx} Live ${index}`] = formula;
        
        resultRow[`Amazon Live Image ${index}`] = formula;
      });
    }

    // Search for existing row to override removed based on prompt instruction to always overwrite starting from row 2
    await sheet.loadHeaderRow();
    const headers = sheet.headerValues;
    const findHeader = (target: string) => {
      return headers.find(h => h.toLowerCase().trim() === target.toLowerCase().trim());
    };

    // Add new row - mapping to actual sheet headers
    const rowToSave: any = {};
    Object.keys(resultRow).forEach(key => {
      const matchingHeader = findHeader(key);
      if (matchingHeader) {
        rowToSave[matchingHeader] = resultRow[key];
      }
    });
    
    // If we couldn't find some headers in the rowToSave, but we have them in resultRow and they are standard, 
    // they might not be added if headers don't exist. But here we assume sheet has them.
    if (Object.keys(rowToSave).length > 0) {
      await sheet.addRow(rowToSave);
    } else {
      // Fallback to original object if no headers matched (unlikely but safe)
      await sheet.addRow(resultRow);
    }
    
    res.json({ success: true });
  } catch (error: any) {
    console.error("Save Audit Error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/sheets/batch-save-audit", async (req, res) => {
  try {
    const { sheetId, mode, audits } = req.body;
    if (!audits || !Array.isArray(audits) || audits.length === 0) {
      return res.json({ success: true, message: "No audits to save." });
    }
    
    if (!sheetId) throw new Error("Missing sheetId in request body");

    const auth = new JWT({
      email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const doc = new GoogleSpreadsheet(sheetId, auth);
    await doc.loadInfo();

    const targetSheetName = mode === "bol" ? "Bol QC Results" : "Amazon QC Results";
    const sheet = doc.sheetsByTitle[targetSheetName] || 
                  doc.sheetsByIndex.find(s => s.title.toLowerCase().trim() === targetSheetName.toLowerCase());

    if (!sheet) {
      throw new Error(`Sheet "${targetSheetName}" not found.`);
    }

    // Clear existing rows to start automatically from row 2
    await sheet.clearRows();
    await sheet.loadHeaderRow();
    
    const rowsToAdd = audits.map((audit: any) => {
      const { identifier, auditResult, liveData, masterData } = audit;
      const bulletMatchCount = (auditResult && Array.isArray(auditResult.bullets)) ? auditResult.bullets.filter((b: any) => b.match).length : 0;
      const bulletMatchText = bulletMatchCount > 0 ? 'Yes' : 'No';
      const timestamp = new Date().toLocaleString();

      const row: any = {
        'Date': timestamp,
        'date': timestamp,
        'EAN': identifier,
        'ean': identifier,
        'ASIN': identifier,
        'asin': identifier,
        'ASIN/EAN': identifier,
        'asin/ean': identifier,
        'Identifier': identifier,
        'identifier': identifier,
        'Mode': mode,
        'mode': mode,
        'Title Match': auditResult?.title?.match ? 'Yes' : 'No',
        'Description Match': liveData?.hasAPlus ? 'A+ content available' : (auditResult?.description?.match ? 'Yes' : 'No'),
        'Bullet Points Match': bulletMatchText,
        'Variation': auditResult?.variations?.match ? 'Yes' : 'No',
        'A+ Content': liveData?.hasAPlus ? 'Yes' : 'No',
        'Buybox Owner': liveData?.buyboxOwner || 'N/A',
        'Score': auditResult?.score ?? 0,
        'Score Grade': getScoreGrade(auditResult?.score ?? 0),
        'Price Live': liveData?.price || 'N/A',
        'price live': liveData?.price || 'N/A',
        'Price': liveData?.price || 'N/A',
        'price': liveData?.price || 'N/A',
        'Shipping Live (Days)': liveData?.shippingDays || 'N/A',
        'Shipping Time': liveData?.shippingDays || 'N/A',
        'shipping time': liveData?.shippingDays || 'N/A',
        'Notes': liveData?.hasAPlus ? 'A+ content available' : ''
      };

      const pfx = mode === 'bol' ? 'Bol' : 'Amazon';
      const shortPfx = mode === 'bol' ? 'BOL' : 'AMZ';

      if (masterData && Array.isArray(masterData.images)) {
        masterData.images.slice(0, 10).forEach((url: string, i: number) => {
          const formula = url ? `=IMAGE("${url}")` : '';
          row[`${pfx} Master Image ${i + 1}`] = formula;
          row[`${shortPfx} Master Image ${i + 1}`] = formula;
          row[`Master Image ${i + 1}`] = formula;
          row[`${shortPfx} IMG ${i + 1}`] = formula;
          row[`image ${mode} data ${i + 1}`] = formula;
          row[`Image ${pfx} Data ${i + 1}`] = formula;
          
          row[`Amazon Master Image ${i + 1}`] = formula; 
        });
      }

      if (liveData && Array.isArray(liveData.images)) {
        const sliceStart = mode === 'bol' ? 0 : 1;
        liveData.images.slice(sliceStart, sliceStart + 10).forEach((url: string, i: number) => {
          const formula = url ? `=IMAGE("${url}")` : '';
          row[`${pfx} Live Image ${i + 1}`] = formula;
          row[`${shortPfx} Live Image ${i + 1}`] = formula;
          row[`Live Image ${i + 1}`] = formula;
          row[`${shortPfx} L IMG ${i + 1}`] = formula;
          row[`image ${mode} live ${i + 1}`] = formula;
          row[`Image ${pfx} Live ${i + 1}`] = formula;
          
          row[`Amazon Live Image ${i + 1}`] = formula; 
        });
      }
      return row;
    });

    await sheet.addRows(rowsToAdd);
    res.json({ success: true, count: rowsToAdd.length });
  } catch (error: any) {
    console.error("Batch Save Audit Error:", error);
    res.status(500).json({ error: error.message });
  }
});

// Vite middleware for development
if (process.env.NODE_ENV !== "production") {
  const { createServer: createViteServer } = await import('vite');
  const vite = await createViteServer({
    server: { middlewareMode: true },
    appType: "spa",
  });
  app.use(vite.middlewares);
} else {
  const distPath = path.join(process.cwd(), 'dist');
  app.use(express.static(distPath));
  app.get('*', (req, res) => {
    res.sendFile(path.join(distPath, 'index.html'));
  });
}

const PORT = process.env.PORT || 3000;
app.listen(Number(PORT), "0.0.0.0", () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
