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
      userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
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
    
    await page.route('**/*.{png,jpg,jpeg,gif,webp,svg,woff,woff2,ttf,otf}', (route) => {
      route.abort();
    });

    try {
      console.log(`Auditing Amazon ${asin} on ${domain} (Target Zip: ${locConfig.zip})...`);
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
      
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
          return !slot.textContent?.includes(zip);
        }, { zip: locConfig.zip });

        if (isRegionalLocked) {
          console.log(`UI Regional Unlock Fallback: Injecting ${locConfig.zip} for ${domain}`);
          // Wait for first click option to be visible
          const locBtn = await page.waitForSelector('#nav-global-location-slot, #glow-ingress-block, #nav-main-ftr-location-slot', { visible: true, timeout: 15000 }).catch(() => null);
          if (locBtn) {
            await locBtn.click({ force: true });
            await page.waitForTimeout(2000);

            // 1. Check if Country Selection Dropdown is shown instead of Zip update input (e.g. from international proxy IP)
            const countryListSelector = '#GLUXCountryList';
            const countryListVisible = await page.locator(countryListSelector).isVisible().catch(() => false);
            if (countryListVisible && locConfig.countryCode) {
              console.log(`Country list dropdown detected inside Amazon location popover. Selecting: ${locConfig.countryCode}`);
              await page.selectOption(countryListSelector, { value: locConfig.countryCode }).catch(() => null);
              await page.waitForTimeout(1000);

              // Click "Go" or "Apply" button for country selection
              const goBtnSelector = 'input[aria-labelledby="GLUXCountryList-announce"], button[name="glowDoneButton"], #GLUXCountryList-announce + input, .a-popover-footer input';
              await page.click(goBtnSelector, { force: true }).catch(() => null);
              await page.waitForTimeout(3000);

              // Reload or re-click to get zip code inputs visible!
              console.log('Re-clicking global location slot after selecting country...');
              await page.click('#nav-global-location-slot, #glow-ingress-block, #nav-main-ftr-location-slot', { force: true }).catch(() => null);
              await page.waitForTimeout(2000);
            }

            // 2. Now handle the standard zip/postcode input field!
            const zipInputSelector = '#GLUXZipUpdateInput, #GLUXZipUpdateInput_0, input[aria-label*="zip"], input[aria-label*="code"], input[name="zipCode"]';
            const inputVisible = await page.waitForSelector(zipInputSelector, { state: 'visible', timeout: 15000 }).catch(() => null);
            
            if (inputVisible) {
              await page.evaluate((sel) => {
                const el = document.querySelector(sel) as HTMLInputElement;
                if (el) {
                  el.value = '';
                  el.focus();
                }
              }, zipInputSelector);
              
              await page.type(zipInputSelector, locConfig.zip, { delay: 60 });
              await page.keyboard.press('Enter');
              
              const applyBtn = '#GLUXZipUpdate input[type="submit"], #GLUXZipUpdate > span > input, #GLUXZipUpdate_Buttons input, #GLUXZipUpdate input.a-button-input, #GLUXZipUpdate_Buttons span.a-button-inner input';
              await page.click(applyBtn).catch(() => null);
              
              // CRITICAL: Wait 1.5s for backend to register zip
              await page.waitForTimeout(1500);
              
              const confirmBtn = '#GLUXConfirmClose, #GLUXConfirmResponse, input[data-action="GLUXConfirmResponse"], .a-popover-footer input, #GLUXConfirmClose input, #GLUXConfirmClose-announce, .a-popover-footer span.a-button-inner input, button[name="glowDoneButton"]';
              const confirmBtnVisible = await page.waitForSelector(confirmBtn, { timeout: 8000 }).catch(() => null);
              if (confirmBtnVisible) {
                await page.click(confirmBtn).catch(() => null);
              }
              
              await page.waitForTimeout(1500);
              
              // FORCE REFRESH/RELOAD to ensure the new delivery address is applied and price/shipping times are updated
              console.log(`Reloading page for ${domain} after location zip code unlock...`);
              await page.reload({ waitUntil: 'domcontentloaded', timeout: 30000 }).catch(() => null);
            } else {
              console.warn("Zip input popover never appeared.");
            }
          }
        }
      } catch (err: any) {
        console.warn("Location UI injection skipped or failed:", err.message);
      }

      await page.waitForLoadState('domcontentloaded', { timeout: 20000 }).catch(() => null);
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

    // FIXED shipping extraction
    let rawShippingTime = "";
    const primaryDelivery = $('#mir-layout-DELIVERY_BLOCK-slot-PRIMARY_DELIVERY_MESSAGE_LARGE').text().trim() ||
                            $('#deliveryBlockMessage').text().trim();
    const secondaryDelivery = $('#mir-layout-DELIVERY_BLOCK-slot-SECONDARY_DELIVERY_MESSAGE_LARGE').text().trim();

    // Promote secondaryDelivery if it signals a faster/premium option.
    // Added: 'tomorrow', 'morgen', 'demain', 'domani', 'jutro', 'mañana' (next-day signals
    // used on .co.uk/.de/.fr/.it/.nl when the 3P seller offers Prime next-day delivery).
    // Also added: 'today', 'vandaag', 'aujourd', 'oggi', 'heute' (same-day signals).
    const isFastDeliverySignal = (text: string) => {
      const t = text.toLowerCase();
      return t.includes('fastest') ||
             t.includes('snelste') ||
             t.includes('rapide') ||
             t.includes('tomorrow') ||
             t.includes('morgen') ||       // DE/NL "tomorrow"
             t.includes('demain') ||       // FR
             t.includes('domani') ||       // IT
             t.includes('jutro') ||        // PL
             t.includes('mañana') ||       // ES
             t.includes('imorgon') ||      // SE
             t.includes('today') ||
             t.includes('vandaag') ||
             t.includes('aujourd') ||
             t.includes('oggi') ||
             t.includes('heute');
    };

    if (secondaryDelivery && isFastDeliverySignal(secondaryDelivery)) {
      rawShippingTime = secondaryDelivery;
    } else if (primaryDelivery && isFastDeliverySignal(primaryDelivery)) {
      // Primary slot itself signals next-day/same-day — use it directly
      rawShippingTime = primaryDelivery;
    } else {
      const deliveryBlock = $('#mir-layout-DELIVERY_BLOCK-slot-PRIMARY_DELIVERY_MESSAGE_LARGE, #mir-layout-DELIVERY_BLOCK, #deliveryBlockMessage');
      const deliveryTimeAttr = deliveryBlock.find('span[data-csa-c-delivery-time]').attr('data-csa-c-delivery-time');
      if (deliveryTimeAttr) {
        rawShippingTime = deliveryTimeAttr;
      } else {
        rawShippingTime = primaryDelivery ||
                         deliveryBlock.find('.a-text-bold').first().text().trim() ||
                         deliveryBlock.text().trim() ||
                         "";
      }
    }
    rawShippingTime = rawShippingTime.replace(/\s+/g, ' ').trim();

    let amazonDesc = $('#productDescription').text().trim();
    if (!amazonDesc) amazonDesc = $('#feature-bullets').text().trim();
    
    const hasAPlus = !!($('#aplus').length || $('#aplus_feature_div').length || $('div[id*="aplus"]').length);

    // --- 3. Buybox Owner Extraction (Modern tabular design first) ---
    let amazonBuyboxOwner = buyBoxContext.find('div[tabular-attribute-name="Sold by"] .tabular-buybox-text').first().text().trim() ||
                            buyBoxContext.find('div[tabular-attribute-name="Verkauf durch"] .tabular-buybox-text').first().text().trim() ||
                            buyBoxContext.find('div[tabular-attribute-name="Shipper / Seller"] .tabular-buybox-text').first().text().trim() ||
                            buyBoxContext.find('div[tabular-attribute-name="Verzonden en verkocht door"] .tabular-buybox-text').first().text().trim() ||
                            buyBoxContext.find('div[tabular-attribute-name="Dispatched from and sold by"] .tabular-buybox-text').first().text().trim() ||
                            buyBoxContext.find('div[offer-display-feature-name="desktop-merchant-info"] .offer-display-feature-text-message').first().text().trim() ||
                            buyBoxContext.find('#sellerProfileTriggerId').first().text().trim() ||
                            buyBoxContext.find('#merchant-info a').first().text().trim() ||
                            $('div[tabular-attribute-name="Sold by"] .tabular-buybox-text').first().text().trim() ||
                            $('div[tabular-attribute-name="Verkauf durch"] .tabular-buybox-text').first().text().trim() ||
                            $('div[tabular-attribute-name="Shipper / Seller"] .tabular-buybox-text').first().text().trim() ||
                            $('div[tabular-attribute-name="Verzonden en verkocht door"] .tabular-buybox-text').first().text().trim() ||
                            $('div[tabular-attribute-name="Dispatched from and sold by"] .tabular-buybox-text').first().text().trim() ||
                            $('div[offer-display-feature-name="desktop-merchant-info"] .offer-display-feature-text-message').first().text().trim() ||
                            $('#sellerProfileTriggerId').first().text().trim() ||
                            $('#merchant-info a').first().text().trim();

    if (!amazonBuyboxOwner) {
      // Get merchant-info text scoped to buybox first, then global fallback
      const merchantInfoEl = buyBoxContext.find('#merchant-info').first();
      const mInfoRaw = merchantInfoEl.text() || $('#merchant-info').first().text();
      const mInfo = mInfoRaw.toLowerCase();

      // FIXED: Only treat as Amazon-sold when Amazon is explicitly the SELLER,
      // not just the dispatcher. FBA sellers dispatch from Amazon but are NOT Amazon.
      // Exact phrases that mean Amazon itself holds the buybox:
      const amazonIsSeller =
        /\bsold by amazon\b/i.test(mInfoRaw) ||
        /\bdispatched from and sold by amazon\b/i.test(mInfoRaw) ||
        /\bverkauf durch amazon\b/i.test(mInfoRaw) ||
        /\bexpédié et vendu par amazon\b/i.test(mInfoRaw) ||
        /\bverzonden en verkocht door amazon\b/i.test(mInfoRaw) ||
        /\bspedito da e venduto da amazon\b/i.test(mInfoRaw) ||
        /\bvendido por amazon\b/i.test(mInfoRaw) ||
        /\bverkauft von amazon\b/i.test(mInfoRaw);

      if (amazonIsSeller) {
        amazonBuyboxOwner = "Amazon";
      } else if (mInfo.length > 0) {
        // Extract the actual seller name from merchant-info.
        // The anchor tag inside merchant-info links to the seller's storefront.
        const sellerLink = merchantInfoEl.find('a').first().text().trim() ||
                           $('#merchant-info a').first().text().trim();
        if (sellerLink && sellerLink.toLowerCase() !== 'amazon') {
          amazonBuyboxOwner = sellerLink;
        } else {
          // Last resort: strip known prefixes and return the raw text
          amazonBuyboxOwner = mInfoRaw
            .replace(/Dispatched from and sold by\s*/i, '')
            .replace(/Dispatched from Amazon\s*\.?\s*/i, '')
            .replace(/Sold by\s*/i, '')
            .replace(/Fulfilled by Amazon\s*\.?\s*/i, '')
            .replace(/\|.*$/s, '')           // strip everything after a pipe
            .trim();
        }
      } else {
        amazonBuyboxOwner =
          buyBoxContext.find('.offer-display-feature-text-message').first().text().trim() ||
          $('#desktop_buybox .offer-display-feature-text-message').first().text().trim() ||
          $('#rightCol .offer-display-feature-text-message').first().text().trim() ||
          "";
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
          // ── Date-based calculation (existing logic, unchanged) ─────────────────
          const dayMatch = rawShippingTime.match(/(\d{1,2})(?:\.?\s*(?:de|di|d')?\s*)(?:Jan|Feb|Mär|Apr|Mai|Jun|Jul|Aug|Sep|Okt|Nov|Dez|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|janv|févr|mars|avr|mai|juin|juil|août|sept|oct|nov|déc|enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre|maggio|giugno|luglio|wrze|paź|listopad|grudzień|styczeń|luty|kwiecień|maj|maj|maj|czerwiec|lipiec|sierpień|maj|maja|marca|kwietnia|lutego|stycznia|maja|maja|mája|maja|maj|maju|lipca|sierpnia|września|października|listopada|grudnia)/i) ||
                           rawShippingTime.match(/(?:Jan|Feb|Mär|Apr|Mai|Jun|Jul|Aug|Sep|Okt|Nov|Dez|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|janv|févr|mars|avr|mai|juin|juil|août|sept|oct|nov|déc|enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre|maggio|giugno|luglio|maj|mai|maj|maj|maja|marca|kwietnia|lutego|stycznia|maja|maja|mája|maja|maj|maju|lipca|sierpnia|września|października|listopada|grudnia)(?:\s*(?:de|di)?\s*)(\d{1,2})/i);

          let targetDate: Date | null = null;

          if (dayMatch) {
            const day = parseInt(dayMatch[1]);
            const monthMatchStr = dayMatch[0].toLowerCase();

            const monthsEn = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
            const monthsDe = ['jan', 'feb', 'mär', 'apr', 'mai', 'jun', 'jul', 'aug', 'sep', 'okt', 'nov', 'dez'];
            const monthsEs = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic', 'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
            const monthsIt = ['gen', 'feb', 'mar', 'apr', 'mag', 'giu', 'lug', 'ago', 'set', 'ott', 'nov', 'dic', 'maggio', 'giugno', 'luglio'];
            const monthsPl = ['sty', 'lut', 'mar', 'kwi', 'maj', 'cze', 'lip', 'sie', 'wrz', 'paź', 'lis', 'gru', 'stycznia', 'lutego', 'marca', 'kwietnia', 'maja', 'czerwca', 'lipca', 'sierpnia', 'września', 'października', 'listopada', 'grudnia'];
            const monthsSe = ['jan', 'feb', 'mar', 'apr', 'maj', 'jun', 'jul', 'aug', 'sep', 'okt', 'nov', 'dec'];
            const monthsNl = ['jan', 'feb', 'maa', 'apr', 'mei', 'jun', 'jul', 'aug', 'sep', 'okt', 'nov', 'dec'];

            let monthIndex = -1;
            [monthsEn, monthsDe, monthsEs, monthsIt, monthsPl, monthsSe, monthsNl].forEach(mList => {
              const idx = mList.findIndex(m => monthMatchStr.includes(m));
              if (idx !== -1) { monthIndex = idx % 12; }
            });

            if (monthIndex !== -1) {
              targetDate = new Date(today.getFullYear(), monthIndex, day);
              if (targetDate < today && monthIndex < 2) {
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
