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

    const amazonLocalizationMap: Record<string, { locale: string; timezoneId: string; city: string; zip: string; currency: string; deliverTo: string[] }> = {
      'amazon.co.uk': { locale: 'en-GB', timezoneId: 'Europe/London', city: 'LND', zip: 'SW1A 1AA', currency: 'GBP', deliverTo: ['Deliver to', 'Livre à'] },
      'amazon.de': { locale: 'de-DE', timezoneId: 'Europe/Berlin', city: 'BER', zip: '10117', currency: 'EUR', deliverTo: ['Lieferung nach', 'Liefern an', 'Deliver to'] },
      'amazon.fr': { locale: 'fr-FR', timezoneId: 'Europe/Paris', city: 'PAR', zip: '75001', currency: 'EUR', deliverTo: ['Livrer à', 'Livraison à', 'Deliver to'] },
      'amazon.it': { locale: 'it-IT', timezoneId: 'Europe/Rome', city: 'ROM', zip: '00118', currency: 'EUR', deliverTo: ['Invia a', 'Consegna a', 'Deliver to'] },
      'amazon.es': { locale: 'es-ES', timezoneId: 'Europe/Madrid', city: 'MAD', zip: '28001', currency: 'EUR', deliverTo: ['Enviar a', 'Entrega en', 'Deliver to'] },
      'amazon.nl': { locale: 'nl-NL', timezoneId: 'Europe/Amsterdam', city: 'AMS', zip: '1011 AB', currency: 'EUR', deliverTo: ['Bezorgen in', 'Deliver to'] },
      'amazon.pl': { locale: 'pl-PL', timezoneId: 'Europe/Warsaw', city: 'WAW', zip: '00-001', currency: 'PLN', deliverTo: ['Dostawa do', 'Wyślij do', 'Deliver to'] },
      'amazon.se': { locale: 'sv-SE', timezoneId: 'Europe/Stockholm', city: 'STO', zip: '111 20', currency: 'SEK', deliverTo: ['Skicka till', 'Leverera till', 'Deliver to'] },
      'amazon.com.be': { locale: 'nl-BE', timezoneId: 'Europe/Brussels', city: 'BRU', zip: '1000', currency: 'EUR', deliverTo: ['Bezorgen in', 'Livrer à', 'Deliver to'] },
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
      await page.goto(url, { waitUntil: 'load', timeout: 70000 });
      
      try {
        const cookieButtons = ['#sp-cc-accept', 'input[name="accept"]', '#cookie-accept', '#accept-cookies', '.a-button-inner input[data-action="accept-cookies"]'];
        for (const selector of cookieButtons) {
          if (await page.isVisible(selector)) {
            await Promise.all([
              page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 10000 }).catch(() => null),
              page.click(selector)
            ]);
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
          // Use more robust location button selector
          const locBtn = await page.waitForSelector('#nav-global-location-slot, #glow-ingress-block, #nav-main-ftr-location-slot', { visible: true, timeout: 15000 }).catch(() => null);
          if (locBtn) {
            await locBtn.click({ force: true });
            
            // Wait for popover to appear - use broader selector for input
            // Sometimes it's a different popover or needs a moment
            await page.waitForTimeout(1000);
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
              
              await page.type(zipInputSelector, locConfig.zip, { delay: 100 });
              
              const applyBtn = '#GLUXZipUpdate input[type="submit"], #GLUXZipUpdate > span > input, #GLUXZipUpdate_Buttons input, #GLUXZipUpdate input.a-button-input, #GLUXZipUpdate_Buttons span.a-button-inner input';
              await page.click(applyBtn).catch(() => null);
              
              // CRITICAL: Wait 1.5s for backend to register zip
              await page.waitForTimeout(2000);
              
              const confirmBtn = '#GLUXConfirmClose, #GLUXConfirmResponse, input[data-action="GLUXConfirmResponse"], .a-popover-footer input, #GLUXConfirmClose input, .a-popover-footer span.a-button-inner input';
              const confirmBtnVisible = await page.waitForSelector(confirmBtn, { timeout: 8000 }).catch(() => null);
              if (confirmBtnVisible) {
                await page.click(confirmBtn).catch(() => null);
              }
              
              await page.waitForTimeout(1000);
              await page.reload({ waitUntil: 'domcontentloaded', timeout: 60000 }).catch(() => null);
            } else {
              console.warn("Zip input popover never appeared.");
            }
          }
        }
      } catch (err) {
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

    let amazonPrice = $('.apex-core-price-identifier .a-offscreen').first().text().trim();
    if (!amazonPrice) {
      amazonPrice = $('.a-price .a-offscreen').first().text().trim() || 
                    $('.apexPriceToPay .a-offscreen').first().text().trim() ||
                    $('#price_inside_buybox').text().trim() ||
                    $('#priceBlockPremiumPrice').text().trim() || 
                    $('.a-price-whole').first().parent().text().trim() || "";
    }
    amazonPrice = amazonPrice.replace(/\s+/g, ' ').trim();

    let listPrice = $('.basisPrice .a-offscreen').text().trim() || "";

    let rawShippingTime = "";
    const deliveryBlock = $('#mir-layout-DELIVERY_BLOCK-slot-PRIMARY_DELIVERY_MESSAGE_LARGE, #mir-layout-DELIVERY_BLOCK, #deliveryBlockMessage');
    const deliveryTimeAttr = deliveryBlock.find('span[data-csa-c-delivery-time]').attr('data-csa-c-delivery-time');
    
    if (deliveryTimeAttr) {
      rawShippingTime = deliveryTimeAttr;
    } else {
      rawShippingTime = deliveryBlock.find('.a-text-bold').first().text().trim() || 
                       deliveryBlock.find('#mir-layout-DELIVERY_BLOCK-slot-PRIMARY_DELIVERY_MESSAGE_LARGE').text().trim() ||
                       deliveryBlock.text().trim() || "";
    }
    rawShippingTime = rawShippingTime.replace(/\s+/g, ' ').trim();

    let amazonDesc = $('#productDescription').text().trim();
    if (!amazonDesc) amazonDesc = $('#feature-bullets').text().trim();
    
    const hasAPlus = !!($('#aplus').length || $('#aplus_feature_div').length || $('div[id*="aplus"]').length);

    const uniqueImages: string[] = [];
    
    // 1. Try to extract from colorImages JSON (most accurate)
    const scriptContent = $('script').map((_, el) => $(el).html()).get().join('\n');
    const colorImagesMatch = scriptContent.match(/'colorImages':\s*({.+?}),?\n/s) || 
                             scriptContent.match(/"colorImages":\s*({.+?}),?\n/s);
    
    if (colorImagesMatch) {
      try {
        const colorImages = JSON.parse(colorImagesMatch[1]);
        const initialImages = colorImages.initial || [];
        initialImages.forEach((img: any) => {
          const url = img.hiRes || img.large || img.main?.url;
          if (url) uniqueImages.push(url);
        });
      } catch (e) { /* ignore parse error */ }
    }

    // 2. Fallback to imageBlock specific selectors if JSON extraction failed or is incomplete
    if (uniqueImages.length === 0) {
      $('#imageBlock img, #altImages li.imageThumbnail img, #main-image-container img').each((_, el) => {
        let src = $(el).attr('src') || $(el).attr('data-old-hires') || $(el).attr('data-a-dynamic-image');
        
        if (src && src.startsWith('{')) {
          try {
            const urls = Object.keys(JSON.parse(src));
            if (urls.length) src = urls[0];
          } catch(e) { src = null; }
        }

        if (src && typeof src === 'string') {
          const cleaned = src.replace(/\._[A-Z0-9,_-]+\./g, '.');
          if (cleaned.includes('media-amazon.com/images/I/')) {
            uniqueImages.push(cleaned);
          }
        }
      });
    }

    const amazonBullets: string[] = [];
    const bulletSelectors = [
      '#feature-bullets ul li',
      '#featurebullets_feature_div ul li',
      '.a-unordered-list.a-vertical.a-spacing-mini li',
      '#productDescription_feature_div ul li',
      '#feature-bullets-content li',
      '[data-feature-name="product-facts"] .a-list-item',
      '.product-facts-title + .a-unordered-list li',
      '#product-facts-grid li',
      '.a-section.a-spacing-medium .a-list-item',
      '#productFactsDesktopExpander .a-list-item',
      '.a-expander-content .a-list-item'
    ];
    $(bulletSelectors.join(', ')).each((_, el) => {
      const text = $(el).find('span.a-list-item').text().trim() || $(el).text().trim();
      if (text && !text.includes('Make sure this fits') && !text.includes('Geben Sie Ihr Modell ein') && !text.includes('Sprawdź, czy pasuje') && text.length > 5) {
        amazonBullets.push(text);
      }
    });

    const variations: string[] = [];
    // Various variation selectors for different categories
    const varSelectors = [
      '#twister .a-row.variation-row',
      '.twister-selection-column',
      'li[id^="color_name_"]',
      'li[id^="size_name_"]',
      '.tp-inline-twister-dim-values-container',
      '#inline-twister-row-all-options',
      '.a-button-toggle',
      '.swatchAvailable',
      '#variation_color_name li',
      '#variation_size_name li',
      '.twisterSwatchWrapper',
      '.swatchDimLink',
      '.swatchSelect',
      '#twister .a-list-item[data-asin]',
      '#inline-twister-row-all-options .a-list-item',
      '.visualSelection .visual-selection-button',
      '.dimension-selection-container',
      '[id^="variation_"]',
      '.twister-selection-container'
    ];
    $(varSelectors.join(', ')).each((_, el) => {
      const text = $(el).text().trim() || $(el).attr('data-asin') || "";
      if (text && text.length > 0) variations.push(text);
    });

    // Calculate shipping days difference
    let shippingDays = "N/A";
    try {
      if (rawShippingTime) {
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        // More robust date matching for ES, IT, PL, SE, etc.
        // Captures "2 de mayo", "el 2 de mayo", "2 di maggio", etc.
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
            if (idx !== -1) {
              monthIndex = idx % 12;
            }
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
      variations: variations.length,
      hasAPlus: hasAPlus,
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
    variations: { match: live.variations > 0 }
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

  return result;
}

// --- Bol.com Helpers ---
const BOL_USER_AGENTS = [
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
  'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36'
];

async function goToProduct(page: any, searchTerm: string): Promise<string> {
  const searchUrl = `https://www.bol.com/nl/nl/s/?searchtext=${encodeURIComponent(searchTerm)}`;
  
  const ua = BOL_USER_AGENTS[Math.floor(Math.random() * BOL_USER_AGENTS.length)];
  const fetchOpts = {
    headers: {
      'User-Agent': ua,
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
      'Accept-Language': 'nl-NL,nl;q=0.9,en-US;q=0.8,en;q=0.7',
      'Cache-Control': 'no-cache',
      'Pragma': 'no-cache'
    }
  };

  let productUrl: string | null = null;
  let searchHtml = '';

  console.log(`🔎 Attempting Bol.com direct search → ${searchUrl}`);
  try {
    const searchRes = await fetch(searchUrl, fetchOpts);
    searchHtml = await searchRes.text();
    const lowerHtml = searchHtml.toLowerCase();
    const isWaf = lowerHtml.includes('ip adres is geblokkeerd') || 
                  lowerHtml.includes('akamai') || 
                  lowerHtml.includes('rustig aan speed racer') ||
                  (searchHtml.length < 5000 && lowerHtml.includes('<title>bol</title>'));

    if (!isWaf) {
      await page.setContent(searchHtml);
      productUrl = await page.evaluate(() => {
        let el = document.querySelector('a.product-title, a.product-item--row__title, a[data-test="product-title"]');
        if (!el) el = document.querySelector('.product-list a[href*="/p/"]');
        if (!el) {
            const links = Array.from(document.querySelectorAll('a[href*="/p/"]'));
            el = links.find(a => {
               const href = (a as HTMLAnchorElement).href;
               return href && !href.includes('review') && !href.includes('login') && href.includes('/p/');
            }) || null;
        }
        return el ? (el as HTMLAnchorElement).href : null;
      });
    } else {
      console.warn('⚠️ Bol.com search blocked (WAF). Trying DuckDuckGo fallback...');
    }
  } catch (e) {
    console.warn(`⚠️ Direct search failed: ${(e as Error).message}. Trying DuckDuckGo fallback...`);
  }

  // Fallback to DuckDuckGo if direct search failed or was blocked
  if (!productUrl) {
    console.log(`🔎 Searching DDG for site:bol.com ${searchTerm}`);
    try {
      const ddgUrl = `https://html.duckduckgo.com/html/?q=site:bol.com+${encodeURIComponent(searchTerm)}`;
      const ddgRes = await fetch(ddgUrl, { headers: { 'User-Agent': ua } });
      const ddgHtml = await ddgRes.text();
      await page.setContent(ddgHtml);

      productUrl = await page.evaluate(() => {
        const links = Array.from(document.querySelectorAll('a.result__snippet, a.result__url'));
        for (const a of links) {
          try {
            const href = (a as HTMLAnchorElement).href;
            if (href.includes('bol.com') && href.includes('/p/')) {
              // Extract from DDG redirect if needed
              if (href.includes('uddg=')) {
                const url = new URL(href);
                const uddg = url.searchParams.get('uddg');
                if (uddg) return decodeURIComponent(uddg);
              }
              return href;
            }
          } catch (e) {}
        }
        return null;
      });
    } catch (e) {
      console.error(`❌ DDG fallback also failed: ${(e as Error).message}`);
    }
  }

  if (!productUrl) {
    throw new Error(`WAF_BLOCKED: Bol.com results blocked and fallback search failed for EAN ${searchTerm}.`);
  }

  // Ensure leading domain if it's relative
  if (productUrl.startsWith('about:blank')) {
       productUrl = productUrl.replace('about:blank', 'https://www.bol.com');
  } else if (!productUrl.startsWith('http')) {
       productUrl = 'https://www.bol.com' + (productUrl.startsWith('/') ? productUrl : '/' + productUrl);
  }

  console.log(`🖱️ Fetching product URL directly: ${productUrl}`);
  const productRes = await fetch(productUrl, fetchOpts).catch(err => {
      throw new Error(`Product fetch failed: ${err.message}`);
  });
  
  const productHtml = await productRes.text();
  const prodLower = productHtml.toLowerCase();
  const isProdWaf = prodLower.includes('ip adres is geblokkeerd') || 
                    prodLower.includes('akamai') ||
                    (productHtml.length < 5000 && prodLower.includes('<title>bol</title>'));

  if (isProdWaf) {
     throw new Error("WAF_BLOCKED: Bol.com blocked the product page request even via direct link.");
  }

  await page.setContent(productHtml);
  return productUrl;
}

async function extractCatalogue(page: any) {
  await page.waitForLoadState('load', { timeout: 45_000 }).catch(() => null);
  await page.waitForLoadState('networkidle', { timeout: 45_000 }).catch(() => null);
  await page.waitForTimeout(2_000);

  await page.mouse.wheel(0, 400);
  await page.waitForTimeout(2_000);

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
    const priceMatch = pageHtml.match(/"offers"\s*:\s*\{[^}]*"@type"\s*:\s*"Offer"[^}]*"price"\s*:\s*"([\d.]+)"[^}]*"priceCurrency"\s*:\s*"EUR"/);
    if (priceMatch) {
      price = priceMatch[1];
    } else {
      // Fallback
      const m = pageHtml.match(/"price"\s*:\s*([\d.]+)\s*,\s*"priceCurrency"\s*:\s*"EUR"/);
      if (m) price = m[1];
    }

    if (price === 'N/A') {
      const all = document.body.innerText;
      const m = all.match(/€\s*([\d.]+,\d{2})/);
      if (m) price = m[1].replace(',', '.');
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
    const allImages = Array.from(document.querySelectorAll('img'));
    allImages.forEach(img => {
      const alt = img.getAttribute('alt') || '';
      if (alt.includes('Afbeelding nummer')) {
        const src = img.src || img.getAttribute('data-src') || '';
        if (src && src.startsWith('http')) imgs.push(src);
      }
    });

    if (imgs.length === 0) {
      // Fallback
      const mainSel = [
        '[data-test="product-main-image"] img',
        '.js_main_product_image',
        '.pdp-main-image img',
        'img.js_main_product_image'
      ];
      mainSel.forEach(sel => {
        const img = document.querySelector(sel) as HTMLImageElement | null;
        if (img && img.src && img.src.startsWith('http')) imgs.push(img.src);
      });

      const thumbSel = [
        '.js_product_media_items img',
        '.pdp-images img',
        '.js_image_container img',
        '.product-images__item img'
      ];
      thumbSel.forEach(sel => {
        const thumbs = Array.from(document.querySelectorAll(sel)) as HTMLImageElement[];
        thumbs.forEach(i => {
          const src = i.src || i.getAttribute('data-src') || i.getAttribute('src');
          if (src && src.includes('media.s-bol.com')) imgs.push(src);
        });
      });
    }

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

    const launchOpts: any = {
      headless: true,
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-blink-features=AutomationControlled'
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
    const context = await browser.newContext({
      userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
      viewport: { width: Math.floor(Math.random() * (1920 - 1366 + 1)) + 1366, height: Math.floor(Math.random() * (1080 - 768 + 1)) + 768 },
      locale: 'nl-NL',
      extraHTTPHeaders: {
        'Accept-Language': 'nl-NL,nl;q=0.9,en-US;q=0.8,en;q=0.7',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'sec-ch-ua': '"Not A(Brand";v="99", "Google Chrome";v="121", "Chromium";v="121"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'document',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-site': 'none',
        'sec-fetch-user': '?1',
        'upgrade-insecure-requests': '1'
      }
    });

    await context.addInitScript(() => {
      // @ts-ignore
      Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
    });

    const page = await context.newPage();
    
    const productUrl = await goToProduct(page, ean);
    const data = await extractCatalogue(page);
    
    const bolShippingDays = calculateBolShippingDays(data.shipping);

    const liveData = {
      title: data.title,
      price: data.price,
      description: data.description,
      images: data.images,
      url: productUrl,
      hasAPlus: false,
      shipping: bolShippingDays !== "N/A" ? `${bolShippingDays} days` : data.shipping,
      shippingDays: bolShippingDays,
      rawShipping: data.shipping,
      variations: data.liveVariations && data.liveVariations.length > 5 ? data.liveVariations.split('|').length || 1 : 0,
      bullets: data.bullets,
      rawVariationsText: data.liveVariations || ''
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
    const { sheetId, sheetName } = req.body;
    const auth = new JWT({
      email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const doc = new GoogleSpreadsheet(sheetId, auth);
    await doc.loadInfo();
    const sheet = doc.sheetsByTitle[sheetName] || doc.sheetsByIndex[0];
    const rows = await sheet.getRows();
    const data = rows.map(row => row.toObject());
    res.json({ data });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/sheets/save-audit", async (req, res) => {
  try {
    const { mode, identifier, auditResult, liveData, masterData } = req.body;
    
    const auth = new JWT({
      email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const spreadsheetId = '1V4lNf30SlBwczSvGX9rfn5eWFH2AvMO4TqMHAHalS7s';
    const doc = new GoogleSpreadsheet(spreadsheetId, auth);
    await doc.loadInfo();
    
    const targetSheetName = mode === "bol" ? "Bol QC results" : "Amazon QC results";
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
      liveData.images.slice(0, 10).forEach((url: string, i: number) => {
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
    const { mode, audits } = req.body;
    if (!audits || !Array.isArray(audits) || audits.length === 0) {
      return res.json({ success: true, message: "No audits to save." });
    }

    const auth = new JWT({
      email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const spreadsheetId = '1V4lNf30SlBwczSvGX9rfn5eWFH2AvMO4TqMHAHalS7s';
    const doc = new GoogleSpreadsheet(spreadsheetId, auth);
    await doc.loadInfo();

    const targetSheetName = mode === "bol" ? "Bol QC results" : "Amazon QC results";
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
        liveData.images.slice(0, 10).forEach((url: string, i: number) => {
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
