import express from 'express';
import { chromium } from 'playwright';
import * as cheerio from 'cheerio';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import axios from 'axios';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';
import stringSimilarity from 'string-similarity';

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

  if (live.images && live.images.length >= (master.images?.count || 1)) result.images.match = true;

  return result;
}

// 3. Audit Bol.com
app.post("/api/audit/bol", async (req, res) => {
  let browser;
  try {
    const { ean, masterData } = req.body;
    const searchUrl = `https://www.bol.com/nl/nl/s/?searchtext=${ean}`;
    const proxyServer = process.env.PROXY_SERVER;
    
    const launchOptions: any = {
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    };
    if (proxyServer) {
      launchOptions.proxy = {
        server: proxyServer,
        username: process.env.PROXY_USERNAME,
        password: process.env.PROXY_PASSWORD
      };
    }

    browser = await chromium.launch(launchOptions);
    const context = await browser.newContext();
    await context.addCookies([
      { name: 'cookie_consent', value: 'true', domain: '.bol.com', path: '/' },
      { name: 'bol_gdpr_consent', value: 'yes', domain: '.bol.com', path: '/' }
    ]);
    const page = await context.newPage();
    await page.goto(searchUrl, { waitUntil: 'networkidle', timeout: 45000 });

    const content = await page.content();
    const $ = cheerio.load(content);
    
    // Find first product
    const firstProduct = $('.product-item--grid, .product-item--list').first();
    let productUrl = firstProduct.find('a.product-title').attr('href');
    if (productUrl && !productUrl.startsWith('http')) productUrl = 'https://www.bol.com' + productUrl;

    if (productUrl) {
      await page.goto(productUrl, { waitUntil: 'networkidle' });
      const productContent = await page.content();
      const $p = cheerio.load(productContent);
      
      const bolTitle = $p('h1.page-heading').text().trim();
      const bolPrice = $p('.promo-price').text().trim().replace(/\s+/g, '');
      const bolDesc = $p('.product-description').text().trim();
      const bolImages: string[] = [];
      $p('.js_bundle_image, .js_product_img').each((_, el) => {
        const src = $p(el).attr('src') || $p(el).attr('data-src');
        if (src) bolImages.push(src);
      });

      const liveData = {
        title: bolTitle,
        price: bolPrice,
        description: bolDesc,
        images: Array.from(new Set(bolImages)),
        url: productUrl,
        hasAPlus: false,
        shippingDays: "N/A",
        variations: 0,
        bullets: []
      };

      const auditResult = await performAudit(masterData, liveData, 'bol');
      res.json({ liveData, auditResult });
    } else {
      res.status(404).json({ error: "Product not found on Bol" });
    }
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
    
    const targetSheetName = "Amazon QC results";
    // Case-insensitive search for the sheet
    const sheet = doc.sheetsByTitle[targetSheetName] || 
                  doc.sheetsByIndex.find(s => s.title.toLowerCase().trim() === targetSheetName.toLowerCase());
    
    if (!sheet) {
      const availableSheets = doc.sheetsByIndex.map(s => `"${s.title}"`).join(', ');
      throw new Error(`Sheet "${targetSheetName}" not found. Available sheets: ${availableSheets}`);
    }

    // Mapping according to instructions
    const bulletMatchCount = (auditResult && Array.isArray(auditResult.bullets)) ? auditResult.bullets.filter((b: any) => b.match).length : 0;
    const bulletMatchText = bulletMatchCount > 0 ? 'Yes' : 'No';

    const timestamp = new Date().toLocaleString();
    const resultRow: any = {
      'Date': timestamp,
      'date': timestamp,
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

    // Images mapping - User wants "each relevant column based on number image"
    if (masterData && masterData.images && Array.isArray(masterData.images)) {
      masterData.images.slice(0, 10).forEach((url: string, i: number) => {
        const formula = url ? `=IMAGE("${url}")` : '';
        const index = i + 1;
        resultRow[`Amazon Master Image ${index}`] = formula;
        resultRow[`AMZ Master Image ${index}`] = formula;
        resultRow[`Master Image ${index}`] = formula;
        resultRow[`AMZ IMG ${index}`] = formula;
        resultRow[`image amazon data ${index}`] = formula;
        resultRow[`Image Amazon Data ${index}`] = formula;
      });
    }

    if (liveData && liveData.images && Array.isArray(liveData.images)) {
      liveData.images.slice(0, 10).forEach((url: string, i: number) => {
        const formula = url ? `=IMAGE("${url}")` : '';
        const index = i + 1;
        resultRow[`Amazon Live Image ${index}`] = formula;
        resultRow[`AMZ Live Image ${index}`] = formula;
        resultRow[`Live Image ${index}`] = formula;
        resultRow[`LIVE IMG ${index}`] = formula;
        resultRow[`image amazon live ${index}`] = formula;
        resultRow[`Image Amazon Live ${index}`] = formula;
      });
    }

    // Search for existing row to override
    const rows = await sheet.getRows();
    const existingRow = rows.find(r => r.get('ASIN/EAN') === identifier || r.get('ASIN') === identifier || r.get('Identifier') === identifier);

    const headers = sheet.headerValues;
    const findHeader = (target: string) => {
      return headers.find(h => h.toLowerCase().trim() === target.toLowerCase().trim());
    };

    if (existingRow) {
      // Update existing row
      Object.keys(resultRow).forEach(key => {
        const matchingHeader = findHeader(key);
        if (matchingHeader) {
          existingRow.set(matchingHeader, resultRow[key]);
        }
      });
      await existingRow.save();
    } else {
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
    }
    
    res.json({ success: true });
  } catch (error: any) {
    console.error("Save Audit Error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/sheets/clear-sheet", async (req, res) => {
  try {
    const { sheetId, sheetName } = req.body;
    
    const auth = new JWT({
      email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const doc = new GoogleSpreadsheet(sheetId || '1V4lNf30SlBwczSvGX9rfn5eWFH2AvMO4TqMHAHalS7s', auth);
    await doc.loadInfo();
    const sheet = doc.sheetsByTitle[sheetName];
    if (!sheet) return res.status(404).json({ error: "Sheet not found" });

    // Get all rows
    const rows = await sheet.getRows();
    // Delete them all
    for (const row of rows) {
      await row.delete();
    }
    
    res.json({ success: true });
  } catch (error: any) {
    console.error("Clear Sheet Error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/sheets/batch-save-audit", async (req, res) => {
  res.json({ success: true });
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
