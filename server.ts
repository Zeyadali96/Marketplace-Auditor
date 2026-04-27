import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import cors from "cors";
import axios from "axios";
import * as cheerio from "cheerio";
import { GoogleSpreadsheet } from "google-spreadsheet";
import { JWT } from "google-auth-library";
import { google } from "googleapis";
import stringSimilarity from "string-similarity";
import sharp from "sharp";
import { differenceInDays, parse } from 'date-fns';
import { chromium } from "playwright-extra";
import stealth from "puppeteer-extra-plugin-stealth";
chromium.use(stealth());

async function startServer() {
  const app = express();
  // If running on Railway, respect its PORT. Otherwise, AI Studio strictly requires 3000.
  const PORT = process.env.IS_RAILWAY === 'true' ? (Number(process.env.PORT) || 8080) : 3000;

  app.use(cors());
  app.use(express.json({ limit: '50mb' }));

  // Health check routes for Railway
  app.get('/health', (req, res) => res.send('OK'));
  app.get('/api/health', (req, res) => res.send('OK'));

  // API Routes
  
  // 1. Fetch Google Sheet Data
  app.post("/api/sheets/fetch", async (req, res) => {
    try {
      const { sheetId, sheetName } = req.body;
      
      const serviceAccountEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
      const privateKey = process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n');

      if (!serviceAccountEmail || !privateKey) {
        return res.status(400).json({ error: "Google Sheets credentials not configured in environment variables." });
      }

      const auth = new JWT({
        email: serviceAccountEmail,
        key: privateKey,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });

      const doc = new GoogleSpreadsheet(sheetId, auth);
      await doc.loadInfo();
      
      const sheet = sheetName ? doc.sheetsByTitle[sheetName] : doc.sheetsByIndex[0];
      
      if (!sheet) {
        return res.status(404).json({ 
          error: `Sheet "${sheetName || 'at index 0'}" not found in spreadsheet. Available sheets: ${Object.keys(doc.sheetsByTitle).join(', ')}` 
        });
      }

      const rows = await sheet.getRows();
      
      const data = rows.map(row => row.toObject());
      res.json({ data, title: doc.title });
    } catch (error: any) {
      console.error("Sheets Error:", error);
      res.status(500).json({ error: error.message });
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
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-gpu', '--single-process']
      };

      if (proxyServer) {
        launchOptions.proxy = {
          server: proxyServer,
          username: proxyUsername || undefined,
          password: proxyPassword || undefined,
        };
      }

      browser = await chromium.launch(launchOptions).catch(err => {
        console.error("AMAZON AUDIT FAILED TO LAUNCH CHROMIUM:", err);
        throw new Error(`Browser launch failed. If you see "libglib" errors, ensure system dependencies are installed. On Railway, the provided nixpacks.toml should fix this. Error: ${err.message}`);
      });
      const context = await browser.newContext({
        userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        viewport: { width: 1920, height: 1080 }
      });

      // Set regional postcodes and language preferences
      const cookies = [
        { name: 'session-id', value: '123-4567890-1234567', domain: `.${domain}`, path: '/' },
        { name: 'ubid-main', value: '123-4567890-1234567', domain: `.${domain}`, path: '/' }
      ];
      
      if (domain.endsWith('.co.uk')) {
        cookies.push({ name: 'lc-main', value: 'en_GB', domain: `.${domain}`, path: '/' });
        cookies.push({ name: 'i18n-prefs', value: 'GBP', domain: `.${domain}`, path: '/' });
        cookies.push({ name: 'sp-cdn', value: '"LND:SW1A 1AA"', domain: `.${domain}`, path: '/' });
      } else if (domain.endsWith('.de')) {
        cookies.push({ name: 'lc-main', value: 'de_DE', domain: `.${domain}`, path: '/' });
        cookies.push({ name: 'i18n-prefs', value: 'EUR', domain: `.${domain}`, path: '/' });
        cookies.push({ name: 'sp-cdn', value: '"BER:10117"', domain: `.${domain}`, path: '/' });
      } else if (domain.endsWith('.pl')) {
        cookies.push({ name: 'lc-main', value: 'pl_PL', domain: `.${domain}`, path: '/' });
        cookies.push({ name: 'i18n-prefs', value: 'PLN', domain: `.${domain}`, path: '/' });
        cookies.push({ name: 'sp-cdn', value: '"WAW:00-001"', domain: `.${domain}`, path: '/' });
      } else if (domain.endsWith('.se')) {
        cookies.push({ name: 'lc-main', value: 'sv_SE', domain: `.${domain}`, path: '/' });
        cookies.push({ name: 'i18n-prefs', value: 'SEK', domain: `.${domain}`, path: '/' });
        cookies.push({ name: 'sp-cdn', value: '"STO:111 20"', domain: `.${domain}`, path: '/' });
      } else {
        cookies.push({ name: 'lc-main', value: 'en_US', domain: `.${domain}`, path: '/' });
      }
      
      await context.addCookies(cookies);
      const page = await context.newPage();
      
      // Optimize loading by blocking unnecessary resources
      await page.route('**/*.{png,jpg,jpeg,gif,webp,svg,woff,woff2,ttf,otf,css}', (route) => {
        const url = route.request().url();
        // Allow main product images if needed, but for scraping HTML we usually don't need them
        // However, we extract image URLs from the script tags or landingImage src, so we don't need the actual image files to load.
        route.abort();
      });

      // Navigate and wait for initial load
      try {
        await page.goto(url, { waitUntil: 'load', timeout: 60000 });
        await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => null);
      } catch (e: any) {
        if (e.name === 'TimeoutError') {
          console.warn("Initial navigation timed out, attempting to proceed with current content...");
        } else {
          throw e;
        }
      }
      
      // Ensure the product title is at least present
      await page.waitForSelector('#productTitle', { timeout: 15000 }).catch(() => {
        console.warn("Product title not found within 15s, page might be slow or blocked.");
      });
      
      // Check for CAPTCHA
      const isCaptcha = await page.evaluate(function() {
        // @ts-ignore
        if (typeof __name === 'undefined') { (window as any).__name = (t: any, v: any) => t; }
        return document.title.includes('Robot Check') || 
               document.body.innerText.includes('Type the characters you see in this image') ||
               document.body.innerText.includes('To discuss automated access to Amazon data please contact');
      });
      
      if (isCaptcha) {
        throw new Error("Amazon blocked the request with a CAPTCHA. Please try again later.");
      }

      // Wait a bit more for dynamic elements (variations, etc.)
      await page.waitForTimeout(3000);
      
      // Get content
      const content = await page.content();
      const $ = cheerio.load(content);

      // 1. Image Extraction (Extracting main and secondary images, deduplicating via getUniqueImages)
      let images: string[] = [];
      
      // High-res main image candidates
      const landingImg = $('#landingImage').attr('data-old-hires') || $('#landingImage').attr('data-a-dynamic-image') || $('#landingImage').attr('src');
      
      if (landingImg) {
        if (landingImg.startsWith('{')) {
          try {
            const dynamicImgs = JSON.parse(landingImg);
            const urls = Object.keys(dynamicImgs);
            if (urls.length > 0) {
              // Extract the largest resolution URL
              const sortedUrls = urls.sort((a, b) => {
                const resA = (dynamicImgs[a] && dynamicImgs[a][0]) || 0;
                const resB = (dynamicImgs[b] && dynamicImgs[b][0]) || 0;
                return resB - resA;
              });
              images.push(sortedUrls[0]);
            }
          } catch(e) {}
        } else if (landingImg.startsWith('http')) {
          images.push(landingImg);
        }
      }
      
      // Gallery images (excluding main image often found here)
      $('#altImages ul li img, #imageBlock_feature_div img, .aplus-v2 img, #imgTagWrapperId img').each((_, el) => {
        const url = $(el).attr('data-old-hires') || $(el).attr('data-a-dynamic-image') || $(el).attr('src');
        if (url && url.startsWith('http') && !url.includes('sprite')) {
          images.push(url);
        } else if (url && url.startsWith('{')) {
          try {
            const dynamicImgs = JSON.parse(url);
            const urls = Object.keys(dynamicImgs);
            if (urls.length > 0) images.push(urls[urls.length - 1]);
          } catch(e) {}
        }
      });
      
      const uniqueImages = getUniqueImages(images, 'amazon');

      // A+ Content Extraction (Exclude Brand Stories)
      const aPlusContainer = $('.aplus-v2, #aplus, #premium-aplus').not('#brandStory_feature_div, .aplus-brand-story-v2, .aplus-brand-story-v1');
      
      // Remove Brand Story sections if they exist within the container
      aPlusContainer.find('.aplus-brand-story-v2, .aplus-brand-story-v1, #brandStory_feature_div, .premium-brand-story, [class*="brand-story"]').remove();
      
      const hasAPlus = aPlusContainer.length > 0;
      let amazonDesc = $('#productDescription').text().trim();
      
      // Extract A+ text ONLY (Simpler revert)
      let aPlusData = null;
      if (hasAPlus) {
        const aPlusText = aPlusContainer.find('p, h3, h4, span, li').map((_, el) => {
          if (el.tagName === 'img') return '';
          if ($(el).closest('.aplus-brand-story-v2, .aplus-brand-story-v1, #brandStory_feature_div, .premium-brand-story, [class*="brand-story"]').length > 0) return '';
          return $(el).text().trim();
        }).get().filter(t => t.length > 0).join('\n\n');
        
        aPlusData = { text: aPlusText };
      }

      // If description is empty but A+ is present, use A+ data
      if (!amazonDesc && aPlusData) {
        amazonDesc = `APLUS_DATA:${JSON.stringify(aPlusData)}`;
      } else if (!amazonDesc) {
        amazonDesc = "No description on detail page";
      }

      // Shipping Extraction (Target: div#mir-layout-DELIVERY_BLOCK)
      const deliverySelectors = [
        '#mir-layout-DELIVERY_BLOCK-slot-PRIMARY_DELIVERY_MESSAGE_LARGE',
        '#mir-layout-DELIVERY_BLOCK',
        '#deliveryBlockMessage',
        '#pfd-desktop-PRIMARY_DELIVERY_MESSAGE_LARGE',
        '#mir-layout-DELIVERY_BLOCK-slot-SECONDARY_DELIVERY_MESSAGE_LARGE',
        '#fastest_delivery_message',
        '#upsell-messaging',
        '#ddmDeliveryMessage',
        '.a-spacing-base .a-text-bold',
        '#pba-delivery-message'
      ];
      
      let rawShippingTime = "";
      for (const selector of deliverySelectors) {
        const el = $(selector);
        if (el.length) {
          rawShippingTime = el.find('span').attr('data-csa-c-delivery-time') || el.attr('data-csa-c-delivery-time') || "";
          if (!rawShippingTime) {
            // Try to find the bold text which usually contains the date
            rawShippingTime = el.find('.a-text-bold').first().text().trim() || el.text().trim();
          }
          if (rawShippingTime) break;
        }
      }
      
      console.log(`Extracted Raw Shipping: ${rawShippingTime}`);
      
      let shippingDaysNum = 0;
      let shippingDaysStr = "N/A";
      if (rawShippingTime) {
        try {
          const today = new Date();
          const currentYear = today.getFullYear();
          
          let dateStr = rawShippingTime;
          // Handle ranges by taking the first date
          if (dateStr.includes('-')) {
            dateStr = dateStr.split('-')[0].trim();
          }
          
          // Comprehensive month mapping for EU Amazon marketplaces
          const euMonthMap: Record<string, number> = {
            // English
            'january': 0, 'february': 1, 'march': 2, 'april': 3, 'may': 4, 'june': 5, 'july': 6, 'august': 7, 'september': 8, 'october': 9, 'november': 10, 'december': 11,
            // Dutch & generic Germanic
            'januari': 0, 'februari': 1, 'maart': 2, 'mei': 4, 'juni': 5, 'juli': 6, 'augustus': 7,
            // French
            'janvier': 0, 'février': 1, 'fevrier': 1, 'mars': 2, 'avril': 3, 'juin': 5, 'juillet': 6, 'août': 7, 'aout': 7, 'décembre': 11, 'decembre': 11,
            // Spanish
            'enero': 0, 'febrero': 1, 'marzo': 2, 'abril': 3, 'mayo': 4, 'junio': 5, 'julio': 6, 'agosto': 7, 'octubre': 9, 'diciembre': 11,
            // German
            'januar': 0, 'märz': 2, 'marz': 2, 'oktober': 9, 'dezember': 11,
            // Italian
            'gennaio': 0, 'febbraio': 1, 'aprile': 3, 'maggio': 4, 'giugno': 5, 'luglio': 6, 'settembre': 8, 'ottobre': 9, 'novembre': 10, 'dicembre': 11,
            // Polish (both nominative and genitive commonly used in dates)
            'styczeń': 0, 'stycznia': 0, 'luty': 1, 'lutego': 1, 'marzec': 2, 'marca': 2, 'kwiecień': 3, 'kwietnia': 3, 'maja': 4, 'czerwiec': 5, 'czerwca': 5, 'lipiec': 6, 'lipca': 6, 'sierpień': 7, 'sierpnia': 7, 'wrzesień': 8, 'września': 8, 'październik': 9, 'października': 9, 'listopada': 10, 'grudzień': 11, 'grudnia': 11
          };

          const s = dateStr.toLowerCase();
          const dayMatch = s.match(/\d+/);
          const monthMatch = s.match(/[a-zżźćńółęąś]+/gi) || [];
          
          let targetDate: Date | null = null;

          if (dayMatch) {
            const day = parseInt(dayMatch[0]);
            let monthIndex = -1;
            
            for (const monthName of monthMatch) {
              const cleanedMonth = monthName.toLowerCase();
              if (euMonthMap[cleanedMonth] !== undefined) {
                monthIndex = euMonthMap[cleanedMonth];
                break;
              }
            }

            if (monthIndex !== -1) {
              targetDate = new Date(currentYear, monthIndex, day);
            } else {
              // Try standard parsing
              const parsed = new Date(dateStr + ` ${currentYear}`);
              if (!isNaN(parsed.getTime())) targetDate = parsed;
            }
          }

          if (!targetDate && /morgen|tomorrow|demain|mañana|manana|jutro|domani|imorgon|heute|today|vandaag|aujourd|hoy|dzisiaj|oggi|idag/i.test(s)) {
            targetDate = new Date();
            if (/morgen|tomorrow|demain|mañana|manana|jutro|domani|imorgon/i.test(s)) {
              targetDate.setDate(targetDate.getDate() + 1);
            }
          }

          if (targetDate) {
            targetDate.setHours(0, 0, 0, 0);
            const todayClean = new Date();
            todayClean.setHours(0, 0, 0, 0);
            const diff = differenceInDays(targetDate, todayClean);
            shippingDaysNum = diff >= 0 ? diff : 0;
            shippingDaysStr = `${shippingDaysNum} Days`;
          }
        } catch (e) {
          console.error("Shipping parsing error:", e);
        }
      }

      // Price Extraction (Target: #corePriceDisplay_desktop_feature_div)
      let priceDisplay = "N/A";
      let listPrice = "N/A";
      
      const corePriceDiv = $('#corePriceDisplay_desktop_feature_div, #apex_desktop_price_feature_div');
      if (corePriceDiv.length) {
        priceDisplay = corePriceDiv.find('span.a-offscreen').first().text().trim() || 
                       corePriceDiv.find('.a-price span.a-offscreen').first().text().trim();
        // List Price / RRP
        listPrice = corePriceDiv.find('.a-price.a-text-price span.a-offscreen').first().text().trim() || 
                    $('.basisPrice .a-offscreen').first().text().trim() || 
                    $('.a-price-range .a-offscreen').first().text().trim() || "N/A";
      }
      
      if (!priceDisplay || priceDisplay === "N/A") {
        priceDisplay = $('#price_inside_buybox').text().trim() || 
                       $('#priceblock_ourprice').text().trim() ||
                       $('#priceblock_dealprice').text().trim() ||
                       $('input[name="items[0.base][customerVisiblePrice][displayString]"]').val() as string || "N/A";
      }

      const variationSelectors = [
        '#twister',
        '#inline-twister-row-size_name',
        '#inline-twister-row-color_name',
        'li.swatchAvailable',
        '#variation_size_name',
        '#variation_color_name',
        '#variation_style_name',
        '#twister-main-container',
        'select#native_dropdown_selected_size_name',
        '.twister-style-2',
        '#variation_edition_name',
        '#variation_pattern_name',
        '#variation_scent_name',
        '#variation_item_package_quantity',
        '#variation_flavor_name',
        '#variation_material_name',
        '.twister-container',
        '[id^="variation_"]',
        '.a-section.twister-row',
        '#native_dropdown_selected_color_name',
        '#native_dropdown_selected_style_name',
        '.swatchSelect',
        '.swatchAvailable',
        '.a-declarative[data-action="a-modal"]', // Often used for size charts/variations
        '[data-variation-name]',
        '.twister-row',
        '#variation_type',
        '#variation_name',
        '.variation-group'
      ];
      let hasVariations = false;
      for (const selector of variationSelectors) {
        if ($(selector).length > 0) {
          hasVariations = true;
          break;
        }
      }
      
      // Fallback: check for any element with "variation" in ID or class inside a twister container
      if (!hasVariations) {
        if ($('[id*="variation"], [class*="variation"]').length > 0 && ($('#twister-main-container, #twister').length > 0 || $('.twister-row').length > 0)) {
          hasVariations = true;
        }
      }
      
      // Additional check for labels which often indicate variations across all categories
      if (!hasVariations) {
        const twisterText = $('#twister-main-container, #twister, #variation_type, #variation_name').text().toLowerCase();
        const variationKeywords = [
          // Basic
          'size:', 'color:', 'style:', 'rozmiar:', 'kolor:', 'styl:', 'grösse:', 'farbe:', 'storlek:', 'färg:',
          'edition:', 'pattern:', 'scent:', 'flavor:', 'material:', 'configuration:', 'capacity:', 'length:',
          'width:', 'height:', 'weight:', 'volume:', 'quantity:', 'package:', 'type:', 'model:', 'version:',
          'platform:', 'format:', 'design:', 'finish:', 'voltage:', 'wattage:', 'amperage:', 'speed:',
          'memory:', 'storage:', 'display:', 'screen:', 'resolution:', 'connectivity:', 'interface:',
          'compatibility:', 'os:', 'language:', 'region:', 'country:', 'age:', 'gender:',
          // Extended Physical
          'thickness:', 'diameter:', 'depth:', 'fabric:', 'material_type:', 'bundle:', 'set:', 'pack:', 'count:',
          'thickness:', 'diameter:', 'depth:', 'item_dimensions:', 'item_weight:', 'item_form:',
          // Tech & Electronics
          'hardware_platform:', 'software_edition:', 'connectivity_technology:', 'wireless_communication_technology:',
          'power_source:', 'battery_type:', 'memory_storage_capacity:', 'display_size:', 'screen_size:',
          // Apparel & Accessories
          'shoe_size:', 'apparel_size:', 'lens_color:', 'frame_color:', 'lens_width:', 'hand_orientation:',
          // Sports & Outdoors
          'shaft_material:', 'flex:', 'grip_size:', 'tension:', 'weight_class:',
          // Beauty & Health
          'scent_description:', 'item_form:', 'unit_count:', 'number_of_items:', 'number_of_pieces:',
          // Localized Polish
          'grubość:', 'średnica:', 'głębokość:', 'tkanina:', 'zestaw:', 'opakowanie:', 'liczba:', 'częstotliwość:',
          'gwarancja:', 'platforma:', 'wersja:', 'rozdzielczość:', 'kompatybilność:', 'język:', 'kraj:', 'wiek:', 'płeć:',
          // Localized German
          'dicke:', 'durchmesser:', 'tiefe:', 'stoff:', 'set:', 'packung:', 'anzahl:', 'frequenz:', 'garantie:',
          'plattform:', 'version:', 'auflösung:', 'kompatibilität:', 'sprache:', 'land:', 'alter:', 'geschlecht:',
          // Localized Swedish
          'tjocklek:', 'diameter:', 'djup:', 'tyg:', 'set:', 'förpackning:', 'antal:', 'frekvens:', 'garanti:',
          'plattform:', 'version:', 'upplösning:', 'kompatibilitet:', 'språk:', 'land:', 'ålder:', 'kön:',
          // Internal Amazon Names
          'size_name', 'color_name', 'style_name', 'item_package_quantity', 'unit_count', 'customer_defined_variation',
          'scent_name', 'flavor_name', 'material_name', 'pattern_name', 'edition_name'
        ];
        if (variationKeywords.some(k => twisterText.includes(k))) {
          hasVariations = true;
        }
      }
      
      // Final check: look for any list or select inside twister that has more than one option
      if (!hasVariations) {
        const twisterOptions = $('#twister-main-container select option, #twister-main-container ul li').length;
        if (twisterOptions > 1) {
          hasVariations = true;
        }
      }

      // Amazon Bullet Points (Refined to prevent duplication and distortion)
      const amazonBullets: string[] = [];
      const primaryBulletSelectors = [
        '#feature-bullets ul li span.a-list-item',
        '#featurebullets_feature_div ul li span.a-list-item',
        '[data-feature-name="featurebullets"] ul li span.a-list-item',
        '#productFactsDesktopExpander ul li span.a-list-item',
        '.a-unordered-list.a-vertical li span.a-list-item'
      ];

      primaryBulletSelectors.forEach(s => {
        $(s).each((_, el) => {
          const text = $(el).text().trim();
          if (text.length > 2 && 
              !text.includes('See more photos') && 
              !text.includes('Check details') && 
              !text.includes('See more') &&
              !amazonBullets.includes(text)) {
            amazonBullets.push(text);
          }
        });
      });

      console.log(`Extracted ${amazonBullets.length} Amazon Bullets`);

      const liveData = {
        title: ($('#productTitle').text() || $('span[id="productTitle"]').text()).trim(),
        description: amazonDesc,
        bullets: amazonBullets,
        price: priceDisplay,
        listPrice: listPrice,
        currency: "", // Removed currency validation
        shipping: shippingDaysStr,
        rawShipping: rawShippingTime || "N/A",
        variations: hasVariations ? 1 : 0,
        hasAPlus: hasAPlus,
        images: uniqueImages
      };

      const auditResult = performAudit(masterData, liveData, 'amazon', domain);
      res.json({ liveData, auditResult });
    } catch (error: any) {
      console.error("Amazon Audit Error:", error);
      res.status(500).json({ 
        error: error.message || "An unexpected error occurred during Amazon audit",
        details: error.stack,
        name: error.name
      });
    } finally {
      if (browser) await browser.close();
    }
  });

  // 3. Audit Bol.com
  app.post("/api/audit/bol", async (req, res) => {
    let browser;
    try {
      const { ean, masterData } = req.body;
      const searchUrl = `https://www.bol.com/nl/nl/s/?searchtext=${ean}`;
      
      console.log(`Starting Bol Audit for EAN: ${ean}`);
      
      const proxyServer = process.env.PROXY_SERVER;
      const proxyUsername = process.env.PROXY_USERNAME;
      const proxyPassword = process.env.PROXY_PASSWORD;
      
      const launchOptions: any = {
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-gpu', '--single-process']
      };

      if (proxyServer) {
        launchOptions.proxy = {
          server: proxyServer,
          username: proxyUsername || undefined,
          password: proxyPassword || undefined,
        };
      }

      browser = await chromium.launch(launchOptions).catch(err => {
        console.error("BOL AUDIT FAILED TO LAUNCH CHROMIUM:", err);
        throw new Error(`Bol Audit: Browser launch failed. Ensure system dependencies are installed. Original error: ${err.message}`);
      });
      const context = await browser.newContext({
        userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        viewport: { width: 1920, height: 1080 },
        extraHTTPHeaders: {
            'Accept-Language': 'en-US,en;q=0.9,nl;q=0.8',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
            'sec-ch-ua': '"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'Upgrade-Insecure-Requests': '1',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'none',
            'Sec-Fetch-User': '?1',
            'Referer': 'https://www.google.com/'
        }
      });

      // Set cookies for language and country
      await context.addCookies([
        { name: 'cookie_consent', value: 'true', domain: '.bol.com', path: '/' },
        { name: 'bol_gdpr_consent', value: 'yes', domain: '.bol.com', path: '/' },
        { name: 'nl_NL', value: 'true', domain: '.bol.com', path: '/' },
        { name: 'language', value: 'nl-NL', domain: '.bol.com', path: '/' },
        { name: 'country', value: 'NL', domain: '.bol.com', path: '/' },
        { name: 'bol_p_incognito', value: 'true', domain: '.bol.com', path: '/' }
      ]);

      const page = await context.newPage();
      
      // Navigate to search results
      console.log(`Navigating to: ${searchUrl}`);
      try {
        await page.goto(searchUrl, { waitUntil: 'networkidle', timeout: 45000 });
      } catch (e: any) {
        if (e.name === 'TimeoutError') {
          console.warn("Bol search navigation timed out, attempting to proceed...");
        } else {
          throw e;
        }
      }
      
      // Check for CAPTCHA or blocking
      const isBlocked = await page.evaluate(function() {
        // @ts-ignore
        if (typeof __name === 'undefined') { (window as any).__name = (t: any, v: any) => t; }
        const text = document.body.innerText;
        return document.title.includes('Robot Check') || 
               text.includes('Type the characters you see in this image') ||
               text.includes('To discuss automated access to Amazon data please contact') ||
               text.includes('Ben je een robot?') ||
               text.includes('rustig aan speed racer') ||
               text.includes('Je gaat iets te snel') ||
               text.includes('is geblokkeerd');
      });

      if (isBlocked) {
        throw new Error("WAF_BLOCKED: Bol.com blocked the request. IP address is blocked by their anti-bot system. Running on cloud providers like Railway requires a residential proxy.");
      }

      // Check if we are on a search page or product page
      let pageTitle = await page.title();
      console.log(`Page Title: ${pageTitle}`);
      
      const isSearchPage = pageTitle.includes('Alle artikelen') || 
                           pageTitle.includes('Zoekresultaten') || 
                           await page.evaluate(function() {
                             // @ts-ignore
                             if (typeof __name === 'undefined') { (window as any).__name = (t: any, v: any) => t; }
                             return !!document.querySelector('.product-list, .product-item--grid, [data-test="product-list"]');
                           });
      
      if (isSearchPage) {
        console.log("Search page detected, looking for product link...");
        // Try to find the first product link - more robust selector
        const firstProductSelector = 'a[href*="/p/"]:not([href*="javascript"]), .product-title a, a[data-test="product-title"], .product-item--grid a, .product-list a';
        await page.waitForSelector(firstProductSelector, { timeout: 10000 }).catch(() => null);
        const firstProductLink = await page.getAttribute(firstProductSelector, 'href').catch(() => null);
        
        if (firstProductLink) {
          const productUrl = firstProductLink.startsWith('http') ? firstProductLink : `https://www.bol.com${firstProductLink}`;
          console.log(`Navigating to product URL: ${productUrl}`);
          try {
            await page.goto(productUrl, { waitUntil: 'domcontentloaded', timeout: 45000 });
          } catch (e: any) {
            if (e.name === 'TimeoutError') {
              console.warn("Bol product navigation timed out, attempting to proceed...");
            } else {
              throw e;
            }
          }
        } else {
          console.warn("No product link found on search page.");
          // Fallback: try to find any link that looks like a product link
          const anyProductLink = await page.evaluate(function() {
            // @ts-ignore
            if (typeof __name === 'undefined') { (window as any).__name = (t: any, v: any) => t; }
            const links = Array.from(document.querySelectorAll('a'));
            const productLink = links.find(a => a.href.includes('/nl/nl/p/') && !a.href.includes('javascript'));
            return productLink ? productLink.getAttribute('href') : null;
          });
          
          if (anyProductLink) {
            const productUrl = anyProductLink.startsWith('http') ? anyProductLink : `https://www.bol.com${anyProductLink}`;
            console.log(`Fallback: Navigating to product URL: ${productUrl}`);
            try {
              await page.goto(productUrl, { waitUntil: 'domcontentloaded', timeout: 45000 });
            } catch (e: any) {
              if (e.name === 'TimeoutError') {
                console.warn("Bol fallback product navigation timed out, attempting to proceed...");
              } else {
                throw e;
              }
            }
          }
        }
      }

      // Ensure we are on a product page
      await page.waitForSelector('div#pdp_main_section, [data-test="title"], h1.page-title, #buyBlockSlot', { timeout: 30000 }).catch(() => {
        console.warn("Product indicators not found, page might be slow or not a product page.");
      });

      // Wait for network idle to ensure hydration
      await page.waitForLoadState('load', { timeout: 15000 }).catch(() => null);
      await page.waitForLoadState('networkidle', { timeout: 15000 }).catch(() => {
        console.warn("Network idle timeout, proceeding with current state.");
      });

      // Extra wait for dynamic content (variations, etc.)
      await page.waitForTimeout(3000);

      // Scroll to media container to trigger lazy loading
      await page.evaluate(() => {
        const media = document.querySelector('.js_product_media_items') || 
                      document.querySelector('.pdp-images') || 
                      document.querySelector('.buy-block');
        if (media) {
          media.scrollIntoView();
          // Scroll a bit more to ensure thumbnails are triggered
          window.scrollBy(0, 300);
        }
      });
      await page.waitForTimeout(2000);

      // Wait for media container specifically
      await page.waitForSelector('.js_product_media_items, .pdp-images, [data-test="product-main-image"]', { timeout: 10000 }).catch(() => null);

      const liveDataRaw = await page.evaluate(function() {
        // @ts-ignore
        if (typeof __name === 'undefined') { (window as any).__name = (t: any, v: any) => t; }
        
        // 1. Title
        const elTitle = document.querySelector('[data-test="title"]') || document.querySelector('h1.page-title') || document.querySelector('h1');
        const title = elTitle ? (elTitle as any).innerText.trim() : "";
        
        // 2. Description
        let description = "";
        
        // Prioritize "Productbeschrijving" heading extraction
        const headings = Array.from(document.querySelectorAll('h2, h3, h4, b, strong, span'));
        const descHeading = headings.find(h => {
          const t = (h as any).innerText || (h as any).textContent || "";
          return t.includes('Productbeschrijving') || t.includes('Product description');
        });
        
        if (descHeading) {
           const parent = descHeading.closest('section') || descHeading.parentElement;
           if (parent) {
              const clone = parent.cloneNode(true) as HTMLElement;
              const UI_SELECTORS = ['.js_description_read_more', '[data-test="read-more"]', '.pdp-description__read-more', 'button', 'a.button--link', '.u-visually-hidden'];
              UI_SELECTORS.forEach(sel => {
                clone.querySelectorAll(sel).forEach(btn => btn.remove());
              });
              const text = (clone as any).innerText || (clone as any).textContent || "";
              description = text.replace(/Productbeschrijving|Product description/i, '').trim();
           }
        }

        if (!description || description.length < 50) {
          const descSelectors = [
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
            '.manufacturer-info', // Bol A+ equivalent
            '.product-info',      // Bol A+ equivalent
            '[data-test="product-info"]',
            '.js_product_info'
          ];
          
          // Try to expand "Lees meer" if it exists
          const readMore = document.querySelector('.js_description_read_more, [data-test="read-more"], .pdp-description__read-more') as HTMLElement;
          if (readMore && typeof readMore.click === 'function') {
            try { readMore.click(); } catch(e){}
          }

          let pooledText = "";
          for (const s of descSelectors) {
            const el = document.querySelector(s);
            if (el) {
               const clone = el.cloneNode(true) as HTMLElement;
               const UI_SELECTORS = ['.js_description_read_more', '[data-test="read-more"]', '.pdp-description__read-more', 'button', 'a.button--link', '.u-visually-hidden'];
               UI_SELECTORS.forEach(sel => {
                 clone.querySelectorAll(sel).forEach(btn => btn.remove());
               });
               const text = (clone as any).innerText || (clone as any).textContent || "";
               if (text.trim().length > 20) {
                  let html = clone.innerHTML;
                  let t = html.replace(/<br\s*\/?>/gi, '\n').replace(/<\/p>/gi, '\n').replace(/<\/div>/gi, '\n').replace(/<[^>]*>?/gm, ' ').replace(/\s+/g, ' ').trim();
                  
                  if (!pooledText.includes(t)) {
                    pooledText += (pooledText ? "\n\n" : "") + t;
                  }
               }
            }
          }
          if (pooledText) description = pooledText;
        }

        // Global clean for generic UI fragments and diacritics normalization (handled by server-side cleanText, but good to prune here too)
        description = description.replace(/toon meer|toon minder/gi, '').trim();

        // 3. Price
        let price = "N/A";
        const promoEl = document.querySelector('#buyBlockSlot [class*="text-promo-text-high"]') || 
                        document.querySelector('.promo-price') || 
                        document.querySelector('[data-test="price"]') ||
                        document.querySelector('.buy-block__price');
        if (promoEl) {
           price = (promoEl as any).innerText.replace(/\s+/g, '').replace('€', '').trim().replace(',', '.');
           // Ensure it's a valid price format
           if (!/^\d+\.\d+$/.test(price) && price.includes('.')) {
              // already okay
           } else if (/^\d+$/.test(price)) {
              // missing cents? Bol usually has ,- for round numbers
           }
        } else {
          const buyBlock = document.querySelector('#buyBlockSlot') || document.querySelector('.buy-block') || document.querySelector('[data-test="buy-block"]');
          if (buyBlock) {
             const text = (buyBlock as any).innerText;
             const match = text.match(/(\d+),(\d{2})/) || text.match(/(\d+),-/);
             if (match) price = match[2] ? `${match[1]}.${match[2]}` : `${match[1]}.00`;
          }
          if (price === "N/A") {
            const meta = document.querySelector('meta[property="product:price:amount"]') || document.querySelector('meta[itemprop="price"]');
            if (meta) price = meta.getAttribute('content') || "N/A";
          }
        }

        // 4. Shipping
        let shippingText = "N/A";
        // Look for "Uiterlijk [Date] in huis" or "Morgen in huis"
        const deliveryPromise = document.querySelector('.delivery-promise') || 
                                document.querySelector('[data-test="delivery-highlight"]') || 
                                document.querySelector('.delivery-delivery-time') || 
                                document.querySelector('[data-test="delivery-promise"]') ||
                                document.querySelector('.buy-block__delivery-text');
                                
        if (deliveryPromise) {
          shippingText = (deliveryPromise as any).innerText.trim();
        } else {
          const bodyText = document.body.innerText;
          const uiterlijkMatch = bodyText.match(/Uiterlijk\s+(.+?)\s+in\s+huis/i);
          if (uiterlijkMatch) shippingText = uiterlijkMatch[0];
        }

        // 5. Images (Extracting main and secondary images based on alt text)
        let images: string[] = [];
        
        document.querySelectorAll('img[alt^="Afbeelding nummer"]').forEach(img => {
          const src = (img as any).src || img.getAttribute('data-src') || img.getAttribute('src');
          if (src && src.includes('media.s-bol.com')) {
            // Normalize to large version if applicable
            const largeSrc = src.replace(/\/\d+x\d+\//, "/large/")
                               .replace("/small/", "/large/")
                               .replace("/slot/", "/large/")
                               .replace("/thumb/", "/large/")
                               .replace("/100x100/", "/large/")
                               .replace("/124x124/", "/large/")
                               .replace("/140x140/", "/large/")
                               .replace("/210x210/", "/large/")
                               .replace("/300x300/", "/large/")
                               .replace("/40x40/", "/large/");
            if (!images.includes(largeSrc)) images.push(largeSrc);
          }
        });

        // Fallback or broader search
        if (images.length === 0) {
          const thumbnailSelectors = [
            '[data-test="product-images-thumbnail"] img',
            '.js_product_img_carousel img',
            '.pdp-images__thumbnail img'
          ];

          thumbnailSelectors.forEach(s => {
            document.querySelectorAll(s).forEach(img => {
              // Ensure we are inside a thumbnail container and NOT in an A+ or recommended section
              const isInsideAplus = !!img.closest('.manufacturer-info, .product-info, [data-test="product-info"], .js_product_info');
              const isInsideRecommended = !!img.closest('.recommendations, .ux-selection-list, [data-test="recommendations"]');
              
              if (!isInsideAplus && !isInsideRecommended) {
                const src = (img as any).src || img.getAttribute('data-src') || img.getAttribute('src');
                if (src && src.includes('media.s-bol.com')) {
                  const largeSrc = src.replace(/\/\d+x\d+\//, "/large/")
                                     .replace("/small/", "/large/")
                                     .replace("/slot/", "/large/")
                                     .replace("/thumb/", "/large/")
                                     .replace("/100x100/", "/large/")
                                     .replace("/124x124/", "/large/")
                                     .replace("/140x140/", "/large/")
                                     .replace("/210x210/", "/large/")
                                     .replace("/40x40/", "/large/");
                  if (!images.includes(largeSrc)) images.push(largeSrc);
                }
              }
            });
          });
        }

        // Fallback for main image IF AND ONLY IF no thumbnails were found (safety net)
        if (images.length === 0) {
          const mainImg = document.querySelector('[data-test="product-main-image"] img') || document.querySelector('.js_main_product_image');
          if (mainImg) {
            const src = (mainImg as any).src || mainImg.getAttribute('src');
            if (src && src.startsWith('http') && !src.includes('pixel')) images.push(src);
          }
        }
        
        // 6. Variations & A+
        const hasVariations = !!(
          document.querySelector('[data-test="variant-selector"]') || 
          document.querySelector('.variant-selector') || 
          document.querySelector('.js_bundle_as_variant_selector') ||
          document.querySelector('[data-test="variant-dropdown"]') ||
          document.querySelector('.js_variant_selector') ||
          document.querySelector('[data-test="variant-list"]') ||
          document.querySelector('.pdp-variant-selector') ||
          document.querySelector('.js_variant_dropdown') ||
          // Broad pattern matching for variation themes
          Array.from(document.querySelectorAll('span, div, h3, label')).some(el => {
            const t = (el as any).innerText || "";
            const themes = [
              'Kies je kleur', 
              'Kies je maat', 
              'Kies je variant', 
              'Kies je bestanddeel', 
              'Kies je model',
              'Kies je type',
              'Kies je uitvoering',
              'Kies je breedte',
              'Kies je lengte',
              'Kies je inhoud',
              'Kies je gewicht',
              'Kies je verpakking',
              'Kies je aantal',
              'Kies je materiaal',
              'Kies je geur',
              'Kies je smaak',
              'Kies je stijl',
              'Kies je set',
              'Kies je platform'
            ];
            return themes.some(theme => t.includes(theme));
          })
        );
        
        const hasAPlus = !!(
          Array.from(document.querySelectorAll('.js_product_description img, .product-description img, .product-description-content img, [data-test="product-info"] img, .js_product_info img, .manufacturer-info img, .product-info img'))
            .some(img => {
              const parent = img.closest('section, div');
              if (parent) {
                const text = (parent as any).innerText || "";
                if (text.includes('Over het merk') || text.includes('Brand Story')) return false;
              }
              return true;
            })
        );

        // 6. Bullet Points / Kenmerken for Bol
        let bullets: string[] = [];
        const featureSelectors = [
          '[data-test="product-features"] li',
          '.product-features li',
          '.js_product_features li',
          '.specs-list__item',
          '.product-specifications li'
        ];
        
        const extractedSet = new Set<string>();
        for (const s of featureSelectors) {
          const items = document.querySelectorAll(s);
          if (items.length > 0) {
            Array.from(items).forEach(li => {
              const t = (li as any).innerText.trim();
              if (t.length > 3) extractedSet.add(t);
            });
          }
        }
        bullets = Array.from(extractedSet);

        return {
          title,
          description,
          price,
          shipping: shippingText,
          images,
          bullets,
          variations: hasVariations ? 1 : 0,
          hasAPlus: hasAPlus ? 1 : 0
        };
      });

      console.log(`Extracted Title: ${liveDataRaw.title}`);
      console.log(`Extracted Price: ${liveDataRaw.price}`);
      console.log(`Has Variations: ${liveDataRaw.variations}`);
      console.log(`Has A+ Content: ${liveDataRaw.hasAPlus}`);
      console.log(`Description Length: ${liveDataRaw.description.length}`);
      console.log(`Images Found: ${liveDataRaw.images?.length || 0}`);

      // Parse shipping days
      let shippingDaysStr = "N/A";
      let shippingDaysNum = 0;
      if (liveDataRaw.shipping && liveDataRaw.shipping !== "N/A") {
        try {
          const s = liveDataRaw.shipping.toLowerCase();
          const today = new Date();
          const currentYear = today.getFullYear();
          today.setHours(0, 0, 0, 0);

          if (s.includes('vandaag') || s.includes('today')) {
            shippingDaysNum = 0;
            shippingDaysStr = "0 Days";
          } else if (s.includes('morgen') || s.includes('tomorrow')) {
            shippingDaysNum = 1;
            shippingDaysStr = "1 Day";
          } else {
            const nlMonthMap: Record<string, number> = {
              'januari': 0, 'februari': 1, 'maart': 2, 'april': 3, 'mei': 4, 'juni': 5,
              'juli': 6, 'augustus': 7, 'september': 8, 'oktober': 9, 'november': 10, 'december': 11
            };
            
            const dayMatch = s.match(/\d+/);
            const monthMatch = s.match(/[a-z]+/g) || [];
            
            let targetDate: Date | null = null;
            if (dayMatch) {
              const day = parseInt(dayMatch[0]);
              let monthIndex = -1;
              for (const monthName of monthMatch) {
                if (nlMonthMap[monthName] !== undefined) {
                  monthIndex = nlMonthMap[monthName];
                  break;
                }
              }
              if (monthIndex !== -1) {
                targetDate = new Date(currentYear, monthIndex, day);
              }
            }
            
            if (targetDate) {
              targetDate.setHours(0, 0, 0, 0);
              const todayClean = new Date();
              todayClean.setHours(0, 0, 0, 0);
              const diff = differenceInDays(targetDate, todayClean);
              shippingDaysNum = diff >= 0 ? diff : 0;
              shippingDaysStr = `${shippingDaysNum} Day${shippingDaysNum === 1 ? '' : 's'}`;
            }
          }
        } catch (e) {}
      }

      const liveData = {
        title: liveDataRaw.title,
        description: liveDataRaw.description || "No description on detail page",
        bullets: liveDataRaw.bullets || [],
        price: liveDataRaw.price,
        shipping: shippingDaysStr,
        rawShipping: liveDataRaw.shipping || "N/A",
        variations: liveDataRaw.variations,
        hasAPlus: liveDataRaw.hasAPlus,
        images: liveDataRaw.images && liveDataRaw.images.length > 0 ? getUniqueImages(liveDataRaw.images, 'bol') : []
      };

      const auditResult = performAudit(masterData, liveData, 'bol');
      res.json({ liveData, auditResult });
    } catch (error: any) {
      console.error("Bol Audit Error:", error);
      res.status(500).json({ 
        error: error.message || "An unexpected error occurred during Bol.com audit",
        details: error.stack,
        name: error.name
      });
    } finally {
      if (browser) await browser.close();
    }
  });

  function prepareRowData(mode: string, auditResult: any, masterData: any, marketplace: string, identifier: string) {
    const getMatchText = (isMatch: boolean) => isMatch ? 'Yes' : 'No';
    const getAPlusText = (hasAPlus: boolean) => hasAPlus ? 'Available' : 'Not Available';
    const getVariationText = (exists: boolean) => exists ? 'Yes' : 'No';
    
    const masterFirst = masterData.images?.[0] || "";
    const liveFirst = auditResult.images?.live?.[0] || "";
    const allLiveImages = (auditResult.images?.live || []).join(", ");
    
    const descMatch = (mode === 'amazon' && auditResult.hasAPlus.live && !auditResult.description.match) 
      ? "A+ content available" 
      : getMatchText(auditResult.description.match);

    const sharedData: any = {
      "Identifier": identifier,
      "SKU": identifier,
      "Title match": getMatchText(auditResult.title.match),
      "Title Match": getMatchText(auditResult.title.match),
      "Description match": descMatch,
      "Description Match": descMatch,
      "Main Image Link": masterFirst,
      "Live Image Links": allLiveImages,
      "Main Image": masterFirst ? `=IMAGE("${masterFirst}")` : "",
      "Live Image": liveFirst ? `=IMAGE("${liveFirst}")` : "",
      
      // Specific requested mappings Target: "Main Live Image", "Image 1", "Live Image 1"
      "Main Live Image": liveFirst ? `=IMAGE("${liveFirst}")` : "",
      "Image 1": masterFirst, 
      "Live Image 1": liveFirst,
      
      "A+ Content": getAPlusText(auditResult.hasAPlus.live),
      "A+": getAPlusText(auditResult.hasAPlus.live),
      "Has A+": getAPlusText(auditResult.hasAPlus.live),
      "Shipping Time": auditResult.shipping.live || "N/A",
      "Shipping": auditResult.shipping.live || "N/A",
      "Delivery Days": auditResult.shipping.live || "N/A",
      "Price": auditResult.price.live || "N/A",
      "Variation": getVariationText(auditResult.variations.live),
      "Variations": getVariationText(auditResult.variations.live),
      "Variation Match": getVariationText(auditResult.variations.live),
      "Bullet Points match": auditResult.bullets ? getMatchText(auditResult.bullets.length > 0 && auditResult.bullets.every((b: any) => b.match)) : "N/A",
      "Bullet Points Match": auditResult.bullets ? getMatchText(auditResult.bullets.length > 0 && auditResult.bullets.every((b: any) => b.match)) : "N/A",
      "Bullets Match": auditResult.bullets ? getMatchText(auditResult.bullets.length > 0 && auditResult.bullets.every((b: any) => b.match)) : "N/A"
    };

    // Bulk Image Columns Implementation
    const prefix = mode === 'amazon' ? 'Amazon' : 'Bol';
    const masterImgs = masterData.images || [];
    const liveImgs = auditResult.images?.live || [];

    // Main image aliases
    sharedData[`${prefix} Main Image`] = masterImgs[0] ? `=IMAGE("${masterImgs[0]}")` : "";
    sharedData[`${prefix} Main Live Image`] = liveImgs[0] ? `=IMAGE("${liveImgs[0]}")` : "";

    // Master Images 1-10
    for (let i = 1; i <= 10; i++) {
      const url = masterImgs[i-1] || "";
      sharedData[`${prefix} Master Image ${i}`] = url ? `=IMAGE("${url}")` : "";
    }

    // Live Images 1-10
    for (let i = 1; i <= 10; i++) {
      const url = liveImgs[i-1] || "";
      sharedData[`${prefix} Live Image ${i}`] = url ? `=IMAGE("${url}")` : "";
    }

    if (mode === 'amazon') {
      const allBulletsMatch = auditResult.bullets && auditResult.bullets.length > 0 && auditResult.bullets.every((b: any) => b.match);
      return {
        "ASIN": identifier,
        "Marketplace": marketplace,
        "Bullet Points match": getMatchText(allBulletsMatch),
        ...sharedData
      };
    } else {
      return {
        "EAN": identifier,
        "Marketplace": "Bol.com",
        ...sharedData
      };
    }
  }

  // 4. Clear Target Tab Data
  app.post("/api/sheets/clear-sheet", async (req, res) => {
    try {
      const { mode } = req.body;
      const targetSheetId = "1V4lNf30SlBwczSvGX9rfn5eWFH2AvMO4TqMHAHalS7s";
      const targetTab = mode === 'amazon' ? 'Amazon QC Results' : 'Bol QC Results';
      
      const serviceAccountEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
      const privateKey = process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n');

      if (!serviceAccountEmail || !privateKey) {
        return res.status(400).json({ error: "Google Sheets credentials not configured." });
      }

      const auth = new JWT({
        email: serviceAccountEmail,
        key: privateKey,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });

      const sheets = google.sheets({ version: 'v4', auth });
      
      // Wipe protocol: Strict Clear A2:AZ1000
      await sheets.spreadsheets.values.clear({
        spreadsheetId: targetSheetId,
        range: `'${targetTab}'!A2:AZ1000`,
      });
      
      res.json({ success: true, message: `Tab "${targetTab}" cleared A2:AZ1000.` });
    } catch (error: any) {
      console.error("Clear Sheet Error:", error);
      res.status(500).json({ error: error.message });
    }
  });

  // 5. Save Audit Result to QC Automation Sheet (Singular)
  app.post("/api/sheets/save-audit", async (req, res) => {
    try {
      const { mode, auditResult, masterData, marketplace, identifier } = req.body;
      const targetSheetId = "1V4lNf30SlBwczSvGX9rfn5eWFH2AvMO4TqMHAHalS7s";
      const targetTab = mode === 'amazon' ? 'Amazon QC Results' : 'Bol QC Results';
      
      const serviceAccountEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
      const privateKey = process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n');

      if (!serviceAccountEmail || !privateKey) {
        return res.status(400).json({ error: "Google Sheets credentials not configured." });
      }

      const auth = new JWT({
        email: serviceAccountEmail,
        key: privateKey,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });

      const sheets = google.sheets({ version: 'v4', auth });
      const doc = new GoogleSpreadsheet(targetSheetId, auth);
      await doc.loadInfo();
      const sheet = doc.sheetsByTitle[targetTab];
      
      if (!sheet) {
        return res.status(404).json({ error: `Tab "${targetTab}" not found.` });
      }

      await sheet.loadHeaderRow();
      const headers = sheet.headerValues;
      const rowData = prepareRowData(mode, auditResult, masterData, marketplace, identifier);
      
      // Map to header order for batch update
      const rowArray = headers.map(h => {
        const key = Object.keys(rowData).find(k => k.toLowerCase() === h.toLowerCase());
        return key ? rowData[key] : "";
      });

      // 1. Wipe protocol: Clear A2:AZ1000
      await sheets.spreadsheets.values.clear({
        spreadsheetId: targetSheetId,
        range: `'${targetTab}'!A2:AZ1000`,
      });

      // 2. Overwrite logic: Start writing at A2
      await sheets.spreadsheets.values.update({
        spreadsheetId: targetSheetId,
        range: `'${targetTab}'!A2`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [rowArray]
        }
      });
      
      res.json({ success: true, tab: targetTab, overrode: true });
    } catch (error: any) {
      console.error("Save Audit Error:", error);
      res.status(500).json({ error: error.message });
    }
  });

  // 6. Batch Save Audit Results
  app.post("/api/sheets/batch-save-audit", async (req, res) => {
    try {
      const { mode, audits } = req.body; // audits is array of { auditResult, masterData, marketplace, identifier }
      const targetSheetId = "1V4lNf30SlBwczSvGX9rfn5eWFH2AvMO4TqMHAHalS7s";
      const targetTab = mode === 'amazon' ? 'Amazon QC Results' : 'Bol QC Results';
      
      const serviceAccountEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
      const privateKey = process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n');

      if (!serviceAccountEmail || !privateKey) {
        return res.status(400).json({ error: "Google Sheets credentials not configured." });
      }

      const auth = new JWT({
        email: serviceAccountEmail,
        key: privateKey,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });

      const sheets = google.sheets({ version: 'v4', auth });
      const doc = new GoogleSpreadsheet(targetSheetId, auth);
      await doc.loadInfo();
      const sheet = doc.sheetsByTitle[targetTab];
      
      if (!sheet) {
        return res.status(404).json({ error: `Tab "${targetTab}" not found.` });
      }

      await sheet.loadHeaderRow();
      const headers = sheet.headerValues;
      const allRows = audits.map((a: any) => {
        const rowData = prepareRowData(mode, a.auditResult, a.masterData, a.marketplace, a.identifier);
        return headers.map(h => {
          const key = Object.keys(rowData).find(k => k.toLowerCase() === h.toLowerCase());
          return key ? rowData[key] : "";
        });
      });

      // 1. Wipe protocol: Clear A2:AZ1000
      await sheets.spreadsheets.values.clear({
        spreadsheetId: targetSheetId,
        range: `'${targetTab}'!A2:AZ1000`,
      });

      // 2. Overwrite logic: Start writing at A2
      await sheets.spreadsheets.values.update({
        spreadsheetId: targetSheetId,
        range: `'${targetTab}'!A2`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: allRows
        }
      });
      
      res.json({ success: true, count: allRows.length });
    } catch (error: any) {
      console.error("Batch Save Error:", error);
      res.status(500).json({ error: error.message });
    }
  });

  // 7. Image Proxy for Google Drive/UserContent
  app.get("/api/proxy-image", async (req, res) => {
    let imageUrl = req.query.url as string;
    try {
      if (!imageUrl) return res.status(400).send("URL is required");

      // Transform Google Drive "view" links to "download" links
      if (imageUrl.includes('drive.google.com')) {
        const fileIdMatch = imageUrl.match(/\/d\/([^\/]+)/);
        if (fileIdMatch) {
          imageUrl = `https://drive.google.com/uc?export=download&id=${fileIdMatch[1]}`;
        }
      }

      const response = await axios.get(imageUrl, {
        responseType: 'arraybuffer',
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
        },
        timeout: 15000
      });

      const contentType = response.headers['content-type'];
      res.setHeader('Content-Type', contentType || 'image/jpeg');
      res.setHeader('Cache-Control', 'public, max-age=86400');
      res.send(response.data);
    } catch (error: any) {
      console.error(`Proxy Error for URL [${imageUrl}]:`, error.message);
      res.status(500).send(`Failed to fetch image: ${error.message}`);
    }
  });

  // 5. Image Comparison Helper
  app.post("/api/images/compare", async (req, res) => {
    try {
      const { url1, url2 } = req.body;
      
      const img1Resp = await axios.get(url1, { responseType: 'arraybuffer' });
      const img2Resp = await axios.get(url2, { responseType: 'arraybuffer' });
      
      const hash1 = await getImageHash(Buffer.from(img1Resp.data));
      const hash2 = await getImageHash(Buffer.from(img2Resp.data));
      
      const similarity = compareHashes(hash1, hash2);
      res.json({ similarity });
    } catch (error: any) {
      res.status(500).json({ error: error.message });
    }
  });

  function cleanText(text: string | undefined): string {
    if (!text) return "";
    
    // 1. Remove HTML tags
    let cleaned = text.replace(/<[^>]*>?/gm, '');
    
    // 2. Whitespace Normalization (Replace non-breaking spaces with standard spaces)
    cleaned = cleaned.replace(/&nbsp;|\xa0|[\u00A0\u1680\u180E\u2000-\u200B\u202F\u205F\u3000\uFEFF]/g, ' ');
    
    // 3. Symbol Erasure (Temporarily ignore TM, R, C, bullets, dashes during comparison)
    // We also remove other common bullet types and symbols
    cleaned = cleaned.replace(/[™®©•\-\·\*+\u2022\u2023\u25E6\u2043\u2219]/g, '');
    
    // 4. Diacritic Normalization
    cleaned = cleaned.normalize('NFD').replace(/[\u0300-\u036f]/g, "");
    
    // 5. Remove UI Fragments
    cleaned = cleaned.replace(/toon meer|toon minder/gi, '');
    
    // 6. Collapse multiple spaces and trim
    return cleaned.replace(/\s+/g, ' ').trim().toLowerCase();
  }

  function getUniqueImages(urlList: string[], mode: 'amazon' | 'bol' = 'amazon') {
    const idToUrlMap = new Map<string, string>();
    const finalImages: string[] = [];
    
    // Filter out icons and unrelated stuff
    const filteredList = urlList.filter(url => {
      if (!url || typeof url !== 'string' || !url.startsWith('http')) return false;
      const lower = url.toLowerCase();
      // Exclude common icon/UI keywords found on Amazon and retail sites
      if (lower.includes('icon') || lower.includes('sprite') || lower.includes('zoom') || 
          lower.includes('play-button') || lower.includes('logo') || lower.includes('video-button') ||
          lower.includes('loading') || lower.includes('transparent-pixel') || lower.includes('dot-') ||
          lower.includes('swatch') || lower.includes('feedback') || lower.includes('magnify') ||
          lower.includes('video-placeholder') || lower.includes('sash') || lower.includes('grey-pixel') ||
          lower.includes('_ss40_') || lower.includes('_sr38,50_')) return false;
      return true;
    });

    for (const url of filteredList) {
      let imgId = "";
      let isAmazon = false;
      let processedUrl = url;
      
      if (url.includes('amazon.com') || url.includes('media-amazon.com')) {
        isAmazon = true;
        try {
          if (url.includes('/I/')) {
            // Extract core ID before any resolution markers like ._AC_...
            const parts = url.split('/I/');
            if (parts.length > 1) {
              imgId = parts[1].split('.')[0];
            }
          } else if (url.includes('/S/')) {
            // S images are often A+ content
            const parts = url.split('/S/');
            if (parts.length > 1) {
              imgId = 'S_' + parts[1].split('.')[0];
            }
          } else {
            const match = url.match(/\/images\/([IS])\/([A-Za-z0-9_-]+)/);
            if (match) imgId = (match[1] === 'S' ? 'S_' : '') + match[2];
          }
        } catch (e) {
          const match = url.match(/\/images\/[IS]\/([A-Za-z0-9_-]+)/);
          if (match) imgId = match[1];
        }
      } else if (url.includes('media.s-bol.com')) {
        const match = url.match(/media\.s-bol\.com\/([A-Za-z0-9_-]+)/);
        if (match) imgId = match[1];
        
        // Normalize Bol URLs to large version
        processedUrl = url.replace(/\/\d+x\d+\//, "/large/")
                          .replace("/small/", "/large/")
                          .replace("/slot/", "/large/")
                          .replace("/thumb/", "/large/")
                          .replace("/100x100/", "/large/")
                          .replace("/124x124/", "/large/")
                          .replace("/140x140/", "/large/")
                          .replace("/210x210/", "/large/")
                          .replace("/40x40/", "/large/");
      } else {
        const match = url.match(/\/([A-Za-z0-9_-]{10,20})/);
        if (match) imgId = match[1];
      }

      if (imgId) {
        if (isAmazon) {
          const existingUrl = idToUrlMap.get(imgId);
          if (!existingUrl) {
            idToUrlMap.set(imgId, url);
          } else {
            // Prioritize higher resolution
            const getRes = (u: string) => {
              if (u.includes('SL1500') || u.includes('original')) return 9999;
              const match = u.match(/[A-Z]+(\d+)/);
              return match ? parseInt(match[1]) : 0;
            };
            if (getRes(url) > getRes(existingUrl)) {
              idToUrlMap.set(imgId, url);
            }
          }
        } else {
          if (!idToUrlMap.has(imgId)) {
            idToUrlMap.set(imgId, processedUrl);
          }
        }
      } else {
        if (!finalImages.includes(url)) {
          finalImages.push(url);
        }
      }
    }

    let result = [...Array.from(idToUrlMap.values()), ...finalImages];
    
    // As per user request: skip Amazon live image 2 if it exists (usually duplicate of image 1)
    if (mode === 'amazon' && result.length >= 2) {
      result.splice(1, 1);
    }
    
    return result;
  }

  function performAudit(master: any, live: any, mode: 'amazon' | 'bol', domain?: string) {
    const results: any = {};
    
    // Title Comparison
    const cleanMasterTitle = cleanText(master.title);
    const cleanLiveTitle = cleanText(live.title);
    const titleSimilarity = stringSimilarity.compareTwoStrings(cleanMasterTitle, cleanLiveTitle);
    results.title = {
      master: master.title,
      live: live.title,
      similarity: titleSimilarity,
      match: titleSimilarity >= 0.95 || cleanMasterTitle === cleanLiveTitle
    };

    // Description Comparison
    const isImageDesc = live.description.startsWith('IMAGE:');
    const isAPlusImages = live.description.startsWith('APLUS_IMAGES:');
    const isAPlusData = live.description.startsWith('APLUS_DATA:');
    const cleanMasterDesc = cleanText(master.description);
    const cleanLiveDesc = (isImageDesc || isAPlusImages || isAPlusData) ? "" : cleanText(live.description);
    const descSimilarity = (isImageDesc || isAPlusImages || isAPlusData) ? 0.5 : stringSimilarity.compareTwoStrings(cleanMasterDesc, cleanLiveDesc);
    
    results.description = {
      master: master.description,
      live: live.description,
      similarity: descSimilarity,
      match: (isImageDesc || isAPlusImages || isAPlusData) ? false : (descSimilarity >= 0.95 || cleanMasterDesc === cleanLiveDesc),
      isImage: isImageDesc || isAPlusImages || isAPlusData,
      isAPlus: isAPlusImages || isAPlusData,
      status: ((isAPlusImages || isAPlusData) && cleanMasterDesc) ? "Manual Check Required: A+ Content Live" : null
    };

    // Currency Validation - REMOVED as per user request
    results.currency = {
      expected: "N/A",
      actual: live.currency || "N/A",
      match: true
    };

    // Bullet Points Comparison (Best-match logic)
    const liveBullets = live.bullets || [];
    results.bullets = (master.bullets || []).map((mb: string) => {
      const cmb = cleanText(mb);
      let bestMatch = { similarity: 0, text: "" };
      
      for (const lb of liveBullets) {
        const clb = cleanText(lb);
        const sim = stringSimilarity.compareTwoStrings(cmb, clb);
        if (sim > bestMatch.similarity) {
          bestMatch = { similarity: sim, text: lb };
        }
      }
      
      return {
        master: mb,
        live: bestMatch.text || "",
        similarity: bestMatch.similarity,
        match: bestMatch.similarity >= 0.90 || cmb === cleanText(bestMatch.text)
      };
    });

    results.hasAPlus = {
      master: !!master.hasAPlus,
      live: !!live.hasAPlus,
      match: !!master.hasAPlus === !!live.hasAPlus
    };

    results.variations = {
      master: !!master.variations,
      live: !!live.variations,
      match: !!master.variations === !!live.variations
    };

    results.price = {
      master: master.price,
      live: live.price,
      match: master.price == live.price
    };

    results.shipping = {
      master: master.shipping,
      live: live.shipping,
      similarity: stringSimilarity.compareTwoStrings(cleanText(master.shipping), cleanText(live.shipping)),
      match: cleanText(master.shipping) === cleanText(live.shipping)
    };

    // Image comparison results
    const masterImages = master.images || [];
    const liveImages = live.images || [];
    const masterFirst = masterImages[0] || "";
    const liveFirst = liveImages[0] || "";
    
    // Simple match logic: compare first image filename or full URL
    const getFilename = (url: string) => url.split('/').pop()?.split('?')[0] || "";
    const isImageMatch = masterFirst && liveFirst && (masterFirst === liveFirst || getFilename(masterFirst) === getFilename(liveFirst));

    results.images = {
      master: masterImages,
      live: liveImages,
      match: isImageMatch
    };

    return results;
  }

  async function getImageHash(buffer: Buffer) {
    // Basic hash: resize to 8x8 and convert to grayscale
    const { data } = await sharp(buffer)
      .resize(8, 8, { fit: 'fill' })
      .grayscale()
      .raw()
      .toBuffer({ resolveWithObject: true });
    return data;
  }

  function compareHashes(h1: Buffer, h2: Buffer) {
    let diff = 0;
    for (let i = 0; i < h1.length; i++) {
      if (Math.abs(h1[i] - h2[i]) > 20) diff++;
    }
    return 1 - (diff / h1.length);
  }

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
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

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer().catch(err => {
  console.error("Failed to start server:", err);
  process.exit(1);
});
