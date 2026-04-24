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
import { chromium as chromiumExtra } from "playwright-extra";
import stealth from "puppeteer-extra-plugin-stealth";

// Configure playwright-extra with stealth plugin — used for ALL scraping
// chromiumExtra.use(stealth()); 

async function startServer() {
  const app = express();
  const PORT = process.env.PORT || 3000;

  app.use(cors());
  app.use(express.json({ limit: '50mb' }));

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

      browser = await chromiumExtra.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-blink-features=AutomationControlled']
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

      // Disable webdriver flag before navigation
      await page.addInitScript(() => {
        Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
      });

      // Block only fonts/tracking — keep images/CSS so data-a-dynamic-image populates
      await page.route('**/*.{woff,woff2,ttf,otf}', route => route.abort());
      await page.route('**/analytics*', route => route.abort());
      await page.route('**/tracking*', route => route.abort());

      // Navigate and wait for DOM + AJAX pricing to load
      try {
        await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
        // Wait for networkidle so AJAX buybox price loads
        await page.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => null);
      } catch (e: any) {
        if (e.name === 'TimeoutError') {
          console.warn("Amazon navigation timed out, proceeding with current content...");
        } else {
          throw e;
        }
      }

      // Ensure the product title is at least present
      await page.waitForSelector('#productTitle', { timeout: 15000 }).catch(() => {
        console.warn("Product title not found within 15s, page might be slow or blocked.");
      });

      // Check for CAPTCHA
      const isCaptcha = await page.evaluate(function () {
        // @ts-ignore
        if (typeof __name === 'undefined') { (window as any).__name = (t: any, v: any) => t; }
        return document.title.includes('Robot Check') ||
          document.body.innerText.includes('Type the characters you see in this image') ||
          document.body.innerText.includes('To discuss automated access to Amazon data please contact');
      });

      if (isCaptcha) {
        throw new Error("Amazon blocked the request with a CAPTCHA. Please try again later.");
      }

      // Get content immediately after domcontentloaded/buybox wait
      let content = await page.content();
      let $ = cheerio.load(content);

      // --- PHASE 1: TOP-PAGE DATA (Price & Shipping) ---

      // Price Extraction
      let priceDisplay = "N/A";
      let listPrice = "N/A";

      const priceSelectors = [
        '#corePriceDisplay_desktop_feature_div .a-offscreen',
        '#corePrice_feature_div .a-offscreen',
        '#price_inside_buybox',
        '#priceblock_ourprice',
        '#priceblock_dealprice',
        'span.a-price span.a-offscreen',
        '.a-price.a-text-price .a-offscreen'
      ];

      for (const selector of priceSelectors) {
        const val = $(selector).first().text().trim();
        if (val && val !== "N/A" && val.length < 20) { // Limit length to avoid title capture
          priceDisplay = val;
          break;
        }
      }

      // Shipping Extraction (Prioritize mir-layout-DELIVERY_BLOCK)
      const deliverySelectors = [
        '#mir-layout-DELIVERY_BLOCK-slot-PRIMARY_DELIVERY_MESSAGE_LARGE',
        '#mir-layout-DELIVERY_BLOCK-slot-SECONDARY_DELIVERY_MESSAGE_LARGE',
        '#mir-layout-DELIVERY_BLOCK',
        '#deliveryBlockMessage',
        '#pfd-desktop-PRIMARY_DELIVERY_MESSAGE_LARGE'
      ];

      let rawShippingTime = "N/A";
      for (const selector of deliverySelectors) {
        const el = $(selector);
        if (el.length) {
          rawShippingTime = el.find('[data-csa-c-delivery-time]').attr('data-csa-c-delivery-time') ||
            el.attr('data-csa-c-delivery-time') ||
            el.find('.a-text-bold').first().text().trim() ||
            el.text().trim();
          if (rawShippingTime && rawShippingTime.length > 3) break;
        }
      }

      console.log(`✅ Top-page data secured: Price=${priceDisplay}, Shipping=${rawShippingTime}`);

      // --- PHASE 2: LOWER-PAGE DATA (A+, Images, Description) ---

      // Human-mimicry: Scroll down to trigger lazy-loaded A+ and modules
      console.log("Scrolling for A+ content and images...");
      await page.mouse.wheel(0, 800);
      await page.waitForTimeout(3000);

      // Re-load content for hydrated elements
      content = await page.content();
      $ = cheerio.load(content);

      // 1. Image Extraction — Strategy A: data-a-dynamic-image (highest resolution JSON)
      let images: string[] = [];
      const dynamicImgJson = $('#landingImage').attr('data-a-dynamic-image');
      if (dynamicImgJson) {
        try {
          const imgMap = JSON.parse(dynamicImgJson);
          // Sort by resolution descending, take highest-res URL
          const sorted = Object.keys(imgMap).sort((a, b) => {
            const [wa, ha] = imgMap[a]; const [wb, hb] = imgMap[b];
            return (wb * hb) - (wa * ha);
          });
          if (sorted[0]) images.push(sorted[0]);
        } catch (e) { console.warn('Failed to parse data-a-dynamic-image'); }
      }

      // Strategy B: data-old-hires on landing image
      const landingHires = $('#landingImage').attr('data-old-hires');
      if (landingHires && landingHires.startsWith('http')) images.push(landingHires);

      // Strategy C: Alt images strip sizing fragments for full-res
      $('#altImages ul li img').each((_, el) => {
        const src = $(el).attr('data-old-hires') || $(el).attr('src') || '';
        if (src.startsWith('http')) images.push(src);
      });

      let uniqueImages = getUniqueImages(images);

      // Specifically for Amazon: Skip the image at index 1 (the second image) 
      // to avoid duplication with index 0 (main image).
      if (uniqueImages.length > 1) {
        console.log("Amazon extraction: Skipping image at index 1 to avoid duplication.");
        uniqueImages = [uniqueImages[0], ...uniqueImages.slice(2)];
      }

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

      // Shipping parsing using the rawShippingTime extracted in Phase 1
      let shippingDaysNum = 0;
      let shippingDaysStr = "N/A";

      if (rawShippingTime && rawShippingTime !== "N/A") {
        try {
          const today = new Date();
          today.setHours(0, 0, 0, 0);
          const currentYear = today.getFullYear();
          let targetDate: Date | null = null;

          // Multi-language month maps
          const monthMap: Record<string, number> = {
            // English
            'january': 0, 'february': 1, 'march': 2, 'april': 3, 'may': 4, 'june': 5,
            'july': 6, 'august': 7, 'september': 8, 'october': 9, 'november': 10, 'december': 11,
            'jan': 0, 'feb': 1, 'mar': 2, 'apr': 3, 'may': 4, 'jun': 5, 'jul': 6, 'aug': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dec': 11,
            // Dutch (Amazon.nl)
            'januari': 0, 'februari': 1, 'maart': 2, 'mei': 4, 'juni': 5, 'juli': 6, 'augustus': 7, 'oktober': 9,
            // German (Amazon.de)
            'januar': 0, 'februar': 1, 'märz': 2, 'mai': 4, 'juni': 5, 'juli': 6, 'oktober': 9, 'dezember': 11,
            'jan.': 0, 'feb.': 1, 'märz': 2, 'apr.': 3, 'mai': 4, 'juni': 5, 'juli': 6, 'aug.': 7, 'sept.': 8, 'okt.': 9, 'nov.': 10, 'dez.': 11,
            // Polish (Amazon.pl)
            'stycznia': 0, 'lutego': 1, 'marca': 2, 'kwietnia': 3, 'maja': 4, 'czerwca': 5, 'lipca': 6, 'sierpnia': 7, 'września': 8, 'października': 9, 'listopada': 10, 'grudnia': 11,
            'styczni': 0, 'lut': 1, 'mar': 2, 'kwie': 3, 'maj': 4, 'czerw': 5, 'lip': 6, 'sierp': 7, 'wrze': 8, 'paź': 9, 'list': 10, 'grud': 11,
            // Swedish (Amazon.se)
            'januari': 0, 'februari': 1, 'mars': 2, 'maj': 4, 'juni': 5, 'juli': 6, 'augusti': 7, 'oktober': 9, 'november': 10, 'december': 11
          };

          const s = rawShippingTime.toLowerCase();

          // Check for "Tomorrow" in various languages
          if (/tomorrow|morgen|jutro|i morgon/i.test(s)) {
            targetDate = new Date(today);
            targetDate.setDate(today.getDate() + 1);
          } else if (/today|vandaag|heute|dzisiaj|idag/i.test(s)) {
            targetDate = new Date(today);
          } else {
            // Find day and month
            const dayMatch = s.match(/(\d+)/);
            let monthIndex = -1;

            for (const [name, index] of Object.entries(monthMap)) {
              if (s.includes(name)) {
                monthIndex = index;
                break;
              }
            }

            if (dayMatch && monthIndex !== -1) {
              const day = parseInt(dayMatch[1]);
              targetDate = new Date(currentYear, monthIndex, day);

              // Handle year rollover (e.g., today is Dec 30, and delivery is Jan 2)
              if (targetDate.getTime() < today.getTime() - (30 * 24 * 60 * 60 * 1000)) {
                targetDate.setFullYear(currentYear + 1);
              }
            }
          }

          if (targetDate) {
            targetDate.setHours(0, 0, 0, 0);
            shippingDaysNum = Math.ceil((targetDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
            if (shippingDaysNum < 0) shippingDaysNum = 0;
            shippingDaysStr = `${shippingDaysNum} Days`;
          }
        } catch (e) {
          console.error("Shipping parsing error:", e);
        }
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

      const launchOptions: any = {
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-blink-features=AutomationControlled']
      };

      browser = await chromiumExtra.launch(launchOptions);

      // --- BOL.COM MOBILE STEALTH CONTEXT ---
      const viewportWidth = Math.floor(Math.random() * (420 - 375 + 1)) + 375;
      const context = await browser.newContext({
        userAgent: 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_4 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Mobile/15E148 Safari/604.1',
        viewport: { width: viewportWidth, height: 844 },
        locale: 'nl-NL',
        extraHTTPHeaders: {
          'Accept-Language': 'nl-NL,nl;q=0.9',
          'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        }
      });

      await context.clearCookies();

      // Mask Automation
      await context.addInitScript(() => {
        // @ts-ignore
        Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
      });

      let page = await context.newPage();

      // --- FORCED BOL.COM BYPASS (Direct Link Method) ---
      const directSearchUrl = `https://www.bol.com/nl/nl/s/?searchtext=${ean}`;
      console.log(`Mobile Direct Jump: ${directSearchUrl} (Viewport: ${viewportWidth}px)`);

      async function attemptNavigation(isRetry = false) {
        try {
          await page.goto(directSearchUrl, { waitUntil: 'commit', timeout: 45000 });
          
          // Cookie Bypass
          const title = await page.title();
          if (title.includes('Privacy') || title.includes('cookies')) {
            console.log("Cookie consent detected, clicking Akkoord...");
            await page.click('button#js-accept-all-cookies, [data-test="consent-assign-all"]').catch(() => null);
            await page.waitForTimeout(2000);
          }

          // Check for Block
          const content = await page.content();
          if (content.includes("IP adres is geblokkeerd") || content.includes("rustig aan speed racer")) {
            if (!isRetry) {
              console.warn("❌ Blocked. Initiating 10-second recovery and retry...");
              await page.waitForTimeout(10000);
              // Update viewport for jitter
              const newWidth = Math.floor(Math.random() * (420 - 375 + 1)) + 375;
              await page.setViewportSize({ width: newWidth, height: 844 });
              return await attemptNavigation(true);
            }
            throw new Error("Bol.com IP Blocked after retry.");
          }

          // Wait for redirect to /p/ (Product Page)
          try {
            await page.waitForURL('**/p/**', { timeout: 8000 });
          } catch (e) {
            // Not redirected? Click first title
            const firstProduct = await page.waitForSelector('[data-test="product-title"]', { timeout: 5000 }).catch(() => null);
            if (firstProduct) {
              await firstProduct.click();
              await page.waitForURL('**/p/**', { timeout: 8000 }).catch(() => null);
            }
          }

        } catch (err: any) {
          if (!isRetry) {
             console.warn(`Navigation failed: ${err.message}. Retrying once...`);
             await page.waitForTimeout(5000);
             return await attemptNavigation(true);
          }
          throw err;
        }
      }

      await attemptNavigation();

      console.log(`✅ Reached Bol.com product page: "${await page.title()}"`);

      // Ensure we are on a product page — wait specifically for the title (2026-compatible selector)
      await page.waitForSelector('[data-test="title"], div#pdp_main_section, h1.page-title, #buyBlockSlot, #pdp_description', { timeout: 15000 }).catch(() => {
        console.warn("Product title or description not found within 15s, page might be slow or not a product page.");
      });

      // Wait for network idle strictly for PDP to ensure Price/Shipping are hydrated
      await page.waitForLoadState('load', { timeout: 45000 }).catch(() => null);
      await page.waitForLoadState('networkidle', { timeout: 45000 }).catch(() => {
        console.warn("Network idle timeout, proceeding with current state.");
      });

      // Human-Mimicry: Force scroll event to trigger lazy-loaded description and modules
      console.log("Human mimicry: scrolling PDP...");
      await page.mouse.wheel(0, 400);
      await page.waitForTimeout(2000);

      // Extra wait for dynamic content (variations, etc.)
      await page.waitForTimeout(3000);

      // Check if page loaded properly
      const pageLoaded = await page.evaluate(() => {
        return document.body.innerText && document.body.innerText.length > 100;
      });

      if (!pageLoaded) {
        throw new Error("Page content did not load properly. Possible block or network issue.");
      }

      // Add another delay to be extra safe before scrolling
      await page.waitForTimeout(2000);

      // Scroll to media container to trigger lazy loading
      await page.evaluate(() => {
        const media = document.querySelector('[data-test="product-main-image"]') ||
          document.querySelector('.js_product_media_items') ||
          document.querySelector('.pdp-images') ||
          document.querySelector('.buy-block') ||
          document.querySelector('[class*="ProductImage"]') ||
          document.querySelector('[id*="image"]');
        if (media) {
          media.scrollIntoView();
          // Scroll a bit more to ensure thumbnails are triggered
          window.scrollBy(0, 300);
        } else {
          window.scrollBy(0, 500);
        }
      });
      await page.waitForTimeout(2500);

      // Wait for media container specifically
      await page.waitForSelector('[data-test="product-main-image"], .js_product_media_items, .pdp-images, [class*="ProductImage"], img', { timeout: 10000 }).catch(() => null);

      const liveDataRaw = await page.evaluate(function () {
        // @ts-ignore
        if (typeof __name === 'undefined') { (window as any).__name = (t: any, v: any) => t; }

        // 1. Title
        let title = document.title.split('|')[0].trim();
        if (!title || title.length < 5) {
          const elTitle = document.querySelector('[data-test="title"]') || document.querySelector('h1.page-title') || document.querySelector('h1');
          title = elTitle ? (elTitle as any).innerText.trim() : "";
        }

        // 2. Description
        let description = "";

        // ... (rest of description logic remains unchanged) ...
        const headings = Array.from(document.querySelectorAll('h2, h3, h4, b, strong, span'));
        const descHeading = headings.find(h => {
          const t = (h as any).innerText || (h as any).textContent || "";
          return t.includes('Productbeschrijving') || t.includes('Product description');
        });

        if (descHeading) {
          const parent = descHeading.closest('section') || descHeading.parentElement;
          if (parent) {
            const clone = parent.cloneNode(true) as HTMLElement;
            const UI_SELECTORS = ['.js_description_read_more', '[data-test="read-more"]', '.pdp-description__read-more', 'button', 'a.button--link'];
            UI_SELECTORS.forEach(sel => {
              clone.querySelectorAll(sel).forEach(btn => btn.remove());
            });
            const text = (clone as any).innerText || (clone as any).textContent || "";
            description = text.replace(/Productbeschrijving|Product description/i, '').trim();
            // Global clean for UI fragments
            description = description.replace(/toon meer|toon minder/gi, '').trim();
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
            '[data-test="product-info"]'
          ];

          // Try to expand "Lees meer" if it exists
          const readMore = document.querySelector('.js_description_read_more, [data-test="read-more"], .pdp-description__read-more');
          if (readMore) (readMore as any).click();

          let pooledText = "";
          for (const s of descSelectors) {
            const el = document.querySelector(s);
            if (el) {
              const clone = el.cloneNode(true) as HTMLElement;
              const UI_SELECTORS = ['.js_description_read_more', '[data-test="read-more"]', '.pdp-description__read-more', 'button', 'a.button--link'];
              UI_SELECTORS.forEach(sel => {
                clone.querySelectorAll(sel).forEach(btn => btn.remove());
              });
              const text = (clone as any).innerText || (clone as any).textContent || "";
              if (text.trim().length > 20) {
                let html = clone.innerHTML;
                let t = html.replace(/<br\s*\/?>/gi, '\n').replace(/<\/p>/gi, '\n').replace(/<\/div>/gi, '\n').replace(/<[^>]*>?/gm, ' ').replace(/\s+/g, ' ').trim();
                // Secondary cleanup for text fragments
                t = t.replace(/toon meer|toon minder/gi, '').trim();

                if (!pooledText.includes(t)) {
                  pooledText += (pooledText ? "\n\n" : "") + t;
                }
              }
            }
          }
          if (pooledText) description = pooledText;
        }

        // 3. Price (Prioritize LD+JSON as requested)
        let price = "N/A";

        try {
          const scripts = Array.from(document.querySelectorAll('script[type="application/ld+json"]'));
          for (const script of scripts) {
            try {
              const data = JSON.parse(script.textContent || "");
              const findPrice = (obj: any): string | null => {
                if (obj.offers) {
                  if (Array.isArray(obj.offers)) return obj.offers[0].price || null;
                  return obj.offers.price || null;
                }
                if (obj["@type"] === "Offer") return obj.price || null;
                return null;
              };

              if (Array.isArray(data)) {
                for (const item of data) {
                  const p = findPrice(item);
                  if (p) { price = String(p); break; }
                }
              } else {
                const p = findPrice(data);
                if (p) price = String(p);
              }
              if (price !== "N/A") break;
            } catch (e) { }
          }
        } catch (e) { }

        if (price === "N/A") {
          // Try multiple comprehensive selectors for BOL price
          const priceSelectors = [
            '[data-test="current-price"]',
            '[data-test="price"]',
            'span[class*="promo-price"]',
            'span[class*="current-price"]',
            'span[itemprop="price"]',
            '#price',
            '.price',
            '[class*="ProductPrice"]',
            'span[class*="Price"]',
            'div[class*="price-content"]',
            '[data-element-type="price"]'
          ];

          for (const selector of priceSelectors) {
            const priceEl = document.querySelector(selector);
            if (priceEl) {
              const text = (priceEl as any).innerText || (priceEl as any).textContent;
              if (text) {
                const match = text.match(/[\d.]+,\d{2}|[\d,]+/);
                if (match) {
                  price = match[0].replace(',', '.');
                  break;
                }
              }
            }
          }

          // Fallback: search in all text
          if (price === "N/A") {
            const allText = document.body.innerText;
            const priceMatch = allText.match(/€\s*([\d.]+,\d{2})/);
            if (priceMatch) price = priceMatch[1].replace(',', '.');
          }
        }

        // 4. Shipping (JSON-LD + User snippet indicators)
        let shippingText = "N/A";

        try {
          const scripts = Array.from(document.querySelectorAll('script[type="application/ld+json"]'));
          for (const script of scripts) {
            try {
              const data = JSON.parse(script.textContent || "");
              const findShipping = (obj: any): string | null => {
                const ship = obj.shippingDetails || obj.offers?.shippingDetails;
                if (ship && ship.deliveryTime) {
                  const dt = ship.deliveryTime;
                  if (dt.transitTime) return `Delivery in ${dt.transitTime.minValue || dt.transitTime.value}-${dt.transitTime.maxValue || dt.transitTime.value} days`;
                }
                return null;
              };

              const s = Array.isArray(data) ? data.map(findShipping).find(x => x) : findShipping(data);
              if (s) { shippingText = s; break; }
            } catch (e) { }
          }
        } catch (e) { }

        if (shippingText === "N/A") {
          // Try multiple comprehensive selectors for BOL shipping
          const shippingSelectors = [
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

          for (const selector of shippingSelectors) {
            const shippingEl = document.querySelector(selector);
            if (shippingEl) {
              const text = (shippingEl as any).innerText || (shippingEl as any).textContent;
              if (text && text.trim().length > 0) {
                shippingText = text.trim();
                break;
              }
            }
          }

          // Fallback: search in body text for date patterns
          if (shippingText === "N/A") {
            const bodyText = document.body.innerText;
            const uiterlijkMatch = bodyText.match(/Uiterlijk\s+(.+?)(?:\s+in\s+huis|$)/i) ||
              bodyText.match(/Morgen\s+in\s+huis/i) ||
              bodyText.match(/Vandaag\s+.*?(?:in|om)/i) ||
              bodyText.match(/Bezorging:\s+(.+?)(?:\n|$)/i);
            if (uiterlijkMatch) shippingText = uiterlijkMatch[0] || uiterlijkMatch[1] || "N/A";
          }
        }

        // 5. Images (Revised extraction to avoid empty results)
        let images: string[] = [];

        // Strategy 1: Preload tags (Often has high-res URLs)
        try {
          const preloads = Array.from(document.querySelectorAll('link[rel="preload"][as="image"]'));
          preloads.forEach(link => {
            const href = link.getAttribute('href');
            if (href && href.includes('media.s-bol.com')) images.push(href);
          });
        } catch (e) { }

        // Strategy 2: Main image selectors
        const mainSelectors = [
          '[data-test="product-main-image"] img',
          '.js_main_product_image',
          '.pdp-main-image img',
          'img.js_main_product_image'
        ];
        for (const s of mainSelectors) {
          const img = document.querySelector(s);
          if (img) {
            const src = (img as any).src || img.getAttribute('src');
            if (src && src.startsWith('http')) images.push(src);
          }
        }

        // Strategy 3: Thumbnails
        const thumbSelectors = [
          '.js_product_media_items img',
          '.pdp-images img',
          '.js_image_container img',
          '.product-images__item img'
        ];
        for (const s of thumbSelectors) {
          const thumbs = Array.from(document.querySelectorAll(s));
          thumbs.forEach(img => {
            const src = (img as any).src || img.getAttribute('data-src') || img.getAttribute('src');
            if (src && src.includes('media.s-bol.com')) {
              images.push(src);
            }
          });
        }

        // 6. Variations & A+
        const hasVariations = !!(
          document.querySelector('[data-test="variant-selector"]') ||
          document.querySelector('[data-test="product-variant"]') ||
          document.querySelector('[data-test="attribute-selector"]') ||
          document.querySelector('.variant-selector') ||
          document.querySelector('.js_bundle_as_variant_selector') ||
          document.querySelector('[data-test="variant-dropdown"]') ||
          document.querySelector('.js_variant_selector') ||
          document.querySelector('[data-test="variant-list"]') ||
          document.querySelector('.pdp-variant-selector') ||
          document.querySelector('.js_variant_dropdown') ||
          document.querySelector('[class*="VariantSelector"]') ||
          document.querySelector('[class*="variant"]') ||
          document.querySelector('select[name*="variant"]') ||
          document.querySelector('select[data-test*="variant"]') ||
          // Broad pattern matching for variation themes and buttons
          Array.from(document.querySelectorAll('button, span, div, h3, label, select')).some(el => {
            const t = ((el as any).innerText || "").toLowerCase();
            const themes = [
              'kies je kleur',
              'kies je maat',
              'kies je variant',
              'kies je bestanddeel',
              'kies je model',
              'kies je type',
              'kies je uitvoering',
              'kies je breedte',
              'kies je lengte',
              'kies je inhoud',
              'kies je gewicht',
              'kies je verpakking',
              'kies je aantal',
              'kies je materiaal',
              'kies je geur',
              'kies je smaak',
              'kies je stijl',
              'kies je set',
              'kies je platform',
              'selecteer',
              'select variant',
              'choose',
              'color',
              'size',
              'option'
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
        } catch (e) { }
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
        images: liveDataRaw.images && liveDataRaw.images.length > 0 ? getUniqueImages(liveDataRaw.images) : []
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

    const descriptionMatchValue = (mode === 'amazon' && auditResult.hasAPlus.live)
      ? "A+ content available"
      : getMatchText(auditResult.description.match);

    const sharedData: any = {
      "Identifier": identifier,
      "SKU": identifier,
      "Title match": getMatchText(auditResult.title.match),
      "Title Match": getMatchText(auditResult.title.match),
      "Description match": descriptionMatchValue,
      "Description Match": descriptionMatchValue,
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
      const url = masterImgs[i - 1] || "";
      sharedData[`${prefix} Master Image ${i}`] = url ? `=IMAGE("${url}")` : "";
    }

    // Live Images 1-10
    for (let i = 1; i <= 10; i++) {
      const url = liveImgs[i - 1] || "";
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

  function cleanText(text: string): string {
    if (!text) return "";
    // 1. Strip HTML tags
    let cleaned = text.replace(/<[^>]*>?/gm, '');

    // 2. Remove raw URLs
    cleaned = cleaned.replace(/https?:\/\/[^\s]+\.(jpg|jpeg|png|gif|webp|svg)/gi, '');

    // 3. Normalize diacritics (treat ë as e, etc.)
    cleaned = cleaned.normalize('NFD').replace(/[\u0300-\u036f]/g, "");

    // 4. Whitespace Normalization (replace non-breaking spaces etc with standard space)
    cleaned = cleaned.replace(/[\u00A0\u1680\u180E\u2000-\u200B\u202F\u205F\u3000\uFEFF]/g, ' ');

    // 5. Aggressive Alphanumeric filter (Removes symbols like ™, ®, ©, bullets, punctuation)
    cleaned = cleaned.replace(/[^a-zA-Z0-9\s]/g, '');

    // 6. Normalize casing and multiple spaces
    return cleaned.toLowerCase().trim().replace(/\s+/g, ' ');
  }

  function getUniqueImages(urlList: string[]) {
    const idToUrlMap = new Map<string, string>();
    const finalImages: string[] = [];

    for (const url of urlList) {
      if (!url || typeof url !== 'string' || !url.startsWith('http')) continue;

      let imgId = "";
      let isAmazon = false;
      let processedUrl = url;

      if (url.includes('amazon.com') || url.includes('media-amazon.com')) {
        isAmazon = true;
        try {
          if (url.includes('/I/')) {
            imgId = url.split('/I/')[1].split('.')[0];
          } else {
            const match = url.match(/\/images\/I\/([A-Za-z0-9_-]+)/);
            if (match) imgId = match[1];
          }
        } catch (e) {
          const match = url.match(/\/images\/I\/([A-Za-z0-9_-]+)/);
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
            const getRes = (u: string) => {
              if (u.includes('SL1500') || u.includes('original')) return 9999;
              const match = u.match(/SL(\d+)/);
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

    return [...Array.from(idToUrlMap.values()), ...finalImages];
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
      match: titleSimilarity >= 0.98
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
      match: (isImageDesc || isAPlusImages || isAPlusData) ? false : (descSimilarity >= 0.98),
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

    // Bullet Points Comparison
    results.bullets = (master.bullets || []).map((mb: string, i: number) => {
      const lb = live.bullets?.[i] || "";
      const cmb = cleanText(mb);
      const clb = cleanText(lb);
      return {
        master: mb,
        live: lb,
        similarity: stringSimilarity.compareTwoStrings(cmb, clb),
        match: cmb === clb
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
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer().catch(err => {
  console.error("Failed to start server:", err);
  process.exit(1);
});
