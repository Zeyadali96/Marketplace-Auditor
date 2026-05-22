import { chromium as chromiumExtra } from 'playwright-extra';
import stealth from 'puppeteer-extra-plugin-stealth';
import * as cheerio from 'cheerio';
import { GoogleGenAI } from '@google/genai';
import dotenv from 'dotenv';

chromiumExtra.use(stealth());
dotenv.config();

export interface AmazonConfig {
  locale: string;
  timezoneId: string;
  city: string;
  zip: string;
  currency: string;
  countryCode?: string;
  deliverTo: string[];
}

export const amazonLocalizationMap: Record<string, AmazonConfig> = {
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

function cleanAndNormalizePrice(priceStr: string): string {
  if (!priceStr) return "";
  let s = priceStr.trim();
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
  const match = s.match(/\d+(\.\d+)?/);
  return match ? match[0] : s.replace(/[^0-9.]/g, '');
}

export async function launchAmazonBrowser(domain: string) {
  const proxyServer = process.env.PROXY_SERVER;
  const launchOptions: any = {
    headless: false,
    args: [
      '--headless=new',
      '--no-sandbox', 
      '--disable-setuid-sandbox', 
      '--disable-dev-shm-usage', 
      '--disable-gpu', 
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

  const browser = await chromiumExtra.launch(launchOptions).catch(err => {
    console.error("AMAZON AUDIT FAILED TO LAUNCH CHROMIUM:", err);
    throw new Error(`Browser launch failed. Error: ${err.message}`);
  });

  const locConfig = amazonLocalizationMap[domain] || { locale: 'en-US', timezoneId: 'America/New_York', city: 'NYC', zip: '10001', currency: 'USD', deliverTo: ['Deliver to'] };

  const context = await browser.newContext({
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
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

  return { browser, context, page, config: locConfig };
}

export async function extractAmazonData(page: any, asin: string, domain: string, config: AmazonConfig): Promise<string> {
  const url = `https://www.${domain}/dp/${asin}`;
  console.log(`[AMAZON EXTRACT] Navigating to ${url}`);
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });

  // Accept cookies
  try {
    const cookieButtons = ['#sp-cc-accept', 'input[name="accept"]', '#cookie-accept', '#accept-cookies', '.a-button-inner input[data-action="accept-cookies"]'];
    for (const selector of cookieButtons) {
      if (await page.isVisible(selector).catch(() => false)) {
        await page.click(selector).catch(() => null);
        await page.waitForTimeout(500);
        break;
      }
    }
  } catch (err) { /* ignored */ }

  const title = await page.title().catch(() => '');
  console.log(`[AMAZON EXTRACT] Page Title: "${title}"`);

  // Handle regional delivery zip code selection
  try {
    const isRegionalLocked = await page.evaluate(({ zip }) => {
      const slot = document.querySelector('#nav-global-location-slot');
      if (!slot) return true;
      return !slot.textContent?.includes(zip);
    }, { zip: config.zip }).catch(() => true);

    if (isRegionalLocked) {
      console.log(`[AMAZON EXTRACT] Region is locked or different. Injecting postcode ${config.zip} for ${domain}`);
      const locBtn = await page.waitForSelector('#nav-global-location-slot, #glow-ingress-block, #nav-main-ftr-location-slot', { visible: true, timeout: 10000 }).catch(() => null);
      if (locBtn) {
        await locBtn.click({ force: true });
        await page.waitForTimeout(2000);

        // Check for country selection dropdown
        const countryListSelector = '#GLUXCountryList';
        const countryListVisible = await page.locator(countryListSelector).isVisible().catch(() => false);
        if (countryListVisible && config.countryCode) {
          console.log(`[AMAZON EXTRACT] Selecting country block: ${config.countryCode}`);
          await page.selectOption(countryListSelector, { value: config.countryCode }).catch(() => null);
          await page.waitForTimeout(1000);

          const goBtnSelector = 'input[aria-labelledby="GLUXCountryList-announce"], button[name="glowDoneButton"], #GLUXCountryList-announce + input, .a-popover-footer input';
          await page.click(goBtnSelector, { force: true }).catch(() => null);
          await page.waitForTimeout(3000);

          console.log('[AMAZON EXTRACT] Re-clicking global location slot...');
          await page.click('#nav-global-location-slot, #glow-ingress-block, #nav-main-ftr-location-slot', { force: true }).catch(() => null);
          await page.waitForTimeout(2000);
        }

        // Fill standard postcode field
        const zipInputSelector = '#GLUXZipUpdateInput, #GLUXZipUpdateInput_0, input[aria-label*="zip"], input[aria-label*="code"], input[name="zipCode"]';
        const inputVisible = await page.waitForSelector(zipInputSelector, { state: 'visible', timeout: 5000 }).catch(() => null);
        
        if (inputVisible) {
          await page.evaluate((sel) => {
            const el = document.querySelector(sel) as HTMLInputElement;
            if (el) {
              el.value = '';
              el.focus();
            }
          }, zipInputSelector);
          
          await page.type(zipInputSelector, config.zip, { delay: 60 });
          await page.keyboard.press('Enter');
          
          const applyBtn = '#GLUXZipUpdate input[type="submit"], #GLUXZipUpdate > span > input, #GLUXZipUpdate_Buttons input, #GLUXZipUpdate input.a-button-input, #GLUXZipUpdate_Buttons span.a-button-inner input';
          await page.click(applyBtn).catch(() => null);
          await page.waitForTimeout(1500);
          
          const confirmBtn = '#GLUXConfirmClose, #GLUXConfirmResponse, input[data-action="GLUXConfirmResponse"], .a-popover-footer input, #GLUXConfirmClose input, #GLUXConfirmClose-announce, .a-popover-footer span.a-button-inner input, button[name="glowDoneButton"]';
          const confirmBtnVisible = await page.waitForSelector(confirmBtn, { timeout: 5000 }).catch(() => null);
          if (confirmBtnVisible) {
            await page.click(confirmBtn).catch(() => null);
          }
          await page.waitForTimeout(1500);
          
          console.log(`[AMAZON EXTRACT] Reloading page for ${domain} post-zipcode change...`);
          await page.reload({ waitUntil: 'domcontentloaded', timeout: 30000 }).catch(() => null);
        }
      }
    }
  } catch (err: any) {
    console.warn("[AMAZON EXTRACT] Location selection skipped/failed:", err.message);
  }

  await page.waitForLoadState('domcontentloaded', { timeout: 15000 }).catch(() => null);
  return await page.content();
}

export function parseAmazonContent(content: string, domain: string, config?: AmazonConfig) {
  const $ = cheerio.load(content);

  // 1. Title Extraction
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
      } catch (e) { /* ignored */ }
    });
  }
  if (!amazonTitle) amazonTitle = $('meta[name="title"]').attr('content')?.split(': Amazon')[0] || "";
  if (!amazonTitle) amazonTitle = $('h1').first().text().trim() || "";

  // 2. Buybox Scoping & Price Extraction
  const buyBoxContext = $('#oneTimeBuyBox').length ? $('#oneTimeBuyBox') :
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

  // 3. Shipping Extraction & Day Calculations
  let rawShippingTime = "";
  const primaryDelivery = $('#mir-layout-DELIVERY_BLOCK-slot-PRIMARY_DELIVERY_MESSAGE_LARGE').text().trim() ||
                          $('#deliveryBlockMessage').text().trim();
  const secondaryDelivery = $('#mir-layout-DELIVERY_BLOCK-slot-SECONDARY_DELIVERY_MESSAGE_LARGE').text().trim();

  const isFastDeliverySignal = (text: string) => {
    const t = text.toLowerCase();
    return t.includes('fastest') || t.includes('snelste') || t.includes('rapide') ||
           t.includes('tomorrow') || t.includes('morgen') || t.includes('demain') ||
           t.includes('domani') || t.includes('jutro') || t.includes('mañana') ||
           t.includes('imorgon') || t.includes('today') || t.includes('vandaag') ||
           t.includes('aujourd') || t.includes('oggi') || t.includes('heute');
  };

  if (secondaryDelivery && isFastDeliverySignal(secondaryDelivery)) {
    rawShippingTime = secondaryDelivery;
  } else if (primaryDelivery && isFastDeliverySignal(primaryDelivery)) {
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

  // 4. Buybox Owner Extraction
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
    const merchantInfoEl = buyBoxContext.find('#merchant-info').first();
    const mInfoRaw = merchantInfoEl.text() || $('#merchant-info').first().text();
    const mInfo = mInfoRaw.toLowerCase();

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
      const sellerLink = merchantInfoEl.find('a').first().text().trim() || $('#merchant-info a').first().text().trim();
      if (sellerLink && sellerLink.toLowerCase() !== 'amazon') {
        amazonBuyboxOwner = sellerLink;
      } else {
        amazonBuyboxOwner = mInfoRaw
          .replace(/Dispatched from and sold by\s*/i, '')
          .replace(/Dispatched from Amazon\s*\.?\s*/i, '')
          .replace(/Sold by\s*/i, '')
          .replace(/Fulfilled by Amazon\s*\.?\s*/i, '')
          .replace(/\|.*$/s, '')
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

  amazonBuyboxOwner = amazonBuyboxOwner
    .replace(/Sold by\s*:?\s*/gi, '')
    .replace(/Venduto da\s*:?\s*/gi, '')
    .replace(/Verkauf durch\s*:?\s*/gi, '')
    .replace(/Dispatched from and sold by\s*/gi, '')
    .replace(/Dispatched from Amazon\.?\s*/gi, '')
    .replace(/Fulfilled by Amazon\.?\s*/gi, '')
    .replace(/\|.*$/s, '')
    .trim();

  // 5. Desc & Hardened Bullets & A+ content
  let amazonDesc = $('#productDescription').text().trim();
  if (!amazonDesc) amazonDesc = $('#feature-bullets').text().trim();
  const hasAPlus = !!($('#aplus').length || $('#aplus_feature_div').length || $('div[id*="aplus"]').length);

  const bulletSet = new Set<string>();
  const bulletSelectors = [
    '#feature-bullets ul li:not(:has(ul))',
    '#featurebullets_feature_div ul li:not(:has(ul))',
    '#feature-bullets-content li:not(:has(ul))',
    '[data-feature-name="product-facts"] .a-list-item',
    '.product-facts-title + .a-unordered-list li:not(:has(ul))',
    '#product-facts-grid li:not(:has(ul))',
    '#productFactsDesktopExpander .a-list-item'
  ];

  $(bulletSelectors.join(', ')).each((_, el) => {
    const $el = $(el);
    const junkContainers = [
      '#customerReviews', '#reviews-medley-footer', '#cm-cr-dp-review-list',
      '.customer_review', '#fbt_x_cl_div', '#legal-disclaimer',
      '#ad-feedback-form-desktop-feature-bullets_secondary_view_div',
      '#reviews-image-gallery-container', '#social-proofing-faceout-feature-div',
      '#dp-ads-center-promo-pc_desktop_view_div', '.cr-widget-FocalReviews',
      '.a-expander-content.a-expander-partial-collapse-content'
    ];
    if ($el.closest(junkContainers.join(', ')).length > 0) return;

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
      return (
        lower.includes('window.ue') || lower.includes('if(window.ue)') ||
        lower.includes('out of 5 stars') || lower.includes('verified purchase') ||
        lower.includes('helpful report') || lower.includes('reviewed in') ||
        lower.includes('not for sale to persons under') || lower.includes('16 years of age') ||
        lower.includes('make sure this fits') || lower.includes('geben sie ihr modell ein') ||
        lower.includes('sprawdź, czy pasuje') || lower.includes('read more') ||
        lower.includes('customer reviews') || lower.includes('by placing an order') ||
        lower.includes('declare that you are') || lower.includes('used responsibly') ||
        t.length < 5 || t.length > 1000 ||
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

  // 6. Image extraction
  const imageMap = new Map<string, string>();
  const getNormalizedInfo = (url: string | undefined) => {
    if (!url || typeof url !== "string" || url.length < 15) return null;
    let cleaned = url.split("?")[0].trim();
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
  if (heroInfo) {
    imageMap.set(heroInfo.baseId, heroInfo.url);
  }

  $("#altImages li.imageThumbnail:not(.videoThumbnail) img, .imageThumbnail img, .altImages img").not("#landingImage").each((_, el) => {
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

  const uniqueImages = Array.from(imageMap.values());

  // 7. Variations Extraction
  const variationsSet = new Set<string>();
  const optionSelectors = [
    '#twister li[data-asin]',
    '#inline-twister-row-all-options li[data-asin]',
    '#twister .a-button-toggle',
    '#twister li[id^="color_name_"]',
    '#twister li[id^="size_name_"]',
    '#twister li[id^="style_name_"]'
  ];

  $(optionSelectors.join(', ')).each((_, el) => {
    const $el = $(el);
    const text = $el.text().trim();
    const asin = $el.attr('data-asin');
    if (asin) {
      variationsSet.add(asin);
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
        const val = $(opt).attr('data-asin') || $(opt).text().trim() || '';
        return val.length > 0 && !val.toLowerCase().includes('select');
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

  return {
    title: amazonTitle || '',
    description: amazonDesc || '',
    bullets: amazonBullets || [],
    price: amazonPrice || 'N/A',
    listPrice: listPrice || 'N/A',
    currency: config ? config.currency : 'USD',
    shipping: shippingDays !== "N/A" ? `${shippingDays} days` : (rawShippingTime || 'N/A'),
    shippingDays: shippingDays,
    rawShipping: rawShippingTime || 'N/A',
    variations: variationsCount,
    hasAPlus: hasAPlus,
    buyboxOwner: amazonBuyboxOwner || 'N/A',
    images: uniqueImages
  };
}

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

    const prompt = `Perform a google search for "amazon ${domain} product ${asin}" or search for "${asin}" directly on amazon.${domain}.
Locate the official product page on amazon.${domain}.
Extract and return a single, exact JSON object with the following schema:
{
  "title": "exact full product title on amazon.${domain}",
  "price": "correct numerical price string e.g. 14.99",
  "listPrice": "correct numerical list/basis price if any, else empty string",
  "shipping": "correct shipping time/delivery message e.g. '3 days' or 'Deliver by Tuesday'",
  "rawShipping": "raw original delivery message from amazon",
  "shippingDays": "numerical number of days e.g. '3', or '0' if same-day, or '1' if tomorrow, else 'N/A'",
  "description": "product description details or bullets if description is missing, first 500 characters",
  "images": ["image url 1", "image url 2"],
  "variants": 0,
  "hasAPlus": false,
  "bullets": ["feature point 1", "feature point 2"],
  "buyboxOwner": "The name of the vendor/brand in the Buybox, e.g. 'Amazon' or the third-party seller name. If not found, use 'N/A'"
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
    const jsonText = rawText.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();

    try {
      const parsed = JSON.parse(jsonText);
      if (parsed && parsed.title) {
        parsed._source = 'gemini-google-search';
        // Clean prices
        parsed.price = cleanAndNormalizePrice(parsed.price || '');
        parsed.listPrice = cleanAndNormalizePrice(parsed.listPrice || '');
        return parsed;
      }
    } catch (parseErr) {
      const jsonMatch = rawText.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        try {
          const extracted = JSON.parse(jsonMatch[0]);
          if (extracted && extracted.title) {
            extracted._source = 'gemini-google-search';
            extracted.price = cleanAndNormalizePrice(extracted.price || '');
            extracted.listPrice = cleanAndNormalizePrice(extracted.listPrice || '');
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
