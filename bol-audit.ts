import { chromium as chromiumExtra } from 'playwright-extra';
import stealth from 'puppeteer-extra-plugin-stealth';
import { GoogleGenAI } from '@google/genai';
import dotenv from 'dotenv';

chromiumExtra.use(stealth());
dotenv.config();

export async function tryBolViaGemini(ean: string): Promise<any | null> {
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

export async function launchBolBrowser() {
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

  const browser = await chromiumExtra.launch(launchOpts);

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

  // Pre-inject cookie consent
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
  return { browser, context, page };
}

export async function goToBolProduct(page: any, ean: string) {
  const searchUrl = `https://www.bol.com/nl/nl/s/?searchtext=${encodeURIComponent(ean)}`;
  console.log(`[BOL BROWSER] Searching for: ${ean}`);

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

  console.log('[BOL BROWSER] Visited homepage to warm Akamai session...');
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

  await waitForAkamai();

  const searchContent = await page.content().catch(() => '');
  if (isHardBlocked(searchContent)) {
    throw new Error('WAF_BLOCKED: Bol.com blocked the request. IP address is blocked by their anti-bot system.');
  }

  // Handle product click if listed in search results
  const pLink = await page.$('a.product-title, [data-test="product-title"], .product-item__title a').catch(() => null);
  if (pLink) {
    console.log('[BOL BROWSER] Search result found, clicking the product page...');
    await Promise.all([
      page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 15_000 }).catch(() => null),
      pLink.click().catch(() => null)
    ]);
  }
}

export async function extractBolCatalogue(page: any) {
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

    let price = 'N/A';
    const pageHtml = document.documentElement.innerHTML;
    const jsonLdMatch = pageHtml.match(/"offers"\s*:\s*\{[^}]*"@type"\s*:\s*"Offer"[^}]*"price"\s*:\s*"([\d.]+)"/) ||
                        pageHtml.match(/"offers"\s*:\s*\{[^}]*"@type"\s*:\s*"Offer"[^}]*"price"\s*:\s*([\d.]+)/);
    
    if (jsonLdMatch) {
      price = jsonLdMatch[1];
    } else {
      const metaPriceMatch = pageHtml.match(/"price"\s*:\s*"([\d.]+)"\s*,\s*"priceCurrency"\s*:\s*"EUR"/) ||
                             pageHtml.match(/"price"\s*:\s*([\d.]+)\s*,\s*"priceCurrency"\s*:\s*"EUR"/);
      if (metaPriceMatch) price = metaPriceMatch[1];
    }

    if (price === 'N/A') {
      const allText = document.body.innerText;
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

    const allImages = Array.from(document.querySelectorAll('img'));
    allImages.forEach(img => {
      const alt = img.getAttribute('alt') || '';
      const src = img.src || img.getAttribute('data-src') || '';
      if (alt.includes('Afbeelding nummer')) {
        if (src && src.startsWith('http')) imgs.push(src);
      }
    });

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

    let variationsData = '';
    const varItems = Array.from(document.querySelectorAll('div, label, a, span, button')).filter(el => {
      const t = (el.textContent || '').toLowerCase();
      return t.includes('kies je ');
    });
    
    if (varItems.length > 0) {
       const container = varItems[0].closest('section, div.variant-container, div.product-variants, [data-test="variants"], [class*="variant"]') || varItems[0].parentElement?.parentElement;
       if (container) {
         variationsData = (container as HTMLElement).innerText.trim().replace(/\s+/g, ' ');
       } else {
         variationsData = "Various options found: " + varItems.map(v => v.textContent?.replace(/\s+/g, ' ').trim()).join(' | ');
       }
    } else {
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
      productUrl: window.location.href,
      liveVariations: variationsData
    };
  });
}

export function calculateBolShippingDays(rawShippingTime: string): string {
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

