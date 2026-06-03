import { chromium as chromiumExtra } from 'playwright-extra';
import stealth from 'puppeteer-extra-plugin-stealth';
import { GoogleGenAI } from '@google/genai';
import { getSimilarity, cleanAndNormalizePrice } from './helpers.ts';

function isRailwayDeployment(): boolean {
  return !!(
    process.env.RAILWAY_STATIC_URL ||
    process.env.RAILWAY_ENVIRONMENT ||
    process.env.RAILWAY_SERVICE_ID ||
    process.env.PORT_BOL_SCRAPE_DIRECT_PROHIBITED ||
    process.env.NODE_ENV === 'production'
  );
}

// Register standard browser stealth plugins on the extra chromium instance
try {
  chromiumExtra.use(stealth());
} catch (err: any) {
  console.warn('[BOL SETUP] Web stealth registrations:', err.message);
}

// ── BOL STRATEGY 1: Gemini Google Search Grounding ───────────────────────────
export async function tryBolViaGemini(ean: string, masterTitle?: string): Promise<any | null> {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.log('[BOL GEMINI] No GEMINI_API_KEY found, skipping.');
    return null;
  }

  const prompt = `You are a professional product intelligence scraper for bol.com.
Locate the correct product matching the EAN "${ean}" and product title "${masterTitle || ''}".

Please follow these exact steps to find the product:
1. Search Google using Google Search tool for 'site:bol.com "${ean}"' or 'bol.com ${ean}'.
2. If that search yields very few results or does not find the specific product page, search Google for 'site:bol.com "${masterTitle || ''}"' or search for 'site:bol.com ${ean} ${masterTitle || ''}'.
3. Find the official product listing page on bol.com and extract information.
4. Extract the live attributes of the product: title, price, shipping time message, description, bullet points, image URLs, product URL, variations, and buyboxOwner (seller name). Ensure they are fully grounded in search results.

Return ONLY a single, exact JSON object matching the following structure. No other conversational text, no other markdown text. Ensure it is wrapped in an exact \`\`\`json markdown block:
{
  "title": "exact full product title on bol.com",
  "price": "correct numerical price string e.g. 14.99",
  "shipping": "correct shipping time/delivery message e.g. 'Morgen in huis' or 'Uiterlijk donderdag 22 mei'",
  "description": "product description details, first 500 characters",
  "images": ["image url 1", "image url 2"],
  "bullets": ["feature point 1", "feature point 2"],
  "productUrl": "the direct final product link on bol.com",
  "liveVariations": "variation options if any, else empty string",
  "buyboxOwner": "The seller name (verkocht door). Defaults to 'bol.com' if sold/shipped by them."
}

Ensure all extracted values (pricing, title, shipping) are fully grounded in search results. If an attribute cannot be found, set it to "N/A" rather than leaving it empty.`;

  const normalizeParsedData = (obj: any) => {
    if (!obj) return obj;
    if (obj.productUrl && typeof obj.productUrl === 'string') {
      const url = obj.productUrl.trim();
      if (url && !url.startsWith('http') && !url.startsWith('//')) {
        obj.productUrl = 'https://www.bol.com' + (url.startsWith('/') ? '' : '/') + url;
      }
    }
    if (Array.isArray(obj.images)) {
      obj.images = obj.images.map((img: any) => {
        if (typeof img === 'string') {
          const trimmed = img.trim();
          if (trimmed && !trimmed.startsWith('http') && !trimmed.startsWith('//')) {
            return 'https://www.bol.com' + (trimmed.startsWith('/') ? '' : '/') + trimmed;
          }
          return trimmed;
        }
        return img;
      }).filter(Boolean);
    }
    return obj;
  };

  const maxAttempts = 3;
  let lastError: any = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const genai = new GoogleGenAI({
        apiKey,
        httpOptions: {
          headers: {
            'User-Agent': 'aistudio-build',
          }
        }
      });

      console.log(`[BOL GEMINI] Generating content with googleSearch grounding (attempt ${attempt}/${maxAttempts}) for EAN:`, ean);

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
          const finalized = normalizeParsedData(parsed);
          finalized._source = 'gemini-google-search';
          return finalized;
        }
      } catch (parseErr) {
        console.log('[BOL GEMINI] Direct JSON parse failed, trying regex extraction. Raw text length:', rawText.length);
        const jsonMatch = rawText.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
          try {
            const extracted = JSON.parse(jsonMatch[0]);
            if (extracted && extracted.title) {
              const finalized = normalizeParsedData(extracted);
              finalized._source = 'gemini-google-search';
              return finalized;
            }
          } catch (_) {}
        }
        console.log('[BOL GEMINI] Regex JSON parse failed:', parseErr);
      }

      lastError = new Error('JSON parsing failed from Gemini output response');

    } catch (e: any) {
      lastError = e;
      console.log(`[BOL GEMINI] Attempt ${attempt} failed:`, e.message);

      const isRateOrQuotaLimit = 
        e.message?.includes('429') || 
        e.message?.includes('RESOURCE_EXHAUSTED') || 
        e.message?.includes('quota') || 
        e.message?.includes('limit');

      if (isRateOrQuotaLimit && attempt < maxAttempts) {
        const backoffMs = attempt * 5000 + Math.floor(Math.random() * 2000);
        console.log(`[BOL GEMINI] Rate limit / quota error detected. Retrying in ${backoffMs}ms...`);
        await new Promise(resolve => setTimeout(resolve, backoffMs));
      } else {
        break;
      }
    }
  }

  if (lastError) {
    console.log('[BOL GEMINI] All Gemini attempts exhausted. Final error:', lastError.message);
  }
  return null;
}

// ── BOL STRATEGY 2: Playwright stealth browser (hardened backup) ─────────────
export async function goToProduct(page: any, searchTerm: string) {
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
      }, { timeout: 6000, polling: 500 });
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
            page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 5000 }).catch(() => null),
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
          await page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 5000 }).catch(() => null);
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
      timeout: 20000
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
    await page.goto(searchUrl, { waitUntil: 'domcontentloaded', timeout: 20000 });
  } catch (_) {}

  await handleCookieConsent();
  await page.waitForTimeout(1200);

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
    }, { timeout: 10000, polling: 500 });
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
        await page.waitForSelector(sel, { state: 'attached', timeout: 5000 });
        foundResults = true;
        console.log(`[BOL BROWSER] Search results found: ${sel}`);
        break;
      } catch (_) {}
    }
    if (!foundResults) {
      await page.waitForTimeout(1000);
      await page.mouse.wheel(0, 500);
      await page.waitForTimeout(1000);
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
      await page.goto(fullUrl, { waitUntil: 'domcontentloaded', timeout: 20000 }).catch(() => null);
      await page.waitForTimeout(1000);
    } else {
      const debugTitle = await page.title().catch(() => '');
      const debugUrl = page.url();
      const debugContent = await page.content().catch(() => '');
      const snippet = debugContent.replace(/\s+/g, ' ').substring(0, 250);
      throw new Error(`No product link found on Bol.com. Title: "${debugTitle}", Snippet: ${snippet}`);
    }
  }
}

export async function extractCatalogue(page: any) {
  let initialContent = await page.content().catch(() => '');
  let initialContentLower = initialContent.toLowerCase();
  if ((initialContentLower.includes('ip adres') && initialContentLower.includes('geblokkeerd')) || 
      initialContentLower.includes('rustig aan speed racer') ||
      initialContentLower.includes('human verification')) {
    throw new Error('WAF_BLOCKED: Bol.com blocked the request. IP address is blocked by their anti-bot system.');
  }

  await page.waitForLoadState('domcontentloaded', { timeout: 15000 }).catch(() => null);
  await page.waitForTimeout(1200);

  await page.mouse.wheel(0, 400);
  await page.waitForTimeout(1200);

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
    
    // 1. Try to find price in JSON-LD
    const jsonLdMatch = pageHtml.match(/"offers"\s*:\s*\{[^}]*"@type"\s*:\s*"Offer"[^}]*"price"\s*:\s*"([\d.]+)"/) ||
                        pageHtml.match(/"offers"\s*:\s*\{[^}]*"@type"\s*:\s*"Offer"[^}]*"price"\s*:\s*([\d.]+)/);
    
    if (jsonLdMatch) {
      price = jsonLdMatch[1];
    } else {
      // Fallback
      const metaPriceMatch = pageHtml.match(/"price"\s*:\s*"([\d.]+)"\s*,\s*"priceCurrency"\s*:\s*"EUR"/) ||
                             pageHtml.match(/"price"\s*:\s*([\d.]+)\s*,\s*"priceCurrency"\s*:\s*"EUR"/);
      if (metaPriceMatch) price = metaPriceMatch[1];
    }

    // 2. Last resort: text-based extraction from the UI
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
    
    // Main image
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

    // Extract thumbnails
    const allImages = Array.from(document.querySelectorAll('img'));
    allImages.forEach(img => {
      const alt = img.getAttribute('alt') || '';
      const src = img.src || img.getAttribute('data-src') || '';
      
      if (alt.includes('Afbeelding nummer')) {
        if (src && src.startsWith('http')) imgs.push(src);
      }
    });

    // Thumbnail fallback
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
      '.product-specifications li',
      '.specs-list-item',
      '[data-test="product-specifications"] li',
      '.key-benefits__list-item',
      '.pdp-specs__row',
      '.key-benefit',
      '.product-features__item',
      '.usp-list li',
      '.usp-item'
    ];
    const bulletSet = new Set<string>();
    bulletSel.forEach(sel => {
      document.querySelectorAll(sel).forEach(li => {
        const txt = (li as HTMLElement).innerText.trim();
        if (txt.length > 3) bulletSet.add(txt);
      });
    });

    // Variations
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

    let buyboxOwner = 'bol.com';
    const sellerSelector = [
      '[data-test="seller-link"]',
      'a[data-test="seller-link"]',
      '.buy-block__seller-name a',
      '.pdp-seller-link',
      'a[id*="seller"]',
      'span[class*="seller"]'
    ];
    for (const sel of sellerSelector) {
      const el = document.querySelector(sel);
      if (el) {
        const txt = (el as HTMLElement).innerText.trim();
        if (txt) {
          buyboxOwner = txt;
          break;
        }
      }
    }
    if (buyboxOwner === 'bol.com') {
      const allText = document.body.innerText;
      const soldByMatch = allText.match(/verkocht\s+door\s*:?\s*([^\n\r]+)/i) || 
                          allText.match(/verkoop\s+door\s*:?\s*([^\n\r]+)/i);
      if (soldByMatch) {
        buyboxOwner = soldByMatch[1].trim();
      }
    }

    return {
      title,
      description,
      price,
      shipping,
      images: Array.from(new Set(imgs)),
      bullets: Array.from(bulletSet),
      liveVariations: variationsData,
      buyboxOwner
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

export async function performBolAudit(master: any, live: any) {
  const result: any = {
    title: { master: master.title, live: live.title, similarity: getSimilarity(master.title, live.title), match: false },
    description: { 
      master: master.description, 
      live: live.description || "", 
      similarity: 0, 
      match: false, 
      isAPlus: false 
    },
    bullets: [],
    price: { master: master.price, live: live.price, match: false },
    shipping: { master: master.shipping, live: live.shipping, match: false, days: live.shippingDays },
    images: { master: master.images, live: live.images, match: false },
    variations: { match: live.variations > 1 }
  };

  result.description.similarity = getSimilarity(master.description || "", live.description || "");
  if (result.description.similarity > 0.6) result.description.match = true;

  if (result.title.similarity > 0.8) result.title.match = true;
  
  const masterBullets = Array.isArray(master.bullets) ? master.bullets.filter(Boolean) : [];
  const liveBullets = Array.isArray(live.bullets) ? live.bullets.filter(Boolean) : [];
  const bulletsResults: any[] = [];
  const matchedLiveIndices = new Set<number>();

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

  const masterPriceCleaned = cleanAndNormalizePrice(String(master.price || ""));
  const livePriceCleaned = cleanAndNormalizePrice(String(live.price || ""));
  const masterPriceNum = parseFloat(masterPriceCleaned) || 0;
  const livePriceNum = parseFloat(livePriceCleaned) || 0;
  if (masterPriceNum > 0 && Math.abs(masterPriceNum - livePriceNum) < 1.0) result.price.match = true;

  if (live.images && live.images.length >= (master.images?.length || 1)) result.images.match = true;

  let scoreValue = 0;
  if (result.title.match) scoreValue += 50;
  if (result.description.match) scoreValue += 50;
  result.score = scoreValue;

  return result;
}

// Main Bol Audit Function
export async function auditBol(ean: string, masterData: any) {
  let browser;
  try {
    let data: any = null;
    let dataSource = 'browser';

    if (isRailwayDeployment()) {
      console.log('[BOL] Railway environment detected. Activating prioritised WAF-evasion routing (Gemini Search Grounding via Google to bypass Akamai IP restrictions).');
    }

    // ── Strategy 1: Gemini Google Search Grounding ─────────────────────────
    console.log('[BOL] Trying Strategy 1: Gemini Google Search Grounding...');
    data = await tryBolViaGemini(ean, masterData?.title);
    if (data) {
      console.log('[BOL] Strategy 1 succeeded via Gemini.');
      dataSource = 'gemini';
    }

    // ── Strategy 2: Playwright stealth browser (hardened) ─────────────────────
    if (!data) {
      try {
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

        // Pre-inject OneTrust consent cookies
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
      } catch (browserError: any) {
        console.error("[BOL] Browser strategy failed:", browserError.message);
        
        // Anti-WAF recovery with fallback to Gemini Search grounding (with 2-second timeout)
        if (browserError.message.includes('WAF_BLOCKED') || browserError.message.includes('anti-bot') || browserError.message.includes('blocked')) {
          console.log('[BOL] Browser block detected. Attempting recovery via relaxed Gemini Google Search Grounding...');
          data = await tryBolViaGemini(ean, masterData?.title);
          if (data && data.title) {
            console.log('[BOL WAF RECOVERY] Re-running Gemini Grounding was successful.');
            dataSource = 'gemini-recovery-after-waf';
          } else {
            throw browserError;
          }
        } else {
          throw browserError;
        }
      }
    }

    const bolShippingDays = calculateBolShippingDays(data.shipping || '');

    const liveData = {
      title: data.title || '',
      price: data.price || 'N/A',
      buyboxOwner: data.buyboxOwner || 'bol.com',
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

    const auditResult = await performBolAudit(masterData, liveData);
    return { liveData, auditResult };

  } finally {
    if (browser) await browser.close();
  }
}
