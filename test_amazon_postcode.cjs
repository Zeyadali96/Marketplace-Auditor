const { chromium } = require('playwright-extra');
const stealth = require('puppeteer-extra-plugin-stealth')();
const cheerio = require('cheerio');
chromium.use(stealth);

async function cleanAndNormalizePrice(priceStr) {
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

(async () => {
  const launchOpts = {
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-blink-features=AutomationControlled']
  };

  if (process.env.PROXY_SERVER) {
    launchOpts.proxy = {
      server: process.env.PROXY_SERVER,
      username: process.env.PROXY_USERNAME,
      password: process.env.PROXY_PASSWORD
    };
    console.log('Using proxy server:', process.env.PROXY_SERVER);
  } else {
    console.log('No proxy environment variable found.');
  }

  const browser = await chromium.launch(launchOpts);

  const configs = [
    { domain: 'amazon.de', asin: 'B00OLZ9TJ8', zip: '10117', locale: 'de-DE', countryCode: 'DE' },
    { domain: 'amazon.co.uk', asin: 'B00OLZ9TJ8', zip: 'SW1A 1AA', locale: 'en-GB', countryCode: 'GB' }
  ];

  for (const config of configs) {
    console.log(`\n=================== TESTING http://www.${config.domain}/dp/${config.asin} ===================`);
    const context = await browser.newContext({
      userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36',
      viewport: { width: 1920, height: 1080 },
      locale: config.locale,
      extraHTTPHeaders: {
        'Accept-Language': `${config.locale},en-GB;q=0.9,en;q=0.8`
      }
    });

    const page = await context.newPage();
    const url = `https://www.${config.domain}/dp/${config.asin}`;

    try {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await page.waitForTimeout(2000);

      const title = await page.title();
      console.log(`Page Title: "${title}"`);
      if (title.includes('Robot') || title.includes('Captcha') || title.includes('Sorry!') || title.includes('Page Not Found')) {
        console.warn('WARNING: Captcha / Robot challenge page detected!');
      }

      // Check existence of some elements
      console.log('Waiting for location slot to be visible...');
      await page.waitForSelector('#nav-global-location-slot, #glow-ingress-block, #nav-main-ftr-location-slot', { timeout: 15000 }).catch(() => null);
      const hasSlot = await page.locator('#nav-global-location-slot').count() > 0;
      const hasIngress = await page.locator('#glow-ingress-block').count() > 0;
      const bodyClass = await page.locator('body').getAttribute('class').catch(() => '');
      console.log(`Presence check: hasSlot=${hasSlot}, hasIngress=${hasIngress}, bodyClass="${bodyClass}"`);

      // Handle cookie accept
      console.log('Checking for cookie banner...');
      const cookieButtons = ['#sp-cc-accept', 'input[name="accept"]', '#cookie-accept', '#accept-cookies', '.a-button-inner input[data-action="accept-cookies"]'];
      for (const selector of cookieButtons) {
        if (await page.locator(selector).isVisible().catch(() => false)) {
          await page.click(selector).catch(() => null);
          console.log('Clicked cookie accept.');
          await page.waitForTimeout(1000);
          break;
        }
      }

      console.log('Location slot initial text:', (await page.locator('#nav-global-location-slot').textContent().catch(() => '')).trim().replace(/\s+/g, ' '));

      // Click location button
      console.log('Clicking global location slot...');
      await page.click('#nav-global-location-slot, #glow-ingress-block, #nav-main-ftr-location-slot', { force: true }).catch(() => null);
      console.log('Waiting for popover input or country list to be visible...');
      await page.waitForSelector('#GLUXZipUpdateInput, #GLUXZipUpdateInput_0, #GLUXCountryList, input[aria-label*="zip"]', { timeout: 10000 }).catch(() => null);
      await page.waitForTimeout(1000);

      const zipInputSelector = '#GLUXZipUpdateInput, #GLUXZipUpdateInput_0, input[aria-label*="zip"], input[aria-label*="code"], input[name="zipCode"]';

      // Check if country list is present instead of zip input
      const countryListSelector = '#GLUXCountryList';
      const isCountryListVisible = await page.locator(countryListSelector).isVisible().catch(() => false);
      if (isCountryListVisible) {
        console.log('Country dropdown list detected inside location popover!');
        try {
          // Select the correct market country
          await page.selectOption(countryListSelector, { value: config.countryCode }).catch(() => null);
          console.log(`Selected country: ${config.countryCode}`);
          await page.waitForTimeout(1000);
          // Try clicking the apply button for country dropdown
          const goBtnSelector = 'input[aria-labelledby="GLUXCountryList-announce"], button[name="glowDoneButton"], #GLUXCountryList-announce + input, .a-popover-footer input';
          await page.click(goBtnSelector, { force: true }).catch(() => null);
          await page.waitForTimeout(3000);
          console.log('Re-clicking location slot after choosing country...');
          await page.click('#nav-global-location-slot, #glow-ingress-block', { force: true }).catch(() => null);
          await page.waitForTimeout(2000);
        } catch (err) {
          console.log('Error changing country dropdown:', err.message);
        }
      }

      const isZipVisible = await page.locator(zipInputSelector).isVisible().catch(() => false);
      if (isZipVisible) {
        console.log(`Zip input is visible! Filling postcode: ${config.zip}`);
        await page.locator(zipInputSelector).fill(config.zip).catch(() => null);
        await page.waitForTimeout(500);

        const applyBtn = '#GLUXZipUpdate input[type="submit"], #GLUXZipUpdate input, input[aria-labelledby="GLUXZipUpdate-announce"], #GLUXZipUpdate-announce + input';
        console.log('Clicking Apply button...');
        await page.click(applyBtn, { force: true }).catch(() => null);
        await page.waitForTimeout(2000);

        const confirmBtn = '#GLUXConfirmClose, #GLUXConfirmResponse, input[data-action="GLUXConfirmResponse"], .a-popover-footer input, #GLUXConfirmClose input, #GLUXConfirmClose-announce, button[name="glowDoneButton"]';
        const isConfirmVisible = await page.locator(confirmBtn).first().isVisible().catch(() => false);
        if (isConfirmVisible) {
          console.log('Done/Confirm button is visible. Clicking Done...');
          await page.locator(confirmBtn).first().click({ force: true }).catch(() => null);
          await page.waitForTimeout(2000);
        }

        console.log('Reloading page to refresh pricing...');
        await page.reload({ waitUntil: 'domcontentloaded', timeout: 30000 }).catch(() => null);
        await page.waitForTimeout(2000);
      } else {
        console.log('Zip code input was NOT visible in popover.');
      }

      const updatedLocText = (await page.locator('#nav-global-location-slot').textContent().catch(() => '')).trim().replace(/\s+/g, ' ');
      console.log('Location slot updated text:', updatedLocText);

      const content = await page.content();
      const $ = cheerio.load(content);

      // Scrape data
      let buyBoxContext = $('#oneTimeBuyBox').length ? $('#oneTimeBuyBox') :
                          $('#newAccordionRow').length ? $('#newAccordionRow') :
                          $('#buyNewSection').length ? $('#buyNewSection') :
                          $('#apex_offerDisplay_desktop').length ? $('#apex_offerDisplay_desktop') :
                          $('#newUnifiedOfferDisplay').length ? $('#newUnifiedOfferDisplay') :
                          $('#corePrice_feature_div').length ? $('#corePrice_feature_div') :
                          $('#desktop_buybox').length ? $('#desktop_buybox') :
                          $('#rightCol').length ? $('#rightCol') :
                          $('body');

      let rawPrice = buyBoxContext.find('.priceToPay .a-offscreen').first().text().trim() ||
                     $('#apex_offerDisplay_desktop .priceToPay .a-offscreen').first().text().trim() ||
                     $('#newUnifiedOfferDisplay .priceToPay .a-offscreen').first().text().trim() ||
                     $('#corePriceDisplay_desktop_feature_div .priceToPay .a-offscreen').first().text().trim() ||
                     $('#corePrice_desktop .priceToPay .a-offscreen').first().text().trim() ||
                     $('#corePrice_feature_div .priceToPay .a-offscreen').first().text().trim() ||
                     $('#buyNewSection .a-price .a-offscreen').first().text().trim() ||
                     $('#desktop_buybox .priceToPay .a-offscreen').first().text().trim() ||
                     $('#rightCol .priceToPay .a-offscreen').first().text().trim() ||
                     "";

      const cleanedPrice = await cleanAndNormalizePrice(rawPrice);

      let rawShippingText = $('#mir-layout-DELIVERY_BLOCK-slot-PRIMARY_DELIVERY_MESSAGE_LARGE').text().trim() ||
                            $('#deliveryBlockMessage').text().trim() ||
                            $('#mir-layout-DELIVERY_BLOCK-slot-SECONDARY_DELIVERY_MESSAGE_LARGE').text().trim() || '';

      let amazonBuyboxOwner = buyBoxContext.find('div[tabular-attribute-name="Sold by"] .tabular-buybox-text').first().text().trim() ||
                              buyBoxContext.find('div[tabular-attribute-name="Verkauf durch"] .tabular-buybox-text').first().text().trim() ||
                              buyBoxContext.find('div[tabular-attribute-name="Verzonden en verkocht door"] .tabular-buybox-text').first().text().trim() ||
                              buyBoxContext.find('#sellerProfileTriggerId').first().text().trim() ||
                              buyBoxContext.find('#merchant-info a').first().text().trim() ||
                              "";

      if (!amazonBuyboxOwner) {
        amazonBuyboxOwner = buyBoxContext.find('#merchant-info').first().text().trim();
      }

      console.log(`RESULT => Raw Price text: "${rawPrice}" | Cleaned: "${cleanedPrice}"`);
      console.log(`RESULT => Shipping text: "${rawShippingText.replace(/\s+/g, ' ')}"`);
      console.log(`RESULT => Buybox Owner: "${amazonBuyboxOwner.replace(/\s+/g, ' ').trim()}"`);

    } catch (err) {
      console.error(`Error auditing ${config.domain}:`, err);
    } finally {
      await context.close();
    }
  }

  await browser.close();
})();
