const { chromium } = require('playwright');
const cheerio = require('cheerio');
const { chromium: chromiumExtra } = require('playwright-extra');
const stealth = require('puppeteer-extra-plugin-stealth')();
chromiumExtra.use(stealth);

(async () => {
    const launchOpts = {
      args: ['--no-sandbox','--disable-setuid-sandbox']
    };
    if (process.env.PROXY_SERVER) {
      launchOpts.proxy = {
        server: process.env.PROXY_SERVER,
        username: process.env.PROXY_USERNAME,
        password: process.env.PROXY_PASSWORD
      };
    }
    const browser = await chromiumExtra.launch(launchOpts);
    
    for (const asin of ['B007671GXU', 'B0087OXZFI']) {
        const context = await browser.newContext({ locale: 'en-GB' });
        const page = await context.newPage();
        await page.goto(`https://www.amazon.co.uk/dp/${asin}`);
        await page.waitForTimeout(3000);
        let content = await page.content();
        
        while(content.includes('continue shopping') || content.includes('Type the characters you see in this image')) {
            await context.close();
            await page.waitForTimeout(1000);
            context = await browser.newContext({ locale: 'en-GB' });
            page = await context.newPage();
            await page.route('**/*', route => {
                if (['image', 'stylesheet', 'font'].includes(route.request().resourceType())) return route.abort();
                route.continue();
            });
            await page.goto(`https://www.amazon.co.uk/dp/${asin}`);
            await page.waitForTimeout(3000);
            content = await page.content();
        }

        const $ = cheerio.load(content);
        
        // Exact logic copied from server.ts
        let amazonPrice = $('#corePriceDisplay_desktop_feature_div .priceToPay .a-offscreen').first().text().trim() ||
                          $('#corePrice_desktop .priceToPay .a-offscreen').first().text().trim() ||
                          $('#corePriceDisplay_desktop_feature_div .a-price .a-offscreen').first().text().trim() ||
                          $('#corePrice_desktop .a-price .a-offscreen').first().text().trim() ||
                          $('#buyNew_noncbb .a-price .a-offscreen').first().text().trim() ||
                          $('#buyNewSection .a-price .a-offscreen').first().text().trim() ||
                          $('#desktop_buybox .a-price .a-offscreen').first().text().trim() ||
                          $('#price_inside_buybox').text().trim() ||
                          $('.apex-core-price-identifier .a-offscreen').first().text().trim() ||
                          $('#desktop_buybox .apexPriceToPay .a-offscreen').first().text().trim() ||
                          $('#desktop_buybox .priceToPay .a-offscreen').first().text().trim() ||
                          $('#rightCol .a-price .a-offscreen').first().text().trim() || "";
        amazonPrice = amazonPrice.replace(/\s+/g, ' ').trim();

        let amazonBuyboxOwner = $('div[tabular-attribute-name="Sold by"] .tabular-buybox-text').first().text().trim() ||
                                $('div[tabular-attribute-name="Verkauf durch"] .tabular-buybox-text').first().text().trim() ||
                                $('div[tabular-attribute-name="Shipper / Seller"] .tabular-buybox-text').first().text().trim() ||
                                $('div[tabular-attribute-name="Verzonden en verkocht door"] .tabular-buybox-text').first().text().trim() ||
                                $('div[tabular-attribute-name="Dispatched from and sold by"] .tabular-buybox-text').first().text().trim() ||
                                $('div[offer-display-feature-name="desktop-merchant-info"] .offer-display-feature-text-message').first().text().trim() ||
                                $('#sellerProfileTriggerId').first().text().trim() ||
                                $('#merchant-info a').first().text().trim();

        if (!amazonBuyboxOwner) {
          let mInfo = $('#merchant-info').first().text().toLowerCase();
          if (mInfo.includes('sold by amazon') || mInfo.includes('verkauf durch amazon') || mInfo.includes('dispatched from and sold by amazon') || mInfo.includes('expédié et vendu par amazon')) {
            amazonBuyboxOwner = "Amazon";
          } else {
            amazonBuyboxOwner = $('#desktop_buybox .offer-display-feature-text-message').first().text().trim() || 
                                $('#rightCol .offer-display-feature-text-message').first().text().trim() ||
                                $('#merchant-info').first().text().trim();
          }
        }
        amazonBuyboxOwner = amazonBuyboxOwner.replace(/Sold by\s*:?\s*/gi, '').replace(/Venduto da\s*:?\s*/gi, '').replace(/Verkauf durch\s*:?\s*/gi, '').trim();

        console.log(`ASIN ${asin} => Price: ${amazonPrice}, Buybox: ${amazonBuyboxOwner}`);

        await context.close();
    }
    await browser.close();
})();
