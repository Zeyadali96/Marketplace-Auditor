const { chromium } = require('playwright');
const cheerio = require('cheerio');
const { chromium: chromiumExtra } = require('playwright-extra');
const stealth = require('puppeteer-extra-plugin-stealth')();
chromiumExtra.use(stealth);

(async () => {
    const browser = await chromiumExtra.launch();
    const context = await browser.newContext({
      userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    });
    
    for (const asin of ['B007671GXU']) {
        console.log(`\n--- ASIN: ${asin} ---`);
        const page = await context.newPage();
        await page.goto(`https://www.amazon.co.uk/dp/${asin}`);
        await page.waitForTimeout(3000);
        const content = await page.content();
        const $ = cheerio.load(content);
        
        console.log("Title:", $('#productTitle').text().trim());
        
        console.log("twister price:", $('input#twister-plus-price-data-price').attr('value'));
        console.log("twister price display:", $('input#twister-plus-price-data-price-display').attr('value'));
        
        console.log("accordion row 1 price:", $('#accordionRow .a-price .a-offscreen').first().text().trim());
        console.log("one-time purchase price:", $('#oneTimeBuyBox .a-price .a-offscreen').first().text().trim());
        console.log("corePriceDisplay priceToPay:", $('#corePriceDisplay_desktop_feature_div .priceToPay .a-offscreen').first().text().trim());
        
        console.log("buybox tabular owner 1:", $('div[tabular-attribute-name="Sold by"] .tabular-buybox-text').first().text().trim());
        console.log("buybox tabular owner Shipper / Seller:", $('div[tabular-attribute-name="Shipper / Seller"] .tabular-buybox-text').first().text().trim());
        console.log("buybox tabular owner Verzonden en verkocht door:", $('div[tabular-attribute-name="Verzonden en verkocht door"] .tabular-buybox-text').first().text().trim());
        console.log("merchant info span:", $('#merchant-info span').first().text().trim());
        console.log("merchant info a:", $('#merchant-info a').first().text().trim());
        console.log("seller profile trigger:", $('#sellerProfileTriggerId').text().trim());

        await page.close();
    }
    await browser.close();
})();
