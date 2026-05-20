const { chromium } = require('playwright-extra');
const stealth = require('puppeteer-extra-plugin-stealth')();
chromium.use(stealth);

(async () => {
    const launchOpts = { args: ['--no-sandbox','--disable-setuid-sandbox'] };
    if (process.env.PROXY_SERVER) {
      launchOpts.proxy = {
        server: process.env.PROXY_SERVER,
        username: process.env.PROXY_USERNAME,
        password: process.env.PROXY_PASSWORD
      };
    }
    const browser = await chromium.launch(launchOpts);
    const context = await browser.newContext({ locale: 'en-GB' });
    const page = await context.newPage();
    await page.goto("https://www.amazon.co.uk/dp/B007671GXU");
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
        await page.goto("https://www.amazon.co.uk/dp/B007671GXU");
        await page.waitForTimeout(3000);
        content = await page.content();
    }
    const $ = require('cheerio').load(content);
    
    console.log("ALL PRICES in body:");
    $('body .a-price').each((i, el) => {
        let text = $(el).children('.a-offscreen').first().text().trim();
        let parents = $(el).parents().map((i, p) => $(p).attr('id') || $(p).attr('class') || '').get().filter(Boolean).slice(0, 3).join(' < ');
        console.log(parents, '||', text);
    });


    await browser.close();
})();
