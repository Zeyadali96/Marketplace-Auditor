const { chromium } = require('playwright-extra');
const stealth = require('puppeteer-extra-plugin-stealth');
chromium.use(stealth());

(async () => {
  try {
    const browser = await chromium.launch({ 
      headless: false, 
      args: ['--headless=new', '--no-sandbox'] 
    });
    console.log("Launched!");
    const context = await browser.newContext({
      userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36',
      viewport: { width: 1920, height: 1080 },
      locale: 'nl-NL',
      extraHTTPHeaders: {
        'Accept-Language': 'nl-NL,nl;q=0.9,en-US;q=0.8,en;q=0.7',
        'sec-ch-ua': '"Chromium";v="136", "Google Chrome";v="136", "Not-A.Brand";v="99"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'upgrade-insecure-requests': '1'
      }
    });
    const page = await context.newPage();
    await page.goto("https://www.bol.com/nl/nl/s/?searchtext=8710755913251", { waitUntil: 'networkidle' });
    const content = await page.content();
    console.log(content.includes('WAF_BLOCKED') ? 'BLOCKED' : 'SUCCESS');
    await browser.close();
  } catch (e) {
    console.error(e);
  }
})();
