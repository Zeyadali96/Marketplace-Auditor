import { chromium } from 'playwright-extra';
import stealth from 'puppeteer-extra-plugin-stealth';
chromium.use(stealth());

(async () => {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext();
  const page = await context.newPage();
  const ean = "0810127261204";
  await page.goto("https://html.duckduckgo.com/html/?q=site:bol.com+" + ean, { waitUntil: 'domcontentloaded' });
  const links = await page.evaluate(() => {
     return Array.from(document.querySelectorAll('a.result__snippet')).map(a => {
        try {
           const url = new URL(a.href);
           const uddg = url.searchParams.get('uddg');
           if (uddg) return decodeURIComponent(uddg);
        } catch (e) {}
        return a.href;
     });
  });
  console.log("Links found:", links);
  await browser.close();
})();
