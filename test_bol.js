import { chromium } from 'playwright-extra';
import stealth from 'puppeteer-extra-plugin-stealth';
chromium.use(stealth());

(async () => {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext();
  const page = await context.newPage();
  await page.goto("https://www.bol.com/nl/nl/s/?searchtext=0810127261204", { waitUntil: 'domcontentloaded' });
  const title = await page.title();
  console.log("Title: " + title);
  const links = await page.evaluate(() => Array.from(document.querySelectorAll('a')).map(a => a.href).filter(h => h.includes('/p/')));
  console.log("Links with /p/:", links.length);
  if (links.length > 0) {
      console.log(links[0]);
  }
  const content = await page.content();
  import('fs').then(fs => fs.writeFileSync('bol_search.html', content));
  await browser.close();
})();
