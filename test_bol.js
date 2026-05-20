import { chromium } from 'playwright-extra';
import stealth from 'puppeteer-extra-plugin-stealth';
chromium.use(stealth());

(async () => {
  // Run Chromium in headful mode on Railway to let Akamai’s JS challenge complete.
  // “headless: false” forces a real browser window (even though no display is attached);
  // Railway’s container provides a virtual frame buffer, so this works.
  const browser = await chromium.launch({
    headless: false,
    args: [
      '--headless=new',
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-accelerated-2d-canvas',
      '--disable-gpu',
      '--window-size=1280,720',
      '--disable-features=IsolateOrigins,site-per-process'   // helps bypass some WAF checks
    ]
  });
  const context = await browser.newContext();
  const page = await context.newPage();
  // Load the page and wait for the full network idle state – this gives Akamai’s JS challenge time to finish.
  await page.goto(
    "https://www.bol.com/nl/nl/s/?searchtext=0810127261204",
    { waitUntil: 'networkidle' }
  );
  // Extra short wait for any final redirects or JS that may still be running.
  await page.waitForTimeout(2000);
  const title = await page.title();
  console.log("Title: " + title);
  const links = await page.evaluate(() => Array.from(document.querySelectorAll('a')).map(a => a.href).filter(h => h.includes('/p/')));
  console.log("Links with /p/:", links.length);
  if (links.length > 0) {
      console.log(links[0]);
  }
  // Capture the final HTML **after** the challenge resolves.
  const content = await page.content();
  import('fs').then(fs => fs.writeFileSync('bol_search.html', content));
  // Give the file system a moment to finish writing before closing.
  await page.waitForTimeout(500);
  await browser.close();
})();
