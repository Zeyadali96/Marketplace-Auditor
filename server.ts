import express from 'express';
import { chromium } from 'playwright';
import { chromium as chromiumExtra } from 'playwright-extra';
import stealth from 'puppeteer-extra-plugin-stealth';
import * as cheerio from 'cheerio';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import axios from 'axios';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';
import stringSimilarity from 'string-similarity';
import { GoogleGenAI } from '@google/genai';
import {
  launchAmazonBrowser,
  extractAmazonData,
  parseAmazonContent,
  tryAmazonViaGemini
} from './amazon-audit.js';
import {
  tryBolViaGemini,
  launchBolBrowser,
  goToBolProduct,
  extractBolCatalogue,
  calculateBolShippingDays
} from './bol-audit.js';

chromiumExtra.use(stealth());

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json({ limit: '50mb' }));

// Helper for string similarity
const getSimilarity = (str1: string, str2: string) => {
  if (!str1 || !str2) return 0;
  return stringSimilarity.compareTwoStrings(str1.toLowerCase(), str2.toLowerCase());
};

// 1. Proxy Image API
app.get('/api/proxy-image', async (req, res) => {
  const imageUrl = req.query.url as string;
  if (!imageUrl) return res.status(400).send('No URL provided');
  
  try {
    const response = await axios.get(imageUrl, {
      responseType: 'arraybuffer',
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
        'Referer': 'https://www.amazon.com/'
      }
    });
    const contentType = response.headers['content-type'];
    res.setHeader('Content-Type', contentType);
    res.send(response.data);
  } catch (error) {
    res.status(500).send('Error proxying image');
  }
});

// Helper to normalize price strings to dots
function cleanAndNormalizePrice(priceStr: string): string {
  if (!priceStr) return "";
  let s = priceStr.trim();
  // Remove space between numbers (e.g. "1 250,50" -> "1250,50")
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
  // Select characters and decimals
  const match = s.match(/\d+(\.\d+)?/);
  return match ? match[0] : s.replace(/[^0-9.]/g, '');
}

// 2. Audit Amazon (MODULARIZED WITH FALLBACK)
app.post("/api/audit/amazon", async (req, res) => {
  let browser;
  try {
    const { asin, marketplace, masterData } = req.body;
    const domain = marketplace || 'amazon.com';
    
    console.log(`[AMAZON] Auditing ${asin} on ${domain}...`);

    let liveData = null;
    let config = null;

    try {
      // Launch browser with localization and optional proxy
      const amzResult = await launchAmazonBrowser(domain);
      browser = amzResult.browser;
      const { page, context, config: amzConfig } = amzResult;
      config = amzConfig;

      try {
        // Extract HTML content with location unlock
        const content = await extractAmazonData(page, asin, domain, config);
        
        // Parse the HTML to extract all data using improved selectors
        liveData = parseAmazonContent(content, domain, config);
      } finally {
        await context.close();
      }
    } catch (browserError) {
      console.warn("[AMAZON] Playwright browser extraction failed:", browserError.message);
    }

    // Try Gemini Fallback if Browser was blocked, failed, or returned empty/malformed results
    const isMalformed = !liveData || !liveData.title || !liveData.price || liveData.price === 'N/A';
    if (isMalformed) {
      console.log(`[AMAZON] Live data was empty/malformed or browser was blocked. Attempting Gemini Google Search fallback...`);
      const geminiData = await tryAmazonViaGemini(asin, domain);
      if (geminiData) {
        liveData = geminiData;
        console.log(`[AMAZON] Gemini Google Search Fallback succeeded!`);
      }
    }

    if (!liveData) {
      throw new Error("Unable to extract Amazon data via browser or Gemini fallback.");
    }

    // Extend liveData with standard fields
    liveData.currency = liveData.currency || (config ? config.currency : 'USD');
    liveData.buyboxOwner = liveData.buyboxOwner || 'N/A';

    // Perform audit comparison
    const auditResult = await performAudit(masterData, liveData, 'amazon', domain);
    res.json({ liveData, auditResult });

  } catch (error) {
    console.error("[AMAZON] Audit Error:", error);
    res.status(500).json({ error: error.message });
  } finally {
    if (browser) await browser.close();
  }
});



// Helper for performAudit
async function performAudit(master: any, live: any, mode: string, domain?: string) {
  const result: any = {
    title: { master: master.title, live: live.title, similarity: getSimilarity(master.title, live.title), match: false },
    description: { 
      master: master.description, 
      live: live.description || (live.hasAPlus ? "A+ Content Detected (Standard description missing)" : ""), 
      similarity: 0, 
      match: false, 
      isAPlus: live.hasAPlus 
    },
    bullets: [],
    price: { master: master.price, live: live.price, match: false },
    shipping: { master: master.shipping, live: live.shipping, match: false, days: live.shippingDays },
    images: { master: master.images, live: live.images, match: false },
    variations: { match: live.variations > 1 }
  };

  result.description.similarity = getSimilarity(master.description || "", live.description || "");

  if (result.title.similarity > 0.8) result.title.match = true;
  if (result.description.similarity > 0.6 || live.hasAPlus) result.description.match = true;
  
  // Bullets match
  if (master.bullets && Array.isArray(master.bullets)) {
    master.bullets.forEach((mb: string) => {
      let bestSim = 0;
      let bestLive = "";
      if (live.bullets && Array.isArray(live.bullets)) {
        live.bullets.forEach((lb: string) => {
          const sim = getSimilarity(mb, lb);
          if (sim > bestSim) {
            bestSim = sim;
            bestLive = lb;
          }
        });
      }
      result.bullets.push({ master: mb, live: bestLive, similarity: bestSim, match: bestSim > 0.7 });
    });
  }

  // Price match (fuzzy)
  const masterPriceNum = parseFloat(String(master.price || "").replace(/[^0-9.]/g, '')) || 0;
  const livePriceNum = parseFloat(String(live.price || "").replace(/[^0-9.]/g, '')) || 0;
  if (masterPriceNum > 0 && Math.abs(masterPriceNum - livePriceNum) < 1.0) result.price.match = true;

  if (live.images && live.images.length >= (master.images?.length || 1)) result.images.match = true;

  // Score calculation
  let scoreValue = 0;
  if (mode === 'amazon') {
    if (result.title.match) scoreValue += 30;
    if (result.description.match) scoreValue += 30;
    const bulletMatchCount = (result.bullets || []).filter((b: any) => b.match).length;
    scoreValue += Math.min(bulletMatchCount * 8, 40);
  } else {
    if (result.title.match) scoreValue += 50;
    if (result.description.match) scoreValue += 50;
  }
  result.score = scoreValue;

  return result;
}

function getScoreGrade(score: number): string {
  if (score > 70) return "excellent";
  if (score >= 50) return "acceptable";
  return "Needs improvement";
}

app.post("/api/audit/bol", async (req, res) => {
  let browser;
  try {
    const { ean, masterData } = req.body;
    if (!ean) throw new Error('Missing "ean" in request body');

    let data: any = null;
    let dataSource = 'browser';

    // ── Strategy 1: Gemini Google Search Grounding ─────────────────────────
    console.log('[BOL] Trying Strategy 1: Gemini Google Search Grounding...');
    data = await tryBolViaGemini(ean);
    if (data) {
      console.log('[BOL] Strategy 1 succeeded via Gemini.');
      dataSource = 'gemini';
    }

    // ── Strategy 2: Playwright stealth browser (hardened) ─────────────────────
    // Add delay before browser strategy to reduce rapid-fire requests
    if (!data) {
      console.log('[BOL] Adding 2-second delay before browser strategy to avoid WAF detection...');
      await new Promise(resolve => setTimeout(resolve, 2000));
      console.log('[BOL] Trying Strategy 2: Playwright stealth browser...');

      const bolResult = await launchBolBrowser();
      browser = bolResult.browser;
      const { page, context } = bolResult;

      try {
        await goToBolProduct(page, ean);
        data = await extractBolCatalogue(page);
        dataSource = 'browser';
      } finally {
        await context.close();
      }
    }

    const bolShippingDays = calculateBolShippingDays(data.shipping || '');

    const liveData = {
      title: data.title || '',
      price: data.price || 'N/A',
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

    const auditResult = await performAudit(masterData, liveData, 'bol');
    res.json({ liveData, auditResult });

  } catch (error: any) {
    res.status(500).json({ error: error.message });
  } finally {
    if (browser) await browser.close();
  }
});

// 4. Sheets APIs
app.post("/api/sheets/fetch", async (req, res) => {
  try {
    const { sheetId, mode, marketplace, sheetName: requestedSheetName } = req.body;
    const auth = new JWT({
      email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const doc = new GoogleSpreadsheet(sheetId, auth);
    await doc.loadInfo();

    const targetTabName = requestedSheetName || (mode === "bol" ? "Product data" : "Amazon Data");
    const sheet = doc.sheetsByTitle[targetTabName];
    if (!sheet) {
      throw new Error(`Sheet tab "${targetTabName}" not found in the spreadsheet.`);
    }

    // Amazon Data tab has a category group row in row 1;
    // the real column headers (ASIN, AMZ title DE, …) are in row 2.
    // Product data tab has normal headers in row 1.
    if (targetTabName === 'Amazon Data') {
      await sheet.loadHeaderRow(2);
    }

    const rows = await sheet.getRows();
    const data = rows.map(row => row.toObject());
    
    res.json({ data });
  } catch (error: any) {
    console.error("Fetch Sheet Error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/sheets/save-audit", async (req, res) => {
  try {
    const { sheetId, mode, identifier, auditResult, liveData, masterData } = req.body;
    
    if (!sheetId) throw new Error("Missing sheetId in request body");

    const auth = new JWT({
      email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const doc = new GoogleSpreadsheet(sheetId, auth);
    await doc.loadInfo();
    
    const targetSheetName = mode === "bol" ? "Bol QC Results" : "Amazon QC Results";
    // Case-insensitive search for the sheet
    const sheet = doc.sheetsByTitle[targetSheetName] || 
                  doc.sheetsByIndex.find(s => s.title.toLowerCase().trim() === targetSheetName.toLowerCase());
    
    if (!sheet) {
      const availableSheets = doc.sheetsByIndex.map(s => `"${s.title}"`).join(', ');
      throw new Error(`Sheet "${targetSheetName}" not found. Available sheets: ${availableSheets}`);
    }

    // Always override everything and start from row 2
    await sheet.clearRows();

    // Mapping according to instructions
    const bulletMatchCount = (auditResult && Array.isArray(auditResult.bullets)) ? auditResult.bullets.filter((b: any) => b.match).length : 0;
    const bulletMatchText = bulletMatchCount > 0 ? 'Yes' : 'No';

    const timestamp = new Date().toLocaleString();
    const resultRow: any = {
      'Date': timestamp,
      'date': timestamp,
      'EAN': identifier,
      'ean': identifier,
      'ASIN': identifier,
      'asin': identifier,
      'ASIN/EAN': identifier,
      'asin/ean': identifier,
      'Identifier': identifier,
      'identifier': identifier,
      'Mode': mode,
      'mode': mode,
      'Title Match': auditResult?.title?.match ? 'Yes' : 'No',
      'Description Match': liveData?.hasAPlus ? 'A+ content available' : (auditResult?.description?.match ? 'Yes' : 'No'),
      'Bullet Points Match': bulletMatchText,
      'Variation': auditResult?.variations?.match ? 'Yes' : 'No',
      'A+ Content': liveData?.hasAPlus ? 'Yes' : 'No',
      'Buybox Owner': liveData?.buyboxOwner || 'N/A',
      'Score': auditResult?.score ?? 0,
      'Score Grade': getScoreGrade(auditResult?.score ?? 0),
      'Price Live': liveData?.price || 'N/A',
      'price live': liveData?.price || 'N/A',
      'Price': liveData?.price || 'N/A',
      'price': liveData?.price || 'N/A',
      'Shipping Live (Days)': liveData?.shippingDays || 'N/A',
      'Shipping Time': liveData?.shippingDays || 'N/A',
      'shipping time': liveData?.shippingDays || 'N/A',
      'Notes': liveData?.hasAPlus ? 'A+ content available' : ''
    };

    const pfx = mode === 'bol' ? 'Bol' : 'Amazon';
    const shortPfx = mode === 'bol' ? 'BOL' : 'AMZ';

    // Images mapping - User wants "each relevant column based on number image"
    if (masterData && masterData.images && Array.isArray(masterData.images)) {
      masterData.images.slice(0, 10).forEach((url: string, i: number) => {
        const formula = url ? `=IMAGE("${url}")` : '';
        const index = i + 1;
        resultRow[`${pfx} Master Image ${index}`] = formula;
        resultRow[`${shortPfx} Master Image ${index}`] = formula;
        resultRow[`Master Image ${index}`] = formula;
        resultRow[`${shortPfx} IMG ${index}`] = formula;
        resultRow[`image ${mode} data ${index}`] = formula;
        resultRow[`Image ${pfx} Data ${index}`] = formula;
        
        resultRow[`Amazon Master Image ${index}`] = formula;
      });
    }

    if (liveData && liveData.images && Array.isArray(liveData.images)) {
      const sliceStart = mode === 'bol' ? 0 : 1;
      liveData.images.slice(sliceStart, sliceStart + 10).forEach((url: string, i: number) => {
        const formula = url ? `=IMAGE("${url}")` : '';
        const index = i + 1;
        resultRow[`${pfx} Live Image ${index}`] = formula;
        resultRow[`${shortPfx} Live Image ${index}`] = formula;
        resultRow[`Live Image ${index}`] = formula;
        resultRow[`${shortPfx} L IMG ${index}`] = formula;
        resultRow[`image ${mode} live ${index}`] = formula;
        resultRow[`Image ${pfx} Live ${index}`] = formula;
        
        resultRow[`Amazon Live Image ${index}`] = formula;
      });
    }

    // Search for existing row to override removed based on prompt instruction to always overwrite starting from row 2
    await sheet.loadHeaderRow();
    const headers = sheet.headerValues;
    const findHeader = (target: string) => {
      return headers.find(h => h.toLowerCase().trim() === target.toLowerCase().trim());
    };

    // Add new row - mapping to actual sheet headers
    const rowToSave: any = {};
    Object.keys(resultRow).forEach(key => {
      const matchingHeader = findHeader(key);
      if (matchingHeader) {
        rowToSave[matchingHeader] = resultRow[key];
      }
    });
    
    // If we couldn't find some headers in the rowToSave, but we have them in resultRow and they are standard, 
    // they might not be added if headers don't exist. But here we assume sheet has them.
    if (Object.keys(rowToSave).length > 0) {
      await sheet.addRow(rowToSave);
    } else {
      // Fallback to original object if no headers matched (unlikely but safe)
      await sheet.addRow(resultRow);
    }
    
    res.json({ success: true });
  } catch (error: any) {
    console.error("Save Audit Error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/sheets/batch-save-audit", async (req, res) => {
  try {
    const { sheetId, mode, audits } = req.body;
    if (!audits || !Array.isArray(audits) || audits.length === 0) {
      return res.json({ success: true, message: "No audits to save." });
    }
    
    if (!sheetId) throw new Error("Missing sheetId in request body");

    const auth = new JWT({
      email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const doc = new GoogleSpreadsheet(sheetId, auth);
    await doc.loadInfo();

    const targetSheetName = mode === "bol" ? "Bol QC Results" : "Amazon QC Results";
    const sheet = doc.sheetsByTitle[targetSheetName] || 
                  doc.sheetsByIndex.find(s => s.title.toLowerCase().trim() === targetSheetName.toLowerCase());

    if (!sheet) {
      throw new Error(`Sheet "${targetSheetName}" not found.`);
    }

    // Clear existing rows to start automatically from row 2
    await sheet.clearRows();
    await sheet.loadHeaderRow();
    
    const rowsToAdd = audits.map((audit: any) => {
      const { identifier, auditResult, liveData, masterData } = audit;
      const bulletMatchCount = (auditResult && Array.isArray(auditResult.bullets)) ? auditResult.bullets.filter((b: any) => b.match).length : 0;
      const bulletMatchText = bulletMatchCount > 0 ? 'Yes' : 'No';
      const timestamp = new Date().toLocaleString();

      const row: any = {
        'Date': timestamp,
        'date': timestamp,
        'EAN': identifier,
        'ean': identifier,
        'ASIN': identifier,
        'asin': identifier,
        'ASIN/EAN': identifier,
        'asin/ean': identifier,
        'Identifier': identifier,
        'identifier': identifier,
        'Mode': mode,
        'mode': mode,
        'Title Match': auditResult?.title?.match ? 'Yes' : 'No',
        'Description Match': liveData?.hasAPlus ? 'A+ content available' : (auditResult?.description?.match ? 'Yes' : 'No'),
        'Bullet Points Match': bulletMatchText,
        'Variation': auditResult?.variations?.match ? 'Yes' : 'No',
        'A+ Content': liveData?.hasAPlus ? 'Yes' : 'No',
        'Buybox Owner': liveData?.buyboxOwner || 'N/A',
        'Score': auditResult?.score ?? 0,
        'Score Grade': getScoreGrade(auditResult?.score ?? 0),
        'Price Live': liveData?.price || 'N/A',
        'price live': liveData?.price || 'N/A',
        'Price': liveData?.price || 'N/A',
        'price': liveData?.price || 'N/A',
        'Shipping Live (Days)': liveData?.shippingDays || 'N/A',
        'Shipping Time': liveData?.shippingDays || 'N/A',
        'shipping time': liveData?.shippingDays || 'N/A',
        'Notes': liveData?.hasAPlus ? 'A+ content available' : ''
      };

      const pfx = mode === 'bol' ? 'Bol' : 'Amazon';
      const shortPfx = mode === 'bol' ? 'BOL' : 'AMZ';

      if (masterData && Array.isArray(masterData.images)) {
        masterData.images.slice(0, 10).forEach((url: string, i: number) => {
          const formula = url ? `=IMAGE("${url}")` : '';
          row[`${pfx} Master Image ${i + 1}`] = formula;
          row[`${shortPfx} Master Image ${i + 1}`] = formula;
          row[`Master Image ${i + 1}`] = formula;
          row[`${shortPfx} IMG ${i + 1}`] = formula;
          row[`image ${mode} data ${i + 1}`] = formula;
          row[`Image ${pfx} Data ${i + 1}`] = formula;
          
          row[`Amazon Master Image ${i + 1}`] = formula; 
        });
      }

      if (liveData && Array.isArray(liveData.images)) {
        const sliceStart = mode === 'bol' ? 0 : 1;
        liveData.images.slice(sliceStart, sliceStart + 10).forEach((url: string, i: number) => {
          const formula = url ? `=IMAGE("${url}")` : '';
          row[`${pfx} Live Image ${i + 1}`] = formula;
          row[`${shortPfx} Live Image ${i + 1}`] = formula;
          row[`Live Image ${i + 1}`] = formula;
          row[`${shortPfx} L IMG ${i + 1}`] = formula;
          row[`image ${mode} live ${i + 1}`] = formula;
          row[`Image ${pfx} Live ${i + 1}`] = formula;
          
          row[`Amazon Live Image ${i + 1}`] = formula; 
        });
      }
      return row;
    });

    await sheet.addRows(rowsToAdd);
    res.json({ success: true, count: rowsToAdd.length });
  } catch (error: any) {
    console.error("Batch Save Audit Error:", error);
    res.status(500).json({ error: error.message });
  }
});

// Vite middleware for development
if (process.env.NODE_ENV !== "production") {
  const { createServer: createViteServer } = await import('vite');
  const vite = await createViteServer({
    server: { middlewareMode: true },
    appType: "spa",
  });
  app.use(vite.middlewares);
} else {
  const distPath = path.join(process.cwd(), 'dist');
  app.use(express.static(distPath));
  app.get('*', (req, res) => {
    res.sendFile(path.join(distPath, 'index.html'));
  });
}

const PORT = process.env.PORT || 3000;
app.listen(Number(PORT), "0.0.0.0", () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
