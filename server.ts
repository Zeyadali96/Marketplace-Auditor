import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import axios from 'axios';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';
import { getScoreGrade } from './helpers.ts';
import { auditAmazon } from './amazon-audit.ts';
import { auditBol } from './bol-audit.ts';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json({ limit: '50mb' }));

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

// 2. Audit Amazon
app.post("/api/audit/amazon", async (req, res) => {
  try {
    const { asin, marketplace, masterData } = req.body;
    if (!asin) throw new Error('Missing "asin" in request body');
    if (!masterData) throw new Error('Missing "masterData" in request body');

    console.log(`[API ROUTE] Handing Amazon audit request for ASIN: ${asin} on marketplace: ${marketplace || 'amazon.com'}`);
    const result = await auditAmazon(asin, marketplace, masterData);
    res.json(result);
  } catch (error: any) {
    console.error("Amazon Audit Route Error:", error);
    res.status(500).json({ error: error.message });
  }
});

// 3. Audit Bol.com
app.post("/api/audit/bol", async (req, res) => {
  try {
    const { ean, masterData } = req.body;
    if (!ean) throw new Error('Missing "ean" in request body');
    if (!masterData) throw new Error('Missing "masterData" in request body');

    console.log(`[API ROUTE] Handing Bol.com audit request for EAN: ${ean}`);
    const result = await auditBol(ean, masterData);
    res.json(result);
  } catch (error: any) {
    console.error("Bol Audit Route Error:", error);
    res.status(500).json({ error: error.message });
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
