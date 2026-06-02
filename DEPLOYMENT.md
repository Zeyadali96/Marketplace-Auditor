# Deployment Guide: Product Audit Tool

This guide provides comprehensive instructions for deploying the **Product Audit Tool** of the **Marketplace Audit Automation** suite to cloud environments like Railway and Google Cloud Platform (Cloud Run, App Engine).

---

## 📋 Application Overview & Technology Stack

The **Product Audit Tool** is a full-stack, TypeScript-powered web application designed to securely audit live product listings across ten European marketplaces plus the US against master product records in Google Sheets.

- **Frontend**: React + Vite + TypeScript + TailwindCSS
- **Backend**: Express.js + Playwright (Headless Crawler / Scraper)
- **APIs**:
  - Google Gemini API (for WAF-resilient Bol.com fallback extraction)
  - Google Sheets v4 API (for Master list processing and synchronization)
- **Scraping Framework**: Playwright & Cheerio for HTML node traversal and geographic location bypasses.

---

## 🛠️ Required Environment Variables

To run the application, configure the following environmental variables within your hosting provider or local `.env` file:

```env
# Google Sheets API Integration
GOOGLE_SERVICE_ACCOUNT_EMAIL=your-service-account@project.iam.gserviceaccount.com
GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
GOOGLE_SHEETS_ID=your-spreadsheet-id

# Google Gemini AI API Key (Required for Bol.com WAF bypass strategy)
GEMINI_API_KEY=AIzaSy...

# Optional: Proxy configurations (For advanced rate-limit protection)
PROXY_SERVER=
PROXY_USERNAME=
PROXY_PASSWORD=
```

---

## 📦 Deployment Instructions

### Option A: Railway (Highly Recommended)
Railway automatically detects Node.js apps, installs standard headless requirements, and manages dependencies easily.

1. **Connect Repository**: Link your GitHub repository directly to your [Railway Dashboard](https://railway.app).
2. **Configure Variables**: Go to the **Variables** tab and paste your `.env` key-value pairs (`GEMINI_API_KEY`, Google Sheets credentials, etc.).
3. **Automatic Lifecycle**:
   - Railway reads `package.json` and runs `npm install`.
   - The included `postinstall` script triggers `npx playwright install chromium` to fetch standard browser binaries.
   - It executes `npm run build` followed by `npm start`.

---

### Option B: Google Cloud Run (Containerized Scaling)
For highly scalable, serverless container execution:

1. **Create Dockerfile**: Here is a minimal, production-ready `Dockerfile` targeting Node 18:
   ```dockerfile
   FROM node:18-slim

   # Install basic OS libraries required for Playwright Chromium running headlessly
   RUN apt-get update && apt-get install -y \
       libglib2.0-0 \
       libnss3 \
       libnspr4 \
       libatk-1.0-0 \
       libatk-bridge2.0-0 \
       libcups2 \
       libdrm2 \
       libxkbcommon0 \
       libxcomposite1 \
       libxdamage1 \
       libxext6 \
       libxfixes3 \
       librandr2 \
       libgbm1 \
       libasound2 \
       libpangocairo-1.0-0 \
       libpango-1.0-0 \
       libcairo2 \
       && rm -rf /var/lib/apt/lists/*

   WORKDIR /app
   COPY package*.json ./
   RUN npm ci --only=production
   RUN npx playwright install chromium
   COPY . .
   RUN npm run build

   EXPOSE 3000
   CMD ["npm", "run", "start"]
   ```
2. **Deploy to Cloud Run**:
   ```bash
   gcloud run deploy product-audit-tool --source . --port 3000 --allow-unauthenticated
   ```

---

### Option C: Google Cloud App Engine
For standard PaaS deployments using App Engine Standard Environment:

1. **Create `app.yaml`**:
   ```yaml
   runtime: nodejs18
   instance_class: F2 # Recommending at least 512MB RAM for headless scraping
   env_variables:
     GOOGLE_SERVICE_ACCOUNT_EMAIL: "your-service-account"
     GOOGLE_SHEETS_ID: "your-id"
     GEMINI_API_KEY: "your-key"
   ```
2. **Deploy via App SDK**:
   ```bash
   gcloud app deploy
   ```

---

## ⚡ API Endpoint References

### 1. Amazon Listing Audit
- **Endpoint**: `/api/audit/amazon`
- **Method**: `POST`
- **Payload Schema**:
  ```json
  {
    "asin": "B00OLZ9TJ8",
    "marketplace": "amazon.co.uk",
    "masterData": {
      "title": "Product Title",
      "price": "29.99",
      "shipping": "2-3 days",
      "description": "Product description",
      "bullets": ["Feature 1", "Feature 2"],
      "images": ["url1", "url2"]
    }
  }
  ```

### 2. Bol.com Listing Audit
- **Endpoint**: `/api/audit/bol`
- **Method**: `POST`
- **Payload Schema**:
  ```json
  {
    "ean": "9781234567890",
    "masterData": {
      "title": "Book Title",
      "price": "14.99",
      "shipping": "Morgen in huis"
    }
  }
  ```

---

## 🔍 Troubleshooting Checklist

- **Playwright errors inside Docker**: If Playwright reports missing system dependencies or libraries (e.g. `libgbm.so.1`), verify your `Dockerfile` includes the necessary dynamic libraries (apt-get packages shown in Option B).
- **Google Sheets Authorization Error**: Ensure your Google Service Account email has been explicitly shared with **Viewer/Editor** permissions on your specific `GOOGLE_SHEETS_ID` sheet.
- **WAF Blocks on Bol.com**: When running from server-grade IP ranges (e.g., AWS/GCP blocks), direct Chrome scraping is heavily protected. The integrated **Gemini 2-Tier Fallback** solves this; keep your `GEMINI_API_KEY` active to ensure flawless Google Search-grounded audits are leveraged automatically.
