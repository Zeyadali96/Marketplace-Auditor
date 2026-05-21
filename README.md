# Product Audit Tool

A full-stack TypeScript, React, and Node.js application that audits product listings on **Amazon** and **Bol.com** against master product data stored securely in **Google Sheets**.

---

## 🚀 Core Features

- **Multi-Market Amazon Audits**: Fast and reliable scraping with custom selectors for European Amazon regions (`.co.uk`, `.de`, `.fr`, `.it`, `.es`, `.nl`, `.pl`, `.se`, `.com.be`) and US (`.com`). It extracts live price, shipping, buybox owner, bullets, images, and variations.
- **WAF-Resilient Bol.com Audits**: High-success rate audits using a smart, resilient 2-tier fallback strategy:
  1. **Gemini URL Context API**: Uses Google's API servers to extract live catalog items without scraping directly, completely bypassing Bol's browser/WAF blocks.
  2. **Playwright Stealth Browser**: A fallback utilizing stealth-evading browser automation, complete with randomized screen profiles, pre-injected cookies to skip consent interfaces, and a 2-second rate delay to avoid rapid-fire bot blocks.
- **Google Sheets Integration**: Automatically pulls target master records (title, target price, target shipping time) from your Google Sheets, compares them instantly using custom score calculations, and persists detailed results back.
- **Interactive Validation Dashboard**: Modern frontend showing visual comparison cards, match indicators (Pass/Fail) for Pricing, Buybox, and Shipping time, along with overall status grade metrics.

---

## 🛠️ Required Environment Variables (`.env`)

Create a `.env` file in the root directory for local development (do not commit actual values to git):

```env
# Google Sheets Integration
GOOGLE_SERVICE_ACCOUNT_EMAIL=your-service-account@project.iam.gserviceaccount.com
GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\nYourSecretPrivateKey...\n-----END PRIVATE KEY-----\n"
GOOGLE_SHEETS_ID=your-google-sheets-spreadsheet-id

# Gemini AI API Key (Required for Bol.com Gemini extraction strategy)
GEMINI_API_KEY=AIzaSy...

# Optional Proxies (for advanced Playwright rate-limit evasion)
PROXY_SERVER=
PROXY_USERNAME=
PROXY_PASSWORD=
```

---

## 💻 Local Setup & Development

### 1. Prerequisite
Ensure you have **Node.js 18+** installed on your system.

### 2. Install Dependencies & Playwright Browsers
To install the workspace dependencies and automatically download the headless Chromium binaries required for Playwright:
```bash
npm install
```
*(This triggers the `postinstall` script which runs `npm run install-browsers` to fetch Chromium)*

### 3. Run the Development Server
```bash
npm run dev
```
Starts Vite + the Express API backend on port `3000`. Open your browser to `http://localhost:3000` to start using the tool locally.

---

## 📦 Production Deployment

### 1. GitHub CI/CD Configuration
Our pre-configured pipeline inside `.github/workflows/deploy.yml` automatically verifies build integrity on any main/master push. Make sure to define these **GitHub Repository Secrets**:
* `GEMINI_API_KEY`
* `GOOGLE_SERVICE_ACCOUNT_EMAIL`
* `GOOGLE_PRIVATE_KEY`
* `GOOGLE_SHEETS_ID`

### 2. Build and Start Commands
To containerize or deploy the app to hosts like Render, Railway, or VPS:
```bash
# Build the React frontend production assets
npm run build

# Start the full-stack server
npm run start
```

---

## 🛠️ Troubleshooting Common Issues

* **"Missing raw private key / certificate block" on local service account setup**: Ensure the `GOOGLE_PRIVATE_KEY` uses literal escape sequences `\n` or the raw multiline blocks correctly formatted for the Node.js `google-auth-library` scope.
* **"WAF BLOCKED" on Bol.com Browser fallback**: This indicates a temporary anti-scrape block. Rest assured that the Gemini URL Context fallback is designed to bypass this with Google API resources first.
* **Playwright fails to boot browser inside container**: Ensure standard Linux dependencies are present or use the designated `mcr.microsoft.com/playwright:vX.Y.Z` base images.
