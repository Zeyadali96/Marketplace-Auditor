/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import { 
  Search, 
  Settings, 
  CheckCircle2, 
  XCircle, 
  X,
  AlertCircle, 
  ExternalLink, 
  RefreshCw, 
  Trash2,
  Play,
  Image as ImageIcon,
  ChevronRight,
  ChevronDown,
  Database,
  ShoppingBag,
  Info
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

type AuditMode = 'amazon' | 'bol';

export default function App() {
  const [sheetId, setSheetId] = useState('1V4lNf30SlBwczSvGX9rfn5eWFH2AvMO4TqMHAHalS7s');
  const [sheetName, setSheetName] = useState('Amazon Data');
  const [mode, setMode] = useState<AuditMode>('amazon');
  const [countryCode, setCountryCode] = useState('DE');
  const [marketplace, setMarketplace] = useState('amazon.de');
  const [data, setData] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  const [auditing, setAuditing] = useState<string | null>(null);
  const [savingRow, setSavingRow] = useState<number | null>(null);
  const [auditResults, setAuditResults] = useState<Record<string, any>>({});
  const [selectedRow, setSelectedRow] = useState<number | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);
  const [range, setRange] = useState({ start: 1, end: 10 });
  const [galleryModal, setGalleryModal] = useState<{ images: string[], title: string } | null>(null);

  const marketplaces = [
    { label: 'United Kingdom', value: 'amazon.co.uk', code: 'UK' },
    { label: 'Germany', value: 'amazon.de', code: 'DE' },
    { label: 'France', value: 'amazon.fr', code: 'FR' },
    { label: 'Italy', value: 'amazon.it', code: 'IT' },
    { label: 'Spain', value: 'amazon.es', code: 'ES' },
    { label: 'Netherlands', value: 'amazon.nl', code: 'NL' },
    { label: 'Belgium', value: 'amazon.com.be', code: 'BE' },
    { label: 'Sweden', value: 'amazon.se', code: 'SE' },
    { label: 'Poland', value: 'amazon.pl', code: 'PL' },
    { label: 'United States', value: 'amazon.com', code: 'US' },
  ];

  const getVal = (row: any, ...keys: string[]) => {
    const rowKeys = Object.keys(row);
    const normalize = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '');
    
    for (const key of keys) {
      const normalizedKey = normalize(key);
      const foundKey = rowKeys.find(rk => normalize(rk) === normalizedKey);
      if (foundKey && row[foundKey] !== undefined && row[foundKey] !== null) {
        return row[foundKey].toString().trim();
      }
    }
    return '';
  };

  const getProxiedUrl = (url: string) => {
    if (!url) return '';
    // Proxy all external images to avoid CSP/Referrer issues in iframe
    if (url.startsWith('http') || url.startsWith('//')) {
      return `/api/proxy-image?url=${encodeURIComponent(url.startsWith('//') ? 'https:' + url : url)}`;
    }
    return url;
  };

  const fetchSheetData = async () => {
    if (!sheetId) return;
    setLoading(true);
    setError(null);
    try {
      const response = await fetch('/api/sheets/fetch', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sheetId, sheetName })
      });
      const result = await response.json();
      if (result.error) {
        const errorMsg = result.details ? `${result.error}\n\nDetails: ${result.details}` : result.error;
        throw new Error(errorMsg);
      }
      setData(result.data);
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const [isClearing, setIsClearing] = useState(false);

  const clearSheet = async () => {
    setIsClearing(true);
    try {
      const resp = await fetch('/api/sheets/clear-sheet', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ mode })
      });
      const res = await resp.json();
      if (res.error) throw new Error(res.error);
    } catch (err: any) {
      setError(`Failed to clear sheet: ${err.message}`);
    } finally {
      setIsClearing(false);
    }
  };

  const runRangeAudit = async () => {
    // Stage 1: Clear (Wipe Protocol Step 1 & 2)
    await clearSheet();

    // Stage 2: Audit Loop (Wipe Protocol Step 3)
    const startIdx = Math.max(0, range.start - 1);
    const endIdx = Math.min(data.length - 1, range.end - 1);
    
    const resultsToBatch = [];
    
    for (let i = startIdx; i <= endIdx; i++) {
      const result = await runAudit(i, true); // true = skip individual save
      if (result) {
        resultsToBatch.push(result);
      }
    }

    // Stage 2: Batch Save (Wipe once, Write all)
    if (resultsToBatch.length > 0) {
      setSavingRow(-1); // Indicator for batch saving
      try {
        const batchSaveResp = await fetch('/api/sheets/batch-save-audit', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            mode,
            audits: resultsToBatch
          })
        });
        const batchSaveResult = await batchSaveResp.json();
        if (batchSaveResult.error) throw new Error(batchSaveResult.error);
      } catch (err: any) {
        setError(`Batch Save Failed: ${err.message}`);
      } finally {
        setSavingRow(null);
      }
    }
  };

  const exportResults = () => {
    const csvRows = [];
    const headers = ['Row', 'Identifier', 'Title Match', 'Description Match', 'Bullet Match %', 'Price', 'Shipping'];
    csvRows.push(headers.join(','));

    Object.entries(auditResults).forEach(([idx, res]: any) => {
      const row = data[parseInt(idx)];
      const id = mode === 'amazon' ? row.ASIN : row.EAN;
      const bulletMatch = res.auditResult.bullets.filter((b: any) => b.match).length / (res.auditResult.bullets.length || 1);
      
      const line = [
        parseInt(idx) + 1,
        id,
        res.auditResult.title.match ? 'YES' : 'NO',
        res.auditResult.description.match ? 'YES' : 'NO',
        Math.round(bulletMatch * 100) + '%',
        res.auditResult.price.live,
        res.auditResult.shipping.live
      ];
      csvRows.push(line.join(','));
    });

    const blob = new Blob([csvRows.join('\n')], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.setAttribute('hidden', '');
    a.setAttribute('href', url);
    a.setAttribute('download', `audit_results_${new Date().toISOString()}.csv`);
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };
  const runAudit = async (rowIndex: number, skipSave = false) => {
    const row = data[rowIndex];
    const asin = getVal(row, 'ASIN', 'asin');
    const ean = getVal(row, 'EAN', 'ean');
    const identifier = mode === 'amazon' ? asin : ean;
    
    if (!identifier) {
      setError(`Missing ${mode === 'amazon' ? 'ASIN' : 'EAN'} for row ${rowIndex + 1}`);
      return null;
    }

    // Extract language code from marketplace (e.g., amazon.de -> DE)
    const langCode = mode === 'amazon' ? countryCode : 'NL';
    const suffix = langCode ? ` ${langCode}` : '';

    const masterData = {
      title: mode === 'bol' 
        ? getVal(row, 'Bol Title', 'Title NL', 'title nl')
        : getVal(row, 'Amazon Title', `AMZ title${suffix}`, `Title${suffix}`, `Amazon Title ${langCode}`),
      description: mode === 'bol'
        ? getVal(row, 'Bol Description', 'Body NL', 'body nl')
        : getVal(row, 'Amazon Description', `AMZ body${suffix}`, `Description${suffix}`, `Amazon Description ${langCode}`),
      bullets: [
        getVal(row, 'Bullet point 1', `Bullet point 1${suffix}`),
        getVal(row, 'Bullet point 2', `Bullet point 2${suffix}`),
        getVal(row, 'Bullet point 3', `Bullet point 3${suffix}`),
        getVal(row, 'Bullet point 4', `Bullet point 4${suffix}`),
        getVal(row, 'Bullet point 5', `Bullet point 5${suffix}`)
      ].filter(Boolean),
      images: mode === 'bol'
        ? [
            getVal(row, 'Bol IMG 1', 'Image NL'), getVal(row, 'Bol IMG 2'), getVal(row, 'Bol IMG 3'),
            getVal(row, 'Bol IMG 4'), getVal(row, 'Bol IMG 5'), getVal(row, 'Bol IMG 6'),
            getVal(row, 'Bol IMG 7'), getVal(row, 'Bol IMG 8'), getVal(row, 'Bol IMG 9'),
            getVal(row, 'Bol IMG 10')
          ].filter(Boolean)
        : [
            getVal(row, 'AMZ IMG 1'), getVal(row, 'AMZ IMG 2'), getVal(row, 'AMZ IMG 3'),
            getVal(row, 'AMZ IMG 4'), getVal(row, 'AMZ IMG 5'), getVal(row, 'AMZ IMG 6'),
            getVal(row, `AMZ IMG 1${suffix}`), getVal(row, `AMZ IMG 2${suffix}`)
          ].filter(Boolean),
      price: getVal(row, 'Price', 'price'),
      shipping: getVal(row, 'Shipping', 'shipping', 'Shipping Time', `Shipping ${langCode}`, `Shipping Time ${langCode}`),
      variations: getVal(row, 'Variations', 'variations') === 'Yes' || getVal(row, 'Variations', 'variations') === true || getVal(row, 'Variations', 'variations') === 'TRUE',
      hasAPlus: getVal(row, 'APlus', 'aplus', 'A+ Content') === 'Yes' || getVal(row, 'APlus', 'aplus', 'A+ Content') === true || getVal(row, 'APlus', 'aplus', 'A+ Content') === 'TRUE'
    };

    setAuditing(rowIndex.toString());
    try {
      const response = await fetch(`/api/audit/${mode}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          [mode === 'amazon' ? 'asin' : 'ean']: identifier,
          marketplace: mode === 'amazon' ? marketplace : undefined,
          masterData
        })
      });
      const result = await response.json();
      if (result.error) {
        const errorMsg = result.details ? `${result.error}\n\nDetails: ${result.details}` : result.error;
        throw new Error(errorMsg);
      }
      
      setAuditResults(prev => ({
        ...prev,
        [rowIndex]: result
      }));

      const auditPayload = {
        mode,
        identifier,
        marketplace,
        auditResult: result.auditResult,
        masterData
      };

      if (!skipSave) {
        // Automatically Save to QC Automation Sheet (Singular Wipe)
        setSavingRow(rowIndex);
        const saveResponse = await fetch('/api/sheets/save-audit', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(auditPayload)
        });
        const saveResult = await saveResponse.json();
        if (saveResult.error) {
          throw new Error(`Failed to save to sheet: ${saveResult.error}`);
        }
      }

      return auditPayload;

    } catch (err: any) {
      setError(err.message);
      return null;
    } finally {
      setAuditing(null);
      if (!skipSave) setSavingRow(null);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="bg-indigo-600 p-2 rounded-lg">
              <ShoppingBag className="text-white w-5 h-5" />
            </div>
            <h1 className="text-xl font-bold tracking-tight">Marketplace Auditor</h1>
          </div>
          
          <div className="flex items-center gap-4">
            {mode === 'amazon' && (
              <select 
                value={marketplace}
                onChange={(e) => {
                  const m = marketplaces.find(mp => mp.value === e.target.value);
                  setMarketplace(e.target.value);
                  if (m) setCountryCode(m.code);
                }}
                className="bg-slate-100 border-none rounded-lg px-3 py-1.5 text-sm font-medium text-slate-700 outline-none focus:ring-2 focus:ring-indigo-500"
              >
                {marketplaces.map(m => (
                  <option key={m.value} value={m.value}>{m.label} ({m.value})</option>
                ))}
              </select>
            )}
            <div className="flex bg-slate-100 p-1 rounded-lg">
              <button 
                onClick={() => { setMode('amazon'); setSheetName('Amazon Data'); }}
                className={`px-4 py-1.5 rounded-md text-sm font-medium transition-all ${mode === 'amazon' ? 'bg-white shadow-sm text-indigo-600' : 'text-slate-500 hover:text-slate-700'}`}
              >
                Amazon
              </button>
              <button 
                onClick={() => { setMode('bol'); setSheetName('Bol Data'); }}
                className={`px-4 py-1.5 rounded-md text-sm font-medium transition-all ${mode === 'bol' ? 'bg-white shadow-sm text-indigo-600' : 'text-slate-500 hover:text-slate-700'}`}
              >
                Bol.com
              </button>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8">
        {/* Config Section */}
        <section className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 mb-8">
          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-2 text-slate-500">
              <Settings className="w-4 h-4" />
              <h2 className="text-sm font-semibold uppercase tracking-wider">Configuration</h2>
            </div>
            <div className="flex items-center gap-2">
              <button 
                onClick={clearSheet}
                disabled={isClearing || auditing !== null}
                className="bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold px-4 py-2 rounded-lg transition-all disabled:opacity-50 flex items-center gap-2"
              >
                {isClearing ? <RefreshCw className="w-3 h-3 animate-spin" /> : <Trash2 className="w-3 h-3" />}
                Clear Sheet
              </button>
              <button 
                onClick={runRangeAudit}
                disabled={data.length === 0 || auditing !== null || isClearing}
                className="bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold px-4 py-2 rounded-lg transition-all shadow-lg shadow-indigo-100 disabled:opacity-50 flex items-center gap-2"
              >
                {auditing !== null ? <RefreshCw className="w-3 h-3 animate-spin" /> : <Play className="w-3 h-3" />}
                Audit Range
              </button>
            </div>
          </div>
          
          <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
            <div className="space-y-1">
              <label className="text-xs font-medium text-slate-500 ml-1">Google Sheet ID</label>
              <div className="relative">
                <Database className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                <input 
                  type="text" 
                  value={sheetId}
                  onChange={(e) => setSheetId(e.target.value)}
                  placeholder="Enter Spreadsheet ID..."
                  className="w-full pl-10 pr-4 py-2 bg-slate-100 border border-slate-200 rounded-xl text-slate-500 cursor-not-allowed outline-none transition-all"
                  readOnly
                />
              </div>
            </div>
            
            <div className="space-y-1">
              <label className="text-xs font-medium text-slate-500 ml-1">Sheet Name</label>
              <input 
                type="text" 
                value={sheetName}
                onChange={(e) => setSheetName(e.target.value)}
                placeholder="e.g. MasterData"
                className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all"
              />
            </div>

            <div className="space-y-1">
              <label className="text-xs font-medium text-slate-500 ml-1">Audit Range (SKU Rows)</label>
              <div className="flex items-center gap-2">
                <input 
                  type="number" 
                  value={range.start}
                  onChange={(e) => setRange(prev => ({ ...prev, start: parseInt(e.target.value) || 1 }))}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 outline-none text-center text-sm font-bold"
                />
                <span className="text-slate-400">to</span>
                <input 
                  type="number" 
                  value={range.end}
                  onChange={(e) => setRange(prev => ({ ...prev, end: parseInt(e.target.value) || 1 }))}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 outline-none text-center text-sm font-bold"
                />
              </div>
            </div>

            <div className="flex items-end">
              <button 
                onClick={fetchSheetData}
                disabled={loading || !sheetId}
                className="w-full bg-slate-900 hover:bg-slate-800 disabled:opacity-50 text-white font-medium py-2 rounded-xl transition-colors flex items-center justify-center gap-2"
              >
                {loading ? <RefreshCw className="w-4 h-4 animate-spin" /> : <Search className="w-4 h-4" />}
                Fetch Master Data
              </button>
            </div>
          </div>

          {error && (
            <motion.div 
              initial={{ opacity: 0, y: -10 }}
              animate={{ opacity: 1, y: 0 }}
              className="mt-4 p-3 bg-red-50 border border-red-100 text-red-600 rounded-xl flex items-center gap-2 text-sm"
            >
              <AlertCircle className="w-4 h-4 flex-shrink-0" />
              {error}
            </motion.div>
          )}
        </section>

        {/* Data Table */}
        {data.length > 0 && (
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200">
                    <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider">Product</th>
                    <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider">{mode === 'amazon' ? 'ASIN' : 'EAN'}</th>
                    <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider text-center">
                      {mode === 'bol' ? 'Live Data' : 'Status'}
                    </th>
                    <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {data.map((row, idx) => (
                    <React.Fragment key={idx}>
                      <tr className={`hover:bg-slate-50 transition-colors ${selectedRow === idx ? 'bg-indigo-50/30' : ''}`}>
                        <td className="px-6 py-4">
                          <div className="font-medium text-slate-900 truncate max-w-xs">{row.Title}</div>
                          <div className="flex items-center gap-2 mt-0.5">
                            <span className="text-xs text-slate-500">Master: {row.Price}</span>
                            {auditResults[idx] && (
                              <span className={`text-xs font-bold ${auditResults[idx].auditResult.price.match ? 'text-green-600' : 'text-red-600'}`}>
                                Live: {auditResults[idx].liveData.price || 'N/A'}
                              </span>
                            )}
                          </div>
                        </td>
                        <td className="px-6 py-4 font-mono text-sm text-slate-600">
                          {mode === 'amazon' ? row.ASIN : row.EAN}
                        </td>
                        <td className="px-6 py-4 text-center">
                          {savingRow === idx ? (
                            <div className="flex flex-col items-center gap-1">
                              <RefreshCw className="w-4 h-4 text-indigo-500 animate-spin" />
                              <span className="text-[10px] font-medium text-indigo-500">Saving to Sheet...</span>
                            </div>
                          ) : auditResults[idx] ? (
                            <div className="flex flex-col items-center gap-2">
                              {mode === 'bol' && auditResults[idx].liveData.images && auditResults[idx].liveData.images.length > 0 && (
                                <div className="flex items-center justify-center gap-1 overflow-x-auto max-w-[180px] pb-1 scrollbar-hide">
                                  {auditResults[idx].liveData.images.slice(0, 5).map((img: string, i: number) => (
                                    <div key={i} className="relative group">
                                      <img 
                                        src={getProxiedUrl(img)} 
                                        className={`w-10 h-10 min-w-[40px] object-contain bg-white rounded border ${i === 0 && !auditResults[idx].auditResult.images.match ? 'border-red-500 ring-1 ring-red-500' : 'border-slate-200'}`} 
                                        alt=""
                                        referrerPolicy="no-referrer"
                                      />
                                      {i === 0 && !auditResults[idx].auditResult.images.match && (
                                        <div className="absolute -top-1 -right-1 bg-red-500 text-white rounded-full p-0.5 shadow-sm">
                                          <X className="w-2 h-2" />
                                        </div>
                                      )}
                                    </div>
                                  ))}
                                  {auditResults[idx].liveData.images.length > 5 && (
                                    <button 
                                      onClick={() => setGalleryModal({ 
                                        images: auditResults[idx].liveData.images, 
                                        title: auditResults[idx].liveData.title 
                                      })}
                                      className="w-10 h-10 min-w-[40px] flex items-center justify-center bg-slate-100 rounded border border-slate-200 text-[10px] font-bold text-slate-500 hover:bg-indigo-50 hover:text-indigo-600 transition-colors"
                                    >
                                      +{auditResults[idx].liveData.images.length - 5}
                                    </button>
                                  )}
                                </div>
                              )}
                              <div className="flex items-center justify-center gap-1">
                                {Object.values(auditResults[idx].auditResult).every((v: any) => v.match || (v.similarity && v.similarity > 0.9)) ? (
                                  <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-700">
                                    <CheckCircle2 className="w-3 h-3" /> Pass
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-medium bg-red-100 text-red-700">
                                    <XCircle className="w-3 h-3" /> Issues
                                  </span>
                                )}
                              </div>
                            </div>
                          ) : (
                            <span className="text-slate-400 text-xs italic">Pending Audit</span>
                          )}
                        </td>
                        <td className="px-6 py-4 text-right">
                          <div className="flex items-center justify-end gap-2">
                            <button 
                              onClick={() => runAudit(idx)}
                              disabled={auditing === idx.toString()}
                              className="p-2 text-indigo-600 hover:bg-indigo-50 rounded-lg transition-colors disabled:opacity-50"
                              title="Run Audit"
                            >
                              {auditing === idx.toString() ? <RefreshCw className="w-4 h-4 animate-spin" /> : <RefreshCw className="w-4 h-4" />}
                            </button>
                            <button 
                              onClick={() => setSelectedRow(selectedRow === idx ? null : idx)}
                              className="p-2 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-colors"
                              title="View Details"
                            >
                              {selectedRow === idx ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
                            </button>
                          </div>
                        </td>
                      </tr>
                      
                      {/* Detailed Audit View */}
                      <AnimatePresence>
                        {selectedRow === idx && (
                          <motion.tr 
                            initial={{ opacity: 0, height: 0 }}
                            animate={{ opacity: 1, height: 'auto' }}
                            exit={{ opacity: 0, height: 0 }}
                          >
                            <td colSpan={4} className="px-6 py-8 bg-slate-50/50 border-b border-slate-200">
                              {!auditResults[idx] ? (
                                <div className="text-center py-12">
                                  <Info className="w-8 h-8 text-slate-300 mx-auto mb-2" />
                                  <p className="text-slate-500">Run audit to see detailed comparison</p>
                                  <button 
                                    onClick={() => runAudit(idx)}
                                    className="mt-4 text-indigo-600 font-medium hover:underline"
                                  >
                                    Start Audit Now
                                  </button>
                                </div>
                              ) : (
                                <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                                  {/* Text Comparison */}
                                  <div className="space-y-6">
                                    <h3 className="text-sm font-bold text-slate-400 uppercase tracking-widest mb-4">Content Validation</h3>
                                    
                                    <ComparisonItem 
                                      label="Title" 
                                      master={auditResults[idx].auditResult.title.master}
                                      live={auditResults[idx].auditResult.title.live}
                                      similarity={auditResults[idx].auditResult.title.similarity}
                                    />

                                    <ComparisonItem 
                                      label="Description" 
                                      master={auditResults[idx].auditResult.description.master}
                                      live={auditResults[idx].auditResult.description.live}
                                      similarity={auditResults[idx].auditResult.description.similarity}
                                      status={auditResults[idx].auditResult.description.status}
                                      isLongText
                                    />

                                    {auditResults[idx].auditResult.bullets && (
                                      <div className="space-y-3">
                                        <label className="text-xs font-semibold text-slate-500">
                                          {mode === 'bol' ? 'Product Specifications / Features' : 'Bullet Points'}
                                        </label>
                                        {auditResults[idx].auditResult.bullets.length === 0 ? (
                                          <div className="text-[10px] text-slate-400 italic">No bullets/features found on live page</div>
                                        ) : (
                                          auditResults[idx].auditResult.bullets.map((b: any, i: number) => (
                                            <ComparisonItem 
                                              key={i}
                                              label={mode === 'bol' ? `Spec ${i+1}` : `Bullet ${i+1}`} 
                                              master={b.master}
                                              live={b.live}
                                              similarity={b.similarity}
                                              mini
                                            />
                                          ))
                                        )}
                                      </div>
                                    )}
                                  </div>

                                  <div className="space-y-6">
                                    <h3 className="text-sm font-bold text-slate-400 uppercase tracking-widest mb-4">Direct Extraction</h3>
                                    
                                    <div className="grid grid-cols-2 gap-4">
                                      <div className="p-4 bg-white rounded-xl border border-slate-200">
                                        <div className="text-xs text-slate-400 mb-1">Price</div>
                                        <div className="text-lg font-bold text-indigo-600">{auditResults[idx].liveData.price || 'N/A'}</div>
                                        {auditResults[idx].liveData.listPrice && auditResults[idx].liveData.listPrice !== 'N/A' && (
                                          <div className="text-[10px] text-slate-400 line-through">List: {auditResults[idx].liveData.listPrice}</div>
                                        )}
                                      </div>
                                      <div className="p-4 bg-white rounded-xl border border-slate-200">
                                        <div className="text-xs text-slate-400 mb-1">Shipping</div>
                                        <div className="text-sm font-bold text-indigo-600">{auditResults[idx].liveData.shipping || 'N/A'}</div>
                                        <div className="text-[10px] text-slate-400 mt-1 truncate" title={auditResults[idx].liveData.rawShipping}>
                                          Live: {auditResults[idx].liveData.rawShipping || 'N/A'}
                                        </div>
                                      </div>
                                    </div>

                                    <div className="grid grid-cols-2 gap-4">
                                      <div className="p-4 bg-white rounded-xl border border-slate-200">
                                        <div className="text-xs text-slate-400 mb-1">Variations</div>
                                        <div className="flex items-center gap-2">
                                          <span className={`text-sm font-bold ${auditResults[idx].liveData.variations > 0 ? 'text-green-600' : 'text-slate-400'}`}>
                                            {auditResults[idx].liveData.variations > 0 ? 'YES' : 'NO'}
                                          </span>
                                          <span className="text-xs text-slate-400">({auditResults[idx].liveData.variations} found)</span>
                                        </div>
                                      </div>
                                      <div className="p-4 bg-white rounded-xl border border-slate-200">
                                        <div className="text-xs text-slate-400 mb-1">A+ Content</div>
                                        <div className={`text-sm font-bold ${auditResults[idx].liveData.hasAPlus ? 'text-green-600' : 'text-slate-400'}`}>
                                          {auditResults[idx].liveData.hasAPlus ? 'PRESENT' : 'MISSING'}
                                        </div>
                                      </div>
                                    </div>

                                    <div className="space-y-4">
                                      <label className="text-xs font-bold text-slate-400 uppercase tracking-widest flex items-center gap-2">
                                        <ImageIcon className="w-4 h-4" /> Image Comparison
                                      </label>
                                      
                                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                        {/* Master Images */}
                                          <div className="bg-white p-4 rounded-xl border border-slate-200">
                                            <div className="text-[10px] font-bold text-slate-400 uppercase mb-3 flex items-center justify-between">
                                              <span>Master Images</span>
                                              <span className="bg-slate-100 px-2 py-0.5 rounded text-slate-500">{auditResults[idx].auditResult.images.master.length}</span>
                                            </div>
                                            <div className="grid grid-cols-3 gap-2">
                                              {auditResults[idx].auditResult.images.master.slice(0, 6).map((img: string, i: number) => (
                                                <div key={i} className="aspect-square bg-slate-50 border border-slate-100 rounded-lg overflow-hidden relative group">
                                                  <img 
                                                    src={getProxiedUrl(img)} 
                                                    alt="" 
                                                    className="w-full h-full object-contain transition-transform group-hover:scale-110" 
                                                    referrerPolicy="no-referrer" 
                                                  />
                                                  <div className="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center">
                                                    <button 
                                                      onClick={() => window.open(img, '_blank')}
                                                      className="p-1.5 bg-white rounded-full text-slate-700 shadow-lg"
                                                    >
                                                      <ExternalLink className="w-3 h-3" />
                                                    </button>
                                                  </div>
                                                </div>
                                              ))}
                                              {auditResults[idx].auditResult.images.master.length === 0 && (
                                                <div className="col-span-3 text-[10px] text-slate-400 italic py-4 text-center">No master images</div>
                                              )}
                                            </div>
                                          </div>

                                          {/* Live Images */}
                                          <div className="bg-white p-4 rounded-xl border border-slate-200">
                                            <div className="text-[10px] font-bold text-slate-400 uppercase mb-3 flex items-center justify-between">
                                              <span>Live Marketplace Images</span>
                                              <button 
                                                onClick={() => setGalleryModal({ 
                                                  images: auditResults[idx].liveData.images, 
                                                  title: auditResults[idx].liveData.title 
                                                })}
                                                className="text-indigo-600 hover:underline flex items-center gap-1"
                                              >
                                                View Gallery ({auditResults[idx].liveData.images?.length || 0})
                                              </button>
                                            </div>
                                            
                                            {/* Main Image Highlight */}
                                            {auditResults[idx].liveData.images && auditResults[idx].liveData.images.length > 0 && (
                                              <div className="mb-4 aspect-square bg-slate-50 border border-slate-200 rounded-xl overflow-hidden relative group">
                                                <div className="absolute top-2 left-2 z-10 bg-indigo-600 text-white text-[9px] font-bold px-2 py-0.5 rounded shadow-sm uppercase">Main View</div>
                                                <img 
                                                  src={getProxiedUrl(auditResults[idx].liveData.images[0])} 
                                                  alt="Main Product" 
                                                  className="w-full h-full object-contain transition-transform group-hover:scale-105" 
                                                  referrerPolicy="no-referrer" 
                                                />
                                                <div className="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center">
                                                  <button 
                                                    onClick={() => window.open(auditResults[idx].liveData.images[0], '_blank')}
                                                    className="p-2 bg-white rounded-full text-slate-700 shadow-lg"
                                                  >
                                                    <ExternalLink className="w-4 h-4" />
                                                  </button>
                                                </div>
                                              </div>
                                            )}

                                            <div className="text-[10px] font-bold text-slate-400 uppercase mb-2">Gallery View</div>
                                            <div className="grid grid-cols-4 gap-2">
                                              {auditResults[idx].liveData.images?.slice(1, 9).map((img: string, i: number) => (
                                                <div key={i} className="aspect-square bg-slate-50 border border-slate-100 rounded-lg overflow-hidden relative group">
                                                  <img 
                                                    src={getProxiedUrl(img)} 
                                                    alt={`Gallery Image ${i+1}`} 
                                                    className="w-full h-full object-contain transition-transform group-hover:scale-110" 
                                                    referrerPolicy="no-referrer" 
                                                  />
                                                  <div className="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center">
                                                    <button 
                                                      onClick={() => window.open(img, '_blank')}
                                                      className="p-1.5 bg-white rounded-full text-slate-700 shadow-lg"
                                                    >
                                                      <ExternalLink className="w-3 h-3" />
                                                    </button>
                                                  </div>
                                                </div>
                                              ))}
                                              {(!auditResults[idx].liveData.images || auditResults[idx].liveData.images.length <= 1) && (
                                                <div className="col-span-4 text-[10px] text-slate-400 italic py-4 text-center">No secondary images</div>
                                              )}
                                            </div>
                                          </div>
                                      </div>
                                    </div>

                                    <div className="pt-4">
                                      <a 
                                        href={mode === 'amazon' ? `https://www.${marketplace}/dp/${row.ASIN}` : `https://www.bol.com/nl/nl/s/?searchtext=${row.EAN}`}
                                        target="_blank"
                                        rel="noopener noreferrer"
                                        className="inline-flex items-center gap-2 text-sm text-indigo-600 font-medium hover:underline"
                                      >
                                        View Live Page <ExternalLink className="w-3 h-3" />
                                      </a>
                                    </div>
                                  </div>
                                </div>
                              )}
                            </td>
                          </motion.tr>
                        )}
                      </AnimatePresence>
                    </React.Fragment>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {data.length === 0 && !loading && (
          <div className="text-center py-24 bg-white rounded-2xl border border-dashed border-slate-300">
            <Database className="w-12 h-12 text-slate-200 mx-auto mb-4" />
            <h3 className="text-lg font-medium text-slate-900">No data loaded</h3>
            <p className="text-slate-500 max-w-xs mx-auto mt-1">Enter your Google Sheet ID above to start the marketplace audit process.</p>
          </div>
        )}
      </main>

      {/* Gallery Modal */}
      <AnimatePresence>
        {galleryModal && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-sm"
          >
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col"
            >
              <div className="p-4 border-b border-slate-100 flex items-center justify-between">
                <div>
                  <h3 className="font-bold text-slate-900 line-clamp-1">{galleryModal.title}</h3>
                  <p className="text-xs text-slate-500">{galleryModal.images.length} Product Images Found</p>
                </div>
                <button 
                  onClick={() => setGalleryModal(null)}
                  className="p-2 hover:bg-slate-100 rounded-full transition-colors"
                >
                  <X className="w-5 h-5 text-slate-500" />
                </button>
              </div>
              
              <div className="flex-1 overflow-y-auto p-6 bg-slate-50">
                <div className="space-y-8">
                  {/* Main Image Section */}
                  {galleryModal.images.length > 0 && (
                    <div className="space-y-3">
                      <h4 className="text-xs font-bold text-slate-400 uppercase tracking-widest">Main Product Image</h4>
                      <div className="max-w-md mx-auto aspect-square bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-sm group relative">
                        <img 
                          src={getProxiedUrl(galleryModal.images[0])} 
                          alt="Main Product" 
                          className="w-full h-full object-contain p-4"
                          referrerPolicy="no-referrer"
                        />
                        <div className="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center">
                          <button 
                            onClick={() => window.open(galleryModal.images[0], '_blank')}
                            className="p-3 bg-white rounded-full text-slate-700 shadow-xl"
                          >
                            <ExternalLink className="w-6 h-6" />
                          </button>
                        </div>
                      </div>
                    </div>
                  )}

                  {/* Gallery Section */}
                  {galleryModal.images.length > 1 && (
                    <div className="space-y-3">
                      <h4 className="text-xs font-bold text-slate-400 uppercase tracking-widest">Secondary Gallery Images</h4>
                      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
                        {galleryModal.images.slice(1).map((img, i) => (
                          <div key={i} className="aspect-square bg-white rounded-xl border border-slate-200 overflow-hidden group relative shadow-sm">
                            <img 
                              src={getProxiedUrl(img)} 
                              alt={`Gallery ${i+1}`} 
                              className="w-full h-full object-contain transition-transform group-hover:scale-110"
                              referrerPolicy="no-referrer"
                            />
                            <div className="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center">
                              <button 
                                onClick={() => window.open(img, '_blank')}
                                className="p-2 bg-white rounded-full text-slate-700 shadow-lg"
                              >
                                <ExternalLink className="w-4 h-4" />
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              </div>
              
              <div className="p-4 border-t border-slate-100 bg-white flex justify-end">
                <button 
                  onClick={() => setGalleryModal(null)}
                  className="px-6 py-2 bg-slate-900 text-white rounded-xl font-bold hover:bg-slate-800 transition-colors"
                >
                  Close Gallery
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

function HighlightDiff({ master, live }: { master: string, live: string }) {
  if (!live) return <span className="italic text-slate-300">Not Found</span>;
  
  const masterWords = (master || '').toLowerCase().split(/\s+/).map(w => w.replace(/[™©®•●▪◦‣■□▢▣▤▥▦▧▨▩▪▫▬▭▮▯▰▱▲△▴▵▶▷▸▹▼▽▾▿◀◁◂◃]/g, '').trim()).filter(Boolean);
  const liveWords = live.split(/\s+/);
  
  return (
    <>
      {liveWords.map((word, i) => {
        const cleanWord = word.toLowerCase().replace(/[™©®•●▪◦‣■□▢▣▤▥▦▧▨▩▪▫▬▭▮▯▰▱▲△▴▵▶▷▸▹▼▽▾▿◀◁◂◃]/g, '').trim();
        const isMatch = masterWords.includes(cleanWord);
        return (
          <span key={i} className={isMatch ? '' : 'font-bold text-red-700 underline decoration-red-300 decoration-2'}>
            {word}{' '}
          </span>
        );
      })}
    </>
  );
}

function ComparisonItem({ label, master, live, similarity, status, isLongText = false, mini = false }: any) {
  const [expanded, setExpanded] = useState(false);
  const isImage = live && live.startsWith('IMAGE:');
  const isAPlusImages = live && live.startsWith('APLUS_IMAGES:');
  const isAPlusData = live && live.startsWith('APLUS_DATA:');
  
  let imageUrls: string[] = [];
  let aPlusText = '';

  if (isImage) {
    imageUrls = [live.replace('IMAGE:', '')];
  } else if (isAPlusImages) {
    imageUrls = live.replace('APLUS_IMAGES:', '').split(',');
  } else if (isAPlusData) {
    try {
      const data = JSON.parse(live.replace('APLUS_DATA:', ''));
      imageUrls = data.images || [];
      aPlusText = data.text || '';
    } catch (e) {}
  }
  
  const isMatch = similarity > 0.99;
  const simColor = isMatch ? 'text-green-600 bg-green-50' : similarity > 0.7 ? 'text-yellow-600 bg-yellow-50' : 'text-red-600 bg-red-50';

  return (
    <div className={`space-y-2 ${mini ? 'bg-white p-3 rounded-xl border border-slate-100 shadow-sm' : ''}`}>
      <div className="flex items-center justify-between">
        <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">{label}</label>
        <div className="flex items-center gap-2">
          {status && <span className="text-[10px] font-bold text-amber-600 bg-amber-50 px-2 py-0.5 rounded-full uppercase flex items-center gap-1"><AlertCircle className="w-2.5 h-2.5" /> {status}</span>}
          {!isMatch && !isImage && !isAPlusImages && !isAPlusData && !status && <span className="text-[10px] font-bold text-red-500 uppercase flex items-center gap-1"><AlertCircle className="w-2.5 h-2.5" /> Mismatch</span>}
          {(isImage || isAPlusImages || isAPlusData) && <span className="text-[10px] font-bold text-indigo-500 uppercase flex items-center gap-1"><ImageIcon className="w-2.5 h-2.5" /> A+ Content</span>}
          <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full uppercase ${simColor}`}>
            {(isImage || isAPlusImages || isAPlusData) ? 'Visual Check' : `${Math.round(similarity * 100)}% Match`}
          </span>
        </div>
      </div>
      
      <div className={`grid grid-cols-1 ${!mini ? 'lg:grid-cols-2' : ''} gap-3`}>
        <div className="space-y-1">
          <div className="text-[9px] text-slate-400 font-bold uppercase tracking-widest ml-1">Master Data</div>
          <div className={`text-xs p-3 bg-slate-50 rounded-xl border border-slate-200 font-medium leading-relaxed ${isLongText && !expanded ? 'line-clamp-3' : ''} ${!isMatch && !isImage && !isAPlusImages && !isAPlusData ? 'border-amber-200 bg-amber-50/30' : ''}`}>
            {master || <span className="italic text-slate-300">Empty</span>}
          </div>
        </div>
        
        {(!isMatch || !mini || isImage || isAPlusImages || isAPlusData) && (
          <div className="space-y-1">
            <div className="text-[9px] text-indigo-400 font-bold uppercase tracking-widest ml-1">Live Listing</div>
            <div className={`text-xs p-3 bg-white rounded-xl border border-slate-200 font-medium leading-relaxed ${isLongText && !expanded ? 'line-clamp-3' : ''} ${!isMatch && !isImage && !isAPlusImages && !isAPlusData ? 'border-red-100 bg-red-50/50' : 'text-slate-700'}`}>
              {(isImage || isAPlusImages || isAPlusData) ? (
                <div className="space-y-2">
                  <p className="text-[10px] text-slate-500 italic mb-2">Standard description missing. Showing A+ Content:</p>
                  
                  {aPlusText && (
                    <div className="mb-4 p-2 bg-slate-50 rounded border border-slate-100 text-[11px] whitespace-pre-wrap max-h-[200px] overflow-y-auto">
                      {aPlusText}
                    </div>
                  )}

                  {imageUrls.length > 0 && (
                    <div className="max-h-[400px] overflow-y-auto space-y-1 rounded-lg border border-slate-100 p-1 bg-slate-50">
                      {imageUrls.map((url, i) => (
                        <img key={i} src={url} alt={`A+ Content ${i+1}`} className="w-full rounded shadow-sm" referrerPolicy="no-referrer" />
                      ))}
                    </div>
                  )}
                </div>
              ) : (
                isMatch ? (live || <span className="italic text-slate-300">Not Found</span>) : <HighlightDiff master={master} live={live} />
              )}
            </div>
          </div>
        )}
      </div>
      
      {isLongText && (String(master || '').length > 150 || String(live || '').length > 150) && (
        <button 
          onClick={() => setExpanded(!expanded)}
          className="text-[10px] text-indigo-600 font-bold hover:text-indigo-700 transition-colors ml-1"
        >
          {expanded ? 'Collapse View' : 'Expand Details'}
        </button>
      )}
    </div>
  );
}
