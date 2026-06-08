import stringSimilarity from 'string-similarity';

export const getSimilarity = (str1: string, str2: string) => {
  if (!str1 || !str2) return 0;
  return stringSimilarity.compareTwoStrings(str1.toLowerCase(), str2.toLowerCase());
};

export function cleanAndNormalizePrice(priceStr: string): string {
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

export function getScoreGrade(score: number): string {
  if (score >= 90) return 'A';
  if (score >= 70) return 'B';
  if (score >= 50) return 'C';
  return 'D';
}
