function parseDateList(cell) {
  if (!cell) return [];
  if (typeof cell !== 'string') return [];
  
  const dates = [];
  
  // Prvo podeli po zarezu
  const parts = cell.split(',').map(p => p.trim()).filter(p => p.length > 0);
  
  for (const part of parts) {
    // Proveri da li je opseg sa "do" (srpski)
    let rangeMatch = part.match(/^(\d{2}-\d{2}-\d{4})\s+do\s+(\d{2}-\d{2}-\d{4})$/i);
    
    // Proveri da li je opseg sa "to" (engleski)
    if (!rangeMatch) {
      rangeMatch = part.match(/^(\d{2}-\d{2}-\d{4})\s+to\s+(\d{2}-\d{2}-\d{4})$/i);
    }
    
    if (rangeMatch) {
      // Proširi opseg u niz pojedinačnih datuma
      const expandedDates = expandDateRange(rangeMatch[1], rangeMatch[2]);
      dates.push(...expandedDates);
      console.log(`      📅 Opseg: ${rangeMatch[1]} do ${rangeMatch[2]} → ${expandedDates.length} datuma`);
      continue;
    }
    
    // Ako nije opseg, dodaj kao pojedinačni datum
    if (/^\d{2}-\d{2}-\d{4}$/.test(part)) {
      dates.push(part);
    }
  }
  
  // Ukloni duplikate i sortiraj
  return [...new Set(dates)].sort((a, b) => {
    const [da, ma, ya] = a.split("-").map(Number);
    const [db, mb, yb] = b.split("-").map(Number);
    return new Date(ya, ma - 1, da) - new Date(yb, mb - 1, db);
  });
}

/**
 * Proširuje opseg datuma u niz pojedinačnih datuma
 */
function expandDateRange(startStr, endStr) {
  const [ds, ms, ys] = startStr.split("-").map(Number);
  const [de, me, ye] = endStr.split("-").map(Number);
  
  const start = new Date(ys, ms - 1, ds);
  const end = new Date(ye, me - 1, de);
  
  if (start > end) {
    throw new Error(`Početni datum (${startStr}) je posle krajnjeg datuma (${endStr}).`);
  }
  
  const dates = [];
  const current = new Date(start);
  
  while (current <= end) {
    const d = current.getDate().toString().padStart(2, '0');
    const m = (current.getMonth() + 1).toString().padStart(2, '0');
    const y = current.getFullYear();
    dates.push(`${d}-${m}-${y}`);
    current.setDate(current.getDate() + 1);
  }
  
  return dates;
}

function shareToNumber(share) {
  const shareMap = {
    "celo": 1,
    "polovina": 0.5,
    "trećina": 1/3,
    "četvrtina": 0.25,
    "osmina": 0.125
  };
  
  const result = shareMap[share.toLowerCase().trim()];
  if (result !== undefined) {
    return result;
  }
  
  const fallbackMap = {
    "1/1": 1, "1": 1,
    "1/2": 0.5, "0.5": 0.5,
    "1/3": 1/3,
    "1/4": 0.25, "0.25": 0.25,
    "1/8": 0.125, "0.125": 0.125
  };
  
  const fallbackResult = fallbackMap[share.toLowerCase().trim()];
  if (fallbackResult !== undefined) {
    return fallbackResult;
  }
  
  throw new Error(
    `Neispravan udeo dežurstva: "${share}". ` +
    `Dozvoljene vrednosti: celo, polovina, trećina, četvrtina, osmina`
  );
}

function shareToString(share) {
  const rounded = Math.round(share * 1000) / 1000;
  
  if (rounded >= 0.999 && rounded <= 1.001) return "celo";
  if (rounded >= 0.499 && rounded <= 0.501) return "polovina";
  if (rounded >= 0.332 && rounded <= 0.335) return "trećina";
  if (rounded >= 0.249 && rounded <= 0.251) return "četvrtina";
  if (rounded >= 0.124 && rounded <= 0.126) return "osmina";
  
  if (rounded >= 0.874 && rounded <= 0.876) return "celo bez osmine";
  if (rounded >= 0.749 && rounded <= 0.751) return "tri četvrtine";
  if (rounded >= 0.666 && rounded <= 0.668) return "dve trećine";
  if (rounded >= 0.541 && rounded <= 0.543) return "polovina i osmina";
  if (rounded >= 0.374 && rounded <= 0.376) return "tri osmine";
  if (rounded >= 0.624 && rounded <= 0.626) return "pet osmina";
  if (rounded >= 0.832 && rounded <= 0.834) return "pet šestina";
  if (rounded <= 0.05) return "minimalno";
  
  if (rounded === 1) return "celo";
  if (rounded === 0.5) return "polovina";
  if (rounded === 0.25) return "četvrtina";
  if (rounded === 0.125) return "osmina";
  
  return `${rounded}`;
}

function isWeekend(dateStr) {
  const [d, m, y] = dateStr.split("-").map(Number);
  const date = new Date(y, m - 1, d);
  return date.getDay() === 0 || date.getDay() === 6;
}

module.exports = { parseDateList, shareToNumber, shareToString, isWeekend, expandDateRange };