function parseDateList(cell) {
  if (!cell) return [];
  if (typeof cell !== 'string') return [];
  
  // Podeli po zarezu, tačka-zarezu ili razmaku
  return cell.split(/[,;]+/)
    .map(v => v.trim())
    .filter(v => v.length > 0)
    .map(v => {
      // Ako je već u formatu DD-MM-YYYY, ostavi kako jeste
      if (/^\d{2}-\d{2}-\d{4}$/.test(v)) return v;
      
      // Ako je u formatu YYYY-MM-DD ili YYYY-MM-DD HH:MM:SS, konvertuj u DD-MM-YYYY
      const isoMatch = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (isoMatch) {
        const [, y, m, d] = isoMatch;
        return `${d}-${m}-${y}`;
      }
      
      // Ako je u formatu DD/MM/YYYY, konvertuj u DD-MM-YYYY
      const slashMatch = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (slashMatch) {
        const [, d, m, y] = slashMatch;
        return `${d.padStart(2, '0')}-${m.padStart(2, '0')}-${y}`;
      }
      
      // Ako je u formatu MM/DD/YYYY (US), konvertuj u DD-MM-YYYY
      const usMatch = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (usMatch) {
        const [, m, d, y] = usMatch;
        return `${d.padStart(2, '0')}-${m.padStart(2, '0')}-${y}`;
      }
      
      // Ako je samo broj (Excel serijski), pokušaj konverziju
      const numVal = Number(v);
      if (!isNaN(numVal) && numVal > 40000 && numVal < 60000) {
        const excelEpoch = new Date(1899, 11, 30);
        const date = new Date(excelEpoch.getTime() + numVal * 86400000);
        const d = date.getDate().toString().padStart(2, '0');
        const m = (date.getMonth() + 1).toString().padStart(2, '0');
        const y = date.getFullYear();
        return `${d}-${m}-${y}`;
      }
      
      // Vrati original ako ništa ne odgovara
      return v;
    });
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
  
  // Fallback za stare vrednosti
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
  
  // Kombinovane vrednosti koje nastaju odsecanjem
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

module.exports = { parseDateList, shareToNumber, shareToString, isWeekend };