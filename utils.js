function parseDateList(cell) {
  if (!cell) return [];
  if (typeof cell !== 'string') return [];
  return cell.split(",").map(v => v.trim()).filter(v => v.length > 0);
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
  // Zaokruži na 3 decimale
  const rounded = Math.round(share * 1000) / 1000;
  
  // Tačna poređenja sa tolerancijom
  if (rounded >= 0.999 && rounded <= 1.001) return "celo";
  if (rounded >= 0.499 && rounded <= 0.501) return "polovina";
  if (rounded >= 0.332 && rounded <= 0.335) return "trećina";
  if (rounded >= 0.249 && rounded <= 0.251) return "četvrtina";
  if (rounded >= 0.124 && rounded <= 0.126) return "osmina";
  
  // Kombinovane vrednosti koje nastaju odsecanjem
  // 0.875 = 1 - 0.125 (celo - osmina)
  if (rounded >= 0.874 && rounded <= 0.876) return "celo bez osmine";
  // 0.75 = 1 - 0.25 (celo - četvrtina) = 3/4
  if (rounded >= 0.749 && rounded <= 0.751) return "tri četvrtine";
  // 0.667 = 2/3 (celo - trećina)
  if (rounded >= 0.666 && rounded <= 0.668) return "dve trećine";
  // 0.542 = približno polovina (0.5) + osmina (0.125) - ali zbog float-a 0.542
  if (rounded >= 0.541 && rounded <= 0.543) return "polovina i osmina";
  // 0.375 = 3/8 (tri osmine)
  if (rounded >= 0.374 && rounded <= 0.376) return "tri osmine";
  // 0.625 = 5/8 (pet osmina)
  if (rounded >= 0.624 && rounded <= 0.626) return "pet osmina";
  // 0.833 = 5/6
  if (rounded >= 0.832 && rounded <= 0.834) return "pet šestina";
  // 0.667 = 2/3
  if (rounded >= 0.666 && rounded <= 0.668) return "dve trećine";
  // 0.333 = 1/3
  if (rounded >= 0.332 && rounded <= 0.335) return "trećina";
  // 0.042 = sitno (osmina/3)
  if (rounded <= 0.05) return "minimalno";
  
  // Ako ništa ne odgovara, prikaži kao razlomak
  return `${rounded}`;
}

function isWeekend(dateStr) {
  const [d, m, y] = dateStr.split("-").map(Number);
  const date = new Date(y, m - 1, d);
  return date.getDay() === 0 || date.getDay() === 6;
}

module.exports = { parseDateList, shareToNumber, shareToString, isWeekend };