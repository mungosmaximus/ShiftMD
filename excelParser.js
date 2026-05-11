// ShiftMD v3.6 - Dvojezični Excel parser sa podrškom za rangove i robustnim parsiranjem datuma
const ExcelJS = require("exceljs");
const { parseDateList } = require("./utils");

console.log("✅ ShiftMD v3.6 - Dvojezični Excel parser (sa rangovima i popravljenim datumima)");

const HEADER_MAP = {
  "ime i prezime": "name", "full name": "name", "name": "name",
  "starešinstvo": "seniority", "seniority": "seniority", "senioritet": "seniority", "years of experience": "seniority", "experience": "seniority",
  "može biti glavni": "canBeChief", "can be chief": "canBeChief", "glavni": "canBeChief", "chief": "canBeChief",
  "tip zaposlenog": "role", "employee type": "role", "role": "role", "position": "role", "tip": "role", "type": "role",
  "udeo dežurstva": "share", "duty share": "share", "share": "share", "udeo": "share",
  "ne može": "unavailable", "unavailable": "unavailable", "ne moze": "unavailable", "cannot work": "unavailable",
  "želi da dežura": "preferred", "preferred": "preferred", "želi": "preferred", "wants to work": "preferred", "wishes": "preferred", "zeli da dezura": "preferred",
  "rang": "rank", "rank": "rank", "redni broj": "rank", "order": "rank", "priority": "rank", "nivo": "rank", "level": "rank"
};

const CHIEF_VALUES = ["da", "može", "moze", "yes", "y", "true", "1"];
const ROLE_SPECIALIST = ["specijalista", "specialist", "attending"];
const ROLE_RESIDENT = ["specijalizant", "resident", "trainee", "fellow"];

const SHARE_MAP = {
  "celo": "celo", "polovina": "polovina", "trećina": "trećina", "cetvrtina": "četvrtina", "četvrtina": "četvrtina", "osmina": "osmina",
  "full": "celo", "whole": "celo", "half": "polovina", "third": "trećina", "quarter": "četvrtina", "eighth": "osmina",
  "1/1": "celo", "1": "celo", "1/2": "polovina", "0.5": "polovina", "1/3": "trećina", "1/4": "četvrtina", "0.25": "četvrtina", "1/8": "osmina", "0.125": "osmina"
};

/**
 * Konvertuje bilo koju vrednost datuma u string formata DD-MM-YYYY
 * Podržava: Date objekte, Excel serijske brojeve, stringove u raznim formatima
 */
function parseDateValue(value) {
  if (value === null || value === undefined || value === '') return '';
  
  // 1. Ako je već string u formatu DD-MM-YYYY
  if (typeof value === 'string') {
    const trimmed = value.trim();
    
    // Proveri da li je već u formatu DD-MM-YYYY ili DD-MM-YYYY HH:MM:SS
    const dateRegex = /^(\d{2})-(\d{2})-(\d{4})/;
    if (dateRegex.test(trimmed)) {
      return trimmed.substring(0, 10); // Uzmi samo DD-MM-YYYY deo
    }
    
    // Proveri da li je u formatu YYYY-MM-DD ili YYYY-MM-DD HH:MM:SS (ISO format)
    const isoRegex = /^(\d{4})-(\d{2})-(\d{2})/;
    const isoMatch = trimmed.match(isoRegex);
    if (isoMatch) {
      const [, y, m, d] = isoMatch;
      return `${d}-${m}-${y}`;
    }
    
    // Proveri da li je u formatu MM/DD/YYYY (US format)
    const usRegex = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/;
    const usMatch = trimmed.match(usRegex);
    if (usMatch) {
      const [, m, d, y] = usMatch;
      return `${d.padStart(2, '0')}-${m.padStart(2, '0')}-${y}`;
    }
    
    // Ako je običan broj (Excel serijski broj)
    const numVal = Number(trimmed);
    if (!isNaN(numVal) && numVal > 40000 && numVal < 60000) {
      return excelSerialToDate(numVal);
    }
    
    return trimmed;
  }
  
  // 2. Ako je Date objekat
  if (value instanceof Date) {
    const d = value.getDate().toString().padStart(2, '0');
    const m = (value.getMonth() + 1).toString().padStart(2, '0');
    const y = value.getFullYear();
    return `${d}-${m}-${y}`;
  }
  
  // 3. Ako je broj (Excel serijski broj datuma)
  if (typeof value === 'number') {
    if (value > 40000 && value < 60000) {
      return excelSerialToDate(value);
    }
    return value.toString();
  }
  
  // 4. Ako je objekat sa richText
  if (value && typeof value === 'object' && value.richText) {
    const text = value.richText.map(t => t.text).join('');
    return parseDateValue(text); // Rekurzivno parsiraj
  }
  
  return String(value).trim();
}

/**
 * Konvertuje Excel serijski broj u datum DD-MM-YYYY
 */
function excelSerialToDate(serial) {
  // Excel datum počinje od 1. januara 1900. (serijski broj 1)
  // Postoji bug: Excel misli da 1900. godina ima 29. februar
  const excelEpoch = new Date(1899, 11, 30); // 30. decembar 1899.
  const date = new Date(excelEpoch.getTime() + serial * 86400000);
  
  const d = date.getDate().toString().padStart(2, '0');
  const m = (date.getMonth() + 1).toString().padStart(2, '0');
  const y = date.getFullYear();
  return `${d}-${m}-${y}`;
}

module.exports = async function parseExcel(path) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) throw new Error("Excel fajl je prazan ili nema listova. / Excel file is empty or has no sheets.");
    
    const headerRow = worksheet.getRow(1);
    const columnMap = {};
    
    headerRow.eachCell((cell, colNumber) => {
      const rawHeader = cell.value;
      if (!rawHeader) return;
      const headerStr = String(rawHeader).trim().toLowerCase();
      const fieldKey = HEADER_MAP[headerStr];
      if (fieldKey) {
        columnMap[colNumber] = fieldKey;
        console.log(`   📋 Kolona ${colNumber}: "${rawHeader}" → ${fieldKey}`);
      } else {
        console.warn(`   ⚠️  Nepoznata kolona: "${rawHeader}" (kolona ${colNumber})`);
      }
    });
    
    const requiredFields = ["name", "seniority", "canBeChief", "role", "share"];
    const missingFields = requiredFields.filter(f => !Object.values(columnMap).includes(f));
    if (missingFields.length > 0) {
      throw new Error(`Nedostaju obavezne kolone / Missing required columns: ${missingFields.join(", ")}`);
    }
    
    const hasRankColumn = Object.values(columnMap).includes("rank");
    console.log(`   📋 Kolona 'rang' ${hasRankColumn ? 'je prisutna' : 'nije prisutna'}`);
    
    const rows = [];
    const totalRows = worksheet.rowCount;
    
    for (let rowNumber = 2; rowNumber <= totalRows; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      let hasData = false;
      row.eachCell(() => { hasData = true; });
      if (!hasData) continue;
      
      const rowData = {};
      row.eachCell((cell, colNumber) => {
        const fieldKey = columnMap[colNumber];
        if (!fieldKey) return;
        
        let cellValue = cell.value;
        
        // Konvertuj datume i brojeve u string
        if (cellValue instanceof Date) {
          const d = cellValue.getDate().toString().padStart(2, '0');
          const m = (cellValue.getMonth() + 1).toString().padStart(2, '0');
          const y = cellValue.getFullYear();
          cellValue = `${d}-${m}-${y}`;
        } else if (typeof cellValue === 'number') {
          // Proveri da li je Excel serijski broj datuma
          if (cellValue > 40000 && cellValue < 60000) {
            cellValue = excelSerialToDate(cellValue);
          } else {
            cellValue = cellValue.toString();
          }
        } else if (cellValue && typeof cellValue === 'object' && cellValue.richText) {
          cellValue = cellValue.richText.map(t => t.text).join('');
        } else if (cellValue === null || cellValue === undefined) {
          cellValue = '';
        } else {
          cellValue = String(cellValue).trim();
        }
        
        // Za kolone sa datumima (unavailable, preferred) - dodatna obrada
        if (fieldKey === 'unavailable' || fieldKey === 'preferred') {
          cellValue = parseDateValue(cellValue);
        }
        
        rowData[fieldKey] = cellValue;
      });
      
      if (!rowData.name || rowData.name === '') continue;
      rows.push(rowData);
    }
    
    if (rows.length === 0) throw new Error("Excel fajl nema podataka. / Excel file has no data.");
    console.log(`   ✅ Učitano ${rows.length} redova podataka`);
    
    return rows.map((row, index) => {
      const lineNum = index + 2;
      if (!row.name) throw new Error(`Red ${lineNum}: Nedostaje ime / Missing name.`);
      
      const seniority = Number(row.seniority);
      if (isNaN(seniority)) throw new Error(`Red ${lineNum}: Neispravno starešinstvo "${row.seniority}" za ${row.name}.`);
      
      const canBeChiefRaw = String(row.canBeChief || '').toLowerCase().trim();
      const canBeChief = CHIEF_VALUES.some(v => canBeChiefRaw === v);
      
      const roleRaw = String(row.role || '').toLowerCase().trim();
      let role;
      if (ROLE_SPECIALIST.some(r => roleRaw.includes(r))) role = "specijalista";
      else if (ROLE_RESIDENT.some(r => roleRaw.includes(r))) role = "specijalizant";
      else throw new Error(`Red ${lineNum}: Neispravan tip zaposlenog "${row.role}" za ${row.name}.`);
      
      const shareRaw = String(row.share || '').toLowerCase().trim();
      const mappedShare = SHARE_MAP[shareRaw];
      if (!mappedShare) throw new Error(`Red ${lineNum}: Neispravan udeo dežurstva "${row.share}" za ${row.name}.`);
      
      const rank = row.rank ? parseInt(row.rank) || 0 : 0;
      
      // Parsiraj datume koristeći parseDateList iz utils
      const unavailableDates = parseDateList(row.unavailable || '');
      const preferredDates = parseDateList(row.preferred || '');
      
      console.log(`   👨‍⚕️ ${row.name}: ${role}, udeo=${mappedShare}, glavni=${canBeChief ? 'DA' : 'NE'}, rang=${rank || '-'}`);
      if (unavailableDates.length > 0) console.log(`      Ne može: ${unavailableDates.join(', ')}`);
      if (preferredDates.length > 0) console.log(`      Želi: ${preferredDates.join(', ')}`);
      
      return {
        name: row.name,
        seniority: seniority,
        canBeChief: canBeChief,
        role: role,
        dutyShare: mappedShare,
        unavailable: unavailableDates,
        preferred: preferredDates,
        rank: rank,
        assignedDates: [],
        weekendCount: 0
      };
    });
  } catch (error) {
    if (error.message.includes("Red") || error.message.includes("Line")) throw error;
    throw new Error(`Greška pri čitanju Excel fajla: ${error.message}`);
  }
};