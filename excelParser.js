// ShiftMD v3.6 - Dvojezični Excel parser sa podrškom za rangove i opsege datuma
const ExcelJS = require("exceljs");
const { parseDateList } = require("./utils");

console.log("✅ ShiftMD v3.6 - Dvojezični Excel parser (sa rangovima i opsezima datuma)");

const HEADER_MAP = {
  "ime i prezime": "name", "full name": "name", "name": "name",
  "starešinstvo": "seniority", "seniority": "seniority", "senioritet": "seniority",
  "years of experience": "seniority", "experience": "seniority",
  "može biti glavni": "canBeChief", "can be chief": "canBeChief", "glavni": "canBeChief", "chief": "canBeChief",
  "tip zaposlenog": "role", "employee type": "role", "role": "role", "position": "role", "tip": "role", "type": "role",
  "udeo dežurstva": "share", "duty share": "share", "share": "share", "udeo": "share",
  "ne može": "unavailable", "unavailable": "unavailable", "ne moze": "unavailable", "cannot work": "unavailable",
  "želi da dežura": "preferred", "preferred": "preferred", "želi": "preferred", "wants to work": "preferred",
  "wishes": "preferred", "zeli da dezura": "preferred",
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
 * Konvertuje Excel serijski broj u datum DD-MM-YYYY
 */
function excelSerialToDate(serial) {
  const excelEpoch = new Date(1899, 11, 30);
  const date = new Date(excelEpoch.getTime() + serial * 86400000);
  const d = date.getDate().toString().padStart(2, '0');
  const m = (date.getMonth() + 1).toString().padStart(2, '0');
  const y = date.getFullYear();
  return `${d}-${m}-${y}`;
}

/**
 * Konvertuje vrednost ćelije u string.
 * Date objekti i Excel serijski brojevi se konvertuju u DD-MM-YYYY.
 * Stringovi se vraćaju NETAKNUTI (mogu sadržati opsege sa "do"/"to").
 */
function cellValueToString(cellValue) {
  if (cellValue === null || cellValue === undefined) {
    return '';
  }

  // Date objekat → DD-MM-YYYY
  if (cellValue instanceof Date) {
    const d = cellValue.getDate().toString().padStart(2, '0');
    const m = (cellValue.getMonth() + 1).toString().padStart(2, '0');
    const y = cellValue.getFullYear();
    return `${d}-${m}-${y}`;
  }

  // Broj → proveri da li je Excel serijski broj datuma
  if (typeof cellValue === 'number') {
    if (cellValue > 40000 && cellValue < 60000) {
      return excelSerialToDate(cellValue);
    }
    return cellValue.toString();
  }

  // Rich text → običan tekst
  if (cellValue && typeof cellValue === 'object' && cellValue.richText) {
    return cellValue.richText.map(t => t.text).join('');
  }

  // String → vrati ga NETAKNUTOG (može sadržati "DD-MM-YYYY do DD-MM-YYYY")
  return String(cellValue).trim();
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
        
        // Konvertuj vrednost ćelije u string (čuvajući opsege!)
        const cellValue = cellValueToString(cell.value);
        rowData[fieldKey] = cellValue;
      });
      
      if (!rowData.name || rowData.name === '') continue;
      rows.push(rowData);
    }
    
    if (rows.length === 0) throw new Error("Excel fajl nema podataka. / Excel file has no data.");
    console.log(`   ✅ Učitano ${rows.length} redova podataka`);
    
    return rows.map((row, index) => {
      const lineNum = index + 2;
      
      // Validacija imena
      if (!row.name) {
        throw new Error(`Red ${lineNum}: Nedostaje ime / Missing name.`);
      }
      
      // Validacija starešinstva
      const seniority = Number(row.seniority);
      if (isNaN(seniority)) {
        throw new Error(`Red ${lineNum}: Neispravno starešinstvo "${row.seniority}" za ${row.name}. / Invalid seniority.`);
      }
      
      // Validacija "Može biti glavni"
      const canBeChiefRaw = String(row.canBeChief || '').toLowerCase().trim();
      const canBeChief = CHIEF_VALUES.some(v => canBeChiefRaw === v);
      
      // Validacija "Tip zaposlenog"
      const roleRaw = String(row.role || '').toLowerCase().trim();
      let role;
      if (ROLE_SPECIALIST.some(r => roleRaw.includes(r))) {
        role = "specijalista";
      } else if (ROLE_RESIDENT.some(r => roleRaw.includes(r))) {
        role = "specijalizant";
      } else {
        throw new Error(`Red ${lineNum}: Neispravan tip zaposlenog "${row.role}" za ${row.name}. / Invalid employee type.`);
      }
      
      // Validacija "Udeo dežurstva"
      const shareRaw = String(row.share || '').toLowerCase().trim();
      const mappedShare = SHARE_MAP[shareRaw];
      if (!mappedShare) {
        throw new Error(`Red ${lineNum}: Neispravan udeo dežurstva "${row.share}" za ${row.name}. / Invalid duty share.`);
      }
      
      // Parsiraj rang
      const rank = row.rank ? parseInt(row.rank) || 0 : 0;
      
      // Parsiraj datume koristeći parseDateList iz utils.js
      // parseDateList sada podržava opsege: "DD-MM-YYYY do DD-MM-YYYY"
      const unavailableDates = parseDateList(row.unavailable || '');
      const preferredDates = parseDateList(row.preferred || '');
      
      console.log(`   👨‍⚕️ ${row.name}: ${role}, udeo=${mappedShare}, glavni=${canBeChief ? 'DA' : 'NE'}, rang=${rank || '-'}`);
      if (unavailableDates.length > 0) console.log(`      Ne može: ${unavailableDates.length} datuma`);
      if (preferredDates.length > 0) console.log(`      Želi: ${preferredDates.length} datuma`);
      
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