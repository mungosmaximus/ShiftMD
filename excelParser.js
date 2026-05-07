// ShiftMD v3.0 - Dvojezični Excel parser (podržava SR i EN zaglavlja)
const ExcelJS = require("exceljs");
const { parseDateList } = require("./utils");

console.log("✅ ShiftMD v3.0 - Dvojezični Excel parser");

// Mapiranje zaglavlja: srpski i engleski nazivi kolona
const HEADER_MAP = {
  // Ime i prezime
  "ime i prezime": "name",
  "full name": "name",
  "name": "name",
  
  // Starešinstvo
  "starešinstvo": "seniority",
  "seniority": "seniority",
  "senioritet": "seniority",
  "years of experience": "seniority",
  "experience": "seniority",
  
  // Može biti glavni
  "može biti glavni": "canBeChief",
  "can be chief": "canBeChief",
  "glavni": "canBeChief",
  "chief": "canBeChief",
  
  // Tip zaposlenog
  "tip zaposlenog": "role",
  "employee type": "role",
  "role": "role",
  "position": "role",
  "tip": "role",
  "type": "role",
  
  // Udeo dežurstva
  "udeo dežurstva": "share",
  "duty share": "share",
  "share": "share",
  "udeo": "share",
  
  // Ne može
  "ne može": "unavailable",
  "unavailable": "unavailable",
  "ne moze": "unavailable",
  "cannot work": "unavailable",
  
  // Želi da dežura
  "želi da dežura": "preferred",
  "preferred": "preferred",
  "želi": "preferred",
  "wants to work": "preferred",
  "wishes": "preferred",
  "zeli da dezura": "preferred"
};

// Mapiranje vrednosti za "Može biti glavni"
const CHIEF_VALUES_SR = ["da", "može", "moze"];
const CHIEF_VALUES_EN = ["yes", "y", "true", "1"];

// Mapiranje vrednosti za "Tip zaposlenog"
const ROLE_SPECIALIST = ["specijalista", "specialist", "specijalista", "attending"];
const ROLE_RESIDENT = ["specijalizant", "resident", "specijalizant", "trainee", "fellow"];

// Mapiranje vrednosti za "Udeo dežurstva"
const SHARE_MAP = {
  // Srpski
  "celo": "celo",
  "polovina": "polovina",
  "trećina": "trećina",
  "cetvrtina": "četvrtina",
  "četvrtina": "četvrtina",
  "osmina": "osmina",
  // Engleski
  "full": "celo",
  "whole": "celo",
  "half": "polovina",
  "third": "trećina",
  "quarter": "četvrtina",
  "eighth": "osmina",
  // Stari formati (razlomci)
  "1/1": "celo",
  "1": "celo",
  "1/2": "polovina",
  "0.5": "polovina",
  "1/3": "trećina",
  "1/4": "četvrtina",
  "0.25": "četvrtina",
  "1/8": "osmina",
  "0.125": "osmina"
};

module.exports = async function parseExcel(path) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path);
    
    const worksheet = workbook.worksheets[0];
    
    if (!worksheet) {
      throw new Error("Excel fajl je prazan ili nema listova. / Excel file is empty or has no sheets.");
    }
    
    // Čitaj header row i mapiraj kolone
    const headerRow = worksheet.getRow(1);
    const columnMap = {}; // { colNumber: fieldKey }
    
    headerRow.eachCell((cell, colNumber) => {
      const rawHeader = cell.value;
      if (!rawHeader) return;
      
      // Normalizuj header: trim, lowercase
      const headerStr = String(rawHeader).trim().toLowerCase();
      
      // Pronađi mapiranje
      const fieldKey = HEADER_MAP[headerStr];
      
      if (fieldKey) {
        columnMap[colNumber] = fieldKey;
        console.log(`   📋 Kolona ${colNumber}: "${rawHeader}" → ${fieldKey}`);
      } else {
        console.warn(`   ⚠️  Nepoznata kolona: "${rawHeader}" (kolona ${colNumber})`);
      }
    });
    
    // Proveri da li su sve obavezne kolone prisutne
    const requiredFields = ["name", "seniority", "canBeChief", "role", "share"];
    const missingFields = requiredFields.filter(f => !Object.values(columnMap).includes(f));
    
    if (missingFields.length > 0) {
      throw new Error(
        `Nedostaju obavezne kolone / Missing required columns: ${missingFields.join(", ")}. ` +
        `Potrebno / Required: Ime i prezime/Full Name, Starešinstvo/Seniority, ` +
        `Može biti glavni/Can be Chief, Tip zaposlenog/Employee Type, Udeo dežurstva/Duty Share`
      );
    }
    
    // Čitaj podatke
    const rows = [];
    const totalRows = worksheet.rowCount;
    
    for (let rowNumber = 2; rowNumber <= totalRows; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      
      // Preskoči prazne redove
      let hasData = false;
      row.eachCell(() => { hasData = true; });
      if (!hasData) continue;
      
      const rowData = {};
      
      row.eachCell((cell, colNumber) => {
        const fieldKey = columnMap[colNumber];
        if (!fieldKey) return;
        
        let cellValue = cell.value;
        
        // Konvertuj Date objekte u string (za datume)
        if (cellValue instanceof Date) {
          const d = cellValue.getDate().toString().padStart(2, '0');
          const m = (cellValue.getMonth() + 1).toString().padStart(2, '0');
          const y = cellValue.getFullYear();
          cellValue = `${d}-${m}-${y}`;
        } else if (typeof cellValue === 'number') {
          cellValue = cellValue.toString();
        } else if (cellValue && typeof cellValue === 'object' && cellValue.richText) {
          cellValue = cellValue.richText.map(t => t.text).join('');
        } else if (cellValue === null || cellValue === undefined) {
          cellValue = '';
        } else {
          cellValue = String(cellValue).trim();
        }
        
        rowData[fieldKey] = cellValue;
      });
      
      // Preskoči redove bez imena
      if (!rowData.name || rowData.name === '') continue;
      
      rows.push(rowData);
    }
    
    if (rows.length === 0) {
      throw new Error("Excel fajl nema podataka. / Excel file has no data.");
    }
    
    console.log(`   ✅ Učitano ${rows.length} redova podataka`);
    
    // Validiraj i konvertuj svaki red
    return rows.map((row, index) => {
      const lineNum = index + 2; // +2 jer je header red 1
      
      // Validacija imena
      if (!row.name) {
        throw new Error(`Red ${lineNum}: Nedostaje ime / Missing name.`);
      }
      
      // Validacija i konverzija starešinstva
      const seniority = Number(row.seniority);
      if (isNaN(seniority)) {
        throw new Error(`Red ${lineNum}: Neispravno starešinstvo za ${row.name}. / Invalid seniority.`);
      }
      
      // Validacija "Može biti glavni"
      const canBeChiefRaw = String(row.canBeChief || '').toLowerCase().trim();
      const canBeChief = CHIEF_VALUES_SR.includes(canBeChiefRaw) || CHIEF_VALUES_EN.includes(canBeChiefRaw);
      
      // Validacija "Tip zaposlenog"
      const roleRaw = String(row.role || '').toLowerCase().trim();
      let role;
      if (ROLE_SPECIALIST.some(r => roleRaw.includes(r))) {
        role = "specijalista";
      } else if (ROLE_RESIDENT.some(r => roleRaw.includes(r))) {
        role = "specijalizant";
      } else {
        throw new Error(
          `Red ${lineNum}: Neispravan tip zaposlenog "${row.role}" za ${row.name}. ` +
          `Dozvoljene vrednosti: specijalista/specialist, specijalizant/resident.`
        );
      }
      
      // Validacija "Udeo dežurstva"
      const shareRaw = String(row.share || '').toLowerCase().trim();
      const mappedShare = SHARE_MAP[shareRaw];
      
      if (!mappedShare) {
        throw new Error(
          `Red ${lineNum}: Neispravan udeo dežurstva "${row.share}" za ${row.name}. ` +
          `Dozvoljene vrednosti: celo/full, polovina/half, trećina/third, četvrtina/quarter, osmina/eighth.`
        );
      }
      
      console.log(`   👨‍⚕️ ${row.name}: ${role}, udeo=${mappedShare}, glavni=${canBeChief ? 'DA' : 'NE'}`);
      
      return {
        name: row.name,
        seniority: seniority,
        canBeChief: canBeChief,
        role: role,
        dutyShare: mappedShare,
        unavailable: parseDateList(row.unavailable || ''),
        preferred: parseDateList(row.preferred || ''),
        assignedDates: [],
        weekendCount: 0
      };
    });
    
  } catch (error) {
    if (error.message.includes("Red") || error.message.includes("Line")) {
      throw error;
    }
    throw new Error(`Greška pri čitanju Excel fajla: ${error.message}`);
  }
};