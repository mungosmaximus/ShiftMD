const ExcelJS = require("exceljs");
const { parseDateList } = require("./utils");

console.log("✅ ShiftMD v2.0 - Excel parser učitan");

module.exports = async function parseExcel(path) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path);
    
    const worksheet = workbook.worksheets[0];
    
    if (!worksheet) {
      throw new Error("Excel fajl je prazan ili nema listova.");
    }
    
    const rows = [];
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      
      const rowData = {};
      row.eachCell((cell, colNumber) => {
        const headerCell = worksheet.getRow(1).getCell(colNumber);
        let headerValue = headerCell.value;
        
        if (typeof headerValue === 'string') {
          headerValue = headerValue.trim();
        }
        
        let cellValue = cell.value;
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
        
        rowData[headerValue] = cellValue;
      });
      
      if (rowData["Ime i prezime"]) {
        rows.push(rowData);
      }
    });
    
    if (!rows || rows.length === 0) {
      throw new Error("Excel fajl nema podataka.");
    }

    return rows.map((row, index) => {
      if (!row["Ime i prezime"]) {
        throw new Error(`Red ${index + 2}: Nedostaje "Ime i prezime".`);
      }
      
      if (!row["Tip zaposlenog"]) {
        throw new Error(`Red ${index + 2}: Nedostaje "Tip zaposlenog" za ${row["Ime i prezime"]}.`);
      }
      
      const role = String(row["Tip zaposlenog"]).toLowerCase().trim();
      if (role !== "specijalista" && role !== "specijalizant") {
        throw new Error(
          `Red ${index + 2}: Neispravan tip zaposlenog "${row["Tip zaposlenog"]}" za ${row["Ime i prezime"]}. ` +
          `Dozvoljene vrednosti: specijalista, specijalizant.`
        );
      }
      
      if (!row["Udeo dežurstva"] && row["Udeo dežurstva"] !== 0) {
        throw new Error(`Red ${index + 2}: Nedostaje "Udeo dežurstva" za ${row["Ime i prezime"]}.`);
      }
      
      const shareValue = String(row["Udeo dežurstva"]).toLowerCase().trim();
      
      const shareMapping = {
        "1/1": "celo", "1": "celo", "celo": "celo",
        "1/2": "polovina", "0.5": "polovina", "polovina": "polovina",
        "1/3": "trećina", "trećina": "trećina",
        "1/4": "četvrtina", "0.25": "četvrtina", "četvrtina": "četvrtina",
        "1/8": "osmina", "0.125": "osmina", "osmina": "osmina"
      };
      
      const mappedShare = shareMapping[shareValue];
      
      console.log(`   Udeo za ${row["Ime i prezime"]}: "${shareValue}" → "${mappedShare}"`);
      
      if (!mappedShare) {
        throw new Error(
          `Red ${index + 2}: Neispravan udeo dežurstva "${row["Udeo dežurstva"]}" za ${row["Ime i prezime"]}. ` +
          `Dozvoljene vrednosti: celo, polovina, trećina, četvrtina, osmina.`
        );
      }

      return {
        name: String(row["Ime i prezime"]).trim(),
        seniority: Number(row["Starešinstvo"]) || 0,
        canBeChief: String(row["Može biti glavni"]).toUpperCase() === "DA",
        role: role,
        dutyShare: mappedShare,
        unavailable: parseDateList(row["Ne može"]),
        preferred: parseDateList(row["Želi da dežura"]),
        assignedDates: [],
        weekendCount: 0
      };
    });
  } catch (error) {
    if (error.message.includes("Red")) {
      throw error;
    }
    throw new Error(`Greška pri čitanju Excel fajla: ${error.message}`);
  }
};