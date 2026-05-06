// ShiftMD v2.0 - Excel Generator
const ExcelJS = require("exceljs");

const MONTHS = {
  "1": "Januar", "2": "Februar", "3": "Mart", "4": "April",
  "5": "Maj", "6": "Jun", "7": "Jul", "8": "Avgust",
  "9": "Septembar", "10": "Oktobar", "11": "Novembar", "12": "Decembar"
};

const DAYS = ["Ponedeljak", "Utorak", "Sreda", "Četvrtak", "Petak", "Subota", "Nedelja"];

module.exports = async function generateExcel(data, filePath, statistics = null) {
  const { hospital, month, year, schedule, dutyStaffCount, allowLessSpecialists } = data;
  
  const monthName = MONTHS[month] || month;
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "ShiftMD v2.0";
  workbook.created = new Date();
  
  const specialistRule = allowLessSpecialists 
    ? "Obavezan glavni specijalista | Dozvoljeno manje od 50% specijalista"
    : "Minimum 50% specijalista u svakom dežurstvu";
  
  const genDate = new Date();
  
  const dates = Object.keys(schedule).sort((a, b) => {
    const [da, ma, ya] = a.split("-").map(Number);
    const [db, mb, yb] = b.split("-").map(Number);
    return new Date(ya, ma - 1, da) - new Date(yb, mb - 1, db);
  });
  
  // =====================================================================
  // SHEET 1: RASPORED
  // =====================================================================
  const wsSchedule = workbook.addWorksheet("Raspored", {
    pageSetup: {
      orientation: "landscape", paperSize: 9,
      fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      margins: { left: 1.5, right: 1.5, top: 1.5, bottom: 1.5, header: 0.5, footer: 0.5 }
    }
  });
  
  wsSchedule.getColumn(1).width = 80;
  wsSchedule.getColumn(2).width = 25;
  wsSchedule.getColumn(3).width = 22;
  wsSchedule.getColumn(4).width = 22;
  
  let r = 1;
  
  wsSchedule.mergeCells(`A${r}:D${r}`);
  wsSchedule.getCell(`A${r}`).value = "RASPORED DEŽURSTAVA";
  wsSchedule.getCell(`A${r}`).font = { bold: true, size: 20, name: "Arial", color: { argb: "FF1E3A5F" } };
  wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center", vertical: "middle" };
  wsSchedule.getRow(r).height = 35; r++;
  
  wsSchedule.mergeCells(`A${r}:D${r}`);
  wsSchedule.getCell(`A${r}`).value = hospital;
  wsSchedule.getCell(`A${r}`).font = { bold: true, size: 16, name: "Arial", color: { argb: "FF2563EB" } };
  wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center", vertical: "middle" };
  wsSchedule.getRow(r).height = 28; r++;
  
  wsSchedule.mergeCells(`A${r}:D${r}`);
  wsSchedule.getCell(`A${r}`).value = `${monthName} ${year}. godine`;
  wsSchedule.getCell(`A${r}`).font = { bold: true, size: 14, name: "Arial", color: { argb: "FF1F2937" } };
  wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center", vertical: "middle" };
  wsSchedule.getRow(r).height = 25; r++;
  
  wsSchedule.mergeCells(`A${r}:D${r}`);
  wsSchedule.getCell(`A${r}`).value = `Broj dežurstava po danu: ${dutyStaffCount} | ${specialistRule}`;
  wsSchedule.getCell(`A${r}`).font = { size: 10, name: "Arial", color: { argb: "FF6B7280" } };
  wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" };
  r++;
  
  wsSchedule.mergeCells(`A${r}:D${r}`);
  wsSchedule.getCell(`A${r}`).value = `Generisano: ${genDate.toLocaleDateString("sr-RS")} u ${genDate.toLocaleTimeString("sr-RS")} | ShiftMD v2.0`;
  wsSchedule.getCell(`A${r}`).font = { size: 8, name: "Arial", color: { argb: "FF9CA3AF" } };
  wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" };
  r++; r++;
  
  if (statistics && statistics.entries && statistics.entries.length > 0) {
    wsSchedule.mergeCells(`A${r}:D${r}`);
    wsSchedule.getCell(`A${r}`).value = "📊 STATISTIKA OPTEREĆENJA";
    wsSchedule.getCell(`A${r}`).font = { bold: true, size: 12, name: "Arial", color: { argb: "FF374151" } };
    r++;
    
    const E = statistics.entries;
    const maxL = Math.max(...E.map(e => e.totalAssignedShare));
    const minL = Math.min(...E.map(e => e.totalAssignedShare));
    const avgL = E.reduce((s, e) => s + e.totalAssignedShare, 0) / E.length;
    const maxD = Math.max(...E.map(e => e.dutyCount || e.assignedDates.length));
    const minD = Math.min(...E.map(e => e.dutyCount || e.assignedDates.length));
    const avgD = E.reduce((s, e) => s + (e.dutyCount || e.assignedDates.length), 0) / E.length;
    const maxW = Math.max(...E.map(e => e.weekendCount));
    const minW = Math.min(...E.map(e => e.weekendCount));
    const avgW = E.reduce((s, e) => s + e.weekendCount, 0) / E.length;
    
    ["", "Opterećenje (skor)", "Broj dežurstava", "Vikend dežurstva"].forEach((h, i) => {
      const cell = wsSchedule.getCell(r, i + 1);
      cell.value = h;
      cell.font = { bold: true, size: 11, name: "Arial", color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF374151" } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    });
    r++;
    
    [
      ["Maksimum", maxL.toFixed(2), String(maxD), String(maxW), "FFFEE2E2"],
      ["Prosek", avgL.toFixed(2), avgD.toFixed(1), avgW.toFixed(1), "FFF3F4F6"],
      ["Minimum", minL.toFixed(2), String(minD), String(minW), "FFDCFCE7"]
    ].forEach(row => {
      row.slice(0, 4).forEach((val, i) => {
        const cell = wsSchedule.getCell(r, i + 1);
        cell.value = val;
        cell.font = { size: 11, name: "Arial" };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: row[4] } };
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      });
      r++;
    });
    
    const loaded = E.filter(e => e.totalAssignedShare > 0).length;
    wsSchedule.mergeCells(`A${r}:D${r}`);
    wsSchedule.getCell(`A${r}`).value = `Ukupno lekara: ${E.length} | Sa dežurstvima: ${loaded} | Bez: ${E.length - loaded}`;
    wsSchedule.getCell(`A${r}`).font = { size: 9, name: "Arial", color: { argb: "FF6B7280" } };
    wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" };
    r++;
    
    wsSchedule.mergeCells(`A${r}:D${r}`);
    wsSchedule.getCell(`A${r}`).value = `Period: ${statistics.firstDate || "—"} do ${statistics.lastDate || "—"} (${statistics.totalDays || 0} dana)`;
    wsSchedule.getCell(`A${r}`).font = { size: 9, name: "Arial", color: { argb: "FF9CA3AF" } };
    wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" };
    r++; r++;
    
    wsSchedule.mergeCells(`A${r}:D${r}`);
    wsSchedule.getCell(`A${r}`).value = "═".repeat(100);
    wsSchedule.getCell(`A${r}`).font = { size: 8, name: "Arial", color: { argb: "FFD1D5DB" } };
    wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" };
    r++; r++;
  }
  
  for (const date of dates) {
    const [day, m, y] = date.split("-").map(Number);
    const dateObj = new Date(y, m - 1, day);
    const dayName = DAYS[dateObj.getDay() === 0 ? 6 : dateObj.getDay() - 1];
    const isW = dateObj.getDay() === 0 || dateObj.getDay() === 6;
    
    wsSchedule.mergeCells(`A${r}:D${r}`);
    const dc = wsSchedule.getCell(`A${r}`);
    dc.value = `${date} (${dayName})`;
    dc.font = { bold: true, size: 13, name: "Arial", color: { argb: isW ? "FFDC2626" : "FF1F2937" } };
    dc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: isW ? "FFFEE2E2" : "FFF3F4F6" } };
    dc.alignment = { horizontal: "left", vertical: "middle" };
    wsSchedule.getRow(r).height = 22; r++;
    
    if (schedule[date]) {
      schedule[date].forEach((slot, si) => {
        wsSchedule.mergeCells(`A${r}:D${r}`);
        wsSchedule.getCell(`A${r}`).value = `  Dežurstvo ${si + 1}:`;
        wsSchedule.getCell(`A${r}`).font = { bold: true, size: 12, name: "Arial", color: { argb: "FF374151" } };
        r++;
        
        if (slot.persons) {
          slot.persons.forEach(person => {
            wsSchedule.mergeCells(`A${r}:D${r}`);
            const cell = wsSchedule.getCell(`A${r}`);
            cell.value = `    ${person.isChief ? "★ " : "• "}${person.name} — ${person.role} (${person.share})${person.isChief ? " — GLAVNI DEŽURNI" : ""}`;
            cell.font = { bold: person.isChief, size: 11, name: "Arial", color: { argb: person.isChief ? "FF059669" : "FF1F2937" } };
            if (person.isChief) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF0FDF4" } };
            r++;
          });
        }
        r++;
      });
    }
    r++;
  }
  
  wsSchedule.mergeCells(`A${r}:D${r}`);
  wsSchedule.getCell(`A${r}`).value = `© ShiftMD v2.0 | Generisano: ${genDate.toLocaleDateString("sr-RS")} u ${genDate.toLocaleTimeString("sr-RS")}`;
  wsSchedule.getCell(`A${r}`).font = { size: 7, name: "Arial", color: { argb: "FFD1D5DB" } };
  wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" };
  
    // =====================================================================
  // SHEET 2: TABELARNI PREGLED
  // =====================================================================
  const wsTable = workbook.addWorksheet("Tabelarni pregled", {
    pageSetup: {
      orientation: "landscape",
      paperSize: 9,
      fitToPage: false,        // ← KLJUČNO: isključi auto-fit!
      fitToWidth: 0,           // ← ne prilagođavaj širinu stranici
      fitToHeight: 0,          // ← ne prilagođavaj visinu stranici
      margins: { left: 1.5, right: 1.5, top: 1.5, bottom: 1.5 }
    }
  });
  
  // SADA će ove širine biti poštovane
  wsTable.getColumn(1).width = 16;   // Datum
  wsTable.getColumn(2).width = 18;   // Dan
  wsTable.getColumn(3).width = 18;   // Dežurstvo
  wsTable.getColumn(4).width = 38;   // Ime i prezime
  wsTable.getColumn(5).width = 20;   // Tip
  wsTable.getColumn(6).width = 16;   // Udeo
  wsTable.getColumn(7).width = 18;   // Glavni dežurni
  
  wsTable.mergeCells("A1:G1");
  wsTable.getCell("A1").value = `${hospital} — Raspored dežurstava: ${monthName} ${year}.`;
  wsTable.getCell("A1").font = { bold: true, size: 14, name: "Arial", color: { argb: "FF1E3A5F" } };
  wsTable.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  wsTable.getRow(1).height = 30;
  
  wsTable.mergeCells("A2:G2");
  wsTable.getCell("A2").value = `Broj dežurstava po danu: ${dutyStaffCount} | ${specialistRule} | ShiftMD v2.0`;
  wsTable.getCell("A2").font = { size: 9, name: "Arial", color: { argb: "FF6B7280" } };
  wsTable.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };
  
  const hRow = wsTable.addRow(["Datum", "Dan", "Dežurstvo", "Ime i prezime", "Tip", "Udeo", "Glavni dežurni"]);
  hRow.font = { bold: true, size: 11, name: "Arial", color: { argb: "FFFFFFFF" } };
  hRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF374151" } };
  hRow.alignment = { horizontal: "center", vertical: "middle" };
  hRow.height = 25;
  
  for (const date of dates) {
    const [day, m, y] = date.split("-").map(Number);
    const dateObj = new Date(y, m - 1, day);
    const dayName = DAYS[dateObj.getDay() === 0 ? 6 : dateObj.getDay() - 1];
    const isWeekendDay = dateObj.getDay() === 0 || dateObj.getDay() === 6;
    
    if (schedule[date]) {
      schedule[date].forEach((slot, slotIndex) => {
        if (slot.persons) {
          slot.persons.forEach((person) => {
            const row = wsTable.addRow([
              date,
              dayName,
              `Dežurstvo ${slotIndex + 1}`,
              person.name,
              person.role === "specijalista" ? "Specijalista" : "Specijalizant",
              person.share,
              person.isChief ? "★ DA" : "NE"
            ]);
            
            if (isWeekendDay) row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF2F2" } };
            if (person.isChief) {
              row.font = { bold: true, color: { argb: "FF059669" } };
              row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF0FDF4" } };
            }
            if (person.role === "specijalista" && !person.isChief) {
              row.getCell(5).font = { color: { argb: "FF2563EB" }, bold: true };
            }
            row.alignment = { vertical: "middle" };
            row.height = 20;
          });
        }
      });
    }
  }
  
  wsTable.eachRow((row, rowNumber) => {
    if (rowNumber <= 3) return;
    row.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFD1D5DB" } },
        left: { style: "thin", color: { argb: "FFD1D5DB" } },
        bottom: { style: "thin", color: { argb: "FFD1D5DB" } },
        right: { style: "thin", color: { argb: "FFD1D5DB" } }
      };
    });
  });
  
  // =====================================================================
  // SHEET 3: STATISTIKA
  // =====================================================================
  if (statistics && statistics.entries && statistics.entries.length > 0) {
    const wsStats = workbook.addWorksheet("Statistika");
    
    wsStats.getColumn(1).width = 35;
    wsStats.getColumn(2).width = 18;
    wsStats.getColumn(3).width = 16;
    wsStats.getColumn(4).width = 20;
    wsStats.getColumn(5).width = 20;
    wsStats.getColumn(6).width = 16;
    
    wsStats.mergeCells("A1:F1");
    wsStats.getCell("A1").value = `STATISTIKA OPTEREĆENJA — ${hospital}`;
    wsStats.getCell("A1").font = { bold: true, size: 14, name: "Arial" };
    wsStats.getCell("A1").alignment = { horizontal: "center" };
    
    wsStats.mergeCells("A2:F2");
    wsStats.getCell("A2").value = `${monthName} ${year}. | ShiftMD v2.0 | Generisano: ${genDate.toLocaleDateString("sr-RS")}`;
    wsStats.getCell("A2").font = { size: 10, name: "Arial", color: { argb: "FF6B7280" } };
    wsStats.getCell("A2").alignment = { horizontal: "center" };
    
    const sRow = wsStats.addRow(["Ime i prezime", "Tip", "Opterećenje", "Broj dežurstava", "Vikend dežurstva", "Može glavni"]);
    sRow.font = { bold: true, size: 11, name: "Arial", color: { argb: "FFFFFFFF" } };
    sRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF374151" } };
    sRow.alignment = { horizontal: "center", vertical: "middle" };
    
    [...statistics.entries].sort((a, b) => b.totalAssignedShare - a.totalAssignedShare).forEach((entry) => {
      const row = wsStats.addRow([
        entry.name,
        entry.role === "specijalista" ? "Specijalista" : "Specijalizant",
        entry.totalAssignedShare.toFixed(2),
        entry.dutyCount || entry.assignedDates.length,
        entry.weekendCount,
        entry.canBeChief ? "DA" : "NE"
      ]);
      
      if (entry.totalAssignedShare > 0) {
        const intensity = Math.min(Math.round(entry.totalAssignedShare * 30), 200);
        const green = (255 - intensity).toString(16).padStart(2, '0');
        row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${green}FF${green}` } };
      }
      row.alignment = { vertical: "middle" };
    });
    
    wsStats.eachRow((row, rowNumber) => {
      if (rowNumber <= 2) return;
      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin" }, left: { style: "thin" },
          bottom: { style: "thin" }, right: { style: "thin" }
        };
      });
    });
  }
  
  await workbook.xlsx.writeFile(filePath);
  console.log(`   ✅ Excel sačuvan: ${filePath}`);
  return true;
};