// ShiftMD v3.6 - Dvojezični Excel Generator
const ExcelJS = require("exceljs");

const MONTHS = {
  sr: { "1": "Januar", "2": "Februar", "3": "Mart", "4": "April", "5": "Maj", "6": "Jun", "7": "Jul", "8": "Avgust", "9": "Septembar", "10": "Oktobar", "11": "Novembar", "12": "Decembar" },
  en: { "1": "January", "2": "February", "3": "March", "4": "April", "5": "May", "6": "June", "7": "July", "8": "August", "9": "September", "10": "October", "11": "November", "12": "December" }
};

const DAYS = {
  sr: ["Ponedeljak", "Utorak", "Sreda", "Četvrtak", "Petak", "Subota", "Nedelja"],
  en: ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
};

const T = {
  sr: {
    sheet_raspored: "Raspored", sheet_table: "Tabelarni pregled", sheet_stats: "Statistika",
    title: "RASPORED DEŽURSTAVA", period: "godine", duty_count: "Broj dežurstava po danu",
    rule_strict: "Minimum 50% specijalista u svakom dežurstvu",
    rule_relaxed: "Obavezan glavni specijalista | Dozvoljeno manje od 50% specijalista",
    generated: "Generisano", stats_title: "📊 STATISTIKA OPTEREĆENJA",
    stats_load: "Opterećenje (skor)", stats_duties: "Broj dežurstava", stats_weekends: "Vikend dežurstva",
    max: "Maksimum", avg: "Prosek", min: "Minimum",
    total_doctors: "Ukupno lekara", with_duties: "Sa dežurstvima", without: "Bez",
    period_label: "Period", days: "dana", duty_slot: "Dežurni", chief: "GLAVNI DEŽURNI",
    specialist: "Specijalista", resident: "Specijalizant", yes: "DA", no: "NE",
    header_date: "Datum", header_day: "Dan", header_slot: "Dežurni",
    header_name: "Ime i prezime", header_type: "Tip", header_share: "Udeo",
    header_chief: "Glavni dežurni", header_rank: "Rang", header_duty_type: "Tip dežurstva",
    footer: "ShiftMD v3.6",
  },
  en: {
    sheet_raspored: "Schedule", sheet_table: "Table View", sheet_stats: "Statistics",
    title: "DUTY SCHEDULE", period: "", duty_count: "Duties per day",
    rule_strict: "Minimum 50% specialists in each duty",
    rule_relaxed: "Mandatory chief specialist | Less than 50% specialists allowed",
    generated: "Generated", stats_title: "📊 WORKLOAD STATISTICS",
    stats_load: "Workload (score)", stats_duties: "Number of Duties", stats_weekends: "Weekend Duties",
    max: "Maximum", avg: "Average", min: "Minimum",
    total_doctors: "Total Physicians", with_duties: "With Duties", without: "Without",
    period_label: "Period", days: "days", duty_slot: "On-call", chief: "CHIEF ON CALL",
    specialist: "Specialist", resident: "Resident", yes: "YES", no: "NO",
    header_date: "Date", header_day: "Day", header_slot: "On-call",
    header_name: "Full Name", header_type: "Type", header_share: "Share",
    header_chief: "Chief on Call", header_rank: "Rank", header_duty_type: "Duty Type",
    footer: "ShiftMD v3.6",
  }
};

const COLORS = {
  first_row_bg: "FFDBEAFE",  // Svetlo plava za prvi red dana
  chief_bg: "FFF0FDF4",      // Glavni dežurni - zelenkasta
  text_black: "FF1F2937",    // Crna boja za sav tekst
  white: "FFFFFFFF",         // Bela
};

function detectLanguage(hospital) {
  if (!hospital) return 'sr';
  const lower = hospital.toLowerCase();
  const enWords = ['hospital', 'clinic', 'medical', 'center', 'centre', 'health', 'general', 'university', 'institute'];
  const srWords = ['bolnica', 'dom', 'zdravlja', 'klinički', 'centar', 'zavod', 'opšta', 'univerzitetski', 'institut'];
  let enScore = 0, srScore = 0;
  enWords.forEach(w => { if (lower.includes(w)) enScore++; });
  srWords.forEach(w => { if (lower.includes(w)) srScore++; });
  return enScore > srScore ? 'en' : 'sr';
}

module.exports = async function generateExcel(data, filePath, statistics = null) {
  const { hospital, month, year, schedule, dutyTypes, hasType2, type1Name, type2Name, staffCount1, staffCount2, allowLess1, allowLess2, useRanked1, useRanked2, outputLang } = data;
  
  const lang = outputLang || detectLanguage(hospital);
  const dict = T[lang];
  const months = MONTHS[lang];
  const days = DAYS[lang];
  
  const monthName = months[month] || month;
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "ShiftMD v3.6";
  workbook.created = new Date();
  
  const genDate = new Date();
  
  const dates = Object.keys(schedule).sort((a, b) => {
    const [da, ma, ya] = a.split("-").map(Number);
    const [db, mb, yb] = b.split("-").map(Number);
    return new Date(ya, ma - 1, da) - new Date(yb, mb - 1, db);
  });
  
  // =====================================================================
  // SHEET 1: RASPORED
  // =====================================================================
  const wsSchedule = workbook.addWorksheet(dict.sheet_raspored, {
    pageSetup: { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 1.5, right: 1.5, top: 1.5, bottom: 1.5, header: 0.5, footer: 0.5 } }
  });
  
  wsSchedule.getColumn(1).width = 80; wsSchedule.getColumn(2).width = 25; wsSchedule.getColumn(3).width = 22; wsSchedule.getColumn(4).width = 22;
  
  let r = 1;
  
  wsSchedule.mergeCells(`A${r}:D${r}`);
  wsSchedule.getCell(`A${r}`).value = dict.title;
  wsSchedule.getCell(`A${r}`).font = { bold: true, size: 20, name: "Arial", color: { argb: COLORS.text_black } };
  wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center", vertical: "middle" };
  wsSchedule.getRow(r).height = 35; r++;
  
  wsSchedule.mergeCells(`A${r}:D${r}`);
  wsSchedule.getCell(`A${r}`).value = hospital;
  wsSchedule.getCell(`A${r}`).font = { bold: true, size: 16, name: "Arial", color: { argb: "FF2563EB" } };
  wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center", vertical: "middle" };
  wsSchedule.getRow(r).height = 28; r++;
  
  wsSchedule.mergeCells(`A${r}:D${r}`);
  wsSchedule.getCell(`A${r}`).value = dict.period ? `${monthName} ${year}. ${dict.period}` : `${monthName} ${year}`;
  wsSchedule.getCell(`A${r}`).font = { bold: true, size: 14, name: "Arial", color: { argb: COLORS.text_black } };
  wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center", vertical: "middle" };
  wsSchedule.getRow(r).height = 25; r++;
  
  let infoText = `${type1Name || "Tip 1"}: ${dict.duty_count}: ${staffCount1 || 1}`;
  if (allowLess1) infoText += ` | ${dict.rule_relaxed}`; else infoText += ` | ${dict.rule_strict}`;
  if (useRanked1) infoText += ` | Rangirani: DA`;
  if (hasType2) {
    infoText += `\n${type2Name || "Tip 2"}: ${dict.duty_count}: ${staffCount2 || 1}`;
    if (allowLess2) infoText += ` | ${dict.rule_relaxed}`; else infoText += ` | ${dict.rule_strict}`;
    if (useRanked2) infoText += ` | Rangirani: DA`;
  }
  
  wsSchedule.mergeCells(`A${r}:D${r}`);
  wsSchedule.getCell(`A${r}`).value = infoText;
  wsSchedule.getCell(`A${r}`).font = { size: 10, name: "Arial", color: { argb: "FF6B7280" } };
  wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" };
  wsSchedule.getRow(r).height = hasType2 ? 35 : 20; r++;
  
  wsSchedule.mergeCells(`A${r}:D${r}`);
  wsSchedule.getCell(`A${r}`).value = `${dict.generated}: ${genDate.toLocaleDateString("sr-RS")} ${genDate.toLocaleTimeString("sr-RS")} | ${dict.footer}`;
  wsSchedule.getCell(`A${r}`).font = { size: 8, name: "Arial", color: { argb: "FF9CA3AF" } };
  wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" }; r++; r++;
  
  // STATISTIKA
  if (statistics && statistics.entries && statistics.entries.length > 0) {
    wsSchedule.mergeCells(`A${r}:D${r}`);
    wsSchedule.getCell(`A${r}`).value = dict.stats_title;
    wsSchedule.getCell(`A${r}`).font = { bold: true, size: 12, name: "Arial", color: { argb: COLORS.text_black } }; r++;
    
    const E = statistics.entries;
    const maxL = Math.max(...E.map(e => e.totalAssignedShare || 0));
    const minL = Math.min(...E.map(e => e.totalAssignedShare || 0));
    const avgL = E.reduce((s, e) => s + (e.totalAssignedShare || 0), 0) / E.length;
    const maxD = Math.max(...E.map(e => e.dutyCount || e.assignedDates?.length || 0));
    const minD = Math.min(...E.map(e => e.dutyCount || e.assignedDates?.length || 0));
    const avgD = E.reduce((s, e) => s + (e.dutyCount || e.assignedDates?.length || 0), 0) / E.length;
    const maxW = Math.max(...E.map(e => e.weekendCount || 0));
    const minW = Math.min(...E.map(e => e.weekendCount || 0));
    const avgW = E.reduce((s, e) => s + (e.weekendCount || 0), 0) / E.length;
    
    ["", dict.stats_load, dict.stats_duties, dict.stats_weekends].forEach((h, i) => {
      const cell = wsSchedule.getCell(r, i + 1);
      cell.value = h;
      cell.font = { bold: true, size: 11, name: "Arial", color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF374151" } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    }); r++;
    
    [
      [dict.max, maxL.toFixed(2), String(maxD), String(maxW), "FFFEE2E2"],
      [dict.avg, avgL.toFixed(2), avgD.toFixed(1), avgW.toFixed(1), "FFF3F4F6"],
      [dict.min, minL.toFixed(2), String(minD), String(minW), "FFDCFCE7"]
    ].forEach(row => {
      row.slice(0, 4).forEach((val, i) => {
        const cell = wsSchedule.getCell(r, i + 1);
        cell.value = val;
        cell.font = { size: 11, name: "Arial", color: { argb: COLORS.text_black } };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: row[4] } };
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      }); r++;
    });
    
    const loaded = E.filter(e => (e.totalAssignedShare || 0) > 0).length;
    wsSchedule.mergeCells(`A${r}:D${r}`);
    wsSchedule.getCell(`A${r}`).value = `${dict.total_doctors}: ${E.length} | ${dict.with_duties}: ${loaded} | ${dict.without}: ${E.length - loaded}`;
    wsSchedule.getCell(`A${r}`).font = { size: 9, name: "Arial", color: { argb: "FF6B7280" } }; wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" }; r++;
    
    wsSchedule.mergeCells(`A${r}:D${r}`);
    wsSchedule.getCell(`A${r}`).value = `${dict.period_label}: ${statistics.firstDate || "—"} ${lang === 'sr' ? 'do' : 'to'} ${statistics.lastDate || "—"} (${statistics.totalDays || 0} ${dict.days})`;
    wsSchedule.getCell(`A${r}`).font = { size: 9, name: "Arial", color: { argb: "FF9CA3AF" } }; wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" }; r++; r++;
    
    wsSchedule.mergeCells(`A${r}:D${r}`);
    wsSchedule.getCell(`A${r}`).value = "═".repeat(100);
    wsSchedule.getCell(`A${r}`).font = { size: 8, name: "Arial", color: { argb: "FFD1D5DB" } }; wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" }; r++; r++;
  }
  
  // RASPORED
  const datesByType = {};
  for (const [date, typeLabel] of Object.entries(dutyTypes || {})) {
    if (!datesByType[typeLabel]) datesByType[typeLabel] = [];
    datesByType[typeLabel].push(date);
  }
  if (Object.keys(datesByType).length === 0) datesByType[lang === 'en' ? "Duties" : "Dežurstva"] = dates;
  
  for (const [typeLabel, typeDates] of Object.entries(datesByType)) {
    if (Object.keys(datesByType).length > 1) {
      wsSchedule.mergeCells(`A${r}:D${r}`);
      wsSchedule.getCell(`A${r}`).value = `▸ ${typeLabel}`;
      wsSchedule.getCell(`A${r}`).font = { bold: true, size: 14, name: "Arial", color: { argb: "FF2563EB" } };
      wsSchedule.getCell(`A${r}`).alignment = { horizontal: "left", vertical: "middle" }; wsSchedule.getRow(r).height = 25; r++;
    }
    
    typeDates.sort((a, b) => {
      const [da, ma, ya] = a.split("-").map(Number); const [db, mb, yb] = b.split("-").map(Number);
      return new Date(ya, ma - 1, da) - new Date(yb, mb - 1, db);
    });
    
    for (const date of typeDates) {
      const [day, m, y] = date.split("-").map(Number);
      const dateObj = new Date(y, m - 1, day);
      const dayName = days[dateObj.getDay() === 0 ? 6 : dateObj.getDay() - 1];
      const isW = dateObj.getDay() === 0 || dateObj.getDay() === 6;
      
      wsSchedule.mergeCells(`A${r}:D${r}`);
      const dc = wsSchedule.getCell(`A${r}`);
      dc.value = `${date} (${dayName})`;
      dc.font = { bold: true, size: 13, name: "Arial", color: { argb: isW ? "FFDC2626" : COLORS.text_black } };
      dc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: isW ? "FFFEE2E2" : "FFF3F4F6" } };
      dc.alignment = { horizontal: "left", vertical: "middle" }; wsSchedule.getRow(r).height = 22; r++;
      
      if (schedule[date]) {
        schedule[date].forEach((slot, si) => {
          wsSchedule.mergeCells(`A${r}:D${r}`);
          wsSchedule.getCell(`A${r}`).value = `  ${dict.duty_slot} ${si + 1}:`;
          wsSchedule.getCell(`A${r}`).font = { bold: true, size: 12, name: "Arial", color: { argb: COLORS.text_black } }; r++;
          
          if (slot.persons) {
            slot.persons.forEach(person => {
              wsSchedule.mergeCells(`A${r}:D${r}`);
              const cell = wsSchedule.getCell(`A${r}`);
              const roleText = person.role === "specijalista" ? dict.specialist : dict.resident;
              const rankText = person.rank ? ` [${dict.header_rank} ${person.rank}]` : "";
              cell.value = `    ${person.isChief ? "★ " : "• "}${person.name} — ${roleText} (${person.share})${rankText}${person.isChief ? " — " + dict.chief : ""}`;
              cell.font = { bold: person.isChief, size: 11, name: "Arial", color: { argb: person.isChief ? "FF059669" : COLORS.text_black } };
              if (person.isChief) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.chief_bg } };
              r++;
            });
          }
          r++;
        });
      }
      r++;
    }
  }
  
  wsSchedule.mergeCells(`A${r}:D${r}`);
  wsSchedule.getCell(`A${r}`).value = `© ${dict.footer} | ${dict.generated}: ${genDate.toLocaleDateString("sr-RS")} ${genDate.toLocaleTimeString("sr-RS")}`;
  wsSchedule.getCell(`A${r}`).font = { size: 7, name: "Arial", color: { argb: "FFD1D5DB" } }; wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" };
  
  // =====================================================================
  // SHEET 2: TABELARNI PREGLED
  // =====================================================================
  const rankedUsed = useRanked1 || useRanked2;
  const extraCols = (hasType2 ? 1 : 0) + (rankedUsed ? 1 : 0);
  const totalCols = 7 + extraCols;
  const lastCol = String.fromCharCode(64 + totalCols);
  
  const wsTable = workbook.addWorksheet(dict.sheet_table, {
    pageSetup: { orientation: "landscape", paperSize: 9, fitToPage: false, fitToWidth: 0, fitToHeight: 0, margins: { left: 1.5, right: 1.5, top: 1.5, bottom: 1.5 } }
  });
  
  wsTable.getColumn(1).width = 18; wsTable.getColumn(2).width = 22; wsTable.getColumn(3).width = 22;
  wsTable.getColumn(4).width = 50; wsTable.getColumn(5).width = 24; wsTable.getColumn(6).width = 20; wsTable.getColumn(7).width = 24;
  let colIdx = 8;
  if (hasType2) { wsTable.getColumn(colIdx).width = 24; colIdx++; }
  if (rankedUsed) { wsTable.getColumn(colIdx).width = 12; }
  
  wsTable.mergeCells(`A1:${lastCol}1`);
  wsTable.getCell("A1").value = `${hospital} — ${dict.title}: ${monthName} ${year}`;
  wsTable.getCell("A1").font = { bold: true, size: 14, name: "Arial", color: { argb: COLORS.text_black } };
  wsTable.getCell("A1").alignment = { horizontal: "center", vertical: "middle" }; wsTable.getRow(1).height = 30;
  
  wsTable.mergeCells(`A2:${lastCol}2`);
  wsTable.getCell("A2").value = hasType2 ? `${type1Name || "Tip 1"} & ${type2Name || "Tip 2"} | ${dict.footer}` : `${dict.footer}`;
  wsTable.getCell("A2").font = { size: 9, name: "Arial", color: { argb: "FF6B7280" } }; wsTable.getCell("A2").alignment = { horizontal: "center" };
  
  const headers = [dict.header_date, dict.header_day, dict.header_slot, dict.header_name, dict.header_type, dict.header_share, dict.header_chief];
  if (hasType2) headers.push(dict.header_duty_type);
  if (rankedUsed) headers.push(dict.header_rank);
  
  const hRow = wsTable.addRow(headers);
  hRow.font = { bold: true, size: 11, name: "Arial", color: { argb: "FFFFFFFF" } };
  hRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF374151" } };
  hRow.alignment = { horizontal: "center", vertical: "middle" }; hRow.height = 25;
  
  // Podaci - SAMO PRVI RED SVAKOG DANA IMA SVETLO PLAVU POZADINU
  let prevDate = "";
  
  for (const date of dates) {
    const [day, m, y] = date.split("-").map(Number);
    const dateObj = new Date(y, m - 1, day);
    const dayName = days[dateObj.getDay() === 0 ? 6 : dateObj.getDay() - 1];
    const dutyTypeForDate = (dutyTypes && dutyTypes[date]) ? dutyTypes[date] : "";
    const isNewDate = (date !== prevDate);
    
    if (schedule[date]) {
      schedule[date].forEach((slot, slotIndex) => {
        if (slot.persons) {
          slot.persons.forEach((person, personIndex) => {
            const roleText = person.role === "specijalista" ? dict.specialist : dict.resident;
            const rowData = [
              date, dayName, `${dict.duty_slot} ${slotIndex + 1}`, person.name, roleText, person.share,
              person.isChief ? `★ ${dict.yes}` : dict.no
            ];
            if (hasType2) rowData.push(dutyTypeForDate || "-");
            if (rankedUsed) rowData.push(person.rank || "-");
            
            const row = wsTable.addRow(rowData);
            row.font = { color: { argb: COLORS.text_black } };
            
            // SAMO PRVI RED DANA (isNewDate + prvi slot + prva osoba) dobija svetlo plavu pozadinu
            if (isNewDate && slotIndex === 0 && personIndex === 0) {
              row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.first_row_bg } };
            }
            
            // Glavni dežurni uvek bold
            if (person.isChief) {
              row.font = { bold: true, color: { argb: COLORS.text_black } };
            }
            
            // Specijalisti koji nisu glavni - plavo bold u koloni Tip
            if (person.role === "specijalista" && !person.isChief) {
              row.getCell(5).font = { color: { argb: "FF2563EB" }, bold: true };
            }
            
            row.alignment = { vertical: "middle" }; row.height = 20;
          });
        }
      });
    }
    
    prevDate = date;
  }
  
  // Border
  wsTable.eachRow((row, rowNumber) => {
    if (rowNumber <= 3) return;
    row.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFD1D5DB" } }, left: { style: "thin", color: { argb: "FFD1D5DB" } },
        bottom: { style: "thin", color: { argb: "FFD1D5DB" } }, right: { style: "thin", color: { argb: "FFD1D5DB" } }
      };
    });
  });
  
  // =====================================================================
  // SHEET 3: STATISTIKA
  // =====================================================================
  if (statistics && statistics.entries && statistics.entries.length > 0) {
    const wsStats = workbook.addWorksheet(dict.sheet_stats);
    
    wsStats.getColumn(1).width = 35; wsStats.getColumn(2).width = 18; wsStats.getColumn(3).width = 20;
    wsStats.getColumn(4).width = 20; wsStats.getColumn(5).width = 20; wsStats.getColumn(6).width = 16;
    if (rankedUsed) wsStats.getColumn(7).width = 12;
    
    wsStats.mergeCells("A1:F1");
    wsStats.getCell("A1").value = `${dict.stats_title} — ${hospital}`;
    wsStats.getCell("A1").font = { bold: true, size: 14, name: "Arial", color: { argb: COLORS.text_black } }; wsStats.getCell("A1").alignment = { horizontal: "center" };
    
    wsStats.mergeCells("A2:F2");
    wsStats.getCell("A2").value = `${monthName} ${year} | ${dict.footer} | ${dict.generated}: ${genDate.toLocaleDateString("sr-RS")}`;
    wsStats.getCell("A2").font = { size: 10, name: "Arial", color: { argb: "FF6B7280" } }; wsStats.getCell("A2").alignment = { horizontal: "center" };
    
    const statHeaders = rankedUsed
      ? [dict.header_name, dict.header_type, dict.stats_load, dict.stats_duties, dict.stats_weekends, dict.header_chief, dict.header_rank]
      : [dict.header_name, dict.header_type, dict.stats_load, dict.stats_duties, dict.stats_weekends, dict.header_chief];
    
    const sRow = wsStats.addRow(statHeaders);
    sRow.font = { bold: true, size: 11, name: "Arial", color: { argb: "FFFFFFFF" } };
    sRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF374151" } }; sRow.alignment = { horizontal: "center", vertical: "middle" };
    
    const sortedEntries = [...statistics.entries].sort((a, b) => (b.totalAssignedShare || 0) - (a.totalAssignedShare || 0));
    const maxLoad = Math.max(...sortedEntries.map(e => e.totalAssignedShare || 0), 1);
    
    sortedEntries.forEach((entry) => {
      const roleText = entry.role === "specijalista" ? dict.specialist : dict.resident;
      const rowData = [entry.name, roleText, (entry.totalAssignedShare || 0).toFixed(2), entry.dutyCount || entry.assignedDates?.length || 0, entry.weekendCount || 0, entry.canBeChief ? dict.yes : dict.no];
      if (rankedUsed) rowData.push(entry.rank || "-");
      
      const row = wsStats.addRow(rowData);
      row.font = { color: { argb: COLORS.text_black } };
      if ((entry.totalAssignedShare || 0) > 0) {
        const intensity = Math.min(Math.round(((entry.totalAssignedShare || 0) / maxLoad) * 200), 200);
        const green = (255 - intensity).toString(16).padStart(2, '0');
        row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${green}FF${green}` } };
      }
      row.alignment = { vertical: "middle" };
    });
    
    wsStats.eachRow((row, rowNumber) => {
      if (rowNumber <= 2) return;
      row.eachCell((cell) => {
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      });
    });
  }
  
  await workbook.xlsx.writeFile(filePath);
  console.log(`   ✅ ${lang === 'en' ? 'Excel saved' : 'Excel sačuvan'}: ${filePath}`);
  return true;
};