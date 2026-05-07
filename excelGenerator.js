// ShiftMD v3.0 - Dvojezični Excel Generator
const ExcelJS = require("exceljs");

const MONTHS = {
  sr: { "1": "Januar", "2": "Februar", "3": "Mart", "4": "April", "5": "Maj", "6": "Jun", "7": "Jul", "8": "Avgust", "9": "Septembar", "10": "Oktobar", "11": "Novembar", "12": "Decembar" },
  en: { "1": "January", "2": "February", "3": "March", "4": "April", "5": "May", "6": "June", "7": "July", "8": "August", "9": "September", "10": "October", "11": "November", "12": "December" }
};

const DAYS = {
  sr: ["Ponedeljak", "Utorak", "Sreda", "Četvrtak", "Petak", "Subota", "Nedelja"],
  en: ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
};

// Prevodne tabele za Excel output
const T = {
  sr: {
    sheet_raspored: "Raspored",
    sheet_table: "Tabelarni pregled",
    sheet_stats: "Statistika",
    title: "RASPORED DEŽURSTAVA",
    period: "godine",
    duty_count: "Broj dežurstava po danu",
    rule_strict: "Minimum 50% specijalista u svakom dežurstvu",
    rule_relaxed: "Obavezan glavni specijalista | Dozvoljeno manje od 50% specijalista",
    generated: "Generisano",
    stats_title: "📊 STATISTIKA OPTEREĆENJA",
    stats_load: "Opterećenje (skor)",
    stats_duties: "Broj dežurstava",
    stats_weekends: "Vikend dežurstva",
    max: "Maksimum",
    avg: "Prosek",
    min: "Minimum",
    total_doctors: "Ukupno lekara",
    with_duties: "Sa dežurstvima",
    without: "Bez",
    period_label: "Period",
    days: "dana",
    duty_slot: "Dežurstvo",
    chief: "GLAVNI DEŽURNI",
    specialist: "Specijalista",
    resident: "Specijalizant",
    yes: "DA",
    no: "NE",
    header_date: "Datum",
    header_day: "Dan",
    header_slot: "Dežurstvo",
    header_name: "Ime i prezime",
    header_type: "Tip",
    header_share: "Udeo",
    header_chief: "Glavni dežurni",
    footer: "ShiftMD v3.0",
    share_full: "celo",
    share_half: "polovina",
    share_third: "trećina",
    share_quarter: "četvrtina",
    share_eighth: "osmina",
    share_three_quarters: "tri četvrtine",
    share_two_thirds: "dve trećine",
    share_half_eighth: "polovina i osmina",
    share_three_eighths: "tri osmine",
    share_five_eighths: "pet osmina",
    share_five_sixths: "pet šestina",
    share_minimal: "minimalno",
    share_whole_minus_eighth: "celo bez osmine",
  },
  en: {
    sheet_raspored: "Schedule",
    sheet_table: "Table View",
    sheet_stats: "Statistics",
    title: "DUTY SCHEDULE",
    period: "",
    duty_count: "Duties per day",
    rule_strict: "Minimum 50% specialists in each duty",
    rule_relaxed: "Mandatory chief specialist | Less than 50% specialists allowed",
    generated: "Generated",
    stats_title: "📊 WORKLOAD STATISTICS",
    stats_load: "Workload (score)",
    stats_duties: "Number of Duties",
    stats_weekends: "Weekend Duties",
    max: "Maximum",
    avg: "Average",
    min: "Minimum",
    total_doctors: "Total Physicians",
    with_duties: "With Duties",
    without: "Without",
    period_label: "Period",
    days: "days",
    duty_slot: "Duty",
    chief: "CHIEF ON CALL",
    specialist: "Specialist",
    resident: "Resident",
    yes: "YES",
    no: "NO",
    header_date: "Date",
    header_day: "Day",
    header_slot: "Duty",
    header_name: "Full Name",
    header_type: "Type",
    header_share: "Share",
    header_chief: "Chief on Call",
    footer: "ShiftMD v3.0",
    share_full: "full",
    share_half: "half",
    share_third: "third",
    share_quarter: "quarter",
    share_eighth: "eighth",
    share_three_quarters: "three quarters",
    share_two_thirds: "two thirds",
    share_half_eighth: "half and eighth",
    share_three_eighths: "three eighths",
    share_five_eighths: "five eighths",
    share_five_sixths: "five sixths",
    share_minimal: "minimal",
    share_whole_minus_eighth: "whole minus eighth",
  }
};

function shareToText(share, lang) {
  const rounded = Math.round(share * 1000) / 1000;
  const dict = T[lang] || T['sr'];
  
  if (rounded >= 0.999 && rounded <= 1.001) return dict.share_full;
  if (rounded >= 0.499 && rounded <= 0.501) return dict.share_half;
  if (rounded >= 0.332 && rounded <= 0.335) return dict.share_third;
  if (rounded >= 0.249 && rounded <= 0.251) return dict.share_quarter;
  if (rounded >= 0.124 && rounded <= 0.126) return dict.share_eighth;
  if (rounded >= 0.874 && rounded <= 0.876) return dict.share_whole_minus_eighth;
  if (rounded >= 0.749 && rounded <= 0.751) return dict.share_three_quarters;
  if (rounded >= 0.666 && rounded <= 0.668) return dict.share_two_thirds;
  if (rounded >= 0.541 && rounded <= 0.543) return dict.share_half_eighth;
  if (rounded >= 0.374 && rounded <= 0.376) return dict.share_three_eighths;
  if (rounded >= 0.624 && rounded <= 0.626) return dict.share_five_eighths;
  if (rounded >= 0.832 && rounded <= 0.834) return dict.share_five_sixths;
  if (rounded <= 0.05) return dict.share_minimal;
  
  return String(rounded);
}

module.exports = async function generateExcel(data, filePath, statistics = null) {
  const { hospital, month, year, schedule, dutyStaffCount, allowLessSpecialists } = data;
  
  // Pokušaj da detektuješ jezik iz imena institucije ili koristi srpski kao default
  const lang = detectLanguage(hospital);
  const dict = T[lang];
  const months = MONTHS[lang];
  const days = DAYS[lang];
  
  const monthName = months[month] || month;
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "ShiftMD v3.0";
  workbook.created = new Date();
  
  const specialistRule = allowLessSpecialists ? dict.rule_relaxed : dict.rule_strict;
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
  const titleCell = wsSchedule.getCell(`A${r}`);
  titleCell.value = dict.title;
  titleCell.font = { bold: true, size: 20, name: "Arial", color: { argb: "FF1E3A5F" } };
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  wsSchedule.getRow(r).height = 35; r++;
  
  wsSchedule.mergeCells(`A${r}:D${r}`);
  wsSchedule.getCell(`A${r}`).value = hospital;
  wsSchedule.getCell(`A${r}`).font = { bold: true, size: 16, name: "Arial", color: { argb: "FF2563EB" } };
  wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center", vertical: "middle" };
  wsSchedule.getRow(r).height = 28; r++;
  
  wsSchedule.mergeCells(`A${r}:D${r}`);
  const periodText = dict.period ? `${monthName} ${year}. ${dict.period}` : `${monthName} ${year}`;
  wsSchedule.getCell(`A${r}`).value = periodText;
  wsSchedule.getCell(`A${r}`).font = { bold: true, size: 14, name: "Arial", color: { argb: "FF1F2937" } };
  wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center", vertical: "middle" };
  wsSchedule.getRow(r).height = 25; r++;
  
  wsSchedule.mergeCells(`A${r}:D${r}`);
  wsSchedule.getCell(`A${r}`).value = `${dict.duty_count}: ${dutyStaffCount} | ${specialistRule}`;
  wsSchedule.getCell(`A${r}`).font = { size: 10, name: "Arial", color: { argb: "FF6B7280" } };
  wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" };
  r++;
  
  wsSchedule.mergeCells(`A${r}:D${r}`);
  wsSchedule.getCell(`A${r}`).value = `${dict.generated}: ${genDate.toLocaleDateString("sr-RS")} ${genDate.toLocaleTimeString("sr-RS")} | ${dict.footer}`;
  wsSchedule.getCell(`A${r}`).font = { size: 8, name: "Arial", color: { argb: "FF9CA3AF" } };
  wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" };
  r++; r++;
  
  // STATISTIKA
  if (statistics && statistics.entries && statistics.entries.length > 0) {
    wsSchedule.mergeCells(`A${r}:D${r}`);
    wsSchedule.getCell(`A${r}`).value = dict.stats_title;
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
    
    ["", dict.stats_load, dict.stats_duties, dict.stats_weekends].forEach((h, i) => {
      const cell = wsSchedule.getCell(r, i + 1);
      cell.value = h;
      cell.font = { bold: true, size: 11, name: "Arial", color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF374151" } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    });
    r++;
    
    [
      [dict.max, maxL.toFixed(2), String(maxD), String(maxW), "FFFEE2E2"],
      [dict.avg, avgL.toFixed(2), avgD.toFixed(1), avgW.toFixed(1), "FFF3F4F6"],
      [dict.min, minL.toFixed(2), String(minD), String(minW), "FFDCFCE7"]
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
    wsSchedule.getCell(`A${r}`).value = `${dict.total_doctors}: ${E.length} | ${dict.with_duties}: ${loaded} | ${dict.without}: ${E.length - loaded}`;
    wsSchedule.getCell(`A${r}`).font = { size: 9, name: "Arial", color: { argb: "FF6B7280" } };
    wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" };
    r++;
    
    wsSchedule.mergeCells(`A${r}:D${r}`);
    wsSchedule.getCell(`A${r}`).value = `${dict.period_label}: ${statistics.firstDate || "—"} ${lang === 'sr' ? 'do' : 'to'} ${statistics.lastDate || "—"} (${statistics.totalDays || 0} ${dict.days})`;
    wsSchedule.getCell(`A${r}`).font = { size: 9, name: "Arial", color: { argb: "FF9CA3AF" } };
    wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" };
    r++; r++;
    
    wsSchedule.mergeCells(`A${r}:D${r}`);
    wsSchedule.getCell(`A${r}`).value = "═".repeat(100);
    wsSchedule.getCell(`A${r}`).font = { size: 8, name: "Arial", color: { argb: "FFD1D5DB" } };
    wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" };
    r++; r++;
  }
  
  // RASPORED
  for (const date of dates) {
    const [day, m, y] = date.split("-").map(Number);
    const dateObj = new Date(y, m - 1, day);
    const dayName = days[dateObj.getDay() === 0 ? 6 : dateObj.getDay() - 1];
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
        wsSchedule.getCell(`A${r}`).value = `  ${dict.duty_slot} ${si + 1}:`;
        wsSchedule.getCell(`A${r}`).font = { bold: true, size: 12, name: "Arial", color: { argb: "FF374151" } };
        r++;
        
        if (slot.persons) {
          slot.persons.forEach(person => {
            wsSchedule.mergeCells(`A${r}:D${r}`);
            const cell = wsSchedule.getCell(`A${r}`);
            const roleText = person.role === "specijalista" ? dict.specialist : dict.resident;
            const shareText = person.share; // Već je tekstualni format iz schedulera
            cell.value = `    ${person.isChief ? "★ " : "• "}${person.name} — ${roleText} (${shareText})${person.isChief ? " — " + dict.chief : ""}`;
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
  wsSchedule.getCell(`A${r}`).value = `© ${dict.footer} | ${dict.generated}: ${genDate.toLocaleDateString("sr-RS")} ${genDate.toLocaleTimeString("sr-RS")}`;
  wsSchedule.getCell(`A${r}`).font = { size: 7, name: "Arial", color: { argb: "FFD1D5DB" } };
  wsSchedule.getCell(`A${r}`).alignment = { horizontal: "center" };
  
  // =====================================================================
  // SHEET 2: TABELARNI PREGLED
  // =====================================================================
  const wsTable = workbook.addWorksheet(dict.sheet_table, {
    pageSetup: {
      orientation: "landscape", paperSize: 9,
      fitToPage: false, fitToWidth: 0, fitToHeight: 0,
      margins: { left: 1.5, right: 1.5, top: 1.5, bottom: 1.5 }
    }
  });
  
  wsTable.getColumn(1).width = 18;
  wsTable.getColumn(2).width = 22;
  wsTable.getColumn(3).width = 22;
  wsTable.getColumn(4).width = 50;
  wsTable.getColumn(5).width = 24;
  wsTable.getColumn(6).width = 20;
  wsTable.getColumn(7).width = 24;
  
  wsTable.mergeCells("A1:G1");
  wsTable.getCell("A1").value = `${hospital} — ${dict.title}: ${monthName} ${year}`;
  wsTable.getCell("A1").font = { bold: true, size: 14, name: "Arial", color: { argb: "FF1E3A5F" } };
  wsTable.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  wsTable.getRow(1).height = 30;
  
  wsTable.mergeCells("A2:G2");
  wsTable.getCell("A2").value = `${dict.duty_count}: ${dutyStaffCount} | ${specialistRule} | ${dict.footer}`;
  wsTable.getCell("A2").font = { size: 9, name: "Arial", color: { argb: "FF6B7280" } };
  wsTable.getCell("A2").alignment = { horizontal: "center" };
  
  const hRow = wsTable.addRow([dict.header_date, dict.header_day, dict.header_slot, dict.header_name, dict.header_type, dict.header_share, dict.header_chief]);
  hRow.font = { bold: true, size: 11, name: "Arial", color: { argb: "FFFFFFFF" } };
  hRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF374151" } };
  hRow.alignment = { horizontal: "center", vertical: "middle" };
  hRow.height = 25;
  
  for (const date of dates) {
    const [day, m, y] = date.split("-").map(Number);
    const dateObj = new Date(y, m - 1, day);
    const dayName = days[dateObj.getDay() === 0 ? 6 : dateObj.getDay() - 1];
    const isWeekendDay = dateObj.getDay() === 0 || dateObj.getDay() === 6;
    
    if (schedule[date]) {
      schedule[date].forEach((slot, slotIndex) => {
        if (slot.persons) {
          slot.persons.forEach((person) => {
            const roleText = person.role === "specijalista" ? dict.specialist : dict.resident;
            const row = wsTable.addRow([
              date,
              dayName,
              `${dict.duty_slot} ${slotIndex + 1}`,
              person.name,
              roleText,
              person.share,
              person.isChief ? `★ ${dict.yes}` : dict.no
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
    const wsStats = workbook.addWorksheet(dict.sheet_stats);
    
    wsStats.getColumn(1).width = 35;
    wsStats.getColumn(2).width = 18;
    wsStats.getColumn(3).width = 16;
    wsStats.getColumn(4).width = 20;
    wsStats.getColumn(5).width = 20;
    wsStats.getColumn(6).width = 16;
    
    wsStats.mergeCells("A1:F1");
    wsStats.getCell("A1").value = `${dict.stats_title} — ${hospital}`;
    wsStats.getCell("A1").font = { bold: true, size: 14, name: "Arial" };
    wsStats.getCell("A1").alignment = { horizontal: "center" };
    
    wsStats.mergeCells("A2:F2");
    wsStats.getCell("A2").value = `${monthName} ${year} | ${dict.footer} | ${dict.generated}: ${genDate.toLocaleDateString("sr-RS")}`;
    wsStats.getCell("A2").font = { size: 10, name: "Arial", color: { argb: "FF6B7280" } };
    wsStats.getCell("A2").alignment = { horizontal: "center" };
    
    const sRow = wsStats.addRow([dict.header_name, dict.header_type, dict.stats_load, dict.stats_duties, dict.stats_weekends, dict.header_chief.replace(dict.chief, dict.yes + "/" + dict.no)]);
    sRow.font = { bold: true, size: 11, name: "Arial", color: { argb: "FFFFFFFF" } };
    sRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF374151" } };
    sRow.alignment = { horizontal: "center", vertical: "middle" };
    
    [...statistics.entries].sort((a, b) => b.totalAssignedShare - a.totalAssignedShare).forEach((entry) => {
      const roleText = entry.role === "specijalista" ? dict.specialist : dict.resident;
      const row = wsStats.addRow([
        entry.name,
        roleText,
        entry.totalAssignedShare.toFixed(2),
        entry.dutyCount || entry.assignedDates.length,
        entry.weekendCount,
        entry.canBeChief ? dict.yes : dict.no
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
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      });
    });
  }
  
  await workbook.xlsx.writeFile(filePath);
  console.log(`   ✅ ${lang === 'en' ? 'Excel saved' : 'Excel sačuvan'}: ${filePath}`);
  return true;
};

// Detekcija jezika na osnovu naziva institucije
function detectLanguage(hospital) {
  if (!hospital) return 'sr';
  
  const lower = hospital.toLowerCase();
  
  // Engleske reči koje indiciraju engleski jezik
  const englishIndicators = ['hospital', 'clinic', 'medical', 'center', 'centre', 'health', 'general', 'university', 'institute'];
  
  // Srpske reči koje indiciraju srpski jezik
  const serbianIndicators = ['bolnica', 'dom', 'zdravlja', 'klinički', 'centar', 'zavod', 'opšta', 'univerzitetski', 'institut'];
  
  let enScore = 0;
  let srScore = 0;
  
  englishIndicators.forEach(word => {
    if (lower.includes(word)) enScore++;
  });
  
  serbianIndicators.forEach(word => {
    if (lower.includes(word)) srScore++;
  });
  
  // Ako ima više engleskih indikatora, koristi engleski
  if (enScore > srScore) return 'en';
  
  // Default: srpski
  return 'sr';
}