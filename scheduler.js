const { shareToNumber, shareToString, isWeekend } = require("./utils");
const { calculateScore } = require("./scoring");

/**
 * ═══════════════════════════════════════════════════════════════
 * KONAČNI SCHEDULER v5.1 - SA UNAPREĐENIM SCORINGOM I STATISTIKOM
 * ═══════════════════════════════════════════════════════════════
 * 
 * REDOSLED DODELE ZA SVAKI DATUM:
 * 
 * 1. GLAVNI DEŽURNI
 *    - Specijalista sa canBeChief=DA
 *    - Prednost: tražio taj dan + najniže starešinstvo + najmanji skor
 *    - Ako više kandidata isto - nasumičan izbor
 * 
 * 2. ŽELJE (preferred dates)
 *    - Ko je tražio taj dan
 *    - Ako više ljudi traži: sortiraj po višem starešinstvu
 *    - Dodaj dok se ne popuni broj dežurstava
 * 
 * 3. POPUNA DO PUNOG BROJA DEŽURSTAVA
 *    - Sortiraj sve dostupne po SKORU (manji = bolji)
 *    - Ako odabrani ima udeo < 1, odmah traži sledećeg da popuni ostatak
 *    - Poštuje "ne može" (unavailable)
 *    - Minimum 2 dana razmaka između dežurstava
 *    - Koristi unapređeni scoring sa svim faktorima
 * 
 * DODATNI MEHANIZMI ZAŠTITE OD BURNOUT-A:
 *   - Maksimalno 30% ukupnog opterećenja po zaposlenom
 *   - Minimum 2 dana razmaka
 *   - Maksimalno 2 vikenda mesečno po zaposlenom (ili proporcionalno)
 *   - Ravnomerna raspodela vikenda
 *   - Eksponencijalna kazna za prekoračenje proseka
 * ═══════════════════════════════════════════════════════════════
 */
module.exports = function generateSchedule(employees, dutyDates, dutyStaffCount = 1, allowLessSpecialists = false) {
  
  const totalDays = dutyDates.length;
  const weekendDays = dutyDates.filter(isWeekend).length;
  const maxLoadPerEmployee = Math.max(dutyStaffCount * 1.2 / Math.max(employees.length, 1), 0.25);
  const maxWeekendsPerEmployee = Math.ceil(weekendDays / Math.max(employees.length, 1)) + 1;
  
  console.log("\n" + "═".repeat(80));
  console.log("🩺 ShiftMD v2.0 - RASPORED DEŽURSTAVA");
  console.log("═".repeat(80));
  console.log(`   Ukupno datuma: ${totalDays} (vikend: ${weekendDays})`);
  console.log(`   Dežurstava po danu: ${dutyStaffCount}`);
  console.log(`   Ukupno lekara: ${employees.length}`);
  console.log(`   Max opterećenje po lekaru: ${(maxLoadPerEmployee * 100).toFixed(0)}%`);
  console.log(`   Max vikenda po lekaru: ${maxWeekendsPerEmployee}`);
  console.log("═".repeat(80) + "\n");

  // ===== INICIJALIZUJ STATISTIKU U EMPLOYEES OBJEKTIMA =====
  employees.forEach(emp => {
    emp.assignedDates = [];
    emp.weekendCount = 0;
    emp.totalAssignedShare = 0;
    emp.lastDutyDate = null;
    emp.dutyCount = 0;           // Broj dežurstava (dana)
    emp.weekendDutyCount = 0;    // Broj vikend dežurstava
  });

  // =====================================================================
  // KREIRANJE TRACKING TABELE
  // =====================================================================
  const trackingTable = employees.map((emp, index) => ({
    id: index,
    employeeRef: emp,            // Referenca na originalni employees objekat
    name: emp.name,
    role: emp.role,
    seniority: emp.seniority || 0,
    canBeChief: emp.canBeChief === true || emp.canBeChief === "DA",
    dutyShare: emp.dutyShare,
    shareNumber: shareToNumber(emp.dutyShare),
    unavailable: Array.isArray(emp.unavailable) ? emp.unavailable : [],
    preferred: Array.isArray(emp.preferred) ? emp.preferred : [],
    
    // TRACKING POLJA
    assignedDates: [],
    weekendCount: 0,
    totalAssignedShare: 0,
    lastDutyDate: null
  }));

  const schedule = {};
  const sortedDates = [...dutyDates].sort((a, b) => {
    const [da, ma, ya] = a.split("-").map(Number);
    const [db, mb, yb] = b.split("-").map(Number);
    return new Date(ya, ma - 1, da) - new Date(yb, mb - 1, db);
  });

  // =====================================================================
  // POMOĆNE FUNKCIJE
  // =====================================================================

  function canWorkOnDate(entry, date) {
    if (entry.unavailable.includes(date)) return false;
    
    if (entry.lastDutyDate) {
      const [d, m, y] = date.split("-").map(Number);
      const [ld, lm, ly] = entry.lastDutyDate.split("-").map(Number);
      const dateObj = new Date(y, m - 1, d);
      const lastDateObj = new Date(ly, lm - 1, ld);
      const diffDays = Math.floor((dateObj - lastDateObj) / (1000 * 60 * 60 * 24));
      
      if (diffDays < 2) return false;
    }
    
    if (entry.totalAssignedShare >= maxLoadPerEmployee * totalDays) return false;
    if (isWeekend(date) && entry.weekendCount >= maxWeekendsPerEmployee) return false;
    
    return true;
  }

  /**
   * Dodeljuje dežurstvo i ažurira OBA objekta (tracking + originalni employee)
   */
  function assignDuty(entry, date, share) {
    // Ažuriraj tracking tabelu
    entry.assignedDates.push(date);
    entry.totalAssignedShare += share;
    entry.lastDutyDate = date;
    if (isWeekend(date)) {
      entry.weekendCount++;
    }
    
    // Ažuriraj originalni employees objekat (ZA STATISTIKU)
    if (entry.employeeRef) {
      entry.employeeRef.assignedDates.push(date);
      entry.employeeRef.totalAssignedShare += share;
      entry.employeeRef.lastDutyDate = date;
      entry.employeeRef.dutyCount = entry.assignedDates.length;
      if (isWeekend(date)) {
        entry.employeeRef.weekendCount++;
        entry.employeeRef.weekendDutyCount = entry.weekendCount;
      }
    }
  }

  function removeDuty(entry, date, share) {
    const idx = entry.assignedDates.lastIndexOf(date);
    if (idx > -1) entry.assignedDates.splice(idx, 1);
    entry.totalAssignedShare = Math.max(0, entry.totalAssignedShare - share);
    if (isWeekend(date)) {
      entry.weekendCount = Math.max(0, entry.weekendCount - 1);
    }
    
    // Ažuriraj originalni employees objekat
    if (entry.employeeRef) {
      const eIdx = entry.employeeRef.assignedDates.lastIndexOf(date);
      if (eIdx > -1) entry.employeeRef.assignedDates.splice(eIdx, 1);
      entry.employeeRef.totalAssignedShare = entry.totalAssignedShare;
      entry.employeeRef.dutyCount = entry.assignedDates.length;
      entry.employeeRef.weekendCount = entry.weekendCount;
      entry.employeeRef.weekendDutyCount = entry.weekendCount;
      entry.employeeRef.lastDutyDate = entry.lastDutyDate;
    }
  }

  function getRankedCandidates(date, excludeList = [], mustBeChief = false, mustBeSpecialist = false, remainingNeeded = null) {
    let candidates = trackingTable.filter(entry => {
      if (excludeList.includes(entry)) return false;
      if (!canWorkOnDate(entry, date)) return false;
      if (mustBeChief && !(entry.role === "specijalista" && entry.canBeChief)) return false;
      if (mustBeSpecialist && entry.role !== "specijalista") return false;
      return true;
    });
    
    candidates.sort((a, b) => 
      calculateScore(a, date, remainingNeeded, trackingTable) - 
      calculateScore(b, date, remainingNeeded, trackingTable)
    );
    
    return candidates;
  }

  function pickBestOrRandom(candidates, date, remainingNeeded = null) {
    if (candidates.length === 0) return null;
    if (candidates.length === 1) return candidates[0];
    
    const byScore = {};
    candidates.forEach(c => {
      const s = calculateScore(c, date, remainingNeeded, trackingTable);
      const roundedScore = Math.round(s * 100) / 100;
      if (!byScore[roundedScore]) byScore[roundedScore] = [];
      byScore[roundedScore].push(c);
    });
    
    const bestScore = Math.min(...Object.keys(byScore).map(Number));
    const bestCandidates = byScore[bestScore];
    
    return bestCandidates[Math.floor(Math.random() * bestCandidates.length)];
  }

  // =====================================================================
  // GLAVNA PETLJA
  // =====================================================================
  for (const date of sortedDates) {
    const isWeekendDay = isWeekend(date);
    const dateNumber = sortedDates.indexOf(date) + 1;
    
    console.log(`\n📅 ${date} (${dateNumber}/${sortedDates.length}) ${isWeekendDay ? '🔴 VIKEND' : '🔵 RADNI DAN'}`);
    console.log("─".repeat(60));
    
    schedule[date] = [];
    const assignedToday = [];
    
    // KORAK 1: GLAVNI DEŽURNI
    console.log("   👑 KORAK 1: Dodela glavnog dežurnog...");
    
    let chiefCandidates = getRankedCandidates(date, assignedToday, true, false, 1.0)
      .filter(c => c.preferred.includes(date));
    
    if (chiefCandidates.length === 0) {
      chiefCandidates = getRankedCandidates(date, assignedToday, true, false, 1.0);
    }
    
    if (chiefCandidates.length > 0) {
      const chosenChief = pickBestOrRandom(chiefCandidates, date, 1.0);
      const share = chosenChief.shareNumber;
      
      const slot = {
        persons: [{
          entry: chosenChief,
          share: share,
          isChief: true
        }],
        filled: share
      };
      
      schedule[date].push(slot);
      assignedToday.push(chosenChief);
      assignDuty(chosenChief, date, share);
      
      const chiefScore = calculateScore(chosenChief, date, 1.0, trackingTable);
      console.log(`   ✅ Glavni: ${chosenChief.name} (${chosenChief.role}, udeo: ${shareToString(share)}, skor: ${chiefScore.toFixed(1)}) ${chosenChief.preferred.includes(date) ? '⭐ ŽELJA' : ''}`);
    } else {
      console.warn(`   ⚠️  NEMA dostupnog glavnog dežurnog za ${date}!`);
    }
    
    // KORAK 1b: POPUNI OSTATAK GLAVNOG SLOTA
    for (const slot of schedule[date]) {
      while (slot.filled < 0.999) {
        const remaining = 1 - slot.filled;
        
        let fillCandidates = getRankedCandidates(date, assignedToday, false, true, remaining);
        if (fillCandidates.length === 0) {
          fillCandidates = getRankedCandidates(date, assignedToday, false, false, remaining);
        }
        
        if (fillCandidates.length === 0) break;
        
        const chosen = fillCandidates[0];
        const assignedShare = Math.min(chosen.shareNumber, remaining);
        
        if (assignedShare <= 0) break;
        
        slot.persons.push({
          entry: chosen,
          share: assignedShare,
          isChief: false
        });
        slot.filled += assignedShare;
        assignedToday.push(chosen);
        assignDuty(chosen, date, assignedShare);
        
        const fillScore = calculateScore(chosen, date, remaining, trackingTable);
        console.log(`   ➕ Dopuna glavnog: ${chosen.name} (${shareToString(assignedShare)}, skor: ${fillScore.toFixed(1)})`);
      }
    }
    
    // KORAK 2: ŽELJE
    console.log("   ⭐ KORAK 2: Dodela prema željama...");
    
    let wishedEntries = trackingTable.filter(c => 
      c.preferred.includes(date) && 
      !assignedToday.includes(c) &&
      canWorkOnDate(c, date)
    );
    
    wishedEntries.sort((a, b) => b.seniority - a.seniority);
    
    let wishedCount = 0;
    for (const entry of wishedEntries) {
      const currentTotal = schedule[date].reduce((sum, s) => sum + s.filled, 0);
      if (currentTotal >= dutyStaffCount - 0.001) break;
      
      const share = entry.shareNumber;
      const remaining = dutyStaffCount - currentTotal;
      const assignedShare = Math.min(share, remaining);
      
      if (assignedShare <= 0) continue;
      
      let targetSlot = schedule[date].find(s => s.filled + assignedShare <= 1.001 && s.filled < 1);
      
      if (!targetSlot) {
        targetSlot = { persons: [], filled: 0 };
        schedule[date].push(targetSlot);
      }
      
      targetSlot.persons.push({
        entry: entry,
        share: assignedShare,
        isChief: false
      });
      targetSlot.filled += assignedShare;
      assignedToday.push(entry);
      assignDuty(entry, date, assignedShare);
      
      console.log(`   ⭐ Želja: ${entry.name} (${shareToString(assignedShare)}, starešinstvo: ${entry.seniority})`);
      wishedCount++;
    }
    
    if (wishedCount === 0) {
      console.log(`   ℹ️  Nema (više) želja za ovaj datum`);
    }
    
    // KORAK 3: POPUNA DO PUNOG BROJA
    console.log("   📋 KORAK 3: Popuna do punog broja...");
    
    let currentTotal = schedule[date].reduce((sum, s) => sum + s.filled, 0);
    
    for (const slot of schedule[date]) {
      while (slot.filled < 0.999 && currentTotal < dutyStaffCount) {
        const remaining = 1 - slot.filled;
        const candidates = getRankedCandidates(date, assignedToday, false, false, remaining);
        
        if (candidates.length === 0) break;
        
        const chosen = candidates[0];
        const assignedShare = Math.min(chosen.shareNumber, remaining);
        
        if (assignedShare <= 0) break;
        
        slot.persons.push({
          entry: chosen,
          share: assignedShare,
          isChief: false
        });
        slot.filled += assignedShare;
        currentTotal += assignedShare;
        assignedToday.push(chosen);
        assignDuty(chosen, date, assignedShare);
        
        const popScore = calculateScore(chosen, date, remaining, trackingTable);
        console.log(`   ➕ Popuna: ${chosen.name} (${shareToString(assignedShare)}, skor: ${popScore.toFixed(1)})`);
      }
    }
    
    while (currentTotal < dutyStaffCount - 0.001) {
      const remaining = dutyStaffCount - currentTotal;
      const candidates = getRankedCandidates(date, assignedToday, false, false, remaining);
      
      if (candidates.length === 0) {
        console.warn(`   ⚠️  Nema više dostupnih kandidata! Popunjeno: ${currentTotal.toFixed(2)}/${dutyStaffCount}`);
        break;
      }
      
      const chosen = candidates[0];
      const assignedShare = Math.min(chosen.shareNumber, remaining);
      
      if (assignedShare <= 0) break;
      
      const newSlot = {
        persons: [{
          entry: chosen,
          share: assignedShare,
          isChief: false
        }],
        filled: assignedShare
      };
      
      schedule[date].push(newSlot);
      currentTotal += assignedShare;
      assignedToday.push(chosen);
      assignDuty(chosen, date, assignedShare);
      
      const newScore = calculateScore(chosen, date, remaining, trackingTable);
      console.log(`   🆕 Novi slot: ${chosen.name} (${shareToString(assignedShare)}, skor: ${newScore.toFixed(1)})`);
    }
    
    // KORAK 4: VALIDACIJA
    if (!allowLessSpecialists) {
      for (let i = 0; i < schedule[date].length; i++) {
        const slot = schedule[date][i];
        const specialistShare = slot.persons
          .filter(p => p.entry.role === "specijalista")
          .reduce((sum, p) => sum + p.share, 0);
        
        if (specialistShare < 0.5) {
          console.warn(`   ⚠️  Slot ${i + 1}: Samo ${(specialistShare * 100).toFixed(0)}% specijalista`);
        }
      }
    }
    
    console.log(`   📊 Popunjeno: ${currentTotal.toFixed(2)}/${dutyStaffCount} dežurstava`);
  }

  // =====================================================================
  // FORMATIRANJE IZLAZA
  // =====================================================================
  for (const date of sortedDates) {
    const slots = schedule[date] || [];
    
    schedule[date] = slots.map(slot => {
      const sortedPersons = [...slot.persons].sort((a, b) => {
        if (a.isChief) return -1;
        if (b.isChief) return 1;
        if (a.entry.role === "specijalista" && b.entry.role !== "specijalista") return -1;
        if (b.entry.role === "specijalista" && a.entry.role !== "specijalista") return 1;
        return 0;
      });
      
      return {
        persons: sortedPersons.map(p => ({
          name: p.entry.name,
          role: p.entry.role,
          share: shareToString(p.share),
          isChief: p.isChief || false
        }))
      };
    });
  }
  
  // =====================================================================
  // STATISTIKA
  // =====================================================================
  console.log("\n" + "═".repeat(80));
  console.log("📊 KONAČNA STATISTIKA OPTEREĆENJA");
  console.log("═".repeat(80));
  
  const sorted = [...trackingTable].sort((a, b) => b.totalAssignedShare - a.totalAssignedShare);
  
  console.log(
    " #  " + 
    "Ime".padEnd(25) + 
    "Opterećenje".padStart(12) + 
    "Dežurstava".padStart(12) + 
    "Vikenda".padStart(8) + 
    "Tip".padStart(16) + 
    "Glavni".padStart(8)
  );
  console.log("─".repeat(85));
  
  sorted.forEach((e, i) => {
    const barLength = Math.min(Math.round(e.totalAssignedShare * 15), 40);
    const bar = "█".repeat(barLength);
    console.log(
      `${(i + 1).toString().padStart(2)}. ` +
      `${e.name.padEnd(23)}` +
      `${e.totalAssignedShare.toFixed(2).padStart(8)} ` +
      `${bar.padEnd(15)}` +
      `${e.assignedDates.length.toString().padStart(4)} dana` +
      `${e.weekendCount.toString().padStart(6)}` +
      `${e.role.padStart(16)}` +
      `${e.canBeChief ? 'DA' : 'NE'.padStart(8)}`
    );
  });
  
  const totalLoad = sorted.reduce((s, e) => s + e.totalAssignedShare, 0);
  const avgLoad = totalLoad / Math.max(sorted.length, 1);
  const maxLoad = sorted.length > 0 ? Math.max(...sorted.map(e => e.totalAssignedShare)) : 0;
  const minLoad = sorted.length > 0 ? Math.min(...sorted.map(e => e.totalAssignedShare)) : 0;
  const loadedCount = sorted.filter(e => e.totalAssignedShare > 0).length;
  
  console.log("─".repeat(85));
  console.log(`   Ukupno opterećenje: ${totalLoad.toFixed(2)} | Prosek: ${avgLoad.toFixed(2)}`);
  console.log(`   Max: ${maxLoad.toFixed(2)} | Min: ${minLoad.toFixed(2)} | Raspon: ${(maxLoad - minLoad).toFixed(2)}`);
  console.log(`   Lekara sa dežurstvima: ${loadedCount}/${sorted.length} | Bez dežurstava: ${sorted.length - loadedCount}`);
  console.log("═".repeat(80) + "\n");
  
  return schedule;
};