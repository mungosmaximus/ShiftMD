// ShiftMD v3.6 - Kompletan scheduler sa ranked modom i pravilom jedne osobe za popunu
const { shareToNumber, shareToString, isWeekend } = require("./utils");
const { calculateScore } = require("./scoring");

module.exports = function generateSchedule(employees, dutyDates, dutyStaffCount = 1, allowLessSpecialists = false, useRankedDuties = false) {
  
  const totalDays = dutyDates.length;
  const weekendDays = dutyDates.filter(isWeekend).length;
  const maxLoadPerEmployee = Math.max(dutyStaffCount * 1.2 / Math.max(employees.length, 1), 0.25);
  const maxWeekendsPerEmployee = Math.ceil(weekendDays / Math.max(employees.length, 1)) + 1;
  const totalSlots = totalDays * dutyStaffCount;
  const avgDaysPerEmployee = totalSlots / Math.max(employees.length, 1);
  const maxDutyDays = Math.ceil(avgDaysPerEmployee) + 2;
  
  if (useRankedDuties) {
    const allRanks = [...new Set(employees.map(e => e.rank).filter(r => r > 0))].sort((a, b) => a - b);
    if (allRanks.length === 0) {
      throw new Error("Kolona 'rang' nije pronađena u Excel fajlu ili su svi rangovi 0. / Column 'rank' not found or all ranks are 0.");
    }
    if (allRanks.length !== dutyStaffCount) {
      throw new Error(`Broj dežurnih po danu (${dutyStaffCount}) ne odgovara broju različitih rangova (${allRanks.length}). Rangovi: ${allRanks.join(', ')}`);
    }
    console.log(`   📊 Ranked mod: ${allRanks.length} rangova [${allRanks.join(', ')}]`);
  }
  
  console.log("\n" + "═".repeat(80));
  console.log("🩺 ShiftMD v3.6 - RASPORED DEŽURSTAVA" + (useRankedDuties ? " (RANGIRANI TIMOVI)" : ""));
  console.log("═".repeat(80));

  employees.forEach(emp => {
    emp.assignedDates = [];
    emp.weekendCount = 0;
    emp.totalAssignedShare = 0;
    emp.lastDutyDate = null;
    emp.dutyCount = 0;
    emp.weekendDutyCount = 0;
  });

  const trackingTable = employees.map((emp, index) => ({
    id: index, employeeRef: emp, name: emp.name, role: emp.role,
    seniority: emp.seniority || 0, canBeChief: emp.canBeChief === true || emp.canBeChief === "DA",
    dutyShare: emp.dutyShare, shareNumber: shareToNumber(emp.dutyShare), rank: emp.rank || 0,
    unavailable: Array.isArray(emp.unavailable) ? emp.unavailable : [],
    preferred: Array.isArray(emp.preferred) ? emp.preferred : [],
    assignedDates: [], weekendCount: 0, totalAssignedShare: 0, lastDutyDate: null
  }));

  const schedule = {};
  const sortedDates = [...dutyDates].sort((a, b) => {
    const [da, ma, ya] = a.split("-").map(Number);
    const [db, mb, yb] = b.split("-").map(Number);
    return new Date(ya, ma - 1, da) - new Date(yb, mb - 1, db);
  });

  function canWorkOnDate(entry, date) {
    if (entry.unavailable.includes(date)) return false;
    if (entry.assignedDates.length >= maxDutyDays) return false;
    if (entry.lastDutyDate) {
      const [d, m, y] = date.split("-").map(Number);
      const [ld, lm, ly] = entry.lastDutyDate.split("-").map(Number);
      if (Math.floor((new Date(y, m - 1, d) - new Date(ly, lm - 1, ld)) / (1000 * 60 * 60 * 24)) < 2) return false;
    }
    if (entry.totalAssignedShare >= maxLoadPerEmployee * totalDays) return false;
    if (isWeekend(date) && entry.weekendCount >= maxWeekendsPerEmployee) return false;
    return true;
  }

  function assignDuty(entry, date, share) {
    entry.assignedDates.push(date); entry.totalAssignedShare += share; entry.lastDutyDate = date;
    if (isWeekend(date)) entry.weekendCount++;
    if (entry.employeeRef) {
      entry.employeeRef.assignedDates.push(date); entry.employeeRef.totalAssignedShare += share;
      entry.employeeRef.lastDutyDate = date; entry.employeeRef.dutyCount = entry.assignedDates.length;
      if (isWeekend(date)) { entry.employeeRef.weekendCount++; entry.employeeRef.weekendDutyCount = entry.weekendCount; }
    }
  }

  function getRankedCandidates(date, excludeList = [], mustBeChief = false, mustBeSpecialist = false, remainingNeeded = null, requiredRank = 0) {
    let candidates = trackingTable.filter(entry => {
      if (excludeList.includes(entry)) return false;
      if (!canWorkOnDate(entry, date)) return false;
      if (mustBeChief && !(entry.role === "specijalista" && entry.canBeChief)) return false;
      if (mustBeSpecialist && entry.role !== "specijalista") return false;
      if (requiredRank > 0 && entry.rank !== requiredRank) return false;
      return true;
    });
    candidates.sort((a, b) => calculateScore(a, date, remainingNeeded, trackingTable) - calculateScore(b, date, remainingNeeded, trackingTable));
    return candidates;
  }

  function pickBestOrRandom(candidates, date, remainingNeeded = null) {
    if (candidates.length === 0) return null;
    if (candidates.length === 1) return candidates[0];
    const byScore = {};
    candidates.forEach(c => {
      const s = calculateScore(c, date, remainingNeeded, trackingTable);
      const rs = Math.round(s * 100) / 100;
      if (!byScore[rs]) byScore[rs] = [];
      byScore[rs].push(c);
    });
    const best = byScore[Math.min(...Object.keys(byScore).map(Number))];
    return best[Math.floor(Math.random() * best.length)];
  }

  function isEmployeeInAnySlot(slots, employee) {
    return slots.some(s => s.persons && s.persons.some(p => p.employee === employee));
  }

  // ===== GLAVNA PETLJA =====
  for (const date of sortedDates) {
    const isWeekendDay = isWeekend(date);
    console.log(`\n📅 ${date} (${sortedDates.indexOf(date) + 1}/${sortedDates.length}) ${isWeekendDay ? '🔴 VIKEND' : '🔵 RADNI DAN'}`);
    console.log("─".repeat(60));
    
    schedule[date] = [];
    const assignedToday = [];
    let slots = [];
    let totalFilled = 0;
    
    if (useRankedDuties) {
      // ===== RANKED MOD =====
      const allRanks = [...new Set(trackingTable.map(e => e.rank).filter(r => r > 0))].sort((a, b) => a - b);
      console.log("   🎯 RANGED MOD: Dodela po rangovima...");
      
      for (const rank of allRanks) {
        let candidates = getRankedCandidates(date, assignedToday, rank === 1, false, null, rank)
          .filter(c => c.preferred.includes(date));
        if (candidates.length === 0) candidates = getRankedCandidates(date, assignedToday, rank === 1, false, null, rank);
        
        if (candidates.length > 0) {
          const chosen = pickBestOrRandom(candidates, date, 1.0);
          const share = chosen.shareNumber;
          
          // Kreiraj novi slot za ovaj rang
          const slot = { persons: [{ employee: chosen, share: share, isChief: rank === 1 && chosen.canBeChief }], filled: share, hasChief: rank === 1 && chosen.canBeChief };
          slots.push(slot);
          totalFilled += share;
          assignedToday.push(chosen);
          assignDuty(chosen, date, share);
          console.log(`   ✅ Rang ${rank}: ${chosen.name} (${chosen.role}, udeo: ${shareToString(share)}) ${chosen.preferred.includes(date) ? '⭐' : ''}`);
          
          // === POPUNI OSTATAK OVOG SLOTA JEDNOM OSOBOM ===
          if (share < 0.999) {
            const remaining = 1 - share;
            
            // Traži JEDNU osobu koja može da pokrije ceo ostatak
            let fillCandidates = getRankedCandidates(date, assignedToday, false, true, remaining);
            if (fillCandidates.length === 0) fillCandidates = getRankedCandidates(date, assignedToday, false, false, remaining);
            
            // Filtriraj samo one čiji udeo MOŽE da pokrije ostatak (jedna osoba)
            const singleFillers = fillCandidates.filter(c => c.shareNumber >= remaining - 0.001);
            
            if (singleFillers.length > 0) {
              // Postoji osoba koja može sama da popuni ostatak
              const filler = singleFillers[0]; // Najbolji po skoru
              const assignedShare = Math.min(filler.shareNumber, remaining);
              
              slot.persons.push({ employee: filler, share: assignedShare, isChief: false });
              slot.filled += assignedShare;
              totalFilled += assignedShare - share + assignedShare; // Korekcija
              assignedToday.push(filler);
              assignDuty(filler, date, assignedShare);
              console.log(`   ➕ Popuna rang ${rank}: ${filler.name} (${shareToString(assignedShare)}) - JEDNA OSOBA`);
            } else if (fillCandidates.length > 0) {
              // Nema idealne osobe, uzmi najbolju dostupnu (jednu)
              const filler = fillCandidates[0];
              const assignedShare = Math.min(filler.shareNumber, remaining);
              
              slot.persons.push({ employee: filler, share: assignedShare, isChief: false });
              slot.filled += assignedShare;
              totalFilled += assignedShare - share + assignedShare;
              assignedToday.push(filler);
              assignDuty(filler, date, assignedShare);
              console.log(`   ➕ Popuna rang ${rank}: ${filler.name} (${shareToString(assignedShare)}) - najbolja dostupna`);
              
              // Ako i dalje ostaje, dodaj još jednu osobu (samo ako je neophodno)
              if (slot.filled < 0.999) {
                const remaining2 = 1 - slot.filled;
                const moreFillers = getRankedCandidates(date, assignedToday, false, false, remaining2)
                  .filter(c => c.shareNumber >= remaining2 - 0.001);
                
                if (moreFillers.length > 0) {
                  const filler2 = moreFillers[0];
                  const assignedShare2 = Math.min(filler2.shareNumber, remaining2);
                  slot.persons.push({ employee: filler2, share: assignedShare2, isChief: false });
                  slot.filled += assignedShare2;
                  totalFilled += assignedShare2;
                  assignedToday.push(filler2);
                  assignDuty(filler2, date, assignedShare2);
                  console.log(`   ➕ Dodatna popuna: ${filler2.name} (${shareToString(assignedShare2)})`);
                }
              }
            }
          }
        } else {
          throw new Error(`Nedostaje lekar sa rangom ${rank} za ${date}.`);
        }
      }
    } else {
      // ===== STANDARD MOD =====
      console.log("   👑 KORAK 1: Glavni dežurni...");
      
      let chiefCandidates = getRankedCandidates(date, assignedToday, true, false, 1.0)
        .filter(c => c.preferred.includes(date));
      if (chiefCandidates.length === 0) chiefCandidates = getRankedCandidates(date, assignedToday, true, false, 1.0);
      
      if (chiefCandidates.length > 0) {
        const chief = pickBestOrRandom(chiefCandidates, date, 1.0);
        const share = chief.shareNumber;
        
        const slot = { persons: [{ employee: chief, share: share, isChief: true }], filled: share, hasChief: true };
        slots.push(slot);
        totalFilled += share;
        assignedToday.push(chief);
        assignDuty(chief, date, share);
        console.log(`   ✅ Glavni: ${chief.name} (${shareToString(share)})`);
        
        // === POPUNI OSTATAK GLAVNOG SLOTA JEDNOM OSOBOM ===
        if (share < 0.999) {
          const remaining = 1 - share;
          
          let fillCandidates = getRankedCandidates(date, assignedToday, false, true, remaining);
          if (fillCandidates.length === 0) fillCandidates = getRankedCandidates(date, assignedToday, false, false, remaining);
          
          // Filtriraj samo one koji mogu sami da popune ostatak
          const singleFillers = fillCandidates.filter(c => c.shareNumber >= remaining - 0.001);
          
          if (singleFillers.length > 0) {
            const filler = singleFillers[0];
            const assignedShare = Math.min(filler.shareNumber, remaining);
            slot.persons.push({ employee: filler, share: assignedShare, isChief: false });
            slot.filled += assignedShare;
            totalFilled += assignedShare;
            assignedToday.push(filler);
            assignDuty(filler, date, assignedShare);
            console.log(`   ➕ Popuna glavnog: ${filler.name} (${shareToString(assignedShare)}) - JEDNA OSOBA`);
          } else if (fillCandidates.length > 0) {
            const filler = fillCandidates[0];
            const assignedShare = Math.min(filler.shareNumber, remaining);
            slot.persons.push({ employee: filler, share: assignedShare, isChief: false });
            slot.filled += assignedShare;
            totalFilled += assignedShare;
            assignedToday.push(filler);
            assignDuty(filler, date, assignedShare);
            console.log(`   ➕ Popuna glavnog: ${filler.name} (${shareToString(assignedShare)}) - najbolja dostupna`);
          }
        }
      } else {
        console.warn(`   ⚠️  NEMA glavnog!`);
      }
      
      // KORAK 2: Želje
      console.log("   ⭐ KORAK 2: Želje...");
      let wished = trackingTable.filter(c => c.preferred.includes(date) && !assignedToday.includes(c) && canWorkOnDate(c, date));
      wished.sort((a, b) => b.seniority - a.seniority);
      
      for (const entry of wished) {
        if (totalFilled >= dutyStaffCount - 0.001) break;
        const share = Math.min(entry.shareNumber, dutyStaffCount - totalFilled);
        if (share <= 0) continue;
        
        // Novi slot za želju
        const slot = { persons: [{ employee: entry, share: share, isChief: false }], filled: share, hasChief: false };
        slots.push(slot);
        totalFilled += share;
        assignedToday.push(entry);
        assignDuty(entry, date, share);
        console.log(`   ⭐ Želja: ${entry.name} (${shareToString(share)})`);
        
        // Popuni ostatak JEDNOM osobom
        if (share < 0.999) {
          const remaining = 1 - share;
          let fillCandidates = getRankedCandidates(date, assignedToday, false, false, remaining);
          const singleFillers = fillCandidates.filter(c => c.shareNumber >= remaining - 0.001);
          
          if (singleFillers.length > 0) {
            const filler = singleFillers[0];
            const assignedShare = Math.min(filler.shareNumber, remaining);
            slot.persons.push({ employee: filler, share: assignedShare, isChief: false });
            slot.filled += assignedShare;
            totalFilled += assignedShare;
            assignedToday.push(filler);
            assignDuty(filler, date, assignedShare);
          } else if (fillCandidates.length > 0) {
            const filler = fillCandidates[0];
            const assignedShare = Math.min(filler.shareNumber, remaining);
            slot.persons.push({ employee: filler, share: assignedShare, isChief: false });
            slot.filled += assignedShare;
            totalFilled += assignedShare;
            assignedToday.push(filler);
            assignDuty(filler, date, assignedShare);
          }
        }
      }
    }
    
    // KORAK 3: Popuna do punog broja
    console.log("   📋 KORAK 3: Popuna do punog broja...");
    
    while (totalFilled < dutyStaffCount - 0.001) {
      const remaining = dutyStaffCount - totalFilled;
      const candidates = getRankedCandidates(date, assignedToday, false, false, remaining);
      
      if (candidates.length === 0) {
        console.warn(`   ⚠️  Nema kandidata! ${totalFilled.toFixed(2)}/${dutyStaffCount}`);
        break;
      }
      
      const chosen = candidates[0];
      const assignedShare = Math.min(chosen.shareNumber, remaining);
      if (assignedShare <= 0) break;
      
      const slot = { persons: [{ employee: chosen, share: assignedShare, isChief: false }], filled: assignedShare, hasChief: false };
      slots.push(slot);
      totalFilled += assignedShare;
      assignedToday.push(chosen);
      assignDuty(chosen, date, assignedShare);
      console.log(`   🆕 Novi slot: ${chosen.name} (${shareToString(assignedShare)})`);
      
      // Popuni ostatak JEDNOM osobom
      if (slot.filled < 0.999 && totalFilled < dutyStaffCount) {
        const slotRemaining = 1 - slot.filled;
        let fillCandidates = getRankedCandidates(date, assignedToday, false, false, slotRemaining);
        const singleFillers = fillCandidates.filter(c => c.shareNumber >= slotRemaining - 0.001);
        
        if (singleFillers.length > 0) {
          const filler = singleFillers[0];
          const fs = Math.min(filler.shareNumber, slotRemaining);
          slot.persons.push({ employee: filler, share: fs, isChief: false });
          slot.filled += fs;
          totalFilled += fs;
          assignedToday.push(filler);
          assignDuty(filler, date, fs);
          console.log(`   ➕ Popuna: ${filler.name} (${shareToString(fs)}) - JEDNA OSOBA`);
        } else if (fillCandidates.length > 0) {
          const filler = fillCandidates[0];
          const fs = Math.min(filler.shareNumber, slotRemaining);
          slot.persons.push({ employee: filler, share: fs, isChief: false });
          slot.filled += fs;
          totalFilled += fs;
          assignedToday.push(filler);
          assignDuty(filler, date, fs);
        }
      }
    }
    
    // Osiguraj glavnog (samo standard mod)
    if (!useRankedDuties) {
      for (const slot of slots) {
        if (!slot.hasChief && !slot.persons.some(p => p.employee.role === "specijalista" && p.employee.canBeChief)) {
          const chief = trackingTable.find(e => e.role === "specijalista" && e.canBeChief && !e.unavailable.includes(date) && canWorkOnDate(e, date) && !isEmployeeInAnySlot(slots, e));
          if (chief) {
            const s = chief.shareNumber;
            if (slot.filled + s <= 1.001) {
              slot.persons.push({ employee: chief, share: s, isChief: true });
              slot.filled += s; slot.hasChief = true; assignedToday.push(chief); assignDuty(chief, date, s);
            }
          }
        }
      }
    }
    
    console.log(`   📊 Popunjeno: ${totalFilled.toFixed(2)}/${dutyStaffCount}`);
    schedule[date] = slots;
  }

  // Formatiranje izlaza
  for (const date of sortedDates) {
    schedule[date] = (schedule[date] || []).map(slot => {
      const chief = slot.persons.find(p => p.isChief);
      const ordered = chief ? [chief, ...slot.persons.filter(p => p !== chief)] : slot.persons;
      return { persons: ordered.map(p => ({ name: p.employee.name, role: p.employee.role, share: shareToString(p.share), isChief: p.isChief || false, rank: p.employee.rank || 0 })) };
    });
  }
  
  // Statistika
  console.log("\n" + "═".repeat(80));
  console.log("📊 STATISTIKA" + (useRankedDuties ? " (RANGIRANI)" : ""));
  console.log("═".repeat(80));
  [...trackingTable].sort((a, b) => b.totalAssignedShare - a.totalAssignedShare).forEach((e, i) => {
    const bar = "█".repeat(Math.min(Math.round(e.totalAssignedShare * 8), 40));
    console.log(`${(i+1).toString().padStart(2)}. ${e.name.padEnd(28)} ${e.totalAssignedShare.toFixed(2).padStart(5)} ${bar}`);
  });
  console.log("═".repeat(80) + "\n");
  
  return schedule;
};