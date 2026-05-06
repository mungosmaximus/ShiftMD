const { shareToNumber, isWeekend } = require("./utils");

/**
 * ═══════════════════════════════════════════════════════════════
 * ShiftMD v2.0 - Napredni Scoring Sistem
 * ═══════════════════════════════════════════════════════════════
 * 
 * CILJ: Minimizirati "cost" (skor) = što manji broj, to bolji kandidat
 * 
 * FAKTORI:
 * 1. Opterećenje (ukupno) - NAJVAŽNIJI
 * 2. Gustina dežurstava - prosek dana između
 * 3. Vikendi - posebno kažnjeni
 * 4. Razmak od poslednjeg vikenda
 * 5. Starešinstvo - kategorijski
 * 6. Udeo dežurstva - bonus za manji udeo
 * 7. Želje - bonus ako je tražio ovaj datum
 * 8. Poređenje sa prosekom - izravnavanje opterećenja
 */

// Težinski faktori (mogu se podešavati)
const WEIGHTS = {
  TOTAL_LOAD: 50,           // Svaka jedinica opterećenja (1 celo dežurstvo = 1.0)
  DUTY_COUNT: 15,           // Svako pojedinačno dežurstvo
  DUTY_COUNT_SQUARED: 2,    // Kvadrat broja dežurstava (kazna za mnogo dežurstava)
  
  WEEKEND_COUNT: 25,        // Svaki vikend dežurstvo (povećano sa 3)
  WEEKEND_GAP_PENALTY: 30,  // Kazna ako je prošli vikend radio
  
  DENSITY_FACTOR: 10,       // Faktor gustine (manji razmak = veća kazna)
  MIN_IDEAL_GAP: 4,         // Idealni razmak između dežurstava (dani)
  
  SENIORITY_CATEGORY: 20,   // Po kategoriji starešinstva
  
  SHARE_MATCH_BONUS: 15,    // Bonus ako udeo tačno odgovara potrebnom
  
  PREFERRED_BONUS: -30,     // Veliki bonus (negativan skor) za željeni datum
  
  GAP_BONUS: 5,             // Bonus po danu razmaka
  GAP_PENALTY_THRESHOLD: 3, // Ispod ovog razmaka = kazna
  GAP_PENALTY: 50,          // Kazna po danu ispod praga
};

/**
 * Kategorije starešinstva
 * Niža kategorija = veći prioritet za dežurstva
 */
function getSeniorityCategory(seniority) {
  if (seniority <= 2) return 1;   // Početnici - najveći prioritet
  if (seniority <= 5) return 2;   // Mlađi
  if (seniority <= 10) return 3;  // Srednji
  if (seniority <= 15) return 4;  // Iskusni
  return 5;                        // Najiskusniji - najmanji prioritet
}

/**
 * Računa prosečan razmak između dežurstava (gustinu)
 * Ako je gustina velika (mali razmak) = loše
 */
function calculateDensity(assignedDates, currentDateStr) {
  if (assignedDates.length < 2) return null;
  
  // Konvertuj sve datume u Date objekte
  const allDates = [];
  
  for (const date of assignedDates) {
    if (typeof date === 'string') {
      const [d, m, y] = date.split("-").map(Number);
      allDates.push(new Date(y, m - 1, d));
    }
  }
  
  // Dodaj i trenutni datum
  const [cd, cm, cy] = currentDateStr.split("-").map(Number);
  allDates.push(new Date(cy, cm - 1, cd));
  
  // Sortiraj hronološki
  allDates.sort((a, b) => a - b);
  
  // Računaj razmake između uzastopnih datuma
  const gaps = [];
  for (let i = 1; i < allDates.length; i++) {
    const gap = Math.floor((allDates[i] - allDates[i-1]) / (1000 * 60 * 60 * 24));
    gaps.push(gap);
  }
  
  // Prosečan razmak
  const avgGap = gaps.reduce((s, g) => s + g, 0) / gaps.length;
  return avgGap;
}

/**
 * Računa dane od poslednjeg vikenda
 * Veći broj = duže nije radio vikend = bolje
 */
function daysSinceLastWeekend(assignedDates, currentDateStr) {
  const weekendDates = assignedDates.filter(d => isWeekend(d));
  if (weekendDates.length === 0) return 999; // Nije još radio vikend - veliki bonus
  
  const [cd, cm, cy] = currentDateStr.split("-").map(Number);
  const currentObj = new Date(cy, cm - 1, cd);
  
  let maxDays = 0;
  for (const wd of weekendDates) {
    const [wd_d, wd_m, wd_y] = wd.split("-").map(Number);
    const weekendObj = new Date(wd_y, wd_m - 1, wd_d);
    const days = Math.floor((currentObj - weekendObj) / (1000 * 60 * 60 * 24));
    if (days > maxDays) maxDays = days;
  }
  
  return maxDays;
}

/**
 * GLAVNA FUNKCIJA - Računa skor (cost) za kandidata
 * MANJI SKOR = BOLJI KANDIDAT
 * 
 * @param {Object} entry - Zapis iz tracking tabele
 * @param {string} date - Datum za koji se računa (format: "DD-MM-YYYY")
 * @param {number|null} remainingNeeded - Koliko još treba popuniti (za bonus udela)
 * @param {Array|null} allEntries - Svi zapisi (za poređenje sa prosekom)
 * @returns {number} Skor (manji = bolji)
 */
function calculateScore(entry, date, remainingNeeded = null, allEntries = null) {
  let score = 0;
  
  // ===== 1. UKUPNO OPTEREĆENJE (NAJVAŽNIJI FAKTOR) =====
  score += entry.totalAssignedShare * WEIGHTS.TOTAL_LOAD;
  
  // ===== 2. BROJ DEŽURSTAVA SA KVADRATNOM KAZNOM =====
  const dutyCount = entry.assignedDates.length;
  score += dutyCount * WEIGHTS.DUTY_COUNT;
  
  // Kvadratna kazna za više od 3 dežurstva (sprečava gomilanje)
  if (dutyCount > 3) {
    score += Math.pow(dutyCount - 3, 2) * WEIGHTS.DUTY_COUNT_SQUARED;
  }
  
  // ===== 3. VIKEND SPECIFIČNI FAKTORI =====
  if (isWeekend(date)) {
    // Kazna za svaki već dodeljeni vikend
    score += entry.weekendCount * WEIGHTS.WEEKEND_COUNT;
    
    // Provera razmaka od poslednjeg vikenda
    const daysSinceWeekend = daysSinceLastWeekend(entry.assignedDates, date);
    if (daysSinceWeekend < 14) {
      // Radio je vikend u poslednjih 14 dana - progresivna kazna
      const penalty = WEIGHTS.WEEKEND_GAP_PENALTY * (1 - daysSinceWeekend / 14);
      score += penalty;
    }
  }
  
  // ===== 4. GUSTINA DEŽURSTAVA =====
  if (dutyCount >= 2) {
    const density = calculateDensity(entry.assignedDates, date);
    if (density !== null && density < WEIGHTS.MIN_IDEAL_GAP) {
      // Kazna za preveliku gustinu (premali razmaci)
      score += (WEIGHTS.MIN_IDEAL_GAP - density) * WEIGHTS.DENSITY_FACTOR;
    }
  }
  
  // ===== 5. RAZMAK OD POSLEDNJEG DEŽURSTVA =====
  if (entry.lastDutyDate) {
    const [d, m, y] = date.split("-").map(Number);
    const [ld, lm, ly] = entry.lastDutyDate.split("-").map(Number);
    const dateObj = new Date(y, m - 1, d);
    const lastDateObj = new Date(ly, lm - 1, ld);
    const diffDays = Math.floor((dateObj - lastDateObj) / (1000 * 60 * 60 * 24));
    
    if (diffDays >= WEIGHTS.GAP_PENALTY_THRESHOLD) {
      // Bonus za dobar razmak
      score -= diffDays * WEIGHTS.GAP_BONUS;
    } else {
      // Kazna za premali razmak
      score += (WEIGHTS.GAP_PENALTY_THRESHOLD - diffDays) * WEIGHTS.GAP_PENALTY;
    }
  } else {
    // Još nije dežurao - veliki bonus
    score -= 50;
  }
  
  // ===== 6. STAREŠINSTVO (KATEGORIJSKI) =====
  const category = getSeniorityCategory(entry.seniority);
  score -= category * WEIGHTS.SENIORITY_CATEGORY;
  // Niža kategorija (mlađi) = više negativnih poena = manji skor = prednost
  
  // ===== 7. UDEO DEŽURSTVA - BONUS ZA POKLAPANJE =====
  if (remainingNeeded !== null && remainingNeeded > 0) {
    const share = shareToNumber(entry.dutyShare);
    const diff = Math.abs(share - remainingNeeded);
    
    if (diff < 0.01) {
      // Savršeno poklapanje
      score -= WEIGHTS.SHARE_MATCH_BONUS;
    } else if (share <= remainingNeeded + 0.01) {
      // Može da stane bez prekoračenja
      score -= WEIGHTS.SHARE_MATCH_BONUS * 0.5;
    }
    // Ako je udeo veći od potrebnog, nema bonusa
  }
  
  // ===== 8. ŽELJE - BONUS AKO JE TRAŽIO OVAJ DATUM =====
  if (entry.preferred && entry.preferred.includes(date)) {
    score += WEIGHTS.PREFERRED_BONUS; // Negativan broj = smanjuje skor
  }
  
  // ===== 9. POREDJENJE SA PROSEKOM (ako imamo sve entry-je) =====
  if (allEntries && allEntries.length > 1) {
    const avgLoad = allEntries.reduce((s, e) => s + (e.totalAssignedShare || 0), 0) / allEntries.length;
    const diffFromAvg = entry.totalAssignedShare - avgLoad;
    
    if (diffFromAvg > 0) {
      // Iznad proseka - progresivna kazna
      score += diffFromAvg * WEIGHTS.TOTAL_LOAD * 1.5;
    } else if (diffFromAvg < 0) {
      // Ispod proseka - bonus
      score += diffFromAvg * WEIGHTS.TOTAL_LOAD * 0.5;
      // diffFromAvg je negativan, pa je rezultat negativan = smanjuje skor
    }
  }
  
  return score;
}

module.exports = { calculateScore, WEIGHTS, getSeniorityCategory };