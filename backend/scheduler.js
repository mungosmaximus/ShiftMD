const score = require("./scoring");
const { shareToNumber, isWeekend } = require("./utils");

module.exports = function generateSchedule(employees, dutyDates) {
  const schedule = {};

  for (const date of dutyDates) {
    let assigned = [];
    let total = 0;

    // 1. ŽELJE
    for (const e of employees) {
      if (
        e.preferred.includes(date) &&
        !e.unavailable.includes(date) &&
        total < 1
      ) {
        const s = shareToNumber(e.dutyShare);
        if (total + s <= 1) {
          assigned.push({ e, share: s });
          total += s;
          e.assignedDates.push(date);
          if (isWeekend(date)) e.weekendCount++;
        }
      }
    }

    // 2. POPUNA
    while (total < 1) {
      const candidates = employees
        .filter(e =>
          !e.unavailable.includes(date) &&
          !assigned.find(a => a.e === e)
        )
        .sort((a, b) => score(b, date) - score(a, date));

      if (!candidates.length) {
        throw new Error(`Nema dostupnih zaposlenih za ${date}`);
      }

      const chosen = candidates[0];
      const remaining = 1 - total;
      const s = Math.min(shareToNumber(chosen.dutyShare), remaining);

      assigned.push({ e: chosen, share: s });
      total += s;
      chosen.assignedDates.push(date);
      if (isWeekend(date)) chosen.weekendCount++;
    }

    // 3. VALIDACIJA ≥ 50% SPECIJALISTA
    const specialistShare = assigned
      .filter(a => a.e.role === "specijalista")
      .reduce((sum, a) => sum + a.share, 0);

    if (specialistShare < 0.5) {
      throw new Error(`Manje od 50% specijalista na dan ${date}`);
    }

    // 4. GLAVNI PRVI
    const chief = assigned.find(a => a.e.canBeChief);
    const ordered = [
      ...(chief ? [chief] : []),
      ...assigned.filter(a => a !== chief)
    ];

    schedule[date] = ordered.map(a => ({
      name: a.e.name,
      role: a.e.role,
      share:
        a.share === 1 ? "1/1" :
        a.share === 0.5 ? "1/2" :
        "1/4"
    }));
  }

  return schedule;
};