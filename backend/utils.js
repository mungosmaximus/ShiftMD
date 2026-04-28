function parseDateList(cell) {
  if (!cell) return [];
  return cell.split(",").map(v => v.trim());
}

function shareToNumber(share) {
  if (share === "1/1") return 1;
  if (share === "1/2") return 0.5;
  if (share === "1/4") return 0.25;
  throw new Error("Neispravan udeo dežurstva");
}

function isWeekend(dateStr) {
  const [d, m, y] = dateStr.split("-");
  const date = new Date(`${y}-${m}-${d}`);
  return date.getDay() === 0 || date.getDay() === 6;
}

module.exports = { parseDateList, shareToNumber, isWeekend };