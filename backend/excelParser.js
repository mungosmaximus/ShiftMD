const XLSX = require("xlsx");
const { parseDateList } = require("./utils");

module.exports = function parseExcel(path) {
  const wb = XLSX.readFile(path);
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet);

  return rows.map(row => ({
    name: row["Ime i prezime"],
    seniority: Number(row["Starešinstvo"]),
    canBeChief: row["Može biti glavni"] === "DA",
    role: row["Tip zaposlenog"], // specijalista | specijalizant
    dutyShare: row["Udeo dežurstva"], // 1/1 | 1/2 | 1/4
    unavailable: parseDateList(row["Ne može"]),
    preferred: parseDateList(row["Želi da dežura"]),
    assignedDates: [],
    weekendCount: 0
  }));
};
``