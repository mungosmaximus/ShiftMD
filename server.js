// ShiftMD v3.6 - Server sa podrškom za dva tipa dežurstava
const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");

const parseExcel = require("./excelParser");
const generateSchedule = require("./scheduler");
const generateExcel = require("./excelGenerator");
const { isWeekend } = require("./utils");

const app = express();

const uploadsDir = path.join(__dirname, "uploads");
const rasporediDir = path.join(__dirname, "rasporedi");

if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir, { recursive: true });
if (!fs.existsSync(rasporediDir)) fs.mkdirSync(rasporediDir, { recursive: true });

const upload = multer({ dest: uploadsDir });

app.use(express.static(__dirname));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use("/rasporedi", express.static(rasporediDir));

app.get("/", (req, res) => res.sendFile(path.join(__dirname, "index.html")));
app.get("/o-sistemu", (req, res) => res.sendFile(path.join(__dirname, "o-sistemu.html")));
app.get("/uputstvo", (req, res) => res.sendFile(path.join(__dirname, "uputstvo.html")));

app.get("/demo-lekari.xlsx", (req, res) => {
  const p = path.join(__dirname, "demo-lekari.xlsx");
  if (fs.existsSync(p)) { res.setHeader("Content-Disposition", "attachment; filename=demo-lekari.xlsx"); res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); fs.createReadStream(p).pipe(res); }
  else res.status(404).send("Demo nije pronađen.");
});
app.get("/demo-physicians.xlsx", (req, res) => {
  const p = path.join(__dirname, "demo-lekari.xlsx");
  if (fs.existsSync(p)) { res.setHeader("Content-Disposition", "attachment; filename=demo-physicians.xlsx"); res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); fs.createReadStream(p).pipe(res); }
  else res.status(404).send("Demo file not found.");
});

app.get("/download/:f", (req, res) => {
  const p = path.join(rasporediDir, req.params.f);
  if (!fs.existsSync(p)) return res.status(404).send("Nije pronađen.");
  res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(req.params.f)}"`);
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  fs.createReadStream(p).pipe(res);
});

app.post("/generate", upload.single("excel"), async (req, res) => {
  try {
    const { hospital, month, year, dutyDates1, dutyStaffCount1, allowLessSpecialists1, useRankedDuties1, enableType2, dutyDates2, dutyStaffCount2, allowLessSpecialists2, useRankedDuties2, customName1, customName2 } = req.body;
    
    if (!hospital || !month || !year || !dutyDates1) return res.status(400).json({ error: "Nedostaju obavezni podaci." });
    if (!req.file) return res.status(400).json({ error: "Excel nije otpremljen." });

    let dates1; try { dates1 = JSON.parse(dutyDates1); } catch (e) { return res.status(400).json({ error: "Neispravni datumi za Tip 1." }); }
    if (!Array.isArray(dates1) || dates1.length === 0) return res.status(400).json({ error: "Morate izabrati bar jedan dan za Tip 1." });

    const hasType2 = enableType2 === "true";
    let dates2 = [];
    
    if (hasType2) {
      try { dates2 = JSON.parse(dutyDates2); } catch (e) { return res.status(400).json({ error: "Neispravni datumi za Tip 2." }); }
      if (!Array.isArray(dates2) || dates2.length === 0) return res.status(400).json({ error: "Morate izabrati bar jedan dan za Tip 2." });
      for (const d of dates1) { if (dates2.includes(d)) return res.status(400).json({ error: `Datum ${d} se preklapa između Tip 1 i Tip 2!` }); }
    }

    const staffCount1 = parseInt(dutyStaffCount1) || 1;
    const allowLess1 = allowLessSpecialists1 === "true";
    const useRanked1 = useRankedDuties1 === "true";
    
    const staffCount2 = hasType2 ? (parseInt(dutyStaffCount2) || 1) : 0;
    const allowLess2 = hasType2 ? (allowLessSpecialists2 === "true") : false;
    const useRanked2 = hasType2 ? (useRankedDuties2 === "true") : false;

    const type1Name = customName1 || "Tip 1";
    const type2Name = customName2 || "Tip 2";

    console.log(`\n🩺 ShiftMD v3.6 - ${hospital} | ${month}/${year}`);
    console.log(`   ${type1Name}: ${dates1.length} dana, ${staffCount1} dežurstava/dan${useRanked1 ? ', rangirani' : ''}`);
    if (hasType2) console.log(`   ${type2Name}: ${dates2.length} dana, ${staffCount2} dežurstava/dan${useRanked2 ? ', rangirani' : ''}`);

    const employees = await parseExcel(req.file.path);
    console.log(`   ✅ Učitano ${employees.length} lekara`);

    const schedule1 = generateSchedule(employees, dates1, staffCount1, allowLess1, useRanked1);
    let schedule2 = {};
    if (hasType2) {
      schedule2 = generateSchedule(employees, dates2, staffCount2, allowLess2, useRanked2);
    }

    const schedule = { ...schedule1, ...schedule2 };
    const dutyTypes = {};
    dates1.forEach(d => dutyTypes[d] = type1Name);
    dates2.forEach(d => dutyTypes[d] = type2Name);

    const allDates = [...dates1, ...dates2].sort((a, b) => {
      const [da, ma, ya] = a.split("-").map(Number);
      const [db, mb, yb] = b.split("-").map(Number);
      return new Date(ya, ma - 1, da) - new Date(yb, mb - 1, db);
    });

    const now = new Date();
    const ts = now.toISOString().replace(/[:.]/g, "-").substring(0, 19);
    const name = hospital.replace(/[^a-zA-Z0-9čćžšđČĆŽŠĐ\s]/g, "").replace(/\s+/g, "_").substring(0, 50);
    const xlFile = `${name}_-_${ts}.xlsx`;
    const xlPath = path.join(rasporediDir, xlFile);

    const resultData = { hospital, month, year, schedule, dutyTypes, hasType2, type1Name, type2Name, staffCount1, staffCount2, allowLess1, allowLess2, useRanked1, useRanked2 };

    const tracking = employees.map(e => ({
      name: e.name, role: e.role, totalAssignedShare: e.totalAssignedShare || 0,
      assignedDates: e.assignedDates || [], dutyCount: e.dutyCount || e.assignedDates?.length || 0,
      weekendCount: e.weekendCount || 0, canBeChief: e.canBeChief, seniority: e.seniority || 0, rank: e.rank || 0
    }));
    const stats = { entries: tracking, firstDate: allDates[0] || "", lastDate: allDates[allDates.length-1] || "", totalDays: allDates.length };

    await generateExcel(resultData, xlPath, stats);
    if (!fs.existsSync(xlPath)) throw new Error("Excel nije kreiran.");

    const fStats = fs.statSync(xlPath);
    console.log(`   ✅ ${xlFile} (${(fStats.size/1024).toFixed(1)} KB)\n`);
    try { fs.unlinkSync(req.file.path); } catch (e) {}

    res.json({
      hospital, month, year,
      schedule,
      dutyTypes,
      hasType2,
      type1Name, type2Name,
      type1Count: dates1.length, type2Count: dates2.length,
      staffCount1, staffCount2,
      useRanked1, useRanked2,
      allowLess1, allowLess2,
      excelFile: xlFile, excelSize: fStats.size, generatedAt: now.toISOString()
    });
  } catch (err) {
    console.error("❌", err.message);
    if (req.file?.path) try { fs.unlinkSync(req.file.path); } catch (e) {}
    res.status(400).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`\n🩺 ShiftMD v3.6 na http://localhost:${PORT}\n`));
process.on("SIGINT", () => { console.log("\n🛑 Zaustavljam..."); process.exit(0); });