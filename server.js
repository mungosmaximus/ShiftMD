// ShiftMD v2.0 - Server za generisanje rasporeda dežurstava
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

if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
  console.log(`✅ ShiftMD: Kreiran folder uploads`);
}
if (!fs.existsSync(rasporediDir)) {
  fs.mkdirSync(rasporediDir, { recursive: true });
  console.log(`✅ ShiftMD: Kreiran folder rasporedi`);
}

const upload = multer({ dest: uploadsDir });

app.use(express.static(__dirname));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use("/rasporedi", express.static(rasporediDir));

// ===== RUTE ZA STRANICE =====

// Početna strana
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

// Strana "O sistemu"
app.get("/o-sistemu", (req, res) => {
  res.sendFile(path.join(__dirname, "o-sistemu.html"));
});

// Strana "Uputstvo"
app.get("/uputstvo", (req, res) => {
  res.sendFile(path.join(__dirname, "uputstvo.html"));
});

// Preuzimanje demo Excel fajla
app.get("/demo-lekari.xlsx", (req, res) => {
  const demoPath = path.join(__dirname, "demo-lekari.xlsx");
  if (fs.existsSync(demoPath)) {
    res.setHeader("Content-Disposition", "attachment; filename=demo-lekari.xlsx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    const fileStream = fs.createReadStream(demoPath);
    fileStream.on("error", (err) => {
      if (!res.headersSent) res.status(500).send("Greška pri čitanju demo fajla.");
    });
    fileStream.pipe(res);
  } else {
    res.status(404).send("Demo fajl nije pronađen. Pokrenite 'node create-demo-excel.js' da ga kreirate.");
  }
});

// ===== API RUTE =====

// Preuzimanje generisanog Excel fajla
app.get("/download/:filename", (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(rasporediDir, filename);
  
  console.log(`📥 ShiftMD: Preuzimanje ${filename}`);
  
  if (!fs.existsSync(filePath)) {
    return res.status(404).send("Fajl nije pronađen.");
  }
  
  const stats = fs.statSync(filePath);
  
  res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(filename)}"`);
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Length", stats.size);
  
  const fileStream = fs.createReadStream(filePath);
  fileStream.on("error", (err) => {
    if (!res.headersSent) res.status(500).send("Greška pri čitanju fajla.");
  });
  fileStream.pipe(res);
});

// Generisanje rasporeda
app.post("/generate", upload.single("excel"), async (req, res) => {
  try {
    const { hospital, month, year, dutyDates, dutyStaffCount, allowLessSpecialists } = req.body;
    
    if (!hospital || !month || !year || !dutyDates) {
      return res.status(400).json({ error: "Nedostaju obavezni podaci." });
    }
    if (!req.file) {
      return res.status(400).json({ error: "Excel fajl nije otpremljen." });
    }

    let dates;
    try { dates = JSON.parse(dutyDates); } catch (e) {
      return res.status(400).json({ error: "Neispravan format datuma." });
    }
    if (!Array.isArray(dates) || dates.length === 0) {
      return res.status(400).json({ error: "Nema izabranih datuma." });
    }

    const staffCount = parseInt(dutyStaffCount) || 1;
    const allowLess = allowLessSpecialists === "true";
    
    console.log("");
    console.log("═".repeat(60));
    console.log("🩺 ShiftMD v2.0 - GENERISANJE RASPOREDA");
    console.log("═".repeat(60));
    console.log(`   Ustanova: ${hospital} | ${month}/${year}`);
    console.log(`   Datuma: ${dates.length} | Dežurstava/dan: ${staffCount}`);
    console.log(`   Opušteno pravilo: ${allowLess ? 'DA' : 'NE'}`);
    
    const employees = await parseExcel(req.file.path);
    console.log(`   ✅ Učitano zaposlenih: ${employees.length}`);
    
    const schedule = generateSchedule(employees, dates, staffCount, allowLess);
    console.log(`   ✅ Raspored generisan`);

    const sortedDates = [...dates].sort((a, b) => {
      const [da, ma, ya] = a.split("-").map(Number);
      const [db, mb, yb] = b.split("-").map(Number);
      return new Date(ya, ma - 1, da) - new Date(yb, mb - 1, db);
    });

    const now = new Date();
    const timestamp = now.toISOString().replace(/[:.]/g, "-").substring(0, 19);
    const safeHospitalName = hospital
      .replace(/[^a-zA-Z0-9čćžšđČĆŽŠĐ\s]/g, "")
      .replace(/\s+/g, "_")
      .substring(0, 50);
    const excelFilename = `${safeHospitalName}_-_${timestamp}.xlsx`;
    const excelPath = path.join(rasporediDir, excelFilename);

    const resultData = { 
      hospital, month, year, 
      dutyStaffCount: staffCount,
      allowLessSpecialists: allowLess,
      schedule 
    };

    const trackingEntries = employees.map(emp => ({
      name: emp.name,
      role: emp.role,
      totalAssignedShare: emp.totalAssignedShare || 0,
      assignedDates: emp.assignedDates || [],
      dutyCount: emp.dutyCount || emp.assignedDates?.length || 0,
      weekendCount: emp.weekendCount || 0,
      canBeChief: emp.canBeChief,
      seniority: emp.seniority || 0
    }));

    const statistics = {
      entries: trackingEntries,
      firstDate: sortedDates[0] || "",
      lastDate: sortedDates[sortedDates.length - 1] || "",
      totalDays: sortedDates.length,
      totalDutySlots: staffCount * sortedDates.length,
      totalWeekendDays: sortedDates.filter(d => isWeekend(d)).length
    };
    
    await generateExcel(resultData, excelPath, statistics);
    
    if (!fs.existsSync(excelPath)) {
      throw new Error("Excel fajl nije uspešno kreiran.");
    }
    
    const fileStats = fs.statSync(excelPath);
    console.log(`   ✅ Excel: ${excelFilename} (${(fileStats.size / 1024).toFixed(1)} KB)`);
    console.log("═".repeat(60));
    console.log("");

    try { fs.unlinkSync(req.file.path); } catch (e) {}
    
    res.json({
      hospital, month, year, 
      dutyStaffCount: staffCount,
      allowLessSpecialists: allowLess,
      schedule,
      excelFile: excelFilename,
      excelSize: fileStats.size,
      generatedAt: now.toISOString()
    });
    
  } catch (err) {
    console.error("❌ ShiftMD greška:", err.message);
    if (req.file && req.file.path) {
      try { fs.unlinkSync(req.file.path); } catch (e) {}
    }
    res.status(400).json({ error: err.message });
  }
});

// Lista svih generisanih rasporeda
app.get("/rasporedi-list", (req, res) => {
  try {
    if (!fs.existsSync(rasporediDir)) {
      return res.json([]);
    }
    
    const files = fs.readdirSync(rasporediDir)
      .filter(f => f.endsWith(".xlsx"))
      .map(f => {
        const filePath = path.join(rasporediDir, f);
        const fileStats = fs.statSync(filePath);
        return {
          filename: f,
          size: fileStats.size,
          sizeFormatted: formatFileSize(fileStats.size),
          created: fileStats.birthtime,
          modified: fileStats.mtime,
          url: `/download/${encodeURIComponent(f)}`
        };
      })
      .sort((a, b) => new Date(b.created) - new Date(a.created));
    
    console.log(`📋 ShiftMD: Lista rasporeda - ${files.length} fajlova`);
    res.json(files);
  } catch (err) {
    console.error("ShiftMD: Greška pri čitanju foldera:", err);
    res.status(500).json({ error: "Greška pri čitanju foldera sa rasporedima." });
  }
});

// ===== HELPER FUNKCIJE =====

function formatFileSize(bytes) {
  if (!bytes || bytes === 0) return '0 B';
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

// ===== POKRETANJE SERVERA =====

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("");
  console.log("╔══════════════════════════════════════════════╗");
  console.log("║        🩺 ShiftMD v2.0 - POKRENUT           ║");
  console.log("╠══════════════════════════════════════════════╣");
  console.log(`║  URL: http://localhost:${PORT}                  ║`);
  console.log("║  /              - Početna strana            ║");
  console.log("║  /o-sistemu     - O sistemu                 ║");
  console.log("║  /uputstvo      - Uputstvo                  ║");
  console.log("║  Udeli: celo, polovina, trećina,            ║");
  console.log("║         četvrtina, osmina                   ║");
  console.log("║  Scoring: napredni, burnout zaštita         ║");
  console.log("╚══════════════════════════════════════════════╝");
  console.log("");
  console.log("▶  Pritisni Ctrl+C za zaustavljanje servera");
  console.log("");
});

// Graceful shutdown
process.on("SIGINT", () => {
  console.log("");
  console.log("🛑 ShiftMD: Server se zaustavlja...");
  process.exit(0);
});

process.on("SIGTERM", () => {
  console.log("");
  console.log("🛑 ShiftMD: Server se zaustavlja...");
  process.exit(0);
});