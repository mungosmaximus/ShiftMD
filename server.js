const express = require("express");
const multer = require("multer");

const parseExcel = require("./backend/excelParser");
const generateSchedule = require("./backend/scheduler");
const generatePDF = require("./backend/pdfGenerator");

const app = express();
const upload = multer({ dest: "uploads/" });

app.use(express.json());
app.use(express.static("frontend"));

let LAST_RESULT = null;

app.post("/generate", upload.single("excel"), (req, res) => {
  try {
    const { hospital, month, year, dutyDates } = req.body;
    const dates = JSON.parse(dutyDates);

    const employees = parseExcel(req.file.path);
    const schedule = generateSchedule(employees, dates);

    LAST_RESULT = { hospital, month, year, schedule };
    res.json(LAST_RESULT);
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

app.get("/pdf", (req, res) => {
  if (!LAST_RESULT) {
    return res.status(400).send("Raspored još nije generisan.");
  }
  generatePDF(LAST_RESULT, res);
});

app.listen(3000, () => {
  console.log("✅ Duty Scheduler v1.4 running at http://localhost:3000");
});