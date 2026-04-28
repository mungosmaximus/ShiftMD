const PDFDocument = require("pdfkit");

module.exports = function generatePDF(data, res) {
  const { hospital, month, year, schedule } = data;

  const doc = new PDFDocument({ margin: 40 });

  res.setHeader(
    "Content-Disposition",
    `attachment; filename=Raspored_dezurstava_${hospital}_${month}-${year}.pdf`
  );
  res.setHeader("Content-Type", "application/pdf");

  doc.pipe(res);

  doc.fontSize(18).text("RASPORED DEŽURSTAVA", { align: "center" });
  doc.moveDown(0.5);
  doc.fontSize(14).text(hospital, { align: "center" });
  doc.fontSize(12).text(`Mesec: ${month} / ${year}`, { align: "center" });
  doc.moveDown(1);

  Object.keys(schedule).forEach(date => {
    doc.fontSize(12).text(date);
    schedule[date].forEach(p => {
      doc.fontSize(10).text(
        `• ${p.name} – ${p.role} (${p.share})`,
        { indent: 20 }
      );
    });
    doc.moveDown(0.5);
  });

  doc.end();
};