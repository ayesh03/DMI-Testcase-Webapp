// server.js
const express = require("express");
const cors = require("cors");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const app = express();
app.use(cors());
app.use(express.json({ limit: "200mb" }));
app.use(express.urlencoded({ limit: "200mb", extended: true }));

const EXCEL_PATH = path.join(__dirname, "DMI-Testcases-Excelsheet.xlsx");
const OUTPUT_NAME = "Updated_DMI-Testcases-Excelsheet.xlsx";

app.post("/generate-report", async (req, res) => {
  try {
    const { screenshots } = req.body;
    if (!Array.isArray(screenshots) || screenshots.length === 0) {
      return res.status(400).json({ error: "No screenshot data received" });
    }

    // Load original workbook
    const srcWb = new ExcelJS.Workbook();
    await srcWb.xlsx.readFile(EXCEL_PATH);

    const sheetName = "V1.4_DMI_Functional_testcase ";
    const srcSheet = srcWb.getWorksheet(sheetName);
    if (!srcSheet) {
      return res.status(404).json({ error: `Worksheet '${sheetName}' not found` });
    }

    // Create a new workbook and copy only this sheet
    const dstWb = new ExcelJS.Workbook();
    const dstSheet = dstWb.addWorksheet(sheetName);

    // Force column widths (I = 40)
    dstSheet.columns = [
      { header: "A", key: "A", width: 10 },
      { header: "B", key: "B", width: 35 },
      { header: "C", key: "C", width: 45 },
      { header: "D", key: "D", width: 45 },
      { header: "E", key: "E", width: 35 },
      { header: "F", key: "F", width: 45 },
      { header: "G", key: "G", width: 40 },
      { header: "H", key: "H", width: 45 },
      { header: "I", key: "I", width: 40 }, // Screenshot column
    ];

    // Copy cell values and styles
    srcSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const dstRow = dstSheet.getRow(rowNumber + 1);
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        dstRow.getCell(colNumber).value = cell.value;
        if (cell.style) {
          try {
            dstRow.getCell(colNumber).style = JSON.parse(JSON.stringify(cell.style));
          } catch {}
        }
      });
      dstRow.commit();
    });

    //  Add header "Screenshot Report"
    dstSheet.spliceRows(1, 0, ["Screenshot Report"]);
    const header = dstSheet.getRow(1);
    header.height = 30;
    //header.getCell(1).value = " Screenshot Report";
    header.getCell(1).font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
    header.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
    header.getCell(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "4472C4" },
    };
    dstSheet.mergeCells("A1:I1");

    // Build TestCaseID → row map (after header offset)
    const idToRow = {};
    dstSheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return; // skip header
      const cell = row.getCell(2);
      let val = cell ? cell.value : null;
      if (val && typeof val === "object" && val.richText) {
        val = val.richText.map(t => t.text).join("");
      }
      if (val) idToRow[val.toString().trim().toUpperCase()] = rowNumber;
    });

    // Add Screenshot column header
    const headerCell = dstSheet.getCell("I3");
    headerCell.value = "Screenshot Image";
    headerCell.font = { bold: true };
    headerCell.alignment = { horizontal: "center", vertical: "middle" };

    //  Add screenshots
    let addedCount = 0;
    for (const entry of screenshots) {
      const { id, base64 } = entry;
      if (!id || !base64) continue;

      const key = id.toString().trim().toUpperCase();
      const targetRow = idToRow[key];
      if (!targetRow) continue;

      // Detect format
      let extension = "jpeg";
      if (base64.startsWith("iVBORw0KGgo")) extension = "png";
      else if (base64.startsWith("/9j/")) extension = "jpeg";

      const cleanBase64 = base64.replace(/^data:image\/\w+;base64,/, "").replace(/\s/g, "");
      const buffer = Buffer.from(cleanBase64, "base64");
      if (!buffer || buffer.length < 100) continue;

      const imageId = dstWb.addImage({ buffer, extension });

      //  Calculate dynamic image and row height
      const colWidth = dstSheet.getColumn(9).width || 40; // Column I width
      const baseImageWidth = colWidth * 7.5;
      const aspectRatio = 4 / 3; // approximate standard ratio
      const dynamicHeightPx = baseImageWidth / aspectRatio; // height from ratio

      // Convert to Excel row height unit (~0.75 px per unit)
      const excelRowHeight = dynamicHeightPx / 0.75;

      const row = dstSheet.getRow(targetRow);
      row.height = Math.max(45, Math.min(excelRowHeight, 120)); // limit max height
      row.commit();

      // Add image inside the cell box
      const heightPx = row.height * 0.75;
      dstSheet.addImage(imageId, {
        tl: { col: 8, row: targetRow - 1 },
        ext: { width: baseImageWidth, height: heightPx },
      });

      addedCount++;
    }

    // Save new Excel
    const outputPath = path.join(__dirname, "Updated_DMI-Testcases-Excelsheet.xlsx");
    await dstWb.xlsx.writeFile(outputPath);

    console.log(` ${addedCount} screenshots added with dynamic row heights.`);
    return res.json({ success: true, file: "Updated_DMI-Testcases-Excelsheet.xlsx", added: addedCount });
  } catch (err) {
    console.error(" Error generating report:", err);
    return res.status(500).json({ error: err.message });
  }
});

// Download endpoint
app.get("/download", (req, res) => {
  const file = path.join(__dirname, OUTPUT_NAME);
  if (fs.existsSync(file)) {
    res.download(file);
  } else {
    res.status(404).send("File not found");
  }
});

const PORT = 5000;
// Serve frontend index.html (for Render or local)
app.use(express.static(__dirname)); // serves index.html, Excel file, etc.
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

app.listen(PORT, () => console.log(` Server running at http://localhost:${PORT}`));
