const ExcelJS = require("exceljs");
const { Readable } = require("stream");

// ============================================================
// HELPER: patch native Node.js res agar punya .status().json()
// (Vercel sudah punya ini built-in, tapi lokal perlu di-patch)
// ============================================================
function patchRes(res) {
  if (typeof res.status === "function") return res;
  res.status = function (code) {
    res.statusCode = code;
    return res;
  };
  res.json = function (data) {
    res.setHeader("Content-Type", "application/json");
    res.end(JSON.stringify(data));
    return res;
  };
  return res;
}

// ============================================================
// CONFIG — ubah path file di sini jika membaca dari filesystem
// (untuk production Vercel, file dikirim via request body/base64)
// ============================================================
const FILE_CONFIG = {
  // Nama sheet yang akan diproses (null = semua sheet)
  targetSheetName: null,

  // Nama file output yang akan dikembalikan ke client
  outputFileName: "output.xlsx",
};

// ============================================================
// PIPELINE STEPS
// Tambahkan step baru di sini secara berurutan.
// Setiap step menerima (workbook, worksheet) dan memodifikasi langsung.
// ============================================================
const STEPS = [
  // ----------------------------------------------------------
  // STEP 1: Ubah semua font MENJADI Arial ukuran 6
  // ----------------------------------------------------------
  async function step1_setFontArialSize6(workbook, worksheet) {
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        // Pertahankan properti font yang sudah ada, hanya timpa name & size
        cell.font = {
          ...(cell.font || {}),
          name: "Arial",
          size: 6,
        };
      });
    });
  },

  // ----------------------------------------------------------
  // STEP 2: Cari kolom A dengan format XXX.XX.XX, ubah warna
  //         font seluruh baris MENJADI #0c0c5e
  // ----------------------------------------------------------
  async function step2_colorCodeRows(workbook, worksheet) {
    const pattern = /^\d{3}\.\d{2}\.[A-Za-z0-9]{2}$/;

    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const cellA = row.getCell(1);
      const value = cellA.value ? String(cellA.value).trim() : "";

      if (pattern.test(value)) {
        row.eachCell({ includeEmpty: true }, (cell) => {
          cell.font = {
            ...(cell.font || {}),
            color: { argb: "FF0c0c5e" },
          };
        });
      }
    });
  },

  // ----------------------------------------------------------
  // STEP 3: Cari kolom A dengan 4 digit angka (contoh: 2175),
  //         ubah warna font seluruh baris MENJADI #0000FF
  // ----------------------------------------------------------
  async function step3_colorFourDigitRows(workbook, worksheet) {
    const pattern = /^\d{4}$/;

    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const cellA = row.getCell(1);
      const value = cellA.value ? String(cellA.value).trim() : "";

      if (pattern.test(value)) {
        row.eachCell({ includeEmpty: true }, (cell) => {
          cell.font = {
            ...(cell.font || {}),
            color: { argb: "FF0000FF" },
          };
        });
      }
    });
  },

  // ----------------------------------------------------------
  // STEP 4: Cari kolom A dengan format XXXX.XXX (4 digit angka,
  //         titik, 3 karakter huruf/angka), ubah warna font
  //         seluruh baris MENJADI #B10301
  // ----------------------------------------------------------
  async function step4_colorCode43Rows(workbook, worksheet) {
    const pattern = /^\d{4}\.[A-Za-z0-9]{3}$/;

    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const cellA = row.getCell(1);
      const value = cellA.value ? String(cellA.value).trim() : "";

      if (pattern.test(value)) {
        row.eachCell({ includeEmpty: true }, (cell) => {
          cell.font = {
            ...(cell.font || {}),
            color: { argb: "FFB10301" },
          };
        });
      }
    });
  },

  // ----------------------------------------------------------
  // STEP 5: Header styling A1:AU3
  //  - Fill #0070C0 pada A1:AT2, font putih
  //  - Fill #BFBFBF pada A3:AT3
  //  - Font Calibri seluruh A1:AU3
  //  - Font size 12 → A1:AT1, AU1:AU3
  //  - Font size 10 → A2:AT2
  //  - Font size 9  → A3:AT3
  // ----------------------------------------------------------
  async function step5_headerStyling(workbook, worksheet) {
    // Helper: apply ke range baris & kolom
    function applyRange(rowStart, rowEnd, colStart, colEnd, applyFn) {
      for (let r = rowStart; r <= rowEnd; r++) {
        const row = worksheet.getRow(r);
        for (let c = colStart; c <= colEnd; c++) {
          applyFn(row.getCell(c));
        }
      }
    }

    // 1. Fill #0070C0 + font putih → A1:AT2 (baris 1-2, col 1-46)
    applyRange(1, 2, 1, 46, (cell) => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0070C0" } };
      cell.font = { ...(cell.font || {}), color: { argb: "FFFFFFFF" } };
    });

    // 2. Fill #BFBFBF → A3:AT3 (baris 3, col 1-46)
    applyRange(3, 3, 1, 46, (cell) => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFBFBFBF" } };
    });

    // 3. Font Calibri → A1:AU3 (baris 1-3, col 1-47)
    applyRange(1, 3, 1, 47, (cell) => {
      cell.font = { ...(cell.font || {}), name: "Calibri" };
    });

    // 4. Font size 12 → A1:AT1 (baris 1, col 1-46)
    applyRange(1, 1, 1, 46, (cell) => {
      cell.font = { ...(cell.font || {}), size: 12 };
    });

    // 5. Font size 10 → A2:AT2 (baris 2, col 1-46)
    applyRange(2, 2, 1, 46, (cell) => {
      cell.font = { ...(cell.font || {}), size: 10 };
    });

    // 6. Font size 9 → A3:AT3 (baris 3, col 1-46)
    applyRange(3, 3, 1, 46, (cell) => {
      cell.font = { ...(cell.font || {}), size: 9 };
    });

    // 7. Font size 12 → AU1:AU3 (baris 1-3, col 47)
    applyRange(1, 3, 47, 47, (cell) => {
      cell.font = { ...(cell.font || {}), size: 12 };
    });

    // 8. Font color #FF0000 + fill #FFFF00 → AU1:AU3 (baris 1-3, col 47)
    applyRange(1, 3, 47, 47, (cell) => {
      cell.font = { ...(cell.font || {}), color: { argb: "FFFF0000" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
    });
  },

  // ----------------------------------------------------------
  // STEP 6: Border
  //  - All border putih  → A1:AT2
  //  - All border hitam  → A3:AT3
  //  - All border hitam  → AU1:AU3
  // ----------------------------------------------------------
  async function step6_borders(workbook, worksheet) {
    function applyBorder(rowStart, rowEnd, colStart, colEnd, borderColor) {
      const border = {
        top:      { style: "thin", color: { argb: borderColor } },
        bottom:   { style: "thin", color: { argb: borderColor } },
        left:     { style: "thin", color: { argb: borderColor } },
        right:    { style: "thin", color: { argb: borderColor } },
      };
      for (let r = rowStart; r <= rowEnd; r++) {
        const row = worksheet.getRow(r);
        for (let c = colStart; c <= colEnd; c++) {
          const cell = row.getCell(c);
          cell.border = { ...border };
        }
      }
    }

    // All border putih → A1:AT2 (baris 1-2, col 1-46)
    applyBorder(1, 2, 1, 46, "FFFFFFFF");

    // All border hitam → A3:AT3 (baris 3, col 1-46)
    applyBorder(3, 3, 1, 46, "FF000000");

    // All border hitam → AU1:AU3 (baris 1-3, col 47)
    applyBorder(1, 3, 47, 47, "FF000000");
  },

  // ----------------------------------------------------------
  // STEP 7: Border hitam outside pada range dinamis (XX = baris
  //         terbawah yang berisi data di kolom B, mulai baris 4)
  // ----------------------------------------------------------
  async function step7_dynamicBorders(workbook, worksheet) {
    // Cari baris terbawah yang berisi data di kolom B (col index 2)
    let lastRow = 4;
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const cellB = row.getCell(2);
      if (row.number >= 4 && cellB.value !== null && cellB.value !== undefined && cellB.value !== "") {
        if (row.number > lastRow) lastRow = row.number;
      }
    });

    console.log(`[Step 7] Last row with data in col B: ${lastRow}`);

    // Helper: outside border pada range
    function outsideBorder(rowStart, rowEnd, colStart, colEnd) {
      const black = { style: "thin", color: { argb: "FF000000" } };

      for (let r = rowStart; r <= rowEnd; r++) {
        const row = worksheet.getRow(r);
        for (let c = colStart; c <= colEnd; c++) {
          const cell = row.getCell(c);
          const current = cell.border || {};
          cell.border = {
            ...current,
            top:    r === rowStart ? black : current.top,
            bottom: r === rowEnd   ? black : current.bottom,
            left:   c === colStart ? black : current.left,
            right:  c === colEnd   ? black : current.right,
          };
        }
      }
    }

    // Helper: bottom border pada range
    function bottomBorder(rowStart, rowEnd, colStart, colEnd) {
      const black = { style: "thin", color: { argb: "FF000000" } };
      for (let r = rowStart; r <= rowEnd; r++) {
        const row = worksheet.getRow(r);
        for (let c = colStart; c <= colEnd; c++) {
          const cell = row.getCell(c);
          cell.border = { ...(cell.border || {}), bottom: black };
        }
      }
    }

    const XX = lastRow;

    // OUTSIDE BORDERS
    outsideBorder(4, XX,  1,  1);  // A4:AXX
    outsideBorder(4, XX,  2, 14);  // B4:NXX
    outsideBorder(4, XX, 15, 15);  // O4:OXX
    outsideBorder(4, XX, 16, 16);  // P4:PXX
    outsideBorder(4, XX, 19, 19);  // S4:SXX
    outsideBorder(4, XX, 21, 21);  // U4:UXX
    outsideBorder(4, XX, 22, 22);  // V4:VXX
    outsideBorder(4, XX, 23, 23);  // W4:WXX
    outsideBorder(4, XX, 24, 24);  // X4:XXX
    outsideBorder(4, XX, 25, 37);  // Y4:AKXX
    outsideBorder(4, XX, 38, 38);  // AL4:ALXX
    outsideBorder(4, XX, 39, 39);  // AM4:AMXX
    outsideBorder(4, XX, 42, 42);  // AP4:APXX
    outsideBorder(4, XX, 44, 44);  // AR4:ARXX
    outsideBorder(4, XX, 45, 45);  // AS4:ASXX
    outsideBorder(4, XX, 46, 46);  // AT4:ATXX
    outsideBorder(4, XX, 47, 47);  // AU4:AUXX

    // BOTTOM BORDER A:AU pada baris XX
    bottomBorder(XX, XX, 1, 47);   // AXX:AUXX
  },

  // ----------------------------------------------------------
  // STEP 8: Hapus isi cell T(XX+1), T(XX+5), T(XX+6)
  //         XX = baris terbawah berisi data di kolom B
  // ----------------------------------------------------------
  async function step8_clearCells(workbook, worksheet) {
    // Cari XX (sama seperti step 7)
    let lastRow = 4;
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const cellB = row.getCell(2);
      if (row.number >= 4 && cellB.value !== null && cellB.value !== undefined && cellB.value !== "") {
        if (row.number > lastRow) lastRow = row.number;
      }
    });

    const XX = lastRow;
    console.log(`[Step 8] XX = ${XX}`);

    const targetRows = [XX + 1, XX + 5, XX + 6];
    for (const r of targetRows) {
      const cell = worksheet.getRow(r).getCell(20); // T = col 20
      cell.value = null;
    }
  },

  // ----------------------------------------------------------
  // STEP 9: Center dan middle align A1:AU3
  // ----------------------------------------------------------
  async function step9_alignHeader(workbook, worksheet) {
    for (let r = 1; r <= 3; r++) {
      const row = worksheet.getRow(r);
      for (let c = 1; c <= 47; c++) { // A=1 sampai AU=47
        const cell = row.getCell(c);
        cell.alignment = {
          ...(cell.alignment || {}),
          horizontal: "center",
          vertical: "middle",
        };
      }
    }
  },

  // ----------------------------------------------------------
  // STEP 10: Cari exact match "Satuan Ukur" dan "Biaya Satuan Ukur"
  //          lalu ubah teks antar kata MENJADI newline (alt+enter)
  //          dan set wrapText: true
  // ----------------------------------------------------------
  async function step10_newlineWords(workbook, worksheet) {
    const targets = {
      "Satuan Ukur":        "Satuan\nUkur",
      "Biaya Satuan Ukur":  "Biaya\nSatuan\nUkur",
    };

    worksheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const val = cell.value;
        if (typeof val !== "string") return;

        const trimmed = val.trim();
        if (targets[trimmed] !== undefined) {
          cell.value = targets[trimmed];
          cell.alignment = {
            ...(cell.alignment || {}),
            wrapText: true,
          };
        }
      });
    });
  },

  // ----------------------------------------------------------
  // STEP 11: Pindahkan teks AQ(XX+1), AQ(XX+5), AQ(XX+6)
  //          ke AP(XX+1), AP(XX+5), AP(XX+6)
  //          XX = baris terbawah berisi data di kolom B
  // ----------------------------------------------------------
  async function step11_moveCells(workbook, worksheet) {
    // Cari XX
    let lastRow = 4;
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const cellB = row.getCell(2);
      if (row.number >= 4 && cellB.value !== null && cellB.value !== undefined && cellB.value !== "") {
        if (row.number > lastRow) lastRow = row.number;
      }
    });

    const XX = lastRow;
    console.log(`[Step 11] XX = ${XX}`);

    const targetRows = [XX + 1, XX + 5, XX + 6];
    for (const r of targetRows) {
      const srcCell = worksheet.getRow(r).getCell(43); // AQ = col 43
      const dstCell = worksheet.getRow(r).getCell(42); // AP = col 42

      // Pindahkan value dan style
      dstCell.value = srcCell.value;
      dstCell.style = JSON.parse(JSON.stringify(srcCell.style || {}));

      // Kosongkan cell sumber
      srcCell.value = null;
    }
  },

  // ----------------------------------------------------------
  // STEP 12: Bold seluruh baris jika kolom A match:
  //          - Code 433: format XXXX.XXX.XXX (4 digit . 3 char . 3 char)
  //          - 3 digit angka (contoh: 051)
  // ----------------------------------------------------------
  async function step12_boldRows(workbook, worksheet) {
    const patternCode433 = /^\d{4}\.[A-Za-z0-9]{3}\.[A-Za-z0-9]{3}$/;
    const pattern3Digit  = /^\d{3}$/;

    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const cellA = row.getCell(1);
      const value = cellA.value ? String(cellA.value).trim() : "";

      if (patternCode433.test(value) || pattern3Digit.test(value)) {
        row.eachCell({ includeEmpty: true }, (cell) => {
          cell.font = { ...(cell.font || {}), bold: true };
        });
      }
    });
  },


  // ----------------------------------------------------------
  // STEP 13: Pada setiap baris Y dimana kolom A match Code 433
  //          (XXXX.XXX.XXX), isi AW(Y) = "524 SEMULA"
  //          dan AW(Y+1) = formula SUM kolom U dari baris-baris
  //          dimana kolom A adalah 6 digit angka diawali "524"
  // ----------------------------------------------------------
  async function step13_fillAWColumn(workbook, worksheet) {
    const patternCode433 = /^\d{4}\.[A-Za-z0-9]{3}\.[A-Za-z0-9]{3}$/;
    const pattern524     = /^524\d{3}$/;
    const code433Rows = [];
    const rows524 = [];
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const val = row.getCell(1).value ? String(row.getCell(1).value).trim() : "";
      if (patternCode433.test(val)) code433Rows.push(row.number);
      if (pattern524.test(val))     rows524.push(row.number);
    });
    console.log(`[Step 13] Code 433 rows:`, code433Rows);
    console.log(`[Step 13] 524xxx rows:`, rows524);
    for (let i = 0; i < code433Rows.length; i++) {
      const Y        = code433Rows[i];
      const nextY    = code433Rows[i + 1] ?? Infinity;
      const colAW    = 49; // AW = col 49
      const cellY = worksheet.getRow(Y).getCell(colAW);
      cellY.value = "524 SEMULA";
      const rows524InRange = rows524.filter(r => r > Y && r < nextY);
      const formulaStr = rows524InRange.length > 0
        ? rows524InRange.map(r => `U${r}`).join("+")
        : "0";
      const cellY1 = worksheet.getRow(Y + 1).getCell(colAW);
      cellY1.value = { formula: `=${formulaStr}` };
    }
  },

  // ----------------------------------------------------------
  // STEP 14: Pada setiap baris Y dimana kolom A match Code 43
  //          (XXXX.XXX), isi AW(Y) = "524 SEMULA"
  //          dan AW(Y+1) = formula penjumlahan AW(Y+1) dari
  //          Step 13 yang berada antara Y dan Code 43 berikutnya
  // ----------------------------------------------------------
  async function step14_fillAWCode43(workbook, worksheet) {
    const patternCode43  = /^\d{4}\.[A-Za-z0-9]{3}$/;
    const patternCode433 = /^\d{4}\.[A-Za-z0-9]{3}\.[A-Za-z0-9]{3}$/;
    const code43Rows  = [];
    const code433Rows = [];
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const val = row.getCell(1).value ? String(row.getCell(1).value).trim() : "";
      if (patternCode43.test(val))  code43Rows.push(row.number);
      if (patternCode433.test(val)) code433Rows.push(row.number);
    });
    console.log(`[Step 14] Code 43 rows:`, code43Rows);
    console.log(`[Step 14] Code 433 rows (from Step 13):`, code433Rows);
    for (let i = 0; i < code43Rows.length; i++) {
      const Y     = code43Rows[i];
      const nextY = code43Rows[i + 1] ?? Infinity;
      const colAW = 49; // AW = col 49
      worksheet.getRow(Y).getCell(colAW).value = "524 SEMULA";
      const step13Cells = code433Rows
        .filter(r => r > Y && r < nextY)
        .map(r => `AW${r + 1}`);
      const formulaStr = step13Cells.length > 0
        ? step13Cells.join("+")
        : "0";
      worksheet.getRow(Y + 1).getCell(colAW).value = { formula: `=${formulaStr}` };
    }
  },

  // ----------------------------------------------------------
  // STEP 15: Pada setiap baris Y dimana kolom A match Code 322
  //          (XXX.XX.XX, contoh: 026.04.DN), isi AW(Y) = "524 SEMULA"
  //          dan AW(Y+1) = formula penjumlahan AW(Y+1) dari Step 14
  //          yang berada antara Y dan Code 322 berikutnya
  // ----------------------------------------------------------
  async function step15_fillAWCode322(workbook, worksheet) {
    const patternCode322 = /^\d{3}\.\d{2}\.[A-Za-z0-9]{2}$/;
    const patternCode43  = /^\d{4}\.[A-Za-z0-9]{3}$/;
    const code322Rows = [];
    const code43Rows  = [];
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const val = row.getCell(1).value ? String(row.getCell(1).value).trim() : "";
      if (patternCode322.test(val)) code322Rows.push(row.number);
      if (patternCode43.test(val))  code43Rows.push(row.number);
    });
    console.log(`[Step 15] Code 322 rows:`, code322Rows);
    console.log(`[Step 15] Code 43 rows (from Step 14):`, code43Rows);
    for (let i = 0; i < code322Rows.length; i++) {
      const Y     = code322Rows[i];
      const nextY = code322Rows[i + 1] ?? Infinity;
      const colAW = 49; // AW = col 49
      worksheet.getRow(Y).getCell(colAW).value = "524 SEMULA";
      const step14Cells = code43Rows
        .filter(r => r > Y && r < nextY)
        .map(r => `AW${r + 1}`);
      const formulaStr = step14Cells.length > 0
        ? step14Cells.join("+")
        : "0";
      worksheet.getRow(Y + 1).getCell(colAW).value = { formula: `=${formulaStr}` };
    }
  },

  // ----------------------------------------------------------
  // STEP 16: Pada setiap baris Y dimana kolom A match Code 433
  //          (XXXX.XXX.XXX), isi AX(Y) = "524 MENJADI"
  //          dan AX(Y+1) = formula SUM kolom AR dari baris-baris
  //          dimana kolom X adalah 6 digit diawali "524",
  //          hanya antara Y dan Code 433 berikutnya
  // ----------------------------------------------------------
  async function step16_fillAXCode433(workbook, worksheet) {
    const patternCode433 = /^\d{4}\.[A-Za-z0-9]{3}\.[A-Za-z0-9]{3}$/;
    const pattern524     = /^524\d{3}$/;
    const code433Rows = [];
    const rows524     = [];
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const valA = row.getCell(1).value  ? String(row.getCell(1).value).trim()  : "";
      const valX = row.getCell(24).value ? String(row.getCell(24).value).trim() : "";
      if (patternCode433.test(valA)) code433Rows.push(row.number);
      if (pattern524.test(valX))     rows524.push(row.number);
    });
    console.log(`[Step 16] Code 433 rows (col A):`, code433Rows);
    console.log(`[Step 16] 524xxx rows (col X):`, rows524);
    for (let i = 0; i < code433Rows.length; i++) {
      const Y     = code433Rows[i];
      const nextY = code433Rows[i + 1] ?? Infinity;
      const colAX = 50; // AX = col 50
      worksheet.getRow(Y).getCell(colAX).value = "524 MENJADI";
      const rows524InRange = rows524.filter(r => r > Y && r < nextY);
      const formulaStr = rows524InRange.length > 0
        ? rows524InRange.map(r => `AR${r}`).join("+")
        : "0";
      worksheet.getRow(Y + 1).getCell(colAX).value = { formula: `=${formulaStr}` };
    }
  },

  // ----------------------------------------------------------
  // STEP 17: Sama seperti Step 14, tapi:
  //          - Kolom AW → AX (col 50)
  //          - "524 SEMULA" → "524 MENJADI"
  //          - AW(Y+1) Step 13 → AX(Y+1) Step 16
  // ----------------------------------------------------------
  async function step17_fillAXCode43(workbook, worksheet) {
    const patternCode43  = /^\d{4}\.[A-Za-z0-9]{3}$/;
    const patternCode433 = /^\d{4}\.[A-Za-z0-9]{3}\.[A-Za-z0-9]{3}$/;
    const code43Rows  = [];
    const code433Rows = [];
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const val = row.getCell(1).value ? String(row.getCell(1).value).trim() : "";
      if (patternCode43.test(val))  code43Rows.push(row.number);
      if (patternCode433.test(val)) code433Rows.push(row.number);
    });
    console.log(`[Step 17] Code 43 rows:`, code43Rows);
    console.log(`[Step 17] Code 433 rows:`, code433Rows);
    for (let i = 0; i < code43Rows.length; i++) {
      const Y     = code43Rows[i];
      const nextY = code43Rows[i + 1] ?? Infinity;
      const colAX = 50; // AX = col 50
      worksheet.getRow(Y).getCell(colAX).value = "524 MENJADI";
      const step16Cells = code433Rows
        .filter(r => r > Y && r < nextY)
        .map(r => `AX${r + 1}`);
      const formulaStr = step16Cells.length > 0
        ? step16Cells.join("+")
        : "0";
      worksheet.getRow(Y + 1).getCell(colAX).value = { formula: `=${formulaStr}` };
    }
  },

  // ----------------------------------------------------------
  // STEP 18: Sama seperti Step 15, tapi:
  //          - Kolom AW → AX (col 50)
  //          - "524 SEMULA" → "524 MENJADI"
  //          - AW(Y+1) Step 14 → AX(Y+1) Step 17
  // ----------------------------------------------------------
  async function step18_fillAXCode322(workbook, worksheet) {
    const patternCode322 = /^\d{3}\.\d{2}\.[A-Za-z0-9]{2}$/;
    const patternCode43  = /^\d{4}\.[A-Za-z0-9]{3}$/;
    const code322Rows = [];
    const code43Rows  = [];
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const val = row.getCell(1).value ? String(row.getCell(1).value).trim() : "";
      if (patternCode322.test(val)) code322Rows.push(row.number);
      if (patternCode43.test(val))  code43Rows.push(row.number);
    });
    console.log(`[Step 18] Code 322 rows:`, code322Rows);
    console.log(`[Step 18] Code 43 rows (from Step 17):`, code43Rows);
    for (let i = 0; i < code322Rows.length; i++) {
      const Y     = code322Rows[i];
      const nextY = code322Rows[i + 1] ?? Infinity;
      const colAX = 50; // AX = col 50
      worksheet.getRow(Y).getCell(colAX).value = "524 MENJADI";
      const step17Cells = code43Rows
        .filter(r => r > Y && r < nextY)
        .map(r => `AX${r + 1}`);
      const formulaStr = step17Cells.length > 0
        ? step17Cells.join("+")
        : "0";
      worksheet.getRow(Y + 1).getCell(colAX).value = { formula: `=${formulaStr}` };
    }
  },

  // ----------------------------------------------------------
  // STEP 19: Selisih AX - AW pada kolom AY
  //          Untuk setiap baris Y (Code 433, Code 43, Code 322)
  //          dan Y+1, isi AY = =AX{row} - AW{row}
  // ----------------------------------------------------------
  async function step19_selisihAY(workbook, worksheet) {
    const patternCode433 = /^\d{4}\.[A-Za-z0-9]{3}\.[A-Za-z0-9]{3}$/;
    const patternCode43  = /^\d{4}\.[A-Za-z0-9]{3}$/;
    const patternCode322 = /^\d{3}\.\d{2}\.[A-Za-z0-9]{2}$/;

    const targetRows  = new Set();
    const triggerRows  = new Set(); // baris Y (Code 433, 43, 322)

    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const val = row.getCell(1).value ? String(row.getCell(1).value).trim() : "";
      if (patternCode433.test(val) || patternCode43.test(val) || patternCode322.test(val)) {
        triggerRows.add(row.number);    // Y
        targetRows.add(row.number);     // Y
        targetRows.add(row.number + 1); // Y+1
      }
    });

    console.log(`[Step 19] Target rows for AY:`, [...targetRows].sort((a,b) => a-b));

    const colAY = 51; // AY = col 51

    for (const r of targetRows) {
      const cell = worksheet.getRow(r).getCell(colAY);
      // Baris Y (trigger row) → "SELISIH", baris Y+1 → formula
      if (triggerRows.has(r)) {
        cell.value = "SELISIH";
      } else {
        cell.value = { formula: `=AX${r}-AW${r}` };
      }
    }
  },
];

// ============================================================
// HELPER: jalankan semua step ke setiap worksheet
// ============================================================
async function runPipeline(workbook) {
  const sheets =
    FILE_CONFIG.targetSheetName
      ? [workbook.getWorksheet(FILE_CONFIG.targetSheetName)]
      : workbook.worksheets;

  for (const worksheet of sheets) {
    if (!worksheet) continue;
    for (const step of STEPS) {
      await step(workbook, worksheet);
    }
  }
}

// ============================================================
// HELPER: parse multipart/form-data secara manual (tanpa multer,
// agar kompatibel dengan Vercel serverless)
// ============================================================
function parseMultipart(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => {
      const body = Buffer.concat(chunks);
      const contentType = req.headers["content-type"] || "";
      const boundaryMatch = contentType.match(/boundary=(.+)$/);

      if (!boundaryMatch) return reject(new Error("No multipart boundary found"));

      const boundary = Buffer.from("--" + boundaryMatch[1]);
      const parts = [];
      let start = 0;

      while (start < body.length) {
        const boundaryIdx = body.indexOf(boundary, start);
        if (boundaryIdx === -1) break;

        const partStart = boundaryIdx + boundary.length + 2; // skip \r\n
        const nextBoundary = body.indexOf(boundary, partStart);
        if (nextBoundary === -1) break;

        const partEnd = nextBoundary - 2; // trim \r\n before next boundary
        const part = body.slice(partStart, partEnd);

        // Pisahkan header dan data
        const headerEnd = part.indexOf(Buffer.from("\r\n\r\n"));
        if (headerEnd === -1) { start = nextBoundary; continue; }

        const headerStr = part.slice(0, headerEnd).toString();
        const data = part.slice(headerEnd + 4);

        const nameMatch = headerStr.match(/name="([^"]+)"/);
        const filenameMatch = headerStr.match(/filename="([^"]+)"/);

        if (nameMatch) {
          parts.push({
            name: nameMatch[1],
            filename: filenameMatch ? filenameMatch[1] : null,
            data,
          });
        }

        start = nextBoundary;
      }

      resolve(parts);
    });
    req.on("error", reject);
  });
}

// ============================================================
// HELPER: baca body JSON (untuk input base64)
// ============================================================
function parseJson(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (c) => chunks.push(c));
    req.on("end", () => {
      try { resolve(JSON.parse(Buffer.concat(chunks).toString())); }
      catch (e) { reject(e); }
    });
    req.on("error", reject);
  });
}

// ============================================================
// MAIN HANDLER (Vercel serverless function)
// ============================================================
module.exports = async function handler(req, res) {
  res = patchRes(res); // pastikan .status() dan .json() tersedia (lokal & Vercel)

  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") return res.status(200).end();

  // Health check
  if (req.method === "GET") {
    return res.status(200).json({
      status: "ok",
      message: "Excel API is running",
      endpoints: {
        "POST /": "Process Excel file",
      },
      accepted_formats: [
        "multipart/form-data  → field: 'file' (xlsx binary)",
        "application/json    → body: { base64: '<base64 string>' }",
      ],
    });
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const workbook = new ExcelJS.Workbook();
    const contentType = req.headers["content-type"] || "";

    // ----------------------------------------------------------
    // INPUT MODE A: multipart/form-data (field name = "file")
    // ----------------------------------------------------------
    if (contentType.includes("multipart/form-data")) {
      const parts = await parseMultipart(req);
      const filePart = parts.find((p) => p.name === "file" && p.filename);

      if (!filePart) {
        return res.status(400).json({ error: "Field 'file' tidak ditemukan dalam form-data" });
      }

      const stream = Readable.from(filePart.data);
      await workbook.xlsx.read(stream);
    }

    // ----------------------------------------------------------
    // INPUT MODE B: application/json dengan { base64: "..." }
    // ----------------------------------------------------------
    else if (contentType.includes("application/json")) {
      const body = await parseJson(req);

      if (!body.base64) {
        return res.status(400).json({ error: "Field 'base64' tidak ditemukan dalam JSON body" });
      }

      const buffer = Buffer.from(body.base64, "base64");
      const stream = Readable.from(buffer);
      await workbook.xlsx.read(stream);

      // Opsional: override nama file output
      if (body.outputFileName) {
        FILE_CONFIG.outputFileName = body.outputFileName;
      }
    }

    // ----------------------------------------------------------
    // INPUT tidak dikenali
    // ----------------------------------------------------------
    else {
      return res.status(400).json({
        error: "Content-Type tidak didukung",
        supported: ["multipart/form-data", "application/json"],
      });
    }

    // ----------------------------------------------------------
    // Jalankan semua step processing
    // ----------------------------------------------------------
    await runPipeline(workbook);

    // ----------------------------------------------------------
    // OUTPUT: kembalikan file xlsx
    // ----------------------------------------------------------
    const outputBuffer = await workbook.xlsx.writeBuffer();

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${FILE_CONFIG.outputFileName}"`);
    res.setHeader("Content-Length", outputBuffer.length);
    return res.status(200).end(outputBuffer);

  } catch (err) {
    console.error("[Excel API Error]", err);
    return res.status(500).json({
      error: "Gagal memproses file Excel",
      detail: err.message,
    });
  }
};