(function () {
  'use strict';

  const cfg = window.TUKIN_CONFIG;
  const rules = window.TukinRules;
  const MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const BORDER_COLOR = 'FF000000';
  const HEADER_FILL = 'FFDDEBF7';
  const MIN_DATA_ROWS = 36;

  function monthLabel(period) {
    return cfg.MONTHS[period.month - 1];
  }

  function recapFileName(period) {
    return `(${monthLabel(period)}) Rekap Potongan Tunjangan Kinerja CPNS ${period.year}.xlsx`;
  }

  function zipFileName(period) {
    return `Tunjangan Kinerja ${monthLabel(period)} ${period.year}.zip`;
  }

  function thinBorder() {
    const side = { style: 'thin', color: { argb: BORDER_COLOR } };
    return { top: side, left: side, bottom: side, right: side };
  }

  function applyHeaderCell(cell, wrapText) {
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF000000' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_FILL } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: Boolean(wrapText) };
    cell.border = thinBorder();
  }

  function applyDataCell(cell, options) {
    const opts = options || {};
    cell.font = opts.font || { name: 'Calibri', size: 11, color: { argb: 'FF000000' } };
    cell.border = thinBorder();
    cell.alignment = {
      horizontal: opts.horizontal,
      vertical: opts.vertical,
      wrapText: Boolean(opts.wrapText)
    };
    if (opts.numFmt) cell.numFmt = opts.numFmt;
    if (opts.fill) cell.fill = opts.fill;
  }

  function configureWorksheet(ws) {
    ws.properties.defaultRowHeight = 15;
    ws.pageSetup.orientation = 'landscape';
    ws.pageSetup.paperSize = 9; // A4
    ws.pageSetup.scale = 55;
    ws.pageSetup.horizontalCentered = true;
    ws.pageSetup.margins = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };

    const widths = {
      A: 6.42578125,
      B: 9.28515625,
      C: 16.140625,
      D: 7.7109375,
      E: 8.43,
      F: 21,
      G: 7.5703125,
      H: 16.5703125,
      I: 20.7109375,
      J: 19.7109375,
      K: 22,
      L: 14.42578125,
      M: 20.42578125,
      N: 41.42578125,
      O: 40.85546875
    };
    Object.entries(widths).forEach(([column, width]) => { ws.getColumn(column).width = width; });

    ws.getRow(1).height = 18.75;
    ws.getRow(2).height = 18.75;
    ws.getRow(9).height = 46.5;
    ws.getRow(10).height = 15.75;
  }

  function buildStaticHeader(ws, period) {
    const range = rules.attendancePeriod(period);

    ws.mergeCells('A1:O1');
    ws.mergeCells('J2:L2');
    ws.mergeCells('A3:B3');
    ws.mergeCells('A5:B5');
    ws.mergeCells('A6:B6');
    ws.mergeCells('I8:K8');

    ws.getCell('A1').value = 'REKAP POTONGAN TUNJANGAN KINERJA';
    ws.getCell('A1').font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FF000000' } };
    ws.getCell('A1').alignment = { horizontal: 'center' };

    ws.getCell('J2').value = cfg.UNIT_WORK;
    ws.getCell('J2').font = { name: 'Calibri', size: 14, color: { argb: 'FF000000' } };
    ws.getCell('J2').alignment = { horizontal: 'center' };

    const labels = [
      ['A3', 'Bulan'],
      ['A4', 'Dasar Absensi'],
      ['A5', 'Unit Kerja'],
      ['A6', 'Unit Organisasi'],
      ['C3', `: ${monthLabel(period)}`],
      ['C4', `: Tanggal 11 Bulan ${cfg.MONTHS[range.start.getMonth()]} s/d 10 Bulan ${cfg.MONTHS[range.end.getMonth()]}`],
      ['C5', `: ${cfg.UNIT_WORK}`],
      ['C6', `: ${cfg.UNIT_ORGANIZATION}`]
    ];
    labels.forEach(([address, value]) => {
      const cell = ws.getCell(address);
      cell.value = value;
      cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF000000' } };
      cell.alignment = { horizontal: 'left' };
    });

    ws.getCell('I8').value = '(Presentase Potongan Berdasarkan SE Sekjen Nomor 13 Tahun 2025)';
    ws.getCell('I8').font = { name: 'Calibri', size: 11, color: { argb: 'FF000000' } };
    ws.getCell('I8').alignment = { horizontal: 'center' };
    ['I8', 'J8', 'K8'].forEach((address) => {
      ws.getCell(address).border = { bottom: { style: 'thin', color: { argb: BORDER_COLOR } } };
    });

    const headers = [
      'No', 'Satker', 'Anak Satker', 'Bulan', 'Tahun', 'NIP', 'SKP', 'TGL',
      '% Potongan \nAbsensi', '% Potongan SKP', '% Potongan Final\n(30% Absen 70% SKP)',
      'Besaran Tukin', 'Potongan Tukin ', 'Nama', 'Unit Kerja, Unit Organisasi'
    ];
    headers.forEach((value, index) => {
      const cell = ws.getCell(9, index + 1);
      cell.value = value;
      applyHeaderCell(cell, index >= 8 && index <= 10);
    });
  }

  function styleDataRow(ws, rowNumber) {
    const styles = [
      { horizontal: 'center' },
      {},
      { horizontal: 'center' },
      { horizontal: 'center' },
      { horizontal: 'center' },
      { horizontal: 'left', numFmt: '@' },
      { vertical: 'middle' },
      { vertical: 'middle', wrapText: true, numFmt: '@' },
      { horizontal: 'center', numFmt: '0.00%' },
      { numFmt: '0%' },
      { numFmt: '0.00%' },
      { vertical: 'middle', numFmt: '#,##0', font: { name: 'Arial', size: 10, color: { argb: 'FF000000' } } },
      { wrapText: true, numFmt: '#,##0', fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } } },
      {},
      {}
    ];
    styles.forEach((style, index) => applyDataCell(ws.getCell(rowNumber, index + 1), style));
  }

  function buildWorkbook(employees, period) {
    if (!window.ExcelJS) throw new Error('Library ExcelJS belum termuat.');

    const wb = new ExcelJS.Workbook();
    wb.creator = 'Pusdatin Kementerian PKP';
    wb.lastModifiedBy = 'Generator Dokumen Pusdatin';
    wb.created = new Date();
    wb.modified = new Date();
    wb.calcProperties.fullCalcOnLoad = true;
    wb.calcProperties.forceFullCalc = true;

    const ws = wb.addWorksheet('PUSDATIN');
    configureWorksheet(ws);
    buildStaticHeader(ws, period);

    const range = rules.attendancePeriod(period);
    const dayAfter = rules.addDays(range.end, 1);
    const sorted = [...employees].sort((a, b) => (a.masterOrder - b.masterOrder) || a.name.localeCompare(b.name, 'id'));
    const rowCount = Math.max(MIN_DATA_ROWS, sorted.length);

    for (let i = 0; i < rowCount; i += 1) {
      const rowNumber = cfg.DATA_START_ROW + i;
      styleDataRow(ws, rowNumber);
      const employee = sorted[i];
      if (!employee) continue;

      const summary = rules.summarizeEmployee(employee);
      const values = [
        i + 1,
        employee.satker || cfg.SATKER,
        String(employee.anakSatker || '').padStart(2, '0'),
        rules.pad2(period.month),
        period.year,
        String(employee.nip || ''),
        cfg.SKP_SCORE,
        rules.formatDate(dayAfter, false),
        summary.attendancePercent / 100,
        summary.skpPercent / 100,
        null,
        summary.tukin,
        null,
        employee.name,
        cfg.UNIT_LABEL
      ];
      values.forEach((value, index) => { ws.getCell(rowNumber, index + 1).value = value; });
      ws.getCell(rowNumber, 11).value = { formula: `(30%*I${rowNumber})+(70%*J${rowNumber})`, result: summary.finalPercent / 100 };
      ws.getCell(rowNumber, 13).value = { formula: `K${rowNumber}*L${rowNumber}`, result: summary.cutAmount };
    }

    return wb;
  }

  async function generateRecap(employees, period) {
    const wb = buildWorkbook(employees, period);
    const buffer = await wb.xlsx.writeBuffer();
    return {
      name: recapFileName(period),
      blob: new Blob([buffer], { type: MIME_XLSX })
    };
  }

  function uniqueName(used, name) {
    const base = String(name || 'file').replace(/[\\/]/g, '-');
    let candidate = base;
    let n = 2;
    while (used.has(candidate.toLowerCase())) {
      const dot = base.lastIndexOf('.');
      candidate = dot > 0 ? `${base.slice(0, dot)} (${n++})${base.slice(dot)}` : `${base} (${n++})`;
    }
    used.add(candidate.toLowerCase());
    return candidate;
  }

  async function generateZip(employees, period) {
    if (!window.JSZip) throw new Error('Library JSZip belum termuat.');
    const recap = await generateRecap(employees, period);
    const zip = new JSZip();
    zip.file(recap.name, recap.blob);
    const root = zip.folder(`PNS ${period.year}`);

    for (const employee of employees) {
      const folder = root.folder(rules.safeFolderName(employee.name));
      const used = new Set();
      for (const source of employee.sourceFiles || []) {
        if (source) folder.file(uniqueName(used, source.name), source);
      }
      for (const record of Object.values(employee.records || {})) {
        for (const evidence of record.evidence || []) {
          if (evidence?.file) folder.file(uniqueName(used, evidence.name || evidence.file.name), evidence.file);
        }
      }
    }

    const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    return { name: zipFileName(period), blob, recap };
  }

  function download(file) {
    saveAs(file.blob, file.name);
  }

  window.TukinGenerator = Object.freeze({ generateRecap, generateZip, download, recapFileName, zipFileName, buildWorkbook });
})();
