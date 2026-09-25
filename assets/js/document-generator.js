(function () {
  'use strict';

  const cfg = window.APP_CONFIG;
  const rules = window.BusinessRules;

  const thinBorder = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } }
  };

  function assertExcelJS() {
    if (!window.ExcelJS) throw new Error('Library generator Excel belum termuat. Muat ulang halaman.');
  }

  function monthLabel(period) {
    return `${cfg.INDONESIAN_MONTHS[period.month - 1]} ${period.year}`;
  }

  function minutesFraction(minutes) {
    return minutes === null || minutes === undefined ? null : Number(minutes) / (24 * 60);
  }

  function excelDateSerial(date) {
    const utcDate = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
    const epoch = Date.UTC(1899, 11, 30);
    return Math.round((utcDate - epoch) / 86400000);
  }

  function colLetter(column) {
    let result = '';
    let value = column;
    while (value > 0) {
      const remainder = (value - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      value = Math.floor((value - 1) / 26);
    }
    return result;
  }

  function applyBorder(rangeCells) {
    rangeCells.forEach((cell) => { cell.border = thinBorder; });
  }

  function eachCellInRange(ws, r1, c1, r2, c2, callback) {
    for (let row = r1; row <= r2; row += 1) {
      for (let col = c1; col <= c2; col += 1) callback(ws.getCell(row, col));
    }
  }

  function styleRange(ws, r1, c1, r2, c2, styleFn) {
    eachCellInRange(ws, r1, c1, r2, c2, styleFn);
  }

  function setBaseFont(ws, maxRow, maxCol) {
    styleRange(ws, 1, 1, maxRow, maxCol, (cell) => {
      cell.font = { name: 'Arial', size: 10 };
      cell.alignment = { vertical: 'middle' };
    });
  }

  function setBorders(ws, r1, c1, r2, c2) {
    styleRange(ws, r1, c1, r2, c2, (cell) => { cell.border = thinBorder; });
  }

  function setFill(cell, argb) {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
  }

  function setColumnWidthsRecap(ws) {
    const px = [35, 24, 40, 72, 255, ...Array(31).fill(31), 70, 18, 95, 105];
    px.forEach((value, index) => { ws.getColumn(index + 1).width = Math.max(3, value / 7); });
  }

  function buildRecapSheet(ws, employees, period, holidays) {
    const firstEmployeeRow = 8;
    const lastEmployeeRow = firstEmployeeRow + employees.length - 1;
    const totalRow = lastEmployeeRow + 1;
    const daysInMonth = new Date(period.year, period.month, 0).getDate();

    ws.name = 'Rekap';
    ws.views = [{ showGridLines: false, state: 'frozen', ySplit: 7 }];
    setColumnWidthsRecap(ws);
    setBaseFont(ws, Math.max(80, totalRow + 5), 40);

    ws.mergeCells('C2:AK2');
    ws.mergeCells('C3:AK3');
    ws.mergeCells('C4:AK4');
    ws.getCell('C2').value = 'REKAPITULASI LEMBUR';
    ws.getCell('C3').value = cfg.WORK_UNIT;
    ws.getCell('C4').value = `BULAN ${cfg.INDONESIAN_MONTHS[period.month - 1].toUpperCase()} ${period.year}`;
    ['C2', 'C3', 'C4'].forEach((a1) => {
      const cell = ws.getCell(a1);
      cell.font = { name: 'Arial', size: 11, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    ws.getCell('AM3').value = 'Bulan';
    ws.getCell('AN3').value = period.month;
    ws.getCell('AM3').alignment = ws.getCell('AN3').alignment = { horizontal: 'center', vertical: 'middle' };

    ['C6:C7', 'D6:D7', 'E6:E7', 'F6:AJ6', 'AK6:AK7', 'AL6:AL7', 'AM6:AM7', 'AN6:AN7'].forEach((range) => ws.mergeCells(range));
    ws.getCell('C6').value = 'NO';
    ws.getCell('D6').value = 'Kode Satker';
    ws.getCell('E6').value = 'NAMA';
    ws.getCell('F6').value = 'TANGGAL KEHADIRAN';
    ws.getCell('AK6').value = 'Total (Jam)';
    ws.getCell('AM6').value = 'Jumlah Hari Lembur';
    ws.getCell('AN6').value = 'Uang Makan (Hari)';

    for (let day = 1; day <= 31; day += 1) ws.getCell(7, 5 + day).value = day <= daysInMonth ? day : '';
    styleRange(ws, 6, 3, 7, 40, (cell) => {
      cell.font = { name: 'Arial', size: 10, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = thinBorder;
    });

    employees.forEach((employee, index) => {
      const row = firstEmployeeRow + index;
      ws.getCell(row, 3).value = index + 1;
      ws.getCell(row, 4).value = employee.satkerCode || '';
      ws.getCell(row, 5).value = employee.name + (employee.nip ? `\n${employee.nip}` : '');
      ws.getCell(row, 5).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      ws.getRow(row).height = 31.5;

      for (let day = 1; day <= 31; day += 1) {
        const key = `${period.year}-${rules.pad2(period.month)}-${rules.pad2(day)}`;
        const record = employee.records[key];
        ws.getCell(row, 5 + day).value = record && record.overtimeHours > 0 ? Number(record.overtimeHours) : null;
        ws.getCell(row, 5 + day).alignment = { horizontal: 'center', vertical: 'middle' };
      }
      const overtimeRecords = Object.values(employee.records || {}).filter((record) => Number(record.overtimeHours || 0) > 0);
      ws.getCell(row, 37).value = overtimeRecords.reduce((sum, record) => sum + Number(record.overtimeHours || 0), 0);
      ws.getCell(row, 39).value = overtimeRecords.length;
      ws.getCell(row, 40).value = overtimeRecords.filter((record) => Number(record.overtimeHours || 0) >= cfg.RULES.MEAL_ALLOWANCE_MIN_HOURS).length;
      [3, 4, 37, 39, 40].forEach((col) => { ws.getCell(row, col).alignment = { horizontal: 'center', vertical: 'middle' }; });
      setBorders(ws, row, 3, row, 40);
    });

    for (let day = 1; day <= daysInMonth; day += 1) {
      const date = new Date(period.year, period.month - 1, day, 12, 0, 0);
      if (rules.isWeekendLike(date, holidays)) {
        const column = 5 + day;
        for (let row = 7; row <= lastEmployeeRow; row += 1) setFill(ws.getCell(row, column), 'FFE7E6E6');
      }
    }

    ws.getCell(totalRow, 5).value = 'TOTAL';
    for (let col = 6; col <= 36; col += 1) {
      let total = 0;
      for (let row = firstEmployeeRow; row <= lastEmployeeRow; row += 1) total += Number(ws.getCell(row, col).value || 0);
      ws.getCell(totalRow, col).value = total || null;
    }
    ws.getCell(totalRow, 37).value = employees.reduce((sum, employee) => sum + Object.values(employee.records || {}).reduce((sub, record) => sub + Number(record.overtimeHours || 0), 0), 0);
    ws.getCell(totalRow, 39).value = employees.reduce((sum, employee) => sum + Object.values(employee.records || {}).filter((record) => Number(record.overtimeHours || 0) > 0).length, 0);
    ws.getCell(totalRow, 40).value = employees.reduce((sum, employee) => sum + Object.values(employee.records || {}).filter((record) => Number(record.overtimeHours || 0) >= cfg.RULES.MEAL_ALLOWANCE_MIN_HOURS).length, 0);
    styleRange(ws, totalRow, 3, totalRow, 40, (cell) => {
      cell.font = { name: 'Arial', size: 10, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = thinBorder;
      setFill(cell, 'FFD9EAD3');
    });
  }

  function activeDailyRecords(employees) {
    const byDate = new Map();
    employees.forEach((employee) => {
      Object.keys(employee.records || {}).sort().forEach((key) => {
        const record = employee.records[key];
        if (!record || Number(record.overtimeHours || 0) <= 0) return;
        if (!byDate.has(key)) byDate.set(key, []);
        byDate.get(key).push({ employee, record });
      });
    });
    byDate.forEach((rows) => rows.sort((a, b) => a.employee.name.localeCompare(b.employee.name, 'id', { sensitivity: 'base' })));
    return byDate;
  }

  function buildDailyDatabase(ws, employees) {
    const byDate = activeDailyRecords(employees);
    ws.state = 'veryHidden';
    ws.views = [{ showGridLines: false }];
    const headers = [
      'Lookup Key', 'Tanggal', 'Employee Key', 'No', 'Nama Pegawai', 'NIP',
      'Jam Datang Kerja', 'Jam Pulang Kerja Disesuaikan', 'Jam Datang Lembur', 'Jam Pulang Aktual',
      'Jam Lembur', 'Status', 'Sumber File', 'Tanggal Selector', 'Label Tanggal Indonesia',
      'Kunci Urutan Aktif', 'Nama View Harian', 'Jam Datang Kerja View', 'Jam Pulang Kerja View',
      'Jam Datang Lembur View', 'Jam Pulang Lembur View', 'Jam Lembur View', 'Employee Key View'
    ];
    ws.addRow(headers);
    ws.getRow(1).font = { name: 'Arial', size: 10, bold: true };
    ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D2E9' } };

    [...byDate.entries()].sort((a, b) => a[0].localeCompare(b[0])).forEach(([key, rows]) => {
      rows.forEach(({ employee, record }, index) => {
        const sequence = index + 1;
        const lookupKey = excelDateSerial(record.date) * 1000 + sequence;
        ws.addRow([
          `${key}|${rules.employeeKey(employee)}`,
          record.date,
          rules.employeeKey(employee),
          sequence,
          employee.name,
          employee.nip,
          minutesFraction(record.inMinutes),
          minutesFraction(record.adjustedWorkEndMinutes),
          minutesFraction(record.overtimeStartMinutes),
          minutesFraction(record.outMinutes),
          Number(record.overtimeHours || 0),
          record.status,
          record.sourceFileName || '',
          null,
          null,
          lookupKey,
          employee.name,
          minutesFraction(record.inMinutes),
          minutesFraction(record.adjustedWorkEndMinutes),
          minutesFraction(record.overtimeStartMinutes),
          minutesFraction(record.outMinutes),
          Number(record.overtimeHours || 0),
          rules.employeeKey(employee)
        ]);
      });
    });

    for (let row = 2; row <= ws.rowCount; row += 1) {
      ws.getCell(row, 2).numFmt = 'dd mmmm yyyy';
      [7, 8, 9, 10, 18, 19, 20, 21].forEach((col) => { ws.getCell(row, col).numFmt = 'hh:mm'; });
    }
  }

  function buildDailySheet(ws, employees, period, databaseSheet) {
    const firstDataRow = 17;
    const dataRows = Math.max(employees.length, 1);
    const lastDataRow = firstDataRow + dataRows - 1;
    const signatureStartRow = Math.max(68, lastDataRow + 3);
    const daysInMonth = new Date(period.year, period.month, 0).getDate();
    const dateOptionsEndRow = daysInMonth + 1;
    const overtimeDates = rules.collectOvertimeDates(employees);
    const initialDate = overtimeDates[0] || new Date(period.year, period.month - 1, 1, 12, 0, 0);

    ws.name = 'Daftar Hadir Harian';
    ws.views = [{ showGridLines: false, state: 'frozen', ySplit: 16 }];
    ws.properties.defaultRowHeight = 18;
    const px = [36, 220, 78, 24, 68, 78, 68, 78, 68, 78, 68, 82, 20, 20, 20];
    px.forEach((value, index) => { ws.getColumn(index + 1).width = Math.max(3, value / 7); });
    setBaseFont(ws, Math.max(90, signatureStartRow + 10), 15);

    ws.getCell('A1').value = 'KEMENTERIAN';
    ws.getCell('C1').value = `:   ${cfg.MINISTRY.toUpperCase()}`;
    ws.getCell('A2').value = 'UNIT ORGANISASI';
    ws.getCell('C2').value = `:   ${cfg.ORGANIZATION_UNIT}`;
    ws.getCell('A3').value = 'UNIT KERJA';
    ws.getCell('C3').value = `:   ${cfg.WORK_UNIT}`;
    ws.getCell('A4').value = 'DI';
    ws.getCell('C4').value = `:   ${cfg.CITY}`;

    ws.mergeCells('A7:L7');
    ws.getCell('A7').value = 'DAFTAR HADIR KERJA LEMBUR';
    ws.getCell('A7').font = { name: 'Arial', size: 10, bold: true, underline: true };
    ws.getCell('A7').alignment = { horizontal: 'center', vertical: 'middle' };

    ws.getCell('C10').value = 'Hari';
    ws.getCell('D10').value = ':';
    ws.getCell('E10').value = { formula: '=IF($M$11="","-",CHOOSE(WEEKDAY($M$11),"Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"))' };
    ws.getCell('C11').value = 'Tanggal';
    ws.getCell('D11').value = ':';
    ws.mergeCells('E11:F11');
    ws.getCell('E11').value = rules.formatIndonesianDate(initialDate, false);
    ws.getCell('E11').alignment = { horizontal: 'left', vertical: 'middle' };

    for (let day = 1; day <= daysInMonth; day += 1) {
      const row = day + 1;
      const date = new Date(period.year, period.month - 1, day, 12, 0, 0);
      ws.getCell(row, 14).value = rules.formatIndonesianDate(date, false);
      ws.getCell(row, 15).value = date;
      ws.getCell(row, 15).numFmt = 'dd mmmm yyyy';
    }
    ws.getCell('E11').dataValidation = {
      type: 'list',
      allowBlank: false,
      formulae: [`$N$2:$N$${dateOptionsEndRow}`],
      showErrorMessage: true,
      errorTitle: 'Tanggal tidak valid',
      error: 'Pilih tanggal dari daftar.'
    };
    ws.getCell('M11').value = { formula: `=IFERROR(VLOOKUP($E$11,$N$2:$O$${dateOptionsEndRow},2,FALSE),"")` };
    ws.getCell('M11').numFmt = 'dd mmmm yyyy';

    ['A14:A16', 'B14:B16', 'C14:G14', 'H14:K14', 'L14:L16', 'C15:C16', 'E15:E16', 'F15:F16', 'G15:G16', 'H15:H16', 'I15:I16', 'J15:J16', 'K15:K16'].forEach((range) => ws.mergeCells(range));
    ws.getCell('A14').value = 'NO';
    ws.getCell('B14').value = 'NAMA PEGAWAI';
    ws.getCell('C14').value = 'DAFTAR HADIR KERJA';
    ws.getCell('H14').value = 'DAFTAR HADIR LEMBUR';
    ws.getCell('L14').value = 'KET';
    ws.getCell('C15').value = 'JAM DATANG';
    ws.getCell('E15').value = 'PARAF';
    ws.getCell('F15').value = 'JAM PULANG';
    ws.getCell('G15').value = 'PARAF';
    ws.getCell('H15').value = 'JAM DATANG';
    ws.getCell('I15').value = 'PARAF';
    ws.getCell('J15').value = 'JAM PULANG';
    ws.getCell('K15').value = 'PARAF';
    styleRange(ws, 14, 1, 16, 12, (cell) => {
      cell.font = { name: 'Arial', size: 10, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = thinBorder;
    });

    const dbLastRow = Math.max(2, databaseSheet.rowCount);
    for (let index = 0; index < dataRows; index += 1) {
      const row = firstDataRow + index;
      const sequence = index + 1;
      const lookup = `INT($M$11)*1000+${sequence}`;
      ws.getCell(row, 1).value = { formula: `=IF(B${row}="","",${sequence})` };
      ws.getCell(row, 2).value = { formula: `=IF($M$11="","",IFERROR(VLOOKUP(${lookup},'Database'!$P$2:$W$${dbLastRow},2,FALSE),""))` };
      ws.getCell(row, 3).value = { formula: `=IF($M$11="","",IFERROR(VLOOKUP(${lookup},'Database'!$P$2:$W$${dbLastRow},3,FALSE),""))` };
      ws.getCell(row, 6).value = { formula: `=IF($M$11="","",IFERROR(VLOOKUP(${lookup},'Database'!$P$2:$W$${dbLastRow},4,FALSE),""))` };
      ws.getCell(row, 8).value = { formula: `=IF($M$11="","",IFERROR(VLOOKUP(${lookup},'Database'!$P$2:$W$${dbLastRow},5,FALSE),""))` };
      ws.getCell(row, 10).value = { formula: `=IF($M$11="","",IFERROR(VLOOKUP(${lookup},'Database'!$P$2:$W$${dbLastRow},6,FALSE),""))` };
      ws.getCell(row, 12).value = { formula: `=IF($M$11="","",IFERROR(VLOOKUP(${lookup},'Database'!$P$2:$W$${dbLastRow},7,FALSE),""))` };
      ws.getCell(row, 13).value = { formula: `=IF($M$11="","",IFERROR(VLOOKUP(${lookup},'Database'!$P$2:$W$${dbLastRow},8,FALSE),""))` };
      [3, 6, 8, 10].forEach((col) => { ws.getCell(row, col).numFmt = 'hh:mm'; });
      ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      styleRange(ws, row, 1, row, 12, (cell) => {
        cell.border = thinBorder;
        if (cell.column !== 2) cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });
      ws.getRow(row).height = 30;
    }

    ws.getCell(signatureStartRow, 8).value = cfg.RESPONSIBLE_TITLE_1;
    ws.getCell(signatureStartRow + 1, 8).value = cfg.RESPONSIBLE_TITLE_2;
    ws.getCell(signatureStartRow + 7, 8).value = cfg.RESPONSIBLE_NAME;
    ws.getCell(signatureStartRow + 7, 8).font = { name: 'Arial', size: 10, underline: true };
    ws.getCell(signatureStartRow + 8, 8).value = `NIP. ${cfg.RESPONSIBLE_NIP}`;
    for (let row = signatureStartRow; row <= signatureStartRow + 8; row += 1) {
      for (let col = 8; col <= 11; col += 1) ws.getCell(row, col).alignment = { horizontal: 'center', vertical: 'middle' };
    }

    [13, 14, 15].forEach((col) => { ws.getColumn(col).hidden = true; });
  }

  function buildSpklLetterhead(ws) {
    ws.mergeCells('A1:E1');
    ws.mergeCells('A2:E2');
    ws.mergeCells('A3:E3');
    ws.mergeCells('A4:E4');
    ws.getCell('A1').value = `KEMENTERIAN ${cfg.MINISTRY}`;
    ws.getCell('A2').value = 'SEKRETARIAT JENDERAL';
    ws.getCell('A3').value = cfg.WORK_UNIT;
    ws.getCell('A4').value = cfg.SPKL_ADDRESS;
    ['A1', 'A2', 'A3'].forEach((a1, index) => {
      ws.getCell(a1).font = { name: 'Arial', size: index === 0 ? 15 : 14, bold: true };
      ws.getCell(a1).alignment = { horizontal: 'center', vertical: 'middle' };
    });
    ws.getCell('A4').font = { name: 'Arial', size: 9 };
    ws.getCell('A4').alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    for (let col = 1; col <= 5; col += 1) ws.getCell(6, col).border = { bottom: { style: 'medium', color: { argb: 'FF000000' } } };
  }

  function collectSpklGroups(employees, weekendOnly, holidays) {
    const groups = [];
    employees.forEach((employee) => {
      const records = Object.keys(employee.records || {}).sort().map((key) => employee.records[key]).filter((record) => {
        if (!record || Number(record.overtimeHours || 0) <= 0) return false;
        return weekendOnly ? rules.isWeekendLike(record.date, holidays) : !rules.isWeekendLike(record.date, holidays);
      });
      if (records.length) groups.push({ employee, records });
    });
    return groups;
  }

  function buildSpklSheet(ws, employees, period, weekendOnly, holidays) {
    const groups = collectSpklGroups(employees, weekendOnly, holidays);
    const dataStartRow = 14;
    const dataRowCount = groups.reduce((sum, group) => sum + group.records.length, 0);
    const monthUpper = cfg.INDONESIAN_MONTHS[period.month - 1].toUpperCase();
    ws.name = weekendOnly ? `SPKL ${monthUpper} WEEKEND`.slice(0, 31) : `SPKL ${monthUpper}`.slice(0, 31);
    ws.views = [{ showGridLines: false, state: 'frozen', ySplit: 13 }];
    [48, 315, 195, 95, 440].forEach((px, index) => { ws.getColumn(index + 1).width = Math.max(5, px / 7); });
    setBaseFont(ws, Math.max(40, dataStartRow + dataRowCount + 5), 5);
    buildSpklLetterhead(ws);

    ws.mergeCells('A8:E8');
    ws.mergeCells('A9:E9');
    ws.mergeCells('A10:E11');
    ws.getCell('A8').value = 'SURAT PERINTAH KERJA LEMBUR';
    ws.getCell('A9').value = 'NOMOR:';
    ws.getCell('A10').value = `Sehubungan dengan adanya kegiatan lembur yang akan dilaksanakan di lingkungan Pusat Data dan Informasi Sekretariat Jenderal Kementerian Perumahan dan Kawasan Permukiman pada Bulan ${cfg.INDONESIAN_MONTHS[period.month - 1]} ${period.year}, dengan ini kami memerintahkan pegawai berikut:`;
    ['A8', 'A9'].forEach((a1) => { ws.getCell(a1).alignment = { horizontal: 'center', vertical: 'top' }; });
    ws.getCell('A10').alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    ['No.', 'Nama/Bidang', 'Waktu Penugasan (tgl)', 'Jumlah Lembur (jam)', 'Uraian Kegiatan'].forEach((value, index) => { ws.getCell(13, index + 1).value = value; });
    styleRange(ws, 13, 1, 13, 5, (cell) => {
      cell.font = { name: 'Arial', size: 10, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = thinBorder;
    });

    if (!dataRowCount) {
      setBorders(ws, dataStartRow, 1, dataStartRow, 5);
      return;
    }

    let currentRow = dataStartRow;
    groups.forEach((group, groupIndex) => {
      const start = currentRow;
      group.records.forEach((record) => {
        ws.getCell(currentRow, 3).value = rules.formatIndonesianDate(record.date, true);
        ws.getCell(currentRow, 4).value = Number(record.overtimeHours || 0);
        ws.getCell(currentRow, 3).numFmt = '@';
        ws.getCell(currentRow, 4).numFmt = '0';
        styleRange(ws, currentRow, 1, currentRow, 5, (cell) => {
          cell.border = thinBorder;
          cell.alignment = { horizontal: cell.column === 5 ? 'left' : 'center', vertical: 'middle', wrapText: true };
        });
        ws.getRow(currentRow).height = 22.5;
        currentRow += 1;
      });
      const end = currentRow - 1;
      if (end > start) {
        ws.mergeCells(start, 1, end, 1);
        ws.mergeCells(start, 2, end, 2);
      }
      ws.getCell(start, 1).value = groupIndex + 1;
      ws.getCell(start, 2).value = group.employee.name;
      ws.getCell(start, 1).alignment = { horizontal: 'center', vertical: 'middle' };
      ws.getCell(start, 2).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
  }

  function newWorkbook() {
    assertExcelJS();
    const workbook = new ExcelJS.Workbook();
    workbook.creator = cfg.APP_NAME;
    workbook.company = cfg.MINISTRY_DISPLAY;
    workbook.created = new Date();
    workbook.modified = new Date();
    workbook.calcProperties.fullCalcOnLoad = true;
    workbook.calcProperties.forceFullCalc = true;
    workbook.calcProperties.calcMode = 'auto';
    return workbook;
  }

  async function workbookToFile(workbook, fileName) {
    const buffer = await workbook.xlsx.writeBuffer();
    return {
      name: fileName,
      blob: new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    };
  }

  async function generateRecap(employees, period, holidays) {
    const recapEmployees = (employees || []).filter((employee) => {
      const totalOvertime = Object.values(employee.records || {}).reduce(
        (sum, record) => sum + Number(record.overtimeHours || 0),
        0
      );
      return totalOvertime > 0;
    });

    const workbook = newWorkbook();
    buildRecapSheet(workbook.addWorksheet('Rekap'), recapEmployees, period, holidays);
    return workbookToFile(workbook, `${cfg.OUTPUT.RECAP_PREFIX} - ${monthLabel(period)}.xlsx`);
  }

  async function generateDaily(employees, period, holidays) {
    const workbook = newWorkbook();
    const daily = workbook.addWorksheet('Daftar Hadir Harian');
    const database = workbook.addWorksheet('Database');
    buildDailyDatabase(database, employees);
    buildDailySheet(daily, employees, period, database);
    return workbookToFile(workbook, `${cfg.OUTPUT.DAILY_PREFIX} - ${monthLabel(period)}.xlsx`);
  }

  async function generateSpkl(employees, period, holidays) {
    const workbook = newWorkbook();
    const weekday = workbook.addWorksheet('SPKL');
    const weekend = workbook.addWorksheet('SPKL WEEKEND');
    buildSpklSheet(weekday, employees, period, false, holidays);
    buildSpklSheet(weekend, employees, period, true, holidays);
    return workbookToFile(workbook, `${cfg.OUTPUT.SPKL_PREFIX} ${monthLabel(period)}.xlsx`);
  }

  async function generateAll(employees, period, holidays) {
    const [recap, daily, spkl] = await Promise.all([
      generateRecap(employees, period, holidays),
      generateDaily(employees, period, holidays),
      generateSpkl(employees, period, holidays)
    ]);
    return { recap, daily, spkl };
  }

  function downloadFile(file) {
    if (!file || !file.blob) return;
    if (window.saveAs) window.saveAs(file.blob, file.name);
    else {
      const url = URL.createObjectURL(file.blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = file.name;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    }
  }

  window.DocumentGenerator = Object.freeze({
    generateRecap,
    generateDaily,
    generateSpkl,
    generateAll,
    downloadFile
  });
})();
