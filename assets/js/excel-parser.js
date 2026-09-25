(function () {
  'use strict';

  const cfg = window.APP_CONFIG;
  const rules = window.BusinessRules;

  function getCellDisplay(sheet, address) {
    const cell = sheet[address];
    if (!cell) return '';
    if (cell.w !== undefined) return String(cell.w);
    if (cell.v instanceof Date) return cell.v.toISOString();
    return cell.v === undefined || cell.v === null ? '' : String(cell.v);
  }

  function parseEmployeeIdentity(text, fallbackFileName) {
    const cleaned = String(text || '').trim();
    let name = cleaned;
    let nip = '';

    let match = cleaned.match(/^(.*?)\s+-\s+(\d{8,})\s*$/);
    if (match) {
      name = match[1].trim();
      nip = match[2].trim();
    } else {
      match = cleaned.match(/(\d{8,})\s*$/);
      if (match) {
        nip = match[1];
        name = cleaned.slice(0, match.index).replace(/[\s-]+$/, '').trim();
      }
    }

    if (!name) {
      name = String(fallbackFileName || '')
        .replace(/\.(xlsx?|xls)$/i, '')
        .replace(/^laporan_presensi_/i, '')
        .replace(/_\d{8}_\d{8}$/i, '')
        .replace(/_/g, ' ')
        .trim();
    }

    if (!name) throw new Error('Identitas pegawai tidak ditemukan.');
    return { name, nip };
  }

  function findHeaderRow(displayRows) {
    const limit = Math.min(displayRows.length, 15);
    for (let r = 0; r < limit; r += 1) {
      const row = (displayRows[r] || []).map((value) => String(value || '').trim().toLowerCase());
      if (row.includes('tanggal') && row.includes('masuk') && row.includes('keluar')) return r;
    }
    throw new Error('Header Tanggal/Masuk/Keluar tidak ditemukan.');
  }

  function findHeaderIndex(header, candidates, fallback) {
    for (const candidate of candidates) {
      const index = header.indexOf(candidate);
      if (index >= 0) return index;
    }
    return fallback >= 0 ? fallback : -1;
  }

  function detectColumns(displayRows) {
    const headerRowIndex = findHeaderRow(displayRows);
    const header = (displayRows[headerRowIndex] || []).map((value) => String(value || '').trim().toLowerCase());
    const date = findHeaderIndex(header, ['tanggal'], 2);
    const inTime = findHeaderIndex(header, ['masuk'], 3);
    const outTime = findHeaderIndex(header, ['keluar'], 4);
    let status = findHeaderIndex(header, ['keterangan', 'status'], -1);
    if (status < 0) status = 21;
    return { date, inTime, outTime, status, dataStartRow: headerRowIndex + 1 };
  }

  function safeLocalDate(year, month, day) {
    const date = new Date(year, month - 1, day, 12, 0, 0);
    if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
    return date;
  }

  function parseDate(rawValue, displayValue) {
    if (rawValue instanceof Date && !Number.isNaN(rawValue.getTime())) {
      return new Date(rawValue.getFullYear(), rawValue.getMonth(), rawValue.getDate(), 12, 0, 0);
    }

    if (typeof rawValue === 'number' && Number.isFinite(rawValue) && window.XLSX && XLSX.SSF) {
      const parsed = XLSX.SSF.parse_date_code(rawValue);
      if (parsed) return safeLocalDate(parsed.y, parsed.m, parsed.d);
    }

    const text = String(displayValue || rawValue || '').trim();
    if (!text || text === '-') return null;

    let match = text.match(/^(\d{4})[-\/]([01]?\d)[-\/]([0-3]?\d)$/);
    if (match) return safeLocalDate(Number(match[1]), Number(match[2]), Number(match[3]));

    match = text.match(/^([0-3]?\d)[-\/]([01]?\d)[-\/](\d{4})$/);
    if (match) return safeLocalDate(Number(match[3]), Number(match[2]), Number(match[1]));

    const parsed = new Date(text);
    if (!Number.isNaN(parsed.getTime())) {
      return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate(), 12, 0, 0);
    }
    return null;
  }

  function parseTimeMinutes(rawValue, displayValue) {
    const text = String(displayValue || '').trim();
    if (text && text !== '-') {
      const match = text.match(/^(\d{1,2})[:.](\d{2})(?::\d{2})?$/);
      if (match) {
        const hour = Number(match[1]);
        const minute = Number(match[2]);
        if (hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) return hour * 60 + minute;
      }
    }

    if (rawValue instanceof Date && !Number.isNaN(rawValue.getTime())) {
      return rawValue.getHours() * 60 + rawValue.getMinutes();
    }

    if (typeof rawValue === 'number' && Number.isFinite(rawValue)) {
      const fraction = rawValue - Math.floor(rawValue);
      return Math.round(fraction * 24 * 60) % (24 * 60);
    }

    return null;
  }

  function parseWorkbookRows(workbook) {
    const firstName = workbook.SheetNames[0];
    if (!firstName) throw new Error('Workbook tidak memiliki sheet.');
    const sheet = workbook.Sheets[firstName];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
    const display = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
    return { sheet, raw, display };
  }

  async function parseAttendanceFile(file, period, holidays) {
    if (!window.XLSX) throw new Error('Library parser Excel belum termuat. Muat ulang halaman.');
    const ext = String(file.name || '').split('.').pop().toLowerCase();
    if (!['xlsx', 'xls'].includes(ext)) throw new Error('Format file harus .xlsx atau .xls.');

    const buffer = await file.arrayBuffer();
    let workbook;
    try {
      workbook = XLSX.read(buffer, { type: 'array', cellDates: true, cellNF: true, cellText: true });
    } catch (error) {
      throw new Error(`File Excel tidak dapat dibaca: ${error.message || error}`);
    }

    const { sheet, raw, display } = parseWorkbookRows(workbook);
    if (raw.length < 6) throw new Error('Format file tidak sesuai: data presensi tidak ditemukan.');

    const identityText = getCellDisplay(sheet, 'A2') || (display[1] && display[1][0]) || '';
    const identity = parseEmployeeIdentity(identityText, file.name);
    const columns = detectColumns(display);
    const records = {};
    let rowsInSelectedPeriod = 0;

    for (let r = columns.dataStartRow; r < raw.length; r += 1) {
      const firstCell = String((display[r] && display[r][0]) || '').trim().toUpperCase();
      if (firstCell === 'TOTAL') break;

      const rawRow = raw[r] || [];
      const displayRow = display[r] || [];
      const date = parseDate(rawRow[columns.date], displayRow[columns.date]);
      if (!date) continue;
      if (date.getFullYear() !== period.year || date.getMonth() + 1 !== period.month) continue;
      rowsInSelectedPeriod += 1;

      const inMinutes = parseTimeMinutes(rawRow[columns.inTime], displayRow[columns.inTime]);
      const outMinutes = parseTimeMinutes(rawRow[columns.outTime], displayRow[columns.outTime]);
      const status = String(displayRow[columns.status] || '').trim();
      const overtime = rules.calculateOvertime(date, inMinutes, outMinutes, status, holidays);
      const key = rules.dateKey(date);
      const record = {
        date,
        inMinutes,
        outMinutes,
        status,
        overtimeHours: overtime.hours,
        originalOvertimeHours: overtime.hours,
        normalEndMinutes: overtime.normalEndMinutes,
        adjustedWorkEndMinutes: overtime.adjustedWorkEndMinutes,
        overtimeStartMinutes: overtime.overtimeStartMinutes,
        isHoliday: rules.isHoliday(date, holidays),
        sourceFileName: file.name
      };

      if (!records[key] || record.overtimeHours >= records[key].overtimeHours) records[key] = record;
    }

    if (!rowsInSelectedPeriod) {
      const month = cfg.INDONESIAN_MONTHS[period.month - 1];
      throw new Error(`Periode data tidak sesuai. Tidak ditemukan data ${month} ${period.year}.`);
    }

    return {
      name: identity.name,
      nip: identity.nip,
      sourceFileName: file.name,
      records
    };
  }

  function mergeEmployeeRecords(target, source) {
    Object.keys(source.records || {}).forEach((key) => {
      if (!target.records[key] || source.records[key].overtimeHours >= target.records[key].overtimeHours) {
        target.records[key] = source.records[key];
      }
    });
    if (!target.nip && source.nip) target.nip = source.nip;
    if (!target.name && source.name) target.name = source.name;
  }

  async function parseFiles(files, period, holidays, onProgress) {
    if (typeof holidays === 'function') {
      onProgress = holidays;
      holidays = [];
    }
    holidays = rules.normalizeHolidays(holidays || []);
    const employeesByKey = new Map();
    const results = [];
    const errors = [];
    const fileList = Array.from(files || []);

    for (let index = 0; index < fileList.length; index += 1) {
      const file = fileList[index];
      if (onProgress) onProgress({ index, total: fileList.length, file, stage: 'reading' });
      try {
        const employee = await parseAttendanceFile(file, period, holidays);
        const key = rules.employeeKey(employee);
        if (!employeesByKey.has(key)) employeesByKey.set(key, employee);
        else mergeEmployeeRecords(employeesByKey.get(key), employee);
        results.push({ fileName: file.name, ok: true, employee: employee.name });
      } catch (error) {
        const message = error && error.message ? error.message : String(error);
        results.push({ fileName: file.name, ok: false, error: message });
        errors.push(`${file.name}: ${message}`);
      }
    }

    const employees = [...employeesByKey.values()].sort((a, b) => a.name.localeCompare(b.name, 'id', { sensitivity: 'base' }));
    return { employees, results, errors };
  }

  window.ExcelParser = Object.freeze({
    parseAttendanceFile,
    parseFiles,
    parseEmployeeIdentity,
    detectColumns,
    parseDate,
    parseTimeMinutes
  });
})();
