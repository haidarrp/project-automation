(function () {
  'use strict';

  const LEAVE_TYPES = Object.freeze([
    'Cuti Tahunan',
    'Cuti Besar',
    'Cuti Sakit',
    'Cuti Melahirkan',
    'Cuti karena Alasan Penting',
    'Cuti Bersama',
    'Cuti di Luar Tanggungan Negara',
    'Lainnya'
  ]);

  // Struktur sengaja dipisahkan agar daftar hari libur/cuti bersama dapat
  // ditambahkan kemudian tanpa mengubah fungsi kalkulasi utama.
  const HOLIDAY_DATES = Object.freeze([]);

  function parseDate(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ''))) return null;
    const date = new Date(`${value}T00:00:00`);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  function toIsoDate(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
  }

  function isWorkingDay(value, holidayDates) {
    const date = value instanceof Date ? new Date(value.getFullYear(), value.getMonth(), value.getDate()) : parseDate(value);
    if (!date) return false;
    const day = date.getDay();
    if (day === 0 || day === 6) return false;
    const holidays = holidayDates instanceof Set ? holidayDates : new Set(holidayDates || HOLIDAY_DATES);
    return !holidays.has(toIsoDate(date));
  }

  function workingDates(startValue, endValue, holidayDates) {
    const start = parseDate(startValue);
    const end = parseDate(endValue);
    if (!start || !end || end < start) return [];
    const holidays = holidayDates instanceof Set ? holidayDates : new Set(holidayDates || HOLIDAY_DATES);
    const dates = [];
    for (const cursor = new Date(start); cursor <= end; cursor.setDate(cursor.getDate() + 1)) {
      if (isWorkingDay(cursor, holidays)) dates.push(toIsoDate(cursor));
    }
    return dates;
  }

  function countWorkingDays(startValue, endValue, holidayDates) {
    return workingDates(startValue, endValue, holidayDates).length;
  }

  function dateRange(startValue, endValue) {
    const start = parseDate(startValue);
    const end = parseDate(endValue);
    if (!start || !end || end < start) return [];
    const dates = [];
    for (const cursor = new Date(start); cursor <= end; cursor.setDate(cursor.getDate() + 1)) dates.push(toIsoDate(cursor));
    return dates;
  }

  window.CutiRules = Object.freeze({
    leaveTypes: LEAVE_TYPES,
    holidayDates: HOLIDAY_DATES,
    parseDate,
    toIsoDate,
    isWorkingDay,
    workingDates,
    countWorkingDays,
    dateRange
  });
})();
