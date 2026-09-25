(function () {
  'use strict';

  const cfg = window.TUKIN_CONFIG;
  const r = cfg.RULES;

  function pad2(value) { return String(value).padStart(2, '0'); }
  function normalizeName(name) { return String(name || '').toLowerCase().replace(/[^a-z0-9]+/g, ''); }
  function dateKey(date) { return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`; }
  function dateFromKey(key) {
    const [y,m,d] = String(key).split('-').map(Number);
    return new Date(y, m - 1, d, 12, 0, 0);
  }
  function addDays(date, days) {
    const out = new Date(date.getFullYear(), date.getMonth(), date.getDate() + days, 12, 0, 0);
    return out;
  }
  function compareDate(a,b) { return dateKey(a).localeCompare(dateKey(b)); }
  function inRange(date, start, end) { return compareDate(date,start) >= 0 && compareDate(date,end) <= 0; }

  function attendancePeriod(tukinPeriod) {
    const start = new Date(tukinPeriod.year, tukinPeriod.month - 3, 11, 12, 0, 0);
    const end = new Date(tukinPeriod.year, tukinPeriod.month - 2, 10, 12, 0, 0);
    return { start, end };
  }

  function formatDate(date, includeDay) {
    const core = `${pad2(date.getDate())} ${cfg.MONTHS[date.getMonth()]} ${date.getFullYear()}`;
    return includeDay ? `${cfg.DAYS[date.getDay()]}, ${core}` : core;
  }

  function formatPeriodRange(range) {
    return `${formatDate(range.start, false)} s.d. ${formatDate(range.end, false)}`;
  }

  function formatMinutes(minutes) {
    if (minutes === null || minutes === undefined || Number.isNaN(Number(minutes))) return '-';
    const v = ((Number(minutes) % 1440) + 1440) % 1440;
    return `${pad2(Math.floor(v / 60))}:${pad2(v % 60)}`;
  }

  function isWeekend(date) { return date.getDay() === 0 || date.getDay() === 6; }
  function isManualHoliday(date, settings) {
    return Array.isArray(settings?.holidays) && settings.holidays.includes(dateKey(date));
  }
  function isSourceHoliday(status) { return /(^|\s)libur($|\s)/i.test(String(status || '').trim()); }
  function isRamadan(date, settings) {
    if (!settings?.ramadanEnabled || !settings.ramadanStart || !settings.ramadanEnd) return false;
    const key = dateKey(date);
    return key >= settings.ramadanStart && key <= settings.ramadanEnd;
  }

  function getSchedule(date, settings) {
    const friday = date.getDay() === 5;
    const ramadan = isRamadan(date, settings);
    const startMinutes = ramadan ? r.RAMADAN_START_MINUTES : r.NORMAL_START_MINUTES;
    const endMinutes = ramadan
      ? (friday ? r.RAMADAN_END_FRIDAY_MINUTES : r.RAMADAN_END_MON_THU_MINUTES)
      : (friday ? r.NORMAL_END_FRIDAY_MINUTES : r.NORMAL_END_MON_THU_MINUTES);
    const flexEndMinutes = startMinutes + r.FLEX_MINUTES;
    return { ramadan, startMinutes, endMinutes, flexEndMinutes };
  }

  function tlFromInTime(inMinutes, schedule) {
    if (inMinutes === null || inMinutes === undefined) {
      return { category: 'TL3', percent: r.TL3_PERCENT, reason: 'Tidak mengisi presensi masuk' };
    }
    if (inMinutes <= schedule.flexEndMinutes) return { category: '-', percent: 0, reason: 'Jam masuk masih dalam batas fleksibilitas' };
    if (inMinutes <= schedule.flexEndMinutes + r.TL1_WINDOW_MINUTES) {
      return { category: 'TL1', percent: r.TL1_PERCENT, reason: `Jam masuk ${formatMinutes(inMinutes)} melewati batas fleksibilitas ${formatMinutes(schedule.flexEndMinutes)}` };
    }
    if (inMinutes <= schedule.flexEndMinutes + r.TL2_WINDOW_MINUTES) {
      return { category: 'TL2', percent: r.TL2_PERCENT, reason: `Jam masuk ${formatMinutes(inMinutes)} masuk kategori keterlambatan tingkat 2` };
    }
    return { category: 'TL3', percent: r.TL3_PERCENT, reason: `Jam masuk ${formatMinutes(inMinutes)} masuk kategori keterlambatan tingkat 3` };
  }

  function pswFromTimes(inMinutes, outMinutes, schedule) {
    if (outMinutes === null || outMinutes === undefined) {
      return { category: 'PSW4', percent: r.PSW4_PERCENT, shortageMinutes: null, requiredEndMinutes: inMinutes === null || inMinutes === undefined ? schedule.endMinutes : schedule.endMinutes + Math.max(0, inMinutes - schedule.startMinutes), reason: 'Tidak mengisi presensi pulang' };
    }

    const requiredEndMinutes = (inMinutes === null || inMinutes === undefined)
      ? schedule.endMinutes
      : schedule.endMinutes + Math.max(0, inMinutes - schedule.startMinutes);
    let actualOut = Number(outMinutes);
    if (inMinutes !== null && inMinutes !== undefined && actualOut < inMinutes) actualOut += 1440;
    const shortageMinutes = Math.max(0, requiredEndMinutes - actualOut);
    if (shortageMinutes <= 0) return { category: '-', percent: 0, shortageMinutes: 0, requiredEndMinutes, reason: `Jam pulang memenuhi kewajiban sampai ${formatMinutes(requiredEndMinutes)}` };
    if (shortageMinutes <= 30) return { category: 'PSW1', percent: r.PSW1_PERCENT, shortageMinutes, requiredEndMinutes, reason: `Kekurangan waktu kerja ${shortageMinutes} menit` };
    if (shortageMinutes <= 60) return { category: 'PSW2', percent: r.PSW2_PERCENT, shortageMinutes, requiredEndMinutes, reason: `Kekurangan waktu kerja ${shortageMinutes} menit` };
    if (shortageMinutes <= 90) return { category: 'PSW3', percent: r.PSW3_PERCENT, shortageMinutes, requiredEndMinutes, reason: `Kekurangan waktu kerja ${shortageMinutes} menit` };
    return { category: 'PSW4', percent: r.PSW4_PERCENT, shortageMinutes, requiredEndMinutes, reason: `Kekurangan waktu kerja ${shortageMinutes} menit` };
  }

  function calculateDay(date, inMinutes, outMinutes, status, settings) {
    const schedule = getSchedule(date, settings);
    const holiday = isWeekend(date) || isManualHoliday(date, settings) || isSourceHoliday(status);
    if (holiday) {
      return {
        schedule, holiday: true, tl: {category:'-',percent:0,reason:'Hari libur'},
        psw: {category:'-',percent:0,shortageMinutes:0,requiredEndMinutes:schedule.endMinutes,reason:'Hari libur'},
        totalPercent: 0,
        reason: 'Hari libur / tanggal merah'
      };
    }
    const tl = tlFromInTime(inMinutes, schedule);
    const psw = pswFromTimes(inMinutes, outMinutes, schedule);
    const totalPercent = Math.min(r.MAX_DAILY_PERCENT, Number((tl.percent + psw.percent).toFixed(2)));
    const reasons = [];
    if (tl.percent > 0) reasons.push(`${tl.category}: ${tl.reason}`);
    if (psw.percent > 0) reasons.push(`${psw.category}: ${psw.reason}`);
    return { schedule, holiday: false, tl, psw, totalPercent, reason: reasons.join(' · ') || 'Tidak ada potongan' };
  }

  function recordFinalPercent(record) {
    return record && record.adjustedPercent !== null && record.adjustedPercent !== undefined
      ? Number(record.adjustedPercent)
      : Number(record?.autoTotalPercent || 0);
  }

  function summarizeEmployee(employee) {
    const records = Object.values(employee.records || {});
    const attendancePercent = Number(records.reduce((sum, rec) => sum + recordFinalPercent(rec), 0).toFixed(2));
    const skpPercent = Number(employee.skpDeductionPercent ?? cfg.DEFAULT_SKP_DEDUCTION);
    const finalPercent = Number((attendancePercent * r.ATTENDANCE_WEIGHT + skpPercent * r.SKP_WEIGHT).toFixed(4));
    const tukin = Number(employee.tukin || cfg.DEFAULT_TUKIN);
    const cutAmount = Math.round(tukin * finalPercent / 100);
    const netAmount = tukin - cutAmount;
    const workDays = records.filter((rec) => !rec.isWeekend && !rec.isHoliday && !isSourceHoliday(rec.status)).length;
    const flaggedRecords = records.filter((rec) => rec.needsVerification).length;
    const adjustedRecords = records.filter((rec) => rec.adjustedPercent !== null && rec.adjustedPercent !== undefined).length;
    return { attendancePercent, skpPercent, finalPercent, tukin, cutAmount, netAmount, workDays, flaggedRecords, adjustedRecords };
  }

  function employeeKey(employee) { return String(employee.nip || '').replace(/\D/g,'') || normalizeName(employee.name); }
  function safeFolderName(name) {
    return String(name || 'Pegawai')
      .replace(/,?\s*(S\.Kom\.?|S\.T\.?|ST\.?|S\.E\.?|S\.Si\.?|S\.Tr\.Kom\.?)$/i, '')
      .replace(/[<>:"/\\|?*]/g, '-')
      .replace(/\s+/g, ' ')
      .trim() || 'Pegawai';
  }
  function normalizeHolidays(values) {
    return [...new Set((values || []).map(String).filter(v => /^\d{4}-\d{2}-\d{2}$/.test(v)))].sort();
  }
  function periodContainsKey(range, key) { return key >= dateKey(range.start) && key <= dateKey(range.end); }

  window.TukinRules = Object.freeze({
    pad2, normalizeName, dateKey, dateFromKey, addDays, inRange, attendancePeriod, formatDate,
    formatPeriodRange, formatMinutes, isWeekend, isManualHoliday, isSourceHoliday, isRamadan,
    getSchedule, calculateDay, recordFinalPercent, summarizeEmployee, employeeKey, safeFolderName,
    normalizeHolidays, periodContainsKey
  });
})();
