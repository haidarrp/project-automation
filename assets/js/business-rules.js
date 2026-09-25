(function () {
  'use strict';

  const cfg = window.APP_CONFIG;

  function normalizeName(name) {
    return String(name || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
  }

  function pad2(value) {
    return String(value).padStart(2, '0');
  }

  function dateKey(date) {
    return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
  }

  function dateFromKey(key) {
    const [y, m, d] = String(key).split('-').map(Number);
    return new Date(y, m - 1, d, 12, 0, 0);
  }

  function employeeKey(employee) {
    return String(employee.nip || '').replace(/\D/g, '') || normalizeName(employee.name);
  }

  function isWeekend(date) {
    const day = date.getDay();
    return day === 0 || day === 6;
  }

  function normalizeHolidays(holidays) {
    if (!holidays) return [];
    const values = holidays instanceof Set ? [...holidays] : Array.isArray(holidays) ? holidays : [];
    return [...new Set(values.map((value) => String(value || '').trim()).filter((value) => /^\d{4}-\d{2}-\d{2}$/.test(value)))].sort();
  }

  function isHoliday(date, holidays) {
    const key = dateKey(date);
    if (holidays instanceof Set) return holidays.has(key);
    return Array.isArray(holidays) && holidays.includes(key);
  }

  function isWeekendLike(date, holidays) {
    return isWeekend(date) || isHoliday(date, holidays);
  }

  function calculateOvertime(date, inMinutes, outMinutes, status, holidays) {
    const r = cfg.RULES;
    const day = date.getDay();
    const weekend = isWeekendLike(date, holidays);
    const statusUpper = String(status || '').trim().toUpperCase();
    const normalEnd = day === 5 ? r.NORMAL_END_FRIDAY_MINUTES : r.NORMAL_END_MON_THU_MINUTES;

    if (inMinutes === null || outMinutes === null) {
      return {
        hours: 0,
        normalEndMinutes: weekend ? null : normalEnd,
        adjustedWorkEndMinutes: null,
        overtimeStartMinutes: null
      };
    }

    if (weekend && r.ALLOW_WEEKEND_OVERTIME) {
      const elapsedMinutes = outMinutes >= inMinutes
        ? outMinutes - inMinutes
        : (24 * 60 - inMinutes) + outMinutes;
      const rawHours = Math.floor(elapsedMinutes / 60);
      const hours = rawHours < r.MIN_OVERTIME_HOURS
        ? 0
        : Math.min(r.MAX_OVERTIME_HOURS, rawHours);
      return {
        hours,
        normalEndMinutes: null,
        adjustedWorkEndMinutes: null,
        overtimeStartMinutes: hours > 0 ? inMinutes : null
      };
    }

    if (r.REQUIRE_WFO_STATUS && statusUpper !== 'WFO') {
      return {
        hours: 0,
        normalEndMinutes: normalEnd,
        adjustedWorkEndMinutes: null,
        overtimeStartMinutes: null
      };
    }

    if (inMinutes > r.LATEST_OVERTIME_ARRIVAL_MINUTES) {
      return {
        hours: 0,
        normalEndMinutes: normalEnd,
        adjustedWorkEndMinutes: null,
        overtimeStartMinutes: null
      };
    }

    const lateMinutes = Math.max(0, inMinutes - r.NORMAL_START_MINUTES);
    const overtimeStart = normalEnd + lateMinutes;
    const eligibleMinutes = Math.max(0, outMinutes - overtimeStart);
    const rawHours = Math.floor(eligibleMinutes / 60);
    const hours = rawHours < r.MIN_OVERTIME_HOURS
      ? 0
      : Math.min(r.MAX_OVERTIME_HOURS, rawHours);

    return {
      hours,
      normalEndMinutes: normalEnd,
      adjustedWorkEndMinutes: hours > 0 ? overtimeStart : null,
      overtimeStartMinutes: hours > 0 ? overtimeStart : null
    };
  }

  function summarize(employees) {
    let totalHours = 0;
    let mealDays = 0;
    let overtimeEmployees = 0;

    employees.forEach((employee) => {
      let hasOvertime = false;
      Object.values(employee.records || {}).forEach((record) => {
        const hours = Number(record.overtimeHours || 0);
        if (hours > 0) {
          hasOvertime = true;
          totalHours += hours;
          if (hours >= cfg.RULES.MEAL_ALLOWANCE_MIN_HOURS) mealDays += 1;
        }
      });
      if (hasOvertime) overtimeEmployees += 1;
    });

    return {
      employees: employees.length,
      overtimeEmployees,
      totalHours,
      mealDays
    };
  }

  function collectOvertimeRows(employees) {
    const rows = [];
    employees.forEach((employee) => {
      Object.keys(employee.records || {}).sort().forEach((key) => {
        const record = employee.records[key];
        if (!record || Number(record.overtimeHours || 0) <= 0) return;
        rows.push({ employee, record, key });
      });
    });
    return rows;
  }

  function collectOvertimeDates(employees) {
    const map = new Map();
    collectOvertimeRows(employees).forEach(({ record, key }) => map.set(key, record.date));
    return [...map.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([, date]) => date);
  }

  function formatMinutes(minutes) {
    if (minutes === null || minutes === undefined || Number.isNaN(Number(minutes))) return '-';
    const value = Number(minutes);
    const h = Math.floor(value / 60) % 24;
    const m = value % 60;
    return `${pad2(h)}:${pad2(m)}`;
  }

  function formatIndonesianDate(date, includeDay) {
    const dayName = cfg.INDONESIAN_DAYS[date.getDay()];
    const monthName = cfg.INDONESIAN_MONTHS[date.getMonth()];
    const core = `${pad2(date.getDate())} ${monthName} ${date.getFullYear()}`;
    return includeDay ? `${dayName}, ${core}` : core;
  }

  async function sha256(text) {
    const input = new TextEncoder().encode(String(text || ''));
    const digest = await crypto.subtle.digest('SHA-256', input);
    return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, '0')).join('');
  }

  async function assignSatkerCodes(employees) {
    const master = window.SATKER_MASTER_HASHED || [];
    const byNip = new Map(master.map((item) => [item.nipHash, item.code]));
    const byName = new Map(master.map((item) => [item.nameHash, item.code]));
    const unmatched = [];

    for (const employee of employees) {
      const nipDigits = String(employee.nip || '').replace(/\D/g, '');
      const nameNormalized = normalizeName(employee.name);
      let code = '';
      if (nipDigits) {
        const nipHash = await sha256(nipDigits);
        if (byNip.has(nipHash)) code = byNip.get(nipHash);
      }
      if (!code && nameNormalized) {
        const nameHash = await sha256(nameNormalized);
        if (byName.has(nameHash)) code = byName.get(nameHash);
      }
      employee.satkerCode = code;
      if (!code) unmatched.push(employee.name + (employee.nip ? ` (${employee.nip})` : ''));
    }

    return unmatched;
  }

  function serializeEmployees(employees) {
    return employees.map((employee) => ({
      name: employee.name,
      nip: employee.nip,
      satkerCode: employee.satkerCode || '',
      records: Object.fromEntries(Object.entries(employee.records || {}).map(([key, record]) => [key, {
        ...record,
        date: key
      }]))
    }));
  }

  function hydrateEmployees(serialized) {
    return (serialized || []).map((employee) => ({
      ...employee,
      records: Object.fromEntries(Object.entries(employee.records || {}).map(([key, record]) => [key, {
        ...record,
        date: dateFromKey(record.date || key)
      }]))
    }));
  }

  window.BusinessRules = Object.freeze({
    normalizeName,
    pad2,
    dateKey,
    dateFromKey,
    employeeKey,
    isWeekend,
    normalizeHolidays,
    isHoliday,
    isWeekendLike,
    calculateOvertime,
    summarize,
    collectOvertimeRows,
    collectOvertimeDates,
    formatMinutes,
    formatIndonesianDate,
    assignSatkerCodes,
    serializeEmployees,
    hydrateEmployees
  });
})();
