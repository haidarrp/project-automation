window.APP_CONFIG = Object.freeze({
  APP_NAME: 'Generator Dokumen Lembur',
  APP_SUBTITLE: 'Pusat Data dan Informasi',
  MINISTRY: 'PERUMAHAN DAN KAWASAN PERMUKIMAN',
  MINISTRY_DISPLAY: 'Kementerian Perumahan dan Kawasan Permukiman',
  ORGANIZATION_UNIT: 'SEKRETARIAT JENDERAL KEMENTERIAN PERUMAHAN DAN KAWASAN PERMUKIMAN',
  WORK_UNIT: 'PUSAT DATA DAN INFORMASI',
  CITY: 'JAKARTA',
  TIME_ZONE: 'Asia/Jakarta',
  LOCALE: 'id-ID',

  RESPONSIBLE_TITLE_1: 'Pejabat yang Bertanggung Jawab',
  RESPONSIBLE_TITLE_2: 'Kepala Pusat Data dan Informasi',
  RESPONSIBLE_NAME: 'Adhita Surya Permana, S.Si., M.T.',
  RESPONSIBLE_NIP: '197804102002121003',
  SPKL_ADDRESS: 'Wisma Mandiri 2, Jl. Kebon Sirih No.83, RT.2/RW.1, Kb. Sirih, Kec. Menteng, Kota Jakarta Pusat, Daerah Khusus Jakarta, 10340',

  RULES: Object.freeze({
    NORMAL_START_MINUTES: 7 * 60 + 30,
    NORMAL_END_MON_THU_MINUTES: 16 * 60,
    NORMAL_END_FRIDAY_MINUTES: 16 * 60 + 30,
    // Migrated from Database.gs executable configuration.
    // IMPORTANT: Main.gs comments mention 09:31 while Database.gs contains 09:00.
    // Keep this confirmation flag visible until the business owner confirms the intended threshold.
    LATEST_OVERTIME_ARRIVAL_MINUTES: 9 * 60,
    LATEST_OVERTIME_ARRIVAL_CONFIRMATION_REQUIRED: true,
    MIN_OVERTIME_HOURS: 1,
    MAX_OVERTIME_HOURS: 4,
    REQUIRE_WFO_STATUS: true,
    ALLOW_WEEKEND_OVERTIME: true,
    MEAL_ALLOWANCE_MIN_HOURS: 2
  }),

  OUTPUT: Object.freeze({
    RECAP_PREFIX: 'Rekapitulasi Lembur',
    DAILY_PREFIX: 'Daftar Hadir Kerja Lembur',
    SPKL_PREFIX: 'SPKL',
    SPKL_WEEKEND_SUFFIX: ' WEEKEND'
  }),

  INDONESIAN_MONTHS: Object.freeze([
    'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
  ]),
  INDONESIAN_DAYS: Object.freeze([
    'Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'
  ]),

  STORAGE_KEY: 'generator-lembur-pusdatin:v1',
  HISTORY_LIMIT: 12
});
