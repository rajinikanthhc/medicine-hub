/**
 * ============================================
 * MEDICINE HUB - CONFIGURATION
 * ============================================
 */

const APP = {
  NAME: 'Medicine Hub',
  VERSION: '1.0.0'
};

const SHEETS = {
  MEDICAL_HISTORY: 'Medical History',
  MEDICINES: 'Medicines',
  DOCTORS: 'Doctors',
  MEDICAL_REPORTS: 'Medical Reports',
  FAMILY_MEMBERS: 'Family Members',
  SETTINGS: 'Settings'
};

const STATUS = {
  ACTIVE: 'Active',
  COMPLETED: 'Completed',
  FOLLOW_UP: 'Follow-up Pending',
  CANCELLED: 'Cancelled'
};

const ID_PREFIX = {
  MEDICAL_HISTORY: 'MH',
  MEDICINE: 'MED',
  DOCTOR: 'DOC',
  REPORT: 'RPT',
  FAMILY: 'FM'
};

const DATE_FORMAT = 'dd-MMM-yyyy';
const TIME_ZONE = Session.getScriptTimeZone();