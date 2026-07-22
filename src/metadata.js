// ============================================================================
// METADATA  — single source of truth for the Project Initiation Portal.
//
// Everything the portal renders (departments, services, project types,
// asset templates, scope fields, option lists) is defined here so the UI
// never hardcodes users, departments, services, or asset types.
//
// This module is served to the frontend via GET /api/metadata and can later
// be replaced by a SharePoint-backed metadata source without touching the UI.
// ============================================================================

// Field types: text | number | textarea | select(options) | date | people
function F(label, key, type, options, meta) {
  var cfg = { label: label, key: key, type: type || 'text', options: null, helper: '', placeholder: '', required: false, default: '' };
  if (Array.isArray(options)) cfg.options = options;
  else if (options && typeof options === 'object' && !Array.isArray(options)) meta = options;
  if (meta && typeof meta === 'object') {
    if (meta.helper) cfg.helper = meta.helper;
    if (meta.placeholder) meta.placeholder && (cfg.placeholder = meta.placeholder);
    if (meta.required) cfg.required = true;
    if (meta.options && !cfg.options) cfg.options = meta.options;
    if (meta.picker) cfg.picker = meta.picker;
    if (meta.default !== undefined) cfg.default = meta.default;
  }
  return cfg;
}

// ---- Option lists derived from the GRC source spreadsheet -----------------
var GRC_AUDIT_TYPES = [
  'Cyber Audit', 'System Audit', 'CSCRF', 'SAQ', 'SAR', 'IS Audit', 'IT/IS Audit',
  'ISO 27001', 'ISO 9001', 'ISO 22301', 'SOC 1', 'SOC 2', 'SOC 2 Type 2',
  'PCI DSS', 'PCI SSS', 'GDPR', 'HIPAA', 'Maturity Assessment', 'Gap Assessment',
  'Surveillance Audit', 'Application Approval', 'EDP Audit', 'Documentation',
  'Assessment', 'Other'
].filter(function (t) { return ['SEBI / Exchange', 'RBI / Bank', 'Enterprise'].indexOf(t) === -1; });
var GRC_ASSESSMENT_TYPES = [
  'Initial Assessment', 'Reassessment', 'Surveillance', 'Renewal/Recertification',
  'Implementation', 'Gap Assessment Only', 'Other'
];
var GRC_REGULATORS = [
  'NSE', 'BSE', 'MCX', 'CDSL', 'NSDL', 'IIBX', 'SEBI', 'RBI', 'NABARD',
  'Third Party', 'NA', 'NCDEX', 'IFCA', 'Primefort', 'General', 'NSEIX',
  'IRDAI', 'E-KYC'
];

// GRC project types. The five explicit rows from the spreadsheet get their
// own curated option sets; the rest reuse the global lists.
var GRC_PROJECT_TYPES = [
  { name: 'SEBI/Exchange', auditTypes: ['Cyber Audit'], assessmentTypes: ['Initial Assessment'], regulators: ['NSE'], deliverables: ['Audit Report', 'Management Letter', 'Remediation Plan'] },
  { name: 'RBI/Bank', auditTypes: ['System Audit'], assessmentTypes: ['Reassessment'], regulators: ['BSE'], deliverables: ['System Audit Report', 'Compliance Certificate'] },
  { name: 'Enterprise', auditTypes: ['CSCRF'], assessmentTypes: ['Surveillance'], regulators: ['MCX'], deliverables: ['CSCRF Assessment', 'Maturity Report'] },
  { name: 'IT-IS Audit', auditTypes: ['SAQ'], assessmentTypes: ['Renewal/Recertification'], regulators: ['CDSL'], deliverables: ['SAQ Report', 'Attestation'] },
  { name: 'Internal', auditTypes: ['SAR'], assessmentTypes: ['Implementation'], regulators: ['NSDL'], deliverables: ['Internal Audit Report'] },
  { name: 'ISO 27001', deliverables: ['Gap Assessment', 'ISO 27001 Report', 'Certification Support'] },
  { name: 'ISO 9001', deliverables: ['Quality Audit Report'] },
  { name: 'ISO 22301', deliverables: ['BCMS Audit Report'] },
  { name: 'SOC 1', deliverables: ['SOC 1 Report'] },
  { name: 'SOC 2', deliverables: ['SOC 2 Report'] },
  { name: 'SOC 2 Type 2', deliverables: ['SOC 2 Type 2 Report'] },
  { name: 'PCI DSS', deliverables: ['RoC', 'AOC'] },
  { name: 'PCI SSS', deliverables: ['PCI SSS Report'] },
  { name: 'GDPR', deliverables: ['GDPR Assessment'] },
  { name: 'HIPAA', deliverables: ['HIPAA Assessment'] },
  { name: 'CSCRF', deliverables: ['CSCRF Assessment'] },
  { name: 'Maturity Assessment', deliverables: ['Maturity Model Report'] },
  { name: 'Gap Assessment', deliverables: ['Gap Report'] },
  { name: 'Surveillance Audit', deliverables: ['Surveillance Report'] },
  { name: 'Application Approval', deliverables: ['Approval Memo'] },
  { name: 'EDP Audit', deliverables: ['EDP Audit Report'] },
  { name: 'Documentation', deliverables: ['Documentation Pack'] },
  { name: 'E-kyc/UDAI', deliverables: ['KYC Report'] },
  { name: 'Pan Approval', deliverables: ['PAN Approval Memo'] },
  { name: 'NFS Audit', deliverables: ['NFS Audit Report'] },
  { name: 'IRDAI', deliverables: ['IRDAI Assessment'] },
  { name: 'Assessment', deliverables: ['Assessment Report'] },
  { name: 'Other', deliverables: [] }
];

// SOC project types with preloaded scope fields (all editable in the UI).
var SOC_PROJECT_TYPES = [
  {
    name: 'SECEON',
    scope: [
      F('Critical Devices', 'criticalDevices', 'number', null, { default: 10 }),
    ]
  },
  {
    name: 'Securonix',
    scope: [
      F('Critical Devices', 'criticalDevices', 'number'),
    ]
  },
  {
    name: 'Gurucul',
    scope: [
      F('Critical Devices', 'criticalDevices', 'number'),
    ]
  },
  {
    name: 'Forensic',
    scope: [
    ]
  },
  {
    name: 'Email Phishing',
    scope: [
    ]
  },
  { name: 'Product Support', scope: [ F('Product', 'product'), F('Support Scope', 'supportScope', 'textarea'), F('Timeline', 'timeline') ] },
  { name: 'Other', scope: [ F('Custom Scope', 'customScope', 'textarea') ] },
  
];

// VAPT Scope of Work (SOW) table.
//
// The VAPT department captures scope as a structured table (one row per scope
// of work). The column definitions below drive the reusable Dynamic SOW Table
// component, so the table is metadata-driven and easy to extend later.
//
// Column semantics:
//   no           — read-only, auto-numbered (1..N) at render & submit time.
//   domainName   — free text (never a fixed option list) so custom assessments
//                  are supported out of the box.
//   quantity     — textarea so a single row can hold multiple quantities
//                  (e.g. "Server - 10\nEndpoint - 10\nFirewall - 10"). Newlines
//                  are preserved exactly as typed.
//   deliverables — free text summary of what is delivered.
//
// Every column except "no" is required. Validation lives in the form and is
// enforced before submit, matching the Power Automate Parse JSON → Select →
// Create HTML Table workflow that consumes the resulting array of rows.
var VAPT_SOW_COLUMNS = [
  { key: 'no', label: 'No.', type: 'auto', readonly: true, autoNumber: true },
  { key: 'domainName', label: 'Domain Name', type: 'text', required: true, placeholder: 'e.g. Internal Network VAPT' },
  { key: 'quantity', label: 'SOW - Quantity', type: 'textarea', required: true, placeholder: 'Server - 10\nEndpoint - 10\nFirewall - 10', preserveNewlines: true },
  { key: 'deliverables', label: 'Deliverables', type: 'text', required: true, placeholder: 'e.g. Internal Network VAPT Report' }
];

// Training services. Single awareness service; delivery mode is captured
// separately (no free-form scope builder for Training).
var TRAINING_SERVICES = [
  { name: 'Cyber Security Awareness Training', mode: ['Virtual', 'Physical']},
  { name: 'Custom', mode: null }
];

// Department-level general info. People fields are Graph-backed pickers.
// `picker.department` pre-filters the people picker to a real Azure AD
// department (discovered in this tenant: Management, Training & Consulting,
// VAPT, SOC, Compliance). Update these to match your directory.
var DEPARTMENT_GENERAL = [
  F('Priority', 'priority', 'select', ['High', 'Medium', 'Low'], { required: true }),
  F('Remarks', 'remarks', 'textarea', null, { required: true })
];

var VAPT_PROJECT_TYPES = ['SEBI', 'IFSCA', 'Enterprise', 'Bank & RBI', 'International', 'White-label', 'Government'];

module.exports = {
  F: F,
  departments: ['VAPT', 'SOC', 'GRC', 'Training'],
  vaptSowColumns: VAPT_SOW_COLUMNS,
  departmentConfig: {
    VAPT: { key: 'VAPT', label: 'VAPT', kind: 'vapt', assessmentTypes: [], sow: { columns: VAPT_SOW_COLUMNS } },
    SOC: { key: 'SOC', label: 'SOC', kind: 'soc', projectTypes: SOC_PROJECT_TYPES },
    GRC: {
      key: 'GRC', label: 'GRC / Compliance', kind: 'grc',
      auditTypes: GRC_AUDIT_TYPES,
      assessmentTypes: GRC_ASSESSMENT_TYPES,
      regulators: GRC_REGULATORS
    },
    Training: { key: 'Training', label: 'Training', kind: 'training', services: TRAINING_SERVICES, allowCustomService: true }
  },
  departmentGeneral: DEPARTMENT_GENERAL,
  peopleDepartments: ['Business Development', 'Project Management', 'Consulting'],
  grcAuditTypes: GRC_AUDIT_TYPES,
  grcAssessmentTypes: GRC_ASSESSMENT_TYPES,
  grcRegulators: GRC_REGULATORS,
  departmentAliases: { Compliance: 'GRC' },
  vaptProjectTypes: VAPT_PROJECT_TYPES
};
