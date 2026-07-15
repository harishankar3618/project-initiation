const SITE_URL = process.env.SHAREPOINT_SITE_URL || 'https://techdefencelabsolutions.sharepoint.com/sites/pmo';

const LIST_NAMES = {
  clients: ['Clients Track', 'Client Master'],
  intake: ['Project Intake', 'Main Tracker', 'Progress tracker list']
};

const CLIENT_FIELD_CANDIDATES = {
  name: ['LinkTitle', 'ClientName', 'Title', 'Client', 'Name'],
  contactPerson: ['ContactPersonName', 'ContactPerson', 'Contact Person', 'PrimaryContact'],
  contactDesignation: ['ContactPersonDesignation', 'ContactDesignation', 'Designation', 'ContactRole'],
  contactNumber: ['ContactPersonNo_x002e_', 'ContactNumber', 'Contact Number', 'Phone', 'Mobile'],
  contactEmail: ['ContactPersonMail', 'ContactPersonMailId', 'ContactEmail', 'Contact Email', 'Email'],
  bdPerson: ['BDPerson', 'BDPersonLookupId', 'BD Person', 'BusinessDevelopmentPerson'],
  trackingId: ['Trackingid', 'TrackingID', 'Tracking Id', 'Tracking', 'TrackingId'],
  quoteId: ['QuoteId', 'QuoteID', 'Quote Id', 'Quote', 'QuoteId'],
  departments: ['DepartmentsInvolved', 'Departments Involved', 'Department', 'Departments'],
  requirement: ['Service_x002f_ProductRequirements', 'Service_x002f_ProductRequirement', 'ServiceProductRequirement', 'Service / Product Requirement', 'Requirement'],
  industry: ['Industry type', 'Industry_x0020_type', 'Industry_x0020_Type', 'IndustryType', 'Industry', 'ClientIndustry', 'Sector'],
  wonDate: ['WONdate', 'WonDate', 'Won Date', 'ClosedWonDate']
};

const PROJECT_FIELD_CANDIDATES = {
  clientKey: ['TrackingID', 'Tracking Id', 'QuoteID', 'Quote Id', 'ClientName', 'Title'],
  department: ['Department', 'DepartmentName', 'Departments', 'ServiceDepartment'],
  service: ['Service', 'ServiceName', 'Title', 'Service / Product'],
  scope: ['Scope', 'ScopeDetails', 'ProjectScope'],
  departments: ['Departments', 'Department', 'DepartmentName']
};

// Main Tracker write mapping. Keys are logical field names; values are the
// SharePoint column *display names* as shown in the list settings. Internal
// names (e.g. "BD_x0020_Person") are resolved at runtime from these display
// names via getListColumns(), so this stays readable and edit-safe.
const MAIN_TRACKER_FIELDS = {
  title: 'Title',
  department: 'Department',
  services: 'Services',
  scope: 'Scope',
  bdPerson: 'BD Person',              // Person field -> written as <internal>LookupId
  contactName: 'Contact Name',
  contactDesignation: 'Contact Person Designation',
  contactEmail: 'Contact Person Mail Id',
  contactPhone: 'Contact Person No.',
  initiationDate: 'Initiation Date',
  priority: 'Priority',
  remarks: 'Remarks',
  status: 'Status',
  scopeJson: 'Scope json',
  industry: 'Industry'
};

// Default value written to the Main Tracker "Status" choice column on initiate.
// Must exist as a choice (or the column must allow fill-in choices).
const DEFAULT_STATUS = 'Initiated';

module.exports = {
  SITE_URL,
  LIST_NAMES,
  CLIENT_FIELD_CANDIDATES,
  PROJECT_FIELD_CANDIDATES,
  MAIN_TRACKER_FIELDS,
  DEFAULT_STATUS
};