const { SITE_URL, LIST_NAMES, CLIENT_FIELD_CANDIDATES, PROJECT_FIELD_CANDIDATES, MAIN_TRACKER_FIELDS, DEFAULT_STATUS } = require('./config');
const { graphGet, graphGetAll, graphPost, graphPatch, mapGraphUser } = require('./graphClient');
const metadata = require('./metadata');

function normalizeText(value) {
  return String(value == null ? '' : value).trim();
}

function toIsoDateTime(value) {
  if (value === undefined || value === null || value === '') return '';
  const d = new Date(value);
  if (isNaN(d.getTime())) return '';
  return d.toISOString();
}

function safeField(fields, candidates) {
  for (let i = 0; i < candidates.length; i += 1) {
    const key = candidates[i];
    if (fields && fields[key] !== undefined && fields[key] !== null && String(fields[key]).trim() !== '') {
      return fields[key];
    }
  }
  return '';
}

function splitList(value) {
  return normalizeText(value)
    .split(/[;,\n]/)
    .map(function (part) { return part.trim(); })
    .filter(Boolean);
}

function dedupe(values) {
  return Array.from(new Set(values.filter(Boolean)));
}

function parseClientItem(fields, itemId, columnMap) {
  return {
    id: String(itemId),
    clientName: normalizeText(safeField(fields, resolveCandidates(CLIENT_FIELD_CANDIDATES.name, columnMap))),
    contactPerson: normalizeText(safeField(fields, resolveCandidates(CLIENT_FIELD_CANDIDATES.contactPerson, columnMap))),
    designation: normalizeText(safeField(fields, resolveCandidates(CLIENT_FIELD_CANDIDATES.contactDesignation, columnMap))),
    phone: normalizeText(safeField(fields, resolveCandidates(CLIENT_FIELD_CANDIDATES.contactNumber, columnMap))),
    email: normalizeText(safeField(fields, resolveCandidates(CLIENT_FIELD_CANDIDATES.contactEmail, columnMap))),
    bdPersonLookupId: normalizeText(safeField(fields, resolveCandidates(['BDPersonLookupId'], columnMap))),
    bdPersonName: normalizeText(safeField(fields, resolveCandidates(['BDPerson'], columnMap))),
    bdPerson: null,
    trackingId: normalizeText(safeField(fields, resolveCandidates(CLIENT_FIELD_CANDIDATES.trackingId, columnMap))),
    quoteId: normalizeText(safeField(fields, resolveCandidates(CLIENT_FIELD_CANDIDATES.quoteId, columnMap))),
    departmentsInvolved: splitList(safeField(fields, resolveCandidates(CLIENT_FIELD_CANDIDATES.departments, columnMap))),
    requirement: normalizeText(safeField(fields, resolveCandidates(CLIENT_FIELD_CANDIDATES.requirement, columnMap))),
    industry: normalizeText(safeField(fields, resolveCandidates(CLIENT_FIELD_CANDIDATES.industry, columnMap))),
    wonDate: normalizeText(safeField(fields, resolveCandidates(CLIENT_FIELD_CANDIDATES.wonDate, columnMap)))
  };
}

async function resolveSharePointPerson(lookupId, displayName, siteId) {
  if (!lookupId) {
    return displayName ? { id: '', name: displayName, email: '', claims: '' } : null;
  }
  try {
    const item = await graphGet(
      '/sites/' + siteId + "/lists('User Information List')/items/" + lookupId +
      '?$expand=fields($select=Title,EMail,Name,UserName)'
    );
    const f = item.fields || {};
    const email = normalizeText(f.EMail);
    const name = normalizeText(f.Title) || displayName;
    const claims = normalizeText(f.Name) || (email ? 'i:0#.f|membership|' + email : '');
    return { id: String(lookupId), name: name, email: email, claims: claims };
  } catch (e) {
    return { id: String(lookupId), name: displayName || '', email: '', claims: '' };
  }
}

function parseInitiatedDepartments(fields) {
  const departmentText = safeField(fields, PROJECT_FIELD_CANDIDATES.departments) || safeField(fields, PROJECT_FIELD_CANDIDATES.department);
  return dedupe(splitList(departmentText));
}

function buildLookupKeys(fields, itemId) {
  return dedupe([
    normalizeText(itemId),
    normalizeText(safeField(fields, PROJECT_FIELD_CANDIDATES.clientKey)),
    normalizeText(safeField(fields, CLIENT_FIELD_CANDIDATES.trackingId)),
    normalizeText(safeField(fields, CLIENT_FIELD_CANDIDATES.quoteId)),
    normalizeText(safeField(fields, CLIENT_FIELD_CANDIDATES.name))
  ]);
}

async function resolveSiteAndLists() {
  const siteUrl = new URL(SITE_URL);
  const sitePath = siteUrl.pathname.replace(/\/$/, '');
  const site = await graphGet('/sites/' + siteUrl.hostname + ':' + sitePath);
  const lists = await graphGetAll('/sites/' + site.id + '/lists?$select=id,displayName,webUrl');

  function findByPreferredNames(preferredNames) {
    for (let i = 0; i < preferredNames.length; i += 1) {
      const match = lists.find(function (list) { return list.displayName === preferredNames[i]; });
      if (match) return match;
    }
    return null;
  }

  const clientsList = findByPreferredNames(LIST_NAMES.clients);
  const intakeList = findByPreferredNames(LIST_NAMES.intake);

  return {
    site: {
      id: site.id,
      webUrl: site.webUrl,
      displayName: site.displayName
    },
    lists: {
      clients: clientsList ? { id: clientsList.id, displayName: clientsList.displayName } : null,
      intake: intakeList ? { id: intakeList.id, displayName: intakeList.displayName } : null
    }
  };
}

function buildClientSelect(columnMap) {
  const names = new Set(['id', 'Title', 'LinkTitle']);
  Object.keys(CLIENT_FIELD_CANDIDATES).forEach(function (group) {
    CLIENT_FIELD_CANDIDATES[group].forEach(function (name) {
      // If this candidate is a known display name, select its real internal
      // name (e.g. "Industry type" -> "IndustryType"/"Industry_x0020_type").
      if (columnMap && columnMap[name]) { names.add(columnMap[name]); return; }
      if (name && !/\s/.test(name)) names.add(name);
    });
  });
  return Array.from(names).join(',');
}

async function loadClients(siteId, listId) {
  if (!listId) return [];
  const columnInfo = await getClientColumns();
  const columnMap = columnInfo.columns || {};
  const select = buildClientSelect(columnMap);
  const items = await graphGetAll('/sites/' + siteId + '/lists/' + listId + '/items?$expand=fields($select=' + select + ')&$top=200');

  const parsed = items.map(function (item) {
    return parseClientItem(item.fields || {}, item.id, columnMap);
  });

  await Promise.all(parsed.map(async function (c) {
    c.bdPerson = await resolveSharePointPerson(c.bdPersonLookupId, c.bdPersonName, siteId);
  }));

  return parsed;
}

async function loadInitiatedDepartments(siteId, listId) {
  if (!listId) return {};
  const items = await graphGetAll('/sites/' + siteId + '/lists/' + listId + '/items?$expand=fields&$top=500');
  const map = {};

  items.forEach(function (item) {
    const fields = item.fields || {};
    const departments = parseInitiatedDepartments(fields);
    const services = normalizeText(safeField(fields, PROJECT_FIELD_CANDIDATES.service));
    const scope = normalizeText(safeField(fields, PROJECT_FIELD_CANDIDATES.scope));
    const keys = buildLookupKeys(fields, item.id);
    departments.forEach(function (department) {
      keys.forEach(function (key) {
        if (!map[key]) map[key] = [];
        if (!map[key].some(function (e) { return e.department === department; })) {
          map[key].push({ department: department, services: services, scope: scope });
        }
      });
    });
  });

  return map;
}

let cachedContext = null;
async function getSiteContext() {
  if (cachedContext) return cachedContext;
  cachedContext = await resolveSiteAndLists();
  return cachedContext;
}

const SYSTEM_COLUMN_NAMES = new Set([
  'LinkTitle', 'LinkTitleNoMenu', 'Edit', 'Attachments', 'Created', 'Modified',
  'Author', 'Editor', 'ContentType', 'ID', 'GUID', 'AppAuthor', 'AppEditor',
  'Composed', 'Composite', 'PermMask', 'UniqueId', 'Version', 'WorkflowVersion',
  'owshiddenversion', 'MetaInfo', 'Restricted', 'OriginatorId', 'ProgId',
  'FileLeafRef', 'FileRef', 'FSObjType', 'SortBehavior', 'DocIcon', 'ServerUrl',
  'EncodedAbsUrl', 'BaseName', 'FileDirRef', 'Created_x0020_Date', 'Modified_x0020_Date',
  'Breakpoint', 'CheckoutUser', 'IsCheckedout', 'VirusStatus'
]);

async function getListColumns(siteId, listId) {
  if (!listId) return { columns: {}, choiceColumns: {} };
  const columns = await graphGetAll(
    '/sites/' + siteId + '/lists/' + listId + '/columns?$select=name,displayName,readOnly'
  );
  const map = {};
  const choices = {};
  columns.forEach(function (col) {
    if (!col.displayName || !col.name) return;
    if (col.readOnly) return;
    if (SYSTEM_COLUMN_NAMES.has(col.name)) return;
    if (map[col.displayName] === undefined) map[col.displayName] = col.name;
    if (col.choice && Array.isArray(col.choice.choices) && col.choice.choices.length && choices[col.displayName] === undefined) {
      choices[col.displayName] = col.choice.choices;
    }
  });
  return { columns: map, choiceColumns: choices };
}

let cachedIntakeColumns = null;
async function getIntakeColumns() {
  if (cachedIntakeColumns) return cachedIntakeColumns;
  const ctx = await getSiteContext();
  const listId = ctx.lists.intake && ctx.lists.intake.id;
  cachedIntakeColumns = await getListColumns(ctx.site.id, listId);
  return cachedIntakeColumns;
}

let cachedClientColumns = null;
async function getClientColumns() {
  if (cachedClientColumns) return cachedClientColumns;
  const ctx = await getSiteContext();
  const listId = ctx.lists.clients && ctx.lists.clients.id;
  cachedClientColumns = await getListColumns(ctx.site.id, listId);
  return cachedClientColumns;
}

// Map a list of candidate names (which may be display names like "Industry
// type" or internal names like "IndustryType") to the list's actual internal
// names. Display names are resolved via the column map; unknown entries are
// passed through unchanged so existing hardcoded internal names keep working.
function resolveCandidates(candidates, columnMap) {
  if (!columnMap) return candidates;
  return candidates.map(function (c) { return columnMap[c] || c; });
}

const sharePointUserCache = {};
async function getSharePointUserIdByEmail(siteId, email) {
  const key = String(email).toLowerCase();
  if (sharePointUserCache[key] !== undefined) return sharePointUserCache[key];
  try {
    const items = await graphGetAll(
      "/sites/" + siteId + "/lists('User Information List')/items?$select=id,EMail&$filter=EMail eq '" + String(email).replace(/'/g, "''") + "'"
    );
    const id = items.length ? String(items[0].id) : '';
    sharePointUserCache[key] = id;
    return id;
  } catch (e) {
    sharePointUserCache[key] = '';
    return '';
  }
}

async function resolvePersonLookupId(person, siteId) {
  if (!person) return '';
  const email = normalizeText(person.email);
  if (email) {
    const byEmail = await getSharePointUserIdByEmail(siteId, email);
    if (byEmail) return byEmail;
  }
  const direct = normalizeText(person.id);
  if (/^\d+$/.test(direct)) return direct;
  return '';
}

// Normalize any department's scope into one standard shape for the "Scope
// json" column: a `services` array where each item has a `Service` name and
// an `assets` array, plus department-specific extras when present.
function buildStandardScope(dept) {
  const services = [];

  if (Array.isArray(dept.assessments)) {
    dept.assessments.forEach(function (a) {
      services.push({
        Service: a.assessment,
        assets: (a.assets || []).map(function (r) { return { label: r.label, count: r.count || 0 }; })
      });
    });
  } else if (Array.isArray(dept.serviceDetails)) {
    dept.serviceDetails.forEach(function (s) {
      const item = { Service: s.type, assets: [] };
      if (s.regulator) item.regulator = s.regulator;
      if (s.assessmentType) item.assessmentType = s.assessmentType;
      if (s.mode) item.mode = s.mode;
      if (s.scope && Object.keys(s.scope).length) item.scope = s.scope;
      services.push(item);
    });
  } else if (Array.isArray(dept.services)) {
    dept.services.forEach(function (svc) {
      services.push({ Service: svc, assets: [] });
    });
  }

  return services;
}

async function buildMainTrackerItem(payload, dept, columns, choiceColumns, siteId) {
  const client = payload.client || {};

  const departmentAliases = (metadata && metadata.departmentAliases) || {};
  const departmentName = departmentAliases[normalizeText(dept.department)] || normalizeText(dept.department);

  const warnings = [];

  function resolveInternal(displayName) {
    if (!displayName) return null;
    if (columns && Object.keys(columns).length) {
      if (columns[displayName] !== undefined) return columns[displayName];
      const lower = String(displayName).toLowerCase();
      const match = Object.keys(columns).find(function (k) { return k.toLowerCase() === lower; });
      if (match) return columns[match];
      return null;
    }
    return columns && columns[displayName] ? columns[displayName] : displayName;
  }

  function allowedChoicesFor(logicalName) {
    const displayName = MAIN_TRACKER_FIELDS[logicalName];
    return (choiceColumns && choiceColumns[displayName]) || null;
  }

  function setField(fields, logicalName, value) {
    if (value === undefined || value === null || value === '') return;
    const internal = resolveInternal(MAIN_TRACKER_FIELDS[logicalName]);
    if (!internal) return;
    const allowed = allowedChoicesFor(logicalName);
    if (allowed) {
      const values = Array.isArray(value)
        ? value.map(normalizeText)
        : normalizeText(value).split(';').map(function (v) { return v.trim(); });
      const present = values.filter(Boolean);
      const bad = present.filter(function (v) { return allowed.indexOf(v) === -1; });
      if (bad.length) {
        warnings.push('Skipped "' + MAIN_TRACKER_FIELDS[logicalName] + '" value(s) not in allowed choices: [' + bad.join(', ') + ']. Allowed: ' + allowed.join(', '));
        return;
      }
    }
    fields[internal] = value;
  }

  async function setPersonField(fields, logicalName, person) {
    const id = await resolvePersonLookupId(person, siteId);
    const internal = resolveInternal(MAIN_TRACKER_FIELDS[logicalName]);
    if (!id) {
      warnings.push('Skipped ' + MAIN_TRACKER_FIELDS[logicalName] + ' — no resolved lookup id for person: ' + (person ? (person.email || person.name || '') : ''));
      return;
    }
    if (!internal) {
      warnings.push('Skipped ' + MAIN_TRACKER_FIELDS[logicalName] + ' — column not resolved.');
      return;
    }
    let exists = true;
    try {
      await graphGet('/sites/' + siteId + "/lists('User Information List')/items/" + id + '?$select=id');
    } catch (e) {
      exists = false;
    }
    if (!exists) {
      warnings.push('Skipped BD Person lookup — id ' + id + ' not found in site User Information List (person: ' + (person && person.email ? person.email : person && person.name) + ').');
      return;
    }
    fields[internal + 'LookupId'] = id;
  }

  // Human-readable "Scope" text + the "Services" column string, derived from
  // the department's natural shape.
  let services = '';
  let scopeText = '';

  if (Array.isArray(dept.assessments)) {
    services = dept.assessments
      .map(function (a) { return a.assessment; })
      .filter(function (v) { return v; })
      .join('; ');
    scopeText = dept.assessments
      .map(function (a) {
        const lines = (a.assets || []).map(function (r) { return '  - ' + r.label + ': ' + (r.count || 0); });
        return a.assessment + ':\n' + lines.join('\n');
      })
      .join('\n');
  } else if (Array.isArray(dept.services)) {
    services = dept.services.join('; ');
    scopeText = dept.scope || '';
  } else {
    scopeText = dept.scope || '';
  }

  // Canonical, machine-readable scope saved to the "Scope json" column.
  // VAPT sends a pre-built scopeJson (the SOW table array string) directly
  // from the form; every other department uses the standardized services
  // shape. This keeps the backend API contract unchanged.
  let scopeJson;
  if (typeof dept.scopeJson === 'string' && dept.scopeJson) {
    scopeJson = dept.scopeJson;
  } else {
    const scopeData = {
      department: departmentName,
      services: buildStandardScope(dept)
    };
    scopeJson = JSON.stringify(scopeData, null, 2);
  }

  const general = dept.general || {};
  const fields = {};

  setField(fields, 'title', client.name);
  setField(fields, 'department', departmentName);
  setField(fields, 'services', services);
  setField(fields, 'scope', scopeText);
  setField(fields, 'scopeJson', scopeJson);
  setField(fields, 'industry', client.industry);
  setField(fields, 'projectType', dept.projectType || '');
  await setPersonField(fields, 'bdPerson', payload.bdPerson);
  setField(fields, 'contactName', client.contact);
  setField(fields, 'contactDesignation', client.designation);
  setField(fields, 'contactEmail', client.email);
  setField(fields, 'contactPhone', client.phone);
  setField(fields, 'initiationDate', new Date().toISOString());
  setField(fields, 'priority', general.priority);
  setField(fields, 'remarks', general.remarks);
  const statusChoices = allowedChoicesFor('status');
  let statusValue = DEFAULT_STATUS;
  if (statusChoices && statusChoices.indexOf(statusValue) === -1) {
    if (statusChoices.length) {
      statusValue = statusChoices[0];
      warnings.push('Status default "' + DEFAULT_STATUS + '" is not an allowed choice; used "' + statusValue + '" instead. Allowed: ' + statusChoices.join(', '));
    }
  }
  setField(fields, 'status', statusValue);

  return { fields: fields, warnings: warnings };
}

async function initiateProject(payload) {
  const ctx = await getSiteContext();
  const listId = ctx.lists.intake && ctx.lists.intake.id;
  if (!listId) throw new Error('Main Tracker (intake) list could not be resolved.');
  const columns = await getIntakeColumns();

  const result = { created: [], errors: [], warnings: [], sent: [] };
  const departments = payload.departments || [];
  for (let i = 0; i < departments.length; i += 1) {
    const dept = departments[i];
    let attemptedFields = null;
    try {
      const built = await buildMainTrackerItem(payload, dept, columns.columns, columns.choiceColumns, ctx.site.id);
      attemptedFields = built.fields;
      built.warnings.forEach(function (w) { result.warnings.push({ department: dept.department, message: w }); });
      result.sent.push({ department: dept.department, fields: attemptedFields });
      const item = await graphPost(
        '/sites/' + ctx.site.id + '/lists/' + listId + '/items',
        { fields: attemptedFields }
      );
      result.created.push(item);
    } catch (error) {
      result.errors.push({
        department: dept.department,
        fields: attemptedFields,
        message: error && error.message ? error.message : 'Unknown error'
      });
    }
  }
  if (!result.created.length && result.errors.length) {
    const err = new Error(result.errors[0].message);
    err.details = result.errors;
    throw err;
  }

  if (result.created.length) {
    try {
      const progress = await getClientInitiationProgress(ctx.site.id, listId, payload.client);
      if (progress && ctx.lists.clients && ctx.lists.clients.id) {
        await updateClientMasterProgress(ctx.site.id, ctx.lists.clients.id, payload.client, progress.progressText);
      }
    } catch (updErr) {
      console.warn('Failed to update Client Master initiation progress:', updErr);
    }
  }

  return result;
}

async function buildBootstrap() {
  const siteInfo = await resolveSiteAndLists();

  const clientsPromise = loadClients(siteInfo.site.id, siteInfo.lists.clients && siteInfo.lists.clients.id);
  const initiatedPromise = loadInitiatedDepartments(siteInfo.site.id, siteInfo.lists.intake && siteInfo.lists.intake.id);

  const clients = await clientsPromise;
  const initiatedDepartmentsByClientId = await initiatedPromise;

  return {
    site: siteInfo.site,
    lists: siteInfo.lists,
    clients: clients,
    initiatedDepartmentsByKey: initiatedDepartmentsByClientId
  };
}

async function getClientInitiationProgress(siteId, intakeListId, client) {
  if (!intakeListId || !client) return null;
  const items = await graphGetAll('/sites/' + siteId + '/lists/' + intakeListId + '/items?$expand=fields&$top=500');
  const clientKeys = [
    String(client.id || ''),
    String(client.trackingId || ''),
    String(client.quoteId || ''),
    String(client.clientName || '')
  ];
  const initiated = new Set();

  items.forEach(function (item) {
    const fields = item.fields || {};
    const departments = dedupe(splitList(safeField(fields, PROJECT_FIELD_CANDIDATES.departments) || safeField(fields, PROJECT_FIELD_CANDIDATES.department)));
    if (!departments.length) return;
    const itemKeys = buildLookupKeys(fields, item.id);
    const matches = clientKeys.some(function (ck) {
      if (!ck) return false;
      return itemKeys.some(function (ik) { return ik === ck; });
    });
    if (matches) {
      departments.forEach(function (d) { initiated.add(d); });
    }
  });

  const total = Array.isArray(client.departmentsInvolved) ? client.departmentsInvolved.length : 0;
  const count = initiated.size;
  return {
    initiatedCount: count,
    totalDepartments: total,
    progressText: total > 0 && count >= total ? 'Done' : (total > 0 ? count + '/' + total : '')
  };
}

async function updateClientMasterProgress(siteId, clientsListId, client, progressText) {
  if (!clientsListId || !client || !client.id || !progressText) return null;
  const columns = await getListColumns(siteId, clientsListId);
  const displayNames = Object.keys(columns.columns);
  let internalName = null;
  for (var i = 0; i < displayNames.length; i++) {
    if (displayNames[i].toLowerCase() === 'initiation progress') {
      internalName = columns.columns[displayNames[i]];
      break;
    }
  }
  if (!internalName) {
    internalName = 'Initiation_x0020_Progress';
  }

  const res = await graphPatch(
    '/sites/' + siteId + '/lists/' + clientsListId + '/items/' + client.id,
    { fields: { [internalName]: progressText } }
  );
  return res;
}

module.exports = {
  buildBootstrap,
  initiateProject,
  getSiteContext,
  getIntakeColumns
};
