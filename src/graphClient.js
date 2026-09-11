const { getAccessToken } = require('./auth');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const GRAPH_RETRIES = 3;
const GRAPH_RETRY_DELAY_BASE = 1000;
const GRAPH_TIMEOUT = 60000;
const SIMPLE_UPLOAD_MAX_BYTES = 4 * 1024 * 1024;
const UPLOAD_CHUNK_BYTES = 10 * 1024 * 1024;
const ATTACHMENT_MAX_BYTES = 250 * 1024 * 1024;

async function sleep(ms) {
  return new Promise(function (resolve) { setTimeout(resolve, ms); });
}

async function graphFetchWithRetry(url, options, attempt) {
  const controller = new AbortController();
  const timeout = setTimeout(function () { controller.abort(); }, GRAPH_TIMEOUT);
  try {
    const response = await fetch(url, Object.assign({}, options, { signal: controller.signal }));
    clearTimeout(timeout);

    if (response.ok) return response;

    const retryAfter = response.headers.get('Retry-After');
    if ((response.status === 429 || response.status === 503) && attempt < GRAPH_RETRIES) {
      const delay = retryAfter ? (parseInt(retryAfter, 10) * 1000) : (GRAPH_RETRY_DELAY_BASE * Math.pow(2, attempt));
      await sleep(delay);
      return graphFetchWithRetry(url, options, attempt + 1);
    }

    return response;
  } catch (error) {
    clearTimeout(timeout);
    if (error.name === 'AbortError') {
      throw new Error('Graph request to ' + url + ' timed out after ' + GRAPH_TIMEOUT + 'ms');
    }
    if (attempt < GRAPH_RETRIES) {
      await sleep(GRAPH_RETRY_DELAY_BASE * Math.pow(2, attempt));
      return graphFetchWithRetry(url, options, attempt + 1);
    }
    throw error;
  }
}

async function graphGet(relativeUrl, extraHeaders) {
  const token = await getAccessToken();
  const headers = Object.assign({
    Authorization: 'Bearer ' + token,
    Accept: 'application/json'
  }, extraHeaders || {});
  const response = await graphFetchWithRetry(GRAPH_BASE + relativeUrl, { headers: headers }, 0);

  if (!response.ok) {
    const text = await response.text();
    throw new Error('Graph GET failed for ' + relativeUrl + ': ' + text);
  }

  return response.json();
}

async function graphWrite(method, relativeUrl, body, extraHeaders) {
  const token = await getAccessToken();
  const headers = Object.assign({
    Authorization: 'Bearer ' + token,
    Accept: 'application/json',
    'Content-Type': 'application/json'
  }, extraHeaders || {});
  const response = await graphFetchWithRetry(GRAPH_BASE + relativeUrl, {
    method: method,
    headers: headers,
    body: body != null ? JSON.stringify(body) : undefined
  }, 0);

  if (!response.ok) {
    const text = await response.text();
    throw new Error('Graph ' + method + ' failed for ' + relativeUrl + ': ' + text);
  }

  if (response.status === 204) return {};
  return response.json();
}

function graphPost(relativeUrl, body, extraHeaders) {
  return graphWrite('POST', relativeUrl, body, extraHeaders);
}

function graphPatch(relativeUrl, body, extraHeaders) {
  return graphWrite('PATCH', relativeUrl, body, extraHeaders);
}

function mapGraphUser(u) {
  const email = u.mail || u.userPrincipalName || '';
  return {
    id: u.id,
    name: u.displayName || email,
    email: email,
    department: u.department || '',
    title: u.jobTitle || '',
    claims: email ? 'i:0#.f|membership|' + email : ''
  };
}

function toPerson(u) {
  return mapGraphUser(u || {});
}

async function graphSearchUsers(query, department) {
  const select = '$select=id,displayName,mail,userPrincipalName,jobTitle,department';

  if (!query && !department) {
    const url = '/users?' + select + '&$top=100';
    const all = await graphGetAll(url);
    return all.map(mapGraphUser);
  }

  const clauses = [];
  if (department) {
    const escaped = String(department).replace(/'/g, "''");
    clauses.push("department eq '" + escaped + "'");
  }
  if (query) {
    const term = String(query).trim().replace(/'/g, "''");
    clauses.push("(startswith(displayName,'" + term + "') or startswith(mail,'" + term + "') or startswith(userPrincipalName,'" + term + "'))");
  }
  const filter = clauses.join(' and ');
  const url = '/users?' + '$filter=' + encodeURIComponent(filter) + '&' + select + '&$top=25';
  const payload = await graphGet(url);
  const users = Array.isArray(payload.value) ? payload.value : [];
  return users.map(mapGraphUser);
}

async function graphGetUserById(userId) {
  const payload = await graphGet('/users/' + userId + '?$select=id,displayName,mail,userPrincipalName,jobTitle,department');
  return mapGraphUser(payload);
}

async function graphGetAll(relativeUrl) {
  const items = [];
  let nextUrl = relativeUrl;

  while (nextUrl) {
    const payload = await graphGet(nextUrl);
    if (Array.isArray(payload.value)) items.push.apply(items, payload.value);
    nextUrl = payload['@odata.nextLink'] ? payload['@odata.nextLink'].replace(GRAPH_BASE, '') : null;
  }

  return items;
}

function graphPath() {
  const args = Array.prototype.slice.call(arguments, 0);
  return '/' + args.map(function (segment) {
    const value = String(segment);
    const queryIndex = value.indexOf('?');
    if (queryIndex === -1) return encodeURIComponent(value);
    return encodeURIComponent(value.slice(0, queryIndex)) + value.slice(queryIndex);
  }).join('/');
}

async function graphBinaryRequest(method, relativeUrl, body, contentType) {
  const token = await getAccessToken();
  const response = await graphFetchWithRetry(GRAPH_BASE + relativeUrl, {
    method: method,
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/json',
      'Content-Type': contentType || 'application/octet-stream'
    },
    body: body
  }, 0);
  if (!response.ok) {
    const text = await response.text();
    throw new Error('Graph binary request failed for ' + relativeUrl + ': ' + text);
  }
  return response.status === 204 ? {} : response.json();
}

async function listDriveChildren(driveId, parentId) {
  return graphGetAll(graphPath('drives', driveId, 'items', parentId, 'children?$select=id,name,size,webUrl,file,folder'));
}

async function findDriveChild(driveId, parentId, name) {
  const children = await listDriveChildren(driveId, parentId);
  return children.find(function (child) { return child.name === name; }) || null;
}

async function ensureDriveFolder(driveId, parentId, name) {
  const existing = await findDriveChild(driveId, parentId, name);
  if (existing) {
    if (!existing.folder) throw new Error('Documents/' + name + ' exists but is not a folder.');
    return existing;
  }

  try {
    return await graphPost(graphPath('drives', driveId, 'items', parentId, 'children'), {
      name: name,
      folder: {},
      '@microsoft.graph.conflictBehavior': 'fail'
    });
  } catch (error) {
    // A concurrent initiation may create the folder between the lookup and POST.
    const createdByAnotherRequest = await findDriveChild(driveId, parentId, name);
    if (createdByAnotherRequest && createdByAnotherRequest.folder) return createdByAnotherRequest;
    throw error;
  }
}

function splitFileName(fileName) {
  const dot = fileName.lastIndexOf('.');
  return dot > 0 ? { stem: fileName.slice(0, dot), extension: fileName.slice(dot) } : { stem: fileName, extension: '' };
}

async function chooseDriveFileName(driveId, parentId, requestedName) {
  const safeName = String(requestedName || 'unnamed').replace(/[\\/\0]/g, '_').trim() || 'unnamed';
  const children = await listDriveChildren(driveId, parentId);
  const used = new Set(children.map(function (child) { return child.name; }));
  if (!used.has(safeName)) return safeName;
  const parts = splitFileName(safeName);
  let index = 1;
  let candidate;
  do {
    candidate = parts.stem + ' (' + index + ')' + parts.extension;
    index += 1;
  } while (used.has(candidate));
  return candidate;
}

function driveFileContentUrl(driveId, parentId, fileName) {
  return '/drives/' + encodeURIComponent(driveId) + '/items/' + encodeURIComponent(parentId) + ':/' + encodeURIComponent(fileName) + ':/content';
}

async function uploadDriveFile(driveId, parentId, fileName, fileBuffer, contentType) {
  const targetUrl = driveFileContentUrl(driveId, parentId, fileName);
  if (fileBuffer.length <= SIMPLE_UPLOAD_MAX_BYTES) {
    return graphBinaryRequest('PUT', targetUrl, fileBuffer, contentType);
  }

  const session = await graphPost(
    '/drives/' + encodeURIComponent(driveId) + '/items/' + encodeURIComponent(parentId) + ':/' + encodeURIComponent(fileName) + ':/createUploadSession',
    { item: { '@microsoft.graph.conflictBehavior': 'fail', name: fileName } }
  );
  let offset = 0;
  let uploaded = null;
  while (offset < fileBuffer.length) {
    const end = Math.min(offset + UPLOAD_CHUNK_BYTES, fileBuffer.length);
    const chunk = fileBuffer.subarray(offset, end);
    const response = await graphFetchWithRetry(session.uploadUrl, {
      method: 'PUT',
      headers: {
        'Content-Length': String(chunk.length),
        'Content-Range': 'bytes ' + offset + '-' + (end - 1) + '/' + fileBuffer.length
      },
      body: chunk
    }, 0);
    if (!response.ok) {
      const text = await response.text();
      throw new Error('Graph upload session failed for ' + fileName + ': ' + text);
    }
    uploaded = await response.json();
    offset = end;
  }
  return uploaded;
}

async function uploadAttachmentToSharePointItem(drive, itemId, file) {
  const fileName = String(file.name || '').replace(/\\/g, '/').split('/').pop() || 'unnamed';
  const fileSize = Number(file.size || 0);
  const contentType = String(file.type || 'application/octet-stream');
  if (!fileSize) throw new Error('File "' + fileName + '" is empty (0 bytes).');
  if (fileSize > ATTACHMENT_MAX_BYTES) throw new Error('File "' + fileName + '" exceeds the maximum allowed size of 250 MB.');
  if (fileSize !== file.buffer.length) throw new Error('File "' + fileName + '" size validation failed.');

  const itemFolder = await ensureDriveFolder(drive.id, drive.attachmentsFolderId, String(itemId));
  if (!drive.itemFolderUrls) drive.itemFolderUrls = {};
  drive.itemFolderUrls[String(itemId)] = itemFolder.webUrl || '';
  const targetName = await chooseDriveFileName(drive.id, itemFolder.id, fileName);
  const uploaded = await uploadDriveFile(drive.id, itemFolder.id, targetName, file.buffer, contentType);
  return {
    name: uploaded.name || targetName,
    size: uploaded.size != null ? uploaded.size : fileSize,
    webUrl: uploaded.webUrl || '',
    driveItemId: uploaded.id || '',
    mainTrackerId: String(itemId),
    attachmentsFolderUrl: itemFolder.webUrl || drive.attachmentsFolderUrl || ''
  };
}

async function resolveDocumentsDrive(siteId) {
  const drive = await graphGet(graphPath('sites', siteId, 'drive?$select=id,name,webUrl,driveType'));
  const root = await graphGet(graphPath('drives', drive.id, 'root?$select=id,name,webUrl'));
  const attachmentsFolder = await ensureDriveFolder(drive.id, root.id, 'Attachments');
  return {
    id: drive.id,
    name: drive.name,
    webUrl: drive.webUrl,
    rootId: root.id,
    attachmentsFolderId: attachmentsFolder.id,
    attachmentsFolderUrl: attachmentsFolder.webUrl || ''
  };
}

function graphDelete(relativeUrl, extraHeaders) {
  return graphWrite('DELETE', relativeUrl, undefined, extraHeaders);
}

module.exports = {
  graphGet,
  graphGetAll,
  graphPost,
  graphPatch,
  graphDelete,
  graphSearchUsers,
  graphGetUserById,
  mapGraphUser,
  toPerson,
  resolveDocumentsDrive,
  uploadAttachmentToSharePointItem,
  SIMPLE_UPLOAD_MAX_BYTES,
  ATTACHMENT_MAX_BYTES
};
