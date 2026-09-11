const { getAccessToken } = require('./auth');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const GRAPH_RETRIES = 3;
const GRAPH_RETRY_DELAY_BASE = 1000;
const GRAPH_TIMEOUT = 60000;
const ATTACHMENT_MAX_BYTES = 50 * 1024 * 1024;

function graphPath() {
  var args = Array.prototype.slice.call(arguments, 0);
  return '/' + args.map(function (s) { return encodeURIComponent(s); }).join('/');
}

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

async function createAttachmentUploadSession(siteId, listId, itemId, fileName, fileSize, contentType) {
  if (fileSize > ATTACHMENT_MAX_BYTES) {
    throw new Error('File "' + fileName + '" exceeds the maximum allowed size of ' + (ATTACHMENT_MAX_BYTES / 1024 / 1024) + ' MB.');
  }

  const sessionBody = {
    item: {
      attachmentType: 'file',
      name: fileName
    }
  };

  const relativeUrl = graphPath('sites', siteId, 'lists', listId, 'items', itemId, 'attachments', 'createUploadSession');
  const payload = await graphPost(relativeUrl, sessionBody);

  if (!payload || !payload.uploadUrl) {
    throw new Error('Graph did not return an upload URL for attachment: ' + fileName);
  }

  return payload.uploadUrl;
}

async function streamBytesToUploadUrl(uploadUrl, buffer, fileName, contentType) {
  const headers = {
    'Content-Type': contentType || 'application/octet-stream',
    'Content-Length': String(Buffer.isBuffer(buffer) ? buffer.length : buffer.byteLength)
  };

  const controller = new AbortController();
  const timeout = setTimeout(function () { controller.abort(); }, GRAPH_TIMEOUT);
  let response;
  try {
    response = await fetch(uploadUrl, {
      method: 'PUT',
      headers: headers,
      body: buffer,
      signal: controller.signal
    });
  } catch (error) {
    clearTimeout(timeout);
    if (error.name === 'AbortError') {
      throw new Error('Attachment upload for "' + fileName + '" timed out after ' + GRAPH_TIMEOUT + 'ms.');
    }
    throw error;
  }
  clearTimeout(timeout);

  if (!response.ok) {
    const text = await response.text();
    const msg = text || ('HTTP ' + response.status);
    throw new Error('Graph attachment upload failed for "' + fileName + '": ' + msg);
  }
}

async function uploadAttachmentToSharePointItem(siteId, listId, itemId, file) {
  const fileName = String(file.name || '').replace(/\\/g, '/').split('/').pop() || 'unnamed';
  const fileSize = Number(file.size || 0);
  const contentType = String(file.type || 'application/octet-stream');

  if (!fileSize) {
    throw new Error('File "' + fileName + '" is empty (0 bytes).');
  }

  let buffer;
  if (typeof Buffer !== 'undefined' && Buffer.isBuffer(file.buffer)) {
    buffer = file.buffer;
  } else if (file.arrayBuffer) {
    buffer = Buffer.from(await file.arrayBuffer());
  } else {
    throw new Error('Unsupported file object for attachment upload: ' + fileName);
  }

  const uploadUrl = await createAttachmentUploadSession(siteId, listId, itemId, fileName, fileSize, contentType);
  await streamBytesToUploadUrl(uploadUrl, buffer, fileName, contentType);
  return { name: fileName, size: fileSize, type: contentType };
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
  createAttachmentUploadSession,
  streamBytesToUploadUrl,
  uploadAttachmentToSharePointItem,
  ATTACHMENT_MAX_BYTES
};
