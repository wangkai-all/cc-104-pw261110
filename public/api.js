/* Thin API client for the importer server. */

async function req(method, url, body, options = {}) {
  const init = { method, ...options };
  if (body !== undefined) {
    if (body instanceof Blob || body instanceof ArrayBuffer) {
      init.body = body;
      init.headers = { 'Content-Type': 'application/octet-stream', ...(options.headers || {}) };
    } else {
      init.body = JSON.stringify(body);
      init.headers = { 'Content-Type': 'application/json', ...(options.headers || {}) };
    }
  }
  const res = await fetch(url, init);
  const text = await res.text();
  let data = null;
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      data = { error: { message: text.slice(0, 300) } };
    }
  }
  if (!res.ok) {
    const err = new Error(data?.error?.message || data?.message || `请求失败（${res.status}）`);
    err.status = res.status;
    err.payload = data;
    throw err;
  }
  return data;
}

export const api = {
  health: () => req('GET', '/api/health'),
  upload: (file) => req('POST', `/api/files?name=${encodeURIComponent(file.name)}`, file),
  files: () => req('GET', '/api/files'),
  file: (id) => req('GET', `/api/files/${id}`),
  suggest: (fileId, body = {}) => req('POST', `/api/files/${fileId}/suggest`, body),
  analyze: (fileId, body) => req('POST', `/api/files/${fileId}/analyze`, body),
  specs: () => req('GET', '/api/specs'),
  saveSpec: (body) => req('POST', '/api/specs', body),
  deleteSpec: (specId) => req('DELETE', `/api/specs/${specId}`),
  applySpec: (specId, fileId) => req('POST', `/api/specs/${specId}/apply`, { fileId }),
  preview: (body) => req('POST', '/api/preview', body),
  createBatch: (body) => req('POST', '/api/batches', body),
  batches: (fileId) => req('GET', `/api/batches${fileId ? `?fileId=${encodeURIComponent(fileId)}` : ''}`),
  batch: (id) => req('GET', `/api/batches/${id}`),
  deleteBatch: (id) => req('DELETE', `/api/batches/${id}`),
  originalUrl: (fileId) => `/api/files/${fileId}/original`,
};
