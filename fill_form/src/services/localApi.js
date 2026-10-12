import axios from 'axios';

// 逐字段调用 LLM 会比较慢，超时放宽到 5 分钟，避免整批建议被中断
const localApi = axios.create({ baseURL: 'http://127.0.0.1:8999', timeout: 300_000 });

export async function listKnowledgeFiles() { const { data } = await localApi.get('/api/knowledge'); return data; }
export async function getKnowledgeFile(fileId) { const { data } = await localApi.get(`/api/knowledge/${fileId}`); return data; }
export async function checkKnowledgeImport(sourcePath) { const { data } = await localApi.post('/api/knowledge/import/check', { source_path: sourcePath }); return data; }
export async function importKnowledgeFile(sourcePath, replaceFileId = null) {
  const { data } = await localApi.post('/api/knowledge/import', { source_path: sourcePath, replace_file_id: replaceFileId });
  return data;
}
export async function deleteKnowledgeFile(fileId) { await localApi.delete(`/api/knowledge/${fileId}`); }
// 首次识别会从 PaddleOCR 官方源下载模型，网络较慢时可能超过普通接口的 5 分钟。
export async function ocrKnowledgePdf(fileId) {
  const { data } = await localApi.post(`/api/knowledge/${fileId}/ocr`, {}, { timeout: 900_000 });
  return data;
}
export async function getKnowledgeOcrStatus(fileId) {
  const { data } = await localApi.get(`/api/knowledge/${fileId}/ocr-status`);
  return data;
}
export async function importFormFile(sourcePath) { const { data } = await localApi.post('/api/forms/import', { source_path: sourcePath }); return data; }
export async function getForm(formId) { const { data } = await localApi.get(`/api/forms/${formId}`); return data; }

export async function generateFormSuggestions(formId, writingStyle) {
  // 不传 writing_style 时，由后端使用「AI 设置」里保存的写作风格
  const payload = writingStyle ? { writing_style: writingStyle } : {};
  const { data } = await localApi.post(`/api/forms/${formId}/suggestions`, payload);
  return data;
}

export async function getLlmSettings() {
  const { data } = await localApi.get('/api/settings/llm');
  return data;
}

export async function saveLlmSettings(settings) {
  const { data } = await localApi.put('/api/settings/llm', settings);
  return data;
}

export async function testLlmConnection(settings) {
  const { data } = await localApi.post('/api/settings/llm/test', settings, { timeout: 30_000 });
  return data;
}

export async function exportFilledForm(formId, values) {
  const { data } = await localApi.post(`/api/forms/${formId}/export`, { values });
  return data;
}

export async function listExports() { const { data } = await localApi.get('/api/exports'); return data; }
export async function deleteExport(exportId) { await localApi.delete(`/api/exports/${exportId}`); }
export async function clearExports() { const { data } = await localApi.delete('/api/exports'); return data; }

export async function getStorageSettings() {
  const { data } = await localApi.get('/api/settings/storage');
  return data;
}

export async function getWritingStyles() {
  const { data } = await localApi.get('/api/settings/writing-styles');
  return data;
}

export async function saveWritingStyles(writingStyles) {
  const { data } = await localApi.put('/api/settings/writing-styles', writingStyles);
  return data;
}

export async function saveStorageSettings(storage) {
  const { data } = await localApi.put('/api/settings/storage', storage);
  return data;
}

export async function getNotificationSettings() {
  const { data } = await localApi.get('/api/settings/notifications');
  return data;
}

export async function saveNotificationSettings(settings) {
  const { data } = await localApi.put('/api/settings/notifications', settings);
  return data;
}
