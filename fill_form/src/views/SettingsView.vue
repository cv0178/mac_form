<script setup>
import { onMounted, ref } from 'vue';
import { getLlmSettings, getNotificationSettings, getStorageSettings, saveLlmSettings, saveNotificationSettings, saveStorageSettings, testLlmConnection } from '../services/localApi';

const baseUrl = ref(''); const model = ref(''); const apiKey = ref('');
const hasApiKey = ref(false); const isLoading = ref(true); const isSaving = ref(false); const isTesting = ref(false); const message = ref(''); const isError = ref(false);
const outputDir = ref(''); const defaultOutputDir = ref(''); const effectiveOutputDir = ref('');
const notificationsEnabled = ref(true); const isSavingNotifications = ref(false);

function flash(text, isErrorFlag = false) {
  message.value = text;
  isError.value = isErrorFlag;
}

async function loadSettings() {
  try {
    const settings = await getLlmSettings();
    baseUrl.value = settings.base_url; model.value = settings.model; hasApiKey.value = settings.has_api_key;
    const storage = await getStorageSettings();
    outputDir.value = storage.output_dir; defaultOutputDir.value = storage.default_output_dir; effectiveOutputDir.value = storage.effective_output_dir;
    const notifications = await getNotificationSettings();
    notificationsEnabled.value = notifications.enabled;
  } catch { flash('无法连接本地后端，请先在 PyCharm 中启动 Python 服务。', true); }
  finally { isLoading.value = false; }
}

async function applyOutputDir(dir) {
  message.value = ''; isError.value = false;
  try {
    const storage = await saveStorageSettings({ output_dir: dir });
    outputDir.value = storage.output_dir; effectiveOutputDir.value = storage.effective_output_dir;
    flash(dir ? '导出目录已更新。' : '已恢复默认导出目录。');
  }
  catch (error) { flash(error.response?.data?.detail ?? '保存导出目录失败，请检查后端服务。', true); }
}

async function chooseOutputDir() {
  const picked = await window.desktopApi?.selectDirectory?.();
  if (picked) await applyOutputDir(picked);
}

function resetOutputDir() { applyOutputDir(''); }

async function updateNotifications() {
  isSavingNotifications.value = true; message.value = ''; isError.value = false;
  try {
    const result = await saveNotificationSettings({ enabled: notificationsEnabled.value });
    notificationsEnabled.value = result.enabled;
    flash(result.enabled ? '已开启填写完成通知。' : '已关闭填写完成通知。');
  }
  catch (error) {
    notificationsEnabled.value = !notificationsEnabled.value;
    flash(error.response?.data?.detail ?? '保存通知设置失败，请检查后端服务。', true);
  }
  finally { isSavingNotifications.value = false; }
}

async function saveSettings() {
  isSaving.value = true; message.value = ''; isError.value = false;
  try {
    const settings = await saveLlmSettings({ base_url: baseUrl.value, model: model.value, api_key: apiKey.value || null });
    hasApiKey.value = settings.has_api_key; apiKey.value = '';
    flash('模型配置已保存到本机。');
  }
  catch (error) { flash(error.response?.data?.detail ?? '保存失败，请检查后端服务。', true); }
  finally { isSaving.value = false; }
}

async function testConnection() {
  isTesting.value = true; message.value = ''; isError.value = false;
  try {
    const result = await testLlmConnection({
      base_url: baseUrl.value,
      model: model.value,
      api_key: apiKey.value || null,
    });
    flash(`连接成功：${result.model} 可用，耗时 ${(result.latency_ms / 1000).toFixed(1)} 秒。`);
  }
  catch (error) {
    flash(error.response?.data?.detail ?? '测试连接失败，请确认后端服务正在运行。', true);
  }
  finally { isTesting.value = false; }
}

onMounted(loadSettings);
</script>

<template>
  <main class="content">
    <header class="page-header">
      <div>
        <p class="page-kicker">AI SETTINGS</p>
        <h1>AI 设置</h1>
        <p class="page-description">配置模型服务后，自动填写时才会把相关资料片段发送给模型。这里管理模型和导出位置。</p>
      </div>
    </header>

    <section v-if="isLoading" class="analysis-state">正在读取本地配置…</section>

    <form v-else class="settings-panel" @submit.prevent="saveSettings">
      <div class="settings-note"><span>⌂</span><p>配置仅保存在本机应用数据目录的 <code>settings.json</code> 中；不会同步到云端。</p></div>

      <label><span>模型服务地址</span><input v-model.trim="baseUrl" type="url" required placeholder="模型地址，例如：https://api.deepseek.com" /><small>应填写兼容 OpenAI Chat Completions 的 API 根地址。</small></label>
      <label><span>模型名称</span><input v-model.trim="model" type="text" required placeholder="例如：你的模型名称，比如deepseek-v4.1-flash" /></label>
      <label><span>API 密钥</span><input v-model="apiKey" type="password" :required="!hasApiKey" :placeholder="hasApiKey ? '已保存密钥；留空则保持不变' : '输入 API 密钥'" /><small v-if="hasApiKey">已保存密钥。为安全起见，客户端不会显示原始密钥。</small></label>

      <footer class="settings-actions model-settings-actions"><button class="secondary-button" type="button" :disabled="isSaving || isTesting" @click="testConnection">{{ isTesting ? '正在测试…' : '测试连接' }}</button><button class="primary-button" type="submit" :disabled="isSaving || isTesting">{{ isSaving ? '正在保存…' : '保存模型配置' }}</button></footer>
      <p v-if="message" :class="isError ? 'api-error' : 'settings-success'">{{ message }}</p>

      <section class="notification-setting">
        <div>
          <strong>填写完成通知</strong>
          <p>表单成功导出后，向 Windows 或 macOS 系统通知中心发送提醒。</p>
        </div>
        <label class="notification-toggle">
          <input v-model="notificationsEnabled" type="checkbox" :disabled="isSavingNotifications" @change="updateNotifications" />
          <span>{{ notificationsEnabled ? '已开启' : '已关闭' }}</span>
        </label>
      </section>

      <section class="storage-setting">
        <div class="storage-heading"><span>导出文件目录</span><code>{{ effectiveOutputDir }}</code></div>
        <small v-if="outputDir">已自定义（默认位置：{{ defaultOutputDir }}）</small>
        <small v-else>当前使用默认位置</small>
        <div class="storage-actions">
          <button class="secondary-button" type="button" @click="chooseOutputDir">更改目录</button>
          <button class="text-button" type="button" :disabled="!outputDir" @click="resetOutputDir">恢复默认</button>
        </div>
        <p class="storage-hint">注意：更改目录后，<b>之前已经导出的历史文件仍留在原位置，不会自动搬过去</b>。程序里的「填写记录」依然可以打开它们。</p>
      </section>
    </form>
  </main>
</template>
