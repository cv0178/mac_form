<script setup>
import { computed, onMounted, ref } from 'vue';
import { clearExports, deleteExport, listExports } from '../services/localApi';

const records = ref([]);
const errorMessage = ref('');
const isLoading = ref(true);
const isClearing = ref(false);

const totalCount = computed(() => records.value.length);

function formatTime(value) {
  return value
    ? new Intl.DateTimeFormat('zh-CN', { dateStyle: 'medium', timeStyle: 'short' }).format(new Date(value))
    : '';
}

function toViewModel(record) {
  const isWord = record.output_name.toLowerCase().endsWith('.docx');
  return {
    ...record,
    icon: isWord ? 'W' : 'X',
    color: isWord ? 'word' : 'excel',
    createdAt: formatTime(record.created_at),
  };
}

async function refresh() {
  try {
    records.value = (await listExports()).map(toViewModel);
    errorMessage.value = '';
  } catch {
    errorMessage.value = '无法连接本地后端，请先启动 Python 服务。';
  } finally {
    isLoading.value = false;
  }
}

async function openFile(record) {
  try {
    await window.desktopApi?.openOutputFile?.(record.output_path);
  } catch {
    errorMessage.value = `无法打开文件，可能已被移动或删除：${record.output_name}`;
  }
}

async function revealFile(record) {
  try {
    await window.desktopApi?.showOutputInFolder?.(record.output_path);
  } catch {
    errorMessage.value = `无法定位文件，可能已被移动或删除：${record.output_name}`;
  }
}

async function removeRecord(id) {
  try {
    await deleteExport(id);
    await refresh();
  } catch {
    errorMessage.value = '删除记录失败，请确认后端服务正在运行。';
  }
}

async function clearAllRecords() {
  const count = records.value.length;
  if (!count) return;
  const confirmed = window.confirm(
    `确定清空全部 ${count} 条填写记录吗？\n\n此操作只会清除应用中的历史列表，不会删除磁盘上的已导出文件，且无法恢复。`,
  );
  if (!confirmed) return;
  isClearing.value = true; errorMessage.value = '';
  try {
    const result = await clearExports();
    records.value = [];
    // 仅在异常竞态下已无记录时才不提示；正常情况不会遮挡空状态说明。
    if (!result.deleted_count) errorMessage.value = '没有可清空的填写记录。';
  } catch (error) {
    errorMessage.value = error.response?.data?.detail ?? '清空记录失败，请确认后端服务正在运行。';
  } finally { isClearing.value = false; }
}

onMounted(refresh);
</script>

<template>
  <main class="content">
    <header class="page-header">
      <div>
        <p class="page-kicker">EXPORT HISTORY</p>
        <h1>填写记录</h1>
        <p class="page-description">这里列出每次生成的填写结果文件，可以随时打开查看。删除记录不会删除磁盘上的文件。</p>
      </div>
      <div class="history-header-actions"><div class="library-count">{{ totalCount }} 次记录</div><button v-if="records.length" class="clear-records-button" type="button" :disabled="isClearing" @click="clearAllRecords">{{ isClearing ? '正在清空…' : '清空记录' }}</button></div>
    </header>

    <section v-if="isLoading" class="analysis-state">正在读取填写记录…</section>

    <template v-else>
      <p v-if="errorMessage" class="api-error">{{ errorMessage }}</p>

      <section v-if="records.length" class="library-section">
        <div class="record-list">
          <article v-for="record in records" :key="record.id" class="record-item">
            <span class="document-icon" :class="record.color">{{ record.icon }}</span>
            <div class="record-main">
              <strong>{{ record.output_name }}</strong>
              <small>
                {{ record.createdAt }} · 表单「{{ record.form_name }}」 · 已填写 {{ record.filled_field_count }} 项<template v-if="record.missing_field_count"> · {{ record.missing_field_count }} 项未填</template>
              </small>
              <small v-if="!record.file_exists" class="record-missing">文件已被移动或删除，无法打开</small>
            </div>
            <div class="record-actions">
              <button class="view-button" type="button" :disabled="!record.file_exists" @click="openFile(record)">打开文件</button>
              <button class="view-button" type="button" :disabled="!record.file_exists" @click="revealFile(record)">所在文件夹</button>
              <button class="remove-button" type="button" @click="removeRecord(record.id)">删除记录</button>
            </div>
          </article>
        </div>
      </section>

      <div v-else class="empty-library">
        <span>◷</span>
        <h2>还没有填写记录</h2>
        <p>在「开始填表」里完成一次自动填写后，导出的文件会出现在这里。</p>
        <RouterLink class="primary-button" to="/">前往开始填表</RouterLink>
      </div>
    </template>
  </main>
</template>
