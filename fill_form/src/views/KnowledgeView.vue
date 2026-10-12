<script setup>
import { computed, onMounted, ref } from "vue";
import {
  checkKnowledgeImport,
  deleteKnowledgeFile,
  getKnowledgeFile,
  getKnowledgeOcrStatus,
  importKnowledgeFile,
  listKnowledgeFiles,
  ocrKnowledgePdf,
} from "../services/localApi";

const files = ref([]);
const isDragging = ref(false);
const isImporting = ref(false);
const isLoadingContent = ref(false);
const ocrFileId = ref("");
const ocrProgress = ref({ progress: 0, message: "" });
const errorMessage = ref("");
const successMessage = ref("");
const selectedRecord = ref(null);
const sameNameConflict = ref(null);
let resolveConflictChoice = null;
const selectedCount = computed(() => files.value.length);

function displayStatus(status) {
  return (
    { parsed: "已解析", failed: "解析失败", pending: "待解析" }[status] ??
    status
  );
}
function formatTime(value) {
  return value
    ? new Intl.DateTimeFormat("zh-CN", {
        dateStyle: "medium",
        timeStyle: "short",
      }).format(new Date(value))
    : "刚刚";
}
function toViewModel(record) {
  const documentType = {
    ".docx": { icon: "W", color: "word" },
    ".xlsx": { icon: "X", color: "excel" },
    ".pdf": { icon: "P", color: "pdf" },
  }[record.extension] ?? { icon: "F", color: "file" };
  const needsOcr =
    record.extension === ".pdf" &&
    record.parse_status === "failed" &&
    String(record.parse_error || "").includes("未提取到可复制文字");
  return {
    ...record,
    ...documentType,
    kind: needsOcr ? "PDF 扫描件 · 请点击右侧“OCR 识别”" : record.kind,
    importedAt: formatTime(record.imported_at),
    status: displayStatus(record.parse_status),
  };
}
async function refreshFiles() {
  try {
    files.value = (await listKnowledgeFiles()).map(toViewModel);
    errorMessage.value = "";
  } catch {
    errorMessage.value =
      "无法连接本地后端，请先在 PyCharm 中启动 Python 服务。";
  }
}
async function addFiles(pickedFiles) {
  isDragging.value = false;
  if (!pickedFiles?.length) return;
  isImporting.value = true;
  errorMessage.value = "";
  successMessage.value = "";
  try {
    const results = [];
    let skippedCount = 0;
    for (const file of pickedFiles) {
      const check = await checkKnowledgeImport(file.path);
      let replaceFileId = null;
      if (check.candidates.length) {
        const choice = await askSameNameConflict(file, check.candidates);
        if (choice.action === "cancel") {
          skippedCount += 1;
          continue;
        }
        replaceFileId = choice.action === "replace" ? choice.fileId : null;
      }
      results.push(await importKnowledgeFile(file.path, replaceFileId));
    }
    await refreshFiles();
    const updatedCount = results.filter(
      (item) => item.import_action === "updated",
    ).length;
    const createdCount = results.filter(
      (item) => item.import_action === "created",
    ).length;
    const unchangedCount = results.filter(
      (item) => item.import_action === "unchanged",
    ).length;
    const messages = [];
    if (createdCount) messages.push(`已导入 ${createdCount} 份新资料`);
    if (updatedCount) messages.push(`已更新 ${updatedCount} 份资料并重新解析`);
    if (unchangedCount)
      messages.push(`${unchangedCount} 份文件内容未变化，未重复导入`);
    const ocrNeededCount = results.filter(
      (item) =>
        item.extension === ".pdf" &&
        item.parse_status === "failed" &&
        String(item.parse_error || "").includes("未提取到可复制文字"),
    ).length;
    if (ocrNeededCount)
      messages.push(
        `${ocrNeededCount} 份扫描 PDF 等待 OCR，请点击右侧“OCR 识别”`,
      );
    if (skippedCount) messages.push(`已取消 ${skippedCount} 份文件的导入`);
    successMessage.value = messages.length ? `${messages.join("；")}。` : "";
  } catch (error) {
    errorMessage.value =
      error.response?.data?.detail ?? "导入失败，请确认后端服务正在运行。";
  } finally {
    isImporting.value = false;
  }
}
function askSameNameConflict(file, candidates) {
  sameNameConflict.value = { file, candidates, selectedId: candidates[0].id };
  return new Promise((resolve) => {
    resolveConflictChoice = resolve;
  });
}
function closeSameNameConflict(action) {
  const choice =
    action === "replace"
      ? { action, fileId: sameNameConflict.value.selectedId }
      : { action };
  sameNameConflict.value = null;
  resolveConflictChoice?.(choice);
  resolveConflictChoice = null;
}
async function chooseFiles() {
  await addFiles((await window.desktopApi?.selectKnowledgeFiles?.()) || []);
}
function handleDrop() {
  isDragging.value = false;
  errorMessage.value =
    "请使用“选择文件”导入资料，以确保获取可靠的本地文件路径。";
}
async function showDetails(fileId) {
  isLoadingContent.value = true;
  errorMessage.value = "";
  try {
    selectedRecord.value = toViewModel(await getKnowledgeFile(fileId));
  } catch (error) {
    errorMessage.value =
      error.response?.data?.detail ??
      "无法读取解析内容，请确认后端服务正在运行。";
  } finally {
    isLoadingContent.value = false;
  }
}
async function removeFile(id) {
  const file = files.value.find((item) => item.id === id);
  const confirmed = window.confirm(
    `确定删除“${file?.name || "这份资料"}”吗？\n\n删除后会移除资料库中的本地副本和已解析内容，无法恢复。`,
  );
  if (!confirmed) return;
  try {
    await deleteKnowledgeFile(id);
    if (selectedRecord.value?.id === id) selectedRecord.value = null;
    await refreshFiles();
  } catch {
    errorMessage.value = "删除失败，请确认后端服务正在运行。";
  }
}
async function runOcr(file) {
  const approved = window.confirm(
    `将为“${file.name}”启动扫描 PDF 识别。首次使用会通过 PaddleOCR 官方源下载模型，下载和识别时间取决于网络与页数；模型仅保存到本机。是否继续？`,
  );
  if (!approved) return;
  ocrFileId.value = file.id;
  ocrProgress.value = { progress: 5, message: "正在准备扫描 PDF 识别…" };
  errorMessage.value = "";
  successMessage.value = "扫描 PDF OCR：5% · 正在准备 OCR 引擎…";
  try {
    const ocrRequest = ocrKnowledgePdf(file.id);
    let completed = false;
    ocrRequest.then(
      () => {
        completed = true;
      },
      () => {
        completed = true;
      },
    );
    while (!completed) {
      try {
        const status = await getKnowledgeOcrStatus(file.id);
        if (status.state !== "idle") {
          ocrProgress.value = status;
          successMessage.value = `扫描 PDF OCR：${status.progress}% · ${status.message}`;
        }
      } catch {
        /* OCR 主请求会给出最终错误；轮询失败时继续等待。 */
      }
      if (!completed)
        await new Promise((resolve) => window.setTimeout(resolve, 700));
    }
    await ocrRequest;
    await refreshFiles();
    successMessage.value = `已完成“${file.name}”的 OCR 识别，可以点击“查看”确认解析内容。`;
  } catch (error) {
    errorMessage.value =
      error.response?.data?.detail ??
      "OCR 识别失败，请确认后端服务正在运行后重试。";
  } finally {
    ocrFileId.value = "";
    ocrProgress.value = { progress: 0, message: "" };
  }
}
onMounted(refreshFiles);
</script>
<template>
  <main class="content">
    <header class="page-header">
      <div>
        <p class="page-kicker">KNOWLEDGE LIBRARY</p>
        <h1>我的资料库</h1>
        <p class="page-description">
          导入个人资料后，AI 会在填写表单时从中检索可靠信息。
        </p>
      </div>
      <div class="library-count">{{ selectedCount }} 份资料</div>
    </header>
    <section
      class="library-upload"
      :class="{ dragging: isDragging }"
      @dragenter.prevent="isDragging = true"
      @dragover.prevent="isDragging = true"
      @dragleave.prevent="isDragging = false"
      @drop.prevent="handleDrop"
    >
      <span class="upload-icon" aria-hidden="true">＋</span>
      <div>
        <h2>导入个人资料</h2>
        <p>
          支持 Word（.docx）、Excel（.xlsx）和 PDF（.pdf）；扫描 PDF
          导入后可选择 OCR 识别。
        </p>
      </div>
      <button
        class="button-like"
        type="button"
        :disabled="isImporting"
        @click="chooseFiles"
      >
        {{ isImporting ? "正在导入…" : "选择文件" }}
      </button>
    </section>
    <p v-if="errorMessage" class="api-error">{{ errorMessage }}</p>
    <p v-if="successMessage" class="api-success">{{ successMessage }}</p>
    <section v-if="ocrFileId" class="ocr-progress-panel" aria-live="polite">
      <div class="ocr-progress-heading">
        <span><i class="ocr-spinner" aria-hidden="true"></i>OCR 正在运行</span>
        <strong>{{ Math.max(5, ocrProgress.progress || 5) }}%</strong>
      </div>
      <div class="ocr-progress-track" :class="{ indeterminate: (ocrProgress.progress || 0) <= 12 }">
        <span :style="{ width: `${Math.max(5, ocrProgress.progress || 5)}%` }"></span>
      </div>
      <p>
        {{ (ocrProgress.progress || 0) <= 12
          ? '正在下载或加载 OCR 模型；此阶段暂无法获得精确下载百分比，请保持程序开启。'
          : ocrProgress.message }}
      </p>
    </section>
    <section class="library-section">
      <div class="section-heading">
        <h2>已导入的资料</h2>
        <p>
          {{ selectedCount }} 份文件会复制到本地资料库；扫描 PDF
          在首次识别时会从官方源下载本地 OCR 模型。
        </p>
      </div>
      <div v-if="files.length" class="file-table" role="table">
        <div class="table-row table-header">
          <span>文件</span><span>导入/更新时间</span><span>解析状态</span
          ><span>操作</span>
        </div>
        <div v-for="file in files" :key="file.id" class="table-row">
          <span class="file-cell"
            ><span class="document-icon" :class="file.color">{{
              file.icon
            }}</span
            ><span
              ><strong>{{ file.name }}</strong
              ><small>{{ file.kind }}</small></span
            ></span
          ><span>{{ file.importedAt }}</span
          ><span
            ><b
              class="status"
              :class="{
                pending: file.status === '待解析',
                failed: file.status === '解析失败',
              }"
              >{{ file.status }}</b
            ></span
          ><span class="row-actions"
            ><button
              v-if="file.extension === '.pdf' && file.parse_status === 'failed'"
              class="ocr-button"
              type="button"
              :disabled="Boolean(ocrFileId)"
              @click="runOcr(file)"
            >
              {{ ocrFileId === file.id ? "OCR 中…" : "OCR 识别" }}</button
            ><button
              class="view-button"
              type="button"
              @click="showDetails(file.id)"
            >
              查看</button
            ><button
              class="remove-button"
              type="button"
              @click="removeFile(file.id)"
            >
              删除
            </button></span
          >
        </div>
      </div>
      <div v-else class="empty-library">
        <span>▣</span>
        <h2>资料库还是空的</h2>
        <p>导入简历、个人信息表、项目经历或 PDF，开始建立你的本地知识库。</p>
      </div>
    </section>
    <div
      v-if="sameNameConflict"
      class="modal-backdrop"
      @click.self="closeSameNameConflict('cancel')"
    >
      <section
        class="content-modal same-name-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="same-name-dialog-title"
      >
        <header class="modal-header">
          <div>
            <p class="page-kicker">SAME FILE NAME</p>
            <h2 id="same-name-dialog-title">发现同名资料</h2>
          </div>
          <button
            class="close-button"
            type="button"
            aria-label="取消导入"
            @click="closeSameNameConflict('cancel')"
          >
            ×
          </button>
        </header>
        <p class="modal-meta">
          正在导入 <strong>{{ sameNameConflict.file.name }}</strong
          >，但资料库中已有同名、不同路径的文件。
        </p>
        <div class="same-name-options">
          <label
            v-for="candidate in sameNameConflict.candidates"
            :key="candidate.id"
            class="same-name-option"
            ><input
              v-model="sameNameConflict.selectedId"
              type="radio"
              :value="candidate.id"
            /><span
              ><strong>{{ candidate.name }}</strong
              ><small>原路径：{{ candidate.source_path }}</small></span
            ></label
          >
        </div>
        <p class="same-name-help">
          “更新已有资料”会用当前文件替换所选资料并重新解析；“作为新资料保留”会保留两份资料。
        </p>
        <footer class="modal-actions">
          <button
            class="secondary-button"
            type="button"
            @click="closeSameNameConflict('cancel')"
          >
            取消</button
          ><button
            class="secondary-button"
            type="button"
            @click="closeSameNameConflict('keep')"
          >
            作为新资料保留</button
          ><button
            class="primary-button"
            type="button"
            @click="closeSameNameConflict('replace')"
          >
            更新已有资料
          </button>
        </footer>
      </section>
    </div>
    <div
      v-if="isLoadingContent || selectedRecord"
      class="modal-backdrop"
      @click.self="selectedRecord = null"
    >
      <section
        class="content-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="content-dialog-title"
      >
        <header class="modal-header">
          <div>
            <p class="page-kicker">PARSED CONTENT</p>
            <h2 id="content-dialog-title">
              {{ isLoadingContent ? "正在读取解析内容…" : selectedRecord.name }}
            </h2>
          </div>
          <button
            class="close-button"
            type="button"
            aria-label="关闭"
            @click="selectedRecord = null"
          >
            ×
          </button>
        </header>
        <p v-if="!isLoadingContent" class="modal-meta">
          {{ selectedRecord.kind }} · {{ selectedRecord.status }} ·
          {{ selectedRecord.importedAt }}
        </p>
        <div v-if="isLoadingContent" class="content-loading">
          正在从本地资料库读取内容…
        </div>
        <pre v-else class="parsed-content">{{
          selectedRecord.extracted_text || "未提取到可显示的文字内容。"
        }}</pre>
      </section>
    </div>
  </main>
</template>
