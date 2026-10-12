<script setup>
import { computed, nextTick, onMounted, ref } from 'vue';
import { exportFilledForm, generateFormSuggestions, getNotificationSettings, getWritingStyles, importFormFile, listKnowledgeFiles, saveWritingStyles } from '../services/localApi';

const selectedFile = ref(null);
const isDragging = ref(false);
const isAnalyzing = ref(false);
const progressText = ref('');
const errorMessage = ref('');
const result = ref(null);
const resultFields = ref([]);
const resultCard = ref(null);
const showFilled = ref(false);

const knowledgeFiles = ref([]);
const knowledgeNames = ref({});
const knowledgeUnavailable = ref(false);
const writingStyles = ref([]);
const activeStyleId = ref('');
const isAddingStyle = ref(false);
const isEditingStyle = ref(false);
const newStyleName = ref('');
const newStyleContent = ref('');
const isSavingStyle = ref(false);

const selectedFileName = computed(() => selectedFile.value?.name ?? '');
const hasKnowledge = computed(() => knowledgeFiles.value.length > 0);
const filledFields = computed(() => resultFields.value.filter((field) => field.suggested_value));
const missingFields = computed(() => resultFields.value.filter((field) => !field.suggested_value));
const activeWritingStyle = computed(() => writingStyles.value.find((item) => item.id === activeStyleId.value));

function setSelectedFile(file) {
  if (file) {
    selectedFile.value = file;
    isDragging.value = false;
    errorMessage.value = '';
    result.value = null;
    resultFields.value = [];
  }
}

function sourceNames(field) {
  return (field.source_ids ?? []).map((id) => knowledgeNames.value[id]).filter(Boolean);
}

async function chooseFormFile() {
  setSelectedFile(await window.desktopApi?.selectFormFile?.());
}

function handleDrop() {
  isDragging.value = false;
  errorMessage.value = '请使用“选择文件”选择表单，以确保获取可靠的本地文件路径。';
}

function openOutputFolder() {
  window.desktopApi?.showOutputInFolder?.(result.value.output_path);
}

async function loadKnowledgeFiles() {
  try {
    knowledgeFiles.value = await listKnowledgeFiles();
    knowledgeNames.value = Object.fromEntries(knowledgeFiles.value.map((item) => [item.id, item.name]));
    knowledgeUnavailable.value = false;
  } catch {
    knowledgeFiles.value = [];
    knowledgeUnavailable.value = true;
  }
}

async function loadWritingStyles() {
  try {
    const data = await getWritingStyles();
    writingStyles.value = data.styles;
    activeStyleId.value = data.active_id;
  } catch {
    // 读取失败时由后端退回默认写作风格
  }
}

async function changeActiveStyle() {
  try {
    await saveWritingStyles({ styles: writingStyles.value, active_id: activeStyleId.value });
  } catch {
    // 保存失败不影响本次填写，后端会沿用上一次选中的风格
  }
}

function newStyleId() {
  return globalThis.crypto?.randomUUID?.() ?? `style-${Date.now()}`;
}

function cancelNewStyle() {
  isAddingStyle.value = false;
  isEditingStyle.value = false;
  newStyleName.value = '';
  newStyleContent.value = '';
}

function editActiveStyle() {
  if (!activeWritingStyle.value) return;
  isAddingStyle.value = true;
  isEditingStyle.value = true;
  newStyleName.value = activeWritingStyle.value.name;
  newStyleContent.value = activeWritingStyle.value.content;
}

async function addWritingStyle() {
  const name = newStyleName.value.trim();
  const content = newStyleContent.value.trim();
  if (!name || !content) {
    errorMessage.value = '请填写写作风格名称和风格要求。';
    return;
  }
  isSavingStyle.value = true;
  errorMessage.value = '';
  try {
    const style = isEditingStyle.value
      ? { ...activeWritingStyle.value, name, content }
      : { id: newStyleId(), name, content };
    const styles = isEditingStyle.value
      ? writingStyles.value.map((item) => (item.id === style.id ? style : item))
      : [...writingStyles.value, style];
    const data = await saveWritingStyles({ styles, active_id: style.id });
    writingStyles.value = data.styles;
    activeStyleId.value = data.active_id;
    cancelNewStyle();
  } catch (error) {
    errorMessage.value = error.response?.data?.detail ?? '新增写作风格失败，请检查后端服务。';
  } finally {
    isSavingStyle.value = false;
  }
}

async function deleteActiveStyle() {
  if (!activeWritingStyle.value) return;
  if (writingStyles.value.length <= 1) {
    errorMessage.value = '至少需要保留一条写作风格。';
    return;
  }
  if (!window.confirm(`确定删除「${activeWritingStyle.value.name}」吗？`)) return;
  errorMessage.value = '';
  try {
    const styles = writingStyles.value.filter((item) => item.id !== activeStyleId.value);
    const data = await saveWritingStyles({ styles, active_id: styles[0].id });
    writingStyles.value = data.styles;
    activeStyleId.value = data.active_id;
  } catch (error) {
    errorMessage.value = error.response?.data?.detail ?? '删除写作风格失败，请检查后端服务。';
  }
}

async function startFilling() {
  if (!selectedFile.value?.path) return;
  isAnalyzing.value = true;
  errorMessage.value = '';
  result.value = null;
  resultFields.value = [];
  try {
    progressText.value = '正在识别表单结构…';
    const form = await importFormFile(selectedFile.value.path);
    if (!form.fields.length) {
      errorMessage.value = '没有在表单中识别到可填写的位置，请确认表格中存在“标题 + 相邻空白单元格”。';
      return;
    }
    progressText.value = `正在根据你的资料自动填写 ${form.fields.length} 个位置…`;
    // 用下拉里选中的那条风格；取不到时由后端使用当前风格
    const style = writingStyles.value.find((item) => item.id === activeStyleId.value);
    const filled = await generateFormSuggestions(form.id, style?.content);
    progressText.value = '正在生成填写结果文件…';
    const exported = await exportFilledForm(
      filled.id,
      filled.fields.map((field) => ({ field_id: field.id, value: field.suggested_value || '' })),
    );
    resultFields.value = filled.fields;
    result.value = { ...exported, formId: filled.id };
    // 导出已完成后再通知；若用户关闭系统通知，也不能影响文件生成。
    try {
      const notificationSettings = await getNotificationSettings();
      if (notificationSettings.enabled) await window.desktopApi?.notifyExportComplete?.(exported.output_name);
    } catch {
      // 系统通知只作为补充提示。
    }
    showFilled.value = false;
    await nextTick();
    resultCard.value?.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  } catch (error) {
    errorMessage.value = error.response?.data?.detail ?? '自动填写失败，请确认 Python 后端正在运行，并检查「设置」。';
  } finally {
    isAnalyzing.value = false;
    progressText.value = '';
  }
}

function resetForNext() {
  selectedFile.value = null;
  result.value = null;
  resultFields.value = [];
  errorMessage.value = '';
  showFilled.value = false;
}

onMounted(() => {
  loadKnowledgeFiles();
  loadWritingStyles();
});
</script>

<template>
  <main class="content">
    <header class="page-header">
      <div>
        <p class="page-kicker">AI FORM ASSISTANT</p>
        <h1>新建智能填写</h1>
        <p class="page-description">选择待填写的文档，系统会自动识别待填位置，根据你的资料库完成填写并导出结果。</p>
      </div>
      <div class="local-mode"><span class="avatar">你</span>本地模式</div>
    </header>

    <section class="task-grid" aria-label="创建填写任务">
      <section
        class="file-dropzone"
        :class="{ dragging: isDragging, selected: selectedFileName }"
        @dragenter.prevent="isDragging = true"
        @dragover.prevent="isDragging = true"
        @dragleave.prevent="isDragging = false"
        @drop.prevent="handleDrop"
      >
        <span class="upload-icon" aria-hidden="true">⇧</span>
        <template v-if="selectedFileName">
          <h2>{{ selectedFileName }}</h2>
          <p>文件已选择，点击可重新选择。</p>
        </template>
        <template v-else>
          <h2>上传待填写的表单</h2>
          <p>请选择 Word（.docx）或 Excel（.xlsx）表单文件。</p>
        </template>
        <button class="button-like" type="button" @click="chooseFormFile">选择文件</button>
      </section>

      <section class="knowledge-card">
        <div class="card-heading">
          <div>
            <h2>本次将使用的资料</h2>
            <p v-if="hasKnowledge">AI 仅从这 {{ knowledgeFiles.length }} 份资料中提取信息。</p>
            <p v-else>AI 需要参考你的资料才能填写内容。</p>
          </div>
          <RouterLink class="text-button" to="/knowledge">管理</RouterLink>
        </div>

        <template v-if="hasKnowledge">
          <ul class="knowledge-list">
            <li v-for="file in knowledgeFiles.slice(0, 4)" :key="file.id">
              <span class="document-icon" :class="file.extension === '.docx' ? 'word' : 'excel'">
                {{ file.extension === '.docx' ? 'W' : 'X' }}
              </span>
              <span class="file-title">{{ file.name }}</span>
            </li>
          </ul>
          <p v-if="knowledgeFiles.length > 4" class="knowledge-more">
            还有 {{ knowledgeFiles.length - 4 }} 份资料未展示
          </p>
        </template>

        <div v-else class="knowledge-empty">
          <p v-if="knowledgeUnavailable">无法连接本地服务，请先启动 Python 后端。</p>
          <template v-else>
            <p>资料库还是空的，暂时无法自动填写。</p>
            <RouterLink class="button-like" to="/knowledge">去导入资料</RouterLink>
          </template>
        </div>
      </section>
    </section>

    <section v-if="writingStyles.length" class="start-style-card">
      <div class="start-style-heading">
        <div>
          <h2>本次写作风格</h2>
          <p>用于生成工作心得、获奖感言、项目介绍等长文本内容。</p>
        </div>
      </div>
      <div class="start-style-controls">
        <label>
          <span>选择风格</span>
          <select v-model="activeStyleId" @change="changeActiveStyle">
            <option v-for="style in writingStyles" :key="style.id" :value="style.id">{{ style.name }}</option>
          </select>
        </label>
        <button v-if="!isAddingStyle" class="secondary-button" type="button" @click="isAddingStyle = true">+ 新增风格</button>
      </div>
      <section v-if="activeWritingStyle && !isAddingStyle" class="selected-style-detail">
        <div>
          <strong>{{ activeWritingStyle.name }}</strong>
          <p>{{ activeWritingStyle.content }}</p>
        </div>
        <div class="selected-style-actions">
          <button class="view-button" type="button" @click="editActiveStyle">编辑</button>
          <button class="remove-button" type="button" :disabled="writingStyles.length <= 1" @click="deleteActiveStyle">删除</button>
        </div>
      </section>
      <form v-if="isAddingStyle" class="start-style-editor" @submit.prevent="addWritingStyle">
        <input v-model.trim="newStyleName" type="text" :placeholder="isEditingStyle ? '风格名称' : '风格名称，例如：正式项目申报'" aria-label="新写作风格名称" />
        <textarea v-model.trim="newStyleContent" rows="3" placeholder="描述这套风格的语气、长度和表达要求" aria-label="新写作风格要求" />
        <div>
          <button class="primary-button" type="submit" :disabled="isSavingStyle">{{ isSavingStyle ? '正在保存…' : isEditingStyle ? '保存修改' : '保存并使用' }}</button>
          <button class="text-button" type="button" @click="cancelNewStyle">取消</button>
        </div>
      </form>
    </section>

    <section v-if="isAnalyzing" class="auto-filling">
      <span class="spinner" aria-hidden="true"></span>
      <div>
        <strong>正在自动填写…</strong>
        <p>{{ progressText }}</p>
      </div>
    </section>

    <p v-if="errorMessage" class="api-error">{{ errorMessage }}</p>

    <section v-if="result" ref="resultCard" class="result-card">
      <div class="result-head">
        <div>
          <strong>填写完成</strong>
          <p>{{ result.output_name }}</p>
          <p class="result-stat">
            已填写 {{ result.filled_field_count }} 项<template v-if="missingFields.length"> · <b>{{ missingFields.length }} 项未填</b></template>
          </p>
        </div>
        <button class="secondary-button" type="button" @click="openOutputFolder">打开所在文件夹</button>
      </div>

      <section v-if="missingFields.length" class="result-block missing-block">
        <h2>未填写 {{ missingFields.length }} 项</h2>
        <ul>
          <li v-for="field in missingFields" :key="field.id">
            <strong>{{ field.label }}</strong>
            <span>{{ field.suggestion_reason || '资料库中没有能匹配的内容' }}</span>
          </li>
        </ul>
      </section>

      <section v-if="filledFields.length" class="result-block">
        <button class="text-button" type="button" @click="showFilled = !showFilled">
          {{ showFilled ? '收起已填写内容' : `查看已填写内容（${filledFields.length} 项）` }}
        </button>
        <ul v-if="showFilled" class="filled-list">
          <li v-for="field in filledFields" :key="field.id">
            <strong>{{ field.label }}</strong>
            <span>{{ field.suggested_value }}</span>
            <small v-if="sourceNames(field).length">来源：{{ sourceNames(field).join('、') }}</small>
          </li>
        </ul>
      </section>

      <div class="result-foot">
        <span>结果未经人工核对，正式提交前建议打开文件确认。</span>
      </div>
    </section>

    <footer class="actions">
      <button v-if="result" class="secondary-button" type="button" @click="resetForNext">再填一份</button>
      <button
        v-else
        class="primary-button"
        type="button"
        :disabled="!selectedFile || isAnalyzing || !hasKnowledge"
        @click="startFilling"
      >
        {{ isAnalyzing ? '正在填写…' : '开始填写' }} <span>→</span>
      </button>
    </footer>
  </main>
</template>
