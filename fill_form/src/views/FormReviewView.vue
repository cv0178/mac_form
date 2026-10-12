<script setup>
import { computed, onMounted, ref } from 'vue';
import { useRoute } from 'vue-router';
import { exportFilledForm, generateFormSuggestions, getForm, getNotificationSettings, getWritingStyles, listKnowledgeFiles } from '../services/localApi';

const route = useRoute();
const form = ref(null);
const editableValues = ref({});
const repeatRows = ref({});
const knowledgeNames = ref({});
const errorMessage = ref('');
const isLoading = ref(true);
const isGenerating = ref(false);
const isExporting = ref(false);
const exportResult = ref(null);
const showDetails = ref(false);

const writingStyle = ref('专业、简洁，突出事实与成果；不得编造资料库中不存在的事实。');

const fields = computed(() => form.value?.fields ?? []);
function rowsForField(field) {
  if (field.field_type !== 'repeat_table') return [];
  const rawValue = editableValues.value[field.id] || field.final_value || field.suggested_value || '';
  try {
    const rows = JSON.parse(rawValue);
    return Array.isArray(rows) ? rows : [];
  } catch {
    return [];
  }
}

function fieldHasValue(field) {
  return field.field_type === 'repeat_table'
    ? (repeatRows.value[field.id] ?? rowsForField(field)).some((row) => Object.values(row).some(Boolean))
    : Boolean(editableValues.value[field.id]);
}

const filledCount = computed(() => fields.value.filter(fieldHasValue).length);
const missingCount = computed(() => fields.value.length - filledCount.value);
// 需要补充的排在前面，方便用户先处理没填上的项
const orderedFields = computed(() =>
  [...fields.value].sort(
    (a, b) => Number(fieldHasValue(a)) - Number(fieldHasValue(b)),
  ),
);

function locationText(location) {
  if (location.file_type === 'xlsx') return `${location.sheet}!${location.cell}`;
  if ('table_index' in location) return `表格 ${location.table_index + 1} · 第 ${location.row_index + 1} 行`;
  return `段落 ${location.paragraph_index + 1}`;
}

function typeText(field) {
  if (field.needs_ai_generation) return 'AI 长文本';
  return { phone: '电话', email: '邮箱', date: '日期', choice: '单选项', repeat_table: '可重复表格', short_text: '短文本' }[field.field_type] ?? '文本';
}

function sourceNames(field) {
  return (field.source_ids ?? []).map((id) => knowledgeNames.value[id]).filter(Boolean);
}

function populateEditableValues(list, preserveValues = false) {
  for (const field of list) {
    if (!preserveValues || !editableValues.value[field.id]) {
      editableValues.value[field.id] = field.final_value || field.suggested_value || '';
    }
    if (field.field_type === 'repeat_table' && (!preserveValues || !repeatRows.value[field.id]?.length)) {
      repeatRows.value[field.id] = rowsForField(field);
    }
  }
}

function addRepeatRow(field) {
  const blankRow = Object.fromEntries((field.location.columns ?? []).map((column) => [column.key, '']));
  if (!repeatRows.value[field.id]) repeatRows.value[field.id] = [];
  repeatRows.value[field.id].push(blankRow);
}

function removeRepeatRow(field, index) {
  repeatRows.value[field.id]?.splice(index, 1);
}

function valueForField(field) {
  if (field.field_type === 'repeat_table') {
    return JSON.stringify((repeatRows.value[field.id] ?? []).filter((row) => Object.values(row).some(Boolean)));
  }
  return editableValues.value[field.id] || '';
}

async function loadKnowledgeNames() {
  try {
    const list = await listKnowledgeFiles();
    knowledgeNames.value = Object.fromEntries(list.map((item) => [item.id, item.name]));
  } catch {
    // 资料名称只用于展示建议来源，获取失败不影响填写流程
  }
}

async function loadWritingStyle() {
  try {
    const { styles, active_id: activeId } = await getWritingStyles();
    const active = styles.find((item) => item.id === activeId);
    if (active) writingStyle.value = active.content;
  } catch {
    // 读取失败时沿用默认写作风格
  }
}

async function loadForm() {
  try {
    form.value = await getForm(route.params.formId);
    populateEditableValues(form.value.fields);
  } catch (error) {
    errorMessage.value = error.response?.data?.detail ?? '无法读取表单分析结果，请确认 Python 后端正在运行。';
  } finally {
    isLoading.value = false;
  }
}

async function generateSuggestions() {
  if (isGenerating.value) return;
  isGenerating.value = true;
  errorMessage.value = '';
  try {
    form.value = await generateFormSuggestions(route.params.formId, writingStyle.value);
    populateEditableValues(form.value.fields, true);
  } catch (error) {
    errorMessage.value = error.response?.data?.detail ?? '自动填写失败，请检查模型服务配置。';
  } finally {
    isGenerating.value = false;
  }
}

async function exportForm() {
  isExporting.value = true;
  errorMessage.value = '';
  exportResult.value = null;
  try {
    exportResult.value = await exportFilledForm(
      route.params.formId,
      form.value.fields.map((field) => ({ field_id: field.id, value: valueForField(field) })),
    );
    // 通知失败不应影响已经成功生成的文件，例如用户在系统层面关闭了应用通知。
    try {
      const notificationSettings = await getNotificationSettings();
      if (notificationSettings.enabled) await window.desktopApi?.notifyExportComplete?.(exportResult.value.output_name);
    } catch {
      // 保持导出成功状态；系统通知只是额外提醒。
    }
    form.value = await getForm(route.params.formId);
  } catch (error) {
    errorMessage.value = error.response?.data?.detail ?? '导出失败，请确认 Python 后端正在运行。';
  } finally {
    isExporting.value = false;
  }
}

function openOutputFolder() {
  window.desktopApi?.showOutputInFolder?.(exportResult.value.output_path);
}

onMounted(async () => {
  await Promise.all([loadForm(), loadKnowledgeNames(), loadWritingStyle()]);
  // 进入页面就自动完成填写，不需要用户再点一次按钮
  if (fields.value.length && !fields.value.some((field) => field.suggested_value)) {
    await generateSuggestions();
  }
});
</script>

<template>
  <main class="content">
    <header class="page-header">
      <div>
        <p class="page-kicker">AUTO FILL</p>
        <h1>确认填写结果</h1>
        <p class="page-description">系统已根据你的资料库自动填写，请核对后导出。导出只会生成副本，不会修改原始表单。</p>
      </div>
      <RouterLink class="secondary-button" to="/">重新选择</RouterLink>
    </header>

    <section v-if="isLoading" class="analysis-state">正在读取表单结构…</section>
    <p v-else-if="errorMessage && !form" class="api-error">{{ errorMessage }}</p>

    <template v-else>
      <section v-if="isGenerating" class="auto-filling">
        <span class="spinner" aria-hidden="true"></span>
        <div>
          <strong>正在根据你的资料自动填写…</strong>
          <p>共 {{ fields.length }} 个填写项，正在逐项检索资料并生成内容，请稍候。</p>
        </div>
      </section>

      <section class="analysis-summary">
        <div><strong>{{ filledCount }}</strong><span>已自动填写</span></div>
        <div><strong>{{ missingCount }}</strong><span>需要你补充</span></div>
        <div><strong>{{ fields.length }}</strong><span>填写项总数</span></div>
      </section>

      <section class="ai-generation">
        <div>
          <h2>写作风格</h2>
          <p>用于生成较长内容（如个人简介、项目经历）。修改后可重新生成。</p>
          <textarea v-model="writingStyle" aria-label="AI 写作风格" rows="2" />
        </div>
        <button class="secondary-button" type="button" :disabled="isGenerating || !fields.length" @click="generateSuggestions">
          {{ isGenerating ? '正在生成…' : '重新生成' }}
        </button>
      </section>

      <p v-if="errorMessage" class="api-error">{{ errorMessage }}</p>

      <section class="review-list">
        <div class="review-heading review-heading-row">
          <div>
            <h2>填写内容</h2>
            <p>标记为「需补充」的项目未能自动填写，请手动补上。</p>
          </div>
          <button v-if="fields.length" class="text-button" type="button" @click="showDetails = !showDetails">
            {{ showDetails ? '隐藏识别详情' : '查看识别详情' }}
          </button>
        </div>

        <div v-if="fields.length" class="fill-list">
          <article
            v-for="field in orderedFields"
            :key="field.id"
            class="fill-item"
            :class="{ missing: !fieldHasValue(field) }"
          >
            <div class="fill-item-head">
              <div class="fill-item-title">
                <strong>{{ field.label }}</strong>
                <span v-if="fieldHasValue(field)" class="fill-badge auto">已自动填写</span>
                <span v-else class="fill-badge missing">需补充</span>
              </div>
              <small v-if="sourceNames(field).length" class="fill-source">
                来源：{{ sourceNames(field).join('、') }}
              </small>
              <small v-else-if="field.suggestion_reason" class="fill-source">{{ field.suggestion_reason }}</small>
            </div>
            <select
              v-if="field.field_type === 'choice'"
              v-model="editableValues[field.id]"
              :aria-label="`${field.label} 的填写选项`"
            >
              <option value="">请选择</option>
              <option v-for="option in field.location.checkbox_options" :key="option" :value="option">{{ option }}</option>
            </select>
            <div v-else-if="field.field_type === 'repeat_table'" class="repeat-table-editor">
              <p>每一行对应一个项目；项目数量超过 Word 模板预留行时，导出文件会自动增加表格行。</p>
              <div v-for="(row, rowIndex) in repeatRows[field.id]" :key="rowIndex" class="repeat-row">
                <label v-for="column in field.location.columns" :key="column.key">
                  <span>{{ column.label }}</span>
                  <textarea v-model="row[column.key]" :aria-label="`${column.label}（第 ${rowIndex + 1} 个项目）`" rows="2" />
                </label>
                <button class="text-button repeat-remove" type="button" @click="removeRepeatRow(field, rowIndex)">删除此项目</button>
              </div>
              <button class="secondary-button repeat-add" type="button" @click="addRepeatRow(field)">+ 新增项目</button>
            </div>
            <textarea
              v-else
              v-model="editableValues[field.id]"
              :aria-label="`${field.label} 的填写内容`"
              :rows="field.needs_ai_generation ? 5 : 2"
              placeholder="请填写内容"
            />
          </article>
        </div>

        <div v-else class="empty-library">
          <span>?</span>
          <h2>未识别到需要填写的位置</h2>
          <p>请确认表单中存在标题旁的空白单元格，或“标题：下划线”形式的填写区域。</p>
        </div>
      </section>

      <section v-if="showDetails && fields.length" class="detect-details">
        <div class="review-heading">
          <h2>识别详情</h2>
          <p>以下为系统识别待填位置时的依据，仅供排查使用。</p>
        </div>
        <div class="detail-table">
          <div class="detail-row detail-header">
            <span>原始标题</span><span>位置</span><span>识别方式</span><span>置信度</span><span>类型</span>
          </div>
          <div v-for="field in fields" :key="field.id" class="detail-row">
            <span><strong>{{ field.label }}</strong></span>
            <span>{{ locationText(field.location) }}</span>
            <span>{{ field.evidence }}</span>
            <span>{{ Math.round(field.confidence * 100) }}%</span>
            <span>{{ typeText(field) }}</span>
          </div>
        </div>
      </section>

      <footer class="export-actions">
        <span>将生成填写副本至本机 <code>outputs</code> 目录。</span>
        <button
          class="primary-button"
          type="button"
          :disabled="!fields.length || isExporting || isGenerating"
          @click="exportForm"
        >
          {{ isExporting ? '正在导出…' : '确认无误，导出文件' }}
        </button>
      </footer>

      <section v-if="exportResult" class="export-success">
        <div>
          <strong>填写结果已生成</strong>
          <p>{{ exportResult.output_name }} · 已填写 {{ exportResult.filled_field_count }} 个字段</p>
        </div>
        <button class="secondary-button" type="button" @click="openOutputFolder">打开所在文件夹</button>
      </section>
    </template>
  </main>
</template>
