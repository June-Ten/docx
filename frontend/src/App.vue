<template>
  <main class="page">
    <section class="toolbar">
      <div>
        <p class="eyebrow">DOCX Comment Editor</p>
        <h1>DOCX 划词高亮批注</h1>
        <p>前端使用 mammoth.js 预览 DOCX，web-highlighter 负责划词高亮，Python 后端生成带 Word 批注的文件。</p>
      </div>

      <label class="upload-button">
        <input type="file" accept=".docx" @change="handleFileChange" />
        选择 DOCX
      </label>
    </section>

    <section v-if="fileName" class="workspace">
      <article class="preview-card">
        <header>
          <div>
            <strong>{{ fileName }}</strong>
            <span>拖拽选择正文，选中后会自动高亮，再在右侧填写批注。</span>
          </div>
        </header>

        <div ref="previewRef" class="doc-preview" v-html="html"></div>
      </article>

      <aside class="panel">
        <section class="current">
          <h2>当前选区</h2>
          <p v-if="!draft" class="muted">请先在预览区划词。</p>
          <blockquote v-else>{{ draft.selectedText }}</blockquote>
        </section>

        <label class="field">
          批注内容
          <textarea v-model="commentText" rows="5" placeholder="请输入批注内容"></textarea>
        </label>

        <label class="field">
          作者
          <input v-model="author" placeholder="Reviewer" />
        </label>

        <button type="button" :disabled="!canSaveComment" @click="saveComment">
          保存当前批注
        </button>

        <section class="comments">
          <h2>批注列表</h2>
          <p v-if="comments.length === 0" class="muted">暂无批注。</p>
          <article v-for="comment in comments" :key="comment.id" class="comment-item">
            <strong>{{ comment.selectedText }}</strong>
            <p>{{ comment.comment }}</p>
            <small>第 {{ comment.occurrence + 1 }} 次出现 · {{ comment.author }}</small>
            <button type="button" class="text-button" @click="removeComment(comment.id)">删除</button>
          </article>
        </section>

        <button class="primary" type="button" :disabled="comments.length === 0 || exporting" @click="exportDocx">
          {{ exporting ? '生成中...' : '生成批注文件' }}
        </button>
      </aside>
    </section>

    <section v-else class="empty-state">
      <p>上传一个 `.docx` 文件后开始预览和批注。</p>
    </section>

    <p v-if="message" class="message">{{ message }}</p>
  </main>
</template>

<script setup>
import Highlighter from 'web-highlighter';
import { computed, nextTick, onBeforeUnmount, ref } from 'vue';
import mammoth from 'mammoth/mammoth.browser';

const API_BASE = import.meta.env.VITE_API_BASE ?? '';

const previewRef = ref(null);
const sourceFile = ref(null);
const fileName = ref('');
const html = ref('');
const comments = ref([]);
const draft = ref(null);
const commentText = ref('');
const author = ref('Reviewer');
const message = ref('');
const exporting = ref(false);

let highlighter = null;

const canSaveComment = computed(() => draft.value && commentText.value.trim());

async function handleFileChange(event) {
  const file = event.target.files?.[0];
  event.target.value = '';
  if (!file) return;

  sourceFile.value = file;
  fileName.value = file.name;
  comments.value = [];
  draft.value = null;
  commentText.value = '';
  message.value = '正在解析 DOCX...';

  try {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer });
    html.value = result.value;
    message.value = result.messages.length ? 'DOCX 已预览，部分格式可能被 mammoth 简化。' : '';
    await nextTick();
    setupHighlighter();
  } catch (error) {
    message.value = error.message || 'DOCX 解析失败';
  }
}

function setupHighlighter() {
  destroyHighlighter();
  if (!previewRef.value) return;

  highlighter = new Highlighter({
    $root: previewRef.value,
    exceptSelectors: ['button', 'input', 'textarea']
  });

  highlighter.on(Highlighter.event.CREATE, ({ sources }) => {
    const source = sources?.[0];
    const selectedText = source?.text?.replace(/\s+/g, ' ').trim();
    if (!source || !selectedText) return;

    draft.value = {
      id: source.id,
      source,
      selectedText,
      occurrence: countOccurrenceBeforeSelection(selectedText)
    };
    commentText.value = '';
    message.value = '';
  });

  highlighter.run();
}

function destroyHighlighter() {
  highlighter = null;
}

function countOccurrenceBeforeSelection(selectedText) {
  const selection = window.getSelection();
  const root = previewRef.value;
  if (!selection || !root || selection.rangeCount === 0) return 0;

  const range = selection.getRangeAt(0);
  const beforeRange = document.createRange();
  beforeRange.selectNodeContents(root);
  beforeRange.setEnd(range.startContainer, range.startOffset);

  const beforeText = normalizeText(beforeRange.toString());
  const needle = normalizeText(selectedText);
  let count = 0;
  let index = beforeText.indexOf(needle);

  while (index >= 0) {
    count += 1;
    index = beforeText.indexOf(needle, index + needle.length);
  }

  return count;
}

function saveComment() {
  if (!canSaveComment.value) return;

  comments.value.push({
    id: draft.value.id,
    selectedText: draft.value.selectedText,
    occurrence: draft.value.occurrence,
    comment: commentText.value.trim(),
    author: author.value.trim() || 'Reviewer',
    initials: getInitials(author.value),
    highlightSource: draft.value.source
  });

  draft.value = null;
  commentText.value = '';
}

function removeComment(id) {
  comments.value = comments.value.filter((comment) => comment.id !== id);
}

async function exportDocx() {
  if (!sourceFile.value) return;

  exporting.value = true;
  message.value = '';

  try {
    const formData = new FormData();
    formData.append('file', sourceFile.value);
    formData.append('comments', JSON.stringify(comments.value));

    const response = await fetch(`${API_BASE}/api/comments/generate`, {
      method: 'POST',
      body: formData
    });

    if (!response.ok) {
      throw new Error(await readError(response));
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName.value.replace(/\.docx$/i, '-comments.docx');
    link.click();
    URL.revokeObjectURL(url);
  } catch (error) {
    message.value = error.message || '生成失败';
  } finally {
    exporting.value = false;
  }
}

function normalizeText(value) {
  return value.replace(/\s+/g, ' ').trim();
}

function getInitials(value) {
  const normalized = value.trim();
  if (!normalized) return 'RV';
  return normalized
    .split(/\s+/)
    .map((part) => part[0])
    .join('')
    .slice(0, 9)
    .toUpperCase();
}

async function readError(response) {
  const text = await response.text();
  try {
    const payload = JSON.parse(text);
    return payload.message || response.statusText || '生成失败';
  } catch {
    return text || response.statusText || '生成失败';
  }
}

onBeforeUnmount(() => {
  destroyHighlighter();
});
</script>
