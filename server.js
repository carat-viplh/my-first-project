/**
 * 批量文件修改对比 & 自动修改工具 - 后端服务
 * 技术栈：Node.js + Express + multer + diff + mammoth + pdf-parse + jszip
 */
import express from 'express';
import multer from 'multer';
import path from 'path';
import fs from 'fs/promises';
import mammoth from 'mammoth';
import pdfParse from 'pdf-parse';
import * as diffLib from 'diff';
import JSZip from 'jszip';

const app = express();
const UPLOAD_DIR = 'uploads';
const PORT = 3000;

const SUPPORTED_EXT = new Set(['.txt', '.docx', '.doc', '.pdf']);

// ---------- 初始化 ----------
app.use(express.json({ limit: '50mb' }));
app.use(express.static('.'));

const storage = multer.diskStorage({
  destination: UPLOAD_DIR,
  filename: (_req, file, cb) => {
    const safe = file.originalname.replace(/[^\w.\-()\u4e00-\u9fa5]/g, '_');
    cb(null, `${Date.now()}-${safe}`);
  },
});
const upload = multer({
  storage,
  limits: { fileSize: 30 * 1024 * 1024 },
});

(async () => {
  try {
    await fs.access(UPLOAD_DIR);
  } catch {
    await fs.mkdir(UPLOAD_DIR, { recursive: true });
  }
})();

// ---------- 1. 文件解析：提取纯文本 ----------
/**
 * 从 txt / docx / pdf 中提取纯文本（去除格式，用于对比）
 */
async function extractText(filePath, originalName) {
  const ext = path.extname(originalName).toLowerCase();
  if (!SUPPORTED_EXT.has(ext)) {
    throw new Error(`不支持的文件格式: ${ext}，请使用 txt、docx、doc、pdf`);
  }
  if (ext === '.txt') {
    return await fs.readFile(filePath, 'utf-8');
  }
  if (ext === '.docx') {
    const result = await mammoth.extractRawText({ path: filePath });
    return result.value;
  }
  if (ext === '.doc') {
    throw new Error('旧版 .doc 请先另存为 .docx 后再上传');
  }
  if (ext === '.pdf') {
    const buffer = await fs.readFile(filePath);
    const data = await pdfParse(buffer);
    return data.text || '';
  }
  throw new Error(`无法解析: ${ext}`);
}

/** 是否支持自动写回修改 */
function canAutoFix(filename) {
  const ext = path.extname(filename).toLowerCase();
  return ext === '.txt' || ext === '.docx';
}

function splitLines(text) {
  return text.split(/\r?\n/);
}

function countLines(text) {
  if (!text) return 0;
  const trimmed = text.replace(/\r?\n$/, '');
  if (!trimmed) return 0;
  return splitLines(trimmed).length;
}

// ---------- 2. 对比算法：生成替换规则与带行号的展示数据 ----------
/**
 * 从 diff 结果中提取「旧文本 -> 新文本」替换规则
 */
function buildReplacements(diffParts) {
  const rules = [];
  for (let i = 0; i < diffParts.length; i++) {
    const part = diffParts[i];
    const next = diffParts[i + 1];

    // 修改：removed 紧接 added
    if (part.removed && next?.added) {
      const oldBlock = part.value.replace(/\r?\n$/, '');
      const newBlock = next.value.replace(/\r?\n$/, '');
      const oldLines = splitLines(oldBlock);
      const newLines = splitLines(newBlock);

      if (oldLines.length === newLines.length && oldLines.length > 1) {
        oldLines.forEach((oldLine, idx) => {
          const newLine = newLines[idx];
          if (oldLine.trim() && oldLine.trim() !== newLine.trim()) {
            rules.push({
              type: 'modify',
              oldText: oldLine.trim(),
              newText: newLine.trim(),
            });
          }
        });
      } else if (oldBlock.trim()) {
        rules.push({
          type: 'modify',
          oldText: oldBlock.trim(),
          newText: newBlock.trim(),
        });
      }
      i++;
      continue;
    }

    // 仅删除
    if (part.removed && part.value.trim()) {
      splitLines(part.value).forEach((line) => {
        const t = line.trim();
        if (t) rules.push({ type: 'remove', oldText: t, newText: '' });
      });
      continue;
    }

    // 仅新增（模板有、原文无）— 批量文件通常无需按行替换，跳过
  }

  // 去重：相同 oldText 保留最后一次 newText
  const map = new Map();
  for (const r of rules) {
    if (r.oldText) map.set(r.oldText, r);
  }
  return [...map.values()];
}

/**
 * 生成带行号的差异展示列表（删除红 / 新增绿 / 修改黄）
 */
function buildDiffDisplay(diffParts) {
  const rows = [];
  let originLine = 1;
  let templateLine = 1;

  for (let i = 0; i < diffParts.length; i++) {
    const part = diffParts[i];
    const next = diffParts[i + 1];

    if (part.removed && next?.added) {
      const oldLines = splitLines(part.value.replace(/\r?\n$/, ''));
      const newLines = splitLines(next.value.replace(/\r?\n$/, ''));
      const maxLen = Math.max(oldLines.length, newLines.length, 1);

      for (let j = 0; j < maxLen; j++) {
        const oldT = (oldLines[j] ?? '').trimEnd();
        const newT = (newLines[j] ?? '').trimEnd();
        if (!oldT && !newT) continue;
        rows.push({
          kind: 'modify',
          originLine: originLine + j,
          templateLine: templateLine + j,
          oldText: oldT,
          newText: newT,
        });
      }
      originLine += countLines(part.value);
      templateLine += countLines(next.value);
      i++;
      continue;
    }

    if (part.removed) {
      const lines = splitLines(part.value.replace(/\r?\n$/, ''));
      lines.forEach((line, j) => {
        if (!line.trim() && lines.length > 1) return;
        rows.push({
          kind: 'remove',
          originLine: originLine + j,
          oldText: line,
        });
      });
      originLine += countLines(part.value);
      continue;
    }

    if (part.added) {
      const lines = splitLines(part.value.replace(/\r?\n$/, ''));
      lines.forEach((line, j) => {
        if (!line.trim() && lines.length > 1) return;
        rows.push({
          kind: 'add',
          templateLine: templateLine + j,
          newText: line,
        });
      });
      templateLine += countLines(part.value);
      continue;
    }

    // 未变化：只推进行号，不展示（避免刷屏）
    originLine += countLines(part.value);
    templateLine += countLines(part.value);
  }

  return rows;
}

// ---------- 3. 批量遍历：检测子文件中需修改的行 ----------
/**
 * 在单个文件文本中查找需要替换的旧内容
 */
function analyzeOneFile(filename, filePath, fileText, replacements) {
  const lines = splitLines(fileText);
  const analysis = {
    filename,
    filePath,
    storedName: path.basename(filePath),
    canAutoFix: canAutoFix(filename),
    needFix: false,
    details: [],
  };

  const seen = new Set();

  for (const rule of replacements) {
    if (!rule.oldText) continue;

    lines.forEach((line, index) => {
      if (!line.includes(rule.oldText)) return;

      const key = `${index + 1}:${rule.oldText}`;
      if (seen.has(key)) return;
      seen.add(key);

      analysis.needFix = true;
      analysis.details.push({
        line: index + 1,
        oldText: rule.oldText,
        newText: rule.newText ?? '',
        type: rule.type,
        lineContent: line.trim(),
      });
    });
  }

  return analysis;
}

// ---------- 4. 自动修改：按行替换 ----------
function escapeXml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/**
 * 修改 .txt：按行将 oldText 替换为 newText
 */
async function applyFixToTxt(filePath, details) {
  let content = await fs.readFile(filePath, 'utf-8');
  const lines = splitLines(content);

  for (const detail of details) {
    const idx = detail.line - 1;
    if (idx < 0 || idx >= lines.length) continue;
    if (lines[idx].includes(detail.oldText)) {
      lines[idx] = lines[idx].replace(detail.oldText, detail.newText);
    }
  }

  const ending = content.includes('\r\n') ? '\r\n' : '\n';
  await fs.writeFile(filePath, lines.join(ending), 'utf-8');
}

/**
 * 修改 .docx：在 document.xml 中做文本替换（简单场景有效）
 */
async function applyFixToDocx(filePath, details) {
  const buffer = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);
  const xmlFile = zip.file('word/document.xml');
  if (!xmlFile) throw new Error('无效的 docx 文件结构');

  let xml = await xmlFile.async('string');

  for (const detail of details) {
    if (!detail.oldText) continue;
    const oldEsc = escapeXml(detail.oldText);
    const newEsc = escapeXml(detail.newText);
    if (xml.includes(oldEsc)) {
      xml = xml.split(oldEsc).join(newEsc);
    } else if (xml.includes(detail.oldText)) {
      xml = xml.split(detail.oldText).join(detail.newText);
    }
  }

  zip.file('word/document.xml', xml);
  const out = await zip.generateAsync({ type: 'nodebuffer' });
  await fs.writeFile(filePath, out);
}

async function applyFixToFile(task) {
  const ext = path.extname(task.filename).toLowerCase();
  if (ext === '.txt') {
    await applyFixToTxt(task.filePath, task.details);
    return;
  }
  if (ext === '.docx') {
    await applyFixToDocx(task.filePath, task.details);
    return;
  }
  throw new Error(`${task.filename} 不支持自动写回，请使用 txt 或 docx`);
}

async function safeUnlink(filePath) {
  try {
    await fs.unlink(filePath);
  } catch {
    /* 忽略 */
  }
}

// ---------- API 路由 ----------

/** 对比原文件与模板文件 */
app.post(
  '/api/compare-template',
  upload.fields([
    { name: 'originFile', maxCount: 1 },
    { name: 'templateFile', maxCount: 1 },
  ]),
  async (req, res) => {
    let originPath;
    let templatePath;
    try {
      if (!req.files?.originFile?.[0] || !req.files?.templateFile?.[0]) {
        return res.status(400).json({ success: false, error: '请上传原始文件和模板文件' });
      }

      originPath = req.files.originFile[0].path;
      const originName = req.files.originFile[0].originalname;
      templatePath = req.files.templateFile[0].path;
      const templateName = req.files.templateFile[0].originalname;

      const originText = await extractText(originPath, originName);
      const templateText = await extractText(templatePath, templateName);

      const diffParts = diffLib.diffLines(originText, templateText);
      const replacements = buildReplacements(diffParts);
      const diffDisplay = buildDiffDisplay(diffParts);

      await safeUnlink(originPath);
      await safeUnlink(templatePath);
      originPath = templatePath = null;

      res.json({
        success: true,
        diff: diffParts,
        diffDisplay,
        replacements,
        stats: {
          modifyCount: diffDisplay.filter((d) => d.kind === 'modify').length,
          removeCount: diffDisplay.filter((d) => d.kind === 'remove').length,
          addCount: diffDisplay.filter((d) => d.kind === 'add').length,
          ruleCount: replacements.length,
        },
      });
    } catch (error) {
      if (originPath) await safeUnlink(originPath);
      if (templatePath) await safeUnlink(templatePath);
      res.status(500).json({ success: false, error: error.message });
    }
  }
);

/** 批量遍历分析 */
app.post('/api/analyze-batch', upload.array('batchFiles', 200), async (req, res) => {
  try {
    if (!req.files?.length) {
      return res.status(400).json({ success: false, error: '请上传至少一个批量文件' });
    }

    const replacements = JSON.parse(req.body.replacements || '[]');
    if (!replacements.length) {
      return res.status(400).json({ success: false, error: '没有可用的模板差异规则，请先完成模板对比' });
    }

    const results = [];
    for (const file of req.files) {
      const fileText = await extractText(file.path, file.originalname);
      results.push(analyzeOneFile(file.originalname, file.path, fileText, replacements));
    }

    res.json({ success: true, results });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

/** 一键批量自动修改 */
app.post('/api/apply-fix', async (req, res) => {
  try {
    const { fileTasks } = req.body;
    if (!fileTasks?.length) {
      return res.status(400).json({ success: false, error: '没有待修改的文件' });
    }

    const fixed = [];
    const skipped = [];

    for (const task of fileTasks) {
      try {
        const stat = await fs.stat(task.filePath).catch(() => null);
        if (!stat?.isFile()) {
          skipped.push({ filename: task.filename, reason: '文件不存在' });
          continue;
        }
        if (!task.canAutoFix) {
          skipped.push({ filename: task.filename, reason: '仅 txt/docx 支持自动写回' });
          continue;
        }
        await applyFixToFile(task);
        fixed.push({
          filename: task.filename,
          storedName: task.storedName || path.basename(task.filePath),
        });
      } catch (err) {
        skipped.push({ filename: task.filename, reason: err.message });
      }
    }

    res.json({
      success: true,
      message: `已完成 ${fixed.length} 个文件修改${skipped.length ? `，${skipped.length} 个跳过` : ''}`,
      fixed,
      skipped,
    });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

/** 下载修改后的文件 */
app.get('/api/download/:storedName', async (req, res) => {
  try {
    const storedName = path.basename(req.params.storedName);
    const filePath = path.join(UPLOAD_DIR, storedName);
    const resolved = path.resolve(filePath);
    if (!resolved.startsWith(path.resolve(UPLOAD_DIR))) {
      return res.status(403).json({ error: '非法路径' });
    }
    await fs.access(resolved);
    res.download(resolved);
  } catch {
    res.status(404).json({ error: '文件不存在' });
  }
});

app.listen(PORT, () => {
  console.log(`✅ 批量文件对比工具已启动: http://localhost:${PORT}`);
});
