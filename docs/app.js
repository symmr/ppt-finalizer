const DEFAULT_FONT = "Noto Sans JP";
const CUSTOM_VALUE = "__custom__";
const PPTX_TYPES = [{
  description: "PowerPoint",
  accept: { "application/vnd.openxmlformats-officedocument.presentationml.presentation": [".pptx"] },
}];
const PRESET_FONTS = [
  "Noto Sans JP",
  "Meiryo",
  "Yu Gothic UI",
  "Segoe UI",
  "Arial",
  "Calibri",
];
const PRESET_SET = new Set(PRESET_FONTS);
const TYPEFACE_CAPTURE_RE = /typeface="([^"]*)"/g;
const TYPEFACE_RE = /typeface="[^"]*"/g;
const TITLE_PH_RE = /<p:ph[^>]*\btype="(?:title|ctrTitle|subTitle)"/;
const BODY_PH_RE = /<p:ph[^>]*\btype="body"/;
const STYLE_BLOCK_RES = [
  { re: /<(?:p|a):titleStyle\b[\s\S]*?<\/(?:p|a):titleStyle>/g, role: "title" },
  { re: /<(?:p|a):bodyStyle\b[\s\S]*?<\/(?:p|a):bodyStyle>/g, role: "body" },
  { re: /<(?:p|a):otherStyle\b[\s\S]*?<\/(?:p|a):otherStyle>/g, role: "body" },
];
const EMBEDDED_FONT_BLOCK_RE = /<p:embeddedFont>([\s\S]*?)<\/p:embeddedFont>/g;
const PREVIEWABLE_IMAGE_RE = /\.(png|jpe?g|gif|webp|bmp)$/i;
const EMBEDDED_FONT_LST_RE = /<p:embeddedFontLst>[\s\S]*?<\/p:embeddedFontLst>/g;
const FONT_REL_RE = /<Relationship[^>]+Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/font"[^>]*\/>\s*/g;
const REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const REL_SLIDE_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const REL_SLIDE_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const REL_NOTES_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const REL_NOTES_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
const REL_FONT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";
const MEDIA_REL_MARKERS = ["/image", "/video", "/audio", "/media"];
const SETTINGS_STORAGE_KEY = "ppt-finalizer-settings";

const supportsFsAccess = "showOpenFilePicker" in window;

let selectedFile = null;
let sourceFileHandle = null;
let parentDirHandle = null;
let pptxFonts = [];
let mediaAnalysis = null;
let cleanupPlan = null;
let currentFileSize = 0;
let selectedTitleFont = DEFAULT_FONT;
let selectedBodyFont = DEFAULT_FONT;
let pptxZipCache = null;
const mediaThumbUrlCache = new Map();

const titleFontSelect = document.getElementById("titleFontSelect");
const titleCustomFontInput = document.getElementById("titleCustomFontInput");
const bodyFontSelect = document.getElementById("bodyFontSelect");
const bodyCustomFontInput = document.getElementById("bodyCustomFontInput");
const scanningMsg = document.getElementById("scanningMsg");
const dropzone = document.getElementById("dropzone");
const openBtn = document.getElementById("openBtn");
const pickBtn = document.getElementById("pickBtn");
const fileInput = document.getElementById("fileInput");
const fileNameEl = document.getElementById("fileName");
const fileModeEl = document.getElementById("fileMode");
const overwriteBtn = document.getElementById("overwriteBtn");
const downloadBtn = document.getElementById("downloadBtn");
const clearBtn = document.getElementById("clearBtn");
const errorMsg = document.getElementById("errorMsg");
const resultPanel = document.getElementById("resultPanel");
const statsEl = document.getElementById("stats");
const fsNote = document.getElementById("fsNote");
const analysisStack = document.getElementById("analysisStack");
const fileAnalysisSize = document.getElementById("fileAnalysisSize");
const fileAnalysisSlides = document.getElementById("fileAnalysisSlides");
const fileAnalysisEstimate = document.getElementById("fileAnalysisEstimate");
const fontAnalysisBadge = document.getElementById("fontAnalysisBadge");
const fontSummary = document.getElementById("fontSummary");
const fontTableBody = document.getElementById("fontTableBody");
const fontEmpty = document.getElementById("fontEmpty");
const mediaAnalysisBadge = document.getElementById("mediaAnalysisBadge");
const mediaSummary = document.getElementById("mediaSummary");
const mediaTableBody = document.getElementById("mediaTableBody");
const mediaEmpty = document.getElementById("mediaEmpty");
const cleanupPanel = document.getElementById("cleanupPanel");
const optRemoveOrphanMedia = document.getElementById("optRemoveOrphanMedia");
const optRemoveUnusedStructure = document.getElementById("optRemoveUnusedStructure");
const optRemoveNotes = document.getElementById("optRemoveNotes");
const optRemoveProperties = document.getElementById("optRemoveProperties");
const orphanMediaPreview = document.getElementById("orphanMediaPreview");
const structurePreview = document.getElementById("structurePreview");
const structureOptionDesc = document.getElementById("structureOptionDesc");
const notesPreview = document.getElementById("notesPreview");
const propertiesPreview = document.getElementById("propertiesPreview");
const mediaThumbTooltip = document.getElementById("mediaThumbTooltip");
const mediaThumbImg = mediaThumbTooltip.querySelector("img");
const mediaThumbFallback = mediaThumbTooltip.querySelector(".media-thumb-fallback");

if (!supportsFsAccess) {
  openBtn.hidden = true;
  pickBtn.hidden = false;
  overwriteBtn.hidden = true;
  fsNote.textContent = "このブラウザでは別名ダウンロードのみ利用できます。上書き保存は Chrome / Edge をご利用ください。";
} else {
  overwriteBtn.hidden = false;
}

function decodeXmlEntities(value) {
  return value
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function escapeXmlAttr(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function isPreviewableMediaPath(path) {
  return PREVIEWABLE_IMAGE_RE.test(path);
}

function mimeForMediaPath(path) {
  const ext = (path.split(".").pop() || "").toLowerCase();
  const map = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    webp: "image/webp",
    bmp: "image/bmp",
  };
  return map[ext] || "application/octet-stream";
}

function revokeMediaThumbUrls() {
  for (const url of mediaThumbUrlCache.values()) {
    URL.revokeObjectURL(url);
  }
  mediaThumbUrlCache.clear();
  mediaThumbTooltip.hidden = true;
  mediaThumbImg.hidden = false;
  mediaThumbImg.removeAttribute("src");
  mediaThumbFallback.hidden = true;
  mediaThumbFallback.textContent = "";
}

async function getMediaBlobUrl(path) {
  if (mediaThumbUrlCache.has(path)) return mediaThumbUrlCache.get(path);
  if (!pptxZipCache?.files[path]) return null;
  const data = await pptxZipCache.files[path].async("uint8array");
  const url = URL.createObjectURL(new Blob([data], { type: mimeForMediaPath(path) }));
  mediaThumbUrlCache.set(path, url);
  return url;
}

function positionMediaThumbTooltip(clientX, clientY) {
  const rect = mediaThumbTooltip.getBoundingClientRect();
  const pad = 14;
  let x = clientX + pad;
  let y = clientY + pad;
  if (x + rect.width > window.innerWidth - 8) {
    x = Math.max(8, clientX - rect.width - pad);
  }
  if (y + rect.height > window.innerHeight - 8) {
    y = Math.max(8, clientY - rect.height - pad);
  }
  mediaThumbTooltip.style.left = `${x}px`;
  mediaThumbTooltip.style.top = `${y}px`;
}

let mediaThumbLoadToken = 0;

async function showMediaThumb(path, clientX, clientY) {
  if (!isPreviewableMediaPath(path)) return;
  const token = ++mediaThumbLoadToken;
  const url = await getMediaBlobUrl(path);
  if (!url || token !== mediaThumbLoadToken) return;
  mediaThumbImg.src = url;
  mediaThumbImg.hidden = false;
  mediaThumbFallback.hidden = true;
  mediaThumbTooltip.hidden = false;
  mediaThumbTooltip.setAttribute("aria-hidden", "false");
  positionMediaThumbTooltip(clientX, clientY);
  requestAnimationFrame(() => positionMediaThumbTooltip(clientX, clientY));
}

function hideMediaThumb() {
  mediaThumbLoadToken += 1;
  mediaThumbTooltip.hidden = true;
  mediaThumbTooltip.setAttribute("aria-hidden", "true");
  mediaThumbImg.removeAttribute("src");
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function formatSizeChange(before, after) {
  const delta = before - after;
  const pct = before > 0 ? (Math.abs(delta) / before) * 100 : 0;
  if (delta > 0) {
    return `${formatBytes(delta)} 削減（−${pct.toFixed(1)}%）`;
  }
  if (delta < 0) {
    return `${formatBytes(-delta)} 増加（+${pct.toFixed(1)}%）`;
  }
  return "変化なし（0%）";
}

function countReplace(str, re, replacement) {
  const matches = str.match(re);
  const count = matches ? matches.length : 0;
  const data = str.replace(re, replacement);
  return { data, count };
}

function addOption(group, value, label) {
  const option = document.createElement("option");
  option.value = value;
  option.textContent = label;
  group.appendChild(option);
}

function replaceTypefacesInString(text, font) {
  const safeFont = escapeXmlAttr(font.trim() || DEFAULT_FONT);
  return text.replace(TYPEFACE_RE, `typeface="${safeFont}"`);
}

function countTypefaces(text) {
  TYPEFACE_RE.lastIndex = 0;
  const matches = text.match(TYPEFACE_RE);
  return matches ? matches.length : 0;
}

function findNextSpShapeStart(xml, fromIndex) {
  let i = fromIndex;
  while (i < xml.length) {
    const idx = xml.indexOf("<p:sp", i);
    if (idx === -1) return -1;
    const next = xml.charAt(idx + 5);
    if (next === ">" || next === " " || next === "/") return idx;
    i = idx + 5;
  }
  return -1;
}

function findBalancedSpBlockEnd(xml, startIndex) {
  const closeTag = "</p:sp>";
  let depth = 1;
  let i = startIndex + 5;
  while (i < xml.length && depth > 0) {
    const nextOpen = findNextSpShapeStart(xml, i);
    const nextClose = xml.indexOf(closeTag, i);
    if (nextClose === -1) return -1;
    if (nextOpen !== -1 && nextOpen < nextClose) {
      depth += 1;
      i = nextOpen + 5;
    } else {
      depth -= 1;
      if (depth === 0) return nextClose + closeTag.length;
      i = nextClose + closeTag.length;
    }
  }
  return -1;
}

function replaceFontsInStyleBlocks(xml, titleFont, bodyFont) {
  let result = xml;
  let replacements = 0;
  for (const { re, role } of STYLE_BLOCK_RES) {
    re.lastIndex = 0;
    result = result.replace(re, (block) => {
      const font = role === "title" ? titleFont : bodyFont;
      replacements += countTypefaces(block);
      return replaceTypefacesInString(block, font);
    });
  }
  return { data: result, count: replacements };
}

function replaceFontsContextual(data, titleFont, bodyFont) {
  const safeTitle = (titleFont || DEFAULT_FONT).trim() || DEFAULT_FONT;
  const safeBody = (bodyFont || DEFAULT_FONT).trim() || DEFAULT_FONT;
  let replacements = 0;
  const spStore = [];
  let masked = "";
  let cursor = 0;
  while (cursor < data.length) {
    const start = findNextSpShapeStart(data, cursor);
    if (start === -1) {
      masked += data.slice(cursor);
      break;
    }
    masked += data.slice(cursor, start);
    const end = findBalancedSpBlockEnd(data, start);
    if (end === -1) {
      masked += data.slice(start);
      break;
    }
    spStore.push(data.slice(start, end));
    masked += `\x00SP${spStore.length - 1}\x00`;
    cursor = end;
  }

  // Apply body font to unmasked XML first; title/body style blocks are updated afterward
  // so their typeface attributes are not overwritten by the global body pass.
  const remaining = countReplace(masked, TYPEFACE_RE, `typeface="${escapeXmlAttr(safeBody)}"`);
  masked = remaining.data;
  replacements += remaining.count;

  const styles = replaceFontsInStyleBlocks(masked, safeTitle, safeBody);
  masked = styles.data;
  replacements += styles.count;

  const restored = masked.replace(/\x00SP(\d+)\x00/g, (_, indexText) => {
    const block = spStore[Number(indexText)];
    let font = safeBody;
    if (TITLE_PH_RE.test(block)) font = safeTitle;
    else if (BODY_PH_RE.test(block)) font = safeBody;
    replacements += countTypefaces(block);
    return replaceTypefacesInString(block, font);
  });

  return { data: restored, count: replacements };
}

function populateFontSelect(selectEl, customInput, previous) {
  selectEl.innerHTML = "";

  const presetGroup = document.createElement("optgroup");
  presetGroup.label = "おすすめ";
  for (const name of PRESET_FONTS) {
    addOption(presetGroup, name, name);
  }
  selectEl.appendChild(presetGroup);

  const pptxOnly = pptxFonts.filter(({ name }) => !PRESET_SET.has(name));
  if (pptxOnly.length > 0) {
    const pptxGroup = document.createElement("optgroup");
    pptxGroup.label = "この PPTX 内";
    for (const { name, count } of pptxOnly) {
      addOption(pptxGroup, name, `${name}（${count} 箇所）`);
    }
    selectEl.appendChild(pptxGroup);
  }

  addOption(selectEl, CUSTOM_VALUE, "その他（自由入力）");

  const knownValues = new Set([
    ...PRESET_FONTS,
    ...pptxOnly.map(({ name }) => name),
  ]);

  if (knownValues.has(previous)) {
    selectEl.value = previous;
    customInput.hidden = true;
    return previous;
  }
  if (previous && previous !== DEFAULT_FONT) {
    selectEl.value = CUSTOM_VALUE;
    customInput.hidden = false;
    customInput.value = previous;
    return previous;
  }
  selectEl.value = DEFAULT_FONT;
  customInput.hidden = true;
  customInput.value = "";
  return DEFAULT_FONT;
}

function rebuildFontDropdowns() {
  selectedTitleFont = populateFontSelect(titleFontSelect, titleCustomFontInput, selectedTitleFont);
  selectedBodyFont = populateFontSelect(bodyFontSelect, bodyCustomFontInput, selectedBodyFont);
}

function getFontFromPicker(selectEl, customInput) {
  if (selectEl.value === CUSTOM_VALUE) {
    return customInput.value.trim() || DEFAULT_FONT;
  }
  return selectEl.value.trim() || DEFAULT_FONT;
}

function getTargetFonts() {
  return {
    title: getFontFromPicker(titleFontSelect, titleCustomFontInput),
    body: getFontFromPicker(bodyFontSelect, bodyCustomFontInput),
  };
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_STORAGE_KEY);
    if (!raw) return;
    const settings = JSON.parse(raw);
    if (typeof settings.titleFont === "string" && settings.titleFont.trim()) {
      selectedTitleFont = settings.titleFont.trim();
    }
    if (typeof settings.bodyFont === "string" && settings.bodyFont.trim()) {
      selectedBodyFont = settings.bodyFont.trim();
    }
    if (typeof settings.removeOrphanMedia === "boolean") {
      optRemoveOrphanMedia.checked = settings.removeOrphanMedia;
    }
    if (typeof settings.removeUnusedStructure === "boolean") {
      optRemoveUnusedStructure.checked = settings.removeUnusedStructure;
    }
    if (typeof settings.removeNotes === "boolean") {
      optRemoveNotes.checked = settings.removeNotes;
    }
    if (typeof settings.removeProperties === "boolean") {
      optRemoveProperties.checked = settings.removeProperties;
    }
  } catch {
    /* ignore corrupt settings */
  }
}

function saveSettings() {
  try {
    localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify({
      titleFont: getFontFromPicker(titleFontSelect, titleCustomFontInput),
      bodyFont: getFontFromPicker(bodyFontSelect, bodyCustomFontInput),
      removeOrphanMedia: optRemoveOrphanMedia.checked,
      removeUnusedStructure: optRemoveUnusedStructure.checked,
      removeNotes: optRemoveNotes.checked,
      removeProperties: optRemoveProperties.checked,
    }));
  } catch {
    /* ignore quota / private mode */
  }
}

function bindFontPicker(selectEl, customInput, getSelected, setSelected) {
  selectEl.addEventListener("change", () => {
    if (selectEl.value === CUSTOM_VALUE) {
      customInput.hidden = false;
      customInput.focus();
      if (!customInput.value.trim()) {
        customInput.value = getSelected() === CUSTOM_VALUE ? "" : getSelected();
      }
      setSelected(customInput.value.trim() || DEFAULT_FONT);
    } else {
      customInput.hidden = true;
      setSelected(selectEl.value);
    }
    saveSettings();
  });
  customInput.addEventListener("input", () => {
    setSelected(customInput.value.trim() || DEFAULT_FONT);
    saveSettings();
  });
}

function updateFileMode() {
  if (!selectedFile) {
    fileModeEl.hidden = true;
    return;
  }
  fileModeEl.hidden = false;
  fileModeEl.textContent = sourceFileHandle
    ? "上書き保存可"
    : "ダウンロードのみ";
}

function setActionState() {
  const hasFile = Boolean(selectedFile);
  downloadBtn.disabled = !hasFile;
  clearBtn.disabled = !hasFile;
  overwriteBtn.disabled = !hasFile || !sourceFileHandle;
  updateFileMode();
}

async function verifyPermission(handle, readWrite = true) {
  const opts = { mode: readWrite ? "readwrite" : "read" };
  if ((await handle.queryPermission(opts)) === "granted") return true;
  if ((await handle.requestPermission(opts)) === "granted") return true;
  return false;
}

async function writeBlobToHandle(handle, blob) {
  const writable = await handle.createWritable({ keepExistingData: false });
  await writable.write(blob);
  await writable.close();
}

function friendlyFsError(err) {
  const msg = err?.message || String(err);
  if (
    err?.name === "InvalidStateError" &&
    msg.includes("state cached in an interface object")
  ) {
    return "ファイルまたはフォルダの状態がディスク上で変わりました。PowerPoint で開いていないか確認し、OneDrive 等の同期が終わってから「ファイルを開く」で選び直してください。";
  }
  return msg;
}

async function runFsOp(fn) {
  try {
    return await fn();
  } catch (err) {
    if (err?.name === "InvalidStateError") {
      parentDirHandle = null;
    }
    throw new Error(friendlyFsError(err));
  }
}

function backupTimestamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

function backupBaseName(originalName) {
  return originalName.replace(/\.pptx$/i, "");
}

async function resolveBackupName(dirHandle, originalName) {
  const stem = backupBaseName(originalName);
  let backupName = `${stem}_backup.pptx`;
  try {
    await dirHandle.getFileHandle(backupName, { create: false });
    backupName = `${stem}_backup_${backupTimestamp()}.pptx`;
  } catch (err) {
    if (err.name !== "NotFoundError") throw err;
  }
  return backupName;
}

async function ensureParentDirectory(fileHandle) {
  if (parentDirHandle) return parentDirHandle;

  if (typeof fileHandle.getParent === "function") {
    try {
      parentDirHandle = await fileHandle.getParent();
      return parentDirHandle;
    } catch {
      /* fall through */
    }
  }

  parentDirHandle = await window.showDirectoryPicker({
    mode: "readwrite",
    startIn: fileHandle,
  });
  return parentDirHandle;
}

async function createBackupInDirectory(dirHandle, originalFile) {
  const backupName = await resolveBackupName(dirHandle, originalFile.name);
  const backupHandle = await dirHandle.getFileHandle(backupName, { create: true });
  await writeBlobToHandle(backupHandle, originalFile);
  return backupName;
}

function getZipEntrySize(entry) {
  if (!entry) return 0;
  const data = entry._data;
  if (!data) return 0;
  if (typeof data.uncompressedSize === "number" && data.uncompressedSize > 0) {
    return data.uncompressedSize;
  }
  return data.length || data.compressedSize || 0;
}

function getZipEntryCompressedSize(entry) {
  if (!entry) return 0;
  const data = entry._data;
  if (!data) return 0;
  if (typeof data.compressedSize === "number" && data.compressedSize >= 0) {
    return data.compressedSize;
  }
  return data.length || 0;
}

function zipEntryExists(zip, path) {
  return Boolean(path && zip.files[path] && !zip.files[path].dir);
}

function parseRelationships(xml) {
  const rels = [];
  const re = /<Relationship\s+([^>]+?)\/?>/g;
  let match;
  while ((match = re.exec(xml)) !== null) {
    const attrs = match[1];
    const id = /(?:\bId|Id)="([^"]+)"/.exec(attrs)?.[1]
      || /Id="([^"]+)"/.exec(attrs)?.[1];
    const type = /Type="([^"]+)"/.exec(attrs)?.[1];
    const target = /Target="([^"]+)"/.exec(attrs)?.[1];
    if (type && target) rels.push({ id, type, target });
  }
  return rels;
}

function resolveZipPath(basePath, target) {
  const normalized = target.replace(/\\/g, "/");
  if (normalized.startsWith("/")) {
    return normalized.replace(/^\//, "");
  }
  const baseDir = basePath.includes("/")
    ? basePath.slice(0, basePath.lastIndexOf("/") + 1)
    : "";
  const parts = (baseDir + normalized).split("/");
  const stack = [];
  for (const part of parts) {
    if (!part || part === ".") continue;
    if (part === "..") stack.pop();
    else stack.push(part);
  }
  return stack.join("/");
}

function isMediaRelationship(type, target = "") {
  const normalizedTarget = String(target || "").replace(/\\/g, "/");
  if (/(?:^|\/)media\//.test(normalizedTarget)) {
    return true;
  }
  return MEDIA_REL_MARKERS.some((marker) => type.includes(marker));
}

function ownerPathFromRelsPath(relsPath) {
  const match = relsPath.match(/^(ppt\/[^/]+)\/_rels\/(.+)\.rels$/);
  if (match) return `${match[1]}/${match[2]}`;
  return null;
}

async function getPresentationSlideOrder(zip) {
  const presRelsFile = zip.file("ppt/_rels/presentation.xml.rels");
  const presFile = zip.file("ppt/presentation.xml");
  if (!presRelsFile || !presFile) return new Map();

  const presRels = parseRelationships(await presRelsFile.async("string"));
  const rIdToSlide = new Map();
  for (const rel of presRels) {
    if (rel.type === REL_SLIDE) {
      rIdToSlide.set(rel.id, resolveZipPath("ppt/presentation.xml", rel.target));
    }
  }

  const presXml = await presFile.async("string");
  const slidePathToNum = new Map();
  const sldIdRe = /<p:sldId\b[^>]*\br:id="([^"]+)"/g;
  let match;
  let index = 0;
  while ((match = sldIdRe.exec(presXml)) !== null) {
    index += 1;
    const slidePath = rIdToSlide.get(match[1]);
    if (slidePath) slidePathToNum.set(slidePath, index);
  }
  return slidePathToNum;
}

function formatSlideLabel(slideNum, via) {
  if (!slideNum) return via;
  return via ? `スライド ${slideNum}（${via}）` : `スライド ${slideNum}`;
}

function addMediaUsage(mediaUsage, mediaPath, label) {
  if (!mediaPath.startsWith("ppt/media/")) return;
  if (!mediaUsage.has(mediaPath)) {
    mediaUsage.set(mediaPath, new Set());
  }
  mediaUsage.get(mediaPath).add(label);
}

function isOrphanUsageLabel(label) {
  return label.startsWith("未使用");
}

function sortUsageLabels(labels) {
  return labels.sort((a, b) => {
    const aOrphan = isOrphanUsageLabel(a);
    const bOrphan = isOrphanUsageLabel(b);
    if (aOrphan !== bOrphan) return aOrphan ? 1 : -1;
    return a.localeCompare(b, "ja");
  });
}

async function analyzeMediaUsage(zip, slidePathToNum) {
  const mediaSizes = new Map();
  for (const path of Object.keys(zip.files)) {
    if (!path.startsWith("ppt/media/") || zip.files[path].dir) continue;
    const entry = zip.files[path];
    mediaSizes.set(path, {
      size: getZipEntrySize(entry),
      compressedSize: getZipEntryCompressedSize(entry),
    });
  }

  const slideToLayout = new Map();
  const layoutToMaster = new Map();
  const notesToSlide = new Map();
  const chartToSlides = new Map();
  const mediaUsage = new Map();

  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith(".rels")) continue;
    const ownerPath = ownerPathFromRelsPath(path);
    if (!ownerPath) continue;

    const rels = parseRelationships(await zip.files[path].async("string"));

    if (ownerPath.startsWith("ppt/slides/")) {
      for (const rel of rels) {
        if (rel.type === REL_SLIDE_LAYOUT) {
          slideToLayout.set(ownerPath, resolveZipPath(ownerPath, rel.target));
        }
        if (rel.type === REL_NOTES_SLIDE) {
          notesToSlide.set(resolveZipPath(ownerPath, rel.target), ownerPath);
        }
        if (rel.type.includes("/chart")) {
          const chartPath = resolveZipPath(ownerPath, rel.target);
          if (!chartToSlides.has(chartPath)) chartToSlides.set(chartPath, []);
          chartToSlides.get(chartPath).push(ownerPath);
        }
      }
    }

    if (ownerPath.startsWith("ppt/slideLayouts/")) {
      for (const rel of rels) {
        if (rel.type === REL_SLIDE_MASTER) {
          layoutToMaster.set(ownerPath, resolveZipPath(ownerPath, rel.target));
        }
      }
    }
  }

  const layoutToSlides = new Map();
  for (const [slidePath, layoutPath] of slideToLayout) {
    if (!layoutToSlides.has(layoutPath)) layoutToSlides.set(layoutPath, []);
    layoutToSlides.get(layoutPath).push(slidePath);
  }

  const masterToSlides = new Map();
  for (const [slidePath, layoutPath] of slideToLayout) {
    const masterPath = layoutToMaster.get(layoutPath);
    if (!masterPath) continue;
    if (!masterToSlides.has(masterPath)) masterToSlides.set(masterPath, []);
    masterToSlides.get(masterPath).push(slidePath);
  }

  function slideLabelsForPaths(slidePaths, via) {
    const labels = new Set();
    for (const slidePath of slidePaths) {
      const num = slidePathToNum.get(slidePath);
      if (num) labels.add(formatSlideLabel(num, via));
    }
    return labels;
  }

  function ownerShortName(ownerPath) {
    return ownerPath.split("/").pop() || ownerPath;
  }

  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith(".rels")) continue;
    const ownerPath = ownerPathFromRelsPath(path);
    if (!ownerPath) continue;

    const rels = parseRelationships(await zip.files[path].async("string"));

    for (const rel of rels) {
      if (!isMediaRelationship(rel.type, rel.target)) continue;
      const mediaPath = resolveZipPath(ownerPath, rel.target);

      if (ownerPath.startsWith("ppt/slides/")) {
        const num = slidePathToNum.get(ownerPath);
        if (num) {
          addMediaUsage(mediaUsage, mediaPath, formatSlideLabel(num, "直接"));
        } else {
          addMediaUsage(
            mediaUsage,
            mediaPath,
            `${ownerShortName(ownerPath)}（表示順外）`
          );
        }
        continue;
      }

      if (ownerPath.startsWith("ppt/slideLayouts/")) {
        const slidePaths = layoutToSlides.get(ownerPath) || [];
        if (slidePaths.length === 0) {
          addMediaUsage(
            mediaUsage,
            mediaPath,
            `未使用レイアウト: ${ownerShortName(ownerPath)}`
          );
        } else {
          for (const label of slideLabelsForPaths(slidePaths, "レイアウト")) {
            addMediaUsage(mediaUsage, mediaPath, label);
          }
        }
        continue;
      }

      if (ownerPath.startsWith("ppt/slideMasters/")) {
        const slidePaths = masterToSlides.get(ownerPath) || [];
        if (slidePaths.length === 0) {
          addMediaUsage(
            mediaUsage,
            mediaPath,
            `未使用マスター: ${ownerShortName(ownerPath)}`
          );
        } else {
          for (const label of slideLabelsForPaths(slidePaths, "マスター")) {
            addMediaUsage(mediaUsage, mediaPath, label);
          }
        }
        continue;
      }

      if (ownerPath.startsWith("ppt/charts/")) {
        const slidePaths = chartToSlides.get(ownerPath) || [];
        for (const label of slideLabelsForPaths(slidePaths, "チャート")) {
          addMediaUsage(mediaUsage, mediaPath, label);
        }
        if (slidePaths.length === 0) {
          addMediaUsage(
            mediaUsage,
            mediaPath,
            `未参照チャート: ${ownerShortName(ownerPath)}`
          );
        }
        continue;
      }

      if (ownerPath.startsWith("ppt/notesSlides/")) {
        const slidePath = notesToSlide.get(ownerPath);
        const num = slidePath ? slidePathToNum.get(slidePath) : null;
        if (num) {
          addMediaUsage(mediaUsage, mediaPath, formatSlideLabel(num, "ノート"));
        } else {
          addMediaUsage(
            mediaUsage,
            mediaPath,
            `ノート: ${ownerShortName(ownerPath)}（対応スライド不明）`
          );
        }
      }
    }
  }

  const items = [...mediaSizes.entries()]
    .map(([path, sizes]) => {
      const usage = mediaUsage.get(path);
      const slides = usage ? sortUsageLabels([...usage]) : [];
      const isOrphan = slides.length > 0 && slides.every(isOrphanUsageLabel);
      return {
        path,
        name: path.split("/").pop(),
        size: sizes.size,
        compressedSize: sizes.compressedSize,
        slides,
        isOrphan,
        isPreviewable: isPreviewableMediaPath(path),
      };
    })
    .sort((a, b) =>
      b.size - a.size ||
      b.compressedSize - a.compressedSize ||
      a.name.localeCompare(b.name, "ja")
    );

  const totalMediaSize = items.reduce((sum, item) => sum + item.size, 0);
  const totalMediaCompressedSize = items.reduce((sum, item) => sum + item.compressedSize, 0);
  const orphanMediaSize = items
    .filter((item) => item.isOrphan)
    .reduce((sum, item) => sum + item.size, 0);
  const orphanMediaCompressedSize = items
    .filter((item) => item.isOrphan)
    .reduce((sum, item) => sum + item.compressedSize, 0);
  const orphanCount = items.filter((item) => item.isOrphan).length;

  return {
    items,
    totalMediaSize,
    totalMediaCompressedSize,
    orphanMediaSize,
    orphanMediaCompressedSize,
    orphanCount,
    slideCount: slidePathToNum.size,
  };
}

function escapeRegex(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function partFileName(partPath) {
  return partPath.split("/").pop();
}

function relsPathForPart(partPath) {
  const idx = partPath.lastIndexOf("/");
  if (idx === -1) return `_rels/${partFileName(partPath)}.rels`;
  return `${partPath.slice(0, idx)}/_rels/${partFileName(partPath)}.rels`;
}

async function buildDeckStructure(zip) {
  const slidePathToNum = await getPresentationSlideOrder(zip);
  const slideToLayout = new Map();
  const layoutToMaster = new Map();
  const allLayouts = new Set();
  const allMasters = new Set();

  for (const path of Object.keys(zip.files)) {
    if (/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(path)) allLayouts.add(path);
    if (/^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(path)) allMasters.add(path);
  }

  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith(".rels")) continue;
    const owner = ownerPathFromRelsPath(path);
    if (!owner) continue;
    const rels = parseRelationships(await zip.files[path].async("string"));

    if (owner.startsWith("ppt/slides/")) {
      for (const rel of rels) {
        if (rel.type === REL_SLIDE_LAYOUT) {
          slideToLayout.set(owner, resolveZipPath(owner, rel.target));
        }
      }
    }
    if (owner.startsWith("ppt/slideLayouts/")) {
      for (const rel of rels) {
        if (rel.type === REL_SLIDE_MASTER) {
          layoutToMaster.set(owner, resolveZipPath(owner, rel.target));
        }
      }
    }
  }

  const usedLayouts = new Set(slideToLayout.values());
  const unusedLayouts = [...allLayouts].filter((layout) => !usedLayouts.has(layout));
  const usedMasters = new Set();
  for (const layout of usedLayouts) {
    const master = layoutToMaster.get(layout);
    if (master) usedMasters.add(master);
  }
  const unusedMasters = [...allMasters].filter((master) => !usedMasters.has(master));

  return {
    slidePathToNum,
    slideToLayout,
    layoutToMaster,
    allLayouts,
    allMasters,
    usedLayouts,
    unusedLayouts,
    usedMasters,
    unusedMasters,
  };
}

async function collectAllReferencedMedia(zip) {
  const referenced = new Set();
  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith(".rels")) continue;
    const owner = ownerPathFromRelsPath(path);
    if (!owner) continue;
    const rels = parseRelationships(await zip.files[path].async("string"));
    for (const rel of rels) {
      if (!isMediaRelationship(rel.type, rel.target)) continue;
      const mediaPath = resolveZipPath(owner, rel.target);
      if (mediaPath.startsWith("ppt/media/")) referenced.add(mediaPath);
    }
  }
  return referenced;
}

async function computePackageOrphanMedia(zip) {
  const referenced = await collectAllReferencedMedia(zip);
  const orphanPaths = [];

  for (const path of Object.keys(zip.files)) {
    if (!path.startsWith("ppt/media/") || zip.files[path].dir) continue;
    if (!referenced.has(path)) orphanPaths.push(path);
  }

  const items = orphanPaths
    .map((path) => ({
      path,
      name: partFileName(path),
      size: zipEntryExists(zip, path) ? getZipEntrySize(zip.files[path]) : 0,
      compressedSize: zipEntryExists(zip, path)
        ? getZipEntryCompressedSize(zip.files[path])
        : 0,
      missing: !zipEntryExists(zip, path),
    }))
    .sort((a, b) => b.size - a.size || Number(a.missing) - Number(b.missing));

  return {
    items,
    totalSize: items.reduce((sum, item) => sum + item.size, 0),
    totalCompressedSize: items.reduce((sum, item) => sum + item.compressedSize, 0),
    missingCount: items.filter((item) => item.missing).length,
    paths: new Set(items.filter((item) => zipEntryExists(zip, item.path)).map((item) => item.path)),
  };
}

async function collectMediaOwners(zip) {
  const mediaOwners = new Map();

  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith(".rels")) continue;
    const owner = ownerPathFromRelsPath(path);
    if (!owner) continue;
    const rels = parseRelationships(await zip.files[path].async("string"));
    for (const rel of rels) {
      if (!isMediaRelationship(rel.type, rel.target)) continue;
      const mediaPath = resolveZipPath(owner, rel.target);
      if (!mediaPath.startsWith("ppt/media/")) continue;
      if (!mediaOwners.has(mediaPath)) mediaOwners.set(mediaPath, new Set());
      mediaOwners.get(mediaPath).add(owner);
    }
  }

  return mediaOwners;
}

async function computeStructureFreedMedia(zip, layoutsToRemove, mastersToRemove) {
  const doomed = new Set([...layoutsToRemove, ...mastersToRemove]);
  const mediaOwners = await collectMediaOwners(zip);
  const freedPaths = [];

  for (const [mediaPath, owners] of mediaOwners) {
    if (owners.size > 0 && [...owners].every((owner) => doomed.has(owner))) {
      freedPaths.push(mediaPath);
    }
  }

  const items = freedPaths
    .map((path) => ({
      path,
      name: partFileName(path),
      size: zipEntryExists(zip, path) ? getZipEntrySize(zip.files[path]) : 0,
      compressedSize: zipEntryExists(zip, path)
        ? getZipEntryCompressedSize(zip.files[path])
        : 0,
      missing: !zipEntryExists(zip, path),
    }))
    .sort((a, b) => b.size - a.size || Number(a.missing) - Number(b.missing));

  return {
    items,
    totalSize: items.reduce((sum, item) => sum + item.size, 0),
    totalCompressedSize: items.reduce((sum, item) => sum + item.compressedSize, 0),
    missingCount: items.filter((item) => item.missing).length,
    paths: new Set(items.filter((item) => zipEntryExists(zip, item.path)).map((item) => item.path)),
  };
}

async function computeNotesInfo(zip) {
  const paths = Object.keys(zip.files).filter(
    (path) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(path)
  );
  const masterPaths = Object.keys(zip.files).filter(
    (path) => /^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(path)
  );
  let bytes = 0;
  let compressedBytes = 0;
  for (const path of [...paths, ...masterPaths]) {
    bytes += getZipEntrySize(zip.files[path]);
    compressedBytes += getZipEntryCompressedSize(zip.files[path]);
    const relsPath = relsPathForPart(path);
    if (zip.files[relsPath]) {
      bytes += getZipEntrySize(zip.files[relsPath]);
      compressedBytes += getZipEntryCompressedSize(zip.files[relsPath]);
    }
  }
  return { count: paths.length, bytes, compressedBytes, paths };
}

function readXmlElementText(xml, fullTag) {
  const escapedTag = fullTag.replace(":", "\\:");
  const re = new RegExp(`<${escapedTag}[^>]*>([\\s\\S]*?)<\\/${escapedTag}>`, "i");
  const match = xml.match(re);
  if (!match) return "";
  return decodeXmlEntities(match[1].replace(/<[^>]+>/g, "")).trim();
}

function setXmlElementText(xml, fullTag, value) {
  const escapedTag = fullTag.replace(":", "\\:");
  const re = new RegExp(`(<${escapedTag}[^>]*>)([\\s\\S]*?)(<\\/${escapedTag}>)`, "i");
  if (!re.test(xml)) return xml;
  return xml.replace(re, `$1${escapeXmlAttr(value)}$3`);
}

async function scanDocumentProperties(zip) {
  const fields = [];
  const coreFile = zip.file("docProps/core.xml");
  if (coreFile) {
    const xml = await coreFile.async("string");
    const coreTags = [
      ["dc:title", "タイトル"],
      ["dc:subject", "件名"],
      ["dc:creator", "作成者"],
      ["dc:description", "説明"],
      ["cp:lastModifiedBy", "最終更新者"],
      ["cp:keywords", "キーワード"],
      ["cp:category", "カテゴリ"],
      ["cp:contentStatus", "コンテンツ状態"],
    ];
    for (const [tag, label] of coreTags) {
      const value = readXmlElementText(xml, tag);
      if (value) fields.push({ label, value });
    }
    const revision = readXmlElementText(xml, "cp:revision");
    if (revision && revision !== "0" && revision !== "1") {
      fields.push({ label: "リビジョン", value: revision });
    }
  }

  const appFile = zip.file("docProps/app.xml");
  if (appFile) {
    const xml = await appFile.async("string");
    const appTags = [
      ["Company", "会社"],
      ["Manager", "管理者"],
      ["HyperlinkBase", "ハイパーリンク基準"],
    ];
    for (const [tag, label] of appTags) {
      const value = readXmlElementText(xml, tag);
      if (value) fields.push({ label, value });
    }
  }

  const customFile = zip.file("docProps/custom.xml");
  if (customFile) {
    const xml = await customFile.async("string");
    const customCount = (xml.match(/<property\b/gi) || []).length;
    if (customCount > 0) {
      fields.push({ label: "カスタムプロパティ", value: `${customCount} 件` });
    }
  }

  if (zip.file("docProps/thumbnail.jpeg")) {
    fields.push({ label: "サムネイル", value: "あり" });
  }

  let removableBytes = 0;
  const thumbFile = zip.file("docProps/thumbnail.jpeg");
  if (thumbFile) {
    removableBytes += getZipEntryCompressedSize(thumbFile);
  }
  if (customFile) {
    const customXml = await customFile.async("string");
    if ((customXml.match(/<property\b/gi) || []).length > 0) {
      removableBytes += Math.max(0, getZipEntryCompressedSize(customFile) - 180);
    }
  }

  return { fields, fieldCount: fields.length, removableBytes };
}

async function clearDocumentProperties(zip) {
  let cleared = 0;
  const corePath = "docProps/core.xml";
  const coreFile = zip.file(corePath);
  if (coreFile) {
    let xml = await coreFile.async("string");
    const clearTags = [
      "dc:title", "dc:subject", "dc:creator", "dc:description",
      "cp:lastModifiedBy", "cp:keywords", "cp:category", "cp:contentStatus",
    ];
    for (const tag of clearTags) {
      const before = readXmlElementText(xml, tag);
      if (before) cleared += 1;
      xml = setXmlElementText(xml, tag, "");
    }
    const revision = readXmlElementText(xml, "cp:revision");
    if (revision && revision !== "0") {
      cleared += 1;
      xml = setXmlElementText(xml, "cp:revision", "0");
    }
    zip.file(corePath, xml);
  }

  const appPath = "docProps/app.xml";
  const appFile = zip.file(appPath);
  if (appFile) {
    let xml = await appFile.async("string");
    for (const tag of ["Company", "Manager", "HyperlinkBase"]) {
      const before = readXmlElementText(xml, tag);
      if (before) cleared += 1;
      xml = setXmlElementText(xml, tag, "");
    }
    zip.file(appPath, xml);
  }

  const customPath = "docProps/custom.xml";
  if (zip.file(customPath)) {
    const customCount = ((await zip.file(customPath).async("string")).match(/<property\b/gi) || []).length;
    if (customCount > 0) cleared += customCount;
    zip.file(
      customPath,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ' +
      'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>'
    );
  }

  if (zip.file("docProps/thumbnail.jpeg")) {
    cleared += 1;
    deleteZipEntry(zip, "docProps/thumbnail.jpeg");
    const ctPath = "[Content_Types].xml";
    const ctFile = zip.file(ctPath);
    if (ctFile) {
      let ctXml = await ctFile.async("string");
      ctXml = ctXml.replace(
        /<Override[^>]+PartName="\/docProps\/thumbnail\.jpeg"[^>]*\/>\s*/g,
        ""
      );
      zip.file(ctPath, ctXml);
    }
  }

  return { cleared };
}

async function removeNotesFromZip(zip) {
  const deleted = new Set();
  const notesPaths = Object.keys(zip.files).filter(
    (path) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(path)
  );
  let bytes = 0;

  for (const path of Object.keys(zip.files)) {
    if (!/^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(path)) continue;
    let relsXml = await zip.files[path].async("string");
    const rels = parseRelationships(relsXml);
    let changed = false;
    for (const rel of rels) {
      if (rel.type === REL_NOTES_SLIDE) {
        relsXml = removeRelationshipById(relsXml, rel.id);
        changed = true;
      }
    }
    if (changed) zip.file(path, relsXml);
  }

  for (const notesPath of notesPaths) {
    bytes += getZipEntrySize(zip.files[notesPath]);
    const relsPath = relsPathForPart(notesPath);
    if (zip.files[relsPath]) bytes += getZipEntrySize(zip.files[relsPath]);
    deleteZipEntry(zip, notesPath);
    deleteZipEntry(zip, relsPath);
    deleted.add(notesPath);
    deleted.add(relsPath);
  }

  const presRelsPath = "ppt/_rels/presentation.xml.rels";
  const presPath = "ppt/presentation.xml";
  const presRelsFile = zip.file(presRelsPath);
  const presFile = zip.file(presPath);
  const notesMasterPaths = Object.keys(zip.files).filter(
    (path) => /^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(path)
  );

  if (presRelsFile && presFile) {
    let relsXml = await presRelsFile.async("string");
    let presXml = await presFile.async("string");
    const rels = parseRelationships(relsXml);
    const masterRIds = new Set();

    for (const rel of rels) {
      if (rel.type === REL_NOTES_MASTER) {
        masterRIds.add(rel.id);
        relsXml = removeRelationshipById(relsXml, rel.id);
      }
    }
    if (masterRIds.size > 0) {
      zip.file(presRelsPath, relsXml);
      for (const rId of masterRIds) {
        presXml = removeXmlElementsByRId(presXml, rId);
      }
      presXml = presXml.replace(/<p:notesMasterIdLst>[\s\S]*?<\/p:notesMasterIdLst>\s*/g, "");
      zip.file(presPath, presXml);
    }
  }

  for (const masterPath of notesMasterPaths) {
    bytes += getZipEntrySize(zip.files[masterPath]);
    const relsPath = relsPathForPart(masterPath);
    if (zip.files[relsPath]) bytes += getZipEntrySize(zip.files[relsPath]);
    deleteZipEntry(zip, masterPath);
    deleteZipEntry(zip, relsPath);
    deleted.add(masterPath);
    deleted.add(relsPath);
  }

  await removeContentTypeOverrides(zip, deleted);
  return { count: notesPaths.length, bytes };
}

async function computeCleanupPlan(zip) {
  const structure = await buildDeckStructure(zip);
  const slideOrphans = await computePackageOrphanMedia(zip);
  const notes = await computeNotesInfo(zip);
  const properties = await scanDocumentProperties(zip);

  const layoutsToRemove = new Set(structure.unusedLayouts);
  for (const master of structure.unusedMasters) {
    for (const layout of structure.allLayouts) {
      if (structure.layoutToMaster.get(layout) === master) layoutsToRemove.add(layout);
    }
  }

  const layoutsOnUsedMasters = structure.unusedLayouts.filter((layout) => {
    const master = structure.layoutToMaster.get(layout);
    return !master || structure.usedMasters.has(master);
  });

  const structureFreedMedia = await computeStructureFreedMedia(
    zip,
    [...layoutsToRemove],
    [...structure.unusedMasters]
  );

  return {
    structure,
    slideOrphanMedia: slideOrphans,
    structureFreedMedia,
    layoutsToRemove: [...layoutsToRemove],
    mastersToRemove: [...structure.unusedMasters],
    unusedLayoutCount: layoutsOnUsedMasters.length,
    unusedMasterCount: structure.unusedMasters.length,
    notes,
    properties,
  };
}

function removeRelationshipById(relsXml, rId) {
  if (!rId) return relsXml;
  return relsXml.replace(
    new RegExp(`<Relationship\\s+[^>]*\\bId="${escapeRegex(rId)}"[^>]*/>\\s*`, "g"),
    ""
  );
}

function removeXmlElementsByRId(xml, rId) {
  if (!rId) return xml;
  return xml.replace(
    new RegExp(`<[^>]+\\br:id="${escapeRegex(rId)}"[^>]*/>\\s*`, "g"),
    ""
  );
}

async function removeContentTypeOverrides(zip, deletedPaths) {
  const ctPath = "[Content_Types].xml";
  const ctFile = zip.file(ctPath);
  if (!ctFile) return;
  let xml = await ctFile.async("string");
  for (const partPath of deletedPaths) {
    if (!partPath.endsWith(".xml")) continue;
    const partName = `/${partPath}`;
    xml = xml.replace(
      new RegExp(`<Override[^>]+PartName="${escapeRegex(partName)}"[^>]*/>\\s*`, "g"),
      ""
    );
  }
  zip.file(ctPath, xml);
}

function deleteZipEntry(zip, path) {
  if (zip.files[path]) delete zip.files[path];
}

async function removeLayoutFromMaster(zip, layoutPath, layoutToMaster) {
  const masterPath = layoutToMaster.get(layoutPath);
  if (!masterPath) return;
  const masterRelsPath = relsPathForPart(masterPath);
  const masterRelsFile = zip.file(masterRelsPath);
  const masterFile = zip.file(masterPath);
  if (!masterRelsFile || !masterFile) return;

  let relsXml = await masterRelsFile.async("string");
  const rels = parseRelationships(relsXml);
  let layoutRId = null;
  for (const rel of rels) {
    if (rel.type === REL_SLIDE_LAYOUT && resolveZipPath(masterPath, rel.target) === layoutPath) {
      layoutRId = rel.id;
      break;
    }
  }
  if (!layoutRId) return;

  relsXml = removeRelationshipById(relsXml, layoutRId);
  zip.file(masterRelsPath, relsXml);

  let masterXml = await masterFile.async("string");
  masterXml = removeXmlElementsByRId(masterXml, layoutRId);
  zip.file(masterPath, masterXml);
}

async function removeMasterFromPresentation(zip, masterPath) {
  const presRelsPath = "ppt/_rels/presentation.xml.rels";
  const presPath = "ppt/presentation.xml";
  const presRelsFile = zip.file(presRelsPath);
  const presFile = zip.file(presPath);
  if (!presRelsFile || !presFile) return;

  let relsXml = await presRelsFile.async("string");
  const rels = parseRelationships(relsXml);
  let masterRId = null;
  for (const rel of rels) {
    if (rel.type === REL_SLIDE_MASTER && resolveZipPath(presPath, rel.target) === masterPath) {
      masterRId = rel.id;
      break;
    }
  }
  if (!masterRId) return;

  relsXml = removeRelationshipById(relsXml, masterRId);
  zip.file(presRelsPath, relsXml);

  let presXml = await presFile.async("string");
  presXml = removeXmlElementsByRId(presXml, masterRId);
  zip.file(presPath, presXml);
}

async function removeUnusedStructure(zip, plan) {
  const deleted = new Set();
  const layoutToMaster = plan.structure.layoutToMaster;

  for (const layoutPath of plan.layoutsToRemove) {
    if (plan.structure.usedMasters.has(layoutToMaster.get(layoutPath))) {
      await removeLayoutFromMaster(zip, layoutPath, layoutToMaster);
    }
    deleteZipEntry(zip, layoutPath);
    deleteZipEntry(zip, relsPathForPart(layoutPath));
    deleted.add(layoutPath);
    deleted.add(relsPathForPart(layoutPath));
  }

  for (const masterPath of plan.mastersToRemove) {
    await removeMasterFromPresentation(zip, masterPath);
    deleteZipEntry(zip, masterPath);
    deleteZipEntry(zip, relsPathForPart(masterPath));
    deleted.add(masterPath);
    deleted.add(relsPathForPart(masterPath));
  }

  await removeContentTypeOverrides(zip, deleted);
  return {
    layoutsRemoved: plan.layoutsToRemove.length,
    mastersRemoved: plan.mastersToRemove.length,
  };
}

async function removePackageOrphanMedia(zip) {
  const referenced = await collectAllReferencedMedia(zip);
  let bytes = 0;
  let count = 0;
  for (const path of Object.keys(zip.files)) {
    if (!path.startsWith("ppt/media/") || zip.files[path].dir) continue;
    if (referenced.has(path)) continue;
    bytes += getZipEntrySize(zip.files[path]);
    deleteZipEntry(zip, path);
    count += 1;
  }
  return { count, bytes };
}

async function removeSlideOrphanMedia(zip, orphanPaths) {
  let bytes = 0;
  let count = 0;
  for (const path of orphanPaths) {
    if (!zip.files[path]) continue;
    bytes += getZipEntrySize(zip.files[path]);
    deleteZipEntry(zip, path);
    count += 1;
  }
  return { count, bytes };
}

async function applyFontReplaceToZip(zip, fonts) {
  let replacements = 0;
  let filesChanged = 0;
  let embeddedFontsRemoved = 0;

  for (const path of Object.keys(zip.files)) {
    const entry = zip.files[path];
    if (entry.dir) continue;

    if (path.startsWith("ppt/fonts/")) {
      embeddedFontsRemoved += 1;
      deleteZipEntry(zip, path);
      continue;
    }

    if (path.endsWith(".xml") || path.endsWith(".rels")) {
      const text = await entry.async("string");
      const { data, count } = replaceFontsInXml(path, text, fonts);
      if (data !== text) filesChanged += 1;
      replacements += count;
      zip.file(path, data);
    }
  }

  return { replacements, filesChanged, embeddedFontsRemoved };
}

function getFinalizeOptions() {
  return {
    removeOrphanMedia: optRemoveOrphanMedia.checked,
    removeUnusedStructure: optRemoveUnusedStructure.checked,
    removeNotes: optRemoveNotes.checked,
    removeProperties: optRemoveProperties.checked,
  };
}

function zipEntryCompressedBytes(zip, path) {
  if (!zipEntryExists(zip, path)) return 0;
  return getZipEntryCompressedSize(zip.files[path]);
}

function computeReductionEstimate(plan, options, zip) {
  if (!plan || !zip) return 0;

  let bytes = 0;

  for (const path of Object.keys(zip.files)) {
    if (path.startsWith("ppt/fonts/") && !zip.files[path].dir) {
      bytes += getZipEntryCompressedSize(zip.files[path]);
    }
  }

  if (options.removeOrphanMedia) {
    for (const item of plan.slideOrphanMedia.items) {
      if (!item.missing) {
        bytes += item.compressedSize || zipEntryCompressedBytes(zip, item.path);
      }
    }
  }

  if (options.removeUnusedStructure) {
    for (const path of plan.layoutsToRemove) {
      bytes += zipEntryCompressedBytes(zip, path);
      bytes += zipEntryCompressedBytes(zip, relsPathForPart(path));
    }
    for (const path of plan.mastersToRemove) {
      bytes += zipEntryCompressedBytes(zip, path);
      bytes += zipEntryCompressedBytes(zip, relsPathForPart(path));
    }
    bytes += plan.structureFreedMedia.totalCompressedSize;
  }

  if (options.removeNotes && plan.notes.count > 0) {
    bytes += plan.notes.compressedBytes || plan.notes.bytes;
  }

  if (options.removeProperties && plan.properties.removableBytes > 0) {
    bytes += plan.properties.removableBytes;
  }

  return bytes;
}

function updateReductionEstimate() {
  if (!fileAnalysisEstimate) return;
  if (!cleanupPlan || !pptxZipCache || !currentFileSize) {
    fileAnalysisEstimate.textContent = "—";
    return;
  }

  const bytes = computeReductionEstimate(cleanupPlan, getFinalizeOptions(), pptxZipCache);
  if (bytes <= 0) {
    fileAnalysisEstimate.textContent = "なし";
    return;
  }

  const after = Math.max(0, currentFileSize - bytes);
  const pct = currentFileSize > 0 ? ((bytes / currentFileSize) * 100).toFixed(1) : "0";
  fileAnalysisEstimate.textContent =
    `約 ${formatBytes(bytes)}（${pct}%・仕上げ後 ${formatBytes(after)} 前後）`;
}

function renderCleanupPreview() {
  if (!cleanupPlan) {
    cleanupPanel.hidden = true;
    return;
  }
  cleanupPanel.hidden = false;
  const slideCount = cleanupPlan.structure.slidePathToNum.size;
  if (structureOptionDesc) {
    structureOptionDesc.textContent =
      `${slideCount} 枚のスライドから参照されていない型・マスター XML および関連するメディアを削除します`;
  }
  orphanMediaPreview.textContent = cleanupPlan.slideOrphanMedia.items.length
    ? `${formatBytes(cleanupPlan.slideOrphanMedia.totalSize)}（${cleanupPlan.slideOrphanMedia.items.length} 件` +
      `${cleanupPlan.slideOrphanMedia.missingCount ? `、参照のみ ${cleanupPlan.slideOrphanMedia.missingCount} 件` : ""}）`
    : "0 B（0 件）";
  const structureMedia = cleanupPlan.structureFreedMedia;
  const structureMediaPart = structureMedia.items.length
    ? ` + 関連メディア ${structureMedia.items.length} 件（${formatBytes(structureMedia.totalCompressedSize)}）`
    : "";
  structurePreview.textContent =
    `${cleanupPlan.unusedLayoutCount} レイアウト + ${cleanupPlan.unusedMasterCount} マスター${structureMediaPart}`;
  notesPreview.textContent = cleanupPlan.notes.count
    ? `${cleanupPlan.notes.count} 件（${formatBytes(cleanupPlan.notes.bytes)}）`
    : "0 件";
  if (cleanupPlan.properties.fieldCount > 0) {
    const sample = cleanupPlan.properties.fields
      .slice(0, 3)
      .map((field) => field.label)
      .join("、");
    const suffix = cleanupPlan.properties.fieldCount > 3 ? " 等" : "";
    propertiesPreview.textContent = `${cleanupPlan.properties.fieldCount} 項目（${sample}${suffix}）`;
  } else {
    propertiesPreview.textContent = "0 項目";
  }
  updateReductionEstimate();
}

function renderFileAnalysis(totalFileSize, slideCount) {
  currentFileSize = totalFileSize;
  analysisStack.hidden = false;
  fileAnalysisSize.textContent = formatBytes(totalFileSize);
  fileAnalysisSlides.textContent = `${slideCount} 枚`;
  updateReductionEstimate();
}

function renderFontAnalysis(fonts) {
  if (!fonts || fonts.length === 0) {
    analysisStack.hidden = false;
    fontAnalysisBadge.textContent = "フォントなし";
    fontSummary.textContent = "";
    fontTableBody.innerHTML = "";
    fontEmpty.hidden = false;
    return;
  }

  analysisStack.hidden = false;
  fontEmpty.hidden = true;
  const totalUses = fonts.reduce((sum, font) => sum + font.count, 0);
  const embeddedKinds = fonts.filter((font) => font.fileBytes > 0).length;
  let totalEmbedded = 0;
  if (pptxZipCache) {
    for (const path of Object.keys(pptxZipCache.files)) {
      if (!path.startsWith("ppt/fonts/") || pptxZipCache.files[path].dir) continue;
      totalEmbedded += getZipEntrySize(pptxZipCache.files[path]);
    }
  }
  fontAnalysisBadge.textContent = `${fonts.length} 種類 / 埋込 ${formatBytes(totalEmbedded)}`;
  fontSummary.textContent =
    `検出フォント: ${fonts.length} 種類 / typeface 合計: ${totalUses} 箇所 / ` +
    `埋込フォント: ${embeddedKinds} 種類（${formatBytes(totalEmbedded)}）`;

  fontTableBody.innerHTML = fonts
    .map((font) => {
      const fileLabel = font.fileBytes > 0 ? formatBytes(font.fileBytes) : "—";
      const fileTitle = font.fileDetail
        ? ` title="${escapeHtml(font.fileDetail)}"`
        : "";
      return `<tr>
      <td>${escapeHtml(font.name)}</td>
      <td class="file-size"${fileTitle}>${escapeHtml(fileLabel)}</td>
      <td class="count">${font.count}</td>
    </tr>`;
    })
    .join("");
}

async function mapEmbeddedFontFileSizes(zip) {
  const typefaceFiles = new Map();

  const presFile = zip.file("ppt/presentation.xml");
  const presRelsFile = zip.file("ppt/_rels/presentation.xml.rels");
  if (!presFile || !presRelsFile) {
    return typefaceFiles;
  }

  const presXml = await presFile.async("string");
  const rels = parseRelationships(await presRelsFile.async("string"));
  const rIdToFontPath = new Map();
  for (const rel of rels) {
    if (rel.type === REL_FONT) {
      rIdToFontPath.set(rel.id, resolveZipPath("ppt/presentation.xml", rel.target));
    }
  }

  EMBEDDED_FONT_BLOCK_RE.lastIndex = 0;
  let block;
  while ((block = EMBEDDED_FONT_BLOCK_RE.exec(presXml)) !== null) {
    const inner = block[1];
    const tfMatch = inner.match(/typeface="([^"]*)"/);
    if (!tfMatch) continue;
    const typeface = decodeXmlEntities(tfMatch[1]).trim();
    if (!typeface) continue;

    const rIds = new Set();
    const ridRe = /\br:id="([^"]+)"/g;
    let ridMatch;
    while ((ridMatch = ridRe.exec(inner)) !== null) {
      rIds.add(ridMatch[1]);
    }

    const files = [];
    let bytes = 0;
    for (const rId of rIds) {
      const fontPath = rIdToFontPath.get(rId);
      if (!fontPath || !zip.files[fontPath] || files.includes(fontPath)) continue;
      files.push(fontPath);
      bytes += getZipEntrySize(zip.files[fontPath]);
    }

    if (files.length === 0) continue;
    typefaceFiles.set(typeface, {
      bytes,
      files: files.map((path) => ({
        path,
        name: partFileName(path),
        size: getZipEntrySize(zip.files[path]),
      })),
    });
  }

  return typefaceFiles;
}

function buildEmbeddedFontDetail(files) {
  return files
    .sort((a, b) => b.size - a.size)
    .map((file) => `${file.name} (${formatBytes(file.size)})`)
    .join(" / ");
}

async function extractFontsFromZipAsync(zip) {
  const fonts = new Map();
  const embeddedMap = await mapEmbeddedFontFileSizes(zip);

  function ensureFont(name) {
    if (!fonts.has(name)) {
      fonts.set(name, { name, count: 0 });
    }
    return fonts.get(name);
  }

  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith(".xml")) continue;
    const text = await zip.files[path].async("string");

    TYPEFACE_CAPTURE_RE.lastIndex = 0;
    let match;
    while ((match = TYPEFACE_CAPTURE_RE.exec(text)) !== null) {
      const name = decodeXmlEntities(match[1]).trim();
      if (name) ensureFont(name).count += 1;
    }
  }

  return [...fonts.values()]
    .map((font) => {
      const embedded = embeddedMap.get(font.name);
      return {
        name: font.name,
        count: font.count,
        fileBytes: embedded?.bytes || 0,
        fileDetail: embedded?.files?.length
          ? buildEmbeddedFontDetail(embedded.files)
          : "",
      };
    })
    .sort((a, b) =>
      b.fileBytes - a.fileBytes ||
      b.count - a.count ||
      a.name.localeCompare(b.name, "ja")
    );
}

async function analyzePptxFile(file) {
  revokeMediaThumbUrls();
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  pptxZipCache = zip;
  const slidePathToNum = await getPresentationSlideOrder(zip);
  const [fonts, media, plan] = await Promise.all([
    extractFontsFromZipAsync(zip),
    analyzeMediaUsage(zip, slidePathToNum),
    computeCleanupPlan(zip),
  ]);
  return { fonts, media, cleanupPlan: plan, totalSize: file.size };
}

function renderMediaAnalysis(analysis) {
  if (!analysis || analysis.items.length === 0) {
    mediaAnalysisBadge.textContent = "メディアなし";
    mediaSummary.textContent = "ppt/media/ 内のメディアはありません。";
    mediaTableBody.innerHTML = "";
    mediaEmpty.hidden = false;
    return;
  }

  mediaEmpty.hidden = true;
  mediaAnalysisBadge.textContent =
    `${formatBytes(analysis.totalMediaCompressedSize)} / 未使用メディア ${formatBytes(analysis.orphanMediaCompressedSize)}`;

  const previewableCount = analysis.items.filter((item) => item.isPreviewable).length;
  if (analysis.items.length > 30) {
    mediaSummary.textContent = previewableCount > 0
      ? "サイズ上位 30 件。画像ファイル名ホバーでプレビュー。"
      : "サイズ上位 30 件。";
  } else {
    mediaSummary.textContent = previewableCount > 0
      ? `${analysis.items.length} 件。画像ファイル名ホバーでプレビュー。`
      : `${analysis.items.length} 件。`;
  }

  mediaTableBody.innerHTML = analysis.items
    .slice(0, 30)
    .map((item) => {
      const slideText = item.slides.length
        ? escapeHtml(item.slides.join("、"))
        : '<span style="color:var(--muted)">（参照元不明）</span>';
      const rowStyle = item.isOrphan ? ' style="opacity:0.92"' : "";
      const nameCell = item.isPreviewable
        ? `<td class="media-name media-name--preview" data-path="${escapeHtml(item.path)}" title="ホバーでプレビュー">${escapeHtml(item.name)}</td>`
        : `<td class="media-name">${escapeHtml(item.name)}</td>`;
      return `<tr${rowStyle}>
        ${nameCell}
        <td class="size-compressed">${formatBytes(item.compressedSize)}</td>
        <td class="slides">${slideText}</td>
      </tr>`;
    })
    .join("");
}

async function extractFontsFromPptx(file) {
  const { fonts } = await analyzePptxFile(file);
  return fonts;
}

function replaceFontsInXml(name, data, fonts) {
  const titleFont = fonts.title.trim() || DEFAULT_FONT;
  const bodyFont = fonts.body.trim() || DEFAULT_FONT;
  let total = 0;

  if (name === "ppt/presentation.xml") {
    let result = countReplace(data, EMBEDDED_FONT_LST_RE, "");
    data = result.data;
    total += result.count;
    data = data.replace(/embedTrueTypeFonts="1"/g, 'embedTrueTypeFonts="0"');
    const contextual = replaceFontsContextual(data, titleFont, bodyFont);
    return { data: contextual.data, count: total + contextual.count };
  }

  if (name === "ppt/_rels/presentation.xml.rels") {
    const result = countReplace(data, FONT_REL_RE, "");
    return { data: result.data, count: result.count };
  }

  if (!name.endsWith(".xml")) {
    return { data, count: 0 };
  }

  const contextual = replaceFontsContextual(data, titleFont, bodyFont);
  return { data: contextual.data, count: contextual.count };
}

async function finalizePptx(file, fonts, options) {
  const titleFont = fonts.title.trim() || DEFAULT_FONT;
  const bodyFont = fonts.body.trim() || DEFAULT_FONT;
  if (!titleFont || !bodyFont) {
    throw new Error("フォント名を入力してください。");
  }
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    throw new Error(".pptx ファイルを選択してください。");
  }

  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const plan = await computeCleanupPlan(zip);
  const embeddedMap = await mapEmbeddedFontFileSizes(zip);
  const embeddedFontKinds = embeddedMap.size;
  const embeddedFontBytes = [...embeddedMap.values()].reduce((sum, item) => sum + item.bytes, 0);

  const fontStats = await applyFontReplaceToZip(zip, fonts);

  let structureStats = { layoutsRemoved: 0, mastersRemoved: 0, mediaRemoved: 0, mediaBytes: 0 };
  if (options.removeUnusedStructure) {
    structureStats = await removeUnusedStructure(zip, plan);
    const structureMediaStats = await removePackageOrphanMedia(zip);
    structureStats.mediaRemoved = structureMediaStats.count;
    structureStats.mediaBytes = structureMediaStats.bytes;
  }

  let mediaStats = { count: 0, bytes: 0 };
  let notesStats = { count: 0, bytes: 0 };
  let propertiesStats = { cleared: 0 };

  if (options.removeNotes) {
    notesStats = await removeNotesFromZip(zip);
  }

  if (options.removeProperties) {
    propertiesStats = await clearDocumentProperties(zip);
  }

  if (options.removeOrphanMedia) {
    mediaStats = await removePackageOrphanMedia(zip);
  }

  const blob = await zip.generateAsync({
    type: "blob",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return {
    blob,
    replacements: fontStats.replacements,
    filesChanged: fontStats.filesChanged,
    embeddedFontsRemoved: fontStats.embeddedFontsRemoved,
    embeddedFontKinds,
    embeddedFontBytes,
    layoutsRemoved: structureStats.layoutsRemoved,
    mastersRemoved: structureStats.mastersRemoved,
    structureMediaRemoved: structureStats.mediaRemoved,
    structureMediaBytes: structureStats.mediaBytes,
    orphanMediaRemoved: mediaStats.count,
    orphanMediaBytes: mediaStats.bytes,
    notesRemoved: notesStats.count,
    notesBytes: notesStats.bytes,
    propertiesCleared: propertiesStats.cleared,
    titleFont,
    bodyFont,
    targetFont: bodyFont,
    sourceFonts: pptxFonts.map((f) => f.name),
    originalSize: file.size,
    outputSize: blob.size,
  };
}

async function loadFromFile(file, handle = null) {
  selectedFile = file;
  sourceFileHandle = handle;
  if (!handle) parentDirHandle = null;

  fileNameEl.textContent = file ? file.name : "";
  hideError();
  resultPanel.hidden = true;
  setActionState();

  if (!file) {
    pptxFonts = [];
    mediaAnalysis = null;
    cleanupPlan = null;
    pptxZipCache = null;
    revokeMediaThumbUrls();
    analysisStack.hidden = true;
    cleanupPanel.hidden = true;
    rebuildFontDropdowns();
    return;
  }

  scanningMsg.hidden = false;
  scanningMsg.textContent = "分析中…";

  try {
    const analysis = await analyzePptxFile(file);
    pptxFonts = analysis.fonts;
    mediaAnalysis = analysis.media;
    cleanupPlan = analysis.cleanupPlan;
    rebuildFontDropdowns();
    renderFileAnalysis(analysis.totalSize, mediaAnalysis?.slideCount ?? 0);
    renderFontAnalysis(analysis.fonts);
    renderMediaAnalysis(mediaAnalysis);
    renderCleanupPreview();
  } catch (err) {
    showError(`分析に失敗しました: ${err.message || err}`);
    pptxFonts = [];
    mediaAnalysis = null;
    cleanupPlan = null;
    pptxZipCache = null;
    revokeMediaThumbUrls();
    analysisStack.hidden = true;
    cleanupPanel.hidden = true;
    rebuildFontDropdowns();
  } finally {
    scanningMsg.hidden = true;
    scanningMsg.textContent = "分析中…";
  }
}

async function openFileWithHandle() {
  if (!supportsFsAccess) {
    fileInput.click();
    return;
  }

  try {
    const [handle] = await window.showOpenFilePicker({
      mode: "readwrite",
      types: PPTX_TYPES,
    });
    if (!(await verifyPermission(handle, true))) {
      throw new Error("ファイルへの書き込み権限が必要です。");
    }
    const file = await handle.getFile();
    parentDirHandle = null;
    await loadFromFile(file, handle);
  } catch (err) {
    if (err.name === "AbortError") return;
    showError(err.message || String(err));
  }
}

async function loadFromDrop(dataTransfer) {
  const item = dataTransfer.items?.[0];
  if (item && typeof item.getAsFileSystemHandle === "function") {
    try {
      const handle = await item.getAsFileSystemHandle();
      if (handle?.kind === "file") {
        if (await verifyPermission(handle, true)) {
          const file = await handle.getFile();
          parentDirHandle = null;
          await loadFromFile(file, handle);
          return;
        }
      }
    } catch {
      /* fall through to File */
    }
  }

  const file = dataTransfer.files?.[0];
  if (file) {
    await loadFromFile(file, null);
  }
}

async function overwriteOriginal() {
  if (!selectedFile || !sourceFileHandle) {
    throw new Error("上書き保存するには「ファイルを開く（上書き可）」で PPTX を選んでください。");
  }

  if (!(await verifyPermission(sourceFileHandle, true))) {
    throw new Error("ファイルへの書き込み権限が必要です。");
  }

  // ディレクトリハンドルのキャッシュは InvalidStateError の原因になりやすい
  parentDirHandle = null;

  const originalFile = await runFsOp(() => sourceFileHandle.getFile());
  const dirHandle = await runFsOp(() => ensureParentDirectory(sourceFileHandle));
  if (!(await verifyPermission(dirHandle, true))) {
    throw new Error("バックアップ先フォルダへの書き込み権限が必要です。");
  }

  const backupName = await runFsOp(() =>
    createBackupInDirectory(dirHandle, originalFile)
  );

  const fonts = getTargetFonts();
  const options = getFinalizeOptions();
  const result = await finalizePptx(originalFile, fonts, options);
  await runFsOp(() => writeBlobToHandle(sourceFileHandle, result.blob));

  selectedFile = new File([result.blob], originalFile.name, {
    type: originalFile.type || "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
  scanningMsg.hidden = false;
  scanningMsg.textContent = "分析中…";
  try {
    const analysis = await analyzePptxFile(selectedFile);
    pptxFonts = analysis.fonts;
    mediaAnalysis = analysis.media;
    cleanupPlan = analysis.cleanupPlan;
    rebuildFontDropdowns();
    renderFileAnalysis(analysis.totalSize, mediaAnalysis?.slideCount ?? 0);
    renderFontAnalysis(analysis.fonts);
    renderMediaAnalysis(mediaAnalysis);
    renderCleanupPreview();
  } finally {
    scanningMsg.hidden = true;
  }
  setActionState();

  return {
    ...result,
    outputName: originalFile.name,
    backupName,
    saveMode: "overwrite",
  };
}

function showError(message) {
  errorMsg.textContent = message;
  errorMsg.hidden = false;
}

function hideError() {
  errorMsg.hidden = true;
  errorMsg.textContent = "";
}

function renderStats(stats) {
  const convertedFontCount = stats.sourceFonts.length;
  const embeddedLine = stats.embeddedFontKinds > 0
    ? `${stats.embeddedFontKinds} 種類（${formatBytes(stats.embeddedFontBytes)}）`
    : "0 種類";
  const backupLine = stats.backupName
    ? `<dt>バックアップ</dt><dd>${escapeHtml(stats.backupName)}（元ファイルと同じフォルダ）</dd>`
    : "";
  const modeLine = stats.saveMode === "overwrite"
    ? `<dt>保存方法</dt><dd>元ファイルに上書き</dd>`
    : `<dt>保存方法</dt><dd>別名ダウンロード</dd>`;
  const sizeLine = stats.originalSize != null && stats.outputSize != null
    ? `<dt>ファイルサイズ</dt><dd>${formatBytes(stats.originalSize)} → ${formatBytes(stats.outputSize)}</dd>
       <dt>サイズ削減</dt><dd>${formatSizeChange(stats.originalSize, stats.outputSize)}</dd>`
    : "";

  const cleanupLines = [];
  if (stats.orphanMediaRemoved > 0) {
    cleanupLines.push(
      `<dt>孤立メディア削除</dt><dd>${stats.orphanMediaRemoved} 件（${formatBytes(stats.orphanMediaBytes)}）</dd>`
    );
  }
  if (stats.layoutsRemoved > 0 || stats.mastersRemoved > 0) {
    let structureDetail =
      `レイアウト ${stats.layoutsRemoved}、マスター ${stats.mastersRemoved}`;
    if (stats.structureMediaRemoved > 0) {
      structureDetail +=
        `、関連メディア ${stats.structureMediaRemoved} 件（${formatBytes(stats.structureMediaBytes)}）`;
    }
    cleanupLines.push(
      `<dt>未使用レイアウト／マスター削除</dt><dd>${structureDetail}</dd>`
    );
  }
  if (stats.notesRemoved > 0) {
    cleanupLines.push(
      `<dt>ノート削除</dt><dd>${stats.notesRemoved} 件（${formatBytes(stats.notesBytes)}）</dd>`
    );
  }
  if (stats.propertiesCleared > 0) {
    cleanupLines.push(
      `<dt>プロパティ削除</dt><dd>${stats.propertiesCleared} 項目</dd>`
    );
  }

  statsEl.innerHTML = `
    <dt>タイトルフォント</dt><dd>${escapeHtml(stats.titleFont)}</dd>
    <dt>本文フォント</dt><dd>${escapeHtml(stats.bodyFont)}</dd>
    <dt>置換されたフォント</dt><dd>${convertedFontCount} フォント</dd>
    <dt>typeface 置換数</dt><dd>${stats.replacements}</dd>
    <dt>埋め込みフォント削除</dt><dd>${embeddedLine}</dd>
    ${cleanupLines.join("")}
    ${sizeLine}
    ${modeLine}
    ${backupLine}
    <dt>出力ファイル</dt><dd>${escapeHtml(stats.outputName)}</dd>
  `;
  resultPanel.hidden = false;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function outputFilename(originalName) {
  return `${originalName.replace(/\.pptx$/i, "")}_finalized.pptx`;
}

function clearAll() {
  selectedFile = null;
  sourceFileHandle = null;
  parentDirHandle = null;
  currentFileSize = 0;
  fileInput.value = "";
  pptxFonts = [];
  mediaAnalysis = null;
  cleanupPlan = null;
  pptxZipCache = null;
  revokeMediaThumbUrls();
  analysisStack.hidden = true;
  cleanupPanel.hidden = true;
  loadFromFile(null, null);
}

async function runDownload() {
  if (!selectedFile) return;
  hideError();
  downloadBtn.disabled = true;
  downloadBtn.textContent = "仕上げ中…";

  try {
    const fonts = getTargetFonts();
    const options = getFinalizeOptions();
    const result = await finalizePptx(selectedFile, fonts, options);
    const outputName = outputFilename(selectedFile.name);
    downloadBlob(result.blob, outputName);
    renderStats({ ...result, outputName, saveMode: "download" });
  } catch (err) {
    showError(err.message || String(err));
  } finally {
    downloadBtn.disabled = !selectedFile;
    downloadBtn.textContent = "仕上げてダウンロード";
  }
}

async function runOverwrite() {
  if (!selectedFile || !sourceFileHandle) return;
  hideError();
  overwriteBtn.disabled = true;
  overwriteBtn.textContent = "保存中…";

  try {
    const result = await overwriteOriginal();
    renderStats(result);
  } catch (err) {
    if (err.name === "AbortError") return;
    showError(err.message || String(err));
  } finally {
    overwriteBtn.disabled = !selectedFile || !sourceFileHandle;
    overwriteBtn.textContent = "仕上げて上書き（バックアップ付き）";
  }
}

function onCleanupOptionChange() {
  saveSettings();
  renderCleanupPreview();
}

optRemoveOrphanMedia.addEventListener("change", onCleanupOptionChange);
optRemoveUnusedStructure.addEventListener("change", onCleanupOptionChange);
optRemoveNotes.addEventListener("change", onCleanupOptionChange);
optRemoveProperties.addEventListener("change", onCleanupOptionChange);

mediaTableBody.addEventListener("mouseover", (event) => {
  const cell = event.target.closest(".media-name[data-path]");
  if (!cell) return;
  showMediaThumb(cell.dataset.path, event.clientX, event.clientY);
});

mediaTableBody.addEventListener("mousemove", (event) => {
  if (mediaThumbTooltip.hidden) return;
  const cell = event.target.closest(".media-name[data-path]");
  if (!cell) return;
  positionMediaThumbTooltip(event.clientX, event.clientY);
});

mediaTableBody.addEventListener("mouseout", (event) => {
  const fromCell = event.target.closest(".media-name[data-path]");
  if (!fromCell) return;
  const toCell = event.relatedTarget?.closest?.(".media-name[data-path]");
  if (toCell === fromCell) return;
  hideMediaThumb();
});

openBtn.addEventListener("click", openFileWithHandle);
pickBtn.addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", () => {
  const file = fileInput.files[0];
  if (file) loadFromFile(file, null);
});

["dragenter", "dragover"].forEach((ev) => {
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  });
});

["dragleave", "drop"].forEach((ev) => {
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
  });
});

dropzone.addEventListener("drop", (e) => {
  loadFromDrop(e.dataTransfer);
});

clearBtn.addEventListener("click", clearAll);
bindFontPicker(
  titleFontSelect,
  titleCustomFontInput,
  () => selectedTitleFont,
  (value) => { selectedTitleFont = value; }
);
bindFontPicker(
  bodyFontSelect,
  bodyCustomFontInput,
  () => selectedBodyFont,
  (value) => { selectedBodyFont = value; }
);
downloadBtn.addEventListener("click", runDownload);
overwriteBtn.addEventListener("click", runOverwrite);

[titleCustomFontInput, bodyCustomFontInput].forEach((input) => {
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && selectedFile && !downloadBtn.disabled) {
      if (sourceFileHandle && !overwriteBtn.disabled) {
        runOverwrite();
      } else {
        runDownload();
      }
    }
  });
});

loadSettings();
rebuildFontDropdowns();
setActionState();
showAppVersion();

async function showAppVersion() {
  const el = document.getElementById("appVersion");
  if (!el) return;
  try {
    const res = await fetch("version.json", { cache: "no-store" });
    if (!res.ok) throw new Error("version fetch failed");
    const data = await res.json();
    const version = String(data.version || "").trim();
    if (!version) throw new Error("empty version");
    const released = String(data.released || "").trim();
    el.textContent = released ? `v${version} (${released})` : `v${version}`;
    el.hidden = false;
  } catch {
    el.hidden = true;
  }
}
