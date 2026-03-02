// JSON to Excel 原生 Rust 版 - 前端逻辑

const { invoke } = window.__TAURI__.core;
const { open, save } = window.__TAURI__.dialog;

// ============ 多语言管理 ============

const LANG_KEY = "json2excel-lang";
let currentLang = localStorage.getItem(LANG_KEY) || "zh";

const i18nData = {
  zh: {
    title: "JSON to Excel 转换工具",
    themeAuto: "跟随系统",
    themeLight: "浅色模式",
    themeDark: "深色模式",
    subtitle: "多种方式导入 JSON，一键转换为 Excel",
    tabFile: "文件",
    inputPlaceholder: "选择或输入 JSON 文件路径...",
    browse: "浏览",
    curlPlaceholder: "粘贴 curl 或 fetch 命令...\n\n示例:\ncurl -X GET 'https://api.example.com/data' \\\n  -H 'Authorization: Bearer token'\n\n或 fetch 格式:\nfetch('https://api.example.com/data', {\n  headers: { 'Authorization': 'Bearer token' }\n})",
    execute: "执行请求",
    jsonPlaceholder: "直接粘贴 JSON 内容...\n\n支持数组或对象结构:\n[\n  { \"name\": \"Alice\", \"age\": 30 },\n  { \"name\": \"Bob\", \"age\": 25 }\n]",
    parseJson: "解析 JSON",
    fieldSelectTitle: "📋 发现多个数组结构，请选择要导出为 Sheet 的字段：",
    outputLabel: "输出文件",
    outputPlaceholder: "选择或输入保存位置...",
    startConvert: "开始转换",
    converting: "转换中...",
    revealInExplorer: "在文件夹中显示",
    errParseFile: "解析文件...",
    errParseJson: "解析 JSON...",
    errParseFail: "❌ 解析失败",
    errNoArray: "⚠️ 未在 JSON 中找到任何数组字段",
    errNeedCurl: "⚠️ 请先粘贴 curl 或 fetch 命令",
    executing: "执行请求...",
    reqSuccess: "✅ 请求成功，正在解析字段...",
    errNeedJson: "⚠️ 请先粘贴 JSON 内容",
    exporting: "正在导出...",
    errOpenViewer: "❌ 无法打开文件管理器",
    items: "项",
    matrix: "矩阵",
    langToggleTooltip: "Switch to English",
    langText: "EN",
    fileFilters: "JSON 文件",
    excelFilters: "Excel 文件",
    selJsonTitle: "选择 JSON 文件",
    selSaveTitle: "选择保存位置",
    convertSuccess: "✅ "
  },
  en: {
    title: "JSON to Excel Converter",
    themeAuto: "Auto Theme",
    themeLight: "Light Theme",
    themeDark: "Dark Theme",
    subtitle: "Import JSON in multiple ways, convert to Excel easily",
    tabFile: "File",
    inputPlaceholder: "Select or input JSON file path...",
    browse: "Browse",
    curlPlaceholder: "Paste curl or fetch command...\n\nExample:\ncurl -X GET 'https://api.example.com/data' \\\n  -H 'Authorization: Bearer token'\n\nOr fetch format:\nfetch('https://api.example.com/data', {\n  headers: { 'Authorization': 'Bearer token' }\n})",
    execute: "Execute Request",
    jsonPlaceholder: "Paste JSON content directly...\n\nSupports array or object structures:\n[\n  { \"name\": \"Alice\", \"age\": 30 },\n  { \"name\": \"Bob\", \"age\": 25 }\n]",
    parseJson: "Parse JSON",
    fieldSelectTitle: "📋 Found multiple array structures, select fields for Sheets:",
    outputLabel: "Output File",
    outputPlaceholder: "Select or input save location...",
    startConvert: "Start Convert",
    converting: "Converting...",
    revealInExplorer: "Reveal in Explorer",
    errParseFile: "Parsing file...",
    errParseJson: "Parsing JSON...",
    errParseFail: "❌ Parse failed",
    errNoArray: "⚠️ No array fields found in JSON",
    errNeedCurl: "⚠️ Please paste a curl or fetch command first",
    executing: "Executing request...",
    reqSuccess: "✅ Request successful, parsing fields...",
    errNeedJson: "⚠️ Please paste JSON content first",
    exporting: "Exporting...",
    errOpenViewer: "❌ Cannot open file manager",
    items: "items",
    matrix: "matrix",
    langToggleTooltip: "切换到中文",
    langText: "中文",
    fileFilters: "JSON Files",
    excelFilters: "Excel Files",
    selJsonTitle: "Select JSON File",
    selSaveTitle: "Select Save Location",
    convertSuccess: "✅ "
  }
};

function t(key) {
  return i18nData[currentLang]?.[key] || key;
}

function updateDOMTranslations() {
  document.querySelectorAll("[data-i18n]").forEach(el => {
    const key = el.getAttribute("data-i18n");
    if (i18nData[currentLang][key]) {
      el.textContent = i18nData[currentLang][key];
    }
  });
  document.querySelectorAll("[data-i18n-placeholder]").forEach(el => {
    const key = el.getAttribute("data-i18n-placeholder");
    if (i18nData[currentLang][key]) {
      el.setAttribute("placeholder", i18nData[currentLang][key]);
    }
  });

  const btnLang = document.querySelector("#btn-lang");
  if (btnLang) {
    btnLang.setAttribute("data-tooltip", t("langToggleTooltip"));
    btnLang.querySelector(".lang-text").textContent = t("langText");
  }

  // Update theme tooltip immediately if possible
  const btnTheme = document.querySelector("#btn-theme");
  if (btnTheme) {
    const labelMapping = { dark: t("themeDark"), light: t("themeLight"), auto: t("themeAuto") };
    btnTheme.setAttribute("data-tooltip", labelMapping[currentMode] || t("themeAuto"));
  }

  // Update field selector if active
  if (rawFieldsOptions.length > 0) {
    showFieldSelector(rawFieldsOptions);
  }

  updateConvertButton();
}

function toggleLang() {
  currentLang = currentLang === "zh" ? "en" : "zh";
  localStorage.setItem(LANG_KEY, currentLang);
  updateDOMTranslations();
}

// ============ 主题管理 ============

const THEME_KEY = "json2excel-theme-mode";
const MODES = ["dark", "light", "auto"];

let currentMode = "auto";
let systemDarkQuery = window.matchMedia("(prefers-color-scheme: dark)");

/**
 * 初始化主题：从 localStorage 恢复偏好并应用
 */
function initTheme() {
  const saved = localStorage.getItem(THEME_KEY);
  currentMode = saved && MODES.includes(saved) ? saved : "auto";
  applyMode(currentMode);

  // 监听系统主题变化（仅 auto 模式生效）
  systemDarkQuery.addEventListener("change", () => {
    if (currentMode === "auto") {
      applyThemeAttribute();
    }
  });
}

/**
 * 应用主题模式
 */
function applyMode(mode) {
  currentMode = mode;
  localStorage.setItem(THEME_KEY, mode);

  // 设置 data-theme-mode 用于控制图标显示
  document.documentElement.setAttribute("data-theme-mode", mode);

  // 应用实际主题
  applyThemeAttribute();

  // 更新 tooltip
  const btn = document.querySelector("#btn-theme");
  if (btn) {
    const labelMapping = { dark: t("themeDark"), light: t("themeLight"), auto: t("themeAuto") };
    btn.setAttribute("data-tooltip", labelMapping[mode] || t("themeAuto"));
  }
}

/**
 * 根据当前模式设置 data-theme 属性
 */
function applyThemeAttribute() {
  let theme;
  if (currentMode === "auto") {
    theme = systemDarkQuery.matches ? "dark" : "light";
  } else {
    theme = currentMode;
  }

  if (theme === "dark") {
    document.documentElement.removeAttribute("data-theme");
  } else {
    document.documentElement.setAttribute("data-theme", theme);
  }
}

/**
 * 循环切换主题: dark → light → auto → dark
 */
function cycleTheme() {
  const nextIndex = (MODES.indexOf(currentMode) + 1) % MODES.length;
  applyMode(MODES[nextIndex]);
}

// ============ 输入模式管理 ============

let activeTab = "file"; // 当前激活的 Tab: file / curl / json
let currentJsonContent = ""; // curl 或 json 模式下保存的 JSON 内容

// DOM 元素
let inputPathEl;
let outputPathEl;
let btnConvert;
let statusMsgEl;
let btnReveal;
let curlInputEl;
let jsonInputEl;
let rawFieldsOptions = []; // 存储从后端拉取的字段信息

/**
 * 切换输入 Tab
 */
function switchTab(tabName) {
  activeTab = tabName;

  // 更新 Tab 按钮状态
  document.querySelectorAll(".input-tab").forEach(tab => {
    tab.classList.toggle("active", tab.dataset.tab === tabName);
  });

  // 切换面板
  document.querySelectorAll(".tab-panel").forEach(panel => {
    panel.classList.remove("active");
  });
  const panel = document.querySelector(`#panel-${tabName}`);
  if (panel) panel.classList.add("active");

  // 切换 Tab 时清除状态
  hideFieldSelector();
  showStatus("", "none");
  btnReveal.style.display = "none";
  rawFieldsOptions = [];
  currentJsonContent = "";
  updateConvertButton();

  // 切换后滚动到底部，确保能看到全部内容
  setTimeout(scrollToBottom, 50);
}

/**
 * 选择输入 JSON 文件并获取字段
 */
async function selectInputFile() {
  const path = await open({
    title: t("selJsonTitle"),
    filters: [{ name: t("fileFilters"), extensions: ["json"] }],
    multiple: false,
  });
  if (path) {
    inputPathEl.value = path;
    if (!outputPathEl.value) {
      outputPathEl.value = path.replace(/\.json$/i, ".xlsx");
    }

    // 载入字段
    await loadFieldsFromFile(path);
  }
}

/**
 * 从文件路径获取字段（文件模式）
 */
async function loadFieldsFromFile(input) {
  setLoading(true, t("errParseFile"));
  hideFieldSelector();
  showStatus("", "none");
  btnReveal.style.display = "none";
  rawFieldsOptions = [];

  try {
    const fields = await invoke("get_json_fields", { input });
    handleFieldsResult(fields);
  } catch (err) {
    showStatus(`${t("errParseFail")}: ${err}`, "error");
    btnConvert.disabled = true;
    // inputPathEl.value = ""; // Removed to allow users to fix the path
  } finally {
    setLoading(false);
  }
}

/**
 * 从 JSON 内容获取字段（curl/json 模式）
 */
async function loadFieldsFromContent(content) {
  setLoading(true, t("errParseJson"));
  hideFieldSelector();
  showStatus("", "none");
  btnReveal.style.display = "none";
  rawFieldsOptions = [];
  currentJsonContent = content;

  try {
    const fields = await invoke("get_json_fields_from_content", { content });
    handleFieldsResult(fields);
  } catch (err) {
    showStatus(`${t("errParseFail")}: ${err}`, "error");
    btnConvert.disabled = true;
    currentJsonContent = "";
  } finally {
    setLoading(false);
  }
}

/**
 * 处理字段结果（共用逻辑）
 */
function handleFieldsResult(fields) {
  if (fields.length === 0) {
    showStatus(t("errNoArray"), "error");
    btnConvert.disabled = true;
    return;
  }

  rawFieldsOptions = fields;

  if (fields.length === 1 && fields[0].path === "root") {
    updateConvertButton();
  } else {
    showFieldSelector(fields);
    updateConvertButton();
  }
}

/**
 * 执行 curl/fetch 命令
 */
async function executeCurl() {
  const command = curlInputEl.value.trim();
  if (!command) {
    showStatus(t("errNeedCurl"), "error");
    return;
  }

  setLoading(true, t("executing"));
  showStatus("", "none");
  hideFieldSelector();
  rawFieldsOptions = [];
  currentJsonContent = "";

  try {
    const jsonContent = await invoke("execute_curl", { command });
    currentJsonContent = jsonContent;

    showStatus(t("reqSuccess"), "success");

    // 解析字段
    await loadFieldsFromContent(jsonContent);
  } catch (err) {
    showStatus(`❌ ${err}`, "error");
    btnConvert.disabled = true;
  } finally {
    setLoading(false);
  }
}

/**
 * 解析粘贴的 JSON
 */
async function parseJsonInput() {
  const jsonText = jsonInputEl.value.trim();
  if (!jsonText) {
    showStatus(t("errNeedJson"), "error");
    return;
  }

  // 尝试格式化 JSON（如果是有效 JSON 的话）
  try {
    const parsed = JSON.parse(jsonText);
    jsonInputEl.value = JSON.stringify(parsed, null, 2);
  } catch {
    // 格式化失败不影响后续逻辑，后端会给出更准确的错误
  }

  await loadFieldsFromContent(jsonText);
}

/**
 * 显示状态消息
 */
function showStatus(message, type) {
  if (type === "none") {
    statusMsgEl.style.display = "none";
    return;
  }
  statusMsgEl.textContent = message;
  statusMsgEl.className = `status-msg ${type}`;
  statusMsgEl.style.display = "block";

  // 消息出现后自动滚动到下方
  setTimeout(scrollToBottom, 50);
}

/**
 * 隐藏字段选择面板
 */
function hideFieldSelector() {
  const panel = document.querySelector("#field-selector");
  if (panel) panel.style.display = "none";
}

/**
 * 显示字段选择面板供用户勾选导出的字段
 */
function showFieldSelector(fields) {
  const panel = document.querySelector("#field-selector");
  const list = document.querySelector("#field-list");

  list.innerHTML = "";
  fields.forEach((field, index) => {
    const label = document.createElement("label");
    label.className = "field-option";
    label.innerHTML = `
      <input type="checkbox" value="${field.path}" checked />
      <span class="field-num">[${index + 1}]</span>
      <span class="field-name">${field.path}</span>
      <span class="field-meta">(${field.count} ${t("items")}${field.is_matrix ? ', ' + t("matrix") : ''})</span>
    `;
    list.appendChild(label);
  });

  const checkboxes = list.querySelectorAll("input[type='checkbox']");
  checkboxes.forEach(cb => cb.addEventListener("change", updateConvertButton));

  panel.style.display = "block";

  // 展开列表后滚动到底部
  setTimeout(scrollToBottom, 100);
}

/**
 * 选取输出文件位置
 */
async function selectOutputFile() {
  const path = await save({
    title: t("selSaveTitle"),
    defaultPath: outputPathEl.value || undefined,
    filters: [{ name: t("excelFilters"), extensions: ["xlsx"] }],
  });
  if (path) {
    outputPathEl.value = path;
    updateConvertButton();
  }
}

/**
 * 更新转换按钮启用状态
 */
function updateConvertButton() {
  let hasInput = false;
  const hasOutput = !!outputPathEl.value;
  let hasFields = false;

  // 根据当前 Tab 判断是否有输入
  switch (activeTab) {
    case "file":
      hasInput = !!inputPathEl.value;
      break;
    case "curl":
      hasInput = !!currentJsonContent; // curl 执行后才有内容
      break;
    case "json":
      hasInput = !!currentJsonContent; // 解析后才有内容
      break;
  }

  const panel = document.querySelector("#field-selector");
  if (panel.style.display === "block") {
    const checked = document.querySelectorAll("#field-list input[type='checkbox']:checked");
    hasFields = checked.length > 0;
  } else {
    hasFields = rawFieldsOptions.length > 0;
  }

  btnConvert.disabled = !(hasInput && hasOutput && hasFields);
}

/**
 * 设置按钮的 loading 状态
 */
function setLoading(loading, text = t("converting")) {
  const btnText = btnConvert.querySelector(".btn-text");
  const btnLoading = btnConvert.querySelector(".btn-loading");

  if (loading) {
    btnText.style.display = "none";
    btnLoading.innerHTML = `
      <svg class="spinner" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/>
      </svg>
      ${text}
    `;
    btnLoading.style.display = "flex";
    btnConvert.disabled = true;
  } else {
    btnText.style.display = "inline";
    btnLoading.style.display = "none";
    updateConvertButton();
  }
}

/**
 * 获取选中的字段路径
 */
function getSelectedPaths() {
  const panel = document.querySelector("#field-selector");
  if (panel.style.display === "block") {
    const checked = document.querySelectorAll("#field-list input[type='checkbox']:checked");
    return Array.from(checked).map(cb => cb.value);
  }
  return rawFieldsOptions.map(f => f.path);
}

/**
 * 执行转换导出
 */
async function convert() {
  const output = outputPathEl.value;
  const selectedPaths = getSelectedPaths();

  if (!output || selectedPaths.length === 0) return;

  setLoading(true, t("exporting"));
  showStatus("", "none");
  btnReveal.style.display = "none";

  try {
    let result;
    if (activeTab === "file") {
      // 文件模式：从文件路径导出
      const input = inputPathEl.value;
      if (!input) return;
      result = await invoke("export_to_excel", { input, output, selectedPaths });
    } else {
      // curl/json 模式：从内容导出
      if (!currentJsonContent) return;
      result = await invoke("export_to_excel_from_content", {
        content: currentJsonContent,
        output,
        selectedPaths,
      });
    }
    showStatus(`${t("convertSuccess")}${result}`, "success");
    btnReveal.style.display = "block";

    // 完成后确保看到结果详情
    setTimeout(scrollToBottom, 50);
  } catch (err) {
    showStatus(`❌ ${err}`, "error");
  } finally {
    setLoading(false);
  }
}

// ============ 初始化 ============

window.addEventListener("DOMContentLoaded", () => {
  // 主题初始化（尽早执行以避免闪烁）
  initTheme();

  inputPathEl = document.querySelector("#input-path");
  outputPathEl = document.querySelector("#output-path");
  btnConvert = document.querySelector("#btn-convert");
  statusMsgEl = document.querySelector("#status-msg");
  btnReveal = document.querySelector("#btn-reveal");
  curlInputEl = document.querySelector("#curl-input");
  jsonInputEl = document.querySelector("#json-input");

  // 文件模式
  document.querySelector("#btn-select-input").addEventListener("click", selectInputFile);

  // 优化：双击输入框也可以打开选择
  inputPathEl.addEventListener("dblclick", selectInputFile);

  // 优化：输入变化时更新按钮状态和处理手动输入
  let inputDebounceTimer;
  inputPathEl.addEventListener("input", () => {
    updateConvertButton();

    // 如果是手动输入且看起来像个路径，尝试自动解析
    clearTimeout(inputDebounceTimer);
    inputDebounceTimer = setTimeout(() => {
      const val = inputPathEl.value.trim();
      if (val && (val.toLowerCase().endsWith(".json") || val.includes("/") || val.includes("\\"))) {
        loadFieldsFromFile(val);
      }
    }, 500);
  });

  outputPathEl.addEventListener("input", updateConvertButton);

  // 输出文件
  document.querySelector("#btn-select-output").addEventListener("click", selectOutputFile);

  // 转换按钮
  btnConvert.addEventListener("click", convert);

  // 文件夹中显示
  btnReveal.addEventListener("click", () => {
    invoke("reveal_in_explorer", { path: outputPathEl.value }).catch(err => {
      showStatus(`${t("errOpenViewer")}: ${err}`, "error");
    });
  });

  // Tab 切换
  document.querySelectorAll(".input-tab").forEach(tab => {
    tab.addEventListener("click", () => switchTab(tab.dataset.tab));
  });

  // Curl 执行按钮
  document.querySelector("#btn-exec-curl").addEventListener("click", executeCurl);

  // JSON 解析按钮
  document.querySelector("#btn-parse-json").addEventListener("click", parseJsonInput);

  // 主题切换按钮
  document.querySelector("#btn-theme").addEventListener("click", cycleTheme);

  // 语言切换按钮
  const btnLang = document.querySelector("#btn-lang");
  if (btnLang) {
    btnLang.addEventListener("click", toggleLang);
  }

  // 自定义标题栏拖拽逻辑
  const titlebar = document.querySelector(".titlebar");
  if (titlebar) {
    titlebar.addEventListener("pointerdown", (e) => {
      // 防止右键点击或其他特殊点击触发拖动
      if (e.button !== 0) return;

      try {
        if (window.__TAURI__.window && window.__TAURI__.window.getCurrentWindow) {
          window.__TAURI__.window.getCurrentWindow().startDragging();
        } else {
          // Fallback to core invoke if window API is not directly exposed
          invoke("plugin:window|start_dragging").catch(err => {
            console.warn("Drag error fallback:", err);
          });
        }
      } catch (err) {
        console.warn("Failed to drag window:", err);
      }
    });
  }

  // 初始化 DOM 翻译
  updateDOMTranslations();
});

/**
 * 自动滚动到页面底部
 */
function scrollToBottom() {
  window.scrollTo({
    top: document.body.scrollHeight,
    behavior: "smooth"
  });
}