const state = { jobId: null, filename: "", sheets: [], previews: {}, values: [], maxStep: 1, executing: false };
const titles = ["", "选择 Excel 文件", "选择 Sheet", "配置表头与字段", "选择拆分值", "设置输出", "执行结果"];

const byId = (id) => document.getElementById(id);
const escapeHtml = (value) => String(value ?? "").replace(/[&<>'"]/g, (char) => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[char]));

function showAlert(message) {
  const alert = byId("alert");
  alert.textContent = message;
  alert.hidden = !message;
}

function setBusy(isBusy, message = "处理中") {
  byId("busy").textContent = message;
  byId("busy").hidden = !isBusy;
  document.querySelectorAll("button").forEach((button) => {
    const step = Number(button.dataset.step || 0);
    button.disabled = isBusy || (step > 0 && step > state.maxStep);
  });
}

function goStep(step) {
  document.querySelectorAll(".panel").forEach((panel) => panel.classList.toggle("active", Number(panel.dataset.panel) === step));
  document.querySelectorAll(".step").forEach((button) => button.classList.toggle("active", Number(button.dataset.step) === step));
  byId("page-title").textContent = titles[step];
  document.querySelectorAll(".step").forEach((button) => { button.disabled = Number(button.dataset.step) > state.maxStep; });
  showAlert("");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

async function api(path, options = {}, showGlobalBusy = true) {
  if (showGlobalBusy) setBusy(true);
  try {
    const response = await fetch(path, options);
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "请求失败");
    return payload;
  } finally {
    if (showGlobalBusy) setBusy(false);
  }
}

byId("excel-file").addEventListener("change", (event) => {
  const file = event.target.files[0];
  byId("upload-zone").classList.toggle("has-file", Boolean(file));
  byId("selected-file").textContent = file?.name || "尚未选择文件";
  byId("selected-file-size").textContent = file ? `${formatFileSize(file.size)} · 已选择，可加载工作簿` : "支持最大 100 MB 的 .xlsx 文件";
  byId("file-picker").textContent = file ? "重新选择文件" : "选择 .xlsx 文件";
});

function formatFileSize(bytes) {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

let pendingPasswordJobId = null;

byId("load-file").addEventListener("click", async () => {
  const file = byId("excel-file").files[0];
  if (!file) return showAlert("请选择 .xlsx 文件");
  const form = new FormData();
  form.append("file", file);
  if (file.path) {
    const sep = file.path.includes("/") ? "/" : "\\";
    form.append("source_dir", file.path.substring(0, file.path.lastIndexOf(sep)));
  }
  try {
    const payload = await api("/api/load", { method: "POST", body: form });
    if (payload.needs_password) {
      pendingPasswordJobId = payload.job_id;
      byId("password-modal-filename").textContent = payload.filename;
      byId("password-input").value = "";
      byId("password-modal").classList.add("open");
      byId("password-input").focus();
      return;
    }
    finishLoad(payload);
  } catch (error) { showAlert(error.message); }
});

byId("password-confirm").addEventListener("click", async () => {
  const password = byId("password-input").value;
  if (!password) return;
  try {
    const payload = await api("/api/load-with-password", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ job_id: pendingPasswordJobId, password }),
    });
    byId("password-modal").classList.remove("open");
    pendingPasswordJobId = null;
    finishLoad(payload);
  } catch (error) {
    showAlert(error.message);
    byId("password-input").value = "";
    byId("password-input").focus();
  }
});

byId("password-cancel").addEventListener("click", () => {
  byId("password-modal").classList.remove("open");
  pendingPasswordJobId = null;
});

byId("password-input").addEventListener("keydown", (event) => {
  if (event.key === "Enter") byId("password-confirm").click();
});

function finishLoad(payload) {
  Object.assign(state, { jobId: payload.job_id, filename: payload.filename, sheets: payload.sheets, previews: {}, values: [], maxStep: 2 });
  byId("file-status").textContent = `已加载：${payload.filename} · ${payload.sheets.length} 个 Sheet`;
  byId("file-status").classList.add("loaded");
  byId("output-dir").value = payload.default_output_dir;
  byId("sheet-list").innerHTML = payload.sheets.map((name) => `<label class="check-item"><input type="checkbox" value="${escapeHtml(name)}"><span>${escapeHtml(name)}</span></label>`).join("");
  goStep(2);
}

function selectedSheets() {
  return [...byId("sheet-list").querySelectorAll("input:checked")].map((input) => input.value);
}

byId("select-all-sheets").addEventListener("change", (event) => {
  byId("sheet-list").querySelectorAll("input").forEach((input) => {
    input.checked = event.target.checked;
  });
});

byId("configure-sheets").addEventListener("click", async () => {
  const selected = selectedSheets();
  if (!selected.length) return showAlert("至少选择一个 Sheet");
  try {
    for (const sheetName of selected) {
      if (!state.previews[sheetName]) {
        state.previews[sheetName] = await api("/api/preview", {
          method: "POST", headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ job_id: state.jobId, sheet_name: sheetName, start_row: 1, max_rows: 100 })
        });
      }
    }
    renderSheetConfigs(selected);
    state.maxStep = Math.max(state.maxStep, 3);
    goStep(3);
  } catch (error) { showAlert(error.message); }
});

function renderSheetConfigs(selected) {
  byId("sheet-configs").innerHTML = selected.map((name, index) => {
    const preview = state.previews[name];
    preview.loading = false;
    const headerOptions = Array.from({ length: Math.min(15, preview.total_rows) }, (_, rowIndex) => rowIndex + 1).map((row) => `<option value="${row}" ${row === preview.suggested_header_row ? "selected" : ""}>第 ${row} 行</option>`).join("");
    const modeOptions = `<option value="full" selected>不拆分，完整保留</option><option value="reference">作为基准 Sheet</option><option value="linked">按关联键匹配</option><option value="direct">直接按本 Sheet 字段拆分</option>`;
    const table = renderPreviewRows(preview.rows, preview.start_row);
    return `<article class="sheet-config" data-sheet="${escapeHtml(name)}" data-index="${index}"><h2>${escapeHtml(name)}</h2><div class="config-controls"><label>处理方式<select class="sheet-mode">${modeOptions}</select></label><label>表头行<select class="header-row">${headerOptions}</select></label><label class="split-column-control">拆分字段<select class="split-column"></select></label><label class="key-column-control">关联键字段<select class="key-column"></select></label></div><div class="preview-meta">已加载 <span class="loaded-rows">${preview.end_row}</span> / ${preview.total_rows} 行</div><div class="preview-wrap"><table class="preview-table"><tbody>${table}</tbody></table></div></article>`;
  }).join("");
  document.querySelectorAll(".sheet-config").forEach((block) => {
    block.querySelector(".header-row").addEventListener("change", () => updateColumns(block));
    block.querySelector(".sheet-mode").addEventListener("change", () => updateModeControls(block));
    block.querySelector(".preview-wrap").addEventListener("scroll", () => maybeLoadMorePreview(block));
    updateColumns(block);
  });
}

function renderPreviewRows(rows, startRow) {
  return rows.map((row, rowIndex) => `<tr><th class="row-number">${startRow + rowIndex}</th>${row.map((value) => `<td>${escapeHtml(value)}</td>`).join("")}</tr>`).join("");
}

async function maybeLoadMorePreview(block) {
  const container = block.querySelector(".preview-wrap");
  const preview = state.previews[block.dataset.sheet];
  const nearBottom = container.scrollTop + container.clientHeight >= container.scrollHeight - 80;
  if (!nearBottom || !preview.has_more || preview.loading) return;

  preview.loading = true;
  try {
    const nextPage = await api("/api/preview", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        job_id: state.jobId,
        sheet_name: block.dataset.sheet,
        start_row: preview.end_row + 1,
        max_rows: 100,
      }),
    }, false);
    block.querySelector("tbody").insertAdjacentHTML("beforeend", renderPreviewRows(nextPage.rows, nextPage.start_row));
    preview.rows.push(...nextPage.rows);
    preview.end_row = nextPage.end_row;
    preview.has_more = nextPage.has_more;
    block.querySelector(".loaded-rows").textContent = preview.end_row;
  } catch (error) {
    showAlert(error.message);
  } finally {
    preview.loading = false;
  }
}

function updateColumns(block) {
  const preview = state.previews[block.dataset.sheet];
  const rowIndex = Number(block.querySelector(".header-row").value) - preview.start_row;
  const row = preview.rows[rowIndex] || [];
  let last = row.length;
  while (last > 0 && (row[last - 1] === null || row[last - 1] === "")) last -= 1;
  const options = row.slice(0, last).map((value, index) => {
    const letter = columnLetter(index + 1);
    const label = `${letter} - ${value ?? "(空列)"}`;
    return `<option value="${index + 1}" data-label="${escapeHtml(label)}">${escapeHtml(label)}</option>`;
  }).join("");
  block.querySelector(".split-column").innerHTML = options;
  block.querySelector(".key-column").innerHTML = options;
  updateModeControls(block);
}

function updateModeControls(block) {
  const mode = block.querySelector(".sheet-mode").value;
  const usesSplitColumn = ["direct", "reference"].includes(mode);
  const usesKeyColumn = ["reference", "linked"].includes(mode);
  const splitColumn = block.querySelector(".split-column");
  const keyColumn = block.querySelector(".key-column");
  block.querySelector(".split-column-control").hidden = !usesSplitColumn;
  block.querySelector(".key-column-control").hidden = !usesKeyColumn;
  splitColumn.disabled = !usesSplitColumn;
  keyColumn.disabled = !usesKeyColumn;
}

function columnLetter(index) {
  let result = "";
  while (index > 0) { index -= 1; result = String.fromCharCode(65 + (index % 26)) + result; index = Math.floor(index / 26); }
  return result;
}

function sheetConfigs() {
  return [...document.querySelectorAll(".sheet-config")].map((block) => {
    const mode = block.querySelector(".sheet-mode").value;
    const splitColumn = block.querySelector(".split-column");
    const keyColumn = block.querySelector(".key-column");
    const usesSplitColumn = ["direct", "reference"].includes(mode);
    const usesKeyColumn = ["reference", "linked"].includes(mode);
    return {
      sheet_name: block.dataset.sheet,
      header_row: Number(block.querySelector(".header-row").value),
      mode,
      split_column_idx: usesSplitColumn ? Number(splitColumn.value) : null,
      split_column_label: usesSplitColumn ? splitColumn.selectedOptions[0]?.dataset.label || "" : "",
      key_column_idx: usesKeyColumn ? Number(keyColumn.value) : null,
      key_column_label: usesKeyColumn ? keyColumn.selectedOptions[0]?.dataset.label || "" : "",
    };
  });
}

function validateSheetConfigs(configs) {
  const references = configs.filter((config) => config.mode === "reference");
  if (references.length > 1) throw new Error("一次任务只能选择一个基准 Sheet");
  if (configs.some((config) => config.mode === "linked") && references.length !== 1) {
    throw new Error("按关联键匹配时必须选择一个基准 Sheet");
  }
}

byId("load-values").addEventListener("click", async () => {
  try {
    const configs = sheetConfigs();
    validateSheetConfigs(configs);
    const payload = await api("/api/split-values", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ job_id: state.jobId, sheet_configs: configs }) });
    state.values = payload.values;
    byId("split-values").innerHTML = payload.values.map((value) => `<label class="check-item"><input type="checkbox" value="${escapeHtml(value)}" checked><span>${escapeHtml(value)}</span></label>`).join("");
    state.maxStep = Math.max(state.maxStep, 4);
    goStep(4);
  } catch (error) { showAlert(error.message); }
});

byId("select-all-values").addEventListener("click", () => byId("split-values").querySelectorAll("input").forEach((input) => { input.checked = true; }));
byId("clear-values").addEventListener("click", () => byId("split-values").querySelectorAll("input").forEach((input) => { input.checked = false; }));
byId("to-output").addEventListener("click", () => {
  if (!byId("split-values").querySelector("input:checked")) return showAlert("至少选择一个拆分值");
  state.maxStep = Math.max(state.maxStep, 5);
  goStep(5);
});

byId("browse-output-dir").addEventListener("click", async () => {
  setBusy(true, "请选择输出目录");
  try {
    const payload = await api("/api/select-output-dir", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ current_path: byId("output-dir").value }),
    }, false);
    if (payload.selected) byId("output-dir").value = payload.path;
  } catch (error) {
    showAlert(error.message);
  } finally {
    setBusy(false);
  }
});

byId("output-encrypt").addEventListener("change", (event) => {
  byId("output-password-group").hidden = !event.target.checked;
});

byId("execute").addEventListener("click", async () => {
  const selectedValues = [...byId("split-values").querySelectorAll("input:checked")].map((input) => input.value);
  const outputTypes = [...document.querySelectorAll('input[name="output-type"]:checked')].map((input) => input.value);
  if (!outputTypes.length) return showAlert("至少选择一种输出版本");
  try {
    const configs = sheetConfigs();
    validateSheetConfigs(configs);
    const encryptChecked = byId("output-encrypt").checked;
    const outputPassword = encryptChecked ? byId("output-password").value : "";
    if (encryptChecked && !outputPassword) return showAlert("请输入加密密码");
    state.maxStep = 6;
    state.executing = true;
    goStep(6);
    prepareProgress();
    setBusy(true, "拆分任务执行中");
    const started = await api("/api/execute", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ job_id: state.jobId, sheet_configs: configs, split_mode: "selected", selected_split_values: selectedValues, output_types: outputTypes, output_dir: byId("output-dir").value, filename_template: byId("filename-template").value, overwrite: byId("overwrite").checked, background: true, output_password: outputPassword }) }, false);
    await pollExecution(started.progress_url);
  } catch (error) {
    updateProgress(0, "拆分失败");
    showAlert(error.message);
  } finally {
    state.executing = false;
    setBusy(false);
  }
});

function prepareProgress() {
  byId("progress-panel").hidden = false;
  byId("result-content").hidden = true;
  updateProgress(0, "任务已提交，正在准备");
}

function updateProgress(percent, message) {
  const normalized = Math.max(0, Math.min(100, Number(percent) || 0));
  byId("execution-progress").value = normalized;
  byId("execution-progress").textContent = `${normalized}%`;
  byId("progress-percent").textContent = `${normalized}%`;
  byId("progress-message").textContent = message;
}

async function pollExecution(progressUrl) {
  while (true) {
    const snapshot = await api(progressUrl, {}, false);
    updateProgress(snapshot.progress, snapshot.message);
    if (snapshot.status === "complete") {
      renderResults(snapshot.result);
      return;
    }
    if (snapshot.status === "failed") throw new Error(snapshot.error || "拆分失败");
    await new Promise((resolve) => setTimeout(resolve, 500));
  }
}

function renderResults(payload) {
  byId("result-content").hidden = false;
  byId("progress-panel").hidden = false;
  updateProgress(100, "拆分完成");
  byId("result-count").textContent = payload.total_files;
  byId("discard-count").textContent = payload.total_discarded;
  byId("unmatched-count").textContent = payload.total_unmatched;
  byId("error-count").textContent = payload.errors.length;
  byId("result-list").innerHTML = payload.results.map((result) => {
    const stats = Object.entries(result.sheet_rows).map(([sheetName, kept]) => {
      const empty = result.discarded_empty_rows[sheetName] || 0;
      const unmatched = result.unmatched_key_rows[sheetName] || 0;
      return `<span>${escapeHtml(sheetName)}：保留 ${kept}，空键 ${empty}，未匹配 ${unmatched}</span>`;
    }).join("");
    const files = result.output_files.map((artifact) => {
      const outputLabel = artifact.output_type === "formula" ? "公式版" : "结果值版";
      return `<div class="result-item"><strong>${outputLabel}</strong><span class="result-path">${escapeHtml(artifact.output_file)}</span><button class="btn-action" data-action="open-file" data-path="${escapeHtml(artifact.output_file)}">打开文件</button><button class="btn-action" data-action="open-folder" data-path="${escapeHtml(artifact.output_file)}">打开所在文件夹</button></div>`;
    }).join("");
    return `<section class="result-group"><h2>${escapeHtml(result.split_value)}</h2><div class="result-stats">${stats}</div>${files}</section>`;
  }).join("");
  const messages = [...payload.warnings, ...payload.errors];
  byId("warning-list").innerHTML = messages.map((message) => `<div>${escapeHtml(message)}</div>`).join("");
}

byId("result-list").addEventListener("click", async (event) => {
  const button = event.target.closest(".btn-action");
  if (!button) return;
  const action = button.dataset.action;
  const path = button.dataset.path;
  try {
    await api(`/api/action/${action}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path }),
    }, false);
  } catch (error) {
    showAlert(error.message);
  }
});

document.querySelectorAll(".back").forEach((button) => button.addEventListener("click", () => goStep(Number(button.dataset.target))));
document.querySelectorAll(".step").forEach((button) => button.addEventListener("click", () => goStep(Number(button.dataset.step))));
byId("restart").addEventListener("click", () => window.location.reload());

/* ===== 合并模式 ===== */
const mergeState = { jobId: null, files: [], sheets: [], selectedSheets: [], failedUnlocks: [], maxStep: 1, executing: false, step: 1 };
const mergeTitles = ["", "添加文件", "选择 Sheet", "字段检查", "输出设置", "执行结果"];
let lastSplitStep = 1;

function setMode(mode) {
  document.querySelectorAll(".mode-switch input").forEach((input) => { input.checked = input.value === mode; });
  const isMerge = mode === "merge";
  byId("steps-split").hidden = isMerge;
  byId("steps-merge").hidden = !isMerge;
  if (isMerge) {
    lastSplitStep = Number(document.querySelector("#steps-split .step.active")?.dataset.step || 1);
    byId("file-status").textContent = mergeState.files.length ? `合并模式：已添加 ${mergeState.files.length} 个文件` : "合并模式：未添加文件";
    byId("file-status").classList.remove("loaded");
    mergeGoStep(mergeState.step);
  } else {
    byId("file-status").textContent = state.jobId ? `已加载：${state.filename} · ${state.sheets.length} 个 Sheet` : "未加载文件";
    byId("file-status").classList.toggle("loaded", Boolean(state.jobId));
    goStep(lastSplitStep);
  }
}

document.querySelectorAll(".mode-switch input").forEach((input) => {
  input.addEventListener("change", (event) => setMode(event.target.value));
});

function mergeGoStep(step) {
  mergeState.step = step;
  document.querySelectorAll(".panel").forEach((panel) => panel.classList.toggle("active", panel.dataset.panel === `merge-${step}`));
  document.querySelectorAll(".mstep").forEach((button) => button.classList.toggle("active", Number(button.dataset.mergeStep) === step));
  byId("page-title").textContent = mergeTitles[step];
  mergeRefreshNav();
  showAlert("");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function mergeRefreshNav() {
  document.querySelectorAll(".mstep").forEach((button) => { button.disabled = Number(button.dataset.mergeStep) > mergeState.maxStep; });
  byId("merge-to-sheets").disabled = !(mergeState.files.length >= 2 && !mergeState.files.some((file) => file.encrypted));
}

function mergeApi(path, body) {
  return api(path, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
}

byId("merge-file-input").addEventListener("change", async (event) => {
  const files = [...event.target.files];
  event.target.value = "";
  if (!files.length) return;
  const form = new FormData();
  files.forEach((file) => {
    form.append("files", file);
    form.append("source_paths", file.path || "");
  });
  if (mergeState.jobId) form.append("job_id", mergeState.jobId);
  try {
    const payload = await api("/api/merge/load", { method: "POST", body: form });
    mergeState.jobId = payload.job_id;
    mergeState.files = payload.files;
    renderMergeFiles();
  } catch (error) { showAlert(error.message); }
});

function renderMergeFiles() {
  byId("merge-file-list").innerHTML = mergeState.files.map((file, index) => `<div class="merge-file-row" data-id="${file.file_id}"><span class="merge-file-name">${file.encrypted ? "🔒 " : ""}${escapeHtml(file.filename)}</span><span class="muted">${formatFileSize(file.size)}${file.encrypted ? " · 已加密" : ""}</span><span class="merge-file-actions"><button type="button" class="secondary small merge-move" data-dir="-1" ${index === 0 ? "disabled" : ""}>上移</button><button type="button" class="secondary small merge-move" data-dir="1" ${index === mergeState.files.length - 1 ? "disabled" : ""}>下移</button><button type="button" class="secondary small merge-remove">移除</button></span></div>`).join("");
  byId("merge-file-count").textContent = mergeState.files.length ? `已添加 ${mergeState.files.length} 个文件` : "尚未添加文件";
  byId("merge-file-picker").textContent = mergeState.files.length ? "继续添加文件" : "添加 .xlsx 文件";
  byId("merge-upload-zone").classList.toggle("has-file", mergeState.files.length > 0);
  byId("merge-password-section").hidden = !mergeState.files.some((file) => file.encrypted);
  byId("file-status").textContent = mergeState.files.length ? `合并模式：已添加 ${mergeState.files.length} 个文件` : "合并模式：未添加文件";
  mergeRefreshNav();
}

byId("merge-file-list").addEventListener("click", async (event) => {
  const button = event.target.closest("button");
  const row = button?.closest(".merge-file-row");
  const fileId = row?.dataset.id;
  if (!fileId) return;
  try {
    if (button.classList.contains("merge-remove")) {
      const payload = await mergeApi("/api/merge/files/remove", { job_id: mergeState.jobId, file_id: fileId });
      mergeState.files = payload.files;
    } else if (button.classList.contains("merge-move")) {
      const index = mergeState.files.findIndex((file) => file.file_id === fileId);
      const target = index + Number(button.dataset.dir);
      if (index < 0 || target < 0 || target >= mergeState.files.length) return;
      const ids = mergeState.files.map((file) => file.file_id);
      [ids[index], ids[target]] = [ids[target], ids[index]];
      const payload = await mergeApi("/api/merge/files/reorder", { job_id: mergeState.jobId, file_ids: ids });
      mergeState.files = payload.files;
    } else {
      return;
    }
    renderMergeFiles();
  } catch (error) { showAlert(error.message); }
});

byId("merge-unlock").addEventListener("click", async () => {
  const password = byId("merge-password").value;
  if (!password) return showAlert("请输入统一密码");
  byId("merge-password").value = "";
  await mergeUnlock({ password });
});

byId("merge-password").addEventListener("keydown", (event) => {
  if (event.key === "Enter") byId("merge-unlock").click();
});

async function mergeUnlock(body) {
  try {
    const payload = await mergeApi("/api/merge/unlock", { job_id: mergeState.jobId, ...body });
    payload.results.forEach((result) => {
      if (!result.success) return;
      const file = mergeState.files.find((item) => item.file_id === result.file_id);
      if (file) file.encrypted = false;
    });
    mergeState.failedUnlocks = payload.results.filter((result) => !result.success);
    renderMergeFiles();
    renderUnlockFailures();
  } catch (error) { showAlert(error.message); }
}

function renderUnlockFailures() {
  byId("merge-unlock-failures").innerHTML = mergeState.failedUnlocks.map((result) => `<div class="merge-failure-row" data-id="${result.file_id}"><span class="merge-file-name">🔒 ${escapeHtml(result.filename)}</span><span class="muted">${escapeHtml(result.error || "解密失败")}，可单独输入密码重试，或移除该文件</span><span class="input-action"><input type="password" placeholder="该文件的密码"><button type="button" class="secondary small merge-retry">重试</button></span></div>`).join("");
}

byId("merge-unlock-failures").addEventListener("click", async (event) => {
  const button = event.target.closest(".merge-retry");
  if (!button) return;
  const row = button.closest(".merge-failure-row");
  const password = row.querySelector("input").value;
  if (!password) return showAlert("请输入该文件的密码");
  await mergeUnlock({ file_passwords: { [row.dataset.id]: password } });
});

byId("merge-to-sheets").addEventListener("click", async () => {
  try {
    const payload = await mergeApi("/api/merge/sheets", { job_id: mergeState.jobId });
    mergeState.sheets = payload.sheets;
    byId("merge-sheet-list").innerHTML = payload.sheets.map((name) => `<label class="check-item"><input type="checkbox" value="${escapeHtml(name)}" checked><span>${escapeHtml(name)}</span></label>`).join("");
    byId("merge-select-all-sheets").checked = true;
    mergeState.maxStep = Math.max(mergeState.maxStep, 2);
    mergeGoStep(2);
  } catch (error) { showAlert(error.message); }
});

byId("merge-select-all-sheets").addEventListener("change", (event) => {
  byId("merge-sheet-list").querySelectorAll("input").forEach((input) => { input.checked = event.target.checked; });
});

function mergeSelectedSheets() {
  return [...byId("merge-sheet-list").querySelectorAll("input:checked")].map((input) => input.value);
}

byId("merge-to-fields").addEventListener("click", () => {
  const selected = mergeSelectedSheets();
  if (!selected.length) return showAlert("至少选择一个 Sheet");
  mergeState.selectedSheets = selected;
  // 只渲染卡片骨架，预览懒加载（展开时才请求），避免多 sheet 时一次性渲染卡顿
  byId("merge-sheet-configs").innerHTML = selected.map((name) => `<article class="merge-sheet-card" data-sheet="${escapeHtml(name)}" data-header-row="1"><header class="merge-card-header"><strong>${escapeHtml(name)}</strong><span class="muted merge-card-source">预览取自第一个包含该 Sheet 的文件</span><label class="merge-card-header-row">表头行<select class="merge-header-row-select">${mergeHeaderRowOptions(15, 1)}</select></label><button type="button" class="secondary small merge-card-toggle">展开预览</button></header><label class="merge-card-identical"><input type="checkbox" class="merge-identical-checkbox"><span>各文件中该 Sheet 内容完全一致，只保留一份（不合并）</span></label><div class="muted merge-identical-hint" hidden>已标记为内容完全一致，合并时只保留第一个文件中的该 Sheet</div><div class="merge-card-preview" hidden></div></article>`).join("");
  byId("merge-plan-results").innerHTML = "";
  mergeState.maxStep = Math.max(mergeState.maxStep, 3);
  mergeGoStep(3);
});

function mergeHeaderRowOptions(totalRows, selectedRow) {
  return Array.from({ length: totalRows }, (_, index) => index + 1)
    .map((row) => `<option value="${row}" ${row === selectedRow ? "selected" : ""}>第 ${row} 行</option>`)
    .join("");
}

function mergeSheetConfigs() {
  return [...document.querySelectorAll(".merge-sheet-card")].map((card) => ({
    sheet_name: card.dataset.sheet,
    header_row: Number(card.dataset.headerRow) || 1,
    identical: card.querySelector(".merge-identical-checkbox").checked,
  }));
}

byId("merge-sheet-configs").addEventListener("click", async (event) => {
  const toggle = event.target.closest(".merge-card-toggle");
  if (!toggle) return;
  const card = toggle.closest(".merge-sheet-card");
  const previewBox = card.querySelector(".merge-card-preview");
  if (!previewBox.hidden) {
    previewBox.hidden = true;
    toggle.textContent = "展开预览";
    return;
  }
  if (!previewBox.dataset.loaded) {
    try {
      const payload = await mergeApi("/api/merge/preview", { job_id: mergeState.jobId, sheet_name: card.dataset.sheet, max_rows: 50 });
      previewBox.innerHTML = renderMergePreviewTable(payload);
      previewBox.dataset.loaded = "1";
      card.querySelector(".merge-card-source").textContent = `预览来源：${payload.source_file}`;
      // 与拆分一致：表头行选项按实际行数生成（最多 15 行）
      const headerRow = Number(card.dataset.headerRow) || 1;
      card.querySelector(".merge-header-row-select").innerHTML = mergeHeaderRowOptions(Math.min(15, payload.total_rows), headerRow);
      markMergeHeaderRow(card);
    } catch (error) {
      showAlert(error.message);
      return;
    }
  }
  previewBox.hidden = false;
  toggle.textContent = "收起预览";
});

byId("merge-sheet-configs").addEventListener("change", (event) => {
  const identicalBox = event.target.closest(".merge-identical-checkbox");
  if (identicalBox) {
    const card = identicalBox.closest(".merge-sheet-card");
    card.querySelector(".merge-card-header-row").hidden = identicalBox.checked;
    card.querySelector(".merge-card-toggle").hidden = identicalBox.checked;
    card.querySelector(".merge-identical-hint").hidden = !identicalBox.checked;
    if (identicalBox.checked) {
      card.querySelector(".merge-card-preview").hidden = true;
      card.querySelector(".merge-card-toggle").textContent = "展开预览";
    }
    return;
  }
  const select = event.target.closest(".merge-header-row-select");
  if (!select) return;
  const card = select.closest(".merge-sheet-card");
  card.dataset.headerRow = select.value;
  markMergeHeaderRow(card);
});

function renderMergePreviewTable(preview) {
  const body = preview.rows.map((row, index) => `<tr data-row="${index + 1}"><th class="row-number">${index + 1}</th>${row.map((value) => `<td>${escapeHtml(value)}</td>`).join("")}</tr>`).join("");
  return `<div class="preview-meta">共 ${preview.total_rows} 行，表头行之上的标题内容会保留到合并结果</div><div class="preview-wrap"><table class="preview-table"><tbody>${body}</tbody></table></div>`;
}

function markMergeHeaderRow(card) {
  const headerRow = Number(card.dataset.headerRow) || 1;
  card.querySelectorAll("tr[data-row]").forEach((tr) => {
    const isHeader = Number(tr.dataset.row) === headerRow;
    tr.classList.toggle("merge-header-selected", isHeader);
    const numberCell = tr.querySelector(".row-number");
    if (numberCell) numberCell.textContent = isHeader ? `${tr.dataset.row} · 表头` : tr.dataset.row;
  });
}

byId("merge-check-plan").addEventListener("click", async () => {
  try {
    const payload = await mergeApi("/api/merge/plan", { job_id: mergeState.jobId, sheet_configs: mergeSheetConfigs() });
    renderMergePlan(payload);
  } catch (error) { showAlert(error.message); }
});

function renderMergePlan(plan) {
  const sheets = plan.sheets.map((sheet) => {
    if (sheet.identical) {
      return `<article class="merge-plan-sheet"><h2>${escapeHtml(sheet.sheet_name)}</h2><div class="muted">已标记为内容完全一致，合并时只保留第一个文件中的该 Sheet，不做字段检查</div></article>`;
    }
    const issues = [];
    Object.entries(sheet.missing_fields).forEach(([file, fields]) => issues.push(`文件 ${escapeHtml(file)} 缺少字段：${escapeHtml(fields.join("、"))}，对应列将留空`));
    Object.entries(sheet.extra_fields).forEach(([file, fields]) => issues.push(`文件 ${escapeHtml(file)} 多出字段：${escapeHtml(fields.join("、"))}，将追加到表头末尾`));
    sheet.missing_files.forEach((file) => issues.push(`文件 ${escapeHtml(file)} 没有该 Sheet，将跳过`));
    const title = `<h2>${escapeHtml(sheet.sheet_name)}（并集 ${sheet.union_headers.length} 个字段）</h2>`;
    if (!issues.length) {
      return `<article class="merge-plan-sheet">${title}<div class="muted">所有文件字段一致</div></article>`;
    }
    return `<article class="merge-plan-sheet">${title}${renderMergeFieldMatrix(sheet)}<ul>${issues.map((line) => `<li>${line}</li>`).join("")}</ul></article>`;
  }).join("");
  const warnings = (plan.warnings || []).length
    ? `<div class="warning-list">${plan.warnings.map((message) => `<div>${escapeHtml(message)}</div>`).join("")}</div>`
    : "";
  byId("merge-plan-results").innerHTML = sheets + warnings;
}

function renderMergeFieldMatrix(sheet) {
  const files = mergeState.files.map((file) => file.filename);
  const header = `<tr><th>文件</th>${sheet.union_headers.map((field) => `<th>${escapeHtml(field)}</th>`).join("")}</tr>`;
  const body = files.map((name) => {
    if (sheet.missing_files.includes(name)) {
      return `<tr><td class="merge-matrix-file">${escapeHtml(name)}</td><td class="merge-field-missing" colspan="${sheet.union_headers.length}">没有该 Sheet</td></tr>`;
    }
    const missing = sheet.missing_fields[name] || [];
    const cells = sheet.union_headers.map((field) => (
      missing.includes(field)
        ? `<td class="merge-field-missing">—</td>`
        : `<td class="merge-field-has">✓</td>`
    )).join("");
    return `<tr><td class="merge-matrix-file">${escapeHtml(name)}</td>${cells}</tr>`;
  }).join("");
  return `<div class="merge-matrix-wrap"><table class="merge-field-matrix">${header}${body}</table></div>`;
}

function mergeParentDir(path) {
  const sep = path.includes("/") ? "/" : "\\";
  return path.substring(0, path.lastIndexOf(sep));
}

byId("merge-to-output").addEventListener("click", () => {
  const first = mergeState.files[0];
  if (first) {
    if (!byId("merge-output-dir").value && first.source_path) {
      byId("merge-output-dir").value = mergeParentDir(first.source_path);
    }
    if (!byId("merge-output-filename").value) {
      byId("merge-output-filename").value = `${first.filename.replace(/\.xlsx$/i, "")}_合并结果.xlsx`;
    }
  }
  mergeState.maxStep = Math.max(mergeState.maxStep, 4);
  mergeGoStep(4);
});

byId("merge-browse-output-dir").addEventListener("click", async () => {
  setBusy(true, "请选择输出目录");
  try {
    const payload = await api("/api/select-output-dir", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ current_path: byId("merge-output-dir").value }),
    }, false);
    if (payload.selected) byId("merge-output-dir").value = payload.path;
  } catch (error) {
    showAlert(error.message);
  } finally {
    setBusy(false);
  }
});

byId("merge-output-encrypt").addEventListener("change", (event) => {
  byId("merge-output-password-group").hidden = !event.target.checked;
});

byId("merge-execute").addEventListener("click", async () => {
  const encryptChecked = byId("merge-output-encrypt").checked;
  const outputPassword = encryptChecked ? byId("merge-output-password").value : "";
  if (encryptChecked && !outputPassword) return showAlert("请输入加密密码");
  const body = {
    job_id: mergeState.jobId,
    sheet_configs: mergeSheetConfigs(),
    include_source_column: byId("merge-include-source").checked,
    skip_duplicate_sheets: byId("merge-skip-duplicates").checked,
    overwrite: byId("merge-overwrite").checked,
    output_password: outputPassword,
    background: true,
  };
  if (byId("merge-output-dir").value.trim()) body.output_dir = byId("merge-output-dir").value.trim();
  if (byId("merge-output-filename").value.trim()) body.output_filename = byId("merge-output-filename").value.trim();
  try {
    mergeState.maxStep = 5;
    mergeState.executing = true;
    mergeGoStep(5);
    mergePrepareProgress();
    setBusy(true, "合并任务执行中");
    const started = await api("/api/merge/execute", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) }, false);
    await mergePollExecution(started.progress_url);
  } catch (error) {
    mergeUpdateProgress(0, "合并失败");
    showAlert(error.message);
  } finally {
    mergeState.executing = false;
    setBusy(false);
    mergeRefreshNav();
  }
});

function mergePrepareProgress() {
  byId("merge-progress-panel").hidden = false;
  byId("merge-result-content").hidden = true;
  mergeUpdateProgress(0, "任务已提交，正在准备");
}

function mergeUpdateProgress(percent, message) {
  const normalized = Math.max(0, Math.min(100, Number(percent) || 0));
  byId("merge-execution-progress").value = normalized;
  byId("merge-execution-progress").textContent = `${normalized}%`;
  byId("merge-progress-percent").textContent = `${normalized}%`;
  byId("merge-progress-message").textContent = message;
}

async function mergePollExecution(progressUrl) {
  while (true) {
    const snapshot = await api(progressUrl, {}, false);
    mergeUpdateProgress(snapshot.progress, snapshot.message);
    if (snapshot.status === "complete") {
      renderMergeResults(snapshot.result);
      return;
    }
    if (snapshot.status === "failed") throw new Error(snapshot.error || "合并失败");
    await new Promise((resolve) => setTimeout(resolve, 500));
  }
}

function renderMergeResults(payload) {
  byId("merge-result-content").hidden = false;
  mergeUpdateProgress(100, "合并完成");
  byId("merge-total-rows").textContent = payload.total_rows;
  byId("merge-sheet-count").textContent = payload.results.length;
  byId("merge-warning-count").textContent = payload.warnings.length;
  byId("merge-error-count").textContent = payload.errors.length;
  const sheetStats = payload.results.map((result) => {
    const sources = Object.entries(result.source_rows).map(([file, rows]) => `${escapeHtml(file)} ${rows} 行`).join("，");
    const skipped = Object.entries(result.skipped_duplicates || {}).map(([file, original]) => `文件 ${escapeHtml(file)}（与 ${escapeHtml(original)} 完全相同，已跳过）`).join("，");
    return `<div class="result-stats"><span>${escapeHtml(result.sheet_name)}：合并 ${result.merged_rows} 行（${sources}）${skipped ? `；跳过重复：${skipped}` : ""}</span></div>`;
  }).join("");
  byId("merge-result-list").innerHTML = `${sheetStats}<div class="result-item"><strong>合并结果</strong><span class="result-path">${escapeHtml(payload.output_file)}</span><span class="merge-result-actions"><button class="btn-action" data-action="open-file" data-path="${escapeHtml(payload.output_file)}">打开文件</button><button class="btn-action" data-action="open-folder" data-path="${escapeHtml(payload.output_file)}">打开所在文件夹</button></span></div>`;
  const messages = [...payload.warnings, ...payload.errors];
  byId("merge-warning-list").innerHTML = messages.map((message) => `<div>${escapeHtml(message)}</div>`).join("");
  mergeRefreshNav();
}

byId("merge-result-list").addEventListener("click", async (event) => {
  const button = event.target.closest(".btn-action");
  if (!button) return;
  try {
    await api(`/api/action/${button.dataset.action}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path: button.dataset.path }),
    }, false);
  } catch (error) {
    showAlert(error.message);
  }
});

document.querySelectorAll(".mback").forEach((button) => button.addEventListener("click", () => mergeGoStep(Number(button.dataset.mergeTarget))));
document.querySelectorAll(".mstep").forEach((button) => button.addEventListener("click", () => mergeGoStep(Number(button.dataset.mergeStep))));
byId("merge-restart").addEventListener("click", () => window.location.reload());
