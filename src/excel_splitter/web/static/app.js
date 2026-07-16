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
