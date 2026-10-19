let tmResult = null;
let tmDbReady = document.getElementById("tm-db-status").dataset.ready === "true";
let tmBusy = false;

function tmNode(tag, value) {
    const node = document.createElement(tag);
    if (value !== undefined) node.textContent = value;
    return node;
}

async function tmRequest(url, options) {
    const response = await fetch(url, options);
    const data = await response.json();
    if (!response.ok || data.error) throw new Error(data.error || "请求失败");
    return data;
}

function syncTmControls() {
    for (const id of ["choose-tm-btn", "tm-input"]) document.getElementById(id).disabled = tmBusy;
    document.getElementById("update-tm-btn").disabled = tmBusy || !document.getElementById("tm-input").files[0];
    for (const id of ["start-btn", "choose-file-btn", "file-input"]) document.getElementById(id).disabled = tmBusy || !tmDbReady;
}

function resetTmPreview() {
    document.getElementById("progress-container").style.display = "none";
    document.getElementById("tm-preview").hidden = true;
    document.getElementById("download-btn").style.display = "none";
    tmResult = null;
}

async function updateTmDatabase() {
    const input = document.getElementById("tm-input");
    const feedback = document.getElementById("tm-db-feedback");
    if (!input.files[0]) return;
    tmBusy = true;
    syncTmControls();
    feedback.style.color = "";
    feedback.textContent = "正在校验并保存翻译记忆库...";
    try {
        const form = new FormData();
        form.append("tm_tmx_file", input.files[0]);
        const data = await tmRequest("/update/tm_db", {method: "POST", body: form});
        const db = data.tm_db;
        const status = document.getElementById("tm-db-status");
        status.replaceChildren(
            tmNode("strong", "当前翻译记忆库：" + db.filename),
            tmNode("small", db.valid_count + " 条有效英德记录，" + db.source_count + " 组英文原文，同原文异译 " + db.conflict_count + " 组；跳过 " + db.skipped_count + " 条。更新时间：" + db.updated_at)
        );
        status.dataset.ready = "true";
        tmDbReady = true;
        input.value = "";
        document.getElementById("tm-file-name").textContent = "点击选择英德翻译记忆库（.tmx）";
        document.getElementById("update-tm-btn").textContent = "更新当前翻译记忆库";
        feedback.textContent = data.message + "，可以上传上一阶段的文件开始匹配。";
        resetTmPreview();
    } catch (error) {
        feedback.textContent = error.message;
        feedback.style.color = "var(--error-color)";
    } finally {
        tmBusy = false;
        syncTmControls();
    }
}

function renderTmRows() {
    if (!tmResult) return;
    const status = document.getElementById("status-filter").value;
    const fragment = document.createDocumentFragment();
    for (const record of tmResult.records) {
        if (status && record.status !== status) continue;
        const row = tmNode("tr");
        for (const value of [record.row, record.key || record.speaker || "无 Key", record.english, record.status]) row.append(tmNode("td", value));
        const cell = tmNode("td");
        if (record.warning) cell.append(tmNode("p", record.warning));
        if (!record.matches.length) cell.append(tmNode("span", record.warning ? "未扫描" : "未找到合适的历史候选"));
        for (const match of record.matches) {
            const details = tmNode("details");
            details.append(tmNode("summary", match.match_type + " · 文字相似度 " + Math.round(match.score * 100) + "%"));
            details.append(tmNode("pre", "【英】" + match.english + "\n【德】" + match.german));
            if (match.warnings.length) details.append(tmNode("p", "核对提示：" + match.warnings.join("；")));
            for (const change of match.changes) details.append(tmNode("p", "差异：当前〔" + (change.current || "无") + "〕／历史〔" + (change.history || "无") + "〕"));
            const sources = tmNode("details");
            sources.append(tmNode("summary", "查看 " + match.origins.length + " 处来源与历史上下文"));
            for (const origin of match.origins) {
                sources.append(tmNode("p", tmResult.tm_db.filename + " / " + (origin.document || "未提供原文档名") + " / TU " + origin.tu + (origin.date ? " / " + origin.date : "")));
                if (origin.previous) sources.append(tmNode("pre", "前文：" + origin.previous));
                if (origin.next) sources.append(tmNode("pre", "后文：" + origin.next));
            }
            details.append(sources);
            cell.append(details);
        }
        row.append(cell);
        fragment.append(row);
    }
    document.getElementById("tm-rows").replaceChildren(fragment);
}

async function startTm() {
    const fileInput = document.getElementById("file-input");
    const status = document.getElementById("status-text");
    const download = document.getElementById("download-btn");
    document.getElementById("progress-container").style.display = "block";
    document.getElementById("tm-preview").hidden = true;
    download.style.display = "none";
    status.style.color = "";
    document.getElementById("progress-bar").style.width = "0%";
    if (!fileInput.files[0]) { status.textContent = "请先选择第四步下载的术语筛选文件"; return; }
    tmBusy = true;
    syncTmControls();
    try {
        const form = new FormData();
        form.append("trans_file", fileInput.files[0]);
        status.textContent = "正在读取英文原文并匹配英德翻译记忆...";
        const {task_id: taskId} = await tmRequest("/upload/tm", {method: "POST", body: form});
        while (true) {
            const task = await tmRequest("/progress/" + taskId);
            status.textContent = task.status;
            document.getElementById("progress-bar").style.width = task.progress + "%";
            if (task.state === "error") throw new Error(task.status);
            if (task.state === "completed") break;
            await new Promise(resolve => setTimeout(resolve, 500));
        }
        tmResult = await tmRequest("/api/tm/" + taskId);
        document.getElementById("tm-summary").textContent = "共 " + tmResult.records.length + " 条；" + Object.entries(tmResult.counts).map(([key, value]) => key + " " + value + " 条").join("，") + "；使用记忆库：" + tmResult.tm_db.filename;
        document.getElementById("status-filter").value = "";
        renderTmRows();
        document.getElementById("tm-preview").hidden = false;
        download.onclick = () => { window.location.href = "/download/" + taskId; };
        download.style.display = "block";
    } catch (error) {
        status.textContent = error.message;
        status.style.color = "var(--error-color)";
    } finally {
        tmBusy = false;
        syncTmControls();
    }
}

document.addEventListener("DOMContentLoaded", () => {
    const fileInput = document.getElementById("file-input");
    document.getElementById("choose-file-btn").onclick = () => fileInput.click();
    document.getElementById("start-btn").onclick = startTm;
    document.getElementById("status-filter").onchange = renderTmRows;
    const tbInput = document.getElementById("tm-input");
    document.getElementById("choose-tm-btn").onclick = () => tbInput.click();
    document.getElementById("update-tm-btn").onclick = updateTmDatabase;
    tbInput.onchange = () => {
        document.getElementById("tm-file-name").textContent = tbInput.files[0] ? tbInput.files[0].name : "点击选择英德翻译记忆库（.tmx）";
        document.getElementById("tm-db-feedback").textContent = "";
        syncTmControls();
    };
    fileInput.onchange = () => {
        document.getElementById("file-name").textContent = fileInput.files[0] ? fileInput.files[0].name : "点击选择第四步下载的术语筛选文件（.xlsx）";
        resetTmPreview();
    };
    syncTmControls();
});
