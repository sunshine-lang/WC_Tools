let termResult = null;
let termDbReady = document.getElementById("term-db-status").dataset.ready === "true";
let termBusy = false;

function termNode(tag, value) {
    const node = document.createElement(tag);
    if (value !== undefined) node.textContent = value;
    return node;
}

async function termRequest(url, options) {
    const response = await fetch(url, options);
    const data = await response.json();
    if (!response.ok || data.error) throw new Error(data.error || "请求失败");
    return data;
}

function syncTermControls() {
    for (const id of ["choose-tb-btn", "tb-input"]) document.getElementById(id).disabled = termBusy;
    document.getElementById("update-tb-btn").disabled = termBusy || !document.getElementById("tb-input").files[0];
    for (const id of ["start-btn", "choose-file-btn", "file-input"]) document.getElementById(id).disabled = termBusy || !termDbReady;
}

function resetTermPreview() {
    document.getElementById("progress-container").style.display = "none";
    document.getElementById("term-preview").hidden = true;
    document.getElementById("download-btn").style.display = "none";
    termResult = null;
}

async function updateTermDatabase() {
    const input = document.getElementById("tb-input");
    const feedback = document.getElementById("term-db-feedback");
    if (!input.files[0]) return;
    termBusy = true;
    syncTermControls();
    feedback.style.color = "";
    feedback.textContent = "正在校验并保存术语库...";
    try {
        const form = new FormData();
        form.append("tb_file", input.files[0]);
        const data = await termRequest("/update/term_db", {method: "POST", body: form});
        const db = data.term_db;
        const status = document.getElementById("term-db-status");
        status.replaceChildren(
            termNode("strong", "当前术语库：" + db.filename),
            termNode("small", db.row_count + " 条有效记录，" + db.term_count + " 个英文词条；同词异译 " + db.conflict_count + " 项。更新时间：" + db.updated_at)
        );
        status.dataset.ready = "true";
        termDbReady = true;
        input.value = "";
        document.getElementById("tb-file-name").textContent = "点击选择英德术语库（.xlsx / .csv）";
        document.getElementById("update-tb-btn").textContent = "更新当前术语库";
        feedback.textContent = data.message + "，可以上传上一阶段的文件开始匹配。";
        resetTermPreview();
    } catch (error) {
        feedback.textContent = error.message;
        feedback.style.color = "var(--error-color)";
    } finally {
        termBusy = false;
        syncTermControls();
    }
}

function renderTermRows() {
    if (!termResult) return;
    const status = document.getElementById("status-filter").value;
    const fragment = document.createDocumentFragment();
    for (const record of termResult.records) {
        if (status && record.status !== status) continue;
        const row = termNode("tr");
        for (const value of [record.row, record.key || record.speaker || "无 Key", record.english, record.status]) row.append(termNode("td", value));
        const cell = termNode("td", record.terms || (record.warning ? "未扫描" : "未找到术语候选"));
        if (record.warning) cell.append(termNode("p", record.warning));
        if (record.matches.length) {
            const details = termNode("details");
            details.append(termNode("summary", "查看命中依据与来源"));
            const list = termNode("ul");
            for (const match of record.matches) {
                list.append(termNode("li", match.matched_text + " → " + match.english + "（" + match.match_type + "匹配）；" + match.sources.join("；")));
            }
            details.append(list);
            cell.append(details);
        }
        row.append(cell);
        fragment.append(row);
    }
    document.getElementById("term-rows").replaceChildren(fragment);
}

async function startTerm() {
    const fileInput = document.getElementById("file-input");
    const status = document.getElementById("status-text");
    const download = document.getElementById("download-btn");
    document.getElementById("progress-container").style.display = "block";
    document.getElementById("term-preview").hidden = true;
    download.style.display = "none";
    status.style.color = "";
    document.getElementById("progress-bar").style.width = "0%";
    if (!fileInput.files[0]) { status.textContent = "请先选择第二步下载的上下文文件"; return; }
    termBusy = true;
    syncTermControls();
    try {
        const form = new FormData();
        form.append("trans_file", fileInput.files[0]);
        status.textContent = "正在读取英文原文并匹配英德术语...";
        const {task_id: taskId} = await termRequest("/upload/term", {method: "POST", body: form});
        while (true) {
            const task = await termRequest("/progress/" + taskId);
            status.textContent = task.status;
            document.getElementById("progress-bar").style.width = task.progress + "%";
            if (task.state === "error") throw new Error(task.status);
            if (task.state === "completed") break;
            await new Promise(resolve => setTimeout(resolve, 500));
        }
        termResult = await termRequest("/api/term/" + taskId);
        document.getElementById("term-summary").textContent = "共 " + termResult.records.length + " 条；" + Object.entries(termResult.counts).map(([key, value]) => key + " " + value + " 条").join("，") + "；使用词库：" + termResult.term_db.filename;
        document.getElementById("status-filter").value = "";
        renderTermRows();
        document.getElementById("term-preview").hidden = false;
        download.onclick = () => { window.location.href = "/download/" + taskId; };
        download.style.display = "block";
    } catch (error) {
        status.textContent = error.message;
        status.style.color = "var(--error-color)";
    } finally {
        termBusy = false;
        syncTermControls();
    }
}

document.addEventListener("DOMContentLoaded", () => {
    const fileInput = document.getElementById("file-input");
    document.getElementById("choose-file-btn").onclick = () => fileInput.click();
    document.getElementById("start-btn").onclick = startTerm;
    document.getElementById("status-filter").onchange = renderTermRows;
    const tbInput = document.getElementById("tb-input");
    document.getElementById("choose-tb-btn").onclick = () => tbInput.click();
    document.getElementById("update-tb-btn").onclick = updateTermDatabase;
    tbInput.onchange = () => {
        document.getElementById("tb-file-name").textContent = tbInput.files[0] ? tbInput.files[0].name : "点击选择英德术语库（.xlsx / .csv）";
        document.getElementById("term-db-feedback").textContent = "";
        syncTermControls();
    };
    fileInput.onchange = () => {
        document.getElementById("file-name").textContent = fileInput.files[0] ? fileInput.files[0].name : "点击选择第二步下载的上下文文件（.xlsx）";
        resetTermPreview();
    };
    syncTermControls();
});
