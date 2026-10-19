let qaResult = null;
let qaBusy = false;

const QA_FILES = {translated: "translated", original: "original"};
const QA_PLACEHOLDERS = {
    translated: "点击选择翻译回填后的文件（.xlsx / .csv，必填）",
    original: "点击选择原始待译文件（选填，用于行对应检查）",
};

function qaNode(tag, value, className) {
    const node = document.createElement(tag);
    if (value !== undefined) node.textContent = value;
    if (className) node.className = className;
    return node;
}

async function qaRequest(url, options) {
    const response = await fetch(url, options);
    const data = await response.json();
    if (!response.ok || data.error) throw new Error(data.error || "请求失败");
    return data;
}

function syncQaControls() {
    const ids = ["start-btn", "choose-translated-btn", "translated-input", "choose-original-btn", "original-input",
        "source-column", "target-column", "key-column", "limit-column", "length-ratio"];
    for (const id of ids) document.getElementById(id).disabled = qaBusy;
}

function resetQaPreview() {
    document.getElementById("progress-container").style.display = "none";
    document.getElementById("qa-preview").hidden = true;
    document.getElementById("download-btn").style.display = "none";
    qaResult = null;
}

function renderQaRows() {
    if (!qaResult) return;
    const level = document.getElementById("level-filter").value;
    const check = document.getElementById("check-filter").value;
    const fragment = document.createDocumentFragment();
    for (const record of qaResult.records) {
        const issues = record.issues.filter(issue => (!level || issue.level === level) && (!check || issue.check === check));
        if (!issues.length) continue;
        const row = qaNode("tr");
        for (const value of [record.row, record.key || "无 Key", record.source, record.target]) row.append(qaNode("td", value));
        const cell = qaNode("td");
        for (const issue of issues) {
            const line = qaNode("p", undefined, "qa-issue");
            line.append(qaNode("span", issue.level, "qa-level qa-level-" + (issue.level === "错误" ? "error" : "warning")));
            line.append(qaNode("span", issue.check + "：" + issue.message));
            cell.append(line);
        }
        row.append(cell);
        fragment.append(row);
    }
    const body = document.getElementById("qa-rows");
    body.replaceChildren(fragment);
    if (!body.childElementCount) {
        const row = qaNode("tr");
        const cell = qaNode("td", "当前筛选条件下没有问题记录");
        cell.colSpan = 5;
        row.append(cell);
        body.append(row);
    }
}

async function startQa() {
    const translated = document.getElementById("translated-input");
    const status = document.getElementById("status-text");
    const download = document.getElementById("download-btn");
    document.getElementById("progress-container").style.display = "block";
    document.getElementById("qa-preview").hidden = true;
    download.style.display = "none";
    status.style.color = "";
    document.getElementById("progress-bar").style.width = "0%";
    if (!translated.files[0]) { status.textContent = "请先选择翻译回填后的文件"; return; }
    qaBusy = true;
    syncQaControls();
    try {
        const form = new FormData();
        form.append("translated_file", translated.files[0]);
        const original = document.getElementById("original-input");
        if (original.files[0]) form.append("original_file", original.files[0]);
        form.append("source_column", document.getElementById("source-column").value);
        form.append("target_column", document.getElementById("target-column").value);
        form.append("key_column", document.getElementById("key-column").value);
        form.append("limit_column", document.getElementById("limit-column").value);
        form.append("length_ratio", document.getElementById("length-ratio").value);
        status.textContent = "正在上传并检查...";
        const {task_id: taskId} = await qaRequest("/upload/qa", {method: "POST", body: form});
        while (true) {
            const task = await qaRequest("/progress/" + taskId);
            status.textContent = task.status;
            document.getElementById("progress-bar").style.width = task.progress + "%";
            if (task.state === "error") throw new Error(task.status);
            if (task.state === "completed") break;
            await new Promise(resolve => setTimeout(resolve, 500));
        }
        qaResult = await qaRequest("/api/qa/" + taskId);
        document.getElementById("qa-summary").textContent =
            "共 " + qaResult.total + " 条，已有译文 " + qaResult.translated + " 条；" +
            Object.entries(qaResult.counts).map(([key, value]) => key + " " + value + " 条").join("，") +
            "；问题记录 " + qaResult.records.length + " 条。";
        const list = document.getElementById("qa-file-issues");
        list.replaceChildren(...qaResult.file_issues.map(issue => qaNode("li", issue.check + "：" + issue.message)));
        const checkFilter = document.getElementById("check-filter");
        const all = qaNode("option", "全部检查项");
        all.value = "";
        checkFilter.replaceChildren(all, ...Object.entries(qaResult.check_counts).map(([name, count]) => {
            const option = qaNode("option", name + "（" + count + "）");
            option.value = name;
            return option;
        }));
        document.getElementById("level-filter").value = "";
        renderQaRows();
        document.getElementById("qa-preview").hidden = false;
        download.onclick = () => { window.location.href = "/download/" + taskId; };
        download.style.display = "block";
    } catch (error) {
        status.textContent = error.message;
        status.style.color = "var(--error-color)";
    } finally {
        qaBusy = false;
        syncQaControls();
    }
}

document.addEventListener("DOMContentLoaded", () => {
    for (const role of Object.values(QA_FILES)) {
        const input = document.getElementById(role + "-input");
        document.getElementById("choose-" + role + "-btn").onclick = () => input.click();
        input.onchange = () => {
            document.getElementById(role + "-name").textContent = input.files[0] ? input.files[0].name : QA_PLACEHOLDERS[role];
            resetQaPreview();
        };
    }
    document.getElementById("start-btn").onclick = startQa;
    document.getElementById("level-filter").onchange = renderQaRows;
    document.getElementById("check-filter").onchange = renderQaRows;
    syncQaControls();
});
