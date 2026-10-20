const cleanInput = document.getElementById("file-input");
const cleanColumn = document.getElementById("clean-column");
let cleanBusy = false;

function cleanNode(tag, value) {
    const node = document.createElement(tag);
    node.textContent = value;
    return node;
}

async function cleanRequest(url, options) {
    const response = await fetch(url, options);
    const data = await response.json();
    if (!response.ok || data.error) throw new Error(data.error || "请求失败，请重试");
    return data;
}

function syncCleanControls() {
    for (const node of [cleanInput, cleanColumn, document.getElementById("choose-file-btn")]) node.disabled = cleanBusy;
    document.getElementById("start-btn").disabled = cleanBusy || !cleanInput.files[0];
}

function resetCleanPreview() {
    document.getElementById("progress-container").style.display = "none";
    document.getElementById("clean-preview").hidden = true;
    document.getElementById("download-btn").style.display = "none";
    syncCleanControls();
}

async function startClean() {
    if (!cleanInput.files[0]) return;
    const status = document.getElementById("status-text");
    const download = document.getElementById("download-btn");
    cleanBusy = true;
    syncCleanControls();
    document.getElementById("progress-container").style.display = "block";
    document.getElementById("clean-preview").hidden = true;
    download.style.display = "none";
    status.style.color = "";
    status.textContent = "正在上传并清洗...";
    document.getElementById("progress-bar").style.width = "0%";
    try {
        const form = new FormData();
        form.append("file", cleanInput.files[0]);
        form.append("column", cleanColumn.value);
        const {task_id: taskId} = await cleanRequest("/upload/tag_clean", {method: "POST", body: form});
        while (true) {
            const task = await cleanRequest("/progress/" + taskId);
            status.textContent = task.status;
            document.getElementById("progress-bar").style.width = task.progress + "%";
            if (task.state === "error") throw new Error(task.status);
            if (task.state === "completed") break;
            await new Promise(resolve => setTimeout(resolve, 500));
        }
        const result = await cleanRequest("/api/tag_clean/" + taskId);
        document.getElementById("clean-summary").textContent = "清洗列：" + result.column + "；检查 " + result.total + " 个文本单元格，修改 " + result.changed + " 个，跳过 " + result.skipped_formulas + " 个公式单元格。";
        const rows = result.records.map(record => {
            const row = document.createElement("tr");
            row.append(cleanNode("td", record.sheet + " 第 " + record.row + " 行\n" + record.key),
                cleanNode("td", record.before), cleanNode("td", record.after));
            return row;
        });
        if (!rows.length) {
            const row = document.createElement("tr"), cell = cleanNode("td", "未发现需要清洗的标签空格，文件内容保持原样。");
            cell.colSpan = 3;
            row.append(cell);
            rows.push(row);
        }
        document.getElementById("clean-rows").replaceChildren(...rows);
        document.getElementById("clean-preview").hidden = false;
        download.onclick = () => { window.location.href = "/download/" + taskId; };
        download.style.display = "block";
    } catch (error) {
        status.textContent = error.message;
        status.style.color = "var(--error-color)";
    } finally {
        cleanBusy = false;
        syncCleanControls();
    }
}

document.getElementById("choose-file-btn").onclick = () => cleanInput.click();
document.getElementById("start-btn").onclick = startClean;
cleanInput.onchange = () => {
    document.getElementById("file-name").textContent = cleanInput.files[0] ? cleanInput.files[0].name : "点击选择文件（.xlsx / UTF-8 .csv）";
    resetCleanPreview();
};
cleanColumn.onchange = resetCleanPreview;
syncCleanControls();
