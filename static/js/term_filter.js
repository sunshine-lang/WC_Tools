let filterResult = null;
let filterBusy = false;
const filterReady = document.getElementById("filter-model").dataset.ready === "true";
const filterInput = document.getElementById("file-input");

function filterNode(tag, text) {
    const node = document.createElement(tag);
    if (text !== undefined) node.textContent = text;
    return node;
}

async function filterRequest(url, options) {
    const response = await fetch(url, options);
    const data = await response.json();
    if (!response.ok || data.error) throw new Error(data.error || "请求失败，请重试");
    return data;
}

function syncFilterControls() {
    for (const id of ["choose-file-btn", "file-input"]) document.getElementById(id).disabled = filterBusy || !filterReady;
    document.getElementById("start-btn").disabled = filterBusy || !filterReady || !filterInput.files[0];
}

function renderFilterRows() {
    if (!filterResult) return;
    const fragment = document.createDocumentFragment();
    for (const record of filterResult.records) {
        const row = filterNode("tr");
        row.append(filterNode("td", record.sheet + " 第 " + record.row + " 行\n" + (record.key || record.speaker || "无 Key")));
        const source = filterNode("td");
        source.append(filterNode("p", record.english || "英文源文为空"));
        if (record.context) {
            const context = filterNode("details");
            context.append(filterNode("summary", "查看本行 Context"), filterNode("pre", record.context));
            source.append(context);
        }
        row.append(source, filterNode("td", record.active || ""), filterNode("td", record.excluded || ""));
        fragment.append(row);
    }
    if (!fragment.childNodes.length) {
        const row = filterNode("tr"), cell = filterNode("td", "没有记录");
        cell.colSpan = 4;
        row.append(cell);
        fragment.append(row);
    }
    document.getElementById("filter-rows").replaceChildren(fragment);
}

async function startFilter() {
    const status = document.getElementById("status-text");
    const download = document.getElementById("download-btn");
    filterBusy = true;
    filterResult = null;
    syncFilterControls();
    document.getElementById("progress-container").style.display = "block";
    document.getElementById("filter-preview").hidden = true;
    download.style.display = "none";
    status.style.color = "";
    document.getElementById("progress-bar").style.width = "0%";
    status.textContent = "正在检查文件并准备语境筛选...";
    try {
        const form = new FormData();
        form.append("file", filterInput.files[0]);
        const {task_id: taskId} = await filterRequest("/upload/term_filter", {method: "POST", body: form});
        while (true) {
            const task = await filterRequest("/progress/" + taskId);
            status.textContent = task.status;
            document.getElementById("progress-bar").style.width = task.progress + "%";
            if (task.state === "error") throw new Error(task.status);
            if (task.state === "completed") break;
            await new Promise(resolve => setTimeout(resolve, 800));
        }
        filterResult = await filterRequest("/api/term_filter/" + taskId);
        const counts = filterResult.term_counts;
        document.getElementById("filter-summary").textContent = "共 " + filterResult.records.length + " 行；生效术语 " + (counts.keep || 0) + " 个，无效术语 " + (counts.exclude || 0) + " 个。";
        renderFilterRows();
        document.getElementById("filter-preview").hidden = false;
        download.onclick = () => { window.location.href = "/download/" + taskId; };
        download.style.display = "block";
        if (filterResult.error_count) status.style.color = "var(--error-color)";
    } catch (error) {
        status.textContent = error.message;
        status.style.color = "var(--error-color)";
    } finally {
        filterBusy = false;
        syncFilterControls();
    }
}

document.getElementById("choose-file-btn").onclick = () => filterInput.click();
document.getElementById("start-btn").onclick = startFilter;
filterInput.onchange = () => {
    document.getElementById("file-name").textContent = filterInput.files[0] ? filterInput.files[0].name : "点击选择第三步下载的文件（.xlsx）";
    document.getElementById("progress-container").style.display = "none";
    document.getElementById("filter-preview").hidden = true;
    document.getElementById("download-btn").style.display = "none";
    filterResult = null;
    syncFilterControls();
};
syncFilterControls();
