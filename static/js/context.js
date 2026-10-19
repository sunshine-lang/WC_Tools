let contextResult = null;
let contextTaskId = null;

function contextNode(tag, value) {
    const node = document.createElement(tag);
    if (value !== undefined) node.textContent = value;
    return node;
}

async function contextRequest(url, options) {
    const response = await fetch(url, options);
    const data = await response.json();
    if (!response.ok || data.error) throw new Error(data.error || "请求失败");
    return data;
}

function renderContextRows() {
    if (!contextResult) return;
    const category = document.getElementById("category-filter").value;
    const status = document.getElementById("status-filter").value;
    const fragment = document.createDocumentFragment();
    for (const record of contextResult.records) {
        if ((category && record.category !== category) || (status && record.status !== status)) continue;
        const row = contextNode("tr");
        for (const value of [record.row, record.key || record.speaker || "无 Key", record.english, record.status]) row.append(contextNode("td", value));
        const cell = contextNode("td");
        const details = contextNode("details");
        details.append(contextNode("summary", record.context ? "查看上下文（" + record.related_count + " 条关联原文）" : "查看未关联原因"));
        if (record.warning) details.append(contextNode("p", "检查提示：" + record.warning));
        details.append(contextNode("pre", record.context || "Context 留空，待资料确认后重新关联。"));
        const source = contextNode("small", "来源：" + record.source + (record.reference_source ? "；补充来源：" + record.reference_source : ""));
        details.append(source);
        const images = contextNode("div");
        images.className = "context-images";
        for (const item of record.images) {
            const link = contextNode("a");
            link.href = "/context/assets/" + contextTaskId + "/" + encodeURIComponent(item.name);
            link.target = "_blank";
            link.rel = "noopener";
            const image = contextNode("img");
            image.loading = "lazy";
            image.src = link.href;
            image.alt = item.caption;
            link.append(image, contextNode("small", item.scope));
            images.append(link);
        }
        details.append(images);
        cell.append(details);
        row.append(cell);
        fragment.append(row);
    }
    document.getElementById("context-rows").replaceChildren(fragment);
}

async function startContext() {
    const mainInput = document.getElementById("file-input");
    const status = document.getElementById("status-text");
    const download = document.getElementById("download-btn");
    const bundle = document.getElementById("bundle-btn");
    const controls = ["start-btn", "choose-file-btn", "file-input", "enhance-input", "reference-input", "guide-input"].map(id => document.getElementById(id));
    document.getElementById("progress-container").style.display = "block";
    document.getElementById("context-preview").hidden = true;
    download.style.display = "none";
    bundle.hidden = true;
    status.style.color = "";
    document.getElementById("progress-bar").style.width = "0%";
    if (!mainInput.files[0]) { status.textContent = "请先选择第一步的分类结果"; return; }
    controls.forEach(control => { control.disabled = true; });
    try {
        const form = new FormData();
        form.append("file", mainInput.files[0]);
        const enhanced = document.getElementById("enhance-input").checked;
        form.append("enhance", enhanced ? "1" : "0");
        if (enhanced) {
            form.append("reference_name", document.getElementById("reference-input").value);
            form.append("guide_name", document.getElementById("guide-input").value);
        }
        status.textContent = enhanced ? "正在读取本目录材料..." : "正在关联当前文件内的英文原文...";
        const {task_id: taskId} = await contextRequest("/upload/context", {method: "POST", body: form});
        while (true) {
            const task = await contextRequest("/progress/" + taskId);
            status.textContent = task.status;
            document.getElementById("progress-bar").style.width = task.progress + "%";
            if (task.state === "error") throw new Error(task.status);
            if (task.state === "completed") break;
            await new Promise(resolve => setTimeout(resolve, 500));
        }
        contextTaskId = taskId;
        contextResult = await contextRequest("/api/context/" + taskId);
        document.getElementById("context-summary").textContent = "共 " + contextResult.records.length + " 条；" +
            Object.entries(contextResult.counts).map(([key, value]) => key + " " + value + " 条").join("，") +
            (contextResult.enhanced ? "；已启用本目录可选补充" : "；仅关联当前上传文件");
        document.getElementById("image-warnings").textContent = contextResult.image_warnings.join("；");
        const filter = document.getElementById("category-filter");
        filter.replaceChildren(new Option("全部分类", ""));
        for (const category of new Set(contextResult.records.map(record => record.category))) filter.add(new Option(category, category));
        document.getElementById("status-filter").value = "";
        renderContextRows();
        document.getElementById("context-preview").hidden = false;
        download.onclick = () => { window.location.href = "/download/" + taskId; };
        download.style.display = "block";
        bundle.hidden = !contextResult.has_bundle;
        bundle.onclick = () => { window.location.href = "/download/context_bundle/" + taskId; };
    } catch (error) {
        status.textContent = error.message;
        status.style.color = "var(--error-color)";
    } finally {
        controls.forEach(control => { control.disabled = false; });
        toggleContextSupplements();
    }
}

document.addEventListener("DOMContentLoaded", () => {
    const fileInput = document.getElementById("file-input");
    document.getElementById("choose-file-btn").onclick = () => fileInput.click();
    document.getElementById("start-btn").onclick = startContext;
    document.getElementById("category-filter").onchange = renderContextRows;
    document.getElementById("status-filter").onchange = renderContextRows;
    for (const id of ["file-input", "reference-input", "guide-input"]) {
        document.getElementById(id).onchange = () => {
            document.getElementById("file-name").textContent = fileInput.files[0] ? fileInput.files[0].name : "点击选择第一步下载的分类文件（.xlsx）";
            document.getElementById("progress-container").style.display = "none";
            document.getElementById("context-preview").hidden = true;
            contextResult = null;
        };
    }
});

function toggleContextSupplements() {
    const enabled = document.getElementById("enhance-input").checked;
    document.getElementById("context-supplements").hidden = !enabled;
    document.getElementById("reference-input").disabled = !enabled;
    document.getElementById("guide-input").disabled = !enabled;
}

document.getElementById("enhance-input").addEventListener("change", () => {
    toggleContextSupplements();
    document.getElementById("progress-container").style.display = "none";
    contextResult = null;
});
