const ruleCategories = ["活动 UI 与商店文案", "集卡名称", "装饰物名称", "活动对白"];
let currentRules = {};
let activeRuleCategory = ruleCategories[0];
let rulesLoaded = false;
let rulesBusy = false;
let classificationBusy = false;

function ruleNode(tag, text, className) {
    const node = document.createElement(tag);
    if (text !== undefined) node.textContent = text;
    if (className) node.className = className;
    return node;
}

function renderRulePanel() {
    const tabs = document.getElementById("rule-tabs");
    tabs.replaceChildren();
    for (const category of [...ruleCategories, "待确认"]) {
        const button = ruleNode("button", category, "tab-btn" + (category === activeRuleCategory ? " active" : ""));
        button.type = "button";
        button.disabled = rulesBusy;
        button.setAttribute("aria-pressed", String(category === activeRuleCategory));
        button.onclick = () => { activeRuleCategory = category; renderRulePanel(); };
        tabs.append(button);
    }
    const editor = document.getElementById("rule-editor");
    editor.replaceChildren();
    if (activeRuleCategory === "待确认") {
        editor.append(ruleNode("p", "未命中规则或同时命中多类规则的文案，会标记为待确认。"));
        return;
    }
    const category = activeRuleCategory;
    const form = ruleNode("form", undefined, "rule-entry");
    const input = ruleNode("input");
    input.placeholder = "输入特征词...";
    input.maxLength = 120;
    input.disabled = rulesBusy || classificationBusy;
    input.setAttribute("aria-label", "输入特征词");
    const add = ruleNode("button", "添加特征");
    add.type = "submit";
    add.disabled = input.disabled;
    form.onsubmit = event => {
        event.preventDefault();
        const pattern = input.value.trim();
        if (!pattern || currentRules[category].includes(pattern)) return;
        saveRuleChange(category, [...currentRules[category], pattern]);
    };
    form.append(input, add);
    const features = ruleNode("div", undefined, "feature-list");
    for (const pattern of currentRules[category]) {
        const tag = ruleNode("span", undefined, "feature-tag");
        const remove = ruleNode("button", "×");
        remove.type = "button";
        remove.disabled = input.disabled;
        remove.setAttribute("aria-label", "删除特征 " + pattern);
        remove.onclick = () => saveRuleChange(category, currentRules[category].filter(item => item !== pattern));
        tag.append(ruleNode("span", pattern), remove);
        features.append(tag);
    }
    if (!features.childElementCount) features.append(ruleNode("span", "当前分类暂无特征词", "rule-help"));
    editor.append(form, features);
    if (category === "活动对白") {
        editor.append(ruleNode("p", "说话人和情绪字段会优先识别为对白。", "rule-help"));
    }
}

async function saveRuleChange(category, patterns) {
    if (rulesBusy || classificationBusy) return;
    rulesBusy = true;
    document.getElementById("start-btn").disabled = true;
    document.getElementById("rule-status").textContent = "正在保存...";
    renderRulePanel();
    try {
        const updated = {...currentRules, [category]: patterns};
        await classificationRequest("/api/key_rules", {
            method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(updated)
        });
        currentRules = updated;
        document.getElementById("rule-status").textContent = "已自动保存";
    } catch (error) {
        document.getElementById("rule-status").textContent = "保存失败：" + error.message;
    } finally {
        rulesBusy = false;
        renderRulePanel();
        document.getElementById("start-btn").disabled = classificationBusy;
    }
}

async function classificationRequest(url, options) {
    const response = await fetch(url, options);
    const data = await response.json();
    if (!response.ok || data.error) throw new Error(data.error || "请求失败");
    return data;
}

function showClassificationPreview(result) {
    document.getElementById("classification-summary").textContent =
        "共 " + result.records.length + " 条 · " +
        Object.entries(result.counts).filter(([, count]) => count > 0)
            .map(([category, count]) => category + " " + count + " 条").join("，");
    const fragment = document.createDocumentFragment();
    for (const record of result.records) {
        const row = document.createElement("tr");
        for (const value of [record.row, record.key || record.speaker || "无 Key", record.english, record.category]) {
            const cell = document.createElement("td");
            cell.textContent = value;
            row.append(cell);
        }
        fragment.append(row);
    }
    document.getElementById("preview-rows").replaceChildren(fragment);
    document.getElementById("classification-preview").hidden = false;
}

async function startClassification() {
    if (!rulesLoaded || rulesBusy) return;
    const fileInput = document.getElementById("file-input");
    const startButton = document.getElementById("start-btn");
    const chooseButton = document.getElementById("choose-file-btn");
    const downloadButton = document.getElementById("download-btn");
    const status = document.getElementById("status-text");
    const progress = document.getElementById("progress-bar");
    document.getElementById("progress-container").style.display = "block";
    document.getElementById("classification-preview").hidden = true;
    downloadButton.style.display = "none";
    progress.style.width = "0%";
    status.style.color = "";
    if (!fileInput.files[0]) {
        status.textContent = "请先选择待处理文件";
        return;
    }
    classificationBusy = true;
    renderRulePanel();
    startButton.disabled = chooseButton.disabled = fileInput.disabled = true;
    status.textContent = "正在上传文件...";
    try {
        const form = new FormData();
        form.append("file", fileInput.files[0]);
        const {task_id: taskId} = await classificationRequest("/upload/key_classifier", {method: "POST", body: form});
        while (true) {
            const task = await classificationRequest("/progress/" + taskId);
            status.textContent = task.status;
            progress.style.width = task.progress + "%";
            if (task.state === "error") throw new Error(task.status);
            if (task.state === "completed") break;
            await new Promise(resolve => setTimeout(resolve, 500));
        }
        showClassificationPreview(await classificationRequest("/api/classification/" + taskId));
        downloadButton.onclick = () => { window.location.href = "/download/" + taskId; };
        downloadButton.style.display = "block";
    } catch (error) {
        status.textContent = error.message;
        status.style.color = "var(--error-color)";
    } finally {
        classificationBusy = false;
        renderRulePanel();
        startButton.disabled = chooseButton.disabled = fileInput.disabled = false;
    }
}

document.addEventListener("DOMContentLoaded", async () => {
    const fileInput = document.getElementById("file-input");
    document.getElementById("choose-file-btn").onclick = () => fileInput.click();
    document.getElementById("start-btn").onclick = startClassification;
    fileInput.onchange = () => {
        const label = document.getElementById("file-name");
        label.textContent = fileInput.files[0] ? fileInput.files[0].name : "点击选择 Excel 或 CSV 文件";
        label.className = fileInput.files[0] ? "file-name" : "";
        document.getElementById("progress-container").style.display = "none";
        document.getElementById("classification-preview").hidden = true;
        document.getElementById("download-btn").style.display = "none";
    };
    try {
        currentRules = await classificationRequest("/api/key_rules");
        rulesLoaded = true;
        renderRulePanel();
        document.getElementById("rule-status").textContent = "";
        document.getElementById("start-btn").disabled = false;
    } catch (error) {
        document.getElementById("rule-status").textContent = "规则加载失败：" + error.message;
    }
});
