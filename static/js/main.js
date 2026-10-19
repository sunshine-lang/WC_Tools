/* /static/js/main.js */
function updateFileInput(inputId, nameId, defaultText) {
    const fileInput = document.getElementById(inputId);
    const nameDisplay = document.getElementById(nameId);
    if (fileInput.files.length > 0) {
        nameDisplay.innerHTML = `<span class="file-name">✓ 已选择: ${fileInput.files[0].name}</span>`;
        nameDisplay.parentElement.style.borderColor = 'var(--success-color)';
        nameDisplay.parentElement.style.backgroundColor = '#f0fdf4';
    } else {
        nameDisplay.textContent = defaultText;
        nameDisplay.parentElement.style.borderColor = 'var(--border-color)';
        nameDisplay.parentElement.style.backgroundColor = '#f8fafc';
    }
}

function startPolling(taskId, btnId) {
    const progressContainer = document.getElementById('progress-container');
    const progressBar = document.getElementById('progress-bar');
    const statusText = document.getElementById('status-text');
    const downloadBtn = document.getElementById('download-btn');
    const startBtn = document.getElementById(btnId);
    progressContainer.style.display = 'block';
    downloadBtn.style.display = 'none';
    progressBar.style.width = '0%';
    statusText.textContent = '初始化中...';
    startBtn.disabled = true;

    function poll() {
        fetch('/progress/' + taskId)
            .then(r => r.json())
            .then(data => {
                if (data.error) {
                    statusText.textContent = '报错: ' + data.error;
                    statusText.style.color = 'var(--error-color)';
                    startBtn.disabled = false;
                    return;
                }
                progressBar.style.width = data.progress + '%';
                statusText.textContent = data.status;

                if (data.state === 'processing') {
                    setTimeout(poll, 1000);
                } else if (data.state === 'completed') {
                    progressBar.style.backgroundColor = 'var(--success-color)';
                    downloadBtn.style.display = 'block';
                    startBtn.disabled = false;
                    startBtn.textContent = '再次处理';

                    downloadBtn.onclick = () => {
                        window.location.href = '/download/' + taskId;
                    };
                } else if (data.state === 'error') {
                    progressBar.style.backgroundColor = 'var(--error-color)';
                    statusText.style.color = 'var(--error-color)';
                    startBtn.disabled = false;
                }
            })
            .catch(err => {
                statusText.textContent = '网络请求失败';
                statusText.style.color = 'var(--error-color)';
                startBtn.disabled = false;
            });
    }
    poll();
}

function submitForm(endpoint, inputs, btnId) {
    const formData = new FormData();
    let allValid = true;

    inputs.forEach(input => {
        const el = document.getElementById(input.id);
        if (!el) return;

        // 如果配置了 type = 'value'，说明是从普通控件（如 select/input text）取值
        if (input.type === 'value') {
            formData.append(input.param, el.value);
        } else {
            // 否则默认是取文件 (file)
            const isRequired = input.required !== undefined ? input.required : true;
            if (isRequired && el.files.length === 0) {
                allValid = false;
            } else if (el.files.length > 0) {
                formData.append(input.param, el.files[0]);
            }
        }
    });

    if (!allValid) {
        alert('请先上传标为“必填”的文件！');
        return;
    }
    const startBtn = document.getElementById(btnId);
    const originalLabel = startBtn.dataset.originalLabel || startBtn.textContent.trim();
    startBtn.disabled = true;
    startBtn.textContent = '上传处理中...';

    fetch(endpoint, {
        method: 'POST',
        body: formData
    })
    .then(r => r.json())
    .then(data => {
        if (data.error) {
            alert(data.error);
            startBtn.disabled = false;
            startBtn.textContent = originalLabel;
        } else {
            startPolling(data.task_id, btnId);
        }
    })
    .catch(err => {
        alert('上传失败，请检查网络或后台服务运行状态。');
        startBtn.disabled = false;
        startBtn.textContent = originalLabel;
    });
}