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
