// Основной JavaScript для веб-приложения валидатора

let currentTaskId = null;
let statusCheckInterval = null;
let isProcessing = false;

// Звуки: из папки static/sounds/ (upload.mp3, start.mp3, complete.mp3) или встроенный сигнал
const SOUND_BASE = '/static/sounds/';
const SOUND_FILES = { upload: 'upload', start: 'start', complete: 'complete' };

function playSound(type) {
    const name = SOUND_FILES[type];
    if (!name) return;
    const audio = new Audio(SOUND_BASE + name + '.mp3');
    let fallbackDone = false;
    function doFallback() {
        if (fallbackDone) return;
        fallbackDone = true;
        fallbackBeep(type);
    }
    audio.addEventListener('error', doFallback, { once: true });
    audio.play().catch(doFallback);
}

function fallbackBeep(type) {
    try {
        const ctx = new (window.AudioContext || window.webkitAudioContext)();
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.connect(gain);
        gain.connect(ctx.destination);
        gain.gain.setValueAtTime(0.15, ctx.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.2);
        if (type === 'upload') {
            osc.frequency.setValueAtTime(523, ctx.currentTime);
            osc.type = 'sine';
            osc.start(ctx.currentTime);
            osc.stop(ctx.currentTime + 0.15);
        } else if (type === 'start') {
            osc.frequency.setValueAtTime(659, ctx.currentTime);
            osc.type = 'sine';
            osc.start(ctx.currentTime);
            osc.stop(ctx.currentTime + 0.12);
            setTimeout(() => {
                const o2 = ctx.createOscillator();
                const g2 = ctx.createGain();
                o2.connect(g2);
                g2.connect(ctx.destination);
                g2.gain.setValueAtTime(0.15, ctx.currentTime);
                g2.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.12);
                o2.frequency.setValueAtTime(784, ctx.currentTime);
                o2.type = 'sine';
                o2.start(ctx.currentTime);
                o2.stop(ctx.currentTime + 0.12);
            }, 120);
        } else if (type === 'complete') {
            [523, 659, 784].forEach((freq, i) => {
                const o = ctx.createOscillator();
                const g = ctx.createGain();
                o.connect(g);
                g.connect(ctx.destination);
                g.gain.setValueAtTime(0.12, ctx.currentTime);
                g.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.25);
                o.frequency.setValueAtTime(freq, ctx.currentTime);
                o.type = 'sine';
                o.start(ctx.currentTime + i * 0.12);
                o.stop(ctx.currentTime + i * 0.12 + 0.25);
            });
        }
    } catch (e) { /* звуки опциональны */ }
}

// Инициализация при загрузке страницы
document.addEventListener('DOMContentLoaded', function() {
    initializeFileUpload();
    initializeForm();
});

// Инициализация загрузки файла
function initializeFileUpload() {
    const fileInput = document.getElementById('fileInput');
    const fileText = document.getElementById('fileText');
    const fileUploadArea = document.getElementById('fileUploadArea');
    const startButton = document.getElementById('startButton');

    // Обработка выбора файла
    fileInput.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            fileText.textContent = `Выбран файл: ${file.name}`;
            fileText.classList.add('has-file');
            startButton.disabled = false;
            playSound('upload');
        } else {
            fileText.textContent = 'Выберите файл или перетащите его сюда';
            fileText.classList.remove('has-file');
            startButton.disabled = true;
        }
    });

    // Drag and Drop
    fileUploadArea.addEventListener('dragover', function(e) {
        e.preventDefault();
        fileUploadArea.classList.add('dragover');
    });

    fileUploadArea.addEventListener('dragleave', function(e) {
        e.preventDefault();
        fileUploadArea.classList.remove('dragover');
    });

    fileUploadArea.addEventListener('drop', function(e) {
        e.preventDefault();
        fileUploadArea.classList.remove('dragover');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            const file = files[0];
            if (isValidFile(file)) {
                fileInput.files = files;
                fileText.textContent = `Выбран файл: ${file.name}`;
                fileText.classList.add('has-file');
                startButton.disabled = false;
                playSound('upload');
            } else {
                showStatus('error', 'Неподдерживаемый формат файла. Используйте .xlsx, .xls или .csv');
            }
        }
    });
}

// Проверка валидности файла
function isValidFile(file) {
    const allowedExtensions = ['xlsx', 'xls', 'csv'];
    const extension = file.name.split('.').pop().toLowerCase();
    return allowedExtensions.includes(extension);
}

// Инициализация формы
function initializeForm() {
    const startButton = document.getElementById('startButton');
    const stopButton = document.getElementById('stopButton');
    
    startButton.addEventListener('click', startValidation);
    stopButton.addEventListener('click', stopValidation);
}

// Запуск валидации
async function startValidation() {
    if (isProcessing) return;
    
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    
    if (!file) {
        showStatus('error', 'Пожалуйста, выберите файл');
        return;
    }

    // Собираем данные формы (SMTP всегда включён)
    const formData = new FormData();
    formData.append('file', file);
    formData.append('timeout', document.getElementById('timeout').value);
    formData.append('include_full_results_sheet', document.getElementById('include_full_results_sheet').checked);
    formData.append('only_valid_emails_sheet', document.getElementById('only_valid_emails_sheet').checked);
    
    const maxEmails = document.getElementById('max_emails').value;
    if (maxEmails) {
        formData.append('max_emails', maxEmails);
    }

    // Обновляем UI
    isProcessing = true;
    document.getElementById('startButton').style.display = 'none';
    document.getElementById('stopButton').style.display = 'inline-block';
    document.getElementById('progressCard').style.display = 'block';
    document.getElementById('downloadContainer').style.display = 'none';
    updateProgress(0, 'Загрузка файла...', '', 0, 0, null);

    try {
        // Загружаем файл
        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        if (!response.ok) {
            throw new Error(data.error || 'Ошибка при загрузке файла');
        }

        currentTaskId = data.task_id;
        showStatus('info', 'Файл загружен, валидация началась');
        playSound('start');
        
        // Начинаем проверку статуса
        startStatusCheck();

    } catch (error) {
        console.error('Ошибка:', error);
        showStatus('error', `Ошибка: ${error.message}`);
        resetUI();
    }
}

// Проверка статуса задачи
function startStatusCheck() {
    if (statusCheckInterval) {
        clearInterval(statusCheckInterval);
    }

    statusCheckInterval = setInterval(async () => {
        try {
            const response = await fetch(`/api/status/${currentTaskId}`);
            const data = await response.json();

            if (!response.ok) {
                throw new Error(data.error || 'Ошибка при получении статуса');
            }

            updateProgress(
                data.progress || 0,
                data.message || '',
                data.current_file || '',
                data.processed || 0,
                data.total || 0,
                data.eta_seconds
            );

            if (data.status === 'completed') {
                clearInterval(statusCheckInterval);
                statusCheckInterval = null;
                playSound('complete');
                showStatus('success', '✓ Проверка завершена успешно!');
                document.getElementById('downloadContainer').style.display = 'block';
                document.getElementById('downloadButton').onclick = () => downloadResult();
                resetUI();
            } else if (data.status === 'error') {
                clearInterval(statusCheckInterval);
                statusCheckInterval = null;
                showStatus('error', `Ошибка: ${data.error || data.message}`);
                resetUI();
            }

        } catch (error) {
            console.error('Ошибка при проверке статуса:', error);
            clearInterval(statusCheckInterval);
            statusCheckInterval = null;
            showStatus('error', `Ошибка: ${error.message}`);
            resetUI();
        }
    }, 2000); // Проверяем каждые 2 секунды
}

// Остановка валидации
function stopValidation() {
    if (statusCheckInterval) {
        clearInterval(statusCheckInterval);
        statusCheckInterval = null;
    }
    showStatus('info', 'Остановка валидации...');
    resetUI();
}

// Сброс UI
function resetUI() {
    isProcessing = false;
    document.getElementById('startButton').style.display = 'inline-block';
    document.getElementById('startButton').disabled = false;
    document.getElementById('stopButton').style.display = 'none';
}

// Обновление прогресса: процент, сообщение, файл, обработано, всего, осталось_сек
function updateProgress(progress, message, currentFile, processed, total, etaSeconds) {
    const progressFill = document.getElementById('progressFill');
    progressFill.style.width = `${progress}%`;

    document.getElementById('progressFile').textContent = currentFile || '—';
    document.getElementById('progressCount').textContent = total ? `${processed} из ${total}` : '—';
    document.getElementById('progressPercent').textContent = total ? `${progress.toFixed(1)}%` : '—';

    let etaStr = '—';
    if (etaSeconds != null && etaSeconds > 0) {
        if (etaSeconds >= 60) {
            etaStr = `${Math.floor(etaSeconds / 60)} мин ${Math.floor(etaSeconds % 60)} сек`;
        } else {
            etaStr = `${Math.floor(etaSeconds)} сек`;
        }
    }
    document.getElementById('progressEta').textContent = etaStr;
}

// Показать статус
function showStatus(type, message) {
    const statusMessage = document.getElementById('statusMessage');
    statusMessage.className = `status-message ${type}`;
    statusMessage.textContent = message;
    
    if (type === 'success') {
        setTimeout(() => {
            statusMessage.style.display = 'none';
        }, 5000);
    }
}

// Скачивание результата
async function downloadResult() {
    if (!currentTaskId) {
        showStatus('error', 'ID задачи не найден');
        return;
    }

    try {
        const response = await fetch(`/api/download/${currentTaskId}`);
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || 'Ошибка при скачивании файла');
        }

        // Получаем имя файла из заголовка
        const contentDisposition = response.headers.get('Content-Disposition');
        let filename = 'result.xlsx';
        if (contentDisposition) {
            const filenameMatch = contentDisposition.match(/filename="?(.+)"?/);
            if (filenameMatch) {
                filename = filenameMatch[1];
            }
        }

        // Создаем blob и скачиваем
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

        showStatus('success', 'Файл скачан успешно!');

    } catch (error) {
        console.error('Ошибка при скачивании:', error);
        showStatus('error', `Ошибка: ${error.message}`);
    }
}
