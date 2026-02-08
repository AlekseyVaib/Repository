"""
Flask веб-приложение для валидации email адресов - SaaS версия
"""

from flask import Flask, render_template, request, jsonify, send_file, session
import os
import uuid
import threading
from datetime import datetime
from werkzeug.utils import secure_filename
import logging
from email_validator import process_excel_file, EmailValidator

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'your-secret-key-change-this-in-production')
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['RESULTS_FOLDER'] = 'results'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB max file size

# Создаем папки если их нет
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['RESULTS_FOLDER'], exist_ok=True)

# Хранилище задач (в продакшене использовать Redis или БД)
tasks = {}

# Разрешенные расширения файлов
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}


def allowed_file(filename):
    """Проверка расширения файла"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def _format_eta(seconds):
    """Форматирование оставшегося времени."""
    if seconds is None or seconds <= 0:
        return ''
    if seconds >= 60:
        return f"{int(seconds // 60)} мин {int(seconds % 60)} сек"
    return f"{int(seconds)} сек"


def process_validation_task(task_id, file_path, options):
    """Обработка задачи валидации в отдельном потоке"""
    try:
        tasks[task_id]['status'] = 'processing'
        tasks[task_id]['progress'] = 0
        tasks[task_id]['message'] = 'Начало обработки файла...'
        tasks[task_id]['current_file'] = ''
        tasks[task_id]['processed'] = 0
        tasks[task_id]['total'] = 0
        tasks[task_id]['eta_seconds'] = None
        
        # Генерируем имя выходного файла
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"validated_{task_id}_{timestamp}.xlsx"
        output_path = os.path.join(app.config['RESULTS_FOLDER'], output_filename)
        
        def progress_callback(current_file, processed, total, percent, eta_seconds):
            tasks[task_id]['progress'] = round(percent, 1)
            tasks[task_id]['current_file'] = current_file
            tasks[task_id]['processed'] = processed
            tasks[task_id]['total'] = total
            tasks[task_id]['eta_seconds'] = eta_seconds
            eta_str = _format_eta(eta_seconds)
            tasks[task_id]['message'] = (
                f"Файл: {current_file} — обработано {processed} из {total} ({percent:.1f}%)"
                + (f", осталось ~{eta_str}" if eta_str else "")
            )
        
        # Запускаем валидацию
        logger.info(f"Запуск валидации для задачи {task_id}")
        
        process_excel_file(
            input_file=file_path,
            output_file=output_path,
            check_smtp=options.get('check_smtp', True),
            timeout=options.get('timeout', 10),
            accept_catch_all=False,  # опция убрана, catch-all помечаются как X
            max_emails=options.get('max_emails'),
            validation_mode=options.get('validation_mode', 'strict'),
            progress_callback=progress_callback
        )
        
        # Обновляем статус задачи
        tasks[task_id]['status'] = 'completed'
        tasks[task_id]['progress'] = 100
        tasks[task_id]['message'] = 'Валидация завершена успешно!'
        tasks[task_id]['eta_seconds'] = 0
        tasks[task_id]['result_file'] = output_filename
        tasks[task_id]['result_path'] = output_path
        
        logger.info(f"Валидация завершена для задачи {task_id}")
        
    except Exception as e:
        logger.error(f"Ошибка при валидации задачи {task_id}: {e}", exc_info=True)
        tasks[task_id]['status'] = 'error'
        tasks[task_id]['message'] = f'Ошибка: {str(e)}'
        tasks[task_id]['error'] = str(e)


@app.route('/')
def index():
    """Главная страница"""
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Загрузка файла и запуск валидации"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Файл не найден'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'Файл не выбран'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'Неподдерживаемый формат файла. Используйте .xlsx, .xls или .csv'}), 400
        
        # Получаем параметры из формы (опция accept_catch_all убрана — catch-all помечаются как X)
        options = {
            'check_smtp': request.form.get('check_smtp', 'true').lower() == 'true',
            'timeout': int(request.form.get('timeout', 10)),
            'validation_mode': request.form.get('validation_mode', 'strict'),
            'max_emails': request.form.get('max_emails') or None
        }
        
        if options['max_emails']:
            try:
                options['max_emails'] = int(options['max_emails'])
            except ValueError:
                options['max_emails'] = None
        
        # Генерируем уникальный ID задачи
        task_id = str(uuid.uuid4())
        
        # Сохраняем файл
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_{filename}")
        file.save(file_path)
        
        # Создаем задачу
        tasks[task_id] = {
            'status': 'pending',
            'progress': 0,
            'message': 'Ожидание начала обработки...',
            'filename': filename,
            'created_at': datetime.now().isoformat()
        }
        
        # Запускаем обработку в отдельном потоке
        thread = threading.Thread(
            target=process_validation_task,
            args=(task_id, file_path, options),
            daemon=True
        )
        thread.start()
        
        return jsonify({
            'task_id': task_id,
            'message': 'Файл загружен, валидация началась'
        }), 200
        
    except Exception as e:
        logger.error(f"Ошибка при загрузке файла: {e}", exc_info=True)
        return jsonify({'error': f'Ошибка: {str(e)}'}), 500


@app.route('/api/status/<task_id>')
def get_status(task_id):
    """Получение статуса задачи"""
    if task_id not in tasks:
        return jsonify({'error': 'Задача не найдена'}), 404
    
    task = tasks[task_id]
    return jsonify({
        'status': task['status'],
        'progress': task.get('progress', 0),
        'message': task.get('message', ''),
        'current_file': task.get('current_file', ''),
        'processed': task.get('processed', 0),
        'total': task.get('total', 0),
        'eta_seconds': task.get('eta_seconds'),
        'result_file': task.get('result_file'),
        'error': task.get('error')
    }), 200


@app.route('/api/download/<task_id>')
def download_file(task_id):
    """Скачивание результата"""
    if task_id not in tasks:
        return jsonify({'error': 'Задача не найдена'}), 404
    
    task = tasks[task_id]
    
    if task['status'] != 'completed':
        return jsonify({'error': 'Задача еще не завершена'}), 400
    
    if 'result_path' not in task:
        return jsonify({'error': 'Файл результата не найден'}), 404
    
    result_path = task['result_path']
    if not os.path.exists(result_path):
        return jsonify({'error': 'Файл не существует'}), 404
    
    return send_file(
        result_path,
        as_attachment=True,
        download_name=task.get('result_file', 'result.xlsx'),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/api/tasks')
def list_tasks():
    """Список всех задач (для отладки)"""
    return jsonify(tasks), 200


if __name__ == '__main__':
    # Для разработки
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
    # Для продакшена используйте:
    # gunicorn -w 4 -b 0.0.0.0:$PORT app:app
