import os
import threading
import uuid
import json
import re
from datetime import datetime
from flask import Flask, request, jsonify, render_template, send_file, redirect
import logging
from flask import cli

from core_tools import TermScanner, TMMatcher, TermFilter, KeyClassifier, ContextProcessor, QAChecker
from core_tools.context_processor import validate_material_name, CONTEXT_POLICY
from core_tools.term_scanner import TERM_POLICY
from core_tools.tm_matcher import TM_POLICY
from core_tools.term_filter import FILTER_POLICY
from core_tools.qa_checker import QA_POLICY

# ==========================================
# Flask 初始化 & 屏蔽系统级输出 (Werkzeug)
# ==========================================
cli.show_server_banner = lambda *x: None
logging.getLogger('werkzeug').setLevel(logging.ERROR)

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')
app.config['DB_FOLDER'] = os.path.join(app.config['UPLOAD_FOLDER'], 'knowledge_db')

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['DB_FOLDER'], exist_ok=True)

# ===== 语料库及配置持久化路径 =====
TERM_DB_META_NAME = 'wc_term_db.json'

TM_DB_META_NAME = 'wc_tm_db.json'

knowledge_db_lock = threading.Lock()
key_classifier = KeyClassifier(os.path.join(app.config['DB_FOLDER'], 'key_rules.json'))


def read_metadata(path):
    try:
        with open(path, 'r', encoding='utf-8') as file:
            return json.load(file)
    except (OSError, ValueError):
        return {}


def write_metadata(path, data):
    temp_path = f"{path}.{uuid.uuid4().hex}.tmp"
    with open(temp_path, 'w', encoding='utf-8') as file:
        json.dump(data, file, ensure_ascii=False, indent=2)
    os.replace(temp_path, path)


def load_term_db():
    metadata = read_metadata(os.path.join(app.config['DB_FOLDER'], TERM_DB_META_NAME))
    stored = metadata.get('stored_filename', '')
    if not isinstance(stored, str) or not re.fullmatch(r'wc_term_[0-9a-f]{32}\.(xlsx|csv)', stored):
        raise ValueError('请先上传并初始化英德术语库')
    scanner = TermScanner(os.path.join(app.config['DB_FOLDER'], stored), source_name=metadata['filename'])
    return scanner, metadata


def get_term_db_status():
    try:
        scanner, metadata = load_term_db()
        return {**scanner.status(), 'updated_at': metadata['updated_at']}
    except Exception as error:
        return {'ready': False, 'error': str(error)}


def load_tm_db():
    metadata = read_metadata(os.path.join(app.config['DB_FOLDER'], TM_DB_META_NAME))
    stored = metadata.get('stored_filename', '')
    if not isinstance(stored, str) or not re.fullmatch(r'wc_tm_[0-9a-f]{32}\.tmx', stored):
        raise ValueError('请先上传并初始化英德翻译记忆库')
    matcher = TMMatcher(os.path.join(app.config['DB_FOLDER'], stored), source_name=metadata['filename'])
    return matcher, metadata


def get_tm_db_status():
    try:
        matcher, metadata = load_tm_db()
        return {**matcher.status(), 'updated_at': metadata['updated_at']}
    except Exception as error:
        return {'ready': False, 'error': str(error)}


tasks = {}
context_processor = ContextProcessor()

@app.route('/')
def index():
    return redirect('/key_classifier')


# ----------------- 页面渲染路由 -----------------
@app.route('/key_classifier')
def page_key_classifier():
    return render_template('key_classifier.html', active_page='key_classifier')


@app.route('/term')
def page_term():
    return render_template('term.html', active_page='term', term_db=get_term_db_status())


@app.route('/tm')
def page_tm():
    return render_template('tm.html', active_page='tm', tm_db=get_tm_db_status())


@app.route('/term_filter')
def page_term_filter():
    try:
        model = TermFilter().status()
    except Exception:
        model = {'ready': False, 'model': '', 'error': '本机模型配置无法读取，请检查后刷新页面'}
    return render_template('term_filter.html', active_page='term_filter', model=model)


@app.route('/api/key_rules', methods=['GET', 'POST'])
def handle_key_rules():
    try:
        if request.method == 'GET':
            return jsonify(key_classifier.load_rules())
        key_classifier.save_rules(request.get_json(silent=True))
        return jsonify({'success': True})
    except (ValueError, OSError) as error:
        return jsonify({'error': str(error)}), 400


@app.route('/upload/key_classifier', methods=['POST'])
def upload_key_classifier():
    file = request.files.get('file')
    if not file or file.filename == '':
        return jsonify({'error': '未选择文件'}), 400
    original_name, extension = os.path.splitext(os.path.basename(file.filename))
    if extension.lower() not in ('.xlsx', '.csv'):
        return jsonify({'error': '请上传 .xlsx 或 UTF-8 CSV 文件'}), 400
    dl_name = f"{original_name}_classified.xlsx"
    task_id = str(uuid.uuid4())
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_in{extension.lower()}")
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_out.xlsx")
    file.save(input_path)
    tasks[task_id] = {
        'progress': 0,
        'status': '等待处理...',
        'state': 'processing',
        'dl_name': dl_name
    }
    threading.Thread(target=key_classifier.process,
                     args=(input_path, output_path, tasks, task_id)).start()
    return jsonify({'task_id': task_id})


@app.route('/api/classification/<task_id>')
def classification_result(task_id):
    task = tasks.get(task_id)
    if not task or task.get('state') != 'completed' or 'records' not in task:
        return jsonify({'error': '分类结果不存在或尚未完成'}), 404
    return jsonify({'records': task['records'], 'counts': task['counts']})


# ----------------- 功能处理 API -----------------
@app.route('/upload/term_filter', methods=['POST'])
def upload_term_filter():
    file = request.files.get('file')
    if not file or not file.filename:
        return jsonify({'error': '请选择第三步下载的 XLSX 文件'}), 400
    name = file.filename.replace(chr(92), '/').rsplit('/', 1)[-1]
    if os.path.splitext(name)[1].lower() != '.xlsx':
        return jsonify({'error': '请上传第三步下载的 XLSX 文件'}), 400
    try:
        validate_material_name(name)
    except ValueError as error:
        return jsonify({'error': str(error)}), 400
    try:
        tool = TermFilter()
    except Exception:
        return jsonify({'error': '本机模型配置无法读取，请检查后重试'}), 400
    if not tool.status()['ready']:
        return jsonify({'error': '请先在本机配置有效的模型接口、模型名称和 API Key'}), 400
    task_id = str(uuid.uuid4())
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{task_id}_filter_in.xlsx')
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{task_id}_filtered.xlsx')
    file.save(input_path)
    tasks[task_id] = {'progress': 0, 'status': '等待语境筛选...', 'state': 'processing',
                      'input_name': name, 'dl_name': os.path.splitext(name)[0] + '_filtered.xlsx'}
    threading.Thread(target=tool.process, args=(input_path, output_path, tasks, task_id)).start()
    return jsonify({'task_id': task_id})


@app.route('/api/term_filter/<task_id>')
def term_filter_result(task_id):
    task = tasks.get(task_id, {})
    if task.get('state') != 'completed' or task.get('filter_policy') != FILTER_POLICY:
        return jsonify({'error': '术语筛选结果不存在或尚未完成'}), 404
    return jsonify({name: task[name] for name in ('records', 'counts', 'term_counts', 'error_count', 'model_requests', 'model')})


@app.route('/update/term_db', methods=['POST'])
def update_term_db():
    uploaded = request.files.get('tb_file')
    if not uploaded or not uploaded.filename:
        return jsonify({'error': '请选择术语库文件'}), 400
    filename = uploaded.filename.replace('\\', '/').rsplit('/', 1)[-1]
    extension = os.path.splitext(filename)[1].lower()
    if extension not in ('.xlsx', '.csv'):
        return jsonify({'error': '术语库仅支持 .xlsx 或 UTF-8 .csv，首行须包含 EN、DE'}), 400
    stored = f'wc_term_{uuid.uuid4().hex}{extension}'
    path = os.path.join(app.config['DB_FOLDER'], stored)
    committed = False
    try:
        validate_material_name(filename)
        uploaded.save(path)
        scanner = TermScanner(path, source_name=filename)
        metadata = {'stored_filename': stored, 'filename': filename,
                    'updated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'sha256': scanner.sha256}
        # Publish only a validated snapshot; failed uploads leave the active library intact.
        with knowledge_db_lock:
            write_metadata(os.path.join(app.config['DB_FOLDER'], TERM_DB_META_NAME), metadata)
            committed = True
        return jsonify({'success': True, 'message': '术语库已保存',
                        'term_db': {**scanner.status(), 'updated_at': metadata['updated_at']}})
    except Exception as error:
        return jsonify({'error': f'术语库更新失败：{error}'}), 400
    finally:
        if not committed and os.path.exists(path):
            os.remove(path)


@app.route('/upload/term', methods=['POST'])
def upload_term():
    trans_file = request.files.get('trans_file')
    if not trans_file or not trans_file.filename:
        return jsonify({'error': '未选择文件'}), 400
    original_name, extension = os.path.splitext(os.path.basename(trans_file.filename))
    if extension.lower() != '.xlsx':
        return jsonify({'error': '请上传第二步下载的 XLSX 文件'}), 400
    try:
        validate_material_name(trans_file.filename)
        scanner, _ = load_term_db()
    except Exception as error:
        return jsonify({'error': str(error)}), 400
    dl_name = f"{original_name}_tb.xlsx"
    task_id = str(uuid.uuid4())
    trans_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_trans.xlsx")
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_tb.xlsx")
    trans_file.save(trans_path)
    tasks[task_id] = {'progress': 0, 'status': '等待处理...', 'state': 'processing',
                      'dl_name': dl_name, 'input_name': os.path.basename(trans_file.filename)}
    threading.Thread(target=scanner.process, args=(trans_path, output_path, tasks, task_id)).start()
    return jsonify({'task_id': task_id})


@app.route('/api/term/<task_id>')
def term_result(task_id):
    task = tasks.get(task_id, {})
    if task.get('state') != 'completed' or task.get('term_policy') != TERM_POLICY:
        return jsonify({'error': '术语结果不存在或尚未完成'}), 404
    return jsonify({'records': task['records'], 'counts': task['counts'], 'term_db': task['term_db']})


# =============== 英德 TM 记忆库 ===============

@app.route('/update/tm_db', methods=['POST'])
def update_tm_db():
    uploaded = request.files.get('tm_tmx_file')
    if not uploaded or not uploaded.filename:
        return jsonify({'error': '请选择英德 TMX 文件'}), 400
    filename = uploaded.filename.replace(chr(92), '/').rsplit('/', 1)[-1]
    if os.path.splitext(filename)[1].lower() != '.tmx':
        return jsonify({'error': '请上传 .tmx 翻译记忆库'}), 400
    stored = f'wc_tm_{uuid.uuid4().hex}.tmx'
    path = os.path.join(app.config['DB_FOLDER'], stored)
    committed = False
    try:
        validate_material_name(filename)
        uploaded.save(path)
        matcher = TMMatcher(path, source_name=filename)
        metadata = {'stored_filename': stored, 'filename': filename,
                    'updated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'sha256': matcher.sha256}
        with knowledge_db_lock:
            write_metadata(os.path.join(app.config['DB_FOLDER'], TM_DB_META_NAME), metadata)
            committed = True
        return jsonify({'success': True, 'message': '英德翻译记忆库已保存',
                        'tm_db': {**matcher.status(), 'updated_at': metadata['updated_at']}})
    except Exception as error:
        return jsonify({'error': f'记忆库更新失败：{error}'}), 400
    finally:
        if not committed and os.path.exists(path):
            os.remove(path)


@app.route('/upload/tm', methods=['POST'])
def upload_tm():
    file = request.files.get('trans_file')
    if not file or not file.filename:
        return jsonify({'error': '请选择第四步术语筛选下载的 XLSX 文件'}), 400
    name = file.filename.replace(chr(92), '/').rsplit('/', 1)[-1]
    if os.path.splitext(name)[1].lower() != '.xlsx':
        return jsonify({'error': '请上传第四步术语筛选下载的 XLSX 文件'}), 400
    try:
        validate_material_name(name)
        matcher, _ = load_tm_db()
    except Exception as error:
        return jsonify({'error': str(error)}), 400
    task_id = str(uuid.uuid4())
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{task_id}_tm_in.xlsx')
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{task_id}_tm.xlsx')
    file.save(input_path)
    tasks[task_id] = {'progress': 0, 'status': '正在准备匹配...', 'state': 'processing',
                      'input_name': name, 'dl_name': os.path.splitext(name)[0] + '_tm.xlsx'}
    threading.Thread(target=matcher.process, args=(input_path, output_path, tasks, task_id)).start()
    return jsonify({'task_id': task_id})


@app.route('/api/tm/<task_id>')
def tm_result(task_id):
    task = tasks.get(task_id, {})
    if task.get('state') != 'completed' or task.get('tm_policy') != TM_POLICY:
        return jsonify({'error': 'TM 结果不存在或尚未完成'}), 404
    return jsonify({'records': task['records'], 'counts': task['counts'],
                    'matched_count': task['matched_count'], 'tm_db': task['tm_db']})


# =============== 译后质量检查 ===============

QA_DEFAULT_COLUMNS = {'source_column': 'game_text_en', 'target_column': 'game_text_de',
                      'key_column': 'key', 'limit_column': 'limit'}


@app.route('/qa')
def page_qa():
    return render_template('qa.html', active_page='qa', defaults=QA_DEFAULT_COLUMNS)


@app.route('/upload/qa', methods=['POST'])
def upload_qa():
    uploads = {'translated': request.files.get('translated_file'), 'original': request.files.get('original_file')}
    if not uploads['translated'] or not uploads['translated'].filename:
        return jsonify({'error': '请选择翻译回填后的文件'}), 400
    names = {}
    for role, uploaded in uploads.items():
        if uploaded and uploaded.filename:
            names[role] = uploaded.filename.replace(chr(92), '/').rsplit('/', 1)[-1]
            if os.path.splitext(names[role])[1].lower() not in ('.xlsx', '.csv'):
                return jsonify({'error': '仅支持 .xlsx 或 UTF-8 .csv 文件'}), 400
    columns = {name: (request.form.get(name) or '').strip() for name in QA_DEFAULT_COLUMNS}
    if not columns['source_column'] or not columns['target_column']:
        return jsonify({'error': '请填写源文列和译文列的表头名称'}), 400
    if columns['source_column'].casefold() == columns['target_column'].casefold():
        return jsonify({'error': '源文列和译文列不能相同'}), 400
    try:
        ratio = float(request.form.get('length_ratio') or 0)
    except ValueError:
        return jsonify({'error': '长度比例须为数字'}), 400
    if ratio and not 1 <= ratio <= 5:
        return jsonify({'error': '长度比例须在 1 到 5 之间，填 0 表示不检查'}), 400

    task_id = str(uuid.uuid4())
    paths = {}
    for role, name in names.items():
        paths[role] = os.path.join(app.config['UPLOAD_FOLDER'], f'{task_id}_qa_{role}{os.path.splitext(name)[1].lower()}')
        uploads[role].save(paths[role])
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{task_id}_qa_report.xlsx')
    checker = QAChecker(columns['source_column'], columns['target_column'],
                        columns['key_column'], columns['limit_column'], length_ratio=ratio)
    tasks[task_id] = {'progress': 0, 'status': '正在准备检查...', 'state': 'processing',
                      'input_name': names['translated'], 'original_name': names.get('original', ''),
                      'dl_name': os.path.splitext(names['translated'])[0] + '_QA报告.xlsx'}
    threading.Thread(target=checker.process,
                     args=(paths['translated'], output_path, tasks, task_id, paths.get('original'))).start()
    return jsonify({'task_id': task_id})


@app.route('/api/qa/<task_id>')
def qa_result(task_id):
    task = tasks.get(task_id, {})
    if task.get('state') != 'completed' or task.get('qa_policy') != QA_POLICY:
        return jsonify({'error': '检查结果不存在或尚未完成'}), 404
    return jsonify({name: task[name] for name in ('records', 'file_issues', 'counts', 'check_counts', 'total', 'translated')})


@app.route('/progress/<task_id>')
def get_progress(task_id):
    task = tasks.get(task_id, {'error': '任务不存在或已过期'})
    if 'asset_paths' in task and task.get('context_policy') != CONTEXT_POLICY:
        return jsonify({'error': '旧版上下文结果已失效，请重新处理'}), 410
    return jsonify({key: value for key, value in task.items()
                    if key not in ('records', 'file_issues', 'asset_paths', 'bundle_file', 'result_file')})


@app.route('/download/<task_id>')
def download_file(task_id):
    if task_id in tasks and 'asset_paths' in tasks[task_id] and tasks[task_id].get('context_policy') != CONTEXT_POLICY:
        return '旧版上下文结果已失效，请按当前材料规则重新处理', 410
    if task_id in tasks and tasks[task_id]['state'] == 'completed':
        file_path = tasks[task_id]['result_file']
        dl_name = tasks[task_id].get('dl_name', 'result.csv')
        return send_file(file_path, as_attachment=True, download_name=dl_name)
    return "文件暂不可用", 400


# ----------------- 上下文关联 -----------------
@app.route('/context')
def page_context():
    return render_template('context.html', active_page='context',
                           context_materials=context_processor.material_status(),
                           material_options=context_processor.material_options())


@app.route('/upload/context', methods=['POST'])
def upload_context():
    file = request.files.get('file')
    if not file or file.filename == '':
        return jsonify({'error': '未上传文件'}), 400

    original_name, extension = os.path.splitext(os.path.basename(file.filename))
    if extension.lower() != '.xlsx':
        return jsonify({'error': '请上传第一步下载的分类结果（.xlsx）'}), 400
    optional_files = {name: request.files.get(name) for name in ('reference_file', 'guide_file')}
    try:
        for uploaded in [file, *optional_files.values()]:
            if uploaded and uploaded.filename:
                validate_material_name(uploaded.filename)
        if any(uploaded and uploaded.filename for uploaded in optional_files.values()):
            raise ValueError('参考材料只支持选择 WC-UI 目录中的现有文件，请使用页面下拉框')
        reference_path = context_processor.resolve_material(request.form.get('reference_name') or context_processor.reference_path.name)
        guide_path = context_processor.resolve_material(request.form.get('guide_name') or context_processor.guide_path.name)
    except ValueError as error:
        return jsonify({'error': str(error)}), 400
    for uploaded in optional_files.values():
        if uploaded and uploaded.filename and os.path.splitext(uploaded.filename)[1].lower() != '.xlsx':
            return jsonify({'error': '参考文件和指导材料须为 .xlsx'}), 400
    dl_name = f"{original_name}_context.xlsx"
    task_id = str(uuid.uuid4())

    input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_in.xlsx")
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_context.xlsx")
    file.save(input_path)

    enhance = request.form.get('enhance') == '1'
    tasks[task_id] = {'progress': 0, 'status': '排队中...', 'state': 'processing',
                      'enhanced': enhance,
                      'dl_name': dl_name, 'input_name': os.path.basename(file.filename),
                      'guide_name': guide_path.name,
                      'reference_name': reference_path.name}
    threading.Thread(target=context_processor.process,
                     args=(input_path, output_path, tasks, task_id,
                           reference_path, guide_path, enhance)).start()
    return jsonify({'task_id': task_id})


@app.route('/api/context/<task_id>')
def context_result(task_id):
    task = tasks.get(task_id)
    if not task or task.get('state') != 'completed' or task.get('context_policy') != CONTEXT_POLICY:
        return jsonify({'error': '上下文结果不存在或尚未完成'}), 404
    return jsonify({'records': task['records'], 'counts': task['counts'],
                    'image_warnings': task['image_warnings'],
                    'reference_name': task['reference_name'], 'has_bundle': bool(task['bundle_file']),
                    'enhanced': task['enhanced']})


@app.route('/context/assets/<task_id>/<filename>')
def context_asset(task_id, filename):
    task = tasks.get(task_id, {})
    path = task.get('asset_paths', {}).get(filename)
    if task.get('state') != 'completed' or task.get('context_policy') != CONTEXT_POLICY or not path:
        return '图片不存在', 404
    return send_file(path)


@app.route('/download/context_bundle/<task_id>')
def context_bundle(task_id):
    task = tasks.get(task_id, {})
    if task.get('state') != 'completed' or task.get('context_policy') != CONTEXT_POLICY or not task.get('bundle_file'):
        return '附件包暂不可用', 404
    return send_file(task['bundle_file'], as_attachment=True,
                     download_name=os.path.splitext(task['dl_name'])[0] + '.zip')


if __name__ == '__main__':
    if os.environ.get('WERKZEUG_RUN_MAIN') == 'true':
        print("\n" + "=" * 45)
        print("  WC Tools 启动成功")
        print("  访问地址: http://127.0.0.1:5001")
        print("  使用 CTRL+C 终止运行")
        print("=" * 45 + "\n")
    app.run(host='0.0.0.0', port=5001, debug=True)
