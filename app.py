"""WC 工作台入口：翻译前准备、标签清洗和译后 QA。"""
import logging
import os
import threading
import uuid

from flask import Flask, jsonify, redirect, render_template, request, send_file
from flask import cli

from core_tools.context_processor import validate_material_name
from core_tools.qa_checker import QAChecker, QA_POLICY
from core_tools.tag_cleaner import TagCleaner, TAG_CLEAN_POLICY
from pipeline_api import bp as pipeline_bp

cli.show_server_banner = lambda *x: None
logging.getLogger('werkzeug').setLevel(logging.ERROR)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
app.register_blueprint(pipeline_bp)
tasks = {}

QA_DEFAULT_COLUMNS = {'source_column': 'game_text_en', 'target_column': 'game_text_de',
                      'key_column': 'key', 'limit_column': 'limit'}


@app.route('/')
def index():
    return redirect('/pipeline')


@app.route('/key_classifier')
@app.route('/context')
@app.route('/term')
@app.route('/term_filter')
@app.route('/tm')
def legacy_prepare():
    """兼容旧书签，准备操作统一进入工作台。"""
    return redirect('/pipeline')


@app.route('/tag_clean')
def page_tag_clean():
    return render_template('tag_clean.html', active_page='tag_clean')



@app.route('/upload/tag_clean', methods=['POST'])
def upload_tag_clean():
    uploaded = request.files.get('file')
    if not uploaded or not uploaded.filename:
        return jsonify({'error': '请选择要清洗的 XLSX 或 UTF-8 CSV 文件'}), 400
    name = uploaded.filename.replace(chr(92), '/').rsplit('/', 1)[-1]
    stem, extension = os.path.splitext(name)
    if extension.lower() not in ('.xlsx', '.csv'):
        return jsonify({'error': '请上传 XLSX 或 UTF-8 CSV 文件'}), 400
    try:
        validate_material_name(name)
        cleaner = TagCleaner(request.form.get('column', 'game_text_de'))
    except ValueError as error:
        return jsonify({'error': str(error)}), 400
    task_id = str(uuid.uuid4())
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{task_id}_clean_in{extension.lower()}')
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{task_id}_cleaned{extension.lower()}')
    uploaded.save(input_path)
    tasks[task_id] = {'progress': 0, 'state': 'processing', 'status': '正在准备清洗...',
                      'input_name': name, 'dl_name': stem + '_cleaned' + extension.lower()}
    threading.Thread(target=cleaner.process, args=(input_path, output_path, tasks, task_id)).start()
    return jsonify({'task_id': task_id})



@app.route('/api/tag_clean/<task_id>')
def tag_clean_result(task_id):
    task = tasks.get(task_id, {})
    if task.get('state') != 'completed' or task.get('tag_clean_policy') != TAG_CLEAN_POLICY:
        return jsonify({'error': '清洗结果不存在或尚未完成'}), 404
    return jsonify({key: task[key] for key in ('records', 'column', 'total', 'changed', 'skipped_formulas')})



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
    return jsonify({key: value for key, value in task.items()
                    if key not in ('records', 'file_issues', 'result_file')})


@app.route('/download/<task_id>')
def download_file(task_id):
    if task_id in tasks and tasks[task_id]['state'] == 'completed':
        file_path = tasks[task_id]['result_file']
        dl_name = tasks[task_id].get('dl_name', 'result.csv')
        return send_file(file_path, as_attachment=True, download_name=dl_name)
    return "文件暂不可用", 400



if __name__ == '__main__':
    print("WC 工作台：http://127.0.0.1:5001/pipeline")
    print("标签清洗：http://127.0.0.1:5001/tag_clean")
    print("译后 QA：http://127.0.0.1:5001/qa")
    debug = os.environ.get('WC_DEBUG') == '1'
    app.run(host='0.0.0.0', port=5001, debug=debug, use_reloader=debug)
