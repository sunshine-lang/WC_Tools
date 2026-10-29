"""翻译准备工作台的 HTTP 接口。"""
from __future__ import annotations

import json

from flask import Blueprint, jsonify, make_response, render_template, request, send_file

from pipeline_service import PipelineError, service


bp = Blueprint('pipeline', __name__)


def _guard(fn):
    try:
        return jsonify(fn())
    except PipelineError as error:
        return jsonify({'error': str(error)}), 400
    except Exception as error:  # noqa: BLE001 — 本地工具，把异常变成可读提示
        return jsonify({'error': f'服务异常：{error}'}), 500


@bp.route('/pipeline')
def page():
    response = make_response(render_template('pipeline.html'))
    response.headers['Cache-Control'] = 'no-store, max-age=0'
    return response


@bp.route('/api/pipeline/bootstrap')
def bootstrap():
    return _guard(lambda: service.bootstrap(
        request.args.get('project') or None,
        request.args.get('batch') or None,
    ))


@bp.route('/api/pipeline/projects', methods=['POST'])
def create_project():
    data = request.get_json(silent=True) or {}
    return _guard(lambda: service.create_project(
        data.get('name', ''), data.get('src'), data.get('tgt'), data.get('use_tm', True)
    ))


@bp.route('/api/pipeline/projects/<project_id>/activate', methods=['POST'])
def activate_project(project_id):
    return _guard(lambda: service.activate_project(project_id))


@bp.route('/api/pipeline/projects/<project_id>/library/<slot>', methods=['POST'])
def upload_library(project_id, slot):
    return _guard(lambda: service.upload_library(project_id, slot, request.files.get('file')))


@bp.route('/api/pipeline/projects/<project_id>/library/<slot>', methods=['DELETE'])
def remove_library(project_id, slot):
    return _guard(lambda: service.remove_library(project_id, slot))


@bp.route('/api/pipeline/projects/<project_id>/batches', methods=['POST'])
def create_batch(project_id):
    def create():
        try:
            mapping = json.loads(request.form.get('columns_map', '{}'))
        except (ValueError, TypeError):
            raise PipelineError('列映射格式无效')
        if not isinstance(mapping, dict):
            raise PipelineError('列映射必须是对象')
        return service.create_batch(project_id, request.files.get('file'), mapping)
    return _guard(create)


@bp.route('/api/pipeline/projects/<project_id>/batches/<batch_id>/activate', methods=['POST'])
def activate_batch(project_id, batch_id):
    return _guard(lambda: service.activate_batch(project_id, batch_id))


@bp.route('/api/pipeline/projects/<project_id>/batches/<batch_id>/clone', methods=['POST'])
def clone_batch(project_id, batch_id):
    return _guard(lambda: service.clone_batch(project_id, batch_id))


@bp.route('/api/pipeline/projects/<project_id>/batches/<batch_id>/state')
def batch_state(project_id, batch_id):
    return _guard(lambda: service.batch_state(project_id, batch_id))


@bp.route('/api/pipeline/projects/<project_id>/batches/<batch_id>/run', methods=['POST'])
def run_batch(project_id, batch_id):
    data = request.get_json(silent=True) or {}
    return _guard(lambda: service.start(
        project_id, batch_id,
        stage=data.get('stage') or None,
        mode=data.get('mode') or 'run',
        then_continue=bool(data.get('then_continue')),
    ))


@bp.route('/api/pipeline/projects/<project_id>/batches/<batch_id>/confirm', methods=['POST'])
def confirm_batch(project_id, batch_id):
    return _guard(lambda: service.confirm(project_id, batch_id))


@bp.route('/api/pipeline/projects/<project_id>/batches/<batch_id>/classifications', methods=['POST'])
def save_classifications(project_id, batch_id):
    data = request.get_json(silent=True) or {}
    return _guard(lambda: service.save_classifications(project_id, batch_id, data.get('changes') if isinstance(data, dict) else None))


@bp.route('/api/pipeline/projects/<project_id>/batches/<batch_id>/reset', methods=['POST'])
def reset_batch(project_id, batch_id):
    return _guard(lambda: service.reset(project_id, batch_id))


@bp.route('/api/pipeline/projects/<project_id>/batches/<batch_id>/final')
def final_artifact(project_id, batch_id):
    try:
        path, name = service.final_artifact(project_id, batch_id)
        print(f'[download] final {batch_id} {name}', flush=True)
        return send_file(path, as_attachment=True, download_name=name)
    except PipelineError as error:
        return jsonify({'error': str(error)}), 404


@bp.route('/api/pipeline/projects/<project_id>/batches/<batch_id>/artifact/<stage_key>')
def artifact(project_id, batch_id, stage_key):
    try:
        path = service.artifact_path(project_id, batch_id, stage_key)
        print(f'[download] {batch_id} {stage_key} {path.name}', flush=True)
        return send_file(path, as_attachment=True, download_name=path.name)
    except PipelineError as error:
        return jsonify({'error': str(error)}), 404
