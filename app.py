"""WC 翻译准备工作台入口。"""
import logging
import os

from flask import Flask, redirect
from flask import cli

from pipeline_api import bp as pipeline_bp

cli.show_server_banner = lambda *x: None
logging.getLogger('werkzeug').setLevel(logging.ERROR)

app = Flask(__name__, static_folder=None)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024
app.register_blueprint(pipeline_bp)


@app.route('/')
def index():
    return redirect('/pipeline')


if __name__ == '__main__':
    print("WC 工作台：http://127.0.0.1:5001/pipeline")
    debug = os.environ.get('WC_DEBUG') == '1'
    app.run(host='0.0.0.0', port=5001, debug=debug, use_reloader=debug)
