#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
问卷转换器 - Web应用程序
基于Flask的Web版本，支持在线文件上传和转换

功能:
- 文件上传接口
- 在线转换服务
- 结果下载
- 进度查询
"""

import sys
import os
import json
import uuid
import threading
import time
from datetime import datetime
from pathlib import Path
from werkzeug.utils import secure_filename
from flask import Flask, request, jsonify, render_template, send_file, send_from_directory, abort

# 环境变量支持
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    from flask_cors import CORS
except ImportError:
    print("警告: flask_cors未安装，跨域功能将不可用")
    CORS = None

# 确保能导入本地模块
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

# 导入转换器
try:
    from survey_converter import SurveyConverter
except ImportError as e:
    print(f"警告: 无法导入SurveyConverter: {e}")
    print("请确保survey_converter.py及其依赖模块存在")
    SurveyConverter = None


# 创建Flask应用
app = Flask(__name__)

# 配置应用
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'survey-converter-secret-key-change-in-production')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB最大文件大小

# 环境检测
IS_PRODUCTION = os.environ.get('FLASK_ENV') == 'production'

# 启用CORS支持
if CORS:
    CORS(app)

# 配置目录
UPLOAD_FOLDER = Path('uploads')
OUTPUT_FOLDER = Path('outputs')
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {'doc', 'docx'}

# 全局变量存储转换任务状态
conversion_tasks = {}


class ConversionTask:
    """转换任务类"""
    
    def __init__(self, task_id, filename, file_path=None):
        self.task_id = task_id
        self.filename = filename
        self.file_path = file_path
        self.status = 'pending'  # pending, processing, completed, failed
        self.progress = 0
        self.message = '等待开始...'
        self.start_time = datetime.now()
        self.end_time = None
        self.result_files = {}
        self.error_message = None
        
    def update_status(self, status, progress=None, message=None):
        """更新任务状态"""
        self.status = status
        if progress is not None:
            self.progress = progress
        if message is not None:
            self.message = message
        if status in ['completed', 'failed']:
            self.end_time = datetime.now()
            
    def to_dict(self):
        """转换为字典格式"""
        return {
            'task_id': self.task_id,
            'filename': self.filename,
            'status': self.status,
            'progress': self.progress,
            'message': self.message,
            'start_time': self.start_time.isoformat(),
            'end_time': self.end_time.isoformat() if self.end_time else None,
            'result_files': self.result_files,
            'error_message': self.error_message
        }


def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def perform_conversion(task_id):
    """执行转换任务（在后台线程中运行）"""
    task = conversion_tasks[task_id]
    
    try:
        task.update_status('processing', 10, '开始解析Word文档...')
        
        # 检查转换器是否可用
        if SurveyConverter is None:
            raise Exception("转换器模块未正确加载，请检查依赖模块")
        
        # 从task对象获取文件路径
        input_file = task.file_path
        
        # 设置输出目录
        output_dir = os.path.join(OUTPUT_FOLDER, task_id)
        os.makedirs(output_dir, exist_ok=True)
        
        # 创建转换器实例
        converter = SurveyConverter()
        
        task.update_status('processing', 30, '解析Word文档...')
        
        # 执行转换
        result = converter.convert(
            word_file=str(input_file),
            output_dir=str(output_dir),
            verbose=False
        )
        
        task.update_status('processing', 80, '生成输出文件...')
        
        if result.get('success'):
            # 转换成功
            json_file = result.get('json_file')
            xml_file = result.get('xml_file')
            
            # 验证文件是否存在
            json_exists = json_file and Path(json_file).exists()
            xml_exists = xml_file and Path(xml_file).exists()
            
            if json_exists or xml_exists:
                task.result_files = {}
                
                if json_exists:
                    json_path = Path(json_file)
                    task.result_files['json'] = {
                        'filename': json_path.name,
                        'path': str(json_path.relative_to(OUTPUT_FOLDER)),
                        'size': json_path.stat().st_size
                    }
                
                if xml_exists:
                    xml_path = Path(xml_file)
                    task.result_files['xml'] = {
                        'filename': xml_path.name,
                        'path': str(xml_path.relative_to(OUTPUT_FOLDER)),
                        'size': xml_path.stat().st_size
                    }
                
                task.update_status('completed', 100, '转换完成！')
            else:
                task.error_message = '输出文件未生成'
                task.update_status('failed', 0, '转换失败: 输出文件未生成')
        else:
            task.error_message = result.get('error', '未知错误')
            task.update_status('failed', 0, f'转换失败: {task.error_message}')
            
    except Exception as e:
        task.error_message = str(e)
        task.update_status('failed', 0, f'发生异常: {str(e)}')
        print(f"转换任务 {task_id} 失败: {e}")
        import traceback
        traceback.print_exc()


@app.route('/')
def index():
    """主页"""
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """文件上传接口"""
    try:
        # 检查是否有文件
        if 'file' not in request.files:
            return jsonify({'error': '没有选择文件'}), 400
        
        file = request.files['file']
        
        # 检查文件名
        if file.filename == '':
            return jsonify({'error': '没有选择文件'}), 400
        
        # 检查文件类型
        if not allowed_file(file.filename):
            return jsonify({'error': '不支持的文件格式，请上传.doc或.docx文件'}), 400
        
        # 生成任务ID和安全文件名
        task_id = str(uuid.uuid4())
        filename = secure_filename(file.filename)
        
        # 创建任务专用目录
        task_dir = OUTPUT_FOLDER / task_id
        task_dir.mkdir(exist_ok=True)
        
        # 保存上传的文件
        upload_path = UPLOAD_FOLDER / f"{task_id}_{filename}"
        file.save(upload_path)
        
        # 创建转换任务
        task = ConversionTask(task_id, filename, file_path=upload_path)
        conversion_tasks[task_id] = task
        
        # 在后台线程中执行转换
        thread = threading.Thread(
            target=perform_conversion,
            args=(task_id,)
        )
        thread.daemon = True
        thread.start()
        
        return jsonify({
            'task_id': task_id,
            'filename': filename,
            'message': '文件上传成功，开始转换...'
        })
        
    except Exception as e:
        return jsonify({'error': f'上传失败: {str(e)}'}), 500


@app.route('/api/status/<task_id>')
def get_status(task_id):
    """获取转换状态"""
    if task_id not in conversion_tasks:
        return jsonify({'error': '任务不存在'}), 404
    
    task = conversion_tasks[task_id]
    return jsonify(task.to_dict())


@app.route('/api/download/<task_id>/<file_type>')
def download_file(task_id, file_type):
    """下载转换结果文件"""
    if task_id not in conversion_tasks:
        return jsonify({'error': '任务不存在'}), 404
    
    task = conversion_tasks[task_id]
    
    if task.status != 'completed':
        return jsonify({'error': '转换尚未完成'}), 400
    
    if file_type not in task.result_files:
        return jsonify({'error': '文件类型不存在'}), 404
    
    file_info = task.result_files[file_type]
    file_path = OUTPUT_FOLDER / file_info['path']
    
    if not file_path.exists():
        return jsonify({'error': '文件不存在'}), 404
    
    return send_file(
        file_path,
        as_attachment=True,
        download_name=file_info['filename']
    )


@app.route('/api/tasks')
def list_tasks():
    """获取所有任务列表"""
    tasks = []
    for task in conversion_tasks.values():
        tasks.append(task.to_dict())
    
    # 按开始时间倒序排列
    tasks.sort(key=lambda x: x['start_time'], reverse=True)
    
    return jsonify(tasks)


@app.route('/api/clear/<task_id>', methods=['DELETE'])
def clear_task(task_id):
    """清除指定任务"""
    if task_id not in conversion_tasks:
        return jsonify({'error': '任务不存在'}), 404
    
    try:
        # 删除相关文件
        task_dir = OUTPUT_FOLDER / task_id
        upload_file = UPLOAD_FOLDER / f"{task_id}_*"
        
        # 删除任务目录
        if task_dir.exists():
            import shutil
            shutil.rmtree(task_dir)
        
        # 删除上传文件
        for file_path in UPLOAD_FOLDER.glob(f"{task_id}_*"):
            file_path.unlink()
        
        # 从内存中删除任务
        del conversion_tasks[task_id]
        
        return jsonify({'message': '任务已清除'})
        
    except Exception as e:
        return jsonify({'error': f'清除失败: {str(e)}'}), 500


@app.route('/static/<path:filename>')
def static_files(filename):
    """静态文件服务"""
    return send_from_directory('static', filename)


@app.errorhandler(413)
def too_large(e):
    """文件过大错误处理"""
    return jsonify({'error': '文件过大，请上传小于50MB的文件'}), 413


@app.errorhandler(500)
def internal_error(e):
    """内部服务器错误处理"""
    return jsonify({'error': '服务器内部错误'}), 500


if __name__ == '__main__':
    print("=" * 60)
    print("问卷转换器 Web 应用程序")
    print("=" * 60)
    
    # 获取端口和主机配置
    port = int(os.environ.get('PORT', 5000))
    host = os.environ.get('HOST', '0.0.0.0')
    debug = not IS_PRODUCTION
    
    print(f"🌐 服务地址: http://{host}:{port}")
    print(f"📁 上传目录: {UPLOAD_FOLDER.absolute()}")
    print(f"📁 输出目录: {OUTPUT_FOLDER.absolute()}")
    print(f"🔧 环境模式: {'生产环境' if IS_PRODUCTION else '开发环境'}")
    print("=" * 60)
    
    # 启动Flask应用
    app.run(
        host=host,
        port=port,
        debug=debug,
        threaded=True
    )