#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
问卷转换器Web应用启动脚本
"""

import os
import sys
import webbrowser
import time
from pathlib import Path

def check_dependencies():
    """检查依赖模块"""
    required_modules = [
        'flask',
        'werkzeug',
        'docx',
        'lxml'
    ]
    
    missing_modules = []
    
    for module in required_modules:
        try:
            if module == 'docx':
                import docx
            elif module == 'lxml':
                import lxml
            elif module == 'flask':
                import flask
            elif module == 'werkzeug':
                import werkzeug
        except ImportError:
            missing_modules.append(module)
    
    if missing_modules:
        print("❌ 缺少以下依赖模块:")
        for module in missing_modules:
            print(f"   - {module}")
        print("\n请运行以下命令安装依赖:")
        print("pip install flask python-docx lxml werkzeug")
        return False
    
    return True

def check_converter_modules():
    """检查转换器模块"""
    required_files = [
        'word_to_json.py',
        'survey_parser.py', 
        'xml_generator.py',
        'survey_converter.py'
    ]
    
    missing_files = []
    current_dir = Path(__file__).parent
    
    for file in required_files:
        if not (current_dir / file).exists():
            missing_files.append(file)
    
    if missing_files:
        print("❌ 缺少以下转换器模块:")
        for file in missing_files:
            print(f"   - {file}")
        print("\n请确保所有转换器模块都在当前目录中")
        return False
    
    return True

def start_web_app():
    """启动Web应用"""
    try:
        print("🚀 启动问卷转换器Web应用...")
        
        # 检查依赖
        if not check_dependencies():
            return False
        
        if not check_converter_modules():
            return False
        
        # 导入并启动应用
        from web_app import app
        
        print("✅ 所有依赖检查通过")
        print("🌐 Web应用正在启动...")
        print("📱 应用地址: http://localhost:5000")
        print("🔄 按 Ctrl+C 停止应用")
        print("-" * 50)
        
        # 延迟打开浏览器
        def open_browser():
            time.sleep(2)
            try:
                webbrowser.open('http://localhost:5000')
                print("🌐 已在浏览器中打开应用")
            except:
                print("⚠️ 无法自动打开浏览器，请手动访问 http://localhost:5000")
        
        import threading
        browser_thread = threading.Thread(target=open_browser)
        browser_thread.daemon = True
        browser_thread.start()
        
        # 启动Flask应用
        app.run(
            host='0.0.0.0',
            port=5000,
            debug=False,
            threaded=True
        )
        
    except KeyboardInterrupt:
        print("\n👋 应用已停止")
        return True
    except Exception as e:
        print(f"❌ 启动失败: {e}")
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("🔄 问卷转换器 Web 应用")
    print("=" * 60)
    
    success = start_web_app()
    
    if not success:
        print("\n按任意键退出...")
        input()
        sys.exit(1)