#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
问卷转换器 - 一体化脚本
将Word文档一键转换为JSON和XML格式

用法:
    python survey_converter.py <word_file>
    
示例:
    python survey_converter.py survey_document.docx
    
输出:
    - survey_document.json (结构化问卷数据)
    - survey_document.xml (XML格式问卷)
"""

import sys
import os
import argparse
from pathlib import Path
import json
from datetime import datetime

# 导入现有模块
from word_to_json import WordToJsonConverter
from survey_parser import SurveyParser
from xml_generator import SurveyXMLGenerator


class SurveyConverter:
    """问卷转换器主类"""
    
    def __init__(self):
        self.word_converter = WordToJsonConverter()
        self.survey_parser = SurveyParser()
        self.xml_generator = SurveyXMLGenerator()
    
    def convert(self, word_file, output_dir=None, verbose=True):
        """
        一键转换Word文档为JSON和XML格式
        
        Args:
            word_file (str): Word文档路径
            output_dir (str): 输出目录，默认为Word文档所在目录
            verbose (bool): 是否显示详细信息
            
        Returns:
            dict: 包含生成文件路径的字典
        """
        try:
            # 验证输入文件
            word_path = Path(word_file)
            if not word_path.exists():
                raise FileNotFoundError(f"Word文档不存在: {word_file}")
            
            if not word_path.suffix.lower() in ['.docx', '.doc']:
                raise ValueError(f"不支持的文件格式: {word_path.suffix}")
            
            # 设置输出目录
            if output_dir is None:
                output_dir = word_path.parent
            else:
                output_dir = Path(output_dir)
                output_dir.mkdir(parents=True, exist_ok=True)
            
            # 生成输出文件名
            base_name = word_path.stem
            raw_json_file = output_dir / f"{base_name}_raw.json"
            structured_json_file = output_dir / f"{base_name}.json"
            xml_file = output_dir / f"{base_name}.xml"
            
            if verbose:
                print(f"🔄 开始转换: {word_file}")
                print(f"📁 输出目录: {output_dir}")
            
            # 步骤1: Word转原始JSON
            if verbose:
                print("📖 步骤1: 解析Word文档...")
            
            self.word_converter.convert_to_json(str(word_path), str(raw_json_file))
            
            if verbose:
                print(f"✅ 原始JSON已生成: {raw_json_file.name}")
            
            # 步骤2: 结构化问卷数据
            if verbose:
                print("🔧 步骤2: 结构化问卷数据...")
            
            self.survey_parser.parse_survey_document(str(word_path), str(structured_json_file))
            
            if verbose:
                print(f"✅ 结构化JSON已生成: {structured_json_file.name}")
            
            # 步骤3: 生成XML
            if verbose:
                print("📝 步骤3: 生成XML格式...")
            
            self.xml_generator.generate_xml(str(structured_json_file), str(xml_file), verbose)
            
            if verbose:
                print(f"✅ XML文件已生成: {xml_file.name}")
            
            # 统计信息
            if verbose:
                self._print_statistics(structured_json_file, xml_file)
            
            # 清理临时文件（可选）
            if raw_json_file.exists():
                raw_json_file.unlink()
                if verbose:
                    print("🧹 已清理临时文件")
            
            result = {
                'word_file': str(word_path),
                'json_file': str(structured_json_file),
                'xml_file': str(xml_file),
                'success': True
            }
            
            if verbose:
                print(f"\n🎉 转换完成！")
                print(f"📊 JSON文件: {structured_json_file}")
                print(f"📄 XML文件: {xml_file}")
            
            return result
            
        except Exception as e:
            error_msg = f"转换失败: {str(e)}"
            if verbose:
                print(f"❌ {error_msg}")
            
            return {
                'word_file': word_file,
                'error': error_msg,
                'success': False
            }
    
    def _print_statistics(self, json_file, xml_file):
        """打印统计信息"""
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            if 'questions' in data:
                questions = data['questions']
                total_questions = len(questions)
                
                # 统计问题类型
                type_counts = {}
                for q in questions:
                    q_type = q.get('question_type', 'Unknown')
                    type_counts[q_type] = type_counts.get(q_type, 0) + 1
                
                print(f"\n📊 转换统计:")
                print(f"   总问题数: {total_questions}")
                for q_type, count in type_counts.items():
                    print(f"   {q_type}: {count}个")
                
                # 文件大小
                json_size = os.path.getsize(json_file)
                xml_size = os.path.getsize(xml_file)
                print(f"   JSON文件大小: {json_size:,} 字节")
                print(f"   XML文件大小: {xml_size:,} 字节")
                
        except Exception as e:
            print(f"⚠️ 统计信息获取失败: {e}")


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='问卷转换器 - 将Word文档一键转换为JSON和XML格式',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法:
  python survey_converter.py survey_document.docx
  python survey_converter.py survey_document.docx --output-dir ./output
  python survey_converter.py survey_document.docx --quiet
        """
    )
    
    parser.add_argument(
        'word_file',
        help='Word文档文件路径 (.docx或.doc格式)'
    )
    
    parser.add_argument(
        '--output-dir', '-o',
        help='输出目录 (默认为Word文档所在目录)',
        default=None
    )
    
    parser.add_argument(
        '--quiet', '-q',
        action='store_true',
        help='静默模式，不显示详细信息'
    )
    
    parser.add_argument(
        '--version', '-v',
        action='version',
        version='问卷转换器 v1.0.0'
    )
    
    # 解析参数
    args = parser.parse_args()
    
    # 创建转换器并执行转换
    converter = SurveyConverter()
    result = converter.convert(
        word_file=args.word_file,
        output_dir=args.output_dir,
        verbose=not args.quiet
    )
    
    # 返回适当的退出码
    sys.exit(0 if result['success'] else 1)


if __name__ == "__main__":
    main()