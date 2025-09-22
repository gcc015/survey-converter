#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
XML生成器 - 将survey_output.json转换为XML格式
根据问题类型生成对应的XML结构
"""

import json
import xml.etree.ElementTree as ET
from xml.dom import minidom
import argparse
import sys

class SurveyXMLGenerator:
    def __init__(self):
        self.root = ET.Element("survey")
    
    def escape_xml_text(self, text):
        """转义XML特殊字符"""
        if not text:
            return ""
        # 替换XML特殊字符
        text = str(text)
        text = text.replace("&", "&amp;")
        text = text.replace("<", "&lt;")
        text = text.replace(">", "&gt;")
        text = text.replace('"', "&quot;")
        text = text.replace("'", "&apos;")
        return text
    
    def generate_single_answer_xml(self, question):
        """生成Single Answer类型的XML"""
        radio = ET.Element("radio")
        radio.set("label", question.get("question_id", ""))
        
        # 添加title
        title = ET.SubElement(radio, "title")
        title.text = self.escape_xml_text(question.get("question_text", ""))
        
        # 添加comment
        comment = ET.SubElement(radio, "comment")
        comment.text = "${res.SCInst}"
        
        # 添加选项
        for option in question.get("question_options", []):
            row = ET.SubElement(radio, "row")
            option_code = option.get("option_code", "")
            row.set("label", f"r{option_code}")
            row.set("value", option_code)
            row.text = self.escape_xml_text(option.get("option_text", ""))
        
        return radio
    
    def generate_multiple_answers_xml(self, question):
        """生成Multiple Answers类型的XML"""
        checkbox = ET.Element("checkbox")
        checkbox.set("label", question.get("question_id", ""))
        checkbox.set("atleast", "1")
        
        # 添加title
        title = ET.SubElement(checkbox, "title")
        title.text = self.escape_xml_text(question.get("question_text", ""))
        
        # 添加comment
        comment = ET.SubElement(checkbox, "comment")
        comment.text = "${res.MRInst}"
        
        # 添加选项
        for option in question.get("question_options", []):
            row = ET.SubElement(checkbox, "row")
            option_code = option.get("option_code", "")
            row.set("label", f"r{option_code}")
            row.text = self.escape_xml_text(option.get("option_text", ""))
        
        return checkbox
    
    def generate_numeric_xml(self, question):
        """生成Numeric类型的XML"""
        number = ET.Element("number")
        number.set("label", question.get("question_id", ""))
        number.set("optional", "0")
        number.set("size", "20")
        number.set("verify", "range(0,99)")
        
        # 添加title
        title = ET.SubElement(number, "title")
        title.text = self.escape_xml_text(question.get("question_text", ""))
        
        return number
    
    def generate_xml(self, json_file_path, xml_file_path, verbose=True):
        """从JSON文件生成XML文件"""
        try:
            # 读取JSON文件
            with open(json_file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            questions = data.get("questions", [])
            print(f"开始处理 {len(questions)} 个问题...")
            
            # 按index排序确保顺序正确
            questions.sort(key=lambda x: x.get("index", 0))
            
            processed_count = {"Single Answer": 0, "Multiple Answers": 0, "Numeric": 0, "Unknown": 0}
            
            for question in questions:
                question_type = question.get("question_type", "")
                question_id = question.get("question_id", "")
                
                if verbose:
                    print(f"处理问题 {question.get('index', 'N/A')}: {question_id} ({question_type})")
                
                # 根据问题类型生成对应的XML元素
                if question_type == "Single Answer":
                    element = self.generate_single_answer_xml(question)
                    processed_count["Single Answer"] += 1
                elif question_type == "Multiple Answers":
                    element = self.generate_multiple_answers_xml(question)
                    processed_count["Multiple Answers"] += 1
                elif question_type == "Numeric":
                    element = self.generate_numeric_xml(question)
                    processed_count["Numeric"] += 1
                else:
                    print(f"  警告: 未知问题类型 '{question_type}'，跳过处理")
                    processed_count["Unknown"] += 1
                    continue
                
                # 添加到根元素
                self.root.append(element)
                
                # 添加suspend元素
                suspend = ET.Element("suspend")
                self.root.append(suspend)
            
            # 生成格式化的XML字符串
            xml_str = ET.tostring(self.root, encoding='unicode')
            
            # 使用minidom美化XML格式
            dom = minidom.parseString(xml_str)
            pretty_xml = dom.toprettyxml(indent="  ", encoding=None)
            
            # 移除空行
            lines = [line for line in pretty_xml.split('\n') if line.strip()]
            pretty_xml = '\n'.join(lines)
            
            # 写入XML文件
            with open(xml_file_path, 'w', encoding='utf-8') as f:
                f.write(pretty_xml)
            
            if verbose:
                print(f"\nXML生成完成！")
                print(f"输出文件: {xml_file_path}")
                print(f"处理统计:")
                for qtype, count in processed_count.items():
                    if count > 0:
                        print(f"  - {qtype}: {count}个")
            
            return True
            
        except Exception as e:
            print(f"生成XML时发生错误: {str(e)}")
            return False

def main():
    parser = argparse.ArgumentParser(description='将survey_output.json转换为XML格式')
    parser.add_argument('input_file', help='输入的JSON文件路径')
    parser.add_argument('output_file', help='输出的XML文件路径')
    
    args = parser.parse_args()
    
    generator = SurveyXMLGenerator()
    success = generator.generate_xml(args.input_file, args.output_file)
    
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()