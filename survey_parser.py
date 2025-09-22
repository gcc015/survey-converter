#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
问卷解析器 - 将Word文档中的问卷内容转换为结构化JSON格式
"""

import re
import json
import sys
import os
from datetime import datetime
from word_to_json import WordToJsonConverter

class SurveyParser:
    """问卷解析器类"""
    
    def __init__(self):
        self.word_converter = WordToJsonConverter()
        self.question_patterns = {
            'answer_logic': r'^(ASK\s+ALL|ASK\s+IF\s+.+)$',
            'question_number': r'^(\d+\.\d+[a-zA-Z]?)\s*(.+)$',  # 支持带字母后缀的题号
            'question_type': r'^(Single\s+Answer|Multiple\s+Answers?|Open\s+Answer|Numeric)\.?\s*.*$',
            'option': r'^(.+?)\s+(\d+)$',
            'question_text': r'^(.+?)(?:\s*\n.*)?$'  # 匹配问题文本，忽略换行后的内容
        }
    
    def parse_survey_document(self, word_file_path, output_file=None):
        """
        解析问卷Word文档
        
        Args:
            word_file_path (str): Word文档路径
            output_file (str, optional): 输出JSON文件路径
            
        Returns:
            dict: 解析结果
        """
        try:
            # 首先使用基础转换器读取Word文档
            word_content = self.word_converter.convert_to_json(word_file_path)
            
            if "error" in word_content:
                return word_content
            
            # 提取段落文本
            paragraphs = []
            for para in word_content["content"]["paragraphs"]:
                text = para["text"].strip()
                if text:  # 只保留非空段落
                    paragraphs.append(text)
            
            # 提取表格中的选项数据
            table_options = self._extract_table_options(word_content.get("content", {}).get("tables", []))
            
            # 解析问卷结构
            questions = self._parse_questions(paragraphs, table_options)
            
            # 构建结果
            result = {
                "file_info": word_content["file_info"],
                "survey_info": {
                    "total_questions": len(questions),
                    "processed_time": datetime.now().isoformat()
                },
                "questions": questions
            }
            
            # 保存到文件（如果指定）
            if output_file:
                with open(output_file, 'w', encoding='utf-8') as f:
                    json.dump(result, f, ensure_ascii=False, indent=2)
                print(f"问卷解析完成! JSON文件已保存到: {output_file}")
            
            return result
            
        except Exception as e:
            error_msg = f"解析问卷文档时发生错误: {str(e)}"
            print(error_msg)
            return {"error": error_msg}
    
    def _parse_questions(self, paragraphs, table_options):
        """
        解析问题列表
        
        Args:
            paragraphs (list): 段落文本列表
            table_options (list): 表格中的选项数据
            
        Returns:
            list: 问题列表
        """
        questions = []
        current_question = None
        question_index = 1
        table_index = 0
        
        i = 0
        while i < len(paragraphs):
            para = paragraphs[i].strip()
            
            # 检查是否是答题逻辑
            if self._is_answer_logic(para):
                if current_question:
                    # 在保存问题前，如果问题类型为空，进行全面搜索
                    if not current_question["question_type"]:
                        self._find_question_type_comprehensive(current_question, paragraphs)
                    questions.append(current_question)
                
                current_question = {
                    "index": question_index,
                    "answer_logic": para,
                    "question_id": "",
                    "question_text": "",
                    "question_type": "",
                    "question_options": []
                }
                question_index += 1
                
                # 查找下一个非空段落作为问题文本
                next_i = i + 1
                while next_i < len(paragraphs):
                    next_para = paragraphs[next_i].strip()
                    if next_para and not self._is_answer_logic(next_para):
                        # 检查是否是带编号的问题
                        match = re.match(self.question_patterns['question_number'], next_para)
                        if match:
                            question_num = match.group(1)
                            question_text = match.group(2).strip()
                            current_question["question_id"] = self._generate_question_id(question_num)
                            current_question["question_text"] = question_text
                            # 立即分配表格选项
                            self._assign_table_options(current_question, table_options)
                        else:
                            # 处理没有编号的问题文本
                            question_text = self._clean_question_text(next_para)
                            if question_text and not self._is_question_type(question_text):
                                current_question["question_text"] = question_text
                                # 为没有编号的问题生成ID，使用1.x格式
                                current_question["question_id"] = f"Q1x{current_question['index']}"
                                # 立即分配表格选项
                                self._assign_table_options(current_question, table_options)
                        
                        # 在找到问题文本后，继续查找问题类型（跳过可能的英文翻译）
                        type_i = next_i + 1
                        while type_i < len(paragraphs) and type_i < next_i + 3:  # 最多向前查找3行
                            type_para = paragraphs[type_i].strip()
                            if type_para and self._is_question_type(type_para):
                                type_match = re.match(self.question_patterns['question_type'], type_para, re.IGNORECASE)
                                if type_match:
                                    current_question["question_type"] = type_match.group(1)
                                break
                            type_i += 1
                        break
                    next_i += 1
            
            # 检查是否是问题类型
            elif self._is_question_type(para) and current_question:
                match = re.match(self.question_patterns['question_type'], para, re.IGNORECASE)
                if match:
                    current_question["question_type"] = match.group(1)  # 提取第一个捕获组
                
                # 根据问题内容分配正确的表格选项
                if current_question["question_text"]:
                    self._assign_table_options(current_question, table_options)
                    
                    # 特殊处理银行问题的Multiple类型
                    if "银行" in current_question["question_text"] and "Multiple" in para:
                        # 银行问题使用最后一个表格
                        if len(table_options) > 0:
                            current_question["question_options"] = table_options[-1]
            
            # 检查是否是选项（段落中的选项）
            elif self._is_option(para) and current_question:
                option = self._parse_option(para)
                if option:
                    current_question["question_options"].append(option)
            
            i += 1
        
        # 添加最后一个问题
        if current_question:
            # 在保存问题前，检查是否需要分配表格选项
            if not current_question["question_options"]:
                self._assign_table_options(current_question, table_options)
            
            # 如果问题类型仍然为空，进行全面搜索
            if not current_question["question_type"]:
                self._find_question_type_comprehensive(current_question, paragraphs)
            
            questions.append(current_question)
        
        return questions
    
    def _is_answer_logic(self, text):
        """检查是否是答题逻辑"""
        return bool(re.match(self.question_patterns['answer_logic'], text, re.IGNORECASE))
    
    def _is_question_text(self, text):
        """检查是否是问题文本"""
        return bool(re.match(self.question_patterns['question_number'], text))
    
    def _is_question_type(self, text):
        """检查是否是问题类型"""
        return bool(re.match(self.question_patterns['question_type'], text, re.IGNORECASE))
    
    def _is_option(self, text):
        """检查是否是选项"""
        return bool(re.match(self.question_patterns['option'], text))
    
    def _parse_option(self, text):
        """
        解析选项
        
        Args:
            text (str): 选项文本
            
        Returns:
            dict: 选项信息
        """
        match = re.match(self.question_patterns['option'], text)
        if match:
            option_text = match.group(1).strip()
            option_code = match.group(2).strip()
            return {
                "option_code": option_code,
                "option_text": option_text
            }
        return None
    
    def _extract_table_options(self, tables):
        """
        从表格中提取选项数据，支持多种表格格式：
        1. 两列格式：第1列option_text，第2列option_code
        2. 四列格式：第1列option_text，第2列option_code，第3列option_text，第4列option_code
        
        Args:
            tables (list): 表格数据列表
            
        Returns:
            list: 选项列表的列表
        """
        table_options = []
        
        for table in tables:
            if "data" in table and isinstance(table["data"], list):
                options = []
                for row in table["data"]:
                    if isinstance(row, list) and len(row) >= 4:
                        # 四列格式：处理第1、2列和第3、4列
                        # 第1列option_text，第2列option_code
                        option_text1 = row[0].get("text", "").strip()
                        option_code1 = row[1].get("text", "").strip()
                        
                        # 第3列option_text，第4列option_code
                        option_text2 = row[2].get("text", "").strip()
                        option_code2 = row[3].get("text", "").strip()
                        
                        # 添加第一对选项
                        if option_text1 and option_code1:
                            options.append({
                                "option_code": option_code1,
                                "option_text": option_text1
                            })
                        
                        # 添加第二对选项
                        if option_text2 and option_code2:
                            options.append({
                                "option_code": option_code2,
                                "option_text": option_text2
                            })
                            
                    elif isinstance(row, list) and len(row) >= 2:
                        # 两列格式：第1列option_text，第2列option_code
                        option_text = row[0].get("text", "").strip()
                        option_code = row[1].get("text", "").strip()
                        
                        # 跳过空行或无效数据
                        if option_text and option_code:
                            options.append({
                                "option_code": option_code,
                                "option_text": option_text
                            })
                    elif isinstance(row, list) and len(row) >= 1:
                        # 处理只有一列的情况（可能是标题行）
                        option_text = row[0].get("text", "").strip()
                        if option_text and not option_text.isdigit():
                            # 如果只有文本没有代码，自动生成代码
                            option_code = str(len(options) + 1)
                            options.append({
                                "option_code": option_code,
                                "option_text": option_text
                            })
                
                if options:
                    table_options.append(options)
        
        return table_options
    
    def _assign_table_options(self, question, table_options):
        """为问题分配对应的表格选项"""
        if not question["question_text"]:
            return
            
        # 根据问题内容关键词分配表格选项
        question_text = question["question_text"].lower()
        
        if "性别" in question["question_text"]:
            # 性别问题使用第一个表格
            if len(table_options) > 0:
                question["question_options"] = table_options[0]
        elif "年龄" in question["question_text"]:
            # 年龄问题使用第二个表格
            if len(table_options) > 1:
                question["question_options"] = table_options[1]
        elif "居留身份" in question["question_text"]:
            # 居留身份问题使用第三个表格
            if len(table_options) > 2:
                question["question_options"] = table_options[2]
        elif "信用卡" in question["question_text"] and "持有任何" in question["question_text"]:
            # 信用卡持有问题使用第四个表格
            if len(table_options) > 3:
                question["question_options"] = table_options[3]
        elif "职业" in question["question_text"]:
            # 职业问题，根据问题索引选择表格
            if question["index"] == 5 and len(table_options) > 4:
                question["question_options"] = table_options[4]
            elif question["index"] == 6 and len(table_options) > 5:
                question["question_options"] = table_options[5]
        elif "年收入" in question["question_text"]:
            # 年收入问题使用第七个表格
            if len(table_options) > 6:
                question["question_options"] = table_options[6]
        elif "银行" in question["question_text"]:
            # 银行问题使用最后一个表格
            if len(table_options) > 0:
                question["question_options"] = table_options[-1]
        elif "信用卡" in question["question_text"] and "持有哪些" in question["question_text"]:
            # 信用卡类型问题，查找包含信用卡名称的表格
            for i, options in enumerate(table_options):
                if options and any("信用卡" in opt.get("option_text", "") or "Mastercard" in opt.get("option_text", "") or "Visa" in opt.get("option_text", "") for opt in options):
                    question["question_options"] = options
                    break
        elif "mmpower" in question_text or "mastercard" in question_text:
            # MMPOWER相关问题
            if "何时取得" in question["question_text"] or "when" in question_text:
                # Q1x8问题：何时取得MMPOWER卡，使用时间选项表格（索引9）
                if len(table_options) > 9:
                    question["question_options"] = table_options[9]
            else:
                # 其他MMPOWER相关问题，使用原有逻辑
                table_index = question["index"] - 2
                if 0 <= table_index < len(table_options):
                    question["question_options"] = table_options[table_index]
        elif "消费" in question["question_text"] or "开销" in question["question_text"]:
            # 消费相关问题，需要找到对应的消费金额选项表格
            # 通常消费问题会有自己的选项表格，不是信用卡表格
            for i, options in enumerate(table_options):
                if options and any("港币" in opt.get("option_text", "") or "元" in opt.get("option_text", "") for opt in options):
                    # 查找包含金额的表格
                    if not any("银行" in opt.get("option_text", "") or "信用卡" in opt.get("option_text", "") for opt in options):
                        question["question_options"] = options
                        break
        else:
            # 通用回退机制 - 按问题索引分配对应的表格
            # 问题索引从1开始，表格索引从0开始
            table_index = question["index"] - 1
            if 0 <= table_index < len(table_options) and table_options[table_index]:
                question["question_options"] = table_options[table_index]
            else:
                # 如果对应索引的表格为空或不存在，尝试找到最近的非空表格
                # 先向后查找
                for i in range(table_index + 1, len(table_options)):
                    if table_options[i]:
                        question["question_options"] = table_options[i]
                        break
                else:
                    # 如果向后没找到，向前查找
                    for i in range(table_index - 1, -1, -1):
                        if table_options[i]:
                            question["question_options"] = table_options[i]
                            break
    
    def _clean_question_text(self, text):
        """
        清理问题文本，移除多余的换行和格式
        
        Args:
            text (str): 原始文本
            
        Returns:
            str: 清理后的问题文本
        """
        # 分割多行文本，只取第一行作为问题文本
        lines = text.split('\n')
        question_text = lines[0].strip()
        
        # 移除问题类型信息
        if re.match(self.question_patterns['question_type'], question_text, re.IGNORECASE):
            return ""
        
        # 移除答题逻辑信息
        if re.match(self.question_patterns['answer_logic'], question_text, re.IGNORECASE):
            return ""
        
        return question_text
    
    def _generate_question_id(self, question_num):
        """
        生成问题ID
        
        Args:
            question_num (str): 问题编号 (如 "1.1", "1.4a", "1.4b")
            
        Returns:
            str: 问题ID (如 "Q1x1", "Q1x4a", "Q1x4b")
        """
        parts = question_num.split('.')
        if len(parts) == 2:
            return f"Q{parts[0]}x{parts[1]}"
        return f"Q{question_num}"
    
    def _find_question_type_comprehensive(self, question, paragraphs):
        """
        全面搜索问题类型
        
        Args:
            question (dict): 当前问题对象
            paragraphs (list): 所有段落列表
        """
        question_text = question.get("question_text", "")
        if not question_text:
            return
        
        # 策略1: 在整个文档中搜索包含问题文本的段落附近的问题类型
        for i, para in enumerate(paragraphs):
            if question_text in para:
                # 向前和向后搜索问题类型（扩大搜索范围到10行）
                search_start = max(0, i - 5)
                search_end = min(len(paragraphs), i + 10)
                
                for j in range(search_start, search_end):
                    search_para = paragraphs[j].strip()
                    if search_para and self._is_question_type(search_para):
                        type_match = re.match(self.question_patterns['question_type'], search_para, re.IGNORECASE)
                        if type_match:
                            question["question_type"] = type_match.group(1)
                            return
                break
        
        # 策略2: 基于问题内容和选项推断问题类型
        self._infer_question_type_from_content(question)
    
    def _infer_question_type_from_content(self, question):
        """
        基于问题内容和选项推断问题类型
        
        Args:
            question (dict): 当前问题对象
        """
        question_text = question.get("question_text", "").lower()
        options = question.get("question_options", [])
        
        # 如果问题文本包含特定关键词，推断类型
        if any(keyword in question_text for keyword in ["年龄", "收入", "金额", "数量", "百分比", "比例"]):
            question["question_type"] = "Numeric"
        elif any(keyword in question_text for keyword in ["选择所有", "多选", "哪些", "哪三个", "前三个"]):
            question["question_type"] = "Multiple Answers"
        elif len(options) == 2 and any("男" in opt.get("option_text", "") or "女" in opt.get("option_text", "") for opt in options):
            question["question_type"] = "Single Answer"
        elif len(options) > 2:
            # 检查选项内容，如果都是具体选择项，推断为Single Answer
            if all(opt.get("option_text", "").strip() for opt in options):
                question["question_type"] = "Single Answer"

    def parse_single_question_example(self):
        """
        返回单个问题的示例格式
        
        Returns:
            dict: 示例问题格式
        """
        return {
            "index": 1,
            "answer_logic": "ASK ALL",
            "question_id": "Q1x1",
            "question_text": "您的性别是?",
            "question_type": "Single Answer",
            "question_options": [
                {
                    "option_code": "1",
                    "option_text": "男"
                },
                {
                    "option_code": "2",
                    "option_text": "女"
                }
            ]
        }

def main():
    """主函数 - 命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="将Word文档中的问卷内容转换为结构化JSON格式",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法:
  python survey_parser.py survey.docx                     # 输出到控制台
  python survey_parser.py survey.docx survey_output.json  # 保存到文件
  python survey_parser.py --example                       # 显示输出格式示例
        """
    )
    
    parser.add_argument(
        'word_file',
        nargs='?',
        help='要解析的问卷Word文件路径 (.docx 或 .doc)'
    )
    
    parser.add_argument(
        'output_file',
        nargs='?',
        help='输出JSON文件路径 (可选，不指定则输出到控制台)'
    )
    
    parser.add_argument(
        '--example',
        action='store_true',
        help='显示问题格式示例'
    )
    
    parser.add_argument(
        '--version',
        action='version',
        version='Survey Parser 1.0.0'
    )
    
    args = parser.parse_args()
    
    if args.example:
        parser_instance = SurveyParser()
        example = parser_instance.parse_single_question_example()
        print("问题格式示例已生成")
        return
    
    if not args.word_file:
        parser.print_help()
        return
    
    print(f"正在解析问卷文档: {args.word_file}")
    
    parser_instance = SurveyParser()
    result = parser_instance.parse_survey_document(args.word_file, args.output_file)
    
    if "error" in result:
        print(f"错误: {result['error']}")
        sys.exit(1)
    else:
        if not args.output_file:
            print("解析完成! JSON内容已生成")

if __name__ == "__main__":
    main()