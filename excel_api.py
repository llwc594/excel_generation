import json
import zipfile

from io import BytesIO

from flask import Flask, request, jsonify, send_file

import math

from collections import defaultdict, OrderedDict
from itertools import combinations
import pandas as pd
import datetime
import os
import shutil
from pathlib import Path
from typing import Any, List, Dict, Union
import openpyxl
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook import Workbook
from typing import Dict, List, Any, Union, Tuple, Optional
import xlwings as xw
from openpyxl.styles import Alignment, Font, Border, Side

from  api_8 import main as main_8
from  api_24 import main as main_24
from  api_48 import main as main_48
from  api_96 import main as main_96

# 常量定义
BASE_PATH = r"./generation_zy"
BASE_PATH1 = r"./逆合成"
crystallize_BASE_PATH = r"/192.168.11.80/立项组/立项组/李伟龙/接口/表格生成/结晶平台"
crystallize_BASE_PATH1 = r"/192.168.11.80/立项组/立项组/李伟龙/接口/模板/结晶平台"

TEMPLATE_PATHS = {
    "wf11": f"{BASE_PATH}/hiwo-任务2.xlsx",
    "wf05": f"{BASE_PATH}/hiwo-任务1.xlsx",
    "sowo": f"{BASE_PATH}/sowo-物料1.xlsx",
    "mate": f"{BASE_PATH}/sowo-任务1.xlsx",
    "information": f"{BASE_PATH}/信息表格.xlsx",
    "mate_info1": f"{BASE_PATH}/hiwo-物料1.xlsx",
    "mate_info2": f"{BASE_PATH}/hiwo-物料2.xlsx",

}
crystallize_TEMPLATE_PATHS = {
    "wf11": f"{crystallize_BASE_PATH}/hiwo-任务1-1.xlsx",
    "wf05": f"{crystallize_BASE_PATH}/hiwo-任务1-2.xlsx",
    "sowo": f"{crystallize_BASE_PATH}/sowo-物料1.xlsx",
    "mate": f"{crystallize_BASE_PATH}/sowo-任务1.xlsx",
    "mate_info1": f"{crystallize_BASE_PATH}/hiwo-物料1-1.xlsx",
    "mate_info2": f"{crystallize_BASE_PATH}/hiwo-物料1-2.xlsx",
    "information":f"{crystallize_BASE_PATH}/信息表格.xlsx",
    "WF10":f"{crystallize_BASE_PATH1}/hiwo任务15-8-24.xlsx",
    "WF11": f"{crystallize_BASE_PATH1}/hiwo物料-8-01.xlsx"
}
# 全局状态


class convert_format():

    def __init__(self):
        self.app = Flask(__name__)
        self._register_routes()
    def _register_routes(self):
        @self.app.route('/api/synthesis_excel', methods=['POST'])
        def handle_data():
            if request.method == 'POST':
                try:

                    data = request.get_json()

                    json_name = '合成平台设置参数.json'
                    json_path = os.path.join(BASE_PATH, json_name)
                    with open(json_path, 'w', encoding='utf-8') as f:
                        json.dump(data, f, ensure_ascii=False)
                    id,condition = data["id"],data["condition"]

                    if condition["experimentLog1"]["wellPlates"] =="8孔板20ml":

                        main_8(id,condition,1)
                        memory_file = BytesIO()
                        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for root, dirs, files in os.walk(BASE_PATH):  # 假设Excel文件在output目录
                                for file in files:
                                    if file.endswith(('.xlsx','.json')):
                                        zf.write(os.path.join(root, file), file)
                        memory_file.seek(0)
                        return send_file(
                            memory_file,
                            mimetype='application/zip',
                            as_attachment=True,
                            download_name=f'{id}.zip')
                    elif condition["experimentLog1"]["wellPlates"] =="24孔板8ml" or condition["experimentLog1"]["wellPlates"] =="24孔板2ml" or condition["experimentLog1"]["wellPlates"] =="24孔板4ml":

                        main_24(id,condition,1)
                        memory_file = BytesIO()
                        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for root, dirs, files in os.walk(BASE_PATH):  # 假设Excel文件在output目录
                                for file in files:
                                    if file.endswith(('.xlsx','.json')):
                                        zf.write(os.path.join(root, file), file)
                        memory_file.seek(0)
                        return send_file(
                            memory_file,
                            mimetype='application/zip',
                            as_attachment=True,
                            download_name=f'{id}.zip')
                    elif condition["experimentLog1"]["wellPlates"] =="48孔板1ml温":

                        main_48(id,condition,1)
                        memory_file = BytesIO()
                        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for root, dirs, files in os.walk(BASE_PATH):  # 假设Excel文件在output目录
                                for file in files:
                                    if file.endswith(('.xlsx','.json')):
                                        zf.write(os.path.join(root, file), file)
                        memory_file.seek(0)
                        return send_file(
                            memory_file,
                            mimetype='application/zip',
                            as_attachment=True,
                            download_name=f'{id}.zip')
                    elif condition["experimentLog1"]["wellPlates"] =="96孔板1ml":

                        main_96(id,condition,1)
                        memory_file = BytesIO()
                        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for root, dirs, files in os.walk(BASE_PATH):  # 假设Excel文件在output目录
                                for file in files:
                                    if file.endswith(('.xlsx','.json')):
                                        zf.write(os.path.join(root, file), file)
                        memory_file.seek(0)
                        return send_file(
                            memory_file,
                            mimetype='application/zip',
                            as_attachment=True,
                            download_name=f'{id}.zip')

                    else:
                        return "仅支持8孔20cm以及24孔8cm"
                except Exception as e:
                    # 捕获异常并返回错误响应
                    return jsonify({'error': str(e)}), 500

        @self.app.route('/api/crystallize_excel', methods=['POST'])
        def crystallize_excel():
            if request.method == 'POST':
                try:
                    data = request.get_json()
                    json_name = '结晶平台设置参数.json'
                    json_path = os.path.join(crystallize_BASE_PATH, json_name)
                    with open(json_path, 'w', encoding='utf-8') as f:
                        json.dump(data, f, ensure_ascii=False)
                    id = data["id"]
                    condition = data["condition"]
                    if condition["experimentLog1"]["wellPlates"] =="8孔板20ml":

                        # 执行主程序
                        main_8(id,condition,2)
                        # 创建内存中的ZIP文件
                        memory_file = BytesIO()
                        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
                            # 遍历生成的Excel文件并添加到ZIP
                            for root, dirs, files in os.walk(crystallize_BASE_PATH):  # 假设Excel文件在output目录
                                for file in files:
                                    if file.endswith(('.xlsx','.json')):
                                        zf.write(os.path.join(root, file), file)

                        # 重置文件指针
                        memory_file.seek(0)

                        # 返回ZIP文件
                        return send_file(
                            memory_file,
                            mimetype='application/zip',
                            as_attachment=True,
                            download_name=f'{id}.zip'
                        )
                    elif condition["experimentLog1"]["wellPlates"] =="24孔板8ml" or condition["experimentLog1"]["wellPlates"] =="24孔板2ml" or condition["experimentLog1"]["wellPlates"] =="24孔板4ml":

                        main_24(id,condition,2)
                        # 创建内存中的ZIP文件
                        memory_file = BytesIO()
                        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
                            # 遍历生成的Excel文件并添加到ZIP
                            for root, dirs, files in os.walk(crystallize_BASE_PATH):  # 假设Excel文件在output目录
                                for file in files:
                                    if file.endswith(('.xlsx','.json')):
                                        zf.write(os.path.join(root, file), file)

                        # 重置文件指针
                        memory_file.seek(0)

                        # 返回ZIP文件
                        return send_file(
                            memory_file,
                            mimetype='application/zip',
                            as_attachment=True,
                            download_name=f'{id}.zip'
                        )
                    elif condition["experimentLog1"]["wellPlates"] == "48孔板1ml温":

                        main_48(id,condition, 2)
                        # 创建内存中的ZIP文件
                        memory_file = BytesIO()
                        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
                            # 遍历生成的Excel文件并添加到ZIP
                            for root, dirs, files in os.walk(crystallize_BASE_PATH):  # 假设Excel文件在output目录
                                for file in files:
                                    if file.endswith(('.xlsx', '.json')):
                                        zf.write(os.path.join(root, file), file)

                        # 重置文件指针
                        memory_file.seek(0)

                        # 返回ZIP文件
                        return send_file(
                            memory_file,
                            mimetype='application/zip',
                            as_attachment=True,
                            download_name=f'{id}.zip'
                        )
                    elif condition["experimentLog1"]["wellPlates"] == "96孔板1ml":

                        main_96(id,condition, 2)
                        # 创建内存中的ZIP文件
                        memory_file = BytesIO()
                        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
                            # 遍历生成的Excel文件并添加到ZIP
                            for root, dirs, files in os.walk(crystallize_BASE_PATH):  # 假设Excel文件在output目录
                                for file in files:
                                    if file.endswith(('.xlsx', '.json')):
                                        zf.write(os.path.join(root, file), file)

                        # 重置文件指针
                        memory_file.seek(0)

                        # 返回ZIP文件
                        return send_file(
                            memory_file,
                            mimetype='application/zip',
                            as_attachment=True,
                            download_name=f'{id}.zip'
                        )

                    else:
                        return "仅支持8孔20cm以及24孔8cm"

                except Exception as e:
                    # 捕获异常并返回错误响应
                    return jsonify({'error': str(e)}), 500


if __name__ =="__main__":
    converter = convert_format()
    # converter.app.run(debug=True, host='0.0.0.0', port=9259)
    converter.app.run(debug=True, host='0.0.0.0', port=10014)
    # converter.main()


