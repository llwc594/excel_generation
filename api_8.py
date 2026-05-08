import json
import zipfile

from io import BytesIO
from openpyxl.styles import PatternFill

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

import subprocess
import os
import time

import os

from pathlib import Path
import aspose.cells as ac


from  api_24 import main as main_24
from  api_48 import main as main_48
from  api_96 import main as main_96
# BASE_PATH = r"D:/api/python/表格生成/逆合成"
# BASE_PATH1 = r"D:/api/python/模板/逆合成"
# crystallize_BASE_PATH = r"D:/api/python/表格生成/结晶平台"
# crystallize_BASE_PATH1 = r"D:/api/python/模板/结晶平台"
# # 常量定义
# BASE_PATH = r"C:/Users/24017/PycharmProjects/pythonProject/接口测试/表格生成/逆合成"
# BASE_PATH1 = r"C:/Users/24017/PycharmProjects/pythonProject/接口测试/模板/逆合成"
# crystallize_BASE_PATH = r"C:/Users/24017/PycharmProjects/pythonProject/接口测试/表格生成/结晶平台"
# crystallize_BASE_PATH1 = r"C:/Users/24017/PycharmProjects/pythonProject/接口测试/模板/结晶平台"

BASE_PATH = r"./generation_zy"
BASE_PATH1 = r"./逆合成"
crystallize_BASE_PATH = r"C:/Users/24017/PycharmProjects/pythonProject/接口测试/表格生成/结晶平台"
crystallize_BASE_PATH1 = r"C:/Users/24017/PycharmProjects/pythonProject/接口测试/模板/结晶平台"
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

class ExcelUtils:
    """Excel文件操作工具类"""

    @staticmethod
    def open_workbook(file_path: str) -> Workbook:
        """打开Excel文件"""
        return openpyxl.load_workbook(file_path)

    @staticmethod
    def save_workbook(workbook: Workbook, save_path: str) -> None:
        """保存Excel文件"""
        Path(save_path).parent.mkdir(parents=True, exist_ok=True)
        workbook.save(save_path)

    @staticmethod
    def delete_file(file_path: str) -> bool:
        """删除指定路径的文件

        Args:
            file_path: 要删除的文件路径

        Returns:
            bool: 是否删除成功（True表示成功，False表示失败）
        """
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                return True
            return False
        except Exception as e:
            print(f"删除文件失败: {e}")
            return False



    @staticmethod
    def copy_workbook(source_path: str, target_path: str) -> None:
        """复制Excel文件"""
        wb = openpyxl.load_workbook(source_path)
        ExcelUtils.save_workbook(wb, target_path)
        print(f"文件已成功复制到: {target_path}")

    @staticmethod
    def modify_cell(workbook: Workbook, sheet_name: str, cell_ref: Union[str, tuple], new_value: Any) -> None:
        """修改单元格值"""

        sheet = workbook[sheet_name]
        if isinstance(cell_ref, str):
            if new_value is None or new_value == "":
                sheet[cell_ref].value = "空"
            else:
                sheet[cell_ref].value = new_value
        else:
            if new_value is None or new_value == "":
                sheet.cell(row=cell_ref[0], column=cell_ref[1]).value = "空"
            else:
                sheet.cell(row=cell_ref[0], column=cell_ref[1]).value = new_value


    @staticmethod
    def insert_row(workbook: Workbook, sheet_name: str, row_num: int, style: Any = None) -> None:
        """插入新行"""
        sheet = workbook[sheet_name]
        sheet.insert_rows(row_num)

        if style:
            for col in range(1, sheet.max_column + 1):
                sheet.cell(row=row_num, column=col)._style = style

    @staticmethod
    def delete_row(workbook: Workbook, sheet_name: str, row_num: int) -> None:
        """删除行"""
        sheet = workbook[sheet_name]
        for i in range(2, row_num):
            sheet.delete_rows(2)

    @staticmethod
    def modify_row(workbook: Workbook, sheet_name: str, row_num: int,
                   new_values: Union[list, dict], start_col: int = 1) -> None:
        """修改整行数据"""
        sheet = workbook[sheet_name]

        if isinstance(new_values, dict):
            for col, value in new_values.items():
                if isinstance(col, str):
                    sheet[f"{col}{row_num}"].value = value
                else:sheet.cell(row=row_num, column=col).value = value
        else:
            for i, value in enumerate(new_values, start=start_col):
                sheet.cell(row=row_num, column=i).value = value
    @staticmethod
    def write_records_to_sheet(workbook: Workbook, sheet_name: str, records: List[dict]) -> None:
        """将记录写入工作表"""
        sheet = workbook[sheet_name]

        # 写入数据
        for record in records:
            row = list(record.values())
            sheet.append(row)

    @staticmethod
    def clear_directory(folder_path: str) -> None:
        """
        清空指定文件夹内的所有文件和子文件夹

        参数:
            folder_path: 要清空的文件夹路径

        异常:
            FileNotFoundError: 当文件夹不存在时抛出
            PermissionError: 当没有权限时抛出
        """
        if not os.path.exists(folder_path):
            raise FileNotFoundError(f"文件夹不存在: {folder_path}")

        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f"删除失败 {file_path}. 原因: {e}")

    @staticmethod
    def clear_directory_(folder_path: str) -> None:
        """
        清空指定文件夹内的所有Excel文件

        参数:
            folder_path: 要清空的文件夹路径

        异常:
            FileNotFoundError: 当文件夹不存在时抛出
            PermissionError: 当没有权限时抛出
        """
        if not os.path.exists(folder_path):
            raise FileNotFoundError(f"文件夹不存在: {folder_path}")

        # Excel文件扩展名列表
        excel_extensions = {'.xlsx', '.xls', '.xlsm', '.xlsb', '.xltx', '.xlt', '.xltm'}

        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            try:
                # 只删除Excel文件，保留其他文件和文件夹
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    # 检查文件扩展名是否为Excel格式
                    file_ext = os.path.splitext(filename)[1].lower()
                    if file_ext in excel_extensions:
                        os.unlink(file_path)
                        print(f"已删除Excel文件: {filename}")
                # 注意：不再删除文件夹，只处理文件
            except Exception as e:
                print(f"删除失败 {file_path}. 原因: {e}")
    @staticmethod
    def batch_repair_excels(folder_path):
        """
        批量修复文件夹内所有Excel文件（直接覆盖原文件）
        参数：
            folder_path: 包含Excel文件的文件夹路径
        返回：
            总耗时(秒)和成功处理文件数
        """
        app = None
        try:
            app = xw.App(visible=False)
            files = [f for f in os.listdir(folder_path)
                     if f.lower().endswith(('.xlsx', '.xls', '.xlsm', '.xlsb'))]
            for filename in files:
                try:
                    file_path = os.path.join(folder_path, filename)
                    wb = app.books.open(file_path)
                    # 直接保存覆盖原文件
                    wb.save()
                    wb.close()
                except Exception as e:
                    print(f"处理失败 {filename}: {str(e)}")
        finally:
            if app is not None: app.quit()

    @staticmethod
    def batch_repair_folder(folder_path, backup=False, extensions=('.xlsx', '.xls', '.xlsm', '.xlsb')):
        def repair_excel(file_path, backup=False):
            """
            修复单个 Excel 文件
            :param file_path: 文件路径
            :param backup: 是否备份原文件（备份为 原文件名.bak）
            :return: 成功返回 True，失败返回 False
            """
            try:
                wb = ac.Workbook(file_path)
                if backup:
                    backup_path = file_path + ".bak"
                    os.rename(file_path, backup_path)
                    print(f"已备份: {backup_path}")
                wb.save(file_path)
                print(f"修复成功: {file_path}")
                return True
            except Exception as e:
                print(f"修复失败: {file_path}, 错误: {e}")
                return False
        """
        批量修复文件夹内所有 Excel 文件
        :param folder_path: 文件夹路径
        :param backup: 是否备份原文件
        :param extensions: 要处理的文件扩展名元组
        """
        if not os.path.isdir(folder_path):
            print(f"错误: {folder_path} 不是有效文件夹")
            return

        files = [f for f in os.listdir(folder_path) 
                if "任务" in f and f.lower().endswith(('.xlsx', '.xls', '.xlsm', '.xlsb'))]
        
        if not files:
            print("未找到 Excel 文件")
            return

        success_count = 0
        for filename in files:
            full_path = os.path.join(folder_path, filename)
            if repair_excel(full_path, backup):
                success_count += 1

        print(f"批量修复完成: 成功 {success_count}/{len(files)}")


    
    
class ReactionProcessor:
    """反应处理核心类"""

    def __init__(self, id,post_data: dict):
        self.post_data = post_data
        self.traceid = id
        self.data_dict = {}
        self.theoretical_volume = {}
        self.solidity = {}
        self.perforated_plate = []
        self.reagent_plates_12 = {}
        self.reagent_plates_12_ = {}
        self.reagent_plates_24 = []
        self.data_list_24 = []
        self.data_list_12 = []
        # 12孔板使用情况
        self.usage_12_hole = 13
        # 判断时候需要第二次任务
        self.number_tasks = 2
        # 八孔板一共需要使用几个孔列
        self.hole_number = 0
        # 记录sowo使用的物料对应的key
        self.name_key_dict = {}
        # self.manage_daily_file(BASE_PATH1)
        # self.manage_daily_file(crystallize_BASE_PATH1)
        # 结晶平台的反溶清
        self.antisolvent = {}
        # 作为判断是逆合成还是结晶平台
        self.type = 0
        # 12孔板占用24孔板次数
        self.usage_24_hole = 0
        self.gun_head_dict = defaultdict(list)

    def process(self, source) -> None:
        """主处理流程"""

        self.type = source
        """处理数据"""
        self._initialize_data()
        """计算反应相似度并分组"""
        self._calculate_similarity()

        self.hole_number = len(self.perforated_plate)
        """生成所有Excel文件"""
        self._generate_excel_files()
        if self.type == 1:

            if self.number_tasks == 1:
                """直接在这里判断如果只有一个文件就打开文件2保存为文件1，"""
                wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info2"])
                ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info1"])
                wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf11"])
                ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf05"])

                ExcelUtils.delete_file(TEMPLATE_PATHS["mate_info2"])
                ExcelUtils.delete_file(TEMPLATE_PATHS["wf11"])

                self._create_excel_xxx(TEMPLATE_PATHS["mate_info1"], TEMPLATE_PATHS["information"], 1)
                #sowo
                self._create_excel_ttt(26)
                #添加空位信息
                self._create_excel_information(26, 7, TEMPLATE_PATHS["information"])
            else:
                self._create_excel_xxx(TEMPLATE_PATHS["mate_info1"], TEMPLATE_PATHS["information"], 1)
                self._create_excel_xxx(TEMPLATE_PATHS["mate_info2"], TEMPLATE_PATHS["information"], 28)
                self._create_excel_ttt(60)
                self._create_excel_information(60, 7, TEMPLATE_PATHS["information"])
            #添加物料表
            self._create_excel_material(TEMPLATE_PATHS["information"])

        elif self.type == 2:

            self._create_excel_xxx(crystallize_TEMPLATE_PATHS["mate_info1"], crystallize_TEMPLATE_PATHS["information"],
                                   1)
            self._create_excel_tttt(26, crystallize_TEMPLATE_PATHS['sowo'], crystallize_TEMPLATE_PATHS['information'])
            self._create_excel_information(26, 7, crystallize_TEMPLATE_PATHS["information"])
            self._create_excel_material(crystallize_TEMPLATE_PATHS["information"])
            method = self.post_data['experimentLog2'].get('postProcessingForm', []).get('method', '0')
            if method == "2":
                self._create_excel_xxxx(fr"{crystallize_BASE_PATH}/hiwo-物料3-1.xlsx",
                                        crystallize_TEMPLATE_PATHS["information"], 2)
            if method == "1":
                self._create_excel_xxxx(fr"{crystallize_BASE_PATH}/hiwo-物料3-1.xlsx",
                                        crystallize_TEMPLATE_PATHS["information"], 2)

    def _create_excel_information(self, start_row, start_col, input_path):

        wb = ExcelUtils.open_workbook(input_path)
        ws = wb['Sheet']
        # 指定起始行列（例如从第3行第2列开始写入）

        for col_idx, letter in enumerate("AB", start=start_col):
            ws.cell(row=start_row, column=col_idx, value=letter).font = Font(bold=True)

        # 添加行标签 (1,2,3...)
        for row_idx in range(start_row + 1, start_row + 5):
            ws.cell(row=row_idx, column=start_col - 1, value=(start_row + 5) - row_idx).font = Font(bold=True)
        thin_border = Border(left=Side(style='thin'),
                             right=Side(style='thin'),
                             top=Side(style='thin'),
                             bottom=Side(style='thin'))

        # 写入数据
        for row_idx, row_data in enumerate(self.perforated_plate, start=1):
            for col_idx, cell_value in enumerate(row_data, start=start_col):
                cell = ws.cell(row=start_row + 5 - row_idx, column=col_idx, value=cell_value)
                cell.border = thin_border
                cell.font = Font(bold=True)

        ExcelUtils.save_workbook(wb, input_path)

    def _create_excel_material(self, input_path):
        def safe_float_convert(value, default=0.0):
            """安全地将值转换为浮点数，处理空字符串或无效输入"""
            try:
                return float(value) if value != '' else default
            except (ValueError, TypeError):
                return default

        book = None
        ws = None
        if os.path.exists(input_path):
            book = load_workbook(input_path)
            if "试剂数据" in book.sheetnames:
                ws = book["试剂数据"]
            else:
                ws = book.create_sheet("试剂数据")
        else:
            book = Workbook()
            ws = book.create_sheet("试剂数据")

        # 获取表头并写入
        headers = [
            "分子#", "物质名称", "分子式", "CAS", "SMILES", "分子量(g/mol)",
            "理论摩尔数(mMol)", "理论质量(mg)", "理论体积(mL)", "密度(g/mL)",
            "配置反应摩尔数(mMol)", "液体试剂配置质量(mg)", "配置体积(mL)", "浓度(mMol/mL)"
        ]
        for col_idx, header in enumerate(headers, 1):
            ws.cell(row=1, column=col_idx, value=header)

        # 遍历数据并写入
        for row, data in enumerate(self.post_data['experimentLog1']['MaterialsTableList'], 2):
            if self.type == 1:
                theoretical_moles = safe_float_convert(data["theoreticalMoles"])
                theoretical_quality = safe_float_convert(data["theoreticalQuality"])
                theoretical_volume = safe_float_convert(data["theoreticalVolume"])
                configuration_moles = safe_float_convert(data["configurationMoles"])
                configuration_quality = safe_float_convert(data["configurationQuality"])
                configuration_volume = safe_float_convert(data["configurationVolume"])
                concentration = safe_float_convert(data["concentration"])
                smiles = data["smiles"]
            else:
                theoretical_moles = safe_float_convert(data["theoreticalMoles"])
                theoretical_quality = safe_float_convert(data["theoreticalQuality"])
                theoretical_volume = safe_float_convert(data["theoreticalVolume"])
                configuration_moles = safe_float_convert(data["theoreticalMoles"])
                configuration_quality = safe_float_convert(data["theoreticalQuality"])
                configuration_volume = safe_float_convert(data["theoreticalVolume"])
                concentration = safe_float_convert(data["concentration"])
                smiles = data["smile"]

            # 计算密度
            density = ""
            if theoretical_volume > 0:
                density = round((theoretical_quality / 1000) / theoretical_volume, 4)

            row_data = {
                "分子#": data["id"],
                "物质名称": data["substance"],
                "分子式": data["molecularFormula"],
                "CAS": data["cas"] if data["cas"] else "",
                "SMILES": smiles,
                "分子量(g/mol)": data.get("molecularWeight", ""),
                # 注意：这里假设分子量字段可能是 "molecularWeight"，原代码使用了 "solvent" 可能是错误的
                "理论摩尔数(mMol)": theoretical_moles,
                "理论质量(mg)": theoretical_quality,
                "理论体积(mL)": theoretical_volume,
                "密度(g/mL)": density,
                "配置反应摩尔数(mMol)": configuration_moles,
                "液体试剂配置质量(mg)": configuration_quality,
                "配置体积(mL)": configuration_volume,
                "浓度(mMol/mL)": concentration
            }

            # 写入数据行
            for col_idx, col_name in enumerate(headers, 1):
                value = row_data.get(col_name, "")
                ws.cell(row=row, column=col_idx, value=value)

        # 保存工作簿（只在所有数据写入完成后保存一次）
        book.save(input_path)

    def manage_daily_file(self, path):
        # 获取当前日期并格式化文件名
        today = datetime.date.today()
        today_file = f"{path}/{today.strftime('%Y%m%d')}.txt"

        # 获取前一天的日期
        yesterday = today - datetime.timedelta(days=1)
        yesterday_file = f"{path}/{yesterday.strftime('%Y%m%d')}.txt"

        # 遍历目录下所有文件
        for filename in os.listdir(path):
            if filename.endswith(".txt") and filename != os.path.basename(today_file):
                file_path = os.path.join(path, filename)
                try:
                    os.remove(file_path)
                    print(f"已删除日志文件: {file_path}")
                except Exception as e:
                    print(f"删除文件{file_path}时出错: {e}")

        # 检查当天文件是否存在
        if not os.path.exists(today_file):
            # 创建新文件
            with open(today_file, 'w') as f:
                f.write(f"日志文件创建时间: {datetime.datetime.now()}/n")
            print(f"已创建新的日志文件: {today_file}")
        else:
            print(f"今日日志文件已存在: {today_file}")

    def _initialize_data(self) -> None:
        """初始化数据"""
        """提取底物数据"""
        self._extract_substrates()
        """提取物料数据"""
        self._extract_materials()

    def _extract_substrates(self) -> None:
        """提取底物数据"""
        counter = 1
        substrates = self.post_data["experimentLog1"]["SubstratesTableList"]
        """当只有一个孔的时候，直接添加孔2的信息"""
        # if all(sub['id'] == 1 for sub in substrates):
        #     new_subs = []
        #     for sub in substrates:
        #         new_sub = sub.copy()
        #         new_sub['id'] = 2
        #         new_subs.append(new_sub)
        #     substrates.extend(new_subs)
        for item in substrates:
            substance = item["substance"].strip()
            hole_id = item["id"]
            if self.type == 1:
                if item["liquidReagentConcentration"] == "液体":
                    volume_str = item.get("singleCockAddVolume", '0')
                    volume = float(volume_str) if volume_str else 0.0
                    if volume == 0 or not volume:
                        continue
                    # 添加到data_dict
                    if str(hole_id) not in self.data_dict:
                        self.data_dict[str(hole_id)] = [{substance: volume}]
                    else:
                        self.data_dict[str(hole_id)].append({substance: volume})

                    # 添加到substance_dict
                    if substance not in self.reagent_plates_12:
                        self.reagent_plates_12[substance] = [{str(hole_id): volume}]
                    else:
                        self.reagent_plates_12[substance].append({str(hole_id): volume})

                elif item["liquidReagentConcentration"] == "固态":
                    volume_str = item.get("reactionQuality", '0')
                    volume = float(volume_str) if volume_str else 0.0
                    if volume == 0 or not volume:
                        continue
                    # 初始化该试剂的存储结构（如果不存在）
                    if substance not in self.solidity:
                        self.solidity[substance] = {}

                    # 存储当前孔位的信息
                    self.solidity[substance][hole_id] = {
                        "weight": volume,

                    }
                    counter += 1
            if self.type == 2:
                if item["type"] == "液体物料" or item["type"] == "溶剂":
                    volume_str = item.get("singleCockAddVolume", '0')
                    volume = float(volume_str) if volume_str else 0.0
                    if volume == 0 or not volume:
                        continue
                    # 添加到data_dict
                    if str(hole_id) not in self.data_dict:
                        self.data_dict[str(hole_id)] = [{substance: volume}]
                    else:
                        self.data_dict[str(hole_id)].append({substance: volume})

                    # 添加到substance_dict
                    if substance not in self.reagent_plates_12:
                        self.reagent_plates_12[substance] = [{str(hole_id): volume}]
                    else:
                        self.reagent_plates_12[substance].append({str(hole_id): volume})

                elif item["type"] == "固体物料":
                    if item["quality"] == 0 or not item["quality"]:
                        continue
                    # 初始化该试剂的存储结构（如果不存在）
                    if substance not in self.solidity:
                        self.solidity[substance] = {}

                    # 存储当前孔位的信息
                    self.solidity[substance][hole_id] = {
                        "weight": item["quality"],

                    }
                    counter += 1
                elif item["type"] == '反溶剂':
                    volume_str = item.get("singleCockAddVolume", '0')
                    volume = float(volume_str) if volume_str else 0.0
                    if volume == 0 or not volume:
                        continue
                    # 初始化该试剂的存储结构（如果不存在）
                    if substance not in self.antisolvent:
                        self.antisolvent[substance] = {}

                    # 存储当前孔位的信息
                    self.antisolvent[substance][hole_id] = {
                        "weight": volume,

                    }

    def _extract_materials(self) -> None:
        """提取物料数据"""
        materials = self.post_data["experimentLog1"]["MaterialsTableList"]
        for item in materials:
            substance = item["substance"].strip()
            self.theoretical_volume[substance] = item["theoreticalVolume"]
    def _hole_json(self) -> None:
        """生成孔位设置json文件"""
        current_time = datetime.datetime.now()

        # 格式化为指定格式
        formatted_time = current_time.strftime("%Y-%m-%d %H:%M:%S")

        hole_json = {
            "code": 200,
            "msg": "success",
            "data": {
                "hole_position": self.perforated_plate
            },
            "time": formatted_time,
            "traceid": self.traceid
        }
        if self.type  ==1:
            json_path = BASE_PATH

        else:
            json_path = crystallize_BASE_PATH
        json_name = '孔位设置参数.json'
        json_path = os.path.join(json_path, json_name)
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(hole_json, f, ensure_ascii=False)

    def _calculate_similarity(self) -> None:
        """计算反应相似度并分组"""
        # 准备数据集
        chemical_sets = {}
        for list_id, chem_list in self.data_dict.items():
            identifiers = {self._create_chemical_identifier(chem) for chem in chem_list}
            chemical_sets[list_id] = identifiers
        # 获取列表长度信息
        lengths = {list_id: len(chem_list) for list_id, chem_list in self.data_dict.items()}
        # 计算所有可能的配对组合
        all_pairs = []
        for (a, b) in combinations(chemical_sets.keys(), 2):
            jaccard, len_sim, combined, common_count, _ = self._calculate_similarity_pair(
                chemical_sets, lengths, a, b)
            all_pairs.append((a, b, combined, jaccard, len_sim, common_count))
        # 按相似度排序所有列表对
        sorted_pairs = sorted(all_pairs, key=lambda x: x[2], reverse=True)
        # 为每个列表找到最相似的匹配
        perforated_plate_ = []
        for pair in sorted_pairs:
            a, b, _, _, _, _ = pair
            if a not in perforated_plate_ and b not in perforated_plate_:
                self.perforated_plate.append([a, b])
                perforated_plate_.extend([a, b])

        all_reaction_ids = set(self.data_dict.keys())
        paired_ids = set(perforated_plate_)  # 已配对的ID
        unpaired_ids = all_reaction_ids - paired_ids  # 未配对的ID

        # 将未配对的单个反应作为独立组添加
        for uid in unpaired_ids:
            self.perforated_plate.append([uid])
        #生成孔位json数据到文件夹
        self._hole_json()
        # 处理24孔板数据
        for perforated in self.perforated_plate:
            if len(perforated) == 2:
                list_1, list_2 = sort_list(self.data_dict[f"{perforated[0]}"], self.data_dict[f"{perforated[1]}"])
                max_length = max(len(list_1), len(list_2))
                llwc_num = 1
                for i in range(max_length):
                    # 处理长度不一致的情况
                    item1 = list_1[i] if i < len(list_1) else None
                    item2 = list_2[i] if i < len(list_2) else None

                    if item1 is not None and item2 is not None and item1 == item2:
                        key = list(item1.keys())[0]
                        if key in list(self.reagent_plates_12_.keys()):
                            self.reagent_plates_12_[f'{key}'].append(perforated[0])
                        else:
                            self.reagent_plates_12_.update({key: [perforated[0]]})
                    else:
                        self.reagent_plates_24.append([item1, item2, perforated[0], llwc_num])
                        llwc_num += 4
            else:
                chem_list = self.data_dict[f"{perforated[0]}"]
                llwc_num = 1
                for chem_dict in chem_list:
                    chem = list(chem_dict.keys())[0]
                    vol = list(chem_dict.values())[0]

                    # 单个反应在24孔板中独占一个位置
                    self.reagent_plates_24.append([
                        {chem: vol},  # 左侧化学物质
                        None,  # 右侧为空
                        perforated[0],  # 孔位ID
                        llwc_num  # 位置编号
                    ])
                    llwc_num += 4

    @staticmethod
    def _create_chemical_identifier(chem_dict: dict) -> str:
        """创建化学物质标识符"""
        name = list(chem_dict.keys())[0]
        value = chem_dict[name]
        return f"{name}_{value}"

    @staticmethod
    def _calculate_similarity_pair(chemical_sets: dict, lengths: dict, a: str, b: str) -> tuple:
        """计算两个列表的相似度"""
        set_a = chemical_sets[a]
        set_b = chemical_sets[b]
        intersection = len(set_a & set_b)
        union = len(set_a | set_b)
        jaccard = intersection / union if union else 0

        len_a, len_b = lengths[a], lengths[b]
        length_sim = min(len_a, len_b) / max(len_a, len_b) if max(len_a, len_b) > 0 else 0

        combined = 0.5 * jaccard + 0.5 * length_sim
        common_names = {identifier.rsplit('_', 1)[0] for identifier in set_a & set_b}

        return jaccard, length_sim, combined, intersection, common_names

    def _generate_excel_files(self) -> None:
        """生成所有Excel文件"""
        if self.type == 1:
            # 添加淬灭剂数据
            self._create_task_quencher()
            # 添加内标数据
            self._create_task_internal_standard()
            # 添加物料表信息
            self.data_list_24, self.data_list_12 = self._create_mate_info()
            # 添加任务表信息
            self._create_w11_task()
            # 添加任务配置
            self._create_task_configuration()
            # 添加稀释数据
            self._create_dilution_data()
            # 添加萃取
            self._create_task_extract()
            # 添加sowo物料数据
            self._create_sowo_edit_stuff()
            # 添加sowo任务数据
            self._create_mate_msg()
        elif self.type == 2:
            # 重新写 函数 生成出表格
            self.data_list_24, self.data_list_12 = self._create_crystallize_mate_info()
            # 添加任务表信息
            self._create_crystallize_wf05_task()

            # 生成温度梯度表格
            self._create_crystallize_task_configuration()

            # 添加sowo物料数据
            self._create_sowo_crystallize_edit_stuff()
            # 添加sowo任务数据
            self._create_mate_msg()

    def _create_excel_xxx(self, input_path, output_path, rows):
        if self.type == 1:
            if rows == 1:wb = Workbook()
            else:wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["information"])
            ws = wb['Sheet']
            # 读取Excel文件
            df_material = pd.read_excel(input_path, sheet_name='物料信息')
            df_holes = pd.read_excel(input_path, sheet_name='孔位信息')
            # 获取物料表第2行和第3行的第2和第8列数据
            material_data = df_material.iloc[0:2, [1, 7]].reset_index(drop=True)
            # 获取孔位表第一列值为1和2的数据(1-6列)
            hole_data_1 = df_holes[df_holes.iloc[:, 0] == 1].iloc[:, [0,1,3,5]]
            hole_data_2 = df_holes[df_holes.iloc[:, 0] == 2].iloc[:, [0,1,3,5]]
            # 写入任务标签
            if rows == 1:
                # 合并单元格并写入"任务一"
                ws.merge_cells(start_row=rows, end_row=25, start_column=1, end_column=1)
                ws.cell(row=rows, column=1, value="任务一")
                alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(row=rows, column=1).alignment = alignment
                add_list = ['堆栈B2','堆栈B1']
            else:
                # 对于第二组数据，调整行号
                # 合并单元格并写入"任务一"
                ws.merge_cells(start_row=rows, end_row=rows + 24, start_column=1, end_column=1)
                ws.cell(row=rows, column=1, value="任务二")
                alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(row=rows, column=1).alignment = alignment
                add_list = ['堆栈H4', '堆栈H3']
            # 写入左侧数据(第2行物料+孔位1)
            # 物料第2行数据
            for col_idx, value in enumerate(material_data.iloc[0], start=2):
                ws.cell(row=rows, column=col_idx, value=value)
            # 添加堆栈
            ws.cell(row=rows, column=col_idx+ 1, value=add_list[0])
            # 孔位1数据
            for row_idx, row in hole_data_1.iterrows():
                for col_idx, value in enumerate(row, start=2):
                    ws.cell(row=row_idx + rows + 1, column=col_idx, value=value)
            # 写入右侧数据(第3行物料+孔位2)
            right_start_col = len(material_data.columns) + 7
            # 物料第3行数据
            for col_idx, value in enumerate(material_data.iloc[1], start=right_start_col):
                ws.cell(row=rows, column=col_idx, value=value)
            # 添加堆栈
            ws.cell(row=rows, column=col_idx + 1, value=add_list[1])
            if rows == 1:
                if self.number_tasks == 1:row_ = -10
                else:row_ = -22
            else:row_ = rows - 11
            # 孔位2数据
            for row_idx, row in hole_data_2.iterrows():
                for col_idx, value in enumerate(row, start=right_start_col):
                    ws.cell(row=row_idx + row_, column=col_idx, value=value)
            # 保存结果
            wb.save(output_path)
            print(f"处理完成，结果已保存到: {output_path}")
        else:
            if rows == 1:
                wb = Workbook()
            else:
                wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["information"])
            add_list = ['堆栈B4','堆栈B3','堆栈B2','堆栈B1']

            ws = wb['Sheet']
            # 读取Excel文件
            df_material = pd.read_excel(input_path, sheet_name='物料信息')
            df_holes = pd.read_excel(input_path, sheet_name='孔位信息')

            # 获取物料表第2行和第3行的第2和第8列数据
            material_data = df_material.iloc[0:5, [1, 7]].reset_index(drop=True)

            # 获取孔位表第一列值为1和2的数据(1-6列)
            hole_data_1 = df_holes[df_holes.iloc[:, 0] == 1].iloc[:, [0,1,3,5]]
            hole_data_2 = df_holes[df_holes.iloc[:, 0] == 2].iloc[:, [0,1,3,5]]
            hole_data_3 = df_holes[df_holes.iloc[:, 0] == 3].iloc[:, [0,1,3,5]]
            hole_data_4 = df_holes[df_holes.iloc[:, 0] == 4].iloc[:, [0,1,3,5]]
            # 写入任务标签
            if rows == 1:
                # 合并单元格并写入"任务一"
                ws.merge_cells(start_row=rows, end_row=25, start_column=1, end_column=1)
                ws.cell(row=rows, column=1, value="任务一")
                alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(row=rows, column=1).alignment = alignment

            else:
                # 对于第二组数据，调整行号
                # 合并单元格并写入"任务一"
                ws.merge_cells(start_row=rows, end_row=rows + 24, start_column=1, end_column=1)
                ws.cell(row=rows, column=1, value="任务二")
                alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(row=rows, column=1).alignment = alignment

            # 写入左侧数据(第2行物料+孔位1)
            # 物料第2行数据
            for col_idx, value in enumerate(material_data.iloc[0], start=2):
                ws.cell(row=rows, column=col_idx, value=value)
            # 添加堆栈
            ws.cell(row=rows, column=col_idx+ 1, value=add_list[0])
            # 孔位1数据
            for row_idx, row in hole_data_1.iterrows():
                for col_idx, value in enumerate(row, start=2):
                    ws.cell(row=row_idx + rows + 1, column=col_idx, value=value)

            # 写入右侧数据(第3行物料+孔位2)

            # 物料第3行数据
            for col_idx, value in enumerate(material_data.iloc[1], start=9):
                ws.cell(row=rows, column=col_idx, value=value)
            # 添加堆栈
            ws.cell(row=rows, column=col_idx+ 1, value=add_list[1])
            if rows == 1:
                if self.number_tasks == 1:
                    row_ = -10
                else:
                    row_ = -22
            else:
                row_ = rows - 11
            # 孔位2数据
            for row_idx, row in hole_data_2.iterrows():
                for col_idx, value in enumerate(row, start=9):
                    ws.cell(row=row_idx + row_, column=col_idx, value=value)

            for col_idx, value in enumerate(material_data.iloc[2], start=16):
                ws.cell(row=rows, column=col_idx, value=value)
            # 添加堆栈
            ws.cell(row=rows, column=col_idx+ 1, value=add_list[2])
            # 孔位1数据
            for row_idx, row in hole_data_3.iterrows():
                for col_idx, value in enumerate(row, start=16):
                    ws.cell(row=row_idx + rows - 35, column=col_idx, value=value)

            for col_idx, value in enumerate(material_data.iloc[3], start=23):
                ws.cell(row=rows, column=col_idx, value=value)
            # 添加堆栈
            ws.cell(row=rows, column=col_idx+ 1, value=add_list[3])
            # 孔位1数据
            for row_idx, row in hole_data_4.iterrows():
                for col_idx, value in enumerate(row, start=23):
                    ws.cell(row=row_idx + rows - 59, column=col_idx, value=value)

            # 保存结果
            wb.save(output_path)
            print(f"处理完成，结果已保存到: {output_path}")

    def _create_excel_xxxx(self, input_path, output_path, rows):
        if rows == 1:
            wb = Workbook()
        else:
            wb = ExcelUtils.open_workbook(output_path)
        ws = wb['Sheet']
        add_list = ['堆栈A3', '堆栈A2']
        # 读取Excel文件
        df_material = pd.read_excel(input_path, sheet_name='物料信息')
        df_holes = pd.read_excel(input_path, sheet_name='孔位信息')

        # 获取物料表第2行和第3行的第2和第8列数据
        material_data = df_material.iloc[1:3, [1, 7]].reset_index(drop=True)

        # 获取孔位表第一列值为1和2的数据(1-6列)
        hole_data_1 = df_holes[df_holes.iloc[:, 0] == 2].iloc[:, [0,1,3,5]]
        hole_data_2 = df_holes[df_holes.iloc[:, 0] == 3].iloc[:, [0,1,3,5]]

        # 写入任务标签
        if rows == 2:
            # 合并单元格并写入"任务一"
            ws.merge_cells(start_row=rows + 36, end_row=61, start_column=1, end_column=1)
            ws.cell(row=rows + 36, column=1, value="反溶剂添加")
            alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=rows + 36, column=1).alignment = alignment

        # 写入左侧数据(第2行物料+孔位1)
        # 物料第2行数据
        for col_idx, value in enumerate(material_data.iloc[0], start=2):
            ws.cell(row=38, column=col_idx, value=value)
        # 添加堆栈
        ws.cell(row=38, column=col_idx+ 1, value=add_list[0])
        # 孔位1数据
        for row_idx, row in hole_data_1.iterrows():
            for col_idx, value in enumerate(row, start=2):
                ws.cell(row=row_idx + 27, column=col_idx, value=value)

        # 写入右侧数据(第3行物料+孔位2)
        right_start_col = len(material_data.columns) + 7
        # 物料第3行数据
        for col_idx, value in enumerate(material_data.iloc[1], start=right_start_col):
            ws.cell(row=38, column=col_idx, value=value)
        #添加堆栈
        ws.cell(row=38, column=col_idx+ 1, value=add_list[1])
        # 孔位2数据
        for row_idx, row in hole_data_2.iterrows():
            for col_idx, value in enumerate(row, start=right_start_col):
                ws.cell(row=row_idx + rows + 1, column=col_idx, value=value)

        # 保存结果
        wb.save(output_path)
        print(f"处理完成，结果已保存到: {output_path}")

    def _create_excel_ttt(self, rowss):
        def get_row_values(sheet, row_num, prefix=""):
            data_list = [cell.value for cell in sheet[row_num]]
            return {
                '试剂板类型': data_list[1],
                '位置编号': data_list[7],
                '信息': f"{prefix}-{'左' if row_num == 2 else '右'}"}
        # 准备数据列表
        data_list_ = []
        for i, key in enumerate(list(self.solidity.keys()), 1):
            if i >= 10:i += 1
            data_list_.append({"sowo": "", "试剂名称": key, "编号": self.name_key_dict[key], "排序": i})
        # 分组处理逻辑
        grouped_data = []
        batch_count = len(data_list_) // 13 + (1 if len(data_list_) % 13 != 0 else 0)
        if batch_count > 1:
            # 多批次处理
            for batch_idx in range(batch_count):
                batch_name = f"第{batch_idx + 1}次任务"
                start_idx = batch_idx * 13
                end_idx = min((batch_idx + 1) * 13, len(data_list_))
                # 添加批次标题行
                grouped_data.append({
                    "sowo": batch_name,
                    "试剂名称": "",
                    "编号": "",
                    "信息": ""})
                # 添加批次数据
                for item in data_list_[start_idx:end_idx]:
                    grouped_data.append({
                        "sowo": "",
                        "试剂名称": item["试剂名称"],
                        "编号": item["编号"],
                        "信息": ""
                    })
        else:
            # 单批次处理
            grouped_data.append({
                "sowo": "第一次任务",
                "试剂名称": "",
                "编号": "",
                "信息": ""
            })
            for item in data_list_:
                grouped_data.append({
                    "sowo": "",
                    "试剂名称": item["试剂名称"],
                    "编号": item["编号"],
                    "信息": ""
                })
        # 处理模板配置
        configs = [
            (TEMPLATE_PATHS["mate_info1"], "1")
        ] if self.number_tasks == 1 else [
            (TEMPLATE_PATHS["mate_info1"], "1"),
            (TEMPLATE_PATHS["mate_info2"], "2")
        ]
        results = []
        for path, prefix in configs:
            wb = ExcelUtils.open_workbook(path)
            sheet = wb["物料信息"]
            results.extend([
                get_row_values(sheet, 2, prefix),
                get_row_values(sheet, 3, prefix)])
        wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["information"])
        sheet = wb['Sheet']
        headers = ['sowo', '试剂', '编号', "信息"]
        ExcelUtils.modify_row(wb, 'Sheet', rowss, headers)
        start_row_idx = rowss + 1  # 数据起始行（表头下一行）
        ExcelUtils.write_records_to_sheet(wb, 'Sheet', grouped_data)
        batch_start_rows = []
        for idx, item in enumerate(grouped_data):
            if item["sowo"]:  # 批次标题行
                batch_start_rows.append(start_row_idx + idx)
        alignment = Alignment(horizontal='center', vertical='center')
        for i, start_row in enumerate(batch_start_rows):
            if i < len(batch_start_rows) - 1:
                end_row = batch_start_rows[i + 1] - 1
            else:
                end_row = start_row_idx + len(grouped_data) - 1
            sheet.merge_cells(f'A{start_row}:A{end_row}')
            cell = sheet.cell(row=start_row, column=1)
            cell.alignment = alignment
        ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["information"])

    def _create_excel_tttt(self, rowss, input_path, output_path):
        # 准备数据列表
        data_list_ = []
        for i, key in enumerate(list(self.solidity.keys()), 1):
            if i >= 10:
                i += 1
            data_list_.append({"sowo": "", "试剂名称": key, "编号": self.name_key_dict[key], "排序": i})

        # 分组处理逻辑
        grouped_data = []
        batch_count = len(data_list_) // 13 + (1 if len(data_list_) % 13 != 0 else 0)

        if batch_count > 1:
            # 多批次处理
            for batch_idx in range(batch_count):
                batch_name = f"第{batch_idx + 1}次任务"
                start_idx = batch_idx * 13
                end_idx = min((batch_idx + 1) * 13, len(data_list_))

                # 添加批次标题行
                grouped_data.append({
                    "sowo": batch_name,
                    "试剂名称": "",
                    "编号": "",
                    "信息": ""
                })

                # 添加批次数据
                for item in data_list_[start_idx:end_idx]:
                    grouped_data.append({
                        "sowo": "",
                        "试剂名称": item["试剂名称"],
                        "编号": item["编号"],
                        "信息": ""
                    })
        else:
            # 单批次处理
            grouped_data.append({
                "sowo": "第一次任务",
                "试剂名称": "",
                "编号": "",
                "信息": ""
            })
            for item in data_list_:
                grouped_data.append({
                    "sowo": "",
                    "试剂名称": item["试剂名称"],
                    "编号": item["编号"],
                    "信息": ""
                })

        # 创建Excel并写入数据
        wb = ExcelUtils.open_workbook(output_path)
        sheet = wb['Sheet']

        # 修改表头
        headers = ['sowo', '试剂', '编号', "信息"]
        ExcelUtils.modify_row(wb, 'Sheet', rowss, headers)

        # 写入分组数据
        start_row_idx = rowss + 1  # 数据起始行（表头下一行）
        ExcelUtils.write_records_to_sheet(wb, 'Sheet', grouped_data)

        # 合并所有批次的单元格并设置居中

        # 找出所有批次标题行的位置
        batch_start_rows = []
        for idx, item in enumerate(grouped_data):
            if item["sowo"]:  # 批次标题行
                batch_start_rows.append(start_row_idx + idx)

        # 确定每个批次的结束行
        alignment = Alignment(horizontal='center', vertical='center')

        for i, start_row in enumerate(batch_start_rows):
            # 计算批次的结束行
            if i < len(batch_start_rows) - 1:
                end_row = batch_start_rows[i + 1] - 1
            else:
                end_row = start_row_idx + len(grouped_data) - 1

            # 合并该批次的单元格
            sheet.merge_cells(f'A{start_row}:A{end_row}')

            # 设置合并单元格的居中样式
            cell = sheet.cell(row=start_row, column=1)
            cell.alignment = alignment

        ExcelUtils.save_workbook(wb, output_path)

    def _create_mate_info(self) -> Tuple[list, list]:
        """创建物料信息Excel"""
        # 初始化数据列表
        data_list_12 = self._process_12_plate()
        data_list_24, result_1 = self._process_24_plate()

        # 文件路径

        # 生成第一个物料表格
        if len(data_list_24) > 6:
            ExcelUtils.copy_workbook(
                f"{BASE_PATH1}/hiwo物料-8.xlsx",
                TEMPLATE_PATHS["mate_info1"]
            )

            wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info1"])
            #在里面修改模版文件，把12孔换成24孔
            self._fill_worksheet(wb, data_list_12[:10], data_list_24[:12])
            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info1"])

            # 生成第二个物料表格
            if data_list_12:
                wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info2"])
                self._create_extended_version(wb, data_list_12, data_list_24[12:], self.number_tasks)
                ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info2"])

        else:
            """self.usage_12_hole是12孔板的使用最小使用孔数，最大孔是12孔，用最大值12减去使用的最小值，等于使用孔数，然后再用12减去它，等于剩余孔数，所以直接用24减去它"""
            # if int(24 - int(self.usage_12_hole)) >= len(data_list_12):
            self.number_tasks = 1
            wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info2"])
            self._create_extended_version(wb, data_list_12, data_list_24, self.number_tasks)
            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info2"])
            """如果24孔使用数量小于6，那么只需要一个任务就可以完成"""

        return result_1, data_list_12

    def _create_crystallize_mate_info(self) -> Tuple[list, list]:
        """创建物料信息Excel"""
        # 初始化数据列表

        data_list_12 = self._process_12_plate()
        data_list_24, result_1 = self._process_24_plate()
        ExcelUtils.copy_workbook(
            f"{crystallize_BASE_PATH1}/hiwo物料-8-02.xlsx",
            crystallize_TEMPLATE_PATHS["mate_info1"]
        )

        # 文件路径
        if data_list_12:
            wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["mate_info1"])
            self._create_extended_version(wb, data_list_12, data_list_24, self.number_tasks)
            ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["mate_info1"])

        # 生成第一个物料表格
        if len(data_list_24) > 18 - self.usage_24_hole:
            """先生成使用12孔的逻辑 ,再生成24孔的 """

            wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["mate_info1"])
            self._fill_worksheet(wb, data_list_12[:10], data_list_24[:18 - self.usage_24_hole])
            ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["mate_info1"])

            # 生成第二个物料表格
            wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["mate_info2"])
            self._fill_worksheet(wb, data_list_12[:10], data_list_24[18 - self.usage_24_hole:])
            ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["mate_info2"])

        else:
            self.number_tasks = 1
            wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["mate_info1"])
            # 12孔板
            self._create_extended_version(wb, data_list_12, data_list_24, self.number_tasks)
            # 24孔板
            self._fill_worksheet(wb, data_list_12[:10], data_list_24)
            ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["mate_info1"])
            """如果24孔使用数量小于6，那么只需要一个任务就可以完成"""

        return result_1, data_list_12

    # def process_plate_stack(self,
    #         file_path: str,
    #         output_path: str,
    #         transfer_records_12: list,
    #         transfer_records_24: list,
    #         transfer_records_2: list,
    #         plate_type: str,
    #         condition: dict,
    #         gun_head_dict: dict,
    #         Task_type =1
    # ):
    #     """
    #     处理 PlateStack 子表：
    #     1. 根据三个列表的长度标黄特定单元格。
    #     2. 根据 plate_type 修改 D2, E2, F2 的值。
    #     """
    #     wb = openpyxl.load_workbook(file_path)
    #     if "PlateStack" not in wb.sheetnames:
    #         raise ValueError("工作簿中不存在 'PlateStack' 子表")
    #     ws = wb["PlateStack"]
    #
    #     yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    #
    #     # ========== 第一部分：标黄规则 ==========
    #     if len(transfer_records_12) > 0:
    #         ws["E3"].fill = yellow_fill
    #
    #     len_24 = len(transfer_records_24)
    #     if Task_type ==1:
    #         if len_24 > 0 and len_24 < 6:
    #             ws["B4"].fill = yellow_fill
    #         elif len_24 > 6:
    #             ws["E4"].fill = yellow_fill
    #             ws["B5"].fill = yellow_fill
    #             if len(transfer_records_2) > 0:
    #                 ws["B4"].fill = yellow_fill
    #     else:
    #         if len_24 > 0 and len_24 < 6:
    #             ws["B4"].fill = yellow_fill
    #         elif len_24 > 6 and len_24 < 12:
    #             ws["B4"].fill = yellow_fill
    #             ws["C4"].fill = yellow_fill
    #         elif len_24 >12:
    #             ws["B4"].fill = yellow_fill
    #             ws["C4"].fill = yellow_fill
    #             ws["D4"].fill = yellow_fill
    #
    #
    #     # ========== 第二部分：根据孔板类型修改 D2, E2, F2 ==========
    #     if plate_type == "8孔":
    #         ws["F2"] = "8孔"
    #         ws["D2"] = "TA1-磁子分装_20ml-1"
    #         ws["E2"] = "TA1-称量载具_20ml-1"
    #     elif plate_type == "96孔":
    #         ws["F2"] = "96孔"
    #         ws["D2"] = "TA1-磁子分装_900ul-1"
    #         ws["E2"] = "TA1-称量载具_900ul-1"
    #     elif plate_type == '24孔':
    #         ws["F2"] = "24孔"
    #         if condition["experimentLog1"]["wellPlates"] == "24孔板8ml":
    #             ws["D2"] = "TA1-磁子分装_8ml-1"
    #             ws["E2"] = "TA1-称量载具_8ml-1"
    #         elif condition["experimentLog1"]["wellPlates"] == "24孔板4ml":
    #             ws["D2"] = "TA1-磁子分装_4ml-1"
    #             ws["E2"] = "TA1-称量载具_4ml-1"
    #         elif condition["experimentLog1"]["wellPlates"] == "24孔板2ml":
    #             ws["D2"] = "TA1-磁子分装_2ml-1"
    #             ws["E2"] = "TA1-称量载具_2ml-1"
    #     else:
    #         print(f"警告：未知的 plate_type '{plate_type}'，未修改 D2/E2/F2")
    #
    #     if "A05" in gun_head_dict:
    #         ws["E6"].fill = yellow_fill
    #     if "B05" in gun_head_dict:
    #         ws["B7"].fill = yellow_fill
    #     if "A04" in gun_head_dict:
    #         ws["C7"].fill = yellow_fill
    #     if "B04" in gun_head_dict:
    #         ws["D7"].fill = yellow_fill
    #
    #     wb.save(output_path)
    #     print(f"处理完成，保存至: {output_path}")

    def process_plate_stack(
            self,
            file_path: str,
            output_path: str,
            fill_cells: List[str],  # 新增参数：需要标黄的单元格列表

            plate_type: str,
            condition: dict,

    ):
        """
        处理 PlateStack 子表：
        1. 根据传入的 fill_cells 列表批量标黄单元格。
        2. 根据 plate_type / condition 修改 D2, E2, F2 的值（逻辑保持不变）。
        """
        wb = openpyxl.load_workbook(file_path)
        if "PlateStack" not in wb.sheetnames:
            raise ValueError("工作簿中不存在 'PlateStack' 子表")
        ws = wb["PlateStack"]

        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # ========== 第一部分：批量标黄（根据外部传入的列表） ==========
        for cell in fill_cells:
            ws[cell].fill = yellow_fill

        # ========== 第二部分：修改 D2, E2, F2 的值（原有逻辑不变） ==========
        if plate_type == "8孔":
            ws["F2"] = "8孔"
            ws["B4"] = "TA1-磁子分装_20ml-1"
            ws["B5"] = "TA1-称量载具_20ml-1"
        elif plate_type == "96孔":
            ws["F2"] = "96孔"
            ws["B4"] = "TA1-磁子分装_900ul-1"
            ws["B5"] = "TA1-称量载具_900ul-1"
        elif plate_type == "48孔":
            ws["F2"] = "48孔"
            ws["B4"] = "TA1-磁子分装_900ul-1"
            ws["B5"] = "TA1-称量载具_900ul-1"
        elif plate_type == '24孔':

            if condition["experimentLog1"]["wellPlates"] == "24孔板8ml":
                ws["F2"] = "24孔板8ml"
                ws["B4"] = "TA1-磁子分装_8ml-1"
                ws["B5"] = "TA1-称量载具_8ml-1"
            elif condition["experimentLog1"]["wellPlates"] == "24孔板4ml":
                ws["F2"] = "24孔板4ml"
                ws["B4"] = "TA1-磁子分装_4ml-1"
                ws["B5"] = "TA1-称量载具_4ml-1"
            elif condition["experimentLog1"]["wellPlates"] == "24孔板2ml":
                ws["F2"] = "24孔板2ml"
                ws["B4"] = "TA1-磁子分装_2ml-1"
                ws["B5"] = "TA1-称量载具_2ml-1"
        else:
            print(f"警告：未知的 plate_type '{plate_type}'，未修改 D2/E2/F2")

        wb.save(output_path)
        print(f"处理完成，保存至: {output_path}")
    def _create_crystallize_wf05_task(self) -> None:
        """创建W05任务Excel,在结晶中只有wf05任务,把所有都写进这个里面"""

        hole_mapping = {}
        for k, perforated_ in enumerate(self.perforated_plate, 1):
            if len(perforated_) == 2:
                for perforated in perforated_:
                    hole_mapping.update({perforated: k})
            # 单个反应的组：分配一个位置
            else:
                hole_mapping.update({perforated_[0]: k})
        # 生成转移数据
        transfer_records_12, transfer_records_24, transfer_records_2 = self._generate_transfer_data(hole_mapping)
        if self.number_tasks == 1:
            """只需要一张任务表wf11"""
            ExcelUtils.copy_workbook(
                f"{crystallize_BASE_PATH1}/hiwo任务05-8-24.xlsx",
                crystallize_TEMPLATE_PATHS["wf11"]
            )
            wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["wf11"])
            ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_12)
            if transfer_records_2:
                ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_2)
            ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_24)
            ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["wf11"])

        else:
            # 写入表格
            if transfer_records_2 or transfer_records_12:
                wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["wf11"])
                ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_12[:10])
                ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_2[:12])
                ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["wf11"])

            if transfer_records_24:
                wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["wf05"])
                ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_24)
                ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["wf05"])

        # fill_cells = []
        #
        # if len(transfer_records_12) > 0:
        #     fill_cells.append("C2")
        # len_24 = len(transfer_records_24)
        # if len_24 > 0 and len_24 <= 6:
        #     fill_cells.append("C3")
        # elif len_24 > 6 and len_24 <= 12:
        #     fill_cells.append("C3")
        #     fill_cells.append("C4")
        # elif len_24 >12:
        #     fill_cells.append("C3")
        #     fill_cells.append("C4")
        #     fill_cells.append("C5")

        fill_cells = []
        if len(transfer_records_12) > 0:
            fill_cells.append('C2')
        len_24 = len(transfer_records_24)
        if len_24 > 0 and len_24 <= 6:

            fill_cells.append('C3')
        elif len_24 > 6 :
            fill_cells.append('C6')
            fill_cells.append('C7')
            if len(transfer_records_2) > 0:
                fill_cells.append('C3')


        if "A05" in self.gun_head_dict:
            fill_cells.append("E4")
        if "B05" in self.gun_head_dict:
            fill_cells.append("E5")
        if "A04" in self.gun_head_dict:
            fill_cells.append("E2")
        if "B04" in self.gun_head_dict:
            fill_cells.append("E3")
            

        self.process_plate_stack(
        file_path=fr"{crystallize_BASE_PATH1}/Materials-stack4列.xlsx",
        output_path=fr"{crystallize_BASE_PATH}/Materials-stack4列.xlsx",
        fill_cells = fill_cells,
        plate_type='8孔',
        condition=self.post_data,

    )
    def add_excel_filter(self, num, type,key_dict = {}):

        result = []
        for group in self.perforated_plate:
            # 计算每组两个键对应的值总和
            sum1 = sum(value for dictionary in self.data_dict[group[0]]
                       for value in dictionary.values())
            # 如果type为1，添加antisolvent值
            if type == 2 :
                sum1 += int(key_dict.get(int(group[0]),{}).get("weight",''))
            try:
                # 计算第二个键的值总和
                sum2 = sum(value for dictionary in self.data_dict[group[1]]
                           for value in dictionary.values())
                # 如果type为1，添加antisolvent值
                if type == 2 :
                    sum2 += int(key_dict.get(int(group[1]),{}).get("weight",''))
            except:
                sum2 = 0

            # 比较并取较大值的键
            max_key = sum1 if sum1 >= sum2 else sum2
            result.append(max_key)

        excel_name = f"{crystallize_BASE_PATH}/hiwo-任务3-{num + 1}.xlsx"
        ExcelUtils.copy_workbook(
            crystallize_TEMPLATE_PATHS["WF10"], excel_name)
        ExcelUtils.copy_workbook(
            crystallize_TEMPLATE_PATHS["WF11"],
            f"{crystallize_BASE_PATH}/hiwo-物料3-{num + 1}.xlsx"
        )
        wb = ExcelUtils.open_workbook(excel_name)
        if type == 1:
            filtrationForm_time = self.post_data.get('experimentLog3', []).get('filtrationForm', []).get('time', 0)
        else:
            filtrationForm_time = self.post_data.get('experimentLog2', []).get('filtrationForm', []).get('time', 0)
        ExcelUtils.modify_cell(wb, "任务参数配置", "M3", int(filtrationForm_time))
        for s, i in enumerate(result, 1):
            pipette_location = 'A05'
            if len(self.gun_head_dict[pipette_location]) >= 12:
                pipette_location = 'B05'
            try:
                gun_head = self.gun_head_dict[pipette_location][-1]
            except:
                gun_head = 0
            ExcelUtils.modify_row(wb, "反应板混匀|||移液信息", s + 1,
                                  ["FY24-B01", s, "GL24-2A01", s, (i / 4) * 1000, pipette_location, gun_head],
                                  start_col=1)
            if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                self.gun_head_dict[pipette_location].append(gun_head + 1)
        ExcelUtils.save_workbook(wb, excel_name)

    def _create_crystallize_task_configuration(self) -> None:
        """创建任务配置Excel"""
        exp_log1 = self.post_data["experimentLog1"]
        exp_log2 = self.post_data["experimentLog2"]
        exp_log3 = self.post_data["experimentLog3"]

        if exp_log1["isType"] == '2':
            # 降温结晶
            excel_name = f"{crystallize_BASE_PATH}/hiwo-任务2-1.xlsx"
            ExcelUtils.copy_workbook(
                f"{crystallize_BASE_PATH1}/hiwo任务01-8-24.xlsx",
                excel_name
            )
            ExcelUtils.copy_workbook(
                f"{crystallize_BASE_PATH1}/hiwo物料-8-02.xlsx",
                f"{crystallize_BASE_PATH}/hiwo-物料2-1.xlsx"
            )
            wb = ExcelUtils.open_workbook(excel_name)
            temperature = int(exp_log3['dissolveForm']['temperature'])

            if temperature < 0:
                temperature = 0
            ExcelUtils.modify_cell(wb, "任务参数配置", "F3", int(temperature))

            time = int(exp_log3['dissolveForm']['time'])
            if time:
                time = int(time)
                if temperature:
                    if int(temperature) <= 26:
                        pass
                    elif int(temperature) <= 60:
                        time += 3
                    elif int(temperature) <= 100:
                        time += 6
                    elif int(temperature) <= 160:
                        time += 10
                    elif int(temperature) <= 200:
                        time += 18

            ExcelUtils.modify_cell(wb, "任务参数配置", "J3", int(time))
            rjSpeed = int(exp_log3['dissolveForm']['rjSpeed'])
            ExcelUtils.modify_cell(wb, "任务参数配置", "L3", int(rjSpeed))
            seal = "关盖" if exp_log3["dissolveForm"].get("seal") == "1" else "不关盖"
            ExcelUtils.modify_cell(wb, "任务参数配置", "E3", seal)
            ExcelUtils.modify_cell(wb, "任务参数配置", "N3", "开盖" if seal == "关盖" else "不开盖")
            ExcelUtils.save_workbook(wb, excel_name)
            if exp_log3['dissolveForm']['dissolve'] == '1':
                # 时间
                dissolveTimeInterval = exp_log3['dissolveForm']['dissolveTimeInterval']
                # 次数
                dissolveMaximum = exp_log3['dissolveForm']['dissolveMaximum']
                # 温度
                dissolveGradient = exp_log3['dissolveForm']['dissolveGradient']


                for num in range(2, int(dissolveMaximum) + 1):
                    temperature += int(dissolveGradient)
                    if temperature <= 0:
                        temperature = 0

                    if dissolveTimeInterval:
                        dissolveTimeInterval = int(dissolveTimeInterval)
                        if temperature:
                            if int(temperature) <= 26:
                                dissolveTimeIntervals = dissolveTimeInterval
                            elif int(temperature) <= 60:
                                dissolveTimeIntervals = dissolveTimeInterval + 3
                            elif int(temperature) <= 100:
                                dissolveTimeIntervals = dissolveTimeInterval + 6
                            elif int(temperature) <= 160:
                                dissolveTimeIntervals = dissolveTimeInterval + 10
                            elif int(temperature) <= 200:
                                dissolveTimeIntervals = dissolveTimeInterval + 18
                            else:
                                dissolveTimeIntervals = dissolveTimeInterval
                    excel_name = f"{crystallize_BASE_PATH}/hiwo-任务2-{num}.xlsx"
                    ExcelUtils.copy_workbook(
                        f"{crystallize_BASE_PATH1}/hiwo任务01-8-24.xlsx",
                        excel_name
                    )
                    ExcelUtils.copy_workbook(
                        f"{crystallize_BASE_PATH1}/hiwo物料-8-02.xlsx",
                        f"{crystallize_BASE_PATH}/hiwo-物料2-{num}.xlsx"
                    )
                    wb = ExcelUtils.open_workbook(excel_name)

                    ExcelUtils.modify_cell(wb, "任务参数配置", "F3", temperature)
                    ExcelUtils.modify_cell(wb, "任务参数配置", "J3", int(dissolveTimeIntervals))
                    ExcelUtils.modify_cell(wb, "任务参数配置", "L3", int(rjSpeed))
                    seal = "关盖" if exp_log3["dissolveForm"].get("seal") == "1" else "不关盖"
                    ExcelUtils.modify_cell(wb, "任务参数配置", "E3", seal)
                    ExcelUtils.modify_cell(wb, "任务参数配置", "N3", "开盖" if seal == "关盖" else "不开盖")
                    ExcelUtils.save_workbook(wb, excel_name)

            # 预冷 降温判断
            excel_name = f"{crystallize_BASE_PATH}/hiwo-任务3-1.xlsx"
            ExcelUtils.copy_workbook(
                f"{crystallize_BASE_PATH1}/hiwo任务01-8-24.xlsx",
                excel_name)
            ExcelUtils.copy_workbook(
                f"{crystallize_BASE_PATH1}/hiwo物料-8-02.xlsx",
                f"{crystallize_BASE_PATH}/hiwo-物料3-1.xlsx"
            )
            wb = ExcelUtils.open_workbook(excel_name)
            temperature = int(exp_log3['postProcessingForm']['temperature'])

            if temperature <= 0:
                temperature = 0

            ExcelUtils.modify_cell(wb, "任务参数配置", "F3", int(temperature))
            time = int(exp_log3['postProcessingForm']['time'])
            if time:
                time = int(time)
                if temperature:
                    if int(temperature) <= 26:
                        pass
                    elif int(temperature) <= 60:
                        time += 3
                    elif int(temperature) <= 100:
                        time += 6
                    elif int(temperature) <= 160:
                        time += 10
                    elif int(temperature) <= 200:
                        time += 18
            ExcelUtils.modify_cell(wb, "任务参数配置", "J3", int(time))
            speed = int(exp_log3['postProcessingForm']['speed'])
            ExcelUtils.modify_cell(wb, "任务参数配置", "L3", int(speed))
            #三阶段统一不关盖不开盖
            # seal = "关盖" if exp_log3["dissolveForm"].get("seal") == "1" else "不关盖"
            # ExcelUtils.modify_cell(wb, "任务参数配置", "E3", seal)
            # ExcelUtils.modify_cell(wb, "任务参数配置", "N3", "开盖" if seal == "关盖" else "不开盖")
            ExcelUtils.save_workbook(wb, excel_name)
            # 梯度降温析出判断
            num = 1
            if exp_log3['postProcessingForm']['precipitate'] == '1':
                # 时间
                dissolveTimeInterval = exp_log3['postProcessingForm']['dissolveTimeInterval']
                # 次数
                dissolveMaximum = exp_log3['postProcessingForm']['dissolveMaximum']
                # 温度
                dissolveGradient = exp_log3['postProcessingForm']['dissolveGradient']


                # num = 1
                for num in range(2, int(dissolveMaximum) + 1):
                    temperature -= int(dissolveGradient)
                    if temperature <= 0:
                        temperature = 0

                    if dissolveTimeInterval:
                        dissolveTimeInterval = int(dissolveTimeInterval)
                        if temperature:
                            if int(temperature) <= 26:
                                dissolveTimeIntervals = dissolveTimeInterval
                            elif int(temperature) <= 60:
                                dissolveTimeIntervals = dissolveTimeInterval + 3
                            elif int(temperature) <= 100:
                                dissolveTimeIntervals = dissolveTimeInterval + 6
                            elif int(temperature) <= 160:
                                dissolveTimeIntervals = dissolveTimeInterval + 10
                            elif int(temperature) <= 200:
                                dissolveTimeIntervals = dissolveTimeInterval + 18
                            else:
                                dissolveTimeIntervals = dissolveTimeInterval
                    excel_name = f"{crystallize_BASE_PATH}/hiwo-任务3-{num}.xlsx"
                    ExcelUtils.copy_workbook(
                        f"{crystallize_BASE_PATH1}/hiwo任务01-8-24.xlsx", excel_name)
                    ExcelUtils.copy_workbook(
                        f"{crystallize_BASE_PATH1}/hiwo物料-8-02.xlsx",
                        f"{crystallize_BASE_PATH}/hiwo-物料3-{num}.xlsx"
                    )
                    wb = ExcelUtils.open_workbook(excel_name)

                    ExcelUtils.modify_cell(wb, "任务参数配置", "F3", temperature)
                    ExcelUtils.modify_cell(wb, "任务参数配置", "J3", int(dissolveTimeIntervals))
                    ExcelUtils.modify_cell(wb, "任务参数配置", "L3", int(rjSpeed))
                    ExcelUtils.save_workbook(wb, excel_name)

            self.add_excel_filter(num, 1)

        if exp_log1["isType"] == '1':
            # 溶析结晶
            try:
                excel_name = f"{crystallize_BASE_PATH}/hiwo-任务2-1.xlsx"
                ExcelUtils.copy_workbook(
                    f"{crystallize_BASE_PATH1}/hiwo任务01-8-24.xlsx",
                    excel_name
                )
                ExcelUtils.copy_workbook(
                    f"{crystallize_BASE_PATH1}/hiwo物料-8-02.xlsx",
                    f"{crystallize_BASE_PATH}/hiwo-物料2-1.xlsx"
                )
                wb = ExcelUtils.open_workbook(excel_name)
                temperature = int(exp_log2['dissolveForm']['temperature'])
                if temperature <= 0:
                    temperature = 0

                ExcelUtils.modify_cell(wb, "任务参数配置", "F3", int(temperature))
                time = int(exp_log2['dissolveForm']['time'])
                if time:
                    time = int(time)
                    if temperature:
                        if int(temperature) <= 26:
                            pass
                        elif int(temperature) <= 60:
                            time += 3
                        elif int(temperature) <= 100:
                            time += 6
                        elif int(temperature) <= 160:
                            time += 10
                        elif int(temperature) <= 200:
                            time += 18

                ExcelUtils.modify_cell(wb, "任务参数配置", "J3", int(time))
                rjSpeed = int(exp_log2['dissolveForm']['rjSpeed'])
                ExcelUtils.modify_cell(wb, "任务参数配置", "L3", int(rjSpeed))
                seal = "关盖" if exp_log2["dissolveForm"].get("seal") == "1" else "不关盖"
                ExcelUtils.modify_cell(wb, "任务参数配置", "E3", seal)
                ExcelUtils.modify_cell(wb, "任务参数配置", "N3", "开盖" if seal == "关盖" else "不开盖")
                ExcelUtils.save_workbook(wb, excel_name)
                # 析出判断
                if exp_log2['dissolveForm']['dissolve'] == '1':
                    # 时间
                    dissolveTimeInterval = exp_log2['dissolveForm']['dissolveTimeInterval']
                    # 次数
                    dissolveMaximum = exp_log2['dissolveForm']['dissolveMaximum']
                    # 温度
                    dissolveGradient = exp_log2['dissolveForm']['dissolveGradient']


                    for num in range(2, int(dissolveMaximum) + 1):
                        temperature += int(dissolveGradient)

                        if dissolveTimeInterval:
                            dissolveTimeInterval = int(dissolveTimeInterval)
                            if temperature:
                                if int(temperature) <= 26:
                                    dissolveTimeIntervals = dissolveTimeInterval
                                elif int(temperature) <= 60:
                                    dissolveTimeIntervals = dissolveTimeInterval + 3
                                elif int(temperature) <= 100:
                                    dissolveTimeIntervals = dissolveTimeInterval + 6
                                elif int(temperature) <= 160:
                                    dissolveTimeIntervals = dissolveTimeInterval + 10
                                elif int(temperature) <= 200:
                                    dissolveTimeIntervals = dissolveTimeInterval + 18
                                else:
                                    dissolveTimeIntervals = dissolveTimeInterval

                        excel_name = f"{crystallize_BASE_PATH}/hiwo-任务2-{num}.xlsx"
                        ExcelUtils.copy_workbook(
                            f"{crystallize_BASE_PATH1}/hiwo任务01-8-24.xlsx",
                            excel_name
                        )
                        ExcelUtils.copy_workbook(
                            f"{crystallize_BASE_PATH1}/hiwo物料-8-02.xlsx",
                            f"{crystallize_BASE_PATH}/hiwo-物料2-{num}.xlsx"
                        )
                        wb = ExcelUtils.open_workbook(excel_name)

                        if temperature < 0:
                            temperature = 0

                        ExcelUtils.modify_cell(wb, "任务参数配置", "F3", temperature)
                        ExcelUtils.modify_cell(wb, "任务参数配置", "J3", int(dissolveTimeIntervals))
                        ExcelUtils.modify_cell(wb, "任务参数配置", "L3", int(rjSpeed))
                        seal = "关盖" if exp_log2["dissolveForm"].get("seal") == "1" else "不关盖"
                        ExcelUtils.modify_cell(wb, "任务参数配置", "E3", seal)
                        ExcelUtils.modify_cell(wb, "任务参数配置", "N3", "开盖" if seal == "关盖" else "不开盖")
                        ExcelUtils.save_workbook(wb, excel_name)
                # 反溶剂添加
                self._create_crystallize_task_quencher()
            except:
                pass

    def _create_crystallize_task_quencher(self):
        '''
        反溶剂正加
        把反溶剂加到反应孔
        生成物料表,把每个孔对应的反溶剂放到24孔对应的孔位中

        反加: 把反应孔试剂都吸到试剂板

        :return:
        '''
        hole_mapping = {}
        for k, perforated_ in enumerate(self.perforated_plate, 1):
            if len(perforated_) == 2:
                for perforated in perforated_:
                    hole_mapping.update({perforated: k})
            # 单个反应的组：分配一个位置
            else:
                hole_mapping.update({perforated_[0]: k})
        postProcessingForm = self.post_data['experimentLog2']['postProcessingForm']
        precoolingTemperature = postProcessingForm['precoolingTemperature']
        time = postProcessingForm['time']
        # temperature = postProcessingForm['temperature']
        speed = postProcessingForm['speed']

        all_solid_records = []
        for reagent_name, hole_data in self.antisolvent.items():
            for hole_id, details in hole_data.items():
                all_solid_records.append({
                    "reagent_name": reagent_name,
                    "hole_id": hole_id,
                    "weight": hole_data[hole_id]['weight']
                })

        if postProcessingForm['method'] == '2':
            #反加
            # 物料模板
            excel_name = fr"{crystallize_BASE_PATH}/hiwo-物料3-1.xlsx"
            ExcelUtils.copy_workbook(
                f"{crystallize_BASE_PATH1}/hiwo-物料模板-8.xlsx",
                excel_name
            )

            wb = ExcelUtils.open_workbook(excel_name)
            key_dict = {}
            for data in all_solid_records:
                key_dict.update({data['hole_id']: data})
            for i, key in enumerate(self.perforated_plate):
                row = 14 + i * 4
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}",
                                       key_dict.get(int(key[0]), []).get("reagent_name", ''))
                if len(key) == 2:
                    ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 2}",
                                           key_dict.get(int(key[1]), []).get("reagent_name", ''))
            ExcelUtils.save_workbook(wb, excel_name)

            # 任务模板
            # if postProcessingForm['precipitate'] =='1':
            excel_name = fr"{crystallize_BASE_PATH}/hiwo-任务3-1.xlsx"
            ExcelUtils.copy_workbook(
                f"{crystallize_BASE_PATH1}/hiwo任务01-8-24.xlsx",
                excel_name
            )
            wb = ExcelUtils.open_workbook(excel_name)


            if int(precoolingTemperature) <= 0:
                precoolingTemperature = 0
            ExcelUtils.modify_cell(wb, "任务参数配置", f"F3", precoolingTemperature)
            ExcelUtils.modify_cell(wb, "任务参数配置", f"J3", time)
            ExcelUtils.modify_cell(wb, "任务参数配置", f"L3", speed)
            # seal = "关盖" if self.post_data['experimentLog2']['dissolveForm'].get("seal") == "1" else "不关盖"
            # ExcelUtils.modify_cell(wb, "任务参数配置", "E3", seal)
            # ExcelUtils.modify_cell(wb, "任务参数配置", "N3", "开盖" if seal == "关盖" else "不开盖")


            #A02加到B02
            record_list = []
            #金属板加到B02
            record_lists = []
            #B02加回金属板
            record_list1 = []
            #加一个过滤逻辑3-2
            record_list2 = []
            pipette_location = "A05"

            for num, list_data in enumerate(self.perforated_plate, 1):
                # 计算每组两个键对应的值总和
                sum1 = sum(value for dictionary in self.data_dict[list_data[0]]
                           for value in dictionary.values())
                try:
                    # 计算第二个键的值总和
                    sum2 = sum(value for dictionary in self.data_dict[list_data[1]]
                               for value in dictionary.values())
                except:
                    sum2 = 0

                # 比较并取较大值的键
                max_key = sum1 if sum1 >= sum2 else sum2

                if len(self.gun_head_dict[pipette_location]) >= 12:
                    pipette_location = "B05"

                gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0
                target_volume = max_key * 1000 / 2 if max_key else 0
                bar_code = "SS-B02"

                #获取反溶剂最大剂量
                target_ids = [int(x) for x in list_data]

                # 收集符合条件的weight值
                weights = []
                for item in key_dict.values():
                    if item['hole_id'] in target_ids:
                        weights.append(item['weight'])

                # 找出最大的weight值
                max_weight = max(weights)* 1000 / 2 if weights else None
                record = {
                    "来源孔板条码": "SS-A03",
                    "来源孔板列Y": num,
                    "目标孔板条码": bar_code,
                    "目标孔板列Y": num,
                    "移液量(ul)": max_weight,
                    "枪头库位": pipette_location,
                    "枪头列Y": gun_head + 1
                }
                record_list.append(record)

                if len(self.gun_head_dict[pipette_location]) >= 12:
                    pipette_location = "B05"

                if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                    self.gun_head_dict[pipette_location].append(gun_head + 1)

                bar_code = "SS-B02"
                gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0
                record = {
                    "来源孔板条码": "FY24-B01",
                    "来源孔板列Y": num,
                    "目标孔板条码": bar_code,
                    "目标孔板列Y": num,
                    "移液量(ul)": target_volume,
                    "枪头库位": pipette_location,
                    "枪头列Y": gun_head + 1
                }
                record_lists.append(record)

                # if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                #     self.gun_head_dict[pipette_location].append(gun_head + 1)
                #
                # gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0
                #总移液量是前面两个的和
                Total_pipetting_volume = max_weight+target_volume
                record = {
                    "来源孔板条码": "SS-B02",
                    "来源孔板列Y": num,
                    "目标孔板条码": 'FY24-B01',
                    "目标孔板列Y": num,
                    "移液量(ul)": Total_pipetting_volume,
                    "枪头库位": pipette_location,
                    "枪头列Y": gun_head + 1
                }
                record_list1.append(record)

                # if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                #     self.gun_head_dict[pipette_location].append(gun_head + 1)
                #
                # gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0
                #总移液量是前面两个的和
                Total_pipetting_volume = max_weight+target_volume
                record = {
                    "来源孔板条码": "FY24-B01",
                    "来源孔板列Y": num,
                    "目标孔板条码": 'GL24-2A01',
                    "目标孔板列Y": num,
                    "移液量(ul)": Total_pipetting_volume,
                    "枪头库位": pipette_location,
                    "枪头列Y": gun_head + 1
                }
                record_list2.append(record)

                if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                    self.gun_head_dict[pipette_location].append(gun_head + 1)
            ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", record_list)
            ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", record_lists)
            ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", record_list1)
            ExcelUtils.save_workbook(wb, excel_name)


            excel_name1 = fr"{crystallize_BASE_PATH}/hiwo-物料3-2.xlsx"
            ExcelUtils.copy_workbook(
                crystallize_TEMPLATE_PATHS["WF11"],
                excel_name1
            )
            excel_name1 = fr"{crystallize_BASE_PATH}/hiwo-任务3-2.xlsx"
            ExcelUtils.copy_workbook(
                crystallize_TEMPLATE_PATHS["WF10"],
                excel_name1
            )
            wb1 = ExcelUtils.open_workbook(excel_name1)

            if type == 1:
                filtrationForm_time = self.post_data.get('experimentLog3', []).get('filtrationForm', []).get('time', 0)
            else:
                filtrationForm_time = self.post_data.get('experimentLog2', []).get('filtrationForm', []).get('time', 0)
            ExcelUtils.modify_cell(wb1, "任务参数配置", "M3", int(filtrationForm_time))


            ExcelUtils.write_records_to_sheet(wb1, "反应板混匀|||移液信息", record_list2)

            ExcelUtils.save_workbook(wb1, excel_name1)

        else:
            #正加
            key_dict = {}
            for data in all_solid_records:
                key_dict.update({data['hole_id']: data})
            excel_name = fr"{crystallize_BASE_PATH}/hiwo-物料3-1.xlsx"
            ExcelUtils.copy_workbook(
                f"{crystallize_BASE_PATH1}/hiwo-物料模板.xlsx",
                excel_name
            )
            wb = ExcelUtils.open_workbook(excel_name)
            if all_solid_records:
                for i, key in enumerate(self.perforated_plate):
                    i =1 + i*2
                    row = 14 + (i-1) * 4

                    ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}",
                                           key_dict.get(int(key[0]), []).get("reagent_name", ''))
                    ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 2}",'')
                    if len(key) == 2:
                        row = 14 + i * 4
                        ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}", '')
                        ExcelUtils.modify_cell(wb, "孔位信息", f"D{row+2}",
                                               key_dict.get(int(key[1]), []).get("reagent_name", ''))

                ExcelUtils.save_workbook(wb, excel_name)
            # if postProcessingForm['precipitate'] == '1':
            excel_name = fr"{crystallize_BASE_PATH}/hiwo-任务3-1.xlsx"
            ExcelUtils.copy_workbook(
                f"{crystallize_BASE_PATH1}/hiwo任务01-8-24.xlsx",
                excel_name
            )
            wb = ExcelUtils.open_workbook(excel_name)

            if int(precoolingTemperature) <= 0:
                precoolingTemperature = 0
            ExcelUtils.modify_cell(wb, "任务参数配置", f"F3", precoolingTemperature)
            ExcelUtils.modify_cell(wb, "任务参数配置", f"J3", time)
            ExcelUtils.modify_cell(wb, "任务参数配置", f"L3", speed)
            # seal = "关盖" if self.post_data['experimentLog2']['dissolveForm'].get("seal") == "1" else "不关盖"
            # ExcelUtils.modify_cell(wb, "任务参数配置", "E3", seal)
            # ExcelUtils.modify_cell(wb, "任务参数配置", "N3", "开盖" if seal == "关盖" else "不开盖")
            record_list = []
            pipette_location = "A05"
            for llwc, list_data in enumerate(self.perforated_plate, 1):
                if len(self.gun_head_dict[pipette_location]) >= 12:
                    pipette_location = "B05"
                # max_num = 0
                # for data in list_data:
                #     num = self.data_dict[data]
                #     key_num = 0
                #     for s in num:
                #         key_num += sum(s.values())
                #     if key_num >= max_num:
                #         max_num = key_num
                for key in list_data:
                    gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0

                    target_volume = int(key_dict.get(int(key), []).get("weight", ''))/2 *1000
                    target_id = int(key_dict.get(int(key), []).get("hole_id", ''))
                    bar_code = "SS-A03"

                    record = {
                        "来源孔板条码": bar_code,
                        "来源孔板列Y": target_id,
                        "目标孔板条码": "FY24-B01",
                        "目标孔板列Y": llwc,
                        "移液量(ul)": target_volume,
                        "枪头库位": pipette_location,
                        "枪头列Y": gun_head + 1
                    }
                    if len(record_list)>=6:
                        record['来源孔板条码']='SS-B02'
                        record['来源孔板列Y']= target_id -6
                    record_list.append(record)
                    if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                        self.gun_head_dict[pipette_location].append(gun_head + 1)
            ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", record_list)
            ExcelUtils.save_workbook(wb, excel_name)
            self.add_excel_filter(1, 2,key_dict)

    def _create_sowo_crystallize_edit_stuff(self) -> None:
        """创建sowoEditStuff Excel并记录物料表号"""
        # 收集所有固态记录
        all_solid_records = []
        for reagent_name, hole_data in self.solidity.items():
            for hole_id, details in hole_data.items():
                all_solid_records.append({
                    "reagent_name": reagent_name,
                    "hole_id": hole_id
                })

        # 按试剂名称去重
        seen = set()
        unique_reagents = [x for x in all_solid_records
                           if x["reagent_name"] not in seen
                           and not seen.add(x["reagent_name"])]

        # 生成物料编码
        day_name = datetime.datetime.now().strftime("%Y%m%d")
        file_path = os.path.join(BASE_PATH1, f"物料key.json")
        data = {
            "daily_counters": {},  # 存储最新计数器
            "materials": OrderedDict()  # 有序字典保存试剂-编码映射
        }

        if os.path.exists(file_path):
            try:
                with open(file_path, 'r') as f:
                    data = json.load(f, object_pairs_hook=OrderedDict)
            except json.JSONDecodeError:
                pass  # 文件损坏时使用默认数据
        # 更新试剂编码映射

        # 初始化当天计数器
        if day_name not in data["daily_counters"]:
            data["daily_counters"][day_name] = 0
            counter = 0

        else:
            counter = data["daily_counters"][day_name]

        need_save = False

        for reagent in unique_reagents:
            reagent_name = reagent["reagent_name"]

            if reagent_name in data["materials"]:
                key = data["materials"][reagent_name]
                self.name_key_dict[reagent_name] = key

            # 如果试剂不存在则生成新编码
            else :
                counter += 1
                key = f"{day_name}-{str(counter).zfill(3)}"
                data["materials"][reagent_name] = key
                self.name_key_dict[reagent_name] = key
                need_save = True
        # 更新计数器并保存数据
        if need_save:
            data["daily_counters"][day_name] = counter
        with open(file_path, 'w') as f:
            json.dump(data, f, indent=2)  # 美观格式保存

        # 创建物料表记录并记录表号
        self.reagent_table_map = {}  # 新增：存储物料->表号的映射
        table_num = 1
        excel_list = []
        name_list = []

        # 第一张表处理
        s = 0
        for record in all_solid_records:
            if record["reagent_name"] in name_list:
                continue

            s += 1
            if s == 10:  # 跳过第10号位置
                s += 1
            if s > 14:  # 第一表最多13个位置
                break

            excel_list.append({
                "Code": self.name_key_dict[record["reagent_name"]],
                "Locate": s,
                "Type": 2
            })
            # 记录物料所属表号 (表1)
            self.reagent_table_map[record["reagent_name"]] = table_num
            name_list.append(record["reagent_name"])

        # 添加溶剂位置
        day_name = datetime.datetime.now().strftime("%Y%m%d")

        ExcelUtils.copy_workbook(
            f"{crystallize_BASE_PATH1}/sowo物料.xlsx",
            crystallize_TEMPLATE_PATHS["sowo"]
        )
        # 保存第一张表
        wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["sowo"])
        ExcelUtils.write_records_to_sheet(wb, "Sheet1", excel_list)
        ExcelUtils.write_records_to_sheet(wb, "Sheet1", [{"Code": f"{day_name}-001", "Locate": 1, "Type": 1}])
        ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["sowo"])

        # 处理第二张表 (如果有超过13种物料)
        excel_list_2 = []
        if len(name_list) < len(unique_reagents):
            table_num = 2  # 表号递增

            s = 0

            # 处理剩余物料
            for record in all_solid_records:
                if record["reagent_name"] in name_list:
                    continue

                s += 1
                if s > 14:  # 第二表也最多13个位置
                    break

                excel_list_2.append({
                    "Code": self.name_key_dict[record["reagent_name"]],
                    "Locate": s,
                    "Type": 2
                })
                # 记录物料所属表号 (表2)
                self.reagent_table_map[record["reagent_name"]] = table_num
                name_list.append(record["reagent_name"])

            # 保存第二张表
            target_path = crystallize_TEMPLATE_PATHS["sowo"].replace("1.xlsx", "2.xlsx")
            ExcelUtils.copy_workbook(f"{crystallize_BASE_PATH1}/sowo物料.xlsx", target_path)
            wb = ExcelUtils.open_workbook(target_path)
            ExcelUtils.write_records_to_sheet(wb, "Sheet1", excel_list_2)
            ExcelUtils.write_records_to_sheet(wb, "Sheet1", [{"Code": f"{day_name}-001", "Locate": 1, "Type": 1}])
            ExcelUtils.save_workbook(wb, target_path)
        self.write_codes_to_inner_hopper_stack(fr"{crystallize_BASE_PATH}/Materials-stack4列.xlsx",fr"{crystallize_BASE_PATH}/Materials-stack4列.xlsx",excel_list,excel_list_2)
        # self.set_hopper_b2_with_utils(fr"{crystallize_BASE_PATH}/Materials-stack4列.xlsx",fr"{crystallize_BASE_PATH}/Materials-stack4列.xlsx",f"{day_name}-001")

    def _process_12_plate(self) -> List[str]:
        """处理12孔板数据"""
        result = []
        result_1 = []
        for substance in self.reagent_plates_12_.keys():
            # list_ = self.reagent_plates_12.get(substance,[])
            try:
                vol = float(self.theoretical_volume.get(substance, 0))
                if vol > 18:
                    num = math.ceil(vol / 18)
                    result.extend([substance] * num)

                elif substance not in result:
                    result.append(substance)
                    # result_1.append([chem1, chem2, item[2], chem3])

            except (ValueError, TypeError):
                continue
        return result

    def _process_24_plate(self) -> Tuple[List[List[Union[str, None]]], list]:
        """处理24孔板数据"""
        result = []
        result_1 = []

        for item in self.reagent_plates_24:
            chem1 = list(item[0].keys())[0] if item[0] else None
            chem2 = list(item[1].keys())[0] if item[1] and len(item) > 1 else None
            chem3 = list(item[0].values())[0] if item[0] else None
            chem4 = list(item[1].values())[0] if item[1] else None

            if chem2 and chem3 == chem4:
                if chem3 <= 8:
                    result.append([chem1, chem2])
                    result_1.append([chem1, chem2, item[2], chem3])
                else:
                    result.append([chem1, chem2])
                    result_1.append([chem1, chem2, item[2], 8])
                    result.append([chem1, chem2])
                    result_1.append([chem1, chem2, item[2], chem3-8])

            else:
                if chem1:
                    if chem3<=8:
                        result.append([chem1, None])
                        result_1.append([chem1, None, item[2], chem3])
                    else:
                        result.append([chem1, None])
                        result_1.append([chem1, None, item[2], 8])
                        result.append([chem1, None])
                        result_1.append([chem1, None, item[2], chem3-8])
                if chem2:
                    if chem4 <= 8:
                        result.append([None, chem2])
                        result_1.append([None, chem2, item[2], chem4])
                    else:
                        result.append([None, chem2])
                        result_1.append([None, chem2, item[2], 8])
                        result.append([None, chem2])
                        result_1.append([None, chem2, item[2], chem4-8])
        return result, result_1

    def _fill_worksheet(self, wb: Workbook, data_12: List[str],
                        data_24: List[List[Union[str, None]]]) -> None:
        """填充工作表数据"""

        if self.type == 1:
            # 填充24孔板数据
            ExcelUtils.modify_row(wb, "物料信息", 2,
                                  ["24孔试剂板-低板", "24孔板", "SS-A02", 1, "堆栈", 1, "A02"],
                                  start_col=2)
            ExcelUtils.delete_row(wb, "孔位信息", 14)
            self._copy_insert_rows(wb, "孔位信息")
            start = 2
        else:
            start = 14
        for i, (chem1, chem2) in enumerate(data_24):
            row = start + i * 4
            ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}", chem1)
            if chem2:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 2}", chem2)

    def _copy_insert_rows(self, wb: Workbook, sheet_name: str) -> None:
        """复制并插入行"""
        sheet = wb[sheet_name]

        # 复制行数据
        copied_rows = []
        for row in sheet.iter_rows(min_row=2, max_row=25, values_only=True):
            copied_rows.append(row)

        # 插入新行
        sheet.insert_rows(2, amount=24)

        # 填充新行数据
        for i, row_data in enumerate(copied_rows, start=2):
            sheet.cell(row=i, column=1, value=1)
            sheet.cell(row=i, column=3, value=f"SS-A02-0{i - 1}")
            for col_idx, value in enumerate(row_data[1:], start=2):
                if col_idx != 3:
                    sheet.cell(row=i, column=col_idx, value=value)

    def _create_extended_version(self, wb: Workbook, data_12: List[str],
                                 data_24: List[List[Union[str, None]]], number_tasks) -> None:
        """创建扩展版本工作表"""
        # 填充12孔板物料数据,最大值取决于物料和内标使用的值
        stop_num = self.usage_12_hole - 1
        for i, chem in enumerate(data_12[:stop_num], 2):
            ExcelUtils.modify_cell(wb, "孔位信息", f"D{i}", chem)

        # 处理12孔板扩展数据，llwc_num：记录12孔板用了多少24孔板的位置
        llwc_num = 0
        if len(data_12) > stop_num:
            for i, chem in enumerate(data_12[stop_num:]):
                row = 14 + i * 4
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}", chem)
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 2}", chem)
                llwc_num += 1
                if llwc_num > 6:
                    """如果12孔板剩余的数量大于6，也就是说一块24孔板不够使用，那么也是应该第三次实验吗？"""
                    pass
                    break
        if data_24:
            for i ,chem in enumerate(data_24):
                row = 14+llwc_num*4+i*4
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}", chem[0])
                if len(chem)==2:
                    ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 2}", chem[1])



        # 处理24孔板扩展数据,从哪开始取决于 12孔板用完用了多少24孔板的孔位
        """
        如果是生成第二个表的逻辑，第一个表使用两块24孔板，第二块表使用一块12孔板一块24孔板
        data_24_start+2 + i * 4： data_24_start默认是12，但是在excel表格里12孔板用完行数是13，应该从14行开始才是24孔板
        如果只需要一个表，也就是使用一块12孔板以及一块24孔板就行，那么该怎么修改这个逻辑

        """
        if number_tasks == 1:
            data_24_start = 0
        else:
            data_24_start = 12

        if len(data_24) > data_24_start:
            for i, (chem1, chem2) in enumerate(data_24[data_24_start:18]):
                row = 12 + 2 + llwc_num * 4 + i * 4
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}", chem1)
                if chem2:
                    ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 2}", chem2)
        if len(data_24) > 18:
            """如果第二个任务的24孔板也不够用，是不是需要第三次任务？"""
            pass
    def _create_w11_task(self) -> None:
        """创建W11任务Excel"""
        # 创建孔位映射
        # hole_mapping = {}
        # position_counter = 1
        #
        # for perforated in self.perforated_plate:
        #     # 两个反应的组：分配连续的两个位置
        #     if len(perforated) == 2:
        #         hole_mapping[perforated[0]] = position_counter
        #         hole_mapping[perforated[1]] = position_counter + 1
        #         position_counter += 2
        #     # 单个反应的组：分配一个位置
        #     elif len(perforated) == 1:
        #         hole_mapping[perforated[0]] = position_counter
        #         position_counter += 1
        hole_mapping = {}
        for k, perforated_ in enumerate(self.perforated_plate, 1):
            if len(perforated_) == 2:
                for perforated in perforated_:
                    hole_mapping.update({perforated: k})
            # 单个反应的组：分配一个位置
            else:
                hole_mapping.update({perforated_[0]: k})
        # 生成转移数据
        transfer_records_12, transfer_records_24, transfer_records_2 = self._generate_transfer_data(hole_mapping)
        if self.number_tasks == 1:
            """只需要一张任务表wf11"""
            wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf11"])
            ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_12)
            if transfer_records_2:
                ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_2)
            ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_24)
            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf11"])
        else:
            # 写入表格

            if transfer_records_2 or transfer_records_12:
                wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf11"])
                ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_12[:10])
                ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_2[:12])
                ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf11"])
            if transfer_records_24:
                ExcelUtils.copy_workbook(
                    f"{BASE_PATH1}/hiwo任务05-8-24.xlsx",
                    TEMPLATE_PATHS["wf05"]
                )
                wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf05"])
                ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_24)
                ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf05"])
        # fill_cells = []
        # if len(transfer_records_12) > 0:
        #     fill_cells.append('E3')
        # len_24 = len(transfer_records_24)
        # if len_24 > 0 and len_24 <= 6:
        #
        #     fill_cells.append('B4')
        # elif len_24 > 6:
        #     fill_cells.append('E4')
        #     fill_cells.append('B5')
        #     if len(transfer_records_2) > 0:
        #         fill_cells.append('B4')
        fill_cells = []
        if len(transfer_records_12) > 0:
            fill_cells.append('C2')
        len_24 = len(transfer_records_24)
        if len_24 > 0 and len_24 <= 6:

            fill_cells.append('C3')
        elif len_24 > 6 :
            fill_cells.append('C6')
            fill_cells.append('C7')
            if len(transfer_records_2) > 0:
                fill_cells.append('C3')
        fill_cells = []
        if len(transfer_records_12) > 0:
            fill_cells.append('C2')
        len_24 = len(transfer_records_24)
        if len_24 > 0 and len_24 <= 6:

            fill_cells.append('C3')
        elif len_24 > 6 :
            fill_cells.append('C6')
            fill_cells.append('C7')
            if len(transfer_records_2) > 0:
                fill_cells.append('C3')

        if "A05" in self.gun_head_dict:
            fill_cells.append("E4")
        if "B05" in self.gun_head_dict:
            fill_cells.append("E5")
        if "A04" in self.gun_head_dict:
            fill_cells.append("E2")
        if "B04" in self.gun_head_dict:
            fill_cells.append("E3")


        self.process_plate_stack(
        file_path=fr"{BASE_PATH1}/Materials-stack4列.xlsx",
        output_path=fr"{BASE_PATH}/Materials-stack4列.xlsx",
        fill_cells = fill_cells,
        plate_type='8孔',
        condition=self.post_data,

    )
    def _generate_transfer_data(self, hole_mapping: dict) -> Tuple[list, list, list]:
        """生成孔板转移数据"""
        transfer_records_12,transfer_records_24,transfer_records_2 = [],[],[]
        pipette_location = "A05"
        # 处理24孔板数据
        for plate_idx, (substance1, substance2, hole_id, amount) in enumerate(self.data_list_24, 1):
            if len(self.gun_head_dict[pipette_location]) >= 12:
                pipette_location = "B05"

            gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0
            target_volume = amount * 1000 / 2 if amount else 0
            if self.type == 1:
                if self.number_tasks == 1:
                    bar_code = "SS-A03"
                else:
                    bar_code = "SS-A02"
                record = {
                    "来源孔板条码": bar_code,
                    "来源孔板列Y": plate_idx,
                    "目标孔板条码": "FY24-B01",
                    "目标孔板列Y": hole_mapping[hole_id],
                    "移液量(ul)": target_volume,
                    "枪头库位": pipette_location,
                    "枪头列Y": gun_head + 1
                }

                if len(transfer_records_24) >= 12:
                    record["来源孔板条码"] = "SS-A03"
                    record['来源孔板列Y'] = plate_idx-12
                    transfer_records_2.append(record)
                else:
                    if len(transfer_records_24) >= 6:
                        record["来源孔板条码"] = "SS-A03"
                        record["来源孔板列Y"] = plate_idx - 6
                    transfer_records_24.append(record)

                if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                    self.gun_head_dict[pipette_location].append(gun_head + 1)
            else:
                record = {
                    "来源孔板条码": "SS-A03",
                    "来源孔板列Y": plate_idx + self.usage_24_hole,
                    "目标孔板条码": "FY24-B01",
                    "目标孔板列Y": hole_mapping[hole_id],
                    "移液量(ul)": target_volume,
                    "枪头库位": pipette_location,
                    "枪头列Y": gun_head + 1
                }

                if len(transfer_records_24) >= 18:
                    transfer_records_2.append(record)
                else:
                    if len(transfer_records_24) >= 6:
                        record["来源孔板条码"] = "SS-B02"
                        record["来源孔板列Y"] = plate_idx - 6
                    if len(transfer_records_24) >= 12:
                        record["来源孔板条码"] = "SS-B03"
                        record["来源孔板列Y"] = plate_idx - 12
                    transfer_records_24.append(record)

                if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                    self.gun_head_dict[pipette_location].append(gun_head + 1)
        # 处理12孔板数据
        num_12 = 0

        for col_idx, (substance, volumes) in enumerate(self.reagent_plates_12_.items(), 1):
            column_remaining = {col_idx: 18}  # 初始列剩余18000ul (18ml)
            current_col = self.data_list_12.index(substance) + 1  # 当前使用的列索引

            for volume in volumes:
                sub_volums = next((item[substance] for item in self.data_dict[volume] if substance in item), None)
                required_volume = sub_volums * 2  # 本次需要转移的总体积 (10ml = 10000ul)
                while required_volume > 0:
                    # 获取当前列的剩余体积（若不存在则初始化为18000）
                    current_remaining = column_remaining.get(current_col, 18)

                    # 当前列试剂不足时切换到新列
                    if current_remaining <= 0:
                        current_col += 1  # 移动到下一列
                        column_remaining[current_col] = 18  # 初始化新列剩余体积
                        continue  # 重新检查新列
                    # 计算本次实际转移量
                    transfer_vol = min(current_remaining, required_volume)
                    if len(self.gun_head_dict[pipette_location]) >= 12:
                        pipette_location = 'B05'
                    try:
                        gun_head = self.gun_head_dict[pipette_location][-1]
                    except:
                        gun_head = 0
                    """transfer_vol是四个枪头吸的量,分别加到两个孔里,所以每个孔两个枪头,每个枪头吸总量的1/4 """
                    target_volume = transfer_vol * 1000 / 4 if sub_volums else 0
                    '''self.data_list_12: 12孔板使用排序,
                       usage_12_hole:12孔板内标和淬灭剂使用情况
                       当前试剂在前n个代表肯定是第一个任务表的 '''
                    if substance in self.data_list_12[:self.usage_12_hole - 1]:
                        num_12 += 1
                        transfer_records_12.append({
                            "来源孔板条码": "R12-A02",
                            "来源孔板列Y": current_col,
                            "目标孔板条码": "FY24-B01",
                            "目标孔板列Y": hole_mapping[volume],
                            "移液量(ul)": target_volume,
                            "枪头库位": pipette_location,
                            "枪头列Y": gun_head + 1
                        })
                        if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                            self.gun_head_dict[pipette_location].append(gun_head + 1)
                    else:
                        '''改成24孔板'''
                        # record["来源孔板条码"] = "SS-A03"
                        # record['来源孔板列Y'] = plate_idx - 12
                        # current_col-12-len(transfer_records_2)
                        transfer_records_2.append({
                            "来源孔板条码": "SS-A03",
                            "来源孔板列Y": current_col-12,
                            "目标孔板条码": "FY24-B01",
                            "目标孔板列Y": hole_mapping[volume],
                            "移液量(ul)": target_volume,
                            "枪头库位": pipette_location,
                            "枪头列Y": gun_head + 1
                        })

                        if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                            self.gun_head_dict[pipette_location].append(gun_head + 1)
                    # 更新剩余体积
                    column_remaining[current_col] = current_remaining - transfer_vol
                    required_volume -= transfer_vol
        return transfer_records_12, transfer_records_24, transfer_records_2

    def _create_task_quencher(self):

        ExcelUtils.copy_workbook(
            f"{BASE_PATH1}/hiwo物料-8.xlsx",
            TEMPLATE_PATHS["mate_info2"]
        )
        ExcelUtils.copy_workbook(
            f"{BASE_PATH1}/hiwo任务11-8-24.xlsx",
            TEMPLATE_PATHS["wf11"]
        )
        # 淬灭剂
        quenchingAgentData = self.post_data['experimentLog3']["quenchingAgentData"]

        wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info2"])
        if quenchingAgentData:
            # 添加淬灭剂到物料表
            quenchingAgent = quenchingAgentData[0]["name"]
            quenchingAgent_Volume = int(float(quenchingAgentData[0]['addVolume']))

            quenchingAgent_Type = True
            if int(quenchingAgent_Volume) * (self.hole_number * 2) * 4 >= 18000:
                quenchingAgent_Type = None

                ExcelUtils.modify_cell(wb, "孔位信息", f'D{13}', quenchingAgent)
                # ExcelUtils.modify_cell(wb, "孔位信息", f'E{13}', 18000)
                ExcelUtils.modify_cell(wb, "孔位信息", f'D{12}', quenchingAgent)
                # ExcelUtils.modify_cell(wb, "孔位信息", f'E{12}', int(quenchingAgent_Volume)*8*4-18000)
            else:
                ExcelUtils.modify_cell(wb, "孔位信息", f'D{13}', quenchingAgent)
                # ExcelUtils.modify_cell(wb, "孔位信息", f'E{13}', int(quenchingAgent_Volume)*8*4)
            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info2"])

            if int(quenchingAgent_Volume) <= 200:
                pipette_location = 'A04'
            else:
                pipette_location = 'A05'
            if len(self.gun_head_dict[pipette_location]) >= 12:
                if pipette_location == 'A04':
                    pipette_location = 'B04'
                if pipette_location == 'A05':
                    pipette_location = 'B05'
            try:
                gun_head = self.gun_head_dict[pipette_location][-1]
            except:
                gun_head = 0
            """添加淬灭剂取液到wf11表格"""
            wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf11"])
            wells = ''
            for age in range(1, self.hole_number + 1):
                if quenchingAgent_Type:
                    ExcelUtils.modify_row(wb, "加淬灭液|||移液信息", age + 1,
                                          ["R12-A02", 12, "FY24-B01", age, int(quenchingAgent_Volume) / 2,
                                           pipette_location, gun_head + 1], start_col=1)
                else:
                    """ 当淬灭剂发在两个孔的时候，每个孔取两次"""
                    if age >= 2:vol = 12
                    else:vol = 11
                    ExcelUtils.modify_row(wb, "加淬灭液|||移液信息", age + 1,
                                          ["R12-A02", vol, "FY24-B01", age, int(quenchingAgent_Volume) / 2,
                                           pipette_location, gun_head + 1], start_col=1)
            if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                self.gun_head_dict[pipette_location].append(gun_head + 1)
            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf11"])

    def _create_task_internal_standard(self):
        # 内标数据
        internalStandardData = self.post_data['experimentLog3']["internalStandardData"]
        if not internalStandardData:
            return
        # 安全地处理淬灭剂孔位信息
        quenching_used_rows = []
        try:
            # 检查淬灭剂数据是否存在且非空
            if self.post_data['experimentLog3']["quenchingAgentData"]:
                quenchingAgent_Volume = int(float(self.post_data['experimentLog3']["quenchingAgentData"][0]['addVolume']))
                '''前端页面的剂量 * 反应的孔数(组*2) '''
                if int(quenchingAgent_Volume) * (self.hole_number * 2) >= 18000:
                    quenching_used_rows = [12, 13]
                else:
                    quenching_used_rows = [13]
        except (KeyError, IndexError):
            # 如果淬灭剂数据不存在或为空，则保持空列表
            quenching_used_rows = []

        internalStandard = internalStandardData[0]["name"]
        internalStandard_Volume = int(float(internalStandardData[0]['addVolume']))
        wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info2"])

        """quenching_used_rows[0]-1: 内标使用的最小孔位-1就是未使用的"""
        if quenching_used_rows:
            internal_standard_used_rows = quenching_used_rows[0] - 1
        else:
            internal_standard_used_rows = 13
        internalStandard_Type = True

        if int(internalStandard_Volume) * (self.hole_number * 2) >= 18000:
            internalStandard_Type = False
            ExcelUtils.modify_cell(wb, "孔位信息", f'D{internal_standard_used_rows}', internalStandard)
            # ExcelUtils.modify_cell(wb, "孔位信息", f'E{13}', 18000)
            ExcelUtils.modify_cell(wb, "孔位信息", f'D{internal_standard_used_rows - 1}', internalStandard)
            # ExcelUtils.modify_cell(wb, "孔位信息", f'E{12}', int(quenchingAgent_Volume)*8*4-18000)
        else:
            ExcelUtils.modify_cell(wb, "孔位信息", f'D{internal_standard_used_rows}', internalStandard)
            # ExcelUtils.modify_cell(wb, "孔位信息", f'E{13}', int(quenchingAgent_Volume)*8*4)
        ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info2"])

        if int(internalStandard_Volume) <= 200:
            pipette_location = 'A04'
        else:
            pipette_location = 'A05'
        if len(self.gun_head_dict[pipette_location]) >= 12:
            if pipette_location == 'A04':
                pipette_location = 'B04'
            if pipette_location == 'A05':
                pipette_location = 'B05'
        try:
            gun_head = self.gun_head_dict[pipette_location][-1]
        except:
            gun_head = 0
        """添加内标取液到wf11表格"""
        wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf11"])

        for age in range(1, self.hole_number + 1):
            if internalStandard_Type:
                ExcelUtils.modify_row(wb, "加内标|||移液信息", age + 1,
                                      ["R12-A02", internal_standard_used_rows - 1, "FY24-B01", age,
                                       int(internalStandard_Volume) / 2,
                                       pipette_location, gun_head + 1], start_col=1)
                self.usage_12_hole = internal_standard_used_rows - 1
            else:
                """ 当内标装在两个孔的时候，每个孔取两次 ，注意！！ internal_standard_used_rows值，是excel的行数，不是12孔板的行数，因为excel是第二行开始的，所以需要减一"""
                if age <= 2:
                    vol = internal_standard_used_rows - 1
                else:
                    vol = internal_standard_used_rows - 2
                ExcelUtils.modify_row(wb, "加内标|||移液信息", age + 1,
                                      ["R12-A02", vol, "FY24-B01", age, int(internalStandard_Volume) / 2,
                                       pipette_location, gun_head + 1], start_col=1)
                self.usage_12_hole = internal_standard_used_rows - 2
        if gun_head + 1 not in self.gun_head_dict[pipette_location]:
            self.gun_head_dict[pipette_location].append(gun_head + 1)
        ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf11"])

    def _create_task_extract(self):
        # 萃取液数据
        extractionData = self.post_data['experimentLog3']["extractionData"]
        if not extractionData:
            return
        # 安全地处理淬灭剂孔位信息
        internal_standard_used_rows = self.usage_12_hole
        internalStandard = extractionData[0]["name"]
        if internalStandard:
            internalStandard_Volume = 100
            wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info2"])
            """quenching_used_rows[0]-1: 内标使用的最小孔位-1就是未使用的"""
            # 添加萃取液到物料表
            ExcelUtils.modify_cell(wb, "孔位信息", f'D{internal_standard_used_rows}', internalStandard)
            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info2"])
            if int(internalStandard_Volume) <= 200:pipette_location = 'A04'
            else:pipette_location = 'A05'
            if len(self.gun_head_dict[pipette_location]) >= 12:
                if pipette_location == 'A04':pipette_location = 'B04'
                if pipette_location == 'A05':pipette_location = 'B05'
            try:gun_head = self.gun_head_dict[pipette_location][-1]
            except:gun_head = 0
            """添加萃取液到任务表格"""
            wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf11"])
            for age in range(1, self.hole_number + 1):
                try:
                    if self.post_data["experimentLog3"].get("productDilutionData", [])[0].get("procedure", {}).get(
                            "diluent3", ''):start = self.hole_number + 1 + age
                    else:start = 1 + age
                except:
                    start = 1 + age
                ExcelUtils.modify_row(wb, "加稀释液-过滤|||移液信息", start,
                                      ["R12-A02", internal_standard_used_rows - 1, "GL96-2A01", age, 100,
                                       pipette_location, gun_head + 1], start_col=1)
                self.usage_12_hole = internal_standard_used_rows - 1
            if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                self.gun_head_dict[pipette_location].append(gun_head + 1)
            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf11"])

    def _create_task_configuration(self) -> None:
        """创建任务配置Excel"""
        exp_log2 = self.post_data["experimentLog2"]
        exp_log3 = self.post_data["experimentLog3"]

        wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf11"])
        # 设置基本参数
        ExcelUtils.modify_cell(wb, "任务参数配置", "D3", "FY24-B01")
        # 温度
        if temp := exp_log2["ReactionConditions"].get("temperature"):
            ExcelUtils.modify_cell(wb, "任务参数配置", "E3", int(temp))
        # 反应时间
        if times := exp_log2["ReactionConditions"].get("time"):
            time = int(times)
            if temp:
                if int(temp)<=26:
                    pass
                elif int(temp)<=60 :
                    time+=3
                elif int(temp)<=100:
                    time+=6
                elif int(temp)<=160:
                    time+=10
                elif int(temp)<=200:
                    time+=18
            ExcelUtils.modify_cell(wb, "任务参数配置", "J3", int(time))
        # 密封
        seal = "关盖" if exp_log2["ReactionConditions"].get("seal") == "1" else "不关盖"
        ExcelUtils.modify_cell(wb, "任务参数配置", "H3", seal)
        ExcelUtils.modify_cell(wb, "任务参数配置", "P3", "开盖" if seal == "关盖" else "不开盖")
        # 转速
        if speed := exp_log2["ReactionConditions"].get("speed"):
            ExcelUtils.modify_cell(wb, "任务参数配置", "L3", int(speed))
        # 冷却时间
        if cool_time := exp_log3.get("coolingTime"):
            ExcelUtils.modify_cell(wb, "任务参数配置", "N3", int(cool_time))
        # 过滤时间
        if filter_data := exp_log3.get("filtrationData"):
            if filter_time := filter_data[0].get("filterTime"):
                ExcelUtils.modify_cell(wb, "任务参数配置", "AC3", int(filter_time))
        ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf11"])


    def write_codes_to_inner_hopper_stack(self,
            file_path: str,
            output_path: str,
            excel_list: list,
            excel_list_2
    ) -> None:
        """
        将 excel_list 中的 Code 按 Locate 位置写入 InnerHopperStack 的 B2:G5 区域，并标黄。

        :param file_path: 原始 Excel 文件路径
        :param output_path: 修改后保存的路径
        :param excel_list: 列表，每个元素为字典，包含 'Code', 'Locate', 'Type'
                           其中 Locate 为 1~24 的整数，对应 B2:G5 的行优先顺序
        """
        environment = self.post_data['experimentLog2']['ReactionConditions']['environment']
        wb = openpyxl.load_workbook(file_path)
        if environment =='2':
            table_name='HopperStack'
        else:
            if len(excel_list) >=5:
                table_name = 'HopperStack'
            else:
                table_name='TransferHopperStack'


        if table_name not in wb.sheetnames:
            raise ValueError(f"工作簿中不存在 {table_name} 子表")
        ws = wb[table_name]

        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # B2 对应 (2,2) —— 行2列2；G2 对应 (2,7)；B5 对应 (5,2)；G5 对应 (5,7)
        # Locate 编号从 1 开始，按行优先：第1行（row=2）列从2到7，第2行（row=3），等等
        all_items = excel_list + excel_list_2

        # 5×5 布局参数
        start_row = 2  # 起始行（第2行）
        start_col = 2  # 起始列（B列 → 列号2）
        cols_per_row = 5  # 每行5列（B~F）
        max_rows = 5  # 共5行（第2~6行）
        max_cells = cols_per_row * max_rows  # 25

        for idx, item in enumerate(all_items):
            if idx >= max_cells:
                print(f"警告：数据超过 {max_cells} 格，已截断")
                break

            row = start_row + idx // cols_per_row
            col = start_col + idx % cols_per_row
            code = item.get('Code')
            if code:
                ws.cell(row=row, column=col).value = code
                ws.cell(row=row, column=col).fill = yellow_fill

        # 如果总数据不足25格，剩余格子保持空白（或可根据需求填写其他内容）
        wb.save(output_path)
        print(f"处理完成，保存至: {output_path}")
    def set_hopper_b2_with_utils(self,file_path: str, output_path: str, value) -> None:
        """
        使用 ExcelUtils 工具类修改第二个子表 (HopperStack) 的 B2 单元格。

        :param file_path: 原始 Excel 文件路径
        :param output_path: 修改后保存的路径（建议另存）
        :param value: 要写入 B2 的值
        """
        # 打开工作簿
        wb = ExcelUtils.open_workbook(file_path)

        # 确定子表名称（优先使用 "HopperStack"，否则取第二个子表）
        if "HopperStack" in wb.sheetnames:
            sheet_name = "HopperStack"
        else:
            if len(wb.worksheets) < 2:
                raise ValueError("工作簿中不存在第二个子表")
            # 获取第二个子表的名称
            sheet_name = wb.worksheets[1].title

        # 修改单元格
        ExcelUtils.modify_cell(wb, sheet_name, "B2", value)

        # 保存
        ExcelUtils.save_workbook(wb, output_path)
        print(f"已将 {sheet_name} 的 B2 修改为 '{value}'，保存至: {output_path}")

    def _create_sowo_edit_stuff(self) -> None:
        """创建sowoEditStuff Excel并记录物料表号"""
        # 收集所有固态记录
        all_solid_records = []
        for reagent_name, hole_data in self.solidity.items():
            for hole_id, details in hole_data.items():
                all_solid_records.append({"reagent_name": reagent_name,"hole_id": hole_id})
        # 按试剂名称去重
        seen = set()
        unique_reagents = [x for x in all_solid_records if x["reagent_name"] not in seen and not seen.add(x["reagent_name"])]
        # 生成物料编码
        day_name = datetime.datetime.now().strftime("%Y%m%d")
        file_path = os.path.join(BASE_PATH1, f"物料key.json")
        data = {
            "daily_counters": {},  # 存储最新计数器
            "materials": OrderedDict()  # 有序字典保存试剂-编码映射
        }

        if os.path.exists(file_path):
            try:
                with open(file_path, 'r') as f:
                    data = json.load(f, object_pairs_hook=OrderedDict)
            except json.JSONDecodeError:
                pass  # 文件损坏时使用默认数据
        # 更新试剂编码映射

        # 初始化当天计数器
        if day_name not in data["daily_counters"]:
            data["daily_counters"][day_name] = 0
            counter = 0

        else:
            counter = data["daily_counters"][day_name]

        need_save = False

        for reagent in unique_reagents:
            reagent_name = reagent["reagent_name"]

            if reagent_name in data["materials"]:
                key = data["materials"][reagent_name]
                self.name_key_dict[reagent_name] = key

            # 如果试剂不存在则生成新编码
            else :
                counter += 1
                key = f"{day_name}-{str(counter).zfill(3)}"
                data["materials"][reagent_name] = key
                self.name_key_dict[reagent_name] = key
                need_save = True
        # 更新计数器并保存数据
        if need_save:
            data["daily_counters"][day_name] = counter
        with open(file_path, 'w') as f:
            json.dump(data, f, indent=2)  # 美观格式保存

        # 创建物料表记录并记录表号
        self.reagent_table_map = {}  # 新增：存储物料->表号的映射
        table_num = 1
        excel_list = []
        name_list = []
        # 第一张表处理
        s = 0
        for record in all_solid_records:
            if record["reagent_name"] in name_list:continue
            s += 1
            if s == 10:s += 1
            if s > 14:break
            excel_list.append({"Code": self.name_key_dict[record["reagent_name"]],"Locate": s,"Type": 2})
            self.reagent_table_map[record["reagent_name"]] = table_num
            name_list.append(record["reagent_name"])
        day_name = datetime.datetime.now().strftime("%Y%m%d")
        ExcelUtils.copy_workbook(f"{BASE_PATH1}/sowo物料.xlsx",TEMPLATE_PATHS["sowo"])
        # 保存第一张表
        wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["sowo"])
        ExcelUtils.write_records_to_sheet(wb, "Sheet1", excel_list)
        ExcelUtils.write_records_to_sheet(wb, "Sheet1", [{"Code": f"{day_name}-001", "Locate": 1, "Type": 1}])
        ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["sowo"])
        # 处理第二张表 (如果有超过13种物料)
        excel_list_2 = []
        if len(name_list) < len(unique_reagents):
            table_num = 2  # 表号递增

            s = 0
            for record in all_solid_records:
                if record["reagent_name"] in name_list:continue
                s += 1
                if s > 14:break
                excel_list_2.append({"Code": self.name_key_dict[record["reagent_name"]],"Locate": s,"Type": 2})
                self.reagent_table_map[record["reagent_name"]] = table_num
                name_list.append(record["reagent_name"])
            # 保存第二张表
            target_path = TEMPLATE_PATHS["sowo"].replace("1.xlsx", "2.xlsx")
            ExcelUtils.copy_workbook(f"{BASE_PATH1}/sowo物料.xlsx", target_path)
            wb = ExcelUtils.open_workbook(target_path)
            ExcelUtils.write_records_to_sheet(wb, "Sheet1", excel_list_2)
            ExcelUtils.write_records_to_sheet(wb, "Sheet1", [{"Code": f"{day_name}-001", "Locate": 1, "Type": 1}])
            ExcelUtils.save_workbook(wb, target_path)
        self.write_codes_to_inner_hopper_stack(fr"{BASE_PATH}/Materials-stack4列.xlsx",fr"{BASE_PATH}/Materials-stack4列.xlsx",excel_list,excel_list_2)
        # self.set_hopper_b2_with_utils(fr"{BASE_PATH}/Materials-stack4列.xlsx",fr"{BASE_PATH}/Materials-stack4列.xlsx",f"{day_name}-001")

    def _create_mate_msg(self) -> None:
        """创建mateMsg Excel并根据物料表号分表"""
        hole_to_coordinate = {
            1: (1, 1), 2: (1, 2), 3: (1, 3), 4: (1, 4),
            5: (2, 1), 6: (2, 2), 7: (2, 3), 8: (2, 4)
        }
        # 创建反应ID到孔位映射
        reaction_to_hole = {}
        for group_index, group in enumerate(self.perforated_plate):
            reaction_to_hole[group[0]] = group_index + 1
            if len(group) > 1:  # 处理成对的情况
                reaction_to_hole[group[1]] = group_index + 5

        # 准备数据行（按表号分组）
        table_data = {1: [], 2: []}  # 存储分表数据
        for reagent_name, hole_data in self.solidity.items():
            for hole_id, details in hole_data.items():
                hole_str = str(hole_id)
                if hole_str in reaction_to_hole:
                    actual_hole = reaction_to_hole[hole_str]
                    x, y = hole_to_coordinate[actual_hole]
                    clean_name = reagent_name.strip()
                    target = details["weight"] if details["weight"] not in [None, ""] else 0
                    # 获取物料所属表号
                    table_num = self.reagent_table_map.get(reagent_name, 1)
                    table_data[table_num].append({
                        "PointX": y,
                        "PointY": x,
                        "UsePlateType": 0,
                        "Name": clean_name,
                        "Target": round(float(target), 1),
                        "Tolerance": 0.5,
                        "UseMode": 0,
                        "Code": self.name_key_dict[reagent_name]
                    })

        # 分别保存每个表
        create_True = 0
        for table_num, data_rows in table_data.items():
            if not data_rows:  # 跳过空表
                create_True += 1
                continue
            # 创建DataFrame并保存
            df = pd.DataFrame(data_rows, columns=[
                "PointX", "PointY", "UsePlateType", "Name",
                "Target", "Tolerance", "UseMode", "Code"
            ])

            df = df.sort_values(by="Name").reset_index(drop=True)
            df['PlaterName'] = None
            # 仅第一行（索引0）写入 '96孔'
            if len(df) > 0:
                df.at[0, 'PlaterName'] = '8孔'
            if self.type == 1:
                # 根据表号生成文件名
                if table_num == 1:
                    output_path = TEMPLATE_PATHS["mate"]
                else:
                    output_path = TEMPLATE_PATHS["mate"].replace("1.xlsx", "2.xlsx")
            else:
                if table_num == 1:
                    output_path = crystallize_TEMPLATE_PATHS["mate"]
                else:
                    output_path = crystallize_TEMPLATE_PATHS["mate"].replace("1.xlsx", "2.xlsx")
            df.to_excel(output_path, index=False)
        if create_True == 2 and self.type == 1:
            df = pd.DataFrame(columns=[
                "PointX", "PointY", "UsePlateType", "Name",
                "Target", "Tolerance", "UseMode", "Code"
            ])
            df['PlaterName'] = None
            # 仅第一行（索引0）写入 '96孔'
            if len(df) > 0:
                df.at[0, 'PlaterName'] = '8孔'
            df.to_excel(TEMPLATE_PATHS["mate"], index=False)
        elif create_True == 2 and self.type == 2:
            df = pd.DataFrame(columns=[
                "PointX", "PointY", "UsePlateType", "Name",
                "Target", "Tolerance", "UseMode", "Code"
            ])
            df['PlaterName'] = None
            # 仅第一行（索引0）写入 '96孔'
            if len(df) > 0:
                df.at[0, 'PlaterName'] = '8孔'
            df.to_excel(crystallize_TEMPLATE_PATHS["mate"], index=False)

    def _create_dilution_data(self) -> None:
        """创建稀释数据"""
        product_dilution_data = self.post_data["experimentLog3"].get("productDilutionData", [])

        for data in product_dilution_data:
            wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf11"])
            procedure = data.get("procedure", {})

            self._process_diluent(wb, "加稀释液-反应|||移液信息", procedure.get("diluent1"), True)
            self._process_diluent(wb, "加稀释液-中转|||移液信息", procedure.get("diluent2"), True, "ZZ-B02")
            self._process_diluent(wb, "加稀释液-过滤|||移液信息", procedure.get("diluent3"), False, "GL96-2A01")

            self._process_mixing(wb, procedure.get("diluent4"), procedure.get("diluent5"))

            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf11"])

    def _process_diluent(self, wb: Workbook, sheet_name: str, diluent: Optional[str], judge, kk="FY24-B01") -> None:
        """处理稀释剂"""
        if not diluent:
            return
        if judge:
            diluent =  int(float(diluent))/ 2
        else:
            diluent = int(float(diluent))
        pipette_location = "A04" if int(diluent) <= 200 else "A05"
        if len(self.gun_head_dict[pipette_location]) >= 12:
            pipette_location = "B04" if pipette_location == "A04" else "B05"

        gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0
        num = 0
        for i in range(2, self.hole_number + 2):
            num += 1
            ExcelUtils.modify_row(wb, sheet_name, i,
                                  ["R1-B03", 1, kk, num, int(diluent), pipette_location, gun_head + 1])
        if gun_head + 1 not in self.gun_head_dict[pipette_location]:
            self.gun_head_dict[pipette_location].append(gun_head + 1)

    def _process_mixing(self, wb: Workbook, diluent4: Optional[str], diluent5: Optional[str]) -> None:
        """处理混合"""
        if not diluent4 or not diluent5:
            return
        diluent4 = int(float(diluent4))
        diluent5 = int(float(diluent5))
        pipette_location = "A04" if int(diluent4) <= 200 else "A05"
        if len(self.gun_head_dict[pipette_location]) >= 12:
            pipette_location = "B04" if pipette_location == "A04" else "B05"
        num = 0
        for i in range(2, (self.hole_number + 1) * 2, 2):
            num += 1
            gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0
            ExcelUtils.modify_row(wb, '反应板混匀|||移液信息', i,
                                  ["FY24-B01", num, "ZZ-B02", num, int(diluent4) / 2, pipette_location, gun_head + 1])
            ExcelUtils.modify_row(wb, '反应板混匀|||移液信息', i + 1,
                                  ["ZZ-B02", num, "GL96-2A01", num, int(diluent5), pipette_location, gun_head + 1])

            if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                self.gun_head_dict[pipette_location].append(gun_head + 1)



def sort_list(list1: list, list2: list) -> Tuple[list, list]:
    """排序两个列表"""
    # 第一级：完全相同的元素
    common_elements = [item for item in list1 if item in list2]

    # 第二级：值相同但SMILES不同的元素
    same_value_diff_smiles = []
    for item1 in list1:
        if item1 not in common_elements:
            for item2 in list2:
                if item2 not in common_elements:
                    if list(item1.values())[0] == list(item2.values())[0]:
                        if item1 not in same_value_diff_smiles:
                            same_value_diff_smiles.append(item1)
                        if item2 not in same_value_diff_smiles:
                            same_value_diff_smiles.append(item2)

    # 第三级：剩余元素
    remaining1 = [item for item in list1
                  if item not in common_elements and item not in same_value_diff_smiles]
    remaining2 = [item for item in list2
                  if item not in common_elements and item not in same_value_diff_smiles]

    # 重新排序
    new_list1 = common_elements + [item for item in same_value_diff_smiles if item in list1] + remaining1
    new_list2 = common_elements + [item for item in same_value_diff_smiles if item in list2] + remaining2

    return new_list1, new_list2


def main(id,post_data: dict,source) -> None:

    """主函数"""
    if source ==1:
        ExcelUtils.clear_directory_(BASE_PATH)

    else:
        ExcelUtils.clear_directory_(crystallize_BASE_PATH)
    # 初始化文件

    # 处理反应
    processor = ReactionProcessor(id,post_data)
    processor.process(source)

    if source ==1:
        # ExcelUtils._repair_with_libreoffice(BASE_PATH)
        ExcelUtils.batch_repair_folder(BASE_PATH, backup=False)
    else:
        # ExcelUtils._repair_with_libreoffice(crystallize_BASE_PATH)
        ExcelUtils.batch_repair_folder(crystallize_BASE_PATH, backup=False)

if __name__ == "__main__":
    ddd = {
    "experimentLog1": {
        "wellPlates": "8孔板20ml",
        "SubstratesTableList": [
            {
                "rowId": "row_1776652981356_77",
                "id": 1,
                "type": "sub",
                "smiles": "O=C(Cl)Cc1ccccc1",
                "substance": "O=C(Cl)Cc1ccccc1",
                "molecularFormula": "C8H7ClO",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 154.6,
                "reactionMoles": 1,
                "reactionQuality": 154.6,
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "rowId": "row_1776652981356_78",
                "id": 1,
                "type": "sub",
                "smiles": "N[C@@H](CSC(c1ccccc1)(c1ccccc1)c1ccccc1)C(=O)O",
                "substance": "N[C@@H](CSC(c1ccccc1)(c1ccccc1)c1ccccc1)C(=O)O",
                "molecularFormula": "C22H21NO2S",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 363.48,
                "reactionMoles": 1,
                "reactionQuality": 363.48,
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "rowId": "row_1776652981356_79",
                "id": 1,
                "type": "reagent",
                "smiles": "",
                "substance": "试剂1,copper(I) bromide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "0.1",
                "solvent": "231",
                "reactionMoles": "0.10",
                "reactionQuality": "23.10",
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "rowId": "row_1776652981356_80",
                "id": 1,
                "type": "solvent",
                "smiles": "",
                "substance": "溶剂1,tetrahydrofuran",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 16,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": None,
                "singleCockAddVolume": "16",
                "cas": "",
                "intensity": None
            },
            {
                "rowId": "row_1776652981356_81",
                "id": 2,
                "type": "sub",
                "smiles": "O=C(Cl)Cc1ccccc1",
                "substance": "O=C(Cl)Cc1ccccc1",
                "molecularFormula": "C8H7ClO",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 154.6,
                "reactionMoles": 1,
                "reactionQuality": 154.6,
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "rowId": "row_1776652981356_82",
                "id": 2,
                "type": "sub",
                "smiles": "N[C@@H](CSC(c1ccccc1)(c1ccccc1)c1ccccc1)C(=O)O",
                "substance": "N[C@@H](CSC(c1ccccc1)(c1ccccc1)c1ccccc1)C(=O)O",
                "molecularFormula": "C22H21NO2S",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 363.48,
                "reactionMoles": 1,
                "reactionQuality": 363.48,
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "rowId": "row_1776652981356_83",
                "id": 2,
                "type": "reagent",
                "smiles": "",
                "substance": "试剂1,Sodium bicarbonate",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2323",
                "reactionMoles": "1.00",
                "reactionQuality": "2323.00",
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "rowId": "row_1776652981356_84",
                "id": 2,
                "type": "reagent",
                "smiles": "",
                "substance": "NaCl",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.5",
                "solvent": "122",
                "reactionMoles": "1.50",
                "reactionQuality": "183.00",
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "rowId": "row_1776652981356_85",
                "id": 2,
                "type": "reagent",
                "smiles": "",
                "substance": "试剂3,Sodium hydroxide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "122",
                "reactionMoles": "1.00",
                "reactionQuality": "122.00",
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "rowId": "row_1776652981356_86",
                "id": 2,
                "type": "solvent",
                "smiles": "",
                "substance": "溶剂1,Benzene",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 16,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": None,
                "singleCockAddVolume": "7.00",
                "cas": "",
                "intensity": None
            },
            {
                "rowId": "row_1776652981356_87",
                "id": 2,
                "type": "solvent",
                "smiles": "",
                "substance": "溶剂2,Water",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 16,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": None,
                "singleCockAddVolume": "3",
                "cas": "",
                "intensity": None
            },
            {
                "rowId": "row_1776652981356_88",
                "id": 2,
                "type": "solvent",
                "smiles": "",
                "substance": "溶剂3,Benzene",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 16,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": None,
                "singleCockAddVolume": "3",
                "cas": "",
                "intensity": None
            },
            {
                "rowId": "row_1776652981356_89",
                "id": 2,
                "type": "solvent",
                "smiles": "",
                "substance": "溶剂4,Water",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 16,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": None,
                "singleCockAddVolume": "3",
                "cas": "",
                "intensity": None
            }
        ],
        "MaterialsTableList": [
            {
                "type": "sub",
                "substance": "O=C(Cl)Cc1ccccc1",
                "smiles": "O=C(Cl)Cc1ccccc1",
                "solvent": 154.6,
                "molecularFormula": "C8H7ClO",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "309.20",
                "theoreticalVolume": "0.00",
                "configurationMoles": "2.60",
                "cas": "",
                "intensity": None,
                "id": 1,
                "configurationVolume": "0.00",
                "configurationQuality": "401.96",
                "concentration": ""
            },
            {
                "type": "sub",
                "substance": "N[C@@H](CSC(c1ccccc1)(c1ccccc1)c1ccccc1)C(=O)O",
                "smiles": "N[C@@H](CSC(c1ccccc1)(c1ccccc1)c1ccccc1)C(=O)O",
                "solvent": 363.48,
                "molecularFormula": "C22H21NO2S",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "726.96",
                "theoreticalVolume": "0.00",
                "configurationMoles": "2.60",
                "cas": "",
                "intensity": None,
                "id": 2,
                "configurationVolume": "0.00",
                "configurationQuality": "945.05",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂1,copper(I) bromide",
                "smiles": "",
                "solvent": "231",
                "molecularFormula": "",
                "theoreticalMoles": "0.10",
                "theoreticalQuality": "23.10",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.13",
                "cas": "",
                "intensity": None,
                "id": 3,
                "configurationVolume": "0.00",
                "configurationQuality": "30.03",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "溶剂1,tetrahydrofuran",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "",
                "theoreticalQuality": "",
                "theoreticalVolume": "16.00",
                "configurationMoles": "",
                "cas": "",
                "intensity": None,
                "id": 4,
                "configurationVolume": "20.80",
                "configurationQuality": "",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂1,Sodium bicarbonate",
                "smiles": "",
                "solvent": "2323",
                "molecularFormula": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "2323.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "1.30",
                "cas": "",
                "intensity": None,
                "id": 5,
                "configurationVolume": "0.00",
                "configurationQuality": "3019.90",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "NaCl",
                "smiles": "",
                "solvent": "122",
                "molecularFormula": "",
                "theoreticalMoles": "1.50",
                "theoreticalQuality": "183.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "1.95",
                "cas": "",
                "intensity": None,
                "id": 6,
                "configurationVolume": "0.00",
                "configurationQuality": "237.90",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂3,Sodium hydroxide",
                "smiles": "",
                "solvent": "122",
                "molecularFormula": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "122.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "1.30",
                "cas": "",
                "intensity": None,
                "id": 7,
                "configurationVolume": "0.00",
                "configurationQuality": "158.60",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "溶剂1,Benzene",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "",
                "theoreticalQuality": "",
                "theoreticalVolume": "7.00",
                "configurationMoles": "",
                "cas": "",
                "intensity": None,
                "id": 8,
                "configurationVolume": "9.10",
                "configurationQuality": "",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "溶剂2,Water",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "",
                "theoreticalQuality": "",
                "theoreticalVolume": "3.00",
                "configurationMoles": "",
                "cas": "",
                "intensity": None,
                "id": 9,
                "configurationVolume": "5.00",
                "configurationQuality": "",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "溶剂3,Benzene",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "",
                "theoreticalQuality": "",
                "theoreticalVolume": "3.00",
                "configurationMoles": "",
                "cas": "",
                "intensity": None,
                "id": 10,
                "configurationVolume": "5.00",
                "configurationQuality": "",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "溶剂4,Water",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "",
                "theoreticalQuality": "",
                "theoreticalVolume": "3.00",
                "configurationMoles": "",
                "cas": "",
                "intensity": None,
                "id": 11,
                "configurationVolume": "5.00",
                "configurationQuality": "",
                "concentration": ""
            }
        ],
        "SynthesisMolReadList": [],
        "reactionParams": []
    },
    "experimentLog2": {
        "ReactionConditions": {
            "temperature": "100",
            "time": "30",
            "speed": "400",
            "seal": "1",
            "environment": "1"
        },
        "lightData": {
            "type": "1",
            "typeList": [
                {
                    "label": "照射方式1",
                    "value": "1"
                },
                {
                    "label": "照射方式2",
                    "value": "2"
                }
            ],
            "list": [
                {
                    "status": "1",
                    "WattageStart": "",
                    "WattageType": "",
                    "time": [
                        {
                            "start": "",
                            "end": ""
                        }
                    ]
                },
                {
                    "status": "1",
                    "WattageStart": "",
                    "WattageType": "",
                    "time": [
                        {
                            "start": "",
                            "end": ""
                        }
                    ]
                },
                {
                    "status": "1",
                    "WattageStart": "",
                    "WattageType": "",
                    "time": [
                        {
                            "start": "",
                            "end": ""
                        }
                    ]
                },
                {
                    "status": "1",
                    "WattageStart": "",
                    "WattageType": "",
                    "time": [
                        {
                            "start": "",
                            "end": ""
                        }
                    ]
                }
            ]
        },
        "electricityData": {
            "type": "1",
            "list": []
        },
        "gasData": {}
    },
    "experimentLog3": {
        "coolingTime": "50",
        "quenchingAgentData": [],
        "internalStandardData": [],
        "productDilutionData": [],
        "extractionData": [
            {
                "name": "",
                "volume": "",
                "time": "",
                "times": "",
                "type": "1"
            }
        ],
        "filtrationData": [
            {
                "name": "",
                "filterTime": ""
            }
        ]
    },
    "experimentLog4": {
        "instrumentName": "",
        "instrumentInstruments": [],
        "NMR": {
            "analyseArr": [
                {
                    "name": "H NMR",
                    "status": False
                },
                {
                    "name": "C NMR",
                    "status": False
                },
                {
                    "name": "F NMR",
                    "status": False
                },
                {
                    "name": "P NMR",
                    "status": False
                },
                {
                    "name": "COSY",
                    "status": False
                },
                {
                    "name": "HSQC",
                    "status": False
                },
                {
                    "name": "HMBC",
                    "status": False
                },
                {
                    "name": "NOESY",
                    "status": False
                }
            ],
            "solvent": {
                "name": "",
                "volume": "",
                "temperature": "",
                "frequency": ""
            }
        },
        "MS": [],
        "LC": [],
        "LCMS": [],
        "GC": [],
        "UV": [
            {
                "wavelengthStart": "",
                "wavelengthEnd": "",
                "velocity": "",
                "width": "",
                "solvent": "",
                "concentration": ""
            }
        ],
        "IR": [
            {
                "resolutionStart": "",
                "resolutionEnd": "",
                "scanNumber": "",
                "Wavenumber": ""
            }
        ],
        "XRD": [
            {
                "startAngle": "",
                "endAngle": "",
                "stepLength": "",
                "scanSpeed": "",
                "scanDuration": "",
                "mode": ""
            }
        ]
    }
}
    main('b606f8d9f73d4153bb532ff297533814',ddd,1)



'''

反溶剂反加的剂量是不是还是按两个孔最大的算，是单枪头还是双枪头
开关盖结晶只有一个的问题
8孔 ：2*4
12kon
24孔：4*6 10ml
48孔：6*8
96孔：8*12
pkill -f excel_api.py
netstat -tulnp | grep 10014
nohup python excel_api.py &
'''