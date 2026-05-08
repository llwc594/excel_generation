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
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
import aspose.cells as ac
# 常量定义
# BASE_PATH = r"D:/api/python/表格生成/逆合成"
# BASE_PATH1 = r"D:/api/python/模板/逆合成"
# crystallize_BASE_PATH = r"D:/api/python/表格生成/结晶平台"
# crystallize_BASE_PATH1 = r"D:/api/python/模板/结晶平台"
#
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
    "sowo": f"{BASE_PATH}/sowo-物料1-1.xlsx",
    "mate": f"{BASE_PATH}/sowo-任务1-1.xlsx",
    "information": f"{BASE_PATH}/信息表格.xlsx",
    "mate_info1": f"{BASE_PATH}/hiwo-物料1.xlsx",
    "mate_info2": f"{BASE_PATH}/hiwo-物料2.xlsx"

}
crystallize_TEMPLATE_PATHS = {
    "wf11": f"{crystallize_BASE_PATH}/hiwo-任务1-1.xlsx",
    "wf05": f"{crystallize_BASE_PATH}/hiwo-任务1-2.xlsx",
    "sowo": f"{crystallize_BASE_PATH}/sowo-物料1-1.xlsx",
    "mate": f"{crystallize_BASE_PATH}/sowo-任务1-1.xlsx",
    "mate_info1": f"{crystallize_BASE_PATH}/hiwo-物料1-1.xlsx",
    "mate_info2": f"{crystallize_BASE_PATH}/hiwo-物料1-2.xlsx",
    "information":f"{crystallize_BASE_PATH}/信息表格.xlsx",
    "WF10":f"{crystallize_BASE_PATH1}/hiwo任务15-8-24.xlsx",
    "WF11": f"{crystallize_BASE_PATH1}/hiwo物料-24-01.xlsx"
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
        # sheet = workbook[sheet_name]
        # if isinstance(cell_ref, str):
        #     sheet[cell_ref].value = new_value or "空"
        # else:
        #     sheet.cell(row=cell_ref[0], column=cell_ref[1]).value = new_value or "空"
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
                else:
                    sheet.cell(row=row_num, column=col).value = value
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

            # 获取所有Excel文件
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
            if app is not None:
                app.quit()

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
        self.reagent_hole = {}
        #记录24孔的相同试剂使用位置
        self.reagent_plates_24_ = {}
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
        # 结晶平台的反溶剂
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
        """计算反应相似度并分组 self.data_dict"""

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
                #sowo物料
                self._create_excel_ttt(26)
                self._create_excel_information(26, 7, TEMPLATE_PATHS["information"])
            else:
                self._create_excel_xxx(TEMPLATE_PATHS["mate_info1"], TEMPLATE_PATHS["information"], 1)
                self._create_excel_xxx(TEMPLATE_PATHS["mate_info2"], TEMPLATE_PATHS["information"], 28)
                # sowo物料
                self._create_excel_ttt(60)
                self._create_excel_information(60, 7, TEMPLATE_PATHS["information"])
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

        for col_idx, letter in enumerate("ABCD", start=start_col):
            ws.cell(row=start_row, column=col_idx, value=letter).font = Font(bold=True)

        # 添加行标签 (1,2,3...)
        for row_idx in range(start_row + 1, start_row + 7):
            ws.cell(row=row_idx, column=start_col - 1, value=(start_row + 7) - row_idx).font = Font(bold=True)
        thin_border = Border(left=Side(style='thin'),
                             right=Side(style='thin'),
                             top=Side(style='thin'),
                             bottom=Side(style='thin'))

        # 写入数据
        for row_idx, row_data in enumerate(self.perforated_plate, start=1):
            for col_idx, cell_value in enumerate(row_data, start=start_col):
                cell = ws.cell(row=start_row + 7 - row_idx, column=col_idx, value=cell_value)
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

    def process_reaction_conditions(self,reaction_conditions):
        """
        处理反应条件数据，按试剂组合分组
        """
        # 将每个反应的条件列表转换为标准化的字符串键
        reaction_groups = defaultdict(list)

        for reaction_id, conditions in reaction_conditions.items():
            # 提取试剂名称和用量
            reagents = []
            for condition in conditions:
                for reagent, amount in condition.items():
                    reagents.append((reagent, amount))

            # 按试剂名称排序，确保相同组合有相同的键
            reagents.sort(key=lambda x: x[0])

            # 创建标准化键
            key_parts = []
            for reagent, amount in reagents:
                # 使用固定格式的字符串表示试剂组合
                key_parts.append(f"{reagent}:{amount}")

            group_key = "|".join(key_parts)
            reaction_groups[group_key].append(reaction_id)

        return reaction_groups

    def assign_to_plate(self,reaction_groups, rows=6, cols=4):
        """
        将反应分配到孔板，优先将相同条件的反应放在同一行
        """
        # 按组大小排序，优先处理大组
        sorted_groups = sorted(reaction_groups.items(),
                               key=lambda x: (len(x[1]), x[0]),
                               reverse=True)

        # 初始化孔板
        plate = [[None for _ in range(cols)] for _ in range(rows)]
        current_row = 0
        current_col = 0

        # 分配策略：尽量将同一组的反应放在同一行
        for group_key, reaction_ids in sorted_groups:
            # 如果当前行已满，换到下一行
            if current_col >= cols and current_row < rows - 1:
                current_row += 1
                current_col = 0

            # 如果已经填满所有行，停止分配
            if current_row >= rows:
                break

            # 计算这组反应需要多少行
            num_reactions = len(reaction_ids)
            rows_needed = (num_reactions + cols - 1) // cols  # 向上取整

            # 检查是否有足够的连续行
            available_rows = rows - current_row
            if rows_needed > available_rows:
                # 如果没有足够的行，调整到下一组有空位的地方
                # 重新从下一行开始寻找空位
                for row in range(current_row, rows):
                    empty_cols = [col for col, cell in enumerate(plate[row]) if cell is None]
                    if len(empty_cols) >= min(cols, num_reactions):
                        # 在这一行开始放置
                        current_row = row
                        current_col = 0
                        for col in range(cols):
                            if plate[current_row][col] is not None:
                                current_col = col + 1
                        break
                else:
                    # 如果没有找到合适的行，跳过这组反应
                    print(f"警告: 无法为组 {group_key[:20]}... 分配所有反应")
                    continue

            # 分配这组反应到孔板
            for i, reaction_id in enumerate(reaction_ids):
                # 如果当前列已满，换到下一行
                if current_col >= cols:
                    current_row += 1
                    current_col = 0

                # 如果超出了孔板范围，停止分配
                if current_row >= rows:
                    break

                # 分配反应到当前孔位
                plate[current_row][current_col] = reaction_id
                current_col += 1

        return plate

    def compress_plate_layout(self,plate):
        """
        压缩孔板布局，移除空行和每行的空孔位
        """
        compressed_plate = []
        for row in plate:
            # 只保留非空的孔位
            compressed_row = [cell for cell in row if cell != None]
            # 如果这一行有数据，则添加到压缩后的孔板中
            if compressed_row:
                compressed_plate.append(compressed_row)
        return compressed_plate

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
        if self.type == 1:
            json_path = BASE_PATH

        else:
            json_path = crystallize_BASE_PATH
        json_name = '孔位设置参数.json'
        json_path = os.path.join(json_path, json_name)
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(hole_json, f, ensure_ascii=False)

    def _calculate_similarity(self) -> None:
        temp_group = []
        for list_id in self.data_dict.keys():
            temp_group.append(list_id)
            if len(temp_group)==4:
                self.perforated_plate.append(temp_group)
                temp_group = []
        if temp_group:
            self.perforated_plate.append(temp_group)
        # 准备数据集
        # reaction_groups = self.process_reaction_conditions(self.data_dict)
        # plate = self.assign_to_plate(reaction_groups, rows=6, cols=4)
        # plate = self.compress_plate_layout(plate)
        # for plate in plate:
        #     self.perforated_plate.append(plate)

        #生成孔位设置json文件
        self._hole_json()
        for perforated in self.perforated_plate:
            group_size = len(perforated)  # 获取当前组的大小
            # 获取当前组所有反应的化学列表
            chem_lists = [self.data_dict[f"{pid}"] for pid in perforated if pid]
            max_length = max(len(chem_list) for chem_list in chem_lists)
            llwc_num = 1
            for i in range(max_length):
                row_items = []
                for j in range(group_size):
                    chem_list = chem_lists[j]
                    item = chem_list[i] if i < len(chem_list) else None
                    row_items.append(item)

                # 补齐到4个位置（用None填充空位）
                while len(row_items) < 4:
                    row_items.append(None)
                all_keys = [list(d.keys())[0] for d in row_items if d ]  # 提取每个字典的第一个键（每个字典只有一个键）
                all_same = all(key == all_keys[0] for key in all_keys)
                # if all_same and len(all_keys)==4:
                #     #在这里获取相同试剂的行数
                #     if all_keys[0] in list(self.reagent_plates_24_.keys()):
                #         self.reagent_plates_24_[f'{all_keys[0]}'].append(perforated[0])
                #     else:
                #         self.reagent_plates_24_.update({all_keys[0]: [perforated[0]]})

                # 添加组ID和位置编号
                self.reagent_plates_24.append(row_items + [perforated[0], llwc_num])
                llwc_num += 4  # 48孔板每行6个位置
            # else:
            #     chem_list = self.data_dict[f"{perforated[0]}"]
            #     llwc_num = 1
            #     for chem_dict in chem_list:
            #         chem = list(chem_dict.keys())[0]
            #         vol = list(chem_dict.values())[0]
            #
            #         # 单个反应在24孔板中独占一个位置
            #         self.reagent_plates_24.append([
            #             {chem: vol},  # 左侧化学物质
            #             None,
            #             None,
            #             None,# 右侧为空
            #             perforated[0],  # 孔位ID
            #             llwc_num  # 位置编号
            #         ])
            #         llwc_num += 4
            '''
            24孔逻辑：都是24孔板：
                1.先分配孔位
                2.根据孔位确定物料的位置
                2.1：判断同横排的试剂是否相同，
            '''

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

    @staticmethod
    def _calculate_similarity_four_sets_core(chemical_sets: dict, lengths: dict, keys: list) -> dict:
        """计算四个列表的核心相似度（四个集合的交集）"""
        if len(keys) != 4:
            raise ValueError("需要提供4个键")

        # 获取四个集合
        sets = [chemical_sets[k] for k in keys]

        # 计算四个集合的交集
        intersection_all = set.intersection(*sets)
        union_all = set.union(*sets)

        # 计算基于四个集合的Jaccard相似度
        jaccard_all = len(intersection_all) / len(union_all) if union_all else 0

        # 计算长度相似度（使用四个长度的最小值和最大值）
        lens = [lengths[k] for k in keys]
        length_sim_all = min(lens) / max(lens) if max(lens) > 0 else 0

        # 计算加权相似度
        combined_all = 0.5 * jaccard_all + 0.5 * length_sim_all

        # 提取共同名称
        common_names_all = {identifier.rsplit('_', 1)[0] for identifier in intersection_all}

        return {
            'jaccard_all': jaccard_all,
            'length_sim_all': length_sim_all,
            'combined_all': combined_all,
            'intersection_all': len(intersection_all),
            'common_names_all': common_names_all,
            'union_all': len(union_all)
        }
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
            if rows == 1:
                wb = Workbook()
            else:
                wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["information"])
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
            # 孔位1数据
            for row_idx, row in hole_data_1.iterrows():
                for col_idx, value in enumerate(row, start=2):
                    ws.cell(row=row_idx + rows + 1, column=col_idx, value=value)

            # 写入右侧数据(第3行物料+孔位2)
            right_start_col = len(material_data.columns) + 7
            # 物料第3行数据
            for col_idx, value in enumerate(material_data.iloc[1], start=right_start_col):
                ws.cell(row=rows, column=col_idx, value=value)

            if rows == 1:
                if self.number_tasks == 1:
                    row_ = -22
                else:
                    row_ = -22
            else:
                row_ = rows - 23

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
            # 孔位1数据
            for row_idx, row in hole_data_1.iterrows():
                for col_idx, value in enumerate(row, start=2):
                    ws.cell(row=row_idx + rows + 1, column=col_idx, value=value)

            # 写入右侧数据(第3行物料+孔位2)

            # 物料第3行数据
            for col_idx, value in enumerate(material_data.iloc[1], start=7):
                ws.cell(row=rows, column=col_idx, value=value)

            if rows == 1:
                if self.number_tasks == 1:
                    row_ = -22
                else:
                    row_ = -22
            else:
                row_ = rows - 11
            # 孔位2数据
            for row_idx, row in hole_data_2.iterrows():
                for col_idx, value in enumerate(row, start=7):
                    ws.cell(row=row_idx + row_, column=col_idx, value=value)

            for col_idx, value in enumerate(material_data.iloc[2], start=12):
                ws.cell(row=rows, column=col_idx, value=value)
            # 孔位1数据
            for row_idx, row in hole_data_3.iterrows():
                for col_idx, value in enumerate(row, start=12):
                    ws.cell(row=row_idx + rows - 46, column=col_idx, value=value)

            for col_idx, value in enumerate(material_data.iloc[3], start=17):
                ws.cell(row=rows, column=col_idx, value=value)
            # 孔位1数据
            for row_idx, row in hole_data_4.iterrows():
                for col_idx, value in enumerate(row, start=17):
                    ws.cell(row=row_idx + rows - 70, column=col_idx, value=value)

            # 保存结果
            wb.save(output_path)
            print(f"处理完成，结果已保存到: {output_path}")

    def _create_excel_xxxx(self, input_path, output_path, rows):
        if rows == 1:
            wb = Workbook()
        else:
            wb = ExcelUtils.open_workbook(output_path)
        ws = wb['Sheet']
        # 读取Excel文件
        df_material = pd.read_excel(input_path, sheet_name='物料信息')
        df_holes = pd.read_excel(input_path, sheet_name='孔位信息')

        # 获取物料表第2行和第3行的第2和第8列数据
        material_data = df_material.iloc[1:3, [1, 7]].reset_index(drop=True)

        # 获取孔位表第一列值为1和2的数据(1-6列)
        hole_data_1 = df_holes[df_holes.iloc[:, 0] == 1].iloc[:, [0,1,3,5]]


        # 写入任务标签
        if rows == 2:
            # 合并单元格并写入"任务一"
            ws.merge_cells(start_row=rows + 36, end_row=61, start_column=6, end_column=6)
            ws.cell(row=rows + 36, column=6, value="反溶剂添加")
            alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=rows + 36, column=6).alignment = alignment

        # 写入左侧数据(第2行物料+孔位1)
        # 物料第2行数据
        for col_idx, value in enumerate(material_data.iloc[0], start=2):
            ws.cell(row=rows + 36, column=col_idx+5, value=value)

        # 孔位1数据
        for row_idx, row in hole_data_1.iterrows():
            for col_idx, value in enumerate(row, start=2):
                ws.cell(row=row_idx + 39, column=col_idx+5, value=value)


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
                get_row_values(sheet, 3, prefix)
            ])

        # 创建Excel并写入数据
        wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["information"])
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
        data_list_12 = []
        data_list_24, result_1 = self._process_24_plate()

        # 文件路径

        # 生成第一个物料表格
        if len(data_list_24) > 12:
            ExcelUtils.copy_workbook(
                f"{BASE_PATH1}/hiwo物料-24.xlsx",
                TEMPLATE_PATHS["mate_info1"]
            )
            wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info1"])
            self._fill_worksheet(wb, [], data_list_24[:12])

            """修改24孔2ml、4ml 模板文件"""
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板2ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-2ml')
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板4ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-4ml')
            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info1"])
            # 生成第二个物料表格
            wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info2"])

            self._create_extended_version(wb, [], data_list_24[12:24], self.number_tasks)
            """修改24孔2ml、4ml 模板文件"""
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板2ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-2ml')
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板4ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-4ml')
            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info2"])

        else:
            """self.usage_12_hole是12孔板的使用最小使用孔数，最大孔是12孔，用最大值12减去使用的最小值，等于使用孔数，然后再用12减去它，等于剩余孔数，所以直接用24减去它"""
            if int(24 - int(self.usage_12_hole)) > len(data_list_12):
                self.number_tasks = 1
                wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info2"])

                self._create_extended_version(wb, data_list_12, data_list_24, self.number_tasks)
                """修改24孔2ml、4ml 模板文件"""
                if self.post_data['experimentLog1']['wellPlates'] == '24孔板2ml':
                    ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-2ml')
                if self.post_data['experimentLog1']['wellPlates'] == '24孔板4ml':
                    ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-4ml')
                ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info2"])
                """如果24孔使用数量小于6，那么只需要一个任务就可以完成"""

        return result_1, data_list_12

    def _create_crystallize_mate_info(self) -> Tuple[list, list]:
        """创建物料信息Excel"""
        # 初始化数据列表

        data_list_12 = []
        data_list_24, result_1 = self._process_24_plate()
        ExcelUtils.copy_workbook(
            f"{crystallize_BASE_PATH1}/hiwo物料-24-02.xlsx",
            crystallize_TEMPLATE_PATHS["mate_info1"]
        )


        # 生成第一个物料表格
        if len(data_list_24) > 24 :


            wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["mate_info1"])
            self._fill_worksheet(wb, [], data_list_24[:24])
            """修改24孔2ml、4ml 模板文件"""
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板2ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-2ml')
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板4ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-4ml')

            ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["mate_info1"])

            # 生成第二个物料表格
            wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["mate_info2"])
            self._fill_worksheet(wb, [], data_list_24[24:])
            """修改24孔2ml、4ml 模板文件"""
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板2ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-2ml')
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板4ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-4ml')
            ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["mate_info2"])

        else:
            self.number_tasks = 1
            wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["mate_info1"])
            # 24孔板
            self._fill_worksheet(wb, [], data_list_24)
            """修改24孔2ml、4ml 模板文件"""
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板2ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-2ml')
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板4ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-4ml')
            ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["mate_info1"])
            """如果24孔使用数量小于6，那么只需要一个任务就可以完成"""

        return result_1, data_list_12

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

        len_24 = len(transfer_records_24)
        CELLS = ["C2", "C3", "C6", "C7", "D2", "D3"]  # 共6个，对应 len_24 在 (0,36] 的6个区间

        if len_24 > 0:
            # 计算需要前几个单元格：每6长度为一个单位，向上取整
            n = (len_24 + 5) // 6  # 等价于 ceil(len_24 / 6)
            n = min(n, len(CELLS))  # 限制不超过最大单元格数（len_24 > 36 时取全部）
            fill_cells = CELLS[:n]  # 直接切片，一次性生成
        else:
            fill_cells = []  # len_24 == 0 时为空

        if "A05" in self.gun_head_dict:
            fill_cells.append("C7")
        if "B05" in self.gun_head_dict:
            fill_cells.append("D7")
        if "A04" in self.gun_head_dict:
            fill_cells.append("E6")
        if "B04" in self.gun_head_dict:
            fill_cells.append("B7")
            
        self.process_plate_stack(
        file_path=fr"{crystallize_BASE_PATH1}/Materials-stack4列.xlsx",
        output_path=fr"{crystallize_BASE_PATH}/Materials-stack4列.xlsx",
        fill_cells = fill_cells,
        plate_type='24孔',
        condition=self.post_data,
        )
    def add_excel_filter(self, num, type,key_dict = {}):
        #新增的过滤逻辑
        result = []
        for group in self.perforated_plate:
            # 计算每组两个键对应的值总和
            sum1 = sum(value for dictionary in self.data_dict[group[0]]
                       for value in dictionary.values())
            if type == 2 :
                sum1 += int(key_dict.get(int(group[0]),{}).get("weight",''))
            try:
                # 计算第二个键的值总和
                sum2 = sum(value for dictionary in self.data_dict[group[1]]
                           for value in dictionary.values())
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
        wb = ExcelUtils.open_workbook(f"{crystallize_BASE_PATH}/hiwo-物料3-{num + 1}.xlsx")
        if self.post_data['experimentLog1']['wellPlates'] == '24孔板2ml':
            ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-2ml')
        if self.post_data['experimentLog1']['wellPlates'] == '24孔板4ml':
            ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-4ml')
        ExcelUtils.save_workbook(wb, f"{crystallize_BASE_PATH}/hiwo-物料3-{num + 1}.xlsx")

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
                                  ["FY24-B01", s, "GL24-2A01", s, i * 1000, pipette_location, gun_head],
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
                f"{crystallize_BASE_PATH1}/hiwo物料-24-02.xlsx",
                f"{crystallize_BASE_PATH}/hiwo-物料2-1.xlsx"
            )
            wb = ExcelUtils.open_workbook(f"{crystallize_BASE_PATH}/hiwo-物料2-1.xlsx")
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板2ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-2ml')
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板4ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-4ml')
            ExcelUtils.save_workbook(wb, f"{crystallize_BASE_PATH}/hiwo-物料2-1.xlsx")
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
                        f"{crystallize_BASE_PATH1}/hiwo物料-24-02.xlsx",
                        f"{crystallize_BASE_PATH}/hiwo-物料2-{num}.xlsx"
                    )
                    wb = ExcelUtils.open_workbook(f"{crystallize_BASE_PATH}/hiwo-物料2-{num}.xlsx")
                    if self.post_data['experimentLog1']['wellPlates'] == '24孔板2ml':
                        ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-2ml')
                    if self.post_data['experimentLog1']['wellPlates'] == '24孔板4ml':
                        ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-4ml')
                    ExcelUtils.save_workbook(wb, f"{crystallize_BASE_PATH}/hiwo-物料2-{num}.xlsx")

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
                f"{crystallize_BASE_PATH1}/hiwo物料-24-02.xlsx",
                f"{crystallize_BASE_PATH}/hiwo-物料3-1.xlsx"
            )
            wb = ExcelUtils.open_workbook(f"{crystallize_BASE_PATH}/hiwo-物料3-1.xlsx")
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板2ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-2ml')
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板4ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-4ml')
            ExcelUtils.save_workbook(wb, f"{crystallize_BASE_PATH}/hiwo-物料3-1.xlsx")

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
            # 三阶段统一不关盖不开盖
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
                        f"{crystallize_BASE_PATH1}/hiwo物料-24-02.xlsx",
                        f"{crystallize_BASE_PATH}/hiwo-物料3-{num}.xlsx"
                    )
                    wb = ExcelUtils.open_workbook(f"{crystallize_BASE_PATH}/hiwo-物料3-{num}.xlsx")
                    if self.post_data['experimentLog1']['wellPlates'] == '24孔板2ml':
                        ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-2ml')
                    if self.post_data['experimentLog1']['wellPlates'] == '24孔板4ml':
                        ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-4ml')
                    ExcelUtils.save_workbook(wb, f"{crystallize_BASE_PATH}/hiwo-物料3-{num}.xlsx")

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
                    f"{crystallize_BASE_PATH1}/hiwo物料-24-02.xlsx",
                    f"{crystallize_BASE_PATH}/hiwo-物料2-1.xlsx"
                )
                wb = ExcelUtils.open_workbook(f"{crystallize_BASE_PATH}/hiwo-物料2-1.xlsx")
                if self.post_data['experimentLog1']['wellPlates'] == '24孔板2ml':
                    ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-2ml')
                if self.post_data['experimentLog1']['wellPlates'] == '24孔板4ml':
                    ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-4ml')
                ExcelUtils.save_workbook(wb, f"{crystallize_BASE_PATH}/hiwo-物料2-1.xlsx")
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
                            f"{crystallize_BASE_PATH1}/hiwo物料-24-02.xlsx",
                            f"{crystallize_BASE_PATH}/hiwo-物料2-{num}.xlsx"
                        )
                        wb = ExcelUtils.open_workbook(f"{crystallize_BASE_PATH}/hiwo-物料2-{num}.xlsx")
                        if self.post_data['experimentLog1']['wellPlates'] == '24孔板2ml':
                            ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-2ml')
                        if self.post_data['experimentLog1']['wellPlates'] == '24孔板4ml':
                            ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-4ml')
                        ExcelUtils.save_workbook(wb, f"{crystallize_BASE_PATH}/hiwo-物料2-{num}.xlsx")

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
            except Exception as e:
                print(e)


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

            for perforated in perforated_:
                hole_mapping.update({perforated: k})

        postProcessingForm = self.post_data['experimentLog2']['postProcessingForm']
        precoolingTemperature = postProcessingForm['precoolingTemperature']
        time = postProcessingForm['time']
        temperature = postProcessingForm['temperature']
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
                f"{crystallize_BASE_PATH1}/hiwo物料-24-02.xlsx",
                excel_name
            )
            wb = ExcelUtils.open_workbook(f"{crystallize_BASE_PATH}/hiwo-物料3-1.xlsx")
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板2ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-2ml')
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板4ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-4ml')
            ExcelUtils.save_workbook(wb, f"{crystallize_BASE_PATH}/hiwo-物料3-1.xlsx")

            #反加也需要物料，不然吸液过去干嘛
            wb = ExcelUtils.open_workbook(excel_name)
            for i, keys in enumerate(self.perforated_plate):
                row = 1 + i * 4
                for key in keys:
                    row+=1
                    data_key = [aa for aa in all_solid_records if aa['hole_id']==int(key)][0]
                    name = data_key['reagent_name']
                    ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}", name)

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
                data_key = [aa for aa in all_solid_records if aa['hole_id'] == int(list_data[0])][0]
                hole_id = data_key['hole_id']
                amount = data_key['weight']
                gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0
                target_volume = max_key * 1000  if max_key else 0
                target_volumes = amount * 1000 if amount else 0
                bar_code = "SS-B02"

                record = {
                    "来源孔板条码": "SS-A03",
                    "来源孔板列Y": hole_mapping[str(hole_id)],
                    "目标孔板条码": bar_code,
                    "目标孔板列Y": hole_mapping[str(hole_id)],
                    "移液量(ul)": target_volumes,
                    "枪头库位": pipette_location,
                    "枪头列Y": gun_head + 1
                }
                record_list.append(record)
                # if len(record_list) > 6:
                #     record["来源孔板条码"] = "SS-A03"
                #     record["来源孔板列Y"] = num - 6
                if len(self.gun_head_dict[pipette_location]) >= 12:
                    pipette_location = "B05"
                if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                    self.gun_head_dict[pipette_location].append(gun_head + 1)


                gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0
                bar_code = "SS-B02"
                record = {
                    "来源孔板条码": "FY24-B01",
                    "来源孔板列Y": hole_mapping[str(hole_id)],
                    "目标孔板条码": bar_code,
                    "目标孔板列Y": hole_mapping[str(hole_id)],
                    "移液量(ul)": target_volume,
                    "枪头库位": pipette_location,
                    "枪头列Y": gun_head + 1
                }
                record_lists.append(record)
                # if len(record_list) > 6:
                #     record["来源孔板条码"] = "SS-A03"
                #     record["来源孔板列Y"] = num - 6
                # if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                #     self.gun_head_dict[pipette_location].append(gun_head + 1)
                #
                # gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0
                #总移液量是前面两个的和
                Total_pipetting_volume = target_volumes+target_volume
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
            wb = ExcelUtils.open_workbook(f"{crystallize_BASE_PATH}/hiwo-物料3-2.xlsx")
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板2ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-2ml')
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板4ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-4ml')
            ExcelUtils.save_workbook(wb, f"{crystallize_BASE_PATH}/hiwo-物料3-2.xlsx")

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
                f"{crystallize_BASE_PATH1}/hiwo物料-24-02.xlsx",
                excel_name
            )
            wb = ExcelUtils.open_workbook(f"{crystallize_BASE_PATH}/hiwo-物料3-1.xlsx")
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板2ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-2ml')
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板4ml':
                ExcelUtils.modify_cell(wb, "物料信息", "B7", '24孔反应板-4ml')
            ExcelUtils.save_workbook(wb, f"{crystallize_BASE_PATH}/hiwo-物料3-1.xlsx")

            wb = ExcelUtils.open_workbook(excel_name)
            if all_solid_records:
                for i, key in enumerate(self.perforated_plate):
                    row = 2 + i * 4

                    ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}",
                                           key_dict.get(int(key[0]), []).get("reagent_name", ''))
                    if len(key) >= 2:
                        ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 1}",
                                               key_dict.get(int(key[1]), []).get("reagent_name", ''))
                    if len(key) >= 3:
                        ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 2}",
                                               key_dict.get(int(key[2]), []).get("reagent_name", ''))
                    if len(key) >= 4:
                        ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 3}",
                                               key_dict.get(int(key[3]), []).get("reagent_name", ''))


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

                gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0
                # target_volume = max_num * 1000 / 2 if max_num else 0
                target_volume = int(key_dict.get(int(list_data[0]), []).get("weight", ''))  * 1000
                bar_code = "SS-A02"

                record = {
                    "来源孔板条码": bar_code,
                    "来源孔板列Y": llwc,
                    "目标孔板条码": "FY24-B01",
                    "目标孔板列Y": llwc,
                    "移液量(ul)": target_volume,
                    "枪头库位": pipette_location,
                    "枪头列Y": gun_head + 1
                }
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
        name_list_len = len(name_list)
        '''
        如果超出13中物料那么这里不需要改
        如果没超出13种物料那么不生成1-2
        '''
        # 处理第二张表 (如果有超过13种物料)

        '''
        做两次判断：
            1.如果物料超过13种，生成物料2-1
            2.如果孔位超过12个，生成物料1-2
            3。如果物料超过13种，且孔位超过12个 两个都生成
        '''
        excel_list_2 = []
        if  len(unique_reagents) >13:
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
            target_path = crystallize_TEMPLATE_PATHS["sowo"].replace("物料1-1.xlsx", "物料2-1.xlsx")
            ExcelUtils.copy_workbook(f"{crystallize_BASE_PATH1}/sowo物料.xlsx", target_path)
            wb = ExcelUtils.open_workbook(target_path)
            ExcelUtils.write_records_to_sheet(wb, "Sheet1", excel_list_2)
            ExcelUtils.write_records_to_sheet(wb, "Sheet1", [{"Code": f"{day_name}-001", "Locate": 1, "Type": 1}])
            ExcelUtils.save_workbook(wb, target_path)
        if self.post_data['experimentLog1']['wellPlates'] == '24孔板8ml':
            if len(self.data_dict) >12:
                ExcelUtils.copy_workbook(
                    crystallize_TEMPLATE_PATHS["sowo"],
                    crystallize_TEMPLATE_PATHS["sowo"].replace('物料1-1',"物料1-2")
                )
            if  len(unique_reagents) >13 and len(self.data_dict) >12:
                #
                ExcelUtils.copy_workbook(
                    crystallize_TEMPLATE_PATHS["sowo"],
                    crystallize_TEMPLATE_PATHS["sowo"].replace('物料1-1',"物料1-2")
                )
                target_path = crystallize_TEMPLATE_PATHS["sowo"].replace("物料1-1.xlsx", "物料2-1.xlsx")
                ExcelUtils.copy_workbook(
                    target_path,
                    crystallize_TEMPLATE_PATHS["sowo"].replace('物料1-1',"物料2-2")
                )
        else:
            pass
        #表三
        self.write_codes_to_inner_hopper_stack(fr"{crystallize_BASE_PATH}/Materials-stack4列.xlsx",fr"{crystallize_BASE_PATH}/Materials-stack4列.xlsx",excel_list,excel_list_2)

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
        reagent_totals = {}

        for row in self.reagent_plates_24:
            # 提取前4个试剂字典
            reagent_dicts = row[:4]
            # 过滤掉None，只保留非空的试剂字典
            valid_reagents = [rd for rd in reagent_dicts if rd is not None]
            if valid_reagents:  # 至少有一个非空试剂
                # 提取试剂名称和剂量
                reagent_names = []
                reagent_volumes = []
                for reagent_dict in valid_reagents:
                    for reagent, volume in reagent_dict.items():
                        reagent_names.append(reagent)
                        reagent_volumes.append(volume)
                # 检查所有非空试剂是否完全相同
                if len(set(reagent_names)) == 1 :
                    # 所有非空试剂完全相同，累加剂量
                    reagent_name = reagent_names[0]
                    reagent_volume = reagent_volumes[0]
                    total_volume = reagent_volume * len(valid_reagents)  # 乘以实际数量
                    if reagent_name in reagent_totals:
                        reagent_totals[reagent_name] += total_volume
                        self.reagent_hole[reagent_name].append(row[4])
                    else:
                        reagent_totals[reagent_name] = total_volume
                        self.reagent_hole[reagent_name] = [row[4]]



        for item in self.reagent_plates_24:

            chem1 = list(item[0].keys())[0] if item[0]  else None
            chem2 = list(item[1].keys())[0] if item[1]  else None
            chem3 = list(item[2].keys())[0] if item[2]  else None
            chem4 = list(item[3].keys())[0] if item[3]  else None

            values1 = list(item[0].values())[0] if item[0] else None
            values2 = list(item[1].values())[0] if item[1] else None
            values3 = list(item[2].values())[0] if item[2] else None
            values4 = list(item[3].values())[0] if item[3] else None
            # 先过滤掉None值再求最大值
            valid_values = [v for v in (values1, values2, values3, values4) if v is not None]
            max_vol = max(valid_values) if valid_values else None

            if chem4 ==chem2 ==chem3==chem1:
                Reagent_name = reagent_totals.get(chem1,"")
                if Reagent_name:
                    #8*4=32:这里的数量是一横排反应的容量，而不是单孔的
                    hole_num = math.ceil(reagent_totals[chem1] / 32)
                    del reagent_totals[chem1]
                    for hole in range(hole_num):
                        result.append([chem1, chem2, chem3, chem4])
                        result_1.append([chem1, chem2, chem3, chem4, [item[4]], max_vol])
            else:
                result.append([chem1, chem2,chem3,chem4])
                result_1.append([chem1, chem2,chem3,chem4, [item[4]], max_vol])

        return result, result_1

    def _fill_worksheet(self, wb: Workbook, data_12: List[str],
                        data_24: List[List[Union[str, None]]]) -> None:
        """填充工作表数据"""

        # if self.type == 1:
        #     # 填充24孔板数据
        #     ExcelUtils.modify_row(wb, "物料信息", 2,
        #                           ["24孔试剂板-低板", "24孔板", "SS-A02", 1, "堆栈", 1, "A02"],
        #                           start_col=2)
        #     ExcelUtils.delete_row(wb, "孔位信息", 14)
        #     self._copy_insert_rows(wb, "孔位信息")
        #     start = 2
        # else:
        #     start = 2
        start = 2
        for i, (chem1, chem2,chem3,chem4) in enumerate(data_24):
            row = start + i * 4
            if chem1:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}", chem1)
            if chem2:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 1}", chem2)
            if chem3:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 2}", chem3)
            if chem4:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 3}", chem4)
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
        # stop_num = self.usage_12_hole - 1
        # for i, chem in enumerate(data_12[:stop_num], 2):
        #     ExcelUtils.modify_cell(wb, "孔位信息", f"D{i}", chem)

        # 处理12孔板扩展数据，llwc_num：记录12孔板用了多少24孔板的位置
        llwc_num = 0
        # if len(data_12) > stop_num:
        #     for i, chem in enumerate(data_12[stop_num:]):
        #         row = 14 + i * 4
        #         ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}", chem)
        #         ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 2}", chem)
        #         llwc_num += 1
        #         if llwc_num > 6:
        #             """如果12孔板剩余的数量大于6，也就是说一块24孔板不够使用，那么也是应该第三次实验吗？"""
        #             pass
        #             break
        if data_24:
            for i ,chem in enumerate(data_24):
                row = 2+llwc_num*4+i*4
                if chem[0]:
                    ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}", chem[0])
                if chem[1]:
                    ExcelUtils.modify_cell(wb, "孔位信息", f"D{row+1}", chem[1])
                if chem[2]:
                    ExcelUtils.modify_cell(wb, "孔位信息", f"D{row+2}", chem[2])
                if chem[3]:
                    ExcelUtils.modify_cell(wb, "孔位信息", f"D{row+3}", chem[3])



        # 处理24孔板扩展数据,从哪开始取决于 12孔板用完用了多少24孔板的孔位
        """
        如果是生成第二个表的逻辑，第一个表使用两块24孔板，第二块表使用一块12孔板一块24孔板
        data_24_start+2 + i * 4： data_24_start默认是12，但是在excel表格里12孔板用完行数是13，应该从14行开始才是24孔板
        如果只需要一个表，也就是使用一块12孔板以及一块24孔板就行，那么该怎么修改这个逻辑

        """
        # if number_tasks == 1:
        #     data_24_start = 0
        # else:
        #     data_24_start = 12
        #
        # if len(data_24) > data_24_start:
        #     for i, (chem1, chem2) in enumerate(data_24[data_24_start:18]):
        #         row = 12 + 2 + llwc_num * 4 + i * 4
        #         ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}", chem1)
        #         if chem2:
        #             ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 2}", chem2)
        # if len(data_24) > 18:
        #     """如果第二个任务的24孔板也不够用，是不是需要第三次任务？"""
        #     pass
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

    def _create_w11_task(self) -> None:
        """创建W11任务Excel"""
        hole_mapping = {}
        for k, perforated_ in enumerate(self.perforated_plate, 1):
            if len(perforated_) == 4:
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
                # ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_12[:10])
                ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_2)
                ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf11"])

            if transfer_records_24:
                ExcelUtils.copy_workbook(
                    f"{BASE_PATH1}/hiwo任务05-8-24.xlsx",
                    TEMPLATE_PATHS["wf05"]
                )
                wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf05"])
                ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_24)
                ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf05"])



        len_24 = len(transfer_records_24)
        CELLS = ["C2", "C3", "C6", "C7", "D2", "D3"]  # 共6个，对应 len_24 在 (0,36] 的6个区间

        if len_24 > 0:
            # 计算需要前几个单元格：每6长度为一个单位，向上取整
            n = (len_24 + 5) // 6  # 等价于 ceil(len_24 / 6)
            n = min(n, len(CELLS))  # 限制不超过最大单元格数（len_24 > 36 时取全部）
            fill_cells = CELLS[:n]  # 直接切片，一次性生成
        else:
            fill_cells = []  # len_24 == 0 时为空

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
            fill_cells=fill_cells,
            plate_type='24孔',
            condition=self.post_data,
        )

    def _generate_transfer_data(self, hole_mapping: dict) -> Tuple[list, list, list]:
        """生成孔板转移数据"""
        transfer_records_12,transfer_records_24,transfer_records_2,pipette_location = [],[],[],"A05"
        # 处理24孔板数据
        add_ = 0
        substance_list =[]
        for plate_idx, (substance1, substance2,substance3,substance4, hole_id, amount) in enumerate(self.data_list_24, 1):
            if len(self.gun_head_dict[pipette_location]) >= 12:
                pipette_location = "B05"

            gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0
            target_volume = amount * 1000  if amount else 0
            if substance1 in list(self.reagent_hole.keys()) and substance1==substance2==substance3==substance4 and substance1 not in substance_list:
                substance_list.append(substance1)
                Quantity_used = 8000
                for range_num in self.reagent_hole[substance1]:
                    add_num = Quantity_used - math.ceil(target_volume)
                    if add_num < 1:
                        add_ += 1
                        Quantity_used = 8000
                    else:
                        Quantity_used -= target_volume
                    plate_idx_ = plate_idx + add_
                    if self.type == 1:
                        bar_code = "SS-A02"
                        record = {
                            "来源孔板条码": bar_code,
                            "来源孔板列Y": plate_idx_,
                            "目标孔板条码": "FY24-B01",
                            "目标孔板列Y": hole_mapping[range_num],
                            "移液量(ul)": target_volume,
                            "枪头库位": pipette_location,
                            "枪头列Y": gun_head + 1}
                        if len(transfer_records_2) >= 6:
                            record["来源孔板条码"] = "SS-A03"
                            record['来源孔板列Y'] = plate_idx_-18
                            transfer_records_2.append(record)
                        elif plate_idx_ >= 12:
                            record["来源孔板条码"] = "SS-A02"
                            record['来源孔板列Y'] = plate_idx_-12
                            transfer_records_2.append(record)
                        else:
                            if plate_idx_>= 6:
                                record["来源孔板条码"] = "SS-A03"
                                record["来源孔板列Y"] = plate_idx_ - 6
                            transfer_records_24.append(record)
                        if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                            self.gun_head_dict[pipette_location].append(gun_head + 1)
                    else:
                        record = {
                            "来源孔板条码": "SS-A02",
                            "来源孔板列Y": plate_idx_ ,
                            "目标孔板条码": "FY24-B01",
                            "目标孔板列Y": hole_mapping[range_num],
                            "移液量(ul)": target_volume,
                            "枪头库位": pipette_location,
                            "枪头列Y": gun_head + 1
                        }

                        if plate_idx_ >18:
                            record["来源孔板条码"] = "SS-B03"
                            record["来源孔板列Y"] = plate_idx - 18
                            transfer_records_2.append(record)
                        else:
                            if plate_idx_ > 6:
                                record["来源孔板条码"] = "SS-A03"
                                record["来源孔板列Y"] = plate_idx - 6
                            if plate_idx_ > 12:
                                record["来源孔板条码"] = "SS-B02"
                                record["来源孔板列Y"] = plate_idx_ - 12
                            transfer_records_24.append(record)

                        if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                            self.gun_head_dict[pipette_location].append(gun_head + 1)
            elif substance1==substance2==substance3==substance4:
                add_ -= 1
                continue
            else:
                if self.type == 1:

                    bar_code = "SS-A02"
                    record = {
                        "来源孔板条码": bar_code,
                        "来源孔板列Y": plate_idx+add_,
                        "目标孔板条码": "FY24-B01",
                        "目标孔板列Y": hole_mapping[hole_id[0]],
                        "移液量(ul)": target_volume,
                        "枪头库位": pipette_location,
                        "枪头列Y": gun_head + 1
                    }
                    if plate_idx+add_> 18:
                        record["来源孔板条码"] = "SS-A03"
                        record['来源孔板列Y'] = plate_idx+add_ - 18
                        transfer_records_2.append(record)

                    elif plate_idx+add_ > 12:
                        record["来源孔板条码"] = "SS-A02"
                        record['来源孔板列Y'] = plate_idx+add_ - 12
                        transfer_records_2.append(record)
                    else:
                        if plate_idx+add_ > 6:
                            record["来源孔板条码"] = "SS-A03"
                            record["来源孔板列Y"] = plate_idx+add_ - 6
                        transfer_records_24.append(record)

                    if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                        self.gun_head_dict[pipette_location].append(gun_head + 1)
                else:
                    record = {
                        "来源孔板条码": "SS-A02",
                        "来源孔板列Y": plate_idx ,
                        "目标孔板条码": "FY24-B01",
                        "目标孔板列Y": hole_mapping[hole_id[0]],
                        "移液量(ul)": target_volume,
                        "枪头库位": pipette_location,
                        "枪头列Y": gun_head + 1
                    }

                    if plate_idx+add_ >= 18:
                        record["来源孔板条码"] = "SS-B03"
                        record["来源孔板列Y"] = plate_idx+add_ - 18
                        transfer_records_2.append(record)
                    else:
                        if plate_idx+add_ >= 6:
                            record["来源孔板条码"] = "SS-A03"
                            record["来源孔板列Y"] = plate_idx+add_ - 6
                        if plate_idx+add_ >12:
                            record["来源孔板条码"] = "SS-B02"
                            record["来源孔板列Y"] = plate_idx+add_ - 12
                        transfer_records_24.append(record)

                    if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                        self.gun_head_dict[pipette_location].append(gun_head + 1)
            # 处理12孔板数据

        return transfer_records_12, transfer_records_24, transfer_records_2

    def _create_task_quencher(self):
        ExcelUtils.copy_workbook(
            f"{BASE_PATH1}/hiwo物料-24.xlsx",
            TEMPLATE_PATHS["mate_info2"]
        )

        ExcelUtils.copy_workbook(
            f"{BASE_PATH1}/hiwo任务11-8-24.xlsx",
            TEMPLATE_PATHS["wf11"]
        )
        # 淬灭剂
        quenchingAgentData = self.post_data['experimentLog3']["quenchingAgentData"]

        wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info2"])
        # ExcelUtils.modify_row(wb, "物料信息", 2,
        #                       ["24孔试剂板-低板", "24孔板", "SS-A02", 1, "堆栈", 1, "A02"],
        #                       start_col=2)
        # ExcelUtils.delete_row(wb, "孔位信息", 14)
        # self._copy_insert_rows(wb, "孔位信息")
        if quenchingAgentData:
            # 添加淬灭剂到物料表
            quenchingAgent = quenchingAgentData[0]["name"]
            quenchingAgent_Volume = int(float(quenchingAgentData[0]['addVolume']))
            quenchingAgent_Type = True
            quenchingAgent_Volume_all = int(quenchingAgent_Volume) * (self.hole_number  * 4)
            #需要添加几个格子的淬灭剂
            quenchingAgent_Volume_num= (quenchingAgent_Volume_all + 8000 - 1) // 8000
            for num in range(quenchingAgent_Volume_num):
                ExcelUtils.modify_cell(wb, "孔位信息", f'D{49-num*4}', quenchingAgent)
                ExcelUtils.modify_cell(wb, "孔位信息", f'D{49-num*4-1}', quenchingAgent)
                ExcelUtils.modify_cell(wb, "孔位信息", f'D{49-num*4-2}', quenchingAgent)
                ExcelUtils.modify_cell(wb, "孔位信息", f'D{49-num*4-3}', quenchingAgent)
            # if quenchingAgent_Volume_all >= 8000:
            #     quenchingAgent_Type = None
            #
            #     ExcelUtils.modify_cell(wb, "孔位信息", f'D{13}', quenchingAgent)
            #     # ExcelUtils.modify_cell(wb, "孔位信息", f'E{13}', 18000)
            #     ExcelUtils.modify_cell(wb, "孔位信息", f'D{12}', quenchingAgent)
            #     # ExcelUtils.modify_cell(wb, "孔位信息", f'E{12}', int(quenchingAgent_Volume)*8*4-18000)
            # else:
            #     ExcelUtils.modify_cell(wb, "孔位信息", f'D{13}', quenchingAgent)
            #     # ExcelUtils.modify_cell(wb, "孔位信息", f'E{13}', int(quenchingAgent_Volume)*8*4)
            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info2"])

            if int(quenchingAgent_Volume) <= 200:pipette_location = 'A04'
            else:pipette_location = 'A05'
            if len(self.gun_head_dict[pipette_location]) >= 12:
                if pipette_location == 'A04':pipette_location = 'B04'
                if pipette_location == 'A05':pipette_location = 'B05'
            try:
                gun_head = self.gun_head_dict[pipette_location][-1]
            except:
                gun_head = 0
            """添加淬灭剂取液到wf11表格"""
            wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf11"])
            plate_counter = 0
            plate_remaining = 8000  # 每个孔板最大容量8ml (8000μl)
            for age in range(1, self.hole_number + 1):
                # if quenchingAgent_Volume_num ==1:
                #     ExcelUtils.modify_row(wb, "加淬灭液|||移液信息", age + 1,
                #                           ["SS-A03", 12-6, "FY24-B01", age, int(quenchingAgent_Volume),
                #                            pipette_location, gun_head + 1], start_col=1)
                # else:
                    vol_per_hole = int(quenchingAgent_Volume)*4
                    # 检查当前孔板剩余容量是否足够
                    if plate_remaining < vol_per_hole:
                        # 切换到下一个孔板
                        plate_counter += 1
                        plate_remaining = 8000  # 重置为新孔板的完整容量
                    # 更新当前孔板剩余容量
                    plate_remaining -= vol_per_hole
                    ExcelUtils.modify_row(wb, "加淬灭液|||移液信息", age + 1,
                                          ["SS-A03", 6-plate_counter, "FY24-B01", age, int(quenchingAgent_Volume),
                                           pipette_location, gun_head + 1], start_col=1)
            if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                self.gun_head_dict[pipette_location].append(gun_head + 1)
            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf11"])

    def _create_task_internal_standard(self):
        # 内标数据
        internalStandardData = self.post_data['experimentLog3']["internalStandardData"]
        if not internalStandardData:
            return

        #获取淬灭剂使用24孔板数
        if self.post_data['experimentLog3']["quenchingAgentData"]:
            quenchingAgent_Volume = int(float(self.post_data['experimentLog3']["quenchingAgentData"][0]['addVolume']))
            quenchingAgent_Volume_num = (int(quenchingAgent_Volume) * (self.hole_number * 4) + 8000 - 1) // 8000
        else:
            quenchingAgent_Volume_num = 0
        internalStandard = internalStandardData[0]["name"]
        internalStandard_Volume = int(float(internalStandardData[0]['addVolume']))
        wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info2"])

        """quenching_used_rows[0]-1: 内标使用的最小孔位-1就是未使用的"""
        # if quenching_used_rows:
        #     internal_standard_used_rows = quenching_used_rows[0] - 1
        # else:
        #     internal_standard_used_rows = 13
        internal_standard_used_rows = 6- quenchingAgent_Volume_num

        internalStandard_Volume_all = int(internalStandard_Volume) * (self.hole_number * 4)
        # 需要添加几个格子的淬灭剂
        internalStandard_Volume_num = (internalStandard_Volume_all + 8000 - 1) // 8000
        for num in range(internalStandard_Volume_num):
            ExcelUtils.modify_cell(wb, "孔位信息", f'D{49-num*4-quenchingAgent_Volume_num*4 }', internalStandard)
            ExcelUtils.modify_cell(wb, "孔位信息", f'D{49-num*4-quenchingAgent_Volume_num*4-1}', internalStandard)
            ExcelUtils.modify_cell(wb, "孔位信息", f'D{49-num*4-quenchingAgent_Volume_num*4-2}', internalStandard)
            ExcelUtils.modify_cell(wb, "孔位信息", f'D{49-num*4-quenchingAgent_Volume_num*4-3}', internalStandard)

        ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info2"])

        if int(internalStandard_Volume) <= 200:pipette_location = 'A04'
        else:pipette_location = 'A05'
        if len(self.gun_head_dict[pipette_location]) >= 12:
            if pipette_location == 'A04':pipette_location = 'B04'
            if pipette_location == 'A05':pipette_location = 'B05'
        try:
            gun_head = self.gun_head_dict[pipette_location][-1]
        except:
            gun_head = 0
        """添加内标取液到wf11表格"""
        wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf11"])
        plate_counter = 0
        plate_remaining = 8000  # 每个孔板最大容量8ml (8000μl)
        for age in range(1, self.hole_number + 1):

            vol_per_hole = int(internalStandard_Volume)*4
            # 检查当前孔板剩余容量是否足够
            if plate_remaining < vol_per_hole:
                plate_counter += 1
                plate_remaining = 8000  # 重置为新孔板的完整容量
            # 更新当前孔板剩余容量
            plate_remaining -= vol_per_hole
            ExcelUtils.modify_row(wb, "加内标|||移液信息", age + 1,["SS-A03", internal_standard_used_rows -plate_counter ,
            "FY24-B01", age, int(internalStandard_Volume),pipette_location, gun_head + 1], start_col=1)
            self.usage_12_hole = internal_standard_used_rows - 2
        if gun_head + 1 not in self.gun_head_dict[pipette_location]:
            self.gun_head_dict[pipette_location].append(gun_head + 1)
        ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf11"])

    def _create_task_extract(self):
        # 萃取液数据
        extractionData = self.post_data['experimentLog3']["extractionData"]
        if not extractionData:
            return

        #淬灭剂使用孔位
        if self.post_data['experimentLog3']["quenchingAgentData"]:
            quenchingAgent_Volume = int(float(self.post_data['experimentLog3']["quenchingAgentData"][0]['addVolume']))
            quenchingAgent_Volume_num = (int(quenchingAgent_Volume) * (self.hole_number * 4) + 8000 - 1) // 8000
        else:
            quenchingAgent_Volume_num = 0
        #淬灭剂使用孔位
        if self.post_data['experimentLog3']["internalStandardData"]:
            internalStandardData_Volume = int(float(self.post_data['experimentLog3']["internalStandardData"][0]['addVolume']))
            internalStandardData_Volume_num = (int(internalStandardData_Volume) * (self.hole_number * 4) + 8000 - 1) // 8000
        else:
            internalStandardData_Volume_num = 0
        # 安全地处理淬灭剂孔位信息
        internal_standard_used_rows = 12-6-quenchingAgent_Volume_num-internalStandardData_Volume_num
        internalStandard = extractionData[0]["name"]
        if internalStandard:
            internalStandard_Volume = 100
            wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info2"])
            # 添加萃取液到物料表
            ExcelUtils.modify_cell(wb, "孔位信息", f'D{49-quenchingAgent_Volume_num*4-internalStandardData_Volume_num*4}', internalStandard)
            ExcelUtils.modify_cell(wb, "孔位信息", f'D{49-quenchingAgent_Volume_num*4-internalStandardData_Volume_num*4 -1}', internalStandard)
            ExcelUtils.modify_cell(wb, "孔位信息", f'D{49-quenchingAgent_Volume_num*4-internalStandardData_Volume_num*4 -2}', internalStandard)
            ExcelUtils.modify_cell(wb, "孔位信息", f'D{49-quenchingAgent_Volume_num*4-internalStandardData_Volume_num*4 -3}', internalStandard)
            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info2"])
            if int(internalStandard_Volume) <= 200:pipette_location = 'A04'
            else:pipette_location = 'A05'
            if len(self.gun_head_dict[pipette_location]) >= 12:
                if pipette_location == 'A04':
                    pipette_location = 'B04'
                if pipette_location == 'A05':
                    pipette_location = 'B05'
            try:gun_head = self.gun_head_dict[pipette_location][-1]
            except:gun_head = 0
            """添加萃取液到任务表格"""
            wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf11"])
            for age in range(1, self.hole_number + 1):
                try:
                    if self.post_data["experimentLog3"].get("productDilutionData", [])[0].get("procedure", {}).get(
                            "diluent3", ''):
                        start = self.hole_number + 1 + age
                    else:
                        start = 1 + age
                except:
                    start = 1 + age
                ExcelUtils.modify_row(wb, "加稀释液-过滤|||移液信息", start,
                                      ["SS-A03", internal_standard_used_rows - 1, "GL96-2A01", age, 100,
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
                if int(temp)>=26:pass
                elif int(temp)<=60:time+=3
                elif int(temp)<=100:time+=6
                elif int(temp)<=160:time+=10
                elif int(temp)<=200:time+=18
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
            excel_list_2:list
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
                all_solid_records.append({
                    "reagent_name": reagent_name,
                    "hole_id": hole_id})
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
        table_num,excel_list,name_list = 1,[],[]
        # 第一张表处理
        s = 0
        for record in all_solid_records:
            if record["reagent_name"] in name_list:
                continue
            s += 1
            if s == 10:s += 1  # 跳过第10号位置
            if s > 14:break  # 第一表最多13个位置
            excel_list.append({
                "Code": self.name_key_dict[record["reagent_name"]],
                "Locate": s,
                "Type": 2})
            # 记录物料所属表号 (表1)
            self.reagent_table_map[record["reagent_name"]] = table_num
            name_list.append(record["reagent_name"])
        # 添加溶剂位置
        day_name = datetime.datetime.now().strftime("%Y%m%d")
        ExcelUtils.copy_workbook(
            f"{BASE_PATH1}/sowo物料.xlsx",
            TEMPLATE_PATHS["sowo"])
        # 保存第一张表
        wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["sowo"])
        ExcelUtils.write_records_to_sheet(wb, "Sheet1", excel_list)
        ExcelUtils.write_records_to_sheet(wb, "Sheet1", [{"Code": f"{day_name}-001", "Locate": 1, "Type": 1}])
        ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["sowo"])
        # 处理第二张表 (如果有超过13种物料)
        excel_list_2 = []
        if  len(unique_reagents) >13:
            table_num = 2  # 表号递增

            s = 0
            # 处理剩余物料
            for record in all_solid_records:
                if record["reagent_name"] in name_list:
                    continue
                s += 1
                if s == 10:s += 1  # 跳过第10号位置
                if s > 14:break  # 第一表最多13个位置
                excel_list_2.append({
                    "Code": self.name_key_dict[record["reagent_name"]],
                    "Locate": s,
                    "Type": 2})
                # 记录物料所属表号 (表2)
                self.reagent_table_map[record["reagent_name"]] = table_num
                name_list.append(record["reagent_name"])
            # 保存第二张表
            target_path = TEMPLATE_PATHS["sowo"].replace("1-1.xlsx", "2-1.xlsx")
            ExcelUtils.copy_workbook(f"{BASE_PATH1}/sowo物料.xlsx", target_path)
            wb = ExcelUtils.open_workbook(target_path)
            ExcelUtils.write_records_to_sheet(wb, "Sheet1", excel_list_2)
            ExcelUtils.write_records_to_sheet(wb, "Sheet1", [{"Code": f"{day_name}-001", "Locate": 1, "Type": 1}])
            ExcelUtils.save_workbook(wb, target_path)
            """只有8ml需要两个表格"""
        if self.post_data['experimentLog1']['wellPlates'] == '24孔板8ml':
            if len(self.data_dict) > 12:
                ExcelUtils.copy_workbook(
                    TEMPLATE_PATHS["sowo"],
                    TEMPLATE_PATHS["sowo"].replace('物料1-1', "物料1-2")
                )
            if len(unique_reagents) > 13 and len(self.data_dict) > 12:
                #
                ExcelUtils.copy_workbook(
                    TEMPLATE_PATHS["sowo"],
                    TEMPLATE_PATHS["sowo"].replace('物料1-1', "物料1-2")
                )
                target_path = TEMPLATE_PATHS["sowo"].replace("物料1-1.xlsx", "物料2-1.xlsx")
                ExcelUtils.copy_workbook(
                    target_path,
                    TEMPLATE_PATHS["sowo"].replace('物料1-1', "物料2-2")
                )
        else:
            pass
        self.write_codes_to_inner_hopper_stack(fr"{BASE_PATH}/Materials-stack4列.xlsx",fr"{BASE_PATH}/Materials-stack4列.xlsx",excel_list,excel_list_2)


    def _create_mate_msg(self) -> None:
        """创建mateMsg Excel并根据物料表号分表"""
        hole_to_coordinate = {
            1: (1, 1), 2: (1, 2), 3: (1, 3), 4: (1, 4),
            5: (2, 1), 6: (2, 2), 7: (2, 3), 8: (2, 4),
            9: (3, 1), 10: (3, 2), 11: (3, 3), 12: (3, 4),
            13: (4, 1), 14: (4, 2), 15: (4, 3), 16: (4, 4),
            17: (5, 1), 18: (5, 2), 19: (5, 3), 20: (5, 4),
            21: (6, 1), 22: (6, 2), 23: (6, 3), 24: (6, 4)}
        # 创建反应ID到孔位映射
        reaction_to_hole = {}
        for group_index, group in enumerate(self.perforated_plate):
            if group[0]:reaction_to_hole[group[0]] = group_index * 4 + 1
            if len(group) >= 2:reaction_to_hole[group[1]] = group_index * 4 + 2
            if len(group) >= 3:reaction_to_hole[group[2]] = group_index * 4 + 3
            if len(group) >= 4:reaction_to_hole[group[3]] = group_index * 4 + 4
        table_data = {1: [], 2: []}  # 存储分表数据
        for reagent_name, hole_data in self.solidity.items():
            for hole_id, details in hole_data.items():
                hole_str = str(hole_id)
                if hole_str in reaction_to_hole:
                    actual_hole = reaction_to_hole[hole_str]
                    x, y = hole_to_coordinate[actual_hole]
                    clean_name = reagent_name.strip()
                    target = details["weight"] if details["weight"] not in [None, ""] else 0
                    table_num = self.reagent_table_map.get(reagent_name, 1)
                    table_data[table_num].append({
                        "PointX": x,
                        "PointY": y,
                        "UsePlateType": 0,
                        "Name": clean_name,
                        "Target": round(float(target), 1),
                        "Tolerance": 0.5,
                        "UseMode": 0,
                        "Code": self.name_key_dict[reagent_name]})
        '''
        如果table_data[2]为空： 生成任务1-1，且table_data[1]>12 生成1-2

        如果table_data[2]不为空
        判断孔位小于12，只生成任务1-1 除了这种情况，其它都得分表
        '''
        if table_data[2] == [] and len(self.data_dict) <=12:
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
                    df.at[0, 'PlaterName'] = self.post_data["experimentLog1"]["wellPlates"]
                if self.type == 1:
                    # 根据表号生成文件名
                    if table_num == 1:
                        output_path = TEMPLATE_PATHS["mate"]
                else:
                    if table_num == 1:
                        output_path = crystallize_TEMPLATE_PATHS["mate"]
                df.to_excel(output_path, index=False)
        elif table_data[1] == [] and table_data[2] == []:
            #如果sowo为空，则需要生成空白任务表格
            if self.type == 1:
                df = pd.DataFrame(columns=[
                    "PointX", "PointY", "UsePlateType", "Name",
                    "Target", "Tolerance", "UseMode", "Code"])
                df['PlaterName'] = None
                # 仅第一行（索引0）写入 '96孔'
                if len(df) > 0:
                    df.at[0, 'PlaterName'] = self.post_data["experimentLog1"]["wellPlates"]
                df.to_excel(TEMPLATE_PATHS["mate"], index=False)
                if len(self.data_dict)>12:
                    df = pd.DataFrame(columns=[
                        "PointX", "PointY", "UsePlateType", "Name",
                        "Target", "Tolerance", "UseMode", "Code"])
                    df.to_excel(TEMPLATE_PATHS["mate"].replace('1-1.xlsx','1-2.xlsx'), index=False)
            elif self.type == 2:
                df = pd.DataFrame(columns=[
                    "PointX", "PointY", "UsePlateType", "Name",
                    "Target", "Tolerance", "UseMode", "Code"])
                df['PlaterName'] = None
                # 仅第一行（索引0）写入 '96孔'
                if len(df) > 0:
                    df.at[0, 'PlaterName'] = self.post_data["experimentLog1"]["wellPlates"]
                df.to_excel(crystallize_TEMPLATE_PATHS["mate"], index=False)
                if len(self.data_dict)>12:
                    df = pd.DataFrame(columns=[
                        "PointX", "PointY", "UsePlateType", "Name",
                        "Target", "Tolerance", "UseMode", "Code"])
                    df['PlaterName'] = None
                    # 仅第一行（索引0）写入 '96孔'
                    if len(df) > 0:
                        df.at[0, 'PlaterName'] = self.post_data["experimentLog1"]["wellPlates"]
                    df.to_excel(crystallize_TEMPLATE_PATHS["mate"].replace('1-1.xlsx','1-2.xlsx'), index=False)
        else:
            if self.post_data['experimentLog1']['wellPlates'] == '24孔板8ml':
                # 按名称排序
                grouped_table_data = {}
                for table_num, data_list in table_data.items():
                    if not data_list:  # 跳过空表
                        continue

                    # 按PointY分组
                    group1 = [item for item in data_list if item["PointY"] in [1, 4]]
                    group2 = [item for item in data_list if item["PointY"] in [2, 3]]

                    # 对每组数据进行排序
                    group1.sort(key=lambda x: (x["PointX"], x["PointY"]))
                    group2.sort(key=lambda x: (x["PointX"], x["PointY"]))

                    # 存储分组数据
                    grouped_table_data[f"{table_num}-1"] = group1  # 例如 "1-1" 或 "2-1"
                    grouped_table_data[f"{table_num}-2"] = group2  # 例如 "1-2" 或 "2-2"
                # 分别保存每个表
                for table_num, data_rows in grouped_table_data.items():
                    # 创建DataFrame并保存
                    df = pd.DataFrame(data_rows, columns=[
                        "PointX", "PointY", "UsePlateType", "Name",
                        "Target", "Tolerance", "UseMode", "Code"
                    ])
                    df['PlaterName'] = None
                    # 仅第一行（索引0）写入 '96孔'
                    if len(df) > 0:
                        df.at[0, 'PlaterName'] = self.post_data["experimentLog1"]["wellPlates"]
                    if self.type == 1:
                        # # 根据表号生成文件名(table_num的值为1-1，1-2，所以永远不会等于1)
                        # if table_num == 1:
                        #     output_path = TEMPLATE_PATHS["mate"].replace("1-1.xlsx", "")
                        #     output_path = f"{output_path}{table_num}.xlsx"
                        # else:
                            # 修改文件名以包含表号
                            base_name = TEMPLATE_PATHS["mate"].replace("1-1.xlsx", "")
                            output_path = f"{base_name}{table_num}.xlsx"
                    else:
                        # if table_num == 1:
                        #     output_path = crystallize_TEMPLATE_PATHS["mate"].replace("1-1.xlsx", "")
                        #     output_path = f"{output_path}{table_num}.xlsx"
                        # else:
                            # 修改文件名以包含表号
                            base_name = crystallize_TEMPLATE_PATHS["mate"].replace("1-1.xlsx", "")
                            output_path = f"{base_name}{table_num}.xlsx"

                    df.to_excel(output_path, index=False)
            else:
                # 分别保存每个表
                for table_num, data_rows in table_data.items():
                    if not data_rows:  # 跳过空表
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
                        df.at[0, 'PlaterName'] = self.post_data["experimentLog1"]["wellPlates"]
                    if self.type == 1:
                        # 根据表号生成文件名
                        if table_num == 1:
                            output_path = TEMPLATE_PATHS["mate"]
                        else:
                            output_path = TEMPLATE_PATHS["mate"].replace("1-1.xlsx", "2-1.xlsx")
                    else:
                        if table_num == 1:
                            output_path = crystallize_TEMPLATE_PATHS["mate"]
                        else:
                            output_path = crystallize_TEMPLATE_PATHS["mate"].replace("1-1.xlsx", "2-1.xlsx")
                    df.to_excel(output_path, index=False)


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
                                  ["FY24-B01", num, "ZZ-B02", num, int(diluent4) , pipette_location, gun_head + 1])
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
        ExcelUtils.batch_repair_folder(BASE_PATH)
    else:
        ExcelUtils.batch_repair_folder(crystallize_BASE_PATH)


if __name__ == "__main__":
    #逆合成
    data10 = {
    "experimentLog1": {
        "wellPlates": "24孔板8ml",
        "SubstratesTableList": [
            {
                "id": 1,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "reagent",
                "smiles": "",
                "substance": "C22H25ClF3NORu",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.6",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "reagent",
                "smiles": "",
                "substance": " hydrogen",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "reagent",
                "smiles": "",
                "substance": " acetic acid",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "solvent",
                "smiles": "",
                "substance": "methanol",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.63",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "solvent",
                "smiles": "",
                "substance": " ethanol",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.63",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "solvent",
                "smiles": "",
                "substance": " water",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.63",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "reagent",
                "smiles": "",
                "substance": "dmap",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.6",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "reagent",
                "smiles": "",
                "substance": " triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "solvent",
                "smiles": "",
                "substance": "tetrahydrofuran",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "reagent",
                "smiles": "",
                "substance": "triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "sub",
                "smiles": "C=CC(C)=O",
                "substance": "C=CC(C)=O",
                "molecularFormula": "C4H6O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 70.09,
                "reactionMoles": 1,
                "reactionQuality": 70.09,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "sub",
                "smiles": "Nc1ccccc1",
                "substance": "Nc1ccccc1",
                "molecularFormula": "C6H7N",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 93.13,
                "reactionMoles": 1,
                "reactionQuality": 93.13,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "reagent",
                "smiles": "",
                "substance": "triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "solvent",
                "smiles": "",
                "substance": "chloroform",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.40",
                "cas": "",
                "intensity": None
            }
        ],
        "MaterialsTableList": [
            {
                "type": "sub",
                "substance": "C=CC(C)=O",
                "smiles": "C=CC(C)=O",
                "solvent": 70.09,
                "molecularFormula": "C4H6O",
                "theoreticalMoles": "22.00",
                "theoreticalQuality": "1541.98",
                "theoreticalVolume": "44.00",
                "configurationMoles": "28.60",
                "configurationQuality": "2004.64",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 1,
                "concentration": "0.50"
            },
            {
                "type": "sub",
                "substance": "Nc1ccccc1",
                "smiles": "Nc1ccccc1",
                "solvent": 93.13,
                "molecularFormula": "C6H7N",
                "theoreticalMoles": "22.00",
                "theoreticalQuality": "2048.86",
                "theoreticalVolume": "22.00",
                "configurationMoles": "28.60",
                "configurationQuality": "2663.54",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 2,
                "concentration": "1.00"
            },
            {
                "type": "reagent",
                "substance": "C22H25ClF3NORu",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.60",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 3,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": " hydrogen",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.50",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 4,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": " acetic acid",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.40",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 5,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "methanol",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.63",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 6,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " ethanol",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.63",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 7,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " water",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.63",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 8,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "LiAlH4",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "8.00",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 9,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "THF",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "50.20",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 10,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "Borane-THF",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "3.00",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 11,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "dmap",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.60",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 12,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": " triethylamine",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.40",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 13,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "tetrahydrofuran",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.40",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 14,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "triethylamine",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 15,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "dichloromethane",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.40",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 16,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "chloroform",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.40",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 17,
                "concentration": ""
            }
        ]
    },
    "experimentLog2": {
        "ReactionConditions": {
            "temperature": "60",
            "time": "30",
            "speed": "700",
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
            ],
            "show": False
        },
        "electricityData": {
            "type": "1",
            "list": [],
            "show": False
        },
        "gasData": {
            "name": "",
            "pressure": "",
            "time": "",
            "show": False
        }
    },
    "experimentLog3": {
        "coolingTime": "60",
        "quenchingAgentData": [
            {
                "name": "llwc",
                "moles": "",
                "molecularWeight": "",
                "addVolume": "300",
                "quencherConcentration": "",
                "quenchVolume": "",
                "quencherQuality": ""
            }
        ],
        "internalStandardData": [
            {
                "name": "llwc1",
                "moles": "",
                "molecularWeight": "",
                "addVolume": "200",
                "internalConcentration": "",
                "internalVolume": "",
                "internalQuality": ""
            }
        ],
        "productDilutionData": [
            {
                "name": "MeCN",
                "beforeConcentration": "",
                "afterConcentration": "",
                "procedure": {
                    "diluent1": "100",
                    "diluent2": "100",
                    "diluent3": "100",
                    "diluent4": "100",
                    "diluent5": "100"
                }
            }
        ],
        "extractionData": [
            {
                "name": "o-Xylene",
                "volume": "100",
                "time": "30",
                "times": "2",
                "type": "2",
                "density": "0.881"
            }
        ],
        "filtrationData": [
            {
                "name": "硅胶",
                "filterTime": "50"
            }
        ]
    },
    "experimentLog4": {
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
        ]
    }
}
    data11 = {
    "experimentLog1": {
        "wellPlates": "8孔板20ml",
        "SubstratesTableList": [
            {
                "id": 1,
                "type": "sub",
                "smiles": "Clc1cc(Cl)cc(Cl)c1",
                "substance": "Clc1cc(Cl)cc(Cl)c1",
                "molecularFormula": "",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": "",
                "reactionMoles": 1,
                "reactionQuality": "",
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "reagent",
                "smiles": "",
                "substance": "tetrakis(triphenylphosphine) palladium(0)",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "67",
                "reactionMoles": "1.00",
                "reactionQuality": "67.00",
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "solvent",
                "smiles": "",
                "substance": "1,4-dioxane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "89",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "16.00",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "sub",
                "smiles": "Clc1cc(Cl)cc(Cl)c1",
                "substance": "Clc1cc(Cl)cc(Cl)c1",
                "molecularFormula": "",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": "",
                "reactionMoles": 1,
                "reactionQuality": "",
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "reagent",
                "smiles": "",
                "substance": "triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 16,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 16,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "14.00",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "sub",
                "smiles": "Clc1cc(Cl)cc(Cl)c1",
                "substance": "Clc1cc(Cl)cc(Cl)c1",
                "molecularFormula": "",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": "",
                "reactionMoles": 1,
                "reactionQuality": "",
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "reagent",
                "smiles": "",
                "substance": "potassium carbonate",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 16,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "solvent",
                "smiles": "",
                "substance": "octane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "120",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "13.00",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "sub",
                "smiles": "Clc1cc(Cl)cc(Cl)c1",
                "substance": "Clc1cc(Cl)cc(Cl)c1",
                "molecularFormula": "",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": "",
                "reactionMoles": 1,
                "reactionQuality": "",
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "reagent",
                "smiles": "",
                "substance": "sodium methylate",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 16,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "solvent",
                "smiles": "",
                "substance": "methanol",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 16,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.00",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "solvent",
                "smiles": "",
                "substance": " hexane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "60",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.00",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "solvent",
                "smiles": "",
                "substance": " chloroform",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 16,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.00",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "solvent",
                "smiles": "",
                "substance": " water",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 16,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.00",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "solvent",
                "smiles": "",
                "substance": " toluene",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "127",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.00",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "sub",
                "smiles": "Clc1cc(Cl)cc(Cl)c1",
                "substance": "Clc1cc(Cl)cc(Cl)c1",
                "molecularFormula": "",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": "81",
                "reactionMoles": "1.00",
                "reactionQuality": "81.00",
                "reactionVolume": 16,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "reagent",
                "smiles": "",
                "substance": "pyridine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 16,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 16,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "6.50",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "solvent",
                "smiles": "",
                "substance": " ethyl acetate",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "219",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 16,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "6.50",
                "cas": "",
                "intensity": None
            }
        ],
        "MaterialsTableList": [
            {
                "type": "sub",
                "substance": "Clc1cc(Cl)cc(Cl)c1",
                "smiles": "Clc1cc(Cl)cc(Cl)c1",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "5.00",
                "theoreticalQuality": "81.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "6.50",
                "cas": "",
                "intensity": None,
                "id": 1,
                "configurationVolume": "0.00",
                "configurationQuality": "105.30",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "tetrakis(triphenylphosphine) palladium(0)",
                "smiles": "",
                "solvent": "67",
                "molecularFormula": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "67.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "1.30",
                "cas": "",
                "intensity": None,
                "id": 2,
                "configurationVolume": "0.00",
                "configurationQuality": "87.10",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "1,4-dioxane",
                "smiles": "",
                "solvent": "89",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "16.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 3,
                "configurationVolume": "20.80",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "triethylamine",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 4,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "dichloromethane",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "20.50",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 5,
                "configurationVolume": "26.65",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "potassium carbonate",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "3.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 6,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "octane",
                "smiles": "",
                "solvent": "120",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "13.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 7,
                "configurationVolume": "16.90",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "sodium methylate",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 8,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "methanol",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "3.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 9,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " hexane",
                "smiles": "",
                "solvent": "60",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "3.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 10,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " chloroform",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "3.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 11,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " water",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "3.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 12,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " toluene",
                "smiles": "",
                "solvent": "127",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "3.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 13,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "pyridine",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "3.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 14,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " ethyl acetate",
                "smiles": "",
                "solvent": "219",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "6.50",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 15,
                "configurationVolume": "8.45",
                "configurationQuality": "0.00",
                "concentration": ""
            }
        ]
    },
    "experimentLog2": {
        "ReactionConditions": {
            "temperature": "120",
            "time": "200",
            "speed": "200",
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
            ],
            "show": True
        },
        "electricityData": {
            "type": "1",
            "list": [],
            "show": True
        },
        "gasData": {
            "name": "",
            "pressure": "",
            "time": "",
            "show": True
        }
    },
    "experimentLog3": {
        "coolingTime": "82",
        "quenchingAgentData": [
            {
                "name": "llwc",
                "moles": "",
                "molecularWeight": "",
                "addVolume": "300",
                "quencherConcentration": "",
                "quenchVolume": "",
                "quencherQuality": ""
            }
        ],
        "internalStandardData": [
            {
                "name": "llwc1",
                "moles": "",
                "molecularWeight": "",
                "addVolume": "400",
                "internalConcentration": "",
                "internalVolume": "",
                "internalQuality": ""
            }
        ],
        "productDilutionData": [],
        "extractionData": [
            {
                "name": "o-Xylene",
                "volume": "100",
                "time": "30",
                "times": "2",
                "type": "2",
                "density": "0.881"
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
        ]
    }
}
    data1 = {
    "experimentLog1": {
        "wellPlates": "24孔板4ml",
        "SubstratesTableList": [
            {
                "id": 1,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "reagent",
                "smiles": "",
                "substance": "sodium",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "reagent",
                "smiles": "",
                "substance": "pyridine!methanol!jones reagent!amberlyst-15",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "122.5",
                "reactionMoles": "2.00",
                "reactionQuality": "245.00",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "solvent",
                "smiles": "",
                "substance": "acetone",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "solvent",
                "smiles": "",
                "substance": " toluene",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "solvent",
                "smiles": "",
                "substance": "PPA",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "4.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "reagent",
                "smiles": "",
                "substance": "pyridine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "reagent",
                "smiles": "",
                "substance": "water!sodium carbonate!triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "solvent",
                "smiles": "",
                "substance": "diethyl ether",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "reagent",
                "smiles": "",
                "substance": "toluene-4-sulfonic acid",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "solvent",
                "smiles": "",
                "substance": "cyclohexane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "reagent",
                "smiles": "",
                "substance": "N,N-dimethyl-4-aminopyridine!dicyclohexyl-carbodiimide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "111.5",
                "reactionMoles": "2.00",
                "reactionQuality": "223.00",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "4.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "reagent",
                "smiles": "",
                "substance": "O-(benzotriazol-1-yl)-N,N,N',N'-tetramethyluronium tetrafluoroborate!triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "solvent",
                "smiles": "",
                "substance": "N,N-dimethyl-formamide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "4.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "reagent",
                "smiles": "",
                "substance": "dmap!dicyclohexyl-carbodiimide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "2.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "reagent",
                "smiles": "",
                "substance": "hydrogenchloride!water!sodium hydroxide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.45",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "solvent",
                "smiles": "",
                "substance": " water",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.45",
                "cas": "",
                "intensity": None
            }


        ],
        "MaterialsTableList": [
            {
                "type": "sub",
                "substance": "CCCCCCCC(=O)O",
                "smiles": "CCCCCCCC(=O)O",
                "solvent": 144.21,
                "molecularFormula": "C8H16O2",
                "theoreticalMoles": "24.00",
                "theoreticalQuality": "3461.04",
                "theoreticalVolume": "48.00",
                "configurationMoles": "31.20",
                "configurationQuality": "4499.28",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 1,
                "concentration": "0.50"
            },
            {
                "type": "reagent",
                "substance": "sodium",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 2,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "pyridine!methanol!jones reagent!amberlyst-15",
                "smiles": "",
                "solvent": "122.5",
                "molecularFormula": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "245.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "2.60",
                "configurationQuality": "318.50",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 3,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "acetone",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.20",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 4,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " toluene",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "4.40",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 5,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "PPA",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "4.40",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 6,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "pyridine",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 7,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "water!sodium carbonate!triethylamine",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 8,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "diethyl ether",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.40",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 9,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "toluene-4-sulfonic acid",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 10,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "cyclohexane",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.40",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 11,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "N,N-dimethyl-4-aminopyridine!dicyclohexyl-carbodiimide",
                "smiles": "",
                "solvent": "111.5",
                "molecularFormula": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "223.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "2.60",
                "configurationQuality": "289.90",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 12,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "dichloromethane",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "22.25",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 13,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "O-(benzotriazol-1-yl)-N,N,N',N'-tetramethyluronium tetrafluoroborate!triethylamine",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 14,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "N,N-dimethyl-formamide",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "13.20",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 15,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "dmap!dicyclohexyl-carbodiimide",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "4.60",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 16,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "hydrogenchloride!water!sodium hydroxide",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.50",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 17,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " water",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.45",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 18,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "potassium carbonate",
                "smiles": "",
                "solvent": "253.2",
                "molecularFormula": "",
                "theoreticalMoles": "3.00",
                "theoreticalQuality": "759.60",
                "theoreticalVolume": "0.00",
                "configurationMoles": "3.90",
                "configurationQuality": "987.48",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 19,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "sodium tris(acetoxy)borohydride",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.30",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 20,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "N,N-dimethyl acetamide",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.05",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 21,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " N,N-dimethyl acetamide",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.05",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 22,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "trimethylamine",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.50",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 23,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "acetonitrile",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.90",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 24,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "silver carbonate",
                "smiles": "",
                "solvent": "185.4",
                "molecularFormula": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "370.80",
                "theoreticalVolume": "0.00",
                "configurationMoles": "2.60",
                "configurationQuality": "482.04",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 25,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "1-methyl-pyrrolidin-2-one",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.20",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 26,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "triethylamine!copper(l) chloride",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.20",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 27,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "sodium hydride",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 28,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "tetrahydrofuran",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.70",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 29,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " tetrahydrofuran",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.70",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 30,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "copper-iron mixed oxide",
                "smiles": "",
                "solvent": "324.5",
                "molecularFormula": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "649.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "2.60",
                "configurationQuality": "843.70",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 31,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " acetonitrile",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.20",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 32,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "trifluoroacetic acid",
                "smiles": "",
                "solvent": "243.1",
                "molecularFormula": "",
                "theoreticalMoles": "2.50",
                "theoreticalQuality": "607.75",
                "theoreticalVolume": "0.00",
                "configurationMoles": "3.25",
                "configurationQuality": "790.08",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 33,
                "concentration": ""
            }
        ]
    },
    "experimentLog2": {
        "ReactionConditions": {
            "temperature": "50",
            "time": "300",
            "speed": "800",
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
            ],
            "show": False
        },
        "electricityData": {
            "type": "1",
            "list": [],
            "show": False
        },
        "gasData": {
            "name": "",
            "pressure": "",
            "time": "",
            "show": False
        }
    },
    "experimentLog3": {
        "coolingTime": "60",
        "quenchingAgentData": [
            {
                "name": "llwc",
                "moles": "2",
                "molecularWeight": "",
                "addVolume": "700",
                "quencherConcentration": "",
                "quenchVolume": "",
                "quencherQuality": ""
            }
        ],
        "internalStandardData": [
            {
                "name": "llwc1",
                "moles": "",
                "molecularWeight": "",
                "addVolume": "500",
                "internalConcentration": "",
                "internalVolume": "",
                "internalQuality": ""
            }
        ],
        "productDilutionData": [
            {
                "name": "MeCN",
                "beforeConcentration": "",
                "afterConcentration": "",
                "procedure": {
                    "diluent1": "100",
                    "diluent2": "100",
                    "diluent3": "20",
                    "diluent4": "50",
                    "diluent5": "100"
                }
            }
        ],
        "extractionData": [
            {
                "name": "o-Xylene",
                "volume": "300",
                "time": "60",
                "times": "3",
                "type": "2",
                "density": "0.881"
            }
        ],
        "filtrationData": [
            {
                "name": "硅胶",
                "filterTime": "70"
            }
        ]
    },
    "experimentLog4": {
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
        ]
    }
}
    data2 = {
    "experimentLog1": {
        "wellPlates": "24孔板8ml",
        "SubstratesTableList": [
            {
                "id": 1,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "reagent",
                "smiles": "",
                "substance": "sodium",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "reagent",
                "smiles": "",
                "substance": "pyridine!methanol!jones reagent!amberlyst-15",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "122.5",
                "reactionMoles": "2.00",
                "reactionQuality": "245.00",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "solvent",
                "smiles": "",
                "substance": "acetone",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "solvent",
                "smiles": "",
                "substance": " toluene",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "solvent",
                "smiles": "",
                "substance": "PPA",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "4.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "reagent",
                "smiles": "",
                "substance": "pyridine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "reagent",
                "smiles": "",
                "substance": "water!sodium carbonate!triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "solvent",
                "smiles": "",
                "substance": "diethyl ether",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "reagent",
                "smiles": "",
                "substance": "toluene-4-sulfonic acid",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "solvent",
                "smiles": "",
                "substance": "cyclohexane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "reagent",
                "smiles": "",
                "substance": "N,N-dimethyl-4-aminopyridine!dicyclohexyl-carbodiimide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "111.5",
                "reactionMoles": "2.00",
                "reactionQuality": "223.00",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "4.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "reagent",
                "smiles": "",
                "substance": "O-(benzotriazol-1-yl)-N,N,N',N'-tetramethyluronium tetrafluoroborate!triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "solvent",
                "smiles": "",
                "substance": "N,N-dimethyl-formamide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "4.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "reagent",
                "smiles": "",
                "substance": "dmap!dicyclohexyl-carbodiimide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "2.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "reagent",
                "smiles": "",
                "substance": "hydrogenchloride!water!sodium hydroxide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.45",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "solvent",
                "smiles": "",
                "substance": " water",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.45",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "reagent",
                "smiles": "",
                "substance": "potassium carbonate",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "3",
                "solvent": "253.2",
                "reactionMoles": "3.00",
                "reactionQuality": "759.60",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "solvent",
                "smiles": "",
                "substance": "N,N-dimethyl-formamide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "4.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "reagent",
                "smiles": "",
                "substance": "sodium tris(acetoxy)borohydride",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "2.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "solvent",
                "smiles": "",
                "substance": "N,N-dimethyl acetamide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.05",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "solvent",
                "smiles": "",
                "substance": " N,N-dimethyl acetamide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.05",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "reagent",
                "smiles": "",
                "substance": "trimethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "2.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "solvent",
                "smiles": "",
                "substance": "acetonitrile",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "reagent",
                "smiles": "",
                "substance": "dmap!dicyclohexyl-carbodiimide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "2.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "reagent",
                "smiles": "",
                "substance": "pyridine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "reagent",
                "smiles": "",
                "substance": "silver carbonate",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "185.4",
                "reactionMoles": "2.00",
                "reactionQuality": "370.80",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "solvent",
                "smiles": "",
                "substance": "1-methyl-pyrrolidin-2-one",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "solvent",
                "smiles": "",
                "substance": " toluene",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "reagent",
                "smiles": "",
                "substance": "potassium carbonate",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "solvent",
                "smiles": "",
                "substance": "N,N-dimethyl-formamide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "4.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "reagent",
                "smiles": "",
                "substance": "triethylamine!copper(l) chloride",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "2.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "reagent",
                "smiles": "",
                "substance": "sodium hydride",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "solvent",
                "smiles": "",
                "substance": "tetrahydrofuran",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.70",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "solvent",
                "smiles": "",
                "substance": " tetrahydrofuran",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.70",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "reagent",
                "smiles": "",
                "substance": "copper-iron mixed oxide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "324.5",
                "reactionMoles": "2.00",
                "reactionQuality": "649.00",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "solvent",
                "smiles": "",
                "substance": " acetonitrile",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "sub",
                "smiles": "CCCCCCCC(=O)O",
                "substance": "CCCCCCCC(=O)O",
                "molecularFormula": "C8H16O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 144.21,
                "reactionMoles": 1,
                "reactionQuality": 144.21,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.50",
                "singleCockAddVolume": "2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "reagent",
                "smiles": "",
                "substance": "trifluoroacetic acid",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2.5",
                "solvent": "243.1",
                "reactionMoles": "2.50",
                "reactionQuality": "607.75",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "4.40",
                "cas": "",
                "intensity": None
            }
        ],
        "MaterialsTableList": [
            {
                "type": "sub",
                "substance": "CCCCCCCC(=O)O",
                "smiles": "CCCCCCCC(=O)O",
                "solvent": 144.21,
                "molecularFormula": "C8H16O2",
                "theoreticalMoles": "24.00",
                "theoreticalQuality": "3461.04",
                "theoreticalVolume": "48.00",
                "configurationMoles": "31.20",
                "configurationQuality": "4499.28",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 1,
                "concentration": "0.50"
            },
            {
                "type": "reagent",
                "substance": "sodium",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 2,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "pyridine!methanol!jones reagent!amberlyst-15",
                "smiles": "",
                "solvent": "122.5",
                "molecularFormula": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "245.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "2.60",
                "configurationQuality": "318.50",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 3,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "acetone",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.20",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 4,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " toluene",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "4.40",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 5,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "PPA",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "4.40",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 6,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "pyridine",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 7,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "water!sodium carbonate!triethylamine",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 8,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "diethyl ether",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.40",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 9,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "toluene-4-sulfonic acid",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 10,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "cyclohexane",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.40",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 11,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "N,N-dimethyl-4-aminopyridine!dicyclohexyl-carbodiimide",
                "smiles": "",
                "solvent": "111.5",
                "molecularFormula": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "223.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "2.60",
                "configurationQuality": "289.90",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 12,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "dichloromethane",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "22.25",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 13,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "O-(benzotriazol-1-yl)-N,N,N',N'-tetramethyluronium tetrafluoroborate!triethylamine",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 14,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "N,N-dimethyl-formamide",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "13.20",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 15,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "dmap!dicyclohexyl-carbodiimide",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "4.60",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 16,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "hydrogenchloride!water!sodium hydroxide",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.50",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 17,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " water",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.45",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 18,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "potassium carbonate",
                "smiles": "",
                "solvent": "253.2",
                "molecularFormula": "",
                "theoreticalMoles": "3.00",
                "theoreticalQuality": "759.60",
                "theoreticalVolume": "0.00",
                "configurationMoles": "3.90",
                "configurationQuality": "987.48",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 19,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "sodium tris(acetoxy)borohydride",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.30",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 20,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "N,N-dimethyl acetamide",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.05",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 21,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " N,N-dimethyl acetamide",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.05",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 22,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "trimethylamine",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.50",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 23,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "acetonitrile",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.90",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 24,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "silver carbonate",
                "smiles": "",
                "solvent": "185.4",
                "molecularFormula": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "370.80",
                "theoreticalVolume": "0.00",
                "configurationMoles": "2.60",
                "configurationQuality": "482.04",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 25,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "1-methyl-pyrrolidin-2-one",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.20",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 26,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "triethylamine!copper(l) chloride",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.20",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 27,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "sodium hydride",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 28,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "tetrahydrofuran",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.70",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 29,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " tetrahydrofuran",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.70",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 30,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "copper-iron mixed oxide",
                "smiles": "",
                "solvent": "324.5",
                "molecularFormula": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "649.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "2.60",
                "configurationQuality": "843.70",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 31,
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": " acetonitrile",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.20",
                "configurationMoles": "0.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "cas": "",
                "intensity": None,
                "id": 32,
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "trifluoroacetic acid",
                "smiles": "",
                "solvent": "243.1",
                "molecularFormula": "",
                "theoreticalMoles": "2.50",
                "theoreticalQuality": "607.75",
                "theoreticalVolume": "0.00",
                "configurationMoles": "3.25",
                "configurationQuality": "790.08",
                "configurationVolume": "0.00",
                "cas": "",
                "intensity": None,
                "id": 33,
                "concentration": ""
            }
        ]
    },
    "experimentLog2": {
        "ReactionConditions": {
            "temperature": "50",
            "time": "300",
            "speed": "800",
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
            ],
            "show": False
        },
        "electricityData": {
            "type": "1",
            "list": [],
            "show": False
        },
        "gasData": {
            "name": "",
            "pressure": "",
            "time": "",
            "show": False
        }
    },
    "experimentLog3": {
        "coolingTime": "60",
        "quenchingAgentData": [
            {
                "name": "llwc",
                "moles": "2",
                "molecularWeight": "",
                "addVolume": "200",
                "quencherConcentration": "",
                "quenchVolume": "",
                "quencherQuality": ""
            }
        ],
        "internalStandardData": [
            {
                "name": "llwc1",
                "moles": "",
                "molecularWeight": "",
                "addVolume": "200",
                "internalConcentration": "",
                "internalVolume": "",
                "internalQuality": ""
            }
        ],
        "productDilutionData": [
            {
                "name": "MeCN",
                "beforeConcentration": "",
                "afterConcentration": "",
                "procedure": {
                    "diluent1": "100",
                    "diluent2": "100",
                    "diluent3": "20",
                    "diluent4": "50",
                    "diluent5": "100"
                }
            }
        ],
        "extractionData": [
            {
                "name": "o-Xylene",
                "volume": "300",
                "time": "60",
                "times": "3",
                "type": "2",
                "density": "0.881"
            }
        ],
        "filtrationData": [
            {
                "name": "硅胶",
                "filterTime": "70"
            }
        ]
    },
    "experimentLog4": {
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
        ]
    }
}
    data3 = {
    "experimentLog1": {
        "wellPlates": "24孔板8ml",
        "SubstratesTableList": [
            {
                "id": 1,
                "type": "固体物料",
                "substance": "零零",
                "molecularFormula": "",
                "molecularWeight": "235",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "470.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 1,
                "type": "液体物料",
                "substance": "ll1",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 1,
                "type": "溶剂",
                "substance": "ll2",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 1,
                "type": "反溶剂",
                "substance": "ll3",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 2,
                "type": "固体物料",
                "substance": "零零",
                "molecularFormula": "",
                "molecularWeight": "235",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "470.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 2,
                "type": "液体物料",
                "substance": "ll2",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 2,
                "type": "溶剂",
                "substance": "ll4",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 2,
                "type": "反溶剂",
                "substance": "ll3",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 3,
                "type": "固体物料",
                "substance": "零零",
                "molecularFormula": "",
                "molecularWeight": "235",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "470.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 3,
                "type": "液体物料",
                "substance": "ll2",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 3,
                "type": "溶剂",
                "substance": "ll1",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 3,
                "type": "反溶剂",
                "substance": "ll3",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 4,
                "type": "固体物料",
                "substance": "零零",
                "molecularFormula": "",
                "molecularWeight": "235",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "470.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 4,
                "type": "液体物料",
                "substance": "ll4",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 4,
                "type": "溶剂",
                "substance": "ll5",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 4,
                "type": "反溶剂",
                "substance": "ll3",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            }
        ],
        "MaterialsTableList": [
            {
                "type": "固体物料",
                "substance": "零零",
                "molecularFormula": "",
                "molecularWeight": "235",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "8.00",
                "theoreticalQuality": "1880.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "2444.00",
                "configurationVolume": "0.00",
                "id": 1,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "ll1",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "3.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 2,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "ll2",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "4.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 3,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "ll3",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "4.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 4,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "ll4",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "3.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 5,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "ll5",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 6,
                "concentration": ""
            }
        ],
        "isType": "1"
    },
    "experimentLog2": {
        "dissolveForm": {
            "temperature": "60",
            "maxTemperature": "80",
            "time": "60",
            "speed": "",
            "rjSpeed": "800",
            "seal": "1",
            "environment": "1",
            "dissolve": "1",
            "dissolveTimeInterval": "30",
            "dissolveMaximum": "3",
            "dissolveGradient": "5"
        },
        "postProcessingForm": {
            "time": "30",
            "temperature": "30",
            "maxTemperature": "",
            "speed": "700",
            "method": "1",
            "precipitate": "1",
            "precoolingTemperature": "30"
        },
        "filtrationForm": {
            "type": "硅胶",
            "time": "30"
        }
    },
    "experimentLog3": {
        "dissolveForm": {},
        "postProcessingForm": {},
        "filtrationForm": {}
    }
}
    #正加
    data4 = {
    "experimentLog1": {
        "wellPlates": "24孔板2ml",
        "SubstratesTableList": [
            {
                "id": 1,
                "type": "固体物料",
                "substance": "111",
                "molecularFormula": "",
                "molecularWeight": "123",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "246.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 1,
                "type": "液体物料",
                "substance": "112",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 1,
                "type": "溶剂",
                "substance": "113",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 1,
                "type": "反溶剂",
                "substance": "114",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 2,
                "type": "固体物料",
                "substance": "121",
                "molecularFormula": "",
                "molecularWeight": "231",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "462.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 2,
                "type": "液体物料",
                "substance": "122",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 2,
                "type": "溶剂",
                "substance": "123",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 2,
                "type": "反溶剂",
                "substance": "124",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 3,
                "type": "固体物料",
                "substance": "125",
                "molecularFormula": "",
                "molecularWeight": "213",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "426.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 3,
                "type": "液体物料",
                "substance": "127",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 3,
                "type": "溶剂",
                "substance": "131",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 3,
                "type": "反溶剂",
                "substance": "132",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 4,
                "type": "固体物料",
                "substance": "133",
                "molecularFormula": "",
                "molecularWeight": "241",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "241.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 4,
                "type": "液体物料",
                "substance": "134",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 4,
                "type": "溶剂",
                "substance": "143",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 4,
                "type": "反溶剂",
                "substance": "132",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 5,
                "type": "固体物料",
                "substance": "154",
                "molecularFormula": "",
                "molecularWeight": "153",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "306.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 5,
                "type": "液体物料",
                "substance": "211",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 5,
                "type": "溶剂",
                "substance": "212",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 5,
                "type": "反溶剂",
                "substance": "213",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 6,
                "type": "固体物料",
                "substance": "233",
                "molecularFormula": "",
                "molecularWeight": "271",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "271.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 6,
                "type": "液体物料",
                "substance": "234",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 6,
                "type": "溶剂",
                "substance": "235",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 6,
                "type": "反溶剂",
                "substance": "243",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 7,
                "type": "固体物料",
                "substance": "241",
                "molecularFormula": "",
                "molecularWeight": "182",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "364.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 7,
                "type": "液体物料",
                "substance": "242",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 7,
                "type": "溶剂",
                "substance": "243",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 7,
                "type": "反溶剂",
                "substance": "244",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 8,
                "type": "固体物料",
                "substance": "311",
                "molecularFormula": "",
                "molecularWeight": "217",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "434.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 8,
                "type": "液体物料",
                "substance": "123",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 8,
                "type": "溶剂",
                "substance": "321",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 8,
                "type": "反溶剂",
                "substance": "323",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 9,
                "type": "固体物料",
                "substance": "453",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 9,
                "type": "液体物料",
                "substance": "444",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 9,
                "type": "溶剂",
                "substance": "432",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 9,
                "type": "反溶剂",
                "substance": "431",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 10,
                "type": "固体物料",
                "substance": "4531",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 10,
                "type": "液体物料",
                "substance": "4441",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 10,
                "type": "溶剂",
                "substance": "4321",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 10,
                "type": "反溶剂",
                "substance": "4311",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 11,
                "type": "固体物料",
                "substance": "4532",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 11,
                "type": "液体物料",
                "substance": "4442",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 11,
                "type": "溶剂",
                "substance": "4322",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 11,
                "type": "反溶剂",
                "substance": "4312",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 12,
                "type": "固体物料",
                "substance": "4533",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 12,
                "type": "液体物料",
                "substance": "4443",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 12,
                "type": "溶剂",
                "substance": "4323",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 12,
                "type": "反溶剂",
                "substance": "4313",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 9,
                "type": "固体物料",
                "substance": "453",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 13,
                "type": "液体物料",
                "substance": "4444",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 13,
                "type": "溶剂",
                "substance": "4324",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 13,
                "type": "反溶剂",
                "substance": "4314",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 14,
                "type": "固体物料",
                "substance": "45356",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 14,
                "type": "液体物料",
                "substance": "4446",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 14,
                "type": "溶剂",
                "substance": "43256",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 14,
                "type": "反溶剂",
                "substance": "4356",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 15,
                "type": "固体物料",
                "substance": "4535",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 15,
                "type": "液体物料",
                "substance": "4445",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 15,
                "type": "溶剂",
                "substance": "4325",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 15,
                "type": "反溶剂",
                "substance": "435",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            }
        ],
        "MaterialsTableList": [
            {
                "type": "固体物料",
                "substance": "111",
                "molecularFormula": "",
                "molecularWeight": "123",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "246.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "319.80",
                "configurationVolume": "0.00",
                "id": 1,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "112",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 2,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "113",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 3,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "114",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 4,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "121",
                "molecularFormula": "",
                "molecularWeight": "231",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "462.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "600.60",
                "configurationVolume": "0.00",
                "id": 5,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "122",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 6,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "123",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "4.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 7,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "124",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 8,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "125",
                "molecularFormula": "",
                "molecularWeight": "213",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "426.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "553.80",
                "configurationVolume": "0.00",
                "id": 9,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "127",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 10,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "131",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 11,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "132",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "3.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 12,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "133",
                "molecularFormula": "",
                "molecularWeight": "241",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "241.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "313.30",
                "configurationVolume": "0.00",
                "id": 13,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "134",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 14,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "143",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 15,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "154",
                "molecularFormula": "",
                "molecularWeight": "153",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "306.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "397.80",
                "configurationVolume": "0.00",
                "id": 16,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "211",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 17,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "212",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 18,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "213",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 19,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "233",
                "molecularFormula": "",
                "molecularWeight": "271",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "271.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "352.30",
                "configurationVolume": "0.00",
                "id": 20,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "234",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 21,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "235",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 22,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "243",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "4.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 23,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "241",
                "molecularFormula": "",
                "molecularWeight": "182",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "364.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "473.20",
                "configurationVolume": "0.00",
                "id": 24,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "242",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 25,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "244",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 26,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "311",
                "molecularFormula": "",
                "molecularWeight": "217",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "434.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "564.20",
                "configurationVolume": "0.00",
                "id": 27,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "321",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 28,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "323",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 29,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "453",
                "molecularFormula": "",
                "molecularWeight": "253",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "253.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "328.90",
                "configurationVolume": "0.00",
                "id": 30,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "444",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 31,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "432",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 32,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "431",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 33,
                "concentration": ""
            }
        ],
        "isType": "1"
    },
    "experimentLog2": {
        "dissolveForm": {
            "temperature": "70",
            "maxTemperature": "110",
            "time": "60",
            "speed": "",
            "rjSpeed": "700",
            "seal": "1",
            "environment": "1",
            "dissolve": "1",
            "dissolveTimeInterval": "30",
            "dissolveMaximum": "4",
            "dissolveGradient": "5"
        },
        "postProcessingForm": {
            "time": "30",
            "temperature": "60",
            "maxTemperature": "",
            "speed": "800",
            "method": "1",
            "precipitate": "1",
            "precoolingTemperature": "40"
        },
        "filtrationForm": {
            "type": "硅胶",
            "time": "80"
        }
    },
    "experimentLog3": {
        "dissolveForm": {
            "temperature": "60",
            "maxTemperature": "110",
            "time": "70",
            "speed": "",
            "rjSpeed": "800",
            "seal": "1",
            "environment": "1",
            "dissolve": "1",
            "dissolveTimeInterval": 30,
            "dissolveMaximum": "2",
            "dissolveGradient": "5"
        },
        "postProcessingForm": {
            "time": "30",
            "maxTemperature": "",
            "preCooling": "1",
            "temperature": "70",
            "speed": "700",
            "precipitate": "1",
            "dissolveTimeInterval": "30",
            "dissolveMaximum": "3",
            "dissolveGradient": "5"
        },
        "filtrationForm": {
            "type": "硅胶",
            "time": "35"
        }
    }
}
    #反加
    data5 = {
        "id":125461682,
    "experimentLog1": {
        "wellPlates": "24孔板2ml",
        "SubstratesTableList": [
            {
                "id": 1,
                "type": "固体物料",
                "substance": "111",
                "molecularFormula": "",
                "molecularWeight": "123",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "246.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 1,
                "type": "液体物料",
                "substance": "112",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 1,
                "type": "溶剂",
                "substance": "113",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 1,
                "type": "反溶剂",
                "substance": "114",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 2,
                "type": "固体物料",
                "substance": "121",
                "molecularFormula": "",
                "molecularWeight": "231",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "462.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 2,
                "type": "液体物料",
                "substance": "122",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 2,
                "type": "溶剂",
                "substance": "123",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 2,
                "type": "反溶剂",
                "substance": "124",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 3,
                "type": "固体物料",
                "substance": "125",
                "molecularFormula": "",
                "molecularWeight": "213",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "426.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 3,
                "type": "液体物料",
                "substance": "127",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 3,
                "type": "溶剂",
                "substance": "131",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 3,
                "type": "反溶剂",
                "substance": "132",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 4,
                "type": "固体物料",
                "substance": "133",
                "molecularFormula": "",
                "molecularWeight": "241",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "241.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 4,
                "type": "液体物料",
                "substance": "134",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 4,
                "type": "溶剂",
                "substance": "143",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 4,
                "type": "反溶剂",
                "substance": "132",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 5,
                "type": "固体物料",
                "substance": "154",
                "molecularFormula": "",
                "molecularWeight": "153",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "306.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 5,
                "type": "液体物料",
                "substance": "211",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 5,
                "type": "溶剂",
                "substance": "212",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 5,
                "type": "反溶剂",
                "substance": "213",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 6,
                "type": "固体物料",
                "substance": "233",
                "molecularFormula": "",
                "molecularWeight": "271",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "271.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 6,
                "type": "液体物料",
                "substance": "234",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 6,
                "type": "溶剂",
                "substance": "235",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 6,
                "type": "反溶剂",
                "substance": "243",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 7,
                "type": "固体物料",
                "substance": "241",
                "molecularFormula": "",
                "molecularWeight": "182",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "364.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 7,
                "type": "液体物料",
                "substance": "242",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 7,
                "type": "溶剂",
                "substance": "243",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 7,
                "type": "反溶剂",
                "substance": "244",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 8,
                "type": "固体物料",
                "substance": "311",
                "molecularFormula": "",
                "molecularWeight": "217",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "434.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 8,
                "type": "液体物料",
                "substance": "123",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 8,
                "type": "溶剂",
                "substance": "321",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 8,
                "type": "反溶剂",
                "substance": "323",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 9,
                "type": "固体物料",
                "substance": "453",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 9,
                "type": "液体物料",
                "substance": "444",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 9,
                "type": "溶剂",
                "substance": "432",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 9,
                "type": "反溶剂",
                "substance": "431",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 10,
                "type": "固体物料",
                "substance": "4531",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 10,
                "type": "液体物料",
                "substance": "4441",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 10,
                "type": "溶剂",
                "substance": "4321",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 10,
                "type": "反溶剂",
                "substance": "4311",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 11,
                "type": "固体物料",
                "substance": "4532",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 11,
                "type": "液体物料",
                "substance": "4442",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 11,
                "type": "溶剂",
                "substance": "4322",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 11,
                "type": "反溶剂",
                "substance": "4312",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 12,
                "type": "固体物料",
                "substance": "4535",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 12,
                "type": "液体物料",
                "substance": "4445",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 12,
                "type": "溶剂",
                "substance": "4325",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 12,
                "type": "反溶剂",
                "substance": "4315",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 13,
                "type": "固体物料",
                "substance": "4533",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 13,
                "type": "液体物料",
                "substance": "4443",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 13,
                "type": "溶剂",
                "substance": "4323",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 13,
                "type": "反溶剂",
                "substance": "4313",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            }
        ],
        "MaterialsTableList": [
            {
                "type": "固体物料",
                "substance": "111",
                "molecularFormula": "",
                "molecularWeight": "123",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "246.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "319.80",
                "configurationVolume": "0.00",
                "id": 1,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "112",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 2,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "113",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 3,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "114",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 4,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "121",
                "molecularFormula": "",
                "molecularWeight": "231",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "462.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "600.60",
                "configurationVolume": "0.00",
                "id": 5,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "122",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 6,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "123",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "4.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 7,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "124",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 8,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "125",
                "molecularFormula": "",
                "molecularWeight": "213",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "426.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "553.80",
                "configurationVolume": "0.00",
                "id": 9,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "127",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 10,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "131",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 11,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "132",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "3.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 12,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "133",
                "molecularFormula": "",
                "molecularWeight": "241",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "241.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "313.30",
                "configurationVolume": "0.00",
                "id": 13,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "134",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 14,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "143",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 15,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "154",
                "molecularFormula": "",
                "molecularWeight": "153",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "306.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "397.80",
                "configurationVolume": "0.00",
                "id": 16,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "211",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 17,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "212",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 18,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "213",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 19,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "233",
                "molecularFormula": "",
                "molecularWeight": "271",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "271.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "352.30",
                "configurationVolume": "0.00",
                "id": 20,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "234",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 21,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "235",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 22,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "243",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "4.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 23,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "241",
                "molecularFormula": "",
                "molecularWeight": "182",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "364.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "473.20",
                "configurationVolume": "0.00",
                "id": 24,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "242",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 25,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "244",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 26,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "311",
                "molecularFormula": "",
                "molecularWeight": "217",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "434.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "564.20",
                "configurationVolume": "0.00",
                "id": 27,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "321",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 28,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "323",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 29,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "453",
                "molecularFormula": "",
                "molecularWeight": "253",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "253.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "328.90",
                "configurationVolume": "0.00",
                "id": 30,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "444",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 31,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "432",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 32,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "431",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 33,
                "concentration": ""
            }
        ],
        "isType": "1"
    },
    "experimentLog2": {
        "dissolveForm": {
            "temperature": "70",
            "maxTemperature": "110",
            "time": "60",
            "speed": "",
            "rjSpeed": "700",
            "seal": "1",
            "environment": "1",
            "dissolve": "1",
            "dissolveTimeInterval": "30",
            "dissolveMaximum": "4",
            "dissolveGradient": "5"
        },
        "postProcessingForm": {
            "time": "30",
            "temperature": "60",
            "maxTemperature": "",
            "speed": "800",
            "method": "2",
            "precipitate": "1",
            "precoolingTemperature": "40"
        },
        "filtrationForm": {
            "type": "硅胶",
            "time": "80"
        }
    },
    "experimentLog3": {
        "dissolveForm": {
            "temperature": "60",
            "maxTemperature": "110",
            "time": "70",
            "speed": "",
            "rjSpeed": "800",
            "seal": "1",
            "environment": "1",
            "dissolve": "1",
            "dissolveTimeInterval": 30,
            "dissolveMaximum": "2",
            "dissolveGradient": "5"
        },
        "postProcessingForm": {
            "time": "30",
            "maxTemperature": "",
            "preCooling": "1",
            "temperature": "70",
            "speed": "700",
            "precipitate": "1",
            "dissolveTimeInterval": "30",
            "dissolveMaximum": "3",
            "dissolveGradient": "5"
        },
        "filtrationForm": {
            "type": "硅胶",
            "time": "35"
        }
    }
}
    #降温
    data = {
    "experimentLog1": {
        "wellPlates": "24孔板2ml",
        "SubstratesTableList": [
            {
                "id": 1,
                "type": "固体物料",
                "substance": "111",
                "molecularFormula": "",
                "molecularWeight": "123",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "246.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 1,
                "type": "液体物料",
                "substance": "112",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 1,
                "type": "溶剂",
                "substance": "113",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 1,
                "type": "反溶剂",
                "substance": "114",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 2,
                "type": "固体物料",
                "substance": "121",
                "molecularFormula": "",
                "molecularWeight": "231",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "462.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 2,
                "type": "液体物料",
                "substance": "122",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 2,
                "type": "溶剂",
                "substance": "123",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 2,
                "type": "反溶剂",
                "substance": "124",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 3,
                "type": "固体物料",
                "substance": "125",
                "molecularFormula": "",
                "molecularWeight": "213",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "426.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 3,
                "type": "液体物料",
                "substance": "127",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 3,
                "type": "溶剂",
                "substance": "131",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 3,
                "type": "反溶剂",
                "substance": "132",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 4,
                "type": "固体物料",
                "substance": "133",
                "molecularFormula": "",
                "molecularWeight": "241",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "241.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 4,
                "type": "液体物料",
                "substance": "134",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 4,
                "type": "溶剂",
                "substance": "143",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 4,
                "type": "反溶剂",
                "substance": "132",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 5,
                "type": "固体物料",
                "substance": "154",
                "molecularFormula": "",
                "molecularWeight": "153",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "306.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 5,
                "type": "液体物料",
                "substance": "211",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 5,
                "type": "溶剂",
                "substance": "212",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 5,
                "type": "反溶剂",
                "substance": "213",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 6,
                "type": "固体物料",
                "substance": "233",
                "molecularFormula": "",
                "molecularWeight": "271",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "271.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 6,
                "type": "液体物料",
                "substance": "234",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 6,
                "type": "溶剂",
                "substance": "235",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 6,
                "type": "反溶剂",
                "substance": "243",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 7,
                "type": "固体物料",
                "substance": "241",
                "molecularFormula": "",
                "molecularWeight": "182",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "364.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 7,
                "type": "液体物料",
                "substance": "242",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 7,
                "type": "溶剂",
                "substance": "243",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 7,
                "type": "反溶剂",
                "substance": "244",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 8,
                "type": "固体物料",
                "substance": "311",
                "molecularFormula": "",
                "molecularWeight": "217",
                "moles": "2",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "434.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 8,
                "type": "液体物料",
                "substance": "123",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 8,
                "type": "溶剂",
                "substance": "321",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 8,
                "type": "反溶剂",
                "substance": "323",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 9,
                "type": "固体物料",
                "substance": "453",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 9,
                "type": "液体物料",
                "substance": "444",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 9,
                "type": "溶剂",
                "substance": "432",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 9,
                "type": "反溶剂",
                "substance": "431",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 10,
                "type": "固体物料",
                "substance": "4531",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 10,
                "type": "液体物料",
                "substance": "4441",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 10,
                "type": "溶剂",
                "substance": "4321",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 10,
                "type": "反溶剂",
                "substance": "4311",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 11,
                "type": "固体物料",
                "substance": "4532",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 11,
                "type": "液体物料",
                "substance": "4442",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 11,
                "type": "溶剂",
                "substance": "4322",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 11,
                "type": "反溶剂",
                "substance": "4312",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 12,
                "type": "固体物料",
                "substance": "45321",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 12,
                "type": "液体物料",
                "substance": "44421",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 12,
                "type": "溶剂",
                "substance": "43221",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 12,
                "type": "反溶剂",
                "substance": "43121",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 13,
                "type": "固体物料",
                "substance": "4533",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 13,
                "type": "液体物料",
                "substance": "4443",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 13,
                "type": "溶剂",
                "substance": "4323",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 13,
                "type": "反溶剂",
                "substance": "4313",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 14,
                "type": "固体物料",
                "substance": "4534",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 14,
                "type": "液体物料",
                "substance": "4444",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 14,
                "type": "溶剂",
                "substance": "4324",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 14,
                "type": "反溶剂",
                "substance": "4314",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 15,
                "type": "固体物料",
                "substance": "4535",
                "molecularFormula": "",
                "molecularWeight": "253",
                "moles": "1",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "253.00",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 15,
                "type": "液体物料",
                "substance": "4445",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "smile": "",
                "intensity": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            },
            {
                "id": 15,
                "type": "溶剂",
                "substance": "4325",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "2"
            },
            {
                "id": 15,
                "type": "反溶剂",
                "substance": "4315",
                "molecularFormula": "",
                "molecularWeight": "",
                "moles": "",
                "intensity": "",
                "smile": "",
                "cas": "",
                "quality": "",
                "Volume": "6.40",
                "moleConcentration": "",
                "singleCockAddVolume": "1"
            }
        ],
        "MaterialsTableList": [
            {
                "type": "固体物料",
                "substance": "111",
                "molecularFormula": "",
                "molecularWeight": "123",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "246.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "319.80",
                "configurationVolume": "0.00",
                "id": 1,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "112",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 2,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "113",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 3,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "114",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 4,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "121",
                "molecularFormula": "",
                "molecularWeight": "231",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "462.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "600.60",
                "configurationVolume": "0.00",
                "id": 5,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "122",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 6,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "123",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "4.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 7,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "124",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 8,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "125",
                "molecularFormula": "",
                "molecularWeight": "213",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "426.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "553.80",
                "configurationVolume": "0.00",
                "id": 9,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "127",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 10,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "131",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 11,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "132",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "3.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 12,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "133",
                "molecularFormula": "",
                "molecularWeight": "241",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "241.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "313.30",
                "configurationVolume": "0.00",
                "id": 13,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "134",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 14,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "143",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 15,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "154",
                "molecularFormula": "",
                "molecularWeight": "153",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "306.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "397.80",
                "configurationVolume": "0.00",
                "id": 16,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "211",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 17,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "212",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 18,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "213",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 19,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "233",
                "molecularFormula": "",
                "molecularWeight": "271",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "271.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "352.30",
                "configurationVolume": "0.00",
                "id": 20,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "234",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 21,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "235",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 22,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "243",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "4.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 23,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "241",
                "molecularFormula": "",
                "molecularWeight": "182",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "364.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "473.20",
                "configurationVolume": "0.00",
                "id": 24,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "242",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 25,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "244",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 26,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "311",
                "molecularFormula": "",
                "molecularWeight": "217",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "434.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "564.20",
                "configurationVolume": "0.00",
                "id": 27,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "321",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 28,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "323",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 29,
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "453",
                "molecularFormula": "",
                "molecularWeight": "253",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "253.00",
                "theoreticalVolume": "0.00",
                "configurationQuality": "328.90",
                "configurationVolume": "0.00",
                "id": 30,
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "444",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 31,
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "432",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "2.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 32,
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "431",
                "molecularFormula": "",
                "molecularWeight": "",
                "cas": "",
                "smile": "",
                "intensity": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.00",
                "configurationQuality": "0.00",
                "configurationVolume": "5.00",
                "id": 33,
                "concentration": ""
            }
        ],
        "isType": "2"
    },
    "experimentLog2": {
        "dissolveForm": {
            "temperature": "70",
            "maxTemperature": "110",
            "time": "60",
            "speed": "",
            "rjSpeed": "700",
            "seal": "1",
            "environment": "1",
            "dissolve": "1",
            "dissolveTimeInterval": "30",
            "dissolveMaximum": "4",
            "dissolveGradient": "5"
        },
        "postProcessingForm": {
            "time": "30",
            "temperature": "60",
            "maxTemperature": "",
            "speed": "800",
            "method": "2",
            "precipitate": "1",
            "precoolingTemperature": "40"
        },
        "filtrationForm": {
            "type": "硅胶",
            "time": "80"
        }
    },
    "experimentLog3": {
        "dissolveForm": {
            "temperature": "60",
            "maxTemperature": "110",
            "time": "70",
            "speed": "",
            "rjSpeed": "800",
            "seal": "1",
            "environment": "1",
            "dissolve": "1",
            "dissolveTimeInterval": 30,
            "dissolveMaximum": "2",
            "dissolveGradient": "5"
        },
        "postProcessingForm": {
            "time": "30",
            "maxTemperature": "",
            "preCooling": "1",
            "temperature": "70",
            "speed": "700",
            "precipitate": "1",
            "dissolveTimeInterval": "30",
            "dissolveMaximum": "3",
            "dissolveGradient": "5"
        },
        "filtrationForm": {
            "type": "硅胶",
            "time": "35"
        }
    }
}
    data7 = {"experimentLog1":{"wellPlates":"8孔板20ml","SubstratesTableList":[{"id":1,"type":"固体物料","substance":"Acetone","molecularFormula":"","molecularWeight":58.08,"moles":"0.8","smile":"CC(=O)C","intensity":0.791,"cas":"67-64-1","quality":"46.46","Volume":"","moleConcentration":"","singleCockAddVolume":""},{"id":1,"type":"液体物料","substance":"dichloromethane","molecularFormula":"","molecularWeight":84.93,"moles":"","smile":"C(Cl)Cl","intensity":1.325,"cas":"75-09-2","quality":"","Volume":"16.00","moleConcentration":"","singleCockAddVolume":"2"},{"id":1,"type":"溶剂","substance":"tetrahydrofuran","molecularFormula":"","molecularWeight":72.11,"moles":"","intensity":0.883,"smile":"C1CCOC1","cas":"109-99-9","quality":"","Volume":"16.00","moleConcentration":"","singleCockAddVolume":"2"},{"id":1,"type":"反溶剂","substance":"benzene","molecularFormula":"","molecularWeight":78.11,"moles":"","intensity":0.874,"smile":"C1=CC=CC=C1","cas":"71-43-2","quality":"","Volume":"16.00","moleConcentration":"","singleCockAddVolume":"2"}],"MaterialsTableList":[{"type":"固体物料","substance":"Acetone","molecularFormula":"","molecularWeight":58.08,"cas":"67-64-1","smile":"CC(=O)C","intensity":0.791,"theoreticalMoles":"0.80","theoreticalQuality":"46.46","theoreticalVolume":"0.00","configurationQuality":"60.40","configurationVolume":"0.00","id":1,"concentration":""},{"type":"液体物料","substance":"dichloromethane","molecularFormula":"","molecularWeight":84.93,"cas":"75-09-2","smile":"C(Cl)Cl","intensity":1.325,"theoreticalMoles":"0.00","theoreticalQuality":"0.00","theoreticalVolume":"2.00","configurationQuality":"0.00","configurationVolume":"5.00","id":2,"concentration":""},{"type":"溶剂","substance":"tetrahydrofuran","molecularFormula":"","molecularWeight":72.11,"cas":"109-99-9","smile":"C1CCOC1","intensity":0.883,"theoreticalMoles":"0.00","theoreticalQuality":"0.00","theoreticalVolume":"2.00","configurationQuality":"0.00","configurationVolume":"5.00","id":3,"concentration":""},{"type":"反溶剂","substance":"benzene","molecularFormula":"","molecularWeight":78.11,"cas":"71-43-2","smile":"C1=CC=CC=C1","intensity":0.874,"theoreticalMoles":"0.00","theoreticalQuality":"0.00","theoreticalVolume":"2.00","configurationQuality":"0.00","configurationVolume":"5.00","id":4,"concentration":""}],"isType":"1"},"experimentLog2":{"dissolveForm":{"temperature":"40","maxTemperature":"80","time":"1","speed":"","rjSpeed":"500","seal":"1","environment":"1","dissolve":"1","dissolveTimeInterval":"3","dissolveMaximum":"2","dissolveGradient":"5"},"postProcessingForm":{"time":"5","temperature":"20","maxTemperature":"","speed":"100","method":"1","precipitate":"1","precoolingTemperature":"20"},"filtrationForm":{"type":"硅胶","time":"1"}},"experimentLog3":{"dissolveForm":{},"postProcessingForm":{},"filtrationForm":{}}}
    data8 ={
    "experimentLog1": {
        "wellPlates": "24孔板8ml",
        "SubstratesTableList": [
            {
                "id": 1,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.43",
                "singleCockAddVolume": "2.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "reagent",
                "smiles": "",
                "substance": "lithium diisopropyl amide",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.00",
                "solvent": "52.4",
                "reactionMoles": "1.00",
                "reactionQuality": "52.40",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "solvent",
                "smiles": "",
                "substance": "tetrahydrofuran",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "4.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.43",
                "singleCockAddVolume": "2.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.60",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.77",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.80",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.77",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.80",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.77",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.80",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.77",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.80",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.77",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.80",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.77",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.80",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.77",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.80",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.77",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.80",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.77",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "reagent",
                "smiles": "",
                "substance": "dmap",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "reagent",
                "smiles": "",
                "substance": " triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "solvent",
                "smiles": "",
                "substance": "tetrahydrofuran",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.50",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.77",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.00",
                "solvent": "0.00",
                "reactionMoles": "1",
                "reactionQuality": "0.00",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "5.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.77",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.80",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.77",
                "singleCockAddVolume": "1.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "2.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.80",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.43",
                "singleCockAddVolume": "2.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "3.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.43",
                "singleCockAddVolume": "2.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.00",
                "singleCockAddVolume": "1.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "2.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.83",
                "singleCockAddVolume": "1.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.00",
                "solvent": "52.4",
                "reactionMoles": "1",
                "reactionQuality": "52.40",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "5.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.83",
                "singleCockAddVolume": "1.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "reagent",
                "smiles": "",
                "substance": "triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "",
                "solvent": "",
                "reactionMoles": "",
                "reactionQuality": "",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "5.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "sub",
                "smiles": "CC(=O)O",
                "substance": "CC(=O)O",
                "molecularFormula": "C2H4O2",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 60.05,
                "reactionMoles": 1,
                "reactionQuality": 60.05,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "0.83",
                "singleCockAddVolume": "1.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "reagent",
                "smiles": "",
                "substance": "triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.00",
                "solvent": "42.5",
                "reactionMoles": "1.00",
                "reactionQuality": "42.50",
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "solvent",
                "smiles": "",
                "substance": "chloroform",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 6.4,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "5.20",
                "cas": "",
                "intensity": None
            }
        ],
        "MaterialsTableList": [
            {
                "type": "sub",
                "substance": "CC(=O)O",
                "smiles": "CC(=O)O",
                "solvent": 60.05,
                "molecularFormula": "C2H4O2",
                "theoreticalMoles": "19.00",
                "theoreticalQuality": "1140.95",
                "theoreticalVolume": "0.00",
                "configurationMoles": "24.70",
                "cas": "",
                "intensity": None,
                "id": 1,
                "configurationVolume": "0.00",
                "configurationQuality": "1483.24",
                "concentration": ""
            },
            {
                "type": "sub",
                "substance": "Cc1csc(=S)s1",
                "smiles": "Cc1csc(=S)s1",
                "solvent": 148.28,
                "molecularFormula": "C4H4S3",
                "theoreticalMoles": "19.00",
                "theoreticalQuality": "2817.32",
                "theoreticalVolume": "28.40",
                "configurationMoles": "24.70",
                "cas": "",
                "intensity": None,
                "id": 2,
                "configurationVolume": "36.92",
                "configurationQuality": "3662.52",
                "concentration": "0.67"
            },
            {
                "type": "reagent",
                "substance": "lithium diisopropyl amide",
                "smiles": "",
                "solvent": "52.4",
                "molecularFormula": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "52.40",
                "theoreticalVolume": "0.00",
                "configurationMoles": "1.30",
                "cas": "",
                "intensity": None,
                "id": 3,
                "configurationVolume": "0.00",
                "configurationQuality": "68.12",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "tetrahydrofuran",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "6.60",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 4,
                "configurationVolume": "8.58",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "LiAlH4",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "15.90",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 5,
                "configurationVolume": "20.67",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "THF",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "57.70",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 6,
                "configurationVolume": "75.01",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "dmap",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.30",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 7,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": " triethylamine",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.30",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 8,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "Borane-THF",
                "smiles": "",
                "solvent": "0.00",
                "molecularFormula": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "52.40",
                "theoreticalVolume": "0.00",
                "configurationMoles": "2.60",
                "cas": "",
                "intensity": None,
                "id": 9,
                "configurationVolume": "0.00",
                "configurationQuality": "68.12",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "triethylamine",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "42.50",
                "theoreticalVolume": "0.00",
                "configurationMoles": "1.30",
                "cas": "",
                "intensity": None,
                "id": 10,
                "configurationVolume": "0.00",
                "configurationQuality": "55.25",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "dichloromethane",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "5.20",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 11,
                "configurationVolume": "6.76",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "chloroform",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "5.20",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 12,
                "configurationVolume": "6.76",
                "configurationQuality": "0.00",
                "concentration": ""
            }
        ]
    },
    "experimentLog2": {
        "ReactionConditions": {
            "temperature": "40",
            "time": "60",
            "speed": "700",
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
            ],
            "show": False
        },
        "electricityData": {
            "type": "1",
            "list": [],
            "show": False
        },
        "gasData": {
            "name": "",
            "pressure": "",
            "time": "",
            "show": False
        }
    },
    "experimentLog3": {
        "coolingTime": "60",
        "quenchingAgentData": [],
        "internalStandardData": [],
        "productDilutionData": [],
        "extractionData": [
            {
                "name": "toluene",
                "volume": "100",
                "time": "50",
                "times": "2",
                "type": "2",
                "density": "0.866"
            }
        ],
        "filtrationData": [
            {
                "name": "硅胶",
                "filterTime": "60"
            }
        ]
    },
    "experimentLog4": {
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
        ]
    }
}
    data9 = {
    "experimentLog1": {
        "wellPlates": "24孔板8ml",
        "SubstratesTableList": [
            {
                "id": 1,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            }
        ],
        "MaterialsTableList": [
            {
                "type": "sub",
                "substance": "X03",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "solvent": "122.10",
                "molecularFormula": "C24H20BNa",
                "theoreticalMoles": "2.40",
                "theoreticalQuality": "293.04",
                "theoreticalVolume": "0.00",
                "configurationMoles": "3.12",
                "cas": "",
                "intensity": None,
                "id": 1,
                "configurationVolume": "0.00",
                "configurationQuality": "380.95",
                "concentration": ""
            },
            {
                "type": "sub",
                "substance": "K2CO3",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "solvent": "138.2",
                "molecularFormula": "C12H10ClI",
                "theoreticalMoles": "3.60",
                "theoreticalQuality": "497.52",
                "theoreticalVolume": "0.00",
                "configurationMoles": "4.69",
                "cas": "",
                "intensity": None,
                "id": 2,
                "configurationVolume": "0.00",
                "configurationQuality": "646.78",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "A04",
                "smiles": "",
                "solvent": "109.13",
                "molecularFormula": "",
                "theoreticalMoles": "3.60",
                "theoreticalQuality": "392.88",
                "theoreticalVolume": "2.40",
                "configurationMoles": "4.69",
                "cas": "",
                "intensity": None,
                "id": 3,
                "configurationVolume": "5.00",
                "configurationQuality": "510.74",
                "concentration": "1.50"
            },
            {
                "type": "solvent",
                "substance": "DMF",
                "smiles": "",
                "solvent": None,
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "45.60",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 4,
                "configurationVolume": "59.28",
                "configurationQuality": "0.00",
                "concentration": ""
            }
        ]
    },
    "experimentLog2": {
        "ReactionConditions": {
            "temperature": "100",
            "time": "720",
            "speed": "720",
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
            ],
            "show": True
        },
        "electricityData": {
            "type": "1",
            "list": [],
            "show": True
        },
        "gasData": {
            "name": "",
            "pressure": "",
            "time": "",
            "show": True
        }
    },
    "experimentLog3": {
        "coolingTime": "60",
        "quenchingAgentData": [],
        "internalStandardData": [
            {
                "name": "93-98-1/644-08-6",
                "moles": "",
                "molecularWeight": "197.23",
                "addVolume": "1000",
                "internalConcentration": "",
                "internalVolume": "",
                "internalQuality": ""
            }
        ],
        "productDilutionData": [
            {
                "name": "MeCN",
                "beforeConcentration": "",
                "afterConcentration": "",
                "procedure": {
                    "diluent1": "1000",
                    "diluent2": "500",
                    "diluent3": "100",
                    "diluent4": "100",
                    "diluent5": "100"
                }
            }
        ],
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
        ]
    }
}
    data10 ={
    "experimentLog1": {
        "wellPlates": "24孔板8ml",
        "SubstratesTableList": [
            {
                "id": 1,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "sub",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "substance": "X03",
                "molecularFormula": "C24H20BNa",
                "limitReagents": True,
                "equivalent": "1.00",
                "solvent": "122.10",
                "reactionMoles": "0.1",
                "reactionQuality": "12.21",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "sub",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "substance": "K2CO3",
                "molecularFormula": "C12H10ClI",
                "limitReagents": False,
                "equivalent": "1.50",
                "solvent": "138.2",
                "reactionMoles": "0.15",
                "reactionQuality": "20.73",
                "reactionVolume": "2",
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "reagent",
                "smiles": "",
                "substance": "A04",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.50",
                "solvent": "109.13",
                "reactionMoles": "0.15",
                "reactionQuality": "16.37",
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.50",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "solvent",
                "smiles": "",
                "substance": "DMF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": None,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": "2",
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "1.90",
                "cas": "",
                "intensity": None
            }
        ],
        "MaterialsTableList": [
            {
                "type": "sub",
                "substance": "X03",
                "smiles": "[Na+].c1ccc([B-](c2ccccc2)(c2ccccc2)c2ccccc2)cc1",
                "solvent": "122.10",
                "molecularFormula": "C24H20BNa",
                "theoreticalMoles": "2.40",
                "theoreticalQuality": "293.04",
                "theoreticalVolume": "0.00",
                "configurationMoles": "3.12",
                "cas": "",
                "intensity": None,
                "id": 1,
                "configurationVolume": "0.00",
                "configurationQuality": "380.95",
                "concentration": ""
            },
            {
                "type": "sub",
                "substance": "K2CO3",
                "smiles": "[Cl-].c1ccc([I+]c2ccccc2)cc1",
                "solvent": "138.2",
                "molecularFormula": "C12H10ClI",
                "theoreticalMoles": "3.60",
                "theoreticalQuality": "497.52",
                "theoreticalVolume": "0.00",
                "configurationMoles": "4.69",
                "cas": "",
                "intensity": None,
                "id": 2,
                "configurationVolume": "0.00",
                "configurationQuality": "646.78",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "A04",
                "smiles": "",
                "solvent": "109.13",
                "molecularFormula": "",
                "theoreticalMoles": "3.60",
                "theoreticalQuality": "392.88",
                "theoreticalVolume": "2.40",
                "configurationMoles": "4.69",
                "cas": "",
                "intensity": None,
                "id": 3,
                "configurationVolume": "5.00",
                "configurationQuality": "510.74",
                "concentration": "1.50"
            },
            {
                "type": "solvent",
                "substance": "DMF",
                "smiles": "",
                "solvent": None,
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "45.60",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 4,
                "configurationVolume": "59.28",
                "configurationQuality": "0.00",
                "concentration": ""
            }
        ]
    },
    "experimentLog2": {
        "ReactionConditions": {
            "temperature": "100",
            "time": "720",
            "speed": "720",
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
            ],
            "show": True
        },
        "electricityData": {
            "type": "1",
            "list": [],
            "show": True
        },
        "gasData": {
            "name": "",
            "pressure": "",
            "time": "",
            "show": True
        }
    },
    "experimentLog3": {
        "coolingTime": "60",
        "quenchingAgentData": [
            {
                "name": "31231231231",
                "moles": "",
                "molecularWeight": "",
                "addVolume": "100",
                "quencherConcentration": "",
                "quenchVolume": "",
                "quencherQuality": ""
            }
        ],
        "internalStandardData": [
            {
                "name": "93-98-1/644-08-6",
                "moles": "",
                "molecularWeight": "197.23",
                "addVolume": "200",
                "internalConcentration": "",
                "internalVolume": "",
                "internalQuality": ""
            }
        ],
        "productDilutionData": [
            {
                "name": "MeCN",
                "beforeConcentration": "",
                "afterConcentration": "",
                "procedure": {
                    "diluent1": "1000",
                    "diluent2": "500",
                    "diluent3": "100",
                    "diluent4": "100",
                    "diluent5": "100"
                }
            }
        ],
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
        ]
    }
} 
    main('1',data10,1)
    # main('1',data4,2)



'''合成平台:
            1.24
'''