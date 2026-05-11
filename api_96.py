import copy
import json
import zipfile
from decimal import Decimal

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

BASE_PATH = r"./generation_zy"
BASE_PATH1 = r"./逆合成"
crystallize_BASE_PATH = r"C:/Users/24017/PycharmProjects/pythonProject/接口测试/表格生成/结晶平台"
crystallize_BASE_PATH1 = r"C:/Users/24017/PycharmProjects/pythonProject/接口测试/模板/结晶平台"

TEMPLATE_PATHS = {
    "wf11": f"{BASE_PATH}/hiwo-任务2-1.xlsx",
    "wf05": f"{BASE_PATH}/hiwo-任务1-1.xlsx",
    "sowo": f"{BASE_PATH}/sowo-物料1-1.xlsx",
    "mate": f"{BASE_PATH}/sowo-任务1-1.xlsx",
    "information": f"{BASE_PATH}/信息表格.xlsx",
    "mate_info1": f"{BASE_PATH}/hiwo-物料1-1.xlsx",
    "mate_info2": f"{BASE_PATH}/hiwo-物料2-1.xlsx"

}
crystallize_TEMPLATE_PATHS = {
    "wf11": f"{crystallize_BASE_PATH}/hiwo-任务1-1.xlsx",
    "wf05": f"{crystallize_BASE_PATH}/hiwo-任务2-1.xlsx",
    "sowo": f"{crystallize_BASE_PATH}/sowo-物料1-1.xlsx",
    "mate": f"{crystallize_BASE_PATH}/sowo-任务1-1.xlsx",
    "mate_info1": f"{crystallize_BASE_PATH}/hiwo-物料1-1.xlsx",
    "mate_info2": f"{crystallize_BASE_PATH}/hiwo-物料1-2.xlsx",
    "information":f"{crystallize_BASE_PATH}/信息表格.xlsx",
    "WF10":f"{crystallize_BASE_PATH1}/hiwo任务15-48-96.xlsx",
    "WF11": f"{crystallize_BASE_PATH1}/hiwo物料-96-01.xlsx"
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
        self.reagent_plates_96 = []
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
        #枪头
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
                #添加hiwo物料到信息表格
                self._create_excel_xxx(TEMPLATE_PATHS["mate_info1"], TEMPLATE_PATHS["information"], 1)
                # 内标等试剂
                self._create_excel_xxx(TEMPLATE_PATHS["mate_info2"], TEMPLATE_PATHS["information"], 100)

                #sowo物料
                self._create_excel_ttt(200)
                self._create_excel_information(200, 7, TEMPLATE_PATHS["information"])

            else:
                self._create_excel_xxx(TEMPLATE_PATHS["mate_info1"], TEMPLATE_PATHS["information"], 1)
                #第二次任务
                self._create_excel_xxx(TEMPLATE_PATHS["mate_info1"].replace('1-1','1-2'), TEMPLATE_PATHS["information"], 100)
                # 内标等试剂
                self._create_excel_xxx(TEMPLATE_PATHS["mate_info2"], TEMPLATE_PATHS["information"], 200)
                # sowo物料
                self._create_excel_ttt(300)
                self._create_excel_information(300, 7, TEMPLATE_PATHS["information"])

            self._create_excel_material(TEMPLATE_PATHS["information"])
        elif self.type == 2:

            self._create_excel_xxx(crystallize_TEMPLATE_PATHS["mate_info1"], crystallize_TEMPLATE_PATHS["information"],1)
            self._create_excel_tttt(100, crystallize_TEMPLATE_PATHS['sowo'], crystallize_TEMPLATE_PATHS['information'])
            self._create_excel_information(100, 7, crystallize_TEMPLATE_PATHS["information"])
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

        for col_idx, letter in enumerate("ABCDEFGH", start=start_col):
            ws.cell(row=start_row, column=col_idx, value=letter).font = Font(bold=True)

        # 添加行标签 (1,2,3...)
        for row_idx in range(start_row + 1, start_row + 13):
            ws.cell(row=row_idx, column=start_col - 1, value=(start_row + 13) - row_idx).font = Font(bold=True)
        thin_border = Border(left=Side(style='thin'),
                             right=Side(style='thin'),
                             top=Side(style='thin'),
                             bottom=Side(style='thin'))
        ExcelUtils.save_workbook(wb, input_path)
        # 写入数据
        for row_idx, row_data in enumerate(self.perforated_plate, start=1):
            for col_idx, cell_value in enumerate(row_data, start=start_col):
                cell = ws.cell(row=start_row + 13 - row_idx, column=col_idx, value=cell_value)
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

    def _extract_materials(self) -> None:
        """提取物料数据"""
        materials = self.post_data["experimentLog1"]["MaterialsTableList"]
        for item in materials:
            substance = item["substance"].strip()
            self.theoretical_volume[substance] = item["theoreticalVolume"]

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

    def assign_to_plate(self,reaction_groups, rows=12, cols=8):
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
        """计算反应相似度并分组"""
        #按相似度排序
        # reaction_groups = self.process_reaction_conditions(self.data_dict)
        # plate = self.assign_to_plate(reaction_groups, rows=12, cols=8)
        # plate = self.compress_plate_layout(plate)
        # for plate in plate:
        #     self.perforated_plate.append(plate)

        #准备数据集按长度排序
        # sorted_keys = sorted(self.data_dict.keys(),
        #                      key=lambda x: len(self.data_dict[x]),
        #                      reverse=True)
        # #列表推导式
        # # self.perforated_plate = [sorted_keys[i:i + 8] for i in range(0, len(sorted_keys), 8)]
        # print(self.data_dict)
        temp_group = []
        for list_id in  self.data_dict.keys():
            temp_group.append(list_id)
            if len(temp_group)==8:
                self.perforated_plate.append(temp_group)
                temp_group = []
        if temp_group:
            self.perforated_plate.append(temp_group)

        #生成孔位设置json文件
        self._hole_json()
        for perforated in copy.deepcopy(self.perforated_plate):
            group_size = len(perforated)  # 获取当前组的大小

            # 获取当前组所有反应的化学列表
            chem_lists = [self.data_dict[f"{pid}"] for pid in perforated]
            max_length = max(len(chem_list) for chem_list in chem_lists)
            llwc_num = 1
            all_reagents = {}
            for chem_list in chem_lists:
                for reagent_dict in chem_list:
                    for reagent, amount in reagent_dict.items():
                        if reagent in all_reagents:
                            all_reagents[reagent] += 1
                        else:
                            all_reagents[reagent] = 1

            # 找出在所有反应中都出现的试剂
            common_reagents = {reagent for reagent, count in all_reagents.items() if count == len(chem_lists)}
            chem_lists_12 = []  # 用于12孔板的相同试剂
            chem_lists_96 = []  # 用于96孔板的不同试剂

            for chem_list in chem_lists:
                common_reagent_list = []
                unique_reagent_list = []

                for reagent_dict in chem_list:
                    reagent = list(reagent_dict.keys())[0]
                    if reagent in common_reagents:
                        common_reagent_list.append(reagent_dict)
                    else:
                        unique_reagent_list.append(reagent_dict)

                chem_lists_12.append(common_reagent_list)
                chem_lists_96.append(unique_reagent_list)
            #12孔板分配{'CCC(C)=O': ['3'], 'LiAlH4': ['3'], 'Nc1ccccc1': ['3']}
            if chem_lists_12 and any(chem_lists_12):  # 如果有相同试剂
                max_common_length = max(len(chem_list) for chem_list in chem_lists_12)
                for  chem_lists_12_  in chem_lists_12[0]:
                    key = list(chem_lists_12_.keys())[0]
                    if key in list(self.reagent_plates_12_.keys()):

                        self.reagent_plates_12_[f'{key}'].append(perforated[0])
                    else:
                        self.reagent_plates_12_.update({key: [perforated[0]]})


            #96孔板分配
            max_unique_length = max(len(chem_list) for chem_list in chem_lists_96)
            llwc_num = 1
            for i in range(max_unique_length):
                row_items = []
                for j in range(group_size):
                    chem_list = chem_lists_96[j]
                    item = chem_list[i] if i < len(chem_list) else None
                    row_items.append(item)

                # 补齐到8个位置（96孔板每行8个位置）
                while len(row_items) < 8:
                    row_items.append(None)

                # 添加组ID和位置编号到96孔板[[{'dichloromethane': 7.0}, {'chloroform': 7.0}, '5', 1]]
                self.reagent_plates_96.append(row_items + [perforated.copy(), llwc_num])
                llwc_num += 8  # 96孔板每行8个位置
        pass



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
            transfer_records_12, transfer_records_24, transfer_records_2 = self._create_w11_task()
            # 添加任务配置
            self._create_task_configuration()
            # 添加稀释数据
            self._create_dilution_data()
            # 添加萃取
            self._create_task_extract()
            # 添加sowo物料数据
            excel_list,excel_list_2 =  self._create_sowo_edit_stuff()
            # 添加sowo任务数据
            self._create_mate_msg()
            self.modify_stack_table(transfer_records_12, transfer_records_24, transfer_records_2, BASE_PATH1, BASE_PATH)
            self.write_codes_to_inner_hopper_stack(fr"{BASE_PATH}/Materials-stack4列.xlsx",
                                               fr"{BASE_PATH}/Materials-stack4列.xlsx", excel_list, excel_list_2)
        elif self.type == 2:
            # 重新写 函数 生成出表格
            self.data_list_24, self.data_list_12 = self._create_crystallize_mate_info()
            # 添加任务表信息
            transfer_records_12, transfer_records_24, transfer_records_2 = self._create_crystallize_wf05_task()

            # 生成温度梯度表格
            self._create_crystallize_task_configuration()

            # 添加sowo物料数据
            excel_list,excel_list_2 =  self._create_sowo_crystallize_edit_stuff()
            # 添加sowo任务数据
            self._create_mate_msg()
            self.modify_stack_table(transfer_records_12, transfer_records_24, transfer_records_2,
                                    crystallize_BASE_PATH1, crystallize_BASE_PATH)
            self.write_codes_to_inner_hopper_stack(fr"{crystallize_BASE_PATH}/Materials-stack4列.xlsx",fr"{crystallize_BASE_PATH}/Materials-stack4列.xlsx",excel_list,excel_list_2)


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
            material_data = df_material.iloc[0:5, [1, 7]].reset_index(drop=True)

            # 获取孔位表第一列值为1和2的数据(1-6列)
            hole_data_1 = df_holes[df_holes.iloc[:, 0] == 1].iloc[:, [0,1,3,5]]
            hole_data_2 = df_holes[df_holes.iloc[:, 0] == 2].iloc[:, [0,1,3,5]]
            hole_data_3 = df_holes[df_holes.iloc[:, 0] == 3].iloc[:, [0,1,3,5]]
            hole_data_4 = df_holes[df_holes.iloc[:, 0] == 4].iloc[:, [0,1,3,5]]
            # 写入任务标签
            if rows == 1:
                # 合并单元格并写入"任务一"
                ws.merge_cells(start_row=rows, end_row=96, start_column=1, end_column=1)
                ws.cell(row=rows, column=1, value="任务一")
                alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(row=rows, column=1).alignment = alignment

            else:
                # 对于第二组数据，调整行号
                # 合并单元格并写入"任务二"
                ws.merge_cells(start_row=rows, end_row=rows + 96, start_column=1, end_column=1)
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

            # 写入右侧数据(第3行物料+孔位2)(表头)
            right_start_col = len(material_data.columns) + 7
            # 物料第3行数据
            for col_idx, value in enumerate(material_data.iloc[1], start=right_start_col):
                ws.cell(row=rows, column=col_idx, value=value)

            if rows == 1:
                if self.number_tasks == 1:
                    row_ = -10
                else:
                    row_ = -22
                # 孔位2数据
                for row_idx, row in hole_data_2.iterrows():
                    for col_idx, value in enumerate(row, start=7):
                        ws.cell(row=row_idx + row_, column=col_idx, value=value)
                for col_idx, value in enumerate(material_data.iloc[2], start=12):
                    ws.cell(row=rows, column=col_idx, value=value)
                for row_idx, row in hole_data_3.iterrows():
                    for col_idx, value in enumerate(row, start=12):
                        ws.cell(row=row_idx + row_ - 96, column=col_idx, value=value)
                for col_idx, value in enumerate(material_data.iloc[3], start=17):
                    ws.cell(row=rows, column=col_idx, value=value)
                for row_idx, row in hole_data_4.iterrows():
                    for col_idx, value in enumerate(row, start=17):
                        ws.cell(row=row_idx + row_ - 192, column=col_idx, value=value)
                # 保存结果
                wb.save(output_path)
            else:
                # 孔位2数据
                for row_idx, row in hole_data_2.iterrows():
                    for col_idx, value in enumerate(row, start=7):
                        ws.cell(row=row_idx + 89, column=col_idx, value=value)

                for col_idx, value in enumerate(material_data.iloc[2], start=12):
                    ws.cell(row=rows, column=col_idx, value=value)
                for row_idx, row in hole_data_3.iterrows():
                    for col_idx, value in enumerate(row, start=12):
                        ws.cell(row=row_idx + 81, column=col_idx, value=value)

                for col_idx, value in enumerate(material_data.iloc[3], start=17):
                    ws.cell(row=rows, column=col_idx, value=value)
                for row_idx, row in hole_data_4.iterrows():
                    for col_idx, value in enumerate(row, start=17):
                        ws.cell(row=row_idx -15, column=col_idx, value=value)

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
                ws.merge_cells(start_row=rows, end_row=96, start_column=1, end_column=1)
                ws.cell(row=rows, column=1, value="任务一")
                alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(row=rows, column=1).alignment = alignment

            else:
                # 对于第二组数据，调整行号
                # 合并单元格并写入"任务一"
                ws.merge_cells(start_row=rows, end_row=rows + 96, start_column=1, end_column=1)
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
            wb.save(output_path)
            if rows == 1:
                if self.number_tasks == 1:
                    row_ = -10
                else:
                    row_ = -10
            else:
                row_ = rows - 11
            # 孔位2数据
            for row_idx, row in hole_data_2.iterrows():
                for col_idx, value in enumerate(row, start=7):
                    ws.cell(row=row_idx + row_, column=col_idx, value=value)
            #添加表头
            for col_idx, value in enumerate(material_data.iloc[2], start=12):
                ws.cell(row=rows, column=col_idx, value=value)
            # 添加孔位数据
            for row_idx, row in hole_data_3.iterrows():
                for col_idx, value in enumerate(row, start=12):
                    ws.cell(row=row_idx + rows - 107, column=col_idx, value=value)
            # 添加表头

            # 添加孔位数据
            for row_idx, row in hole_data_4.iterrows():
                for col_idx, value in enumerate(row, start=17):
                    ws.cell(row=row_idx + rows - 203, column=col_idx, value=value)

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
            ws.merge_cells(start_row=rows + 120, end_row=140, start_column=6, end_column=6)
            ws.cell(row=rows + 120, column=6, value="反溶剂添加")
            alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=rows + 120, column=6).alignment = alignment

        # 写入左侧数据(第2行物料+孔位1)
        # 物料第2行数据
        for col_idx, value in enumerate(material_data.iloc[0], start=2):
            ws.cell(row=rows + 120, column=col_idx+5, value=value)

        # 孔位1数据
        for row_idx, row in hole_data_1.iterrows():
            for col_idx, value in enumerate(row, start=2):
                ws.cell(row=row_idx + 120, column=col_idx+5, value=value)


        # 保存结果
        wb.save(output_path)
        print(f"处理完成，结果已保存到: {output_path}")

    def _create_excel_ttt(self, rowss):
        def get_row_values(sheet, row_num, prefix=""):
            data_list = [cell.value for cell in sheet[row_num]]
            return {
                '试剂板类型': data_list[1],
                '位置编号': data_list[7],
                '信息': f"{prefix}-{'左' if row_num == 2 else '右'}"
            }

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
        """逆合成创建物料信息Excel"""
        # 初始化数据列表
        data_list_12 = self._process_12_plate()
        data_list_24, result_1 = self._process_24_plate()
        # 文件路径
        ExcelUtils.copy_workbook(
            f"{BASE_PATH1}/hiwo物料-96-01.xlsx",
            TEMPLATE_PATHS["mate_info1"]
        )
        '''
        这里不做长度判断，如果96孔试剂版够使用 max_num会为0，否则会把剩余的96孔试剂版添加到第二个物料表
        '''
        wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info1"])
        max_num = self._fill_worksheet(wb, data_list_12, data_list_24)
        ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info1"])
        if max_num < 36:
            ExcelUtils.copy_workbook(
                f"{BASE_PATH1}/hiwo物料-96-02.xlsx",
                TEMPLATE_PATHS["mate_info2"])
            # 生成第二个物料表格
            wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info2"])
            self._create_extended_version(wb, [], data_list_24[max_num:], max_num)
            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info2"])
        else:
            self.number_tasks = 1
        # 生成第一个物料表格
        #判断错误：如果12孔板超出需要占用，而96孔板剩余的不够使用，会有错误
        # if len(data_list_24) > 36:
        #     '''
        #     先写入12孔数据，如果12孔数据超出写入第一个96孔板
        #     根据占用的孔位再写入96孔数据
        #     如果没超出，直接写入
        #     3块96板试剂版总共3*12=36行
        #
        #     使用：1.先写入12孔板，和96孔板
        #         2.做判断，12孔板是否不够用，96孔板是否不够用，是否需要第二个物料表
        #         3.需要：超出的12孔板以及96 统统写到第二个
        #         4.不需要：12孔多出x个，96孔剩余x个，把12孔多出数据填充到96孔
        #         会不会任务表不好找12孔数据
        #     '''
        #     wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info1"])
        #     max_num = self._fill_worksheet(wb, data_list_12, data_list_24)
        #     ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info1"])
        #     if max_num>0:
        #         # 生成第二个物料表格
        #         wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info1"])
        #
        #         self._create_extended_version(wb, [], data_list_24[max_num:],max_num)
        #         ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info1"].replace('1-1','1-2'))
        #     else:
        #         self.number_tasks = 1
        # else:
        #
        #         self.number_tasks = 1
        #         wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info1"])
        #
        #         max_num = self._fill_worksheet(wb, data_list_12, data_list_24)
        #         ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["mate_info1"])
        #         """如果24孔使用数量小于6，那么只需要一个任务就可以完成"""

        return result_1, data_list_12

    def _create_crystallize_mate_info(self) -> Tuple[list, list]:
        """合成平台创建物料信息Excel"""
        # 初始化数据列表

        data_list_12 = self._process_12_plate()
        data_list_24, result_1 = self._process_24_plate()
        ExcelUtils.copy_workbook(
            f"{crystallize_BASE_PATH1}/hiwo物料-96-01.xlsx",
            crystallize_TEMPLATE_PATHS["mate_info1"]
        )

        wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["mate_info1"])
        max_num = self._fill_worksheet(wb, data_list_12, data_list_24)
        ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["mate_info1"])
        if max_num < 36:

            # 生成第二个物料表格
            ExcelUtils.copy_workbook(
                f"{crystallize_BASE_PATH1}/hiwo物料-96-02.xlsx",
                crystallize_TEMPLATE_PATHS["mate_info2"])
            wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["mate_info2"])
            self._create_extended_version(wb, [], data_list_24[max_num:], max_num)
            ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["mate_info2"])
        else:
            self.number_tasks = 1
        # 生成第一个物料表格 96：8*12
        # if len(data_list_24) > 36 :
        #
        #     wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["mate_info1"])
        #     max_num = self._fill_worksheet(wb, data_list_12, data_list_24)
        #     ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["mate_info1"])
        #     if max_num > 0:
        #
        #         # 生成第二个物料表格
        #         ExcelUtils.copy_workbook(
        #             f"{crystallize_BASE_PATH1}/物料-96-02.xlsx",
        #             crystallize_TEMPLATE_PATHS["mate_info2"])
        #         wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["mate_info2"])
        #         self._create_extended_version(wb, [], data_list_24[max_num:],max_num)
        #         ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["mate_info2"])
        #     else:
        #         self.number_tasks = 1
        # else:
        #     self.number_tasks = 1
        #     wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["mate_info1"])
        #     # 24孔板
        #     max_num = self._fill_worksheet(wb, data_list_12, data_list_24)
        #     ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["mate_info1"])
        #     """如果24孔使用数量小于6，那么只需要一个任务就可以完成"""

        return result_1, data_list_12
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
        # for item in excel_list:
        #     locate = item.get('Locate')
        #     code = item.get('Code')
        #     if locate is None or code is None:
        #         continue  # 跳过无效项
        #     if not (1 <= locate <= 24):
        #         print(f"警告：Locate={locate} 超出 B2:G5 范围，已跳过")
        #         continue
        #
        #     # 计算行、列
        #     # 每行6个位置（B~G），列偏移：0->B(2),1->C(3),...,5->G(7)
        #     row_offset = (locate - 1) // 5  # 0,1,2,3 对应第2,3,4,5行
        #     col_offset = (locate - 1) % 5  # 0~5 对应列号2~7
        #
        #     row = 2 + row_offset
        #     col = 2 + col_offset
        #     print(f"{row}:{col}")
        #     cell = ws.cell(row=row, column=col)
        #     cell.value = code
        #     cell.fill = yellow_fill
        #
        # wb.save(output_path)
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

    def _create_crystallize_wf05_task(self) :
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

        ExcelUtils.copy_workbook(
            f"{crystallize_BASE_PATH1}/hiwo任务05-96-01-02.xlsx",
            crystallize_TEMPLATE_PATHS["wf11"]
        )
        wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["wf11"])
        ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_12)
        ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_24)
        ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["wf11"])
        if transfer_records_2:
            ExcelUtils.copy_workbook(
                f"{crystallize_BASE_PATH1}/hiwo任务05-96-01-02.xlsx",
                crystallize_TEMPLATE_PATHS["wf05"]
            )
            wb = ExcelUtils.open_workbook(crystallize_TEMPLATE_PATHS["wf05"])
            ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_2)
        ExcelUtils.save_workbook(wb, crystallize_TEMPLATE_PATHS["wf05"])
        return transfer_records_12, transfer_records_24, transfer_records_2

    def modify_stack_table(self,transfer_records_12, transfer_records_24, transfer_records_2,BASE_PATH_1,BASE_PATH_2):
        fill_cells = []
        if len(transfer_records_12) > 0:
            fill_cells.append('C2')
        len_24 = len(transfer_records_24)
        CELLS_ROW4 = ["C3", "C4", "C5"]  # 按顺序，对应每12长度一个单元格

        if len_24 > 0:
            n = (len_24 - 1) // 12 + 1  # 向上取整，得到需要填充的单元格数量（1~3）
            n = min(n, len(CELLS_ROW4))
            fill_cells.extend(CELLS_ROW4[:n])

        len_2 = len(transfer_records_2)
        CELLS_ROW4 = ["C6", "C7", "D2","D3"]  # 按顺序，对应每12长度一个单元格

        if len_2 > 0:
            n = (len_2 - 1) // 12 + 1  # 向上取整，得到需要填充的单元格数量（1~3）
            n = min(n, len(CELLS_ROW4))
            fill_cells.extend(CELLS_ROW4[:n])

        quenchingAgentData = self.post_data['experimentLog3']["quenchingAgentData"]
        internalStandardData = self.post_data['experimentLog3']["internalStandardData"]
        if quenchingAgentData and internalStandardData:
            if 'C6' in fill_cells:
                fill_cells.append("D4")
            else:
                fill_cells.append("C6")

        if "A05" in self.gun_head_dict:
            fill_cells.append("E4")
        if "B05" in self.gun_head_dict:
            fill_cells.append("E5")
        if "A04" in self.gun_head_dict:
            fill_cells.append("E2")
        if "B04" in self.gun_head_dict:
            fill_cells.append("E3")

        self.process_plate_stack(
        file_path=fr"{BASE_PATH_1}/Materials-stack4列.xlsx",
        output_path=fr"{BASE_PATH_2}/Materials-stack4列.xlsx",
        fill_cells = fill_cells,
        plate_type='96孔',
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
        wb = ExcelUtils.open_workbook(excel_name)
        if type == 1:
            filtrationForm_time = self.post_data.get('experimentLog3', []).get('filtrationForm', []).get('time', 0)
        else:
            filtrationForm_time = self.post_data.get('experimentLog2', []).get('filtrationForm', []).get('time', 0)
        ExcelUtils.modify_cell(wb, "任务参数配置", "M3", int(filtrationForm_time))
        for s, i in enumerate(result, 1):
            target_volume = i * 1000
            if int(target_volume) <= 200:
                pipette_location = 'A04'
            else:
                pipette_location = 'A05'
            if len(self.gun_head_dict[pipette_location]) >= 12:
                if pipette_location == 'A04': pipette_location = 'B04'
                if pipette_location == 'A05': pipette_location = 'B05'

            gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0

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
                f"{crystallize_BASE_PATH1}/hiwo任务01-48-96.xlsx",
                excel_name
            )
            ExcelUtils.copy_workbook(
                f"{crystallize_BASE_PATH1}/hiwo物料-96-02.xlsx",
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
                        f"{crystallize_BASE_PATH1}/hiwo任务01-48-96.xlsx",
                        excel_name
                    )
                    ExcelUtils.copy_workbook(
                        f"{crystallize_BASE_PATH1}/hiwo物料-96-02.xlsx",
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
                f"{crystallize_BASE_PATH1}/hiwo任务01-48-96.xlsx",
                excel_name)
            ExcelUtils.copy_workbook(
                f"{crystallize_BASE_PATH1}/hiwo物料-96-02.xlsx",
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
                        f"{crystallize_BASE_PATH1}/hiwo任务01-48-96.xlsx", excel_name)
                    ExcelUtils.copy_workbook(
                        f"{crystallize_BASE_PATH1}/hiwo物料-96-02.xlsx",
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
                    f"{crystallize_BASE_PATH1}/hiwo任务01-48-96.xlsx",
                    excel_name
                )
                ExcelUtils.copy_workbook(
                    f"{crystallize_BASE_PATH1}/hiwo物料-96-02.xlsx",
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
                            f"{crystallize_BASE_PATH1}/hiwo任务01-48-96.xlsx",
                            excel_name
                        )
                        ExcelUtils.copy_workbook(
                            f"{crystallize_BASE_PATH1}/hiwo物料-96-02.xlsx",
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
                f"{crystallize_BASE_PATH1}/hiwo物料-96-02.xlsx",
                excel_name
            )
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
                f"{crystallize_BASE_PATH1}/hiwo任务01-48-96.xlsx",
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
                data_key = [aa for aa in all_solid_records if aa['hole_id'] == int(list_data[0])][0]
                hole_id = data_key['hole_id']
                amount = data_key['weight']

                target_volume = max_key * 1000  if max_key else 0
                target_volumes = amount * 1000 if amount else 0
                bar_code = "SS-B02"

                if int(target_volumes) <= 200:
                    pipette_location = 'A04'
                else:
                    pipette_location = 'A05'
                if len(self.gun_head_dict[pipette_location]) >= 12:
                    if pipette_location == 'A04': pipette_location = 'B04'
                    if pipette_location == 'A05': pipette_location = 'B05'

                gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0

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
                f"{crystallize_BASE_PATH1}/hiwo物料-96-02.xlsx",
                excel_name
            )
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
                f"{crystallize_BASE_PATH1}/hiwo任务01-48-96.xlsx",
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

            for llwc, list_data in enumerate(self.perforated_plate, 1):
                target_volume = int(key_dict.get(int(list_data[0]), []).get("weight", ''))  * 1000
                bar_code = "SS-A02"
                if int(target_volume) <= 200:
                    pipette_location = 'A04'
                else:
                    pipette_location = 'A05'
                if len(self.gun_head_dict[pipette_location]) >= 12:
                    if pipette_location == 'A04': pipette_location = 'B04'
                    if pipette_location == 'A05': pipette_location = 'B05'

                gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0

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

    def _create_sowo_crystallize_edit_stuff(self) :
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

        return excel_list,excel_list_2
        #表三
        # self.write_codes_to_inner_hopper_stack(fr"{crystallize_BASE_PATH}/Materials-stack4列.xlsx",fr"{crystallize_BASE_PATH}/Materials-stack4列.xlsx",excel_list,excel_list_2)

    def _process_12_plate(self) -> List[str]:
        """处理12孔板数据"""
        result = []
        result_1 = []
        for substance in self.reagent_plates_12_.keys():
            # list_ = self.reagent_plates_12.get(substance,[])
            hole_data = {}
            for data in self.perforated_plate:
                hole_data.update({data[0]: len(data)})
            vol = 0
            for i in self.reagent_plates_12_.get(substance,[]):
                hole_num = hole_data[i]
                sub_volums = next((item[substance] for item in self.data_dict[i] if substance in item), None)
                required_volume =Decimal(str(sub_volums)) * Decimal(hole_num)
                vol+=required_volume
            try:
                if vol > 18:
                    num = math.ceil(vol / 18)
                    result.extend([substance] * num)
                elif substance not in result:
                    result.append(substance)

            except (ValueError, TypeError):
                continue
        return result

    def _process_24_plate(self) -> Tuple[List[List[Union[str, None]]], list]:
        """处理24孔板数据"""
        result = []
        result_1 = []
        for item in self.reagent_plates_96:
            chem1 = list(item[0].keys())[0] if item[0]  else None
            chem2 = list(item[1].keys())[0] if item[1]  else None
            chem3 = list(item[2].keys())[0] if item[2]  else None
            chem4 = list(item[3].keys())[0] if item[3]  else None
            chem5 = list(item[4].keys())[0] if item[4]  else None
            chem6 = list(item[5].keys())[0] if item[5]  else None
            chem7 = list(item[6].keys())[0] if item[6]  else None
            chem8 = list(item[7].keys())[0] if item[7]  else None

            values1 = list(item[0].values())[0] if item[0] else None
            values2 = list(item[1].values())[0] if item[1] else None
            values3 = list(item[2].values())[0] if item[2] else None
            values4 = list(item[3].values())[0] if item[3] else None
            values5 = list(item[4].values())[0] if item[4] else None
            values6 = list(item[5].values())[0] if item[5] else None
            values7 = list(item[6].values())[0] if item[6] else None
            values8 = list(item[7].values())[0] if item[7] else None

            # 先过滤掉None值再求最大值
            valid_values = [v for v in (values1, values2, values3, values4,values5,values6,values7,values8) if v is not None]
            max_vol = max(valid_values) if valid_values else None

            result.append([chem1, chem2,chem3,chem4,chem5,chem6,chem7,chem8])
            result_1.append([chem1, chem2,chem3,chem4,chem5,chem6,chem7,chem8, item[8], max_vol])

        return result, result_1

    def _fill_worksheet(self, wb: Workbook, data_12: List[str],
                        data_24: List[List[Union[str, None]]]) -> int:
        """填充工作表数据"""
        """创建扩展版本工作表"""
        # 填充12孔板物料数据,最大值取决于物料和内标使用的值
        stop_num = self.usage_12_hole
        for i, chem in enumerate(data_12[:stop_num], 2):
            ExcelUtils.modify_cell(wb, "孔位信息", f"D{i}", chem)

        # 处理12孔板扩展数据，llwc_num：记录12孔板用了多少24孔板的位置
        llwc_num,start = 0,14

        if len(data_12) > stop_num:
            for i, chem in enumerate(data_12[stop_num:]):
                row = 14 + i * 8

                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}", chem)
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 1}",chem)
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 2}", chem)
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 3}", chem)
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 4}", chem)
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 5}", chem)
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 6}", chem)
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 7}", chem)
                start = row
                llwc_num += 1
                if llwc_num > 36:
                    """如果12孔板剩余的数量大于6，也就是说一块24孔板不够使用，那么也是应该第三次实验吗？"""
                    pass
                    break

        max_num = 36-llwc_num
        for i, (chem1, chem2,chem3,chem4,chem5,chem6,chem7,chem8) in enumerate(data_24[:max_num]):
            row = start + i * 4
            if chem1:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}", chem1)
            if chem2:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 1}", chem2)
            if chem3:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 2}", chem3)
            if chem4:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 3}", chem4)
            if chem5:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 4}", chem5)
            if chem6:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 5}", chem6)
            if chem7:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 6}", chem7)
            if chem8:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 7}", chem8)
        return max_num
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
                                 data_24: List[List[Union[str, None]]],max_num) -> None:

        # """创建扩展版本工作表"""
        # # 填充12孔板物料数据,最大值取决于物料和内标使用的值
        # stop_num = self.usage_12_hole
        # for i, chem in enumerate(data_12[:stop_num], 2):
        #     ExcelUtils.modify_cell(wb, "孔位信息", f"D{i}", chem)
        #
        # # 处理12孔板扩展数据，llwc_num：记录12孔板用了多少24孔板的位置
        # llwc_num,start = 0,14
        # if len(data_12) > stop_num:
        #     for i, chem in enumerate(data_12[stop_num:]):
        #         row = 14 + i * 8
        #         ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}", chem)
        #         ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 1}",chem)
        #         ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 2}", chem)
        #         ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 3}", chem)
        #         ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 4}", chem)
        #         ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 5}", chem)
        #         ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 6}", chem)
        #         ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 7}", chem)
        #         start = row
        #         llwc_num += 1
        #         if llwc_num > 6:
        #             """如果12孔板剩余的数量大于6，也就是说一块24孔板不够使用，那么也是应该第三次实验吗？"""
        #             pass
        #             break

        for i, (chem1, chem2,chem3,chem4,chem5,chem6,chem7,chem8) in enumerate(data_24):
            row =  i * 8
            if chem1:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row}", chem1)
            if chem2:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 1}", chem2)
            if chem3:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 2}", chem3)
            if chem4:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 3}", chem4)
            if chem5:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 4}", chem5)
            if chem6:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 5}", chem6)
            if chem7:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 6}", chem7)
            if chem8:
                ExcelUtils.modify_cell(wb, "孔位信息", f"D{row + 7}", chem8)


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

    def _create_w11_task(self) :
        """创建W11任务Excel"""
        ExcelUtils.copy_workbook(
            f"{BASE_PATH1}/hiwo任务05-96-01-02.xlsx",
            TEMPLATE_PATHS["wf05"]
        )
        hole_mapping = {}
        for k, perforated_group in enumerate(self.perforated_plate, 1):
            for perforated in perforated_group:
                hole_mapping[perforated] = k
        # 生成转移数据
        transfer_records_12, transfer_records_24, transfer_records_2 = self._generate_transfer_data(hole_mapping)
        '''
        同理：只需要判断transfer_records_2是否有值就可以
        '''
        wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf05"])
        ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_12)
        ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_24)
        ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf05"])
        if transfer_records_2:
            TEMPLATE_NAME = TEMPLATE_PATHS["wf11"]
            ExcelUtils.copy_workbook(f"{BASE_PATH1}/hiwo任务05-96-01-02.xlsx", TEMPLATE_NAME)
            wb = ExcelUtils.open_workbook(TEMPLATE_NAME)
            ExcelUtils.write_records_to_sheet(wb, "移液|||移液信息", transfer_records_2)
            ExcelUtils.save_workbook(wb, TEMPLATE_NAME)

        return transfer_records_12, transfer_records_24, transfer_records_2

    def _generate_transfer_data(self, hole_mapping: dict) -> Tuple[list, list, list]:
        """生成孔板转移数据"""
        transfer_records_12 = []
        transfer_records_24 = []
        transfer_records_2 = []

        hole_data = {}
        for data in self.perforated_plate:
            hole_data.update({data[0]:len(data)})
        # 处理12孔板数据
        num_12 = 0
        for col_idx, (substance, volumes) in enumerate(self.reagent_plates_12_.items(), 1):
            column_remaining = {col_idx: 18}  # 初始列剩余18000ul (18ml)
            current_col = self.data_list_12.index(substance) + 1  # 当前使用的列索引
            add_12_hole = []
            for volume in volumes:
                sub_volums = next((item[substance] for item in self.data_dict[volume] if substance in item), None)
                hole_num = hole_data[volume]
                required_volume = Decimal(str(sub_volums)) * Decimal(hole_num)
                remaining_volume = Decimal(required_volume)
                while remaining_volume > 0:
                    # 获取当前列的剩余体积（若不存在则初始化为18000）
                    current_remaining = Decimal(str(column_remaining.get(current_col, 18)))
                    # 当前列试剂不足时切换到新列
                    if current_remaining <= 0:
                        current_col += 1  # 移动到下一列
                        column_remaining[current_col] = 18  # 初始化新列剩余体积
                        continue  # 重新检查新列
                    # 计算本次实际转移量
                    transfer_vol = min(current_remaining, remaining_volume)

                    target_volume = transfer_vol * 1000  if sub_volums else 0
                    if int(target_volume) <= 200:
                        pipette_location = 'A04'
                    else:
                        pipette_location = 'A05'
                    if len(self.gun_head_dict[pipette_location]) >= 12:
                        if pipette_location == 'A04': pipette_location = 'B04'
                        if pipette_location == 'A05': pipette_location = 'B05'

                    gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0
                    if pipette_location not in add_12_hole:
                        add_12_hole.append(pipette_location)
                        gun_head = gun_head + 1
                        if gun_head  not in self.gun_head_dict[pipette_location]:
                            self.gun_head_dict[pipette_location].append(gun_head )
                    '''self.data_list_12: 12孔板使用排序,
                       usage_12_hole:12孔板内标和淬灭剂使用情况
                       当前试剂在前n个代表肯定是第一个任务表的 
                    '''
                    # 更新列剩余体积

                    if substance in self.data_list_12:

                        transfer_records_12.append({
                            "来源孔板条码": "A02-12",
                            "来源孔板列Y": current_col,
                            "目标孔板条码": "B01-FY96",
                            "目标孔板列Y": hole_mapping[volume],
                            "移液量(ul)": target_volume,
                            "枪头库位": pipette_location,
                            "枪头列Y": gun_head
                        })

                    else:
                        '''
                        改成96孔板
                        不存在这种情况，substance是取自data_list_12
                        '''
                        num_12 += 1
                        transfer_records_24.append({
                            "来源孔板条码": "A03-96",
                            "来源孔板列Y": current_col-12,
                            "目标孔板条码": "B01-FY96",
                            "目标孔板列Y": hole_mapping[volume],
                            "移液量(ul)": target_volume,
                            "枪头库位": pipette_location,
                            "枪头列Y": gun_head + 1
                        })

                        if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                            self.gun_head_dict[pipette_location].append(gun_head + 1)
                    # 更新剩余体积
                    column_remaining[current_col] = float(current_remaining - transfer_vol)
                    remaining_volume -= transfer_vol

        # 处理96孔板数据
        #num_12:12孔板占用孔位数
        num_12
        for plate_idx, (substance1, substance2,substance3,substance4,substance5,substance6,substance7
                        ,substance8, hole_id, amount) in enumerate(self.data_list_24, 1):

            target_volume = amount * 1000  if amount else 0
            if int(target_volume) <= 200:
                pipette_location = 'A04'
            else:
                pipette_location = 'A05'
            if len(self.gun_head_dict[pipette_location]) >= 12:
                if pipette_location == 'A04': pipette_location = 'B04'
                if pipette_location == 'A05': pipette_location = 'B05'

            gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0

            plate_id = plate_idx + num_12
            if self.type == 1:
                bar_code = "A03-96"
                record = {
                    "来源孔板条码": bar_code,
                    "来源孔板列Y": plate_id,
                    "目标孔板条码": "B01-FY96",
                    "目标孔板列Y": hole_mapping[hole_id[0]],
                    "移液量(ul)": target_volume,
                    "枪头库位": pipette_location,
                    "枪头列Y": gun_head + 1}
                if len(transfer_records_24) >= 36:
                    record["来源孔板条码"] = "A02-96"
                    record['来源孔板列Y'] = plate_id-36
                    transfer_records_2.append(record)
                elif len(transfer_records_24) >= 24:
                    record["来源孔板条码"] = "B03-96"
                    record['来源孔板列Y'] = plate_id-24
                    transfer_records_24.append(record)
                else:
                    if len(transfer_records_24) >= 12:
                        record["来源孔板条码"] = "B02-96"
                        record["来源孔板列Y"] = plate_id - 12
                    transfer_records_24.append(record)

                if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                    self.gun_head_dict[pipette_location].append(gun_head + 1)

            else:

                record = {
                    "来源孔板条码": "A03-96",
                    "来源孔板列Y": plate_id + self.usage_24_hole,
                    "目标孔板条码": "B01-FY96",
                    "目标孔板列Y": hole_mapping[hole_id[0]],
                    "移液量(ul)": target_volume,
                    "枪头库位": pipette_location,
                    "枪头列Y": gun_head + 1
                }

                if len(transfer_records_24) >= 36:
                    record["来源孔板条码"] = "A02-96"
                    record["来源孔板列Y"] = plate_id - 36
                    transfer_records_2.append(record)
                else:
                    if len(transfer_records_24) >= 12:
                        record["来源孔板条码"] = "B02-96"
                        record["来源孔板列Y"] = plate_id - 12
                    if len(transfer_records_24) >= 24:
                        record["来源孔板条码"] = "B03-96"
                        record["来源孔板列Y"] = plate_id - 24
                    transfer_records_24.append(record)

                if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                    self.gun_head_dict[pipette_location].append(gun_head + 1)

        return transfer_records_12, transfer_records_24, transfer_records_2

    def _create_task_quencher(self):

        ExcelUtils.copy_workbook(
            f"{BASE_PATH1}/hiwo物料-96-03.xlsx",TEMPLATE_PATHS["mate_info2"])
        ExcelUtils.copy_workbook(
            f"{BASE_PATH1}/hiwo任务11-96-03.xlsx",TEMPLATE_PATHS["wf11"])
        # 淬灭剂
        quenchingAgentData = self.post_data['experimentLog3']["quenchingAgentData"]
        wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info2"])
        if quenchingAgentData:
            # 添加淬灭剂到物料表
            quenchingAgent = quenchingAgentData[0]["name"]
            quenchingAgent_Volume = int(float(quenchingAgentData[0]['addVolume']))
            quenchingAgent_Type = True
            # quenchingAgent_Volume_all = int(quenchingAgent_Volume) * (self.hole_number  * 8)
            #计算所有孔位淬灭剂的使用量
            quenchingAgent_Volume_all = quenchingAgent_Volume * sum(len(sublist) for sublist in self.perforated_plate)
            #需要添加几个格子的淬灭剂
            #(quenchingAgent_Volume_all + 18000 -1) 实现取整效果 0.1==1，1.1==2
            quenchingAgent_Volume_num= (quenchingAgent_Volume_all + 18000 -1) // 18000
            for num in range(quenchingAgent_Volume_num):
                ExcelUtils.modify_cell(wb, "孔位信息", f'D{13-num}', quenchingAgent)
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
            """添加淬灭剂取液到wf11任务表格"""
            wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf11"])
            plate_counter = 0
            plate_remaining = 18000  # 每个孔板最大容量8ml (8000μl)
            for age in range(1, self.hole_number + 1):
                if quenchingAgent_Volume_num ==1:
                    ExcelUtils.modify_row(wb, "加淬灭液|||移液信息", age + 1,
                                          ["SS-A03", 12, "B01-FY96", age, int(quenchingAgent_Volume),
                                           pipette_location, gun_head + 1], start_col=1)
                else:
                    #获取当前横排需要用的量
                    vol_per_hole = len(self.perforated_plate[age-1]) * int(quenchingAgent_Volume)
                    # 检查当前孔板剩余容量是否足够
                    if plate_remaining < vol_per_hole:
                        # 切换到下一个孔板
                        plate_counter += 1
                        plate_remaining = 18000  # 重置为新孔板的完整容量

                    # 更新当前孔板剩余容量
                    plate_remaining -= vol_per_hole

                    ExcelUtils.modify_row(wb, "加淬灭液|||移液信息", age + 1,
                                          ["SS-A03", 12-plate_counter, "B01-FY96", age, int(quenchingAgent_Volume),
                                           pipette_location, gun_head + 1], start_col=1)

            if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                self.gun_head_dict[pipette_location].append(gun_head + 1)
            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf11"])

    def _create_task_internal_standard(self):
        # 内标数据
        internalStandardData = self.post_data['experimentLog3']["internalStandardData"]
        if not internalStandardData:
            return

        if self.post_data['experimentLog3']["quenchingAgentData"]:
            quenchingAgent_Volume = int(float(self.post_data['experimentLog3']["quenchingAgentData"][0]['addVolume']))
            # quenchingAgent_Volume_num = (int(quenchingAgent_Volume) * (self.hole_number * 8) + 18000 - 1) // 18000
            quenchingAgent_Volume_all = int(quenchingAgent_Volume) * sum(len(sublist) for sublist in self.perforated_plate)
            #需要添加几个格子的淬灭剂
            quenchingAgent_Volume_num= (quenchingAgent_Volume_all + 18000 -1) // 18000
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
        internal_standard_used_rows = 13- quenchingAgent_Volume_num
        # 计算所有孔位内标的使用量
        internalStandard_Volume_all = int(internalStandard_Volume) * sum(len(sublist) for sublist in self.perforated_plate)

        # 需要添加几个格子的淬灭剂
        internalStandard_Volume_num = (internalStandard_Volume_all + 18000 - 1) // 18000
        for num in range(internalStandard_Volume_num):
            ExcelUtils.modify_cell(wb, "孔位信息", f'D{internal_standard_used_rows - num}', internalStandard)

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
        plate_counter = 0
        plate_remaining = 18000  # 每个孔板最大容量20ml 取(18000μl)
        for age in range(1, self.hole_number + 1):
            if internalStandard_Volume_num ==1:
                ExcelUtils.modify_row(wb, "加内标|||移液信息", age + 1,
                                      ["SS-A03", internal_standard_used_rows , "B01-FY96", age,
                                       int(internalStandard_Volume) ,
                                       pipette_location, gun_head + 1], start_col=1)

            else:

                vol_per_hole = len(self.perforated_plate[age - 1]) * int(internalStandard_Volume)
                # 检查当前孔板剩余容量是否足够
                if plate_remaining < vol_per_hole:
                    # 切换到下一个孔板
                    plate_counter += 1
                    plate_remaining = 18000  # 重置为新孔板的完整容量

                # 更新当前孔板剩余容量
                plate_remaining -= vol_per_hole

                ExcelUtils.modify_row(wb, "加内标|||移液信息", age + 1,
                                      ["SS-A03", internal_standard_used_rows -plate_counter , "B01-FY96", age, int(internalStandard_Volume),
                                       pipette_location, gun_head + 1], start_col=1)
                # self.usage_12_hole = internal_standard_used_rows - int(internalStandard_Volume_num)
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
        internal_standard_used_rows = 12-quenchingAgent_Volume_num-internalStandardData_Volume_num
        internalStandard = extractionData[0]["name"]
        if internalStandard:
            internalStandard_Volume = 100
            wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["mate_info2"])
            """quenching_used_rows[0]-1: 内标使用的最小孔位-1就是未使用的"""
            # 添加萃取液到物料表
            ExcelUtils.modify_cell(wb, "孔位信息", f'D{internal_standard_used_rows+1}', internalStandard)
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
                                      ["A02-12", internal_standard_used_rows , "2A01-GL96", age, 100,
                                       pipette_location, gun_head + 1], start_col=1)
                # self.usage_12_hole = internal_standard_used_rows - 1

            if gun_head + 1 not in self.gun_head_dict[pipette_location]:
                self.gun_head_dict[pipette_location].append(gun_head + 1)
            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf11"])

    def _create_task_configuration(self) -> None:
        """创建任务配置Excel"""
        exp_log2 = self.post_data["experimentLog2"]
        exp_log3 = self.post_data["experimentLog3"]

        wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf11"])

        # 设置基本参数
        ExcelUtils.modify_cell(wb, "任务参数配置", "D3", "B01-FY96")
        # 温度
        if temp := exp_log2["ReactionConditions"].get("temperature"):
            ExcelUtils.modify_cell(wb, "任务参数配置", "E3", int(temp))
        # 反应时间
        if times := exp_log2["ReactionConditions"].get("time"):
            time = int(times)
            if temp:
                if int(temp)>=26:
                    pass
                elif int(temp)<=60:
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

    def _create_sowo_edit_stuff(self) :
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
            f"{BASE_PATH1}/sowo物料.xlsx",
            TEMPLATE_PATHS["sowo"]
        )
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
                if s == 10:  # 跳过第10号位置
                    s += 1
                if s > 14:  # 第一表最多13个位置
                    break

                # s += 1
                # if s > 14:  # 第二表也最多13个位置
                #     break

                excel_list_2.append({
                    "Code": self.name_key_dict[record["reagent_name"]],
                    "Locate": s,
                    "Type": 2
                })
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

        return excel_list,excel_list_2
        # self.write_codes_to_inner_hopper_stack(fr"{BASE_PATH}/Materials-stack4列.xlsx",fr"{BASE_PATH}/Materials-stack4列.xlsx",excel_list,excel_list_2)


    def _create_mate_msg(self) -> None:
        """创建mateMsg Excel并根据物料表号分表"""
        hole_to_coordinate = {
            1: (1, 1), 2: (1, 2), 3: (1, 3), 4: (1, 4), 5: (1, 5), 6: (1, 6), 7: (1, 7), 8: (1, 8),
            9: (2, 1), 10: (2, 2), 11: (2, 3), 12: (2, 4), 13: (2, 5), 14: (2, 6), 15: (2, 7), 16: (2, 8),
            17: (3, 1), 18: (3, 2), 19: (3, 3), 20: (3, 4), 21: (3, 5), 22: (3, 6), 23: (3, 7), 24: (3, 8),
            25: (4, 1), 26: (4, 2), 27: (4, 3), 28: (4, 4), 29: (4, 5), 30: (4, 6), 31: (4, 7), 32: (4, 8),
            33: (5, 1), 34: (5, 2), 35: (5, 3), 36: (5, 4), 37: (5, 5), 38: (5, 6), 39: (5, 7), 40: (5, 8),
            41: (6, 1), 42: (6, 2), 43: (6, 3), 44: (6, 4), 45: (6, 5), 46: (6, 6), 47: (6, 7), 48: (6, 8),
            49: (7, 1), 50: (7, 2), 51: (7, 3), 52: (7, 4), 53: (7, 5), 54: (7, 6), 55: (7, 7), 56: (7, 8),
            57: (8, 1), 58: (8, 2), 59: (8, 3), 60: (8, 4), 61: (8, 5), 62: (8, 6), 63: (8, 7), 64: (8, 8),
            65: (9, 1), 66: (9, 2), 67: (9, 3), 68: (9, 4), 69: (9, 5), 70: (9, 6), 71: (9, 7), 72: (9, 8),
            73: (10, 1), 74: (10, 2), 75: (10, 3), 76: (10, 4), 77: (10, 5), 78: (10, 6), 79: (10, 7), 80: (10, 8),
            81: (11, 1), 82: (11, 2), 83: (11, 3), 84: (11, 4), 85: (11, 5), 86: (11, 6), 87: (11, 7), 88: (11, 8),
            89: (12, 1), 90: (12, 2), 91: (12, 3), 92: (12, 4), 93: (12, 5), 94: (12, 6), 95: (12, 7), 96: (12, 8)
        }
        # 创建反应ID到孔位映射
        reaction_to_hole = {}
        for group_index, group in enumerate(self.perforated_plate):
            if group[0]:
                reaction_to_hole[group[0]] = group_index * 8 + 1
            if len(group) >= 2:
                reaction_to_hole[group[1]] = group_index * 8 + 2
            if len(group) >= 3:
                reaction_to_hole[group[2]] = group_index * 8 + 3
            if len(group) >= 4:
                reaction_to_hole[group[3]] = group_index * 8 + 4
            if len(group) >= 5:
                reaction_to_hole[group[4]] = group_index * 8 + 5
            if len(group) >= 6:
                reaction_to_hole[group[5]] = group_index * 8 + 6
            if len(group) >= 7:
                reaction_to_hole[group[6]] = group_index * 8 + 7
            if len(group) >= 8:
                reaction_to_hole[group[7]] = group_index * 8 + 8


        # 准备所有数据行
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
                df.at[0, 'PlaterName'] = '96孔'
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
                    output_path = crystallize_TEMPLATE_PATHS["mate"].replace("1-1.xlsx", "2-1.xlsx")
            df.to_excel(output_path, index=False)
        if create_True == 2 and self.type == 1:
            df = pd.DataFrame(columns=[
                "PointX", "PointY", "UsePlateType", "Name",
                "Target", "Tolerance", "UseMode", "Code"
            ])
            df['PlaterName'] = None
            # 仅第一行（索引0）写入 '96孔'
            if len(df) > 0:
                df.at[0, 'PlaterName'] = '96孔'
            df.to_excel(TEMPLATE_PATHS["mate"], index=False)
        elif create_True == 2 and self.type == 2:
            df = pd.DataFrame(columns=[
                "PointX", "PointY", "UsePlateType", "Name",
                "Target", "Tolerance", "UseMode", "Code"
            ])
            df['PlaterName'] = None
            # 仅第一行（索引0）写入 '96孔'
            if len(df) > 0:
                df.at[0, 'PlaterName'] = '96孔'
            df.to_excel(crystallize_TEMPLATE_PATHS["mate"], index=False)


    def _create_dilution_data(self) -> None:
        """创建稀释数据"""
        product_dilution_data = self.post_data["experimentLog3"].get("productDilutionData", [])

        for data in product_dilution_data:
            wb = ExcelUtils.open_workbook(TEMPLATE_PATHS["wf11"])
            procedure = data.get("procedure", {})

            self._process_diluent(wb, "加稀释液-反应|||移液信息", procedure.get("diluent1"), True)
            self._process_diluent(wb, "加稀释液-中转|||移液信息", procedure.get("diluent2"), True, "B02-96")
            self._process_diluent(wb, "加稀释液-过滤|||移液信息", procedure.get("diluent3"), False, "2A01-GL96")
            #反应板混匀
            self._process_mixing(wb, procedure.get("diluent4"), procedure.get("diluent5"))

            ExcelUtils.save_workbook(wb, TEMPLATE_PATHS["wf11"])

    def _process_diluent(self, wb: Workbook, sheet_name: str, diluent: Optional[str], judge, kk="B01-FY96") -> None:
        """处理稀释剂"""
        if not diluent:
            return

        diluent = int(float(diluent))
        pipette_location = "A04" if diluent <= 200 else "A05"
        if len(self.gun_head_dict[pipette_location]) >= 12:
            pipette_location = "B04" if pipette_location == "A04" else "B05"

        gun_head = self.gun_head_dict[pipette_location][-1] if self.gun_head_dict[pipette_location] else 0
        num = 0
        for i in range(2, self.hole_number + 2):
            num += 1
            ExcelUtils.modify_row(wb, sheet_name, i,
                                  ["B03-1", 1, kk, num, int(diluent), pipette_location, gun_head + 1])
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
                                  ["B01-FY96", num, "B02-96", num, int(diluent4) , pipette_location, gun_head + 1])
            ExcelUtils.modify_row(wb, '反应板混匀|||移液信息', i + 1,
                                  ["B02-96", num, "2A01-GL96", num, int(diluent5), pipette_location, gun_head + 1])

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

    #逆合成一次任务
    data96 = {
    "experimentLog1": {
        "wellPlates": "96孔板1ml",
        "SubstratesTableList": [
            {
                "id": 1,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
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
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
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
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
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
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
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
                "equivalent": "1.00",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
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
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
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
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
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
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.50",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
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
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
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
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
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
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
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
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
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
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
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
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.50",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "reagent",
                "smiles": "",
                "substance": "dmap",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
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
                "substance": " triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "",
                "reactionMoles": "1.00",
                "reactionQuality": "0.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "solvent",
                "smiles": "",
                "substance": "tetrahydrofuran",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
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
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
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
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
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
                "equivalent": "2",
                "solvent": "3",
                "reactionMoles": "2.00",
                "reactionQuality": "6.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "2",
                "reactionMoles": "2.00",
                "reactionQuality": "4.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.5",
                "solvent": "1",
                "reactionMoles": "1.50",
                "reactionQuality": "1.50",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "3",
                "reactionMoles": "2.00",
                "reactionQuality": "6.00",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 25,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 25,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 25,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 25,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 26,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 26,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 26,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 26,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 27,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 27,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 27,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 27,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 28,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 28,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 28,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "3",
                "reactionMoles": "2.00",
                "reactionQuality": "6.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 28,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 29,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 29,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 29,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 29,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.50",
                "cas": "",
                "intensity": None
            },
            {
                "id": 30,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 30,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 30,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "3",
                "reactionMoles": "2.00",
                "reactionQuality": "6.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 30,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 31,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 31,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.67",
                "singleCockAddVolume": "0.6",
                "cas": "",
                "intensity": None
            },
            {
                "id": 31,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "0.4",
                "solvent": "2",
                "reactionMoles": "0.40",
                "reactionQuality": "0.80",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 31,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 32,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 32,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 32,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 32,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 33,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 33,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 33,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 33,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 34,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 34,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 34,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "2",
                "reactionMoles": "2.00",
                "reactionQuality": "4.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 34,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 35,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 35,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 35,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 35,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 36,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 36,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 36,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "2",
                "reactionMoles": "2.00",
                "reactionQuality": "4.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 36,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 37,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 37,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 37,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.2",
                "solvent": "3",
                "reactionMoles": "1.20",
                "reactionQuality": "3.60",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 37,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 38,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 38,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 38,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 38,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.50",
                "cas": "",
                "intensity": None
            },
            {
                "id": 39,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 39,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 39,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 39,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 40,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 40,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 40,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 40,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 41,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 41,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 41,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 41,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 42,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 42,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 42,
                "type": "reagent",
                "smiles": "",
                "substance": "triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 42,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 43,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 43,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 43,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "",
                "reactionMoles": "1.00",
                "reactionQuality": "0.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 43,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 44,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 44,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 44,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 44,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 45,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 45,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 45,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "3",
                "reactionMoles": "2.00",
                "reactionQuality": "6.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 45,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 46,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 46,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 46,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 46,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 47,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 47,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 47,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 47,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 48,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "7.14",
                "singleCockAddVolume": "0.14",
                "cas": "",
                "intensity": None
            },
            {
                "id": 48,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.76",
                "singleCockAddVolume": "0.21",
                "cas": "",
                "intensity": None
            },
            {
                "id": 48,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 48,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.45",
                "cas": "",
                "intensity": None
            },
            {
                "id": 49,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 49,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "7.69",
                "singleCockAddVolume": "0.13",
                "cas": "",
                "intensity": None
            },
            {
                "id": 49,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 49,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "",
                "intensity": None
            },
            {
                "id": 50,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 50,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 50,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 50,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 51,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 51,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 51,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 51,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 52,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 52,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 52,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 52,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 53,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 53,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 53,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 53,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 54,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 54,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 54,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 54,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 55,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 55,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 55,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 55,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.50",
                "cas": "",
                "intensity": None
            },
            {
                "id": 56,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 56,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 56,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 56,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 57,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 57,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 57,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 57,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 58,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 58,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 58,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "0.",
                "solvent": "3",
                "reactionMoles": "0.00",
                "reactionQuality": "0.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 58,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.50",
                "cas": "",
                "intensity": None
            },
            {
                "id": 59,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 59,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 59,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "2",
                "reactionMoles": "2.00",
                "reactionQuality": "4.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 59,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 60,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 60,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.03",
                "singleCockAddVolume": "0.33",
                "cas": "",
                "intensity": None
            },
            {
                "id": 60,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "1",
                "reactionMoles": "2.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 60,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "",
                "intensity": None
            },
            {
                "id": 61,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 61,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 61,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 61,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 62,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 62,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 62,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 62,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 63,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 63,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 63,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 63,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.50",
                "cas": "",
                "intensity": None
            },
            {
                "id": 64,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 64,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 64,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 64,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 65,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 65,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 65,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 65,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 66,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 66,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 66,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 66,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 67,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 67,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 67,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 67,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 68,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 68,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 68,
                "type": "reagent",
                "smiles": "",
                "substance": "triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "",
                "reactionMoles": "1.00",
                "reactionQuality": "0.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 68,
                "type": "solvent",
                "smiles": "",
                "substance": "chloroform",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "2",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 69,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 69,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "",
                "intensity": None
            },
            {
                "id": 69,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 69,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "",
                "intensity": None
            }
        ],
        "MaterialsTableList": [
            {
                "type": "sub",
                "substance": "Cc1csc(=S)s1",
                "smiles": "Cc1csc(=S)s1",
                "solvent": 148.28,
                "molecularFormula": "C4H4S3",
                "theoreticalMoles": "69.00",
                "theoreticalQuality": "10231.32",
                "theoreticalVolume": "15.14",
                "configurationMoles": "89.70",
                "cas": "",
                "intensity": None,
                "id": 1,
                "configurationVolume": "19.68",
                "configurationQuality": "13300.72",
                "concentration": "4.56"
            },
            {
                "type": "sub",
                "substance": "O=C1CCC(=O)N1Br",
                "smiles": "O=C1CCC(=O)N1Br",
                "solvent": 177.99,
                "molecularFormula": "C4H4BrNO2",
                "theoreticalMoles": "69.00",
                "theoreticalQuality": "12281.31",
                "theoreticalVolume": "19.30",
                "configurationMoles": "89.70",
                "cas": "",
                "intensity": None,
                "id": 2,
                "configurationVolume": "25.09",
                "configurationQuality": "15965.70",
                "concentration": "3.58"
            },
            {
                "type": "reagent",
                "substance": "LiAlH4",
                "smiles": "",
                "solvent": "1",
                "molecularFormula": "",
                "theoreticalMoles": "72.10",
                "theoreticalQuality": "151.90",
                "theoreticalVolume": "0.00",
                "configurationMoles": "93.73",
                "cas": "",
                "intensity": None,
                "id": 3,
                "configurationVolume": "0.00",
                "configurationQuality": "197.47",
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
                "theoreticalVolume": "19.86",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 4,
                "configurationVolume": "25.82",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "Borane-THF",
                "smiles": "",
                "solvent": "3",
                "molecularFormula": "",
                "theoreticalMoles": "3.00",
                "theoreticalQuality": "5.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "3.90",
                "cas": "",
                "intensity": None,
                "id": 5,
                "configurationVolume": "0.00",
                "configurationQuality": "6.50",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "dmap",
                "smiles": "",
                "solvent": "3",
                "molecularFormula": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "3.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "1.30",
                "cas": "",
                "intensity": None,
                "id": 6,
                "configurationVolume": "0.00",
                "configurationQuality": "3.90",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": " triethylamine",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.10",
                "configurationMoles": "1.30",
                "cas": "",
                "intensity": None,
                "id": 7,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": "10.00"
            },
            {
                "type": "solvent",
                "substance": "tetrahydrofuran",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.20",
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
                "substance": "triethylamine",
                "smiles": "",
                "solvent": "2",
                "molecularFormula": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "2.00",
                "theoreticalVolume": "0.20",
                "configurationMoles": "2.60",
                "cas": "",
                "intensity": None,
                "id": 9,
                "configurationVolume": "5.00",
                "configurationQuality": "2.60",
                "concentration": "10.00"
            },
            {
                "type": "solvent",
                "substance": "dichloromethane",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.10",
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
                "substance": "chloroform",
                "smiles": "",
                "solvent": "2",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.30",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 11,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            }
        ]
    },
    "experimentLog2": {
        "ReactionConditions": {
            "temperature": "55",
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
        "coolingTime": "70",
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
                "volume": "200",
                "time": "30",
                "times": "4",
                "type": "2",
                "density": "0.881"
            }
        ],
        "filtrationData": [
            {
                "name": "硅胶",
                "filterTime": "40"
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
    #96孔逆合成二次任务
    data96_1 = {
        "id":46136543,
    "experimentLog1": {
        "wellPlates": "96孔板1ml",
        "SubstratesTableList": [
            {
                "id": 1,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 1,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 2,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
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
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 3,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
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
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
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
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 4,
                "type": "solvent",
                "smiles": "",
                "substance": "thf4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
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
                "equivalent": "1.00",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 5,
                "type": "solvent",
                "smiles": "",
                "substance": "o-Xylene",
                "molecularFormula": "CC1=CC=CC=C1C",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 106.16,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "95-47-6",
                "intensity": 0.881
            },
            {
                "id": 6,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
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
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 6,
                "type": "solvent",
                "smiles": "",
                "substance": "toluene",
                "molecularFormula": "CC1=CC=CC=C1",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 92.14,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "108-88-3",
                "intensity": 0.866
            },
            {
                "id": 7,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
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
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 7,
                "type": "solvent",
                "smiles": "",
                "substance": "diethylamine",
                "molecularFormula": "CCNCC",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 73.14,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "109-89-7",
                "intensity": 0.707
            },
            {
                "id": 8,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 8,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
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
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.50",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 9,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
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
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 10,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
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
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 11,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 12,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
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
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 13,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 14,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
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
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 15,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
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
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 16,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
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
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.50",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "reagent",
                "smiles": "",
                "substance": "dmap",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
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
                "substance": " triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "",
                "reactionMoles": "1.00",
                "reactionQuality": "0.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 17,
                "type": "solvent",
                "smiles": "",
                "substance": "tetrahydrofuran",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 18,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 19,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
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
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 20,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
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
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 21,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
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
                "equivalent": "2",
                "solvent": "3",
                "reactionMoles": "2.00",
                "reactionQuality": "6.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
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
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "2",
                "reactionMoles": "2.00",
                "reactionQuality": "4.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 22,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 23,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.5",
                "solvent": "1",
                "reactionMoles": "1.50",
                "reactionQuality": "1.50",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 24,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "3",
                "reactionMoles": "2.00",
                "reactionQuality": "6.00",
                "reactionVolume": 0.8,
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
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 25,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 25,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 25,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 25,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 26,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 26,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 26,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 26,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 27,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 27,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 27,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 27,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 28,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 28,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 28,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "3",
                "reactionMoles": "2.00",
                "reactionQuality": "6.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 28,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 29,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 29,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 29,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 29,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.50",
                "cas": "",
                "intensity": None
            },
            {
                "id": 30,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 30,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 30,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "3",
                "reactionMoles": "2.00",
                "reactionQuality": "6.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 30,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 31,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 31,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "1.67",
                "singleCockAddVolume": "0.6",
                "cas": "",
                "intensity": None
            },
            {
                "id": 31,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "0.4",
                "solvent": "2",
                "reactionMoles": "0.40",
                "reactionQuality": "0.80",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 31,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 32,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 32,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 32,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 32,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 33,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 33,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 33,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 33,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 34,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 34,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 34,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "2",
                "reactionMoles": "2.00",
                "reactionQuality": "4.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 34,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 35,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 35,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 35,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 35,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 36,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 36,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 36,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "2",
                "reactionMoles": "2.00",
                "reactionQuality": "4.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 36,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 37,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 37,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.00",
                "singleCockAddVolume": "0.5",
                "cas": "",
                "intensity": None
            },
            {
                "id": 37,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1.2",
                "solvent": "3",
                "reactionMoles": "1.20",
                "reactionQuality": "3.60",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 37,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 38,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 38,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 38,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 38,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.50",
                "cas": "",
                "intensity": None
            },
            {
                "id": 39,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 39,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 39,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 39,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 40,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 40,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 40,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 40,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 41,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 41,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 41,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 41,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 42,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 42,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 42,
                "type": "reagent",
                "smiles": "",
                "substance": "triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 42,
                "type": "solvent",
                "smiles": "",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.10",
                "cas": "",
                "intensity": None
            },
            {
                "id": 43,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 43,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 43,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "",
                "reactionMoles": "1.00",
                "reactionQuality": "0.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 43,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 44,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 44,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 44,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 44,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 45,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 45,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 45,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "3",
                "reactionMoles": "2.00",
                "reactionQuality": "6.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 45,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 46,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 46,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 46,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 46,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 47,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 47,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 47,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 47,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 48,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "7.14",
                "singleCockAddVolume": "0.14",
                "cas": "",
                "intensity": None
            },
            {
                "id": 48,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.76",
                "singleCockAddVolume": "0.21",
                "cas": "",
                "intensity": None
            },
            {
                "id": 48,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 48,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.45",
                "cas": "",
                "intensity": None
            },
            {
                "id": 49,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 49,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "7.69",
                "singleCockAddVolume": "0.13",
                "cas": "",
                "intensity": None
            },
            {
                "id": 49,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 49,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "",
                "intensity": None
            },
            {
                "id": 50,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 50,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 50,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 50,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 51,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 51,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 51,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 51,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 52,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 52,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 52,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 52,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 53,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 53,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 53,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 53,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.20",
                "cas": "",
                "intensity": None
            },
            {
                "id": 54,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 54,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 54,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 54,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 55,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 55,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 55,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 55,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.50",
                "cas": "",
                "intensity": None
            },
            {
                "id": 56,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 56,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 56,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 56,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 57,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 57,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 57,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 57,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 58,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 58,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 58,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "0.",
                "solvent": "3",
                "reactionMoles": "0.00",
                "reactionQuality": "0.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 58,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.50",
                "cas": "",
                "intensity": None
            },
            {
                "id": 59,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "2.50",
                "singleCockAddVolume": "0.4",
                "cas": "",
                "intensity": None
            },
            {
                "id": 59,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 59,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "2",
                "reactionMoles": "2.00",
                "reactionQuality": "4.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 59,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 60,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 60,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.03",
                "singleCockAddVolume": "0.33",
                "cas": "",
                "intensity": None
            },
            {
                "id": 60,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "2",
                "solvent": "1",
                "reactionMoles": "2.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 60,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "",
                "intensity": None
            },
            {
                "id": 61,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 61,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 61,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 61,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 62,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 62,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 62,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 62,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 63,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 63,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 63,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 63,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.50",
                "cas": "",
                "intensity": None
            },
            {
                "id": 64,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 64,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 64,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 64,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 65,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 65,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 65,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "2",
                "reactionMoles": "1.00",
                "reactionQuality": "2.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 65,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 66,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "N,N-dimethyl-formamide",
                "molecularFormula": "CN(C)C=O",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 73.09,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "68-12-2",
                "intensity": 0.944
            },
            {
                "id": 66,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 66,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 66,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 67,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "10.00",
                "singleCockAddVolume": "0.1",
                "cas": "",
                "intensity": None
            },
            {
                "id": 67,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 67,
                "type": "reagent",
                "smiles": "",
                "substance": "LiAlH4",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "3",
                "reactionMoles": "1.00",
                "reactionQuality": "3.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 67,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.40",
                "cas": "",
                "intensity": None
            },
            {
                "id": 68,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 68,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "3.33",
                "singleCockAddVolume": "0.3",
                "cas": "",
                "intensity": None
            },
            {
                "id": 68,
                "type": "reagent",
                "smiles": "",
                "substance": "triethylamine",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "",
                "reactionMoles": "1.00",
                "reactionQuality": "0.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 68,
                "type": "solvent",
                "smiles": "",
                "substance": "chloroform",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "2",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.30",
                "cas": "",
                "intensity": None
            },
            {
                "id": 69,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 69,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "O=C1CCC(=O)N1Br",
                "molecularFormula": "C4H4BrNO2",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 177.99,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "",
                "intensity": None
            },
            {
                "id": 69,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 69,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "",
                "intensity": None
            },
            {
                "id": 70,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 70,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "benzene",
                "molecularFormula": "C1=CC=CC=C1",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 78.11,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "71-43-2",
                "intensity": 0.874
            },
            {
                "id": 70,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 70,
                "type": "solvent",
                "smiles": "",
                "substance": "THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": None,
                "solvent": "",
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "",
                "intensity": None
            },
            {
                "id": 71,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Cc1csc(=S)s1",
                "molecularFormula": "C4H4S3",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 148.28,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "",
                "intensity": None
            },
            {
                "id": 71,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "benzene",
                "molecularFormula": "C1=CC=CC=C1",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 78.11,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "71-43-2",
                "intensity": 0.874
            },
            {
                "id": 71,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 71,
                "type": "solvent",
                "smiles": "",
                "substance": "N,N-dimethyl-formamide",
                "molecularFormula": "CN(C)C=O",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 73.09,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "68-12-2",
                "intensity": 0.944
            },
            {
                "id": 72,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Hexane",
                "molecularFormula": "CCCCCC",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 86.18,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "110-54-3",
                "intensity": 0.659
            },
            {
                "id": 72,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "benzene",
                "molecularFormula": "C1=CC=CC=C1",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 78.11,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "71-43-2",
                "intensity": 0.874
            },
            {
                "id": 72,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 72,
                "type": "solvent",
                "smiles": "",
                "substance": "N,N-dimethyl-formamide",
                "molecularFormula": "CN(C)C=O",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 73.09,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "68-12-2",
                "intensity": 0.944
            },
            {
                "id": 73,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Hexane",
                "molecularFormula": "CCCCCC",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 86.18,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "110-54-3",
                "intensity": 0.659
            },
            {
                "id": 73,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "benzene",
                "molecularFormula": "C1=CC=CC=C1",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 78.11,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "71-43-2",
                "intensity": 0.874
            },
            {
                "id": 73,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 73,
                "type": "solvent",
                "smiles": "",
                "substance": "N,N-dimethyl-formamide",
                "molecularFormula": "CN(C)C=O",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 73.09,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "68-12-2",
                "intensity": 0.944
            },
            {
                "id": 74,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "Hexane",
                "molecularFormula": "CCCCCC",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 86.18,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "110-54-3",
                "intensity": 0.659
            },
            {
                "id": 74,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "benzene",
                "molecularFormula": "C1=CC=CC=C1",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 78.11,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "71-43-2",
                "intensity": 0.874
            },
            {
                "id": 74,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 74,
                "type": "solvent",
                "smiles": "",
                "substance": "N,N-dimethyl-formamide",
                "molecularFormula": "CN(C)C=O",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 73.09,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "68-12-2",
                "intensity": 0.944
            },
            {
                "id": 75,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "toluene",
                "molecularFormula": "CC1=CC=CC=C1",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 92.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "108-88-3",
                "intensity": 0.866
            },
            {
                "id": 75,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "xylene",
                "molecularFormula": "Cc1ccccc1C",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 106.17,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "1330-20-7",
                "intensity": 0.864
            },
            {
                "id": 75,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 75,
                "type": "solvent",
                "smiles": "",
                "substance": "Sodium hydroxide",
                "molecularFormula": "[OH-].[Na+]",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 40,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "1310-73-2",
                "intensity": 2.13
            },
            {
                "id": 76,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "toluene",
                "molecularFormula": "CC1=CC=CC=C1",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 92.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "108-88-3",
                "intensity": 0.866
            },
            {
                "id": 76,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "xylene",
                "molecularFormula": "Cc1ccccc1C",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 106.17,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "1330-20-7",
                "intensity": 0.864
            },
            {
                "id": 76,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 76,
                "type": "solvent",
                "smiles": "",
                "substance": "Sodium hydroxide",
                "molecularFormula": "[OH-].[Na+]",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 40,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "1310-73-2",
                "intensity": 2.13
            },
            {
                "id": 77,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "toluene",
                "molecularFormula": "CC1=CC=CC=C1",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 92.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "108-88-3",
                "intensity": 0.866
            },
            {
                "id": 77,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "xylene",
                "molecularFormula": "Cc1ccccc1C",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 106.17,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "1330-20-7",
                "intensity": 0.864
            },
            {
                "id": 77,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 77,
                "type": "solvent",
                "smiles": "",
                "substance": "Sodium hydroxide",
                "molecularFormula": "[OH-].[Na+]",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 40,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "1310-73-2",
                "intensity": 2.13
            },
            {
                "id": 78,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "toluene",
                "molecularFormula": "CC1=CC=CC=C1",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 92.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "108-88-3",
                "intensity": 0.866
            },
            {
                "id": 78,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "xylene",
                "molecularFormula": "Cc1ccccc1C",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 106.17,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "1330-20-7",
                "intensity": 0.864
            },
            {
                "id": 78,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 78,
                "type": "solvent",
                "smiles": "",
                "substance": "Sodium hydroxide",
                "molecularFormula": "[OH-].[Na+]",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 40,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "1310-73-2",
                "intensity": 2.13
            },
            {
                "id": 79,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "toluene",
                "molecularFormula": "CC1=CC=CC=C1",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 92.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "108-88-3",
                "intensity": 0.866
            },
            {
                "id": 79,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "xylene",
                "molecularFormula": "Cc1ccccc1C",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 106.17,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "1330-20-7",
                "intensity": 0.864
            },
            {
                "id": 79,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 79,
                "type": "solvent",
                "smiles": "",
                "substance": "Sodium hydroxide",
                "molecularFormula": "[OH-].[Na+]",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 40,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "1310-73-2",
                "intensity": 2.13
            },
            {
                "id": 80,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "toluene",
                "molecularFormula": "CC1=CC=CC=C1",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 92.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "108-88-3",
                "intensity": 0.866
            },
            {
                "id": 80,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "xylene",
                "molecularFormula": "Cc1ccccc1C",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 106.17,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "1330-20-7",
                "intensity": 0.864
            },
            {
                "id": 80,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 80,
                "type": "solvent",
                "smiles": "",
                "substance": "Sodium hydroxide",
                "molecularFormula": "[OH-].[Na+]",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 40,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "1310-73-2",
                "intensity": 2.13
            },
            {
                "id": 81,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "toluene",
                "molecularFormula": "CC1=CC=CC=C1",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 92.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "108-88-3",
                "intensity": 0.866
            },
            {
                "id": 81,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "xylene",
                "molecularFormula": "Cc1ccccc1C",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 106.17,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "1330-20-7",
                "intensity": 0.864
            },
            {
                "id": 81,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 81,
                "type": "solvent",
                "smiles": "",
                "substance": "Sodium hydroxide",
                "molecularFormula": "[OH-].[Na+]",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 40,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "1310-73-2",
                "intensity": 2.13
            },
            {
                "id": 82,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "toluene",
                "molecularFormula": "CC1=CC=CC=C1",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 92.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "108-88-3",
                "intensity": 0.866
            },
            {
                "id": 82,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "xylene",
                "molecularFormula": "Cc1ccccc1C",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 106.17,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "1330-20-7",
                "intensity": 0.864
            },
            {
                "id": 82,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 82,
                "type": "solvent",
                "smiles": "",
                "substance": "Sodium hydroxide",
                "molecularFormula": "[OH-].[Na+]",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 40,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "1310-73-2",
                "intensity": 2.13
            },
            {
                "id": 83,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "toluene",
                "molecularFormula": "CC1=CC=CC=C1",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 92.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "108-88-3",
                "intensity": 0.866
            },
            {
                "id": 83,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "xylene",
                "molecularFormula": "Cc1ccccc1C",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 106.17,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "1330-20-7",
                "intensity": 0.864
            },
            {
                "id": 83,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 83,
                "type": "solvent",
                "smiles": "",
                "substance": "Sodium hydroxide",
                "molecularFormula": "[OH-].[Na+]",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 40,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "1310-73-2",
                "intensity": 2.13
            },
            {
                "id": 84,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "toluene",
                "molecularFormula": "CC1=CC=CC=C1",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 92.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "108-88-3",
                "intensity": 0.866
            },
            {
                "id": 84,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "xylene",
                "molecularFormula": "Cc1ccccc1C",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 106.17,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "1330-20-7",
                "intensity": 0.864
            },
            {
                "id": 84,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 84,
                "type": "solvent",
                "smiles": "",
                "substance": "Sodium hydroxide",
                "molecularFormula": "[OH-].[Na+]",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 40,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "1310-73-2",
                "intensity": 2.13
            },
            {
                "id": 85,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "toluene",
                "molecularFormula": "CC1=CC=CC=C1",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 92.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "108-88-3",
                "intensity": 0.866
            },
            {
                "id": 85,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "Acetone",
                "molecularFormula": "CC(=O)C",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 58.08,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "67-64-1",
                "intensity": 0.791
            },
            {
                "id": 85,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 85,
                "type": "solvent",
                "smiles": "",
                "substance": "Sodium hydroxide",
                "molecularFormula": "[OH-].[Na+]",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 40,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "1310-73-2",
                "intensity": 2.13
            },
            {
                "id": 86,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "toluene",
                "molecularFormula": "CC1=CC=CC=C1",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 92.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "108-88-3",
                "intensity": 0.866
            },
            {
                "id": 86,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "Acetone",
                "molecularFormula": "CC(=O)C",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 58.08,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "67-64-1",
                "intensity": 0.791
            },
            {
                "id": 86,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 86,
                "type": "solvent",
                "smiles": "",
                "substance": "Copper(II) sulfate pentahydrate",
                "molecularFormula": "[Cu++].5[OH2]",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 249.68,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "7758-99-8",
                "intensity": 2.284
            },
            {
                "id": 87,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "toluene",
                "molecularFormula": "CC1=CC=CC=C1",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 92.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "108-88-3",
                "intensity": 0.866
            },
            {
                "id": 87,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "Acetone",
                "molecularFormula": "CC(=O)C",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 58.08,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "67-64-1",
                "intensity": 0.791
            },
            {
                "id": 87,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 87,
                "type": "solvent",
                "smiles": "",
                "substance": "Copper(II) sulfate pentahydrate",
                "molecularFormula": "[Cu++].5[OH2]",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 249.68,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "7758-99-8",
                "intensity": 2.284
            },
            {
                "id": 88,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "toluene",
                "molecularFormula": "CC1=CC=CC=C1",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 92.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "108-88-3",
                "intensity": 0.866
            },
            {
                "id": 88,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "Acetone",
                "molecularFormula": "CC(=O)C",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 58.08,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "67-64-1",
                "intensity": 0.791
            },
            {
                "id": 88,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 88,
                "type": "solvent",
                "smiles": "",
                "substance": "Potassium dihydrogen phosphate",
                "molecularFormula": "OP(=O)([O-])O[H].K+",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 136.09,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "7778-77-0",
                "intensity": 2.338
            },
            {
                "id": 89,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "diethylamine",
                "molecularFormula": "CCNCC",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 73.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "109-89-7",
                "intensity": 0.707
            },
            {
                "id": 89,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "Acetone",
                "molecularFormula": "CC(=O)C",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 58.08,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "67-64-1",
                "intensity": 0.791
            },
            {
                "id": 89,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 89,
                "type": "solvent",
                "smiles": "",
                "substance": "Potassium dihydrogen phosphate",
                "molecularFormula": "OP(=O)([O-])O[H].K+",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 136.09,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "7778-77-0",
                "intensity": 2.338
            },
            {
                "id": 90,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "diethylamine",
                "molecularFormula": "CCNCC",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 73.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "109-89-7",
                "intensity": 0.707
            },
            {
                "id": 90,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "Hydrochloric Acid",
                "molecularFormula": "Cl",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 36.46,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "7647-01-0",
                "intensity": 1.045
            },
            {
                "id": 90,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 90,
                "type": "solvent",
                "smiles": "",
                "substance": "Potassium dihydrogen phosphate",
                "molecularFormula": "OP(=O)([O-])O[H].K+",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 136.09,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "7778-77-0",
                "intensity": 2.338
            },
            {
                "id": 91,
                "type": "sub",
                "smiles": "Cc1csc(=S)s1",
                "substance": "diethylamine",
                "molecularFormula": "CCNCC",
                "limitReagents": True,
                "equivalent": 1,
                "solvent": 73.14,
                "reactionMoles": 1,
                "reactionQuality": 148.28,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "5.00",
                "singleCockAddVolume": "0.2",
                "cas": "109-89-7",
                "intensity": 0.707
            },
            {
                "id": 91,
                "type": "sub",
                "smiles": "O=C1CCC(=O)N1Br",
                "substance": "Hydrochloric Acid",
                "molecularFormula": "Cl",
                "limitReagents": False,
                "equivalent": 1,
                "solvent": 36.46,
                "reactionMoles": 1,
                "reactionQuality": 177.99,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "4.35",
                "singleCockAddVolume": "0.23",
                "cas": "7647-01-0",
                "intensity": 1.045
            },
            {
                "id": 91,
                "type": "reagent",
                "smiles": "",
                "substance": "Borane-THF",
                "molecularFormula": "",
                "limitReagents": None,
                "equivalent": "1",
                "solvent": "1",
                "reactionMoles": "1.00",
                "reactionQuality": "1.00",
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "固态",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "",
                "cas": "",
                "intensity": None
            },
            {
                "id": 91,
                "type": "solvent",
                "smiles": "",
                "substance": "Potassium dihydrogen phosphate",
                "molecularFormula": "OP(=O)([O-])O[H].K+",
                "limitReagents": None,
                "equivalent": None,
                "solvent": 136.09,
                "reactionMoles": None,
                "reactionQuality": None,
                "reactionVolume": 0.8,
                "liquidReagentConcentration": "液体",
                "reactionMoleConcentration": "",
                "singleCockAddVolume": "0.37",
                "cas": "7778-77-0",
                "intensity": 2.338
            }
        ],
        "MaterialsTableList": [
            {
                "type": "sub",
                "substance": "Cc1csc(=S)s1",
                "smiles": "Cc1csc(=S)s1",
                "solvent": 148.28,
                "molecularFormula": "C4H4S3",
                "theoreticalMoles": "70.00",
                "theoreticalQuality": "10379.60",
                "theoreticalVolume": "15.34",
                "configurationMoles": "91.00",
                "cas": "",
                "intensity": None,
                "id": 1,
                "configurationVolume": "19.94",
                "configurationQuality": "13493.48",
                "concentration": "4.56"
            },
            {
                "type": "sub",
                "substance": "O=C1CCC(=O)N1Br",
                "smiles": "O=C1CCC(=O)N1Br",
                "solvent": 177.99,
                "molecularFormula": "C4H4BrNO2",
                "theoreticalMoles": "69.00",
                "theoreticalQuality": "12281.31",
                "theoreticalVolume": "19.30",
                "configurationMoles": "89.70",
                "cas": "",
                "intensity": None,
                "id": 2,
                "configurationVolume": "25.09",
                "configurationQuality": "15965.70",
                "concentration": "3.58"
            },
            {
                "type": "reagent",
                "substance": "LiAlH4",
                "smiles": "",
                "solvent": "1",
                "molecularFormula": "",
                "theoreticalMoles": "72.10",
                "theoreticalQuality": "151.90",
                "theoreticalVolume": "0.00",
                "configurationMoles": "93.73",
                "cas": "",
                "intensity": None,
                "id": 3,
                "configurationVolume": "0.00",
                "configurationQuality": "197.47",
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
                "theoreticalVolume": "19.13",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 4,
                "configurationVolume": "24.87",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "thf4",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.30",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 5,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "o-Xylene",
                "smiles": "",
                "solvent": 106.16,
                "molecularFormula": "CC1=CC=CC=C1C",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.20",
                "configurationMoles": "0.00",
                "cas": "95-47-6",
                "intensity": 0.881,
                "id": 6,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "toluene",
                "smiles": "",
                "solvent": 92.14,
                "molecularFormula": "CC1=CC=CC=C1",
                "theoreticalMoles": "14.00",
                "theoreticalQuality": "2075.92",
                "theoreticalVolume": "3.10",
                "configurationMoles": "18.20",
                "cas": "108-88-3",
                "intensity": 0.866,
                "id": 7,
                "configurationVolume": "5.00",
                "configurationQuality": "2698.70",
                "concentration": "4.52"
            },
            {
                "type": "solvent",
                "substance": "diethylamine",
                "smiles": "",
                "solvent": 73.14,
                "molecularFormula": "CCNCC",
                "theoreticalMoles": "3.00",
                "theoreticalQuality": "444.84",
                "theoreticalVolume": "0.90",
                "configurationMoles": "3.90",
                "cas": "109-89-7",
                "intensity": 0.707,
                "id": 8,
                "configurationVolume": "5.00",
                "configurationQuality": "578.29",
                "concentration": "3.33"
            },
            {
                "type": "reagent",
                "substance": "Borane-THF",
                "smiles": "",
                "solvent": "3",
                "molecularFormula": "",
                "theoreticalMoles": "25.00",
                "theoreticalQuality": "27.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "32.50",
                "cas": "",
                "intensity": None,
                "id": 9,
                "configurationVolume": "0.00",
                "configurationQuality": "35.10",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "dmap",
                "smiles": "",
                "solvent": "3",
                "molecularFormula": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "3.00",
                "theoreticalVolume": "0.00",
                "configurationMoles": "1.30",
                "cas": "",
                "intensity": None,
                "id": 10,
                "configurationVolume": "0.00",
                "configurationQuality": "3.90",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": " triethylamine",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.10",
                "configurationMoles": "1.30",
                "cas": "",
                "intensity": None,
                "id": 11,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": "10.00"
            },
            {
                "type": "solvent",
                "substance": "tetrahydrofuran",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.20",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 12,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "triethylamine",
                "smiles": "",
                "solvent": "2",
                "molecularFormula": "",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "2.00",
                "theoreticalVolume": "0.20",
                "configurationMoles": "2.60",
                "cas": "",
                "intensity": None,
                "id": 13,
                "configurationVolume": "5.00",
                "configurationQuality": "2.60",
                "concentration": "10.00"
            },
            {
                "type": "solvent",
                "substance": "dichloromethane",
                "smiles": "",
                "solvent": "",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.10",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 14,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "sub",
                "substance": "N,N-dimethyl-formamide",
                "smiles": "Cc1csc(=S)s1",
                "solvent": 73.09,
                "molecularFormula": "CN(C)C=O",
                "theoreticalMoles": "1.00",
                "theoreticalQuality": "148.28",
                "theoreticalVolume": "1.68",
                "configurationMoles": "1.30",
                "cas": "68-12-2",
                "intensity": 0.944,
                "id": 15,
                "configurationVolume": "5.00",
                "configurationQuality": "192.76",
                "concentration": "0.60"
            },
            {
                "type": "solvent",
                "substance": "chloroform",
                "smiles": "",
                "solvent": "2",
                "molecularFormula": "",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.30",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 16,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "sub",
                "substance": "benzene",
                "smiles": "O=C1CCC(=O)N1Br",
                "solvent": 78.11,
                "molecularFormula": "C1=CC=CC=C1",
                "theoreticalMoles": "5.00",
                "theoreticalQuality": "889.95",
                "theoreticalVolume": "1.15",
                "configurationMoles": "6.50",
                "cas": "71-43-2",
                "intensity": 0.874,
                "id": 17,
                "configurationVolume": "5.00",
                "configurationQuality": "1156.94",
                "concentration": "4.35"
            },
            {
                "type": "sub",
                "substance": "Hexane",
                "smiles": "Cc1csc(=S)s1",
                "solvent": 86.18,
                "molecularFormula": "CCCCCC",
                "theoreticalMoles": "3.00",
                "theoreticalQuality": "444.84",
                "theoreticalVolume": "0.60",
                "configurationMoles": "3.90",
                "cas": "110-54-3",
                "intensity": 0.659,
                "id": 18,
                "configurationVolume": "5.00",
                "configurationQuality": "578.29",
                "concentration": "5.00"
            },
            {
                "type": "sub",
                "substance": "xylene",
                "smiles": "O=C1CCC(=O)N1Br",
                "solvent": 106.17,
                "molecularFormula": "Cc1ccccc1C",
                "theoreticalMoles": "10.00",
                "theoreticalQuality": "1779.90",
                "theoreticalVolume": "2.30",
                "configurationMoles": "13.00",
                "cas": "1330-20-7",
                "intensity": 0.864,
                "id": 19,
                "configurationVolume": "5.00",
                "configurationQuality": "2313.87",
                "concentration": "4.35"
            },
            {
                "type": "solvent",
                "substance": "Sodium hydroxide",
                "smiles": "",
                "solvent": 40,
                "molecularFormula": "[OH-].[Na+]",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "4.07",
                "configurationMoles": "0.00",
                "cas": "1310-73-2",
                "intensity": 2.13,
                "id": 20,
                "configurationVolume": "5.29",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "sub",
                "substance": "Acetone",
                "smiles": "O=C1CCC(=O)N1Br",
                "solvent": 58.08,
                "molecularFormula": "CC(=O)C",
                "theoreticalMoles": "5.00",
                "theoreticalQuality": "889.95",
                "theoreticalVolume": "1.15",
                "configurationMoles": "6.50",
                "cas": "67-64-1",
                "intensity": 0.791,
                "id": 21,
                "configurationVolume": "5.00",
                "configurationQuality": "1156.94",
                "concentration": "4.35"
            },
            {
                "type": "solvent",
                "substance": "Copper(II) sulfate pentahydrate",
                "smiles": "",
                "solvent": 249.68,
                "molecularFormula": "[Cu++].5[OH2]",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.74",
                "configurationMoles": "0.00",
                "cas": "7758-99-8",
                "intensity": 2.284,
                "id": 22,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "Potassium dihydrogen phosphate",
                "smiles": "",
                "solvent": 136.09,
                "molecularFormula": "OP(=O)([O-])O[H].K+",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "1.48",
                "configurationMoles": "0.00",
                "cas": "7778-77-0",
                "intensity": 2.338,
                "id": 23,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "sub",
                "substance": "Hydrochloric Acid",
                "smiles": "O=C1CCC(=O)N1Br",
                "solvent": 36.46,
                "molecularFormula": "Cl",
                "theoreticalMoles": "2.00",
                "theoreticalQuality": "355.98",
                "theoreticalVolume": "0.46",
                "configurationMoles": "2.60",
                "cas": "7647-01-0",
                "intensity": 1.045,
                "id": 24,
                "configurationVolume": "5.00",
                "configurationQuality": "462.77",
                "concentration": "4.35"
            }
        ]
    },
    "experimentLog2": {
        "ReactionConditions": {
            "temperature": "55",
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
        "coolingTime": "70",
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
                "volume": "200",
                "time": "30",
                "times": "4",
                "type": "2",
                "density": "0.881"
            }
        ],
        "filtrationData": [
            {
                "name": "硅胶",
                "filterTime": "40"
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
    # 结晶—溶析-正加
    data_96 = {
    "experimentLog1": {
        "wellPlates": "96孔板1ml",
        "SubstratesTableList": [
            {
                "id": 1,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "2",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 1,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 1,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.4"
            },
            {
                "id": 1,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 2,
                "type": "固体物料",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "smile": "CCCCCC",
                "intensity": 0.659,
                "cas": "110-54-3",
                "quality": "1",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 2,
                "type": "液体物料",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "smile": "CCNCC",
                "intensity": 0.707,
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 2,
                "type": "溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.4"
            },
            {
                "id": 2,
                "type": "反溶剂",
                "substance": "triethylamine",
                "molecularFormula": "",
                "molecularWeight": 101.19,
                "moles": "",
                "intensity": 0.728,
                "smile": "CCN(CC)CC",
                "cas": "121-44-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.14"
            },
            {
                "id": 3,
                "type": "固体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "2.2",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 3,
                "type": "液体物料",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "smile": "CC(=O)C",
                "intensity": 0.791,
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 3,
                "type": "溶剂",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "intensity": 0.707,
                "smile": "CCNCC",
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.13"
            },
            {
                "id": 3,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.4"
            },
            {
                "id": 4,
                "type": "固体物料",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "smile": "CCNCC",
                "intensity": 0.707,
                "cas": "109-89-7",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 4,
                "type": "液体物料",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "smile": "CCCCCC",
                "intensity": 0.659,
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.45"
            },
            {
                "id": 4,
                "type": "溶剂",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "intensity": 0.881,
                "smile": "CC1=CC=CC=C1C",
                "cas": "95-47-6",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.1"
            },
            {
                "id": 4,
                "type": "反溶剂",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "intensity": 0.866,
                "smile": "CC1=CC=CC=C1",
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 5,
                "type": "固体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "2.2",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 5,
                "type": "液体物料",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "smile": "CCCCCC",
                "intensity": 0.659,
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 5,
                "type": "溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.1"
            },
            {
                "id": 5,
                "type": "反溶剂",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "intensity": 0.881,
                "smile": "CC1=CC=CC=C1C",
                "cas": "95-47-6",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.3"
            },
            {
                "id": 6,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "2",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 6,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 6,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.1"
            },
            {
                "id": 6,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 7,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 7,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "2",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 7,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.3"
            },
            {
                "id": 7,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 8,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 8,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 8,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 8,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.26"
            },
            {
                "id": 9,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 9,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "2",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 9,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 9,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.1"
            },
            {
                "id": 10,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "2",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 10,
                "type": "液体物料",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "smile": "CC(=O)C",
                "intensity": 0.791,
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 10,
                "type": "溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.4"
            },
            {
                "id": 10,
                "type": "反溶剂",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "intensity": 0.866,
                "smile": "CC1=CC=CC=C1",
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.1"
            },
            {
                "id": 11,
                "type": "固体物料",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "smile": "CC(=O)C",
                "intensity": 0.791,
                "cas": "67-64-1",
                "quality": "2",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 11,
                "type": "液体物料",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "smile": "C1=CC=CC=C1",
                "intensity": 0.874,
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 11,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.4"
            },
            {
                "id": 11,
                "type": "反溶剂",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "intensity": 0.707,
                "smile": "CCNCC",
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 12,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 12,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "2",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 12,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.1"
            },
            {
                "id": 12,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.3"
            },
            {
                "id": 13,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 13,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "2",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 13,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.5"
            },
            {
                "id": 13,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.1"
            },
            {
                "id": 14,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "2",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 14,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 14,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.3"
            },
            {
                "id": 14,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 15,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 15,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "2",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 15,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 15,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.1"
            },
            {
                "id": 16,
                "type": "固体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 16,
                "type": "液体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "2",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 16,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 16,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.2"
            },
            {
                "id": 17,
                "type": "固体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "2",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 17,
                "type": "液体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 17,
                "type": "溶剂",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "intensity": 0.707,
                "smile": "CCNCC",
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 17,
                "type": "反溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 18,
                "type": "固体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "0.21",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 18,
                "type": "液体物料",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "smile": "CC(=O)C",
                "intensity": 0.791,
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 18,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 18,
                "type": "反溶剂",
                "substance": "triethylamine",
                "molecularFormula": "",
                "molecularWeight": 101.19,
                "moles": "",
                "intensity": 0.728,
                "smile": "CCN(CC)CC",
                "cas": "121-44-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 19,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "0.21",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 19,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 19,
                "type": "溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 19,
                "type": "反溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 20,
                "type": "固体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "0.21",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 20,
                "type": "液体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 20,
                "type": "溶剂",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "intensity": 0.864,
                "smile": "Cc1ccccc1C",
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 20,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 21,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "0.21",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 21,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 21,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 21,
                "type": "反溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 22,
                "type": "固体物料",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "smile": "C1=CC=CC=C1",
                "intensity": 0.874,
                "cas": "71-43-2",
                "quality": "0.21",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 22,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 22,
                "type": "溶剂",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "intensity": 0.864,
                "smile": "Cc1ccccc1C",
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 22,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 23,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "0.21",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 23,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 23,
                "type": "溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 23,
                "type": "反溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 24,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "0.21",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 24,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 24,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 24,
                "type": "反溶剂",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "intensity": 0.866,
                "smile": "CC1=CC=CC=C1",
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.23"
            },
            {
                "id": 25,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "0.21",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 25,
                "type": "液体物料",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "smile": "CCNCC",
                "intensity": 0.707,
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.23"
            },
            {
                "id": 25,
                "type": "溶剂",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "intensity": 0.864,
                "smile": "Cc1ccccc1C",
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.23"
            },
            {
                "id": 25,
                "type": "反溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.23"
            },
            {
                "id": 26,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "0.21",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 26,
                "type": "液体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.23"
            },
            {
                "id": 26,
                "type": "溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.23"
            },
            {
                "id": 26,
                "type": "反溶剂",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "intensity": 0.864,
                "smile": "Cc1ccccc1C",
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.23"
            },
            {
                "id": 27,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "0.23",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 27,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.23"
            },
            {
                "id": 27,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.23"
            },
            {
                "id": 27,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.23"
            },
            {
                "id": 28,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "0.23",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 28,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.23"
            },
            {
                "id": 28,
                "type": "溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.23"
            },
            {
                "id": 28,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.23"
            },
            {
                "id": 29,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "0.23",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 29,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.23"
            },
            {
                "id": 29,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.23"
            },
            {
                "id": 29,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.23"
            },
            {
                "id": 30,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "0.3",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 30,
                "type": "液体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "0.3",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.3"
            },
            {
                "id": 30,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 30,
                "type": "反溶剂",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "intensity": 0.866,
                "smile": "CC1=CC=CC=C1",
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 31,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "0.3",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 31,
                "type": "液体物料",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "smile": "CC(=O)C",
                "intensity": 0.791,
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 31,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 31,
                "type": "反溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 32,
                "type": "固体物料",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "smile": "CCCCCC",
                "intensity": 0.659,
                "cas": "110-54-3",
                "quality": "0.3",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 32,
                "type": "液体物料",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "smile": "Cc1ccccc1C",
                "intensity": 0.864,
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 32,
                "type": "溶剂",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "intensity": 0.707,
                "smile": "CCNCC",
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 32,
                "type": "反溶剂",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "intensity": 0.866,
                "smile": "CC1=CC=CC=C1",
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 33,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "1",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 33,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 33,
                "type": "溶剂",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "intensity": 0.881,
                "smile": "CC1=CC=CC=C1C",
                "cas": "95-47-6",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 33,
                "type": "反溶剂",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "intensity": 0.864,
                "smile": "Cc1ccccc1C",
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 34,
                "type": "固体物料",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "smile": "CC(=O)C",
                "intensity": 0.791,
                "cas": "67-64-1",
                "quality": "1",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 34,
                "type": "液体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 34,
                "type": "溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 34,
                "type": "反溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 35,
                "type": "固体物料",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "smile": "CCCCCC",
                "intensity": 0.659,
                "cas": "110-54-3",
                "quality": "1",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 35,
                "type": "液体物料",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "smile": "CCCCCC",
                "intensity": 0.659,
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 35,
                "type": "溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 35,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 36,
                "type": "固体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "1",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 36,
                "type": "液体物料",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "smile": "Cc1ccccc1C",
                "intensity": 0.864,
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 36,
                "type": "溶剂",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "intensity": 0.707,
                "smile": "CCNCC",
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 36,
                "type": "反溶剂",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "intensity": 0.707,
                "smile": "CCNCC",
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.25"
            },
            {
                "id": 37,
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 37,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 37,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 37,
                "type": "反溶剂",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "intensity": 0.864,
                "smile": "Cc1ccccc1C",
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 38,
                "type": "固体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 38,
                "type": "液体物料",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "smile": "CCCCCC",
                "intensity": 0.659,
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 38,
                "type": "溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 38,
                "type": "反溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 39,
                "type": "固体物料",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "smile": "CCNCC",
                "intensity": 0.707,
                "cas": "109-89-7",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 39,
                "type": "液体物料",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "smile": "C1=CC=CC=C1",
                "intensity": 0.874,
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 39,
                "type": "溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 39,
                "type": "反溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 40,
                "type": "固体物料",
                "substance": "triethylamine",
                "molecularFormula": "",
                "molecularWeight": 101.19,
                "moles": "",
                "smile": "CCN(CC)CC",
                "intensity": 0.728,
                "cas": "121-44-8",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 40,
                "type": "液体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 40,
                "type": "溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 40,
                "type": "反溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 41,
                "type": "固体物料",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "smile": "CCNCC",
                "intensity": 0.707,
                "cas": "109-89-7",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 41,
                "type": "液体物料",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "smile": "Cc1ccccc1C",
                "intensity": 0.864,
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 41,
                "type": "溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 41,
                "type": "反溶剂",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "intensity": 0.881,
                "smile": "CC1=CC=CC=C1C",
                "cas": "95-47-6",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 42,
                "type": "固体物料",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "smile": "C1=CC=CC=C1",
                "intensity": 0.874,
                "cas": "71-43-2",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 42,
                "type": "液体物料",
                "substance": "N,N-dimethyl-formamide",
                "molecularFormula": "",
                "molecularWeight": 73.09,
                "moles": "",
                "smile": "CN(C)C=O",
                "intensity": 0.944,
                "cas": "68-12-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 42,
                "type": "溶剂",
                "substance": "carbon dioxide",
                "molecularFormula": "",
                "molecularWeight": 44.009,
                "moles": "",
                "intensity": 1.53,
                "smile": "C(=O)=O",
                "cas": "124-38-9",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 42,
                "type": "反溶剂",
                "substance": "butan-1-ol",
                "molecularFormula": "",
                "molecularWeight": 74.12,
                "moles": "",
                "intensity": 0.775,
                "smile": "CC(C)(C)O",
                "cas": "75-65-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 43,
                "type": "固体物料",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "smile": "C1=CC=CC=C1",
                "intensity": 0.874,
                "cas": "71-43-2",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 43,
                "type": "液体物料",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "smile": "Cc1ccccc1C",
                "intensity": 0.864,
                "cas": "1330-20-7",
                "quality": "1.5",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 43,
                "type": "溶剂",
                "substance": "ethanol",
                "molecularFormula": "",
                "molecularWeight": 46.07,
                "moles": "",
                "intensity": 0.785,
                "smile": "CCO",
                "cas": "64-17-5",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 43,
                "type": "反溶剂",
                "substance": "dimethyl sulfoxide",
                "molecularFormula": "",
                "molecularWeight": 78.14,
                "moles": "",
                "intensity": 1.1,
                "smile": "CS(=O)C",
                "cas": "67-68-5",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 44,
                "type": "固体物料",
                "substance": "triethylamine",
                "molecularFormula": "",
                "molecularWeight": 101.19,
                "moles": "",
                "smile": "CCN(CC)CC",
                "intensity": 0.728,
                "cas": "121-44-8",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 44,
                "type": "液体物料",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "smile": "Cc1ccccc1C",
                "intensity": 0.864,
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 44,
                "type": "溶剂",
                "substance": "triethylamine",
                "molecularFormula": "",
                "molecularWeight": 101.19,
                "moles": "",
                "intensity": 0.728,
                "smile": "CCN(CC)CC",
                "cas": "121-44-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 44,
                "type": "反溶剂",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "intensity": 0.864,
                "smile": "Cc1ccccc1C",
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 45,
                "type": "固体物料",
                "substance": "acetonitrile",
                "molecularFormula": "",
                "molecularWeight": 41.05,
                "moles": "",
                "smile": "CC#N",
                "intensity": 0.786,
                "cas": "75-05-8",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 45,
                "type": "液体物料",
                "substance": "acetonitrile",
                "molecularFormula": "",
                "molecularWeight": 41.05,
                "moles": "",
                "smile": "CC#N",
                "intensity": 0.786,
                "cas": "75-05-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 45,
                "type": "溶剂",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "intensity": 0.707,
                "smile": "CCNCC",
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 45,
                "type": "反溶剂",
                "substance": "1-methyl-pyrrolidin-2-one",
                "molecularFormula": "",
                "molecularWeight": 99.13,
                "moles": "",
                "intensity": 1.027,
                "smile": "CN1CCCC1=O",
                "cas": "872-50-4",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 46,
                "type": "固体物料",
                "substance": "dimethyl sulfoxide",
                "molecularFormula": "",
                "molecularWeight": 78.14,
                "moles": "",
                "smile": "CS(=O)C",
                "intensity": 1.1,
                "cas": "67-68-5",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 46,
                "type": "液体物料",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "smile": "CCNCC",
                "intensity": 0.707,
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 46,
                "type": "溶剂",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "molecularWeight": 84.93,
                "moles": "",
                "intensity": 1.325,
                "smile": "C(Cl)Cl",
                "cas": "75-09-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 46,
                "type": "反溶剂",
                "substance": "butan-1-ol",
                "molecularFormula": "",
                "molecularWeight": 74.12,
                "moles": "",
                "intensity": 0.775,
                "smile": "CC(C)(C)O",
                "cas": "75-65-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 47,
                "type": "固体物料",
                "substance": "acetonitrile",
                "molecularFormula": "",
                "molecularWeight": 41.05,
                "moles": "",
                "smile": "CC#N",
                "intensity": 0.786,
                "cas": "75-05-8",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 47,
                "type": "液体物料",
                "substance": "water",
                "molecularFormula": "",
                "molecularWeight": 18.015,
                "moles": "",
                "smile": "O",
                "intensity": 0.995,
                "cas": "7732-18-5",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 47,
                "type": "溶剂",
                "substance": "ethanol",
                "molecularFormula": "",
                "molecularWeight": 46.07,
                "moles": "",
                "intensity": 0.785,
                "smile": "CCO",
                "cas": "64-17-5",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 47,
                "type": "反溶剂",
                "substance": "1,2-dichloro-ethane",
                "molecularFormula": "",
                "molecularWeight": 98.96,
                "moles": "",
                "intensity": 1.245,
                "smile": "C(CCl)Cl",
                "cas": "107-06-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 48,
                "type": "固体物料",
                "substance": "tetrahydrofuran",
                "molecularFormula": "",
                "molecularWeight": 72.11,
                "moles": "",
                "smile": "C1CCOC1",
                "intensity": 0.883,
                "cas": "109-99-9",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 48,
                "type": "液体物料",
                "substance": "carbon dioxide",
                "molecularFormula": "",
                "molecularWeight": 44.009,
                "moles": "",
                "smile": "C(=O)=O",
                "intensity": 1.53,
                "cas": "124-38-9",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 48,
                "type": "溶剂",
                "substance": "pyridine",
                "molecularFormula": "",
                "molecularWeight": 79.1,
                "moles": "",
                "intensity": 0.983,
                "smile": "C1=CC=NC=C1",
                "cas": "110-86-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 48,
                "type": "反溶剂",
                "substance": "Diethyl Ether",
                "molecularFormula": "",
                "molecularWeight": 74.12,
                "moles": "",
                "intensity": 0.713,
                "smile": "CCOCC",
                "cas": "60-29-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 49,
                "type": "固体物料",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "smile": "C(=O)(C(F)(F)F)O",
                "intensity": 1.489,
                "cas": "76-05-1",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 49,
                "type": "液体物料",
                "substance": "N-Methylmorpholine",
                "molecularFormula": "",
                "molecularWeight": 101.15,
                "moles": "",
                "smile": "CN1CCOCC1",
                "intensity": 0.92,
                "cas": "109-02-4",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 49,
                "type": "溶剂",
                "substance": "N-Methylmorpholine",
                "molecularFormula": "",
                "molecularWeight": 101.15,
                "moles": "",
                "intensity": 0.92,
                "smile": "CN1CCOCC1",
                "cas": "109-02-4",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 49,
                "type": "反溶剂",
                "substance": "pyridine",
                "molecularFormula": "",
                "molecularWeight": 79.1,
                "moles": "",
                "intensity": 0.983,
                "smile": "C1=CC=NC=C1",
                "cas": "110-86-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 50,
                "type": "固体物料",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "smile": "C(=O)(C(F)(F)F)O",
                "intensity": 1.489,
                "cas": "76-05-1",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 50,
                "type": "液体物料",
                "substance": "Potassium dihydrogen phosphate",
                "molecularFormula": "",
                "molecularWeight": 136.09,
                "moles": "",
                "smile": "OP(=O)([O-])O[H].K+",
                "intensity": 2.338,
                "cas": "7778-77-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 50,
                "type": "溶剂",
                "substance": "carbon dioxide",
                "molecularFormula": "",
                "molecularWeight": 44.009,
                "moles": "",
                "intensity": 1.53,
                "smile": "C(=O)=O",
                "cas": "124-38-9",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 50,
                "type": "反溶剂",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "intensity": 1.489,
                "smile": "C(=O)(C(F)(F)F)O",
                "cas": "76-05-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 51,
                "type": "固体物料",
                "substance": "Oxalic acid dihydrate",
                "molecularFormula": "",
                "molecularWeight": 126.07,
                "moles": "",
                "smile": "OC(=O)C(=O)O.O.O",
                "intensity": 1.65,
                "cas": "6153-56-6",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 51,
                "type": "液体物料",
                "substance": "Sodium hydroxide",
                "molecularFormula": "",
                "molecularWeight": 40,
                "moles": "",
                "smile": "[OH-].[Na+]",
                "intensity": 2.13,
                "cas": "1310-73-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 51,
                "type": "溶剂",
                "substance": "Potassium dihydrogen phosphate",
                "molecularFormula": "",
                "molecularWeight": 136.09,
                "moles": "",
                "intensity": 2.338,
                "smile": "OP(=O)([O-])O[H].K+",
                "cas": "7778-77-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 51,
                "type": "反溶剂",
                "substance": "Barium sulfate",
                "molecularFormula": "",
                "molecularWeight": 233.39,
                "moles": "",
                "intensity": 4.49,
                "smile": "[Ba++].[O-]S(=O)(=O)[O-]",
                "cas": "7727-43-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 52,
                "type": "固体物料",
                "substance": "Sodium bicarbonate",
                "molecularFormula": "",
                "molecularWeight": 84.007,
                "moles": "",
                "smile": "C(=O)(O)[O-].[Na+]",
                "intensity": 2.2,
                "cas": "144-55-8",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 52,
                "type": "液体物料",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "molecularWeight": 84.93,
                "moles": "",
                "smile": "C(Cl)Cl",
                "intensity": 1.325,
                "cas": "75-09-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 52,
                "type": "溶剂",
                "substance": "ethanol",
                "molecularFormula": "",
                "molecularWeight": 46.07,
                "moles": "",
                "intensity": 0.785,
                "smile": "CCO",
                "cas": "64-17-5",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 52,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 53,
                "type": "固体物料",
                "substance": "N-Methylmorpholine",
                "molecularFormula": "",
                "molecularWeight": 101.15,
                "moles": "",
                "smile": "CN1CCOCC1",
                "intensity": 0.92,
                "cas": "109-02-4",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 53,
                "type": "液体物料",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "smile": "C(=O)(C(F)(F)F)O",
                "intensity": 1.489,
                "cas": "76-05-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 53,
                "type": "溶剂",
                "substance": "isopropyl alcohol",
                "molecularFormula": "",
                "molecularWeight": 60.1,
                "moles": "",
                "intensity": 0.785,
                "smile": "CC(C)O",
                "cas": "67-63-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 53,
                "type": "反溶剂",
                "substance": "tetrahydrofuran",
                "molecularFormula": "",
                "molecularWeight": 72.11,
                "moles": "",
                "intensity": 0.883,
                "smile": "C1CCOC1",
                "cas": "109-99-9",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 54,
                "type": "固体物料",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "smile": "CCNCC",
                "intensity": 0.707,
                "cas": "109-89-7",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 54,
                "type": "液体物料",
                "substance": "methanol",
                "molecularFormula": "",
                "molecularWeight": 32.042,
                "moles": "",
                "smile": "CO",
                "intensity": 0.791,
                "cas": "67-56-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 54,
                "type": "溶剂",
                "substance": "methanol",
                "molecularFormula": "",
                "molecularWeight": 32.042,
                "moles": "",
                "intensity": 0.791,
                "smile": "CO",
                "cas": "67-56-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 54,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 55,
                "type": "固体物料",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "smile": "C1=CC=CC=C1",
                "intensity": 0.874,
                "cas": "71-43-2",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 55,
                "type": "液体物料",
                "substance": "ethanol",
                "molecularFormula": "",
                "molecularWeight": 46.07,
                "moles": "",
                "smile": "CCO",
                "intensity": 0.785,
                "cas": "64-17-5",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 55,
                "type": "溶剂",
                "substance": "acetic acid",
                "molecularFormula": "",
                "molecularWeight": 60.05,
                "moles": "",
                "intensity": 1.045,
                "smile": "CC(=O)O",
                "cas": "64-19-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 55,
                "type": "反溶剂",
                "substance": "isopropyl alcohol",
                "molecularFormula": "",
                "molecularWeight": 60.1,
                "moles": "",
                "intensity": 0.785,
                "smile": "CC(C)O",
                "cas": "67-63-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 56,
                "type": "固体物料",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "smile": "Cc1ccccc1C",
                "intensity": 0.864,
                "cas": "1330-20-7",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 56,
                "type": "液体物料",
                "substance": "triethylamine",
                "molecularFormula": "",
                "molecularWeight": 101.19,
                "moles": "",
                "smile": "CCN(CC)CC",
                "intensity": 0.728,
                "cas": "121-44-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 56,
                "type": "溶剂",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "molecularWeight": 84.93,
                "moles": "",
                "intensity": 1.325,
                "smile": "C(Cl)Cl",
                "cas": "75-09-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 56,
                "type": "反溶剂",
                "substance": "acetic acid",
                "molecularFormula": "",
                "molecularWeight": 60.05,
                "moles": "",
                "intensity": 1.045,
                "smile": "CC(=O)O",
                "cas": "64-19-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 57,
                "type": "固体物料",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "molecularWeight": 84.93,
                "moles": "",
                "smile": "C(Cl)Cl",
                "intensity": 1.325,
                "cas": "75-09-2",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 57,
                "type": "液体物料",
                "substance": "water",
                "molecularFormula": "",
                "molecularWeight": 18.015,
                "moles": "",
                "smile": "O",
                "intensity": 0.995,
                "cas": "7732-18-5",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 57,
                "type": "溶剂",
                "substance": "acetonitrile",
                "molecularFormula": "",
                "molecularWeight": 41.05,
                "moles": "",
                "intensity": 0.786,
                "smile": "CC#N",
                "cas": "75-05-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 57,
                "type": "反溶剂",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "intensity": 0.707,
                "smile": "CCNCC",
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 58,
                "type": "固体物料",
                "substance": "carbon dioxide",
                "molecularFormula": "",
                "molecularWeight": 44.009,
                "moles": "",
                "smile": "C(=O)=O",
                "intensity": 1.53,
                "cas": "124-38-9",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 58,
                "type": "液体物料",
                "substance": "1-methyl-pyrrolidin-2-one",
                "molecularFormula": "",
                "molecularWeight": 99.13,
                "moles": "",
                "smile": "CN1CCCC1=O",
                "intensity": 1.027,
                "cas": "872-50-4",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 58,
                "type": "溶剂",
                "substance": "Diethyl Ether",
                "molecularFormula": "",
                "molecularWeight": 74.12,
                "moles": "",
                "intensity": 0.713,
                "smile": "CCOCC",
                "cas": "60-29-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 58,
                "type": "反溶剂",
                "substance": "1-methyl-pyrrolidin-2-one",
                "molecularFormula": "",
                "molecularWeight": 99.13,
                "moles": "",
                "intensity": 1.027,
                "smile": "CN1CCCC1=O",
                "cas": "872-50-4",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 59,
                "type": "固体物料",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "smile": "CC(=O)C",
                "intensity": 0.791,
                "cas": "67-64-1",
                "quality": "1.4",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 59,
                "type": "液体物料",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "smile": "CCNCC",
                "intensity": 0.707,
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 59,
                "type": "溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 59,
                "type": "反溶剂",
                "substance": "triethylamine",
                "molecularFormula": "",
                "molecularWeight": 101.19,
                "moles": "",
                "intensity": 0.728,
                "smile": "CCN(CC)CC",
                "cas": "121-44-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 60,
                "type": "固体物料",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "smile": "CCNCC",
                "intensity": 0.707,
                "cas": "109-89-7",
                "quality": "1.4",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 60,
                "type": "液体物料",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "smile": "CC(=O)C",
                "intensity": 0.791,
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 60,
                "type": "溶剂",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "intensity": 0.707,
                "smile": "CCNCC",
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 60,
                "type": "反溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 61,
                "type": "固体物料",
                "substance": "N,N-dimethyl-formamide",
                "molecularFormula": "",
                "molecularWeight": 73.09,
                "moles": "",
                "smile": "CN(C)C=O",
                "intensity": 0.944,
                "cas": "68-12-2",
                "quality": "1.4",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 61,
                "type": "液体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "cas": "95-47-6",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 61,
                "type": "溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 61,
                "type": "反溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 62,
                "type": "固体物料",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "smile": "CC(=O)C",
                "intensity": 0.791,
                "cas": "67-64-1",
                "quality": "1.4",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 62,
                "type": "液体物料",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "smile": "CCNCC",
                "intensity": 0.707,
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 62,
                "type": "溶剂",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "intensity": 0.866,
                "smile": "CC1=CC=CC=C1",
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 62,
                "type": "反溶剂",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "intensity": 0.864,
                "smile": "Cc1ccccc1C",
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 63,
                "type": "固体物料",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "smile": "C1=CC=CC=C1",
                "intensity": 0.874,
                "cas": "71-43-2",
                "quality": "1.4",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 63,
                "type": "液体物料",
                "substance": "Oxalic acid dihydrate",
                "molecularFormula": "",
                "molecularWeight": 126.07,
                "moles": "",
                "smile": "OC(=O)C(=O)O.O.O",
                "intensity": 1.65,
                "cas": "6153-56-6",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 63,
                "type": "溶剂",
                "substance": "Barium sulfate",
                "molecularFormula": "",
                "molecularWeight": 233.39,
                "moles": "",
                "intensity": 4.49,
                "smile": "[Ba++].[O-]S(=O)(=O)[O-]",
                "cas": "7727-43-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 63,
                "type": "反溶剂",
                "substance": "Sodium bicarbonate",
                "molecularFormula": "",
                "molecularWeight": 84.007,
                "moles": "",
                "intensity": 2.2,
                "smile": "C(=O)(O)[O-].[Na+]",
                "cas": "144-55-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 64,
                "type": "固体物料",
                "substance": "Hydrochloric Acid",
                "molecularFormula": "",
                "molecularWeight": 36.46,
                "moles": "",
                "smile": "Cl",
                "intensity": 1.045,
                "cas": "7647-01-0",
                "quality": "1.4",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 64,
                "type": "液体物料",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "smile": "C(=O)(C(F)(F)F)O",
                "intensity": 1.489,
                "cas": "76-05-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 64,
                "type": "溶剂",
                "substance": "water",
                "molecularFormula": "",
                "molecularWeight": 18.015,
                "moles": "",
                "intensity": 0.995,
                "smile": "O",
                "cas": "7732-18-5",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 64,
                "type": "反溶剂",
                "substance": "Potassium dihydrogen phosphate",
                "molecularFormula": "",
                "molecularWeight": 136.09,
                "moles": "",
                "intensity": 2.338,
                "smile": "OP(=O)([O-])O[H].K+",
                "cas": "7778-77-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 65,
                "type": "固体物料",
                "substance": "Diethyl Ether",
                "molecularFormula": "",
                "molecularWeight": 74.12,
                "moles": "",
                "smile": "CCOCC",
                "intensity": 0.713,
                "cas": "60-29-7",
                "quality": "1.4",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 65,
                "type": "液体物料",
                "substance": "Oxalic acid dihydrate",
                "molecularFormula": "",
                "molecularWeight": 126.07,
                "moles": "",
                "smile": "OC(=O)C(=O)O.O.O",
                "intensity": 1.65,
                "cas": "6153-56-6",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 65,
                "type": "溶剂",
                "substance": "Sodium acetate",
                "molecularFormula": "",
                "molecularWeight": 82.03,
                "moles": "",
                "intensity": 1.528,
                "smile": "CC(=O)[O-].[Na+]",
                "cas": "127-09-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 65,
                "type": "反溶剂",
                "substance": "Copper(II) sulfate pentahydrate",
                "molecularFormula": "",
                "molecularWeight": 249.68,
                "moles": "",
                "intensity": 2.284,
                "smile": "[Cu++].5[OH2]",
                "cas": "7758-99-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 66,
                "type": "固体物料",
                "substance": "carbon dioxide",
                "molecularFormula": "",
                "molecularWeight": 44.009,
                "moles": "",
                "smile": "C(=O)=O",
                "intensity": 1.53,
                "cas": "124-38-9",
                "quality": "1.4",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 66,
                "type": "液体物料",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "smile": "C(=O)(C(F)(F)F)O",
                "intensity": 1.489,
                "cas": "76-05-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 66,
                "type": "溶剂",
                "substance": "Calcium carbonate",
                "molecularFormula": "",
                "molecularWeight": 100.09,
                "moles": "",
                "intensity": 2.71,
                "smile": "[Ca++]([O-]C(=O)[O-])([O-]C(=O)[O-])",
                "cas": "471-34-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 66,
                "type": "反溶剂",
                "substance": "Potassium iodide",
                "molecularFormula": "",
                "molecularWeight": 166,
                "moles": "",
                "intensity": 3.13,
                "smile": "[I-].[K+]",
                "cas": "7681-11-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 67,
                "type": "固体物料",
                "substance": "Potassium dihydrogen phosphate",
                "molecularFormula": "",
                "molecularWeight": 136.09,
                "moles": "",
                "smile": "OP(=O)([O-])O[H].K+",
                "intensity": 2.338,
                "cas": "7778-77-0",
                "quality": "1.4",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 67,
                "type": "液体物料",
                "substance": "Potassium dihydrogen phosphate",
                "molecularFormula": "",
                "molecularWeight": 136.09,
                "moles": "",
                "smile": "OP(=O)([O-])O[H].K+",
                "intensity": 2.338,
                "cas": "7778-77-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 67,
                "type": "溶剂",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "intensity": 1.489,
                "smile": "C(=O)(C(F)(F)F)O",
                "cas": "76-05-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 67,
                "type": "反溶剂",
                "substance": "Potassium iodide",
                "molecularFormula": "",
                "molecularWeight": 166,
                "moles": "",
                "intensity": 3.13,
                "smile": "[I-].[K+]",
                "cas": "7681-11-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 68,
                "type": "固体物料",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "smile": "C1=CC=CC=C1",
                "intensity": 0.874,
                "cas": "71-43-2",
                "quality": "1.4",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 68,
                "type": "液体物料",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "smile": "Cc1ccccc1C",
                "intensity": 0.864,
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 68,
                "type": "溶剂",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "intensity": 0.864,
                "smile": "Cc1ccccc1C",
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 68,
                "type": "反溶剂",
                "substance": "tetrahydrofuran",
                "molecularFormula": "",
                "molecularWeight": 72.11,
                "moles": "",
                "intensity": 0.883,
                "smile": "C1CCOC1",
                "cas": "109-99-9",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 69,
                "type": "固体物料",
                "substance": "1-methyl-pyrrolidin-2-one",
                "molecularFormula": "",
                "molecularWeight": 99.13,
                "moles": "",
                "smile": "CN1CCCC1=O",
                "intensity": 1.027,
                "cas": "872-50-4",
                "quality": "1.4",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 69,
                "type": "液体物料",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "smile": "C(=O)(C(F)(F)F)O",
                "intensity": 1.489,
                "cas": "76-05-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 69,
                "type": "溶剂",
                "substance": "pyridine",
                "molecularFormula": "",
                "molecularWeight": 79.1,
                "moles": "",
                "intensity": 0.983,
                "smile": "C1=CC=NC=C1",
                "cas": "110-86-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 69,
                "type": "反溶剂",
                "substance": "1,2-dichloro-ethane",
                "molecularFormula": "",
                "molecularWeight": 98.96,
                "moles": "",
                "intensity": 1.245,
                "smile": "C(CCl)Cl",
                "cas": "107-06-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 70,
                "type": "固体物料",
                "substance": "N-Methylmorpholine",
                "molecularFormula": "",
                "molecularWeight": 101.15,
                "moles": "",
                "smile": "CN1CCOCC1",
                "intensity": 0.92,
                "cas": "109-02-4",
                "quality": "1.4",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 70,
                "type": "液体物料",
                "substance": "triethylamine",
                "molecularFormula": "",
                "molecularWeight": 101.19,
                "moles": "",
                "smile": "CCN(CC)CC",
                "intensity": 0.728,
                "cas": "121-44-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 70,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 70,
                "type": "反溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 71,
                "type": "固体物料",
                "substance": "triethylamine",
                "molecularFormula": "",
                "molecularWeight": 101.19,
                "moles": "",
                "smile": "CCN(CC)CC",
                "intensity": 0.728,
                "cas": "121-44-8",
                "quality": "1.4",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 71,
                "type": "液体物料",
                "substance": "1-methyl-pyrrolidin-2-one",
                "molecularFormula": "",
                "molecularWeight": 99.13,
                "moles": "",
                "smile": "CN1CCCC1=O",
                "intensity": 1.027,
                "cas": "872-50-4",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 71,
                "type": "溶剂",
                "substance": "N-Methylmorpholine",
                "molecularFormula": "",
                "molecularWeight": 101.15,
                "moles": "",
                "intensity": 0.92,
                "smile": "CN1CCOCC1",
                "cas": "109-02-4",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 71,
                "type": "反溶剂",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "intensity": 1.489,
                "smile": "C(=O)(C(F)(F)F)O",
                "cas": "76-05-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.21"
            },
            {
                "id": 72,
                "type": "固体物料",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "smile": "C1=CC=CC=C1",
                "intensity": 0.874,
                "cas": "71-43-2",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 72,
                "type": "液体物料",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "smile": "Cc1ccccc1C",
                "intensity": 0.864,
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 72,
                "type": "溶剂",
                "substance": "triethylamine",
                "molecularFormula": "",
                "molecularWeight": 101.19,
                "moles": "",
                "intensity": 0.728,
                "smile": "CCN(CC)CC",
                "cas": "121-44-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 72,
                "type": "反溶剂",
                "substance": "Copper(II) sulfate pentahydrate",
                "molecularFormula": "",
                "molecularWeight": 249.68,
                "moles": "",
                "intensity": 2.284,
                "smile": "[Cu++].5[OH2]",
                "cas": "7758-99-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 73,
                "type": "固体物料",
                "substance": "Copper(II) sulfate pentahydrate",
                "molecularFormula": "",
                "molecularWeight": 249.68,
                "moles": "",
                "smile": "[Cu++].5[OH2]",
                "intensity": 2.284,
                "cas": "7758-99-8",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 73,
                "type": "液体物料",
                "substance": "Sodium bicarbonate",
                "molecularFormula": "",
                "molecularWeight": 84.007,
                "moles": "",
                "smile": "C(=O)(O)[O-].[Na+]",
                "intensity": 2.2,
                "cas": "144-55-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 73,
                "type": "溶剂",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "intensity": 1.489,
                "smile": "C(=O)(C(F)(F)F)O",
                "cas": "76-05-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 73,
                "type": "反溶剂",
                "substance": "N-Methylmorpholine",
                "molecularFormula": "",
                "molecularWeight": 101.15,
                "moles": "",
                "intensity": 0.92,
                "smile": "CN1CCOCC1",
                "cas": "109-02-4",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 74,
                "type": "固体物料",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "smile": "C(=O)(C(F)(F)F)O",
                "intensity": 1.489,
                "cas": "76-05-1",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 74,
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 74,
                "type": "溶剂",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "intensity": 0.864,
                "smile": "Cc1ccccc1C",
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 74,
                "type": "反溶剂",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "intensity": 1.489,
                "smile": "C(=O)(C(F)(F)F)O",
                "cas": "76-05-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 75,
                "type": "固体物料",
                "substance": "Hydrochloric Acid",
                "molecularFormula": "",
                "molecularWeight": 36.46,
                "moles": "",
                "smile": "Cl",
                "intensity": 1.045,
                "cas": "7647-01-0",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 75,
                "type": "液体物料",
                "substance": "water",
                "molecularFormula": "",
                "molecularWeight": 18.015,
                "moles": "",
                "smile": "O",
                "intensity": 0.995,
                "cas": "7732-18-5",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 75,
                "type": "溶剂",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "intensity": 0.866,
                "smile": "CC1=CC=CC=C1",
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 75,
                "type": "反溶剂",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "intensity": 1.489,
                "smile": "C(=O)(C(F)(F)F)O",
                "cas": "76-05-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 76,
                "type": "固体物料",
                "substance": "dimethyl sulfoxide",
                "molecularFormula": "",
                "molecularWeight": 78.14,
                "moles": "",
                "smile": "CS(=O)C",
                "intensity": 1.1,
                "cas": "67-68-5",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 76,
                "type": "液体物料",
                "substance": "Sodium bicarbonate",
                "molecularFormula": "",
                "molecularWeight": 84.007,
                "moles": "",
                "smile": "C(=O)(O)[O-].[Na+]",
                "intensity": 2.2,
                "cas": "144-55-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 76,
                "type": "溶剂",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "moles": "",
                "intensity": 0.866,
                "smile": "CC1=CC=CC=C1",
                "cas": "108-88-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 76,
                "type": "反溶剂",
                "substance": "1,2-dichloro-ethane",
                "molecularFormula": "",
                "molecularWeight": 98.96,
                "moles": "",
                "intensity": 1.245,
                "smile": "C(CCl)Cl",
                "cas": "107-06-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 77,
                "type": "固体物料",
                "substance": "Potassium iodide",
                "molecularFormula": "",
                "molecularWeight": 166,
                "moles": "",
                "smile": "[I-].[K+]",
                "intensity": 3.13,
                "cas": "7681-11-0",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 77,
                "type": "液体物料",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "smile": "C1=CC=CC=C1",
                "intensity": 0.874,
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 77,
                "type": "溶剂",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "intensity": 0.707,
                "smile": "CCNCC",
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 77,
                "type": "反溶剂",
                "substance": "Oxalic acid dihydrate",
                "molecularFormula": "",
                "molecularWeight": 126.07,
                "moles": "",
                "intensity": 1.65,
                "smile": "OC(=O)C(=O)O.O.O",
                "cas": "6153-56-6",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 78,
                "type": "固体物料",
                "substance": "Potassium iodide",
                "molecularFormula": "",
                "molecularWeight": 166,
                "moles": "",
                "smile": "[I-].[K+]",
                "intensity": 3.13,
                "cas": "7681-11-0",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 78,
                "type": "液体物料",
                "substance": "Barium sulfate",
                "molecularFormula": "",
                "molecularWeight": 233.39,
                "moles": "",
                "smile": "[Ba++].[O-]S(=O)(=O)[O-]",
                "intensity": 4.49,
                "cas": "7727-43-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 78,
                "type": "溶剂",
                "substance": "Hydrochloric Acid",
                "molecularFormula": "",
                "molecularWeight": 36.46,
                "moles": "",
                "intensity": 1.045,
                "smile": "Cl",
                "cas": "7647-01-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 78,
                "type": "反溶剂",
                "substance": "Sodium hydroxide",
                "molecularFormula": "",
                "molecularWeight": 40,
                "moles": "",
                "intensity": 2.13,
                "smile": "[OH-].[Na+]",
                "cas": "1310-73-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 79,
                "type": "固体物料",
                "substance": "N-Methylmorpholine",
                "molecularFormula": "",
                "molecularWeight": 101.15,
                "moles": "",
                "smile": "CN1CCOCC1",
                "intensity": 0.92,
                "cas": "109-02-4",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 79,
                "type": "液体物料",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "smile": "C1=CC=CC=C1",
                "intensity": 0.874,
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 79,
                "type": "溶剂",
                "substance": "Calcium carbonate",
                "molecularFormula": "",
                "molecularWeight": 100.09,
                "moles": "",
                "intensity": 2.71,
                "smile": "[Ca++]([O-]C(=O)[O-])([O-]C(=O)[O-])",
                "cas": "471-34-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 79,
                "type": "反溶剂",
                "substance": "Oxalic acid dihydrate",
                "molecularFormula": "",
                "molecularWeight": 126.07,
                "moles": "",
                "intensity": 1.65,
                "smile": "OC(=O)C(=O)O.O.O",
                "cas": "6153-56-6",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 80,
                "type": "固体物料",
                "substance": "Barium sulfate",
                "molecularFormula": "",
                "molecularWeight": 233.39,
                "moles": "",
                "smile": "[Ba++].[O-]S(=O)(=O)[O-]",
                "intensity": 4.49,
                "cas": "7727-43-7",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 80,
                "type": "液体物料",
                "substance": "Potassium iodide",
                "molecularFormula": "",
                "molecularWeight": 166,
                "moles": "",
                "smile": "[I-].[K+]",
                "intensity": 3.13,
                "cas": "7681-11-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 80,
                "type": "溶剂",
                "substance": "Copper(II) sulfate pentahydrate",
                "molecularFormula": "",
                "molecularWeight": 249.68,
                "moles": "",
                "intensity": 2.284,
                "smile": "[Cu++].5[OH2]",
                "cas": "7758-99-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 80,
                "type": "反溶剂",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "moles": "",
                "intensity": 0.881,
                "smile": "CC1=CC=CC=C1C",
                "cas": "95-47-6",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 81,
                "type": "固体物料",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "smile": "CCCCCC",
                "intensity": 0.659,
                "cas": "110-54-3",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 81,
                "type": "液体物料",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "moles": "",
                "smile": "Cc1ccccc1C",
                "intensity": 0.864,
                "cas": "1330-20-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 81,
                "type": "溶剂",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "intensity": 1.489,
                "smile": "C(=O)(C(F)(F)F)O",
                "cas": "76-05-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 81,
                "type": "反溶剂",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "intensity": 0.707,
                "smile": "CCNCC",
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 82,
                "type": "固体物料",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "smile": "CCCCCC",
                "intensity": 0.659,
                "cas": "110-54-3",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 82,
                "type": "液体物料",
                "substance": "carbon dioxide",
                "molecularFormula": "",
                "molecularWeight": 44.009,
                "moles": "",
                "smile": "C(=O)=O",
                "intensity": 1.53,
                "cas": "124-38-9",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 82,
                "type": "溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 82,
                "type": "反溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 83,
                "type": "固体物料",
                "substance": "N-Methylmorpholine",
                "molecularFormula": "",
                "molecularWeight": 101.15,
                "moles": "",
                "smile": "CN1CCOCC1",
                "intensity": 0.92,
                "cas": "109-02-4",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 83,
                "type": "液体物料",
                "substance": "Sodium bicarbonate",
                "molecularFormula": "",
                "molecularWeight": 84.007,
                "moles": "",
                "smile": "C(=O)(O)[O-].[Na+]",
                "intensity": 2.2,
                "cas": "144-55-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 83,
                "type": "溶剂",
                "substance": "Hydrochloric Acid",
                "molecularFormula": "",
                "molecularWeight": 36.46,
                "moles": "",
                "intensity": 1.045,
                "smile": "Cl",
                "cas": "7647-01-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 83,
                "type": "反溶剂",
                "substance": "water",
                "molecularFormula": "",
                "molecularWeight": 18.015,
                "moles": "",
                "intensity": 0.995,
                "smile": "O",
                "cas": "7732-18-5",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 84,
                "type": "固体物料",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "smile": "C1=CC=CC=C1",
                "intensity": 0.874,
                "cas": "71-43-2",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 84,
                "type": "液体物料",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "smile": "C1=CC=CC=C1",
                "intensity": 0.874,
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 84,
                "type": "溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 84,
                "type": "反溶剂",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "intensity": 1.489,
                "smile": "C(=O)(C(F)(F)F)O",
                "cas": "76-05-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 85,
                "type": "固体物料",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "smile": "C(=O)(C(F)(F)F)O",
                "intensity": 1.489,
                "cas": "76-05-1",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 85,
                "type": "液体物料",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "smile": "CCCCCC",
                "intensity": 0.659,
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 85,
                "type": "溶剂",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 85,
                "type": "反溶剂",
                "substance": "Diethyl Ether",
                "molecularFormula": "",
                "molecularWeight": 74.12,
                "moles": "",
                "intensity": 0.713,
                "smile": "CCOCC",
                "cas": "60-29-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 86,
                "type": "固体物料",
                "substance": "triethylamine",
                "molecularFormula": "",
                "molecularWeight": 101.19,
                "moles": "",
                "smile": "CCN(CC)CC",
                "intensity": 0.728,
                "cas": "121-44-8",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 86,
                "type": "液体物料",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "smile": "CC(=O)C",
                "intensity": 0.791,
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 86,
                "type": "溶剂",
                "substance": "Diethyl Ether",
                "molecularFormula": "",
                "molecularWeight": 74.12,
                "moles": "",
                "intensity": 0.713,
                "smile": "CCOCC",
                "cas": "60-29-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 86,
                "type": "反溶剂",
                "substance": "Diethyl Ether",
                "molecularFormula": "",
                "molecularWeight": 74.12,
                "moles": "",
                "intensity": 0.713,
                "smile": "CCOCC",
                "cas": "60-29-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 87,
                "type": "固体物料",
                "substance": "Hydrochloric Acid",
                "molecularFormula": "",
                "molecularWeight": 36.46,
                "moles": "",
                "smile": "Cl",
                "intensity": 1.045,
                "cas": "7647-01-0",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 87,
                "type": "液体物料",
                "substance": "Sodium chloride",
                "molecularFormula": "",
                "molecularWeight": 58.44,
                "moles": "",
                "smile": "[Na+].[Cl-]",
                "intensity": 2.16,
                "cas": "7647-14-5",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 87,
                "type": "溶剂",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "intensity": 1.489,
                "smile": "C(=O)(C(F)(F)F)O",
                "cas": "76-05-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 87,
                "type": "反溶剂",
                "substance": "Sodium bicarbonate",
                "molecularFormula": "",
                "molecularWeight": 84.007,
                "moles": "",
                "intensity": 2.2,
                "smile": "C(=O)(O)[O-].[Na+]",
                "cas": "144-55-8",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 88,
                "type": "固体物料",
                "substance": "Hydrochloric Acid",
                "molecularFormula": "",
                "molecularWeight": 36.46,
                "moles": "",
                "smile": "Cl",
                "intensity": 1.045,
                "cas": "7647-01-0",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 88,
                "type": "液体物料",
                "substance": "Hydrochloric Acid",
                "molecularFormula": "",
                "molecularWeight": 36.46,
                "moles": "",
                "smile": "Cl",
                "intensity": 1.045,
                "cas": "7647-01-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 88,
                "type": "溶剂",
                "substance": "Potassium dihydrogen phosphate",
                "molecularFormula": "",
                "molecularWeight": 136.09,
                "moles": "",
                "intensity": 2.338,
                "smile": "OP(=O)([O-])O[H].K+",
                "cas": "7778-77-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 88,
                "type": "反溶剂",
                "substance": "Diethyl Ether",
                "molecularFormula": "",
                "molecularWeight": 74.12,
                "moles": "",
                "intensity": 0.713,
                "smile": "CCOCC",
                "cas": "60-29-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 89,
                "type": "固体物料",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "molecularWeight": 84.93,
                "moles": "",
                "smile": "C(Cl)Cl",
                "intensity": 1.325,
                "cas": "75-09-2",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 89,
                "type": "液体物料",
                "substance": "Potassium iodide",
                "molecularFormula": "",
                "molecularWeight": 166,
                "moles": "",
                "smile": "[I-].[K+]",
                "intensity": 3.13,
                "cas": "7681-11-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 89,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 89,
                "type": "反溶剂",
                "substance": "N,N-dimethyl-formamide",
                "molecularFormula": "",
                "molecularWeight": 73.09,
                "moles": "",
                "intensity": 0.944,
                "smile": "CN(C)C=O",
                "cas": "68-12-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 90,
                "type": "固体物料",
                "substance": "butan-1-ol",
                "molecularFormula": "",
                "molecularWeight": 74.12,
                "moles": "",
                "smile": "CC(C)(C)O",
                "intensity": 0.775,
                "cas": "75-65-0",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 90,
                "type": "液体物料",
                "substance": "1-methyl-pyrrolidin-2-one",
                "molecularFormula": "",
                "molecularWeight": 99.13,
                "moles": "",
                "smile": "CN1CCCC1=O",
                "intensity": 1.027,
                "cas": "872-50-4",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 90,
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "moles": "",
                "intensity": 0.791,
                "smile": "CC(=O)C",
                "cas": "67-64-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 90,
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "intensity": 0.874,
                "smile": "C1=CC=CC=C1",
                "cas": "71-43-2",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 91,
                "type": "固体物料",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "moles": "",
                "smile": "C1=CC=CC=C1",
                "intensity": 0.874,
                "cas": "71-43-2",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 91,
                "type": "液体物料",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "moles": "",
                "smile": "CCNCC",
                "intensity": 0.707,
                "cas": "109-89-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 91,
                "type": "溶剂",
                "substance": "1-methyl-pyrrolidin-2-one",
                "molecularFormula": "",
                "molecularWeight": 99.13,
                "moles": "",
                "intensity": 1.027,
                "smile": "CN1CCCC1=O",
                "cas": "872-50-4",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 91,
                "type": "反溶剂",
                "substance": "Calcium carbonate",
                "molecularFormula": "",
                "molecularWeight": 100.09,
                "moles": "",
                "intensity": 2.71,
                "smile": "[Ca++]([O-]C(=O)[O-])([O-]C(=O)[O-])",
                "cas": "471-34-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 92,
                "type": "固体物料",
                "substance": "water",
                "molecularFormula": "",
                "molecularWeight": 18.015,
                "moles": "",
                "smile": "O",
                "intensity": 0.995,
                "cas": "7732-18-5",
                "quality": "1.5",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 92,
                "type": "液体物料",
                "substance": "Potassium dihydrogen phosphate",
                "molecularFormula": "",
                "molecularWeight": 136.09,
                "moles": "",
                "smile": "OP(=O)([O-])O[H].K+",
                "intensity": 2.338,
                "cas": "7778-77-0",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 92,
                "type": "液体物料",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "moles": "",
                "intensity": 0.659,
                "smile": "CCCCCC",
                "cas": "110-54-3",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 92,
                "type": "溶剂",
                "substance": "Calcium carbonate",
                "molecularFormula": "",
                "molecularWeight": 100.09,
                "moles": "",
                "intensity": 2.71,
                "smile": "[Ca++]([O-]C(=O)[O-])([O-]C(=O)[O-])",
                "cas": "471-34-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 92,
                "type": "反溶剂",
                "substance": "water",
                "molecularFormula": "",
                "molecularWeight": 18.015,
                "moles": "",
                "intensity": 0.995,
                "smile": "O",
                "cas": "7732-18-5",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 93,
                "type": "固体物料",
                "substance": "Sodium acetate",
                "molecularFormula": "",
                "molecularWeight": 82.03,
                "moles": "",
                "smile": "CC(=O)[O-].[Na+]",
                "intensity": 1.528,
                "cas": "127-09-3",
                "quality": "1.5",
                "Volume": "",
                "moleConcentration": "",
                "singleCockAddVolume": ""
            },
            {
                "id": 93,
                "type": "液体物料",
                "substance": "N-Methylmorpholine",
                "molecularFormula": "",
                "molecularWeight": 101.15,
                "moles": "",
                "smile": "CN1CCOCC1",
                "intensity": 0.92,
                "cas": "109-02-4",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 93,
                "type": "溶剂",
                "substance": "Diethyl Ether",
                "molecularFormula": "",
                "molecularWeight": 74.12,
                "moles": "",
                "intensity": 0.713,
                "smile": "CCOCC",
                "cas": "60-29-7",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            },
            {
                "id": 93,
                "type": "反溶剂",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "moles": "",
                "intensity": 1.489,
                "smile": "C(=O)(C(F)(F)F)O",
                "cas": "76-05-1",
                "quality": "",
                "Volume": "0.80",
                "moleConcentration": "",
                "singleCockAddVolume": "0.15"
            }
        ],
        "MaterialsTableList": [
            {
                "type": "固体物料",
                "substance": "o-Xylene",
                "molecularFormula": "",
                "molecularWeight": 106.16,
                "cas": "95-47-6",
                "smile": "CC1=CC=CC=C1C",
                "intensity": 0.881,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "16.85",
                "theoreticalVolume": "2.83",
                "id": 1,
                "configurationVolume": "5.00",
                "configurationQuality": "21.91",
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "toluene",
                "molecularFormula": "",
                "molecularWeight": 92.14,
                "cas": "108-88-3",
                "smile": "CC1=CC=CC=C1",
                "intensity": 0.866,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "19.32",
                "theoreticalVolume": "5.69",
                "id": 2,
                "configurationVolume": "7.40",
                "configurationQuality": "25.12",
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "Acetone",
                "molecularFormula": "",
                "molecularWeight": 58.08,
                "cas": "67-64-1",
                "smile": "CC(=O)C",
                "intensity": 0.791,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "5.80",
                "theoreticalVolume": "7.63",
                "id": 3,
                "configurationVolume": "9.92",
                "configurationQuality": "7.54",
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "benzene",
                "molecularFormula": "",
                "molecularWeight": 78.11,
                "cas": "71-43-2",
                "smile": "C1=CC=CC=C1",
                "intensity": 0.874,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "12.01",
                "theoreticalVolume": "7.06",
                "id": 4,
                "configurationVolume": "9.18",
                "configurationQuality": "15.61",
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "Hexane",
                "molecularFormula": "",
                "molecularWeight": 86.18,
                "cas": "110-54-3",
                "smile": "CCCCCC",
                "intensity": 0.659,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "5.30",
                "theoreticalVolume": "5.12",
                "id": 5,
                "configurationVolume": "6.66",
                "configurationQuality": "6.89",
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "diethylamine",
                "molecularFormula": "",
                "molecularWeight": 73.14,
                "cas": "109-89-7",
                "smile": "CCNCC",
                "intensity": 0.707,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "7.40",
                "theoreticalVolume": "3.43",
                "id": 6,
                "configurationVolume": "5.00",
                "configurationQuality": "9.62",
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "triethylamine",
                "molecularFormula": "",
                "molecularWeight": 101.19,
                "cas": "121-44-8",
                "smile": "CCN(CC)CC",
                "intensity": 0.728,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "5.90",
                "theoreticalVolume": "1.34",
                "id": 7,
                "configurationVolume": "5.00",
                "configurationQuality": "7.67",
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "xylene",
                "molecularFormula": "",
                "molecularWeight": 106.17,
                "cas": "1330-20-7",
                "smile": "Cc1ccccc1C",
                "intensity": 0.864,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "3.00",
                "theoreticalVolume": "3.76",
                "id": 8,
                "configurationVolume": "5.00",
                "configurationQuality": "3.90",
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "N,N-dimethyl-formamide",
                "molecularFormula": "",
                "molecularWeight": 73.09,
                "cas": "68-12-2",
                "smile": "CN(C)C=O",
                "intensity": 0.944,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.40",
                "theoreticalVolume": "0.36",
                "id": 9,
                "configurationVolume": "5.00",
                "configurationQuality": "1.82",
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "carbon dioxide",
                "molecularFormula": "",
                "molecularWeight": 44.009,
                "cas": "124-38-9",
                "smile": "C(=O)=O",
                "intensity": 1.53,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "2.90",
                "theoreticalVolume": "0.78",
                "id": 10,
                "configurationVolume": "5.00",
                "configurationQuality": "3.77",
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "butan-1-ol",
                "molecularFormula": "",
                "molecularWeight": 74.12,
                "cas": "75-65-0",
                "smile": "CC(C)(C)O",
                "intensity": 0.775,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.50",
                "theoreticalVolume": "0.42",
                "id": 11,
                "configurationVolume": "5.00",
                "configurationQuality": "1.95",
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "ethanol",
                "molecularFormula": "",
                "molecularWeight": 46.07,
                "cas": "64-17-5",
                "smile": "CCO",
                "intensity": 0.785,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.84",
                "id": 12,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "dimethyl sulfoxide",
                "molecularFormula": "",
                "molecularWeight": 78.14,
                "cas": "67-68-5",
                "smile": "CS(=O)C",
                "intensity": 1.1,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "3.00",
                "theoreticalVolume": "0.21",
                "id": 13,
                "configurationVolume": "5.00",
                "configurationQuality": "3.90",
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "acetonitrile",
                "molecularFormula": "",
                "molecularWeight": 41.05,
                "cas": "75-05-8",
                "smile": "CC#N",
                "intensity": 0.786,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "3.00",
                "theoreticalVolume": "0.42",
                "id": 14,
                "configurationVolume": "5.00",
                "configurationQuality": "3.90",
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "1-methyl-pyrrolidin-2-one",
                "molecularFormula": "",
                "molecularWeight": 99.13,
                "cas": "872-50-4",
                "smile": "CN1CCCC1=O",
                "intensity": 1.027,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.40",
                "theoreticalVolume": "1.14",
                "id": 15,
                "configurationVolume": "5.00",
                "configurationQuality": "1.82",
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "dichloromethane",
                "molecularFormula": "",
                "molecularWeight": 84.93,
                "cas": "75-09-2",
                "smile": "C(Cl)Cl",
                "intensity": 1.325,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "3.00",
                "theoreticalVolume": "0.63",
                "id": 16,
                "configurationVolume": "5.00",
                "configurationQuality": "3.90",
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "water",
                "molecularFormula": "",
                "molecularWeight": 18.015,
                "cas": "7732-18-5",
                "smile": "O",
                "intensity": 0.995,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.50",
                "theoreticalVolume": "1.08",
                "id": 17,
                "configurationVolume": "5.00",
                "configurationQuality": "1.95",
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "1,2-dichloro-ethane",
                "molecularFormula": "",
                "molecularWeight": 98.96,
                "cas": "107-06-2",
                "smile": "C(CCl)Cl",
                "intensity": 1.245,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.57",
                "id": 18,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "tetrahydrofuran",
                "molecularFormula": "",
                "molecularWeight": 72.11,
                "cas": "109-99-9",
                "smile": "C1CCOC1",
                "intensity": 0.883,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.50",
                "theoreticalVolume": "0.42",
                "id": 19,
                "configurationVolume": "5.00",
                "configurationQuality": "1.95",
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "pyridine",
                "molecularFormula": "",
                "molecularWeight": 79.1,
                "cas": "110-86-1",
                "smile": "C1=CC=NC=C1",
                "intensity": 0.983,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.63",
                "id": 20,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "Diethyl Ether",
                "molecularFormula": "",
                "molecularWeight": 74.12,
                "cas": "60-29-7",
                "smile": "CCOCC",
                "intensity": 0.713,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.40",
                "theoreticalVolume": "1.17",
                "id": 21,
                "configurationVolume": "5.00",
                "configurationQuality": "1.82",
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "Trifluoroacetic Acid",
                "molecularFormula": "",
                "molecularWeight": 114.02,
                "cas": "76-05-1",
                "smile": "C(=O)(C(F)(F)F)O",
                "intensity": 1.489,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "6.00",
                "theoreticalVolume": "2.52",
                "id": 22,
                "configurationVolume": "5.00",
                "configurationQuality": "7.80",
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "N-Methylmorpholine",
                "molecularFormula": "",
                "molecularWeight": 101.15,
                "cas": "109-02-4",
                "smile": "CN1CCOCC1",
                "intensity": 0.92,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "5.90",
                "theoreticalVolume": "0.93",
                "id": 23,
                "configurationVolume": "5.00",
                "configurationQuality": "7.67",
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "Potassium dihydrogen phosphate",
                "molecularFormula": "",
                "molecularWeight": 136.09,
                "cas": "7778-77-0",
                "smile": "OP(=O)([O-])O[H].K+",
                "intensity": 2.338,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.40",
                "theoreticalVolume": "1.14",
                "id": 24,
                "configurationVolume": "5.00",
                "configurationQuality": "1.82",
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "Oxalic acid dihydrate",
                "molecularFormula": "",
                "molecularWeight": 126.07,
                "cas": "6153-56-6",
                "smile": "OC(=O)C(=O)O.O.O",
                "intensity": 1.65,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.50",
                "theoreticalVolume": "0.72",
                "id": 25,
                "configurationVolume": "5.00",
                "configurationQuality": "1.95",
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "Sodium hydroxide",
                "molecularFormula": "",
                "molecularWeight": 40,
                "cas": "1310-73-2",
                "smile": "[OH-].[Na+]",
                "intensity": 2.13,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.36",
                "id": 26,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "Barium sulfate",
                "molecularFormula": "",
                "molecularWeight": 233.39,
                "cas": "7727-43-7",
                "smile": "[Ba++].[O-]S(=O)(=O)[O-]",
                "intensity": 4.49,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.50",
                "theoreticalVolume": "0.57",
                "id": 27,
                "configurationVolume": "5.00",
                "configurationQuality": "1.95",
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "Sodium bicarbonate",
                "molecularFormula": "",
                "molecularWeight": 84.007,
                "cas": "144-55-8",
                "smile": "C(=O)(O)[O-].[Na+]",
                "intensity": 2.2,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.50",
                "theoreticalVolume": "0.81",
                "id": 28,
                "configurationVolume": "5.00",
                "configurationQuality": "1.95",
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "isopropyl alcohol",
                "molecularFormula": "",
                "molecularWeight": 60.1,
                "cas": "67-63-0",
                "smile": "CC(C)O",
                "intensity": 0.785,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.42",
                "id": 29,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "methanol",
                "molecularFormula": "",
                "molecularWeight": 32.042,
                "cas": "67-56-1",
                "smile": "CO",
                "intensity": 0.791,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.42",
                "id": 30,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "acetic acid",
                "molecularFormula": "",
                "molecularWeight": 60.05,
                "cas": "64-19-7",
                "smile": "CC(=O)O",
                "intensity": 1.045,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.42",
                "id": 31,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "固体物料",
                "substance": "Hydrochloric Acid",
                "molecularFormula": "",
                "molecularWeight": 36.46,
                "cas": "7647-01-0",
                "smile": "Cl",
                "intensity": 1.045,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "5.90",
                "theoreticalVolume": "0.45",
                "id": 32,
                "configurationVolume": "5.00",
                "configurationQuality": "7.67",
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "Sodium acetate",
                "molecularFormula": "",
                "molecularWeight": 82.03,
                "cas": "127-09-3",
                "smile": "CC(=O)[O-].[Na+]",
                "intensity": 1.528,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.50",
                "theoreticalVolume": "0.21",
                "id": 33,
                "configurationVolume": "5.00",
                "configurationQuality": "1.95",
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "Copper(II) sulfate pentahydrate",
                "molecularFormula": "",
                "molecularWeight": 249.68,
                "cas": "7758-99-8",
                "smile": "[Cu++].5[OH2]",
                "intensity": 2.284,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.50",
                "theoreticalVolume": "0.51",
                "id": 34,
                "configurationVolume": "5.00",
                "configurationQuality": "1.95",
                "concentration": ""
            },
            {
                "type": "溶剂",
                "substance": "Calcium carbonate",
                "molecularFormula": "",
                "molecularWeight": 100.09,
                "cas": "471-34-1",
                "smile": "[Ca++]([O-]C(=O)[O-])([O-]C(=O)[O-])",
                "intensity": 2.71,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.66",
                "id": 35,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            },
            {
                "type": "反溶剂",
                "substance": "Potassium iodide",
                "molecularFormula": "",
                "molecularWeight": 166,
                "cas": "7681-11-0",
                "smile": "[I-].[K+]",
                "intensity": 3.13,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "3.00",
                "theoreticalVolume": "0.72",
                "id": 36,
                "configurationVolume": "5.00",
                "configurationQuality": "3.90",
                "concentration": ""
            },
            {
                "type": "液体物料",
                "substance": "Sodium chloride",
                "molecularFormula": "",
                "molecularWeight": 58.44,
                "cas": "7647-14-5",
                "smile": "[Na+].[Cl-]",
                "intensity": 2.16,
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.00",
                "theoreticalVolume": "0.15",
                "id": 37,
                "configurationVolume": "5.00",
                "configurationQuality": "0.00",
                "concentration": ""
            }
        ],
        "isType": "1"
    },
    "experimentLog2": {
        "dissolveForm": {
            "temperature": "70",
            "maxTemperature": "120",
            "time": "60",
            "speed": "",
            "rjSpeed": "700",
            "seal": "1",
            "environment": "1",
            "dissolve": "1",
            "dissolveTimeInterval": "30",
            "dissolveMaximum": "2",
            "dissolveGradient": "5"
        },
        "postProcessingForm": {
            "time": "30",
            "temperature": "40",
            "maxTemperature": "",
            "speed": "800",
            "method": "1",
            "precipitate": "1",
            "precoolingTemperature": "20"
        },
        "filtrationForm": {
            "type": "硅胶",
            "time": "50"
        }
    },
    "experimentLog3": {
        "dissolveForm": {},
        "postProcessingForm": {},
        "filtrationForm": {}
    }
}
    data_96_2 = {
    "id": "ad7ace5d207e4cc29f0f18fa8583ccf8",
    "condition": {
        "experimentLog1": {
            "SubstratesTableList": [
                {
                    "limitReagents": False,
                    "equivalent": "1.00",
                    "smiles": "sodium tetrahydroborate",
                    "substance": "试剂1、sodium tetrahydroborate",
                    "type": "reagent",
                    "liquidReagentConcentration": "固态",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_1",
                    "reactionQuality": "3.53",
                    "cas": "",
                    "molecularFormula": "NaBH4",
                    "singleCockAddVolume": "",
                    "reactionMoleConcentration": "",
                    "solvent": "37.83",
                    "reactionMoles": "0.09",
                    "id": 1
                },
                {
                    "limitReagents": False,
                    "smiles": "",
                    "substance": "溶剂1、",
                    "type": "solvent",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_2",
                    "cas": "",
                    "molecularFormula": "",
                    "singleCockAddVolume": 0.18671,
                    "reactionMoleConcentration": "",
                    "solvent": "0.00",
                    "id": 1
                },
                {
                    "limitReagents": True,
                    "equivalent": "1.00",
                    "smiles": "NCc1ccccc1",
                    "substance": "NCc1ccccc1",
                    "type": "sub",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_3",
                    "reactionQuality": "10.00",
                    "cas": "",
                    "molecularFormula": "C7H9N",
                    "singleCockAddVolume": 0.06223,
                    "reactionMoleConcentration": "",
                    "solvent": "107.15",
                    "reactionMoles": "0.09",
                    "id": 1
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.00",
                    "smiles": "Fc1ccccn1",
                    "substance": "Fc1ccccn1",
                    "type": "sub",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_4",
                    "reactionQuality": "9.06",
                    "cas": "",
                    "molecularFormula": "C5H4FN",
                    "singleCockAddVolume": 0.06223,
                    "reactionMoleConcentration": "",
                    "solvent": "97.09",
                    "reactionMoles": "0.09",
                    "id": 1
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.00",
                    "smiles": "diethyl 2,6-dimethyl-1,4-dihydropyridine-3,5-dicarboxylate",
                    "substance": "试剂1、diethyl 2,6-dimethyl-1,4-dihydropyridine-3,5-dicarboxylate",
                    "type": "reagent",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_5",
                    "reactionQuality": "26.07",
                    "cas": "",
                    "molecularFormula": "C15H21NO4",
                    "singleCockAddVolume": 0.06223,
                    "reactionMoleConcentration": "",
                    "solvent": "279.33",
                    "reactionMoles": "0.09",
                    "id": 2
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.00",
                    "smiles": "silver trifluoromethanesulfonate",
                    "substance": "试剂2、silver trifluoromethanesulfonate",
                    "type": "reagent",
                    "liquidReagentConcentration": "固态",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_6",
                    "reactionQuality": "23.99",
                    "cas": "",
                    "molecularFormula": "AgCF3O3S",
                    "singleCockAddVolume": "",
                    "reactionMoleConcentration": "",
                    "solvent": "256.94",
                    "reactionMoles": "0.09",
                    "id": 2
                },
                {
                    "limitReagents": False,
                    "smiles": "dichloromethane",
                    "substance": "溶剂1、dichloromethane",
                    "type": "solvent",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_7",
                    "cas": "",
                    "molecularFormula": "CH2Cl2",
                    "singleCockAddVolume": 0.12448,
                    "reactionMoleConcentration": "",
                    "solvent": "84.93",
                    "id": 2
                },
                {
                    "limitReagents": True,
                    "equivalent": "1.00",
                    "smiles": "NCc1ccccc1",
                    "substance": "NCc1ccccc1",
                    "type": "sub",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_8",
                    "reactionQuality": "10.00",
                    "cas": "",
                    "molecularFormula": "C7H9N",
                    "singleCockAddVolume": 0.06223,
                    "reactionMoleConcentration": "",
                    "solvent": "107.15",
                    "reactionMoles": "0.09",
                    "id": 2
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.00",
                    "smiles": "Fc1ccccn1",
                    "substance": "Fc1ccccn1",
                    "type": "sub",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_9",
                    "reactionQuality": "9.06",
                    "cas": "",
                    "molecularFormula": "C5H4FN",
                    "singleCockAddVolume": 0.06223,
                    "reactionMoleConcentration": "",
                    "solvent": "97.09",
                    "reactionMoles": "0.09",
                    "id": 2
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.00",
                    "smiles": "1.1 wt% Pd/NiO",
                    "substance": "试剂1、1.1 wt% Pd/NiO",
                    "type": "reagent",
                    "liquidReagentConcentration": "固态",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_10",
                    "reactionQuality": "0.00",
                    "cas": "",
                    "molecularFormula": "Pd/NiO",
                    "singleCockAddVolume": "",
                    "reactionMoleConcentration": "",
                    "solvent": "0.00",
                    "reactionMoles": "0.09",
                    "id": 3
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.00",
                    "smiles": "hydrogen",
                    "substance": "试剂2、hydrogen",
                    "type": "reagent",
                    "liquidReagentConcentration": "气体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_11",
                    "reactionQuality": "0.19",
                    "cas": "",
                    "molecularFormula": "H2",
                    "singleCockAddVolume": "",
                    "reactionMoleConcentration": "",
                    "solvent": "2.02",
                    "reactionMoles": "0.09",
                    "id": 3
                },
                {
                    "limitReagents": False,
                    "smiles": "ethanol",
                    "substance": "溶剂1、ethanol",
                    "type": "solvent",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_12",
                    "cas": "",
                    "molecularFormula": "C2H6O",
                    "singleCockAddVolume": 0.12448,
                    "reactionMoleConcentration": "",
                    "solvent": "46.07",
                    "id": 3
                },
                {
                    "limitReagents": True,
                    "equivalent": "1.00",
                    "smiles": "NCc1ccccc1",
                    "substance": "NCc1ccccc1",
                    "type": "sub",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_13",
                    "reactionQuality": "10.00",
                    "cas": "",
                    "molecularFormula": "C7H9N",
                    "singleCockAddVolume": 0.06223,
                    "reactionMoleConcentration": "",
                    "solvent": "107.15",
                    "reactionMoles": "0.09",
                    "id": 3
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.00",
                    "smiles": "Fc1ccccn1",
                    "substance": "Fc1ccccn1",
                    "type": "sub",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_14",
                    "reactionQuality": "9.06",
                    "cas": "",
                    "molecularFormula": "C5H4FN",
                    "singleCockAddVolume": 0.06223,
                    "reactionMoleConcentration": "",
                    "solvent": "97.09",
                    "reactionMoles": "0.09",
                    "id": 3
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.00",
                    "smiles": "sodium tetrahydroborate",
                    "substance": "试剂1、sodium tetrahydroborate",
                    "type": "reagent",
                    "liquidReagentConcentration": "固态",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_15",
                    "reactionQuality": "3.53",
                    "cas": "",
                    "molecularFormula": "NaBH4",
                    "singleCockAddVolume": "",
                    "reactionMoleConcentration": "",
                    "solvent": "37.83",
                    "reactionMoles": "0.09",
                    "id": 4
                },
                {
                    "limitReagents": False,
                    "smiles": "water",
                    "substance": "溶剂1、water",
                    "type": "solvent",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_16",
                    "cas": "",
                    "molecularFormula": "H2O",
                    "singleCockAddVolume": 0.18671,
                    "reactionMoleConcentration": "",
                    "solvent": "18.02",
                    "id": 4
                },
                {
                    "limitReagents": True,
                    "equivalent": "1.00",
                    "smiles": "NCc1ccccc1",
                    "substance": "NCc1ccccc1",
                    "type": "sub",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_17",
                    "reactionQuality": "10.00",
                    "cas": "",
                    "molecularFormula": "C7H9N",
                    "singleCockAddVolume": 0.06223,
                    "reactionMoleConcentration": "",
                    "solvent": "107.15",
                    "reactionMoles": "0.09",
                    "id": 4
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.00",
                    "smiles": "Fc1ccccn1",
                    "substance": "Fc1ccccn1",
                    "type": "sub",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_18",
                    "reactionQuality": "9.06",
                    "cas": "",
                    "molecularFormula": "C5H4FN",
                    "singleCockAddVolume": 0.06223,
                    "reactionMoleConcentration": "",
                    "solvent": "97.09",
                    "reactionMoles": "0.09",
                    "id": 4
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.20",
                    "smiles": "diethyl 2,6-dimethyl-1,4-dihydropyridine-3,5-dicarboxylate",
                    "substance": "试剂1、diethyl 2,6-dimethyl-1,4-dihydropyridine-3,5-dicarboxylate",
                    "type": "reagent",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_19",
                    "reactionQuality": "31.29",
                    "cas": "",
                    "molecularFormula": "C15H21NO4",
                    "singleCockAddVolume": 0.07468,
                    "reactionMoleConcentration": "",
                    "solvent": "279.33",
                    "reactionMoles": "0.11",
                    "id": 5
                },
                {
                    "limitReagents": False,
                    "smiles": "toluene",
                    "substance": "溶剂1、toluene",
                    "type": "solvent",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_20",
                    "cas": "",
                    "molecularFormula": "C7H8",
                    "singleCockAddVolume": 0.09958,
                    "reactionMoleConcentration": "",
                    "solvent": "92.14",
                    "id": 5
                },
                {
                    "limitReagents": True,
                    "equivalent": "1.00",
                    "smiles": "NCc1ccccc1",
                    "substance": "NCc1ccccc1",
                    "type": "sub",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_21",
                    "reactionQuality": "10.00",
                    "cas": "",
                    "molecularFormula": "C7H9N",
                    "singleCockAddVolume": 0.06223,
                    "reactionMoleConcentration": "",
                    "solvent": "107.15",
                    "reactionMoles": "0.09",
                    "id": 5
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.20",
                    "smiles": "Fc1ccccn1",
                    "substance": "Fc1ccccn1",
                    "type": "sub",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_22",
                    "reactionQuality": "10.88",
                    "cas": "",
                    "molecularFormula": "C5H4FN",
                    "singleCockAddVolume": 0.07468,
                    "reactionMoleConcentration": "",
                    "solvent": "97.09",
                    "reactionMoles": "0.11",
                    "id": 5
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.20",
                    "smiles": "5A molecular sieve",
                    "substance": "试剂1、5A molecular sieve",
                    "type": "reagent",
                    "liquidReagentConcentration": "固态",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_23",
                    "reactionQuality": "",
                    "cas": "",
                    "molecularFormula": "",
                    "singleCockAddVolume": "",
                    "reactionMoleConcentration": "",
                    "solvent": "",
                    "reactionMoles": "0.11",
                    "id": 6
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.20",
                    "smiles": "diethyl 2,6-dimethyl-1,4-dihydropyridine-3,5-dicarboxylate",
                    "substance": "试剂2、diethyl 2,6-dimethyl-1,4-dihydropyridine-3,5-dicarboxylate",
                    "type": "reagent",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_24",
                    "reactionQuality": "31.29",
                    "cas": "",
                    "molecularFormula": "C15H21NO4",
                    "singleCockAddVolume": 0.07468,
                    "reactionMoleConcentration": "",
                    "solvent": "279.33",
                    "reactionMoles": "0.11",
                    "id": 6
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.20",
                    "smiles": "thiourea",
                    "substance": "试剂3、thiourea",
                    "type": "reagent",
                    "liquidReagentConcentration": "固态",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_25",
                    "reactionQuality": "8.53",
                    "cas": "",
                    "molecularFormula": "CH4N2S",
                    "singleCockAddVolume": "",
                    "reactionMoleConcentration": "",
                    "solvent": "76.12",
                    "reactionMoles": "0.11",
                    "id": 6
                },
                {
                    "limitReagents": False,
                    "smiles": "toluene",
                    "substance": "溶剂1、toluene",
                    "type": "solvent",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_26",
                    "cas": "",
                    "molecularFormula": "C7H8",
                    "singleCockAddVolume": 0.09958,
                    "reactionMoleConcentration": "",
                    "solvent": "92.14",
                    "id": 6
                },
                {
                    "limitReagents": True,
                    "equivalent": "1.00",
                    "smiles": "NCc1ccccc1",
                    "substance": "NCc1ccccc1",
                    "type": "sub",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_27",
                    "reactionQuality": "10.00",
                    "cas": "",
                    "molecularFormula": "C7H9N",
                    "singleCockAddVolume": 0.06223,
                    "reactionMoleConcentration": "",
                    "solvent": "107.15",
                    "reactionMoles": "0.09",
                    "id": 6
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.20",
                    "smiles": "Fc1ccccn1",
                    "substance": "Fc1ccccn1",
                    "type": "sub",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_28",
                    "reactionQuality": "10.88",
                    "cas": "",
                    "molecularFormula": "C5H4FN",
                    "singleCockAddVolume": 0.07468,
                    "reactionMoleConcentration": "",
                    "solvent": "97.09",
                    "reactionMoles": "0.11",
                    "id": 6
                },
                {
                    "limitReagents": False,
                    "equivalent": "0.05",
                    "smiles": "Pd(OAc)2",
                    "substance": "试剂1、Pd(OAc)2",
                    "type": "reagent",
                    "liquidReagentConcentration": "固态",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_29",
                    "reactionQuality": "1.05",
                    "cas": "",
                    "molecularFormula": "C4H6O4Pd",
                    "singleCockAddVolume": "",
                    "reactionMoleConcentration": "",
                    "solvent": "224.51",
                    "reactionMoles": "0.00",
                    "id": 7
                },
                {
                    "limitReagents": False,
                    "equivalent": "0.05",
                    "smiles": "Josiphos SL-J009-1",
                    "substance": "试剂2、Josiphos SL-J009-1",
                    "type": "reagent",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_30",
                    "reactionQuality": "3.28",
                    "cas": "",
                    "molecularFormula": "C43H44FeP2",
                    "singleCockAddVolume": 0.00311,
                    "reactionMoleConcentration": "",
                    "solvent": "702.59",
                    "reactionMoles": "0.00",
                    "id": 7
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.50",
                    "smiles": "NaOtBu",
                    "substance": "试剂3、NaOtBu",
                    "type": "reagent",
                    "liquidReagentConcentration": "固态",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_31",
                    "reactionQuality": "13.46",
                    "cas": "",
                    "molecularFormula": "C4H9NaO",
                    "singleCockAddVolume": "",
                    "reactionMoleConcentration": "",
                    "solvent": "96.11",
                    "reactionMoles": "0.14",
                    "id": 7
                },
                {
                    "limitReagents": False,
                    "smiles": "",
                    "substance": "溶剂1、GDME",
                    "type": "solvent",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_32",
                    "cas": "",
                    "molecularFormula": "C6H14O3",
                    "singleCockAddVolume": 0.17157,
                    "reactionMoleConcentration": "",
                    "solvent": "134.17",
                    "id": 7
                },
                {
                    "limitReagents": True,
                    "equivalent": "1.00",
                    "smiles": "NCc1ccccc1",
                    "substance": "NCc1ccccc1",
                    "type": "sub",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_33",
                    "reactionQuality": "10.00",
                    "cas": "",
                    "molecularFormula": "C7H9N",
                    "singleCockAddVolume": 0.06223,
                    "reactionMoleConcentration": "",
                    "solvent": "107.15",
                    "reactionMoles": "0.09",
                    "id": 7
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.20",
                    "smiles": "Fc1ccccn1",
                    "substance": "Fc1ccccn1",
                    "type": "sub",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_34",
                    "reactionQuality": "10.88",
                    "cas": "",
                    "molecularFormula": "C5H4FN",
                    "singleCockAddVolume": 0.07468,
                    "reactionMoleConcentration": "",
                    "solvent": "97.09",
                    "reactionMoles": "0.11",
                    "id": 7
                },
                {
                    "limitReagents": False,
                    "equivalent": "0.03",
                    "smiles": "Pd2(dba)3",
                    "substance": "试剂1、Pd2(dba)3",
                    "type": "reagent",
                    "liquidReagentConcentration": "固态",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_35",
                    "reactionQuality": "2.14",
                    "cas": "",
                    "molecularFormula": "C54H45O3Pd2",
                    "singleCockAddVolume": "",
                    "reactionMoleConcentration": "",
                    "solvent": "915.72",
                    "reactionMoles": "0.00",
                    "id": 8
                },
                {
                    "limitReagents": False,
                    "equivalent": "0.05",
                    "smiles": "",
                    "substance": "试剂2、BIDIME",
                    "type": "reagent",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_36",
                    "reactionQuality": "1.34",
                    "cas": "",
                    "molecularFormula": "C20H18N2",
                    "singleCockAddVolume": 0.00311,
                    "reactionMoleConcentration": "",
                    "solvent": "286.37",
                    "reactionMoles": "0.00",
                    "id": 8
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.50",
                    "smiles": "NaOtBu",
                    "substance": "试剂3、NaOtBu",
                    "type": "reagent",
                    "liquidReagentConcentration": "固态",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_37",
                    "reactionQuality": "13.46",
                    "cas": "",
                    "molecularFormula": "C4H9NaO",
                    "singleCockAddVolume": "",
                    "reactionMoleConcentration": "",
                    "solvent": "96.11",
                    "reactionMoles": "0.14",
                    "id": 8
                },
                {
                    "limitReagents": False,
                    "smiles": "toluene",
                    "substance": "溶剂1、toluene",
                    "type": "solvent",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_38",
                    "cas": "",
                    "molecularFormula": "C7H8",
                    "singleCockAddVolume": 0.09958,
                    "reactionMoleConcentration": "",
                    "solvent": "92.14",
                    "id": 8
                },
                {
                    "limitReagents": True,
                    "equivalent": "1.00",
                    "smiles": "NCc1ccccc1",
                    "substance": "NCc1ccccc1",
                    "type": "sub",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_39",
                    "reactionQuality": "10.00",
                    "cas": "",
                    "molecularFormula": "C7H9N",
                    "singleCockAddVolume": 0.06223,
                    "reactionMoleConcentration": "",
                    "solvent": "107.15",
                    "reactionMoles": "0.09",
                    "id": 8
                },
                {
                    "limitReagents": False,
                    "equivalent": "1.20",
                    "smiles": "Fc1ccccn1",
                    "substance": "Fc1ccccn1",
                    "type": "sub",
                    "liquidReagentConcentration": "液体",
                    "reactionVolume": 0.31117,
                    "rowId": "row_1768549764507_40",
                    "reactionQuality": "10.88",
                    "cas": "",
                    "molecularFormula": "C5H4FN",
                    "singleCockAddVolume": 0.07468,
                    "reactionMoleConcentration": "",
                    "solvent": "97.09",
                    "reactionMoles": "0.11",
                    "id": 8
                }
            ],
            "wellPlates": "96孔板1ml",
            "MaterialsTableList": [
                {
                    "type": "reagent",
                    "substance": "试剂1、sodium tetrahydroborate",
                    "smiles": "sodium tetrahydroborate",
                    "solvent": "37.83",
                    "molecularFormula": "NaBH4",
                    "theoreticalMoles": "0.18",
                    "theoreticalQuality": "7.06",
                    "theoreticalVolume": "0.00",
                    "configurationMoles": "0.24",
                    "cas": "",
                    "id": 1,
                    "configurationVolume": "0.00",
                    "configurationQuality": "9.18",
                    "concentration": ""
                },
                {
                    "type": "solvent",
                    "substance": "溶剂1、",
                    "smiles": "",
                    "solvent": "0.00",
                    "molecularFormula": "",
                    "theoreticalMoles": "0.00",
                    "theoreticalQuality": "0.00",
                    "theoreticalVolume": "0.19",
                    "configurationMoles": "0.00",
                    "cas": "",
                    "id": 2,
                    "configurationVolume": "5.00",
                    "configurationQuality": "0.00",
                    "concentration": ""
                },
                {
                    "type": "sub",
                    "substance": "NCc1ccccc1",
                    "smiles": "NCc1ccccc1",
                    "solvent": "107.15",
                    "molecularFormula": "C7H9N",
                    "theoreticalMoles": "0.72",
                    "theoreticalQuality": "80.00",
                    "theoreticalVolume": "0.48",
                    "configurationMoles": "0.96",
                    "cas": "",
                    "id": 3,
                    "configurationVolume": "5.00",
                    "configurationQuality": "104.00",
                    "concentration": "1.50"
                },
                {
                    "type": "sub",
                    "substance": "Fc1ccccn1",
                    "smiles": "Fc1ccccn1",
                    "solvent": "97.09",
                    "molecularFormula": "C5H4FN",
                    "theoreticalMoles": "0.80",
                    "theoreticalQuality": "79.76",
                    "theoreticalVolume": "0.52",
                    "configurationMoles": "1.04",
                    "cas": "",
                    "id": 4,
                    "configurationVolume": "5.00",
                    "configurationQuality": "103.69",
                    "concentration": "1.54"
                },
                {
                    "type": "reagent",
                    "substance": "试剂1、diethyl 2,6-dimethyl-1,4-dihydropyridine-3,5-dicarboxylate",
                    "smiles": "diethyl 2,6-dimethyl-1,4-dihydropyridine-3,5-dicarboxylate",
                    "solvent": "279.33",
                    "molecularFormula": "C15H21NO4",
                    "theoreticalMoles": "0.20",
                    "theoreticalQuality": "57.36",
                    "theoreticalVolume": "0.13",
                    "configurationMoles": "0.26",
                    "cas": "",
                    "id": 5,
                    "configurationVolume": "5.00",
                    "configurationQuality": "74.57",
                    "concentration": "1.54"
                },
                {
                    "type": "reagent",
                    "substance": "试剂2、silver trifluoromethanesulfonate",
                    "smiles": "silver trifluoromethanesulfonate",
                    "solvent": "256.94",
                    "molecularFormula": "AgCF3O3S",
                    "theoreticalMoles": "0.09",
                    "theoreticalQuality": "23.99",
                    "theoreticalVolume": "0.00",
                    "configurationMoles": "0.12",
                    "cas": "",
                    "id": 6,
                    "configurationVolume": "0.00",
                    "configurationQuality": "31.19",
                    "concentration": ""
                },
                {
                    "type": "solvent",
                    "substance": "溶剂1、dichloromethane",
                    "smiles": "dichloromethane",
                    "solvent": "84.93",
                    "molecularFormula": "CH2Cl2",
                    "theoreticalMoles": "0.00",
                    "theoreticalQuality": "0.00",
                    "theoreticalVolume": "0.12",
                    "configurationMoles": "0.00",
                    "cas": "",
                    "id": 7,
                    "configurationVolume": "5.00",
                    "configurationQuality": "0.00",
                    "concentration": ""
                },
                {
                    "type": "reagent",
                    "substance": "试剂1、1.1 wt% Pd/NiO",
                    "smiles": "1.1 wt% Pd/NiO",
                    "solvent": "0.00",
                    "molecularFormula": "Pd/NiO",
                    "theoreticalMoles": "0.09",
                    "theoreticalQuality": "0.00",
                    "theoreticalVolume": "0.00",
                    "configurationMoles": "0.12",
                    "cas": "",
                    "id": 8,
                    "configurationVolume": "0.00",
                    "configurationQuality": "0.00",
                    "concentration": ""
                },
                {
                    "type": "reagent",
                    "substance": "试剂2、hydrogen",
                    "smiles": "hydrogen",
                    "solvent": "2.02",
                    "molecularFormula": "H2",
                    "theoreticalMoles": "0.09",
                    "theoreticalQuality": "0.19",
                    "theoreticalVolume": "0.00",
                    "configurationMoles": "0.12",
                    "cas": "",
                    "id": 9,
                    "configurationVolume": "0.00",
                    "configurationQuality": "0.25",
                    "concentration": ""
                },
                {
                    "type": "solvent",
                    "substance": "溶剂1、ethanol",
                    "smiles": "ethanol",
                    "solvent": "46.07",
                    "molecularFormula": "C2H6O",
                    "theoreticalMoles": "0.00",
                    "theoreticalQuality": "0.00",
                    "theoreticalVolume": "0.12",
                    "configurationMoles": "0.00",
                    "cas": "",
                    "id": 10,
                    "configurationVolume": "5.00",
                    "configurationQuality": "0.00",
                    "concentration": ""
                },
                {
                    "type": "solvent",
                    "substance": "溶剂1、water",
                    "smiles": "water",
                    "solvent": "18.02",
                    "molecularFormula": "H2O",
                    "theoreticalMoles": "0.00",
                    "theoreticalQuality": "0.00",
                    "theoreticalVolume": "0.19",
                    "configurationMoles": "0.00",
                    "cas": "",
                    "id": 11,
                    "configurationVolume": "5.00",
                    "configurationQuality": "0.00",
                    "concentration": ""
                },
                {
                    "type": "solvent",
                    "substance": "溶剂1、toluene",
                    "smiles": "toluene",
                    "solvent": "92.14",
                    "molecularFormula": "C7H8",
                    "theoreticalMoles": "0.00",
                    "theoreticalQuality": "0.00",
                    "theoreticalVolume": "0.30",
                    "configurationMoles": "0.00",
                    "cas": "",
                    "id": 12,
                    "configurationVolume": "5.00",
                    "configurationQuality": "0.00",
                    "concentration": ""
                },
                {
                    "type": "reagent",
                    "substance": "试剂1、5A molecular sieve",
                    "smiles": "5A molecular sieve",
                    "solvent": "",
                    "molecularFormula": "",
                    "theoreticalMoles": "0.11",
                    "theoreticalQuality": "0.00",
                    "theoreticalVolume": "0.00",
                    "configurationMoles": "0.14",
                    "cas": "",
                    "id": 13,
                    "configurationVolume": "0.00",
                    "configurationQuality": "0.00",
                    "concentration": ""
                },
                {
                    "type": "reagent",
                    "substance": "试剂2、diethyl 2,6-dimethyl-1,4-dihydropyridine-3,5-dicarboxylate",
                    "smiles": "diethyl 2,6-dimethyl-1,4-dihydropyridine-3,5-dicarboxylate",
                    "solvent": "279.33",
                    "molecularFormula": "C15H21NO4",
                    "theoreticalMoles": "0.11",
                    "theoreticalQuality": "31.29",
                    "theoreticalVolume": "0.07",
                    "configurationMoles": "0.14",
                    "cas": "",
                    "id": 14,
                    "configurationVolume": "5.00",
                    "configurationQuality": "40.68",
                    "concentration": "1.57"
                },
                {
                    "type": "reagent",
                    "substance": "试剂3、thiourea",
                    "smiles": "thiourea",
                    "solvent": "76.12",
                    "molecularFormula": "CH4N2S",
                    "theoreticalMoles": "0.11",
                    "theoreticalQuality": "8.53",
                    "theoreticalVolume": "0.00",
                    "configurationMoles": "0.14",
                    "cas": "",
                    "id": 15,
                    "configurationVolume": "0.00",
                    "configurationQuality": "11.09",
                    "concentration": ""
                },
                {
                    "type": "reagent",
                    "substance": "试剂1、Pd(OAc)2",
                    "smiles": "Pd(OAc)2",
                    "solvent": "224.51",
                    "molecularFormula": "C4H6O4Pd",
                    "theoreticalMoles": "0.00",
                    "theoreticalQuality": "1.05",
                    "theoreticalVolume": "0.00",
                    "configurationMoles": "0.00",
                    "cas": "",
                    "id": 16,
                    "configurationVolume": "0.00",
                    "configurationQuality": "1.37",
                    "concentration": ""
                },
                {
                    "type": "reagent",
                    "substance": "试剂2、Josiphos SL-J009-1",
                    "smiles": "Josiphos SL-J009-1",
                    "solvent": "702.59",
                    "molecularFormula": "C43H44FeP2",
                    "theoreticalMoles": "0.00",
                    "theoreticalQuality": "3.28",
                    "theoreticalVolume": "0.00",
                    "configurationMoles": "0.00",
                    "cas": "",
                    "id": 17,
                    "configurationVolume": "0.00",
                    "configurationQuality": "4.26",
                    "concentration": ""
                },
                {
                    "type": "reagent",
                    "substance": "试剂3、NaOtBu",
                    "smiles": "NaOtBu",
                    "solvent": "96.11",
                    "molecularFormula": "C4H9NaO",
                    "theoreticalMoles": "0.28",
                    "theoreticalQuality": "26.92",
                    "theoreticalVolume": "0.00",
                    "configurationMoles": "0.36",
                    "cas": "",
                    "id": 18,
                    "configurationVolume": "0.00",
                    "configurationQuality": "35.00",
                    "concentration": ""
                },
                {
                    "type": "solvent",
                    "substance": "溶剂1、GDME",
                    "smiles": "",
                    "solvent": "134.17",
                    "molecularFormula": "C6H14O3",
                    "theoreticalMoles": "0.00",
                    "theoreticalQuality": "0.00",
                    "theoreticalVolume": "0.17",
                    "configurationMoles": "0.00",
                    "cas": "",
                    "id": 19,
                    "configurationVolume": "5.00",
                    "configurationQuality": "0.00",
                    "concentration": ""
                },
                {
                    "type": "reagent",
                    "substance": "试剂1、Pd2(dba)3",
                    "smiles": "Pd2(dba)3",
                    "solvent": "915.72",
                    "molecularFormula": "C54H45O3Pd2",
                    "theoreticalMoles": "0.00",
                    "theoreticalQuality": "2.14",
                    "theoreticalVolume": "0.00",
                    "configurationMoles": "0.00",
                    "cas": "",
                    "id": 20,
                    "configurationVolume": "0.00",
                    "configurationQuality": "2.78",
                    "concentration": ""
                },
                {
                    "type": "reagent",
                    "substance": "试剂2、BIDIME",
                    "smiles": "",
                    "solvent": "286.37",
                    "molecularFormula": "C20H18N2",
                    "theoreticalMoles": "0.00",
                    "theoreticalQuality": "1.34",
                    "theoreticalVolume": "0.00",
                    "configurationMoles": "0.00",
                    "cas": "",
                    "id": 21,
                    "configurationVolume": "0.00",
                    "configurationQuality": "1.74",
                    "concentration": ""
                }
            ]
        },
        "experimentLog2": {
            "ReactionConditions": {
                "environment": "1",
                "temperature": "40",
                "seal": "1",
                "time": "8",
                "speed": "900"
            }
        },
        "experimentLog3": {
            "quenchingAgentData": [
                {
                    "quenchVolume": "19.20",
                    "quencherQuality": "592.00",
                    "name": "Ammonium chloride(aq)",
                    "addVolume": "582.00",
                    "molecularWeight": "251.60",
                    "moles": "",
                    "quencherConcentration": "52.40"
                }
            ],
            "productDilutionData": [
                {
                    "name": "MeCN",
                    "afterConcentration": "3.73 × 10⁻⁸",
                    "procedure": {
                        "diluent5": "10.00",
                        "diluent4": "10.00",
                        "diluent3": "190.00",
                        "diluent2": "1800.00",
                        "diluent1": "190.00"
                    },
                    "beforeConcentration": "0.30"
                }
            ],
            "internalStandardData": [
                {
                    "internalConcentration": "0.00",
                    "name": "BENZANILIDE",
                    "addVolume": "190.00",
                    "molecularWeight": "197.23",
                    "internalVolume": "28240.00",
                    "moles": "0.09",
                    "internalQuality": "2736.52"
                }
            ],
            "filtrationData": [
                {
                    "name": "0.22微米",
                    "filterTime": "3"
                }
            ],
            "extractionData": [],
            "coolingTime": "10"
        },
        "experimentLog4": {}
    }
}
    data_96_111 ={
    "experimentLog1": {
        "wellPlates": "96孔板1ml",
        "SubstratesTableList": [
            {
                "limitReagents": False,
                "equivalent": "0.05",
                "smiles": "Pd(OAc)2",
                "substance": "试剂1,Pd(OAc)2",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_1",
                "intensity": None,
                "reactionQuality": "0.52",
                "cas": "",
                "molecularFormula": "C4H6O4Pd",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "224.51",
                "reactionMoles": "0.00",
                "id": 1
            },
            {
                "limitReagents": False,
                "equivalent": "0.05",
                "smiles": "Josiphos SL-J009-1",
                "substance": "试剂2,Josiphos SL-J009-1",
                "type": "reagent",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_2",
                "intensity": None,
                "reactionQuality": "1.64",
                "cas": "",
                "molecularFormula": "C44H42FeP2",
                "singleCockAddVolume": "0.00",
                "reactionMoleConcentration": "",
                "solvent": "704.58",
                "reactionMoles": "0.00",
                "id": 1
            },
            {
                "limitReagents": False,
                "equivalent": "2.00",
                "smiles": "NaOtBu",
                "substance": "试剂3,NaOtBu",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_3",
                "intensity": None,
                "reactionQuality": "8.97",
                "cas": "",
                "molecularFormula": "C4H9ONa",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "96.11",
                "reactionMoles": "0.09",
                "id": 1
            },
            {
                "limitReagents": False,
                "equivalent": None,
                "smiles": "",
                "substance": "溶剂1,GDME",
                "type": "solvent",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_4",
                "intensity": None,
                "reactionQuality": None,
                "cas": "",
                "molecularFormula": "C6H14O3",
                "singleCockAddVolume": "0.09",
                "reactionMoleConcentration": "",
                "solvent": "134.17",
                "reactionMoles": None,
                "id": 1
            },
            {
                "limitReagents": True,
                "equivalent": "1.00",
                "smiles": "NCc1ccccc1",
                "substance": "NCc1ccccc1",
                "type": "sub",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_5",
                "intensity": None,
                "reactionQuality": "5.00",
                "cas": "",
                "molecularFormula": "C7H9N",
                "singleCockAddVolume": "0.03",
                "reactionMoleConcentration": "",
                "solvent": "107.15",
                "reactionMoles": "0.05",
                "id": 1
            },
            {
                "limitReagents": False,
                "equivalent": "1.50",
                "smiles": "Fc1ccccn1",
                "substance": "Fc1ccccn1",
                "type": "sub",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_6",
                "intensity": None,
                "reactionQuality": "10.30",
                "cas": "",
                "molecularFormula": "C9H6FN",
                "singleCockAddVolume": "0.04",
                "reactionMoleConcentration": "1.75",
                "solvent": "147.15",
                "reactionMoles": "0.07",
                "id": 1
            },
            {
                "limitReagents": False,
                "equivalent": "0.03",
                "smiles": "Pd2(dba)3",
                "substance": "试剂1,Pd2(dba)3",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_7",
                "intensity": None,
                "reactionQuality": "1.07",
                "cas": "",
                "molecularFormula": "C48H42O3Pd2",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "915.68",
                "reactionMoles": "0.00",
                "id": 2
            },
            {
                "limitReagents": False,
                "equivalent": "0.05",
                "smiles": "",
                "substance": "试剂2,BIDIME",
                "type": "reagent",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_8",
                "intensity": None,
                "reactionQuality": "0.67",
                "cas": "",
                "molecularFormula": "C20H16O2",
                "singleCockAddVolume": "0.00",
                "reactionMoleConcentration": "",
                "solvent": "288.34",
                "reactionMoles": "0.00",
                "id": 2
            },
            {
                "limitReagents": False,
                "equivalent": "2.00",
                "smiles": "NaOtBu",
                "substance": "试剂3,NaOtBu",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_9",
                "intensity": None,
                "reactionQuality": "8.97",
                "cas": "",
                "molecularFormula": "C4H9ONa",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "96.11",
                "reactionMoles": "0.09",
                "id": 2
            },
            {
                "limitReagents": False,
                "equivalent": None,
                "smiles": "toluene",
                "substance": "溶剂1,toluene",
                "type": "solvent",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_10",
                "intensity": None,
                "reactionQuality": None,
                "cas": "",
                "molecularFormula": "C7H8",
                "singleCockAddVolume": "0.08",
                "reactionMoleConcentration": "",
                "solvent": "92.14",
                "reactionMoles": None,
                "id": 2
            },
            {
                "limitReagents": True,
                "equivalent": "1.00",
                "smiles": "NCc1ccccc1",
                "substance": "NCc1ccccc1",
                "type": "sub",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_11",
                "intensity": None,
                "reactionQuality": "5.00",
                "cas": "",
                "molecularFormula": "C7H9N",
                "singleCockAddVolume": "0.03",
                "reactionMoleConcentration": "",
                "solvent": "107.15",
                "reactionMoles": "0.05",
                "id": 2
            },
            {
                "limitReagents": False,
                "equivalent": "1.50",
                "smiles": "Fc1ccccn1",
                "substance": "Fc1ccccn1",
                "type": "sub",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_12",
                "intensity": None,
                "reactionQuality": "10.30",
                "cas": "",
                "molecularFormula": "C9H6FN",
                "singleCockAddVolume": "0.05",
                "reactionMoleConcentration": "",
                "solvent": "147.15",
                "reactionMoles": "0.07",
                "id": 2
            },
            {
                "limitReagents": False,
                "equivalent": "0.05",
                "smiles": "Pd(OAc)2",
                "substance": "试剂1,Pd(OAc)2",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_13",
                "intensity": None,
                "reactionQuality": "0.52",
                "cas": "",
                "molecularFormula": "C4H6O4Pd",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "224.51",
                "reactionMoles": "0.00",
                "id": 3
            },
            {
                "limitReagents": False,
                "equivalent": "0.05",
                "smiles": "Josiphos SL-J001-1",
                "substance": "试剂2,Josiphos SL-J001-1",
                "type": "reagent",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_14",
                "intensity": None,
                "reactionQuality": "1.64",
                "cas": "",
                "molecularFormula": "C44H42FeP2",
                "singleCockAddVolume": "0.00",
                "reactionMoleConcentration": "",
                "solvent": "704.58",
                "reactionMoles": "0.00",
                "id": 3
            },
            {
                "limitReagents": False,
                "equivalent": "2.00",
                "smiles": "NaOtBu",
                "substance": "试剂3,NaOtBu",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_15",
                "intensity": None,
                "reactionQuality": "8.97",
                "cas": "",
                "molecularFormula": "C4H9ONa",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "96.11",
                "reactionMoles": "0.09",
                "id": 3
            },
            {
                "limitReagents": False,
                "equivalent": None,
                "smiles": "",
                "substance": "溶剂1,GDME",
                "type": "solvent",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_16",
                "intensity": None,
                "reactionQuality": None,
                "cas": "",
                "molecularFormula": "C6H14O3",
                "singleCockAddVolume": "0.08",
                "reactionMoleConcentration": "",
                "solvent": "134.17",
                "reactionMoles": None,
                "id": 3
            },
            {
                "limitReagents": True,
                "equivalent": "1.00",
                "smiles": "NCc1ccccc1",
                "substance": "NCc1ccccc1",
                "type": "sub",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_17",
                "intensity": None,
                "reactionQuality": "5.00",
                "cas": "",
                "molecularFormula": "C7H9N",
                "singleCockAddVolume": "0.03",
                "reactionMoleConcentration": "",
                "solvent": "107.15",
                "reactionMoles": "0.05",
                "id": 3
            },
            {
                "limitReagents": False,
                "equivalent": "1.50",
                "smiles": "Fc1ccccn1",
                "substance": "Fc1ccccn1",
                "type": "sub",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_18",
                "intensity": None,
                "reactionQuality": "10.30",
                "cas": "",
                "molecularFormula": "C9H6FN",
                "singleCockAddVolume": "0.05",
                "reactionMoleConcentration": "",
                "solvent": "147.15",
                "reactionMoles": "0.07",
                "id": 3
            },
            {
                "limitReagents": False,
                "equivalent": "0.05",
                "smiles": "C22H26Cl2NPPd",
                "substance": "试剂1,C22H26Cl2NPPd",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_19",
                "intensity": None,
                "reactionQuality": "1.17",
                "cas": "",
                "molecularFormula": "C22H26Cl2NPPd",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "500.73",
                "reactionMoles": "0.00",
                "id": 4
            },
            {
                "limitReagents": False,
                "equivalent": "2.00",
                "smiles": "NaOtBu",
                "substance": "试剂2,NaOtBu",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_20",
                "intensity": None,
                "reactionQuality": "8.97",
                "cas": "",
                "molecularFormula": "C4H9ONa",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "96.11",
                "reactionMoles": "0.09",
                "id": 4
            },
            {
                "limitReagents": False,
                "equivalent": None,
                "smiles": "toluene",
                "substance": "溶剂1,toluene",
                "type": "solvent",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_21",
                "intensity": None,
                "reactionQuality": None,
                "cas": "",
                "molecularFormula": "C7H8",
                "singleCockAddVolume": "0.08",
                "reactionMoleConcentration": "",
                "solvent": "92.14",
                "reactionMoles": None,
                "id": 4
            },
            {
                "limitReagents": True,
                "equivalent": "1.00",
                "smiles": "NCc1ccccc1",
                "substance": "NCc1ccccc1",
                "type": "sub",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_22",
                "intensity": None,
                "reactionQuality": "5.00",
                "cas": "",
                "molecularFormula": "C7H9N",
                "singleCockAddVolume": "0.03",
                "reactionMoleConcentration": "",
                "solvent": "107.15",
                "reactionMoles": "0.05",
                "id": 4
            },
            {
                "limitReagents": False,
                "equivalent": "1.50",
                "smiles": "Fc1ccccn1",
                "substance": "Fc1ccccn1",
                "type": "sub",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_23",
                "intensity": None,
                "reactionQuality": "10.30",
                "cas": "",
                "molecularFormula": "C9H6FN",
                "singleCockAddVolume": "0.05",
                "reactionMoleConcentration": "",
                "solvent": "147.15",
                "reactionMoles": "0.07",
                "id": 4
            },
            {
                "limitReagents": False,
                "equivalent": "0.05",
                "smiles": "Pd(dba)2",
                "substance": "试剂1,Pd(dba)2",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_24",
                "intensity": None,
                "reactionQuality": "1.07",
                "cas": "",
                "molecularFormula": "C26H22O2Pd",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "456.86",
                "reactionMoles": "0.00",
                "id": 5
            },
            {
                "limitReagents": False,
                "equivalent": "0.10",
                "smiles": "[tert-butyloxy]bis(N,N-diisopropylamino)phosphane",
                "substance": "试剂2,[tert-butyloxy]bis(N,N-diisopropylamino)phosphane",
                "type": "reagent",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_25",
                "intensity": None,
                "reactionQuality": "1.49",
                "cas": "",
                "molecularFormula": "C17H39N2OP",
                "singleCockAddVolume": "0.00",
                "reactionMoleConcentration": "",
                "solvent": "318.48",
                "reactionMoles": "0.00",
                "id": 5
            },
            {
                "limitReagents": False,
                "equivalent": "2.00",
                "smiles": "NaOtBu",
                "substance": "试剂3,NaOtBu",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_26",
                "intensity": None,
                "reactionQuality": "8.97",
                "cas": "",
                "molecularFormula": "C4H9NaO",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "96.11",
                "reactionMoles": "0.09",
                "id": 5
            },
            {
                "limitReagents": False,
                "equivalent": None,
                "smiles": "toluene",
                "substance": "溶剂1,toluene",
                "type": "solvent",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_27",
                "intensity": None,
                "reactionQuality": None,
                "cas": "",
                "molecularFormula": "C7H8",
                "singleCockAddVolume": "0.06",
                "reactionMoleConcentration": "",
                "solvent": "92.14",
                "reactionMoles": None,
                "id": 5
            },
            {
                "limitReagents": True,
                "equivalent": "1.00",
                "smiles": "NCc1ccccc1",
                "substance": "NCc1ccccc1",
                "type": "sub",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_28",
                "intensity": None,
                "reactionQuality": "5.00",
                "cas": "",
                "molecularFormula": "C7H9N",
                "singleCockAddVolume": "0.03",
                "reactionMoleConcentration": "",
                "solvent": "107.15",
                "reactionMoles": "0.05",
                "id": 5
            },
            {
                "limitReagents": False,
                "equivalent": "1.00",
                "smiles": "Fc1ccccn1",
                "substance": "Fc1ccccn1",
                "type": "sub",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_29",
                "intensity": None,
                "reactionQuality": "4.53",
                "cas": "",
                "molecularFormula": "C5H4FN",
                "singleCockAddVolume": "0.03",
                "reactionMoleConcentration": "",
                "solvent": "97.09",
                "reactionMoles": "0.05",
                "id": 5
            },
            {
                "limitReagents": False,
                "equivalent": "2.00",
                "smiles": "potassium tert-butylate",
                "substance": "试剂1,potassium tert-butylate",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_30",
                "intensity": None,
                "reactionQuality": "10.48",
                "cas": "",
                "molecularFormula": "C4H9KO",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "112.21",
                "reactionMoles": "0.09",
                "id": 6
            },
            {
                "limitReagents": False,
                "equivalent": "0.10",
                "smiles": "copper diacetate",
                "substance": "试剂2,copper diacetate",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_31",
                "intensity": None,
                "reactionQuality": "0.85",
                "cas": "",
                "molecularFormula": "C4H6CuO4",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "181.63",
                "reactionMoles": "0.00",
                "id": 6
            },
            {
                "limitReagents": False,
                "equivalent": None,
                "smiles": "1,4-dioxane",
                "substance": "溶剂1,1,4-dioxane",
                "type": "solvent",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_32",
                "intensity": None,
                "reactionQuality": None,
                "cas": "",
                "molecularFormula": "C4H8O2",
                "singleCockAddVolume": "0.06",
                "reactionMoleConcentration": "",
                "solvent": "88.11",
                "reactionMoles": None,
                "id": 6
            },
            {
                "limitReagents": True,
                "equivalent": "1.00",
                "smiles": "NCc1ccccc1",
                "substance": "NCc1ccccc1",
                "type": "sub",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_33",
                "intensity": None,
                "reactionQuality": "5.00",
                "cas": "",
                "molecularFormula": "C7H9N",
                "singleCockAddVolume": "0.03",
                "reactionMoleConcentration": "",
                "solvent": "107.15",
                "reactionMoles": "0.05",
                "id": 6
            },
            {
                "limitReagents": False,
                "equivalent": "1.00",
                "smiles": "Fc1ccccn1",
                "substance": "Fc1ccccn1",
                "type": "sub",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_34",
                "intensity": None,
                "reactionQuality": "4.53",
                "cas": "",
                "molecularFormula": "C5H4FN",
                "singleCockAddVolume": "0.03",
                "reactionMoleConcentration": "",
                "solvent": "97.09",
                "reactionMoles": "0.05",
                "id": 6
            },
            {
                "limitReagents": False,
                "equivalent": "0.05",
                "smiles": "bis[dichloro(pentamethylcyclopentadienyl)iridium(III)]",
                "substance": "试剂1,bis[dichloro(pentamethylcyclopentadienyl)iridium(III)]",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_35",
                "intensity": None,
                "reactionQuality": "1.86",
                "cas": "",
                "molecularFormula": "C20H30Cl4Ir2",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "796.78",
                "reactionMoles": "0.00",
                "id": 7
            },
            {
                "limitReagents": False,
                "equivalent": "0.10",
                "smiles": "2-(2-(diphenylphosphanyl)phenyl)benzo[d]oxazole",
                "substance": "试剂2,2-(2-(diphenylphosphanyl)phenyl)benzo[d]oxazole",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_36",
                "intensity": None,
                "reactionQuality": "1.77",
                "cas": "",
                "molecularFormula": "C25H18NOP",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "379.39",
                "reactionMoles": "0.00",
                "id": 7
            },
            {
                "limitReagents": False,
                "equivalent": "2.00",
                "smiles": "potassium hydroxide",
                "substance": "试剂3,potassium hydroxide",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_37",
                "intensity": None,
                "reactionQuality": "5.24",
                "cas": "",
                "molecularFormula": "KOH",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "56.11",
                "reactionMoles": "0.09",
                "id": 7
            },
            {
                "limitReagents": False,
                "equivalent": None,
                "smiles": "toluene",
                "substance": "溶剂1,toluene",
                "type": "solvent",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_38",
                "intensity": None,
                "reactionQuality": None,
                "cas": "",
                "molecularFormula": "C7H8",
                "singleCockAddVolume": "0.06",
                "reactionMoleConcentration": "",
                "solvent": "92.14",
                "reactionMoles": None,
                "id": 7
            },
            {
                "limitReagents": True,
                "equivalent": "1.00",
                "smiles": "NCc1ccccc1",
                "substance": "NCc1ccccc1",
                "type": "sub",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_39",
                "intensity": None,
                "reactionQuality": "5.00",
                "cas": "",
                "molecularFormula": "C7H9N",
                "singleCockAddVolume": "0.03",
                "reactionMoleConcentration": "",
                "solvent": "107.15",
                "reactionMoles": "0.05",
                "id": 7
            },
            {
                "limitReagents": False,
                "equivalent": "1.00",
                "smiles": "Fc1ccccn1",
                "substance": "Fc1ccccn1",
                "type": "sub",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_40",
                "intensity": None,
                "reactionQuality": "4.53",
                "cas": "",
                "molecularFormula": "C5H4FN",
                "singleCockAddVolume": "0.03",
                "reactionMoleConcentration": "",
                "solvent": "97.09",
                "reactionMoles": "0.05",
                "id": 7
            },
            {
                "limitReagents": False,
                "equivalent": "0.05",
                "smiles": "chloro(1,5-cyclooctadiene)rhodium(I) dimer",
                "substance": "试剂1,chloro(1,5-cyclooctadiene)rhodium(I) dimer",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_41",
                "intensity": None,
                "reactionQuality": "1.15",
                "cas": "",
                "molecularFormula": "C16H24Cl2Rh2",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "493.98",
                "reactionMoles": "0.00",
                "id": 8
            },
            {
                "limitReagents": False,
                "equivalent": "0.10",
                "smiles": "1-mesityl-3-((1-(4-(trifluoromethyl)phenyl)-1H-1,2,3-triazol-4-yl)methyl)-1H-imidazol-3-ium bromide",
                "substance": "试剂2,1-mesityl-3-((1-(4-(trifluoromethyl)phenyl)-1H-1,2,3-triazol-4-yl)methyl)-1H-imidazol-3-ium bromide",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_42",
                "intensity": None,
                "reactionQuality": "2.48",
                "cas": "",
                "molecularFormula": "C24H22BrF3N6",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "531.37",
                "reactionMoles": "0.00",
                "id": 8
            },
            {
                "limitReagents": False,
                "equivalent": "2.00",
                "smiles": "potassium tert-butylate",
                "substance": "试剂3,potassium tert-butylate",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_43",
                "intensity": None,
                "reactionQuality": "10.48",
                "cas": "",
                "molecularFormula": "C4H9KO",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "112.21",
                "reactionMoles": "0.09",
                "id": 8
            },
            {
                "limitReagents": False,
                "equivalent": "0.10",
                "smiles": "sodium tetrakis[(3,5-di-trifluoromethyl)phenyl]borate",
                "substance": "试剂4,sodium tetrakis[(3,5-di-trifluoromethyl)phenyl]borate",
                "type": "reagent",
                "liquidReagentConcentration": "固态",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_44",
                "intensity": None,
                "reactionQuality": "4.14",
                "cas": "",
                "molecularFormula": "C32H12BF24Na",
                "singleCockAddVolume": "",
                "reactionMoleConcentration": "",
                "solvent": "886.21",
                "reactionMoles": "0.00",
                "id": 8
            },
            {
                "limitReagents": False,
                "equivalent": None,
                "smiles": "(2)H8-toluene",
                "substance": "溶剂1,(2)H8-toluene",
                "type": "solvent",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_45",
                "intensity": None,
                "reactionQuality": None,
                "cas": "",
                "molecularFormula": "C7D8",
                "singleCockAddVolume": "0.06",
                "reactionMoleConcentration": "",
                "solvent": "100.19",
                "reactionMoles": None,
                "id": 8
            },
            {
                "limitReagents": True,
                "equivalent": "1.00",
                "smiles": "NCc1ccccc1",
                "substance": "NCc1ccccc1",
                "type": "sub",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_46",
                "intensity": None,
                "reactionQuality": "5.00",
                "cas": "",
                "molecularFormula": "C7H9N",
                "singleCockAddVolume": "0.03",
                "reactionMoleConcentration": "",
                "solvent": "107.15",
                "reactionMoles": "0.05",
                "id": 8
            },
            {
                "limitReagents": False,
                "equivalent": "1.00",
                "smiles": "Fc1ccccn1",
                "substance": "Fc1ccccn1",
                "type": "sub",
                "liquidReagentConcentration": "液体",
                "reactionVolume": "0.16",
                "rowId": "row_1775807379310_47",
                "intensity": None,
                "reactionQuality": "4.53",
                "cas": "",
                "molecularFormula": "C5H4FN",
                "singleCockAddVolume": "0.03",
                "reactionMoleConcentration": "",
                "solvent": "97.09",
                "reactionMoles": "0.05",
                "id": 8
            }
        ],
        "MaterialsTableList": [
            {
                "type": "reagent",
                "substance": "试剂1,Pd(OAc)2",
                "smiles": "Pd(OAc)2",
                "solvent": "224.51",
                "molecularFormula": "C4H6O4Pd",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.04",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 1,
                "configurationVolume": "0.00",
                "configurationQuality": "1.35",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂2,Josiphos SL-J009-1",
                "smiles": "Josiphos SL-J009-1",
                "solvent": "704.58",
                "molecularFormula": "C44H42FeP2",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.64",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 2,
                "configurationVolume": "0.00",
                "configurationQuality": "2.13",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂3,NaOtBu",
                "smiles": "NaOtBu",
                "solvent": "96.11",
                "molecularFormula": "C4H9ONa",
                "theoreticalMoles": "0.36",
                "theoreticalQuality": "35.88",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.48",
                "cas": "",
                "intensity": None,
                "id": 3,
                "configurationVolume": "0.00",
                "configurationQuality": "46.64",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "溶剂1,GDME",
                "smiles": "",
                "solvent": "134.17",
                "molecularFormula": "C6H14O3",
                "theoreticalMoles": "",
                "theoreticalQuality": "",
                "theoreticalVolume": "0.17",
                "configurationMoles": "",
                "cas": "",
                "intensity": None,
                "id": 4,
                "configurationVolume": "5.00",
                "configurationQuality": "",
                "concentration": ""
            },
            {
                "type": "sub",
                "substance": "NCc1ccccc1",
                "smiles": "NCc1ccccc1",
                "solvent": "107.15",
                "molecularFormula": "C7H9N",
                "theoreticalMoles": "0.40",
                "theoreticalQuality": "40.00",
                "theoreticalVolume": "0.24",
                "configurationMoles": "0.53",
                "cas": "",
                "intensity": None,
                "id": 5,
                "configurationVolume": "5.00",
                "configurationQuality": "52.00",
                "concentration": "1.67"
            },
            {
                "type": "sub",
                "substance": "Fc1ccccn1",
                "smiles": "Fc1ccccn1",
                "solvent": "147.15",
                "molecularFormula": "C9H6FN",
                "theoreticalMoles": "0.48",
                "theoreticalQuality": "59.32",
                "theoreticalVolume": "0.31",
                "configurationMoles": "0.60",
                "cas": "",
                "intensity": None,
                "id": 6,
                "configurationVolume": "5.00",
                "configurationQuality": "77.12",
                "concentration": "1.55"
            },
            {
                "type": "reagent",
                "substance": "试剂1,Pd2(dba)3",
                "smiles": "Pd2(dba)3",
                "solvent": "915.68",
                "molecularFormula": "C48H42O3Pd2",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.07",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 7,
                "configurationVolume": "0.00",
                "configurationQuality": "1.39",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂2,BIDIME",
                "smiles": "",
                "solvent": "288.34",
                "molecularFormula": "C20H16O2",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.67",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 8,
                "configurationVolume": "0.00",
                "configurationQuality": "0.87",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "溶剂1,toluene",
                "smiles": "toluene",
                "solvent": "92.14",
                "molecularFormula": "C7H8",
                "theoreticalMoles": "",
                "theoreticalQuality": "",
                "theoreticalVolume": "0.28",
                "configurationMoles": "",
                "cas": "",
                "intensity": None,
                "id": 9,
                "configurationVolume": "5.00",
                "configurationQuality": "",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂2,Josiphos SL-J001-1",
                "smiles": "Josiphos SL-J001-1",
                "solvent": "704.58",
                "molecularFormula": "C44H42FeP2",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.64",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 10,
                "configurationVolume": "0.00",
                "configurationQuality": "2.13",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂1,C22H26Cl2NPPd",
                "smiles": "C22H26Cl2NPPd",
                "solvent": "500.73",
                "molecularFormula": "C22H26Cl2NPPd",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.17",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 11,
                "configurationVolume": "0.00",
                "configurationQuality": "1.52",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂2,NaOtBu",
                "smiles": "NaOtBu",
                "solvent": "96.11",
                "molecularFormula": "C4H9ONa",
                "theoreticalMoles": "0.09",
                "theoreticalQuality": "8.97",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.12",
                "cas": "",
                "intensity": None,
                "id": 12,
                "configurationVolume": "0.00",
                "configurationQuality": "11.66",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂1,Pd(dba)2",
                "smiles": "Pd(dba)2",
                "solvent": "456.86",
                "molecularFormula": "C26H22O2Pd",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.07",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 13,
                "configurationVolume": "0.00",
                "configurationQuality": "1.39",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂2,[tert-butyloxy]bis(N,N-diisopropylamino)phosphane",
                "smiles": "[tert-butyloxy]bis(N,N-diisopropylamino)phosphane",
                "solvent": "318.48",
                "molecularFormula": "C17H39N2OP",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.49",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 14,
                "configurationVolume": "0.00",
                "configurationQuality": "1.94",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂1,potassium tert-butylate",
                "smiles": "potassium tert-butylate",
                "solvent": "112.21",
                "molecularFormula": "C4H9KO",
                "theoreticalMoles": "0.09",
                "theoreticalQuality": "10.48",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.12",
                "cas": "",
                "intensity": None,
                "id": 15,
                "configurationVolume": "0.00",
                "configurationQuality": "13.62",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂2,copper diacetate",
                "smiles": "copper diacetate",
                "solvent": "181.63",
                "molecularFormula": "C4H6CuO4",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "0.85",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 16,
                "configurationVolume": "0.00",
                "configurationQuality": "1.10",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "溶剂1,1,4-dioxane",
                "smiles": "1,4-dioxane",
                "solvent": "88.11",
                "molecularFormula": "C4H8O2",
                "theoreticalMoles": "",
                "theoreticalQuality": "",
                "theoreticalVolume": "0.06",
                "configurationMoles": "",
                "cas": "",
                "intensity": None,
                "id": 17,
                "configurationVolume": "5.00",
                "configurationQuality": "",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂1,bis[dichloro(pentamethylcyclopentadienyl)iridium(III)]",
                "smiles": "bis[dichloro(pentamethylcyclopentadienyl)iridium(III)]",
                "solvent": "796.78",
                "molecularFormula": "C20H30Cl4Ir2",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.86",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 18,
                "configurationVolume": "0.00",
                "configurationQuality": "2.42",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂2,2-(2-(diphenylphosphanyl)phenyl)benzo[d]oxazole",
                "smiles": "2-(2-(diphenylphosphanyl)phenyl)benzo[d]oxazole",
                "solvent": "379.39",
                "molecularFormula": "C25H18NOP",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.77",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 19,
                "configurationVolume": "0.00",
                "configurationQuality": "2.30",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂3,potassium hydroxide",
                "smiles": "potassium hydroxide",
                "solvent": "56.11",
                "molecularFormula": "KOH",
                "theoreticalMoles": "0.09",
                "theoreticalQuality": "5.24",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.12",
                "cas": "",
                "intensity": None,
                "id": 20,
                "configurationVolume": "0.00",
                "configurationQuality": "6.81",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂1,chloro(1,5-cyclooctadiene)rhodium(I) dimer",
                "smiles": "chloro(1,5-cyclooctadiene)rhodium(I) dimer",
                "solvent": "493.98",
                "molecularFormula": "C16H24Cl2Rh2",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "1.15",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 21,
                "configurationVolume": "0.00",
                "configurationQuality": "1.49",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂2,1-mesityl-3-((1-(4-(trifluoromethyl)phenyl)-1H-1,2,3-triazol-4-yl)methyl)-1H-imidazol-3-ium bromide",
                "smiles": "1-mesityl-3-((1-(4-(trifluoromethyl)phenyl)-1H-1,2,3-triazol-4-yl)methyl)-1H-imidazol-3-ium bromide",
                "solvent": "531.37",
                "molecularFormula": "C24H22BrF3N6",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "2.48",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 22,
                "configurationVolume": "0.00",
                "configurationQuality": "3.22",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂3,potassium tert-butylate",
                "smiles": "potassium tert-butylate",
                "solvent": "112.21",
                "molecularFormula": "C4H9KO",
                "theoreticalMoles": "0.09",
                "theoreticalQuality": "10.48",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.12",
                "cas": "",
                "intensity": None,
                "id": 23,
                "configurationVolume": "0.00",
                "configurationQuality": "13.62",
                "concentration": ""
            },
            {
                "type": "reagent",
                "substance": "试剂4,sodium tetrakis[(3,5-di-trifluoromethyl)phenyl]borate",
                "smiles": "sodium tetrakis[(3,5-di-trifluoromethyl)phenyl]borate",
                "solvent": "886.21",
                "molecularFormula": "C32H12BF24Na",
                "theoreticalMoles": "0.00",
                "theoreticalQuality": "4.14",
                "theoreticalVolume": "0.00",
                "configurationMoles": "0.00",
                "cas": "",
                "intensity": None,
                "id": 24,
                "configurationVolume": "0.00",
                "configurationQuality": "5.38",
                "concentration": ""
            },
            {
                "type": "solvent",
                "substance": "溶剂1,(2)H8-toluene",
                "smiles": "(2)H8-toluene",
                "solvent": "100.19",
                "molecularFormula": "C7D8",
                "theoreticalMoles": "",
                "theoreticalQuality": "",
                "theoreticalVolume": "0.06",
                "configurationMoles": "",
                "cas": "",
                "intensity": None,
                "id": 25,
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
            "environment": "1",
            "temperature": "40",
            "seal": "1",
            "time": "8",
            "speed": "900"
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
        "quenchingAgentData": [],
        "productDilutionData": [
            {
                "name": "MeCN",
                "afterConcentration": "1",
                "procedure": {
                    "diluent5": "10.00",
                    "diluent4": "10.00",
                    "diluent3": "190.00",
                    "diluent2": "1800.00",
                    "diluent1": "190.00"
                },
                "beforeConcentration": "0.30"
            }
        ],
        "internalStandardData": [
            {
                "internalConcentration": "0.12",
                "name": "BENZANILIDE",
                "addVolume": "190.00",
                "molecularWeight": "197.23",
                "internalVolume": "28.24",
                "moles": "0.02",
                "internalQuality": "684.06"
            }
        ],
        "filtrationData": [
            {
                "name": "0.22微米",
                "filterTime": "3"
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
        "coolingTime": "10"
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
    main('1564',data96,1)
    # main('1564',data_96,2)


