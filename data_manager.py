import io
import math
import os
import re
from datetime import datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal
from typing import Any

import pandas as pd

MONEY_QUANT = Decimal("0.01")


def parse_money_decimal(val: Any) -> Decimal:
    """从单元格值解析为 Decimal（不单独 quantize，便于多格相加后再统一到分）。"""
    if val is None:
        return Decimal("0")
    if isinstance(val, Decimal):
        return val
    if isinstance(val, float) and pd.isna(val):
        return Decimal("0")
    if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
        return Decimal("0")
    if isinstance(val, bool):
        return Decimal(int(val))
    if isinstance(val, int):
        return Decimal(val)
    if isinstance(val, float):
        # 已参与过 float 运算的数：先规范到分再转 Decimal，避免二进制尾数污染累加
        return Decimal(str(round_money_half_up(val)))
    s = str(val).strip().replace("￥", "").replace("¥", "").replace(",", "").replace("，", "")
    if not s or s.lower() in ("nan", "none", "nat") or s in ("-", "—", "NA", "None"):
        return Decimal("0")
    try:
        return Decimal(s)
    except Exception:
        return Decimal("0")


def sum_parsed_money(values: Any) -> float:
    """对一列/可迭代单元格金额用 Decimal 求和，最后再四舍五入到分（与 Excel SUM 精度更一致）。"""
    if values is None:
        return 0.0
    total = Decimal("0")
    for x in values:
        total += parse_money_decimal(x)
    return round_money_half_up(total)


def round_money_half_up(val: Any) -> float:
    """金额四舍五入到分（ROUND_HALF_UP），用于解析、汇总、写回与界面。"""
    if val is None:
        return 0.0
    if isinstance(val, Decimal):
        try:
            return float(val.quantize(MONEY_QUANT, rounding=ROUND_HALF_UP))
        except Exception:
            return 0.0
    if isinstance(val, str):
        try:
            return round_money_half_up(parse_money_decimal(val))
        except Exception:
            return 0.0
    if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
        return 0.0
    if isinstance(val, bool):
        return float(int(val))
    if isinstance(val, (int, float)):
        try:
            d = Decimal(str(val))
        except Exception:
            return 0.0
        return float(d.quantize(MONEY_QUANT, rounding=ROUND_HALF_UP))
    s = str(val).strip().replace("￥", "").replace("¥", "").replace(",", "").replace("，", "")
    if not s or s.lower() in ("nan", "none", "nat") or s in ("-", "—", "NA", "None"):
        return 0.0
    try:
        d = Decimal(s)
    except Exception:
        return 0.0
    return float(d.quantize(MONEY_QUANT, rounding=ROUND_HALF_UP))

# 资产表标准列名：与 sheet_normalize / gui 一致；供应商各月可能用「截至YYYY年M月D日应计费天数」
BILLING_DAYS_COLUMN = "应计费天数"
_LEGACY_BILLING_DAY_HEADER = "截至2026年4月30日应计费天数"
_VENDOR_BILL_HEADER_RE = re.compile(r"^截至\d{4}年\d{1,2}月\d{1,2}日应计费天数$")
_LEGACY_USE_DAY_HEADER = "截至2026年4月30日使用天数"
_VENDOR_USE_DAYS_HEADER_RE = re.compile(r"^截至\d{4}年\d{1,2}月\d{1,2}日使用天数$")


def _merge_vendor_billing_columns(df: pd.DataFrame) -> pd.DataFrame:
    """把「截至…应计费天数」等供应商表头并入标准列「应计费天数」。"""
    if df is None or len(df.columns) == 0:
        return df
    vendors = []
    for c in list(df.columns):
        s = str(c).strip()
        if s == BILLING_DAYS_COLUMN:
            continue
        if s == _LEGACY_BILLING_DAY_HEADER or _VENDOR_BILL_HEADER_RE.fullmatch(s):
            vendors.append(c)
    if not vendors:
        return df
    if BILLING_DAYS_COLUMN not in df.columns:
        df = df.rename(columns={vendors[0]: BILLING_DAYS_COLUMN})
        vendors = vendors[1:]
    for c in vendors:
        if c not in df.columns:
            continue
        m = df[BILLING_DAYS_COLUMN].astype(str).str.strip() == ""
        df.loc[m, BILLING_DAYS_COLUMN] = df.loc[m, c].astype(str)
        df = df.drop(columns=[c])
    return df


def _merge_vendor_use_days_columns(df: pd.DataFrame) -> pd.DataFrame:
    """把「截至…使用天数」等供应商表头并入标准列「使用天数」（与 sheet_normalize 一致）。"""
    if df is None or len(df.columns) == 0:
        return df
    vendors = []
    for c in list(df.columns):
        s = str(c).strip()
        if s == "使用天数":
            continue
        if s == _LEGACY_USE_DAY_HEADER or _VENDOR_USE_DAYS_HEADER_RE.fullmatch(s):
            vendors.append(c)
    if not vendors:
        return df
    if "使用天数" not in df.columns:
        df = df.rename(columns={vendors[0]: "使用天数"})
        vendors = vendors[1:]
    for c in vendors:
        if c not in df.columns:
            continue
        m = df["使用天数"].astype(str).str.strip() == ""
        df.loc[m, "使用天数"] = df.loc[m, c].astype(str)
        df = df.drop(columns=[c])
    return df


def _merge_vendor_total_cost_columns(df: pd.DataFrame) -> pd.DataFrame:
    """把「本期合计费用」并入标准列「合计费用」（与 sheet_normalize 表头别名一致，避免 reindex 丢列）。"""
    if df is None or len(df.columns) == 0:
        return df
    if "本期合计费用" not in df.columns:
        return df
    if "合计费用" not in df.columns:
        return df.rename(columns={"本期合计费用": "合计费用"})

    def _empty_cost(s: pd.Series) -> pd.Series:
        t = s.astype(str).str.strip()
        return t.eq("") | t.str.lower().isin(("nan", "none", "nat"))

    m = _empty_cost(df["合计费用"])
    df.loc[m, "合计费用"] = df.loc[m, "本期合计费用"].astype(str)
    return df.drop(columns=["本期合计费用"])


class AssetDataManager:
    def __init__(self):
        self.asset_types = {
            'database': {
                'name': '数据库资产',
                'columns': [
                    '唯一编号', '服务目录编号', '服务类型', '服务名称', '数量', '单位', '单价',
                    '备注', '开通时间', '关停时间', '计费开始时间', '本期计费开始时间',
                    '本期计费截至日期', '本月核算天数', '使用天数', '本月核算金额',
                    BILLING_DAYS_COLUMN, '合计费用',
                ]
            }
        }
        self.current_type = 'database'
        self.data = pd.DataFrame(columns=self.asset_types[self.current_type]['columns'])
        self.load_data()
    
    def load_data(self):
        """加载数据"""
        filename = f"{self.asset_types[self.current_type]['name']}.csv"
        if os.path.exists(filename):
            self.data = pd.read_csv(filename, encoding='utf-8-sig')
        self.data = self.data.fillna("")
        self.data = _merge_vendor_billing_columns(self.data)
        self.data = _merge_vendor_use_days_columns(self.data)
        self.data = _merge_vendor_total_cost_columns(self.data)
        cols = self.asset_types[self.current_type]["columns"]
        for c in cols:
            if c not in self.data.columns:
                self.data[c] = ""
        self.data = self.data.reindex(columns=list(cols), fill_value="")
    
    def save_data(self):
        """保存数据"""
        filename = f"{self.asset_types[self.current_type]['name']}.csv"
        self.data.to_csv(filename, index=False, encoding='utf-8-sig')
        return filename
    
    def generate_unique_id(self):
        """生成唯一编号"""
        if len(self.data) == 0:
            return 'DB0001'
        max_id = max([int(str(row['唯一编号'])[2:]) for _, row in self.data.iterrows() if row['唯一编号']])
        return f"DB{str(max_id + 1).zfill(4)}"
    
    def add_asset(self, asset_info):
        """添加资产记录"""
        if '唯一编号' not in asset_info or not asset_info['唯一编号']:
            asset_info['唯一编号'] = self.generate_unique_id()
        
        new_row = pd.DataFrame([asset_info])
        self.data = pd.concat([self.data, new_row], ignore_index=True)
        self.save_data()
        return True
    
    def update_asset(self, index, asset_info):
        """更新资产记录"""
        if 0 <= index < len(self.data):
            for key, value in asset_info.items():
                if key in self.data.columns:
                    self.data.loc[index, key] = value
            self.save_data()
            return True
        return False
    
    def delete_asset(self, index):
        """删除资产记录"""
        if 0 <= index < len(self.data):
            self.data = self.data.drop(index).reset_index(drop=True)
            self.save_data()
            return True
        return False
    
    def clear_all_data(self):
        """清除所有数据"""
        self.data = pd.DataFrame(columns=self.asset_types[self.current_type]['columns'])
        self.save_data()
        return True
    
    def search_assets(self, conditions):
        """根据条件搜索资产"""
        result = self.data.copy()
        
        for key, value in conditions.items():
            if value and key in result.columns:
                if isinstance(result[key].iloc[0], str):
                    result = result[result[key].str.contains(str(value), case=False, na=False)]
                else:
                    result = result[result[key] == value]
        
        return result
    
    def get_statistics(self):
        """获取统计信息"""
        stats = {
            'total_count': len(self.data),
            'service_types': self.data['服务类型'].unique().tolist() if '服务类型' in self.data.columns else [],
            'service_names': self.data['服务名称'].unique().tolist() if '服务名称' in self.data.columns else []
        }
        
        if "合计费用" in self.data.columns:
            stats["total_cost"] = sum_parsed_money(self.data["合计费用"])

            stats["cost_by_type"] = {}
            for service_type in stats["service_types"]:
                type_data = self.data[self.data["服务类型"] == service_type]
                stats["cost_by_type"][service_type] = sum_parsed_money(type_data["合计费用"])
        
        return stats
    
    def calculate_billing_start_date(self, open_date_str):
        """计算计费开始时间（开通时间+90天）"""
        if not open_date_str:
            return ''
        try:
            # 尝试多种日期格式
            date_formats = ['%Y/%m/%d', '%Y-%m-%d', '%Y年%m月%d日', '%d/%m/%Y', '%d-%m-%Y']
            open_date = None
            for fmt in date_formats:
                try:
                    open_date = datetime.strptime(str(open_date_str), fmt)
                    break
                except:
                    continue
            
            if open_date:
                billing_start_date = open_date + timedelta(days=90)
                return billing_start_date.strftime('%Y/%m/%d')
            return ''
        except:
            return ''
    
    def _import_dataframe(self, df):
        """将已规范化的 DataFrame 合并进当前资产表（与 CSV 导入逻辑一致）。"""
        df = df.fillna("")
        df = _merge_vendor_billing_columns(df)
        df = _merge_vendor_use_days_columns(df)
        df = _merge_vendor_total_cost_columns(df)

        new_count = 0
        update_count = 0

        for _, row in df.iterrows():
            unique_id = str(row["唯一编号"]).strip() if "唯一编号" in df.columns else ""

            if "开通时间" in df.columns and "计费开始时间" in df.columns:
                open_date = row["开通时间"]
                billing_start = self.calculate_billing_start_date(open_date)
                row["计费开始时间"] = billing_start

            if unique_id and "唯一编号" in self.data.columns:
                existing_index = self.data[self.data["唯一编号"] == unique_id].index
                if len(existing_index) > 0:
                    for col in df.columns:
                        if col in self.data.columns:
                            self.data.loc[existing_index[0], col] = row[col]
                    update_count += 1
                    continue

            if not unique_id:
                unique_id = self.generate_unique_id()
                row["唯一编号"] = unique_id

            new_row = pd.DataFrame([row])
            self.data = pd.concat([self.data, new_row], ignore_index=True)
            new_count += 1

        self.save_data()
        return True, f"导入完成！新增 {new_count} 条，更新 {update_count} 条"

    def import_from_excel(self, excel_path):
        """从 Excel 导入：先读表再经 UTF-8 CSV 往返，与直接 read_excel 迭代相比更稳定。"""
        try:
            df = pd.read_excel(
                excel_path, sheet_name=0, dtype=str, keep_default_na=False
            )
            df = df.fillna("")
            buf = io.BytesIO()
            df.to_csv(buf, index=False, encoding="utf-8-sig")
            buf.seek(0)
            df = pd.read_csv(buf, encoding="utf-8-sig", dtype=str, keep_default_na=False)
            return self._import_dataframe(df)
        except Exception as e:
            return False, str(e)

    def import_from_csv(self, csv_path):
        """从CSV导入数据（支持多种编码）"""
        try:
            encodings = ["utf-8-sig", "gbk", "gb2312", "cp1252", "utf-8"]
            df = None

            for encoding in encodings:
                try:
                    df = pd.read_csv(csv_path, encoding=encoding)
                    break
                except UnicodeDecodeError:
                    continue

            if df is None:
                return False, "无法识别文件编码，请确保文件为UTF-8或GBK编码"

            return self._import_dataframe(df)
        except Exception as e:
            return False, str(e)
    
    def export_to_excel(self, excel_path):
        """导出数据到Excel"""
        try:
            self.data.to_excel(excel_path, index=False)
            return True, f"数据已成功导出到 {excel_path}"
        except Exception as e:
            return False, str(e)
    
    def export_csv_template(self, csv_path):
        """导出标准CSV模板（仅表头）"""
        try:
            columns = self.asset_types[self.current_type]['columns']
            df = pd.DataFrame(columns=columns)
            df.to_csv(csv_path, index=False, encoding='utf-8-sig')
            return True, f"CSV模板已成功导出到 {csv_path}"
        except Exception as e:
            return False, str(e)
