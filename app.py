import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
import re
from io import BytesIO

# --- 核心逻辑类 ---
class AmazonCouponTool:
    @staticmethod
    def clean_asin_input(text):
        """处理 ASIN 列表：自动识别换行/逗号/分号，转为分号连接"""
        if not text: return ""
        asins = re.split(r'[;,\s\n\t]+', str(text).strip())
        clean_list = [a.strip().upper() for a in asins if len(a.strip()) == 10 and a.upper().startswith('B')]
        return ";".join(list(dict.fromkeys(clean_list)))

    @staticmethod
    def get_template_info(file):
        """解析模板：提取标题(Row 7)、规则(Row 5)、示例/下拉(Row 8/9)"""
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        
        headers = []
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=7, column=col).value
            if val: headers.append(val)
        
        rules = {ws.cell(row=7, column=i).value: ws.cell(row=5, column=i).value 
                 for i in range(1, len(headers) + 1)}
        
        options = {}
        dropdown_fields = ["折扣类型", "限制每位买家只能兑换一次", "优惠券类型", "目标买家", "叠加使用的促销"]
        
        for i in range(1, len(headers) + 1):
            header_name = ws.cell(row=7, column=i).value
            if any(field in str(header_name) for field in dropdown_fields):
                v8 = ws.cell(row=8, column=i).value
                v9 = ws.cell(row=9, column=i).value
                opt_list = list(dict.fromkeys(filter(None, [str(v8) if v8 else None, str(v9) if v9 else None])))
                options[header_name] = opt_list if opt_list else ["无预设选项"]
                
        return headers, rules, options, wb

# --- Streamlit UI ---
st.set_page_config(page_title="Amazon Coupon Tool", layout="wide")
st.title("🚀 Amazon Coupon 自动化提报系统")

with st.sidebar:
    st.header("📂 必备文件上传")
    all_listing = st.file_uploader("1. ALL Listing Report", type=['txt', 'csv'])
    coupon_template = st.file_uploader("2. 空白 Coupon 模板", type=['xlsx'])

tab1, tab2 = st.tabs(["🔵 阶段 1：生成提报", "🔴 阶段 2：报错解析修复"])

if tab1:
    if not coupon_template:
        st.info("请先上传模板。")
    else:
        try:
            headers, rules, options, wb_template = AmazonCouponTool.get_template_info(coupon_template)
            
            with st.form("coupon_input_form"):
                user_inputs = {}
                col1, col2 = st.columns(2)
                for idx, name in enumerate(headers):
                    target_col = col1 if idx % 2 == 0 else col2
                    help_msg = rules.get(name, "")
                    if "ASIN 列表" in str(name):
                        raw_asin = target_col.text_area(f"📍 {name}", help=help_msg)
                        user_inputs[name] = AmazonCouponTool.clean_asin_input(raw_asin)
                    elif any(field in str(name) for field in ["折扣类型", "限制每位买家只能兑换一次", "优惠券类型", "目标买家", "叠加使用的促销"]):
                        opts = options.get(name, ["请参考模板"])
                        user_inputs[name] = target_col.selectbox(f"🔽 {name}", options=opts, help=help_msg)
                    else:
                        user_inputs[name] = target_col.text_input(f"✍️ {name}", help=help_msg)

                submit_btn = st.form_submit_button("🔥 生成标准上传文件")

                if submit_btn:
                    ws = wb_template.active
                    write_row = 10
                    while ws.cell(row=write_row, column=1).value:
                        write_row += 1
                    
                    # 修复后的样式克隆逻辑
                    for col_idx, header_name in enumerate(headers, 1):
                        val = user_inputs.get(header_name, "")
                        new_cell = ws.cell(row=write_row, column=col_idx, value=val)
                        source_cell = ws.cell(row=9, column=col_idx)
                        if source_cell:
                            try:
                                if source_cell.font: new_cell.font = copy(source_cell.font)
                                if source_cell.border: new_cell.border = copy(source_cell.border)
                                if source_cell.fill: new_cell.fill = copy(source_cell.fill)
                                if source_cell.alignment: new_cell.alignment = copy(source_cell.alignment)
                                if source_cell.number_format: new_cell.number_format = source_cell.number_format
                            except:
                                pass # 如果某项样式不支持克隆，则跳过，保证程序不崩
                    
                    output = BytesIO()
                    wb_template.save(output)
                    st.download_button(label="📥 下载提报文件", data=output.getvalue(), file_name="Coupon_Ready.xlsx")
                    st.balloons()
        except Exception as e:
            st.error(f"解析过程中出现问题: {e}")
