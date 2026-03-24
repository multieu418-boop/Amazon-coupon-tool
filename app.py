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
        """处理用户输入的 ASIN 列表，输出分号分隔的字符串"""
        if not text: return ""
        # 兼容换行、逗号、分号、空格
        asins = re.split(r'[;,\s\n]+', str(text).strip())
        clean_list = [a.strip().upper() for a in asins if len(a.strip()) == 10]
        return ";".join(list(dict.fromkeys(clean_list)))

    @staticmethod
    def parse_template_config(file):
        """读取模板第 5/7/8/9 行逻辑"""
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        headers = [cell.value for cell in ws[7] if cell.value] # 第7行标题
        rules = {ws.cell(row=7, column=i).value: ws.cell(row=5, column=i).value 
                 for i in range(1, len(headers) + 1)} # 第5行规则映射
        
        # 提取下拉选项 (从第8/9行获取示例值，实际开发可扩展为读取数据有效性列表)
        options = {}
        for i in range(1, len(headers) + 1):
            val8 = ws.cell(row=8, column=i).value
            val9 = ws.cell(row=9, column=i).value
            if val8 or val9:
                options[ws.cell(row=7, column=i).value] = list(filter(None, [val8, val9]))
        
        return headers, rules, options, wb

# --- Streamlit UI 配置 ---
st.set_page_config(page_title="Amazon Coupon 集成工具", layout="wide")
st.title("🚀 Amazon Coupon 自动化提报系统")

# 侧边栏：必备文件上传
with st.sidebar:
    st.header("📂 必备文件上传")
    all_listing = st.file_uploader("1. ALL Listing Report", type=['txt', 'csv'])
    coupon_template = st.file_uploader("2. 空白 Coupon 模板", type=['xlsx'])
    
    if all_listing: st.success("✅ All Listing 已就绪")
    if coupon_template: st.success("✅ 模板已解析")

tab1, tab2 = st.tabs(["🔵 阶段 1：生成提报", "🔴 阶段 2：报错解析修复"])

# --- 阶段 1：提报生成 ---
with tab1:
    if not coupon_template:
        st.info("请在左侧上传 Coupon 模板以开始。")
    else:
        headers, rules, options, wb_template = AmazonCouponTool.parse_template_config(coupon_template)
        
        # 动态表单生成
        st.subheader("📝 填写 Coupon 需求")
        with st.form("coupon_form"):
            user_data = {}
            col1, col2 = st.columns(2)
            
            for i, header in enumerate(headers):
                target_col = col1 if i % 2 == 0 else col2
                help_text = rules.get(header, "")
                
                # ASIN 特殊处理
                if "ASIN" in str(header).upper():
                    raw_asin = target_col.text_area(f"{header}", help=help_text, placeholder="粘贴 ASIN 列表...")
                    user_data[header] = AmazonCouponTool.clean_asin_input(raw_asin)
                # 下拉框处理
                elif header in options:
                    user_data[header] = target_col.selectbox(f"{header}", options[header], help=help_text)
                # 普通输入
                else:
                    user_data[header] = target_col.text_input(f"{header}", help=help_text)
            
            submit = st.form_submit_button("生成并导出提报文件")
            
            if submit:
                # 写入逻辑：克隆样式并追加到第10行
                ws = wb_template.active
                target_row = 10
                # 寻找真正的空白行
                while ws.cell(row=target_row, column=1).value:
                    target_row += 1
                
                for col_idx, header in enumerate(headers, 1):
                    new_cell = ws.cell(row=target_row, column=col_idx, value=user_data[header])
                    # 克隆第9行的样式
                    source_cell = ws.cell(row=9, column=col_idx)
                    if source_cell.has_style:
                        new_cell.font = copy(source_cell.font)
                        new_cell.border = copy(source_cell.border)
                        new_cell.fill = copy(source_cell.fill)
                        new_cell.alignment = copy(source_cell.alignment)
                
                output = BytesIO()
                wb_template.save(output)
                st.download_button("📥 点击下载生成的提报文件", data=output.getvalue(), file_name="Coupon_Upload_Ready.xlsx")

# --- 阶段 2：报错修复 (简化版逻辑展示) ---
with tab2:
    error_file = st.file_uploader("上传亚马逊报错回传文件", type=['xlsx'])
    if error_file and all_listing:
        st.warning("⚠️ 正在开发：此处将自动解析批注并匹配 All Listing 价格...")
        st.write("已识别报错 ASIN，计算中...")