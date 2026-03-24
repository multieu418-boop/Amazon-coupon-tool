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
        return ";".join(list(dict.fromkeys(clean_list))) # 去重并连接

    @staticmethod
    def get_template_info(file):
        """解析模板：提取标题(Row 7)、规则(Row 5)、示例/下拉(Row 8/9)"""
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        
        # 提取第7行所有标题名
        headers = []
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=7, column=col).value
            if val:
                headers.append(val)
        
        # 提取第5行规则提示
        rules = {ws.cell(row=7, column=i).value: ws.cell(row=5, column=i).value 
                 for i in range(1, len(headers) + 1)}
        
        # 提取第8/9行作为下拉选项 (针对你指定的字段)
        options = {}
        dropdown_fields = ["折扣类型", "限制每位买家只能兑换一次", "优惠券类型", "目标买家", "叠加使用的促销"]
        
        for i in range(1, len(headers) + 1):
            header_name = ws.cell(row=7, column=i).value
            if any(field in str(header_name) for field in dropdown_fields):
                # 获取第8行和第9行的值作为下拉选项，保留原样（中文/英文）
                v8 = ws.cell(row=8, column=i).value
                v9 = ws.cell(row=9, column=i).value
                opt_list = list(dict.fromkeys(filter(None, [str(v8) if v8 else None, str(v9) if v9 else None])))
                options[header_name] = opt_list
                
        return headers, rules, options, wb

# --- Streamlit UI 界面 ---
st.set_page_config(page_title="Amazon Coupon Tool", layout="wide")
st.title("🚀 Amazon Coupon 自动化提报系统 (GitHub 版)")

# 侧边栏上传
with st.sidebar:
    st.header("📂 必备文件上传")
    all_listing = st.file_uploader("1. ALL Listing Report", type=['txt', 'csv'])
    coupon_template = st.file_uploader("2. 空白 Coupon 模板", type=['xlsx'])
    if coupon_template:
        st.success("✅ 模板解析成功")

tab1, tab2 = st.tabs(["🔵 阶段 1：生成提报", "🔴 阶段 2：报错解析修复"])

# --- 阶段 1：生成逻辑 ---
with tab1:
    if not coupon_template:
        st.info("请先在左侧上传 Amazon 原始模板文件。")
    else:
        headers, rules, options, wb_template = AmazonCouponTool.get_template_info(coupon_template)
        
        st.subheader("📝 填写 Coupon 需求内容")
        
        with st.form("coupon_input_form"):
            user_inputs = {}
            # 按照你指定的标题名进行逻辑分类
            col1, col2 = st.columns(2)
            
            for idx, name in enumerate(headers):
                target_col = col1 if idx % 2 == 0 else col2
                help_msg = rules.get(name, "参考模板要求填写")
                
                # 1. 自动处理 ASIN 列表 (文本域)
                if "ASIN 列表" in str(name):
                    raw_asin = target_col.text_area(f"📍 {name}", help=help_msg, placeholder="直接粘贴 ASIN，系统自动处理格式...")
                    user_inputs[name] = AmazonCouponTool.clean_asin_input(raw_asin)
                
                # 2. 下拉选择字段 (根据你的要求锁定这5个)
                elif any(field in str(name) for field in ["折扣类型", "限制每位买家只能兑换一次", "优惠券类型", "目标买家", "叠加使用的促销"]):
                    opts = options.get(name, ["请参考模板第8行"])
                    user_inputs[name] = target_col.selectbox(f"🔽 {name}", options=opts, help=help_msg)
                
                # 3. 其他手动输入字段 (数值、日期、名称等)
                else:
                    user_inputs[name] = target_col.text_input(f"✍️ {name}", help=help_msg)

            submit_btn = st.form_submit_button("🔥 生成标准上传文件")

            if submit_btn:
                ws = wb_template.active
                # 定位写入行：从第10行开始向下找空行
                write_row = 10
                while ws.cell(row=write_row, column=1).value:
                    write_row += 1
                
                # 写入并克隆样式
                for col_idx, header_name in enumerate(headers, 1):
                    val = user_inputs.get(header_name, "")
                    new_cell = ws.cell(row=write_row, column=col_idx, value=val)
                    
                    # 关键：完全保留第9行的原始样式
                    source_cell = ws.cell(row=9, column=col_idx)
                    if source_cell.has_style:
                        new_cell.font = copy(source_cell.font)
                        new_cell.border = copy(source_cell.border)
                        new_cell.fill = copy(source_cell.fill)
                        new_cell.alignment = copy(source_cell.alignment)
                        new_cell.number_format = source_cell.number_format
                
                # 导出文件
                output = BytesIO()
                wb_template.save(output)
                st.download_button(
                    label="📥 点击下载提报文件 (可以直接上传亚马逊)",
                    data=output.getvalue(),
                    file_name="Amazon_Coupon_Ready.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.balloons()

# --- 阶段 2：报错解析 (预留接口) ---
with tab2:
    st.header("解析亚马逊报错批注")
    error_excel = st.file_uploader("上传亚马逊退回的报错文件", type=['xlsx'])
    if error_excel and all_listing:
        st.info("系统正在读取批注中的 ASIN 和报错原因...")
        # 此处后续可根据具体的 DE/UK 报错文本进行正则匹配优化
