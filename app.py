import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
import re
from io import BytesIO
import math

# --- 页面基础配置 ---
st.set_page_config(page_title="Cupshe Coupon 综合工具", layout="wide")

class CouponProcessor:
    @staticmethod
    def clean_asin_input(text):
        if not text: return ""
        asins = re.split(r'[;,\s\n]+', str(text).strip())
        clean_list = [a.strip().upper() for a in asins if len(a.strip()) == 10]
        return ";".join(list(dict.fromkeys(clean_list)))

    @staticmethod
    def parse_error_details(comment_text):
        error_map = {}
        if not comment_text: return error_map
        blocks = re.split(r'([A-Z0-9]{10})\n', str(comment_text))
        if len(blocks) > 1:
            for i in range(1, len(blocks), 2):
                asin = blocks[i].strip()
                content = blocks[i+1]
                req_p_match = re.search(r'(?:要求的净价格|当前净价格|要求的最高商品价格)：[^\d]*([\d\.]+)', content)
                req_p = float(req_p_match.group(1)) if req_p_match else None
                reason_part = re.split(r'(?:要求的净价格|当前净价格|要求的最高商品价格)', content)[0]
                reason = reason_part.strip().replace('\n', ' ')
                # 核心自动化规则：没有验证价格的一律初始化为“剔除”
                auto_exclude = "没有经验证的参考价" in reason
                error_map[asin] = {
                    "req_price": req_p, 
                    "reason": reason, 
                    "default_decision": "剔除" if auto_exclude else "保留"
                }
        return error_map

# --- 侧边栏：文件上传 ---
with st.sidebar:
    st.header("📂 核心数据源")
    all_listing_file = st.file_uploader("1. All Listing 报告 (txt/csv)", type=['txt', 'csv', 'xlsx'])
    template_file = st.file_uploader("2. Coupon 模板 (xlsx)", type=['xlsx'])
    st.divider()
    if st.button("🔄 重置所有会话"):
        st.session_state.clear()
        st.rerun()

st.title("🎯 Cupshe Amazon Coupon 自动化管理系统")

# 定义 Tab
tab1, tab2 = st.tabs(["🔵 第一阶段：生成提报", "🔴 第二阶段：报错修复"])

# --- 第一阶段逻辑 ---
with tab1:
    if not template_file:
        st.info("请在左侧上传 Coupon 模板以开启表单。")
    else:
        # 解析模板：获取第7行标题，以及第8/9行的下拉选项
        wb_gen = openpyxl.load_workbook(template_file, data_only=True)
        ws_gen = wb_gen.active
        headers = [cell.value for cell in ws_gen[7] if cell.value]
        
        # 提取下拉选项逻辑
        dropdown_options = {}
        for i in range(1, len(headers) + 1):
            val8 = ws_gen.cell(row=8, column=i).value
            val9 = ws_gen.cell(row=9, column=i).value
            if val8 or val9:
                # 过滤掉重复项和空项，形成下拉列表
                opts = list(dict.fromkeys(filter(None, [str(val8), str(val9)])))
                dropdown_options[headers[i-1]] = opts

        st.subheader("📝 快速生成提报需求")
        with st.form("quick_gen_form"):
            user_data = {}
            col1, col2 = st.columns(2)
            for i, h in enumerate(headers):
                target_col = col1 if i % 2 == 0 else col2
                # 特殊处理 ASIN
                if "ASIN" in str(h).upper():
                    raw_asin = target_col.text_area(f"{h}", placeholder="支持换行、逗号或空格分隔...")
                    user_data[h] = CouponProcessor.clean_asin_input(raw_asin)
                # 处理下拉选项（从第8/9行提取出来的）
                elif h in dropdown_options:
                    user_data[h] = target_col.selectbox(f"{h}", options=dropdown_options[h])
                # 普通文本输入
                else:
                    user_data[h] = target_col.text_input(f"{h}")
            
            if st.form_submit_button("🚀 生成并导出提报文件"):
                target_row = 10
                # 写入逻辑并克隆第9行样式
                for col_idx, h in enumerate(headers, 1):
                    new_cell = ws_gen.cell(row=target_row, column=col_idx, value=user_data[h])
                    source_cell = ws_gen.cell(row=9, column=col_idx)
                    if source_cell.has_style:
                        new_cell.font, new_cell.border, new_cell.fill, new_cell.alignment = \
                            copy(source_cell.font), copy(source_cell.border), copy(source_cell.fill), copy(source_cell.alignment)
                
                out_gen = BytesIO()
                wb_gen.save(out_gen)
                st.download_button("📥 下载生成的文件", out_gen.getvalue(), "New_Coupon_Request.xlsx")

# --- 第二阶段逻辑 ---
with tab2:
    if not all_listing_file or not template_file:
        st.info("请确保左侧已上传 All Listing 报告及带报错批注的模板。")
    else:
        # 使用 Session State 确保多人并行不冲突
        if 'repair_data' not in st.session_state:
            with st.spinner("正在交叉解析数据..."):
                # 读取 Listing (自动识别编码)
                for enc in ['utf-8', 'utf-16', 'gbk', 'utf-8-sig']:
                    try:
                        all_listing_file.seek(0)
                        df_l = pd.read_csv(all_listing_file, sep='\t', encoding=enc) if all_listing_file.name.endswith('.txt') else pd.read_excel(all_listing_file)
                        df_l.columns = [c.lower().strip() for c in df_l.columns]
                        break
                    except: continue
                
                template_file.seek(0)
                wb_err = openpyxl.load_workbook(template_file, data_only=True)
                ws_err = wb_err.active
                err_headers = [cell.value for cell in ws_err[7]]
                
                # 定位核心列
                asin_col = next((c for c in df_l.columns if 'asin' in c), None)
                price_col = next((c for c in df_l.columns if 'price' in c or '价格' in c), None)
                e_asin_idx = next((i for i, h in enumerate(err_headers) if h and 'ASIN' in str(h)), 0)
                e_disc_idx = next((i for i, h in enumerate(err_headers) if h and '折扣' in str(h) and '数值' in str(h)), 2)

                repair_rows = []
                for r_idx in range(10, ws_err.max_row + 1):
                    asin_val = ws_err.cell(row=r_idx, column=e_asin_idx+1).value
                    if not asin_val: continue
                    
                    comment = ws_err.cell(row=r_idx, column=ws_err.max_column).comment
                    err_map = CouponProcessor.parse_error_details(comment.text if comment else "")
                    asins = [a.strip() for a in str(asin_val).replace(',', ';').split(';') if a.strip()]
                    
                    for a in asins:
                        p_match = df_l[df_l[asin_col] == a][price_col].values if asin_col else []
                        orig_p = p_match[0] if len(p_match) > 0 else 0
                        info = err_map.get(a, {})
                        
                        # 自动计算建议折扣
                        curr_d = ws_err.cell(row=r_idx, column=e_disc_idx+1).value or 0.05
                        suggested = curr_d
                        if a in err_map and orig_p and info.get('req_price'):
                            needed = math.ceil(((float(orig_p) - float(info.get('req_price'))) / float(orig_p)) * 100)
                            suggested = needed / 100 if curr_d < 1 else needed

                        repair_rows.append({
                            "决策": info.get('default_decision', "保留"),
                            "ASIN": a, "状态": "❌ 报错" if a in err_map else "✅ 正常",
                            "报错原因": info.get('reason', "-"), "拟提报折扣": suggested,
                            "Listing原价": orig_p, "原始行号": r_idx
                        })
                st.session_state.repair_data = pd.DataFrame(repair_rows)
                st.session_state.err_headers = err_headers

        if st.session_state.repair_data is not None:
            # 渲染决策台
            edited_df = st.data_editor(
                st.session_state.repair_data,
                column_config={
                    "决策": st.column_config.SelectboxColumn("决策", options=["保留", "剔除"]),
                    "原始行号": None, "拟提报折扣": st.column_config.NumberColumn("拟提报折扣", format="%.2f")
                },
                disabled=["ASIN", "状态", "报错原因", "Listing原价"],
                hide_index=True, use_container_width=True, key="repair_table"
            )

            if not edited_df.equals(st.session_state.repair_data):
                st.session_state.repair_data = edited_df
                st.rerun()

            if st.button("🚀 导出全行修复版 Excel", use_container_width=True):
                template_file.seek(0)
                wb_out = openpyxl.load_workbook(template_file)
                ws_out = wb_out.active
                
                # 备份数据 & 清空 10 行以后
                backup = {}
                for r in edited_df['原始行号'].unique():
                    backup[r] = [ws_out.cell(row=r, column=c).value for c in range(1, ws_out.max_column + 1)]
                
                for r in range(10, ws_out.max_row + 1):
                    for c in range(1, ws_out.max_column + 1): ws_out.cell(row=r, column=c).value = None

                # 回写逻辑
                final_keep = edited_df[edited_df['决策'] == "保留"]
                curr_row = 10
                a_col_idx, d_col_idx = 1, 3
                for i, h in enumerate(st.session_state.err_headers, 1):
                    if h and 'ASIN' in str(h): a_col_idx = i
                    if h and '折扣' in str(h) and '数值' in str(h): d_col_idx = i

                for (orig_line, disc), group in final_keep.groupby(['原始行号', '拟提报折扣']):
                    orig_vals = backup.get(orig_line)
                    for c_idx, val in enumerate(orig_vals, 1):
                        target_cell = ws_out.cell(row=curr_row, column=c_idx, value=val)
                        source_cell = ws_out.cell(row=orig_line, column=c_idx)
                        if source_cell.has_style:
                            target_cell.font, target_cell.border, target_cell.fill, target_cell.number_format, target_cell.alignment = \
                                copy(source_cell.font), copy(source_cell.border), copy(source_cell.fill), copy(source_cell.number_format), copy(source_cell.alignment)
                    
                    ws_out.cell(row=curr_row, column=a_col_idx).value = ";".join(group['ASIN'].tolist())
                    ws_out.cell(row=curr_row, column=d_col_idx).value = disc
                    curr_row += 1
                
                out_repair = BytesIO()
                wb_out.save(out_repair)
                st.success("文件生成成功！")
                st.download_button("📥 下载修复文件", out_repair.getvalue(), "Coupon_Fixed_Full.xlsx")
