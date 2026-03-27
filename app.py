import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
import re
from io import BytesIO
import math
from datetime import datetime

# --- 页面基础配置 ---
st.set_page_config(page_title="Cupshe Amazon Coupon 综合管理系统", layout="wide")

# --- 核心逻辑类 ---
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
        # 匹配 ASIN 及其下方的报错内容
        blocks = re.split(r'([A-Z0-9]{10})\n', str(comment_text))
        if len(blocks) > 1:
            for i in range(1, len(blocks), 2):
                asin = blocks[i].strip()
                content = blocks[i+1]
                # 提取价格要求
                req_p_match = re.search(r'(?:要求的净价格|当前净价格|要求的最高商品价格)：[^\d]*([\d\.]+)', content)
                req_p = float(req_p_match.group(1)) if req_p_match else None
                # 提取报错原因
                reason_part = re.split(r'(?:要求的净价格|当前净价格|要求的最高商品价格)', content)[0]
                reason = reason_part.strip().replace('\n', ' ')
                auto_exclude = "没有经验证的参考价" in reason
                error_map[asin] = {
                    "req_price": req_p, 
                    "reason": reason, 
                    "default_decision": "剔除" if auto_exclude else "保留"
                }
        return error_map

# --- 侧边栏 ---
with st.sidebar:
    st.header("📂 核心数据源")
    all_listing_file = st.file_uploader("1. All Listing 报告", type=['txt', 'csv', 'xlsx'])
    template_file = st.file_uploader("2. Coupon 模板 (带报错批注)", type=['xlsx'])
    st.divider()
    
    # 仅在第二阶段显示的筛选器
    if template_file and all_listing_file:
        st.header("⚙️ 第二阶段筛选配置")
        status_sel = st.multiselect("ASIN 状态筛选", ["✅ 正常", "❌ 批注报错"], default=["✅ 正常", "❌ 批注报错"])
        discount_limit = st.slider("折扣力度预警线 (%)", 5, 50, 30) / 100
        reason_kw = st.text_input("报错原因关键词过滤")

    if st.button("🔄 重置所有会话"):
        st.session_state.clear()
        st.rerun()

st.title("🎯 Cupshe Amazon Coupon 自动化管理系统")

tab1, tab2 = st.tabs(["🔵 第一阶段：生成提报", "🔴 第二阶段：报错修复"])

# --- 第一阶段：生成提报 ---
with tab1:
    if not template_file:
        st.info("💡 请先上传模板文件。")
    else:
        wb_gen = openpyxl.load_workbook(template_file, data_only=True)
        ws_gen = wb_gen.active
        headers = [cell.value for cell in ws_gen[7] if cell.value]
        
        MANUAL_KEYWORDS = ["数值", "金额", "名称", "预算", "满减金额"]
        CALENDAR_KEYWORDS = ["日期"]
        
        dropdown_options = {}
        for i in range(1, len(headers) + 1):
            h_text = str(headers[i-1])
            is_date = any(k in h_text for k in CALENDAR_KEYWORDS)
            # 排除“类型”，确保折扣类型依然是下拉框
            is_manual = any(k in h_text for k in MANUAL_KEYWORDS) or ("折扣" in h_text and "类型" not in h_text)
            
            if not is_date and not is_manual:
                val8 = ws_gen.cell(row=8, column=i).value
                val9 = ws_gen.cell(row=9, column=i).value
                if val8 or val9:
                    opts = list(dict.fromkeys(filter(None, [str(val8), str(val9)])))
                    if len(opts) >= 1: dropdown_options[h_text] = opts

        with st.form("gen_form"):
            user_data = {}
            col1, col2 = st.columns(2)
            for i, h in enumerate(headers):
                target_col = col1 if i % 2 == 0 else col2
                h_str = str(h)
                if "ASIN" in h_str.upper():
                    raw_asin = target_col.text_area(f"{h}", placeholder="粘贴 ASIN...")
                    user_data[h] = CouponProcessor.clean_asin_input(raw_asin)
                elif any(k in h_str for k in CALENDAR_KEYWORDS):
                    picked_date = target_col.date_input(f"{h}", value=datetime.now(), key=f"date_p_{i}")
                    user_data[h] = picked_date.strftime("%Y-%m-%d")
                elif any(k in h_str for k in MANUAL_KEYWORDS) or ("折扣" in h_str and "类型" not in h_str):
                    user_data[h] = target_col.text_input(f"{h}", placeholder="手动输入...")
                elif h_str in dropdown_options:
                    user_data[h] = target_col.selectbox(f"{h}", options=dropdown_options[h_str])
                else:
                    user_data[h] = target_col.text_input(f"{h}")
            
            if st.form_submit_button("🚀 生成提报文件"):
                for col_idx, h in enumerate(headers, 1):
                    new_cell = ws_gen.cell(row=10, column=col_idx, value=user_data[h])
                    source_cell = ws_gen.cell(row=9, column=col_idx)
                    if source_cell.has_style:
                        new_cell.font, new_cell.border, new_cell.fill, new_cell.alignment = \
                            copy(source_cell.font), copy(source_cell.border), copy(source_cell.fill), copy(source_cell.alignment)
                
                out_gen = BytesIO()
                wb_gen.save(out_gen)
                st.session_state.gen_file = out_gen.getvalue()
                st.session_state.gen_ready = True

        if st.session_state.get('gen_ready'):
            st.download_button("📥 下载生成的提报 Excel", st.session_state.gen_file, "New_Coupon.xlsx", use_container_width=True)

# --- 第二阶段：报错修复 (采用您提供的独立完整逻辑) ---
with tab2:
    if not all_listing_file or not template_file:
        st.info("💡 请上传数据源以进入修复模式。")
    else:
        if 'master_df' not in st.session_state:
            with st.spinner("正在深度解析模板信息..."):
                # 读取 Listing
                for enc in ['utf-8', 'utf-16', 'gbk', 'utf-8-sig']:
                    try:
                        all_listing_file.seek(0)
                        if all_listing_file.name.endswith('.txt'):
                            df_l = pd.read_csv(all_listing_file, sep='\t', encoding=enc)
                        else:
                            df_l = pd.read_excel(all_listing_file)
                        df_l.columns = [c.lower().strip() for c in df_l.columns]
                        break
                    except: continue
                
                # 读取模板
                template_file.seek(0)
                wb_err = openpyxl.load_workbook(template_file, data_only=True)
                ws_err = wb_err.active
                headers = [cell.value for cell in ws_err[7]]
                
                asin_col = next((c for c in df_l.columns if 'asin' in c), None)
                price_col = next((c for c in df_l.columns if 'price' in c or '价格' in c), None)
                e_asin_h = next((c for c in headers if 'ASIN' in str(c)), headers[0])
                e_disc_h = next((c for c in headers if '折扣' in str(c) and '数值' in str(c)), None)

                rows = []
                for r_idx, row in enumerate(ws_err.iter_rows(min_row=10), 10):
                    vals = [cell.value for cell in row]
                    if not any(vals): continue
                    # 批注在最后一列或对应单元格
                    comment_obj = row[-1].comment
                    comment_text = comment_obj.text if comment_obj else ""
                    row_dict = {headers[i]: v for i, v in enumerate(vals) if i < len(headers)}
                    
                    asins = [a.strip() for a in str(row_dict.get(e_asin_h, "")).replace(',', ';').split(';') if a.strip()]
                    err_map = CouponProcessor.parse_error_details(comment_text)
                    
                    for a in asins:
                        p_match = df_l[df_l[asin_col] == a][price_col].values if asin_col else []
                        orig_p = p_match[0] if len(p_match) > 0 else None
                        info = err_map.get(a, {})
                        is_err = a in err_map
                        curr_d = row_dict.get(e_disc_h, 0.05)
                        suggested = curr_d
                        if is_err and orig_p and info.get('req_price'):
                            needed = math.ceil(((float(orig_p) - float(info.get('req_price'))) / float(orig_p)) * 100)
                            suggested = needed / 100 if curr_d < 1 else needed

                        rows.append({
                            "决策": info.get('default_decision', "保留"),
                            "ASIN": a, "状态": "❌ 批注报错" if is_err else "✅ 正常",
                            "详细报错原因": info.get('reason', "-"), "拟提报折扣": suggested,
                            "Listing原价": orig_p, "要求净价": info.get('req_price'),
                            "原始行号": r_idx
                        })
                st.session_state.master_df = pd.DataFrame(rows)
                st.session_state.orig_headers = headers

        # 应用侧边栏筛选
        if st.session_state.master_df is not None:
            mask = st.session_state.master_df['状态'].isin(status_sel)
            if reason_kw:
                mask = mask & st.session_state.master_df['详细报错原因'].str.contains(reason_kw, case=False)
            
            df_filtered = st.session_state.master_df[mask].copy()

            st.subheader("🛠️ 修复决策台")
            edited = st.data_editor(
                df_filtered,
                column_config={
                    "决策": st.column_config.SelectboxColumn("决策", options=["保留", "剔除"]),
                    "拟提报折扣": st.column_config.NumberColumn("折扣数值", format="%.2f"),
                    "详细报错原因": st.column_config.TextColumn("报错原因", width="large"),
                    "原始行号": None
                },
                disabled=['ASIN', '状态', '详细报错原因', 'Listing原价', '要求净价'],
                hide_index=True, use_container_width=True, key="repair_editor"
            )

            # 同步编辑内容
            if not edited.equals(df_filtered):
                for idx in edited.index:
                    st.session_state.master_df.loc[idx, '决策'] = edited.loc[idx, '决策']
                    st.session_state.master_df.loc[idx, '拟提报折扣'] = edited.loc[idx, '拟提报折扣']
                st.rerun()

            if st.button("🚀 生成并导出完整信息 Excel", use_container_width=True):
                template_file.seek(0)
                wb_out = openpyxl.load_workbook(template_file)
                ws_out = wb_out.active
                
                # 备份原始行
                row_backup = {}
                for r_idx in st.session_state.master_df['原始行号'].unique():
                    row_backup[r_idx] = [ws_out.cell(row=r_idx, column=c).value for c in range(1, ws_out.max_column + 1)]
                
                # 清空旧数据
                for r in range(10, ws_out.max_row + 1):
                    for c in range(1, ws_out.max_column + 1): ws_out.cell(row=r, column=c).value = None

                final_keep = st.session_state.master_df[st.session_state.master_df['决策'] == "保留"]
                
                # 定位列
                a_idx, d_idx = 1, 3
                for i, h in enumerate(st.session_state.orig_headers, 1):
                    if h and 'ASIN' in str(h): a_idx = i
                    if h and '折扣' in str(h) and '数值' in str(h): d_idx = i

                curr_r = 10
                for (orig_line, disc), group in final_keep.groupby(['原始行号', '拟提报折扣']):
                    orig_vals = row_backup.get(orig_line)
                    if orig_vals:
                        for c_idx, val in enumerate(orig_vals, 1):
                            target_cell = ws_out.cell(row=curr_r, column=c_idx, value=val)
                            source_cell = ws_out.cell(row=orig_line, column=c_idx)
                            if source_cell.has_style:
                                target_cell.font, target_cell.border, target_cell.fill, target_cell.number_format, target_cell.alignment = \
                                    copy(source_cell.font), copy(source_cell.border), copy(source_cell.fill), copy(source_cell.number_format), copy(source_cell.alignment)
                        
                        ws_out.cell(row=curr_r, column=a_idx).value = ";".join(group['ASIN'].tolist())
                        ws_out.cell(row=curr_r, column=d_idx).value = disc
                        curr_r += 1

                out_repair = BytesIO()
                wb_out.save(out_repair)
                st.success("✅ 修复版文件已生成！")
                st.download_button("📥 点击下载修复后的完整 Excel", out_repair.getvalue(), "Coupon_Full_Fixed.xlsx")
