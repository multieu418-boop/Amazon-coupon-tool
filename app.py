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
st.set_page_config(page_title="Cupshe Amazon Coupon 集成系统", layout="wide")

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
                auto_exclude = "没有经验证的参考价" in reason
                error_map[asin] = {
                    "req_price": req_p, 
                    "reason": reason, 
                    "default_decision": "剔除" if auto_exclude else "保留"
                }
        return error_map

# --- 侧边栏：严格区分文件用途 ---
with st.sidebar:
    st.header("📂 第一阶段数据源")
    site_template = st.file_uploader("上传：站点空白 Coupon 模板", type=['xlsx'], key="template_gen")
    
    st.header("📂 第二阶段数据源")
    all_listing_file = st.file_uploader("上传：All Listing 报告", type=['txt', 'csv', 'xlsx'], key="listing_rep")
    error_feedback_file = st.file_uploader("上传：Amazon 返回的报错文件", type=['xlsx'], key="error_rep")
    
    st.divider()
    if error_feedback_file and all_listing_file:
        st.header("⚙️ 修复筛选配置")
        status_sel = st.multiselect("ASIN 状态筛选", ["✅ 正常", "❌ 批注报错"], default=["✅ 正常", "❌ 批注报错"])
        reason_kw = st.text_input("报错原因关键词过滤")

    if st.button("🔄 清空所有上传"):
        st.session_state.clear()
        st.rerun()

st.title("🎯 Cupshe Amazon Coupon 自动化管理系统")

tab1, tab2 = st.tabs(["🔵 第一阶段：生成提报", "🔴 第二阶段：报错修复"])

# --- 第一阶段：生成提报 (使用 site_template) ---
with tab1:
    if not site_template:
        st.warning("👈 请在左侧上传【站点空白 Coupon 模板】以开始。")
    else:
        wb_gen = openpyxl.load_workbook(site_template, data_only=True)
        ws_gen = wb_gen.active
        headers = [cell.value for cell in ws_gen[7] if cell.value]
        
        MANUAL_KEYWORDS = ["数值", "金额", "名称", "预算", "满减金额"]
        CALENDAR_KEYWORDS = ["日期"]
        
        dropdown_options = {}
        for i in range(1, len(headers) + 1):
            h_text = str(headers[i-1])
            is_date = any(k in h_text for k in CALENDAR_KEYWORDS)
            is_manual = any(k in h_text for k in MANUAL_KEYWORDS) or ("折扣" in h_text and "类型" not in h_text)
            
            if not is_date and not is_manual:
                val8 = ws_gen.cell(row=8, column=i).value
                val9 = ws_gen.cell(row=9, column=i).value
                if val8 or val9:
                    opts = list(dict.fromkeys(filter(None, [str(val8), str(val9)])))
                    if len(opts) >= 1: dropdown_options[h_text] = opts

        st.subheader("📝 录入新优惠券需求")
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
                    picked_date = target_col.date_input(f"{h}", value=datetime.now(), key=f"date_gen_{i}")
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
                st.session_state.gen_file_data = out_gen.getvalue()
                st.session_state.gen_done = True

        if st.session_state.get('gen_done'):
            st.download_button("📥 下载生成的文件", st.session_state.gen_file_data, "Coupon_Upload.xlsx", use_container_width=True)

# --- 第二阶段：报错修复 (使用 all_listing_file 和 error_feedback_file) ---
with tab2:
    if not all_listing_file or not error_feedback_file:
        st.warning("👈 请在左侧同时上传【All Listing 报告】和【报错文件】以开始修复。")
    else:
        if 'master_df' not in st.session_state:
            with st.spinner("深度解析报错批注中..."):
                # 1. 解析 Listing
                for enc in ['utf-8', 'utf-16', 'gbk', 'utf-8-sig']:
                    try:
                        all_listing_file.seek(0)
                        df_l = pd.read_csv(all_listing_file, sep='\t', encoding=enc) if all_listing_file.name.endswith('.txt') else pd.read_excel(all_listing_file)
                        df_l.columns = [c.lower().strip() for c in df_l.columns]
                        break
                    except: continue
                
                # 2. 解析报错文件
                error_feedback_file.seek(0)
                wb_err = openpyxl.load_workbook(error_feedback_file, data_only=True)
                ws_err = wb_err.active
                err_headers = [cell.value for cell in ws_err[7]]
                
                asin_col = next((c for c in df_l.columns if 'asin' in c), None)
                price_col = next((c for c in df_l.columns if 'price' in c or '价格' in c), None)
                e_asin_h = next((c for c in err_headers if 'ASIN' in str(c)), err_headers[0])
                e_disc_h = next((c for c in err_headers if '折扣' in str(c) and '数值' in str(c)), None)

                rows = []
                for r_idx, row in enumerate(ws_err.iter_rows(min_row=10), 10):
                    vals = [cell.value for cell in row]
                    if not any(vals): continue
                    comment_text = row[-1].comment.text if row[-1].comment else ""
                    row_dict = {err_headers[i]: v for i, v in enumerate(vals) if i < len(err_headers)}
                    asins = [a.strip() for a in str(row_dict.get(e_asin_h, "")).replace(',', ';').split(';') if a.strip()]
                    err_map = CouponProcessor.parse_error_details(comment_text)
                    
                    for a in asins:
                        p_match = df_l[df_l[asin_col] == a][price_col].values if asin_col else []
                        orig_p = p_match[0] if len(p_match) > 0 else None
                        info = err_map.get(a, {})
                        suggested = row_dict.get(e_disc_h, 0.05)
                        if a in err_map and orig_p and info.get('req_price'):
                            needed = math.ceil(((float(orig_p) - float(info.get('req_price'))) / float(orig_p)) * 100)
                            suggested = needed / 100 if suggested < 1 else max(needed, 5)

                        rows.append({
                            "决策": info.get('default_decision', "保留"), "ASIN": a, 
                            "状态": "❌ 批注报错" if a in err_map else "✅ 正常",
                            "详细报错原因": info.get('reason', "-"), "拟提报折扣": suggested,
                            "Listing原价": orig_p, "原始行号": r_idx
                        })
                st.session_state.master_df = pd.DataFrame(rows)
                st.session_state.orig_err_headers = err_headers

        if st.session_state.get('master_df') is not None:
            mask = st.session_state.master_df['状态'].isin(status_sel)
            if reason_kw: mask = mask & st.session_state.master_df['详细报错原因'].str.contains(reason_kw, case=False)
            
            df_fix = st.session_state.master_df[mask].copy()
            st.subheader("🛠️ 修复决策台")
            edited = st.data_editor(
                df_fix,
                column_config={"决策": st.column_config.SelectboxColumn("决策", options=["保留", "剔除"]), "原始行号": None},
                disabled=['ASIN', '状态', '详细报错原因', 'Listing原价'],
                hide_index=True, use_container_width=True, key="edit_fix"
            )

            if not edited.equals(df_fix):
                for idx in edited.index:
                    st.session_state.master_df.loc[idx, '决策'] = edited.loc[idx, '决策']
                    st.session_state.master_df.loc[idx, '拟提报折扣'] = edited.loc[idx, '拟提报折扣']
                st.rerun()

            if st.button("🚀 生成修复版 Excel", use_container_width=True):
                error_feedback_file.seek(0)
                wb_out = openpyxl.load_workbook(error_feedback_file)
                ws_out = wb_out.active
                row_backup = {r: [ws_out.cell(row=r, column=c).value for c in range(1, ws_out.max_column+1)] for r in st.session_state.master_df['原始行号'].unique()}
                for r in range(10, ws_out.max_row+1):
                    for c in range(1, ws_out.max_column+1): ws_out.cell(row=r, column=c).value = None
                
                final_keep = st.session_state.master_df[st.session_state.master_df['决策'] == "保留"]
                a_idx = next(i for i, h in enumerate(st.session_state.orig_err_headers, 1) if h and 'ASIN' in str(h))
                d_idx = next(i for i, h in enumerate(st.session_state.orig_err_headers, 1) if h and '折扣' in str(h) and '数值' in str(h))
                
                curr_r = 10
                for (orig_l, disc), group in final_keep.groupby(['原始行号', '拟提报折扣']):
                    vals = row_backup.get(orig_l)
                    for c_idx, v in enumerate(vals, 1):
                        target = ws_out.cell(row=curr_r, column=c_idx, value=v)
                        source = ws_out.cell(row=orig_l, column=c_idx)
                        if source.has_style:
                            target.font, target.border, target.fill, target.number_format, target.alignment = \
                                copy(source.font), copy(source.border), copy(source.fill), copy(source.number_format), copy(source.alignment)
                    ws_out.cell(row=curr_r, column=a_idx).value = ";".join(group['ASIN'].tolist())
                    ws_out.cell(row=curr_r, column=d_idx).value = disc
                    curr_r += 1
                
                out_fix = BytesIO()
                wb_out.save(out_fix)
                st.download_button("📥 下载修复结果", out_fix.getvalue(), "Fixed_Coupon.xlsx")
