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
        # 解析批注中的 ASIN 和 报错详情
        blocks = re.split(r'([A-Z0-9]{10})\n', str(comment_text))
        if len(blocks) > 1:
            for i in range(1, len(blocks), 2):
                asin = blocks[i].strip()
                content = blocks[i+1]
                # 提取要求的价格（兼容多种描述）
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

# --- 侧边栏 ---
with st.sidebar:
    st.header("📂 第一阶段数据源")
    site_template = st.file_uploader("上传：站点空白 Coupon 模板 (导出底稿将参考此文件)", type=['xlsx'], key="template_gen")
    
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

# --- 第一阶段：生成提报 ---
with tab1:
    if not site_template:
        st.warning("👈 请在左侧上传【站点空白 Coupon 模板】。")
    else:
        wb_gen = openpyxl.load_workbook(site_template, data_only=True)
        ws_gen = wb_gen.active
        headers = [cell.value for cell in ws_gen[7] if cell.value]
        
        MANUAL_KEYWORDS = ["数值", "金额", "名称", "预算", "满减金额"]
        CALENDAR_KEYWORDS = ["日期"]
        
        dropdown_options = {}
        for i, h in enumerate(headers, 1):
            h_text = str(h)
            is_date = any(k in h_text for k in CALENDAR_KEYWORDS)
            is_manual = any(k in h_text for k in MANUAL_KEYWORDS) or ("折扣" in h_text and "类型" not in h_text)
            if not is_date and not is_manual:
                val8, val9 = ws_gen.cell(row=8, column=i).value, ws_gen.cell(row=9, column=i).value
                opts = list(dict.fromkeys(filter(None, [str(val8), str(val9)])))
                if opts: dropdown_options[h_text] = opts

        with st.form("gen_form"):
            user_data, col1, col2 = {}, *st.columns(2)
            for i, h in enumerate(headers):
                target_col = col1 if i % 2 == 0 else col2
                h_str = str(h)
                if "ASIN" in h_str.upper():
                    user_data[h] = target_col.text_area(f"{h}", placeholder="粘贴 ASIN...")
                elif any(k in h_str for k in CALENDAR_KEYWORDS):
                    user_data[h] = target_col.date_input(f"{h}", value=datetime.now(), key=f"d_g_{i}").strftime("%Y-%m-%d")
                elif any(k in h_str for k in MANUAL_KEYWORDS) or ("折扣" in h_str and "类型" not in h_str):
                    user_data[h] = target_col.text_input(f"{h}")
                elif h_str in dropdown_options:
                    user_data[h] = target_col.selectbox(f"{h}", options=dropdown_options[h_str])
                else:
                    user_data[h] = target_col.text_input(f"{h}")
            
            if st.form_submit_button("🚀 生成提报文件"):
                for col_idx, h in enumerate(headers, 1):
                    val = CouponProcessor.clean_asin_input(user_data[h]) if "ASIN" in str(h).upper() else user_data[h]
                    ws_gen.cell(row=10, column=col_idx, value=val)
                out_gen = BytesIO()
                wb_gen.save(out_gen)
                st.session_state.gen_file = out_gen.getvalue()
                st.session_state.gen_done = True

        if st.session_state.get('gen_done'):
            st.download_button("📥 下载提报 Excel", st.session_state.gen_file, "Coupon_Upload.xlsx")

# --- 第二阶段：报错修复 ---
with tab2:
    if not all_listing_file or not error_feedback_file:
        st.warning("👈 请在左侧上传【All Listing 报告】和【亚马逊返回的报错文件】。")
    elif not site_template:
        st.error("⚠️ 请在左侧【第一阶段】处上传站点空白模板，作为修复后的导出底稿。")
    else:
        if 'master_df' not in st.session_state:
            with st.spinner("深度比对数据中..."):
                for enc in ['utf-8', 'utf-16', 'gbk', 'utf-8-sig']:
                    try:
                        all_listing_file.seek(0)
                        df_l = pd.read_csv(all_listing_file, sep='\t', encoding=enc) if all_listing_file.name.endswith('.txt') else pd.read_excel(all_listing_file)
                        df_l.columns = [c.lower().strip() for c in df_l.columns]
                        break
                    except: continue
                
                error_feedback_file.seek(0)
                wb_err = openpyxl.load_workbook(error_feedback_file, data_only=True)
                ws_err = wb_err.active
                err_headers = [cell.value for cell in ws_err[7] if cell.value]
                
                asin_col = next((c for c in df_l.columns if 'asin' in c), None)
                price_col = next((c for c in df_l.columns if 'price' in c or '价格' in c), None)
                e_asin_h = next((h for h in err_headers if h and 'ASIN' in str(h)), None)
                e_disc_h = next((h for h in err_headers if h and '折扣' in str(h) and '数值' in str(h)), None)

                rows = []
                header_map = {h: i for i, h in enumerate([cell.value for cell in ws_err[7]]) if h}
                
                for r_idx in range(10, ws_err.max_row + 1):
                    row_vals = [ws_err.cell(row=r_idx, column=c).value for c in range(1, ws_err.max_column + 1)]
                    if not any(row_vals): continue
                    
                    # 定位 M/N 列或最后一列（带批注的那一列）
                    comment_cell = ws_err.cell(row=r_idx, column=ws_err.max_column)
                    comment_text = comment_cell.comment.text if comment_cell and comment_cell.comment else ""
                    
                    asin_val = ws_err.cell(row=r_idx, column=header_map[e_asin_h]+1).value
                    asins = [a.strip() for a in str(asin_val).replace(',', ';').replace('\n', ';').split(';') if a.strip()]
                    err_map = CouponProcessor.parse_error_details(comment_text)
                    
                    for a in asins:
                        p_match = df_l[df_l[asin_col] == a][price_col].values if asin_col else []
                        orig_p = p_match[0] if len(p_match) > 0 else 0
                        info = err_map.get(a, {})
                        
                        curr_d = ws_err.cell(row=r_idx, column=header_map[e_disc_h]+1).value if e_disc_h in header_map else 0.05
                        suggested = curr_d
                        if a in err_map and orig_p and info.get('req_price'):
                            needed = math.ceil(((float(orig_p) - float(info.get('req_price'))) / float(orig_p)) * 100)
                            suggested = needed / 100 if float(curr_d or 0) < 1 else max(needed, 5)

                        # 【修复1】确保 req_price 被填入 rows 列表
                        rows.append({
                            "决策": info.get('default_decision', "保留"), 
                            "ASIN": a, 
                            "状态": "❌ 批注报错" if a in err_map else "✅ 正常",
                            "详细报错原因": info.get('reason', "-"), 
                            "要求净价格": info.get('req_price', "-"), 
                            "拟提报折扣": suggested,
                            "Listing原价": orig_p, 
                            "原始行号": r_idx
                        })
                st.session_state.master_df = pd.DataFrame(rows)

        if st.session_state.get('master_df') is not None:
            # 【修复2】联动状态筛选逻辑
            mask = st.session_state.master_df['状态'].isin(status_sel)
            if reason_kw:
                mask = mask & st.session_state.master_df['详细报错原因'].str.contains(reason_kw, case=False)
            
            df_show = st.session_state.master_df[mask].copy()
            
            st.subheader("🛠️ 修复决策台")
            edited = st.data_editor(
                df_show,
                column_config={
                    "决策": st.column_config.SelectboxColumn("决策", options=["保留", "剔除"]),
                    "要求净价格": st.column_config.NumberColumn("要求净价格", format="%.2f"),
                    "原始行号": None
                },
                disabled=['ASIN', '状态', '详细报错原因', '要求净价格', 'Listing原价'],
                hide_index=True, use_container_width=True, key="fix_edit"
            )

            # 同步编辑结果
            if not edited.equals(df_show):
                for idx in edited.index:
                    st.session_state.master_df.loc[idx, '决策'] = edited.loc[idx, '决策']
                    st.session_state.master_df.loc[idx, '拟提报折扣'] = edited.loc[idx, '拟提报折扣']
                st.rerun()

            if st.button("🚀 生成纯净修复版 Excel", use_container_width=True):
                # 使用第一阶段底稿
                site_template.seek(0)
                wb_final = openpyxl.load_workbook(site_template)
                ws_final = wb_final.active
                
                # 参考报错文件提取非 ASIN 数据
                error_feedback_file.seek(0)
                wb_err_ref = openpyxl.load_workbook(error_feedback_file, data_only=True)
                ws_err_ref = wb_err_ref.active
                
                final_headers = [ws_final.cell(row=7, column=c).value for c in range(1, ws_final.max_column + 1)]
                a_idx = next((i for i, h in enumerate(final_headers, 1) if h and 'ASIN' in str(h)), 1)
                d_idx = next((i for i, h in enumerate(final_headers, 1) if h and '折扣' in str(h) and '数值' in str(h)), 3)

                final_keep = st.session_state.master_df[st.session_state.master_df['决策'] == "保留"]
                curr_row = 10
                
                for (orig_l, disc), group in final_keep.groupby(['原始行号', '拟提报折扣']):
                    for c_idx in range(1, len(final_headers) + 1):
                        orig_val = ws_err_ref.cell(row=orig_l, column=c_idx).value
                        target_cell = ws_final.cell(row=curr_row, column=c_idx, value=orig_val)
                        
                        ref_style_cell = ws_final.cell(row=9, column=c_idx)
                        if ref_style_cell.has_style:
                            target_cell.font, target_cell.border, target_cell.fill, target_cell.alignment = \
                                copy(ref_style_cell.font), copy(ref_style_cell.border), copy(ref_style_cell.fill), copy(ref_style_cell.alignment)
                    
                    ws_final.cell(row=curr_row, column=a_idx).value = ";".join(group['ASIN'].tolist())
                    ws_final.cell(row=curr_row, column=d_idx).value = disc
                    curr_row += 1
                
                out_fix = BytesIO()
                wb_final.save(out_fix)
                st.success("✅ 文件已成功基于空白底稿生成。")
                st.download_button("📥 下载纯净版修复结果", out_fix.getvalue(), "Fixed_Submission_Clean.xlsx")
