import streamlit as st
import pandas as pd
import openpyxl
from copy import copy
import re
import io
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
        """需求1：精准提取“要求”的价格"""
        error_map = {}
        if not comment_text: return error_map
        blocks = re.split(r'([A-Z0-9]{10})\n', str(comment_text))
        if len(blocks) > 1:
            for i in range(1, len(blocks), 2):
                asin = blocks[i].strip()
                content = blocks[i+1]
                # 精准匹配“要求”后的数字
                req_match = re.search(r'(?:要求的(?:净价格|最高商品价格)|Required net price|Maximum product price allowed)：?\s*[^\d]*([\d\.]+)', content)
                req_p = float(req_match.group(1)) if req_match else None
                reason_part = re.split(r'(?:要求的|当前|Maximum|Required)', content)[0]
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
    st.header("📂 阶段一：上传空白模板")
    site_template = st.file_uploader("站点空白 Coupon 模板 (导出底稿)", type=['xlsx'], key="base_tpl")
    
    st.header("📂 阶段二：上传报错源")
    all_listing_file = st.file_uploader("All Listing 报告", type=['txt', 'csv', 'xlsx'], key="l_rep")
    error_feedback_file = st.file_uploader("亚马逊返回的报错文件 (带批注)", type=['xlsx'], key="e_rep")
    
    st.divider()
    if error_feedback_file and all_listing_file:
        st.header("⚙️ 辅助筛选")
        status_sel = st.multiselect("ASIN 状态", ["✅ 正常", "❌ 批注报错"], default=["✅ 正常", "❌ 批注报错"])
        reason_kw = st.text_input("报错原因关键词搜索 (可选)")

    if st.sidebar.button("🔄 重置并清空缓存"):
        st.session_state.clear()
        st.rerun()

st.title("🎯 Cupshe Amazon Coupon 自动化管理系统")
tab1, tab2 = st.tabs(["🔵 第一阶段：提报生成", "🔴 第二阶段：报错修复"])

# --- 第一阶段逻辑 (保持不变) ---
with tab1:
    if not site_template:
        st.warning("👈 请先在左侧上传【站点空白模板】")
    else:
        wb_gen = openpyxl.load_workbook(site_template, data_only=True)
        ws_gen = wb_gen.active
        headers = [ws_gen.cell(row=7, column=c).value for c in range(1, ws_gen.max_column + 1) if ws_gen.cell(row=7, column=c).value]
        with st.form("gen_form"):
            user_input = {}
            col1, col2 = st.columns(2)
            for i, h in enumerate(headers):
                target_col = col1 if i % 2 == 0 else col2
                if "ASIN" in str(h).upper():
                    user_input[h] = target_col.text_area(h, placeholder="粘贴 ASIN...")
                else:
                    user_input[h] = target_col.text_input(h)
            if st.form_submit_button("🚀 生成提报文件"):
                for c_idx, h in enumerate(headers, 1):
                    val = CouponProcessor.clean_asin_input(user_input[h]) if "ASIN" in str(h).upper() else user_input[h]
                    ws_gen.cell(row=10, column=c_idx, value=val)
                out_gen = io.BytesIO()
                wb_gen.save(out_gen)
                st.session_state.gen_data = out_gen.getvalue()
        if "gen_data" in st.session_state:
            st.download_button("📥 下载提报文件", st.session_state.gen_data, "Initial_Coupon.xlsx")

# --- 第二阶段逻辑 (优化修复) ---
with tab2:
    if not all_listing_file or not error_feedback_file or not site_template:
        st.warning("👈 请确保左侧已上传：1.空白底稿 2.Listing报告 3.报错文件")
    else:
        # 1. 自动解析 (仅在文件变更时运行)
        if 'master_df' not in st.session_state:
            with st.spinner("正在精准匹配 ASIN 数据..."):
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
                err_headers = [ws_err.cell(row=7, column=c).value for c in range(1, ws_err.max_column + 1)]
                a_col_idx = next((i for i, h in enumerate(err_headers) if h and 'ASIN' in str(h)), 0)
                d_col_idx = next((i for i, h in enumerate(err_headers) if h and '折扣' in str(h) and '数值' in str(h)), 2)

                rows = []
                for r_idx in range(10, ws_err.max_row + 1):
                    row_cells = [ws_err.cell(row=r_idx, column=c).value for c in range(1, ws_err.max_column + 1)]
                    if not any(row_cells): continue
                    comment_cell = ws_err.cell(row=r_idx, column=ws_err.max_column)
                    comment_text = comment_cell.comment.text if comment_cell and comment_cell.comment else ""
                    err_map = CouponProcessor.parse_error_details(comment_text)
                    asin_str = str(ws_err.cell(row=r_idx, column=a_col_idx+1).value)
                    asins = [a.strip() for a in asin_str.replace(',', ';').replace('\n', ';').split(';') if a.strip()]
                    for a in asins:
                        asin_col = next((c for c in df_l.columns if 'asin' in c), None)
                        p_col = next((c for c in df_l.columns if 'price' in c or '价格' in c), None)
                        p_match = df_l[df_l[asin_col] == a][p_col].values if asin_col else []
                        orig_p = p_match[0] if len(p_match) > 0 else 0
                        info = err_map.get(a, {})
                        curr_d = ws_err.cell(row=r_idx, column=d_col_idx+1).value or 0.05
                        suggested = curr_d
                        if info.get('req_price') and orig_p:
                            needed = math.ceil(((float(orig_p) - float(info['req_price'])) / float(orig_p)) * 100)
                            suggested = needed / 100 if float(curr_d) < 1 else max(needed, 5)
                        rows.append({
                            "决策": info.get('default_decision', "保留"), "ASIN": a, 
                            "状态": "❌ 批注报错" if a in err_map else "✅ 正常",
                            "详细报错原因": info.get('reason', "-"), "要求净价格": info.get('req_price', "-"),
                            "拟提报折扣": suggested, "Listing原价": orig_p, "原始行号": r_idx
                        })
                st.session_state.master_df = pd.DataFrame(rows)

        # 2. 决策工作台
        mask = st.session_state.master_df['状态'].isin(status_sel)
        if reason_kw:
            mask = mask & st.session_state.master_df['详细报错原因'].str.contains(reason_kw, case=False)
        df_show = st.session_state.master_df[mask].copy()
        
        st.subheader("🛠️ 修复决策台")
        edited = st.data_editor(
            df_show,
            column_config={
                "决策": st.column_config.SelectboxColumn("决策", options=["保留", "剔除"]),
                "拟提报折扣": st.column_config.NumberColumn("拟提报折扣", format="%.2f"),
                "原始行号": None
            },
            disabled=['ASIN', '状态', '详细报错原因', '要求净价格', 'Listing原价'],
            hide_index=True, use_container_width=True, key="fix_editor_integrated"
        )

        # 只要编辑器有动作就同步状态
        if not edited.equals(df_show):
            for idx in edited.index:
                st.session_state.master_df.loc[idx, '决策'] = edited.loc[idx, '决策']
                st.session_state.master_df.loc[idx, '拟提报折扣'] = edited.loc[idx, '拟提报折扣']
            st.rerun()

        st.divider()
        
        # 3. 增强版生成与下载区域
        col_gen, col_dl = st.columns(2)
        
        with col_gen:
            if st.button("🚀 生成纯净修复版 Excel", use_container_width=True, type="primary"):
                with st.status("正在从空白底稿构建文件...", expanded=True) as status:
                    try:
                        site_template.seek(0)
                        wb_final = openpyxl.load_workbook(site_template)
                        ws_final = wb_final.active
                        error_feedback_file.seek(0)
                        wb_err_ref = openpyxl.load_workbook(error_feedback_file, data_only=True)
                        ws_err_ref = wb_err_ref.active
                        
                        f_headers = [ws_final.cell(row=7, column=c).value for c in range(1, ws_final.max_column + 1)]
                        a_idx = next((i for i, h in enumerate(f_headers, 1) if h and 'ASIN' in str(h)), 1)
                        d_idx = next((i for i, h in enumerate(f_headers, 1) if h and '折扣' in str(h) and '数值' in str(h)), 3)
                        
                        final_keep = st.session_state.master_df[st.session_state.master_df['决策'] == "保留"]
                        
                        if final_keep.empty:
                            st.error("当前决策下没有保留任何 ASIN，无法生成。")
                        else:
                            curr_row = 10
                            for (orig_line, disc), group in final_keep.groupby(['原始行号', '拟提报折扣']):
                                for c_idx in range(1, len(f_headers) + 1):
                                    orig_val = ws_err_ref.cell(row=orig_line, column=c_idx).value
                                    target_cell = ws_final.cell(row=curr_row, column=c_idx, value=orig_val)
                                    ref_style = ws_final.cell(row=9, column=c_idx)
                                    if ref_style.has_style:
                                        target_cell.font, target_cell.border = copy(ref_style.font), copy(ref_style.border)
                                        target_cell.fill, target_cell.alignment = copy(ref_style.fill), copy(ref_style.alignment)

                                ws_final.cell(row=curr_row, column=a_idx).value = ";".join(group['ASIN'].tolist())
                                ws_final.cell(row=curr_row, column=d_idx).value = disc
                                curr_row += 1
                            
                            out_fix = io.BytesIO()
                            wb_final.save(out_fix)
                            st.session_state.fix_data_ready = out_fix.getvalue()
                            status.update(label="✅ 文件已成功生成！可点击右侧下载。", state="complete", expanded=False)
                    except Exception as e:
                        st.error(f"处理出错: {str(e)}")

        with col_dl:
            if "fix_data_ready" in st.session_state:
                st.download_button(
                    label="📥 点击下载修复后的纯净 Excel",
                    data=st.session_state.fix_data_ready,
                    file_name=f"Fixed_Coupon_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.info("请先点击左侧【🚀 生成】按钮开始处理")
