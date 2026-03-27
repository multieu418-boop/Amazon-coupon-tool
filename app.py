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
                req_match = re.search(r'(?:要求的(?:净价格|最高商品价格)|Required net price|Maximum product price allowed)：?\s*[^\d]*([\d\.]+)', content)
                req_p = float(req_match.group(1)) if req_match else None
                reason_part = re.split(r'(?:要求的|当前|Maximum|Required)', content)[0]
                reason = reason_part.strip().replace('\n', ' ')
                auto_exclude = "没有经验证的参考价" in reason
                error_map[asin] = {"req_price": req_p, "reason": reason, "default_decision": "剔除" if auto_exclude else "保留"}
        return error_map

# --- 侧边栏 ---
with st.sidebar:
    st.header("📂 原始底稿")
    site_template = st.file_uploader("1. 上传空白 Coupon 模板 (导出用)", type=['xlsx'], key="tpl")
    st.header("📂 报错修复源")
    all_listing_file = st.file_uploader("2. 上传 All Listing 报告", type=['txt', 'csv', 'xlsx'], key="list")
    error_feedback_file = st.file_uploader("3. 上传亚马逊报错文件", type=['xlsx'], key="err")
    
    if st.button("🔄 重置所有数据"):
        st.session_state.clear()
        st.rerun()

st.title("🎯 Cupshe Amazon Coupon 自动化管理系统")
tab1, tab2 = st.tabs(["🔵 第一阶段：提报生成", "🔴 第二阶段：报错修复"])

# --- 第一阶段：提报生成 ---
with tab1:
    if site_template:
        wb_gen = openpyxl.load_workbook(site_template, data_only=True)
        ws_gen = wb_gen.active
        headers = [ws_gen.cell(row=7, column=c).value for c in range(1, ws_gen.max_column + 1) if ws_gen.cell(row=7, column=c).value]
        with st.form("gen_form"):
            user_input = {}
            c1, c2 = st.columns(2)
            for i, h in enumerate(headers):
                t_col = c1 if i % 2 == 0 else c2
                user_input[h] = t_col.text_area(h) if "ASIN" in str(h).upper() else t_col.text_input(h)
            if st.form_submit_button("🚀 生成初始提报"):
                for idx, h in enumerate(headers, 1):
                    val = CouponProcessor.clean_asin_input(user_input[h]) if "ASIN" in str(h).upper() else user_input[h]
                    ws_gen.cell(row=10, column=idx, value=val)
                out = io.BytesIO(); wb_gen.save(out)
                st.session_state.gen_file = out.getvalue()
        if "gen_file" in st.session_state:
            st.download_button("📥 下载提报文件", st.session_state.gen_file, "Initial_Upload.xlsx")

# --- 第二阶段：报错修复 ---
with tab2:
    if not all_listing_file or not error_feedback_file or not site_template:
        st.warning("⚠️ 请确保左侧上传了：空白模板、Listing 报告、报错文件")
    else:
        # 数据解析 (存储在 session_state 防止刷新消失)
        if 'master_df' not in st.session_state:
            with st.spinner("数据深度解析中..."):
                # 解析 Listing
                for enc in ['utf-8', 'utf-16', 'gbk', 'utf-8-sig']:
                    try:
                        all_listing_file.seek(0)
                        df_l = pd.read_csv(all_listing_file, sep='\t', encoding=enc) if all_listing_file.name.endswith('.txt') else pd.read_excel(all_listing_file)
                        df_l.columns = [c.lower().strip() for c in df_l.columns]; break
                    except: continue
                # 解析报错
                error_feedback_file.seek(0)
                wb_err = openpyxl.load_workbook(error_feedback_file, data_only=True)
                ws_err = wb_err.active
                e_h = [ws_err.cell(row=7, column=c).value for c in range(1, ws_err.max_column + 1)]
                a_idx = next((i for i, h in enumerate(e_h) if h and 'ASIN' in str(h)), 0)
                d_idx = next((i for i, h in enumerate(e_h) if h and '折扣' in str(h) and '数值' in str(h)), 2)
                
                rows = []
                for r in range(10, ws_err.max_row + 1):
                    if not any([ws_err.cell(row=r, column=c).value for c in range(1, ws_err.max_column+1)]): continue
                    comm = ws_err.cell(row=r, column=ws_err.max_column).comment.text if ws_err.cell(row=r, column=ws_err.max_column).comment else ""
                    err_map = CouponProcessor.parse_error_details(comm)
                    asin_str = str(ws_err.cell(row=r, column=a_idx+1).value)
                    for a in [a.strip() for a in asin_str.replace(',',';').split(';') if a.strip()]:
                        p_col = next((c for c in df_l.columns if 'price' in c or '价格' in c), None)
                        a_col = next((c for c in df_l.columns if 'asin' in c), None)
                        p_match = df_l[df_l[a_col] == a][p_col].values if a_col else []
                        orig_p = p_match[0] if len(p_match) > 0 else 0
                        info = err_map.get(a, {})
                        curr_d = ws_err.cell(row=r, column=d_idx+1).value or 0.05
                        suggested = curr_d
                        if info.get('req_price') and orig_p:
                            needed = math.ceil(((float(orig_p) - float(info['req_price'])) / float(orig_p)) * 100)
                            suggested = needed / 100 if float(curr_d) < 1 else max(needed, 5)
                        rows.append({"决策": info.get('default_decision', "保留"), "ASIN": a, "状态": "❌ 批注报错" if a in err_map else "✅ 正常",
                                    "原因": info.get('reason', "-"), "要求净价格": info.get('req_price', "-"), "拟提报折扣": suggested, "Listing原价": orig_p, "原始行号": r})
                st.session_state.master_df = pd.DataFrame(rows)

        # 渲染决策台
        st.subheader("🛠️ 修复决策台")
        # 增加一个保存按钮，显式更新数据，避免组件自动 rerun 导致的按钮消失
        if st.button("💾 保存决策修改 (修改完表格后请点这里)"):
            st.toast("决策已锁定", icon="🔒")

        edited_df = st.data_editor(
            st.session_state.master_df,
            column_config={"决策": st.column_config.SelectboxColumn("决策", options=["保留", "剔除"]), "原始行号": None},
            disabled=['ASIN', '状态', '原因', '要求净价格', 'Listing原价'],
            hide_index=True, use_container_width=True, key="main_editor"
        )
        st.session_state.master_df = edited_df # 实时同步到 session

        # --- 核心导出区：使用独立 Form 确保按钮永远渲染 ---
        st.divider()
        st.subheader("📦 导出纯净修复文件")
        
        # 只要数据解析了，这个 Form 就会显示在页面底部，哪怕你不动任何数据
        with st.form("export_section"):
            st.info("💡 提示：点击下方按钮将基于【空白模板】生成结果，不包含亚马逊的 Upload Results 列。")
            submit_gen = st.form_submit_button("🚀 第一步：执行生成修复文件", use_container_width=True)
            
            if submit_gen:
                with st.status("正在从空白底稿构建文件...", expanded=True) as status:
                    site_template.seek(0)
                    wb_final = openpyxl.load_workbook(site_template)
                    ws_final = wb_final.active
                    error_feedback_file.seek(0)
                    wb_err_ref = openpyxl.load_workbook(error_feedback_file, data_only=True)
                    ws_err_ref = wb_err_ref.active
                    
                    f_h = [ws_final.cell(row=7, column=c).value for c in range(1, ws_final.max_column + 1)]
                    final_a_idx = next((i for i, h in enumerate(f_h, 1) if h and 'ASIN' in str(h)), 1)
                    final_d_idx = next((i for i, h in enumerate(f_h, 1) if h and '折扣' in str(h) and '数值' in str(h)), 3)
                    
                    final_keep = st.session_state.master_df[st.session_state.master_df['决策'] == "保留"]
                    curr_row = 10
                    for (orig_l, disc), group in final_keep.groupby(['原始行号', '拟提报折扣']):
                        for c_idx in range(1, len(f_h) + 1):
                            val = ws_err_ref.cell(row=orig_l, column=c_idx).value
                            target = ws_final.cell(row=curr_row, column=c_idx, value=val)
                            ref_s = ws_final.cell(row=9, column=c_idx)
                            if ref_s.has_style:
                                target.font, target.border = copy(ref_s.font), copy(ref_s.border)
                                target.fill, target.alignment = copy(ref_s.fill), copy(ref_s.alignment)
                        ws_final.cell(row=curr_row, column=final_a_idx).value = ";".join(group['ASIN'].tolist())
                        ws_final.cell(row=curr_row, column=final_d_idx).value = disc
                        curr_row += 1
                    
                    out_fix = io.BytesIO()
                    wb_final.save(out_fix)
                    st.session_state.final_excel = out_fix.getvalue()
                    status.update(label="✅ 生成成功！请点击下方的下载按钮。", state="complete", expanded=False)

        # 下载按钮放在 Form 外面，只要生成过就一直显示
        if "final_excel" in st.session_state:
            st.download_button(
                label="📥 第二步：点击下载纯净修复版 Excel",
                data=st.session_state.final_excel,
                file_name=f"Fixed_Coupon_{datetime.now().strftime('%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
