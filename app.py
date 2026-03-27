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

# 初始化 Session State 状态锁
if "reset_key" not in st.session_state:
    st.session_state.reset_key = 0
if "master_df" not in st.session_state:
    st.session_state.master_df = None

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
                req_match = re.search(r'(?:要求的(?:净价格|最高商品价格)|Required net price|Maximum product price allowed)：?\s*[^\d]*([\d\.]+)', content)
                req_p = float(req_match.group(1)) if req_match else None
                reason_part = re.split(r'(?:要求的|当前|Maximum|Required)', content)[0]
                reason = reason_part.strip().replace('\n', ' ')
                auto_exclude = "没有经验证的参考价" in reason
                error_map[asin] = {"req_price": req_p, "reason": reason, "default_decision": "剔除" if auto_exclude else "保留"}
        return error_map

# --- 侧边栏 ---
with st.sidebar:
    st.header("📂 核心文件上传")
    rk = st.session_state.reset_key

    site_template = st.file_uploader("1. 空白模板 (底稿)", type=['xlsx'], key=f"tpl_{rk}")
    all_listing_file = st.file_uploader("2. Listing 报告", type=['txt', 'csv', 'xlsx'], key=f"list_{rk}")
    error_feedback_file = st.file_uploader("3. 亚马逊报错文件", type=['xlsx'], key=f"err_{rk}")
    
    st.divider()

    if st.button("🗑️ 清空所有上传", use_container_width=True, type="secondary"):
        st.session_state.reset_key += 1
        st.session_state.master_df = None

        # 清空所有生成结果
        for key in ["final_excel", "gen_file", "main_editor_v4"]:
            if key in st.session_state:
                del st.session_state[key]

        st.toast("✅ 已清空所有上传内容", icon="🧹")
        st.success("页面已重置")
        st.rerun()

st.title("🎯 Cupshe Amazon Coupon 自动化管理系统")
tab1, tab2 = st.tabs(["🔵 第一阶段：提报生成", "🔴 第二阶段：报错修复"])

# --- 第一阶段 ---
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
            st.download_button("📥 下载初始提报", st.session_state.gen_file, "Initial.xlsx")

# --- 第二阶段 ---
with tab2:
    if not all_listing_file or not error_feedback_file or not site_template:
        st.info("💡 请在左侧上传【空白模板】、【Listing报告】和【报错文件】以开启修复功能。")
    else:
        # 数据解析
        if st.session_state.master_df is None:
            with st.spinner("正在解析数据..."):
                for enc in ['utf-8', 'utf-16', 'gbk', 'utf-8-sig']:
                    try:
                        all_listing_file.seek(0)
                        df_l = pd.read_csv(all_listing_file, sep='\t', encoding=enc) if all_listing_file.name.endswith('.txt') else pd.read_excel(all_listing_file)
                        df_l.columns = [c.lower().strip() for c in df_l.columns]; break
                    except: continue

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
                        a_col = next((c for c in df_l.columns if 'asin' in c), None)
                        p_col = next((c for c in df_l.columns if 'price' in c or '价格' in c), None)
                        p_match = df_l[df_l[a_col] == a][p_col].values if a_col else []
                        orig_p = p_match[0] if len(p_match) > 0 else 0

                        info = err_map.get(a, {})
                        curr_d = ws_err.cell(row=r, column=d_idx+1).value or 0.05
                        suggested = curr_d

                        if info.get('req_price') and orig_p:
                            needed = math.ceil(((float(orig_p) - float(info['req_price'])) / float(orig_p)) * 100)
                            suggested = needed / 100 if float(curr_d) < 1 else max(needed, 5)

                        rows.append({
                            "决策": info.get('default_decision', "保留"),
                            "ASIN": a,
                            "状态": "❌ 批注报错" if a in err_map else "✅ 正常",
                            "原因": info.get('reason', "-"),
                            "要求净价格": info.get('req_price', "-"),
                            "拟提报折扣": suggested,
                            "Listing原价": orig_p,
                            "原始行号": r
                        })

                st.session_state.master_df = pd.DataFrame(rows)

        # 表格
        st.subheader("🛠️ 修复决策台")
        edited_df = st.data_editor(
            st.session_state.master_df,
            column_config={"决策": st.column_config.SelectboxColumn("决策", options=["保留", "剔除"]), "原始行号": None},
            disabled=['ASIN', '状态', '原因', '要求净价格', 'Listing原价'],
            hide_index=True,
            use_container_width=True,
            key="main_editor_v4"
        )
        st.session_state.master_df = edited_df

        st.divider()

        # ✅ 强制显示生成区（核心修复）
        st.subheader("📦 纯净导出区")

        col_btn, col_msg = st.columns([1, 2])

        if col_btn.button("🚀 执行生成", use_container_width=True, type="primary"):
            msg_placeholder = col_msg.empty()

            try:
                msg_placeholder.warning("⏳ 正在生成，请稍候...")

                site_template.seek(0)
                wb_final = openpyxl.load_workbook(site_template)
                ws_final = wb_final.active

                error_feedback_file.seek(0)
                wb_err_ref = openpyxl.load_workbook(error_feedback_file, data_only=True)
                ws_err_ref = wb_err_ref.active

                f_h = [ws_final.cell(row=7, column=c).value for c in range(1, ws_final.max_column + 1)]
                f_a_idx = next((i for i, h in enumerate(f_h, 1) if h and 'ASIN' in str(h)), 1)
                f_d_idx = next((i for i, h in enumerate(f_h, 1) if h and '折扣' in str(h) and '数值' in str(h)), 3)

                final_keep = st.session_state.master_df[st.session_state.master_df['决策'] == "保留"]

                curr_row = 10
                for (orig_l, disc), group in final_keep.groupby(['原始行号', '拟提报折扣']):
                    for c_idx in range(1, len(f_h) + 1):
                        val = ws_err_ref.cell(row=orig_l, column=c_idx).value
                        ws_final.cell(row=curr_row, column=c_idx, value=val)

                    ws_final.cell(row=curr_row, column=f_a_idx).value = ";".join(group['ASIN'].tolist())
                    ws_final.cell(row=curr_row, column=f_d_idx).value = disc
                    curr_row += 1

                out_fix = io.BytesIO()
                wb_final.save(out_fix)

                st.session_state.final_excel = out_fix.getvalue()

                msg_placeholder.success("✅ 已生成成功，下面可以下载")

            except Exception as e:
                msg_placeholder.error(f"❌ 生成失败: {str(e)}")

        # 下载按钮（稳定显示）
        if "final_excel" in st.session_state:
            st.download_button(
                label="📥 下载修复后的Excel",
                data=st.session_state.final_excel,
                file_name=f"Fixed_{datetime.now().strftime('%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
