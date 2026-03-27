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
                req_p = None
                req_match = re.search(r'要求的(?:净价格|最高商品价格)：[^\d]*([\d\.]+)', content)
                if not req_match:
                    req_match = re.search(r'(?:Maximum product price allowed|Required net price)：[^\d]*([\d\.]+)', content)
                if req_match:
                    req_p = float(req_match.group(1))
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
    st.header("📂 数据上传区")
    site_template = st.file_uploader("上传：站点空白 Coupon 模板", type=['xlsx'], key="template_gen")
    all_listing_file = st.file_uploader("上传：All Listing 报告", type=['txt', 'csv', 'xlsx'], key="listing_rep")
    error_feedback_file = st.file_uploader("上传：Amazon 报错文件", type=['xlsx'], key="error_rep")
    
    st.divider()
    if error_feedback_file and all_listing_file:
        st.header("⚙️ 修复筛选配置")
        status_sel = st.multiselect("ASIN 状态筛选", ["✅ 正常", "❌ 批注报错"], default=["✅ 正常", "❌ 批注报错"])
        reason_kw = st.text_input("报错原因关键词过滤")

    if st.button("🔄 清空所有数据"):
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
            if not any(k in h_text for k in CALENDAR_KEYWORDS) and not any(k in h_text for k in MANUAL_KEYWORDS):
                opts = list(dict.fromkeys(filter(None, [str(ws_gen.cell(row=8, column=i).value), str(ws_gen.cell(row=9, column=i).value)])))
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
    if not all_listing_file or not error_feedback_file or not site_template:
        st.warning("👈 请确保左侧已上传：1.空白模板 2.All Listing报告 3.报错文件")
    else:
        # 1. 初始数据解析
        if 'master_df' not in st.session_state:
            with st.spinner("深度解析报错中..."):
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
                raw_h = [ws_err.cell(row=7, column=c).value for c in range(1, ws_err.max_column + 1)]
                h_map = {h: i for i, h in enumerate(raw_h) if h}
                asin_h = next((h for h in raw_h if h and 'ASIN' in str(h)), None)
                disc_h = next((h for h in raw_h if h and '折扣' in str(h) and '数值' in str(h)), None)
                rows = []
                for r_idx in range(10, ws_err.max_row + 1):
                    if not any([ws_err.cell(row=r_idx, column=c).value for c in range(1, ws_err.max_column + 1)]): continue
                    comm = ws_err.cell(row=r_idx, column=ws_err.max_column).comment.text if ws_err.cell(row=r_idx, column=ws_err.max_column).comment else ""
                    asin_val = ws_err.cell(row=r_idx, column=h_map[asin_h]+1).value
                    asins = [a.strip() for a in str(asin_val).replace(',', ';').replace('\n', ';').split(';') if a.strip()]
                    err_map = CouponProcessor.parse_error_details(comm)
                    for a in asins:
                        asin_col = next((c for c in df_l.columns if 'asin' in c), None)
                        p_col = next((c for c in df_l.columns if 'price' in c or '价格' in c), None)
                        p_match = df_l[df_l[asin_col] == a][p_col].values if asin_col else []
                        orig_p = p_match[0] if len(p_match) > 0 else 0
                        info = err_map.get(a, {})
                        curr_d = ws_err.cell(row=r_idx, column=h_map[disc_h]+1).value if disc_h in h_map else 0.05
                        suggested = curr_d
                        if info.get('req_price') and orig_p:
                            needed = math.ceil(((float(orig_p) - float(info['req_price'])) / float(orig_p)) * 100)
                            suggested = needed / 100 if float(curr_d or 0) < 1 else max(needed, 5)
                        rows.append({
                            "决策": info.get('default_decision', "保留"), "ASIN": a, 
                            "状态": "❌ 批注报错" if a in err_map else "✅ 正常",
                            "详细报错原因": info.get('reason', "-"), "要求净价格": info.get('req_price', "-"),
                            "拟提报折扣": suggested, "Listing原价": orig_p, "原始行号": r_idx
                        })
                st.session_state.master_df = pd.DataFrame(rows)

        # 2. 决策工作台
        mask = st.session_state.master_df['状态'].isin(status_sel)
        if reason_kw: mask = mask & st.session_state.master_df['详细报错原因'].str.contains(reason_kw, case=False)
        df_show = st.session_state.master_df[mask].copy()

        st.subheader("🛠️ 修复决策台")
        edited = st.data_editor(df_show, hide_index=True, use_container_width=True, key="fix_editor_final",
                                column_config={"决策": st.column_config.SelectboxColumn("决策", options=["保留", "剔除"]),
                                              "拟提报折扣": st.column_config.NumberColumn("拟提报折扣", format="%.2f"),
                                              "原始行号": None},
                                disabled=['ASIN', '状态', '详细报错原因', '要求净价格', 'Listing原价'])

        # 实时同步编辑结果
        if not edited.equals(df_show):
            for idx in edited.index:
                st.session_state.master_df.loc[idx, '决策'] = edited.loc[idx, '决策']
                st.session_state.master_df.loc[idx, '拟提报折扣'] = edited.loc[idx, '拟提报折扣']
            st.rerun()

        st.divider()

        # 3. 改进的导出逻辑（彻底解决无反应问题）
        exp_col1, exp_col2 = st.columns(2)
        
        with exp_col1:
            # 点击生成按钮
            if st.button("🚀 第一步：执行生成 (Generate)", use_container_width=True, type="primary"):
                try:
                    with st.status("正在构建修复版文件...", expanded=False) as status:
                        site_template.seek(0)
                        wb_final = openpyxl.load_workbook(site_template)
                        ws_final = wb_final.active
                        error_feedback_file.seek(0)
                        wb_err_ref = openpyxl.load_workbook(error_feedback_file, data_only=True)
                        ws_err_ref = wb_err_ref.active
                        
                        f_h = [ws_final.cell(row=7, column=c).value for c in range(1, ws_final.max_column + 1)]
                        a_idx = next((i for i, h in enumerate(f_h, 1) if h and 'ASIN' in str(h)), 1)
                        d_idx = next((i for i, h in enumerate(f_h, 1) if h and '折扣' in str(h) and '数值' in str(h)), 3)
                        
                        # 只导出“保留”的 ASIN
                        final_keep = st.session_state.master_df[st.session_state.master_df['决策'] == "保留"]
                        curr_row = 10
                        for (orig_l, disc), group in final_keep.groupby(['原始行号', '拟提报折扣']):
                            for c_idx in range(1, len(f_h) + 1):
                                val = ws_err_ref.cell(row=orig_l, column=c_idx).value
                                target = ws_final.cell(row=curr_row, column=c_idx, value=val)
                                ref = ws_final.cell(row=9, column=c_idx)
                                if ref.has_style:
                                    target.font, target.border = copy(ref.font), copy(ref.border)
                                    target.fill, target.alignment = copy(ref.fill), copy(ref.alignment)
                            
                            ws_final.cell(row=curr_row, column=a_idx).value = ";".join(group['ASIN'].tolist())
                            ws_final.cell(row=curr_row, column=d_idx).value = disc
                            curr_row += 1
                        
                        out_io = BytesIO()
                        wb_final.save(out_io)
                        st.session_state.final_blob = out_io.getvalue()
                        status.update(label="文件生成完毕！", state="complete")
                    st.toast("✅ 修复版 Excel 已就绪", icon="🔥")
                except Exception as e:
                    st.error(f"生成失败: {e}")

        with exp_col2:
            # 下载按钮：只要 final_blob 存在就始终显示
            if "final_blob" in st.session_state:
                st.download_button(
                    label="📥 第二步：点击下载文件 (Download)",
                    data=st.session_state.final_blob,
                    file_name=f"Fixed_Coupon_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.info("请先点击左侧【🚀 生成】按钮")
