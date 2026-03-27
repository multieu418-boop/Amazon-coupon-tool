import streamlit as st
import pandas as pd
import openpyxl
from copy import copy
import re
import io
import math
from datetime import datetime

# --- 页面基础配置 ---
st.set_page_config(page_title="Cupshe Amazon Coupon 自动化修复系统", layout="wide")

# --- 1. 核心解析与导出逻辑 (使用你提供的无损导出方案) ---
def parse_error_details(comment_text):
    error_map = {}
    if not comment_text: return error_map
    # 支持中文和英文关键词提取
    blocks = re.split(r'([A-Z0-9]{10})\n', str(comment_text))
    if len(blocks) > 1:
        for i in range(1, len(blocks), 2):
            asin = blocks[i].strip()
            content = blocks[i+1]
            req_p_match = re.search(r'(?:要求的(?:净价格|最高商品价格)|Required net price|Maximum product price allowed)：?\s*[^\d]*([\d\.]+)', content)
            req_p = float(req_p_match.group(1)) if req_p_match else None
            reason_part = re.split(r'(?:要求的|当前|Maximum|Required)', content)[0]
            reason = reason_part.strip().replace('\n', ' ')
            auto_exclude = "没有经验证的参考价" in reason
            error_map[asin] = {"req_price": req_p, "reason": reason, "default_decision": "剔除" if auto_exclude else "保留"}
    return error_map

def generate_excel_lossless(e_file, master_df, orig_headers):
    """你提供的无损导出函数"""
    e_file.seek(0)
    wb = openpyxl.load_workbook(e_file)
    ws = wb.active
    
    row_data_backup = {}
    for r_idx in master_df['原始行号'].unique():
        row_cells = [ws.cell(row=r_idx, column=c).value for c in range(1, ws.max_column + 1)]
        row_data_backup[r_idx] = row_cells

    # 清空第10行以后
    if ws.max_row >= 10:
        for r in range(10, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                ws.cell(row=r, column=c).value = None

    final_keep = master_df[master_df['决策'] == "保留"]
    if final_keep.empty: return None

    # 定位列
    a_idx, d_idx = 1, 3
    for i, h in enumerate(orig_headers, 1):
        if h and 'ASIN' in str(h): a_idx = i
        if h and '折扣' in str(h) and '数值' in str(h): d_idx = i

    curr_r = 10
    for (orig_line, disc), group in final_keep.groupby(['原始行号', '拟提报折扣']):
        orig_row_values = row_data_backup.get(orig_line)
        if orig_row_values:
            for c_idx, val in enumerate(orig_row_values, 1):
                target_cell = ws.cell(row=curr_r, column=c_idx)
                target_cell.value = val
                source_cell = ws.cell(row=orig_line, column=c_idx)
                if source_cell.has_style:
                    target_cell.font, target_cell.border = copy(source_cell.font), copy(source_cell.border)
                    target_cell.fill, target_cell.alignment = copy(source_cell.fill), copy(source_cell.alignment)
                    target_cell.number_format = copy(source_cell.number_format)
        
        ws.cell(row=curr_r, column=a_idx).value = ";".join(group['ASIN'].tolist())
        ws.cell(row=curr_r, column=d_idx).value = disc
        curr_r += 1

    if ws.max_row >= curr_r:
        ws.delete_rows(curr_r, ws.max_row - curr_r + 1)
    
    out_io = io.BytesIO()
    wb.save(out_io)
    return out_io.getvalue()

# --- 2. 主程序界面 ---
st.title("🎯 Cupshe Amazon Coupon 自动化管理系统")

# 初始化数据容器
if 'master_df' not in st.session_state:
    st.session_state.master_df = None
if 'orig_headers' not in st.session_state:
    st.session_state.orig_headers = None

# --- 侧边栏 ---
with st.sidebar:
    st.header("⚙️ 筛选与重置")
    status_sel = st.multiselect("ASIN 状态筛选", ["✅ 正常", "❌ 批注报错"], default=["✅ 正常", "❌ 批注报错"])
    reason_kw = st.text_input("报错原因关键词过滤")
    if st.button("🔄 重置并清空所有上传"):
        st.session_state.clear()
        st.rerun()

# --- 文件上传区 ---
col_u1, col_u2 = st.columns(2)
with col_u1:
    l_file = st.file_uploader("1. 上传 All Listing 报告", type=['txt', 'csv', 'xlsx'])
with col_u2:
    e_file = st.file_uploader("2. 上传带批注的报错文件", type=['xlsx'])

# --- 核心逻辑运行 ---
if l_file and e_file:
    # 仅在初次上传或重置后解析数据
    if st.session_state.master_df is None:
        with st.spinner("数据解析中..."):
            # 解析 Listing
            for enc in ['utf-8', 'utf-16', 'gbk', 'utf-8-sig']:
                try:
                    l_file.seek(0)
                    df_l = pd.read_csv(l_file, sep='\t', encoding=enc) if l_file.name.endswith('.txt') else pd.read_excel(l_file)
                    df_l.columns = [c.lower().strip() for c in df_l.columns]
                    break
                except: continue
            
            # 解析报错模板
            e_file.seek(0)
            wb = openpyxl.load_workbook(e_file, data_only=True)
            ws = wb.active
            headers = [cell.value for cell in ws[7]]
            
            asin_col = next((c for c in df_l.columns if 'asin' in c), None)
            price_col = next((c for c in df_l.columns if 'price' in c or '价格' in c), None)
            e_asin_col = next((c for c in headers if 'ASIN' in str(c)), headers[0])
            e_disc_col = next((c for c in headers if '折扣' in str(c) and '数值' in str(h)), None)

            rows = []
            for r_idx, row in enumerate(ws.iter_rows(min_row=10), 10):
                vals = [cell.value for cell in row]
                if not any(vals): continue
                comment = row[-1].comment.text if row[-1].comment else ""
                row_dict = {headers[i]: v for i, v in enumerate(vals) if i < len(headers)}
                asins = [a.strip() for a in str(row_dict.get(e_asin_col, "")).replace(',', ';').split(';') if a.strip()]
                err_map = parse_error_details(comment)
                
                for a in asins:
                    p_match = df_l[df_l[asin_col] == a][price_col].values if asin_col else []
                    orig_p = p_match[0] if len(p_match) > 0 else 0
                    info = err_map.get(a, {})
                    is_err = a in err_map
                    curr_d = row_dict.get(e_disc_col, 0.05)
                    suggested = curr_d
                    if is_err and orig_p and info.get('req_price'):
                        needed = math.ceil(((float(orig_p) - float(info.get('req_price'))) / float(orig_p)) * 100)
                        suggested = needed / 100 if curr_d < 1 else needed
                    
                    rows.append({
                        "决策": info.get('default_decision', "保留"), "ASIN": a, 
                        "状态": "❌ 批注报错" if is_err else "✅ 正常",
                        "详细报错原因": info.get('reason', "-"), "拟提报折扣": suggested,
                        "Listing原价": orig_p, "要求净价": info.get('req_price'), "原始行号": r_idx
                    })
            st.session_state.master_df = pd.DataFrame(rows)
            st.session_state.orig_headers = headers

    # --- 渲染决策台 ---
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
                "拟提报折扣": st.column_config.NumberColumn("提报折扣", format="%.2f"),
                "原始行号": None
            },
            disabled=['ASIN', '状态', '详细报错原因', 'Listing原价', '要求净价'],
            hide_index=True, use_container_width=True, key="editor_integrated"
        )

        # 只要编辑器有变动，同步回 master_df
        if not edited.equals(df_filtered):
            for idx in edited.index:
                st.session_state.master_df.loc[idx, '决策'] = edited.loc[idx, '决策']
                st.session_state.master_df.loc[idx, '拟提报折扣'] = edited.loc[idx, '拟提报折扣']
            st.rerun()

        # --- 生成与导出区 (不再嵌套在 if edited 逻辑内，确保按钮永远在) ---
        st.markdown("---")
        st.subheader("📦 导出纯净修复文件")
        
        if st.button("🚀 生成并导出完整信息 Excel", use_container_width=True, type="primary"):
            with st.spinner("正在基于空白底稿构建文件..."):
                file_data = generate_excel_lossless(e_file, st.session_state.master_df, st.session_state.orig_headers)
                if file_data:
                    st.success("✅ 文件已成功生成！")
                    st.download_button(
                        label="📥 点击下载修复后的完整 Excel",
                        data=file_data,
                        file_name=f"Fixed_Coupon_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.error("没有可导出的项（可能所有决策都被改成了剔除）。")

else:
    st.info("💡 请先上传 Listing 报告和报错文件。")
