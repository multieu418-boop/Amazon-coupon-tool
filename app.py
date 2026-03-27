import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
import re
from io import BytesIO
import math

# --- 全局配置 ---
st.set_page_config(page_title="Cupshe Coupon 工具箱", layout="wide")

class CouponLogic:
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
                # 核心逻辑：自动识别剔除项
                auto_exclude = "没有经验证的参考价" in reason
                error_map[asin] = {
                    "req_price": req_p, 
                    "reason": reason, 
                    "default_decision": "剔除" if auto_exclude else "保留"
                }
        return error_map

# --- 侧边栏：统一上传区 ---
with st.sidebar:
    st.header("📂 数据源配置")
    all_listing = st.file_uploader("1. All Listing 报告", type=['txt', 'csv', 'xlsx'])
    template_file = st.file_uploader("2. Coupon 模板 (含报错或空白)", type=['xlsx'])
    st.divider()
    if st.button("🔄 重置系统（清空所有数据）"):
        st.session_state.clear()
        st.rerun()

st.title("🎯 Cupshe Amazon Coupon 综合管理系统")

tab1, tab2 = st.tabs(["🔵 阶段 1：生成新提报", "🔴 阶段 2：报错解析修复"])

# --- 阶段 1：提报生成 ---
with tab1:
    if not template_file:
        st.info("请在左侧上传空白模板以开启提报生成功能。")
    else:
        # 复用第一阶段逻辑
        wb_t = openpyxl.load_workbook(template_file, data_only=True)
        ws_t = wb_t.active
        headers = [cell.value for cell in ws_t[7] if cell.value]
        
        st.subheader("📝 录入提报需求")
        with st.form("gen_form"):
            user_inputs = {}
            c1, c2 = st.columns(2)
            for i, h in enumerate(headers):
                target = c1 if i % 2 == 0 else c2
                if "ASIN" in str(h).upper():
                    raw = target.text_area(f"{h}", placeholder="粘贴 ASIN 列表...")
                    user_inputs[h] = CouponLogic.clean_asin_input(raw)
                else:
                    user_inputs[h] = target.text_input(f"{h}")
            
            if st.form_submit_button("🚀 生成提报文件"):
                target_row = 10
                while ws_t.cell(row=target_row, column=1).value: target_row += 1
                for idx, h in enumerate(headers, 1):
                    cell = ws_t.cell(row=target_row, column=idx, value=user_inputs[h])
                    source = ws_t.cell(row=9, column=idx)
                    if source.has_style:
                        cell.font, cell.border, cell.fill, cell.alignment = copy(source.font), copy(source.border), copy(source.fill), copy(source.alignment)
                
                out = BytesIO()
                wb_t.save(out)
                st.download_button("📥 下载提报文件", out.getvalue(), "New_Coupon.xlsx")

# --- 阶段 2：报错修复 ---
with tab2:
    if not all_listing or not template_file:
        st.info("请在左侧上传 All Listing 和 带批注的报错模板。")
    else:
        # 初始化处理逻辑 (保持多人独立使用)
        if 'repair_df' not in st.session_state:
            with st.spinner("正在交叉比对数据..."):
                # 读取 Listing
                for enc in ['utf-8', 'utf-16', 'gbk', 'utf-8-sig']:
                    try:
                        all_listing.seek(0)
                        df_l = pd.read_csv(all_listing, sep='\t', encoding=enc) if all_listing.name.endswith('.txt') else pd.read_excel(all_listing)
                        df_l.columns = [c.lower().strip() for c in df_l.columns]
                        break
                    except: continue
                
                # 读取模板
                template_file.seek(0)
                wb_err = openpyxl.load_workbook(template_file, data_only=True)
                ws_err = wb_err.active
                err_headers = [cell.value for cell in ws_err[7]]
                
                asin_col = next((c for c in df_l.columns if 'asin' in c), None)
                price_col = next((c for c in df_l.columns if 'price' in c or '价格' in c), None)
                e_asin_idx = next((i for i, h in enumerate(err_headers) if h and 'ASIN' in str(h)), 0)
                e_disc_idx = next((i for i, h in enumerate(err_headers) if h and '折扣' in str(h) and '数值' in str(h)), 2)

                rows = []
                for r_idx in range(10, ws_err.max_row + 1):
                    asin_raw = ws_err.cell(row=r_idx, column=e_asin_idx+1).value
                    if not asin_raw: continue
                    comment = ws_err.cell(row=r_idx, column=ws_err.max_column).comment
                    comment_text = comment.text if comment else ""
                    err_map = CouponLogic.parse_error_details(comment_text)
                    
                    asins = [a.strip() for a in str(asin_raw).replace(',', ';').split(';') if a.strip()]
                    for a in asins:
                        p_match = df_l[df_l[asin_col] == a][price_col].values if asin_col else []
                        orig_p = p_match[0] if len(p_match) > 0 else 0
                        info = err_map.get(a, {})
                        
                        # 建议折扣计算
                        curr_d = ws_err.cell(row=r_idx, column=e_disc_idx+1).value or 0.05
                        suggested = curr_d
                        if a in err_map and orig_p and info.get('req_price'):
                            needed = math.ceil(((float(orig_p) - float(info.get('req_price'))) / float(orig_p)) * 100)
                            suggested = needed / 100 if curr_d < 1 else needed

                        rows.append({
                            "决策": info.get('default_decision', "保留"),
                            "ASIN": a, "状态": "❌ 报错" if a in err_map else "✅ 正常",
                            "报错原因": info.get('reason', "-"), "拟提报折扣": suggested,
                            "Listing原价": orig_p, "原始行号": r_idx
                        })
                st.session_state.repair_df = pd.DataFrame(rows)
                st.session_state.orig_headers = err_headers

        if st.session_state.repair_df is not None:
            # 渲染编辑器
            edited = st.data_editor(
                st.session_state.repair_df,
                column_config={"原始行号": None, "决策": st.column_config.SelectboxColumn("决策", options=["保留", "剔除"])},
                disabled=["ASIN", "状态", "报错原因", "Listing原价"],
                hide_index=True, use_container_width=True, key="repair_editor"
            )

            if not edited.equals(st.session_state.repair_df):
                st.session_state.repair_df = edited
                st.rerun()

            if st.button("🚀 导出修复后的完整 Excel", use_container_width=True):
                # 此处执行全行复制逻辑 (逻辑同上一版，略)
                st.write("正在生成文件...")
                # (建议将上一版中的 generate_excel 函数放入并调用)
