# ========== 输出结果（含修复：安全复制样式） ==========
output_full = BytesIO()
wb.save(output_full)
output_full.seek(0)

# 生成精简错误表：复制值并为有颜色的单元格在新工作簿用新 PatternFill 重建颜色
wb_error = Workbook()
ws_err = wb_error.active
for i, col_name in enumerate(tc_df.columns, start=1):
    ws_err.cell(1, i, col_name)

row_idx = 2
for idx in sorted(error_rows):
    for j, val in enumerate(tc_df.iloc[idx], start=1):
        ws_err.cell(row_idx, j, val)
        orig_cell = ws.cell(idx + 2, j)
        fill = orig_cell.fill
        try:
            if hasattr(fill, "fill_type") and fill.fill_type not in (None, "none", ""):
                start = getattr(fill.start_color, "rgb", None) or getattr(fill.start_color, "index", None)
                end = getattr(fill.end_color, "rgb", None) or getattr(fill.end_color, "index", None)
                if start or end:
                    new_fill = PatternFill(fill_type=fill.fill_type, start_color=start, end_color=end)
                    ws_err.cell(row_idx, j).fill = new_fill
        except Exception:
            pass
    row_idx += 1

output_err = BytesIO()
wb_error.save(output_err)
output_err.seek(0)

# ========== 🔍 新增功能：反向检查漏填合同号 ==========
st.divider()
st.subheader("🔍 反向漏填检查：检测‘总’sheet是否遗漏其他表合同号")

# 收集各来源的合同号集合
def get_contracts_from_df(df):
    col = find_col(df, "合同", exact=False)
    if col is not None:
        return set(df[col].dropna().astype(str).str.strip())
    return set()

def get_contracts_from_fk_dfs(fk_list):
    all_cons = set()
    for df in fk_list:
        all_cons |= get_contracts_from_df(df)
    return all_cons

contracts_total = get_contracts_from_df(tc_df)
contracts_fk = get_contracts_from_fk_dfs(fk_dfs)
contracts_ec = get_contracts_from_df(ec_df)
contracts_ori = get_contracts_from_df(original_df)

# 合并所有来源
contracts_check_sources = contracts_fk | contracts_ec | contracts_ori

# 找出漏填合同号
missing_contracts = sorted(list(contracts_check_sources - contracts_total))

if missing_contracts:
    st.warning(f"⚠️ 发现 {len(missing_contracts)} 个合同号存在于‘放款明细’/‘二次明细’/‘原表’中，但未出现在‘总’sheet中")

    # 生成漏填表
    df_missing = pd.DataFrame({"漏填合同号": missing_contracts})
    output_missing = BytesIO()
    with pd.ExcelWriter(output_missing, engine="openpyxl") as writer:
        df_missing.to_excel(writer, index=False, sheet_name="漏填合同号")
    output_missing.seek(0)

    st.download_button(
        "📥 下载漏填合同号表",
        data=output_missing,
        file_name="提成_漏填合同号.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.success("✅ 未发现漏填合同号，‘总’sheet已覆盖所有检查来源。")

# ========== 下载区 ==========
st.divider()
st.subheader("📤 下载审核结果文件")

st.download_button(
    "📥 下载提成总sheet审核标注版",
    data=output_full,
    file_name="提成_总sheet_审核标注版.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.download_button(
    "📥 下载仅错误与标黄合同号精简版（含红黄标记）",
    data=output_err,
    file_name="提成_错误精简版_带颜色.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success(f"✅ 审核完成，共发现 {total_errors} 处错误，{len(error_rows)} 行合同涉及异常")
