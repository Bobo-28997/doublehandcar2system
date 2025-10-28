# =====================================
# 六、执行审核（稳健版 + 优化进度条）
# =====================================
wb = Workbook()
ws = wb.active
for i, col_name in enumerate(tc_df.columns, start=1):
    ws.cell(1, i, col_name)

red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

total_errors = 0
n = len(tc_df)
progress = st.progress(0)
status = st.empty()

for idx, row in tc_df.iterrows():
    contract_no = row.get(contract_col_main)
    if pd.isna(contract_no):
        continue

    row_has_error = False

    for main_kw, (src, ref_kw, tol, mult) in MAPPING.items():
        main_col = find_col(tc_df, main_kw, exact=("期限" in main_kw))
        if not main_col:
            continue

        ref_row = get_ref_row(contract_no, src)
        if ref_row is None:
            continue

        ref_col = find_col(ref_row, ref_kw)
        if not ref_col:
            continue

        main_val = row[main_col]
        ref_val = ref_row[ref_col]

        # ✅ 日期字段比对：只比对年月日
        if "日期" in main_kw or main_kw == "二次交接":
            try:
                main_dt = pd.to_datetime(main_val, errors='coerce').normalize()
                ref_dt = pd.to_datetime(ref_val, errors='coerce').normalize()
            except:
                main_dt = ref_dt = pd.NaT

            if pd.isna(main_dt) or pd.isna(ref_dt) or main_dt != ref_dt:
                row_has_error = True
                ws.cell(idx + 2, list(tc_df.columns).index(main_col) + 1).fill = red_fill
                total_errors += 1

        # 数值字段比对
        elif isinstance(normalize_num(main_val), (int, float)) or isinstance(normalize_num(ref_val), (int, float)):
            m = normalize_num(main_val)
            r = normalize_num(ref_val)
            if m is not None and r is not None:
                if "期限" in main_kw:
                    r = r * mult
                    if abs(m - r) > tol:
                        row_has_error = True
                        total_errors += 1
                        ws.cell(idx + 2, list(tc_df.columns).index(main_col) + 1).fill = red_fill
                else:
                    if abs(m - r) > tol:
                        row_has_error = True
                        total_errors += 1
                        ws.cell(idx + 2, list(tc_df.columns).index(main_col) + 1).fill = red_fill
            else:
                if normalize_text(main_val) != normalize_text(ref_val):
                    row_has_error = True
                    total_errors += 1
                    ws.cell(idx + 2, list(tc_df.columns).index(main_col) + 1).fill = red_fill
        else:
            if normalize_text(main_val) != normalize_text(ref_val):
                row_has_error = True
                total_errors += 1
                ws.cell(idx + 2, list(tc_df.columns).index(main_col) + 1).fill = red_fill

    # 标黄合同号
    if row_has_error:
        ws.cell(idx + 2, list(tc_df.columns).index(contract_col_main) + 1).fill = yellow_fill

    # 写入原数据
    for j, val in enumerate(row, start=1):
        ws.cell(idx + 2, j, val)

    # 更新进度条
    progress.progress((idx + 1) / n)
    if (idx + 1) % 10 == 0 or (idx + 1) == n:
        status.text(f"审核进度：{idx + 1}/{n}")
