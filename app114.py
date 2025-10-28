# =====================================
# Streamlit App: 提成表总sheet自动审核（标红错误格 + 标黄合同号 + 精简错误下载）
# 修复：复制颜色时为目标工作簿创建新的 PatternFill（避免跨工作簿复用样式对象）
# =====================================
import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata, re

try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill
except ImportError:
    st.error("❌ openpyxl 未安装，请执行 pip install openpyxl")
    st.stop()

st.title("📊 提成表『总』sheet 自动审核工具（标红错误格 + 标黄合同号 + 精简错误下载）")

# ========== 上传文件 ==========
uploaded_files = st.file_uploader(
    "请上传包含“提成”、“放款明细”、“二次明细”和“原表”的xlsx文件",
    type="xlsx", accept_multiple_files=True
)

if not uploaded_files or len(uploaded_files) < 4:
    st.warning("⚠️ 请至少上传 提成表、放款明细、二次明细、原表 四个文件")
    st.stop()

# ========== 工具函数 ==========
def find_file(files_list, keyword):
    for f in files_list:
        if keyword in f.name:
            return f
    return None

def normalize_text(val):
    if pd.isna(val):
        return ""
    s = str(val)
    s = re.sub(r'[\n\r\t ]+', '', s)
    s = s.replace('\u3000', '')
    return ''.join(unicodedata.normalize('NFKC', ch) for ch in s).lower().strip()

def normalize_num(val):
    if pd.isna(val):
        return None
    s = str(val).replace(",", "").replace("%", "").strip()
    if s in ["", "-", "nan"]:
        return None
    try:
        return float(s)
    except:
        return None

def find_col(df_like, keyword, exact=False):
    key = keyword.strip().lower()
    columns = df_like.columns if hasattr(df_like, "columns") else df_like.index
    for col in columns:
        cname = str(col).strip().lower()
        if (exact and cname == key) or (not exact and key in cname):
            return col
    return None

# ========== 读取文件 ==========
tc_file = find_file(uploaded_files, "提成")
fk_file = find_file(uploaded_files, "放款明细")
ec_file = find_file(uploaded_files, "二次明细")
original_file = find_file(uploaded_files, "原表")

tc_xls = pd.ExcelFile(tc_file)
sheet_total = next((s for s in tc_xls.sheet_names if "总" in s), None)
tc_df = pd.read_excel(tc_file, sheet_name=sheet_total)

fk_xls = pd.ExcelFile(fk_file)
fk_dfs = [pd.read_excel(fk_file, sheet_name=s) for s in fk_xls.sheet_names if "潮掣" in s]

ec_xls = pd.ExcelFile(ec_file)
ec_df = pd.concat([pd.read_excel(ec_file, sheet_name=s) for s in ec_xls.sheet_names], ignore_index=True)

original_xls = pd.ExcelFile(original_file)
original_df = pd.read_excel(original_xls)

st.success("✅ 文件读取完成")

# ========== 定义映射 ==========
MAPPING = {
    "放款日期": ("放款明细", "放款日期", 0, 1),
    "提报人员": ("放款明细", "提报人员", 0, 1),
    "城市经理": ("放款明细", "城市经理", 0, 1),
    "租赁本金": ("放款明细", "租赁本金", 0, 1),
    "收益率": ("放款明细", "xirr", 0.005, 1),
    "期限": ("放款明细", "租赁期限/年", 0.5, 12),
    "家访": ("放款明细", "家访", 0, 1),
    "人员类型": ("放款明细", "类型", 0, 1),
    "二次交接": ("二次明细", "出本流程时间", 0, 1),
}

contract_col_main = find_col(tc_df, "合同")

# ========== 比对辅助函数 ==========
def get_ref_row(contract_no, source_type):
    contract_no = str(contract_no).strip()
    if source_type == "放款明细":
        for df in fk_dfs:
            col = find_col(df, "合同")
            if col is None:
                continue
            res = df[df[col].astype(str).str.strip() == contract_no]
            if not res.empty:
                return res.iloc[0]
    elif source_type == "二次明细":
        col = find_col(ec_df, "合同")
        if col is not None:
            res = ec_df[ec_df[col].astype(str).str.strip() == contract_no]
            if not res.empty:
                return res.iloc[0]
    elif source_type == "原表":
        col = find_col(original_df, "合同", exact=False)
        if col is not None:
            res = original_df[original_df[col].astype(str).str.strip() == contract_no]
            if not res.empty:
                return res.iloc[0]
    return None

# ========== 执行审核 ==========
wb = Workbook()
ws = wb.active
for i, col_name in enumerate(tc_df.columns, start=1):
    ws.cell(1, i, col_name)

red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

total_errors = 0
error_rows = set()

progress = st.progress(0)
status = st.empty()

person_type_col = find_col(tc_df, "人员类型", exact=True)
n = len(tc_df)

for idx, row in tc_df.iterrows():
    contract_no = row.get(contract_col_main)
    if pd.isna(contract_no):
        continue

    row_has_error = False

    for main_kw, (src, ref_kw, tol, mult) in MAPPING.items():
        exact_main = "期限" in main_kw or main_kw == "人员类型"
        main_col = find_col(tc_df, main_kw, exact=exact_main)
        if not main_col:
            continue

        if main_kw == "收益率":
            person_type = str(row[person_type_col]).strip()
            if person_type == "轻卡":
                ref_row = get_ref_row(contract_no, "原表")
                ref_kw = "年化NIM"
            else:
                ref_row = get_ref_row(contract_no, src)
        else:
            ref_row = get_ref_row(contract_no, src)

        if ref_row is None:
            continue

        ref_col = find_col(ref_row, ref_kw, exact=(main_kw == "人员类型"))
        if not ref_col:
            continue

        main_val = row[main_col]
        ref_val = ref_row[ref_col]

        if "日期" in main_kw or main_kw == "二次交接":
            try:
                main_dt = pd.to_datetime(main_val, errors='coerce').normalize()
                ref_dt = pd.to_datetime(ref_val, errors='coerce').normalize()
            except:
                main_dt = ref_dt = pd.NaT
            if pd.isna(main_dt) or pd.isna(ref_dt) or main_dt != ref_dt:
                row_has_error = True
                total_errors += 1
                ws.cell(idx + 2, list(tc_df.columns).index(main_col) + 1).fill = red_fill
        else:
            m = normalize_num(main_val)
            r = normalize_num(ref_val)
            if main_kw == "收益率" and m is not None and r is not None:
                if m > 1: m /= 100
                if r > 1: r /= 100
            if m is not None and r is not None:
                if "期限" in main_kw:
                    r *= mult
                if abs(m - r) > tol:
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
        error_rows.add(idx)

    for j, val in enumerate(row, start=1):
        ws.cell(idx + 2, j, val)

    if (idx + 1) % 10 == 0 or (idx + 1) == n:
        progress.progress((idx + 1) / n)
        status.text(f"审核进度：{idx + 1}/{n}")

# ========== 输出结果（含修复：安全复制样式） ==========
# 注意：只有在这里——所有工作簿已写入完成之后——才创建 BytesIO 并准备下载内容

# 先把完整带标注的工作簿保存到内存
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
        # 写值
        ws_err.cell(row_idx, j, val)
        # 从主工作表读取原始单元格（注意：ws 是原工作簿的sheet）
        try:
            orig_cell = ws.cell(idx + 2, j)
            fill = orig_cell.fill
            if hasattr(fill, "fill_type") and fill.fill_type not in (None, "none", ""):
                # 尝试取出 start_color / end_color 的 rgb 或 index
                start = getattr(fill.start_color, "rgb", None) or getattr(fill.start_color, "index", None)
                end = getattr(fill.end_color, "rgb", None) or getattr(fill.end_color, "index", None)
                if start or end:
                    new_fill = PatternFill(fill_type=fill.fill_type, start_color=start, end_color=end)
                    ws_err.cell(row_idx, j).fill = new_fill
        except Exception:
            # 容错：若读取原单元格或复制样式失败，只保留值，不阻断流程
            pass
    row_idx += 1

output_err = BytesIO()
wb_error.save(output_err)
output_err.seek(0)

# ========== 🔍 反向漏填检查（仅使用放款明细中包含“潮掣”的sheet） ==========
st.divider()
st.subheader("🔍 反向漏填检查（仅基于放款明细中包含“潮掣”的sheet）")

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
# 只使用 fk_dfs（你之前只读取了包含“潮掣”的 sheets）
contracts_fk = get_contracts_from_fk_dfs(fk_dfs)

# 只对放款明细（潮掣 sheets）中的合同号进行反向检查
missing_contracts = sorted(list(contracts_fk - contracts_total))

if missing_contracts:
    st.warning(f"⚠️ 发现 {len(missing_contracts)} 个合同号存在于放款明细（含“潮掣”的sheet）中，但未出现在提成表的‘总’sheet中")
    df_missing = pd.DataFrame({"漏填合同号": missing_contracts})
    output_missing = BytesIO()
    with pd.ExcelWriter(output_missing, engine="openpyxl") as writer:
        df_missing.to_excel(writer, index=False, sheet_name="漏填合同号")
    output_missing.seek(0)
    st.download_button(
        "📥 下载漏填合同号表（基于放款明细-潮掣）",
        data=output_missing,
        file_name="提成_漏填合同号_基于放款明细_潮掣.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.success("✅ 未发现漏填合同号（基于放款明细-潮掣）。")

# ========== 下载区 ==========
st.divider()
st.subheader("📤 下载审核结果文件")

# 仅当流非空时才显示下载按钮（防止提前创建或变量未定义）
if output_full is not None:
    st.download_button(
        "📥 下载提成总sheet审核标注版",
        data=output_full,
        file_name="提成_总sheet_审核标注版.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if error_rows and output_err is not None:
    st.download_button(
        "📥 下载仅错误与标黄合同号精简版（含红黄标记）",
        data=output_err,
        file_name="提成_错误精简版_带颜色.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.success(f"✅ 审核完成，共发现 {total_errors} 处错误，{len(error_rows)} 行合同涉及异常")
