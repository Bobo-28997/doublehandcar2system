# =====================================
# Streamlit App: 提成表总sheet自动审核（标红错误格 + 标黄合同号）
# =====================================
import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata, re

# 安全导入 openpyxl
try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill
except ImportError:
    st.error("❌ openpyxl 未安装，请执行 pip install openpyxl")
    st.stop()

# -------------------------------
# 页面标题
# -------------------------------
st.title("📊 提成表『总』sheet 自动审核工具（标红错误格 + 标黄合同号）")

# =====================================
# 一、上传文件
# =====================================
uploaded_files = st.file_uploader(
    "请上传包含“提成”、“放款明细”、“二次明细”和“原表”的xlsx文件",
    type="xlsx", accept_multiple_files=True
)

if not uploaded_files or len(uploaded_files) < 4:
    st.warning("⚠️ 请至少上传 提成表、放款明细、二次明细、原表 四个文件")
    st.stop()
else:
    st.success("✅ 文件上传完成")

# =====================================
# 二、工具函数
# =====================================
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
    """在DataFrame中查找包含关键字的列"""
    key = keyword.strip().lower()
    columns = df_like.columns if hasattr(df_like, "columns") else df_like.index
    for col in columns:
        cname = str(col).strip().lower()
        if (exact and cname == key) or (not exact and key in cname):
            return col
    return None

# =====================================
# 三、读取文件
# =====================================
tc_file = find_file(uploaded_files, "提成")
fk_file = find_file(uploaded_files, "放款明细")
ec_file = find_file(uploaded_files, "二次明细")
original_file = find_file(uploaded_files, "原表")

if not (tc_file and fk_file and ec_file and original_file):
    st.error("❌ 文件缺失，请确保文件名中包含 提成、放款明细、二次明细、原表")
    st.stop()

# 提成表：自动选择含“总”的sheet
tc_xls = pd.ExcelFile(tc_file)
sheet_total = next((s for s in tc_xls.sheet_names if "总" in s), None)
if sheet_total is None:
    st.error("❌ 提成文件中未找到包含“总”的sheet")
    st.stop()
tc_df = pd.read_excel(tc_file, sheet_name=sheet_total)

# 放款明细：仅取包含“潮掣”的sheet
fk_xls = pd.ExcelFile(fk_file)
fk_sheets = [s for s in fk_xls.sheet_names if "潮掣" in s]
if not fk_sheets:
    st.error("❌ 放款明细文件中未找到包含“潮掣”的sheet")
    st.stop()
fk_dfs = [pd.read_excel(fk_file, sheet_name=s) for s in fk_sheets]

# 二次明细：合并所有sheet
ec_xls = pd.ExcelFile(ec_file)
ec_df_list = [pd.read_excel(ec_file, sheet_name=s) for s in ec_xls.sheet_names]
ec_df = pd.concat(ec_df_list, ignore_index=True)

original_xls = pd.ExcelFile(original_file)
original_df = pd.read_excel(original_xls)

st.success(f"✅ 文件读取完成，提成总sheet {len(tc_df)} 行，原表 {len(original_df)} 行")

# =====================================
# 四、字段映射定义
# =====================================
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
if not contract_col_main:
    st.error("❌ 提成总sheet 未找到合同号列")
    st.stop()

# =====================================
# 五、主比对函数（含原表模糊匹配）
# =====================================
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
        col = find_col(original_df, "合同", exact=False)  # 模糊匹配“合同”
        if col is not None:
            res = original_df[original_df[col].astype(str).str.strip() == contract_no]
            if not res.empty:
                return res.iloc[0]
    return None

# =====================================
# 六、执行审核
# =====================================
try:
    wb = Workbook()
except Exception as e:
    st.error(f"❌ Workbook 初始化失败: {e}")
    st.stop()

ws = wb.active
for i, col_name in enumerate(tc_df.columns, start=1):
    ws.cell(1, i, col_name)

red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

total_errors = 0
n = len(tc_df)
progress = st.progress(0)
status = st.empty()

person_type_col = find_col(tc_df, "人员类型", exact=True)

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

        # 判断使用哪个参考表
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

        # 日期比对
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

    if row_has_error:
        ws.cell(idx + 2, list(tc_df.columns).index(contract_col_main) + 1).fill = yellow_fill

    for j, val in enumerate(row, start=1):
        ws.cell(idx + 2, j, val)

    if (idx + 1) % 10 == 0 or (idx + 1) == n:
        progress.progress((idx + 1) / n)
        status.text(f"审核进度：{idx + 1}/{n}")

# =====================================
# 七、输出结果
# =====================================
output = BytesIO()
wb.save(output)
output.seek(0)

st.download_button(
    "📥 下载提成总sheet审核标注版",
    data=output,
    file_name="提成_总sheet_审核标注版.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success(f"✅ 审核完成，共发现 {total_errors} 处错误")
