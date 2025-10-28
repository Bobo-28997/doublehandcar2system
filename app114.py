# =====================================
# Streamlit App: “提成表总sheet自动审核（标红错误格 + 标黄合同号）”
# =====================================
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from io import BytesIO
import unicodedata, re

# -------------------------------
# 页面标题
# -------------------------------
st.title("📊 提成表『总』sheet 自动审核工具（标红错误格 + 标黄合同号）")


# =====================================
# 一、上传文件
# =====================================
uploaded_files = st.file_uploader(
    "请上传包含“提成”、“放款明细”、“二次明细”的xlsx文件",
    type="xlsx", accept_multiple_files=True
)

if not uploaded_files or len(uploaded_files) < 3:
    st.warning("⚠️ 请至少上传 提成表、放款明细、二次明细 三个文件")
    st.stop()
else:
    st.success("✅ 文件上传完成")


# =====================================
# 二、工具函数
# =====================================
def find_file(files_list, keyword):
    """通过关键词找到文件"""
    for f in files_list:
        if keyword in f.name:
            return f
    return None

def normalize_text(val):
    """标准化文本"""
    if pd.isna(val):
        return ""
    s = str(val)
    s = re.sub(r'[\n\r\t ]+', '', s)
    s = s.replace('\u3000', '')
    s = ''.join(unicodedata.normalize('NFKC', ch) for ch in s)
    return s.lower().strip()

def normalize_num(val):
    """标准化数值"""
    if pd.isna(val):
        return None
    s = str(val).replace(",", "").replace("%", "").strip()
    if s in ["", "-", "nan"]:
        return None
    try:
        return float(s)
    except:
        return None

def find_col(df, keyword, exact=False):
    """模糊/精确匹配列名"""
    key = keyword.strip().lower()
    for col in df.columns:
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

if not (tc_file and fk_file and ec_file):
    st.error("❌ 文件缺失，请确保文件名中包含 “提成”、“放款明细”、“二次明细”")
    st.stop()

# --- 提成总sheet ---
tc_xls = pd.ExcelFile(tc_file)
if "总" not in tc_xls.sheet_names:
    st.error("❌ 提成文件中未找到 sheet『总』")
    st.stop()
tc_df = pd.read_excel(tc_file, sheet_name="总")

# --- 放款明细中含“潮掣”的两个sheet ---
fk_xls = pd.ExcelFile(fk_file)
fk_sheets = [s for s in fk_xls.sheet_names if "潮掣" in s]
if not fk_sheets:
    st.error("❌ 放款明细文件中未找到包含“潮掣”的sheet")
    st.stop()
fk_dfs = [pd.read_excel(fk_file, sheet_name=s) for s in fk_sheets]

# --- 二次明细：读取文件名包含“二次”的所有sheet并合并 ---
ec_xls = pd.ExcelFile(ec_file)
ec_sheets = ec_xls.sheet_names  # 读取所有sheet
if not ec_sheets:
    st.error("❌ 二次明细文件中没有sheet")
    st.stop()
# 读取所有sheet并合并
ec_df_list = [pd.read_excel(ec_file, sheet_name=s) for s in ec_sheets]
ec_df = pd.concat(ec_df_list, ignore_index=True)
st.success(f"✅ 成功读取 二次明细文件中 {len(ec_sheets)} 个 sheet，共 {len(ec_df)} 行数据")



# =====================================
# 四、字段映射定义
# =====================================
# 「总」字段 → 对应明细字段
MAPPING = {
    "放款日期": ("放款明细", "放款日期", 0, 1),       # 日期比对
    "提报人员": ("放款明细", "提报人员", 0, 1),
    "城市经理": ("放款明细", "城市经理", 0, 1),
    "租赁本金": ("放款明细", "租赁本金", 0, 1),
    "收益率":   ("放款明细", "xirr", 0.005, 1),     # 收益率误差0.005
    "期限":     ("放款明细", "租赁期限/年", 0.5, 12), # 年×12，允许±0.5月
    "二次交接": ("二次明细", "出本流程时间", 0, 1),
}

contract_col_main = find_col(tc_df, "合同")
if not contract_col_main:
    st.error("❌ 『总』sheet 未找到合同号列")
    st.stop()


# =====================================
# 五、主比对函数
# =====================================
def get_ref_row(contract_no, source_type):
    """根据合同号从不同明细中取对应行"""
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
    return None


# =====================================
# 六、执行审核
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

        ref_col = find_col(ref_row.index.to_series(), ref_kw)
        if not ref_col:
            continue

        main_val = row[main_col]
        ref_val = ref_row[ref_col]

        # 日期字段比对
        if "日期" in main_kw:
            main_dt = pd.to_datetime(main_val, errors="coerce")
            ref_dt = pd.to_datetime(ref_val, errors="coerce")
            if pd.isna(main_dt) or pd.isna(ref_dt) or not (
                main_dt.year == ref_dt.year and main_dt.month == ref_dt.month and main_dt.day == ref_dt.day
            ):
                row_has_error = True
                total_errors += 1
                ws.cell(idx + 2, list(tc_df.columns).index(main_col) + 1).fill = red_fill

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

    # 若该行有错误，合同号标黄
    if row_has_error:
        ws.cell(idx + 2, list(tc_df.columns).index(contract_col_main) + 1).fill = yellow_fill

    # 填写原数据
    for j, val in enumerate(row, start=1):
        ws.cell(idx + 2, j, val)

    progress.progress((idx + 1) / n)
    if (idx + 1) % 10 == 0 or idx + 1 == n:
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

