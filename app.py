
from __future__ import annotations

import io
import math
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# =========================================================
# CẤU HÌNH CHUNG
# =========================================================
st.set_page_config(
    page_title="GT35 Professional",
    page_icon="🐷",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_TITLE = "GT35 PROFESSIONAL"
PROJECT_NAME = "Dự án giảm giá thành trại hậu bị"
TARGET_SAVING = 35_000.0
DEFAULT_DATA_FILES = [
    "GT35 Theo doi du lieu Dashboard.xlsx",
    "GT35%20Theo%20doi%20du%20lieu%20Dashboard.xlsx",
    "GT35 Input Data Dashboard.xlsx",
    "GT35%20Input%20Data%20Dashboard.xlsx",
]
INPUT_SHEET_CANDIDATES = ["02 INPUT DATA", "02_INPUT_DATA", "INPUT DATA", "DATA"]

INTEGER_COLUMNS = {
    "Năm", "Tháng", "Tuần", "Quy mô", "Số heo nhập", "Số heo xuất",
    "Số heo chọn giống", "Số heo bán thịt", "Số heo chết", "Tổng số loại thải",
    "Số loại thải sớm", "Số heo giống đã bán", "Số heo thịt đã bán",
    "Số heo bệnh", "Số heo điều trị", "Số heo chuyển", "Số heo chết khi chuyển",
    "Số heo bị thương khi chuyển", "Số lao động", "Số nhân viên đã đào tạo",
    "Số lỗi do con người", "Số heo chưa bán", "Số lần thiếu hàng",
    "Số lần mua khẩn cấp", "Số ý tưởng Kaizen đề xuất",
    "Số ý tưởng Kaizen đã triển khai",
}

DATE_COLUMNS = {
    "Ngày cập nhật dữ liệu", "Hạn hoàn thành", "Ngày bắt đầu",
    "Ngày kết thúc", "Actual Finish", "Start Date", "Deadline",
}

PERCENT_COLUMNS = {
    "Tỷ lệ chọn giống (%)", "Tỷ lệ chọn giống mục tiêu (%)",
    "Tỷ lệ heo thịt (%)", "Tỷ lệ heo không đạt (%)",
    "Tỷ lệ loại thải sớm (%)", "Tỷ lệ chết (%)", "Tổng tỷ lệ hao hụt (%)",
    "Tỷ lệ mắc bệnh (%)", "Điểm đánh giá an toàn sinh học (%)",
    "Độ ẩm trung bình (%)", "Điểm thông gió (%)",
    "Điểm tình trạng máng ăn (%)", "Điểm hệ thống nước (%)",
    "Điểm đánh giá chuồng trại (%)", "Tỷ lệ tuân thủ vaccine (%)",
    "Tỷ lệ điều trị thành công (%)", "Điểm tuân thủ SOP (%)",
    "Điểm KPI nhân viên (%)", "Tỷ lệ đạt mục tiêu tiết kiệm (%)",
    "Tỷ lệ đầy đủ dữ liệu (%)", "Tiến độ (%)",
}

MONEY_COLUMNS = {
    "Chi phí thức ăn", "Chi phí thức ăn/con", "Tiết kiệm từ ADG (đồng/con)",
    "Chi phí hao hụt", "Tiết kiệm do giảm hao hụt (đồng/con)",
    "Chi phí thiệt hại do bệnh", "Chi phí hao hụt do chuyển heo",
    "Chi phí thuốc", "Chi phí vaccine", "Chi phí thuốc/con",
    "Chi phí vaccine/con", "Chi phí thuốc hết hạn", "Chi phí nhân công",
    "Doanh thu heo giống", "Doanh thu heo thịt", "Giá bán heo giống bình quân",
    "Giá bán heo thịt bình quân", "Tổng doanh thu", "Lợi nhuận gộp",
    "Chi phí điện", "Chi phí nước", "Chi phí vật tư", "Chi phí bảo trì",
    "Chi phí khác", "Chi phí điện nước vật tư/con", "Giá trị tồn kho cám",
    "Giá trị tồn kho thuốc", "Giá trị tồn kho vaccine", "Giá trị tồn kho vật tư",
    "Giá trị hàng hết hạn", "Tiết kiệm mua hàng", "Tổng chi phí",
    "Giá thành/con", "Giá thành cơ sở/con", "Mục tiêu tiết kiệm (đồng/con)",
    "Tiết kiệm thực tế (đồng/con)", "Tổng tiền tiết kiệm (đồng)",
    "Tiết kiệm Kaizen đã xác nhận",
}

MODULES = [
    {
        "id": 1,
        "name": "Giảm FCR và chi phí thức ăn",
        "kpis": ["FCR", "FCR mục tiêu", "Chênh lệch FCR", "Chi phí thức ăn/con",
                 "Cám hao hụt (kg)", "Chênh lệch tồn kho cám (kg)"],
        "primary": "FCR",
        "lower_is_better": True,
    },
    {
        "id": 2,
        "name": "Tăng ADG và giảm ngày nuôi",
        "kpis": ["ADG (g/ngày)", "ADG mục tiêu (g/ngày)", "Số ngày nuôi/con xuất",
                 "Số ngày giảm", "Tiết kiệm từ ADG (đồng/con)"],
        "primary": "ADG (g/ngày)",
        "lower_is_better": False,
    },
    {
        "id": 3,
        "name": "Tăng tỷ lệ chọn giống và bán thịt",
        "kpis": ["Tỷ lệ chọn giống (%)", "Tỷ lệ chọn giống mục tiêu (%)",
                 "Tỷ lệ heo thịt (%)", "Tỷ lệ heo không đạt (%)",
                 "Số heo chọn giống", "Số heo bán thịt"],
        "primary": "Tỷ lệ chọn giống (%)",
        "lower_is_better": False,
    },
    {
        "id": 4,
        "name": "Giảm hao hụt và loại thải",
        "kpis": ["Tỷ lệ chết (%)", "Tỷ lệ loại thải sớm (%)",
                 "Tổng tỷ lệ hao hụt (%)", "Chi phí hao hụt",
                 "Tiết kiệm do giảm hao hụt (đồng/con)"],
        "primary": "Tổng tỷ lệ hao hụt (%)",
        "lower_is_better": True,
    },
    {
        "id": 5,
        "name": "Kiểm soát bệnh dịch",
        "kpis": ["Tỷ lệ mắc bệnh (%)", "Điểm đánh giá an toàn sinh học (%)",
                 "Chi phí thiệt hại do bệnh", "Tình trạng PRRS",
                 "Tình trạng APP", "Tình trạng Mycoplasma"],
        "primary": "Tỷ lệ mắc bệnh (%)",
        "lower_is_better": True,
    },
    {
        "id": 6,
        "name": "Quản lý chuồng trại",
        "kpis": ["Mật độ nuôi (con/m²)", "Nhiệt độ trung bình (°C)",
                 "Độ ẩm trung bình (%)", "Điểm thông gió (%)",
                 "Điểm tình trạng máng ăn (%)", "Điểm hệ thống nước (%)",
                 "Điểm đánh giá chuồng trại (%)"],
        "primary": "Điểm đánh giá chuồng trại (%)",
        "lower_is_better": False,
    },
    {
        "id": 7,
        "name": "Quy trình chuyển heo",
        "kpis": ["Số heo chuyển", "Số heo chết khi chuyển",
                 "Số heo bị thương khi chuyển",
                 "Hao hụt khối lượng sau chuyển (kg/con)",
                 "Chi phí hao hụt do chuyển heo"],
        "primary": "Chi phí hao hụt do chuyển heo",
        "lower_is_better": True,
    },
    {
        "id": 8,
        "name": "Vaccine và thuốc",
        "kpis": ["Chi phí thuốc/con", "Chi phí vaccine/con",
                 "Tỷ lệ tuân thủ vaccine (%)", "Tỷ lệ điều trị thành công (%)",
                 "Lượng kháng sinh sử dụng", "Chi phí thuốc hết hạn"],
        "primary": "Chi phí thuốc/con",
        "lower_is_better": True,
    },
    {
        "id": 9,
        "name": "Năng lực nhân viên",
        "kpis": ["Số heo/lao động", "Giờ đào tạo", "Số nhân viên đã đào tạo",
                 "Điểm tuân thủ SOP (%)", "Điểm KPI nhân viên (%)",
                 "Số lỗi do con người", "Chi phí nhân công"],
        "primary": "Điểm tuân thủ SOP (%)",
        "lower_is_better": False,
    },
    {
        "id": 10,
        "name": "Bán heo giống và heo thịt",
        "kpis": ["Doanh thu heo giống", "Doanh thu heo thịt",
                 "Giá bán heo giống bình quân", "Giá bán heo thịt bình quân",
                 "Tổng doanh thu", "Lợi nhuận gộp", "Số heo chưa bán",
                 "Số ngày chậm bán bình quân"],
        "primary": "Lợi nhuận gộp",
        "lower_is_better": False,
    },
    {
        "id": 11,
        "name": "Điện, nước và vật tư",
        "kpis": ["Điện tiêu thụ (kWh)", "Chi phí điện",
                 "Nước tiêu thụ (m³)", "Chi phí nước", "Chi phí vật tư",
                 "Chi phí bảo trì", "Chi phí khác",
                 "Chi phí điện nước vật tư/con"],
        "primary": "Chi phí điện nước vật tư/con",
        "lower_is_better": True,
    },
    {
        "id": 12,
        "name": "Tồn kho và mua hàng",
        "kpis": ["Giá trị tồn kho cám", "Giá trị tồn kho thuốc",
                 "Giá trị tồn kho vaccine", "Giá trị tồn kho vật tư",
                 "Giá trị hàng hết hạn", "Số lần thiếu hàng",
                 "Số lần mua khẩn cấp", "Tiết kiệm mua hàng",
                 "Số ngày quay vòng tồn kho"],
        "primary": "Giá trị hàng hết hạn",
        "lower_is_better": True,
    },
    {
        "id": 13,
        "name": "Dashboard và chất lượng dữ liệu",
        "kpis": ["Tỷ lệ đầy đủ dữ liệu (%)", "Nộp dữ liệu đúng hạn",
                 "Trạng thái chất lượng dữ liệu", "Dashboard đã cập nhật",
                 "Báo cáo đã phê duyệt"],
        "primary": "Tỷ lệ đầy đủ dữ liệu (%)",
        "lower_is_better": False,
    },
    {
        "id": 14,
        "name": "Kaizen và cải tiến liên tục",
        "kpis": ["Số ý tưởng Kaizen đề xuất", "Số ý tưởng Kaizen đã triển khai",
                 "Tiết kiệm Kaizen đã xác nhận", "Hành động cải tiến chính",
                 "Người phụ trách", "Hạn hoàn thành",
                 "Trạng thái hành động", "Tiến độ (%)"],
        "primary": "Tiết kiệm Kaizen đã xác nhận",
        "lower_is_better": False,
    },
]


# =========================================================
# GIAO DIỆN
# =========================================================
st.markdown(
    """
    <style>
    :root {--gt35-green:#146c43; --gt35-light:#eaf6ef; --gt35-dark:#0d3b2a;}
    .stApp {background: #f6f8f7;}
    [data-testid="stSidebar"] {background: linear-gradient(180deg,#0d3b2a,#146c43);}
    [data-testid="stSidebar"] * {color:white;}
    .gt35-title {
        padding: 18px 22px; border-radius: 16px;
        background: linear-gradient(120deg,#0d3b2a,#1d8a58);
        color:white; margin-bottom:16px;
        box-shadow: 0 8px 24px rgba(13,59,42,.15);
    }
    .gt35-title h1 {margin:0; font-size:30px;}
    .gt35-title p {margin:4px 0 0 0; opacity:.9;}
    .metric-card {
        background:white; border-radius:14px; padding:16px;
        border-left:5px solid #1d8a58;
        box-shadow:0 4px 14px rgba(0,0,0,.06); min-height:112px;
    }
    .metric-label {font-size:13px; color:#5d6b64; margin-bottom:8px;}
    .metric-value {font-size:25px; font-weight:700; color:#0d3b2a;}
    .metric-delta {font-size:13px; margin-top:5px;}
    .good {color:#14804a;} .bad {color:#c93c37;} .neutral {color:#68736d;}
    .module-box {
        background:white; border:1px solid #dce8e1; border-radius:12px;
        padding:13px; margin-bottom:8px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# HÀM TIỆN ÍCH
# =========================================================
def clean_column_name(name: object) -> str:
    return " ".join(str(name).replace("\n", " ").strip().split())


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    result.columns = [clean_column_name(c) for c in result.columns]
    result = result.loc[:, ~result.columns.str.startswith("Unnamed:")]
    return result


def find_sheet(excel_file: pd.ExcelFile) -> str:
    normalized = {s.strip().upper(): s for s in excel_file.sheet_names}
    for candidate in INPUT_SHEET_CANDIDATES:
        if candidate.upper() in normalized:
            return normalized[candidate.upper()]
    for sheet in excel_file.sheet_names:
        upper = sheet.upper()
        if "INPUT" in upper and "DATA" in upper:
            return sheet
    return excel_file.sheet_names[0]


@st.cache_data(show_spinner=False)
def read_excel_bytes(file_bytes: bytes) -> pd.DataFrame:
    bio = io.BytesIO(file_bytes)
    excel = pd.ExcelFile(bio)
    sheet = find_sheet(excel)
    df = pd.read_excel(bio, sheet_name=sheet)
    return prepare_data(df)


@st.cache_data(show_spinner=False)
def read_excel_path(path_text: str, modified_time: float) -> pd.DataFrame:
    del modified_time
    path = Path(path_text)
    excel = pd.ExcelFile(path)
    sheet = find_sheet(excel)
    df = pd.read_excel(path, sheet_name=sheet)
    return prepare_data(df)


def prepare_data(df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_columns(df)
    df = df.dropna(how="all").copy()

    # Loại dòng công thức/rác không có Trại, Năm, Tuần.
    key_cols = [c for c in ["Trại", "Năm", "Tuần", "Farm", "Year", "Week"] if c in df.columns]
    if key_cols:
        df = df[df[key_cols].notna().any(axis=1)].copy()

    aliases = {
        "Year": "Năm", "Month": "Tháng", "Week": "Tuần",
        "Region": "Khu vực", "Farm": "Trại", "Capacity": "Quy mô",
        "Farm Manager": "Quản lý trại", "Data Update Date": "Ngày cập nhật dữ liệu",
        "Intake pigs": "Số heo nhập", "Output pigs": "Số heo xuất",
        "Selected pigs": "Số heo chọn giống", "Meat pigs": "Số heo bán thịt",
        "Deaths": "Số heo chết", "Total culled": "Tổng số loại thải",
        "Early culled": "Số loại thải sớm", "Feed kg": "Lượng cám sử dụng (kg)",
        "Feed waste kg": "Cám hao hụt (kg)", "Weight gain kg": "Tăng khối lượng (kg)",
        "Pig-days": "Tổng ngày-con", "Feed cost": "Chi phí thức ăn",
        "Feed cost/pig": "Chi phí thức ăn/con", "Cost per pig": "Giá thành/con",
        "Baseline cost/pig": "Giá thành cơ sở/con",
        "Actual saving (VND/pig)": "Tiết kiệm thực tế (đồng/con)",
        "Total saving (VND)": "Tổng tiền tiết kiệm (đồng)",
        "GT35 Status": "Trạng thái GT35",
    }
    df = df.rename(columns={k: v for k, v in aliases.items() if k in df.columns})

    for col in DATE_COLUMNS.intersection(df.columns):
        df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=False)

    protected_text = {
        "Khu vực", "Trại", "Quản lý trại", "Nguyên nhân chính không đạt chọn giống",
        "Nguyên nhân chết chính", "Nguyên nhân loại thải chính", "Tình trạng PRRS",
        "Tình trạng APP", "Tình trạng Mycoplasma", "Bệnh khác",
        "Nộp dữ liệu đúng hạn", "Trạng thái chất lượng dữ liệu",
        "Dashboard đã cập nhật", "Báo cáo đã phê duyệt",
        "Hành động cải tiến chính", "Người phụ trách",
        "Trạng thái hành động", "Liên kết minh chứng", "Ghi chú",
        "Trạng thái GT35",
    }
    for col in df.columns:
        if col not in protected_text and col not in DATE_COLUMNS:
            converted = pd.to_numeric(df[col], errors="coerce")
            # Chỉ thay khi cột chủ yếu là dữ liệu số.
            non_empty = df[col].notna().sum()
            numeric_count = converted.notna().sum()
            if non_empty > 0 and numeric_count / non_empty >= 0.65:
                df[col] = converted

    # Chuẩn hóa phần trăm: dữ liệu dạng 0.35 sẽ đổi thành 35; dữ liệu 35 giữ nguyên.
    for col in PERCENT_COLUMNS.intersection(df.columns):
        numeric = pd.to_numeric(df[col], errors="coerce")
        valid = numeric.dropna()
        if not valid.empty and valid.abs().quantile(0.9) <= 1.5:
            numeric = numeric * 100
        df[col] = numeric

    # Không fillna(0): ô thiếu phải giữ NaN.
    return df.reset_index(drop=True)


def safe_divide(numerator: pd.Series, denominator: pd.Series, multiplier: float = 1.0) -> pd.Series:
    numerator = pd.to_numeric(numerator, errors="coerce")
    denominator = pd.to_numeric(denominator, errors="coerce")
    result = numerator.div(denominator.where(denominator.ne(0))) * multiplier
    return result.replace([np.inf, -np.inf], np.nan)


def calculate_missing_kpis(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    def has(*cols: str) -> bool:
        return all(c in out.columns for c in cols)

    if "FCR" not in out.columns and has("Lượng cám sử dụng (kg)", "Tăng khối lượng (kg)"):
        out["FCR"] = safe_divide(out["Lượng cám sử dụng (kg)"], out["Tăng khối lượng (kg)"])

    if "ADG (g/ngày)" not in out.columns and has("Tăng khối lượng (kg)", "Tổng ngày-con"):
        out["ADG (g/ngày)"] = safe_divide(out["Tăng khối lượng (kg)"] * 1000, out["Tổng ngày-con"])

    if "Tỷ lệ chọn giống (%)" not in out.columns and has("Số heo chọn giống", "Số heo xuất"):
        out["Tỷ lệ chọn giống (%)"] = safe_divide(out["Số heo chọn giống"], out["Số heo xuất"], 100)

    if "Tỷ lệ chết (%)" not in out.columns and has("Số heo chết", "Số heo nhập"):
        out["Tỷ lệ chết (%)"] = safe_divide(out["Số heo chết"], out["Số heo nhập"], 100)

    if "Tổng tỷ lệ hao hụt (%)" not in out.columns and has("Số heo chết", "Tổng số loại thải", "Số heo nhập"):
        out["Tổng tỷ lệ hao hụt (%)"] = safe_divide(
            out["Số heo chết"].fillna(0) + out["Tổng số loại thải"].fillna(0),
            out["Số heo nhập"],
            100,
        )

    if "Chi phí thức ăn/con" not in out.columns and has("Chi phí thức ăn", "Số heo xuất"):
        out["Chi phí thức ăn/con"] = safe_divide(out["Chi phí thức ăn"], out["Số heo xuất"])

    if "Tiết kiệm thực tế (đồng/con)" not in out.columns and has("Giá thành cơ sở/con", "Giá thành/con"):
        out["Tiết kiệm thực tế (đồng/con)"] = out["Giá thành cơ sở/con"] - out["Giá thành/con"]

    if "Tổng tiền tiết kiệm (đồng)" not in out.columns and has("Tiết kiệm thực tế (đồng/con)", "Số heo xuất"):
        out["Tổng tiền tiết kiệm (đồng)"] = (
            out["Tiết kiệm thực tế (đồng/con)"] * out["Số heo xuất"]
        )

    if "Tỷ lệ đạt mục tiêu tiết kiệm (%)" not in out.columns and "Tiết kiệm thực tế (đồng/con)" in out.columns:
        out["Tỷ lệ đạt mục tiêu tiết kiệm (%)"] = (
            out["Tiết kiệm thực tế (đồng/con)"] / TARGET_SAVING * 100
        )

    return out


def fmt_number(value: object, decimals: int = 0) -> str:
    if value is None or pd.isna(value):
        return "—"
    try:
        number = float(value)
        if decimals == 0:
            return f"{number:,.0f}".replace(",", ".")
        raw = f"{number:,.{decimals}f}"
        return raw.replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return str(value)


def fmt_money(value: object, suffix: str = "đ") -> str:
    return "—" if value is None or pd.isna(value) else f"{fmt_number(value)} {suffix}"


def fmt_percent(value: object, decimals: int = 2) -> str:
    return "—" if value is None or pd.isna(value) else f"{fmt_number(value, decimals)} %"


def fmt_date(value: object) -> str:
    if value is None or pd.isna(value):
        return "—"
    parsed = pd.to_datetime(value, errors="coerce")
    return "—" if pd.isna(parsed) else parsed.strftime("%d/%m/%Y")


def weighted_average(df: pd.DataFrame, value_col: str, weight_col: str = "Số heo xuất") -> float:
    if value_col not in df.columns:
        return np.nan
    values = pd.to_numeric(df[value_col], errors="coerce")
    if weight_col not in df.columns:
        return values.mean()
    weights = pd.to_numeric(df[weight_col], errors="coerce")
    mask = values.notna() & weights.notna() & weights.gt(0)
    if not mask.any():
        return values.mean()
    return float(np.average(values[mask], weights=weights[mask]))


def sum_col(df: pd.DataFrame, col: str) -> float:
    if col not in df.columns:
        return np.nan
    series = pd.to_numeric(df[col], errors="coerce")
    return series.sum(min_count=1)


def mean_col(df: pd.DataFrame, col: str) -> float:
    if col not in df.columns:
        return np.nan
    return pd.to_numeric(df[col], errors="coerce").mean()


def metric_card(label: str, value: str, delta: str = "", state: str = "neutral") -> None:
    st.markdown(
        f"""
        <div class="metric-card">
          <div class="metric-label">{label}</div>
          <div class="metric-value">{value}</div>
          <div class="metric-delta {state}">{delta}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def available_module_columns(df: pd.DataFrame, module: dict) -> list[str]:
    return [c for c in module["kpis"] if c in df.columns and df[c].notna().any()]


def module_completeness(df: pd.DataFrame, module: dict) -> float:
    available = available_module_columns(df, module)
    if not module["kpis"]:
        return 0.0
    return len(available) / len(module["kpis"]) * 100


def latest_vs_previous(
    df: pd.DataFrame, value_col: str, lower_is_better: bool
) -> tuple[float, float, str, str]:
    if value_col not in df.columns:
        return np.nan, np.nan, "Chưa có dữ liệu", "neutral"

    temp = df.copy()
    order_cols = [c for c in ["Năm", "Tuần", "Ngày cập nhật dữ liệu"] if c in temp.columns]
    if order_cols:
        temp = temp.sort_values(order_cols)
    values = pd.to_numeric(temp[value_col], errors="coerce").dropna()
    if values.empty:
        return np.nan, np.nan, "Chưa có dữ liệu", "neutral"

    current = float(values.iloc[-1])
    previous = float(values.iloc[-2]) if len(values) > 1 else np.nan
    if pd.isna(previous):
        return current, previous, "Chưa có kỳ so sánh", "neutral"

    delta = current - previous
    if abs(delta) < 1e-12:
        return current, previous, "Không thay đổi", "neutral"

    improved = delta < 0 if lower_is_better else delta > 0
    arrow = "↓" if delta < 0 else "↑"
    label = f"{arrow} {fmt_number(abs(delta), 2)} so với kỳ trước"
    return current, previous, label, "good" if improved else "bad"


def format_display_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    display = df.copy()
    for col in display.columns:
        if col in DATE_COLUMNS:
            display[col] = display[col].apply(fmt_date)
        elif col in INTEGER_COLUMNS:
            display[col] = display[col].apply(
                lambda x: "" if pd.isna(x) else str(int(float(x)))
            )
        elif col in PERCENT_COLUMNS:
            display[col] = display[col].apply(
                lambda x: "" if pd.isna(x) else fmt_percent(x)
            )
        elif col in MONEY_COLUMNS:
            display[col] = display[col].apply(
                lambda x: "" if pd.isna(x) else fmt_number(x)
            )
    return display


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    export_df = df.copy()
    for col in DATE_COLUMNS.intersection(export_df.columns):
        export_df[col] = pd.to_datetime(export_df[col], errors="coerce")

    with pd.ExcelWriter(output, engine="openpyxl", datetime_format="DD/MM/YYYY") as writer:
        export_df.to_excel(writer, sheet_name="DATA FILTERED", index=False)
        ws = writer.book["DATA FILTERED"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)

        for idx, col in enumerate(export_df.columns, start=1):
            max_len = max(
                [len(str(col))]
                + [
                    len(str(v))
                    for v in export_df[col].dropna().astype(str).head(500)
                ]
            )
            ws.column_dimensions[ws.cell(1, idx).column_letter].width = min(max_len + 2, 35)

            if col in DATE_COLUMNS:
                for row in range(2, ws.max_row + 1):
                    ws.cell(row, idx).number_format = "DD/MM/YYYY"
            elif col in INTEGER_COLUMNS:
                for row in range(2, ws.max_row + 1):
                    ws.cell(row, idx).number_format = "0"
            elif col in MONEY_COLUMNS:
                for row in range(2, ws.max_row + 1):
                    ws.cell(row, idx).number_format = '#,##0'
            elif col in PERCENT_COLUMNS:
                # Dữ liệu đã chuẩn hóa ở thang 0–100.
                for row in range(2, ws.max_row + 1):
                    ws.cell(row, idx).number_format = '0.00" %"'

    return output.getvalue()


# =========================================================
# NẠP DỮ LIỆU
# =========================================================
def locate_default_file() -> Path | None:
    for filename in DEFAULT_DATA_FILES:
        path = Path(filename)
        if path.exists():
            return path
    xlsx_files = [p for p in Path(".").glob("*.xlsx") if not p.name.startswith("~$")]
    return xlsx_files[0] if xlsx_files else None


with st.sidebar:
    st.markdown("## 🐷 GT35")
    st.caption("Professional Dashboard")
    uploaded_file = st.file_uploader(
        "Tải dữ liệu Excel",
        type=["xlsx", "xlsm"],
        help="Ưu tiên sheet 02 INPUT DATA.",
    )
    st.markdown("---")

try:
    if uploaded_file is not None:
        raw_df = read_excel_bytes(uploaded_file.getvalue())
        data_source_name = uploaded_file.name
    else:
        default_file = locate_default_file()
        if default_file is None:
            st.info(
                "Chưa có dữ liệu. Hãy tải file Excel GT35 ở thanh bên hoặc đặt file Excel cùng thư mục với app.py."
            )
            st.stop()
        raw_df = read_excel_path(str(default_file), default_file.stat().st_mtime)
        data_source_name = default_file.name

    df = calculate_missing_kpis(raw_df)
except Exception as exc:
    st.error(f"Không thể đọc dữ liệu: {exc}")
    st.stop()


# =========================================================
# BỘ LỌC TOÀN HỆ THỐNG
# =========================================================
with st.sidebar:
    years = sorted(
        pd.to_numeric(df.get("Năm", pd.Series(dtype=float)), errors="coerce")
        .dropna().astype(int).unique().tolist()
    )
    selected_years = st.multiselect("Năm", years, default=years)

    regions = sorted(df.get("Khu vực", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
    selected_regions = st.multiselect("Khu vực", regions, default=regions)

    farms = sorted(df.get("Trại", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
    selected_farms = st.multiselect("Trại", farms, default=farms)

    weeks = sorted(
        pd.to_numeric(df.get("Tuần", pd.Series(dtype=float)), errors="coerce")
        .dropna().astype(int).unique().tolist()
    )
    selected_weeks = st.multiselect(
        "Tuần",
        weeks,
        default=weeks,
        format_func=lambda x: str(int(x)),
    )

filtered = df.copy()
if selected_years and "Năm" in filtered.columns:
    filtered = filtered[pd.to_numeric(filtered["Năm"], errors="coerce").isin(selected_years)]
if selected_regions and "Khu vực" in filtered.columns:
    filtered = filtered[filtered["Khu vực"].astype(str).isin(selected_regions)]
if selected_farms and "Trại" in filtered.columns:
    filtered = filtered[filtered["Trại"].astype(str).isin(selected_farms)]
if selected_weeks and "Tuần" in filtered.columns:
    filtered = filtered[pd.to_numeric(filtered["Tuần"], errors="coerce").isin(selected_weeks)]

with st.sidebar:
    st.caption(f"Nguồn: {data_source_name}")
    st.caption(f"{len(filtered):,} dòng sau lọc".replace(",", "."))
    st.markdown("---")
    menu = st.radio(
        "Chức năng",
        [
            "Tổng hợp điều hành",
            "Dashboard 14 hạng mục",
            "Chất lượng dữ liệu",
            "Dữ liệu chi tiết",
        ],
    )


st.markdown(
    f"""
    <div class="gt35-title">
      <h1>{APP_TITLE}</h1>
      <p>{PROJECT_NAME} · Dữ liệu không có được giữ trống, không tự đổi thành 0</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if filtered.empty:
    st.warning("Không có dữ liệu phù hợp với bộ lọc.")
    st.stop()


# =========================================================
# 1. TỔNG HỢP ĐIỀU HÀNH
# =========================================================
if menu == "Tổng hợp điều hành":
    total_output = sum_col(filtered, "Số heo xuất")
    current_cost = weighted_average(filtered, "Giá thành/con")
    baseline_cost = weighted_average(filtered, "Giá thành cơ sở/con")
    actual_saving = (
        baseline_cost - current_cost
        if not pd.isna(baseline_cost) and not pd.isna(current_cost)
        else weighted_average(filtered, "Tiết kiệm thực tế (đồng/con)")
    )
    total_saving = sum_col(filtered, "Tổng tiền tiết kiệm (đồng)")
    achievement = (
        actual_saving / TARGET_SAVING * 100
        if not pd.isna(actual_saving)
        else np.nan
    )

    if pd.isna(actual_saving):
        direction_text, direction_state = "Chưa đủ dữ liệu để đánh giá", "neutral"
    elif actual_saving > 0:
        direction_text, direction_state = f"Giảm {fmt_money(actual_saving, 'đ/con')}", "good"
    elif actual_saving < 0:
        direction_text, direction_state = f"Tăng {fmt_money(abs(actual_saving), 'đ/con')}", "bad"
    else:
        direction_text, direction_state = "Không thay đổi", "neutral"

    st.subheader("Kết quả chung")
    cols = st.columns(6)
    with cols[0]:
        metric_card("Giá thành hiện tại", fmt_money(current_cost, "đ/con"))
    with cols[1]:
        metric_card("Giá thành cơ sở", fmt_money(baseline_cost, "đ/con"))
    with cols[2]:
        metric_card("Kết quả chung", direction_text, state=direction_state)
    with cols[3]:
        metric_card("Mục tiêu tiết kiệm", fmt_money(TARGET_SAVING, "đ/con"))
    with cols[4]:
        metric_card("Mức đạt mục tiêu", fmt_percent(achievement))
    with cols[5]:
        metric_card("Tổng tiền tiết kiệm", fmt_money(total_saving))

    st.markdown("### KPI sản xuất chính")
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        metric_card("FCR", fmt_number(weighted_average(filtered, "FCR"), 3))
    with k2:
        metric_card("ADG", f"{fmt_number(weighted_average(filtered, 'ADG (g/ngày)'), 0)} g/ngày")
    with k3:
        metric_card("Tỷ lệ chọn giống", fmt_percent(weighted_average(filtered, "Tỷ lệ chọn giống (%)")))
    with k4:
        metric_card("Tỷ lệ chết", fmt_percent(weighted_average(filtered, "Tỷ lệ chết (%)")))
    with k5:
        metric_card("Tổng hao hụt", fmt_percent(weighted_average(filtered, "Tổng tỷ lệ hao hụt (%)")))

    left, right = st.columns([1.5, 1])

    with left:
        st.markdown("### Xu hướng giá thành")
        if "Giá thành/con" in filtered.columns and filtered["Giá thành/con"].notna().any():
            trend = filtered.copy()
            if "Tuần" in trend.columns:
                trend["Tuần"] = pd.to_numeric(trend["Tuần"], errors="coerce")
                group_cols = [c for c in ["Năm", "Tuần"] if c in trend.columns]
                trend_df = (
                    trend.dropna(subset=["Giá thành/con"])
                    .groupby(group_cols, dropna=False)["Giá thành/con"]
                    .mean()
                    .reset_index()
                )
                trend_df["Kỳ"] = trend_df.apply(
                    lambda r: f"{int(r['Năm'])}-{int(r['Tuần'])}"
                    if "Năm" in r and pd.notna(r.get("Năm"))
                    else str(int(r["Tuần"])),
                    axis=1,
                )
                fig = px.line(
                    trend_df, x="Kỳ", y="Giá thành/con", markers=True,
                    labels={"Giá thành/con": "đ/con"},
                )
                if not pd.isna(baseline_cost):
                    fig.add_hline(
                        y=baseline_cost, line_dash="dash",
                        annotation_text="Baseline",
                    )
                fig.update_layout(height=380, margin=dict(l=10, r=10, t=20, b=10))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Chưa có cột Tuần để vẽ xu hướng.")
        else:
            st.info("Chưa có dữ liệu Giá thành/con.")

    with right:
        st.markdown("### Tiến độ dữ liệu 14 hạng mục")
        progress_rows = []
        for module in MODULES:
            completeness = module_completeness(filtered, module)
            progress_rows.append(
                {
                    "Hạng mục": f"{module['id']:02d}. {module['name']}",
                    "Đầy đủ (%)": completeness,
                }
            )
        progress_df = pd.DataFrame(progress_rows)
        fig = px.bar(
            progress_df.sort_values("Đầy đủ (%)"),
            x="Đầy đủ (%)", y="Hạng mục", orientation="h",
            range_x=[0, 100],
        )
        fig.update_layout(height=380, margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Tăng hay giảm theo từng hạng mục")
    summary_rows = []
    for module in MODULES:
        primary = module["primary"]
        current, previous, trend_label, state = latest_vs_previous(
            filtered, primary, module["lower_is_better"]
        )
        summary_rows.append(
            {
                "STT": module["id"],
                "Hạng mục": module["name"],
                "KPI chính": primary,
                "Hiện tại": current,
                "Kỳ trước": previous,
                "Đánh giá": trend_label,
                "Mức đầy đủ dữ liệu": module_completeness(filtered, module),
            }
        )
    summary_df = pd.DataFrame(summary_rows)
    st.dataframe(
        summary_df.style.format(
            {
                "Hiện tại": lambda x: "—" if pd.isna(x) else fmt_number(x, 2),
                "Kỳ trước": lambda x: "—" if pd.isna(x) else fmt_number(x, 2),
                "Mức đầy đủ dữ liệu": lambda x: fmt_percent(x, 0),
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("### Pareto chi phí")
    cost_candidates = [
        "Chi phí thức ăn", "Chi phí thuốc", "Chi phí vaccine",
        "Chi phí nhân công", "Chi phí điện", "Chi phí nước",
        "Chi phí vật tư", "Chi phí bảo trì", "Chi phí khác",
        "Chi phí hao hụt", "Chi phí thiệt hại do bệnh",
        "Chi phí hao hụt do chuyển heo",
    ]
    pareto_rows = []
    for col in cost_candidates:
        value = sum_col(filtered, col)
        if not pd.isna(value) and value != 0:
            pareto_rows.append({"Nhóm chi phí": col, "Giá trị": max(value, 0)})
    if pareto_rows:
        pareto = pd.DataFrame(pareto_rows).sort_values("Giá trị", ascending=False)
        total = pareto["Giá trị"].sum()
        if total > 0:
            pareto["Lũy kế (%)"] = pareto["Giá trị"].cumsum() / total * 100
            fig = go.Figure()
            fig.add_bar(x=pareto["Nhóm chi phí"], y=pareto["Giá trị"], name="Chi phí")
            fig.add_scatter(
                x=pareto["Nhóm chi phí"], y=pareto["Lũy kế (%)"],
                name="Lũy kế (%)", yaxis="y2", mode="lines+markers",
            )
            fig.update_layout(
                yaxis=dict(title="Chi phí"),
                yaxis2=dict(title="Lũy kế (%)", overlaying="y", side="right", range=[0, 110]),
                height=430,
                margin=dict(l=10, r=10, t=20, b=80),
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Chưa có chi phí dương để lập Pareto.")
    else:
        st.info("Chưa có dữ liệu chi phí để lập Pareto.")


# =========================================================
# 2. DASHBOARD 14 HẠNG MỤC
# =========================================================
elif menu == "Dashboard 14 hạng mục":
    selected_module_name = st.selectbox(
        "Chọn hạng mục",
        [f"{m['id']:02d}. {m['name']}" for m in MODULES],
    )
    module_id = int(selected_module_name.split(".")[0])
    module = next(m for m in MODULES if m["id"] == module_id)

    st.subheader(selected_module_name)
    available_cols = available_module_columns(filtered, module)
    missing_cols = [c for c in module["kpis"] if c not in available_cols]
    completeness = module_completeness(filtered, module)

    a, b, c = st.columns(3)
    with a:
        metric_card("Mức đầy đủ dữ liệu", fmt_percent(completeness, 0))
    with b:
        metric_card("KPI có dữ liệu", f"{len(available_cols)}/{len(module['kpis'])}")
    with c:
        current, previous, trend_label, state = latest_vs_previous(
            filtered, module["primary"], module["lower_is_better"]
        )
        metric_card(module["primary"], fmt_number(current, 2), trend_label, state)

    if missing_cols:
        st.warning("Chưa có dữ liệu: " + ", ".join(missing_cols))
    else:
        st.success("Hạng mục này đã có đủ các trường dữ liệu theo thiết kế.")

    numeric_available = [
        c for c in available_cols
        if pd.api.types.is_numeric_dtype(filtered[c])
    ]
    if numeric_available:
        st.markdown("### KPI hiện tại")
        card_cols = st.columns(min(4, len(numeric_available)))
        for i, col in enumerate(numeric_available[:8]):
            value = weighted_average(filtered, col)
            with card_cols[i % len(card_cols)]:
                if col in MONEY_COLUMNS:
                    value_text = fmt_money(value)
                elif col in PERCENT_COLUMNS:
                    value_text = fmt_percent(value)
                elif col in INTEGER_COLUMNS:
                    value_text = fmt_number(value)
                else:
                    value_text = fmt_number(value, 2)
                metric_card(col, value_text)

        primary = module["primary"]
        if primary in filtered.columns and pd.api.types.is_numeric_dtype(filtered[primary]):
            st.markdown("### Xu hướng KPI chính")
            temp = filtered.copy()
            temp[primary] = pd.to_numeric(temp[primary], errors="coerce")
            temp = temp.dropna(subset=[primary])
            if "Tuần" in temp.columns and not temp.empty:
                group_cols = [c for c in ["Năm", "Tuần"] if c in temp.columns]
                trend = temp.groupby(group_cols, dropna=False)[primary].mean().reset_index()
                trend["Kỳ"] = trend.apply(
                    lambda r: f"{int(r['Năm'])}-{int(r['Tuần'])}"
                    if "Năm" in r and pd.notna(r.get("Năm"))
                    else str(int(r["Tuần"])),
                    axis=1,
                )
                fig = px.line(trend, x="Kỳ", y=primary, markers=True)
                fig.update_layout(height=380)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Chưa có đủ dữ liệu tuần để vẽ xu hướng.")

        if "Trại" in filtered.columns and primary in filtered.columns:
            st.markdown("### So sánh theo trại")
            ranking = (
                filtered.groupby("Trại", dropna=False)[primary]
                .mean()
                .dropna()
                .sort_values(ascending=not module["lower_is_better"])
                .reset_index()
            )
            if not ranking.empty:
                fig = px.bar(ranking, x="Trại", y=primary)
                fig.update_layout(height=380)
                st.plotly_chart(fig, use_container_width=True)

    else:
        st.info("Hạng mục này chưa có dữ liệu số để phân tích. Các hạng mục khác vẫn hoạt động bình thường.")

    detail_cols = [
        c for c in ["Năm", "Tháng", "Tuần", "Khu vực", "Trại", "Ngày cập nhật dữ liệu"]
        + module["kpis"]
        if c in filtered.columns
    ]
    st.markdown("### Dữ liệu chi tiết hạng mục")
    st.dataframe(
        format_display_dataframe(filtered[detail_cols]),
        use_container_width=True,
        hide_index=True,
    )


# =========================================================
# 3. CHẤT LƯỢNG DỮ LIỆU
# =========================================================
elif menu == "Chất lượng dữ liệu":
    st.subheader("Kiểm tra chất lượng dữ liệu")

    module_quality = []
    for module in MODULES:
        available = available_module_columns(filtered, module)
        missing = [c for c in module["kpis"] if c not in available]
        completeness = module_completeness(filtered, module)
        status = (
            "Đầy đủ" if completeness == 100
            else "Có thể phân tích một phần" if completeness >= 50
            else "Thiếu nhiều dữ liệu"
        )
        module_quality.append(
            {
                "STT": module["id"],
                "Hạng mục": module["name"],
                "Số KPI có dữ liệu": len(available),
                "Tổng KPI": len(module["kpis"]),
                "Đầy đủ (%)": completeness,
                "Trạng thái": status,
                "Trường còn thiếu": ", ".join(missing) if missing else "",
            }
        )

    quality_df = pd.DataFrame(module_quality)
    completed = int((quality_df["Đầy đủ (%)"] == 100).sum())
    partial = int(((quality_df["Đầy đủ (%)"] >= 50) & (quality_df["Đầy đủ (%)"] < 100)).sum())
    weak = int((quality_df["Đầy đủ (%)"] < 50).sum())

    x, y, z = st.columns(3)
    with x:
        metric_card("Hạng mục đầy đủ", f"{completed}/14", state="good")
    with y:
        metric_card("Phân tích một phần", f"{partial}/14", state="neutral")
    with z:
        metric_card("Thiếu nhiều dữ liệu", f"{weak}/14", state="bad" if weak else "good")

    st.dataframe(
        quality_df.style.format({"Đầy đủ (%)": lambda x: fmt_percent(x, 0)}),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("### Ô trống theo cột")
    missing_report = (
        pd.DataFrame(
            {
                "Cột dữ liệu": filtered.columns,
                "Số dòng trống": [int(filtered[c].isna().sum()) for c in filtered.columns],
                "Tỷ lệ trống (%)": [filtered[c].isna().mean() * 100 for c in filtered.columns],
            }
        )
        .sort_values("Tỷ lệ trống (%)", ascending=False)
    )
    st.dataframe(
        missing_report.style.format({"Tỷ lệ trống (%)": lambda x: fmt_percent(x, 1)}),
        use_container_width=True,
        hide_index=True,
        height=500,
    )
    st.info(
        "Ô trống được hiểu là chưa có dữ liệu. Số 0 chỉ được hiểu là giá trị thực tế bằng 0."
    )


# =========================================================
# 4. DỮ LIỆU CHI TIẾT
# =========================================================
elif menu == "Dữ liệu chi tiết":
    st.subheader("Dữ liệu sau lọc")

    export_bytes = dataframe_to_excel_bytes(filtered)
    st.download_button(
        "⬇️ Xuất Excel",
        data=export_bytes,
        file_name=f"GT35_Bao_cao_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )

    display_df = format_display_dataframe(filtered)
    st.dataframe(display_df, use_container_width=True, hide_index=True, height=650)

    st.caption(
        "Ngày hiển thị theo dd/mm/yyyy. Năm, tháng, tuần và số lượng không có phần .000."
    )
