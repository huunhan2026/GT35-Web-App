import io
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from field_config import FIELD_DEFS, GROUP_ORDER
from database import (
    is_supabase_configured, login_user, logout_user, get_current_user,
    seed_farms, load_farms, save_farms, save_record, load_records, delete_record
)

st.set_page_config(
    page_title="GT35 – Quản lý trại hậu bị",
    page_icon="🐷", layout="wide", initial_sidebar_state="expanded"
)

INITIAL_FARMS = [{'region': 'Đông Nam Bộ', 'name': 'TAN HUNG 1', 'capacity': 0, 'manager': '', 'active': True}, {'region': 'Miền Tây', 'name': 'LOC TAN 4', 'capacity': 0, 'manager': '', 'active': True}, {'region': 'Nam Trung Bộ', 'name': 'LOC THIEN', 'capacity': 0, 'manager': '', 'active': True}, {'region': 'Khác', 'name': 'SONG LUY', 'capacity': 0, 'manager': '', 'active': True}]

BY_KEY = {f["key"]: f for f in FIELD_DEFS}
BY_VI = {f["vi"]: f for f in FIELD_DEFS}
KEY_BY_VI = {f["vi"]: f["key"] for f in FIELD_DEFS}
VI_BY_KEY = {f["key"]: f["vi"] for f in FIELD_DEFS}
FORMULA_KEYS = [f["key"] for f in FIELD_DEFS if f["is_formula"]]


DATE_FIELDS = {
    f["vi"] for f in FIELD_DEFS
    if "Ngày" in f["vi"] or "Hạn hoàn thành" in f["vi"]
}

def is_blank(value):
    try:
        return value is None or value == "" or pd.isna(value)
    except Exception:
        return value is None or value == ""

def format_date_ddmmyyyy(value):
    if is_blank(value):
        return ""
    dt = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(dt):
        return str(value)
    return dt.strftime("%d/%m/%Y")

def format_number(value, decimals=0):
    if is_blank(value):
        return "—"
    try:
        number = float(value)
        if decimals == 0:
            return f"{number:,.0f}".replace(",", ".")
        s = f"{number:,.{decimals}f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(value)

def numeric_column(df, vi_name):
    key = KEY_BY_VI.get(vi_name)
    if not key or key not in df.columns:
        return pd.Series(dtype=float)
    return pd.to_numeric(df[key], errors="coerce")

def sum_if_available(df, vi_name):
    s = numeric_column(df, vi_name).dropna()
    return None if s.empty else float(s.sum())

def mean_if_available(df, vi_name):
    s = numeric_column(df, vi_name).dropna()
    return None if s.empty else float(s.mean())

def format_display_table(df):
    out = df.copy()
    for col in out.columns:
        if col in DATE_FIELDS or col in ("updated_at", "created_at", "Ngày cập nhật"):
            out[col] = out[col].apply(format_date_ddmmyyyy)
        elif col in ("Năm", "Tháng", "Tuần", "Quy mô"):
            out[col] = out[col].apply(
                lambda x: "" if is_blank(x) else str(int(float(x)))
            )
    return out

def inject_css():
    st.markdown("""
    <style>
    .stApp {background:#f4f8f5;}
    [data-testid="stSidebar"] {background:#103d2c;}
    [data-testid="stSidebar"] * {color:white;}
    .main-title {font-size:2rem;font-weight:800;color:#174c36;margin-bottom:.15rem;}
    .sub-title {color:#557267;margin-bottom:1rem;}
    div[data-testid="stMetric"] {background:white;border:1px solid #dce9e1;
      padding:14px;border-radius:12px;box-shadow:0 2px 8px rgba(20,80,55,.06);}
    .group-title {background:#dcefe4;border-left:6px solid #1f7a4d;padding:9px 12px;
      border-radius:8px;font-weight:700;color:#174c36;margin:8px 0;}
    </style>
    """, unsafe_allow_html=True)

def val(record, vi, default=0):
    k = KEY_BY_VI.get(vi)
    v = record.get(k, default) if k else default
    try:
        if v is None or v == "": return default
        return float(v)
    except Exception:
        return default

def setv(record, vi, value):
    k = KEY_BY_VI.get(vi)
    if k: record[k] = value

def safe_div(a,b,mult=1):
    try:
        return a/b*mult if b else None
    except Exception:
        return None

def recalculate(record):
    # Key formulas matching the Excel logic
    output_pigs = val(record,"Số heo xuất")
    intake = val(record,"Số heo nhập")
    selected = val(record,"Số heo chọn giống")
    meat = val(record,"Số heo bán thịt")
    deaths = val(record,"Số heo chết")
    culled = val(record,"Tổng số loại thải")
    early = val(record,"Số loại thải sớm")
    feed = val(record,"Lượng cám sử dụng (kg)")
    gain = val(record,"Tăng khối lượng (kg)")
    pig_days = val(record,"Tổng ngày-con")
    feed_cost = val(record,"Chi phí thức ăn")
    medicine_cost = val(record,"Chi phí thuốc")
    vaccine_cost = val(record,"Chi phí vaccine")
    labor_cost = val(record,"Chi phí nhân công")
    electricity_cost = val(record,"Chi phí điện")
    water_cost = val(record,"Chi phí nước")
    material_cost = val(record,"Chi phí vật tư")
    maintenance_cost = val(record,"Chi phí bảo trì")
    other_cost = val(record,"Chi phí khác")
    breeder_rev = val(record,"Doanh thu heo giống")
    meat_rev = val(record,"Doanh thu heo thịt")
    baseline = val(record,"Giá thành cơ sở/kg")
    target_saving = val(record,"Mục tiêu tiết kiệm (đồng/kg)")

    setv(record,"Chi phí thức ăn/kg",safe_div(feed_cost,output_pigs))
    setv(record,"FCR",safe_div(feed,gain))
    fcr = val(record,"FCR")
    fcr_target = val(record,"FCR mục tiêu")
    setv(record,"Chênh lệch FCR",fcr-fcr_target if fcr_target else fcr)
    setv(record,"ADG (g/ngày)",safe_div(gain*1000,pig_days))
    days = safe_div(pig_days,output_pigs)
    setv(record,"Số ngày nuôi/con xuất",days)
    target_days = val(record,"Số ngày nuôi mục tiêu/con xuất")
    setv(record,"Số ngày giảm",(target_days-days) if (days is not None and target_days) else (-days if days else None))
    setv(record,"Tỷ lệ chọn giống (%)",safe_div(selected,output_pigs))
    setv(record,"Tỷ lệ heo thịt (%)",safe_div(meat,output_pigs))
    setv(record,"Tỷ lệ heo không đạt (%)",safe_div(max(output_pigs-selected-meat,0),output_pigs))
    setv(record,"Tỷ lệ loại thải sớm (%)",safe_div(early,culled))
    setv(record,"Tỷ lệ chết (%)",safe_div(deaths,intake))
    setv(record,"Tổng tỷ lệ hao hụt (%)",safe_div(deaths+culled,intake))
    sick = val(record,"Số heo bệnh")
    setv(record,"Tỷ lệ mắc bệnh (%)",safe_div(sick,intake))
    setv(record,"Chi phí thuốc/kg",safe_div(medicine_cost,output_pigs))
    setv(record,"Chi phí vaccine/kg",safe_div(vaccine_cost,output_pigs))
    labor = val(record,"Số lao động")
    setv(record,"Số heo/lao động",safe_div(output_pigs,labor))
    setv(record,"Tổng doanh thu",breeder_rev+meat_rev)
    total_cost = feed_cost+medicine_cost+vaccine_cost+labor_cost+electricity_cost+water_cost+material_cost+maintenance_cost+other_cost
    setv(record,"Tổng chi phí",total_cost)
    setv(record,"Lợi nhuận gộp",(breeder_rev+meat_rev)-total_cost)
    setv(record,"Chi phí điện nước vật tư/kg",safe_div(electricity_cost+water_cost+material_cost,output_pigs))
    cost_pig = safe_div(total_cost,output_pigs)
    setv(record,"Giá thành/kg",cost_pig)
    actual_saving = (baseline-cost_pig) if (baseline and cost_pig is not None) else None
    setv(record,"Tiết kiệm thực tế (đồng/kg)",actual_saving)
    setv(record,"Tổng tiền tiết kiệm (đồng)",actual_saving*output_pigs if actual_saving is not None else None)
    setv(record,"Tỷ lệ đạt mục tiêu tiết kiệm (%)",safe_div(actual_saving,target_saving) if actual_saving is not None else None)
    if actual_saving is not None:
        setv(record,"Trạng thái GT35","ĐẠT" if actual_saving >= target_saving else "CHƯA ĐẠT")
    return record

def new_record(farm_row=None):
    r = {f["key"]: None for f in FIELD_DEFS}
    today = date.today()
    setv(r,"Năm",today.year)
    setv(r,"Tháng",today.month)
    setv(r,"Ngày cập nhật dữ liệu",today.strftime("%d/%m/%Y"))
    if farm_row is not None:
        setv(r,"Khu vực",farm_row.get("region",""))
        setv(r,"Trại",farm_row.get("name",""))
        setv(r,"Quy mô",farm_row.get("capacity",0))
        setv(r,"Quản lý trại",farm_row.get("manager",""))
    return r

def format_for_editor(field, value):
    vi = field["vi"]
    if "Ngày" in vi or "Hạn hoàn thành" in vi:
        return value
    return value

def render_group_editor(record, group, idx):
    fields = [f for f in FIELD_DEFS if f["group"] == group]
    editable = [f for f in fields if not f["is_formula"]]
    derived = [f for f in fields if f["is_formula"]]

    st.markdown(f'<div class="group-title">{group}</div>', unsafe_allow_html=True)

    if editable:
        row = {}
        for f in editable:
            value = record.get(f["key"])
            if f["vi"] in DATE_FIELDS and not is_blank(value):
                dt = pd.to_datetime(value, errors="coerce", dayfirst=True)
                value = dt.date() if not pd.isna(dt) else value
            row[f["vi"]] = value

        editor_df = pd.DataFrame([row])
        column_config = {}
        for f in editable:
            vi = f["vi"]
            if vi in DATE_FIELDS:
                column_config[vi] = st.column_config.DateColumn(
                    vi, format="DD/MM/YYYY"
                )
            elif vi in ("Năm", "Tháng", "Tuần", "Quy mô"):
                column_config[vi] = st.column_config.NumberColumn(
                    vi, format="%d"
                )
            elif any(token in vi for token in [
                "Số ", "Chi phí", "Doanh thu", "Giá ", "Lượng ", "Tổng ",
                "Tỷ lệ", "Điểm ", "Mật độ", "Nhiệt độ", "Độ ẩm",
                "Khoảng cách", "Thời gian", "Hao hụt", "Tiến độ"
            ]) and "Trạng thái" not in vi and "Nguyên nhân" not in vi:
                column_config[vi] = st.column_config.NumberColumn(
                    vi, format="%.3f"
                )

        edited = st.data_editor(
            editor_df,
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            key=f"group_editor_{idx}",
            column_config=column_config
        )

        for f in editable:
            value = edited.iloc[0][f["vi"]]
            if f["vi"] in DATE_FIELDS and not is_blank(value):
                value = format_date_ddmmyyyy(value)
            record[f["key"]] = value

    record = recalculate(record)

    if derived:
        show = pd.DataFrame([
            {f["vi"]: record.get(f["key"]) for f in derived}
        ])
        show = format_display_table(show)
        st.caption("Các chỉ tiêu tự động tính")
        st.dataframe(show, hide_index=True, use_container_width=True)

    return record

def login_page():
    st.markdown('<div class="main-title">GT35 – Quản lý trại hậu bị</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Nhập liệu theo đúng cấu trúc Input Data • Quản lý trại • Dashboard</div>', unsafe_allow_html=True)
    with st.form("login"):
        email=st.text_input("Email")
        password=st.text_input("Mật khẩu",type="password")
        submit=st.form_submit_button("Đăng nhập",use_container_width=True)
    if submit:
        ok,msg=login_user(email.strip(),password)
        if ok: st.success(msg); st.rerun()
        else: st.error(msg)
    st.info("Chạy thử: admin@gt35.local / admin123")

def farm_management_page(user):
    st.header("Quản lý danh sách trại")
    if user.get("role") not in ("admin","manager"):
        st.warning("Chỉ Admin hoặc Manager được thay đổi danh sách trại.")
        st.dataframe(load_farms(include_inactive=False),use_container_width=True,hide_index=True)
        return
    st.write("Bạn có thể thêm dòng mới, đổi tên trại, khu vực, quy mô, quản lý trại hoặc khóa trại.")
    df=load_farms(include_inactive=True)
    if df.empty:
        df=pd.DataFrame(columns=["id","region","name","capacity","manager","active"])
    edited=st.data_editor(
        df[["id","region","name","capacity","manager","active"]],
        num_rows="dynamic",hide_index=True,use_container_width=True,
        column_config={
            "id": st.column_config.NumberColumn("ID",disabled=True),
            "region": st.column_config.TextColumn("Khu vực",required=True),
            "name": st.column_config.TextColumn("Tên trại",required=True),
            "capacity": st.column_config.NumberColumn("Quy mô",min_value=0,step=100),
            "manager": st.column_config.TextColumn("Quản lý trại"),
            "active": st.column_config.CheckboxColumn("Đang hoạt động"),
        }
    )
    if st.button("Lưu danh sách trại",type="primary"):
        ok,msg=save_farms(edited)
        st.success(msg) if ok else st.error(msg)
        if ok: st.rerun()

def input_page(user):
    st.header("Nhập dữ liệu giống sheet 02 INPUT DATA")
    farms=load_farms(include_inactive=False)
    if farms.empty:
        st.warning("Chưa có trại. Hãy vào Quản lý trại để thêm trại.")
        return

    names=farms["name"].tolist()
    assigned=user.get("farm","ALL")
    if assigned not in ("ALL","",None) and assigned in names:
        names=[assigned]

    c1,c2,c3=st.columns([2,1,1])
    farm_name=c1.selectbox("Chọn trại",names)
    mode=c2.radio("Kiểu nhập",["Theo 14 nhóm","Bảng Excel"],horizontal=False)
    reset=c3.button("Tạo phiếu mới",use_container_width=True)

    farm_row=farms[farms["name"]==farm_name].iloc[0].to_dict()
    record_key=f"record_{farm_name}"
    if reset or record_key not in st.session_state:
        st.session_state[record_key]=new_record(farm_row)
    record=st.session_state[record_key]

    # Always sync farm master fields
    setv(record,"Khu vực",farm_row.get("region",""))
    setv(record,"Trại",farm_name)
    setv(record,"Quy mô",farm_row.get("capacity",0))
    setv(record,"Quản lý trại",farm_row.get("manager",""))

    if mode=="Theo 14 nhóm":
        tabs=st.tabs([g.replace("THÔNG TIN CHUNG","THÔNG TIN CHUNG") for g in GROUP_ORDER])
        for idx,(tab,group) in enumerate(zip(tabs,GROUP_ORDER)):
            with tab:
                record=render_group_editor(record,group,idx)
    else:
        group=st.selectbox("Chọn nhóm cột để nhập",GROUP_ORDER)
        record=render_group_editor(record,group,100+GROUP_ORDER.index(group))
        st.info("Dữ liệu của các nhóm khác vẫn được giữ trong phiếu. Chọn nhóm khác để tiếp tục nhập.")

    record=recalculate(record)
    st.session_state[record_key]=record

    st.divider()
    c1,c2,c3,c4=st.columns(4)
    year=int(val(record,"Năm",datetime.now().year))
    month=int(val(record,"Tháng",datetime.now().month))
    week_raw=record.get(KEY_BY_VI.get("Tuần"))
    week=str(week_raw).replace(".0","") if week_raw not in (None,"") else ""
    c1.metric("Năm",year)
    c2.metric("Tháng",month)
    c3.metric("Tuần",week or "Chưa nhập")
    c4.metric("Trại",farm_name)

    status=st.selectbox("Trạng thái phiếu",["Nháp","Đã gửi","Đã duyệt","Đã khóa"])
    if st.button("LƯU TOÀN BỘ PHIẾU",type="primary",use_container_width=True):
        if not week:
            st.error("Cần nhập Tuần trong nhóm THÔNG TIN CHUNG.")
        else:
            meta={"year":year,"month":month,"week":week,"region":farm_row.get("region",""),"farm":farm_name}
            ok,msg=save_record(meta,record,user.get("email","unknown"),status)
            st.success(msg) if ok else st.error(msg)

def import_excel_page(user):
    st.header("Nhập từ file Excel đúng mẫu")
    st.write("Sau khi chọn file, hệ thống sẽ hiển thị toàn bộ dữ liệu theo từng nhóm của sheet 02 INPUT DATA.")
    with open("Mau Input Data GT35.xlsx","rb") as f:
        st.download_button("Tải file mẫu gốc",f.read(),"Mau Input Data GT35.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    uploaded=st.file_uploader("Chọn file GT35 Excel",type=["xlsx"])
    if not uploaded:
        return

    try:
        book=load_workbook(uploaded,data_only=True)
    except Exception as e:
        st.error(f"Không đọc được file Excel: {e}")
        return

    if "02 INPUT DATA" not in book.sheetnames:
        st.error("Không tìm thấy sheet '02 INPUT DATA'.")
        return

    sh=book["02 INPUT DATA"]
    headers=[sh.cell(3,c).value for c in range(1,sh.max_column+1)]
    rows=[]
    for r in range(4,sh.max_row+1):
        values=[sh.cell(r,c).value for c in range(1,sh.max_column+1)]
        if not any(v not in (None,"") for v in values):
            continue
        record={f["key"]: None for f in FIELD_DEFS}
        for h,v in zip(headers,values):
            if h and str(h).strip() in KEY_BY_VI:
                record[KEY_BY_VI[str(h).strip()]]=v
        rows.append(recalculate(record))

    if not rows:
        st.warning("File không có dòng dữ liệu trong sheet 02 INPUT DATA.")
        return

    # Tóm tắt nhanh
    summary=[]
    for record in rows:
        summary.append({
            "Năm": record.get(KEY_BY_VI.get("Năm")),
            "Tháng": record.get(KEY_BY_VI.get("Tháng")),
            "Tuần": record.get(KEY_BY_VI.get("Tuần")),
            "Khu vực": record.get(KEY_BY_VI.get("Khu vực")),
            "Trại": record.get(KEY_BY_VI.get("Trại")),
        })
    st.success(f"Đã đọc {len(rows)} dòng và {len(FIELD_DEFS)} cột dữ liệu.")
    st.subheader("Kiểm tra tổng quan")
    st.dataframe(pd.DataFrame(summary),use_container_width=True,hide_index=True,height=300)

    # Hiển thị toàn bộ dữ liệu theo từng nhóm/tab
    st.subheader("Dữ liệu chi tiết theo nhóm")
    tabs=st.tabs(GROUP_ORDER)
    for tab,group in zip(tabs,GROUP_ORDER):
        with tab:
            fields=[f for f in FIELD_DEFS if f["group"]==group]
            group_data=[]
            for record in rows:
                group_data.append({f["vi"]: record.get(f["key"]) for f in fields})
            group_df=format_display_table(pd.DataFrame(group_data))
            st.caption(f"{len(fields)} cột • {len(group_df)} dòng")
            st.dataframe(group_df,use_container_width=True,hide_index=True,height=420)

    st.divider()
    confirm=st.checkbox("Tôi đã kiểm tra dữ liệu và đồng ý nhập toàn bộ vào hệ thống")
    if st.button("NHẬP TOÀN BỘ DỮ LIỆU",type="primary",use_container_width=True,disabled=not confirm):
        ok_count=0
        errors=[]
        for record in rows:
            year=int(val(record,"Năm",0))
            month=int(val(record,"Tháng",0))
            week=str(record.get(KEY_BY_VI.get("Tuần")) or "").replace(".0","")
            farm=str(record.get(KEY_BY_VI.get("Trại")) or "").strip()
            region=str(record.get(KEY_BY_VI.get("Khu vực")) or "").strip()
            if not year or not week or not farm:
                errors.append(f"Thiếu Năm/Tuần/Trại: {farm}-{week}")
                continue
            ok,msg=save_record(
                {"year":year,"month":month,"week":week,"region":region,"farm":farm},
                record,user.get("email","unknown"),"Đã gửi"
            )
            if ok:
                ok_count+=1
            else:
                errors.append(msg)
        if errors:
            st.warning(f"Đã nhập {ok_count} dòng; lỗi {len(errors)} dòng. Lỗi đầu tiên: {errors[0]}")
        else:
            st.success(f"Đã nhập thành công {ok_count} dòng với toàn bộ {len(FIELD_DEFS)} cột.")

def records_page(user):
    st.header("Dữ liệu và báo cáo")
    df=load_records()
    if df.empty:
        st.info("Chưa có dữ liệu.")
        return

    c1,c2,c3,c4=st.columns(4)
    farm=c1.selectbox(
        "Trại",
        ["Tất cả"]+sorted(df["farm"].dropna().astype(str).unique().tolist())
    )
    year=c2.selectbox(
        "Năm",
        ["Tất cả"]+sorted(
            df["year"].dropna().astype(str).unique().tolist(),
            reverse=True
        )
    )
    week=c3.text_input("Tìm tuần")
    status_options=["Tất cả"]
    if "status" in df.columns:
        status_options += sorted(
            df["status"].dropna().astype(str).unique().tolist()
        )
    status_filter=c4.selectbox("Trạng thái",status_options)

    view=df.copy()
    if farm!="Tất cả":
        view=view[view["farm"].astype(str)==farm]
    if year!="Tất cả":
        view=view[view["year"].astype(str)==year]
    if week:
        view=view[
            view["week"].astype(str).str.contains(week,na=False)
        ]
    if status_filter!="Tất cả" and "status" in view.columns:
        view=view[
            view["status"].astype(str)==status_filter
        ]

    fixed=[
        "id","year","month","week","region","farm",
        "status","created_by","updated_at"
    ]
    fixed=[c for c in fixed if c in view.columns]
    data_cols=[
        f["key"] for f in FIELD_DEFS
        if f["key"] in view.columns
    ]
    display=view[fixed+data_cols].rename(columns=VI_BY_KEY)
    display=format_display_table(display)

    st.caption(f"Đang hiển thị {len(view)} bản ghi.")
    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
        height=480
    )

    output=io.BytesIO()
    with pd.ExcelWriter(
        output,
        engine="openpyxl",
        datetime_format="DD/MM/YYYY"
    ) as writer:
        display.to_excel(
            writer,
            index=False,
            sheet_name="Du lieu GT35"
        )
        ws=writer.book["Du lieu GT35"]
        for cell in ws[1]:
            cell.fill=PatternFill("solid",fgColor="1F6B4A")
            cell.font=Font(color="FFFFFF",bold=True)
            cell.alignment=Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
        ws.freeze_panes="A2"
        for i,col in enumerate(display.columns,1):
            ws.column_dimensions[get_column_letter(i)].width=min(
                28,max(12,len(str(col))+2)
            )

    st.download_button(
        "Xuất Excel theo bộ lọc",
        output.getvalue(),
        f"Bao cao GT35 {datetime.now():%d%m%Y_%H%M}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if user.get("role")=="admin":
        st.divider()
        st.subheader("🗑️ Quản lý xóa dữ liệu — Chỉ Admin")
        st.warning(
            "Dữ liệu đã xóa không thể khôi phục trên web. "
            "Hãy xuất Excel sao lưu trước khi xóa."
        )

        delete_mode=st.radio(
            "Chọn cách xóa",
            [
                "Xóa một bản ghi theo ID",
                "Xóa toàn bộ dữ liệu Nháp đang lọc",
                "Xóa toàn bộ dữ liệu đang lọc"
            ],
            key="admin_delete_mode"
        )

        if delete_mode=="Xóa một bản ghi theo ID":
            rid=st.number_input(
                "ID bản ghi cần xóa",
                min_value=1,
                step=1,
                key="delete_record_id"
            )
            confirm_one=st.checkbox(
                f"Tôi xác nhận xóa bản ghi ID {int(rid)}",
                key="confirm_delete_one"
            )
            if st.button(
                "XÓA BẢN GHI",
                type="primary",
                disabled=not confirm_one,
                key="delete_one_button"
            ):
                ok,msg=delete_record(int(rid))
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

        elif delete_mode=="Xóa toàn bộ dữ liệu Nháp đang lọc":
            if "status" in view.columns:
                draft_view=view[
                    view["status"].astype(str).str.strip().str.lower()=="nháp"
                ]
            else:
                draft_view=view.iloc[0:0]

            st.info(
                f"Có {len(draft_view)} bản ghi Nháp phù hợp với bộ lọc hiện tại."
            )

            confirm_text=st.text_input(
                'Nhập chính xác "XOA NHAP" để xác nhận',
                key="confirm_delete_drafts"
            )
            can_delete=(
                len(draft_view)>0
                and confirm_text.strip().upper()=="XOA NHAP"
            )

            if st.button(
                f"XÓA {len(draft_view)} BẢN GHI NHÁP",
                type="primary",
                disabled=not can_delete,
                key="delete_drafts_button"
            ):
                deleted=0
                errors=[]
                for rid in draft_view["id"].dropna().tolist():
                    ok,msg=delete_record(int(rid))
                    if ok:
                        deleted+=1
                    else:
                        errors.append(f"ID {rid}: {msg}")

                if errors:
                    st.warning(
                        f"Đã xóa {deleted} bản ghi; "
                        f"có {len(errors)} lỗi. "
                        f"Lỗi đầu tiên: {errors[0]}"
                    )
                else:
                    st.success(f"Đã xóa {deleted} bản ghi Nháp.")
                st.rerun()

        else:
            st.error(
                "Tùy chọn này sẽ xóa toàn bộ bản ghi "
                "đang hiển thị theo bộ lọc hiện tại."
            )
            st.write(f"Số bản ghi sẽ bị xóa: **{len(view)}**")

            confirm_text=st.text_input(
                'Nhập chính xác "XOA TAT CA" để xác nhận',
                key="confirm_delete_all"
            )
            can_delete=(
                len(view)>0
                and confirm_text.strip().upper()=="XOA TAT CA"
            )

            if st.button(
                f"XÓA {len(view)} BẢN GHI ĐANG LỌC",
                type="primary",
                disabled=not can_delete,
                key="delete_all_button"
            ):
                deleted=0
                errors=[]
                for rid in view["id"].dropna().tolist():
                    ok,msg=delete_record(int(rid))
                    if ok:
                        deleted+=1
                    else:
                        errors.append(f"ID {rid}: {msg}")

                if errors:
                    st.warning(
                        f"Đã xóa {deleted} bản ghi; "
                        f"có {len(errors)} lỗi. "
                        f"Lỗi đầu tiên: {errors[0]}"
                    )
                else:
                    st.success(f"Đã xóa {deleted} bản ghi.")
                st.rerun()
    else:
        st.caption("Chức năng xóa dữ liệu chỉ dành cho Admin.")

def dashboard_page():
    st.header("Dashboard GT35")

    df = load_records()

    # Luôn hiển thị Dashboard tổng hợp và 14 hạng mục, kể cả khi chưa có dữ liệu.
    # Khi chưa có bản ghi, mỗi phần sẽ báo "Chưa có dữ liệu" thay vì dừng toàn trang.
    if df.empty:
        st.info(
            "Chưa có bản ghi đã lưu trong hệ thống. "
            "Hãy vào 'Nhập từ Excel', đánh dấu xác nhận và bấm "
            "'NHẬP TOÀN BỘ DỮ LIỆU', hoặc lưu phiếu tại 'Nhập liệu Input Data'."
        )
        df = pd.DataFrame()

    c1, c2 = st.columns(2)
    farm_options = ["Tất cả"]
    year_options = ["Tất cả"]
    if not df.empty:
        if "farm" in df.columns:
            farm_options += sorted(df["farm"].dropna().astype(str).unique().tolist())
        if "year" in df.columns:
            year_options += sorted(
                df["year"].dropna().astype(str).unique().tolist(), reverse=True
            )

    farm = c1.selectbox("Lọc trại", farm_options, key="dfarm")
    year = c2.selectbox("Lọc năm", year_options, key="dyear")

    view = df.copy()
    if not view.empty:
        if farm != "Tất cả" and "farm" in view.columns:
            view = view[view["farm"].astype(str) == farm]
        if year != "Tất cả" and "year" in view.columns:
            view = view[view["year"].astype(str) == year]

    module_groups = [
        g for g in GROUP_ORDER
        if "THÔNG TIN CHUNG" not in g.upper()
    ][:14]

    tabs = st.tabs(["TỔNG HỢP"] + module_groups)

    with tabs[0]:
        if view.empty:
            st.warning(
                "Dashboard đang hiển thị cấu trúc báo cáo. "
                "Chưa có dữ liệu đã lưu để tính kết quả tổng hợp."
            )

        output = sum_if_available(view, "Số heo xuất")
        total_cost = sum_if_available(view, "Tổng chi phí")
        feed = sum_if_available(view, "Lượng cám sử dụng (kg)")
        gain = sum_if_available(view, "Tăng khối lượng (kg)")
        pig_days = sum_if_available(view, "Tổng ngày-con")
        deaths = sum_if_available(view, "Số heo chết")
        intake = sum_if_available(view, "Số heo nhập")

        cost_per_pig = (
            total_cost / output
            if total_cost is not None and output not in (None, 0)
            else None
        )
        fcr = (
            feed / gain
            if feed is not None and gain not in (None, 0)
            else None
        )
        adg = (
            gain * 1000 / pig_days
            if gain is not None and pig_days not in (None, 0)
            else None
        )
        death_rate = (
            deaths / intake * 100
            if deaths is not None and intake not in (None, 0)
            else None
        )
        baseline = mean_if_available(view, "Giá thành cơ sở/kg")
        saving = (
            baseline - cost_per_pig
            if baseline is not None and cost_per_pig is not None
            else None
        )

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Heo xuất", format_number(output, 0))
        m2.metric("Giá thành/kg", "—" if cost_per_pig is None else f"{format_number(cost_per_pig, 0)} đ")
        m3.metric("FCR", "—" if fcr is None else format_number(fcr, 3))
        m4.metric("ADG", "—" if adg is None else f"{format_number(adg, 0)} g/ngày")
        m5.metric("Tỷ lệ chết", "—" if death_rate is None else f"{format_number(death_rate, 2)}%")

        st.subheader("Kết quả chung")
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Giá thành cơ sở/kg", "—" if baseline is None else f"{format_number(baseline, 0)} đ")
        r2.metric("Giá thành hiện tại/kg", "—" if cost_per_pig is None else f"{format_number(cost_per_pig, 0)} đ")

        if saving is None:
            r3.metric("Kết quả", "Chưa đủ dữ liệu")
        elif saving > 0:
            r3.metric("Kết quả", f"GIẢM {format_number(saving, 0)} đ/kg")
        elif saving < 0:
            r3.metric("Kết quả", f"TĂNG {format_number(abs(saving), 0)} đ/kg")
        else:
            r3.metric("Kết quả", "KHÔNG THAY ĐỔI")
        r4.metric("Mục tiêu", "35.000 đ/kg")

        quality_rows = []
        for group in module_groups:
            fields = [f for f in FIELD_DEFS if f["group"] == group]
            available = [
                f for f in fields
                if f["key"] in view.columns and view[f["key"]].notna().any()
            ]
            pct = len(available) / len(fields) * 100 if fields else 0
            quality_rows.append({
                "Hạng mục": group,
                "Số chỉ tiêu có dữ liệu": len(available),
                "Tổng chỉ tiêu": len(fields),
                "Mức đầy đủ (%)": pct
            })

        st.subheader("Mức đầy đủ dữ liệu 14 hạng mục")
        quality_df = pd.DataFrame(quality_rows)
        st.dataframe(
            quality_df.style.format({"Mức đầy đủ (%)": "{:.0f}%"}),
            use_container_width=True,
            hide_index=True
        )

    for tab, group in zip(tabs[1:], module_groups):
        with tab:
            fields = [f for f in FIELD_DEFS if f["group"] == group]
            available_fields = [
                f for f in fields
                if f["key"] in view.columns and view[f["key"]].notna().any()
            ]
            missing_fields = [f["vi"] for f in fields if f not in available_fields]

            st.markdown(
                f'<div class="group-title">{group}</div>',
                unsafe_allow_html=True
            )

            a, b, c = st.columns(3)
            a.metric("Chỉ tiêu có dữ liệu", f"{len(available_fields)}/{len(fields)}")
            b.metric("Mức đầy đủ", f"{(len(available_fields) / len(fields) * 100 if fields else 0):.0f}%")
            c.metric("Trạng thái", "Có thể phân tích" if available_fields else "Chưa có dữ liệu")

            if missing_fields:
                with st.expander("Các chỉ tiêu chưa có dữ liệu"):
                    st.write(", ".join(missing_fields))

            if not available_fields:
                st.info("Hạng mục này chưa có dữ liệu. Các hạng mục khác vẫn được phân tích bình thường.")
                continue

            numeric_fields = []
            for f in available_fields:
                s = pd.to_numeric(view[f["key"]], errors="coerce")
                if s.notna().any():
                    numeric_fields.append(f)

            if numeric_fields:
                st.subheader("KPI của hạng mục")
                metric_cols = st.columns(min(4, len(numeric_fields)))
                for idx, field in enumerate(numeric_fields[:8]):
                    vi = field["vi"]
                    value = mean_if_available(view, vi)
                    with metric_cols[idx % len(metric_cols)]:
                        if "Chi phí" in vi or "Doanh thu" in vi or "Giá " in vi or "Tiết kiệm" in vi or "Lợi nhuận" in vi:
                            shown = "—" if value is None else f"{format_number(value, 0)} đ"
                        elif "Tỷ lệ" in vi or "(%)" in vi:
                            if value is None:
                                shown = "—"
                            else:
                                pct = value * 100 if abs(value) <= 1.5 else value
                                shown = f"{format_number(pct, 2)}%"
                        else:
                            shown = "—" if value is None else format_number(value, 3)
                        st.metric(vi, shown)

            detail_cols = [c for c in ["year", "month", "week", "region", "farm"] if c in view.columns]
            detail_cols += [f["key"] for f in available_fields]
            detail = view[detail_cols].rename(columns=VI_BY_KEY)
            detail = detail.rename(columns={
                "year": "Năm",
                "month": "Tháng",
                "week": "Tuần",
                "region": "Khu vực",
                "farm": "Trại"
            })
            detail = format_display_table(detail)

            st.subheader("Dữ liệu chi tiết")
            st.dataframe(
                detail,
                use_container_width=True,
                hide_index=True,
                height=380
            )


def _ai_prepare_cost_data(df):
    """
    Chuẩn hóa dữ liệu cho AI Platform mà không thay đổi dữ liệu gốc.

    Ưu tiên dùng cột "Giá thành/kg" đã lưu. Nếu cột này trống,
    hệ thống tự tính giống Dashboard: Tổng chi phí / Số heo xuất.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    out = df.copy()

    cost_key = KEY_BY_VI.get("Giá thành/kg")
    total_cost_key = KEY_BY_VI.get("Tổng chi phí")
    output_key = KEY_BY_VI.get("Số heo xuất")
    baseline_key = KEY_BY_VI.get("Giá thành cơ sở/kg")

    stored_cost = pd.Series(np.nan, index=out.index, dtype=float)
    if cost_key and cost_key in out.columns:
        stored_cost = pd.to_numeric(out[cost_key], errors="coerce")

    calculated_cost = pd.Series(np.nan, index=out.index, dtype=float)
    if (
        total_cost_key and total_cost_key in out.columns
        and output_key and output_key in out.columns
    ):
        total_cost = pd.to_numeric(out[total_cost_key], errors="coerce")
        output_pigs = pd.to_numeric(out[output_key], errors="coerce")
        calculated_cost = total_cost.div(output_pigs.where(output_pigs > 0))

    out["_cost_per_kg"] = stored_cost.fillna(calculated_cost)
    out["_cost_source"] = np.where(
        stored_cost.notna(),
        "Giá thành/kg đã lưu",
        np.where(calculated_cost.notna(), "Tổng chi phí / Số heo xuất", "Không có")
    )

    if baseline_key and baseline_key in out.columns:
        out["_baseline_per_kg"] = pd.to_numeric(out[baseline_key], errors="coerce")
    else:
        out["_baseline_per_kg"] = np.nan

    if "year" in out.columns:
        out["_year"] = pd.to_numeric(out["year"], errors="coerce")
    else:
        year_key = KEY_BY_VI.get("Năm")
        if year_key and year_key in out.columns:
            out["_year"] = pd.to_numeric(out[year_key], errors="coerce")
        else:
            out["_year"] = np.nan

    if "week" in out.columns:
        out["_week"] = pd.to_numeric(out["week"], errors="coerce")
    else:
        week_key = KEY_BY_VI.get("Tuần")
        if week_key and week_key in out.columns:
            out["_week"] = pd.to_numeric(out[week_key], errors="coerce")
        else:
            out["_week"] = np.nan

    if "farm" not in out.columns:
        farm_key = KEY_BY_VI.get("Trại")
        if farm_key and farm_key in out.columns:
            out["farm"] = out[farm_key]
        else:
            out["farm"] = "Không xác định"

    out["farm"] = out["farm"].fillna("Không xác định").astype(str)
    out = out.replace([np.inf, -np.inf], np.nan)
    out = out.dropna(subset=["_cost_per_kg", "_year", "_week"])
    out = out[out["_cost_per_kg"] > 0]
    return out

def _ai_weekly_cost(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=["year", "week", "cost_per_kg"])

    weekly = (
        df.groupby(["_year", "_week"], as_index=False)
        .agg(cost_per_kg=("_cost_per_kg", "mean"))
        .rename(columns={"_year": "year", "_week": "week"})
        .sort_values(["year", "week"])
    )
    weekly["period"] = (
        weekly["year"].astype(int).astype(str)
        + "-W"
        + weekly["week"].astype(int).astype(str).str.zfill(2)
    )
    return weekly.reset_index(drop=True)


def _ai_linear_forecast(weekly_df, periods=4, lookback=8):
    """Dự báo xu hướng tuyến tính đơn giản và có thể giải thích."""
    if weekly_df is None or weekly_df.empty:
        return None, None

    values = pd.to_numeric(weekly_df["cost_per_kg"], errors="coerce").dropna()
    if len(values) < 2:
        return None, None

    recent = values.tail(min(lookback, len(values))).to_numpy(dtype=float)
    x = np.arange(len(recent), dtype=float)
    slope, intercept = np.polyfit(x, recent, 1)
    future_x = len(recent) + int(periods) - 1
    forecast = max(float(intercept + slope * future_x), 0.0)
    return forecast, float(slope)


def _ai_baseline(df):
    if df is None or df.empty:
        return None
    s = pd.to_numeric(df.get("_baseline_per_kg"), errors="coerce").dropna()
    return None if s.empty else float(s.mean())


def _ai_status_text(saving, target=35000.0):
    if saving is None:
        return "Chưa đủ dữ liệu"
    return "ĐẠT" if saving >= target else "CHƯA ĐẠT"


def _ai_forecast_panel(data, title, forecast_weeks, target_saving=35000.0, system_average=None):
    st.markdown(f"### {title}")
    weekly = _ai_weekly_cost(data)
    if len(weekly) < 2:
        st.warning("Cần tối thiểu 2 tuần có dữ liệu giá thành/kg để dự báo.")
        return

    current_cost = float(weekly["cost_per_kg"].tail(min(4, len(weekly))).mean())
    forecast_cost, slope = _ai_linear_forecast(weekly, periods=forecast_weeks)
    baseline = _ai_baseline(data)

    if forecast_cost is None:
        st.warning("Chưa đủ dữ liệu hợp lệ để dự báo.")
        return

    saving = baseline - forecast_cost if baseline is not None else None
    gap = target_saving - saving if saving is not None else None

    cols = st.columns(5)
    cols[0].metric("Giá thành hiện tại", f"{format_number(current_cost, 0)} đ/kg")
    cols[1].metric(
        f"Dự báo sau {forecast_weeks} tuần",
        f"{format_number(forecast_cost, 0)} đ/kg",
        delta=f"{format_number(forecast_cost-current_cost, 0)} đ/kg"
    )
    cols[2].metric("Baseline", "—" if baseline is None else f"{format_number(baseline, 0)} đ/kg")
    cols[3].metric("Mức giảm dự báo", "—" if saving is None else f"{format_number(saving, 0)} đ/kg")
    cols[4].metric("Trạng thái GT35", _ai_status_text(saving, target_saving))

    if saving is None:
        st.info("Chưa có Giá thành cơ sở/kg nên chưa thể xác định mức giảm so với baseline.")
    elif saving >= target_saving:
        st.success(
            f"Dự báo đạt mục tiêu: giảm {format_number(saving, 0)} đ/kg, "
            f"cao hơn yêu cầu {format_number(saving-target_saving, 0)} đ/kg."
        )
    else:
        st.warning(
            f"Dự báo chưa đạt mục tiêu giảm ≥ {format_number(target_saving, 0)} đ/kg. "
            f"Cần giảm thêm {format_number(max(gap, 0), 0)} đ/kg."
        )

    if slope > 0:
        st.error(f"Xu hướng giá thành đang tăng khoảng {format_number(slope, 0)} đ/kg mỗi tuần.")
    elif slope < 0:
        st.success(f"Xu hướng giá thành đang giảm khoảng {format_number(abs(slope), 0)} đ/kg mỗi tuần.")
    else:
        st.info("Xu hướng giá thành gần như không thay đổi.")

    if system_average is not None:
        diff = forecast_cost - system_average
        if diff > 0:
            st.warning(f"Dự báo cao hơn bình quân toàn hệ thống {format_number(diff, 0)} đ/kg.")
        else:
            st.success(f"Dự báo thấp hơn bình quân toàn hệ thống {format_number(abs(diff), 0)} đ/kg.")

    chart = weekly[["period", "cost_per_kg"]].copy()
    chart = chart.rename(columns={"period": "Tuần", "cost_per_kg": "Giá thành thực tế (đ/kg)"})
    st.line_chart(chart.set_index("Tuần"), use_container_width=True)


def _ai_value_or_dash(value, decimals=0, suffix=""):
    """Định dạng an toàn; dữ liệu thiếu hiển thị dấu gạch ngang."""
    if value is None:
        return "—"
    try:
        if pd.isna(value):
            return "—"
        return f"{format_number(float(value), decimals)}{suffix}"
    except Exception:
        return "—"


def _ai_farm_summary(data, forecast_weeks, target_saving):
    """Tạo một dòng kết quả cho một trại; cho phép dữ liệu khuyết."""
    weekly = _ai_weekly_cost(data)
    current = None
    forecast = None
    slope = None

    if not weekly.empty:
        current = float(weekly["cost_per_kg"].tail(min(4, len(weekly))).mean())
        forecast, slope = _ai_linear_forecast(weekly, periods=forecast_weeks)

    baseline = _ai_baseline(data)
    saving = (
        baseline - forecast
        if baseline is not None and forecast is not None
        else None
    )

    return {
        "Số tuần có dữ liệu": int(len(weekly)),
        "Giá thành hiện tại (đ/kg)": current,
        "Giá thành dự báo (đ/kg)": forecast,
        "Baseline (đ/kg)": baseline,
        "Mức giảm dự báo (đ/kg)": saving,
        "Xu hướng/tuần (đ/kg)": slope,
        "GT35": _ai_status_text(saving, target_saving),
    }



def _ai_mean_metric(df, vi_name):
    """Lấy trung bình một chỉ tiêu theo tên tiếng Việt; thiếu dữ liệu trả về None."""
    key = KEY_BY_VI.get(vi_name)
    if df is None or df.empty or not key or key not in df.columns:
        return None
    s = pd.to_numeric(df[key], errors="coerce").replace([np.inf, -np.inf], np.nan).dropna()
    return None if s.empty else float(s.mean())


def _ai_pct_display(value):
    """Hiển thị tỷ lệ nhất quán dù dữ liệu đang lưu dạng 0–1 hay 0–100."""
    if value is None or pd.isna(value):
        return "—"
    shown = value * 100 if abs(value) <= 1.5 else value
    return f"{format_number(shown, 2)}%"


def _ai_relative_score(actual, reference, bad_when="high"):
    """
    Chấm điểm mức lệch so với tham chiếu.
    0: không xấu hơn; 1: nhẹ; 2: trung bình; 3: cao; 4: rất cao.
    """
    if actual is None or reference is None:
        return 0
    try:
        actual = float(actual)
        reference = float(reference)
        if reference == 0:
            return 0
        deviation = (
            (actual - reference) / abs(reference)
            if bad_when == "high"
            else (reference - actual) / abs(reference)
        )
        if deviation <= 0:
            return 0
        if deviation < 0.05:
            return 1
        if deviation < 0.10:
            return 2
        if deviation < 0.20:
            return 3
        return 4
    except Exception:
        return 0


def _ai_priority_label(score):
    if score >= 4:
        return "RẤT CAO"
    if score >= 3:
        return "CAO"
    if score >= 2:
        return "TRUNG BÌNH"
    return "THẤP"


def _ai_build_cost_causes(farm_data, system_data):
    """
    Tạo danh sách nguyên nhân dựa trên dữ liệu thật.
    Ưu tiên so với mục tiêu; nếu không có mục tiêu thì so với toàn hệ thống.
    Không tự tạo số liệu cho chỉ tiêu bị thiếu.
    """
    causes = []
    missing = []

    def add_cause(name, actual_vi, target_vi, bad_when, unit, action, system_compare=True):
        actual = _ai_mean_metric(farm_data, actual_vi)
        target = _ai_mean_metric(farm_data, target_vi) if target_vi else None
        reference_name = "mục tiêu"
        reference = target

        if reference is None and system_compare:
            reference = _ai_mean_metric(system_data, actual_vi)
            reference_name = "bình quân hệ thống"

        if actual is None:
            missing.append(actual_vi)
            return

        if reference is None:
            missing.append(f"{target_vi or actual_vi} (tham chiếu)")
            return

        score = _ai_relative_score(actual, reference, bad_when)
        if score <= 0:
            return

        if unit == "%":
            actual_text = _ai_pct_display(actual)
            reference_text = _ai_pct_display(reference)
        elif unit:
            actual_text = f"{format_number(actual, 3)} {unit}"
            reference_text = f"{format_number(reference, 3)} {unit}"
        else:
            actual_text = format_number(actual, 3)
            reference_text = format_number(reference, 3)

        direction = "cao hơn" if bad_when == "high" else "thấp hơn"
        causes.append({
            "score": score,
            "Nguyên nhân": name,
            "Mức ưu tiên": _ai_priority_label(score),
            "Bằng chứng": (
                f"{actual_vi}: {actual_text}, {direction} "
                f"{reference_name}: {reference_text}."
            ),
            "Hành động đề xuất": action,
        })

    add_cause(
        "FCR cao",
        "FCR",
        "FCR mục tiêu",
        "high",
        "",
        "Kiểm tra hao hụt cám, độ chính xác cấp cám, máng ăn, chất lượng thức ăn và phân nhóm heo theo khối lượng.",
    )
    add_cause(
        "ADG thấp",
        "ADG (g/ngày)",
        "ADG mục tiêu (g/ngày)",
        "low",
        "g/ngày",
        "Kiểm tra lượng ăn, sức khỏe đàn, mật độ nuôi, nhiệt độ chuồng và chương trình dinh dưỡng.",
    )
    add_cause(
        "Ngày nuôi kéo dài",
        "Số ngày nuôi/con xuất",
        "Số ngày nuôi mục tiêu/con xuất",
        "high",
        "ngày",
        "Rà soát tăng trọng theo giai đoạn, tuổi/khối lượng nhập, bệnh mãn tính và thời điểm xuất bán.",
    )
    add_cause(
        "Tỷ lệ chọn giống thấp",
        "Tỷ lệ chọn giống (%)",
        "Tỷ lệ chọn giống mục tiêu (%)",
        "low",
        "%",
        "Phân tích nguyên nhân không đạt, chất lượng đầu vào, chân móng, ngoại hình, tăng trưởng và tiêu chuẩn chọn giống.",
    )
    add_cause(
        "Tỷ lệ chết cao",
        "Tỷ lệ chết (%)",
        None,
        "high",
        "%",
        "Phân tích nguyên nhân chết theo tuần; kiểm tra PRRS, APP, Mycoplasma, điều trị sớm và an toàn sinh học.",
    )
    add_cause(
        "Tỷ lệ loại thải sớm cao",
        "Tỷ lệ loại thải sớm (%)",
        None,
        "high",
        "%",
        "Phân loại nguyên nhân loại thải sớm, kiểm tra chất lượng heo nhập và biện pháp can thiệp ở giai đoạn đầu.",
    )
    add_cause(
        "Tổng hao hụt cao",
        "Tổng tỷ lệ hao hụt (%)",
        None,
        "high",
        "%",
        "Lập Pareto chết và loại thải; tập trung xử lý 2–3 nguyên nhân chiếm tỷ trọng lớn nhất.",
    )
    add_cause(
        "Tỷ lệ mắc bệnh cao",
        "Tỷ lệ mắc bệnh (%)",
        None,
        "high",
        "%",
        "Kiểm tra mô hình bệnh theo tuần, hiệu quả điều trị, vaccine, thông khí và điều kiện chuồng nuôi.",
    )
    add_cause(
        "Chi phí thức ăn/kg cao",
        "Chi phí thức ăn/kg",
        None,
        "high",
        "đ/kg",
        "Rà soát giá cám, FCR, hao hụt tồn kho, định mức cấp cám và hiệu quả từng công thức thức ăn.",
    )
    add_cause(
        "Chi phí thuốc/kg cao",
        "Chi phí thuốc/kg",
        None,
        "high",
        "đ/kg",
        "Rà soát nhóm thuốc chi phí cao, phác đồ điều trị, tỷ lệ tái điều trị và nguyên nhân bệnh chính.",
    )
    add_cause(
        "Chi phí vaccine/kg cao",
        "Chi phí vaccine/kg",
        None,
        "high",
        "đ/kg",
        "Kiểm tra chương trình vaccine, hao hụt sử dụng, liều lượng và mức phù hợp với nguy cơ dịch tễ.",
    )
    add_cause(
        "Chi phí điện nước vật tư/kg cao",
        "Chi phí điện nước vật tư/kg",
        None,
        "high",
        "đ/kg",
        "Kiểm tra điện quạt, bơm nước, rò rỉ, định mức vật tư và các thiết bị tiêu thụ bất thường.",
    )

    causes = sorted(causes, key=lambda x: x["score"], reverse=True)[:5]
    return causes, sorted(set(missing))


def _ai_overall_priority(farm_summary, causes):
    """Xếp hạng mức độ ưu tiên chung của trại."""
    scores = [c["score"] for c in causes]
    max_score = max(scores) if scores else 0
    high_count = sum(s >= 3 for s in scores)

    slope = farm_summary.get("Xu hướng/tuần (đ/kg)")
    gt35 = farm_summary.get("GT35")

    if slope is not None and slope > 0:
        max_score = max(max_score, 3)
    if gt35 == "CHƯA ĐẠT":
        max_score = max(max_score, 2)
    if high_count >= 3 or max_score >= 4:
        return "RẤT CAO"
    if high_count >= 1 or max_score >= 3:
        return "CAO"
    if max_score >= 2:
        return "TRUNG BÌNH"
    return "THẤP"


def _ai_write_farm_comment(selected_farm, farm_summary, causes, missing):
    """Tạo nhận xét tự động, chỉ dựa trên dữ liệu hiện có."""
    priority = _ai_overall_priority(farm_summary, causes)
    current = farm_summary.get("Giá thành hiện tại (đ/kg)")
    forecast = farm_summary.get("Giá thành dự báo (đ/kg)")
    slope = farm_summary.get("Xu hướng/tuần (đ/kg)")
    gt35 = farm_summary.get("GT35")

    sentences = [
        f"Trại **{selected_farm}** có mức độ ưu tiên **{priority}**."
    ]

    if current is not None:
        sentences.append(
            f"Giá thành hiện tại khoảng **{format_number(current, 0)} đ/kg**."
        )
    if forecast is not None:
        sentences.append(
            f"Giá thành dự báo khoảng **{format_number(forecast, 0)} đ/kg**."
        )
    if slope is not None:
        if slope > 0:
            sentences.append(
                f"Xu hướng đang tăng khoảng **{format_number(slope, 0)} đ/kg/tuần**."
            )
        elif slope < 0:
            sentences.append(
                f"Xu hướng đang giảm khoảng **{format_number(abs(slope), 0)} đ/kg/tuần**."
            )
        else:
            sentences.append("Xu hướng giá thành gần như ổn định.")

    if gt35 == "ĐẠT":
        sentences.append("Dự báo hiện đạt mục tiêu GT35.")
    elif gt35 == "CHƯA ĐẠT":
        sentences.append("Dự báo hiện chưa đạt mục tiêu GT35.")

    if causes:
        top_names = ", ".join(c["Nguyên nhân"] for c in causes[:3])
        sentences.append(f"Các vấn đề cần ưu tiên trước gồm: **{top_names}**.")
    else:
        sentences.append(
            "Chưa phát hiện chỉ tiêu xấu hơn mục tiêu hoặc bình quân hệ thống từ dữ liệu hiện có."
        )

    if missing:
        sentences.append(
            "Một số chỉ tiêu còn thiếu nên nhận xét chỉ dựa trên phần dữ liệu đã có."
        )

    return " ".join(sentences), priority


def _ai_show_farm_diagnostics(
    farm_data,
    system_data,
    selected_farm,
    farm_summary,
):
    """Hiển thị 4 chức năng AI mới cho từng trại."""
    causes, missing = _ai_build_cost_causes(farm_data, system_data)
    comment, priority = _ai_write_farm_comment(
        selected_farm,
        farm_summary,
        causes,
        missing,
    )

    st.divider()
    st.markdown("## 🧠 AI phân tích nguyên nhân và hành động")

    c1, c2 = st.columns([1, 3])
    c1.metric("Mức độ ưu tiên", priority)
    with c2:
        st.markdown("### 1. Nhận xét tự động")
        st.info(comment)

    st.markdown("### 2. Xếp hạng mức độ ưu tiên")
    priority_order = {
        "RẤT CAO": 1,
        "CAO": 2,
        "TRUNG BÌNH": 3,
        "THẤP": 4,
    }

    if causes:
        cause_df = pd.DataFrame([
            {
                "Thứ tự": idx + 1,
                "Nguyên nhân": cause["Nguyên nhân"],
                "Mức ưu tiên": cause["Mức ưu tiên"],
                "Bằng chứng": cause["Bằng chứng"],
            }
            for idx, cause in enumerate(causes)
        ])
        cause_df["_order"] = cause_df["Mức ưu tiên"].map(priority_order).fillna(9)
        cause_df = cause_df.sort_values(
            ["_order", "Thứ tự"]
        ).drop(columns="_order")
        st.dataframe(cause_df, use_container_width=True, hide_index=True)

        st.markdown("### 3. Top 5 nguyên nhân làm tăng giá thành")
        for idx, cause in enumerate(causes, start=1):
            st.markdown(
                f"**{idx}. {cause['Nguyên nhân']} — {cause['Mức ưu tiên']}**  \n"
                f"{cause['Bằng chứng']}"
            )

        st.markdown("### 4. Hành động đề xuất cho từng nguyên nhân")
        action_df = pd.DataFrame([
            {
                "Ưu tiên": idx + 1,
                "Nguyên nhân": cause["Nguyên nhân"],
                "Hành động đề xuất": cause["Hành động đề xuất"],
            }
            for idx, cause in enumerate(causes)
        ])
        st.dataframe(action_df, use_container_width=True, hide_index=True)
    else:
        st.success(
            "Chưa phát hiện nguyên nhân bất lợi rõ ràng so với mục tiêu "
            "hoặc bình quân toàn hệ thống từ dữ liệu hiện có."
        )

    if missing:
        with st.expander("Các chỉ tiêu thiếu dữ liệu, không dùng để kết luận"):
            st.write(", ".join(missing))


def _ai_clamp(value, low, high):
    try:
        return max(low, min(high, float(value)))
    except Exception:
        return low


def _ai_decision_plan(farm_data, system_data, selected_farm, farm_summary):
    """Giai đoạn 2: chuyển nguyên nhân thành quyết định ưu tiên có thể hành động."""
    causes, missing = _ai_build_cost_causes(farm_data, system_data)
    st.divider()
    st.markdown("## 🎯 AI Decision – Quyết định ưu tiên")

    if not causes:
        st.success(
            "Chưa phát hiện vấn đề bất lợi rõ ràng từ dữ liệu hiện có. "
            "Tiếp tục theo dõi xu hướng và bổ sung các chỉ tiêu còn thiếu."
        )
        return causes

    decision_rows = []
    for idx, cause in enumerate(causes, start=1):
        level = cause["Mức ưu tiên"]
        if level == "RẤT CAO":
            timing = "Thực hiện ngay trong tuần"
            owner = "Quản lý trại + Bộ phận chuyên môn"
        elif level == "CAO":
            timing = "Triển khai trong 1–2 tuần"
            owner = "Quản lý trại"
        elif level == "TRUNG BÌNH":
            timing = "Lập kế hoạch trong tháng"
            owner = "Phụ trách hạng mục"
        else:
            timing = "Theo dõi định kỳ"
            owner = "Nhân viên phụ trách"

        decision_rows.append({
            "Thứ tự": idx,
            "Mức ưu tiên": level,
            "Vấn đề cần quyết định": cause["Nguyên nhân"],
            "Quyết định đề xuất": cause["Hành động đề xuất"],
            "Thời hạn": timing,
            "Đầu mối": owner,
            "Bằng chứng": cause["Bằng chứng"],
        })

    st.dataframe(
        pd.DataFrame(decision_rows),
        use_container_width=True,
        hide_index=True,
    )

    top = causes[0]
    st.warning(
        f"**Quyết định ưu tiên số 1 cho trại {selected_farm}:** "
        f"{top['Nguyên nhân']}. {top['Hành động đề xuất']}"
    )
    if missing:
        st.caption(
            "Quyết định chỉ dựa trên dữ liệu hiện có. Chỉ tiêu thiếu: "
            + ", ".join(missing)
        )
    return causes


def _ai_simulation_panel(farm_data, selected_farm, farm_summary):
    """Giai đoạn 3: mô phỏng tác động khi cải thiện KPI.

    Đây là mô hình độ nhạy minh bạch, không phải mô hình nhân quả đã hiệu chỉnh.
    """
    st.divider()
    st.markdown("## 🧪 AI Simulation – Mô phỏng phương án")

    current_cost = farm_summary.get("Giá thành hiện tại (đ/kg)")
    if current_cost is None or pd.isna(current_cost) or current_cost <= 0:
        st.info("Thiếu giá thành hiện tại nên chưa thể chạy mô phỏng.")
        return None

    fcr = _ai_mean_metric(farm_data, "FCR")
    adg = _ai_mean_metric(farm_data, "ADG (g/ngày)")
    death = _ai_mean_metric(farm_data, "Tỷ lệ chết (%)")
    days = _ai_mean_metric(farm_data, "Số ngày nuôi/con xuất")
    feed_cost_kg = _ai_mean_metric(farm_data, "Chi phí thức ăn/kg")

    # Chuẩn hóa tỷ lệ chết về phần trăm để hiển thị/mô phỏng.
    death_pct = None
    if death is not None:
        death_pct = death * 100 if abs(death) <= 1.5 else death

    feed_share = 0.65
    if feed_cost_kg is not None and current_cost > 0:
        feed_share = _ai_clamp(feed_cost_kg / current_cost, 0.30, 0.85)

    st.caption(
        "Mô phỏng dùng độ nhạy: FCR tác động theo tỷ trọng thức ăn; "
        "ADG, ngày nuôi và tỷ lệ chết tác động theo hệ số quản trị thận trọng. "
        "Kết quả là ước tính hỗ trợ quyết định, không thay thế thử nghiệm đối chứng."
    )

    col1, col2 = st.columns(2)
    with col1:
        if fcr is not None and fcr > 0:
            sim_fcr = st.slider(
                "FCR sau cải tiến",
                min_value=max(0.5, round(float(fcr) * 0.75, 2)),
                max_value=round(float(fcr), 2),
                value=round(float(fcr), 2),
                step=0.01,
                key=f"sim_fcr_{selected_farm}",
            )
        else:
            sim_fcr = None
            st.text_input("FCR sau cải tiến", value="Thiếu dữ liệu", disabled=True)

        if adg is not None and adg > 0:
            sim_adg = st.slider(
                "ADG sau cải tiến (g/ngày)",
                min_value=int(round(float(adg))),
                max_value=max(int(round(float(adg) * 1.35)), int(round(float(adg))) + 1),
                value=int(round(float(adg))),
                step=5,
                key=f"sim_adg_{selected_farm}",
            )
        else:
            sim_adg = None
            st.text_input("ADG sau cải tiến", value="Thiếu dữ liệu", disabled=True)

    with col2:
        if days is not None and days > 0:
            sim_days = st.slider(
                "Số ngày nuôi sau cải tiến",
                min_value=max(1, int(round(float(days) * 0.75))),
                max_value=int(round(float(days))),
                value=int(round(float(days))),
                step=1,
                key=f"sim_days_{selected_farm}",
            )
        else:
            sim_days = None
            st.text_input("Số ngày nuôi sau cải tiến", value="Thiếu dữ liệu", disabled=True)

        if death_pct is not None and death_pct >= 0:
            sim_death = st.slider(
                "Tỷ lệ chết sau cải tiến (%)",
                min_value=0.0,
                max_value=max(round(float(death_pct), 2), 0.01),
                value=round(float(death_pct), 2),
                step=0.05,
                key=f"sim_death_{selected_farm}",
            )
        else:
            sim_death = None
            st.text_input("Tỷ lệ chết sau cải tiến", value="Thiếu dữ liệu", disabled=True)

    fcr_saving = 0.0
    if fcr is not None and sim_fcr is not None and fcr > 0:
        fcr_saving = current_cost * feed_share * max((float(fcr) - sim_fcr) / float(fcr), 0)

    adg_saving = 0.0
    if adg is not None and sim_adg is not None and adg > 0:
        adg_saving = current_cost * 0.10 * max((sim_adg - float(adg)) / float(adg), 0)

    days_saving = 0.0
    if days is not None and sim_days is not None and days > 0:
        days_saving = current_cost * (1 - feed_share) * 0.25 * max((float(days) - sim_days) / float(days), 0)

    death_saving = 0.0
    if death_pct is not None and sim_death is not None and death_pct > 0:
        death_saving = current_cost * 0.08 * max((death_pct - sim_death) / death_pct, 0)

    total_saving = min(
        fcr_saving + adg_saving + days_saving + death_saving,
        current_cost * 0.30,
    )
    simulated_cost = max(current_cost - total_saving, 0)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Giá thành hiện tại", f"{format_number(current_cost, 0)} đ/kg")
    m2.metric("Tiết kiệm mô phỏng", f"{format_number(total_saving, 0)} đ/kg")
    m3.metric("Giá thành sau mô phỏng", f"{format_number(simulated_cost, 0)} đ/kg")
    m4.metric("Tỷ trọng thức ăn dùng", f"{format_number(feed_share * 100, 1)}%")

    impact = pd.DataFrame([
        {"Phương án": "Giảm FCR", "Tiết kiệm ước tính (đ/kg)": fcr_saving},
        {"Phương án": "Tăng ADG", "Tiết kiệm ước tính (đ/kg)": adg_saving},
        {"Phương án": "Rút ngắn ngày nuôi", "Tiết kiệm ước tính (đ/kg)": days_saving},
        {"Phương án": "Giảm tỷ lệ chết", "Tiết kiệm ước tính (đ/kg)": death_saving},
    ])
    impact["Tiết kiệm ước tính (đ/kg)"] = impact["Tiết kiệm ước tính (đ/kg)"].round(0)
    impact = impact.sort_values("Tiết kiệm ước tính (đ/kg)", ascending=False)
    st.dataframe(impact, use_container_width=True, hide_index=True)

    return {
        "current_cost": current_cost,
        "simulated_cost": simulated_cost,
        "total_saving": total_saving,
        "impact": impact,
    }


def _ai_copilot_answer(question, selected_farm, farm_summary, causes, system_summary, simulation):
    """Giai đoạn 4: trả lời câu hỏi bằng dữ liệu và quy tắc nội bộ, không gọi API ngoài."""
    q = (question or "").strip().lower()
    current = farm_summary.get("Giá thành hiện tại (đ/kg)")
    forecast = farm_summary.get("Giá thành dự báo (đ/kg)")
    saving = farm_summary.get("Mức giảm dự báo (đ/kg)")
    slope = farm_summary.get("Xu hướng/tuần (đ/kg)")
    gt35 = farm_summary.get("GT35")

    if any(k in q for k in ["tại sao", "nguyên nhân", "vì sao", "tăng giá"]):
        if not causes:
            return "Chưa phát hiện nguyên nhân bất lợi rõ ràng từ dữ liệu hiện có."
        lines = [f"Các nguyên nhân ưu tiên của trại **{selected_farm}**:"]
        for i, c in enumerate(causes[:5], 1):
            lines.append(f"{i}. **{c['Nguyên nhân']}** ({c['Mức ưu tiên']}): {c['Bằng chứng']}")
        return "\n\n".join(lines)

    if any(k in q for k in ["làm gì", "hành động", "ưu tiên", "quyết định", "khắc phục"]):
        if not causes:
            return "Chưa có vấn đề nổi bật để đề xuất hành động riêng. Hãy tiếp tục theo dõi và bổ sung dữ liệu."
        lines = [f"Thứ tự hành động đề xuất cho **{selected_farm}**:"]
        for i, c in enumerate(causes[:5], 1):
            lines.append(f"{i}. **{c['Nguyên nhân']}**: {c['Hành động đề xuất']}")
        return "\n\n".join(lines)

    if any(k in q for k in ["đạt gt35", "mục tiêu", "35.000", "35000"]):
        if saving is None:
            return "Chưa có baseline hoặc dự báo phù hợp nên chưa đánh giá được mục tiêu GT35."
        gap = 35000 - saving
        if gt35 == "ĐẠT":
            return f"Trại **{selected_farm}** đang được dự báo **ĐẠT GT35**, mức giảm khoảng **{format_number(saving, 0)} đ/kg**."
        return f"Trại **{selected_farm}** hiện **CHƯA ĐẠT GT35**; cần giảm thêm khoảng **{format_number(max(gap, 0), 0)} đ/kg**."

    if any(k in q for k in ["so sánh", "hệ thống", "trại khác"]):
        system_current = system_summary.get("Giá thành hiện tại (đ/kg)")
        if current is None or system_current is None:
            return "Thiếu dữ liệu để so sánh trại với toàn hệ thống."
        diff = current - system_current
        direction = "cao hơn" if diff > 0 else "thấp hơn" if diff < 0 else "bằng"
        return (
            f"Giá thành hiện tại của **{selected_farm}** là **{format_number(current, 0)} đ/kg**, "
            f"{direction} bình quân hệ thống **{format_number(abs(diff), 0)} đ/kg**."
        )

    if any(k in q for k in ["mô phỏng", "phương án", "tiết kiệm"]):
        if not simulation:
            return "Hãy điều chỉnh các thanh mô phỏng ở mục AI Simulation để có kết quả phương án."
        return (
            f"Theo phương án đang chọn, giá thành của **{selected_farm}** có thể từ "
            f"**{format_number(simulation['current_cost'], 0)}** xuống khoảng "
            f"**{format_number(simulation['simulated_cost'], 0)} đ/kg**, "
            f"tương ứng tiết kiệm chỉ báo **{format_number(simulation['total_saving'], 0)} đ/kg**."
        )

    parts = [f"Tóm tắt trại **{selected_farm}**:"]
    if current is not None:
        parts.append(f"Giá thành hiện tại **{format_number(current, 0)} đ/kg**.")
    if forecast is not None:
        parts.append(f"Dự báo **{format_number(forecast, 0)} đ/kg**.")
    if slope is not None:
        trend = "tăng" if slope > 0 else "giảm" if slope < 0 else "ổn định"
        parts.append(f"Xu hướng {trend} khoảng **{format_number(abs(slope), 0)} đ/kg/tuần**.")
    if causes:
        parts.append("Ưu tiên chính: " + ", ".join(c["Nguyên nhân"] for c in causes[:3]) + ".")
    parts.append("Bạn có thể hỏi: 'Tại sao giá thành tăng?', 'Nên làm gì trước?', 'Có đạt GT35 không?' hoặc 'So sánh với hệ thống'.")
    return " ".join(parts)


def _ai_copilot_panel(farm_data, system_data, selected_farm, farm_summary, simulation):
    st.divider()
    st.markdown("## 💬 GT35 Copilot – Hỏi đáp dữ liệu trại")
    st.caption(
        "Copilot này trả lời bằng dữ liệu đang có và bộ quy tắc GT35 nội bộ; "
        "không gửi dữ liệu ra dịch vụ AI bên ngoài."
    )

    causes, _ = _ai_build_cost_causes(farm_data, system_data)
    system_summary = _ai_farm_summary(system_data, 4, 35000.0)
    question = st.text_area(
        "Nhập câu hỏi",
        placeholder=(
            "Ví dụ: Tại sao giá thành trại này tăng? Nên ưu tiên hành động nào? "
            "Trại có đạt GT35 không?"
        ),
        key=f"copilot_question_{selected_farm}",
    )
    if st.button(
        "🤖 HỎI GT35 COPILOT",
        type="primary",
        use_container_width=True,
        key=f"copilot_run_{selected_farm}",
    ):
        answer = _ai_copilot_answer(
            question,
            selected_farm,
            farm_summary,
            causes,
            system_summary,
            simulation,
        )
        st.session_state[f"copilot_answer_{selected_farm}"] = answer

    answer = st.session_state.get(f"copilot_answer_{selected_farm}")
    if answer:
        st.info(answer)


def _ai_show_stages_2_3_4(farm_data, system_data, selected_farm, farm_summary):
    """Hiển thị liền mạch Giai đoạn 2, 3 và 4 sau phần phân tích Giai đoạn 1."""
    _ai_decision_plan(
        farm_data=farm_data,
        system_data=system_data,
        selected_farm=selected_farm,
        farm_summary=farm_summary,
    )
    simulation = _ai_simulation_panel(
        farm_data=farm_data,
        selected_farm=selected_farm,
        farm_summary=farm_summary,
    )
    _ai_copilot_panel(
        farm_data=farm_data,
        system_data=system_data,
        selected_farm=selected_farm,
        farm_summary=farm_summary,
        simulation=simulation,
    )


def _ai_show_individual_farm_analysis(
    farm_data,
    selected_farm,
    forecast_weeks,
    target_saving,
    system_data,
):
    """Phân tích từng trại; chỉ tiêu thiếu sẽ hiển thị khuyết thay vì gây lỗi."""
    st.markdown(f"### 🤖 Kết quả AI – Trại {selected_farm}")

    farm_summary = _ai_farm_summary(
        farm_data,
        forecast_weeks,
        target_saving,
    )
    system_summary = _ai_farm_summary(
        system_data,
        forecast_weeks,
        target_saving,
    )

    weeks_available = farm_summary["Số tuần có dữ liệu"]
    current = farm_summary["Giá thành hiện tại (đ/kg)"]
    forecast = farm_summary["Giá thành dự báo (đ/kg)"]
    baseline = farm_summary["Baseline (đ/kg)"]
    saving = farm_summary["Mức giảm dự báo (đ/kg)"]
    slope = farm_summary["Xu hướng/tuần (đ/kg)"]

    cols = st.columns(5)
    cols[0].metric(
        "Giá thành hiện tại",
        _ai_value_or_dash(current, 0, " đ/kg"),
    )
    cols[1].metric(
        f"Dự báo sau {forecast_weeks} tuần",
        _ai_value_or_dash(forecast, 0, " đ/kg"),
    )
    cols[2].metric(
        "Baseline",
        _ai_value_or_dash(baseline, 0, " đ/kg"),
    )
    cols[3].metric(
        "Mức giảm dự báo",
        _ai_value_or_dash(saving, 0, " đ/kg"),
    )
    cols[4].metric(
        "Trạng thái GT35",
        farm_summary["GT35"],
    )

    if weeks_available < 2:
        st.warning(
            f"Trại {selected_farm} hiện chỉ có {weeks_available} tuần dữ liệu hợp lệ. "
            "Cần tối thiểu 2 tuần để dự báo xu hướng. Các chỉ tiêu thiếu được để khuyết (—)."
        )
    else:
        if saving is None:
            st.info(
                "Trại chưa có Giá thành cơ sở/kg nên chưa tính được mức giảm và trạng thái GT35."
            )
        elif saving >= target_saving:
            st.success(
                f"Dự báo đạt mục tiêu GT35: giảm {format_number(saving, 0)} đ/kg."
            )
        else:
            st.warning(
                f"Dự báo chưa đạt GT35; cần giảm thêm "
                f"{format_number(max(target_saving-saving, 0), 0)} đ/kg."
            )

        if slope is None:
            st.info("Chưa đủ dữ liệu để xác định xu hướng theo tuần.")
        elif slope > 0:
            st.error(
                f"Giá thành có xu hướng tăng khoảng {format_number(slope, 0)} đ/kg/tuần."
            )
        elif slope < 0:
            st.success(
                f"Giá thành có xu hướng giảm khoảng {format_number(abs(slope), 0)} đ/kg/tuần."
            )
        else:
            st.info("Giá thành gần như không thay đổi theo tuần.")

    st.subheader("So sánh trại với toàn hệ thống")
    comparison_rows = []
    comparison_fields = [
        "Giá thành hiện tại (đ/kg)",
        "Giá thành dự báo (đ/kg)",
        "Baseline (đ/kg)",
        "Mức giảm dự báo (đ/kg)",
        "Xu hướng/tuần (đ/kg)",
    ]

    for field in comparison_fields:
        farm_value = farm_summary.get(field)
        system_value = system_summary.get(field)
        difference = None
        if farm_value is not None and system_value is not None:
            try:
                if not pd.isna(farm_value) and not pd.isna(system_value):
                    difference = float(farm_value) - float(system_value)
            except Exception:
                difference = None

        comparison_rows.append({
            "Chỉ tiêu": field,
            "Trại": _ai_value_or_dash(farm_value, 0),
            "Toàn hệ thống": _ai_value_or_dash(system_value, 0),
            "Chênh lệch": _ai_value_or_dash(difference, 0),
            "Ghi chú": (
                "Thiếu dữ liệu để so sánh"
                if difference is None
                else (
                    "Trại cao hơn hệ thống"
                    if difference > 0
                    else "Trại thấp hơn hệ thống"
                    if difference < 0
                    else "Bằng hệ thống"
                )
            ),
        })

    st.dataframe(
        pd.DataFrame(comparison_rows),
        use_container_width=True,
        hide_index=True,
    )

    weekly = _ai_weekly_cost(farm_data)
    if not weekly.empty:
        chart = weekly[["period", "cost_per_kg"]].copy()
        chart = chart.rename(columns={
            "period": "Tuần",
            "cost_per_kg": "Giá thành thực tế (đ/kg)",
        })
        st.subheader("Xu hướng giá thành của trại")
        st.line_chart(chart.set_index("Tuần"), use_container_width=True)

    st.subheader("Mức đầy đủ dữ liệu của trại")
    completeness = []
    checks = [
        ("Giá thành/kg", "_cost_per_kg"),
        ("Giá thành cơ sở/kg", "_baseline_per_kg"),
        ("Năm", "_year"),
        ("Tuần", "_week"),
        ("Trại", "farm"),
    ]
    for label, col in checks:
        count = int(farm_data[col].notna().sum()) if col in farm_data.columns else 0
        completeness.append({
            "Chỉ tiêu": label,
            "Số bản ghi có dữ liệu": count,
            "Trạng thái": "Có dữ liệu" if count > 0 else "Thiếu dữ liệu",
        })
    st.dataframe(
        pd.DataFrame(completeness),
        use_container_width=True,
        hide_index=True,
    )

    _ai_show_farm_diagnostics(
        farm_data=farm_data,
        system_data=system_data,
        selected_farm=selected_farm,
        farm_summary=farm_summary,
    )

    _ai_show_stages_2_3_4(
        farm_data=farm_data,
        system_data=system_data,
        selected_farm=selected_farm,
        farm_summary=farm_summary,
    )


def ai_platform_page(user):
    """Trang AI Platform chỉ dành cho Admin."""
    if user.get("role") != "admin":
        st.error("Chức năng AI Platform chỉ dành cho Admin.")
        return

    st.header("🤖 AI Platform – Dự báo chi phí GT35")
    st.success("Đang chạy bản GT35 AI FULL GIAI ĐOẠN 1–4 – 27/07/2026")
    st.caption(
        "Chọn phạm vi, sau đó bấm nút AI tại đúng phạm vi cần phân tích. "
        "Mục tiêu GT35: mức giảm so với baseline ≥ 35.000 VND/kg."
    )

    raw = load_records()
    if raw is None or raw.empty:
        st.warning("Hệ thống chưa có dữ liệu đã lưu.")
        return

    data = _ai_prepare_cost_data(raw)
    if data.empty:
        st.error(
            "AI chưa tạo được dữ liệu giá thành theo tuần. "
            "Cần có Năm, Tuần và một trong hai nguồn: "
            "Giá thành/kg hoặc Tổng chi phí cùng Số heo xuất."
        )
        return

    c1, c2, c3 = st.columns(3)
    years = sorted(
        data["_year"].dropna().astype(int).unique().tolist(),
        reverse=True,
    )
    selected_year = c1.selectbox(
        "Năm phân tích",
        ["Tất cả"] + years,
        key="ai_year",
    )
    forecast_weeks = c2.selectbox(
        "Số tuần dự báo",
        [1, 2, 4, 8],
        index=2,
        key="ai_weeks",
    )
    target_saving = c3.number_input(
        "Mục tiêu giảm (VND/kg)",
        min_value=0.0,
        value=35000.0,
        step=1000.0,
        key="ai_target",
    )

    if selected_year != "Tất cả":
        data = data[data["_year"] == int(selected_year)]

    if data.empty:
        st.warning("Không có dữ liệu phù hợp với năm đã chọn.")
        return

    scope = st.radio(
        "Phạm vi phân tích",
        ["Tổng toàn bộ trại", "Từng trại"],
        horizontal=True,
        key="ai_scope",
    )

    if scope == "Tổng toàn bộ trại":
        st.markdown("### 🤖 Phân tích tổng toàn bộ trại")
        run_total = st.button(
            "🤖 YÊU CẦU AI PHÂN TÍCH TỔNG TRẠI",
            type="primary",
            use_container_width=True,
            key="run_ai_total",
        )
        if run_total:
            st.session_state["ai_total_ready"] = True

        if not st.session_state.get("ai_total_ready", False):
            st.info("Bấm nút AI để chạy phân tích tổng toàn bộ trại.")
            return

        with st.spinner("AI đang phân tích tổng toàn bộ trại..."):
            _ai_forecast_panel(
                data,
                "Dự báo tổng toàn bộ trại",
                forecast_weeks,
                target_saving=target_saving,
            )

        st.subheader("So sánh và xếp hạng từng trại")
        rows = []
        for farm_name in sorted(
            data["farm"].dropna().astype(str).unique().tolist()
        ):
            farm_data = data[data["farm"].astype(str) == farm_name]
            summary = _ai_farm_summary(
                farm_data,
                forecast_weeks,
                target_saving,
            )
            rows.append({
                "Trại": farm_name,
                "Số tuần có dữ liệu": summary["Số tuần có dữ liệu"],
                "Giá thành hiện tại (đ/kg)": _ai_value_or_dash(
                    summary["Giá thành hiện tại (đ/kg)"], 0
                ),
                "Giá thành dự báo (đ/kg)": _ai_value_or_dash(
                    summary["Giá thành dự báo (đ/kg)"], 0
                ),
                "Baseline (đ/kg)": _ai_value_or_dash(
                    summary["Baseline (đ/kg)"], 0
                ),
                "Mức giảm dự báo (đ/kg)": _ai_value_or_dash(
                    summary["Mức giảm dự báo (đ/kg)"], 0
                ),
                "Xu hướng/tuần (đ/kg)": _ai_value_or_dash(
                    summary["Xu hướng/tuần (đ/kg)"], 0
                ),
                "GT35": summary["GT35"],
            })

        st.dataframe(
            pd.DataFrame(rows),
            use_container_width=True,
            hide_index=True,
        )

    else:
        farms = sorted(
            data["farm"].dropna().astype(str).unique().tolist()
        )
        if not farms:
            st.warning("Không có tên trại hợp lệ trong dữ liệu.")
            return

        selected_farm = st.selectbox(
            "Chọn trại cần AI phân tích",
            farms,
            key="ai_farm",
        )
        farm_data = data[data["farm"].astype(str) == selected_farm]

        st.markdown(f"### 🤖 AI phân tích trại {selected_farm}")
        run_farm = st.button(
            f"🤖 YÊU CẦU AI PHÂN TÍCH TRẠI {selected_farm}",
            type="primary",
            use_container_width=True,
            key=f"run_ai_farm_{selected_farm}",
        )
        farm_ready_key = f"ai_farm_ready_{selected_farm}"
        if run_farm:
            st.session_state[farm_ready_key] = True

        if not st.session_state.get(farm_ready_key, False):
            st.info(
                f"Bấm nút AI để phân tích riêng trại {selected_farm}. "
                "Chỉ tiêu thiếu dữ liệu sẽ hiển thị dấu — và ghi chú thiếu dữ liệu."
            )
            return

        with st.spinner(f"AI đang phân tích trại {selected_farm}..."):
            _ai_show_individual_farm_analysis(
                farm_data=farm_data,
                selected_farm=selected_farm,
                forecast_weeks=forecast_weeks,
                target_saving=target_saving,
                system_data=data,
            )

    st.divider()
    st.caption(
        "Dự báo dùng xu hướng tuyến tính từ dữ liệu lịch sử. "
        "Chỉ tiêu không đủ dữ liệu được để khuyết (—), không làm dừng ứng dụng."
    )

def main():
    inject_css()
    seed_farms(INITIAL_FARMS)
    user=get_current_user()
    if not user:
        login_page(); return
    with st.sidebar:
        st.markdown("## 🐷 GT35 WEB V4")
        st.caption("Build: GT35-AI-FULL-1-4-20260727")
        st.write(f"**{user.get('email')}**")
        st.caption(f"Vai trò: {user.get('role','user')}")
        menu_items=[
            "Dashboard","Nhập liệu Input Data","Nhập từ Excel",
            "Dữ liệu & Báo cáo","Quản lý trại"
        ]
        if user.get("role")=="admin":
            menu_items.append("AI Platform")
        page=st.radio("Chức năng",menu_items)
        st.divider()
        st.caption("Dữ liệu: "+("Supabase trực tuyến" if is_supabase_configured() else "SQLite chạy thử"))
        if st.button("Đăng xuất",use_container_width=True):
            logout_user(); st.rerun()

    st.markdown('<div class="main-title">GT35 – Quản lý trại hậu bị</div>',unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Đầu vào đầy đủ theo các nhóm của sheet 02 INPUT DATA</div>',unsafe_allow_html=True)
    if page=="Dashboard": dashboard_page()
    elif page=="Nhập liệu Input Data": input_page(user)
    elif page=="Nhập từ Excel": import_excel_page(user)
    elif page=="Dữ liệu & Báo cáo": records_page(user)
    elif page=="AI Platform": ai_platform_page(user)
    else: farm_management_page(user)

if __name__=="__main__":
    main()
