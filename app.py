import io
from datetime import datetime, date

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
    baseline = val(record,"Giá thành cơ sở/con")
    target_saving = val(record,"Mục tiêu tiết kiệm (đồng/con)")

    setv(record,"Chi phí thức ăn/con",safe_div(feed_cost,output_pigs))
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
    setv(record,"Chi phí thuốc/con",safe_div(medicine_cost,output_pigs))
    setv(record,"Chi phí vaccine/con",safe_div(vaccine_cost,output_pigs))
    labor = val(record,"Số lao động")
    setv(record,"Số heo/lao động",safe_div(output_pigs,labor))
    setv(record,"Tổng doanh thu",breeder_rev+meat_rev)
    total_cost = feed_cost+medicine_cost+vaccine_cost+labor_cost+electricity_cost+water_cost+material_cost+maintenance_cost+other_cost
    setv(record,"Tổng chi phí",total_cost)
    setv(record,"Lợi nhuận gộp",(breeder_rev+meat_rev)-total_cost)
    setv(record,"Chi phí điện nước vật tư/con",safe_div(electricity_cost+water_cost+material_cost,output_pigs))
    cost_pig = safe_div(total_cost,output_pigs)
    setv(record,"Giá thành/con",cost_pig)
    actual_saving = (baseline-cost_pig) if (baseline and cost_pig is not None) else None
    setv(record,"Tiết kiệm thực tế (đồng/con)",actual_saving)
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
    setv(r,"Ngày cập nhật dữ liệu",today.isoformat())
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
        df = pd.DataFrame([{f["vi"]: record.get(f["key"]) for f in editable}])
        edited = st.data_editor(
            df, hide_index=True, use_container_width=True,
            num_rows="fixed", key=f"group_editor_{idx}",
            column_config={
                f["vi"]: st.column_config.NumberColumn(f["vi"], format="%.3f")
                for f in editable
                if any(token in f["vi"] for token in ["Số ","Chi phí","Doanh thu","Giá ","Lượng ","Tổng ","Tỷ lệ","Điểm ","Mật độ","Nhiệt độ","Độ ẩm","Khoảng cách","Thời gian","Hao hụt","Tiến độ","Quy mô","Năm","Tháng","Tuần"])
                and "Trạng thái" not in f["vi"] and "Nguyên nhân" not in f["vi"]
            }
        )
        for f in editable:
            record[f["key"]] = edited.iloc[0][f["vi"]]

    record = recalculate(record)
    if derived:
        show = pd.DataFrame([{f["vi"]: record.get(f["key"]) for f in derived}])
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
            group_df=pd.DataFrame(group_data)
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
    c1,c2,c3=st.columns(3)
    farm=c1.selectbox("Trại",["Tất cả"]+sorted(df["farm"].dropna().astype(str).unique().tolist()))
    year=c2.selectbox("Năm",["Tất cả"]+sorted(df["year"].dropna().astype(str).unique().tolist(),reverse=True))
    week=c3.text_input("Tìm tuần")
    view=df.copy()
    if farm!="Tất cả": view=view[view["farm"].astype(str)==farm]
    if year!="Tất cả": view=view[view["year"].astype(str)==year]
    if week: view=view[view["week"].astype(str).str.contains(week,na=False)]

    fixed=["id","year","month","week","region","farm","status","created_by","updated_at"]
    data_cols=[f["key"] for f in FIELD_DEFS if f["key"] in view.columns]
    display=view[fixed+data_cols].rename(columns=VI_BY_KEY)
    st.dataframe(display,use_container_width=True,hide_index=True,height=480)

    output=io.BytesIO()
    with pd.ExcelWriter(output,engine="openpyxl") as writer:
        display.to_excel(writer,index=False,sheet_name="Du lieu GT35")
        ws=writer.book["Du lieu GT35"]
        for cell in ws[1]:
            cell.fill=PatternFill("solid",fgColor="1F6B4A")
            cell.font=Font(color="FFFFFF",bold=True)
            cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        ws.freeze_panes="A2"
        for i,col in enumerate(display.columns,1):
            ws.column_dimensions[get_column_letter(i)].width=min(28,max(12,len(str(col))+2))
    st.download_button("Xuất Excel theo bộ lọc",output.getvalue(),
                       f"Bao cao GT35 {datetime.now():%Y%m%d_%H%M}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if user.get("role") in ("admin","manager"):
        rid=st.number_input("ID cần xóa",min_value=1,step=1)
        if st.button("Xóa bản ghi"):
            ok,msg=delete_record(rid)
            st.success(msg) if ok else st.error(msg)
            if ok: st.rerun()

def dashboard_page():
    st.header("Dashboard GT35")
    df=load_records()
    if df.empty:
        st.info("Chưa có dữ liệu.")
        return
    c1,c2=st.columns(2)
    farm=c1.selectbox("Lọc trại",["Tất cả"]+sorted(df["farm"].dropna().astype(str).unique().tolist()),key="dfarm")
    year=c2.selectbox("Lọc năm",["Tất cả"]+sorted(df["year"].dropna().astype(str).unique().tolist(),reverse=True),key="dyear")
    v=df.copy()
    if farm!="Tất cả": v=v[v["farm"].astype(str)==farm]
    if year!="Tất cả": v=v[v["year"].astype(str)==year]

    def col(vi):
        return KEY_BY_VI.get(vi)
    output=pd.to_numeric(v.get(col("Số heo xuất"),0),errors="coerce").fillna(0).sum()
    cost=pd.to_numeric(v.get(col("Tổng chi phí"),0),errors="coerce").fillna(0).sum()
    feed=pd.to_numeric(v.get(col("Lượng cám sử dụng (kg)"),0),errors="coerce").fillna(0).sum()
    gain=pd.to_numeric(v.get(col("Tăng khối lượng (kg)"),0),errors="coerce").fillna(0).sum()
    pigdays=pd.to_numeric(v.get(col("Tổng ngày-con"),0),errors="coerce").fillna(0).sum()
    deaths=pd.to_numeric(v.get(col("Số heo chết"),0),errors="coerce").fillna(0).sum()
    intake=pd.to_numeric(v.get(col("Số heo nhập"),0),errors="coerce").fillna(0).sum()

    m1,m2,m3,m4,m5=st.columns(5)
    m1.metric("Heo xuất",f"{output:,.0f}")
    m2.metric("Giá thành/con",f"{(cost/output if output else 0):,.0f} đ")
    m3.metric("FCR",f"{(feed/gain if gain else 0):.3f}")
    m4.metric("ADG",f"{(gain*1000/pigdays if pigdays else 0):,.0f} g/ngày")
    m5.metric("Tỷ lệ chết",f"{(deaths/intake*100 if intake else 0):.2f}%")

    chart=v.copy()
    chart["Kỳ"]=chart["year"].astype(str)+"-"+chart["week"].astype(str)
    cp=col("Giá thành/con"); fcr=col("FCR"); adg=col("ADG (g/ngày)")
    for c in [cp,fcr,adg]:
        if c in chart: chart[c]=pd.to_numeric(chart[c],errors="coerce")
    agg={cp:"mean",fcr:"mean",adg:"mean"}
    agg={k:v for k,v in agg.items() if k in chart.columns}
    if agg:
        chart=chart.groupby("Kỳ").agg(agg).sort_index()
        if cp in chart: st.subheader("Giá thành/con theo tuần"); st.line_chart(chart[[cp]])
        if fcr in chart: st.subheader("FCR theo tuần"); st.line_chart(chart[[fcr]])
        if adg in chart: st.subheader("ADG theo tuần"); st.line_chart(chart[[adg]])

def main():
    inject_css()
    seed_farms(INITIAL_FARMS)
    user=get_current_user()
    if not user:
        login_page(); return
    with st.sidebar:
        st.markdown("## 🐷 GT35 WEB V4")
        st.write(f"**{user.get('email')}**")
        st.caption(f"Vai trò: {user.get('role','user')}")
        page=st.radio("Chức năng",[
            "Dashboard","Nhập liệu Input Data","Nhập từ Excel",
            "Dữ liệu & Báo cáo","Quản lý trại"
        ])
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
    else: farm_management_page(user)

if __name__=="__main__":
    main()
