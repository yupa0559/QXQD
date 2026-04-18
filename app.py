"""
外观不良智能管理系统 V3 - Python/Streamlit 版
对应原 HTML V3，功能完整等价
"""

import streamlit as st
import json
import time
import requests
import io
from datetime import date, datetime
from typing import Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ════════════════════════════════════════════════
# CONFIG
# ════════════════════════════════════════════════
DEEPSEEK_KEY = "sk-c285636d450b468ba3b437891743d3d5"
DS_API       = "https://api.deepseek.com/v1/chat/completions"

SB_URL   = "https://scgscpfjpdhtqdzkszty.supabase.co"
SB_KEY   = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InNjZ3NjcGZqcGRodHFkemtzenR5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzU4MjkxMzEsImV4cCI6MjA5MTQwNTEzMX0.Yl93_IKvqfL50GJduPfaK-oijnSN5zZCwXYifMO_4U0"
SB_TABLE = "defect_records"

SEV_LABEL  = {"low": "低·轻微", "medium": "中·一般", "high": "高·严重", "critical": "致命·报废"}
SHIFT_LABEL = {"day": "☀ 白班", "night": "🌙 夜班"}

DEFAULT_TYPES = ["缺料", "缩水", "批锋", "气纹", "拉伤", "划痕", "氧化", "开裂", "变形", "其它"]
DEFAULT_MOLDS = [
    {"id": "M24-001", "name": "M24-001", "partNo": "", "partVer": "", "productName": ""},
    {"id": "M24-002", "name": "M24-002", "partNo": "", "partVer": "", "productName": ""},
]

# ════════════════════════════════════════════════
# SESSION STATE 初始化
# ════════════════════════════════════════════════
def init_state():
    if "data_store" not in st.session_state:
        st.session_state.data_store = []
    if "mold_list" not in st.session_state:
        st.session_state.mold_list = DEFAULT_MOLDS.copy()
    if "defect_types" not in st.session_state:
        st.session_state.defect_types = DEFAULT_TYPES.copy()
    if "saved_checklists" not in st.session_state:
        st.session_state.saved_checklists = []
    if "sync_status" not in st.session_state:
        st.session_state.sync_status = ("ing", "未连接")
    if "edit_record" not in st.session_state:
        st.session_state.edit_record = None
    if "active_tab" not in st.session_state:
        st.session_state.active_tab = "records"
    if "synced_once" not in st.session_state:
        st.session_state.synced_once = False

# ════════════════════════════════════════════════
# SUPABASE HTTP 工具
# ════════════════════════════════════════════════
def sb_headers():
    return {
        "apikey": SB_KEY,
        "Authorization": f"Bearer {SB_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=representation",
    }

def sb_get_all() -> list:
    url = f"{SB_URL}/rest/v1/{SB_TABLE}?select=*&order=id.desc"
    try:
        r = requests.get(url, headers=sb_headers(), timeout=10)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        st.session_state.sync_status = ("err", f"拉取失败：{e}")
        return []

def sb_upsert(record: dict) -> bool:
    row = local_to_cloud(record)
    url = f"{SB_URL}/rest/v1/{SB_TABLE}"
    try:
        r = requests.post(
            url,
            headers={**sb_headers(), "Prefer": "resolution=merge-duplicates,return=minimal"},
            json=row,
            timeout=10,
        )
        r.raise_for_status()
        return True
    except Exception as e:
        st.session_state.sync_status = ("err", f"上传失败：{e}")
        return False

def sb_delete(record_id: int) -> bool:
    url = f"{SB_URL}/rest/v1/{SB_TABLE}?id=eq.{record_id}"
    try:
        r = requests.delete(url, headers=sb_headers(), timeout=10)
        r.raise_for_status()
        return True
    except Exception as e:
        st.session_state.sync_status = ("err", f"删除失败：{e}")
        return False

def sb_test_connection() -> tuple[bool, str]:
    url = f"{SB_URL}/rest/v1/{SB_TABLE}?select=id&limit=1"
    try:
        r = requests.get(url, headers=sb_headers(), timeout=8)
        r.raise_for_status()
        return True, "✅ 连接成功"
    except Exception as e:
        return False, f"❌ 失败：{e}"

# ════════════════════════════════════════════════
# 字段映射
# ════════════════════════════════════════════════
def local_to_cloud(r: dict) -> dict:
    return {
        "id":          r["id"],
        "mold_id":     r.get("moldId", ""),
        "part_no":     r.get("partNo", ""),
        "part_ver":    r.get("partVer", ""),
        "type":        r.get("type", ""),
        "severity":    r.get("severity", "medium"),
        "date":        r.get("date", ""),
        "shift":       r.get("shift", ""),
        "machine_no":  r.get("machineNo", ""),
        "description": r.get("description", ""),
        "reason":      r.get("reason", ""),
        "images":      json.dumps(r.get("images", []), ensure_ascii=False),
        "updated_at":  datetime.utcnow().isoformat(),
    }

def cloud_to_local(r: dict) -> dict:
    images = []
    try:
        images = json.loads(r.get("images") or "[]")
    except Exception:
        pass
    return {
        "id":          r["id"],
        "moldId":      r.get("mold_id", ""),
        "partNo":      r.get("part_no", ""),
        "partVer":     r.get("part_ver", ""),
        "type":        r.get("type", ""),
        "severity":    r.get("severity", "medium"),
        "date":        r.get("date", ""),
        "shift":       r.get("shift", ""),
        "machineNo":   r.get("machine_no", ""),
        "description": r.get("description", ""),
        "reason":      r.get("reason", ""),
        "images":      images,
        "_synced":     True,
    }

# ════════════════════════════════════════════════
# 同步
# ════════════════════════════════════════════════
def sync_from_supabase():
    st.session_state.sync_status = ("ing", "同步中…")
    rows = sb_get_all()
    if not rows and st.session_state.sync_status[0] == "err":
        return
    cloud_records = [cloud_to_local(r) for r in rows]
    cloud_ids = {str(r["id"]) for r in cloud_records}
    # 保留本地未同步的
    local_pending = [r for r in st.session_state.data_store
                     if not r.get("_synced") and str(r["id"]) not in cloud_ids]
    st.session_state.data_store = cloud_records + local_pending
    pushed = 0
    for rec in local_pending:
        if sb_upsert(rec):
            rec["_synced"] = True
            pushed += 1
    msg = f"已同步 {len(cloud_records)} 条"
    if pushed:
        msg += f"，补传 {pushed} 条"
    st.session_state.sync_status = ("ok", msg)
    st.session_state.synced_once = True

# ════════════════════════════════════════════════
# DeepSeek 调用
# ════════════════════════════════════════════════
def ds_call(system_prompt: str, user_prompt: str) -> Optional[str]:
    try:
        r = requests.post(
            DS_API,
            headers={"Content-Type": "application/json", "Authorization": f"Bearer {DEEPSEEK_KEY}"},
            json={
                "model": "deepseek-reasoner",
                "max_tokens": 8000,
                "temperature": 1,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
            },
            timeout=120,
        )
        data = r.json()
        if "error" in data:
            st.error(f"DeepSeek 错误：{data['error']['message']}")
            return None
        return data["choices"][0]["message"]["content"]
    except Exception as e:
        st.error(f"请求失败：{e}")
        return None

# ════════════════════════════════════════════════
# 文件编号生成
# ════════════════════════════════════════════════
def gen_file_no() -> str:
    now = datetime.now()
    yy = str(now.year)[-2:]
    mm = str(now.month).zfill(2)
    seq = str(len(st.session_state.saved_checklists) + 1).zfill(3)
    return f"QCL-{yy}{mm}-{seq}"

# ════════════════════════════════════════════════
# Excel 导出 - 台账
# ════════════════════════════════════════════════
def export_excel(records: list) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "外观不良台账"

    header = ["序号", "模具号", "零件号", "版本", "不良类型", "严重程度",
              "发现日期", "班次", "机床号", "现象描述", "原因分析", "同步状态"]
    col_widths = [5, 12, 12, 8, 10, 10, 12, 8, 10, 32, 22, 8]

    # 表头样式
    hdr_fill = PatternFill("solid", fgColor="1565C0")
    hdr_font = Font(color="FFFFFF", bold=True, size=10)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ci, (h, w) in enumerate(zip(header, col_widths), 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.row_dimensions[1].height = 22

    for ri, d in enumerate(records, 2):
        row_data = [
            ri - 1,
            d.get("moldId", ""),
            d.get("partNo", ""),
            d.get("partVer", ""),
            d.get("type", ""),
            SEV_LABEL.get(d.get("severity", ""), d.get("severity", "")),
            d.get("date", ""),
            "白班" if d.get("shift") == "day" else ("夜班" if d.get("shift") == "night" else ""),
            d.get("machineNo", ""),
            d.get("description", ""),
            d.get("reason", ""),
            "已同步" if d.get("_synced") else "本地",
        ]
        for ci, val in enumerate(row_data, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=(ci in (10, 11)))
        if ri % 2 == 0:
            for ci in range(1, len(header) + 1):
                ws.cell(row=ri, column=ci).fill = PatternFill("solid", fgColor="F0F4FA")

    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ════════════════════════════════════════════════
# Excel 导出 - 缺陷清单
# ════════════════════════════════════════════════
def export_checklist_excel(cl: dict) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "缺陷清单"

    hdr_fill = PatternFill("solid", fgColor="1565C0")
    hdr_font = Font(color="FFFFFF", bold=True, size=10)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 元信息行
    meta = [
        ["文件编号", cl.get("fileNo", ""), "", "文件名称", cl.get("title", "")],
        ["创建日期", cl.get("date", ""), "", "版本", cl.get("revision", "A")],
        [],
    ]
    for row in meta:
        ws.append(row)

    # 表头
    cols = ["序号", "检查项目", "模具/零件号", "不良类型", "严重程度",
            "判定标准", "频次", "检查方法", "处置方法", "备注"]
    col_widths = [5, 26, 16, 12, 10, 26, 8, 16, 26, 14]
    header_row = ws.max_row + 1
    for ci, (h, w) in enumerate(zip(cols, col_widths), 1):
        cell = ws.cell(row=header_row, column=ci, value=h)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
        ws.column_dimensions[get_column_letter(ci)].width = w

    for item in cl.get("items", []):
        row = [
            item.get("no", ""),
            item.get("item", ""),
            item.get("mold", ""),
            item.get("type", ""),
            item.get("severity", ""),
            item.get("standard", ""),
            item.get("frequency", ""),
            item.get("method", ""),
            item.get("action", ""),
            item.get("remark", ""),
        ]
        ri = ws.max_row + 1
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ════════════════════════════════════════════════
# UI 工具
# ════════════════════════════════════════════════
def sev_badge(sev: str) -> str:
    colors = {
        "low":      ("#107c10", "#e6f4ea"),
        "medium":   ("#c55a11", "#fff4e5"),
        "high":     ("#c72a2a", "#fde8e8"),
        "critical": ("#7b5ea7", "#f0e8fa"),
    }
    fc, bg = colors.get(sev, ("#333", "#eee"))
    return f'<span style="background:{bg};color:{fc};padding:2px 9px;border-radius:12px;font-size:11px;font-weight:700;">{SEV_LABEL.get(sev, sev)}</span>'

def sync_badge():
    status, msg = st.session_state.sync_status
    color_map = {"ok": "#2e7d32", "ing": "#bf360c", "err": "#c62828"}
    bg_map    = {"ok": "#e6f4ea", "ing": "#fff4e5", "err": "#fde8e8"}
    color = color_map.get(status, "#333")
    bg    = bg_map.get(status, "#eee")
    st.markdown(
        f'<div style="display:inline-block;background:{bg};color:{color};padding:4px 12px;'
        f'border-radius:20px;font-size:12px;font-weight:600;border:1px solid {color}40;">'
        f'● {msg}</div>',
        unsafe_allow_html=True,
    )

# ════════════════════════════════════════════════
# 各功能区渲染
# ════════════════════════════════════════════════
def render_sidebar_form():
    """左侧数据录入表单"""
    edit = st.session_state.edit_record

    with st.form("record_form", clear_on_submit=False):
        st.markdown("#### 📝 数据录入")

        mold_options = ["— 请选择模具 —"] + [
            f"{m['name']}" + (f"  {m['partNo']}" if m.get("partNo") else "") +
            (f"  {m['productName']}" if m.get("productName") else "")
            for m in st.session_state.mold_list
        ]
        mold_idx = 0
        if edit and edit.get("moldId"):
            mold_ids = [m["id"] for m in st.session_state.mold_list]
            if edit["moldId"] in mold_ids:
                mold_idx = mold_ids.index(edit["moldId"]) + 1

        sel_mold = st.selectbox("模具编号 *", mold_options, index=mold_idx)

        # 带出模具信息
        mold_obj = None
        if sel_mold != "— 请选择模具 —":
            mold_name = sel_mold.split("  ")[0].strip()
            mold_obj = next((m for m in st.session_state.mold_list if m["name"] == mold_name), None)

        if mold_obj:
            st.caption(f"零件号：{mold_obj.get('partNo','—')} ｜ 版本：{mold_obj.get('partVer','—')} ｜ {mold_obj.get('productName','')}")

        col1, col2 = st.columns(2)
        with col1:
            rec_date = st.date_input("日期 *",
                value=datetime.strptime(edit["date"], "%Y-%m-%d").date() if edit and edit.get("date") else date.today())
            severity = st.selectbox("严重程度 *",
                options=list(SEV_LABEL.keys()),
                format_func=lambda x: SEV_LABEL[x],
                index=list(SEV_LABEL.keys()).index(edit.get("severity","medium")) if edit else 1)
        with col2:
            shift = st.selectbox("班次", ["", "day", "night"],
                format_func=lambda x: {"":"— 请选择 —","day":"☀ 白班","night":"🌙 夜班"}.get(x,x),
                index=["", "day", "night"].index(edit.get("shift","")) if edit else 0)
            defect_type = st.selectbox("不良类型 *", st.session_state.defect_types,
                index=st.session_state.defect_types.index(edit["type"]) if edit and edit.get("type") in st.session_state.defect_types else 0)

        machine_no = st.text_input("机床号", value=edit.get("machineNo","") if edit else "")
        description = st.text_area("现象描述 *", value=edit.get("description","") if edit else "", height=80)
        reason = st.text_area("原因分析", value=edit.get("reason","") if edit else "", height=60)

        images_up = st.file_uploader("上传图片", type=["jpg","jpeg","png","webp"],
                                      accept_multiple_files=True)

        col_s, col_c = st.columns(2)
        with col_s:
            submitted = st.form_submit_button("✔ 提交记录" if not edit else "✔ 保存修改",
                                              use_container_width=True, type="primary")
        with col_c:
            cancelled = st.form_submit_button("✖ 取消编辑", use_container_width=True)

    if cancelled:
        st.session_state.edit_record = None
        st.rerun()

    if submitted:
        if sel_mold == "— 请选择模具 —" and not edit:
            st.error("请先选择模具号")
            return
        if not description.strip():
            st.error("请填写现象描述")
            return

        mold_id = mold_obj["id"] if mold_obj else (edit.get("moldId","") if edit else "")
        part_no  = mold_obj.get("partNo","") if mold_obj else (edit.get("partNo","") if edit else "")
        part_ver = mold_obj.get("partVer","") if mold_obj else (edit.get("partVer","") if edit else "")

        # 处理图片
        images = edit.get("images", []) if edit else []
        if images_up:
            import base64
            for f in images_up:
                b64 = base64.b64encode(f.read()).decode()
                mime = f.type or "image/jpeg"
                images.append(f"data:{mime};base64,{b64}")

        record = {
            "id":          edit["id"] if edit else int(time.time() * 1000),
            "moldId":      mold_id,
            "partNo":      part_no,
            "partVer":     part_ver,
            "type":        defect_type,
            "severity":    severity,
            "date":        str(rec_date),
            "shift":       shift,
            "machineNo":   machine_no.strip(),
            "description": description.strip(),
            "reason":      reason.strip(),
            "images":      images,
            "_synced":     False,
        }

        with st.spinner("提交中…"):
            ok = sb_upsert(record)
        if ok:
            record["_synced"] = True
            st.session_state.sync_status = ("ok", "已同步")
        else:
            st.warning("本地已保存，云端同步失败")

        ds = st.session_state.data_store
        if edit:
            idx = next((i for i, r in enumerate(ds) if r["id"] == edit["id"]), None)
            if idx is not None:
                ds[idx] = record
        else:
            ds.insert(0, record)

        st.session_state.edit_record = None
        st.success("✅ 记录已保存")
        st.rerun()


def render_records_tab():
    """台账卡片视图 + 筛选"""
    # ── 筛选栏 ──
    with st.container():
        fc1, fc2, fc3, fc4, fc5 = st.columns([2, 1.5, 1.5, 1.5, 1.5])
        mold_opts = ["全部模具"] + [m["name"] for m in st.session_state.mold_list]
        with fc1:
            f_mold = st.selectbox("模具", mold_opts, label_visibility="collapsed")
        with fc2:
            f_sev = st.selectbox("严重程度", ["全部"] + list(SEV_LABEL.keys()),
                format_func=lambda x: "全部严重程度" if x=="全部" else SEV_LABEL.get(x,x),
                label_visibility="collapsed")
        with fc3:
            f_shift = st.selectbox("班次", ["全部","day","night"],
                format_func=lambda x: {"全部":"全部班次","day":"白班","night":"夜班"}.get(x,x),
                label_visibility="collapsed")
        with fc4:
            f_from = st.date_input("开始日期", value=None, label_visibility="collapsed")
        with fc5:
            f_to   = st.date_input("结束日期", value=None, label_visibility="collapsed")

    # ── 过滤 ──
    ds = st.session_state.data_store
    filtered = [d for d in ds if
        (f_mold == "全部模具" or d.get("moldId") == f_mold.split("  ")[0]) and
        (f_sev  == "全部"     or d.get("severity") == f_sev) and
        (f_shift == "全部"    or d.get("shift") == f_shift) and
        (not f_from or d.get("date","") >= str(f_from)) and
        (not f_to   or d.get("date","") <= str(f_to))
    ]

    # ── 统计栏 ──
    sev_cnt = {k: sum(1 for d in filtered if d.get("severity")==k) for k in SEV_LABEL}
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📋 总记录", len(filtered))
    c2.metric("🔴 严重/致命", sev_cnt["high"] + sev_cnt["critical"])
    c3.metric("🟠 一般", sev_cnt["medium"])
    c4.metric("🟢 轻微", sev_cnt["low"])

    # ── 操作按钮 ──
    op1, op2, op3 = st.columns([1, 1, 4])
    with op1:
        if st.button("🔄 同步云端", use_container_width=True):
            with st.spinner("同步中…"):
                sync_from_supabase()
            st.rerun()
    with op2:
        if filtered:
            xlsx_bytes = export_excel(filtered)
            st.download_button("📥 导出Excel", data=xlsx_bytes,
                file_name=f"外观不良台账_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

    # ── 卡片 ──
    if not filtered:
        st.info("暂无记录，请在左侧录入数据")
        return

    cols_per_row = 3
    rows = [filtered[i:i+cols_per_row] for i in range(0, len(filtered), cols_per_row)]
    for row in rows:
        cols = st.columns(cols_per_row)
        for col, d in zip(cols, row):
            with col:
                sev = d.get("severity","medium")
                sev_colors = {
                    "low":("#107c10","#e6f4ea"),
                    "medium":("#c55a11","#fff4e5"),
                    "high":("#c72a2a","#fde8e8"),
                    "critical":("#7b5ea7","#f0e8fa"),
                }
                fc, bg = sev_colors.get(sev, ("#333","#eee"))
                synced_tag = "☁ 已同步" if d.get("_synced") else "💾 本地"
                shift_txt = {"day":"☀ 白班","night":"🌙 夜班"}.get(d.get("shift",""),"")
                imgs = d.get("images",[])

                st.markdown(f"""
<div style="background:#fff;border:1px solid #d4dae6;border-radius:10px;
  overflow:hidden;margin-bottom:4px;box-shadow:0 1px 3px rgba(0,0,0,.09);">
  <div style="background:{bg};padding:8px 12px;display:flex;justify-content:space-between;align-items:center;">
    <span style="color:{fc};font-weight:700;font-size:11px;">{SEV_LABEL.get(sev,sev)}</span>
    <span style="font-size:10px;color:#1565c0;font-weight:700;font-family:monospace;">{d.get('moldId','')}</span>
  </div>
  <div style="padding:10px 12px;">
    <div style="font-size:13px;font-weight:700;color:#1e2a3b;margin-bottom:4px;">{d.get('type','')}</div>
    <div style="font-size:11px;color:#445068;line-height:1.5;
      display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;">
      {d.get('description','')}
    </div>
    <div style="margin-top:7px;font-size:10px;color:#7f8fa8;display:flex;gap:8px;flex-wrap:wrap;">
      <span>📅 {d.get('date','')}</span>
      {f'<span>{shift_txt}</span>' if shift_txt else ''}
      {f'<span>🔧 {d.get("machineNo","")}</span>' if d.get("machineNo") else ''}
      <span style="color:{'#2e7d32' if d.get('_synced') else '#bf360c'}">{synced_tag}</span>
      {f'<span>🖼 {len(imgs)}张</span>' if imgs else ''}
    </div>
  </div>
</div>""", unsafe_allow_html=True)

                bc1, bc2 = st.columns(2)
                with bc1:
                    if st.button("✏ 编辑", key=f"edit_{d['id']}", use_container_width=True):
                        st.session_state.edit_record = d
                        st.rerun()
                with bc2:
                    if st.button("🗑 删除", key=f"del_{d['id']}", use_container_width=True):
                        st.session_state.data_store = [r for r in st.session_state.data_store if r["id"] != d["id"]]
                        with st.spinner("云端删除…"):
                            sb_delete(d["id"])
                        st.rerun()

                # 图片展示
                if imgs:
                    with st.expander(f"查看图片（{len(imgs)}张）"):
                        for src in imgs[:4]:
                            st.image(src, use_column_width=True)


def render_checklist_tab():
    """AI 缺陷清单"""
    c1, c2 = st.columns([3, 1])
    with c1:
        st.markdown("##### 🤖 AI 缺陷检查清单")
    with c2:
        if st.session_state.saved_checklists:
            xlsx_bytes = export_checklist_excel(st.session_state.saved_checklists[0])
            st.download_button("📥 导出清单", data=xlsx_bytes,
                file_name=f"缺陷清单_{st.session_state.saved_checklists[0].get('fileNo','')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

    if st.button("🤖 AI 生成缺陷清单", type="primary"):
        ds = st.session_state.data_store
        summary = "\n".join([
            f"模具:{d.get('moldId','')} 零件:{d.get('partNo','-')} 类型:{d.get('type','')} "
            f"程度:{SEV_LABEL.get(d.get('severity',''),'')} 描述:{d.get('description','')}"
            for d in ds[:60]
        ])
        mold_info = "、".join([
            f"{m['name']}(零件:{m.get('partNo','-')} 版本:{m.get('partVer','-')})"
            for m in st.session_state.mold_list
        ])
        file_no = gen_file_no()
        today   = datetime.now().strftime("%Y/%m/%d")

        with st.spinner("DeepSeek-R1 分析中，请稍候…"):
            result = ds_call(
                "你是资深精密铸造质量工程师，熟悉IATF16949。请严格只输出合法JSON，不加任何说明文字或markdown符号。",
                f"根据以下不良记录，生成正式缺陷检查清单：\n模具：{mold_info}\n记录：\n{summary}\n\n"
                f'输出JSON格式：{{"title":"外观不良检查清单","fileNo":"{file_no}","date":"{today}",'
                f'"revision":"A","scope":"适用范围","items":[{{"no":1,"item":"检查项目","mold":"模具号",'
                f'"type":"类型","severity":"程度","standard":"判定标准","frequency":"频次",'
                f'"method":"方法","action":"处置","remark":"备注"}}],"notes":["注意事项"]}}',
            )
        if result:
            try:
                parsed = json.loads(result.replace("```json","").replace("```","").strip())
                st.session_state.saved_checklists.insert(0, parsed)
                st.success("✅ 清单已生成")
                st.rerun()
            except Exception:
                st.markdown(f"```\n{result}\n```")

    # 渲染最新清单
    if st.session_state.saved_checklists:
        cl = st.session_state.saved_checklists[0]
        st.markdown(f"""
**{cl.get('title','')}** `{cl.get('fileNo','')}` Rev.{cl.get('revision','A')} · {cl.get('date','')}

> {cl.get('scope','')}
""")
        items = cl.get("items", [])
        if items:
            import pandas as pd
            df = pd.DataFrame([{
                "序号": it.get("no",""), "检查项目": it.get("item",""),
                "模具/零件": it.get("mold","-"), "类型": it.get("type","-"),
                "程度": it.get("severity","-"), "判定标准": it.get("standard","-"),
                "频次": it.get("frequency","-"), "方法": it.get("method","-"),
                "处置": it.get("action","-"), "备注": it.get("remark",""),
            } for it in items])
            st.dataframe(df, use_container_width=True, hide_index=True)

        notes = cl.get("notes", [])
        if notes:
            st.markdown("**注意事项：**")
            for i, n in enumerate(notes, 1):
                st.markdown(f"{i}. {n}")
    else:
        st.info("暂无清单，点击「AI 生成缺陷清单」")

    # 上传模板
    with st.expander("📄 上传缺陷清单模板（Excel/CSV/TXT）"):
        up = st.file_uploader("选择文件", type=["xlsx","xls","csv","txt"], key="template_up")
        if up and st.button("🤖 AI 解析模板"):
            if up.name.endswith((".xlsx",".xls")):
                import openpyxl as ox
                wb = ox.load_workbook(io.BytesIO(up.read()))
                ws = wb.active
                content = "\n".join([
                    ",".join([str(c.value or "") for c in row])
                    for row in ws.iter_rows()
                ])
            else:
                content = up.read().decode("utf-8", errors="ignore")

            file_no = gen_file_no()
            today   = datetime.now().strftime("%Y/%m/%d")
            with st.spinner("DeepSeek-R1 解析中…"):
                result = ds_call(
                    "你是质量工程师。将模板整理为标准缺陷清单JSON，严格只输出合法JSON，不加任何说明。",
                    f'模板内容：\n{content[:4000]}\n\n输出：{{"title":"...","fileNo":"{file_no}",'
                    f'"date":"{today}","revision":"A","scope":"...","items":[{{"no":1,"item":"...",'
                    f'"mold":"...","type":"...","severity":"...","standard":"...","frequency":"...",'
                    f'"method":"...","action":"...","remark":"..."}}],"notes":[]}}',
                )
            if result:
                try:
                    parsed = json.loads(result.replace("```json","").replace("```","").strip())
                    st.session_state.saved_checklists.insert(0, parsed)
                    st.success("✅ 解析成功")
                    st.rerun()
                except Exception:
                    st.code(result)


def render_report_tab():
    """AI 质量报告"""
    st.markdown("##### 📊 AI 质量分析报告")
    rc1, rc2 = st.columns(2)
    with rc1:
        report_type = st.selectbox("报告周期", ["本周", "本月", "自定义"])
    
    from_date, to_date = None, None
    if report_type == "自定义":
        d1, d2 = st.columns(2)
        with d1:
            from_date = st.date_input("开始日期")
        with d2:
            to_date = st.date_input("结束日期")

    if st.button("🤖 AI 生成报告", type="primary"):
        now = datetime.now()
        if report_type == "本周":
            day = now.weekday()
            from_date = date(now.year, now.month, now.day - day)
            to_date   = date.today()
        elif report_type == "本月":
            from_date = date(now.year, now.month, 1)
            to_date   = date.today()
        
        if not from_date or not to_date:
            st.warning("请选择日期范围")
            return

        ds = st.session_state.data_store
        data = [d for d in ds if str(from_date) <= d.get("date","") <= str(to_date)]
        by_sev = {k: sum(1 for d in data if d.get("severity")==k) for k in SEV_LABEL}
        summary = "\n".join([
            f"[{d.get('date','')}]{'白班' if d.get('shift')=='day' else '夜班' if d.get('shift')=='night' else ''} "
            f"模具:{d.get('moldId','')} 类型:{d.get('type','')} 程度:{SEV_LABEL.get(d.get('severity',''),'')} "
            f"机床:{d.get('machineNo','-')} 描述:{d.get('description','')} "
            f"{'分析:'+d.get('reason') if d.get('reason') else ''}"
            for d in data
        ])

        with st.spinner("DeepSeek-R1 撰写报告，请稍候（约30-60秒）…"):
            result = ds_call(
                "你是精密铸造工厂质量经理，负责撰写专业质量周报/月报用于管理层会议汇报。报告需要数据分析、趋势研判、风险预警和改善建议，语言简洁专业，结构清晰。",
                f"请根据 {from_date} 至 {to_date} 的外观不良数据生成质量分析报告：\n\n"
                f"统计摘要：总计{len(data)}条 | 轻微{by_sev['low']} | 一般{by_sev['medium']} | "
                f"严重{by_sev['high']} | 致命{by_sev['critical']}\n\n"
                f"详细记录：\n{summary or '（该时段无记录）'}\n\n"
                f"报告要求：①执行摘要 ②数据概览（表格） ③主要不良分析 ④趋势与风险 ⑤改善建议 ⑥重点关注项",
            )
        if result:
            st.text_area("📋 报告内容", value=result, height=520)
            txt_bytes = result.encode("utf-8")
            st.download_button("⬇ 导出TXT", data=txt_bytes,
                file_name=f"质量报告_{date.today()}.txt", mime="text/plain")


def render_mold_manager():
    """模具管理"""
    st.markdown("##### 🔧 模具管理")
    with st.form("add_mold"):
        mc1, mc2, mc3, mc4, mc5 = st.columns([2,2,1,2,1])
        with mc1: m_name = st.text_input("模具号 *", placeholder="如 1507#")
        with mc2: m_part = st.text_input("零件号", placeholder="如 AD77311211")
        with mc3: m_ver  = st.text_input("版本", placeholder="A-1")
        with mc4: m_pname= st.text_input("产品名称（选填）")
        with mc5:
            st.markdown("<br>", unsafe_allow_html=True)
            add_m = st.form_submit_button("添加", use_container_width=True)
    if add_m:
        m_name = m_name.strip().upper()
        if not m_name:
            st.error("请输入模具号")
        elif any(m["id"] == m_name for m in st.session_state.mold_list):
            st.error("该模具号已存在")
        else:
            st.session_state.mold_list.append({
                "id": m_name, "name": m_name,
                "partNo": m_part.strip(), "partVer": m_ver.strip(),
                "productName": m_pname.strip()
            })
            st.success("已添加")
            st.rerun()

    if st.session_state.mold_list:
        import pandas as pd
        df = pd.DataFrame([{
            "模具号": m["name"],
            "零件号": m.get("partNo",""),
            "版本": m.get("partVer",""),
            "产品名称": m.get("productName",""),
        } for m in st.session_state.mold_list])
        st.dataframe(df, use_container_width=True, hide_index=True)

        del_idx = st.selectbox("删除模具", ["— 请选择 —"] + [m["name"] for m in st.session_state.mold_list])
        if del_idx != "— 请选择 —" and st.button(f"🗑 确认删除 {del_idx}"):
            st.session_state.mold_list = [m for m in st.session_state.mold_list if m["name"] != del_idx]
            st.rerun()


def render_type_manager():
    """不良类型管理"""
    st.markdown("##### 🏷 不良类型管理")
    with st.form("add_type"):
        t1, t2 = st.columns([3,1])
        with t1: new_type = st.text_input("新类型名称", label_visibility="collapsed", placeholder="输入新类型…")
        with t2: add_t = st.form_submit_button("添加", use_container_width=True)
    if add_t and new_type.strip():
        if new_type.strip() not in st.session_state.defect_types:
            st.session_state.defect_types.append(new_type.strip())
            st.rerun()

    for i, t in enumerate(st.session_state.defect_types):
        c1, c2 = st.columns([4,1])
        c1.write(t)
        if c2.button("删除", key=f"del_type_{i}"):
            st.session_state.defect_types.pop(i)
            st.rerun()

# ════════════════════════════════════════════════
# 主程序
# ════════════════════════════════════════════════
def main():
    st.set_page_config(
        page_title="外观不良智能管理系统 V3",
        page_icon="🏭",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # 注入最小CSS
    st.markdown("""
    <style>
    [data-testid="stSidebar"] { min-width: 320px; max-width: 340px; }
    .block-container { padding-top: 1rem; padding-bottom: 1rem; }
    div[data-testid="metric-container"] {
        background: #f8f9fb; border: 1px solid #d4dae6;
        border-radius: 8px; padding: 10px 14px; box-shadow: 0 1px 3px rgba(0,0,0,.07);
    }
    .stButton > button { border-radius: 7px; font-size: 12px; }
    .stDataFrame { font-size: 12px; }
    </style>
    """, unsafe_allow_html=True)

    init_state()

    # 首次自动同步
    if not st.session_state.synced_once:
        with st.spinner("连接 Supabase…"):
            sync_from_supabase()

    # ── 侧边栏 ──
    with st.sidebar:
        # 品牌头
        st.markdown("""
<div style="background:linear-gradient(135deg,#1565c0,#00695c);padding:12px 14px;border-radius:10px;margin-bottom:14px;">
  <div style="color:#fff;font-size:15px;font-weight:700;">🏭 外观不良管理 V3</div>
  <div style="color:rgba(255,255,255,.7);font-size:10px;font-family:monospace;">APPEARANCE DEFECT SYSTEM · PY</div>
</div>""", unsafe_allow_html=True)

        # 同步状态
        sync_badge()
        st.markdown("")

        render_sidebar_form()

        st.divider()
        with st.expander("🔧 模具管理"):
            render_mold_manager()
        with st.expander("🏷 类型管理"):
            render_type_manager()
        with st.expander("☁ 云端连接测试"):
            if st.button("🔌 测试 Supabase 连接"):
                ok, msg = sb_test_connection()
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)

    # ── 主区域 ──
    # Topbar
    tb1, tb2 = st.columns([3, 1])
    with tb1:
        st.markdown("## 外观不良智能管理系统")
    with tb2:
        sync_badge()

    tab1, tab2, tab3 = st.tabs(["📋 不良台账", "📄 AI缺陷清单", "📊 AI质量报告"])
    with tab1:
        render_records_tab()
    with tab2:
        render_checklist_tab()
    with tab3:
        render_report_tab()


if __name__ == "__main__":
    main()
