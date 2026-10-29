import io
import json
import os
from datetime import date, datetime
from urllib.parse import quote
import altair as alt
import pandas as pd
import streamlit as st

from core import (
    prepare_wo_dataframe, coordinator_map_from_excel, coordinators_master_from_excel, fmes_from_excel,
    make_fme_message, make_fme_group_message, feedback_json_records, clean_text
)
from database import (
    get_client, upsert_many, replace_many, select_all, coordinator_for_site, fme_for_cedula,
    upsert_activity, insert_cycle, message_status, set_message_sent,
    save_ows_feedback, set_closed_management
)

st.set_page_config(page_title="Gestión Rechazos FME V12", page_icon="🛠️", layout="wide")

def apply_enterprise_style():
    """Visual system for a clear, professional operations dashboard."""
    st.markdown("""
    <style>
      :root { --navy:#102a43; --blue:#1769aa; --cyan:#12b8c4; --ink:#243b53; --muted:#627d98; --line:#d9e2ec; }
      .stApp { background: linear-gradient(135deg, #f7fafc 0%, #eef5f9 100%); color: var(--ink); }
      [data-testid="stHeader"] { background: rgba(247,250,252,.88); border-bottom: 1px solid var(--line); }
      [data-testid="stSidebar"] { background: linear-gradient(180deg, #102a43 0%, #163d5c 100%); }
      [data-testid="stSidebar"] * { color: #f7fafc !important; }
      [data-testid="stSidebar"] .stRadio label { border-radius: 8px; padding: .42rem .55rem; margin: .08rem 0; }
      [data-testid="stSidebar"] .stRadio label:hover { background: rgba(255,255,255,.12); }
      h1, h2, h3 { color: var(--navy) !important; letter-spacing: -.02em; }
      h1 { font-weight: 800 !important; } h2 { font-weight: 750 !important; padding-top: .4rem; }
      [data-testid="stMetric"] { background: #fff; border: 1px solid var(--line); border-radius: 14px; padding: 1rem 1.15rem; box-shadow: 0 3px 12px rgba(16,42,67,.06); }
      [data-testid="stMetricLabel"] { color: var(--muted) !important; font-weight: 650; }
      [data-testid="stMetricValue"] { color: var(--navy) !important; font-weight: 800; }
      .stButton > button, .stDownloadButton > button, .stLinkButton > a { border-radius: 8px; border: 0; font-weight: 700; transition: all .18s ease; }
      .stButton > button[kind="primary"] { background: linear-gradient(135deg, var(--blue), #0f83a7); box-shadow: 0 4px 10px rgba(23,105,170,.22); }
      .stButton > button[kind="primary"]:hover { transform: translateY(-1px); filter: brightness(1.05); }
      [data-testid="stDataFrame"] { border: 1px solid var(--line); border-radius: 12px; overflow: hidden; background: #fff; }
      [data-testid="stFileUploaderDropzone"] { background: #fff !important; border: 1.5px dashed #9fb3c8 !important; border-radius: 12px !important; }
      [data-testid="stFileUploaderDropzone"] * { color: var(--ink) !important; }
      [data-testid="stFileUploaderDropzone"] button { background: #edf6ff !important; border: 1px solid #9fc5e8 !important; color: var(--blue) !important; border-radius: 8px !important; }
      [data-testid="stFileUploaderFile"] { background: #edf6ff !important; border: 1px solid #b9d6ee !important; border-radius: 10px !important; box-shadow: none !important; }
      [data-testid="stFileUploaderFile"] * { color: var(--navy) !important; }
      [data-testid="stFileUploaderFile"] button { background: transparent !important; border: 0 !important; }
      .stTextInput input, .stTextArea textarea, [data-baseweb="select"] > div { background: #fff !important; color: var(--ink) !important; border-radius: 8px !important; border-color: #bcccdc !important; }
      [data-testid="stTextInput"] label, [data-testid="stTextArea"] label { color: var(--ink) !important; font-weight: 650 !important; }
      [data-testid="stTextInput"] button { color: var(--ink) !important; }
      [data-testid="stTextInput"] [data-baseweb="input"], [data-testid="stTextInput"] [data-baseweb="base-input"] { background: #fff !important; border: 1.5px solid #9fb3c8 !important; border-radius: 12px !important; box-shadow: none !important; }
      [data-testid="stTextInput"] [data-baseweb="input"] input, [data-testid="stTextInput"] [data-baseweb="base-input"] input { background: transparent !important; color: var(--ink) !important; -webkit-text-fill-color: var(--ink) !important; }
      [data-baseweb="input"] > div, [data-baseweb="base-input"], [data-baseweb="select"] > div,
      [data-testid="stDateInput"] input, [data-testid="stTimeInput"] input,
      [data-testid="stNumberInput"] input, [data-testid="stSelectbox"] input,
      [data-testid="stMultiSelect"] input { background: #fff !important; color: var(--ink) !important; border-color: #bcccdc !important; }
      [data-baseweb="input"] input, [data-baseweb="select"] input, [data-baseweb="textarea"] textarea { color: var(--ink) !important; -webkit-text-fill-color: var(--ink) !important; }
      [data-testid="stDateInput"] button, [data-testid="stTimeInput"] button,
      [data-testid="stNumberInput"] button, [data-testid="stSelectbox"] svg, [data-testid="stMultiSelect"] svg { color: var(--ink) !important; fill: var(--ink) !important; }
      [data-testid="stDateInput"] label, [data-testid="stTimeInput"] label,
      [data-testid="stNumberInput"] label, [data-testid="stSelectbox"] label, [data-testid="stMultiSelect"] label { color: var(--ink) !important; font-weight: 650 !important; }
      [data-baseweb="tag"] { background: #e6f0f8 !important; color: var(--navy) !important; }
      [data-testid="stExpander"] { background: #fff; border: 1px solid var(--line); border-radius: 10px; }
      [data-testid="stAlert"] { border-radius: 10px; } hr { border-color: var(--line); margin: 2rem 0; }
      #MainMenu, footer { visibility: hidden; }
    </style>
    """, unsafe_allow_html=True)


apply_enterprise_style()

TO_EMAILS = ["griselda.dolores.castillo@eysglobal.co", "camilo.jimenez@huawei.com"]
CC_EMAILS = ["sebastian.rolong_costa@manpowercolombia.com", "ali.villa_costa@manpowercolombia.com"]


def coordinator_filter_options(values):
    """Keep operational catch-all coordinators at the end of filter lists."""
    last = {"isabel hernandez", "fija", "tlf"}
    options = sorted({str(value).strip() for value in values if str(value).strip()}, key=str.casefold)
    return [value for value in options if value.casefold() not in last] + [value for value in options if value.casefold() in last]


def short_fme_label(value):
    """Compact label for charts; keeps the full FME name in data views."""
    parts = clean_text(value).split()
    if len(parts) < 3:
        return " ".join(parts)
    surname_parts = 2 if parts[0].casefold() in {"de", "del", "la", "las", "los"} and len(parts) >= 4 else 1
    first_name_index = surname_parts + 1
    if first_name_index >= len(parts):
        return " ".join(parts[:2])
    return f"{parts[first_name_index]} {' '.join(parts[:surname_parts])}"


def feedback_stage_label(value):
    try:
        quantity = int(value)
    except (TypeError, ValueError):
        return "Sin ciclo definido"
    if quantity == 0:
        return "Auditoría inicial"
    if quantity == 1:
        return "Primer feedback"
    if quantity == 2:
        return "Segundo feedback"
    if quantity == 3:
        return "Tercer feedback"
    return f"Feedback #{quantity}"


def result_label(value):
    labels = {
        "RECHAZADO": "Rechazado",
        "RECHAZADO_NUEVAMENTE": "Rechazado nuevamente",
        "APROBADO": "Aprobado",
        "PENDIENTE_REVISION_FEEDBACK": "Pendiente de revisión",
        "SIN_CLASIFICAR": "Sin clasificar",
    }
    text = clean_text(value).upper()
    return labels.get(text, clean_text(value).replace("_", " ").title() or "Sin clasificar")


def render_horizontal_chart(data, label_column, value_column, color, height=None):
    """A compact horizontal bar chart with readable values and labels."""
    if data.empty:
        st.info("No hay datos para mostrar con el filtro actual.")
        return
    chart_height = height or max(230, len(data) * 42)
    base = alt.Chart(data).encode(
        y=alt.Y(f"{label_column}:N", sort="-x", title=None, axis=alt.Axis(labelLimit=230, labelFontSize=12)),
        x=alt.X(f"{value_column}:Q", title=None, axis=alt.Axis(gridColor="#d9e2ec", tickMinStep=1)),
        tooltip=[alt.Tooltip(f"{label_column}:N", title=label_column), alt.Tooltip(f"{value_column}:Q", title=value_column, format=".0f")],
    )
    bars = base.mark_bar(color=color, cornerRadiusEnd=7, size=24)
    labels = base.mark_text(align="left", baseline="middle", dx=7, color="#102a43", fontWeight=700).encode(
        text=alt.Text(f"{value_column}:Q", format=".0f")
    )
    st.altair_chart(
        (bars + labels).properties(height=chart_height).configure_view(strokeOpacity=0).configure_axis(domain=False),
        use_container_width=True,
    )


def render_selectable_coordinator_chart(data):
    """Interactive coordinator chart that returns the coordinator clicked by the user."""
    chart_data = data.rename(columns={"coordinator": "Coordinador", "rechazos": "Rechazos"})
    selection = alt.selection_point(name="coordinator_click", fields=["Coordinador"], on="click")
    base = alt.Chart(chart_data).encode(
        y=alt.Y("Coordinador:N", sort="-x", title=None, axis=alt.Axis(labelLimit=230, labelFontSize=12)),
        x=alt.X("Rechazos:Q", title=None, axis=alt.Axis(gridColor="#d9e2ec", tickMinStep=1)),
        tooltip=[alt.Tooltip("Coordinador:N", title="Coordinador"), alt.Tooltip("Rechazos:Q", title="Rechazos", format=".0f")],
    )
    bars = base.mark_bar(color="#1769aa", cornerRadiusEnd=7, size=24).add_params(selection)
    labels = base.mark_text(align="left", baseline="middle", dx=7, color="#102a43", fontWeight=700).encode(text=alt.Text("Rechazos:Q", format=".0f"))
    event = st.altair_chart(
        (bars + labels).properties(height=max(300, len(chart_data) * 42)).configure_view(strokeOpacity=0).configure_axis(domain=False),
        use_container_width=True,
        key="dashboard_coordinator_chart",
        on_select="rerun",
        selection_mode="coordinator_click",
    )
    points = event.selection.get("coordinator_click", []) if event else []
    return points[0].get("Coordinador") if points else None


def login_gate():
    try:
        pw = st.secrets.get("APP_PASSWORD") or os.getenv("APP_PASSWORD", "")
    except Exception:
        pw = os.getenv("APP_PASSWORD", "")
    if not pw:
        return True
    if st.session_state.get("auth_ok"):
        return True
    st.title("🛠️ Gestión de Rechazos FME")
    entered = st.text_input("Contraseña", type="password")
    if st.button("Ingresar", type="primary"):
        if entered == pw:
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta")
    return False


def db_ok():
    try:
        get_client()
        return True
    except Exception as e:
        st.error(str(e))
        st.info("Configura los Secrets y ejecuta schema.sql en Supabase antes de usar la app.")
        return False


def activities_df():
    data = select_all("activities", order="updated_at", desc=True)
    return pd.DataFrame(data)


def feedback_df():
    data = select_all("ows_feedbacks", order="created_at", desc=True)
    return pd.DataFrame(data)


def render_dashboard():
    df = activities_df()
    st.header("📊 Dashboard")
    if df.empty:
        st.info("Aún no hay actividades en la base de datos.")
        return
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Actividades", len(df))
    c2.metric("Rechazadas", int(df["current_result"].astype(str).str.contains("RECHAZ", na=False).sum()))
    c3.metric("Aprobadas", int((df["current_result"] == "APROBADO").sum()))
    c4.metric("Closed", int(df["task_status"].astype(str).str.lower().eq("closed").sum()))
    st.dataframe(df[[c for c in ["assignment_id","task_id","site_id","coordinator","fme_nombre","feedback_qty","current_result","current_remark","task_status"] if c in df.columns]], use_container_width=True, hide_index=True)


def import_wo():
    st.header("📥 Importar WO Result Query")
    st.caption("Carga el archivo, valida el resumen y guarda las actividades en Supabase.")
    up = st.file_uploader("WO Result Query (.xlsx)", type=["xlsx"], key="wo")
    if not up:
        return
    try:
        df = prepare_wo_dataframe(up)
    except Exception as e:
        st.error(f"No se pudo leer el archivo: {e}")
        return

    total, rejected, approved = st.columns(3)
    total.metric("Actividades detectadas", f"{len(df):,}")
    rejected.metric("Rechazos detectados", int(df["current_result"].astype(str).str.contains("RECHAZ", na=False).sum()))
    approved.metric("Aprobadas", int(df["current_result"].astype(str).str.upper().eq("APROBADO").sum()))
    save_requested = st.button("Guardar/actualizar en Supabase", type="primary", key="save_wo", use_container_width=True)

    with st.expander("Vista previa del archivo", expanded=False):
        columns = ["assignment_id", "site_id", "fme_nombre", "feedback_qty", "current_result", "remark_source", "current_remark"]
        st.dataframe(df[[c for c in columns if c in df]].head(50), use_container_width=True, hide_index=True, height=360)

    if save_requested:
        progress = st.progress(0, text="Guardando actividades en Supabase...")
        try:
            for index, (_, row) in enumerate(df.iterrows(), start=1):
                record = row.to_dict()
                coordinator = coordinator_for_site(record["site_id"])
                upsert_activity(record, coordinator)
                insert_cycle(record, coordinator)
                progress.progress(index / len(df), text=f"Guardando {index:,} de {len(df):,} actividades...")
            progress.empty()
            st.success(f"{len(df):,} actividades procesadas. Los ciclos anteriores no se eliminan.")
        except Exception as e:
            progress.empty()
            st.error(f"No se pudo guardar el archivo en Supabase: {e}")


def render_messages():
    st.header("📲 Mensajes FME")
    df = activities_df()
    if df.empty: return
    df = df[df["current_result"].astype(str).str.contains("RECHAZ", na=False)].copy()
    coords = coordinator_filter_options(df["coordinator"].dropna().unique().tolist())
    selected = st.multiselect("Coordinador", coords)
    if selected: df = df[df["coordinator"].isin(selected)]
    state = st.selectbox("Estado mensaje", ["Todos","Pendientes","Enviados"])
    cards=[]
    for _, r in df.iterrows():
        msg_state = message_status(r["assignment_id"], r.get("feedback_qty",0))
        sent = bool(msg_state and msg_state.get("sent"))
        if state == "Pendientes" and sent: continue
        if state == "Enviados" and not sent: continue
        fme = fme_for_cedula(clean_text(r.get("fme_cedula")))
        cards.append((r, fme, sent, msg_state))
    c1,c2,c3 = st.columns(3)
    c1.metric("Visibles", len(cards)); c2.metric("Pendientes", sum(not x[2] for x in cards)); c3.metric("Enviados", sum(x[2] for x in cards))
    grouped = {}
    for card in cards:
        record, fme, _, _ = card
        key = clean_text(record.get("fme_cedula")) or clean_text(record.get("fme_nombre")) or "FME sin identificar"
        grouped.setdefault(key, []).append(card)

    for _, fme_cards in sorted(grouped.items(), key=lambda group: clean_text(group[1][0][0].get("fme_nombre")).casefold()):
        first_record, first_fme, _, _ = fme_cards[0]
        fme_name = first_record.get("fme_nombre", "FME sin identificar")
        phone = (first_fme or {}).get("celular", "Sin celular registrado")
        pending = sum(not card[2] for card in fme_cards)
        with st.expander(f"{fme_name} · {len(fme_cards)} actividad(es) · {pending} pendiente(s)", expanded=False):
            st.caption(f"Coordinador: {first_record.get('coordinator','')} · Celular: {phone}")
            pending_records = [card[0] for card in fme_cards if not card[2]]
            if pending_records:
                grouped_message = make_fme_group_message(pending_records, fme_name)
                st.markdown("**Mensaje consolidado para el FME**")
                st.text_area("Resumen de rechazos", grouped_message, height=280, key=f"group_msg_{clean_text(first_record.get('fme_cedula')) or clean_text(fme_name)}")
                if phone and phone != "Sin celular registrado":
                    st.link_button("📲 Abrir WhatsApp con todos los rechazos pendientes", f"https://wa.me/{phone}?text={quote(grouped_message)}", use_container_width=True)
                if st.button(f"✅ Marcar {len(pending_records)} actividad(es) como enviadas", key=f"group_sent_{clean_text(first_record.get('fme_cedula')) or clean_text(fme_name)}", type="primary", use_container_width=True):
                    for group_record, _, group_sent, _ in fme_cards:
                        if not group_sent:
                            set_message_sent(group_record["assignment_id"], group_record.get("feedback_qty", 0), True, grouped_message)
                    st.rerun()
            else:
                st.success("Todas las actividades visibles de este FME ya están marcadas como enviadas.")
                if st.button(f"↩️ Marcar {len(fme_cards)} actividad(es) como pendientes", key=f"group_undo_{clean_text(first_record.get('fme_cedula')) or clean_text(fme_name)}", use_container_width=True):
                    for group_record, _, _, _ in fme_cards:
                        set_message_sent(group_record["assignment_id"], group_record.get("feedback_qty", 0), False)
                    st.rerun()
            st.divider()
            for r, fme, sent, msg_state in []:
                with st.container(border=True):
                    st.markdown(f"**{r['assignment_id']}** · Site: `{r.get('site_id','')}` · Feedback Qty: **{r.get('feedback_qty',0)}** · Estado: **{r.get('current_result','')}**")
                    st.caption(f"Observación vigente ({r.get('remark_source','')}): {r.get('current_remark','') or 'Sin observación'}")
                    msg = make_fme_message(r)
                    st.text_area("Mensaje", msg, height=150, key=f"msg_{r['assignment_id']}_{r.get('feedback_qty',0)}")
                    activity_phone = (fme or {}).get("celular", "")
                    a,b = st.columns(2)
                    if activity_phone:
                        a.link_button("📲 Abrir WhatsApp", f"https://wa.me/{activity_phone}?text={quote(msg)}")
                    else:
                        a.warning("Celular no encontrado en la BD")
                    if not sent:
                        if b.button("✅ Marcar como enviado", key=f"sent_{r['assignment_id']}_{r.get('feedback_qty',0)}"):
                            set_message_sent(r["assignment_id"], r.get("feedback_qty",0), True, msg)
                            st.rerun()
                    else:
                        b.success(f"Enviado {msg_state.get('sent_at','')[:16] if msg_state else ''}")
                        if b.button("↩️ Marcar pendiente", key=f"undo_{r['assignment_id']}_{r.get('feedback_qty',0)}"):
                            set_message_sent(r["assignment_id"], r.get("feedback_qty",0), False, msg)
                            st.rerun()


def render_messages_v2():
    """One consolidated WhatsApp management card per FME."""
    _section("Mensajes FME", "Gestiona un único mensaje consolidado por técnico.")
    df = activities_df()
    if df.empty:
        st.info("No hay actividades para gestionar.")
        return
    df = df[_is_rejection(df["current_result"])].copy()
    if df.empty:
        st.success("No hay rechazos activos para notificar.")
        return
    left, right = st.columns(2)
    selected = left.multiselect("Coordinador", coordinator_filter_options(df["coordinator"].dropna().unique().tolist()), key="msg_coord_filter")
    state = right.selectbox("Estado de envío", ["Todos", "Pendientes", "Enviados"], key="msg_state_filter")
    if selected:
        df = df[df["coordinator"].isin(selected)]

    groups = {}
    for _, row in df.iterrows():
        status = message_status(row["assignment_id"], row.get("feedback_qty", 0))
        sent = bool(status and status.get("sent"))
        if state == "Pendientes" and sent:
            continue
        if state == "Enviados" and not sent:
            continue
        fme = fme_for_cedula(clean_text(row.get("fme_cedula")))
        group_key = clean_text(row.get("fme_cedula")) or clean_text(row.get("fme_nombre")) or "FME sin identificar"
        groups.setdefault(group_key, []).append((row, fme, sent))

    all_cards = [card for group in groups.values() for card in group]
    c1, c2, c3 = st.columns(3)
    c1.metric("FME visibles", len(groups))
    c2.metric("Actividades pendientes", sum(not card[2] for card in all_cards))
    c3.metric("Actividades enviadas", sum(card[2] for card in all_cards))
    if not groups:
        st.info("No hay mensajes que coincidan con los filtros seleccionados.")
        return

    for group_key, cards in sorted(groups.items(), key=lambda item: clean_text(item[1][0][0].get("fme_nombre")).casefold()):
        first_row, first_fme, _ = cards[0]
        name = clean_text(first_row.get("fme_nombre")) or "FME sin identificar"
        phone = clean_text((first_fme or {}).get("celular"))
        pending_cards = [card for card in cards if not card[2]]
        records = [card[0] for card in (pending_cards or cards)]
        message = make_fme_group_message(records, name)
        with st.expander(f"{name} · {len(cards)} actividad(es) · {len(pending_cards)} pendiente(s)", expanded=False):
            st.caption(f"Coordinador: {first_row.get('coordinator','')} · Celular: {phone or 'Sin celular registrado'}")
            st.text_area("Mensaje consolidado", message, height=300, key=f"fme_group_message_{group_key}")
            if pending_cards:
                if phone:
                    st.link_button("📲 Abrir WhatsApp con el mensaje consolidado", f"https://wa.me/{phone}?text={quote(message)}", use_container_width=True)
                else:
                    st.warning("Este FME no tiene un celular registrado.")
                if st.button(f"✅ Marcar {len(pending_cards)} actividad(es) como enviadas", key=f"fme_group_sent_{group_key}", type="primary", use_container_width=True):
                    for record, _, _ in pending_cards:
                        set_message_sent(record["assignment_id"], record.get("feedback_qty", 0), True, message)
                    st.rerun()
            else:
                st.success("Todas las actividades de este FME están marcadas como enviadas.")
                if st.button(f"↩️ Marcar {len(cards)} actividad(es) como pendientes", key=f"fme_group_pending_{group_key}", use_container_width=True):
                    for record, _, _ in cards:
                        set_message_sent(record["assignment_id"], record.get("feedback_qty", 0), False)
                    st.rerun()


def render_feedback():
    st.header("📝 Feedback OWS")
    acts = activities_df()
    if acts.empty: return
    labels = {f"{r['assignment_id']} | {r.get('site_id','')} | {r.get('fme_nombre','')}": r for _,r in acts.iterrows()}
    label = st.selectbox("Actividad", list(labels.keys()))
    r = labels[label]
    st.info(f"Ciclo actual detectado: Feedback Qty = {r.get('feedback_qty',0)} · {r.get('current_result','')}")
    with st.form("feedback_form"):
        fecha = st.date_input("Fecha de feedback", value=date.today())
        estado = st.selectbox("Estado", ["Pendiente","Aprobado","Rechazado","Otro"])
        feedback = st.text_area("Feedback (comentario que dejaste en OWS)", height=160)
        ok = st.form_submit_button("Guardar feedback", type="primary")
    if ok:
        if not feedback.strip():
            st.error("El comentario de feedback es obligatorio.")
        else:
            payload = {
                "assignment_id": r["assignment_id"], "site_id": r.get("site_id",""),
                "fme_id": r.get("fme_nombre","") or r.get("fme_id_raw",""),
                "coordinador": r.get("coordinator",""), "fecha_feedback": fecha.isoformat(),
                "estado": estado, "feedback": feedback.strip(),
            }
            save_ows_feedback(payload)
            st.success("Feedback guardado en Supabase.")
    hist = feedback_df()
    if not hist.empty:
        st.subheader("Historial")
        st.dataframe(hist[[c for c in ["assignment_id","site_id","fme_id","coordinador","fecha_feedback","estado","feedback"] if c in hist.columns]], use_container_width=True, hide_index=True)
        records = feedback_json_records(hist.rename(columns={}))
        st.download_button("⬇️ Exportar JSON", json.dumps(records, ensure_ascii=False, indent=2), "feedback_ows.json", "application/json")


def closed_email_body(r, closed_at):
    audit = clean_text(r.get("audit_time"))
    remark = clean_text(r.get("current_remark"))
    if any(k in remark.lower() for k in ["imagen", "foto", "fotogr", "evidencia"]):
        purpose = "Las evidencias fotográficas solicitadas pueden ser remitidas por correo. Para ajustes que requieran modificar formularios o información dentro de OWS, la actividad cerrada no permite realizar cambios adicionales."
    else:
        purpose = "Las correcciones solicitadas requieren modificar información o formularios de la actividad. Al encontrarse en estado Closed, no es posible realizar dichos cambios en OWS, por lo que se solicita gestionar la exclusión correspondiente."
    return f"""Buenos días, cordial saludo.\n\nDe acuerdo con lo mencionado con anterioridad con respecto al cierre de actividades, esta fue cerrada automáticamente por el sistema el día {closed_at}. Posteriormente, fue auditada por parte de GNOC el día {audit}, quedando en estado rechazada.\n\nActividad: {r.get('assignment_id','')}\nEstado: Closed\nSitio: {r.get('site_id','')}\n\nAudit/Feedback Remark:\n{remark}\n\n{purpose}\n\nAgradecemos gestionar lo correspondiente con el fin de evitar un impacto negativo en la medición de los KPI."""


def render_closed():
    st.header("🔒 Closed / Exclusiones")
    df = activities_df()
    if df.empty: return
    df = df[(df["task_status"].astype(str).str.lower()=="closed") & df["current_result"].astype(str).str.contains("RECHAZ", na=False)]
    if df.empty:
        st.success("No hay actividades Closed rechazadas pendientes en la vista actual.")
        return
    labels={f"{r['assignment_id']} | {r.get('site_id','')}":r for _,r in df.iterrows()}
    label=st.selectbox("Actividad Closed", list(labels.keys()))
    r=labels[label]
    closed_date=st.date_input("Fecha de cierre consultada en OWS", value=date.today())
    closed_time=st.time_input("Hora de cierre")
    closed_at=datetime.combine(closed_date,closed_time).isoformat(sep=" ")
    subject=f"EXCLUSIÓN PARA {r['assignment_id']} EN ESTADO CLOSED"
    body=closed_email_body(r,closed_at)
    st.text_input("Para", value="; ".join(TO_EMAILS), disabled=True)
    st.text_input("CC", value="; ".join(CC_EMAILS), disabled=True)
    st.text_input("Asunto", value=subject, disabled=True)
    st.text_area("Correo", value=body, height=420)
    if st.button("💾 Guardar gestión Closed"):
        set_closed_management({"assignment_id":r["assignment_id"],"closed_at":closed_at,"email_status":"LISTO","email_subject":subject,"email_body":body})
        st.success("Gestión Closed guardada en Supabase.")


def render_view():
    st.header("🔎 Vista coordinadores")
    df=activities_df()
    if df.empty:return
    coords=coordinator_filter_options(df["coordinator"].dropna().unique().tolist())
    sel=st.multiselect("Coordinador",coords)
    if sel: df=df[df["coordinator"].isin(sel)]
    result=st.multiselect("Resultado", sorted(df["current_result"].dropna().unique().tolist()))
    if result: df=df[df["current_result"].isin(result)]
    fqty=st.multiselect("Feedback Qty", sorted(df["feedback_qty"].dropna().unique().tolist()))
    if fqty: df=df[df["feedback_qty"].isin(fqty)]
    site=st.text_input("Filtrar Site ID contiene")
    if site: df=df[df["site_id"].astype(str).str.contains(site,case=False,na=False)]
    st.subheader("Resumen")
    summary=df.groupby("coordinator",dropna=False).size().reset_index(name="Cantidad").sort_values("Cantidad",ascending=False)
    st.dataframe(summary,use_container_width=True,hide_index=True)
    st.subheader("Detalle")
    st.dataframe(df[[c for c in ["assignment_id","task_id","site_id","coordinator","fme_nombre","feedback_qty","current_result","current_remark","task_status"] if c in df.columns]],use_container_width=True,hide_index=True)


def render_admin():
    st.header("⚙️ Migración / actualización de maestros")
    st.info("El uso diario solo requiere el WO Result Query. Esta sección se usa cuando cambie DATA COORDINADORES o CUADRILLAS.")

    st.subheader("1. DATA COORDINADORES - NUEVA.xlsx")
    st.caption("Prioridad exacta: Libro3[CODIGO] → Libro3[CodigoNetco] → Hoja1[SITIO]. Los coordinadores #N/A se ignoran.")
    coord=st.file_uploader("DATA COORDINADORES - NUEVA.xlsx",type=["xlsx"],key="coord")
    if coord:
        try:
            coord.seek(0)
            lookup_df=coordinator_map_from_excel(coord)
            coord.seek(0)
            master_df=coordinators_master_from_excel(coord)
            metric_codes, metric_master = st.columns(2)
            metric_codes.metric("Códigos únicos detectados", f"{len(lookup_df):,}")
            metric_master.metric("Registros maestros", f"{len(master_df):,}")
            if not lookup_df.empty and st.button("Guardar/actualizar coordinadores en Supabase", type="primary", key="save_coordinators"):
                try:
                    # Replace derived lookup so sites removed from the new master do not remain active.
                    nmap=replace_many("site_coordinator_map",lookup_df.to_dict("records"),"site_code",on_conflict="site_code")
                    nmaster=0
                    if not master_df.empty:
                        # This table is the detailed V12 master used by buscar_coordinador().
                        nmaster=replace_many("coordinadores",master_df.to_dict("records"),"coordinador")
                    st.success(f"Actualización completa: {nmap} códigos de búsqueda y {nmaster} registros maestros guardados.")
                except Exception as e:
                    st.error(f"No se pudo migrar. Ejecuta primero schema_v12.sql en Supabase. Detalle: {e}")
            with st.expander("Vista previa de códigos a actualizar", expanded=True):
                st.caption("Se muestran los primeros 50 registros. La prioridad 1 tiene preferencia sobre 2 y 3.")
                st.dataframe(lookup_df.head(50),use_container_width=True,hide_index=True,height=360)
        except Exception as e:
            st.error(f"No se pudo leer DATA COORDINADORES: {e}")
            lookup_df=pd.DataFrame(); master_df=pd.DataFrame()

    st.divider()
    st.subheader("2. CUADRILLAS ACT. NUEVO PROY. 2026.xlsx")
    fme=st.file_uploader("CUADRILLAS ACT. NUEVO PROY. 2026.xlsx",type=["xlsx"],key="fme")
    if fme and st.button("Actualizar FME / celulares"):
        try:
            df=fmes_from_excel(fme)
            n=upsert_many("fmes",df.to_dict("records"),on_conflict="cedula")
            st.success(f"{n} FME actualizados.")
        except Exception as e:
            st.error(f"No se pudo migrar CUADRILLAS: {e}")

    st.divider()
    st.caption("Después de actualizar estos maestros, vuelve a Importar WO para que las actividades nuevas tomen el coordinador vigente.")


def _is_rejection(series):
    return series.astype(str).str.contains("RECHAZ", case=False, na=False)


def _fme_visible_rejections(df):
    """Rechazos que debe ver el técnico, incluidos los de auditoría inicial.

    Algunas importaciones antiguas conservaron el texto de rechazo en los campos
    originales de auditoría aunque `current_result` no quedó clasificado como
    RECHAZADO. En la vista FME esos casos de feedback 0 también son pendientes
    que el técnico debe corregir.
    """
    if df.empty:
        return pd.Series(False, index=df.index, dtype=bool)
    by_current_result = _is_rejection(df.get("current_result", pd.Series("", index=df.index)))
    feedback_qty = pd.to_numeric(df.get("feedback_qty", pd.Series(0, index=df.index)), errors="coerce").fillna(0)
    audit_text = (
        df.get("audit_status", pd.Series("", index=df.index)).fillna("").astype(str)
        + " "
        + df.get("audit_remark_original", pd.Series("", index=df.index)).fillna("").astype(str)
    )
    initial_audit_rejection = feedback_qty.eq(0) & audit_text.str.contains("RECHAZ|REJECT", case=False, na=False)
    return by_current_result | initial_audit_rejection


def _section(title, detail=""):
    st.header(title)
    if detail:
        st.caption(detail)


def _xlsx_bytes(df, sheet_name):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return buffer.getvalue()


def render_dashboard_v2():
    _section("Dashboard", "Resumen ejecutivo de gestión, comunicación y ciclos de feedback.")
    df = activities_df()
    if df.empty:
        st.info("Aún no hay actividades. Usa Importar WO para iniciar la operación.")
        return
    coordinator_options = ["Todos los coordinadores"] + coordinator_filter_options(df["coordinator"].dropna().unique())
    current_coordinator = st.session_state.get("dashboard_selected_coordinator", "Todos los coordinadores")
    if current_coordinator not in coordinator_options:
        current_coordinator = "Todos los coordinadores"
    selected_coordinator = st.selectbox(
        "Coordinador para el Dashboard",
        coordinator_options,
        index=coordinator_options.index(current_coordinator),
        key="dashboard_coordinator_widget",
    )
    st.session_state["dashboard_selected_coordinator"] = selected_coordinator
    if selected_coordinator != "Todos los coordinadores":
        df = df[df["coordinator"].eq(selected_coordinator)].copy()
    st.caption(f"Vista actual: **{selected_coordinator}**")
    messages = pd.DataFrame(select_all("fme_messages"))
    closed = pd.DataFrame(select_all("closed_management"))
    assignment_ids = set(df["assignment_id"].astype(str))
    if not messages.empty:
        messages = messages[messages["assignment_id"].astype(str).isin(assignment_ids)]
    if not closed.empty:
        closed = closed[closed["assignment_id"].astype(str).isin(assignment_ids)]
    rejected = _is_rejection(df["current_result"])
    sent = int(messages.get("sent", pd.Series(dtype=bool)).fillna(False).sum()) if not messages.empty else 0
    active = int(rejected.sum())
    approved = int(df["current_result"].astype(str).str.upper().eq("APROBADO").sum())
    again = int(df["current_result"].astype(str).str.upper().eq("RECHAZADO_NUEVAMENTE").sum())
    closed_count = int(df["task_status"].astype(str).str.lower().eq("closed").sum())
    feedback_zero = int(df["feedback_qty"].fillna(0).astype(int).eq(0).sum())
    cards = [
        ("Rechazos activos", active), ("Pendientes de notificar", max(active - sent, 0)),
        ("WhatsApp enviados", sent), ("Pendientes de feedback", feedback_zero),
        ("Aprobados", approved), ("Rechazados nuevamente", again),
        ("Actividades Closed", closed_count),
        ("Exclusiones pendientes", int(closed.get("email_status", pd.Series(dtype=str)).astype(str).str.upper().eq("PENDIENTE").sum()) if not closed.empty else 0),
    ]
    for row in (cards[:4], cards[4:]):
        for column, (label, value) in zip(st.columns(4), row):
            column.metric(label, value)
    left, right = st.columns(2)
    with left:
        st.subheader("Rechazos por coordinador")
        st.caption("Toca una barra para actualizar el Dashboard y los FME de ese coordinador.")
        coordinator_chart = (
            df[rejected]
            .groupby("coordinator", dropna=False)
            .size()
            .sort_values(ascending=False)
            .head(12)
            .reset_index(name="rechazos")
        )
        coordinator_chart["coordinator"] = coordinator_chart["coordinator"].fillna("Sin coordinador")
        clicked_coordinator = render_selectable_coordinator_chart(coordinator_chart)
        if clicked_coordinator and clicked_coordinator != st.session_state.get("dashboard_selected_coordinator"):
            st.session_state["dashboard_selected_coordinator"] = clicked_coordinator
            st.rerun()
    with right:
        st.subheader("Estado actual de las actividades")
        status_chart = df["current_result"].fillna("SIN_CLASIFICAR").map(result_label).value_counts()
        status_chart = status_chart.rename_axis("Estado").reset_index(name="Actividades")
        render_horizontal_chart(status_chart, "Estado", "Actividades", "#e07a3f")
    left, right = st.columns(2)
    with left:
        st.subheader("Actividades por etapa de feedback")
        feedback_chart = df["feedback_qty"].fillna(0).astype(int).map(feedback_stage_label).value_counts()
        stage_order = ["Auditoría inicial", "Primer feedback", "Segundo feedback", "Tercer feedback"]
        feedback_chart = feedback_chart.reindex(stage_order + [label for label in feedback_chart.index if label not in stage_order]).dropna()
        feedback_chart = feedback_chart.rename_axis("Etapa").reset_index(name="Actividades")
        render_horizontal_chart(feedback_chart, "Etapa", "Actividades", "#12a6a6")
    with right:
        st.subheader("Rechazos por FME")
        fme_chart = df[rejected].copy()
        fme_chart["fme_corto"] = fme_chart["fme_nombre"].map(short_fme_label).replace("", "FME sin identificar")
        fme_chart = fme_chart.groupby("fme_corto", dropna=False).size().sort_values(ascending=False).head(12).rename_axis("FME").reset_index(name="Rechazos")
        render_horizontal_chart(fme_chart, "FME", "Rechazos", "#6856a8")


def render_rechazos():
    _section("Rechazos", "Aplica filtros para priorizar la gestión operativa.")
    df = activities_df()
    if df.empty:
        st.info("No hay registros para consultar.")
        return
    with st.container(border=True):
        a, b, c, d = st.columns(4)
        coords = a.multiselect("Coordinador", coordinator_filter_options(df["coordinator"].dropna().astype(str).unique()), key="r_coord")
        fmes = b.multiselect("FME", sorted(df["fme_nombre"].dropna().astype(str).unique()), key="r_fme")
        states = c.multiselect("Estado", sorted(df["current_result"].dropna().astype(str).unique()), key="r_state")
        quantities = d.multiselect("Feedback Qty", sorted(df["feedback_qty"].dropna().unique()), key="r_qty")
        e, f, g, h = st.columns(4)
        site = e.text_input("Site ID", key="r_site")
        assignment = f.text_input("Assignment ID", key="r_assignment")
        lifecycle = g.selectbox("Actividad", ["Todos", "Closed", "Open"], key="r_lifecycle")
        cycle = h.selectbox("Ciclo", ["Todos", "Con feedback", "Sin feedback"], key="r_cycle")
    view = df.copy()
    if coords: view = view[view["coordinator"].isin(coords)]
    if fmes: view = view[view["fme_nombre"].isin(fmes)]
    if states: view = view[view["current_result"].isin(states)]
    if quantities: view = view[view["feedback_qty"].isin(quantities)]
    if site: view = view[view["site_id"].astype(str).str.contains(site, case=False, na=False)]
    if assignment: view = view[view["assignment_id"].astype(str).str.contains(assignment, case=False, na=False)]
    if lifecycle != "Todos": view = view[view["task_status"].astype(str).str.lower().eq(lifecycle.lower())]
    if cycle == "Con feedback": view = view[view["feedback_qty"].fillna(0).astype(int).gt(0)]
    if cycle == "Sin feedback": view = view[view["feedback_qty"].fillna(0).astype(int).eq(0)]
    st.caption(f"Mostrando {len(view):,} de {len(df):,} registros")
    cols = ["assignment_id", "site_id", "coordinator", "fme_nombre", "feedback_qty", "current_result", "remark_source", "current_remark", "task_status", "audit_time"]
    st.dataframe(view[[c for c in cols if c in view]], use_container_width=True, hide_index=True, height=520)


def render_historial():
    _section("Historial", "Línea de tiempo completa de la auditoría y sus ciclos de feedback.")
    acts = activities_df()
    if acts.empty:
        st.info("No hay actividades disponibles.")
        return
    labels = {f"{r['assignment_id']} | {r.get('site_id', '')}": r["assignment_id"] for _, r in acts.iterrows()}
    selected = st.selectbox("Actividad", list(labels.keys()))
    cycles = pd.DataFrame(select_all("audit_cycles"))
    rows = cycles[cycles["assignment_id"].eq(labels[selected])].sort_values("feedback_qty") if not cycles.empty else pd.DataFrame()
    if rows.empty:
        st.info("No se han registrado ciclos para esta actividad todavía.")
        return
    for _, row in rows.iterrows():
        cycle = int(row.get("feedback_qty", 0))
        title = "Audit inicial" if cycle == 0 else f"Feedback Qty {cycle}"
        status = str(row.get("result", ""))
        icon = "✅" if status.upper() == "APROBADO" else ("🔴" if "RECHAZ" in status.upper() else "🟠")
        with st.container(border=True):
            st.markdown(f"**{icon} {title}**  ·  {status}")
            st.caption(f"{row.get('audit_time', '')} · {row.get('remark_source', '')}")
            st.write(row.get("remark", "Sin observación"))


def render_portal_operativo():
    _section("Portal de Coordinadores y FME", "Consulta operativa de actividades, rechazos y ciclos de feedback.")
    df = activities_df()
    if df.empty:
        st.info("No hay actividades disponibles todavía.")
        return
    coordinator_tab, fme_tab = st.tabs(["Vista para coordinador", "Vista para FME"])

    with coordinator_tab:
        coordinators = coordinator_filter_options(df["coordinator"].dropna().unique())
        coordinator = st.selectbox("Selecciona tu nombre", coordinators, key="portal_coordinator")
        view = df[df["coordinator"].eq(coordinator)].copy()
        rejected = _is_rejection(view["current_result"])
        feedback_pending = view["current_result"].astype(str).str.upper().eq("PENDIENTE_REVISION_FEEDBACK")
        a, b, c, d = st.columns(4)
        a.metric("Actividades", len(view))
        b.metric("Rechazos activos", int(rejected.sum()))
        c.metric("Pendientes de feedback", int(feedback_pending.sum()))
        d.metric("Actividades Closed", int(view["task_status"].astype(str).str.lower().eq("closed").sum()))
        st.subheader("Resumen por FME")
        summary = (
            view.assign(Rechazo=rejected)
            .groupby("fme_nombre", dropna=False)
            .agg(Actividades=("assignment_id", "count"), Rechazos=("Rechazo", "sum"))
            .reset_index()
            .rename(columns={"fme_nombre": "FME"})
            .sort_values(["Rechazos", "Actividades"], ascending=False)
        )
        st.dataframe(summary, use_container_width=True, hide_index=True)
        st.subheader("Actividades a gestionar")
        columns = ["assignment_id", "site_id", "fme_nombre", "feedback_qty", "current_result", "current_remark", "task_status"]
        show = view[rejected][[column for column in columns if column in view]].copy()
        if "feedback_qty" in show:
            show["feedback_qty"] = show["feedback_qty"].map(feedback_stage_label)
        if "current_result" in show:
            show["current_result"] = show["current_result"].map(result_label)
        st.dataframe(show, use_container_width=True, hide_index=True, height=420)

    with fme_tab:
        fmes = sorted(df["fme_nombre"].dropna().astype(str).unique(), key=str.casefold)
        fme_name = st.selectbox("Selecciona el FME", fmes, key="portal_fme")
        view = df[df["fme_nombre"].eq(fme_name)].copy()
        # Incluye rechazos de auditoría inicial aunque provengan de una carga
        # anterior cuyo `current_result` no se haya clasificado correctamente.
        rejected = _fme_visible_rejections(view)
        a, b, c = st.columns(3)
        a.metric("Actividades asignadas", len(view))
        b.metric("Rechazos por corregir", int(rejected.sum()))
        c.metric("Aprobadas", int(view["current_result"].astype(str).str.upper().eq("APROBADO").sum()))
        st.info("Revisa cada observación, realiza la corrección solicitada y confirma la gestión con tu coordinador.")
        st.subheader("Mis rechazos por corregir")
        st.caption("Se incluyen los rechazos iniciales (sin feedback) y los rechazos posteriores a feedback.")
        rejected_view = view[rejected].sort_values(["feedback_qty", "audit_time"], ascending=[True, False])
        if rejected_view.empty:
            st.success("No tienes rechazos activos en este momento.")
        else:
            for _, row in rejected_view.iterrows():
                stage = feedback_stage_label(row.get("feedback_qty", 0))
                with st.expander(f"{row.get('site_id', '')} · {stage} · {result_label(row.get('current_result', ''))}", expanded=False):
                    st.caption(f"Actividad: {row.get('assignment_id','')} · Coordinador: {row.get('coordinator','')}")
                    st.markdown("**Observación vigente**")
                    st.write(row.get("current_remark", "Sin observación registrada"))


def render_exportar():
    _section("Exportar", "Descarga información operativa sin modificar los datos de Supabase.")
    df = activities_df()
    feedback = feedback_df()
    rejected = df[_is_rejection(df["current_result"])].copy() if not df.empty else df
    left, right = st.columns(2)
    with left:
        st.download_button("Exportar Excel de actividades", _xlsx_bytes(df, "Actividades"), "actividades_fme.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, disabled=df.empty)
        st.caption("Base completa de actividades importadas.")
        st.download_button("Exportar JSON de feedback", json.dumps(feedback_json_records(feedback), ensure_ascii=False, indent=2), "feedback_ows.json", "application/json", use_container_width=True, disabled=feedback.empty)
        st.caption("Registros listos para compartir o integrar en OWS.")
    with right:
        st.download_button("Exportar rechazos", _xlsx_bytes(rejected, "Rechazos"), "rechazos_fme.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, disabled=rejected.empty)
        st.caption("Actividades cuyo estado actual contiene rechazo.")
        messages = pd.DataFrame(select_all("fme_messages"))
        st.download_button("Exportar mensajes FME", _xlsx_bytes(messages, "Mensajes"), "mensajes_fme.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, disabled=messages.empty)
        st.caption("Trazabilidad de notificaciones enviadas y pendientes.")


if not login_gate():
    st.stop()
if not db_ok():
    st.stop()

st.sidebar.title("Gestión Rechazos · V12")
st.sidebar.caption("Centro de control operativo")
page=st.sidebar.radio("Módulo",["Dashboard","Portal Coordinadores / FME","Rechazos","Importar WO","Mensajes FME","Feedback OWS","Closed / Exclusiones","Vista coordinadores","Historial","Exportar","Administración / Migración"])
if page=="Dashboard": render_dashboard_v2()
elif page=="Portal Coordinadores / FME": render_portal_operativo()
elif page=="Rechazos": render_rechazos()
elif page=="Importar WO": import_wo()
elif page=="Mensajes FME": render_messages_v2()
elif page=="Feedback OWS": render_feedback()
elif page=="Closed / Exclusiones": render_closed()
elif page=="Vista coordinadores": render_view()
elif page=="Historial": render_historial()
elif page=="Exportar": render_exportar()
elif page=="Administración / Migración": render_admin()
