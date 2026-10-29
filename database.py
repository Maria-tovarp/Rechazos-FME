import os
from datetime import datetime, timezone
import streamlit as st
from supabase import create_client


def config_value(name):
    """Read Streamlit Secrets locally or environment variables on Render."""
    try:
        return st.secrets.get(name) or os.getenv(name)
    except Exception:
        return os.getenv(name)


def get_client():
    url = config_value("SUPABASE_URL")
    key = config_value("SUPABASE_KEY")
    if not url or not key:
        raise RuntimeError("Configura SUPABASE_URL y SUPABASE_KEY en los Secrets de Streamlit.")
    return create_client(url, key)


def chunked(seq, size=500):
    for i in range(0, len(seq), size):
        yield seq[i:i+size]


def upsert_many(table, rows, on_conflict=None):
    if not rows:
        return 0
    sb = get_client()
    count = 0
    for batch in chunked(rows, 400):
        q = sb.table(table).upsert(batch, on_conflict=on_conflict) if on_conflict else sb.table(table).upsert(batch)
        q.execute()
        count += len(batch)
    return count



def delete_all(table, key_column):
    """Delete all rows from a table using a server-side key filter."""
    sb = get_client()
    sb.table(table).delete().neq(key_column, "__never_match__").execute()


def replace_many(table, rows, key_column, on_conflict=None):
    delete_all(table, key_column)
    return upsert_many(table, rows, on_conflict=on_conflict)

def select_all(table, columns="*", order=None, desc=False):
    sb = get_client()
    q = sb.table(table).select(columns)
    if order:
        q = q.order(order, desc=desc)
    return q.execute().data or []


def coordinator_for_site(site_id):
    """Use the V12 SQL function first; fall back to the derived map table."""
    sb = get_client()
    site_id = (site_id or "").strip()
    if not site_id:
        return {"coordinator":"TLF","source":"Asignación TLF","priority":99}
    try:
        r = sb.rpc("buscar_coordinador", {"p_site_id": site_id}).execute().data or []
        if r:
            return {"coordinator": r[0].get("coordinador", "TLF"), "source": r[0].get("fuente", "Asignación TLF"), "priority": 1}
    except Exception:
        pass
    try:
        r = sb.table("site_coordinator_map").select("coordinator,source,priority").eq("site_code", site_id).limit(1).execute().data
        if r:
            return r[0]
    except Exception:
        pass
    return {"coordinator":"TLF","source":"Asignación TLF","priority":99}


def fme_for_cedula(cedula):
    if not cedula:
        return None
    sb = get_client()
    r = sb.table("fmes").select("*").eq("cedula", cedula).limit(1).execute().data
    return r[0] if r else None


def upsert_activity(row, coordinator):
    payload = dict(row)
    payload["audit_time"] = row.get("audit_time").isoformat() if getattr(row.get("audit_time"), "isoformat", None) else None
    payload["coordinator"] = coordinator.get("coordinator", "TLF")
    payload["coordinator_source"] = coordinator.get("source", "Asignación TLF")
    payload["updated_at"] = datetime.now(timezone.utc).isoformat()
    payload.pop("needs_fme_message", None)
    upsert_many("activities", [payload], on_conflict="assignment_id")


def insert_cycle(row, coordinator):
    payload = {
        "assignment_id": row["assignment_id"],
        "feedback_qty": int(row.get("feedback_qty", 0)),
        "result": row.get("current_result", ""),
        "remark": row.get("current_remark", ""),
        "remark_source": row.get("remark_source", ""),
        "audit_time": row.get("audit_time").isoformat() if getattr(row.get("audit_time"), "isoformat", None) else None,
        "coordinator": coordinator.get("coordinator", "TLF"),
    }
    upsert_many("audit_cycles", [payload], on_conflict="assignment_id,feedback_qty")


def message_status(assignment_id, feedback_qty):
    sb = get_client()
    r = sb.table("fme_messages").select("*").eq("assignment_id", assignment_id).eq("feedback_qty", int(feedback_qty)).limit(1).execute().data
    return r[0] if r else None


def set_message_sent(assignment_id, feedback_qty, sent, message_text=""):
    payload = {
        "assignment_id": assignment_id,
        "feedback_qty": int(feedback_qty),
        "sent": bool(sent),
        "sent_at": datetime.now(timezone.utc).isoformat() if sent else None,
        "message_text": message_text,
    }
    upsert_many("fme_messages", [payload], on_conflict="assignment_id,feedback_qty")


def save_ows_feedback(payload):
    sb = get_client()
    sb.table("ows_feedbacks").insert(payload).execute()


def set_closed_management(payload):
    upsert_many("closed_management", [payload], on_conflict="assignment_id")
