import re
import io
import json
import unicodedata
from datetime import datetime
import pandas as pd

NA_VALUES = {"N/A", "NA", "N.A", "N.A.", "APROBADO", "APPROVE", "APPROVED"}


def clean_text(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def normalize(value):
    txt = clean_text(value).upper()
    txt = unicodedata.normalize("NFKD", txt)
    return "".join(c for c in txt if not unicodedata.combining(c)).strip()


COORDINATOR_ALIASES = {
    "BORIS DEL CRISTO": "Boris Del Cristo",
    "ISABEL HERNANDEZ": "Isabel Hernandez",
    "FIJA": "Fija",
}


def canonical_coordinator(value):
    """Keep approved coordinator names consistent across Excel imports."""
    text = " ".join(clean_text(value).split())
    return COORDINATOR_ALIASES.get(normalize(text), text)


def is_na_approval(value):
    txt = normalize(value).replace(" ", "")
    return txt in {x.replace(" ", "") for x in NA_VALUES}


def parse_int(value, default=0):
    if value is None or (isinstance(value, float) and pd.isna(value)) or clean_text(value) == "":
        return default
    try:
        return int(float(value))
    except Exception:
        return default


def extract_cedula(fme_id):
    text = clean_text(fme_id)
    m = re.search(r"\((\d{5,15})\)\s*$", text)
    if m:
        return m.group(1)
    nums = re.findall(r"\d{5,15}", text)
    return nums[-1] if nums else ""


def extract_fme_name(fme_id):
    text = clean_text(fme_id)
    return re.sub(r"\s*\(\d{5,15}\)\s*$", "", text).strip()


def normalize_phone(value):
    digits = re.sub(r"\D", "", clean_text(value))
    if digits.endswith(".0"):
        digits = digits[:-2]
    if len(digits) == 10 and digits.startswith("3"):
        return "57" + digits
    if len(digits) == 12 and digits.startswith("57"):
        return digits
    return digits


def classify_wo_row(row):
    """Classifies current audit/feedback lifecycle.

    cycle 0 = initial audit.
    Feedback Qty 1 = first feedback review, 2 = second, etc.
    If Feedback Remark is N/A in a feedback cycle -> approved.
    If Feedback Remark has text -> rejected again and that text is current remark.
    """
    qty = parse_int(row.get("Feedback Qty"), 0)
    audit_status = normalize(row.get("Audit status"))
    audit_remark = clean_text(row.get("Audit remark"))
    feedback_remark = clean_text(row.get("Feedback Remark"))

    if qty >= 1:
        if is_na_approval(feedback_remark) or is_na_approval(audit_status):
            return {
                "feedback_qty": qty,
                "result": "APROBADO",
                "current_remark": "N/A",
                "remark_source": "Feedback Remark",
                "needs_fme_message": False,
            }
        if feedback_remark:
            return {
                "feedback_qty": qty,
                "result": "RECHAZADO_NUEVAMENTE",
                "current_remark": feedback_remark,
                "remark_source": "Feedback Remark",
                "needs_fme_message": True,
            }
        return {
            "feedback_qty": qty,
            "result": "PENDIENTE_APROBACION",
            "current_remark": "",
            "remark_source": "Feedback Remark",
            "needs_fme_message": False,
        }

    if is_na_approval(audit_status) or is_na_approval(audit_remark):
        return {
            "feedback_qty": 0,
            "result": "APROBADO",
            "current_remark": "N/A",
            "remark_source": "Audit remark",
            "needs_fme_message": False,
        }

    if "REJECT" in audit_status or audit_remark:
        return {
            "feedback_qty": 0,
            "result": "RECHAZADO",
            "current_remark": audit_remark,
            "remark_source": "Audit remark",
            "needs_fme_message": True,
        }

    return {
        "feedback_qty": 0,
        "result": audit_status or "SIN_CLASIFICAR",
        "current_remark": audit_remark,
        "remark_source": "Audit remark",
        "needs_fme_message": False,
    }


def prepare_wo_dataframe(file_or_path):
    df = pd.read_excel(file_or_path)
    required = ["Assignment ID", "Site ID", "FME ID", "Audit remark", "Audit status", "Feedback Remark", "Feedback Qty"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {', '.join(missing)}")
    rows = []
    for _, r in df.iterrows():
        c = classify_wo_row(r)
        rows.append({
            "assignment_id": clean_text(r.get("Assignment ID")),
            "task_id": clean_text(r.get("Task ID")),
            "site_id": clean_text(r.get("Site ID")),
            "task_status": clean_text(r.get("Task status")).lower(),
            "fme_id_raw": clean_text(r.get("FME ID")),
            "fme_cedula": extract_cedula(r.get("FME ID")),
            "fme_nombre": extract_fme_name(r.get("FME ID")),
            "audit_time": pd.to_datetime(r.get("Audit time"), errors="coerce"),
            "audit_status": clean_text(r.get("Audit status")),
            "audit_remark_original": clean_text(r.get("Audit remark")),
            "feedback_remark_latest": clean_text(r.get("Feedback Remark")),
            "feedback_qty": c["feedback_qty"],
            "current_result": c["result"],
            "current_remark": c["current_remark"],
            "remark_source": c["remark_source"],
            "needs_fme_message": c["needs_fme_message"],
            "region": clean_text(r.get("Region")),
            "fm_office": clean_text(r.get("FM Office")),
            "audit_type": clean_text(r.get("Audit type")),
            "task_title": clean_text(r.get("Task title")),
        })
    out = pd.DataFrame(rows)
    out = out[out["assignment_id"] != ""].copy()
    return out


def mark_advanced_feedbacks_pending(df, previous_activities):
    """Marks a feedback cycle that advanced since the last WO as awaiting approval.

    A higher `Feedback Qty` with no new feedback remark means the FME submitted
    the correction and the activity awaits GNOC approval. A feedback remark
    that explicitly reports a rejection must always remain an active rejection,
    regardless of the cycle number.
    """
    if df.empty or previous_activities is None or previous_activities.empty:
        return df, 0
    previous_qty = {
        clean_text(row.get("assignment_id")): parse_int(row.get("feedback_qty"), 0)
        for _, row in previous_activities.iterrows()
    }
    output = df.copy()
    advanced = 0
    for index, row in output.iterrows():
        assignment_id = clean_text(row.get("assignment_id"))
        old_qty = previous_qty.get(assignment_id)
        new_qty = parse_int(row.get("feedback_qty"), 0)
        feedback_remark = clean_text(row.get("feedback_remark_latest"))
        if old_qty is not None and new_qty > old_qty and not feedback_remark:
            output.at[index, "current_result"] = "PENDIENTE_APROBACION"
            output.at[index, "current_remark"] = "Feedback registrado. Pendiente de validación y aprobación."
            output.at[index, "remark_source"] = "Seguimiento de feedback"
            output.at[index, "needs_fme_message"] = False
            advanced += 1
    return output, advanced


def _valid_coordinator(value):
    txt = clean_text(value)
    return bool(txt) and normalize(txt) not in {"#N/A", "N/A", "NA", "NAN", "-", "0", "0.0"}


def coordinators_master_from_excel(file_or_path):
    """Read DATA COORDINADORES - NUEVA.xlsx.

    Principal sheet: Libro3
      Coordinador, CODIGO, CodigoNetco and site metadata.
    Fallback sheet: Hoja1
      SITIO, COORDINADOR.

    Returns rows ready for the `coordinadores` table.
    """
    xls = pd.ExcelFile(file_or_path)
    if "Libro3" not in xls.sheet_names:
        raise ValueError("El archivo debe contener la hoja 'Libro3'.")

    main = pd.read_excel(file_or_path, sheet_name="Libro3")
    required = ["Coordinador", "CODIGO", "CodigoNetco"]
    missing = [c for c in required if c not in main.columns]
    if missing:
        raise ValueError(f"En Libro3 faltan columnas: {', '.join(missing)}")

    rows = []
    for _, r in main.iterrows():
        coord = canonical_coordinator(r.get("Coordinador"))
        if not _valid_coordinator(coord):
            continue
        codigo = clean_text(r.get("CODIGO"))
        netco = clean_text(r.get("CodigoNetco"))
        if not codigo and not netco:
            continue
        rows.append({
            "codigo": codigo or None,
            "codigo_netco": netco or None,
            "sitio_alt": None,
            "coordinador": coord,
            "zona": clean_text(r.get("Zona")) or None,
            "departamento": clean_text(r.get("DEPARTAMENTO")) or None,
            "municipio": clean_text(r.get("MUNICIPIO")) or None,
            "nombre_sitio": clean_text(r.get("NOMBRE")) or None,
            "fuente": "Libro3",
            "activo": True,
        })

    if "Hoja1" in xls.sheet_names:
        h1 = pd.read_excel(file_or_path, sheet_name="Hoja1")
        if "SITIO" in h1.columns and "COORDINADOR" in h1.columns:
            for _, r in h1.iterrows():
                site = clean_text(r.get("SITIO"))
                coord = canonical_coordinator(r.get("COORDINADOR"))
                if site and _valid_coordinator(coord):
                    rows.append({
                        "codigo": None,
                        "codigo_netco": None,
                        "sitio_alt": site,
                        "coordinador": coord,
                        "zona": None,
                        "departamento": None,
                        "municipio": None,
                        "nombre_sitio": None,
                        "fuente": "Hoja1",
                        "activo": True,
                    })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    dedup_cols = ["codigo", "codigo_netco", "sitio_alt", "coordinador"]
    return df.drop_duplicates(subset=dedup_cols, keep="first").reset_index(drop=True)


def coordinator_map_from_excel(file_or_path):
    """Derived lookup table with the exact BUSCARX priority.

    1) Libro3[CODIGO]
    2) Libro3[CodigoNetco]
    3) Hoja1[SITIO]
    """
    xls = pd.ExcelFile(file_or_path)
    if "Libro3" not in xls.sheet_names:
        raise ValueError("El archivo debe contener la hoja 'Libro3'.")
    main = pd.read_excel(file_or_path, sheet_name="Libro3")
    required = ["Coordinador", "CODIGO", "CodigoNetco"]
    missing = [c for c in required if c not in main.columns]
    if missing:
        raise ValueError(f"En Libro3 faltan columnas: {', '.join(missing)}")

    frames = []
    for _, r in main.iterrows():
        coord = canonical_coordinator(r.get("Coordinador"))
        if not _valid_coordinator(coord):
            continue
        codigo = clean_text(r.get("CODIGO"))
        netco = clean_text(r.get("CodigoNetco"))
        if codigo and normalize(codigo) not in {"-", "NAN"}:
            frames.append({"site_code": codigo, "coordinator": coord, "source": "Libro3-CODIGO", "priority": 1})
        if netco and normalize(netco) not in {"-", "NAN"}:
            frames.append({"site_code": netco, "coordinator": coord, "source": "Libro3-CodigoNetco", "priority": 2})

    if "Hoja1" in xls.sheet_names:
        h1 = pd.read_excel(file_or_path, sheet_name="Hoja1")
        if "SITIO" in h1.columns and "COORDINADOR" in h1.columns:
            for _, r in h1.iterrows():
                site = clean_text(r.get("SITIO"))
                coord = canonical_coordinator(r.get("COORDINADOR"))
                if site and _valid_coordinator(coord):
                    frames.append({"site_code": site, "coordinator": coord, "source": "Hoja1-SITIO", "priority": 3})

    df = pd.DataFrame(frames)
    if df.empty:
        return df
    df["site_code_norm"] = df["site_code"].map(normalize)
    df = df.sort_values(["site_code_norm", "priority"]).drop_duplicates("site_code_norm", keep="first")
    return df.drop(columns=["site_code_norm"])

def fmes_from_excel(file_or_path):
    xls = pd.ExcelFile(file_or_path)
    rows = []
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(file_or_path, sheet_name=sheet)
        except Exception:
            continue
        cols = {normalize(c): c for c in df.columns}
        ced_col = cols.get("CEDULA")
        nom_col = cols.get("NOMBRE") or cols.get("NOMBRE ")
        cel_col = cols.get("CELULAR")
        if not ced_col or not nom_col or not cel_col:
            continue
        for _, r in df.iterrows():
            ced = re.sub(r"\D", "", clean_text(r.get(ced_col)).replace(".0", ""))
            if not ced:
                continue
            rows.append({
                "cedula": ced,
                "nombre": clean_text(r.get(nom_col)),
                "celular": normalize_phone(r.get(cel_col)),
                "correo": clean_text(r.get(cols.get("CORREOS", ""))) if cols.get("CORREOS") else "",
                "grupo": clean_text(r.get(cols.get("GRUPO", ""))) if cols.get("GRUPO") else "",
                "zona": clean_text(r.get(cols.get("ZONA", ""))) if cols.get("ZONA") else "",
                "source_sheet": sheet,
            })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    # Prefer rows with a phone.
    df["has_phone"] = df["celular"].astype(str).str.len().gt(0)
    df = df.sort_values(["cedula", "has_phone"], ascending=[True, False]).drop_duplicates("cedula", keep="first")
    return df.drop(columns=["has_phone"])


def feedback_json_records(df):
    cols = ["assignment_id", "site_id", "fme_id", "coordinador", "fecha_feedback", "estado", "feedback"]
    out = []
    for _, r in df.iterrows():
        out.append({c: ("" if pd.isna(r.get(c)) else str(r.get(c))) for c in cols})
    return out


def make_fme_message(row):
    cycle = parse_int(row.get("feedback_qty"), 0)
    if cycle == 0:
        intro = "La actividad fue rechazada en auditoría GNOC."
    else:
        intro = f"La actividad fue rechazada nuevamente después del feedback #{cycle}."
    return (
        f"Buen día. {intro}\n\n"
        f"Sitio: {clean_text(row.get('site_id'))}\n"
        f"Actividad: {clean_text(row.get('assignment_id'))}\n"
        f"Observación: {clean_text(row.get('current_remark'))}\n\n"
        "Por favor validar y realizar la corrección correspondiente."
    )


def make_fme_group_message(rows, fme_name):
    """Build one clear WhatsApp message for all pending activities of an FME."""
    greeting = clean_text(fme_name) or "equipo"
    lines = [
        f"Buen día, {greeting}.",
        "",
        "Te comparto los rechazos de auditoría pendientes de corrección:",
        "",
    ]
    for number, row in enumerate(rows, start=1):
        task = clean_text(row.get("task_id")) or clean_text(row.get("assignment_id"))
        lines.extend([
            f"{number}. Sitio: {clean_text(row.get('site_id'))}",
            f"Task ID: {task}",
            f"Audit Remark: {clean_text(row.get('current_remark')) or 'Sin observación registrada'}",
            "",
        ])
    lines.extend([
        "Por favor realizar las correcciones indicadas y confirmar cuando queden gestionadas.",
        "Gracias.",
    ])
    return "\n".join(lines)
