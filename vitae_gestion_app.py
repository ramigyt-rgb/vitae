# vitae_gestion_app.py
# Ejecutar en VS Code / terminal:
#   pip install streamlit pandas plotly openpyxl
#   streamlit run vitae_gestion_app.py

from __future__ import annotations

import sqlite3
import json
import os
from pathlib import Path
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st

# =========================================================
# CONFIG GENERAL
# =========================================================

APP_TITLE = "Sistema de Gestión"
DB_PATH = Path("vitae_gestion.db")
DATE_FMT = "%Y-%m-%d"
TECH_COLUMNS = ["id", "created_at", "updated_at"]

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-title {
        font-size: 2.1rem;
        font-weight: 800;
        margin-bottom: 0.1rem;
    }
    .subtitle {
        color: #6b7280;
        margin-bottom: 1.2rem;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.35rem !important;
    }
    .small-muted {
        color: #6b7280;
        font-size: 0.88rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# DEFINICIÓN DE MÓDULOS
# =========================================================

MODULES: Dict[str, Dict[str, Any]] = {
    "Caja VMR": {
        "table": "caja_vmr",
        "empresa": "VMR",
        "tipo": "flujo",
        "descripcion": "Movimientos de efectivo de Vitae Medicina Reproductiva.",
        "fields": [
            ("fecha", "date", True),
            ("concepto", "text", True),
            ("categoria", "select", True, ["Ingreso", "Egreso", "Retiro", "Aporte", "Otro"]),
            ("medio", "select", True, ["Efectivo", "Transferencia", "Tarjeta", "Cheque", "Otro"]),
            ("ingreso", "money", False),
            ("egreso", "money", False),
            ("responsable", "text", False),
            ("observaciones", "textarea", False),
        ],
    },
    "Banco Macro VMR": {
        "table": "banco_macro_vmr",
        "empresa": "VMR",
        "tipo": "flujo",
        "descripcion": "Movimientos bancarios de Banco Macro pertenecientes a VMR.",
        "fields": [
            ("fecha", "date", True),
            ("concepto", "text", True),
            ("tipo_movimiento", "select", True, ["Crédito", "Débito", "Transferencia", "Débito automático", "Impuesto", "Otro"]),
            ("referencia", "text", False),
            ("ingreso", "money", False),
            ("egreso", "money", False),
            ("conciliado", "bool", False),
            ("observaciones", "textarea", False),
        ],
    },
    "Cuenta Corriente VMR": {
        "table": "cuenta_corriente_vmr",
        "empresa": "VMR",
        "tipo": "cuenta_corriente",
        "descripcion": "Cuentas por cobrar y pagar de VMR.",
        "fields": [
            ("fecha", "date", True),
            ("persona_entidad", "text", True),
            ("concepto", "text", True),
            ("tipo", "select", True, ["A cobrar", "A pagar"]),
            ("importe", "money", True),
            ("pagado", "money", False),
            ("vencimiento", "date", False),
            ("estado", "select", True, ["Pendiente", "Parcial", "Pagado", "Vencido"]),
            ("observaciones", "textarea", False),
        ],
    },

    "Agenda Quirófano": {

        "table": "agenda_quirofano",
    
        "empresa": "VM",
    
        "tipo": "quirófano",
    
        "descripcion": "Agenda diaria, semanal y mensual de cirugías.",
    
        "fields": [

        ("fecha", "date", True),
    
        ("hora_inicio", "text", True),
    
        ("hora_fin", "text", True),
    
        ("duracion_min", "number", True),
    
        ("sala", "select", True, ["Quirófano 1"]),
    
        ("paciente", "text", True),
    
        ("procedimiento", "text", True),
    
        ("medico", "text", True),
    
        ("anestesista", "text", False),
    
        ("estado", "select", True,
    
         ["Programada", "En curso", "Finalizada", "Suspendida", "Cancelada"]),
    
        ("observaciones", "textarea", False),
    
    ],
  },
    "Facturación VMR": {
        "table": "facturacion_vmr",
        "empresa": "VMR",
        "tipo": "facturacion",
        "descripcion": "Control de facturación de procedimientos/pacientes de VMR según planilla quirófano.",
        "fields": [
            ("mes", "text", True),
            ("afiliado", "text", True),
            ("obra_social", "text", True),
            ("procedimiento", "text", True),
            ("medico_responsable", "text", True),
            ("fecha_factura", "date", False),
            ("numero_factura", "text", False),
            ("vencimiento", "date", False),
            ("fecha_pago", "date", False),
            ("valor_pesos", "money", True),
            ("valor_usd", "money", False),
            ("estado", "select", True, ["Pendiente", "Completo", "Parcial", "Vencido", "Anulado"]),
            ("observaciones", "textarea", False),
        ],
    },
    "Caja VM": {
        "table": "caja_vm",
        "empresa": "VM",
        "tipo": "flujo",
        "descripcion": "Movimientos de efectivo de Vitae Medical.",
        "fields": [
            ("fecha", "date", True),
            ("concepto", "text", True),
            ("categoria", "select", True, ["Ingreso", "Egreso", "Retiro", "Aporte", "Otro"]),
            ("medio", "select", True, ["Efectivo", "Transferencia", "Tarjeta", "Cheque", "Otro"]),
            ("ingreso", "money", False),
            ("egreso", "money", False),
            ("responsable", "text", False),
            ("observaciones", "textarea", False),
        ],
    },
    "Banco Galicia VM": {
        "table": "banco_galicia_vm",
        "empresa": "VM",
        "tipo": "flujo",
        "descripcion": "Movimientos bancarios de Banco Galicia pertenecientes a VM.",
        "fields": [
            ("fecha", "date", True),
            ("concepto", "text", True),
            ("tipo_movimiento", "select", True, ["Crédito", "Débito", "Transferencia", "Débito automático", "Impuesto", "Otro"]),
            ("referencia", "text", False),
            ("ingreso", "money", False),
            ("egreso", "money", False),
            ("conciliado", "bool", False),
            ("observaciones", "textarea", False),
        ],
    },
    "Cuenta Corriente VM": {
        "table": "cuenta_corriente_vm",
        "empresa": "VM",
        "tipo": "cuenta_corriente",
        "descripcion": "Cuentas por cobrar y pagar de VM.",
        "fields": [
            ("fecha", "date", True),
            ("persona_entidad", "text", True),
            ("concepto", "text", True),
            ("tipo", "select", True, ["A cobrar", "A pagar"]),
            ("importe", "money", True),
            ("pagado", "money", False),
            ("vencimiento", "date", False),
            ("estado", "select", True, ["Pendiente", "Parcial", "Pagado", "Vencido"]),
            ("observaciones", "textarea", False),
        ],
    },
    "Facturación VM": {
        "table": "facturacion_vm",
        "empresa": "VM",
        "tipo": "facturacion",
        "descripcion": "Control de facturación de procedimientos/pacientes de Vitae Medical según planilla quirófano.",
        "fields": [
            ("mes", "text", True),
            ("afiliado", "text", True),
            ("obra_social", "text", True),
            ("procedimiento", "text", True),
            ("medico_responsable", "text", True),
            ("fecha_factura", "date", False),
            ("numero_factura", "text", False),
            ("vencimiento", "date", False),
            ("fecha_pago", "date", False),
            ("valor_pesos", "money", True),
            ("valor_usd", "money", False),
            ("estado", "select", True, ["Pendiente", "Completo", "Parcial", "Vencido", "Anulado"]),
            ("observaciones", "textarea", False),
        ],
    },
    "Deudas Impositivas VMR": {
        "table": "deudas_impositivas_vmr",
        "empresa": "VMR",
        "tipo": "deuda",
        "descripcion": "IVA, Ganancias, cargas sociales, autónomos, monotributo u otros impuestos VMR.",
        "fields": [
            ("fecha", "date", True),
            ("impuesto", "select", True, ["IVA", "Ganancias", "Ingresos Brutos", "SUSS", "Monotributo", "Autónomos", "Municipal", "Otro"]),
            ("periodo", "text", True),
            ("importe", "money", True),
            ("pagado", "money", False),
            ("vencimiento", "date", True),
            ("estado", "select", True, ["Pendiente", "Parcial", "Pagado", "Plan de pago", "Vencido"]),
            ("observaciones", "textarea", False),
        ],
    },
    "Deudas Impositivas VM": {
        "table": "deudas_impositivas_vm",
        "empresa": "VM",
        "tipo": "deuda",
        "descripcion": "IVA, Ganancias, cargas sociales, autónomos, monotributo u otros impuestos VM.",
        "fields": [
            ("fecha", "date", True),
            ("impuesto", "select", True, ["IVA", "Ganancias", "Ingresos Brutos", "SUSS", "Monotributo", "Autónomos", "Municipal", "Otro"]),
            ("periodo", "text", True),
            ("importe", "money", True),
            ("pagado", "money", False),
            ("vencimiento", "date", True),
            ("estado", "select", True, ["Pendiente", "Parcial", "Pagado", "Plan de pago", "Vencido"]),
            ("observaciones", "textarea", False),
        ],
    },
    "Planes de pagos y préstamos": {
        "table": "planes_pagos_prestamos",
        "empresa": "VITAE",
        "tipo": "deuda",
        "descripcion": "Planes AFIP/ARCA, bancos, financieras, préstamos internos y externos.",
        "fields": [
            ("fecha", "date", True),
            ("empresa", "select", True, ["VMR", "VM", "VITAE"]),
            ("acreedor", "text", True),
            ("detalle", "text", True),
            ("cuotas_totales", "int", False),
            ("cuotas_pagadas", "int", False),
            ("importe_total", "money", True),
            ("saldo", "money", True),
            ("proximo_vencimiento", "date", False),
            ("estado", "select", True, ["Activo", "Finalizado", "Mora", "Refinanciado"]),
            ("observaciones", "textarea", False),
        ],
    },
    "Pagos pendientes Vitae": {
        "table": "pagos_pendientes_vitae",
        "empresa": "VITAE",
        "tipo": "pago_pendiente",
        "descripcion": "Pagos pendientes globales de la empresa.",
        "fields": [
            ("fecha", "date", True),
            ("empresa", "select", True, ["VMR", "VM", "VITAE"]),
            ("proveedor", "text", True),
            ("concepto", "text", True),
            ("importe", "money", True),
            ("pagado", "money", False),
            ("vencimiento", "date", False),
            ("prioridad", "select", True, ["Alta", "Media", "Baja"]),
            ("estado", "select", True, ["Pendiente", "Parcial", "Pagado", "Vencido"]),
            ("observaciones", "textarea", False),
        ],
    },
    "Gastos comunes": {
        "table": "gastos_comunes",
        "empresa": "VITAE",
        "tipo": "gasto",
        "descripcion": "Gastos compartidos entre VMR y VM.",
        "fields": [
            ("fecha", "date", True),
            ("rubro", "select", True, ["Luz", "Agua", "Gas", "Internet", "Limpieza", "Mantenimiento", "Sueldos", "Insumos", "Alquiler", "Otro"]),
            ("concepto", "text", True),
            ("importe", "money", True),
            ("porcentaje_vmr", "number", False),
            ("porcentaje_vm", "number", False),
            ("pagado", "bool", False),
            ("observaciones", "textarea", False),
        ],
    },
    "Vencimientos": {
        "table": "vencimientos",
        "empresa": "VITAE",
        "tipo": "vencimiento",
        "descripcion": "Calendario general de vencimientos administrativos, impositivos, contratos, habilitaciones e insumos.",
        "fields": [
            ("fecha", "date", True),
            ("empresa", "select", True, ["VMR", "VM", "VITAE"]),
            ("tipo_vencimiento", "select", True, ["Impuesto", "Servicio", "Contrato", "Habilitación", "Seguro", "Medicamento", "Mantenimiento", "Otro"]),
            ("detalle", "text", True),
            ("importe", "money", False),
            ("vencimiento", "date", True),
            ("estado", "select", True, ["Pendiente", "Realizado", "Vencido", "Reprogramado"]),
            ("observaciones", "textarea", False),
        ],
    },
    "Valores Alquileres": {
        "table": "valores_alquileres",
        "empresa": "VITAE",
        "tipo": "contrato_valor",
        "descripcion": "Control de alquileres, aumentos, valores mensuales y vigencia.",
        "fields": [
            ("fecha", "date", True),
            ("inmueble_area", "text", True),
            ("locador", "text", False),
            ("valor_mensual", "money", True),
            ("periodo", "text", True),
            ("fecha_desde", "date", False),
            ("fecha_hasta", "date", False),
            ("proximo_aumento", "date", False),
            ("estado", "select", True, ["Vigente", "Finalizado", "A renegociar"]),
            ("observaciones", "textarea", False),
        ],
    },
    "Tareas Pendientes": {
        "table": "tareas_pendientes",
        "empresa": "VITAE",
        "tipo": "tarea",
        "descripcion": "Seguimiento de tareas administrativas y operativas.",
        "fields": [
            ("fecha", "date", True),
            ("empresa", "select", True, ["VMR", "VM", "VITAE"]),
            ("tarea", "text", True),
            ("responsable", "text", False),
            ("prioridad", "select", True, ["Alta", "Media", "Baja"]),
            ("vencimiento", "date", False),
            ("estado", "select", True, ["Pendiente", "En proceso", "Finalizada", "Cancelada"]),
            ("observaciones", "textarea", False),
        ],
    },
    "Deuda total": {
        "table": "deuda_total_manual",
        "empresa": "VITAE",
        "tipo": "deuda",
        "descripcion": "Carga manual de deudas no contempladas en otros módulos. El tablero también calcula deuda total automática.",
        "fields": [
            ("fecha", "date", True),
            ("empresa", "select", True, ["VMR", "VM", "VITAE"]),
            ("acreedor", "text", True),
            ("concepto", "text", True),
            ("importe_original", "money", True),
            ("pagado", "money", False),
            ("saldo", "money", True),
            ("vencimiento", "date", False),
            ("estado", "select", True, ["Pendiente", "Parcial", "Pagado", "Vencido", "Refinanciado"]),
            ("observaciones", "textarea", False),
        ],
    },
    "Contratos": {
        "table": "contratos",
        "empresa": "VITAE",
        "tipo": "contrato",
        "descripcion": "Contratos con profesionales, proveedores, alquileres, servicios y convenios.",
        "fields": [
            ("fecha", "date", True),
            ("empresa", "select", True, ["VMR", "VM", "VITAE"]),
            ("contraparte", "text", True),
            ("tipo_contrato", "select", True, ["Profesional", "Proveedor", "Alquiler", "Servicio", "Convenio", "Otro"]),
            ("detalle", "text", True),
            ("inicio", "date", False),
            ("fin", "date", False),
            ("valor", "money", False),
            ("estado", "select", True, ["Vigente", "Vencido", "A renovar", "Finalizado"]),
            ("archivo_link", "text", False),
            ("observaciones", "textarea", False),
        ],
    },
    "Honorarios médicos": {
        "table": "honorarios_medicos",
        "empresa": "VITAE",
        "tipo": "honorario",
        "descripcion": "Honorarios por médico, prestación, estado de pago y empresa.",
        "fields": [
            ("fecha", "date", True),
            ("empresa", "select", True, ["VMR", "VM", "VITAE"]),
            ("medico", "text", True),
            ("paciente", "text", False),
            ("procedimiento", "text", True),
            ("importe", "money", True),
            ("pagado", "money", False),
            ("fecha_pago", "date", False),
            ("estado", "select", True, ["Pendiente", "Parcial", "Pagado"]),
            ("observaciones", "textarea", False),
        ],
    },
    "Gine Vitae": {
        "table": "gine_vitae",
        "empresa": "VM",
        "tipo": "unidad_medica",
        "descripcion": "Gestión de pacientes, prácticas, derivaciones y cirugías de la unidad Gine Vitae.",
        "fields": [
            ("fecha", "date", True),
            ("paciente", "text", True),
            ("dni", "text", False),
            ("telefono", "text", False),
            ("medico", "text", False),
            ("obra_social", "text", False),
            ("practica", "select", True, ["Consulta", "Control anual", "PAP", "Colposcopía", "HPV", "Ecografía", "Histeroscopía", "LEEP", "DIU", "Cirugía", "Otro"]),
            ("estado", "select", True, ["Pendiente", "Turno dado", "Realizado", "Derivado a quirófano", "Cancelado"]),
            ("importe", "money", False),
            ("cobrado", "money", False),
            ("proxima_accion", "text", False),
            ("observaciones", "textarea", False),
        ],
    },
}

# =========================================================
# BASE DE DATOS
# =========================================================

def connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def sql_type(field_type: str) -> str:
    if field_type in {"money", "number"}:
        return "REAL"
    if field_type == "int":
        return "INTEGER"
    if field_type == "bool":
        return "INTEGER"
    return "TEXT"

def init_db() -> None:
    """Crea tablas y agrega columnas nuevas si actualizás el esquema.

    SQLite no borra columnas viejas, pero la app muestra/exporta solo las columnas vigentes.
    """
    with connect() as conn:
        for cfg in MODULES.values():
            table = cfg["table"]
            columns = ["id INTEGER PRIMARY KEY AUTOINCREMENT", "created_at TEXT", "updated_at TEXT"]
            for field in cfg["fields"]:
                name, ftype = field[0], field[1]
                columns.append(f"{name} {sql_type(ftype)}")
            conn.execute(f"CREATE TABLE IF NOT EXISTS {table} ({', '.join(columns)})")

            existing_cols = [row[1] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()]
            for field in cfg["fields"]:
                name, ftype = field[0], field[1]
                if name not in existing_cols:
                    conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {sql_type(ftype)}")
        conn.execute("""
        CREATE TABLE IF NOT EXISTS app_settings (
            key TEXT PRIMARY KEY,
            value TEXT,
            updated_at TEXT
        )
        """)
        conn.commit()

def get_df(table: str) -> pd.DataFrame:
    with connect() as conn:
        try:
            return pd.read_sql_query(f"SELECT * FROM {table} ORDER BY id DESC", conn)
        except Exception:
            return pd.DataFrame()

def insert_row(table: str, data: Dict[str, Any]) -> None:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    data = {**data, "created_at": now, "updated_at": now}
    cols = list(data.keys())
    placeholders = ", ".join(["?"] * len(cols))
    sql = f"INSERT INTO {table} ({', '.join(cols)}) VALUES ({placeholders})"
    with connect() as conn:
        conn.execute(sql, [data[c] for c in cols])
        conn.commit()

def bulk_insert_rows(table: str, rows: List[Dict[str, Any]]) -> int:
    if not rows:
        return 0
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    clean_rows = [{**row, "created_at": now, "updated_at": now} for row in rows]
    cols = list(clean_rows[0].keys())
    placeholders = ", ".join(["?"] * len(cols))
    sql = f"INSERT INTO {table} ({', '.join(cols)}) VALUES ({placeholders})"
    values = [[row.get(c, "") for c in cols] for row in clean_rows]
    with connect() as conn:
        conn.executemany(sql, values)
        conn.commit()
    return len(clean_rows)

def replace_table_rows(table: str, rows: List[Dict[str, Any]]) -> int:
    with connect() as conn:
        conn.execute(f"DELETE FROM {table}")
        conn.commit()
    return bulk_insert_rows(table, rows)

def update_row(table: str, row_id: int, data: Dict[str, Any]) -> None:
    data = {**data, "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    sets = ", ".join([f"{k} = ?" for k in data.keys()])
    sql = f"UPDATE {table} SET {sets} WHERE id = ?"
    with connect() as conn:
        conn.execute(sql, [*data.values(), row_id])
        conn.commit()

def delete_row(table: str, row_id: int) -> None:
    with connect() as conn:
        conn.execute(f"DELETE FROM {table} WHERE id = ?", (row_id,))
        conn.commit()

# =========================================================
# HELPERS
# =========================================================

def normalize_money_string(value: Any) -> str:
    if value is None:
        return "0"
    try:
        if pd.isna(value):
            return "0"
    except Exception:
        pass
    text = str(value).strip()
    if text == "":
        return "0"
    text = text.replace("$", "").replace("ARS", "").replace("USD", "")
    text = text.replace(" ", "").replace("\u00a0", "")
    text = text.replace("(", "-").replace(")", "")
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(".", "").replace(",", ".")
    return text

def money(value: Any) -> float:
    try:
        if value is None:
            return 0.0
        if pd.isna(value) or value == "":
            return 0.0
        if isinstance(value, str):
            value = normalize_money_string(value)
        return float(value)
    except Exception:
        return 0.0

def fmt_money(value: Any) -> str:
    return f"$ {money(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def parse_date(value: Any) -> date | None:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if isinstance(value, str) and value.strip() == "":
        return None
    try:
        parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
        if pd.isna(parsed):
            return None
        return parsed.date()
    except Exception:
        return None

def clean_for_db(value: Any, ftype: str) -> Any:
    if ftype == "date":
        parsed = parse_date(value)
        return parsed.strftime(DATE_FMT) if parsed else ""
    if isinstance(value, date):
        return value.strftime(DATE_FMT)
    if ftype == "bool":
        return 1 if value else 0
    if ftype in {"money", "number"}:
        return float(value or 0)
    if ftype == "int":
        return int(value or 0)
    return value or ""

def default_value(ftype: str, options: List[str] | None = None) -> Any:
    if ftype == "date":
        return date.today()
    if ftype in {"money", "number"}:
        return 0.0
    if ftype == "int":
        return 0
    if ftype == "bool":
        return False
    if ftype == "select":
        return options[0] if options else ""
    return ""

def input_field(field: Tuple, prefix: str, existing: Dict[str, Any] | None = None) -> Any:
    name, ftype, required = field[0], field[1], field[2]
    options = field[3] if len(field) > 3 else None
    label = name.replace("_", " ").title() + (" *" if required else "")
    key = f"{prefix}_{name}"
    old = existing.get(name) if existing else None

    if ftype == "date":
        value = parse_date(old) if old else date.today()
        if value is None:
            value = date.today()
        return st.date_input(label, value=value, key=key)
    if ftype == "money":
        return st.number_input(label, min_value=0.0, step=1000.0, value=money(old), key=key)
    if ftype == "number":
        return st.number_input(label, step=1.0, value=float(money(old)), key=key)
    if ftype == "int":
        return st.number_input(label, min_value=0, step=1, value=int(money(old)), key=key)
    if ftype == "bool":
        return st.checkbox(label, value=bool(old), key=key)
    if ftype == "select":
        idx = 0
        if options and old in options:
            idx = options.index(old)
        return st.selectbox(label, options or [], index=idx, key=key)
    if ftype == "textarea":
        return st.text_area(label, value=str(old or ""), key=key)
    return st.text_input(label, value=str(old or ""), key=key)

def validate_required(cfg: Dict[str, Any], data: Dict[str, Any]) -> List[str]:
    errors = []
    for field in cfg["fields"]:
        name, ftype, required = field[0], field[1], field[2]
        if required and ftype not in {"money", "number", "int", "bool"} and not data.get(name):
            errors.append(name.replace("_", " ").title())
        if required and ftype in {"money", "number"} and money(data.get(name)) <= 0:
            errors.append(name.replace("_", " ").title())
    return errors

def get_field_names(cfg: Dict[str, Any]) -> List[str]:
    return [field[0] for field in cfg["fields"]]

def business_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    cols = [c for c in df.columns if c not in TECH_COLUMNS]
    return df[cols].copy()

def module_business_df(df: pd.DataFrame, cfg: Dict[str, Any]) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    field_names = get_field_names(cfg)
    calc_cols = ["saldo", "saldo_movimiento"]
    cols = [c for c in field_names + calc_cols if c in df.columns]
    return df[cols].copy()

def show_business_table(df: pd.DataFrame, height: int | None = None, **kwargs: Any) -> None:
    if height is not None:
        st.dataframe(business_df(df), use_container_width=True, hide_index=True, height=height, **kwargs)
    else:
        st.dataframe(business_df(df), use_container_width=True, hide_index=True, **kwargs)

def show_module_table(df: pd.DataFrame, cfg: Dict[str, Any], **kwargs: Any) -> None:
    st.dataframe(module_business_df(df, cfg), use_container_width=True, hide_index=True, **kwargs)

def add_balance_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    if "ingreso" in df.columns and "egreso" in df.columns:
        df["saldo_movimiento"] = df["ingreso"].apply(money) - df["egreso"].apply(money)
    if "importe" in df.columns and "pagado" in df.columns:
        df["saldo"] = df["importe"].apply(money) - df["pagado"].apply(money)
    if "valor_pesos" in df.columns:
        df["valor_pesos"] = df["valor_pesos"].apply(money)
    if "valor_usd" in df.columns:
        df["valor_usd"] = df["valor_usd"].apply(money)
    if "importe_total" in df.columns and "saldo" in df.columns:
        df["saldo"] = df["saldo"].apply(money)
    if "importe_original" in df.columns and "saldo" in df.columns:
        df["saldo"] = df["saldo"].apply(money)
    return df

def first_available_date_col(df: pd.DataFrame, module_name: str) -> str | None:
    if module_name in ["Facturación VMR", "Facturación VM"] and "fecha_factura" in df.columns:
        return "fecha_factura"
    for candidate in ["fecha", "vencimiento", "fecha_pago", "fecha_cobro", "proximo_vencimiento", "fecha_desde", "fecha_hasta"]:
        if candidate in df.columns:
            return candidate
    return None

def apply_filters(df: pd.DataFrame, module_name: str) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    st.subheader("Filtros")
    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)

    with c1:
        search = st.text_input("Buscar texto", key=f"search_{module_name}")
    with c2:
        estado = "Todos"
        if "estado" in df.columns:
            estados = [str(x).strip() for x in df["estado"].dropna().unique().tolist() if str(x).strip() != ""]
            estado = st.selectbox("Estado", ["Todos"] + sorted(estados), key=f"estado_{module_name}")
    with c3:

        obra_social = "Todos"

        if "obra_social" in df.columns:

            obra_social = st.selectbox(

                "Obra Social",

                ["Todos"] + sorted(df["obra_social"].dropna().astype(str).unique().tolist())

            )

    with c4:

        procedimiento = "Todos"

        if "procedimiento" in df.columns:

            procedimiento = st.selectbox(

                "Procedimiento",

                ["Todos"] + sorted(df["procedimiento"].dropna().astype(str).unique().tolist())

            )

    with c5:

        medico = "Todos"

        if "medico_responsable" in df.columns:

            medico = st.selectbox(

                "Médico",

                ["Todos"] + sorted(df["medico_responsable"].dropna().astype(str).unique().tolist())

            )
    with c6:
        fecha_desde = st.date_input("Desde", value=date.today() - timedelta(days=3650), key=f"desde_{module_name}")
    with c7:
        fecha_hasta = st.date_input("Hasta", value=date.today() + timedelta(days=3650), key=f"hasta_{module_name}")

    if search:
        mask = df.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)
        df = df[mask]
    if "estado" in df.columns and estado != "Todos":
        df = df[df["estado"].astype(str).str.strip() == estado]
    if obra_social != "Todos":

        df = df[df["obra_social"].astype(str) == obra_social]

    if procedimiento != "Todos":

        df = df[df["procedimiento"].astype(str) == procedimiento]

    if medico != "Todos":

        df = df[df["medico_responsable"].astype(str) == medico]
           
    fecha_col = first_available_date_col(df, module_name)
    if fecha_col:
        fechas = pd.to_datetime(df[fecha_col], errors="coerce")
        desde_ts = pd.Timestamp(fecha_desde)
        hasta_ts = pd.Timestamp(fecha_hasta)
        # Conserva filas sin fecha para que no desaparezcan registros importados con fecha_factura vacía.
        df = df[fechas.isna() | ((fechas >= desde_ts) & (fechas <= hasta_ts))]
    return df

# =========================================================
# IMPORTADOR EXCEL / CSV
# =========================================================

def clean_tabular_sheet(df_raw: pd.DataFrame) -> pd.DataFrame:
    if df_raw.empty:
        return df_raw
    raw = df_raw.copy().dropna(how="all").dropna(axis=1, how="all")
    if raw.empty:
        return raw

    header_keywords = {
        "mes", "afiliado", "obra social", "procedimiento", "medico", "médico",
        "fecha factura", "factura", "vencimiento", "fecha pago", "valor", "estado",
        "cliente", "paciente", "importe", "concepto", "comprobante"
    }

    best_idx = raw.index[0]
    best_score = -1.0
    for idx, row in raw.iterrows():
        values = [str(x).strip().lower() for x in row.tolist() if pd.notna(x) and str(x).strip() != ""]
        if not values:
            continue
        joined = " | ".join(values)
        score = sum(1 for kw in header_keywords if kw in joined) + min(len(values), 10) * 0.05
        if score > best_score:
            best_score = score
            best_idx = idx

    header_values = raw.loc[best_idx].tolist()
    columns: List[str] = []
    used: Dict[str, int] = {}
    for i, value in enumerate(header_values):
        name = str(value).strip() if pd.notna(value) and str(value).strip() else f"Columna_{i + 1}"
        name = name.replace("\n", " ").replace("  ", " ").strip()
        if name in used:
            used[name] += 1
            name = f"{name}_{used[name]}"
        else:
            used[name] = 1
        columns.append(name)

    cleaned = raw.loc[raw.index > best_idx].copy()
    cleaned.columns = columns
    cleaned = cleaned.dropna(how="all")
    cleaned = cleaned.loc[:, [not str(c).lower().startswith("columna_") or not cleaned[c].isna().all() for c in cleaned.columns]]
    return cleaned.reset_index(drop=True)

def read_uploaded_sheet(uploaded_file: Any) -> Dict[str, pd.DataFrame]:
    filename = uploaded_file.name.lower()
    if filename.endswith(".csv"):
        try:
            raw = pd.read_csv(uploaded_file, sep=None, engine="python", header=None)
        except Exception:
            uploaded_file.seek(0)
            raw = pd.read_csv(uploaded_file, header=None)
        return {"CSV": clean_tabular_sheet(raw)}
    raw_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)
    return {name: clean_tabular_sheet(df) for name, df in raw_sheets.items()}

def field_label(field: Tuple) -> str:
    name, _ftype, required = field[0], field[1], field[2]
    return f"{name.replace('_', ' ').title()}{' *' if required else ''}"

def auto_guess_column(target_name: str, source_columns: List[str]) -> str:
    norm_target = target_name.lower().replace("_", " ")
    aliases = {
        "fecha": ["fecha", "dia", "día", "date"],
        "mes": ["mes", "periodo", "período"],
        "afiliado": ["afiliado", "paciente", "cliente", "nombre", "apellido y nombre"],
        "obra_social": ["obra social", "os", "prepaga"],
        "procedimiento": ["procedimiento", "practica", "práctica", "prestacion", "prestación"],
        "medico_responsable": ["medico responsable", "médico responsable", "medico", "médico", "doctor", "profesional", "responsable"],
        "fecha_factura": ["fecha factura", "fecha de factura", "fecha", "factura fecha"],
        "numero_factura": ["n° factura", "nº factura", "n factura", "numero factura", "número factura", "factura", "comprobante"],
        "fecha_pago": ["fecha pago", "fecha de pago", "pago fecha"],
        "valor_pesos": ["valor $", "valor pesos", "valor ars", "importe", "monto", "total", "valor"],
        "valor_usd": ["valor usd", "usd", "dolares", "dólares"],
        "cliente": ["cliente", "paciente", "nombre", "razon social", "razón social"],
        "concepto": ["concepto", "detalle", "descripcion", "descripción", "movimiento", "observacion"],
        "detalle": ["detalle", "concepto", "descripcion", "descripción"],
        "persona_entidad": ["persona", "entidad", "cliente", "proveedor", "paciente", "nombre"],
        "proveedor": ["proveedor", "acreedor", "contraparte", "entidad"],
        "acreedor": ["acreedor", "proveedor", "banco", "entidad"],
        "contraparte": ["contraparte", "proveedor", "profesional", "locador"],
        "medico": ["medico", "médico", "doctor", "profesional"],
        "importe": ["importe", "monto", "total", "valor", "debe", "saldo"],
        "importe_total": ["importe total", "total", "monto", "importe"],
        "importe_original": ["importe original", "deuda", "total", "importe", "monto"],
        "valor": ["valor", "importe", "monto", "total"],
        "valor_mensual": ["valor mensual", "alquiler", "importe", "monto", "total"],
        "ingreso": ["ingreso", "entradas", "haber", "credito", "crédito", "cobro"],
        "egreso": ["egreso", "salidas", "debe", "debito", "débito", "pago"],
        "pagado": ["pagado", "pago", "abonado", "cancelado"],
        "cobrado": ["cobrado", "cobro", "pagado", "abonado"],
        "saldo": ["saldo", "pendiente", "resta", "deuda"],
        "estado": ["estado", "situacion", "situación", "status"],
        "vencimiento": ["vencimiento", "vence", "fecha vencimiento"],
        "proximo_vencimiento": ["proximo vencimiento", "próximo vencimiento", "vencimiento", "vence"],
        "observaciones": ["observaciones", "observacion", "obs", "nota", "comentario"],
        "responsable": ["responsable", "usuario", "encargado"],
        "dni": ["dni", "documento"],
        "telefono": ["telefono", "teléfono", "celular", "whatsapp"],
        "practica": ["practica", "práctica", "prestacion", "prestación", "procedimiento"],
        "periodo": ["periodo", "período", "mes"],
        "comprobante": ["comprobante", "factura", "n factura", "n° factura", "nº factura"],
    }
    candidates = aliases.get(target_name, [norm_target])
    normalized_sources = {str(col).lower().replace("_", " ").strip(): col for col in source_columns}
    for cand in candidates:
        cand = cand.lower().strip()
        if cand in normalized_sources:
            return normalized_sources[cand]
    for cand in candidates:
        cand = cand.lower().strip()
        for src_norm, original in normalized_sources.items():
            if cand in src_norm or src_norm in cand:
                return original
    return "No usar"

def normalize_select_value(value: Any, options: List[str]) -> str:
    if value is None:
        return options[0] if options else ""
    try:
        if pd.isna(value):
            return options[0] if options else ""
    except Exception:
        pass
    text = str(value).strip()
    if text == "":
        return options[0] if options else ""
    for opt in options:
        if text.lower() == opt.lower():
            return opt
    aliases = {
        "cobrado": "Cobrado", "pagado": "Pagado", "pendiente": "Pendiente", "vencido": "Vencido",
        "parcial": "Parcial", "completo": "Completo", "completa": "Completo",
        "realizado": "Realizado", "finalizada": "Finalizada", "finalizado": "Finalizado",
        "alta": "Alta", "media": "Media", "baja": "Baja",
        "credito": "Crédito", "crédito": "Crédito", "debito": "Débito", "débito": "Débito",
    }
    wanted = aliases.get(text.lower())
    if wanted and wanted in options:
        return wanted
    return options[0] if options else text

def clean_import_value(value: Any, field: Tuple) -> Any:
    _name, ftype = field[0], field[1]
    options = field[3] if len(field) > 3 else None
    if ftype == "date":
        parsed = parse_date(value)
        return parsed.strftime(DATE_FMT) if parsed else ""
    if ftype in {"money", "number"}:
        num = pd.to_numeric(normalize_money_string(value), errors="coerce")
        return 0.0 if pd.isna(num) else float(num)
    if ftype == "int":
        num = pd.to_numeric(normalize_money_string(value), errors="coerce")
        return 0 if pd.isna(num) else int(num)
    if ftype == "bool":
        if value is None:
            return 0
        try:
            if pd.isna(value):
                return 0
        except Exception:
            pass
        return 1 if str(value).strip().lower() in ["1", "true", "si", "sí", "x", "ok", "pagado", "conciliado"] else 0
    if ftype == "select":
        return normalize_select_value(value, options or [])
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()

def render_importer(module_name: str, cfg: Dict[str, Any]) -> None:
    table = cfg["table"]
    st.subheader("Importar planilla Excel / CSV")
    st.caption("Subí una planilla, elegí la hoja, mapeá columnas y guardala dentro de este módulo.")

    uploaded_file = st.file_uploader("Subir archivo", type=["xlsx", "xls", "csv"], key=f"upload_{table}")
    if uploaded_file is None:
        st.info("Acepta Excel con varias hojas o CSV.")
        return

    try:
        sheets = read_uploaded_sheet(uploaded_file)
    except Exception as e:
        st.error(f"No pude leer el archivo. Detalle: {e}")
        return

    sheet_names = list(sheets.keys())
    selected_sheet = st.selectbox("Hoja a importar", sheet_names, key=f"sheet_{table}")
    df_original = sheets[selected_sheet].copy().dropna(how="all")
    df_original.columns = [str(c).strip() for c in df_original.columns]

    if df_original.empty:
        st.warning("La hoja seleccionada está vacía.")
        return

    st.markdown("#### Vista previa")
    show_business_table(
        df_original,
        height=700,
    )

    columnas = df_original.columns.tolist()
    st.markdown("#### Mapeo de columnas")
    mapping: Dict[str, str] = {}
    cols = st.columns(2)
    for i, field in enumerate(cfg["fields"]):
        name = field[0]
        guessed = auto_guess_column(name, columnas)
        options = ["No usar"] + columnas
        index = options.index(guessed) if guessed in options else 0
        with cols[i % 2]:
            mapping[name] = st.selectbox(field_label(field), options, index=index, key=f"map_{table}_{name}")

    with st.expander("Opciones avanzadas"):
        modo = st.radio("Modo de importación", ["Agregar a registros existentes", "Reemplazar módulo completo"], key=f"modo_import_{table}")
        saltar_filas_vacias = st.checkbox("Saltar filas completamente vacías", value=True, key=f"skip_empty_{table}")
        validar_obligatorios = st.checkbox("Validar campos obligatorios", value=False, key=f"valid_required_{table}")

    rows: List[Dict[str, Any]] = []
    rejected_rows: List[Dict[str, Any]] = []
    for idx, source_row in df_original.iterrows():
        if saltar_filas_vacias and source_row.isna().all():
            continue
        new_row: Dict[str, Any] = {}
        for field in cfg["fields"]:
            name = field[0]
            mapped_col = mapping.get(name, "No usar")
            if mapped_col == "No usar":
                new_row[name] = clean_for_db(default_value(field[1], field[3] if len(field) > 3 else None), field[1])
                if field[1] == "date" and not field[2]:
                    new_row[name] = ""
            else:
                new_row[name] = clean_import_value(source_row.get(mapped_col), field)
        errors = validate_required(cfg, new_row) if validar_obligatorios else []
        if errors:
            rejected_rows.append({"fila_excel": idx + 2, "motivo": ", ".join(errors), **new_row})
        else:
            rows.append(new_row)

    st.markdown("#### Previsualización final")
    preview_df = pd.DataFrame(rows)
    if preview_df.empty:
        st.warning("No hay filas válidas para importar con el mapeo actual.")
    else:
        show_business_table(preview_df.head(50))
        st.success(f"Filas listas para importar: {len(rows)}")

    if rejected_rows:
        with st.expander(f"Filas rechazadas: {len(rejected_rows)}"):
            show_business_table(pd.DataFrame(rejected_rows))

    col_a, col_b = st.columns([1, 2])
    with col_a:
        confirm_import = st.checkbox("Confirmo la importación", key=f"confirm_import_{table}")
    with col_b:
        st.caption("Si reemplazás el módulo completo, se borran los registros anteriores de este módulo.")

    if st.button("Importar planilla al módulo", type="primary", disabled=(not confirm_import or not rows), key=f"btn_import_{table}"):
        count = replace_table_rows(table, rows) if modo == "Reemplazar módulo completo" else bulk_insert_rows(table, rows)
        st.success(f"Importación completada. Registros importados en {module_name}: {count}")
        st.rerun()

# =========================================================
# VISTAS
# =========================================================

def render_header() -> None:



    col1, col2 = st.columns([6.5, 1.2])



    with col1:



        st.markdown(



            '<div class="main-title">🏥 Sistema de Gestión | VITAE </div>',



            unsafe_allow_html=True



        )



        st.markdown(



            '<div class="subtitle">VMR · Vitae Medicina Reproductiva | VM · Vitae Medical</div>',



            unsafe_allow_html=True



        )



    with col2:



        logo_path = Path("logo_vitae.png")



        if logo_path.exists():



            st.markdown(



                """



                <style>



                .vitae-logo img {



                    width: 170px !important;



                    max-width: 170px !important;



                }



                </style>



                """,



                unsafe_allow_html=True,



            )



            st.markdown('<div class="vitae-logo">', unsafe_allow_html=True)



            st.image(str(logo_path))



            st.markdown('</div>', unsafe_allow_html=True)



        else:



            st.warning("Logo no encontrado")
def get_setting(key: str, default: Any = None) -> Any:

    with connect() as conn:

        row = conn.execute("SELECT value FROM app_settings WHERE key = ?", (key,)).fetchone()

    if row is None:

        return default

    try:

        return json.loads(row["value"])

    except Exception:

        return row["value"]

def set_setting(key: str, value: Any) -> None:

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    payload = json.dumps(value, ensure_ascii=False)

    with connect() as conn:

        conn.execute(

            """

            INSERT INTO app_settings (key, value, updated_at)

            VALUES (?, ?, ?)

            ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at

            """,

            (key, payload, now),

        )

        conn.commit()

DEFAULT_FACT_LABELS = {

    "mes": "Mes",

    "afiliado": "Paciente / Afiliado",

    "obra_social": "Obra social",

    "procedimiento": "Procedimiento",

    "medico_responsable": "Médico",

    "fecha_factura": "Fecha factura",

    "numero_factura": "N° factura",

    "vencimiento": "Vencimiento",

    "fecha_pago": "Fecha pago",

    "valor_pesos": "Valor facturado",

    "valor_usd": "Valor USD",

    "estado": "Estado",

    "observaciones": "Observaciones",

}

def get_fact_labels(module_name: str, cfg: Dict[str, Any]) -> Dict[str, str]:

    key = f"labels_{cfg['table']}"

    saved = get_setting(key, {})

    labels = DEFAULT_FACT_LABELS.copy()

    if isinstance(saved, dict):

        labels.update(saved)

    return labels

def rename_fact_df(df: pd.DataFrame, labels: Dict[str, str]) -> pd.DataFrame:

    return df.rename(columns={c: labels.get(c, c.replace("_", " ").title()) for c in df.columns})

def format_facturacion_table(df: pd.DataFrame, labels: Dict[str, str]) -> pd.DataFrame:

    if df.empty:

        return df

    show = df.copy()

    if "mes" in show.columns:
        show["mes"] = pd.to_datetime(
            show["mes"],
            errors="coerce"            
        ).dt.strftime("%d/%m/%Y")

    show = show.drop(
        columns=[
            "id",
            "created_at",
            "updated_at"
        ],
        errors="ignore"
    )


    for col in ["fecha_factura", "vencimiento", "fecha_pago"]:

        if col in show.columns:

            show[col] = pd.to_datetime(show[col], errors="coerce").dt.strftime("%d/%m/%Y")

            show[col] = show[col].fillna("")

    for col in ["valor_pesos"]:

        if col in show.columns:

            show[col] = show[col].apply(fmt_money)

    if "valor_usd" in show.columns:

        show["valor_usd"] = show["valor_usd"].apply(lambda x: f"USD {money(x):,.2f}")

    

    show = show.rename(columns={c: labels.get(c, c.replace("_", " ").title()) for c in show.columns})

    return show

def render_facturacion_pro(module_name: str, cfg: Dict[str, Any]) -> None:

    table = cfg["table"]

    render_header()

    st.header(module_name)

    st.caption(cfg["descripcion"])

    labels = get_fact_labels(module_name, cfg)

    tab_panel, tab_cargar, tab_importar, tab_editar, tab_columnas, tab_exportar = st.tabs([

        "📊 Panel PRO",

        "➕ Cargar",

        "📥 Importar",

        "✏️ Editar tabla",

        "🏷️ Editar columnas",

        "📤 Exportar",

    ])

    with tab_panel:



        df = add_balance_columns(get_df(table))

        filtered = df.copy()

        if df.empty:



            st.warning("No hay registros cargados.")



        else:



            filtered = apply_filters(df, module_name)



            total_facturado = filtered["valor_pesos"].apply(money).sum()



            cobrado = filtered[



                filtered["estado"].astype(str).str.lower().isin(["completo", "pagado", "cobrado"])



            ]["valor_pesos"].apply(money).sum()



            pendiente = total_facturado - cobrado



            pacientes = len(filtered)



            ticket_promedio = total_facturado / pacientes if pacientes > 0 else 0



            c1, c2, c3, c4, = st.columns(4)



            c1.metric("💰 Facturado", fmt_money(total_facturado))



            c2.metric("✅ Cobrado", fmt_money(cobrado))



            c3.metric("⏳ Pendiente", fmt_money(pendiente))



            c4.metric("👥 Pacientes", pacientes)



            #c5.metric("📊 Ticket Promedio", fmt_money(ticket_promedio))



        st.divider()

        if "mes" in filtered.columns:

            filtered = filtered.sort_values(

                by="mes",

                ascending=False,

                na_position="last"

            )

        col_orden = "mes" if "mes" in filtered.columns else "Mes" if "Mes" in filtered.columns else None

        if col_orden:

            filtered[col_orden] = pd.to_datetime(

                filtered[col_orden],

                errors="coerce",

                dayfirst=True

            )

            filtered = filtered.sort_values(

                by=col_orden,

                ascending=False,

                na_position="last"

            )

        if not filtered.empty:
            st.markdown("### Tabla limpia")
            st.dataframe(
                format_facturacion_table(filtered, labels),
                use_container_width=True,
                hide_index=True,
            )

    st.divider()


    st.markdown("### Gráficos útiles")

    g1, g2 = st.columns(2)

    if "fecha_factura" in filtered.columns and "valor_pesos" in filtered.columns:

        graph = filtered.copy()

        graph["fecha_factura"] = pd.to_datetime(graph["fecha_factura"], errors="coerce")

        graph = graph[graph["fecha_factura"].notna()]

        if not graph.empty:

            graph["Mes"] = graph["fecha_factura"].dt.to_period("M").astype(str)

            chart = graph.groupby("Mes")["valor_pesos"].sum().reset_index()

            fig = px.bar(chart, x="Mes", y="valor_pesos", title="Facturación por mes")

            g1.plotly_chart(fig, use_container_width=True)

    if "obra_social" in filtered.columns and "valor_pesos" in filtered.columns:

        chart = filtered.groupby("obra_social")["valor_pesos"].apply(lambda x: x.apply(money).sum()).reset_index()

        chart = chart.sort_values("valor_pesos", ascending=False).head(10)

        fig = px.bar(chart, x="obra_social", y="valor_pesos", title="Facturación por obra social")

        g2.plotly_chart(fig, use_container_width=True)

    g3, g4 = st.columns(2)

    if "medico_responsable" in filtered.columns and "valor_pesos" in filtered.columns:

        chart = filtered.groupby("medico_responsable")["valor_pesos"].apply(lambda x: x.apply(money).sum()).reset_index()

        chart = chart.sort_values("valor_pesos", ascending=False).head(10)

        fig = px.bar(chart, x="medico_responsable", y="valor_pesos", title="Facturación por médico")

        g3.plotly_chart(fig, use_container_width=True)

    if "procedimiento" in filtered.columns and "valor_pesos" in filtered.columns:

        chart = filtered.groupby("procedimiento")["valor_pesos"].apply(lambda x: x.apply(money).sum()).reset_index()

        chart = chart.sort_values("valor_pesos", ascending=False).head(10)

        fig = px.bar(chart, x="procedimiento", y="valor_pesos", title="Facturación por procedimiento")

        g4.plotly_chart(fig, use_container_width=True)
    
    with tab_cargar:

        st.subheader("Nuevo registro")

        with st.form(f"form_add_{table}", clear_on_submit=True):

            data: Dict[str, Any] = {}

            cols = st.columns(2)

            for i, field in enumerate(cfg["fields"]):

                with cols[i % 2]:

                    raw = input_field(field, f"add_{table}")

                    data[field[0]] = clean_for_db(raw, field[1])

            submitted = st.form_submit_button("Guardar registro", type="primary")

            if submitted:

                errors = validate_required(cfg, data)

                if errors:

                    st.error("Faltan completar campos obligatorios: " + ", ".join(errors))

                else:

                    insert_row(table, data)

                    st.success("Registro guardado correctamente.")

                    st.rerun()

    with tab_importar:

        render_importer(module_name, cfg)
    
    with tab_editar:

        st.subheader("Editar registros cargados")

        df = get_df(table)

        if df.empty:

            st.warning("No hay registros para editar.")

        else:

            df_edit = df.copy()

            columnas_ocultas = ["created_at", "updated_at"]

            df_edit = df_edit.drop(columns=columnas_ocultas, errors="ignore")

            # Ordenar desde la fecha actual hacia atrás

            if "mes" in df_edit.columns:

                df_edit["mes"] = pd.to_datetime(df_edit["mes"], errors="coerce").dt.strftime("%Y-%m-%d")


            if "mes" in df_edit.columns:

                df_edit["mes"] = pd.to_datetime(

                    df_edit["mes"],

                    errors="coerce"

                )

                df_edit = df_edit.sort_values(

                    by="mes",

                    ascending=False

                )

                df_edit["mes"] = df_edit["mes"].dt.strftime("%Y-%m-%d")
              
            estado_editor = "Todos"

            if "estado" in df_edit.columns:

                estados = sorted(

                    df_edit["estado"].dropna().astype(str).unique().tolist()

                )

                estado_editor = st.selectbox(

                    "Filtrar por estado",

                    ["Todos"] + estados,

                    key=f"estado_editor_{table}"

                )

                if estado_editor != "Todos":

                    df_edit = df_edit[

                        df_edit["estado"].astype(str) == estado_editor

                    ]

            edited_df = st.data_editor(

                df_edit,
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                disabled=["id"]
            )
            col1, col2 = st.columns(2)
            with col1:
                guardar = st.button("Guardar cambios", type="primary", key=f"guardar_editor_{table}")
            with col2:
               st.warning("Si borrás filas en la tabla y guardás, se eliminan de la base.")
            if guardar:



                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")



                rows = edited_df.to_dict("records")



                with connect() as conn:



                    conn.execute(f"DELETE FROM {table}")



                    for row in rows:



                        row.pop("id", None)



                        row["created_at"] = now



                        row["updated_at"] = now



                        cols = list(row.keys())



                        placeholders = ", ".join(["?"] * len(cols))



                        conn.execute(



                            f"INSERT INTO {table} ({', '.join(cols)}) VALUES ({placeholders})",



                            [row[c] for c in cols]



                        )



                    conn.commit()



                st.success("Tabla actualizada correctamente.")



                st.rerun()



    with tab_columnas:

        st.subheader("Editar nombres visibles de columnas")

        with st.form(f"form_labels_{table}"):

            new_labels = {}

            cols = st.columns(2)

            for i, field in enumerate(cfg["fields"]):

                name = field[0]

                with cols[i % 2]:

                    new_labels[name] = st.text_input(

                        name,

                        value=labels.get(name, name.replace("_", " ").title()),

                        key=f"label_{table}_{name}",

                    )

            save_labels = st.form_submit_button("Guardar nombres de columnas", type="primary")

            if save_labels:

                set_setting(f"labels_{table}", new_labels)

                st.success("Nombres de columnas actualizados.")

                st.rerun()

    with tab_exportar:

        df = add_balance_columns(get_df(table))
        

        if df.empty:

            st.info("No hay datos para exportar.")

        else:

            export_df = format_facturacion_table(df, labels)

            csv = export_df.to_csv(index=False).encode("utf-8-sig")

            st.download_button("Descargar CSV", data=csv, file_name=f"{table}.csv", mime="text/csv")

            xlsx_path = Path(f"{table}.xlsx")

            export_df.to_excel(xlsx_path, index=False)

            with open(xlsx_path, "rb") as f:

                st.download_button(

                    "Descargar Excel",

                    data=f,

                    file_name=f"{table}.xlsx",

                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

                )
def render_dashboard() -> None:
    render_header()
    st.markdown("### Resumen General")
    dfs = {name: add_balance_columns(get_df(cfg["table"])) for name, cfg in MODULES.items()}
    def total_mod(nombre):

        df = dfs.get(nombre, pd.DataFrame())

        if df.empty:

            return 0.0

        if "saldo" in df.columns:

            return df["saldo"].apply(money).sum()

        if "saldo_movimiento" in df.columns:

            return df["saldo_movimiento"].apply(money).sum()

        if "importe" in df.columns:

            return df["importe"].apply(money).sum()

        if "valor_pesos" in df.columns:

            return df["valor_pesos"].apply(money).sum()

        if "monto" in df.columns:

            return df["monto"].apply(money).sum()

        return 0.0

    caja_vmr = total_mod("Caja VMR")

    banco_vmr = total_mod("Banco Macro VMR")

    caja_vm = total_mod("Caja VM")

    banco_vm = total_mod("Banco Galicia VM")

    gine_vitae = total_mod("Gine Vitae")

    pagos_pendientes = total_mod("Pagos pendientes Vitae")

    planes_pago = total_mod("Planes de pagos y préstamos")

    honorarios = total_mod("Honorarios médicos")

    deuda_imp_vmr = total_mod("Deudas Impositivas VMR")

    deuda_imp_vm = total_mod("Deudas Impositivas VM")

    liquidez_total = caja_vmr + banco_vmr + caja_vm + banco_vm + gine_vitae

    deuda_total_global = pagos_pendientes + planes_pago + honorarios + deuda_imp_vmr + deuda_imp_vm
    caja_bancos = 0.0
    ingresos_mes = 0.0
    egresos_mes = 0.0
    facturacion_mes = 0.0
    cobrado_mes = 0.0
    a_cobrar = 0.0
    a_pagar = 0.0
    deuda_total = 0.0
    vencidos = 0
    tareas_pend = 0
    pacientes_mes = 0
    medicos_activos = set()
    hoy = pd.Timestamp.today().normalize()
    inicio_mes = hoy.replace(day=1)
    fin_mes = inicio_mes + pd.offsets.MonthEnd(0)
    estados_cerrados = ["pagado", "cobrado", "completo", "realizado", "finalizada", "finalizado", "anulado", "cancelado"]
    for name, df in dfs.items():
        if df.empty:
            continue
        if "fecha" in df.columns:
            fechas = pd.to_datetime(df["fecha"], errors="coerce")
        elif "fecha_factura" in df.columns:
            fechas = pd.to_datetime(df["fecha_factura"], errors="coerce")
        else:
            fechas = pd.Series([pd.NaT] * len(df), index=df.index)
        es_mes = fechas.notna() & (fechas >= inicio_mes) & (fechas <= fin_mes)
        if name in ["Caja VMR", "Caja VM", "Banco Macro VMR", "Banco Galicia VM"]:
            ingresos = df["ingreso"].apply(money).sum() if "ingreso" in df.columns else 0
            egresos = df["egreso"].apply(money).sum() if "egreso" in df.columns else 0
            caja_bancos += ingresos - egresos
            if "ingreso" in df.columns:
                ingresos_mes += df.loc[es_mes, "ingreso"].apply(money).sum()
            if "egreso" in df.columns:
                egresos_mes += df.loc[es_mes, "egreso"].apply(money).sum()
        if name in ["Facturación VMR", "Facturación VM"]:
            if "valor_pesos" in df.columns:
                total_facturado = df["valor_pesos"].apply(money).sum()
                facturacion_mes += df.loc[es_mes, "valor_pesos"].apply(money).sum()
                estado = df["estado"].astype(str).str.lower().str.strip() if "estado" in df.columns else pd.Series([""] * len(df), index=df.index)
                cobrado = df[estado.isin(["completo", "cobrado", "pagado"])]["valor_pesos"].apply(money).sum()
                cobrado_mes += df.loc[es_mes & estado.isin(["completo", "cobrado", "pagado"]), "valor_pesos"].apply(money).sum()
                a_cobrar += max(0, total_facturado - cobrado)
                pacientes_mes += int(es_mes.sum())
                if "medico_responsable" in df.columns:
                    medicos_activos.update(
                        df.loc[es_mes, "medico_responsable"]
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .replace("", pd.NA)
                        .dropna()
                        .tolist()
                    )
        if name in ["Cuenta Corriente VMR", "Cuenta Corriente VM"]:
            if "tipo" in df.columns and "importe" in df.columns:
                tipo = df["tipo"].astype(str).str.lower()
                pagado = df["pagado"].apply(money) if "pagado" in df.columns else 0
                saldo = df["importe"].apply(money) - pagado
                a_cobrar += saldo[tipo.eq("a cobrar")].sum()
                a_pagar += saldo[tipo.eq("a pagar")].sum()
        if name in ["Deudas Impositivas VMR", "Deudas Impositivas VM", "Planes de pagos y préstamos", "Pagos pendientes Vitae", "Deuda total", "Honorarios médicos"]:
            if "saldo" in df.columns:
                deuda_total += df["saldo"].apply(money).sum()
            elif "importe" in df.columns:
                pagado = df["pagado"].apply(money) if "pagado" in df.columns else 0
                deuda_total += max(0, df["importe"].apply(money).sum() - pagado.sum())
        if "vencimiento" in df.columns:
            venc = pd.to_datetime(df["vencimiento"], errors="coerce")
            estado = df["estado"].astype(str).str.lower().str.strip() if "estado" in df.columns else pd.Series([""] * len(df), index=df.index)
            vencidos += int((venc.notna() & (venc < hoy) & (~estado.isin(estados_cerrados))).sum())
        if name == "Tareas Pendientes" and "estado" in df.columns:
            tareas_pend += int(df[~df["estado"].isin(["Finalizada", "Cancelada"])].shape[0])
    resultado_mes = ingresos_mes + cobrado_mes - egresos_mes
    pendiente_cobro = a_cobrar
    promedio_facturacion = facturacion_mes / pacientes_mes if pacientes_mes > 0 else 0
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Liquidez actual", fmt_money(caja_bancos))
    c2.metric("Facturación mes", fmt_money(facturacion_mes))
    c3.metric("Cobrado mes", fmt_money(cobrado_mes))
    c4.metric("A cobrar", fmt_money(pendiente_cobro))
    c5.metric("Resultado mes", fmt_money(resultado_mes))
    c6, c7, c8, c9, c10 = st.columns(5)
    c6.metric("A pagar", fmt_money(a_pagar))
    c7.metric("Deuda total", fmt_money(deuda_total))
    c8.metric("Vencidos / críticos", vencidos)
    c9.metric("Tareas pendientes", tareas_pend)
    c10.metric("Promedio por paciente", fmt_money(promedio_facturacion))
    

    # ======================
    
    # RESUMEN VMR
    
    # ======================
    
    def render_resumen_empresa(titulo, empresa):

        mods = {
    
            name: dfs.get(name, pd.DataFrame())
    
            for name, cfg in MODULES.items()
    
            if cfg.get("empresa") == empresa
    
        }
    
        liquidez = 0
    
        facturacion = 0
    
        cobrado = 0
    
        a_cobrar = 0
    
        a_pagar_emp = 0
    
        deuda_emp = 0
    
        vencidos_emp = 0
    
        tareas_emp = 0
    
        pacientes = 0
    
        for name, df in mods.items():
    
            if df.empty:
    
                continue
    
            tipo = MODULES[name].get("tipo", "")
    
            if tipo in ["caja", "banco"]:
    
                liquidez += total_mod(name)
    
            if "valor_pesos" in df.columns:
    
                facturacion += df["valor_pesos"].apply(money).sum()
    
                pacientes += len(df)
    
                if "estado" in df.columns:
    
                    cobrado += df[df["estado"].astype(str).str.lower().isin(["completo", "cobrado", "pagado"])]["valor_pesos"].apply(money).sum()
    
                    a_cobrar += df[df["estado"].astype(str).str.lower().isin(["pendiente", "parcial", "vencido"])]["valor_pesos"].apply(money).sum()
    
            if "monto" in df.columns and "estado" in df.columns:
    
                a_pagar_emp += df[df["estado"].astype(str).str.lower().isin(["pendiente", "vencido"])]["monto"].apply(money).sum()
    
            if "vencimiento" in df.columns:
    
                vencidos_emp += len(df)
    
            if name == "Tareas Pendientes" and "estado" in df.columns:
    
                tareas_emp += len(df[~df["estado"].astype(str).str.lower().isin(["finalizada", "cancelada"])])
    
        resultado = cobrado - a_pagar_emp
    
        promedio = facturacion / pacientes if pacientes > 0 else 0
    
        st.divider()
    
        st.markdown(f"### {titulo}")
    
        r1, r2, r3, r4, r5 = st.columns(5)
    
        r1.metric("Liquidez actual", fmt_money(liquidez))
    
        r2.metric("Facturación mes", fmt_money(facturacion))
    
        r3.metric("Cobrado mes", fmt_money(cobrado))
    
        r4.metric("A cobrar", fmt_money(a_cobrar))
    
        r5.metric("Resultado mes", fmt_money(resultado))
    
        r6, r7, r8, r9, r10 = st.columns(5)
    
        r6.metric("A pagar", fmt_money(a_pagar_emp))
    
        r7.metric("Deuda total", fmt_money(deuda_emp))
    
        r8.metric("Vencidos / críticos", vencidos_emp)
    
        r9.metric("Tareas pendientes", tareas_emp)
    
        r10.metric("Promedio por paciente", fmt_money(promedio))
    
    render_resumen_empresa("Resumen VMR", "VMR")
    
    render_resumen_empresa("Resumen VM", "VM")
    

    st.divider()
    o1, o2, o3 = st.columns(3)
    o1.metric("Pacientes / procedimientos del mes", pacientes_mes)
    o2.metric("Médicos activos del mes", len(medicos_activos))
    o3.metric("Pendiente neto de cobro", fmt_money(pendiente_cobro))
    st.divider()
    st.markdown("### Facturación por módulo")
    rows = []
    for name, df in dfs.items():
        if df.empty:
            continue
        if "valor_pesos" in df.columns:
            total = df["valor_pesos"].apply(money).sum()
        elif "importe" in df.columns:
            total = df["importe"].apply(money).sum()
        else:
           total = 0
        if total > 0:
            rows.append({
                "Módulo": name,
                "Empresa": MODULES[name]["empresa"],
                "Total": total,
                "Registros": len(df),
            })
    resumen = pd.DataFrame(rows)
    if resumen.empty:
        st.info("Todavía no hay datos económicos cargados.")
    else:
        st.dataframe(resumen, use_container_width=True, hide_index=True)
        fig = px.bar(
            resumen,
            x="Módulo",
            y="Total",
            color="Empresa",
            title="Importes registrados por módulo"
        )
        fig.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(fig, use_container_width=True)
    st.divider()
    st.markdown("### Próximos vencimientos")
    venc_rows = []
    for name, df in dfs.items():
        if df.empty or "vencimiento" not in df.columns:
            continue
        temp = df.copy()
        temp["vencimiento_dt"] = pd.to_datetime(temp["vencimiento"], errors="coerce")
        temp = temp[temp["vencimiento_dt"].notna()]
        temp = temp[temp["vencimiento_dt"] >= hoy]
        for _, row in temp.iterrows():
            venc_rows.append({
                "Módulo": name,
                "Vencimiento": row.get("vencimiento_dt").strftime(DATE_FMT),
                "Detalle": row.get("concepto") or row.get("detalle") or row.get("tarea") or row.get("acreedor") or row.get("afiliado") or row.get("procedimiento") or "",
                "Obra Social": row.get("obra_social") or row.get("obra social") or row.get("os") or "",
                "Importe": row.get("importe") or row.get("saldo") or row.get("valor") or row.get("valor_pesos") or 0,
                "Estado": row.get("estado", ""),
            })
    venc_df = pd.DataFrame(venc_rows)
    if venc_df.empty:
        st.success("No hay vencimientos cargados para los próximos 30 días.")
    else:
        st.dataframe(venc_df.sort_values("Vencimiento"), use_container_width=True, hide_index=True)
        st.divider()

        a_cobrar_vmr = 0
    
        a_cobrar_vm = 0
    
        for _, row in venc_df.iterrows():
    
            importe = money(row["Importe"])
    
            if "VMR" in str(row["Módulo"]):
    
                a_cobrar_vmr += importe
    
            elif "VM" in str(row["Módulo"]):
    
                a_cobrar_vm += importe
    
        c1, c2, c3 = st.columns(3)
    
        c1.metric("💰 A cobrar VMR", fmt_money(a_cobrar_vmr))
    
        c2.metric("💰 A cobrar VM", fmt_money(a_cobrar_vm))
    
        c3.metric("💰 Total a cobrar", fmt_money(a_cobrar_vmr + a_cobrar_vm))
def render_agenda_quirofano(module_name: str, cfg: dict) -> None:

    table = cfg["table"]

    render_header()

    st.header(module_name)

    st.caption(cfg.get("descripcion", ""))

    tab_agenda, tab_cargar, tab_registros = st.tabs([

        "📅 Agenda",

        "➕ Cargar cirugía",

        "📋 Registros"

    ])

    df = get_df(table)

    if not df.empty:

        df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")

        df["hora_inicio_dt"] = pd.to_datetime(

            df["fecha"].dt.strftime("%Y-%m-%d") + " " + df["hora_inicio"].astype(str),

            errors="coerce"

        )

        df["hora_fin_dt"] = pd.to_datetime(

            df["fecha"].dt.strftime("%Y-%m-%d") + " " + df["hora_fin"].astype(str),

            errors="coerce"

        )

    with tab_agenda:

        st.subheader("Agenda quirúrgica")

        c1, c2, c3 = st.columns(3)

        vista = c1.selectbox(

            "Vista",

            ["Día", "Semana", "Mes"],

            key="vista_agenda_qx"

        )

        fecha_sel = c2.date_input(

            "Fecha",

            value=date.today(),

            key="fecha_agenda_qx"

        )

        sala_sel = c3.selectbox(

            "Sala",

            ["Todas", "Quirófano 1"],

            key="sala_agenda_qx"

        )

        agenda = df.copy()

        if not agenda.empty:

            if vista == "Día":

                agenda = agenda[agenda["fecha"].dt.date == fecha_sel]

            elif vista == "Semana":

                inicio_semana = fecha_sel - timedelta(days=fecha_sel.weekday())

                fin_semana = inicio_semana + timedelta(days=6)

                agenda = agenda[

                    (agenda["fecha"].dt.date >= inicio_semana) &

                    (agenda["fecha"].dt.date <= fin_semana)

                ]

            elif vista == "Mes":

                agenda = agenda[

                    (agenda["fecha"].dt.month == fecha_sel.month) &

                    (agenda["fecha"].dt.year == fecha_sel.year)

                ]

            if sala_sel != "Todas" and "sala" in agenda.columns:

                agenda = agenda[agenda["sala"] == sala_sel]

        total = len(agenda)

        programadas = len(agenda[agenda["estado"].astype(str).str.lower() == "programada"]) if not agenda.empty and "estado" in agenda.columns else 0

        finalizadas = len(agenda[agenda["estado"].astype(str).str.lower() == "finalizada"]) if not agenda.empty and "estado" in agenda.columns else 0

        canceladas = len(agenda[agenda["estado"].astype(str).str.lower().isin(["cancelada", "suspendida"])]) if not agenda.empty and "estado" in agenda.columns else 0

        ocupacion_min = 0

        if not agenda.empty and "duracion_min" in agenda.columns:

            ocupacion_min = agenda["duracion_min"].apply(money).sum()

        ocupacion_pct = min((ocupacion_min / (13 * 60)) * 100, 100)

        k1, k2, k3, k4, k5 = st.columns(5)

        k1.metric("Cirugías", total)

        k2.metric("Programadas", programadas)

        k3.metric("Finalizadas", finalizadas)

        k4.metric("Canceladas/Suspendidas", canceladas)

        k5.metric("Ocupación estimada", f"{ocupacion_pct:.1f}%")

        st.divider()

        if agenda.empty:

            st.info("No hay cirugías cargadas para esta vista.")

        else:

            agenda = agenda.sort_values(by=["fecha", "hora_inicio_dt"], ascending=True)

            st.markdown("### Cronograma")

            for _, row in agenda.iterrows():

                estado = str(row.get("estado", "Programada"))

                paciente = row.get("paciente", "")

                proc = row.get("procedimiento", "")

                medico = row.get("medico", "")

                anest = row.get("anestesista", "")

                sala = row.get("sala", "")

                hi = row.get("hora_inicio", "")

                hf = row.get("hora_fin", "")

                fecha_txt = row["fecha"].strftime("%d/%m/%Y") if pd.notna(row["fecha"]) else ""

                if estado.lower() == "finalizada":

                    color = "#DCFCE7"

                elif estado.lower() in ["cancelada", "suspendida"]:

                    color = "#FEE2E2"

                elif estado.lower() == "en curso":

                    color = "#DBEAFE"

                else:

                    color = "#FEF9C3"

                with st.container(border=True):

                                st.markdown(f"**{fecha_txt} | {hi} - {hf}**")
                            
                                st.markdown(f"### {proc}")
                            
                                st.write(f"Paciente: {paciente}")
                            
                                st.write(f"Médico: {medico}")
                            
                                st.write(f"Anestesista: {anest}")
                            
                                st.write(f"Sala: {sala}")
                            
                                st.write(f"Estado: {estado}")
                            
                st.markdown("### Tabla agenda")
                        
                                    
                st.dataframe(
                    
                                    
                                agenda.drop(columns=["hora_inicio_dt", "hora_fin_dt"], errors="ignore"),
                    
                                use_container_width=True,
                    
                                hide_index=True
                    
                            )
                    
                with st.form(f"form_qx_{datetime.now().timestamp()}", clear_on_submit=True):

                        fecha = st.date_input("Fecha", key=f"qx_fecha_{datetime.now().timestamp()}")
                
                        hora_inicio = st.text_input("Hora inicio", key=f"qx_hi_{datetime.now().timestamp()}")
                
                        hora_fin = st.text_input("Hora fin", key=f"qx_hf_{datetime.now().timestamp()}")
                
                        duracion_min = st.number_input("Duración min", value=0.0, key=f"qx_dur_{datetime.now().timestamp()}")
                
                        sala = st.selectbox("Sala", ["Quirófano 1"], key=f"qx_sala_{datetime.now().timestamp()}")
                
                        paciente = st.text_input("Paciente", key=f"qx_paciente_{datetime.now().timestamp()}")
                
                        procedimiento = st.text_input("Procedimiento", key=f"qx_proc_{datetime.now().timestamp()}")
                
                        medico = st.text_input("Médico", key=f"qx_medico_{datetime.now().timestamp()}")
                
                            
                        anestesista = st.text_input("Anestesista", key=f"qx_anest_{datetime.now().timestamp()}")
                
                        estado = st.selectbox("Estado", ["Programada", "En curso", "Finalizada", "Suspendida", "Cancelada"], key=f"qx_estado_{datetime.now().timestamp()}")
                
                        observaciones = st.text_area("Observaciones", key=f"qx_obs_{datetime.now().timestamp()}")
                
                        submitted = st.form_submit_button("Guardar cirugía", type="primary")
                
                        if submitted:
                
                            data = {
                
                                "fecha": fecha,
                
                                "hora_inicio": hora_inicio,
                
                                "hora_fin": hora_fin,
                
                                "duracion_min": duracion_min,
                
                                "sala": sala,
                 
                                "paciente": paciente,
                
                                "procedimiento": procedimiento,
                
                                "medico": medico,
                
                                "anestesista": anestesista,
                
                                "estado": estado,
                
                                "observaciones": observaciones,
                
                            }
                
                            insert_row(table, data)
                
                            st.success("Cirugía guardada correctamente.")
                
                            st.rerun()
                
                with tab_registros:
                
                        st.subheader("Registros cargados")
                
                        if df.empty:
                
                            st.warning("No hay registros cargados.")
                
                        else:
                
                            st.dataframe(
                
                                df.drop(columns=["hora_inicio_dt", "hora_fin_dt"], errors="ignore"),
                
                                use_container_width=True,
                
                                hide_index=True
                
                            )     
                                                          
def render_module(module_name: str) -> None:
        cfg = MODULES[module_name]
                    
                
        if cfg.get("tipo") == "quirófano":
                    
            render_agenda_quirofano(module_name, cfg)
                    
            return
        if module_name in ["Facturación VMR", "Facturación VM"]:
            render_facturacion_pro(module_name, cfg)
                            
            return
        table = cfg["table"]
        render_header()
        st.header(module_name)
        st.caption(cfg["descripcion"])
                
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["➕ Cargar", "📥 Importar planilla", "📋 Registros", "✏️ Editar / Eliminar", "📤 Exportar"])
                
        with tab1:
            st.subheader("Nuevo registro")
            with st.form(f"form_add_{table}", clear_on_submit=True):
                data: Dict[str, Any] = {}
                    
                cols = st.columns(2)
                    
                for i, field in enumerate(cfg["fields"]):
                    with cols[i % 2]:
                            
                        raw = input_field(field, f"add_{table}")
                                    
                        data[field[0]] = clean_for_db(raw, field[1])
                submitted = st.form_submit_button("Guardar registro", type="primary")
                if submitted:
                    errors = validate_required(cfg, data)
                    if errors:
                        st.error("Faltan completar campos obligatorios: " + ", ".join(errors))
                    else:
                        insert_row(table, data)
                        st.success("Registro guardado correctamente.")
                                    
                        st.rerun()

            with tab2:
                render_importer(module_name, cfg)

            with tab3:
                df = add_balance_columns(get_df(table))
                if df.empty:
                    st.warning("Todavía no hay registros cargados en este módulo.")
                else:
                    filtered = apply_filters(df, module_name)
                    st.subheader("Indicadores del módulo")
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Registros visibles", len(filtered))
                    if "ingreso" in filtered.columns:
                        m2.metric("Ingresos", fmt_money(filtered["ingreso"].apply(money).sum()))
                    if "egreso" in filtered.columns:
                        m3.metric("Egresos", fmt_money(filtered["egreso"].apply(money).sum()))
                    if "saldo_movimiento" in filtered.columns:
                        m4.metric("Saldo", fmt_money(filtered["saldo_movimiento"].apply(money).sum()))
                    elif "saldo" in filtered.columns:
                        m4.metric("Saldo", fmt_money(filtered["saldo"].apply(money).sum()))
                    elif "importe" in filtered.columns:
                        m4.metric("Total importe", fmt_money(filtered["importe"].apply(money).sum()))
                    elif "valor_pesos" in filtered.columns:
                        m4.metric("Total facturado", fmt_money(filtered["valor_pesos"].apply(money).sum()))
        
                    show = filtered.copy()
        
                    cols = [
        
                        "afiliado",
        
                        "obra_social",
        
                        "procedimiento",
        
                        "medico_responsable",
        
                        "fecha_factura",
        
                        "valor_pesos",
        
                        "estado"
        
                    ]
        
                    show = show[[c for c in cols if c in show.columns]]
        
                    show = show.rename(columns={
        
                        "afiliado": "Paciente",
        
                        "obra_social": "Obra Social",
        
                        "procedimiento": "Procedimiento",
        
                        "medico_responsable": "Médico",
        
                        "fecha_factura": "Fecha Factura",
        
                        "valor_pesos": "Valor Facturado",
        
                        "estado": "Estado"
        
                    })
        
                    if "Valor Facturado" in show.columns:
        
                        show["Valor Facturado"] = show["Valor Facturado"].apply(fmt_money)
        
                    st.dataframe(
        
                        show,
        
                        use_container_width=True,
        
                        hide_index=True
        
                    )
        
                    fecha_col = first_available_date_col(filtered, module_name)
                    if fecha_col and not filtered.empty:
                        graph = filtered.copy()
                        graph[fecha_col] = pd.to_datetime(graph[fecha_col], errors="coerce")
                        graph = graph[graph[fecha_col].notna()]
                        y_col = None
                        if "saldo_movimiento" in graph.columns:
                            y_col = "saldo_movimiento"
                        elif "saldo" in graph.columns:
                            y_col = "saldo"
                        elif "importe" in graph.columns:
                            y_col = "importe"
                        elif "valor_pesos" in graph.columns:
                            y_col = "valor_pesos"
                        if y_col and not graph.empty:
                            chart = graph.groupby(graph[fecha_col].dt.date)[y_col].sum().reset_index()
                            fig = px.line(chart, x=fecha_col, y=y_col, markers=True, title=f"Evolución: {module_name}")
                            st.plotly_chart(fig, use_container_width=True)
        
            with tab4:
                df = get_df(table)
                if df.empty:
                    st.warning("No hay registros para editar.")
                else:
                    ids = df["id"].tolist()
                    selected_id = st.selectbox("Seleccionar ID", ids, key=f"edit_select_{table}")
                    row = df[df["id"] == selected_id].iloc[0].to_dict()
                    st.caption(f"Creado: {row.get('created_at', '')} | Última actualización: {row.get('updated_at', '')}")
        
                    with st.form(f"form_edit_{table}"):
                        data: Dict[str, Any] = {}
                        cols = st.columns(2)
                        for i, field in enumerate(cfg["fields"]):
                            with cols[i % 2]:
                                raw = input_field(field, f"edit_{table}_{selected_id}", row)
                                data[field[0]] = clean_for_db(raw, field[1])
                        save = st.form_submit_button("Guardar cambios", type="primary")
                        if save:
                            errors = validate_required(cfg, data)
                            if errors:
                                st.error("Faltan completar campos obligatorios: " + ", ".join(errors))
                            else:
                                update_row(table, int(selected_id), data)
                                st.success("Registro actualizado correctamente.")
                                st.rerun()
        
                    st.warning("Zona de eliminación")
                    confirm = st.checkbox("Confirmo que quiero eliminar este registro", key=f"confirm_delete_{table}_{selected_id}")
                    if st.button("Eliminar registro", disabled=not confirm, type="secondary"):
                        delete_row(table, int(selected_id))
                        st.success("Registro eliminado.")
                        st.rerun()
        
            with tab5:
                df = add_balance_columns(get_df(table))
                if df.empty:
                    st.info("No hay datos para exportar.")
                else:
                    incluir_tecnicas = st.checkbox("Incluir columnas técnicas id / created_at / updated_at", value=False, key=f"export_tech_{table}")
                    export_df = df if incluir_tecnicas else module_business_df(df, cfg)
                    csv = export_df.to_csv(index=False).encode("utf-8-sig")
                    st.download_button("Descargar CSV", data=csv, file_name=f"{table}.csv", mime="text/csv")
                    xlsx_path = Path(f"{table}.xlsx")
                    export_df.to_excel(xlsx_path, index=False)
                    with open(xlsx_path, "rb") as f:
                        st.download_button(
                            "Descargar Excel",
                            data=f,
                            file_name=f"{table}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

def render_admin() -> None:
    render_header()
    st.header("⚙️ Administración")
    st.write("Herramientas de mantenimiento local.")

    st.subheader("Base de datos")
    st.code(str(DB_PATH.resolve()))

    if st.button("Inicializar / reparar tablas"):
        init_db()
        st.success("Tablas verificadas correctamente.")

    st.subheader("Carga de datos de ejemplo")
    if st.button("Crear ejemplos mínimos"):
        seed_examples()
        st.success("Datos de ejemplo creados.")
        st.rerun()

    st.subheader("Exportación global")
    all_data = {}
    for name, cfg in MODULES.items():
        df = get_df(cfg["table"])
        if not df.empty:
            all_data[name[:31]] = module_business_df(add_balance_columns(df), cfg)
    if all_data:
        export_path = Path("vitae_export_global.xlsx")
        with pd.ExcelWriter(export_path, engine="openpyxl") as writer:
            for sheet, df in all_data.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
        with open(export_path, "rb") as f:
            st.download_button(
                "Descargar Excel global",
                data=f,
                file_name="vitae_export_global.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.info("Todavía no hay datos para exportación global.")

    st.subheader("Borrado total")
    st.error("Esta acción elimina todos los registros de todos los módulos.")
    password = st.text_input(



    "Contraseña de administrador",



    type="password"



    )



    erase = st.checkbox(



        "Confirmo borrado total de la base local"



    )



    if st.button("Borrar todos los datos"):



        if password != "Vitae2026!":



            st.error("Contraseña incorrecta")



            st.stop()



        if not erase:



            st.error("Debes confirmar el borrado")



            st.stop()



        with connect() as conn:



            for cfg in MODULES.values():



                conn.execute(f"DELETE FROM {cfg['table']}")



            conn.commit()



        st.success("Base vaciada.")



        st.rerun()

def render_configuracion() -> None:

    render_header()

    st.header("Configuración")

    tab1, tab2, tab3, tab4 = st.tabs([

        "👤 Usuarios",

        "🔐 Permisos",

        "🏢 Empresas",

        "⚙️ Sistema"

    ])

    with tab1:

        st.subheader("Usuarios")

        st.info("Acá irá la gestión de usuarios.")

    with tab2:

        st.subheader("Permisos")

        st.info("Acá irá la gestión de permisos.")

    with tab3:

        st.subheader("Empresas")

        st.info("Acá irá la gestión de empresas.")

    with tab4:

        st.subheader("Sistema")

        st.info("Acá irá la configuración general del sistema.")

def seed_examples() -> None:
    examples = [
        ("caja_vmr", {"fecha": date.today().strftime(DATE_FMT), "concepto": "Ingreso muestra fertilidad", "categoria": "Ingreso", "medio": "Efectivo", "ingreso": 150000, "egreso": 0, "responsable": "Administración", "observaciones": "Ejemplo"}),
        ("banco_galicia_vm", {"fecha": date.today().strftime(DATE_FMT), "concepto": "Pago proveedor quirófano", "tipo_movimiento": "Débito", "referencia": "OP-001", "ingreso": 0, "egreso": 80000, "conciliado": 1, "observaciones": "Ejemplo"}),
        ("pagos_pendientes_vitae", {"fecha": date.today().strftime(DATE_FMT), "empresa": "VITAE", "proveedor": "Proveedor insumos", "concepto": "Insumos médicos", "importe": 120000, "pagado": 0, "vencimiento": (date.today() + timedelta(days=7)).strftime(DATE_FMT), "prioridad": "Alta", "estado": "Pendiente", "observaciones": "Ejemplo"}),
        ("tareas_pendientes", {"fecha": date.today().strftime(DATE_FMT), "empresa": "VM", "tarea": "Revisar stock quirófano", "responsable": "Enfermería", "prioridad": "Alta", "vencimiento": (date.today() + timedelta(days=3)).strftime(DATE_FMT), "estado": "Pendiente", "observaciones": "Ejemplo"}),
    ]
    for table, data in examples:
        insert_row(table, data)

# =========================================================
# APP
# =========================================================

def main() -> None:
    init_db()

    st.sidebar.title("VITAE")
    st.sidebar.caption("Sistema interno de gestión")

    page = st.sidebar.radio("Navegación", ["Dashboard Global", "Módulos", "Administración", "Configuración"])

    if page == "Dashboard Global":
    
        render_dashboard()
    
    elif page == "Módulos":
    
        empresas = ["Todos", "VMR", "VM", "VITAE"]
    
        empresa_filter = st.sidebar.selectbox("Empresa", empresas)
    
        module_names = list(MODULES.keys())
    
        if empresa_filter != "Todos":
    
            module_names = [
    
                m for m in module_names
    
                if MODULES[m]["empresa"] == empresa_filter or MODULES[m]["empresa"] == "VITAE"
    
            ]
    
        module_name = st.sidebar.selectbox("Módulo", module_names)
    
        render_module(module_name)
    
    elif page == "Administración":
    
        render_admin()
    
    elif page == "Configuración":
    
        render_configuracion()

    st.sidebar.divider()
    st.sidebar.markdown("**Módulos incluidos**")
    st.sidebar.caption(f"{len(MODULES)} módulos activos")

if __name__ == "__main__":
    main()
