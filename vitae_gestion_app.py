# vitae_gestion_app.py
# Ejecutar en VS Code / terminal:
#   pip install streamlit pandas plotly openpyxl
#   streamlit run vitae_gestion_app.py

from __future__ import annotations

import sqlite3
from pathlib import Path
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st

# =========================================================
# CONFIG GENERAL
# =========================================================

APP_TITLE = "VITAE | Sistema Integral de Gestión | IMPORTADOR OK"
DB_PATH = Path("vitae_gestion.db")
DATE_FMT = "%Y-%m-%d"

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
    """Crea tablas y agrega automáticamente columnas nuevas si el módulo cambió.

    Esto permite actualizar Facturación VMR/VM sin tener que borrar la base cada vez.
    SQLite no elimina columnas viejas automáticamente, pero la app muestra/exporta solo las columnas del módulo actual.
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
        conn.commit()


def get_df(table: str) -> pd.DataFrame:
    with connect() as conn:
        try:
            df = pd.read_sql_query(f"SELECT * FROM {table} ORDER BY id DESC", conn)
        except Exception:
            df = pd.DataFrame()
    return df


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

def money(value: Any) -> float:
    try:
        if pd.isna(value) or value == "":
            return 0.0
        if isinstance(value, str):
            value = normalize_money_string(value)
        return float(value)
    except Exception:
        return 0.0


def fmt_money(value: Any) -> str:
    return f"$ {money(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


TECH_COLUMNS = ["id", "created_at", "updated_at"]


def business_df(df: pd.DataFrame) -> pd.DataFrame:
    """Oculta columnas técnicas para que el usuario vea solo la gestión real."""
    if df is None or df.empty:
        return df
    cols = [c for c in df.columns if c not in TECH_COLUMNS]
    return df[cols].copy()


def module_business_df(df: pd.DataFrame, cfg: Dict[str, Any]) -> pd.DataFrame:
    """Muestra solo las columnas vigentes del módulo y cálculos útiles.

    Esto evita que queden visibles columnas viejas si el esquema cambió, por ejemplo
    cliente/importe/cobrado antiguos de Facturación.
    """
    if df is None or df.empty:
        return df
    field_names = get_field_names(cfg)
    calc_cols = ["saldo", "saldo_movimiento"]
    cols = [c for c in field_names + calc_cols if c in df.columns]
    return df[cols].copy()


def show_module_table(df: pd.DataFrame, cfg: Dict[str, Any], **kwargs: Any) -> None:
    st.dataframe(module_business_df(df, cfg), use_container_width=True, hide_index=True, **kwargs)


def show_business_table(df: pd.DataFrame, **kwargs: Any) -> None:
    st.dataframe(business_df(df), use_container_width=True, hide_index=True, **kwargs)


def get_field_names(cfg: Dict[str, Any]) -> List[str]:
    return [field[0] for field in cfg["fields"]]


def parse_date(value: Any) -> date | None:
    if value in (None, "", pd.NaT):
        return None
    try:
        return pd.to_datetime(value, dayfirst=True, errors="coerce").date()
    except Exception:
        return None


def clean_for_db(value: Any, ftype: str) -> Any:
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
        if old in options:
            idx = options.index(old)
        return st.selectbox(label, options, index=idx, key=key)
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


def apply_filters(df: pd.DataFrame, module_name: str) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    st.subheader("Filtros")
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        search = st.text_input("Buscar texto", key=f"search_{module_name}")
    with c2:
        estado = "Todos"
        if "estado" in df.columns:
            estados = ["Todos"] + sorted([x for x in df["estado"].dropna().unique().tolist() if x != ""])
            estado = st.selectbox("Estado", estados, key=f"estado_{module_name}")
    with c3:
        fecha_desde = st.date_input("Desde", value=date.today() - timedelta(days=365), key=f"desde_{module_name}")
    with c4:
        fecha_hasta = st.date_input("Hasta", value=date.today() + timedelta(days=365), key=f"hasta_{module_name}")

    if search:
        mask = df.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)
        df = df[mask]
    if "estado" in df.columns and estado != "Todos":
        df = df[df["estado"] == estado]
    if "fecha" in df.columns:
        f = pd.to_datetime(df["fecha"], errors="coerce").dt.date
        df = df[(f >= fecha_desde) & (f <= fecha_hasta)]
    return df

# =========================================================
# IMPORTADOR PROFESIONAL EXCEL / CSV
# =========================================================

def normalize_money_string(value: Any) -> str:
    """Convierte formatos argentinos $ 1.234.567,89 o 1234567.89 a string numérico válido."""
    if value is None or pd.isna(value):
        return "0"
    text = str(value).strip()
    if text == "":
        return "0"
    text = text.replace("$", "").replace("ARS", "").replace("USD", "")
    text = text.replace(" ", "").replace("\u00a0", "")
    text = text.replace("(", "-").replace(")", "")

    if "," in text and "." in text:
        # Formato AR: 1.234,56
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        # Formato US: 1,234.56
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(".", "").replace(",", ".")
    return text


def clean_tabular_sheet(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Detecta encabezados aunque la planilla tenga títulos arriba o filas combinadas.

    Ejemplo: una hoja con título "CONTROL PROCEDIMIENTOS..." y encabezados en la fila 4
    se convierte automáticamente en una tabla con columnas MES, AFILIADO, OBRA SOCIAL, etc.
    """
    if df_raw.empty:
        return df_raw

    raw = df_raw.copy().dropna(how="all").dropna(axis=1, how="all")
    if raw.empty:
        return raw

    header_keywords = {
        "mes", "afiliado", "obra social", "procedimiento", "medico", "médico",
        "fecha factura", "factura", "vencimiento", "fecha pago", "valor", "estado",
        "cliente", "paciente", "importe", "concepto"
    }

    best_idx = raw.index[0]
    best_score = -1
    for idx, row in raw.iterrows():
        values = [str(x).strip().lower() for x in row.tolist() if pd.notna(x) and str(x).strip() != ""]
        if not values:
            continue
        joined_values = " | ".join(values)
        score = sum(1 for kw in header_keywords if kw in joined_values)
        score += min(len(values), 8) * 0.05
        if score > best_score:
            best_score = score
            best_idx = idx

    header_values = raw.loc[best_idx].tolist()
    columns = []
    used: Dict[str, int] = {}
    for i, value in enumerate(header_values):
        name = str(value).strip() if pd.notna(value) and str(value).strip() else f"Columna_{i+1}"
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
    """Devuelve un diccionario hoja -> dataframe. Soporta CSV y Excel con múltiples hojas."""
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
    name, ftype, required = field[0], field[1], field[2]
    req = " *" if required else ""
    return f"{name.replace('_', ' ').title()}{req}"


def auto_guess_column(target_name: str, source_columns: List[str]) -> str:
    norm_target = target_name.lower().replace("_", " ")
    aliases = {
        "fecha": ["fecha", "dia", "día", "date"],
        "mes": ["mes", "periodo", "período"],
        "afiliado": ["afiliado", "paciente", "cliente", "nombre", "apellido y nombre"],
        "medico_responsable": ["medico responsable", "médico responsable", "medico", "médico", "doctor", "profesional", "responsable"],
        "fecha_factura": ["fecha factura", "fecha de factura", "fecha", "factura fecha"],
        "numero_factura": ["n° factura", "nº factura", "n factura", "numero factura", "número factura", "factura", "comprobante"],
        "fecha_pago": ["fecha pago", "fecha de pago", "pago fecha"],
        "valor_pesos": ["valor $", "valor pesos", "valor ars", "importe", "monto", "total", "valor"],
        "valor_usd": ["valor usd", "usd", "dolares", "dólares"],
        "concepto": ["concepto", "detalle", "descripcion", "descripción", "movimiento", "observacion"],
        "detalle": ["detalle", "concepto", "descripcion", "descripción"],
        "descripcion": ["descripcion", "descripción", "detalle", "concepto"],
        "cliente": ["cliente", "paciente", "nombre", "razon social", "razón social"],
        "paciente": ["paciente", "cliente", "nombre"],
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
        "obra_social": ["obra social", "os", "prepaga"],
        "practica": ["practica", "práctica", "prestacion", "prestación", "procedimiento"],
        "procedimiento": ["procedimiento", "practica", "práctica", "prestacion", "prestación"],
        "periodo": ["periodo", "período", "mes"],
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
    if value is None or pd.isna(value) or str(value).strip() == "":
        return options[0] if options else ""
    text = str(value).strip()
    for opt in options:
        if text.lower() == opt.lower():
            return opt
    text_low = text.lower()
    aliases = {
        "si": "Sí", "sí": "Sí", "true": "Sí", "no": "No", "false": "No",
        "cobrado": "Cobrado", "pagado": "Pagado", "pendiente": "Pendiente", "vencido": "Vencido",
        "parcial": "Parcial", "completo": "Completo", "completa": "Completo", "activo": "Activo", "finalizado": "Finalizado",
        "alta": "Alta", "media": "Media", "baja": "Baja",
        "credito": "Crédito", "crédito": "Crédito", "debito": "Débito", "débito": "Débito",
    }
    wanted = aliases.get(text_low)
    if wanted and wanted in options:
        return wanted
    return options[0] if options else text


def clean_import_value(value: Any, field: Tuple) -> Any:
    name, ftype = field[0], field[1]
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
        if value is None or pd.isna(value):
            return 0
        return 1 if str(value).strip().lower() in ["1", "true", "si", "sí", "x", "ok", "pagado", "conciliado"] else 0
    if ftype == "select":
        return normalize_select_value(value, options or [])
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def render_importer(module_name: str, cfg: Dict[str, Any]) -> None:
    table = cfg["table"]
    st.subheader("Importar planilla Excel / CSV")
    st.caption("Subí una planilla, elegí la hoja, mapeá las columnas y guardala dentro de este módulo. No hace falta que todas tus planillas tengan el mismo formato.")

    uploaded_file = st.file_uploader(
        "Subir archivo",
        type=["xlsx", "xls", "csv"],
        key=f"upload_{table}",
    )

    if uploaded_file is None:
        st.info("Acepta Excel con varias hojas o CSV. Primero subí el archivo y después asignás cada columna.")
        return

    try:
        sheets = read_uploaded_sheet(uploaded_file)
    except Exception as e:
        st.error(f"No pude leer el archivo. Revisá que sea Excel/CSV válido. Detalle: {e}")
        return

    sheet_names = list(sheets.keys())
    selected_sheet = st.selectbox("Hoja a importar", sheet_names, key=f"sheet_{table}")
    df_original = sheets[selected_sheet].copy()
    df_original = df_original.dropna(how="all")
    df_original.columns = [str(c).strip() for c in df_original.columns]

    if df_original.empty:
        st.warning("La hoja seleccionada está vacía.")
        return

    st.markdown("#### Vista previa")
    show_business_table(df_original.head(30))

    columnas = df_original.columns.tolist()
    st.markdown("#### Mapeo de columnas")
    st.caption("La app intenta detectar columnas automáticamente. Corregí lo que haga falta. Los campos con * son recomendados para que el módulo quede bien cargado.")

    mapping: Dict[str, str] = {}
    cols = st.columns(2)
    for i, field in enumerate(cfg["fields"]):
        name = field[0]
        guessed = auto_guess_column(name, columnas)
        options = ["No usar"] + columnas
        index = options.index(guessed) if guessed in options else 0
        with cols[i % 2]:
            mapping[name] = st.selectbox(
                field_label(field),
                options,
                index=index,
                key=f"map_{table}_{name}",
            )

    with st.expander("Opciones avanzadas"):
        modo = st.radio(
            "Modo de importación",
            ["Agregar a registros existentes", "Reemplazar módulo completo"],
            key=f"modo_import_{table}",
        )
        saltar_filas_vacias = st.checkbox("Saltar filas completamente vacías", value=True, key=f"skip_empty_{table}")
        validar_obligatorios = st.checkbox("Validar campos obligatorios", value=False, key=f"valid_required_{table}")
        st.caption("Si activás validar obligatorios, la app no importará filas que no tengan los campos obligatorios del módulo.")

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
        missing_core = []
        for core in ["importe", "cobrado", "cliente", "concepto", "fecha"]:
            if core in mapping and mapping.get(core) == "No usar":
                missing_core.append(core)
        if missing_core:
            st.warning("Atención: estos campos importantes quedaron en 'No usar': " + ", ".join(missing_core) + ". Si no los mapeás, entrarán vacíos o en cero.")

    if rejected_rows:
        with st.expander(f"Filas rechazadas: {len(rejected_rows)}"):
            show_business_table(pd.DataFrame(rejected_rows))

    col_a, col_b = st.columns([1, 2])
    with col_a:
        confirm_import = st.checkbox("Confirmo la importación", key=f"confirm_import_{table}")
    with col_b:
        st.caption("Revisá la previsualización antes de confirmar. Si reemplazás el módulo completo, se borran los registros anteriores de este módulo.")

    if st.button("Importar planilla al módulo", type="primary", disabled=(not confirm_import or not rows), key=f"btn_import_{table}"):
        if modo == "Reemplazar módulo completo":
            count = replace_table_rows(table, rows)
        else:
            count = bulk_insert_rows(table, rows)
        st.success(f"Importación completada. Registros importados en {module_name}: {count}")
        st.rerun()

# =========================================================
# VISTAS
# =========================================================

def render_header() -> None:
    st.markdown('<div class="main-title">🏥 VITAE | Sistema Integral de Gestión</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">VMR · Vitae Medicina Reproductiva | VM · Vitae Medical</div>', unsafe_allow_html=True)


def render_dashboard() -> None:
    render_header()
    st.info("Tablero global calculado automáticamente desde todos los módulos cargados.")

    dfs = {name: add_balance_columns(get_df(cfg["table"])) for name, cfg in MODULES.items()}

    ingresos = 0.0
    egresos = 0.0
    deuda = 0.0
    a_cobrar = 0.0
    vencidos = 0
    tareas_pend = 0

    deuda_modules = [
        "Deudas Impositivas VMR", "Deudas Impositivas VM", "Planes de pagos y préstamos",
        "Pagos pendientes Vitae", "Deuda total", "Honorarios médicos"
    ]

    for name, df in dfs.items():
        if df.empty:
            continue
        if "ingreso" in df.columns:
            ingresos += df["ingreso"].apply(money).sum()
        if "egreso" in df.columns:
            egresos += df["egreso"].apply(money).sum()
        if name in ["Facturación VMR", "Facturación VM"] and "valor_pesos" in df.columns:
            total_facturado = df["valor_pesos"].apply(money).sum()
            cobrado_estimado = df[df.get("estado", "").astype(str).str.lower().isin(["completo", "cobrado", "pagado"])]["valor_pesos"].apply(money).sum() if "estado" in df.columns else 0
            ingresos += cobrado_estimado
            a_cobrar += max(0, total_facturado - cobrado_estimado)
        elif "importe" in df.columns and name in ["Gine Vitae"]:
            cob = df["cobrado"].apply(money).sum() if "cobrado" in df.columns else 0
            ingresos += cob
            a_cobrar += max(0, df["importe"].apply(money).sum() - cob)
        if name in deuda_modules:
            if "saldo" in df.columns:
                deuda += df["saldo"].apply(money).sum()
            elif "importe" in df.columns:
                pag = df["pagado"].apply(money).sum() if "pagado" in df.columns else 0
                deuda += max(0, df["importe"].apply(money).sum() - pag)
        if "vencimiento" in df.columns:
            venc = pd.to_datetime(df["vencimiento"], errors="coerce")
            estado = (
                df["estado"].astype(str).str.lower()
                if "estado" in df.columns
                else pd.Series([""] * len(df), index=df.index)
            )
            estados_cerrados = ["pagado", "cobrado", "realizado", "finalizada", "finalizado", "anulado", "cancelado"]
            vencidos += int(((venc < pd.Timestamp.today().normalize()) & (~estado.isin(estados_cerrados))).sum())
        if name == "Tareas Pendientes" and "estado" in df.columns:
            tareas_pend += int(df[~df["estado"].isin(["Finalizada", "Cancelada"])].shape[0])

    saldo = ingresos - egresos
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Ingresos registrados", fmt_money(ingresos))
    c2.metric("Egresos registrados", fmt_money(egresos))
    c3.metric("Saldo operativo", fmt_money(saldo))
    c4.metric("Deuda total estimada", fmt_money(deuda))
    c5.metric("A cobrar", fmt_money(a_cobrar))

    c6, c7, c8 = st.columns(3)
    c6.metric("Vencidos / críticos", vencidos)
    c7.metric("Tareas pendientes", tareas_pend)
    c8.metric("Base de datos", str(DB_PATH))

    st.divider()

    rows = []
    for name, df in dfs.items():
        cfg = MODULES[name]
        total_importe = df["importe"].apply(money).sum() if "importe" in df.columns else (df["valor_pesos"].apply(money).sum() if "valor_pesos" in df.columns else 0)
        total_saldo = df["saldo"].apply(money).sum() if "saldo" in df.columns else 0
        total_ing = df["ingreso"].apply(money).sum() if "ingreso" in df.columns else 0
        total_egr = df["egreso"].apply(money).sum() if "egreso" in df.columns else 0
        rows.append({
            "Módulo": name,
            "Empresa": cfg["empresa"],
            "Registros": len(df),
            "Ingresos": total_ing,
            "Egresos": total_egr,
            "Importes": total_importe,
            "Saldos": total_saldo,
        })
    summary = pd.DataFrame(rows)
    st.subheader("Resumen por módulo")
    show_business_table(summary)

    chart_df = summary[summary["Registros"] > 0]
    if not chart_df.empty:
        fig = px.bar(chart_df, x="Módulo", y="Registros", color="Empresa", title="Registros cargados por módulo")
        fig.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Próximos vencimientos")
    venc_rows = []
    for name, df in dfs.items():
        if df.empty or "vencimiento" not in df.columns:
            continue
        temp = df.copy()
        temp["vencimiento_dt"] = pd.to_datetime(temp["vencimiento"], errors="coerce")
        temp = temp[temp["vencimiento_dt"].notna()]
        hoy_ts = pd.Timestamp.today().normalize()
        limite_ts = hoy_ts + pd.Timedelta(days=30)
        temp = temp[(temp["vencimiento_dt"] >= hoy_ts) & (temp["vencimiento_dt"] <= limite_ts)]
        for _, row in temp.iterrows():
            venc_rows.append({
                "Módulo": name,
                "Vencimiento": row.get("vencimiento_dt").strftime("%Y-%m-%d") if pd.notna(row.get("vencimiento_dt")) else row.get("vencimiento"),
                "Detalle": row.get("concepto") or row.get("detalle") or row.get("tarea") or row.get("acreedor") or row.get("afiliado") or row.get("procedimiento") or "",
                "Importe": row.get("importe") or row.get("saldo") or row.get("valor") or row.get("valor_pesos") or 0,
                "Estado": row.get("estado", ""),
            })
    venc_df = pd.DataFrame(venc_rows)
    if venc_df.empty:
        st.success("No hay vencimientos cargados para los próximos 30 días.")
    else:
        show_business_table(venc_df.sort_values("Vencimiento"))


def render_module(module_name: str) -> None:
    cfg = MODULES[module_name]
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
            m1.metric("Registros", len(filtered))
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

            show_module_table(filtered, cfg)

            fecha_col = "fecha" if "fecha" in filtered.columns else ("fecha_factura" if "fecha_factura" in filtered.columns else None)
            if fecha_col and not filtered.empty:
                graph = filtered.copy()
                graph[fecha_col] = pd.to_datetime(graph[fecha_col], errors="coerce")
                y_col = None
                if "saldo_movimiento" in graph.columns:
                    y_col = "saldo_movimiento"
                elif "saldo" in graph.columns:
                    y_col = "saldo"
                elif "importe" in graph.columns:
                    y_col = "importe"
                elif "valor_pesos" in graph.columns:
                    y_col = "valor_pesos"
                if y_col:
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
            st.download_button(
                "Descargar CSV",
                data=csv,
                file_name=f"{table}.csv",
                mime="text/csv",
            )
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
            all_data[name[:31]] = df
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
    erase = st.checkbox("Confirmo borrado total de la base local")
    if st.button("Borrar todos los datos", disabled=not erase):
        with connect() as conn:
            for cfg in MODULES.values():
                conn.execute(f"DELETE FROM {cfg['table']}")
            conn.commit()
        st.success("Base vaciada.")
        st.rerun()


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

    page = st.sidebar.radio(
        "Navegación",
        ["Dashboard Global", "Módulos", "Administración"],
    )

    if page == "Dashboard Global":
        render_dashboard()
    elif page == "Módulos":
        empresas = ["Todos", "VMR", "VM", "VITAE"]
        empresa_filter = st.sidebar.selectbox("Empresa", empresas)
        module_names = list(MODULES.keys())
        if empresa_filter != "Todos":
            module_names = [m for m in module_names if MODULES[m]["empresa"] == empresa_filter or MODULES[m]["empresa"] == "VITAE"]
        module_name = st.sidebar.selectbox("Módulo", module_names)
        render_module(module_name)
    else:
        render_admin()

    st.sidebar.divider()
    st.sidebar.markdown("**Módulos incluidos**")
    st.sidebar.caption(f"{len(MODULES)} módulos activos")


if __name__ == "__main__":
    main()
