# app_streamlit.py
import io
from datetime import datetime
import pandas as pd
import streamlit as st

from core_consolidacion import MESES_ES  # solo para usar nombres de meses


st.set_page_config(page_title="Data Workbench de Ventas", layout="wide")


# ===================== UTILIDADES =====================

def leer_excels_subidos(uploaded_files):
    """Combina todos los Excels subidos en un único DataFrame."""
    frames = []
    for up in uploaded_files:
        df = pd.read_excel(up, sheet_name=0)
        df["_archivo_origen"] = up.name
        frames.append(df)
    if not frames:
        raise ValueError("No se pudo leer ningún archivo.")
    return pd.concat(frames, ignore_index=True)


def get_schema_mapping(df):
    """UI: deja al usuario mapear qué columna es qué cosa."""
    st.subheader("Mapear columnas (esquema)")

    cols = ["<Ninguna>"] + list(df.columns)

    def select(label, default):
        return st.selectbox(label, options=cols,
                            index=cols.index(default) if default in cols else 0)

    c1, c2, c3 = st.columns(3)
    with c1:
        fecha_col = select("Columna de fecha de comprobante", "Fecha")
        ticket_col = select("Columna de N° de comprobante / ticket", "Ticket")
        cliente_col = select("Columna de cliente", "IdCliente")

    with c2:
        prod_col = select("Columna de IdArticulo / SKU", "IdArticulo")
        desc_col = select("Columna de descripción producto", "Descripcion")
        depto_col = select("Columna de departamento / familia", "Departamento")

    with c3:
        cant_col = select("Columna de cantidad", "Cantidad")
        precio_col = select("Columna de precio unitario", "PrecioUnitario")
        total_col = select("Columna de total de línea / ticket", "Total")

    c4, c5 = st.columns(2)
    with c4:
        sucursal_col = select("Columna de sucursal / unidad de negocio", "Sucursal")
    with c5:
        vendedor_col = select("Columna de vendedor / cajero", "Vendedor")

    schema = {
        "fecha": None if fecha_col == "<Ninguna>" else fecha_col,
        "ticket": None if ticket_col == "<Ninguna>" else ticket_col,
        "cliente": None if cliente_col == "<Ninguna>" else cliente_col,
        "producto": None if prod_col == "<Ninguna>" else prod_col,
        "descripcion": None if desc_col == "<Ninguna>" else desc_col,
        "departamento": None if depto_col == "<Ninguna>" else depto_col,
        "cantidad": None if cant_col == "<Ninguna>" else cant_col,
        "precio": None if precio_col == "<Ninguna>" else precio_col,
        "total": None if total_col == "<Ninguna>" else total_col,
        "sucursal": None if sucursal_col == "<Ninguna>" else sucursal_col,
        "vendedor": None if vendedor_col == "<Ninguna>" else vendedor_col,
    }
    return schema


def get_columnas_adicionales_config():
    """UI para que el usuario seleccione año y columnas mensuales adicionales."""
    st.subheader("Configuración de columnas adicionales")
    
    # Selector de año
    anio_actual = datetime.now().year
    anio_seleccionado = st.selectbox(
        "Seleccionar año para columnas mensuales",
        options=list(range(2024, 2031)),
        index=list(range(2024, 2031)).index(anio_actual) if anio_actual <= 2030 else 0
    )
    
    # Columnas mensuales disponibles
    meses_disponibles = [
        f"ENERO {anio_seleccionado}",
        f"FEBRERO {anio_seleccionado}",
        f"MARZO {anio_seleccionado}",
        f"ABRIL {anio_seleccionado}",
        f"MAYO {anio_seleccionado}",
        f"JUNIO {anio_seleccionado}",
        f"JULIO {anio_seleccionado}",
        f"AGOSTO {anio_seleccionado}",
        f"SEPTIEMBRE {anio_seleccionado}",
    ]
    
    # Columnas de totales
    columnas_totales = [
        "TOTAL HIPER",
        "TOTAL CORRIENTES",
        "TOTAL CONSOLIDADO"
    ]
    
    todas_columnas = meses_disponibles + columnas_totales
    
    columnas_seleccionadas = st.multiselect(
        "Seleccionar columnas adicionales a incluir en exportación",
        options=todas_columnas,
        default=[]
    )
    
    return {
        "anio": anio_seleccionado,
        "columnas": columnas_seleccionadas,
        "meses": [c for c in columnas_seleccionadas if c in meses_disponibles],
        "totales": [c for c in columnas_seleccionadas if c in columnas_totales]
    }


def agregar_columnas_adicionales(df_resultado, config_cols, df_original, schema):
    """Agrega las columnas adicionales seleccionadas al DataFrame de resultados."""
    if not config_cols["columnas"]:
        return df_resultado
    
    df = df_resultado.copy()
    anio = config_cols["anio"]
    
    # Verificar si el df original tiene datos de fecha
    if schema.get("fecha") and schema.get("total"):
        df_orig = df_original.copy()
        df_orig[schema["fecha"]] = pd.to_datetime(df_orig[schema["fecha"]], errors="coerce")
        df_orig = df_orig.dropna(subset=[schema["fecha"]])
        
        # Agregar columnas mensuales
        for mes_col in config_cols["meses"]:
            # Extraer número de mes del nombre
            mes_nombre = mes_col.split()[0]
            mes_map = {
                "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4,
                "MAYO": 5, "JUNIO": 6, "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9
            }
            mes_num = mes_map.get(mes_nombre)
            
            if mes_num:
                # Filtrar datos del mes y año específico
                mask = (df_orig[schema["fecha"]].dt.year == anio) & \
                       (df_orig[schema["fecha"]].dt.month == mes_num)
                total_mes = df_orig.loc[mask, schema["total"]].sum()
                df[mes_col] = total_mes
        
        # Agregar columnas de totales (si hay información de sucursal)
        if schema.get("sucursal"):
            for total_col in config_cols["totales"]:
                if total_col == "TOTAL HIPER":
                    # Filtrar por sucursales tipo "HIPER" (ajustar según tu lógica)
                    mask = df_orig[schema["sucursal"]].str.contains("HIPER", case=False, na=False)
                    df[total_col] = df_orig.loc[mask, schema["total"]].sum()
                elif total_col == "TOTAL CORRIENTES":
                    # Filtrar por sucursales en Corrientes (ajustar según tu lógica)
                    mask = df_orig[schema["sucursal"]].str.contains("CORRIENTES", case=False, na=False)
                    df[total_col] = df_orig.loc[mask, schema["total"]].sum()
                elif total_col == "TOTAL CONSOLIDADO":
                    df[total_col] = df_orig[schema["total"]].sum()
        else:
            # Si no hay sucursal, poner totales generales
            for total_col in config_cols["totales"]:
                if total_col == "TOTAL CONSOLIDADO":
                    df[total_col] = df_orig[schema["total"]].sum()
                else:
                    df[total_col] = 0  # No se puede calcular sin info de sucursal
    
    return df


# ===================== ACCIONES =====================

def accion_totales_por_periodo(df, schema, periodo="mes", fecha_inicio=None, fecha_fin=None):
    if not schema["fecha"] or not schema["total"]:
        return "Requiere columna de fecha y total."

    d = df.copy()
    d[schema["fecha"]] = pd.to_datetime(d[schema["fecha"]], errors="coerce")
    d = d.dropna(subset=[schema["fecha"]])

    if fecha_inicio:
        d = d[d[schema["fecha"]] >= pd.to_datetime(fecha_inicio)]
    if fecha_fin:
        d = d[d[schema["fecha"]] <= pd.to_datetime(fecha_fin)]

    if d.empty:
        return pd.DataFrame(columns=["Periodo", "TotalFacturado"])

    if periodo == "dia":
        d["Periodo"] = d[schema["fecha"]].dt.date
    elif periodo == "mes":
        d["Periodo"] = d[schema["fecha"]].dt.to_period("M").astype(str)
    elif periodo == "rango":
        total = d[schema["total"]].sum()
        return pd.DataFrame([{"Periodo": f"{fecha_inicio}–{fecha_fin}", "TotalFacturado": total}])
    else:
        d["Periodo"] = d[schema["fecha"]]

    tabla = (
        d.groupby("Periodo", as_index=False)[schema["total"]]
        .sum()
        .rename(columns={schema["total"]: "TotalFacturado"})
        .sort_values("Periodo")
    )
    return tabla


def accion_unidades_totales(df, schema, por="producto"):
    if not schema["cantidad"]:
        return "Requiere columna de cantidad."

    if por == "producto" and schema["producto"]:
        grupo = schema["producto"]
        nombre = "IdArticulo"
    elif por == "categoria" and schema["departamento"]:
        grupo = schema["departamento"]
        nombre = "Categoria"
    elif por == "vendedor" and schema["vendedor"]:
        grupo = schema["vendedor"]
        nombre = "Vendedor"
    else:
        return "No se ha mapeado la columna necesaria para esta agregación."

    tabla = (
        df.groupby(grupo, as_index=False)[schema["cantidad"]]
        .sum()
        .rename(columns={grupo: nombre, schema["cantidad"]: "UnidadesVendidas"})
        .sort_values("UnidadesVendidas", ascending=False)
    )
    return tabla


def accion_conteo_tickets(df, schema, por="producto"):
    if not schema["ticket"]:
        return "Requiere columna de ticket."

    d = df.copy()

    if por == "producto" and schema["producto"]:
        grupo = [schema["producto"]]
        nombre = "IdArticulo"
    elif por == "dia" and schema["fecha"]:
        d[schema["fecha"]] = pd.to_datetime(d[schema["fecha"]], errors="coerce")
        d = d.dropna(subset=[schema["fecha"]])
        d["Dia"] = d[schema["fecha"]].dt.date
        grupo = ["Dia"]
        nombre = "Dia"
    elif por == "vendedor" and schema["vendedor"]:
        grupo = [schema["vendedor"]]
        nombre = "Vendedor"
    else:
        return "No se ha mapeado la columna necesaria para esta agregación."

    tabla = (
        d.groupby(grupo)[schema["ticket"]]
        .nunique()
        .reset_index(name="CantidadTickets")
        .rename(columns={grupo[0]: nombre})
        .sort_values("CantidadTickets", ascending=False)
    )
    return tabla


def accion_productos_unicos(df, schema):
    if not schema["producto"]:
        return "Requiere columna de IdArticulo."
    n = df[schema["producto"]].nunique()
    lista = (
        df[[schema["producto"], schema.get("descripcion", schema["producto"])]]
        .drop_duplicates()
        .reset_index(drop=True)
    )
    return n, lista


def accion_productos_unicos_mes(df, schema):
    if not schema["producto"] or not schema["fecha"]:
        return "Requiere IdArticulo y fecha."
    d = df.copy()
    d[schema["fecha"]] = pd.to_datetime(d[schema["fecha"]], errors="coerce")
    d = d.dropna(subset=[schema["fecha"]])
    d["Mes"] = d[schema["fecha"]].dt.to_period("M").astype(str)
    tabla = (
        d.groupby("Mes")[schema["producto"]]
        .nunique()
        .reset_index(name="ProductosUnicos")
        .sort_values("Mes")
    )
    return tabla


def accion_clientes_unicos(df, schema):
    if not schema["cliente"]:
        return "Requiere columna de cliente."
    return df[schema["cliente"]].nunique()


def accion_clientes_unicos_mes(df, schema):
    if not schema["cliente"] or not schema["fecha"]:
        return "Requiere cliente y fecha."
    d = df.copy()
    d[schema["fecha"]] = pd.to_datetime(d[schema["fecha"]], errors="coerce")
    d = d.dropna(subset=[schema["fecha"]])
    d["Mes"] = d[schema["fecha"]].dt.to_period("M").astype(str)
    tabla = (
        d.groupby("Mes")[schema["cliente"]]
        .nunique()
        .reset_index(name="ClientesUnicos")
        .sort_values("Mes")
    )
    return tabla


def accion_clientes_recurrentes(df, schema, min_veces=2):
    if not schema["cliente"] or not schema["ticket"]:
        return "Requiere cliente y ticket."
    tickets_por_cliente = (
        df.groupby(schema["cliente"])[schema["ticket"]]
        .nunique()
        .reset_index(name="Compras")
    )
    recurrentes = tickets_por_cliente[tickets_por_cliente["Compras"] >= min_veces]
    return recurrentes


def accion_precio_promedio_producto(df, schema):
    if not schema["producto"] or not schema["precio"]:
        return "Requiere IdArticulo y precio unitario."
    tabla = (
        df.groupby(schema["producto"], as_index=False)[schema["precio"]]
        .mean()
        .rename(columns={schema["producto"]: "IdArticulo",
                         schema["precio"]: "PrecioPromedio"})
        .sort_values("PrecioPromedio", ascending=False)
    )
    return tabla


def accion_ticket_promedio_por(df, schema, por="dia"):
    if not schema["ticket"] or not schema["total"]:
        return "Requiere ticket y total."

    d = df.copy()
    totales_por_ticket = (
        d.groupby(schema["ticket"], as_index=False)[schema["total"]]
        .sum()
        .rename(columns={schema["total"]: "TotalTicket"})
    )

    if por == "dia" and schema["fecha"]:
        d[schema["fecha"]] = pd.to_datetime(d[schema["fecha"]], errors="coerce")
        d = d.dropna(subset=[schema["fecha"]])
        mapa_ticket_dia = d.groupby(schema["ticket"])[schema["fecha"]].min().dt.date
        totales_por_ticket["Dia"] = totales_por_ticket[schema["ticket"]].map(mapa_ticket_dia)
        tabla = (
            totales_por_ticket.groupby("Dia")["TotalTicket"]
            .mean()
            .reset_index(name="TicketPromedio")
        )
    elif por == "vendedor" and schema["vendedor"]:
        mapa_ticket_vend = d.groupby(schema["ticket"])[schema["vendedor"]].first()
        totales_por_ticket["Vendedor"] = totales_por_ticket[schema["ticket"]].map(mapa_ticket_vend)
        tabla = (
            totales_por_ticket.groupby("Vendedor")["TotalTicket"]
            .mean()
            .reset_index(name="TicketPromedio")
        )
    else:
        return "Falta mapear fecha o vendedor."

    return tabla


def accion_participacion(df, schema, nivel="producto"):
    if not schema["total"]:
        return "Requiere columna total."

    if nivel == "producto" and schema["producto"]:
        grupo = schema["producto"]
        nombre = "IdArticulo"
    elif nivel == "familia" and schema["departamento"]:
        grupo = schema["departamento"]
        nombre = "Familia"
    else:
        return "No se ha mapeado la columna necesaria."

    d = df.copy()
    tabla = (
        d.groupby(grupo)[schema["total"]]
        .sum()
        .reset_index(name="Total")
        .rename(columns={grupo: nombre})
    )
    total_general = tabla["Total"].sum()
    tabla["Participacion_%"] = (tabla["Total"] / total_general * 100).round(2)
    tabla = tabla.sort_values("Total", ascending=False)
    return tabla


def accion_segmentacion_sucursal(df, schema):
    if not schema["sucursal"] or not schema["total"]:
        return "Requiere sucursal y total."
    tabla = (
        df.groupby(schema["sucursal"], as_index=False)[schema["total"]]
        .sum()
        .rename(columns={schema["sucursal"]: "Sucursal",
                         schema["total"]: "TotalFacturado"})
        .sort_values("TotalFacturado", ascending=False)
    )
    return tabla


def accion_maestro_productos(df, schema):
    cols = []
    for key in ["producto", "descripcion", "departamento"]:
        col = schema.get(key)
        if col:
            cols.append(col)
    if not cols:
        return "No hay columnas de producto / descripción / departamento."

    maestro = df[cols].drop_duplicates().reset_index(drop=True)
    return maestro


def accion_ventas_duplicadas(df, schema):
    subset = []
    for key in ["ticket", "fecha", "cliente", "total"]:
        col = schema.get(key)
        if col:
            subset.append(col)
    if len(subset) < 2:
        return "Se necesitan al menos dos columnas (ej. ticket y fecha) para detectar duplicados."

    d = df.copy()
    if schema.get("fecha"):
        d[schema["fecha"]] = pd.to_datetime(d[schema["fecha"]], errors="coerce")

    duplicados_mask = d.duplicated(subset=subset, keep=False)
    dup = d[duplicados_mask].sort_values(subset)
    return dup


def accion_normalizar_fechas(df, schema):
    if not schema["fecha"]:
        return "Requiere columna de fecha."
    d = df.copy()
    d[schema["fecha"]] = pd.to_datetime(d[schema["fecha"]], errors="coerce")
    d["Fecha_normalizada"] = d[schema["fecha"]].dt.date
    d["Mes"] = d[schema["fecha"]].dt.to_period("M").astype(str)
    d["Anio"] = d[schema["fecha"]].dt.year
    return d


def accion_tabla_mensual(df, schema):
    if not schema["fecha"] or not schema["total"]:
        return "Requiere fecha y total."
    d = df.copy()
    d[schema["fecha"]] = pd.to_datetime(d[schema["fecha"]], errors="coerce")
    d = d.dropna(subset=[schema["fecha"]])
    d["Anio"] = d[schema["fecha"]].dt.year
    d["Mes"] = d[schema["fecha"]].dt.month

    tabla = (
        d.groupby(["Anio", "Mes"], as_index=False)[schema["total"]]
        .sum()
        .rename(columns={schema["total"]: "TotalFacturado"})
        .sort_values(["Anio", "Mes"])
    )
    return tabla


def accion_comparacion_mensual(df, schema):
    base = accion_tabla_mensual(df, schema)
    if isinstance(base, str):
        return base

    base = base.copy()
    base["Periodo"] = base["Anio"].astype(str) + "-" + base["Mes"].astype(str).str.zfill(2)
    base = base.sort_values(["Anio", "Mes"]).reset_index(drop=True)

    base["TotalMesAnterior"] = base["TotalFacturado"].shift(1)
    base["Delta_vs_MesAnterior"] = (base["TotalFacturado"] - base["TotalMesAnterior"]).fillna(0)
    base["TotalMismoMes_AñoAnterior"] = base["TotalFacturado"].shift(12)
    base["Delta_vs_AñoAnterior"] = (base["TotalFacturado"] - base["TotalMismoMes_AñoAnterior"]).fillna(0)

    return base


def accion_top_bottom(df, schema, nivel="producto", n=10):
    if not schema["total"]:
        return "Requiere total."

    d = df.copy()

    if nivel == "producto" and schema["producto"]:
        grupo = schema["producto"]
        nombre = "IdArticulo"
        g = d.groupby(grupo)[schema["total"]].sum().reset_index(name="Total")
        g = g.rename(columns={grupo: nombre})
    elif nivel == "dia" and schema["fecha"]:
        d[schema["fecha"]] = pd.to_datetime(d[schema["fecha"]], errors="coerce")
        d = d.dropna(subset=[schema["fecha"]])
        d["Dia"] = d[schema["fecha"]].dt.date
        g = d.groupby("Dia")[schema["total"]].sum().reset_index(name="Total")
    elif nivel == "vendedor" and schema["vendedor"]:
        grupo = schema["vendedor"]
        nombre = "Vendedor"
        g = d.groupby(grupo)[schema["total"]].sum().reset_index(name="Total")
        g = g.rename(columns={grupo: nombre})
    else:
        return "Faltan columnas para este análisis."

    top = g.sort_values("Total", ascending=False).head(n)
    bottom = g.sort_values("Total", ascending=True).head(n)
    return top, bottom


def accion_sumatoria_ventas_mensuales_por_idarticulo(df, schema):
    """Sumatoria de ventas mensuales por IdArticulo.
    Mapea cada mes como columna, filtrando por la fecha del Excel fuente.
    """
    if not schema["producto"] or not schema["fecha"] or not schema["total"]:
        return "Requiere IdArticulo, fecha y total."
    
    d = df.copy()
    d[schema["fecha"]] = pd.to_datetime(d[schema["fecha"]], errors="coerce")
    d = d.dropna(subset=[schema["fecha"]])
    
    # Crear columnas de año y mes
    d["Anio"] = d[schema["fecha"]].dt.year
    d["Mes"] = d[schema["fecha"]].dt.month
    d["Mes_Nombre"] = d[schema["fecha"]].dt.strftime("%B").map({
        "January": "ENERO", "February": "FEBRERO", "March": "MARZO",
        "April": "ABRIL", "May": "MAYO", "June": "JUNIO",
        "July": "JULIO", "August": "AGOSTO", "September": "SEPTIEMBRE",
        "October": "OCTUBRE", "November": "NOVIEMBRE", "December": "DICIEMBRE"
    })
    
    # Crear etiqueta Mes-Año
    d["Mes_Anio"] = d["Mes_Nombre"] + " " + d["Anio"].astype(str)
    
    # Agrupar por IdArticulo y Mes_Anio, sumando el total
    resultado = d.groupby([schema["producto"], "Mes_Anio"], as_index=False)[schema["total"]].sum()
    resultado = resultado.rename(columns={schema["producto"]: "IdArticulo", schema["total"]: "Venta"})
    
    # Pivotar para que cada mes sea una columna
    tabla_pivot = resultado.pivot_table(
        index="IdArticulo",
        columns="Mes_Anio",
        values="Venta",
        aggfunc="sum",
        fill_value=0
    ).reset_index()
    
    # Reordenar columnas de meses en orden cronológico
    meses_orden = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
                   "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
    
    columnas_ordenadas = ["IdArticulo"]
    for col in tabla_pivot.columns[1:]:
        # Extraer mes y año de la columna
        parts = col.split(" ")
        if len(parts) == 2:
            mes_nombre, anio = parts
            if mes_nombre in meses_orden:
                columnas_ordenadas.append(col)
    
    tabla_pivot = tabla_pivot[[c for c in columnas_ordenadas if c in tabla_pivot.columns]]
    
    # Agregar columna de TOTAL
    tabla_pivot["TOTAL"] = tabla_pivot.iloc[:, 1:].sum(axis=1)
    
    return tabla_pivot

# ===================== REGISTRO DE ACCIONES =====================

ACCIONES = {
    "Totales facturados por mes": {
        "fn": lambda df, schema, **kw: accion_totales_por_periodo(df, schema, "mes", **kw),
        "tipo": "tabla",
        "descripcion": "Suma de ventas por mes calendario.",
    },
    "Totales facturados por día": {
        "fn": lambda df, schema, **kw: accion_totales_por_periodo(df, schema, "dia", **kw),
        "tipo": "tabla",
        "descripcion": "Suma de ventas por día.",
    },
    "Totales facturados en rango": {
        "fn": lambda df, schema, **kw: accion_totales_por_periodo(df, schema, "rango", **kw),
        "tipo": "tabla",
        "descripcion": "Total de ventas en el rango de fechas.",
    },
    "Unidades por producto": {
        "fn": lambda df, schema, **kw: accion_unidades_totales(df, schema, "producto"),
        "tipo": "tabla",
        "descripcion": "Unidades totales vendidas por IdArticulo.",
    },
    "Unidades por categoría": {
        "fn": lambda df, schema, **kw: accion_unidades_totales(df, schema, "categoria"),
        "tipo": "tabla",
        "descripcion": "Unidades totales por familia/departamento.",
    },
    "Unidades por vendedor": {
        "fn": lambda df, schema, **kw: accion_unidades_totales(df, schema, "vendedor"),
        "tipo": "tabla",
        "descripcion": "Unidades totales vendidas por vendedor.",
    },
    "Tickets por producto": {
        "fn": lambda df, schema, **kw: accion_conteo_tickets(df, schema, "producto"),
        "tipo": "tabla",
        "descripcion": "Número de tickets en los que aparece cada producto.",
    },
    "Tickets por día": {
        "fn": lambda df, schema, **kw: accion_conteo_tickets(df, schema, "dia"),
        "tipo": "tabla",
        "descripcion": "Número de tickets por día.",
    },
    "Tickets por vendedor": {
        "fn": lambda df, schema, **kw: accion_conteo_tickets(df, schema, "vendedor"),
        "tipo": "tabla",
        "descripcion": "Número de tickets atendidos por cada vendedor.",
    },
    "Productos únicos vendidos": {
        "fn": lambda df, schema, **kw: accion_productos_unicos(df, schema),
        "tipo": "mixto",
        "descripcion": "Cantidad y lista de productos distintos vendidos en el periodo.",
    },
    "Productos únicos por mes": {
        "fn": lambda df, schema, **kw: accion_productos_unicos_mes(df, schema),
        "tipo": "tabla",
        "descripcion": "Cantidad de productos distintos vendidos por mes.",
    },
    "Clientes únicos (KPI)": {
        "fn": lambda df, schema, **kw: accion_clientes_unicos(df, schema),
        "tipo": "kpi",
        "descripcion": "Número de clientes distintos en el periodo.",
    },
    "Clientes únicos por mes": {
        "fn": lambda df, schema, **kw: accion_clientes_unicos_mes(df, schema),
        "tipo": "tabla",
        "descripcion": "Cantidad de clientes distintos por mes.",
    },
    "Clientes recurrentes (>=2 compras)": {
        "fn": lambda df, schema, **kw: accion_clientes_recurrentes(df, schema, min_veces=2),
        "tipo": "tabla",
        "descripcion": "Clientes con dos o más compras.",
    },
    "Precio promedio por producto": {
        "fn": lambda df, schema, **kw: accion_precio_promedio_producto(df, schema),
        "tipo": "tabla",
        "descripcion": "Precio promedio de venta por IdArticulo.",
    },
    "Ticket promedio por día": {
        "fn": lambda df, schema, **kw: accion_ticket_promedio_por(df, schema, "dia"),
        "tipo": "tabla",
        "descripcion": "Ticket promedio por día.",
    },
    "Ticket promedio por vendedor": {
        "fn": lambda df, schema, **kw: accion_ticket_promedio_por(df, schema, "vendedor"),
        "tipo": "tabla",
        "descripcion": "Ticket promedio por vendedor.",
    },
    "Participación por producto": {
        "fn": lambda df, schema, **kw: accion_participacion(df, schema, "producto"),
        "tipo": "tabla",
        "descripcion": "Participación porcentual de cada producto en el total facturado.",
    },
    "Participación por familia": {
        "fn": lambda df, schema, **kw: accion_participacion(df, schema, "familia"),
        "tipo": "tabla",
        "descripcion": "Participación de cada familia/departamento en el total.",
    },
    "Segmentación por sucursal": {
        "fn": lambda df, schema, **kw: accion_segmentacion_sucursal(df, schema),
        "tipo": "tabla",
        "descripcion": "Total facturado por sucursal o unidad de negocio.",
    },
    "Maestro de productos": {
        "fn": lambda df, schema, **kw: accion_maestro_productos(df, schema),
        "tipo": "tabla",
        "descripcion": "Lista única de productos con sus datos maestros.",
    },
    "Ventas duplicadas": {
        "fn": lambda df, schema, **kw: accion_ventas_duplicadas(df, schema),
        "tipo": "tabla",
        "descripcion": "Filas potencialmente duplicadas según ticket/fecha/cliente/total.",
    },
    "Normalizar fechas": {
        "fn": lambda df, schema, **kw: accion_normalizar_fechas(df, schema),
        "tipo": "tabla",
        "descripcion": "Añade columnas Fecha_normalizada, Mes y Anio.",
    },
    "Tabla mensual": {
        "fn": lambda df, schema, **kw: accion_tabla_mensual(df, schema),
        "tipo": "tabla",
        "descripcion": "Total facturado por año y mes.",
    },
    "Comparación vs mes anterior y año anterior": {
        "fn": lambda df, schema, **kw: accion_comparacion_mensual(df, schema),
        "tipo": "tabla",
        "descripcion": "Agrega columnas con diferencias vs mes anterior y mismo mes del año anterior.",
    },
    "Top/bottom productos": {
        "fn": lambda df, schema, **kw: accion_top_bottom(df, schema, "producto", n=10),
        "tipo": "mixto",
        "descripcion": "Top y bottom 10 productos por total facturado.",
    },
        "Sumatoria Ventas mensuales por IdArticulo": {
        "fn": lambda df, schema, **kw: accion_sumatoria_ventas_mensuales_por_idarticulo(df, schema),
        "tipo": "tabla",
        "descripcion": "Ventas mensuales agregadas por IdArticulo, con cada mes como columna.",
    },
}


# ===================== UI PRINCIPAL =====================

st.sidebar.header("1. Subir archivos")
uploaded_files = st.sidebar.file_uploader(
    "Excels de ventas", type=["xls", "xlsx"], accept_multiple_files=True
)

if not uploaded_files:
    st.info("Sube uno o más archivos Excel para comenzar.")
    st.stop()

st.sidebar.success(f"{len(uploaded_files)} archivo(s) cargado(s).")

df = leer_excels_subidos(uploaded_files)
st.write("Vista previa de datos combinados:", df.head())

schema = get_schema_mapping(df)

# Configuración de columnas adicionales
config_cols_adicionales = get_columnas_adicionales_config()

# Filtros de fecha para acciones que lo usen
usar_rango = False
rango_inicio = None
rango_fin = None
if schema["fecha"]:
    with st.expander("Filtros de fecha opcionales"):
        usar_rango = st.checkbox("Aplicar rango de fechas en acciones que lo soportan", value=False)
        if usar_rango:
            c1, c2 = st.columns(2)
            with c1:
                rango_inicio = st.date_input("Desde", value=df[schema["fecha"]].min())
            with c2:
                rango_fin = st.date_input("Hasta", value=df[schema["fecha"]].max())

st.markdown("---")
st.subheader("Elegir acciones a ejecutar")

acciones_disponibles = list(ACCIONES.keys())
acciones_sel = st.multiselect(
    "Selecciona una o varias acciones:",
    options=acciones_disponibles,
    default=["Totales facturados por mes", "Unidades por producto"],
)

ejecutar = st.button("Ejecutar análisis")

if ejecutar and acciones_sel:
    resultados_para_exportar = {}
    tabs = st.tabs(acciones_sel)

    for tab, nombre_accion in zip(tabs, acciones_sel):
        meta = ACCIONES[nombre_accion]
        fn = meta["fn"]
        with tab:
            st.markdown(f"**{nombre_accion}**")
            st.caption(meta["descripcion"])
            try:
                kwargs = {}
                if "rango" in nombre_accion and usar_rango:
                    kwargs["fecha_inicio"] = rango_inicio
                    kwargs["fecha_fin"] = rango_fin

                res = fn(df, schema, **kwargs)
                tipo = meta["tipo"]

                if isinstance(res, str):
                    st.warning(res)
                    continue

                if tipo == "tabla":
                    # Agregar columnas adicionales si están configuradas
                    if config_cols_adicionales["columnas"]:
                        res = agregar_columnas_adicionales(res, config_cols_adicionales, df, schema)
                    st.dataframe(res)
                    resultados_para_exportar[nombre_accion] = res
                elif tipo == "kpi":
                    st.metric(label=nombre_accion, value=res)
                elif tipo == "mixto":
                    if nombre_accion.startswith("Productos únicos"):
                        n, tabla = res
                        st.metric("Cantidad de productos únicos", n)
                        if config_cols_adicionales["columnas"]:
                            tabla = agregar_columnas_adicionales(tabla, config_cols_adicionales, df, schema)
                        st.dataframe(tabla)
                        resultados_para_exportar[nombre_accion] = tabla
                    else:
                        top, bottom = res
                        st.write("Top N:")
                        if config_cols_adicionales["columnas"]:
                            top = agregar_columnas_adicionales(top, config_cols_adicionales, df, schema)
                        st.dataframe(top)
                        st.write("Bottom N:")
                        if config_cols_adicionales["columnas"]:
                            bottom = agregar_columnas_adicionales(bottom, config_cols_adicionales, df, schema)
                        st.dataframe(bottom)
                        resultados_para_exportar[nombre_accion + "_TOP"] = top
                        resultados_para_exportar[nombre_accion + "_BOTTOM"] = bottom
            except Exception as e:
                st.error(f"Error en '{nombre_accion}': {e}")

    # Exportación conjunta a Excel
    if resultados_para_exportar:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            for nombre, df_res in resultados_para_exportar.items():
                sheet = nombre[:31]  # límite Excel
                df_res.to_excel(writer, sheet_name=sheet, index=False)
        buffer.seek(0)

        ts = datetime.now().strftime("%Y%m%d_%H%M")
        st.download_button(
            "📥 Descargar resultados en Excel",
            data=buffer,
            file_name=f"Resultados_Analisis_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
