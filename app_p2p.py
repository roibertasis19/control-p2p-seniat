import streamlit as st
import time
import hmac
import hashlib
import requests
import pandas as pd
from io import BytesIO
from datetime import datetime, timezone, timedelta

st.set_page_config(page_title="Sistema Contable P2P - Binance", layout="wide", page_icon="📊")

BASE_URL = "https://api.binance.com"

def generar_firma(query_string, secret):
    return hmac.new(secret.encode("utf-8"), query_string.encode("utf-8"), hashlib.sha256).hexdigest()

def consultar_ventana_p2p(api_key, api_secret, start_ts, end_ts):
    endpoint = "/sapi/v1/c2c/orderMatch/listUserOrderHistory"
    ordenes_ventana = []
    pagina = 1

    while True:
        timestamp = int(time.time() * 1000)
        query_string = f"startTimestamp={start_ts}&endTimestamp={end_ts}&page={pagina}&rows=100&timestamp={timestamp}"
        signature = generar_firma(query_string, api_secret)

        headers = {"X-MBX-APIKEY": api_key}
        url = f"{BASE_URL}{endpoint}?{query_string}&signature={signature}"

        try:
            respuesta = requests.get(url, headers=headers, timeout=15).json()
        except Exception as e:
            return None, f"Error de conexión: {e}"

        if respuesta.get("code") == "000000":
            ordenes = respuesta.get("data", [])
            if not ordenes:
                break
            ordenes_ventana.extend(ordenes)
            if len(ordenes) < 100:
                break
            pagina += 1
            time.sleep(0.25)
        else:
            return None, f"Error de API Binance: {respuesta.get('message', respuesta)}"

    return ordenes_ventana, None

@st.cache_data(ttl=60, show_spinner=False)
def obtener_historial_por_bloques(api_key, api_secret, dt_inicio, dt_fin, dias_bloque=15):
    todas_las_ordenes = []
    ordenes_vistas = set()
    cursor = dt_inicio

    while cursor <= dt_fin:
        fin_bloque = min(cursor + timedelta(days=dias_bloque) - timedelta(milliseconds=1), dt_fin)
        start_ts = int(cursor.timestamp() * 1000)
        end_ts = int(fin_bloque.timestamp() * 1000)

        ordenes_bloque, error = consultar_ventana_p2p(api_key, api_secret, start_ts, end_ts)
        if error:
            return None, f"Fallo en bloque {cursor.date()} a {fin_bloque.date()}: {error}"

        for orden in (ordenes_bloque or []):
            order_id = str(orden.get("orderNumber"))
            if order_id not in ordenes_vistas:
                ordenes_vistas.add(order_id)
                todas_las_ordenes.append(orden)

        cursor = fin_bloque + timedelta(milliseconds=1)

    return todas_las_ordenes, None

def procesar_datos(ordenes_raw, pct_comision_banco, zona_horaria_ve=-4):
    registros = []
    tz_ve = timezone(timedelta(hours=zona_horaria_ve))

    for orden in ordenes_raw:
        if orden.get("orderStatus") == "COMPLETED":
            dt_ve = datetime.fromtimestamp(orden["createTime"] / 1000.0, tz=tz_ve)
            fecha_str = dt_ve.strftime("%Y-%m-%d %H:%M:%S")

            tipo = "COMPRA" if orden["tradeType"] == "BUY" else "VENTA"
            cantidad_usdt = float(orden["amount"])
            tasa_bs = float(orden["unitPrice"])
            total_bs = float(orden["totalPrice"])
            comision_binance_usdt = float(orden.get("commission", 0))

            comision_banco_bs = (total_bs * (pct_comision_banco / 100.0)) if tipo == "COMPRA" else 0.0
            usdt_netos = (cantidad_usdt - comision_binance_usdt) if tipo == "COMPRA" else (cantidad_usdt + comision_binance_usdt)

            registros.append({
                "Fecha_Hora_VE": fecha_str,
                "Numero_Orden": str(orden["orderNumber"]),
                "Operacion": tipo,
                "Activo": orden["asset"],
                "Cantidad_Orden_USDT": cantidad_usdt,
                "Fee_Binance_USDT": comision_binance_usdt,
                "USDT_Netos_Billetera": usdt_netos,
                "Tasa_Bs": tasa_bs,
                "Monto_Orden_Bs": total_bs,
                "Comision_Bancaria_Bs": round(comision_banco_bs, 2),
                "Flujo_Total_Banco_Bs": round(total_bs + comision_banco_bs if tipo == "COMPRA" else total_bs, 2),
                "Banco_Metodo": orden.get("payMethodName", "No especificado")
            })

    df = pd.DataFrame(registros)
    if not df.empty:
        df = df.sort_values("Fecha_Hora_VE", ascending=False).reset_index(drop=True)
    return df

def generar_excel_contador(df, resumen_dict, df_bancos):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame([resumen_dict]).to_excel(writer, sheet_name="1_Resumen_Fiscal_SENIAT", index=False)
        df.to_excel(writer, sheet_name="2_Libro_Diario_Operaciones", index=False)
        df_bancos.to_excel(writer, sheet_name="3_Conciliacion_Bancos", index=False)
    return output.getvalue()

st.title("📊 Control Operativo y Contable P2P (SENIAT)")
st.markdown("Consulta por bloques desde Binance API, conciliación bancaria y exportación en Excel.")

tz_ve = timezone(timedelta(hours=-4))
hoy_ve = datetime.now(tz_ve).date()

with st.sidebar:
    st.header("🔑 Credenciales API")
    st.caption("Usa claves de **solo lectura**. Se mantienen solo en tu sesión activa.")
    api_key = st.text_input("Binance API Key", type="password")
    api_secret = st.text_input("Binance API Secret", type="password")

    st.divider()
    st.header("⚙️ Parámetros Contables")

    fecha_inicio = st.date_input("Fecha Inicio", value=hoy_ve)
    fecha_fin = st.date_input("Fecha Fin", value=hoy_ve)

    dias_bloque = st.selectbox(
        "Bloque de consulta API (días)",
        options=[7, 15, 30],
        index=1,
        help="Divide rangos antiguos en bloques para evitar límites o timeouts en consultas largas."
    )

    pct_banco = st.number_input(
        "Comisión Bancaria promedio en Compras (%)",
        min_value=0.0, max_value=5.0, value=0.30, step=0.05,
        help="Ej: 0.30% por transferencias interbancarias o Pago Móvil al pagar órdenes de compra."
    )

    btn_consultar = st.button("🔄 Sincronizar con Binance", type="primary", use_container_width=True)

st.subheader("🔎 Filtro Rápido de Operaciones")
col_b1, col_b2 = st.columns(2)
buscar_orden = col_b1.text_input("Buscar por N° de Orden", placeholder="Ej: 22956139037815148544")
buscar_banco = col_b2.text_input("Filtrar por Banco / Método", placeholder="Ej: Banesco, Mercantil, Pago Móvil")

if "df_actual" not in st.session_state:
    st.session_state.df_actual = None
if "rango_actual" not in st.session_state:
    st.session_state.rango_actual = (fecha_inicio, fecha_fin)

if btn_consultar:
    if not api_key or not api_secret:
        st.error("⚠️ Por favor ingresa tu API Key y API Secret en el panel izquierdo.")
    elif fecha_inicio > fecha_fin:
        st.error("⚠️ La Fecha Inicio no puede ser mayor que la Fecha Fin.")
    else:
        dt_inicio = datetime.combine(fecha_inicio, datetime.min.time(), tzinfo=tz_ve)
        dt_fin = datetime.combine(fecha_fin, datetime.max.time(), tzinfo=tz_ve)

        with st.spinner("Consultando historial P2P en Binance por bloques..."):
            ordenes, error = obtener_historial_por_bloques(api_key, api_secret, dt_inicio, dt_fin, dias_bloque)

        if error:
            st.error(error)
        elif not ordenes:
            st.warning("No se encontraron órdenes en el rango de fechas seleccionado.")
        else:
            df_proc = procesar_datos(ordenes, pct_banco)
            if df_proc.empty:
                st.warning("Hubo órdenes en el período, pero ninguna con estado 'COMPLETED'.")
            else:
                st.session_state.df_actual = df_proc
                st.session_state.rango_actual = (fecha_inicio, fecha_fin)

df_mostrar = st.session_state.df_actual

if df_mostrar is not None and not df_mostrar.empty:
    df_filtrado = df_mostrar.copy()
    if buscar_orden.strip():
        df_filtrado = df_filtrado[df_filtrado["Numero_Orden"].str.contains(buscar_orden.strip(), case=False, na=False)]
    if buscar_banco.strip():
        df_filtrado = df_filtrado[df_filtrado["Banco_Metodo"].str.contains(buscar_banco.strip(), case=False, na=False)]

    if df_filtrado.empty:
        st.info("No hay operaciones que coincidan con el filtro de búsqueda actual.")
    else:
        compras = df_filtrado[df_filtrado["Operacion"] == "COMPRA"]
        ventas = df_filtrado[df_filtrado["Operacion"] == "VENTA"]

        usdt_comprados_netos = compras["USDT_Netos_Billetera"].sum()
        bs_pagados_ordenes = compras["Monto_Orden_Bs"].sum()
        bs_comisiones_banco = compras["Comision_Bancaria_Bs"].sum()
        costo_total_compras_bs = bs_pagados_ordenes + bs_comisiones_banco
        tasa_prom_compra = costo_total_compras_bs / usdt_comprados_netos if usdt_comprados_netos > 0 else 0

        usdt_vendidos_brutos = ventas["USDT_Netos_Billetera"].sum()
        usdt_entregados_cliente = ventas["Cantidad_Orden_USDT"].sum()
        bs_recibidos_ventas = ventas["Monto_Orden_Bs"].sum()
        tasa_prom_venta = bs_recibidos_ventas / usdt_entregados_cliente if usdt_entregados_cliente > 0 else 0

        costo_de_lo_vendido_bs = usdt_vendidos_brutos * tasa_prom_compra
        ganancia_neta_bs = bs_recibidos_ventas - costo_de_lo_vendido_bs
        ganancia_neta_usdt = ganancia_neta_bs / tasa_prom_venta if tasa_prom_venta > 0 else 0
        margen_porcentaje = (ganancia_neta_bs / costo_de_lo_vendido_bs * 100) if costo_de_lo_vendido_bs > 0 else 0

        inventario_usdt = usdt_comprados_netos - usdt_vendidos_brutos
        valor_inventario_bs = inventario_usdt * tasa_prom_compra

        f_ini, f_fin = st.session_state.rango_actual
        st.divider()
        st.subheader(f"📈 Resumen Operativo ({f_ini} al {f_fin})")
        st.caption(f"Mostrando {len(df_filtrado)} operaciones")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Ganancia Neta Real (Bs.)", f"{ganancia_neta_bs:,.2f} Bs", f"{margen_porcentaje:.2f}% margen")
        col2.metric("Ganancia Neta Est. (USDT)", f"${ganancia_neta_usdt:,.2f} USDT")
        col3.metric("Tasa Prom. Compra / Venta", f"{tasa_prom_compra:.2f} / {tasa_prom_venta:.2f} Bs")
        col4.metric("Inventario Remanente", f"{inventario_usdt:,.2f} USDT", f"Valuado en {valor_inventario_bs:,.2f} Bs")

        st.divider()

        df_bancos = df_filtrado.groupby(["Banco_Metodo", "Operacion"]).agg(
            Cantidad_Operaciones=("Numero_Orden", "count"),
            Total_USDT=("Cantidad_Orden_USDT", "sum"),
            Monto_Ordenes_Bs=("Monto_Orden_Bs", "sum"),
            Comisiones_Banco_Bs=("Comision_Bancaria_Bs", "sum"),
            Movimiento_Real_Cuenta_Bs=("Flujo_Total_Banco_Bs", "sum")
        ).reset_index()

        tab1, tab2, tab3 = st.tabs(["📑 Libro Diario de Operaciones", "🏦 Conciliación por Banco", "🏛️ Vista Previa SENIAT"])

        with tab1:
            st.dataframe(df_filtrado, use_container_width=True)

        with tab2:
            st.info("Usa esta tabla para cuadrar exactamente los estados de cuenta de tus bancos nacionales.")
            st.dataframe(df_bancos, use_container_width=True)

        with tab3:
            resumen_fiscal = {
                "Fecha_Inicio": str(f_ini),
                "Fecha_Fin": str(f_fin),
                "Total_Operaciones": len(df_filtrado),
                "Ingresos_Brutos_Bancarios_Bs": round(bs_recibidos_ventas, 2),
                "Costo_Adquisicion_USDT_Vendidos_Bs": round(costo_de_lo_vendido_bs, 2),
                "Gastos_Comisiones_Bancarias_Bs": round(bs_comisiones_banco, 2),
                "Utilidad_Neta_Gravable_Bs": round(ganancia_neta_bs, 2),
                "Utilidad_Equivalente_USDT": round(ganancia_neta_usdt, 2),
                "Inventario_Final_USDT": round(inventario_usdt, 4),
                "Valor_Inventario_Final_Bs": round(valor_inventario_bs, 2),
                "Total_Comisiones_Binance_USDT": round(df_filtrado["Fee_Binance_USDT"].sum(), 4)
            }
            st.json(resumen_fiscal)

        excel_bytes = generar_excel_contador(df_filtrado, resumen_fiscal, df_bancos)
        st.download_button(
            label="📥 Descargar Reporte Contable Excel (.xlsx) para Contador",
            data=excel_bytes,
            file_name=f"Contabilidad_P2P_SENIAT_{f_ini}_a_{f_fin}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )