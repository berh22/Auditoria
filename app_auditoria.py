import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime

# Configuración global del sistema
st.set_page_config(page_title="Sistema de Control Transaccional - Cheques", layout="wide")

# ==========================================
# FUNCIÓN MAESTRA DE SANITIZACIÓN (ANTI-CRASH)
# ==========================================
def sanitizar_para_streamlit(df):
    if df is None or df.empty:
        return df
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].fillna('').astype(str).str.strip()
    return df

def limpiar_columna_monto(df, columna):
    if columna in df.columns:
        df[columna] = df[columna].astype(str).str.replace('$', '', regex=False)
        df[columna] = df[columna].str.replace(',', '', regex=False)
        df[columna] = df[columna].str.strip()
        df[columna] = pd.to_numeric(df[columna], errors='coerce').fillna(0)
    return df

# ==========================================
# 1. BASE DE DATOS INTERNA (PERSISTENCIA)
# ==========================================
if "MASTER_CHEQUES" not in st.session_state:
    st.session_state["MASTER_CHEQUES"] = pd.DataFrame(columns=[
        "EMPRESA", "FECHA DE EMISIÓN", "FOLIO", "MONTO", "BENEFICIARIO", 
        "CUENTA CON SELLO", "FECHA DE COBRO", "CONCEPTO", "OBSERVACIONES", 
        "CAUSA DE LA BAJA", "FECHA DE BAJA", "ESTATUS_CHEQUE"
    ])

# ==========================================
# 2. PANTALLA DE LOGIN
# ==========================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.write("### 🛡️ Control de Accesos - Auditoría")
        with st.form("login"):
            u = st.text_input("Usuario:")
            p = st.text_input("Contraseña:", type="password")
            if st.form_submit_button("Ingresar"):
                if u == "auditor_interno" and p == "ControlTrans2026*":
                    st.session_state["autenticado"] = True
                    st.rerun()
                else: st.error("Credenciales incorrectas.")
    st.stop()

# ==========================================
# 3. BARRA LATERAL Y NAVEGACIÓN
# ==========================================
st.sidebar.title("🛡️ Sistema Integral")
if st.sidebar.button("🔒 Cerrar Sesión"):
    st.session_state["autenticado"] = False
    st.rerun()

st.sidebar.markdown("---")
opcion = st.sidebar.radio("Menú Principal:", ["📊 Dashboard y Gráficos", "📥 Alimentar / Capturar Cheques", "🔍 Auditoría y Alertas Forenses"])

LISTA_EMPRESAS = [
    "Bajio Tropper", "Bajio TCS", "Bajio TCS Gerencia", 
    "Bajio Somnus", "Bajio Idea", "Bajio Areng", 
    "BBVA Tropper", "BBVA Sab Solutions"
]
CHEQUERAS_INACTIVAS = ["BBVA Tropper", "BBVA Sab Solutions"]

# ==========================================
# MÓDULO: ALIMENTAR / CAPTURAR CHEQUES
# ==========================================
if opcion == "📥 Alimentar / Capturar Cheques":
    st.title("📥 Administración y Alimentación de la Base de Datos")
    st.markdown("---")
    
    tab_carga, tab_manual, tab_exportar = st.tabs(["📁 Subir Archivo (Excel/Contpaq)", "✍️ Captura Manual Individual", "💾 Exportar Base Maestra"])
    
    with tab_carga:
        st.subheader("Cargar Nuevos Cheques masivamente")
        archivo_nuevo = st.file_uploader("Sube el Excel con los cheques nuevos del mes:", type=["xlsx"])
        empresa_carga = st.selectbox("¿A qué empresa corresponden estos cheques?", LISTA_EMPRESAS, key="emp_carga")
        
        if archivo_nuevo:
            df_raw = pd.read_excel(archivo_nuevo, header=None)
            
            if df_raw.empty:
                st.warning("⚠️ El archivo cargado parece estar vacío.")
            else:
                keywords = ["FOLIO", "BENEFICIARIO", "MONTO", "FECHA DE EMISIÓN", "CONCEPTO"]
                header_idx = None
                
                for idx, row in df_raw.iterrows():
                    row_str = [str(cell).upper() for cell in row.values]
                    matches = sum(1 for kw in keywords if any(kw in cell for cell in row_str))
                    if matches >= 2:
                        header_idx = idx
                        break
                
                if header_idx is not None:
                    df_temp = df_raw.iloc[header_idx+1:].copy()
                    
                    # Encabezados limpios en mayúsculas
                    encabezados_crudos = [str(c).strip().upper() for c in df_raw.iloc[header_idx].values]
                    
                    # Forzar nombres únicos para columnas basura de la derecha (nan)
                    encabezados_unicos = []
                    conteos = {}
                    for col in encabezados_crudos:
                        if col in conteos:
                            conteos[col] += 1
                            encabezados_unicos.append(f"{col}_{conteos[col]}")
                        else:
                            conteos[col] = 0
                            encabezados_unicos.append(col)
                    
                    df_temp.columns = encabezados_unicos
                    
                    st.info(f"📋 ¡Formato de Chequera Detectado! Omitimos las filas de título del banco.")
                    st.write("Vista previa de las primeras filas identificadas:")
                    st.dataframe(df_temp.head(3), width='stretch')
                    
                    if st.button("🚀 Confirmar e Integrar estas filas al Histórico Master"):
                        # 🎯 DICCIONARIO DE ALIAS EXACTOS (Evita colisiones de nombres)
                        mapeo_columnas = {
                            "FECHA DE EMISIÓN": ["FECHA DE EMISIÓN", "FECHA DE EMISION", "FECHA EMISION", "FECHA EMISIÓN"],
                            "FOLIO": ["FOLIO", "NUMERO", "NÚMERO", "NÚM", "NUM", "CHEQUE", "FOLIOS"],
                            "MONTO": ["MONTO", "IMPORTE", "CANTIDAD", "CARGOS"],
                            "BENEFICIARIO": ["BENEFICIARIO", "NOMBRE", "PROVEEDOR", "A FAVOR DE"],
                            "CUENTA CON SELLO": ["CUENTA CON SELLO", "SELLO", "CON SELLO"],
                            "FECHA DE COBRO": ["FECHA DE COBRO (DEPOSITO EN CUENTA)", "FECHA DE COBRO", "FECHA COBRO", "COBRO", "DEPOSITO"],
                            "CONCEPTO": ["CONCEPTO", "MOTIVO", "DESCRIPCION", "DESCRIPCIÓN"],
                            "OBSERVACIONES": ["OBSERVACIONES", "NOTAS", "COMENTARIOS", "OBSERVACION"]
                        }
                        
                        columnas_nuevas = {}
                        for col_excel in df_temp.columns:
                            for col_estandar, lista_alias in mapeo_columnas.items():
                                if col_excel in lista_alias: # ⚡ CAMBIO CRÍTICO: Búsqueda exacta para evitar choques
                                    columnas_nuevas[col_excel] = col_estandar
                                    break
                        df_temp = df_temp.rename(columns=columnas_nuevas)
                        
                        for col_estandar in st.session_state["MASTER_CHEQUES"].columns:
                            if col_estandar not in df_temp.columns:
                                df_temp[col_estandar] = np.nan if "FECHA" in col_estandar else ""
                        
                        # Filtro seguro para remover las filas en blanco del cierre de mes
                        df_temp = df_temp[df_temp["FECHA DE EMISIÓN"].notna()]
                        df_temp = df_temp[df_temp["FECHA DE EMISIÓN"].astype(str).str.strip() != ""]
                        df_temp = df_temp[df_temp["FECHA DE EMISIÓN"].astype(str).str.upper().str.strip() != "NAN"]
                        
                        df_temp = limpiar_columna_monto(df_temp, 'MONTO')
                        df_temp["FOLIO"] = pd.to_numeric(df_temp["FOLIO"], errors='coerce').fillna(0).astype(int)
                        df_temp["EMPRESA"] = empresa_carga
                        
                        # Detección inteligente de cancelados
                        df_temp["ESTATUS_CHEQUE"] = df_temp.apply(
                            lambda r: "CANCELADO" if "CANCELADO" in str(r.get("OBSERVACIONES", "")).upper() or \
                                                     "CANCELADO" in str(r.get("BENEFICIARIO", "")).upper() or \
                                                     "CANCELADO" in str(r.get("CONCEPTO", "")).upper() \
                                                  else "EMITIDO", axis=1
                        )
                        
                        df_temp = df_temp.reindex(columns=st.session_state["MASTER_CHEQUES"].columns)
                        df_temp = sanitizar_para_streamlit(df_temp)
                        
                        st.session_state["MASTER_CHEQUES"] = pd.concat([st.session_state["MASTER_CHEQUES"], df_temp], ignore_index=True)
                        st.session_state["MASTER_CHEQUES"] = st.session_state["MASTER_CHEQUES"].drop_duplicates(subset=["EMPRESA", "FOLIO"], keep="last")
                        
                        st.success(f"🎉 ¡Éxito total! Se cargaron **{len(df_temp)}** cheques reales en la base de datos de {empresa_carga}.")
                        st.rerun()
                else:
                    st.error("❌ No encontramos la fila de encabezados en el archivo. Asegúrate de que el archivo incluya las columnas estándares de la chequera.")

    with tab_manual:
        st.subheader("Formulario de Captura Individual")
        with st.form("registro_manual", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            emp_m = c1.selectbox("Empresa:", LISTA_EMPRESAS)
            fecha_m = c2.date_input("Fecha de Emisión:")
            folio_m = c3.number_input("Folio del Cheque:", step=1, min_value=1)
            
            c4, c5, c6 = st.columns(3)
            monto_m = c4.number_input("Monto ($):", min_value=0.0, format="%.2f")
            benef_m = c5.text_input("Beneficiario:").upper()
            sello_m = c6.selectbox("¿Cuenta con Sello 'Abono en Cuenta'?", ["SI", "NO"])
            
            c7, c8 = st.columns(2)
            concepto_m = c7.selectbox("Concepto del Cheque:", [
                "1. Comprobación de gastos", "2. Finiquito", "3. Abono a financiera", 
                "4. Apertura de fondo fijo", "5. Pago de nómina", "6. Gasto por comprobar", 
                "7. Apoyo por defunción", "8. Gratificaciones", "9. Pensión alimenticia", 
                "10. Retención a proveedor para abono a financiera", "11. Pago de trámites", "12. Otro motivo"
            ])
            estatus_m = c8.selectbox("Estatus Inicial:", ["EMITIDO", "CANCELADO"])
            obs_m = st.text_area("Observaciones:")
            
            baja_causa, baja_fecha = "", ""
            if "Finiquito" in concepto_m:
                st.markdown("##### ⚠️ Datos Adicionales por Baja de Personal")
                cx1, cx2 = st.columns(2)
                baja_causa = cx1.text_input("Causa de la Baja:")
                baja_fecha = cx2.date_input("Fecha de la Baja:")
            
            if st.form_submit_button("Guardar Cheque en el Sistema"):
                nuevo_registro = {
                    "EMPRESA": emp_m, "FECHA DE EMISIÓN": pd.to_datetime(fecha_m), "FOLIO": int(folio_m),
                    "MONTO": monto_m, "BENEFICIARIO": benef_m, "CUENTA CON SELLO": sello_m,
                    "FECHA DE COBRO": np.nan, "CONCEPTO": concepto_m, "OBSERVACIONES": obs_m,
                    "CAUSA DE LA BAJA": baja_causa, "FECHA DE BAJA": str(baja_fecha), "ESTATUS_CHEQUE": estatus_m
                }
                df_manual = pd.DataFrame([nuevo_registro])
                df_manual = sanitizar_para_streamlit(df_manual)
                st.session_state["MASTER_CHEQUES"] = pd.concat([st.session_state["MASTER_CHEQUES"], df_manual], ignore_index=True)
                st.success(f"¡Cheque Folio {folio_m} guardado exitosamente!")
                st.rerun()

    with tab_exportar:
        st.subheader("Descargar Base de Datos Actualizada")
        st.write(f"Registros totales almacenados en la nube actualmente: **{len(st.session_state['MASTER_CHEQUES'])}**")
        
        if not st.session_state["MASTER_CHEQUES"].empty:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                st.session_state["MASTER_CHEQUES"].to_excel(writer, index=False, sheet_name="Master_Cheques")
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Descargar Master de Cheques en Excel",
                data=processed_data,
                file_name=f"Master_Cheques_Auditoria_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    st.markdown("---")
    st.subheader("📋 Base de Datos Actual del Sistema")
    if not st.session_state["MASTER_CHEQUES"].empty:
        st.dataframe(st.session_state["MASTER_CHEQUES"], width='stretch')
    else:
        st.info("La base de datos central está vacía. En cuanto subas un Excel válido y des clic en el botón de confirmar, verás la lista aquí.")

# ==========================================
# MÓDULO: AUDITORÍA Y ALERTAS FORENSES
# ==========================================
elif opcion == "🔍 Auditoría y Alertas Forenses":
    st.title("🔍 Escáner Forense de Riesgos de Control Interno")
    st.markdown("---")
    
    df_master = st.session_state["MASTER_CHEQUES"].copy()
    df_master = sanitizar_para_streamlit(df_master)
    
    if df_master.empty:
        st.info("No hay datos en el sistema. Ve al módulo de Carga para alimentar la base.")
        st.stop()

    st.subheader("⚡ Cruce de Validación Externa (Banco)")
    archivo_banco = st.file_uploader("Sube el Estado de Cuenta del Banco para buscar cobros indebidos:", type=["xlsx"])
    
    folios_banco_cobrados = []
    if archivo_banco:
        df_banco = pd.read_excel(archivo_banco)
        df_banco.columns = df_banco.columns.astype(str).str.upper().str.strip()
        df_banco = limpiar_columna_monto(df_banco, 'CARGOS')
        df_banco = limpiar_columna_monto(df_banco, 'SALDO')
        df_banco = sanitizar_para_streamlit(df_banco)
        
        for col_b in df_banco.columns:
            if df_banco[col_b].dtype in [np.int64, np.float64]:
                folios_banco_cobrados.extend(df_banco[col_b].dropna().astype(int).tolist())
            else:
                texto_completo = " ".join(df_banco[col_b].astype(str).tolist())
                numeros = [int(s) for s in re.findall(r'\b\d{1,6}\b', texto_completo)]
                folios_banco_cobrados.extend(numeros)
        folios_banco_cobrados = set(folios_banco_cobrados)
        st.success("¡Estado de cuenta indexado correctamente para el cruce de folios!")

    st.markdown("---")
    
    alertas_bbva = df_master[df_master["EMPRESA"].isin(CHEQUERAS_INACTIVAS)]
    
    df_master["CONCEPTO_STR"] = df_master["CONCEPTO"].fillna("").astype(str)
    finiquitos = df_master[df_master["CONCEPTO_STR"].str.contains("Finiquito", na=False)]
    dup_finiquitos = finiquitos[finiquitos.duplicated(subset=["BENEFICIARIO"], keep=False)]
    
    cancelados = df_master[df_master["ESTATUS_CHEQUE"] == "CANCELADO"].copy()
    cancelados["FOLIO_INT"] = pd.to_numeric(cancelados["FOLIO"], errors='coerce').fillna(-1).astype(int)
    cancelados["¿COBRADO_EN_BANCO?"] = cancelados["FOLIO_INT"].apply(lambda x: "🚨 SÍ COBRADO" if x in folios_banco_cobrados and x != -1 else "OK (No cobrado)")
    fraudes_cancelados = cancelados[cancelados["¿COBRADO_EN_BANCO?"] == "🚨 SÍ COBRADO"]
    
    gratif = df_master[df_master["CONCEPTO_STR"].str.contains("Gratificaciones", na=False)]
    
    gastos = df_master[df_master["CONCEPTO_STR"].str.contains("Gasto por comprobar", na=False)].copy()
    gastos["FECHA DE EMISIÓN"] = pd.to_datetime(gastos["FECHA DE EMISIÓN"], errors='coerce')
    fecha_auditoria = pd.to_datetime("2026-06-05")
    gastos["DIAS_PASADOS"] = (fecha_auditoria - gastos["FECHA DE EMISIÓN"]).dt.days
    gastos_vencidos = gastos[(gastos["DIAS_PASADOS"] > 7) & (gastos["FECHA DE COBRO"].replace('', np.nan).isna())]

    c1, c2, c3 = st.columns(3)
    c1.metric("🚨 Finiquitos Duplicados", f"{len(dup_finiquitos)} alertas", delta_color="inverse")
    c2.metric("💀 Cheques Cancelados COBRADOS", f"{len(fraudes_cancelados)} alertas", delta_color="inverse")
    c3.metric("🚫 Gratificaciones Halladas", f"{len(gratif)} alertas")

    with st.expander("🚨 ALERTA CRÍTICA: Cheques Cancelados que presentan movimiento en Banco"):
        if not fraudes_cancelados.empty:
            st.error("¡ALERTA DE FRAUDE! Los siguientes folios están marcados como CANCELADOS, pero aparecen cobrados en el Banco:")
            st.dataframe(fraudes_cancelados[["EMPRESA", "FOLIO", "MONTO", "BENEFICIARIO", "¿COBRADO_EN_BANCO?", "OBSERVACIONES"]], width='stretch')
        else: st.success("Todo en orden. Ningún cheque cancelado presenta flujos en el banco.")

    with st.expander("👤 ALERTA: Finiquitos Duplicados a la misma persona"):
        if not dup_finiquitos.empty:
            st.warning("Se detectó duplicidad de Finiquitos para un mismo operador/colaborador:")
            st.dataframe(dup_finiquitos.sort_values(by="BENEFICIARIO")[["EMPRESA", "FECHA DE EMISIÓN", "FOLIO", "MONTO", "BENEFICIARIO", "CAUSA DE LA BAJA"]], width='stretch')
        else: st.success("Sin duplicados. Cada finiquito corresponde a un beneficiario único.")

    with st.expander("🏦 ALERTA: Uso de Chequeras Inactivas (BBVA)"):
        if not alertas_bbva.empty:
            st.error("Se detectaron movimientos en cuentas congeladas de BBVA:")
            st.dataframe(alertas_bbva, width='stretch')
        else: st.success("Cuentas inactivas sin anomalías.")

    with st.expander("⏳ Control de Gastos por Comprobar (> 1 Semana Pendientes)"):
        if not gastos_vencidos.empty:
            st.warning("Efectivo flotante sin comprobantes fiscales:")
            st.dataframe(gastos_vencidos[["EMPRESA", "FOLIO", "BENEFICIARIO", "MONTO", "DIAS_PASADOS"]], width='stretch')
        else: st.success("Gastos por comprobar al corriente.")

# ==========================================
# MÓDULO: DASHBOARD Y GRÁFICOS
# ==========================================
elif opcion == "📊 Dashboard y Gráficos":
    st.title("📊 Métricas Ejecutivas e Indicadores Mensuales")
    st.markdown("---")
    
    df_m = st.session_state["MASTER_CHEQUES"].copy()
    df_m = sanitizar_para_streamlit(df_m)
    
    if df_m.empty:
        st.info("Carga datos en el sistema para habilitar las gráficas interactivas.")
        st.stop()
        
    df_m["MONTO"] = pd.to_numeric(df_m["MONTO"], errors="coerce").fillna(0)
    df_m["FECHA DE EMISIÓN"] = pd.to_datetime(df_m["FECHA DE EMISIÓN"], errors='coerce')
    df_m["MES"] = df_m["FECHA DE EMISIÓN"].dt.strftime('%Y-%m').fillna("Sin Fecha")

    emp_filtro = st.multiselect("Filtrar Gráficos por Empresa:", ["TODAS"] + LISTA_EMPRESAS, default=["TODAS"])
    if "TODAS" not in emp_filtro:
        df_m = df_m[df_m["EMPRESA"].isin(emp_filtro)]

    l1, l2 = st.columns(2)
    with l1:
        st.subheader("💰 Monto de Cheques por Concepto")
        conceptos_monto = df_m.groupby("CONCEPTO")["MONTO"].sum()
        st.bar_chart(conceptos_monto)
    with l2:
        st.subheader("🏢 Volumen de Emisión por Empresa")
        empresa_volumen = df_m.groupby("EMPRESA")["MONTO"].sum()
        st.bar_chart(empresa_volumen)

    l3, l4 = st.columns(2)
    with l3:
        st.subheader("📅 Comportamiento de Gastos Mensual ($)")
        mensual_monto = df_m.groupby("MES")["MONTO"].sum()
        st.line_chart(mensual_monto)
    with l4:
        st.subheader("🏆 Top 5 Beneficiarios con Mayor Flujo ($)")
        top_benef = df_m.groupby("BENEFICIARIO")["MONTO"].sum().sort_values(ascending=False).head(5)
        st.bar_chart(top_benef)
