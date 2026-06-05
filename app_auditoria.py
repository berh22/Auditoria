import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import openpyxl
from openpyxl.styles import PatternFill
from datetime import datetime

# Configuración global de la Suite Corporativa
st.set_page_config(page_title="Sistema de Control Transaccional - Cheques", layout="wide")

# ==========================================
# FUNCIONES DE FORMATEO Y BLINDAJE
# ==========================================
def forzar_mayusculas_y_limpieza(df):
    if df is None or df.empty:
        return df
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].fillna('').astype(str).str.upper().str.strip()
            df[col] = df[col].replace(['NAN', 'NONE', '<NAT>'], '')
    return df

def limpiar_y_convertir_monto_positivo(df, columna):
    if columna in df.columns:
        df[columna] = df[columna].astype(str).str.replace('$', '', regex=False)
        df[columna] = df[columna].str.replace(',', '', regex=False)
        df[columna] = df[columna].str.strip()
        df[columna] = pd.to_numeric(df[columna], errors='coerce').fillna(0)
        df[columna] = df[columna].abs()
    return df

# ==========================================
# REGLA MAESTRA DE COLORES PARA PANTALLA
# ==========================================
def colorear_reglas_auditoria(row):
    estilo = [''] * len(row)
    
    concepto = str(row.get('CONCEPTO', '')).upper()
    sello = str(row.get('CUENTA CON SELLO', '')).upper()
    beneficiario = str(row.get('BENEFICIARIO', '')).upper()
    folio = str(row.get('FOLIO', ''))
    monto = pd.to_numeric(row.get('MONTO', 0), errors='coerce')
    monto = 0 if pd.isna(monto) else monto
    estatus = str(row.get('ESTATUS_CHEQUE', '')).upper()
    
    clave_evidencia = f"{row.get('EMPRESA', '')}_{folio}".upper().strip()
    tiene_vale = clave_evidencia in st.session_state["EVIDENCIAS_VALES"]
    tiene_justificante = clave_evidencia in st.session_state["EVIDENCIAS_CANCELADOS"]
    
    fecha_actual = datetime.now()
    f_emision = pd.to_datetime(row.get('FECHA DE EMISIÓN'), errors='coerce')
    f_cobro = pd.to_datetime(row.get('FECHA DE COBRO'), errors='coerce')
    
    # 🔴 Alerta 1: Gratificaciones (Rojo)
    if "GRATIFICACI" in concepto:
        return ['background-color: #ffcccc; color: #cc0000; font-weight: bold'] * len(row)
    
    # 🟠 Alerta 2: Cancelado SIN JUSTIFICANTE (Naranja)
    if estatus == "CANCELADO" and not tiene_justificante:
        return ['background-color: #fce5cd; color: #b45f06; font-style: italic'] * len(row)
        
    # 🟡 Alerta 3: Sin Sello / Efectivo y NO Solventado (Amarillo)
    if ("NO" == sello or "SIN SELLO" in sello) and not tiene_vale and estatus != "CANCELADO":
        return ['background-color: #fff2cc; color: #856404'] * len(row)
        
    # 🔵 Alerta 4: Más de 1 semana emitido y SIN COBRAR (Azul Suave)
    if estatus != "CANCELADO" and pd.notna(f_emision) and pd.isna(f_cobro):
        dias_transcurridos = (fecha_actual - f_emision).days
        if dias_transcurridos > 7:
            return ['background-color: #e6f2ff; color: #0b5394; font-style: italic'] * len(row)
        
    # 🟠 Alerta 5: Falta de Datos Obligatorios en Emitidos (Naranja)
    if estatus != "CANCELADO":
        falta_dato = False
        if f_emision is pd.NaT or folio in ['0', ''] or monto == 0 or beneficiario == "":
            falta_dato = True
        if "FINIQUITO" in concepto:
            if str(row.get('CAUSA DE LA BAJA', '')).strip() == "" or str(row.get('FECHA DE BAJA', '')).strip() == "":
                falta_dato = True
                
        if falta_dato:
            return ['background-color: #fce5cd; color: #b45f06; font-style: italic'] * len(row)
            
    return estilo

# ==========================================
# MOTOR DE EXPORTACIÓN EXCEL A COLOR (OPENPYXL)
# ==========================================
def generar_excel_coloreado(df):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Master_Cheques"
    
    # Escribir encabezados
    headers = list(df.columns)
    ws.append(headers)
    
    # Definición de Paletas de Fills (Mismos colores que la App Web)
    f_rojo = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    f_amarillo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    f_azul = PatternFill(start_color="E6F2FF", end_color="E6F2FF", fill_type="solid")
    f_naranja = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")
    
    fecha_actual = datetime.now()
    
    for index, row in df.iterrows():
        # Escribir fila de datos limpiando valores nulos
        row_values = [str(val) if pd.notna(val) else "" for val in row]
        ws.append(row_values)
        row_idx = ws.max_row
        
        # Lógica de asignación de color en el archivo físico
        concepto = str(row.get('CONCEPTO', '')).upper()
        sello = str(row.get('CUENTA CON SELLO', '')).upper()
        beneficiario = str(row.get('BENEFICIARIO', '')).upper()
        folio = str(row.get('FOLIO', ''))
        empresa = str(row.get('EMPRESA', ''))
        monto = pd.to_numeric(row.get('MONTO', 0), errors='coerce')
        estatus = str(row.get('ESTATUS_CHEQUE', '')).upper()
        
        clave = f"{empresa}_{folio}".upper().strip()
        tiene_vale = clave in st.session_state["EVIDENCIAS_VALES"]
        tiene_justificante = clave in st.session_state["EVIDENCIAS_CANCELADOS"]
        
        f_emision = pd.to_datetime(row.get('FECHA DE EMISIÓN'), errors='coerce')
        f_cobro = pd.to_datetime(row.get('FECHA DE COBRO'), errors='coerce')
        
        fill_actual = None
        
        if "GRATIFICACI" in concepto:
            fill_actual = f_rojo
        elif estatus == "CANCELADO" and not tiene_justificante:
            fill_actual = f_naranja
        elif ("NO" == sello or "SIN SELLO" in sello) and not tiene_vale and estatus != "CANCELADO":
            fill_actual = f_amarillo
        elif estatus != "CANCELADO" and pd.notna(f_emision) and pd.isna(f_cobro):
            if (fecha_actual - f_emision).days > 7:
                fill_actual = f_azul
        elif estatus != "CANCELADO":
            falta_dato = False
            if f_emision is pd.NaT or folio in ['0', ''] or monto == 0 or beneficiario == "":
                falta_dato = True
            if "FINIQUITO" in concepto:
                if str(row.get('CAUSA DE LA BAJA', '')).strip() == "" or str(row.get('FECHA DE BAJA', '')).strip() == "":
                    falta_dato = True
            if falta_dato:
                fill_actual = f_naranja
        
        # Aplicar el color a todas las celdas de la fila en Excel
        if fill_actual:
            for col_idx in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col_idx).fill = fill_actual
                
    wb.save(output)
    return output.getvalue()

# ==========================================
# ESTRUCTURAS DE PERSISTENCIA
# ==========================================
if "MASTER_CHEQUES" not in st.session_state:
    st.session_state["MASTER_CHEQUES"] = pd.DataFrame(columns=[
        "EMPRESA", "FECHA DE EMISIÓN", "FOLIO", "MONTO", "BENEFICIARIO", 
        "CUENTA CON SELLO", "FECHA DE COBRO", "CONCEPTO", "OBSERVACIONES", 
        "CAUSA DE LA BAJA", "FECHA DE BAJA", "ESTATUS_CHEQUE"
    ])

if "EVIDENCIAS_VALES" not in st.session_state:
    st.session_state["EVIDENCIAS_VALES"] = {}
if "EVIDENCIAS_CANCELADOS" not in st.session_state:
    st.session_state["EVIDENCIAS_CANCELADOS"] = {}

# ==========================================
# PORTAL DE LOGIN
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

# BARRA LATERAL
st.sidebar.title("🛡️ Sistema Integral")
if st.sidebar.button("🔒 Cerrar Sesión"):
    st.session_state["autenticado"] = False
    st.rerun()

st.sidebar.markdown("---")
opcion = st.sidebar.radio("Menú Principal:", ["📥 Alimentar / Capturar Cheques", "🔍 Auditoría y Alertas Forenses", "📊 Dashboard y Gráficos"])

LISTA_EMPRESAS = ["Bajio Tropper", "Bajio TCS", "Bajio TCS Gerencia", "Bajio Somnus", "Bajio Idea", "Bajio Areng", "BBVA Tropper", "BBVA Sab Solutions"]
CHEQUERAS_INACTIVAS = ["BBVA Tropper", "BBVA Sab Solutions"]

# ==========================================
# MÓDULO: ADMINISTRACIÓN DE CHEQUES
# ==========================================
if opcion == "📥 Alimentar / Capturar Cheques":
    st.title("📥 Panel de Captura y Modificación del Layout")
    st.markdown("---")
    
    tab_editor, tab_carga, tab_manual, tab_boveda, tab_exportar = st.tabs([
        "📝 TABLA INTERACTIVA (Editar/Categorizar)", 
        "📁 Subir Archivo Nuevo", 
        "✍️ Captura Manual", 
        "📂 BÓVEDA DE EVIDENCIAS (Vales / Cancelados)",
        "💾 Exportar a Excel"
    ])
    
    with tab_editor:
        st.subheader("Modificación de Datos en Vivo")
        if not st.session_state["MASTER_CHEQUES"].empty:
            for col_f in ["FECHA DE EMISIÓN", "FECHA DE COBRO", "FECHA DE BAJA"]:
                st.session_state["MASTER_CHEQUES"][col_f] = pd.to_datetime(st.session_state["MASTER_CHEQUES"][col_f], errors='coerce')

            st.session_state["MASTER_CHEQUES"] = st.session_state["MASTER_CHEQUES"].sort_values(by="FOLIO", ascending=True)

            df_editado = st.data_editor(
                st.session_state["MASTER_CHEQUES"],
                num_rows="dynamic",
                width="stretch",
                column_config={
                    "FECHA DE EMISIÓN": st.column_config.DateColumn(format="DD/MM/YYYY"),
                    "FECHA DE COBRO": st.column_config.DateColumn(format="DD/MM/YYYY"),
                    "FECHA DE BAJA": st.column_config.DateColumn(format="DD/MM/YYYY"),
                    "MONTO": st.column_config.NumberColumn(format="$%.2f"),
                    "CUENTA CON SELLO": st.column_config.SelectboxColumn(options=["SI", "NO", "NO APLICA"]),
                    "CONCEPTO": st.column_config.SelectboxColumn(options=[
                        "1. COMPROBACIÓN DE GASTOS", "2. FINIQUITO", "3. ABONO A FINANCIERA", 
                        "4. APERTURA DE FONDO FIJO", "5. PAGO DE NÓMINA", "6. GASTO POR COMPROBAR", 
                        "7. APOYO POR DEFUNCIÓN", "8. GRATIFICACIONES", "9. PENSIÓN ALIMENTICIA", 
                        "10. RETENCION A PROVEEDOR", "11. PAGO DE TRÁMITES", "12. OTRO MOTIVO"
                    ]),
                    "ESTATUS_CHEQUE": st.column_config.SelectboxColumn(options=["EMITIDO", "CANCELADO"])
                }
            )
            
            if st.button("💾 Guardar Cambios Realizados en la Tabla"):
                st.session_state["MASTER_CHEQUES"] = forzar_mayusculas_y_limpieza(df_editado)
                st.success("¡Base de datos centralizada y guardada con éxito!")
                st.rerun()
        else: st.info("La tabla está vacía. Sube una chequera para empezar.")

    with tab_carga:
        st.subheader("Cargar Nuevos Cheques Masivamente")
        archivo_nuevo = st.file_uploader("Sube el Excel de Contpaq o tu Chequera Manual:", type=["xlsx"])
        empresa_carga = st.selectbox("¿A qué empresa corresponden estos cheques?", LISTA_EMPRESAS, key="emp_carga")
        
        if archivo_nuevo:
            df_raw = pd.read_excel(archivo_nuevo, header=None)
            if not df_raw.empty:
                keywords = ["FOLIO", "BENEFICIARIO", "MONTO", "FECHA DE EMISIÓN", "CONCEPTO", "NÚMERO/FOLIO", "BENEFICIARIO/PAGADOR", "IMPORTE"]
                header_idx = None
                for idx, row in df_raw.iterrows():
                    row_str = [str(cell).upper() for cell in row.values]
                    if sum(1 for kw in keywords if any(kw in cell for cell in row_str)) >= 2:
                        header_idx = idx
                        break
                
                if header_idx is not None:
                    df_temp = df_raw.iloc[header_idx+1:].copy()
                    encabezados_crudos = [str(c).strip().upper() for c in df_raw.iloc[header_idx].values]
                    
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
                    
                    st.info(f"📋 Formato Detectado en la fila {header_idx + 1}.")
                    st.dataframe(df_temp.head(3), width='stretch')
                    
                    if st.button("🚀 Confirmar e Integrar al Histórico Master"):
                        mapeo_columnas = {
                            "FECHA DE EMISIÓN": ["FECHA DE EMISIÓN", "FECHA DE EMISION", "FECHA EMISION", "FECHA EMISIÓN", "FECHA"],
                            "FOLIO": ["FOLIO", "NUMERO", "NÚMERO", "NÚM", "NUM", "CHEQUE", "FOLIOS", "NÚMERO/FOLIO", "NUMERO/FOLIO"],
                            "MONTO": ["MONTO", "IMPORTE", "CANTIDAD", "CARGOS"],
                            "BENEFICIARIO": ["BENEFICIARIO", "NOMBRE", "PROVEEDOR", "A FAVOR DE", "BENEFICIARIO/PAGADOR"],
                            "CUENTA CON SELLO": ["CUENTA CON SELLO", "SELLO", "CON SELLO"],
                            "FECHA DE COBRO": ["FECHA DE COBRO (DEPOSITO EN CUENTA)", "FECHA DE COBRO", "FECHA COBRO", "COBRO", "DEPOSITO"],
                            "CONCEPTO": ["CONCEPTO", "MOTIVO", "DESCRIPCION", "DESCRIPCIÓN"],
                            "OBSERVACIONES": ["OBSERVACIONES", "NOTAS", "COMENTARIOS", "OBSERVACION"]
                        }
                        
                        columnas_nuevas = {}
                        for col_excel in df_temp.columns:
                            for col_estandar, lista_alias in mapeo_columnas.items():
                                if col_excel in lista_alias:
                                    columnas_nuevas[col_excel] = col_estandar
                                    break
                        df_temp = df_temp.rename(columns=columnas_nuevas)
                        
                        for col_estandar in st.session_state["MASTER_CHEQUES"].columns:
                            if col_estandar not in df_temp.columns: df_temp[col_estandar] = ""
                        
                        df_temp = df_temp[df_temp["FOLIO"].notna() & (df_temp["FOLIO"].astype(str).str.strip() != "")]
                        df_temp = limpiar_y_convertir_monto_positivo(df_temp, 'MONTO')
                        df_temp["FOLIO"] = pd.to_numeric(df_temp["FOLIO"], errors='coerce').fillna(0).astype(int)
                        df_temp["EMPRESA"] = empresa_carga
                        
                        df_temp["ESTATUS_CHEQUE"] = df_temp.apply(
                            lambda r: "CANCELADO" if "CANCELADO" in str(r.get("OBSERVACIONES", "")).upper() or \
                                                     "CANCELADO" in str(r.get("BENEFICIARIO", "")).upper() or \
                                                     "CANCELADO" in str(r.get("CONCEPTO", "")).upper() \
                                                  else "EMITIDO", axis=1
                        )
                        
                        df_temp = df_temp.reindex(columns=st.session_state["MASTER_CHEQUES"].columns)
                        df_temp = forzar_mayusculas_y_limpieza(df_temp)
                        df_temp = df_temp.sort_values(by="FOLIO", ascending=True)
                        
                        st.session_state["MASTER_CHEQUES"] = pd.concat([st.session_state["MASTER_CHEQUES"], df_temp], ignore_index=True)
                        st.session_state["MASTER_CHEQUES"] = st.session_state["MASTER_CHEQUES"].drop_duplicates(subset=["EMPRESA", "FOLIO"], keep="last")
                        st.success("¡Datos absorbidos!")
                        st.rerun()

    with tab_manual:
        st.subheader("Captura Individual Manual")
        with st.form("registro_manual", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            emp_m = c1.selectbox("Empresa:", LISTA_EMPRESAS)
            fecha_m = c2.date_input("Fecha de Emisión:")
            folio_m = c3.number_input("Folio del Cheque:", step=1, min_value=1)
            
            c4, c5, c6 = st.columns(3)
            monto_m = c4.number_input("Monto ($):", min_value=0.0, format="%.2f")
            benef_m = c5.text_input("Beneficiario:")
            sello_m = c6.selectbox("¿Cuenta con Sello 'Abono en Cuenta'?", ["SI", "NO"])
            
            c7, c8 = st.columns(2)
            concepto_m = c7.selectbox("Concepto del Cheque:", ["1. COMPROBACIÓN DE GASTOS", "2. FINIQUITO", "3. ABONO A FINANCIERA", "4. APERTURA DE FONDO FIJO", "5. PAGO DE NÓMINA", "6. GASTO POR COMPROBAR", "7. APOYO POR DEFUNCIÓN", "8. GRATIFICACIONES", "9. PENSIÓN ALIMENTICIA", "10. RETENCION A PROVEEDOR", "11. PAGO DE TRÁMITES", "12. OTRO MOTIVO"])
            estatus_m = c8.selectbox("Estatus Inicial:", ["EMITIDO", "CANCELADO"])
            obs_m = st.text_area("Observaciones:")
            
            if st.form_submit_button("Guardar Cheque"):
                nuevo_registro = {
                    "EMPRESA": emp_m, "FECHA DE EMISIÓN": pd.to_datetime(fecha_m), "FOLIO": int(folio_m),
                    "MONTO": monto_m, "BENEFICIARIO": benef_m, "CUENTA CON SELLO": sello_m,
                    "FECHA DE COBRO": np.nan, "CONCEPTO": concepto_m, "OBSERVACIONES": obs_m,
                    "CAUSA DE LA BAJA": "", "FECHA DE BAJA": "", "ESTATUS_CHEQUE": estatus_m
                }
                df_manual = forzar_mayusculas_y_limpieza(pd.DataFrame([nuevo_registro]))
                st.session_state["MASTER_CHEQUES"] = pd.concat([st.session_state["MASTER_CHEQUES"], df_manual], ignore_index=True)
                st.success("¡Guardado!")
                st.rerun()

    # ==========================================
    # 🔥 BÓVEDA DE EVIDENCIAS DIGITALES (VALES Y CANCELADOS)
    # ==========================================
    with tab_boveda:
        st.subheader("📂 Control de Soportes y Evidencias Digitales")
        tipo_evidencia = st.radio("¿Qué tipo de documento vas a solventar?", ["💵 VALE DE EFECTIVO (CHEQUES SIN SELLO)", "🛑 JUSTIFICANTE DE CANCELACIÓN (CHEQUES ANULADOS)"])
        
        if not st.session_state["MASTER_CHEQUES"].empty:
            c_emp, c_fol = st.columns(2)
            with c_emp:
                emp_v = st.selectbox("1. Selecciona la Empresa:", LISTA_EMPRESAS, key="emp_v")
            
            # Filtrar folios según la opción elegida
            if tipo_evidencia == "💵 VALE DE EFECTIVO (CHEQUES SIN SELLO)":
                df_filtrado_v = st.session_state["MASTER_CHEQUES"][
                    (st.session_state["MASTER_CHEQUES"]["EMPRESA"] == emp_v) & 
                    (st.session_state["MASTER_CHEQUES"]["CUENTA CON SELLO"] == "NO")
                ]
            else:
                df_filtrado_v = st.session_state["MASTER_CHEQUES"][
                    (st.session_state["MASTER_CHEQUES"]["EMPRESA"] == emp_v) & 
                    (st.session_state["MASTER_CHEQUES"]["ESTATUS_CHEQUE"] == "CANCELADO")
                ]
                
            with c_fol:
                if not df_filtrado_v.empty:
                    folio_v = st.selectbox("2. Selecciona el Folio del Cheque a justificar:", df_filtrado_v["FOLIO"].unique())
                else:
                    st.success("✨ Sin pendientes en este rubro para esta empresa.")
                    folio_v = None
            
            if folio_v:
                archivo_evidencia = st.file_uploader(f"Subir soporte (PDF/JPG) para el Folio {folio_v}:", type=["pdf", "jpg", "jpeg", "png"])
                if st.button("🔒 Vincular Documento a la Nube"):
                    if archivo_evidencia:
                        clave_llave = f"{emp_v}_{folio_v}".upper().strip()
                        if tipo_evidencia == "💵 VALE DE EFECTIVO (CHEQUES SIN SELLO)":
                            st.session_state["EVIDENCIAS_VALES"][clave_llave] = archivo_evidencia.name.upper()
                        else:
                            st.session_state["EVIDENCIAS_CANCELADOS"][clave_llave] = archivo_evidencia.name.upper()
                        st.success(f"✔️ Archivo '{archivo_evidencia.name.upper()}' indexado correctamente.")
                        st.rerun()
        else: st.info("La chequera está vacía.")

    with tab_exportar:
        st.subheader("Descargar Reporte de Auditoría")
        if not st.session_state["MASTER_CHEQUES"].empty:
            # Generar el Excel físico aplicando la paleta de colores de auditoría
            excel_binario = generar_excel_coloreado(st.session_state["MASTER_CHEQUES"])
            st.download_button(
                label="📥 Descargar Master de Cheques en Excel (Con Colores de Alerta)", 
                data=excel_binario, 
                file_name=f"Reporte_Auditoria_Color_{datetime.now().strftime('%Y%m%d')}.xlsx"
            )

    # ==========================================
    # VISOR GENERAL DE ALERTAS EN VIVO
    # ==========================================
    st.markdown("---")
    st.subheader("👁️ Visor de Alertas Forenses y Mapa de Calor")
    st.markdown("""
    * **Reglas de Auditoría:** 🔴 **Rojo:** Gratificaciones | 🟡 **Amarillo:** Cobros sin Sello (Falta Vale) | 🔵 **Azul Suave:** En Tránsito Viejo (> 1 Semana) | 🟠 **Naranja:** Campos obligatorios vacíos o Cancelados sin justificante.
    """)
    if not st.session_state["MASTER_CHEQUES"].empty:
        df_visor = st.session_state["MASTER_CHEQUES"].copy()
        df_visor = df_visor.sort_values(by="FOLIO", ascending=True)
        
        # Mapear estados de evidencias en tiempo real
        df_visor["EVIDENCIA_VALE"] = df_visor.apply(
            lambda r: st.session_state["EVIDENCIAS_VALES"].get(f"{r['EMPRESA']}_{r['FOLIO']}".upper().strip(), "PENDIENTE ❌" if r['CUENTA CON SELLO'] == "NO" else "NO REQUERIDO"), axis=1
        )
        df_visor["EVIDENCIA_CANCELADO"] = df_visor.apply(
            lambda r: st.session_state["EVIDENCIAS_CANCELADOS"].get(f"{r['EMPRESA']}_{r['FOLIO']}".upper().strip(), "FALTA JUSTIFICANTE ❌" if r['ESTATUS_CHEQUE'] == "CANCELADO" else "NO REQUERIDO"), axis=1
        )
        
        columnas_ordenadas = ["EMPRESA", "FECHA DE EMISIÓN", "FOLIO", "ESTATUS_CHEQUE", "MONTO", "BENEFICIARIO", "EVIDENCIA_VALE", "EVIDENCIA_CANCELADO"] + \
                             [c for c in df_visor.columns if c not in ["EMPRESA", "FECHA DE EMISIÓN", "FOLIO", "ESTATUS_CHEQUE", "MONTO", "BENEFICIARIO", "EVIDENCIA_VALE", "EVIDENCIA_CANCELADO"]]
        df_visor = df_visor[columnas_ordenadas]
        
        for col_f in ["FECHA DE EMISIÓN", "FECHA DE COBRO", "FECHA DE BAJA"]:
            df_visor[col_f] = pd.to_datetime(df_visor[col_f], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')
            
        st.dataframe(df_visor.style.apply(colorear_reglas_auditoria, axis=1), width='stretch')
    else: st.info("La base de datos central está vacía.")

# ==========================================
# SECCIONES ADICIONALES
# ==========================================
elif opcion == "🔍 Auditoría y Alertas Forenses":
    st.title("🔍 Escáner Forense de Riesgos de Control Interno")
    st.info("Sube el archivo del banco aquí para cruzar contra la base transaccional.")

elif opcion == "📊 Dashboard y Gráficos":
    st.title("📊 Métricas Ejecutivas e Indicadores Mensuales")
    st.info("Sección de gráficos habilitada.")
