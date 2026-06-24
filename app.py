import streamlit as st
import pandas as pd
from datetime import datetime
import os
import requests

# --- REGISTRO HORA DE INICIO (MODIFICADO: OBJETO NATIVO DATETIME PARA RESTA MATEMÁTICA)
if "hora_inicio" not in st.session_state:
    st.session_state["hora_inicio"] = datetime.now()
# -----------------------------------------------------

st.set_page_config(page_title="Precio de alimentos por territorialidades alimentarias", page_icon="🌽", layout="centered")

# --- CONVERSIÓN DE UNIDADES COMERCIALES A GRAMOS ---
DICCIONARIO_UNIDADES_GRAMOS = {
    "Kilogramo (kg)": 1000,
    "Libra (500g)": 500,
    "Litro (L)": 1000,
    "Medio litro (500ml)": 500,
    "Unidad": 100,           
    "Atado": 300,                       
}

# --- LECTURA DIRECTA DE DATOS UNIFICADOS ---
try:
    df_muni = pd.read_excel("municipios.xlsx")
    df_muni.columns = df_muni.columns.str.strip()
    for col in df_muni.columns:
        if "territorialidad" in col.lower():
            df_muni = df_muni.rename(columns={col: "Territorialidad_Estandar"})

    df_alimentos = pd.read_excel("alimentos.xlsx")
    df_alimentos.columns = [str(c).strip() for c in df_alimentos.columns]

    # SOLUCIÓN DE RAÍZ: Limpieza masiva de ArrowString en la memoria del servidor
    for col in df_muni.columns:
        df_muni[col] = df_muni[col].astype(str).str.replace("[", "").str.replace("]", "").str.replace("'", "").str.strip()

    for col in df_alimentos.columns:
        df_alimentos[col] = df_alimentos[col].astype(str).str.replace("[", "").str.replace("]", "").str.replace("'", "").str.strip()

    # Normalizar los nombres de las columnas clave adaptado a tu estructura
    for col in df_alimentos.columns:
        col_limpia = col.lower().replace("\n", " ").replace("\r", " ")
        if "territorialidad" in col_limpia:
            df_alimentos = df_alimentos.rename(columns={col: "Territorialidad_Estandar"})
        if "grupo" in col_limpia and "sub" not in col_limpia:
            df_alimentos = df_alimentos.rename(columns={col: "Grupo"})
        if "subgrupo" in col_limpia:
            df_alimentos = df_alimentos.rename(columns={col: "Subgrupo"})
        if "alimento" in col_limpia:
            df_alimentos = df_alimentos.rename(columns={col: "Alimento"})
        if "bruto" in col_limpia:
            df_alimentos = df_alimentos.rename(columns={col: "Persona gr/dia (bruto)"})
        if "persona" in col_limpia and "semana" in col_limpia:
            df_alimentos = df_alimentos.rename(columns={col: "Persona gr/semana"})
        if "hogar" in col_limpia and "semana" in col_limpia:
            df_alimentos = df_alimentos.rename(columns={col: "Hogar kg/semana"})
        if "descri" in col_limpia:
            df_alimentos = df_alimentos.rename(columns={col: "Descripción"})
            
except Exception as e:
    st.error(f"❌ Error leyendo archivos Excel. Detalle: {e}")
    st.stop()

st.title("🌽Precio de alimentos por territorialidades alimentarias")
st.write("Formulario para la recolección de precios de alimentos de las 13 territorialidades colombianas discriminados por tienda, plaza de mercado y supermercado. Indicaciones:")
st.write("1. Ingrese a la aplicación con su email institucional.")
st.write("2. Registre los precios de los alimentos para los tres expendios.")
st.write("4. Escriba cualquier observación que tenga sobre la recolección de datos o la aplicación.")
st.write("3. Guarde y envie la encuesta completa.")

st.subheader("1. Validación Institucional y Ubicación")
correo_estudiante = st.text_input("Correo Institucional (@udea.edu.co) *", key="corr_est").strip().lower()
es_correo_valido = correo_estudiante.endswith("@udea.edu.co") and len(correo_estudiante) > 12

if not es_correo_valido and correo_estudiante != "":
    st.error("❌ Acceso Denegado. Debes ingresar un correo institucional válido terminado en @udea.edu.co")

nombre_estudiante = ""
depto_sel = "---"
municipio_sel = "---"
territorialidad = ""

if es_correo_valido:
    st.success("✅ Correo institucional validado con éxito.")
    nombre_estudiante = st.text_input("Nombre Completo del Estudiante *", key="nom_est").strip()
    
    lista_departamentos = sorted(df_muni["Departamento"].unique())
    depto_sel = st.selectbox("Selecciona el Departamento *", ["---"] + lista_departamentos, key="dep_box")

    if depto_sel != "---":
        df_filtrado_muni = df_muni[df_muni["Departamento"] == depto_sel]
        lista_municipios = sorted(df_filtrado_muni["Municipio"].unique())
        municipio_sel = st.selectbox("Selecciona el Municipio *", ["---"] + lista_municipios, key="mun_box")
        if municipio_sel != "---":
            # Extraemos la territorialidad como un valor purificado de texto plano
            territorialidad = str(df_filtrado_muni.loc[df_filtrado_muni["Municipio"] == municipio_sel, "Territorialidad_Estandar"].values[0]).strip()
            st.success(f"📍 **Territorialidad de tu región:** {territorialidad}")
else:
    st.info("🔒 Por favor, ingresa tu correo institucional para desbloquear el formulario de ubicación y precios.")

st.divider()
# --- SECCIÓN 2: ALIMENTOS ACTIVOS ---
if es_correo_valido and territorialidad != "":
    df_region = df_alimentos[df_alimentos["Territorialidad_Estandar"] == territorialidad]
    
    df_region["Persona gr/dia (bruto)"] = pd.to_numeric(df_region["Persona gr/dia (bruto)"], errors='coerce').fillna(0.0)
    df_region["Persona gr/semana"] = pd.to_numeric(df_region["Persona gr/semana"], errors='coerce').fillna(0.0)
    df_region["Hogar kg/semana"] = pd.to_numeric(df_region["Hogar kg/semana"], errors='coerce').fillna(0.0)
    
    df_lista_obligatoria = df_region[df_region["Persona gr/dia (bruto)"] > 0].drop_duplicates(subset=["Alimento"])
    
    if not df_lista_obligatoria.empty:
        st.subheader(f"2. Registro Obligatorio de Alimentos ({len(df_lista_obligatoria)} productos)")
        st.warning("⚠️ Debes completar cada alimento. Si no encontraste un producto en algún establecimiento, déjalo en 0.")
        
        datos_capturados = {}
        
        for index, fila in df_lista_obligatoria.iterrows():
            alimento_nombre = fila["Alimento"]
            grupo = fila.get("Grupo", "General")
            subgrupo = fila.get("Subgrupo", "General")
            desc_texto = str(fila.get("Descripción", "Variedad específica del producto."))
            
            n_dia_bruto = round(float(fila["Persona gr/dia (bruto)"]), 1) 
            n_sem_persona = round(float(fila["Persona gr/semana"]), 1)
            n_sem_hogar = round(float(fila["Hogar kg/semana"]), 1)
            
            if desc_texto == "nan" or desc_texto == "" or desc_texto == "None":
                desc_texto = f"Variedad correspondiente a {alimento_nombre} ({subgrupo})."

            st.markdown(f"### 🍲 {alimento_nombre}")
            st.markdown(f"**Persona:** {n_dia_bruto} g/día (Bruto) | **Hogar:** {n_sem_hogar} kg/semana")
            st.caption(f"🔬 *Descripción:* {desc_texto}")
            
            col1, col2, col3, col4 = st.columns([1.5, 1, 1, 1])
            with col1:
                u_sel = st.selectbox(f"Unidad ({alimento_nombre})", ["---"] + list(DICCIONARIO_UNIDADES_GRAMOS.keys()), key=f"uni_{index}")
            with col2:
                p_tienda = st.number_input(f"Tienda ($)", min_value=0, step=500, value=0, key=f"tien_{index}")
            with col3:
                p_super = st.number_input(f"Super ($)", min_value=0, step=500, value=0, key=f"sup_{index}")
            with col4:
                p_plaza = st.number_input(f"Plaza ($)", min_value=0, step=500, value=0, key=f"plaz_{index}")
            st.markdown("---")
            
            datos_capturados[alimento_nombre] = {
                "unidad": u_sel, "tienda": p_tienda, "super": p_super, "plaza": p_plaza,
                "n_dia_bruto": n_dia_bruto, "n_sem_persona": n_sem_persona, "n_sem_hogar": n_sem_hogar,
                "subgrupo": subgrupo, "grupo": grupo
            }
        # --- REQUERIMIENTO 2: AGREGAR CAJÓN DE OBSERVACIONES EN INTERFAZ ---
        st.subheader("3. Observaciones del Formulario")
        observaciones_usuario = st.text_area(
            "Escribe aquí comentarios adicionales sobre los precios, novedades o inconvenientes en los establecimientos:", 
            placeholder="Ej. Algunos productos específicos no se encontraban disponibles en la plaza de mercado...",
            key="txt_observaciones"
        ).strip()

        st.markdown("<br>", unsafe_allow_html=True)
        enviar = st.button("💾 GUARDAR ENCUESTA COMPLETA", type="primary")

        if enviar:
            if not nombre_estudiante:
                st.error("❌ Por favor, digita tu nombre completo en la Sección 1.")
            elif depto_sel == "---" or municipio_sel == "---":
                st.error("❌ Por favor, selecciona el departamento y municipio en la Sección 1.")
            else:
                errores_validacion = False
                filas_para_guardar = []
                
                # --- REQUERIMIENTO 1: MÁQUINA DE TIEMPO AUTOMÁTICA ---
                hora_fin_dt = datetime.now()
                hora_inicio_dt = st.session_state["hora_inicio"]
                
                # Operación matemática directa para extraer la diferencia total en segundos
                tiempo_total_segundos = int((hora_fin_dt - hora_inicio_dt).total_seconds())
                
                # Formatear marcas temporales en texto estándar para almacenamiento estructurado
                str_hora_inicio = hora_inicio_dt.strftime("%Y-%m-%d %H:%M:%S")
                str_fecha_hora_envio = hora_fin_dt.strftime("%Y-%m-%d %H:%M:%S")
                
                id_encuesta = hora_fin_dt.strftime("%Y%m%d%H%M%S")

                for alim, inputs in datos_capturados.items():
                    if inputs["unidad"] == "---":
                        st.error(f"❌ Te falta seleccionar la Unidad de Medida para el alimento: **{alim}**.")
                        errores_validacion = True
                        break
                    if inputs["tienda"] == 0 and inputs["super"] == 0 and inputs["plaza"] == 0:
                        st.error(f"❌ Debes registrar al menos un precio válido para el alimento: **{alim}**.")
                        errores_validacion = True
                        break
                    
                    g_comerciales = DICCIONARIO_UNIDADES_GRAMOS[inputs["unidad"]]
                    g_dia_bruto = inputs["n_dia_bruto"]
                    g_sem_persona = inputs["n_sem_persona"]
                    g_sem_hogar = inputs["n_sem_hogar"] * 1000  
                    
                    c_dia_tienda = (inputs["tienda"] / g_comerciales) * g_dia_bruto if inputs["tienda"] > 0 else 0
                    c_dia_super = (inputs["super"] / g_comerciales) * g_dia_bruto if inputs["super"] > 0 else 0
                    c_dia_plaza = (inputs["plaza"] / g_comerciales) * g_dia_bruto if inputs["plaza"] > 0 else 0
                    
                    c_hogar_tienda = (inputs["tienda"] / g_comerciales) * g_sem_hogar if inputs["tienda"] > 0 else 0
                    c_hogar_super = (inputs["super"] / g_comerciales) * g_sem_hogar if inputs["super"] > 0 else 0
                    c_hogar_plaza = (inputs["plaza"] / g_comerciales) * g_sem_hogar if inputs["plaza"] > 0 else 0
                    
                    # Se anexan al final del arreglo 'str_hora_inicio', 'str(tiempo_total_segundos)' y 'observaciones_usuario'
                    filas_para_guardar.append([
                        id_encuesta, str_fecha_hora_envio, correo_estudiante,
                        nombre_estudiante, depto_sel, municipio_sel, territorialidad, 
                        inputs["grupo"], inputs["subgrupo"], alim, inputs["unidad"], str(g_comerciales),
                        str(inputs["tienda"]), str(inputs["super"]), str(inputs["plaza"]),
                        str(g_dia_bruto), str(g_sem_persona), str(inputs["n_sem_hogar"]),
                        str(round(c_dia_tienda, 2)), str(round(c_dia_super, 2)), str(round(c_dia_plaza, 2)),
                        str(round(c_hogar_tienda, 2)), str(round(c_hogar_super, 2)), str(round(c_hogar_plaza, 2)),
                        str_hora_inicio, str(tiempo_total_segundos), observaciones_usuario
                    ])
                
                if not errores_validacion:
                    try:
                        # TU ENLACE OFICIAL DE CARTERO DE GOOGLE APPS SCRIPT VINCULADO DEFINITIVAMENTE:
                        url_cartero = "https://script.google.com/macros/s/AKfycbwSOZaezhfjyP5c4LUNMkACAZ02urRqmhpDClcTTvhAmkWHnXgM6LpY2Ld442BO5GKK/exec"
                        
                        respuesta = requests.post(url_cartero, json=filas_para_guardar)
                        if respuesta.status_code == 200:
                            st.success("🎉 ¡Excelente! Los precios fueron guardados y enviados")
                        else:
                            st.error("❌ Error al transmitir los datos. Verifica la implementación del Script.")
                    except Exception as err:
                        st.error(f"❌ Error de red: {err}")
    else:
        st.warning(f"⚠️ No hay alimentos con necesidad en bruto mayor a 0 gramos para la territorialidad '{territorialidad}'.")
elif es_correo_valido:
    st.info("💡 Por favor, define el departamento y municipio para generar la lista de alimentos de tu territorialidad.")


    
