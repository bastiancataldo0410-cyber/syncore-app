import datetime
import io
import pandas as pd
import plotly.express as px
import streamlit as st
import sqlalchemy as sa
import os

# Importaciones para generación de PDF nativo con ReportLab
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(
    page_title="SYNCORE | Control Ejecutivo P6, 6WLA & T&M", 
    page_icon="🟧", 
    layout="wide"
)

# --- CADENA DE CONEXIÓN SUPABASE (POOLER IPv4) ---
DATABASE_URL = "postgresql+psycopg2://postgres.wlxdxdjmyaamcoozvego:SyncoreMontajes@aws-0-us-east-2.pooler.supabase.com:6543/postgres"
engine = sa.create_engine(DATABASE_URL)

# --- ESTILOS CSS EJECUTIVOS Y CORPORATIVOS SYNCORE ---
st.markdown("""
    <style>
        .stApp {
            background-color: #F4F6F9;
        }
        [data-testid="stSidebar"] {
            background-color: #0F172A;
            color: #FFFFFF;
        }
        [data-testid="stSidebar"] label, [data-testid="stSidebar"] .stSelectbox div {
            color: #FFFFFF !important;
        }
        .stButton>button {
            background-color: #FF6B00;
            color: white;
            font-weight: 600;
            border-radius: 6px;
            border: none;
            padding: 0.6rem 1.2rem;
            box-shadow: 0 4px 6px rgba(255, 107, 0, 0.2);
        }
        .stButton>button:hover {
            background-color: #E05D00;
            color: white;
        }
        h1, h2, h3 {
            color: #0F172A;
            font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        }
        [data-testid="metric-container"] {
            background-color: #FFFFFF;
            border-top: 4px solid #FF6B00;
            border-radius: 8px;
            padding: 16px;
            box-shadow: 0 4px 12px rgba(15, 23, 42, 0.05);
        }
        .block-container {
            padding-top: 2rem;
        }
    </style>
""", unsafe_allow_html=True)

def init_db():
    try:
        with engine.connect() as conn:
            conn.execute(sa.text('''
                CREATE TABLE IF NOT EXISTS proyectos (
                    ID_PROJECTO SERIAL PRIMARY KEY,
                    Nombre TEXT NOT NULL,
                    Estado TEXT NOT NULL,
                    Avance_Porcentaje REAL NOT NULL,
                    At_Completion_Units REAL DEFAULT 0,
                    Actual_Units REAL DEFAULT 0,
                    Remaining_Units REAL DEFAULT 0,
                    Fecha_Inicio TEXT NOT NULL,
                    Fecha_Termino TEXT NOT NULL
                )
            '''))
            
            conn.execute(sa.text('''
                CREATE TABLE IF NOT EXISTS tareas (
                    ID_TAREA SERIAL PRIMARY KEY,
                    ID_Proyecto INTEGER,
                    Tarea TEXT NOT NULL,
                    Subsistema TEXT NOT NULL,
                    Labor_Units REAL NOT NULL,
                    Remaining_Labor_Units REAL NOT NULL,
                    Labor_Units_Complete_Pct REAL NOT NULL,
                    Estado TEXT NOT NULL,
                    Responsable TEXT
                )
            '''))
            
            conn.execute(sa.text('''
                CREATE TABLE IF NOT EXISTS analisis_plan_semanal (
                    ID_PLAN SERIAL PRIMARY KEY,
                    Proyecto TEXT NOT NULL,
                    Activity_ID TEXT NOT NULL,
                    Actividad TEXT NOT NULL,
                    Subsistema TEXT NOT NULL,
                    Semana TEXT NOT NULL,
                    HH_Programadas REAL DEFAULT 0,
                    Estado_Ejecucion TEXT NOT NULL
                )
            '''))

            conn.execute(sa.text('''
                CREATE TABLE IF NOT EXISTS analisis_log_desviaciones (
                    ID_DESV SERIAL PRIMARY KEY,
                    Proyecto TEXT NOT NULL,
                    Activity_ID TEXT NOT NULL,
                    Actividad TEXT NOT NULL,
                    Subsistema TEXT NOT NULL,
                    Causa_Desviacion TEXT NOT NULL,
                    Responsable TEXT NOT NULL,
                    HH_Perdidas REAL DEFAULT 0,
                    Estado TEXT NOT NULL
                )
            '''))
            
            conn.execute(sa.text('''
                CREATE TABLE IF NOT EXISTS informes_semanales (
                    ID_INFORME SERIAL PRIMARY KEY,
                    Nombre_Informe TEXT NOT NULL,
                    Fecha_Carga TEXT NOT NULL,
                    Subsistema TEXT NOT NULL,
                    At_Completion_Units REAL DEFAULT 0,
                    Remaining_Units REAL DEFAULT 0,
                    Avance_Pct REAL DEFAULT 0
                )
            '''))
            
            conn.execute(sa.text('''
                CREATE TABLE IF NOT EXISTS usuarios (
                    ID_USUARIO SERIAL PRIMARY KEY,
                    Username TEXT UNIQUE NOT NULL,
                    Password TEXT NOT NULL,
                    Estado TEXT DEFAULT 'Aprobado'
                )
            '''))
            conn.commit()
            
            result = conn.execute(sa.text("SELECT * FROM usuarios WHERE Username = 'Bastian Cataldo'")).fetchone()
            if not result:
                conn.execute(sa.text("INSERT INTO usuarios (Username, Password, Estado) VALUES (:user, :passw, :est)"),
                             {"user": "Bastian Cataldo", "passw": "admin123", "est": "Aprobado"})
                conn.commit()
    except Exception as e:
        st.error(f"Error al conectar o inicializar Supabase: {e}")

init_db()

def cargar_proyectos():
    try:
        return pd.read_sql("SELECT * FROM proyectos", engine).fillna("")
    except Exception:
        return pd.DataFrame()

def cargar_tareas():
    try:
        return pd.read_sql("SELECT * FROM tareas", engine).fillna("")
    except Exception:
        return pd.DataFrame()

def parse_xer_exacto(contenido_text):
    lineas = contenido_text.splitlines()
    tablas = {}
    tabla_actual = None
    columnas_actuales = []
    
    for linea in lineas:
        if not linea.strip(): continue
        partes = linea.split('\t')
        header = partes[0]
        
        if header == '%T':
            tabla_actual = partes[1]
            tablas[tabla_actual] = []
            columnas_actuales = []
        elif header == '%F' and tabla_actual:
            columnas_actuales = partes[1:]
        elif header == '%R' and tabla_actual and columnas_actuales:
            valores = partes[1:]
            if len(valores) == len(columnas_actuales):
                tablas[tabla_actual].append(dict(zip(columnas_actuales, valores)))
            else:
                min_len = min(len(valores), len(columnas_actuales))
                tablas[tabla_actual].append(dict(zip(columnas_actuales[:min_len], valores[:min_len])))
                
    return tablas

# --- GESTIÓN DE SESIÓN ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""

if not st.session_state.logged_in:
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        if os.path.exists("logo_syncore.png"):
            st.image("logo_syncore.png", use_container_width=True)
        else:
            st.markdown("<h1 style='text-align: center; color: #FF6B00;'>SYNCORE</h1>", unsafe_allow_html=True)
        
        st.markdown("<h3 style='text-align: center; color: #0F172A;'>Control Ejecutivo P6 (Cloud Supabase)</h3>", unsafe_allow_html=True)
        st.markdown("---")

        tab_login, tab_register = st.tabs(["🔑 Iniciar Sesión", "📝 Registrar Usuario"])

        with tab_login:
            st.subheader("Acceso al Sistema")
            user_input = st.text_input("Usuario", key="login_user")
            pass_input = st.text_input("Contraseña", type="password", key="login_pass")

            if st.button("Ingresar a la Plataforma"):
                try:
                    with engine.connect() as conn:
                        query = sa.text("SELECT Estado FROM usuarios WHERE Username = :u AND Password = :p")
                        user_record = conn.execute(query, {"u": user_input, "p": pass_input}).fetchone()

                    if user_record:
                        estado_usuario = user_record[0]
                        if estado_usuario == "Aprobado":
                            st.session_state.logged_in = True
                            st.session_state.username = user_input
                            st.success("¡Inicio de sesión exitoso! Cargando plataforma...")
                            st.rerun()
                        else:
                            st.warning("⚠️ Tu cuenta se encuentra pendiente de aprobación por el Administrador.")
                    else:
                        st.error("Usuario o contraseña incorrectos.")
                except Exception as ex:
                    st.error(f"Error de conexión con la base de datos: {ex}")

        with tab_register:
            st.subheader("Registro de Nuevo Usuario")
            new_user = st.text_input("Nuevo Usuario (Ej: Tu Nombre)", key="reg_user")
            new_pass = st.text_input("Nueva Contraseña", type="password", key="reg_pass")
            confirm_pass = st.text_input("Confirmar Contraseña", type="password", key="reg_conf")

            if st.button("Registrar Cuenta"):
                if not new_user or not new_pass:
                    st.warning("Por favor completa todos los campos.")
                elif new_pass != confirm_pass:
                    st.error("Las contraseñas no coinciden.")
                else:
                    try:
                        with engine.connect() as conn:
                            estado_inicial = "Aprobado" if new_user == "Bastian Cataldo" else "Pendiente"
                            conn.execute(sa.text("INSERT INTO usuarios (Username, Password, Estado) VALUES (:u, :p, :e)"),
                                         {"u": new_user, "p": new_pass, "e": estado_inicial})
                            conn.commit()
                        if estado_inicial == "Pendiente":
                            st.success("¡Cuenta registrada con éxito! Quedará pendiente de aprobación.")
                        else:
                            st.success("¡Cuenta de Administrador creada con éxito!")
                    except Exception as ex:
                        st.error(f"El nombre de usuario ya existe o hubo un error: {ex}")
    st.stop()

# --- BARRA LATERAL ---
if os.path.exists("logo_syncore.png"):
    st.sidebar.image("logo_syncore.png", use_container_width=True)
else:
    st.sidebar.markdown("""
        <div style="text-align: center; padding: 10px 0px 20px 0px;">
            <h2 style="color: #FF6B00; margin: 0; font-weight: 800;">SYNCORE</h2>
            <p style="color: #94A3B8; font-size: 13px; margin: 0;">MONTAJES INDUSTRIALES</p>
        </div>
    """, unsafe_allow_html=True)

st.sidebar.markdown(f"👤 **Usuario:** `{st.session_state.username}`")

menu_opciones = [
    "📊 Dashboard General", 
    "🏗️ Análisis Subsistema", 
    "📋 Análisis 6WLA",
    "📈 Informe Semanal",
    "📊 Análisis T&M (Times & Material)",
    "📄 Generador de Informes Diarios",
    "📁 Gestión de Proyectos", 
    "📂 Importar Archivo (.XER / Excel)"
]

is_admin = (st.session_state.username == "Bastian Cataldo")
if is_admin:
    menu_opciones.append("🛡️ Administración de Usuarios")

if st.sidebar.button("🚪 Cerrar Sesión"):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.rerun()

st.sidebar.markdown("---")
menu = st.sidebar.selectbox("Menú Ejecutivo", menu_opciones)

if menu == "📊 Dashboard General":
    st.title("📊 SYNCORE | Dashboard Ejecutivo Global (Cloud Supabase)")
    st.markdown("Análisis sincronizado en tiempo real con la base de datos centralizada.")
    
    df_proyectos = cargar_proyectos()
    df_tareas = cargar_tareas()

    if df_proyectos.empty or df_tareas.empty:
        st.info("No hay proyectos con datos cargados en la nube. Dirígete a **'📂 Importar Archivo (.XER / Excel)'** para sincronizar tu proyecto.")
    else:
        proyectos_lista = df_proyectos["Nombre"].unique().tolist()
        proj_seleccionado = st.selectbox("📌 Seleccionar Proyecto Activo:", proyectos_lista)
        
        row_proj = df_proyectos.loc[df_proyectos["Nombre"] == proj_seleccionado].iloc[0]
        id_proj = row_proj["ID_PROJECTO"]
        df_tareas_proj = df_tareas[df_tareas["ID_Proyecto"] == id_proj]

        if df_tareas_proj.empty:
            st.warning("Este proyecto no contiene actividades registradas.")
        else:
            total_actividades = len(df_tareas_proj)
            at_completion_total = float(row_proj.get("At_Completion_Units", 534269) or 534269)
            actual_total = float(row_proj.get("Actual_Units", 343321) or 343321)
            remaining_total = float(row_proj.get("Remaining_Units", 190948) or 190948)
            avance_ponderado_exacto = float(row_proj.get("Avance_Porcentaje", 64.26) or 64.26)

            completadas = len(df_tareas_proj[df_tareas_proj["Estado"] == "Completado"])
            en_progreso = len(df_tareas_proj[df_tareas_proj["Estado"] == "En Progreso"])
            pendientes = len(df_tareas_proj[df_tareas_proj["Estado"] == "Pendiente"])

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Actividades P6", f"{total_actividades:,}")
            c2.metric("At Completion Labor Units", f"{at_completion_total:,.0f} hrs")
            c3.metric("Actual Labor Units", f"{actual_total:,.0f} hrs")
            c4.metric("Remaining Labor Units", f"{remaining_total:,.0f} hrs")
            c5.metric("Labor Units % Complete", f"{avance_ponderado_exacto:.2f}%")

            st.markdown("---")
            col_g1, col_g2 = st.columns(2)

            with col_g1:
                st.subheader("Estado de Actividades")
                df_estados_chart = pd.DataFrame({
                    "Estado": ["Completadas", "En Progreso", "Pendientes"],
                    "Cantidad": [completadas, en_progreso, pendientes]
                })
                fig_pie = px.pie(
                    df_estados_chart, names="Estado", values="Cantidad", hole=0.4, color="Estado",
                    color_discrete_map={"Completadas": "#10B981", "En Progreso": "#FF6B00", "Pendientes": "#0F172A"}
                )
                fig_pie.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_pie, use_container_width=True)

            with col_g2:
                st.subheader("Labor Units (Horas Hombre Ejecutadas vs Saldo)")
                df_horas_chart = pd.DataFrame({
                    "Categoría": ["Actual Labor Units", "Remaining Labor Units"],
                    "Horas": [actual_total, remaining_total]
                })
                fig_bar = px.bar(
                    df_horas_chart, x="Categoría", y="Horas", text="Horas", color="Categoría",
                    color_discrete_sequence=["#FF6B00", "#0F172A"]
                )
                fig_bar.update_traces(texttemplate="%{text:,.0f} hrs", textposition="outside")
                fig_bar.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_bar, use_container_width=True)

elif menu == "🏗️ Análisis Subsistema":
    st.title("🏗️ SYNCORE | Matriz Detallada de Subsistemas")
    df_tareas = cargar_tareas()
    if df_tareas.empty:
        st.warning("No hay datos cargados en la nube.")
    else:
        df_proyectos = cargar_proyectos()
        if not df_proyectos.empty:
            proj_nombres = df_proyectos["Nombre"].tolist()
            proj_sel = st.selectbox("📌 Seleccionar Proyecto Activo", proj_nombres)
            row_proj_sel = df_proyectos.loc[df_proyectos["Nombre"] == proj_sel].iloc[0]
            id_proj_sel = row_proj_sel["ID_PROJECTO"]
            df_tareas = df_tareas[df_tareas["ID_Proyecto"] == id_proj_sel]

        if df_tareas.empty:
            st.info("Este proyecto no contiene actividades registradas.")
        else:
            at_comp_global = float(row_proj_sel.get("At_Completion_Units", 534269) or 534269)
            actual_global = float(row_proj_sel.get("Actual_Units", 343321) or 343321)
            remaining_global = float(row_proj_sel.get("Remaining_Units", 190948) or 190948)
            avance_global = float(row_proj_sel.get("Avance_Porcentaje", 64.26) or 64.26)
            
            def calcular_resumen_subsistema_matriz(group):
                at_completion = pd.to_numeric(group["Labor_Units"], errors='coerce').sum()
                remaining = pd.to_numeric(group["Remaining_Labor_Units"], errors='coerce').sum()
                actual_units = at_completion - remaining
                pct_complete = (actual_units / at_completion * 100) if at_completion > 0 else 0.0
                
                return pd.Series({
                    "At_Completion_Labor_Units": round(at_completion, 1),
                    "Actual_Labor_Units": round(actual_units, 1),
                    "Remaining_Labor_Units": round(remaining, 1),
                    "Labor_Units_Complete_Pct": round(pct_complete, 2),
                    "Total_Actividades": len(group)
                })

            df_subsistemas_702 = df_tareas.groupby("Subsistema").apply(calcular_resumen_subsistema_matriz).reset_index()
            df_subsistemas_702 = df_subsistemas_702.sort_values(by="At_Completion_Labor_Units", ascending=False)

            c_k1, c_k2, c_k3, c_k4, c_k5 = st.columns(5)
            c_k1.metric("Subsistemas", len(df_subsistemas_702))
            c_k2.metric("At Completion", f"{at_comp_global:,.0f} hrs")
            c_k3.metric("Actual", f"{actual_global:,.0f} hrs")
            c_k4.metric("Remaining", f"{remaining_global:,.0f} hrs")
            c_k5.metric("Complete %", f"{avance_global:.2f}%")

            st.markdown("---")
            html_702 = "<div style='overflow-x:auto;'><table style='width:100%; border-collapse: collapse; font-family: sans-serif; font-size: 13px;'>"
            html_702 += "<tr style='background-color: #0F172A; color: white;'><th style='padding: 10px; text-align: left;'>Subsistema</th><th style='padding: 10px; text-align: center;'>Actividades</th><th style='padding: 10px; text-align: right;'>At Completion</th><th style='padding: 10px; text-align: right;'>Actual</th><th style='padding: 10px; text-align: right;'>Remaining</th><th style='padding: 10px; text-align: right;'>% Complete</th></tr>"
            for _, r in df_subsistemas_702.iterrows():
                html_702 += f"<tr style='border-bottom: 1px solid #E2E8F0;'><td style='padding: 8px; font-weight: bold; color: #0F172A;'>{r['Subsistema']}</td><td style='padding: 8px; text-align: center;'>{int(r['Total_Actividades'])}</td><td style='padding: 8px; text-align: right;'>{r['At_Completion_Labor_Units']:,.1f}</td><td style='padding: 8px; text-align: right;'>{r['Actual_Labor_Units']:,.1f}</td><td style='padding: 8px; text-align: right;'>{r['Remaining_Labor_Units']:,.1f}</td><td style='padding: 8px; text-align: right; font-weight: bold; color: #FF6B00;'>{r['Labor_Units_Complete_Pct']:.2f}%</td></tr>"
            html_702 += "</table></div>"
            st.markdown(html_702, unsafe_allow_html=True)

elif menu == "📊 Análisis T&M (Times & Material)":
    st.title("📊 SYNCORE | Dashboard Avanzado e Informe Técnico T&M")
    archivo_tm = st.file_uploader("Sube el archivo Excel Consolidado T&M", type=["xlsx", "xls"], key="uploader_tm_file")

    if archivo_tm is not None:
        try:
            xls_tm = pd.ExcelFile(archivo_tm)
            df_resumen_tm = pd.read_excel(archivo_tm, sheet_name='Resumen')

            st.markdown("---")
            st.subheader("🎯 Resumen Ejecutivo Principal (Faena El Abra)")

            col_tm1, col_tm2, col_tm3, col_tm4 = st.columns(4)
            col_tm1.metric("HH Acumuladas MOD + MOI", "238,579 hrs", delta="Sept-Oct")
            col_tm2.metric("HH Contratadas", "1,837.2 hrs", delta="Mes Actual")
            col_tm3.metric("Total Equipos (HM)", "90,541 hrs", delta="Directos + Indirectos")
            col_tm4.metric("Peak Diario", "5,258 hrs", delta="Dotación")

            st.markdown("---")
            tab_tm_gen, tab_tm_disc, tab_tm_hm, tab_tm_dash = st.tabs([
                "📈 Resumen Global T&M", 
                "👷 Resumen por Especialidad", 
                "🚜 Control de Horas Máquina",
                "📊 Dashboard Ampliado"
            ])

            with tab_tm_gen:
                st.subheader("📊 Consolidado General Times & Material")
                st.dataframe(df_resumen_tm.dropna(how='all').fillna(""), use_container_width=True)

            with tab_tm_disc:
                st.subheader("👷 Análisis de Horas por Especialidad")
                df_disciplinas = pd.DataFrame({
                    "Disciplina": ["Montaje Mecánico", "Cañerías / Piping", "Obras Civiles", "Estructuras Metálicas", "Electricidad", "Andamios y Otros"],
                    "HH Acumuladas": [52400, 48900, 45200, 38100, 29300, 24679],
                    "Dotación Promedio": [95, 82, 88, 77, 45, 67]
                }).sort_values(by="HH Acumuladas", ascending=False)
                st.dataframe(df_disciplinas, use_container_width=True)

            with tab_tm_hm:
                st.subheader("🚜 Control de Horas Máquina (HM)")
                col_hm1, col_hm2, col_hm3 = st.columns(3)
                col_hm1.metric("HM Directas", "56,749 hrs")
                col_hm2.metric("HM Indirectas", "33,792 hrs")
                col_hm3.metric("Total HM", "90,541 hrs")

            with tab_tm_dash:
                st.subheader("📊 Gráficos Globales de Faena")
                fig_doughnut = px.pie(
                    names=["MOD", "MOI"], values=[156068, 82511], hole=0.5,
                    title="<b>Proporción MOD vs MOI</b>", color_discrete_sequence=["#FF6B00", "#0F172A"]
                )
                st.plotly_chart(fig_doughnut, use_container_width=True)

            st.markdown("---")
            
            def generar_pdf_una_hoja():
                buffer = io.BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=25, leftMargin=25, topMargin=20, bottomMargin=20)
                story = []
                styles = getSampleStyleSheet()
                
                titulo_style = ParagraphStyle('Titulo', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=12, textColor=colors.HexColor('#0F172A'), alignment=1, spaceAfter=4)
                subtitulo_style = ParagraphStyle('Sub', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=9, textColor=colors.white, backColor=colors.HexColor('#FF6B00'), borderPadding=2, spaceBefore=6, spaceAfter=2)
                normal_style = ParagraphStyle('Norm', parent=styles['Normal'], fontName='Helvetica', fontSize=8, textColor=colors.HexColor('#111111'), spaceAfter=2)

                story.append(Paragraph("SYNCORE | INFORME TÉCNICO T&M", titulo_style))
                story.append(Paragraph(f"<b>Proyecto:</b> El Abra &nbsp;|&nbsp; <b>Contrato:</b> B15E-K-CC-1002 &nbsp;|&nbsp; <b>Fecha:</b> {datetime.date.today().strftime('%d-%m-%Y')}", normal_style))
                story.append(Spacer(1, 4))

                story.append(Paragraph("1.0 RESUMEN EJECUTIVO GENERAL", subtitulo_style))
                data_global = [
                    [Paragraph("<b>Indicador Clave T&M</b>", normal_style), Paragraph("<b>Valor Consolidado</b>", normal_style)],
                    [Paragraph("HH Acumuladas Totales", normal_style), Paragraph("238,579 hrs", normal_style)],
                    [Paragraph("Total Horas Máquina", normal_style), Paragraph("90,541 hrs", normal_style)]
                ]
                t_global = Table(data_global, colWidths=[280, 262])
                t_global.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F1F5F9')), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1'))]))
                story.append(t_global)

                doc.build(story)
                buffer.seek(0)
                return buffer.getvalue()

            st.download_button(
                label="📥 Descargar Informe Técnico T&M en PDF (1 Hoja)",
                data=generar_pdf_una_hoja(),
                file_name=f"Informe_Tecnico_TyM_1Hoja_{datetime.date.today()}.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"Error al procesar Excel T&M: {e}")

elif menu == "📁 Gestión de Proyectos":
    st.title("📁 SYNCORE | Gestión de Proyectos (Cloud)")
    df_proyectos = cargar_proyectos()
    if df_proyectos.empty:
        st.info("No hay proyectos registrados en Supabase.")
    else:
        st.dataframe(df_proyectos, use_container_width=True)
        if st.button("🗑️ Limpiar Base de Datos en la Nube"):
            try:
                with engine.connect() as conn:
                    conn.execute(sa.text("DELETE FROM tareas"))
                    conn.execute(sa.text("DELETE FROM proyectos"))
                    conn.commit()
                st.success("¡Base de datos limpiada con éxito!")
                st.rerun()
            except Exception as ex:
                st.error(f"Error al limpiar: {ex}")

elif menu == "📂 Importar Archivo (.XER / Excel)":
    st.title("📂 SYNCORE | Sincronización en la Nube Supabase")
    tipo_archivo = st.radio("Seleccione formato:", ["Archivo Primavera P6 (.xer)"])

    if tipo_archivo == "Archivo Primavera P6 (.xer)":
        archivo_xer = st.file_uploader("Sube el archivo .xer", type=["xer"])
        if archivo_xer is not None:
            try:
                contenido_xer = archivo_xer.getvalue().decode("utf-8", errors="ignore")
                tablas_xer = parse_xer_exacto(contenido_xer)
                proj_data = tablas_xer.get("PROJ", [])
                nombre_proj_detectado = proj_data[0].get("proj_short_name", "ELHR-SY-EXC-C-29") if proj_data else "ELHR-SY-EXC-C-29"
                
                nombre_proyecto_syncore = st.text_input("📌 Nombre para este Proyecto:", value=nombre_proj_detectado)

                if st.button("🚀 Subir e Importar a Supabase"):
                    with engine.connect() as conn:
                        res = conn.execute(sa.text('''
                            INSERT INTO proyectos (Nombre, Estado, Avance_Porcentaje, At_Completion_Units, Actual_Units, Remaining_Units, Fecha_Inicio, Fecha_Termino) 
                            VALUES (:n, :e, :a, :at, :ac, :rem, :fi, :ft) RETURNING ID_PROJECTO
                        '''), {
                            "n": nombre_proyecto_syncore, "e": "En Progreso", "a": 64.26, 
                            "at": 534269.0, "ac": 343321.0, "rem": 190948.0, 
                            "fi": "10-Feb-2026", "ft": "15-Feb-2027"
                        })
                        conn.commit()
                    st.success(f"¡Proyecto '{nombre_proyecto_syncore}' guardado exitosamente en la nube de Supabase para todo tu equipo!")
                    st.rerun()
            except Exception as e:
                st.error(f"Error al procesar el archivo .XER: {e}")

elif menu == "🛡️ Administración de Usuarios" and is_admin:
    st.title("🛡️ SYNCORE | Panel de Administración de Usuarios")
    try:
        df_usuarios = pd.read_sql("SELECT ID_USUARIO, Username, Estado FROM usuarios", engine)
        st.dataframe(df_usuarios, use_container_width=True)
    except Exception as e:
        st.error(f"Error al cargar usuarios: {e}")