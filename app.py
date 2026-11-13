"""SYNCORE | Control Ejecutivo P6, 6WLA e Informes — versión en un solo archivo.

Ejecutar:  streamlit run syncore_app.py
Requisitos: pip install streamlit==1.45.1 "pandas<3" plotly openpyxl reportlab pillow
            (para PostgreSQL en producción, además: psycopg2-binary y la variable DATABASE_URL)

Cada sección "MÓDULO" agrupa sus funciones en un espacio de nombres (db.*, auth.*, xer.*...).
"""
from __future__ import annotations

import os
from types import SimpleNamespace

import streamlit as st

st.set_page_config(page_title="SYNCORE | Control Ejecutivo P6", page_icon="🟧", layout="wide")


# ======================================================================
# MÓDULO: xer
# ======================================================================
def _modulo_xer():
    from dataclasses import dataclass, field
    
    ESTADOS_P6 = {
        "TK_Complete": "Completado",
        "TK_Active": "En Progreso",
        "TK_NotStart": "No Iniciado",
    }
    
    TIPOS_P6 = {
        "TT_Task": "Tarea",
        "TT_Rsrc": "Tarea de recurso",
        "TT_Mile": "Hito de inicio",
        "TT_FinMile": "Hito de término",
        "TT_LOE": "Level of Effort",
        "TT_WBS": "Resumen WBS",
    }
    
    SIN_CODIGO = "(Sin código asignado)"
    
    
    class XerError(ValueError):
        """El archivo no es un .XER válido o le faltan tablas obligatorias."""
    
    
    def decodificar(datos: bytes) -> str:
        """P6 exporta en UTF-8 o en la página de códigos de Windows (cp1252)."""
        for enc in ("utf-8-sig", "cp1252"):
            try:
                return datos.decode(enc)
            except UnicodeDecodeError:
                continue
        return datos.decode("latin-1")
    
    
    _R = "\ufffd"  # carácter de reemplazo: P6 exportó sin la codificación correcta
    _REPARACIONES = [  # (patrón sin distinguir mayúsculas, reemplazo en minúsculas)
        (rf"i{_R}n", "ión"),
        (rf"ca{_R}er{_R}a", "cañería"), (rf"ca{_R}er", "cañer"), (rf"ca{_R}a", "caña"), (rf"ba{_R}o", "baño"),
        (rf"(?<!\w){_R}rea", "área"), (rf"el{_R}ctric", "eléctric"), (rf"mec{_R}nic", "mecánic"), (rf"v{_R}lvul", "válvul"),
        (rf"hormig{_R}n", "hormigón"), (rf"ingenier{_R}a", "ingeniería"), (rf"tuber{_R}a", "tubería"),
        (rf"topograf{_R}a", "topografía"), (rf"topogr{_R}fic", "topográfic"), (rf"t{_R}rmic", "térmic"),
        (rf"fr{_R}a", "fría"), (rf"l{_R}nea", "línea"), (rf"(?<!\w){_R}ptic", "óptic"), (rf"(\d){_R}", r"\1°"),
        (rf"\bn{_R}", "n°"), (rf"m{_R}dulo", "módulo"), (rf"cami{_R}n", "camión"), (rf"gr{_R}a", "grúa"),
        (rf"c{_R}mara", "cámara"), (rf"vac{_R}o", "vacío"), (rf"env{_R}o", "envío"), (rf"st{_R}tic", "státic"),
        (rf"\bd{_R}a", "día"), (rf"nitr{_R}geno", "nitrógeno"), (rf"desag{_R}e", "desagüe"), (rf"c{_R}bic", "cúbic"),
    ]
    
    
    def reparar_texto(texto: str) -> str:
        """Recupera tildes y eñes perdidas en exportaciones de P6 con codificación incorrecta."""
        import re
    
        if _R not in texto:
            return texto
    
        def con_caso(rep):
            def f(m):
                s = m.expand(rep)
                orig = m.group(0)
                letras = [ch for ch in orig if ch.isalpha()]
                if letras and all(ch.isupper() for ch in letras) and len(letras) > 1:
                    return s.upper()
                if orig[:1].isupper():
                    return s[:1].upper() + s[1:]
                if orig[:1] == _R:  # palabra que empieza con la letra perdida: mayúscula al inicio de un campo
                    previo = m.string[max(0, m.start() - 2):m.start()]
                    if not previo or previo[-1] in "\t\n(" or previo in (". ", "- "):
                        return s[:1].upper() + s[1:]
                return s
            return f
    
        for patron, rep in _REPARACIONES:
            texto = re.sub(patron, con_caso(rep), texto, flags=re.IGNORECASE)
        return texto
    
    
    def leer_tablas(texto: str) -> dict[str, list[dict[str, str]]]:
        """Convierte el texto .XER en {tabla: [filas como dict]}."""
        texto = reparar_texto(texto)
        tablas: dict[str, list[dict[str, str]]] = {}
        actual = None
        columnas: list[str] = []
        for linea in texto.splitlines():
            if not linea:
                continue
            partes = linea.split("\t")
            marca = partes[0]
            if marca == "%T" and len(partes) > 1:
                actual = partes[1].strip()
                tablas[actual] = []
                columnas = []
            elif marca == "%F" and actual:
                columnas = [c.strip() for c in partes[1:]]
            elif marca == "%R" and actual and columnas:
                valores = partes[1:]
                if len(valores) < len(columnas):
                    valores += [""] * (len(columnas) - len(valores))
                tablas[actual].append(dict(zip(columnas, valores)))
        if "TASK" not in tablas:
            raise XerError("El archivo no contiene la tabla TASK. ¿Es un .XER exportado desde P6?")
        return tablas
    
    
    def _num(v) -> float:
        if v in (None, ""):
            return 0.0
        try:
            return float(str(v).replace(",", "."))
        except ValueError:
            return 0.0
    
    
    def _fecha(v: str | None) -> str | None:
        """'2026-10-23 16:00' -> '2026-10-23'."""
        v = (v or "").strip()
        if not v or v.startswith("1899"):
            return None
        return v[:10]
    
    
    def _etiqueta_codigo(code: dict) -> str:
        corto = (code.get("short_name") or "").strip()
        nombre = (code.get("actv_code_name") or "").strip()
        if not nombre:
            return corto or SIN_CODIGO
        if not corto or nombre.upper().startswith(corto.upper()):
            return nombre
        return f"{corto}: {nombre}"
    
    
    @dataclass
    class ProyectoXer:
        proj_id: str
        codigo: str
        nombre: str
        fecha_corte: str | None
        fecha_inicio: str | None
        fecha_termino: str | None
        actividades: list[dict] = field(default_factory=list)
        # [(task_code, tipo_codigo, valor)]
        codigos: list[tuple[str, str, str]] = field(default_factory=list)
        tipos_codigo: list[str] = field(default_factory=list)
        fuente_horas: str = "TASK"
    
        @property
        def actual(self) -> float:
            return sum(a["actual"] for a in self.actividades)
    
        @property
        def remaining(self) -> float:
            return sum(a["remaining"] for a in self.actividades)
    
        @property
        def at_completion(self) -> float:
            return self.actual + self.remaining
    
        @property
        def budget(self) -> float:
            return sum(a["budget"] for a in self.actividades)
    
        @property
        def avance_pct(self) -> float:
            atc = self.at_completion
            return self.actual / atc * 100 if atc > 0 else 0.0
    
    
    def listar_proyectos(tablas: dict) -> list[dict]:
        salida = []
        for p in tablas.get("PROJECT", []) or tablas.get("PROJ", []):
            salida.append({"proj_id": p.get("proj_id", ""), "codigo": p.get("proj_short_name", "").strip()})
        if not salida:
            ids = sorted({t.get("proj_id", "") for t in tablas["TASK"]})
            salida = [{"proj_id": i, "codigo": f"Proyecto {i}"} for i in ids]
        return salida
    
    
    def extraer_proyecto(tablas: dict, proj_id: str | None = None) -> ProyectoXer:
        proyectos = listar_proyectos(tablas)
        if not proyectos:
            raise XerError("El archivo no contiene proyectos.")
        if proj_id is None:
            proj_id = proyectos[0]["proj_id"]
    
        fila_proj = next(
            (p for p in (tablas.get("PROJECT", []) or tablas.get("PROJ", [])) if p.get("proj_id") == proj_id), {}
        )
        codigo = (fila_proj.get("proj_short_name") or f"Proyecto {proj_id}").strip()
    
        # El nombre largo del proyecto es el nodo raíz de la WBS.
        nombre = codigo
        wbs_nombres: dict[str, str] = {}
        for w in tablas.get("PROJWBS", []):
            if w.get("proj_id") != proj_id:
                continue
            wbs_nombres[w.get("wbs_id", "")] = (w.get("wbs_name") or "").strip()
            if w.get("proj_node_flag") == "Y" and w.get("wbs_name"):
                nombre = w["wbs_name"].strip()
    
        tareas = [t for t in tablas["TASK"] if t.get("proj_id") == proj_id]
        if not tareas:
            raise XerError(f"El proyecto {codigo} no tiene actividades.")
    
        # Horas: P6 resume las Labor Units en TASK (act_work_qty / remain_work_qty).
        # Si vinieran vacías, se suman las asignaciones de recursos de tipo Labor.
        usar_rsrc = all(_num(t.get("act_work_qty")) == 0 and _num(t.get("remain_work_qty")) == 0 for t in tareas)
        horas_rsrc: dict[str, list[float]] = {}
        if usar_rsrc and tablas.get("TASKRSRC"):
            labor_ids = {r.get("rsrc_id") for r in tablas.get("RSRC", []) if r.get("rsrc_type") == "RT_Labor"}
            for r in tablas["TASKRSRC"]:
                es_labor = r.get("rsrc_type") == "RT_Labor" or r.get("rsrc_id") in labor_ids
                if r.get("proj_id") != proj_id or not es_labor:
                    continue
                h = horas_rsrc.setdefault(r.get("task_id", ""), [0.0, 0.0, 0.0])
                h[0] += _num(r.get("act_reg_qty")) + _num(r.get("act_ot_qty"))
                h[1] += _num(r.get("remain_qty"))
                h[2] += _num(r.get("target_qty"))
    
        actividades = []
        id_a_codigo = {}
        for t in tareas:
            task_code = (t.get("task_code") or t.get("task_id", "")).strip()
            id_a_codigo[t.get("task_id", "")] = task_code
            if usar_rsrc:
                actual, remaining, budget = horas_rsrc.get(t.get("task_id", ""), [0.0, 0.0, 0.0])
            else:
                actual = _num(t.get("act_work_qty"))
                remaining = _num(t.get("remain_work_qty"))
                budget = _num(t.get("target_work_qty"))
            actividades.append(
                {
                    "task_code": task_code,
                    "nombre": (t.get("task_name") or "").strip(),
                    "tipo": TIPOS_P6.get(t.get("task_type", ""), t.get("task_type", "")),
                    "estado": ESTADOS_P6.get(t.get("status_code", ""), t.get("status_code", "")),
                    "wbs": wbs_nombres.get(t.get("wbs_id", ""), ""),
                    "budget": budget,
                    "actual": actual,
                    "remaining": remaining,
                    "inicio": _fecha(t.get("act_start_date")) or _fecha(t.get("early_start_date")) or _fecha(t.get("target_start_date")),
                    "fin": _fecha(t.get("act_end_date")) or _fecha(t.get("early_end_date")) or _fecha(t.get("target_end_date")),
                    "holgura_total_hr": _num(t.get("total_float_hr_cnt")) if t.get("total_float_hr_cnt") not in (None, "") else None,
                }
            )
    
        # Códigos de actividad (Subsistemas, Sistema, Área, Fase...)
        tipos = {
            ty.get("actv_code_type_id"): (ty.get("actv_code_type") or "").strip()
            for ty in tablas.get("ACTVTYPE", [])
            if ty.get("actv_code_type_scope") != "AS_Project" or ty.get("proj_id") == proj_id
        }
        codes = {c.get("actv_code_id"): c for c in tablas.get("ACTVCODE", [])}
        codigos = []
        tipos_usados = set()
        for a in tablas.get("TASKACTV", []):
            if a.get("proj_id") not in (None, "", proj_id):
                continue
            tc = id_a_codigo.get(a.get("task_id", ""))
            tipo = tipos.get(a.get("actv_code_type_id"))
            code = codes.get(a.get("actv_code_id"))
            if tc is None or not tipo or not code:
                continue
            codigos.append((tc, tipo, _etiqueta_codigo(code)))
            tipos_usados.add(tipo)
    
        return ProyectoXer(
            proj_id=proj_id,
            codigo=codigo,
            nombre=nombre,
            fecha_corte=_fecha(fila_proj.get("last_recalc_date")) or _fecha(fila_proj.get("next_data_date")),
            fecha_inicio=_fecha(fila_proj.get("plan_start_date")) or min((a["inicio"] for a in actividades if a["inicio"]), default=None),
            fecha_termino=_fecha(fila_proj.get("scd_end_date")) or _fecha(fila_proj.get("plan_end_date"))
            or max((a["fin"] for a in actividades if a["fin"]), default=None),
            actividades=actividades,
            codigos=codigos,
            tipos_codigo=sorted(tipos_usados),
            fuente_horas="TASKRSRC (Labor)" if usar_rsrc else "TASK",
        )
    
    
    def tipo_codigo_por_defecto(tipos: list[str]) -> str | None:
        """Prefiere el código de subsistemas si existe."""
        for criterio in (lambda t: t.lower().endswith("subsistemas") and "wbs" not in t.lower(),
                         lambda t: "subsistema" in t.lower()):
            for t in tipos:
                if criterio(t):
                    return t
        return tipos[0] if tipos else None

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


xer = _modulo_xer()

# ======================================================================
# MÓDULO: db
# ======================================================================
def _modulo_db():
    import os
    import sqlite3
    from contextlib import contextmanager
    from datetime import datetime
    
    import pandas as pd
    
    DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
    ES_POSTGRES = DATABASE_URL.startswith(("postgres://", "postgresql://"))
    def _carpeta_app() -> str:
        """Carpeta de la app (la del archivo principal), para que la base no dependa desde dónde se ejecuta Streamlit."""
        try:
            aqui = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            return os.getcwd()
        return os.path.dirname(aqui) if os.path.basename(aqui) == "core" else aqui
    
    
    SQLITE_PATH = os.environ.get("SQLITE_PATH") or os.path.join(_carpeta_app(), "syncore.db")
    
    if ES_POSTGRES:
        import psycopg2
        import psycopg2.extras
    
    _PK = "SERIAL PRIMARY KEY" if ES_POSTGRES else "INTEGER PRIMARY KEY AUTOINCREMENT"
    _REAL = "DOUBLE PRECISION" if ES_POSTGRES else "REAL"
    
    ESQUEMA = [
        f"""CREATE TABLE IF NOT EXISTS contratos (
            id {_PK},
            codigo TEXT UNIQUE NOT NULL,
            nombre TEXT NOT NULL,
            activo INTEGER NOT NULL DEFAULT 1,
            creado_en TEXT NOT NULL
        )""",
        """CREATE TABLE IF NOT EXISTS usuario_contratos (
            usuario_id INTEGER NOT NULL,
            contrato_id INTEGER NOT NULL,
            PRIMARY KEY (usuario_id, contrato_id)
        )""",
        f"""CREATE TABLE IF NOT EXISTS usuarios (
            id {_PK},
            username TEXT UNIQUE NOT NULL,
            nombre TEXT,
            password_hash TEXT NOT NULL,
            rol TEXT NOT NULL DEFAULT 'lector',
            estado TEXT NOT NULL DEFAULT 'Pendiente',
            intentos_fallidos INTEGER NOT NULL DEFAULT 0,
            bloqueado_hasta TEXT,
            creado_en TEXT NOT NULL
        )""",
        f"""CREATE TABLE IF NOT EXISTS proyectos (
            id {_PK},
            codigo TEXT UNIQUE NOT NULL,
            nombre TEXT NOT NULL,
            origen TEXT NOT NULL DEFAULT 'P6',
            creado_en TEXT NOT NULL
        )""",
        f"""CREATE TABLE IF NOT EXISTS cortes (
            id {_PK},
            proyecto_id INTEGER NOT NULL REFERENCES proyectos(id) ON DELETE CASCADE,
            fecha_corte TEXT NOT NULL,
            fecha_inicio TEXT,
            fecha_termino TEXT,
            archivo TEXT,
            cargado_por TEXT,
            cargado_en TEXT NOT NULL,
            budget {_REAL} NOT NULL DEFAULT 0,
            actual {_REAL} NOT NULL DEFAULT 0,
            remaining {_REAL} NOT NULL DEFAULT 0,
            n_actividades INTEGER NOT NULL DEFAULT 0,
            UNIQUE (proyecto_id, fecha_corte)
        )""",
        f"""CREATE TABLE IF NOT EXISTS actividades (
            id {_PK},
            corte_id INTEGER NOT NULL REFERENCES cortes(id) ON DELETE CASCADE,
            task_code TEXT NOT NULL,
            nombre TEXT,
            tipo TEXT,
            estado TEXT,
            wbs TEXT,
            budget {_REAL} NOT NULL DEFAULT 0,
            actual {_REAL} NOT NULL DEFAULT 0,
            remaining {_REAL} NOT NULL DEFAULT 0,
            inicio TEXT,
            fin TEXT,
            holgura_total_hr {_REAL}
        )""",
        "CREATE INDEX IF NOT EXISTS ix_act_corte ON actividades(corte_id)",
        """CREATE TABLE IF NOT EXISTS actividad_codigos (
            corte_id INTEGER NOT NULL REFERENCES cortes(id) ON DELETE CASCADE,
            task_code TEXT NOT NULL,
            tipo_codigo TEXT NOT NULL,
            valor TEXT NOT NULL
        )""",
        "CREATE INDEX IF NOT EXISTS ix_cod_corte ON actividad_codigos(corte_id, tipo_codigo)",
        f"""CREATE TABLE IF NOT EXISTS informes (
            id {_PK},
            nombre TEXT NOT NULL,
            proyecto_id INTEGER REFERENCES proyectos(id) ON DELETE SET NULL,
            fecha_corte TEXT NOT NULL,
            archivo TEXT,
            hoja TEXT,
            cargado_por TEXT,
            cargado_en TEXT NOT NULL
        )""",
        f"""CREATE TABLE IF NOT EXISTS informe_items (
            id {_PK},
            informe_id INTEGER NOT NULL REFERENCES informes(id) ON DELETE CASCADE,
            item TEXT NOT NULL,
            at_completion {_REAL} NOT NULL DEFAULT 0,
            remaining {_REAL} NOT NULL DEFAULT 0,
            avance_pct {_REAL} NOT NULL DEFAULT 0
        )""",
        "CREATE INDEX IF NOT EXISTS ix_items_inf ON informe_items(informe_id)",
        f"""CREATE TABLE IF NOT EXISTS restricciones (
            id {_PK},
            proyecto TEXT NOT NULL,
            actividad TEXT NOT NULL,
            subsistema TEXT,
            semana TEXT,
            restriccion TEXT NOT NULL,
            responsable TEXT,
            estado TEXT NOT NULL,
            fecha_compromiso TEXT,
            actualizado_por TEXT,
            actualizado_en TEXT NOT NULL
        )""",
        # ---- Informe diario
        f"""CREATE TABLE IF NOT EXISTS diario_archivos (
            id {_PK},
            tipo TEXT NOT NULL,
            nombre TEXT NOT NULL,
            fecha TEXT NOT NULL,
            numero INTEGER,
            contenido TEXT NOT NULL,
            creado_por TEXT,
            creado_en TEXT NOT NULL
        )""",
        f"""CREATE TABLE IF NOT EXISTS semanal_archivos (
            id {_PK},
            tipo TEXT NOT NULL,
            nombre TEXT NOT NULL,
            fecha TEXT NOT NULL,
            numero INTEGER,
            contenido TEXT NOT NULL,
            creado_por TEXT,
            creado_en TEXT NOT NULL
        )""",
        f"""CREATE TABLE IF NOT EXISTS cubicacion_planos (
            id {_PK},
            plano TEXT NOT NULL,
            rev TEXT,
            disciplina TEXT NOT NULL,
            descripcion TEXT,
            datos TEXT NOT NULL,
            hh REAL,
            creado_por TEXT,
            creado_en TEXT NOT NULL
        )""",
        f"""CREATE TABLE IF NOT EXISTS cubicacion_cargas (
            id {_PK},
            fecha TEXT NOT NULL,
            archivo TEXT NOT NULL,
            disciplinas TEXT,
            datos TEXT NOT NULL,
            creado_por TEXT,
            creado_en TEXT NOT NULL
        )""",
        f"""CREATE TABLE IF NOT EXISTS diario_reportes (
            id {_PK},
            fecha TEXT NOT NULL,
            archivo TEXT NOT NULL,
            datos TEXT NOT NULL,
            cargado_por TEXT,
            cargado_en TEXT NOT NULL
        )""",
        "CREATE INDEX IF NOT EXISTS ix_drep_fecha ON diario_reportes(fecha)",
        f"""CREATE TABLE IF NOT EXISTS diario_informes (
            id {_PK},
            fecha TEXT NOT NULL,
            numero INTEGER NOT NULL,
            resumen TEXT NOT NULL,
            archivo_id INTEGER,
            emitido_por TEXT,
            emitido_en TEXT NOT NULL
        )""",
        """CREATE TABLE IF NOT EXISTS diario_config (
            clave TEXT PRIMARY KEY,
            valor TEXT NOT NULL
        )""",
        # ---- Sesiones persistentes (sobreviven a refrescar la página)
        """CREATE TABLE IF NOT EXISTS sesiones (
            token_hash TEXT PRIMARY KEY,
            usuario_id INTEGER NOT NULL,
            navegador TEXT,
            creada_en TEXT NOT NULL,
            expira_en TEXT NOT NULL
        )""",
        # ---- Times & Material
        f"""CREATE TABLE IF NOT EXISTS tm_cargas (
            id {_PK},
            archivo TEXT NOT NULL,
            fecha_corte TEXT,
            contrato TEXT,
            datos TEXT NOT NULL,
            cargado_por TEXT,
            cargado_en TEXT NOT NULL
        )""",
        # ---- LOG de desviación de recursos
        f"""CREATE TABLE IF NOT EXISTS desv_cargas (
            id {_PK},
            archivo TEXT NOT NULL,
            fecha_corte TEXT,
            revision TEXT,
            hh REAL,
            registros INTEGER,
            datos TEXT NOT NULL,
            cargado_por TEXT,
            cargado_en TEXT NOT NULL
        )""",
    ]
    
    
    def ahora() -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    
    def _sql(q: str) -> str:
        return q.replace("?", "%s") if ES_POSTGRES else q
    
    
    def _conectar():
        if ES_POSTGRES:
            url = DATABASE_URL.replace("postgres://", "postgresql://", 1)
            return psycopg2.connect(url)
        conn = sqlite3.connect(SQLITE_PATH, timeout=30)
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = WAL")
        return conn
    
    
    @contextmanager
    def transaccion():
        conn = _conectar()
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
    
    
    def ejecutar(q: str, params: tuple = (), conn=None, _reintento: bool = True):
        if conn is None:
            try:
                with transaccion() as c:
                    return ejecutar(q, params, c)
            except Exception as e:
                if _reintento and _falta_esquema(e):
                    init_db()
                    return ejecutar(q, params, None, _reintento=False)
                raise
        cur = conn.cursor()
        cur.execute(_sql(q), params)
        return cur
    
    
    def insertar(q: str, params: tuple, conn) -> int:
        """INSERT ... RETURNING id (SQLite >= 3.35 y PostgreSQL)."""
        cur = conn.cursor()
        cur.execute(_sql(q + " RETURNING id"), params)
        return cur.fetchone()[0]
    
    
    def insertar_muchos(q: str, filas: list[tuple], conn):
        if not filas:
            return
        cur = conn.cursor()
        if ES_POSTGRES:
            # execute_values envía miles de filas en pocos viajes a la base.
            base, _, valores = q.partition("VALUES")
            psycopg2.extras.execute_values(cur, base + "VALUES %s", filas, page_size=1000)
        else:
            cur.executemany(q, filas)
    
    
    NUMERICAS = {"budget", "actual", "remaining", "at_completion", "avance_pct", "holgura_total_hr", "n_actividades",
                 "cortes", "items", "hh_gastadas"}
    
    
    def _falta_esquema(e: Exception) -> bool:
        """La base no tiene una tabla o columna: se creó con otra versión, o el archivo se borró o se cambió."""
        m = str(e).lower()
        return "no such table" in m or "no such column" in m or "does not exist" in m
    
    
    def consultar(q: str, params: tuple = (), _reintento: bool = True) -> pd.DataFrame:
        conn = _conectar()
        try:
            cur = conn.cursor()
            try:
                cur.execute(_sql(q), params)
            except Exception as e:
                if _reintento and _falta_esquema(e):
                    conn.close()
                    init_db()  # base creada con una versión anterior: agrega las columnas nuevas
                    return consultar(q, params, _reintento=False)
                raise
            cols = [d[0] for d in cur.description]
            df = pd.DataFrame(cur.fetchall(), columns=cols)
            for col in NUMERICAS.intersection(cols):  # columnas vacías llegan como object
                df[col] = pd.to_numeric(df[col], errors="coerce")
            return df
        finally:
            conn.close()
    
    
    def uno(q: str, params: tuple = (), _reintento: bool = True):
        conn = _conectar()
        try:
            cur = conn.cursor()
            try:
                cur.execute(_sql(q), params)
            except Exception as e:
                if _reintento and _falta_esquema(e):
                    conn.close()
                    init_db()
                    return uno(q, params, _reintento=False)
                raise
            return cur.fetchone()
        finally:
            conn.close()
    
    
    def identidad_bd() -> str:
        """Cambia si cambia la estructura o si el archivo SQLite se borra o se reemplaza."""
        if ES_POSTGRES:
            return VERSION_ESQUEMA
        try:
            st_ = os.stat(SQLITE_PATH)
            return f"{VERSION_ESQUEMA}|{SQLITE_PATH}|{st_.st_ino}|{st_.st_ctime_ns}"
        except OSError:
            return f"{VERSION_ESQUEMA}|{SQLITE_PATH}|nuevo"
    
    
    TABLAS_POR_CONTRATO = ("proyectos", "informes", "restricciones", "diario_archivos", "diario_reportes", "diario_informes",
                           "tm_cargas", "desv_cargas", "semanal_archivos",
                           "cubicacion_cargas", "cubicacion_planos")
    
    MIGRACIONES = [  # columnas agregadas después de la primera versión
        ("informe_items", "hh_gastadas", _REAL),
    ] + [(t, "contrato_id", "INTEGER") for t in TABLAS_POR_CONTRATO] + [
        ("usuarios", "cargo", "TEXT"),
        ("usuarios", "firma", "TEXT"),                 # PNG en base64 (firma para el informe diario)
        ("usuarios", "firma_actualizada", "TEXT"),
    ]
    
    
    VERSION_ESQUEMA = f"{len(ESQUEMA)}-{len(MIGRACIONES)}-" + "-".join(f"{t}.{c}" for t, c, _ in MIGRACIONES)
    
    
    def _columnas(conn, tabla: str) -> set[str]:
        cur = conn.cursor()
        if ES_POSTGRES:
            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = %s", (tabla,))
            return {r[0] for r in cur.fetchall()}
        cur.execute(f"PRAGMA table_info({tabla})")
        return {r[1] for r in cur.fetchall()}
    
    
    def _respaldo_antes_de_contratos():
        """Antes de pasar una base antigua a contratos, deja una copia completa junto a ella (solo SQLite, una vez)."""
        if ES_POSTGRES or not os.path.exists(SQLITE_PATH):
            return
        try:
            conn = sqlite3.connect(SQLITE_PATH, timeout=30)
            try:
                tablas = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
                if "usuarios" not in tablas or not conn.execute("SELECT 1 FROM usuarios LIMIT 1").fetchone():
                    return
                if "diario_config" in tablas and conn.execute(
                        "SELECT 1 FROM diario_config WHERE clave = '_migracion_contratos'").fetchone():
                    return
            finally:
                conn.close()
            destino = SQLITE_PATH.replace(".db", "") + f"_antes_de_contratos_{datetime.now():%Y%m%d_%H%M%S}.db"
            with open(destino, "wb") as f:
                f.write(respaldo_sqlite())
        except Exception:
            pass                                     # el respaldo es una precaución: no impide abrir la app
    
    
    def init_db():
        _respaldo_antes_de_contratos()
        with transaccion() as conn:
            for q in ESQUEMA:
                conn.cursor().execute(q)
            for tabla, columna, tipo in MIGRACIONES:
                if columna not in _columnas(conn, tabla):
                    conn.cursor().execute(f"ALTER TABLE {tabla} ADD COLUMN {columna} {tipo}")
            _migrar_a_contratos(conn)
    
    
    CONTRATO_INICIAL = ("B1SE-K-CC-1002", "Construcción, montaje y pre-comisionamiento planta de calentamiento de refino")
    CLAVES_GLOBALES = ("resumen_informe_",)          # claves de configuración que no dependen del contrato
    
    
    def _migrar_a_contratos(conn):
        """Bases creadas antes de los contratos: todo lo existente pasa al contrato inicial (una sola vez, sin perder datos)."""
        import json
        fila = ejecutar("SELECT MIN(id) FROM contratos", (), conn).fetchone()
        cid = fila[0] if fila else None
        hay_datos = any(ejecutar(f"SELECT 1 FROM {t} LIMIT 1", (), conn).fetchone() for t in TABLAS_POR_CONTRATO)
        if cid is None:
            if not hay_datos and not ejecutar("SELECT 1 FROM usuarios LIMIT 1", (), conn).fetchone():
                # instalación nueva: el administrador crea su primer contrato; no hay nada que migrar
                if not ejecutar("SELECT 1 FROM diario_config WHERE clave = '_migracion_contratos'", (), conn).fetchone():
                    ejecutar("INSERT INTO diario_config (clave, valor) VALUES ('_migracion_contratos', 'null')", (), conn)
                return
            tm = ejecutar("SELECT contrato FROM tm_cargas WHERE contrato IS NOT NULL AND contrato <> '' LIMIT 1", (), conn).fetchone()
            codigo, nombre = (tm[0], CONTRATO_INICIAL[1]) if tm else CONTRATO_INICIAL
            cid = insertar("INSERT INTO contratos (codigo, nombre, activo, creado_en) VALUES (?, ?, 1, ?)", (codigo, nombre, ahora()), conn)
        for t in TABLAS_POR_CONTRATO:
            ejecutar(f"UPDATE {t} SET contrato_id = ? WHERE contrato_id IS NULL", (cid,), conn)
        hecho = ejecutar("SELECT valor FROM diario_config WHERE clave = '_migracion_contratos'", (), conn).fetchone()
        if not hecho:
            for (uid,) in ejecutar("SELECT id FROM usuarios", (), conn).fetchall():
                if not ejecutar("SELECT 1 FROM usuario_contratos WHERE usuario_id = ? AND contrato_id = ?", (uid, cid), conn).fetchone():
                    ejecutar("INSERT INTO usuario_contratos (usuario_id, contrato_id) VALUES (?, ?)", (uid, cid), conn)
            for clave, valor in ejecutar("SELECT clave, valor FROM diario_config", (), conn).fetchall():
                if clave.startswith("c") and ":" in clave or clave.startswith(CLAVES_GLOBALES) or clave.startswith("_"):
                    continue
                nueva = f"c{cid}:{clave}"
                if not ejecutar("SELECT 1 FROM diario_config WHERE clave = ?", (nueva,), conn).fetchone():
                    ejecutar("INSERT INTO diario_config (clave, valor) VALUES (?, ?)", (nueva, valor), conn)
                ejecutar("DELETE FROM diario_config WHERE clave = ?", (clave,), conn)
            ejecutar("INSERT INTO diario_config (clave, valor) VALUES ('_migracion_contratos', ?)", (json.dumps(cid),), conn)
    
    
    # ---------------------------------------------------------------- Contratos
    
    def contratos(solo_activos: bool = False) -> pd.DataFrame:
        q = "SELECT id, codigo, nombre, activo, creado_en FROM contratos"
        return consultar(q + (" WHERE activo = 1" if solo_activos else "") + " ORDER BY codigo")
    
    
    def contrato(cid: int) -> dict | None:
        fila = uno("SELECT id, codigo, nombre, activo FROM contratos WHERE id = ?", (int(cid),))
        return dict(zip(["id", "codigo", "nombre", "activo"], fila)) if fila else None
    
    
    def crear_contrato(codigo: str, nombre: str) -> int:
        codigo = codigo.strip()
        if not codigo:
            raise ValueError("El código del contrato no puede estar vacío.")
        if uno("SELECT id FROM contratos WHERE LOWER(codigo) = LOWER(?)", (codigo,)):
            raise ValueError(f"Ya existe el contrato {codigo}.")
        with transaccion() as conn:
            return insertar("INSERT INTO contratos (codigo, nombre, activo, creado_en) VALUES (?, ?, 1, ?)",
                            (codigo, nombre.strip() or codigo, ahora()), conn)
    
    
    def actualizar_contrato(cid: int, codigo: str, nombre: str, activo: bool):
        codigo = codigo.strip()
        otro = uno("SELECT id FROM contratos WHERE LOWER(codigo) = LOWER(?) AND id <> ?", (codigo, int(cid)))
        if not codigo or otro:
            raise ValueError("Código vacío o repetido.")
        ejecutar("UPDATE contratos SET codigo = ?, nombre = ?, activo = ? WHERE id = ?", (codigo, nombre.strip() or codigo, int(activo), int(cid)))
    
    
    def contratos_de_usuario(uid: int, es_admin: bool) -> pd.DataFrame:
        """Contratos activos que el usuario puede ver (el administrador ve todos)."""
        if es_admin:
            return contratos(solo_activos=True)
        return consultar("""SELECT c.id, c.codigo, c.nombre, c.activo, c.creado_en FROM contratos c
                            JOIN usuario_contratos u ON u.contrato_id = c.id
                            WHERE u.usuario_id = ? AND c.activo = 1 ORDER BY c.codigo""", (int(uid),))
    
    
    def usuarios_de_contrato(cid: int) -> list[int]:
        return consultar("SELECT usuario_id FROM usuario_contratos WHERE contrato_id = ?", (int(cid),))["usuario_id"].astype(int).tolist()
    
    
    def asignar_usuarios(cid: int, usuario_ids: list[int]):
        with transaccion() as conn:
            ejecutar("DELETE FROM usuario_contratos WHERE contrato_id = ?", (int(cid),), conn)
            for uid in set(int(u) for u in usuario_ids):
                ejecutar("INSERT INTO usuario_contratos (usuario_id, contrato_id) VALUES (?, ?)", (uid, int(cid)), conn)
    
    
    def contratos_ids_de_usuario(uid: int) -> list[int]:
        return consultar("SELECT contrato_id FROM usuario_contratos WHERE usuario_id = ?", (int(uid),))["contrato_id"].astype(int).tolist()
    
    
    def asignar_contratos_a_usuario(uid: int, contrato_ids: list[int]):
        with transaccion() as conn:
            ejecutar("DELETE FROM usuario_contratos WHERE usuario_id = ?", (int(uid),), conn)
            for cid in set(int(c) for c in contrato_ids):
                ejecutar("INSERT INTO usuario_contratos (usuario_id, contrato_id) VALUES (?, ?)", (int(uid), cid), conn)
    
    
    def resumen_datos_contrato(cid: int) -> dict:
        """Cuántos registros tiene un contrato en cada módulo (para confirmar antes de borrarlo)."""
        nombres = {"proyectos": "proyectos P6", "informes": "informes semanales", "restricciones": "restricciones 6WLA",
                   "diario_informes": "informes diarios emitidos", "diario_reportes": "reportes de terreno",
                   "tm_cargas": "cargas T&M", "desv_cargas": "cargas del LOG de desviaciones",
                   "semanal_archivos": "informes semanales generados",
                   "cubicacion_cargas": "planillas de avance cargadas (cubicación)",
                   "cubicacion_planos": "cubicaciones hechas desde planos"}
        return {n: uno(f"SELECT COUNT(*) FROM {t} WHERE contrato_id = ?", (int(cid),))[0] for t, n in nombres.items()}
    
    
    def borrar_contrato(cid: int) -> str | None:
        """Borra el contrato y TODA su información. Antes deja una copia de la base (SQLite); devuelve su ruta."""
        cid = int(cid)
        copia = None
        if not ES_POSTGRES and os.path.exists(SQLITE_PATH):
            fila = uno("SELECT codigo FROM contratos WHERE id = ?", (cid,))
            etiqueta = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in (fila[0] if fila else str(cid)))
            copia = SQLITE_PATH.replace(".db", "") + f"_antes_de_borrar_{etiqueta}_{datetime.now():%Y%m%d_%H%M%S}.db"
            with open(copia, "wb") as f:
                f.write(respaldo_sqlite())
        with transaccion() as conn:
            for (pid,) in ejecutar("SELECT id FROM proyectos WHERE contrato_id = ?", (cid,), conn).fetchall():
                for (corte,) in ejecutar("SELECT id FROM cortes WHERE proyecto_id = ?", (pid,), conn).fetchall():
                    _borrar_corte(corte, conn)
            for (iid,) in ejecutar("SELECT id FROM informes WHERE contrato_id = ?", (cid,), conn).fetchall():
                ejecutar("DELETE FROM informe_items WHERE informe_id = ?", (iid,), conn)
                ejecutar("DELETE FROM diario_config WHERE clave = ?", (f"resumen_informe_{iid}",), conn)
            for t in TABLAS_POR_CONTRATO:
                ejecutar(f"DELETE FROM {t} WHERE contrato_id = ?", (cid,), conn)
            ejecutar("DELETE FROM diario_config WHERE clave LIKE ?", (f"c{cid}:%",), conn)
            ejecutar("DELETE FROM usuario_contratos WHERE contrato_id = ?", (cid,), conn)
            ejecutar("DELETE FROM contratos WHERE id = ?", (cid,), conn)
        return copia
    
    
    def config(contrato_id: int, clave: str, defecto=None):
        """Configuración propia de un contrato (calendario, encabezados del informe diario, valor HH…)."""
        return config_diario(f"c{int(contrato_id)}:{clave}", defecto)
    
    
    def guardar_config(contrato_id: int, clave: str, valor):
        guardar_config_diario(f"c{int(contrato_id)}:{clave}", valor)
    
    
    # ---------------------------------------------------------------- Proyectos P6
    
    def guardar_corte_p6(proy, archivo: str, usuario: str, reemplazar: bool = True, *, contrato_id: int) -> tuple[int, bool]:
        """Guarda un proyecto leído del .XER como un corte (fecha de datos).
    
        Si ya existe un corte con la misma fecha, se reemplaza. Devuelve (corte_id, reemplazado).
        Nunca borra otros proyectos ni otros cortes.
        """
        fecha = proy.fecha_corte or datetime.now().strftime("%Y-%m-%d")
        with transaccion() as conn:
            fila = ejecutar("SELECT id, contrato_id FROM proyectos WHERE codigo = ?", (proy.codigo,), conn).fetchone()
            if fila and fila[1] is not None and int(fila[1]) != int(contrato_id):
                otro = ejecutar("SELECT codigo FROM contratos WHERE id = ?", (fila[1],), conn).fetchone()
                raise ValueError(f"El proyecto {proy.codigo} ya está cargado en el contrato {otro[0] if otro else fila[1]}.")
            if fila:
                proyecto_id = fila[0]
                ejecutar("UPDATE proyectos SET nombre = ? WHERE id = ?", (proy.nombre, proyecto_id), conn)
            else:
                proyecto_id = insertar(
                    "INSERT INTO proyectos (codigo, nombre, origen, creado_en, contrato_id) VALUES (?, ?, ?, ?, ?)",
                    (proy.codigo, proy.nombre, "P6", ahora(), int(contrato_id)), conn,
                )
            previo = ejecutar(
                "SELECT id FROM cortes WHERE proyecto_id = ? AND fecha_corte = ?", (proyecto_id, fecha), conn
            ).fetchone()
            if previo:
                if not reemplazar:
                    raise ValueError(f"Ya existe un corte al {fecha} para {proy.codigo}.")
                _borrar_corte(previo[0], conn)
    
            corte_id = insertar(
                """INSERT INTO cortes (proyecto_id, fecha_corte, fecha_inicio, fecha_termino, archivo,
                       cargado_por, cargado_en, budget, actual, remaining, n_actividades)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (proyecto_id, fecha, proy.fecha_inicio, proy.fecha_termino, archivo, usuario, ahora(),
                 proy.budget, proy.actual, proy.remaining, len(proy.actividades)),
                conn,
            )
            insertar_muchos(
                """INSERT INTO actividades (corte_id, task_code, nombre, tipo, estado, wbs, budget, actual,
                       remaining, inicio, fin, holgura_total_hr) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                [(corte_id, a["task_code"], a["nombre"], a["tipo"], a["estado"], a["wbs"], a["budget"],
                  a["actual"], a["remaining"], a["inicio"], a["fin"], a["holgura_total_hr"])
                 for a in proy.actividades],
                conn,
            )
            insertar_muchos(
                "INSERT INTO actividad_codigos (corte_id, task_code, tipo_codigo, valor) VALUES (?, ?, ?, ?)",
                [(corte_id, tc, tipo, valor) for tc, tipo, valor in proy.codigos],
                conn,
            )
        return corte_id, bool(previo)
    
    
    def _borrar_corte(corte_id: int, conn):
        # Borrado explícito: no depende de que el motor tenga ON DELETE CASCADE activo.
        for tabla in ("actividad_codigos", "actividades"):
            ejecutar(f"DELETE FROM {tabla} WHERE corte_id = ?", (corte_id,), conn)
        ejecutar("DELETE FROM cortes WHERE id = ?", (corte_id,), conn)
    
    
    def borrar_corte(corte_id: int):
        with transaccion() as conn:
            _borrar_corte(corte_id, conn)
    
    
    def borrar_proyecto(proyecto_id: int):
        with transaccion() as conn:
            for (cid,) in ejecutar("SELECT id FROM cortes WHERE proyecto_id = ?", (proyecto_id,), conn).fetchall():
                _borrar_corte(cid, conn)
            ejecutar("UPDATE informes SET proyecto_id = NULL WHERE proyecto_id = ?", (proyecto_id,), conn)
            ejecutar("DELETE FROM proyectos WHERE id = ?", (proyecto_id,), conn)
    
    
    def proyectos(contrato_id: int) -> pd.DataFrame:
        return consultar(
            """SELECT p.id, p.codigo, p.nombre, p.origen, COUNT(c.id) AS cortes, MAX(c.fecha_corte) AS ultimo_corte
               FROM proyectos p LEFT JOIN cortes c ON c.proyecto_id = p.id
               WHERE p.contrato_id = ?
               GROUP BY p.id, p.codigo, p.nombre, p.origen ORDER BY p.codigo""", (int(contrato_id),)
        )
    
    
    def cortes(proyecto_id: int) -> pd.DataFrame:
        df = consultar(
            """SELECT id, fecha_corte, fecha_inicio, fecha_termino, archivo, cargado_por, cargado_en,
                      budget, actual, remaining, n_actividades
               FROM cortes WHERE proyecto_id = ? ORDER BY fecha_corte""",
            (proyecto_id,),
        )
        if not df.empty:
            df["at_completion"] = df["actual"] + df["remaining"]
            df["avance_pct"] = (df["actual"] / df["at_completion"].where(df["at_completion"] > 0) * 100).fillna(0)
        return df
    
    
    def actividades(corte_id: int) -> pd.DataFrame:
        df = consultar(
            """SELECT task_code, nombre, tipo, estado, wbs, budget, actual, remaining, inicio, fin, holgura_total_hr
               FROM actividades WHERE corte_id = ? ORDER BY task_code""",
            (corte_id,),
        )
        if not df.empty:
            df["at_completion"] = df["actual"] + df["remaining"]
            df["avance_pct"] = (df["actual"] / df["at_completion"].where(df["at_completion"] > 0) * 100).fillna(0)
        return df
    
    
    def tipos_codigo(corte_id: int) -> list[str]:
        df = consultar("SELECT DISTINCT tipo_codigo FROM actividad_codigos WHERE corte_id = ? ORDER BY tipo_codigo", (corte_id,))
        return df["tipo_codigo"].tolist()
    
    
    def codigos(corte_id: int, tipo_codigo: str) -> pd.DataFrame:
        return consultar(
            "SELECT task_code, valor FROM actividad_codigos WHERE corte_id = ? AND tipo_codigo = ?",
            (corte_id, tipo_codigo),
        )
    
    
    # ---------------------------------------------------------- Informes semanales
    
    def guardar_informe(nombre, proyecto_id, fecha_corte, archivo, hoja, usuario, items: pd.DataFrame, *, contrato_id: int) -> int:
        with transaccion() as conn:
            inf_id = insertar(
                """INSERT INTO informes (nombre, proyecto_id, fecha_corte, archivo, hoja, cargado_por, cargado_en, contrato_id)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (nombre, proyecto_id, fecha_corte, archivo, hoja, usuario, ahora(), int(contrato_id)), conn,
            )
            insertar_muchos(
                "INSERT INTO informe_items (informe_id, item, at_completion, remaining, avance_pct, hh_gastadas) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                [(inf_id, str(r.item), float(r.at_completion), float(r.remaining), float(r.avance_pct),
                  None if pd.isna(getattr(r, "hh_gastadas", None)) else float(r.hh_gastadas))
                 for r in items.itertuples()],
                conn,
            )
        return inf_id
    
    
    def informes(contrato_id: int) -> pd.DataFrame:
        return consultar(
            """SELECT i.id, i.nombre, i.fecha_corte, p.codigo AS proyecto, i.archivo, i.hoja, i.cargado_por, i.cargado_en,
                      COUNT(t.id) AS items, COALESCE(SUM(t.at_completion), 0) AS at_completion,
                      COALESCE(SUM(t.remaining), 0) AS remaining, SUM(t.hh_gastadas) AS hh_gastadas
               FROM informes i
               LEFT JOIN proyectos p ON p.id = i.proyecto_id
               LEFT JOIN informe_items t ON t.informe_id = i.id
               WHERE i.contrato_id = ?
               GROUP BY i.id, i.nombre, i.fecha_corte, p.codigo, i.archivo, i.hoja, i.cargado_por, i.cargado_en
               ORDER BY i.fecha_corte, i.id""", (int(contrato_id),)
        )
    
    
    def informe_items(ids: list[int]) -> pd.DataFrame:
        if not ids:
            return pd.DataFrame(columns=["informe_id", "nombre", "fecha_corte", "item", "at_completion", "remaining",
                                         "avance_pct", "hh_gastadas"])
        marcas = ",".join("?" * len(ids))
        return consultar(
            f"""SELECT t.informe_id, i.nombre, i.fecha_corte, t.item, t.at_completion, t.remaining, t.avance_pct, t.hh_gastadas
                FROM informe_items t JOIN informes i ON i.id = t.informe_id
                WHERE t.informe_id IN ({marcas}) ORDER BY i.fecha_corte, t.id""",
            tuple(int(i) for i in ids),
        )
    
    
    def borrar_informe(inf_id: int):
        with transaccion() as conn:
            ejecutar("DELETE FROM informe_items WHERE informe_id = ?", (inf_id,), conn)
            ejecutar("DELETE FROM informes WHERE id = ?", (inf_id,), conn)
    
    
    # ---------------------------------------------------------------- 6WLA
    
    COLS_6WLA = ["proyecto", "actividad", "subsistema", "semana", "restriccion", "responsable", "estado", "fecha_compromiso"]
    
    
    def guardar_restricciones(df: pd.DataFrame, usuario: str, *, contrato_id: int):
        with transaccion() as conn:
            insertar_muchos(
                f"""INSERT INTO restricciones ({", ".join(COLS_6WLA)}, actualizado_por, actualizado_en, contrato_id)
                    VALUES ({", ".join("?" * (len(COLS_6WLA) + 3))})""",
                [tuple(None if pd.isna(r[c]) else str(r[c]) for c in COLS_6WLA) + (usuario, ahora(), int(contrato_id))
                 for _, r in df.iterrows()],
                conn,
            )
    
    
    def restricciones(contrato_id: int, proyecto: str | None = None) -> pd.DataFrame:
        q = f"SELECT id, {', '.join(COLS_6WLA)}, actualizado_por, actualizado_en FROM restricciones WHERE contrato_id = ?"
        if proyecto:
            return consultar(q + " AND proyecto = ? ORDER BY semana, id", (int(contrato_id), proyecto))
        return consultar(q + " ORDER BY proyecto, semana, id", (int(contrato_id),))
    
    
    def actualizar_restricciones(cambios: list[tuple[int, str, str, str | None]], usuario: str):
        """cambios: [(id, estado, responsable, fecha_compromiso)]"""
        with transaccion() as conn:
            for rid, estado, responsable, fecha in cambios:
                ejecutar(
                    """UPDATE restricciones SET estado = ?, responsable = ?, fecha_compromiso = ?,
                           actualizado_por = ?, actualizado_en = ? WHERE id = ?""",
                    (estado, responsable, fecha, usuario, ahora(), int(rid)), conn,
                )
    
    
    def borrar_restricciones(proyecto: str, *, contrato_id: int):
        ejecutar("DELETE FROM restricciones WHERE proyecto = ? AND contrato_id = ?", (proyecto, int(contrato_id)))
    
    
    # ---------------------------------------------------------------- Informe diario
    
    def _b64(datos: bytes) -> str:
        import base64
        return base64.b64encode(datos).decode()
    
    
    def _de_b64(texto: str) -> bytes:
        import base64
        return base64.b64decode(texto)
    
    
    ARCHIVOS_A_CONSERVAR = 4
    
    
    def guardar_base_diario(nombre: str, fecha: str, numero: int, datos: bytes, usuario: str, tipo: str = "base", *,
                            contrato_id: int) -> int:
        """Guarda la planilla base vigente (la última versión queda como base del próximo informe)."""
        with transaccion() as conn:
            aid = insertar(
                "INSERT INTO diario_archivos (tipo, nombre, fecha, numero, contenido, creado_por, creado_en, contrato_id) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                (tipo, nombre, fecha, numero, _b64(datos), usuario, ahora(), int(contrato_id)), conn,
            )
            # Solo se conservan las últimas versiones de cada contrato (cada archivo pesa varios MB)
            viejos = ejecutar("SELECT id FROM diario_archivos WHERE contrato_id = ? ORDER BY id DESC", (int(contrato_id),),
                              conn).fetchall()[ARCHIVOS_A_CONSERVAR:]
            for (vid,) in viejos:
                ejecutar("DELETE FROM diario_archivos WHERE id = ?", (vid,), conn)
                ejecutar("UPDATE diario_informes SET archivo_id = NULL WHERE archivo_id = ?", (vid,), conn)
        return aid
    
    
    def guardar_base_semanal(nombre: str, fecha: str, numero: int, datos: bytes, usuario: str, tipo: str = "base", *,
                             contrato_id: int) -> int:
        """Guarda el informe semanal que sirve de base para la semana siguiente (se conservan las últimas versiones)."""
        with transaccion() as conn:
            aid = insertar(
                "INSERT INTO semanal_archivos (tipo, nombre, fecha, numero, contenido, creado_por, creado_en, contrato_id) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                (tipo, nombre, fecha, numero, _b64(datos), usuario, ahora(), int(contrato_id)), conn,
            )
            viejos = ejecutar("SELECT id FROM semanal_archivos WHERE contrato_id = ? ORDER BY id DESC", (int(contrato_id),),
                              conn).fetchall()[ARCHIVOS_A_CONSERVAR:]
            for (vid,) in viejos:
                ejecutar("DELETE FROM semanal_archivos WHERE id = ?", (vid,), conn)
        return aid
    
    
    def base_semanal_vigente(contrato_id: int) -> dict | None:
        fila = uno("SELECT id, tipo, nombre, fecha, numero, creado_por, creado_en FROM semanal_archivos WHERE contrato_id = ? "
                   "ORDER BY id DESC LIMIT 1", (int(contrato_id),))
        if not fila:
            return None
        return dict(zip(["id", "tipo", "nombre", "fecha", "numero", "creado_por", "creado_en"], fila))
    
    
    def archivo_semanal(aid: int) -> bytes | None:
        fila = uno("SELECT contenido FROM semanal_archivos WHERE id = ?", (int(aid),))
        return _de_b64(fila[0]) if fila else None
    
    
    def archivos_semanal(contrato_id: int) -> pd.DataFrame:
        return consultar("SELECT id, tipo, nombre, fecha, numero, creado_por, creado_en FROM semanal_archivos WHERE contrato_id = ? "
                         "ORDER BY id DESC", (int(contrato_id),))
    
    
    # ---------------------------------------------------------------- Cubicación (planillas de avance por disciplina)
    
    def guardar_carga_cubicacion(fecha: str, archivo: str, disciplinas: str, datos: dict, usuario: str, *,
                                 contrato_id: int) -> int:
        """Guarda lo leído de una planilla de avance (JSON, no el archivo). Reemplaza la carga del mismo archivo y fecha."""
        import json
        with transaccion() as conn:
            ejecutar("DELETE FROM cubicacion_cargas WHERE contrato_id = ? AND fecha = ? AND archivo = ?",
                     (int(contrato_id), fecha, archivo), conn)
            return insertar(
                "INSERT INTO cubicacion_cargas (fecha, archivo, disciplinas, datos, creado_por, creado_en, contrato_id) "
                "VALUES (?, ?, ?, ?, ?, ?, ?)",
                (fecha, archivo, disciplinas, json.dumps(datos, ensure_ascii=False), usuario, ahora(), int(contrato_id)), conn)
    
    
    def guardar_cub_plano(plano: str, rev: str, disciplina: str, descripcion: str, datos: dict, hh: float,
                          usuario: str, *, contrato_id: int) -> int:
        """Guarda la cubicación de un plano. Reemplaza la del mismo plano y revisión."""
        import json
        with transaccion() as conn:
            ejecutar("DELETE FROM cubicacion_planos WHERE contrato_id = ? AND plano = ? AND COALESCE(rev, '') = ?",
                     (int(contrato_id), plano, rev or ""), conn)
            return insertar(
                "INSERT INTO cubicacion_planos (plano, rev, disciplina, descripcion, datos, hh, creado_por, creado_en, "
                "contrato_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (plano, rev or "", disciplina, descripcion, json.dumps(datos, ensure_ascii=False, default=str), float(hh),
                 usuario, ahora(), int(contrato_id)), conn)
    
    
    def cub_planos(contrato_id: int) -> pd.DataFrame:
        return consultar("SELECT id, plano, rev, disciplina, descripcion, hh, creado_por, creado_en FROM cubicacion_planos "
                         "WHERE contrato_id = ? ORDER BY plano, rev", (int(contrato_id),))
    
    
    def cub_plano(cid_cub: int) -> dict | None:
        import json
        fila = uno("SELECT datos FROM cubicacion_planos WHERE id = ?", (int(cid_cub),))
        return json.loads(fila[0]) if fila else None
    
    
    def borrar_cub_plano(cid_cub: int, *, contrato_id: int):
        ejecutar("DELETE FROM cubicacion_planos WHERE id = ? AND contrato_id = ?", (int(cid_cub), int(contrato_id)))
    
    
    def cargas_cubicacion(contrato_id: int) -> pd.DataFrame:
        return consultar("SELECT id, fecha, archivo, disciplinas, creado_por, creado_en FROM cubicacion_cargas "
                         "WHERE contrato_id = ? ORDER BY fecha DESC, id", (int(contrato_id),))
    
    
    def carga_cubicacion(cid_carga: int) -> dict | None:
        import json
        fila = uno("SELECT datos FROM cubicacion_cargas WHERE id = ?", (int(cid_carga),))
        return json.loads(fila[0]) if fila else None
    
    
    def borrar_carga_cubicacion(cid_carga: int, *, contrato_id: int):
        ejecutar("DELETE FROM cubicacion_cargas WHERE id = ? AND contrato_id = ?", (int(cid_carga), int(contrato_id)))
    
    
    def base_diario_vigente(contrato_id: int) -> dict | None:
        fila = uno("SELECT id, tipo, nombre, fecha, numero, creado_por, creado_en FROM diario_archivos WHERE contrato_id = ? "
                   "ORDER BY id DESC LIMIT 1", (int(contrato_id),))
        if not fila:
            return None
        return dict(zip(["id", "tipo", "nombre", "fecha", "numero", "creado_por", "creado_en"], fila))
    
    
    def archivo_diario(aid: int) -> bytes | None:
        fila = uno("SELECT contenido FROM diario_archivos WHERE id = ?", (int(aid),))
        return _de_b64(fila[0]) if fila else None
    
    
    def archivos_diario(contrato_id: int) -> pd.DataFrame:
        return consultar("SELECT id, tipo, nombre, fecha, numero, creado_por, creado_en FROM diario_archivos WHERE contrato_id = ? "
                         "ORDER BY id DESC", (int(contrato_id),))
    
    
    def guardar_reporte_diario(fecha: str, archivo: str, datos: dict, usuario: str, *, contrato_id: int) -> int:
        import json
        with transaccion() as conn:
            previo = ejecutar("SELECT id FROM diario_reportes WHERE fecha = ? AND archivo = ? AND contrato_id = ?",
                              (fecha, archivo, int(contrato_id)), conn).fetchone()
            if previo:
                ejecutar("UPDATE diario_reportes SET datos = ?, cargado_por = ?, cargado_en = ? WHERE id = ?",
                         (json.dumps(datos, ensure_ascii=False), usuario, ahora(), previo[0]), conn)
                return previo[0]
            return insertar("INSERT INTO diario_reportes (fecha, archivo, datos, cargado_por, cargado_en, contrato_id) VALUES (?, ?, ?, ?, ?, ?)",
                            (fecha, archivo, json.dumps(datos, ensure_ascii=False), usuario, ahora(), int(contrato_id)), conn)
    
    
    def reportes_diario(fecha: str, contrato_id: int) -> list[dict]:
        import json
        df = consultar("SELECT id, archivo, datos, cargado_por, cargado_en FROM diario_reportes WHERE fecha = ? AND contrato_id = ? "
                       "ORDER BY id", (fecha, int(contrato_id)))
        return [{"id": int(r.id), "archivo": r.archivo, "datos": json.loads(r.datos), "cargado_por": r.cargado_por,
                 "cargado_en": r.cargado_en} for r in df.itertuples()]
    
    
    def actualizar_reporte_diario(rid: int, datos: dict):
        import json
        ejecutar("UPDATE diario_reportes SET datos = ? WHERE id = ?", (json.dumps(datos, ensure_ascii=False), int(rid)))
    
    
    def borrar_reporte_diario(rid: int):
        ejecutar("DELETE FROM diario_reportes WHERE id = ?", (int(rid),))
    
    
    def registrar_informe_diario(fecha: str, numero: int, resumen: dict, archivo_id: int, usuario: str, *, contrato_id: int):
        import json
        ejecutar("INSERT INTO diario_informes (fecha, numero, resumen, archivo_id, emitido_por, emitido_en, contrato_id) "
                 "VALUES (?, ?, ?, ?, ?, ?, ?)",
                 (fecha, int(numero), json.dumps(resumen, ensure_ascii=False), archivo_id, usuario, ahora(), int(contrato_id)))
    
    
    def informes_diario(contrato_id: int) -> pd.DataFrame:
        import json
        df = consultar("SELECT id, fecha, numero, resumen, archivo_id, emitido_por, emitido_en FROM diario_informes WHERE contrato_id = ? "
                       "ORDER BY fecha DESC, id DESC", (int(contrato_id),))
        if not df.empty:
            res = df["resumen"].map(json.loads)
            for k in ("directos_dia", "directos_noche", "indirectos_dia", "indirectos_noche", "hh_mod", "hh_moi", "reportes"):
                df[k] = res.map(lambda d: d.get(k))
        return df.drop(columns=["resumen"]) if not df.empty else df
    
    
    def config_diario(clave: str, defecto=None):
        import json
        fila = uno("SELECT valor FROM diario_config WHERE clave = ?", (clave,))
        return json.loads(fila[0]) if fila else defecto
    
    
    def guardar_config_diario(clave: str, valor):
        import json
        with transaccion() as conn:
            if ejecutar("SELECT clave FROM diario_config WHERE clave = ?", (clave,), conn).fetchone():
                ejecutar("UPDATE diario_config SET valor = ? WHERE clave = ?", (json.dumps(valor, ensure_ascii=False), clave), conn)
            else:
                ejecutar("INSERT INTO diario_config (clave, valor) VALUES (?, ?)", (clave, json.dumps(valor, ensure_ascii=False)), conn)
    
    
    # ---------------------------------------------------------------- Times & Material
    
    def guardar_tm(archivo: str, fecha_corte: str | None, contrato: str, datos: dict, usuario: str, *, contrato_id: int) -> int:
        import json
        with transaccion() as conn:
            return insertar("INSERT INTO tm_cargas (archivo, fecha_corte, contrato, datos, cargado_por, cargado_en, contrato_id) "
                            "VALUES (?, ?, ?, ?, ?, ?, ?)",
                            (archivo, fecha_corte, contrato, json.dumps(datos, ensure_ascii=False, default=str), usuario, ahora(),
                             int(contrato_id)), conn)
    
    
    def cargas_tm(contrato_id: int) -> pd.DataFrame:
        return consultar("SELECT id, archivo, fecha_corte, contrato, cargado_por, cargado_en FROM tm_cargas WHERE contrato_id = ? "
                         "ORDER BY id DESC", (int(contrato_id),))
    
    
    def carga_tm(cid: int) -> dict | None:
        import json
        fila = uno("SELECT datos FROM tm_cargas WHERE id = ?", (int(cid),))
        return json.loads(fila[0]) if fila else None
    
    
    def borrar_carga_tm(cid: int):
        ejecutar("DELETE FROM tm_cargas WHERE id = ?", (int(cid),))
    
    
    # ---------------------------------------------------------------- LOG de desviación de recursos
    
    def guardar_desviaciones(archivo: str, fecha_corte: str | None, revision: str, hh: float, registros: int, datos: dict,
                             usuario: str, *, contrato_id: int) -> int:
        import json
        with transaccion() as conn:
            return insertar("INSERT INTO desv_cargas (archivo, fecha_corte, revision, hh, registros, datos, cargado_por, cargado_en, "
                            "contrato_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                            (archivo, fecha_corte, revision, hh, registros, json.dumps(datos, ensure_ascii=False, default=str),
                             usuario, ahora(), int(contrato_id)), conn)
    
    
    def cargas_desviaciones(contrato_id: int) -> pd.DataFrame:
        return consultar("SELECT id, archivo, fecha_corte, revision, hh, registros, cargado_por, cargado_en FROM desv_cargas "
                         "WHERE contrato_id = ? ORDER BY fecha_corte DESC, id DESC", (int(contrato_id),))
    
    
    def carga_desviaciones(cid: int) -> dict | None:
        import json
        fila = uno("SELECT datos FROM desv_cargas WHERE id = ?", (int(cid),))
        return json.loads(fila[0]) if fila else None
    
    
    def borrar_carga_desviaciones(cid: int):
        ejecutar("DELETE FROM desv_cargas WHERE id = ?", (int(cid),))
    
    
    # ---------------------------------------------------------------- respaldo y restauración (SQLite)
    
    TABLAS_CLAVE = ("usuarios", "proyectos", "cortes")
    
    
    def respaldo_sqlite() -> bytes:
        """Copia consistente de la base SQLite (aunque la app esté en uso)."""
        import os
        import tempfile
        if ES_POSTGRES:
            raise RuntimeError("El respaldo desde la app solo está disponible con SQLite; en PostgreSQL usa los respaldos del proveedor.")
        with tempfile.TemporaryDirectory() as tmp:
            destino = os.path.join(tmp, "respaldo.db")
            origen = sqlite3.connect(SQLITE_PATH)
            copia = sqlite3.connect(destino)
            try:
                origen.backup(copia)
            finally:
                copia.close()
                origen.close()
            with open(destino, "rb") as f:
                return f.read()
    
    
    def resumen_sqlite(datos: bytes) -> dict:
        """Valida un respaldo y cuenta sus registros, sin tocar la base en uso."""
        import os
        import tempfile
        if not datos.startswith(b"SQLite format 3"):
            raise ValueError("El archivo no es un respaldo válido de SYNCORE (.db).")
        with tempfile.TemporaryDirectory() as tmp:
            ruta = os.path.join(tmp, "r.db")
            with open(ruta, "wb") as f:
                f.write(datos)
            conn = sqlite3.connect(ruta)
            try:
                if conn.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
                    raise ValueError("El respaldo está dañado (falló la verificación de integridad).")
                tablas = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
                if "usuarios" not in tablas:
                    raise ValueError("El archivo no tiene la estructura de SYNCORE (falta la tabla de usuarios).")
                cuenta = {}
                for t in ("contratos", "usuarios", "proyectos", "cortes", "informes", "restricciones", "diario_informes", "diario_archivos", "tm_cargas", "desv_cargas"):
                    if t in tablas:
                        cuenta[t] = conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
                return cuenta
            finally:
                conn.close()
    
    
    def restaurar_sqlite(datos: bytes) -> str:
        """Reemplaza la base en uso por el respaldo. Antes guarda una copia de la base actual junto a ella."""
        import os
        import tempfile
        if ES_POSTGRES:
            raise RuntimeError("La restauración desde la app solo está disponible con SQLite.")
        resumen_sqlite(datos)
        copia_previa = SQLITE_PATH.replace(".db", "") + f"_antes_de_restaurar_{datetime.now():%Y%m%d_%H%M%S}.db"
        if os.path.exists(SQLITE_PATH):
            with open(copia_previa, "wb") as f:
                f.write(respaldo_sqlite())
        with tempfile.TemporaryDirectory() as tmp:
            ruta = os.path.join(tmp, "r.db")
            with open(ruta, "wb") as f:
                f.write(datos)
            origen = sqlite3.connect(ruta)
            destino = sqlite3.connect(SQLITE_PATH, timeout=30)
            try:
                origen.backup(destino)
            finally:
                destino.close()
                origen.close()
        init_db()  # agrega tablas o columnas nuevas si el respaldo es de una versión anterior
        return copia_previa

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


db = _modulo_db()

# ======================================================================
# MÓDULO: calendario
# ======================================================================
def _modulo_calendario():
    from contextvars import ContextVar
    from datetime import date, datetime, timedelta
    from functools import lru_cache
    
    import pandas as pd
    
    
    INICIO_DEFECTO = date(2026, 3, 9)
    ANCLAS_DEFECTO = ((1, date(2026, 3, 21)), (30, date(2026, 10, 9)))
    _CLAVE = "calendario_semanas"
    
    
    # Contrato en uso en esta ejecución (cada sesión de la app corre en su propio hilo, así que no se mezclan usuarios).
    _CONTRATO: ContextVar[int | None] = ContextVar("contrato_calendario", default=None)
    
    
    def usar_contrato(contrato_id: int | None):
        """Fija el contrato cuyo calendario usan las funciones de este módulo (lo llama la app al inicio de cada carga)."""
        _CONTRATO.set(int(contrato_id) if contrato_id is not None else None)
    
    
    def contrato_actual() -> int | None:
        return _CONTRATO.get()
    
    
    def configuracion(contrato_id: int | None = None) -> tuple[date, tuple[tuple[int, date], ...]]:
        """(inicio del proyecto, anclas (semana, fecha de corte) ordenadas por semana) del contrato."""
        cid = contrato_id if contrato_id is not None else _CONTRATO.get()
        try:
            v = db.config(cid, _CLAVE) if cid is not None else db.config_diario(_CLAVE)
            if v:
                ini = date.fromisoformat(v["inicio"])
                anclas = tuple(sorted((int(n), date.fromisoformat(f)) for n, f in v["anclas"]))
                if anclas and anclas[0][0] == 1:
                    return ini, anclas
        except Exception:
            pass
        return INICIO_DEFECTO, ANCLAS_DEFECTO
    
    
    def validar(ini: date, anclas: list[tuple[int, date]]) -> str | None:
        anclas = sorted(anclas)
        if not anclas or anclas[0][0] != 1:
            return "Debe indicarse la fecha de corte de la semana 1."
        if anclas[0][1] < ini:
            return "La semana 1 no puede terminar antes del inicio del proyecto."
        for (n1, f1), (n2, f2) in zip(anclas, anclas[1:]):
            if n2 == n1:
                return f"La semana {n2} aparece dos veces."
            dias = (f2 - f1).days
            if dias <= 7 * (n2 - n1 - 1):
                return f"La semana {n2} terminaría antes de que empiece (revisa su fecha de corte)."
            if dias > 7 * (n2 - n1) + 6:
                return f"La semana {n2} tendría más de 13 días (revisa su fecha de corte)."
        return None
    
    
    def guardar(ini: date, anclas: list[tuple[int, date]], contrato_id: int | None = None):
        err = validar(ini, anclas)
        if err:
            raise ValueError(err)
        cid = contrato_id if contrato_id is not None else _CONTRATO.get()
        valor = {"inicio": ini.isoformat(), "anclas": [[int(n), f.isoformat()] for n, f in sorted(anclas)]}
        if cid is None:
            db.guardar_config_diario(_CLAVE, valor)
        else:
            db.guardar_config(cid, _CLAVE, valor)
    
    
    def inicio() -> date:
        return configuracion()[0]
    
    
    def _a_fecha(f) -> date | None:
        if f is None or (isinstance(f, float) and pd.isna(f)):
            return None
        if isinstance(f, datetime):
            return f.date()
        if isinstance(f, date):
            return f
        try:
            return pd.Timestamp(str(f)[:10]).date()
        except (ValueError, TypeError):
            return None
    
    
    def _fin(n: int, anclas) -> date:
        k, f = anclas[0]
        for kk, ff in anclas:
            if kk <= n:
                k, f = kk, ff
        return f + timedelta(days=7 * (n - k))
    
    
    @lru_cache(maxsize=4096)
    def _semana(f: date, ini: date, anclas) -> int:
        if f < ini:
            return 0
        k, d = anclas[0]
        for kk, ff in anclas:
            if ff < f:
                k, d = kk, ff
        n = k if f <= d else k + -(-(f - d).days // 7)
        while n > 1 and _fin(n - 1, anclas) >= f:          # por si una ancla posterior acorta una semana
            n -= 1
        while _fin(n, anclas) < f:
            n += 1
        return n
    
    
    def semana(f, cfg=None) -> int | None:
        """Número de semana del proyecto (1 = primera semana; 0 = antes del inicio)."""
        f = _a_fecha(f)
        if f is None:
            return None
        ini, anclas = cfg or configuracion()
        return _semana(f, ini, anclas)
    
    
    def fin_semana(n: int, cfg=None) -> date:
        return _fin(n, (cfg or configuracion())[1])
    
    
    def inicio_semana(n: int, cfg=None) -> date:
        ini, anclas = cfg or configuracion()
        return ini if n <= 1 else _fin(n - 1, anclas) + timedelta(days=1)
    
    
    def dias_semana(n: int, cfg=None) -> int:
        return (fin_semana(n, cfg) - inicio_semana(n, cfg)).days + 1
    
    
    def etiqueta(f, cfg=None, con_fechas: bool = False) -> str:
        cfg = cfg or configuracion()
        n = semana(f, cfg)
        if n is None:
            return "–"
        txt = f"S{n:02d}"
        if con_fechas and n >= 1:
            txt += f" ({inicio_semana(n, cfg):%d-%m} al {fin_semana(n, cfg):%d-%m})"
        return txt
    
    
    def texto(f, cfg=None) -> str:
        """'Semana 33 del proyecto' (vacío si la fecha no es válida o es anterior al inicio)."""
        n = semana(f, cfg)
        return f"Semana {n} del proyecto" if n is not None and n >= 1 else ""
    
    
    def columna_semana(fechas: pd.Series, cfg=None) -> pd.Series:
        cfg = cfg or configuracion()
        return fechas.map(lambda f: semana(f, cfg))
    
    
    def tabla(desde_semana: int = 1, hasta_semana: int | None = None, cfg=None) -> pd.DataFrame:
        cfg = cfg or configuracion()
        hasta_semana = hasta_semana or max(semana(date.today(), cfg) + 6, desde_semana)
        filas = [{"Semana": f"S{n:02d}", "Desde": inicio_semana(n, cfg), "Hasta (corte)": fin_semana(n, cfg),
                  "Días": dias_semana(n, cfg)} for n in range(desde_semana, hasta_semana + 1)]
        return pd.DataFrame(filas)

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


calendario = _modulo_calendario()

# ======================================================================
# MÓDULO: auth
# ======================================================================
def _modulo_auth():
    import base64
    import hashlib
    import hmac
    import os
    import secrets
    from datetime import datetime, timedelta
    
    
    ROLES = {
        "admin": "Administrador",
        "planificador": "Planificador (carga y edita)",
        "lector": "Lector (solo consulta)",
    }
    MAX_INTENTOS = 5
    MINUTOS_BLOQUEO = 15
    LARGO_MINIMO = 8
    
    _N, _R, _P = 2**14, 8, 1
    
    
    def hash_password(password: str) -> str:
        salt = secrets.token_bytes(16)
        h = hashlib.scrypt(password.encode(), salt=salt, n=_N, r=_R, p=_P, dklen=32)
        return f"scrypt${_N}${_R}${_P}${base64.b64encode(salt).decode()}${base64.b64encode(h).decode()}"
    
    
    def verificar_password(password: str, guardado: str) -> bool:
        try:
            _, n, r, p, salt, h = guardado.split("$")
            calc = hashlib.scrypt(password.encode(), salt=base64.b64decode(salt), n=int(n), r=int(r), p=int(p),
                                  dklen=len(base64.b64decode(h)))
            return hmac.compare_digest(calc, base64.b64decode(h))
        except (ValueError, TypeError):
            return False
    
    
    def validar_password(password: str) -> str | None:
        if len(password) < LARGO_MINIMO:
            return f"La contraseña debe tener al menos {LARGO_MINIMO} caracteres."
        if password.isalpha() or password.isdigit():
            return "La contraseña debe combinar letras y números."
        return None
    
    
    def hay_usuarios() -> bool:
        return db.uno("SELECT COUNT(*) FROM usuarios")[0] > 0
    
    
    def crear_usuario(username: str, password: str, nombre: str = "", rol: str = "lector", estado: str = "Pendiente"):
        username = username.strip()
        if not username:
            raise ValueError("El usuario no puede estar vacío.")
        if rol not in ROLES:
            raise ValueError("Rol inválido.")
        err = validar_password(password)
        if err:
            raise ValueError(err)
        if db.uno("SELECT id FROM usuarios WHERE LOWER(username) = LOWER(?)", (username,)):
            raise ValueError("Ese nombre de usuario ya existe.")
        db.ejecutar(
            "INSERT INTO usuarios (username, nombre, password_hash, rol, estado, creado_en) VALUES (?, ?, ?, ?, ?, ?)",
            (username, nombre.strip() or username, hash_password(password), rol, estado, db.ahora()),
        )
    
    
    def asegurar_admin_inicial():
        """Crea el administrador desde variables de entorno si aún no existe ninguno."""
        user = os.environ.get("ADMIN_USERNAME", "").strip()
        pwd = os.environ.get("ADMIN_PASSWORD", "")
        if not user or not pwd:
            return
        if db.uno("SELECT id FROM usuarios WHERE rol = 'admin'"):
            return
        if db.uno("SELECT id FROM usuarios WHERE LOWER(username) = LOWER(?)", (user,)):
            db.ejecutar("UPDATE usuarios SET rol = 'admin', estado = 'Aprobado' WHERE LOWER(username) = LOWER(?)", (user,))
            return
        crear_usuario(user, pwd, rol="admin", estado="Aprobado")
    
    
    def autenticar(username: str, password: str) -> tuple[dict | None, str | None]:
        """Devuelve (usuario, None) o (None, mensaje de error)."""
        fila = db.uno(
            "SELECT id, username, nombre, password_hash, rol, estado, intentos_fallidos, bloqueado_hasta "
            "FROM usuarios WHERE LOWER(username) = LOWER(?)",
            (username.strip(),),
        )
        generico = "Usuario o contraseña incorrectos."
        if not fila:
            hash_password(password)  # tiempo de respuesta similar exista o no el usuario
            return None, generico
        uid, uname, nombre, phash, rol, estado, intentos, bloqueado = fila
        if bloqueado and datetime.fromisoformat(bloqueado) > datetime.now():
            return None, f"Cuenta bloqueada temporalmente por intentos fallidos. Intenta de nuevo después de las {bloqueado[11:16]}."
        if not verificar_password(password, phash):
            intentos = (intentos or 0) + 1
            hasta = None
            if intentos >= MAX_INTENTOS:
                hasta = (datetime.now() + timedelta(minutes=MINUTOS_BLOQUEO)).isoformat(timespec="seconds")
                intentos = 0
            db.ejecutar("UPDATE usuarios SET intentos_fallidos = ?, bloqueado_hasta = ? WHERE id = ?", (intentos, hasta, uid))
            return None, generico
        db.ejecutar("UPDATE usuarios SET intentos_fallidos = 0, bloqueado_hasta = NULL WHERE id = ?", (uid,))
        if estado != "Aprobado":
            return None, "Tu cuenta está registrada pero pendiente de aprobación por un administrador."
        return {"id": uid, "username": uname, "nombre": nombre or uname, "rol": rol}, None
    
    
    def cambiar_password(uid: int, nueva: str, conservar_sesion: str | None = None):
        err = validar_password(nueva)
        if err:
            raise ValueError(err)
        db.ejecutar("UPDATE usuarios SET password_hash = ? WHERE id = ?", (hash_password(nueva), uid))
        cerrar_sesiones_usuario(uid, excepto=conservar_sesion)
    
    
    def listar():
        return db.consultar("SELECT id, username, nombre, rol, estado, creado_en FROM usuarios ORDER BY username")
    
    
    def n_admins_activos() -> int:
        return db.uno("SELECT COUNT(*) FROM usuarios WHERE rol = 'admin' AND estado = 'Aprobado'")[0]
    
    
    def actualizar(uid: int, rol: str, estado: str):
        db.ejecutar("UPDATE usuarios SET rol = ?, estado = ? WHERE id = ?", (rol, estado, uid))
    
    
    def eliminar(uid: int):
        cerrar_sesiones_usuario(uid)
        db.ejecutar("DELETE FROM usuario_contratos WHERE usuario_id = ?", (uid,))
        db.ejecutar("DELETE FROM usuarios WHERE id = ?", (uid,))
    
    
    def puede(usuario: dict | None, accion: str) -> bool:
        """accion: 'ver', 'cargar' o 'administrar'."""
        if not usuario:
            return False
        rol = usuario.get("rol")
        if accion == "ver":
            return True
        if accion == "cargar":
            return rol in ("admin", "planificador")
        if accion == "administrar":
            return rol == "admin"
        return False
    
    
    # ---------------------------------------------------------------- sesiones persistentes
    # El navegador conserva un token aleatorio en la URL (?s=...). En la base solo se guarda su hash (SHA-256),
    # atado al navegador que inició sesión, con vencimiento que se renueva con el uso.
    
    HORAS_SESION = float(os.environ.get("SESION_HORAS", "12"))
    DIAS_RECORDAR = float(os.environ.get("SESION_DIAS_RECORDAR", "7"))
    
    
    def _hash_token(token: str) -> str:
        return hashlib.sha256(token.encode()).hexdigest()
    
    
    def _huella_navegador(navegador: str | None) -> str:
        return hashlib.sha256((navegador or "").encode()).hexdigest()[:32]
    
    
    def crear_sesion(uid: int, navegador: str | None, recordar: bool = False) -> str:
        token = secrets.token_urlsafe(32)
        duracion = timedelta(days=DIAS_RECORDAR) if recordar else timedelta(hours=HORAS_SESION)
        ahora_ = datetime.now()
        db.ejecutar("DELETE FROM sesiones WHERE expira_en < ?", (ahora_.isoformat(timespec="seconds"),))
        db.ejecutar(
            "INSERT INTO sesiones (token_hash, usuario_id, navegador, creada_en, expira_en) VALUES (?, ?, ?, ?, ?)",
            (_hash_token(token), uid, _huella_navegador(navegador) + ("|R" if recordar else ""),
             ahora_.isoformat(timespec="seconds"), (ahora_ + duracion).isoformat(timespec="seconds")),
        )
        return token
    
    
    def validar_sesion(token: str | None, navegador: str | None) -> dict | None:
        """Devuelve el usuario de un token vigente (y renueva su vencimiento) o None."""
        if not token or len(token) < 20:
            return None
        fila = db.uno(
            "SELECT s.usuario_id, s.navegador, s.expira_en, u.username, u.nombre, u.rol, u.estado "
            "FROM sesiones s JOIN usuarios u ON u.id = s.usuario_id WHERE s.token_hash = ?",
            (_hash_token(token),),
        )
        if not fila:
            return None
        uid, nav, expira, username, nombre, rol, estado = fila
        huella, _, recordar = (nav or "").partition("|")
        if datetime.fromisoformat(expira) < datetime.now() or estado != "Aprobado" or huella != _huella_navegador(navegador):
            cerrar_sesion(token)
            return None
        duracion = timedelta(days=DIAS_RECORDAR) if recordar == "R" else timedelta(hours=HORAS_SESION)
        db.ejecutar("UPDATE sesiones SET expira_en = ? WHERE token_hash = ?",
                    ((datetime.now() + duracion).isoformat(timespec="seconds"), _hash_token(token)))
        return {"id": uid, "username": username, "nombre": nombre or username, "rol": rol}
    
    
    def cerrar_sesion(token: str | None):
        if token:
            db.ejecutar("DELETE FROM sesiones WHERE token_hash = ?", (_hash_token(token),))
    
    
    def cerrar_sesiones_usuario(uid: int, excepto: str | None = None):
        if excepto:
            db.ejecutar("DELETE FROM sesiones WHERE usuario_id = ? AND token_hash <> ?", (uid, _hash_token(excepto)))
        else:
            db.ejecutar("DELETE FROM sesiones WHERE usuario_id = ?", (uid,))
    
    
    # ---------------------------------------------------------------- firma para informes
    
    def guardar_firma(uid: int, png: bytes | None, cargo: str | None = None):
        """Cada usuario guarda solo su propia firma (la vista nunca permite cargar la de otro)."""
        import base64 as _b64
        if png is None:
            db.ejecutar("UPDATE usuarios SET firma = NULL, firma_actualizada = ? WHERE id = ?", (db.ahora(), int(uid)))
        else:
            db.ejecutar("UPDATE usuarios SET firma = ?, firma_actualizada = ? WHERE id = ?",
                        (_b64.b64encode(png).decode(), db.ahora(), int(uid)))
        if cargo is not None:
            db.ejecutar("UPDATE usuarios SET cargo = ? WHERE id = ?", (cargo.strip(), int(uid)))
    
    
    def guardar_cargo(uid: int, cargo: str):
        db.ejecutar("UPDATE usuarios SET cargo = ? WHERE id = ?", (cargo.strip(), int(uid)))
    
    
    def firma(uid: int) -> dict:
        """{"png": bytes|None, "cargo": str, "nombre": str, "actualizada": str}"""
        import base64 as _b64
        fila = db.uno("SELECT nombre, username, cargo, firma, firma_actualizada FROM usuarios WHERE id = ?", (int(uid),))
        if not fila:
            return {"png": None, "cargo": "", "nombre": "", "actualizada": None}
        nombre, username, cargo, f, act = fila
        return {"png": _b64.b64decode(f) if f else None, "cargo": cargo or "", "nombre": nombre or username, "actualizada": act}

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


auth = _modulo_auth()

# ======================================================================
# MÓDULO: analisis
# ======================================================================
def _modulo_analisis():
    import io
    import re
    import unicodedata
    
    import pandas as pd
    
    SIN_CODIGO = xer.SIN_CODIGO
    
    
    def pct(actual: float, at_completion: float) -> float:
        return actual / at_completion * 100 if at_completion and at_completion > 0 else 0.0
    
    
    def resumen_por_codigo(act: pd.DataFrame, cods: pd.DataFrame) -> pd.DataFrame:
        """Agrupa actividades por el valor de un código (Subsistema, Sistema, Área...).
    
        Labor Units % Complete = Actual / (Actual + Remaining), igual que P6.
        """
        df = act.merge(cods.drop_duplicates("task_code"), on="task_code", how="left")
        df["valor"] = df["valor"].fillna(SIN_CODIGO)
        g = df.groupby("valor", dropna=False).agg(
            actividades=("task_code", "count"),
            completadas=("estado", lambda s: int((s == "Completado").sum())),
            budget=("budget", "sum"),
            actual=("actual", "sum"),
            remaining=("remaining", "sum"),
            inicio=("inicio", "min"),
            fin=("fin", "max"),
        ).reset_index()
        g["at_completion"] = g["actual"] + g["remaining"]
        g["avance_pct"] = [pct(a, t) for a, t in zip(g["actual"], g["at_completion"])]
        return g.sort_values("valor").reset_index(drop=True)
    
    
    # ------------------------------------------------------------------ Excel
    
    def _norm(s) -> str:
        s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
        return re.sub(r"\s+", " ", s).strip()
    
    
    def detectar_fila_encabezado(raw: pd.DataFrame, max_filas: int = 30) -> int:
        """Fila (0-based) con más celdas de texto: suele ser el encabezado en planillas corporativas."""
        mejor, puntaje = 0, -1
        for i in range(min(max_filas, len(raw))):
            fila = raw.iloc[i]
            textos = sum(1 for v in fila if isinstance(v, str) and v.strip() and not _es_numero(v))
            if textos > puntaje:
                mejor, puntaje = i, textos
        return mejor
    
    
    def _es_numero(v) -> bool:
        try:
            float(str(v).replace(",", "."))
            return True
        except ValueError:
            return False
    
    
    # ---------- lectura directa del .xlsx (celdas combinadas y fórmulas), sin cargar todo el libro
    
    def _col_a_num(letras: str) -> int:
        n = 0
        for ch in letras:
            n = n * 26 + ord(ch) - 64
        return n
    
    
    def _xml_hoja(archivo_bytes: bytes, hoja: str) -> str | None:
        import zipfile
        from xml.sax.saxutils import escape as esc
    
        try:
            z = zipfile.ZipFile(io.BytesIO(archivo_bytes))
            wb = z.read("xl/workbook.xml").decode("utf-8", "ignore")
            rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
        except (zipfile.BadZipFile, KeyError):
            return None  # .xls antiguo u otro formato
        for nombre_xml in {esc(hoja, {'"': "&quot;"}), hoja}:
            m = re.search(r'<sheet\b[^>]*\bname="' + re.escape(nombre_xml) + r'"[^>]*\br:id="([^"]+)"', wb) or \
                re.search(r'<sheet\b[^>]*\br:id="([^"]+)"[^>]*\bname="' + re.escape(nombre_xml) + '"', wb)
            if m:
                break
        else:
            return None
        rid = m.group(1)
        t = re.search(r'<Relationship\b[^>]*\bId="' + re.escape(rid) + r'"[^>]*\bTarget="([^"]+)"', rels) or \
            re.search(r'<Relationship\b[^>]*\bTarget="([^"]+)"[^>]*\bId="' + re.escape(rid) + '"', rels)
        if not t:
            return None
        destino = t.group(1).lstrip("/")
        destino = destino if destino.startswith("xl/") else "xl/" + destino
        try:
            return z.read(destino).decode("utf-8", "ignore")
        except KeyError:
            return None
    
    
    def celdas_combinadas(xml: str | None) -> list[tuple[int, int, int, int]]:
        """[(fila1, col1, fila2, col2)] 1-based."""
        if not xml:
            return []
        salida = []
        for ref in re.findall(r'<mergeCell\b[^>]*\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', xml):
            c1, r1, c2, r2 = ref
            salida.append((int(r1), _col_a_num(c1), int(r2), _col_a_num(c2)))
        return salida
    
    
    def formulas_columna(xml: str | None, col: int) -> dict[int, str]:
        """{fila: fórmula} de una columna (1-based)."""
        if not xml:
            return {}
        from openpyxl.utils import get_column_letter
    
        letra = get_column_letter(col)
        salida = {}
        patron = r'<c\b[^>]*\br="' + letra + r'(\d+)"[^>]*>(?:(?!</c>).)*?<f\b[^>]*>([^<]*)</f>'
        for fila, f in re.findall(patron, xml, flags=re.S):
            salida[int(fila)] = f
        return salida
    
    
    def es_formula_suma(f: str, letra: str) -> bool:
        """SUM(...) o suma de varias celdas de la misma columna: fila de subtotal en el Excel."""
        f = f.upper()
        if "SUM(" in f and "SUMPRODUCT" not in f and "SUMIF" not in f:
            return True
        refs = re.findall(r"(?<![A-Z])\$?" + letra + r"\$?\d+", f)
        return len(refs) >= 2 and re.fullmatch(r"[=+\-\s$A-Z0-9()]*", f.replace(" ", "")) is not None
    
    
    def leer_hoja(archivo_bytes: bytes, hoja: str, fila_encabezado: int | None = None) -> tuple[pd.DataFrame, int]:
        """Lee una hoja con encabezados de una o varias filas (celdas combinadas).
    
        El índice del DataFrame es el número de fila de Excel (1-based) y df.attrs["columnas_excel"]
        mapea cada nombre de columna a su número de columna en Excel.
        """
        raw = pd.read_excel(io.BytesIO(archivo_bytes), sheet_name=hoja, header=None)
        xml = _xml_hoja(archivo_bytes, hoja)
        merges = celdas_combinadas(xml)
        if fila_encabezado is None:
            fila_encabezado = detectar_fila_encabezado(raw)
        h = fila_encabezado
    
        def es_texto(v):
            return isinstance(v, str) and v.strip() != "" and not _es_numero(v)
    
        def tiene_numeros(i):
            return any(isinstance(v, (int, float)) and not pd.isna(v) and not isinstance(v, bool) for v in raw.iloc[i])
    
        def bajo_combinada_o_vacia(i_sub, j):
            """La celda (i_sub, j) completa el encabezado: arriba hay una celda combinada o vacía."""
            fx, cx = i_sub + 1, j + 1
            for r1, c1, r2, c2 in merges:
                if c1 <= cx <= c2 and r1 <= fx - 1 <= r2 and (c2 > c1 or r2 >= fx):
                    return True
            v = raw.iat[i_sub - 1, j]
            return not es_texto(v) and (v is None or (isinstance(v, float) and pd.isna(v)))
    
        # filas de encabezado debajo (máximo 2): solo texto, sin números, y bajo celdas combinadas o vacías
        fin = h
        while fin + 1 < len(raw) and fin - h < 2 and not tiene_numeros(fin + 1):
            textos_j = [j for j, v in enumerate(raw.iloc[fin + 1]) if es_texto(v)]
            if not textos_j or not all(bajo_combinada_o_vacia(fin + 1, j) for j in textos_j):
                break
            fin += 1
        # fila de grupo arriba (PO, POM 2, POM 5...): textos cortos dentro de celdas combinadas horizontales
        ini = h
        if h > 0 and merges and not tiene_numeros(h - 1):
            arriba = h  # fila Excel (1-based) de la fila superior
            en_merge = [m for m in merges if m[0] == arriba and m[3] > m[1]]
            textos = [v for v in raw.iloc[h - 1] if es_texto(v)]
            if en_merge and textos and all(len(str(t)) <= 30 for t in textos):
                ini = h - 1
    
        def valor(i, j):
            """Valor de la celda (fila i, columna j, 0-based) considerando celdas combinadas."""
            fila_x, col_x = i + 1, j + 1
            for r1, c1, r2, c2 in merges:
                if r1 <= fila_x <= r2 and c1 <= col_x <= c2:
                    i, j = r1 - 1, c1 - 1
                    break
            if i < len(raw) and j < raw.shape[1]:
                return raw.iat[i, j]
            return None
    
        nombres, vistos = [], {}
        for j in range(raw.shape[1]):
            partes = []
            for i in range(ini, fin + 1):
                v = valor(i, j)
                if es_texto(v):
                    t = re.sub(r"\s+", " ", str(v)).strip()
                    if not partes or partes[-1] != t:
                        partes.append(t)
            n = " · ".join(partes) if partes else f"Columna {j + 1}"
            if n in vistos:
                vistos[n] += 1
                n = f"{n} ({vistos[n]})"
            else:
                vistos[n] = 0
            nombres.append(n)
        df = raw.iloc[fin + 1:].copy()
        df.columns = nombres
        df.index = [i + 1 for i in range(fin + 1, len(raw))]  # número de fila en Excel
        df = df.dropna(how="all").dropna(axis=1, how="all")
        df.attrs["columnas_excel"] = {n: j + 1 for j, n in enumerate(nombres) if n in df.columns}
        df.attrs["filas_encabezado"] = (ini + 1, fin + 1)
        return df, fila_encabezado
    
    
    def filas_suma_excel(archivo_bytes: bytes, hoja: str, columnas_excel) -> set[int] | None:
        """Filas cuya celda, en alguna de las columnas indicadas, es una fórmula de suma. None si no se puede saber."""
        if isinstance(columnas_excel, int):
            columnas_excel = [columnas_excel]
        columnas_excel = [c for c in (columnas_excel or []) if c]
        if not columnas_excel:
            return None
        xml = _xml_hoja(archivo_bytes, hoja)
        if not xml:
            return None
        from openpyxl.utils import get_column_letter
    
        filas: set[int] = set()
        for col in columnas_excel:
            letra = get_column_letter(col)
            filas |= {fila for fila, f in formulas_columna(xml, col).items() if es_formula_suma(f, letra)}
        return filas
    
    
    SUGERENCIAS = {
        "item": ["subsistema", "itemizado", "item", "partida", "codigo"],
        "descripcion": ["descripcion", "nombre", "alcance"],
        "at_completion": ["at completion", "total hh", "hh totales", "hh total", "hh ppto", "presupuesto", "budget"],
        "remaining": ["remaining", "saldo", "restante", "por ejecutar", "pendiente"],
        "avance": ["% avance", "avance", "% complete", "complete", "cumplimiento", "%"],
    }
    _EXCLUIR_PERIODO = ("anterior", "periodo", "semana")
    
    
    def sugerir_columna(columnas: list[str], clave: str, excluir: set[str] = frozenset()) -> str | None:
        cols = [(c, _norm(c)) for c in columnas if c not in excluir]
        if clave in ("actual", "gastadas"):
            raiz = r"\bgan|actual|ejecutad" if clave == "actual" else r"\bgas\b|gastad"
            mejor, puntaje = None, -1
            for c, n in cols:
                if not re.search(raiz, n) or "%" in n or any(k in n for k in _EXCLUIR_PERIODO) or "clp" in n:
                    continue
                p = (2 if "acumul" in n else 0) + (1 if "hh" in n else 0)
                if p > puntaje:
                    mejor, puntaje = c, p
            return mejor
        for palabra in SUGERENCIAS[clave]:
            coinciden = [c for c, n in cols if palabra in n and "clp" not in n and "precio" not in n]
            if clave == "avance":
                coinciden = [c for c in coinciden if not any(k in _norm(c) for k in _EXCLUIR_PERIODO)] or coinciden
                acum = [c for c in coinciden if "acumul" in _norm(c)]
                coinciden = acum or coinciden
            if coinciden:
                # con varias revisiones del presupuesto (PO, POM 2... POM 5) la vigente es la última
                return coinciden[-1] if clave in ("at_completion", "avance") else coinciden[0]
        return None
    
    
    def a_numero(serie: pd.Series) -> pd.Series:
        """Acepta 1.234,5 / 1,234.5 / 64% / 0,64."""
        def conv(v):
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return float("nan")
            if isinstance(v, (int, float)):
                return float(v)
            s = str(v).strip().replace("%", "").replace(" ", "")
            if not s or s in ("-", "--"):
                return float("nan")
            if "," in s and "." in s:
                s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
            elif "," in s:
                s = s.replace(",", ".")  # en Chile la coma es decimal: "0,64"
            elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):
                s = s.replace(".", "")  # "1.200" o "12.345.678" son miles
            try:
                return float(s)
            except ValueError:
                return float("nan")
        return serie.map(conv)
    
    
    def escala_porcentaje(serie: pd.Series) -> float:
        """1 si los valores vienen como 0-100, 100 si vienen como fracción 0-1."""
        v = serie.dropna()
        v = v[v != 0]
        if v.empty:
            return 1.0
        return 100.0 if v.abs().max() <= 1.5 else 1.0
    
    
    def detectar_subtotales(valores: list[float], tol: float = 0.5, max_iter: int = 6) -> dict[int, str]:
        """Posiciones cuyo valor es la suma de un bloque contiguo de filas (títulos de grupo o subtotales).
    
        Se aplica por capas: primero los grupos internos y luego los externos, para planillas con varios niveles.
        Devuelve {posición: motivo}.
        """
        marcadas: dict[int, str] = {}
        for _ in range(max_iter):
            vivas = [i for i in range(len(valores)) if i not in marcadas]
            nuevas = {}
            for k, i in enumerate(vivas):
                v = valores[i]
                if not v or v <= tol:
                    continue
                # fila de grupo arriba de sus hijas
                acum, n = 0.0, 0
                for j in vivas[k + 1:]:
                    if j in nuevas:
                        break
                    acum += valores[j]
                    n += 1
                    if abs(acum - v) <= tol and n >= 2:
                        nuevas[i] = "Suma de las filas siguientes (título de grupo)"
                        break
                    if acum > v + tol:
                        break
                if i in nuevas:
                    continue
                # subtotal debajo de sus filas
                acum, n = 0.0, 0
                for j in reversed(vivas[:k]):
                    if j in nuevas:
                        break
                    acum += valores[j]
                    n += 1
                    if abs(acum - v) <= tol and n >= 2:
                        nuevas[i] = "Suma de las filas anteriores (subtotal)"
                        break
                    if acum > v + tol:
                        break
            if not nuevas:
                break
            marcadas.update(nuevas)
        return marcadas
    
    
    def normalizar_informe(df: pd.DataFrame, col_item: str, col_atc: str, col_rem: str | None = None,
                           col_actual: str | None = None, col_avance: str | None = None,
                           excluir_totales: bool = True, excluir_subtotales: bool = True,
                           col_desc: str | None = None, col_gastadas: str | None = None,
                           filas_suma: set[int] | None = None) -> tuple[pd.DataFrame, list[str]]:
        """Devuelve (items, advertencias).
    
        items: item, at_completion, remaining, actual, avance_pct, incluir, motivo, fila_origen.
        Las filas de total y las que son suma de otras quedan con incluir=False (se pueden reactivar en pantalla).
        """
        avisos: list[str] = []
        def texto(col):
            return df[col].map(lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip())
    
        item = texto(col_item)
        if col_desc and col_desc != col_item:
            desc = texto(col_desc)
            item = (item + " " + desc).str.strip()
        out = pd.DataFrame({"item": item, "at_completion": a_numero(df[col_atc])})
        out = out[(out["item"] != "") & (out["item"].str.lower() != "nan")]
        # Filas sin presupuesto pero con horas ganadas o gastadas (p. ej. trabajos nuevos) se conservan con At Completion 0.
        otras = [c for c in (col_rem, col_actual, col_gastadas) if c]
        con_horas = pd.Series(False, index=out.index)
        for col in otras:
            v = a_numero(df.loc[out.index, col])
            con_horas |= v.notna() & (v != 0)
        sin_num = out["at_completion"].isna() & ~con_horas
        if sin_num.any():
            avisos.append(f"Se ignoraron {int(sin_num.sum())} filas sin horas (títulos o filas vacías).")
        out = out[~sin_num].copy()
        out["at_completion"] = out["at_completion"].fillna(0)
        idx = out.index
    
        if col_rem:
            out["remaining"] = a_numero(df.loc[idx, col_rem]).fillna(0)
            out["actual"] = out["at_completion"] - out["remaining"]
        elif col_actual:
            out["actual"] = a_numero(df.loc[idx, col_actual]).fillna(0)
            out["remaining"] = out["at_completion"] - out["actual"]
        elif col_avance:
            p = a_numero(df.loc[idx, col_avance])
            factor = escala_porcentaje(p)
            if factor == 100:
                avisos.append("El avance venía como fracción (0–1); se convirtió a porcentaje.")
            p = (p * factor).fillna(0)
            out["actual"] = out["at_completion"] * p / 100
            out["remaining"] = out["at_completion"] - out["actual"]
        else:
            raise ValueError("Debes indicar la columna de Remaining, de Actual o de % Avance.")
    
        out["avance_pct"] = [pct(a, t) for a, t in zip(out["actual"], out["at_completion"])]
        out["hh_gastadas"] = a_numero(df.loc[idx, col_gastadas]) if col_gastadas else float("nan")
        out["fila_origen"] = out.index
        out = out.reset_index(drop=True)
        out["incluir"] = True
        out["motivo"] = ""
    
        if excluir_totales:
            es_total = out["item"].map(_norm).str.contains(r"(?:^|\b)(?:sub\s*-?\s*)?total(?:es)?\b|^resumen\b", regex=True)
            out.loc[es_total, ["incluir", "motivo"]] = [False, "Fila de total"]
            if es_total.any():
                avisos.append(f"Se excluyeron {int(es_total.sum())} filas de total.")
        if excluir_subtotales and filas_suma:
            # El Excel indica sus subtotales con fórmulas SUM(...): es la señal más confiable.
            es_suma = out["incluir"] & out["fila_origen"].isin(filas_suma)
            out.loc[es_suma, ["incluir", "motivo"]] = [False, "Fórmula de suma en el Excel (subtotal)"]
            if es_suma.any():
                avisos.append(f"Se excluyeron {int(es_suma.sum())} filas de subtotal (fórmulas SUM del Excel) "
                              "para no contar las horas dos veces.")
        elif excluir_subtotales:
            vivas = out.index[out["incluir"]].tolist()
            marcas = detectar_subtotales(out.loc[vivas, "at_completion"].tolist())
            for pos, motivo in marcas.items():
                out.loc[vivas[pos], ["incluir", "motivo"]] = [False, motivo]
            if marcas:
                avisos.append(f"Se excluyeron {len(marcas)} filas que son la suma de otras (títulos de grupo o subtotales), "
                              "para no contar las horas dos veces. Revísalas en la tabla.")
    
        incl = out[out["incluir"]]
        if (incl["remaining"] < -0.5).any():
            sin_ppto = int(((incl["remaining"] < -0.5) & (incl["at_completion"] == 0)).sum())
            otros = int((incl["remaining"] < -0.5).sum()) - sin_ppto
            if sin_ppto:
                avisos.append(f"{sin_ppto} ítems tienen HH ganadas pero At Completion 0 (sin presupuesto en esa columna). "
                              "Revisa la planilla.")
            if otros:
                avisos.append(f"{otros} ítems tienen HH ganadas mayores que su At Completion. Revisa la planilla.")
        return out, avisos
    
    
    def productividad(items: pd.DataFrame) -> float | None:
        """PF = HH gastadas / HH ganadas (sobre 1 = se gastan más horas de las ganadas)."""
        if "incluir" in items:
            items = items[items["incluir"].astype(bool)]
        if "hh_gastadas" not in items or items["hh_gastadas"].isna().all():
            return None
        ganadas = items["actual"].sum() if "actual" in items else (items["at_completion"] - items["remaining"]).sum()
        return items["hh_gastadas"].sum() / ganadas if ganadas > 0 else None
    
    
    def avance_ponderado(items: pd.DataFrame) -> float:
        if "incluir" in items:
            items = items[items["incluir"].astype(bool)]
        return pct(items["actual"].sum() if "actual" in items else (items["at_completion"] - items["remaining"]).sum(),
                   items["at_completion"].sum())
    
    
    def a_excel(hojas: dict[str, pd.DataFrame]) -> bytes:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            for nombre, df in hojas.items():
                df.to_excel(xw, sheet_name=nombre[:31], index=False)
                ws = xw.sheets[nombre[:31]]
                for i, col in enumerate(df.columns, start=1):
                    largo = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)])
                    ws.column_dimensions[ws.cell(1, i).column_letter].width = min(60, largo + 2)
        return buf.getvalue()
    
    
    # ------------------------------------------------------- Informe ejecutivo por ítem
    
    def _clave_codigo(valor: str) -> str:
        """'SS-2963-01-01-003: Boiler 001' -> 'SS-2963-01-01-003'."""
        return str(valor).split(":")[0].strip()
    
    
    def _filtrar(act: pd.DataFrame, cods: pd.DataFrame, valor: str) -> pd.DataFrame:
        if valor == SIN_CODIGO:
            return act[~act["task_code"].isin(set(cods["task_code"]))]
        return act[act["task_code"].isin(set(cods.loc[cods["valor"] == valor, "task_code"]))]
    
    
    def _fmt(x, dec=0):
        s = f"{x:,.{dec}f}"
        return s.replace(",", "§").replace(".", ",").replace("§", ".")
    
    
    def _fecha_txt(s) -> str:
        if s is None or (isinstance(s, float) and pd.isna(s)) or not str(s).strip():
            return "–"
        try:
            return pd.Timestamp(str(s)[:10]).strftime("%d-%m-%Y")
        except ValueError:
            return str(s)
    
    
    def datos_informe_item(act: pd.DataFrame, cods: pd.DataFrame, valor: str, tipo: str, proyecto: dict, corte: dict,
                           act_prev: pd.DataFrame | None = None, cods_prev: pd.DataFrame | None = None,
                           fecha_prev: str | None = None, cods_esp: pd.DataFrame | None = None, tipo_esp: str | None = None,
                           restricciones: pd.DataFrame | None = None, horas_dia: float = 8.0) -> dict:
        """Reúne todo lo que necesita el informe ejecutivo de un subsistema (o cualquier valor de código)."""
        for df in (act, act_prev):
            if df is not None and "at_completion" not in df:
                df["at_completion"] = df["actual"] + df["remaining"]
        sub = _filtrar(act, cods, valor).copy()
        atc, actual, rem = sub["at_completion"].sum(), sub["actual"].sum(), sub["remaining"].sum()
        budget = sub["budget"].sum()
        avance = pct(actual, atc)
        atc_proy, actual_proy = act["at_completion"].sum(), act["actual"].sum()
        avance_proy = pct(actual_proy, atc_proy)
    
        resumen_items = resumen_por_codigo(act, cods)
        resumen_items = resumen_items.sort_values("at_completion", ascending=False).reset_index(drop=True)
        con_horas = resumen_items[resumen_items["at_completion"] > 0].reset_index(drop=True)
        ranking = next((i + 1 for i, v in enumerate(con_horas["valor"]) if v == valor), None)
    
        estados = sub["estado"].value_counts()
        d = {
            "valor": valor, "tipo": tipo, "proyecto_codigo": proyecto.get("codigo", ""),
            "proyecto_nombre": proyecto.get("nombre", ""), "fecha_corte": corte.get("fecha_corte"),
            "actividades": int(len(sub)), "completadas": int(estados.get("Completado", 0)),
            "en_progreso": int(estados.get("En Progreso", 0)), "no_iniciadas": int(estados.get("No Iniciado", 0)),
            "budget": budget, "at_completion": atc, "actual": actual, "remaining": rem, "avance_pct": avance,
            "avance_proyecto": avance_proy, "participacion_pct": pct(atc, atc_proy),
            "ranking": ranking, "n_items_con_horas": int(len(con_horas)),
            "inicio": sub["inicio"].dropna().min() if sub["inicio"].notna().any() else None,
            "fin": sub["fin"].dropna().max() if sub["fin"].notna().any() else None,
            "fecha_prev": fecha_prev, "avance_prev": None, "delta_pp": None, "hh_ganadas": None,
        }
    
        if act_prev is not None and cods_prev is not None and not act_prev.empty:
            prev = _filtrar(act_prev, cods_prev, valor)
            if not prev.empty:
                d["avance_prev"] = pct(prev["actual"].sum(), prev["at_completion"].sum())
                d["delta_pp"] = avance - d["avance_prev"]
                d["hh_ganadas"] = actual - prev["actual"].sum()
    
        # Desglose por especialidad / disciplina dentro del subsistema
        d["tipo_esp"] = tipo_esp
        d["por_especialidad"] = None
        if cods_esp is not None and tipo_esp:
            e = resumen_por_codigo(sub, cods_esp)
            d["por_especialidad"] = e[e["actividades"] > 0].sort_values("at_completion", ascending=False).reset_index(drop=True)
    
        pend = sub[sub["estado"] != "Completado"]
        d["lista_en_progreso"] = sub[sub["estado"] == "En Progreso"].sort_values("remaining", ascending=False).head(15)
        d["lista_proximas"] = sub[sub["estado"] == "No Iniciado"].sort_values(["inicio", "task_code"]).head(12)
        crit = pend[pend["holgura_total_hr"].notna()].copy()
        crit["holgura_dias"] = crit["holgura_total_hr"] / horas_dia
        d["lista_criticas"] = crit.sort_values("holgura_dias").head(10)
        d["n_criticas"] = int((crit["holgura_dias"] <= 0).sum())
    
        d["restricciones"] = None
        d["n_restr_pend"] = d["n_restr_vencidas"] = 0
        if restricciones is not None and not restricciones.empty:
            clave = _clave_codigo(valor).lower()
            r = restricciones[restricciones["subsistema"].fillna("").astype(str).str.strip().str.lower().map(
                lambda s: bool(s) and (s == valor.lower() or _clave_codigo(s) == clave))]
            if proyecto.get("codigo"):
                r = r[r["proyecto"].astype(str) == str(proyecto["codigo"])] if (r["proyecto"].astype(str) == str(proyecto["codigo"])).any() else r
            liberadas = {"liberada", "liberado", "resuelto", "resuelta", "completado", "completada", "cerrada", "cerrado"}
            r = r[~r["estado"].fillna("").str.strip().str.lower().isin(liberadas)]
            hoy = pd.Timestamp.today().strftime("%Y-%m-%d")
            d["restricciones"] = r
            d["n_restr_pend"] = int(len(r))
            d["n_restr_vencidas"] = int((r["fecha_compromiso"].notna() & (r["fecha_compromiso"].astype(str).str[:10] < hoy)).sum())
    
        d["resumen"] = _redactar_resumen(d)
        return d
    
    
    def _redactar_resumen(d: dict) -> list[str]:
        """Resumen ejecutivo en lenguaje natural, construido solo con los datos calculados."""
        s = []
        nombre = d["valor"]
        if d["at_completion"] <= 0:
            s.append(f"{nombre} no tiene horas hombre (Labor Units) asignadas en el programa; el avance se informa por estado de actividades.")
        elif d["avance_pct"] >= 99.995:
            s.append(f"{nombre} se encuentra completado: {_fmt(d['actual'])} HH ejecutadas de {_fmt(d['at_completion'])} HH.")
        else:
            dif = d["avance_pct"] - d["avance_proyecto"]
            comp = ("por sobre el" if dif > 0.5 else "por debajo del" if dif < -0.5 else "en línea con el")
            s.append(f"{nombre} registra un avance de {_fmt(d['avance_pct'], 2)}% en Labor Units "
                     f"({_fmt(d['actual'])} de {_fmt(d['at_completion'])} HH), {comp} avance global del proyecto "
                     f"({_fmt(d['avance_proyecto'], 2)}%" + (f"; diferencia de {_fmt(abs(dif), 1)} pp)." if dif and abs(dif) > 0.5 else ")."))
        if d["ranking"]:
            s.append(f"Representa el {_fmt(d['participacion_pct'], 1)}% de las horas del proyecto y es el n° {d['ranking']} "
                     f"de {d['n_items_con_horas']} en volumen de HH.")
        if d["delta_pp"] is not None:
            if abs(d["delta_pp"]) < 0.005:
                s.append(f"No registra avance respecto del corte anterior ({_fecha_txt(d['fecha_prev'])}).")
            else:
                s.append(f"Respecto del corte anterior ({_fecha_txt(d['fecha_prev'])}) avanzó {_fmt(d['delta_pp'], 2)} pp, "
                         f"equivalentes a {_fmt(d['hh_ganadas'])} HH ganadas en el período.")
        s.append(f"De sus {d['actividades']} actividades, {d['completadas']} están completadas, {d['en_progreso']} en progreso "
                 f"y {d['no_iniciadas']} sin iniciar.")
        if d["remaining"] > 0:
            s.append(f"Quedan {_fmt(d['remaining'])} HH por ejecutar; la última actividad termina según programa el {_fecha_txt(d['fin'])}.")
        if d["n_criticas"]:
            s.append(f"{d['n_criticas']} actividad(es) pendiente(s) tienen holgura total nula o negativa y condicionan el término.")
        if d["budget"] > 0 and abs(d["at_completion"] - d["budget"]) / d["budget"] > 0.05:
            dif = d["at_completion"] - d["budget"]
            s.append(f"Las HH At Completion {'superan' if dif > 0 else 'están bajo'} el Budget en {_fmt(abs(dif))} HH "
                     f"({_fmt(abs(pct(dif, d['budget'])), 1)}%).")
        if d["n_restr_pend"]:
            s.append(f"Tiene {d['n_restr_pend']} restricción(es) 6WLA pendiente(s)"
                     + (f", {d['n_restr_vencidas']} con fecha de compromiso vencida." if d["n_restr_vencidas"] else "."))
        return s
    
    
    # ------------------------------------------------------- Resumen del informe semanal (avance actual)
    
    def _norm_txt(v) -> str:
        import unicodedata
        s = unicodedata.normalize("NFKD", str(v if v is not None else "")).encode("ascii", "ignore").decode().upper()
        return re.sub(r"\s+", " ", s).strip()
    
    
    def leer_resumen_semanal(datos: bytes) -> dict | None:
        """Lee la tabla «Resumen» del informe semanal: HH ganadas, HH gastadas y PF del acumulado anterior, del período
        actual y del acumulado actual, por fila (Avance Construcción PO, Aumento Forecast + adicionales, Total…).
        También el N° del informe y las fechas del período. Devuelve None si el archivo no trae esa tabla."""
        import openpyxl
        from datetime import date, datetime
        wb = openpyxl.load_workbook(io.BytesIO(datos), read_only=True, data_only=True)
        try:
            hojas = sorted(wb.worksheets, key=lambda w: 0 if "RESUMEN" in _norm_txt(w.title) else 1)
            numero, per_desde, per_hasta = None, None, None
            for ws in wb.worksheets[:3]:
                filas = list(ws.iter_rows(max_row=40, max_col=30, values_only=True))
                for i, fila in enumerate(filas):
                    for j, v in enumerate(fila):
                        t = _norm_txt(v)
                        m = re.search(r"INFORME SEMANAL N\s*[°ºO]?\s*(\d+)", t)
                        if m and numero is None:
                            numero = int(m.group(1))
                        if t.startswith("PERIODO DEL") and per_desde is None:
                            resto = [x for x in fila[j + 1:] if isinstance(x, (datetime, date))]
                            per_desde = resto[0] if resto else None
                            sig = filas[i + 1] if i + 1 < len(filas) else ()
                            fechas = [x for x in sig if isinstance(x, (datetime, date))]
                            per_hasta = fechas[0] if fechas else None
            a_fecha = lambda x: x.date() if isinstance(x, datetime) else x
            for ws in hojas:
                filas = list(ws.iter_rows(max_row=200, max_col=40, values_only=True))
                for i, fila in enumerate(filas):
                    nombres = [_norm_txt(v) for v in fila]
                    gan = [j for j, n in enumerate(nombres) if n == "HH GANADAS"]
                    gas = [j for j, n in enumerate(nombres) if n == "HH GASTADAS"]
                    if len(gan) < 2 or len(gas) < 2:
                        continue
                    # títulos de grupo y fechas en las filas de arriba
                    grupos = []
                    for j in gan:
                        titulo, fecha = "", None
                        for jj in (j, j + 1, j - 1):           # el título del grupo suele estar en la columna de «HH Ganadas»
                            for k in range(i - 1, max(-1, i - 5), -1):
                                v = filas[k][jj] if 0 <= jj < len(filas[k]) else None
                                if isinstance(v, str) and v.strip() and not titulo:
                                    titulo = v.strip()
                                if isinstance(v, (datetime, date)) and fecha is None:
                                    fecha = a_fecha(v)
                            if titulo:
                                break
                        grupos.append((j, titulo, fecha))
                    def clave(t):
                        t = _norm_txt(t)
                        if "ANTERIOR" in t:
                            return "anterior"
                        if "ACUMULADO" in t:
                            return "acumulado"
                        return "periodo"
                    col_desc = next((j for j in range(gan[0] - 1, -1, -1) if nombres[j]), max(0, gan[0] - 1))
                    res_filas = []
                    for fila2 in filas[i + 1:i + 30]:
                        desc = fila2[col_desc] if col_desc < len(fila2) else None
                        if desc is None or not str(desc).strip():
                            if res_filas and all(v is None for v in fila2[gan[0]:gan[-1] + 3]):
                                continue
                            continue
                        if isinstance(desc, (int, float)):
                            break
                        vals = {}
                        for j, titulo, _ in grupos:
                            g, s = fila2[j] if j < len(fila2) else None, fila2[j + 1] if j + 1 < len(fila2) else None
                            if isinstance(g, (int, float)) or isinstance(s, (int, float)):
                                g, s = float(g or 0), float(s or 0)
                                vals[clave(titulo)] = {"ganadas": g, "gastadas": s, "pf": s / g if g else None}
                        if not vals:
                            if res_filas:
                                break
                            continue
                        res_filas.append({"descripcion": str(desc).strip(), **vals})
                    if not res_filas:
                        continue
                    fechas = {clave(t): f for _, t, f in grupos}
                    total = next((f for f in res_filas if "TOTAL" in _norm_txt(f["descripcion"])), res_filas[-1])
                    return {"hoja": ws.title.strip(), "numero": numero, "desde": a_fecha(per_desde), "hasta": a_fecha(per_hasta) or fechas.get("acumulado"),
                            "fecha_anterior": fechas.get("anterior"), "fecha_actual": fechas.get("acumulado"),
                            "filas": res_filas, "total": total["descripcion"]}
            return None
        finally:
            wb.close()

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


analisis = _modulo_analisis()

# ======================================================================
# MÓDULO: informe_pdf
# ======================================================================
def _modulo_informe_pdf():
    import io
    import os
    from datetime import datetime
    
    import pandas as pd
    from reportlab.graphics.shapes import Drawing, Rect, String
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table,
                                    TableStyle)
    from xml.sax.saxutils import escape
    
    CAL = calendario
    
    NARANJO = colors.HexColor("#FF6B00")
    AZUL = colors.HexColor("#0F172A")
    GRIS = colors.HexColor("#64748B")
    GRIS_CLARO = colors.HexColor("#E2E8F0")
    FONDO = colors.HexColor("#F8FAFC")
    VERDE = colors.HexColor("#10B981")
    ROJO = colors.HexColor("#DC2626")
    
    ANCHO, ALTO = A4
    MARGEN = 16 * mm
    UTIL = ANCHO - 2 * MARGEN
    
    E = {
        "titulo": ParagraphStyle("titulo", fontName="Helvetica-Bold", fontSize=17, leading=21, textColor=AZUL),
        "subtitulo": ParagraphStyle("subtitulo", fontName="Helvetica", fontSize=9.5, leading=13, textColor=GRIS),
        "h2": ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=11.5, leading=15, textColor=AZUL, spaceBefore=10, spaceAfter=5),
        "cuerpo": ParagraphStyle("cuerpo", fontName="Helvetica", fontSize=9.2, leading=13, textColor=AZUL),
        "vineta": ParagraphStyle("vineta", fontName="Helvetica", fontSize=9.2, leading=13, textColor=AZUL, leftIndent=10,
                                 bulletIndent=0, spaceAfter=2.5),
        "celda": ParagraphStyle("celda", fontName="Helvetica", fontSize=7.6, leading=9.4, textColor=AZUL),
        "celda_der": ParagraphStyle("celda_der", fontName="Helvetica", fontSize=7.6, leading=9.4, textColor=AZUL, alignment=TA_RIGHT),
        "cab": ParagraphStyle("cab", fontName="Helvetica-Bold", fontSize=7.6, leading=9.4, textColor=colors.white),
        "cab_der": ParagraphStyle("cab_der", fontName="Helvetica-Bold", fontSize=7.6, leading=9.4, textColor=colors.white, alignment=TA_RIGHT),
        "kpi_valor": ParagraphStyle("kpi_valor", fontName="Helvetica-Bold", fontSize=14, leading=17, textColor=AZUL),
        "kpi_valor_big": ParagraphStyle("kpi_valor_big", fontName="Helvetica-Bold", fontSize=22, leading=25, textColor=NARANJO),
        "kpi_etq": ParagraphStyle("kpi_etq", fontName="Helvetica", fontSize=7.4, leading=9, textColor=GRIS),
        "vacio": ParagraphStyle("vacio", fontName="Helvetica-Oblique", fontSize=8.5, leading=11, textColor=GRIS),
    }
    
    
    def _n(x, dec=0) -> str:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return "–"
        s = f"{x:,.{dec}f}"
        return s.replace(",", "§").replace(".", ",").replace("§", ".")
    
    
    def _f(s) -> str:
        if s is None or (isinstance(s, float) and pd.isna(s)) or not str(s).strip():
            return "–"
        try:
            return pd.Timestamp(str(s)[:10]).strftime("%d-%m-%Y")
        except ValueError:
            return str(s)
    
    
    def _p(texto, estilo="celda") -> Paragraph:
        return Paragraph(escape("" if texto is None or (isinstance(texto, float) and pd.isna(texto)) else str(texto)), E[estilo])
    
    
    # ------------------------------------------------------------------ piezas
    
    def _kpi(etiqueta: str, valor: str, grande=False, ancho=None) -> Table:
        t = Table([[Paragraph(escape(valor), E["kpi_valor_big" if grande else "kpi_valor"])],
                   [Paragraph(escape(etiqueta), E["kpi_etq"])]], colWidths=[ancho])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.white),
            ("LINEABOVE", (0, 0), (-1, 0), 2.5, NARANJO),
            ("BOX", (0, 0), (-1, -1), 0.4, GRIS_CLARO),
            ("LEFTPADDING", (0, 0), (-1, -1), 7), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, 0), 6), ("BOTTOMPADDING", (0, -1), (-1, -1), 6),
        ]))
        return t
    
    
    def _fila_kpis(kpis: list[tuple[str, str, bool]]) -> Table:
        sep = 3 * mm
        ancho = (UTIL - sep * (len(kpis) - 1)) / len(kpis)
        celdas, anchos = [], []
        for i, (etq, val, grande) in enumerate(kpis):
            if i:
                celdas.append("")
                anchos.append(sep)
            celdas.append(_kpi(etq, val, grande, ancho))
            anchos.append(ancho)
        t = Table([celdas], colWidths=anchos)
        t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                               ("VALIGN", (0, 0), (-1, -1), "TOP")]))
        return t
    
    
    def _barra_comparativa(d: dict) -> Drawing:
        """Barras horizontales: avance del ítem vs avance del proyecto (y corte anterior si existe)."""
        corto = str(d["valor"]).split(":")[-1].strip() or str(d["valor"])
        filas = [(corto if len(corto) <= 26 else corto[:25] + "…", d["avance_pct"], NARANJO)]
        if d.get("avance_prev") is not None:
            filas.append((f"Corte anterior ({_f(d['fecha_prev'])})", d["avance_prev"], colors.HexColor("#FDBA74")))
        filas.append(("Proyecto completo", d["avance_proyecto"], AZUL))
        alto_fila, x0, ancho_barra = 15, 118, UTIL - 160
        dr = Drawing(UTIL, alto_fila * len(filas) + 6)
        for i, (etq, v, col) in enumerate(filas):
            y = dr.height - (i + 1) * alto_fila
            dr.add(String(0, y + 3.5, etq, fontName="Helvetica", fontSize=8, fillColor=GRIS))
            dr.add(Rect(x0, y + 1, ancho_barra, 9, fillColor=colors.HexColor("#EEF2F7"), strokeColor=None))
            dr.add(Rect(x0, y + 1, ancho_barra * max(0, min(v, 100)) / 100, 9, fillColor=col, strokeColor=None))
            dr.add(String(x0 + ancho_barra + 6, y + 2.5, f"{_n(v, 1)}%", fontName="Helvetica-Bold", fontSize=8.5, fillColor=AZUL))
        return dr
    
    
    def _grafico_especialidad(df: pd.DataFrame) -> Drawing:
        df = df.head(12)
        alto_fila, x0 = 17, 150
        ancho_barra = UTIL - x0 - 95
        max_hh = max(df["at_completion"].max(), 1)
        dr = Drawing(UTIL, alto_fila * len(df) + 16)
        dr.add(String(x0, dr.height - 9, "Barra = HH At Completion · tramo naranjo = HH ejecutadas", fontName="Helvetica",
                      fontSize=6.8, fillColor=GRIS))
        for i, r in enumerate(df.itertuples()):
            y = dr.height - 16 - (i + 1) * alto_fila
            etq = str(r.valor)
            etq = etq if len(etq) <= 34 else etq[:33] + "…"
            dr.add(String(0, y + 5, etq, fontName="Helvetica", fontSize=7.6, fillColor=AZUL))
            w = ancho_barra * r.at_completion / max_hh
            dr.add(Rect(x0, y + 2, max(w, 0.5), 10, fillColor=colors.HexColor("#CBD5E1"), strokeColor=None))
            dr.add(Rect(x0, y + 2, max(w * min(r.avance_pct, 100) / 100, 0), 10, fillColor=NARANJO, strokeColor=None))
            dr.add(String(x0 + w + 5, y + 4.5, f"{_n(r.avance_pct, 1)}%  ·  {_n(r.at_completion)} HH",
                          fontName="Helvetica", fontSize=7.2, fillColor=AZUL))
        return dr
    
    
    def _tabla(df: pd.DataFrame, columnas: list[tuple[str, str, float, str]], vacio: str):
        """columnas: (campo, encabezado, ancho relativo, 'izq'|'der'|'num'|'num1'|'pct'|'fecha')."""
        if df is None or df.empty:
            return Paragraph(escape(vacio), E["vacio"])
        total = sum(c[2] for c in columnas)
        anchos = [UTIL * c[2] / total for c in columnas]
        filas = [[Paragraph(escape(c[1]), E["cab"] if c[3] in ("izq", "fecha") else E["cab_der"]) for c in columnas]]
        for _, r in df.iterrows():
            fila = []
            for campo, _, _, fmt in columnas:
                v = r.get(campo)
                if fmt == "num":
                    fila.append(_p(_n(v), "celda_der"))
                elif fmt == "num1":
                    fila.append(_p(_n(v, 1), "celda_der"))
                elif fmt == "pct":
                    fila.append(_p(f"{_n(v, 1)}%", "celda_der"))
                elif fmt == "fecha":
                    fila.append(_p(_f(v)))
                else:
                    fila.append(_p(v))
            filas.append(fila)
        t = Table(filas, colWidths=anchos, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), AZUL),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, FONDO]),
            ("LINEBELOW", (0, 1), (-1, -1), 0.3, GRIS_CLARO),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3.2), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.2),
            ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))
        return t
    
    
    def _seccion(titulo: str, contenido) -> KeepTogether:
        return KeepTogether([Paragraph(escape(titulo), E["h2"]), contenido])
    
    
    # ------------------------------------------------------------------ documento
    
    def generar_pdf(d: dict, autor: str = "", observaciones: str = "", acciones: str = "", logo: str | None = None) -> bytes:
        buf = io.BytesIO()
        generado = datetime.now().strftime("%d-%m-%Y %H:%M")
    
        def marco(canvas, doc):
            canvas.saveState()
            # banda superior
            canvas.setFillColor(AZUL)
            canvas.rect(0, ALTO - 20 * mm, ANCHO, 20 * mm, stroke=0, fill=1)
            canvas.setFillColor(NARANJO)
            canvas.rect(0, ALTO - 21.2 * mm, ANCHO, 1.2 * mm, stroke=0, fill=1)
            if logo and os.path.exists(logo):
                canvas.drawImage(logo_fondo_oscuro(logo), MARGEN, ALTO - 17 * mm, height=14 * mm, width=40 * mm, preserveAspectRatio=True,
                                 anchor="w", mask="auto")
            else:
                canvas.setFillColor(NARANJO)
                canvas.setFont("Helvetica-Bold", 17)
                canvas.drawString(MARGEN, ALTO - 12.5 * mm, "SYNCORE")
                canvas.setFillColor(colors.HexColor("#94A3B8"))
                canvas.setFont("Helvetica", 7)
                canvas.drawString(MARGEN, ALTO - 16.5 * mm, "MONTAJES INDUSTRIALES")
            canvas.setFillColor(colors.white)
            canvas.setFont("Helvetica-Bold", 10.5)
            canvas.drawRightString(ANCHO - MARGEN, ALTO - 10.5 * mm, "INFORME EJECUTIVO DE AVANCE")
            canvas.setFont("Helvetica", 8)
            canvas.setFillColor(colors.HexColor("#CBD5E1"))
            canvas.drawRightString(ANCHO - MARGEN, ALTO - 15 * mm, f"Data Date {_f(d['fecha_corte'])} ({CAL.etiqueta(d['fecha_corte'])})  ·  {d['proyecto_codigo']}")
            # pie
            canvas.setStrokeColor(GRIS_CLARO)
            canvas.line(MARGEN, 13 * mm, ANCHO - MARGEN, 13 * mm)
            canvas.setFont("Helvetica", 7)
            canvas.setFillColor(GRIS)
            canvas.drawString(MARGEN, 9 * mm, f"SYNCORE · Control Ejecutivo P6 · Generado el {generado}" + (f" por {autor}" if autor else ""))
            canvas.drawRightString(ANCHO - MARGEN, 9 * mm, f"Página {doc.page}")
            canvas.restoreState()
    
        doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=MARGEN, rightMargin=MARGEN, topMargin=28 * mm,
                                bottomMargin=18 * mm, title=f"Informe ejecutivo {d['valor']}", author=autor or "SYNCORE")
        h = []
        h.append(Paragraph(escape(str(d["valor"])), E["titulo"]))
        h.append(Paragraph(escape(f"{d['tipo']}  ·  Proyecto {d['proyecto_codigo']} — {d['proyecto_nombre']}"), E["subtitulo"]))
        h.append(Spacer(1, 7 * mm))
    
        delta = "–" if d["delta_pp"] is None else f"{'+' if d['delta_pp'] >= 0 else ''}{_n(d['delta_pp'], 2)} pp"
        h.append(_fila_kpis([
            ("Labor Units % Complete", f"{_n(d['avance_pct'], 2)}%", True),
            ("At Completion (HH)", _n(d["at_completion"]), False),
            ("Actual (HH)", _n(d["actual"]), False),
            ("Remaining (HH)", _n(d["remaining"]), False),
        ]))
        h.append(Spacer(1, 3 * mm))
        h.append(_fila_kpis([
            ("Actividades (completadas)", f"{d['actividades']} ({d['completadas']})", False),
            ("Variación vs corte anterior", delta, False),
            ("Inicio", _f(d["inicio"]), False),
            ("Término programado", _f(d["fin"]), False),
        ]))
        h.append(Spacer(1, 5 * mm))
        h.append(_barra_comparativa(d))
    
        h.append(Paragraph("Resumen ejecutivo", E["h2"]))
        for frase in d["resumen"]:
            h.append(Paragraph(escape(frase), E["vineta"], bulletText="•"))
    
        if observaciones.strip():
            h.append(_seccion("Análisis del planificador", _caja(observaciones)))
        if acciones.strip():
            h.append(_seccion("Acciones y compromisos", _caja(acciones)))
    
        esp = d.get("por_especialidad")
        if esp is not None and not esp.empty and len(esp) > 1:
            h.append(_seccion(f"Avance por {d['tipo_esp']}", _grafico_especialidad(esp)))
    
        h.append(PageBreak())
        h.append(_seccion("Actividades en progreso (mayor saldo primero)", _tabla(d["lista_en_progreso"], [
            ("task_code", "Activity ID", 1.3, "izq"), ("nombre", "Actividad", 4, "izq"), ("fin", "Término", 1, "fecha"),
            ("at_completion", "At Compl.", 0.9, "num"), ("remaining", "Remaining", 0.9, "num"), ("avance_pct", "% Av.", 0.7, "pct"),
        ], "No hay actividades en progreso.")))
        h.append(_seccion("Actividades pendientes con menor holgura total", _tabla(d["lista_criticas"], [
            ("task_code", "Activity ID", 1.3, "izq"), ("nombre", "Actividad", 4, "izq"), ("estado", "Estado", 1, "izq"),
            ("fin", "Término", 1, "fecha"), ("holgura_dias", "Holgura (d)", 0.9, "num1"),
        ], "No hay actividades pendientes con holgura informada.")))
        h.append(_seccion("Próximas actividades por iniciar", _tabla(d["lista_proximas"], [
            ("task_code", "Activity ID", 1.3, "izq"), ("nombre", "Actividad", 4.2, "izq"), ("inicio", "Inicio", 1, "fecha"),
            ("fin", "Término", 1, "fecha"), ("at_completion", "HH", 0.8, "num"),
        ], "No quedan actividades sin iniciar.")))
        if esp is not None and not esp.empty:
            h.append(_seccion(f"Detalle por {d['tipo_esp']}", _tabla(esp, [
                ("valor", d["tipo_esp"], 3.6, "izq"), ("actividades", "Activ.", 0.7, "num"),
                ("at_completion", "At Compl.", 1, "num"), ("actual", "Actual", 1, "num"),
                ("remaining", "Remaining", 1, "num"), ("avance_pct", "% Av.", 0.8, "pct"),
            ], "")))
        if d.get("restricciones") is not None:
            h.append(_seccion("Restricciones 6WLA pendientes", _tabla(d["restricciones"], [
                ("actividad", "Actividad", 2.4, "izq"), ("restriccion", "Restricción", 3, "izq"),
                ("responsable", "Responsable", 1.3, "izq"), ("semana", "Semana", 0.8, "izq"),
                ("fecha_compromiso", "Compromiso", 1, "fecha"),
            ], "Sin restricciones pendientes asociadas a este ítem.")))
    
        h.append(Spacer(1, 6 * mm))
        h.append(Paragraph(escape(
            "Fuente: programa Primavera P6 cargado en SYNCORE. % Complete = Actual / (Actual + Remaining) en Labor Units. "
            "Holgura total expresada en días de 8 horas."), E["subtitulo"]))
        doc.build(h, onFirstPage=marco, onLaterPages=marco)
        return buf.getvalue()
    
    
    _LOGOS_OSCUROS: dict = {}
    
    
    def logo_fondo_oscuro(ruta: str):
        """Logo listo para la franja azul del encabezado: recortado y con el azul marino pasado a blanco
        (sobre fondo oscuro "SYN" y "MONTAJES INDUSTRIALES" no se leían). Devuelve un ImageReader o la ruta."""
        try:
            clave = (ruta, os.path.getmtime(ruta))
            if clave not in _LOGOS_OSCUROS:
                from PIL import Image
                from reportlab.lib.utils import ImageReader
                im = Image.open(ruta).convert("RGBA")
                caja = im.getchannel("A").getbbox()
                if caja:
                    im = im.crop(caja)
                px = im.load()
                for y in range(im.height):
                    for x in range(im.width):
                        r, g, b, a = px[x, y]
                        if a and b > r:
                            px[x, y] = (255, 255, 255, a)
                buf = io.BytesIO()
                im.save(buf, "PNG")
                buf.seek(0)
                _LOGOS_OSCUROS[clave] = ImageReader(buf)
            return _LOGOS_OSCUROS[clave]
        except Exception:
            return ruta
    
    
    def _caja(texto: str) -> Table:
        parrafos = [Paragraph(escape(linea.strip()) or "&nbsp;", E["cuerpo"]) for linea in texto.strip().splitlines()]
        t = Table([[parrafos]], colWidths=[UTIL])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), FONDO), ("LINEBEFORE", (0, 0), (0, -1), 2.5, NARANJO),
                               ("LEFTPADDING", (0, 0), (-1, -1), 9), ("TOPPADDING", (0, 0), (-1, -1), 6),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
        return t

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


informe_pdf = _modulo_informe_pdf()

# ======================================================================
# MÓDULO: xlsx_xml
# ======================================================================
def _modulo_xlsx_xml():
    import io
    import re
    import zipfile
    from datetime import date, datetime
    from xml.sax.saxutils import escape
    
    _RE_ROW = re.compile(r"<row\b([^>]*?)(?:/>|>(.*?)</row>)", re.S)
    _RE_CELL = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
    _RE_ATTR = re.compile(r'(\w+(?::\w+)?)="([^"]*)"')
    
    
    def col_a_num(letras: str) -> int:
        n = 0
        for ch in letras:
            n = n * 26 + ord(ch) - 64
        return n
    
    
    def num_a_col(n: int) -> str:
        s = ""
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s
    
    
    def separar_ref(ref: str) -> tuple[str, int]:
        m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
        return m.group(1), int(m.group(2))
    
    
    def fecha_a_serial(d: date | datetime) -> int:
        if isinstance(d, datetime):
            d = d.date()
        return (d - date(1899, 12, 30)).days
    
    
    def serial_a_fecha(v) -> date | None:
        try:
            from datetime import timedelta
            return date(1899, 12, 30) + timedelta(days=int(float(v)))
        except (TypeError, ValueError):
            return None
    
    
    def _attrs(s: str) -> dict[str, str]:
        return dict(_RE_ATTR.findall(s or ""))
    
    
    def _xml_texto(valor: str) -> str:
        # Caracteres de control no permitidos en XML
        valor = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(valor))
        return escape(valor)
    
    
    class Libro:
        """Paquete .xlsx en memoria con acceso por nombre de hoja."""
    
        def __init__(self, datos: bytes):
            self._zip = zipfile.ZipFile(io.BytesIO(datos))
            self._partes: dict[str, bytes] = {}
            self._eliminadas: set[str] = set()
            self._sst: list[str] | None = None
            wb = self.leer("xl/workbook.xml")
            rels = self.leer("xl/_rels/workbook.xml.rels")
            destinos = {a["Id"]: a["Target"] for a in (_attrs(m) for m in re.findall(r"<Relationship\b([^>]*)/?>", rels))}
            self.hojas: dict[str, str] = {}
            for m in re.findall(r"<sheet\b([^>]*)/?>", wb):
                a = _attrs(m)
                t = destinos.get(a.get("r:id"), "")
                t = t.lstrip("/")
                self.hojas[_desescapar(a.get("name", ""))] = t if t.startswith("xl/") else "xl/" + t
    
        # ---------------------------------------------------------- partes
        def leer(self, nombre: str) -> str:
            if nombre in self._partes:
                return self._partes[nombre].decode("utf-8")
            return self._zip.read(nombre).decode("utf-8")
    
        def leer_bytes(self, nombre: str) -> bytes:
            return self._partes.get(nombre) or self._zip.read(nombre)
    
        def escribir(self, nombre: str, contenido: str | bytes):
            self._partes[nombre] = contenido.encode("utf-8") if isinstance(contenido, str) else contenido
    
        def existe(self, nombre: str) -> bool:
            return nombre in self._partes or nombre in self._zip.namelist()
    
        def eliminar(self, nombre: str):
            self._eliminadas.add(nombre)
            self._partes.pop(nombre, None)
    
        def hoja_xml(self, hoja: str) -> str:
            return self.leer(self.hojas[hoja])
    
        def guardar_hoja(self, hoja: str, xml: str):
            self.escribir(self.hojas[hoja], xml)
    
        # ---------------------------------------------------------- textos compartidos
        @property
        def sst(self) -> list[str]:
            if self._sst is None:
                self._sst = []
                if self.existe("xl/sharedStrings.xml"):
                    xml = self.leer("xl/sharedStrings.xml")
                    for si in re.findall(r"<si>(.*?)</si>", xml, flags=re.S):
                        self._sst.append(_desescapar("".join(re.findall(r"<t\b[^>]*>(.*?)</t>", si, flags=re.S))))
            return self._sst
    
        def valor_celda(self, attrs: dict, interior: str | None):
            """Valor Python de una celda (texto, número o None). Para fórmulas devuelve el valor guardado en caché."""
            if not interior:
                return None
            t = attrs.get("t")
            if t == "inlineStr":
                return _desescapar("".join(re.findall(r"<t\b[^>]*>(.*?)</t>", interior, flags=re.S)))
            m = re.search(r"<v>(.*?)</v>", interior, flags=re.S)
            if not m:
                return None
            v = m.group(1)
            if t == "s":
                try:
                    return self.sst[int(v)]
                except (ValueError, IndexError):
                    return None
            if t in ("str", "e"):
                return _desescapar(v)
            if t == "b":
                return v == "1"
            try:
                f = float(v)
                return int(f) if f.is_integer() else f
            except ValueError:
                return v
    
        # ---------------------------------------------------------- guardar
        def bytes(self, recalcular: bool = True) -> bytes:
            if recalcular:
                self._forzar_recalculo()
            salida = io.BytesIO()
            with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as z:
                for info in self._zip.infolist():
                    if info.filename in self._eliminadas:
                        continue
                    datos = self._partes.get(info.filename)
                    if datos is None:
                        datos = self._zip.read(info.filename)
                    comp = zipfile.ZIP_STORED if info.compress_type == zipfile.ZIP_STORED else zipfile.ZIP_DEFLATED
                    z.writestr(info.filename, datos, compress_type=comp)
                for nombre, datos in self._partes.items():
                    if nombre not in self._zip.namelist() and nombre not in self._eliminadas:
                        z.writestr(nombre, datos)
            return salida.getvalue()
    
        def _forzar_recalculo(self):
            """Elimina la cadena de cálculo y pide a Excel recalcular todo al abrir."""
            if self.existe("xl/calcChain.xml"):
                self.eliminar("xl/calcChain.xml")
                rels = self.leer("xl/_rels/workbook.xml.rels")
                rels = re.sub(r'<Relationship\b[^>]*Target="calcChain\.xml"[^>]*/>', "", rels)
                self.escribir("xl/_rels/workbook.xml.rels", rels)
                ct = self.leer("[Content_Types].xml")
                ct = re.sub(r'<Override\b[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', "", ct)
                self.escribir("[Content_Types].xml", ct)
            wb = self.leer("xl/workbook.xml")
            if "<calcPr" in wb:
                def ajustar(m):
                    s = m.group(0)
                    s = re.sub(r'\sfullCalcOnLoad="[^"]*"', "", s)
                    return s.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
                wb = re.sub(r"<calcPr\b[^>]*/?>", ajustar, wb, count=1)
            else:
                wb = wb.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
            self.escribir("xl/workbook.xml", wb)
    
    
    def _desescapar(s: str) -> str:
        import html
        return html.unescape(s)
    
    
    # ------------------------------------------------------------------ filas y celdas
    
    def filas(xml: str):
        """Itera (numero_fila, atributos_fila, xml_completo_fila, inicio, fin)."""
        ini_datos = xml.find("<sheetData")
        fin_datos = xml.find("</sheetData>")
        for m in _RE_ROW.finditer(xml, ini_datos, fin_datos if fin_datos > 0 else len(xml)):
            a = _attrs(m.group(1))
            yield int(a.get("r", 0)), a, m.group(0), m.start(), m.end()
    
    
    def celdas(fila_xml: str) -> dict[str, tuple[dict, str | None, str]]:
        """{columna: (atributos, interior, xml_completo)}"""
        salida = {}
        interior = re.sub(r"^<row\b[^>]*>", "", fila_xml)
        for m in _RE_CELL.finditer(interior):
            a = _attrs(m.group(1))
            col, _ = separar_ref(a["r"])
            salida[col] = (a, m.group(2), m.group(0))
        return salida
    
    
    def xml_celda(ref: str, valor, estilo: str | None = None, formula: str | None = None) -> str:
        s = f' s="{estilo}"' if estilo not in (None, "") else ""
        if formula is not None:
            # fórmula con su resultado ya calculado (valor), para que se vea bien aunque el archivo no se recalcule
            f = f"<f>{_xml_texto(formula)}</f>"
            if valor is None or valor == "" or (isinstance(valor, float) and valor != valor):
                return f'<c r="{ref}"{s}>{f}</c>'
            if isinstance(valor, bool):
                return f'<c r="{ref}"{s} t="b">{f}<v>{int(valor)}</v></c>'
            if isinstance(valor, (int, float)):
                num = int(valor) if float(valor).is_integer() else repr(float(valor))
                return f'<c r="{ref}"{s}>{f}<v>{num}</v></c>'
            return f'<c r="{ref}"{s} t="str">{f}<v>{_xml_texto(str(valor))}</v></c>'
        if valor is None or valor == "":
            return f'<c r="{ref}"{s}/>'
        if isinstance(valor, bool):
            return f'<c r="{ref}"{s} t="b"><v>{int(valor)}</v></c>'
        if isinstance(valor, (date, datetime)):
            return f'<c r="{ref}"{s}><v>{fecha_a_serial(valor)}</v></c>'
        if isinstance(valor, (int, float)):
            if isinstance(valor, float) and valor != valor:  # NaN
                return f'<c r="{ref}"{s}/>'
            return f'<c r="{ref}"{s}><v>{repr(float(valor)) if isinstance(valor, float) and not float(valor).is_integer() else int(valor)}</v></c>'
        return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{_xml_texto(valor)}</t></is></c>'
    
    
    def fijar_celdas(xml: str, valores: dict[str, object], estilos: dict[str, str] | None = None) -> str:
        """Escribe valores en celdas existentes o nuevas conservando el estilo de cada celda.
    
        valores: {"AD6": date(...), "B16": "texto", "AB34": 123, "X1": ("=SUM(A1:A2)",) -> fórmula}
        estilos: estilo (atributo s) para las celdas que no existen todavía, p. ej. copiado de la fila de arriba.
        """
        estilos = estilos or {}
        por_fila: dict[int, dict[str, object]] = {}
        for ref, v in valores.items():
            col, fila = separar_ref(ref)
            por_fila.setdefault(fila, {})[col] = v
    
        partes, ultimo = [], 0
        existentes = set()
        for num, a, fila_xml, ini, fin in filas(xml):
            if num not in por_fila:
                continue
            existentes.add(num)
            partes.append(xml[ultimo:ini])
            partes.append(_reemplazar_en_fila(fila_xml, num, por_fila[num], estilos))
            ultimo = fin
        partes.append(xml[ultimo:])
        xml = "".join(partes)
        faltantes = sorted(set(por_fila) - existentes)
        for num in faltantes:
            nueva = _reemplazar_en_fila(f'<row r="{num}"></row>', num, por_fila[num], estilos)
            xml = _insertar_fila(xml, num, nueva)
        return xml
    
    
    def _reemplazar_en_fila(fila_xml: str, num: int, nuevos: dict[str, object], estilos: dict | None = None) -> str:
        cab = re.match(r"<row\b[^>]*?(/?)>", fila_xml)
        apertura = cab.group(0)
        if apertura.endswith("/>"):
            apertura = apertura[:-2] + ">"
        actuales = celdas(fila_xml)
        todas = {}
        for col, (a, interior, completo) in actuales.items():
            todas[col] = completo
        for col, v in nuevos.items():
            ref = f"{col}{num}"
            estilo = actuales[col][0].get("s") if col in actuales else (estilos or {}).get(ref)
            if isinstance(v, tuple) and v and isinstance(v[0], str) and v[0].startswith("="):
                todas[col] = xml_celda(ref, v[1] if len(v) > 1 else None, estilo, formula=v[0][1:])
            else:
                todas[col] = xml_celda(ref, v, estilo)
        cuerpo = "".join(todas[c] for c in sorted(todas, key=col_a_num))
        apertura = re.sub(r'\sspans="[^"]*"', "", apertura)
        return f"{apertura}{cuerpo}</row>"
    
    
    def _insertar_fila(xml: str, num: int, fila_xml: str) -> str:
        for n, a, completo, ini, fin in filas(xml):
            if n > num:
                return xml[:ini] + fila_xml + xml[ini:]
        pos = xml.find("</sheetData>")
        if pos < 0:
            return xml.replace("<sheetData/>", f"<sheetData>{fila_xml}</sheetData>")
        return xml[:pos] + fila_xml + xml[pos:]
    
    
    def agregar_filas_al_final(xml: str, filas_xml: list[str]) -> str:
        pos = xml.find("</sheetData>")
        xml = xml[:pos] + "".join(filas_xml) + xml[pos:]
        return actualizar_dimension(xml)
    
    
    def ultima_fila(xml: str) -> int:
        ultimo = 0
        for m in re.finditer(r'<row\b[^>]*\br="(\d+)"', xml):
            ultimo = int(m.group(1))
        return ultimo
    
    
    def actualizar_dimension(xml: str) -> str:
        ult = ultima_fila(xml)
        def f(m):
            ref = m.group(1)
            if ":" in ref:
                a, b = ref.split(":")
                col_b, _ = separar_ref(b)
                return f'<dimension ref="{a}:{col_b}{ult}"/>'
            return m.group(0)
        return re.sub(r'<dimension ref="([^"]+)"\s*/>', f, xml, count=1)
    
    
    def renumerar_fila(fila_xml: str, nuevo: int) -> str:
        """Cambia el número de fila y las referencias de sus celdas."""
        fila_xml = re.sub(r'(<row\b[^>]*\br=")\d+(")', rf"\g<1>{nuevo}\g<2>", fila_xml, count=1)
        return re.sub(r'(<c\b[^>]*\br="[A-Z]+)\d+(")', rf"\g<1>{nuevo}\g<2>", fila_xml)
    
    
    def quitar_formulas_relativas(fila_xml: str) -> str:
        """En filas copiadas, las fórmulas compartidas que dependen de otra fila se convierten en valores."""
        def f(m):
            cel = m.group(0)
            if "<f" not in cel:
                return cel
            if re.search(r'<f\b[^>]*t="shared"', cel):
                cel = re.sub(r"<f\b[^>]*/>|<f\b[^>]*>.*?</f>", "", cel, flags=re.S)
            return cel
        return _RE_CELL.sub(f, fila_xml)
    
    
    # ------------------------------------------------------------------ insertar filas (como «Insertar fila» de Excel)
    
    _RE_REF_FX = re.compile(
        r"(?<![A-Za-z0-9_.\$])"
        r"((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?"         # hoja (opcional)
        r"(\$?[A-Z]{1,3}\$?\d+|\$?\d+(?=:\$?\d+))"                 # celda, o fila de un rango de filas completas
        r"(?::(\$?[A-Z]{1,3}\$?\d+|\$?\d+))?"                      # fin del rango
        r"(?![A-Za-z0-9_(!])")
    
    
    def _mover_ref(ref: str, en: int, n: int) -> str:
        m = re.fullmatch(r"(\$?[A-Z]{0,3}\$?)(\d+)", ref)
        if not m:
            return ref
        fila = int(m.group(2))
        return f"{m.group(1)}{fila + n if fila >= en else fila}"
    
    
    def mover_refs_formula(formula: str, hoja: str, en: int, n: int, propia: bool) -> str:
        """Corre n filas las referencias a filas >= en de la hoja «hoja» (las sin hoja cuentan si propia=True).
        Un rango que cruza la fila de inserción se agranda, igual que en Excel."""
        def f(m):
            h, a, b = m.groups()
            nombre = None
            if h:
                nombre = h[:-1]
                if nombre.startswith("'"):
                    nombre = nombre[1:-1].replace("''", "'")
            if (nombre is None and not propia) or (nombre is not None and _desescapar(nombre) != hoja):
                return m.group(0)
            return (h or "") + _mover_ref(a, en, n) + (":" + _mover_ref(b, en, n) if b else "")
        partes = formula.split('"')
        for i in range(0, len(partes), 2):                       # fuera de los textos entre comillas
            partes[i] = _RE_REF_FX.sub(f, partes[i])
        return '"'.join(partes)
    
    
    def _mover_sqref(sqref: str, en: int, n: int) -> str:
        return " ".join(":".join(_mover_ref(p, en, n) for p in r.split(":")) for r in sqref.split())
    
    
    _RE_REF_REL = re.compile(r"(?<![A-Za-z0-9_.!'\$])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_!])")
    
    
    def desplazar_relativas(formula: str, dfila: int) -> str:
        """Desplaza las filas relativas (sin $) de todas las referencias de la fórmula."""
        def f(m):
            c_abs, col, f_abs, fila = m.groups()
            return f"{c_abs}{col}{f_abs}{int(fila) + (0 if f_abs else dfila)}"
        partes = formula.split('"')
        for i in range(0, len(partes), 2):
            partes[i] = _RE_REF_REL.sub(f, partes[i])
        return '"'.join(partes)
    
    
    def insertar_filas(L: "Libro", hoja: str, en: int, n: int, modelo: int | None = None) -> bool:
        """Inserta n filas en la hoja a partir de la fila «en» (la que estaba en «en» baja n filas), con el formato,
        las celdas combinadas y las fórmulas de la fila «modelo» (por defecto, la propia «en»). Corre celdas, fórmulas
        (de todo el libro), combinadas, formatos condicionales, validaciones, saltos de página, imágenes y nombres
        definidos (área de impresión). Devuelve False si no se pudo hacer sin romper fórmulas compartidas."""
        if n <= 0:
            return True
        modelo = modelo or en
        xml = L.hoja_xml(hoja)
        # fórmulas compartidas que cruzan la fila de inserción: no se pueden correr por partes
        for m in re.finditer(r'<f\b[^>]*t="shared"[^>]*ref="([^"]+)"', xml):
            r = m.group(1).split(":")
            if len(r) == 2 and separar_ref(r[0])[1] < en <= separar_ref(r[1])[1]:
                return False
    
        def fx(m):
            attrs, texto = m.group(1), m.group(2)
            attrs = re.sub(r'\bref="([^"]+)"', lambda mm: f'ref="{_mover_sqref(mm.group(1), en, n)}"', attrs)
            if texto is None:
                return f"<f{attrs}/>"
            return f"<f{attrs}>{mover_refs_formula(_desescapar(texto), hoja, en, n, True).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')}</f>"
        re_f = re.compile(r"<f\b([^>]*?)(?:/>|>(.*?)</f>)", re.S)
    
        ini, fin = xml.find("<sheetData"), xml.find("</sheetData>")
        datos = xml[ini:fin]
        fila_modelo = next((completo for num, _, completo, _, _ in filas(xml) if num == modelo), None)
    
        def fila_nueva(m):
            completo = m.group(0)
            num = int(_attrs(m.group(1)).get("r", 0))
            if num >= en:
                completo = renumerar_fila(completo, num + n)
            return re_f.sub(fx, completo)
        datos = _RE_ROW.sub(fila_nueva, datos)
        # filas copiadas de la fila modelo
        copias = []
        if fila_modelo:
            base = quitar_formulas_relativas(fila_modelo)
            base = re.sub(r'(<f\b[^>]*?)\s(?:t="shared"|ref="[^"]*"|si="\d+")', r"\1", base)
            for k in range(n):
                nueva = renumerar_fila(base, en + k)
                nueva = re.sub(r"<f\b([^>]*)>(.*?)</f>",
                               lambda mm: f"<f{mm.group(1)}>" + desplazar_relativas(_desescapar(mm.group(2)), en + k - modelo)
                               .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;") + "</f>", nueva, flags=re.S)
                copias.append(nueva)
        if copias:
            pos = next((m.start() for m in _RE_ROW.finditer(datos) if int(_attrs(m.group(1)).get("r", 0)) >= en + n), None)
            pos = len(datos) if pos is None else pos
            datos = datos[:pos] + "".join(copias) + datos[pos:]
        xml = xml[:ini] + datos + xml[fin:]
        resto_ini = xml.find("</sheetData>")
        cabeza, resto = xml[:resto_ini], xml[resto_ini:]
        # celdas combinadas
        combinadas_modelo = [r for r in re.findall(r'<mergeCell ref="([^"]+)"', resto)
                             if all(separar_ref(p)[1] == modelo for p in r.split(":"))]
        resto = re.sub(r'(<mergeCell ref=")([^"]+)(")', lambda mm: mm.group(1) + _mover_sqref(mm.group(2), en, n) + mm.group(3), resto)
        if combinadas_modelo:
            extra = []
            for k in range(n):
                for r in combinadas_modelo:
                    a, b = r.split(":")
                    extra.append(f'<mergeCell ref="{separar_ref(a)[0]}{en + k}:{separar_ref(b)[0]}{en + k}"/>')
            resto = re.sub(r'<mergeCells count="(\d+)"', lambda mm: f'<mergeCells count="{int(mm.group(1)) + len(extra)}"', resto, count=1)
            resto = resto.replace("</mergeCells>", "".join(extra) + "</mergeCells>", 1)
        # formatos condicionales, validaciones y fórmulas que contienen
        resto = re.sub(r'(\bsqref=")([^"]+)(")', lambda mm: mm.group(1) + _mover_sqref(mm.group(2), en, n) + mm.group(3), resto)
        resto = re.sub(r"(<xm:sqref>)([^<]+)(</xm:sqref>)", lambda mm: mm.group(1) + _mover_sqref(mm.group(2), en, n) + mm.group(3), resto)
        resto = re.sub(r"(<(formula|formula1|formula2|xm:f)>)(.*?)(</\2>)",
                       lambda mm: mm.group(1) + mover_refs_formula(mm.group(3), hoja, en, n, True) + mm.group(4), resto, flags=re.S)
        # saltos de página (id = fila tras la que se corta)
        resto = re.sub(r'(<brk id=")(\d+)(")', lambda mm: f"{mm.group(1)}{int(mm.group(2)) + n if int(mm.group(2)) >= en else mm.group(2)}{mm.group(3)}", resto)
        resto = re.sub(r'(<hyperlink\b[^>]*\bref=")([^"]+)(")', lambda mm: mm.group(1) + _mover_sqref(mm.group(2), en, n) + mm.group(3), resto)
        xml = actualizar_dimension(cabeza + resto)
        L.guardar_hoja(hoja, xml)
        # imágenes y formas (anclas 0-based)
        rels = L.hojas[hoja].replace("worksheets/", "worksheets/_rels/") + ".rels"
        if L.existe(rels):
            for t in re.findall(r'Target="\.\./drawings/(drawing\d+\.xml)"', L.leer(rels)):
                ruta = f"xl/drawings/{t}"
                d = L.leer(ruta)
                d = re.sub(r"(<xdr:row>)(\d+)(</xdr:row>)",
                           lambda mm: f"{mm.group(1)}{int(mm.group(2)) + n if int(mm.group(2)) + 1 >= en else mm.group(2)}{mm.group(3)}", d)
                L.escribir(ruta, d)
        # nombres definidos (área de impresión, filtros…) y fórmulas de otras hojas que apuntan a esta
        wb = L.leer("xl/workbook.xml")
        wb = re.sub(r"(<definedName\b[^>]*>)(.*?)(</definedName>)",
                    lambda mm: mm.group(1) + mover_refs_formula(_desescapar(mm.group(2)), hoja, en, n, False)
                    .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;") + mm.group(3), wb, flags=re.S)
        L.escribir("xl/workbook.xml", wb)
        clave = escape(hoja)
        for otra, ruta in L.hojas.items():
            if otra == hoja:
                continue
            x2 = L.leer(ruta)
            if clave not in x2:
                continue
            x2n = re_f.sub(lambda mm: mm.group(0) if mm.group(2) is None else
                           f"<f{mm.group(1)}>" + mover_refs_formula(_desescapar(mm.group(2)), hoja, en, n, False)
                           .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;") + "</f>", x2)
            if x2n != x2:
                L.escribir(ruta, x2n)
        return True

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


xlsx_xml = _modulo_xlsx_xml()

# ======================================================================
# MÓDULO: diario
# ======================================================================
def _modulo_diario():
    import io
    import math
    import re
    import unicodedata
    from dataclasses import dataclass, field
    from datetime import date, datetime, time, timedelta
    
    X = xlsx_xml
    
    HORAS_JORNADA = 11
    ESTADOS_TRABAJANDO = {"T", "TN", "TT", "TC"}
    ESTADOS = {
        "T": "Trabajando turno día", "TN": "Trabajando turno noche", "TT": "Trabajando (TT)", "TC": "Trabajando en curso (TC)",
        "D": "Descanso", "L": "Licencia", "PSG": "Permiso sin goce", "PCG": "Permiso con goce", "F": "Falla",
        "AC": "En acreditación", "V": "Vacaciones",
    }
    
    
    def norm(s) -> str:
        s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().upper()
        return re.sub(r"\s+", " ", s).strip()
    
    
    def norm_rut(s) -> str:
        s = str(s or "").upper().replace(".", "").replace(",", "").replace(" ", "").strip()
        return s.lstrip("0")
    
    
    # Texto que viene en vez de horas → estado en BD Personal (mismos códigos que usa el informe oficial)
    ESTADO_POR_NOTA = [("SINDICAL", "TC"), ("LICENCIA", "L"), ("LM", "L"), ("NO SUBIO", "F"), ("FALTA", "F"), ("AUSENTE", "F"),
                       ("INASISTENCIA", "F"), ("VACACION", "V"), ("PERMISO", "PCG")]
    
    
    def estado_por_nota(nota: str) -> str | None:
        n = norm(nota)
        for clave, estado in ESTADO_POR_NOTA:
            if re.search(rf"\b{clave}", n):
                return estado
        return None
    
    
    def grupo_turno(turno) -> str:
        """«14x14 a» → «14X14 A»; «14X14 B DESFASADO» se mantiene como grupo propio."""
        return re.sub(r"\s+", " ", norm(turno)).strip()
    
    
    def ciclo_turno(grupo: str) -> int | None:
        m = re.search(r"(\d+)\s*X\s*(\d+)", grupo or "")
        return int(m.group(1)) + int(m.group(2)) if m else None
    
    
    def _trabaja(par) -> bool | None:
        if not par or not par[1]:
            return None
        return par[0] / par[1] >= 0.5
    
    
    def turno_trabaja(historial: dict, grupo: str, fecha: date) -> bool | None:
        """¿El grupo de turno trabaja ese día? Se toma el mismo día de ciclos anteriores en el historial de BD Personal."""
        ciclo = ciclo_turno(grupo)
        dias = historial.get(grupo) or {}
        if not ciclo or not dias:
            return None
        for k in range(1, 5):
            r = _trabaja(dias.get(fecha - timedelta(days=ciclo * k)))
            if r is not None:
                return r
        return None
    
    
    def _tokens_nombre(nombre: str) -> frozenset[str]:
        return frozenset(t for t in re.split(r"[^A-Z]+", norm(nombre)) if len(t) > 1)
    
    
    # ====================================================================== reportes de terreno
    
    @dataclass
    class Trabajador:
        rut: str
        nombre: str
        cargo: str
        horas: dict[str, float]          # {letra actividad: horas}
        total: float
        nota: str = ""                   # texto en vez de horas (p. ej. "Permiso sindical", "Licencia")
    
    
    @dataclass
    class Actividad:
        codigo: str
        area: str
        descripcion: str
        estado: str = ""
        extra: str = ""        # resto de la fila del descriptor de tareas (Activity ID, subsistema, tags…)
    
    
    @dataclass
    class ReporteTerreno:
        archivo: str
        fecha: date | None
        supervisor: str
        turno_texto: str
        turno: str                       # "Día" o "Noche"
        area_reporte: str
        trabajadores: list[Trabajador] = field(default_factory=list)
        actividades: list[Actividad] = field(default_factory=list)
        observaciones: str = ""
        avisos: list[str] = field(default_factory=list)
        tipo: str = "terreno"            # terreno | uso_recurso | adicionales | programacion
        especialidad: str = ""           # cuando el archivo la indica (programación, adicionales)
        uso_recurso: list[dict] = field(default_factory=list)   # filas para «7.0 Uso de recurso»
        materiales: list[dict] = field(default_factory=list)    # materiales / consumibles declarados (formulario UDR)
        precom: bool = False             # reporte de terreno de precomisionamiento (formulario «Trabajos de pre-com»)
    
        @property
        def total_hh(self) -> float:
            return sum(t.total for t in self.trabajadores)
    
        def a_dict(self) -> dict:
            return {
                "archivo": self.archivo, "fecha": self.fecha.isoformat() if self.fecha else None,
                "supervisor": self.supervisor, "turno_texto": self.turno_texto, "turno": self.turno,
                "area_reporte": self.area_reporte, "observaciones": self.observaciones, "avisos": self.avisos,
                "trabajadores": [t.__dict__ for t in self.trabajadores],
                "actividades": [a.__dict__ for a in self.actividades],
                "tipo": self.tipo, "especialidad": self.especialidad, "uso_recurso": self.uso_recurso,
                "materiales": self.materiales, "precom": self.precom,
            }
    
        @staticmethod
        def de_dict(d: dict) -> "ReporteTerreno":
            r = ReporteTerreno(d["archivo"], date.fromisoformat(d["fecha"]) if d.get("fecha") else None, d.get("supervisor", ""),
                               d.get("turno_texto", ""), d.get("turno", "Día"), d.get("area_reporte", ""),
                               observaciones=d.get("observaciones", ""), avisos=d.get("avisos", []),
                               tipo=d.get("tipo", "terreno"), especialidad=d.get("especialidad", ""),
                               uso_recurso=d.get("uso_recurso", []), materiales=d.get("materiales", []),
                               precom=d.get("precom", False))
            r.trabajadores = [Trabajador(**t) for t in d.get("trabajadores", [])]
            r.actividades = [Actividad(**a) for a in d.get("actividades", [])]
            return r
    
    
    def _num(v) -> float:
        if v is None or v == "":
            return 0.0
        if isinstance(v, (int, float)):
            return float(v)
        try:
            return float(str(v).replace(",", "."))
        except ValueError:
            return 0.0
    
    
    def _txt(v) -> str:
        if v is None:
            return ""
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        return str(v).replace("_x000D_", " ").strip()
    
    
    def _es_hora(v) -> bool:
        if isinstance(v, (int, float)):
            return True
        try:
            float(str(v).replace(",", ".").strip())
            return True
        except ValueError:
            return False
    
    
    def _grilla(ws, max_vacias: int = 60, max_filas: int = 600) -> list[list]:
        """Valores de la hoja hasta muchas filas vacías seguidas (algunas planillas declaran miles de filas con formato)."""
        grid, vacias = [], 0
        for fila in ws.iter_rows(max_row=max_filas, max_col=70, values_only=True):
            fila = list(fila)
            if any(v not in (None, "") for v in fila):
                grid.extend([[]] * vacias)
                vacias = 0
                grid.append(fila)
            else:
                vacias += 1
                if vacias > max_vacias:
                    break
        return grid
    
    
    def _fecha_en(grid, archivo: str, filas: int = 10) -> date | None:
        for fila in grid[:filas]:
            for v in fila:
                if isinstance(v, datetime) and 2020 <= v.year <= 2100:
                    return v.date()
        m = re.search(r"(\d{1,2})[-_.](\d{1,2})[-_.](20\d{2}|\d{2})\b", archivo)
        if m:
            a = int(m.group(3))
            return date(a + 2000 if a < 100 else a, int(m.group(2)), int(m.group(1)))
        return None
    
    
    def _cabecera(grid, requeridas: list[set[str]], max_fila: int = 80) -> tuple[int, list[str]] | None:
        """Primera fila cuyas celdas contienen, para cada grupo, al menos uno de sus nombres (normalizados)."""
        for i, f in enumerate(grid[:max_fila]):
            nombres = [norm(v) for v in f]
            if all(any(n in grupo for n in nombres) for grupo in requeridas):
                return i, nombres
        return None
    
    
    def _col(nombres: list[str], *opciones: str, empieza: bool = False) -> int | None:
        for o in opciones:
            for j, n in enumerate(nombres):
                if (n.startswith(o) if empieza else n == o):
                    return j
        return None
    
    
    def _tipo_hoja(grid) -> str:
        """Reconoce el formato por su contenido, no por el nombre del archivo."""
        plano = " ".join(norm(v) for f in grid[:25] for v in f if isinstance(v, str))
        personal = _cabecera(grid, [{"RUT"}, {"NOMBRE"}])
        if ("USO DE RECURSO" in plano or "TRABAJOS DE PRE-COM" in plano or "TRABAJOS DE PRECOM" in plano
                or "TRABAJOS ADICIONALES" in plano) and (personal or "ACTIVIDADES" in plano):
            return "uso_recurso"
        if personal:
            i, nombres = personal
            letras = [n for n in nombres if re.fullmatch(r"[A-Z]{1,2}", n)]
            if _col(nombres, "CARGO") is not None and len(letras) >= 2:
                return "terreno"
            return "nomina"
        if _cabecera(grid, [{"ACTIVIDAD", "ACTIVIDADES", "DESCRIPCION"}, {"PERSONAL", "DOTACION", "CANTIDAD DE PERSONAS"},
                            {"HH", "HH GASTO", "HH GASTADAS", "HH DESVIADAS", "HORAS HOMBRE", "HH TOTALES"}]):
            return "adicionales"
        if _cabecera(grid, [{"ACTIVIDAD", "ACTIVIDADES"}, {"ESPECIALIDAD"}]):
            return "programacion"
        return ""
    
    
    _ABC = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ"
    
    
    def _letra(v) -> str:
        """Código de actividad del reporte (A, B, … N, Ñ, O …). A diferencia de norm(), la Ñ NO se convierte en N:
        en los reportes «N» y «Ñ» son actividades distintas."""
        s = str(v if v is not None else "").strip().upper().replace(" ", "")
        s = unicodedata.normalize("NFC", s)
        return s if re.fullmatch(r"[A-ZÑ]{1,2}", s) else norm(s).replace(" ", "")
    
    
    def _orden_letra(s: str):
        return (len(s), [_ABC.index(ch) if ch in _ABC else 99 for ch in s])
    
    
    def _turno_de(textos: list[str]) -> str:
        return "Noche" if any("NOCHE" in norm(t) for t in textos) else "Día"
    
    
    def _leer_terreno(grid, archivo: str) -> ReporteTerreno:
        fecha = _fecha_en(grid, archivo)
        supervisor = turno_txt = ""
        for fila in grid[:6]:
            for v in fila:
                t = _txt(v)
                if norm(t).startswith("SUPERVISOR"):
                    supervisor = t.split(":", 1)[-1].strip()
                if "TURNO" in norm(t):
                    turno_txt = t.split(":", 1)[-1].strip() if ":" in t else t
        rep = ReporteTerreno(archivo, fecha, supervisor, turno_txt, _turno_de([turno_txt, archivo]), "")
    
        fila_cab, cab = _cabecera(grid, [{"RUT"}, {"NOMBRE"}, {"CARGO"}])
        c_rut, c_nom, c_car = cab.index("RUT"), cab.index("NOMBRE"), cab.index("CARGO")
        c_tot = next((j for j, v in enumerate(cab) if v.startswith("TOTAL") and j > c_car), None)
        crudo = [_letra(v) for v in grid[fila_cab]] + [""] * len(cab)
        letras = {j: crudo[j] for j, v in enumerate(cab) if re.fullmatch(r"[A-ZÑ]{1,2}", crudo[j]) and j > c_car
                  and (c_tot is None or j < c_tot)}
        i = fila_cab + 1
        while i < len(grid):
            f = grid[i] + [None] * 70
            inicio = [norm(x) for x in f[:3]]
            if any(x.startswith("TOTAL") or x.startswith("DESCRIPCION") or x.startswith("AREA") for x in inicio) or inicio[0] == "ACT":
                break
            rut, nom = _txt(f[c_rut]), _txt(f[c_nom])
            if nom.startswith("#") and not rut:          # filas de fórmula sin datos (#N/A)
                nom = ""
            if (rut or nom) and not norm(nom).startswith("TOTAL"):
                horas, notas = {}, []
                for j, l in letras.items():
                    v = f[j]
                    if v in (None, ""):
                        continue
                    if _es_hora(v):
                        if _num(v):
                            horas[l] = horas.get(l, 0) + _num(v)
                    else:
                        notas.append(_txt(v))
                total = _num(f[c_tot]) if c_tot is not None and _es_hora(f[c_tot]) else 0.0
                total = total or sum(horas.values())
                rep.trabajadores.append(Trabajador(norm_rut(rut), nom, _txt(f[c_car]), horas, total, " · ".join(dict.fromkeys(notas))))
            i += 1
    
        # Descripción de actividades: fila con «ACT» o con «Área/Subsistema» y «Descripción»
        fila_act = None
        for k in range(i, len(grid)):
            n = [norm(v) for v in grid[k]]
            if (n and n[0] == "ACT") or (any(x.startswith("AREA") for x in n) and any("DESCRIPCION" in x for x in n)):
                fila_act = k
                break
        orden_letras = sorted(set(letras.values()), key=_orden_letra)
        if fila_act is not None:
            cab_a = [norm(v) for v in grid[fila_act]]
            c_area = next((j for j, v in enumerate(cab_a) if v.startswith("AREA")), 1)
            c_desc = next((j for j, v in enumerate(cab_a) if "DESCRIPCION" in v), 6)
            c_est = next((j for j, v in enumerate(cab_a) if v.startswith("STATUS") or v.startswith("ESTADO")), None)
            k, pos = fila_act + 1, 0
            while k < len(grid):
                f = grid[k] + [None] * 40
                if norm(f[0]).startswith("OBSERVACION") or norm(f[0]) == "IMAGEN":
                    break
                desc = _txt(f[c_desc])
                cod = _letra(f[0])
                if desc:
                    if not re.fullmatch(r"[A-ZÑ]{1,2}", cod):
                        corr = orden_letras[pos] if pos < len(orden_letras) else cod
                        rep.avisos.append(f"Actividad n°{pos + 1} con código '{_txt(f[0])}': se asumió '{corr}' por su posición.")
                        cod = corr
                    extra = " ".join(_txt(v) for j, v in enumerate(f) if j not in (0, c_desc, c_area) and _txt(v))
                    rep.actividades.append(Actividad(cod, _txt(f[c_area]), desc, _txt(f[c_est]) if c_est is not None else "", extra))
                if desc or cod:
                    pos += 1
                k += 1
            obs = [_txt(v) for f in grid[k + 1:] for v in f if _txt(v)]
            rep.observaciones = "\n".join(obs)
    
        areas = [a.area for a in rep.actividades if a.area]
        rep.area_reporte = max(set(areas), key=areas.count) if areas else ""
        usadas = {l for t in rep.trabajadores for l in t.horas}
        sin_desc = sorted(usadas - {a.codigo for a in rep.actividades})
        if sin_desc:
            rep.avisos.append(f"Hay horas en actividades sin descripción: {', '.join(sin_desc)}.")
        for t in rep.trabajadores:
            if t.nota and not t.total:
                rep.avisos.append(f"{t.nombre}: «{t.nota}» (sin horas, no se marca presente).")
        if fecha is None:
            rep.avisos.append("El reporte no indica fecha.")
        return rep
    
    
    def _especialidad_udr(cargos: list[str], especialidades: list[str], defecto: str = "") -> str:
        """Disciplina del UDR como en el informe: Eléctrica, Instrumentación, Mecánica, Cañería, Andamios, lototo, Obras Civiles…"""
        texto = " ".join(norm(c) for c in cargos)
        esp = " ".join(norm(e) for e in especialidades) + " " + norm(defecto)
        if "LOTO" in esp:
            return "lototo"
        if texto.count("CANERI") > len(cargos) / 2:
            return "Cañería"
        for clave, nombre in (("ANDAMI", "Andamios"), ("INSTRUMENT", "Instrumentación"), ("ELECTRIC", "Eléctrica"),
                              ("MECANIC", "Mecánica"), ("CANERI", "Cañería"), ("PIPING", "Cañería"), ("CIVIL", "Obras Civiles"),
                              ("SOLDA", "Soldadura"), ("PINTU", "Pintura"), ("AISLA", "Aislación")):
            if clave in esp:
                return nombre
        return (especialidades[0] if especialidades and especialidades[0] else defecto).title()
    
    
    def _leer_uso_recurso(grid, archivo: str) -> ReporteTerreno:
        """Formato UDR («Reporte: uso de recursos trabajos adicionales»): personal con RUT y HH, y sus actividades."""
        fecha = _fecha_en(grid, archivo, 15)
        etiquetas: dict[str, str] = {}
        for f in grid[:14]:
            for j, v in enumerate(f):
                t = norm(v)
                for clave in ("AREA", "DESCRIPCION", "TRABAJO"):
                    if t.startswith(clave) and clave not in etiquetas:
                        etiquetas[clave] = next((_txt(x) for x in f[j + 1:] if _txt(x) and not norm(x).startswith("N FOLIO")), "")
        rep = ReporteTerreno(archivo, fecha, "", "", _turno_de([archivo] + list(etiquetas.values())), etiquetas.get("AREA", ""),
                             tipo="uso_recurso")
        i, cab = _cabecera(grid, [{"RUT"}, {"NOMBRE"}])
        c_rut, c_nom = cab.index("RUT"), cab.index("NOMBRE")
        c_car = _col(cab, "CATEGORIA", "CARGO")
        c_esp = _col(cab, "ESPECIALIDAD")
        c_tot = _col(cab, "HH TOTALES", "TOTAL", "HH TOTAL", empieza=True)
        c_hh = _col(cab, "HH NORMALES", "HH", empieza=True)
        personas, esp_p = [], []
        for f in grid[i + 1:]:
            f = f + [None] * 40
            if any(norm(x).startswith(("MATERIALES", "ACTIVIDADES", "INGENIERIA")) for x in f[:6] if isinstance(x, str)):
                break
            rut, nom = _txt(f[c_rut]), _txt(f[c_nom])
            if not (rut or nom):
                continue
            hh = _num(f[c_tot]) if c_tot is not None else 0.0
            hh = hh or (_num(f[c_hh]) if c_hh is not None else 0.0)
            cargo = _txt(f[c_car]) if c_car is not None else ""
            if not hh and not _txt(f[c_rut]) and re.fullmatch(r"\d+", nom or ""):
                continue                                       # fila de totales / numeración
            personas.append(Trabajador(norm_rut(rut), nom, cargo, {}, hh))
            esp_p.append(_txt(f[c_esp]) if c_esp is not None else "")
            if c_esp is not None and _txt(f[c_esp]) and not rep.especialidad:
                rep.especialidad = _txt(f[c_esp])
        rep.trabajadores = personas
        # Materiales del formulario (N° | Descripción | Unidad | Cantidad)
        k_mat = next((k for k, f in enumerate(grid) if any(norm(x) == "MATERIALES" for x in f[:4] if isinstance(x, str))), None)
        if k_mat is not None:
            cm = _cabecera(grid[k_mat:k_mat + 4], [{"DESCRIPCION"}, {"CANTIDAD"}])
            if cm:
                k0, cab_m = cm
                c_d, c_u, c_c = _col(cab_m, "DESCRIPCION"), _col(cab_m, "UNIDAD"), _col(cab_m, "CANTIDAD")
                for f in grid[k_mat + k0 + 1:]:
                    f = list(f) + [None] * 40
                    if any(norm(x).startswith(("ACTIVIDADES", "INGENIERIA", "COMENTARIO")) for x in f[:4] if isinstance(x, str)):
                        break
                    d_ = _txt(f[c_d]) if c_d is not None else ""
                    if d_ and norm(d_) != "DESCRIPCION":
                        rep.materiales.append({"descripcion": d_, "unidad": _txt(f[c_u]) if c_u is not None else "",
                                               "cantidad": _num(f[c_c]) if c_c is not None else None})
        # Actividades del UDR: la tabla bajo el título «Actividades» (no la de materiales)
        k_act = next((k for k, f in enumerate(grid) if any(norm(x) == "ACTIVIDADES" for x in f[:4] if isinstance(x, str))), None)
        act = _cabecera(grid[k_act:] if k_act is not None else grid, [{"DESCRIPCION"}, {"TOTAL", "HM", "CANTIDAD"}], max_fila=len(grid))
        if act and k_act is not None:
            act = (act[0] + k_act, act[1])
        descripciones = []
        if act:
            k, ca = act
            for f in grid[k + 1:]:
                f = f + [None] * 40
                if any(norm(x).startswith(("INGENIERIA", "COMENTARIO")) for x in f[:6] if isinstance(x, str)):
                    break
                d = _txt(f[_col(ca, "DESCRIPCION")])
                if d and d not in descripciones and norm(d) != "DESCRIPCION":
                    descripciones.append(d)
        for d in descripciones:
            rep.actividades.append(Actividad("", rep.area_reporte, d))
        # Equipos del formulario (N° | Descripción | Cantidad | HM | … | Total) -> observaciones y HM del UDR
        equipos_txt, hm_total = [], 0.0
        k_eq = next((k for k, f in enumerate(grid) if any(norm(x) == "EQUIPOS" for x in f[:4] if isinstance(x, str))), None)
        if k_eq is not None:
            ce = _cabecera(grid[k_eq:k_eq + 4], [{"DESCRIPCION"}, {"CANTIDAD"}])
            if ce:
                k0, cab_e = ce
                c_d, c_c = _col(cab_e, "DESCRIPCION"), _col(cab_e, "CANTIDAD")
                c_t = _col(cab_e, "TOTAL") or _col(cab_e, "HM", empieza=True)
                for f in grid[k_eq + k0 + 1:]:
                    f = list(f) + [None] * 40
                    if any(norm(x).startswith(("ACTIVIDADES", "INGENIERIA", "COMENTARIO", "MATERIALES")) for x in f[:4] if isinstance(x, str)):
                        break
                    d_ = _txt(f[c_d]) if c_d is not None else ""
                    if d_:
                        cant = _num(f[c_c]) if c_c is not None else 0
                        equipos_txt.append(f"{d_} X{int(cant)}" if cant and f" X{int(cant)}" not in norm(d_) else d_)
                        hm_total += _num(f[c_t]) if c_t is not None else 0
        plano_tit = norm(" ".join(str(x) for f in grid[:12] for x in f if isinstance(x, str)))
        es_precom = "PRE-COM" in plano_tit or "PRECOM" in norm(etiquetas.get("DESCRIPCION", ""))
        if es_precom:
            hoja = norm(archivo.split(" › ")[-1])
            etiquetas["AREA_REAL"] = etiquetas.get("AREA", "")
            etiquetas["AREA"] = "Apoyo a Precom" if "APOYO" in hoja or "APOYO" in norm(etiquetas.get("TRABAJO", "")) else "Precom y protocolos"
        titulo = etiquetas.get("AREA") or etiquetas.get("TRABAJO") or etiquetas.get("DESCRIPCION") or ""
        detalle = "; ".join(descripciones) or etiquetas.get("DESCRIPCION") or etiquetas.get("TRABAJO") or "Uso de recurso"
        base = {"descripcion": titulo or detalle, "detalle": detalle, "turno": rep.turno, "origen": archivo,
                "equipos": " / ".join(equipos_txt), "hm": hm_total or None,
                **({"area": "Precom"} if es_precom else {}),
                "respaldo": respaldo_udr(" ".join(etiquetas.values()) + " " + detalle)}
        if personas:
            # una fila por formulario (hoja), como en el informe oficial; directos / indirectos se separan al consolidar
            # Conductores / operadores del formulario = equipos del UDR (camión aljibe X4 …) con sus HM
            eq_op: dict[str, list] = {}
            for p_ in personas:
                m_ = re.match(r"(?:CONDUCTOR|OPERADOR|CHOFER)\s+(?:DE\s+)?(.+)", norm(p_.cargo))
                if m_:
                    eq_op.setdefault(m_.group(1).replace("ALGIBE", "ALJIBE").strip(), []).append(p_.total)
            for nombre_eq, hs in eq_op.items():
                equipos_txt.append(f"{nombre_eq} X{len(hs)}")
                hm_total += sum(hs)
            esp = _especialidad_udr([p_.cargo for p_ in personas], esp_p, rep.especialidad)
            if es_precom:
                # también es un reporte de terreno: cada persona trabajó en las actividades del formulario (BD Personal / Master)
                rep.precom = True
                rep.supervisor = next((p_.nombre for p_ in personas if "SUPERV" in norm(p_.cargo)), "")
                rep.actividades = [Actividad("PRECOM", "Precom", detalle, "", " ".join(etiquetas.values()))]
                for p_ in personas:
                    if p_.total:
                        p_.horas = {"PRECOM": p_.total}
            rep.uso_recurso.append({**base, "equipos": " / ".join(equipos_txt), "hm": hm_total or None,
                                    "directos": len(personas), "hh_directos": sum(p.total for p in personas),
                                    "especialidad": esp, "ruts": [(p.rut, p.total, p.cargo) for p in personas]})
        elif descripciones or titulo:
            rep.uso_recurso.append({**base, "directos": None, "hh_directos": None, "especialidad": rep.especialidad})
        return rep
    
    
    def _leer_adicionales(grid, archivo: str) -> ReporteTerreno:
        """Tabla de adicionales / desviaciones: una fila por actividad con personal y HH (sin nombres)."""
        i, cab = _cabecera(grid, [{"ACTIVIDAD", "ACTIVIDADES", "DESCRIPCION"}, {"PERSONAL", "DOTACION", "CANTIDAD DE PERSONAS"},
                                  {"HH", "HH GASTO", "HH GASTADAS", "HH DESVIADAS", "HORAS HOMBRE", "HH TOTALES"}])
        c_act = _col(cab, "ACTIVIDAD", "ACTIVIDADES", "DESCRIPCION")
        c_per = _col(cab, "PERSONAL", "DOTACION", "CANTIDAD DE PERSONAS")
        c_hh = _col(cab, "HH GASTO", "HH GASTADAS", "HH DESVIADAS", "HH TOTALES", "HORAS HOMBRE", "HH")
        c_fec, c_jor = _col(cab, "FECHA"), _col(cab, "JORNADA", "TURNO")
        c_esp, c_area = _col(cab, "ESPECIALIDAD"), _col(cab, "AREA", "SECTOR")
        c_resp = _col(cab, "RESPALDO", "REFERENCIA", "OBSERVACION")
        rep = ReporteTerreno(archivo, None, "", "", "Día", "", tipo="adicionales")
        turnos, fechas = [], []
        for f in grid[i + 1:]:
            f = f + [None] * 40
            desc = _txt(f[c_act])
            if not desc:
                continue
            jor = _txt(f[c_jor]) if c_jor is not None else ""
            turnos.append(jor)
            if c_fec is not None and isinstance(f[c_fec], datetime):
                fechas.append(f[c_fec].date())
            esp = _txt(f[c_esp]) if c_esp is not None else ""
            area = _txt(f[c_area]) if c_area is not None else ""
            rep.uso_recurso.append({"descripcion": desc, "directos": _num(f[c_per]), "hh_directos": _num(f[c_hh]),
                                    "especialidad": esp, "area": area, "turno": _turno_de([jor]),
                                    "respaldo": _txt(f[c_resp]) if c_resp is not None else "", "origen": archivo})
            rep.actividades.append(Actividad("", area, desc, "", " ".join(_txt(v) for v in f if _txt(v) and _txt(v) not in (desc, area))))
            if esp and not rep.especialidad:
                rep.especialidad = esp
        rep.fecha = max(set(fechas), key=fechas.count) if fechas else _fecha_en(grid, archivo)
        rep.turno = "Noche" if turnos and all(_turno_de([t]) == "Noche" for t in turnos) else "Día"
        return rep
    
    
    def _leer_programacion(grid, archivo: str) -> ReporteTerreno:
        """Programación de actividades (sin personal): actividades por especialidad y sector."""
        i, cab = _cabecera(grid, [{"ACTIVIDAD", "ACTIVIDADES"}, {"ESPECIALIDAD"}])
        c_act, c_esp = _col(cab, "ACTIVIDADES", "ACTIVIDAD"), _col(cab, "ESPECIALIDAD")
        c_area = _col(cab, "SECTOR", "AREA", "AREA/SUBSISTEMA")
        titulo = " ".join(_txt(v) for f in grid[:i] for v in f if isinstance(v, str))
        rep = ReporteTerreno(archivo, _fecha_en(grid, archivo), "", titulo, _turno_de([titulo, archivo]), "", tipo="programacion")
        for f in grid[i + 1:]:
            f = f + [None] * 40
            desc = re.sub(r"\s+", " ", _txt(f[c_act]).lstrip("*")).strip()
            if not desc:
                continue
            rep.actividades.append(Actividad("", _txt(f[c_area]) if c_area is not None else "", desc))
            if not rep.especialidad and _txt(f[c_esp]):
                rep.especialidad = _txt(f[c_esp])
        rep.area_reporte = rep.actividades[0].area if rep.actividades else ""
        rep.avisos.append("Programación (sin personal): sus actividades quedan como sugerencia, sin marcar, para el texto del informe.")
        return rep
    
    
    LECTORES = {"terreno": _leer_terreno, "uso_recurso": _leer_uso_recurso, "adicionales": _leer_adicionales,
                "programacion": _leer_programacion}
    NOMBRE_TIPO = {"terreno": "Reporte de terreno", "uso_recurso": "Uso de recurso (UDR)", "adicionales": "Adicionales / desviaciones",
                   "programacion": "Programación de actividades", "precom": "Reporte de terreno Precom (+ UDR)",
                   "pm_materiales": "PM de materiales (PDF)"}
    
    
    # ---------------------------------------------------------------- PM de materiales (PDF)
    
    _PM_DESC = re.compile(r"DESCRIP|MATERIAL|DETALLE|ARTICULO|PRODUCTO|INSUMO|ITEM DESC|NOMBRE")
    _PM_CANT = re.compile(r"^CANT|CANTIDAD|CANT\.|^QTY|SOLICITAD|ENTREGAD|DESPACHAD")
    _PM_UNID = re.compile(r"^U\.?M\.?$|^UN\.?$|^UND|UNIDAD|^UNID|^U/M|^MED")
    _PM_TRAB = re.compile(r"ACTIVITY|ACTIVIDAD|\bOT\b|ORDEN DE TRABAJO|TRABAJO|NDC|UDR|DESTINO|USO|APLICACION|AREA")
    _PM_CODIGO = re.compile(r"^(COD|CODIGO|SKU|N° PARTE|NRO PARTE|PART)")
    _UNIDADES_PM = r"UN|UND|UNID|UNIDAD(?:ES)?|C/U|KG|KGS|GR|TON|M|MT|MTS|ML|M2|M3|LT|LTS|L|GL|GAL|ROLLO|ROLLOS|RLL|CJ|CAJA|CAJAS|PZA|PZ|JGO|JUEGO|PAR|BOLSA|BOLSAS|SACO|SACOS|TIRA|TIRAS|PAQ|PQT|PAQUETE|BARRA|BARRAS|PLANCHA|PL|TAMBOR|BIDON|TARRO|GLB|GLOBAL|HR|HRS|SET|KIT|CU"
    
    
    def _num_pm(v) -> float | None:
        s = str(v or "").strip().replace(" ", "")
        if not s:
            return None
        if re.fullmatch(r"\d{1,3}(\.\d{3})+(,\d+)?", s):          # 1.234,5
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None
    
    
    def _cabecera_pm(texto: str) -> dict:
        """Datos generales del PM: número, solicitante, Activity ID / OT / NDC y fecha."""
        t = texto or ""
        info = {}
        m = re.search(r"\b(?:P\.?\s?M\.?|PEDIDO DE MATERIAL(?:ES)?|PARTE DE MATERIAL(?:ES)?|VALE(?: DE CONSUMO| DE MATERIALES)?|"
                      r"SOLICITUD DE MATERIAL(?:ES)?)\s*(?:N\s*[°ºo]\.?|NRO\.?|NUM\.?|#)?\s*[:\-]?\s*([A-Z]{0,4}-?\d[\d\-]{1,14})", t, re.I)
        if m:
            info["pm"] = m.group(1).strip("-")
        m = re.search(r"(?:SOLICITAD[OA] POR|SOLICITANTE|PEDIDO POR|RESPONSABLE|SUPERVISOR)\s*[:\-]?\s*([^\n]{3,60})", t, re.I)
        if m:
            info["solicitado"] = re.split(r"\s{3,}|\b(?:FECHA|CARGO|FIRMA|AREA|ÁREA)\b", m.group(1).strip(), flags=re.I)[0].strip(" :-")
        m = re.search(r"\bSY-\d{4}[A-Z]{3}\d{3,5}\b", t)
        if m:
            info["trabajo"] = m.group(0)
        else:
            m = re.search(r"\b(?:NDC|UDR)\s*(?:N\s*[°ºo]\.?)?\s*[:\-]?\s*(\d{3,5})", t, re.I)
            if m:
                info["trabajo"] = f"NDC {m.group(1)}"
            else:
                m = re.search(r"(?:\bOT\b|ORDEN DE TRABAJO)\s*(?:N\s*[°ºo]\.?)?\s*[:\-]?\s*([A-Z0-9\-]{3,20})", t, re.I)
                if m:
                    info["trabajo"] = f"OT {m.group(1)}"
        m = re.search(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", t)
        if m:
            try:
                a = int(m.group(3)); a = a + 2000 if a < 100 else a
                info["fecha"] = date(a, int(m.group(2)), int(m.group(1)))
            except ValueError:
                pass
        return info
    
    
    def _materiales_de_tabla(tabla: list[list]) -> list[dict]:
        filas = [[re.sub(r"\s+", " ", str(c or "")).strip() for c in f] for f in tabla if f]
        for i, f in enumerate(filas[:8]):
            cab = [norm(c) for c in f]
            c_d = next((j for j, c in enumerate(cab) if _PM_DESC.search(c) and not _PM_CODIGO.search(c)), None)
            c_q = next((j for j, c in enumerate(cab) if _PM_CANT.search(c) and j != c_d), None)
            if c_d is None or c_q is None:
                continue
            c_u = next((j for j, c in enumerate(cab) if _PM_UNID.search(c) and j not in (c_d, c_q)), None)
            c_t = next((j for j, c in enumerate(cab) if _PM_TRAB.search(c) and j not in (c_d, c_q, c_u)), None)
            c_c = next((j for j, c in enumerate(cab) if _PM_CODIGO.search(c) and j not in (c_d, c_q, c_u, c_t)), None)
            out = []
            for g in filas[i + 1:]:
                g = g + [""] * (len(f) - len(g))
                d_, q_ = g[c_d], _num_pm(g[c_q])
                if not d_ or norm(d_).startswith(("TOTAL", "OBSERVAC", "FIRMA")):
                    continue
                if q_ is None:
                    if out and not any(g[j] for j in range(len(g)) if j != c_d):     # descripción que sigue en otra línea
                        out[-1]["descripcion"] += " " + d_
                    continue
                m = {"descripcion": d_, "cantidad": q_, "unidad": g[c_u] if c_u is not None else ""}
                if c_t is not None:                     # la tabla trae su propia columna de trabajo / Activity ID
                    m["trabajo"] = g[c_t]
                if c_c is not None and g[c_c]:
                    m["codigo"] = g[c_c]
                out.append(m)
            if out:
                return out
        return []
    
    
    def _materiales_de_texto(texto: str) -> list[dict]:
        """Respaldo para PDF sin tabla dibujada: líneas «[ítem] descripción cantidad unidad» o «… unidad cantidad»."""
        out = []
        for linea in (texto or "").splitlines():
            l = linea.strip()
            m = re.match(rf"^(?:\d{{1,3}}[.)]?\s+)?(?:[A-Z0-9\-]{{4,}}\s+)?(.{{4,}}?)\s+(\d+(?:[.,]\d+)?)\s+({_UNIDADES_PM})\.?$", l, re.I) \
                or re.match(rf"^(?:\d{{1,3}}[.)]?\s+)?(?:[A-Z0-9\-]{{4,}}\s+)?(.{{4,}}?)\s+({_UNIDADES_PM})\.?\s+(\d+(?:[.,]\d+)?)$", l, re.I)
            if not m:
                continue
            d_, a, b = m.groups()
            q_, u_ = (a, b) if _num_pm(a) is not None else (b, a)
            if re.search(r"FECHA|TOTAL|PAGINA|PÁGINA|RUT|FONO", d_, re.I):
                continue
            out.append({"descripcion": d_.strip(" -:"), "cantidad": _num_pm(q_), "unidad": u_.upper()})
        return out
    
    
    def leer_pm_materiales(datos: bytes, archivo: str) -> ReporteTerreno:
        """PM (pedido / parte de materiales) en PDF → reporte con la lista de materiales para la sección 8.0."""
        textos, mats = [], []
        try:
            import pdfplumber
        except ImportError:
            pdfplumber = None
        if pdfplumber is not None:
            with pdfplumber.open(io.BytesIO(datos)) as pdf:
                for pag in pdf.pages:
                    textos.append(pag.extract_text() or "")
                    tablas = pag.extract_tables() or []
                    if not tablas:
                        tablas = pag.extract_tables({"vertical_strategy": "text", "horizontal_strategy": "text"}) or []
                    for t in tablas:
                        mats += _materiales_de_tabla(t)
        else:                                   # sin pdfplumber: solo texto, con pypdf
            try:
                from pypdf import PdfReader
            except ImportError:
                raise ValueError("falta el componente para leer PDF. Cierra la app y vuelve a abrirla con "
                                 "Iniciar_SYNCORE.bat (lo instala solo), o ejecuta: python -m pip install pdfplumber")
            textos = [p.extract_text() or "" for p in PdfReader(io.BytesIO(datos)).pages]
        texto = "\n".join(textos)
        if not texto.strip():
            raise ValueError("el PDF no tiene texto (parece escaneado o una foto); no se pueden leer los materiales.")
        if not mats:
            mats = _materiales_de_texto(texto)
        if not mats:
            raise ValueError("no se encontró la tabla de materiales (columnas descripción y cantidad).")
        info = _cabecera_pm(texto)
        nombre = re.sub(r"\.pdf$", "", archivo, flags=re.I)
        ref = f"PM {info['pm']}" if info.get("pm") else nombre
        for m in mats:
            m.setdefault("trabajo", info.get("trabajo", ""))
            m["solicitado"] = info.get("solicitado", "")
            m["observacion"] = ref + (f" · cód. {m.pop('codigo')}" if m.get("codigo") else "")
        rep = ReporteTerreno(archivo, info.get("fecha"), info.get("solicitado", ""), "", "Día", "", tipo="pm_materiales")
        rep.materiales = mats
        return rep
    
    
    def nombre_tipo(r) -> str:
        """Rótulo del formato de un reporte (objeto o dict guardado)."""
        precom = r.get("precom") if isinstance(r, dict) else getattr(r, "precom", False)
        tipo = r.get("tipo", "terreno") if isinstance(r, dict) else r.tipo
        return NOMBRE_TIPO["precom"] if precom else NOMBRE_TIPO.get(tipo, tipo)
    
    
    def leer_reportes(datos: bytes, archivo: str) -> tuple[list[ReporteTerreno], list[str]]:
        """Lee TODAS las hojas de un archivo y reconoce cada formato por su contenido.
    
        Devuelve los reportes encontrados (uno por hoja con datos; p. ej. una hoja por capataz) y notas sobre las hojas omitidas.
        """
        import openpyxl
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(io.BytesIO(datos), data_only=True, read_only=True)
        reportes, notas = [], []
        try:
            hojas = [(ws.title, _grilla(ws)) for ws in wb.worksheets if getattr(ws, "sheet_state", "visible") == "visible"]
        finally:
            wb.close()
        utiles = [(t, g, _tipo_hoja(g)) for t, g in hojas]
        con_datos = [(t, g, tipo) for t, g, tipo in utiles if tipo in LECTORES]
        for titulo, grid, tipo in utiles:
            if tipo not in LECTORES:
                if grid:
                    notas.append(f"Hoja «{titulo.strip()}»: " + ("nómina de personal, se omite." if tipo == "nomina"
                                                                else "no tiene un formato reconocible, se omite."))
                continue
            nombre = archivo if len(con_datos) == 1 else f"{archivo} › {titulo.strip()}"
            if tipo == "uso_recurso":
                # varios formularios UDR uno debajo del otro en la misma hoja: cada uno es un uso de recurso
                inicios = [k for k, f in enumerate(grid) if any("USO DE RECURSO" in norm(x) for x in f if isinstance(x, str))]
                inicios = [k for n_, k in enumerate(inicios) if n_ == 0 or k - inicios[n_ - 1] > 8]
                if len(inicios) > 1:
                    bloques = [grid[max(0, a - 4):(b - 4 if b else None)] for a, b in zip(inicios, inicios[1:] + [None])]
                    for n_, g_ in enumerate(bloques, 1):
                        try:
                            r_ = LECTORES[tipo](g_, f"{nombre} › UDR {n_}")
                        except Exception as e:
                            notas.append(f"Hoja «{titulo.strip()}», UDR {n_}: no se pudo leer ({e}).")
                            continue
                        if r_.uso_recurso:
                            reportes.append(r_)
                    continue
            try:
                rep = LECTORES[tipo](grid, nombre)
            except Exception as e:                                   # una hoja con problemas no bloquea las demás
                notas.append(f"Hoja «{titulo.strip()}»: no se pudo leer ({e}).")
                continue
            if tipo == "terreno" and (not rep.trabajadores or not any(t.total for t in rep.trabajadores)) and len(con_datos) > 1:
                notas.append(f"Hoja «{titulo.strip()}»: reporte sin horas registradas (copia o plantilla), se omite.")
                continue
            if tipo == "terreno" and not rep.trabajadores and not rep.actividades:
                notas.append(f"Hoja «{titulo.strip()}»: formato de reporte de terreno, pero sin personal ni actividades.")
                continue
            if tipo in ("adicionales", "programacion") and not rep.uso_recurso and not rep.actividades:
                notas.append(f"Hoja «{titulo.strip()}»: {NOMBRE_TIPO[tipo].lower()} sin registros, se omite.")
                continue
            reportes.append(rep)
        # Una hoja con otra fecha que el resto del archivo suele ser una copia del día anterior
        fechas = [r.fecha for r in reportes if r.fecha]
        if len(set(fechas)) > 1:
            del_nombre = _fecha_en([], archivo)
            comun = del_nombre if del_nombre in fechas else max(set(fechas), key=fechas.count)
            if del_nombre in fechas or fechas.count(comun) > len(fechas) / 2:
                for r in [r for r in reportes if r.fecha and r.fecha != comun and r.tipo == "uso_recurso"]:
                    # un formulario UDR con otra fecha es uno ya informado que quedó en el archivo: no se vuelve a contar
                    notas.append(f"Hoja «{r.archivo.split(' › ')[-1]}»: formulario con fecha {r.fecha:%d-%m-%Y} (el archivo es del "
                                 f"{comun:%d-%m-%Y}); se omite por ser de otro día.")
                    reportes.remove(r)
                for r in [r for r in reportes if r.fecha and r.fecha != comun]:
                    r.avisos.append(f"La hoja indica fecha {r.fecha:%d-%m-%Y} y el resto del archivo {comun:%d-%m-%Y}; "
                                    f"se asume {comun:%d-%m-%Y} (fecha no actualizada en la hoja).")
                    r.fecha = comun
        if not reportes:
            raise ValueError("no se reconoció ningún formato con datos (reporte de terreno, uso de recurso, adicionales o "
                             "programación). Hojas: " + ", ".join(f"«{t.strip()}»" for t, _ in hojas))
        return reportes, notas
    
    
    # ====================================================================== tarja de asistencia
    
    @dataclass
    class Marca:
        rut: str
        nombre: str
        estado: str            # T, TN, TT, TC, F, D, L, V, PCG, PSG; «P» = presente sin código (se decide día/noche)
        noche: bool | None     # None: la tarja no dice el turno
        texto: str             # lo que venía en la tarja (para mostrarlo)
        acreditado: str = ""   # columna ACREDITADO de la tarja (SI / NO / N/A), si la trae
        ficha: dict = field(default_factory=dict)   # datos personales de la tarja en columnas de BD Personal (E..AA)
    
    
    # Códigos de asistencia más comunes → estado de BD Personal. «P» = asistió (el turno se decide después).
    _CODIGOS_TARJA = {
        "T": "T", "TD": "T", "TN": "TN", "N": "TN", "TT": "TT", "TC": "TC", "P": "P", "X": "P", "A": "P", "AS": "P", "1": "P",
        "SI": "P", "S": "P", "PRESENTE": "P", "ASISTE": "P", "ASISTIO": "P", "TRABAJA": "P", "OK": "P", "✓": "P", "V/": "P",
        "F": "F", "FA": "F", "FALLA": "F", "FALTA": "F", "AUSENTE": "F", "INASISTENCIA": "F", "0": "F", "NO": "F", "NS": "F",
        "D": "D", "DE": "D", "DESCANSO": "D", "LIBRE": "D", "BAJADA": "D", "L": "L", "LM": "L", "LICENCIA": "L",
        "V": "V", "VAC": "V", "VACACIONES": "V", "VACACION": "V", "PCG": "PCG", "PERMISO": "PCG", "PSG": "PSG",
        "AC": "AC", "ACRED": "AC", "EN ACREDITACION": "AC",
    }
    
    
    def estado_tarja(v) -> tuple[str | None, bool | None]:
        """(estado, noche) a partir de lo que trae una celda de la tarja; (None, None) si está vacía o no se entiende."""
        if v is None or v == "":
            return None, None
        if isinstance(v, (datetime, time)) or (isinstance(v, str) and re.fullmatch(r"\s*\d{1,2}:\d{2}(:\d{2})?\s*", v)):
            h = v.hour if isinstance(v, (datetime, time)) else int(v.strip().split(":")[0])
            return "P", h >= 17 or h < 4                      # marcó entrada: asistió (entrada de tarde/noche = turno noche)
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            if 0 < v < 1:                                     # hora guardada como fracción del día
                return "P", v * 24 >= 17 or v * 24 < 4
            return ("P", None) if v > 0 else ("F", None)
        t = norm(v)
        if t in _CODIGOS_TARJA:
            e = _CODIGOS_TARJA[t]
            return e, (True if e == "TN" else None)
        noche = True if "NOCHE" in t else (False if re.search(r"\bDIA\b", t) else None)
        for clave, e in (("TELETRAB", "TT"), ("CURSO", "TC"), ("SINDICAL", "TC"), ("LICENCIA", "L"), ("VACACION", "V"),
                         ("PERMISO SIN", "PSG"), ("PERMISO", "PCG"), ("FALTA", "F"), ("FALLA", "F"), ("AUSEN", "F"),
                         ("INASIST", "F"), ("NO ASIST", "F"), ("NO SUBIO", "F"), ("DESCANSO", "D"), ("BAJADA", "D"),
                         ("PRESENT", "P"), ("ASIST", "P"), ("TRABAJ", "P"), ("NOCHE", "P"), ("DIA", "P")):
            if clave in t:
                return e, noche
        return None, None
    
    
    def _rut_valido(v) -> str | None:
        r = norm_rut(v)
        return r if re.fullmatch(r"\d{6,9}-?[\dK]", r) else None
    
    
    def _con_guion(r: str) -> str:
        return r if "-" in r else f"{r[:-1]}-{r[-1]}"
    
    
    def _celda_fecha(v, fecha: date) -> bool:
        """¿La celda de encabezado corresponde al día pedido? (fecha completa o número de día)."""
        if isinstance(v, datetime):
            return v.date() == fecha
        if isinstance(v, date):
            return v == fecha
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return float(v).is_integer() and int(v) == fecha.day
        t = _txt(v)
        m = re.search(r"(\d{1,2})[-/.](\d{1,2})(?:[-/.](\d{2,4}))?", t)
        if m:
            return int(m.group(1)) == fecha.day and int(m.group(2)) == fecha.month
        m = re.fullmatch(r"(?:[A-Za-zÁÉÍÓÚáéíóú]{1,3}\.?\s*)?(\d{1,2})", t)
        return bool(m) and int(m.group(1)) == fecha.day
    
    
    def _es_dia(v) -> bool:
        if isinstance(v, (datetime, date)):
            return True
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return float(v).is_integer() and 1 <= v <= 31
        return bool(re.fullmatch(r"(?:[A-Za-zÁÉÍÓÚáéíóú]{1,3}\.?\s*)?\d{1,2}([-/.]\d{1,2}([-/.]\d{2,4})?)?", _txt(v)))
    
    
    # Encabezado de la tarja -> columna de BD Personal (mismo maestro de personal)
    _TARJA_A_BD = (("ROL", "E"), ("TURNO", "F"), ("VIGENTE", "G"), ("V - F", "G"), ("TIPO DE TURNO", "H"), ("NOMBRE", "I"),
                   ("CARGO", "J"), ("ESPECIALIDAD", "K"), ("AREA", "L"), ("FECHA DE NACIMIENTO", "M"), ("COMUNA", "N"),
                   ("REGION", "O"), ("LOCAL", "P"), ("DIRECCION", "Q"), ("TELEFONO", "R"), ("FECHA CONTRATACION", "S"),
                   ("FECHA DE CONTRATACION", "S"), ("FECHA FINIQUITO", "T"), ("FECHA DE FINIQUITO", "T"), ("CAUSAL", "U"),
                   ("EMPRESA", "V"), ("MOI/MOD", "W"), ("CORREO", "X"), ("GENERO", "Y"), ("ESTATUS", "Z"), ("ACREDITADO", "AA"))
    
    
    def _cols_ficha(nombres: list[str]) -> dict[int, str]:
        out = {}
        for j, n in enumerate(nombres):
            if not n:
                continue
            for pref, col in _TARJA_A_BD:
                if (n == pref or n.startswith(pref + " ") or n.startswith(pref + "/") or n.startswith(pref)) and col not in out.values():
                    if pref == "TURNO" and n != "TURNO":
                        continue                               # «TURNO» exacto (14X14 A); «TIPO DE TURNO» va aparte
                    out[j] = col
                    break
        return out
    
    
    def _leer_tarja_grilla(grid, fecha: date, origen: str) -> tuple[dict[str, Marca], list[str]]:
        notas: list[str] = []
        cab = None
        for i, f in enumerate(grid[:60]):
            nombres = [norm(v) for v in f]
            if any(n == "RUT" or n.startswith("RUT ") or n in ("R.U.T", "R.U.T.", "RUN", "RUT TRABAJADOR") for n in nombres):
                cab = i, nombres
                break
        filas = grid
        if cab is None:                                       # sin encabezado: se busca la columna con RUT
            conteo: dict[int, int] = {}
            for f in grid[:200]:
                for j, v in enumerate(f):
                    if _rut_valido(v):
                        conteo[j] = conteo.get(j, 0) + 1
            if not conteo or max(conteo.values()) < 3:
                return {}, []
            j_rut = max(conteo, key=conteo.get)
            i0, nombres = -1, []
        else:
            i0, nombres = cab
            j_rut = next(j for j, n in enumerate(nombres) if n == "RUT" or n.startswith("RUT ") or n in ("R.U.T", "R.U.T.", "RUN", "RUT TRABAJADOR"))
        j_nom = next((j for j, n in enumerate(nombres) if "NOMBRE" in n or n in ("TRABAJADOR", "APELLIDOS")), None)
        cols_ficha = _cols_ficha(nombres) if len(_cols_ficha(nombres)) >= 6 else {}
        j_acr = next((j for j, n in enumerate(nombres) if n in ("ACREDITADO", "ACREDITADA", "ACREDITACION", "ACREDITADO SI/NO")), None)
    
        # ¿Formato mensual (una columna por día)? Se busca en el encabezado y en las 2 filas de arriba.
        j_dia = None
        for k in range(max(0, i0 - 2), i0 + 1) if i0 >= 0 else []:
            fila = grid[k]
            dias = [j for j, v in enumerate(fila) if j != j_rut and _es_dia(v)]
            if len(dias) >= 5:
                cand = [j for j in dias if _celda_fecha(fila[j], fecha)]
                if cand:
                    j_dia = cand[0]
                    notas.append(f"{origen}: formato mensual, se usó la columna del día {fecha:%d-%m}.")
                else:
                    notas.append(f"{origen}: tiene columnas por día pero ninguna corresponde al {fecha:%d-%m-%Y}.")
                    return {}, notas
                break
    
        j_fecha = j_est = j_ent = j_jor = None
        if j_dia is None:
            def buscar(*claves, excluir=()):
                for j, n in enumerate(nombres):
                    if j in (j_rut, j_nom):
                        continue
                    if any(c in n for c in claves) and not any(e in n for e in excluir):
                        return j
                return None
            j_fecha = buscar("FECHA", excluir=("NACIM", "INGRESO", "CONTRAT", "TERMINO"))
            j_est = buscar("ESTADO", "ASISTENCIA", "ASIST", "MARCA", "CODIGO", "PRESENTE", "CONDICION", "SITUACION", "OBSERV")
            j_ent = buscar("ENTRADA", "INGRESO", "HORA")
            j_jor = buscar("JORNADA", "DIA/NOCHE", "DIA / NOCHE", "TIPO DE TURNO", "TURNO D", "TURNO")
            if j_jor is not None and j_jor == j_est:
                j_jor = None
    
        marcas: dict[str, Marca] = {}
        otras_fechas = 0
        for f in filas[i0 + 1:]:
            if j_rut >= len(f):
                continue
            rut = _rut_valido(f[j_rut])
            if not rut:
                continue
            rut = _con_guion(rut)
            nombre = _txt(f[j_nom]) if j_nom is not None and j_nom < len(f) else ""
            val = lambda j: f[j] if j is not None and j < len(f) else None
            if j_fecha is not None and val(j_fecha) not in (None, ""):
                fv = val(j_fecha)
                fv = fv.date() if isinstance(fv, datetime) else fv
                if isinstance(fv, date) and fv != fecha or (not isinstance(fv, date) and not _celda_fecha(fv, fecha)):
                    otras_fechas += 1
                    continue
            if j_dia is not None:
                est, noche = estado_tarja(val(j_dia)); texto = _txt(val(j_dia))
            elif j_est is not None and val(j_est) not in (None, ""):
                est, noche = estado_tarja(val(j_est)); texto = _txt(val(j_est))
            elif j_ent is not None:
                est, noche = estado_tarja(val(j_ent)); texto = _txt(val(j_ent))
                est = est or "F"
            else:
                est, noche, texto = "P", None, "en la lista"        # lista simple de asistentes
            ficha = {col: (val(j).date().isoformat() if isinstance(val(j), datetime) else
                           val(j).isoformat() if isinstance(val(j), date) else val(j))
                     for j, col in cols_ficha.items() if val(j) not in (None, "")}      # (fechas en texto ISO: se guarda en JSON)
            if est is None:
                if ficha:                                      # sin marca del día, pero con su ficha (vigente / finiquitado)
                    marcas.setdefault(rut, Marca(rut, nombre, "", None, "", "", ficha))
                continue
            if j_jor is not None and val(j_jor) not in (None, ""):
                jt = norm(val(j_jor))
                if "NOCHE" in jt or jt in ("N", "TN"):
                    noche = True
                elif "DIA" in jt or jt in ("D", "TD"):
                    noche = False if noche is None else noche
            previo = marcas.get(rut)
            if previo and previo.estado in ("P", "T", "TN", "TT", "TC") and est not in ("P", "T", "TN", "TT", "TC"):
                continue                                       # dos marcas del mismo día: manda la de asistencia
            acr = norm(val(j_acr)) if j_acr is not None else ""
            marcas[rut] = Marca(rut, nombre, est, noche, texto, acr if acr not in ("", "#N/A", "0") else "", ficha)
        if otras_fechas:
            notas.append(f"{origen}: {otras_fechas} fila(s) de otras fechas no se consideraron.")
        return marcas, notas
    
    
    def _grilla_tarja(ws, fecha: date, max_filas: int = 6000) -> list[list]:
        """Como _grilla, pero con todas las columnas necesarias: en la tarja mensual/anual hay una columna por día y la del
        informe puede estar más allá de la columna 300. Se ubica la columna de la fecha en el encabezado y se corta ahí."""
        tope = 70
        for fila in ws.iter_rows(max_row=40, max_col=800, values_only=True):
            for j, v in enumerate(fila):
                if v is not None and _celda_fecha(v, fecha) and isinstance(v, (datetime, date)):
                    tope = max(tope, j + 2)
        grid, vacias = [], 0
        for fila in ws.iter_rows(max_row=max_filas, max_col=tope, values_only=True):
            fila = list(fila)
            if any(v not in (None, "") for v in fila):
                grid.extend([[]] * vacias)
                vacias = 0
                grid.append(fila)
            else:
                vacias += 1
                if vacias > 60:
                    break
        return grid
    
    
    # ---------------------------------------------------------------------- tarja de maquinaria
    
    def estado_maquina(v) -> dict:
        """Código o texto de la tarja de maquinaria -> cambios de BD Maquinaria (estado / observación / disponible)."""
        if v is None or v == "":
            return {}
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return {"estado": "OPERANDO", "disponible": "SI"} if v > 0 else {"estado": "STAND BY"}
        t = norm(v)
        if t in ("T", "TD", "D/T", "X", "1", "SI", "OP", "O", "OPERANDO", "OPERATIVO", "OPERATIVA", "TRABAJANDO", "DIA"):
            return {"estado": "OPERANDO", "disponible": "SI", "observacion": ""}
        if t in ("TN", "N", "NOCHE") or t.startswith("TURNO NOCHE"):
            return {"estado": "OPERANDO", "disponible": "SI", "observacion": "Turno Noche"}
        if t in ("TDN", "DN", "T/N", "TD/TN", "24H", "24 H") or ("DIA" in t and "NOCHE" in t):
            return {"estado": "OPERANDO", "disponible": "SI", "observacion": "Turno Dia y Noche"}
        if t in ("M", "MA", "MANT") or t.startswith("MANTEN"):
            return {"estado": "MANTENCION"}
        if t in ("P", "PA") or t.startswith("PANNE") or t.startswith("FALLA") or t.startswith("AVERIA"):
            return {"estado": "PANNE"}
        if t in ("SB", "S", "STB", "STAND BY", "STANDBY", "DETENIDO", "SIN OPERADOR", "0", "NO") or t.startswith("STAND"):
            return {"estado": "STAND BY"}
        if t.startswith("DESMOV") or t in ("DM", "DES"):
            return {"estado": "DESMOVILIZADO"}
        if t in ("ND", "NO DISPONIBLE"):
            return {"disponible": "NO"}
        return {}
    
    
    def leer_tarja_maquinaria(datos: bytes, archivo: str, fecha: date) -> tuple[dict[str, dict], list[str]]:
        """Tarja / reporte de maquinaria en formato libre: una fila por equipo con su PATENTE y
        - una columna por día (planilla mensual) con el código del día (T, TN, M, P, SB, D…), o
        - columnas ESTADO / OPERANDO / DISPONIBLE / TURNO / OBSERVACIÓN / HORÓMETRO (lista del día, o con columna FECHA).
        Devuelve {patente normalizada: {campo: valor}} con los cambios para BD Maquinaria."""
        import openpyxl
        import warnings
        grillas = []
        if archivo.lower().endswith((".csv", ".txt")):
            import csv
            texto = datos.decode("utf-8-sig", errors="replace")
            sep = ";" if texto.count(";") > texto.count(",") else ","
            grillas.append((archivo, [list(r) for r in csv.reader(io.StringIO(texto), delimiter=sep)]))
        else:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                wb = openpyxl.load_workbook(io.BytesIO(datos), read_only=True, data_only=True)
                for ws in [w for w in wb.worksheets if w.sheet_state == "visible"] or wb.worksheets:
                    grillas.append((f"{archivo} › {ws.title}", _grilla_tarja(ws, fecha)))
            wb.close()
        salida: dict[str, dict] = {}
        notas: list[str] = []
        for origen, g in grillas:
            cab = next(((i, [norm(v) for v in f]) for i, f in enumerate(g[:40])
                        if any(n in ("PATENTE", "PPU", "PLACA", "PATENTE / CODIGO", "N° PATENTE") or n.startswith("PATENTE") for n in
                               (norm(x) for x in f))), None)
            if not cab:
                continue
            i0, nombres = cab
            jp = next(j for j, n in enumerate(nombres) if n in ("PPU", "PLACA") or n.startswith("PATENTE") or n == "N° PATENTE")
            jdia = next((j for j, v in enumerate(g[i0]) if j != jp and isinstance(v, (datetime, date)) and _celda_fecha(v, fecha)), None)
            if jdia is None and sum(1 for v in g[i0] if _es_dia(v)) >= 5:
                jdia = next((j for j, v in enumerate(g[i0]) if j != jp and _es_dia(v) and _celda_fecha(v, fecha)), None)
            def col(*claves, excluir=()):
                return next((j for j, n in enumerate(nombres) if j != jp and any(c in n for c in claves)
                             and not any(e in n for e in excluir)), None)
            j_est = col("ESTADO", "CONDICION", "SITUACION", "OPERANDO", "OPERATIV", "STATUS")
            j_disp = col("DISPONIBLE")
            j_turno = col("TURNO", "JORNADA")
            j_obs = col("OBSERVACION", "COMENTARIO")
            j_hor = col("HOROMETRO", "HM", "KILOMETRAJE")
            j_fecha = col("FECHA", excluir=("FABRIC", "ANO", "ADQUI"))
            j_tipo = col("TIPO", "EQUIPO", "DESCRIPCION")
            n_ok = 0
            for f in g[i0 + 1:]:
                f = list(f) + [None] * 5
                pat = _clave_patente(f[jp])
                if len(pat) < 4:
                    continue
                if j_fecha is not None and f[j_fecha] not in (None, "") and jdia is None:
                    fv = f[j_fecha].date() if isinstance(f[j_fecha], datetime) else f[j_fecha]
                    if not ((isinstance(fv, date) and fv == fecha) or (not isinstance(fv, date) and _celda_fecha(fv, fecha))):
                        continue
                c: dict = {}
                if jdia is not None:
                    c.update(estado_maquina(f[jdia]))
                else:
                    if j_est is not None:
                        c.update(estado_maquina(f[j_est]))
                    if j_disp is not None and norm(f[j_disp]) in ("SI", "NO"):
                        c["disponible"] = norm(f[j_disp])
                    if j_turno is not None and _txt(f[j_turno]):
                        tt = norm(f[j_turno])
                        if "NOCHE" in tt and "DIA" in tt:
                            c["observacion"] = "Turno Dia y Noche"
                        elif "NOCHE" in tt or tt == "TN":
                            c["observacion"] = "Turno Noche"
                    if j_obs is not None and _txt(f[j_obs]) and "observacion" not in c:
                        c["observacion"] = _txt(f[j_obs])
                if j_hor is not None and _num(f[j_hor]):
                    h_ = _num(f[j_hor])
                    c["horometro"] = str(int(h_)) if float(h_).is_integer() else str(h_)
                if c:
                    c["_tipo"] = _txt(f[j_tipo]) if j_tipo is not None else ""
                    c["_patente"] = _txt(f[jp])
                    salida[pat] = {**salida.get(pat, {}), **c}
                    n_ok += 1
            if n_ok:
                notas.append(f"{origen}: {n_ok} equipo(s)" + (" (columna del día)" if jdia is not None else "") + ".")
        if not salida:
            raise ValueError(f"No se encontró una columna PATENTE con el estado de los equipos para el {fecha:%d-%m-%Y}.")
        return salida, notas
    
    
    def _vigente(m: "Marca") -> bool:
        f = m.ficha or {}
        g, z = norm(f.get("G")), norm(f.get("Z"))
        return (g.startswith("VIGENTE") or z.startswith("CONTRATADO")) and not g.startswith("FINIQUITADO")
    
    
    def _finiquitado(m: "Marca") -> bool:
        f = m.ficha or {}
        return norm(f.get("G")).startswith("FINIQUITADO") or norm(f.get("Z")).startswith("FINIQUITADO")
    
    
    def leer_tarja(datos: bytes, archivo: str, fecha: date) -> tuple[dict[str, Marca], list[str]]:
        """Lee una tarja de asistencia en cualquiera de los formatos habituales:
        - lista del día: RUT (+ nombre) + estado/asistencia/marca, o + hora de entrada, o solo la lista de los que asistieron;
        - lista de varios días con columna FECHA (se toma la del informe);
        - planilla mensual con una columna por día (1..31 o fechas).
        Acepta .xlsx/.xlsm y .csv. Devuelve {rut: Marca} y notas de lectura."""
        nombre = archivo.lower()
        grillas: list[tuple[str, list[list]]] = []
        if nombre.endswith((".csv", ".txt")):
            import csv
            texto = None
            for cod in ("utf-8-sig", "latin-1"):
                try:
                    texto = datos.decode(cod)
                    break
                except UnicodeDecodeError:
                    pass
            sep = ";" if texto.count(";") > texto.count(",") else ","
            grillas.append((archivo, [list(r) for r in csv.reader(io.StringIO(texto), delimiter=sep)]))
        else:
            import openpyxl
            import warnings
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                wb = openpyxl.load_workbook(io.BytesIO(datos), read_only=True, data_only=True)
                hojas = [ws for ws in wb.worksheets if ws.sheet_state == "visible"] or wb.worksheets
                preferidas = [ws for ws in hojas if "TARJA" in norm(ws.title) or "ASIST" in norm(ws.title)]
                for ws in preferidas or hojas:
                    grillas.append((f"{archivo} › {ws.title}", _grilla_tarja(ws, fecha)))
            wb.close()
        marcas: dict[str, Marca] = {}
        notas: list[str] = []
        for origen, g in grillas:
            m, n = _leer_tarja_grilla(g, fecha, origen)
            notas += n
            for rut, mk in m.items():
                if rut not in marcas or (marcas[rut].estado not in ("P", "T", "TN", "TT", "TC")):
                    marcas[rut] = mk
        if marcas and not any(m.estado in ("P", "T", "TN", "TT", "TC") for m in marcas.values()):
            notas.append(f"Ojo: para el {fecha:%d-%m-%Y} la tarja aún no tiene a nadie marcado como trabajando "
                         "(solo descansos/ausencias). Revisa que sea la versión actualizada del día.")
        if not marcas:
            raise ValueError("No se encontró una columna de RUT con datos de asistencia para "
                             f"el {fecha:%d-%m-%Y}. Revisa que el archivo tenga RUT y la marca del día.")
        return marcas, notas
    
    
    def leer_reporte_terreno(datos: bytes, archivo: str) -> ReporteTerreno:
        """Compatibilidad: el primer reporte del archivo."""
        return leer_reportes(datos, archivo)[0][0]
    
    
    # ====================================================================== planilla oficial
    
    HOJA_INFORME = "Informe Diario"
    HOJA_PERSONAL = "BD Personal"
    HOJA_ACTIVIDADES = "BD Actividades MOD"
    HOJA_MASTER = "Master Actividades"          # formato nuevo (T&M): una fila por persona y actividad, con Activity ID de P6
    HOJA_P6 = "P6"
    CRITERIO_DEFECTO = "Planificada (Itemizado / Programa)"
    
    # Columnas de BD Personal (encabezado en la fila 3 de la planilla)
    COL_P = {"mes": "B", "fecha": "C", "rut": "D", "rol": "E", "turno": "F", "vf": "G", "tipo_turno": "H", "nombre": "I",
             "cargo": "J", "especialidad": "K", "area": "L", "empresa": "V", "moimod": "W", "estatus": "Z",
             "acreditado": "AA", "faena": "AB"}
    # En BD Actividades MOD las columnas B..AA son las mismas de BD Personal C..AB (una a la izquierda)
    COL_A_EXTRA = {"supervisor": "AB", "capataz": "AC", "area_trabajo": "AD", "sub_area": "AE", "actividad": "AF", "hh": "AG"}
    
    
    @dataclass
    class Persona:
        fila: int
        rut: str
        nombre: str
        cargo: str
        especialidad: str
        tipo_turno: str
        moimod: str
        acreditado: str
        estatus: str
        faena: str
        turno: str = ""                  # sistema de turno (columna TURNO de BD Personal): «14X14 A», «15X13 B», «8X6 A»…
        ficha: dict = field(default_factory=dict)   # datos de la tarja para una persona nueva (columna de BD Personal -> valor)
    
    
    @dataclass
    class FilaCargo:
        fila: int
        cargo: str
        seccion: str                     # "MOD" o "MOI"
        formulas: dict[str, str]         # {"L": "...", "O": ..., "R": ..., "U": ...}
    
    
    @dataclass
    class FilaEquipo:
        fila: int
        nombre: str
        grupo: str                       # OOCC, MONTAJE, INDIRECTOS
        cantidad: float | None
        modelo: str
        patente: str
        dia: float
        noche: float
        hm_anterior: float | None
        observacion: str
        col_noche: str
    
    
    @dataclass
    class Plantilla:
        datos: bytes
        fecha: date
        numero: int
        condicion: str
        personas: list[Persona]
        cargos: list[FilaCargo]
        equipos: list[FilaEquipo]
        valores: dict[str, object]       # celdas manuales del informe (números, textos)
        refs: dict[str, str]             # rótulo -> celda donde va el valor
        listas: dict[str, list[str]]
        filas_interferencia: list[int]
        filas_udr: list[int]
        filas_persona_xml: dict[int, str]
        estilos_actividades: dict[str, str]
        actividades_texto: str
        fotos: list[dict]                # [{"media": "xl/media/image5.png", "rel": "rId5", "leyenda": "B325", "aspecto": 1.4}]
        totales: dict[str, float] = field(default_factory=dict)
        ultimo_ndc: int | None = None    # último N° NDC del «Uso de recurso» de la planilla base (para seguir la numeración)
        udr_formulas: dict[str, str] = field(default_factory=dict)   # fórmulas de la 1ª fila de «Uso de recurso» (si usan BD UDR)
        # Historial de BD Personal por sistema de turno: {«14X14 A»: {fecha: (personas trabajando, personas)}}
        historial_turnos: dict[str, dict[date, tuple[int, int]]] = field(default_factory=dict)
        # ---- formato nuevo (Informe Diario T&M)
        formato: str = "clasico"                     # "clasico" (BD Actividades MOD) o "master" (Master Actividades + P6)
        col_hh: str = "X"                            # columna «HH Faena» de las tablas de personal (X antiguo, Z nuevo)
        con_formula: set = field(default_factory=set)        # celdas del informe que son fórmula (no se sobrescriben)
        criterios: list = field(default_factory=list)        # tipos de actividad (CRITERIO) del control T&M
        filas_tipo_actividad: dict = field(default_factory=dict)   # fila de la sección 1.1 -> criterio
        p6: dict = field(default_factory=dict)               # Activity ID -> {"C": nombre, "K": área, "I": subsistema, ...}
        master_formulas: dict = field(default_factory=dict)  # columna -> fórmula de la última fila de Master Actividades
        master_fila_modelo: int = 0
        udr_cols: dict = field(default_factory=dict)         # BD UDR: clave -> columna (según encabezados)
        udr_fila_cab: int = 0
        cab_interferencia: dict = field(default_factory=dict)  # columna -> rótulo del encabezado de la tabla de interferencias
        validacion: dict = field(default_factory=dict)       # {"filas": {"Elaborado por": 362, ...}, "cols": {"nombre": "L", ...}}
        historial_master: dict = field(default_factory=dict)  # comentario normalizado -> (Activity ID, criterio) de días anteriores
        maquinaria: list = field(default_factory=list)       # BD Maquinaria del último día: una fila por patente
        materiales: dict = field(default_factory=dict)       # {"filas": [...], "cols": {"descripcion": "B", ...}}
    
    
    def _celdas_fila(L: X.Libro, fila_xml: str) -> dict[str, object]:
        return {c: L.valor_celda(a, i) for c, (a, i, _) in X.celdas(fila_xml).items()}
    
    
    def _formula(interior: str | None) -> str | None:
        if not interior:
            return None
        m = re.search(r"<f\b[^>]*>(.*?)</f>", interior, flags=re.S)
        return X._desescapar(m.group(1)) if m else None
    
    
    _MAQ_COLS = {"patente": "C", "modelo": "D", "tipo": "E", "anio": "F", "contratado": "G", "acreditado": "H", "disponible": "I",
                 "estado": "J", "operador": "K", "observacion": "L", "empresa": "M", "turno": "N", "horometro": "Q"}
    
    
    def _leer_maquinaria(L: "X.Libro", fecha) -> list[dict]:
        """Filas de BD Maquinaria del día del informe base (o del último día con datos)."""
        x = L.hoja_xml("BD Maquinaria")
        filas = []
        for n, a, fx, _, _ in X.filas(x):
            if n < 2:
                continue
            v_ = _celdas_fila(L, fx)
            if isinstance(v_.get("B"), (int, float)) and v_.get("C"):
                filas.append((int(v_["B"]), n, fx, v_))
        if not filas:
            return []
        obj = X.fecha_a_serial(fecha) if fecha else None
        dia = obj if obj in {f[0] for f in filas} else max(f[0] for f in filas)
        out = []
        for f_, n, fx, v_ in filas:
            if f_ == dia:
                d = {k: (str(v_.get(c)).strip() if v_.get(c) is not None else "") for k, c in _MAQ_COLS.items()}
                d["_fila"], d["_xml"] = n, fx
                out.append(d)
        return out
    
    
    def _clave_patente(p) -> str:
        return re.sub(r"[^A-Z0-9]", "", str(p or "").upper())
    
    
    def maquinaria_del_dia(pl: "Plantilla", cambios: dict | None = None) -> list[dict]:
        """BD Maquinaria del nuevo día: la del día anterior con los cambios de estado indicados (por patente).
        Los equipos marcados DESMOVILIZADO salen de la base."""
        cambios = cambios or {}
        out = []
        for d in pl.maquinaria:
            c = cambios.get(_clave_patente(d["patente"]), {})
            nd = {**d, **{k: v for k, v in c.items() if k in _MAQ_COLS}}
            if norm(nd.get("estado")).startswith("DESMOVILIZ"):
                continue
            if norm(nd.get("estado")).startswith(("MANTEN", "PANNE")):
                nd["disponible"] = "SI"                        # sigue disponible para el contrato
            out.append(nd)
        return out
    
    
    # Estados que cuentan como equipo disponible y trabajado (11 HM en su turno): en mantención o panne el equipo sigue
    # contratado en obra (criterio del contrato); solo DESMOVILIZADO lo saca.
    ESTADOS_EQUIPO_CUENTA = ("OPERANDO", "MANTENCION", "PANNE")
    
    
    def cuenta_equipo(d: dict) -> bool:
        e = norm(d.get("estado"))
        if e in ("MANTENCION", "PANNE") or e.startswith("MANTEN"):
            return True
        return norm(d.get("disponible")) == "SI" and e == "OPERANDO"
    
    
    def operando_dia(d: dict) -> bool:
        obs = norm(d.get("observacion"))
        return cuenta_equipo(d) and not ("NOCHE" in obs and "DIA" not in obs)
    
    
    def operando_noche(d: dict) -> bool:
        return cuenta_equipo(d) and norm(d.get("turno")) == "NOCHE"
    
    
    def _leer_historial_master(L: "X.Libro", xa: str, dias: int = 45) -> dict:
        """Master Actividades de los días anteriores: descripción (comentario) -> (Activity ID, criterio) más reciente.
        Es lo que el equipo ya clasificó a mano; sirve para repetir ese criterio en las actividades iguales o parecidas."""
        filas = []
        for n, a, fx, _, _ in X.filas(xa):
            if n < 14:
                continue
            v_ = _celdas_fila(L, fx)
            f_, com = v_.get("B"), v_.get("S")
            if isinstance(f_, (int, float)) and isinstance(com, str) and com.strip():
                filas.append((f_, com, v_.get("C"), v_.get("G")))
        if not filas:
            return {}
        ultimo = max(f for f, *_ in filas)
        hist: dict = {}
        for f_, com, aid, crit in sorted(filas, key=lambda x: x[0]):
            if f_ >= ultimo - dias and isinstance(crit, str) and crit.strip():
                hist[norm(com)[:200]] = (str(aid).strip() if aid not in (None, "", 0) else "", crit.strip())
        return hist
    
    
    def _clave_udr(t: str, vistos: set) -> str | None:
        """Encabezado de BD UDR -> clave del uso de recurso."""
        if t == "FECHA":
            return "fecha"
        if "NDC" in t:
            return "ndc"
        if t.startswith("TITULO") or t.startswith("DESCRIPCION ACTIVIDAD"):
            return "descripcion"
        if t.startswith("DESCRIPCION"):
            return "detalle" if "detalle" not in vistos else "materiales"
        if t.startswith("DCN") or "RFI" in t or "RESPALDO" in t:
            return "respaldo"
        if t in ("AREA", "AREA / FRENTE"):
            return "area"
        if t.startswith("CRITERIO"):
            return "criterio"
        if t == "DIRECTOS":
            return "directos"
        if t == "INDIRECTOS":
            return "indirectos"
        if t.startswith("DISCIPLINA") or t.startswith("ESPECIALIDAD"):
            return "especialidad"
        if t == "HORAS":
            return "horas"
        if t in ("HH", "HH DIRECTAS"):
            return "hh_directos"
        if t == "HH INDIRECTAS":
            return "hh_indirectos"
        if t.startswith("IMPACTO"):
            return "impacto"
        if t.startswith("EQUIPOS"):
            return "equipos"
        if t == "HM":
            return "hm"
        if t.startswith("COMENTARIO"):
            return "comentario"
        return None
    
    
    def leer_plantilla(datos: bytes) -> Plantilla:
        L = X.Libro(datos)
        for h in (HOJA_INFORME, HOJA_PERSONAL):
            if h not in L.hojas:
                raise ValueError(f"La planilla no tiene la hoja '{h}'. ¿Es el informe diario oficial del contrato?")
        formato = "master" if HOJA_MASTER in L.hojas else "clasico"
        if formato == "clasico" and HOJA_ACTIVIDADES not in L.hojas:
            raise ValueError(f"La planilla no tiene la hoja '{HOJA_ACTIVIDADES}' ni '{HOJA_MASTER}'. ¿Es el informe diario oficial?")
    
        # ---------------- hoja del informe
        xi = L.hoja_xml(HOJA_INFORME)
        filas_inf: dict[int, dict[str, object]] = {}
        form_inf: dict[int, dict[str, str]] = {}
        for n, a, fx, _, _ in X.filas(xi):
            cs = X.celdas(fx)
            filas_inf[n] = {c: L.valor_celda(at, inn) for c, (at, inn, _) in cs.items()}
            form_inf[n] = {c: f for c, (at, inn, _) in cs.items() if (f := _formula(inn))}
    
        def buscar(texto: str, col: str | None = None, desde: int = 0) -> int | None:
            t = norm(texto)
            for n in sorted(filas_inf):
                if n < desde:
                    continue
                for c, v in filas_inf[n].items():
                    if (col is None or c == col) and isinstance(v, str) and norm(v).startswith(t):
                        return n
            return None
    
        refs: dict[str, str] = {}
        f_fecha = buscar("FECHA", "AA")
        f_num = buscar("N° INFORME", "AA") or buscar("N INFORME", "AA")
        f_cond = buscar("CONDICION CLIMATICA", "AA")
        refs.update(fecha=f"AD{f_fecha}", numero=f"AD{f_num}", condicion=f"AD{f_cond}")
        for clave, rotulo in [("mod_acred", "MOD CONTRATADO ACREDITADO"), ("moi_acred", "MOI CONTRATADO ACREDITADO"),
                              ("mod_noacred", "MOD CONTRATADO NO ACREDITADO"), ("moi_noacred", "MOI CONTRATADO NO ACREDITADO"),
                              ("mod_total", "TOTAL CONTRATADOS MOD"), ("moi_total", "TOTAL CONTRATADOS MOI"),
                              ("proy_total", "TOTAL CONTRATADOS PROYECTO"), ("hh_constr", "MOD TRABAJANDO CONSTRUCCION"),
                              ("hh_precom", "MOD TRABAJANDO PRECOM"), ("hh_mod", "TOTAL MOD TRABAJANDO"),
                              ("hh_moi", "MOI TRABAJANDO"), ("cursos", "PERSONAL EN CURSOS"), ("lluvia", "CONDICION DE LLUVIA"),
                              ("alerta1", "ALERTA NIVEL 1"), ("alerta2", "ALERTA NIVEL 2")]:
            n = buscar(rotulo, "U") or (buscar(rotulo.replace("ALERTA ", "ALERTA VIENTO "), "U") if clave.startswith("alerta") else None)
            if n:
                refs[clave] = str(n)
        n = buscar("MOI TRABAJANDO", "U")
        if n:
            refs["hh_total"] = str(n + 1)
        elif "hh_mod" in refs:                              # formato nuevo: sin fila MOI, «TOTAL» bajo «TOTAL MOD TRABAJANDO»
            n = next((k for k in range(int(refs["hh_mod"]) + 1, int(refs["hh_mod"]) + 4)
                      if norm(filas_inf.get(k, {}).get("U")) == "TOTAL"), None)
            if n:
                refs["hh_total"] = str(n)
        n = buscar("RESGUARDO", "U")
        if n:
            # La fila de valores es la siguiente con número en AB
            refs["resguardo"] = str(next(k for k in range(n + 1, n + 6) if isinstance(filas_inf.get(k, {}).get("AB"), (int, float))))
        f_act = buscar("DESCRIPCION DE ACTIVIDADES", "C")
        refs["actividades"] = f"B{f_act + 1}" if f_act else "B16"
    
        # Filas de cargos (MOD y MOI)
        cargos: list[FilaCargo] = []
        totales_filas = {}
        col_hh = "X"
        for seccion, cab in (("MOD", "DETALLE PERSONAL DIRECTO"), ("MOI", "DETALLE PERSONAL INDIRECTO")):
            ini = buscar(cab, "C")
            if not ini:
                continue
            fin = buscar("TOTAL ESTE INFORME", "B", ini)
            for k in range(ini + 1, fin):
                v = filas_inf.get(k, {}).get("B")
                if isinstance(v, str) and v.strip():
                    cargos.append(FilaCargo(k, v.strip(), seccion, {c: form_inf.get(k, {}).get(c, "") for c in "LORU"}))
            totales_filas[seccion] = (fin, fin + 1, fin + 2)   # este informe, anterior, acumulado
            if seccion == "MOD":
                col_hh = next((c for c, v in filas_inf.get(ini, {}).items() if norm(v).startswith("HH FAENA")), "X")
        refs["mod_tot"], refs["mod_ant"], refs["mod_acum"] = map(str, totales_filas.get("MOD", (0, 0, 0)))
        refs["moi_tot"], refs["moi_ant"], refs["moi_acum"] = map(str, totales_filas.get("MOI", (0, 0, 0)))
    
        # Equipos
        equipos: list[FilaEquipo] = []
        for grupo, cab, fin_txt in (("OOCC", "DETALLE DE EQUIPOS DIRECTOS OOCC", "TOTALES EQUIPOS OOCC"),
                                    ("MONTAJE", "DETALLE DE EQUIPOS DIRECTOS MONTAJE", "TOTALES EQUIPOS MONTAJE"),
                                    ("INDIRECTOS", "DETALLE DE EQUIPOS INDIRECTOS", "TOTALES EQUIPOS INDIRECTOS")):
            ini = buscar(cab, "C")
            fin = buscar(fin_txt, "B", ini or 0) if ini else None
            if not ini or not fin:
                continue
            col_noche = "Y" if grupo == "INDIRECTOS" else "X"
            for k in range(ini + 1, fin):
                fv = filas_inf.get(k, {})
                nombre = fv.get("B")
                if not (isinstance(nombre, str) and nombre.strip()):
                    continue
                def n_(c):
                    v = fv.get(c)
                    return float(v) if isinstance(v, (int, float)) else 0.0
                equipos.append(FilaEquipo(k, nombre.strip(), grupo, fv.get("L") if isinstance(fv.get("L"), (int, float)) else None,
                                          str(fv.get("O") or "").strip(), str(fv.get("R") or "").strip(), n_("U"), n_(col_noche),
                                          float(fv["Y"]) if grupo != "INDIRECTOS" and isinstance(fv.get("Y"), (int, float)) else None,
                                          str(fv.get("AB") or "").strip(), col_noche))
    
        # Interferencias y uso de recurso
        def es_titulo(k):                                   # «5.0 INTERFERENCIAS…», «6.0 TRABAJOS EXTRAORDINARIOS…»
            v_ = filas_inf.get(k, {}).get("B")
            return isinstance(v_, str) and re.match(r"\s*\d+\.\d+\s+[A-ZÁÉÍÓÚ]", v_) is not None
        f_int = buscar("DETALLE DE INTERFERENCIA", "B")
        f_udr_sec = buscar("7.0 USO DE RECURSO", "B") or (next((k for k in sorted(filas_inf) if k > f_int and es_titulo(k)), None)
                                                         if f_int else None)
        filas_int = list(range(f_int + 1, f_udr_sec)) if f_int and f_udr_sec else []
        cab_int = {c: str(v).strip() for c, v in filas_inf.get(f_int, {}).items() if isinstance(v, str)} if f_int else {}
        f_udr = buscar("N° NDC", "B") or buscar("N NDC", "B")
        f_fotos = buscar("8.0 REGISTRO FOTOGRAFICO", "B") or buscar("9.0 REGISTRO FOTOGRAFICO", "B") or buscar("REGISTRO FOTOGRAFICO", "B")
        # Validación / firmas del formato nuevo
        validacion = {}
        f_val = buscar("8.0 VALIDACION DEL INFORME", "B") or buscar("VALIDACION DEL INFORME", "B")
        if f_val:
            cab = filas_inf.get(f_val + 1, {})
            cols_val = {k: next((c for c, v in cab.items() if norm(v) == r), None)
                        for k, r in (("nombre", "NOMBRE"), ("cargo", "CARGO"), ("fecha", "FECHA"), ("firma", "FIRMA"))}
            filas_val = {}
            for k in range(f_val + 2, f_val + 8):
                t_ = norm(filas_inf.get(k, {}).get("B"))
                for rol, clave in (("Elaborado por", "PREPARADO"), ("Revisado por", "REVISADO"), ("Aprobado por", "APROBADO")):
                    if t_.startswith(clave):
                        filas_val[rol] = k
            if filas_val and cols_val["nombre"]:
                validacion = {"filas": filas_val, "cols": cols_val, "titulo": f_val}
        # 8.0 Materiales y consumibles
        materiales = {}
        f_mat = buscar("8.0 MATERIALES", "B") or buscar("MATERIALES Y CONSUMIBLES", "B")
        if f_mat:
            cab = filas_inf.get(f_mat + 1, {})
            cols_m = {}
            for c_, v_ in cab.items():
                t_ = norm(v_)
                for k_, pref in (("descripcion", "DESCRIPCION"), ("unidad", "UNIDAD"), ("cantidad", "CANTIDAD"),
                                 ("trabajo", "ATCIVITY"), ("trabajo", "ACTIVITY"), ("solicitado", "SOLICITADO"),
                                 ("observacion", "OBSERVACION")):
                    if t_.startswith(pref) and k_ not in cols_m:
                        cols_m[k_] = c_
            fin_m = next((k for k in sorted(filas_inf) if k > f_mat + 1 and es_titulo(k)), f_mat + 12)
            if cols_m.get("descripcion"):
                materiales = {"filas": list(range(f_mat + 2, fin_m)), "cols": cols_m}
        if validacion:
            f_est = next((k for k in range(validacion["titulo"], validacion["titulo"] + 8)
                          if norm(filas_inf.get(k, {}).get("B")).startswith("ESTADO DE VALIDACION")), None)
            if f_est:
                obs_c = next((c_ for c_, v_ in filas_inf.get(f_est, {}).items() if norm(v_).startswith("OBSERVACIONES")), None)
                validacion.update(fila_estado=f_est, col_obs=obs_c, rotulo_obs=filas_inf[f_est].get(obs_c) if obs_c else None)
        if not f_fotos:
            # sin título de registro fotográfico: las fotos son las imágenes bajo la última sección del informe
            f_fotos = max([k for k in sorted(filas_inf) if es_titulo(k)] + [validacion.get("titulo", 0) + 7 if validacion else 0]) or None
        filas_udr = []
        if f_udr:
            fin_udr = next((k for k in sorted(filas_inf) if k > f_udr and es_titulo(k)), f_udr + 60)
            for k in range(f_udr + 1, fin_udr):
                if filas_inf.get(k, {}).get("B") is None and not form_inf.get(k, {}).get("B"):
                    break                                   # fin de la tabla «Uso de recurso»
                filas_udr.append(k)
        # ¿La tabla toma sus datos de la hoja BD UDR? (B = 'BD UDR'!C<n> y el resto con BUSCARX)
        udr_formulas = {}
        if filas_udr and "BD UDR" in (form_inf.get(filas_udr[0], {}).get("B") or "").upper() and "BD UDR" in L.hojas:
            udr_formulas = {c: f for c, f in form_inf.get(filas_udr[0], {}).items()}
    
        valores = {}
        for clave in ("fecha", "numero", "condicion", "actividades"):
            col, fila = X.separar_ref(refs[clave])
            valores[clave] = filas_inf.get(fila, {}).get(col)
        for clave in ("mod_acred", "moi_acred", "mod_noacred", "moi_noacred", "mod_total", "moi_total", "proy_total"):
            if clave in refs:
                valores[clave] = filas_inf[int(refs[clave])].get("AB")
        for clave in ("hh_constr", "hh_precom", "hh_mod", "hh_moi", "hh_total", "resguardo"):
            if clave in refs:
                fv = filas_inf.get(int(refs[clave]), {})
                valores[clave + "_ant"] = fv.get("AB")
                valores[clave + "_act"] = fv.get("AD")
                valores[clave + "_acum"] = fv.get("AF")
        for clave in ("cursos", "lluvia", "alerta1", "alerta2"):
            if clave in refs:
                fv = filas_inf.get(int(refs[clave]), {})
                valores[clave] = (fv.get("AB"), fv.get("AD"))
        for sec in ("mod", "moi"):
            fv = filas_inf.get(int(refs[f"{sec}_ant"]), {})
            valores[f"{sec}_ant"] = {c: fv.get(c) for c in "LORUXZ"}
            valores[f"{sec}_ant"]["HH"] = fv.get(col_hh)
    
        fecha = valores["fecha"]
        fecha = X.serial_a_fecha(fecha) if isinstance(fecha, (int, float)) else fecha
        if not isinstance(fecha, date):
            raise ValueError("No se pudo leer la fecha del informe (celda FECHA).")
    
        # ---------------- listas desplegables
        listas = {"condicion": [], "areas": [], "interferencia": []}
        if "Listas" in L.hojas:
            for n, a, fx, _, _ in X.filas(L.hoja_xml("Listas")):
                v = _celdas_fila(L, fx)
                if n <= 3:
                    continue
                for clave, col in (("condicion", "H"), ("areas", "M"), ("interferencia", "F")):
                    if isinstance(v.get(col), str) and v[col].strip():
                        listas[clave].append(v[col].strip())
    
        # ---------------- plantel del último día (BD Personal)
        xp = L.hoja_xml(HOJA_PERSONAL)
        serial = X.fecha_a_serial(fecha)
        personas, filas_xml = [], {}
        patron_fecha = re.compile(r'<c r="C\d+"[^>]*><v>' + str(serial) + r"</v>")
        historial: dict[str, dict[date, list[int]]] = {}
        re_c = {col: re.compile(rf'<c r="{col}\d+"([^>]*?)(?:/>|>(.*?)</c>)', re.S) for col in ("C", "F", "AB")}
    
        def celda(col, fx):
            m = re_c[col].search(fx)
            return L.valor_celda(X._attrs(m.group(1)), m.group(2)) if m else None
    
        for n, a, fx, _, _ in X.filas(xp):
            if n < 4:
                continue
            if not patron_fecha.search(fx):
                # historial: cuántos de cada sistema de turno estaban trabajando cada día
                f_ = celda("C", fx)
                g_ = grupo_turno(celda("F", fx))
                if isinstance(f_, (int, float)) and g_:
                    d_ = X.serial_a_fecha(f_)
                    par = historial.setdefault(g_, {}).setdefault(d_, [0, 0])
                    par[1] += 1
                    par[0] += str(celda("AB", fx) or "").strip().upper() in ESTADOS_TRABAJANDO
                continue
            v = _celdas_fila(L, fx)
            p = Persona(n, norm_rut(v.get("D")), str(v.get("I") or "").strip(), str(v.get("J") or "").strip(),
                        str(v.get("K") or "").strip(), str(v.get("H") or "").strip(), str(v.get("W") or "").strip(),
                        str(v.get("AA") or "").strip(), str(v.get("Z") or "").strip(), str(v.get("AB") or "").strip(),
                        str(v.get("F") or "").strip())
            personas.append(p)
            g_ = grupo_turno(p.turno)
            if g_:
                par = historial.setdefault(g_, {}).setdefault(fecha, [0, 0])
                par[1] += 1
                par[0] += p.faena.upper() in ESTADOS_TRABAJANDO
            filas_xml[n] = fx
        if not personas:
            raise ValueError(f"BD Personal no tiene filas con la fecha del informe ({fecha:%d-%m-%Y}).")
    
        # Estilos (y en el formato nuevo, fórmulas) de la última fila de la hoja de actividades
        xa = L.hoja_xml(HOJA_MASTER if formato == "master" else HOJA_ACTIVIDADES)
        ult, n_ult = None, 0
        for n, a, fx, _, _ in X.filas(xa):
            ult, n_ult = fx, n
        estilos = {c: at.get("s", "") for c, (at, _, _) in X.celdas(ult).items()} if ult else {}
        master_formulas, criterios = {}, []
        if formato == "master":
            master_formulas = {c: f for c, (at, inn, _) in X.celdas(ult).items() if (f := _formula(inn))} if ult else {}
            for n, a, fx, _, _ in X.filas(xa):
                if n >= 13:
                    break
                v_ = _celdas_fila(L, fx).get("B")
                if isinstance(v_, str) and v_.strip() and norm(v_) != "FECHA":
                    criterios.append(v_.strip())
        historial_master = _leer_historial_master(L, xa) if formato == "master" else {}
        # Sección 1.1 (HH por tipo de actividad): fila -> criterio
        filas_tipo = {}
        f_11 = buscar("1.1 HH POR TIPO", "B")
        if f_11:
            for k in range(f_11 + 2, f_11 + 30):
                t_ = filas_inf.get(k, {}).get("B")
                if not isinstance(t_, str) or norm(t_).startswith("TOTAL"):
                    break
                filas_tipo[k] = t_.strip()
            criterios = criterios or list(filas_tipo.values())
        # P6: Activity ID -> nombre, área, subsistema (la columna B es =TRIM(A) con su valor en caché)
        p6 = {}
        if HOJA_P6 in L.hojas:
            for n, a, fx, _, _ in X.filas(L.hoja_xml(HOJA_P6)):
                if n < 2:
                    continue
                v_ = _celdas_fila(L, fx)
                aid = str(v_.get("B") or v_.get("A") or "").strip()
                if aid and (v_.get("C") or not aid.startswith("SY-")) and aid not in p6:
                    p6[aid] = {c: x for c, x in v_.items() if x not in (None, "")}
        # BD UDR: columnas según sus encabezados
        udr_cols, udr_cab = {}, 0
        if "BD UDR" in L.hojas:
            for n, a, fx, _, _ in X.filas(L.hoja_xml("BD UDR")):
                v_ = _celdas_fila(L, fx)
                if any(norm(x) in ("N° NDC", "N NDC") for x in v_.values()):
                    udr_cab = n
                    vistos = set()
                    for c in sorted(v_, key=X.col_a_num):
                        clave = _clave_udr(norm(v_[c]), vistos)
                        if clave:
                            udr_cols[clave] = c
                            vistos.add(clave)
                    break
                if n > 30:
                    break
        con_formula = {f"{c}{n}" for n, fs in form_inf.items() for c in fs}
    
        # Fotos: imágenes ancladas bajo "8.0 REGISTRO FOTOGRÁFICO"
        fotos = _ubicar_fotos(L, xi, f_fotos, filas_inf)
    
        pl = Plantilla(datos, fecha, int(valores["numero"] or 0), str(valores["condicion"] or ""), personas, cargos, equipos,
                       valores, refs, listas, filas_int, filas_udr, filas_xml, estilos, str(valores.get("actividades") or ""), fotos)
        pl.totales = calcular_totales(pl, personas)
        pl.historial_turnos = {g: {d: tuple(v) for d, v in dias.items()} for g, dias in historial.items()}
        ndcs = [filas_inf.get(f, {}).get("B") for f in filas_udr]
        ndcs = [int(v) for v in ndcs if isinstance(v, (int, float)) and v > 0]
        pl.udr_formulas = udr_formulas
        pl.formato, pl.col_hh, pl.con_formula, pl.criterios = formato, col_hh, con_formula, criterios
        pl.filas_tipo_actividad, pl.p6, pl.master_formulas, pl.master_fila_modelo = filas_tipo, p6, master_formulas, n_ult
        pl.udr_cols, pl.udr_fila_cab, pl.cab_interferencia, pl.validacion = udr_cols, udr_cab, cab_int, validacion
        pl.historial_master = historial_master
        pl.maquinaria = _leer_maquinaria(L, fecha) if "BD Maquinaria" in L.hojas else []
        pl.materiales = materiales
        if udr_formulas:                                    # el último NDC registrado en BD UDR
            for n_, a_, fx_, _, _ in X.filas(L.hoja_xml("BD UDR")):
                c_ = X.celdas(fx_).get("C")
                if c_ and n_ > 2:
                    v_ = L.valor_celda(c_[0], c_[1])
                    if isinstance(v_, (int, float)) and v_ > 0:
                        ndcs.append(int(v_))
        pl.ultimo_ndc = max(ndcs) if ndcs else None
        return pl
    
    
    def _ubicar_fotos(L: X.Libro, xi: str, fila_sec: int | None, filas_inf) -> list[dict]:
        if not fila_sec:
            return []
        ruta_rels = L.hojas[HOJA_INFORME].replace("worksheets/", "worksheets/_rels/") + ".rels"
        if not L.existe(ruta_rels):
            return []
        rels = L.leer(ruta_rels)
        m = re.search(r'Id="([^"]+)"[^>]*Target="\.\./drawings/(drawing\d+\.xml)"|Target="\.\./drawings/(drawing\d+\.xml)"[^>]*Id="([^"]+)"', rels)
        if not m:
            return []
        dibujo = "xl/drawings/" + (m.group(2) or m.group(3))
        dxml = L.leer(dibujo)
        drel = L.leer(dibujo.replace("drawings/", "drawings/_rels/") + ".rels")
        destinos = dict(re.findall(r'Id="([^"]+)"[^>]*Target="\.\./media/([^"]+)"', drel))
        destinos.update({b: a for a, b in re.findall(r'Target="\.\./media/([^"]+)"[^>]*Id="([^"]+)"', drel)})
        anchos = _anchos_columnas(xi)
        fotos = []
        for anc in re.findall(r"<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>", dxml, flags=re.S):
            fr = re.search(r"<xdr:from><xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>", anc, flags=re.S)
            to = re.search(r"<xdr:to><xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>", anc, flags=re.S)
            emb = re.search(r'r:embed="([^"]+)"', anc)
            if not (fr and to and emb) or int(fr.group(2)) + 1 < fila_sec:
                continue
            c1, r1, c2, r2 = int(fr.group(1)), int(fr.group(2)), int(to.group(1)), int(to.group(2))
            ancho = sum(anchos.get(c, 9.0) for c in range(c1 + 1, c2 + 1)) * 7
            alto = (r2 - r1) * 15
            fotos.append({"rel": emb.group(1), "media": "xl/media/" + destinos.get(emb.group(1), ""), "fila": r1 + 1, "col": c1 + 1,
                          "fila_fin": r2 + 1, "aspecto": ancho / alto if alto else 1.33, "dibujo": dibujo})
        fotos.sort(key=lambda f: (f["fila"], f["col"]))
        # Leyenda: primera fila con texto bajo cada foto, en la celda combinada que la contiene
        for f in fotos:
            for k in range(f["fila_fin"], f["fila_fin"] + 4):
                fila = filas_inf.get(k, {})
                cand = [c for c in fila if X.col_a_num(c) <= f["col"] + 2 and isinstance(fila[c], str) and fila[c].strip()]
                if cand:
                    f["leyenda"] = f"{max(cand, key=X.col_a_num)}{k}"
                    break
            f.setdefault("leyenda", None)
        return fotos
    
    
    def _anchos_columnas(xml: str) -> dict[int, float]:
        anchos = {}
        for a in re.findall(r"<col\b([^>]*)/>", xml):
            d = X._attrs(a)
            for c in range(int(d.get("min", 0)), int(d.get("max", 0)) + 1):
                anchos[c] = float(d.get("width", 9))
        return anchos
    
    
    # ---------------------------------------------------------------- evaluación de COUNTIFS de la planilla
    
    _COL_CRIT = re.compile(r"'?([^'!]+)'?!\$?([A-Z]+):\$?([A-Z]+)")
    
    
    def _args(s: str) -> list[str]:
        args, prof, comillas, actual = [], 0, False, ""
        for ch in s:
            if ch == '"':
                comillas = not comillas
            if not comillas:
                if ch == "(":
                    prof += 1
                elif ch == ")":
                    prof -= 1
                elif ch == "," and prof == 0:
                    args.append(actual.strip())
                    actual = ""
                    continue
            actual += ch
        if actual.strip():
            args.append(actual.strip())
        return args
    
    
    def evaluar_countifs(formula: str, filas_bd: list[dict], cargo: str) -> int | None:
        """Evalúa una suma de COUNTIFS sobre BD Personal. Devuelve None si la fórmula no es de ese tipo."""
        if not formula or "COUNTIFS" not in formula.upper():
            return None
        total = 0
        for m in re.finditer(r"COUNTIFS\(", formula, flags=re.I):
            ini = m.end()
            prof, k = 1, ini
            while k < len(formula) and prof:
                if formula[k] == "(":
                    prof += 1
                elif formula[k] == ")":
                    prof -= 1
                k += 1
            args = _args(formula[ini:k - 1])
            criterios = []
            for rango, crit in zip(args[0::2], args[1::2]):
                mc = _COL_CRIT.search(rango)
                if not mc or norm(mc.group(1)) != norm(HOJA_PERSONAL):
                    return None
                col = mc.group(2)
                if crit.startswith('"'):
                    criterios.append((col, crit.strip('"')))
                elif re.search(r"\$?AD\$?\d+", crit):
                    criterios.append((col, "__FECHA__"))
                elif re.search(r"!?\$?B\$?\d+$", crit):
                    criterios.append((col, "__CARGO__"))
                else:
                    return None
            for f in filas_bd:
                ok = True
                for col, crit in criterios:
                    v = f.get(col)
                    if crit == "__FECHA__":
                        continue  # las filas ya vienen filtradas por fecha
                    if crit == "__CARGO__":
                        ok = norm(v) == norm(cargo)
                    elif crit == "*":
                        ok = isinstance(v, str) and v != "" or (v is not None and not isinstance(v, str))
                    else:
                        ok = norm(v) == norm(crit)
                    if not ok:
                        break
                if ok:
                    total += 1
        return total
    
    
    def _filas_bd(personas: list[Persona]) -> list[dict]:
        return [{"C": None, "D": p.rut, "H": p.tipo_turno, "I": p.nombre, "J": p.cargo, "K": p.especialidad, "W": p.moimod,
                 "Z": p.estatus, "AA": p.acreditado, "AB": p.faena} for p in personas]
    
    
    def calcular_totales(pl: Plantilla, personas: list[Persona]) -> dict:
        """Replica los conteos por cargo de la planilla (contratado, acreditado, faena día y noche, HH)."""
        bd = _filas_bd(personas)
        por_cargo: dict[str, list[dict]] = {}
        for f in bd:
            por_cargo.setdefault(norm(f["J"]), []).append(f)
        res = {"MOD": {c: 0 for c in "LORU"}, "MOI": {c: 0 for c in "LORU"}, "por_cargo": {}}
        for fc in pl.cargos:
            filas = por_cargo.get(norm(fc.cargo), [])
            vals = {}
            for col in "LORU":
                v = evaluar_countifs(fc.formulas.get(col, ""), filas, fc.cargo)
                vals[col] = v or 0
                res[fc.seccion][col] += vals[col]
            vals["X"] = (vals["R"] + vals["U"]) * HORAS_JORNADA
            res["por_cargo"][fc.fila] = vals
        for sec in ("MOD", "MOI"):
            res[sec]["X"] = (res[sec]["R"] + res[sec]["U"]) * HORAS_JORNADA
        cargos_inf = {norm(fc.cargo) for fc in pl.cargos}
        res["sin_fila"] = sorted({p.cargo for p in personas if norm(p.cargo) not in cargos_inf and p.cargo})
        return res
    
    
    # ====================================================================== consolidación
    
    @dataclass
    class Cambio:
        rut: str
        nombre: str
        cargo: str
        especialidad: str
        antes: str
        ahora: str
        motivo: str
    
    
    @dataclass
    class Consolidado:
        fecha: date
        personas: list[Persona]                      # plantel del nuevo día (estado actualizado)
        cambios: list[Cambio]
        nuevos: list[dict]                           # en reportes pero no en BD Personal
        actividades: list[dict]                      # filas para el texto del informe
        filas_actividad: list[dict]                  # filas para BD Actividades MOD
        hh_constr: float
        hh_precom: float
        totales: dict
        avisos: list[str]
        especialidades_con_reporte: list[str]
        especialidades_sin_reporte: dict[str, int]
        uso_recurso: list[dict] = field(default_factory=list)   # sugerencias para «7.0 Uso de recurso»
        materiales: list[dict] = field(default_factory=list)    # sugerencias para «8.0 Materiales y consumibles»
        interferencias: list[dict] = field(default_factory=list)  # sugerencias para «5.0 Interferencias y detenciones»
        clasificacion: dict = field(default_factory=dict)       # cursos / lluvia / alertas: (cantidad, HH)
    
    
    def clave_area(rep_area: str, especialidad: str) -> str:
        return f"{norm(especialidad)}|{norm(rep_area)}"
    
    
    # ====================================================================== Activity ID y área (P6)
    
    _STOP_P6 = set("de del la las el los y en a al para con por un una se realiza realizan sector linea lineas cant gl m2 m3 ml kg "
                   "mts tramo no".split())
    _DISC_P6 = {"CAN": ("CANERIA", "PIPING", "SOLDADURA"), "MEC": ("MECANIC", "SOLDADURA", "MONTAJE"), "ELE": ("ELECTRIC",),
                "TIC": ("INSTRUMENT",), "CIV": ("CIVIL", "OO-CC", "OOCC", "OO.CC"), "MOV": ("CIVIL", "OO-CC", "OOCC", "MOVIMIENTO"),
                "EST": ("ESTRUCTURA", "MONTAJE"), "ARQ": ("ARQUITECT",)}
    _MESES_P6 = {m: i + 1 for i, m in enumerate("JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC".split())}
    _RE_ID = re.compile(r"\bS\s*Y\s*[-_ ]?\s*(\d{4})\s*[-_ ]?\s*([A-Z]{3,6})\s*[-_ ]?\s*(\d{3,4})\b")
    _RE_SUB = re.compile(r"(?:\bS\s*S\s*[-_ ]?\s*)?\b(29\d\d)\s*[-_]\s*(\d{2})\s*[-_]\s*([A-Z]\d|\d{2})\s*[-_]\s*(\d{3}[A-Z]?)\b")
    
    
    def _tokens_p6(t) -> list[str]:
        t = norm(t).lower()
        return [w if w.isdigit() else w[:6] for w in re.findall(r"[a-z]+|\d+", t) if w not in _STOP_P6 and not (w.isalpha() and len(w) < 3)]
    
    
    def _num_tags(t) -> set[str]:
        """Números de tags (líneas, equipos), sin códigos de área (29xx) ni cantidades chicas."""
        return {str(int(n)) for n in re.findall(r"\d{2,}", norm(t)) if int(n) > 20 and not (2900 <= int(n) <= 2999)}
    
    
    def _fecha_p6(v):
        m = re.search(r"(\d{1,2})-([A-Za-z]{3})-(\d{2})", str(v or ""))
        if not m:
            return None, False
        try:
            return date(2000 + int(m.group(3)), _MESES_P6.get(m.group(2).upper(), 1), int(m.group(1))), str(v).strip().endswith("A")
        except ValueError:
            return None, False
    
    
    def _caldera(t) -> int | None:
        m = re.search(r"(?:CALDERA|BOILER|CHIMENEA)\s*(?:N\s*[°º]?\s*|#\s*)?(\d)\b", norm(t))
        return int(m.group(1)) if m else None
    
    
    def _indice_p6(pl: "Plantilla") -> dict:
        ix = getattr(pl, "_indice_p6", None)
        if ix is not None:
            return ix
        items, df, por_sub, por_norm = [], {}, {}, {}
        for aid, v in pl.p6.items():
            por_norm[re.sub(r"[^A-Z0-9]", "", aid.upper())] = aid
            sub, area = str(v.get("I") or "").strip(), v.get("K")
            if sub and area:
                por_sub.setdefault(re.sub(r"[^A-Z0-9]", "", sub.upper().removeprefix("SS")), {}).setdefault(area, 0)
                por_sub[re.sub(r"[^A-Z0-9]", "", sub.upper().removeprefix("SS"))][area] += 1
            nombre = str(v.get("C") or "")
            if not aid.startswith("SY-") or not nombre:
                continue
            tk = set(_tokens_p6(nombre))
            ini, _ = _fecha_p6(v.get("E"))
            fin, fin_real = _fecha_p6(v.get("F"))
            items.append({"id": aid, "nombre": nombre, "tk": tk, "num": _num_tags(nombre), "area": area, "sub": sub,
                          "disc": re.sub(r"^SY-\d{4}([A-Z]{3}).*", r"\1", aid), "ini": ini, "fin": fin, "fin_real": fin_real,
                          "caldera": _caldera(f"{nombre} {area}")})
            for w in tk:
                df[w] = df.get(w, 0) + 1
        n = max(1, len(items))
        idf = {w: math.log((n + 1) / (c + 0.5)) for w, c in df.items()}
        for it in items:
            it["norma"] = math.sqrt(sum(idf.get(w, 1) ** 2 for w in it["tk"])) or 1
        ix = {"items": items, "idf": idf, "por_sub": {k: max(v, key=v.get) for k, v in por_sub.items()}, "por_norm": por_norm,
              "areas": sorted({str(v.get("K")) for v in pl.p6.values() if v.get("K")})}
        pl._indice_p6 = ix
        return ix
    
    
    def _area_por_texto(t: str, areas: list[str]) -> str | None:
        """Área del P6 a partir del texto (Caldera N -> Boiler N; nombre de área escrito tal cual)."""
        n = _caldera(t)
        if n and f"Boiler {n}" in areas:
            return f"Boiler {n}"
        tt = norm(t)
        for a in areas:
            if len(a) > 3 and norm(a) in tt:
                return a
        if re.search(r"\bTRENCH\b|\bCELDA\s*1[5-8]\b", tt):
            return next((a for a in areas if a.upper().startswith("TRENCH")), None)
        return None
    
    
    _CODIGO_POR_PALABRA = (
        (2963, ("CALDERA", "BOILER", "ECONOMIZADOR", "CHIMENEA", "CHIMENA", "SOPLADOR", "CALENTAMIENTO", "INTERCAMBIADOR",
                "QUEMADOR")),
        (2962, ("CANALETA", "CELDA", "LIRA", "REFINO", "TRENCH", "LIXIVIA")),
        (2964, ("INCENDIO", "CONTRAINCENDIO", "RCI")),
        (2965, ("AGUA FRESCA", "AGUAS FRESCAS")),
        (2966, ("SALA ELECTRICA", "SUBESTACION", "DISTRIBUCION")),
        (2967, ("IIFF", "II.FF", "INSTALACION DE FAENA", "INSTALACIONES DE FAENA", "INFRAESTRUCTURA", "OFICINA", "MODULAR",
                "COMEDOR", "CAMARIN", "BODEGA")),
    )
    _ELECTRICO = ("CIRCUITO", "CABLEADO", "CONEXIONADO", "ACOMETIDA", "TENDIDO", "CANALIZACION", "ATERRAMIENTO", "INSTRUMENTO",
                  "ESCALERILLA", "LUMINARIA", "TABLERO", "MALLA A TIERRA", "BANCODUCTO", "BANCO DUCTO", "CONDUIT")
    _PIPING = ("LINEA", "SPOOL", "CANERIA", "SOLDADURA", "FLANGE", "UNION", "CARRETE", "VALVULA", "TEST PACK", "SPECTACLE")
    
    
    def _historial_encabezados(pl: "Plantilla") -> list[tuple[str, set]]:
        h = getattr(pl, "_hist_enc", None)
        if h is None:
            h, actual = [], None
            for linea in (pl.actividades_texto or "").splitlines():
                l = linea.strip()
                if l.endswith(":") and not re.match(r"^\d+\s*[-.)]", l):
                    actual = None if norm(l) in ("NOTA:", "NOTAS:", "TURNO NOCHE:") else l[:-1].strip()
                elif actual and l:
                    h.append((actual, set(_tokens_p6(re.sub(r"^\d+\s*[-.)]\s*", "", l)))))
            pl._hist_enc = h
        return h
    
    
    _RE_UDR = re.compile(r"\(\s*ADICIONAL|TRABAJOS? ADICIONAL|ADICIONAL PRECOM|[-–:]\s*ADICIONAL\b|\bADICIONAL\s*\)?\s*$|"
                         r"\bUDR\b|\bNDC\b|EXTRAORDINARI|FUERA DE ALCANCE|TRABAJO EXTRA\b|\bPRECOM\b")
    _RE_NO_UDR = re.compile(r"\bDETENCION|\bDESVIACION|\bRETRASO|\bINTERFERENCIA|\bESPERA\b|\bCHARLA\b")
    
    
    def area_udr(texto: str, area: str = "") -> str:
        """Área / frente como se escribe en la sección UDR del informe: «Caldera N°4», «Caldera N°3 y N°4», «Gasco», «Lira N°3»,
        «Celda 15», «Celda 17 y 18», «Precom»… (si no hay pista, el área del reporte)."""
        t = norm(texto)
        if "GASCO" in t:
            return "Gasco"
        m = re.search(r"\b(?:A|HACIA|HASTA)\s+(?:LA\s+)?CELDA\s*(\d+)", t)           # traslado: el destino
        if m:
            return f"Celda {m.group(1)}"
        def numeros(patron):
            m = re.search(patron, t)
            return re.findall(r"\d+", m.group(1)) if m else []
        n = numeros(r"\bLIRA\s*(?:N\s*[°º]?\s*)?(\d+(?:\s*(?:Y|,|-)\s*\d+)*)")
        if n:
            return "Lira " + " y ".join(f"N°{x}" for x in n)
        n = numeros(r"\bCALDERAS?\s*(?:N\s*[°º]?\s*|#\s*)?(\d(?:\s*(?:Y|,|-)\s*(?:N\s*[°º]?\s*)?\d)*)\b")
        if n:
            return "Caldera " + " y ".join(f"N°{x}" for x in dict.fromkeys(n))
        n = numeros(r"\bCELDAS?\s*(\d+(?:\s*(?:Y|,|-|A)\s*\d+)*)")
        if n:
            return "Celda " + " y ".join(dict.fromkeys(n))
        for clave, nombre in (("MODULO DE DESCARGA", "Módulo de Descarga"), ("SALA ELECTRICA", "Sala Eléctrica"),
                              ("INTERCAMBIADOR", "Intercambiadores de Calor"), ("AGUA FRESCA", "Agua Fresca"),
                              ("CANALETA STOP", "Canaleta Stop")):
            if clave in t:
                return nombre
        if re.search(r"\bPRECOM\b", t):
            return "Precom"
        a_ = (area or "").strip()
        if a_.isupper() and len(a_.replace(".", "")) > 5:
            return a_[:1] + a_[1:].lower()
        return a_
    
    
    def respaldo_udr(texto: str) -> str:
        """Respaldo citado en el texto: «Meeting note 107», «RFI 12», «DCN 5»."""
        m = re.search(r"\b(ME+A?T(?:E|I)?NG\s*NOTES?|MEETING\s*NOTES?|RFI|DCN)\s*(?:N\s*[°º]?\s*)?(\d+)", norm(texto))
        if not m:
            return ""
        tipo = "Meeting note" if m.group(1).startswith("M") else m.group(1)
        return f"{tipo} {m.group(2)}"
    
    
    def es_udr(texto: str, activity_id: str | None = None, pl: "Plantilla | None" = None) -> bool:
        """¿La actividad corresponde a un uso de recurso / trabajo adicional (UDR / NDC)?"""
        if activity_id and ("UDR" in activity_id.upper() or (pl and str(pl.p6.get(activity_id, {}).get("K", "")).upper() == "UDR")):
            return True
        t = norm(texto)
        return bool(_RE_UDR.search(t)) and not _RE_NO_UDR.search(t)
    
    
    # Criterio (tipo de actividad del control T&M) según el texto. Cada regla: (palabras del criterio, patrón del texto)
    _REGLAS_CRITERIO = (
        (("TRASLADO", "CAMBIO DE TURNO"), r"\b[BV]AJADA\b|\bSUBIDA\b|CAMBIO DE TURNO|ENTREGA DE TURNO|RECAMBIO|TRASLADO (DE|DEL) PERSONAL|"
                                           r"TRASLADO (A|HACIA|DESDE) (FAENA|CALAMA|CAMPAMENTO|HOTEL)|\bII\.?FF HACIA"),
        (("HABILITACION", "INDUCCION", "CURSO"), r"\bCURSO|INDUCCION|CAPACITACION|HABILITACION|HIPOBARIA|EXAMEN|\bACREDITACION|"
                                                  r"CHARLA DE INDUCCION|INHABILITAD"),
        (("DETENCION", "MANDANTE"), r"DETENCION|DETENIDO|PARALIZACION|SUSPENSION|ALERTA|TORMENTA|\bVIENTO|LLUVIA|PRESURIZACION|"
                                     r"DESPRESURIZACION|RETRASO EN EL INICIO|SIN PERMISO|ESPERA DE PERMISO"),
        (("INTERFERENCIA",), r"INTERFERENCIA|OTRO CONTRATISTA|OTRA EMPRESA|\bESPERA\b|\bESPERANDO\b|FRENTE OCUPADO|SIN FRENTE"),
        (("DESVIO", "EXTRAORDINARIA"), r"DESVI(O|ACION) DE (RECURSOS? )?(A |PARA )?ADICIONAL"),
        (("DESVIO", "ALCANCE"), r"DESVI(O|ACION) DE RECURSO|\bDESVIAD|APOYO A OTRA (CUADRILLA|ESPECIALIDAD|AREA)|PRESTAMO DE PERSONAL"),
        (("RETRABAJO", "MANDANTE"), r"RETRABAJO.*(MANDANTE|CLIENTE|INGENIERIA|DISENO)|CAMBIO DE INGENIERIA|MODIFICACION DE INGENIERIA"),
        (("RETRABAJO", "CONTRATISTA"), r"RETRABAJO|REHACER|CORRECCION DE (SOLDADURA|MONTAJE|INSTALACION)|REPARACION DE SOLDADURA|"
                                        r"\bREPARACION\b.*\b(DEFECTO|RECHAZ)"),
    )
    
    
    def criterio_sugerido(pl: "Plantilla", texto: str, udr: bool = False) -> str | None:
        """Criterio del control T&M (Planificada, Extraordinaria, Interferencia, Desvío, Detención, Habilitación, Traslado,
        Retrabajo) según la descripción; None si no hay pista (queda «Planificada»)."""
        crits = list(pl.criterios or [])
        def buscar(*palabras):
            return next((c for c in crits if all(p in norm(c) for p in palabras)), None)
        t = norm(texto)
        for palabras, patron in _REGLAS_CRITERIO:
            if re.search(patron, t):
                c = buscar(*palabras)
                if c:
                    return c
        if udr:
            return buscar("EXTRAORDINARIA")
        return None
    
    
    def encabezado_sugerido(pl: "Plantilla", desc: str, area: str = "", especialidad: str = "", activity_id: str | None = None) -> str:
        """Encabezado de la descripción de actividades («2963 Sistema de Calentamiento», «2960 Tendido y Conexionado…»):
        1) misma actividad (o muy parecida) en el informe anterior; 2) código de área (Activity ID, tag 29xx o palabras clave)
        y disciplina, entre los encabezados del informe anterior y la lista de áreas de la planilla."""
        previos = encabezados_previos(pl.actividades_texto or "")
        opciones = list(dict.fromkeys(previos + list(pl.listas.get("areas", []))))
        if not opciones:
            return ""
        tk = set(_tokens_p6(desc))
        if tk:
            mejor, enc = 0.0, ""
            for e, t in _historial_encabezados(pl):
                if t:
                    sim = len(tk & t) / len(tk | t)
                    if sim > mejor:
                        mejor, enc = sim, e
            if mejor >= 0.6:
                return enc
        texto = norm(f"{desc} {area}")
        esp = norm(especialidad)
        electrico = any(k in esp for k in ("ELECTRIC", "INSTRUMENT")) or any(re.search(rf"\b{k}", texto) for k in _ELECTRICO)
        piping = (any(k in esp for k in ("PIPING", "CANERIA", "SOLDADURA")) or any(re.search(rf"\b{k}", texto) for k in _PIPING)) \
            and not electrico
        def con_codigo(cod, palabra=None):
            cands = [o for o in opciones if re.match(rf"\s*{cod}\b", o)]
            if palabra:
                cands = [o for o in cands if palabra in norm(o)]
            return cands[0] if cands else None
        if electrico and con_codigo(2960):
            return con_codigo(2960)
        codigo = None
        if activity_id and pl.p6.get(activity_id, {}).get("G"):
            try:
                codigo = int(pl.p6[activity_id]["G"])
            except (TypeError, ValueError):
                codigo = None
        if codigo is None:
            cods = [int(c) for c in re.findall(r"(?<!\d)(29[0-9]\d)(?!\d)", texto) if c != "2960"]
            codigo = max(set(cods), key=cods.count) if cods else None
        if codigo is None:
            for cod, palabras in _CODIGO_POR_PALABRA:
                if any(p in texto for p in palabras):
                    codigo = cod
                    break
        if codigo is None:
            piping_unico = [o for o in opciones if "PIPING" in norm(o)]
            return piping_unico[0] if piping and len(piping_unico) == 1 else ""
        if piping and con_codigo(codigo, "PIPING"):
            return con_codigo(codigo, "PIPING")
        no_piping = [o for o in opciones if re.match(rf"\s*{codigo}\b", o) and "PIPING" not in norm(o)]
        return no_piping[0] if no_piping else (con_codigo(codigo) or (con_codigo(2900) if codigo == 2967 else "") or "")
    
    
    _RE_GENERICA = re.compile(r"\bCHARLA|DOCUMENT|\bORDEN Y ASEO|\bASEO\b|PERMISO DE TRABAJO|\bAST\b|\bART\b|\bPTS\b|"
                              r"\bREPORT|INSPECCION DE HERRAMIENTAS|CHEQUEO DE HERRAMIENTAS")
    
    
    def es_generica(desc: str) -> bool:
        """Actividades de apoyo de cada cuadrilla (charla, documentos, aseo): van con el Activity ID de su trabajo principal."""
        return bool(_RE_GENERICA.search(norm(desc)))
    
    
    def _match_historial(pl: "Plantilla", desc: str):
        """(Activity ID, criterio, similitud) de la actividad más parecida en el historial de Master Actividades."""
        h = getattr(pl, "historial_master", None)
        if not h or not desc:
            return None
        exacto = h.get(norm(desc)[:200])
        if exacto:
            return exacto[0], exacto[1], 1.0
        idx = getattr(pl, "_idx_hist", None)
        if idx is None:
            idx = [(set(_tokens_p6(k)), v) for k, v in h.items()]
            pl._idx_hist = idx
        tk = set(_tokens_p6(desc))
        if len(tk) < 2:
            return None
        mejor, val = 0.0, None
        for t, v in idx:
            if t:
                sim = len(tk & t) / len(tk | t)
                if sim > mejor:
                    mejor, val = sim, v
        return (val[0], val[1], mejor) if val and mejor >= 0.6 else None
    
    
    def identificar_actividad(pl: "Plantilla", desc: str, area: str = "", extra: str = "", especialidad: str = "",
                              fecha: date | None = None) -> dict:
        """Activity ID, subsistema y área (P6) de una actividad de un reporte de terreno.
        Orden: ID escrito en el descriptor > subsistema escrito (da el área) > coincidencia clara con el nombre de la actividad P6.
        Devuelve {"activity_id", "subsistema", "area_p6", "nombre_p6", "origen_id"} (vacíos si no hay nada confiable)."""
        vacio = {"activity_id": None, "subsistema": None, "area_p6": None, "nombre_p6": None, "origen_id": None, "criterio": None}
        if not pl.p6:
            return vacio
        ix = _indice_p6(pl)
        texto = f"{desc} {area} {extra}"
        up = norm(texto)
        r = dict(vacio)
        for m in _RE_ID.finditer(up):
            aid = ix["por_norm"].get(f"SY{m.group(1)}{m.group(2)}{m.group(3)}")
            if aid:
                v = pl.p6[aid]
                r.update(activity_id=aid, subsistema=v.get("I"), area_p6=v.get("K"), nombre_p6=v.get("C"), origen_id="Reporte")
                return r
        hist = _match_historial(pl, desc)
        if hist:
            aid, crit, sim = hist
            r["criterio"] = crit
            if aid and not es_generica(desc):           # charla, documentos, orden y aseo: llevan el ID de la cuadrilla
                v = pl.p6.get(aid, {})
                r.update(activity_id=aid, subsistema=v.get("I"), area_p6=v.get("K"), nombre_p6=v.get("C") or aid,
                         origen_id=f"Informe anterior ({int(sim * 100)}%)")
                return r
        for m in _RE_SUB.finditer(up):
            clave = "".join(m.groups())
            if clave in ix["por_sub"]:
                r.update(subsistema=f"SS-{m.group(1)}-{m.group(2)}-{m.group(3)}-{m.group(4)}", area_p6=ix["por_sub"][clave])
                break
        if not r["area_p6"]:
            r["area_p6"] = _area_por_texto(f"{area} {desc}", ix["areas"])
        # coincidencia por texto con las actividades del P6 (solo si es clara)
        tk = set(_tokens_p6(f"{desc} {area}"))
        if tk:
            idf = ix["idf"]
            nums = _num_tags(desc)
            norma = math.sqrt(sum(idf.get(w, 1) ** 2 for w in tk)) or 1
            cal = _caldera(f"{desc} {area}")
            esp = norm(especialidad)
            puntajes = []
            for it in ix["items"]:
                comun = tk & it["tk"]
                if not comun:
                    continue
                sc = sum(idf.get(w, 1) ** 2 for w in comun) / (norma * it["norma"]) + 0.3 * len(nums & it["num"])
                if cal and it["caldera"]:
                    sc += 0.12 if cal == it["caldera"] else -0.3
                if r["area_p6"] and it["area"] and it["area"] != r["area_p6"]:
                    sc -= 0.1
                if esp and any(k in esp for k in _DISC_P6.get(it["disc"], ())):
                    sc += 0.08
                if fecha and it["fin"] and it["fin_real"] and (fecha - it["fin"]).days > 10:
                    sc -= 0.15
                if fecha and it["ini"] and (it["ini"] - fecha).days > 30:
                    sc -= 0.1
                puntajes.append((sc, it))
            if puntajes:
                puntajes.sort(key=lambda x: -x[0])
                sc, it = puntajes[0]
                margen = sc - (puntajes[1][0] if len(puntajes) > 1 else 0)
                if sc >= 0.7 and margen >= 0.15:
                    r.update(activity_id=it["id"], nombre_p6=it["nombre"], origen_id=f"Sugerido ({min(99, int(sc * 100))}%)",
                             area_p6=it["area"] or r["area_p6"], subsistema=r["subsistema"] or it["sub"])
        return r
    
    
    def consolidar(pl: Plantilla, reportes: list[ReporteTerreno], fecha: date,
                   ajustes_estado: dict[str, str] | None = None, encabezados: dict[str, str] | None = None,
                   agregar_nuevos: bool = True, tarja: dict | None = None) -> Consolidado:
        """Estado de cada persona, por prioridad: ajuste manual > tarja de asistencia > reportes de terreno > turno."""
        ajustes_estado = ajustes_estado or {}
        encabezados = encabezados or {}
        avisos: list[str] = []
        plantel = {p.rut: p for p in pl.personas}
    
        # RUT que no está en BD Personal: se busca a la persona por su nombre (RUT con error de tipeo o en otro formato)
        por_nombre: dict[frozenset, list[str]] = {}
        for p in pl.personas:
            por_nombre.setdefault(_tokens_nombre(p.nombre), []).append(p.rut)
        for rep in reportes:
            for t in rep.trabajadores:
                if t.rut in plantel or not t.nombre:
                    continue
                tk = _tokens_nombre(t.nombre)
                cand = por_nombre.get(tk, [])
                if not cand and len(tk) >= 2:
                    cand = [r for k, rs in por_nombre.items() if tk <= k for r in rs]
                # mismo número de RUT con otro dígito verificador (error de tipeo): se prefiere ese, si hay coincidencia de nombre
                cuerpo = str(t.rut or "").split("-")[0]
                mismo_num = [r for r in plantel if cuerpo and r.split("-")[0] == cuerpo]
                if len(mismo_num) == 1 and (len(cand) != 1 or cand[0] != mismo_num[0]):
                    if not cand or mismo_num[0] in cand or tk & _tokens_nombre(plantel[mismo_num[0]].nombre):
                        cand = mismo_num
                if len(cand) == 1:
                    avisos.append(f"{t.nombre}: el RUT del reporte ({t.rut or 'vacío'}) no está en BD Personal; se asoció por nombre a "
                                  f"{plantel[cand[0]].nombre} ({cand[0]}). Revisa el RUT en el reporte {rep.archivo}.")
                    t.rut = cand[0]
    
        # Tarja de asistencia: RUT tal como está en BD Personal (mismo número con otro formato o dígito verificador)
        marcas: dict[str, Marca] = {}
        fuera_tarja = []
        nuevos_tarja: list[tuple[str, Marca]] = []
        finiquitados: dict[str, Marca] = {}
        if tarja:
            por_numero = {}
            for r in plantel:
                por_numero.setdefault(r.split("-")[0], []).append(r)
            for rut, m in tarja.items():
                m = m if isinstance(m, Marca) else Marca(**m)
                r = rut if rut in plantel else None
                if r is None:
                    cand = por_numero.get(rut.split("-")[0], [])
                    if len(cand) == 1:
                        r = cand[0]
                        avisos.append(f"Tarja: el RUT {rut} ({m.nombre or 'sin nombre'}) se asoció a {plantel[r].nombre} ({r}).")
                if r is None and m.nombre:
                    tk = _tokens_nombre(m.nombre)
                    cand = por_nombre.get(tk, []) or ([x for k, xs in por_nombre.items() if tk <= k for x in xs] if len(tk) >= 2 else [])
                    if len(cand) == 1:
                        r = cand[0]
                        avisos.append(f"Tarja: el RUT {rut} ({m.nombre}) no está en BD Personal; se asoció por nombre a {plantel[r].nombre} ({r}).")
                if r is None:
                    if _vigente(m) and m.ficha.get("J"):
                        nuevos_tarja.append((rut, m))              # contratado según la tarja: se agrega a BD Personal
                    elif m.estado in ("P", "T", "TN", "TT", "TC"):     # solo importa si figura trabajando
                        fuera_tarja.append(f"{m.nombre or '–'} ({rut}: {m.texto or m.estado})")
                    continue
                if m.ficha and _finiquitado(m):
                    finiquitados[r] = m
                    continue
                marcas[r] = m
            if fuera_tarja:
                avisos.append(f"{len(fuera_tarja)} persona(s) trabajando según la tarja no están en BD Personal (no se cuentan; "
                              "agrégalas a BD Personal si corresponde): " + "; ".join(fuera_tarja[:15])
                              + (" …" if len(fuera_tarja) > 15 else "") + ".")
    
        presentes: dict[str, tuple[ReporteTerreno, Trabajador]] = {}
        ausentes: dict[str, tuple[ReporteTerreno, Trabajador]] = {}
        for rep in reportes:
            if rep.fecha and rep.fecha != fecha:
                avisos.append(f"{rep.archivo}: la fecha del reporte ({rep.fecha:%d-%m-%Y}) no coincide con el informe ({fecha:%d-%m-%Y}).")
            for t in rep.trabajadores:
                if not t.total and t.nota:
                    ausentes.setdefault(t.rut, (rep, t))     # permiso, licencia, «no subió»: no está en faena
                    continue
                if t.rut in presentes:
                    rep0, t0 = presentes[t.rut]
                    if rep.tipo == "terreno" and rep0.tipo != "terreno":
                        presentes[t.rut] = (rep, t)           # el reporte de terreno manda sobre un UDR
                    elif rep.tipo == "terreno" and rep0.archivo.split(" › ")[0] == rep.archivo.split(" › ")[0]:
                        t0.total += t.total                   # misma persona en otra hoja del mismo archivo: se suman sus horas
                    elif rep.tipo == "terreno":
                        otro = presentes[t.rut][0].archivo
                        avisos.append(f"{t.nombre} ({t.rut}) aparece en dos reportes: {otro} y {rep.archivo}. Se usa el primero.")
                    continue
                presentes[t.rut] = (rep, t)
    
        def esp_de(rep: ReporteTerreno) -> str:
            if rep.especialidad and rep.tipo != "terreno":
                return rep.especialidad
            esp = [plantel[t.rut].especialidad for t in rep.trabajadores if t.rut in plantel and plantel[t.rut].especialidad]
            return max(set(esp), key=esp.count) if esp else (rep.area_reporte or "SIN ESPECIALIDAD")
    
        esp_reporte = {id(r): esp_de(r) for r in reportes}
        # Cobertura por especialidad Y turno: sin reporte de noche, el personal de noche mantiene su estado.
        cubiertas = {(norm(plantel[r].especialidad), presentes[r][0].turno) for r in presentes
                     if r in plantel and norm(plantel[r].moimod) == "DIRECTO"
                     and (presentes[r][0].tipo == "terreno" or getattr(presentes[r][0], "precom", False))}
    
        def turno_de(p: Persona) -> str:
            return "Noche" if norm(p.tipo_turno) == "NOCHE" or p.faena.upper() == "TN" else "Día"
    
        personas, cambios = [], []
        sin_reporte: dict[str, int] = {}
        sin_marca: list[str] = []
        cambios_acr: list[str] = []
        if finiquitados:
            avisos.append(f"{len(finiquitados)} persona(s) de BD Personal figuran FINIQUITADAS en la tarja y salen del día "
                          "(no se cuentan como contratadas): " + "; ".join(f"{plantel[r].nombre} ({r})" for r in list(finiquitados)[:15])
                          + (" …" if len(finiquitados) > 15 else "") + ".")
        for p in pl.personas:
            if p.rut in finiquitados:
                continue
            q = Persona(**p.__dict__)
            motivo = ""
            if p.rut in presentes:
                rep, t = presentes[p.rut]
                noche = rep.turno == "Noche"
                q.faena = "TN" if noche else "T"
                nuevo_tt = "NOCHE" if noche else "DIA"
                if norm(p.tipo_turno) != nuevo_tt:
                    q.tipo_turno = nuevo_tt
                    avisos.append(f"{p.nombre}: figura en BD Personal con turno {p.tipo_turno or '–'} pero viene en un reporte de turno "
                                  f"{'noche' if noche else 'día'}; se registra como {nuevo_tt}.")
                motivo = f"Reporte {rep.archivo}"
            elif p.rut in ausentes and estado_por_nota(ausentes[p.rut][1].nota):
                q.faena = estado_por_nota(ausentes[p.rut][1].nota)
                motivo = f"Reporte {ausentes[p.rut][0].archivo}: «{ausentes[p.rut][1].nota}»"
            elif norm(p.moimod) == "DIRECTO" and (norm(p.especialidad), turno_de(p)) in cubiertas:
                if p.faena.upper() in ESTADOS_TRABAJANDO:
                    q.faena = "D"
                    motivo = "No aparece en los reportes de su especialidad"
            else:
                if p.faena.upper() in ESTADOS_TRABAJANDO and norm(p.moimod) == "DIRECTO":
                    k = f"{p.especialidad or 'SIN ESPECIALIDAD'} ({turno_de(p).lower()})"
                    sin_reporte[k] = sin_reporte.get(k, 0) + 1
            cubierto = norm(p.moimod) == "DIRECTO" and (norm(p.especialidad), turno_de(p)) in cubiertas
            if not motivo and not cubierto and p.rut not in presentes and p.rut not in ausentes:
                # Sin reporte: subidas y bajadas según el sistema de turno (14x14, 15x13, 8x6…) y su historial
                g = grupo_turno(p.turno)
                hoy = turno_trabaja(pl.historial_turnos, g, fecha)
                ayer = _trabaja((pl.historial_turnos.get(g) or {}).get(pl.fecha))
                if hoy is not None and ayer is not None and hoy != ayer:
                    if hoy and p.faena.upper() == "D":
                        q.faena = "TN" if norm(p.tipo_turno) == "NOCHE" else "T"
                        motivo = f"Sube según su turno ({p.turno})"
                    elif not hoy and p.faena.upper() in ESTADOS_TRABAJANDO:
                        q.faena = "D"
                        motivo = f"Baja según su turno ({p.turno})"
            if p.rut in marcas and marcas[p.rut].estado:
                m = marcas[p.rut]
                antes = q.faena
                if m.estado in ("P", "T", "TN"):
                    if m.estado == "TN" or m.noche is True:
                        noche = True
                    elif m.estado == "T" or m.noche is False:
                        noche = False
                    elif p.rut in presentes:
                        noche = presentes[p.rut][0].turno == "Noche"
                    else:
                        noche = norm(p.tipo_turno) == "NOCHE"
                    if m.estado == "P" and p.rut not in presentes and p.faena.upper() in ("TT", "TC"):
                        q.faena = p.faena                      # asistió y seguía en teletrabajo / curso: se mantiene
                    else:
                        q.faena = "TN" if noche else "T"
                        q.tipo_turno = "NOCHE" if noche else "DIA"
                else:
                    q.faena = m.estado
                    if p.rut in presentes:
                        avisos.append(f"{p.nombre} ({p.rut}): tiene horas en {presentes[p.rut][0].archivo}, pero la tarja dice "
                                      f"«{m.texto or m.estado}»; se usa la tarja ({m.estado}).")
                if q.faena != antes or not motivo:
                    motivo = f"Tarja: «{m.texto or m.estado}»"
                if m.acreditado and norm(m.acreditado) != norm(q.acreditado):
                    cambios_acr.append(f"{p.nombre} ({q.acreditado or '–'} → {m.acreditado})")
                    q.acreditado = m.acreditado
            elif marcas and q.faena.upper() in ("T", "TN") and p.rut not in presentes:
                sin_marca.append(f"{p.nombre} ({p.faena}→{q.faena})")
            if p.rut in ajustes_estado and ajustes_estado[p.rut]:
                q.faena = ajustes_estado[p.rut]
                if q.faena == "TN":
                    q.tipo_turno = "NOCHE"
                elif q.faena in ("T", "TT", "TC") and norm(q.tipo_turno) == "NOCHE":
                    q.tipo_turno = "DIA"
                motivo = "Ajuste manual"
            if q.faena != p.faena:
                cambios.append(Cambio(p.rut, p.nombre, p.cargo, p.especialidad, p.faena, q.faena, motivo))
            personas.append(q)
    
        if cambios_acr:
            avisos.append(f"Acreditación tomada de la tarja: {len(cambios_acr)} cambio(s) respecto del día anterior: "
                          + "; ".join(cambios_acr[:20]) + (" …" if len(cambios_acr) > 20 else "") + ".")
        if sin_marca:
            avisos.append(f"{len(sin_marca)} persona(s) quedan en faena sin reporte y sin marca en la tarja (revísalas en la tabla "
                          "de cambios si no asistieron): " + "; ".join(sin_marca[:20]) + (" …" if len(sin_marca) > 20 else "") + ".")
    
        nuevos = []
        for rut, (rep, t) in presentes.items():
            if rut in plantel:
                continue
            noche = rep.turno == "Noche"
            nuevos.append({"rut": rut, "nombre": t.nombre, "cargo": t.cargo, "especialidad": esp_reporte[id(rep)],
                           "turno": "NOCHE" if noche else "DIA", "faena": "TN" if noche else "T", "archivo": rep.archivo})
        ya = {q.rut for q in personas}
        agregados = []
        for rut, m in nuevos_tarja:
            if rut in ya:
                continue
            f_ = m.ficha
            noche = m.noche if m.noche is not None else norm(f_.get("H")) == "NOCHE"
            est = m.estado or "D"
            faena = ("TN" if noche else "T") if est == "P" else est
            q = Persona(-1, rut, m.nombre or str(f_.get("I") or ""), str(f_.get("J") or ""), str(f_.get("K") or ""),
                        "NOCHE" if noche else str(f_.get("H") or "DIA"), str(f_.get("W") or "Directo"),
                        str(f_.get("AA") or m.acreditado or "SI"), str(f_.get("Z") or "CONTRATADO"), faena, str(f_.get("F") or ""),
                        ficha=f_)
            personas.append(q)
            ya.add(rut)
            agregados.append(f"{q.nombre} ({q.cargo})")
            nuevos = [n_ for n_ in nuevos if n_["rut"] != rut]
        if agregados:
            avisos.append(f"{len(agregados)} persona(s) contratadas según la tarja que no estaban en BD Personal se agregaron con "
                          "los datos de la tarja: " + "; ".join(agregados[:15]) + (" …" if len(agregados) > 15 else "") + ".")
        if nuevos:
            avisos.append(f"{len(nuevos)} persona(s) de los reportes no están en BD Personal del día anterior.")
        if agregar_nuevos:
            cargos_mod = {norm(fc.cargo) for fc in pl.cargos if fc.seccion == "MOD"}
            for nv in nuevos:
                personas.append(Persona(-1, nv["rut"], nv["nombre"], nv["cargo"], nv["especialidad"], nv["turno"],
                                        "Directo" if norm(nv["cargo"]) in cargos_mod else "Indirecto", "SI", "CONTRATADO", nv["faena"]))
    
        if sin_reporte:
            avisos.append("Especialidades con personal directo trabajando y sin reporte cargado (se mantiene su estado del día "
                          "anterior): " + ", ".join(f"{k} ({v})" for k, v in sorted(sin_reporte.items())) + ".")
    
        # Quien quedó sin trabajar (falla, licencia, permiso, descanso… por la tarja o un ajuste manual) no aporta horas:
        # se quitan del reporte de terreno las horas que tenía anotadas.
        estado_final = {q.rut: q.faena.upper() for q in personas}
        sin_horas: dict[str, tuple[str, str, float]] = {}
        for rut, (rep, t) in presentes.items():
            e = estado_final.get(rut)
            if e and e not in ESTADOS_TRABAJANDO and t.total:
                sin_horas[rut] = (plantel[rut].nombre if rut in plantel else t.nombre, e, t.total)
        if sin_horas:
            avisos.append(f"Se quitaron {_fmt_h(sum(v[2] for v in sin_horas.values()))} HH de los reportes de terreno de "
                          f"{len(sin_horas)} persona(s) que no trabajaron: " +
                          "; ".join(f"{n} ({ESTADOS.get(e, e)}, {_fmt_h(h)} h)" for n, e, h in sin_horas.values()) + ".")
    
        # Actividades para el texto del informe
        actividades = []
        for rep in reportes:
            if rep.tipo in ("uso_recurso", "adicionales"):
                continue                                      # sus trabajos van a «7.0 Uso de recurso»
            esp = esp_reporte[id(rep)]
            rep_udr = es_udr(f"{rep.archivo} {rep.area_reporte}")         # reporte completo de trabajos adicionales / UDR
            horas_por_letra: dict[str, float] = {}
            for t in rep.trabajadores:
                if t.rut in sin_horas:
                    continue
                for l, h in t.horas.items():
                    horas_por_letra[l] = horas_por_letra.get(l, 0) + h
            for a in rep.actividades:
                k = clave_area(a.area or rep.area_reporte, esp)
                ident = identificar_actividad(pl, a.descripcion, a.area or rep.area_reporte, getattr(a, "extra", ""), esp, fecha)
                marca_id = bool(re.search(r"PENDIENTE\s+UDR|\bUDR\b|TRABAJO ADICIONAL|\bNDC\b", norm(getattr(a, "extra", ""))))
                udr = rep.tipo == "terreno" and not es_generica(a.descripcion) and (marca_id or norm(ident.get("criterio") or "").startswith("EXTRAORDINARIA")
                                                 or es_udr(f"{a.descripcion} {a.area}", ident.get("activity_id"), pl)
                                                 or (rep_udr and a.descripcion and not _RE_NO_UDR.search(norm(a.descripcion))))
                dot, hh_j, hh_r = 0, 0.0, 0.0
                for t in rep.trabajadores:                     # dotación y HH (reportadas y jornada repartida) de esta actividad
                    h = t.horas.get(a.codigo, 0) if a.codigo else 0
                    tot = sum(t.horas.values())
                    if h and tot and t.rut not in sin_horas:
                        dot += 1
                        hh_j += HORAS_JORNADA * h / tot
                        hh_r += h
                enc = "Turno Noche" if rep.turno == "Noche" else encabezados.get(k, "")
                if not enc:
                    enc = encabezado_sugerido(pl, a.descripcion, f"{a.area or rep.area_reporte} {getattr(a, 'extra', '')}", esp,
                                              ident.get("activity_id"))
                actividades.append({"incluir": bool(a.descripcion) and rep.tipo != "programacion", "encabezado": enc, "descripcion": a.descripcion.strip(),
                                    "especialidad": esp, "area": a.area, "hh": horas_por_letra.get(a.codigo, 0.0),
                                    "turno": rep.turno, "clave": k, "archivo": rep.archivo, "codigo": a.codigo, **ident,
                                    "udr": udr, "dotacion": dot, "hh_jornada": round(hh_j, 1), "hh_reportadas": round(hh_r, 1),
                                    "criterio_sug": ident.get("criterio") or
                                                    criterio_sugerido(pl, f"{a.descripcion} {a.area} {getattr(a, 'extra', '')}", udr)})
                if udr:
                    actividades[-1]["incluir"] = False             # va a «Uso de recurso / NDC», no a la descripción
    
        # Filas de BD Actividades MOD: todo directo en faena
        filas_act = []
        for q in personas:
            if norm(q.moimod) != "DIRECTO" or q.faena.upper() not in ("T", "TN"):
                continue
            fila = {"rut": q.rut, "persona": q, "supervisor": "", "capataz": "", "area_trabajo": "", "sub_area": "",
                    "actividad": "", "hh": float(HORAS_JORNADA)}
            if q.rut in presentes:
                rep, t = presentes[q.rut]
                desc = {a.codigo: a.descripcion for a in rep.actividades}
                capataz = next((x.nombre for x in rep.trabajadores if norm(x.cargo).startswith("CAPATAZ")), "")
                fila.update(supervisor=rep.supervisor, capataz=capataz, area_trabajo=rep.area_reporte,
                            actividad="\n".join(f"{desc.get(l, l)} ({_fmt_h(h)} h)" for l, h in t.horas.items()),
                            detalle=[(rep.archivo, l, h, desc.get(l, l)) for l, h in t.horas.items() if h],
                            tipo_reporte=rep.tipo)
                # Como la planilla oficial: cada persona en faena cuenta una jornada completa (HORAS_JORNADA), aunque el
                # reporte de terreno anote menos horas; las horas por actividad quedan en el texto de la actividad.
            filas_act.append(fila)
        hh_precom = sum(f["hh"] for f in filas_act if "PRECOM" in norm(f["persona"].cargo))
        hh_constr = sum(f["hh"] for f in filas_act) - hh_precom
    
        totales = calcular_totales(pl, personas)
        if totales["sin_fila"]:
            avisos.append("Cargos que no tienen fila en el informe (no se cuentan en él): " + ", ".join(totales["sin_fila"]) + ".")
        uso_recurso, vistos = [], set()
        def _textos(u):
            return [set(_tokens_p6(t)) for t in (u.get("descripcion"), u.get("detalle")) if t]
        def _parecido(u1, u2):
            return any(a and b and len(a & b) / len(a | b) >= 0.6 for a in _textos(u1) for b in _textos(u2))
        # primero los formularios UDR (traen RUT), después las tablas de adicionales
        for u in sorted((u for rep in reportes for u in rep.uso_recurso), key=lambda u: 0 if u.get("ruts") else 1):
            if True:
                k = (norm(u.get("descripcion")), float(u.get("directos") or 0), float(u.get("hh_directos") or 0),
                     tuple(sorted(r[0] for r in u.get("ruts", []))) or norm(u.get("detalle")))
                dup = next((x for x in uso_recurso if _parecido(x, u)
                            and not (x.get("ruts") and u.get("ruts") and {r[0] for r in x["ruts"]} != {r[0] for r in u["ruts"]}) and
                            abs(float(x.get("hh_directos") or 0) + float(x.get("hh_indirectos") or 0)
                                - float(u.get("hh_directos") or 0)) < 0.6), None)
                if k in vistos or dup:
                    avisos.append(f"Uso de recurso repetido en {u.get('origen')}: «{u.get('detalle') or u.get('descripcion')}» "
                                  "(se considera una vez).")
                    continue
                vistos.add(k)
                ident = identificar_actividad(pl, f"{u.get('descripcion') or ''} {u.get('detalle') or ''}", u.get("area") or "",
                                              u.get("respaldo") or "", u.get("especialidad") or "", fecha)
                u = {**u}
                if u.get("ruts"):                              # directos / indirectos según BD Personal (MOD / MOI)
                    d_, i_ = [], []
                    for rut, hh, cargo in u["ruts"]:
                        p_ = plantel.get(rut)
                        es_ind = (p_ is not None and norm(p_.moimod) == "INDIRECTO") or bool(re.search(
                            r"SUPERV|PREVENC|JEFE|ADMINIST|BODEG|ENCARGADO|CONTROL DOC|CONDUCTOR|CHOFER|OPERADOR|PLANIFIC",
                            norm(cargo)))
                        (i_ if es_ind else d_).append(hh)
                    u.update(directos=len(d_) or None, hh_directos=sum(d_) or None, indirectos=len(i_) or None,
                             hh_indirectos=sum(i_) or None)
                texto_u = f"{u.get('descripcion') or ''} {u.get('detalle') or ''}"
                # convención de la sección UDR del informe; manda el título del formulario («Precom y protocolos» -> Precom)
                u["area"] = area_udr(u.get("descripcion") or "") or area_udr(u.get("detalle") or "", u.get("area") or "")
                u["respaldo"] = u.get("respaldo") or respaldo_udr(texto_u)
                uso_recurso.append(u)
        # Actividades marcadas como adicionales / UDR en los reportes de terreno -> uso de recurso (sin duplicar)
        adicional = next((c for c in pl.criterios if norm(c).startswith("EXTRAORDINARIA")), "Extraordinaria / Adicional")
        for a in actividades:
            if not a.get("udr"):
                continue
            area_a = area_udr(f"{a['descripcion']} {a.get('area') or ''}")
            dup = next((u for u in uso_recurso if not u.get("desde_actividad") and _parecido(u, {"descripcion": a["descripcion"]})
                        and (not u.get("area") or not area_a or norm(u["area"]) == norm(area_a))), None)
            if dup:
                if not dup.get("directos") and a["dotacion"]:
                    dup.update(directos=a["dotacion"], hh_directos=a["hh_reportadas"] or a["hh_jornada"])
                continue
            uso_recurso.append({"descripcion": a["descripcion"], "detalle": a["descripcion"], "directos": a["dotacion"] or None,
                                "hh_directos": a["hh_reportadas"] or a["hh_jornada"] or None,
                                "especialidad": _especialidad_udr([], [a["especialidad"]], a["especialidad"]),
                                "area": area_udr(f"{a['descripcion']} {a.get('area') or ''}", a.get("area") or ""),
                                "turno": a["turno"],
                                "origen": a["archivo"], "respaldo": respaldo_udr(a["descripcion"]), "criterio": adicional,
                                "desde_actividad": True})
        for u in uso_recurso:
            u.setdefault("criterio", adicional)
        for rep in reportes:
            for t in rep.trabajadores if rep.tipo == "terreno" else []:
                if t.rut in sin_horas:
                    continue
                if t.total and abs(t.total - HORAS_JORNADA) > 0.01:
                    avisos.append(f"{t.nombre} ({rep.archivo}) registra {_fmt_h(t.total)} h (jornada de {HORAS_JORNADA} h).")
            for av in rep.avisos:
                avisos.append(f"{rep.archivo}: {av}")
        # Materiales declarados en los formularios UDR
        materiales = []
        for rep in reportes:
            for m in getattr(rep, "materiales", []):
                titulo = next((u.get("descripcion") for u in rep.uso_recurso if u.get("descripcion")), "")
                base = {"trabajo": titulo, "solicitado": rep.supervisor or "", "observacion": rep.archivo}
                materiales.append({**base, **{k: v for k, v in m.items() if v not in ("", None)}})
        # Interferencias y detenciones: actividades de los reportes clasificadas así (texto o historial)
        interferencias = []
        tipos_int = pl.listas.get("interferencia") or []
        for a in actividades:
            crit = norm(a.get("criterio_sug") or "")
            if not (crit.startswith("INTERFERENCIA") or crit.startswith("DETENCION")):
                continue
            t_ = norm(a["descripcion"])
            horas = re.findall(r"(\d{1,2})[:.](\d{2})\s*(?:HRS?|AM|PM)?", t_)
            if re.search(r"VIENTO|LLUVIA|TORMENTA|ALERTA|CLIMA", t_):
                buscado = "CLIMA"
            elif "MANDANTE" in crit or re.search(r"GASCO|MANDANTE|CLIENTE", t_):
                buscado = "CLIENTE"
            elif re.search(r"CONTRATISTA|EMPRESA|TERCERO", t_):
                buscado = "TERCERO"
            else:
                buscado = "CONDICION"
            tipo = next((x for x in tipos_int if norm(x) == buscado), tipos_int[0] if tipos_int else "")
            igual = next((x for x in interferencias if norm(x["detalle"]) == t_), None)
            if igual:
                igual["dotacion"] = (igual["dotacion"] or 0) + (a.get("dotacion") or 0)
                igual["hh"] = (igual["hh"] or 0) + (a.get("hh_reportadas") or 0)
                continue
            interferencias.append({"detalle": a["descripcion"], "area": area_udr(a["descripcion"], a.get("area") or ""),
                                   "equipo": a.get("especialidad") or "",
                                   "inicio": f"{horas[0][0]}:{horas[0][1]}" if horas else "",
                                   "termino": f"{horas[1][0]}:{horas[1][1]}" if len(horas) > 1 else "",
                                   "horas": "", "dotacion": a.get("dotacion") or None,
                                   "hh": a.get("hh_reportadas") or None, "tipo": tipo, "origen": a["archivo"]})
        # Resumen clasificación: cursos (habilitación), lluvia y alertas de viento (personas y HH reportadas)
        clasif = {"cursos": [set(), 0.0], "lluvia": [set(), 0.0], "alerta1": [set(), 0.0], "alerta2": [set(), 0.0]}
        for rep in reportes:
            if rep.tipo != "terreno":
                continue
            desc = {a.codigo: norm(f"{a.descripcion} {getattr(a, 'extra', '')}") for a in rep.actividades if a.codigo}
            for t in rep.trabajadores:
                for l, h in t.horas.items():
                    d_ = desc.get(l, "")
                    k = ("cursos" if re.search(r"\bCURSO|INDUCCION|CAPACITACION|HABILITACION|HIPOBARIA|INHABILITAD", d_) else
                         "alerta2" if re.search(r"ALERTA.*(NIVEL\s*2|N\s*2\b|\b2\b)", d_) else
                         "alerta1" if re.search(r"ALERTA|VIENTO", d_) else
                         "lluvia" if re.search(r"LLUVIA|PRECIPITACION|TORMENTA ELECTRICA", d_) else None)
                    if k and h:
                        clasif[k][0].add(t.rut or t.nombre)
                        clasif[k][1] += h
        for a in actividades:                                  # habilitación según el historial (p. ej. «personal inhabilitado»)
            if norm(a.get("criterio_sug") or "").startswith("HABILITACION") and not clasif["cursos"][0]:
                clasif["cursos"][1] += a.get("hh_reportadas") or 0
                clasif["cursos"][0].update(f"{a['archivo']}#{i}" for i in range(a.get("dotacion") or 0))
        clasificacion = {k: (len(v[0]), round(v[1], 1)) for k, v in clasif.items()}
        return Consolidado(fecha, personas, cambios, nuevos, actividades, filas_act, hh_constr, hh_precom, totales, avisos,
                           sorted({esp_reporte[id(r)] for r in reportes if r.tipo == "terreno"}), sin_reporte, uso_recurso,
                           materiales, interferencias, clasificacion)
    
    
    def _fmt_h(h: float) -> str:
        return str(int(h)) if float(h).is_integer() else f"{h:.1f}".replace(".", ",")
    
    
    def texto_actividades(actividades: list[dict], nota: str = "") -> str:
        """Arma el texto de la sección 1.0 con el mismo estilo del informe: encabezado y numeración por sistema."""
        grupos: dict[str, list[str]] = {}
        orden: list[str] = []
        for a in actividades:
            if not a.get("incluir") or not str(a.get("descripcion", "")).strip():
                continue
            enc = (a.get("encabezado") or "Sin sistema asignado").strip().rstrip(":")
            if enc not in grupos:
                grupos[enc] = []
                orden.append(enc)
            d = re.sub(r"\s+", " ", str(a["descripcion"])).strip()
            if d and d not in grupos[enc]:
                grupos[enc].append(d)
        # "Turno Noche" va al final, como en el informe
        orden = [e for e in orden if norm(e) != "TURNO NOCHE"] + [e for e in orden if norm(e) == "TURNO NOCHE"]
        lineas = []
        for enc in orden:
            lineas.append(f"{enc}:")
            lineas += [f"{i}- {d}" for i, d in enumerate(grupos[enc], start=1)]
        if nota.strip():
            lineas.append("Nota:")
            lineas += [f"{i}- {d.strip()}" for i, d in enumerate([x for x in nota.splitlines() if x.strip()], start=1)]
        return "\n".join(lineas)
    
    
    def encabezados_previos(texto: str) -> list[str]:
        salida = []
        for linea in texto.splitlines():
            l = linea.strip()
            if l.endswith(":") and not re.match(r"^\d+\s*-", l) and norm(l) not in ("NOTA:", "TURNO NOCHE:"):
                salida.append(l[:-1].strip())
        return list(dict.fromkeys(salida))
    
    
    # ====================================================================== generación del informe
    
    @dataclass
    class Entradas:
        fecha: date
        numero: int
        condicion: str
        texto_actividades: str
        clasificacion: dict[str, tuple[float | None, float | None]] = field(default_factory=dict)  # cursos/lluvia/alerta1/alerta2
        resguardo_hoy: float = 0.0
        equipos: dict[int, dict] = field(default_factory=dict)       # fila -> {"dia":, "noche":, "observacion":}
        interferencias: list[dict] = field(default_factory=list)
        uso_recurso: list[dict] = field(default_factory=list)
        fotos: list[tuple[bytes | None, str]] = field(default_factory=list)   # [(imagen, leyenda)] en el orden de las fotos
        fotos_en_blanco: bool = True        # las fotos no reemplazadas se dejan en blanco (no se repiten las del día anterior)
        # Bloque de firmas al final del informe: [{"rol": "Elaborado por", "nombre":, "cargo":, "firma": png|None, "fecha":}]
        firmas: list[dict] = field(default_factory=list)
        # Formato nuevo: Activity ID y criterio de cada actividad de los reportes {"archivo|código": {"id":, "criterio":}}
        asignacion: dict[str, dict] = field(default_factory=dict)
        # BD Maquinaria: cambios por patente {"VPXP76": {"estado": "MANTENCION", "observacion": "...", ...}}
        maquinaria: dict[str, dict] = field(default_factory=dict)
        materiales: list[dict] = field(default_factory=list)          # 8.0 Materiales y consumibles
        validacion: dict = field(default_factory=dict)                 # {"estado": ..., "observacion": ...}
    
    
    def valores_cierre(pl: Plantilla) -> dict:
        """Acumulados con los que cierra el informe base (son el 'anterior' del nuevo informe)."""
        v = pl.valores
        t = pl.totales
    
        def n(x):
            return float(x) if isinstance(x, (int, float)) else 0.0
        return {
            "hh_constr": n(v.get("hh_constr_ant")) + n(v.get("hh_constr_act")),
            "hh_precom": n(v.get("hh_precom_ant")) + n(v.get("hh_precom_act")),
            "hh_moi": n(v["moi_ant"].get("HH", v["moi_ant"].get("X"))) + t["MOI"]["X"],
            "hh_mod_x": n(v["mod_ant"].get("HH", v["mod_ant"].get("X"))) + t["MOD"]["X"],
            "resguardo": n(v.get("resguardo_acum")),
            "MOD": t["MOD"], "MOI": t["MOI"],
        }
    
    
    def ampliar_secciones(pl: Plantilla, n_udr: int = 0, n_mat: int = 0) -> Plantilla:
        """Si hay más UDR o más materiales que filas en sus secciones del informe, inserta las filas que faltan
        (con su formato y fórmulas, corriendo todo lo de abajo) y vuelve a leer la planilla."""
        falta_u = n_udr - len(pl.filas_udr) if pl.filas_udr else 0
        filas_m = (pl.materiales or {}).get("filas") or []
        falta_m = n_mat - len(filas_m) if filas_m else 0
        if falta_u <= 0 and falta_m <= 0:
            return pl
        L = X.Libro(pl.datos)
        hecho = False
        # de abajo hacia arriba, para que las filas de la sección de arriba sigan siendo válidas
        for fila, falta in sorted([(filas_m[-1] if filas_m else 0, falta_m), (pl.filas_udr[-1] if pl.filas_udr else 0, falta_u)],
                                  reverse=True):
            if falta > 0 and fila:
                hecho = X.insertar_filas(L, HOJA_INFORME, fila, falta) or hecho
        if not hecho:
            return pl
        nueva = leer_plantilla(L.bytes(recalcular=True))
        if len(nueva.filas_udr) < len(pl.filas_udr) or len((nueva.materiales or {}).get("filas") or []) < len(filas_m):
            return pl
        return nueva
    
    
    def ampliar_udr(pl: Plantilla, total: int) -> Plantilla:
        return ampliar_secciones(pl, n_udr=total)
    
    
    def generar_informe(pl: Plantilla, cons: Consolidado, ent: Entradas) -> bytes:
        pl = ampliar_secciones(pl, len(ent.uso_recurso), len(ent.materiales))
        L = X.Libro(pl.datos)
        cierre = valores_cierre(pl)
        r = pl.refs
        v: dict[str, object] = {
            r["fecha"]: ent.fecha, r["numero"]: ent.numero, r["condicion"]: ent.condicion,
            r["actividades"]: ent.texto_actividades,
        }
        # Resumen de personal: "cantidad anterior" = conteos del informe base
        mod, moi = cierre["MOD"], cierre["MOI"]
        ant = {"mod_acred": mod["O"], "moi_acred": moi["O"], "mod_noacred": mod["L"] - mod["O"], "moi_noacred": moi["L"] - moi["O"],
               "mod_total": mod["L"], "moi_total": moi["L"]}
        ant["proy_total"] = ant["mod_total"] + ant["moi_total"]
        for clave, val in ant.items():
            if clave in r:
                v[f"AB{r[clave]}"] = val
        if "mod_acred" in r:
            v[f"AB{int(r['mod_acred']) + 2}"] = ant["mod_acred"] + ant["moi_acred"]       # subtotal acreditados
        if "mod_noacred" in r:
            v[f"AB{int(r['mod_noacred']) + 2}"] = ant["mod_noacred"] + ant["moi_noacred"]
        # HH gastadas
        hh_mod_ant = cierre["hh_constr"] + cierre["hh_precom"]
        resumen = {f"AB{r['hh_constr']}": cierre["hh_constr"], f"AD{r['hh_constr']}": cons.hh_constr,
                   f"AB{r['hh_precom']}": cierre["hh_precom"], f"AD{r['hh_precom']}": cons.hh_precom,
                   f"AB{r['hh_mod']}": hh_mod_ant}
        if "hh_moi" in r:
            resumen[f"AB{r['hh_moi']}"] = cierre["hh_moi"]
        if "hh_total" in r:
            resumen[f"AB{r['hh_total']}"] = hh_mod_ant + (cierre["hh_moi"] if "hh_moi" in r else 0)
        if "resguardo" in r:
            resumen[f"AB{r['resguardo']}"] = cierre["resguardo"]
            resumen[f"AF{r['resguardo']}"] = cierre["resguardo"] + (ent.resguardo_hoy or 0)
        # las celdas que en la planilla son fórmula (formato nuevo: HH del período = suma de la tabla de personal) se respetan
        v.update({k: x for k, x in resumen.items() if k not in pl.con_formula})
        for clave in ("cursos", "lluvia", "alerta1", "alerta2"):
            if clave in r:
                cant, hh = ent.clasificacion.get(clave, (None, None))
                v[f"AB{r[clave]}"] = cant if cant else None
                v[f"AD{r[clave]}"] = hh if hh else None
        # "Total informe anterior" en las tablas de personal
        for sec, tot in (("mod", mod), ("moi", moi)):
            fila = r[f"{sec}_ant"]
            for c in "LORU":
                v[f"{c}{fila}"] = tot[c]
        v[f"{pl.col_hh}{r['mod_ant']}"] = cierre["hh_mod_x"]
        v[f"{pl.col_hh}{r['moi_ant']}"] = cierre["hh_moi"]
        if pl.col_hh != "X":                                  # formato nuevo: X = total personas en faena
            v[f"X{r['mod_ant']}"] = mod["R"] + mod["U"]
            v[f"X{r['moi_ant']}"] = moi["R"] + moi["U"]
        # Equipos: HM anterior = acumulado del informe base; día/noche según lo indicado
        maq = maquinaria_del_dia(pl, ent.maquinaria) if pl.maquinaria else []
        por_patente = {_clave_patente(d["patente"]): d for d in maq}
        bajas = {_clave_patente(d["patente"]) for d in pl.maquinaria} - set(por_patente)
        for e in pl.equipos:
            cambio = ent.equipos.get(e.fila, {})
            dia, noche = cambio.get("dia", e.dia), cambio.get("noche", e.noche)
            pk = _clave_patente(e.patente)
            if pk and pk in por_patente:                       # equipo con patente: manda BD Maquinaria (siempre calzan)
                d = por_patente[pk]
                dia = 1 if operando_dia(d) else 0
                if not cuenta_equipo(d):
                    noche = 0
                if norm(d.get("estado")) not in ("OPERANDO", "") and (cambio.get("observacion") or "") == (e.observacion or ""):
                    cambio = {**cambio, "observacion": d["estado"].capitalize() + (f": {d['observacion']}" if d.get("observacion") else "")}
            elif pk and pk in bajas:
                dia = noche = 0
                cambio = {**cambio, "observacion": f"Desmovilizado {ent.fecha:%d-%m-%Y}"}
                if f"L{e.fila}" not in pl.con_formula:
                    v[f"L{e.fila}"] = None
            v[f"U{e.fila}"] = dia if dia else None
            if f"{e.col_noche}{e.fila}" not in pl.con_formula:   # la planilla la calcula desde BD Maquinaria
                v[f"{e.col_noche}{e.fila}"] = noche if noche else None
            if e.hm_anterior is not None:
                v[f"Y{e.fila}"] = e.hm_anterior + (e.dia + e.noche) * HORAS_JORNADA
            if "observacion" in cambio:
                v[f"AB{e.fila}"] = cambio["observacion"] or None
        # 8.0 Materiales y consumibles
        if pl.materiales:
            cm = pl.materiales["cols"]
            for i, fila in enumerate(pl.materiales["filas"]):
                d = ent.materiales[i] if i < len(ent.materiales) else {}
                for k, col in cm.items():
                    val = d.get(k)
                    v[f"{col}{fila}"] = val if val not in ("", None) else None
        # 8.0 Validación: estado y observaciones
        if pl.validacion.get("fila_estado") and ent.validacion:
            fe = pl.validacion["fila_estado"]
            if ent.validacion.get("estado"):
                v[f"L{fe}"] = ent.validacion["estado"]
            if pl.validacion.get("col_obs") and ent.validacion.get("observacion"):
                v[f"{pl.validacion['col_obs']}{fe}"] = f"{pl.validacion.get('rotulo_obs') or 'Observaciones:'} {ent.validacion['observacion']}"
        # Interferencias
        for i, fila in enumerate(pl.filas_interferencia):
            d = ent.interferencias[i] if i < len(ent.interferencias) else {}
            for col, k in (("B", "detalle"), ("L", "area"), ("O", "equipo"), ("R", "inicio"), ("U", "termino"), ("X", "horas"),
                           ("AB", "dotacion"), ("AE", "hh"), ("AG", "tipo")):
                v[f"{col}{fila}"] = d.get(k) if d.get(k) not in ("", None) else None
        # Uso de recurso (NDC)
        if pl.udr_formulas:
            # Como en la planilla oficial: las filas van a la hoja BD UDR y el informe las toma con fórmulas
            primera = _agregar_udr(L, pl, ent)
            f0 = pl.filas_udr[0]
            for i, fila in enumerate(pl.filas_udr):
                d = ent.uso_recurso[i] if i < len(ent.uso_recurso) else {}
                ref_ = f"'BD UDR'!C{primera + i}"
                v[f"B{fila}"] = (f'=IF({ref_}="","",{ref_})', d.get("ndc") if d else "")
                clave_col = {c: k for k, c in pl.udr_cols.items()}
                for col, fx in pl.udr_formulas.items():
                    if col == "B":
                        continue
                    ret = re.findall(r"'BD UDR'!\$?([A-Z]+):", fx)
                    k = clave_col.get(ret[-1]) if ret else None
                    val = d.get(k) if k and d.get(k) not in ("", None) else None
                    fx_ = re.sub(rf"\bB{f0}\b", f"B{fila}", fx)
                    if not fx_.upper().startswith("IFERROR("):        # filas sin NDC quedan en blanco (no #N/A)
                        fx_ = f'IFERROR({fx_},"")'
                    v[f"{col}{fila}"] = ("=" + fx_, val if d else "")
        else:
            for i, fila in enumerate(pl.filas_udr):
                d = ent.uso_recurso[i] if i < len(ent.uso_recurso) else {}
                for col, k in (("B", "ndc"), ("D", "descripcion"), ("O", "directos"), ("S", "hh_directos"), ("U", "indirectos"),
                               ("Y", "hh_indirectos"), ("AB", "equipos"), ("AG", "hm")):
                    v[f"{col}{fila}"] = d.get(k) if d.get(k) not in ("", None) else None
        # Fotos y leyendas
        fotos = list(ent.fotos) + [(None, "")] * max(0, len(pl.fotos) - len(ent.fotos))
        if ent.fotos_en_blanco:
            fotos = [(img if img else _imagen_blanca(), ley) for img, ley in fotos]
        for i, f in enumerate(pl.fotos):
            if f.get("leyenda"):
                img, ley = fotos[i]
                if ent.fotos_en_blanco or img or ley:
                    v[f["leyenda"]] = ley or None
    
        fijos = {}
        if pl.formato == "master":
            filas_master = filas_master_actividades(pl, cons, ent)
            # sección 1.1: N° actividades (Activity ID distintos), dotación y HH por criterio (caché de sus fórmulas)
            for fila, crit in pl.filas_tipo_actividad.items():
                sel = [m for m in filas_master if norm(m["criterio"]) == norm(crit)]
                fijos[f"L{fila}"] = len({m["id"] for m in sel if m["id"]})
                fijos[f"O{fila}"] = len(sel)
                fijos[f"R{fila}"] = round(sum(m["hh"] for m in sel), 2)
        xi = X.fijar_celdas(L.hoja_xml(HOJA_INFORME), v)
        L.guardar_hoja(HOJA_INFORME, _cachear_resultados(xi, L, cons, fijos))
        _agregar_personal(L, pl, cons, ent.fecha)
        if pl.maquinaria:
            _agregar_maquinaria(L, pl, ent)
        if pl.formato == "master":
            _agregar_master(L, pl, cons, ent.fecha, filas_master)
        else:
            _agregar_actividades(L, pl, cons, ent.fecha)
        _reemplazar_fotos(L, pl, fotos)
        if ent.firmas:
            if pl.validacion:
                _firmas_en_validacion(L, pl, ent.firmas)
            else:
                _agregar_firmas(L, ent.firmas)
        return L.bytes(recalcular=True)
    
    
    # ====================================================================== firmas
    
    def preparar_firma(datos: bytes) -> bytes:
        """Imagen de firma lista para el informe: fondo blanco a transparente, recortada a la tinta y en PNG."""
        from PIL import Image, ImageOps
        img = ImageOps.exif_transpose(Image.open(io.BytesIO(datos))).convert("RGBA")
        if max(img.size) > 1600:
            img.thumbnail((1600, 1600))
        gris = img.convert("L")
        px, pg = img.load(), gris.load()
        for y in range(img.height):
            for x in range(img.width):
                r, g, b, a = px[x, y]
                claro = pg[x, y]
                if a < 20 or claro > 200:                   # papel o fondo: transparente
                    px[x, y] = (255, 255, 255, 0)
                else:                                       # tinta: se oscurece un poco para que se lea al imprimir
                    k = max(0, min(255, int((claro - 40) * 0.8)))
                    px[x, y] = (min(r, k + 20), min(g, k + 20), min(b, k + 40), 255)
        caja = img.getchannel("A").getbbox()
        if not caja:
            raise ValueError("No se distingue una firma en la imagen. Usa una foto o escaneo con fondo claro y tinta oscura.")
        img = img.crop(caja)
        m = 8
        lienzo = Image.new("RGBA", (img.width + 2 * m, img.height + 2 * m), (255, 255, 255, 0))
        lienzo.paste(img, (m, m))
        if lienzo.width > 700:
            lienzo = lienzo.resize((700, int(lienzo.height * 700 / lienzo.width)))
        buf = io.BytesIO()
        lienzo.save(buf, "PNG", optimize=True)
        return buf.getvalue()
    
    
    _COLS_FIRMA = [("B", "K", 1), ("O", "X", 14), ("Y", "AI", 24)]   # columnas (inicio, fin, índice 0-based de la columna inicial)
    _EMU_CM = 360000
    
    
    def _agregar_firmas(L: X.Libro, firmas: list[dict]):
        """Bloque «FIRMAS» bajo el área de impresión del informe (o reutiliza el que ya existe) con nombre, cargo, fecha e imagen."""
        xi = L.hoja_xml(HOJA_INFORME)
        wb = L.leer("xl/workbook.xml")
        m = re.search(r"(<definedName[^>]*_xlnm\.Print_Area[^>]*>'?" + re.escape(HOJA_INFORME) + r"'?!\$A\$1:\$([A-Z]+)\$)(\d+)(</definedName>)", wb)
        fin_impresion = int(m.group(3)) if m else max(n for n, *_ in X.filas(xi))
        # ¿Ya hay un bloque de firmas (informe generado antes desde esta misma base)?
        fila0 = None
        for n, a, fx, _, _ in X.filas(xi):
            if "FIRMAS DEL INFORME" in fx:
                fila0 = n
                break
        if fila0 is None:
            fila0 = fin_impresion + 2
        estilo_titulo = estilo_texto = None
        for n, a, fx, _, _ in X.filas(xi):
            if n == 6:
                cs = X.celdas(fx)
                estilo_titulo = cs.get("D", ({},))[0].get("s")
                estilo_texto = cs.get("H", ({},))[0].get("s")
                break
        v = {f"B{fila0}": "FIRMAS DEL INFORME"}
        for k in (2, 3, 4):
            v[f"AJ{fila0 + k}"] = " "            # crea las filas del espacio de la firma (fuera del área visible del bloque)
        for i, (c1, c2, _) in enumerate(_COLS_FIRMA):
            f = firmas[i] if i < len(firmas) else None
            v[f"{c1}{fila0 + 1}"] = (f or {}).get("rol")
            v[f"{c1}{fila0 + 5}"] = (f or {}).get("nombre")
            v[f"{c1}{fila0 + 6}"] = (f or {}).get("cargo")
            fe = (f or {}).get("fecha")
            v[f"{c1}{fila0 + 7}"] = (f"Fecha: {fe:%d-%m-%Y %H:%M}" if hasattr(fe, "strftime") else fe) if f else None
        xi = X.fijar_celdas(xi, v)
        # estilos, alto de filas y celdas combinadas del bloque
        for ref in [r for r, val in v.items() if val not in (None, " ")]:
            fila_ref = int(re.sub(r"[A-Z]", "", ref))
            s_ = estilo_titulo if fila_ref in (fila0, fila0 + 1) else estilo_texto
            if s_:
                xi = re.sub(rf'<c r="{ref}"(?: s="\d+")?', f'<c r="{ref}" s="{s_}"', xi, count=1)
        for k in (2, 3, 4):
            xi = _alto_fila(xi, fila0 + k, 26)
        usados = _COLS_FIRMA[:max(1, len(firmas))]
        nuevas = [f"B{fila0}:AI{fila0}"] + [f"{c1}{fila0 + k}:{c2}{fila0 + k}" for c1, c2, _ in usados for k in (1, 5, 6, 7)]
        existentes = set(re.findall(r'<mergeCell ref="([^"]+)"', xi))
        nuevas = [r for r in nuevas if r not in existentes]
        if nuevas:
            if "<mergeCells" in xi:
                xi = re.sub(r'<mergeCells count="(\d+)">', lambda mm: f'<mergeCells count="{int(mm.group(1)) + len(nuevas)}">', xi, count=1)
                xi = xi.replace("</mergeCells>", "".join(f'<mergeCell ref="{r}"/>' for r in nuevas) + "</mergeCells>", 1)
            else:
                xi = xi.replace("</sheetData>", "</sheetData><mergeCells count=\"%d\">%s</mergeCells>" %
                                (len(nuevas), "".join(f'<mergeCell ref="{r}"/>' for r in nuevas)), 1)
        L.guardar_hoja(HOJA_INFORME, X.actualizar_dimension(xi))
        if m:   # el bloque queda dentro del área de impresión (y del PDF)
            nuevo_fin = max(int(m.group(3)), fila0 + 8)
            L.escribir("xl/workbook.xml", wb.replace(m.group(0), f"{m.group(1)}{nuevo_fin}{m.group(4)}", 1))
    
        alto_caja = 3 * 26 * 12700 - 70000
        slots = []
        for i, f in enumerate(firmas[:len(_COLS_FIRMA)]):
            if f.get("firma"):
                c_ini = _COLS_FIRMA[i][2]
                slots.append((i, f["firma"], c_ini, c_ini + 5, fila0 + 1, 70000, fila0 + 4, alto_caja))
        _poner_imagenes_firma(L, xi, slots)
    
    
    def _alto_fila(xi: str, fila: int, alto: float) -> str:
        def f(m):
            tag = re.sub(r'\s(?:ht|customHeight)="[^"]*"', "", m.group(0))
            return tag[:-1] + f' ht="{alto}" customHeight="1">' if not tag.endswith("/>") else tag[:-2] + f' ht="{alto}" customHeight="1"/>'
        return re.sub(rf'<row r="{fila}"[^>]*?/?>', f, xi, count=1)
    
    
    def _poner_imagenes_firma(L: X.Libro, xi: str, slots: list[tuple]):
        """slots: (índice, png, col inicial 0-based, col final, fila desde 0-based, desplazamiento EMU, fila hasta, alto EMU)."""
        rels_hoja = L.hojas[HOJA_INFORME].replace("worksheets/", "worksheets/_rels/") + ".rels"
        rh = L.leer(rels_hoja)
        md = re.search(r'Target="\.\./drawings/(drawing\d+\.xml)"', rh)
        if not md:
            return
        dib = f"xl/drawings/{md.group(1)}"
        rels_dib = f"xl/drawings/_rels/{md.group(1)}.rels"
        xd, rd = L.leer(dib), L.leer(rels_dib)
        for etiqueta in ("oneCellAnchor", "twoCellAnchor"):     # firmas de una generación anterior sobre esta misma base
            xd = re.sub(rf'<xdr:{etiqueta}\b[^>]*>(?:(?!</xdr:{etiqueta}>).)*?name="Firma SYNCORE[^"]*"(?:(?!</xdr:{etiqueta}>).)*</xdr:{etiqueta}>',
                        "", xd, flags=re.S)
        ids = [int(x) for x in re.findall(r'<xdr:cNvPr id="(\d+)"', xd)] or [1]
        from PIL import Image
        # ancho de columnas (para que la imagen no se deforme al ajustarse al recuadro de celdas)
        anchos = {}
        for mm in re.finditer(r'<col\b[^>]*min="(\d+)"[^>]*max="(\d+)"[^>]*width="([\d.]+)"', xi):
            for k in range(int(mm.group(1)), int(mm.group(2)) + 1):
                anchos[k] = float(mm.group(3))
        defecto = float((re.search(r'defaultColWidth="([\d.]+)"', xi) or [None, "8.43"])[1])
        col_emu = lambda k: int((anchos.get(k + 1, defecto) * 7 + 5) * 9525)     # k: índice 0-based
        for i, png, c_ini, c_fin, f_desde, off, f_hasta, alto_caja in slots:
            ancho_caja = sum(col_emu(k) for k in range(c_ini, c_fin))
            img = Image.open(io.BytesIO(png)).convert("RGBA")
            # se escala dentro del recuadro y se completa con transparente hasta su proporción
            esc = min(ancho_caja * 0.95 / img.width, alto_caja * 0.9 / img.height)
            w, h = int(img.width * esc), int(img.height * esc)
            lienzo_w = 900
            lienzo_h = max(1, int(lienzo_w * alto_caja / ancho_caja))
            k_ = lienzo_w / ancho_caja
            firma_img = img.resize((max(1, int(w * k_)), max(1, int(h * k_))))
            lienzo = Image.new("RGBA", (lienzo_w, lienzo_h), (255, 255, 255, 0))
            lienzo.paste(firma_img, (int(0.02 * lienzo_w), lienzo_h - firma_img.height - int(0.04 * lienzo_h)), firma_img)
            buf = io.BytesIO()
            lienzo.save(buf, "PNG", optimize=True)
            nombre_media = f"firma_syncore_{i + 1}_{abs(hash(png)) % 10**8}.png"
            L.escribir(f"xl/media/{nombre_media}", buf.getvalue())
            rid = f"rIdFirma{i + 1}"
            rd = re.sub(rf'<Relationship[^>]*Id="{rid}"[^>]*/>', "", rd)
            rd = rd.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                            f'relationships/image" Target="../media/{nombre_media}"/></Relationships>')
            nid = max(ids) + 1 + i
            xd = xd.replace("</xdr:wsDr>", (
                f'<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>{c_ini}</xdr:col><xdr:colOff>0</xdr:colOff>'
                f'<xdr:row>{f_desde}</xdr:row><xdr:rowOff>{off}</xdr:rowOff></xdr:from><xdr:to><xdr:col>{c_fin}</xdr:col>'
                f'<xdr:colOff>0</xdr:colOff><xdr:row>{f_hasta}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr>'
                f'<xdr:cNvPr id="{nid}" name="Firma SYNCORE {i + 1}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>'
                f'</xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                f'r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/>'
                f'<a:ext cx="{ancho_caja}" cy="{alto_caja}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
                f'<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>'))
        L.escribir(dib, xd)
        L.escribir(rels_dib, rd)
        ct = L.leer("[Content_Types].xml")
        if 'Extension="png"' not in ct:
            L.escribir("[Content_Types].xml", ct.replace("<Default ", '<Default Extension="png" ContentType="image/png"/><Default ', 1))
    
    
    def _firmas_en_validacion(L: X.Libro, pl: Plantilla, firmas: list[dict]):
        """Formato nuevo: nombre, cargo, fecha y firma en la tabla «VALIDACIÓN DEL INFORME» (Preparado / Revisado / Aprobado)."""
        xi = L.hoja_xml(HOJA_INFORME)
        cols, filas = pl.validacion["cols"], pl.validacion["filas"]
        v, slots, altas = {}, [], set()
        for i, f in enumerate(firmas):
            fila = filas.get(f.get("rol"))
            if not fila:
                continue
            v[f"{cols['nombre']}{fila}"] = f.get("nombre") or None
            if cols.get("cargo"):
                v[f"{cols['cargo']}{fila}"] = f.get("cargo") or None
            fe = f.get("fecha")
            if cols.get("fecha"):
                v[f"{cols['fecha']}{fila}"] = fe.strftime("%d-%m-%Y %H:%M") if hasattr(fe, "strftime") else (fe or None)
            if f.get("firma") and cols.get("firma"):
                c_ini = X.col_a_num(cols["firma"]) - 1
                fin = re.search(rf'<mergeCell ref="{cols["firma"]}{fila}:([A-Z]+){fila}"/>', xi)
                c_fin = X.col_a_num(fin.group(1)) if fin else c_ini + 4
                slots.append((i, f["firma"], c_ini, c_fin, fila - 1, 20000, fila, 40 * 12700 - 30000))
                altas.add(fila)
        xi = X.fijar_celdas(xi, v)
        for fila in altas:                                   # la fila de quien firma se agranda para que la firma se lea
            xi = _alto_fila(xi, fila, 40)
        L.guardar_hoja(HOJA_INFORME, xi)
        if slots:
            _poner_imagenes_firma(L, xi, slots)
    
    
    def _ocultar_y_filtrar(xml: str, filas_ocultar: set[int], fecha: date) -> str:
        def f(m):
            tag = m.group(0)
            n = int(re.search(r'\br="(\d+)"', tag).group(1))
            if n in filas_ocultar and 'hidden="1"' not in tag:
                tag = tag.replace("<row ", '<row hidden="1" ', 1)
            return tag
        xml = re.sub(r"<row\b[^>]*>", f, xml)
        xml = re.sub(r"(<dateGroupItem\b[^>]*\byear=\")\d+(\"[^>]*\bmonth=\")\d+(\"[^>]*\bday=\")\d+",
                     rf"\g<1>{fecha.year}\g<2>{fecha.month}\g<3>{fecha.day}", xml)
        ult = X.ultima_fila(xml)
        xml = re.sub(r'(<autoFilter ref="[A-Z]+\d+:[A-Z]+)\d+(")', rf"\g<1>{ult}\g<2>", xml, count=1)
        return X.actualizar_dimension(xml)
    
    
    def _agregar_personal(L: X.Libro, pl: Plantilla, cons: Consolidado, fecha: date):
        xp = L.hoja_xml(HOJA_PERSONAL)
        n = X.ultima_fila(xp)
        serial = X.fecha_a_serial(fecha)
        mes = X.fecha_a_serial(fecha.replace(day=1))
        base = next(iter(pl.filas_persona_xml.values()))
        nuevas = []
        for p in cons.personas:
            n += 1
            fx = pl.filas_persona_xml.get(p.fila, base)
            fx = X.quitar_formulas_relativas(X.renumerar_fila(fx, n)).replace(' hidden="1"', "")
            cambios = {"B": mes, "C": serial, "AB": p.faena, "H": p.tipo_turno}
            if p.fila < 0:  # persona nueva: los datos conocidos (de la tarja, si los trae)
                vacias = {c: None for c in ("E", "F", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "X", "Y", "AC")}
                cambios.update(vacias)
                cambios.update({"D": p.rut, "G": "VIGENTE", "I": p.nombre, "J": p.cargo, "K": p.especialidad, "W": p.moimod,
                                "Z": p.estatus, "AA": p.acreditado})
                for col, val in (p.ficha or {}).items():
                    if col in ("H", "AB"):
                        continue
                    if isinstance(val, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", val):
                        val = X.fecha_a_serial(date.fromisoformat(val))
                    cambios[col] = val
            nuevas.append(X._reemplazar_en_fila(fx, n, cambios))
        xp = X.agregar_filas_al_final(xp, nuevas)
        xp = _ocultar_y_filtrar(xp, set(pl.filas_persona_xml), fecha)
        L.guardar_hoja(HOJA_PERSONAL, xp)
        cons._filas_personal_nuevas = {p.rut: fx for p, fx in zip(cons.personas, nuevas)}
        cons._rango_personal = (n - len(nuevas) + 1, n)
    
    
    def _agregar_udr(L: X.Libro, pl: Plantilla, ent: "Entradas") -> int:
        """Escribe el uso de recurso del día en la hoja BD UDR, a continuación del último registro. Devuelve la primera fila."""
        hoja = "BD UDR"
        xu = L.hoja_xml(hoja)
        ultima, estilos = 2, {}
        for n, a, fx, _, _ in X.filas(xu):
            cs = X.celdas(fx)
            c_ = cs.get("C")
            if c_ and n > 2 and L.valor_celda(c_[0], c_[1]) not in (None, ""):
                ultima = n
                estilos = {col: at.get("s") for col, (at, _, _) in cs.items()}
        primera = ultima + 1
        valores = {}
        cols = pl.udr_cols or {"fecha": "B", "ndc": "C", "descripcion": "D", "respaldo": "E", "detalle": "F", "directos": "G",
                               "indirectos": "H", "especialidad": "I", "hh_directos": "K", "equipos": "L", "hm": "M"}
        for i, d in enumerate(ent.uso_recurso):
            n = primera + i
            dd = {**d, "fecha": ent.fecha, "detalle": d.get("detalle") or d.get("descripcion")}
            fila = {col: dd.get(clave) for clave, col in cols.items() if clave in dd}
            for col, val in fila.items():
                valores[f"{col}{n}"] = val if val not in ("",) else None
        if valores:
            xu = X.fijar_celdas(xu, valores)
            # estilo de la última fila con datos para las celdas nuevas que no tenían formato
            for i in range(len(ent.uso_recurso)):
                n = primera + i
                for col, s_ in estilos.items():
                    if s_:
                        xu = re.sub(rf'<c r="{col}{n}"(?![^>]*\bs=)', f'<c r="{col}{n}" s="{s_}"', xu, count=1)
            L.guardar_hoja(hoja, X.actualizar_dimension(xu))
        return primera
    
    
    def _agregar_maquinaria(L: X.Libro, pl: Plantilla, ent: "Entradas"):
        """Copia la BD Maquinaria del día anterior al nuevo día (con los cambios de estado de la pantalla)."""
        hoja = "BD Maquinaria"
        x = L.hoja_xml(hoja)
        n = X.ultima_fila(x)
        serial = X.fecha_a_serial(ent.fecha)
        nuevas = []
        for d in maquinaria_del_dia(pl, ent.maquinaria):
            n += 1
            fx = d["_xml"]
            fila = re.sub(r'<row r="\d+"', f'<row r="{n}"', fx, count=1)
            fila = re.sub(r'(<c r="[A-Z]+)\d+"', rf'\g<1>{n}"', fila)
            fila = re.sub(rf'(<c r="B{n}"[^>]*>)<v>[^<]*</v>', rf"\g<1><v>{serial}</v>", fila, count=1)
            cambios = dict(ent.maquinaria.get(_clave_patente(d["patente"]), {}))
            orig = next((o for o in pl.maquinaria if _clave_patente(o["patente"]) == _clave_patente(d["patente"])), {})
            if d.get("disponible") != orig.get("disponible") or "disponible" in cambios:
                cambios["disponible"] = d["disponible"]           # mantención / panne: disponible SI
            for k, val in cambios.items():
                col = _MAQ_COLS.get(k)
                if not col:
                    continue
                m = re.search(rf'<c r="{col}{n}"([^>]*?)(?:/>|>.*?</c>)', fila, re.S)
                estilo = (re.search(r's="(\d+)"', m.group(1)) or [None, None])[1] if m else None
                nueva = X.xml_celda(f"{col}{n}", val or None, estilo)
                if m:
                    fila = fila.replace(m.group(0), nueva)
                else:                                          # celda nueva: en su lugar según la columna
                    pos = None
                    for mm in re.finditer(r'<c r="([A-Z]+)\d+"', fila):
                        if X.col_a_num(mm.group(1)) > X.col_a_num(col):
                            pos = mm.start()
                            break
                    fila = fila[:pos] + nueva + fila[pos:] if pos is not None else fila.replace("</row>", nueva + "</row>")
            nuevas.append(fila)
        x = X.agregar_filas_al_final(x, nuevas)
        x = re.sub(r'(<autoFilter ref="[A-Z]+\d+:[A-Z]+)\d+"', rf'\g<1>{n}"', x)
        L.guardar_hoja(hoja, x)
    
    
    def filas_master_actividades(pl: Plantilla, cons: Consolidado, ent: "Entradas") -> list[dict]:
        """Formato nuevo: una fila por persona directa en faena y actividad, con su Activity ID y criterio. Las horas anotadas en
        el reporte se reparten de forma proporcional en la jornada (HORAS_JORNADA), igual que la planilla oficial: así el total
        por actividad cuadra con las HH de la tarja (control T&M)."""
        canon = {norm(c): c for c in pl.criterios}
        def crit(c):
            return canon.get(norm(c), c) if c else CRITERIO_DEFECTO
        adicional = next((c for c in pl.criterios if norm(c).startswith("EXTRAORDINARIA")), "Extraordinaria / Adicional")
        sug_crit = {(a["archivo"], a["codigo"]): a.get("criterio_sug") for a in cons.actividades if a.get("criterio_sug")}
        filas = []
        for f in cons.filas_actividad:
            jornada = float(f.get("hh") or HORAS_JORNADA)
            det = [d for d in (f.get("detalle") or []) if d[2]]
            total = sum(d[2] for d in det)
            base = {"rut": f["rut"], "persona": f["persona"], "capataz": f.get("capataz") or f.get("supervisor") or ""}
            if not det or total <= 0:
                filas.append({**base, "id": "", "criterio": CRITERIO_DEFECTO, "hh": jornada,
                              "comentario": "Sin detalle de actividades en los reportes (jornada completa)"})
                continue
            partes = []
            for archivo, cod, h, desc in det:
                a = ent.asignacion.get(f"{archivo}|{cod}", {})
                c_ = a.get("criterio") or sug_crit.get((archivo, cod)) or (adicional if f.get("tipo_reporte") == "adicionales" else "")
                aid = a.get("id")
                aid = aid.strip() if isinstance(aid, str) else ""
                if not aid:                                    # convención del informe: agrupadores del P6 por tipo
                    cr_ = norm(c_ if isinstance(c_, str) else "")
                    aid = ("Uso de Recursos" if cod == "PRECOM" and "Uso de Recursos" in pl.p6 else
                           "Adicionales UDR" if cr_.startswith("EXTRAORDINARIA") and "Adicionales UDR" in pl.p6 else
                           "OTROS" if cr_.startswith("HABILITACION") and "OTROS" in pl.p6 else "")
                area = a.get("area") if isinstance(a.get("area"), str) else ""
                partes.append({**base, "id": aid, "criterio": crit(c_ if isinstance(c_, str) else ""),
                               "hh": round(jornada * h / total, 2),
                               "comentario": desc if aid or not area else f"{desc} [Área: {area}]"})
            partes[-1]["hh"] = round(partes[-1]["hh"] + jornada - sum(p_["hh"] for p_ in partes), 2)
            principal = max((p_ for p_ in partes if p_["id"] and not es_generica(p_["comentario"])), key=lambda p_: p_["hh"],
                            default=None)
            for p_ in partes:
                if principal and es_generica(p_["comentario"]) and (not p_["id"] or p_.get("id_auto")):
                    p_["id"] = principal["id"]
            filas += partes
        return filas
    
    
    _DATO_PERSONA = {"D": "rut", "I": "nombre", "J": "cargo", "K": "especialidad", "W": "moimod", "H": "tipo_turno", "F": "turno",
                     "AA": "acreditado", "Z": "estatus", "AB": "faena"}
    
    
    def _agregar_master(L: X.Libro, pl: Plantilla, cons: Consolidado, fecha: date, filas: list[dict]):
        xa = L.hoja_xml(HOJA_MASTER)
        n = X.ultima_fila(xa)
        serial_prev = X.fecha_a_serial(pl.fecha)
        patron = re.compile(r'<c r="B\d+"[^>]*><v>' + str(serial_prev) + r"</v>")
        ocultar = {num for num, a, fx, _, _ in X.filas(xa) if num > 13 and patron.search(fx)}
        est, modelo = pl.estilos_actividades, pl.master_fila_modelo
        ini, fin = getattr(cons, "_rango_personal", (4, 4))
        serial = X.fecha_a_serial(fecha)
    
        def formula(col, fila):
            f = pl.master_formulas.get(col)
            if not f:
                return None
            f = re.sub(rf"(?<![$\dA-Z])([A-Z]{{1,3}}){modelo}(?!\d)", rf"\g<1>{fila}", f)
            f = re.sub(r"'BD Personal'!\$([A-Z]+)\$\d+:\$([A-Z]+)\$\d+",
                       lambda m: f"'BD Personal'!${m.group(1)}${ini}:${m.group(2)}${fin}", f)
            return f
    
        def cache(f, fila_d):
            ret = re.findall(r"!\$?([A-Z]+)\$?\d*:", f)
            col = ret[-1] if ret else None
            if "'P6'" in f or "P6!" in f:
                return (pl.p6.get(fila_d["id"]) or {}).get(col, "") if fila_d["id"] else ""
            return getattr(fila_d["persona"], _DATO_PERSONA.get(col, ""), "") if col in _DATO_PERSONA else ""
    
        nuevas = []
        for d in filas:
            n += 1
            vals = {"B": serial, "C": d["id"] or None, "G": d["criterio"], "I": d["capataz"] or None, "J": d["rut"], "O": d["hh"],
                    "S": d["comentario"] or None}
            cel = []
            for k in range(2, 20):                          # B..S
                col = X.num_a_col(k)
                f = formula(col, n)
                if f:
                    cel.append(X.xml_celda(f"{col}{n}", cache(f, d), est.get(col), formula=f))
                else:
                    cel.append(X.xml_celda(f"{col}{n}", vals.get(col), est.get(col)))
            nuevas.append(f'<row r="{n}">{"".join(cel)}</row>')
        xa = X.agregar_filas_al_final(xa, nuevas)
        xa = _ocultar_y_filtrar(xa, ocultar, fecha)
        L.guardar_hoja(HOJA_MASTER, xa)
    
    
    def _agregar_actividades(L: X.Libro, pl: Plantilla, cons: Consolidado, fecha: date):
        xa = L.hoja_xml(HOJA_ACTIVIDADES)
        n = X.ultima_fila(xa)
        serial_prev = X.fecha_a_serial(pl.fecha)
        ocultar = set()
        patron = re.compile(r'<c r="B\d+"[^>]*><v>' + str(serial_prev) + r"</v>")
        for num, a, fx, _, _ in X.filas(xa):
            if num > 2 and patron.search(fx):
                ocultar.add(num)
        est = pl.estilos_actividades
        nuevas = []
        for f in cons.filas_actividad:
            n += 1
            origen = X.celdas(cons._filas_personal_nuevas[f["rut"]])
            cel = []
            for k in range(2, 28):                      # B..AA  <-  C..AB de BD Personal
                col_a, col_p = X.num_a_col(k), X.num_a_col(k + 1)
                s = est.get(col_a, "")
                sattr = f' s="{s}"' if s else ""
                if col_p in origen:
                    at, interior, _ = origen[col_p]
                    t = f' t="{at["t"]}"' if at.get("t") else ""
                    interior = re.sub(r"<f\b[^>]*/>|<f\b[^>]*>.*?</f>", "", interior or "", flags=re.S)
                    cel.append(f'<c r="{col_a}{n}"{sattr}{t}>{interior}</c>' if interior else f'<c r="{col_a}{n}"{sattr}/>')
                else:
                    cel.append(f'<c r="{col_a}{n}"{sattr}/>')
            for clave, col in COL_A_EXTRA.items():
                cel.append(X.xml_celda(f"{col}{n}", f.get(clave) or None, est.get(col)))
            for col in ("AH", "AI"):
                cel.append(X.xml_celda(f"{col}{n}", None, est.get(col)))
            nuevas.append(f'<row r="{n}" ht="15" customHeight="1">{"".join(cel)}</row>')
        xa = X.agregar_filas_al_final(xa, nuevas)
        xa = _ocultar_y_filtrar(xa, ocultar, fecha)
        L.guardar_hoja(HOJA_ACTIVIDADES, xa)
    
    
    def _reemplazar_fotos(L: X.Libro, pl: Plantilla, fotos: list[tuple[bytes | None, str]]):
        from PIL import Image, ImageOps
    
        hechos = False
        for i, f in enumerate(pl.fotos):
            if i >= len(fotos) or not fotos[i][0] or not f.get("media"):
                continue
            try:
                original = Image.open(io.BytesIO(L.leer_bytes(f["media"])))
                aspecto, ancho_obj = original.width / original.height, min(1400, max(800, original.width))
            except Exception:
                aspecto, ancho_obj = f.get("aspecto", 1.33), 1200
            img = ImageOps.exif_transpose(Image.open(io.BytesIO(fotos[i][0]))).convert("RGB")
            # Recorte central al formato del recuadro de la planilla
            if img.width / img.height > aspecto:
                nuevo_ancho = int(img.height * aspecto)
                x0 = (img.width - nuevo_ancho) // 2
                img = img.crop((x0, 0, x0 + nuevo_ancho, img.height))
            else:
                nuevo_alto = int(img.width / aspecto)
                y0 = (img.height - nuevo_alto) // 2
                img = img.crop((0, y0, img.width, y0 + nuevo_alto))
            img = img.resize((ancho_obj, int(ancho_obj / aspecto)))
            buf = io.BytesIO()
            img.save(buf, "JPEG", quality=85, optimize=True)
            destino = re.sub(r"\.\w+$", "", f["media"]) + f"_foto{i + 1}.jpeg"
            L.escribir(destino, buf.getvalue())
            rels_path = f["dibujo"].replace("drawings/", "drawings/_rels/") + ".rels"
            rels = L.leer(rels_path)
            rels = re.sub(r'(<Relationship\b[^>]*Id="' + re.escape(f["rel"]) + r'"[^>]*Target=")[^"]+(")',
                          rf"\g<1>../media/{destino.rsplit('/', 1)[-1]}\g<2>", rels)
            rels = re.sub(r'(<Relationship\b[^>]*Target=")[^"]+("[^>]*Id="' + re.escape(f["rel"]) + '")',
                          rf"\g<1>../media/{destino.rsplit('/', 1)[-1]}\g<2>", rels)
            L.escribir(rels_path, rels)
            hechos = True
        if hechos:
            ct = L.leer("[Content_Types].xml")
            if 'Extension="jpeg"' not in ct:
                ct = ct.replace("<Default ", '<Default Extension="jpeg" ContentType="image/jpeg"/><Default ', 1)
                L.escribir("[Content_Types].xml", ct)
    
    
    def _imagen_blanca() -> bytes:
        from PIL import Image
        buf = io.BytesIO()
        Image.new("RGB", (400, 300), (255, 255, 255)).save(buf, "JPEG", quality=70)
        return buf.getvalue()
    
    
    def paginas_informe(datos: bytes) -> int:
        L = X.Libro(datos)
        xml = L.hoja_xml(HOJA_INFORME)
        m = re.search(r'<rowBreaks\b[^>]*count="(\d+)"', xml)
        return int(m.group(1)) + 1 if m else 1
    
    
    _FUNC_NUEVAS = re.compile(r"_xlfn\.(?:XLOOKUP|XMATCH|UNIQUE|_xlws\.FILTER|FILTER|SORT|LET|LAMBDA)", re.I)
    
    
    def _sin_funciones_nuevas(datos: bytes) -> bytes:
        """LibreOffice (versiones de Debian / Windows antiguas) no conoce BUSCARX, UNICOS, FILTRAR…: muestra #¿NOMBRE?.
        Para el PDF, esas fórmulas de la hoja del informe se reemplazan por su resultado ya calculado por la app."""
        try:
            L = X.Libro(datos)
            xi = L.hoja_xml(HOJA_INFORME)
        except Exception:
            return datos
        if not _FUNC_NUEVAS.search(xi):
            return datos
    
        def f(m):
            cel = m.group(0)
            fm = re.search(r"<f\b[^>]*(?:/>|>(.*?)</f>)", cel, re.S)
            if not fm or not _FUNC_NUEVAS.search(fm.group(0)):
                return cel
            cel = cel.replace(fm.group(0), "")
            if re.search(r'\st="str"', cel):                  # texto calculado: queda como texto
                vm = re.search(r"<v>(.*?)</v>", cel, re.S)
                texto = vm.group(1) if vm else ""
                cel = re.sub(r'\st="str"', ' t="inlineStr"', cel, count=1)
                cel = re.sub(r"<v>.*?</v>", "", cel, flags=re.S)
                cel = cel.replace("</c>", f'<is><t xml:space="preserve">{texto}</t></is></c>') if texto else re.sub(r' t="inlineStr"', "", cel)
            elif re.search(r'\st="e"', cel):                   # error en caché: celda vacía
                cel = re.sub(r'\st="e"', "", cel)
                cel = re.sub(r"<v>.*?</v>", "", cel, flags=re.S)
            return cel
        xi = X._RE_CELL.sub(f, xi)
        L.guardar_hoja(HOJA_INFORME, xi)
        return L.bytes(recalcular=True)
    
    
    def _solo_hoja_informe(datos: bytes) -> bytes:
        """Copia para el PDF: las demás hojas quedan ocultas (LibreOffice no las imprime), así se exporta el informe completo
        sin depender de contar sus páginas."""
        L = X.Libro(datos)
        wb = L.leer("xl/workbook.xml")
        def f(m):
            tag = m.group(0)
            if f'name="{HOJA_INFORME}"' in tag:
                return re.sub(r'\sstate="[^"]*"', "", tag)
            tag = re.sub(r'\sstate="[^"]*"', "", tag)
            return tag.replace("<sheet ", '<sheet state="hidden" ', 1)
        wb = re.sub(r"<sheet\b[^>]*/>", f, wb)
        idx = [m.group(0) for m in re.finditer(r"<sheet\b[^>]*/>", wb)]
        pos = next((i for i, t in enumerate(idx) if f'name="{HOJA_INFORME}"' in t), 0)
        wb = re.sub(r'activeTab="\d+"', f'activeTab="{pos}"', wb)
        wb = re.sub(r'firstSheet="\d+"', f'firstSheet="{pos}"', wb)
        L.escribir("xl/workbook.xml", wb)
        return L.bytes(recalcular=True)
    
    
    def exportar_pdf(datos: bytes, timeout: int = 300) -> bytes | None:
        """PDF del informe (solo la hoja del informe) con LibreOffice, si está instalado. Recalcula las fórmulas."""
        import os
        import shutil
        import subprocess
        import tempfile
    
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        for ruta in (r"C:\Program Files\LibreOffice\program\soffice.exe", r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
            if not soffice and os.path.exists(ruta):
                soffice = ruta
        if not soffice:
            return None
        datos = _solo_hoja_informe(_sin_funciones_nuevas(datos))
        with tempfile.TemporaryDirectory() as tmp:
            perfil = os.path.join(tmp, "perfil", "user")
            os.makedirs(perfil)
            with open(os.path.join(perfil, "registrymodifications.xcu"), "w", encoding="utf-8") as f:
                f.write('<?xml version="1.0" encoding="UTF-8"?><oor:items xmlns:oor="http://openoffice.org/2001/registry" '
                        'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                        '<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse">'
                        '<value>0</value></prop></item></oor:items>')
            entrada = os.path.join(tmp, "informe.xlsx")
            with open(entrada, "wb") as f:
                f.write(datos)
            perfil_url = "file:///" + os.path.join(tmp, "perfil").replace("\\", "/").lstrip("/")
            try:
                subprocess.run([soffice, f"-env:UserInstallation={perfil_url}", "--headless", "--convert-to",
                                "pdf:calc_pdf_Export", "--outdir", tmp, entrada],
                               check=True, timeout=timeout, capture_output=True)
            except (subprocess.SubprocessError, OSError):
                return None
            salida = os.path.join(tmp, "informe.pdf")
            return open(salida, "rb").read() if os.path.exists(salida) else None
    
    
    # ---------------------------------------------------------------- resultados en caché
    
    _REF = re.compile(r"(?<![A-Z!'])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])")
    
    
    def _desplazar(formula: str, dfila: int, dcol: int) -> str:
        """Desplaza referencias relativas (fórmulas compartidas de Excel)."""
        def f(m):
            c_abs, col, f_abs, fila = m.groups()
            col_n = X.col_a_num(col) + (0 if c_abs else dcol)
            fila_n = int(fila) + (0 if f_abs else dfila)
            return f"{c_abs}{X.num_a_col(col_n)}{f_abs}{fila_n}"
        return _REF.sub(f, formula)
    
    
    def _cachear_resultados(xml: str, L: X.Libro, cons: Consolidado, fijos: dict | None = None) -> str:
        """Escribe en cada fórmula de la hoja del informe su resultado, para que el archivo muestre los números
        correctos aunque se abra sin recalcular (vista protegida de Excel, visores, previsualización)."""
        valores, formulas, compartidas, cache_previo = {}, {}, {}, {}
        for n, a, fx, _, _ in X.filas(xml):
            for col, (at, inn, _) in X.celdas(fx).items():
                ref = f"{col}{n}"
                m = re.search(r"<f\b([^>]*)(?:/>|>(.*?)</f>)", inn or "", re.S)
                if m:
                    previo = L.valor_celda(at, inn)
                    if isinstance(previo, (int, float)) and not isinstance(previo, bool):
                        cache_previo[ref] = previo
                    fa = X._attrs(m.group(1))
                    texto = X._desescapar(m.group(2) or "")
                    if fa.get("t") == "shared":
                        if texto:
                            compartidas[fa.get("si")] = (ref, texto)
                            formulas[ref] = texto
                        else:
                            formulas[ref] = ("__shared__", fa.get("si"))
                    else:
                        formulas[ref] = texto
                else:
                    valores[ref] = L.valor_celda(at, inn)
        for ref, f in list(formulas.items()):
            if isinstance(f, tuple):
                maestro = compartidas.get(f[1])
                if not maestro:
                    del formulas[ref]
                    continue
                mc, mf = X.separar_ref(maestro[0])
                c_, f_ = X.separar_ref(ref)
                formulas[ref] = _desplazar(maestro[1], f_ - mf, X.col_a_num(c_) - X.col_a_num(mc))
    
        por_cargo = cons.totales.get("por_cargo", {})
        resultado: dict[str, float] = dict(fijos or {})
        en_curso: set[str] = set()
    
        def valor(ref: str):
            ref = ref.replace("$", "")
            if ref in resultado:
                return resultado[ref]
            if ref in formulas:
                if ref in en_curso:
                    raise ValueError("circular")
                en_curso.add(ref)
                try:
                    r = evaluar(ref, formulas[ref])
                finally:
                    en_curso.discard(ref)
                resultado[ref] = r
                return r
            v = valores.get(ref)
            return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0
    
        def rango(a: str, b: str):
            ca, fa_ = X.separar_ref(a.replace("$", ""))
            cb, fb = X.separar_ref(b.replace("$", ""))
            return [valor(f"{X.num_a_col(c)}{f}") for f in range(fa_, fb + 1) for c in range(X.col_a_num(ca), X.col_a_num(cb) + 1)]
    
        def evaluar(ref: str, f: str):
            if "COUNTIFS" in f.upper():
                col, fila = X.separar_ref(ref)
                if fila in por_cargo and col in por_cargo[fila]:
                    return por_cargo[fila][col]
                if ref in cache_previo:      # COUNTIFS sobre otra base (p. ej. BD Maquinaria): se mantiene su valor
                    return cache_previo[ref]
                raise ValueError("countifs sin valor")
            if "!" in f:
                if ref in cache_previo:
                    return cache_previo[ref]
                raise ValueError("referencia a otra hoja")
            expr = f.lstrip("=")
            def suma(m):
                partes = []
                for arg in _args(m.group(1)):
                    r_ = re.fullmatch(r"\$?([A-Z]+\$?\d+):\$?([A-Z]+\$?\d+)", arg.strip())
                    partes.append(f"__sum(__rango('{r_.group(1)}','{r_.group(2)}'))" if r_ else f"({arg})")
                return "(" + "+".join(partes or ["0"]) + ")"
            anterior = None
            while anterior != expr:
                anterior = expr
                expr = re.sub(r"SUM\(([^()]*)\)", suma, expr, flags=re.I)
            if re.search(r"[A-Z]{2,}\(", expr.replace("__sum(", "").replace("__rango(", "")):
                raise ValueError("función no soportada")
            expr = _REF.sub(lambda m: f"__v('{m.group(2)}{m.group(4)}')", expr)
            return eval(expr, {"__builtins__": {}}, {"__v": valor, "__rango": rango, "__sum": sum})
    
        for ref in formulas:
            try:
                valor(ref)
            except Exception:
                resultado.pop(ref, None)
    
        def escribir(m):
            cel = m.group(0)
            a = X._attrs(m.group(1))
            ref = a.get("r")
            if ref not in resultado or "<f" not in cel:
                return cel
            v = resultado[ref]
            v = int(v) if isinstance(v, float) and v.is_integer() else v
            cel = re.sub(r"<v>.*?</v>", "", cel, flags=re.S)
            cel = re.sub(r'\st="(?:str|s|e|b|n)"', "", cel, count=1)
            return cel.replace("</c>", f"<v>{v}</v></c>")
        return X._RE_CELL.sub(escribir, xml)

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


diario = _modulo_diario()

# ======================================================================
# MÓDULO: tm
# ======================================================================
def _modulo_tm():
    import io
    import re
    import unicodedata
    from dataclasses import dataclass, field
    from datetime import date, datetime
    
    import pandas as pd
    
    CAL = calendario
    
    HOJAS_MO = {"DB MOD-MOI": "Construcción", "DB MOD-MOI PRECOM": "Precom"}
    HOJA_HM = "DB HORAS MAQUINAS"
    GRUPOS_HM = ["Directos OOCC", "Directos Montaje", "Indirectos"]
    
    
    def norm(s) -> str:
        s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().upper()
        return re.sub(r"\s+", " ", s).strip()
    
    
    @dataclass
    class ConsolidadoTM:
        archivo: str
        fecha_corte: date | None
        fecha_inicio: date | None
        contrato: str
        nombre: str
        mo: pd.DataFrame                     # fecha, ambito, tipo, cargo, hh
        dotacion: pd.DataFrame               # fecha, ambito, tipo, dotacion
        acum_inicial: dict[str, float]       # "ambito|tipo" -> HH acumuladas antes del primer día de la matriz
        eq: pd.DataFrame                     # fecha, grupo, equipo, hm, cantidad
        eq_inicial: dict[str, float]
        resumen: dict = field(default_factory=dict)
        avisos: list[str] = field(default_factory=list)
    
        # ---------------------------------------------------------- serialización
        def a_dict(self) -> dict:
            def df(d: pd.DataFrame):
                d = d.copy()
                if "fecha" in d:
                    d["fecha"] = d["fecha"].astype(str)
                return d.to_dict("split")
            res = {}
            for k, v in self.resumen.items():
                res[k] = {"__df__": df(v)} if isinstance(v, pd.DataFrame) else v
            return {"archivo": self.archivo, "fecha_corte": str(self.fecha_corte) if self.fecha_corte else None,
                    "fecha_inicio": str(self.fecha_inicio) if self.fecha_inicio else None, "contrato": self.contrato,
                    "nombre": self.nombre, "mo": df(self.mo), "dotacion": df(self.dotacion), "acum_inicial": self.acum_inicial,
                    "eq": df(self.eq), "eq_inicial": self.eq_inicial, "resumen": res, "avisos": self.avisos}
    
        @staticmethod
        def de_dict(d: dict) -> "ConsolidadoTM":
            def df(x):
                t = pd.DataFrame(x["data"], columns=x["columns"])
                if "fecha" in t:
                    t["fecha"] = pd.to_datetime(t["fecha"]).dt.date
                return t
            res = {k: (df(v["__df__"]) if isinstance(v, dict) and "__df__" in v else v) for k, v in d.get("resumen", {}).items()}
            ter = res.get("terminos")
            if isinstance(ter, pd.DataFrame) and "Término proyectado" in ter:
                ter["Término proyectado"] = pd.to_datetime(ter["Término proyectado"], errors="coerce").dt.date
                ter["Término proyectado"] = ter["Término proyectado"].where(ter["Término proyectado"].notna(), None)
            fc = date.fromisoformat(d["fecha_corte"]) if d.get("fecha_corte") else None
            fi = date.fromisoformat(d["fecha_inicio"]) if d.get("fecha_inicio") else None
            return ConsolidadoTM(d["archivo"], fc, fi, d.get("contrato", ""), d.get("nombre", ""), df(d["mo"]), df(d["dotacion"]),
                                 d.get("acum_inicial", {}), df(d["eq"]), d.get("eq_inicial", {}), res, d.get("avisos", []))
    
        # ---------------------------------------------------------- indicadores
        def hh_acumuladas(self, hasta: date | None = None, ambito: str | None = None, tipo: str | None = None) -> float:
            hasta = hasta or self.fecha_corte
            m = self.mo[self.mo["fecha"] <= hasta] if hasta else self.mo
            ini = 0.0
            for k, v in self.acum_inicial.items():
                a, t = k.split("|")
                if (ambito is None or a == ambito) and (tipo is None or t == tipo):
                    ini += v or 0
            if ambito:
                m = m[m["ambito"] == ambito]
            if tipo:
                m = m[m["tipo"] == tipo]
            return ini + float(m["hh"].sum())
    
    
    def _fecha(v) -> date | None:
        if isinstance(v, datetime):
            return v.date()
        if isinstance(v, date):
            return v
        return None
    
    
    def _matriz(ws_rows: list[tuple]) -> tuple[dict[int, date], list[dict]]:
        """Separa una hoja DB en secciones: {nombre, filas resumen, items}."""
        fila_fechas = next((r for r in ws_rows if sum(isinstance(v, (datetime, date)) for v in r) > 5), None)
        if fila_fechas is None:
            raise ValueError("no se encontró la fila de fechas")
        fechas = {j: _fecha(v) for j, v in enumerate(fila_fechas) if _fecha(v)}
        secciones, actual = [], None
        for r in ws_rows:
            lab = str(r[1]).strip() if len(r) > 1 and r[1] not in (None, "") else ""
            u = norm(lab)
            if not lab:
                continue
            if u.startswith("PERSONAL ") or u.startswith("EQUIPOS "):
                actual = {"nombre": lab, "acum": {}, "parc": {}, "dot": {}, "items": {}}
                secciones.append(actual)
                continue
            if actual is None or u.startswith("MATRIZ") or u.startswith("CONSOLIDADO") or u.startswith("TOTAL"):
                continue
            vals = {fechas[j]: float(r[j]) for j in fechas if j < len(r) and isinstance(r[j], (int, float)) and not isinstance(r[j], bool)}
            if "ACUMULAD" in u:
                # valor previo al primer día: columna inmediatamente anterior a la primera fecha con parcial
                actual["acum"] = vals
            elif "PARCIAL" in u:
                actual["parc"] = vals
            elif u.startswith("DOTACI"):
                actual["dot"] = vals
            else:
                base = lab
                k = 2
                while lab in actual["items"]:
                    lab = f"{base} ({k})"
                    k += 1
                actual["items"][lab] = vals
        return fechas, secciones
    
    
    def _ultimo_dia(secciones: list[dict]) -> date | None:
        dias = [d for s in secciones for d, v in s["parc"].items() if v]
        return max(dias) if dias else None
    
    
    def leer_tm(datos: bytes, archivo: str = "") -> ConsolidadoTM:
        import openpyxl
    
        wb = openpyxl.load_workbook(io.BytesIO(datos), read_only=True, data_only=True)
        hojas = {norm(n): n for n in wb.sheetnames}
        avisos: list[str] = []
        filas_mo, filas_dot, acum_ini = [], [], {}
        corte_dias = []
    
        for hoja, ambito in HOJAS_MO.items():
            nombre = hojas.get(norm(hoja))
            if not nombre:
                avisos.append(f"No está la hoja '{hoja}'.")
                continue
            fechas, secs = _matriz(list(wb[nombre].iter_rows(values_only=True)))
            ult = _ultimo_dia(secs)
            if ult:
                corte_dias.append(ult)
            for s in secs:
                tipo = "Indirecto" if "INDIRECT" in norm(s["nombre"]) else "Directo"
                dias = sorted(s["acum"])
                primero_parcial = min((d for d, v in s["parc"].items() if v), default=None)
                previo = [d for d in dias if primero_parcial and d < primero_parcial]
                acum_ini[f"{ambito}|{tipo}"] = s["acum"][previo[-1]] if previo else 0.0
                suma = {}
                for cargo, vals in s["items"].items():
                    for d, v in vals.items():
                        if ult and d > ult:
                            continue
                        if v:
                            filas_mo.append((d, ambito, tipo, cargo, v))
                            suma[d] = suma.get(d, 0) + v
                for d, v in s["dot"].items():
                    if ult and d > ult:
                        continue
                    if v:
                        filas_dot.append((d, ambito, tipo, v))
                difs = [d for d, v in s["parc"].items() if v and abs(v - suma.get(d, 0)) > 0.5]
                if difs:
                    avisos.append(f"{hoja} / {s['nombre']}: en {len(difs)} días la suma por cargo no coincide con HH PARCIALES "
                                  f"(ej. {difs[0]:%d-%m}).")
    
        filas_eq, eq_ini = [], {}
        nombre = hojas.get(norm(HOJA_HM))
        if nombre:
            fechas, secs = _matriz(list(wb[nombre].iter_rows(values_only=True)))
            ult = _ultimo_dia(secs)
            nombres_vistos = {}
            for i, s in enumerate(secs):
                u = norm(s["nombre"])
                # La planilla repite "EQUIPOS DIRECTOS OOCC" en la sección de montaje: se usa el orden.
                if "INDIRECT" in u:
                    grupo = "Indirectos"
                elif "MONTAJE" in u:
                    grupo = "Directos Montaje"
                elif u in nombres_vistos:
                    grupo = "Directos Montaje"
                else:
                    grupo = "Directos OOCC"
                nombres_vistos[u] = True
                primero = min((d for d, v in s["parc"].items() if v), default=None)
                previo = [d for d in sorted(s["acum"]) if primero and d < primero]
                eq_ini[grupo] = s["acum"][previo[-1]] if previo else 0.0
                # ¿Los ítems son HM o cantidades de equipos? (en indirectos la planilla multiplica la cantidad por 11 h)
                factor = 1.0
                muestras = [(s["parc"].get(d, 0), sum(v.get(d, 0) for v in s["items"].values())) for d in s["parc"] if s["parc"].get(d)]
                if muestras:
                    ratios = [p / q for p, q in muestras if q]
                    if ratios and all(abs(r - ratios[0]) < 0.01 for r in ratios) and abs(ratios[0] - 1) > 0.01:
                        factor = round(ratios[0], 4)
                for equipo, vals in s["items"].items():
                    for d, v in vals.items():
                        if ult and d > ult:
                            continue
                        if v:
                            filas_eq.append((d, grupo, equipo.strip(), v * factor, v if factor != 1 else None))
        else:
            avisos.append(f"No está la hoja '{HOJA_HM}'.")
    
        mo = pd.DataFrame(filas_mo, columns=["fecha", "ambito", "tipo", "cargo", "hh"])
        dot = pd.DataFrame(filas_dot, columns=["fecha", "ambito", "tipo", "dotacion"])
        eq = pd.DataFrame(filas_eq, columns=["fecha", "grupo", "equipo", "hm", "cantidad"])
        if mo.empty:
            raise ValueError("El archivo no tiene datos de mano de obra (hojas DB MOD-MOI).")
    
        contrato, nombre_c = "", ""
        for n in wb.sheetnames:
            if norm(n).startswith("DASHBOARD"):
                for r in wb[n].iter_rows(min_row=1, max_row=8, values_only=True):
                    if r and norm(r[0]).startswith("CONTRATO") and len(r) > 1 and r[1]:
                        contrato = str(r[1]).strip()
                    if r and norm(r[0]).startswith("NOMBRE") and len(r) > 1 and r[1]:
                        nombre_c = str(r[1]).strip()
                if contrato:
                    break
        if not contrato:
            m = re.search(r"[A-Z0-9]{3,5}-K-CC-\d{3,5}", archivo or "")
            contrato = m.group(0) if m else ""
    
        resumen = _leer_resumen(wb, hojas)
        corte = max(corte_dias) if corte_dias else None
        return ConsolidadoTM(archivo, corte, min(mo["fecha"]), contrato, nombre_c, mo, dot, acum_ini, eq, eq_ini, resumen, avisos)
    
    
    # ------------------------------------------------------------------ hoja Resumen
    
    def _leer_resumen(wb, hojas) -> dict:
        nombre = hojas.get("RESUMEN")
        if not nombre:
            return {}
        filas = [list(r) for r in wb[nombre].iter_rows(values_only=True)]
        res: dict = {}
    
        def buscar(txt: str, col_min: int = 0, col_max: int = 99):
            t = norm(txt)
            for i, f in enumerate(filas):
                for j, v in enumerate(f):
                    if col_min <= j <= col_max and isinstance(v, str) and norm(v).startswith(t):
                        return i, j
            return None
    
        def tabla(i_cab: int, j_ini: int, j_fin: int) -> pd.DataFrame:
            cab = [str(filas[i_cab][j]).strip() if j < len(filas[i_cab]) and filas[i_cab][j] is not None else f"col{j}"
                   for j in range(j_ini, j_fin + 1)]
            out = []
            for f in filas[i_cab + 1:]:
                vals = [f[j] if j < len(f) else None for j in range(j_ini, j_fin + 1)]
                if all(v in (None, "") for v in vals):
                    break
                out.append(vals)
            return pd.DataFrame(out, columns=cab)
    
        t = buscar("RESUMEN TIMES")
        if t:
            res["titulo"] = str(filas[t[0]][t[1]]).strip()
        for clave, rot, j_rango in (("mensual_mo", "TIMES & MATERIAL MOD + MOI", (0, 8)),
                                    ("mensual_eq", "TIMES & MATERIAL EQUIPOS", (0, 8))):
            p = buscar(rot, *j_rango)
            if p:
                i_cab = p[0] + 1
                fila = filas[i_cab]
                j_item = next((j for j, v in enumerate(fila) if norm(v) == "ITEM"), 1)
                j_tot = next((j for j, v in enumerate(fila) if norm(v).startswith("TOTAL") and j > j_item), j_item + 6)
                df = tabla(i_cab, j_item, j_tot)
                df = df.rename(columns={df.columns[0]: "Item", df.columns[1]: "Descripción"})
                df["Descripción"] = df["Descripción"].astype(str).str.replace(r"(?<=[a-z])(?=[A-Z])", " ", regex=True)
                res[clave] = df
        p = buscar("TIMES & MATERIAL MOD + MOI CONTRATADOS")
        if p:
            i_cab = p[0] + 1
            fila = filas[i_cab]
            j_ini = next((j for j, v in enumerate(fila) if j >= p[1] - 1 and norm(v) == "ITEM"), p[1] - 1)
            j_fin = max(j for j, v in enumerate(fila) if v not in (None, ""))
            df = tabla(i_cab, j_ini, j_fin)
            df = df.rename(columns={df.columns[0]: "Item", df.columns[1]: "Descripción"})
            df = df.rename(columns={c: "Diciembre" for c in df.columns if norm(c) == "DCIEMBRE"})
            df["Descripción"] = df["Descripción"].astype(str).str.replace(r"(?<=[a-z])(?=[A-Z])", " ", regex=True)
            res["contratados"] = df
        p = buscar("TERMINO PROYECTADO")
        if p:
            df = tabla(p[0], p[1] - 1, p[1])
            df.columns = ["Área", "Término proyectado"]
            df["Término proyectado"] = df["Término proyectado"].map(_fecha)
            res["terminos"] = df.dropna(subset=["Área"])
        p = buscar("AVANCE PROGRAMA")
        if p:
            i, j = p
            df = tabla(i, j, j + 4)
            df.columns = ["Curva", "HH acumulado", "% acumulado", "HH parcial", "% parcial"]
            res["avance"] = df
            arriba = filas[i - 1] if i > 0 else []
            res["avance_fecha"] = str(next((_fecha(v) for v in arriba if _fecha(v)), "") or "")
            res["avance_notas"] = [str(v).strip() for v in arriba if isinstance(v, str) and v.strip()]
        p = buscar("CIERRE SEMANAL")
        if p:
            i, j = p
            fila = filas[i]
            res["cierre_fecha"] = str(next((_fecha(v) for v in fila[j:] if _fecha(v)), "") or "")
            out = []
            for f in filas[i + 1:]:
                desc = f[j] if j < len(f) else None
                if desc in (None, ""):
                    if out:
                        break
                    continue
                out.append({"Ítem": f[j - 1] if j - 1 < len(f) else "", "Indicador": str(desc).strip(),
                            "Valor": f[j + 1] if j + 1 < len(f) else None})
            res["cierre"] = pd.DataFrame(out)
        return res
    
    
    # ------------------------------------------------------------------ cálculos para el tablero y el informe
    
    def filtrar(tm: ConsolidadoTM, desde: date, hasta: date, ambitos: list[str] | None = None):
        mo = tm.mo[(tm.mo["fecha"] >= desde) & (tm.mo["fecha"] <= hasta)]
        dot = tm.dotacion[(tm.dotacion["fecha"] >= desde) & (tm.dotacion["fecha"] <= hasta)]
        eq = tm.eq[(tm.eq["fecha"] >= desde) & (tm.eq["fecha"] <= hasta)]
        if ambitos:
            mo = mo[mo["ambito"].isin(ambitos)]
            dot = dot[dot["ambito"].isin(ambitos)]
        return mo, dot, eq
    
    
    def diario_mo(mo: pd.DataFrame) -> pd.DataFrame:
        if mo.empty:
            return pd.DataFrame(columns=["fecha", "Directo", "Indirecto", "Total"])
        d = mo.pivot_table(index="fecha", columns="tipo", values="hh", aggfunc="sum", fill_value=0).reset_index()
        for c in ("Directo", "Indirecto"):
            if c not in d:
                d[c] = 0.0
        d["Total"] = d["Directo"] + d["Indirecto"]
        return d.sort_values("fecha")
    
    
    def dotacion_diaria(dot: pd.DataFrame) -> pd.DataFrame:
        if dot.empty:
            return pd.DataFrame(columns=["fecha", "Directo", "Indirecto", "Total"])
        d = dot.pivot_table(index="fecha", columns="tipo", values="dotacion", aggfunc="sum", fill_value=0).reset_index()
        for c in ("Directo", "Indirecto"):
            if c not in d:
                d[c] = 0.0
        d["Total"] = d["Directo"] + d["Indirecto"]
        return d.sort_values("fecha")
    
    
    def por_cargo(mo: pd.DataFrame) -> pd.DataFrame:
        if mo.empty:
            return pd.DataFrame(columns=["cargo", "tipo", "ambito", "hh", "dias", "hh_dia", "participacion"])
        g = mo.groupby(["cargo", "tipo", "ambito"]).agg(hh=("hh", "sum"), dias=("fecha", "nunique")).reset_index()
        g["hh_dia"] = g["hh"] / g["dias"]
        g["participacion"] = g["hh"] / g["hh"].sum() * 100
        return g.sort_values("hh", ascending=False).reset_index(drop=True)
    
    
    def mensual(mo: pd.DataFrame, eq: pd.DataFrame) -> pd.DataFrame:
        def mes(s):
            return pd.to_datetime(s).dt.to_period("M").astype(str)
        a = mo.assign(mes=mes(mo["fecha"])).pivot_table(index="mes", columns="tipo", values="hh", aggfunc="sum", fill_value=0) \
            if not mo.empty else pd.DataFrame()
        b = eq.assign(mes=mes(eq["fecha"])).groupby("mes")["hm"].sum().rename("HM equipos") if not eq.empty \
            else pd.Series(dtype=float, name="HM equipos")
        if a.empty:
            t = b.rename_axis("mes").reset_index()
        else:
            t = a.join(b, how="outer").fillna(0).rename_axis("mes").reset_index()
        if "HM equipos" not in t:
            t["HM equipos"] = 0.0
        for c in ("Directo", "Indirecto"):
            if c not in t:
                t[c] = 0.0
        t["HH totales"] = t["Directo"] + t["Indirecto"]
        return t
    
    
    def semanal(mo: pd.DataFrame, eq: pd.DataFrame) -> pd.DataFrame:
        """HH directas, indirectas y HM por semana del proyecto (semana 1 = semana del inicio del proyecto)."""
        ini = CAL.configuracion()
        if mo.empty and eq.empty:
            return pd.DataFrame(columns=["semana", "desde", "hasta", "Directo", "Indirecto", "HH totales", "HM equipos"])
        a = mo.assign(n=CAL.columna_semana(mo["fecha"], ini)).pivot_table(index="n", columns="tipo", values="hh", aggfunc="sum",
                                                                           fill_value=0) if not mo.empty else pd.DataFrame()
        b = eq.assign(n=CAL.columna_semana(eq["fecha"], ini)).groupby("n")["hm"].sum().rename("HM equipos") if not eq.empty \
            else pd.Series(dtype=float, name="HM equipos")
        t = (b.to_frame() if a.empty else a.join(b, how="outer")).fillna(0)
        for c in ("Directo", "Indirecto", "HM equipos"):
            if c not in t:
                t[c] = 0.0
        t = t.rename_axis("n").reset_index().sort_values("n")
        t["HH totales"] = t["Directo"] + t["Indirecto"]
        t["desde"] = [CAL.inicio_semana(int(n), ini) for n in t["n"]]
        t["hasta"] = [CAL.fin_semana(int(n), ini) for n in t["n"]]
        t["semana"] = [f"S{int(n):02d}" for n in t["n"]]
        return t[["semana", "desde", "hasta", "Directo", "Indirecto", "HH totales", "HM equipos"]]
    
    
    def por_equipo(eq: pd.DataFrame) -> pd.DataFrame:
        if eq.empty:
            return pd.DataFrame(columns=["equipo", "grupo", "hm", "dias", "hm_dia"])
        g = eq.groupby(["equipo", "grupo"]).agg(hm=("hm", "sum"), dias=("fecha", "nunique")).reset_index()
        g["hm_dia"] = g["hm"] / g["dias"]
        return g.sort_values("hm", ascending=False).reset_index(drop=True)
    
    
    def indicadores(tm: ConsolidadoTM, desde: date, hasta: date, ambitos: list[str] | None = None) -> dict:
        mo, dot, eq = filtrar(tm, desde, hasta, ambitos)
        dm, dd = diario_mo(mo), dotacion_diaria(dot)
        n_dias = max(1, (hasta - desde).days + 1)
        # período anterior de igual duración (para comparar)
        from datetime import timedelta
        p_hasta = desde - timedelta(days=1)
        p_desde = p_hasta - timedelta(days=n_dias - 1)
        mo_p, dot_p, eq_p = filtrar(tm, p_desde, p_hasta, ambitos)
        dias_reg = dm["fecha"].nunique() if not dm.empty else 0
        dias_reg_p = mo_p["fecha"].nunique() if not mo_p.empty else 0
        hh = float(mo["hh"].sum())
        hh_p = float(mo_p["hh"].sum())
        pico = dm.loc[dm["Total"].idxmax()] if not dm.empty else None
        amb = ambitos or ["Construcción", "Precom"]
        return {
            "desde": desde, "hasta": hasta, "dias": dias_reg,
            "hh": hh, "hh_directo": float(mo.loc[mo["tipo"] == "Directo", "hh"].sum()),
            "hh_indirecto": float(mo.loc[mo["tipo"] == "Indirecto", "hh"].sum()),
            "hh_dia": hh / dias_reg if dias_reg else 0.0,
            "hh_dia_previo": hh_p / dias_reg_p if dias_reg_p else None, "previo_desde": p_desde, "previo_hasta": p_hasta,
            "dot_prom": float(dd["Total"].mean()) if not dd.empty else 0.0,
            "dot_prom_dir": float(dd["Directo"].mean()) if not dd.empty else 0.0,
            "dot_prom_ind": float(dd["Indirecto"].mean()) if not dd.empty else 0.0,
            "dot_max": float(dd["Total"].max()) if not dd.empty else 0.0,
            "pico_fecha": pico["fecha"] if pico is not None else None, "pico_hh": float(pico["Total"]) if pico is not None else 0.0,
            "hh_acum": sum(tm.hh_acumuladas(hasta, a) for a in amb),
            "hh_acum_constr": tm.hh_acumuladas(hasta, "Construcción"), "hh_acum_precom": tm.hh_acumuladas(hasta, "Precom"),
            "hh_acum_mod": sum(tm.hh_acumuladas(hasta, a, "Directo") for a in amb),
            "hm": float(eq["hm"].sum()), "hm_dir": float(eq.loc[eq["grupo"] != "Indirectos", "hm"].sum()),
            "hm_ind": float(eq.loc[eq["grupo"] == "Indirectos", "hm"].sum()),
            "hm_previo": float(eq_p["hm"].sum()),
            "hm_acum": sum(tm.eq_inicial.values()) + float(tm.eq.loc[tm.eq["fecha"] <= hasta, "hm"].sum()),
        }
    
    
    def _n(x, dec=0):
        s = f"{x:,.{dec}f}"
        return s.replace(",", "§").replace(".", ",").replace("§", ".")
    
    
    def _f(d) -> str:
        try:
            return pd.Timestamp(d).strftime("%d-%m-%Y")
        except (TypeError, ValueError):
            return str(d or "–")
    
    
    def valor_cierre(tm: ConsolidadoTM, texto: str):
        c = tm.resumen.get("cierre")
        if c is None or c.empty:
            return None
        fila = c[c["Indicador"].map(norm).str.contains(norm(texto), regex=False)]
        v = fila["Valor"].iloc[0] if not fila.empty else None
        return float(v) if isinstance(v, (int, float)) else None
    
    
    def resumen_narrativo(tm: ConsolidadoTM, ind: dict, cargos: pd.DataFrame, equipos: pd.DataFrame) -> list[str]:
        s = []
        s.append(f"Entre el {_f(ind['desde'])} y el {_f(ind['hasta'])} ({ind['dias']} días con registro) se informaron "
                 f"{_n(ind['hh'])} HH de mano de obra: {_n(ind['hh_directo'])} HH directas y {_n(ind['hh_indirecto'])} HH indirectas "
                 f"({_n(ind['hh_indirecto'] / ind['hh'] * 100 if ind['hh'] else 0, 1)}% indirectas).")
        s.append(f"La dotación promedio fue de {_n(ind['dot_prom'])} personas ({_n(ind['dot_prom_dir'])} directas y "
                 f"{_n(ind['dot_prom_ind'])} indirectas), con un máximo de {_n(ind['dot_max'])}.")
        if ind["pico_fecha"] is not None:
            s.append(f"El día de mayor actividad fue el {_f(ind['pico_fecha'])}, con {_n(ind['pico_hh'])} HH.")
        if ind["hh_dia_previo"]:
            var = (ind["hh_dia"] / ind["hh_dia_previo"] - 1) * 100
            s.append(f"El promedio diario ({_n(ind['hh_dia'])} HH/día) {'aumentó' if var > 0 else 'disminuyó'} "
                     f"{_n(abs(var), 1)}% respecto del período anterior de igual duración ({_n(ind['hh_dia_previo'])} HH/día).")
        s.append(f"Al {_f(ind['hasta'])} el contrato acumula {_n(ind['hh_acum_constr'])} HH de construcción y "
                 f"{_n(ind['hh_acum_precom'])} HH de precomisionamiento.")
        directos = cargos[cargos["tipo"] == "Directo"]
        if not directos.empty:
            top = directos.head(3)
            part = top["hh"].sum() / directos["hh"].sum() * 100
            s.append("Los cargos directos con más horas fueron " + ", ".join(f"{r.cargo.title()} ({_n(r.hh)} HH)" for r in top.itertuples())
                     + f", que concentran el {_n(part, 1)}% de las HH directas.")
        if ind["hm"]:
            s.append(f"Se registraron {_n(ind['hm'])} HM de equipos ({_n(ind['hm_dir'])} directas y {_n(ind['hm_ind'])} indirectas); "
                     f"el acumulado del contrato llega a {_n(ind['hm_acum'])} HM.")
            if not equipos.empty:
                e = equipos.iloc[0]
                s.append(f"El equipo con mayor uso fue {str(e.equipo).strip().title()} con {_n(e.hm)} HM ({e.grupo.lower()}).")
        av = tm.resumen.get("avance")
        if isinstance(av, pd.DataFrame) and not av.empty:
            def fila(t):
                f_ = av[av["Curva"].map(norm).str.startswith(norm(t))]
                return f_.iloc[0] if not f_.empty else None
            prog, real = fila("Avance Programado"), fila("Avance Real")
            if prog is not None and real is not None and isinstance(prog["% acumulado"], (int, float)) and isinstance(real["% acumulado"], (int, float)):
                dif = (real["% acumulado"] - prog["% acumulado"]) * 100
                s.append(f"Según el resumen al {_f(tm.resumen.get('avance_fecha'))}, el avance real acumulado es "
                         f"{_n(real['% acumulado'] * 100, 1)}% frente a {_n(prog['% acumulado'] * 100, 1)}% programado "
                         f"(diferencia de {_n(dif, 1)} pp).")
        pf = valor_cierre(tm, "PF PERIODO")
        if pf:
            s.append(f"El PF del período de construcción es {_n(pf, 2)}: "
                     + ("se gastaron más horas de las ganadas." if pf > 1.05 else "las horas gastadas están en línea con las ganadas."
                        if pf >= 0.95 else "se ganaron más horas de las gastadas."))
        ter = tm.resumen.get("terminos")
        if isinstance(ter, pd.DataFrame) and not ter.empty and ind["hasta"]:
            venc = ter[ter["Término proyectado"].notna() & (ter["Término proyectado"] < ind["hasta"])]
            prox = ter[ter["Término proyectado"].notna() & (ter["Término proyectado"] >= ind["hasta"])].sort_values("Término proyectado")
            if not venc.empty:
                s.append("Áreas con término proyectado ya cumplido a la fecha de corte: "
                         + ", ".join(f"{r['Área']} ({_f(r['Término proyectado'])})" for _, r in venc.iterrows()) + ".")
            if not prox.empty:
                r = prox.iloc[0]
                s.append(f"El próximo término proyectado es {r['Área']} el {_f(r['Término proyectado'])}.")
        return s

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


tm = _modulo_tm()

# ======================================================================
# MÓDULO: tm_pdf
# ======================================================================
def _modulo_tm_pdf():
    import io
    import os
    from datetime import date, datetime
    
    import pandas as pd
    from reportlab.graphics.charts.barcharts import VerticalBarChart
    from reportlab.graphics.charts.legends import Legend
    from reportlab.graphics.charts.linecharts import HorizontalLineChart
    from reportlab.graphics.shapes import Drawing, Rect, String
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.platypus import KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer
    from xml.sax.saxutils import escape
    
    P = informe_pdf
    CAL = calendario
    T = tm
    
    AZUL_MEDIO = colors.HexColor("#3B82F6")
    GRIS_BARRA = colors.HexColor("#94A3B8")
    SERIES = [P.NARANJO, P.AZUL, AZUL_MEDIO, GRIS_BARRA]
    
    
    def _etiquetas_fechas(fechas: list, max_etiquetas: int = 12) -> list[str]:
        paso = max(1, len(fechas) // max_etiquetas + (1 if len(fechas) % max_etiquetas else 0))
        return [pd.Timestamp(f).strftime("%d-%m") if i % paso == 0 else "" for i, f in enumerate(fechas)]
    
    
    def _leyenda(dr: Drawing, items: list[tuple], x: float, y: float):
        lg = Legend()
        lg.x, lg.y = max(40, P.UTIL - 78 * len(items)), y
        lg.alignment = "right"
        lg.columnMaximum = 1
        lg.deltax = 70
        lg.fontName, lg.fontSize = "Helvetica", 7
        lg.dx, lg.dy = 6, 6
        lg.colorNamePairs = items
        dr.add(lg)
    
    
    def grafico_barras_apiladas(fechas: list, series: list[tuple[str, list[float]]], titulo_y: str, alto: float = 62 * mm) -> Drawing:
        dr = Drawing(P.UTIL, alto)
        bc = VerticalBarChart()
        bc.x, bc.y, bc.width, bc.height = 34, 26, P.UTIL - 50, alto - 50
        bc.data = [s[1] for s in series] or [[0]]
        bc.categoryAxis.style = "stacked"
        bc.categoryAxis.categoryNames = _etiquetas_fechas(fechas)
        bc.categoryAxis.labels.fontName, bc.categoryAxis.labels.fontSize = "Helvetica", 6
        bc.categoryAxis.labels.angle, bc.categoryAxis.labels.dy, bc.categoryAxis.labels.dx = 45, -8, -6
        bc.categoryAxis.strokeColor = P.GRIS_CLARO
        bc.valueAxis.labels.fontName, bc.valueAxis.labels.fontSize = "Helvetica", 6.5
        bc.valueAxis.labelTextFormat = lambda v: P._n(v)
        bc.valueAxis.valueMin = 0
        bc.valueAxis.strokeColor = P.GRIS_CLARO
        bc.valueAxis.gridStrokeColor = P.GRIS_CLARO
        bc.valueAxis.visibleGrid = True
        bc.barSpacing, bc.groupSpacing = 0, 1.2
        for i in range(len(series)):
            bc.bars[i].fillColor = SERIES[i % len(SERIES)]
            bc.bars[i].strokeColor = None
        dr.add(bc)
        dr.add(String(0, alto - 10, titulo_y, fontName="Helvetica", fontSize=7, fillColor=P.GRIS))
        _leyenda(dr, [(SERIES[i % len(SERIES)], s[0]) for i, s in enumerate(series)], P.UTIL - 10, alto - 4)
        return dr
    
    
    def grafico_linea(fechas: list, series: list[tuple[str, list[float]]], titulo_y: str, alto: float = 50 * mm) -> Drawing:
        dr = Drawing(P.UTIL, alto)
        lc = HorizontalLineChart()
        lc.x, lc.y, lc.width, lc.height = 34, 26, P.UTIL - 50, alto - 50
        lc.data = [s[1] for s in series] or [[0]]
        lc.categoryAxis.categoryNames = _etiquetas_fechas(fechas)
        lc.categoryAxis.labels.fontName, lc.categoryAxis.labels.fontSize = "Helvetica", 6
        lc.categoryAxis.labels.angle, lc.categoryAxis.labels.dy, lc.categoryAxis.labels.dx = 45, -8, -6
        lc.categoryAxis.strokeColor = P.GRIS_CLARO
        lc.valueAxis.labels.fontName, lc.valueAxis.labels.fontSize = "Helvetica", 6.5
        lc.valueAxis.labelTextFormat = lambda v: P._n(v)
        lc.valueAxis.valueMin = 0
        lc.valueAxis.visibleGrid = True
        lc.valueAxis.gridStrokeColor = P.GRIS_CLARO
        lc.valueAxis.strokeColor = P.GRIS_CLARO
        for i in range(len(series)):
            lc.lines[i].strokeColor = SERIES[i % len(SERIES)]
            lc.lines[i].strokeWidth = 1.6
        dr.add(lc)
        dr.add(String(0, alto - 10, titulo_y, fontName="Helvetica", fontSize=7, fillColor=P.GRIS))
        _leyenda(dr, [(SERIES[i % len(SERIES)], s[0]) for i, s in enumerate(series)], P.UTIL - 10, alto - 4)
        return dr
    
    
    def grafico_ranking(etiquetas: list[str], valores: list[float], sufijo: str = "HH", colores: list | None = None) -> Drawing:
        alto_fila, x0 = 14, 165
        ancho = P.UTIL - x0 - 70
        maximo = max(valores) if valores else 1
        dr = Drawing(P.UTIL, alto_fila * len(valores) + 6)
        for i, (et, v) in enumerate(zip(etiquetas, valores)):
            y = dr.height - (i + 1) * alto_fila
            et = str(et).strip().title()
            et = et if len(et) <= 40 else et[:39] + "…"
            dr.add(String(0, y + 3.5, et, fontName="Helvetica", fontSize=7.4, fillColor=P.AZUL))
            w = ancho * v / maximo if maximo else 0
            dr.add(Rect(x0, y + 1, max(w, 0.5), 9, fillColor=(colores[i] if colores else P.NARANJO), strokeColor=None))
            dr.add(String(x0 + w + 5, y + 3, f"{P._n(v)} {sufijo}", fontName="Helvetica", fontSize=7, fillColor=P.AZUL))
        return dr
    
    
    def _num(v, dec=0):
        if isinstance(v, (int, float)) and not pd.isna(v):
            return P._n(v, dec)
        return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)
    
    
    def generar_informe_tecnico(tm: T.ConsolidadoTM, desde: date, hasta: date, ambitos: list[str] | None = None,
                                autor: str = "", observaciones: str = "", recomendaciones: str = "",
                                anexo_cargos: bool = True, logo: str | None = None, titulo: str = "") -> bytes:
        ambitos = ambitos or ["Construcción", "Precom"]
        mo, dot, eq = T.filtrar(tm, desde, hasta, ambitos)
        ind = T.indicadores(tm, desde, hasta, ambitos)
        dm, dd = T.diario_mo(mo), T.dotacion_diaria(dot)
        cargos, equipos = T.por_cargo(mo), T.por_equipo(eq)
        narrativa = T.resumen_narrativo(tm, ind, cargos, equipos)
        titulo = titulo or "INFORME TÉCNICO TIMES & MATERIAL"
        generado = datetime.now().strftime("%d-%m-%Y %H:%M")
    
        def marco(canvas, doc):
            A, H = A4
            canvas.saveState()
            canvas.setFillColor(P.AZUL)
            canvas.rect(0, H - 20 * mm, A, 20 * mm, stroke=0, fill=1)
            canvas.setFillColor(P.NARANJO)
            canvas.rect(0, H - 21.2 * mm, A, 1.2 * mm, stroke=0, fill=1)
            if logo and os.path.exists(logo):
                canvas.drawImage(P.logo_fondo_oscuro(logo), P.MARGEN, H - 17 * mm, height=14 * mm, width=40 * mm, preserveAspectRatio=True,
                                 anchor="w", mask="auto")
            else:
                canvas.setFillColor(P.NARANJO)
                canvas.setFont("Helvetica-Bold", 17)
                canvas.drawString(P.MARGEN, H - 12.5 * mm, "SYNCORE")
                canvas.setFillColor(colors.HexColor("#94A3B8"))
                canvas.setFont("Helvetica", 7)
                canvas.drawString(P.MARGEN, H - 16.5 * mm, "MONTAJES INDUSTRIALES")
            canvas.setFillColor(colors.white)
            canvas.setFont("Helvetica-Bold", 10.5)
            canvas.drawRightString(A - P.MARGEN, H - 10.5 * mm, titulo[:60])
            canvas.setFont("Helvetica", 8)
            canvas.setFillColor(colors.HexColor("#CBD5E1"))
            canvas.drawRightString(A - P.MARGEN, H - 15 * mm, f"{tm.contrato}  ·  Período {T._f(desde)} al {T._f(hasta)}")
            canvas.setStrokeColor(P.GRIS_CLARO)
            canvas.line(P.MARGEN, 13 * mm, A - P.MARGEN, 13 * mm)
            canvas.setFont("Helvetica", 7)
            canvas.setFillColor(P.GRIS)
            canvas.drawString(P.MARGEN, 9 * mm, f"SYNCORE · Fuente: {tm.archivo[:70]} · Generado el {generado}" + (f" por {autor}" if autor else ""))
            canvas.drawRightString(A - P.MARGEN, 9 * mm, f"Página {doc.page}")
            canvas.restoreState()
    
        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=P.MARGEN, rightMargin=P.MARGEN, topMargin=28 * mm,
                                bottomMargin=18 * mm, title=titulo, author=autor or "SYNCORE")
        E = P.E
        h = [Paragraph(escape(titulo.title() if titulo.isupper() else titulo), E["titulo"]),
             Paragraph(escape(f"Contrato {tm.contrato} — {tm.nombre}"), E["subtitulo"]),
             Paragraph(escape(f"Período informado: {T._f(desde)} al {T._f(hasta)} ({CAL.etiqueta(desde)}"
                              f"{'' if CAL.semana(desde) == CAL.semana(hasta) else ' a ' + CAL.etiqueta(hasta)}) · Ámbito: {', '.join(ambitos)} · "
                              f"Datos al {T._f(tm.fecha_corte)}"), E["subtitulo"]),
             Spacer(1, 6 * mm)]
    
        var = ""
        if ind["hh_dia_previo"]:
            v = (ind["hh_dia"] / ind["hh_dia_previo"] - 1) * 100
            var = f"{'+' if v >= 0 else ''}{P._n(v, 1)}%"
        h.append(P._fila_kpis([("HH del período", P._n(ind["hh"]), True), ("HH promedio diario", P._n(ind["hh_dia"]), False),
                               ("Dotación promedio", P._n(ind["dot_prom"]), False), ("Peak HH diario", P._n(ind["pico_hh"]), False)]))
        h.append(Spacer(1, 3 * mm))
        pf = T.valor_cierre(tm, "PF PERIODO")
        h.append(P._fila_kpis([("HH acumuladas contrato", P._n(ind["hh_acum"]), False),
                               ("HM equipos del período", P._n(ind["hm"]), False),
                               ("PF período construcción", P._n(pf, 2) if pf else "–", False),
                               ("HH/día vs período anterior", var or "–", False)]))
    
        h.append(Paragraph("1. Resumen ejecutivo", E["h2"]))
        for frase in narrativa:
            h.append(Paragraph(escape(frase), E["vineta"], bulletText="•"))
        if observaciones.strip():
            h.append(P._seccion("Observaciones", P._caja(observaciones)))
    
        # ---- Mano de obra
        h.append(PageBreak())
        h.append(Paragraph("2. Mano de obra", E["h2"]))
        if not dm.empty:
            h.append(grafico_barras_apiladas(list(dm["fecha"]), [("Directas", list(dm["Directo"])), ("Indirectas", list(dm["Indirecto"]))],
                                             "HH diarias"))
        if not dd.empty:
            h.append(Spacer(1, 2 * mm))
            h.append(grafico_linea(list(dd["fecha"]), [("Dotación total", list(dd["Total"])), ("Directa", list(dd["Directo"])),
                                                       ("Indirecta", list(dd["Indirecto"]))], "Dotación diaria (personas)"))
        mens = T.mensual(mo, eq)
        if not mens.empty:
            mens = mens.assign(mes=pd.to_datetime(mens["mes"]).dt.strftime("%m-%Y"))
            h.append(P._seccion("HH y HM por mes", P._tabla(mens, [("mes", "Mes", 1, "izq"), ("Directo", "HH directas", 1.2, "num"),
                                                                    ("Indirecto", "HH indirectas", 1.2, "num"),
                                                                    ("HH totales", "HH totales", 1.2, "num"),
                                                                    ("HM equipos", "HM equipos", 1.2, "num")], "")))
        sem = T.semanal(mo, eq)
        if len(sem) > 1:
            sem = sem.assign(periodo=[f"{a:%d-%m} al {b:%d-%m}" for a, b in zip(sem["desde"], sem["hasta"])])
            h.append(P._seccion("HH y HM por semana del proyecto", P._tabla(sem, [
                ("semana", "Semana", 0.7, "izq"), ("periodo", "Fechas", 1.2, "izq"), ("Directo", "HH directas", 1.1, "num"),
                ("Indirecto", "HH indirectas", 1.1, "num"), ("HH totales", "HH totales", 1.1, "num"),
                ("HM equipos", "HM equipos", 1.1, "num")], "")))
        amb = mo.groupby(["ambito", "tipo"])["hh"].sum().reset_index()
        if len(amb):
            amb["participacion"] = amb["hh"] / amb["hh"].sum() * 100
            h.append(P._seccion("Distribución por ámbito", P._tabla(amb, [("ambito", "Ámbito", 1.4, "izq"), ("tipo", "Tipo", 1, "izq"),
                                                                         ("hh", "HH del período", 1.2, "num"),
                                                                         ("participacion", "% del total", 1, "pct")], "")))
        if not cargos.empty:
            top = cargos[cargos["tipo"] == "Directo"].head(12)
            if not top.empty:
                h.append(_bloque("Cargos directos con más HH en el período", grafico_ranking(list(top["cargo"]), list(top["hh"]))))
            topi = cargos[cargos["tipo"] == "Indirecto"].head(10)
            if not topi.empty:
                h.append(_bloque("Cargos indirectos con más HH en el período",
                                 grafico_ranking(list(topi["cargo"]), list(topi["hh"]), colores=[P.AZUL] * len(topi))))
    
        # ---- Equipos
        if not eq.empty:
            h.append(PageBreak())
            h.append(Paragraph("3. Equipos y maquinaria", E["h2"]))
            de = eq.pivot_table(index="fecha", columns="grupo", values="hm", aggfunc="sum", fill_value=0).reset_index().sort_values("fecha")
            grupos = [g for g in T.GRUPOS_HM if g in de]
            h.append(grafico_barras_apiladas(list(de["fecha"]), [(g, list(de[g])) for g in grupos], "HM diarias"))
            pg = eq.groupby("grupo")["hm"].sum().reindex(grupos).reset_index()
            pg["participacion"] = pg["hm"] / pg["hm"].sum() * 100
            h.append(P._seccion("HM por grupo", P._tabla(pg, [("grupo", "Grupo", 2, "izq"), ("hm", "HM del período", 1.2, "num"),
                                                              ("participacion", "% del total", 1, "pct")], "")))
            tope = equipos.head(12)
            h.append(_bloque("Equipos con más horas máquina", grafico_ranking(list(tope["equipo"]), list(tope["hm"]), "HM",
                                                                               colores=[AZUL_MEDIO] * len(tope))))
    
        # ---- Avance, cierre y términos (hoja Resumen)
        av, ci, te = tm.resumen.get("avance"), tm.resumen.get("cierre"), tm.resumen.get("terminos")
        if any(isinstance(x, pd.DataFrame) and not x.empty for x in (av, ci, te)):
            h.append(PageBreak())
            h.append(Paragraph("4. Avance, productividad y términos proyectados", E["h2"]))
            if isinstance(av, pd.DataFrame) and not av.empty:
                t = av.copy()
                for c in ("% acumulado", "% parcial"):
                    t[c] = pd.to_numeric(t[c], errors="coerce") * 100
                notas = " · ".join(tm.resumen.get("avance_notas", [])[:2])
                h.append(P._seccion(f"Avance del programa al {T._f(tm.resumen.get('avance_fecha'))}" + (f" ({notas})" if notas else ""),
                                    P._tabla(t, [("Curva", "Curva", 3, "izq"), ("HH acumulado", "HH acumulado", 1.1, "num"),
                                                 ("% acumulado", "% acum.", 0.8, "pct"), ("HH parcial", "HH parcial", 1, "num"),
                                                 ("% parcial", "% parcial", 0.8, "pct")], "")))
            if isinstance(ci, pd.DataFrame) and not ci.empty:
                t = ci.copy()
                t["Valor"] = [(_num(v, 2) if "PF" in T.norm(i) else _num(v)) for v, i in zip(t["Valor"], t["Indicador"])]
                h.append(P._seccion(f"Cierre semanal al {T._f(tm.resumen.get('cierre_fecha'))}",
                                    P._tabla(t, [("Ítem", "Ítem", 0.5, "izq"), ("Indicador", "Indicador", 3, "izq"),
                                                 ("Valor", "Valor", 1, "der")], "")))
            if isinstance(te, pd.DataFrame) and not te.empty:
                t = te.copy()
                t["dias"] = [(pd.Timestamp(x) - pd.Timestamp(hasta)).days if x else None for x in t["Término proyectado"]]
                t["estado"] = ["Sin fecha" if d is None else "Fecha cumplida" if d < 0 else "Próximos 15 días" if d <= 15 else "En plazo"
                               for d in t["dias"]]
                h.append(P._seccion("Términos proyectados por área", P._tabla(
                    t, [("Área", "Área", 2.4, "izq"), ("Término proyectado", "Término proyectado", 1.2, "fecha"),
                        ("dias", f"Días al {T._f(hasta)}", 1, "num"), ("estado", "Estado", 1.2, "izq")], "")))
            mc = tm.resumen.get("contratados")
            if isinstance(mc, pd.DataFrame) and not mc.empty:
                cols = [c for c in mc.columns if c not in ("Item",)]
                h.append(P._seccion("Dotación contratada promedio por mes (hoja Resumen)", P._tabla(
                    mc, [(cols[0], "Descripción", 2.2, "izq")] + [(c, c, 1, "num1") for c in cols[1:]], "")))
    
        if recomendaciones.strip():
            h.append(P._seccion("5. Conclusiones y recomendaciones", P._caja(recomendaciones)))
    
        if anexo_cargos and not cargos.empty:
            h.append(PageBreak())
            h.append(Paragraph("Anexo A. HH por cargo en el período", E["h2"]))
            h.append(P._tabla(cargos, [("cargo", "Cargo", 3.2, "izq"), ("tipo", "Tipo", 0.9, "izq"), ("ambito", "Ámbito", 1.1, "izq"),
                                       ("hh", "HH", 0.9, "num"), ("dias", "Días", 0.6, "num"), ("hh_dia", "HH/día", 0.8, "num1"),
                                       ("participacion", "% total", 0.8, "pct")], ""))
            if not equipos.empty:
                h.append(Paragraph("Anexo B. HM por equipo en el período", E["h2"]))
                h.append(P._tabla(equipos, [("equipo", "Equipo", 3, "izq"), ("grupo", "Grupo", 1.4, "izq"), ("hm", "HM", 0.9, "num"),
                                            ("dias", "Días", 0.6, "num"), ("hm_dia", "HM/día", 0.8, "num1")], ""))
    
        h.append(Spacer(1, 6 * mm))
        h.append(Paragraph(escape("Fuente: consolidado Times & Material (hojas DB MOD-MOI, DB MOD-MOI PRECOM, DB HORAS MAQUINAS y "
                                  "Resumen). Las cifras de mano de obra y equipos se recalculan desde las bases diarias."), E["subtitulo"]))
        doc.build(h, onFirstPage=marco, onLaterPages=marco)
        return buf.getvalue()
    
    
    def _bloque(titulo: str, contenido) -> KeepTogether:
        return KeepTogether([Paragraph(escape(titulo), P.E["h2"]), contenido])

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


tm_pdf = _modulo_tm_pdf()

# ======================================================================
# MÓDULO: desviaciones
# ======================================================================
def _modulo_desviaciones():
    import io
    import re
    import unicodedata
    from dataclasses import dataclass, field
    from datetime import date, datetime, timedelta
    
    import pandas as pd
    
    CAL = calendario
    
    COLUMNAS = {  # campo interno -> encabezados aceptados (normalizados)
        "n": ["N", "NO", "N°", "NUMERO", "ITEM"],
        "fecha": ["FECHA"],
        "jornada": ["JORNADA", "TURNO"],
        "especialidad": ["ESPECIALIDAD", "DISCIPLINA"],
        "area": ["AREA", "SECTOR"],
        "actividad": ["ACTIVIDAD", "DESCRIPCION", "DESCRIPCION DE LA ACTIVIDAD"],
        "dotacion": ["DOTACION", "PERSONAS", "CANTIDAD DE PERSONAS"],
        "hh": ["HH DESVIADAS", "HH", "HORAS HOMBRE"],
        "hh_gestionadas": ["HH DESVIADAS GESTIONADAS", "HH GESTIONADAS"],
        "observacion": ["OBSERVACION", "OBSERVACIONES", "SOLICITANTE"],
        "referencia": ["REFERENCIA", "MOTIVO", "TIPO", "ORIGEN"],
    }
    NOMBRES = {"n": "N°", "fecha": "Fecha", "jornada": "Jornada", "especialidad": "Especialidad", "area": "Área",
               "actividad": "Actividad", "dotacion": "Dotación", "hh": "HH desviadas", "hh_gestionadas": "HH gestionadas",
               "observacion": "Observación", "referencia": "Motivo"}
    CAMPOS_TOTAL = {"DOTACION": "dotacion", "HH DESVIADAS": "hh", "HH DESVIADAS GESTIONADAS": "hh_gestionadas"}
    HH_MAX_PERSONA = 12.5     # una jornada no supera ~12 HH por persona
    
    
    def norm(s) -> str:
        s = unicodedata.normalize("NFKD", str(s if s is not None else "")).encode("ascii", "ignore").decode().upper()
        s = s.replace("°", "").replace("º", "")
        return re.sub(r"\s+", " ", re.sub(r"[^\w ]", " ", s)).strip()
    
    
    def _texto(v) -> str:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return ""
        return re.sub(r"\s+", " ", str(v).replace("\t", " ")).strip()
    
    
    def _num(v):
        if v is None or v == "":
            return None
        if isinstance(v, (int, float)):
            return None if pd.isna(v) else float(v)
        s = str(v).strip().replace(" ", "")
        if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", s):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None
    
    
    def _fecha(v):
        if isinstance(v, datetime):
            return v.date()
        if isinstance(v, date):
            return v
        if isinstance(v, (int, float)) and not pd.isna(v) and 20000 < v < 80000:      # número de serie de Excel
            return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v))).date()
        s = _texto(v)
        if not s:
            return None
        try:
            return pd.to_datetime(s, dayfirst=True).date()
        except (ValueError, TypeError):
            return None
    
    
    @dataclass
    class LogDesviaciones:
        archivo: str
        fecha_corte: date | None
        fecha_inicio: date | None
        revision: str
        hoja: str
        df: pd.DataFrame                                  # un registro por fila del LOG, ya limpio
        calidad: pd.DataFrame                             # n, fila_excel, problema, detalle
        unificaciones: list[str] = field(default_factory=list)
        totales: list[dict] = field(default_factory=list)  # {origen, campo, planilla, calculado}
        avisos: list[str] = field(default_factory=list)
    
        def a_dict(self) -> dict:
            d = self.df.copy()
            d["fecha"] = d["fecha"].astype(str)
            return {"archivo": self.archivo, "fecha_corte": str(self.fecha_corte) if self.fecha_corte else None,
                    "fecha_inicio": str(self.fecha_inicio) if self.fecha_inicio else None, "revision": self.revision,
                    "hoja": self.hoja, "df": d.to_dict("split"), "calidad": self.calidad.to_dict("split"),
                    "unificaciones": self.unificaciones, "totales": self.totales, "avisos": self.avisos}
    
        @staticmethod
        def de_dict(x: dict) -> "LogDesviaciones":
            d = pd.DataFrame(x["df"]["data"], columns=x["df"]["columns"])
            d["fecha"] = pd.to_datetime(d["fecha"], errors="coerce").dt.date
            d["fecha"] = d["fecha"].where(d["fecha"].notna(), None)
            cal = pd.DataFrame(x["calidad"]["data"], columns=x["calidad"]["columns"])
            f = lambda k: date.fromisoformat(x[k]) if x.get(k) else None
            return LogDesviaciones(x["archivo"], f("fecha_corte"), f("fecha_inicio"), x.get("revision", ""), x.get("hoja", ""),
                                   d, cal, x.get("unificaciones", []), x.get("totales", []), x.get("avisos", []))
    
    
    # ------------------------------------------------------------------ lectura
    
    def _filas(ws, max_vacias: int = 300):
        """Valores de la hoja; se detiene tras muchas filas vacías (estas planillas declaran rangos de un millón de filas)."""
        vacias = 0
        for i, fila in enumerate(ws.iter_rows(values_only=True), start=1):
            if any(v not in (None, "") for v in fila):
                vacias = 0
                yield i, fila
            else:
                vacias += 1
                if vacias > max_vacias:
                    return
    
    
    def _buscar_encabezado(filas) -> tuple[int, dict[str, int]] | None:
        for pos, (_, fila) in enumerate(filas[:40]):
            nombres = [norm(v) for v in fila]
            if not any(n in ("HH DESVIADAS",) or n.startswith("HH DESVIADA") for n in nombres):
                continue
            mapa: dict[str, int] = {}
            for campo, opciones in COLUMNAS.items():
                for j, n in enumerate(nombres):
                    if n and n in [norm(o) for o in opciones] and j not in mapa.values():
                        mapa[campo] = j
                        break
            if "hh" in mapa and "fecha" in mapa:
                return pos, mapa
        return None
    
    
    def _canonico(serie: pd.Series, campo: str, unificaciones: list[str], formato=None) -> pd.Series:
        """Une variantes que solo difieren en mayúsculas, tildes o espacios; conserva la forma más usada."""
        s = serie.map(_texto)
        claves = s.map(norm)
        elegido = {}
        for clave, grupo in s.groupby(claves):
            if not clave:
                continue
            conteo = grupo.value_counts()
            # Preferir la escritura con mayúsculas y minúsculas ("Caldera 4") a TODO MAYÚSCULAS o todo minúsculas,
            # salvo en códigos cortos (IIFF, HEX-005), donde manda la forma más usada.
            def puntaje(v):
                mixto = not (v.isupper() or v.islower()) and (len(v) > 6 or " " in v)
                return (mixto, conteo[v])
            preferido = max(conteo.index, key=puntaje)
            if not puntaje(preferido)[0] and " " in preferido and (preferido.isupper() or preferido.islower()):
                preferido = preferido[:1].upper() + preferido[1:].lower()
            nombre = formato(clave, preferido) if formato else preferido
            elegido[clave] = nombre
            variantes = sorted(set(grupo) - {nombre}) if len(conteo) > 1 or conteo.index[0] != nombre else []
            if variantes:
                unificaciones.append(f"{campo}: " + ", ".join(f"«{v}»" for v in variantes) + f" → «{nombre}»")
        return claves.map(lambda k: elegido.get(k, ""))
    
    
    def _jornada(clave: str, original: str) -> str:
        return {"DIA": "Día", "NOCHE": "Noche", "GENERAL": "General"}.get(clave, original[:1].upper() + original[1:].lower())
    
    
    def _totales_pivote(wb, hoja_detalle: str) -> list[dict]:
        """Totales 'Total general' de las tablas dinámicas (columnas 'Suma de ...') en las demás hojas."""
        res = []
        for ws in wb.worksheets:
            if ws.title == hoja_detalle:
                continue
            cab = None
            for _, fila in _filas(ws, 60):
                nombres = [norm(v) for v in fila]
                if any(n.startswith("SUMA DE ") for n in nombres):
                    cab = {j: CAMPOS_TOTAL.get(n[len("SUMA DE "):]) for j, n in enumerate(nombres) if n.startswith("SUMA DE ")}
                    continue
                if cab and nombres and "TOTAL GENERAL" in nombres:
                    for j, campo in cab.items():
                        if campo and j < len(fila) and _num(fila[j]) is not None:
                            res.append({"origen": f"Hoja {ws.title.strip()} (tabla dinámica)", "campo": campo, "planilla": _num(fila[j])})
                    cab = None
        return res
    
    
    def leer_log(datos: bytes, archivo: str = "") -> LogDesviaciones:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(datos), data_only=True, read_only=True)
        hoja, filas, enc = None, [], None
        for ws in wb.worksheets:
            fs = list(_filas(ws))
            e = _buscar_encabezado(fs)
            if e:
                hoja, filas, enc = ws.title, fs, e
                break
        if not enc:
            raise ValueError("No encontré una hoja con las columnas 'Fecha' y 'HH Desviadas'. ¿Es el LOG de desviación de recursos?")
        pos, mapa = enc
        fila_cab = filas[pos][0]
        totales = []
        # Totales de la planilla escritos sobre el encabezado (SUBTOTAL de la fila 2)
        for _, fila in filas[max(0, pos - 3):pos]:
            for campo in ("dotacion", "hh", "hh_gestionadas"):
                j = mapa.get(campo)
                if j is not None and j < len(fila) and _num(fila[j]) is not None and not isinstance(fila[j], str):
                    totales.append({"origen": f"Hoja {hoja.strip()} (total sobre el encabezado)", "campo": campo, "planilla": _num(fila[j])})
    
        registros = []
        for nfila, fila in filas[pos + 1:]:
            val = lambda campo: fila[mapa[campo]] if campo in mapa and mapa[campo] < len(fila) else None
            if all(_texto(val(c)) == "" for c in ("fecha", "actividad", "hh", "especialidad")):
                continue
            if norm(val("n")).startswith("TOTAL") or norm(val("especialidad")).startswith("TOTAL"):
                continue
            registros.append({
                "fila_excel": nfila, "n": _texto(val("n")).removesuffix(".0"), "fecha": _fecha(val("fecha")),
                "jornada": val("jornada"), "especialidad": val("especialidad"), "area": val("area"),
                "actividad": _texto(val("actividad")), "dotacion": _num(val("dotacion")), "hh": _num(val("hh")),
                "hh_gestionadas": _num(val("hh_gestionadas")), "observacion": _texto(val("observacion")),
                "referencia": val("referencia")})
        totales += _totales_pivote(wb, hoja)
        wb.close()
        if not registros:
            raise ValueError(f"La hoja «{hoja}» tiene el encabezado del LOG (fila {fila_cab}) pero no tiene registros.")
    
        df = pd.DataFrame(registros)
        unif: list[str] = []
        df["jornada"] = _canonico(df["jornada"], "Jornada", unif, _jornada)
        df["especialidad"] = _canonico(df["especialidad"], "Especialidad", unif)
        df["area"] = _canonico(df["area"], "Área", unif)
        df["referencia"] = _canonico(df["referencia"], "Motivo", unif)
        for c in ("dotacion", "hh", "hh_gestionadas"):
            df[c] = pd.to_numeric(df[c], errors="coerce")
    
        calidad = _control_calidad(df)
        unicos = {}
        for t in totales:
            unicos.setdefault((t["origen"], t["campo"], t["planilla"]), t)
        totales = list(unicos.values())
        for t in totales:
            t["calculado"] = float(df[t["campo"]].sum())
        fechas = [f for f in df["fecha"] if f]
        fecha_corte = max(fechas) if fechas else None
        rev = re.search(r"rev[\s._-]*([0-9a-z]+)", archivo, re.I)
        avisos = []
        if "referencia" not in mapa:
            avisos.append("El archivo no trae la columna Referencia (motivo); los análisis por motivo quedan vacíos.")
        return LogDesviaciones(archivo, fecha_corte, min(fechas) if fechas else None, rev.group(1).upper() if rev else "",
                               hoja, df, calidad, unif, totales, avisos)
    
    
    def _control_calidad(df: pd.DataFrame) -> pd.DataFrame:
        p = []
    
        def agregar(mascara, problema, detalle=lambda r: ""):
            for _, r in df[mascara].iterrows():
                p.append({"n": r["n"], "fila_excel": int(r["fila_excel"]), "fecha": r["fecha"], "problema": problema, "detalle": detalle(r)})
    
        sin_hh = df["hh"].isna() & df["hh_gestionadas"].isna()
        agregar(sin_hh, "Sin HH desviadas", lambda r: r["actividad"][:90])
        agregar(df["fecha"].isna(), "Sin fecha")
        agregar(df["referencia"] == "", "Sin motivo (Referencia)")
        agregar(df["especialidad"] == "", "Sin especialidad")
        agregar((df["hh"].fillna(0) > 0) & (df["dotacion"].fillna(0) <= 0), "HH sin dotación", lambda r: f"{r['hh']:g} HH")
        hh_p = df["hh"] / df["dotacion"].where(df["dotacion"] > 0)
        agregar(hh_p > HH_MAX_PERSONA, f"Más de {HH_MAX_PERSONA:g} HH por persona",
                lambda r: f"{r['hh']:g} HH / {r['dotacion']:g} personas")
        num = df["n"][df["n"] != ""]
        agregar(df["n"].isin(num[num.duplicated()]), "N° repetido")
        claves = ["fecha", "jornada", "especialidad", "area", "actividad", "dotacion", "hh"]
        agregar(df.duplicated(claves, keep="first") & df["actividad"].ne(""), "Posible registro duplicado",
                lambda r: "misma fecha, jornada, área, actividad, dotación y HH que otro registro")
        return pd.DataFrame(p, columns=["n", "fila_excel", "fecha", "problema", "detalle"])
    
    
    # ------------------------------------------------------------------ análisis
    
    def filtrar(log: LogDesviaciones, desde: date | None = None, hasta: date | None = None, especialidades=None,
                motivos=None, jornadas=None, areas=None) -> pd.DataFrame:
        d = log.df
        if desde:
            d = d[d["fecha"].notna() & (d["fecha"] >= desde)]
        if hasta:
            d = d[d["fecha"].notna() & (d["fecha"] <= hasta)]
        for col, vals in (("especialidad", especialidades), ("referencia", motivos), ("jornada", jornadas), ("area", areas)):
            if vals:
                d = d[d[col].isin(vals)]
        return d
    
    
    def indicadores(d: pd.DataFrame, desde: date | None = None, hasta: date | None = None) -> dict:
        hh = float(d["hh"].sum())
        dias_dato = d["fecha"].dropna().nunique()
        if desde and hasta:
            dias = (hasta - desde).days + 1
        else:
            dias = dias_dato
        diario = d.groupby("fecha")["hh"].sum()
        out = {"hh": hh, "registros": int(len(d)), "dotacion": float(d["dotacion"].sum()),
               "hh_gestionadas": float(d["hh_gestionadas"].sum()), "dias": dias, "dias_con_registro": dias_dato,
               "hh_dia": hh / dias if dias else 0.0, "pico_fecha": None, "pico_hh": 0.0,
               "motivo_top": None, "motivo_top_hh": 0.0, "esp_top": None, "esp_top_hh": 0.0,
               "hh_por_persona": hh / float(d["dotacion"].sum()) if d["dotacion"].sum() else 0.0}
        if not diario.empty:
            out["pico_fecha"], out["pico_hh"] = diario.idxmax(), float(diario.max())
        for campo, k in (("referencia", "motivo_top"), ("especialidad", "esp_top")):
            g = d[d[campo] != ""].groupby(campo)["hh"].sum()
            if not g.empty and g.max() > 0:
                out[k], out[k + "_hh"] = g.idxmax(), float(g.max())
        return out
    
    
    def por(d: pd.DataFrame, campo: str, vacio: str = "(sin dato)") -> pd.DataFrame:
        if d.empty:
            return pd.DataFrame(columns=[campo, "hh", "registros", "dotacion", "hh_gestionadas", "participacion"])
        g = d.assign(**{campo: d[campo].replace("", vacio)}).groupby(campo).agg(
            hh=("hh", "sum"), registros=("hh", "size"), dotacion=("dotacion", "sum"), hh_gestionadas=("hh_gestionadas", "sum"))
        g = g.reset_index().sort_values("hh", ascending=False)
        total = g["hh"].sum()
        g["participacion"] = g["hh"] / total * 100 if total else 0.0
        return g
    
    
    def diario(d: pd.DataFrame) -> pd.DataFrame:
        """HH por fecha y jornada, con acumulado."""
        if d.empty:
            return pd.DataFrame(columns=["fecha", "Total", "Acumulado"])
        t = d[d["fecha"].notna()].assign(jornada=lambda x: x["jornada"].replace("", "Sin jornada"))
        p = t.pivot_table(index="fecha", columns="jornada", values="hh", aggfunc="sum", fill_value=0)
        orden = [c for c in ("Día", "Noche") if c in p] + [c for c in p.columns if c not in ("Día", "Noche")]
        p = p[orden]
        p["Total"] = p.sum(axis=1)
        p = p.sort_index().reset_index()
        p["Acumulado"] = p["Total"].cumsum()
        return p
    
    
    def semanal(d: pd.DataFrame) -> pd.DataFrame:
        if d.empty:
            return pd.DataFrame(columns=["semana", "desde", "hh", "registros"])
        ini = CAL.configuracion()
        t = d[d["fecha"].notna()].copy()
        t["n_sem"] = [CAL.semana(x, ini) for x in t["fecha"]]
        g = t.groupby("n_sem").agg(hh=("hh", "sum"), registros=("hh", "size")).reset_index()
        g["desde"] = [CAL.inicio_semana(int(n), ini) for n in g["n_sem"]]
        g["semana"] = [f"S{int(n):02d} ({CAL.inicio_semana(int(n), ini):%d-%m} al {CAL.fin_semana(int(n), ini):%d-%m})" for n in g["n_sem"]]
        return g[["semana", "desde", "hh", "registros"]]
    
    
    def matriz(d: pd.DataFrame, filas: str = "especialidad", columnas: str = "referencia") -> pd.DataFrame:
        if d.empty:
            return pd.DataFrame()
        t = d.assign(**{filas: d[filas].replace("", "(sin dato)"), columnas: d[columnas].replace("", "(sin dato)")})
        p = t.pivot_table(index=filas, columns=columnas, values="hh", aggfunc="sum", fill_value=0)
        p = p.loc[p.sum(axis=1).sort_values(ascending=False).index, p.sum().sort_values(ascending=False).index]
        p["Total"] = p.sum(axis=1)
        return p
    
    
    def mayores(d: pd.DataFrame, n: int = 15) -> pd.DataFrame:
        return d.sort_values("hh", ascending=False).head(n)
    
    
    def diferencias_totales(log: LogDesviaciones, tolerancia: float = 0.5) -> list[dict]:
        return [t for t in log.totales if abs((t["planilla"] or 0) - (t["calculado"] or 0)) > tolerancia]
    
    
    def comparar(nuevo: LogDesviaciones, anterior: LogDesviaciones) -> pd.DataFrame:
        """Registros agregados, eliminados o modificados entre dos revisiones del LOG (por N°; sin N°, por contenido)."""
        campos = ["fecha", "jornada", "especialidad", "area", "actividad", "dotacion", "hh", "hh_gestionadas", "referencia"]
    
        def clave(df):
            k = df["n"].where(df["n"] != "", None)
            alt = df[["fecha", "jornada", "area", "actividad"]].astype(str).agg("|".join, axis=1)
            return k.fillna("~" + alt)
    
        a = anterior.df.assign(_k=clave(anterior.df)).drop_duplicates("_k").set_index("_k")
        b = nuevo.df.assign(_k=clave(nuevo.df)).drop_duplicates("_k").set_index("_k")
        filas = []
        for k in b.index.difference(a.index, sort=False):
            r = b.loc[k]
            filas.append({"cambio": "Agregado", "n": r["n"], "fecha": r["fecha"], "especialidad": r["especialidad"],
                          "actividad": r["actividad"], "hh_antes": None, "hh_ahora": r["hh"], "detalle": ""})
        for k in a.index.difference(b.index, sort=False):
            r = a.loc[k]
            filas.append({"cambio": "Eliminado", "n": r["n"], "fecha": r["fecha"], "especialidad": r["especialidad"],
                          "actividad": r["actividad"], "hh_antes": r["hh"], "hh_ahora": None, "detalle": ""})
        for k in b.index.intersection(a.index, sort=False):
            ra, rb = a.loc[k], b.loc[k]
            cambios = [f"{NOMBRES[c]}: {_txt(ra[c])} → {_txt(rb[c])}" for c in campos if not _igual(ra[c], rb[c])]
            if cambios:
                filas.append({"cambio": "Modificado", "n": rb["n"], "fecha": rb["fecha"], "especialidad": rb["especialidad"],
                              "actividad": rb["actividad"], "hh_antes": ra["hh"], "hh_ahora": rb["hh"], "detalle": "; ".join(cambios)})
        return pd.DataFrame(filas, columns=["cambio", "n", "fecha", "especialidad", "actividad", "hh_antes", "hh_ahora", "detalle"])
    
    
    def _vacio(v) -> bool:
        return v is None or v == "" or (isinstance(v, float) and pd.isna(v))
    
    
    def _igual(a, b) -> bool:
        if _vacio(a) and _vacio(b):
            return True
        if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not _vacio(a) and not _vacio(b):
            return abs(float(a) - float(b)) < 1e-9
        return a == b
    
    
    def _txt(v) -> str:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "–"
        if isinstance(v, float):
            return f"{v:g}"
        if isinstance(v, date):
            return v.strftime("%d-%m-%Y")
        s = str(v)
        return s if len(s) <= 60 else s[:57] + "…"
    
    
    def _n(x, dec=0) -> str:
        s = f"{x:,.{dec}f}"
        return s.replace(",", "§").replace(".", ",").replace("§", ".")
    
    
    def _f(x) -> str:
        return x.strftime("%d-%m-%Y") if x else "–"
    
    
    def cruce_tm(d: pd.DataFrame, mo_tm: pd.DataFrame) -> dict | None:
        """Participación de las HH desviadas en las HH de mano de obra directa (construcción) del T&M, en las mismas fechas."""
        if mo_tm is None or mo_tm.empty or d.empty:
            return None
        directa = mo_tm[(mo_tm["tipo"] == "Directo") & (mo_tm["ambito"] == "Construcción")]
        fechas = set(d["fecha"].dropna())
        comunes = sorted(fechas & set(directa["fecha"]))
        if not comunes:
            return None
        hh_tm = float(directa[directa["fecha"].isin(comunes)]["hh"].sum())
        hh_dv = float(d[d["fecha"].isin(comunes)]["hh"].sum())
        dia_tm = directa[directa["fecha"].isin(comunes)].groupby("fecha")["hh"].sum()
        dia_dv = d[d["fecha"].isin(comunes)].groupby("fecha")["hh"].sum()
        serie = pd.DataFrame({"HH directas T&M": dia_tm, "HH desviadas": dia_dv}).fillna(0).reset_index().rename(columns={"index": "fecha"})
        serie["% desviado"] = serie["HH desviadas"] / serie["HH directas T&M"].where(serie["HH directas T&M"] > 0) * 100
        return {"desde": comunes[0], "hasta": comunes[-1], "dias": len(comunes), "hh_tm": hh_tm, "hh_desviadas": hh_dv,
                "pct": hh_dv / hh_tm * 100 if hh_tm else None, "serie": serie}
    
    
    def resumen_narrativo(log: LogDesviaciones, d: pd.DataFrame, ind: dict, desde, hasta, cruce: dict | None = None,
                          valor_hh: float | None = None) -> list[str]:
        if d.empty:
            return ["No hay registros de desviación en el período y filtros seleccionados."]
        s1, s2 = CAL.semana(desde), CAL.semana(hasta)
        sem = f" (semana {s1} del proyecto)" if s1 == s2 else f" (semanas {s1} a {s2} del proyecto)"
        fr = [f"Entre el {_f(desde)} y el {_f(hasta)}{sem if s1 and s1 >= 1 else ''} se registraron {_n(ind['registros'])} desviaciones que suman "
              f"{_n(ind['hh'])} HH, con una dotación acumulada de {_n(ind['dotacion'])} personas-turno "
              f"(promedio {_n(ind['hh_dia'], 1)} HH por día calendario)."]
        if ind["pico_fecha"]:
            fr.append(f"El día con más HH desviadas fue el {_f(ind['pico_fecha'])}, con {_n(ind['pico_hh'], 1)} HH.")
        mot = por(d[d["referencia"] != ""], "referencia")
        if not mot.empty:
            top = mot.head(3)
            partes = [f"{r.referencia} ({_n(r.hh)} HH, {_n(r.participacion, 1)}%)" for r in top.itertuples()]
            fr.append(f"Los motivos principales son {', '.join(partes[:-1]) + ' y ' + partes[-1] if len(partes) > 1 else partes[0]}"
                      f", que concentran el {_n(top['participacion'].sum(), 1)}% del total.")
        esp = por(d[d["especialidad"] != ""], "especialidad")
        if not esp.empty:
            r = esp.iloc[0]
            fr.append(f"La especialidad más afectada es {r['especialidad']} con {_n(r['hh'])} HH ({_n(r['participacion'], 1)}%).")
        jor = por(d[d["jornada"] != ""], "jornada")
        if len(jor) > 1:
            fr.append("Por jornada: " + ", ".join(f"{r.jornada} {_n(r.hh)} HH ({_n(r.participacion, 1)}%)" for r in jor.itertuples()) + ".")
        if ind["hh_gestionadas"]:
            fr.append(f"Se registran {_n(ind['hh_gestionadas'], 1)} HH desviadas gestionadas.")
        if cruce and cruce.get("pct") is not None:
            fr.append(f"En los {cruce['dias']} días con datos en el Times & Material, las HH desviadas equivalen al "
                      f"{_n(cruce['pct'], 1)}% de las HH de mano de obra directa de construcción "
                      f"({_n(cruce['hh_desviadas'])} de {_n(cruce['hh_tm'])} HH).")
        if valor_hh:
            fr.append(f"A un valor referencial de ${_n(valor_hh)} por HH, las desviaciones del período representan "
                      f"${_n(ind['hh'] * valor_hh)} (neto).")
        dif = diferencias_totales(log)
        if dif:
            fr.append("Los totales de la planilla original no cuadran con el detalle: " + "; ".join(
                f"{t['origen']} indica {_n(t['planilla'], 1)} {NOMBRES[t['campo']]} y el detalle suma {_n(t['calculado'], 1)}"
                for t in dif[:2]) + ". Este informe usa los valores del detalle.")
        return fr
    
    
    def tabla_detalle(d: pd.DataFrame) -> pd.DataFrame:
        """Detalle con nombres legibles, para mostrar y exportar."""
        t = d.drop(columns=["fila_excel"], errors="ignore").rename(columns=NOMBRES)
        return t[[c for c in NOMBRES.values() if c in t]]
    
    
    # ------------------------------------------------------------------ cruce con el informe semanal (PF)
    
    def impacto_pf(log: LogDesviaciones, gastadas: float, ganadas: float, hasta: date, desde: date | None = None,
                   especialidades=None, motivos=None) -> dict | None:
        """Cuánto habría sido el PF si las HH desviadas no se hubieran gastado.
    
        PF = HH gastadas / HH ganadas. Las HH desviadas se gastaron en trabajos fuera de alcance que no generan avance,
        así que el PF sin desviaciones es (HH gastadas - HH desviadas) / HH ganadas.
        desde=None: acumulado (todo el LOG hasta la fecha de corte). Con desde: solo las desviaciones de (desde, hasta].
        """
        if not ganadas or ganadas <= 0 or gastadas is None or pd.isna(gastadas):
            return None
        d = filtrar(log, None, hasta, especialidades, motivos)
        if desde is not None:
            d = d[d["fecha"] > desde]
        hh = float(d["hh"].sum())
        pf = gastadas / ganadas
        pf_sin = (gastadas - hh) / ganadas
        avisos = []
        if log.fecha_inicio and (desde is None or log.fecha_inicio > desde + timedelta(days=1)):
            avisos.append(f"El LOG registra desviaciones desde el {_f(log.fecha_inicio)}; las anteriores no se descuentan, "
                          "así que la mejora real del PF podría ser mayor.")
        if log.fecha_corte and log.fecha_corte < hasta:
            avisos.append(f"El LOG llega hasta el {_f(log.fecha_corte)} y el informe corta el {_f(hasta)}: faltan días por registrar.")
        if hh > gastadas:
            avisos.append("Las HH desviadas superan las HH gastadas del período: revisa que el LOG y el informe tengan el mismo alcance.")
        mot = por(d, "referencia", "(sin motivo)")
        mot["puntos_pf"] = mot["hh"] / ganadas
        esp = por(d, "especialidad")
        esp["puntos_pf"] = esp["hh"] / ganadas
        return {"desde": desde, "hasta": hasta, "gastadas": float(gastadas), "ganadas": float(ganadas), "hh_desviadas": hh,
                "registros": int(len(d)), "pf": pf, "pf_sin": pf_sin, "mejora": pf - pf_sin,
                "mejora_pct": (pf - pf_sin) / pf * 100 if pf else 0.0,
                "pct_gastadas": hh / gastadas * 100 if gastadas else 0.0, "por_motivo": mot, "por_especialidad": esp,
                "avisos": avisos}
    
    
    def texto_impacto(r: dict, periodo: str) -> str:
        return (f"{periodo}: el PF fue {_n(r['pf'], 2)}. Sin las {_n(r['hh_desviadas'])} HH desviadas "
                f"({_n(r['pct_gastadas'], 1)}% de las HH gastadas) habría sido {_n(r['pf_sin'], 2)}, "
                f"una mejora de {_n(r['mejora'], 2)} puntos ({_n(r['mejora_pct'], 1)}%).")
    
    
    def recuento_semana(log: LogDesviaciones, desde: date, hasta: date, especialidades=None, motivos=None) -> dict:
        """Desviaciones de la semana del informe (desde y hasta inclusive) y comparación con la semana anterior."""
        d = filtrar(log, desde, hasta, especialidades, motivos)
        largo = (hasta - desde).days + 1
        prev = filtrar(log, desde - timedelta(days=largo), desde - timedelta(days=1), especialidades, motivos)
        ind = indicadores(d, desde, hasta)
        hh_prev = float(prev["hh"].sum()) if len(prev) else None
        texto = []
        if d.empty:
            texto.append(f"No hay desviaciones registradas entre el {_f(desde)} y el {_f(hasta)}.")
        else:
            texto.append(f"Entre el {_f(desde)} y el {_f(hasta)} se registraron {_n(ind['registros'])} desviaciones por "
                         f"{_n(ind['hh'])} HH, en {ind['dias_con_registro']} de {largo} días.")
            if hh_prev:
                v = (ind["hh"] / hh_prev - 1) * 100
                texto.append(f"Frente a la semana anterior ({_n(hh_prev)} HH) es {'un aumento' if v >= 0 else 'una baja'} de {_n(abs(v), 1)}%.")
            if ind["pico_fecha"]:
                texto.append(f"El día con más HH desviadas fue el {_f(ind['pico_fecha'])} ({_n(ind['pico_hh'], 1)} HH).")
            mot = por(d[d["referencia"] != ""], "referencia")
            if not mot.empty:
                texto.append("Principales motivos: " + ", ".join(f"{r.referencia} {_n(r.hh)} HH ({_n(r.participacion, 1)}%)"
                                                                for r in mot.head(3).itertuples()) + ".")
            esp = por(d[d["especialidad"] != ""], "especialidad")
            if not esp.empty:
                texto.append("Especialidades más afectadas: " + ", ".join(f"{r.especialidad} {_n(r.hh)} HH"
                                                                          for r in esp.head(3).itertuples()) + ".")
        return {"d": d, "ind": ind, "hh_prev": hh_prev, "por_motivo": por(d, "referencia", "(sin motivo)"),
                "por_especialidad": por(d, "especialidad"), "diario": diario(d), "mayores": mayores(d, 10), "texto": texto,
                "desde": desde, "hasta": hasta}

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


desviaciones = _modulo_desviaciones()

# ======================================================================
# MÓDULO: desviaciones_pdf
# ======================================================================
def _modulo_desviaciones_pdf():
    import io
    import os
    from datetime import date, datetime
    
    import pandas as pd
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.platypus import KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer
    from xml.sax.saxutils import escape
    
    CAL = calendario
    D = desviaciones
    P = informe_pdf
    TP = tm_pdf
    
    
    def _f(x) -> str:
        return x.strftime("%d-%m-%Y") if isinstance(x, date) else "–"
    
    
    def _semanas(desde, hasta) -> str:
        a, b = CAL.etiqueta(desde), CAL.etiqueta(hasta)
        return f"semana {a}" if a == b else f"semanas {a} a {b}"
    
    
    def generar_informe(log: D.LogDesviaciones, d: pd.DataFrame, desde: date, hasta: date, filtros: str = "",
                        autor: str = "", contrato: str = "", observaciones: str = "", conclusiones: str = "",
                        cruce: dict | None = None, valor_hh: float | None = None, anexo: bool = True,
                        logo: str | None = None, titulo: str = "") -> bytes:
        titulo = titulo or "INFORME DE DESVIACIÓN DE RECURSOS"
        ind = D.indicadores(d, desde, hasta)
        narr = D.resumen_narrativo(log, d, ind, desde, hasta, cruce, valor_hh)
        generado = datetime.now().strftime("%d-%m-%Y %H:%M")
    
        def marco(canvas, doc):
            A, H = A4
            canvas.saveState()
            canvas.setFillColor(P.AZUL)
            canvas.rect(0, H - 20 * mm, A, 20 * mm, stroke=0, fill=1)
            canvas.setFillColor(P.NARANJO)
            canvas.rect(0, H - 21.2 * mm, A, 1.2 * mm, stroke=0, fill=1)
            if logo and os.path.exists(logo):
                canvas.drawImage(P.logo_fondo_oscuro(logo), P.MARGEN, H - 17 * mm, height=14 * mm, width=40 * mm, preserveAspectRatio=True,
                                 anchor="w", mask="auto")
            else:
                canvas.setFillColor(P.NARANJO)
                canvas.setFont("Helvetica-Bold", 17)
                canvas.drawString(P.MARGEN, H - 12.5 * mm, "SYNCORE")
            canvas.setFillColor(colors.white)
            canvas.setFont("Helvetica-Bold", 10.5)
            canvas.drawRightString(A - P.MARGEN, H - 10.5 * mm, titulo[:60])
            canvas.setFont("Helvetica", 8)
            canvas.setFillColor(colors.HexColor("#CBD5E1"))
            canvas.drawRightString(A - P.MARGEN, H - 15 * mm, (f"{contrato}  ·  " if contrato else "") + f"Período {_f(desde)} al {_f(hasta)}")
            canvas.setStrokeColor(P.GRIS_CLARO)
            canvas.line(P.MARGEN, 13 * mm, A - P.MARGEN, 13 * mm)
            canvas.setFont("Helvetica", 7)
            canvas.setFillColor(P.GRIS)
            canvas.drawString(P.MARGEN, 9 * mm, f"SYNCORE · Fuente: {log.archivo[:70]} · Generado el {generado}" + (f" por {autor}" if autor else ""))
            canvas.drawRightString(A - P.MARGEN, 9 * mm, f"Página {doc.page}")
            canvas.restoreState()
    
        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=P.MARGEN, rightMargin=P.MARGEN, topMargin=28 * mm,
                                bottomMargin=18 * mm, title=titulo, author=autor or "SYNCORE")
        E = P.E
        rev = f" · Revisión {log.revision}" if log.revision else ""
        h = [Paragraph(escape(titulo.capitalize() if titulo.isupper() else titulo), E["titulo"]),
             Paragraph(escape(f"Período informado: {_f(desde)} al {_f(hasta)} ({_semanas(desde, hasta)}) · "
                              f"LOG al {_f(log.fecha_corte)}{rev}"), E["subtitulo"])]
        if filtros:
            h.append(Paragraph(escape(f"Filtros: {filtros}"), E["subtitulo"]))
        h.append(Spacer(1, 6 * mm))
        h.append(P._fila_kpis([("HH desviadas", P._n(ind["hh"]), True), ("Registros", P._n(ind["registros"]), False),
                               ("Dotación (personas-turno)", P._n(ind["dotacion"]), False),
                               ("HH promedio por día", P._n(ind["hh_dia"], 1), False)]))
        h.append(Spacer(1, 3 * mm))
        k2 = [("Día con más HH", f"{P._n(ind['pico_hh'])} · {_f(ind['pico_fecha'])}" if ind["pico_fecha"] else "–", False),
              ("Motivo principal", str(ind["motivo_top"] or "–"), False),
              ("HH gestionadas", P._n(ind["hh_gestionadas"], 1), False)]
        if cruce and cruce.get("pct") is not None:
            k2.append(("% de HH directas T&M", f"{P._n(cruce['pct'], 1)}%", False))
        elif valor_hh:
            k2.append(("Valorización referencial", f"${P._n(ind['hh'] * valor_hh / 1e6, 1)} MM", False))
        h.append(P._fila_kpis(k2))
    
        h.append(Paragraph("1. Resumen ejecutivo", E["h2"]))
        for frase in narr:
            h.append(Paragraph(escape(frase), E["vineta"], bulletText="•"))
        if observaciones.strip():
            h.append(P._seccion("Observaciones", P._caja(observaciones)))
    
        # ---- evolución
        dia = D.diario(d)
        if not dia.empty:
            h.append(PageBreak())
            h.append(Paragraph("2. Evolución en el período", E["h2"]))
            jornadas = [c for c in dia.columns if c not in ("fecha", "Total", "Acumulado")]
            h.append(TP.grafico_barras_apiladas(list(dia["fecha"]), [(j, list(dia[j])) for j in jornadas], "HH desviadas por día"))
            h.append(Spacer(1, 2 * mm))
            h.append(TP.grafico_linea(list(dia["fecha"]), [("HH desviadas acumuladas", list(dia["Acumulado"]))], "HH acumuladas"))
            sem = D.semanal(d)
            if len(sem) > 1:
                h.append(P._seccion("HH desviadas por semana", P._tabla(sem, [("semana", "Semana", 2, "izq"),
                                                                            ("registros", "Registros", 1, "num"),
                                                                            ("hh", "HH desviadas", 1.2, "num1")], "")))
    
        # ---- motivo y especialidad
        mot = D.por(d, "referencia", "(sin motivo)")
        esp = D.por(d, "especialidad")
        if not mot.empty:
            h.append(PageBreak())
            h.append(Paragraph("3. Desviaciones por motivo y especialidad", E["h2"]))
            h.append(KeepTogether([Paragraph("HH por motivo", E["h2"]),
                                   TP.grafico_ranking(list(mot["referencia"]), list(mot["hh"]))]))
            cols = [("hh", "HH", 1, "num1"), ("participacion", "% del total", 0.9, "pct"), ("registros", "Registros", 0.8, "num"),
                    ("dotacion", "Dotación", 0.8, "num")]
            h.append(P._seccion("Detalle por motivo", P._tabla(mot, [("referencia", "Motivo", 2.2, "izq")] + cols, "")))
            h.append(KeepTogether([Paragraph("HH por especialidad", E["h2"]),
                                   TP.grafico_ranking(list(esp["especialidad"]), list(esp["hh"]), colores=[P.AZUL] * len(esp))]))
            h.append(P._seccion("Detalle por especialidad", P._tabla(esp, [("especialidad", "Especialidad", 2.2, "izq")] + cols, "")))
            m = D.matriz(d)
            if not m.empty:
                motivos_m = [c for c in m.columns if c != "Total"]
                if len(motivos_m) > 8:                     # los 7 motivos principales y el resto agrupado
                    m = pd.concat([m[motivos_m[:7]], m[motivos_m[7:]].sum(axis=1).rename("Otros"), m["Total"]], axis=1)
                t = m.reset_index()
                motivos = [c for c in m.columns if c != "Total"]
                ancho = [("especialidad", "Especialidad", 1.6, "izq")] + [(c, str(c)[:12], 0.75, "num") for c in motivos] + \
                        [("Total", "Total", 0.8, "num")]
                h.append(P._seccion("HH por especialidad y motivo", P._tabla(t, ancho, "")))
    
        area = D.por(d, "area").head(15)
        if not area.empty:
            h.append(KeepTogether([Paragraph("4. Áreas con más HH desviadas", E["h2"]),
                                   TP.grafico_ranking(list(area["area"]), list(area["hh"]), colores=[TP.AZUL_MEDIO] * len(area))]))
        top = D.mayores(d, 12)
        if not top.empty:
            h.append(P._seccion("Registros con más HH desviadas", P._tabla(top, [
                ("fecha", "Fecha", 0.8, "fecha"), ("especialidad", "Especialidad", 1, "izq"), ("area", "Área", 1.1, "izq"),
                ("actividad", "Actividad", 3.6, "izq"), ("referencia", "Motivo", 0.9, "izq"), ("dotacion", "Dot.", 0.5, "num"),
                ("hh", "HH", 0.6, "num1")], "")))
    
        if cruce and cruce.get("pct") is not None:
            s = cruce["serie"]
            h.append(PageBreak())
            h.append(Paragraph("5. Comparación con el Times & Material", E["h2"]))
            h.append(Paragraph(escape(f"Del {_f(cruce['desde'])} al {_f(cruce['hasta'])} ({cruce['dias']} días con datos en ambos registros), "
                                      f"las HH desviadas suman {P._n(cruce['hh_desviadas'])} HH frente a {P._n(cruce['hh_tm'])} HH de "
                                      f"mano de obra directa de construcción: {P._n(cruce['pct'], 1)}%."), E["cuerpo"]))
            h.append(Spacer(1, 2 * mm))
            h.append(TP.grafico_linea(list(s["fecha"]), [("HH directas T&M", list(s["HH directas T&M"])),
                                                         ("HH desviadas", list(s["HH desviadas"]))], "HH por día"))
    
        cal = log.calidad
        dif = D.diferencias_totales(log)
        if not cal.empty or dif or log.unificaciones:
            h.append(Paragraph("Control de calidad del LOG", E["h2"]))
            for t in dif:
                h.append(Paragraph(escape(f"{t['origen']}: {P._n(t['planilla'], 1)} {D.NOMBRES[t['campo']]} en la planilla; "
                                          f"el detalle suma {P._n(t['calculado'], 1)}."), E["vineta"], bulletText="•"))
            if log.unificaciones:
                h.append(Paragraph(escape(f"Se unificaron {len(log.unificaciones)} nombres escritos de distintas formas "
                                          "(mayúsculas, tildes o espacios)."), E["vineta"], bulletText="•"))
            if not cal.empty:
                h.append(P._tabla(cal, [("n", "N°", 0.5, "izq"), ("fila_excel", "Fila", 0.5, "num"), ("fecha", "Fecha", 0.8, "fecha"),
                                        ("problema", "Observación", 1.6, "izq"), ("detalle", "Detalle", 3, "izq")], ""))
    
        if conclusiones.strip():
            h.append(P._seccion("Conclusiones y acciones", P._caja(conclusiones)))
    
        if anexo and not d.empty:
            h.append(PageBreak())
            h.append(Paragraph("Anexo. Detalle de registros por motivo", E["h2"]))
            orden = list(mot["referencia"]) if not mot.empty else []
            dd = d.assign(referencia=d["referencia"].replace("", "(sin motivo)"))
            for motivo in orden:
                g = dd[dd["referencia"] == motivo].sort_values(["fecha", "n"])
                h.append(Paragraph(escape(f"{motivo} — {P._n(g['hh'].sum(), 1)} HH en {len(g)} registros"), E["h2"]))
                h.append(P._tabla(g, [("n", "N°", 0.45, "izq"), ("fecha", "Fecha", 1.0, "fecha"), ("jornada", "Jornada", 0.75, "izq"),
                                      ("especialidad", "Especialidad", 1.05, "izq"), ("area", "Área", 1.25, "izq"),
                                      ("actividad", "Actividad", 2.85, "izq"), ("observacion", "Observación", 1.3, "izq"),
                                      ("dotacion", "Dot.", 0.45, "num"), ("hh", "HH", 0.55, "num1")], ""))
    
        h.append(Spacer(1, 6 * mm))
        h.append(Paragraph(escape(f"Fuente: LOG de desviación de recursos, hoja «{log.hoja.strip()}». Las cifras se calculan desde el "
                                  "detalle de registros, no desde las tablas dinámicas de la planilla."), E["subtitulo"]))
        doc.build(h, onFirstPage=marco, onLaterPages=marco)
        return buf.getvalue()

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


desviaciones_pdf = _modulo_desviaciones_pdf()

# ======================================================================
# MÓDULO: cubicacion
# ======================================================================
def _modulo_cubicacion():
    import io
    import re
    import unicodedata
    from dataclasses import asdict, dataclass, field
    from datetime import date, datetime
    
    DISCIPLINAS = ["Obras civiles", "Estructura", "Mecánica", "Piping", "Eléctricos"]
    OTRAS = "Otras"
    _RE_ITEM = re.compile(r"^\d{2}\.\d{2}\.\d{3}\.\d{3,4}[A-Z]?(\.\d+)?$")
    _EPS = 1e-6
    
    
    def norm(s) -> str:
        s = unicodedata.normalize("NFKD", str(s if s is not None else "")).encode("ascii", "ignore").decode().upper()
        return re.sub(r"\s+", " ", s.replace("\n", " ")).strip()
    
    
    def _num(v) -> float:
        if isinstance(v, bool) or v is None:
            return 0.0
        if isinstance(v, (int, float)):
            return float(v) if v == v else 0.0          # NaN → 0
        s = str(v).strip().replace(" ", "")
        if not s or s.startswith("#"):
            return 0.0
        if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", s):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return 0.0
    
    
    def _txt(v) -> str:
        if v is None:
            return ""
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return re.sub(r"\s+", " ", str(v)).strip()
    
    
    def disciplina_de(especialidad: str, item: str = "") -> str:
        """Disciplina a partir de la columna Especialidad del itemizado (o del capítulo del ítem si viene vacía)."""
        e = norm(especialidad)
        if "CIVIL" in e:
            return "Obras civiles"
        if "ESTRUC" in e:
            return "Estructura"
        if "MECAN" in e:
            return "Mecánica"
        if "CANER" in e or "PIPING" in e:
            return "Piping"
        if "ELECTR" in e or "INSTRUM" in e:
            return "Eléctricos"
        if e:
            return OTRAS
        cap = item[:2]
        return {"00": "Obras civiles", "10": "Obras civiles", "20": "Estructura", "30": "Mecánica", "40": "Mecánica",
                "50": "Mecánica", "51": "Piping", "52": "Piping", "53": "Piping", "81": "Piping",
                "60": "Eléctricos", "70": "Eléctricos"}.get(cap, OTRAS)
    
    
    @dataclass
    class ItemCub:
        item: str
        disciplina: str
        especialidad: str = ""
        hoja_control: str = ""
        descripcion: str = ""
        unidad: str = ""
        contrato: float = 0.0        # cantidad del contrato (itemizado EC-002)
        rend: float = 0.0            # HH por unidad
        pom: float = 0.0             # cubicación POM
        anterior: float = 0.0        # avance acumulado weekly de la semana anterior
        acumulado: float = 0.0       # avance acumulado weekly a la fecha
        edp: float = 0.0             # avance acumulado EDP
        planilla: str = ""           # archivo de origen
    
        @property
        def semana(self) -> float:
            return self.acumulado - self.anterior
    
        @property
        def pom_fuera_de_escala(self) -> bool:
            # POM 20 veces sobre el contrato y un avance que sigue al contrato (no al POM): error de digitación
            return self.contrato > _EPS and self.pom > self.contrato * 20 and self.acumulado <= self.contrato * 1.5
    
        @property
        def pom_valida(self) -> float:
            """POM para calcular HH; si está fuera de escala (error de la planilla) se usa la cantidad del contrato."""
            return self.contrato if self.pom_fuera_de_escala else self.pom
    
        @property
        def hh_pom(self) -> float:
            return self.pom_valida * self.rend
    
        @property
        def hh_ganadas(self) -> float:
            return self.acumulado * self.rend
    
    
    @dataclass
    class FilaDetalle:
        hoja: str
        fila: int
        item: str
        plano: str = ""
        rev: str = ""
        tag: str = ""
        descripcion: str = ""
        unidad: str = ""
        cantidad: float | None = None      # cubicación de la fila (POM / peso total / cantidad)
        acumulado: float | None = None     # avance acumulado weekly de la fila (si la hoja lo trae)
    
    
    @dataclass
    class Hallazgo:
        disciplina: str
        tipo: str
        item: str
        detalle: str
        planilla: str = ""
        gravedad: str = "Revisar"          # «Error», «Revisar» o «Info»
    
    
    @dataclass
    class PlanillaCub:
        archivo: str
        hoja_itemizado: str
        items: dict[str, ItemCub] = field(default_factory=dict)
        detalle: list[FilaDetalle] = field(default_factory=list)
        hojas_detalle: dict[str, str] = field(default_factory=dict)     # hoja → aviso de lectura ("" = ok)
        hallazgos: list[Hallazgo] = field(default_factory=list)
        fecha: str = ""                                                 # fecha de corte deducida del nombre (dd-mm-aaaa)
    
        @property
        def disciplinas(self) -> list[str]:
            vistas = {it.disciplina for it in self.items.values()}
            return [d for d in DISCIPLINAS + [OTRAS] if d in vistas]
    
        def a_dict(self) -> dict:
            return {"archivo": self.archivo, "hoja_itemizado": self.hoja_itemizado, "fecha": self.fecha,
                    "items": [asdict(i) for i in self.items.values()], "detalle": [asdict(f) for f in self.detalle],
                    "hojas_detalle": self.hojas_detalle, "hallazgos": [asdict(h) for h in self.hallazgos]}
    
        @classmethod
        def de_dict(cls, d: dict) -> "PlanillaCub":
            p = cls(d["archivo"], d.get("hoja_itemizado", ""), fecha=d.get("fecha", ""))
            p.items = {i["item"]: ItemCub(**i) for i in d.get("items", [])}
            p.detalle = [FilaDetalle(**f) for f in d.get("detalle", [])]
            p.hojas_detalle = dict(d.get("hojas_detalle", {}))
            p.hallazgos = [Hallazgo(**h) for h in d.get("hallazgos", [])]
            return p
    
    
    def fecha_de_nombre(nombre: str) -> str:
        """Última fecha dd-mm-aaaa del nombre del archivo («… (24-10-2026 al 30-10-2026)» → 30-10-2026)."""
        f = re.findall(r"(\d{1,2})[-_.](\d{1,2})[-_.](\d{4})", nombre or "")
        if not f:
            return ""
        d, m, a = f[-1]
        try:
            return date(int(a), int(m), int(d)).strftime("%d-%m-%Y")
        except ValueError:
            return ""
    
    
    # ------------------------------------------------------------------ lectura
    
    def _abrir(datos: bytes):
        import openpyxl
        return openpyxl.load_workbook(io.BytesIO(datos), read_only=True, data_only=True)
    
    
    def _filas(ws, max_col=120, max_row=6000):
        return [list(r) for r in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col, values_only=True)]
    
    
    def _hoja_itemizado(wb) -> str | None:
        return next((h for h in wb.sheetnames if norm(h) == "ITEMIZADO"), None) or \
            next((h for h in wb.sheetnames if "ITEMIZADO" in norm(h)), None)
    
    
    def _leer_itemizado(filas: list[list]) -> list[dict]:
        """Filas del itemizado EC-002 con sus columnas (por nombre de columna)."""
        f_cab = None
        for i, r in enumerate(filas[:60]):
            n = [norm(v) for v in r]
            es_cab = "ITEM" in n and any(x.startswith("ESPECIALIDA") for x in n)
            if es_cab:
                f_cab = i          # si la fila siguiente repite los títulos (títulos en dos filas), manda la de abajo
            elif f_cab is not None:
                break
        if f_cab is None:
            raise ValueError("no se encontró la fila de títulos del itemizado (columnas «Especialidad» e «Ítem»).")
        cab = [norm(v) for v in filas[f_cab]]
    
        def col(*pref, desde=0):
            for p in pref:
                for j in range(desde, len(cab)):
                    if cab[j] == p or (p.endswith("*") and cab[j].startswith(p[:-1])):
                        return j
            return None
    
        c_item, c_esp = cab.index("ITEM"), col("ESPECIALIDA*")
        c_ctrl, c_desc, c_un = col("PLANILLA DE CONTROL*"), col("DESCRIPCION"), col("UNIDAD")
        c_cant, c_rend, c_pom = col("CANTIDAD"), col("RENDIMIENTO"), col("CUBICACION*")
        # grupos de avance en la fila de arriba: el 1.er «AVANCE ACUMULADO» es WEEKLY, el 2.º es EDP
        acu, ant = [], []
        for fg in range(f_cab - 1, max(f_cab - 4, -1), -1):
            grupos = [norm(v) for v in filas[fg]]
            acu = [j for j, g in enumerate(grupos) if g.startswith("AVANCE ACUMULADO") or g.startswith("AVNCE ACUMULADO")]
            ant = [j for j, g in enumerate(grupos) if g.startswith("AVANCE ANTERIOR")]
            if acu:
                break
        if not acu:
            raise ValueError("no se encontró el grupo «AVANCE ACUMULADO» en el itemizado.")
    
        def cant_de(j):           # la columna CANT del grupo (o la del propio título del grupo)
            return next((k for k in range(j, min(j + 3, len(cab))) if cab[k].startswith("CANT")), j)
    
        c_acu, c_ant = cant_de(acu[0]), (cant_de(ant[0]) if ant else None)
        c_edp = cant_de(acu[1]) if len(acu) > 1 else None
    
        def hh_de(c):              # la columna «HH GAN» que sigue a una CANT
            if c is None:
                return None
            return next((k for k in range(c + 1, min(c + 3, len(cab))) if cab[k].startswith("HH GAN")), None)
    
        h_acu, h_ant = hh_de(c_acu), hh_de(c_ant)
        out = []
        for i in range(f_cab + 1, len(filas)):
            r = filas[i]
            g = lambda j: r[j] if j is not None and j < len(r) else None
            item = _txt(g(c_item))
            if not _RE_ITEM.match(item):
                continue
            rend = _num(g(c_rend))
            anterior, acumulado = _num(g(c_ant)), _num(g(c_acu))
            # algunas planillas (mecánica) dejan la CANT en blanco y solo llenan las HH ganadas: cantidad = HH / rend.
            if abs(anterior) < _EPS and rend > _EPS and abs(_num(g(h_ant))) > _EPS:
                anterior = _num(g(h_ant)) / rend
            if abs(acumulado) < _EPS and rend > _EPS and abs(_num(g(h_acu))) > _EPS:
                acumulado = _num(g(h_acu)) / rend
            out.append({"fila": i + 1, "item": item, "especialidad": _txt(g(c_esp)), "control": _txt(g(c_ctrl)),
                        "descripcion": _txt(g(c_desc)), "unidad": _txt(g(c_un)), "contrato": _num(g(c_cant)),
                        "rend": rend, "pom": _num(g(c_pom)), "anterior": anterior,
                        "acumulado": acumulado, "edp": _num(g(c_edp))})
        if not out:
            raise ValueError("el itemizado no tiene ítems del EC-002.")
        return out
    
    
    # columnas de las hojas de control (por el título normalizado)
    def _buscar(cab: list[str], reglas) -> int | None:
        for regla in reglas:
            for j, t in enumerate(cab):
                if t and regla(t):
                    return j
        return None
    
    
    def _cabecera_detalle(filas) -> tuple[int, int] | None:
        """(fila de títulos, columna del ítem): la columna con más códigos EC-002 y la fila de títulos sobre el 1.er dato."""
        conteo: dict[int, list[int]] = {}
        for i, r in enumerate(filas[:400]):
            for j, v in enumerate(r[:40]):
                if v is not None and _RE_ITEM.match(_txt(v)):
                    conteo.setdefault(j, []).append(i)
        if not conteo:
            return None
        j = max(conteo, key=lambda k: len(conteo[k]))
        r0 = conteo[j][0]
        for i in range(r0 - 1, max(r0 - 6, -1), -1):
            textos = sum(1 for v in filas[i] if isinstance(v, str) and re.search("[A-Za-z]{2}", v))
            if textos >= 5:
                return i, j
        return None
    
    
    def _leer_hoja_detalle(nombre: str, filas: list[list]) -> tuple[list[FilaDetalle], str]:
        cab_ = _cabecera_detalle(filas)
        if cab_ is None:
            return [], "sin filas con ítem del EC-002"
        i, c_item = cab_
        cab = [norm(v) for v in filas[i]]
        c_plano = _buscar(cab, [lambda t: t in ("N DE ISOMETRICO", "N ISOMETRICO") or t.startswith("N DE ISOMETRICO"),
                                lambda t: t.startswith("PLANO DETALLE"), lambda t: t.startswith("PLANO"),
                                lambda t: t in ("N DE PLANO", "NDE PLANO", "NPLANO"), lambda t: t == "DWG"])
        c_rev = None
        if c_plano is not None:
            c_rev = next((j for j in range(c_plano + 1, min(c_plano + 3, len(cab))) if cab[j].startswith("REV")), None)
        c_tag = _buscar(cab, [lambda t: t == "TAG", lambda t: t == "C ID" or t == "C_ID", lambda t: t.startswith("TAG")])
        c_desc = _buscar(cab, [lambda t: t in ("DESCRIPCION", "DESCRIPSION"),
                               lambda t: t.startswith("DESCRIPCION MAT") or t == "DESCRIPCION 1",
                               lambda t: t.startswith("DESCRIPCION /") or t.startswith("DESIGNACION"),
                               lambda t: t.startswith("ELEMENTO"), lambda t: t.startswith("EQUIPO / ELEMENTO")])
        c_elem = _buscar(cab, [lambda t: t.startswith("EQUIPO / ELEMENTO"), lambda t: t == "ELEMENTO",
                               lambda t: t.startswith("ELEMENTO /")])
        if c_elem == c_desc:
            c_elem = None
        c_un = _buscar(cab, [lambda t: t.startswith("UNIDAD"), lambda t: t == "UND"])
        c_cant = _buscar(cab, [lambda t: t.startswith("POM"), lambda t: t == "PESO TOTAL", lambda t: t == "CUBICACION",
                               lambda t: t.startswith("CUBICACION PLANO"), lambda t: t.startswith("QUANTITY"),
                               lambda t: t.startswith("CANTIDAD MTS") or t.startswith("CANTIDAD _MTS"),
                               lambda t: t in ("CANT", "CANT.", "CANTIDAD"), lambda t: t == "QTY"])
        c_acu = _buscar(cab, [lambda t: t.startswith("CANTIDAD WEEKLY ACUMULADO"), lambda t: t == "AV TOTAL ACUMULADO",
                              lambda t: t.startswith("TOTAL ACUMULADO"), lambda t: t == "AVANCE ACUMULADO"])
        filas_out = []
        for k in range(i + 1, len(filas)):
            r = filas[k]
            g = lambda j: r[j] if j is not None and j < len(r) else None
            item = _txt(g(c_item))
            if not _RE_ITEM.match(item):
                continue
            filas_out.append(FilaDetalle(
                hoja=nombre, fila=k + 1, item=item, plano=_txt(g(c_plano)), rev=_txt(g(c_rev)), tag=_txt(g(c_tag)),
                descripcion=" · ".join(dict.fromkeys(x for x in (_txt(g(c_desc)), _txt(g(c_elem))) if x))[:140],
                unidad=_txt(g(c_un)),
                cantidad=_num(g(c_cant)) if c_cant is not None else None,
                acumulado=_num(g(c_acu)) if c_acu is not None else None))
        aviso = "" if filas_out else "sin filas con ítem"
        if filas_out and c_acu is None:
            aviso = "sin columna de avance acumulado reconocible"
        return filas_out, aviso
    
    
    def _clave_hoja(s) -> str:
        return re.sub(r"[^A-Z0-9]", "", norm(s))
    
    
    def _resolver_hoja(nombre: str, hojas: list[str]) -> str | None:
        """Hoja del archivo que corresponde a la «Planilla de control» (tolera espacios, guiones y sufijos)."""
        k = _clave_hoja(nombre)
        if not k:
            return None
        exactas = [h for h in hojas if _clave_hoja(h) == k]
        if exactas:
            return exactas[0]
        pref = [h for h in hojas if _clave_hoja(h).startswith(k) or (len(_clave_hoja(h)) > 4 and k.startswith(_clave_hoja(h)))]
        return pref[0] if pref else None
    
    
    def leer_planilla(datos: bytes, archivo: str = "", con_detalle: bool = True) -> PlanillaCub:
        """Lee una planilla de avance (cualquier disciplina): itemizado + hojas de control que usa."""
        wb = _abrir(datos)
        try:
            h_itm = _hoja_itemizado(wb)
            if not h_itm:
                raise ValueError("no tiene la hoja de itemizado EC-002 («Itemizado»).")
            filas = _leer_itemizado(_filas(wb[h_itm], max_col=80, max_row=3000))
            resol = {f["control"]: _resolver_hoja(f["control"], wb.sheetnames) for f in filas if f["control"]}
            p = PlanillaCub(archivo, h_itm, fecha=fecha_de_nombre(archivo))
            controladas = [f for f in filas if resol.get(f["control"])]
            principales = {disciplina_de(f["especialidad"], f["item"]) for f in controladas}
            for f in filas:
                disc = disciplina_de(f["especialidad"], f["item"])
                propia = bool(resol.get(f["control"]))
                con_avance = abs(f["acumulado"]) > _EPS or abs(f["anterior"]) > _EPS
                # sin hoja de control: solo se toma si tiene avance y es de una disciplina que esta planilla controla
                # (las demás filas son copia del itemizado de otras especialidades)
                if not propia and not (con_avance and (disc in principales or not principales)):
                    continue
                it = p.items.get(f["item"])
                if it is None:
                    p.items[f["item"]] = ItemCub(
                        item=f["item"], disciplina=disc, especialidad=f["especialidad"],
                        hoja_control=(resol.get(f["control"]) or f["control"]).strip(), descripcion=f["descripcion"],
                        unidad=f["unidad"], contrato=f["contrato"], rend=f["rend"], pom=f["pom"], anterior=f["anterior"],
                        acumulado=f["acumulado"], edp=f["edp"], planilla=archivo)
                else:                      # ítem repetido en el itemizado (filas partidas): se suman cantidades
                    for k in ("contrato", "pom", "anterior", "acumulado", "edp"):
                        setattr(it, k, getattr(it, k) + f[k])
            if not p.items:
                raise ValueError("ningún ítem del itemizado está controlado por esta planilla (sin «Planilla de control» "
                                 "ni avance).")
            if con_detalle:
                usadas = [h for h in wb.sheetnames if h.strip() in {it.hoja_control for it in p.items.values()}]
                for h in usadas:
                    try:
                        ws = wb[h]
                        filas_d, aviso = _leer_hoja_detalle(h.strip(), _filas(ws, max_col=110, max_row=5000))
                    except Exception as e:                   # una hoja rara no debe botar toda la planilla
                        filas_d, aviso = [], f"no se pudo leer ({e})"
                    p.detalle += filas_d
                    p.hojas_detalle[h.strip()] = aviso
            p.hallazgos = revisar(p)
            return p
        finally:
            wb.close()
    
    
    # ------------------------------------------------------------------ revisión
    
    def revisar(p: PlanillaCub) -> list[Hallazgo]:
        """Errores típicos de una planilla de avance."""
        H = []
        a = lambda disc, tipo, item, det, grav="Revisar": H.append(Hallazgo(disc, tipo, item, det, p.archivo, grav))
        for it in p.items.values():
            d = it.disciplina
            if it.pom_valida > _EPS and it.acumulado > it.pom_valida * 1.01 + 1e-3:
                a(d, "Avance mayor que la cubicación", it.item,
                  f"Acumulado {_f(it.acumulado)} {it.unidad} > POM {_f(it.pom_valida)} "
                  f"({it.acumulado / it.pom_valida:.0%}).", "Error")
            if it.pom <= _EPS and it.acumulado > _EPS:
                a(d, "Avance sin cubicación", it.item, f"Acumulado {_f(it.acumulado)} {it.unidad} y POM en 0.", "Error")
            if it.acumulado < it.anterior - 1e-3:
                a(d, "Retroceso", it.item, f"Acumulado {_f(it.acumulado)} menor que el anterior {_f(it.anterior)}.",
                  "Error")
            if it.pom_fuera_de_escala:
                a(d, "POM fuera de escala", it.item,
                  f"POM {_f(it.pom)} vs contrato {_f(it.contrato)} {it.unidad}: parece un error de la planilla. Para las HH "
                  f"se usa la cantidad del contrato.", "Error")
            elif it.contrato > _EPS and it.pom > it.contrato * 1.10:
                a(d, "Mayor obra (POM sobre contrato)", it.item,
                  f"POM {_f(it.pom)} vs contrato {_f(it.contrato)} {it.unidad} (+{it.pom / it.contrato - 1:.0%}).", "Info")
            if it.rend <= _EPS and it.pom > _EPS:
                a(d, "Sin rendimiento", it.item, "El ítem tiene cubicación pero rendimiento 0: no suma HH.")
        for disc in p.disciplinas:
            con = [i for i in p.items.values() if i.disciplina == disc and i.acumulado > _EPS]
            sin_ant = [i for i in con if abs(i.anterior) < _EPS]
            if len(con) > 2 and len(sin_ant) >= 0.8 * len(con):
                a(disc, "Planilla sin avance anterior", "",
                  f"{len(sin_ant)} de {len(con)} ítems con avance no traen «AVANCE ANTERIOR»: el avance de la semana se "
                  f"calcula contra la carga anterior guardada en la app.")
        # detalle vs itemizado
        por_item: dict[str, list[FilaDetalle]] = {}
        for f in p.detalle:
            por_item.setdefault(f.item, []).append(f)
        for item, fs in por_item.items():
            it = p.items.get(item)
            hoja = fs[0].hoja
            if it is None:
                a(disciplina_de("", item), "Ítem del detalle fuera del itemizado", item,
                  f"La hoja «{hoja}» usa el ítem {item}, que esta planilla no controla en su itemizado.")
                continue
            acus = [f.acumulado for f in fs if f.acumulado is not None]
            if acus and len(acus) == len(fs) and norm(hoja) == norm(it.hoja_control):
                s = sum(acus)
                if abs(s - it.acumulado) > max(0.01, abs(it.acumulado) * 0.005):
                    a(it.disciplina, "Detalle no cuadra con el itemizado", item,
                      f"«{hoja}» suma {_f(s)} pero el itemizado dice {_f(it.acumulado)} {it.unidad}.", "Error")
            # mismo TAG con la misma descripción y plano: fila duplicada (un equipo con varios componentes es normal)
            tags: dict[tuple, list[int]] = {}
            for f in fs:
                if f.tag and len(f.tag) > 3:
                    tags.setdefault((norm(f.tag), norm(f.descripcion), norm(f.plano)), []).append(f.fila)
            for (tag, _, _), filas in tags.items():
                if len(filas) > 1:
                    a(it.disciplina, "TAG repetido", item,
                      f"TAG {tag} aparece {len(filas)} veces con la misma descripción en «{hoja}» "
                      f"(filas {', '.join(map(str, filas[:6]))}).")
            sin_plano = [f.fila for f in fs if not f.plano]
            if sin_plano and len(sin_plano) < len(fs):
                a(it.disciplina, "Fila sin plano", item,
                  f"{len(sin_plano)} fila(s) de «{hoja}» sin plano de referencia (p. ej. fila {sin_plano[0]}).")
        for it in p.items.values():
            if it.hoja_control and norm(it.hoja_control) not in {norm(h) for h in p.hojas_detalle} and it.acumulado > _EPS:
                a(it.disciplina, "Sin hoja de control", it.item,
                  f"Tiene avance pero su planilla de control «{it.hoja_control}» no está en el archivo.")
        return H
    
    
    def cruzar(planillas: list[PlanillaCub]) -> list[Hallazgo]:
        """Ítems con avance en más de una planilla (doble conteo)."""
        vistos: dict[str, list[tuple[str, float, str]]] = {}
        for p in planillas:
            for it in p.items.values():
                if abs(it.acumulado) > _EPS:
                    vistos.setdefault(it.item, []).append((p.archivo, it.acumulado, it.disciplina))
        H = []
        for item, v in vistos.items():
            if len({x[0] for x in v}) > 1:
                H.append(Hallazgo(v[0][2], "Ítem en dos planillas", item,
                                  "Con avance en: " + "; ".join(f"{a} ({_f(c)})" for a, c, _ in v)
                                  + ". Se usa la primera.", "", "Error"))
        return H
    
    
    def _f(x: float) -> str:
        s = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return s[:-3] if s.endswith(",00") else s
    
    
    # ------------------------------------------------------------------ consolidado
    
    def consolidar(planillas: list[PlanillaCub]) -> dict[str, ItemCub]:
        """Un ítem por código (si viene en dos planillas manda la primera cargada)."""
        out: dict[str, ItemCub] = {}
        for p in planillas:
            for k, it in p.items.items():
                out.setdefault(k, it)
        return out
    
    
    def resumen_disciplinas(items: dict[str, ItemCub]) -> list[dict]:
        filas = []
        for d in DISCIPLINAS + [OTRAS]:
            its = [i for i in items.values() if i.disciplina == d]
            if not its and d == OTRAS:
                continue
            hh_pom = sum(i.hh_pom for i in its)
            hh_gan = sum(i.hh_ganadas for i in its)
            hh_ant = sum(i.anterior * i.rend for i in its)
            filas.append({"disciplina": d, "items": len(its), "hh_contrato": sum(i.contrato * i.rend for i in its),
                          "hh_pom": hh_pom, "hh_ganadas": hh_gan, "hh_semana": hh_gan - hh_ant,
                          "avance": hh_gan / hh_pom if hh_pom > _EPS else 0.0,
                          "avance_anterior": hh_ant / hh_pom if hh_pom > _EPS else 0.0,
                          "con_avance_semana": sum(1 for i in its if i.semana > 1e-6)})
        return filas
    
    
    def excel_resumen(planillas: list[PlanillaCub], fecha: str = "") -> bytes:
        """Consolidado descargable: resumen por disciplina, ítems, detalle y revisión."""
        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
        items = consolidar(planillas)
        wb = openpyxl.Workbook()
        F, FB = Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True, color="FFFFFF")
        FILL = PatternFill("solid", fgColor="06205C")
    
        def hoja(ws, titulo, cols, filas, anchos, fmt=None):
            ws["A1"] = titulo
            ws["A1"].font = Font(name="Arial", size=13, bold=True, color="06205C")
            for j, c in enumerate(cols, 1):
                x = ws.cell(3, j, c)
                x.font, x.fill = FB, FILL
                x.alignment = Alignment(wrap_text=True, vertical="center")
            for i, r in enumerate(filas, 4):
                for j, v in enumerate(r, 1):
                    x = ws.cell(i, j, v)
                    x.font = F
                    if fmt and j in fmt:
                        x.number_format = fmt[j]
            for j, w in enumerate(anchos, 1):
                ws.column_dimensions[get_column_letter(j)].width = w
            ws.freeze_panes = "A4"
            return 4 + len(filas)
    
        ws = wb.active
        ws.title = "Resumen"
        res = resumen_disciplinas(items)
        n = len(res)
        filas = []
        for k, r in enumerate(res):
            f = 4 + k
            filas.append([r["disciplina"], r["items"], r["hh_contrato"], r["hh_pom"], r["hh_ganadas"], r["hh_semana"],
                          f"=IF(D{f}=0,0,E{f}/D{f})"])
        fin = hoja(ws, f"Avance por disciplina{' al ' + fecha if fecha else ''}",
                   ["Disciplina", "Ítems", "HH contrato", "HH POM", "HH ganadas acum.", "HH ganadas semana", "% avance (HH)"],
                   filas, [18, 8, 14, 14, 16, 16, 13], {3: "#,##0", 4: "#,##0", 5: "#,##0", 6: "#,##0", 7: "0.0%"})
        ws.cell(fin, 1, "TOTAL").font = Font(name="Arial", size=10, bold=True)
        for j, L in ((2, "B"), (3, "C"), (4, "D"), (5, "E"), (6, "F")):
            ws.cell(fin, j, f"=SUM({L}4:{L}{fin - 1})").number_format = "#,##0"
        ws.cell(fin, 7, f"=IF(D{fin}=0,0,E{fin}/D{fin})").number_format = "0.0%"
        ws.cell(fin + 2, 1, "HH = cantidad × rendimiento del itemizado EC-002. Fuente: " +
                "; ".join(p.archivo for p in planillas)).font = Font(name="Arial", size=9, italic=True)
        _ = n
    
        ws = wb.create_sheet("Ítems")
        filas = []
        for k, i in enumerate(sorted(items.values(), key=lambda x: (DISCIPLINAS.index(x.disciplina)
                                                                    if x.disciplina in DISCIPLINAS else 9, x.item))):
            f = 4 + k
            filas.append([i.disciplina, i.item, i.descripcion, i.unidad, i.contrato, i.pom, i.anterior, i.acumulado,
                          f"=H{f}-G{f}", f"=IF(F{f}=0,0,H{f}/F{f})", i.rend, f"=F{f}*K{f}", f"=H{f}*K{f}", i.hoja_control,
                          i.planilla])
        hoja(ws, "Ítems por disciplina", ["Disciplina", "Ítem", "Descripción", "Unidad", "Cantidad contrato", "POM",
                                          "Acum. anterior", "Acum. actual", "Semana", "% avance", "Rend. HH/u", "HH POM",
                                          "HH ganadas", "Planilla de control", "Archivo"],
             filas, [14, 17, 46, 7, 12, 12, 12, 12, 11, 9, 9, 11, 11, 24, 30],
             {5: "#,##0.00", 6: "#,##0.00", 7: "#,##0.00", 8: "#,##0.00", 9: "#,##0.00", 10: "0.0%", 11: "0.000",
              12: "#,##0", 13: "#,##0"})
    
        ws = wb.create_sheet("Detalle")
        filas = [[p.items[f.item].disciplina if f.item in p.items else "", f.hoja, f.fila, f.item, f.plano, f.rev, f.tag,
                  f.descripcion, f.unidad, f.cantidad, f.acumulado] for p in planillas for f in p.detalle]
        hoja(ws, "Detalle de las hojas de control", ["Disciplina", "Hoja", "Fila", "Ítem", "Plano", "Rev", "TAG",
                                                      "Descripción / elemento", "Unidad", "Cubicación", "Acumulado"],
             filas, [14, 24, 6, 17, 30, 5, 18, 46, 7, 12, 12], {10: "#,##0.00", 11: "#,##0.00"})
    
        ws = wb.create_sheet("Revisión")
        hall = [h for p in planillas for h in p.hallazgos] + cruzar(planillas)
        hoja(ws, "Revisión de las planillas", ["Gravedad", "Disciplina", "Tipo", "Ítem", "Detalle", "Planilla"],
             [[h.gravedad, h.disciplina, h.tipo, h.item, h.detalle, h.planilla] for h in hall], [9, 14, 30, 17, 80, 34])
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()
    
    
    def excel_tablas(tablas: dict, titulo: str = "") -> bytes:
        """Excel simple con una hoja por tabla (DataFrames), títulos en azul SYNCORE y totales de las columnas numéricas."""
        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        F, FB = Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True, color="FFFFFF")
        for nombre, df in tablas.items():
            ws = wb.create_sheet(str(nombre)[:31])
            ws["A1"] = f"{titulo} · {nombre}" if titulo else str(nombre)
            ws["A1"].font = Font(name="Arial", size=13, bold=True, color="06205C")
            cols = list(df.columns)
            for j, c in enumerate(cols, 1):
                x = ws.cell(3, j, str(c))
                x.font, x.fill = FB, PatternFill("solid", fgColor="06205C")
                x.alignment = Alignment(wrap_text=True, vertical="center")
            n = len(df)
            for i, fila in enumerate(df.itertuples(index=False), 4):
                for j, v in enumerate(fila, 1):
                    if v is not None and isinstance(v, float) and v != v:
                        v = None
                    x = ws.cell(i, j, v.item() if hasattr(v, "item") else v)
                    x.font = F
                    if isinstance(v, float):
                        x.number_format = "#,##0.00"
            if n:
                ws.cell(4 + n, 1, "TOTAL").font = Font(name="Arial", size=10, bold=True)
                for j, c in enumerate(cols, 1):
                    if j > 1 and str(df[c].dtype).startswith("float") and str(c).lower() not in ("fila", "rev", "% avance"):
                        L = get_column_letter(j)
                        x = ws.cell(4 + n, j, f"=SUM({L}4:{L}{3 + n})")
                        x.font, x.number_format = Font(name="Arial", size=10, bold=True), "#,##0.00"
            for j, c in enumerate(cols, 1):
                largo = max([len(str(c))] + [len(str(v)) for v in df[c].head(200)]) if n else len(str(c))
                ws.column_dimensions[get_column_letter(j)].width = min(60, max(8, largo + 2))
            ws.freeze_panes = "A4"
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()
    
    
    def fecha_texto(d: date | datetime | str) -> str:
        return d.strftime("%d-%m-%Y") if hasattr(d, "strftime") else str(d)

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


cubicacion = _modulo_cubicacion()

# ======================================================================
# MÓDULO: cubicar_plano
# ======================================================================
def _modulo_cubicar_plano():
    import io
    import re
    from dataclasses import dataclass
    
    CUB = cubicacion
    
    norm = CUB.norm
    _num = CUB._num
    
    # ------------------------------------------------------------------ plantillas por disciplina
    # columnas: (nombre, tipo, ayuda). Tipos: "texto", "num", "tipo" (lista TIPOS[disciplina])
    COMUNES_INI = [("Lámina", "texto", "Lámina u hoja del plano (p. ej. 1 DE 5)"),
                   ("Tipo", "tipo", "Qué se está cubicando: define cómo se calcula la cantidad")]
    COMUNES_FIN = [("Ítem EC-002", "item", "Opcional: si lo dejas en blanco, la app propone uno"),
                   ("Observación", "texto", "")]
    
    PLANTILLAS: dict[str, list[tuple[str, str, str]]] = {
        "Piping": COMUNES_INI + [
            ("NPS", "texto", 'Diámetro nominal, p. ej. 24" o 2"'),
            ("Descripción", "texto", "Como aparece en el BOM (PIPE .500 WALL…, BALL 300# RF…, S608CB6…)"),
            ("ID / Spool", "texto", "ID del corte, TAG de válvula o spool"),
            ("Cantidad", "num", "N° de piezas (en cañería, deja 1 por corte)"),
            ("Largo (mm)", "num", "Largo del corte de cañería en mm (tabla CUT PIPE LENGTHS)"),
            ("kg c/u", "num", "Peso por soporte en kg (si el ítem se paga por kg)"),
        ] + COMUNES_FIN,
        "Obras civiles": COMUNES_INI + [
            ("Elemento", "texto", "Fundación, pedestal, losa, zanja…"),
            ("Cantidad", "num", "N° de elementos iguales"),
            ("Largo (m)", "num", ""), ("Ancho (m)", "num", ""), ("Alto (m)", "num", "Alto o espesor"),
            ("Descuento (m³)", "num", "Volumen a descontar (huecos, pedestales…)"),
            ("kg c/u", "num", "Para armadura, pernos o insertos pagados por kg"),
        ] + COMUNES_FIN,
        "Estructura": COMUNES_INI + [
            ("Marca", "texto", "Marca o pieza del plano"),
            ("Perfil", "texto", "HEB 300, L 50x50x5, grating…"),
            ("Cantidad", "num", "N° de piezas"),
            ("Largo (m)", "num", "Largo de cada pieza"),
            ("kg/m", "num", "Peso lineal del perfil"),
            ("kg c/u", "num", "Peso por pieza (si lo trae la lista de materiales, manda sobre largo × kg/m)"),
        ] + COMUNES_FIN,
        "Mecánica": COMUNES_INI + [
            ("TAG", "texto", "TAG del equipo"),
            ("Equipo", "texto", "Descripción del equipo o componente"),
            ("Cantidad", "num", "N° de unidades"),
            ("Peso (ton)", "num", "Peso por unidad, informativo"),
        ] + COMUNES_FIN,
        "Eléctricos": COMUNES_INI + [
            ("TAG / circuito", "texto", "TAG del cable, tramo o equipo"),
            ("Descripción", "texto", "Sección, tipo de cable, tamaño de escalerilla…"),
            ("Cantidad", "num", "N° de tramos o unidades"),
            ("Largo (m)", "num", "Largo de cada tramo (cables, escalerillas, conduit, malla)"),
        ] + COMUNES_FIN,
    }
    
    # tipo → cómo se calcula ("m_mm" metros desde mm, "m", "m2", "m3", "kg", "un")
    TIPOS: dict[str, dict[str, str]] = {
        "Piping": {"Cañería": "m_mm", "Válvula": "un", "Soporte": "kg_o_un", "Fitting": "un", "Flange": "un",
                   "Olet": "un", "Perno / empaquetadura": "un", "Aislación": "m_mm", "Otro": "un"},
        "Obras civiles": {"Excavación": "m3", "Relleno": "m3", "Emplantillado": "m3", "Hormigón": "m3",
                          "Mortero / grout": "m3", "Moldaje": "m2", "Geomembrana": "m2", "Armadura": "kg",
                          "Pernos / insertos (kg)": "kg", "Pernos / insertos (un)": "un", "Prefabricado": "un", "Otro": "un"},
        "Estructura": {"Perfil / pieza": "kg", "Grating / piso": "kg", "Baranda / escalera": "kg", "Otro (un)": "un"},
        "Mecánica": {"Equipo": "un", "Componente": "un", "Otro": "un"},
        "Eléctricos": {"Cable": "m", "Escalerilla / bandeja": "m", "Conduit": "m", "Banco de ductos": "m",
                       "Malla a tierra": "m", "Soldadura exotérmica": "un", "Luminaria": "un", "Tablero / equipo": "un",
                       "Soporte": "un", "Otro": "un"},
    }
    # materiales que van incluidos en el metro de cañería: se listan, pero no tienen ítem de HH propio
    SIN_ITEM = {"Fitting", "Flange", "Olet", "Perno / empaquetadura"}
    UNIDAD_TXT = {"m_mm": "m", "m": "m", "m2": "m²", "m3": "m³", "kg": "kg", "un": "un"}
    DISCIPLINAS = list(PLANTILLAS)
    
    # palabras clave del tipo para buscar el ítem en el catálogo
    _CLAVES_TIPO = {
        "Cañería": ["PIPE"], "Válvula": ["VALVE"], "Soporte": ["SOPORTE", "SHOE", "GUIA"], "Aislación": ["AISLACION"],
        "Excavación": ["EXCAVACION"], "Relleno": ["RELLENO"], "Emplantillado": ["EMPLANTILLADO"],
        "Hormigón": ["HORMIGON", "PEDESTAL", "FUNDACION", "MURO", "LOSA", "RADIER", "CANALETA"], "Mortero / grout": ["MORTERO", "GROUT"], "Moldaje": ["MOLDAJE"],
        "Geomembrana": ["GEOMEMBRANA"], "Armadura": ["REFUERZO DE HORMIGON", "ARMADURA", "ACERO DE REFUERZO", "FIERRO"],
        "Pernos / insertos (kg)": ["PERNO", "INSERTO"], "Pernos / insertos (un)": ["PERNO", "INSERTO"],
        "Prefabricado": ["PREFABRICAD"], "Grating / piso": ["GRATING", "PISO"], "Baranda / escalera": ["BARANDA",
                                                                                                        "ESCALERA"],
        "Cable": ["CABLE"], "Escalerilla / bandeja": ["ESCALERILLA", "BANDEJA"], "Conduit": ["CONDUIT"],
        "Banco de ductos": ["BANCO DE DUCTOS"], "Malla a tierra": ["DESNUDO", "MALLA"],
        "Soldadura exotérmica": ["EXOTERMICA"], "Luminaria": ["LUMINARIA", "PROYECTOR"], "Tablero / equipo": ["TABLERO"],
    }
    _CAPITULO_TIPO = {"Excavación": ("00",), "Relleno": ("00",), "Geomembrana": ("00",), "Emplantillado": ("10",),
                      "Hormigón": ("10",), "Mortero / grout": ("10",), "Moldaje": ("10",), "Armadura": ("10",),
                      "Pernos / insertos (kg)": ("10",), "Pernos / insertos (un)": ("10",), "Prefabricado": ("10",),
                      "Cañería": ("51",), "Válvula": ("52",), "Soporte": ("52",)}
    _VALVULAS = {"BALL": "BALL", "BUTTERFLY": "BUTTERFLY", "CHECK": "CHECK", "GATE": "GATE", "GLOBE": "GLOBE",
                 "PLUG": "PLUG", "KNIFE": "KNIFE"}
    
    
    def columnas(disciplina: str) -> list[tuple[str, str, str]]:
        return PLANTILLAS[disciplina]
    
    
    def fila_vacia(disciplina: str) -> dict:
        return {n: (None if t == "num" else "") for n, t, _ in PLANTILLAS[disciplina]}
    
    
    # ------------------------------------------------------------------ catálogo EC-002
    
    def catalogo_de_planilla(datos: bytes) -> list[dict]:
        """Todo el itemizado EC-002 de una planilla de avance (o cualquier Excel con la hoja de itemizado)."""
        wb = CUB._abrir(datos)
        try:
            h = CUB._hoja_itemizado(wb)
            if not h:
                raise ValueError("el archivo no tiene una hoja de itemizado EC-002.")
            filas = CUB._leer_itemizado(CUB._filas(wb[h], max_col=80, max_row=3000))
        finally:
            wb.close()
        vistos, out = set(), []
        for f in filas:
            if f["item"] in vistos:
                continue
            vistos.add(f["item"])
            out.append({"item": f["item"], "descripcion": f["descripcion"], "unidad": f["unidad"], "rend": f["rend"],
                        "contrato": f["contrato"], "disciplina": CUB.disciplina_de(f["especialidad"], f["item"])})
        return out
    
    
    def unir_catalogos(*listas) -> list[dict]:
        out: dict[str, dict] = {}
        for ls in listas:
            for c in ls or []:
                if c.get("item") and (c["item"] not in out or (not out[c["item"]].get("rend") and c.get("rend"))):
                    out[c["item"]] = c
        return sorted(out.values(), key=lambda c: c["item"])
    
    
    _VACIAS = {"PARA", "DE", "DEL", "LA", "LAS", "EL", "LOS", "Y", "CON", "EN", "POR", "A", "AL", "TIPO"}
    
    
    def _tokens(s: str) -> set[str]:
        out = set()
        for t in re.split(r"[^A-Z0-9#/.]+", norm(s)):
            if len(t) < 2 or t in _VACIAS:
                continue
            if len(t) > 4 and t.isalpha():
                t = t[:-2] if t.endswith("ES") and len(t) > 5 else t.rstrip("S")      # plural → singular
            out.add(t)
        return out
    
    
    def _medida(s: str) -> str:
        """Diámetro en pulgadas normalizado: 24" / 24 / 1 1/2" / 1-1/2 → "24", "1 1/2"."""
        s = norm(s).replace("”", '"').replace("''", '"').replace("-", " ")
        m = re.search(r"(\d+(?:\s+\d/\d)?|\d/\d)\s*(?:\"|IN\b|PULG)", s) or re.fullmatch(r"\s*(\d+(?:\s+\d/\d)?|\d/\d)\s*",
                                                                                      s)
        return re.sub(r"\s+", " ", m.group(1)).strip() if m else ""
    
    
    def _medida_desc(desc: str) -> str:
        d = norm(desc).replace("”", '"').replace("½", " 1/2")
        m = re.search(r"[–-]\s*(\d+(?:\s*\d/\d)?|\d/\d)\s*\"\s*$", d) or re.search(r"(\d+(?:\s+\d/\d)?|\d/\d)\s*\"", d)
        return re.sub(r"\s+", " ", m.group(1)).strip() if m else ""
    
    
    def sugerir_item(disciplina: str, fila: dict, catalogo: list[dict], contexto: str = "",
                     unidad: str = "") -> tuple[str, float]:
        """(ítem, puntaje) más parecido del catálogo para la fila. Puntaje 0 = sin sugerencia.
        contexto: descripción del plano (p. ej. «Fundaciones calderas») para elegir entre ítems por ubicación."""
        tipo = fila.get("Tipo") or ""
        texto = " ".join(str(fila.get(k) or "") for k in ("Descripción", "Elemento", "Perfil", "Equipo", "TAG / circuito",
                                                              "TAG", "Marca"))
        tok_ctx = _tokens(contexto) - _tokens(texto)
        tok = _tokens(texto)
        medida = _medida(str(fila.get("NPS") or "")) or _medida(texto)
        claves = _CLAVES_TIPO.get(tipo, [])
        candidatos = [c for c in catalogo if c["disciplina"] == disciplina or
                      (disciplina == "Piping" and c["item"][:2] in ("51", "52", "53"))]
        mejor, puntaje = "", 0.0
        for c in candidatos:
            d = norm(c["descripcion"])
            p = 0.0
            if unidad and c.get("unidad") and not _unidad_compatible(unidad, c["unidad"]):
                continue
            if tipo in _CAPITULO_TIPO and c["item"][:2] not in _CAPITULO_TIPO[tipo]:
                continue
            if claves:
                if not any(k in d for k in claves):
                    continue
                p += 3
            if tipo == "Válvula":
                v = next((k for k in _VALVULAS if k in norm(texto)), None)
                if v:
                    p += 3 if v in d else -3
            if medida:
                md = _medida_desc(c["descripcion"])
                if md:
                    p += 4 if md == medida else -4
            td = _tokens(d)
            p += min(len(tok & td), 6) * 0.7 + min(len(tok_ctx & td), 4) * 0.8
            if c.get("rend"):
                p += 0.2
            if p > puntaje:
                mejor, puntaje = c["item"], p
        return mejor, puntaje
    
    
    # ------------------------------------------------------------------ cálculo
    
    @dataclass
    class FilaCalc:
        n: int
        lamina: str
        tipo: str
        descripcion: str
        unidad: str
        cantidad: float
        item: str
        item_sugerido: bool
        item_desc: str
        unidad_item: str
        rend: float
        hh: float
        aviso: str
    
    
    def _desc_fila(disciplina: str, f: dict) -> str:
        claves = {"Piping": ("NPS", "Descripción", "ID / Spool"), "Obras civiles": ("Elemento",),
                  "Estructura": ("Marca", "Perfil"), "Mecánica": ("TAG", "Equipo"),
                  "Eléctricos": ("TAG / circuito", "Descripción")}[disciplina]
        return " · ".join(str(f.get(k)).strip() for k in claves if str(f.get(k) or "").strip())
    
    
    def _unidad_compatible(calc: str, unidad_item: str) -> bool:
        u = norm(unidad_item).replace("3", "3").replace("²", "2")
        if not u:
            return True
        grupos = {"m": {"M", "ML", "MTS", "MT"}, "m²": {"M2"}, "m³": {"M3"}, "kg": {"KG", "KGS"},
                  "un": {"UN", "UND", "EA", "C/U", "CU", "U", "UNID", "GL", "LS", "PZA"}}
        u = u.replace("²", "2").replace("³", "3")
        return u in grupos.get(calc, set())
    
    
    def cantidad_fila(disciplina: str, f: dict) -> tuple[float, str, str]:
        """(cantidad, unidad, aviso) de una fila según su tipo."""
        tipo = f.get("Tipo") or ""
        modo = TIPOS[disciplina].get(tipo)
        if not modo:
            return 0.0, "", "Elige el tipo"
        cant = _num(f.get("Cantidad"))
        if cant == 0 and modo in ("m_mm", "m", "m3", "m2", "kg") and disciplina != "Mecánica":
            cant = 1.0                                    # en largos y volúmenes, la cantidad vacía es 1
        L, A, H = _num(f.get("Largo (m)")), _num(f.get("Ancho (m)")), _num(f.get("Alto (m)"))
        if modo == "m_mm":
            mm = _num(f.get("Largo (mm)"))
            return cant * mm / 1000, "m", ("" if mm else "Falta el largo (mm)")
        if modo == "m":
            return cant * L, "m", ("" if L else "Falta el largo (m)")
        if modo == "m2":
            return cant * L * A, "m²", ("" if L and A else "Faltan largo y ancho")
        if modo == "m3":
            v = cant * L * A * H - _num(f.get("Descuento (m³)"))
            return v, "m³", ("" if L and A and H else "Faltan largo, ancho o alto")
        if modo == "kg":
            kgcu = _num(f.get("kg c/u"))
            if kgcu:
                return cant * kgcu, "kg", ""
            kgm = _num(f.get("kg/m"))
            if kgm and L:
                return cant * L * kgm, "kg", ""
            return 0.0, "kg", "Falta el peso (kg c/u, o largo × kg/m)"
        if modo == "kg_o_un":
            kgcu = _num(f.get("kg c/u"))
            return (cant * kgcu, "kg", "") if kgcu else (cant, "un", "")
        return cant, "un", ("" if cant else "Falta la cantidad")
    
    
    def _vacia(f: dict) -> bool:
        return all(v in (None, "") or (isinstance(v, float) and v != v) for k, v in f.items() if k != "Tipo")
    
    
    def calcular(disciplina: str, filas: list[dict], catalogo: list[dict], contexto: str = "") -> list[FilaCalc]:
        por_item = {c["item"]: c for c in catalogo}
        out = []
        for n, f in enumerate(filas, 1):
            if _vacia(f):
                continue
            cant, unidad, aviso = cantidad_fila(disciplina, f)
            item = str(f.get("Ítem EC-002") or "").split(" · ")[0].strip()
            sugerido = False
            if not item and catalogo and f.get("Tipo") not in SIN_ITEM:
                item, p = sugerir_item(disciplina, f, catalogo, contexto, unidad)
                sugerido = bool(item)
                if p < 5 and item:
                    aviso = (aviso + "; " if aviso else "") + "ítem sugerido con poca certeza: revísalo"
            c = por_item.get(item, {})
            if item and not c and catalogo:
                aviso = (aviso + "; " if aviso else "") + "ítem no está en el catálogo"
            rend = float(c.get("rend") or 0)
            if c and not _unidad_compatible(unidad, c.get("unidad", "")):
                aviso = (aviso + "; " if aviso else "") + f"el ítem se paga en {c.get('unidad')} y la fila está en {unidad}"
            out.append(FilaCalc(n, str(f.get("Lámina") or ""), str(f.get("Tipo") or ""), _desc_fila(disciplina, f), unidad,
                                cant, item, sugerido, c.get("descripcion", ""), c.get("unidad", ""), rend, cant * rend,
                                aviso))
        return out
    
    
    def resumen_items(calc: list[FilaCalc]) -> list[dict]:
        r: dict[tuple, dict] = {}
        for f in calc:
            k = (f.item or "(sin ítem)", f.unidad)
            d = r.setdefault(k, {"Ítem": k[0], "Descripción ítem": f.item_desc, "Unidad": f.unidad, "Cantidad": 0.0,
                                 "Rend. HH/u": f.rend, "HH": 0.0, "Filas": 0})
            d["Cantidad"] += f.cantidad
            d["HH"] += f.hh
            d["Filas"] += 1
        return sorted(r.values(), key=lambda d: d["Ítem"])
    
    
    def comparar(nuevo: list[dict], anterior: list[dict]) -> list[dict]:
        """Por ítem: cantidad de la cubicación anterior vs la nueva."""
        a = {(d["Ítem"], d["Unidad"]): d for d in anterior}
        b = {(d["Ítem"], d["Unidad"]): d for d in nuevo}
        out = []
        for k in sorted(set(a) | set(b)):
            ca, cb = a.get(k, {}).get("Cantidad", 0.0), b.get(k, {}).get("Cantidad", 0.0)
            rend = (b.get(k) or a.get(k))["Rend. HH/u"]
            out.append({"Ítem": k[0], "Unidad": k[1], "Anterior": ca, "Nueva": cb, "Diferencia": cb - ca,
                        "Diferencia HH": (cb - ca) * rend})
        return out
    
    
    # ------------------------------------------------------------------ datos del plano
    
    _RE_PLANO = re.compile(r"[A-Z0-9]{3,}(?:-[A-Z0-9]{1,6}){3,}")
    
    
    def meta_de_nombre(nombre: str) -> dict:
        base = re.sub(r"\.pdf$", "", nombre or "", flags=re.I)
        b = norm(base).replace(" ", "")
        m = _RE_PLANO.search(b)
        plano = m.group(0) if m else norm(base)[:60]
        plano = re.sub(r"[-_]?REV.*$", "", plano)
        rev = re.search(r"REV\.?\s*([A-Z0-9]{1,3})", norm(base))
        area = re.search(r"-(\d{4})-", plano)
        disc = ""
        for clave, d in (("-PIP-", "Piping"), ("-CIV-", "Obras civiles"), ("-STR-", "Estructura"), ("-EST-", "Estructura"),
                         ("-MEC-", "Mecánica"), ("-ELE-", "Eléctricos"), ("-ELC-", "Eléctricos"), ("-EL-", "Eléctricos"),
                         ("-INS-", "Eléctricos")):
            if clave in plano + "-":
                disc = d
                break
        return {"plano": plano, "rev": rev.group(1) if rev else "", "area": area.group(1) if area else "",
                "disciplina": disc}
    
    
    def paginas_pdf(datos: bytes) -> int:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(datos)) as pdf:
            return len(pdf.pages)
    
    
    def imagen_pagina(datos: bytes, pagina: int, dpi: int = 110) -> bytes:
        """PNG de una página del plano (para verlo al lado de la planilla)."""
        import pdfplumber
        with pdfplumber.open(io.BytesIO(datos)) as pdf:
            im = pdf.pages[pagina].to_image(resolution=dpi).original
            buf = io.BytesIO()
            im.save(buf, format="PNG")
            return buf.getvalue()
    
    
    # ------------------------------------------------------------------ Excel
    
    def excel(meta: dict, disciplina: str, filas: list[dict], calc: list[FilaCalc],
              comparacion: list[dict] | None = None) -> bytes:
        """Cubicación del plano: resumen por ítem (con fórmulas de HH), filas calculadas y lo digitado."""
        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
        F, FB = Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True)
        FH, FILL = Font(name="Arial", size=10, bold=True, color="FFFFFF"), PatternFill("solid", fgColor="06205C")
        AZUL = Font(name="Arial", size=10, color="0000FF")
        wb = openpyxl.Workbook()
    
        def cab(ws, fila, cols, anchos):
            for j, c in enumerate(cols, 1):
                x = ws.cell(fila, j, c)
                x.font, x.fill = FH, FILL
                x.alignment = Alignment(wrap_text=True, vertical="center")
            for j, w in enumerate(anchos, 1):
                ws.column_dimensions[get_column_letter(j)].width = w
    
        # Resumen
        ws = wb.active
        ws.title = "Resumen"
        ws["A1"] = f"Cubicación {disciplina} · {meta.get('plano', '')} Rev. {meta.get('rev', '')}"
        ws["A1"].font = Font(name="Arial", size=13, bold=True, color="06205C")
        datos_meta = [("N° plano", meta.get("plano", "")), ("Revisión", meta.get("rev", "")), ("Área", meta.get("area", "")),
                      ("Descripción / línea", meta.get("descripcion", "")), ("Disciplina", disciplina),
                      ("Cubicó", meta.get("usuario", "")), ("Fecha", meta.get("fecha", ""))]
        for i, (k, v) in enumerate(datos_meta, 3):
            ws.cell(i, 1, k).font = FB
            ws.cell(i, 2, v).font = F
        f0 = 3 + len(datos_meta) + 1
        cab(ws, f0, ["Ítem EC-002", "Descripción ítem", "Unidad", "Cantidad", "Rend. HH/u", "HH", "Filas"],
            [18, 50, 8, 13, 11, 12, 7])
        res = resumen_items(calc)
        for i, d in enumerate(res, f0 + 1):
            ws.cell(i, 1, d["Ítem"]).font = F
            ws.cell(i, 2, d["Descripción ítem"]).font = F
            ws.cell(i, 3, d["Unidad"]).font = F
            ws.cell(i, 4, f"=SUMIFS(Filas!$F:$F,Filas!$G:$G,A{i},Filas!$E:$E,C{i})").number_format = "#,##0.000"
            x = ws.cell(i, 5, d["Rend. HH/u"])
            x.number_format, x.font = "0.0000", AZUL
            ws.cell(i, 6, f"=D{i}*E{i}").number_format = "#,##0.0"
            ws.cell(i, 7, f"=COUNTIFS(Filas!$G:$G,A{i},Filas!$E:$E,C{i})")
        fin = f0 + 1 + len(res)
        ws.cell(fin, 1, "TOTAL HH").font = FB
        x = ws.cell(fin, 6, f"=SUM(F{f0 + 1}:F{fin - 1})")
        x.font, x.number_format = FB, "#,##0.0"
        ws.cell(fin + 2, 1, "Rendimiento (azul): itemizado EC-002 del contrato. Cantidad y HH se recalculan desde la hoja "
                            "«Filas».").font = Font(name="Arial", size=9, italic=True)
    
        # Filas calculadas
        wf = wb.create_sheet("Filas")
        cab(wf, 1, ["N°", "Lámina", "Tipo", "Descripción", "Unidad", "Cantidad", "Ítem EC-002", "Ítem sugerido",
                    "Rend. HH/u", "HH", "Aviso"], [5, 10, 18, 50, 7, 12, 18, 10, 10, 10, 45])
        for i, c in enumerate(calc, 2):
            for j, v in enumerate([c.n, c.lamina, c.tipo, c.descripcion, c.unidad, round(c.cantidad, 6), c.item,
                                   "sí" if c.item_sugerido else "", c.rend, f"=F{i}*I{i}", c.aviso], 1):
                x = wf.cell(i, j, v)
                x.font = F
            wf.cell(i, 6).number_format = "#,##0.000"
            wf.cell(i, 10).number_format = "#,##0.0"
        wf.freeze_panes = "A2"
    
        # Lo digitado
        wd = wb.create_sheet("Digitado")
        cols = [n for n, _, _ in PLANTILLAS[disciplina]]
        cab(wd, 1, cols, [max(9, min(40, len(c) + 4)) for c in cols])
        for i, f in enumerate([f for f in filas if not _vacia(f)], 2):
            for j, c in enumerate(cols, 1):
                v = f.get(c)
                if isinstance(v, float) and v != v:
                    v = None
                wd.cell(i, j, v).font = F
        wd.freeze_panes = "A2"
    
        if comparacion:
            wc = wb.create_sheet("Comparación")
            cab(wc, 1, ["Ítem", "Unidad", "Anterior", "Nueva", "Diferencia", "Diferencia HH"], [18, 8, 12, 12, 12, 13])
            for i, d in enumerate(comparacion, 2):
                for j, k in enumerate(["Ítem", "Unidad", "Anterior", "Nueva"], 1):
                    wc.cell(i, j, d[k]).font = F
                wc.cell(i, 5, f"=D{i}-C{i}").number_format = "#,##0.000;[Red]-#,##0.000"
                wc.cell(i, 6, round(d["Diferencia HH"], 2)).number_format = "#,##0.0;[Red]-#,##0.0"
                for j in (3, 4):
                    wc.cell(i, j).number_format = "#,##0.000"
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


cubicar_plano = _modulo_cubicar_plano()

# ======================================================================
# MÓDULO: lector_planos
# ======================================================================
def _modulo_lector_planos():
    import base64
    import io
    import json
    import re
    import zlib
    from dataclasses import dataclass, field
    
    import numpy as np
    
    _PLANTILLAS_B64 = "eNrtmkmvZTsNhf9LMXggnVvs9MmQHkQj+kalGiIYMeL/i7PX59g59xXjYsDgbvnsJnEce3nZuZ8+ffjhh0d6fLo+ljoe18c62vPaSn1e+7pv5Pu6en5eU+6P9PF6SvffTNWkdmWT0v1yej54e/6+5v19uq9Xrhr2ecnzfqek7LNdSJ8/Pz59+IVpc8UImm+WPYN/kvswqbT9Xt0PW90q9bI/Hbm56vvTZVM8l6X5f/Xl+ZN/mn2ustKezJ92fzr86dLar+eTZDZKJpXnk2R6T1fhMhXy0/KhRL+nSD2/Tn4VrVDyuN+bbck0ewosvPL9TLqN1O9P7i+qbt6WyLVLCSl6X2JuqfTTD4+MVY7t72HXvg2hvWVBe/nF3SSNbXXNcc1jffk5hv8oz6H9R3vdCXMZLcjk1kKWnyCnlEIuK+TeQ17V5ZzDJXMMmUfERbna4au/dl8Jq/TDUW5BgZS0TUkuWbWFWeaZOcl7l8u1T3+nje7far80oqS+xe0zP3i/QbIJbo4dCANbeyyXIGKNGK8QL8hz+ju11v1pXctHbMdM7nXMKFWTvjB5upjlFSbXFfKM++V4vwg7TJ5xv2oVJsuXTV52X+b5o5snfKme0bVG/Ej1cMXbQcJJ6+GX+RmB4bHHW7d6AX/lGLnOw5nb+U0+PT732AFkQA/ZdkwywIdc4xXAD5ldQgYAkdkx5OXTbn/6yYdH+X/AHwH/OGzxtLRJjuSp7MyYZjYp57GlUTbip50FSnfsv5ZngT1y8534/Regd0xPc5EC9y2s3w/r5nl4ah3Fnu80nzFPliFltpza++QcboY8h/Je9lQuTUoBBthTQKArJLpSjwbQLg2itAg8Lu1VV/oaFie3V4oz6OFq2fTddvmDgS9ruO3/nEYp8Lq6hcybOZ/mlZaXJizpcIWaXEEmEYfpI+mGYHpomfcbS98kabnm9vPZunT67RcSwpg7ZTePn+I0JeVhfvS2g2dLINi1go8VGaQWcre2TpZuzmaekhT58ZcUqR6qa0u5YrC8J82t+Px6P8lgTRSmjOpyzxhPm2bZKflse5N+Z86b857xEnqlvpzHoIGShEx8Tdm854AkuZ3ML3M1Od9gnzCGUG8sX1JVurjqfJyTQ3GnJivKtLOw+yt8T7PPsPRKjm6rACTZGXFdws/JNsg7ExumCNZgjajWwu+c/pRXkXn+5rFtfxAwrSZdGIqrp3/yZjFQh2APB6p2BWj1q7o8RCKQp8g68rqSyQZK13ZMYwG1eCRfs3vuTTJpTQRacXaQEwZC3eW+VjRn0+7h/NAGRmsKEazd00GwfvMlN06OSl4SEE5CmiQYy+6thUpGCQtyNQXvRNRsEV0tbufZjgQtVX604UZ+mAspQr6rOyW1d9oc/rcRoLp5m8Ka52NM/0zwco+1vX56AcHMo+M6yU1zA1kzENmFCg5zQfm0b5esgUyRxnaClZMMkFxuWiBybzHM6PEtQXLLO9y7KVUbK6qW+8y7BwrKx4XqZbyyZsJTmuYWpZF2qrs5uhIIqaYC7XKvGiVfX9Ln4zaSdnUpoi7txmJCBTV+pAhZ2E0Rtdhd3V7d8v6O9UsvLtYGNJGKhQpLcc/uWvW3NZBedZfc0ytBgd4QsCG3UVwmuU3oNDE0raSElO5rwigAXhEDTiQsrGJUa8u41LZWfrEWJSHVB6vKM3IDDpc9TUHbK0V9hx5rbCkhew6oopxiMXjx5ZeN82eK/2YrNERmyPG3Ve8f2ifuy58vmg6pbdy6f8CTOl/IcnssCCk/LiHVpkwEMmw207tYnrguvE3qtRdvV/YsnuaUIyCqXC0OYJUk7wlAQX6WFQvefFHMC6MqBcb2vSSt1ujRh5E+5TCX6TOqM7jNzi1ianPmBl8nBxoP0juNpkGhchDI9+q8bG9YEqtcntC3NmcXY6f+7FUweadXjzwDxEIpDOAQbqKiCl2ik2bTOJE7o7R7HakEImHzS6e/HKB5mWdsNfCuoTXnq0ZVREMMe+G2pCs4JDL8Q09bc9LQBahDWa+v7ME3IIJS6a/vKQAsje7OGj5F0uA4EcQVtYn7iovpq62qEyqSLGrgYTPNg2WvCLvH/4oiss207bIM5jCa5Z4Ub1Biqpma3V+pbgNWgEr426VveiOTQV6mJ5pS6cMkR3mbXiq9HR4Eel5nr0ZZxeQ5XD5u72KpB58ePUIfeQU4t2P4fvCATqAOCo+fe7n+duygdyCT1yCRjS34fM8u37ErqLhv1+WbdflWnazMasUaRTc+Y0V3/S8FeD0qyhXFe4lX8BnkWmJIfObsVodsNuA2VrAqcebgeLKEFVaALcOcwFKPZniJihmbGIVM0+Vvu/CfXtrJexOit5G6M6DqCWcKsuhaQDEKtGJ4TsgxqXkFyqejrH/pMr0jjapBZB1FiE0ornMwv3UyvLVczoZ9t+wcMz3DKG43D530rKW7y7TFkTcd+uYlcRk6iicMsoKmG5OcL/THUUs0tEnPg7TNOCvI2SzRQ5+MU6knFRiKz2W99b11P9uMHzqD70HymndwOZug4VQ8rzkNaFAD/R4wcLnXalAxb/eIJk1KRM0H6Ybz9QRBW1sjzZH8FKA12VdjUhvDYIkdLEnazU6Rqdyp4zQeyDNp7Ho1YGQ8Z+995aPf8Z0Pj7rRcNBfEQKYvIbLIIDJQgBkEMBkLQa5xCuWNaCWJYYEAZDb8f6t4pujtpE6+HJktCTOA88o09sBnK+06yh+SHewc75RlFBmNnX4jB3fZDZ7OW1OQM+/ZJ8Z2dqTyECcZHMzZCBOMmWsyTPGtGYQ8ogx4W4mW4LRmCOFcfJqYZySwjqFkgLkGW4d87aLwqu6dSpxSAkAUkEXAQ5suq0DOj6tIw8aO8W3nakzTQo4ReuOtgU4W14PAdTVKOxyemogDyrUyK0afPlBZe0n2ZEyv/Qs+nWS6EG7erCuCJ80PXpyUK7ioUOfBk/JcVBSZwqCAbuZNIgiM/aD6dQcrcV68BdhRaY4pSxNccBQhDKwErBkyCuqts08U+Tecr7QhebmRc083+XO75+5c4dzDo+1M1lLKGN7bI7unpEtOSxkqx8OYQ6rvIbD3mRrO2zX2Wy2RYXHfnfXNb2EXncP7M2bm1f1dhQtDXz6oi3evB/sZ8Z6KOCkM0snSFnlUgiRhwe5d3hjdOv0vS9Yah0KeX27O3QolCHoLQ6Los3feaG6UpBu06qFVtccrtV16HR2E5L1gpI1WayZcDvA3N2EtayZkDYpT7dc0HqJU3AUdjMKsly9h54MrV4C40H7kXGopyyl/vy+AGxRxHF6ZMWgFRRxSkQEczRE5NIhNfmsbGpQR+t+Qi+PuagupNLjhfwswp8akkXKMxcnaZTBJHFFE+jUKW5WNKPAKprNYE/XIZO1JW77JBodRRan8Szbk9r/fliqW8NwukyEIxthBodb4FAa0d6wKm2+eLzwZBYKBFoa8kfY2ByP3VrZdONfR61sZx49Dq0SjIjT0RwUySpBtmlEzZhIZslr+e7ScGm6tEwKuGb81/8EqRHdqfgZXfWE4DG/PDW08a1/SpnDk0RZL6T0H2YCDGMlUsAbJ+koxtESqQ8oHn7S0sAZzgYgJ3qFkxy+JEzrPpvUQDQ5fVdcZkZTqsb91uN+H/HtoHVIUX6N+AcYeIMOWGnIcDhqSCUaSq7PSVgDl7vfIMD/fbiI7Q9HVyU6siaP8JFLa92jHSYkwSJDhJBJtyX5WanJJb61A/qtoZV/cjra1RLpNIE8tOKTJ00cNCX/fwo6d8kPQVKi1dGd3y9oLTBhDf44Sv7ny39BfAWC8zWD9vPnz/8BfmWJDQ=="
    _UMBRAL = 0.08          # distancia máxima para aceptar una letra
    _ESPACIO = 0.75         # separación (en alturas de letra) que se considera espacio
    
    
    def _plantillas():
        datos = json.loads(zlib.decompress(base64.b64decode(_PLANTILLAS_B64)))
        out: dict[int, list] = {}
        for ch, n, flat in datos:
            out.setdefault(n, []).append((ch, np.array(flat, float).reshape(-1, 2)))
        return out
    
    
    _CACHE: dict = {}
    
    
    # ------------------------------------------------------------------ geometría
    
    class _Pagina:
        def __init__(self, pg):
            self.pg = pg
            tr = [np.array(c["pts"], float) for c in pg.curves if c.get("pts")]
            tr += [np.array([(l["x0"], l["top"]), (l["x1"], l["bottom"])], float) for l in pg.lines]
            self.tr = tr
            self.bb = (np.array([(t[:, 0].min(), t[:, 1].min(), t[:, 0].max(), t[:, 1].max()) for t in tr])
                       if tr else np.zeros((0, 4)))
            self.chars = pg.chars
    
        def trazos(self, a, ya, b, yb, m=0.3):
            bb = self.bb
            if not len(bb):
                return []
            k = (bb[:, 0] >= a + m) & (bb[:, 2] <= b - m) & (bb[:, 1] >= ya + m) & (bb[:, 3] <= yb - m)
            return [self.tr[i] for i in np.where(k)[0]]
    
        def texto_real(self, a, ya, b, yb) -> str | None:
            cs = [c for c in self.chars if c["x0"] >= a and c["x1"] <= b and c["top"] >= ya and c["bottom"] <= yb]
            if not cs:
                return None
            cs.sort(key=lambda c: (round(c["top"] / 3), c["x0"]))
            return re.sub(r"\s+", " ", "".join(c["text"] for c in cs)).strip()
    
    
    def _segmentos(pg):
        segs = [(l["x0"], l["top"], l["x1"], l["bottom"]) for l in pg.lines]
        for r in pg.rects:
            x0, t, x1, b = r["x0"], r["top"], r["x1"], r["bottom"]
            segs += [(x0, t, x1, t), (x0, b, x1, b), (x0, t, x0, b), (x1, t, x1, b)]
        for c in pg.curves:
            pts = c.get("pts") or []
            for (a, b_), (c_, d) in zip(pts, pts[1:]):
                segs.append((min(a, c_), min(b_, d), max(a, c_), max(b_, d)))
        H = [s for s in segs if abs(s[3] - s[1]) < 0.6 and s[2] - s[0] > 20]
        V = [s for s in segs if abs(s[2] - s[0]) < 0.6 and s[3] - s[1] > 6]
        return H, V
    
    
    def _tablas(H, V):
        """Tablas = grupos de líneas horizontales con el mismo ancho; filas entre líneas; celdas por verticales completas."""
        grupos: dict = {}
        for x0, y, x1, _ in H:
            grupos.setdefault((round(x0), round(x1)), []).append(y)
        bloques = []
        for (x0, x1), ys in grupos.items():
            u = []
            for y in sorted(ys):
                if not u or y - u[-1] > 1.0:
                    u.append(y)
            b = [u[0]]
            for y in u[1:]:
                if y - b[-1] > 70:
                    if len(b) >= 3:
                        bloques.append((x0, x1, b))
                    b = [y]
                else:
                    b.append(y)
            if len(b) >= 3:
                bloques.append((x0, x1, b))
        out = []
        for x0, x1, ys in bloques:
            filas = []
            for ya, yb in zip(ys, ys[1:]):
                xs = sorted({round(s[0], 1) for s in V if x0 - 1 <= s[0] <= x1 + 1 and s[1] <= ya + 1.2 and s[3] >= yb - 1.2})
                xs = [x for i, x in enumerate(xs) if i == 0 or x - xs[i - 1] > 2]
                if not xs or xs[0] > x0 + 2:
                    xs = [x0] + xs
                if xs[-1] < x1 - 2:
                    xs = xs + [x1]
                filas.append((ya, yb, [(a, b) for a, b in zip(xs, xs[1:]) if b - a > 3]))
            out.append((x0, x1, filas))
        return out
    
    
    def _lineas(trs, h: float | None = None):
        if not trs:
            return []
        hs = [t[:, 1].max() - t[:, 1].min() for t in trs]
        grandes = [x for x in hs if x > 2.0]
        if h:
            hm = h                                  # altura de letra de la tabla
        else:
            tope = max(grandes) if grandes else 0
            hm = float(np.median([x for x in grandes if x > 0.45 * tope])) if grandes else 0
        altos = sorted([(t[:, 1].min(), t[:, 1].max()) for t, x in zip(trs, hs) if 0.6 * hm <= x <= 1.12 * hm and x > 2.0])
        bandas = []
        for y0, y1 in altos:
            if bandas and y0 < bandas[-1][1] - 0.5:
                bandas[-1][1] = max(bandas[-1][1], y1)
            else:
                bandas.append([y0, y1])
        if not bandas:
            bandas = [[min(t[:, 1].min() for t in trs), max(t[:, 1].max() for t in trs)]]
        L = [{"y0": b[0], "y1": b[1], "t": []} for b in bandas]
        for t in trs:
            yc = (t[:, 1].min() + t[:, 1].max()) / 2
    
            def dist(l):
                h = l["y1"] - l["y0"]
                return 0 if l["y0"] - 0.3 * h <= yc <= l["y1"] + 0.3 * h else min(abs(yc - l["y0"]), abs(yc - l["y1"])) + 100
            min(L, key=dist)["t"].append(t)
        out = []
        for l in L:
            g = []
            for t in sorted(l["t"], key=lambda t: t[:, 0].min()):
                x0, x1 = t[:, 0].min(), t[:, 0].max()
                xc = (x0 + x1) / 2
                if g:
                    G = g[-1]
                    gc = (G["x0"] + G["x1"]) / 2
                    solape = min(x1, G["x1"]) - max(x0, G["x0"])
                    if (G["x0"] - 0.05 <= xc <= G["x1"] + 0.05 or x0 - 0.05 <= gc <= x1 + 0.05
                            or solape > 0.3 * max(0.01, min(x1 - x0, G["x1"] - G["x0"])) or x0 <= G["x1"] + 0.25):
                        G["t"].append(t)
                        G["x0"], G["x1"] = min(G["x0"], x0), max(G["x1"], x1)
                        continue
                g.append({"x0": x0, "x1": x1, "t": [t]})
            if g:
                out.append({"y0": l["y0"], "y1": l["y1"], "g": g})
        return out
    
    
    def _remuestrear(t, n=12):
        d = np.sqrt(((t[1:] - t[:-1]) ** 2).sum(1))
        s = np.concatenate([[0], np.cumsum(d)])
        if s[-1] == 0:
            return np.repeat(t[:1], n, 0)
        u = np.linspace(0, s[-1], n)
        return np.stack([np.interp(u, s, t[:, 0]), np.interp(u, s, t[:, 1])], 1)
    
    
    def _rasgo(g, top, h):
        ts = sorted(g["t"], key=lambda t: (round(t[:, 0].min(), 1), t[:, 1].min()))
        P = np.concatenate([_remuestrear(t) for t in ts])
        return len(ts), np.stack([(P[:, 0] - g["x0"]) / h, (P[:, 1] - top) / h], 1)
    
    
    def _nube(P, n):
        """Puntos del glifo repartidos según el largo de cada trazo (para comparar sin importar el orden del trazo)."""
        out = []
        for b in np.split(P, n):
            largo = float(np.sqrt(((b[1:] - b[:-1]) ** 2).sum(1)).sum())
            k = max(2, int(round(largo * 14)))
            out.append(_remuestrear(b, k))
        return np.concatenate(out)
    
    
    def _chamfer(A, B):
        D = np.sqrt(((A[:, None, :] - B[None, :, :]) ** 2).sum(2))
        return float((D.min(1).mean() + D.min(0).mean()) / 2)
    
    
    def _clasificar(n, P):
        if "pl" not in _CACHE:
            _CACHE["pl"] = _plantillas()
        PL = _CACHE["pl"]
        mejor = (9.0, "?")
        # el mismo trazo puede venir dibujado en sentido contrario: se prueba también invertido
        Pr = np.concatenate([b[::-1] for b in np.split(P, n)]) if n else P
        for ch, Q in PL.get(n, []):
            if Q.shape == P.shape:
                d = min(float(np.sqrt(((Q - P) ** 2).sum(1)).mean()), float(np.sqrt(((Q - Pr) ** 2).sum(1)).mean()))
                if d < mejor[0]:
                    mejor = (d, ch)
        if mejor[0] > _UMBRAL:
            # trazo dibujado en otro orden o con otra cantidad de polilíneas: comparación por forma
            A = _nube(P, n)
            for m, lista in PL.items():
                for ch, Q in lista:
                    d = _chamfer(A, _nube(Q, m)) * 1.4
                    if d < mejor[0]:
                        mejor = (d, ch)
        return mejor
    
    
    def _cero_o(s: str) -> str:
        """En la fuente, 0 y O son iguales: es O en palabras sin otros dígitos (OLETS, BORE), 0 en lo demás."""
        def tok(m):
            t = m.group(0)
            otros = re.sub(r"0", "", t)
            return t if any(ch.isdigit() for ch in otros) or not any(ch.isalpha() for ch in otros) else t.replace("0", "O")
        return re.sub(r"[^\s]+", tok, s)
    
    
    def _texto_celda(pag: _Pagina, a, ya, b, yb, h) -> tuple[str, float]:
        real = pag.texto_real(a, ya, b, yb)
        if real is not None:
            return real, 0.0
        partes, peor = [], 0.0
        for l in _lineas(pag.trazos(a, ya, b, yb), h):
            s, prev = "", None
            for g in l["g"]:
                n, P = _rasgo(g, l["y1"] - h, h)
                d, ch = _clasificar(n, P)
                if d > _UMBRAL:
                    ch = "?"
                peor = max(peor, d)
                if prev is not None and g["x0"] - prev > _ESPACIO * h:
                    s += " "
                s += ch
                prev = g["x1"]
            partes.append(s.replace("''", '"'))
        return _cero_o(" ".join(p for p in partes if p)), peor
    
    
    def _altura(pag, filas) -> float:
        hs = []
        for ya, yb, celdas in filas:
            for a, b in celdas:
                for l in _lineas(pag.trazos(a, ya, b, yb)):
                    if len(l["g"]) >= 2:
                        hs.append(l["y1"] - l["y0"])
        return float(np.median(hs)) if hs else 7.0
    
    
    # ------------------------------------------------------------------ isométrico
    
    SECCIONES = {"PIPE": "Cañería", "FITTINGS": "Fitting", "OLETS": "Olet", "FLANGES": "Flange",
                 "FASTENERS": "Perno / empaquetadura", "VALVES": "Válvula", "PIPE SUPPORTS": "Soporte",
                 "SUPPORTS": "Soporte", "CUT PIPE LENGHTS": "CORTES", "CUT PIPE LENGTHS": "CORTES",
                 "CUT LENGTHS": "CORTES", "GASKETS": "Perno / empaquetadura", "BOLTS": "Perno / empaquetadura"}
    
    
    _BOM = ["ID", "QTY", "ND", "ITEM CODE", "DESCRIPTION", "SPEC"]
    _CABECERAS = {"CORTES": {3: ["ID", "ND", "LENGTH"]}, "Soporte": {6: ["ID", "QTY", "ND", "SUP 1", "SUP 2", "SUP 3"]},
                  **{k: {6: _BOM, 5: ["ID", "QTY", "ND", "ITEM CODE", "DESCRIPTION"]}
                     for k in ("Cañería", "Fitting", "Olet", "Flange", "Perno / empaquetadura", "Válvula")}}
    
    
    @dataclass
    class FilaLeida:
        lamina: int
        seccion: str
        datos: dict
        dudosa: bool = False
    
    
    @dataclass
    class Lectura:
        paginas: int = 0
        filas: list[FilaLeida] = field(default_factory=list)
        avisos: list[str] = field(default_factory=list)
        tablas: int = 0
        glifos_dudosos: int = 0
    
    
    def _num(s) -> float:
        s = str(s or "").upper().replace(",", ".")
        m = re.search(r"\d+(\.\d+)?", s.replace(" ", ""))
        return float(m.group(0)) if m else 0.0
    
    
    def _nd(s: str) -> str:
        s = (s or "").replace(" ", "")
        return s if s.endswith('"') or "X" in s or not s else s + '"'
    
    
    def leer_isometrico(datos: bytes, progreso=None) -> Lectura:
        import pdfplumber
        lec = Lectura()
        with pdfplumber.open(io.BytesIO(datos)) as pdf:
            lec.paginas = len(pdf.pages)
            for ip, pg in enumerate(pdf.pages):
                if progreso:
                    progreso(ip, lec.paginas)
                pag = _Pagina(pg)
                H, V = _segmentos(pg)
                for x0, x1, filas in _tablas(H, V):
                    if x1 - x0 > 900:
                        continue
                    h = _altura(pag, filas)
                    seccion, cab = None, None
                    for ya, yb, celdas in filas:
                        if len(celdas) == 1:
                            t, _ = _texto_celda(pag, *celdas[0][:1], ya, celdas[0][1], yb, h) if False else                             _texto_celda(pag, celdas[0][0], ya, celdas[0][1], yb, h)
                            clave = re.sub(r"[^A-Z ]", "", t.upper()).strip()
                            clave = re.sub(r"\s+", " ", clave)
                            if clave in SECCIONES:
                                seccion, cab = SECCIONES[clave], None
                                lec.tablas += 1
                            elif clave.startswith("CUT PIPE") or clave.startswith("CUT LENG"):
                                seccion, cab = "CORTES", None
                            continue
                        if seccion is None:
                            continue
                        textos = [_texto_celda(pag, a, ya, b, yb, h) for a, b in celdas]
                        nombres = [re.sub(r"[^A-Z0-9 ]", "", t.upper()).strip() for t, _ in textos]
                        if nombres and nombres[0] == "ID":
                            canon = _CABECERAS.get(seccion, {}).get(len(nombres))
                            cab = (canon or nombres, celdas)
                            continue
                        if cab is None:
                            continue
                        # datos con los límites de la cabecera (una línea del dibujo puede cortar una celda)
                        if len(celdas) != len(cab[1]):
                            textos = [_texto_celda(pag, a, ya, b, yb, h) for a, b in cab[1]]
                        if not any(t for t, _ in textos):
                            continue
                        d = {n: t for n, (t, _) in zip(cab[0], textos)}
                        utiles = [(t, x) for n, (t, x) in zip(cab[0], textos) if n != "SPEC"]
                        dudosa = any(x > _UMBRAL or "?" in t for t, x in utiles)
                        lec.glifos_dudosos += sum(t.count("?") for t, _ in utiles)
                        lec.filas.append(FilaLeida(ip + 1, seccion, d, dudosa))
        _validar(lec)
        return lec
    
    
    def _validar(lec: Lectura):
        """Los cortes de cada diámetro deben sumar lo que dice el BOM (el BOM redondea hacia arriba a 0,1 m)."""
        for lam in sorted({f.lamina for f in lec.filas}):
            cortes, bom = {}, {}
            for f in lec.filas:
                if f.lamina != lam:
                    continue
                if f.seccion == "CORTES":
                    nd = _nd(f.datos.get("ND", ""))
                    cortes[nd] = cortes.get(nd, 0) + _num(f.datos.get("LENGTH") or f.datos.get("LENGHT")) / 1000
                elif f.seccion == "Cañería":
                    nd = _nd(f.datos.get("ND", ""))
                    bom[nd] = bom.get(nd, 0) + _num(f.datos.get("QTY"))
            for nd in sorted(set(cortes) | set(bom)):
                c, b = cortes.get(nd, 0), bom.get(nd, 0)
                if b and not (-0.005 <= b - c <= 0.15):
                    lec.avisos.append(f"Lámina {lam}, {nd}: los cortes suman {c:.3f} m y el BOM dice {b:.1f} m. "
                                      f"Revisa los largos de esa lámina.")
                elif c and not b:
                    lec.avisos.append(f"Lámina {lam}, {nd}: hay cortes ({c:.3f} m) pero el BOM no trae cañería de ese "
                                      f"diámetro.")
    
    
    def a_planilla(lec: Lectura, peso_soporte: float | None = None) -> list[dict]:
        """Filas para la planilla de cubicación de piping."""
        out = []
        n_pag = lec.paginas
        desc_pipe = {}
        for f in lec.filas:
            if f.seccion == "Cañería":
                desc_pipe[(f.lamina, _nd(f.datos.get("ND", "")))] = f.datos.get("DESCRIPTION", "")
        for f in lec.filas:
            d = f.datos
            lam = f"{f.lamina} DE {n_pag}"
            obs = "revisar lectura" if f.dudosa else ""
            if f.seccion == "CORTES":
                nd = _nd(d.get("ND", ""))
                out.append({"Lámina": lam, "Tipo": "Cañería", "NPS": nd,
                            "Descripción": desc_pipe.get((f.lamina, nd), "PIPE"), "ID / Spool": d.get("ID", ""),
                            "Cantidad": 1.0, "Largo (mm)": _num(d.get("LENGTH") or d.get("LENGHT")), "kg c/u": None,
                            "Ítem EC-002": "", "Observación": obs})
            elif f.seccion == "Cañería":
                continue                                       # la cañería se cubica por cortes
            elif f.seccion == "Soporte":
                sup = " + ".join(x for x in (d.get("SUP 1", ""), d.get("SUP 2", ""), d.get("SUP 3", "")) if x)
                out.append({"Lámina": lam, "Tipo": "Soporte", "NPS": _nd(d.get("ND", "")), "Descripción": sup,
                            "ID / Spool": d.get("ID", ""), "Cantidad": _num(d.get("QTY")), "Largo (mm)": None,
                            "kg c/u": peso_soporte or None, "Ítem EC-002": "", "Observación": obs})
            else:
                out.append({"Lámina": lam, "Tipo": f.seccion, "NPS": _nd(d.get("ND", "")),
                            "Descripción": d.get("DESCRIPTION", ""),
                            "ID / Spool": (d.get("ITEM CODE") or "").replace(" ", ""),
                            "Cantidad": _num(d.get("QTY")), "Largo (mm)": None, "kg c/u": None, "Ítem EC-002": "",
                            "Observación": obs})
        return out

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


lector_planos = _modulo_lector_planos()

# ======================================================================
# MÓDULO: semanal
# ======================================================================
def _modulo_semanal():
    import io
    import re
    import unicodedata
    from collections import defaultdict
    from dataclasses import dataclass, field
    from datetime import date, datetime, timedelta
    
    CUB = cubicacion
    X = xlsx_xml
    
    HOJA_RESUMEN = "1.- Resumen Ejecutivo"
    HOJA_PORTADA = "0.Portada"
    HOJA_SEGURIDAD = "2. Seguridad "
    HOJA_CURVA = "3.Curva de Avance"
    HOJA_DOTACION = "4.0 Dotación"
    HOJA_CURVA_DOT = "5. Curva Dotación"
    HORAS_DIA = 11
    COLS_REV = ["D", "E", "F", "G", "H", "I"]          # Rev, Fecha, Aprobó, Revisó, Preparó, Emitido para
    
    
    def norm(s) -> str:
        s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().upper()
        return re.sub(r"\s+", " ", s).strip()
    
    
    def _hoja(L: X.Libro, nombre: str) -> str | None:
        """Nombre real de la hoja (tolera espacios y mayúsculas distintas)."""
        if nombre in L.hojas:
            return nombre
        n = norm(nombre)
        return next((h for h in L.hojas if norm(h) == n), None)
    
    
    class _Hoja:
        """Celdas de una hoja: valor en caché, fórmula y estilo."""
    
        def __init__(self, L: X.Libro, nombre: str):
            self.nombre = nombre
            self.v: dict[str, object] = {}
            self.f: dict[str, str] = {}
            self.s: dict[str, str] = {}
            xml = L.hoja_xml(nombre)
            for num, _, fila_xml, _, _ in X.filas(xml):
                for col, (a, interior, _) in X.celdas(fila_xml).items():
                    ref = f"{col}{num}"
                    if a.get("s"):
                        self.s[ref] = a["s"]
                    if interior:
                        m = re.search(r"<f\b[^>]*>(.*?)</f>", interior, flags=re.S)
                        if m:
                            self.f[ref] = X._desescapar(m.group(1))
                        self.v[ref] = L.valor_celda(a, interior)
    
        def txt(self, ref) -> str:
            v = self.v.get(ref)
            return "" if v is None else str(v)
    
        def buscar(self, patron: str, cols: str = "ABCDEFGHIJKLMN", desde: int = 1, hasta: int = 2000) -> int | None:
            """Primera fila (>= desde) donde alguna de las columnas contiene el texto (normalizado)."""
            p = norm(patron)
            filas = sorted({X.separar_ref(r)[1] for r in self.v})
            for f in filas:
                if f < desde or f > hasta:
                    continue
                for c in cols:
                    if p in norm(self.v.get(f"{c}{f}")):
                        return f
            return None
    
    
    def _fecha(v) -> date | None:
        if isinstance(v, datetime):
            return v.date()
        if isinstance(v, date):
            return v
        if isinstance(v, (int, float)) and 20000 < v < 80000:
            return X.serial_a_fecha(v)
        return None
    
    
    def _num(v) -> float | None:
        if isinstance(v, bool):
            return None
        if isinstance(v, (int, float)):
            return float(v)
        try:
            return float(str(v).replace(",", "."))
        except (TypeError, ValueError):
            return None
    
    
    # ------------------------------------------------------------------ plantilla
    
    @dataclass
    class PlantillaSemanal:
        datos: bytes
        hojas: dict[str, str]                      # nombre lógico → nombre real
        numero: int
        inicio: date
        fin: date
        # Portada
        revisiones: list[dict] = field(default_factory=list)
        filas_rev: list[int] = field(default_factory=list)
        # Resumen ejecutivo
        celda_titulo: str = "I3"
        celdas_titulo: list[str] = field(default_factory=list)
        celda_inicio: str = "H4"
        financieros: dict[str, float] = field(default_factory=dict)       # celda → valor
        etiquetas_fin: dict[str, str] = field(default_factory=dict)       # celda → etiqueta
        edp: list[dict] = field(default_factory=list)                     # {fila, nombre, monto}
        criticas: list[dict] = field(default_factory=list)                # {fila, texto, obs}
        alertas: list[dict] = field(default_factory=list)                 # {fila, tema, detalle}
        hitos: list[dict] = field(default_factory=list)                   # {fila, id, nombre, tipo, programada, real, comentario}
        # Seguridad
        seg_ultima: int = 0
        seg_hh_acum: float = 0.0
        seg_filas: dict[str, int] = field(default_factory=dict)           # etiqueta → fila del cuadro de índices
        seg_valores: dict[str, tuple] = field(default_factory=dict)       # etiqueta → (acum. anterior, acum. contrato)
        historial_hh: list[tuple] = field(default_factory=list)           # (fecha, HH acumuladas)
        # Curva de avance
        curva_actual: int = 0
        curva_fechas: dict[int, date] = field(default_factory=dict)
        # Dotación 4.0
        dot_directos: dict[str, int] = field(default_factory=dict)        # cargo normalizado → fila
        dot_indirectos: dict[str, int] = field(default_factory=dict)
        dot_nombres: dict[int, str] = field(default_factory=dict)
        dot_vacias_dir: list[int] = field(default_factory=list)
        dot_vacias_ind: list[int] = field(default_factory=list)
        dot_previo: dict[str, tuple] = field(default_factory=dict)        # cargo → (contratado, faena, descanso)
        # Curva de dotación
        mod_ultima: int = 0
        moi_ultima: int = 0
        mod_acum: float = 0.0
        moi_acum: float = 0.0
        hh_semana_previa: dict[str, float] = field(default_factory=dict)
        # 7. Avance EC-002 Itemizado
        itm_hoja: str | None = None
        itm_cols_ant: list[str] = field(default_factory=list)          # CANT, HH GAN, HH GAS, PF, % PO, % POM 5
        itm_cols_acu: list[str] = field(default_factory=list)
        itm_filas: dict[str, list[int]] = field(default_factory=dict)   # ítem → filas
        itm_info: dict[str, dict] = field(default_factory=dict)         # ítem → descripción, unidad, rendimiento, cant. actual
        itm_roll: list[int] = field(default_factory=list)               # filas (ítems y subtotales) que pasan a «anterior»
        avisos: list[str] = field(default_factory=list)
    
    
    def leer_plantilla(datos: bytes) -> PlantillaSemanal:
        L = X.Libro(datos)
        nombres = {}
        for clave, nombre in (("resumen", HOJA_RESUMEN), ("portada", HOJA_PORTADA), ("seguridad", HOJA_SEGURIDAD),
                              ("curva", HOJA_CURVA), ("dotacion", HOJA_DOTACION), ("curva_dot", HOJA_CURVA_DOT)):
            h = _hoja(L, nombre)
            if not h:
                raise ValueError(f"No se encontró la hoja «{nombre}». ¿Es un informe semanal de avance?")
            nombres[clave] = h
    
        R = _Hoja(L, nombres["resumen"])
        celda_titulo = next((r for r, v in R.v.items() if "INFORME SEMANAL N" in norm(v) and X.separar_ref(r)[1] < 10), None)
        if not celda_titulo:
            raise ValueError("No se encontró el título «INFORME SEMANAL N°…» en el Resumen Ejecutivo.")
        m = re.search(r"N\s*[°ºO]?\s*(\d+)", norm(R.v[celda_titulo]))
        numero = int(m.group(1)) if m else 0
        fila_t = X.separar_ref(celda_titulo)[1]
        celda_inicio, inicio = None, None
        for f in (fila_t + 1, fila_t, fila_t + 2):
            for c in "ABCDEFGHIJ":
                d = _fecha(R.v.get(f"{c}{f}"))
                if d and f"{c}{f}" not in R.f:
                    celda_inicio, inicio = f"{c}{f}", d
                    break
            if inicio:
                break
        if not inicio:
            raise ValueError("No se encontró la fecha de inicio de la semana en el Resumen Ejecutivo.")
        pl = PlantillaSemanal(datos, nombres, numero, inicio, inicio + timedelta(days=6), celda_titulo=celda_titulo,
                              celda_inicio=celda_inicio)
        # el título se repite en algunas hojas (fuera del área de impresión): se actualizan todos los que no son fórmula
        for clave, h in nombres.items():
            H = R if clave == "resumen" else _Hoja(L, h)
            for ref, v in H.v.items():
                if X.separar_ref(ref)[1] < 10 and ref not in H.f and re.search(r"INFORME SEMANAL N\s*[°º]", str(v or "")):
                    pl.celdas_titulo.append(f"{clave}!{ref}")
    
        # --- Portada: registro de revisiones
        P = _Hoja(L, nombres["portada"])
        f_rev = P.buscar("REGISTRO DE REVISIONES")
        if f_rev:
            f_cab = next((f for f in range(f_rev, f_rev + 5) if norm(P.v.get(f"D{f}")) == "REV"), f_rev + 2)
            for f in range(f_cab + 1, f_cab + 12):
                if not any(P.v.get(f"{c}{f}") not in (None, "", " ") for c in COLS_REV):
                    if pl.filas_rev:
                        break
                    continue
                pl.filas_rev.append(f)
                pl.revisiones.append({"rev": P.txt(f"D{f}"), "fecha": _fecha(P.v.get(f"E{f}")), "aprobo": P.txt(f"F{f}"),
                                      "reviso": P.txt(f"G{f}"), "preparo": P.txt(f"H{f}"), "emitido": P.txt(f"I{f}")})
    
        # --- Datos financieros
        f_fin = R.buscar("Datos Financieros")
        if f_fin:
            for f in range(f_fin, f_fin + 15):
                et = norm(R.v.get(f"D{f}"))
                if et and f"E{f}" in R.v and f"E{f}" not in R.f and _num(R.v.get(f"E{f}")) is not None:
                    pl.financieros[f"E{f}"] = _num(R.v[f"E{f}"])
                    pl.etiquetas_fin[f"E{f}"] = R.txt(f"D{f}")
                et_i = R.txt(f"I{f}")
                if norm(et_i).startswith("EDP"):
                    pl.edp.append({"fila": f, "nombre": et_i.strip(), "monto": _num(R.v.get(f"J{f}"))})
    
        # --- Actividades críticas / Consideraciones y alertas
        f_cri = R.buscar("Actividades Cr", cols="DEFGH")
        f_ale = R.buscar("Consideraciones y Alertas", cols="DEFGH")
        f_hit = R.buscar("Hitos Contractuales", cols="CD")
        if f_cri:
            fin_cri = (f_ale or f_cri + 15) - 1
            for f in range(f_cri + 1, fin_cri):
                if f"E{f}" in R.v or f"F{f}" in R.v:
                    pl.criticas.append({"fila": f, "texto": R.txt(f"F{f}").strip(), "obs": R.txt(f"L{f}").strip()})
        if f_ale:
            fin_ale = (f_hit or f_ale + 15) - 1
            for f in range(f_ale + 1, fin_ale):
                if R.txt(f"F{f}").strip() or R.txt(f"L{f}").strip():
                    pl.alertas.append({"fila": f, "tema": R.txt(f"F{f}").strip(), "detalle": R.txt(f"L{f}").strip()})
        if f_hit:
            f_cab = R.buscar("ID de actividad", cols="DEF", desde=f_hit)
            if f_cab:
                for f in range(f_cab + 1, f_cab + 40):
                    ide = R.txt(f"E{f}").strip()
                    if not re.match(r"[A-Z]{2}-", ide):
                        if pl.hitos:
                            break
                        continue
                    pl.hitos.append({"fila": f, "id": ide, "nombre": R.txt(f"F{f}").strip(), "tipo": R.txt(f"J{f}").strip(),
                                     "programada": _fecha(R.v.get(f"K{f}")), "real": _fecha(R.v.get(f"L{f}")),
                                     "comentario": R.txt(f"M{f}").strip()})
    
        # --- Seguridad
        S = _Hoja(L, nombres["seguridad"])
        for f in range(10, 400):
            d, n = _fecha(S.v.get(f"L{f}")), _num(S.v.get(f"N{f}"))
            if d and n is not None:
                pl.seg_ultima, pl.seg_hh_acum = f, n
                pl.historial_hh.append((d, n))
        if not pl.seg_ultima:
            pl.avisos.append("2. Seguridad: no se encontró la tabla de HH acumuladas (columnas L y N).")
        for et in ("N° De Accidentes STP", "N° De Accidentes CTP", "Días Perdidos", "Total trabajadores", "HH Trabajadas"):
            f = S.buscar(et, cols="BC")
            if f:
                pl.seg_filas[et] = f
                pl.seg_valores[et] = (_num(S.v.get(f"F{f}")), _num(S.v.get(f"H{f}")))
    
        # --- Curva de avance: fila de la semana en curso = la que tiene las fórmulas del itemizado en AD
        C = _Hoja(L, nombres["curva"])
        for f in range(55, 400):
            d = _fecha(C.v.get(f"C{f}"))
            if d:
                pl.curva_fechas[f] = d
            if "ITEMIZADO" in norm(C.f.get(f"AD{f}", "")):
                pl.curva_actual = f
        if not pl.curva_actual:
            pl.avisos.append("3. Curva de Avance: no se encontró la fila de la semana con fórmulas del itemizado (columna AD).")
    
        # --- 4.0 Dotación
        D = _Hoja(L, nombres["dotacion"])
        f_dir = D.buscar("MANO DE OBRA DIRECTA", cols="BC")
        f_ind = D.buscar("MANO DE OBRA INDIRECTA", cols="BC")
        f_tot = D.buscar("Directos", cols="BC", desde=(f_ind or 1) + 1)
        if f_dir and f_ind and f_tot:
            for f in range(f_dir + 1, f_ind):
                _registrar_cargo(pl, D, f, pl.dot_directos, pl.dot_vacias_dir, ("D", "F", "H"))
            for f in range(f_ind + 1, f_tot - 1):
                _registrar_cargo(pl, D, f, pl.dot_indirectos, pl.dot_vacias_ind, ("E", "G", "I"))
        else:
            pl.avisos.append("4.0 Dotación: no se encontraron las secciones de mano de obra directa e indirecta.")
    
        # --- 5. Curva Dotación (MOD arriba, MOI abajo)
        CD = _Hoja(L, nombres["curva_dot"])
        f_moi = CD.buscar("5.2", cols="ABCD") or 10 ** 6
        for f in range(10, 600):
            if _num(CD.v.get(f"N{f}")) is not None and _fecha(CD.v.get(f"E{f}")):
                if f < f_moi:
                    pl.mod_ultima, pl.mod_acum = f, _num(CD.v[f"N{f}"])
                else:
                    pl.moi_ultima, pl.moi_acum = f, _num(CD.v[f"N{f}"])
        for clave, fila in (("MOD", pl.mod_ultima), ("MOI", pl.moi_ultima)):
            if fila:
                pl.hh_semana_previa[clave] = _num(CD.v.get(f"L{fila}")) or 0.0
            else:
                pl.avisos.append(f"5. Curva Dotación: no se encontró la tabla de HH reales {clave}.")
        _leer_itemizado(L, pl)
        return pl
    
    
    def _leer_itemizado(L: X.Libro, pl: PlantillaSemanal):
        """Hoja «7.Avance EC-002 Itemizado»: columnas de avance anterior y acumulado, y filas de cada ítem."""
        hoja = next((h for h in L.hojas if norm(h) == "7.AVANCE EC-002 ITEMIZADO"), None)
        if not hoja:
            pl.avisos.append("No se encontró la hoja «7.Avance EC-002 Itemizado»: el avance por ítem queda igual.")
            return
        H = _Hoja(L, hoja)
        f_cab = next((f for f in range(1, 40) if any(norm(H.v.get(f"{c}{f}")) == "AVANCE ACUMULADO"
                                                     for c in ("AZ", "BA", "BB", "AY"))), None)
        if not f_cab:
            pl.avisos.append("7. Itemizado: no se encontraron las columnas «AVANCE ANTERIOR / ACUMULADO».")
            return
        col_ant = next((X.separar_ref(r)[0] for r, v in H.v.items() if X.separar_ref(r)[1] == f_cab and norm(v) == "AVANCE ANTERIOR"), None)
        col_acu = next((X.separar_ref(r)[0] for r, v in H.v.items() if X.separar_ref(r)[1] == f_cab and norm(v) == "AVANCE ACUMULADO"), None)
        if not col_ant or not col_acu:
            return
        pl.itm_hoja = hoja
        pl.itm_cols_ant = [X.num_a_col(X.col_a_num(col_ant) + i) for i in range(6)]
        pl.itm_cols_acu = [X.num_a_col(X.col_a_num(col_acu) + i) for i in range(6)]
        c_cant = pl.itm_cols_acu[0]
        for ref, v in H.v.items():
            col, f = X.separar_ref(ref)
            if col != "F" or f <= f_cab + 2 or not re.match(r"\d{2}\.\d{2}\.\d{3}\.\d{3,4}", str(v or "").strip()):
                continue
            item = str(v).strip()
            pl.itm_filas.setdefault(item, []).append(f)
            info = pl.itm_info.setdefault(item, {"descripcion": H.txt(f"M{f}").strip(), "unidad": H.txt(f"N{f}").strip(),
                                                 "especialidad": H.txt(f"E{f}").strip(), "rend": _num(H.v.get(f"AE{f}")),
                                                 "cant": 0.0, "pom": 0.0})
            info["cant"] += _num(H.v.get(f"{c_cant}{f}")) or 0.0
            info["pom"] += _num(H.v.get(f"AD{f}")) or 0.0
            if not info["rend"]:
                info["rend"] = _num(H.v.get(f"AE{f}"))
        for filas in pl.itm_filas.values():
            filas.sort()
        pl.itm_roll = sorted({X.separar_ref(r)[1] for r in H.v
                              if X.separar_ref(r)[0] in pl.itm_cols_acu and X.separar_ref(r)[1] > f_cab + 2})
    
    
    # ------------------------------------------------------------------ planillas de avance (piping, …)
    
    def leer_planilla_avance(datos: bytes, archivo: str = "") -> dict:
        """Planilla de avance de una especialidad (piping, civil, mecánica-estructuras, eléctrica): de su itemizado EC-002
        toma, por ítem, la cantidad acumulada WEEKLY (CANT bajo el primer «AVANCE ACUMULADO»). Solo los ítems que la
        planilla controla (su «Planilla de control» es una hoja del archivo, o tienen avance).
        Devuelve {"items": {ítem: {"cant": acumulada, "anterior": cant. anterior, "descripcion": …}}, "hoja": …}."""
        p = CUB.leer_planilla(datos, archivo, con_detalle=False)
        return desde_cubicacion(p)
    
    
    def desde_cubicacion(p) -> dict:
        items = {k: {"cant": it.acumulado, "anterior": it.anterior, "descripcion": it.descripcion,
                     "disciplina": it.disciplina}
                 for k, it in p.items.items() if abs(it.acumulado) > 1e-9 or abs(it.anterior) > 1e-9}
        if not items:
            raise ValueError("no se encontraron ítems con avance en el itemizado de la planilla.")
        return {"archivo": p.archivo, "hoja": p.hoja_itemizado, "items": items}
    
    
    def _registrar_cargo(pl, D: _Hoja, f: int, destino: dict, vacias: list, cols: tuple):
        nombre = D.txt(f"B{f}").strip()
        if not nombre:
            if destino and any(f"{c}{f}" in D.s for c in cols):      # fila libre entre cargos (no la de separación)
                vacias.append(f)
            return
        if norm(nombre).startswith("MANO DE OBRA"):
            return
        destino[norm(nombre)] = f
        pl.dot_nombres[f] = nombre
        pl.dot_previo[norm(nombre)] = tuple(_num(D.v.get(f"{c}{f}")) for c in cols)
    
    
    # ------------------------------------------------------------------ datos de la semana (BD Personal)
    
    @dataclass
    class DatosPersonal:
        filas: list[tuple]                  # (fecha, cargo, moi_mod, tipo_turno, codigo, acreditado)
        fechas: set[date]
    
    
    def leer_bd_personal(datos_diario: bytes) -> DatosPersonal:
        """Lee la hoja BD Personal de un informe diario (todas las fechas)."""
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(datos_diario), read_only=True, data_only=True)
        if "BD Personal" not in wb.sheetnames:
            raise ValueError("El informe diario no tiene la hoja «BD Personal».")
        ws = wb["BD Personal"]
        cab, idx, filas = None, {}, []
        for r in ws.iter_rows(min_row=1, max_col=40, values_only=True):
            if cab is None:
                n = [norm(x) for x in r]
                if "RUT" in n and "CARGO" in n:
                    cab = n
                    def col(*nombres, empieza=False):
                        for o in nombres:
                            for j, x in enumerate(n):
                                if (x.startswith(o) if empieza else x == o):
                                    return j
                        return None
                    idx = {"fecha": col("FECHA"), "cargo": col("CARGO"), "tipo": col("MOI/MOD"),
                           "turno": col("TIPO DE TURNO"), "cod": col("PERSONAL EN FAENA", empieza=True),
                           "acr": col("ACREDITADO")}
                    if None in (idx["fecha"], idx["cargo"], idx["tipo"], idx["cod"]):
                        raise ValueError("BD Personal sin las columnas Fecha, Cargo, MOI/MOD o Personal en Faena.")
                continue
            d = _fecha(r[idx["fecha"]]) if idx["fecha"] < len(r) else None
            if not d:
                continue
            g = lambda k: (r[idx[k]] if idx.get(k) is not None and idx[k] < len(r) else None)
            filas.append((d, str(g("cargo") or "").strip(), norm(g("tipo")), norm(g("turno")), norm(g("cod")), norm(g("acr"))))
        wb.close()
        return DatosPersonal(filas, {f[0] for f in filas})
    
    
    def _en_faena(fila) -> bool:
        _, _, tipo, turno, cod, acr = fila
        if turno.startswith("NOCHE"):
            return cod == "TN"
        if tipo.startswith("IND"):
            return cod in ("T", "TT", "TC") and acr == "SI"
        return cod == "T"
    
    
    @dataclass
    class Semana:
        inicio: date
        fin: date
        numero: int
        hh_mod: float = 0.0
        hh_moi: float = 0.0
        dotacion: dict[str, dict] = field(default_factory=dict)   # cargo → {tipo, contratado, faena, descanso}
        dias_sin_datos: list[date] = field(default_factory=list)
    
    
    def calcular_semana(pl: PlantillaSemanal, bd: DatosPersonal, inicio: date | None = None) -> Semana:
        inicio = inicio or pl.inicio + timedelta(days=7)
        fin = inicio + timedelta(days=6)
        dias = {inicio + timedelta(days=i) for i in range(7)}
        sem = Semana(inicio, fin, pl.numero + 1, dias_sin_datos=sorted(dias - bd.fechas))
        n_dias = max(1, len(dias & bd.fechas))
        acum = defaultdict(lambda: {"tipo": "", "contratado": 0, "faena": 0})
        for fila in bd.filas:
            if fila[0] not in dias:
                continue
            faena = _en_faena(fila)
            if faena:
                if fila[2].startswith("IND"):
                    sem.hh_moi += HORAS_DIA
                else:
                    sem.hh_mod += HORAS_DIA
            a = acum[fila[1]]
            a["tipo"] = "Indirecto" if fila[2].startswith("IND") else "Directo"
            a["contratado"] += 1
            a["faena"] += faena
        for cargo, a in acum.items():
            c, f = round(a["contratado"] / n_dias, 1), round(a["faena"] / n_dias, 1)
            sem.dotacion[cargo] = {"tipo": a["tipo"], "contratado": c, "faena": f, "descanso": round(c - f, 1)}
        return sem
    
    
    # ------------------------------------------------------------------ entradas editables
    
    @dataclass
    class EntradasSemanal:
        numero: int
        inicio: date
        revisiones: list[dict]
        financieros: dict[str, float]
        edp: list[dict]
        criticas: list[dict]
        alertas: list[dict]
        hitos: list[dict]
        seguridad: dict[str, float]          # etiqueta → valor del período (accidentes, días perdidos)
        hh_mod: float
        hh_moi: float
        dotacion: dict[str, dict]
        avance_items: dict[str, float] = field(default_factory=dict)   # ítem → cantidad acumulada nueva (planillas)
        fuentes_avance: list[str] = field(default_factory=list)
    
    
    def hitos_desde_p6(pl: PlantillaSemanal, actividades) -> list[dict]:
        """Fecha real de los hitos terminados según el último programa P6 cargado (actividades del corte)."""
        salida = [dict(h) for h in pl.hitos]
        if actividades is None or len(actividades) == 0:
            return salida
        por_id = {str(r["task_code"]).strip(): r for _, r in actividades.iterrows()}
        for h in salida:
            r = por_id.get(h["id"].strip())
            if r is not None and str(r.get("estado", "")).startswith("Complet") and not h.get("real"):
                d = _fecha(datetime.fromisoformat(str(r["fin"])[:10])) if r.get("fin") else None
                if d:
                    h["real"] = d
                    h["origen"] = "P6"
        return salida
    
    
    # ------------------------------------------------------------------ generación
    
    def _cargos_nuevos(pl: PlantillaSemanal, dotacion: dict) -> dict[str, list[str]]:
        nuevos = {"Directo": [], "Indirecto": []}
        for cargo, d in dotacion.items():
            k = norm(cargo)
            if k not in pl.dot_directos and k not in pl.dot_indirectos:
                nuevos["Indirecto" if d.get("tipo") == "Indirecto" else "Directo"].append(cargo)
        return nuevos
    
    
    def _ampliar_dotacion(pl: PlantillaSemanal, dotacion: dict) -> PlantillaSemanal:
        """Si hay cargos nuevos y no quedan filas libres en 4.0 Dotación, inserta las filas que faltan dentro de la sección
        (así los totales =SUM(...) las incluyen) y vuelve a leer la planilla."""
        nuevos = _cargos_nuevos(pl, dotacion)
        falta_d = len(nuevos["Directo"]) - len(pl.dot_vacias_dir)
        falta_i = len(nuevos["Indirecto"]) - len(pl.dot_vacias_ind)
        if falta_d <= 0 and falta_i <= 0:
            return pl
        L = X.Libro(pl.datos)
        for filas, falta in sorted([(list(pl.dot_indirectos.values()), falta_i), (list(pl.dot_directos.values()), falta_d)],
                                   key=lambda t: -(max(t[0]) if t[0] else 0)):
            if falta > 0 and filas:
                X.insertar_filas(L, pl.hojas["dotacion"], max(filas), falta)
        return _marcar_copias(leer_plantilla(L.bytes(recalcular=True)))
    
    
    def _marcar_copias(pl: PlantillaSemanal) -> PlantillaSemanal:
        """Tras insertar filas, el mismo cargo aparece repetido (copias de la fila modelo): todas menos la última quedan libres."""
        for seccion, vacias in (("dot_directos", "dot_vacias_dir"), ("dot_indirectos", "dot_vacias_ind")):
            por_cargo = defaultdict(list)
            for f in sorted(pl.dot_nombres):
                nombre = norm(pl.dot_nombres[f])
                if getattr(pl, seccion).get(nombre) is not None:
                    por_cargo[nombre].append(f)
            for nombre, filas in por_cargo.items():
                if len(filas) > 1:
                    getattr(pl, seccion)[nombre] = filas[-1]
                    getattr(pl, vacias).extend(filas[:-1])
            getattr(pl, vacias).sort()
        return pl
    
    
    def generar(pl: PlantillaSemanal, ent: EntradasSemanal) -> tuple[bytes, list[str]]:
        pl = _ampliar_dotacion(pl, ent.dotacion)
        L = X.Libro(pl.datos)
        avisos: list[str] = []
        fin = ent.inicio + timedelta(days=6)
    
        # --- Resumen ejecutivo
        R = _Hoja(L, pl.hojas["resumen"])
        def titulo_nuevo(t):
            return re.sub(r"(N\s*[°º]\s*)\d+", lambda m: f"{m.group(1)}{ent.numero}", t)
        por_hoja: dict[str, dict] = defaultdict(dict)
        for item in pl.celdas_titulo:
            clave, ref = item.split("!")
            if clave != "resumen":
                por_hoja[clave][ref] = titulo_nuevo(_Hoja(L, pl.hojas[clave]).txt(ref))
        for clave, cam in por_hoja.items():
            L.guardar_hoja(pl.hojas[clave], X.fijar_celdas(L.hoja_xml(pl.hojas[clave]), cam))
        cambios: dict[str, object] = {pl.celda_inicio: ent.inicio}
        for item in pl.celdas_titulo:
            clave, ref = item.split("!")
            if clave == "resumen":
                cambios[ref] = titulo_nuevo(R.txt(ref))
        for ref, v in ent.financieros.items():
            cambios[ref] = v
        for e in ent.edp:
            cambios[f"J{e['fila']}"] = e.get("monto")
        for c in ent.criticas:
            cambios[f"F{c['fila']}"] = c.get("texto") or None
            cambios[f"L{c['fila']}"] = c.get("obs") or None
        for a in ent.alertas:
            cambios[f"F{a['fila']}"] = a.get("tema") or None
            cambios[f"L{a['fila']}"] = a.get("detalle") or None
        for h in ent.hitos:
            cambios[f"L{h['fila']}"] = h.get("real")
            cambios[f"M{h['fila']}"] = h.get("comentario") or None
        L.guardar_hoja(pl.hojas["resumen"], X.fijar_celdas(L.hoja_xml(pl.hojas["resumen"]), cambios))
    
        # --- Portada
        P = _Hoja(L, pl.hojas["portada"])
        cambios, estilos = {}, {}
        filas_rev = list(pl.filas_rev)
        while len(filas_rev) < len(ent.revisiones):
            filas_rev.append((filas_rev[-1] if filas_rev else 26) + 1)
        for i, f in enumerate(filas_rev):
            r = ent.revisiones[i] if i < len(ent.revisiones) else {}
            for col, k in zip(COLS_REV, ("rev", "fecha", "aprobo", "reviso", "preparo", "emitido")):
                cambios[f"{col}{f}"] = r.get(k) or None
                if pl.filas_rev:
                    estilos[f"{col}{f}"] = P.s.get(f"{col}{pl.filas_rev[-1]}")
        L.guardar_hoja(pl.hojas["portada"], X.fijar_celdas(L.hoja_xml(pl.hojas["portada"]), cambios, estilos))
    
        # --- Seguridad
        S = _Hoja(L, pl.hojas["seguridad"])
        if pl.seg_ultima:
            u, n = pl.seg_ultima, pl.seg_ultima + 1
            hh_sem = ent.hh_mod + ent.hh_moi
            cambios = {f"L{n}": (f"=+L{u}+7", X.fecha_a_serial(fin)), f"M{n}": (f"=+M{u}+1",), f"N{n}": round(pl.seg_hh_acum + hh_sem),
                       f"O{n}": (f"=+(N{n}/11)*12",), f"Q{n}": (f"=+O{n}/12/7",), f"R{n}": (f"=+Q{n}-Q{u}",)}
            estilos = {f"{c}{n}": S.s.get(f"{c}{u}") for c in "LMNOQR"}
            # Acumulado anterior = acumulado del informe anterior (filas «Total trabajadores» y «HH Trabajadas»)
            q_u, o_u = _num(S.v.get(f"Q{u}")), _num(S.v.get(f"O{u}"))
            if "Total trabajadores" in pl.seg_filas:
                f = pl.seg_filas["Total trabajadores"]
                cambios[f"F{f}"] = q_u
                cambios[f"H{f}"] = (f"=+Q{n}",)
            if "HH Trabajadas" in pl.seg_filas:
                f = pl.seg_filas["HH Trabajadas"]
                cambios[f"F{f}"] = o_u
                cambios[f"H{f}"] = (f"=+O{n}",)
            for et in ("N° De Accidentes STP", "N° De Accidentes CTP", "Días Perdidos"):
                if et in pl.seg_filas:
                    f = pl.seg_filas[et]
                    previo = pl.seg_valores[et][1] or 0
                    cambios[f"F{f}"] = previo
                    cambios[f"H{f}"] = previo + (ent.seguridad.get(et) or 0)
            L.guardar_hoja(pl.hojas["seguridad"], X.fijar_celdas(L.hoja_xml(pl.hojas["seguridad"]), cambios, estilos))
    
        # --- Curva de avance: congelar la semana anterior y pasar las fórmulas a la nueva
        C = _Hoja(L, pl.hojas["curva"])
        if pl.curva_actual:
            u, n = pl.curva_actual, pl.curva_actual + 1
            if pl.curva_fechas.get(n) and pl.curva_fechas[n] != fin:
                avisos.append(f"3. Curva de Avance: la fila siguiente es del {pl.curva_fechas[n]:%d-%m-%Y} y la semana nueva "
                              f"termina el {fin:%d-%m-%Y}; revisa la tabla.")
            cambios = {}
            for col in ("AB", "AC", "AD", "AE", "AF", "AG"):
                ref = f"{col}{u}"
                if ref in C.f:
                    cambios[ref] = C.v.get(ref)                                  # queda en valor
                    cambios[f"{col}{n}"] = ("=" + X.desplazar_relativas(C.f[ref], 1),)
            L.guardar_hoja(pl.hojas["curva"], X.fijar_celdas(L.hoja_xml(pl.hojas["curva"]), cambios,
                                                              {f"{c}{n}": C.s.get(f"{c}{u}") for c in ("AB", "AC", "AD", "AE", "AF", "AG")}))
    
        # --- 4.0 Dotación
        cambios = {}
        vistos = set()
        libres = {"Directo": list(pl.dot_vacias_dir), "Indirecto": list(pl.dot_vacias_ind)}
        D = _Hoja(L, pl.hojas["dotacion"])
        estilos = {}
        for cargo, d in sorted(ent.dotacion.items()):
            k = norm(cargo)
            directo = d.get("tipo") != "Indirecto"
            fila = (pl.dot_directos if directo else pl.dot_indirectos).get(k) or \
                (pl.dot_indirectos if directo else pl.dot_directos).get(k)
            if fila is None:
                lista = libres["Directo" if directo else "Indirecto"]
                if not lista:
                    avisos.append(f"4.0 Dotación: no hay fila libre para el cargo nuevo «{cargo}»; agrégalo a mano.")
                    continue
                fila = lista.pop(0)
                cambios[f"B{fila}"] = cargo
            vistos.add(fila)
            es_dir = fila in pl.dot_directos.values() or (fila not in pl.dot_indirectos.values() and directo)
            if es_dir:
                cambios[f"D{fila}"], cambios[f"F{fila}"], cambios[f"H{fila}"] = d["contratado"], d["faena"], d["descanso"]
            else:
                cambios[f"E{fila}"], cambios[f"G{fila}"] = d["contratado"], d["faena"]
                if f"I{fila}" not in D.f:
                    cambios[f"I{fila}"] = (f"=+E{fila}-G{fila}",)
        for k, fila in list(pl.dot_directos.items()):
            if fila not in vistos:
                cambios[f"D{fila}"] = cambios[f"F{fila}"] = cambios[f"H{fila}"] = None
        for k, fila in list(pl.dot_indirectos.items()):
            if fila not in vistos:
                cambios[f"E{fila}"] = cambios[f"G{fila}"] = None
                if f"I{fila}" not in D.f:
                    cambios[f"I{fila}"] = None
        L.guardar_hoja(pl.hojas["dotacion"], X.fijar_celdas(L.hoja_xml(pl.hojas["dotacion"]), cambios, estilos))
    
        # --- 5. Curva Dotación
        CD = _Hoja(L, pl.hojas["curva_dot"])
        cambios, estilos = {}, {}
        for u, acum, hh in ((pl.mod_ultima, pl.mod_acum, ent.hh_mod), (pl.moi_ultima, pl.moi_acum, ent.hh_moi)):
            if not u:
                continue
            n = u + 1
            cambios[f"L{n}"] = (f"=+N{n}-N{u}", hh)
            cambios[f"M{n}"] = (f"=+L{n}/77", hh / 77)
            cambios[f"N{n}"] = round(acum + hh)
            for c in "LMN":
                estilos[f"{c}{n}"] = CD.s.get(f"{c}{u}")
        L.guardar_hoja(pl.hojas["curva_dot"], X.fijar_celdas(L.hoja_xml(pl.hojas["curva_dot"]), cambios, estilos))
    
        # --- 7. Itemizado: el acumulado del informe anterior pasa a «avance anterior»; el acumulado nuevo sale de las
        #     planillas de avance cargadas (los ítems sin planilla quedan con el mismo acumulado: período 0).
        if pl.itm_hoja:
            I = _Hoja(L, pl.itm_hoja)
            cambios = {}
            for f in pl.itm_roll:
                if True:
                    for ca, cu in zip(pl.itm_cols_ant, pl.itm_cols_acu):
                        v = I.v.get(f"{cu}{f}")
                        if f"{ca}{f}" in I.f:
                            continue
                        cambios[f"{ca}{f}"] = v if isinstance(v, (int, float)) and not isinstance(v, bool) else None
            c_cant = pl.itm_cols_acu[0]
            for item, cant in ent.avance_items.items():
                filas = pl.itm_filas.get(item)
                if not filas:
                    continue
                libres = [f for f in filas if f"{c_cant}{f}" not in I.f]
                if not libres:
                    avisos.append(f"7. Itemizado: el ítem {item} tiene fórmula en la cantidad acumulada; no se cambió.")
                    continue
                otras = sum(_num(I.v.get(f"{c_cant}{f}")) or 0 for f in filas if f != libres[0])
                cambios[f"{c_cant}{libres[0]}"] = round(cant - otras, 6)
            L.guardar_hoja(pl.itm_hoja, X.fijar_celdas(L.hoja_xml(pl.itm_hoja), cambios))
    
        return L.bytes(recalcular=True), avisos
    
    
    def entradas_sugeridas(pl: PlantillaSemanal, sem: Semana, actividades=None) -> EntradasSemanal:
        fin = sem.fin
        revs = [dict(r) for r in pl.revisiones[:1]] or [{"rev": "Rev.A", "aprobo": "", "reviso": "", "preparo": "", "emitido": ""}]
        revs[0].update({"rev": "Rev.A", "fecha": fin})
        return EntradasSemanal(
            numero=sem.numero, inicio=sem.inicio, revisiones=revs, financieros=dict(pl.financieros),
            edp=[dict(e) for e in pl.edp], criticas=[dict(c) for c in pl.criticas], alertas=[dict(a) for a in pl.alertas],
            hitos=hitos_desde_p6(pl, actividades), seguridad={"N° De Accidentes STP": 0, "N° De Accidentes CTP": 0, "Días Perdidos": 0},
            hh_mod=sem.hh_mod, hh_moi=sem.hh_moi, dotacion=sem.dotacion)

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


semanal = _modulo_semanal()

# ======================================================================
# MÓDULO: comun
# ======================================================================
def _modulo_comun():
    import pandas as pd
    import streamlit as st
    
    
    NARANJO = "#FF6B00"
    AZUL_OSCURO = "#0F172A"
    AZUL = "#3B82F6"
    VERDE = "#10B981"
    GRIS = "#94A3B8"
    PIZARRA = "#64748B"   # serie neutra: se lee sobre fondo claro y oscuro (AZUL_OSCURO desaparece en modo noche)
    COLORES_ESTADO = {"Completado": VERDE, "En Progreso": NARANJO, "No Iniciado": PIZARRA}
    
    
    # --------------------------------------------------------------- sesión
    
    def usuario() -> dict | None:
        return st.session_state.get("usuario")
    
    
    def exigir(accion: str) -> bool:
        if auth.puede(usuario(), accion):
            return True
        st.warning("Tu rol no tiene permiso para esta acción.")
        return False
    
    
    def contrato() -> dict | None:
        """Contrato activo de esta sesión (lo elige el usuario en la barra lateral)."""
        return st.session_state.get("contrato")
    
    
    def contrato_id() -> int:
        ct = contrato()
        if not ct:
            st.warning("Elige un contrato en la barra lateral.")
            st.stop()
        return int(ct["id"])
    
    
    def nombre_usuario() -> str:
        u = usuario()
        return u["username"] if u else ""
    
    
    def limpiar_cache():
        st.cache_data.clear()
    
    
    # --------------------------------------------------------------- formato (es-CL)
    
    def num(x, dec: int = 0) -> str:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return "–"
        s = f"{x:,.{dec}f}"
        return s.replace(",", "§").replace(".", ",").replace("§", ".")
    
    
    def porc(x, dec: int = 2) -> str:
        return f"{num(x, dec)}%"
    
    
    def fecha(s) -> str:
        if not s or (isinstance(s, float) and pd.isna(s)):
            return "–"
        s = str(s)[:10]
        try:
            return pd.Timestamp(s).strftime("%d-%m-%Y")
        except ValueError:
            return s
    
    
    def layout_grafico(fig, alto: int | None = None):
        fig.update_layout(
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=10, r=10, t=40, b=10), legend_title_text="",
            font=dict(family="Inter, Helvetica, Arial, sans-serif"),
            separators=",.",
        )
        if alto:
            fig.update_layout(height=alto)
        return fig
    
    
    # --------------------------------------------------------------- datos (con caché)
    
    def proyectos() -> pd.DataFrame:
        return _proyectos(contrato_id())
    
    
    @st.cache_data(ttl=600, show_spinner=False)
    def _proyectos(cid: int) -> pd.DataFrame:
        return db.proyectos(cid)
    
    
    @st.cache_data(ttl=600, show_spinner=False)
    def cortes(proyecto_id: int) -> pd.DataFrame:
        return db.cortes(proyecto_id)
    
    
    @st.cache_data(ttl=600, show_spinner=False)
    def actividades(corte_id: int) -> pd.DataFrame:
        return db.actividades(corte_id)
    
    
    @st.cache_data(ttl=600, show_spinner=False)
    def tipos_codigo(corte_id: int) -> list[str]:
        return db.tipos_codigo(corte_id)
    
    
    @st.cache_data(ttl=600, show_spinner=False)
    def codigos(corte_id: int, tipo: str) -> pd.DataFrame:
        return db.codigos(corte_id, tipo)
    
    
    # --------------------------------------------------------------- selectores
    
    def selector_proyecto_corte(clave: str) -> tuple[pd.Series, pd.Series, pd.DataFrame] | None:
        """Selector de proyecto y fecha de corte. Devuelve (proyecto, corte, todos_los_cortes)."""
        ps = proyectos()
        if ps.empty:
            st.info("Aún no hay proyectos cargados. Ve a **Importar .XER** para sincronizar un programa de Primavera P6.")
            return None
        c1, c2 = st.columns([2, 1])
        etiquetas = [f"{r.codigo} — {r.nombre}" if r.nombre != r.codigo else r.codigo for r in ps.itertuples()]
        i = c1.selectbox("Proyecto", range(len(ps)), format_func=lambda k: etiquetas[k], key=f"{clave}_proy",
                         index=_indice_guardado("proy_idx", len(ps)))
        st.session_state["proy_idx"] = i
        proyecto = ps.iloc[i]
        cs = cortes(int(proyecto["id"]))
        if cs.empty:
            st.warning("Este proyecto no tiene cortes cargados.")
            return None
        orden = list(range(len(cs)))[::-1]  # más reciente primero
        j = c2.selectbox("Fecha de corte (Data Date)", orden, format_func=lambda k: fecha(cs.iloc[k]["fecha_corte"]),
                         key=f"{clave}_corte")
        return proyecto, cs.iloc[j], cs
    
    
    def _indice_guardado(clave: str, n: int) -> int:
        i = st.session_state.get(clave, 0)
        return i if isinstance(i, int) and 0 <= i < n else 0
    
    
    def selector_tipo_codigo(corte_id: int, clave: str) -> str | None:
        tipo_codigo_por_defecto = xer.tipo_codigo_por_defecto
    
        tipos = tipos_codigo(corte_id)
        if not tipos:
            st.info("Las actividades de este corte no tienen códigos de actividad asignados en P6.")
            return None
        defecto = tipo_codigo_por_defecto(tipos)
        return st.selectbox("Agrupar por código de actividad P6", tipos, index=tipos.index(defecto), key=clave,
                            help="Los códigos salen del propio .XER (Subsistemas, Sistema, Área, Fase…).")
    
    
    COLUMNAS_HORAS = {
        "budget": st.column_config.NumberColumn("Budget (HH)", format="%.0f"),
        "at_completion": st.column_config.NumberColumn("At Completion (HH)", format="%.0f"),
        "actual": st.column_config.NumberColumn("Actual (HH)", format="%.0f"),
        "remaining": st.column_config.NumberColumn("Remaining (HH)", format="%.0f"),
        "avance_pct": st.column_config.ProgressColumn("% Avance", format="%.2f%%", min_value=0, max_value=100),
    }

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


comun = _modulo_comun()

# ======================================================================
# MÓDULO: tema
# ======================================================================
def _modulo_tema():
    import base64
    import io
    import os
    
    import streamlit as st
    import streamlit.components.v1 as components
    
    VERSION = "3.3"
    FECHA_VERSION = "13-11-2026"
    NARANJO = "#FA4D09"
    AZUL = "#06205C"
    
    CSS = """<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    :root {
      --sy-naranjo:#FA4D09; --sy-naranjo2:#FF7A3D; --sy-naranjo-hover:#E04406;
      --sy-azul:#06205C; --sy-azul2:#0B2F7A; --sy-azul3:#163E94; --sy-azul-osc:#041640;
      --sy-fondo:#F3F5F9; --sy-tarjeta:#FFFFFF; --sy-tarjeta-borde:#E3E8F1; --sy-titulo:#06205C;
      --sy-sombra:0 6px 16px rgba(6,32,92,.06); --sy-suave:#64708A; --sy-campo:#FAFBFD;
      --sy-curva:cubic-bezier(.22,.61,.36,1);
    }
    html[data-sy-tema="oscuro"] {
      --sy-naranjo:#FF5A1A; --sy-naranjo2:#FF8A50; --sy-naranjo-hover:#FF6F33;
      --sy-fondo:#0B1222; --sy-tarjeta:#121B30; --sy-tarjeta-borde:rgba(255,255,255,.08); --sy-titulo:#E8EEFB;
      --sy-sombra:0 6px 18px rgba(0,0,0,.35); --sy-suave:#94A3B8; --sy-campo:#0F172A;
    }
    html, body, [data-testid="stApp"], button, input, textarea, select {font-family:'Inter','Segoe UI',system-ui,sans-serif;}
    [data-testid="stApp"], [data-testid="stHeader"] {background:var(--sy-fondo);}
    .block-container {padding-top:2.2rem;}
    
    /* ---------- Animaciones ---------- */
    @keyframes sy-sube {from {opacity:0; transform:translateY(14px);} to {opacity:1; transform:none;}}
    @keyframes sy-aparece {from {opacity:0;} to {opacity:1;}}
    @keyframes sy-flota {0%,100% {transform:rotate(45deg) translateY(0);} 50% {transform:rotate(45deg) translateY(-14px);}}
    @keyframes sy-gira {from {transform:rotate(45deg);} to {transform:rotate(405deg);}}
    @keyframes sy-entra-rombo {from {opacity:0; transform:rotate(0deg) scale(.6);} to {opacity:1; transform:rotate(45deg) scale(1);}}
    @keyframes sy-late {0%,100% {box-shadow:0 0 0 0 rgba(250,77,9,.45);} 50% {box-shadow:0 0 0 7px rgba(250,77,9,0);}}
    @keyframes sy-brillo {from {background-position:-200% 0;} to {background-position:200% 0;}}
    @keyframes sy-desliza {from {opacity:0; transform:translateX(-12px);} to {opacity:1; transform:none;}}
    [data-testid="stMain"] .block-container > div {animation:sy-aparece .45s var(--sy-curva);}
    @media (prefers-reduced-motion: reduce) {*, *::before, *::after {animation:none !important; transition:none !important;}}
    
    /* ---------- Títulos ---------- */
    [data-testid="stMain"] h1 {color:var(--sy-titulo); font-weight:800; letter-spacing:-.5px; position:relative; padding-left:18px;
      animation:sy-desliza .5s var(--sy-curva);}
    [data-testid="stMain"] h1::before {content:""; position:absolute; left:0; top:22%; width:9px; height:9px; border-radius:2px;
      background:var(--sy-naranjo); transform:rotate(45deg); box-shadow:0 0 0 4px rgba(250,77,9,.15);}
    [data-testid="stMain"] h2, [data-testid="stMain"] h3 {color:var(--sy-titulo); font-weight:700;}
    
    /* ---------- Tarjetas de indicadores ---------- */
    [data-testid="stMetric"] {background:var(--sy-tarjeta); border:1px solid var(--sy-tarjeta-borde); border-left:4px solid var(--sy-azul);
      border-radius:14px; padding:16px 18px; box-shadow:var(--sy-sombra); position:relative; overflow:hidden;
      animation:sy-sube .55s var(--sy-curva) both; transition:transform .25s var(--sy-curva), box-shadow .25s;}
    html[data-sy-tema="oscuro"] [data-testid="stMetric"] {border-left-color:#4C79E0;}
    [data-testid="stMetric"]::after {content:""; position:absolute; right:-26px; top:-26px; width:70px; height:70px; border-radius:16px;
      transform:rotate(45deg); background:rgba(6,32,92,.06); transition:transform .4s var(--sy-curva);}
    [data-testid="stMetric"]:hover {transform:translateY(-3px); box-shadow:0 12px 26px rgba(6,32,92,.12);}
    [data-testid="stMetric"]:hover::after {transform:rotate(90deg) scale(1.15);}
    [data-testid="stColumn"]:nth-child(even) [data-testid="stMetric"] {border-left-color:var(--sy-naranjo);}
    [data-testid="stColumn"]:nth-child(even) [data-testid="stMetric"]::after {background:rgba(250,77,9,.12);}
    [data-testid="stColumn"]:nth-child(2) [data-testid="stMetric"] {animation-delay:.07s;}
    [data-testid="stColumn"]:nth-child(3) [data-testid="stMetric"] {animation-delay:.14s;}
    [data-testid="stColumn"]:nth-child(4) [data-testid="stMetric"] {animation-delay:.21s;}
    [data-testid="stColumn"]:nth-child(n+5) [data-testid="stMetric"] {animation-delay:.28s;}
    [data-testid="stMetricLabel"] p {color:var(--sy-suave) !important; font-weight:600;}
    [data-testid="stMetricValue"] {color:var(--sy-titulo); font-weight:800;}
    
    /* ---------- Botones ---------- */
    [data-testid^="stBaseButton-primary"] {background:linear-gradient(90deg,var(--sy-naranjo),var(--sy-naranjo2)) !important;
      border:0 !important; color:#FFFFFF !important; border-radius:10px !important; font-weight:600 !important;
      box-shadow:0 8px 18px rgba(250,77,9,.28); transition:transform .2s var(--sy-curva), box-shadow .2s, filter .2s !important;}
    [data-testid^="stBaseButton-primary"]:hover {transform:translateY(-2px); box-shadow:0 12px 24px rgba(250,77,9,.38); filter:brightness(1.05);}
    [data-testid^="stBaseButton-primary"]:active {transform:translateY(0);}
    [data-testid^="stBaseButton-primary"] * {color:#FFFFFF !important;}
    [data-testid="stMain"] [data-testid^="stBaseButton-secondary"] {border-radius:10px !important; border:1.5px solid var(--sy-tarjeta-borde) !important;
      background:var(--sy-tarjeta) !important; font-weight:600 !important; transition:all .2s var(--sy-curva) !important;}
    [data-testid="stMain"] [data-testid^="stBaseButton-secondary"]:hover {border-color:var(--sy-azul3) !important; color:var(--sy-azul3) !important;
      transform:translateY(-1px);}
    html[data-sy-tema="oscuro"] [data-testid="stMain"] [data-testid^="stBaseButton-secondary"]:hover {border-color:#7FA2F0 !important; color:#C9D8FA !important;}
    [data-testid="stDownloadButton"] button {border-radius:10px !important;}
    
    /* ---------- Pestañas tipo píldora ---------- */
    [data-baseweb="tab-list"] {background:var(--sy-tarjeta); border:1px solid var(--sy-tarjeta-borde); border-radius:12px; padding:5px; gap:4px;
      box-shadow:var(--sy-sombra);}
    [data-baseweb="tab"] {background:transparent; border-radius:9px !important; padding:8px 14px !important; height:auto !important; transition:background .25s, color .25s;}
    [data-baseweb="tab"] p {font-weight:600; color:var(--sy-suave);}
    [data-baseweb="tab"]:hover {background:rgba(6,32,92,.06);}
    [data-baseweb="tab"][aria-selected="true"] {background:var(--sy-azul) !important;}
    html[data-sy-tema="oscuro"] [data-baseweb="tab"][aria-selected="true"] {background:var(--sy-azul3) !important;}
    [data-baseweb="tab"][aria-selected="true"] p {color:#FFFFFF !important;}
    [data-baseweb="tab-highlight"], [data-baseweb="tab-border"] {display:none !important;}
    [data-baseweb="tab-panel"] {animation:sy-sube .35s var(--sy-curva);}
    
    /* ---------- Campos, expansores, formularios, avisos ---------- */
    [data-baseweb="input"], [data-baseweb="select"] > div, [data-baseweb="textarea"], [data-baseweb="base-input"] {border-radius:10px !important;}
    [data-testid="stMain"] [data-baseweb="input"]:focus-within, [data-testid="stMain"] [data-baseweb="select"] > div:focus-within {
      border-color:var(--sy-azul3) !important; box-shadow:0 0 0 4px rgba(22,62,148,.12);}
    [data-testid="stExpander"] details {border-radius:12px; border:1px solid var(--sy-tarjeta-borde); background:var(--sy-tarjeta);
      box-shadow:var(--sy-sombra); transition:box-shadow .25s;}
    [data-testid="stExpander"] details:hover {box-shadow:0 10px 22px rgba(6,32,92,.09);}
    [data-testid="stExpander"] summary:hover p {color:var(--sy-naranjo);}
    [data-testid="stForm"] {border-radius:14px; border:1px solid var(--sy-tarjeta-borde); background:var(--sy-tarjeta);}
    [data-testid="stAlert"] {border-radius:12px;}
    [data-testid="stFileUploaderDropzone"] {border-radius:12px; border:1.5px dashed rgba(22,62,148,.35); transition:border-color .2s, background .2s;}
    [data-testid="stFileUploaderDropzone"]:hover {border-color:var(--sy-naranjo); background:rgba(250,77,9,.04);}
    [data-testid="stDataFrame"], [data-testid="stTable"] {border-radius:12px; overflow:hidden;}
    [data-testid="stProgress"] [role="progressbar"] > div > div > div {background:linear-gradient(90deg,var(--sy-azul3),var(--sy-naranjo)) !important;}
    
    /* ---------- Barra lateral (siempre azul SYNCORE) ---------- */
    [data-testid="stSidebar"] {background:linear-gradient(180deg,#041640 0%,#06205C 60%,#0A2A6E 100%) !important;
      border-right:0; overflow:hidden;}
    [data-testid="stSidebar"]::before {content:""; position:absolute; width:220px; height:220px; border:2px solid rgba(250,77,9,.35);
      border-radius:30px; right:-120px; bottom:-70px; animation:sy-gira 40s linear infinite; pointer-events:none;}
    [data-testid="stSidebar"]::after {content:""; position:absolute; width:110px; height:110px; background:rgba(250,77,9,.85);
      border-radius:22px; right:-50px; bottom:-55px; transform:rotate(45deg); pointer-events:none;}
    [data-testid="stSidebarContent"] {display:flex; flex-direction:column; background:transparent !important;
      background-image:linear-gradient(rgba(255,255,255,.025) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,.025) 1px,transparent 1px) !important;
      background-size:40px 40px !important;}
    [data-testid="stSidebarHeader"] {order:0;}
    [data-testid="stSidebarUserContent"] {order:1; padding-bottom:.4rem !important;}
    [data-testid="stSidebarNav"] {order:2; padding-bottom:1rem;}
    [data-testid="stSidebarContent"]::after {content:"SYNCORE v__VERSION__ · Control Ejecutivo P6"; order:3; margin-top:auto;
      padding:14px 24px 18px; font-size:11.5px; color:#7F8FB5; letter-spacing:.2px; position:relative; z-index:1;}
    [data-testid="stSidebar"] :is(p, span, label, small, li, h1, h2, h3, div[data-testid="stCaptionContainer"]) {color:#D5DDEE !important;}
    [data-testid="stSidebar"] [data-baseweb="select"] > div {background:rgba(255,255,255,.07) !important; border-color:rgba(255,255,255,.18) !important;}
    [data-testid="stSidebar"] [data-baseweb="select"] svg {fill:#D5DDEE !important; color:#D5DDEE !important;}
    [data-testid="stSidebar"] [data-testid^="stBaseButton-secondary"] {background:rgba(255,255,255,.04) !important; border:1px solid rgba(255,255,255,.22) !important;
      border-radius:10px !important; transition:all .2s var(--sy-curva) !important;}
    [data-testid="stSidebar"] [data-testid^="stBaseButton-secondary"]:hover {background:rgba(250,77,9,.18) !important; border-color:var(--sy-naranjo) !important;}
    [data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] svg, [data-testid="stSidebarHeader"] svg {color:#D5DDEE !important; fill:#D5DDEE !important;}
    
    /* Menú de páginas */
    [data-testid="stNavSectionHeader"], [data-testid="stSidebarNavItems"] header {font-size:11px !important; letter-spacing:1.4px; font-weight:700 !important;
      text-transform:uppercase; margin-top:14px !important; padding-left:22px !important;}
    [data-testid="stNavSectionHeader"] *, [data-testid="stSidebarNavItems"] header * {color:#7F8FB5 !important;}
    [data-testid="stSidebarNavSeparator"] {display:none;}
    [data-testid="stSidebarNavLink"] {border-radius:10px !important; margin:2px 10px !important; padding:8px 12px !important; position:relative;
      transition:background .25s var(--sy-curva), transform .25s var(--sy-curva), box-shadow .25s !important; animation:sy-desliza .4s var(--sy-curva) both;}
    [data-testid="stSidebarNavLink"]::before {content:""; flex:0 0 auto; width:11px; height:11px; margin-right:12px; border-radius:3px;
      border:1.6px solid #7F8FB5; transform:rotate(45deg); transition:all .3s var(--sy-curva);}
    [data-testid="stSidebarNavLink"] :is([data-testid="stIconEmoji"], [data-testid="stIconMaterial"]) {display:none !important;}
    [data-testid="stSidebarNavLink"]:hover {background:rgba(255,255,255,.08) !important; transform:translateX(4px);}
    [data-testid="stSidebarNavLink"]:hover::before {border-color:var(--sy-naranjo2); transform:rotate(135deg);}
    [data-testid="stSidebarNavLink"][data-sy-activo="1"] {background:linear-gradient(90deg,rgba(250,77,9,.95),rgba(255,122,61,.85)) !important;
      box-shadow:0 8px 18px rgba(250,77,9,.35);}
    [data-testid="stSidebarNavLink"][data-sy-activo="1"]::before {border-color:#fff; background:#fff;}
    [data-testid="stSidebarNavLink"][data-sy-activo="1"] span {color:#FFFFFF !important; font-weight:600 !important;}
    [data-testid="stSidebarNavItems"] li:nth-child(2) [data-testid="stSidebarNavLink"] {animation-delay:.03s;}
    [data-testid="stSidebarNavItems"] li:nth-child(3) [data-testid="stSidebarNavLink"] {animation-delay:.06s;}
    [data-testid="stSidebarNavItems"] li:nth-child(4) [data-testid="stSidebarNavLink"] {animation-delay:.09s;}
    [data-testid="stSidebarNavItems"] li:nth-child(n+5) [data-testid="stSidebarNavLink"] {animation-delay:.12s;}
    
    /* Tarjeta de usuario */
    .sy-usuario {display:flex; gap:12px; align-items:center; background:rgba(255,255,255,.06); border:1px solid rgba(255,255,255,.09);
      border-radius:12px; padding:10px 12px; animation:sy-sube .5s var(--sy-curva) both;}
    .sy-usuario .sy-av {width:38px; height:38px; flex:0 0 38px; border-radius:10px; background:linear-gradient(135deg,var(--sy-naranjo),var(--sy-naranjo2));
      display:flex; align-items:center; justify-content:center; font-weight:800; color:#fff !important; font-size:14px;}
    .sy-usuario b {display:block; color:#fff !important; font-size:14px; line-height:1.2;}
    .sy-usuario small {color:#A9B6D3 !important; font-size:11.5px;}
    .sy-etq {font-size:11px; letter-spacing:1.4px; font-weight:700; color:#7F8FB5 !important; margin:14px 0 -6px 2px;}
    .sy-pie {color:#7F8FB5 !important; font-size:.72rem;}
    
    /* ---------- Logo: una versión para cada fondo ---------- */
    .sy-logo {display:block; width:100%; height:auto; margin:0 auto .5rem auto;}
    .sy-logo-noche {display:none;}
    html[data-sy-tema="oscuro"] .sy-logo-dia {display:none;}
    html[data-sy-tema="oscuro"] .sy-logo-noche {display:block;}
    [data-testid="stSidebar"] .sy-logo-dia {display:none !important;}
    [data-testid="stSidebar"] .sy-logo-noche {display:block !important; max-width:210px; margin:0 0 .3rem 0; animation:sy-aparece .6s ease;}
    .sy-marca {text-align:center; margin:0;} .sy-marca b {color:var(--sy-naranjo); font-size:2.2rem;}
    .sy-marca span {display:block; color:var(--sy-suave);}
    
    /* ---------- Pantalla de inicio de sesión ---------- */
    [data-testid="stMain"] .block-container:has(.sy-login-izq) {padding:0 !important; max-width:100% !important;}
    [data-testid="stApp"]:has(.sy-login-izq) [data-testid="stHeader"] {background:transparent !important;}
    [data-testid="stHorizontalBlock"]:has(.sy-login-izq) {gap:0 !important; min-height:100vh; align-items:stretch;}
    [data-testid="stColumn"]:has(.sy-login-izq) {padding:0 !important;}
    [data-testid="stColumn"]:has(.sy-login-izq) [data-testid="stVerticalBlock"], [data-testid="stColumn"]:has(.sy-login-izq) [data-testid="stElementContainer"],
    [data-testid="stColumn"]:has(.sy-login-izq) [data-testid="stMarkdownContainer"], [data-testid="stColumn"]:has(.sy-login-izq) [data-testid="stMarkdown"] {height:100%;}
    [data-testid="stColumn"]:has(.sy-login-der) {padding:9vh 6vw 4vh 6vw !important; background:var(--sy-tarjeta); position:relative;
      border-left:6px solid var(--sy-naranjo); animation:sy-sube .7s .15s var(--sy-curva) both;}
    [data-testid="stColumn"]:has(.sy-login-der) [data-testid="stForm"] {border:0; padding:0; background:transparent;}
    [data-testid="stColumn"]:has(.sy-login-der) [data-baseweb="input"] {height:46px; background:var(--sy-campo);}
    .sy-login-izq {position:relative; overflow:hidden; min-height:100vh; height:100%; padding:7vh 4.5vw; color:#fff;
      background:linear-gradient(140deg,#041640 0%,#06205C 45%,#0B2F7A 100%);}
    .sy-login-izq .sy-grilla {position:absolute; inset:0; background-image:linear-gradient(rgba(255,255,255,.035) 1px,transparent 1px),
      linear-gradient(90deg,rgba(255,255,255,.035) 1px,transparent 1px); background-size:48px 48px;}
    .sy-login-izq .sy-r {position:absolute; border-radius:28px; transform:rotate(45deg);}
    .sy-r1 {width:420px; height:420px; border:2px solid rgba(250,77,9,.55); right:-130px; top:-120px; animation:sy-entra-rombo 1.1s var(--sy-curva) both, sy-gira 60s 1.1s linear infinite;}
    .sy-r2 {width:260px; height:260px; background:rgba(250,77,9,.92); right:80px; bottom:-100px; border-radius:36px !important;
      animation:sy-entra-rombo 1s .25s var(--sy-curva) both, sy-flota 7s 1.3s ease-in-out infinite;}
    .sy-r3 {width:520px; height:520px; background:rgba(255,255,255,.04); left:-230px; bottom:-210px; animation:sy-entra-rombo 1.2s .1s var(--sy-curva) both;}
    .sy-r4 {width:140px; height:140px; border:2px solid rgba(255,255,255,.25); left:52%; top:36%; border-radius:20px !important;
      animation:sy-entra-rombo 1s .4s var(--sy-curva) both, sy-flota 6s 1.4s ease-in-out infinite;}
    .sy-r5 {width:58px; height:58px; background:#fff; opacity:.9; left:68%; top:27%; border-radius:12px !important;
      animation:sy-entra-rombo .9s .55s var(--sy-curva) both, sy-flota 5s 1.5s ease-in-out infinite;}
    .sy-login-izq img {position:relative; height:82px; width:auto; animation:sy-aparece .8s ease both;}
    .sy-login-izq h1 {position:relative; margin-top:30vh; font-size:2.9rem; font-weight:800; line-height:1.12; color:#fff !important;
      max-width:540px; padding:0 !important; animation:sy-sube .8s .35s var(--sy-curva) both;}
    .sy-login-izq h1::before {display:none;}
    .sy-login-izq h1 span {color:var(--sy-naranjo);}
    .sy-login-izq p {position:relative; margin-top:16px; color:#C9D3EA !important; font-size:1.05rem; max-width:480px; line-height:1.55;
      animation:sy-sube .8s .5s var(--sy-curva) both;}
    .sy-chips {position:relative; display:flex; flex-wrap:wrap; gap:10px; margin-top:24px; animation:sy-sube .8s .65s var(--sy-curva) both;}
    .sy-chips span {border:1px solid rgba(255,255,255,.25); padding:6px 14px; border-radius:30px; font-size:13px; color:#E6ECF8;
      transition:all .25s;}
    .sy-chips span:hover {background:rgba(250,77,9,.2); border-color:var(--sy-naranjo);}
    .sy-version {display:inline-flex; align-items:center; gap:9px; background:rgba(250,77,9,.1); color:var(--sy-naranjo) !important; font-weight:700;
      font-size:12.5px; padding:6px 13px; border-radius:20px; letter-spacing:.3px;}
    .sy-version i {width:8px; height:8px; border-radius:50%; background:var(--sy-naranjo); animation:sy-late 2s infinite;}
    .sy-bienv {font-size:2rem; font-weight:800; color:var(--sy-titulo); margin:16px 0 2px 0; padding:0 !important; letter-spacing:-.5px;}
    .sy-bienv-sub {color:var(--sy-suave); margin-bottom:10px;}
    .sy-login-pie {display:flex; justify-content:space-between; color:#98A2B6; font-size:12.5px; margin-top:22px;}
    @media (max-width: 900px) {.sy-login-izq {min-height:auto; padding:32px 24px;} .sy-login-izq h1 {margin-top:40px; font-size:2rem;}
      [data-testid="stColumn"]:has(.sy-login-der) {padding:32px 20px !important; border-left:0; border-top:6px solid var(--sy-naranjo);}}
    
    /* El iframe del detector no ocupa espacio */
    .element-container:has(> iframe[height="0"]), [data-testid="stElementContainer"]:has(iframe[height="0"]) {display:none;}
    </style>""".replace("__VERSION__", VERSION)
    
    
    def panel_login(logo_b64: str | None) -> str:
        """Panel azul de la pantalla de acceso, con las figuras animadas de la marca."""
        logo = f"<img alt='SYNCORE' src='data:image/png;base64,{logo_b64}'>" if logo_b64 else "<b style='font-size:2rem'>SYNCORE</b>"
        return ("<div class='sy-login-izq'><div class='sy-grilla'></div><div class='sy-r sy-r3'></div><div class='sy-r sy-r1'></div>"
                "<div class='sy-r sy-r4'></div><div class='sy-r sy-r5'></div><div class='sy-r sy-r2'></div>"
                f"{logo}<h1>Control Ejecutivo de <span>Proyectos</span></h1>"
                "<p>Programación P6, informes diarios, Times &amp; Material y desviaciones de recursos en una sola plataforma.</p>"
                "<div class='sy-chips'><span>Informe diario</span><span>6WLA</span><span>T&amp;M</span><span>Dashboard</span>"
                "<span>Desviaciones</span></div></div>")
    
    
    def logo_blanco(ruta_logo: str) -> str | None:
        logos = _logos(ruta_logo, os.path.getmtime(ruta_logo)) if os.path.exists(ruta_logo) else None
        return logos[1] if logos else None
    
    
    def tarjeta_usuario(nombre: str, rol: str) -> str:
        import html as _h
        partes = [p for p in str(nombre or "").split() if p]
        ini = "".join(p[0] for p in (partes[:1] + partes[-1:] if len(partes) > 1 else partes)).upper() or "?"
        return (f"<div class='sy-usuario'><div class='sy-av'>{_h.escape(ini)}</div>"
                f"<div><b>{_h.escape(str(nombre or ''))}</b><small>{_h.escape(str(rol or ''))}</small></div></div>")
    
    
    # Se inserta en el documento principal (no en el iframe) para que siga funcionando aunque el iframe se recree.
    _DETECTOR = """<script>
    (function () {
      var d = window.parent.document;
      if (d.getElementById("sy-tema-detector")) return;
      function sy() {
        function lum(c) { var m = (c || "").match(/[0-9.]+/g); if (!m) return 0; return (0.2126 * m[0] + 0.7152 * m[1] + 0.0722 * m[2]) / 255; }
        function tema() {
          var a = document.querySelector('[data-testid="stApp"]'); if (!a) return;
          var t = lum(getComputedStyle(a).color) > 0.5 ? "oscuro" : "claro";
          if (document.documentElement.getAttribute("data-sy-tema") !== t) document.documentElement.setAttribute("data-sy-tema", t);
        }
        function ruta(h) { return new URL(h, location.href).pathname.replace(/[/]+$/, ""); }
        function activo() {
          var ls = document.querySelectorAll('a[data-testid="stSidebarNavLink"]'); if (!ls.length) return;
          var p = ruta(location.href), hay = false;
          ls.forEach(function (a) {
            var on = ruta(a.href) === p; if (on) hay = true;
            if ((a.getAttribute("data-sy-activo") === "1") !== on) a.setAttribute("data-sy-activo", on ? "1" : "0");
          });
          if (!hay) ls[0].setAttribute("data-sy-activo", "1");
        }
        tema(); activo(); setInterval(function () { tema(); activo(); }, 300);
      }
      var s = d.createElement("script");
      s.id = "sy-tema-detector";
      s.textContent = "(" + sy.toString() + ")();";
      d.head.appendChild(s);
    })();
    </script>"""
    
    
    def aplicar():
        """Estilos de la marca + detector del tema activo. Llamar una vez por ejecución, antes de dibujar."""
        st.markdown(CSS, unsafe_allow_html=True)
        components.html(_DETECTOR, height=0)
    
    
    @st.cache_resource(show_spinner=False)
    def _logos(ruta: str, modificado: float) -> tuple[str, str] | None:
        """Devuelve (logo para fondo claro, logo para fondo oscuro) en base64, recortados al contenido.
        La versión nocturna se obtiene pasando a blanco el azul marino del logo; el naranjo se mantiene."""
        try:
            from PIL import Image
            im = Image.open(ruta).convert("RGBA")
        except Exception:
            return None
        caja = im.getchannel("A").getbbox()
        if caja:
            m = 24
            im = im.crop((max(0, caja[0] - m), max(0, caja[1] - m), min(im.width, caja[2] + m), min(im.height, caja[3] + m)))
        noche = im.copy()
        px = noche.load()
        for y in range(noche.height):
            for x in range(noche.width):
                r, g, b, a = px[x, y]
                if a and b > r:                      # tonos azules (el naranjo tiene r > b)
                    px[x, y] = (241, 245, 249, a)
    
        def b64(img):
            buf = io.BytesIO()
            img.save(buf, "PNG", optimize=True)
            return base64.b64encode(buf.getvalue()).decode()
    
        return b64(im), b64(noche)
    
    
    def marca(ruta_logo: str, ancho_max: int | None = None):
        """Logo de SYNCORE que cambia según el modo; si no hay archivo, la marca en texto."""
        logos = _logos(ruta_logo, os.path.getmtime(ruta_logo)) if os.path.exists(ruta_logo) else None
        if not logos:
            st.markdown("<p class='sy-marca'><b>SYNCORE</b><span>Control Ejecutivo P6</span></p>", unsafe_allow_html=True)
            return
        estilo = f" style='max-width:{ancho_max}px'" if ancho_max else ""
        dia, noche = logos
        st.markdown(
            f"<img class='sy-logo sy-logo-dia' alt='SYNCORE'{estilo} src='data:image/png;base64,{dia}'>"
            f"<img class='sy-logo sy-logo-noche' alt='SYNCORE'{estilo} src='data:image/png;base64,{noche}'>",
            unsafe_allow_html=True)
    
    
    AYUDA = "🌗 Modo día / noche: menú ⋮ (arriba a la derecha) → Settings → Choose app theme."

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


tema = _modulo_tema()

# ======================================================================
# MÓDULO: dashboard
# ======================================================================
def _modulo_dashboard():
    import pandas as pd
    import plotly.express as px
    import plotly.graph_objects as go
    import streamlit as st
    
    CAL = calendario
    c = comun
    
    
    def vista():
        st.title("Dashboard ejecutivo")
        sel = c.selector_proyecto_corte("dash")
        if not sel:
            return
        proyecto, corte, cortes = sel
        act = c.actividades(int(corte["id"]))
        if act.empty:
            st.warning("Este corte no contiene actividades.")
            return
    
        previo = cortes[cortes["fecha_corte"] < corte["fecha_corte"]]
        previo = previo.iloc[-1] if not previo.empty else None
    
        st.caption(
            f"Data Date **{c.fecha(corte['fecha_corte'])}** (**{CAL.texto(corte['fecha_corte'])}**) · Inicio **{c.fecha(corte['fecha_inicio'])}** · "
            f"Término programado **{c.fecha(corte['fecha_termino'])}** · Archivo `{corte['archivo'] or '–'}`"
        )
    
        k = st.columns(5)
        k[0].metric("Actividades", c.num(len(act)))
        k[1].metric("At Completion", f"{c.num(corte['at_completion'])} HH")
        k[2].metric("Actual", f"{c.num(corte['actual'])} HH",
                    delta=f"{c.num(corte['actual'] - previo['actual'])} HH vs corte anterior" if previo is not None else None)
        k[3].metric("Remaining", f"{c.num(corte['remaining'])} HH")
        k[4].metric("Labor Units % Complete", c.porc(corte["avance_pct"]),
                    delta=f"{c.num(corte['avance_pct'] - previo['avance_pct'], 2)} pp" if previo is not None else None)
    
        if abs(corte["at_completion"] - corte["budget"]) > 0.5:
            desv = corte["at_completion"] - corte["budget"]
            st.caption(f"Budget Labor Units: {c.num(corte['budget'])} HH · desviación At Completion vs Budget: "
                       f"{'+' if desv > 0 else ''}{c.num(desv)} HH ({c.porc(analisis.pct(desv, corte['budget']), 1)})")
    
        st.divider()
        g1, g2 = st.columns(2)
        with g1:
            st.subheader("Estado de actividades")
            est = act["estado"].value_counts().rename_axis("Estado").reset_index(name="Actividades")
            fig = px.pie(est, names="Estado", values="Actividades", hole=0.5, color="Estado",
                         color_discrete_map=c.COLORES_ESTADO)
            fig.update_traces(textinfo="value+percent")
            st.plotly_chart(c.layout_grafico(fig, 340), use_container_width=True, key="g_dashboard_1")
        with g2:
            if len(cortes) > 1:
                st.subheader("Evolución del avance por corte")
                ev = cortes.copy()
                ev["Fecha"] = pd.to_datetime(ev["fecha_corte"])
                fig = go.Figure()
                fig.add_scatter(x=ev["Fecha"], y=ev["avance_pct"], mode="lines+markers+text",
                                text=[c.porc(v, 1) for v in ev["avance_pct"]], textposition="top center",
                                line=dict(color=c.NARANJO, width=3), name="% avance real")
                fig.update_yaxes(title="% Labor Units", range=[0, 105])
                st.plotly_chart(c.layout_grafico(fig, 340), use_container_width=True, key="g_dashboard_2")
            else:
                st.subheader("Horas hombre")
                df = pd.DataFrame({"Categoría": ["Actual", "Remaining"], "HH": [corte["actual"], corte["remaining"]]})
                fig = px.bar(df, x="HH", y="Categoría", orientation="h", text=[c.num(v) for v in df["HH"]],
                             color="Categoría", color_discrete_sequence=[c.NARANJO, c.PIZARRA])
                fig.update_layout(showlegend=False)
                st.plotly_chart(c.layout_grafico(fig, 340), use_container_width=True, key="g_dashboard_3")
                st.caption("Carga un .XER de otra fecha de corte para ver la evolución del avance.")
    
        st.divider()
        tipo = c.selector_tipo_codigo(int(corte["id"]), "dash_tipo")
        if tipo:
            r = analisis.resumen_por_codigo(act, c.codigos(int(corte["id"]), tipo))
            st.subheader(f"Resumen por {tipo}")
            st.dataframe(
                r[["valor", "actividades", "completadas", "at_completion", "actual", "remaining", "avance_pct"]],
                hide_index=True, use_container_width=True, height=min(600, 38 + 35 * len(r)),
                column_config={"valor": st.column_config.TextColumn(tipo, width="large"),
                               "actividades": "Actividades", "completadas": "Completadas", **c.COLUMNAS_HORAS},
            )
    
        proximas = act[(act["estado"] != "Completado") & act["holgura_total_hr"].notna()].nsmallest(10, "holgura_total_hr")
        if not proximas.empty:
            st.subheader("Actividades pendientes con menor holgura total")
            proximas = proximas.assign(holgura_dias=proximas["holgura_total_hr"] / 8)
            st.dataframe(
                proximas[["task_code", "nombre", "estado", "inicio", "fin", "holgura_dias", "remaining"]],
                hide_index=True, use_container_width=True,
                column_config={"task_code": "Activity ID", "nombre": "Actividad", "estado": "Estado",
                               "inicio": "Inicio", "fin": "Término",
                               "holgura_dias": st.column_config.NumberColumn("Holgura total (días de 8 h)", format="%.1f"),
                               "remaining": c.COLUMNAS_HORAS["remaining"]},
            )

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


dashboard = _modulo_dashboard()

# ======================================================================
# MÓDULO: subsistemas
# ======================================================================
def _modulo_subsistemas():
    import os
    
    import plotly.express as px
    import streamlit as st
    
    c = comun
    
    
    def vista():
        st.title("Avance por subsistema")
        sel = c.selector_proyecto_corte("sub")
        if not sel:
            return
        proyecto, corte, cortes = sel
        cid = int(corte["id"])
        act = c.actividades(cid)
        tipo = c.selector_tipo_codigo(cid, "sub_tipo")
        if not tipo or act.empty:
            return
        cods = c.codigos(cid, tipo)
        r = analisis.resumen_por_codigo(act, cods)
    
        con_horas = r[r["at_completion"] > 0]
        k = st.columns(4)
        k[0].metric("Valores de código", len(r))
        k[1].metric("At Completion", f"{c.num(r['at_completion'].sum())} HH")
        k[2].metric("Actual", f"{c.num(r['actual'].sum())} HH")
        k[3].metric("Avance ponderado", c.porc(analisis.pct(r["actual"].sum(), r["at_completion"].sum())))
    
        t1, t2, t3, t4 = st.tabs(["Matriz de avance", "Gráfico", "Actividades por valor", "📄 Informe ejecutivo"])
        with t4:
            _informe_ejecutivo(proyecto, corte, cortes, act, cods, tipo, r)
        with t1:
            st.dataframe(
                r[["valor", "actividades", "completadas", "inicio", "fin", "budget", "at_completion", "actual", "remaining", "avance_pct"]],
                hide_index=True, use_container_width=True, height=min(700, 38 + 35 * len(r)),
                column_config={"valor": st.column_config.TextColumn(tipo, width="large"), "actividades": "Actividades",
                               "completadas": "Completadas", "inicio": "Inicio", "fin": "Término", **c.COLUMNAS_HORAS},
            )
            st.download_button(
                "Descargar Excel", key="dl_sub",
                data=analisis.a_excel({"Resumen": r, "Actividades": act.merge(cods, on="task_code", how="left")}),
                file_name=f"{proyecto['codigo']}_{corte['fecha_corte']}_{tipo[:20]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with t2:
            solo_horas = st.checkbox("Ocultar valores sin horas", value=True)
            datos = (con_horas if solo_horas else r).sort_values("avance_pct")
            datos = datos.assign(etiqueta=[c.porc(v, 1) for v in datos["avance_pct"]])
            fig = px.bar(datos, x="avance_pct", y="valor", orientation="h", text="etiqueta",
                         color="avance_pct", color_continuous_scale=[c.PIZARRA, c.AZUL, c.NARANJO], range_color=[0, 100],
                         hover_data={"at_completion": ":,.0f", "actual": ":,.0f", "remaining": ":,.0f", "etiqueta": False},
                         labels={"avance_pct": "% avance (Labor Units)", "valor": "", "at_completion": "At Completion",
                                 "actual": "Actual", "remaining": "Remaining"})
            fig.update_layout(coloraxis_showscale=False, yaxis=dict(tickfont=dict(size=11)))
            fig.update_xaxes(range=[0, 110])
            st.plotly_chart(c.layout_grafico(fig, max(400, 24 * len(datos))), use_container_width=True, key="g_subsistemas_1")
        with t3:
            valor = st.selectbox("Valor", r["valor"].tolist(), key="sub_valor")
            if valor == analisis.SIN_CODIGO:
                det = act[~act["task_code"].isin(set(cods["task_code"]))]
            else:
                det = act[act["task_code"].isin(set(cods.loc[cods["valor"] == valor, "task_code"]))]
            estados = st.multiselect("Estado", sorted(det["estado"].unique()), default=sorted(det["estado"].unique()))
            det = det[det["estado"].isin(estados)]
            st.caption(f"{len(det)} actividades")
            st.dataframe(
                det[["task_code", "nombre", "tipo", "estado", "inicio", "fin", "at_completion", "actual", "remaining", "avance_pct"]],
                hide_index=True, use_container_width=True,
                column_config={"task_code": "Activity ID", "nombre": st.column_config.TextColumn("Actividad", width="large"),
                               "tipo": "Tipo", "estado": "Estado", "inicio": "Inicio", "fin": "Término", **c.COLUMNAS_HORAS},
            )
    
    
    def _restricciones():
        return _restricciones_de(c.contrato_id())
    
    
    @st.cache_data(ttl=600, show_spinner=False)
    def _restricciones_de(cid: int):
        return db.restricciones(cid)
    
    
    def _informe_ejecutivo(proyecto, corte, cortes, act, cods, tipo, resumen):
        st.markdown("Genera un **PDF ejecutivo** del ítem elegido: KPIs, comparación con el corte anterior, resumen "
                    "redactado a partir de los datos, desglose por especialidad, actividades clave y restricciones 6WLA.")
        cid = int(corte["id"])
        opciones = resumen.sort_values("at_completion", ascending=False)["valor"].tolist()
        a, b = st.columns([3, 2])
        valor = a.selectbox(f"Ítem ({tipo})", opciones, key="ie_valor")
        otros = [t for t in c.tipos_codigo(cid) if t != tipo]
        sug = next((t for t in otros if any(k in t.lower() for k in ("especialidad", "disciplina"))), None)
        ops_esp = ["(sin desglose)"] + otros
        tipo_esp = b.selectbox("Desglosar por", ops_esp, index=ops_esp.index(sug) if sug else 0, key="ie_esp")
        tipo_esp = None if tipo_esp == "(sin desglose)" else tipo_esp
    
        previos = cortes[cortes["fecha_corte"] < corte["fecha_corte"]]
        act_prev = cods_prev = fecha_prev = None
        if not previos.empty:
            p = previos.iloc[-1]
            act_prev, cods_prev, fecha_prev = c.actividades(int(p["id"])), c.codigos(int(p["id"]), tipo), p["fecha_corte"]
    
        d = analisis.datos_informe_item(
            act.copy(), cods, valor, tipo, {"codigo": proyecto["codigo"], "nombre": proyecto["nombre"]},
            {"fecha_corte": corte["fecha_corte"]}, act_prev, cods_prev, fecha_prev,
            c.codigos(cid, tipo_esp) if tipo_esp else None, tipo_esp, _restricciones())
    
        k = st.columns(4)
        k[0].metric("% Avance", c.porc(d["avance_pct"]),
                    delta=None if d["delta_pp"] is None else f"{c.num(d['delta_pp'], 2)} pp vs {c.fecha(fecha_prev)}")
        k[1].metric("At Completion", f"{c.num(d['at_completion'])} HH")
        k[2].metric("Remaining", f"{c.num(d['remaining'])} HH")
        k[3].metric("Actividades", f"{d['actividades']} ({d['completadas']} compl.)")
    
        st.markdown("**Resumen ejecutivo (automático)**")
        for frase in d["resumen"]:
            st.markdown(f"- {frase}")
    
        obs = st.text_area("Análisis del planificador (opcional)", key=f"ie_obs_{valor}", height=100,
                           placeholder="Ej.: El rendimiento de montaje se mantuvo; el atraso del quemador afecta la puesta en marcha.")
        acc = st.text_area("Acciones y compromisos (opcional)", key=f"ie_acc_{valor}", height=100,
                           placeholder="Ej.: 1. Abastecimiento confirma despacho antes del 05-11.")
        autor = (c.usuario() or {}).get("nombre", "")
        pdf = informe_pdf.generar_pdf(d, autor=autor, observaciones=obs, acciones=acc,
                                      logo="logo_syncore.png" if os.path.exists("logo_syncore.png") else None)
        corto = analisis._clave_codigo(valor).replace("/", "-").replace(" ", "_")[:40]
        st.download_button("📄 Descargar informe ejecutivo (PDF)", pdf, type="primary", key="ie_dl",
                           file_name=f"Informe_{corto}_{corte['fecha_corte']}.pdf", mime="application/pdf")
        if d["delta_pp"] is None:
            st.caption("Para ver la variación semanal, carga también el .XER del corte anterior.")

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


subsistemas = _modulo_subsistemas()

# ======================================================================
# MÓDULO: importar
# ======================================================================
def _modulo_importar():
    import datetime as dt
    
    import pandas as pd
    import streamlit as st
    
    c = comun
    
    
    def vista():
        st.title("Importar programa")
        if not c.exigir("cargar"):
            return
        t1, t2 = st.tabs(["Primavera P6 (.xer)", "Planilla Excel de actividades"])
        with t1:
            _xer()
        with t2:
            _excel()
    
    
    @st.cache_data(show_spinner="Leyendo .XER…", max_entries=3)
    def _leer_xer(datos: bytes):
        return xer.leer_tablas(xer.decodificar(datos))
    
    
    def _xer():
        st.markdown(
            "Exporta desde P6 con **File → Export → Primavera PM (XER)**. Cada importación se guarda como un "
            "**corte** según su Data Date: así se construye el historial de avance. Si subes otra vez un archivo con la "
            "misma Data Date, ese corte se reemplaza. Los demás proyectos y cortes no se tocan."
        )
        archivo = st.file_uploader("Archivo .xer", type=["xer"], key="up_xer")
        if archivo is None:
            return
        try:
            tablas = _leer_xer(archivo.getvalue())
        except xer.XerError as e:
            st.error(str(e))
            return
    
        disponibles = xer.listar_proyectos(tablas)
        idx = 0
        if len(disponibles) > 1:
            idx = st.selectbox("El archivo contiene varios proyectos. Elige cuál importar:", range(len(disponibles)),
                               format_func=lambda i: disponibles[i]["codigo"])
        try:
            proy = xer.extraer_proyecto(tablas, disponibles[idx]["proj_id"])
        except xer.XerError as e:
            st.error(str(e))
            return
    
        st.success(f"Proyecto **{proy.codigo}** — {proy.nombre}")
        k = st.columns(4)
        k[0].metric("Data Date", c.fecha(proy.fecha_corte))
        k[1].metric("Actividades", c.num(len(proy.actividades)))
        k[2].metric("At Completion", f"{c.num(proy.at_completion)} HH")
        k[3].metric("Labor Units % Complete", c.porc(proy.avance_pct))
        st.caption(f"Inicio {c.fecha(proy.fecha_inicio)} · Término programado {c.fecha(proy.fecha_termino)} · "
                   f"Horas leídas desde {proy.fuente_horas}")
    
        if proy.tipos_codigo:
            conteo = pd.DataFrame(proy.codigos, columns=["task_code", "Código de actividad", "valor"])
            resumen = conteo.groupby("Código de actividad").agg(
                Valores=("valor", "nunique"), Actividades=("task_code", "nunique")).reset_index()
            with st.expander(f"Códigos de actividad encontrados ({len(resumen)})"):
                st.dataframe(resumen, hide_index=True, use_container_width=True)
        else:
            st.warning("El archivo no trae códigos de actividad (ACTVCODE/TASKACTV). Solo se podrá ver el avance total.")
    
        if "\ufffd" in " ".join(a["nombre"] for a in proy.actividades[:500]) + " ".join(v for _, _, v in proy.codigos[:2000]):
            st.info("Algunos textos traen caracteres dañados (\ufffd) desde el propio archivo de P6. Para conservar tildes y "
                    "ñ, revisa la configuración regional del equipo que exporta el .XER.")
    
        if not proy.fecha_corte:
            st.warning("El archivo no indica Data Date; se usará la fecha de hoy como fecha de corte.")
    
        existe = db.uno(
            "SELECT c.id FROM cortes c JOIN proyectos p ON p.id = c.proyecto_id WHERE p.codigo = ? AND c.fecha_corte = ?",
            (proy.codigo, proy.fecha_corte or dt.date.today().isoformat()),
        )
        if existe:
            st.warning(f"Ya existe un corte de {proy.codigo} al {c.fecha(proy.fecha_corte)}. Al importar se reemplazará.")
    
        if st.button("Importar a la plataforma", type="primary", key="btn_xer"):
            try:
                with st.spinner("Guardando…"):
                    _, reemplazado = db.guardar_corte_p6(proy, archivo.name, c.nombre_usuario(), contrato_id=c.contrato_id())
            except ValueError as e:
                st.error(str(e))
                return
            c.limpiar_cache()
            st.success(f"Corte al {c.fecha(proy.fecha_corte)} {'reemplazado' if reemplazado else 'importado'} correctamente. "
                       "Revisa el Dashboard.")
    
    
    def _excel():
        st.markdown("Para proyectos que no vienen de P6. Indica qué columna corresponde a cada dato; "
                    "no se inventa ningún valor faltante.")
        archivo = st.file_uploader("Planilla Excel", type=["xlsx", "xlsm", "xls"], key="up_excel_act")
        if archivo is None:
            return
        datos = archivo.getvalue()
        hojas = pd.ExcelFile(archivo).sheet_names
        hoja = st.selectbox("Hoja", hojas, key="xl_hoja")
        df, fila = analisis.leer_hoja(datos, hoja)
        fila = st.number_input("Fila del encabezado (1 = primera fila)", 1, 100, fila + 1, key="xl_fila") - 1
        df, _ = analisis.leer_hoja(datos, hoja, fila)
        st.dataframe(df.head(8), use_container_width=True)
        cols = df.columns.tolist()
        NADA = "(no aplica)"
    
        def elegir(etiqueta, clave, opcional=False, col=st):
            sug = analisis.sugerir_columna(cols, clave) if clave else None
            ops = ([NADA] if opcional else []) + cols
            return col.selectbox(etiqueta, ops, index=ops.index(sug) if sug in ops else 0, key=f"xl_{etiqueta}")
    
        a, b = st.columns(2)
        col_act = a.selectbox("Actividad", cols, key="xl_act")
        col_sub = elegir("Subsistema / agrupación", "item", True, b)
        col_atc = elegir("At Completion (HH)", "at_completion", False, a)
        col_rem = elegir("Remaining (HH)", "remaining", True, b)
        col_av = elegir("% Avance", "avance", True, a)
        codigo = b.text_input("Código del proyecto", archivo.name.rsplit(".", 1)[0])
        nombre = a.text_input("Nombre del proyecto", codigo)
        fecha_corte = b.date_input("Fecha de corte", dt.date.today(), format="DD-MM-YYYY")
    
        if col_rem == NADA and col_av == NADA:
            st.info("Indica la columna de Remaining o la de % Avance.")
            return
        items, avisos = analisis.normalizar_informe(
            df, col_act, col_atc, col_rem=None if col_rem == NADA else col_rem,
            col_avance=None if col_av == NADA else col_av)
        for aviso in avisos:
            st.caption("⚠️ " + aviso)
        items = items[items["incluir"]].reset_index(drop=True)
        st.metric("Avance ponderado", c.porc(analisis.avance_ponderado(items)))
    
        if st.button("Importar planilla", type="primary", key="btn_xl"):
            proy = xer.ProyectoXer(proj_id="excel", codigo=codigo.strip(), nombre=nombre.strip() or codigo,
                                   fecha_corte=fecha_corte.isoformat(), fecha_inicio=None, fecha_termino=None)
            for i, r in enumerate(items.itertuples(), start=1):
                tc = f"XL-{i:05d}"
                proy.actividades.append({"task_code": tc, "nombre": r.item, "tipo": "Tarea",
                                         "estado": "Completado" if r.avance_pct >= 99.999 else ("No Iniciado" if r.actual <= 0 else "En Progreso"),
                                         "wbs": "", "budget": r.at_completion, "actual": r.actual, "remaining": r.remaining,
                                         "inicio": None, "fin": None, "holgura_total_hr": None})
            if col_sub != NADA:
                for a, fila_origen in zip(proy.actividades, items["fila_origen"]):
                    v = df.at[fila_origen, col_sub]
                    v = str(v).strip() if pd.notna(v) and str(v).strip() else analisis.SIN_CODIGO
                    proy.codigos.append((a["task_code"], "Subsistema (Excel)", v))
            try:
                db.guardar_corte_p6(proy, archivo.name, c.nombre_usuario(), contrato_id=c.contrato_id())
            except ValueError as e:
                st.error(str(e))
                return
            db.ejecutar("UPDATE proyectos SET origen = 'Excel' WHERE codigo = ?", (proy.codigo,))
            c.limpiar_cache()
            st.success("Planilla importada.")

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


importar = _modulo_importar()

# ======================================================================
# MÓDULO: vcubicar
# ======================================================================
def _modulo_vcubicar():
    import datetime as dt
    from dataclasses import asdict
    
    import pandas as pd
    import streamlit as st
    
    C = cubicacion
    P = cubicar_plano
    LP = lector_planos
    c = comun
    
    ICONOS = {"Obras civiles": "🧱", "Estructura": "🏗️", "Mecánica": "⚙️", "Piping": "🔧", "Eléctricos": "⚡"}
    _WIDGETS_META = ("cp_disc", "cp_plano", "cp_rev", "cp_area", "cp_desc")
    
    
    def seccion():
        cid = c.contrato_id()
        _aplicar_pendiente()
        catalogo = _catalogo(cid)
        st.markdown("Carga el **plano o isométrico** (PDF). En isométricos de piping la app **lee las tablas y arma la "
                    "cubicación sola**; en las demás disciplinas digitas lo que ves en el plano. En ambos casos calcula "
                    "cantidad, ítem EC-002, rendimiento y HH, sin depender de ninguna planilla existente.")
    
        # 1. plano
        st.subheader("1 · Plano")
        up = st.file_uploader("Plano o isométrico (PDF)", type=["pdf"], key="up_cp_pdf")
        if up is not None and st.session_state.get("cp_pdf_nombre") != up.name:
            st.session_state["cp_pdf_nombre"] = up.name
            st.session_state["cp_pdf"] = up.getvalue()
            m = P.meta_de_nombre(up.name)
            st.session_state["cp_pendiente"] = {"meta": {"plano": m["plano"], "rev": m["rev"], "area": m["area"],
                                                         "disciplina": m["disciplina"] or st.session_state.get("cp_disc")}}
            st.rerun()
        a, b, d, e = st.columns([1.1, 2.2, 0.6, 0.7])
        disc = a.selectbox("Disciplina", P.DISCIPLINAS, format_func=lambda x: f"{ICONOS[x]} {x}", key="cp_disc")
        plano = b.text_input("N° de plano", key="cp_plano", placeholder="EAL2LDHR-DD-2962-PIP-ISO-1001")
        rev = d.text_input("Rev.", key="cp_rev")
        area = e.text_input("Área", key="cp_area")
        desc = st.text_input("Descripción / línea / ubicación", key="cp_desc",
                             placeholder='2962-RA-S13-0001-24"-N · Fundaciones calderas · Pipe rack…',
                             help="Ayuda a proponer el ítem EC-002 correcto (los ítems del contrato van por ubicación).")
        if st.session_state.get("cp_pdf"):
            _lectura_automatica(cid, disc)
        if not catalogo:
            st.warning("No hay **itemizado EC-002** cargado: la app calcula cantidades pero no puede proponer ítems ni HH. "
                       "Carga cualquier planilla de avance en **📥 Cargar planillas**, o sube aquí un Excel con la hoja "
                       "de itemizado.")
        with st.expander(f"Itemizado EC-002 del contrato ({len(catalogo)} ítems)" if catalogo else
                         "Cargar itemizado EC-002", expanded=not catalogo):
            up_cat = st.file_uploader("Excel con la hoja de itemizado (planilla de avance o informe semanal)",
                                      type=["xlsx", "xlsm"], key="up_cp_cat")
            if up_cat is not None and st.button("Usar este itemizado", key="btn_cp_cat"):
                try:
                    nuevo = P.catalogo_de_planilla(up_cat.getvalue())
                    db.guardar_config(cid, "catalogo_ec002", P.unir_catalogos(catalogo, nuevo))
                    st.success(f"Itemizado cargado: {len(nuevo)} ítems.")
                    catalogo = _catalogo(cid)
                except Exception as ex:
                    st.error(f"{up_cat.name}: {ex}")
            if catalogo:
                st.dataframe(pd.DataFrame(catalogo)[["item", "descripcion", "unidad", "rend", "disciplina"]],
                             hide_index=True, use_container_width=True, height=250,
                             column_config={"item": "Ítem", "descripcion": "Descripción", "unidad": "Unidad",
                                            "rend": st.column_config.NumberColumn("Rend. HH/u", format="%.4f"),
                                            "disciplina": "Disciplina"})
    
        # 2. planilla
        st.subheader("2 · Cubicación")
        pdf = st.session_state.get("cp_pdf")
        lado = False
        if pdf:
            lado = st.toggle("Plano al lado de la planilla", value=False, key="cp_lado",
                             help="Apagado: el plano arriba y la planilla abajo, a todo el ancho.")
        if pdf and lado:
            izq, der = st.columns([1, 1.25])
            with izq:
                _visor(pdf)
            with der:
                filas = _editor(disc, catalogo)
        else:
            if pdf:
                with st.expander(f"📄 {st.session_state.get('cp_pdf_nombre', 'Plano')}", expanded=True):
                    _visor(pdf)
            filas = _editor(disc, catalogo)
    
        calc = P.calcular(disc, filas, catalogo, contexto=f"{desc} {plano}")
        _resultado(cid, disc, filas, calc, {"plano": plano, "rev": rev, "area": area, "descripcion": desc}, catalogo)
        st.divider()
        _guardadas(cid)
    
    
    # ------------------------------------------------------------------ lectura automática
    
    def _lectura_automatica(cid, disc):
        pdf = st.session_state["cp_pdf"]
        with st.container(border=True):
            if disc != "Piping":
                st.caption("🤖 La lectura automática está disponible para **isométricos de piping** (tablas CUT PIPE "
                           "LENGTHS, BILL OF MATERIALS, VALVES y PIPE SUPPORTS). Para esta disciplina digita los "
                           "elementos en la planilla de abajo.")
                return
            st.markdown("**🤖 Cubicación automática del isométrico**")
            a, b = st.columns([2, 1])
            a.caption("La app lee las tablas de todas las láminas (cortes, válvulas, soportes y materiales), arma la "
                      "planilla y revisa que los cortes de cada diámetro sumen lo que dice el BOM. Sin internet ni OCR: "
                      "reconoce las letras de la fuente del plano.")
            peso = b.number_input("Peso por soporte (kg)", min_value=0.0, step=1.0, key="cp_peso_sop",
                                  value=float(db.config(cid, "peso_soporte_kg", 113.0) or 0),
                                  help="Los soportes se pagan por kg. La planilla de piping usa 113 kg para S608CB6-24\".")
            if st.button("🤖 Leer isométrico y cubicar", type="primary", key="btn_cp_auto"):
                barra = st.progress(0.0, text="Leyendo láminas…")
                try:
                    lec = LP.leer_isometrico(pdf, lambda i, n: barra.progress(i / max(n, 1), text=f"Lámina {i + 1} de {n}…"))
                except Exception as ex:
                    barra.empty()
                    st.error(f"No se pudo leer el plano: {ex}")
                    return
                barra.empty()
                filas = LP.a_planilla(lec, peso or None)
                if peso:
                    db.guardar_config(cid, "peso_soporte_kg", float(peso))
                st.session_state["cp_lectura"] = {
                    "paginas": lec.paginas, "filas": len(filas), "avisos": lec.avisos,
                    "dudosas": sum(1 for f in filas if f.get("Observación")),
                    "cortes": sum(1 for f in filas if f["Tipo"] == "Cañería"),
                    "metros": sum((f.get("Largo (mm)") or 0) / 1000 for f in filas if f["Tipo"] == "Cañería"),
                    "valvulas": sum(f.get("Cantidad") or 0 for f in filas if f["Tipo"] == "Válvula"),
                    "soportes": sum(f.get("Cantidad") or 0 for f in filas if f["Tipo"] == "Soporte")}
                if not filas:
                    st.warning("No encontré tablas de isométrico en este PDF. Digita los elementos en la planilla.")
                    return
                st.session_state["cp_pendiente"] = {"meta": {"disciplina": "Piping"}, "filas": filas}
                st.rerun()
            r = st.session_state.get("cp_lectura")
            if r:
                msg = (f"Leí {r['paginas']} lámina(s): {r['cortes']} cortes de cañería ({c.num(r['metros'], 3)} m), "
                       f"{c.num(r['valvulas'])} válvulas, {c.num(r['soportes'])} soportes y el resto de los materiales. "
                       f"La planilla de abajo quedó llena: revísala y guarda.")
                (st.warning if r["avisos"] or r["dudosas"] else st.success)(msg)
                for av in r["avisos"]:
                    st.warning("⚠️ " + av)
                if r["dudosas"]:
                    st.info(f"{r['dudosas']} fila(s) tienen letras que no reconocí con seguridad: están marcadas con "
                            f"«revisar lectura» en la columna Observación.")
                elif not r["avisos"]:
                    st.caption("✅ En todas las láminas los cortes de cada diámetro cuadran con la cañería del BOM.")
    
    
    # ------------------------------------------------------------------ estado
    
    def _aplicar_pendiente():
        """Valores para los widgets (al abrir una cubicación o subir un PDF) antes de dibujarlos."""
        pend = st.session_state.pop("cp_pendiente", None)
        if not pend:
            return
        meta = pend.get("meta", {})
        for k, campo in zip(_WIDGETS_META, ("disciplina", "plano", "rev", "area", "descripcion")):
            if campo in meta and meta[campo] is not None:
                if k == "cp_disc" and meta[campo] not in P.DISCIPLINAS:
                    continue
                st.session_state[k] = meta[campo]
        if "filas" in pend:
            disc = meta.get("disciplina") or st.session_state.get("cp_disc") or P.DISCIPLINAS[0]
            st.session_state[f"cp_base_{disc}"] = _df(disc, pend["filas"])
            st.session_state["cp_ver"] = st.session_state.get("cp_ver", 0) + 1
    
    
    def _catalogo(cid) -> list[dict]:
        cat = db.config(cid, "catalogo_ec002", None)
        if cat:
            return cat
        # sin catálogo propio: lo que haya en las planillas de avance guardadas (ítems controlados)
        cargas = db.cargas_cubicacion(cid)
        items = []
        for r in cargas.itertuples():
            d = db.carga_cubicacion(int(r.id)) or {}
            items += [{"item": i["item"], "descripcion": i["descripcion"], "unidad": i["unidad"], "rend": i["rend"],
                       "contrato": i["contrato"], "disciplina": i["disciplina"]} for i in d.get("items", [])]
        return P.unir_catalogos(items)
    
    
    def _df(disc: str, filas: list[dict] | None = None) -> pd.DataFrame:
        cols = P.columnas(disc)
        filas = filas or [P.fila_vacia(disc) for _ in range(8)]
        df = pd.DataFrame([{n: f.get(n) for n, _, _ in cols} for f in filas], columns=[n for n, _, _ in cols])
        for n, t, _ in cols:
            df[n] = pd.to_numeric(df[n], errors="coerce") if t == "num" else df[n].fillna("").astype(str)
        return df
    
    
    def _opciones_item(disc: str, catalogo: list[dict]) -> list[str]:
        sel = [x for x in catalogo if x["disciplina"] == disc or (disc == "Piping" and x["item"][:2] in ("51", "52", "53"))]
        return [""] + [f"{x['item']} · {x['descripcion'][:70]}" for x in sel]
    
    
    def _editor(disc: str, catalogo: list[dict]) -> list[dict]:
        base_k = f"cp_base_{disc}"
        if base_k not in st.session_state:
            st.session_state[base_k] = _df(disc)
        base = st.session_state[base_k]
        opciones = _opciones_item(disc, catalogo)
        for v in base["Ítem EC-002"].unique():            # ítems guardados que ya no están en el catálogo
            if v and v not in opciones:
                opciones.append(v)
        conf = {}
        for n, t, ayuda in P.columnas(disc):
            if t == "num":
                conf[n] = st.column_config.NumberColumn(n, help=ayuda or None, min_value=0.0, format="%.3f")
            elif t == "tipo":
                conf[n] = st.column_config.SelectboxColumn(n, help=ayuda or None, options=list(P.TIPOS[disc]),
                                                           required=False, width="medium")
            elif t == "item":
                conf[n] = st.column_config.SelectboxColumn(n, help=ayuda or None, options=opciones, width="large")
            else:
                conf[n] = st.column_config.TextColumn(n, help=ayuda or None)
        st.caption("Agrega filas con el **+** al final de la tabla. Puedes pegar desde Excel (Ctrl+V) varias filas a la vez. "
                   + {"Piping": "En cañería, una fila por corte (tabla CUT PIPE LENGTHS, en mm).",
                      "Obras civiles": "Dimensiones en metros: la cantidad sale en m³, m² o kg según el tipo.",
                      "Estructura": "El peso sale de kg c/u, o de largo × kg/m.",
                      "Mecánica": "Una fila por equipo o TAG.",
                      "Eléctricos": "Cables, escalerillas y conduit en metros; equipos y luminarias por unidad."}[disc])
        ed = st.data_editor(base, key=f"cp_ed_{disc}_{st.session_state.get('cp_ver', 0)}", num_rows="dynamic",
                            hide_index=True, use_container_width=True, column_config=conf,
                            height=min(560, 38 + 35 * (len(base) + 2)))
        a, b, _ = st.columns([1, 1, 3])
        if a.button("➕ 10 filas más", key="cp_mas"):
            st.session_state[base_k] = pd.concat([ed, _df(disc, [P.fila_vacia(disc) for _ in range(10)])],
                                                 ignore_index=True)
            st.session_state["cp_ver"] = st.session_state.get("cp_ver", 0) + 1
            st.rerun()
        if b.button("🧹 Empezar de nuevo", key="cp_limpiar"):
            st.session_state[base_k] = _df(disc)
            st.session_state["cp_ver"] = st.session_state.get("cp_ver", 0) + 1
            st.rerun()
        return [{k: (None if isinstance(v, float) and v != v else v) for k, v in r.items()}
                for r in ed.to_dict("records")]
    
    
    @st.cache_data(show_spinner="Dibujando el plano…", max_entries=30)
    def _png(pdf: bytes, pagina: int, dpi: int) -> bytes:
        return P.imagen_pagina(pdf, pagina, dpi)
    
    
    @st.cache_data(max_entries=10)
    def _n_paginas(pdf: bytes) -> int:
        return P.paginas_pdf(pdf)
    
    
    def _visor(pdf: bytes):
        try:
            n = _n_paginas(pdf)
        except Exception as ex:
            st.error(f"No se pudo abrir el PDF: {ex}")
            return
        a, b = st.columns([1, 1])
        pag = a.number_input(f"Lámina / página (de {n})", min_value=1, max_value=n, value=1, step=1, key="cp_pag") \
            if n > 1 else 1
        dpi = b.select_slider("Nitidez", options=[90, 120, 160, 200], value=120, key="cp_dpi",
                              help="Más nitidez = más lento. Usa 160–200 para leer tablas chicas.")
        try:
            st.image(_png(pdf, int(pag) - 1, int(dpi)), use_container_width=True)
        except Exception as ex:
            st.error(f"No se pudo dibujar la página: {ex}")
    
    
    # ------------------------------------------------------------------ resultado
    
    def _resultado(cid, disc, filas, calc, meta, catalogo):
        if not calc:
            st.info("Digita al menos una fila para ver la cubicación.")
            return
        res = P.resumen_items(calc)
        hh = sum(f.hh for f in calc)
        por_u = {}
        for f in calc:
            por_u[f.unidad] = por_u.get(f.unidad, 0) + f.cantidad
        k = st.columns(4)
        k[0].metric("Filas", len(calc))
        k[1].metric("Ítems EC-002", len({f.item for f in calc if f.item}))
        k[2].metric("HH", c.num(hh, 1))
        k[3].metric("Cantidades", " · ".join(f"{c.num(v, 2)} {u}" for u, v in por_u.items() if u) or "—")
        avisos = [f for f in calc if f.aviso]
        t1, t2, t3 = st.tabs(["Resumen por ítem", f"Filas calculadas{' ⚠️' if avisos else ''}", "Comparar"])
        with t1:
            st.dataframe(pd.DataFrame(res), hide_index=True, use_container_width=True,
                         column_config={"Cantidad": st.column_config.NumberColumn(format="%.3f"),
                                        "Rend. HH/u": st.column_config.NumberColumn(format="%.4f"),
                                        "HH": st.column_config.NumberColumn(format="%.1f"),
                                        "Descripción ítem": st.column_config.TextColumn(width="large")})
        with t2:
            if avisos:
                st.warning(f"{len(avisos)} fila(s) con avisos: revisa la columna «Aviso». Los ítems marcados ✳ los "
                           f"propuso la app; si no corresponden, elígelos en la columna «Ítem EC-002» de la planilla.")
            st.dataframe(pd.DataFrame([{"N°": f.n, "Lámina": f.lamina, "Tipo": f.tipo, "Descripción": f.descripcion,
                                        "Cantidad": f.cantidad, "Unidad": f.unidad,
                                        "Ítem EC-002": (f.item + " ✳") if f.item_sugerido else f.item,
                                        "Descripción ítem": f.item_desc, "Rend.": f.rend, "HH": f.hh, "Aviso": f.aviso}
                                       for f in calc]), hide_index=True, use_container_width=True,
                         column_config={"Cantidad": st.column_config.NumberColumn(format="%.3f"),
                                        "HH": st.column_config.NumberColumn(format="%.1f"),
                                        "Rend.": st.column_config.NumberColumn(format="%.4f")})
        comparacion = None
        with t3:
            comparacion = _comparar(cid, meta, res)
        a, b = st.columns(2)
        if a.button("💾 Guardar cubicación", type="primary", key="btn_cp_guardar", disabled=not meta["plano"].strip()):
            if c.exigir("cargar"):
                datos = {"meta": {**meta, "disciplina": disc}, "disciplina": disc, "filas": filas,
                         "resumen": res, "calc": [asdict(f) for f in calc]}
                db.guardar_cub_plano(meta["plano"].strip(), meta["rev"].strip(), disc, meta["descripcion"], datos, hh,
                                     c.nombre_usuario(), contrato_id=cid)
                st.success(f"Cubicación de {meta['plano']} Rev. {meta['rev'] or '—'} guardada.")
        if not meta["plano"].strip():
            a.caption("Escribe el N° de plano para guardar.")
        meta_x = {**meta, "usuario": c.nombre_usuario(), "fecha": dt.date.today().strftime("%d-%m-%Y")}
        b.download_button("⬇️ Exportar cubicación (Excel)", P.excel(meta_x, disc, filas, calc, comparacion),
                          file_name=f"Cubicacion_{(meta['plano'] or 'plano').strip()}"
                                    f"{'_Rev' + meta['rev'].strip() if meta['rev'].strip() else ''}.xlsx",
                          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_cp")
    
    
    def _comparar(cid, meta, res) -> list[dict] | None:
        plano = meta["plano"].strip()
        if not plano:
            st.caption("Escribe el N° de plano para comparar con otras revisiones o con la planilla de avance.")
            return None
        out = None
        guardadas = db.cub_planos(cid)
        otras = guardadas[(guardadas["plano"].map(C.norm) == C.norm(plano)) &
                          (guardadas["rev"].fillna("").map(C.norm) != C.norm(meta["rev"]))]
        st.markdown("**Contra otra revisión guardada del mismo plano**")
        if otras.empty:
            st.caption("No hay otra revisión de este plano guardada.")
        else:
            etiquetas = {int(r.id): f"Rev. {r.rev or '—'} · {str(r.creado_en)[:10]}" for r in otras.itertuples()}
            sel = st.selectbox("Revisión", list(etiquetas), format_func=etiquetas.get, key="cp_cmp_rev")
            ant = (db.cub_plano(int(sel)) or {}).get("resumen", [])
            out = P.comparar(res, ant)
            _tabla_cmp(pd.DataFrame(out).rename(columns={"Anterior": f"{etiquetas[int(sel)].split(' · ')[0]}",
                                                         "Nueva": f"Rev. {meta['rev'] or 'actual'}"}))
        st.markdown("**Contra la planilla de avance**")
        cargas = db.cargas_cubicacion(cid)
        if cargas.empty:
            st.caption("No hay planillas de avance cargadas.")
            return out
        fecha = cargas["fecha"].max()
        filas = []
        for r in cargas[cargas["fecha"] == fecha].itertuples():
            d = db.carga_cubicacion(int(r.id)) or {}
            filas += [f for f in d.get("detalle", []) if C.norm(f.get("plano")) == C.norm(plano)]
        if not filas:
            st.caption(f"La planilla de avance del {c.fecha(fecha)} no tiene filas con el plano {plano}.")
            return out
        plan: dict[str, float] = {}
        for f in filas:
            plan[f["item"]] = plan.get(f["item"], 0) + float(f.get("cantidad") or 0)
        nuevo = {}
        for d in res:
            nuevo[d["Ítem"]] = nuevo.get(d["Ítem"], 0) + d["Cantidad"]
        rend = {d["Ítem"]: d["Rend. HH/u"] for d in res}
        df = pd.DataFrame([{"Ítem": i, "Planilla": plan.get(i, 0.0), "Plano": nuevo.get(i, 0.0),
                            "Diferencia": nuevo.get(i, 0.0) - plan.get(i, 0.0),
                            "Diferencia HH": (nuevo.get(i, 0.0) - plan.get(i, 0.0)) * rend.get(i, 0.0)}
                           for i in sorted(set(plan) | set(nuevo))])
        st.caption(f"Planilla de avance del {c.fecha(fecha)}: {len(filas)} filas con este plano. «Planilla» es la "
                   f"cubicación de esas filas; la diferencia positiva es cantidad que falta en la planilla.")
        _tabla_cmp(df)
        return out
    
    
    def _tabla_cmp(df: pd.DataFrame):
        st.dataframe(df, hide_index=True, use_container_width=True,
                     column_config={k: st.column_config.NumberColumn(format="%.3f") for k in df.columns
                                    if k not in ("Ítem", "Unidad", "Diferencia HH")}
                     | {"Diferencia HH": st.column_config.NumberColumn(format="%.1f")})
    
    
    # ------------------------------------------------------------------ guardadas
    
    def _guardadas(cid):
        st.subheader("Cubicaciones guardadas")
        g = db.cub_planos(cid)
        if g.empty:
            st.caption("Ninguna todavía. Las que guardes quedan aquí para abrirlas, compararlas o exportarlas.")
            return
        st.dataframe(g.assign(creado_en=g["creado_en"].astype(str).str[:16]).drop(columns="id"), hide_index=True,
                     use_container_width=True,
                     column_config={"plano": "N° plano", "rev": "Rev.", "disciplina": "Disciplina",
                                    "descripcion": "Descripción", "hh": st.column_config.NumberColumn("HH", format="%.1f"),
                                    "creado_por": "Cubicó", "creado_en": "Guardada"})
        etiquetas = {int(r.id): f"{r.plano} · Rev. {r.rev or '—'} ({r.disciplina})" for r in g.itertuples()}
        a, b, d = st.columns([3, 1, 1])
        sel = a.selectbox("Cubicación", list(etiquetas), format_func=etiquetas.get, key="cp_sel")
        if b.button("📂 Abrir", key="btn_cp_abrir"):
            datos = db.cub_plano(int(sel)) or {}
            st.session_state["cp_pendiente"] = {"meta": datos.get("meta", {}), "filas": datos.get("filas", [])}
            st.rerun()
        if d.button("🗑️ Eliminar", key="btn_cp_borrar") and c.exigir("cargar"):
            db.borrar_cub_plano(int(sel), contrato_id=cid)
            st.rerun()
        # consolidado de todas
        filas, items = [], []
        for r in g.itertuples():
            datos = db.cub_plano(int(r.id)) or {}
            for x in datos.get("resumen", []):
                items.append({"Plano": r.plano, "Rev": r.rev, "Disciplina": r.disciplina, **x})
            for x in datos.get("calc", []):
                filas.append({"Plano": r.plano, "Rev": r.rev, "Disciplina": r.disciplina, "Lámina": x["lamina"],
                              "Tipo": x["tipo"], "Descripción": x["descripcion"], "Cantidad": x["cantidad"],
                              "Unidad": x["unidad"], "Ítem EC-002": x["item"], "Rend.": x["rend"], "HH": x["hh"]})
        df_items = pd.DataFrame(items)
        tot = (df_items.groupby(["Disciplina", "Ítem", "Descripción ítem", "Unidad"], dropna=False)
               .agg(Cantidad=("Cantidad", "sum"), HH=("HH", "sum"), Planos=("Plano", "nunique")).reset_index()
               if not df_items.empty else df_items)
        st.download_button("⬇️ Exportar todas las cubicaciones (Excel)",
                           C.excel_tablas({"Total por ítem": tot, "Por plano e ítem": df_items, "Filas": pd.DataFrame(filas)},
                                          "Cubicación desde planos"),
                           file_name=f"Cubicaciones_planos_{dt.date.today().strftime('%d-%m-%Y')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_cp_todas")

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


vcubicar = _modulo_vcubicar()

# ======================================================================
# MÓDULO: vcubicacion
# ======================================================================
def _modulo_vcubicacion():
    import datetime as dt
    
    import pandas as pd
    import plotly.graph_objects as go
    import streamlit as st
    
    C = cubicacion
    CP = cubicar_plano
    c = comun
    VCP = vcubicar
    
    ICONOS = {"Obras civiles": "🧱", "Estructura": "🏗️", "Mecánica": "⚙️", "Piping": "🔧", "Eléctricos": "⚡", "Otras": "📦"}
    NAVY, NARANJO, GRIS = "#06205C", "#FA4D09", "#C9D2E3"
    
    
    def vista():
        st.title("Cubicación por disciplina")
        st.caption("**Cubica un plano nuevo** (pestaña 📝) o carga cada semana las **planillas de avance** de cada especialidad (civil, mecánica y estructuras, "
                   "piping, eléctrica). La app lee su itemizado EC-002 y sus hojas de control, calcula el avance por "
                   "disciplina en HH ganadas y revisa errores típicos. Lo cargado aquí también alimenta el itemizado del "
                   "informe semanal.")
        cid = c.contrato_id()
        cargas = db.cargas_cubicacion(cid)
        t_plano, t_av, t_rev, t_det, t_carga = st.tabs(["📝 Cubicar desde plano", "📊 Avance por disciplina",
                                                        "🔎 Revisión de planillas", "📐 Detalle por plano",
                                                        "📥 Cargar planillas"])
        with t_plano:
            VCP.seccion()
        with t_carga:
            if c.exigir("cargar"):
                _cargar(cid, cargas)
        if cargas.empty:
            for t in (t_av, t_rev, t_det):
                with t:
                    st.info("Aún no hay planillas cargadas. Súbelas en la pestaña **📥 Cargar planillas**.")
            return
        fechas = sorted(cargas["fecha"].unique(), reverse=True)
        with t_av:
            fecha = st.selectbox("Semana al", fechas, format_func=c.fecha, key="cub_fecha")
        planillas = _planillas(cid, cargas, fecha)
        previas = [f for f in fechas if f < fecha]
        previo = _acumulados(cid, cargas, previas[0]) if previas else {}
        items = _con_anterior(C.consolidar(planillas), previo)
        with t_av:
            _avance(planillas, items, fecha, previas[0] if previas else None)
        with t_rev:
            _revision(planillas, items)
        with t_det:
            _detalle(planillas)
    
    
    # ------------------------------------------------------------------ datos
    
    @st.cache_data(show_spinner="Leyendo planilla guardada…", max_entries=40)
    def _carga(carga_id: int) -> dict | None:
        return db.carga_cubicacion(carga_id)
    
    
    def _planillas(cid, cargas: pd.DataFrame, fecha: str) -> list[C.PlanillaCub]:
        out = []
        for r in cargas[cargas["fecha"] == fecha].sort_values("id").itertuples():
            d = _carga(int(r.id))
            if d:
                out.append(C.PlanillaCub.de_dict(d))
        return out
    
    
    def _acumulados(cid, cargas, fecha) -> dict[str, float]:
        return {k: it.acumulado for k, it in C.consolidar(_planillas(cid, cargas, fecha)).items()}
    
    
    def _con_anterior(items: dict[str, C.ItemCub], previo: dict[str, float]) -> dict[str, C.ItemCub]:
        """Si la planilla no trae «avance anterior», se usa el acumulado de la carga anterior guardada en la app."""
        import dataclasses
        out = {}
        for k, it in items.items():
            if abs(it.anterior) < 1e-9 and k in previo:
                it = dataclasses.replace(it, anterior=previo[k])
            out[k] = it
        return out
    
    
    # ------------------------------------------------------------------ cargar
    
    def _cargar(cid, cargas):
        st.markdown("Sube una o varias planillas de avance de la **misma semana**. De cada archivo se guarda lo leído "
                    "(ítems, detalle y revisión), no el Excel completo.")
        archivos = st.file_uploader("Planillas de avance (.xlsx)", type=["xlsx", "xlsm"], accept_multiple_files=True,
                                    key="up_cub")
        detectadas = sorted({C.fecha_de_nombre(a.name) for a in archivos or [] if C.fecha_de_nombre(a.name)})
        defecto = dt.datetime.strptime(detectadas[-1], "%d-%m-%Y").date() if detectadas else _viernes_anterior()
        fecha = st.date_input("Fecha de corte (último día de la semana)", value=defecto, format="DD-MM-YYYY",
                              key="cub_fecha_carga")
        if len(detectadas) > 1:
            st.warning(f"Los nombres de los archivos traen fechas distintas ({', '.join(detectadas)}). Se guardarán "
                       f"todos con la fecha de corte elegida.")
        for tipo, m in st.session_state.pop("cub_msgs", []):
            (st.success if tipo == "ok" else st.error)(m)
        if archivos and st.button(f"Leer y guardar {len(archivos)} planilla(s)", type="primary", key="btn_cub"):
            ok, msgs = 0, []
            for a in archivos:
                try:
                    with st.spinner(f"Leyendo {a.name}…"):
                        p = C.leer_planilla(a.getvalue(), a.name)
                    try:                         # el itemizado completo queda como catálogo EC-002 del contrato
                        db.guardar_config(cid, "catalogo_ec002", CP.unir_catalogos(
                            db.config(cid, "catalogo_ec002", []), CP.catalogo_de_planilla(a.getvalue())))
                    except Exception:
                        pass
                    db.guardar_carga_cubicacion(fecha.isoformat(), a.name, ", ".join(p.disciplinas), p.a_dict(),
                                                c.nombre_usuario(), contrato_id=cid)
                    ok += 1
                    msgs.append(("ok", f"✅ {a.name}: {len(p.items)} ítems ({', '.join(p.disciplinas)}), "
                                       f"{len(p.detalle)} filas de detalle, {len(p.hallazgos)} observaciones."))
                except Exception as e:
                    msgs.append(("error", f"{a.name}: {e}"))
            st.session_state["cub_msgs"] = msgs
            if ok:
                c.limpiar_cache()
                st.session_state["cub_fecha"] = fecha.isoformat()
            st.rerun()
        st.divider()
        st.markdown("**Planillas guardadas**")
        if cargas.empty:
            st.caption("Ninguna todavía.")
            return
        vista = cargas.assign(fecha=cargas["fecha"].map(c.fecha), creado_en=cargas["creado_en"].astype(str).str[:16])
        st.dataframe(vista[["fecha", "archivo", "disciplinas", "creado_por", "creado_en"]], hide_index=True,
                     use_container_width=True,
                     column_config={"fecha": "Semana al", "archivo": "Archivo", "disciplinas": "Disciplinas",
                                    "creado_por": "Cargó", "creado_en": "Fecha de carga"})
        etiquetas = {int(r.id): f"{c.fecha(r.fecha)} · {r.archivo}" for r in cargas.itertuples()}
        a, b = st.columns([3, 1])
        sel = a.selectbox("Eliminar una carga", list(etiquetas), format_func=etiquetas.get, key="cub_borrar_sel")
        if b.button("Eliminar", key="btn_cub_borrar"):
            db.borrar_carga_cubicacion(int(sel), contrato_id=cid)
            c.limpiar_cache()
            st.rerun()
    
    
    def _viernes_anterior() -> dt.date:
        hoy = dt.date.today()
        return hoy - dt.timedelta(days=(hoy.weekday() - 4) % 7)
    
    
    # ------------------------------------------------------------------ avance
    
    def _avance(planillas, items, fecha, fecha_prev):
        res = {r["disciplina"]: r for r in C.resumen_disciplinas(items)}
        con_planilla = {d for p in planillas for d in p.disciplinas}
        tot_pom = sum(r["hh_pom"] for r in res.values())
        tot_gan = sum(r["hh_ganadas"] for r in res.values())
        tot_sem = sum(r["hh_semana"] for r in res.values())
        k = st.columns(4)
        k[0].metric("HH POM (cubicadas)", c.num(tot_pom))
        k[1].metric("HH ganadas acumuladas", c.num(tot_gan))
        k[2].metric("Avance físico (HH)", c.porc(100 * tot_gan / tot_pom if tot_pom else 0, 1))
        k[3].metric("HH ganadas en la semana", c.num(tot_sem))
        st.caption(f"Planillas de la semana al {c.fecha(fecha)}: " + " · ".join(p.archivo for p in planillas)
                   + (f". El avance de la semana se mide contra la carga del {c.fecha(fecha_prev)} cuando la planilla no "
                      f"trae avance anterior." if fecha_prev else "."))
    
        sin_ant = sorted({h.disciplina for p in planillas for h in p.hallazgos if h.tipo == "Planilla sin avance anterior"})
        if sin_ant and not fecha_prev:
            st.warning(f"La planilla de {', '.join(sin_ant)} no trae «avance anterior» y no hay una carga de la semana "
                       f"previa en la app: sus HH de la semana muestran todo el acumulado. Desde la próxima carga se "
                       f"calculan contra esta.")
        filas = []
        for d in C.DISCIPLINAS + [C.OTRAS]:
            r = res.get(d)
            if d == C.OTRAS and (not r or not r["items"]):
                continue
            if not r or d not in con_planilla:
                filas.append({"Disciplina": f"{ICONOS[d]} {d}", "Estado": "Sin planilla cargada", "Ítems": 0,
                              "HH POM": None, "HH ganadas": None, "HH semana": None, "Avance": None, "Avance anterior": None})
                continue
            filas.append({"Disciplina": f"{ICONOS[d]} {d}", "Estado": "Cargada", "Ítems": r["items"], "HH POM": r["hh_pom"],
                          "HH ganadas": r["hh_ganadas"], "HH semana": r["hh_semana"], "Avance": 100 * r["avance"],
                          "Avance anterior": 100 * r["avance_anterior"]})
        df = pd.DataFrame(filas)
        st.dataframe(df, hide_index=True, use_container_width=True, column_config={
            "HH POM": st.column_config.NumberColumn(format="%.0f"),
            "HH ganadas": st.column_config.NumberColumn(format="%.0f"),
            "HH semana": st.column_config.NumberColumn(format="%.0f"),
            "Avance": st.column_config.ProgressColumn("Avance (HH)", format="%.1f%%", min_value=0, max_value=100),
            "Avance anterior": st.column_config.ProgressColumn("Semana anterior", format="%.1f%%", min_value=0,
                                                               max_value=100)})
        g = df.dropna(subset=["Avance"])
        if not g.empty:
            fig = go.Figure()
            fig.add_bar(y=g["Disciplina"], x=g["Avance anterior"], name="Semana anterior", orientation="h",
                        marker_color=GRIS)
            fig.add_bar(y=g["Disciplina"], x=g["Avance"], name="Acumulado", orientation="h", marker_color=NAVY,
                        text=[f"{v:.1f}%" for v in g["Avance"]], textposition="outside")
            fig.update_layout(barmode="group", xaxis=dict(range=[0, 110], ticksuffix="%"), yaxis=dict(autorange="reversed"),
                              title="Avance físico por disciplina (HH ganadas / HH POM)")
            st.plotly_chart(c.layout_grafico(fig, 300 + 30 * len(g)), use_container_width=True)
        st.download_button("⬇️ Descargar consolidado (Excel)", C.excel_resumen(planillas, c.fecha(fecha)),
                           file_name=f"Cubicacion_por_disciplina_{c.fecha(fecha)}.xlsx", key="dl_cub",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
        st.subheader("Ítems de una disciplina")
        disponibles = [d for d in C.DISCIPLINAS + [C.OTRAS] if any(i.disciplina == d for i in items.values())]
        a, b, f = st.columns([1.2, 2, 1])
        disc = a.selectbox("Disciplina", disponibles, format_func=lambda d: f"{ICONOS[d]} {d}", key="cub_disc")
        buscar = b.text_input("Buscar ítem o descripción", key="cub_buscar")
        solo = f.checkbox("Solo con avance en la semana", key="cub_solo")
        sel = [i for i in items.values() if i.disciplina == disc]
        if buscar:
            q = C.norm(buscar)
            sel = [i for i in sel if q in C.norm(i.item) or q in C.norm(i.descripcion)]
        if solo:
            sel = [i for i in sel if i.semana > 1e-6]
        tabla = pd.DataFrame([{"Ítem": i.item, "Descripción": i.descripcion, "Un.": i.unidad, "Contrato": i.contrato,
                               "POM": i.pom, "Anterior": i.anterior, "Acumulado": i.acumulado, "Semana": i.semana,
                               "Avance": 100 * i.acumulado / i.pom_valida if i.pom_valida else 0,
                               "HH ganadas": i.hh_ganadas, "Planilla de control": i.hoja_control}
                              for i in sorted(sel, key=lambda x: x.item)])
        if tabla.empty:
            st.caption("Sin ítems para el filtro.")
            return
        st.dataframe(tabla, hide_index=True, use_container_width=True, height=min(600, 38 + 35 * len(tabla)),
                     column_config={**{k: st.column_config.NumberColumn(format="%.2f")
                                       for k in ("Contrato", "POM", "Anterior", "Acumulado", "Semana")},
                                    "HH ganadas": st.column_config.NumberColumn(format="%.0f"),
                                    "Descripción": st.column_config.TextColumn(width="large"),
                                    "Avance": st.column_config.ProgressColumn("% avance", format="%.0f%%", min_value=0,
                                                                              max_value=100)})
    
    
    # ------------------------------------------------------------------ revisión
    
    def _revision(planillas, items):
        hall = [h for p in planillas for h in p.hallazgos] + C.cruzar(planillas)
        if not hall:
            st.success("Sin observaciones en las planillas de esta semana.")
            return
        n = {g: sum(1 for h in hall if h.gravedad == g) for g in ("Error", "Revisar", "Info")}
        k = st.columns(3)
        k[0].metric("🔴 Errores", n["Error"], help="Avance sobre la cubicación, retrocesos, detalle que no cuadra…")
        k[1].metric("🟠 Para revisar", n["Revisar"], help="TAG repetidos, filas sin plano, ítems sin rendimiento…")
        k[2].metric("ℹ️ Informativas", n["Info"], help="Mayor obra: cubicación POM más de 10 % sobre el contrato.")
        a, b = st.columns(2)
        discs = sorted({h.disciplina for h in hall}, key=lambda d: (C.DISCIPLINAS + [C.OTRAS]).index(d)
                       if d in C.DISCIPLINAS + [C.OTRAS] else 9)
        f_disc = a.multiselect("Disciplina", discs, default=discs, key="cub_rev_disc")
        f_grav = b.multiselect("Gravedad", ["Error", "Revisar", "Info"], default=["Error", "Revisar"], key="cub_rev_grav")
        sel = [h for h in hall if h.disciplina in f_disc and h.gravedad in f_grav]
        orden = {"Error": 0, "Revisar": 1, "Info": 2}
        df = pd.DataFrame([{"Gravedad": {"Error": "🔴 Error", "Revisar": "🟠 Revisar", "Info": "ℹ️ Info"}[h.gravedad],
                            "Disciplina": h.disciplina, "Tipo": h.tipo, "Ítem": h.item, "Detalle": h.detalle,
                            "Planilla": h.planilla, "_o": orden[h.gravedad]} for h in sel])
        if df.empty:
            st.caption("Sin observaciones para el filtro.")
            return
        df = df.sort_values(["_o", "Disciplina", "Tipo", "Ítem"]).drop(columns="_o")
        st.dataframe(df, hide_index=True, use_container_width=True, height=min(650, 38 + 35 * len(df)),
                     column_config={"Detalle": st.column_config.TextColumn(width="large")})
        resumen = df.groupby(["Disciplina", "Tipo"]).size().reset_index(name="Cantidad")
        with st.expander("Resumen por tipo"):
            st.dataframe(resumen, hide_index=True, use_container_width=True)
    
    
    # ------------------------------------------------------------------ detalle
    
    def _detalle(planillas):
        filas = []
        for p in planillas:
            for f in p.detalle:
                it = p.items.get(f.item)
                filas.append({"Disciplina": it.disciplina if it else C.disciplina_de("", f.item), "Hoja": f.hoja,
                              "Fila": f.fila, "Ítem": f.item, "Plano": f.plano, "Rev": f.rev, "TAG": f.tag,
                              "Descripción": f.descripcion, "Un.": f.unidad, "Cubicación": f.cantidad,
                              "Acumulado": f.acumulado})
        if not filas:
            st.info("Las planillas cargadas no traen hojas de detalle legibles.")
            return
        df = pd.DataFrame(filas)
        a, b, f = st.columns([1.2, 1.6, 2])
        discs = [d for d in C.DISCIPLINAS + [C.OTRAS] if d in set(df["Disciplina"])]
        disc = a.selectbox("Disciplina", discs, format_func=lambda d: f"{ICONOS[d]} {d}", key="cub_det_disc")
        df = df[df["Disciplina"] == disc]
        hoja = b.selectbox("Hoja de control", ["(todas)"] + list(dict.fromkeys(df["Hoja"])), key="cub_det_hoja")
        if hoja != "(todas)":
            df = df[df["Hoja"] == hoja]
        q = f.text_input("Buscar plano, TAG, ítem o descripción", key="cub_det_buscar")
        if q:
            qn = C.norm(q)
            df = df[df.apply(lambda r: any(qn in C.norm(r[k]) for k in ("Plano", "TAG", "Ítem", "Descripción")), axis=1)]
        avisos = {h: a for p in planillas for h, a in p.hojas_detalle.items() if a}
        if hoja in avisos or (hoja == "(todas)" and any(h in avisos for h in set(df["Hoja"]))):
            st.caption("ℹ️ " + " · ".join(f"«{h}»: {a}" for h, a in avisos.items() if hoja in ("(todas)", h)
                                          and h in set(df["Hoja"])))
        por_plano = (df.assign(Plano=df["Plano"].replace("", "(sin plano)"))
                     .groupby(["Plano", "Rev"], dropna=False)
                     .agg(Filas=("Fila", "count"), Ítems=("Ítem", "nunique"), Cubicación=("Cubicación", "sum"),
                          Acumulado=("Acumulado", "sum")).reset_index()
                     .sort_values("Filas", ascending=False))
        k = st.columns(3)
        k[0].metric("Filas", len(df))
        k[1].metric("Planos", int(por_plano.shape[0]))
        k[2].metric("Ítems", int(df["Ítem"].nunique()))
        plano_item = (df.assign(Plano=df["Plano"].replace("", "(sin plano)"))
                      .groupby(["Plano", "Rev", "Ítem", "Un."], dropna=False)
                      .agg(Filas=("Fila", "count"), Cubicación=("Cubicación", "sum"), Acumulado=("Acumulado", "sum"))
                      .reset_index())
        plano_item["% avance"] = (plano_item["Acumulado"] / plano_item["Cubicación"]).where(plano_item["Cubicación"] > 0)
        st.download_button(
            "⬇️ Exportar cubicación a Excel (filtro actual)",
            C.excel_tablas({"Por plano": por_plano, "Por plano e ítem": plano_item, "Filas": df.drop(columns="Disciplina")},
                           f"Cubicación {disc}" + (f" · {hoja}" if hoja != "(todas)" else "")),
            file_name=f"Cubicacion_{C.norm(disc).replace(' ', '_')}{'_' + C.norm(hoja)[:25].replace(' ', '_') if hoja != '(todas)' else ''}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_cub_det")
        st.markdown("**Por plano**")
        st.dataframe(por_plano, hide_index=True, use_container_width=True, height=min(400, 38 + 35 * len(por_plano)),
                     column_config={"Cubicación": st.column_config.NumberColumn(format="%.2f"),
                                    "Acumulado": st.column_config.NumberColumn(format="%.2f")})
        st.markdown("**Filas de la planilla**")
        st.dataframe(df.drop(columns="Disciplina"), hide_index=True, use_container_width=True,
                     height=min(600, 38 + 35 * len(df)),
                     column_config={"Cubicación": st.column_config.NumberColumn(format="%.2f"),
                                    "Acumulado": st.column_config.NumberColumn(format="%.2f"),
                                    "Descripción": st.column_config.TextColumn(width="large")})

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


vcubicacion = _modulo_vcubicacion()

# ======================================================================
# MÓDULO: vsemanal
# ======================================================================
def _modulo_vsemanal():
    import datetime as dt
    
    import pandas as pd
    import streamlit as st
    
    SM = semanal
    CUB = cubicacion
    c = comun
    
    MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    
    
    @st.cache_resource(show_spinner="Leyendo el informe semanal base…", max_entries=2)
    def _plantilla(archivo_id: int):
        return SM.leer_plantilla(db.archivo_semanal(archivo_id))
    
    
    @st.cache_resource(show_spinner="Leyendo la BD Personal del último informe diario…", max_entries=2)
    def _bd_personal(archivo_id: int):
        return SM.leer_bd_personal(db.archivo_diario(archivo_id))
    
    
    def _actividades_p6():
        """Actividades del corte P6 más reciente del contrato (para la fecha real de los hitos)."""
        pr = c.proyectos()
        if pr is None or pr.empty:
            return None
        mejor = None
        for _, p in pr.iterrows():
            cts = c.cortes(int(p["id"]))
            if cts is not None and not cts.empty:
                u = cts.iloc[-1]
                if mejor is None or str(u["fecha_corte"]) > str(mejor["fecha_corte"]):
                    mejor = u
        return c.actividades(int(mejor["id"])) if mejor is not None else None
    
    
    def crear():
        cid = c.contrato_id()
        base = db.base_semanal_vigente(cid)
        _planilla_base(base, cid)
        if not base:
            return
        try:
            pl = _plantilla(int(base["id"]))
        except Exception as e:
            st.error(f"No se pudo leer el informe base: {e}")
            return
        for a in pl.avisos:
            st.warning(a)
    
        inicio = pl.inicio + dt.timedelta(days=7)
        fin = inicio + dt.timedelta(days=6)
        st.markdown(f"### Informe semanal N° {pl.numero + 1} · {c.fecha(inicio)} al {c.fecha(fin)}")
    
        # Datos de la semana desde el último informe diario
        base_d = db.base_diario_vigente(cid)
        sem, bd = None, None
        if base_d:
            try:
                bd = _bd_personal(int(base_d["id"]))
                sem = SM.calcular_semana(pl, bd, inicio)
            except Exception as e:
                st.error(f"No se pudo leer la BD Personal del informe diario: {e}")
        if sem is None:
            st.info("No hay informe diario base cargado: las HH y la dotación se ingresan a mano.")
            sem = SM.Semana(inicio, fin, pl.numero + 1)
        elif sem.dias_sin_datos:
            st.warning("La BD Personal del último informe diario no tiene los días "
                       + ", ".join(c.fecha(d) for d in sem.dias_sin_datos)
                       + ". Genera los informes diarios que faltan o corrige las HH a mano.")
    
        clave = f"sem_{base['id']}_{inicio}"
        if clave not in st.session_state:
            st.session_state[clave] = SM.entradas_sugeridas(pl, sem, _actividades_p6())
        ent: SM.EntradasSemanal = st.session_state[clave]
    
        tabs = st.tabs(["0. Portada", "1. Resumen ejecutivo", "2. Seguridad", "3. Curva de avance", "4.0 Dotación",
                        "5. Curva dotación", "7. Itemizado (planillas de avance)", "✅ Revisar y generar"])
        with tabs[0]:
            _portada(ent)
        with tabs[1]:
            _resumen(pl, ent)
        with tabs[2]:
            _seguridad(pl, ent)
        with tabs[3]:
            _curva(pl, fin)
        with tabs[4]:
            _dotacion(pl, ent)
        with tabs[5]:
            _curva_dotacion(pl, ent)
        with tabs[6]:
            _itemizado(pl, ent, clave)
        with tabs[7]:
            _generar(base, pl, ent, sem, cid)
    
    
    # ------------------------------------------------------------------ planilla base
    
    def _planilla_base(base, cid):
        with st.expander("📄 Informe semanal base (el último emitido)", expanded=not base):
            if base:
                st.caption(f"Base vigente: **{base['nombre']}** · N° {base['numero']} · semana al {c.fecha(base['fecha'])} · "
                           f"{'generado por el sistema' if base['tipo'] == 'generado' else 'cargado'} el {base['creado_en'][:16]}")
                st.download_button("Descargar base", db.archivo_semanal(int(base["id"])), file_name=base["nombre"],
                                   mime=MIME_XLSX, key="dl_base_sem")
            archivo = st.file_uploader("Cargar o reemplazar el informe semanal base (.xlsx)", type=["xlsx", "xlsm"],
                                       key="up_base_sem")
            if archivo is not None:
                try:
                    pl = SM.leer_plantilla(archivo.getvalue())
                except Exception as e:
                    st.error(f"No se pudo leer el informe: {e}")
                    return
                st.success(f"Informe N° {pl.numero} · semana del {c.fecha(pl.inicio)} al {c.fecha(pl.fin)}.")
                if st.button("Usar como base", type="primary", key="btn_base_sem"):
                    db.guardar_base_semanal(archivo.name, pl.fin.isoformat(), pl.numero, archivo.getvalue(),
                                            c.nombre_usuario(), "base", contrato_id=cid)
                    st.cache_resource.clear()
                    st.rerun()
    
    
    # ------------------------------------------------------------------ secciones
    
    def _tabla(registros: list[dict], columnas: dict, clave: str, dinamica: bool = False, **cfg) -> list[dict]:
        df = pd.DataFrame(registros, columns=list(columnas)) if registros else pd.DataFrame(columns=list(columnas))
        ed = st.data_editor(df, column_config=columnas, hide_index=True, use_container_width=True,
                            num_rows="dynamic" if dinamica else "fixed", key=clave, **cfg)
        out = []
        for r in ed.to_dict("records"):
            out.append({k: (None if (isinstance(v, float) and pd.isna(v)) or v is pd.NaT else v) for k, v in r.items()})
        return out
    
    
    def _portada(ent):
        st.caption("Registro de revisiones de la portada. Se propone la primera revisión del informe anterior con la fecha "
                   "de cierre de la semana.")
        cols = {"rev": st.column_config.TextColumn("Rev"), "fecha": st.column_config.DateColumn("Fecha", format="DD-MM-YYYY"),
                "aprobo": st.column_config.TextColumn("Aprobó"), "reviso": st.column_config.TextColumn("Revisó"),
                "preparo": st.column_config.TextColumn("Preparó"), "emitido": st.column_config.TextColumn("Emitido para", width="large")}
        filas = _tabla(ent.revisiones, cols, "sem_rev", dinamica=True)
        ent.revisiones = [r for r in filas if any(v not in (None, "") for v in r.values())]
    
    
    def _resumen(pl, ent):
        st.markdown("**1.1 Resumen de avances** — se calcula solo desde la hoja 7 (itemizado EC-002), con el avance de las "
                    "planillas cargadas en la pestaña «7. Itemizado». La hoja POM 5 no se toca.")
        st.markdown("**1.2 Datos financieros**")
        c1, c2 = st.columns(2)
        with c1:
            for ref, et in pl.etiquetas_fin.items():
                ent.financieros[ref] = st.number_input(et, value=float(ent.financieros.get(ref) or 0), step=1.0, format="%.0f",
                                                       key=f"fin_{ref}")
        with c2:
            st.caption("Estados de pago facturados")
            cols = {"nombre": st.column_config.TextColumn("EDP", disabled=True),
                    "monto": st.column_config.NumberColumn("Monto parcial ($)", format="%.0f")}
            filas = _tabla([{"nombre": e["nombre"], "monto": e["monto"]} for e in ent.edp], cols, "sem_edp")
            for e, f in zip(ent.edp, filas):
                e["monto"] = f["monto"]
    
        st.markdown("**1.3 Actividades críticas planificadas**")
        cols = {"texto": st.column_config.TextColumn("Actividad", width="large"),
                "obs": st.column_config.TextColumn("Observación", width="large")}
        filas = _tabla([{"texto": x["texto"], "obs": x["obs"]} for x in ent.criticas], cols, "sem_cri")
        for x, f in zip(ent.criticas, filas):
            x.update(f)
        st.markdown("**Consideraciones y alertas**")
        cols = {"tema": st.column_config.TextColumn("Tema", width="medium"),
                "detalle": st.column_config.TextColumn("Detalle", width="large")}
        filas = _tabla([{"tema": x["tema"], "detalle": x["detalle"]} for x in ent.alertas], cols, "sem_ale")
        for x, f in zip(ent.alertas, filas):
            x.update(f)
    
        st.markdown("**1.4 Hitos contractuales** — la fecha real se completa sola cuando el hito está terminado en el "
                    "último programa P6 cargado.")
        cols = {"id": st.column_config.TextColumn("ID", disabled=True), "nombre": st.column_config.TextColumn("Hito", disabled=True),
                "tipo": st.column_config.TextColumn("Tipo", disabled=True),
                "programada": st.column_config.DateColumn("Programa", format="DD-MM-YYYY", disabled=True),
                "real": st.column_config.DateColumn("Real", format="DD-MM-YYYY"),
                "comentario": st.column_config.TextColumn("Comentario", width="large")}
        filas = _tabla([{k: h.get(k) for k in cols} for h in ent.hitos], cols, "sem_hit")
        for h, f in zip(ent.hitos, filas):
            h["real"] = f["real"].date() if hasattr(f["real"], "date") and not isinstance(f["real"], dt.date) else f["real"]
            h["comentario"] = f["comentario"]
        st.caption("1.5 Resumen de dotación: se toma de la hoja 4.0 Dotación.")
    
    
    def _seguridad(pl, ent):
        hh = ent.hh_mod + ent.hh_moi
        c1, c2, c3 = st.columns(3)
        c1.metric("HH de la semana (MOD + MOI)", c.num(hh))
        c2.metric("HH acumuladas anterior", c.num(pl.seg_hh_acum))
        c3.metric("HH acumuladas nuevas", c.num(pl.seg_hh_acum + hh))
        st.caption("Las HH se toman de la BD Personal del informe diario (11 HH por persona y día). Se corrigen en la "
                   "pestaña 5. Curva dotación.")
        st.markdown("**Accidentes del período** (el acumulado anterior es el acumulado del informe pasado)")
        cols = st.columns(3)
        for col, et in zip(cols, ("N° De Accidentes STP", "N° De Accidentes CTP", "Días Perdidos")):
            prev = (pl.seg_valores.get(et) or (0, 0))[1] or 0
            ent.seguridad[et] = col.number_input(f"{et} (acum. anterior {c.num(prev)})", min_value=0.0,
                                                 value=float(ent.seguridad.get(et) or 0), step=1.0, key=f"seg_{et}")
    
    
    def _curva(pl, fin):
        if not pl.curva_actual:
            st.warning("No se encontró la fila de la semana en la hoja 3. Curva de Avance.")
            return
        st.markdown(f"La fila de la semana al **{c.fecha(pl.curva_fechas.get(pl.curva_actual))}** queda con sus valores y "
                    f"la fila al **{c.fecha(fin)}** toma las fórmulas del itemizado EC-002 (HH ganadas, gastadas y % real).")
        st.caption("El avance de la semana sale de las planillas cargadas en la pestaña «7. Itemizado». Los ítems sin "
                   "planilla quedan sin avance en la semana.")
    
    
    def _dotacion(pl, ent):
        st.caption("Promedio de la semana por cargo (BD Personal del informe diario): contratado = personas registradas, "
                   "en faena = día T (indirectos también TT/TC acreditados) o noche TN, descanso = el resto.")
        filas = [{"cargo": k, "tipo": d["tipo"], "contratado": d["contratado"], "faena": d["faena"], "descanso": d["descanso"],
                  "previo": (pl.dot_previo.get(SM.norm(k)) or (None,))[0],
                  "nuevo": SM.norm(k) not in pl.dot_directos and SM.norm(k) not in pl.dot_indirectos}
                 for k, d in sorted(ent.dotacion.items(), key=lambda kv: (kv[1]["tipo"], kv[0]))]
        cols = {"cargo": st.column_config.TextColumn("Cargo", width="large"),
                "tipo": st.column_config.SelectboxColumn("Tipo", options=["Directo", "Indirecto"]),
                "contratado": st.column_config.NumberColumn("Contratado", format="%.1f"),
                "faena": st.column_config.NumberColumn("En faena", format="%.1f"),
                "descanso": st.column_config.NumberColumn("En descanso", format="%.1f", disabled=True),
                "previo": st.column_config.NumberColumn("Contratado sem. anterior", format="%.1f", disabled=True),
                "nuevo": st.column_config.CheckboxColumn("Cargo nuevo", disabled=True)}
        out = _tabla(filas, cols, "sem_dot", dinamica=True)
        nueva = {}
        for r in out:
            if not r.get("cargo"):
                continue
            ctr, fae = float(r.get("contratado") or 0), float(r.get("faena") or 0)
            nueva[r["cargo"]] = {"tipo": r.get("tipo") or "Directo", "contratado": ctr, "faena": fae, "descanso": round(ctr - fae, 1)}
        ent.dotacion = nueva
        tot = pd.DataFrame(nueva.values())
        if not tot.empty:
            g = tot.groupby("tipo")[["contratado", "faena", "descanso"]].sum().round(1)
            st.dataframe(g, use_container_width=True)
        nuevos = [k for k in nueva if SM.norm(k) not in pl.dot_directos and SM.norm(k) not in pl.dot_indirectos]
        if nuevos:
            st.info(f"{len(nuevos)} cargo(s) nuevo(s) se agregan a la hoja 4.0: " + ", ".join(nuevos))
    
    
    def _curva_dotacion(pl, ent):
        st.caption("HH reales de la semana (11 HH por persona y día, mismo criterio de las secciones 2.0 y 3.0 del informe "
                   "diario). Puedes corregirlas.")
        c1, c2 = st.columns(2)
        ent.hh_mod = c1.number_input("HH MOD de la semana", min_value=0.0, value=float(ent.hh_mod), step=11.0, key="sem_hhmod")
        c1.caption(f"Semana anterior: {c.num(pl.hh_semana_previa.get('MOD'))} HH · acumulado nuevo "
                   f"{c.num(pl.mod_acum + ent.hh_mod)} HH")
        ent.hh_moi = c2.number_input("HH MOI de la semana", min_value=0.0, value=float(ent.hh_moi), step=11.0, key="sem_hhmoi")
        c2.caption(f"Semana anterior: {c.num(pl.hh_semana_previa.get('MOI'))} HH · acumulado nuevo "
                   f"{c.num(pl.moi_acum + ent.hh_moi)} HH")
        c1.metric("Dotación MOD promedio", c.num(ent.hh_mod / 77, 1))
        c2.metric("Dotación MOI promedio", c.num(ent.hh_moi / 77, 1))
    
    
    def _itemizado(pl, ent, clave):
        if not pl.itm_hoja:
            st.warning("El informe base no tiene la hoja «7.Avance EC-002 Itemizado».")
            return
        st.markdown("Sube las **planillas de avance** de la semana (piping, civil, mecánica-estructuras, eléctrica). Del "
                    "itemizado EC-002 de cada una se "
                    "toma la **cantidad acumulada WEEKLY** de cada ítem y pasa a la columna *Avance acumulado* del itemizado "
                    "EC-002. El acumulado del informe anterior pasa a *Avance anterior*. Los ítems que no vienen en ninguna "
                    "planilla quedan sin avance en la semana. La hoja POM 5 no se toca.")
        archivos = st.file_uploader("Planillas de avance (.xlsx)", type=["xlsx", "xlsm"], accept_multiple_files=True,
                                    key=f"up_plan_{clave}")
        k_pl = f"planillas_{clave}"
        cargas = db.cargas_cubicacion(c.contrato_id())
        if not cargas.empty:
            fechas = sorted(cargas["fecha"].unique(), reverse=True)
            a, b = st.columns([2, 1])
            f_cub = a.selectbox("…o usar las planillas guardadas en **Cubicación por disciplina** (semana al)", fechas,
                                format_func=c.fecha, key=f"sel_cub_{clave}")
            if b.button("Usar planillas guardadas", key="btn_plan_cub"):
                leidas = {}
                for r in cargas[cargas["fecha"] == f_cub].sort_values("id").itertuples():
                    d = db.carga_cubicacion(int(r.id))
                    try:
                        leidas[r.archivo] = SM.desde_cubicacion(CUB.PlanillaCub.de_dict(d))
                    except Exception as e:
                        st.error(f"{r.archivo}: {e}")
                st.session_state[k_pl] = leidas
        if archivos and st.button(f"Leer {len(archivos)} planilla(s)", type="primary", key="btn_plan"):
            leidas, errores = {}, []
            for a in archivos:
                try:
                    leidas[a.name] = _planilla(a.getvalue(), a.name)
                except Exception as e:
                    errores.append(f"{a.name}: {e}")
            for e in errores:
                st.error(e)
            st.session_state[k_pl] = leidas
        leidas = st.session_state.get(k_pl) or {}
        if not leidas:
            ent.avance_items, ent.fuentes_avance = {}, []
            st.info("Sin planillas cargadas: el itemizado pasa a la semana nueva sin avance (período 0).")
            return
    
        avance, fuente, filas, fuera = {}, {}, [], []
        for nombre, p in leidas.items():
            for item, d in p["items"].items():
                if item not in pl.itm_filas:
                    fuera.append(item)
                    continue
                avance[item] = avance.get(item, 0.0) + d["cant"]
                fuente[item] = nombre
        for item, cant in sorted(avance.items()):
            info = pl.itm_info.get(item, {})
            ant = info.get("cant") or 0.0
            rend = info.get("rend") or 0.0
            filas.append({"item": item, "descripcion": info.get("descripcion", ""), "unidad": info.get("unidad", ""),
                          "anterior": round(ant, 3), "acumulado": round(cant, 3), "periodo": round(cant - ant, 3),
                          "hh": round((cant - ant) * rend, 1), "pom": round(info.get("pom") or 0, 3),
                          "planilla": fuente[item]})
        df = pd.DataFrame(filas)
        cols = {"item": st.column_config.TextColumn("Ítem", disabled=True),
                "descripcion": st.column_config.TextColumn("Descripción", disabled=True, width="large"),
                "unidad": st.column_config.TextColumn("Un.", disabled=True, width="small"),
                "anterior": st.column_config.NumberColumn("Acumulado informe anterior", format="%.2f", disabled=True),
                "acumulado": st.column_config.NumberColumn("Acumulado nuevo", format="%.2f"),
                "periodo": st.column_config.NumberColumn("Período", format="%.2f", disabled=True),
                "hh": st.column_config.NumberColumn("HH ganadas período", format="%.0f", disabled=True),
                "pom": st.column_config.NumberColumn("Cantidad POM 5", format="%.2f", disabled=True),
                "planilla": st.column_config.TextColumn("Planilla", disabled=True)}
        ed = st.data_editor(df, column_config=cols, hide_index=True, use_container_width=True, key=f"ed_itm_{clave}")
        ent.avance_items = {r["item"]: float(r["acumulado"] or 0) for r in ed.to_dict("records")}
        ent.fuentes_avance = list(leidas)
        per = df["periodo"].sum() if not df.empty else 0
        k = st.columns(4)
        k[0].metric("Ítems con avance", len(df))
        k[1].metric("Ítems que avanzan esta semana", int((df["periodo"] > 0.0001).sum()) if not df.empty else 0)
        k[2].metric("HH ganadas en la semana", c.num(df["hh"].sum() if not df.empty else 0))
        k[3].metric("Ítems con retroceso", int((df["periodo"] < -0.0001).sum()) if not df.empty else 0,
                    help="Acumulado nuevo menor que el del informe anterior: revisa la planilla.")
        if per < 0 or (not df.empty and (df["periodo"] < -0.0001).any()):
            st.warning("Hay ítems con acumulado menor al del informe anterior (período negativo). Revisa si la planilla "
                       "corresponde a esta semana.")
        for nombre, p in leidas.items():
            dif = [it for it, d in p["items"].items() if it in pl.itm_info
                   and abs(d["anterior"] - (pl.itm_info[it].get("cant") or 0)) > 0.05]
            if dif:
                st.caption(f"ℹ️ {nombre}: en {len(dif)} ítem(s) el «avance anterior» de la planilla no coincide con el acumulado "
                           f"del informe anterior (p. ej. {', '.join(dif[:4])}).")
        if fuera:
            st.caption(f"⚠️ {len(fuera)} ítem(s) de las planillas no existen en el itemizado del informe y se omiten: "
                       + ", ".join(sorted(set(fuera))[:8]))
    
    
    @st.cache_data(show_spinner="Leyendo planilla de avance…", max_entries=6)
    def _planilla(datos: bytes, nombre: str):
        return SM.leer_planilla_avance(datos, nombre)
    
    
    def _generar(base, pl, ent, sem, cid):
        st.markdown("**Revisión**")
        fila = lambda ok, t, d: st.markdown(f"{'✅' if ok else '⚠️'} **{t}** — {d}")
        fila(True, "Portada", f"{len(ent.revisiones)} revisión(es)")
        fila(True, "1. Resumen", f"{len(ent.criticas)} actividades críticas · {len(ent.alertas)} alertas · "
                                 f"{sum(1 for h in ent.hitos if h.get('real'))} de {len(ent.hitos)} hitos con fecha real")
        fila(pl.seg_ultima > 0, "2. Seguridad", f"{c.num(ent.hh_mod + ent.hh_moi)} HH en la semana")
        fila(pl.curva_actual > 0, "3. Curva de avance", "rueda a la semana nueva (itemizado sin cambios)")
        fila(bool(ent.dotacion), "4.0 Dotación", f"{len(ent.dotacion)} cargos")
        fila(ent.hh_mod > 0 and ent.hh_moi > 0, "5. Curva dotación", f"MOD {c.num(ent.hh_mod)} HH · MOI {c.num(ent.hh_moi)} HH")
        fila(bool(ent.avance_items), "7. Itemizado", (f"{len(ent.avance_items)} ítems desde " + ", ".join(ent.fuentes_avance))
             if ent.avance_items else "sin planillas de avance: la semana queda sin avance")
        if sem.dias_sin_datos:
            fila(False, "Datos diarios", "faltan días en la BD Personal: " + ", ".join(c.fecha(d) for d in sem.dias_sin_datos))
        st.caption("Equipos, dashboards T&M, programa, 6 Weeks, fotos, NDC y RFI quedan como en el informe anterior. "
                   "Las HH gastadas por ítem del itemizado no se cambian.")
    
        if st.button(f"Generar informe semanal N° {ent.numero}", type="primary", key="btn_gen_sem"):
            try:
                with st.spinner("Generando…"):
                    datos, avisos = SM.generar(pl, ent)
            except Exception as e:
                st.error(f"No se pudo generar: {e}")
                return
            nombre = f"{ent.numero}. Informe Semanal {ent.inicio + dt.timedelta(days=6):%d-%m-%Y}_ Rev.A.xlsx"
            st.session_state["sem_generado"] = (nombre, datos, avisos)
        if st.session_state.get("sem_generado"):
            nombre, datos, avisos = st.session_state["sem_generado"]
            for a in avisos:
                st.warning(a)
            st.success(f"Informe listo: {nombre}")
            st.download_button("⬇️ Descargar Excel", datos, file_name=nombre, mime=MIME_XLSX, key="dl_sem")
            if st.button("Usar como base para la próxima semana", key="btn_usar_sem"):
                db.guardar_base_semanal(nombre, (ent.inicio + dt.timedelta(days=6)).isoformat(), ent.numero, datos,
                                        c.nombre_usuario(), "generado", contrato_id=cid)
                st.session_state.pop("sem_generado", None)
                st.cache_resource.clear()
                st.success("Listo: el informe generado es la nueva base.")
                st.rerun()

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


vsemanal = _modulo_vsemanal()

# ======================================================================
# MÓDULO: informe
# ======================================================================
def _modulo_informe():
    import datetime as dt
    
    import pandas as pd
    import plotly.express as px
    import streamlit as st
    
    DV = desviaciones
    CAL = calendario
    c = comun
    VS = vsemanal
    
    HOJA_PREFERIDA = "7.Avance EC-002 Itemizado"
    NADA = "(no aplica)"
    MODOS = {"actual": "HH ganadas / Actual acumulado", "remaining": "Remaining / Saldo", "avance": "% Avance"}
    
    
    def vista():
        st.title("Informe semanal")
        st.caption("**Crear informe**: arma el informe de la semana nueva sobre el último emitido · "
                   "**Avance y comparación**: curva de avance y diferencias entre dos semanas · "
                   "**Registrar en el historial**: guarda el itemizado de un informe para compararlo · "
                   "**Informes guardados**: lista y eliminación.")
        t_crear, t_hist, t_carga, t_lista = st.tabs(["📝 Crear informe", "📈 Avance y comparación",
                                                     "📥 Registrar en el historial", "🗂️ Informes guardados"])
        with t_crear:
            if c.exigir("cargar"):
                VS.crear()
        with t_carga:
            if c.exigir("cargar"):
                _cargar()
        with t_hist:
            _historial()
        with t_lista:
            _guardados()
    
    
    @st.cache_data(show_spinner="Leyendo hoja…", max_entries=6)
    def _leer(datos: bytes, hoja: str, fila: int | None):
        df, fila_det = analisis.leer_hoja(datos, hoja, fila)
        return df, fila_det, dict(df.attrs.get("columnas_excel", {})), df.attrs.get("filas_encabezado")
    
    
    @st.cache_data(show_spinner=False, max_entries=12)
    def _filas_suma(datos: bytes, hoja: str, columnas: tuple):
        return analisis.filas_suma_excel(datos, hoja, list(columnas))
    
    
    @st.cache_data(show_spinner=False, max_entries=3)
    def _hojas(datos: bytes):
        import io
        return pd.ExcelFile(io.BytesIO(datos)).sheet_names
    
    
    def _hoja_sugerida(hojas: list[str]) -> str:
        objetivo = analisis._norm(HOJA_PREFERIDA)
        exacta = [h for h in hojas if analisis._norm(h) == objetivo]
        if exacta:
            return exacta[0]
        itemizado = [h for h in hojas if "itemizado" in analisis._norm(h)]
        if itemizado:
            return min(itemizado, key=len)
        return next((h for h in hojas if "avance" in h.lower()), hojas[0])
    
    
    def _fecha_desde_nombre(nombre: str) -> dt.date | None:
        import re
        m = re.search(r"(\d{1,2})[-_.](\d{1,2})[-_.](20\d{2})", nombre)
        if m:
            try:
                return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                return None
        return None
    
    
    def _paso(n: int, titulo: str, detalle: str = ""):
        st.markdown(f"#### {n}. {titulo}")
        if detalle:
            st.caption(detalle)
    
    
    def _cargar():
        st.markdown("Guarda el avance por ítem (hoja de itemizado) de un informe semanal ya emitido. Con dos o más informes "
                    "guardados se arma la curva de avance y la comparación entre semanas.")
        _paso(1, "Sube el Excel del informe", f"La app busca sola la hoja **{HOJA_PREFERIDA}** (o la de itemizado más "
              "parecida), la fila de encabezado y las columnas. Solo cambia algo si la sugerencia no calza.")
        archivo = st.file_uploader("Excel del informe semanal", type=["xlsx", "xlsm", "xls"], key="up_inf")
        if archivo is None:
            return
        datos = archivo.getvalue()
        hojas = _hojas(datos)
        sugerida = _hoja_sugerida(hojas)
    
        _paso(2, "Revisa lo que se detectó")
        with st.expander("Hoja y fila de encabezado", expanded=False):
            hoja = st.selectbox("Hoja", hojas, index=hojas.index(sugerida), key="inf_hoja")
            _, fila_auto, _, _ = _leer(datos, hoja, None)
            fila = st.number_input("Fila del encabezado (1 = primera fila del Excel)", 1, 200, fila_auto + 1,
                                   key=f"inf_fila_{hoja}",
                                   help="Si el encabezado ocupa varias filas, indica la fila principal; las combinadas se leen solas.") - 1
        df, _, col_excel, filas_enc = _leer(datos, hoja, fila)
        if df.empty:
            st.warning("La hoja no tiene datos bajo esa fila de encabezado. Ábrela en «Hoja y fila de encabezado» y corrígela.")
            return
        enc = f"filas {filas_enc[0]} a {filas_enc[1]}" if filas_enc and filas_enc[0] != filas_enc[1] else f"fila {fila + 1}"
    
        cols = df.columns.tolist()
    
        def elegir(col, etiqueta, clave, opcional, key=None):
            sug = analisis.sugerir_columna(cols, clave)
            ops = ([NADA] if opcional else []) + cols
            v = col.selectbox(etiqueta, ops, index=ops.index(sug) if sug in ops else 0, key=key or f"inf_{clave}_{hoja}")
            return None if v == NADA else v
    
        exp_cols = st.expander("Columnas que se usan", expanded=False)
        with exp_cols:
            st.dataframe(df.head(15), use_container_width=True, height=220)
        a, b = exp_cols.columns(2)
        col_item = elegir(a, "Ítem / código", "item", False)
        col_desc = elegir(b, "Descripción (opcional, se agrega al ítem)", "descripcion", True)
        col_atc = elegir(a, "At Completion / presupuesto vigente (HH)", "at_completion", False)
        modos = [m for m in ("actual", "remaining", "avance") if analisis.sugerir_columna(cols, m)] + \
                [m for m in ("actual", "remaining", "avance") if not analisis.sugerir_columna(cols, m)]
        modo = b.selectbox("¿Cómo viene el avance?", modos, format_func=MODOS.get, key=f"inf_modo_{hoja}")
        col_medida = elegir(a, f"Columna de {MODOS[modo]}", modo, False, key=f"inf_med_{modo}_{hoja}")
        col_gas = elegir(b, "HH gastadas acumuladas (opcional, para el PF)", "gastadas", True)
        a, b = exp_cols.columns(2)
        excluir = a.checkbox("Excluir filas de total", value=True, key="inf_excl_tot")
        excluir_sub = b.checkbox("Excluir subtotales (fórmulas SUM o filas que suman otras)", value=True,
                                 key="inf_excl_sub",
                                 help="Evita contar dos veces las horas cuando la planilla tiene varios niveles de agrupación.")
        st.caption(f"Hoja **{hoja}** · encabezado en {enc} · ítem: **{col_item}** · At Completion: **{col_atc}** · "
                   f"avance: **{col_medida}** ({MODOS[modo]})" + (f" · HH gastadas: **{col_gas}**" if col_gas else ""))
    
        usadas = tuple(col_excel.get(x) for x in (col_atc, col_medida, col_gas) if x and col_excel.get(x))
        filas_suma = _filas_suma(datos, hoja, usadas) if excluir_sub else None
        try:
            items, avisos = analisis.normalizar_informe(
                df, col_item, col_atc,
                col_rem=col_medida if modo == "remaining" else None,
                col_actual=col_medida if modo == "actual" else None,
                col_avance=col_medida if modo == "avance" else None,
                excluir_totales=excluir, excluir_subtotales=excluir_sub,
                col_desc=col_desc, col_gastadas=col_gas, filas_suma=filas_suma)
        except Exception as e:  # columnas mal mapeadas
            st.error(f"No se pudieron interpretar las columnas elegidas: {e}")
            return
        for aviso in avisos:
            st.caption("⚠️ " + aviso)
        if items.empty:
            st.warning("No quedaron ítems con valores numéricos. Revisa el mapeo de columnas.")
            return
    
        exp_filas = st.expander("Filas que se suman (marca o desmarca «Incluir»)", expanded=False)
        clave_editor = f"inf_ed_{archivo.name}_{hoja}_{fila}_{col_item}_{col_atc}_{modo}_{col_medida}_{excluir}_{excluir_sub}"
        columnas_tabla = ["incluir", "fila_origen", "item", "at_completion", "actual", "remaining", "avance_pct"]
        if col_gas:
            columnas_tabla.append("hh_gastadas")
        columnas_tabla.append("motivo")
        editado = exp_filas.data_editor(
            items[columnas_tabla],
            hide_index=True, use_container_width=True, height=360, key=clave_editor,
            disabled=[x for x in columnas_tabla if x != "incluir"],
            column_config={"incluir": st.column_config.CheckboxColumn("Incluir", width="small"),
                           "fila_origen": st.column_config.NumberColumn("Fila Excel", format="%d", width="small"),
                           "item": st.column_config.TextColumn("Ítem", width="large"),
                           "actual": st.column_config.NumberColumn("HH ganadas / Actual", format="%.0f"),
                           "hh_gastadas": st.column_config.NumberColumn("HH gastadas", format="%.0f"),
                           "motivo": st.column_config.TextColumn("Excluida por", width="medium"),
                           **{k: v for k, v in c.COLUMNAS_HORAS.items() if k != "actual"}},
        )
        items = items.assign(incluir=editado["incluir"].astype(bool).values)
        finales = items[items["incluir"]]
    
        pf = analisis.productividad(finales)
        _paso(3, "Confirma los totales", "Compáralos con los totales de tu Excel antes de guardar.")
        k = st.columns(5 if pf is not None else 4)
        k[0].metric("Ítems incluidos", f"{len(finales)} de {len(items)}")
        k[1].metric("At Completion", f"{c.num(finales['at_completion'].sum())} HH")
        k[2].metric("HH ganadas", f"{c.num(finales['actual'].sum())} HH")
        k[3].metric("Avance ponderado por HH", c.porc(analisis.avance_ponderado(finales)))
        if pf is not None:
            k[4].metric("PF (gastadas / ganadas)", c.num(pf, 2), help=f"HH gastadas: {c.num(finales['hh_gastadas'].sum())}")
        if len(finales) < len(items):
            st.caption(f"Si se sumaran todas las filas el At Completion sería {c.num(items['at_completion'].sum())} HH; "
                       "las filas excluidas son totales o subtotales.")
    
        control = st.number_input("Total At Completion del Excel, para validar (opcional)", min_value=0.0, value=0.0,
                                  step=1.0, format="%.0f", key=f"inf_ctrl_{hoja}",
                                  help="Escribe el total que muestra la planilla; la app te dirá si cuadra.")
        if control > 0:
            dif = finales["at_completion"].sum() - control
            if abs(dif) <= 1:
                st.success(f"Cuadra con el total del Excel: {c.num(control)} HH.")
            else:
                st.warning(f"No cuadra: la suma incluida es {c.num(finales['at_completion'].sum())} HH, "
                           f"{'+' if dif > 0 else ''}{c.num(dif)} HH respecto del total del Excel. "
                           "Revisa las filas marcadas en la tabla o la columna elegida como At Completion.")
    
        resumen = _resumen_semanal(archivo.getvalue())
        _paso(4, "Guarda el informe en el historial")
        a, b, d = st.columns([2, 1, 2])
        fecha_corte = b.date_input("Fecha de corte", (resumen or {}).get("hasta") or _fecha_desde_nombre(archivo.name) or dt.date.today(),
                                   format="DD-MM-YYYY", key="inf_fecha")
        semana = CAL.semana(fecha_corte)
        b.caption(CAL.texto(fecha_corte))
        num = f"N°{resumen['numero']} " if resumen and resumen.get("numero") else ""
        nombre = a.text_input("Nombre del informe", f"Informe semanal {num}S{semana:02d} al {fecha_corte:%d-%m-%Y}", key="inf_nombre")
        ps = c.proyectos()
        opciones = [None] + ps["id"].tolist()
        proy_id = d.selectbox("Proyecto (opcional)", opciones, key="inf_proy",
                              format_func=lambda i: "Sin asociar" if i is None else ps.set_index("id").at[i, "codigo"])
        if st.button("Guardar en el historial", type="primary", key="btn_inf", disabled=finales.empty):
            inf_id = db.guardar_informe(nombre.strip() or f"Informe {fecha_corte}", None if proy_id is None else int(proy_id),
                                        fecha_corte.isoformat(), archivo.name, hoja, c.nombre_usuario(), finales,
                                        contrato_id=c.contrato_id())
            if resumen:
                db.guardar_config_diario(f"resumen_informe_{inf_id}", resumen_a_json(resumen))
            c.limpiar_cache()
            st.success(f"'{nombre}' guardado. Ya aparece en el historial.")
    
        if pf is not None or resumen:
            inf = _informes()
            previos = inf[pd.to_datetime(inf["fecha_corte"]).dt.date < fecha_corte] if not inf.empty else inf
            previo = None
            if not previos.empty:
                p_ = previos.iloc[-1]
                previo = (pd.Timestamp(p_["fecha_corte"]).date(), p_["hh_gastadas"], p_["at_completion"] - p_["remaining"])
            with st.expander("📉 Opcional: impacto de las desviaciones de recursos en el PF de este informe", expanded=False):
                cruce_desviaciones(fecha_corte, float(finales["hh_gastadas"].sum() if "hh_gastadas" in finales else 0),
                                   float(finales["actual"].sum()), previo, "carga", resumen)
    
    
    def _informes():
        return _informes_de(c.contrato_id())
    
    
    @st.cache_data(ttl=600, show_spinner=False)
    def _informes_de(cid: int):
        return db.informes(cid)
    
    
    @st.cache_data(ttl=600, show_spinner=False)
    def _items(ids: tuple):
        return db.informe_items(list(ids))
    
    
    def _historial():
        inf = _informes()
        if inf.empty:
            st.info("Aún no hay informes en el historial. Guárdalos en la pestaña **📥 Registrar en el historial**.")
            return
        inf = inf.copy()
        inf["avance_pct"] = [analisis.pct(t - r, t) for t, r in zip(inf["at_completion"], inf["remaining"])]
        inf["etiqueta"] = [f"{n} ({c.fecha(f)})" for n, f in zip(inf["nombre"], inf["fecha_corte"])]
    
        proyectos = sorted(inf["proyecto"].fillna("Sin asociar").unique())
        if len(proyectos) > 1:
            filtro = st.selectbox("Proyecto", proyectos, key="hist_proy")
            inf = inf[inf["proyecto"].fillna("Sin asociar") == filtro]
    
        # --- Situación al último informe
        ult = inf.sort_values("fecha_corte").iloc[-1]
        ant = inf.sort_values("fecha_corte").iloc[-2] if len(inf) > 1 else None
        st.subheader(f"Situación al {c.fecha(ult['fecha_corte'])}")
        st.caption(f"Último informe del historial: {ult['nombre']}")
        gan_u = ult["at_completion"] - ult["remaining"]
        k = st.columns(4)
        k[0].metric("Avance acumulado", c.porc(ult["avance_pct"]),
                    delta=None if ant is None else f"{c.num(ult['avance_pct'] - ant['avance_pct'], 2)} pp en la semana")
        k[1].metric("HH ganadas acumuladas", c.num(gan_u),
                    delta=None if ant is None else f"{c.num(gan_u - (ant['at_completion'] - ant['remaining']))} HH en la semana")
        k[2].metric("At Completion", f"{c.num(ult['at_completion'])} HH")
        pf_u = ult["hh_gastadas"] / gan_u if pd.notna(ult.get("hh_gastadas")) and gan_u > 0 else None
        k[3].metric("PF acumulado", c.num(pf_u, 2) if pf_u is not None else "–",
                    help="HH gastadas / HH ganadas. Sobre 1: se gastan más horas de las que se ganan.")
    
        st.subheader("Curva de avance")
        ev = inf.assign(Fecha=pd.to_datetime(inf["fecha_corte"]))
        fig = px.line(ev, x="Fecha", y="avance_pct", markers=True, text=[c.porc(v, 1) for v in ev["avance_pct"]],
                      hover_name="nombre", labels={"avance_pct": "% avance ponderado"})
        fig.update_traces(line=dict(color=c.NARANJO, width=3), textposition="top center")
        fig.update_yaxes(range=[0, 105])
        st.plotly_chart(c.layout_grafico(fig, 320), use_container_width=True, key="inf_curva_avance")
    
        st.subheader("Comparar dos semanas")
        st.caption("Por defecto se comparan los dos últimos informes. Cambia cualquiera de los dos para ver otro período.")
        ids = inf["id"].tolist()
        etiquetas = dict(zip(inf["id"], inf["etiqueta"]))
        a, b = st.columns(2)
        id_a = a.selectbox("Desde (informe anterior)", ids, index=max(0, len(ids) - 2), format_func=etiquetas.get, key="cmp_a")
        id_b = b.selectbox("Hasta (informe más reciente)", ids, index=len(ids) - 1, format_func=etiquetas.get, key="cmp_b")
        items = _items(tuple(sorted({int(id_a), int(id_b)})))
        ia = items[items["informe_id"] == id_a].groupby("item", sort=False)[["at_completion", "remaining"]].sum()
        ib = items[items["informe_id"] == id_b].groupby("item", sort=False)[["at_completion", "remaining"]].sum()
        comp = ia.join(ib, how="outer", lsuffix="_a", rsuffix="_b").fillna(0).reset_index()
        comp["avance_a"] = [analisis.pct(t - r, t) for t, r in zip(comp["at_completion_a"], comp["remaining_a"])]
        comp["avance_b"] = [analisis.pct(t - r, t) for t, r in zip(comp["at_completion_b"], comp["remaining_b"])]
        comp["hh_ganadas"] = (comp["at_completion_b"] - comp["remaining_b"]) - (comp["at_completion_a"] - comp["remaining_a"])
        comp["delta_pp"] = comp["avance_b"] - comp["avance_a"]
    
        tot_a = inf.set_index("id").loc[id_a]
        tot_b = inf.set_index("id").loc[id_b]
        def pf_de(tot):
            gan = tot["at_completion"] - tot["remaining"]
            return tot["hh_gastadas"] / gan if pd.notna(tot.get("hh_gastadas")) and gan > 0 else None
    
        pf_a, pf_b = pf_de(tot_a), pf_de(tot_b)
        k = st.columns(4 if pf_b is not None else 3)
        k[0].metric("Avance «desde»", c.porc(tot_a["avance_pct"]))
        k[1].metric("Avance «hasta»", c.porc(tot_b["avance_pct"]),
                    delta=f"{c.num(tot_b['avance_pct'] - tot_a['avance_pct'], 2)} pp")
        k[2].metric("HH ganadas en el período", c.num(comp["hh_ganadas"].sum()))
        if pf_b is not None:
            k[3].metric("PF acumulado", c.num(pf_b, 2),
                        delta=None if pf_a is None else f"{c.num(pf_b - pf_a, 2)}", delta_color="inverse",
                        help="HH gastadas / HH ganadas. Sobre 1: se gastan más horas de las que se ganan.")
    
        resumen_b = resumen_de_json(db.config_diario(f"resumen_informe_{int(id_b)}"))
        if pf_b is not None or resumen_b:
            fa, fb = pd.Timestamp(tot_a["fecha_corte"]).date(), pd.Timestamp(tot_b["fecha_corte"]).date()
            previo = (fa, tot_a["hh_gastadas"], tot_a["at_completion"] - tot_a["remaining"]) if fa < fb and pf_a is not None else None
            with st.expander("📉 Impacto de las desviaciones de recursos en el PF y recuento de la semana", expanded=False):
                cruce_desviaciones(fb, float(tot_b["hh_gastadas"] or 0) if pd.notna(tot_b["hh_gastadas"]) else 0.0,
                                   float(tot_b["at_completion"] - tot_b["remaining"]), previo, "hist", resumen_b)
    
        top = comp[comp["delta_pp"] != 0].nlargest(15, "hh_ganadas")
        if not top.empty:
            fig = px.bar(top.sort_values("hh_ganadas"), x="hh_ganadas", y="item", orientation="h",
                         labels={"hh_ganadas": "HH ganadas en el período", "item": ""},
                         color_discrete_sequence=[c.NARANJO], title="Ítems con más avance en el período")
            st.plotly_chart(c.layout_grafico(fig, max(300, 28 * len(top))), use_container_width=True, key="inf_top_items")
    
        st.markdown("**Detalle por ítem**")
        st.dataframe(
            comp[["item", "at_completion_b", "avance_a", "avance_b", "delta_pp", "hh_ganadas"]],
            hide_index=True, use_container_width=True, height=min(600, 38 + 35 * len(comp)),
            column_config={
                "item": st.column_config.TextColumn("Ítem", width="large"),
                "at_completion_b": st.column_config.NumberColumn("At Completion (HH)", format="%.0f"),
                "avance_a": st.column_config.NumberColumn("% desde", format="%.2f%%"),
                "avance_b": st.column_config.ProgressColumn("% hasta", format="%.2f%%", min_value=0, max_value=100),
                "delta_pp": st.column_config.NumberColumn("Δ pp", format="%+.2f"),
                "hh_ganadas": st.column_config.NumberColumn("HH ganadas", format="%.0f"),
            },
        )
        st.download_button("Descargar comparación (Excel)", analisis.a_excel({"Comparación": comp, "Informes": inf}),
                           file_name="comparacion_informes.xlsx", key="dl_cmp",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    
    
    def _guardados():
        inf = _informes()
        if inf.empty:
            st.info("Aún no hay informes en el historial.")
            return
        inf = inf.copy()
        inf["avance_pct"] = [analisis.pct(t - r, t) for t, r in zip(inf["at_completion"], inf["remaining"])]
        ids = inf["id"].tolist()
        etiquetas = {i: f"{n} ({c.fecha(f)})" for i, n, f in zip(inf["id"], inf["nombre"], inf["fecha_corte"])}
        st.dataframe(inf[["nombre", "fecha_corte", "proyecto", "items", "at_completion", "avance_pct", "archivo", "cargado_por"]],
                     hide_index=True, use_container_width=True,
                     column_config={"nombre": "Informe", "fecha_corte": "Fecha de corte", "proyecto": "Proyecto",
                                    "items": "Ítems", "archivo": "Archivo", "cargado_por": "Cargado por",
                                    **c.COLUMNAS_HORAS})
        if c.usuario() and c.usuario()["rol"] in ("admin", "planificador"):
            borrar = st.selectbox("Eliminar un informe", ids, format_func=etiquetas.get, key="inf_borrar")
            if st.button("Eliminar informe seleccionado", key="btn_inf_borrar"):
                db.borrar_informe(int(borrar))
                c.limpiar_cache()
                st.rerun()
    
    
    # ------------------------------------------------------------------ cruce con el LOG de desviaciones
    
    @st.cache_resource(show_spinner=False, max_entries=4)
    def _log_desviaciones(cid: int):
        return DV.LogDesviaciones.de_dict(db.carga_desviaciones(cid))
    
    
    @st.cache_data(show_spinner=False, max_entries=4)
    def _resumen_semanal(datos: bytes):
        try:
            return analisis.leer_resumen_semanal(datos)
        except Exception:
            return None
    
    
    def resumen_a_json(r: dict) -> dict:
        return {k: (v.isoformat() if isinstance(v, dt.date) else v) for k, v in r.items()}
    
    
    def resumen_de_json(r: dict | None) -> dict | None:
        if not r:
            return None
        out = dict(r)
        for k in ("desde", "hasta", "fecha_anterior", "fecha_actual"):
            if out.get(k):
                out[k] = dt.date.fromisoformat(str(out[k])[:10])
        return out
    
    
    def cruce_desviaciones(fecha_corte, gastadas: float, ganadas: float, previo, clave: str, resumen: dict | None = None):
        """PF real frente al PF sin las HH desviadas, para la semana del informe y el acumulado, y recuento de la semana.
    
        Si el Excel trae la tabla «Resumen» (acumulado anterior / período actual / acumulado actual), se usan sus HH ganadas y
        gastadas; si no, la semana se obtiene restando el informe anterior guardado."""
        st.markdown("#### Impacto de las desviaciones de recursos en el PF")
        try:
            cargas = db.cargas_desviaciones(c.contrato_id())
        except Exception:
            cargas = pd.DataFrame()
        if cargas.empty:
            st.info("Carga el LOG de desviación de recursos (Terreno → Desviación de recursos) para ver cuánto habría mejorado "
                    "el PF sin esas desviaciones.")
            return
        hasta = (resumen or {}).get("hasta") or fecha_corte
        cubre = cargas[cargas["fecha_corte"].fillna("") >= str(hasta)]
        cid = int((cubre.iloc[-1] if not cubre.empty else cargas.iloc[0])["id"])   # el LOG más antiguo que cubra la fecha
        log = _log_desviaciones(cid)
    
        fila = None
        if resumen and resumen.get("filas"):
            nombres = [f["descripcion"] for f in resumen["filas"]]
            defecto = nombres.index(resumen.get("total")) if resumen.get("total") in nombres else len(nombres) - 1
            f0, f1, f2 = st.columns([1.4, 1, 1])
            elegida = f0.selectbox("Fila del resumen del informe", nombres, index=defecto, key=f"cruce_fila_{clave}")
            fila = resumen["filas"][nombres.index(elegida)]
        else:
            f1, f2 = st.columns(2)
        esp = f1.multiselect("Especialidades a descontar (todas si vacío)", sorted(x for x in log.df["especialidad"].unique() if x),
                             key=f"cruce_esp_{clave}")
        mot = f2.multiselect("Motivos a descontar (todos si vacío)", sorted(x for x in log.df["referencia"].unique() if x),
                             key=f"cruce_mot_{clave}")
    
        semana, acum, desde = None, None, None
        if fila:
            desde = resumen.get("desde") or ((resumen.get("fecha_anterior") + dt.timedelta(days=1)) if resumen.get("fecha_anterior") else None)
            per, acu = fila.get("periodo"), fila.get("acumulado")
            if per and desde:
                semana = DV.impacto_pf(log, per["gastadas"], per["ganadas"], hasta, desde - dt.timedelta(days=1), esp, mot)
            if acu:
                acum = DV.impacto_pf(log, acu["gastadas"], acu["ganadas"], hasta, None, esp, mot)
            _tabla_resumen(resumen, fila, semana, acum)
        else:
            if gastadas and ganadas:
                acum = DV.impacto_pf(log, gastadas, ganadas, hasta, None, esp, mot)
            if previo is not None:
                f_prev, g_prev, e_prev = previo
                dg, de = gastadas - float(g_prev or 0), ganadas - float(e_prev or 0)
                if g_prev is not None and not pd.isna(g_prev) and de > 0 and dg > 0:
                    semana = DV.impacto_pf(log, dg, de, hasta, f_prev, esp, mot)
                    desde = f_prev + dt.timedelta(days=1)
                else:
                    st.caption("No se pudo calcular el PF de la semana: las HH ganadas o gastadas no aumentaron respecto del informe anterior.")
            else:
                st.caption("Para el PF de la semana, el Excel debe traer la tabla «Resumen» del informe o debe existir el informe "
                           "de la semana anterior en el historial.")
        if not semana and not acum:
            st.info("El informe no trae HH gastadas y ganadas para calcular el PF.")
            return
        st.caption(f"LOG usado: {log.archivo} (al {c.fecha(log.fecha_corte)}). PF sin desviaciones = (HH gastadas − HH desviadas) / "
                   "HH ganadas: las horas desviadas se gastaron en trabajos fuera de alcance que no generan avance.")
    
        bloques = []
        if semana:
            n = CAL.semana(hasta)
            num = f"informe N°{resumen['numero']} · " if resumen and resumen.get("numero") else ""
            bloques.append((f"Semana S{n:02d} ({num}{c.fecha(desde)} al {c.fecha(hasta)})", semana))
        if acum:
            bloques.append((f"Acumulado al {c.fecha(hasta)}", acum))
        cols = st.columns(len(bloques))
        for col, (titulo, r) in zip(cols, bloques):
            with col:
                st.markdown(f"**{titulo}**")
                k = st.columns(3)
                k[0].metric("PF real", c.num(r["pf"], 2))
                k[1].metric("PF sin desviaciones", c.num(r["pf_sin"], 2), delta=f"-{c.num(r['mejora'], 2)}", delta_color="inverse",
                            help="Menor es mejor: PF bajo 1 significa que se gana más de lo que se gasta.")
                k[2].metric("HH desviadas", c.num(r["hh_desviadas"]), help=f"{c.porc(r['pct_gastadas'], 1)} de las HH gastadas · "
                                                                              f"{c.num(r['registros'])} registros del LOG")
                st.caption(DV.texto_impacto(r, "En el período"))
                for av in r["avisos"]:
                    st.caption("⚠️ " + av)
    
        rec = None
        if semana and desde:
            rec = DV.recuento_semana(log, desde, hasta, esp, mot)
            _recuento(rec, semana, clave)
        else:
            m = acum["por_motivo"]
            m = m[m["hh"] > 0]
            if not m.empty:
                _grafico_puntos(m, "acumulado", clave)
    
        tablas = {}
        if semana:
            tablas.update({"Semana por motivo": semana["por_motivo"], "Semana por especialidad": semana["por_especialidad"]})
        if acum:
            tablas.update({"Acumulado por motivo": acum["por_motivo"], "Acumulado por especialidad": acum["por_especialidad"]})
        if rec is not None and not rec["d"].empty:
            tablas.update({"Semana por día": rec["diario"], "Semana detalle": DV.tabla_detalle(rec["d"])})
        resumen_xl = pd.DataFrame([{"Período": t, "HH gastadas": r["gastadas"], "HH ganadas": r["ganadas"], "PF real": r["pf"],
                                    "HH desviadas": r["hh_desviadas"], "PF sin desviaciones": r["pf_sin"], "Mejora": r["mejora"]}
                                   for t, r in bloques])
        st.download_button("Descargar el cruce y el recuento en Excel", analisis.a_excel({"Resumen": resumen_xl, **tablas}),
                           file_name=f"Cruce_PF_desviaciones_{hasta:%d-%m-%Y}.xlsx", key=f"cruce_dl_{clave}",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    
    def _tabla_resumen(resumen: dict, fila: dict, semana, acum):
        """Sector de avance del informe (acumulado anterior / período actual / acumulado actual) con el PF sin desviaciones."""
        nombres = {"anterior": f"Acumulado anterior ({c.fecha(resumen.get('fecha_anterior'))})",
                   "periodo": f"Período actual ({c.fecha(resumen.get('desde'))} al {c.fecha(resumen.get('hasta'))})",
                   "acumulado": f"Acumulado actual ({c.fecha(resumen.get('fecha_actual') or resumen.get('hasta'))})"}
        filas = []
        for f in resumen["filas"]:
            r = {"Descripción": f["descripcion"]}
            for k, titulo in nombres.items():
                v = f.get(k)
                if v:
                    r[f"{titulo} · HH ganadas"] = v["ganadas"]
                    r[f"{titulo} · HH gastadas"] = v["gastadas"]
                    r[f"{titulo} · PF"] = v["pf"]
            filas.append(r)
        t = pd.DataFrame(filas)
        st.markdown(f"**Avance del informe{' N°' + str(resumen['numero']) if resumen.get('numero') else ''}** "
                    f"(tabla «Resumen» de la hoja {resumen.get('hoja', '')})")
        fmt = {col: st.column_config.NumberColumn(col, format="%.2f" if col.endswith("PF") else "%.0f")
               for col in t.columns if col != "Descripción"}
        st.dataframe(t, hide_index=True, use_container_width=True, column_config=fmt)
        partes = []
        if semana:
            partes.append(f"PF del período {c.num(semana['pf'], 2)} → **{c.num(semana['pf_sin'], 2)}** sin desviaciones")
        if acum:
            partes.append(f"PF acumulado {c.num(acum['pf'], 2)} → **{c.num(acum['pf_sin'], 2)}**")
        if partes:
            st.caption(f"Para «{fila['descripcion']}»: " + " · ".join(partes) + ".")
    
    
    def _grafico_puntos(m: pd.DataFrame, alcance: str, clave: str):
        m = m.sort_values("puntos_pf")
        fig = px.bar(m, x="puntos_pf", y="referencia", orientation="h",
                     text=[f"{v:.3f}".replace(".", ",") for v in m["puntos_pf"]],
                     title=f"Puntos de PF que aporta cada motivo ({alcance})",
                     color_discrete_sequence=[c.NARANJO], labels={"puntos_pf": "Puntos de PF", "referencia": ""})
        st.plotly_chart(c.layout_grafico(fig, max(280, 30 * len(m))), use_container_width=True, key=f"cruce_puntos_{alcance}_{clave}")
    
    
    def _recuento(rec: dict, semana: dict, clave: str):
        st.markdown(f"#### Recuento de la semana ({c.fecha(rec['desde'])} al {c.fecha(rec['hasta'])})")
        ind = rec["ind"]
        k = st.columns(4)
        delta = None
        if rec["hh_prev"]:
            delta = f"{c.num((ind['hh'] / rec['hh_prev'] - 1) * 100, 1)}% vs semana anterior"
        k[0].metric("HH desviadas", c.num(ind["hh"]), delta=delta, delta_color="inverse")
        k[1].metric("Registros", c.num(ind["registros"]))
        k[2].metric("Días con desviaciones", f"{ind['dias_con_registro']} de {ind['dias']}")
        k[3].metric("Día con más HH", c.fecha(ind["pico_fecha"]) if ind["pico_fecha"] else "–",
                    help=f"{c.num(ind['pico_hh'], 1)} HH" if ind["pico_fecha"] else None)
        for frase in rec["texto"]:
            st.markdown(f"- {frase}")
        if rec["d"].empty:
            return
        a, b = st.columns(2)
        with a:
            dia = rec["diario"]
            jornadas = [x for x in dia.columns if x not in ("fecha", "Total", "Acumulado")]
            fig = px.bar(dia, x="fecha", y=jornadas, barmode="stack", title="HH desviadas por día de la semana",
                         labels={"value": "HH", "fecha": "", "variable": ""},
                         color_discrete_map={"Día": c.NARANJO, "Noche": c.AZUL, "General": c.PIZARRA})
            fig.update_xaxes(tickformat="%a %d-%m")
            st.plotly_chart(c.layout_grafico(fig, 320), use_container_width=True, key=f"cruce_dias_{clave}")
        with b:
            m = semana["por_motivo"]
            m = m[m["hh"] > 0]
            if not m.empty:
                _grafico_puntos(m, "semana", clave)
        a, b = st.columns(2)
        num = {"hh": st.column_config.NumberColumn("HH", format="%.1f"), "participacion": st.column_config.NumberColumn("%", format="%.1f%%"),
               "registros": "Registros", "puntos_pf": st.column_config.NumberColumn("Puntos de PF", format="%.3f")}
        with a:
            st.markdown("**Por motivo**")
            st.dataframe(semana["por_motivo"][["referencia", "hh", "participacion", "registros", "puntos_pf"]], hide_index=True,
                         use_container_width=True, column_config={"referencia": "Motivo", **num})
        with b:
            st.markdown("**Por especialidad**")
            st.dataframe(semana["por_especialidad"][["especialidad", "hh", "participacion", "registros", "puntos_pf"]], hide_index=True,
                         use_container_width=True, column_config={"especialidad": "Especialidad", **num})
        st.markdown("**Desviaciones más grandes de la semana**")
        st.dataframe(DV.tabla_detalle(rec["mayores"]), hide_index=True, use_container_width=True,
                     column_config={"Fecha": st.column_config.DateColumn(format="DD-MM-YYYY"),
                                    "Actividad": st.column_config.TextColumn(width="large")})

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


informe = _modulo_informe()

# ======================================================================
# MÓDULO: wla
# ======================================================================
def _modulo_wla():
    import datetime as dt
    
    import pandas as pd
    import plotly.express as px
    import streamlit as st
    
    CAL = calendario
    c = comun
    
    ESTADOS = ["Pendiente", "En gestión", "Liberada"]
    LIBERADAS = {"liberada", "liberado", "resuelto", "resuelta", "completado", "completada", "cerrada", "cerrado"}
    CAMPOS = {  # campo interno: (etiqueta, obligatorio, palabras clave)
        "proyecto": ("Proyecto", False, ["proyecto"]),
        "actividad": ("Actividad", True, ["actividad", "activity", "tarea"]),
        "subsistema": ("Subsistema", False, ["subsistema", "sistema"]),
        "semana": ("Semana", False, ["semana", "week"]),
        "restriccion": ("Restricción", True, ["restriccion", "restricción", "constraint"]),
        "responsable": ("Responsable", False, ["responsable"]),
        "estado": ("Estado", True, ["estado", "status"]),
        "fecha_compromiso": ("Fecha compromiso", False, ["compromiso", "fecha"]),
    }
    NADA = "(no aplica)"
    
    
    def _es_liberada(s: pd.Series) -> pd.Series:
        return s.fillna("").str.strip().str.lower().isin(LIBERADAS)
    
    
    def vista():
        st.title("Lookahead 6 semanas (6WLA)")
        hoy = dt.date.today()
        n = CAL.semana(hoy)
        if n and n >= 1:
            st.caption(f"Semana actual del proyecto: **S{n:02d}** ({CAL.inicio_semana(n):%d-%m} al {CAL.fin_semana(n):%d-%m}) · "
                       f"ventana 6WLA: **S{n:02d} a S{n + 5:02d}** (hasta el {CAL.fin_semana(n + 5):%d-%m-%Y})")
        t1, t2 = st.tabs(["Control de restricciones", "Cargar planilla"])
        with t2:
            if c.exigir("cargar"):
                _cargar()
        with t1:
            _control()
    
    
    def _restricciones():
        return _restricciones_de(c.contrato_id())
    
    
    @st.cache_data(ttl=600, show_spinner=False)
    def _restricciones_de(cid: int):
        return db.restricciones(cid)
    
    
    def _control():
        df = _restricciones()
        if df.empty:
            st.info("No hay restricciones registradas. Carga tu planilla en la pestaña **Cargar planilla**.")
            return
        proyectos = sorted(df["proyecto"].unique())
        proy = st.selectbox("Proyecto", proyectos, key="wla_proy")
        df = df[df["proyecto"] == proy].copy()
        lib = _es_liberada(df["estado"])
        hoy = dt.date.today().isoformat()
        vencidas = (~lib) & df["fecha_compromiso"].notna() & (df["fecha_compromiso"].astype(str).str[:10] < hoy)
    
        k = st.columns(4)
        k[0].metric("Restricciones", len(df))
        k[1].metric("Liberadas", int(lib.sum()))
        k[2].metric("Pendientes", int((~lib).sum()), delta=f"{int(vencidas.sum())} vencidas" if vencidas.any() else None,
                    delta_color="inverse")
        k[3].metric("Tasa de liberación", c.porc(analisis.pct(lib.sum(), len(df)), 1))
    
        g1, g2 = st.columns(2)
        with g1:
            fig = px.histogram(df.assign(semana=df["semana"].fillna("Sin semana")), x="semana", color="estado",
                               barmode="stack", title="Restricciones por semana",
                               labels={"semana": "Semana", "count": "Restricciones"},
                               color_discrete_sequence=[c.NARANJO, c.PIZARRA, c.VERDE, c.AZUL, c.GRIS])
            fig.update_xaxes(categoryorder="category ascending")
            st.plotly_chart(c.layout_grafico(fig, 330), use_container_width=True, key="g_wla_1")
        with g2:
            pend = df[~lib].assign(responsable=df["responsable"].fillna("Sin responsable"))
            if not pend.empty:
                r = pend["responsable"].value_counts().rename_axis("Responsable").reset_index(name="Pendientes")
                fig = px.bar(r.sort_values("Pendientes"), x="Pendientes", y="Responsable", orientation="h",
                             title="Pendientes por responsable", color_discrete_sequence=[c.PIZARRA])
                st.plotly_chart(c.layout_grafico(fig, 330), use_container_width=True, key="g_wla_2")
    
        st.subheader("Detalle")
        solo_pend = st.toggle("Mostrar solo pendientes", value=False)
        vista_df = df[~lib] if solo_pend else df
        puede_editar = c.usuario() and c.usuario()["rol"] in ("admin", "planificador")
        if puede_editar:
            st.caption("Puedes editar **Estado**, **Responsable** y **Fecha compromiso** directamente en la tabla y luego guardar.")
        estados_opciones = sorted(set(ESTADOS) | set(df["estado"].dropna().unique()))
        editado = st.data_editor(
            vista_df[["id", "actividad", "subsistema", "semana", "restriccion", "responsable", "estado", "fecha_compromiso",
                      "actualizado_por", "actualizado_en"]],
            hide_index=True, use_container_width=True, key=f"wla_editor_{proy}_{solo_pend}",
            disabled=not puede_editar or ["id", "actividad", "subsistema", "semana", "restriccion", "actualizado_por", "actualizado_en"],
            column_config={
                "id": None, "actividad": st.column_config.TextColumn("Actividad", width="medium"),
                "subsistema": "Subsistema", "semana": "Semana",
                "restriccion": st.column_config.TextColumn("Restricción", width="large"),
                "responsable": "Responsable",
                "estado": st.column_config.SelectboxColumn("Estado", options=estados_opciones, required=True),
                "fecha_compromiso": st.column_config.TextColumn("Fecha compromiso", help="AAAA-MM-DD"),
                "actualizado_por": "Actualizado por", "actualizado_en": "Actualizado",
            },
        )
        if puede_editar and st.button("Guardar cambios", type="primary", key="wla_guardar"):
            def limpio(v):
                return None if v is None or (not isinstance(v, str) and pd.isna(v)) or str(v).strip() == "" else str(v).strip()
    
            orig = vista_df.set_index("id")
            cambios = []
            for r in editado.itertuples():
                o = orig.loc[r.id]
                nuevo = (limpio(r.estado), limpio(r.responsable), limpio(r.fecha_compromiso))
                if nuevo != (limpio(o["estado"]), limpio(o["responsable"]), limpio(o["fecha_compromiso"])):
                    cambios.append((r.id, nuevo[0] or "Pendiente", nuevo[1], nuevo[2]))
            if cambios:
                db.actualizar_restricciones(cambios, c.nombre_usuario())
                c.limpiar_cache()
                st.success(f"{len(cambios)} restricciones actualizadas.")
                st.rerun()
            else:
                st.info("No hay cambios que guardar.")
    
        st.download_button("Descargar Excel", analisis.a_excel({"6WLA": df.drop(columns=["id"])}),
                           file_name=f"6WLA_{proy}.xlsx", key="wla_dl",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if c.usuario() and c.usuario()["rol"] == "admin":
            with st.expander("Eliminar restricciones del proyecto"):
                if st.checkbox(f"Confirmo que quiero borrar todas las restricciones de {proy}", key="wla_conf"):
                    if st.button("Eliminar", key="wla_borrar"):
                        db.borrar_restricciones(proy, contrato_id=c.contrato_id())
                        c.limpiar_cache()
                        st.rerun()
    
    
    def _cargar():
        archivo = st.file_uploader("Planilla 6WLA", type=["xlsx", "xlsm", "xls"], key="up_6wla")
        if archivo is None:
            st.caption("Columnas esperadas: Actividad, Restricción, Estado (obligatorias) y opcionalmente Proyecto, "
                       "Subsistema, Semana, Responsable, Fecha compromiso.")
            return
        datos = archivo.getvalue()
        hojas = pd.ExcelFile(archivo).sheet_names
        hoja = st.selectbox("Hoja", hojas, key="wla_hoja")
        _, fila = analisis.leer_hoja(datos, hoja)
        fila = st.number_input("Fila del encabezado", 1, 200, fila + 1, key=f"wla_fila_{hoja}") - 1
        df, _ = analisis.leer_hoja(datos, hoja, fila)
        st.dataframe(df.head(8), use_container_width=True)
        cols = df.columns.tolist()
        norm = {col: analisis._norm(col) for col in cols}
    
        mapeo = {}
        grid = st.columns(2)
        for i, (campo, (etiqueta, obligatorio, claves)) in enumerate(CAMPOS.items()):
            sug = next((col for k in claves for col in cols if analisis._norm(k) in norm[col]), None)
            ops = ([] if obligatorio else [NADA]) + cols
            mapeo[campo] = grid[i % 2].selectbox(etiqueta + (" *" if obligatorio else ""), ops,
                                                 index=ops.index(sug) if sug in ops else 0, key=f"wla_map_{campo}")
        proy_fijo = None
        if mapeo["proyecto"] == NADA:
            ps = c.proyectos()
            sugerido = ps["codigo"].iloc[0] if not ps.empty else ""
            proy_fijo = st.text_input("Proyecto para todas las filas", sugerido, key="wla_proy_fijo")
    
        salida = pd.DataFrame(index=df.index)
        for campo, col in mapeo.items():
            salida[campo] = None if col == NADA else df[col]
        if proy_fijo is not None:
            salida["proyecto"] = proy_fijo.strip() or "Sin proyecto"
        salida = salida[salida["actividad"].notna() & salida["restriccion"].notna()]
        salida["estado"] = salida["estado"].fillna("Pendiente").astype(str).str.strip()
        if mapeo["fecha_compromiso"] != NADA:
            f = pd.to_datetime(salida["fecha_compromiso"], errors="coerce", dayfirst=True)
            salida["fecha_compromiso"] = f.dt.strftime("%Y-%m-%d").where(f.notna(), None)
        st.caption(f"{len(salida)} restricciones válidas para importar.")
        if st.button("Importar restricciones", type="primary", key="wla_imp") and len(salida):
            db.guardar_restricciones(salida, c.nombre_usuario(), contrato_id=c.contrato_id())
            c.limpiar_cache()
            st.success(f"{len(salida)} restricciones importadas.")

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


wla = _modulo_wla()

# ======================================================================
# MÓDULO: vdiario
# ======================================================================
def _modulo_vdiario():
    import datetime as dt
    import re
    
    import pandas as pd
    import streamlit as st
    
    D = diario
    CAL = calendario
    c = comun
    
    MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    OPCIONES_ESTADO = list(D.ESTADOS)
    
    
    def _limpio(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        if isinstance(v, str) and not v.strip():
            return None
        return v
    
    
    def _filas_limpias(registros: list[dict]) -> list[dict]:
        salida = []
        for r in registros:
            r = {k: _limpio(v) for k, v in r.items()}
            if any(v is not None for v in r.values()):
                salida.append(r)
        return salida
    
    
    # ------------------------------------------------------------------ borrador del informe (guardar y continuar)
    
    _NO_GUARDAR = ("_gen", "_pdf_btn", "_emitir", "_dl", "_dl_pdf", "_aj_btn", "_aj_rut", "_aj_est", "_xlsx", "_pdf")
    
    
    def _a_json(v):
        if v is None or isinstance(v, (str, bool)):
            return v
        if isinstance(v, (int, float)):
            return None if isinstance(v, float) and pd.isna(v) else v
        if hasattr(v, "item"):                                  # numpy
            return _a_json(v.item())
        if isinstance(v, (dt.date, dt.datetime, pd.Timestamp)):
            return None if pd.isna(v) else v.isoformat()
        if isinstance(v, dt.time):
            return v.strftime("%H:%M")
        return str(v)
    
    
    def _clave_borrador(clave: str) -> str:
        return "borrador_" + clave
    
    
    def _restaurar_borrador(clave: str):
        """Una vez por sesión: si hay un borrador guardado de este informe, vuelve a poner lo que se había avanzado."""
        marca = f"{clave}__restaurado"
        if st.session_state.get(marca):
            return
        st.session_state[marca] = True
        b = db.config(c.contrato_id(), _clave_borrador(clave), None)
        if not b:
            return
        for k, v in (b.get("widgets") or {}).items():
            if k not in st.session_state:
                st.session_state[k] = v
        if b.get("ajustes") is not None:
            st.session_state[f"{clave}_ajustes"] = dict(b["ajustes"])
        st.session_state[f"{clave}__editores"] = b.get("editores") or {}
        st.session_state[f"{clave}__texto"] = b.get("texto") or {}
        st.session_state[f"{clave}__info"] = {"guardado": b.get("guardado"), "usuario": b.get("usuario"), "restaurado": True}
    
    
    def _guardar_borrador(clave: str):
        widgets = {}
        for k, v in st.session_state.items():
            k = str(k)
            if not k.startswith(clave + "_") or "__" in k or "_ed_" in k or "_texto_" in k or k.endswith(_NO_GUARDAR):
                continue
            if isinstance(v, (str, int, float, bool)) and not (isinstance(v, float) and pd.isna(v)):
                widgets[k] = v
        editores = {k.split("__res__", 1)[1]: v for k, v in st.session_state.items()
                    if str(k).startswith(f"{clave}__res__")}
        ahora = dt.datetime.now().strftime("%d-%m-%Y %H:%M")
        db.guardar_config(c.contrato_id(), _clave_borrador(clave), {
            "widgets": widgets, "editores": editores, "ajustes": st.session_state.get(f"{clave}_ajustes") or {},
            "texto": st.session_state.get(f"{clave}__texto_actual") or {}, "guardado": ahora, "usuario": c.nombre_usuario()})
        st.session_state[f"{clave}__info"] = {"guardado": ahora, "usuario": c.nombre_usuario()}
    
    
    def _borrar_borrador(clave: str):
        db.guardar_config(c.contrato_id(), _clave_borrador(clave), None)
        for k in [k for k in st.session_state if str(k).startswith(clave)]:
            del st.session_state[k]
    
    
    def _aplicar_borrador(df: pd.DataFrame, guardado: list[dict], id_col: str | None) -> pd.DataFrame:
        df = df.reset_index(drop=True).copy()
        if id_col and id_col in df.columns and guardado and all(id_col in r for r in guardado):
            usados = set()
            for r in guardado:
                filas = [i for i in df.index[df[id_col].astype(str) == str(r[id_col])] if i not in usados]
                if not filas:
                    continue
                i = filas[0]
                usados.add(i)
                for col, v in r.items():
                    if col in df.columns and col != id_col:
                        try:
                            df.at[i, col] = v
                        except (ValueError, TypeError):
                            df[col] = df[col].astype(object)
                            df.at[i, col] = v
            return df
        nuevo = pd.DataFrame([{col: r.get(col) for col in df.columns} for r in guardado], columns=df.columns)
        for col in df.columns:
            if pd.api.types.is_bool_dtype(df[col]):
                nuevo[col] = nuevo[col].fillna(False).astype(bool)
            elif pd.api.types.is_numeric_dtype(df[col]):
                nuevo[col] = pd.to_numeric(nuevo[col], errors="coerce")
        return nuevo
    
    
    def _editor(clave: str, nombre: str, df: pd.DataFrame, id_col: str | None = None, **kw) -> pd.DataFrame:
        """st.data_editor que retoma lo guardado en el borrador y deja su resultado listo para guardar."""
        guardado = (st.session_state.get(f"{clave}__editores") or {}).get(nombre)
        if guardado is not None:
            df = _aplicar_borrador(df, guardado, id_col)
        ed = st.data_editor(df, key=f"{clave}_ed_{nombre}", **kw)
        st.session_state[f"{clave}__res__{nombre}"] = [{k: _a_json(v) for k, v in r.items()} for r in ed.to_dict("records")]
        return ed
    
    
    def _barra_borrador(clave: str):
        info = st.session_state.get(f"{clave}__info")
        a, b, d = st.columns([1.2, 1.2, 3])
        if a.button("💾 Guardar avance", key=f"{clave}__btn_guardar", type="primary",
                    help="Guarda lo que llevas del informe. Si se recarga la página o se cierra la app, continúas desde aquí."):
            _guardar_borrador(clave)
            st.toast("Avance guardado")
            info = st.session_state.get(f"{clave}__info")
        if info and b.button("Descartar borrador", key=f"{clave}__btn_descartar",
                             help="Borra lo guardado y vuelve a partir desde los reportes."):
            _borrar_borrador(clave)
            st.rerun()
        if info:
            txt = f"Último guardado: {info['guardado']} por {info['usuario']}"
            d.caption(("♻️ Se retomó el avance guardado. " if info.get("restaurado") else "💾 ") + txt)
        else:
            d.caption("Aún no has guardado este informe. Usa «Guardar avance» para no perder lo que llevas si se recarga la página.")
    
    
    def vista():
        st.title("Informe diario")
        if not c.exigir("cargar"):
            return
        base = db.base_diario_vigente(c.contrato_id())
        t_gen, t_rep, t_base, t_hist = st.tabs(["Revisar y generar", "Reportes de terreno", "Planilla base", "Historial"])
        with t_base:
            _planilla_base(base)
        if not base:
            with t_gen:
                st.info("Primero carga la planilla del último informe emitido en la pestaña **Planilla base**.")
            return
        pl = _plantilla(int(base["id"]))
        fecha = _selector_fecha(pl)
        with t_rep:
            _reportes(fecha, pl)
        with t_gen:
            _generar(base, pl, fecha)
        with t_hist:
            _historial()
    
    
    @st.cache_resource(show_spinner="Leyendo la planilla base…", max_entries=2)
    def _plantilla(archivo_id: int):
        return D.leer_plantilla(db.archivo_diario(archivo_id))
    
    
    def _selector_fecha(pl) -> dt.date:
        sugerida = pl.fecha + dt.timedelta(days=1)
        with st.sidebar:
            st.markdown("**Informe diario**")
            fecha = st.date_input("Fecha del informe a generar", sugerida, format="DD-MM-YYYY", key="diario_fecha")
            st.caption(CAL.texto(fecha))
        return fecha
    
    
    # ------------------------------------------------------------------ planilla base
    
    def _planilla_base(base):
        st.markdown(
            "La planilla base es el **último informe diario emitido** (el .xlsx oficial con sus hojas BD). El informe nuevo se "
            "construye sobre ella: se conservan intactos el formato, los logos, las firmas y las fórmulas. Cada vez que emites un "
            "informe desde aquí, ese archivo pasa a ser la base del día siguiente.")
        if base:
            pl = _plantilla(int(base["id"]))
            k = st.columns(4)
            k[0].metric("Informe base", f"N° {pl.numero}")
            k[1].metric("Fecha", c.fecha(pl.fecha))
            k[2].metric("Plantel en BD Personal", c.num(len(pl.personas)))
            k[3].metric("Cargos en el informe", len(pl.cargos))
            t = pl.totales
            st.caption(f"Directos en faena: {t['MOD']['R']} día / {t['MOD']['U']} noche ({c.num(t['MOD']['X'])} HH) · "
                       f"Indirectos: {t['MOI']['R']} día / {t['MOI']['U']} noche ({c.num(t['MOI']['X'])} HH) · "
                       f"Origen: {'generado por el sistema' if base['tipo'] == 'generado' else 'cargado'} el {base['creado_en'][:16]}")
            st.download_button("Descargar planilla base", db.archivo_diario(int(base["id"])), file_name=base["nombre"],
                               mime=MIME_XLSX, key="dl_base")
        archivo = st.file_uploader("Cargar o reemplazar la planilla base (.xlsx del último informe emitido)", type=["xlsx", "xlsm"],
                                   key="up_base_diario")
        if archivo is not None:
            try:
                pl = D.leer_plantilla(archivo.getvalue())
            except Exception as e:
                st.error(f"No se pudo leer la planilla: {e}")
                return
            st.success(f"Informe N° {pl.numero} del {c.fecha(pl.fecha)} · {len(pl.personas)} personas en BD Personal · "
                       f"{len(pl.cargos)} cargos · {len(pl.equipos)} equipos · {len(pl.fotos)} fotos.")
            if st.button("Usar como planilla base", type="primary", key="btn_base"):
                db.guardar_base_diario(archivo.name, pl.fecha.isoformat(), pl.numero, archivo.getvalue(), c.nombre_usuario(), "base",
                                       contrato_id=c.contrato_id())
                st.cache_data.clear()
                st.rerun()
    
    
    # ------------------------------------------------------------------ reportes
    
    def _reportes(fecha: dt.date, pl):
        st.markdown(f"Reportes del **{c.fecha(fecha)}**. Sube todos los archivos del día de una vez: reportes de terreno de "
                    "cada especialidad y turno (aunque traigan una hoja por capataz), uso de recurso (UDR), adicionales y "
                    "programación de actividades. La app reconoce el formato de cada hoja por su contenido.")
        archivos = st.file_uploader("Reportes diarios de terreno (.xlsx)", type=["xlsx", "xlsm"], accept_multiple_files=True,
                                    key=f"up_rep_{fecha}")
        if archivos and st.button(f"Agregar {len(archivos)} reporte(s)", type="primary", key="btn_rep"):
            errores, resumen = [], []
            for a in archivos:
                try:
                    leidos, notas = D.leer_reportes(a.getvalue(), a.name)
                    for r in leidos:
                        db.guardar_reporte_diario(fecha.isoformat(), r.archivo, r.a_dict(), c.nombre_usuario(),
                                                  contrato_id=c.contrato_id())
                    tipos = {}
                    for r in leidos:
                        tipos[D.nombre_tipo(r)] = tipos.get(D.nombre_tipo(r), 0) + 1
                    resumen.append(f"**{a.name}**: " + ", ".join(f"{v} {k.lower()}" for k, v in tipos.items()) +
                                   (" · " + " · ".join(notas) if notas else ""))
                except Exception as e:
                    errores.append(f"{a.name}: {e}")
            st.session_state[f"rep_resumen_{fecha}"] = resumen
            for e in errores:
                st.error(e)
            if not errores:
                st.rerun()
        with st.expander("📦 PM de materiales (PDF) → 8.0 Materiales y consumibles", expanded=False):
            st.caption("Sube todos los PM del día de una vez. Cada material de cada PDF pasa a la sección 8.0 del informe "
                       "(descripción, unidad, cantidad, Activity ID / trabajo, solicitante y N° de PM). Si son más filas de "
                       "las que trae la planilla, la sección se amplía.")
            pms = st.file_uploader("PM de materiales (.pdf)", type=["pdf"], accept_multiple_files=True, key=f"up_pm_{fecha}")
            if pms and st.button(f"Agregar {len(pms)} PM", type="primary", key="btn_pm"):
                errores, resumen = [], []
                for a in pms:
                    try:
                        r = D.leer_pm_materiales(a.getvalue(), a.name)
                        if r.fecha and r.fecha != fecha:
                            resumen.append(f"⚠️ **{a.name}** tiene fecha {c.fecha(r.fecha)}; se agrega igual al {c.fecha(fecha)}.")
                        db.guardar_reporte_diario(fecha.isoformat(), r.archivo, r.a_dict(), c.nombre_usuario(),
                                                  contrato_id=c.contrato_id())
                        resumen.append(f"**{a.name}**: {len(r.materiales)} material(es)"
                                       + (f" · solicitado por {r.supervisor}" if r.supervisor else ""))
                    except Exception as e:
                        errores.append(f"{a.name}: {e}")
                st.session_state[f"rep_resumen_{fecha}"] = resumen
                for e in errores:
                    st.error(e)
                if not errores:
                    st.rerun()
    
        if st.session_state.get(f"rep_resumen_{fecha}"):
            with st.expander("Resultado de la última carga", expanded=True):
                for linea in st.session_state[f"rep_resumen_{fecha}"]:
                    st.markdown(f"- {linea}")
    
        _tarja(fecha, pl)
        _tarja_maquinaria(fecha, pl)
    
        reps = db.reportes_diario(fecha.isoformat(), c.contrato_id())
        if not reps:
            st.info("Aún no hay reportes cargados para esta fecha.")
            return
        plantel = {p.rut: p for p in pl.personas}
        filas = []
        for r in reps:
            d = r["datos"]
            trab = d.get("trabajadores", [])
            esp = [plantel[t["rut"]].especialidad for t in trab if t["rut"] in plantel]
            tipo = d.get("tipo", "terreno")
            especialidad = d.get("especialidad") if tipo != "terreno" and d.get("especialidad") else \
                (max(set(esp), key=esp.count) if esp else "–")
            filas.append({"id": r["id"], "archivo": r["archivo"], "tipo": D.nombre_tipo(d), "especialidad": especialidad,
                          "turno": d.get("turno", "Día"), "supervisor": d.get("supervisor", ""), "fecha_reporte": d.get("fecha"),
                          "personas": len(trab), "hh": sum(t.get("total", 0) for t in trab),
                          "fuera_plantel": sum(1 for t in trab if t["rut"] not in plantel),
                          "actividades": len(d.get("actividades", [])) + len(d.get("uso_recurso", [])),
                          "materiales": len(d.get("materiales", [])),
                          "avisos": " · ".join(d.get("avisos", []))})
        df = pd.DataFrame(filas)
        df.insert(0, "quitar", False)
        editado = st.data_editor(
            df, hide_index=True, use_container_width=True, key=f"ed_rep_{fecha}_{st.session_state.get('rep_ver', 0)}",
            disabled=[x for x in df.columns if x not in ("turno", "quitar")],
            column_config={"quitar": st.column_config.CheckboxColumn("Quitar", width="small",
                                                                     help="Marca los reportes que quieres eliminar"),
                           "id": None, "archivo": "Archivo", "tipo": "Formato", "especialidad": "Especialidad",
                           "turno": st.column_config.SelectboxColumn("Turno", options=["Día", "Noche"], required=True),
                           "supervisor": "Supervisor", "fecha_reporte": "Fecha del reporte", "personas": "Personas",
                           "hh": st.column_config.NumberColumn("HH", format="%.0f"), "fuera_plantel": "No están en BD Personal",
                           "actividades": "Actividades", "materiales": "Materiales", "avisos": st.column_config.TextColumn("Avisos", width="large")})
        k = st.columns(3)
        k[0].metric("Reportes", len(df))
        k[1].metric("Personas", int(df["personas"].sum()))
        k[2].metric("HH reportadas", c.num(df["hh"].sum()))
        a, b = st.columns(2)
        if a.button("Guardar cambios de turno", key="btn_turnos"):
            for fila, r in zip(editado.itertuples(), reps):
                if fila.turno != r["datos"].get("turno"):
                    r["datos"]["turno"] = fila.turno
                    db.actualizar_reporte_diario(r["id"], r["datos"])
            st.rerun()
        marcados = [int(i) for i, q in zip(editado["id"], editado["quitar"]) if q]
        if b.button(f"🗑️ Quitar seleccionados ({len(marcados)})", key="btn_del_sel", disabled=not marcados,
                    type="primary" if marcados else "secondary"):
            for rid in marcados:
                db.borrar_reporte_diario(rid)
            st.session_state["rep_ver"] = st.session_state.get("rep_ver", 0) + 1
            st.toast(f"Se quitaron {len(marcados)} reporte(s).")
            st.rerun()
        with b.expander(f"Quitar todos los reportes del {fecha:%d-%m-%Y}"):
            st.caption(f"Elimina los {len(df)} reportes cargados para este día (los archivos originales no se tocan; puedes volver "
                       "a cargarlos).")
            if st.checkbox("Sí, quitar todos", key=f"conf_del_todos_{fecha}") and st.button("Quitar todos", key="btn_del_todos"):
                for rid in df["id"].tolist():
                    db.borrar_reporte_diario(int(rid))
                st.session_state["rep_ver"] = st.session_state.get("rep_ver", 0) + 1
                st.rerun()
    
    
    def _clave_tarja(fecha: dt.date) -> str:
        return f"tarja_{fecha.isoformat()}"
    
    
    def tarja_del_dia(fecha: dt.date) -> dict | None:
        return db.config(c.contrato_id(), _clave_tarja(fecha))
    
    
    def _tarja(fecha: dt.date, pl):
        st.markdown("##### Tarja de asistencia (opcional, recomendada)")
        st.caption("Con la tarja, la app usa la asistencia real del día: quién marcó (T / TN), faltas, licencias, permisos y descansos. "
                   "Manda sobre los reportes de terreno y sobre la rotación de turnos. Sirve la lista del día (RUT + estado u hora "
                   "de entrada), una lista de varios días con columna FECHA, o la planilla mensual con una columna por día.")
        actual = tarja_del_dia(fecha)
        archivo = st.file_uploader("Tarja de asistencia (.xlsx, .xlsm o .csv)", type=["xlsx", "xlsm", "csv"], key=f"up_tarja_{fecha}")
        if archivo and st.button("Cargar tarja", type="primary", key=f"btn_tarja_{fecha}"):
            try:
                marcas, notas = D.leer_tarja(archivo.getvalue(), archivo.name, fecha)
            except Exception as e:
                st.error(f"{archivo.name}: {e}")
            else:
                from dataclasses import asdict
                db.guardar_config(c.contrato_id(), _clave_tarja(fecha), {
                    "archivo": archivo.name, "cargado_por": c.nombre_usuario(), "cargado": dt.datetime.now().strftime("%d-%m-%Y %H:%M"),
                    "notas": notas, "marcas": {r: asdict(m) for r, m in marcas.items()}})
                st.rerun()
        if not actual:
            return
        marcas = actual.get("marcas", {})
        plantel = {p.rut for p in pl.personas}
        cuenta: dict[str, int] = {}
        for m in marcas.values():
            if not m["estado"]:
                continue                                      # solo ficha (sin marca del día)
            e = m["estado"] if m["estado"] != "P" else ("TN" if m.get("noche") else "Asistió")
            cuenta[e] = cuenta.get(e, 0) + 1
        numeros = {r.split("-")[0] for r in plantel}
        fuera = sum(1 for r in marcas if r not in plantel and r.split("-")[0] not in numeros)
        st.success(f"Tarja cargada: **{actual['archivo']}** ({actual.get('cargado', '')}, {actual.get('cargado_por', '')}) · "
                   f"{len(marcas)} personas · " + " · ".join(f"{D.ESTADOS.get(k, k)}: {v}" for k, v in sorted(cuenta.items()))
                   + (f" · {fuera} RUT no están en BD Personal" if fuera else ""))
        for n in actual.get("notas", []):
            st.caption(n)
        vig = [m for m in marcas.values() if m.get("ficha") and D._vigente(D.Marca(**m))]
        if vig:
            dir_t = sum(1 for m in vig if D.norm(m["ficha"].get("W")) == "DIRECTO")
            ind_t = len(vig) - dir_t
            dir_b = sum(1 for p in pl.personas if D.norm(p.moimod) == "DIRECTO")
            nuevos = sum(1 for r, m in marcas.items() if m in vig and r not in plantel and r.split("-")[0] not in numeros)
            fin = sum(1 for r, m in marcas.items() if r in plantel and m.get("ficha") and D._finiquitado(D.Marca(**m)))
            st.info(f"Contratados vigentes según la tarja: **{dir_t} directos** y **{ind_t} indirectos**. BD Personal del informe "
                    f"anterior tiene {dir_b} directos. Al generar: se agregan **{nuevos}** contratado(s) nuevo(s) con sus datos de la "
                    f"tarja y salen **{fin}** finiquitado(s).")
        with st.expander("Ver tarja leída"):
            st.dataframe(pd.DataFrame([{"RUT": r, "Nombre": m["nombre"], "Estado": m["estado"],
                                        "Turno": {True: "Noche", False: "Día"}.get(m.get("noche"), "–"), "En la tarja": m["texto"],
                                        "En BD Personal": "Sí" if r in plantel or r.split("-")[0] in numeros else "No"}
                                       for r, m in marcas.items()]), hide_index=True, use_container_width=True)
        if st.button("Quitar tarja", key=f"btn_quitar_tarja_{fecha}"):
            db.guardar_config(c.contrato_id(), _clave_tarja(fecha), None)
            st.rerun()
    
    
    def tarja_maquinaria_del_dia(fecha: dt.date) -> dict | None:
        return db.config(c.contrato_id(), f"tarja_maq_{fecha.isoformat()}")
    
    
    def _tarja_maquinaria(fecha: dt.date, pl):
        st.markdown("##### Tarja de maquinaria (opcional)")
        st.caption("Estado de cada equipo por PATENTE para el día: operando (T / TN / día y noche), mantención, panne, stand by o "
                   "desmovilizado. Sirve la planilla mensual (una columna por día) o una lista del día con estado, turno y observación. "
                   "Llena «Maquinaria del día» (BD Maquinaria) y con ella la sección 4.0.")
        actual = tarja_maquinaria_del_dia(fecha)
        archivo = st.file_uploader("Tarja de maquinaria (.xlsx, .xlsm o .csv)", type=["xlsx", "xlsm", "csv"], key=f"up_tarja_maq_{fecha}")
        if archivo and st.button("Cargar tarja de maquinaria", type="primary", key=f"btn_tarja_maq_{fecha}"):
            try:
                datos_m, notas = D.leer_tarja_maquinaria(archivo.getvalue(), archivo.name, fecha)
            except Exception as e:
                st.error(f"{archivo.name}: {e}")
            else:
                db.guardar_config(c.contrato_id(), f"tarja_maq_{fecha.isoformat()}", {
                    "archivo": archivo.name, "cargado_por": c.nombre_usuario(), "cargado": dt.datetime.now().strftime("%d-%m-%Y %H:%M"),
                    "notas": notas, "equipos": datos_m})
                st.rerun()
        if not actual:
            return
        eq = actual.get("equipos", {})
        en_bd = {D._clave_patente(d["patente"]) for d in pl.maquinaria}
        fuera = [v.get("_patente") or k for k, v in eq.items() if k not in en_bd]
        cuenta: dict[str, int] = {}
        for v in eq.values():
            e_ = v.get("estado") or "Sin estado"
            if e_ == "OPERANDO" and v.get("observacion"):
                e_ += f" ({v['observacion']})"
            cuenta[e_] = cuenta.get(e_, 0) + 1
        st.success(f"Tarja de maquinaria: **{actual['archivo']}** ({actual.get('cargado', '')}) · {len(eq)} equipos · "
                   + " · ".join(f"{k}: {n}" for k, n in sorted(cuenta.items())))
        if fuera:
            st.warning(f"{len(fuera)} patente(s) de la tarja no están en BD Maquinaria del informe (no se pueden ubicar en la "
                       f"sección 4.0): {', '.join(fuera[:20])}" + (" …" if len(fuera) > 20 else ""))
        faltan = [d["patente"] for d in pl.maquinaria if D._clave_patente(d["patente"]) not in eq]
        if faltan:
            st.caption(f"{len(faltan)} equipo(s) de BD Maquinaria no vienen en la tarja (se mantiene su estado del día anterior): "
                       + ", ".join(faltan[:20]) + (" …" if len(faltan) > 20 else ""))
        if st.button("Quitar tarja de maquinaria", key=f"btn_quitar_tarja_maq_{fecha}"):
            db.guardar_config(c.contrato_id(), f"tarja_maq_{fecha.isoformat()}", None)
            st.rerun()
    
    
    # ------------------------------------------------------------------ revisar y generar
    
    def _generar(base, pl, fecha: dt.date):
        if fecha <= pl.fecha:
            st.warning(f"La fecha elegida ({c.fecha(fecha)}) no es posterior a la planilla base ({c.fecha(pl.fecha)}). "
                       "Cambia la fecha en la barra lateral.")
            return
        if fecha != pl.fecha + dt.timedelta(days=1):
            st.warning(f"La planilla base es del {c.fecha(pl.fecha)}: los acumulados se arrastrarán desde ese día. "
                       "Si hubo informes intermedios, carga primero el último emitido como planilla base.")
        reps = [D.ReporteTerreno.de_dict(r["datos"]) for r in db.reportes_diario(fecha.isoformat(), c.contrato_id())]
        if not reps:
            st.info(f"Carga los reportes de terreno del {c.fecha(fecha)} en la pestaña **Reportes de terreno**.")
            return
    
        clave = f"diario_{fecha}_{base['id']}"
        _restaurar_borrador(clave)
        _barra_borrador(clave)
        ajustes = st.session_state.setdefault(f"{clave}_ajustes", {})
        encabezados = db.config(c.contrato_id(), "encabezados", {})
        tj = tarja_del_dia(fecha)
        cons = D.consolidar(pl, reps, fecha, ajustes, encabezados, agregar_nuevos=st.session_state.get(f"{clave}_nuevos", True),
                            tarja=(tj or {}).get("marcas"))
        t = cons.totales
        if tj:
            st.caption(f"✅ Asistencia según la tarja **{tj['archivo']}** ({len(tj.get('marcas', {}))} personas).")
        else:
            st.caption("Sin tarja de asistencia para este día: el estado sale de los reportes de terreno y de la rotación de turnos. "
                       "Cárgala en **Reportes de terreno** para un conteo más exacto.")
    
        st.subheader(f"Informe N° {pl.numero + 1} · {c.fecha(fecha)}")
        k = st.columns(5)
        k[0].metric("Directos turno día", t["MOD"]["R"], delta=t["MOD"]["R"] - pl.totales["MOD"]["R"])
        k[1].metric("Directos turno noche", t["MOD"]["U"], delta=t["MOD"]["U"] - pl.totales["MOD"]["U"])
        k[2].metric("HH construcción", c.num(cons.hh_constr))
        k[3].metric("HH precom", c.num(cons.hh_precom))
        k[4].metric("Indirectos en faena", t["MOI"]["R"] + t["MOI"]["U"])
        st.caption(f"Reportes: {len(reps)} · Especialidades con reporte: {', '.join(cons.especialidades_con_reporte)}")
        if cons.avisos:
            with st.expander(f"⚠️ Revisar ({len(cons.avisos)} avisos)", expanded=True):
                for a in cons.avisos:
                    st.markdown(f"- {a}")
    
    
        pestañas = st.tabs(["1.0 Encabezado y resumen", "Descripción de actividades", "1.1 Control T&M", "2.0 MOD · 3.0 MOI",
                            "4.0 Equipos", "5.0 Interferencias", "6.0 UDR", "8.0 Materiales", "8.0 Validación · Fotos",
                            "✅ Revisar y generar"])
        t10, tdesc, t11, tper, t40, t50, t60, t80, tval, tgen = pestañas
    
        with tper:
            st.caption("Estado de cada persona en BD Personal (tarja de asistencia > reportes de terreno > rotación de turno). "
                       "Corrige aquí cualquier estado (T: día, TN: noche, D: descanso, L: licencia, PSG/PCG: permisos, F: falla).")
            # ---- 1. Personal
            st.caption("Se actualiza BD Personal según los reportes. Corrige aquí cualquier estado (T: día, TN: noche, D: descanso, "
                       "L: licencia, PSG/PCG: permisos, F: falla).")
            if cons.cambios:
                dfc = pd.DataFrame([x.__dict__ for x in cons.cambios])
                ed = st.data_editor(dfc, hide_index=True, use_container_width=True, height=280,
                                    key=f"{clave}_ed_cambios_{abs(hash(frozenset(ajustes.items())))}",
                                    disabled=["rut", "nombre", "cargo", "especialidad", "antes", "motivo"],
                                    column_config={"rut": "RUT", "nombre": "Nombre", "cargo": "Cargo", "especialidad": "Especialidad",
                                                   "antes": "Estado anterior",
                                                   "ahora": st.column_config.SelectboxColumn("Estado nuevo", options=OPCIONES_ESTADO, required=True),
                                                   "motivo": "Motivo"})
                nuevos_aj = {r.rut: r.ahora for r, o in zip(ed.itertuples(), cons.cambios) if r.ahora != o.ahora}
                if nuevos_aj:
                    ajustes.update(nuevos_aj)
                    st.rerun()
            with st.expander("Ajustar el estado de otra persona"):
                nombres = {p.rut: f"{p.nombre} · {p.cargo} ({p.faena})" for p in cons.personas}
                a, b, d = st.columns([3, 1, 1])
                rut = a.selectbox("Persona", list(nombres), format_func=nombres.get, key=f"{clave}_aj_rut")
                est = b.selectbox("Estado", OPCIONES_ESTADO, key=f"{clave}_aj_est")
                if d.button("Aplicar", key=f"{clave}_aj_btn"):
                    ajustes[rut] = est
                    st.rerun()
            if cons.nuevos:
                st.warning(f"{len(cons.nuevos)} persona(s) de los reportes no están en BD Personal.")
                st.dataframe(pd.DataFrame(cons.nuevos), hide_index=True, use_container_width=True)
                st.checkbox("Agregarlas a BD Personal con los datos del reporte", value=True, key=f"{clave}_nuevos")
    
    
            _tablas_personal(pl, cons)
    
        with tdesc:
            # ---- 2. Actividades
            st.markdown("##### Descripción de actividades realizadas y observaciones (bloque 1.0 del informe)")
            previos = D.encabezados_previos(pl.actividades_texto)
            opciones = list(dict.fromkeys(previos + pl.listas.get("areas", []) + ["Turno Noche"] +
                                          [a["encabezado"] for a in cons.actividades if a["encabezado"]]))
            dfa = pd.DataFrame(cons.actividades)
            master = pl.formato == "master"
            recordadas = db.config(c.contrato_id(), "actividad_asignacion", {}) if master else {}
            if not dfa.empty:
                columnas = ["incluir", "encabezado", "descripcion", "especialidad", "turno", "hh", "clave", "archivo", "codigo"]
                cfg = {"incluir": st.column_config.CheckboxColumn("Incluir", width="small"),
                       "encabezado": st.column_config.SelectboxColumn("Sistema / encabezado", options=opciones, width="medium"),
                       "descripcion": st.column_config.TextColumn("Actividad", width="large"),
                       "especialidad": "Especialidad", "turno": "Turno",
                       "hh": st.column_config.NumberColumn("HH", format="%.0f"), "clave": None, "archivo": None, "codigo": None}
                if master:
                    # Formato T&M: cada actividad se asocia a su Activity ID del P6 y a su criterio (tipo de actividad)
                    def _clave_desc(t):
                        return re.sub(r"\d+([.,]\d+)?", "#", D.norm(t))
                    recordadas_sn = {_clave_desc(k): v for k, v in recordadas.items()}
                    def sugerir(fila):
                        previo = recordadas.get(D.norm(fila["descripcion"])) or recordadas_sn.get(_clave_desc(fila["descripcion"])) or {}
                        aid, origen = fila.get("activity_id"), fila.get("origen_id")
                        if origen != "Reporte" and previo.get("id"):
                            aid, origen = previo["id"], "Asignado antes"
                        v = pl.p6.get(aid) or {}
                        area = v.get("K") or fila.get("area_p6")
                        return pd.Series({"activity_id": aid if aid in pl.p6 else None, "area_p6": area,
                                          "subsistema": v.get("I") or fila.get("subsistema"),
                                          "origen_id": origen if aid in pl.p6 else None,
                                          "criterio": previo.get("criterio") or _limpio(fila.get("criterio_sug"))
                                                      or (adicional if fila.get("udr") else D.CRITERIO_DEFECTO)})
                    adicional = next((x for x in pl.criterios if D.norm(x).startswith("EXTRAORDINARIA")), "Extraordinaria / Adicional")
                    sug = dfa.apply(lambda r: sugerir(r.to_dict()), axis=1)
                    dfa = dfa.drop(columns=[k for k in sug.columns if k in dfa]).join(sug)
                    columnas += ["activity_id", "area_p6", "subsistema", "origen_id", "criterio"]
                    cfg["activity_id"] = st.column_config.SelectboxColumn("Activity ID (P6)", options=sorted(pl.p6), width="medium")
                    areas = sorted({str(v.get("K")) for v in pl.p6.values() if v.get("K")})
                    cfg["area_p6"] = st.column_config.SelectboxColumn("Área (P6)", options=areas, width="small")
                    cfg["subsistema"] = st.column_config.TextColumn("Subsistema", width="small")
                    cfg["origen_id"] = st.column_config.TextColumn("Origen ID", width="small",
                                                                   help="Reporte: venía escrito en el descriptor · Asignado antes: lo elegiste "
                                                                        "otro día para esta actividad · Sugerido: coincidencia clara con el P6")
                    cfg["criterio"] = st.column_config.SelectboxColumn("Criterio", options=pl.criterios or [D.CRITERIO_DEFECTO],
                                                                       width="medium", required=True)
                ed_a = _editor(clave, "act", dfa[columnas], "clave", hide_index=True, use_container_width=True, height=320,
                                      disabled=["especialidad", "turno", "hh", "clave", "subsistema", "origen_id"], column_config=cfg)
                acts = ed_a.to_dict("records")
                if master:
                    for a_ in acts:                               # el área sigue al Activity ID elegido
                        v_ = pl.p6.get(_limpio(a_.get("activity_id")) or "")
                        if v_ and v_.get("K"):
                            a_["area_p6"] = v_["K"]
                    sin_id = sum(1 for a in acts if not _limpio(a.get("activity_id")))
                    st.caption(f"Formato T&M: las HH de cada persona van a «Master Actividades» repartidas entre sus actividades, con el "
                               f"Activity ID y el criterio de esta tabla (se recuerdan para los próximos días)."
                               + (f" ⚠️ {sin_id} actividad(es) sin Activity ID." if sin_id else ""))
                faltan = sum(1 for a in acts if a["incluir"] and not a["encabezado"])
                if faltan:
                    st.caption(f"⚠️ {faltan} actividades sin sistema asignado. El sistema elegido se recuerda para los próximos días.")
            else:
                acts = []
            nota = st.text_area("Nota (opcional)", key=f"{clave}_nota", height=70)
            texto = D.texto_actividades(acts, nota)
            previo = st.session_state.get(f"{clave}__texto") or {}
            inicial = previo.get("final") if previo.get("base") == texto and previo.get("final") is not None else texto
            texto_final = st.text_area("Texto que irá en el informe (puedes editarlo)", inicial, height=260, key=f"{clave}_texto_{hash(texto)}")
            st.session_state[f"{clave}__texto_actual"] = {"base": texto, "final": texto_final}
    
    
        with t10:
            numero = st.number_input("N° de informe", value=pl.numero + 1, step=1, key=f"{clave}_num")
            st.caption(f"Fecha {c.fecha(fecha)} · {CAL.texto(fecha) or ''}")
            # ---- 3. Datos del día
            a, b = st.columns(2)
            conds = pl.listas.get("condicion") or [pl.condicion]
            condicion = a.selectbox("Condición climática", conds, index=conds.index(pl.condicion) if pl.condicion in conds else 0,
                                    key=f"{clave}_cond")
            resg0 = sum(float(cons.clasificacion.get(k_, (0, 0))[1] or 0) for k_ in ("lluvia", "alerta1", "alerta2"))
            resguardo = b.number_input("HH de resguardo por condición climática del día", 0.0, value=resg0, step=1.0, key=f"{clave}_resg",
                                       help="Se propone la suma de HH en lluvia y alertas de viento leídas de los reportes.")
            cls = {}
            cols = st.columns(4)
            for col, (k_, rot) in zip(cols, [("cursos", "Personal en cursos"), ("lluvia", "Condición de lluvia"),
                                             ("alerta1", "Alerta nivel 1"), ("alerta2", "Alerta nivel 2")]):
                col.markdown(f"**{rot}**")
                c0, h0 = cons.clasificacion.get(k_, (0, 0.0))
                cant = col.number_input("Cantidad", 0, value=int(c0 or 0), step=1, key=f"{clave}_{k_}_c")
                hh = col.number_input("HH", 0.0, value=float(h0 or 0), step=1.0, key=f"{clave}_{k_}_h")
                cls[k_] = (cant, hh)
    
    
            _resumen_10(pl, cons)
    
        with t40:
            cambios_maq, derivado = {}, {}
            if pl.maquinaria:
                with st.expander(f"Maquinaria del día · BD Maquinaria ({len(pl.maquinaria)} equipos con patente)", expanded=False):
                    st.caption("Se arrastra el estado del día anterior. Lo que cambies aquí va a BD Maquinaria y a la sección 4.0 "
                               "(Turno día, Turno noche y observación), así ambas siempre calzan. DESMOVILIZADO saca al equipo de la base.")
                    tmq = (tarja_maquinaria_del_dia(fecha) or {}).get("equipos", {})
                    if tmq:
                        st.caption(f"✅ Estados tomados de la tarja de maquinaria ({len(tmq)} equipos).")
                    dfm = pd.DataFrame([{k: ({**d, **{a: b for a, b in tmq.get(D._clave_patente(d["patente"]), {}).items()
                                                      if not a.startswith("_")}}).get(k, "")
                                         for k in ("patente", "tipo", "modelo", "disponible", "estado", "observacion", "turno", "operador")}
                                        for d in pl.maquinaria])
                    ed_m = _editor(
                        clave, "maq", dfm, "patente", hide_index=True, use_container_width=True, height=320, disabled=["patente", "tipo", "modelo"],
                        column_config={"patente": "Patente", "tipo": "Equipo", "modelo": "Modelo / marca",
                                       "disponible": st.column_config.SelectboxColumn("Disponible", options=["SI", "NO"]),
                                       "estado": st.column_config.SelectboxColumn(
                                           "Estado", options=sorted({"OPERANDO", "MANTENCION", "PANNE", "STAND BY", "DESMOVILIZADO"}
                                                                    | {d["estado"] for d in pl.maquinaria if d["estado"]})),
                                       "observacion": st.column_config.TextColumn("Observación (Turno Noche / Turno Día y Noche…)"),
                                       "turno": st.column_config.TextColumn("Turno"), "operador": st.column_config.TextColumn("Operador (si/no)")})
                    for orig, r in zip(pl.maquinaria, ed_m.to_dict("records")):
                        dif = {k: (_limpio(r[k]) or "") for k in ("disponible", "estado", "observacion", "turno", "operador")
                               if (_limpio(r[k]) or "") != (orig.get(k) or "")}
                        hor = tmq.get(D._clave_patente(orig["patente"]), {}).get("horometro")
                        if hor:
                            dif["horometro"] = hor
                        if dif:
                            cambios_maq[D._clave_patente(orig["patente"])] = dif
                    hoy = D.maquinaria_del_dia(pl, cambios_maq)
                    derivado = {D._clave_patente(d["patente"]): d for d in hoy}
                    # control con la tarja / BD Personal: operadores en faena por tipo de equipo
                    ops = {}
                    for p_ in cons.personas:
                        cg = D.norm(p_.cargo)
                        if re.match(r"(OPERADOR|CONDUCTOR|CHOFER)\b", cg):
                            tipo = re.sub(r"^(OPERADOR|CONDUCTOR|CHOFER)\s+(DE\s+)?", "", cg).strip()
                            t_ = p_.faena.upper()
                            ops.setdefault(tipo, [0, 0])
                            ops[tipo][0] += t_ in ("T", "TC")
                            ops[tipo][1] += t_ == "TN"
                    filas_ctrl = []
                    for tipo, (od, on) in sorted(ops.items()):
                        eqs = [d for d in hoy if D.norm(d["tipo"]).startswith(tipo) or tipo.startswith(re.sub(r"\s*\d+$", "", D.norm(d["tipo"])))]
                        if not eqs:
                            continue
                        ed_, en_ = sum(D.operando_dia(d) for d in eqs), sum(D.operando_noche(d) for d in eqs)
                        if ed_ > od or en_ > on:
                            filas_ctrl.append({"Tipo de equipo": tipo.title(), "Equipos día": ed_, "Operadores día (tarja)": od,
                                               "Equipos noche": en_, "Operadores noche (tarja)": on})
                    if filas_ctrl:
                        st.warning("Hay equipos operando sin operadores suficientes en faena según la tarja / BD Personal. Revisa su "
                                   "estado o la asistencia de los operadores:")
                        st.dataframe(pd.DataFrame(filas_ctrl), hide_index=True, use_container_width=True)
                    else:
                        st.caption("✅ Cada tipo de equipo operando tiene operadores en faena para su turno (según tarja / BD Personal).")
    
            with st.expander("Equipos y maquinaria (se arrastra el día anterior)"):
                def _dia(e):
                    d = derivado.get(D._clave_patente(e.patente)) if e.patente else None
                    if d is not None:
                        return 1 if D.operando_dia(d) else 0
                    if e.patente and pl.maquinaria and D._clave_patente(e.patente) in {D._clave_patente(x["patente"]) for x in pl.maquinaria}:
                        return 0                                   # desmovilizado hoy
                    return e.dia
                dfe = pd.DataFrame([{"fila": e.fila, "grupo": e.grupo, "equipo": e.nombre, "patente": e.patente, "dia": _dia(e),
                                     "noche": e.noche, "observacion": e.observacion} for e in pl.equipos])
                if pl.maquinaria:
                    st.caption("Los equipos con patente toman el Turno día de «Maquinaria del día» (BD Maquinaria); cámbialos allá.")
                ed_e = _editor(clave, "eq", dfe, "fila", hide_index=True, use_container_width=True, height=320,
                                      disabled=["grupo", "equipo", "patente"],
                                      column_config={"fila": None, "grupo": "Grupo", "equipo": "Equipo", "patente": "Patente",
                                                     "dia": st.column_config.NumberColumn("Turno día", min_value=0, step=1),
                                                     "noche": st.column_config.NumberColumn("Turno noche", min_value=0, step=1),
                                                     "observacion": "Observación"})
                equipos = {int(r.fila): {"dia": _limpio(r.dia) or 0, "noche": _limpio(r.noche) or 0,
                                         "observacion": _limpio(r.observacion) or ""} for r in ed_e.itertuples()}
    
    
        with t50:
            with st.expander(f"Interferencias (hasta {len(pl.filas_interferencia)})"):
                dfi = pd.DataFrame([{k: x.get(k) for k in ("detalle", "area", "equipo", "inicio", "termino", "horas", "dotacion", "hh", "tipo")}
                                    for x in cons.interferencias],
                                   columns=["detalle", "area", "equipo", "inicio", "termino", "horas", "dotacion", "hh", "tipo"])
                if cons.interferencias:
                    st.caption(f"{len(cons.interferencias)} interferencia(s) / detención(es) tomadas de los reportes de terreno "
                               "(actividades con criterio Interferencia o Detención). Revisa horas de inicio y término.")
                rot = pl.cab_interferencia or {}
                r_ = lambda col, defecto: (rot.get(col) or defecto).replace("\n", " ")
                ed_i = _editor(clave, "int", dfi, None, num_rows="dynamic", use_container_width=True,
                                      column_config={"detalle": r_("B", "Detalle"), "area": r_("L", "Área"), "equipo": r_("O", "Equipo"),
                                                     "inicio": r_("R", "Inicio"), "termino": r_("U", "Término"), "horas": r_("X", "Horas"),
                                                     "tipo": st.column_config.SelectboxColumn(r_("AG", "Tipo"), options=pl.listas.get("interferencia") or None),
                                                     "dotacion": st.column_config.NumberColumn(r_("AB", "Dotación"), step=1),
                                                     "hh": st.column_config.NumberColumn(r_("AE", "HH"))})
                interferencias = _filas_limpias(ed_i.to_dict("records"))
                if len(interferencias) > len(pl.filas_interferencia):
                    st.warning(f"La planilla tiene espacio para {len(pl.filas_interferencia)} interferencias; se escribirán las primeras.")
    
    
        with t60:
            with st.expander(f"Uso de recurso / NDC (la planilla trae {len(pl.filas_udr)} filas y se amplía si hay más)", expanded=True):
                sug = list(cons.uso_recurso)
                # actividades que tú marcaste con criterio «Extraordinaria / Adicional» también van a UDR
                por_clave = {(a["archivo"], a["codigo"]): a for a in cons.actividades}
                for a_ in acts if master else []:
                    orig = por_clave.get((a_.get("archivo"), a_.get("codigo")), {})
                    if orig.get("udr") or not D.norm(_limpio(a_.get("criterio")) or "").startswith("EXTRAORDINARIA"):
                        continue
                    sug.append({"descripcion": a_["descripcion"], "directos": orig.get("dotacion") or None,
                                "hh_directos": orig.get("hh_reportadas") or orig.get("hh_jornada") or None, "especialidad": orig.get("especialidad"),
                                "area": D.area_udr(a_["descripcion"], orig.get("area") or ""), "respaldo": D.respaldo_udr(a_["descripcion"]),
                                "criterio": a_["criterio"], "desde_actividad": True})
                n0 = (pl.ultimo_ndc or 0) + 1
                dfu = pd.DataFrame([{"ndc": (n0 + i) if pl.ultimo_ndc else None, "descripcion": u.get("descripcion"),
                                     "detalle": u.get("detalle") or u.get("descripcion"),
                                     "directos": u.get("directos"), "hh_directos": u.get("hh_directos"), "indirectos": u.get("indirectos"),
                                     "hh_indirectos": u.get("hh_indirectos"), "equipos": u.get("equipos") or None, "hm": u.get("hm"),
                                     "especialidad": u.get("especialidad"), "respaldo": u.get("respaldo")} for i, u in enumerate(sug)],
                                   columns=["ndc", "descripcion", "detalle", "directos", "hh_directos", "indirectos", "hh_indirectos",
                                            "equipos", "hm", "especialidad", "respaldo"])
                extra_udr = [k for k in ("area", "criterio", "impacto") if k in pl.udr_cols]
                for k in extra_udr:
                    dfu[k] = [u.get(k) for u in sug] if sug else pd.Series(dtype=object)
                for k in ("descripcion", "detalle", "equipos", "especialidad", "respaldo", "area"):
                    if k in dfu:                                  # celdas de texto vacías en blanco (no «None»)
                        dfu[k] = dfu[k].astype(object).where(dfu[k].notna(), "")
                for k in ("ndc", "directos", "hh_directos", "indirectos", "hh_indirectos", "hm"):
                    dfu[k] = pd.to_numeric(dfu[k], errors="coerce")
                if sug:
                    n_act = sum(1 for u in sug if u.get("desde_actividad"))
                    st.caption(f"{len(sug)} fila(s): {len(sug) - n_act} de los archivos de uso de recurso / adicionales y {n_act} "
                               "actividad(es) de los reportes de terreno marcadas como adicional / UDR / NDC (no se repiten en la "
                               "descripción de actividades)"
                               + (f". La numeración sigue el correlativo del informe anterior (último N° {pl.ultimo_ndc})." if pl.ultimo_ndc
                                  else ". No se encontró el último N° NDC del informe anterior: completa la numeración.")
                               + " Revisa, completa o agrega las demás.")
                ed_u = _editor(clave, "udr", dfu, None, num_rows="dynamic", use_container_width=True,
                                      column_config={"ndc": st.column_config.NumberColumn("N° NDC", step=1),
                                                     "descripcion": st.column_config.TextColumn("Título / obra", width="medium"),
                                                     "detalle": st.column_config.TextColumn("Descripción del trabajo", width="large"),
                                                     "directos": "Directos", "hh_directos": "HH directos", "indirectos": "Indirectos",
                                                     "hh_indirectos": "HH indirectos", "equipos": "Equipos", "hm": "HM",
                                                     "especialidad": "Disciplina", "respaldo": "Respaldo (RFI / DCN / solicitante)",
                                                     "area": "Área / frente",
                                                     "criterio": st.column_config.SelectboxColumn("Criterio", options=pl.criterios or None),
                                                     "impacto": st.column_config.SelectboxColumn("Activity ID / impacto",
                                                                                                 options=sorted(pl.p6) or None)})
                udr = _filas_limpias(ed_u.to_dict("records"))
    
    
        with t80:
            materiales = _materiales(pl, cons, clave)
    
        with tval:
            firmas = _firmas(clave, bool(pl.validacion))
            validacion = {}
            if pl.validacion.get("fila_estado"):
                a_, b_ = st.columns([1, 2])
                validacion["estado"] = a_.selectbox("Estado de validación del informe", ["", "Pendiente de revisión", "Revisado",
                                                                                         "Aprobado", "Aprobado con observaciones",
                                                                                         "Rechazado"], key=f"{clave}_val_est")
                validacion["observacion"] = b_.text_input("Observaciones de la validación", key=f"{clave}_val_obs")
            with st.expander(f"Registro fotográfico ({len(pl.fotos)} fotos)"):
                fotos = []
                cols = st.columns(3)
                for i in range(len(pl.fotos)):
                    col = cols[i % 3]
                    img = col.file_uploader(f"Foto {i + 1}", type=["jpg", "jpeg", "png"], key=f"{clave}_foto_{i}")
                    ley = col.text_input(f"Leyenda foto {i + 1}", key=f"{clave}_ley_{i}")
                    fotos.append((img.getvalue() if img else None, ley))
                en_blanco = st.checkbox("Dejar en blanco las fotos que no reemplace (no repetir las del informe anterior)", value=True,
                                        key=f"{clave}_blanco")
    
        asignacion = {f"{a['archivo']}|{a['codigo']}": {"id": _limpio(a.get("activity_id")) or "",
                                                        "criterio": _limpio(a.get("criterio")) or "",
                                                        "area": _limpio(a.get("area_p6")) or ""} for a in acts if master}
        ent = D.Entradas(fecha, int(numero), condicion, texto_final, cls, resguardo, equipos, interferencias, udr, fotos, en_blanco,
                         firmas, asignacion, cambios_maq, materiales, validacion)
    
        with t11:
            if not master:
                st.info("La planilla base no tiene la sección 1.1 (formato antiguo).")
            if master:
                filas_m = D.filas_master_actividades(pl, cons, ent)
                if filas_m:
                    dm = pd.DataFrame(filas_m)
                    res = dm.groupby("criterio").agg(actividades=("id", lambda x: len({i for i in x if i})), dotacion=("rut", "count"),
                                                     hh=("hh", "sum")).reset_index()
                    if True:
                        tot_act, tot_tab = dm["hh"].sum(), cons.hh_constr + cons.hh_precom
                        k1, k2, k3 = st.columns(3)
                        k1.metric("HH MOD por actividad", c.num(tot_act))
                        k2.metric("HH MOD tabla de personal (tarja)", c.num(tot_tab))
                        k3.metric("Diferencia", c.num(tot_act - tot_tab),
                                  help="Debe ser 0: «OK - CONCILIADO» en el informe.")
                        if abs(tot_act - tot_tab) < 0.5:
                            st.success("OK - CONCILIADO: las HH por actividad calzan con las HH de la tarja.")
                        else:
                            st.error("REVISAR: las HH por actividad no calzan con la tabla de personal.")
                        res["%"] = (res["hh"] / res["hh"].sum() * 100).round(1)
                        st.dataframe(res.rename(columns={"criterio": "Tipo de actividad", "actividades": "N° actividades",
                                                         "dotacion": "Dotación MOD", "hh": "HH MOD", "%": "% HH MOD"}), hide_index=True,
                                     use_container_width=True)
    
        with tgen:
            _revision(pl, cons, ent, acts, tj, tarja_maquinaria_del_dia(fecha))
            if st.button("Generar informe", type="primary", key=f"{clave}_gen"):
                _guardar_borrador(clave)
                with st.spinner("Generando el informe sobre la planilla oficial…"):
                    st.session_state[f"{clave}_xlsx"] = D.generar_informe(pl, cons, ent)
                # recordar los sistemas elegidos para las actividades
                for a_ in acts:
                    if a_["encabezado"] and a_["turno"] != "Noche" and a_["encabezado"] != "Turno Noche":
                        encabezados[a_["clave"]] = a_["encabezado"]
                db.guardar_config(c.contrato_id(), "encabezados", encabezados)
                if master:
                    for a_ in acts:
                        aid, cr = _limpio(a_.get("activity_id")), _limpio(a_.get("criterio"))
                        if aid or (cr and cr != D.CRITERIO_DEFECTO):
                            recordadas[D.norm(a_["descripcion"])] = {"id": aid or "", "criterio": cr}
                    db.guardar_config(c.contrato_id(), "actividad_asignacion", recordadas)
            datos = st.session_state.get(f"{clave}_xlsx")
            if datos:
                nombre = f"K-CC-1002-Informe diario {fecha:%d-%m-%Y}-Rev.0.xlsx"
                st.success("Informe generado. Al abrirlo, Excel recalcula todas las fórmulas de la planilla.")
                a, b, d = st.columns(3)
                a.download_button("Descargar Excel", datos, file_name=nombre, mime=MIME_XLSX, key=f"{clave}_dl", type="primary")
                if b.button("Generar PDF", key=f"{clave}_pdf_btn"):
                    with st.spinner("Convirtiendo a PDF…"):
                        pdf = D.exportar_pdf(datos)
                    if pdf:
                        st.session_state[f"{clave}_pdf"] = pdf
                    else:
                        st.info("Este equipo no tiene LibreOffice para convertir a PDF. Abre el Excel y usa Archivo → Exportar → PDF.")
                if st.session_state.get(f"{clave}_pdf"):
                    b.download_button("Descargar PDF", st.session_state[f"{clave}_pdf"], file_name=nombre.replace(".xlsx", ".pdf"),
                                      mime="application/pdf", key=f"{clave}_dl_pdf")
                if d.button("Emitir: dejarlo como base del día siguiente", key=f"{clave}_emitir"):
                    aid = db.guardar_base_diario(nombre, fecha.isoformat(), int(numero), datos, c.nombre_usuario(), "generado",
                                                 contrato_id=c.contrato_id())
                    db.registrar_informe_diario(fecha.isoformat(), int(numero), {
                        "directos_dia": t["MOD"]["R"], "directos_noche": t["MOD"]["U"], "indirectos_dia": t["MOI"]["R"],
                        "indirectos_noche": t["MOI"]["U"], "hh_mod": cons.hh_constr + cons.hh_precom, "hh_moi": t["MOI"]["X"],
                        "reportes": len(reps), "firmas": [f"{f['rol']}: {f['nombre']}" + (" (firmado)" if f.get("firma") else "")
                                                          for f in ent.firmas]}, aid, c.nombre_usuario(), contrato_id=c.contrato_id())
                    _borrar_borrador(clave)
                    st.cache_data.clear()
                    st.success(f"Informe N° {int(numero)} emitido. Ya es la base para el {c.fecha(fecha + dt.timedelta(days=1))}.")
                    st.rerun()
    
    
    
    
    def _tablas_personal(pl, cons):
        """2.0 Mano de obra directa / 3.0 Mano de obra indirecta, por cargo, como en la planilla."""
        pc = cons.totales.get("por_cargo", {})
        for sec, titulo in (("MOD", "2.0 Mano de obra directa"), ("MOI", "3.0 Mano de obra indirecta")):
            filas = []
            for fc in [x for x in pl.cargos if x.seccion == sec]:
                v = pc.get(fc.fila, {})
                if not any(v.get(k) for k in "LORU"):
                    continue
                filas.append({"Cargo": fc.cargo, "Contratados": v.get("L", 0), "Acreditados": v.get("O", 0),
                              "Turno día": v.get("R", 0), "Turno noche": v.get("U", 0),
                              "En faena": v.get("R", 0) + v.get("U", 0), "HH faena": v.get("X", 0)})
            t = cons.totales[sec]
            ant = pl.totales[sec]
            st.markdown(f"**{titulo}** · en faena {t['R'] + t['U']} (día {t['R']} · noche {t['U']}) · HH {c.num(t['X'])} · "
                        f"informe anterior: {ant['R'] + ant['U']} en faena")
            if filas:
                st.dataframe(pd.DataFrame(filas), hide_index=True, use_container_width=True, height=min(420, 38 + 35 * len(filas)),
                             column_config={"HH faena": st.column_config.NumberColumn(format="%.0f")})
    
    
    def _resumen_10(pl, cons):
        """Resumen HH personal en obra (bloque derecho de 1.0)."""
        t, ta = cons.totales, pl.totales
        cierre = D.valores_cierre(pl)
        st.markdown("##### Resumen HH personal en obra")
        cant = pd.DataFrame([
            {"Ítem": "MOD contratado acreditado", "Anterior": ta["MOD"]["O"], "Actual": t["MOD"]["O"]},
            {"Ítem": "MOI contratado acreditado", "Anterior": ta["MOI"]["O"], "Actual": t["MOI"]["O"]},
            {"Ítem": "MOD contratado no acreditado", "Anterior": ta["MOD"]["L"] - ta["MOD"]["O"], "Actual": t["MOD"]["L"] - t["MOD"]["O"]},
            {"Ítem": "MOI contratado no acreditado", "Anterior": ta["MOI"]["L"] - ta["MOI"]["O"], "Actual": t["MOI"]["L"] - t["MOI"]["O"]},
            {"Ítem": "Total contratados proyecto", "Anterior": ta["MOD"]["L"] + ta["MOI"]["L"], "Actual": t["MOD"]["L"] + t["MOI"]["L"]}])
        hh = pd.DataFrame([
            {"Ítem": "MOD trabajando construcción", "Acumulado anterior": cierre["hh_constr"], "Actual": cons.hh_constr,
             "Acumulado": cierre["hh_constr"] + cons.hh_constr},
            {"Ítem": "MOD trabajando precom", "Acumulado anterior": cierre["hh_precom"], "Actual": cons.hh_precom,
             "Acumulado": cierre["hh_precom"] + cons.hh_precom},
            {"Ítem": "Total MOD trabajando", "Acumulado anterior": cierre["hh_constr"] + cierre["hh_precom"],
             "Actual": cons.hh_constr + cons.hh_precom,
             "Acumulado": cierre["hh_constr"] + cierre["hh_precom"] + cons.hh_constr + cons.hh_precom}])
        a, b = st.columns(2)
        a.dataframe(cant, hide_index=True, use_container_width=True)
        b.dataframe(hh, hide_index=True, use_container_width=True,
                    column_config={k: st.column_config.NumberColumn(format="%.0f") for k in ("Acumulado anterior", "Actual", "Acumulado")})
        st.caption("Se calcula desde BD Personal del día (tarja y reportes). Para cambiar un número, corrige el estado de la persona "
                   "en la pestaña «2.0 MOD · 3.0 MOI».")
    
    
    def _materiales(pl, cons, clave) -> list[dict]:
        if not pl.materiales:
            st.info("La planilla base no tiene la sección de materiales y consumibles.")
            return []
        cap = len(pl.materiales["filas"])
        st.caption(f"La planilla trae {cap} filas y se amplía si hay más. Se proponen los materiales de los PM de materiales "
                   "(PDF) y de los formularios UDR del día.")
        cols = ["descripcion", "unidad", "cantidad", "trabajo", "solicitado", "observacion"]
        dfm = pd.DataFrame([{k: m.get(k) for k in cols} for m in cons.materiales], columns=cols)
        for k in ("descripcion", "unidad", "trabajo", "solicitado", "observacion"):
            dfm[k] = dfm[k].astype(object).where(dfm[k].notna(), "")
        dfm["cantidad"] = pd.to_numeric(dfm["cantidad"], errors="coerce")
        ed = _editor(clave, "mat", dfm, None, num_rows="dynamic", use_container_width=True,
                            column_config={"descripcion": st.column_config.TextColumn("Descripción material / consumible", width="large"),
                                           "unidad": "Unidad", "cantidad": st.column_config.NumberColumn("Cantidad"),
                                           "trabajo": st.column_config.TextColumn("Activity ID o trabajo"),
                                           "solicitado": "Solicitado por / Aprobado por", "observacion": "Observación"})
        out = _filas_limpias(ed.to_dict("records"))
        if len(out) > cap:
            st.info(f"{len(out)} materiales: se agregan {len(out) - cap} fila(s) a la sección 8.0 al generar.")
        return out
    
    
    def _revision(pl, cons, ent, acts, tj, tmq):
        """Lista de control antes de generar: cada sección de la planilla y su fuente."""
        filas = []
        def fila(seccion, ok, detalle, fuente):
            filas.append({"": "✅" if ok is True else ("⚠️" if ok is False else "ℹ️"), "Sección": seccion, "Detalle": detalle,
                          "Fuente": fuente})
        fila("Tarja de trabajadores", bool(tj), f"{len((tj or {}).get('marcas', {}))} personas" if tj else "No cargada: el estado sale "
             "de reportes y rotación de turnos", "Reportes de terreno → Tarja de asistencia")
        fila("Tarja de maquinaria", bool(tmq) if pl.maquinaria else None,
             f"{len((tmq or {}).get('equipos', {}))} equipos" if tmq else "No cargada: se arrastra BD Maquinaria del día anterior",
             "Reportes de terreno → Tarja de maquinaria")
        sin = cons.especialidades_sin_reporte
        fila("Reportes de terreno", not sin, f"Especialidades con reporte: {len(cons.especialidades_con_reporte)}"
             + (f" · sin reporte: {', '.join(f'{k} ({v})' for k, v in sin.items())}" if sin else ""), "Reportes de terreno")
        sin_enc = sum(1 for a in acts if a.get("incluir") and not a.get("encabezado"))
        fila("1.0 Descripción de actividades", sin_enc == 0, f"{sum(1 for a in acts if a.get('incluir'))} actividades"
             + (f" · {sin_enc} sin encabezado" if sin_enc else ""), "Reportes de terreno + informe anterior")
        if pl.formato == "master":
            fm = D.filas_master_actividades(pl, cons, ent)
            dif = sum(m["hh"] for m in fm) - (cons.hh_constr + cons.hh_precom)
            sin_id = sum(1 for m in fm if not m["id"])
            fila("1.1 Control T&M", abs(dif) < 0.5, f"Diferencia HH actividad vs tarja: {c.num(dif)}"
                 + (f" · {sin_id} fila(s) sin Activity ID" if sin_id else ""), "Master Actividades (HH de la tarja repartidas)")
        fila("2.0 / 3.0 Personal", not cons.nuevos, f"MOD en faena {cons.totales['MOD']['R'] + cons.totales['MOD']['U']} · "
             f"MOI {cons.totales['MOI']['R'] + cons.totales['MOI']['U']}" + (f" · {len(cons.nuevos)} persona(s) fuera de BD Personal"
                                                                              if cons.nuevos else ""), "Tarja + reportes + BD Personal")
        fila("4.0 Equipos", None, f"{len(pl.equipos)} filas · {len(pl.maquinaria)} equipos con patente", "BD Maquinaria / tarja de maquinaria")
        fila("5.0 Interferencias", len(ent.interferencias) <= len(pl.filas_interferencia),
             f"{len(ent.interferencias)} de {len(pl.filas_interferencia)} filas", "Reportes (criterio Interferencia / Detención)")
        extra_u = max(0, len(ent.uso_recurso) - len(pl.filas_udr))
        fila("6.0 UDR", bool(pl.filas_udr) or not ent.uso_recurso,
             f"{len(ent.uso_recurso)} UDR · N° {(pl.ultimo_ndc or 0) + 1} a {(pl.ultimo_ndc or 0) + len(ent.uso_recurso)}"
             + (f" · se agregan {extra_u} fila(s) a la sección (la planilla trae {len(pl.filas_udr)})" if extra_u else ""),
             "Formularios UDR, adicionales y actividades marcadas")
        if pl.materiales:
            extra_m = max(0, len(ent.materiales) - len(pl.materiales["filas"]))
            fila("8.0 Materiales", True, f"{len(ent.materiales)} material(es)"
                 + (f" · se agregan {extra_m} fila(s) a la sección (la planilla trae {len(pl.materiales['filas'])})" if extra_m else ""),
                 "PM de materiales (PDF) + formularios UDR")
        elab = next((f for f in ent.firmas if f.get("rol") == "Elaborado por"), None)
        fila("8.0 Validación", bool(elab and elab.get("firma")), "Firmado por " + elab["nombre"] if elab and elab.get("firma")
             else "Sin firma de quien elabora", "Mi cuenta → Firma para informes")
        st.dataframe(pd.DataFrame(filas), hide_index=True, use_container_width=True,
                     column_config={"": st.column_config.TextColumn(width="small"), "Detalle": st.column_config.TextColumn(width="large")})
    
    
    def _historial():
        df = db.informes_diario(c.contrato_id())
        if df.empty:
            st.info("Aún no se han emitido informes desde el sistema.")
            return
        st.dataframe(df.drop(columns=["id", "archivo_id"]), hide_index=True, use_container_width=True,
                     column_config={"fecha": "Fecha", "numero": "N°", "emitido_por": "Emitido por", "emitido_en": "Emitido el",
                                    "directos_dia": "Directos día", "directos_noche": "Directos noche",
                                    "indirectos_dia": "Indirectos día", "indirectos_noche": "Indirectos noche",
                                    "hh_mod": st.column_config.NumberColumn("HH MOD", format="%.0f"),
                                    "hh_moi": st.column_config.NumberColumn("HH MOI", format="%.0f"), "reportes": "Reportes"})
        arch = db.archivos_diario(c.contrato_id())
        if not arch.empty:
            st.caption("Se conservan los últimos archivos emitidos:")
            i = st.selectbox("Archivo", arch["id"].tolist(),
                             format_func=lambda x: f"{arch.set_index('id').at[x, 'nombre']} ({arch.set_index('id').at[x, 'creado_en'][:16]})",
                             key="hist_arch")
            st.download_button("Descargar", db.archivo_diario(int(i)), file_name=arch.set_index("id").at[i, "nombre"],
                               mime=MIME_XLSX, key="hist_dl")
    
    
    def _firmas(clave: str, validacion: bool = False) -> list[dict]:
        """Bloque de firmas: quien genera firma con su propia firma (guardada en Mi cuenta); los demás roles quedan con
        nombre y cargo para firmar aparte. Nadie puede poner la firma de otra persona."""
        import datetime as _dt
        st.subheader("Firmas")
        u = c.usuario()
        mia = auth.firma(u["id"])
        filas = []
        a, b = st.columns([1, 2])
        if mia["png"]:
            a.image(mia["png"], width=180)
            firmar = b.checkbox(f"Firmar como «Elaborado por»: **{mia['nombre']}**" + (f" — {mia['cargo']}" if mia["cargo"] else ""),
                                value=True, key=f"{clave}_firmar")
            if firmar:
                filas.append({"rol": "Elaborado por", "nombre": mia["nombre"], "cargo": mia["cargo"], "firma": mia["png"],
                              "fecha": _dt.datetime.now()})
        else:
            b.info("Aún no cargas tu firma. Hazlo en **Cuenta → Mi cuenta → Firma para informes** y aparecerá aquí.")
            if b.checkbox("Incluir mi nombre como «Elaborado por» (sin firma)", key=f"{clave}_firmar_sin"):
                filas.append({"rol": "Elaborado por", "nombre": mia["nombre"], "cargo": mia["cargo"], "firma": None,
                              "fecha": _dt.datetime.now()})
        with st.expander("Agregar espacios para «Revisado por» y «Aprobado por» (firman aparte)"):
            for rol in ("Revisado por", "Aprobado por"):
                x, y = st.columns(2)
                nom = x.text_input(f"{rol}: nombre", key=f"{clave}_{rol}_nom")
                car = y.text_input(f"{rol}: cargo", key=f"{clave}_{rol}_car")
                if nom.strip():
                    filas.append({"rol": rol, "nombre": nom.strip(), "cargo": car.strip(), "firma": None, "fecha": ""})
        if filas:
            st.caption("Van a la tabla «8.0 Validación del informe» (Elaborado = «Preparado por (Contratista)»)." if validacion else
                       "El bloque «Firmas del informe» va al final del informe, dentro del área de impresión y del PDF.")
        return filas

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


vdiario = _modulo_vdiario()

# ======================================================================
# MÓDULO: vtm
# ======================================================================
def _modulo_vtm():
    import datetime as dt
    import os
    
    import pandas as pd
    import plotly.express as px
    import plotly.graph_objects as go
    import streamlit as st
    
    T = tm
    CAL = calendario
    c = comun
    
    COLORES_GRUPO = {"Directos OOCC": c.NARANJO, "Directos Montaje": c.PIZARRA, "Indirectos": c.AZUL}
    COLORES_TIPO = {"Directo": c.NARANJO, "Indirecto": c.PIZARRA}
    
    
    @st.cache_resource(show_spinner=False, max_entries=4)
    def _carga(cid: int):
        return T.ConsolidadoTM.de_dict(db.carga_tm(cid))
    
    
    def vista():
        st.title("Times & Material")
        cargas = db.cargas_tm(c.contrato_id())
        with st.expander("Cargar consolidado Times & Material", expanded=cargas.empty):
            _subir()
        if cargas.empty:
            st.info("Carga el consolidado mensual (.xlsx con las hojas DB MOD-MOI, DB MOD-MOI PRECOM y DB HORAS MAQUINAS).")
            return
    
        etiquetas = {int(r.id): f"{r.archivo} · datos al {c.fecha(r.fecha_corte)} · cargado {str(r.cargado_en)[:16]}"
                     for r in cargas.itertuples()}
        a, b = st.columns([3, 1])
        cid = a.selectbox("Archivo cargado", list(etiquetas), format_func=etiquetas.get, key="tm_carga")
        tm = _carga(int(cid))
        if c.usuario() and c.usuario()["rol"] == "admin":
            with b.popover("Eliminar carga"):
                if st.button("Eliminar esta carga", key="tm_borrar"):
                    db.borrar_carga_tm(int(cid))
                    st.cache_resource.clear()
                    st.rerun()
    
        # ---- filtros
        corte = tm.fecha_corte or max(tm.mo["fecha"])
        inicio = tm.fecha_inicio or min(tm.mo["fecha"])
        opciones = {"Mes de la fecha de corte": (corte.replace(day=1), corte),
                    "Últimos 7 días": (corte - dt.timedelta(days=6), corte),
                    "Últimos 30 días": (corte - dt.timedelta(days=29), corte),
                    "Todo el registro": (inicio, corte), "Personalizado": None}
        f1, f2, f3 = st.columns([1.4, 1.6, 1.4])
        per = f1.selectbox("Período", list(opciones), key="tm_periodo")
        if opciones[per] is None:
            rango = f2.date_input("Desde / hasta", (corte.replace(day=1), corte), min_value=inicio, max_value=corte,
                                  format="DD-MM-YYYY", key="tm_rango")
            desde, hasta = (rango if isinstance(rango, (list, tuple)) and len(rango) == 2 else (corte.replace(day=1), corte))
        else:
            desde, hasta = opciones[per]
            desde = max(desde, inicio)
            f2.caption(f"Del **{c.fecha(desde)}** al **{c.fecha(hasta)}** · datos hasta el {c.fecha(corte)}")
        ambitos = f3.multiselect("Ámbito", ["Construcción", "Precom"], default=["Construcción", "Precom"], key="tm_ambito") or \
            ["Construcción", "Precom"]
    
        mo, dot, eq = T.filtrar(tm, desde, hasta, ambitos)
        ind = T.indicadores(tm, desde, hasta, ambitos)
        st.caption(f"Contrato **{tm.contrato}** — {tm.nombre} · Período en semanas del proyecto: **{CAL.etiqueta(desde)}**" +
                   ("" if CAL.semana(desde) == CAL.semana(hasta) else f" a **{CAL.etiqueta(hasta)}**") +
                   f" · datos hasta la {CAL.texto(corte).lower()}")
        for av in tm.avisos:
            st.caption("⚠️ " + av)
    
        k = st.columns(5)
        var = None
        if ind["hh_dia_previo"]:
            var = (ind["hh_dia"] / ind["hh_dia_previo"] - 1) * 100
        k[0].metric("HH del período", c.num(ind["hh"]), delta=None if var is None else f"{c.num(var, 1)}% HH/día vs período anterior")
        k[1].metric("HH acumuladas contrato", c.num(ind["hh_acum"]),
                    help=f"Construcción {c.num(ind['hh_acum_constr'])} · Precom {c.num(ind['hh_acum_precom'])}")
        k[2].metric("Dotación promedio", c.num(ind["dot_prom"]),
                    help=f"Directa {c.num(ind['dot_prom_dir'])} · Indirecta {c.num(ind['dot_prom_ind'])} · Máx. {c.num(ind['dot_max'])}")
        k[3].metric("Peak HH diario", c.num(ind["pico_hh"]), help=f"El {c.fecha(ind['pico_fecha'])}" if ind["pico_fecha"] else None)
        k[4].metric("HM equipos del período", c.num(ind["hm"]),
                    delta=f"{c.num(ind['hm'] - ind['hm_previo'])} vs período anterior" if ind["hm_previo"] else None)
        pf, pf_total = T.valor_cierre(tm, "PF PERIODO"), T.valor_cierre(tm, "PF TOTAL")
        av = tm.resumen.get("avance")
        if pf or (isinstance(av, pd.DataFrame) and not av.empty):
            k2 = st.columns(4)
            if isinstance(av, pd.DataFrame) and not av.empty:
                fila = lambda t: av[av["Curva"].map(T.norm).str.startswith(T.norm(t))]
                prog, real = fila("Avance Programado"), fila("Avance Real")
                if not real.empty and not prog.empty:
                    r_, p_ = float(real["% acumulado"].iloc[0]) * 100, float(prog["% acumulado"].iloc[0]) * 100
                    k2[0].metric(f"Avance real al {c.fecha(tm.resumen.get('avance_fecha'))}", c.porc(r_, 1),
                                 delta=f"{c.num(r_ - p_, 1)} pp vs programado", delta_color="normal")
                    k2[1].metric("Avance programado Rev.0", c.porc(p_, 1))
            if pf:
                k2[2].metric("PF período construcción", c.num(pf, 2), help="HH gastadas / HH ganadas (hoja Resumen)")
            if pf_total:
                k2[3].metric("PF total semana", c.num(pf_total, 2))
    
        t_mo, t_eq, t_av, t_inf, t_dat = st.tabs(["Mano de obra", "Equipos", "Avance y cierre", "📄 Informe técnico", "Datos"])
        with t_mo:
            _mano_obra(mo, dot)
        with t_eq:
            _equipos(eq)
        with t_av:
            _avance(tm, hasta)
        with t_inf:
            _informe(tm, desde, hasta, ambitos, ind, mo, eq)
        with t_dat:
            st.download_button("Descargar datos en Excel", analisis.a_excel({
                "Mano de obra": mo.assign(semana=[CAL.etiqueta(f) for f in mo["fecha"]]), "Por semana": T.semanal(mo, eq), "Dotación": dot, "Equipos": eq, "HH por cargo": T.por_cargo(mo), "HM por equipo": T.por_equipo(eq)}),
                file_name=f"TM_{tm.contrato}_{desde}_{hasta}.xlsx", key="tm_dl_xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.dataframe(mo, hide_index=True, use_container_width=True, height=400)
    
    
    def _subir():
        if not c.exigir("cargar"):
            return
        archivo = st.file_uploader("Consolidado Times & Material (.xlsx)", type=["xlsx", "xlsm"], key="tm_up")
        if archivo is None:
            return
        try:
            tm = T.leer_tm(archivo.getvalue(), archivo.name)
        except Exception as e:
            st.error(f"No se pudo leer el archivo: {e}")
            return
        st.success(f"Datos del {c.fecha(tm.fecha_inicio)} al {c.fecha(tm.fecha_corte)} · {tm.mo['cargo'].nunique()} cargos · "
                   f"{tm.eq['equipo'].nunique()} equipos · contrato {tm.contrato or '–'}")
        for av in tm.avisos:
            st.caption("⚠️ " + av)
        if st.button("Guardar y ver dashboard", type="primary", key="tm_guardar"):
            db.guardar_tm(archivo.name, str(tm.fecha_corte) if tm.fecha_corte else None, tm.contrato, tm.a_dict(), c.nombre_usuario(),
                        contrato_id=c.contrato_id())
            st.rerun()
    
    
    def _mano_obra(mo: pd.DataFrame, dot: pd.DataFrame):
        if mo.empty:
            st.info("No hay registros de mano de obra en el período.")
            return
        dm, dd = T.diario_mo(mo), T.dotacion_diaria(dot)
        fig = go.Figure()
        for tipo in ("Directo", "Indirecto"):
            fig.add_bar(x=dm["fecha"], y=dm[tipo], name=f"HH {tipo.lower()}s", marker_color=COLORES_TIPO[tipo])
        if not dd.empty:
            fig.add_scatter(x=dd["fecha"], y=dd["Total"], name="Dotación total", yaxis="y2", mode="lines+markers",
                            line=dict(color=c.VERDE, width=2))
        fig.update_layout(barmode="stack", title="HH diarias y dotación", yaxis=dict(title="HH"),
                          yaxis2=dict(title="Personas", overlaying="y", side="right", showgrid=False),
                          legend=dict(orientation="h", y=-0.15))
        st.plotly_chart(c.layout_grafico(fig, 380), use_container_width=True, key="g_tm_1")
    
        a, b = st.columns(2)
        with a:
            amb = mo.groupby(["ambito", "tipo"])["hh"].sum().reset_index()
            fig = px.sunburst(amb, path=["ambito", "tipo"], values="hh", title="Distribución de HH",
                              color="tipo", color_discrete_map=COLORES_TIPO)
            fig.update_traces(textinfo="label+percent entry")
            st.plotly_chart(c.layout_grafico(fig, 360), use_container_width=True, key="g_tm_2")
        with b:
            m = T.mensual(mo, pd.DataFrame(columns=["fecha", "hm"]))
            fig = px.bar(m, x="mes", y=["Directo", "Indirecto"], title="HH por mes", barmode="stack",
                         color_discrete_map=COLORES_TIPO, labels={"value": "HH", "mes": "", "variable": ""})
            st.plotly_chart(c.layout_grafico(fig, 360), use_container_width=True, key="g_tm_3")
    
        sem = T.semanal(mo, pd.DataFrame(columns=["fecha", "hm"]))
        if len(sem) > 1:
            fig = px.bar(sem, x="semana", y=["Directo", "Indirecto"], barmode="stack", title="HH por semana del proyecto",
                         color_discrete_map=COLORES_TIPO, labels={"value": "HH", "semana": "", "variable": ""},
                         hover_data={"desde": True, "hasta": True})
            st.plotly_chart(c.layout_grafico(fig, 340), use_container_width=True, key="g_tm_4")
    
        cargos = T.por_cargo(mo)
        tipo = st.radio("Ranking de cargos", ["Directo", "Indirecto"], horizontal=True, key="tm_rank_tipo")
        top = cargos[cargos["tipo"] == tipo].head(15).sort_values("hh")
        fig = px.bar(top, x="hh", y="cargo", orientation="h", text=[c.num(v) for v in top["hh"]], title=f"Top 15 cargos ({tipo.lower()}s)",
                     color_discrete_sequence=[COLORES_TIPO[tipo]], labels={"hh": "HH del período", "cargo": ""})
        st.plotly_chart(c.layout_grafico(fig, max(320, 26 * len(top))), use_container_width=True, key="g_tm_5")
        st.dataframe(cargos, hide_index=True, use_container_width=True, height=380,
                     column_config={"cargo": st.column_config.TextColumn("Cargo", width="large"), "tipo": "Tipo", "ambito": "Ámbito",
                                    "hh": st.column_config.NumberColumn("HH", format="%.0f"), "dias": "Días",
                                    "hh_dia": st.column_config.NumberColumn("HH/día", format="%.1f"),
                                    "participacion": st.column_config.ProgressColumn("% del total", format="%.1f%%", min_value=0,
                                                                                     max_value=float(max(cargos["participacion"].max(), 1)))})
    
    
    def _equipos(eq: pd.DataFrame):
        if eq.empty:
            st.info("No hay registros de horas máquina en el período.")
            return
        de = eq.groupby(["fecha", "grupo"])["hm"].sum().reset_index()
        fig = px.bar(de, x="fecha", y="hm", color="grupo", title="HM diarias por grupo", color_discrete_map=COLORES_GRUPO,
                     labels={"hm": "HM", "fecha": "", "grupo": ""})
        fig.update_layout(legend=dict(orientation="h", y=-0.15))
        st.plotly_chart(c.layout_grafico(fig, 360), use_container_width=True, key="g_tm_6")
        a, b = st.columns([1, 1.4])
        with a:
            g = eq.groupby("grupo")["hm"].sum().reset_index()
            fig = px.pie(g, names="grupo", values="hm", hole=0.5, title="HM por grupo", color="grupo", color_discrete_map=COLORES_GRUPO)
            st.plotly_chart(c.layout_grafico(fig, 340), use_container_width=True, key="g_tm_7")
        equipos = T.por_equipo(eq)
        with b:
            top = equipos.head(12).sort_values("hm")
            fig = px.bar(top, x="hm", y="equipo", color="grupo", orientation="h", title="Equipos con más HM",
                         color_discrete_map=COLORES_GRUPO, labels={"hm": "HM", "equipo": "", "grupo": ""})
            fig.update_layout(showlegend=False)
            st.plotly_chart(c.layout_grafico(fig, 340), use_container_width=True, key="g_tm_8")
        st.dataframe(equipos, hide_index=True, use_container_width=True, height=360,
                     column_config={"equipo": "Equipo", "grupo": "Grupo", "hm": st.column_config.NumberColumn("HM", format="%.0f"),
                                    "dias": "Días con uso", "hm_dia": st.column_config.NumberColumn("HM/día", format="%.1f")})
    
    
    def _avance(tm, hasta):
        av, ci, te, mc = (tm.resumen.get(k) for k in ("avance", "cierre", "terminos", "contratados"))
        if not any(isinstance(x, pd.DataFrame) and not x.empty for x in (av, ci, te, mc)):
            st.info("El archivo no trae la hoja Resumen con avance y cierre semanal.")
            return
        st.caption(f"Valores tomados de la hoja Resumen ({tm.resumen.get('titulo', '')}).")
        if isinstance(av, pd.DataFrame) and not av.empty:
            st.subheader(f"Avance del programa al {c.fecha(tm.resumen.get('avance_fecha'))}")
            if tm.resumen.get("avance_notas"):
                st.caption(" · ".join(tm.resumen["avance_notas"][:2]))
            curvas = av[~av["Curva"].astype(str).str.contains("∆|Δ|v/s", regex=True)].copy()
            curvas["%"] = pd.to_numeric(curvas["% acumulado"], errors="coerce") * 100
            fig = px.bar(curvas, x="%", y="Curva", orientation="h", text=[c.porc(v, 1) for v in curvas["%"]],
                         color_discrete_sequence=[c.NARANJO], labels={"%": "% acumulado", "Curva": ""})
            fig.update_xaxes(range=[0, 110])
            st.plotly_chart(c.layout_grafico(fig, 280), use_container_width=True, key="g_tm_9")
            st.dataframe(av.assign(**{"% acumulado": pd.to_numeric(av["% acumulado"], errors="coerce") * 100,
                                      "% parcial": pd.to_numeric(av["% parcial"], errors="coerce") * 100}),
                         hide_index=True, use_container_width=True,
                         column_config={"HH acumulado": st.column_config.NumberColumn(format="%.0f"),
                                        "HH parcial": st.column_config.NumberColumn(format="%.0f"),
                                        "% acumulado": st.column_config.NumberColumn(format="%.1f%%"),
                                        "% parcial": st.column_config.NumberColumn(format="%.2f%%")})
        a, b = st.columns(2)
        if isinstance(ci, pd.DataFrame) and not ci.empty:
            with a:
                st.subheader(f"Cierre semanal al {c.fecha(tm.resumen.get('cierre_fecha'))}")
                st.dataframe(ci.assign(Valor=pd.to_numeric(ci["Valor"], errors="coerce")), hide_index=True, use_container_width=True,
                             column_config={"Valor": st.column_config.NumberColumn(format="%.2f")})
        if isinstance(te, pd.DataFrame) and not te.empty:
            with b:
                st.subheader("Términos proyectados")
                t = te.copy()
                t["Días"] = [(pd.Timestamp(x) - pd.Timestamp(hasta)).days if x else None for x in t["Término proyectado"]]
                t["Estado"] = ["Sin fecha" if d is None else "🔴 Fecha cumplida" if d < 0 else "🟠 Próximos 15 días" if d <= 15
                               else "🟢 En plazo" for d in t["Días"]]
                st.dataframe(t, hide_index=True, use_container_width=True,
                             column_config={"Término proyectado": st.column_config.DateColumn(format="DD-MM-YYYY")})
        if isinstance(mc, pd.DataFrame) and not mc.empty:
            st.subheader("Dotación contratada promedio por mes")
            st.dataframe(mc, hide_index=True, use_container_width=True)
    
    
    def _informe(tm, desde, hasta, ambitos, ind, mo, eq):
        st.markdown(f"Informe técnico del **{c.fecha(desde)}** al **{c.fecha(hasta)}** ({', '.join(ambitos)}). Usa los filtros de "
                    "arriba para cambiar el período o el ámbito.")
        narrativa = T.resumen_narrativo(tm, ind, T.por_cargo(mo), T.por_equipo(eq))
        st.markdown("**Resumen ejecutivo (automático)**")
        for frase in narrativa:
            st.markdown(f"- {frase}")
        titulo = st.text_input("Título", "INFORME TÉCNICO TIMES & MATERIAL", key="tm_inf_titulo")
        obs = st.text_area("Observaciones (opcional)", key="tm_inf_obs", height=110)
        rec = st.text_area("Conclusiones y recomendaciones (opcional)", key="tm_inf_rec", height=110)
        anexo = st.checkbox("Incluir anexos con el detalle por cargo y por equipo", value=True, key="tm_inf_anexo")
        if st.button("Generar informe técnico (PDF)", type="primary", key="tm_inf_btn"):
            with st.spinner("Generando…"):
                st.session_state["tm_inf_pdf"] = tm_pdf.generar_informe_tecnico(
                    tm, desde, hasta, ambitos, (c.usuario() or {}).get("nombre", ""), obs, rec, anexo,
                    logo="logo_syncore.png" if os.path.exists("logo_syncore.png") else None, titulo=titulo)
                st.session_state["tm_inf_nombre"] = f"Informe_tecnico_TM_{tm.contrato}_{desde:%d-%m-%Y}_al_{hasta:%d-%m-%Y}.pdf"
        if st.session_state.get("tm_inf_pdf"):
            st.download_button("📄 Descargar informe técnico", st.session_state["tm_inf_pdf"],
                               file_name=st.session_state.get("tm_inf_nombre", "Informe_tecnico_TM.pdf"), mime="application/pdf",
                               key="tm_inf_dl")

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


vtm = _modulo_vtm()

# ======================================================================
# MÓDULO: vdesv
# ======================================================================
def _modulo_vdesv():
    import datetime as dt
    import os
    
    import pandas as pd
    import plotly.express as px
    import plotly.graph_objects as go
    import streamlit as st
    
    D = desviaciones
    DP = desviaciones_pdf
    T = tm
    c = comun
    
    COLORES_JORNADA = {"Día": c.NARANJO, "Noche": c.AZUL, "General": c.PIZARRA, "Sin jornada": c.GRIS}
    LOGO = "logo_syncore.png"
    
    
    @st.cache_resource(show_spinner=False, max_entries=6)
    def _log(cid: int) -> D.LogDesviaciones:
        return D.LogDesviaciones.de_dict(db.carga_desviaciones(cid))
    
    
    @st.cache_resource(show_spinner=False, max_entries=2)
    def _tm(cid: int):
        return T.ConsolidadoTM.de_dict(db.carga_tm(cid))
    
    
    def _ultimo_tm():
        try:
            cargas = db.cargas_tm(c.contrato_id())
        except Exception:
            return None
        if cargas.empty:
            return None
        return _tm(int(cargas.sort_values(["fecha_corte", "id"], ascending=False).iloc[0]["id"]))
    
    
    def vista():
        st.title("Desviación de recursos")
        cargas = db.cargas_desviaciones(c.contrato_id())
        with st.expander("Cargar LOG de desviación de recursos", expanded=cargas.empty):
            _subir()
        if cargas.empty:
            st.info("Carga el LOG de desviación de recursos (.xlsx con la hoja de detalle: Fecha, Jornada, Especialidad, Área, "
                    "Actividad, Dotación, HH Desviadas, Referencia).")
            return
    
        etiquetas = {int(r.id): f"{r.archivo} · al {c.fecha(r.fecha_corte)}" + (f" · rev {r.revision}" if r.revision else "") +
                     f" · {c.num(r.hh)} HH · cargado {str(r.cargado_en)[:16]}" for r in cargas.itertuples()}
        a, b = st.columns([3, 1])
        cid = int(a.selectbox("LOG cargado", list(etiquetas), format_func=etiquetas.get, key="dv_carga"))
        log = _log(cid)
        if c.usuario() and c.usuario()["rol"] == "admin":
            with b.popover("Eliminar carga"):
                if st.button("Eliminar esta carga", key="dv_borrar"):
                    db.borrar_carga_desviaciones(cid)
                    st.cache_resource.clear()
                    st.rerun()
    
        # ---- filtros
        corte = log.fecha_corte or dt.date.today()
        inicio = log.fecha_inicio or corte
        opciones = {"Todo el LOG": (inicio, corte), "Mes de la fecha de corte": (corte.replace(day=1), corte),
                    "Últimos 7 días": (corte - dt.timedelta(days=6), corte), "Últimos 30 días": (corte - dt.timedelta(days=29), corte),
                    "Personalizado": None}
        f1, f2, f3, f4, f5 = st.columns([1.2, 1.4, 1.3, 1.3, 1])
        per = f1.selectbox("Período", list(opciones), key="dv_periodo")
        if opciones[per] is None:
            rango = f2.date_input("Desde / hasta", (inicio, corte), min_value=inicio, max_value=corte, format="DD-MM-YYYY", key="dv_rango")
            desde, hasta = rango if isinstance(rango, (list, tuple)) and len(rango) == 2 else (inicio, corte)
        else:
            desde, hasta = opciones[per]
            desde = max(desde, inicio)
            f2.caption(f"Del **{c.fecha(desde)}** al **{c.fecha(hasta)}**")
        esp = f3.multiselect("Especialidad", sorted(x for x in log.df["especialidad"].unique() if x), key="dv_esp")
        mot = f4.multiselect("Motivo", sorted(x for x in log.df["referencia"].unique() if x), key="dv_mot")
        jor = f5.multiselect("Jornada", sorted(x for x in log.df["jornada"].unique() if x), key="dv_jor")
        d = D.filtrar(log, desde, hasta, esp, mot, jor)
        filtros = " · ".join(x for x in (", ".join(esp), ", ".join(mot), ", ".join(jor)) if x)
    
        dif = D.diferencias_totales(log)
        if dif:
            partes = {}
            for t in dif:
                partes.setdefault(t["campo"], (t["planilla"], t["calculado"], []))[2].append(t["origen"].split(" (")[0].replace("Hoja ", ""))
            st.warning("**Los totales de la planilla no cuadran con el detalle:** " + "; ".join(
                f"{D.NOMBRES[k]} {c.num(p, 1)} en la planilla ({', '.join(dict.fromkeys(o))}) frente a {c.num(q, 1)} en el detalle"
                for k, (p, q, o) in partes.items()) + ". Probablemente la tabla dinámica no se actualizó; la app usa el detalle.")
        if len(log.calidad):
            st.caption(f"⚠️ {len(log.calidad)} observaciones de calidad en el LOG · revisa la pestaña **Control de calidad**.")
    
        # ---- indicadores
        ind = D.indicadores(d, desde, hasta)
        largo = (hasta - desde).days + 1
        previo = D.filtrar(log, desde - dt.timedelta(days=largo), desde - dt.timedelta(days=1), esp, mot, jor)
        ind_prev = D.indicadores(previo, desde - dt.timedelta(days=largo), desde - dt.timedelta(days=1)) if len(previo) else None
        tm = _ultimo_tm()
        cruce = D.cruce_tm(d, tm.mo) if tm is not None else None
        valor_hh = float(db.config(c.contrato_id(), "desv_valor_hh", 0) or 0)
    
        k = st.columns(3)
        delta = None
        if ind_prev and ind_prev["hh_dia"]:
            delta = f"{c.num((ind['hh_dia'] / ind_prev['hh_dia'] - 1) * 100, 1)}% HH/día vs período anterior"
        k[0].metric("HH desviadas", c.num(ind["hh"]), delta=delta, delta_color="inverse")
        k[1].metric("HH promedio por día", c.num(ind["hh_dia"], 1),
                    help=f"Peak {c.num(ind['pico_hh'], 1)} HH el {c.fecha(ind['pico_fecha'])}" if ind["pico_fecha"] else None)
        if cruce and cruce.get("pct") is not None:
            k[2].metric("% de las HH directas T&M", c.porc(cruce["pct"], 1),
                        help=f"{c.num(cruce['hh_desviadas'])} HH desviadas / {c.num(cruce['hh_tm'])} HH de mano de obra directa de "
                             f"construcción, del {c.fecha(cruce['desde'])} al {c.fecha(cruce['hasta'])} (último consolidado T&M cargado).")
        elif valor_hh:
            k[2].metric("Valorización referencial", f"${c.num(ind['hh'] * valor_hh / 1e6, 1)} MM", help=f"A ${c.num(valor_hh)} por HH")
        else:
            k[2].metric("HH gestionadas", c.num(ind["hh_gestionadas"], 1))
    
        t1, t2, t3, t4, t5, t6 = st.tabs(["Resumen", "Especialidad y área", "Detalle", "Control de calidad", "Comparar revisiones",
                                          "📄 Informe PDF"])
        with t1:
            _resumen(d, cruce)
        with t2:
            _especialidad_area(d)
        with t3:
            _detalle(log, d, desde, hasta)
        with t4:
            _calidad(log)
        with t5:
            _comparar(log, cid, cargas)
        with t6:
            _informe(log, d, desde, hasta, filtros, ind, cruce, valor_hh, tm)
    
    
    def _subir():
        if not c.exigir("cargar"):
            return
        archivo = st.file_uploader("LOG de desviación de recursos (.xlsx)", type=["xlsx", "xlsm"], key="dv_up")
        if archivo is None:
            return
        try:
            log = D.leer_log(archivo.getvalue(), archivo.name)
        except Exception as e:
            st.error(f"No se pudo leer el archivo: {e}")
            return
        st.success(f"{c.num(len(log.df))} registros del {c.fecha(log.fecha_inicio)} al {c.fecha(log.fecha_corte)} · "
                   f"{c.num(log.df['hh'].sum())} HH desviadas" + (f" · revisión {log.revision}" if log.revision else "") +
                   f" · hoja «{log.hoja.strip()}»")
        for t in D.diferencias_totales(log):
            st.caption(f"⚠️ {t['origen']}: {c.num(t['planilla'], 1)} {D.NOMBRES[t['campo']]}; el detalle suma {c.num(t['calculado'], 1)}.")
        if log.unificaciones:
            st.caption(f"Se unificarán {len(log.unificaciones)} nombres escritos de distintas formas (detalle en Control de calidad).")
        if len(log.calidad):
            st.caption(f"{len(log.calidad)} observaciones de calidad (sin motivo, posibles duplicados, etc.).")
        for av in log.avisos:
            st.caption("⚠️ " + av)
        if st.button("Guardar y ver dashboard", type="primary", key="dv_guardar"):
            db.guardar_desviaciones(archivo.name, str(log.fecha_corte) if log.fecha_corte else None, log.revision,
                                    float(log.df["hh"].sum()), len(log.df), log.a_dict(), c.nombre_usuario(),
                                    contrato_id=c.contrato_id())
            st.rerun()
    
    
    def _resumen(d: pd.DataFrame, cruce):
        if d.empty:
            st.info("No hay registros en el período y filtros seleccionados.")
            return
        dia = D.diario(d)
        fig = go.Figure()
        for j in [x for x in dia.columns if x not in ("fecha", "Total", "Acumulado")]:
            fig.add_bar(x=dia["fecha"], y=dia[j], name=j, marker_color=COLORES_JORNADA.get(j, c.GRIS))
        fig.add_scatter(x=dia["fecha"], y=dia["Acumulado"], name="Acumulado", yaxis="y2", mode="lines", line=dict(color=c.VERDE, width=2))
        fig.update_layout(barmode="stack", title="HH desviadas por día y acumulado", yaxis=dict(title="HH por día"),
                          yaxis2=dict(title="HH acumuladas", overlaying="y", side="right", showgrid=False),
                          legend=dict(orientation="h", y=-0.15))
        st.plotly_chart(c.layout_grafico(fig, 380), use_container_width=True, key="g_desviaciones_1")
    
        a, b = st.columns(2)
        mot = D.por(d, "referencia", "(sin motivo)")
        with a:
            fig = px.bar(mot.sort_values("hh"), x="hh", y="referencia", orientation="h", title="HH por motivo",
                         text=[f"{c.num(v)} · {c.porc(p, 1)}" for v, p in zip(mot.sort_values('hh')["hh"], mot.sort_values('hh')["participacion"])],
                         color_discrete_sequence=[c.NARANJO], labels={"hh": "HH", "referencia": ""})
            st.plotly_chart(c.layout_grafico(fig, max(320, 30 * len(mot))), use_container_width=True, key="g_desviaciones_2")
        esp = D.por(d, "especialidad")
        with b:
            fig = px.bar(esp.sort_values("hh"), x="hh", y="especialidad", orientation="h", title="HH por especialidad",
                         text=[c.num(v) for v in esp.sort_values("hh")["hh"]], color_discrete_sequence=[c.AZUL],
                         labels={"hh": "HH", "especialidad": ""})
            st.plotly_chart(c.layout_grafico(fig, max(320, 30 * len(esp))), use_container_width=True, key="g_desviaciones_3")
    
        a, b = st.columns(2)
        with a:
            st.markdown("**HH por semana**")
            st.dataframe(D.semanal(d).drop(columns=["desde"]), hide_index=True, use_container_width=True,
                         column_config={"semana": "Semana", "registros": "Registros",
                                        "hh": st.column_config.NumberColumn("HH desviadas", format="%.1f")})
        with b:
            jor = D.por(d, "jornada")
            fig = px.pie(jor, names="jornada", values="hh", hole=0.55, title="HH por jornada", color="jornada",
                         color_discrete_map=COLORES_JORNADA)
            st.plotly_chart(c.layout_grafico(fig, 300), use_container_width=True, key="g_desviaciones_4")
    
        if cruce and cruce.get("pct") is not None:
            s = cruce["serie"]
            fig = go.Figure()
            fig.add_bar(x=s["fecha"], y=s["HH directas T&M"], name="HH directas construcción (T&M)", marker_color=c.PIZARRA, opacity=0.55)
            fig.add_bar(x=s["fecha"], y=s["HH desviadas"], name="HH desviadas", marker_color=c.NARANJO)
            fig.add_scatter(x=s["fecha"], y=s["% desviado"], name="% desviado", yaxis="y2", mode="lines+markers",
                            line=dict(color=c.VERDE, width=2))
            fig.update_layout(barmode="overlay", title=f"HH desviadas frente a las HH directas del T&M ({c.porc(cruce['pct'], 1)} en el período)",
                              yaxis=dict(title="HH"), yaxis2=dict(title="%", overlaying="y", side="right", showgrid=False),
                              legend=dict(orientation="h", y=-0.15))
            st.plotly_chart(c.layout_grafico(fig, 360), use_container_width=True, key="g_desviaciones_5")
    
    
    def _especialidad_area(d: pd.DataFrame):
        if d.empty:
            st.info("No hay registros en el período y filtros seleccionados.")
            return
        m = D.matriz(d)
        if not m.empty:
            mm = m.drop(columns="Total")
            fig = px.imshow(mm, text_auto=".0f", aspect="auto", color_continuous_scale=["#FFF3EA", c.NARANJO, "#9A3412"],
                            labels=dict(x="Motivo", y="Especialidad", color="HH"), title="HH por especialidad y motivo")
            st.plotly_chart(c.layout_grafico(fig, max(320, 36 * len(mm) + 120)), use_container_width=True, key="g_desviaciones_6")
        a, b = st.columns(2)
        num = {"hh": st.column_config.NumberColumn("HH", format="%.1f"), "participacion": st.column_config.NumberColumn("% del total", format="%.1f%%"),
               "registros": "Registros", "dotacion": st.column_config.NumberColumn("Dotación", format="%.0f"),
               "hh_gestionadas": st.column_config.NumberColumn("HH gestionadas", format="%.1f")}
        with a:
            st.markdown("**Por especialidad**")
            st.dataframe(D.por(d, "especialidad"), hide_index=True, use_container_width=True,
                         column_config={"especialidad": "Especialidad", **num})
        with b:
            st.markdown("**Por área**")
            st.dataframe(D.por(d, "area"), hide_index=True, use_container_width=True, height=420, column_config={"area": "Área", **num})
        st.markdown("**Registros con más HH desviadas**")
        st.dataframe(D.tabla_detalle(D.mayores(d, 15)), hide_index=True, use_container_width=True,
                     column_config={"Fecha": st.column_config.DateColumn(format="DD-MM-YYYY")})
    
    
    def _detalle(log, d: pd.DataFrame, desde, hasta):
        buscar = st.text_input("Buscar en actividad, área u observación", key="dv_buscar", placeholder="Ej.: RFI-183, Caldera 2, soportes")
        t = d
        if buscar.strip():
            q = buscar.strip().lower()
            t = d[d[["actividad", "area", "observacion", "referencia", "n"]].astype(str).apply(
                lambda col: col.str.lower().str.contains(q, regex=False)).any(axis=1)]
        st.caption(f"{c.num(len(t))} registros · {c.num(t['hh'].sum(), 1)} HH")
        st.dataframe(D.tabla_detalle(t), hide_index=True, use_container_width=True, height=480,
                     column_config={"Fecha": st.column_config.DateColumn(format="DD-MM-YYYY"),
                                    "Actividad": st.column_config.TextColumn(width="large")})
        st.download_button("Descargar en Excel (detalle y resúmenes)", analisis.a_excel({
            "Detalle": D.tabla_detalle(t), "Por motivo": D.por(t, "referencia", "(sin motivo)"),
            "Por especialidad": D.por(t, "especialidad"), "Por área": D.por(t, "area"), "Por día": D.diario(t),
            "Especialidad x motivo": D.matriz(t).reset_index() if not t.empty else pd.DataFrame(),
            "Control de calidad": log.calidad}),
            file_name=f"Desviaciones_{desde:%d-%m-%Y}_al_{hasta:%d-%m-%Y}.xlsx", key="dv_dl_xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    
    def _calidad(log):
        tot = pd.DataFrame(log.totales)
        if not tot.empty:
            st.markdown("**Totales de la planilla frente al detalle**")
            tot = tot.assign(campo=tot["campo"].map(D.NOMBRES), diferencia=tot["calculado"] - tot["planilla"])
            tot["estado"] = ["✅ Cuadra" if abs(x) <= 0.5 else "❌ No cuadra" for x in tot["diferencia"]]
            st.dataframe(tot, hide_index=True, use_container_width=True,
                         column_config={"origen": "Origen", "campo": "Dato", "planilla": st.column_config.NumberColumn("Planilla", format="%.1f"),
                                        "calculado": st.column_config.NumberColumn("Detalle", format="%.1f"),
                                        "diferencia": st.column_config.NumberColumn("Diferencia", format="%.1f"), "estado": "Estado"})
        st.markdown(f"**Observaciones en los registros** ({len(log.calidad)})")
        if log.calidad.empty:
            st.success("No se encontraron registros con problemas.")
        else:
            st.dataframe(log.calidad, hide_index=True, use_container_width=True,
                         column_config={"n": "N°", "fila_excel": "Fila Excel", "fecha": st.column_config.DateColumn("Fecha", format="DD-MM-YYYY"),
                                        "problema": "Observación", "detalle": "Detalle"})
        st.markdown(f"**Nombres unificados** ({len(log.unificaciones)})")
        if log.unificaciones:
            st.caption("Variantes que solo diferían en mayúsculas, tildes o espacios se agrupan bajo un mismo nombre.")
            for u in log.unificaciones:
                st.markdown(f"- {u}")
    
    
    def _comparar(log, cid: int, cargas: pd.DataFrame):
        otras = cargas[cargas["id"] != cid]
        if otras.empty:
            st.info("Carga otra revisión del LOG para ver qué registros se agregaron, cambiaron o se eliminaron.")
            return
        previas = otras[otras["fecha_corte"].fillna("") <= (str(log.fecha_corte) if log.fecha_corte else "9999")]
        defecto = int((previas if not previas.empty else otras).iloc[0]["id"])
        ids = list(otras["id"].astype(int))
        etiqueta = {int(r.id): f"{r.archivo} · al {c.fecha(r.fecha_corte)}" + (f" · rev {r.revision}" if r.revision else "") +
                    f" · cargado {str(r.cargado_en)[:16]}" for r in otras.itertuples()}
        otro = int(st.selectbox("Comparar con", ids, index=ids.index(defecto), format_func=etiqueta.get, key="dv_cmp"))
        anterior = _log(otro)
        cmp_ = D.comparar(log, anterior)
        k = st.columns(4)
        k[0].metric("Registros agregados", c.num((cmp_["cambio"] == "Agregado").sum()))
        k[1].metric("Modificados", c.num((cmp_["cambio"] == "Modificado").sum()))
        k[2].metric("Eliminados", c.num((cmp_["cambio"] == "Eliminado").sum()))
        k[3].metric("Diferencia de HH", c.num(log.df["hh"].sum() - anterior.df["hh"].sum(), 1))
        if cmp_.empty:
            st.success("Las dos revisiones tienen los mismos registros.")
            return
        st.dataframe(cmp_, hide_index=True, use_container_width=True, height=440,
                     column_config={"cambio": "Cambio", "n": "N°", "fecha": st.column_config.DateColumn("Fecha", format="DD-MM-YYYY"),
                                    "especialidad": "Especialidad", "actividad": st.column_config.TextColumn("Actividad", width="large"),
                                    "hh_antes": st.column_config.NumberColumn("HH antes", format="%.1f"),
                                    "hh_ahora": st.column_config.NumberColumn("HH ahora", format="%.1f"), "detalle": "Qué cambió"})
    
    
    def _informe(log, d, desde, hasta, filtros, ind, cruce, valor_hh, tm):
        st.markdown(f"Informe del **{c.fecha(desde)}** al **{c.fecha(hasta)}**" + (f" (filtros: {filtros})" if filtros else "") +
                    ". Usa los filtros de arriba para cambiar el período, la especialidad o el motivo.")
        a, b = st.columns(2)
        nuevo_valor = a.number_input("Valor HH referencial ($, neto) para valorizar", min_value=0.0, value=valor_hh, step=1000.0,
                                     format="%.0f", key="dv_valor_hh", help="Déjalo en 0 si no quieres valorizar. Se recuerda para la próxima vez.")
        if nuevo_valor != valor_hh and auth.puede(c.usuario(), "cargar"):
            db.guardar_config(c.contrato_id(), "desv_valor_hh", nuevo_valor)
            valor_hh = nuevo_valor
        usar_tm = b.checkbox("Incluir comparación con el Times & Material", value=cruce is not None, disabled=cruce is None, key="dv_inf_tm",
                             help=None if cruce else "Carga un consolidado T&M que cubra estas fechas para habilitarla.")
        narr = D.resumen_narrativo(log, d, ind, desde, hasta, cruce if usar_tm else None, valor_hh or None)
        st.markdown("**Resumen ejecutivo (automático)**")
        for frase in narr:
            st.markdown(f"- {frase}")
        titulo = st.text_input("Título", "INFORME DE DESVIACIÓN DE RECURSOS", key="dv_inf_titulo")
        obs = st.text_area("Observaciones (opcional)", key="dv_inf_obs", height=100)
        concl = st.text_area("Conclusiones y acciones (opcional)", key="dv_inf_concl", height=100)
        anexo = st.checkbox("Incluir anexo con el detalle de todos los registros, agrupados por motivo", value=True, key="dv_inf_anexo")
        if st.button("Generar informe (PDF)", type="primary", key="dv_inf_btn"):
            with st.spinner("Generando…"):
                st.session_state["dv_inf_pdf"] = DP.generar_informe(
                    log, d, desde, hasta, filtros, (c.usuario() or {}).get("nombre", ""), tm.contrato if tm is not None else "",
                    obs, concl, cruce if usar_tm else None, valor_hh or None, anexo, LOGO if os.path.exists(LOGO) else None, titulo)
                st.session_state["dv_inf_nombre"] = f"Informe_desviacion_recursos_{desde:%d-%m-%Y}_al_{hasta:%d-%m-%Y}.pdf"
        if st.session_state.get("dv_inf_pdf"):
            st.download_button("📄 Descargar informe", st.session_state["dv_inf_pdf"], mime="application/pdf",
                               file_name=st.session_state.get("dv_inf_nombre", "Informe_desviacion_recursos.pdf"), key="dv_inf_dl")

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


vdesv = _modulo_vdesv()

# ======================================================================
# MÓDULO: admin
# ======================================================================
def _modulo_admin():
    import streamlit as st
    import datetime as dt
    
    import pandas as pd
    
    CAL = calendario
    DF = diario
    c = comun
    
    
    def proyectos():
        st.title("Proyectos y cortes")
        cfg = CAL.configuracion()
        hoy = CAL.semana(dt.date.today(), cfg)
        st.caption(f"Contrato **{c.contrato()['codigo']}** · inicio del proyecto: **{cfg[0]:%d-%m-%Y}** · hoy es la **{CAL.etiqueta(dt.date.today(), cfg, True)}**. Todas las "
                   "semanas de la app (informe semanal, T&M, desviaciones, 6WLA, PDF) se cuentan con este calendario.")
        with st.expander("Calendario de semanas del proyecto"):
            st.dataframe(CAL.tabla(max(1, hoy - 8), hoy + 6, cfg), hide_index=True, use_container_width=True,
                         column_config={"Desde": st.column_config.DateColumn(format="DD-MM-YYYY"),
                                        "Hasta (corte)": st.column_config.DateColumn(format="DD-MM-YYYY")})
            if auth.puede(c.usuario(), "administrar"):
                st.caption("Cada fila fija la fecha de corte de una semana; las siguientes avanzan de 7 en 7 días hasta la próxima fila. "
                           "Agrega una fila cuando cambie el día de corte (por ejemplo, la S30 cortó el viernes 09-10).")
                ini = st.date_input("Inicio del proyecto", cfg[0], format="DD-MM-YYYY", key="adm_cal_ini")
                anclas = st.data_editor(pd.DataFrame([{"Semana": n, "Fecha de corte": f} for n, f in cfg[1]]), num_rows="dynamic",
                                        hide_index=True, key="adm_cal_anclas",
                                        column_config={"Semana": st.column_config.NumberColumn(min_value=1, step=1, format="%d"),
                                                       "Fecha de corte": st.column_config.DateColumn(format="DD-MM-YYYY")})
                if st.button("Guardar calendario", key="adm_cal_btn"):
                    filas = [(int(r["Semana"]), pd.Timestamp(r["Fecha de corte"]).date()) for _, r in anclas.dropna().iterrows()]
                    try:
                        CAL.guardar(ini, filas)
                        st.cache_data.clear()
                        st.cache_resource.clear()
                        st.success("Calendario guardado.")
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))
        ps = c.proyectos()
        if ps.empty:
            st.info("No hay proyectos cargados.")
            return
        st.dataframe(ps.drop(columns=["id"]), hide_index=True, use_container_width=True,
                     column_config={"codigo": "Código", "nombre": "Nombre", "origen": "Origen", "cortes": "Cortes",
                                    "ultimo_corte": "Último corte"})
        i = st.selectbox("Ver cortes de", range(len(ps)), format_func=lambda k: ps.iloc[k]["codigo"], key="adm_proy")
        p = ps.iloc[i]
        cs = c.cortes(int(p["id"]))
        st.dataframe(
            cs[["fecha_corte", "n_actividades", "at_completion", "actual", "remaining", "avance_pct", "archivo", "cargado_por", "cargado_en"]],
            hide_index=True, use_container_width=True,
            column_config={"fecha_corte": "Data Date", "n_actividades": "Actividades", "archivo": "Archivo",
                           "cargado_por": "Cargado por", "cargado_en": "Cargado el", **c.COLUMNAS_HORAS},
        )
        if not auth.puede(c.usuario(), "administrar"):
            return
        st.divider()
        st.subheader("Eliminar")
        a, b = st.columns(2)
        with a:
            if not cs.empty:
                j = st.selectbox("Corte", range(len(cs)), format_func=lambda k: c.fecha(cs.iloc[k]["fecha_corte"]), key="adm_corte")
                if st.checkbox("Confirmo eliminar este corte", key="adm_conf_corte") and st.button("Eliminar corte", key="adm_del_corte"):
                    db.borrar_corte(int(cs.iloc[j]["id"]))
                    c.limpiar_cache()
                    st.rerun()
        with b:
            if st.checkbox(f"Confirmo eliminar {p['codigo']} con todos sus cortes", key="adm_conf_proy") and \
                    st.button("Eliminar proyecto", key="adm_del_proy"):
                db.borrar_proyecto(int(p["id"]))
                c.limpiar_cache()
                st.rerun()
    
    
    def usuarios():
        st.title("Usuarios")
        if not c.exigir("administrar"):
            return
        df = auth.listar()
        pendientes = df[df["estado"] == "Pendiente"]
        if not pendientes.empty:
            st.warning(f"{len(pendientes)} cuenta(s) esperando aprobación.")
        st.dataframe(df.drop(columns=["id"]).assign(rol=df["rol"].map(auth.ROLES)), hide_index=True, use_container_width=True,
                     column_config={"username": "Usuario", "nombre": "Nombre", "rol": "Rol", "estado": "Estado",
                                    "creado_en": "Creado"})
    
        t1, t2 = st.tabs(["Editar usuario", "Crear usuario"])
        with t1:
            i = st.selectbox("Usuario", range(len(df)), format_func=lambda k: df.iloc[k]["username"], key="usr_sel")
            u = df.iloc[i]
            es_yo = int(u["id"]) == c.usuario()["id"]
            a, b = st.columns(2)
            roles = list(auth.ROLES)
            rol = a.selectbox("Rol", roles, index=roles.index(u["rol"]), format_func=auth.ROLES.get, key=f"usr_rol_{u['id']}")
            estados = ["Aprobado", "Pendiente", "Bloqueado"]
            estado = b.selectbox("Estado", estados, index=estados.index(u["estado"]) if u["estado"] in estados else 0,
                                 key=f"usr_est_{u['id']}")
            deja_sin_admin = u["rol"] == "admin" and u["estado"] == "Aprobado" and (rol != "admin" or estado != "Aprobado") \
                and auth.n_admins_activos() <= 1
            if st.button("Guardar cambios", type="primary", key="usr_guardar"):
                if deja_sin_admin:
                    st.error("Debe quedar al menos un administrador activo.")
                else:
                    auth.actualizar(int(u["id"]), rol, estado)
                    st.success("Usuario actualizado.")
                    st.rerun()
            cts = db.contratos()
            if not cts.empty:
                if rol == "admin":
                    st.caption("Los administradores ven todos los contratos.")
                else:
                    actuales = [i for i in db.contratos_ids_de_usuario(int(u["id"])) if i in set(cts["id"].astype(int))]
                    nombres = {int(r.id): f"{r.codigo}" + ("" if r.activo else " (inactivo)") for r in cts.itertuples()}
                    sel = st.multiselect("Contratos con acceso", list(nombres), default=actuales, format_func=nombres.get,
                                         key=f"usr_cts_{u['id']}")
                    if st.button("Guardar accesos", key=f"usr_cts_btn_{u['id']}"):
                        db.asignar_contratos_a_usuario(int(u["id"]), sel)
                        st.success("Accesos actualizados.")
            with st.expander("Restablecer contraseña"):
                nueva = st.text_input("Nueva contraseña temporal", type="password", key=f"usr_pwd_{u['id']}")
                if st.button("Restablecer", key="usr_reset"):
                    try:
                        auth.cambiar_password(int(u["id"]), nueva)
                        st.success("Contraseña restablecida. Comunícasela al usuario por un canal seguro.")
                    except ValueError as e:
                        st.error(str(e))
            with st.expander("Eliminar usuario"):
                if es_yo:
                    st.caption("No puedes eliminar tu propia cuenta.")
                elif st.checkbox(f"Confirmo eliminar a {u['username']}", key=f"usr_conf_{u['id']}") and \
                        st.button("Eliminar", key="usr_del"):
                    if u["rol"] == "admin" and auth.n_admins_activos() <= 1:
                        st.error("Debe quedar al menos un administrador activo.")
                    else:
                        auth.eliminar(int(u["id"]))
                        st.rerun()
        with t2:
            with st.form("usr_nuevo", clear_on_submit=True):
                a, b = st.columns(2)
                user = a.text_input("Usuario")
                nombre = b.text_input("Nombre completo")
                pwd = a.text_input("Contraseña temporal", type="password")
                rol = b.selectbox("Rol", list(auth.ROLES), format_func=auth.ROLES.get, index=2)
                cts = db.contratos(solo_activos=True)
                nombres = {int(r.id): r.codigo for r in cts.itertuples()}
                actual = (c.contrato() or {}).get("id")
                sel = st.multiselect("Contratos con acceso (no aplica a administradores)", list(nombres),
                                     default=[actual] if actual in nombres else [], format_func=nombres.get)
                if st.form_submit_button("Crear usuario", type="primary"):
                    try:
                        auth.crear_usuario(user, pwd, nombre, rol, "Aprobado")
                        fila = db.uno("SELECT id FROM usuarios WHERE LOWER(username) = LOWER(?)", (user.strip(),))
                        if fila and sel:
                            db.asignar_contratos_a_usuario(int(fila[0]), sel)
                        st.success(f"Usuario {user} creado.")
                    except ValueError as e:
                        st.error(str(e))
    
    
    def contratos():
        st.title("Contratos")
        if not c.exigir("administrar"):
            return
        st.caption("Cada contrato tiene sus propios proyectos P6, informes semanales, informe diario, Times & Material, LOG de "
                   "desviaciones, 6WLA y calendario de semanas. Los usuarios solo ven los contratos a los que tienen acceso; "
                   "los administradores ven todos.")
        cts = db.contratos()
        usuarios = auth.listar()
        no_admin = usuarios[(usuarios["rol"] != "admin")]
        nombres_u = {int(r.id): f"{r.nombre} ({r.username})" for r in no_admin.itertuples()}
        if not cts.empty:
            tabla = cts.assign(usuarios=[len(db.usuarios_de_contrato(int(i))) for i in cts["id"]],
                               estado=["Activo" if a else "Inactivo" for a in cts["activo"]])
            st.dataframe(tabla[["codigo", "nombre", "estado", "usuarios", "creado_en"]], hide_index=True, use_container_width=True,
                         column_config={"codigo": "Código", "nombre": "Nombre", "estado": "Estado",
                                        "usuarios": "Usuarios con acceso", "creado_en": "Creado"})
    
        t_nuevo, t_editar = st.tabs(["Nuevo contrato", "Editar contrato"])
        with t_nuevo:
            with st.form("ct_nuevo", clear_on_submit=True):
                a, b = st.columns([1, 2])
                codigo = a.text_input("Código del contrato", placeholder="Ej.: B1SE-K-CC-1003")
                nombre = b.text_input("Nombre del contrato")
                a, b = st.columns(2)
                ini = a.date_input("Inicio del proyecto (primer día de la semana 1)", dt.date.today(), format="DD-MM-YYYY")
                fin1 = b.date_input("Fecha de corte de la semana 1", dt.date.today() + dt.timedelta(days=6), format="DD-MM-YYYY",
                                    help="Desde ahí las semanas avanzan de 7 en 7 días. Si después cambia el día de corte, se ajusta "
                                         "en Proyectos y cortes → Calendario de semanas.")
                acceso = st.multiselect("Usuarios con acceso", list(nombres_u), format_func=nombres_u.get)
                if st.form_submit_button("Crear contrato", type="primary"):
                    err = CAL.validar(ini, [(1, fin1)])
                    if err:
                        st.error(err)
                    else:
                        try:
                            cid = db.crear_contrato(codigo, nombre)
                            CAL.guardar(ini, [(1, fin1)], contrato_id=cid)
                            db.asignar_usuarios(cid, acceso)
                            st.cache_data.clear()
                            st.success(f"Contrato {codigo.strip()} creado. Elígelo en la barra lateral para empezar a cargar información.")
                            st.rerun()
                        except ValueError as e:
                            st.error(str(e))
        with t_editar:
            if cts.empty:
                st.info("Aún no hay contratos.")
                return
            ids = cts["id"].astype(int).tolist()
            cid = int(st.selectbox("Contrato", ids, format_func=lambda i: cts.set_index("id").at[i, "codigo"], key="ct_sel"))
            fila = cts.set_index("id").loc[cid]
            a, b = st.columns([1, 2])
            codigo = a.text_input("Código", fila["codigo"], key=f"ct_cod_{cid}")
            nombre = b.text_input("Nombre", fila["nombre"], key=f"ct_nom_{cid}")
            activo = st.checkbox("Activo (desmarca para ocultarlo sin borrar sus datos)", value=bool(fila["activo"]), key=f"ct_act_{cid}")
            acceso = st.multiselect("Usuarios con acceso", list(nombres_u), format_func=nombres_u.get,
                                    default=[u for u in db.usuarios_de_contrato(cid) if u in nombres_u], key=f"ct_usr_{cid}")
            if st.button("Guardar cambios", type="primary", key=f"ct_guardar_{cid}"):
                try:
                    db.actualizar_contrato(cid, codigo, nombre, activo)
                    db.asignar_usuarios(cid, acceso)
                    st.cache_data.clear()
                    st.success("Contrato actualizado.")
                    st.rerun()
                except ValueError as e:
                    st.error(str(e))
    
            with st.expander("🗑️ Eliminar este contrato"):
                datos = db.resumen_datos_contrato(cid)
                st.warning(f"Se borrarán **definitivamente** el contrato **{fila['codigo']}** y toda su información: " +
                           ", ".join(f"{v} {k}" for k, v in datos.items()) +
                           ", además de su planilla del informe diario, calendario y configuración. Si solo quieres ocultarlo, "
                           "desmarca «Activo» arriba.")
                if not db.ES_POSTGRES:
                    st.caption("Antes de borrar se guarda automáticamente una copia completa de la base en el servidor.")
                conf = st.text_input(f"Para confirmar, escribe el código del contrato: {fila['codigo']}", key=f"ct_del_txt_{cid}")
                if st.button("Eliminar contrato definitivamente", key=f"ct_del_btn_{cid}",
                             disabled=conf.strip() != str(fila["codigo"]).strip()):
                    copia = db.borrar_contrato(cid)
                    if (c.contrato() or {}).get("id") == cid:
                        st.session_state.pop("contrato", None)
                        try:
                            del st.query_params["c"]
                        except KeyError:
                            pass
                    st.cache_data.clear()
                    st.cache_resource.clear()
                    st.success(f"Contrato {fila['codigo']} eliminado." + (f" Copia previa: {copia}" if copia else ""))
                    st.rerun()
    
    
    def cuenta():
        st.title("Mi cuenta")
        u = c.usuario()
        st.write(f"**Usuario:** {u['username']}  \n**Rol:** {auth.ROLES.get(u['rol'], u['rol'])}")
        _firma_propia(u)
        with st.form("cambiar_pwd", clear_on_submit=True):
            actual = st.text_input("Contraseña actual", type="password")
            nueva = st.text_input("Nueva contraseña", type="password", help=f"Mínimo {auth.LARGO_MINIMO} caracteres, con letras y números.")
            conf = st.text_input("Confirmar nueva contraseña", type="password")
            if st.form_submit_button("Cambiar contraseña", type="primary"):
                ok, err = auth.autenticar(u["username"], actual)
                if not ok:
                    st.error(err or "La contraseña actual no es correcta.")
                elif nueva != conf:
                    st.error("Las contraseñas nuevas no coinciden.")
                else:
                    try:
                        auth.cambiar_password(u["id"], nueva, conservar_sesion=st.session_state.get("token_sesion"))
                        st.success("Contraseña actualizada.")
                    except ValueError as e:
                        st.error(str(e))
    
    
    def respaldo():
        st.title("Respaldo de datos")
        if not c.exigir("administrar"):
            return
        if db.ES_POSTGRES:
            st.info("La base está en PostgreSQL: los respaldos los hace el proveedor de la base de datos.")
            return
        st.markdown("Descarga una copia completa de la base (usuarios, proyectos, informes, informe diario, Times & Material). "
                    "Sirve como respaldo y para **trasladar los datos** de un equipo a otro, por ejemplo de tu computador al servidor web.")
        if st.button("Preparar respaldo", key="resp_prep"):
            st.session_state["resp_datos"] = db.respaldo_sqlite()
        if st.session_state.get("resp_datos"):
            datos = st.session_state["resp_datos"]
            st.download_button(f"Descargar respaldo ({len(datos) / 1e6:.1f} MB)", datos, type="primary", key="resp_dl",
                               file_name=f"syncore_respaldo_{db.ahora()[:16].replace(' ', '_').replace(':', '')}.db",
                               mime="application/octet-stream")
            st.caption("Guárdalo en un lugar seguro: contiene datos del personal.")
    
        st.divider()
        st.subheader("Restaurar un respaldo")
        st.warning("Restaurar **reemplaza todos los datos actuales** por los del respaldo (incluidos los usuarios y contraseñas). "
                   "Antes se guarda automáticamente una copia de la base actual en el servidor.")
        archivo = st.file_uploader("Archivo de respaldo (.db)", type=["db"], key="resp_up")
        if archivo is None:
            return
        try:
            cuenta = db.resumen_sqlite(archivo.getvalue())
        except ValueError as e:
            st.error(str(e))
            return
        st.success("Respaldo válido: " + ", ".join(f"{v} {k.replace('_', ' ')}" for k, v in cuenta.items()))
        if st.checkbox("Confirmo que quiero reemplazar los datos actuales por este respaldo", key="resp_conf") and \
                st.button("Restaurar", type="primary", key="resp_btn"):
            previa = db.restaurar_sqlite(archivo.getvalue())
            st.cache_data.clear()
            st.cache_resource.clear()
            st.session_state.clear()
            st.success(f"Datos restaurados. La base anterior quedó guardada como {previa}. Vuelve a iniciar sesión.")
            st.rerun()
    
    
    def _firma_propia(u):
        st.subheader("Firma para informes")
        st.caption("Tu firma se usa solo cuando **tú** generas un informe diario y marcas «Firmar como Elaborado por». "
                   "Nadie más puede usarla. Es una firma electrónica simple (imagen), no una firma electrónica avanzada con certificado.")
        actual = auth.firma(u["id"])
        a, b = st.columns([1, 2])
        with a:
            if actual["png"]:
                st.image(actual["png"], caption=f"Firma actual · {str(actual['actualizada'] or '')[:16]}", use_container_width=True)
            else:
                st.info("Sin firma cargada.")
        with b:
            cargo = st.text_input("Cargo que aparece bajo tu firma", actual["cargo"], key="firma_cargo",
                                  placeholder="Ej.: Jefe de Oficina Técnica")
            archivo = st.file_uploader("Imagen de tu firma (foto o escaneo, tinta oscura sobre papel blanco)",
                                       type=["png", "jpg", "jpeg"], key="firma_up")
            nueva = None
            if archivo is not None:
                try:
                    nueva = DF.preparar_firma(archivo.getvalue())
                    st.image(nueva, caption="Así quedará (fondo transparente, recortada)", width=260)
                except Exception as e:
                    st.error(str(e))
            x, y = st.columns(2)
            if x.button("Guardar firma y cargo", type="primary", key="firma_guardar", disabled=nueva is None and cargo == actual["cargo"]):
                if nueva is not None:
                    auth.guardar_firma(u["id"], nueva, cargo)
                else:
                    auth.guardar_cargo(u["id"], cargo)
                st.success("Guardado.")
                st.rerun()
            if actual["png"] and y.button("Eliminar mi firma", key="firma_borrar"):
                auth.guardar_firma(u["id"], None)
                st.rerun()

    return SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})


admin = _modulo_admin()

# ======================================================================
# APLICACIÓN
# ======================================================================
CAL = calendario

try:   # los borradores del informe diario reponen valores de widgets: sin aviso amarillo de Streamlit por eso
    from streamlit import config as _st_config
    _st_config.set_option("global.disableWidgetStateDuplicationWarning", True)
except Exception:
    pass

# Estilos de la marca con modo día / noche (estáticos: no incluyen datos de usuarios ni de archivos).
tema.aplicar()


@st.cache_resource
def _iniciar(identidad: str):
    # La identidad de la base entra en la clave de caché: si cambia la estructura o el archivo se borra o
    # reemplaza, las tablas se vuelven a crear aunque el servidor de Streamlit no se haya reiniciado.
    db.init_db()
    auth.asegurar_admin_inicial()
    return True


_iniciar(db.identidad_bd())
LOGO = "logo_syncore.png"


def _marca(ancho_max=None):
    tema.marca(LOGO, ancho_max)


def pantalla_acceso():
    izq, col = st.columns([1.15, 1], gap="small")
    with izq:
        st.markdown(tema.panel_login(tema.logo_blanco(LOGO)), unsafe_allow_html=True)
    with col:
        st.markdown(f"<div class='sy-login-der'></div><span class='sy-version'><i></i>SYNCORE · Versión {tema.VERSION}</span>"
                    "<div class='sy-bienv'>Bienvenido</div><div class='sy-bienv-sub'>Ingresa con tu usuario para continuar.</div>",
                    unsafe_allow_html=True)
        _formularios_acceso()
        st.markdown(f"<div class='sy-login-pie'><span>© 2026 Syncore Montajes Industriales</span>"
                    f"<span>v{tema.VERSION} · {tema.FECHA_VERSION}</span></div>", unsafe_allow_html=True)


def _formularios_acceso():
    """Alta del primer administrador, inicio de sesión y solicitud de acceso."""
    if not auth.hay_usuarios():
        st.subheader("Configuración inicial")
        st.caption("No existe ningún usuario. Crea la cuenta de administrador.")
        with st.form("setup"):
            u = st.text_input("Usuario administrador")
            p = st.text_input("Contraseña", type="password")
            p2 = st.text_input("Confirmar contraseña", type="password")
            if st.form_submit_button("Crear administrador", type="primary"):
                if p != p2:
                    st.error("Las contraseñas no coinciden.")
                else:
                    try:
                        if auth.hay_usuarios():  # otra sesión se adelantó
                            st.error("El administrador ya fue creado. Recarga la página.")
                        else:
                            auth.crear_usuario(u, p, rol="admin", estado="Aprobado")
                            st.success("Administrador creado. Ya puedes ingresar.")
                            st.rerun()
                    except ValueError as e:
                        st.error(str(e))
        return

    t_login, t_reg = st.tabs(["Iniciar sesión", "Solicitar acceso"])
    with t_login:
        with st.form("login"):
            u = st.text_input("Usuario")
            p = st.text_input("Contraseña", type="password")
            recordar = st.checkbox(f"Mantener la sesión iniciada en este equipo ({auth.DIAS_RECORDAR:g} días)",
                                   help="Úsalo solo en equipos personales.")
            if st.form_submit_button("Ingresar", type="primary", use_container_width=True):
                usuario, err = auth.autenticar(u, p)
                if usuario:
                    token = auth.crear_sesion(usuario["id"], _navegador(), recordar)
                    st.session_state["usuario"] = usuario
                    st.session_state["token_sesion"] = token
                    st.query_params["s"] = token
                    st.rerun()
                else:
                    st.error(err)
    with t_reg:
        with st.form("registro", clear_on_submit=False):
            u = st.text_input("Usuario")
            n = st.text_input("Nombre completo")
            p = st.text_input("Contraseña", type="password",
                              help=f"Mínimo {auth.LARGO_MINIMO} caracteres, con letras y números.")
            p2 = st.text_input("Confirmar contraseña", type="password")
            if st.form_submit_button("Enviar solicitud"):
                if p != p2:
                    st.error("Las contraseñas no coinciden.")
                else:
                    try:
                        auth.crear_usuario(u, p, n)
                        st.success("Solicitud enviada. Un administrador debe aprobar tu cuenta.")
                    except ValueError as e:
                        st.error(str(e))


def _navegador() -> str:
    try:
        return st.context.headers.get("User-Agent", "") or ""
    except Exception:
        return ""


def _recuperar_sesion():
    """Al refrescar la página Streamlit crea una sesión nueva: se recupera el usuario desde el token de la URL."""
    if comun.usuario():
        return
    try:
        token = st.query_params.get("s")
    except Exception:
        token = None
    if not token:
        return
    usuario = auth.validar_sesion(token, _navegador())
    if usuario:
        st.session_state["usuario"] = usuario
        st.session_state["token_sesion"] = token
    else:
        try:
            del st.query_params["s"]
        except Exception:
            pass


_CLAVES_SESION = {"usuario", "token_sesion", "contrato", "_contrato_sel"}


def _selector_contrato(usuario) -> dict | None:
    """Contrato activo de la sesión. Al cambiar de contrato se limpian los filtros y selecciones del anterior."""
    cts = db.contratos_de_usuario(usuario["id"], auth.puede(usuario, "administrar"))
    if cts.empty:
        st.session_state.pop("contrato", None)
        CAL.usar_contrato(None)
        return None
    ids = cts["id"].astype(int).tolist()
    actual = (st.session_state.get("contrato") or {}).get("id")
    if actual is None:
        try:
            actual = int(st.query_params.get("c") or 0)
        except (TypeError, ValueError):
            actual = None
    fila = {int(r.id): r for r in cts.itertuples()}
    if len(ids) == 1:
        elegido = ids[0]
        st.markdown(f"**{fila[elegido].codigo}**")
    else:
        elegido = int(st.selectbox("Contrato", ids, index=ids.index(actual) if actual in ids else 0, key="_contrato_sel",
                                   format_func=lambda i: fila[i].codigo, help="Todo lo que ves y cargas queda en este contrato."))
    previo = (st.session_state.get("contrato") or {}).get("id")
    if previo is not None and previo != elegido:
        for k in [k for k in st.session_state if k not in _CLAVES_SESION]:
            del st.session_state[k]
    ct = {"id": elegido, "codigo": fila[elegido].codigo, "nombre": fila[elegido].nombre}
    st.session_state["contrato"] = ct
    CAL.usar_contrato(elegido)
    if st.query_params.get("c") != str(elegido):
        st.query_params["c"] = str(elegido)
    return ct


def main():
    _recuperar_sesion()
    usuario = comun.usuario()
    if not usuario:
        pantalla_acceso()
        return

    # Revalidar en cada carga: si un admin bloquea o cambia el rol, aplica de inmediato.
    fila = db.uno("SELECT rol, estado FROM usuarios WHERE id = ?", (usuario["id"],))
    if not fila or fila[1] != "Aprobado":
        auth.cerrar_sesion(st.session_state.get("token_sesion"))
        st.session_state.clear()
        st.query_params.clear()
        st.rerun()
    usuario["rol"] = fila[0]
    # Al cambiar de página Streamlit limpia la URL: se vuelve a poner el token para que un refresco no cierre la sesión.
    token = st.session_state.get("token_sesion")
    if token and st.query_params.get("s") != token:
        st.query_params["s"] = token

    with st.sidebar:
        _marca()
        st.markdown(tema.tarjeta_usuario(usuario["nombre"], auth.ROLES.get(usuario["rol"], usuario["rol"])), unsafe_allow_html=True)
        st.markdown("<div class='sy-etq'>CONTRATO</div>", unsafe_allow_html=True)
        contrato = _selector_contrato(usuario)
        if contrato:
            st.caption(contrato["nombre"])
        st.markdown(f"<p class='sy-pie'>{tema.AYUDA}</p>", unsafe_allow_html=True)
        if st.button("Cerrar sesión", use_container_width=True):
            auth.cerrar_sesion(st.session_state.get("token_sesion"))
            st.session_state.clear()
            st.query_params.clear()
            st.rerun()

    es_admin = auth.puede(usuario, "administrar")
    if not contrato:
        cuenta = [st.Page(admin.cuenta, title="Mi cuenta", icon="🔑", url_path="cuenta", default=not es_admin)]
        if es_admin:
            cuenta = [st.Page(admin.contratos, title="Contratos", icon="🏢", url_path="contratos", default=True),
                      st.Page(admin.usuarios, title="Usuarios", icon="🛡️", url_path="usuarios"),
                      st.Page(admin.respaldo, title="Respaldo de datos", icon="💾", url_path="respaldo")] + cuenta
            st.info("Crea el primer contrato para empezar a cargar información.")
        else:
            st.warning("Tu cuenta aún no tiene contratos asignados. Pide a un administrador que te dé acceso.")
        st.navigation({"Cuenta": cuenta}, expanded=True).run()
        return

    paginas = {
        "Análisis": [
            st.Page(dashboard.vista, title="Dashboard", icon="📊", url_path="dashboard", default=True),
            st.Page(subsistemas.vista, title="Avance por subsistema", icon="🏗️", url_path="subsistemas"),
            st.Page(informe.vista, title="Informe semanal", icon="📈", url_path="informe-semanal"),
            st.Page(vcubicacion.vista, title="Cubicación por disciplina", icon="📐", url_path="cubicacion"),
            st.Page(wla.vista, title="Lookahead 6WLA", icon="📋", url_path="6wla"),
        ],
        "Terreno": [
            st.Page(vdiario.vista, title="Informe diario", icon="🗓️", url_path="informe-diario"),
            st.Page(vtm.vista, title="Times & Material", icon="⏱️", url_path="times-material"),
            st.Page(vdesv.vista, title="Desviación de recursos", icon="⚠️", url_path="desviaciones"),
        ],
        "Datos": [
            st.Page(importar.vista, title="Importar .XER / Excel", icon="📂", url_path="importar"),
            st.Page(admin.proyectos, title="Proyectos y cortes", icon="📁", url_path="proyectos"),
        ],
        "Cuenta": [st.Page(admin.cuenta, title="Mi cuenta", icon="🔑", url_path="cuenta")],
    }
    if es_admin:
        paginas["Cuenta"].append(st.Page(admin.contratos, title="Contratos", icon="🏢", url_path="contratos"))
        paginas["Cuenta"].append(st.Page(admin.usuarios, title="Usuarios", icon="🛡️", url_path="usuarios"))
        paginas["Cuenta"].append(st.Page(admin.respaldo, title="Respaldo de datos", icon="💾", url_path="respaldo"))
    st.navigation(paginas, expanded=True).run()


main()
