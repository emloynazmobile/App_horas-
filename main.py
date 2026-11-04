# -*- coding: utf-8 -*-
"""
REGISTRO DE HORAS PRODUCTIVAS (RDP) · Kivy v10 (APK Android)
============================================================
Cambios respecto a v9:
  · FIX «Nuevo» (pantalla inicial): ahora abre el formulario de trabajo nuevo.
    La navegación lo confundía con una pantalla que necesita un trabajo ya
    existente y volvía a la lista sin abrir nada.
  · FIX «Volver»: tras guardar un trabajo o un integrante ya no regresa al
    formulario recién usado, y al volver a la lista ya no reaparecen pantallas
    antiguas (la lista es la raíz del historial; los formularios no se apilan).
  · ZOOM: el tamaño que antes se veía como 80 % es ahora el 100 % (ZOOM_BASE).
    El zoom ya guardado se convierte una sola vez para que todo se vea igual.
    Rango del zoom: 60 % – 250 %.
  · QUITAR INTEGRANTES: en la pantalla del trabajo se toca un integrante y se
    pulsa «−Integ.». El J'B no se puede quitar. Si tiene horas, se pide
    confirmación y se borran junto con él (en una sola transacción).
  · Los trabajos se eliminan desde la pantalla inicial (botón «Eliminar»).

Ruta principal: /storage/emulated/0/_zMi_Caja/_PE_1/

buildozer.spec (resumen):
  requirements = python3,kivy==2.3.0,openpyxl,plyer,android,pyjnius
  android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,READ_MEDIA_IMAGES,MANAGE_EXTERNAL_STORAGE
  android.api = 33
  android.minapi = 21
  android.archs = arm64-v8a,armeabi-v7a
"""

import os
import re
import shutil
import sqlite3
import sys
import traceback
import unicodedata
from datetime import date, datetime, timedelta

from kivy.app import App
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.graphics import Color, Rectangle, Line, Mesh
from kivy.metrics import dp, sp
from kivy.properties import NumericProperty, StringProperty
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.screenmanager import Screen, ScreenManager, SlideTransition
from kivy.uix.scrollview import ScrollView
from kivy.uix.spinner import Spinner
from kivy.uix.textinput import TextInput
from kivy.utils import platform

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except Exception:
    Workbook = load_workbook = None
    Font = PatternFill = Alignment = Border = Side = None
    get_column_letter = None
    OPENPYXL_OK = False

try:
    from plyer import filechooser
    PLYER_OK = True
except Exception:
    filechooser = None
    PLYER_OK = False

# =====================================================================
# 1. CONFIGURACIÓN Y UTILIDADES
# =====================================================================
APP_TITULO = "Registro de Horas Productivas (RDP)"
DB_NOMBRE = "base_datos_normalizada.db"

RUTA_PRINCIPAL = "/storage/emulated/0/_zMi_Caja/_PE_1/"
CARPETA_DATOS = RUTA_PRINCIPAL

RUTAS_ORIGEN_BD = [
    "/storage/emulated/0/_ zMi_Caja/_PE_1/" + DB_NOMBRE,
    "/storage/emulated/0/zMi_Caja/PE_1/" + DB_NOMBRE,
    "/storage/emulated/0/Download/" + DB_NOMBRE,
    "/storage/emulated/0/Documents/" + DB_NOMBRE,
]

MAX_HORAS = 10.0
EPS = 1e-9
# Zoom: el 100 % que se muestra equivale al tamaño que hasta v9 se veía como 80 %.
ZOOM_BASE = 0.8
ZOOM_MIN = 60
ZOOM_MAX = 250
ZOOM_PASO = 10
LETRAS_DIA = "LMMJVSD"

RUTA_LOG = None
AVISO_CARPETA = ""
RUTA_DIAGNOSTICO = None


class ErrorDatos(Exception):
    pass


def _asegurar_carpeta_principal():
    os.makedirs(RUTA_PRINCIPAL, exist_ok=True)
    if not os.access(RUTA_PRINCIPAL, os.W_OK):
        raise OSError("Sin permiso de escritura en %s" % RUTA_PRINCIPAL)
    return RUTA_PRINCIPAL


def _bd_tiene_datos(ruta_bd):
    try:
        con = sqlite3.connect(ruta_bd, timeout=3)
        try:
            f = con.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name='trabajador'"
            ).fetchone()
            if not f:
                return -1
            return con.execute("SELECT COUNT(*) FROM trabajador").fetchone()[0]
        finally:
            con.close()
    except Exception:
        return -1


def _recuperar_bd_si_falta():
    destino = os.path.join(RUTA_PRINCIPAL, DB_NOMBRE)
    n_destino = _bd_tiene_datos(destino) if os.path.isfile(destino) else -1
    if os.path.isfile(destino) and n_destino > 0:
        return (False, "")
    candidatas = []
    for origen in RUTAS_ORIGEN_BD:
        if not os.path.isfile(origen):
            continue
        if os.path.abspath(origen) == os.path.abspath(destino):
            continue
        n = _bd_tiene_datos(origen)
        try:
            tam = os.path.getsize(origen)
        except OSError:
            tam = 0
        candidatas.append((n, tam, origen))
    if not candidatas:
        return (False, "")
    candidatas.sort(key=lambda t: (t[0], t[1]), reverse=True)
    n_mejor, tam_mejor, origen_mejor = candidatas[0]
    if n_mejor > 0:
        try:
            os.makedirs(RUTA_PRINCIPAL, exist_ok=True)
            if os.path.isfile(destino):
                try:
                    shutil.copy2(destino, destino + ".antes_de_recuperar")
                except Exception:
                    pass
            shutil.copy2(origen_mejor, destino)
            return (True,
                    "BD recuperada desde %s (%d trabajadores, %d bytes)"
                    % (origen_mejor, n_mejor, tam_mejor))
        except Exception as e:
            return (False, "Error copiando BD: %s" % e)
    if os.path.isfile(origen_mejor):
        try:
            os.makedirs(RUTA_PRINCIPAL, exist_ok=True)
            if os.path.isfile(destino):
                try:
                    shutil.copy2(destino, destino + ".antes_de_recuperar")
                except Exception:
                    pass
            shutil.copy2(origen_mejor, destino)
            return (False,
                    "Copia sin datos desde %s (%d bytes). Reemplázala." % (
                        origen_mejor, tam_mejor))
        except Exception as e:
            return (False, "Error copiando BD: %s" % e)
    return (False, "")


def carpeta_base():
    global AVISO_CARPETA
    try:
        _asegurar_carpeta_principal()
    except OSError as e:
        try:
            from android.storage import app_storage_path
            alt = app_storage_path()
            os.makedirs(alt, exist_ok=True)
            AVISO_CARPETA = "No se pudo usar %s (%s). Usando %s" % (
                RUTA_PRINCIPAL, e, alt)
            return alt
        except Exception:
            pass
        raise
    recuperada, aviso = _recuperar_bd_si_falta()
    if aviso:
        AVISO_CARPETA = (AVISO_CARPETA + "\n" + aviso).strip()
    return RUTA_PRINCIPAL


def registrar_error():
    try:
        if RUTA_LOG:
            with open(RUTA_LOG, "a", encoding="utf-8") as f:
                f.write("\n[%s] (rdp kivy v10)\n" % datetime.now().isoformat(timespec="seconds"))
                f.write(traceback.format_exc())
    except Exception:
        pass


def limpiar(texto):
    return re.sub(r"\s+", " ", str(texto if texto is not None else "")).strip()


def normalizar(texto):
    s = unicodedata.normalize("NFD", str(texto if texto is not None else ""))
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s.casefold().strip()


def fmt_fecha(d):
    return "%02d/%02d/%04d" % (d.day, d.month, d.year)


def fmt_h(x):
    x = round(float(x), 2)
    return "%d" % x if x == int(x) else ("%.2f" % x).rstrip("0")


def fmt_monto(x):
    x = float(x)
    return format(x, ",.0f") if abs(x - round(x)) < 1e-9 else format(x, ",.2f")


def parsear_fecha(texto):
    t = limpiar(texto)
    for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            d = datetime.strptime(t, formato).date()
        except ValueError:
            continue
        if not 1990 <= d.year <= 2100:
            raise ErrorDatos("El año %d parece un error de tecleo." % d.year)
        return d
    raise ErrorDatos("Fecha no válida: «%s»." % t)


def parsear_horas(texto):
    t = limpiar(texto).replace(",", ".")
    if not re.match(r"^\d+(\.\d{1,2})?$", t):
        raise ErrorDatos("Las horas deben ser un número (ej. 8 o 4.5). 0 borra.")
    h = round(float(t), 2)
    if h > MAX_HORAS + EPS:
        raise ErrorDatos("Máximo %s h por día." % fmt_h(MAX_HORAS))
    return h


def parsear_monto(texto):
    t = limpiar(texto).replace(" ", "").replace(",", ".")
    if not re.match(r"^\d+(\.\d{1,2})?$", t):
        raise ErrorDatos("El Salario Creado debe ser un número.")
    return round(float(t), 2)


def periodo_valido(periodo):
    return bool(re.match(r"^\d{4}-(0[1-9]|1[0-2])$", periodo or ""))


def rango_periodo(periodo):
    if not periodo_valido(periodo):
        raise ErrorDatos("Periodo no válido. Usa AAAA-MM.")
    y, m = int(periodo[:4]), int(periodo[5:7])
    ya, ma = (y, m - 1) if m > 1 else (y - 1, 12)
    return date(ya, ma, 21), date(y, m, 20)


def periodo_de_fecha(d):
    y, m = d.year, d.month
    if d.day >= 21:
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return "%04d-%02d" % (y, m)


def periodo_desplazado(periodo, delta):
    y, m = int(periodo[:4]), int(periodo[5:7])
    idx = y * 12 + (m - 1) + delta
    return "%04d-%02d" % (idx // 12, idx % 12 + 1)


# =====================================================================
# 2. ESQUEMA
# =====================================================================
ESQUEMA_BASE = """
CREATE TABLE IF NOT EXISTS cargo (id INTEGER PRIMARY KEY, nombre TEXT NOT NULL UNIQUE);
CREATE TABLE IF NOT EXISTS cargo_aei (id INTEGER PRIMARY KEY, nombre TEXT NOT NULL UNIQUE);
CREATE TABLE IF NOT EXISTS cuadrilla (id INTEGER PRIMARY KEY, codigo TEXT NOT NULL UNIQUE);
CREATE TABLE IF NOT EXISTS trabajador (
    id INTEGER PRIMARY KEY, nombre TEXT NOT NULL,
    cargo_id INTEGER NOT NULL REFERENCES cargo(id),
    cargo_nota TEXT, cargo_aei_id INTEGER REFERENCES cargo_aei(id),
    cuadrilla_id INTEGER REFERENCES cuadrilla(id),
    albergado INTEGER NOT NULL DEFAULT 0 CHECK (albergado IN (0, 1))
);
CREATE TABLE IF NOT EXISTS mes (
    id INTEGER PRIMARY KEY CHECK (id BETWEEN 1 AND 12),
    nombre TEXT NOT NULL UNIQUE, abrev TEXT NOT NULL UNIQUE,
    dias INTEGER NOT NULL CHECK (dias BETWEEN 28 AND 31)
);
CREATE TABLE IF NOT EXISTS dia_semana (
    id INTEGER PRIMARY KEY CHECK (id BETWEEN 1 AND 7),
    nombre TEXT NOT NULL UNIQUE, abrev TEXT NOT NULL UNIQUE
);
CREATE TABLE IF NOT EXISTS incidencia (
    id INTEGER PRIMARY KEY, codigo TEXT NOT NULL UNIQUE,
    valor INTEGER NOT NULL DEFAULT 0, descripcion TEXT
);
CREATE TABLE IF NOT EXISTS registro_incidencia (
    id INTEGER PRIMARY KEY, trabajador_id INTEGER NOT NULL REFERENCES trabajador(id),
    fecha TEXT NOT NULL, incidencia_id INTEGER NOT NULL REFERENCES incidencia(id),
    UNIQUE (trabajador_id, fecha)
);
CREATE INDEX IF NOT EXISTS ix_registro_fecha ON registro_incidencia(fecha);
"""

ESQUEMA_RDP = """
CREATE TABLE IF NOT EXISTS parametro (clave TEXT PRIMARY KEY, valor TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS objeto_obra (id INTEGER PRIMARY KEY, codigo TEXT NOT NULL UNIQUE);
CREATE TABLE IF NOT EXISTS trabajo (
    id INTEGER PRIMARY KEY,
    periodo TEXT NOT NULL CHECK (periodo GLOB '[0-9][0-9][0-9][0-9]-[0-1][0-9]'
        AND substr(periodo, 6, 2) BETWEEN '01' AND '12'),
    objeto_obra_id INTEGER NOT NULL REFERENCES objeto_obra(id),
    salario_creado REAL NOT NULL CHECK (salario_creado >= 0),
    UNIQUE (periodo, objeto_obra_id)
);
CREATE TABLE IF NOT EXISTS integrante (
    trabajo_id INTEGER NOT NULL REFERENCES trabajo(id),
    trabajador_id INTEGER NOT NULL REFERENCES trabajador(id),
    es_jb INTEGER NOT NULL DEFAULT 0 CHECK (es_jb IN (0, 1)),
    orden INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY (trabajo_id, trabajador_id)
);
CREATE TABLE IF NOT EXISTS horas_trabajo (
    id INTEGER PRIMARY KEY, trabajo_id INTEGER NOT NULL, trabajador_id INTEGER NOT NULL,
    fecha TEXT NOT NULL, horas REAL NOT NULL CHECK (horas > 0 AND horas <= 10),
    UNIQUE (trabajo_id, trabajador_id, fecha),
    FOREIGN KEY (trabajo_id, trabajador_id) REFERENCES integrante (trabajo_id, trabajador_id)
);
CREATE INDEX IF NOT EXISTS ix_horas_trabajador_fecha ON horas_trabajo(trabajador_id, fecha);
CREATE TRIGGER IF NOT EXISTS trg_horas_ins_limite
BEFORE INSERT ON horas_trabajo
WHEN (SELECT COALESCE(SUM(horas), 0) FROM horas_trabajo
      WHERE trabajador_id = NEW.trabajador_id AND fecha = NEW.fecha) + NEW.horas > 10.000001
BEGIN SELECT RAISE(ABORT, 'LIMITE_10H'); END;
CREATE TRIGGER IF NOT EXISTS trg_horas_upd_limite
BEFORE UPDATE ON horas_trabajo
WHEN (SELECT COALESCE(SUM(horas), 0) FROM horas_trabajo
      WHERE trabajador_id = NEW.trabajador_id AND fecha = NEW.fecha AND id <> OLD.id)
     + NEW.horas > 10.000001
BEGIN SELECT RAISE(ABORT, 'LIMITE_10H'); END;
CREATE TRIGGER IF NOT EXISTS trg_horas_ins_periodo
BEFORE INSERT ON horas_trabajo
WHEN NOT EXISTS (SELECT 1 FROM trabajo t WHERE t.id = NEW.trabajo_id
    AND NEW.fecha BETWEEN date(t.periodo || '-01', '-1 month', '+20 days') AND t.periodo || '-20')
BEGIN SELECT RAISE(ABORT, 'FUERA_PERIODO'); END;
CREATE TRIGGER IF NOT EXISTS trg_horas_upd_periodo
BEFORE UPDATE ON horas_trabajo
WHEN NOT EXISTS (SELECT 1 FROM trabajo t WHERE t.id = NEW.trabajo_id
    AND NEW.fecha BETWEEN date(t.periodo || '-01', '-1 month', '+20 days') AND t.periodo || '-20')
BEGIN SELECT RAISE(ABORT, 'FUERA_PERIODO'); END;
CREATE TRIGGER IF NOT EXISTS trg_horas_ins_domingo
BEFORE INSERT ON horas_trabajo
WHEN strftime('%w', NEW.fecha) = '0'
BEGIN SELECT RAISE(ABORT, 'DOMINGO_NO_PERMITIDO'); END;
CREATE TRIGGER IF NOT EXISTS trg_horas_upd_domingo
BEFORE UPDATE ON horas_trabajo
WHEN strftime('%w', NEW.fecha) = '0'
BEGIN SELECT RAISE(ABORT, 'DOMINGO_NO_PERMITIDO'); END;
"""

SQL_TRABAJO = """
SELECT tr.id, tr.periodo, o.codigo, tr.salario_creado,
       (SELECT COALESCE(SUM(h.horas), 0) FROM horas_trabajo h
         WHERE h.trabajo_id = tr.id AND strftime('%w', h.fecha) <> '0') AS horas,
       (SELECT COUNT(*) FROM integrante i WHERE i.trabajo_id = tr.id) AS integrantes
FROM trabajo tr JOIN objeto_obra o ON o.id = tr.objeto_obra_id
"""


# =====================================================================
# 3. CAPA DE DATOS
# =====================================================================
class BaseDatos:
    def __init__(self, ruta):
        self.ruta = ruta
        self.aviso = ""
        self.con = sqlite3.connect(ruta, timeout=10)
        self.con.row_factory = sqlite3.Row
        self.con.execute("PRAGMA foreign_keys = ON")
        self._preparar()

    def _preparar(self):
        self.con.executescript(ESQUEMA_BASE)
        self.con.executescript(ESQUEMA_RDP)
        n_trab = self._leer("SELECT COUNT(*) FROM trabajador")[0][0]
        if n_trab == 0:
            self.aviso = ("La BD no tiene trabajadores cargados.\n"
                          "Copia tu base desde «app_trabajadores».")
        self._asegurar_incidencias()
        self._migrar_zoom()
        self._detectar_jb()
        self._limpiar_domingos()
        self.meses = {r["id"]: r["nombre"] for r in self._leer("SELECT id, nombre FROM mes")}
        self.dias = {r["id"]: r["abrev"] for r in self._leer("SELECT id, abrev FROM dia_semana")}
        if self.verificar() != "ok":
            raise sqlite3.DatabaseError("La base de datos está dañada.")

    def _asegurar_incidencias(self):
        try:
            with self.con:
                if self.get_param("incidencia_numerica") != "1":
                    for codigo, valor in [("1", 1), ("2", 2), ("3", 3), ("6", 6),
                                          ("7", 7), ("9", 9), ("10", 10)]:
                        self.con.execute(
                            "INSERT OR IGNORE INTO incidencia (codigo, valor) VALUES (?,?)",
                            (codigo, valor))
                    self.con.execute("INSERT OR REPLACE INTO parametro VALUES ('incidencia_numerica', '1')")
                if self.get_param("incidencia_i0") != "1":
                    self.con.execute(
                        "INSERT OR IGNORE INTO incidencia (codigo, valor, descripcion)"
                        " VALUES ('I-0', 0, 'Interrupto al 0%')")
                    self.con.execute("INSERT OR REPLACE INTO parametro VALUES ('incidencia_i0', '1')")
                if self.get_param("incidencia_alb_tr") != "1":
                    self.con.execute(
                        "INSERT OR IGNORE INTO incidencia (codigo, valor, descripcion)"
                        " VALUES ('ALB', 8, 'Albergamiento')")
                    self.con.execute(
                        "INSERT OR IGNORE INTO incidencia (codigo, valor, descripcion)"
                        " VALUES ('TR', 8, 'Trayecto')")
                    self.con.execute("INSERT OR REPLACE INTO parametro VALUES ('incidencia_alb_tr', '1')")
        except sqlite3.Error:
            registrar_error()

    def _detectar_jb(self):
        if self.get_param("jb_id") is not None:
            return
        candidatos = [r["id"] for r in self._leer(
            "SELECT id, cargo_nota FROM trabajador WHERE cargo_nota IS NOT NULL")
            if re.sub(r"[^a-z]", "", normalizar(r["cargo_nota"])) == "jb"]
        if len(candidatos) == 1:
            self.set_param("jb_id", str(candidatos[0]))

    def _limpiar_domingos(self):
        try:
            with self.con:
                n = self.con.execute(
                    "DELETE FROM horas_trabajo WHERE strftime('%w', fecha) = '0'").rowcount
            if n:
                print("Limpieza domingos: %d registro(s) eliminados." % n)
        except sqlite3.Error:
            registrar_error()

    def verificar(self):
        fila = self.con.execute("PRAGMA integrity_check").fetchone()
        if fila[0] != "ok":
            return fila[0]
        fallos = self.con.execute("PRAGMA foreign_key_check").fetchall()
        return "ok" if not fallos else "Hay %d referencias rotas." % len(fallos)

    def cerrar(self):
        try:
            self.con.close()
        except Exception:
            pass

    @staticmethod
    def _traducir(e, msg_unico=None, msg_fk=None):
        t = str(e).upper()
        if "LIMITE_10H" in t:
            return "Máximo 10 h/día por trabajador."
        if "FUERA_PERIODO" in t:
            return "La fecha está fuera del periodo."
        if "DOMINGO_NO_PERMITIDO" in t:
            return "Los domingos no suman a producción."
        if "UNIQUE" in t:
            return msg_unico or "Ya existe un registro con ese valor."
        if "FOREIGN KEY" in t:
            return msg_fk or "El elemento está en uso o la referencia no existe."
        if "CHECK" in t:
            return "Algún valor no es válido."
        if "NOT NULL" in t:
            return "Falta un dato obligatorio."
        return "Error de integridad: %s" % e

    def _leer(self, sql, params=()):
        return self.con.execute(sql, params).fetchall()

    def _escribir(self, sql, params=(), msg_unico=None, msg_fk=None):
        try:
            with self.con:
                return self.con.execute(sql, params)
        except sqlite3.IntegrityError as e:
            raise ErrorDatos(self._traducir(e, msg_unico, msg_fk)) from e
        except sqlite3.OperationalError as e:
            raise ErrorDatos("Error de escritura: %s" % e) from e

    def get_param(self, clave):
        f = self._leer("SELECT valor FROM parametro WHERE clave=?", (clave,))
        return f[0][0] if f else None

    def set_param(self, clave, valor):
        self._escribir("INSERT OR REPLACE INTO parametro (clave, valor) VALUES (?,?)",
                       (clave, str(valor)))

    def get_zoom(self):
        try:
            pct = int(self.get_param("zoom_pct") or "100")
        except (TypeError, ValueError):
            return 100
        return max(ZOOM_MIN, min(ZOOM_MAX, pct))

    def set_zoom(self, pct):
        pct = max(ZOOM_MIN, min(ZOOM_MAX, int(pct)))
        self.set_param("zoom_pct", str(pct))
        return pct

    def _migrar_zoom(self):
        """Hasta v9 el 100 % era el factor 1.0; ahora el 100 % equivale al 80 % de antes.
        El zoom guardado se convierte UNA sola vez (80 % -> 100 %, 100 % -> 125 %...)
        para que la pantalla se siga viendo exactamente igual."""
        try:
            if self.get_param("zoom_escala") == "2":
                return
            viejo = self.get_param("zoom_pct")
            with self.con:
                if viejo is not None:
                    try:
                        nuevo = round(int(viejo) / ZOOM_BASE)
                    except (TypeError, ValueError):
                        nuevo = 100
                    nuevo = max(ZOOM_MIN, min(ZOOM_MAX, nuevo))
                    self.con.execute("INSERT OR REPLACE INTO parametro (clave, valor) VALUES ('zoom_pct', ?)",
                                     (str(nuevo),))
                self.con.execute("INSERT OR REPLACE INTO parametro (clave, valor) VALUES ('zoom_escala', '2')")
        except sqlite3.Error:
            registrar_error()

    def jb(self):
        v = self.get_param("jb_id")
        if v is None or not v.isdigit():
            return None
        f = self._leer("SELECT id, nombre FROM trabajador WHERE id=?", (int(v),))
        return f[0] if f else None

    def establecer_jb(self, trabajador_id):
        if trabajador_id is None:
            raise ErrorDatos("Selecciona al trabajador que será el J'B.")
        if not self._leer("SELECT 1 FROM trabajador WHERE id=?", (trabajador_id,)):
            raise ErrorDatos("El trabajador no existe.")
        self.set_param("jb_id", trabajador_id)

    def listar_trabajadores(self):
        filas = self._leer("SELECT id, nombre FROM trabajador")
        return sorted(filas, key=lambda f: normalizar(f["nombre"]))

    def nombre_mes(self, periodo):
        return self.meses.get(int(periodo[5:7]), periodo[5:7])

    def abrev_dia(self, d):
        return self.dias.get(d.isoweekday(), "")

    def periodo_inicial(self):
        actual = periodo_de_fecha(date.today())
        if self._leer("SELECT 1 FROM trabajo WHERE periodo=?", (actual,)):
            return actual
        f = self._leer("SELECT MAX(periodo) FROM trabajo")
        return f[0][0] if f and f[0][0] else actual

    def _obra_id(self, codigo):
        f = self.con.execute("SELECT id FROM objeto_obra WHERE codigo=?", (codigo,)).fetchone()
        if f:
            return f[0]
        return self.con.execute("INSERT INTO objeto_obra (codigo) VALUES (?)", (codigo,)).lastrowid

    def listar_codigos_obra(self):
        return [r[0] for r in self._leer("SELECT codigo FROM objeto_obra ORDER BY codigo")]

    def listar_trabajos(self, periodo):
        rango_periodo(periodo)
        return self._leer(SQL_TRABAJO + " WHERE tr.periodo = ? ORDER BY o.codigo", (periodo,))

    def obtener_trabajo(self, trabajo_id):
        if trabajo_id is None:
            return None
        f = self._leer(SQL_TRABAJO + " WHERE tr.id = ?", (trabajo_id,))
        return f[0] if f else None

    def crear_trabajo(self, periodo, codigo, salario_txt):
        rango_periodo(periodo)
        codigo = limpiar(codigo).upper()
        if not codigo:
            raise ErrorDatos("El Objeto de Obra es obligatorio.")
        salario = parsear_monto(salario_txt)
        jb = self.jb()
        if jb is None:
            raise ErrorDatos("Define primero al J'B en Ajustes.")
        try:
            with self.con:
                obra = self._obra_id(codigo)
                tid = self.con.execute(
                    "INSERT INTO trabajo (periodo, objeto_obra_id, salario_creado) VALUES (?,?,?)",
                    (periodo, obra, salario)).lastrowid
                self.con.execute(
                    "INSERT INTO integrante (trabajo_id, trabajador_id, es_jb, orden)"
                    " VALUES (?,?,1,1)", (tid, jb["id"]))
        except sqlite3.IntegrityError as e:
            raise ErrorDatos(self._traducir(e,
                msg_unico="Ya existe un trabajo con el Objeto de Obra %s en este periodo." % codigo)) from e
        return tid

    def editar_trabajo(self, trabajo_id, codigo, salario_txt):
        codigo = limpiar(codigo).upper()
        if not codigo:
            raise ErrorDatos("El Objeto de Obra es obligatorio.")
        salario = parsear_monto(salario_txt)
        if not self.obtener_trabajo(trabajo_id):
            raise ErrorDatos("El trabajo ya no existe.")
        try:
            with self.con:
                obra = self._obra_id(codigo)
                self.con.execute("UPDATE trabajo SET objeto_obra_id=?, salario_creado=? WHERE id=?",
                                 (obra, salario, trabajo_id))
        except sqlite3.IntegrityError as e:
            raise ErrorDatos(self._traducir(e,
                msg_unico="Ya existe un trabajo con el Objeto de Obra %s en este periodo." % codigo)) from e

    def eliminar_trabajo(self, trabajo_id):
        with self.con:
            self.con.execute("DELETE FROM horas_trabajo WHERE trabajo_id=?", (trabajo_id,))
            self.con.execute("DELETE FROM integrante WHERE trabajo_id=?", (trabajo_id,))
            self.con.execute("DELETE FROM trabajo WHERE id=?", (trabajo_id,))

    def listar_integrantes(self, trabajo_id):
        return self._leer(
            "SELECT i.trabajador_id, t.nombre, i.es_jb, i.orden,"
            " (SELECT COALESCE(SUM(h.horas), 0) FROM horas_trabajo h"
            "   WHERE h.trabajo_id = i.trabajo_id AND h.trabajador_id = i.trabajador_id"
            "     AND strftime('%w', h.fecha) <> '0') AS horas,"
            " (SELECT COUNT(*) FROM horas_trabajo h"
            "   WHERE h.trabajo_id = i.trabajo_id AND h.trabajador_id = i.trabajador_id"
            "     AND strftime('%w', h.fecha) <> '0') AS dias"
            " FROM integrante i JOIN trabajador t ON t.id = i.trabajador_id"
            " WHERE i.trabajo_id = ? ORDER BY i.es_jb DESC, i.orden, t.nombre",
            (trabajo_id,))

    def candidatos_integrante(self, trabajo_id):
        filas = self._leer(
            "SELECT id, nombre FROM trabajador WHERE id NOT IN"
            " (SELECT trabajador_id FROM integrante WHERE trabajo_id = ?)", (trabajo_id,))
        return sorted(filas, key=lambda f: normalizar(f["nombre"]))

    def agregar_integrante(self, trabajo_id, trabajador_id):
        if trabajador_id is None:
            raise ErrorDatos("Selecciona un trabajador.")
        if not self.obtener_trabajo(trabajo_id):
            raise ErrorDatos("El trabajo ya no existe.")
        orden = self._leer("SELECT COALESCE(MAX(orden), 0) + 1 FROM integrante WHERE trabajo_id=?",
                           (trabajo_id,))[0][0]
        self._escribir(
            "INSERT INTO integrante (trabajo_id, trabajador_id, es_jb, orden) VALUES (?,?,0,?)",
            (trabajo_id, trabajador_id, orden),
            msg_unico="Ese trabajador ya forma parte de la cuadrilla.",
            msg_fk="El trabajador no existe.")

    def quitar_integrante(self, trabajo_id, trabajador_id, borrar_horas=False):
        """Quita a un trabajador de la cuadrilla de UN trabajo (no lo borra del catálogo).
        El J'B no se puede quitar. Si tiene horas en este trabajo: con borrar_horas=False
        se rechaza; con borrar_horas=True se borran junto con el integrante, todo o nada."""
        f = self._leer("SELECT es_jb FROM integrante WHERE trabajo_id=? AND trabajador_id=?",
                       (trabajo_id, trabajador_id))
        if not f:
            raise ErrorDatos("El integrante ya no está en la cuadrilla.")
        if f[0]["es_jb"]:
            raise ErrorDatos("El J'B debe permanecer en la cuadrilla.")
        n = self._leer("SELECT COUNT(*) FROM horas_trabajo WHERE trabajo_id=? AND trabajador_id=?",
                       (trabajo_id, trabajador_id))[0][0]
        if n and not borrar_horas:
            raise ErrorDatos("Tiene horas registradas en %d día(s). Bórralas primero." % n)
        try:
            with self.con:
                self.con.execute("DELETE FROM horas_trabajo WHERE trabajo_id=? AND trabajador_id=?",
                                 (trabajo_id, trabajador_id))
                self.con.execute("DELETE FROM integrante WHERE trabajo_id=? AND trabajador_id=?",
                                 (trabajo_id, trabajador_id))
        except sqlite3.Error as e:
            raise ErrorDatos(self._traducir(e)) from e

    def _max_produccion_dia(self, trabajador_id, iso):
        inc = self._leer(
            "SELECT i.valor, i.codigo FROM registro_incidencia r"
            " JOIN incidencia i ON i.id = r.incidencia_id"
            " WHERE r.trabajador_id=? AND r.fecha=?", (trabajador_id, iso))
        if inc:
            return inc[0][0], inc[0][1]
        return MAX_HORAS, None

    def guardar_horas(self, trabajo_id, objetivo, desde_txt, hasta_txt, horas_txt, omitir_finde):
        tr = self.obtener_trabajo(trabajo_id)
        if not tr:
            raise ErrorDatos("El trabajo ya no existe.")
        ini, fin = rango_periodo(tr["periodo"])
        desde = parsear_fecha(desde_txt)
        hasta = parsear_fecha(hasta_txt) if limpiar(hasta_txt) else desde
        if hasta < desde:
            raise ErrorDatos("La fecha final no puede ser anterior a la inicial.")
        if desde < ini or hasta > fin:
            raise ErrorDatos("Las fechas deben estar dentro del periodo.")
        horas = parsear_horas(horas_txt)
        fechas = [desde + timedelta(days=k) for k in range((hasta - desde).days + 1)]
        fechas = [d for d in fechas if d.isoweekday() != 7]
        if omitir_finde:
            fechas = [d for d in fechas if d.isoweekday() != 6]
        if not fechas:
            raise ErrorDatos("El rango solo contiene domingos o sábados omitidos.")
        integrantes = {r["trabajador_id"]: r for r in self.listar_integrantes(trabajo_id)}
        if objetivo is None:
            objetivos = [i for i, r in integrantes.items() if not r["es_jb"]]
            if not objetivos:
                raise ErrorDatos("La cuadrilla no tiene integrantes además del J'B.")
        else:
            if objetivo not in integrantes:
                raise ErrorDatos("Ese trabajador no forma parte de la cuadrilla.")
            objetivos = [objetivo]
        try:
            with self.con:
                for w in objetivos:
                    for d in fechas:
                        iso = d.isoformat()
                        if horas == 0:
                            self.con.execute(
                                "DELETE FROM horas_trabajo WHERE trabajo_id=? AND trabajador_id=? AND fecha=?",
                                (trabajo_id, w, iso))
                            continue
                        max_prod, codigo_inc = self._max_produccion_dia(w, iso)
                        otras = self.con.execute(
                            "SELECT COALESCE(SUM(horas), 0) FROM horas_trabajo"
                            " WHERE trabajador_id=? AND fecha=? AND trabajo_id<>?"
                            "   AND strftime('%w', fecha) <> '0'",
                            (w, iso, trabajo_id)).fetchone()[0]
                        if otras + horas > max_prod + EPS:
                            if codigo_inc and max_prod == 0:
                                raise ErrorDatos(
                                    "%s el %s: incidencia «%s» con 0 h." % (
                                        integrantes[w]["nombre"], fmt_fecha(d), codigo_inc))
                            disponible = max(0, max_prod - otras)
                            raise ErrorDatos(
                                "%s el %s: máx %s h (inc. «%s»), ya %s h. "
                                "Hasta %s h." % (
                                    integrantes[w]["nombre"], fmt_fecha(d), fmt_h(max_prod),
                                    codigo_inc or "-", fmt_h(otras), fmt_h(disponible)))
                        cur = self.con.execute(
                            "UPDATE horas_trabajo SET horas=? WHERE trabajo_id=? AND trabajador_id=? AND fecha=?",
                            (horas, trabajo_id, w, iso))
                        if cur.rowcount == 0:
                            self.con.execute(
                                "INSERT INTO horas_trabajo (trabajo_id, trabajador_id, fecha, horas)"
                                " VALUES (?,?,?,?)", (trabajo_id, w, iso, horas))
        except sqlite3.IntegrityError as e:
            raise ErrorDatos(self._traducir(e)) from e
        return len(objetivos) * len(fechas)

    def matriz_trabajo(self, trabajo_id):
        tr = self.obtener_trabajo(trabajo_id)
        if not tr:
            raise ErrorDatos("El trabajo ya no existe.")
        ini, fin = rango_periodo(tr["periodo"])
        integrantes = self.listar_integrantes(trabajo_id)
        dias = [ini + timedelta(days=k) for k in range((fin - ini).days + 1)]
        horas = {(r["trabajador_id"], r["fecha"]): r["horas"] for r in self._leer(
            "SELECT trabajador_id, fecha, horas FROM horas_trabajo"
            " WHERE trabajo_id=? AND strftime('%w', fecha) <> '0'", (trabajo_id,))}
        incidencias = {}
        ids = [i["trabajador_id"] for i in integrantes]
        if ids:
            marcas = ",".join("?" * len(ids))
            for r in self._leer(
                    "SELECT r.trabajador_id, r.fecha, i.codigo, i.valor FROM registro_incidencia r"
                    " JOIN incidencia i ON i.id = r.incidencia_id"
                    " WHERE r.trabajador_id IN (%s) AND r.fecha BETWEEN ? AND ?" % marcas,
                    ids + [ini.isoformat(), fin.isoformat()]):
                incidencias[(r["trabajador_id"], r["fecha"])] = (r["codigo"], r["valor"])
        return {"trabajo": tr, "integrantes": integrantes, "dias": dias,
                "horas": horas, "incidencias": incidencias}

    def guardar_hora_dia(self, trabajo_id, trabajador_id, fecha, horas):
        tr = self.obtener_trabajo(trabajo_id)
        if not tr:
            raise ErrorDatos("El trabajo ya no existe.")
        d = fecha if isinstance(fecha, date) else parsear_fecha(fecha)
        if d.isoweekday() == 7:
            raise ErrorDatos("Los domingos no suman a producción.")
        ini, fin = rango_periodo(tr["periodo"])
        if not ini <= d <= fin:
            raise ErrorDatos("La fecha está fuera del periodo.")
        if not self._leer("SELECT 1 FROM integrante WHERE trabajo_id=? AND trabajador_id=?",
                          (trabajo_id, trabajador_id)):
            raise ErrorDatos("Ese trabajador no forma parte de la cuadrilla.")
        if horas is not None and horas < 0:
            raise ErrorDatos("Las horas no pueden ser negativas.")
        if horas is not None and horas > MAX_HORAS + EPS:
            raise ErrorDatos("Máximo %s h por día." % fmt_h(MAX_HORAS))
        iso = d.isoformat()
        max_prod, codigo_inc = self._max_produccion_dia(trabajador_id, iso)
        try:
            with self.con:
                if not horas or horas <= 0:
                    self.con.execute(
                        "DELETE FROM horas_trabajo WHERE trabajo_id=? AND trabajador_id=? AND fecha=?",
                        (trabajo_id, trabajador_id, iso))
                else:
                    if max_prod <= 0:
                        raise ErrorDatos(
                            "Incidencia «%s»: 0 h trabajadas. No se puede producir." % (
                                codigo_inc or "-"))
                    otras = self.con.execute(
                        "SELECT COALESCE(SUM(horas), 0) FROM horas_trabajo"
                        " WHERE trabajador_id=? AND fecha=? AND trabajo_id<>?"
                        "   AND strftime('%w', fecha) <> '0'",
                        (trabajador_id, iso, trabajo_id)).fetchone()[0]
                    if otras + horas > max_prod + EPS:
                        disponible = max(0, max_prod - otras)
                        raise ErrorDatos(
                            "Máx %s h (inc. «%s»), ya %s h. Puede hasta %s h." % (
                                fmt_h(max_prod), codigo_inc or "-",
                                fmt_h(otras), fmt_h(disponible)))
                    cur = self.con.execute(
                        "UPDATE horas_trabajo SET horas=? WHERE trabajo_id=? AND trabajador_id=? AND fecha=?",
                        (horas, trabajo_id, trabajador_id, iso))
                    if cur.rowcount == 0:
                        self.con.execute(
                            "INSERT INTO horas_trabajo (trabajo_id, trabajador_id, fecha, horas)"
                            " VALUES (?,?,?,?)", (trabajo_id, trabajador_id, iso, horas))
        except sqlite3.IntegrityError as e:
            raise ErrorDatos(self._traducir(e)) from e

    def resumen_periodo(self, periodo):
        rango_periodo(periodo)
        filas = self._leer(
            "SELECT t.id, t.nombre, COALESCE(SUM(h.horas), 0) AS horas,"
            " COUNT(DISTINCT h.fecha) AS dias, COUNT(DISTINCT i.trabajo_id) AS trabajos"
            " FROM integrante i"
            " JOIN trabajo tr ON tr.id = i.trabajo_id AND tr.periodo = ?"
            " JOIN trabajador t ON t.id = i.trabajador_id"
            " LEFT JOIN horas_trabajo h ON h.trabajo_id = i.trabajo_id"
            "     AND h.trabajador_id = i.trabajador_id AND strftime('%w', h.fecha) <> '0'"
            " GROUP BY t.id", (periodo,))
        return sorted(filas, key=lambda f: normalizar(f["nombre"]))

    def detalle_trabajador(self, periodo, trabajador_id):
        return self._leer(
            "SELECT o.codigo, COALESCE(SUM(h.horas), 0) AS horas"
            " FROM integrante i"
            " JOIN trabajo tr ON tr.id = i.trabajo_id AND tr.periodo = ?"
            " JOIN objeto_obra o ON o.id = tr.objeto_obra_id"
            " LEFT JOIN horas_trabajo h ON h.trabajo_id = i.trabajo_id"
            "     AND h.trabajador_id = i.trabajador_id AND strftime('%w', h.fecha) <> '0'"
            " WHERE i.trabajador_id = ? GROUP BY o.codigo ORDER BY o.codigo",
            (periodo, trabajador_id))

    def totales_periodo(self, periodo):
        f = self._leer(
            "SELECT COALESCE(SUM(tr.salario_creado), 0),"
            " (SELECT COALESCE(SUM(h.horas), 0) FROM horas_trabajo h"
            "  JOIN trabajo x ON x.id = h.trabajo_id"
            "  WHERE x.periodo = ? AND strftime('%w', h.fecha) <> '0')"
            " FROM trabajo tr WHERE tr.periodo = ?", (periodo, periodo))
        return f[0][0], f[0][1]

    def estadisticas(self):
        res = {}
        for t in ("trabajador", "objeto_obra", "trabajo", "integrante", "horas_trabajo"):
            res[t] = self._leer("SELECT COUNT(*) FROM %s" % t)[0][0]
        return res

    def _subcarpeta(self, nombre):
        carpeta = os.path.join(os.path.dirname(os.path.abspath(self.ruta)), nombre)
        os.makedirs(carpeta, exist_ok=True)
        return carpeta

    def respaldar(self):
        destino = os.path.join(self._subcarpeta("respaldos"),
                               "respaldo_%s.db" % datetime.now().strftime("%Y%m%d_%H%M%S"))
        copia = sqlite3.connect(destino)
        try:
            self.con.backup(copia)
        finally:
            copia.close()
        return destino

    def exportar_periodo_xlsx(self, periodo):
        if not OPENPYXL_OK:
            raise ErrorDatos("Falta openpyxl. Instálalo con: pip install openpyxl")
        ini, fin = rango_periodo(periodo)
        trabajos = self.listar_trabajos(periodo)
        if not trabajos:
            raise ErrorDatos("No hay trabajos en este periodo.")
        dias = [ini + timedelta(days=k) for k in range((fin - ini).days + 1)]
        wb = Workbook()
        ws = wb.active
        ws.title = "RDP %s" % periodo
        encabezado_fill = PatternFill("solid", fgColor="DCE6F1")
        titulo_fill = PatternFill("solid", fgColor="FFE08A")
        borde = Border(*(Side(border_style="thin", color="7F7F7F"),) * 4)
        center = Alignment(horizontal="center", vertical="center")
        r = 1
        ws.cell(row=r, column=1, value="RDP · %s %s · %s al %s" % (
            self.nombre_mes(periodo), periodo[:4], fmt_fecha(ini), fmt_fecha(fin))
        ).font = Font(bold=True, size=12)
        ws.cell(row=r, column=1).fill = titulo_fill
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4 + len(dias))
        r += 2
        for tr in trabajos:
            ws.cell(row=r, column=1, value="Objeto de Obra").font = Font(bold=True)
            ws.cell(row=r, column=2, value=tr["codigo"]).font = Font(bold=True)
            ws.cell(row=r, column=3, value="Salario Creado").font = Font(bold=True)
            ws.cell(row=r, column=4, value=fmt_monto(tr["salario_creado"]))
            for c in range(1, 5):
                ws.cell(row=r, column=c).fill = encabezado_fill
                ws.cell(row=r, column=c).border = borde
            r += 1
            ws.cell(row=r, column=1, value="Nro").font = Font(bold=True)
            ws.cell(row=r, column=2, value="Nombre").font = Font(bold=True)
            ws.cell(row=r, column=3, value="ID").font = Font(bold=True)
            ws.cell(row=r, column=4, value="Total").font = Font(bold=True)
            for i, d in enumerate(dias):
                c = ws.cell(row=r, column=5 + i, value=LETRAS_DIA[d.isoweekday() - 1])
                c.font = Font(bold=True)
                c.alignment = center
            for c in range(1, 5 + len(dias)):
                ws.cell(row=r, column=c).fill = encabezado_fill
                ws.cell(row=r, column=c).border = borde
            r += 1
            for i, d in enumerate(dias):
                c = ws.cell(row=r, column=5 + i, value=d.day)
                c.alignment = center
                c.border = borde
            r += 1
            horas = {(x["trabajador_id"], x["fecha"]): x["horas"] for x in self._leer(
                "SELECT trabajador_id, fecha, horas FROM horas_trabajo"
                " WHERE trabajo_id=? AND strftime('%w', fecha) <> '0'", (tr["id"],))}
            for n, i in enumerate(self.listar_integrantes(tr["id"]), start=1):
                ws.cell(row=r, column=1, value=n).alignment = center
                ws.cell(row=r, column=2, value=i["nombre"])
                ws.cell(row=r, column=3, value=i["trabajador_id"]).alignment = center
                ws.cell(row=r, column=4, value=fmt_h(i["horas"])).alignment = center
                for k, d in enumerate(dias):
                    clave = (i["trabajador_id"], d.isoformat())
                    if clave in horas:
                        ws.cell(row=r, column=5 + k, value=horas[clave]).alignment = center
                for c in range(1, 5 + len(dias)):
                    ws.cell(row=r, column=c).border = borde
                r += 1
            ws.cell(row=r, column=3, value="TOTAL").font = Font(bold=True)
            ws.cell(row=r, column=4, value=fmt_h(tr["horas"])).font = Font(bold=True)
            for c in range(1, 5 + len(dias)):
                ws.cell(row=r, column=c).fill = encabezado_fill
                ws.cell(row=r, column=c).border = borde
            r += 2
        ws.column_dimensions["A"].width = 5
        ws.column_dimensions["B"].width = 30
        ws.column_dimensions["C"].width = 8
        ws.column_dimensions["D"].width = 12
        for i in range(len(dias)):
            ws.column_dimensions[get_column_letter(5 + i)].width = 5
        destino = os.path.join(self._subcarpeta("exportaciones"),
                               "horas_rdp_%s_%s.xlsx" % (
                                   periodo, datetime.now().strftime("%Y%m%d_%H%M%S")))
        wb.save(destino)
        return destino

    def importar_periodo_xlsx(self, ruta, periodo):
        if not OPENPYXL_OK:
            raise ErrorDatos("Falta openpyxl.")
        if not periodo_valido(periodo):
            raise ErrorDatos("Periodo no válido.")
        if not os.path.isfile(ruta):
            raise ErrorDatos("El archivo no existe: %s" % ruta)
        try:
            wb = load_workbook(ruta, data_only=True)
        except Exception as e:
            raise ErrorDatos("No se pudo abrir: %s" % e)
        ws = wb.active
        ini, fin = rango_periodo(periodo)
        dias = [ini + timedelta(days=k) for k in range((fin - ini).days + 1)]
        n_dias = len(dias)
        filas = list(ws.iter_rows(values_only=True))
        idx = 0
        trabajos_importados = []
        horas_importadas = 0
        integrantes_nuevos = 0
        errores_por_fila = []
        ids_validos = {r["id"] for r in self._leer("SELECT id FROM trabajador")}
        while idx < len(filas):
            fila = filas[idx]
            es_cabecera = any(isinstance(v, str) and "objeto de obra" in v.lower()
                              for v in fila if v is not None)
            if not es_cabecera:
                idx += 1
                continue
            codigo = None
            salario = 0
            celdas = list(fila)
            for k, v in enumerate(celdas):
                if isinstance(v, str) and "objeto de obra" in v.lower():
                    for j in range(k + 1, len(celdas)):
                        if isinstance(celdas[j], str) and celdas[j].strip():
                            codigo = celdas[j].strip().upper()
                            break
                if isinstance(v, str) and "salario creado" in v.lower():
                    for j in range(k + 1, len(celdas)):
                        if isinstance(celdas[j], (int, float)):
                            salario = float(celdas[j])
                            break
            if not codigo:
                idx += 1
                continue
            idx += 2
            if idx >= len(filas):
                break
            fila_numeros = filas[idx]
            col_dias = []
            for k in range(4, min(len(fila_numeros), 4 + n_dias + 2)):
                v = fila_numeros[k]
                if isinstance(v, (int, float)) and 1 <= int(v) <= 31:
                    col_dias.append((k, int(v)))
            idx += 1
            try:
                existente = self._leer(
                    "SELECT tr.id, tr.salario_creado FROM trabajo tr"
                    " JOIN objeto_obra o ON o.id = tr.objeto_obra_id"
                    " WHERE tr.periodo=? AND o.codigo=?", (periodo, codigo))
                if existente:
                    tid = existente[0]["id"]
                    if salario and abs(salario - existente[0]["salario_creado"]) > 1e-6:
                        self._escribir("UPDATE trabajo SET salario_creado=? WHERE id=?",
                                       (salario, tid))
                else:
                    tid = self.crear_trabajo(periodo, codigo, str(salario))
                trabajos_importados.append(codigo)
            except ErrorDatos as e:
                errores_por_fila.append("Cabecera %s: %s" % (codigo, e))
                continue
            integrantes_actuales = {r["trabajador_id"] for r in self.listar_integrantes(tid)}
            while idx < len(filas):
                fila_int = filas[idx]
                if not fila_int:
                    idx += 1
                    continue
                if isinstance(fila_int[3] if len(fila_int) > 3 else "",
                              str) and str(fila_int[3]).upper() == "TOTAL":
                    idx += 1
                    break
                if len(fila_int) < 3:
                    idx += 1
                    continue
                try:
                    id_int = int(fila_int[2])
                except (TypeError, ValueError):
                    idx += 1
                    continue
                if id_int not in ids_validos:
                    idx += 1
                    continue
                if id_int not in integrantes_actuales:
                    try:
                        self.agregar_integrante(tid, id_int)
                        integrantes_actuales.add(id_int)
                        integrantes_nuevos += 1
                    except ErrorDatos:
                        idx += 1
                        continue
                for (col_k, dia_num) in col_dias:
                    if col_k >= len(fila_int):
                        continue
                    val = fila_int[col_k]
                    if val is None or val == "":
                        continue
                    try:
                        horas = float(val)
                    except (TypeError, ValueError):
                        continue
                    if horas <= 0:
                        continue
                    fecha = None
                    for d in dias:
                        if d.day == dia_num:
                            fecha = d
                            break
                    if fecha is None:
                        continue
                    try:
                        self.guardar_hora_dia(tid, id_int, fecha, horas)
                        horas_importadas += 1
                    except ErrorDatos as e:
                        errores_por_fila.append("Fila %d: %s" % (idx + 1, e))
                idx += 1
        return {"trabajos": trabajos_importados,
                "integrantes_nuevos": integrantes_nuevos,
                "horas": horas_importadas, "errores": errores_por_fila}


# =====================================================================
# 4. WIDGETS Y DIÁLOGOS
# =====================================================================
def aviso(mensaje, titulo="Aviso"):
    box = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(8))
    lbl = Label(text=str(mensaje), halign="left", valign="top",
                text_size=(Window.width * 0.8, None))
    box.add_widget(lbl)
    b = Button(text="Cerrar", size_hint_y=None, height=dp(44))
    box.add_widget(b)
    popup = Popup(title=titulo, content=box, size_hint=(0.9, 0.5))
    b.bind(on_release=lambda *_: popup.dismiss())
    popup.open()


def confirmar(mensaje, al_confirmar, titulo="Confirmar"):
    box = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(8))
    box.add_widget(Label(text=str(mensaje), halign="left", valign="top",
                          text_size=(Window.width * 0.8, None)))
    fila = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(4))
    b_si = Button(text="Sí")
    b_no = Button(text="No")
    fila.add_widget(b_si)
    fila.add_widget(b_no)
    box.add_widget(fila)
    popup = Popup(title=titulo, content=box, size_hint=(0.9, 0.4))

    def _si(*_):
        popup.dismiss()
        try:
            al_confirmar()
        except Exception:
            registrar_error()

    b_si.bind(on_release=_si)
    b_no.bind(on_release=lambda *_: popup.dismiss())
    popup.open()


def popup_texto_copiable(texto, titulo="Texto"):
    box = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(8))
    ti = TextInput(text=texto, readonly=True, multiline=True,
                   font_size=sp(11), background_color=(1, 1, 1, 1),
                   foreground_color=(0, 0, 0, 1))
    box.add_widget(ti)
    fila = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(4))
    b_cerrar = Button(text="Cerrar")
    fila.add_widget(b_cerrar)
    box.add_widget(fila)
    popup = Popup(title=titulo, content=box, size_hint=(0.95, 0.9))
    b_cerrar.bind(on_release=lambda *_: popup.dismiss())
    popup.open()


class BotonTriangulo(Button):
    def __init__(self, direccion="izq", **kw):
        super().__init__(**kw)
        self.direccion = direccion
        self.background_normal = ""
        self.background_color = (0.25, 0.45, 0.65, 1)
        self.size_hint_x = None
        self.bind(pos=self._redibujar, size=self._redibujar)

    def _redibujar(self, *_):
        self.canvas.after.clear()
        with self.canvas.after:
            Color(1, 1, 1, 1)
            cx, cy = self.center_x, self.center_y
            r = min(self.width, self.height) * 0.25
            if self.direccion == "izq":
                pts = [cx + r * 0.6, cy + r, cx + r * 0.6, cy - r, cx - r * 0.6, cy]
            else:
                pts = [cx - r * 0.6, cy + r, cx - r * 0.6, cy - r, cx + r * 0.6, cy]
            Line(points=pts, width=dp(1.2), close=True)
            vertices = []
            for i in range(0, len(pts), 2):
                vertices.extend([pts[i], pts[i + 1], 0, 0])
            Mesh(vertices=vertices, indices=[0, 1, 2], mode="triangle_fan")


class FilaTrabajo(BoxLayout):
    def __init__(self, app, fila, pantalla, **kw):
        super().__init__(orientation="horizontal", size_hint_y=None,
                         height=app.escala(dp(44)), spacing=dp(2),
                         padding=[dp(4), 0], **kw)
        self.app = app
        self.fila = fila
        self.pantalla = pantalla
        self.trabajo_id = fila["id"]
        self.color_normal = (0.95, 0.95, 0.95, 1)
        self.color_sel = (0.70, 0.85, 1.0, 1)
        with self.canvas.before:
            self._color = Color(*self.color_normal)
            self._rect = Rectangle(pos=self.pos, size=self.size)
        self.bind(pos=self._actualizar_rect, size=self._actualizar_rect)
        esc = app.escala
        self.add_widget(self._lbl(str(fila["codigo"]), 3, esc))
        self.add_widget(self._lbl(fmt_monto(fila["salario_creado"]), 2, esc))
        self.add_widget(self._lbl(fmt_h(fila["horas"]), 1, esc))
        self.add_widget(self._lbl(str(fila["integrantes"]), 1, esc))
        self.bind(on_touch_down=self._tocar)

    def _lbl(self, texto, size_x, esc):
        lbl = Label(text=texto, size_hint_x=size_x, color=(0, 0, 0, 1),
                    halign="left", valign="middle", font_size=esc(sp(13)))
        lbl.bind(size=lambda inst, v: setattr(inst, "text_size", (v[0], v[1])))
        return lbl

    def _actualizar_rect(self, *_):
        self._rect.pos = self.pos
        self._rect.size = self.size

    def _tocar(self, inst, touch):
        if self.collide_point(*touch.pos):
            self.pantalla._seleccionar(self.trabajo_id)
            if touch.is_double_tap:
                self.pantalla._abrir()
            return True
        return False

    def marcar(self, seleccionado):
        self._color.rgba = self.color_sel if seleccionado else self.color_normal


class FilaIntegrante(BoxLayout):
    """Fila tocable de la cuadrilla (nombre · rol · producción) en la pantalla del trabajo."""

    def __init__(self, app, fila, pantalla, **kw):
        super().__init__(orientation="horizontal", size_hint_y=None,
                         height=app.escala(dp(36)), spacing=dp(2),
                         padding=[dp(4), 0], **kw)
        self.app = app
        self.pantalla = pantalla
        self.trabajador_id = fila["trabajador_id"]
        self.color_normal = (0.95, 0.95, 0.95, 1)
        self.color_sel = (0.70, 0.85, 1.0, 1)
        with self.canvas.before:
            self._color = Color(*self.color_normal)
            self._rect = Rectangle(pos=self.pos, size=self.size)
        self.bind(pos=self._actualizar_rect, size=self._actualizar_rect)
        esc = app.escala
        self.add_widget(self._lbl(str(fila["nombre"]), 4, esc))
        self.add_widget(self._lbl("J'B" if fila["es_jb"] else "", 1, esc))
        self.add_widget(self._lbl(fmt_h(fila["horas"]), 1, esc))
        self.bind(on_touch_down=self._tocar)

    def _lbl(self, texto, size_x, esc):
        lbl = Label(text=texto, size_hint_x=size_x, color=(0, 0, 0, 1),
                    halign="left", valign="middle", font_size=esc(sp(13)))
        lbl.bind(size=lambda inst, v: setattr(inst, "text_size", (v[0], v[1])))
        return lbl

    def _actualizar_rect(self, *_):
        self._rect.pos = self.pos
        self._rect.size = self.size

    def _tocar(self, inst, touch):
        if self.collide_point(*touch.pos):
            self.pantalla._seleccionar_integrante(self.trabajador_id)
            return True
        return False

    def marcar(self, seleccionado):
        self._color.rgba = self.color_sel if seleccionado else self.color_normal


class CeldaDia(TextInput):
    def __init__(self, clave, valor_inicial, estado, app, **kw):
        super().__init__(**kw)
        self.clave = clave
        self.previo = valor_inicial or ""
        self.estado = estado
        self.app = app
        self.text = valor_inicial or ""
        self.multiline = False
        self.input_type = "number"
        self.halign = "center"
        self.padding = [dp(2), dp(6), dp(2), dp(6)]
        self.background_normal = ""
        self.background_active = ""
        self.background_color = self._color_fondo()
        self.foreground_color = (0, 0, 0, 1)
        self.cursor_color = (0, 0, 0, 1)
        self.readonly = bool(self.estado.get("bloqueada", False))
        self.font_size = app.escala(sp(14))
        self.bind(on_text_validate=self._confirmar)
        self.bind(focus=self._al_cambiar_foco)
        self.bind(on_text=self._al_teclear)

    def _color_fondo(self):
        return self.estado.get("fondo", (1, 1, 1, 1))

    def _al_teclear(self, inst, texto):
        if self.readonly:
            self.text = self.previo
            return
        limpio = "".join(ch for ch in texto if ch.isdigit() or ch in ".,")
        if len(limpio) > 5 or limpio.count(",") + limpio.count(".") > 1:
            self.text = self.text[:-1] if self.text else ""
        elif limpio != texto:
            self.text = limpio

    def _al_cambiar_foco(self, inst, tiene):
        if not tiene and not self.readonly:
            self._confirmar(None)

    def _confirmar(self, *_):
        if self.readonly:
            return
        texto = self.text.strip()
        if texto == self.previo:
            return
        try:
            self.app.bd.guardar_hora_dia(
                self.app.trabajo_id_actual, self.clave[0], self.clave[1],
                0 if not texto else round(float(texto.replace(",", ".")), 2))
            self.previo = texto
            self.app.actualizar_totales_rdp()
        except ErrorDatos as e:
            self.text = self.previo
            aviso(str(e))
        except Exception as e:
            registrar_error()
            self.text = self.previo
            aviso("Error: %s" % e)


class CalendarioRDP(GridLayout):
    def __init__(self, app, trabajo_id, wid, **kw):
        super().__init__(cols=7, size_hint_y=None, spacing=dp(1), **kw)
        self.app = app
        self.bd = app.bd
        self.trabajo_id = trabajo_id
        self.wid = wid
        self.celdas = {}
        self.horas = {}
        self.incidencias = {}
        self._construir()

    def _construir(self):
        for w in self.children[:]:
            self.remove_widget(w)
        datos = self.bd.matriz_trabajo(self.trabajo_id)
        self.horas = dict(datos["horas"])
        self.incidencias = dict(datos["incidencias"])
        ini, fin = rango_periodo(datos["trabajo"]["periodo"])
        isos = [d.isoformat() for d in datos["dias"]]
        lunes = ini - timedelta(days=ini.weekday())
        semanas = (fin - lunes).days // 7 + 1
        hoy = date.today().isoformat()
        iso_hoy = hoy if isos and isos[0] <= hoy <= isos[-1] else None
        esc = self.app.escala

        for letra in LETRAS_DIA:
            lbl = Label(text=letra, bold=True, size_hint_y=None,
                        height=esc(dp(24)), color=(0, 0, 0, 1),
                        font_size=esc(sp(12)))
            with lbl.canvas.before:
                Color(0.86, 0.90, 0.95, 1)
                Rectangle(pos=lbl.pos, size=lbl.size)
            lbl.bind(pos=self._redibujar_fondo, size=self._redibujar_fondo)
            self.add_widget(lbl)

        for s in range(semanas):
            for c in range(7):
                d = lunes + timedelta(days=7 * s + c)
                self.add_widget(self._crear_celda(d, ini, fin, iso_hoy, esc))

        self.height = esc(dp(24)) * (semanas + 1) + esc(dp(6))

    @staticmethod
    def _redibujar_fondo(inst, *_):
        inst.canvas.before.clear()
        with inst.canvas.before:
            Color(0.86, 0.90, 0.95, 1)
            Rectangle(pos=inst.pos, size=inst.size)

    def _crear_celda(self, d, ini, fin, iso_hoy, esc):
        dentro = ini <= d <= fin
        iso = d.isoformat()
        es_domingo = d.isoweekday() == 7
        clave = (self.wid, iso) if dentro else None
        info_inc = self.incidencias.get(clave) if clave else None
        bloquea_inc = info_inc is not None and info_inc[1] <= 0
        bloqueada = (not dentro) or es_domingo or bloquea_inc

        if not dentro:
            fondo = (0.75, 0.75, 0.75, 1)
        elif iso == iso_hoy:
            fondo = (1.0, 0.95, 0.77, 1)
        elif es_domingo or bloquea_inc:
            fondo = (0.85, 0.85, 0.85, 1)
        else:
            tramo_inicial = d < date(int(self.app.periodo[:4]),
                                     int(self.app.periodo[5:7]), 1)
            fondo = (0.886, 0.937, 0.855, 1) if tramo_inicial else (0.886, 0.937, 0.667, 1)

        box = BoxLayout(orientation="vertical", size_hint_y=None,
                        height=esc(dp(58)), spacing=esc(dp(1)))
        lbl_fecha = Label(text=d.strftime("%d/%m"), size_hint_y=None,
                          height=esc(dp(16)), halign="left", valign="middle",
                          color=(0, 0, 0, 1), font_size=esc(sp(11)))
        lbl_fecha.bind(size=lambda i, v: setattr(i, "text_size", (v[0], None)))
        box.add_widget(lbl_fecha)

        if dentro:
            valor = fmt_h(self.horas.get(clave, 0)) if self.horas.get(clave) else ""
            celda = CeldaDia(
                clave=clave, valor_inicial=valor,
                estado={"bloqueada": bloqueada, "fondo": fondo},
                app=self.app, size_hint_y=None, height=esc(dp(38)))
            self.celdas[clave] = celda
            box.add_widget(celda)
        else:
            box.add_widget(Label())
        return box


# =====================================================================
# 5. PANTALLAS
# =====================================================================
class PantallaBase(Screen):
    def __init__(self, app, **kw):
        super().__init__(**kw)
        self.app = app
        self.bd = app.bd
        self._lbl_estado = None
        self._lbl_zoom = None
        self._texto_estado = ""

    def barra_superior(self, titulo):
        esc = self.app.escala
        barra = BoxLayout(size_hint_y=None, height=esc(dp(48)),
                          padding=[esc(dp(4)), esc(dp(4))],
                          spacing=esc(dp(4)))
        with barra.canvas.before:
            Color(0.20, 0.40, 0.60, 1)
            Rectangle(pos=barra.pos, size=barra.size)
        barra.bind(pos=self._redib, size=self._redib)

        b_trab = Button(text="Trabajos", size_hint_x=None,
                        width=esc(dp(80)), font_size=esc(sp(12)))
        b_trab.bind(on_release=lambda *_: self.app.ir("trabajos"))
        b_res = Button(text="Resumen", size_hint_x=None,
                       width=esc(dp(80)), font_size=esc(sp(12)))
        b_res.bind(on_release=lambda *_: self.app.ir("resumen"))
        b_aj = Button(text="Ajustes", size_hint_x=None,
                      width=esc(dp(80)), font_size=esc(sp(12)))
        b_aj.bind(on_release=lambda *_: self.app.ir("ajustes"))

        barra.add_widget(b_trab)
        barra.add_widget(b_res)
        barra.add_widget(b_aj)
        barra.add_widget(Label(text=titulo, font_size=esc(sp(13))))

        b_menos = Button(text="A−", size_hint_x=None, width=esc(dp(44)),
                         font_size=esc(sp(14)))
        b_menos.bind(on_release=lambda *_: self.app.cambiar_zoom(-ZOOM_PASO))
        b_mas = Button(text="A+", size_hint_x=None, width=esc(dp(44)),
                       font_size=esc(sp(14)))
        b_mas.bind(on_release=lambda *_: self.app.cambiar_zoom(+ZOOM_PASO))
        barra.add_widget(b_menos)
        barra.add_widget(b_mas)
        return barra

    def barra_estado(self):
        """Dos líneas: mensaje (línea 1) + "Zoom: N %" a la derecha (línea 2)."""
        esc = self.app.escala
        caja = BoxLayout(orientation="vertical",
                         size_hint_y=None, height=esc(dp(56)),
                         padding=[esc(dp(6)), esc(dp(2))],
                         spacing=esc(dp(0)))
        with caja.canvas.before:
            Color(0.10, 0.10, 0.10, 1)
            Rectangle(pos=caja.pos, size=caja.size)
        caja.bind(pos=self._redib_estado, size=self._redib_estado)

        # Línea 1: mensaje
        lbl1 = Label(text=self._texto_estado, color=(1, 1, 1, 1),
                     halign="left", valign="middle",
                     font_size=esc(sp(12)),
                     size_hint_y=None, height=esc(dp(26)))
        lbl1.bind(size=lambda i, v: setattr(i, "text_size", (v[0], None)))
        self._lbl_estado = lbl1

        # Línea 2: zoom a la derecha
        fila2 = BoxLayout(size_hint_y=None, height=esc(dp(24)),
                          spacing=esc(dp(4)))
        lbl2 = Label(text="", color=(0.9, 0.9, 0.9, 1),
                     halign="right", valign="middle",
                     font_size=esc(sp(11)),
                     size_hint=(1, 1))
        lbl2.bind(size=lambda i, v: setattr(i, "text_size", (v[0], None)))
        self._lbl_zoom = lbl2
        fila2.add_widget(lbl2)

        caja.add_widget(lbl1)
        caja.add_widget(fila2)
        self.actualizar_zoom_estado()
        return caja

    @staticmethod
    def _redib(inst, *_):
        inst.canvas.before.clear()
        with inst.canvas.before:
            Color(0.20, 0.40, 0.60, 1)
            Rectangle(pos=inst.pos, size=inst.size)

    @staticmethod
    def _redib_estado(inst, *_):
        inst.canvas.before.clear()
        with inst.canvas.before:
            Color(0.10, 0.10, 0.10, 1)
            Rectangle(pos=inst.pos, size=inst.size)

    def estado(self, texto):
        self._texto_estado = str(texto)
        if self._lbl_estado is not None:
            self._lbl_estado.text = self._texto_estado

    def actualizar_zoom_estado(self):
        if self._lbl_zoom is not None:
            self._lbl_zoom.text = "Zoom: %d %%" % self.app.zoom_pct


class PantallaTrabajos(PantallaBase):
    def __init__(self, app, **kw):
        super().__init__(app, name="trabajos", **kw)
        self._filas_ui = {}
        self._construir()

    def _construir(self):
        self.clear_widgets()
        self._filas_ui.clear()
        esc = self.app.escala
        root = BoxLayout(orientation="vertical")
        root.add_widget(self.barra_superior("Trabajos"))

        fila_per = BoxLayout(size_hint_y=None, height=esc(dp(48)),
                             padding=[esc(dp(4)), esc(dp(4))],
                             spacing=esc(dp(4)))
        btn_prev = BotonTriangulo("izq")
        btn_prev.width = esc(dp(44))
        btn_prev.height = esc(dp(44))
        btn_prev.bind(on_release=lambda *_: self._mover_periodo(-1))
        lbl_per = Label(text=self._texto_periodo(), font_size=esc(sp(13)))
        btn_next = BotonTriangulo("der")
        btn_next.width = esc(dp(44))
        btn_next.height = esc(dp(44))
        btn_next.bind(on_release=lambda *_: self._mover_periodo(+1))
        fila_per.add_widget(btn_prev)
        fila_per.add_widget(lbl_per)
        fila_per.add_widget(btn_next)
        root.add_widget(fila_per)

        cab = BoxLayout(size_hint_y=None, height=esc(dp(28)),
                        spacing=esc(dp(2)), padding=[esc(dp(4)), 0])
        with cab.canvas.before:
            Color(0.86, 0.90, 0.95, 1)
            Rectangle(pos=cab.pos, size=cab.size)
        cab.bind(pos=lambda i, v: self._repintar(i),
                 size=lambda i, v: self._repintar(i))
        for texto, size_x in [("Obra", 3), ("Salario", 2), ("Prod.", 1), ("Integ.", 1)]:
            lbl = Label(text=texto, bold=True, size_hint_x=size_x,
                        halign="left", valign="middle", font_size=esc(sp(12)))
            lbl.bind(size=lambda inst, v: setattr(inst, "text_size", (v[0], v[1])))
            cab.add_widget(lbl)
        root.add_widget(cab)

        scroll = ScrollView()
        cont = GridLayout(cols=1, size_hint_y=None, spacing=esc(dp(1)))
        cont.bind(minimum_height=cont.setter("height"))
        for f in self.bd.listar_trabajos(self.app.periodo):
            fila_ui = FilaTrabajo(self.app, f, self)
            cont.add_widget(fila_ui)
            self._filas_ui[f["id"]] = fila_ui
        scroll.add_widget(cont)
        root.add_widget(scroll)

        salario, horas = self.bd.totales_periodo(self.app.periodo)
        root.add_widget(Label(
            text="%d trabajo(s) · Salario %s · Producción %s h" % (
                len(self._filas_ui), fmt_monto(salario), fmt_h(horas)),
            size_hint_y=None, height=esc(dp(28)),
            font_size=esc(sp(12))))

        fila_botones = BoxLayout(size_hint_y=None, height=esc(dp(48)),
                                  padding=[esc(dp(4)), esc(dp(4))],
                                  spacing=esc(dp(4)))
        for texto, cb in [("Nuevo", self._nuevo), ("Abrir", self._abrir),
                          ("RDP", self._hoja), ("Editar", self._editar),
                          ("Eliminar", self._eliminar), ("Salir", self._salir)]:
            b = Button(text=texto, font_size=esc(sp(13)))
            b.bind(on_release=lambda inst, cb=cb: cb())
            fila_botones.add_widget(b)
        root.add_widget(fila_botones)
        root.add_widget(self.barra_estado())
        self.add_widget(root)

        if self.app.seleccion_trabajo in self._filas_ui:
            self._seleccionar(self.app.seleccion_trabajo)
        else:
            self.app.seleccion_trabajo = None

        self.estado("Selecciona un trabajo para Abrir / RDP / Editar / Eliminar.")

    @staticmethod
    def _repintar(box):
        box.canvas.before.clear()
        with box.canvas.before:
            Color(0.86, 0.90, 0.95, 1)
            Rectangle(pos=box.pos, size=box.size)

    def _texto_periodo(self):
        ini, fin = rango_periodo(self.app.periodo)
        return "%s %s · %s – %s" % (self.bd.nombre_mes(self.app.periodo),
                                     self.app.periodo[:4],
                                     fmt_fecha(ini), fmt_fecha(fin))

    def _mover_periodo(self, delta):
        self.app.periodo = periodo_desplazado(self.app.periodo, delta)
        self.app.seleccion_trabajo = None
        self.app.limpiar_pila()
        self._construir()

    def _seleccionar(self, trabajo_id):
        self.app.seleccion_trabajo = trabajo_id
        for tid, fila in self._filas_ui.items():
            fila.marcar(tid == trabajo_id)
        self.estado("Trabajo seleccionado: id %s" % trabajo_id)

    def _nuevo(self):
        self.app.ir("form_trabajo", trabajo_id=None)

    def _abrir(self):
        tid = self.app.seleccion_trabajo
        if not tid:
            aviso("Selecciona un trabajo de la lista.")
            return
        self.app.ir("pantalla_trabajo", trabajo_id=tid)

    def _hoja(self):
        tid = self.app.seleccion_trabajo
        if not tid:
            aviso("Selecciona un trabajo.")
            return
        self.app.ir("hoja_rdp", trabajo_id=tid)

    def _editar(self):
        tid = self.app.seleccion_trabajo
        if not tid:
            aviso("Selecciona un trabajo.")
            return
        self.app.ir("form_trabajo", trabajo_id=tid)

    def _eliminar(self):
        tid = self.app.seleccion_trabajo
        if not tid:
            aviso("Selecciona un trabajo.")
            return
        tr = self.bd.obtener_trabajo(tid)
        if not tr:
            return

        def _si():
            try:
                self.bd.eliminar_trabajo(tid)
                self.app.seleccion_trabajo = None
                self._construir()
                self.estado("Trabajo eliminado.")
            except Exception as e:
                aviso("No se pudo eliminar: %s" % e)

        confirmar("¿Eliminar el trabajo %s con su cuadrilla y %s h de producción?"
                  % (tr["codigo"], fmt_h(tr["horas"])), _si)

    def _salir(self):
        confirmar("¿Deseas cerrar la aplicación?", self.app.cerrar_app, "Salir")


class PantallaTrabajo(PantallaBase):
    def __init__(self, app, trabajo_id=None, **kw):
        super().__init__(app, name="pantalla_trabajo", **kw)
        self.trabajo_id = trabajo_id
        self._filas_integ = {}
        self._sel_integrante = None
        self._construir()

    def _construir(self):
        self.clear_widgets()
        self._filas_integ.clear()
        esc = self.app.escala
        if not self.trabajo_id:
            self.add_widget(Label(text="Sin trabajo seleccionado."))
            return
        tr = self.bd.obtener_trabajo(self.trabajo_id)
        if not tr:
            self.add_widget(Label(text="El trabajo ya no existe."))
            return
        root = BoxLayout(orientation="vertical")
        root.add_widget(self.barra_superior("Trabajo %s" % tr["codigo"]))
        root.add_widget(Label(
            text="Salario: %s · Producción: %s h" % (
                fmt_monto(tr["salario_creado"]), fmt_h(tr["horas"])),
            size_hint_y=None, height=esc(dp(30)), font_size=esc(sp(13))))

        cab = BoxLayout(size_hint_y=None, height=esc(dp(28)),
                        spacing=esc(dp(2)), padding=[esc(dp(4)), 0])
        with cab.canvas.before:
            Color(0.86, 0.90, 0.95, 1)
            Rectangle(pos=cab.pos, size=cab.size)
        cab.bind(pos=lambda i, v: PantallaTrabajos._repintar(i),
                 size=lambda i, v: PantallaTrabajos._repintar(i))
        for texto, size_x in [("Cuadrilla", 4), ("Rol", 1), ("Prod.", 1)]:
            lbl = Label(text=texto, bold=True, size_hint_x=size_x,
                        halign="left", valign="middle", font_size=esc(sp(12)))
            lbl.bind(size=lambda inst, v: setattr(inst, "text_size", (v[0], v[1])))
            cab.add_widget(lbl)
        root.add_widget(cab)

        scroll = ScrollView()
        cont = GridLayout(cols=1, size_hint_y=None, spacing=esc(dp(1)))
        cont.bind(minimum_height=cont.setter("height"))
        integrantes = self.bd.listar_integrantes(self.trabajo_id)
        for i in integrantes:
            fila_ui = FilaIntegrante(self.app, i, self)
            cont.add_widget(fila_ui)
            self._filas_integ[i["trabajador_id"]] = fila_ui
        scroll.add_widget(cont)
        root.add_widget(scroll)

        fila = BoxLayout(size_hint_y=None, height=esc(dp(44)),
                          spacing=esc(dp(4)), padding=[esc(dp(4)), esc(dp(4))])
        for texto, cb in [("RDP", self._hoja), ("Rango", self._rango),
                          ("+Integ.", self._agregar), ("\u2212Integ.", self._quitar),
                          ("Volver", self._volver)]:
            b = Button(text=texto, font_size=esc(sp(13)))
            b.bind(on_release=lambda inst, cb=cb: cb())
            fila.add_widget(b)
        root.add_widget(fila)
        root.add_widget(self.barra_estado())
        self.add_widget(root)

        if self._sel_integrante in self._filas_integ:
            self._seleccionar_integrante(self._sel_integrante)
        else:
            self._sel_integrante = None
            self.estado("Trabajo %s · Integrantes: %d · Toca uno para quitarlo." % (
                tr["codigo"], len(integrantes)))

    def _seleccionar_integrante(self, trabajador_id):
        self._sel_integrante = trabajador_id
        for wid, fila in self._filas_integ.items():
            fila.marcar(wid == trabajador_id)
        self.estado("Integrante seleccionado: id %s" % trabajador_id)

    def _hoja(self):
        self.app.ir("hoja_rdp", trabajo_id=self.trabajo_id, trabajador_id=self._sel_integrante)

    def _rango(self):
        self.app.ir("form_horas", trabajo_id=self.trabajo_id, trabajador_id=None)

    def _agregar(self):
        self.app.ir("form_integrante", trabajo_id=self.trabajo_id)

    def _quitar(self):
        wid = self._sel_integrante
        if wid is None:
            aviso("Toca un integrante de la lista para quitarlo de la cuadrilla.")
            return
        fila = next((i for i in self.bd.listar_integrantes(self.trabajo_id)
                     if i["trabajador_id"] == wid), None)
        if fila is None:
            self._sel_integrante = None
            self._construir()
            return
        if fila["es_jb"]:
            aviso("El J'B debe permanecer en la cuadrilla de todos los trabajos.\n"
                  "Si cambió el J'B, defínelo en Ajustes.")
            return
        mensaje = "¿Quitar a %s de la cuadrilla de este trabajo?" % fila["nombre"]
        if fila["dias"]:
            mensaje += ("\n\nSe borrarán también sus %d día(s) de producción (%s h) "
                        "en este trabajo." % (fila["dias"], fmt_h(fila["horas"])))
        mensaje += "\n\nEl trabajador no se elimina del catálogo."

        def _si():
            try:
                self.bd.quitar_integrante(self.trabajo_id, wid, borrar_horas=True)
                self._sel_integrante = None
                self._construir()
                self.estado("Integrante quitado de la cuadrilla.")
            except ErrorDatos as e:
                aviso(str(e))
            except Exception as e:
                registrar_error()
                aviso("No se pudo quitar: %s" % e)

        confirmar(mensaje, _si, "Quitar integrante")

    def _volver(self):
        self.app.volver()


class FormTrabajo(PantallaBase):
    def __init__(self, app, trabajo_id=None, **kw):
        super().__init__(app, name="form_trabajo", **kw)
        self.trabajo_id = trabajo_id
        self._construir()

    def _construir(self):
        self.clear_widgets()
        esc = self.app.escala
        tr = None if self.trabajo_id is None else self.bd.obtener_trabajo(self.trabajo_id)
        root = BoxLayout(orientation="vertical")
        root.add_widget(self.barra_superior(
            "Nuevo trabajo" if tr is None else "Editar trabajo"))
        fila_obra = BoxLayout(size_hint_y=None, height=esc(dp(44)),
                              spacing=esc(dp(4)), padding=[esc(dp(4)), esc(dp(4))])
        fila_obra.add_widget(Label(text="Objeto de Obra:", font_size=esc(sp(13))))
        self.ti_obra = TextInput(text="" if tr is None else str(tr["codigo"]),
                                 multiline=False, font_size=esc(sp(13)))
        fila_obra.add_widget(self.ti_obra)
        root.add_widget(fila_obra)
        fila_sal = BoxLayout(size_hint_y=None, height=esc(dp(44)),
                             spacing=esc(dp(4)), padding=[esc(dp(4)), esc(dp(4))])
        fila_sal.add_widget(Label(text="Salario Creado:", font_size=esc(sp(13))))
        self.ti_sal = TextInput(
            text="" if tr is None else fmt_h(tr["salario_creado"]),
            multiline=False, input_type="number", font_size=esc(sp(13)))
        fila_sal.add_widget(self.ti_sal)
        root.add_widget(fila_sal)
        if tr is None:
            jb = self.bd.jb()
            root.add_widget(Label(
                text="J'B: %s" % (jb["nombre"] if jb else "sin definir"),
                size_hint_y=None, height=esc(dp(28)), font_size=esc(sp(12))))
        fila = BoxLayout(size_hint_y=None, height=esc(dp(48)),
                          spacing=esc(dp(4)), padding=[esc(dp(4)), esc(dp(4))])
        b1 = Button(text="Guardar", font_size=esc(sp(13)))
        b1.bind(on_release=lambda *_: self._guardar())
        b2 = Button(text="Volver", font_size=esc(sp(13)))
        b2.bind(on_release=lambda *_: self.app.volver())
        fila.add_widget(b1)
        fila.add_widget(b2)
        root.add_widget(fila)
        root.add_widget(self.barra_estado())
        self.add_widget(root)
        self.estado("Complete los datos del trabajo y pulse Guardar.")

    def _guardar(self):
        try:
            if self.trabajo_id is None:
                tid = self.bd.crear_trabajo(self.app.periodo,
                                            self.ti_obra.text, self.ti_sal.text)
            else:
                self.bd.editar_trabajo(self.trabajo_id,
                                       self.ti_obra.text, self.ti_sal.text)
                tid = self.trabajo_id
            # Refrescar explícitamente la lista de trabajos y abrir el nuevo
            self.app.seleccion_trabajo = tid
            self.app.ir("pantalla_trabajo", trabajo_id=tid)
            self.app.aviso_global("Trabajo guardado (id %s)." % tid)
        except ErrorDatos as e:
            self.estado(str(e))
            aviso(str(e))
        except Exception as e:
            registrar_error()
            self.estado("Error inesperado: %s" % e)
            aviso("Error inesperado: %s" % e)


class FormIntegrante(PantallaBase):
    def __init__(self, app, trabajo_id=None, **kw):
        super().__init__(app, name="form_integrante", **kw)
        self.trabajo_id = trabajo_id
        self._construir()

    def _construir(self):
        self.clear_widgets()
        esc = self.app.escala
        root = BoxLayout(orientation="vertical")
        root.add_widget(self.barra_superior("Agregar integrante"))
        candidatos = {"%d · %s" % (t["id"], t["nombre"]): t["id"]
                      for t in self.bd.candidatos_integrante(self.trabajo_id)}
        self.candidatos = candidatos
        self.sp = Spinner(text="Selecciona" if candidatos else "(sin candidatos)",
                          values=list(candidatos),
                          size_hint_y=None, height=esc(dp(44)),
                          font_size=esc(sp(13)))
        root.add_widget(self.sp)
        fila = BoxLayout(size_hint_y=None, height=esc(dp(48)),
                          spacing=esc(dp(4)), padding=[esc(dp(4)), esc(dp(4))])
        b1 = Button(text="Agregar", font_size=esc(sp(13)))
        b1.bind(on_release=lambda *_: self._agregar())
        b2 = Button(text="Volver", font_size=esc(sp(13)))
        b2.bind(on_release=lambda *_: self.app.volver())
        fila.add_widget(b1)
        fila.add_widget(b2)
        root.add_widget(fila)
        root.add_widget(self.barra_estado())
        self.add_widget(root)
        if candidatos:
            self.estado("Selecciona un trabajador y pulse Agregar.")
        else:
            self.estado("No hay trabajadores disponibles para agregar.")

    def _agregar(self):
        try:
            wid = self.candidatos.get(self.sp.text)
            if wid is None:
                aviso("Selecciona un trabajador.")
                return
            self.bd.agregar_integrante(self.trabajo_id, wid)
            self.app.ir("pantalla_trabajo", trabajo_id=self.trabajo_id)
            self.app.aviso_global("Integrante agregado (id %s)." % wid)
        except ErrorDatos as e:
            self.estado(str(e))
            aviso(str(e))
        except Exception as e:
            registrar_error()
            self.estado("Error inesperado: %s" % e)
            aviso("Error inesperado: %s" % e)


class FormHoras(PantallaBase):
    def __init__(self, app, trabajo_id=None, trabajador_id=None, **kw):
        super().__init__(app, name="form_horas", **kw)
        self.trabajo_id = trabajo_id
        self.trabajador_id = trabajador_id
        self._construir()

    def _construir(self):
        self.clear_widgets()
        esc = self.app.escala
        root = BoxLayout(orientation="vertical")
        root.add_widget(self.barra_superior("Registrar rango"))
        ini, fin = rango_periodo(self.app.periodo)
        self.ti_desde = TextInput(text=ini.isoformat(), multiline=False,
                                  font_size=esc(sp(13)))
        self.ti_hasta = TextInput(text="", multiline=False, font_size=esc(sp(13)))
        self.ti_horas = TextInput(text="8", multiline=False,
                                  input_type="number", font_size=esc(sp(13)))
        for etiqueta, ti in [("Desde:", self.ti_desde),
                             ("Hasta:", self.ti_hasta),
                             ("Horas:", self.ti_horas)]:
            fila = BoxLayout(size_hint_y=None, height=esc(dp(44)),
                             spacing=esc(dp(4)),
                             padding=[esc(dp(4)), esc(dp(4))])
            fila.add_widget(Label(text=etiqueta, font_size=esc(sp(13))))
            fila.add_widget(ti)
            root.add_widget(fila)
        self.objetivos = {"(Toda la cuadrilla)": None}
        for i in self.bd.listar_integrantes(self.trabajo_id):
            if i["es_jb"]:
                continue
            self.objetivos["%d · %s" % (i["trabajador_id"], i["nombre"])] = i["trabajador_id"]
        self.sp_obj = Spinner(text=list(self.objetivos)[0],
                              values=list(self.objetivos),
                              size_hint_y=None, height=esc(dp(44)),
                              font_size=esc(sp(13)))
        fila_obj = BoxLayout(size_hint_y=None, height=esc(dp(44)),
                              spacing=esc(dp(4)), padding=[esc(dp(4)), esc(dp(4))])
        fila_obj.add_widget(Label(text="Quién:", font_size=esc(sp(13))))
        fila_obj.add_widget(self.sp_obj)
        root.add_widget(fila_obj)
        fila = BoxLayout(size_hint_y=None, height=esc(dp(48)),
                          spacing=esc(dp(4)), padding=[esc(dp(4)), esc(dp(4))])
        b1 = Button(text="Guardar", font_size=esc(sp(13)))
        b1.bind(on_release=lambda *_: self._guardar())
        b2 = Button(text="Volver", font_size=esc(sp(13)))
        b2.bind(on_release=lambda *_: self.app.volver())
        fila.add_widget(b1)
        fila.add_widget(b2)
        root.add_widget(fila)
        root.add_widget(self.barra_estado())
        self.add_widget(root)
        self.estado("Rango de producción · domingos omitidos.")

    def _guardar(self):
        try:
            objetivo = self.objetivos.get(self.sp_obj.text)
            n = self.bd.guardar_horas(self.trabajo_id, objetivo,
                                      self.ti_desde.text, self.ti_hasta.text,
                                      self.ti_horas.text, True)
            self.app.ir("hoja_rdp", trabajo_id=self.trabajo_id)
            aviso("%d registro(s) guardado(s)." % n)
        except ErrorDatos as e:
            self.estado(str(e))
            aviso(str(e))


class PantallaHojaRDP(PantallaBase):
    def __init__(self, app, trabajo_id=None, foco_trabajador=None, **kw):
        super().__init__(app, name="hoja_rdp", **kw)
        self.trabajo_id = trabajo_id
        self.foco_trabajador = foco_trabajador
        self.calendario = None
        self.sp = None
        self.lbl_sub = None
        self.btn_rango = None
        self._construir()

    def _construir(self):
        self.clear_widgets()
        esc = self.app.escala
        if not self.trabajo_id:
            self.add_widget(Label(text="Sin trabajo seleccionado."))
            return
        tr = self.bd.obtener_trabajo(self.trabajo_id)
        if not tr:
            self.add_widget(Label(text="El trabajo ya no existe."))
            return
        root = BoxLayout(orientation="vertical")
        root.add_widget(self.barra_superior("RDP · %s" % tr["codigo"]))
        root.add_widget(Label(
            text="Salario: %s" % fmt_monto(tr["salario_creado"]),
            size_hint_y=None, height=esc(dp(28)), font_size=esc(sp(13))))

        integrantes = self.bd.listar_integrantes(self.trabajo_id)
        if not integrantes:
            root.add_widget(Label(text="Este trabajo no tiene integrantes.",
                                  size_hint_y=None, height=esc(dp(28))))
            root.add_widget(self.boton_volver(self.app.volver))
            root.add_widget(self.barra_estado())
            self.add_widget(root)
            self.estado("Trabajo sin integrantes. Agrega uno primero.")
            return

        self.etiquetas = {"%d · %s" % (i["trabajador_id"], i["nombre"]): i["trabajador_id"]
                          for i in integrantes}
        ids = list(self.etiquetas.values())
        self.wid = self.foco_trabajador if self.foco_trabajador in ids else ids[0]
        etq_inicial = next(k for k, v in self.etiquetas.items() if v == self.wid)
        fila_sel = BoxLayout(size_hint_y=None, height=esc(dp(44)),
                             spacing=esc(dp(4)), padding=[esc(dp(4)), esc(dp(4))])
        self.sp = Spinner(text=etq_inicial, values=list(self.etiquetas),
                          size_hint_y=None, height=esc(dp(44)),
                          font_size=esc(sp(13)))
        self.sp.bind(text=self._elegir)
        fila_sel.add_widget(Label(text="Trabajador:",
                                  size_hint_x=None, width=esc(dp(90)),
                                  font_size=esc(sp(13))))
        fila_sel.add_widget(self.sp)
        root.add_widget(fila_sel)

        self.lbl_sub = Label(text="", size_hint_y=None, height=esc(dp(24)),
                              font_size=esc(sp(12)))
        root.add_widget(self.lbl_sub)

        scroll = ScrollView()
        self.calendario = CalendarioRDP(self.app, self.trabajo_id, self.wid)
        scroll.add_widget(self.calendario)
        root.add_widget(scroll)

        fila_botones = BoxLayout(size_hint_y=None, height=esc(dp(48)),
                                  spacing=esc(dp(4)), padding=[esc(dp(4)), esc(dp(4))])
        self.btn_rango = Button(text="Registrar rango", font_size=esc(sp(13)))
        self.btn_rango.bind(on_release=lambda *_: self._rango())
        b_volver = Button(text="Volver", font_size=esc(sp(13)))
        b_volver.bind(on_release=lambda *_: self.app.volver())
        fila_botones.add_widget(self.btn_rango)
        fila_botones.add_widget(b_volver)
        root.add_widget(fila_botones)
        root.add_widget(self.barra_estado())
        self.add_widget(root)
        self.app.pantalla_rdp = self
        self.app.trabajo_id_actual = self.trabajo_id
        self.actualizar_totales()
        self.estado("Selecciona un trabajador para editar horas.")
        self._actualizar_boton_rango()

    def _actualizar_boton_rango(self):
        if self.btn_rango:
            self.btn_rango.disabled = (self.wid is None)

    def _elegir(self, spinner, texto):
        wid = self.etiquetas.get(texto)
        if wid is None:
            return
        if wid != self.wid:
            self.wid = wid
            if self.calendario:
                self.calendario.wid = wid
                self.calendario._construir()
            self.actualizar_totales()
            self._actualizar_boton_rango()
            self.estado("Trabajador seleccionado: %s" % texto)

    def _rango(self):
        if self.wid is None:
            aviso("Selecciona un trabajador primero.")
            return
        self.app.ir("form_horas", trabajo_id=self.trabajo_id,
                    trabajador_id=self.wid)

    def actualizar_totales(self):
        if not self.calendario:
            return
        total = sum(h for (w, _), h in self.calendario.horas.items() if w == self.wid)
        nombre = next((k for k, v in self.etiquetas.items() if v == self.wid), "")
        if self.lbl_sub:
            self.lbl_sub.text = "%s · Producción: %s h" % (nombre, fmt_h(total))
        self.actualizar_zoom_estado()


class PantallaResumen(PantallaBase):
    def __init__(self, app, **kw):
        super().__init__(app, name="resumen", **kw)
        self._construir()

    def _construir(self):
        self.clear_widgets()
        esc = self.app.escala
        root = BoxLayout(orientation="vertical")
        root.add_widget(self.barra_superior("Resumen"))
        filas = self.bd.resumen_periodo(self.app.periodo)
        scroll = ScrollView()
        cont = GridLayout(cols=4, size_hint_y=None, spacing=esc(dp(1)))
        cont.bind(minimum_height=cont.setter("height"))
        for h in ("Trabajador", "Prod.", "Días", "Trab."):
            cont.add_widget(Label(text=h, bold=True, size_hint_y=None,
                                  height=esc(dp(28)), font_size=esc(sp(12))))
        for f in filas:
            cont.add_widget(Label(text=f["nombre"], size_hint_y=None,
                                  height=esc(dp(30)), font_size=esc(sp(13))))
            cont.add_widget(Label(text=fmt_h(f["horas"]), size_hint_y=None,
                                  height=esc(dp(30)), font_size=esc(sp(13))))
            cont.add_widget(Label(text=str(f["dias"]), size_hint_y=None,
                                  height=esc(dp(30)), font_size=esc(sp(13))))
            cont.add_widget(Label(text=str(f["trabajos"]), size_hint_y=None,
                                  height=esc(dp(30)), font_size=esc(sp(13))))
        scroll.add_widget(cont)
        root.add_widget(scroll)
        salario, horas = self.bd.totales_periodo(self.app.periodo)
        root.add_widget(Label(
            text="%d trab. · Prod. %s h · Salario %s" % (
                len(filas), fmt_h(horas), fmt_monto(salario)),
            size_hint_y=None, height=esc(dp(28)), font_size=esc(sp(12))))
        fila = BoxLayout(size_hint_y=None, height=esc(dp(48)),
                          spacing=esc(dp(4)), padding=[esc(dp(4)), esc(dp(4))])
        b1 = Button(text="Exportar XLSX", font_size=esc(sp(13)))
        b1.bind(on_release=lambda *_: self._exportar())
        b2 = Button(text="Importar XLSX", font_size=esc(sp(13)))
        b2.bind(on_release=lambda *_: self._importar())
        b3 = Button(text="Volver", font_size=esc(sp(13)))
        b3.bind(on_release=lambda *_: self.app.volver())
        fila.add_widget(b1)
        fila.add_widget(b2)
        fila.add_widget(b3)
        root.add_widget(fila)
        root.add_widget(self.barra_estado())
        self.add_widget(root)
        self.estado("Resumen del periodo %s" % self.app.periodo)

    def _exportar(self):
        try:
            ruta = self.bd.exportar_periodo_xlsx(self.app.periodo)
            self.estado("Exportado: %s" % ruta)
            aviso("Archivo creado:\n%s" % ruta)
        except ErrorDatos as e:
            self.estado(str(e))
            aviso(str(e))

    def _importar(self):
        if not PLYER_OK:
            aviso("Falta plyer. Instálalo con: pip install plyer")
            return
        try:
            filechooser.open_file(on_selection=self._al_elegir,
                                   filters=[("Excel", "*.xlsx")])
        except Exception as e:
            aviso("No se pudo abrir el selector: %s" % e)

    def _al_elegir(self, seleccion):
        if not seleccion:
            return
        try:
            res = self.bd.importar_periodo_xlsx(seleccion[0], self.app.periodo)
            self._construir()
            aviso("Importado: %d trabajos, %d horas, %d integrantes nuevos.\nErrores: %d" % (
                len(res["trabajos"]), res["horas"],
                res["integrantes_nuevos"], len(res["errores"])))
        except ErrorDatos as e:
            aviso(str(e))


class PantallaAjustes(PantallaBase):
    def __init__(self, app, **kw):
        super().__init__(app, name="ajustes", **kw)
        self._construir()

    def _construir(self):
        self.clear_widgets()
        esc = self.app.escala
        root = BoxLayout(orientation="vertical")
        root.add_widget(self.barra_superior("Ajustes"))
        e = self.bd.estadisticas()
        root.add_widget(Label(
            text=("BD: %s\n\nTrabajadores: %d · Obras: %d · Trabajos: %d\n"
                  "Integrantes: %d · Prod.: %d") % (
                self.bd.ruta, e["trabajador"], e["objeto_obra"], e["trabajo"],
                e["integrante"], e["horas_trabajo"]),
            size_hint_y=None, height=esc(dp(140)), font_size=esc(sp(12)),
            halign="left", valign="top"))
        for texto, cb in [
            ("Diagnóstico de BD", self._diagnostico),
            ("Crear respaldo", self._respaldar),
            ("Verificar integridad", self._verificar),
            ("Copiar BD a carpeta de la app", self._copiar_bd),
            ("Volver", self.app.volver),
        ]:
            b = Button(text=texto, size_hint_y=None, height=esc(dp(48)),
                       font_size=esc(sp(13)))
            b.bind(on_release=lambda inst, cb=cb: cb())
            root.add_widget(b)
        root.add_widget(self.barra_estado())
        self.add_widget(root)
        self.estado("Ajustes y herramientas de BD.")

    def _generar_diagnostico(self):
        lineas = []
        lineas.append("=== DIAGNÓSTICO BD RDP v10 ===")
        lineas.append("Fecha: %s" % datetime.now().isoformat(timespec="seconds"))
        lineas.append("")
        lineas.append("== Rutas ==")
        lineas.append("RUTA_PRINCIPAL = %r" % RUTA_PRINCIPAL)
        lineas.append("CARPETA_DATOS configurada = %r" % CARPETA_DATOS)
        lineas.append("AVISO_CARPETA = %r" % AVISO_CARPETA)
        lineas.append("Plataforma = %s" % platform)
        lineas.append("Ruta que usa BaseDatos = %r" % self.bd.ruta)
        lineas.append("RUTA_LOG = %r" % RUTA_LOG)
        lineas.append("")
        lineas.append("== Carpeta ==")
        carpeta = os.path.dirname(os.path.abspath(self.bd.ruta)) or "."
        lineas.append("Carpeta = %r" % carpeta)
        lineas.append("Existe = %s" % os.path.isdir(carpeta))
        lineas.append("Acceso escritura = %s" % os.access(carpeta, os.W_OK))
        try:
            contenido = sorted(os.listdir(carpeta))
            lineas.append("Contenido (%d):" % len(contenido))
            for nombre in contenido[:50]:
                ruta_item = os.path.join(carpeta, nombre)
                try:
                    tam = os.path.getsize(ruta_item)
                except Exception:
                    tam = -1
                lineas.append("  · %s (%d bytes)" % (nombre, tam))
        except Exception as ex:
            lineas.append("Error listando: %s" % ex)
        lineas.append("")
        lineas.append("== Archivo de la BD ==")
        existe = os.path.isfile(self.bd.ruta)
        lineas.append("Existe = %s" % existe)
        if existe:
            lineas.append("Tamaño = %d bytes" % os.path.getsize(self.bd.ruta))
        lineas.append("")
        lineas.append("== Permisos Android ==")
        try:
            from android.permissions import check_permission, Permission
            for p in (Permission.READ_EXTERNAL_STORAGE,
                      Permission.WRITE_EXTERNAL_STORAGE):
                try:
                    lineas.append("%s = %s" % (p, check_permission(p)))
                except Exception as ex:
                    lineas.append("%s = (error: %s)" % (p, ex))
        except Exception as ex:
            lineas.append("(no disponible) %s" % ex)
        lineas.append("")
        lineas.append("== Rutas de origen para Copiar BD ==")
        for r in RUTAS_ORIGEN_BD:
            lineas.append("  %s → existe=%s" % (r, os.path.isfile(r)))
        lineas.append("")
        lineas.append("== Tablas ==")
        for tabla in ("trabajador", "objeto_obra", "trabajo",
                      "integrante", "horas_trabajo", "parametro",
                      "incidencia", "registro_incidencia"):
            try:
                n = self.bd._leer("SELECT COUNT(*) FROM %s" % tabla)[0][0]
                lineas.append("  %-22s = %d filas" % (tabla, n))
            except Exception as ex:
                lineas.append("  %-22s = ERROR: %s" % (tabla, ex))
        lineas.append("")
        lineas.append("== Módulos ==")
        lineas.append("openpyxl = %s" % OPENPYXL_OK)
        lineas.append("plyer = %s" % PLYER_OK)
        lineas.append("")
        lineas.append("== Fin del diagnóstico ==")
        texto = "\n".join(lineas)
        print(texto)
        global RUTA_DIAGNOSTICO
        try:
            RUTA_DIAGNOSTICO = os.path.join(carpeta, "diagnostico_bd.txt")
            with open(RUTA_DIAGNOSTICO, "w", encoding="utf-8") as f:
                f.write(texto)
            lineas.append("")
            lineas.append("Guardado también en: %s" % RUTA_DIAGNOSTICO)
            texto = "\n".join(lineas)
        except Exception as ex:
            print("No se pudo guardar diagnostico_bd.txt: %s" % ex)
        return texto

    def _diagnostico(self):
        texto = self._generar_diagnostico()
        popup_texto_copiable(texto, "Diagnóstico (selecciona y copia)")

    def _copiar_bd(self):
        destino = self.bd.ruta
        for origen in RUTAS_ORIGEN_BD:
            if os.path.isfile(origen) and os.path.abspath(origen) != os.path.abspath(destino):
                try:
                    os.makedirs(os.path.dirname(destino), exist_ok=True)
                    if os.path.isfile(destino):
                        try:
                            shutil.copy2(destino, destino + ".antes_de_copiar")
                        except Exception:
                            pass
                    shutil.copy2(origen, destino)
                    aviso("BD copiada desde:\n%s\na:\n%s\n\nReinicia la app." % (
                        origen, destino))
                    return
                except Exception as ex:
                    aviso("Error copiando desde %s:\n%s" % (origen, ex))
                    return
        aviso("No se encontró la BD en ninguna ruta conocida.\n"
              "Rutas buscadas:\n" + "\n".join(RUTAS_ORIGEN_BD or
                                               ["(ninguna en este sistema)"]))

    def _respaldar(self):
        aviso("Respaldo creado:\n%s" % self.bd.respaldar())

    def _verificar(self):
        r = self.bd.verificar()
        aviso("Integridad correcta ✔" if r == "ok" else "Problema: %s" % r)


# =====================================================================
# 6. APP
# =====================================================================
class RDPApp(App):
    titulo = StringProperty(APP_TITULO)
    factor_zoom = NumericProperty(ZOOM_BASE)

    def __init__(self, **kw):
        super().__init__(**kw)
        self.bd = None
        self.periodo = "2026-10"
        self.zoom_pct = 100
        self.seleccion_trabajo = None
        self.pantalla_rdp = None
        self.trabajo_id_actual = None
        self._pila = []
        self._firma_actual_guardada = ("trabajos", {})

    def escala(self, valor):
        try:
            f = float(self.factor_zoom)
        except Exception:
            f = 1.0
        if f < 0.4:
            f = 0.4
        if f > 2.0:
            f = 2.0
        return valor * f

    def build(self):
        if platform == "android":
            try:
                from android.permissions import request_permissions, Permission
                request_permissions(
                    [Permission.READ_EXTERNAL_STORAGE,
                     Permission.WRITE_EXTERNAL_STORAGE],
                    self._al_permisos)
            except Exception:
                pass
        try:
            Window.bind(on_keyboard=self._al_boton_atras)
        except Exception:
            pass
        return self._construir_ui()

    def aviso_global(self, texto):
        """Muestra un aviso breve en la barra de estado de la pantalla actual."""
        try:
            pantalla = self.sm.current_screen
            if hasattr(pantalla, "estado"):
                pantalla.estado(texto)
        except Exception:
            pass

    def _al_permisos(self, permisos, resultados):
        print("Permisos:", list(zip(permisos, resultados)))
        Clock.schedule_once(lambda dt: self._recargar_ui(), 0.5)

    def _recargar_ui(self):
        try:
            if self.root is None:
                return
            for w in list(self.root.children):
                self.root.remove_widget(w)
            nuevo = self._construir_ui()
            if hasattr(nuevo, "children"):
                for w in list(nuevo.children):
                    nuevo.remove_widget(w)
                    self.root.add_widget(w)
        except Exception:
            registrar_error()

    def _construir_ui(self):
        try:
            carpeta = carpeta_base()
            global RUTA_LOG
            RUTA_LOG = os.path.join(carpeta, "errores.log")
            ruta_bd = os.path.join(carpeta, DB_NOMBRE)
            self.bd = BaseDatos(ruta_bd)
            self.periodo = self.bd.periodo_inicial()
            self.zoom_pct = self.bd.get_zoom()
            self.factor_zoom = ZOOM_BASE * self.zoom_pct / 100.0
        except Exception as e:
            registrar_error()
            return Label(text="Error al abrir BD:\n%s\n\n%s" % (e, AVISO_CARPETA))

        self.sm = ScreenManager(transition=SlideTransition(duration=0.15))
        self.sm.add_widget(PantallaTrabajos(self))
        self.sm.add_widget(PantallaResumen(self))
        self.sm.add_widget(PantallaAjustes(self))
        self._pila = []
        self._firma_actual_guardada = ("trabajos", {})
        avisos = []
        if AVISO_CARPETA:
            avisos.append(AVISO_CARPETA)
        if getattr(self.bd, "aviso", ""):
            avisos.append(self.bd.aviso)
        if avisos:
            Clock.schedule_once(
                lambda dt: aviso("\n\n".join(avisos), "Aviso"), 0.3)
        return self.sm

    def _crear_pantalla(self, nombre, kwargs):
        if nombre == "trabajos":
            return PantallaTrabajos(self)
        if nombre == "resumen":
            return PantallaResumen(self)
        if nombre == "ajustes":
            return PantallaAjustes(self)
        if nombre == "pantalla_trabajo":
            return PantallaTrabajo(self, trabajo_id=kwargs.get("trabajo_id"))
        if nombre == "form_trabajo":
            return FormTrabajo(self, trabajo_id=kwargs.get("trabajo_id"))
        if nombre == "form_integrante":
            return FormIntegrante(self, trabajo_id=kwargs.get("trabajo_id"))
        if nombre == "form_horas":
            return FormHoras(self, trabajo_id=kwargs.get("trabajo_id"),
                             trabajador_id=kwargs.get("trabajador_id"))
        if nombre == "hoja_rdp":
            return PantallaHojaRDP(self, trabajo_id=kwargs.get("trabajo_id"),
                                    foco_trabajador=kwargs.get("trabajador_id"))
        return None

    PANTALLAS_PRINCIPALES = ("trabajos", "resumen", "ajustes")
    PANTALLAS_FORMULARIO = ("form_trabajo", "form_integrante", "form_horas")

    def ir(self, nombre, **kwargs):
        # Estas pantallas SÍ necesitan un trabajo existente. «form_trabajo» no está en
        # la lista: sin trabajo_id es el formulario de trabajo NUEVO (antes se desviaba
        # a la lista y por eso «Nuevo» nunca abría nada).
        if nombre in ("pantalla_trabajo", "form_integrante", "form_horas", "hoja_rdp"):
            if kwargs.get("trabajo_id") is None:
                nombre = "trabajos"
                kwargs = {}
        actual = self.sm.current if hasattr(self, "sm") else None
        if nombre == "trabajos":
            # La lista es la raíz: aquí empieza un historial nuevo. Sin esto, la pantalla
            # de la que se venía quedaba guardada y reaparecía en el siguiente «Volver».
            self._pila = []
        else:
            # Los formularios no se apilan: al guardar, «Volver» no debe reabrirlos.
            if (actual and actual not in self.PANTALLAS_PRINCIPALES
                    and actual not in self.PANTALLAS_FORMULARIO):
                self._pila.append(self._firma_actual_guardada)
            # Si se regresa a la pantalla del tope del historial, se descarta esa entrada
            # (es volver, no avanzar): así «Volver» no repite la misma pantalla.
            if self._pila and self._pila[-1] == (nombre, dict(kwargs)):
                self._pila.pop()
        for pantalla in list(self.sm.screens):
            if pantalla.name == nombre:
                self.sm.remove_widget(pantalla)
        pantalla = self._crear_pantalla(nombre, kwargs)
        if pantalla is None:
            return
        self.sm.add_widget(pantalla)
        self.sm.current = nombre
        self._firma_actual_guardada = (nombre, dict(kwargs))
        if nombre != "hoja_rdp":
            self.pantalla_rdp = None

    def volver(self):
        if self._pila:
            nombre, kwargs = self._pila.pop()
            for pantalla in list(self.sm.screens):
                if pantalla.name == nombre:
                    self.sm.remove_widget(pantalla)
            pantalla = self._crear_pantalla(nombre, kwargs)
            if pantalla is not None:
                self.sm.add_widget(pantalla)
                self.sm.current = nombre
                self._firma_actual_guardada = (nombre, dict(kwargs))
                if nombre != "hoja_rdp":
                    self.pantalla_rdp = None
                return
        self.ir("trabajos")

    def limpiar_pila(self):
        self._pila = []

    def _al_boton_atras(self, window, key, scancode=None, codepoint=None, modifiers=None):
        try:
            if key == 27:
                if self.sm.current == "trabajos":
                    confirmar("¿Deseas cerrar la aplicación?",
                              self.cerrar_app, "Salir")
                else:
                    self.volver()
                return True
        except Exception:
            pass
        return False

    def cerrar_app(self):
        try:
            if self.bd:
                self.bd.cerrar()
        except Exception:
            pass
        try:
            App.get_running_app().stop()
        except Exception:
            pass

    def cambiar_zoom(self, delta):
        try:
            nuevo = self.bd.set_zoom(self.zoom_pct + delta)
        except Exception:
            return
        if nuevo == self.zoom_pct:
            return
        self.zoom_pct = nuevo
        self.factor_zoom = ZOOM_BASE * self.zoom_pct / 100.0
        # Actualizar la línea de estado "Zoom: N %" en la pantalla actual
        try:
            pantalla = self.sm.current_screen
            if hasattr(pantalla, "actualizar_zoom_estado"):
                pantalla.actualizar_zoom_estado()
        except Exception:
            pass
        # Reconstruir la pantalla con el nuevo factor
        nombre, kwargs = self._firma_actual_guardada
        for pantalla in list(self.sm.screens):
            if pantalla.name == nombre:
                self.sm.remove_widget(pantalla)
        nueva = self._crear_pantalla(nombre, kwargs)
        if nueva is None:
            self.ir("trabajos")
            return
        self.sm.add_widget(nueva)
        self.sm.current = nombre

    def actualizar_totales_rdp(self):
        if self.pantalla_rdp:
            try:
                self.pantalla_rdp.actualizar_totales()
            except Exception:
                pass

    def on_stop(self):
        if self.bd:
            self.bd.cerrar()


def main():
    RDPApp().run()


if __name__ == "__main__":
    main()