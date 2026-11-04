#!/usr/bin/env python3
"""
Buscador de negocios con Google Maps (Places API) + prospectos guardados en SQLite.
Interfaz gráfica con Tkinter — un solo archivo ejecutable.

Archivos que crea junto al programa (o en la carpeta indicada por la variable
de entorno PROSPECCION_DATOS):
  prospectos.db        base SQLite (historial, guardados, estados, notas, uso de API)
  palabras_clave.json  términos de búsqueda por tipo de negocio (opcional, editable)
  informes/            exportaciones a Excel / CSV
Clave de Google: variable GOOGLE_MAPS_API_KEY o archivo .env con GOOGLE_MAPS_API_KEY=...
"""

from __future__ import annotations

import codecs
import csv
import json
import os
import queue
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
import tkinter as tk
import urllib.error
import urllib.request
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from tkinter import font as tkfont
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Callable, Iterable, Optional

from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

CIUDADES_POR_PAIS: dict[str, list[str]] = {
    "España": [
        "Madrid", "Barcelona", "Valencia", "Sevilla", "Zaragoza", "Málaga", "Murcia",
        "Palma", "Las Palmas", "Bilbao", "Alicante", "Córdoba", "Valladolid", "Vigo",
        "Gijón", "Granada", "Vitoria", "Elche", "Oviedo", "Badalona",
    ],
    "México": [
        "Ciudad de México", "Guadalajara", "Monterrey", "Puebla", "Tijuana", "León",
        "Ciudad Juárez", "Zapopan", "Mérida", "Cancún", "Querétaro", "San Luis Potosí",
        "Aguascalientes", "Hermosillo", "Saltillo", "Morelia", "Culiacán", "Veracruz",
        "Acapulco", "Chihuahua", "Toluca", "Reynosa", "Torreón", "Oaxaca",
    ],
    "Colombia": [
        "Bogotá", "Medellín", "Cali", "Barranquilla", "Cartagena", "Cúcuta",
        "Bucaramanga", "Pereira", "Santa Marta", "Ibagué", "Manizales", "Villavicencio",
        "Pasto", "Montería", "Armenia", "Neiva", "Sincelejo", "Valledupar", "Popayán",
        "Tunja", "Riohacha", "Florencia", "Yopal", "Quibdó", "San Andrés",
    ],
    "Argentina": [
        "Buenos Aires", "Córdoba", "Rosario", "Mendoza", "La Plata", "San Miguel de Tucumán",
        "Mar del Plata", "Salta", "Santa Fe", "San Juan", "Resistencia", "Neuquén",
        "Santiago del Estero", "Corrientes", "Bahía Blanca", "Posadas", "Paraná",
        "Formosa", "San Salvador de Jujuy", "La Rioja",
    ],
    "Perú": [
        "Lima", "Arequipa", "Trujillo", "Chiclayo", "Piura", "Cusco", "Iquitos",
        "Huancayo", "Tacna", "Juliaca", "Pucallpa", "Cajamarca", "Chimbote", "Ayacucho",
        "Tarapoto", "Huánuco", "Ica", "Puno", "Tumbes", "Huaraz",
    ],
    "Venezuela": [
        "Caracas", "Maracaibo", "Valencia", "Barquisimeto", "Maracay", "Ciudad Guayana",
        "Barcelona", "Maturín", "Ciudad Bolívar", "San Cristóbal", "Cumaná", "Mérida",
        "Cabimas", "Barinas", "Turmero", "Puerto La Cruz", "Petare", "Los Teques",
    ],
    "Chile": [
        "Santiago", "Valparaíso", "Concepción", "La Serena", "Antofagasta", "Temuco",
        "Rancagua", "Talca", "Arica", "Iquique", "Puerto Montt", "Coquimbo", "Osorno",
        "Valdivia", "Chillán", "Copiapó", "Calama", "Quilpué", "Viña del Mar",
    ],
    "Ecuador": [
        "Quito", "Guayaquil", "Cuenca", "Santo Domingo", "Machala", "Durán", "Manta",
        "Portoviejo", "Loja", "Ambato", "Esmeraldas", "Quevedo", "Riobamba", "Milagro",
        "Ibarra", "Latacunga", "Tulcán", "Babahoyo", "Santa Elena",
    ],
    "Guatemala": [
        "Ciudad de Guatemala", "Mixco", "Villa Nueva", "Quetzaltenango", "Escuintla",
        "San Juan Sacatepéquez", "Villa Canales", "Petapa", "Cobán", "Huehuetenango",
        "Chimaltenango", "Amatitlán", "Chinautla", "Antigua Guatemala", "Jalapa",
    ],
    "Cuba": [
        "La Habana", "Santiago de Cuba", "Camagüey", "Holguín", "Santa Clara", "Guantánamo",
        "Bayamo", "Las Tunas", "Cienfuegos", "Pinar del Río", "Matanzas", "Ciego de Ávila",
        "Sancti Spíritus", "Manzanillo", "Artemisa",
    ],
    "Bolivia": [
        "La Paz", "Santa Cruz", "Cochabamba", "Sucre", "Oruro", "Tarija", "Potosí",
        "Trinidad", "El Alto", "Montero", "Warnes", "Sacaba", "Quillacollo", "Riberalta",
    ],
    "República Dominicana": [
        "Santo Domingo", "Santiago de los Caballeros", "La Romana", "San Pedro de Macorís",
        "Puerto Plata", "San Francisco de Macorís", "La Vega", "Higüey", "San Cristóbal",
        "Moca", "Barahona", "Boca Chica", "Baní", "Azua",
    ],
    "Honduras": [
        "Tegucigalpa", "San Pedro Sula", "La Ceiba", "Choloma", "El Progreso",
        "Comayagua", "Choluteca", "Puerto Cortés", "Danlí", "Juticalpa", "Villanueva",
        "Tela", "Santa Rosa de Copán", "Siguatepeque",
    ],
    "Paraguay": [
        "Asunción", "Ciudad del Este", "San Lorenzo", "Luque", "Capiatá", "Lambaré",
        "Fernando de la Mora", "Encarnación", "Pedro Juan Caballero", "Caaguazú",
        "Coronel Oviedo", "Itauguá", "Villarrica", "Concepción",
    ],
    "El Salvador": [
        "San Salvador", "Santa Ana", "San Miguel", "Soyapango", "Mejicanos", "Santa Tecla",
        "Apopa", "Delgado", "Sonsonate", "San Martín", "Ilopango", "Cojutepeque",
        "Usulután", "Ahuachapán", "La Unión",
    ],
    "Nicaragua": [
        "Managua", "León", "Masaya", "Tipitapa", "Chinandega", "Matagalpa", "Estelí",
        "Granada", "Juigalpa", "Jinotepe", "Bluefields", "Rivas", "Ocotal", "Diriamba",
    ],
    "Costa Rica": [
        "San José", "Alajuela", "Cartago", "Heredia", "Limón", "Puntarenas", "Liberia",
        "Desamparados", "San Isidro", "Curridabat", "San Carlos", "Pérez Zeledón",
        "Grecia", "San Ramón", "Quesada",
    ],
    "Panamá": [
        "Ciudad de Panamá", "San Miguelito", "David", "Colón", "Santiago", "La Chorrera",
        "Chitré", "Penonomé", "Aguadulce", "Arraiján", "Changuinola", "La Concepción",
    ],
    "Uruguay": [
        "Montevideo", "Salto", "Paysandú", "Las Piedras", "Rivera", "Maldonado",
        "Tacuarembó", "Melo", "Mercedes", "Artigas", "Minas", "San José de Mayo",
        "Durazno", "Florida", "Treinta y Tres",
    ],
    "Guinea Ecuatorial": ["Malabo", "Bata", "Ebebiyín", "Mongomo", "Aconibe", "Evinayong"],
    "Puerto Rico": [
        "San Juan", "Bayamón", "Carolina", "Ponce", "Caguas", "Guaynabo", "Mayagüez",
        "Trujillo Alto", "Arecibo", "Fajardo", "Humacao", "Vega Baja", "Cayey", "Aguadilla",
    ],
}

PAISES = list(CIUDADES_POR_PAIS.keys())

_PAIS_REGION: dict[str, str] = {
    "Colombia": "CO", "México": "MX", "Argentina": "AR", "Chile": "CL", "Perú": "PE",
    "Ecuador": "EC", "Venezuela": "VE", "Bolivia": "BO", "Costa Rica": "CR",
    "El Salvador": "SV", "Guatemala": "GT", "Honduras": "HN", "Nicaragua": "NI",
    "Paraguay": "PY", "Panamá": "PA", "Uruguay": "UY", "Guinea Ecuatorial": "GQ",
    "Puerto Rico": "PR", "España": "ES", "Cuba": "CU", "República Dominicana": "DO",
}

# Palabras clave por defecto: UNA consulta por término ("<término> en <ciudad>, <país>").
# Se pueden cambiar desde el botón «Palabras clave» o editando palabras_clave.json.
PALABRAS_CLAVE_DEFECTO: dict[str, list[str]] = {
    "Comidas rápidas": [
        "comida rápida", "fast food", "hamburguesas", "salchipapas", "arepas",
        "perros calientes", "pizza", "pollo broaster", "pollo asado", "comida corrida",
        "empanadas", "tacos", "sándwiches", "snacks",
    ],
    "Restaurantes": [
        "restaurante", "comida", "asadero", "cafetería", "pizzería", "marisquería",
        "sushi", "comida china", "parrilla", "menú del día", "almuerzos", "brunch",
        "comida típica",
    ],
    "Barberías": [
        "barbería", "barber shop", "peluquería para hombres", "peluquería barbería",
        "corte de cabello hombre", "barbería para niños", "fade", "afeitado", "grooming",
    ],
    "Odontólogos": [
        "odontólogo", "clínica dental", "dentista", "odontología", "consultorio dental",
        "odontopediatra", "ortodoncia", "endodoncia", "periodoncia", "implantes dentales",
        "blanqueamiento dental", "estética dental", "carillas dentales",
    ],
    "Médicos": [
        "médico", "médica", "doctor", "doctora", "médico general", "consultorio médico",
        "medicina general", "consulta médica", "clínica médica", "pediatra", "ginecólogo",
        "dermatólogo", "cardiólogo", "ortopedista", "oftalmólogo", "urólogo", "nutriólogo",
        "nutricionista", "especialista médico",
    ],
    "Psicólogos": [
        "psicólogo", "psicóloga", "psicólogo clínico", "consultorio psicológico",
        "consulta psicológica", "terapeuta", "terapia de pareja", "psicoterapia",
        "psiquiatra", "psicología infantil",
    ],
    "Inmobiliarias": [
        "inmobiliaria", "bienes raíces", "finca raíz", "constructora", "apartamentos en venta",
        "casas en venta", "arriendo apartamentos", "arriendos", "corretaje de propiedades",
        "lotes en venta", "oficinas en arriendo", "propiedades",
    ],
}

MENSAJES: dict[str, str] = {
    "Comidas rápidas": (
        "Hola, vi {nombre} y creo que podrías automatizar tus pedidos por WhatsApp "
        "para no perder clientes cuando no estás disponible. Tengo una solución que "
        "responde automáticamente y toma pedidos 24/7."
    ),
    "Restaurantes": (
        "Hola, vi {nombre} y noté que podrías mejorar la experiencia de tus clientes "
        "con reservas y respuestas automáticas por WhatsApp. Te ayudo a no perder "
        "mesas ni consultas fuera de horario."
    ),
    "Barberías": (
        "Hola, vi {nombre} y creo que un bot de WhatsApp para agendar citas y enviar "
        "recordatorios te ahorraría horas al día y reduciría las citas perdidas."
    ),
    "Odontólogos": (
        "Hola, vi {nombre} y creo que automatizar citas y recordatorios por WhatsApp "
        "mejoraría la experiencia de tus pacientes y liberaría tiempo de tu recepción."
    ),
    "Médicos": (
        "Hola, vi {nombre} y creo que un asistente por WhatsApp para agendar citas, "
        "enviar recordatorios y responder consultas frecuentes liberaría a tu recepción "
        "y evitaría perder pacientes fuera de horario."
    ),
    "Psicólogos": (
        "Hola, vi {nombre} y creo que un asistente por WhatsApp para agendar sesiones, "
        "confirmar asistencia y atender el primer contacto te reduciría los no-shows "
        "y evitaría perder pacientes nuevos fuera de horario."
    ),
    "Inmobiliarias": (
        "Hola, vi {nombre} y creo que un asistente por WhatsApp que califique leads "
        "y responda consultas 24/7 te ayudaría a cerrar más negocios sin perder "
        "interesados fuera de horario."
    ),
}
MENSAJE_GENERICO = (
    "Hola, vi {nombre} y creo que podrías automatizar tu atención por WhatsApp."
)

ESTADOS = ("nuevo", "contactado", "respondió", "descartado")

MIN_RESEÑAS_OPCIONES = ["0", "5", "10", "20", "50", "100"]
OPCIONES_POR_PAGINA = ["10", "25", "50", "100", "200", "300", "400", "500", "600", "todos"]
PAGINA_TODOS = 10_000

# Umbrales de «Problemas detectados» (fáciles de cambiar)
POCAS_RESEÑAS_UMBRAL = 10

# Comportamiento de búsqueda / API
MAX_PAGINAS_POR_CONSULTA = 3        # tope de Google: 3 páginas × 20 = 60 por consulta
UMBRAL_OMITIR_TERMINO = 3           # veces seguidas aportando 0 negocios nuevos
TOPE_PETICIONES_MAX = 200           # tope duro de peticiones al llenar una página
TOPE_CARGAR_TODO = 150              # tope de «Cargar todo y exportar»
MAX_REINTENTOS = 4
TRAS_SIN_ESTRICTOS = 4              # peticiones seguidas sin hallar más que cumplan el mínimo de reseñas
MAX_ENLACES_EXCEL = 60_000          # Excel admite ~65 530 hipervínculos por hoja
ESPERA_REINTENTO_BASE = 1.0         # segundos (se duplica en cada reintento)
ESPERA_TOKEN_PAGINA = 0.5           # pausa antes de pedir la página siguiente
ESPERA_TOKEN_REINTENTO = 1.5        # si Google aún no activó el token: espera 1.5 s, 3 s, 4.5 s
COSTO_ESTIMADO_POR_PETICION_USD = 0.035
CREDITO_MENSUAL_USD = 200.0

_DIR_PROGRAMA = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.environ.get("PROSPECCION_DATOS") or _DIR_PROGRAMA
INFORMES_DIR = os.path.join(_BASE_DIR, "informes")
RUTA_DB = os.path.join(_BASE_DIR, "prospectos.db")
RUTA_PALABRAS = os.path.join(_BASE_DIR, "palabras_clave.json")
RUTA_USO_JSON_ANTIGUO = os.path.join(_BASE_DIR, "api_usage.json")

CAMPOS_EDITABLES = ("nombre", "ciudad", "telefono", "resenas", "web", "direccion", "link_maps")

COLUMNAS_EXPORT = (
    "Negocio", "Ciudad", "Teléfono", "Reseñas", "Web / Instagram",
    "Problemas detectados", "Estado", "Notas", "Link Google Maps", "Dirección",
)


# ---------------------------------------------------------------------------
# Modelo de datos
# ---------------------------------------------------------------------------

def _ahora() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def clasificar_web(url: str) -> str:
    """'' (sin web) | 'instagram' | 'facebook' | 'web'."""
    u = (url or "").strip().lower()
    if not u:
        return ""
    if "instagram.com" in u:
        return "instagram"
    if "facebook.com" in u or "fb.com" in u or "fb.me" in u:
        return "facebook"
    return "web"


@dataclass
class Negocio:
    place_id: str
    nombre: str
    ciudad: str = ""
    tipo: str = ""
    pais: str = ""
    telefono: str = ""
    web: str = ""
    resenas: int = 0
    activo: bool = True
    link_maps: str = ""
    direccion: str = ""
    # --- datos del usuario (SQLite) ---
    estado: str = "nuevo"
    notas: str = ""
    guardado: bool = False
    guardado_en: str = ""
    visto_antes: bool = False
    primera_vez: str = ""
    ultima_actualizacion: str = ""

    @property
    def problemas(self) -> list[str]:
        return detectar_problemas(self)

    @property
    def problemas_texto(self) -> str:
        return " · ".join(self.problemas)


def detectar_problemas(n: Negocio) -> list[str]:
    """Señales objetivas a partir de los datos que reporta Google (sin azar)."""
    problemas: list[str] = []
    web = clasificar_web(n.web)
    if not web:
        problemas.append("Sin web")
    elif web == "instagram":
        problemas.append("Solo Instagram")
    elif web == "facebook":
        problemas.append("Solo Facebook")
    if not (n.telefono or "").strip():
        problemas.append("Sin teléfono")
    if n.resenas <= 0:
        problemas.append("Sin reseñas")
    elif n.resenas < POCAS_RESEÑAS_UMBRAL:
        problemas.append("Pocas reseñas")
    return problemas


def mensaje_para(n: Negocio) -> str:
    return MENSAJES.get(n.tipo, MENSAJE_GENERICO).format(nombre=n.nombre)


def _leer_google_api_key() -> str:
    key = os.environ.get("GOOGLE_MAPS_API_KEY", "").strip()
    if key:
        return key
    for carpeta in dict.fromkeys((_DIR_PROGRAMA, _BASE_DIR)):
        env_path = os.path.join(carpeta, ".env")
        if os.path.isfile(env_path):
            try:
                with open(env_path, encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line.startswith("GOOGLE_MAPS_API_KEY="):
                            valor = line.split("=", 1)[1].strip().strip('"').strip("'")
                            if valor:
                                return valor
            except OSError:
                continue
    return ""


# ---------------------------------------------------------------------------
# Palabras clave (palabras_clave.json)
# ---------------------------------------------------------------------------

def normalizar_terminos(terminos: Iterable[str]) -> list[str]:
    """Quita espacios sobrantes, vacíos y repetidos (sin distinguir mayúsculas)."""
    vistos: set[str] = set()
    limpios: list[str] = []
    for t in terminos:
        if not isinstance(t, str):
            continue
        t = re.sub(r"\s+", " ", t).strip()
        if t and t.lower() not in vistos:
            vistos.add(t.lower())
            limpios.append(t)
    return limpios


def cargar_palabras_clave(ruta: str) -> tuple[dict[str, list[str]], str, bool]:
    """Devuelve (palabras, aviso, archivo_invalido)."""
    palabras = {k: list(v) for k, v in PALABRAS_CLAVE_DEFECTO.items()}
    if not os.path.isfile(ruta):
        return palabras, "", False
    try:
        with open(ruta, encoding="utf-8-sig") as f:
            datos = json.load(f)
    except (OSError, ValueError) as exc:
        return palabras, f"No se pudo leer palabras_clave.json ({exc}); se usan los términos por defecto.", True
    if not isinstance(datos, dict):
        return palabras, "palabras_clave.json debe ser un objeto {\"Tipo\": [\"término\", ...]}; se usan los términos por defecto.", True
    for tipo, terminos in datos.items():
        if not isinstance(tipo, str) or not tipo.strip():
            continue
        if isinstance(terminos, str):
            terminos = [terminos]
        if not isinstance(terminos, list):
            continue
        limpios = normalizar_terminos(terminos)
        if limpios:
            palabras[tipo.strip()] = limpios
    return palabras, "", False


def guardar_palabras_clave(ruta: str, palabras: dict[str, list[str]], respaldar_invalido: bool = False) -> None:
    if respaldar_invalido and os.path.isfile(ruta):
        try:
            shutil.copyfile(ruta, ruta + ".bak")
        except OSError:
            pass
    tmp = ruta + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(palabras, f, ensure_ascii=False, indent=2)
    os.replace(tmp, ruta)


# ---------------------------------------------------------------------------
# Base de datos SQLite
# ---------------------------------------------------------------------------

_ESQUEMA = """
CREATE TABLE IF NOT EXISTS prospectos (
    place_id TEXT PRIMARY KEY,
    nombre TEXT NOT NULL,
    pais TEXT DEFAULT '',
    ciudad TEXT DEFAULT '',
    tipo TEXT DEFAULT '',
    telefono TEXT DEFAULT '',
    web TEXT DEFAULT '',
    resenas INTEGER DEFAULT 0,
    direccion TEXT DEFAULT '',
    link_maps TEXT DEFAULT '',
    problemas TEXT DEFAULT '',
    estado TEXT DEFAULT 'nuevo',
    notas TEXT DEFAULT '',
    guardado INTEGER DEFAULT 0,
    guardado_en TEXT DEFAULT '',
    ediciones TEXT DEFAULT '{}',
    primera_vez TEXT DEFAULT '',
    ultima_actualizacion TEXT DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_prospectos_guardado ON prospectos(guardado);
CREATE TABLE IF NOT EXISTS busquedas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    fecha TEXT, pais TEXT, ciudad TEXT, tipo TEXT,
    peticiones INTEGER DEFAULT 0,
    nuevos INTEGER DEFAULT 0,
    repetidos INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS uso_api (
    mes TEXT PRIMARY KEY,
    peticiones INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS terminos (
    clave TEXT PRIMARY KEY,
    ceros_seguidos INTEGER DEFAULT 0,
    nuevos_total INTEGER DEFAULT 0,
    ejecuciones INTEGER DEFAULT 0,
    ultima TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS meta (
    clave TEXT PRIMARY KEY,
    valor TEXT
);
"""


def _json_dict(texto: str) -> dict:
    try:
        d = json.loads(texto or "{}")
        return d if isinstance(d, dict) else {}
    except ValueError:
        return {}


class BaseDatos:
    """Acceso a SQLite. Un solo hilo a la vez (RLock): lo usan la interfaz y el hilo de búsqueda."""

    def __init__(self, ruta: str) -> None:
        self.ruta = ruta
        self._lock = threading.RLock()
        self._con = sqlite3.connect(ruta, check_same_thread=False, timeout=15)
        self._con.row_factory = sqlite3.Row
        with self._lock, self._con:
            self._con.executescript(_ESQUEMA)
        self._migrar_uso_json()

    def cerrar(self) -> None:
        with self._lock:
            try:
                self._con.close()
            except sqlite3.Error:
                pass

    # --- migración del antiguo api_usage.json (una sola vez) ---
    def _migrar_uso_json(self) -> None:
        with self._lock, self._con:
            ya = self._con.execute("SELECT valor FROM meta WHERE clave='uso_json_migrado'").fetchone()
            if ya or not os.path.isfile(RUTA_USO_JSON_ANTIGUO):
                if not ya:
                    self._con.execute("INSERT OR REPLACE INTO meta VALUES('uso_json_migrado','0')")
                return
            try:
                with open(RUTA_USO_JSON_ANTIGUO, encoding="utf-8") as f:
                    datos = json.load(f)
                if not isinstance(datos, dict):
                    raise ValueError("formato inesperado")
                mes = str(datos.get("mes", ""))
                del_mes = max(0, int(datos.get("peticiones_mes", 0)))
                total = max(0, int(datos.get("total_peticiones", 0)))
                if mes and del_mes:
                    self._con.execute(
                        "INSERT INTO uso_api(mes, peticiones) VALUES(?, ?) "
                        "ON CONFLICT(mes) DO UPDATE SET peticiones = peticiones + excluded.peticiones",
                        (mes, del_mes),
                    )
                anterior = total - del_mes
                if anterior > 0:
                    self._con.execute(
                        "INSERT INTO uso_api(mes, peticiones) VALUES('anterior', ?) "
                        "ON CONFLICT(mes) DO UPDATE SET peticiones = peticiones + excluded.peticiones",
                        (anterior,),
                    )
            except (OSError, ValueError, TypeError):
                pass
            self._con.execute("INSERT OR REPLACE INTO meta VALUES('uso_json_migrado','1')")

    # --- conversión fila <-> Negocio ---
    @staticmethod
    def _aplicar_usuario(n: Negocio, row: sqlite3.Row) -> None:
        n.estado = row["estado"] if row["estado"] in ESTADOS else "nuevo"
        n.notas = row["notas"] or ""
        n.guardado = bool(row["guardado"])
        n.guardado_en = row["guardado_en"] or ""
        n.primera_vez = row["primera_vez"] or ""
        n.ultima_actualizacion = row["ultima_actualizacion"] or ""
        for campo, valor in _json_dict(row["ediciones"]).items():
            if campo not in CAMPOS_EDITABLES:
                continue
            if campo == "resenas":
                try:
                    n.resenas = int(valor)
                except (TypeError, ValueError):
                    pass
            else:
                setattr(n, campo, str(valor))

    def _fila_a_negocio(self, row: sqlite3.Row) -> Negocio:
        n = Negocio(
            place_id=row["place_id"], nombre=row["nombre"], ciudad=row["ciudad"] or "",
            tipo=row["tipo"] or "", pais=row["pais"] or "", telefono=row["telefono"] or "",
            web=row["web"] or "", resenas=int(row["resenas"] or 0), link_maps=row["link_maps"] or "",
            direccion=row["direccion"] or "", visto_antes=True,
        )
        self._aplicar_usuario(n, row)
        return n

    # --- resultados de Google ---
    def upsert_google(self, negocios: list[Negocio], pais: str, ciudad: str, tipo: str) -> tuple[int, int]:
        """Guarda/actualiza lo que devolvió Google. Nunca toca estado, notas, guardado ni ediciones.
        Deja en cada Negocio los datos del usuario y la marca visto_antes. Devuelve (nuevos, repetidos)."""
        nuevos = repetidos = 0
        ahora = _ahora()
        with self._lock, self._con:
            for n in negocios:
                n.pais, n.ciudad, n.tipo = pais, ciudad, tipo
                crudo = (n.nombre, n.telefono, n.web, n.resenas, n.direccion, n.link_maps)
                row = self._con.execute(
                    "SELECT estado, notas, guardado, guardado_en, ediciones, primera_vez, ultima_actualizacion "
                    "FROM prospectos WHERE place_id = ?", (n.place_id,)
                ).fetchone()
                if row is None:
                    n.visto_antes = False
                    n.estado, n.notas, n.guardado, n.guardado_en = "nuevo", "", False, ""
                    n.primera_vez = ahora
                    n.ultima_actualizacion = ahora
                    self._con.execute(
                        "INSERT INTO prospectos(place_id, nombre, pais, ciudad, tipo, telefono, web, resenas, "
                        "direccion, link_maps, problemas, estado, primera_vez, ultima_actualizacion) "
                        "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        (n.place_id, *crudo[:1], pais, ciudad, tipo, *crudo[1:],
                         n.problemas_texto, "nuevo", ahora, ahora),
                    )
                    nuevos += 1
                else:
                    self._aplicar_usuario(n, row)
                    n.visto_antes = True
                    n.ultima_actualizacion = ahora
                    self._con.execute(
                        "UPDATE prospectos SET nombre=?, pais=?, ciudad=?, tipo=?, telefono=?, web=?, resenas=?, "
                        "direccion=?, link_maps=?, problemas=?, ultima_actualizacion=? WHERE place_id=?",
                        (crudo[0], pais, ciudad, tipo, crudo[1], crudo[2], crudo[3], crudo[4], crudo[5],
                         n.problemas_texto, ahora, n.place_id),
                    )
                    repetidos += 1
        return nuevos, repetidos

    def obtener(self, place_id: str) -> Optional[Negocio]:
        with self._lock:
            row = self._con.execute("SELECT * FROM prospectos WHERE place_id=?", (place_id,)).fetchone()
        return self._fila_a_negocio(row) if row else None

    def sincronizar(self, negocios: Iterable[Negocio]) -> None:
        """Refresca estado/notas/guardado/ediciones de objetos ya en pantalla (p. ej. tras cambiar de pestaña)."""
        lista = list(negocios)
        por_id = {n.place_id: n for n in lista}
        ids = list(por_id)
        with self._lock:
            for i in range(0, len(ids), 400):
                trozo = ids[i:i + 400]
                marcas = ",".join("?" * len(trozo))
                filas = self._con.execute(
                    f"SELECT * FROM prospectos WHERE place_id IN ({marcas})", trozo
                ).fetchall()
                for row in filas:
                    self._aplicar_usuario(por_id[row["place_id"]], row)

    # --- acciones del usuario ---
    def set_guardado(self, place_id: str, valor: bool) -> Optional[str]:
        guardado_en = _ahora() if valor else ""
        with self._lock, self._con:
            cur = self._con.execute(
                "UPDATE prospectos SET guardado=?, guardado_en=? WHERE place_id=?",
                (1 if valor else 0, guardado_en, place_id),
            )
        return guardado_en if cur.rowcount else None

    def set_estado(self, place_id: str, estado: str) -> bool:
        if estado not in ESTADOS:
            raise ValueError(f"Estado inválido: {estado}")
        with self._lock, self._con:
            cur = self._con.execute("UPDATE prospectos SET estado=? WHERE place_id=?", (estado, place_id))
        return bool(cur.rowcount)

    def set_notas(self, place_id: str, notas: str) -> bool:
        with self._lock, self._con:
            cur = self._con.execute("UPDATE prospectos SET notas=? WHERE place_id=?", (notas, place_id))
        return bool(cur.rowcount)

    def set_edicion(self, n: Negocio, campo: str, valor: object) -> bool:
        """Guarda una corrección manual (sobrevive a futuras actualizaciones desde Google)."""
        if campo not in CAMPOS_EDITABLES:
            raise ValueError(f"Campo no editable: {campo}")
        with self._lock, self._con:
            row = self._con.execute("SELECT ediciones FROM prospectos WHERE place_id=?", (n.place_id,)).fetchone()
            if row is None:
                return False
            ed = _json_dict(row["ediciones"])
            ed[campo] = valor
            self._con.execute(
                "UPDATE prospectos SET ediciones=?, problemas=? WHERE place_id=?",
                (json.dumps(ed, ensure_ascii=False), n.problemas_texto, n.place_id),
            )
        return True

    def listar_guardados(self) -> list[Negocio]:
        with self._lock:
            filas = self._con.execute(
                "SELECT * FROM prospectos WHERE guardado=1 ORDER BY guardado_en DESC, nombre COLLATE NOCASE"
            ).fetchall()
        return [self._fila_a_negocio(r) for r in filas]

    def contar(self) -> dict[str, int]:
        with self._lock:
            c = self._con
            return {
                "prospectos": c.execute("SELECT COUNT(*) FROM prospectos").fetchone()[0],
                "guardados": c.execute("SELECT COUNT(*) FROM prospectos WHERE guardado=1").fetchone()[0],
                "busquedas": c.execute("SELECT COUNT(*) FROM busquedas").fetchone()[0],
            }

    # --- historial de búsquedas ---
    def iniciar_busqueda(self, pais: str, ciudad: str, tipo: str) -> int:
        with self._lock, self._con:
            cur = self._con.execute(
                "INSERT INTO busquedas(fecha, pais, ciudad, tipo) VALUES(?,?,?,?)",
                (_ahora(), pais, ciudad, tipo),
            )
            return int(cur.lastrowid)

    def sumar_busqueda(self, busqueda_id: int, peticiones: int = 0, nuevos: int = 0, repetidos: int = 0) -> None:
        with self._lock, self._con:
            self._con.execute(
                "UPDATE busquedas SET peticiones=peticiones+?, nuevos=nuevos+?, repetidos=repetidos+? WHERE id=?",
                (peticiones, nuevos, repetidos, busqueda_id),
            )

    def peticiones_ultima_busqueda(self) -> int:
        with self._lock:
            row = self._con.execute("SELECT peticiones FROM busquedas ORDER BY id DESC LIMIT 1").fetchone()
        return int(row[0]) if row else 0

    # --- uso de la API ---
    @staticmethod
    def mes_actual() -> str:
        return time.strftime("%Y-%m")

    def registrar_peticion(self) -> None:
        with self._lock, self._con:
            self._con.execute(
                "INSERT INTO uso_api(mes, peticiones) VALUES(?, 1) "
                "ON CONFLICT(mes) DO UPDATE SET peticiones = peticiones + 1",
                (self.mes_actual(),),
            )

    def uso_mes(self) -> int:
        with self._lock:
            row = self._con.execute("SELECT peticiones FROM uso_api WHERE mes=?", (self.mes_actual(),)).fetchone()
        return int(row[0]) if row else 0

    def uso_total(self) -> int:
        with self._lock:
            return int(self._con.execute("SELECT COALESCE(SUM(peticiones),0) FROM uso_api").fetchone()[0])

    # --- rendimiento de cada término ---
    def ceros_seguidos(self, clave: str) -> int:
        with self._lock:
            row = self._con.execute("SELECT ceros_seguidos FROM terminos WHERE clave=?", (clave,)).fetchone()
        return int(row[0]) if row else 0

    def registrar_termino(self, clave: str, nuevos: int) -> None:
        with self._lock, self._con:
            self._con.execute(
                "INSERT INTO terminos(clave, ceros_seguidos, nuevos_total, ejecuciones, ultima) "
                "VALUES(?, ?, ?, 1, ?) "
                "ON CONFLICT(clave) DO UPDATE SET "
                "ceros_seguidos = CASE WHEN ? = 0 THEN ceros_seguidos + 1 ELSE 0 END, "
                "nuevos_total = nuevos_total + ?, ejecuciones = ejecuciones + 1, ultima = ?",
                (clave, 1 if nuevos == 0 else 0, nuevos, _ahora(), nuevos, nuevos, _ahora()),
            )

    def terminos_omitidos(self) -> int:
        with self._lock:
            return int(self._con.execute(
                "SELECT COUNT(*) FROM terminos WHERE ceros_seguidos >= ?", (UMBRAL_OMITIR_TERMINO,)
            ).fetchone()[0])

    def reiniciar_terminos(self) -> None:
        with self._lock, self._con:
            self._con.execute("DELETE FROM terminos")


class ApiUsageTracker:
    """Registra peticiones a Places API (Text Search); los datos viven en SQLite."""

    def __init__(self, db: BaseDatos) -> None:
        self._db = db
        self.peticiones_sesion = 0

    def registrar_peticion(self) -> None:
        self.peticiones_sesion += 1
        self._db.registrar_peticion()

    def resumen_texto(self) -> str:
        mes = self._db.mes_actual()
        peticiones_mes = self._db.uso_mes()
        total = self._db.uso_total()
        costo_mes = peticiones_mes * COSTO_ESTIMADO_POR_PETICION_USD
        credito_restante = max(0.0, CREDITO_MENSUAL_USD - costo_mes)
        uso_pct = min(100.0, (costo_mes / CREDITO_MENSUAL_USD) * 100) if CREDITO_MENSUAL_USD else 0
        cont = self._db.contar()
        return (
            "Uso de Google Places API (Text Search)\n"
            "────────────────────────────────────\n\n"
            f"Sesión actual:              {self.peticiones_sesion} peticiones\n"
            f"Última búsqueda:            {self._db.peticiones_ultima_busqueda()} peticiones\n"
            f"Este mes ({mes}):           {peticiones_mes} peticiones\n"
            f"Total histórico (app):      {total} peticiones\n\n"
            "Estimado de costo\n"
            "────────────────────────────────────\n"
            f"Costo estimado este mes:    ${costo_mes:.2f} USD\n"
            f"Crédito mensual referencia: ${CREDITO_MENSUAL_USD:.0f} USD\n"
            f"Crédito restante estimado:  ${credito_restante:.2f} USD\n"
            f"Uso estimado del crédito:   {uso_pct:.1f}%\n\n"
            "Base de datos local\n"
            "────────────────────────────────────\n"
            f"Prospectos vistos:          {cont['prospectos']}\n"
            f"Prospectos guardados:       {cont['guardados']}\n"
            f"Búsquedas realizadas:       {cont['busquedas']}\n"
            f"Términos omitidos (0 nuevos): {self._db.terminos_omitidos()}\n\n"
            "Notas\n"
            "────────────────────────────────────\n"
            "• Cada consulta (o página de resultados) a Google Maps cuenta como 1 petición.\n"
            "• El costo real depende de tu plan en Google Cloud.\n"
            "• Google no expone la cuota exacta restante con solo la API key.\n"
            "• Revisa el uso real en:\n"
            "  console.cloud.google.com → APIs y servicios → Panel"
        )


# ---------------------------------------------------------------------------
# Google Maps Places API
# ---------------------------------------------------------------------------

class BusquedaCancelada(Exception):
    """El usuario pulsó «Cancelar»."""


class TokenPaginaInvalido(RuntimeError):
    """Google rechazó el token de la página siguiente (expirado o aún no activo)."""


def _mensaje_error_api(codigo: int, detalle: str) -> str:
    texto = detalle.strip()
    estado = ""
    try:
        err = json.loads(detalle).get("error", {})
        texto = err.get("message", texto) or texto
        estado = err.get("status", "")
    except (ValueError, AttributeError):
        pass
    pista = ""
    if codigo in (401, 403):
        pista = ("\nRevisa que la API key sea válida, que «Places API (New)» esté habilitada "
                 "y que la facturación esté activa en Google Cloud.")
    elif codigo == 429:
        pista = "\nSe superó la cuota de Google; espera un momento o revisa tus límites."
    return f"Google Places API error ({codigo}{' ' + estado if estado else ''}): {texto[:400]}{pista}"


class ClienteGooglePlaces:
    """Cliente mínimo de Text Search (New): reintentos, cancelación y conteo de uso."""

    URL = "https://places.googleapis.com/v1/places:searchText"
    # nextPageToken DEBE ir en el field mask; si no, Google nunca devuelve la página 2 y 3.
    FIELD_MASK = (
        "places.id,places.displayName,places.formattedAddress,"
        "places.internationalPhoneNumber,places.nationalPhoneNumber,"
        "places.websiteUri,places.userRatingCount,places.businessStatus,"
        "places.googleMapsUri,nextPageToken"
    )

    def __init__(self, api_key: str, uso: ApiUsageTracker, cancelar: Optional[threading.Event] = None) -> None:
        if not api_key:
            raise ValueError(
                "Falta la API key de Google Maps.\n"
                "Crea un archivo .env con GOOGLE_MAPS_API_KEY=tu_clave\n"
                "o define la variable de entorno GOOGLE_MAPS_API_KEY.\n"
                "Activa Places API (New) en Google Cloud Console."
            )
        self._api_key = api_key
        self._uso = uso
        self._cancelar = cancelar or threading.Event()

    def cancelado(self) -> bool:
        return self._cancelar.is_set()

    def esperar(self, segundos: float) -> None:
        """Pausa que se interrumpe al cancelar."""
        if segundos <= 0:
            if self._cancelar.is_set():
                raise BusquedaCancelada()
            return
        if self._cancelar.wait(segundos):
            raise BusquedaCancelada()

    def post(self, body: dict) -> dict:
        datos = json.dumps(body).encode("utf-8")
        ultimo_error = "sin respuesta"
        invalid_token = 0
        intento = 0
        while True:
            if self._cancelar.is_set():
                raise BusquedaCancelada()
            req = urllib.request.Request(
                self.URL, data=datos, method="POST",
                headers={
                    "Content-Type": "application/json",
                    "X-Goog-Api-Key": self._api_key,
                    "X-Goog-FieldMask": self.FIELD_MASK,
                },
            )
            espera = ESPERA_REINTENTO_BASE * (2 ** intento)
            try:
                with urllib.request.urlopen(req, timeout=30) as resp:
                    respuesta = json.loads(resp.read().decode("utf-8"))
                self._uso.registrar_peticion()
                return respuesta
            except urllib.error.HTTPError as exc:
                detalle = exc.read().decode("utf-8", errors="replace")
                if exc.code == 400 and body.get("pageToken") and invalid_token < 3:
                    invalid_token += 1          # el token tarda un poco en activarse
                    ultimo_error = detalle
                    self.esperar(ESPERA_TOKEN_REINTENTO * invalid_token)
                    continue
                if exc.code == 400 and body.get("pageToken"):
                    raise TokenPaginaInvalido(_mensaje_error_api(exc.code, detalle)) from exc
                if exc.code in (429, 500, 502, 503, 504) and intento < MAX_REINTENTOS:
                    ultimo_error = detalle
                    retry_after = exc.headers.get("Retry-After") if exc.headers else None
                    if retry_after and str(retry_after).isdigit():
                        espera = max(espera, min(60.0, float(retry_after)))
                    intento += 1
                    self.esperar(espera)
                    continue
                raise RuntimeError(_mensaje_error_api(exc.code, detalle)) from exc
            except (urllib.error.URLError, OSError, ValueError) as exc:
                ultimo_error = str(exc)
                if intento < MAX_REINTENTOS:
                    intento += 1
                    self.esperar(espera)
                    continue
                raise RuntimeError(f"Error de red al consultar Google Maps: {ultimo_error[:300]}") from exc


def place_a_negocio(place: dict, pais: str, ciudad: str, tipo: str) -> Optional[Negocio]:
    place_id = str(place.get("id") or "").strip()
    nombre = str((place.get("displayName") or {}).get("text") or "").strip()
    if not place_id or not nombre:
        return None
    telefono = str(place.get("internationalPhoneNumber") or place.get("nationalPhoneNumber") or "").strip()
    web = str(place.get("websiteUri") or "").strip()
    try:
        resenas = int(place.get("userRatingCount") or 0)
    except (TypeError, ValueError):
        resenas = 0
    activo = place.get("businessStatus", "") != "CLOSED_PERMANENTLY"
    link = str(place.get("googleMapsUri") or "").strip() or f"https://www.google.com/maps/place/?q=place_id:{place_id}"
    return Negocio(
        place_id=place_id, nombre=nombre, ciudad=ciudad, tipo=tipo, pais=pais,
        telefono=telefono, web=web, resenas=resenas, activo=activo, link_maps=link,
        direccion=str(place.get("formattedAddress") or "").strip(),
    )


# ---------------------------------------------------------------------------
# Sesión de búsqueda (caché en RAM + paginación sin repetir)
# ---------------------------------------------------------------------------

class SesionBusqueda:
    """Una búsqueda (país, ciudad, tipo). Todo lo que devuelve Google se guarda en `raw` (RAM):
    paginar, cambiar el mínimo de reseñas o exportar trabajan sobre esos datos sin volver a consultar.
    Solo se hacen peticiones nuevas cuando la caché no alcanza para llenar una página."""

    def __init__(
        self,
        cliente: ClienteGooglePlaces,
        db: BaseDatos,
        pais: str,
        ciudad: str,
        tipo: str,
        terminos: list[str],
        min_resenas: int,
        limite: int,
        progreso: Optional[Callable[[str], None]] = None,
    ) -> None:
        self.cliente = cliente
        self.db = db
        self.pais, self.ciudad, self.tipo = pais, ciudad, tipo
        self.terminos = normalizar_terminos(terminos) or [tipo]
        self.min_resenas = max(0, min_resenas)
        self.limite = max(1, limite)
        self.progreso = progreso
        self.region = _PAIS_REGION.get(pais)
        self.raw: dict[str, Negocio] = {}
        self.vistos: set[str] = set()
        self.usados: set[str] = set()
        self.paginas: dict[int, list[Negocio]] = {}
        self.peticiones = 0
        self.agotado = False
        self.error: Optional[str] = None
        self.cancelado = False
        self.omitidos: list[str] = []
        # estado del recorrido (un término = hasta 3 páginas de Google)
        self._i = 0
        self._pn = 0
        self._tok: Optional[str] = None
        self._nuevos_termino = 0
        self.busqueda_id = db.iniciar_busqueda(pais, ciudad, tipo)

    # --- utilidades ---
    def _clave(self, termino: str) -> str:
        return f"{self.pais}|{self.ciudad}|{self.tipo}|{termino.lower()}"

    def _consulta(self, termino: str) -> str:
        return f"{termino} en {self.ciudad}, {self.pais}"

    def _avisar(self, texto: str) -> None:
        if self.progreso:
            try:
                self.progreso(texto)
            except Exception:
                pass

    def _candidatos(self) -> list[Negocio]:
        return [n for n in self.raw.values() if n.place_id not in self.usados and n.activo]

    def _n_estrictos(self) -> int:
        return sum(1 for n in self._candidatos() if n.resenas >= self.min_resenas)

    def tope_peticiones_pagina(self) -> int:
        if self.limite >= PAGINA_TODOS:
            return TOPE_PETICIONES_MAX
        return min(TOPE_PETICIONES_MAX, max(12, self.limite // 5))

    def estimar_peticiones_restantes(self) -> int:
        """Cota superior de las peticiones que faltan para agotar los términos."""
        if self.agotado:
            return 0
        total = 0
        for idx in range(self._i, len(self.terminos)):
            if idx == self._i and self._pn > 0:
                total += max(0, MAX_PAGINAS_POR_CONSULTA - self._pn)
            elif self.db.ceros_seguidos(self._clave(self.terminos[idx])) < UMBRAL_OMITIR_TERMINO:
                total += MAX_PAGINAS_POR_CONSULTA
        return total

    def hay_mas(self) -> bool:
        return (not self.agotado) or bool(self._candidatos())

    # --- una petición a Google ---
    def _ingerir(self, places: list) -> int:
        nuevos: list[Negocio] = []
        for place in places if isinstance(places, list) else []:
            if not isinstance(place, dict):
                continue
            pid = str(place.get("id") or "").strip()
            if not pid or pid in self.vistos:
                continue
            self.vistos.add(pid)
            negocio = place_a_negocio(place, self.pais, self.ciudad, self.tipo)
            if negocio is None or not negocio.activo:
                continue
            nuevos.append(negocio)
        if nuevos:
            n_nuevos, n_repetidos = self.db.upsert_google(nuevos, self.pais, self.ciudad, self.tipo)
            self.db.sumar_busqueda(self.busqueda_id, nuevos=n_nuevos, repetidos=n_repetidos)
            for negocio in nuevos:
                self.raw[negocio.place_id] = negocio
        return len(nuevos)

    def _cerrar_termino(self, registrar: bool) -> None:
        if registrar:
            self.db.registrar_termino(self._clave(self.terminos[self._i]), self._nuevos_termino)
        self._i += 1
        self._pn = 0
        self._tok = None
        self._nuevos_termino = 0

    def _avanzar(self) -> int:
        """Hace como máximo UNA petición. Devuelve 1 si la hizo, 0 si no (omitió términos o terminó)."""
        while self._i < len(self.terminos):
            termino = self.terminos[self._i]
            if self._pn == 0 and self._tok is None:
                if self.db.ceros_seguidos(self._clave(termino)) >= UMBRAL_OMITIR_TERMINO:
                    self.omitidos.append(termino)
                    self._i += 1
                    continue
            body: dict = {
                "textQuery": self._consulta(termino),
                "languageCode": "es",
                "pageSize": 20,
            }
            if self.region:
                body["regionCode"] = self.region
            if self._tok:
                body["pageToken"] = self._tok
                self.cliente.esperar(ESPERA_TOKEN_PAGINA)
            try:
                data = self.cliente.post(body)
            except TokenPaginaInvalido:
                self._cerrar_termino(registrar=False)   # abandona el resto de este término
                return 0
            if not isinstance(data, dict):
                data = {}
            self.peticiones += 1
            self.db.sumar_busqueda(self.busqueda_id, peticiones=1)
            self._pn += 1
            self._nuevos_termino += self._ingerir(data.get("places") or [])
            tok = data.get("nextPageToken")
            self._tok = tok if isinstance(tok, str) and tok else None
            self._avisar(
                f"«{termino}» · {self.peticiones} peticiones · {len(self.raw)} negocios en caché"
            )
            if not self._tok or self._pn >= MAX_PAGINAS_POR_CONSULTA:
                self._cerrar_termino(registrar=True)
            return 1
        self.agotado = True
        return 0

    # --- páginas ---
    def construir_pagina(self, permitir_api: bool = True) -> list[Negocio]:
        """Arma la siguiente página: primero los que cumplen el mínimo de reseñas; si no alcanza,
        completa con los de más reseñas que ya estén en caché (se relaja el mínimo)."""
        self.error = None
        self.cancelado = False
        if permitir_api:
            tope = self.tope_peticiones_pagina()
            llamadas = 0
            sin_estrictos = 0
            try:
                while not self.agotado and llamadas < tope:
                    estrictos = self._n_estrictos()
                    if estrictos >= self.limite:
                        break
                    if sin_estrictos >= TRAS_SIN_ESTRICTOS and len(self._candidatos()) >= self.limite:
                        break               # el mínimo es muy exigente: se completa con los de más reseñas
                    self._avanzar()
                    llamadas += 1
                    sin_estrictos = 0 if self._n_estrictos() > estrictos else sin_estrictos + 1
            except BusquedaCancelada:
                self.cancelado = True
            except RuntimeError as exc:
                self.error = str(exc)
            except Exception as exc:        # noqa: BLE001 — no perder lo que ya está en caché
                self.error = f"Error inesperado: {exc}"
        candidatos = self._candidatos()
        pagina = [n for n in candidatos if n.resenas >= self.min_resenas][: self.limite]
        if len(pagina) < self.limite:
            ids = {n.place_id for n in pagina}
            resto = sorted((n for n in candidatos if n.place_id not in ids), key=lambda n: -n.resenas)
            pagina += resto[: self.limite - len(pagina)]
        if pagina:
            self.paginas[len(self.paginas) + 1] = pagina
            self.usados.update(n.place_id for n in pagina)
        return pagina

    def reconfigurar(self, min_resenas: int, limite: int) -> None:
        """Cambia mínimo de reseñas / tamaño de página y reparte de nuevo lo que ya está en caché."""
        self.min_resenas = max(0, min_resenas)
        self.limite = max(1, limite)
        self.paginas.clear()
        self.usados.clear()

    def cargar_todo(self, tope: int = TOPE_CARGAR_TODO) -> int:
        """Pide todo lo que falte (hasta `tope` peticiones) y reparte el resto en páginas desde la caché."""
        error, cancelado = None, False
        hechas = 0
        try:
            while not self.agotado and hechas < tope:
                self._avanzar()
                hechas += 1
        except BusquedaCancelada:
            cancelado = True
        except RuntimeError as exc:
            error = str(exc)
        except Exception as exc:            # noqa: BLE001
            error = f"Error inesperado: {exc}"
        while self.construir_pagina(permitir_api=False):
            pass
        self.error, self.cancelado = error, cancelado
        return hechas

    def todos_los_negocios(self) -> list[Negocio]:
        salida: list[Negocio] = []
        for numero in sorted(self.paginas):
            salida.extend(self.paginas[numero])
        return salida


# ---------------------------------------------------------------------------
# Exportación (Excel / CSV)
# ---------------------------------------------------------------------------

def _sanitizar_nombre(parte: str) -> str:
    invalidos = '<>:"/\\|?*'
    s = (parte or "").strip().replace(" ", "-")
    for c in invalidos:
        s = s.replace(c, "")
    return s or "sin-nombre"


def _limpio(valor: object) -> object:
    """Quita caracteres que Excel no admite y recorta textos más largos que una celda."""
    if isinstance(valor, str):
        return ILLEGAL_CHARACTERS_RE.sub("", valor)[:32_000]
    return valor


def filas_exportacion(negocios: Iterable[Negocio]) -> list[tuple]:
    return [
        tuple(_limpio(v) for v in (
            n.nombre, n.ciudad, n.telefono, n.resenas, n.web, n.problemas_texto,
            n.estado.capitalize(), n.notas, n.link_maps, n.direccion))
        for n in negocios
    ]


def _seguro_csv(valor: object) -> object:
    """Evita que Excel interprete un texto como fórmula al abrir el CSV."""
    if not isinstance(valor, str) or not valor:
        return valor
    if valor[0] in "=@\t\r":
        return "'" + valor
    if valor[0] in "+-" and not re.fullmatch(r"[+\-][\d\s().\-]+", valor):
        return "'" + valor
    return valor


def ruta_unica(ruta: str) -> str:
    if not os.path.exists(ruta):
        return ruta
    base, ext = os.path.splitext(ruta)
    k = 2
    while os.path.exists(f"{base}-{k}{ext}"):
        k += 1
    return f"{base}-{k}{ext}"


def nombre_exportacion(partes: list[str], extension: str) -> str:
    ahora = datetime.now()
    sufijo = f"{ahora.strftime('%d-%m-%Y')}-{ahora.strftime('%H-%M')}"
    return "-".join(_sanitizar_nombre(p) for p in partes[:-1] + [sufijo, partes[-1]]) + f".{extension}"


def exportar_excel(ruta: str, negocios: list[Negocio]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"
    filas = filas_exportacion(negocios)
    num_cols = len(COLUMNAS_EXPORT)

    borde_fino = Side(style="thin", color="D0D0D0")
    borde = Border(left=borde_fino, right=borde_fino, top=borde_fino, bottom=borde_fino)
    relleno_cabecera = PatternFill(fill_type="solid", fgColor="0F6CBD")
    fuente_cabecera = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    relleno_fila_alt = PatternFill(fill_type="solid", fgColor="EEF4FA")
    alin_cabecera = Alignment(horizontal="center", vertical="center", wrap_text=True)
    alin_datos = Alignment(vertical="center", wrap_text=True)
    fuente_link = Font(color="0563C1", underline="single", name="Calibri", size=11)
    columnas_link = {COLUMNAS_EXPORT.index("Web / Instagram"), COLUMNAS_EXPORT.index("Link Google Maps")}

    for col, titulo in enumerate(COLUMNAS_EXPORT, start=1):
        celda = ws.cell(row=1, column=col, value=titulo)
        celda.fill, celda.font, celda.alignment, celda.border = relleno_cabecera, fuente_cabecera, alin_cabecera, borde

    enlaces = 0
    for fila_idx, valores in enumerate(filas, start=2):
        zebra = fila_idx % 2 == 0
        for col, valor in enumerate(valores, start=1):
            celda = ws.cell(row=fila_idx, column=col, value=valor)
            if isinstance(valor, str) and valor.startswith("="):
                celda.data_type = "s"       # texto, nunca fórmula
            celda.border = borde
            celda.alignment = alin_datos
            if zebra:
                celda.fill = relleno_fila_alt
            if ((col - 1) in columnas_link and isinstance(valor, str) and enlaces < MAX_ENLACES_EXCEL
                    and valor.lower().startswith(("http://", "https://")) and len(valor) <= 2000):
                celda.hyperlink = valor
                celda.font = fuente_link
                enlaces += 1

    for col in range(1, num_cols + 1):
        ancho = len(str(COLUMNAS_EXPORT[col - 1]))
        for fila_idx in range(2, ws.max_row + 1):
            valor = ws.cell(row=fila_idx, column=col).value
            if valor is not None:
                ancho = max(ancho, len(str(valor)))
        ws.column_dimensions[get_column_letter(col)].width = min(ancho + 3, 55)

    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"
    if filas:
        ws.auto_filter.ref = f"A1:{get_column_letter(num_cols)}{ws.max_row}"
    wb.save(ruta)


def exportar_csv(ruta: str, negocios: list[Negocio]) -> None:
    with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(COLUMNAS_EXPORT)
        for fila in filas_exportacion(negocios):
            w.writerow([_seguro_csv(v) for v in fila])


def abrir_ruta(ruta: str) -> None:
    """Abre una carpeta o archivo con el programa del sistema."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(ruta)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", ruta])
        else:
            subprocess.Popen(["xdg-open", ruta])
    except Exception:
        pass


def abrir_url(url: str) -> bool:
    url = (url or "").strip()
    if url.lower().startswith(("http://", "https://")):
        webbrowser.open(url)
        return True
    return False


# ---------------------------------------------------------------------------
# Importar informes (Excel / CSV)
# ---------------------------------------------------------------------------

MAX_FILAS_INFORME = 50_000            # filas de datos que se muestran como máximo
MAX_FILAS_BRUTAS_INFORME = 300_000    # filas que se recorren como máximo (por si hay millones vacías)


def _celda_a_texto(v: object) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "Sí" if v else "No"
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        return str(int(v))
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M") if (v.hour or v.minute or v.second) else v.strftime("%Y-%m-%d")
    texto = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(v))
    return re.sub(r"[\r\n]+", " ", texto)


def _aviso_limite() -> str:
    return f"Se muestran las primeras {MAX_FILAS_INFORME:,} filas; el archivo tiene más.".replace(",", ".")


def _leer_xlsx(ruta: str) -> tuple[list[list[str]], str]:
    from openpyxl import load_workbook
    wb = load_workbook(ruta, read_only=True, data_only=True)
    try:
        hojas = list(wb.worksheets)
        if not hojas:
            raise ValueError("El libro de Excel no tiene hojas de datos.")
        ws = next((h for h in hojas if h.title == "Leads"), hojas[0])
        try:
            ws.reset_dimensions()       # algunos archivos declaran un tamaño incorrecto
        except Exception:               # noqa: BLE001
            pass
        filas: list[list[str]] = []
        aviso = ""
        brutas = 0
        for fila in ws.iter_rows(values_only=True):
            brutas += 1
            if brutas > MAX_FILAS_BRUTAS_INFORME:
                aviso = _aviso_limite()
                break
            vals = [_celda_a_texto(v) for v in fila]
            if not any(v.strip() for v in vals):
                continue
            if len(filas) >= MAX_FILAS_INFORME + 1:      # +1 por la fila de encabezados
                aviso = _aviso_limite()
                break
            filas.append(vals)
        if len(hojas) > 1:
            extra = f"Hoja «{ws.title}» (el archivo tiene {len(hojas)} hojas)."
            aviso = f"{aviso} {extra}".strip()
        return filas, aviso
    finally:
        wb.close()


def _leer_csv(ruta: str) -> tuple[list[list[str]], str]:
    with open(ruta, "rb") as f:
        muestra = f.read(65536)
    if muestra.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        codificacion = "utf-16"
    else:
        try:
            codecs.getincrementaldecoder("utf-8-sig")().decode(muestra, final=False)
            codificacion = "utf-8-sig"
        except UnicodeDecodeError:
            codificacion = "cp1252"
    texto = muestra.decode(codificacion, errors="replace")
    primera = next((l for l in texto.splitlines() if l.strip()), "")
    delimitador = max(",;\t|", key=primera.count) if any(d in primera for d in ",;\t|") else ","
    try:
        csv.field_size_limit(2 ** 31 - 1)
    except (OverflowError, ValueError):
        pass
    filas: list[list[str]] = []
    aviso = ""
    with open(ruta, "r", encoding=codificacion, errors="replace", newline="") as f:
        for fila in csv.reader((linea.replace("\x00", "") for linea in f), delimiter=delimitador):
            vals = [_celda_a_texto(v) for v in fila]
            if not any(v.strip() for v in vals):
                continue
            if len(filas) >= MAX_FILAS_INFORME + 1:
                aviso = _aviso_limite()
                break
            filas.append(vals)
    return filas, aviso


def leer_informe(ruta: str) -> tuple[list[str], list[list[str]], str]:
    """Lee un informe .xlsx/.xlsm/.csv. Devuelve (encabezados, filas, aviso). Lanza ValueError con un mensaje claro."""
    if not os.path.isfile(ruta):
        raise ValueError("El archivo no existe.")
    if os.path.getsize(ruta) == 0:
        raise ValueError("El archivo está vacío.")
    ext = os.path.splitext(ruta)[1].lower()
    if ext == ".xls":
        raise ValueError("El formato .xls (Excel antiguo) no se puede abrir. Guárdalo como .xlsx o .csv desde Excel.")
    if ext not in (".xlsx", ".xlsm", ".csv", ".tsv", ".txt"):
        raise ValueError(f"Formato no admitido ({ext or 'sin extensión'}). Usa un informe .xlsx o .csv.")
    try:
        filas, aviso = _leer_xlsx(ruta) if ext in (".xlsx", ".xlsm") else _leer_csv(ruta)
    except ValueError:
        raise
    except PermissionError as exc:
        raise ValueError("No se puede leer el archivo: está bloqueado o no tienes permiso (¿abierto en otro programa?).") from exc
    except OSError as exc:
        raise ValueError(f"No se pudo leer el archivo: {exc}") from exc
    except Exception as exc:            # noqa: BLE001 — zip dañado, libro protegido con contraseña, CSV corrupto…
        raise ValueError(
            f"No se pudo leer el archivo; puede estar dañado, protegido con contraseña o no ser un informe válido ({type(exc).__name__})."
        ) from exc
    if not filas:
        raise ValueError("El archivo no contiene filas con datos.")
    ncols = max(max((i for i, v in enumerate(f) if v.strip()), default=-1) + 1 for f in filas)
    cabecera = filas[0]
    encabezados: list[str] = []
    usados: set[str] = set()
    for i in range(ncols):
        h = cabecera[i].strip() if i < len(cabecera) else ""
        if not h:
            h = f"Columna {i + 1}"
        base, k = h, 2
        while h.lower() in usados:
            h = f"{base} ({k})"
            k += 1
        usados.add(h.lower())
        encabezados.append(h)
    datos = [(f + [""] * ncols)[:ncols] for f in filas[1:]]
    return encabezados, datos, aviso


# ---------------------------------------------------------------------------
# Interfaz gráfica — widgets
# ---------------------------------------------------------------------------

COLORES = {
    "primario": "#0F6CBD", "primario_osc": "#0B4F8C", "primario_claro": "#DCEBFA",
    "fondo": "#F3F5F8", "tarjeta": "#FFFFFF", "texto": "#1F2937", "suave": "#6B7280",
    "ok": "#1A7F37", "error": "#C0392B", "aviso": "#B7791F", "borde": "#D5DCE5",
    "zebra": "#F6F9FC", "sel": "#CFE3F7",
    "contactado": "#FFF4D6", "respondió": "#E3F4E8", "descartado": "#EDEEF0",
}


@dataclass
class Columna:
    id: str
    titulo: str
    ancho: int
    anchor: str = "w"
    editable: bool = False
    numerica: bool = False


class TablaDatos(ttk.Frame):
    """Treeview con orden por columna, filas cebra, color por estado, copiar celda y edición en la celda."""

    def __init__(
        self,
        parent: tk.Widget,
        columnas: list[Columna],
        *,
        on_estrella: Optional[Callable[[str], None]] = None,
        on_editar: Optional[Callable[[str, str, str], None]] = None,
        on_contexto: Optional[Callable[[tk.Event], None]] = None,
        on_doble_estado: Optional[Callable[[str, int, int], None]] = None,
        on_seleccion: Optional[Callable[[], None]] = None,
        on_mensaje: Optional[Callable[[str], None]] = None,
    ) -> None:
        super().__init__(parent)
        self.columnas = columnas
        self._on_estrella, self._on_editar = on_estrella, on_editar
        self._on_contexto, self._on_doble_estado = on_contexto, on_doble_estado
        self._on_seleccion, self._on_mensaje = on_seleccion, on_mensaje
        self._filas: dict[str, tuple] = {}
        self._estado: dict[str, str] = {}
        self._orden: list[str] = []
        self._sort_col: Optional[str] = None
        self._sort_rev = False
        self._editor: Optional[tuple[tk.Entry, str, int, str]] = None
        self._cerrando = False
        self.celda_activa: Optional[tuple[str, int]] = None

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(
            self, columns=[c.id for c in columnas], show="headings",
            selectmode="extended", style="Datos.Treeview",
        )
        for c in columnas:
            self.tree.heading(c.id, text=c.titulo, anchor=c.anchor, command=lambda cid=c.id: self.ordenar(cid))
            self.tree.column(c.id, width=c.ancho, minwidth=40, anchor=c.anchor, stretch=False)
        self.tree.tag_configure("par", background=COLORES["tarjeta"])
        self.tree.tag_configure("impar", background=COLORES["zebra"])
        self.tree.tag_configure("estado_contactado", background=COLORES["contactado"])
        self.tree.tag_configure("estado_respondió", background=COLORES["respondió"])
        self.tree.tag_configure("estado_descartado", background=COLORES["descartado"], foreground=COLORES["suave"])

        self._vsb = ttk.Scrollbar(self, orient="vertical", command=self._yview)
        self._hsb = ttk.Scrollbar(self, orient="horizontal", command=self._xview)
        self.tree.configure(yscrollcommand=self._vsb.set, xscrollcommand=self._hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        self._vsb.grid(row=0, column=1, sticky="ns")
        self._hsb.grid(row=1, column=0, sticky="ew")

        self.tree.bind("<Button-1>", self._click, add="+")
        self.tree.bind("<Double-1>", self._doble)
        self.tree.bind("<Button-3>", self._contexto)
        self.tree.bind("<Button-2>", self._contexto)
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._on_seleccion() if self._on_seleccion else None)
        self.tree.bind("<Delete>", self._borrar_celda)
        self.tree.bind("<Control-c>", self._copiar)
        self.tree.bind("<Control-C>", self._copiar)
        self.tree.bind("<Command-c>", self._copiar)
        for ev in ("<MouseWheel>", "<Button-4>", "<Button-5>", "<Configure>"):
            self.tree.bind(ev, lambda _e: self.cerrar_editor(True), add="+")

    # --- datos ---
    def set_filas(self, filas: list[tuple[str, tuple, str]]) -> None:
        self.cerrar_editor(False)
        self._filas = {iid: tuple(vals) for iid, vals, _e in filas}
        self._estado = {iid: est for iid, _v, est in filas}
        self._orden = [iid for iid, _v, _e in filas]
        if self._sort_col:
            self._aplicar_orden()
        self._pintar(conservar_seleccion=False)

    def limpiar(self) -> None:
        self.set_filas([])

    def actualizar_fila(self, iid: str, valores: tuple, estado: str) -> None:
        if iid not in self._filas:
            return
        self._filas[iid] = tuple(valores)
        self._estado[iid] = estado
        self.tree.item(iid, values=self._filas[iid], tags=self._tags(iid, self._orden.index(iid)))

    def iids(self) -> list[str]:
        return list(self._orden)

    def seleccion(self) -> list[str]:
        return [i for i in self.tree.selection() if i in self._filas]

    def valor_celda(self, iid: str, ci: int) -> str:
        fila = self._filas.get(iid)
        return "" if fila is None or ci >= len(fila) else str(fila[ci])

    def ver(self, iid: str) -> None:
        if iid in self._filas:
            self.tree.see(iid)

    def _tags(self, iid: str, pos: int) -> tuple[str, ...]:
        estado = self._estado.get(iid, "nuevo")
        if estado and estado != "nuevo":
            return (f"estado_{estado}",)
        return ("par" if pos % 2 == 0 else "impar",)

    def _pintar(self, conservar_seleccion: bool = True) -> None:
        sel = [i for i in self.tree.selection() if i in self._filas] if conservar_seleccion else []
        self.tree.delete(*self.tree.get_children())
        for pos, iid in enumerate(self._orden):
            self.tree.insert("", "end", iid=iid, values=self._filas[iid], tags=self._tags(iid, pos))
        if sel:
            self.tree.selection_set(sel)
        self.celda_activa = None
        self._rotular_cabeceras()

    # --- orden ---
    def ordenar(self, col_id: str) -> None:
        if self._sort_col == col_id:
            self._sort_rev = not self._sort_rev
        else:
            self._sort_col, self._sort_rev = col_id, False
        self.cerrar_editor(False)
        self._aplicar_orden()
        self._pintar()

    def _aplicar_orden(self) -> None:
        ci = next(i for i, c in enumerate(self.columnas) if c.id == self._sort_col)
        numerica = self.columnas[ci].numerica

        def clave(iid: str):
            v = self.valor_celda(iid, ci)
            if numerica:
                try:
                    return int(v)
                except ValueError:
                    return -1
            return v.lower()

        con_valor = [i for i in self._orden if self.valor_celda(i, ci) != ""]
        vacios = [i for i in self._orden if self.valor_celda(i, ci) == ""]
        con_valor.sort(key=clave, reverse=self._sort_rev)
        self._orden = con_valor + vacios

    def _rotular_cabeceras(self) -> None:
        for c in self.columnas:
            flecha = (" ▼" if self._sort_rev else " ▲") if c.id == self._sort_col else ""
            self.tree.heading(c.id, text=c.titulo + flecha)

    # --- eventos ---
    def _celda(self, event: tk.Event) -> Optional[tuple[str, int]]:
        if self.tree.identify_region(event.x, event.y) != "cell":
            return None
        iid = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not iid or not col or iid not in self._filas:
            return None
        ci = int(col.lstrip("#")) - 1
        return (iid, ci) if 0 <= ci < len(self.columnas) else None

    def _click(self, event: tk.Event) -> None:
        celda = self._celda(event)
        if celda is None:
            self.celda_activa = None
            return
        iid, ci = celda
        self.celda_activa = celda
        col = self.columnas[ci]
        if col.id == "estrella" and self._on_estrella:
            # un doble clic llega como <Double-1> (no pasa por aquí): no vuelve a alternar la estrella
            self._on_estrella(iid)
        elif self._on_mensaje:
            valor = self.valor_celda(iid, ci)
            pista = " · doble clic para editar" if col.editable else ""
            self._on_mensaje(f"{col.titulo}: {valor[:90]}  —  Ctrl+C para copiar{pista}")

    def _doble(self, event: tk.Event) -> Optional[str]:
        celda = self._celda(event)
        if celda is None:
            return None
        iid, ci = celda
        col = self.columnas[ci]
        if col.editable:
            self._editar(iid, ci)
            return "break"
        if col.id == "estado" and self._on_doble_estado:
            self._on_doble_estado(iid, event.x_root, event.y_root)
            return "break"
        return None

    def _contexto(self, event: tk.Event) -> str:
        celda = self._celda(event)
        if celda is not None:
            iid, _ci = celda
            if iid not in self.tree.selection():
                self.tree.selection_set(iid)
            self.celda_activa = celda
            if self._on_contexto:
                self._on_contexto(event)
        return "break"

    def texto_copiar(self) -> str:
        if self.celda_activa and self.celda_activa[0] in self._filas:
            return self.valor_celda(*self.celda_activa)
        filas = ["\t".join(str(v) for v in self._filas[i]) for i in self.seleccion()]
        return "\n".join(filas)

    def _copiar(self, _event: Optional[tk.Event] = None) -> str:
        texto = self.texto_copiar()
        if texto:
            self.clipboard_clear()
            self.clipboard_append(texto)
            if self._on_mensaje:
                self._on_mensaje("Copiado al portapapeles")
        return "break"

    def _borrar_celda(self, _event: Optional[tk.Event] = None) -> str:
        """Supr vacía la celda activa si es editable (igual que editarla y dejarla en blanco)."""
        if self.celda_activa and self.celda_activa[0] in self._filas and self._on_editar:
            iid, ci = self.celda_activa
            col = self.columnas[ci]
            if col.editable and self.valor_celda(iid, ci) != "":
                self._on_editar(iid, col.id, "")
        return "break"

    # --- edición en la celda ---
    def _editar(self, iid: str, ci: int) -> None:
        self.cerrar_editor(False)
        bbox = self.tree.bbox(iid, f"#{ci + 1}")
        if not bbox:
            return
        x, y, w, h = bbox
        valor = self.valor_celda(iid, ci)
        entry = tk.Entry(
            self.tree, relief="solid", bd=1, highlightthickness=2,
            highlightcolor=COLORES["primario"], highlightbackground=COLORES["primario"], bg="white",
        )
        entry.insert(0, valor)
        entry.place(x=x, y=y, width=max(w, 80), height=h)
        entry.focus_set()
        entry.select_range(0, "end")
        self._editor = (entry, iid, ci, valor)
        entry.bind("<Return>", lambda _e: self.cerrar_editor(True))
        entry.bind("<Escape>", lambda _e: self.cerrar_editor(False))
        entry.bind("<FocusOut>", lambda _e: self.cerrar_editor(True))
        if self._on_mensaje:
            self._on_mensaje(f"Editando {self.columnas[ci].titulo}: Enter guarda · Esc cancela")

    def cerrar_editor(self, guardar: bool) -> None:
        if self._editor is None or self._cerrando:
            return
        self._cerrando = True
        try:
            entry, iid, ci, original = self._editor
            self._editor = None
            nuevo = entry.get()
            entry.destroy()
            if guardar and nuevo != original and self._on_editar:
                self._on_editar(iid, self.columnas[ci].id, nuevo)
        finally:
            self._cerrando = False

    def _yview(self, *args: str) -> None:
        self.cerrar_editor(True)
        self.tree.yview(*args)

    def _xview(self, *args: str) -> None:
        self.cerrar_editor(True)
        self.tree.xview(*args)


COLS_BUSQUEDA = [
    Columna("estrella", "★", 44, "center"),
    Columna("negocio", "Negocio", 210, "w", editable=True),
    Columna("ciudad", "Ciudad", 110, "w", editable=True),
    Columna("telefono", "Teléfono", 140, "w", editable=True),
    Columna("resenas", "Reseñas", 88, "center", editable=True, numerica=True),
    Columna("web", "Web / Instagram", 190, "w", editable=True),
    Columna("problemas", "Problemas detectados", 240, "w"),
    Columna("estado", "Estado", 100, "center"),
    Columna("historial", "Historial", 84, "center"),
    Columna("link_maps", "Link Google Maps", 170, "w", editable=True),
    Columna("direccion", "Dirección", 250, "w", editable=True),
]

COLS_GUARDADOS = [
    Columna("estrella", "★", 44, "center"),
    Columna("negocio", "Negocio", 210, "w", editable=True),
    Columna("ciudad", "Ciudad", 110, "w", editable=True),
    Columna("telefono", "Teléfono", 140, "w", editable=True),
    Columna("resenas", "Reseñas", 88, "center", editable=True, numerica=True),
    Columna("web", "Web / Instagram", 190, "w", editable=True),
    Columna("problemas", "Problemas detectados", 240, "w"),
    Columna("estado", "Estado", 100, "center"),
    Columna("notas", "Notas", 220, "w"),
    Columna("guardado_en", "Guardado el", 120, "center"),
    Columna("link_maps", "Link Google Maps", 170, "w", editable=True),
    Columna("direccion", "Dirección", 250, "w", editable=True),
]

CAMPO_DE_COLUMNA = {
    "negocio": "nombre", "ciudad": "ciudad", "telefono": "telefono", "resenas": "resenas",
    "web": "web", "link_maps": "link_maps", "direccion": "direccion",
}


def _valores_busqueda(n: Negocio) -> tuple:
    return (
        "★" if n.guardado else "☆", n.nombre, n.ciudad, n.telefono, n.resenas, n.web,
        n.problemas_texto, n.estado.capitalize(), "Ya visto" if n.visto_antes else "Nuevo",
        n.link_maps, n.direccion,
    )


def _valores_guardados(n: Negocio) -> tuple:
    nota = (n.notas.strip().splitlines() or [""])[0]
    return (
        "★" if n.guardado else "☆", n.nombre, n.ciudad, n.telefono, n.resenas, n.web,
        n.problemas_texto, n.estado.capitalize(), nota, (n.guardado_en or "")[:16],
        n.link_maps, n.direccion,
    )


def _texto_busqueda(n: Negocio) -> str:
    return " ".join([
        n.nombre, n.ciudad, n.telefono, n.web, n.link_maps, n.direccion, str(n.resenas),
        n.problemas_texto, n.estado, n.notas,
    ]).lower()


# ---------------------------------------------------------------------------
# Interfaz gráfica — aplicación
# ---------------------------------------------------------------------------

class ProspeccionApp(tk.Tk):
    """Buscador de negocios + prospectos guardados."""

    def __init__(self, ruta_db: str = RUTA_DB, ruta_palabras: str = RUTA_PALABRAS) -> None:
        super().__init__()
        self.title("Prospección de Clientes — Buscador de negocios")
        self.geometry("1280x780")
        self.minsize(1020, 640)

        self._aviso_db = ""
        try:
            self._db = BaseDatos(ruta_db)
        except sqlite3.OperationalError:
            raise
        except sqlite3.DatabaseError:
            respaldo = f"{ruta_db}.corrupta-{time.strftime('%Y%m%d-%H%M%S')}"
            os.replace(ruta_db, respaldo)
            self._db = BaseDatos(ruta_db)
            self._aviso_db = f"La base de datos estaba dañada; se guardó una copia en:\n{respaldo}\nSe creó una base nueva."
        self._uso = ApiUsageTracker(self._db)
        self._hilo: Optional[threading.Thread] = None
        self._after_id: Optional[str] = None
        self._ruta_palabras = ruta_palabras
        self._palabras, self._aviso_json, self._json_invalido = cargar_palabras_clave(ruta_palabras)
        self._cancelar = threading.Event()
        self._cola: "queue.Queue[tuple]" = queue.Queue()
        self._sesion: Optional[SesionBusqueda] = None
        self._pagina_actual = 0
        self._ocupado = False
        self._visibles: list[Negocio] = []
        self._por_iid: dict[str, Negocio] = {}
        self._guardados: dict[str, Negocio] = {}
        self._nota_id: Optional[str] = None
        self._suprimir_filtro = False
        self._editor_palabras: Optional[tk.Toplevel] = None
        # --- informe importado (pestaña creada la primera vez que se importa) ---
        self.tab_i: Optional[ttk.Frame] = None
        self.tabla_i: Optional[TablaDatos] = None
        self._inf_filas: list[tuple] = []
        self._inf_estados: list[str] = []
        self._inf_texto: list[str] = []
        self._inf_aviso = ""
        self._inf_after: Optional[str] = None
        self._suprimir_filtro_i = False

        self._configurar_estilos()
        self._build_ui()
        self._centrar_ventana()
        self.protocol("WM_DELETE_WINDOW", self._cerrar)
        self.bind("<F5>", lambda _e: self._on_buscar())
        self.bind("<Control-o>", lambda _e: self._importar_informe())
        self.bind("<Control-f>", lambda _e: self.entry_filtro_b.focus_set())
        self._after_id = self.after(100, self._procesar_cola)
        self._refrescar_guardados()
        self._actualizar_botones()
        self._actualizar_uso()
        if self._aviso_json:
            self._set_estado(self._aviso_json, "aviso")
        if self._aviso_db:
            self.after(300, lambda: messagebox.showwarning("Base de datos", self._aviso_db))

    # ------------------------------------------------------------------ estilos
    def _configurar_estilos(self) -> None:
        familias = set(tkfont.families(self))
        fam = next(
            (f for f in ("Segoe UI", "SF Pro Text", "Helvetica Neue", "Noto Sans", "DejaVu Sans", "Arial") if f in familias),
            tkfont.nametofont("TkDefaultFont").actual("family"),
        )
        for nombre in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkFixedFont"):
            try:
                tkfont.nametofont(nombre).configure(size=10)
                if nombre != "TkFixedFont":
                    tkfont.nametofont(nombre).configure(family=fam)
            except tk.TclError:
                pass
        self.f_peq = (fam, 9)
        self.f_neg = (fam, 10, "bold")
        self.f_tit = (fam, 16, "bold")
        c = COLORES
        st = ttk.Style(self)
        if "clam" in st.theme_names():
            st.theme_use("clam")
        self.configure(background=c["fondo"])
        st.configure(".", background=c["fondo"], foreground=c["texto"], font=(fam, 10), bordercolor=c["borde"])
        st.configure("TFrame", background=c["fondo"])
        st.configure("Card.TFrame", background=c["tarjeta"], relief="solid", borderwidth=1, bordercolor=c["borde"])
        st.configure("TLabel", background=c["fondo"], foreground=c["texto"])
        st.configure("Card.TLabel", background=c["tarjeta"])
        st.configure("Suave.TLabel", foreground=c["suave"], font=(fam, 9))
        st.configure("CardSuave.TLabel", background=c["tarjeta"], foreground=c["suave"], font=(fam, 9))
        st.configure("Negrita.TLabel", font=(fam, 10, "bold"))
        st.configure("Header.TFrame", background=c["primario"])
        st.configure("HeaderTitulo.TLabel", background=c["primario"], foreground="white", font=(fam, 16, "bold"))
        st.configure("HeaderSub.TLabel", background=c["primario"], foreground="#CFE3F7", font=(fam, 9))
        st.configure("TButton", padding=(12, 6), background="#FFFFFF", bordercolor=c["borde"], focuscolor="#FFFFFF")
        st.map("TButton", background=[("active", c["primario_claro"]), ("disabled", "#EEF0F3")],
               foreground=[("disabled", "#A0A7B1")])
        st.configure("Primary.TButton", background=c["primario"], foreground="white", bordercolor=c["primario"],
                     font=(fam, 10, "bold"), padding=(16, 7))
        st.map("Primary.TButton", background=[("active", c["primario_osc"]), ("disabled", "#9DBFE0")],
               foreground=[("disabled", "#FFFFFF")])
        st.configure("Header.TButton", background="#2A7FCB", foreground="white", bordercolor="#2A7FCB", padding=(12, 5))
        st.map("Header.TButton", background=[("active", c["primario_osc"])])
        st.configure("TMenubutton", padding=(12, 6), background="#FFFFFF", bordercolor=c["borde"], arrowcolor=c["texto"])
        st.map("TMenubutton", background=[("active", c["primario_claro"]), ("disabled", "#EEF0F3")],
               foreground=[("disabled", "#A0A7B1")])
        st.configure("TCombobox", fieldbackground="#FFFFFF", background="#FFFFFF", padding=4)
        st.map("TCombobox", fieldbackground=[("readonly", "#FFFFFF")], foreground=[("readonly", c["texto"])])
        st.configure("TEntry", fieldbackground="#FFFFFF", padding=5)
        st.configure("TNotebook", background=c["fondo"], borderwidth=0, tabmargins=(0, 0, 0, 0))
        st.configure("TNotebook.Tab", padding=(22, 9), background="#E3E8EF", foreground=c["suave"], font=(fam, 10, "bold"))
        st.map("TNotebook.Tab", background=[("selected", c["tarjeta"])], foreground=[("selected", c["primario"])])
        st.configure("Datos.Treeview", rowheight=28, background=c["tarjeta"], fieldbackground=c["tarjeta"],
                     foreground=c["texto"], bordercolor=c["borde"], borderwidth=1, font=(fam, 10))
        st.map("Datos.Treeview", background=[("selected", c["sel"])], foreground=[("selected", c["texto"])])
        st.configure("Datos.Treeview.Heading", background="#E6ECF4", foreground=c["texto"], font=(fam, 10, "bold"),
                     relief="flat", padding=(6, 7))
        st.map("Datos.Treeview.Heading", background=[("active", "#D6E2F0")])
        st.configure("Horizontal.TProgressbar", background=c["primario"], troughcolor="#E3E8EF", bordercolor=c["borde"])
        # Tk 8.6.9: los colores por etiqueta del Treeview no se aplicaban sin este arreglo
        def _sin_bug(opcion: str):
            return [e for e in st.map("Treeview", query_opt=opcion) if e[:2] != ("!disabled", "!selected")]
        st.map("Treeview", foreground=_sin_bug("foreground"), background=_sin_bug("background"))

    # --------------------------------------------------------------------- UI
    def _centrar_ventana(self) -> None:
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = max(0, (self.winfo_screenwidth() // 2) - (w // 2))
        y = max(0, (self.winfo_screenheight() // 2) - (h // 2))
        self.geometry(f"+{x}+{y}")

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        # --- cabecera ---
        cab = ttk.Frame(self, style="Header.TFrame", padding=(18, 12))
        cab.grid(row=0, column=0, sticky="ew")
        cab.columnconfigure(0, weight=1)
        titulos = ttk.Frame(cab, style="Header.TFrame")
        titulos.grid(row=0, column=0, sticky="w")
        ttk.Label(titulos, text="Buscador de negocios", style="HeaderTitulo.TLabel").pack(anchor="w")
        ttk.Label(titulos, text="Google Maps · prospectos guardados en SQLite", style="HeaderSub.TLabel").pack(anchor="w")
        ttk.Button(cab, text="Importar informe", style="Header.TButton", command=self._importar_informe).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(cab, text="Palabras clave", style="Header.TButton", command=self._abrir_editor_palabras).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(cab, text="Info de uso", style="Header.TButton", command=self._mostrar_info_api).grid(row=0, column=3)

        # --- pestañas ---
        self.nb = ttk.Notebook(self)
        self.nb.grid(row=1, column=0, sticky="nsew", padx=12, pady=(10, 4))
        self.tab_b = ttk.Frame(self.nb, padding=(0, 10, 0, 0))
        self.tab_g = ttk.Frame(self.nb, padding=(0, 10, 0, 0))
        self.nb.add(self.tab_b, text="Búsqueda")
        self.nb.add(self.tab_g, text="★ Guardados")
        self.nb.bind("<<NotebookTabChanged>>", self._on_cambio_pestana)
        self._build_tab_busqueda()
        self._build_tab_guardados()

        # --- barra de estado ---
        barra = ttk.Frame(self, padding=(14, 4, 14, 8))
        barra.grid(row=2, column=0, sticky="ew")
        barra.columnconfigure(0, weight=1)
        self.lbl_estado = ttk.Label(barra, text="Listo para buscar", style="Suave.TLabel")
        self.lbl_estado.grid(row=0, column=0, sticky="w")
        self.progreso = ttk.Progressbar(barra, mode="indeterminate", length=140)
        self.progreso.grid(row=0, column=1, padx=10)
        self.progreso.grid_remove()
        self.lbl_uso = ttk.Label(barra, text="", style="Suave.TLabel")
        self.lbl_uso.grid(row=0, column=2, sticky="e")

    def _crear_menu_exportar(self, parent: tk.Widget, origen: str) -> ttk.Menubutton:
        mb = ttk.Menubutton(parent, text="Exportar")
        menu = tk.Menu(mb, tearoff=0)
        for fmt, etiqueta in (("xlsx", "Excel"), ("csv", "CSV")):
            menu.add_command(label=f"{etiqueta} · todo lo cargado" if origen == "busqueda" else f"{etiqueta} · todos los guardados",
                             command=lambda f=fmt: self._exportar(origen, "todo", f))
            menu.add_command(label=f"{etiqueta} · solo lo filtrado en la tabla", command=lambda f=fmt: self._exportar(origen, "filtrado", f))
            menu.add_command(label=f"{etiqueta} · solo lo seleccionado", command=lambda f=fmt: self._exportar(origen, "seleccion", f))
            menu.add_separator()
        if origen == "busqueda":
            menu.add_command(label="Cargar todo y exportar a Excel…", command=self._on_cargar_todo)
        else:
            menu.delete(menu.index("end"))     # quita el último separador
        mb["menu"] = menu
        return mb

    def _build_tab_busqueda(self) -> None:
        t = self.tab_b
        t.columnconfigure(0, weight=1)
        t.rowconfigure(3, weight=1)

        card = ttk.Frame(t, style="Card.TFrame", padding=(14, 12))
        card.grid(row=0, column=0, sticky="ew")
        for col in (1, 3, 5):
            card.columnconfigure(col, weight=1)
        ttk.Label(card, text="País", style="Card.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.var_pais = tk.StringVar(value="Colombia")
        self.combo_pais = ttk.Combobox(card, textvariable=self.var_pais, values=PAISES, state="readonly", width=16)
        self.combo_pais.grid(row=0, column=1, sticky="ew", padx=(0, 14))
        self.combo_pais.bind("<<ComboboxSelected>>", self._on_pais_change)

        ttk.Label(card, text="Ciudad", style="Card.TLabel").grid(row=0, column=2, sticky="w", padx=(0, 6))
        self.var_ciudad = tk.StringVar(value=CIUDADES_POR_PAIS["Colombia"][0])
        self.combo_ciudad = ttk.Combobox(card, textvariable=self.var_ciudad, values=CIUDADES_POR_PAIS["Colombia"], state="readonly", width=16)
        self.combo_ciudad.grid(row=0, column=3, sticky="ew", padx=(0, 14))

        ttk.Label(card, text="Tipo de negocio", style="Card.TLabel").grid(row=0, column=4, sticky="w", padx=(0, 6))
        self.var_tipo = tk.StringVar(value="Médicos" if "Médicos" in self._palabras else next(iter(self._palabras)))
        self.combo_tipo = ttk.Combobox(card, textvariable=self.var_tipo, values=list(self._palabras), state="readonly", width=16)
        self.combo_tipo.grid(row=0, column=5, sticky="ew", padx=(0, 14))
        self.combo_tipo.bind("<<ComboboxSelected>>", lambda _e: self._actualizar_resumen_terminos())

        ttk.Label(card, text="Mín. reseñas", style="Card.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=(10, 0))
        self.var_min = tk.StringVar(value="5")
        self.combo_min = ttk.Combobox(card, textvariable=self.var_min, values=MIN_RESEÑAS_OPCIONES, state="readonly", width=6)
        self.combo_min.grid(row=1, column=1, sticky="w", padx=(0, 14), pady=(10, 0))
        self.combo_min.bind("<<ComboboxSelected>>", self._on_cambio_reaplicar)

        ttk.Label(card, text="Por página", style="Card.TLabel").grid(row=1, column=2, sticky="w", padx=(0, 6), pady=(10, 0))
        self.var_pp = tk.StringVar(value="25")
        self.combo_pp = ttk.Combobox(card, textvariable=self.var_pp, values=OPCIONES_POR_PAGINA, state="readonly", width=6)
        self.combo_pp.grid(row=1, column=3, sticky="w", padx=(0, 14), pady=(10, 0))
        self.combo_pp.bind("<<ComboboxSelected>>", self._on_cambio_reaplicar)

        botones = ttk.Frame(card, style="Card.TFrame")
        botones.grid(row=1, column=4, columnspan=2, sticky="e", pady=(10, 0))
        self.btn_buscar = ttk.Button(botones, text="Buscar negocios", style="Primary.TButton", command=self._on_buscar)
        self.btn_buscar.pack(side="left")
        self.btn_cancelar = ttk.Button(botones, text="Cancelar", command=self._on_cancelar, state="disabled")
        self.btn_cancelar.pack(side="left", padx=(8, 0))

        self.lbl_terminos = ttk.Label(card, text="", style="CardSuave.TLabel")
        self.lbl_terminos.grid(row=2, column=0, columnspan=6, sticky="w", pady=(8, 0))

        # barra de herramientas
        tools = ttk.Frame(t)
        tools.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        tools.columnconfigure(1, weight=1)
        ttk.Label(tools, text="Filtrar en la página").grid(row=0, column=0, padx=(0, 8))
        self.var_filtro_b = tk.StringVar()
        self.var_filtro_b.trace_add("write", lambda *_: self._on_filtro_b())
        self.entry_filtro_b = ttk.Entry(tools, textvariable=self.var_filtro_b, width=34)
        self.entry_filtro_b.grid(row=0, column=1, sticky="w")
        self.lbl_leads_b = ttk.Label(tools, text="0 leads", style="Negrita.TLabel", foreground=COLORES["ok"])
        self.lbl_leads_b.grid(row=0, column=2, padx=14)
        self.btn_guardar_b = ttk.Button(tools, text="★ Guardar / quitar", command=lambda: self._guardar_seleccion("busqueda"))
        self.btn_guardar_b.grid(row=0, column=3, padx=4)
        self.btn_msg_b = ttk.Button(tools, text="Copiar mensaje", command=self._on_copiar_mensaje)
        self.btn_msg_b.grid(row=0, column=4, padx=4)
        self.mb_export_b = self._crear_menu_exportar(tools, "busqueda")
        self.mb_export_b.grid(row=0, column=5, padx=(4, 0))

        # paginación
        pag = ttk.Frame(t)
        pag.grid(row=2, column=0, sticky="w", pady=(8, 8))
        self.btn_ant = ttk.Button(pag, text="◀ Anterior", command=self._pagina_anterior)
        self.btn_ant.pack(side="left")
        self.lbl_pagina = ttk.Label(pag, text="Página —", style="Negrita.TLabel", padding=(14, 0))
        self.lbl_pagina.pack(side="left")
        self.btn_sig = ttk.Button(pag, text="Siguiente ▶", command=self._pagina_siguiente)
        self.btn_sig.pack(side="left")
        ttk.Label(pag, text="  ·  Clic en ☆ para guardar  ·  doble clic en una celda para editar  ·  clic derecho para más acciones",
                  style="Suave.TLabel").pack(side="left", padx=(10, 0))

        self.tabla_b = TablaDatos(
            t, COLS_BUSQUEDA,
            on_estrella=lambda iid: self._toggle_guardado(self._por_iid.get(iid)),
            on_editar=lambda iid, col, val: self._aplicar_edicion(self._por_iid.get(iid), col, val),
            on_contexto=lambda ev: self._menu_contexto("busqueda", ev),
            on_doble_estado=lambda iid, x, y: self._menu_estado_en(("busqueda", x, y)),
            on_seleccion=self._actualizar_botones,
            on_mensaje=self._mensaje_tabla,
        )
        self.tabla_b.grid(row=3, column=0, sticky="nsew")
        self._actualizar_resumen_terminos()

    def _build_tab_guardados(self) -> None:
        t = self.tab_g
        t.columnconfigure(0, weight=1)
        t.rowconfigure(1, weight=1)

        tools = ttk.Frame(t)
        tools.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        tools.columnconfigure(1, weight=1)
        ttk.Label(tools, text="Filtrar").grid(row=0, column=0, padx=(0, 8))
        self.var_filtro_g = tk.StringVar()
        self.var_filtro_g.trace_add("write", lambda *_: self._refrescar_guardados())
        self.entry_filtro_g = ttk.Entry(tools, textvariable=self.var_filtro_g, width=34)
        self.entry_filtro_g.grid(row=0, column=1, sticky="w")
        self.lbl_leads_g = ttk.Label(tools, text="0 guardados", style="Negrita.TLabel", foreground=COLORES["ok"])
        self.lbl_leads_g.grid(row=0, column=2, padx=14)
        self.mb_estado_g = ttk.Menubutton(tools, text="Cambiar estado")
        self.mb_estado_g["menu"] = self._crear_menu_estado(self.mb_estado_g, "guardados")
        self.mb_estado_g.grid(row=0, column=3, padx=4)
        self.btn_tel_g = ttk.Button(tools, text="Copiar teléfono", command=lambda: self._copiar_telefonos("guardados"))
        self.btn_tel_g.grid(row=0, column=4, padx=4)
        self.btn_quitar_g = ttk.Button(tools, text="Quitar de guardados", command=lambda: self._quitar_guardados())
        self.btn_quitar_g.grid(row=0, column=5, padx=4)
        self.mb_export_g = self._crear_menu_exportar(tools, "guardados")
        self.mb_export_g.grid(row=0, column=6, padx=(4, 0))

        self.tabla_g = TablaDatos(
            t, COLS_GUARDADOS,
            on_estrella=lambda iid: self._quitar_guardados([self._guardados[iid]] if iid in self._guardados else []),
            on_editar=lambda iid, col, val: self._aplicar_edicion(self._guardados.get(iid), col, val),
            on_contexto=lambda ev: self._menu_contexto("guardados", ev),
            on_doble_estado=lambda iid, x, y: self._menu_estado_en(("guardados", x, y)),
            on_seleccion=self._on_sel_guardados,
            on_mensaje=self._mensaje_tabla,
        )
        self.tabla_g.grid(row=1, column=0, sticky="nsew")

        panel = ttk.Frame(t, style="Card.TFrame", padding=(12, 8))
        panel.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        panel.columnconfigure(0, weight=1)
        self.lbl_nota = ttk.Label(panel, text="Notas — selecciona un prospecto", style="Card.TLabel", font=self.f_neg)
        self.lbl_nota.grid(row=0, column=0, sticky="w")
        self.txt_nota = tk.Text(panel, height=3, wrap="word", relief="solid", bd=1, state="disabled",
                                highlightthickness=1, highlightbackground=COLORES["borde"], highlightcolor=COLORES["primario"])
        self.txt_nota.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.txt_nota.bind("<FocusOut>", lambda _e: self._guardar_nota_pendiente())
        self.txt_nota.bind("<Control-Return>", lambda _e: (self._guardar_nota_pendiente(), "break")[1])
        self.btn_nota = ttk.Button(panel, text="Guardar nota", command=self._guardar_nota_pendiente, state="disabled")
        self.btn_nota.grid(row=1, column=1, padx=(10, 0), sticky="s")

    # --------------------------------------------------------------- utilidades UI
    def _set_estado(self, texto: str, nivel: str = "normal") -> None:
        color = {"normal": COLORES["suave"], "ok": COLORES["ok"], "aviso": COLORES["aviso"], "error": COLORES["error"]}[nivel]
        self.lbl_estado.config(text=texto, foreground=color)

    def _mensaje_tabla(self, texto: str) -> None:
        if not self._ocupado:
            self._set_estado(texto)

    def _actualizar_uso(self) -> None:
        self.lbl_uso.config(text=f"Peticiones: {self._uso.peticiones_sesion} esta sesión · {self._db.uso_mes()} este mes")

    def _tipo_terminos(self, tipo: str) -> list[str]:
        return self._palabras.get(tipo) or [tipo]

    def _actualizar_resumen_terminos(self) -> None:
        n = len(self._tipo_terminos(self.var_tipo.get()))
        self.lbl_terminos.config(
            text=f"«{self.var_tipo.get()}»: {n} términos de búsqueda · 1 consulta por término (hasta 3 páginas de 20 resultados cada una)"
        )

    def _on_pais_change(self, *_: object) -> None:
        ciudades = CIUDADES_POR_PAIS.get(self.var_pais.get(), [])
        if ciudades:
            self.combo_ciudad["values"] = ciudades
            self.var_ciudad.set(ciudades[0])

    def _limite_pagina(self) -> int:
        if self.var_pp.get().strip().lower() == "todos":
            return PAGINA_TODOS
        try:
            return max(1, int(self.var_pp.get()))
        except ValueError:
            return 10

    def _min_resenas(self) -> int:
        try:
            return max(0, int(self.var_min.get()))
        except ValueError:
            return 0

    def _actualizar_botones(self) -> None:
        busy = self._ocupado
        s = self._sesion
        tiene = s is not None and self._pagina_actual > 0 and self._pagina_actual in s.paginas
        sel_b = bool(self.tabla_b.seleccion())
        sel_g = bool(self.tabla_g.seleccion())
        est = lambda cond: "normal" if cond else "disabled"
        self.btn_buscar.config(state=est(not busy))
        self.btn_cancelar.config(state=est(busy))
        self.btn_ant.config(state=est(tiene and not busy and self._pagina_actual > 1))
        sig = tiene and ((self._pagina_actual + 1) in s.paginas or s.hay_mas())
        self.btn_sig.config(state=est(sig and not busy))
        self.lbl_pagina.config(text=f"Página {self._pagina_actual}" if tiene else "Página —")
        self.mb_export_b.config(state=est(s is not None and not busy))
        self.btn_guardar_b.config(state=est(sel_b and not busy))
        self.btn_msg_b.config(state=est(sel_b))
        self.mb_estado_g.config(state=est(sel_g))
        self.btn_tel_g.config(state=est(sel_g))
        self.btn_quitar_g.config(state=est(sel_g))
        self.mb_export_g.config(state=est(bool(self._guardados)))

    def _actualizar_contador_b(self) -> None:
        visibles = len(self._visibles)
        s = self._sesion
        if s is not None and self._pagina_actual > 0:
            texto = f"Pág. {self._pagina_actual}: {visibles} leads · {len(s.raw)} en caché"
        elif s is not None:
            texto = "0 leads"
        else:
            texto = "0 leads"
        self.lbl_leads_b.config(text=texto, foreground=COLORES["ok"] if visibles else COLORES["error"])

    # -------------------------------------------------- hilos / cola de mensajes
    def _lanzar(self, trabajo: Callable[[], object], al_terminar: Callable[[object, Optional[BaseException]], None], texto: str) -> None:
        self._cerrar_editores()
        self._ocupado = True
        self._cancelar.clear()
        self.progreso.grid()
        self.progreso.start(12)
        self._set_estado(texto, "aviso")
        self._actualizar_botones()

        def run() -> None:
            try:
                self._cola.put(("fin", al_terminar, trabajo(), None))
            except BaseException as exc:      # noqa: BLE001 — se informa en la interfaz
                self._cola.put(("fin", al_terminar, None, exc))

        self._hilo = threading.Thread(target=run, daemon=True)
        self._hilo.start()

    def _cerrar_editores(self) -> None:
        self.tabla_b.cerrar_editor(True)
        self.tabla_g.cerrar_editor(True)

    def _procesar_cola(self) -> None:
        try:
            while True:
                item = self._cola.get_nowait()
                if item[0] == "progreso":
                    if self._ocupado:
                        self._set_estado(item[1], "aviso")
                        self._actualizar_uso()
                elif item[0] == "fin":
                    _, callback, resultado, exc = item
                    self._ocupado = False
                    self.progreso.stop()
                    self.progreso.grid_remove()
                    self._actualizar_uso()
                    try:
                        callback(resultado, exc)
                    except Exception as err:   # noqa: BLE001
                        messagebox.showerror("Error", f"Ocurrió un error inesperado:\n{err}")
                    self._actualizar_botones()
        except queue.Empty:
            pass
        try:
            self._after_id = self.after(100, self._procesar_cola)
        except tk.TclError:
            pass

    def _on_cancelar(self) -> None:
        self._cancelar.set()
        self._set_estado("Cancelando… (termina la petición en curso)", "aviso")
        self.btn_cancelar.config(state="disabled")

    # ------------------------------------------------------------- búsqueda
    def _on_buscar(self) -> None:
        if self._ocupado:
            return
        self._guardar_nota_pendiente()
        limite = self._limite_pagina()
        if limite >= PAGINA_TODOS:
            tope = TOPE_PETICIONES_MAX
            if not messagebox.askyesno(
                "Cargar todos los resultados",
                f"Con «todos» la búsqueda puede usar hasta {tope} peticiones "
                f"(≈ ${tope * COSTO_ESTIMADO_POR_PETICION_USD:.2f} USD estimados).\n\n¿Continuar?",
            ):
                return
        try:
            cliente = ClienteGooglePlaces(_leer_google_api_key(), self._uso, self._cancelar)
        except ValueError as exc:
            messagebox.showerror("Falta la API key", str(exc))
            return
        pais, ciudad, tipo = self.var_pais.get(), self.var_ciudad.get(), self.var_tipo.get()
        sesion = SesionBusqueda(
            cliente, self._db, pais, ciudad, tipo, self._tipo_terminos(tipo),
            self._min_resenas(), limite, progreso=lambda txt: self._cola.put(("progreso", txt)),
        )
        self._lanzar(
            sesion.construir_pagina,
            lambda res, exc: self._fin_pagina(sesion, True, res, exc),
            f"Consultando Google Maps — {ciudad}: {tipo}…",
        )

    def _pagina_siguiente(self) -> None:
        s = self._sesion
        if s is None or self._ocupado:
            return
        siguiente = self._pagina_actual + 1
        if siguiente in s.paginas:
            self._mostrar_pagina(siguiente)
        elif s.hay_mas():
            self._lanzar(
                s.construir_pagina,
                lambda res, exc: self._fin_pagina(s, False, res, exc),
                f"Consultando Google Maps (página {siguiente})…",
            )

    def _pagina_anterior(self) -> None:
        if self._sesion is not None and not self._ocupado and self._pagina_actual > 1:
            self._mostrar_pagina(self._pagina_actual - 1)

    def _fin_pagina(self, sesion: SesionBusqueda, nueva: bool, pagina: object, exc: Optional[BaseException]) -> None:
        if exc is not None:
            messagebox.showerror("Error", f"No se pudo completar la búsqueda:\n{exc}")
            self._set_estado("Error en la búsqueda", "error")
            return
        if nueva:
            self._sesion = sesion
            self._pagina_actual = 0
        if not pagina:
            self._render_busqueda()
            if nueva:
                self._set_estado("Sin resultados para esta búsqueda", "error")
            else:
                messagebox.showinfo("Sin más resultados", "Google Maps no devolvió más negocios nuevos para esta búsqueda.")
            if sesion.error:
                messagebox.showwarning("Búsqueda incompleta", sesion.error)
            return
        self._mostrar_pagina(max(sesion.paginas))
        if sesion.error:
            messagebox.showwarning(
                "Búsqueda incompleta",
                f"Google devolvió un error; se muestran los resultados que ya estaban cargados.\n\n{sesion.error}",
            )
        if sesion.cancelado:
            self._set_estado(f"Búsqueda cancelada — se muestran {len(pagina)} resultados ya cargados", "aviso")

    def _on_cambio_reaplicar(self, *_: object) -> None:
        """Cambiar el mínimo de reseñas o el tamaño de página reparte de nuevo lo que ya está en caché."""
        s = self._sesion
        if s is None or self._ocupado:
            return
        self._cerrar_editores()
        s.reconfigurar(self._min_resenas(), self._limite_pagina())
        pagina = s.construir_pagina(permitir_api=False)
        self._pagina_actual = 0
        if pagina:
            self._mostrar_pagina(1)
            extra = " · pulsa «Siguiente» para pedir más a Google" if s.hay_mas() and not s.agotado else ""
            self._set_estado(f"Reaplicado sobre {len(s.raw)} negocios en caché (sin consultar Google){extra}", "ok")
        else:
            self._render_busqueda()
        self._actualizar_botones()

    def _mostrar_pagina(self, numero: int) -> None:
        s = self._sesion
        if s is None or numero not in s.paginas:
            return
        self._pagina_actual = numero
        self._db.sincronizar(s.paginas[numero])
        self._suprimir_filtro = True
        self.var_filtro_b.set("")
        self._suprimir_filtro = False
        self._render_busqueda()
        omit = f" · {len(s.omitidos)} términos omitidos" if s.omitidos else ""
        self._set_estado(
            f"Página {numero}: {len(s.paginas[numero])} leads · {s.ciudad} — {s.tipo} · {s.peticiones} peticiones{omit}",
            "ok",
        )
        self._actualizar_botones()

    def _negocios_pagina(self) -> list[Negocio]:
        s = self._sesion
        return s.paginas.get(self._pagina_actual, []) if s is not None else []

    def _on_filtro_b(self) -> None:
        if not self._suprimir_filtro:
            self._render_busqueda()

    def _render_busqueda(self) -> None:
        termino = self.var_filtro_b.get().strip().lower()
        base = self._negocios_pagina()
        self._visibles = [n for n in base if not termino or termino in _texto_busqueda(n)]
        self._por_iid = {n.place_id: n for n in self._visibles}
        self.tabla_b.set_filas([(n.place_id, _valores_busqueda(n), n.estado) for n in self._visibles])
        self._actualizar_contador_b()
        self._actualizar_botones()

    # -------------------------------------------------- guardados / estados / edición
    def _tabla(self, origen: str) -> TablaDatos:
        return self.tabla_b if origen == "busqueda" else self.tabla_g

    def _mapa(self, origen: str) -> dict[str, Negocio]:
        return self._por_iid if origen == "busqueda" else self._guardados

    def _seleccionados(self, origen: str) -> list[Negocio]:
        mapa = self._mapa(origen)
        return [mapa[i] for i in self._tabla(origen).seleccion() if i in mapa]

    def _propagar(self, n: Negocio) -> None:
        """Copia los datos del usuario a las demás copias en memoria del mismo prospecto y repinta."""
        copias: list[Negocio] = []
        if self._sesion is not None and n.place_id in self._sesion.raw:
            copias.append(self._sesion.raw[n.place_id])
        if n.place_id in self._guardados:
            copias.append(self._guardados[n.place_id])
        for o in copias:
            if o is n:
                continue
            o.estado, o.notas, o.guardado, o.guardado_en = n.estado, n.notas, n.guardado, n.guardado_en
            for campo in CAMPOS_EDITABLES:
                setattr(o, campo, getattr(n, campo))
        if n.place_id in self._por_iid:
            o = self._por_iid[n.place_id]
            self.tabla_b.actualizar_fila(n.place_id, _valores_busqueda(o), o.estado)
        if n.place_id in self._guardados:
            o = self._guardados[n.place_id]
            self.tabla_g.actualizar_fila(n.place_id, _valores_guardados(o), o.estado)

    def _toggle_guardado(self, n: Optional[Negocio]) -> None:
        if n is None:
            return
        nuevo = not n.guardado
        guardado_en = self._db.set_guardado(n.place_id, nuevo)
        if guardado_en is None:
            self._set_estado("No se pudo guardar: el prospecto no está en la base de datos", "error")
            return
        n.guardado, n.guardado_en = nuevo, guardado_en
        self._propagar(n)
        self._refrescar_guardados()
        self._set_estado(f"«{n.nombre}» {'guardado en ★ Guardados' if nuevo else 'quitado de Guardados'}", "ok")

    def _guardar_seleccion(self, origen: str) -> None:
        sel = self._seleccionados(origen)
        if not sel:
            return
        guardar = any(not n.guardado for n in sel)
        cambiados = 0
        for n in sel:
            if n.guardado != guardar:
                ge = self._db.set_guardado(n.place_id, guardar)
                if ge is not None:
                    n.guardado, n.guardado_en = guardar, ge
                    self._propagar(n)
                    cambiados += 1
        self._refrescar_guardados()
        self._set_estado(f"{cambiados} prospecto(s) {'guardados' if guardar else 'quitados de Guardados'}", "ok")

    def _quitar_guardados(self, negocios: Optional[list[Negocio]] = None) -> None:
        sel = negocios if negocios is not None else self._seleccionados("guardados")
        if not sel:
            return
        con_datos = [n for n in sel if n.notas.strip() or n.estado != "nuevo"]
        if con_datos and not messagebox.askyesno(
            "Quitar de guardados",
            f"{len(con_datos)} de los prospectos tienen notas o un estado distinto de «nuevo».\n"
            "Se conservan en la base de datos y reaparecerán si los vuelves a guardar.\n\n¿Quitarlos de Guardados?",
        ):
            return
        self._guardar_nota_pendiente()
        for n in sel:
            if self._db.set_guardado(n.place_id, False) is not None:
                n.guardado, n.guardado_en = False, ""
                self._propagar(n)
        self._refrescar_guardados()
        self._set_estado(f"{len(sel)} prospecto(s) quitados de Guardados", "ok")

    def _crear_menu_estado(self, master: tk.Widget, origen: str) -> tk.Menu:
        menu = tk.Menu(master, tearoff=0)
        for estado in ESTADOS:
            menu.add_command(label=estado.capitalize(), command=lambda e=estado: self._cambiar_estado(origen, e))
        return menu

    def _menu_estado_en(self, destino: tuple[str, int, int]) -> None:
        origen, x, y = destino
        menu = self._crear_menu_estado(self, origen)
        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def _cambiar_estado(self, origen: str, estado: str) -> None:
        sel = self._seleccionados(origen)
        for n in sel:
            if self._db.set_estado(n.place_id, estado):
                n.estado = estado
                self._propagar(n)
        if sel:
            self._set_estado(f"Estado «{estado}» aplicado a {len(sel)} prospecto(s)", "ok")

    def _aplicar_edicion(self, n: Optional[Negocio], col_id: str, valor: str) -> None:
        if n is None or col_id not in CAMPO_DE_COLUMNA:
            return
        campo = CAMPO_DE_COLUMNA[col_id]
        valor = valor.strip()
        if campo == "resenas":
            if not valor.isdigit():
                self._set_estado("Reseñas debe ser un número entero; no se cambió", "error")
                self._propagar(n)
                return
            nuevo: object = int(valor)
        elif campo == "nombre" and not valor:
            self._set_estado("El nombre no puede quedar vacío; no se cambió", "error")
            self._propagar(n)
            return
        else:
            nuevo = valor
        setattr(n, campo, nuevo)
        self._db.set_edicion(n, campo, nuevo)
        self._propagar(n)
        self._set_estado("Cambio guardado (se conserva aunque Google actualice el prospecto)", "ok")

    def _refrescar_guardados(self) -> None:
        self._guardar_nota_pendiente()
        termino = self.var_filtro_g.get().strip().lower() if hasattr(self, "var_filtro_g") else ""
        todos = self._db.listar_guardados()
        self._guardados = {n.place_id: n for n in todos if not termino or termino in _texto_busqueda(n)}
        self.tabla_g.set_filas([(i, _valores_guardados(n), n.estado) for i, n in self._guardados.items()])
        self.nb.tab(self.tab_g, text=f"★ Guardados ({len(todos)})")
        self.lbl_leads_g.config(
            text=f"{len(self._guardados)} guardados" + (f" de {len(todos)}" if len(todos) != len(self._guardados) else ""),
            foreground=COLORES["ok"] if todos else COLORES["suave"],
        )
        self._nota_id = None
        self._on_sel_guardados()

    def _on_cambio_pestana(self, _event: Optional[tk.Event] = None) -> None:
        self._cerrar_editores()
        self._guardar_nota_pendiente()
        indice = self.nb.index(self.nb.select())
        if indice == 1:
            self._refrescar_guardados()
        elif indice == 0:
            self._db.sincronizar(self._negocios_pagina())
            for n in self._visibles:
                self.tabla_b.actualizar_fila(n.place_id, _valores_busqueda(n), n.estado)

    # --- notas ---
    def _on_sel_guardados(self) -> None:
        self._guardar_nota_pendiente()
        sel = self.tabla_g.seleccion()
        self.txt_nota.config(state="normal")
        self.txt_nota.delete("1.0", "end")
        if len(sel) == 1 and sel[0] in self._guardados:
            n = self._guardados[sel[0]]
            self._nota_id = n.place_id
            self.txt_nota.insert("1.0", n.notas)
            self.lbl_nota.config(text=f"Notas — {n.nombre}")
            self.btn_nota.config(state="normal")
        else:
            self._nota_id = None
            self.txt_nota.config(state="disabled")
            self.lbl_nota.config(text="Notas — selecciona un prospecto" if not sel else "Notas — selecciona un solo prospecto")
            self.btn_nota.config(state="disabled")
        self._actualizar_botones()

    def _guardar_nota_pendiente(self) -> None:
        if not self._nota_id or not hasattr(self, "txt_nota"):
            return
        n = self._guardados.get(self._nota_id)
        if n is None:
            return
        texto = self.txt_nota.get("1.0", "end-1c")
        if texto != n.notas:
            if self._db.set_notas(n.place_id, texto):
                n.notas = texto
                self._propagar(n)
                self._set_estado("Nota guardada", "ok")

    # ---------------------------------------------------- menú contextual y portapapeles
    def _menu_contexto(self, origen: str, event: tk.Event) -> None:
        sel = self._seleccionados(origen)
        if not sel:
            return
        n = sel[0]
        menu = tk.Menu(self, tearoff=0)
        todos_guardados = all(x.guardado for x in sel)
        etiqueta = f" ({len(sel)})" if len(sel) > 1 else ""
        menu.add_command(
            label=("☆ Quitar de guardados" if todos_guardados else "★ Guardar") + etiqueta,
            command=(lambda: self._quitar_guardados()) if (origen == "guardados") else (lambda: self._guardar_seleccion(origen)),
        )
        menu.add_cascade(label="Cambiar estado", menu=self._crear_menu_estado(menu, origen))
        menu.add_separator()
        menu.add_command(label="Copiar celda", command=lambda: self._tabla(origen)._copiar())
        menu.add_command(label="Copiar teléfono" + etiqueta, command=lambda: self._copiar_telefonos(origen))
        if origen == "busqueda":
            menu.add_command(label="Copiar mensaje", command=self._on_copiar_mensaje)
        menu.add_separator()
        menu.add_command(label="Abrir en Google Maps", command=lambda: self._abrir(n.link_maps, "enlace de Maps"),
                         state="normal" if n.link_maps else "disabled")
        menu.add_command(label="Abrir sitio web", command=lambda: self._abrir(n.web, "sitio web"),
                         state="normal" if n.web.lower().startswith(("http://", "https://")) else "disabled")
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _abrir(self, url: str, nombre: str) -> None:
        if not abrir_url(url):
            self._set_estado(f"No hay un {nombre} válido para abrir", "error")

    def _copiar(self, texto: str, mensaje: str) -> None:
        self.clipboard_clear()
        self.clipboard_append(texto)
        self._set_estado(mensaje, "ok")

    def _copiar_telefonos(self, origen: str) -> None:
        tels = [n.telefono for n in self._seleccionados(origen) if n.telefono.strip()]
        if tels:
            self._copiar("\n".join(tels), f"{len(tels)} teléfono(s) copiado(s)")
        else:
            self._set_estado("Los prospectos seleccionados no tienen teléfono", "aviso")

    def _on_copiar_mensaje(self) -> None:
        sel = self._seleccionados("busqueda") or self._seleccionados("guardados")
        if sel:
            self._copiar(mensaje_para(sel[0]), "Mensaje copiado al portapapeles")

    # ------------------------------------------------------------------ exportación
    def _lista_export(self, origen: str, modo: str) -> list[Negocio]:
        if origen == "busqueda":
            if modo == "todo":
                return self._sesion.todos_los_negocios() if self._sesion else []
            mapa = self._por_iid
        else:
            if modo == "todo":
                return list(self._db.listar_guardados())
            mapa = self._guardados
        tabla = self._tabla(origen)
        ids = tabla.iids() if modo == "filtrado" else tabla.seleccion()
        return [mapa[i] for i in ids if i in mapa]

    def _exportar(self, origen: str, modo: str, fmt: str) -> None:
        if self._ocupado:
            return
        self._cerrar_editores()
        lista = self._lista_export(origen, modo)
        if not lista:
            messagebox.showinfo("Exportar", "No hay prospectos para exportar con esa opción.")
            return
        self._escribir_export(lista, origen, fmt)

    def _escribir_export(self, lista: list[Negocio], origen: str, fmt: str) -> None:
        self._guardar_nota_pendiente()
        self._db.sincronizar(lista)
        if origen == "busqueda" and self._sesion is not None:
            s = self._sesion
            leads = "todos" if s.limite >= PAGINA_TODOS else str(len(lista))
            partes = [s.pais, s.ciudad, leads, s.tipo]
        else:
            partes = ["Guardados", str(len(lista)), "prospectos"]
        try:
            os.makedirs(INFORMES_DIR, exist_ok=True)
            ruta = ruta_unica(os.path.join(INFORMES_DIR, nombre_exportacion(partes, fmt)))
            self.config(cursor="watch")
            self.update_idletasks()
            (exportar_excel if fmt == "xlsx" else exportar_csv)(ruta, lista)
        except (OSError, ValueError) as exc:
            messagebox.showerror("Error", f"No se pudo guardar el archivo:\n{exc}")
            self._set_estado("Error al exportar", "error")
            return
        finally:
            self.config(cursor="")
        self._set_estado(f"Exportado: {ruta}", "ok")
        if messagebox.askyesno("Exportar", f"Informe guardado ({len(lista)} prospectos):\n{ruta}\n\n¿Abrir la carpeta?"):
            abrir_ruta(INFORMES_DIR)

    def _on_cargar_todo(self) -> None:
        s = self._sesion
        if s is None or self._ocupado:
            messagebox.showinfo("Exportar", "Primero haz una búsqueda.")
            return
        faltan = s.estimar_peticiones_restantes()
        if faltan > 0:
            usar = min(faltan, TOPE_CARGAR_TODO)
            if not messagebox.askyesno(
                "Cargar todo y exportar",
                f"Para cargar todo el resultado faltan hasta {faltan} peticiones a Google "
                f"(≈ ${faltan * COSTO_ESTIMADO_POR_PETICION_USD:.2f} USD).\n"
                f"Esta operación usará como máximo {usar} (tope {TOPE_CARGAR_TODO}).\n\n¿Continuar?",
            ):
                return
        self._lanzar(
            lambda: s.cargar_todo(TOPE_CARGAR_TODO if faltan > 0 else 0),
            lambda res, exc: self._fin_cargar_todo(s, exc),
            "Cargando todos los resultados…",
        )

    def _fin_cargar_todo(self, s: SesionBusqueda, exc: Optional[BaseException]) -> None:
        if exc is not None:
            messagebox.showerror("Error", f"No se pudo completar la carga:\n{exc}")
            return
        if self._pagina_actual in s.paginas:
            self._render_busqueda()
        lista = s.todos_los_negocios()
        if s.error:
            messagebox.showwarning("Carga incompleta", f"Google devolvió un error; se exportará lo ya cargado.\n\n{s.error}")
        elif s.cancelado:
            self._set_estado("Carga cancelada — se exporta lo ya cargado", "aviso")
        elif not s.agotado:
            messagebox.showinfo("Tope alcanzado", f"Se alcanzó el tope de {TOPE_CARGAR_TODO} peticiones; se exporta lo cargado hasta ahora.")
        if lista:
            self._escribir_export(lista, "busqueda", "xlsx")
        else:
            messagebox.showinfo("Exportar", "No hay resultados para exportar.")

    # --------------------------------------------------------------- diálogos
    def _centrar_sobre(self, ventana: tk.Toplevel) -> None:
        ventana.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - ventana.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - ventana.winfo_height()) // 2
        ventana.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _mostrar_info_api(self) -> None:
        ventana = tk.Toplevel(self)
        ventana.title("Info — Uso de Google Places API")
        ventana.geometry("560x640")
        ventana.transient(self)
        ventana.configure(background=COLORES["fondo"])
        marco = ttk.Frame(ventana, padding=16)
        marco.pack(fill="both", expand=True)
        texto = tk.Text(marco, wrap="word", font=("Consolas", 10), relief="solid", bd=1, bg="#FFFFFF", padx=10, pady=10)
        texto.pack(fill="both", expand=True)

        def actualizar() -> None:
            texto.config(state="normal")
            texto.delete("1.0", tk.END)
            texto.insert("1.0", self._uso.resumen_texto() + f"\n\nBase de datos: {self._db.ruta}")
            texto.config(state="disabled")

        actualizar()
        barra = ttk.Frame(marco)
        barra.pack(fill="x", pady=(12, 0))
        ttk.Button(barra, text="Actualizar", command=actualizar).pack(side="left")
        ttk.Button(barra, text="Cerrar", command=ventana.destroy).pack(side="right")
        self._centrar_sobre(ventana)

    def _abrir_editor_palabras(self) -> None:
        if self._editor_palabras is not None and self._editor_palabras.winfo_exists():
            self._editor_palabras.lift()
            return
        v = tk.Toplevel(self)
        self._editor_palabras = v
        v.title("Palabras clave por tipo de negocio")
        v.geometry("680x640")
        v.transient(self)
        v.configure(background=COLORES["fondo"])
        marco = ttk.Frame(v, padding=16)
        marco.pack(fill="both", expand=True)
        marco.columnconfigure(0, weight=1)
        marco.rowconfigure(2, weight=1)

        fila = ttk.Frame(marco)
        fila.grid(row=0, column=0, sticky="ew")
        fila.columnconfigure(1, weight=1)
        ttk.Label(fila, text="Tipo de negocio").grid(row=0, column=0, padx=(0, 8))
        var_t = tk.StringVar(value=self.var_tipo.get())
        combo = ttk.Combobox(fila, textvariable=var_t, values=list(self._palabras), state="readonly")
        combo.grid(row=0, column=1, sticky="ew")
        ttk.Label(marco, text="Un término por línea. Cada término hace 1 consulta («término en ciudad, país»).",
                  style="Suave.TLabel").grid(row=1, column=0, sticky="w", pady=(10, 4))
        txt = tk.Text(marco, wrap="word", relief="solid", bd=1, padx=8, pady=8)
        txt.grid(row=2, column=0, sticky="nsew")
        lbl = ttk.Label(marco, text="", style="Suave.TLabel")
        lbl.grid(row=3, column=0, sticky="w", pady=(6, 0))

        def cargar(*_: object) -> None:
            txt.delete("1.0", "end")
            txt.insert("1.0", "\n".join(self._palabras.get(var_t.get(), [])))
            btn_elim.config(state="disabled" if var_t.get() in PALABRAS_CLAVE_DEFECTO else "normal")
            lbl.config(text=f"{len(self._palabras.get(var_t.get(), []))} términos guardados")

        def refrescar_combos() -> None:
            combo["values"] = list(self._palabras)
            self.combo_tipo["values"] = list(self._palabras)
            if self.var_tipo.get() not in self._palabras:
                self.var_tipo.set(next(iter(self._palabras)))
            self._actualizar_resumen_terminos()

        def persistir() -> bool:
            try:
                guardar_palabras_clave(self._ruta_palabras, self._palabras, self._json_invalido)
                self._json_invalido = False
                return True
            except OSError as exc:
                messagebox.showerror("Error", f"No se pudo guardar palabras_clave.json:\n{exc}", parent=v)
                return False

        def guardar() -> None:
            terminos = normalizar_terminos(txt.get("1.0", "end").splitlines())
            if not terminos:
                messagebox.showwarning("Palabras clave", "Escribe al menos un término.", parent=v)
                return
            self._palabras[var_t.get()] = terminos
            if persistir():
                refrescar_combos()
                cargar()
                lbl.config(text=f"Guardado: {len(terminos)} términos")

        def restaurar() -> None:
            defecto = PALABRAS_CLAVE_DEFECTO.get(var_t.get())
            if defecto:
                txt.delete("1.0", "end")
                txt.insert("1.0", "\n".join(defecto))
                lbl.config(text="Valores por defecto cargados — pulsa «Guardar» para aplicarlos")

        def nuevo() -> None:
            nombre = simpledialog.askstring("Nuevo tipo", "Nombre del nuevo tipo de negocio:", parent=v)
            nombre = re.sub(r"\s+", " ", nombre or "").strip()
            if not nombre:
                return
            if nombre in self._palabras:
                var_t.set(nombre)
                cargar()
                return
            self._palabras[nombre] = [nombre.lower()]
            if persistir():
                refrescar_combos()
                var_t.set(nombre)
                cargar()

        def eliminar() -> None:
            tipo = var_t.get()
            if tipo in PALABRAS_CLAVE_DEFECTO:
                return
            if messagebox.askyesno("Eliminar tipo", f"¿Eliminar el tipo «{tipo}» y sus términos?", parent=v):
                self._palabras.pop(tipo, None)
                if persistir():
                    refrescar_combos()
                    var_t.set(next(iter(self._palabras)))
                    cargar()

        def reiniciar_omitidos() -> None:
            self._db.reiniciar_terminos()
            lbl.config(text="Listo: se volverán a consultar los términos que se habían omitido")

        barra = ttk.Frame(marco)
        barra.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        ttk.Button(barra, text="Guardar", style="Primary.TButton", command=guardar).pack(side="left")
        ttk.Button(barra, text="Valores por defecto", command=restaurar).pack(side="left", padx=6)
        ttk.Button(barra, text="Nuevo tipo…", command=nuevo).pack(side="left")
        btn_elim = ttk.Button(barra, text="Eliminar tipo", command=eliminar)
        btn_elim.pack(side="left", padx=6)
        barra2 = ttk.Frame(marco)
        barra2.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(barra2, text="Volver a consultar términos omitidos (aportaban 0 nuevos)", command=reiniciar_omitidos).pack(side="left")
        ttk.Button(barra2, text="Cerrar", command=v.destroy).pack(side="right")
        combo.bind("<<ComboboxSelected>>", cargar)
        cargar()
        self._centrar_sobre(v)

    # ------------------------------------------------------- importar informes
    def _build_tab_informe(self) -> None:
        t = self.tab_i = ttk.Frame(self.nb, padding=(0, 10, 0, 0))
        t.columnconfigure(0, weight=1)
        t.rowconfigure(1, weight=1)
        tools = ttk.Frame(t)
        tools.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        tools.columnconfigure(3, weight=1)
        ttk.Label(tools, text="Filtrar").grid(row=0, column=0, padx=(0, 8))
        self.var_filtro_i = tk.StringVar()
        self.var_filtro_i.trace_add("write", lambda *_: self._on_filtro_informe())
        ttk.Entry(tools, textvariable=self.var_filtro_i, width=34).grid(row=0, column=1, sticky="w")
        self.lbl_inf_cuenta = ttk.Label(tools, text="", style="Negrita.TLabel", foreground=COLORES["ok"])
        self.lbl_inf_cuenta.grid(row=0, column=2, padx=14)
        ttk.Button(tools, text="Importar otro…", command=self._importar_informe).grid(row=0, column=4, padx=4)
        ttk.Button(tools, text="Cerrar informe", command=self._cerrar_informe).grid(row=0, column=5, padx=(4, 0))
        self.lbl_inf_archivo = ttk.Label(t, text="", style="Suave.TLabel")
        self.lbl_inf_archivo.grid(row=2, column=0, sticky="w", pady=(6, 0))

    def _importar_informe(self) -> None:
        self._cerrar_editores()
        carpeta = INFORMES_DIR if os.path.isdir(INFORMES_DIR) else os.path.expanduser("~")
        ruta = filedialog.askopenfilename(
            parent=self, title="Importar informe", initialdir=carpeta,
            filetypes=[("Informes (Excel o CSV)", "*.xlsx *.xlsm *.csv"), ("Todos los archivos", "*.*")],
        )
        if not ruta:
            return
        self.config(cursor="watch")
        self.update_idletasks()
        try:
            encabezados, filas, aviso = leer_informe(ruta)
        except Exception as exc:        # noqa: BLE001 — archivo dañado, bloqueado, formato raro…
            messagebox.showerror("Importar informe", f"No se pudo abrir el archivo:\n\n{exc}")
            self._set_estado("No se pudo importar el informe", "error")
            return
        finally:
            self.config(cursor="")
        self._mostrar_informe(ruta, encabezados, filas, aviso)

    def _mostrar_informe(self, ruta: str, encabezados: list[str], filas: list[list[str]], aviso: str) -> None:
        if self.tab_i is None:
            self._build_tab_informe()
        assert self.tab_i is not None
        if str(self.tab_i) not in self.nb.tabs():
            self.nb.add(self.tab_i, text="Informe")
        nombre = os.path.basename(ruta)
        self.nb.tab(self.tab_i, text="Informe: " + (nombre if len(nombre) <= 26 else nombre[:23] + "…"))
        idx_estado = next((i for i, h in enumerate(encabezados) if h.strip().lower() == "estado"), None)
        columnas: list[Columna] = []
        for i, titulo in enumerate(encabezados):
            valores = [f[i] for f in filas]
            no_vacios = [v for v in valores if v.strip()]
            numerica = bool(no_vacios) and all(re.fullmatch(r"-?\d+", v.strip()) for v in no_vacios)
            largo = max([0] + [len(v) for v in valores[:300]])
            columnas.append(Columna(
                f"c{i}", titulo, max(70, min(360, max(len(titulo) * 9 + 34, largo * 7 + 28))),
                "center" if (numerica or i == idx_estado) else "w", numerica=numerica,
            ))
        if self._inf_after is not None:
            self.after_cancel(self._inf_after)
            self._inf_after = None
        if self.tabla_i is not None:
            self.tabla_i.destroy()
        self.tabla_i = TablaDatos(self.tab_i, columnas, on_mensaje=self._mensaje_tabla)
        self.tabla_i.grid(row=1, column=0, sticky="nsew")
        self._inf_filas = [tuple(f) for f in filas]
        self._inf_texto = [" ".join(f).lower() for f in filas]
        self._inf_estados = [
            (f[idx_estado].strip().lower() if idx_estado is not None and f[idx_estado].strip().lower() in ESTADOS else "nuevo")
            for f in filas
        ]
        self._inf_aviso = aviso
        self._suprimir_filtro_i = True
        self.var_filtro_i.set("")
        self._suprimir_filtro_i = False
        self.lbl_inf_archivo.config(text=ruta)
        self._render_informe()
        self.nb.select(self.tab_i)
        total = f"{len(filas):,}".replace(",", ".")
        self._set_estado(f"Informe importado: {nombre} · {total} filas" + (f" · {aviso}" if aviso else ""), "aviso" if aviso else "ok")

    def _on_filtro_informe(self) -> None:
        if self._suprimir_filtro_i:
            return
        if self._inf_after is not None:
            self.after_cancel(self._inf_after)
        self._inf_after = self.after(250, self._render_informe)

    def _render_informe(self) -> None:
        self._inf_after = None
        if self.tabla_i is None:
            return
        termino = self.var_filtro_i.get().strip().lower()
        indices = [i for i, t in enumerate(self._inf_texto) if not termino or termino in t]
        self.tabla_i.set_filas([(str(i), self._inf_filas[i], self._inf_estados[i]) for i in indices])
        total = len(self._inf_filas)
        texto = f"{len(indices):,} de {total:,} filas · {len(self.tabla_i.columnas)} columnas".replace(",", ".")
        if self._inf_aviso:
            texto += f" · {self._inf_aviso}"
        self.lbl_inf_cuenta.config(text=texto, foreground=COLORES["ok"] if indices or not total else COLORES["error"])

    def _cerrar_informe(self) -> None:
        if self.tab_i is None:
            return
        if self._inf_after is not None:
            self.after_cancel(self._inf_after)
            self._inf_after = None
        if self.tabla_i is not None:
            self.tabla_i.destroy()
            self.tabla_i = None
        self._inf_filas, self._inf_estados, self._inf_texto, self._inf_aviso = [], [], [], ""
        if str(self.tab_i) in self.nb.tabs():
            self.nb.forget(self.tab_i)
        self.nb.select(self.tab_b)
        self._set_estado("Informe cerrado")

    # ------------------------------------------------------------------ cierre
    def _cerrar(self) -> None:
        try:
            self._guardar_nota_pendiente()
            self._cerrar_editores()
        finally:
            self._nota_id = None
            self._cancelar.set()
            try:
                self.progreso.stop()
            except tk.TclError:
                pass
            if self._after_id is not None:
                try:
                    self.after_cancel(self._after_id)
                except tk.TclError:
                    pass
            if self._hilo is not None and self._hilo.is_alive():
                self._hilo.join(timeout=3.0)
            self._db.cerrar()
            self.destroy()


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------

def main() -> None:
    try:
        app = ProspeccionApp()
    except Exception as exc:            # noqa: BLE001 — con pythonw un error al abrir sería invisible
        raiz = tk.Tk()
        raiz.withdraw()
        messagebox.showerror(
            "No se pudo iniciar",
            f"No se pudo abrir la base de datos o iniciar la aplicación:\n\n{exc}\n\n"
            f"Carpeta de datos: {_BASE_DIR}\n(Puedes cambiarla con la variable de entorno PROSPECCION_DATOS.)",
        )
        raiz.destroy()
        return
    app.mainloop()


if __name__ == "__main__":
    main()