#!/usr/bin/env python3
"""
Backend y utilidades para ProspeccionApp (app_v2).
Prefijo concatenado delante de la interfaz gráfica.
"""

from __future__ import annotations

import csv
import json
import os
import queue
import re
import sqlite3
import subprocess
import threading
import time
import urllib.error
import urllib.request
import webbrowser
from dataclasses import dataclass, field
from datetime import datetime
from tkinter import font as tkfont
from tkinter import messagebox, simpledialog, ttk
from typing import Callable, Optional
import tkinter as tk

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Rutas y constantes
# ---------------------------------------------------------------------------

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RUTA_DB = os.path.join(_BASE_DIR, "prospectos.db")
RUTA_PALABRAS = os.path.join(_BASE_DIR, "palabras_clave.json")
INFORMES_DIR = os.path.join(_BASE_DIR, "informes")

COLORES = {
    "fondo": "#F0F3F7",
    "tarjeta": "#FFFFFF",
    "texto": "#1F2933",
    "suave": "#6B7785",
    "borde": "#D0D7E2",
    "primario": "#1E6BB8",
    "primario_claro": "#E8F1FB",
    "primario_osc": "#155A9A",
    "sel": "#D6E8FA",
    "ok": "#1B7A4E",
    "aviso": "#B86E00",
    "error": "#C0392B",
}

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
    "Puerto Rico": "PR", "España": "ES", "Cuba": "CU",
}

MIN_RESEÑAS_OPCIONES = ["0", "5", "10", "20", "50", "100"]
OPCIONES_POR_PAGINA = ["10", "25", "50", "100", "200", "300", "400", "500", "600", "todos"]

PAGINA_TODOS = 9999
TOPE_PETICIONES_MAX = 80
TOPE_CARGAR_TODO = 40
COSTO_ESTIMADO_POR_PETICION_USD = 0.035

PALABRAS_CLAVE_DEFECTO: dict[str, list[str]] = {
    "Comidas rápidas": [
        "comida rápida", "fast food", "hamburguesas", "salchipapas", "arepas", "perros calientes",
        "pizza a domicilio", "pollo broaster", "comida corrida", "snacks", "empanadas", "tacos",
    ],
    "Restaurantes": [
        "restaurante", "restaurantes", "comida", "asadero", "cafetería", "pizzería",
        "marisquería", "sushi", "comida china", "parrilla", "menú del día", "brunch",
    ],
    "Barberías": [
        "barbería", "barber shop", "peluquería barbería", "corte de cabello", "barbería hombres",
        "barbería niños", "fade", "afeitado", "grooming",
    ],
    "Odontólogos": [
        "odontólogo", "clínica dental", "dentista", "odontología", "consultorio dental",
        "ortodoncia", "endodoncia", "implantes dentales", "blanqueamiento dental",
    ],
    "Médicos": [
        "consultorio médico", "médico consultorio", "medicina general", "médico general",
        "médico", "consulta médica", "médico particular", "especialista médico", "clínica médica",
    ],
    "Psicólogos": [
        "consultorio psicológico", "psicólogo consultorio", "psicólogo clínico", "consulta psicológica",
        "psicólogo", "psicóloga", "psicólogo particular", "psicóloga clínica", "consultorio de psicología",
    ],
    "Inmobiliarias": [
        "inmobiliaria", "bienes raíces", "apartamentos", "casas en venta", "arriendo apartamentos",
        "corretaje", "propiedades", "lotes en venta", "oficinas en arriendo",
    ],
}

COLS_BUSQUEDA = (
    ("fav", "★", 36),
    ("nombre", "Negocio", 210),
    ("ciudad", "Ciudad", 110),
    ("tipo", "Tipo", 110),
    ("telefono", "Teléfono", 130),
    ("resenas", "Reseñas", 70),
    ("web", "Web", 160),
    ("direccion", "Dirección", 220),
    ("estado", "Estado", 110),
)

COLS_GUARDADOS = (
    ("fav", "★", 36),
    ("nombre", "Negocio", 200),
    ("ciudad", "Ciudad", 100),
    ("tipo", "Tipo", 100),
    ("telefono", "Teléfono", 120),
    ("estado", "Estado", 110),
    ("guardado_en", "Guardado", 130),
    ("web", "Web", 150),
    ("direccion", "Dirección", 200),
)

ESTADOS = ("nuevo", "contactado", "interesado", "no_interesado", "cliente")
CAMPOS_EDITABLES = ["nombre", "telefono", "resenas", "web", "direccion"]
CAMPO_DE_COLUMNA = {
    "nombre": "nombre",
    "telefono": "telefono",
    "resenas": "resenas",
    "web": "web",
    "direccion": "direccion",
}

_COLUMNAS_EXPORT = (
    "Negocio",
    "Ciudad",
    "Tipo",
    "Teléfono",
    "Reseñas",
    "Web",
    "Link Google Maps",
    "Dirección",
    "Estado",
    "Notas",
    "Guardado",
)

_MENSAJES_POR_TIPO: dict[str, str] = {
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

_COLORES_ESTADO = {
    "nuevo": "#1F2933",
    "contactado": "#1E6BB8",
    "interesado": "#1B7A4E",
    "no_interesado": "#8A6D3B",
    "cliente": "#6B3FA0",
}


# ---------------------------------------------------------------------------
# Modelo
# ---------------------------------------------------------------------------

@dataclass
class Negocio:
    place_id: str
    nombre: str
    ciudad: str
    tipo: str
    telefono: str = ""
    web: str = ""
    resenas: int = 0
    link_maps: str = ""
    direccion: str = ""
    estado: str = "nuevo"
    notas: str = ""
    guardado: bool = False
    guardado_en: str = ""
    actividad_reciente: bool = True


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _leer_google_api_key() -> str:
    key = os.environ.get("GOOGLE_MAPS_API_KEY", "").strip()
    if key:
        return key
    env_path = os.path.join(_BASE_DIR, ".env")
    if os.path.isfile(env_path):
        with open(env_path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if line.startswith("GOOGLE_MAPS_API_KEY="):
                    valor = line.split("=", 1)[1].strip().strip('"').strip("'")
                    if valor:
                        return valor
    raise ValueError(
        "Falta la API key de Google Maps.\n"
        "Crea un archivo .env con GOOGLE_MAPS_API_KEY=tu_clave\n"
        "o define la variable de entorno GOOGLE_MAPS_API_KEY.\n"
        "Activa Places API (New) en Google Cloud Console."
    )


def cargar_palabras_clave(ruta: str) -> tuple[dict[str, list[str]], Optional[str], bool]:
    defecto = {k: list(v) for k, v in PALABRAS_CLAVE_DEFECTO.items()}
    if not os.path.isfile(ruta):
        return defecto, None, False
    try:
        with open(ruta, encoding="utf-8") as f:
            datos = json.load(f)
    except (json.JSONDecodeError, OSError, UnicodeDecodeError) as exc:
        return defecto, f"palabras_clave.json inválido ({exc}); se usan valores por defecto", True
    if not isinstance(datos, dict):
        return defecto, "palabras_clave.json no es un objeto; se usan valores por defecto", True
    resultado: dict[str, list[str]] = {}
    for clave, valor in datos.items():
        if not isinstance(clave, str):
            continue
        if isinstance(valor, list):
            terminos = normalizar_terminos([str(x) for x in valor])
            if terminos:
                resultado[clave] = terminos
        elif isinstance(valor, str) and valor.strip():
            resultado[clave] = [valor.strip()]
    if not resultado:
        return defecto, "palabras_clave.json vacío; se usan valores por defecto", True
    return resultado, None, False


def guardar_palabras_clave(ruta: str, palabras: dict[str, list[str]], json_invalido: bool) -> None:
    if json_invalido and os.path.isfile(ruta):
        respaldo = ruta + ".bak"
        try:
            with open(ruta, "rb") as origen, open(respaldo, "wb") as destino:
                destino.write(origen.read())
        except OSError:
            pass
    limpio = {k: normalizar_terminos(v) for k, v in palabras.items() if normalizar_terminos(v)}
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(limpio, f, ensure_ascii=False, indent=2)


def normalizar_terminos(lines) -> list[str]:
    vistos: set[str] = set()
    salida: list[str] = []
    for line in lines:
        termino = re.sub(r"\s+", " ", str(line)).strip()
        if not termino:
            continue
        clave = termino.casefold()
        if clave in vistos:
            continue
        vistos.add(clave)
        salida.append(termino)
    return salida


def _texto_busqueda(n: Negocio) -> str:
    partes = (
        n.nombre, n.ciudad, n.tipo, n.telefono, n.web, n.direccion,
        n.estado, n.notas, n.guardado_en, str(n.resenas),
    )
    return " ".join(partes).lower()


def _valores_busqueda(n: Negocio) -> tuple:
    return (
        "★" if n.guardado else "☆",
        n.nombre,
        n.ciudad,
        n.tipo,
        n.telefono,
        str(n.resenas),
        n.web,
        n.direccion,
        n.estado,
    )


def _valores_guardados(n: Negocio) -> tuple:
    return (
        "★",
        n.nombre,
        n.ciudad,
        n.tipo,
        n.telefono,
        n.estado,
        n.guardado_en,
        n.web,
        n.direccion,
    )


def mensaje_para(n: Negocio) -> str:
    plantilla = _MENSAJES_POR_TIPO.get(
        n.tipo,
        "Hola, vi {nombre} y creo que podrías automatizar tu atención por WhatsApp.",
    )
    return plantilla.format(nombre=n.nombre)


def abrir_url(url: str) -> bool:
    url = (url or "").strip()
    if not url:
        return False
    if not re.match(r"^https?://", url, flags=re.I):
        if url.startswith("www."):
            url = "https://" + url
        else:
            return False
    try:
        webbrowser.open(url)
        return True
    except Exception:
        return False


def abrir_ruta(path: str) -> None:
    path = os.path.abspath(path)
    if os.name == "nt":
        os.startfile(path)  # type: ignore[attr-defined]
        return
    try:
        subprocess.Popen(["xdg-open", path])
    except OSError:
        webbrowser.open(path)


def ruta_unica(path: str) -> str:
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    n = 2
    while True:
        candidato = f"{base} ({n}){ext}"
        if not os.path.exists(candidato):
            return candidato
        n += 1


def _sanitizar_parte(parte: str) -> str:
    invalidos = '<>:"/\\|?*'
    s = str(parte).strip().replace(" ", "-")
    for c in invalidos:
        s = s.replace(c, "")
    s = re.sub(r"-{2,}", "-", s).strip("-")
    return s or "sin-nombre"


def nombre_exportacion(partes, fmt: str) -> str:
    fecha = datetime.now().strftime("%d-%m-%Y-%H-%M")
    cuerpo = "-".join(_sanitizar_parte(p) for p in partes if str(p).strip())
    ext = "xlsx" if fmt.lower() in ("xlsx", "excel") else "csv"
    return f"{cuerpo}-{fecha}.{ext}" if cuerpo else f"exportacion-{fecha}.{ext}"


def exportar_excel(ruta: str, lista: list[Negocio]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Prospectos"
    borde_fino = Side(style="thin", color="D0D0D0")
    borde = Border(left=borde_fino, right=borde_fino, top=borde_fino, bottom=borde_fino)
    relleno_cabecera = PatternFill(fill_type="solid", fgColor="1E6BB8")
    fuente_cabecera = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    relleno_alt = PatternFill(fill_type="solid", fgColor="EEF4FA")
    alin_cab = Alignment(horizontal="center", vertical="center", wrap_text=True)
    alin_dat = Alignment(vertical="center", wrap_text=True)

    for col, titulo in enumerate(_COLUMNAS_EXPORT, start=1):
        celda = ws.cell(row=1, column=col, value=titulo)
        celda.fill = relleno_cabecera
        celda.font = fuente_cabecera
        celda.alignment = alin_cab
        celda.border = borde

    for fila_idx, n in enumerate(lista, start=2):
        valores = (
            n.nombre, n.ciudad, n.tipo, n.telefono, n.resenas, n.web,
            n.link_maps, n.direccion, n.estado, n.notas,
            "sí" if n.guardado else "no",
        )
        zebra = fila_idx % 2 == 0
        for col, valor in enumerate(valores, start=1):
            celda = ws.cell(row=fila_idx, column=col, value=valor)
            celda.border = borde
            celda.alignment = alin_dat
            if zebra:
                celda.fill = relleno_alt

    for col in range(1, len(_COLUMNAS_EXPORT) + 1):
        letra = get_column_letter(col)
        ancho = len(str(_COLUMNAS_EXPORT[col - 1]))
        for fila_idx in range(2, ws.max_row + 1):
            valor = ws.cell(row=fila_idx, column=col).value
            if valor is not None:
                ancho = max(ancho, len(str(valor)))
        ws.column_dimensions[letra].width = min(ancho + 3, 55)

    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"
    if lista:
        ultima = get_column_letter(len(_COLUMNAS_EXPORT))
        ws.auto_filter.ref = f"A1:{ultima}{ws.max_row}"
    wb.save(ruta)


def exportar_csv(ruta: str, lista: list[Negocio]) -> None:
    with open(ruta, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(_COLUMNAS_EXPORT)
        for n in lista:
            writer.writerow([
                n.nombre, n.ciudad, n.tipo, n.telefono, n.resenas, n.web,
                n.link_maps, n.direccion, n.estado, n.notas,
                "sí" if n.guardado else "no",
            ])


def _clave_termino(pais: str, ciudad: str, termino: str) -> str:
    return f"{pais.casefold()}|{ciudad.casefold()}|{termino.casefold()}"


# ---------------------------------------------------------------------------
# Base de datos
# ---------------------------------------------------------------------------

class BaseDatos:
    """SQLite thread-safe para prospectos, uso de API y términos omitidos."""

    def __init__(self, ruta: str) -> None:
        self._ruta = ruta
        self._lock = threading.Lock()
        self._conn = sqlite3.connect(ruta, check_same_thread=False)
        self._conn.row_factory = sqlite3.Row
        with self._lock:
            self._crear_esquema()

    @property
    def ruta(self) -> str:
        return self._ruta

    def _crear_esquema(self) -> None:
        cur = self._conn.cursor()
        cur.executescript(
            """
            CREATE TABLE IF NOT EXISTS prospectos (
                place_id TEXT PRIMARY KEY,
                nombre TEXT NOT NULL DEFAULT '',
                ciudad TEXT NOT NULL DEFAULT '',
                tipo TEXT NOT NULL DEFAULT '',
                telefono TEXT NOT NULL DEFAULT '',
                web TEXT NOT NULL DEFAULT '',
                resenas INTEGER NOT NULL DEFAULT 0,
                link_maps TEXT NOT NULL DEFAULT '',
                direccion TEXT NOT NULL DEFAULT '',
                estado TEXT NOT NULL DEFAULT 'nuevo',
                notas TEXT NOT NULL DEFAULT '',
                guardado INTEGER NOT NULL DEFAULT 0,
                guardado_en TEXT NOT NULL DEFAULT '',
                actividad INTEGER NOT NULL DEFAULT 1,
                edit_nombre INTEGER NOT NULL DEFAULT 0,
                edit_telefono INTEGER NOT NULL DEFAULT 0,
                edit_resenas INTEGER NOT NULL DEFAULT 0,
                edit_web INTEGER NOT NULL DEFAULT 0,
                edit_direccion INTEGER NOT NULL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS api_uso (
                mes TEXT PRIMARY KEY,
                peticiones INTEGER NOT NULL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS terminos_omitidos (
                clave TEXT PRIMARY KEY
            );
            """
        )
        self._conn.commit()

    def _mes_actual(self) -> str:
        return time.strftime("%Y-%m")

    def uso_mes(self) -> int:
        with self._lock:
            fila = self._conn.execute(
                "SELECT peticiones FROM api_uso WHERE mes = ?",
                (self._mes_actual(),),
            ).fetchone()
            return int(fila["peticiones"]) if fila else 0

    def registrar_uso(self) -> None:
        with self._lock:
            mes = self._mes_actual()
            self._conn.execute(
                """
                INSERT INTO api_uso (mes, peticiones) VALUES (?, 1)
                ON CONFLICT(mes) DO UPDATE SET peticiones = peticiones + 1
                """,
                (mes,),
            )
            self._conn.commit()

    def _fila_a_negocio(self, fila: sqlite3.Row) -> Negocio:
        return Negocio(
            place_id=fila["place_id"],
            nombre=fila["nombre"],
            ciudad=fila["ciudad"],
            tipo=fila["tipo"],
            telefono=fila["telefono"],
            web=fila["web"],
            resenas=int(fila["resenas"] or 0),
            link_maps=fila["link_maps"],
            direccion=fila["direccion"],
            estado=fila["estado"] or "nuevo",
            notas=fila["notas"] or "",
            guardado=bool(fila["guardado"]),
            guardado_en=fila["guardado_en"] or "",
            actividad_reciente=bool(fila["actividad"]),
        )

    def sincronizar(self, negocios: list[Negocio]) -> None:
        if not negocios:
            return
        with self._lock:
            for n in negocios:
                fila = self._conn.execute(
                    "SELECT * FROM prospectos WHERE place_id = ?",
                    (n.place_id,),
                ).fetchone()
                if fila is None:
                    self._conn.execute(
                        """
                        INSERT INTO prospectos (
                            place_id, nombre, ciudad, tipo, telefono, web, resenas,
                            link_maps, direccion, estado, notas, guardado, guardado_en, actividad
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            n.place_id, n.nombre, n.ciudad, n.tipo, n.telefono, n.web, int(n.resenas),
                            n.link_maps, n.direccion, n.estado or "nuevo", n.notas or "",
                            1 if n.guardado else 0, n.guardado_en or "",
                            1 if n.actividad_reciente else 0,
                        ),
                    )
                else:
                    nombre = fila["nombre"] if fila["edit_nombre"] else n.nombre
                    telefono = fila["telefono"] if fila["edit_telefono"] else n.telefono
                    resenas = fila["resenas"] if fila["edit_resenas"] else int(n.resenas)
                    web = fila["web"] if fila["edit_web"] else n.web
                    direccion = fila["direccion"] if fila["edit_direccion"] else n.direccion
                    self._conn.execute(
                        """
                        UPDATE prospectos SET
                            nombre = ?, ciudad = ?, tipo = ?, telefono = ?, web = ?,
                            resenas = ?, link_maps = ?, direccion = ?, actividad = ?
                        WHERE place_id = ?
                        """,
                        (
                            nombre, n.ciudad, n.tipo, telefono, web, int(resenas),
                            n.link_maps, direccion,
                            1 if n.actividad_reciente else 0, n.place_id,
                        ),
                    )
                    n.nombre = nombre
                    n.telefono = telefono
                    n.resenas = int(resenas)
                    n.web = web
                    n.direccion = direccion
                    n.estado = fila["estado"] or "nuevo"
                    n.notas = fila["notas"] or ""
                    n.guardado = bool(fila["guardado"])
                    n.guardado_en = fila["guardado_en"] or ""
            self._conn.commit()

    def set_guardado(self, place_id: str, valor: bool) -> Optional[str]:
        with self._lock:
            fila = self._conn.execute(
                "SELECT place_id FROM prospectos WHERE place_id = ?",
                (place_id,),
            ).fetchone()
            if fila is None:
                return None
            guardado_en = datetime.now().strftime("%Y-%m-%d %H:%M") if valor else ""
            self._conn.execute(
                "UPDATE prospectos SET guardado = ?, guardado_en = ? WHERE place_id = ?",
                (1 if valor else 0, guardado_en, place_id),
            )
            self._conn.commit()
            return guardado_en

    def set_estado(self, place_id: str, estado: str) -> bool:
        with self._lock:
            cur = self._conn.execute(
                "UPDATE prospectos SET estado = ? WHERE place_id = ?",
                (estado, place_id),
            )
            self._conn.commit()
            return cur.rowcount > 0

    def set_edicion(self, n: Negocio, campo: str, valor) -> None:
        if campo not in CAMPOS_EDITABLES:
            return
        col_edit = f"edit_{campo}"
        with self._lock:
            self._conn.execute(
                f"UPDATE prospectos SET {campo} = ?, {col_edit} = 1 WHERE place_id = ?",
                (valor, n.place_id),
            )
            self._conn.commit()

    def listar_guardados(self) -> list[Negocio]:
        with self._lock:
            filas = self._conn.execute(
                "SELECT * FROM prospectos WHERE guardado = 1 ORDER BY guardado_en DESC, nombre COLLATE NOCASE"
            ).fetchall()
            return [self._fila_a_negocio(f) for f in filas]

    def set_notas(self, place_id: str, texto: str) -> bool:
        with self._lock:
            cur = self._conn.execute(
                "UPDATE prospectos SET notas = ? WHERE place_id = ?",
                (texto, place_id),
            )
            self._conn.commit()
            return cur.rowcount > 0

    def termino_omitido(self, clave: str) -> bool:
        with self._lock:
            fila = self._conn.execute(
                "SELECT 1 FROM terminos_omitidos WHERE clave = ?",
                (clave,),
            ).fetchone()
            return fila is not None

    def marcar_omitido(self, clave: str) -> None:
        with self._lock:
            self._conn.execute(
                "INSERT OR IGNORE INTO terminos_omitidos (clave) VALUES (?)",
                (clave,),
            )
            self._conn.commit()

    def reiniciar_terminos(self) -> None:
        with self._lock:
            self._conn.execute("DELETE FROM terminos_omitidos")
            self._conn.commit()

    def cerrar(self) -> None:
        with self._lock:
            try:
                self._conn.close()
            except sqlite3.Error:
                pass


# ---------------------------------------------------------------------------
# Uso de API
# ---------------------------------------------------------------------------

class ApiUsageTracker:
    CREDITO_MENSUAL_USD = 200.0

    def __init__(self, db: BaseDatos) -> None:
        self._db = db
        self.peticiones_sesion = 0

    def registrar_peticion(self) -> None:
        self.peticiones_sesion += 1
        self._db.registrar_uso()

    def resumen_texto(self) -> str:
        mes = time.strftime("%Y-%m")
        peticiones_mes = self._db.uso_mes()
        costo_mes = peticiones_mes * COSTO_ESTIMADO_POR_PETICION_USD
        credito_restante = max(0.0, self.CREDITO_MENSUAL_USD - costo_mes)
        uso_pct = min(100.0, (costo_mes / self.CREDITO_MENSUAL_USD) * 100) if self.CREDITO_MENSUAL_USD else 0.0
        return (
            "Uso de Google Places API (Text Search)\n"
            "────────────────────────────────────\n\n"
            f"Sesión actual:              {self.peticiones_sesion} peticiones\n"
            f"Este mes ({mes}):           {peticiones_mes} peticiones\n\n"
            "Estimado de costo\n"
            "────────────────────────────────────\n"
            f"Costo estimado este mes:    ${costo_mes:.2f} USD\n"
            f"Crédito mensual referencia: ${self.CREDITO_MENSUAL_USD:.0f} USD\n"
            f"Crédito restante estimado:  ${credito_restante:.2f} USD\n"
            f"Uso estimado del crédito:   {uso_pct:.1f}%\n\n"
            "Notas\n"
            "────────────────────────────────────\n"
            "• Cada consulta a Google Maps cuenta como 1 petición.\n"
            "• El costo real depende de tu plan en Google Cloud.\n"
            "• Google no expone la cuota exacta restante con solo la API key.\n"
            "• Revisa el uso real en:\n"
            "  console.cloud.google.com → APIs y servicios → Panel"
        )


# ---------------------------------------------------------------------------
# Cliente Google Places (New)
# ---------------------------------------------------------------------------

class ClienteGooglePlaces:
    _FIELD_MASK = (
        "places.id,places.displayName,places.formattedAddress,"
        "places.internationalPhoneNumber,places.nationalPhoneNumber,"
        "places.websiteUri,places.userRatingCount,places.businessStatus,"
        "places.googleMapsUri"
    )

    def __init__(
        self,
        api_key: str,
        uso: ApiUsageTracker,
        cancelar: threading.Event,
    ) -> None:
        self._api_key = api_key
        self._uso = uso
        self._cancelar = cancelar

    def _request(self, body: dict) -> dict:
        if self._cancelar.is_set():
            raise RuntimeError("Búsqueda cancelada")
        req = urllib.request.Request(
            "https://places.googleapis.com/v1/places:searchText",
            data=json.dumps(body).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "X-Goog-Api-Key": self._api_key,
                "X-Goog-FieldMask": self._FIELD_MASK,
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detalle = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"Google Places API error ({exc.code}): {detalle}") from exc
        self._uso.registrar_peticion()
        return data

    def _place_a_negocio(self, place: dict, ciudad: str, tipo: str) -> Optional[Negocio]:
        place_id = (place.get("id") or "").strip()
        nombre = ((place.get("displayName") or {}).get("text") or "").strip()
        if not place_id or not nombre:
            return None
        telefono = (
            place.get("internationalPhoneNumber")
            or place.get("nationalPhoneNumber")
            or ""
        ).strip()
        web = (place.get("websiteUri") or "").strip()
        resenas = int(place["userRatingCount"]) if "userRatingCount" in place else 0
        status = place.get("businessStatus", "")
        actividad = status != "CLOSED_PERMANENTLY"
        link_maps = (place.get("googleMapsUri") or "").strip()
        if not link_maps and place_id:
            link_maps = f"https://www.google.com/maps/place/?q=place_id:{place_id}"
        direccion = (place.get("formattedAddress") or "").strip()
        return Negocio(
            place_id=place_id,
            nombre=nombre,
            ciudad=ciudad,
            tipo=tipo,
            telefono=telefono,
            web=web,
            resenas=resenas,
            link_maps=link_maps,
            direccion=direccion,
            actividad_reciente=actividad,
        )

    def buscar_consulta(
        self,
        text_query: str,
        region: str,
        ciudad: str = "",
        tipo: str = "",
    ) -> list[Negocio]:
        """Hasta 3 páginas × 20 resultados para una consulta de texto."""
        resultados: list[Negocio] = []
        vistos: set[str] = set()
        page_token: Optional[str] = None
        for pagina in range(3):
            if self._cancelar.is_set():
                break
            body: dict = {
                "textQuery": text_query,
                "languageCode": "es",
                "regionCode": region or "CO",
                "maxResultCount": 20,
            }
            if page_token:
                body["pageToken"] = page_token
            data = self._request(body)
            places = data.get("places") or []
            if not places:
                break
            for place in places:
                negocio = self._place_a_negocio(place, ciudad, tipo)
                if negocio is None or negocio.place_id in vistos:
                    continue
                vistos.add(negocio.place_id)
                resultados.append(negocio)
            page_token = data.get("nextPageToken")
            if not page_token:
                break
            if pagina < 2:
                time.sleep(0.35)
        return resultados


# ---------------------------------------------------------------------------
# Sesión de búsqueda paginada
# ---------------------------------------------------------------------------

class SesionBusqueda:
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
        self.pais = pais
        self.ciudad = ciudad
        self.tipo = tipo
        self.terminos = list(terminos)
        self.min_resenas = max(0, int(min_resenas))
        self.limite = max(1, int(limite))
        self.progreso = progreso or (lambda _t: None)
        self.paginas: dict[int, list[Negocio]] = {}
        self.raw: dict[str, Negocio] = {}
        self.omitidos: set[str] = set()
        self.error: Optional[str] = None
        self.cancelado = False
        self.peticiones = 0
        self.agotado = False
        self._cola: list[Negocio] = []
        self._idx_termino = 0
        self._orden: list[str] = []
        self._region = _PAIS_REGION.get(pais, "CO")
        for termino in self.terminos:
            clave = _clave_termino(pais, ciudad, termino)
            if db.termino_omitido(clave):
                self.omitidos.add(termino)

    def _cumple(self, n: Negocio) -> bool:
        return bool(n.actividad_reciente) and int(n.resenas) >= self.min_resenas

    def _avisar(self, texto: str) -> None:
        try:
            self.progreso(texto)
        except Exception:
            pass

    def _peticiones_cliente(self) -> int:
        return self.cliente._uso.peticiones_sesion

    def _siguiente_termino(self) -> Optional[str]:
        while self._idx_termino < len(self.terminos):
            termino = self.terminos[self._idx_termino]
            self._idx_termino += 1
            if termino in self.omitidos:
                continue
            return termino
        return None

    def _incorporar(self, negocios: list[Negocio]) -> int:
        nuevos = 0
        pendientes_sync: list[Negocio] = []
        for n in negocios:
            if n.place_id in self.raw:
                continue
            self.raw[n.place_id] = n
            self._orden.append(n.place_id)
            pendientes_sync.append(n)
            if self._cumple(n):
                self._cola.append(n)
                nuevos += 1
        if pendientes_sync:
            self.db.sincronizar(pendientes_sync)
        return nuevos

    def construir_pagina(self, permitir_api: bool = True) -> list[Negocio]:
        if self.cliente._cancelar.is_set():
            self.cancelado = True
        pagina: list[Negocio] = []
        while len(pagina) < self.limite and self._cola:
            pagina.append(self._cola.pop(0))

        while (
            permitir_api
            and len(pagina) < self.limite
            and not self.agotado
            and not self.cliente._cancelar.is_set()
        ):
            termino = self._siguiente_termino()
            if termino is None:
                self.agotado = True
                break
            consulta = f"{termino} en {self.ciudad}, {self.pais}"
            self._avisar(f"Consultando «{termino}» en {self.ciudad}…")
            antes = self._peticiones_cliente()
            try:
                hallados = self.cliente.buscar_consulta(
                    consulta, self._region, ciudad=self.ciudad, tipo=self.tipo,
                )
            except RuntimeError as exc:
                if self.cliente._cancelar.is_set() or "cancelad" in str(exc).lower():
                    self.cancelado = True
                    break
                self.error = str(exc)
                break
            except Exception as exc:  # noqa: BLE001
                self.error = str(exc)
                break
            self.peticiones += max(0, self._peticiones_cliente() - antes)
            nuevos_ids = 0
            for n in hallados:
                if n.place_id not in self.raw:
                    nuevos_ids += 1
            self._incorporar(hallados)
            if nuevos_ids == 0:
                clave = _clave_termino(self.pais, self.ciudad, termino)
                self.db.marcar_omitido(clave)
                self.omitidos.add(termino)
            while len(pagina) < self.limite and self._cola:
                pagina.append(self._cola.pop(0))
            time.sleep(0.25)

        if self.cliente._cancelar.is_set():
            self.cancelado = True

        if not pagina:
            if not self._cola and (self.agotado or self._idx_termino >= len(self.terminos)):
                self.agotado = True
            return []

        numero = len(self.paginas) + 1
        self.paginas[numero] = pagina
        return pagina

    def hay_mas(self) -> bool:
        if self._cola:
            return True
        if self.agotado:
            return False
        restantes = [
            t for t in self.terminos[self._idx_termino :]
            if t not in self.omitidos
        ]
        return bool(restantes)

    def reconfigurar(self, min_resenas: int, limite: int) -> None:
        self.min_resenas = max(0, int(min_resenas))
        self.limite = max(1, int(limite))
        self.paginas = {}
        self._cola = []
        for pid in self._orden:
            n = self.raw.get(pid)
            if n is not None and self._cumple(n):
                self._cola.append(n)

    def todos_los_negocios(self) -> list[Negocio]:
        return [self.raw[pid] for pid in self._orden if pid in self.raw and self._cumple(self.raw[pid])]

    def estimar_peticiones_restantes(self) -> int:
        restantes = [
            t for t in self.terminos[self._idx_termino :]
            if t not in self.omitidos
        ]
        return len(restantes) * 3

    def cargar_todo(self, tope_peticiones: int) -> None:
        inicio = self.peticiones
        while not self.agotado and not self.cliente._cancelar.is_set():
            if tope_peticiones > 0 and (self.peticiones - inicio) >= tope_peticiones:
                break
            if not self.hay_mas() and not self._cola:
                self.agotado = True
                break
            pagina = self.construir_pagina(permitir_api=True)
            if not pagina and (self.agotado or self.error or self.cancelado):
                break
            if not pagina and not self.hay_mas():
                self.agotado = True
                break


# ---------------------------------------------------------------------------
# Tabla editable
# ---------------------------------------------------------------------------

class TablaDatos(ttk.Frame):
    def __init__(
        self,
        parent,
        cols,
        on_estrella: Optional[Callable[[str], None]] = None,
        on_editar: Optional[Callable[[str, str, str], None]] = None,
        on_contexto: Optional[Callable[[tk.Event], None]] = None,
        on_doble_estado: Optional[Callable[[str, int, int], None]] = None,
        on_seleccion: Optional[Callable[[], None]] = None,
        on_mensaje: Optional[Callable[[str], None]] = None,
    ) -> None:
        super().__init__(parent)
        self.cols = list(cols)
        self.col_ids = [c[0] for c in self.cols]
        self.on_estrella = on_estrella
        self.on_editar = on_editar
        self.on_contexto = on_contexto
        self.on_doble_estado = on_doble_estado
        self.on_seleccion = on_seleccion
        self.on_mensaje = on_mensaje
        self._editor: Optional[ttk.Entry] = None
        self._edit_iid: Optional[str] = None
        self._edit_col: Optional[str] = None
        self._estados: dict[str, str] = {}

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            self,
            columns=self.col_ids,
            show="headings",
            style="Datos.Treeview",
            selectmode="extended",
        )
        ys = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")

        for col_id, heading, width in self.cols:
            self.tree.heading(col_id, text=heading, anchor="w")
            stretch = col_id not in ("fav", "resenas")
            self.tree.column(col_id, width=width, minwidth=28, stretch=stretch, anchor="w")

        for estado, color in _COLORES_ESTADO.items():
            self.tree.tag_configure(estado, foreground=color)

        self.tree.bind("<Button-1>", self._on_click, add=True)
        self.tree.bind("<Double-1>", self._on_doble)
        self.tree.bind("<Button-3>", self._on_right)
        self.tree.bind("<<TreeviewSelect>>", self._on_sel)
        self.tree.bind("<Control-c>", lambda _e: self._copiar())
        self.tree.bind("<Escape>", lambda _e: (self.cerrar_editor(False), "break")[1])

    def _avisar(self, texto: str) -> None:
        if self.on_mensaje:
            self.on_mensaje(texto)

    def _on_sel(self, _event=None) -> None:
        if self.on_seleccion:
            self.on_seleccion()

    def _identificar(self, event) -> tuple[Optional[str], Optional[str]]:
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return None, None
        fila = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not fila or not col:
            return None, None
        try:
            idx = int(col.replace("#", "")) - 1
        except ValueError:
            return None, None
        if idx < 0 or idx >= len(self.col_ids):
            return None, None
        return fila, self.col_ids[idx]

    def _on_click(self, event) -> None:
        iid, col = self._identificar(event)
        if iid and col == "fav" and self.on_estrella:
            self.tree.selection_set(iid)
            self.on_estrella(iid)
            return "break"

    def _on_doble(self, event) -> None:
        iid, col = self._identificar(event)
        if not iid or not col:
            return
        self.tree.selection_set(iid)
        if col == "estado" and self.on_doble_estado:
            self.on_doble_estado(iid, event.x_root, event.y_root)
            return "break"
        if col in CAMPO_DE_COLUMNA:
            self._abrir_editor(iid, col)
            return "break"

    def _on_right(self, event) -> None:
        iid = self.tree.identify_row(event.y)
        if iid:
            if iid not in self.tree.selection():
                self.tree.selection_set(iid)
            self.tree.focus(iid)
        if self.on_contexto:
            self.on_contexto(event)

    def _abrir_editor(self, iid: str, col_id: str) -> None:
        self.cerrar_editor(True)
        bbox = self.tree.bbox(iid, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        valor = ""
        try:
            valores = self.tree.item(iid, "values")
            idx = self.col_ids.index(col_id)
            valor = str(valores[idx]) if idx < len(valores) else ""
        except (ValueError, IndexError, tk.TclError):
            valor = ""
        entry = ttk.Entry(self.tree)
        entry.insert(0, valor)
        entry.select_range(0, "end")
        entry.focus_set()
        entry.place(x=x, y=y, width=max(w, 60), height=h)
        self._editor = entry
        self._edit_iid = iid
        self._edit_col = col_id

        def confirmar(_e=None) -> None:
            self.cerrar_editor(True)

        def cancelar(_e=None) -> None:
            self.cerrar_editor(False)

        entry.bind("<Return>", confirmar)
        entry.bind("<FocusOut>", confirmar)
        entry.bind("<Escape>", cancelar)

    def cerrar_editor(self, guardar: bool) -> None:
        if self._editor is None:
            return
        entry = self._editor
        iid, col = self._edit_iid, self._edit_col
        valor = entry.get()
        self._editor = None
        self._edit_iid = None
        self._edit_col = None
        try:
            entry.destroy()
        except tk.TclError:
            pass
        if guardar and iid and col and self.on_editar:
            self.on_editar(iid, col, valor)

    def set_filas(self, filas: list[tuple]) -> None:
        self.cerrar_editor(False)
        self.tree.delete(*self.tree.get_children())
        self._estados.clear()
        for iid, valores, estado in filas:
            self._estados[iid] = estado
            self.tree.insert("", "end", iid=iid, values=valores, tags=(estado,))

    def actualizar_fila(self, iid: str, valores, estado: str) -> None:
        if not self.tree.exists(iid):
            return
        self._estados[iid] = estado
        self.tree.item(iid, values=valores, tags=(estado,))

    def seleccion(self) -> list[str]:
        return list(self.tree.selection())

    def iids(self) -> list[str]:
        return list(self.tree.get_children())

    def _copiar(self) -> None:
        foco = self.tree.focus()
        sel = self.seleccion()
        if not sel and not foco:
            return
        textos: list[str] = []
        for iid in sel or ([foco] if foco else []):
            valores = self.tree.item(iid, "values")
            textos.append("\t".join(str(v) for v in valores))
        texto = "\n".join(textos)
        try:
            root = self.winfo_toplevel()
            root.clipboard_clear()
            root.clipboard_append(texto)
            self._avisar("Copiado al portapapeles")
        except tk.TclError:
            pass

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

        self._db = BaseDatos(ruta_db)
        self._uso = ApiUsageTracker(self._db)
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

        self._configurar_estilos()
        self._build_ui()
        self._centrar_ventana()
        self.protocol("WM_DELETE_WINDOW", self._cerrar)
        self.bind("<F5>", lambda _e: self._on_buscar())
        self.bind("<Control-f>", lambda _e: self.entry_filtro_b.focus_set())
        self.after(100, self._procesar_cola)
        self._refrescar_guardados()
        self._actualizar_botones()
        self._actualizar_uso()
        if self._aviso_json:
            self._set_estado(self._aviso_json, "aviso")

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
        ttk.Button(cab, text="Palabras clave", style="Header.TButton", command=self._abrir_editor_palabras).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(cab, text="Info de uso", style="Header.TButton", command=self._mostrar_info_api).grid(row=0, column=2)

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
        mb = ttk.Menubutton(parent, text="Exportar  ▾")
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
        for col in (1, 3, 5, 7, 9):
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

        ttk.Label(card, text="Mín. reseñas", style="Card.TLabel").grid(row=0, column=6, sticky="w", padx=(0, 6))
        self.var_min = tk.StringVar(value="5")
        self.combo_min = ttk.Combobox(card, textvariable=self.var_min, values=MIN_RESEÑAS_OPCIONES, state="readonly", width=6)
        self.combo_min.grid(row=0, column=7, sticky="w", padx=(0, 14))
        self.combo_min.bind("<<ComboboxSelected>>", self._on_cambio_reaplicar)

        ttk.Label(card, text="Por página", style="Card.TLabel").grid(row=0, column=8, sticky="w", padx=(0, 6))
        self.var_pp = tk.StringVar(value="25")
        self.combo_pp = ttk.Combobox(card, textvariable=self.var_pp, values=OPCIONES_POR_PAGINA, state="readonly", width=6)
        self.combo_pp.grid(row=0, column=9, sticky="w", padx=(0, 14))
        self.combo_pp.bind("<<ComboboxSelected>>", self._on_cambio_reaplicar)

        botones = ttk.Frame(card, style="Card.TFrame")
        botones.grid(row=0, column=10, sticky="e")
        self.btn_buscar = ttk.Button(botones, text="Buscar negocios", style="Primary.TButton", command=self._on_buscar)
        self.btn_buscar.pack(side="left")
        self.btn_cancelar = ttk.Button(botones, text="Cancelar", command=self._on_cancelar, state="disabled")
        self.btn_cancelar.pack(side="left", padx=(8, 0))

        self.lbl_terminos = ttk.Label(card, text="", style="CardSuave.TLabel")
        self.lbl_terminos.grid(row=1, column=0, columnspan=11, sticky="w", pady=(8, 0))

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
        self.mb_estado_g = ttk.Menubutton(tools, text="Cambiar estado  ▾")
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

        threading.Thread(target=run, daemon=True).start()

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
            self.after(100, self._procesar_cola)
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
        self._guardar_nota_pendiente()
        if self.nb.index(self.nb.select()) == 1:
            self._refrescar_guardados()
        else:
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
        except OSError as exc:
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
        v.geometry("560x620")
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
        ttk.Button(barra, text="Cerrar", command=v.destroy).pack(side="right")
        ttk.Button(marco, text="Volver a consultar términos omitidos (aportaban 0 nuevos)", command=reiniciar_omitidos).grid(
            row=5, column=0, sticky="w", pady=(10, 0))
        combo.bind("<<ComboboxSelected>>", cargar)
        cargar()
        self._centrar_sobre(v)

    # ------------------------------------------------------------------ cierre
    def _cerrar(self) -> None:
        try:
            self._guardar_nota_pendiente()
            self.tabla_b.cerrar_editor(True)
            self.tabla_g.cerrar_editor(True)
        finally:
            self._cancelar.set()
            self._db.cerrar()
            self.destroy()


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------

def main() -> None:
    app = ProspeccionApp()
    app.mainloop()


if __name__ == "__main__":
    main()
