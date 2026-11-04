#!/usr/bin/env python3
"""
Herramienta de prospección de clientes para vender chatbots (WhatsApp automation).
Interfaz gráfica con Tkinter — un solo archivo ejecutable.
"""

from __future__ import annotations

import json
import math
import os
import random
import threading
import time
import tkinter as tk
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from datetime import datetime
from tkinter import messagebox, ttk
from typing import Callable, Iterable, Optional

from openpyxl import Workbook
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
    "Puerto Rico": "PR",
}


TIPOS_NEGOCIO = [
    "Comidas rápidas",
    "Restaurantes",
    "Barberías",
    "Odontólogos",
    "Médicos",
    "Inmobiliarias",
]

COLUMNAS_TABLA = (
    "negocio",
    "ciudad",
    "whatsapp",
    "reseñas",
    "web_instagram",
    "link_google_maps",
    "direccion",
)

COLUMNAS_EXPORT = (
    "Negocio",
    "Ciudad",
    "WhatsApp",
    "Reseñas",
    "Web / Instagram",
    "Link Google Maps",
    "Dirección",
)

MIN_RESEÑAS_OPCIONES = ["0", "5", "10", "20", "50", "100"]

TIPOS_PRESENCIA = ["Todos", "Instagram", "Página web", "Facebook"]

OPCIONES_POR_PAGINA = ["10", "25", "50", "100", "200", "300", "400", "500", "600", "todos"]
MAX_CRUDO_POR_LOTE = 500

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INFORMES_DIR = os.path.join(_BASE_DIR, "informes")

# Variantes de consulta para obtener más negocios reales (Google ~60 por consulta)
_CONSULTAS_EXTRA: dict[str, list[str]] = {
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
    "Inmobiliarias": [
        "inmobiliaria", "bienes raíces", "apartamentos", "casas en venta", "arriendo apartamentos",
        "corretaje", "propiedades", "lotes en venta", "oficinas en arriendo",
    ],
}


# ---------------------------------------------------------------------------
# Modelo de datos
# ---------------------------------------------------------------------------

@dataclass
class Negocio:
    nombre: str
    ciudad: str
    tipo: str
    whatsapp: str
    web_instagram: str
    reseñas: int
    actividad_reciente: bool
    link_google_maps: str = ""
    direccion: str = ""
    tipo_presencia: str = ""
    place_id: str = ""
    problema: str = ""
    mensaje: str = ""


def _leer_google_api_key() -> str:
    key = os.environ.get("GOOGLE_MAPS_API_KEY", "").strip()
    if key:
        return key
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.isfile(env_path):
        with open(env_path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("GOOGLE_MAPS_API_KEY="):
                    return line.split("=", 1)[1].strip().strip('"').strip("'")
    return ""


def _widget_bg(widget: tk.Widget) -> str:
    try:
        return str(widget.cget("background"))
    except tk.TclError:
        return "#f0f0f0"


# ---------------------------------------------------------------------------
# Seguimiento de uso de Google Places API
# ---------------------------------------------------------------------------

class ApiUsageTracker:
    """Registra peticiones locales a Places API (Text Search)."""

    ARCHIVO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "api_usage.json")
    CREDITO_MENSUAL_USD = 200.0
    COSTO_ESTIMADO_POR_PETICION_USD = 0.035

    def __init__(self) -> None:
        self.peticiones_sesion = 0
        self.peticiones_ultima_busqueda = 0
        self._contando_busqueda = False
        self._datos = self._cargar()

    def _mes_actual(self) -> str:
        return time.strftime("%Y-%m")

    def _cargar(self) -> dict:
        if os.path.isfile(self.ARCHIVO):
            try:
                with open(self.ARCHIVO, encoding="utf-8") as f:
                    datos = json.load(f)
            except (json.JSONDecodeError, OSError):
                datos = {}
        else:
            datos = {}

        if datos.get("mes") != self._mes_actual():
            datos["mes"] = self._mes_actual()
            datos["peticiones_mes"] = 0

        datos.setdefault("total_peticiones", 0)
        datos.setdefault("peticiones_mes", 0)
        datos.setdefault("ultima_busqueda", 0)
        return datos

    def _guardar(self) -> None:
        try:
            with open(self.ARCHIVO, "w", encoding="utf-8") as f:
                json.dump(self._datos, f, indent=2)
        except OSError:
            pass

    def iniciar_busqueda(self) -> None:
        self._contando_busqueda = True
        self.peticiones_ultima_busqueda = 0

    def finalizar_busqueda(self) -> None:
        self._contando_busqueda = False
        self._datos["ultima_busqueda"] = self.peticiones_ultima_busqueda
        self._guardar()

    def registrar_peticion(self) -> None:
        self.peticiones_sesion += 1
        self._datos["total_peticiones"] = self._datos.get("total_peticiones", 0) + 1

        if self._datos.get("mes") != self._mes_actual():
            self._datos["mes"] = self._mes_actual()
            self._datos["peticiones_mes"] = 0

        self._datos["peticiones_mes"] = self._datos.get("peticiones_mes", 0) + 1

        if self._contando_busqueda:
            self.peticiones_ultima_busqueda += 1

        self._guardar()

    def resumen_texto(self) -> str:
        mes = self._datos.get("mes", self._mes_actual())
        peticiones_mes = self._datos.get("peticiones_mes", 0)
        total = self._datos.get("total_peticiones", 0)
        costo_mes = peticiones_mes * self.COSTO_ESTIMADO_POR_PETICION_USD
        credito_restante = max(0.0, self.CREDITO_MENSUAL_USD - costo_mes)
        uso_pct = min(100.0, (costo_mes / self.CREDITO_MENSUAL_USD) * 100) if self.CREDITO_MENSUAL_USD else 0

        return (
            "Uso de Google Places API (Text Search)\n"
            "────────────────────────────────────\n\n"
            f"Sesión actual:              {self.peticiones_sesion} peticiones\n"
            f"Última búsqueda:            {self._datos.get('ultima_busqueda', 0)} peticiones\n"
            f"Este mes ({mes}):           {peticiones_mes} peticiones\n"
            f"Total histórico (app):      {total} peticiones\n\n"
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
# Google Maps Places API (datos verificados)
# ---------------------------------------------------------------------------

class BusinessDataProvider:
    """Obtiene negocios reales desde Google Maps Places API."""

    _TIPO_BUSQUEDA: dict[str, str] = {
        "Comidas rápidas": "comida rápida",
        "Restaurantes": "restaurante",
        "Barberías": "barbería",
        "Odontólogos": "odontólogo clínica dental",
        "Médicos": "médico consultorio",
        "Inmobiliarias": "inmobiliaria",
    }

    _FIELD_MASK = (
        "places.id,places.displayName,places.formattedAddress,"
        "places.internationalPhoneNumber,places.nationalPhoneNumber,"
        "places.websiteUri,places.userRatingCount,places.businessStatus,"
        "places.googleMapsUri"
    )

    _PROBLEMAS: dict[str, list[str]] = {
        "Comidas rápidas": [
            "Responde manualmente WhatsApp en horas pico",
            "No tiene automatización de pedidos",
            "Pierde pedidos cuando no está disponible",
            "No confirma pedidos automáticamente",
        ],
        "Restaurantes": [
            "No tiene sistema de reservas",
            "Responde manualmente WhatsApp",
            "No envía menú automático a nuevos clientes",
            "Sin recordatorios de reserva por WhatsApp",
        ],
        "Barberías": [
            "Agenda citas manualmente por WhatsApp",
            "No tiene recordatorios automáticos de cita",
            "Responde tarde cuando está atendiendo",
            "No tiene sistema de reservas online",
        ],
        "Odontólogos": [
            "Agenda citas solo por llamada o WhatsApp manual",
            "No envía recordatorios automáticos de cita",
            "Responde consultas fuera de horario manualmente",
            "Sin confirmación automática de citas",
        ],
        "Médicos": [
            "Agenda citas solo por llamada o WhatsApp manual",
            "No envía recordatorios automáticos de cita a pacientes",
            "Recepción satura con consultas repetidas por WhatsApp",
            "Pierde pacientes que escriben fuera de horario de consulta",
        ],
        "Inmobiliarias": [
            "Responde consultas de propiedades manualmente",
            "No califica leads automáticamente",
            "Pierde interesados fuera de horario laboral",
            "Sin seguimiento automatizado por WhatsApp",
        ],
    }

    _MENSAJES: dict[str, str] = {
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
        "Inmobiliarias": (
            "Hola, vi {nombre} y creo que un asistente por WhatsApp que califique leads "
            "y responda consultas 24/7 te ayudaría a cerrar más negocios sin perder "
            "interesados fuera de horario."
        ),
    }

    def __init__(self) -> None:
        self._api_key = _leer_google_api_key()
        self.usage = ApiUsageTracker()

    def buscar(
        self,
        ciudad: str,
        tipo: str,
        pais: str = "Colombia",
        cantidad: int = MAX_CRUDO_POR_LOTE,
        tipo_presencia: str = "Todos",
        excluir_ids: Optional[set[str]] = None,
        desde_consulta: int = 0,
    ) -> tuple[list[Negocio], int]:
        if not self._api_key:
            raise ValueError(
                "Falta la API key de Google Maps.\n"
                "Crea un archivo .env con GOOGLE_MAPS_API_KEY=tu_clave\n"
                "o define la variable de entorno GOOGLE_MAPS_API_KEY.\n"
                "Activa Places API (New) en Google Cloud Console."
            )
        self.usage.iniciar_busqueda()
        try:
            return self._buscar_google_maps(
                ciudad, tipo, pais, cantidad, tipo_presencia,
                excluir_ids or set(), desde_consulta,
            )
        finally:
            self.usage.finalizar_busqueda()

    def _request_places(self, body: dict) -> dict:
        url = "https://places.googleapis.com/v1/places:searchText"
        req = urllib.request.Request(
            url,
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
                self.usage.registrar_peticion()
                return data
        except urllib.error.HTTPError as exc:
            detalle = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"Google Places API error ({exc.code}): {detalle}") from exc

    @staticmethod
    def _clasificar_presencia(website: str) -> tuple[str, str]:
        url = website.strip()
        if not url:
            return "", ""
        lower = url.lower()
        if "instagram.com" in lower:
            return url, "instagram"
        if "facebook.com" in lower or "fb.com" in lower or "fb.me" in lower:
            return url, "facebook"
        return url, "web"

    def _place_a_negocio(self, place: dict, ciudad: str, tipo: str) -> Optional[Negocio]:
        place_id = place.get("id", "").strip()
        nombre = (place.get("displayName") or {}).get("text", "").strip()
        if not place_id or not nombre:
            return None

        telefono = (
            place.get("internationalPhoneNumber")
            or place.get("nationalPhoneNumber")
            or ""
        ).strip()

        website = (place.get("websiteUri") or "").strip()
        web_url, presencia = self._clasificar_presencia(website)

        reseñas = int(place["userRatingCount"]) if "userRatingCount" in place else 0
        status = place.get("businessStatus", "")
        actividad = status != "CLOSED_PERMANENTLY"
        link_maps = (place.get("googleMapsUri") or "").strip()
        if not link_maps and place_id:
            link_maps = f"https://www.google.com/maps/place/?q=place_id:{place_id}"
        direccion = (place.get("formattedAddress") or "").strip()

        return Negocio(
            nombre=nombre,
            ciudad=ciudad,
            tipo=tipo,
            whatsapp=telefono,
            web_instagram=web_url,
            reseñas=reseñas,
            actividad_reciente=actividad,
            link_google_maps=link_maps,
            direccion=direccion,
            tipo_presencia=presencia,
            place_id=place_id,
        )

    def _consultas_busqueda(self, ciudad: str, pais: str, tipo: str) -> list[str]:
        terminos = _CONSULTAS_EXTRA.get(tipo, [self._TIPO_BUSQUEDA.get(tipo, tipo)])
        consultas: list[str] = []
        vistos: set[str] = set()
        for termino in terminos:
            for plantilla in (
                f"{termino} en {ciudad}, {pais}",
                f"{termino} {ciudad} {pais}",
                f"{termino} {ciudad}",
            ):
                clave = plantilla.lower()
                if clave not in vistos:
                    vistos.add(clave)
                    consultas.append(plantilla)
        return consultas

    def _paginar_consulta(
        self,
        consulta: str,
        ciudad: str,
        tipo: str,
        region: str,
        ids_vistos: set[str],
        resultados: list[Negocio],
        meta: int,
    ) -> None:
        page_token: Optional[str] = None
        paginas = 0
        max_paginas = 3  # tope de la API (~60 por consulta)

        while paginas < max_paginas and len(resultados) < meta:
            faltan = meta - len(resultados)
            body: dict = {
                "textQuery": consulta,
                "languageCode": "es",
                "regionCode": region,
                "maxResultCount": min(20, max(faltan, 1)),
            }
            if page_token:
                body["pageToken"] = page_token

            data = self._request_places(body)
            places = data.get("places", [])
            if not places:
                break

            for place in places:
                pid = place.get("id", "")
                if not pid or pid in ids_vistos:
                    continue
                ids_vistos.add(pid)
                negocio = self._place_a_negocio(place, ciudad, tipo)
                if negocio is None:
                    continue
                resultados.append(negocio)
                if len(resultados) >= meta:
                    return

            page_token = data.get("nextPageToken")
            if not page_token:
                break
            paginas += 1
            time.sleep(0.35)

    def _buscar_google_maps(
        self,
        ciudad: str,
        tipo: str,
        pais: str,
        cantidad: int,
        tipo_presencia: str,
        excluir_ids: set[str],
        desde_consulta: int,
    ) -> tuple[list[Negocio], int]:
        del tipo_presencia
        meta = max(1, cantidad)
        resultados: list[Negocio] = []
        ids_vistos: set[str] = set(excluir_ids)
        region = _PAIS_REGION.get(pais, "CO")
        todas = self._consultas_busqueda(ciudad, pais, tipo)
        consultas = todas[desde_consulta:]
        max_consultas = min(len(consultas), max(8, (meta + 15) // 15))
        consultas = consultas[:max_consultas]
        ejecutadas = 0

        for consulta in consultas:
            if len(resultados) >= meta:
                break
            ejecutadas += 1
            self._paginar_consulta(
                consulta, ciudad, tipo, region, ids_vistos, resultados, meta
            )
            time.sleep(0.4)

        return resultados, desde_consulta + ejecutadas


# ---------------------------------------------------------------------------
# Análisis y filtrado
# ---------------------------------------------------------------------------

class LeadAnalyzer:
    """Genera problemas detectados y mensajes de venta personalizados."""

    def __init__(self, provider: BusinessDataProvider) -> None:
        self._provider = provider

    def analizar(self, negocio: Negocio) -> Negocio:
        problemas = BusinessDataProvider._PROBLEMAS.get(negocio.tipo, ["Responde manualmente WhatsApp"])
        plantilla = BusinessDataProvider._MENSAJES.get(
            negocio.tipo,
            "Hola, vi {nombre} y creo que podrías automatizar tu atención por WhatsApp.",
        )
        negocio.problema = random.choice(problemas)
        negocio.mensaje = plantilla.format(nombre=negocio.nombre)
        return negocio


class LeadFilter:
    """Filtros obligatorios de calidad de prospecto."""

    @staticmethod
    def cumple(negocio: Negocio, min_reseñas: int, tipo_presencia: str) -> bool:
        del tipo_presencia  # presencia online no excluye; celda vacía si Google no la reporta
        if not negocio.actividad_reciente:
            return False
        return negocio.reseñas >= min_reseñas

    @staticmethod
    def filtrar(
        negocios: Iterable[Negocio],
        min_reseñas: int,
        tipo_presencia: str,
    ) -> list[Negocio]:
        return [
            n for n in negocios
            if LeadFilter.cumple(n, min_reseñas, tipo_presencia)
        ]

    @staticmethod
    def deduplicar(negocios: Iterable[Negocio]) -> list[Negocio]:
        place_ids: set[str] = set()
        unicos: list[Negocio] = []
        for n in negocios:
            if n.place_id:
                if n.place_id in place_ids:
                    continue
                place_ids.add(n.place_id)
            unicos.append(n)
        return unicos


# ---------------------------------------------------------------------------
# Exportación Excel
# ---------------------------------------------------------------------------

def _sanitizar_nombre(parte: str) -> str:
    invalidos = '<>:"/\\|?*'
    s = parte.strip().replace(" ", "-")
    for c in invalidos:
        s = s.replace(c, "")
    return s


def exportar_excel(ruta: str, negocios: list[Negocio]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"

    filas = [
        (n.nombre, n.ciudad, n.whatsapp, n.reseñas, n.web_instagram, n.link_google_maps, n.direccion)
        for n in negocios
    ]
    num_cols = len(COLUMNAS_EXPORT)

    borde_fino = Side(style="thin", color="D0D0D0")
    borde = Border(left=borde_fino, right=borde_fino, top=borde_fino, bottom=borde_fino)
    relleno_cabecera = PatternFill(fill_type="solid", fgColor="0078D4")
    fuente_cabecera = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    relleno_fila_alt = PatternFill(fill_type="solid", fgColor="EEF4FA")
    alin_cabecera = Alignment(horizontal="center", vertical="center", wrap_text=True)
    alin_datos = Alignment(vertical="center", wrap_text=True)

    for col, titulo in enumerate(COLUMNAS_EXPORT, start=1):
        celda = ws.cell(row=1, column=col, value=titulo)
        celda.fill = relleno_cabecera
        celda.font = fuente_cabecera
        celda.alignment = alin_cabecera
        celda.border = borde

    for fila_idx, valores in enumerate(filas, start=2):
        zebra = fila_idx % 2 == 0
        for col, valor in enumerate(valores, start=1):
            celda = ws.cell(row=fila_idx, column=col, value=valor)
            celda.border = borde
            celda.alignment = alin_datos
            if zebra:
                celda.fill = relleno_fila_alt

    for col in range(1, num_cols + 1):
        letra = get_column_letter(col)
        ancho = len(str(COLUMNAS_EXPORT[col - 1]))
        for fila_idx in range(2, ws.max_row + 1):
            valor = ws.cell(row=fila_idx, column=col).value
            if valor is not None:
                ancho = max(ancho, len(str(valor)))
        ws.column_dimensions[letra].width = min(ancho + 3, 55)

    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"
    if filas:
        ultima_col = get_column_letter(num_cols)
        ws.auto_filter.ref = f"A1:{ultima_col}{ws.max_row}"

    wb.save(ruta)


# ---------------------------------------------------------------------------
# Interfaz gráfica
# ---------------------------------------------------------------------------

class SpinnerIndicator(tk.Canvas):
    """Indicador circular con puntos (estilo Material)."""

    PUNTOS = 8

    def __init__(self, parent: tk.Widget, size: int = 34, **kwargs) -> None:
        bg = _widget_bg(parent)
        super().__init__(
            parent,
            width=size,
            height=size,
            highlightthickness=0,
            bg=bg,
            **kwargs,
        )
        self._size = size
        self._cx = size / 2
        self._cy = size / 2
        self._radio = size / 2 - 7
        self._radio_punto = 3
        self._paso = 0
        self._running = False
        self._colores = ("#0078d4", "#2d93dc", "#5aade4", "#87c7ec", "#b3dff5", "#d9ecfa")
        self._puntos: list[int] = []

        for i in range(self.PUNTOS):
            angulo = math.radians(i * 360 / self.PUNTOS - 90)
            x = self._cx + self._radio * math.cos(angulo)
            y = self._cy + self._radio * math.sin(angulo)
            r = self._radio_punto
            punto = self.create_oval(
                x - r, y - r, x + r, y + r,
                fill=self._colores[-1],
                outline="",
            )
            self._puntos.append(punto)

    def start(self) -> None:
        if self._running:
            return
        self._running = True
        self.pack(side="left", padx=(10, 0))
        self._tick()

    def stop(self) -> None:
        self._running = False
        self.pack_forget()

    def _tick(self) -> None:
        if not self._running:
            return
        for i, punto_id in enumerate(self._puntos):
            distancia = (self._paso - i) % self.PUNTOS
            indice = max(0, min(distancia, len(self._colores) - 1))
            self.itemconfigure(punto_id, fill=self._colores[indice])
        self._paso = (self._paso + 1) % self.PUNTOS
        self.after(85, self._tick)


class ProspeccionApp(tk.Tk):
    """Aplicación principal de prospección de clientes."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Prospección de Clientes — Chatbot CRM")
        self.geometry("1200x720")
        self.minsize(900, 550)

        self._provider = BusinessDataProvider()
        self._analyzer = LeadAnalyzer(self._provider)
        self._leads: list[Negocio] = []
        self._leads_visibles: list[Negocio] = []
        self._iid_a_negocio: dict[str, Negocio] = {}
        self._ultima_celda: Optional[tuple[str, int]] = None
        self._celda_widget: Optional[tk.Widget] = None
        self._modo_celda: Optional[str] = None  # "seleccion" | "edicion"
        self._edit_guardando = False
        self._ultimo_click: Optional[tuple[float, str, int]] = None
        self._buscando = False
        self._exportando = False
        self._pagina_actual = 0
        self._paginas: dict[int, list[Negocio]] = {}
        self._ids_global: set[str] = set()
        self._hay_mas_paginas = False
        self._params_busqueda: Optional[dict] = None
        self._consulta_offsets: dict[str, int] = {}

        self._build_ui()
        self._centrar_ventana()

    def _centrar_ventana(self) -> None:
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"+{x}+{y}")

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # --- Panel superior: selectores ---
        frame_top = ttk.LabelFrame(self, text="Búsqueda", padding=10)
        frame_top.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        for col in (1, 3, 5, 7, 9):
            frame_top.columnconfigure(col, weight=1)

        ttk.Label(frame_top, text="País:").grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.var_pais = tk.StringVar(value="Colombia")
        self.combo_pais = ttk.Combobox(
            frame_top, textvariable=self.var_pais, values=PAISES, state="readonly", width=16
        )
        self.combo_pais.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        self.combo_pais.bind("<<ComboboxSelected>>", self._on_pais_change)

        ttk.Label(frame_top, text="Ciudad:").grid(row=0, column=2, sticky="w", padx=(0, 5))
        self.var_ciudad = tk.StringVar(value=CIUDADES_POR_PAIS["Colombia"][0])
        self.combo_ciudad = ttk.Combobox(
            frame_top,
            textvariable=self.var_ciudad,
            values=CIUDADES_POR_PAIS["Colombia"],
            state="readonly",
            width=16,
        )
        self.combo_ciudad.grid(row=0, column=3, sticky="ew", padx=(0, 10))
        self.combo_ciudad.bind("<<ComboboxSelected>>", self._on_ciudad_change)

        ttk.Label(frame_top, text="Tipo de negocio:").grid(row=0, column=4, sticky="w", padx=(0, 5))
        self.var_tipo = tk.StringVar(value=TIPOS_NEGOCIO[0])
        self.combo_tipo = ttk.Combobox(
            frame_top, textvariable=self.var_tipo, values=TIPOS_NEGOCIO, state="readonly", width=16
        )
        self.combo_tipo.grid(row=0, column=5, sticky="ew", padx=(0, 10))

        ttk.Label(frame_top, text="Por página:").grid(row=0, column=6, sticky="w", padx=(0, 5))
        self.var_por_pagina = tk.StringVar(value=OPCIONES_POR_PAGINA[0])
        self.combo_por_pagina = ttk.Combobox(
            frame_top,
            textvariable=self.var_por_pagina,
            values=OPCIONES_POR_PAGINA,
            state="readonly",
            width=6,
        )
        self.combo_por_pagina.grid(row=0, column=7, sticky="w", padx=(0, 10))

        frame_btn = tk.Frame(frame_top)
        frame_btn.grid(row=0, column=8, padx=5, sticky="w")
        self.btn_buscar = ttk.Button(frame_btn, text="Buscar negocios", command=self._on_buscar)
        self.btn_buscar.pack(side="left")
        self._spinner = SpinnerIndicator(frame_btn)

        self.lbl_estado = ttk.Label(frame_top, text="Listo para buscar", foreground="#555")
        self.lbl_estado.grid(row=0, column=9, sticky="e", padx=(10, 0))

        ttk.Label(frame_top, text="Mín. reseñas:").grid(
            row=1, column=0, sticky="w", padx=(0, 5), pady=(8, 0)
        )
        self.var_min_resenas = tk.StringVar(value="5")
        self.combo_min_resenas = ttk.Combobox(
            frame_top,
            textvariable=self.var_min_resenas,
            values=MIN_RESEÑAS_OPCIONES,
            state="readonly",
            width=16,
        )
        self.combo_min_resenas.grid(row=1, column=1, sticky="ew", padx=(0, 10), pady=(8, 0))

        ttk.Label(frame_top, text="Presencia online:").grid(
            row=1, column=2, sticky="w", padx=(0, 5), pady=(8, 0)
        )
        self.var_presencia = tk.StringVar(value=TIPOS_PRESENCIA[0])
        self.combo_presencia = ttk.Combobox(
            frame_top,
            textvariable=self.var_presencia,
            values=TIPOS_PRESENCIA,
            state="readonly",
            width=16,
        )
        self.combo_presencia.grid(row=1, column=3, sticky="ew", padx=(0, 10), pady=(8, 0))

        ttk.Label(
            frame_top,
            text="Instagram · Página web · Facebook · Todos",
            foreground="#666",
            font=("Segoe UI", 8),
        ).grid(row=1, column=4, columnspan=4, sticky="w", padx=(0, 5), pady=(8, 0))

        # --- Panel filtros / acciones ---
        frame_actions = ttk.Frame(self, padding=(10, 0))
        frame_actions.grid(row=1, column=0, sticky="ew")
        frame_actions.columnconfigure(1, weight=1)

        ttk.Label(frame_actions, text="Filtrar en tabla:").grid(row=0, column=0, padx=(0, 5))
        self.var_filtro_tabla = tk.StringVar()
        self.var_filtro_tabla.trace_add("write", lambda *_: self._aplicar_filtro_tabla())
        self.entry_filtro = ttk.Entry(frame_actions, textvariable=self.var_filtro_tabla, width=30)
        self.entry_filtro.grid(row=0, column=1, sticky="ew", padx=(0, 15))

        self.lbl_leads = ttk.Label(
            frame_actions,
            text="0 leads listos para contactar",
            font=("Segoe UI", 10, "bold"),
            foreground="#1a7f37",
        )
        self.lbl_leads.grid(row=0, column=2, padx=10)

        self.btn_copiar = ttk.Button(
            frame_actions, text="Copiar mensaje", command=self._on_copiar_mensaje, state="disabled"
        )
        self.btn_copiar.grid(row=0, column=3, padx=5)

        self.btn_exportar = ttk.Button(
            frame_actions, text="Exportar a Excel", command=self._on_exportar, state="disabled"
        )
        self.btn_exportar.grid(row=0, column=4, padx=5)

        self.btn_info = ttk.Button(frame_actions, text="Info", command=self._mostrar_info_api)
        self.btn_info.grid(row=0, column=5, padx=5)

        frame_pag = ttk.Frame(frame_actions)
        frame_pag.grid(row=1, column=0, columnspan=6, sticky="w", pady=(8, 0))

        self.btn_pag_anterior = ttk.Button(
            frame_pag, text="◀ Anterior", command=self._pagina_anterior, state="disabled"
        )
        self.btn_pag_anterior.pack(side="left")

        self.lbl_pagina = ttk.Label(
            frame_pag,
            text="Página —",
            font=("Segoe UI", 9, "bold"),
            padding=(12, 0),
        )
        self.lbl_pagina.pack(side="left")

        self.btn_pag_siguiente = ttk.Button(
            frame_pag, text="Siguiente ▶", command=self._pagina_siguiente, state="disabled"
        )
        self.btn_pag_siguiente.pack(side="left")

        ttk.Label(
            frame_pag,
            text="  ·  sin repetir entre páginas",
            foreground="#666",
            font=("Segoe UI", 8),
        ).pack(side="left", padx=(8, 0))

        # --- Tabla ---
        frame_tabla = ttk.Frame(self, padding=10)
        frame_tabla.grid(row=2, column=0, sticky="nsew")
        frame_tabla.rowconfigure(0, weight=1)
        frame_tabla.columnconfigure(0, weight=1)

        headers = {
            "negocio": "Negocio",
            "ciudad": "Ciudad",
            "whatsapp": "WhatsApp",
            "web_instagram": "Web / Instagram",
            "link_google_maps": "Link Google Maps",
            "direccion": "Dirección",
            "reseñas": "Reseñas",
        }
        self._headers_tabla = headers
        widths = {
            "negocio": 180,
            "ciudad": 110,
            "whatsapp": 140,
            "reseñas": 80,
            "web_instagram": 150,
            "link_google_maps": 150,
            "direccion": 220,
        }

        self.tree = ttk.Treeview(
            frame_tabla,
            columns=COLUMNAS_TABLA,
            show="headings",
            selectmode="browse",
        )
        for col in COLUMNAS_TABLA:
            self.tree.heading(
                col, text=headers[col], anchor="center",
                command=lambda c=col: self._ordenar_por(c),
            )
            self.tree.column(
                col,
                width=widths[col],
                minwidth=60,
                anchor="center",
                stretch=False,
            )

        vsb = ttk.Scrollbar(frame_tabla, orient="vertical", command=self._on_scroll_vertical)
        hsb = ttk.Scrollbar(frame_tabla, orient="horizontal", command=self._on_scroll_horizontal)
        self._vsb = vsb
        self._hsb = hsb
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.tree.bind("<<TreeviewSelect>>", self._on_seleccion)
        self.tree.bind("<Button-1>", self._on_click_celda, add="+")
        self.tree.bind("<Control-c>", self._on_ctrl_c)
        self.tree.bind("<Delete>", self._on_borrar_celda)
        self.tree.bind("<Double-1>", self._on_doble_clic, add="+")
        self.tree.bind("<Configure>", lambda _e: self._ocultar_celda_ui())
        self.tree.bind("<MouseWheel>", lambda _e: self._ocultar_celda_ui())

        # --- Barra inferior ---
        frame_bottom = ttk.Frame(self, padding=(10, 0, 10, 10))
        frame_bottom.grid(row=3, column=0, sticky="ew")

        self.lbl_filtros = ttk.Label(
            frame_bottom,
            text="",
            foreground="#666",
            font=("Segoe UI", 9),
        )
        self.lbl_filtros.pack(anchor="w")
        self._actualizar_texto_filtros()

        self._sort_reverse: dict[str, bool] = {c: False for c in COLUMNAS_TABLA}

    def _mostrar_info_api(self) -> None:
        ventana = tk.Toplevel(self)
        ventana.title("Info — Cuota Google Places API")
        ventana.geometry("460x430")
        ventana.resizable(False, False)
        ventana.transient(self)
        ventana.grab_set()

        marco = ttk.Frame(ventana, padding=16)
        marco.pack(fill="both", expand=True)

        texto = tk.Text(
            marco,
            wrap="word",
            font=("Consolas", 10),
            relief="flat",
            bg="#fafafa",
            padx=10,
            pady=10,
        )
        texto.pack(fill="both", expand=True)

        def actualizar_contenido() -> None:
            texto.config(state="normal")
            texto.delete("1.0", tk.END)
            texto.insert("1.0", self._provider.usage.resumen_texto())
            texto.config(state="disabled")

        actualizar_contenido()

        barra = ttk.Frame(marco)
        barra.pack(fill="x", pady=(12, 0))
        ttk.Button(barra, text="Actualizar", command=actualizar_contenido).pack(side="left")
        ttk.Button(barra, text="Cerrar", command=ventana.destroy).pack(side="right")

        ventana.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - ventana.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - ventana.winfo_height()) // 2
        ventana.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _actualizar_texto_filtros(self) -> None:
        min_res = self.var_min_resenas.get() if hasattr(self, "var_min_resenas") else "5"
        self.lbl_filtros.config(
            text=(
                f"Fuente: Google Maps · {self.var_por_pagina.get()} por página · "
                f"mín. {min_res} reseñas · campos vacíos si Google no los reporta · "
                "sin duplicados  |  1 clic: seleccionar  |  2 clics: editar"
            )
        )

    # --- Búsqueda y paginación ---

    def _on_pais_change(self, *_: object) -> None:
        ciudades = CIUDADES_POR_PAIS.get(self.var_pais.get(), [])
        if ciudades:
            self.combo_ciudad["values"] = ciudades
            self.var_ciudad.set(ciudades[0])

    def _on_ciudad_change(self, *_: object) -> None:
        ciudad = self.var_ciudad.get()
        for pais, ciudades in CIUDADES_POR_PAIS.items():
            if ciudad in ciudades:
                self.var_pais.set(pais)
                self.combo_ciudad["values"] = ciudades
                break

    def _limite_pagina(self) -> int:
        if self.var_por_pagina.get().strip().lower() == "todos":
            return 10_000
        try:
            return max(1, int(self.var_por_pagina.get()))
        except ValueError:
            return 10

    def _cantidad_cruda(self, faltan_filtrados: int, limite_pagina: int) -> int:
        objetivo = max(faltan_filtrados, limite_pagina)
        return min(max(int(objetivo * 2), objetivo + 30), MAX_CRUDO_POR_LOTE)

    @staticmethod
    def _clave_consulta(pais: str, ciudad: str, tipo: str) -> str:
        return f"{pais}|{ciudad}|{tipo}"

    @staticmethod
    def _min_reseñas_relajado(min_reseñas: int) -> int:
        if min_reseñas <= 0:
            return 0
        if min_reseñas >= 20:
            return max(0, min_reseñas - 20)
        if min_reseñas >= 5:
            return max(0, min_reseñas - 5)
        return 0

    def _params_actuales(self) -> dict:
        return {
            "pais": self.var_pais.get(),
            "ciudad": self.var_ciudad.get(),
            "tipo": self.var_tipo.get(),
            "min_reseñas": int(self.var_min_resenas.get()),
            "tipo_presencia": self.var_presencia.get(),
            "limite_pagina": self._limite_pagina(),
            "por_pagina": self.var_por_pagina.get(),
        }

    def _iniciar_carga(self, pagina: int, reiniciar: bool) -> None:
        if self._buscando:
            return
        if reiniciar:
            self._consulta_offsets.clear()

        self._buscando = True
        self.btn_buscar.config(state="disabled")
        self.btn_pag_anterior.config(state="disabled")
        self.btn_pag_siguiente.config(state="disabled")
        self._spinner.start()
        self.lbl_estado.config(
            text=f"Consultando Google Maps (página {pagina})...",
            foreground="#b8860b",
        )

        threading.Thread(
            target=self._cargar_pagina_en_hilo,
            args=(pagina, reiniciar, self._params_actuales()),
            daemon=True,
        ).start()

    def _obtener_leads_pagina(self, params: dict, excluir_ids: set[str]) -> tuple[list[Negocio], bool]:
        limite = params["limite_pagina"]
        pagina_leads: list[Negocio] = []
        ids_pagina: set[str] = set()
        excluir = set(excluir_ids)
        pool: list[Negocio] = []

        def _incorporar(min_reseñas: int) -> None:
            analizados = [self._analyzer.analizar(n) for n in pool]
            filtrados = LeadFilter.filtrar(
                analizados, min_reseñas, params["tipo_presencia"]
            )
            for negocio in LeadFilter.deduplicar(filtrados):
                if not negocio.place_id or negocio.place_id in excluir_ids:
                    continue
                if negocio.place_id in ids_pagina:
                    continue
                pagina_leads.append(negocio)
                ids_pagina.add(negocio.place_id)
                if len(pagina_leads) >= limite:
                    break

        def _api(ciudad: str, cantidad: int) -> None:
            clave = self._clave_consulta(params["pais"], ciudad, params["tipo"])
            desde = self._consulta_offsets.get(clave, 0)
            crudos, siguiente = self._provider.buscar(
                ciudad,
                params["tipo"],
                pais=params["pais"],
                cantidad=cantidad,
                tipo_presencia=params["tipo_presencia"],
                excluir_ids=excluir,
                desde_consulta=desde,
            )
            self._consulta_offsets[clave] = siguiente
            for negocio in crudos:
                if negocio.place_id:
                    excluir.add(negocio.place_id)
            pool.extend(crudos)

        def _rellenar(ciudad: str, cantidad: int) -> None:
            _api(ciudad, cantidad)
            _incorporar(params["min_reseñas"])
            if len(pagina_leads) < limite:
                _incorporar(self._min_reseñas_relajado(params["min_reseñas"]))
            if len(pagina_leads) < limite and params["min_reseñas"] > 0:
                _incorporar(0)

        agotado = False
        for _ in range(12):
            if len(pagina_leads) >= limite:
                break
            antes_pool, antes_leads = len(pool), len(pagina_leads)
            _rellenar(
                params["ciudad"],
                self._cantidad_cruda(limite - len(pagina_leads), limite),
            )
            if len(pool) == antes_pool and len(pagina_leads) == antes_leads:
                agotado = True
                break

        hay_mas = len(pagina_leads) >= limite or not agotado
        return pagina_leads[:limite], hay_mas

    def _cargar_pagina_en_hilo(self, pagina: int, reiniciar: bool, params: dict) -> None:
        try:
            if reiniciar:
                excluir: set[str] = set()
            else:
                excluir = set(self._ids_global)

            leads, hay_mas = self._obtener_leads_pagina(params, excluir)

            self.after(0, lambda: self._finalizar_carga_pagina({
                "ok": True,
                "pagina": pagina,
                "reiniciar": reiniciar,
                "params": params,
                "leads": leads,
                "hay_mas": hay_mas,
                "estado": (
                    f"Página {pagina}: {len(leads)} leads · "
                    f"{params['ciudad']} — {params['tipo']}"
                ),
            }))
        except Exception as exc:
            self.after(0, lambda: self._finalizar_carga_pagina({
                "ok": False,
                "error": str(exc),
            }))

    def _on_buscar(self) -> None:
        self._iniciar_carga(pagina=1, reiniciar=True)

    def _pagina_anterior(self) -> None:
        if self._pagina_actual <= 1:
            return
        self._mostrar_pagina(self._pagina_actual - 1)

    def _pagina_siguiente(self) -> None:
        if self._params_busqueda is None:
            return
        siguiente = self._pagina_actual + 1
        if siguiente in self._paginas:
            self._mostrar_pagina(siguiente)
            return
        if not self._hay_mas_paginas:
            return
        self._iniciar_carga(pagina=siguiente, reiniciar=False)

    def _mostrar_pagina(self, pagina: int) -> None:
        if pagina not in self._paginas:
            return
        self._pagina_actual = pagina
        self._leads = self._paginas[pagina]
        self._leads_visibles = list(self._leads)
        self.var_filtro_tabla.set("")
        self._render_tabla(self._leads_visibles)
        self._actualizar_contador()
        self._actualizar_controles_paginacion()
        total_cargados = sum(len(p) for p in self._paginas.values())
        estado = f"Página {pagina}: {len(self._leads)} leads"
        if total_cargados > len(self._leads):
            estado += f" · {total_cargados} en total"
        self.lbl_estado.config(
            text=estado,
            foreground="#1a7f37" if self._leads else "#c0392b",
        )
        self.btn_exportar.config(state="normal" if self._leads else "disabled")

    def _actualizar_controles_paginacion(self) -> None:
        if self._pagina_actual <= 0:
            self.lbl_pagina.config(text="Página —")
            self.btn_pag_anterior.config(state="disabled")
            self.btn_pag_siguiente.config(state="disabled")
            return

        self.lbl_pagina.config(text=f"Página {self._pagina_actual}")
        self.btn_pag_anterior.config(
            state="normal" if self._pagina_actual > 1 else "disabled"
        )
        puede_siguiente = (
            (self._pagina_actual + 1) in self._paginas or self._hay_mas_paginas
        )
        self.btn_pag_siguiente.config(
            state="normal" if puede_siguiente and not self._buscando else "disabled"
        )

    def _finalizar_carga_pagina(self, resultado: dict) -> None:
        self._spinner.stop()
        self.btn_buscar.config(state="normal")
        self._buscando = False

        if not resultado.get("ok"):
            messagebox.showerror(
                "Error",
                f"No se pudo completar la búsqueda:\n{resultado.get('error', 'Error desconocido')}",
            )
            self.lbl_estado.config(text="Error en búsqueda", foreground="#c0392b")
            self._actualizar_controles_paginacion()
            return

        pagina = resultado["pagina"]
        if resultado.get("reiniciar"):
            self._paginas.clear()
            self._ids_global.clear()
            self._params_busqueda = resultado["params"]

        if not resultado["leads"] and pagina > 1:
            self._hay_mas_paginas = False
            messagebox.showinfo(
                "Sin más resultados",
                "Google Maps no devolvió más negocios nuevos para esta búsqueda.",
            )
            self._actualizar_controles_paginacion()
            return

        self._paginas[pagina] = resultado["leads"]
        for negocio in resultado["leads"]:
            if negocio.place_id:
                self._ids_global.add(negocio.place_id)

        self._hay_mas_paginas = resultado["hay_mas"]
        self._mostrar_pagina(pagina)
        self.lbl_estado.config(
            text=resultado["estado"],
            foreground="#1a7f37" if resultado["leads"] else "#c0392b",
        )
        self._actualizar_texto_filtros()

    def _render_tabla(self, negocios: list[Negocio]) -> None:
        self._ocultar_celda_ui()
        self.tree.delete(*self.tree.get_children())
        self._iid_a_negocio.clear()
        self._ultima_celda = None
        for i, n in enumerate(negocios):
            iid = f"lead_{i}"
            self._iid_a_negocio[iid] = n
            self.tree.insert("", "end", iid=iid, values=(
                n.nombre,
                n.ciudad,
                n.whatsapp,
                n.reseñas,
                n.web_instagram,
                n.link_google_maps,
                n.direccion,
            ))

    def _actualizar_contador(self) -> None:
        visibles = len(self._leads_visibles)
        if self._pagina_actual > 0:
            texto = f"Pág. {self._pagina_actual}: {visibles} leads"
            total_cargados = sum(len(p) for p in self._paginas.values())
            if total_cargados > visibles:
                texto += f" · {total_cargados} en total"
        else:
            texto = f"{visibles} leads listos para contactar"
        self.lbl_leads.config(text=texto)

    def _aplicar_filtro_tabla(self) -> None:
        if not self._leads:
            return
        termino = self.var_filtro_tabla.get().strip().lower()
        if not termino:
            self._leads_visibles = list(self._leads)
        else:
            self._leads_visibles = [
                n for n in self._leads
                if termino in " ".join([
                    n.nombre, n.ciudad, n.whatsapp, n.web_instagram,
                    n.link_google_maps, n.direccion, str(n.reseñas),
                ]).lower()
            ]
        self._render_tabla(self._leads_visibles)
        self._actualizar_contador()

    def _ordenar_por(self, columna: str) -> None:
        idx = COLUMNAS_TABLA.index(columna)
        reverse = self._sort_reverse[columna]
        self._sort_reverse[columna] = not reverse

        def clave(n: Negocio) -> str | int:
            valores = (
                n.nombre, n.ciudad, n.whatsapp, n.reseñas,
                n.web_instagram, n.link_google_maps, n.direccion,
            )
            val = valores[idx]
            return val if isinstance(val, (int, float)) else str(val).lower()

        self._leads_visibles.sort(key=clave, reverse=reverse)
        self._render_tabla(self._leads_visibles)

    # --- Selección y copiado ---

    def _on_scroll_vertical(self, *args: str) -> None:
        self._ocultar_celda_ui()
        self.tree.yview(*args)

    def _on_scroll_horizontal(self, *args: str) -> None:
        self._ocultar_celda_ui()
        self.tree.xview(*args)

    def _ocultar_celda_ui(self) -> None:
        if self._celda_widget is not None:
            self._celda_widget.destroy()
            self._celda_widget = None
        self._modo_celda = None

    def _bbox_celda(self, row: str, col_idx: int) -> Optional[tuple[int, int, int, int]]:
        bbox = self.tree.bbox(row, f"#{col_idx + 1}")
        return bbox if bbox else None

    def _obtener_texto_celda(self, row: str, col_idx: int) -> str:
        valores = self.tree.item(row, "values")
        if col_idx < 0 or col_idx >= len(valores):
            return ""
        return str(valores[col_idx])

    def _valores_desde_negocio(self, n: Negocio) -> tuple:
        return (
            n.nombre, n.ciudad, n.whatsapp, n.reseñas,
            n.web_instagram, n.link_google_maps, n.direccion,
        )

    def _establecer_valor_celda(self, row: str, col_idx: int, valor: str) -> None:
        negocio = self._iid_a_negocio.get(row)
        if negocio is None:
            return

        campos: list[tuple[str, object]] = [
            ("nombre", valor),
            ("ciudad", valor),
            ("whatsapp", valor),
            ("reseñas", int(valor) if valor.isdigit() else 0),
            ("web_instagram", valor),
            ("link_google_maps", valor),
            ("direccion", valor),
        ]
        attr, val = campos[col_idx]
        setattr(negocio, attr, val)

        valores = list(self._valores_desde_negocio(negocio))
        self.tree.item(row, values=valores)

    def _mostrar_seleccion_celda(self, row: str, col_idx: int, texto: str) -> None:
        bbox = self._bbox_celda(row, col_idx)
        if not bbox:
            return

        self._ocultar_celda_ui()
        self._ultima_celda = (row, col_idx)
        self._modo_celda = "seleccion"
        x, y, w, h = bbox

        frame = tk.Frame(
            self.tree,
            bg="#cce5ff",
            highlightthickness=2,
            highlightbackground="#0078d4",
        )
        label = tk.Label(
            frame,
            text=texto,
            bg="#cce5ff",
            fg="black",
            anchor="center",
            font=("Segoe UI", 9),
            padx=2,
        )
        label.pack(fill="both", expand=True)
        frame.place(x=x, y=y, width=max(w, 40), height=h)

        for widget in (frame, label):
            widget.bind(
                "<Button-1>",
                lambda _e, r=row, c=col_idx: self._manejar_click_celda(r, c),
            )
            widget.bind(
                "<Double-1>",
                lambda _e, r=row, c=col_idx: self._abrir_edicion_celda(r, c),
            )

        self._celda_widget = frame
        self.tree.selection_set(row)
        self.tree.focus_set()

        col_name = self._headers_tabla[COLUMNAS_TABLA[col_idx]]
        self.lbl_estado.config(
            text=f"Celda seleccionada: {col_name} — Ctrl+C copiar · Supr borrar · Doble clic editar",
            foreground="#555",
        )

    def _mostrar_edicion_celda(self, row: str, col_idx: int, texto: str) -> None:
        bbox = self._bbox_celda(row, col_idx)
        if not bbox:
            return

        self._ocultar_celda_ui()
        self._ultima_celda = (row, col_idx)
        self._modo_celda = "edicion"
        x, y, w, h = bbox

        entry = tk.Entry(
            self.tree,
            font=("Segoe UI", 9),
            borderwidth=1,
            relief="solid",
            highlightthickness=2,
            highlightcolor="#0078d4",
            highlightbackground="#0078d4",
            bg="white",
        )
        entry.insert(0, texto)
        entry.place(x=x, y=y, width=max(w, 40), height=h)
        entry.focus_set()
        entry.selection_range(0, tk.END)

        entry.bind("<Return>", lambda _e: self._guardar_edicion_celda())
        entry.bind("<FocusOut>", lambda _e: self.after(80, self._guardar_edicion_celda))
        entry.bind("<Escape>", lambda _e: self._cancelar_edicion_celda(row, col_idx))

        self._celda_widget = entry
        self.tree.selection_set(row)

        col_name = self._headers_tabla[COLUMNAS_TABLA[col_idx]]
        self.lbl_estado.config(
            text=f"Editando: {col_name} — Enter guardar · Escape cancelar",
            foreground="#555",
        )

    def _guardar_edicion_celda(self) -> None:
        if self._edit_guardando:
            return
        if self._modo_celda != "edicion" or not isinstance(self._celda_widget, tk.Entry):
            return
        if not self._ultima_celda:
            return

        self._edit_guardando = True
        try:
            row, col_idx = self._ultima_celda
            valor = self._celda_widget.get()
            self._establecer_valor_celda(row, col_idx, valor)
            self._mostrar_seleccion_celda(row, col_idx, self._obtener_texto_celda(row, col_idx))
        finally:
            self._edit_guardando = False

    def _cancelar_edicion_celda(self, row: str, col_idx: int) -> None:
        texto = self._obtener_texto_celda(row, col_idx)
        self._mostrar_seleccion_celda(row, col_idx, texto)

    def _copiar_celda_seleccionada(self) -> None:
        if not self._ultima_celda:
            return
        row, col_idx = self._ultima_celda
        texto = self._obtener_texto_celda(row, col_idx)
        self.clipboard_clear()
        self.clipboard_append(texto)
        col_name = self._headers_tabla[COLUMNAS_TABLA[col_idx]]
        self.lbl_estado.config(
            text=f"Celda copiada: {col_name}",
            foreground="#1a7f37",
        )

    def _borrar_contenido_celda(self) -> None:
        if not self._ultima_celda or self._modo_celda != "seleccion":
            return
        row, col_idx = self._ultima_celda
        self._establecer_valor_celda(row, col_idx, "")
        self._mostrar_seleccion_celda(row, col_idx, "")
        col_name = self._headers_tabla[COLUMNAS_TABLA[col_idx]]
        self.lbl_estado.config(
            text=f"Contenido borrado: {col_name}",
            foreground="#1a7f37",
        )

    def _obtener_seleccionado(self) -> Optional[Negocio]:
        sel = self.tree.selection()
        if not sel:
            return None
        return self._iid_a_negocio.get(sel[0])

    def _on_seleccion(self, _event: Optional[tk.Event] = None) -> None:
        tiene = self._obtener_seleccionado() is not None
        self.btn_copiar.config(state="normal" if tiene else "disabled")

    def _celda_desde_evento(self, event: tk.Event) -> Optional[tuple[str, int, str]]:
        if self.tree.identify_region(event.x, event.y) != "cell":
            return None
        row = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not row or not col:
            return None
        col_idx = int(col.lstrip("#")) - 1
        valores = self.tree.item(row, "values")
        if col_idx < 0 or col_idx >= len(valores):
            return None
        return row, col_idx, str(valores[col_idx])

    def _es_doble_clic(self, row: str, col_idx: int) -> bool:
        ahora = time.time()
        if (
            self._ultimo_click is not None
            and self._ultimo_click[1] == row
            and self._ultimo_click[2] == col_idx
            and (ahora - self._ultimo_click[0]) < 0.45
        ):
            self._ultimo_click = None
            return True
        self._ultimo_click = (ahora, row, col_idx)
        return False

    def _abrir_edicion_celda(self, row: str, col_idx: int) -> None:
        texto = self._obtener_texto_celda(row, col_idx)
        self._ultimo_click = None
        self._mostrar_edicion_celda(row, col_idx, texto)

    def _manejar_click_celda(self, row: str, col_idx: int) -> None:
        texto = self._obtener_texto_celda(row, col_idx)
        if self._es_doble_clic(row, col_idx):
            self._mostrar_edicion_celda(row, col_idx, texto)
        else:
            self._mostrar_seleccion_celda(row, col_idx, texto)

    def _on_click_celda(self, event: tk.Event) -> None:
        celda = self._celda_desde_evento(event)
        if not celda:
            self._ocultar_celda_ui()
            self._ultima_celda = None
            self._ultimo_click = None
            return

        row, col_idx, _texto = celda
        self._manejar_click_celda(row, col_idx)

    def _on_doble_clic(self, event: tk.Event) -> None:
        celda = self._celda_desde_evento(event)
        if celda:
            row, col_idx, _texto = celda
            self._abrir_edicion_celda(row, col_idx)
        return "break"

    def _on_borrar_celda(self, event: tk.Event) -> Optional[str]:
        if self._modo_celda == "edicion":
            return None
        self._borrar_contenido_celda()
        return "break"

    def _on_copiar_mensaje(self) -> None:
        negocio = self._obtener_seleccionado()
        if negocio:
            self.clipboard_clear()
            self.clipboard_append(negocio.mensaje)
            self.lbl_estado.config(text="Mensaje copiado al portapapeles", foreground="#1a7f37")

    def _on_ctrl_c(self, event: tk.Event) -> Optional[str]:
        if self._modo_celda == "edicion" and isinstance(self._celda_widget, tk.Entry):
            try:
                if self.focus_get() == self._celda_widget:
                    return None
            except KeyError:
                pass
        if self._ultima_celda and self._modo_celda == "seleccion":
            self._copiar_celda_seleccionada()
            return "break"
        self._copiar_fila_seleccionada()
        return "break"

    def _copiar_fila_seleccionada(self, _event: Optional[tk.Event] = None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        valores = self.tree.item(sel[0], "values")
        texto = "\t".join(str(v) for v in valores)
        self.clipboard_clear()
        self.clipboard_append(texto)
        self.lbl_estado.config(text="Fila copiada (Ctrl+C)", foreground="#1a7f37")

    # --- Exportación ---

    def _on_exportar(self) -> None:
        if self._exportando or self._buscando:
            return
        if not self._leads or self._params_busqueda is None:
            messagebox.showinfo("Exportar", "No hay leads para exportar.")
            return

        self._exportando = True
        self.btn_exportar.config(state="disabled")
        self._spinner.start()
        self.lbl_estado.config(text="Exportando a Excel...", foreground="#b8860b")

        threading.Thread(
            target=self._exportar_en_hilo,
            args=(dict(self._params_busqueda),),
            daemon=True,
        ).start()

    def _exportar_en_hilo(self, params: dict) -> None:
        todos: list[Negocio] = []
        paginas_extra: dict[int, list[Negocio]] = {}
        try:
            for p in sorted(self._paginas.keys()):
                todos.extend(self._paginas[p])
            excluir = {n.place_id for n in todos if n.place_id}
            hay_mas = self._hay_mas_paginas
            pagina = max(self._paginas.keys(), default=0) + 1
            while hay_mas:
                leads, hay_mas = self._obtener_leads_pagina(params, excluir)
                if not leads:
                    hay_mas = False
                    break
                paginas_extra[pagina] = leads
                todos.extend(leads)
                for negocio in leads:
                    if negocio.place_id:
                        excluir.add(negocio.place_id)
                pagina += 1

            pais = params["pais"]
            ciudad = params["ciudad"]
            tipo = params["tipo"]
            leads = (
                "todos"
                if params.get("por_pagina", "").strip().lower() == "todos"
                else str(len(todos))
            )
            ahora = datetime.now()
            nombre = (
                f"{_sanitizar_nombre(pais)}-{_sanitizar_nombre(ciudad)}-{leads}-"
                f"{ahora.strftime('%d-%m-%Y')}-{ahora.strftime('%H-%M')}-"
                f"{_sanitizar_nombre(tipo)}.xlsx"
            )
            os.makedirs(INFORMES_DIR, exist_ok=True)
            ruta = os.path.join(INFORMES_DIR, nombre)
            exportar_excel(ruta, todos)
            self.after(
                0,
                lambda: self._finalizar_exportacion(
                    True, ruta, len(todos), None, paginas_extra, hay_mas
                ),
            )
        except OSError as exc:
            self.after(
                0,
                lambda: self._finalizar_exportacion(
                    False, "", len(todos), str(exc), {}, False
                ),
            )

    def _finalizar_exportacion(
        self,
        ok: bool,
        ruta: str,
        cantidad: int,
        error: Optional[str],
        paginas_extra: dict[int, list[Negocio]],
        hay_mas: bool,
    ) -> None:
        for pagina, leads in paginas_extra.items():
            self._paginas[pagina] = leads
            for negocio in leads:
                if negocio.place_id:
                    self._ids_global.add(negocio.place_id)
        if paginas_extra:
            self._hay_mas_paginas = hay_mas
        self._spinner.stop()
        self._exportando = False
        self.btn_exportar.config(state="normal" if self._leads else "disabled")
        if ok:
            messagebox.showinfo(
                "Exportar",
                f"Informe guardado ({cantidad} leads):\n{ruta}",
            )
            self.lbl_estado.config(text=f"Exportado: {ruta}", foreground="#1a7f37")
        else:
            messagebox.showerror("Error", f"No se pudo guardar el archivo:\n{error}")
            self.lbl_estado.config(text="Error al exportar", foreground="#c0392b")


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------

def main() -> None:
    app = ProspeccionApp()
    app.mainloop()


if __name__ == "__main__":
    main()
