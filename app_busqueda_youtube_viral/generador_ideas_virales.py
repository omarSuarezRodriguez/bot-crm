"""
Buscador de Shorts Virales (YouTube) - Por categoría
----------------------------------------------------
Busca Shorts REALES en YouTube y los ordena por la métrica que elijas
(vistas, likes o comentarios) dentro del periodo que elijas, y solo en las
categorías de YouTube que marques (se muestran con su ID y nombre).
La búsqueda se guía únicamente por la categoría de YouTube.
Cada búsqueda con resultados se guarda como informe Excel (.xlsx) en la carpeta
"informes", junto a este script.
No usa ningún LLM: solo búsqueda.

Requisitos:   pip install customtkinter requests openpyxl
Archivo .env (junto a este script):
    YOUTUBE_API_KEY=tu_llave_de_youtube
"""

import math
import os
import queue
import re
import sys
import threading
from datetime import datetime, timedelta, timezone

import customtkinter as ctk
import requests

try:  # el informe Excel es opcional: si falta openpyxl, la búsqueda sigue funcionando
    import openpyxl
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    openpyxl = None

# ----------------------------------------------------------------------------
# Configuración
# ----------------------------------------------------------------------------
YT_BASE = "https://www.googleapis.com/youtube/v3"
MAX_DURACION_SEG = 60       # Shorts clásicos <= 60 s
MAX_PAGINAS_POR_ORDEN = 5   # páginas máximas por cada (categoría, orden)
ORDENES = ("viewCount", "relevance")
CANTIDADES = ("5", "10", "50", "100", "200")

# Ancla técnica (NO es un tema): sesga la búsqueda hacia Shorts, ya que la API
# no tiene un filtro "solo Shorts". La selección real se hace por duración.
ANCLA_SHORTS = "#shorts"

# Categorías de YouTube que puedes marcar en la app: ID -> nombre
CATEGORIAS_YT = {
    "28": "Ciencia y tecnología",
    "27": "Educación",
    "26": "Cómo hacer y estilo",
    "22": "Gente y blogs",
    "24": "Entretenimiento",
    "20": "Videojuegos",
}
CATEGORIAS_YT_DEFECTO = ("28", "27")   # marcadas al abrir la app

CARPETA_INFORMES = "informes"   # se crea junto al script
EXCEL_MAX_CHARS = 32000         # límite de Excel por celda: 32.767

PERIODOS = ("Hoy", "Esta semana", "Este mes", "Este año", "Siempre")
CRITERIOS = {"Vistas": "vistas", "Likes": "likes", "Comentarios": "comentarios"}
IDIOMAS = {"Español": "es", "Inglés": "en"}


class ErrorApp(Exception):
    """Error mostrable al usuario. fatal=True aborta toda la búsqueda."""

    def __init__(self, mensaje, fatal=False):
        super().__init__(mensaje)
        self.fatal = fatal


def cargar_env(ruta=".env"):
    """Lee un archivo .env simple (CLAVE=valor) sin dependencias extra."""
    base = os.path.dirname(os.path.abspath(__file__))
    for r in (ruta, os.path.join(base, ".env")):
        if os.path.exists(r):
            with open(r, encoding="utf-8") as f:
                for linea in f:
                    linea = linea.strip()
                    if linea and not linea.startswith("#") and "=" in linea:
                        k, v = linea.split("=", 1)
                        os.environ.setdefault(k.strip(), v.split("#")[0].strip().strip('"').strip("'"))
            break


# ----------------------------------------------------------------------------
# Búsqueda en YouTube
# ----------------------------------------------------------------------------
def yt_get(endpoint, **params):
    key = os.getenv("YOUTUBE_API_KEY", "").strip()
    if not key:
        raise ErrorApp("Falta YOUTUBE_API_KEY en el archivo .env", fatal=True)
    params["key"] = key
    try:
        r = requests.get(f"{YT_BASE}/{endpoint}", params=params, timeout=30)
    except requests.RequestException as e:
        raise ErrorApp(f"Error de red: {e}", fatal=True)
    if r.status_code != 200:
        try:
            msg = r.json()["error"]["message"]
        except Exception:
            msg = r.text[:200]
        # 401/403 (llave inválida, cuota agotada, API sin habilitar): no tiene sentido seguir
        fatal = r.status_code in (401, 403) or "api key" in msg.lower()
        raise ErrorApp(f"YouTube API ({r.status_code}): {msg}", fatal=fatal)
    return r.json()


def duracion_seg(iso):
    m = re.fullmatch(r"PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?", iso or "")
    if not m:
        return 0
    h, mi, s = (int(x) if x else 0 for x in m.groups())
    return h * 3600 + mi * 60 + s


def fecha_desde(periodo):
    """Inicio del periodo en calendario UTC (no ventanas móviles de N días)."""
    if periodo not in PERIODOS:
        raise ErrorApp(f"Periodo no válido: {periodo}")
    if periodo == "Siempre":
        return None
    ahora = datetime.now(timezone.utc)
    if periodo == "Hoy":
        inicio = ahora.replace(hour=0, minute=0, second=0, microsecond=0)
    elif periodo == "Esta semana":
        inicio = (ahora - timedelta(days=ahora.weekday())).replace(
            hour=0, minute=0, second=0, microsecond=0)
    elif periodo == "Este mes":
        inicio = ahora.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    else:  # Este año
        inicio = ahora.replace(month=1, day=1, hour=0, minute=0, second=0, microsecond=0)
    return inicio.strftime("%Y-%m-%dT%H:%M:%SZ")


# ----------------------------------------------------------------------------
# Detección de idioma (sin dependencias ni cuota extra)
# ----------------------------------------------------------------------------
# relevanceLanguage de YouTube es solo una pista, no un filtro. Por eso se
# detecta el idioma real de cada video con 3 señales, de más a menos fiable:
#   1) idioma de audio declarado   2) idioma del texto   3) idioma por defecto
_PALABRAS_BRUTO = {
    "es": "el la los las del una unos unas por con para como pero sus este esta esto son muy "
          "está tu mi te qué cómo cuando todo hay tiene puedes ahora solo sin sobre usa usar "
          "más también segundos truco trucos aprende herramientas gratis tus mejor mejores "
          "hacer haz puedo nunca siempre estos estas esos esas ese esa nuevo nueva cosas",
    "en": "the and of to in is you that it for with on this are your how what why can be at or "
          "from have not but by an all will just use using now new best tips tricks secret "
          "free learn stop make need know every these those thing things never always",
    "pt": "os as da do das dos um uma não é em com para você que como mais mas seu sua isso "
          "este essa ela ele são está muito também aprenda dicas truque truques segundos",
    "fr": "le les des du un une et est que qui pour pas vous nous dans sur avec ce cette sont "
          "plus très aussi comment astuce astuces secondes apprenez",
    "it": "il lo gli dei delle della che per non una sono come più anche questo questa "
          "con molto trucchi secondi impara",
    "de": "der die das und ist nicht ein eine mit für auf den dem des sich auch wie sie "
          "ich wir sekunden tipps lernen",
}
_PALABRAS = {k: set(v.split()) for k, v in _PALABRAS_BRUTO.items()}
# Solo cuentan las palabras exclusivas de un idioma (descarta las que se comparten)
_PALABRAS = {
    k: {w for w in ws if sum(w in o for o in _PALABRAS.values()) == 1}
    for k, ws in _PALABRAS.items()
}
_IGNORAR_AUDIO = {"", "und", "zxx", "mul"}


def _primario(codigo):
    """'es-419' -> 'es'. Devuelve '' si el código es vacío o no informativo."""
    c = (codigo or "").strip().lower().replace("_", "-").split("-")[0]
    return "" if c in _IGNORAR_AUDIO else c


def _idioma_por_texto(texto):
    """'es', 'en', ... si está claro; 'otro' si es de otra escritura; '?' si no se sabe."""
    texto = re.sub(r"https?://\S+|www\.\S+|[#@]\w+", " ", texto or "").lower()
    letras = re.findall(r"[^\W\d_]+", texto)[:150]   # limpiar primero y luego recortar
    texto = " ".join(letras)
    if not letras:
        return "?"
    total = sum(len(x) for x in letras)
    no_latinas = sum(1 for x in letras for ch in x if ord(ch) > 0x024F)
    if total and no_latinas / total > 0.5:
        return "otro"
    puntos = {k: sum(1 for w in letras if w in ws) for k, ws in _PALABRAS.items()}
    puntos["es"] += 2 * len(re.findall(r"[ñ¿¡]", texto))
    ordenados = sorted(puntos.values(), reverse=True)
    mejor = max(puntos, key=puntos.get)
    if ordenados[0] >= 1 and ordenados[0] > ordenados[1]:
        return mejor
    return "?"


def detectar_idioma(sn):
    """Código de idioma del video ('es', 'en', 'pt', 'otro'...) o None si no se puede saber."""
    audio = _primario(sn.get("defaultAudioLanguage"))
    if audio:
        return audio
    texto = " ".join([sn.get("title") or "", (sn.get("description") or "")[:3000],
                      " ".join(sn.get("tags") or [])])
    por_texto = _idioma_por_texto(texto)
    if por_texto != "?":
        return por_texto
    # defaultLanguage es poco fiable (muchos creadores lo dejan en 'en'): solo como último recurso
    return _primario(sn.get("defaultLanguage")) or None


def nombre_categoria(cid):
    if not cid:
        return "sin categoría"
    return f"{cid} · {CATEGORIAS_YT.get(cid, 'Otra categoría')}"


def plan_busqueda(cantidad, n_cats):
    """(objetivo por grupo, nº de búsquedas, cuota máxima estimada en unidades)."""
    grupos = max(n_cats, 1)
    pool = min(max(cantidad * 3, 50), MAX_PAGINAS_POR_ORDEN * 50)
    objetivo_grupo = max(math.ceil(pool / grupos), 25)
    paginas = min(math.ceil(objetivo_grupo / 50), MAX_PAGINAS_POR_ORDEN)
    busquedas = grupos * len(ORDENES) * paginas
    # cada search cuesta 100; cada lote de videos.list cuesta 1 (como mucho 1 lote por search)
    return objetivo_grupo, busquedas, busquedas * 101


def _recolectar_ids(orden, desde, codigo_idioma, objetivo, categoria,
                    progreso, base_frac, span_frac, etiqueta):
    """Pagina search.list (filtrado por categoría) y acumula videoIds únicos."""
    ids, token, paginas = [], None, 0
    while len(ids) < objetivo and paginas < MAX_PAGINAS_POR_ORDEN:
        p = dict(
            part="snippet",
            q=ANCLA_SHORTS,
            type="video",
            videoDuration="short",
            videoCategoryId=categoria,
            order=orden,
            maxResults=50,
            relevanceLanguage=codigo_idioma,
        )
        if desde:
            p["publishedAfter"] = desde
        if token:
            p["pageToken"] = token
        data = yt_get("search", **p)
        for it in data.get("items", []):
            vid = (it.get("id") or {}).get("videoId")
            if vid and vid not in ids:
                ids.append(vid)
        paginas += 1
        token = data.get("nextPageToken")
        progreso(
            base_frac + span_frac * min(len(ids) / max(objetivo, 1), 1),
            f"Analizando tendencias actuales en redes... ({etiqueta} · {orden})",
        )
        if not token:
            break
    return ids


def buscar_shorts(cantidad, periodo, idioma, criterio, categorias, progreso):
    """
    categorias: lista de (id, nombre); debe tener al menos una.
    Devuelve (videos, avisos).
    """
    if criterio not in CRITERIOS:
        raise ErrorApp(f"Criterio no válido: {criterio}")
    if idioma not in IDIOMAS:
        raise ErrorApp(f"Idioma no válido: {idioma}")
    cantidad = int(cantidad)
    if cantidad < 1:
        raise ErrorApp("La cantidad debe ser al menos 1.")
    grupos = list(categorias or [])
    if not grupos:
        raise ErrorApp("Marca al menos una categoría de YouTube.")

    avisos = []
    desde = fecha_desde(periodo)
    codigo_idioma = IDIOMAS[idioma]
    objetivo_grupo, _, _ = plan_busqueda(cantidad, len(grupos))

    # 1) Candidatos: una búsqueda por (categoría, orden); search no ordena por likes/comentarios
    pasos = [(cid, nom, orden) for cid, nom in grupos for orden in ORDENES]
    ids, vistos = [], set()
    for k, (cid, nom, orden) in enumerate(pasos):
        etiqueta = f"{cid} · {nom}"
        try:
            nuevos = _recolectar_ids(
                orden, desde, codigo_idioma, objetivo_grupo, cid, progreso,
                0.05 + 0.50 * k / len(pasos), 0.50 / len(pasos), etiqueta)
        except ErrorApp as e:
            if e.fatal:
                raise
            avisos.append(f"Búsqueda omitida ({etiqueta} · {orden}): {e}")
            continue
        for vid in nuevos:
            if vid not in vistos:
                vistos.add(vid)
                ids.append(vid)

    if not ids:
        if avisos:
            raise ErrorApp(avisos[-1])
        return [], avisos
    progreso(0.58, f"Analizando tendencias actuales en redes... ({len(ids)} candidatos únicos)")

    # 2) Métricas reales, categoría real y filtro por duración de Short
    videos = []
    for i in range(0, len(ids), 50):
        lote = ids[i:i + 50]
        data = yt_get("videos", part="snippet,statistics,contentDetails", id=",".join(lote))
        for v in data.get("items", []):
            seg = duracion_seg((v.get("contentDetails") or {}).get("duration"))
            if not (0 < seg <= MAX_DURACION_SEG):
                continue
            sn = v.get("snippet") or {}
            st = v.get("statistics") or {}
            likes = st.get("likeCount")
            comentarios = st.get("commentCount")
            videos.append({
                "titulo": sn.get("title", ""),
                "canal": sn.get("channelTitle", ""),
                "fecha": (sn.get("publishedAt") or "")[:10],
                "categoria": sn.get("categoryId"),
                "idioma": detectar_idioma(sn),
                "duracion": seg,
                "vistas": int(st.get("viewCount", 0) or 0),
                "likes": int(likes) if likes is not None else None,
                "comentarios": int(comentarios) if comentarios is not None else None,
                "url": f"https://youtube.com/shorts/{v['id']}",
            })
        progreso(0.60 + 0.30 * min((i + len(lote)) / len(ids), 1),
                 "Midiendo métricas reales de cada video...")

    # 3) Confirmar la categoría con el dato real del video (la búsqueda de YouTube a veces se cuela)
    validas = {cid for cid, _ in grupos}
    antes = len(videos)
    videos = [v for v in videos if v["categoria"] in validas]
    if antes - len(videos):
        avisos.append(f"{antes - len(videos)} videos descartados por no estar en las categorías marcadas.")

    # 4) Filtro ESTRICTO de idioma: solo el elegido; si no se puede determinar, se descarta
    antes = len(videos)
    desconocido = sum(1 for v in videos if v["idioma"] is None)
    videos = [v for v in videos if v["idioma"] == codigo_idioma]
    if antes - len(videos):
        avisos.append(f"{antes - len(videos)} videos descartados por no estar en {idioma.lower()}"
                      + (f" ({desconocido} por idioma indeterminado)." if desconocido else "."))

    # 5) Orden descendente real; una métrica oculta (None) no entra en su propio ranking
    clave = CRITERIOS[criterio]
    if clave in ("likes", "comentarios"):
        ocultos = sum(1 for v in videos if v[clave] is None)
        videos = [v for v in videos if v[clave] is not None]
        if ocultos:
            avisos.append(f"{ocultos} videos omitidos por tener {clave} ocultos.")
    videos.sort(key=lambda x: (x[clave] if x[clave] is not None else -1, x["vistas"]), reverse=True)
    return videos[:cantidad], avisos


# ----------------------------------------------------------------------------
# Informes en Excel
# ----------------------------------------------------------------------------
def carpeta_base():
    """Carpeta donde vive el .py (o el .exe si se empaqueta con PyInstaller)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _poner_texto(ws, fila, col, valor):
    """Escribe SIEMPRE como texto: evita que un título que empiece por '=' se ejecute como
    fórmula, quita caracteres ilegales para Excel y respeta el límite de longitud."""
    s = ILLEGAL_CHARACTERS_RE.sub("", "" if valor is None else str(valor))[:EXCEL_MAX_CHARS]
    celda = ws.cell(row=fila, column=col, value=s)
    celda.data_type = "s"
    return celda


def _ruta_unica(carpeta, nombre):
    base, ext = os.path.splitext(nombre)
    ruta, n = os.path.join(carpeta, nombre), 2
    while os.path.exists(ruta):
        ruta = os.path.join(carpeta, f"{base}_{n}{ext}")
        n += 1
    return ruta


def guardar_informe(videos, cfg, avisos, base=None):
    """Guarda la búsqueda en informes/<fecha>.xlsx y devuelve la ruta. Lanza ErrorApp si falla."""
    if openpyxl is None:
        raise ErrorApp("Falta la librería openpyxl (pip install openpyxl).")
    carpeta = os.path.join(base or carpeta_base(), CARPETA_INFORMES)
    try:
        os.makedirs(carpeta, exist_ok=True)
    except OSError as e:
        raise ErrorApp(f"No se pudo crear la carpeta '{CARPETA_INFORMES}': {e}")

    ahora = datetime.now()
    clave = CRITERIOS[cfg["criterio"]]
    nombre = f"informe_{ahora:%Y-%m-%d_%H-%M-%S}_{clave}_{IDIOMAS[cfg['idioma']]}.xlsx"
    ruta = _ruta_unica(carpeta, nombre)

    wb = openpyxl.Workbook()

    # --- Hoja 1: Resultados ---
    ws = wb.active
    ws.title = "Resultados"
    columnas = [  # (encabezado, ancho, clave en el dict del video)
        ("#", 6, None), ("Título", 60, "titulo"), ("Canal", 28, "canal"),
        ("Categoría", 24, "categoria"), ("Idioma", 9, "idioma"), ("Duración (s)", 13, "duracion"),
        ("Publicado", 12, "fecha"), ("Vistas", 13, "vistas"), ("Likes", 11, "likes"),
        ("Comentarios", 13, "comentarios"), ("Enlace", 38, "url"),
    ]
    relleno = PatternFill("solid", fgColor="1F4E78")
    for c, (enc, ancho, k) in enumerate(columnas, 1):
        marca = " ★" if k == clave else ""
        celda = _poner_texto(ws, 1, c, enc + marca)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = relleno
        celda.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(c)].width = ancho
    for f, v in enumerate(videos, 2):
        ws.cell(row=f, column=1, value=f - 1)
        _poner_texto(ws, f, 2, v["titulo"])
        _poner_texto(ws, f, 3, v["canal"])
        _poner_texto(ws, f, 4, nombre_categoria(v["categoria"]))
        _poner_texto(ws, f, 5, v["idioma"] or "")
        ws.cell(row=f, column=6, value=v["duracion"])
        try:
            d = ws.cell(row=f, column=7, value=datetime.strptime(v["fecha"], "%Y-%m-%d").date())
            d.number_format = "yyyy-mm-dd"
        except ValueError:
            _poner_texto(ws, f, 7, v["fecha"])
        for col, k in ((8, "vistas"), (9, "likes"), (10, "comentarios")):
            ws.cell(row=f, column=col, value=v[k]).number_format = "#,##0"   # None = oculto -> vacío
        enlace = _poner_texto(ws, f, 11, v["url"])
        enlace.hyperlink = v["url"]
        enlace.style = "Hyperlink"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columnas))}{max(len(videos) + 1, 2)}"

    # --- Hoja 2: Resumen ---
    rs = wb.create_sheet("Resumen")
    filas = [
        ("Fecha del informe", f"{ahora:%Y-%m-%d %H:%M:%S}"),
        ("Categorías", ", ".join(f"{cid} · {nom}" for cid, nom in cfg["categorias"])),
        ("Periodo (calendario UTC)", cfg["periodo"]),
        ("Idioma", cfg["idioma"]),
        ("Ordenado por", cfg["criterio"]),
        ("Cantidad pedida", cfg["cantidad"]),
        ("Cantidad encontrada", len(videos)),
        ("Nota", "Likes/comentarios ocultos por el creador se muestran como celda vacía."),
    ]
    filas += [("Aviso", a) for a in avisos]
    for f, (etiqueta, valor) in enumerate(filas, 1):
        _poner_texto(rs, f, 1, etiqueta).font = Font(bold=True)
        celda = _poner_texto(rs, f, 2, valor)
        celda.alignment = Alignment(wrap_text=True, vertical="top")
    rs.column_dimensions["A"].width = 26
    rs.column_dimensions["B"].width = 100

    # Guardado atómico: un archivo a medias nunca queda con el nombre final
    tmp = os.path.join(carpeta, f".tmp_{os.getpid()}_{os.path.basename(ruta)}")
    try:
        wb.save(tmp)
        os.replace(tmp, ruta)
    except Exception as e:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass
        raise ErrorApp(f"No se pudo guardar el informe: {e}")
    return ruta


# ----------------------------------------------------------------------------
# Formato de salida
# ----------------------------------------------------------------------------
def fmt(n):
    if n >= 1_000_000:
        return f"{n / 1_000_000:.1f}M"
    if n >= 1_000:
        return f"{n / 1_000:.1f}K"
    return str(n)


def _fmt_metrica(valor):
    return "oculto" if valor is None else fmt(valor)


def formatear(videos, criterio):
    clave = CRITERIOS[criterio]
    etiquetas = [("vistas", "👁", "vistas"), ("likes", "👍", "likes"), ("comentarios", "💬", "comentarios")]
    salida = []
    for i, v in enumerate(videos, 1):
        metricas = " · ".join(
            f"{'⭐ ' if k == clave else ''}{icono} {_fmt_metrica(v[k])} {nombre}"
            for k, icono, nombre in etiquetas)
        salida.append(
            f"#{i} · {metricas}\n"
            f"📌 {v['titulo']}\n"
            f"📺 {v['canal']} · {v['duracion']} s · publicado {v['fecha']} · 🌐 {v['idioma']}\n"
            f"🏷 {nombre_categoria(v['categoria'])}\n"
            f"🔗 {v['url']}\n"
            f"{'─' * 60}\n")
    return "\n".join(salida)


# ----------------------------------------------------------------------------
# Interfaz
# ----------------------------------------------------------------------------
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.title("Buscador de Shorts Virales · Por categoría")
        self.geometry("1150x780")
        self.minsize(980, 640)
        self.cola = queue.Queue()
        self.ocupado = False

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._panel_izquierdo()
        self._panel_derecho()
        self.after(100, self.procesar_cola)

    # -- Panel izquierdo: configuración (con scroll) + botón fijo abajo -------
    def _panel_izquierdo(self):
        cont = ctk.CTkFrame(self, width=350, corner_radius=0)
        cont.grid(row=0, column=0, sticky="nsew")
        cont.grid_propagate(False)
        cont.grid_rowconfigure(0, weight=1)
        cont.grid_columnconfigure(0, weight=1)

        p = ctk.CTkScrollableFrame(cont, fg_color="transparent")
        p.grid(row=0, column=0, sticky="nsew")

        self.btn = ctk.CTkButton(
            cont, text="🔥 Buscar Tendencias Virales", height=56,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="#e8590c", hover_color="#c24a0a", command=self.iniciar)
        self.btn.grid(row=1, column=0, sticky="ew", padx=16, pady=14)

        ctk.CTkLabel(p, text="⚙️ Configuración de Búsqueda",
                     font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=10, pady=(10, 6))

        ctk.CTkLabel(p, text="Cantidad de resultados").pack(anchor="w", padx=10, pady=(8, 0))
        self.cmb_cant = ctk.CTkComboBox(p, values=list(CANTIDADES), state="readonly")
        self.cmb_cant.set("10")
        self.cmb_cant.pack(fill="x", padx=10, pady=4)

        ctk.CTkLabel(p, text="Ordenar por").pack(anchor="w", padx=10, pady=(8, 0))
        self.cmb_criterio = ctk.CTkComboBox(p, values=list(CRITERIOS), state="readonly")
        self.cmb_criterio.set("Vistas")
        self.cmb_criterio.pack(fill="x", padx=10, pady=4)

        ctk.CTkLabel(p, text="Periodo").pack(anchor="w", padx=10, pady=(8, 0))
        self.cmb_periodo = ctk.CTkComboBox(p, values=list(PERIODOS), state="readonly")
        self.cmb_periodo.set("Esta semana")
        self.cmb_periodo.pack(fill="x", padx=10, pady=4)

        ctk.CTkLabel(p, text="Idioma de los videos").pack(anchor="w", padx=10, pady=(8, 0))
        self.cmb_idioma = ctk.CTkComboBox(p, values=list(IDIOMAS), state="readonly")
        self.cmb_idioma.set("Español")
        self.cmb_idioma.pack(fill="x", padx=10, pady=4)

        ctk.CTkLabel(p, text="Categoría de YouTube (ID · nombre)").pack(anchor="w", padx=10, pady=(14, 0))
        self.vars_cat = {}
        for cid, nombre in CATEGORIAS_YT.items():
            var = ctk.IntVar(value=1 if cid in CATEGORIAS_YT_DEFECTO else 0)
            self.vars_cat[cid] = var
            ctk.CTkCheckBox(p, text=f"{cid} · {nombre}", variable=var).pack(anchor="w", padx=10, pady=2)

        fila = ctk.CTkFrame(p, fg_color="transparent")
        fila.pack(fill="x", padx=10, pady=(6, 0))
        ctk.CTkButton(fila, text="Todas", width=80, command=lambda: self._marcar_cats(1)).pack(side="left")
        ctk.CTkButton(fila, text="Ninguna", width=80, command=lambda: self._marcar_cats(0)).pack(side="left", padx=6)
        ctk.CTkLabel(
            p, text="Marca al menos una categoría. "
                    "Cada categoría marcada añade búsquedas (más cuota).",
            wraplength=290, justify="left", text_color="gray",
        ).pack(anchor="w", padx=10, pady=(6, 10))

    def _marcar_cats(self, valor):
        for var in self.vars_cat.values():
            var.set(valor)

    # -- Panel derecho: progreso y resultados -------------------------------
    def _panel_derecho(self):
        p = ctk.CTkFrame(self, fg_color="transparent")
        p.grid(row=0, column=1, sticky="nsew", padx=16, pady=16)
        p.grid_columnconfigure(0, weight=1)
        p.grid_rowconfigure(2, weight=1)

        cab = ctk.CTkFrame(p, fg_color="transparent")
        cab.grid(row=0, column=0, sticky="ew")
        cab.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(cab, text="📈 Resultados",
                     font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, sticky="w")
        ctk.CTkButton(cab, text="Copiar todo", width=110, command=self.copiar).grid(row=0, column=1)

        estado = ctk.CTkFrame(p, fg_color="transparent")
        estado.grid(row=1, column=0, sticky="ew", pady=(10, 8))
        estado.grid_columnconfigure(0, weight=1)
        self.lbl_estado = ctk.CTkLabel(
            estado, text="Listo. El periodo usa calendario UTC (a inicios de mes hay menos Shorts).", anchor="w")
        self.lbl_estado.grid(row=0, column=0, sticky="ew")
        self.barra = ctk.CTkProgressBar(estado)
        self.barra.set(0)
        self.barra.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        self.resultados = ctk.CTkTextbox(p, wrap="word", font=ctk.CTkFont(size=14))
        self.resultados.grid(row=2, column=0, sticky="nsew")

    # -- Lógica --------------------------------------------------------------
    def iniciar(self):
        if self.ocupado:
            return
        if not os.getenv("YOUTUBE_API_KEY", "").strip():
            self.lbl_estado.configure(text="⚠️ Falta YOUTUBE_API_KEY en el archivo .env")
            return
        try:
            cantidad = int(self.cmb_cant.get())
        except ValueError:
            self.lbl_estado.configure(text="⚠️ Cantidad no válida.")
            return

        categorias = [(cid, CATEGORIAS_YT[cid]) for cid, var in self.vars_cat.items() if var.get() == 1]
        if not categorias:
            self.lbl_estado.configure(text="⚠️ Marca al menos una categoría de YouTube.")
            return

        cfg = dict(cantidad=cantidad, periodo=self.cmb_periodo.get(), idioma=self.cmb_idioma.get(),
                   criterio=self.cmb_criterio.get(), categorias=categorias)

        _, _, cuota = plan_busqueda(cantidad, len(categorias))
        self.ocupado = True
        self.btn.configure(state="disabled", text="Buscando...")
        self.resultados.delete("1.0", "end")
        self.barra.set(0)
        self.lbl_estado.configure(
            text=f"Analizando tendencias actuales en redes... (cuota máx. estimada ~{cuota:,} de 10.000)".replace(",", "."))
        threading.Thread(target=self.trabajo, args=(cfg,), daemon=True).start()

    def trabajo(self, cfg):
        q = self.cola

        def progreso(frac, texto):
            q.put(("estado", texto, float(frac)))

        try:
            videos, avisos = buscar_shorts(
                cfg["cantidad"], cfg["periodo"], cfg["idioma"],
                cfg["criterio"], cfg["categorias"], progreso)
            if not videos:
                detalle = ("\n" + "\n".join(avisos)) if avisos else ""
                raise ErrorApp("No se encontraron Shorts para ese periodo y esas categorías. "
                               "Prueba más categorías o un periodo más amplio." + detalle)
            if len(videos) < cfg["cantidad"]:
                avisos.insert(0, f"Solo se encontraron {len(videos)} Shorts reales que cumplen los filtros "
                                 f"(pediste {cfg['cantidad']}). Prueba un periodo más amplio o más categorías.")
            ruta = None
            try:
                ruta = guardar_informe(videos, cfg, avisos)
            except Exception as e:  # el informe nunca debe tumbar la búsqueda
                avisos.append(f"No se guardó el informe Excel: {e}")
            cabecera = "".join(f"ℹ️ {a}\n" for a in avisos)
            q.put(("texto", (cabecera + "\n" if cabecera else "") + formatear(videos, cfg["criterio"])))
            guardado = f" · 📁 {CARPETA_INFORMES}/{os.path.basename(ruta)}" if ruta else ""
            progreso(1.0, f"✅ Listo: {len(videos)} Shorts ordenados por {cfg['criterio'].lower()} "
                          f"({cfg['periodo'].lower()}){guardado}")
            q.put(("fin",))
        except Exception as e:
            q.put(("error", str(e)))

    def procesar_cola(self):
        try:
            while True:
                tipo, *datos = self.cola.get_nowait()
                if tipo == "estado":
                    self.lbl_estado.configure(text=datos[0])
                    self.barra.set(datos[1])
                elif tipo == "texto":
                    self.resultados.insert("end", datos[0])
                elif tipo == "fin":
                    self._terminar()
                elif tipo == "error":
                    self.lbl_estado.configure(text=f"❌ {datos[0]}")
                    self._terminar()
        except queue.Empty:
            pass
        self.after(100, self.procesar_cola)

    def _terminar(self):
        self.ocupado = False
        self.btn.configure(state="normal", text="🔥 Buscar Tendencias Virales")

    def copiar(self):
        texto = self.resultados.get("1.0", "end").strip()
        if not texto:
            self.lbl_estado.configure(text="⚠️ No hay resultados para copiar.")
            return
        self.clipboard_clear()
        self.clipboard_append(texto)
        self.lbl_estado.configure(text="📋 Resultados copiados al portapapeles.")


if __name__ == "__main__":
    cargar_env()
    App().mainloop()