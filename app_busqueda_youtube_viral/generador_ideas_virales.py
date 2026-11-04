"""
Buscador de Shorts Virales (YouTube) - Programación y Tecnología
----------------------------------------------------------------
Busca Shorts REALES en YouTube y los ordena por la métrica que elijas
(vistas, likes o comentarios) dentro del periodo que elijas.
No usa ningún LLM: solo búsqueda.

Requisitos:   pip install customtkinter requests
Archivo .env (junto a este script):
    YOUTUBE_API_KEY=tu_llave_de_youtube
"""

import os
import queue
import re
import threading
from datetime import datetime, timedelta, timezone

import customtkinter as ctk
import requests

# ----------------------------------------------------------------------------
# Configuración
# ----------------------------------------------------------------------------
YT_BASE = "https://www.googleapis.com/youtube/v3"
MAX_DURACION_SEG = 60   # Shorts clásicos ≤ 60 s
MAX_PAGINAS_POR_ORDEN = 5  # por cada modo de búsqueda (relevance / viewCount)
CANTIDADES = ("5", "10", "50", "100", "200")
# Solo en categorías predefinidas (tema libre NO fuerza categoría)
VIDEO_CATEGORY_TECH = "28"

PERIODOS = ("Hoy", "Esta semana", "Este mes", "Este año", "Siempre")
CRITERIOS = {"Vistas": "vistas", "Likes": "likes", "Comentarios": "comentarios"}
IDIOMAS = {"Español": "es", "Inglés": "en"}

CATEGORIAS = {
    "Trucos de PC": "trucos de PC windows tips",
    "Herramientas IA": "herramientas de inteligencia artificial AI tools",
    "Atajos productivos": "atajos de teclado productividad shortcuts",
    "Ciberseguridad": "ciberseguridad hacking tips",
    "Setups tech": "setup tech escritorio programador",
}


class ErrorApp(Exception):
    pass


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
        raise ErrorApp("Falta YOUTUBE_API_KEY en el archivo .env")
    params["key"] = key
    r = requests.get(f"{YT_BASE}/{endpoint}", params=params, timeout=30)
    if r.status_code != 200:
        try:
            msg = r.json()["error"]["message"]
        except Exception:
            msg = r.text[:200]
        raise ErrorApp(f"YouTube API ({r.status_code}): {msg}")
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
            hour=0, minute=0, second=0, microsecond=0
        )
    elif periodo == "Este mes":
        inicio = ahora.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    elif periodo == "Este año":
        inicio = ahora.replace(month=1, day=1, hour=0, minute=0, second=0, microsecond=0)
    else:
        raise ErrorApp(f"Periodo no válido: {periodo}")
    return inicio.strftime("%Y-%m-%dT%H:%M:%SZ")


_STOP = {
    "shorts", "short", "para", "como", "cómo", "with", "from", "this", "that",
    "your", "about", "free", "gratis", "the", "and", "for", "del", "los",
    "las", "una", "uno", "que", "con", "por", "de", "en", "el", "la",
}
_CORTOS = {"ai", "ia", "pc", "os", "ui", "ux", "api", "sql", "css", "js"}


def _stem_es(tok):
    """Stemming mínimo ES/EN para plurales comunes."""
    if len(tok) <= 4:
        return tok
    for suf in ("ciones", "ción", "ments", "ment", "ies", "es", "s"):
        if tok.endswith(suf) and len(tok) - len(suf) >= 4:
            return tok[: -len(suf)]
    return tok


def _tokens_consulta(consulta):
    toks = []
    for w in re.findall(r"[a-záéíóúñü0-9]+", consulta.lower(), flags=re.IGNORECASE):
        if w in _CORTOS:
            toks.append(w)
        elif len(w) > 2 and w not in _STOP:
            toks.append(w)
    # dedupe preservando orden
    vistos, out = set(), []
    for t in toks:
        if t not in vistos:
            vistos.add(t)
            out.append(t)
    return out


def _match_token(texto, tok):
    """Match por palabra completa o prefijo del stem (herramienta/herramientas)."""
    stem = _stem_es(tok)
    patrones = {tok, stem}
    if tok in {"ai", "ia"}:
        patrones.update({"ai", "ia"})
    return any(
        re.search(rf"(?<![a-z0-9áéíóúñü]){re.escape(p)}[a-z0-9áéíóúñü]*", texto)
        for p in patrones
        if len(p) >= 2
    )


def _parece_relevante(titulo, canal, consulta):
    """
    Filtro de tema: si hay ≥2 tokens, exige al menos 2 coincidencias
    (evita que solo 'ai' deje pasar cualquier canal con AI).
    Con 1 token, basta 1 coincidencia.
    """
    toks = _tokens_consulta(consulta)
    if not toks:
        return True
    texto = f"{titulo} {canal}".lower()
    hits = sum(1 for t in toks if _match_token(texto, t))
    necesario = 2 if len(toks) >= 2 else 1
    return hits >= necesario


def _recolectar_ids(consulta, orden, desde, codigo_idioma, objetivo, filtrar_tech, progreso, base_frac, span_frac):
    """Página search.list y acumula videoIds únicos."""
    ids, token, paginas = [], None, 0
    while len(ids) < objetivo and paginas < MAX_PAGINAS_POR_ORDEN:
        p = dict(
            part="snippet",
            q=f"{consulta} shorts",
            type="video",
            videoDuration="short",
            order=orden,
            maxResults=50,
            relevanceLanguage=codigo_idioma,
        )
        if filtrar_tech:
            p["videoCategoryId"] = VIDEO_CATEGORY_TECH
        if desde:
            p["publishedAfter"] = desde
        if token:
            p["pageToken"] = token
        data = yt_get("search", **p)
        for it in data.get("items", []):
            vid = it.get("id", {}).get("videoId")
            if vid and vid not in ids:
                ids.append(vid)
        paginas += 1
        token = data.get("nextPageToken")
        progreso(
            base_frac + span_frac * min(len(ids) / max(objetivo, 1), 1),
            f"Analizando tendencias actuales en redes... ({len(ids)} candidatos · {orden})",
        )
        if not token:
            break
    return ids


def buscar_shorts(consulta, cantidad, periodo, idioma, criterio, progreso, filtrar_tech=False):
    if criterio not in CRITERIOS:
        raise ErrorApp(f"Criterio no válido: {criterio}")
    if idioma not in IDIOMAS:
        raise ErrorApp(f"Idioma no válido: {idioma}")
    cantidad = int(cantidad)
    if cantidad < 1:
        raise ErrorApp("La cantidad debe ser al menos 1.")

    # Pool amplio: search no ordena por likes/comentarios; fusionamos relevance ∪ viewCount
    # y luego ordenamos con estadísticas reales (videos.list).
    objetivo = min(max(cantidad * 2, 40), MAX_PAGINAS_POR_ORDEN * 50)
    desde = fecha_desde(periodo)
    codigo_idioma = IDIOMAS[idioma]
    consulta = (consulta or "").strip()
    if not consulta:
        raise ErrorApp("La consulta de búsqueda está vacía.")

    ids_vista = _recolectar_ids(
        consulta, "viewCount", desde, codigo_idioma, objetivo, filtrar_tech, progreso, 0.05, 0.25
    )
    ids_rel = _recolectar_ids(
        consulta, "relevance", desde, codigo_idioma, objetivo, filtrar_tech, progreso, 0.30, 0.25
    )
    ids = []
    for vid in ids_vista + ids_rel:
        if vid not in ids:
            ids.append(vid)

    progreso(0.58, f"Analizando tendencias actuales en redes... ({len(ids)} candidatos únicos)")
    if not ids:
        return []

    # 2) Métricas reales y filtro por duración de Short
    videos = []
    total_ids = len(ids)
    for i in range(0, total_ids, 50):
        lote = ids[i:i + 50]
        data = yt_get("videos", part="snippet,statistics,contentDetails", id=",".join(lote))
        for v in data.get("items", []):
            details = v.get("contentDetails", {})
            seg = duracion_seg(details.get("duration"))
            if not (0 < seg <= MAX_DURACION_SEG):
                continue
            sn = v.get("snippet", {})
            st = v.get("statistics", {})
            likes = st.get("likeCount")
            comentarios = st.get("commentCount")
            videos.append({
                "titulo": sn.get("title", ""),
                "canal": sn.get("channelTitle", ""),
                "fecha": (sn.get("publishedAt") or "")[:10],
                "duracion": seg,
                "vistas": int(st.get("viewCount", 0) or 0),
                "likes": int(likes) if likes is not None else None,
                "comentarios": int(comentarios) if comentarios is not None else None,
                "url": f"https://youtube.com/shorts/{v['id']}",
            })
        progreso(
            0.60 + 0.30 * min((i + len(lote)) / max(total_ids, 1), 1),
            "Midiendo métricas reales de cada video...",
        )

    # 3) Preferir resultados alineados al tema; si el filtro vacía, se conserva el pool
    alineados = [v for v in videos if _parece_relevante(v["titulo"], v["canal"], consulta)]
    if alineados:
        videos = alineados

    # 4) Orden descendente real; excluye métrica oculta (None) del ranking de esa métrica
    clave = CRITERIOS[criterio]
    if clave in ("likes", "comentarios"):
        videos = [v for v in videos if v[clave] is not None]
    videos.sort(
        key=lambda x: (x[clave] if x[clave] is not None else -1, x["vistas"]),
        reverse=True,
    )
    return videos[:cantidad]


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
    if valor is None:
        return "oculto"
    return fmt(valor)


def formatear(videos, criterio):
    clave = CRITERIOS[criterio]
    etiquetas = [
        ("vistas", "👁", "vistas"),
        ("likes", "👍", "likes"),
        ("comentarios", "💬", "comentarios"),
    ]
    salida = []
    for i, v in enumerate(videos, 1):
        metricas = " · ".join(
            f"{'⭐ ' if k == clave else ''}{icono} {_fmt_metrica(v[k])} {nombre}"
            for k, icono, nombre in etiquetas
        )
        salida.append(
            f"#{i} · {metricas}\n"
            f"📌 {v['titulo']}\n"
            f"📺 {v['canal']} · {v['duracion']} s · publicado {v['fecha']}\n"
            f"🔗 {v['url']}\n"
            f"{'─' * 60}\n"
        )
    return "\n".join(salida)


# ----------------------------------------------------------------------------
# Interfaz
# ----------------------------------------------------------------------------
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.title("Buscador de Shorts Virales · Programación y Tecnología")
        self.geometry("1100x720")
        self.minsize(960, 600)
        self.cola = queue.Queue()
        self.ocupado = False

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._panel_izquierdo()
        self._panel_derecho()
        self.after(100, self.procesar_cola)

    # -- Panel izquierdo: configuración -------------------------------------
    def _panel_izquierdo(self):
        p = ctk.CTkFrame(self, width=330, corner_radius=0)
        p.grid(row=0, column=0, sticky="nsew")
        p.grid_propagate(False)

        ctk.CTkLabel(
            p,
            text="⚙️ Configuración de Búsqueda",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(anchor="w", padx=20, pady=(20, 10))

        ctk.CTkLabel(p, text="Cantidad de resultados").pack(anchor="w", padx=20, pady=(10, 0))
        self.cmb_cant = ctk.CTkComboBox(p, values=list(CANTIDADES), state="readonly")
        self.cmb_cant.set("10")
        self.cmb_cant.pack(fill="x", padx=20, pady=(4, 8))

        ctk.CTkLabel(p, text="Ordenar por").pack(anchor="w", padx=20, pady=(8, 0))
        self.cmb_criterio = ctk.CTkComboBox(p, values=list(CRITERIOS), state="readonly")
        self.cmb_criterio.set("Vistas")
        self.cmb_criterio.pack(fill="x", padx=20, pady=4)

        ctk.CTkLabel(p, text="Periodo").pack(anchor="w", padx=20, pady=(8, 0))
        self.cmb_periodo = ctk.CTkComboBox(p, values=list(PERIODOS), state="readonly")
        self.cmb_periodo.set("Esta semana")
        self.cmb_periodo.pack(fill="x", padx=20, pady=4)

        ctk.CTkLabel(p, text="Idioma de los videos").pack(anchor="w", padx=20, pady=(8, 0))
        self.cmb_idioma = ctk.CTkComboBox(p, values=list(IDIOMAS), state="readonly")
        self.cmb_idioma.set("Español")
        self.cmb_idioma.pack(fill="x", padx=20, pady=4)

        ctk.CTkLabel(p, text="Foco").pack(anchor="w", padx=20, pady=(12, 0))
        self.foco = ctk.StringVar(value="categoria")
        ctk.CTkRadioButton(
            p,
            text="Categoría Predefinida",
            variable=self.foco,
            value="categoria",
            command=self.alternar_foco,
        ).pack(anchor="w", padx=20, pady=(4, 2))
        ctk.CTkRadioButton(
            p,
            text="Tema Específico",
            variable=self.foco,
            value="tema",
            command=self.alternar_foco,
        ).pack(anchor="w", padx=20, pady=2)

        self.cont_foco = ctk.CTkFrame(p, fg_color="transparent")
        self.cont_foco.pack(fill="x", padx=20, pady=6)
        self.cmb_cat = ctk.CTkComboBox(self.cont_foco, values=list(CATEGORIAS), state="readonly")
        self.cmb_cat.set("Herramientas IA")
        self.ent_tema = ctk.CTkEntry(
            self.cont_foco,
            placeholder_text="Escribe tu tema (ej: Docker, Python...)",
        )
        self.cmb_cat.pack(fill="x")

        self.btn = ctk.CTkButton(
            p,
            text="🔥 Buscar Tendencias Virales",
            height=56,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="#e8590c",
            hover_color="#c24a0a",
            command=self.iniciar,
        )
        self.btn.pack(fill="x", padx=20, pady=(24, 20), side="bottom")

    def alternar_foco(self):
        if self.foco.get() == "categoria":
            self.ent_tema.pack_forget()
            self.cmb_cat.pack(fill="x")
        else:
            self.cmb_cat.pack_forget()
            self.ent_tema.pack(fill="x")

    # -- Panel derecho: progreso y resultados -------------------------------
    def _panel_derecho(self):
        p = ctk.CTkFrame(self, fg_color="transparent")
        p.grid(row=0, column=1, sticky="nsew", padx=16, pady=16)
        p.grid_columnconfigure(0, weight=1)
        p.grid_rowconfigure(2, weight=1)

        cab = ctk.CTkFrame(p, fg_color="transparent")
        cab.grid(row=0, column=0, sticky="ew")
        cab.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            cab,
            text="📈 Resultados",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).grid(row=0, column=0, sticky="w")
        ctk.CTkButton(cab, text="Copiar todo", width=110, command=self.copiar).grid(row=0, column=1)

        estado = ctk.CTkFrame(p, fg_color="transparent")
        estado.grid(row=1, column=0, sticky="ew", pady=(10, 8))
        estado.grid_columnconfigure(0, weight=1)
        self.lbl_estado = ctk.CTkLabel(
            estado,
            text="Listo. El periodo usa calendario UTC (a inicios de mes hay menos Shorts).",
            anchor="w",
        )
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
        if self.foco.get() == "categoria":
            cat = self.cmb_cat.get()
            if cat not in CATEGORIAS:
                self.lbl_estado.configure(text="⚠️ Elige una categoría válida.")
                return
            consulta = CATEGORIAS[cat]
            filtrar_tech = True
        else:
            consulta = self.ent_tema.get().strip()
            if not consulta:
                self.lbl_estado.configure(text="⚠️ Escribe un tema específico.")
                return
            filtrar_tech = False

        try:
            cantidad = int(self.cmb_cant.get())
        except ValueError:
            self.lbl_estado.configure(text="⚠️ Cantidad no válida.")
            return

        cfg = dict(
            cantidad=cantidad,
            periodo=self.cmb_periodo.get(),
            idioma=self.cmb_idioma.get(),
            criterio=self.cmb_criterio.get(),
            consulta=consulta,
            filtrar_tech=filtrar_tech,
        )
        self.ocupado = True
        self.btn.configure(state="disabled", text="Buscando...")
        self.resultados.delete("1.0", "end")
        self.barra.set(0)
        self.lbl_estado.configure(text="Analizando tendencias actuales en redes...")
        threading.Thread(target=self.trabajo, args=(cfg,), daemon=True).start()

    def trabajo(self, cfg):
        q = self.cola

        def progreso(frac, texto):
            q.put(("estado", texto, float(frac)))

        try:
            progreso(0.02, "Analizando tendencias actuales en redes...")
            videos = buscar_shorts(
                cfg["consulta"],
                cfg["cantidad"],
                cfg["periodo"],
                cfg["idioma"],
                cfg["criterio"],
                progreso,
                filtrar_tech=cfg.get("filtrar_tech", False),
            )
            if not videos:
                raise ErrorApp(
                    "No se encontraron Shorts para esa búsqueda y periodo. "
                    "Prueba otro tema o un periodo más amplio."
                )
            aviso = ""
            if len(videos) < cfg["cantidad"]:
                aviso = (
                    f"ℹ️ Solo se encontraron {len(videos)} Shorts reales que cumplen los filtros "
                    f"(pediste {cfg['cantidad']}). Prueba un periodo más amplio para obtener más.\n\n"
                )
            q.put(("texto", aviso + formatear(videos, cfg["criterio"])))
            progreso(
                1.0,
                f"✅ Listo: {len(videos)} Shorts ordenados por {cfg['criterio'].lower()} "
                f"({cfg['periodo'].lower()}).",
            )
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
