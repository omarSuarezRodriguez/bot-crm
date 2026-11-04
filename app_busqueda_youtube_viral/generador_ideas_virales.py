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
MAX_DURACION_SEG = 60   # sube a 180 si quieres incluir Shorts de hasta 3 min
MAX_PAGINAS = 10        # cada página de búsqueda cuesta 100 unidades de cuota

PERIODOS = {"Hoy": 1, "Esta semana": 7, "Este mes": 30, "Este año": 365, "Siempre": None}
CRITERIOS = {"Vistas": "vistas", "Likes": "likes", "Comentarios": "comentarios"}
IDIOMAS = {"Español": "es", "Inglés": "en", "Cualquiera": None}

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
    params["key"] = os.getenv("YOUTUBE_API_KEY", "")
    r = requests.get(f"{YT_BASE}/{endpoint}", params=params, timeout=20)
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
    dias = PERIODOS[periodo]
    if dias is None:
        return None
    return (datetime.now(timezone.utc) - timedelta(days=dias)).strftime("%Y-%m-%dT%H:%M:%SZ")


def buscar_shorts(consulta, cantidad, periodo, idioma, criterio, progreso):
    objetivo = min(max(cantidad * 2, 30), 300)  # se pide de más: luego se filtra por duración
    desde = fecha_desde(periodo)
    codigo_idioma = IDIOMAS[idioma]

    # 1) Buscar IDs de videos, paginando
    ids, token, paginas = [], None, 0
    while len(ids) < objetivo and paginas < MAX_PAGINAS:
        p = dict(part="snippet", q=f"{consulta} #shorts", type="video",
                 videoDuration="short", order="viewCount", maxResults=50)
        if desde:
            p["publishedAfter"] = desde
        if codigo_idioma:
            p["relevanceLanguage"] = codigo_idioma
        if token:
            p["pageToken"] = token
        data = yt_get("search", **p)
        for it in data.get("items", []):
            vid = it["id"].get("videoId")
            if vid and vid not in ids:
                ids.append(vid)
        paginas += 1
        token = data.get("nextPageToken")
        progreso(0.05 + 0.60 * min(len(ids) / objetivo, 1),
                 f"Analizando tendencias actuales en redes... ({len(ids)} videos encontrados)")
        if not token:
            break

    # 2) Métricas reales y filtro por duración
    videos = []
    for i in range(0, len(ids), 50):
        data = yt_get("videos", part="snippet,statistics,contentDetails", id=",".join(ids[i:i + 50]))
        for v in data.get("items", []):
            seg = duracion_seg(v["contentDetails"].get("duration"))
            if not (0 < seg <= MAX_DURACION_SEG):
                continue
            st = v.get("statistics", {})
            videos.append({
                "titulo": v["snippet"]["title"],
                "canal": v["snippet"]["channelTitle"],
                "fecha": v["snippet"]["publishedAt"][:10],
                "duracion": seg,
                "vistas": int(st.get("viewCount", 0)),
                "likes": int(st.get("likeCount", 0)),
                "comentarios": int(st.get("commentCount", 0)),
                "url": f"https://youtube.com/shorts/{v['id']}",
            })
        progreso(0.65 + 0.30 * min((i + 50) / max(len(ids), 1), 1), "Midiendo métricas reales de cada video...")

    # 3) Orden por la métrica elegida (desempate: vistas)
    clave = CRITERIOS[criterio]
    videos.sort(key=lambda x: (x[clave], x["vistas"]), reverse=True)
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


def formatear(videos, criterio):
    clave = CRITERIOS[criterio]
    etiquetas = [("vistas", "👁", "vistas"), ("likes", "👍", "likes"), ("comentarios", "💬", "comentarios")]
    salida = []
    for i, v in enumerate(videos, 1):
        metricas = " · ".join(
            f"{'⭐ ' if k == clave else ''}{icono} {fmt(v[k])} {nombre}" for k, icono, nombre in etiquetas)
        salida.append(
            f"#{i} · {metricas}\n"
            f"📌 {v['titulo']}\n"
            f"📺 {v['canal']} · {v['duracion']} s · publicado {v['fecha']}\n"
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

        ctk.CTkLabel(p, text="⚙️ Configuración de Búsqueda",
                     font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=20, pady=(20, 10))

        # Cantidad
        self.lbl_cant = ctk.CTkLabel(p, text="Cantidad de resultados: 20")
        self.lbl_cant.pack(anchor="w", padx=20, pady=(10, 0))
        self.sld_cant = ctk.CTkSlider(
            p, from_=1, to=200, number_of_steps=199,
            command=lambda v: self.lbl_cant.configure(text=f"Cantidad de resultados: {int(v)}"))
        self.sld_cant.set(20)
        self.sld_cant.pack(fill="x", padx=20, pady=(4, 8))

        # Criterio y periodo
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
        self.cmb_idioma.set("Cualquiera")
        self.cmb_idioma.pack(fill="x", padx=20, pady=4)

        # Foco
        ctk.CTkLabel(p, text="Foco").pack(anchor="w", padx=20, pady=(12, 0))
        self.foco = ctk.StringVar(value="categoria")
        ctk.CTkRadioButton(p, text="Categoría Predefinida", variable=self.foco, value="categoria",
                           command=self.alternar_foco).pack(anchor="w", padx=20, pady=(4, 2))
        ctk.CTkRadioButton(p, text="Tema Específico", variable=self.foco, value="tema",
                           command=self.alternar_foco).pack(anchor="w", padx=20, pady=2)

        self.cont_foco = ctk.CTkFrame(p, fg_color="transparent")
        self.cont_foco.pack(fill="x", padx=20, pady=6)
        self.cmb_cat = ctk.CTkComboBox(self.cont_foco, values=list(CATEGORIAS), state="readonly")
        self.cmb_cat.set("Herramientas IA")
        self.ent_tema = ctk.CTkEntry(self.cont_foco, placeholder_text="Escribe tu tema (ej: Docker, Python...)")
        self.cmb_cat.pack(fill="x")

        # Botón principal
        self.btn = ctk.CTkButton(p, text="🔥 Buscar Tendencias Virales", height=56,
                                 font=ctk.CTkFont(size=16, weight="bold"),
                                 fg_color="#e8590c", hover_color="#c24a0a", command=self.iniciar)
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
        ctk.CTkLabel(cab, text="📈 Resultados",
                     font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, sticky="w")
        ctk.CTkButton(cab, text="Copiar todo", width=110, command=self.copiar).grid(row=0, column=1)

        estado = ctk.CTkFrame(p, fg_color="transparent")
        estado.grid(row=1, column=0, sticky="ew", pady=(10, 8))
        estado.grid_columnconfigure(0, weight=1)
        self.lbl_estado = ctk.CTkLabel(estado, text="Listo. Configura tu búsqueda y pulsa el botón.", anchor="w")
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
        if not os.getenv("YOUTUBE_API_KEY"):
            self.lbl_estado.configure(text="⚠️ Falta YOUTUBE_API_KEY en el archivo .env")
            return
        if self.foco.get() == "categoria":
            consulta = CATEGORIAS[self.cmb_cat.get()]
        else:
            consulta = self.ent_tema.get().strip()
            if not consulta:
                self.lbl_estado.configure(text="⚠️ Escribe un tema específico.")
                return
        cfg = dict(cantidad=int(self.sld_cant.get()), periodo=self.cmb_periodo.get(),
                   idioma=self.cmb_idioma.get(), criterio=self.cmb_criterio.get(), consulta=consulta)
        self.ocupado = True
        self.btn.configure(state="disabled", text="Buscando...")
        self.resultados.delete("1.0", "end")
        self.barra.set(0)
        self.lbl_estado.configure(text="Analizando tendencias actuales en redes...")
        threading.Thread(target=self.trabajo, args=(cfg,), daemon=True).start()

    def trabajo(self, cfg):
        q = self.cola

        def progreso(frac, texto):
            q.put(("estado", texto, frac))

        try:
            progreso(0.02, "Analizando tendencias actuales en redes...")
            videos = buscar_shorts(cfg["consulta"], cfg["cantidad"], cfg["periodo"],
                                   cfg["idioma"], cfg["criterio"], progreso)
            if not videos:
                raise ErrorApp("No se encontraron Shorts para esa búsqueda y periodo. "
                               "Prueba otro tema o un periodo más amplio.")
            aviso = ""
            if len(videos) < cfg["cantidad"]:
                aviso = (f"ℹ️ Solo se encontraron {len(videos)} Shorts reales que cumplen los filtros "
                         f"(pediste {cfg['cantidad']}). Prueba un periodo más amplio para obtener más.\n\n")
            q.put(("texto", aviso + formatear(videos, cfg["criterio"])))
            progreso(1.0, f"✅ Listo: {len(videos)} Shorts ordenados por {cfg['criterio'].lower()} "
                          f"({cfg['periodo'].lower()}).")
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
        self.clipboard_clear()
        self.clipboard_append(self.resultados.get("1.0", "end"))
        self.lbl_estado.configure(text="📋 Resultados copiados al portapapeles.")


if __name__ == "__main__":
    cargar_env()
    App().mainloop()
