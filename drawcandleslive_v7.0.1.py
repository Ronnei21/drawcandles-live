import os
import sys
import json
import time
import math
import tkinter as tk
from tkinter import ttk, colorchooser, messagebox, filedialog
from dataclasses import dataclass
from typing import Optional, List, Tuple
from datetime import datetime, timedelta
from supabase_config import SUPABASE_URL, SUPABASE_ANON_KEY
from supabase import create_client
import win32api
import win32con
import win32gui

try:
    import winsound
    HAS_WINSOUND = True
except ImportError:
    HAS_WINSOUND = False

# ============================================================
# Projetor de Candles — v31
# ------------------------------------------------------------
# Novidades da v31:
# - Zoom horizontal com roda do mouse (independente por painel).
# - Grade acompanha o zoom.
# - Limites de zoom: 25% a 500%.
# - Ponto sob o cursor fica fixo ao dar zoom.
# - Tudo da v30 preservado.
# ============================================================

OVERLAY_TRANSPARENT = "#010203"
COLLAPSED_CLOCK_MIN_WIDTH = 180
COLLAPSED_CLOCK_MAX_WIDTH = 360
COLLAPSED_CLOCK_MONITOR_RATIO = 0.24
COLLAPSED_CLOCK_ASPECT = 342 / 2048
ICONE_COLLAPSE_SIZE = 40
DEFAULT_LINE_COLOR = "#e8b93e"

COR_CANDLE_ALTA = "#1B7A2E"
COR_CANDLE_BAIXA = "#D32828"
COR_BORDA_ESCURA = "#FFFFFF"
COR_BORDA_CLARA = "#000000"
COR_TEXTO_ESCURO = "#FFFFFF"
COR_TEXTO_CLARO = "#000000"

# Tema escuro
TEMA_ESCURO = {
    "painel_fundo": "#101419",
    "painel_grade": "#1b2229",
    "painel_label_fonte": "#8fd3ff",
    "painel_label_projetado": "#cfd8e3",
    "painel_label_excluido": "#5b6b7b",
    "borda_candle": COR_BORDA_ESCURA,
}

# Tema claro
TEMA_CLARO = {
    "painel_fundo": "#f0f0f0",
    "painel_grade": "#dcdcdc",
    "painel_label_fonte": "#1a4d80",
    "painel_label_projetado": "#333333",
    "painel_label_excluido": "#7a7a7a",
    "borda_candle": COR_BORDA_CLARA,
}

# Limites de zoom
ZOOM_MIN = 0.25
ZOOM_MAX = 5.0
ZOOM_PASSO = 1.15

ALERT_FILE = os.path.join(os.path.expanduser("~"),
                          ".projetor_candles_alertas.json")

CONFIG_FILE = os.path.join(os.path.expanduser("~"),
                           ".projetor_candles_config.json")

DEFAULT_ALERT_SOUND = r"C:\Windows\Media\Alarm01.wav"

DEFAULT_ALERTS = [
    {"hora": "09:00", "cor": "#ff3b3b", "anticipar_seg": 60,
     "som": DEFAULT_ALERT_SOUND},
]

CLOCK_BG = "#0d1117"
CLOCK_FG_NORMAL = "#e7edf4"
CLOCK_FG_ALERT = "#ff3b3b"
CLOCK_MUTE_BG = "#7a2a2a"
CLOCK_MUTE_FG = "#ffffff"

CLOCK_DEFAULT_W = 190
CLOCK_DEFAULT_H = 64
CLOCK_MIN_W = 110
CLOCK_MIN_H = 44

BTN_BG = "#171a20"
BTN_BG_HOVER = "#30363c"
BTN_BG_ATIVO = "#51472b"
BTN_FG = "#e4e7e9"

BTN_DESTAQUE_BG = "#5c4a2e"
BTN_DESTAQUE_HOVER = "#7a6238"

ICONE_ARQUIVO = "icone.png"

MAX_HISTORICO = 50

ESPESSURAS = [
    ("Fina", 2),
    ("Média", 5),
    ("Grossa", 10),
    ("Extra", 20),
    ("Enorme", 30),
    ("Gigante", 50),
]


def caminho_recurso(nome):
    if getattr(sys, "_MEIPASS", None):
        return os.path.join(sys._MEIPASS, nome)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), nome)


def carregar_config():
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {"tema": "escuro"}


def salvar_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as exc:
        print("Erro ao salvar config:", exc)


try:
    from PIL import ImageGrab, ImageTk, Image, ImageDraw, ImageFilter
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


@dataclass
class Candle:
    open_y: float
    close_y: float
    high_y: float
    low_y: float
    x: float
    width: float = 18.0
    painel: object = None

    @property
    def body_top(self):
        return min(self.open_y, self.close_y)

    @property
    def body_bottom(self):
        return max(self.open_y, self.close_y)

    @property
    def bullish(self):
        return self.close_y < self.open_y


@dataclass
class Fib:
    x0: float
    y0: float
    x1: float
    y1: float
    espessura: int = 2
    color: Optional[str] = None
    painel: object = None
    estilo: str = "padrao"


@dataclass
class HLine:
    y: float
    color: str = DEFAULT_LINE_COLOR
    espessura: int = 2
    painel: object = None


@dataclass
class SegLine:
    x0: float
    y0: float
    x1: float
    y1: float
    color: str
    espessura: int = 2
    painel: object = None


@dataclass
class ArrowShape:
    x0: float
    y0: float
    x1: float
    y1: float
    color: str
    espessura: int = 2
    painel: object = None


@dataclass
class CircleShape:
    x0: float
    y0: float
    x1: float
    y1: float
    color: str
    espessura: int = 2
    painel: object = None


@dataclass
class RectShape:
    x0: float
    y0: float
    x1: float
    y1: float
    color: str
    espessura: int = 2
    painel: object = None

@dataclass
class TextShape:
    x: float
    y: float
    texto: str
    color: str
    espessura: int = 2
    painel: object = None
    tamanho_fonte: int = 14

@dataclass
class HighlightShape:
    x0: float
    y0: float
    x1: float
    y1: float
    color: str = "#f5d90a"
    painel: object = None    


# ============================================================
# Templates de candle (pacote fechado — não configurável)
# ============================================================
# Cada template: (nome, cor_corpo_alta, cor_corpo_baixa, cor_borda)
CANDLE_TEMPLATES = [
    ("Padrão",     "#1B7A2E", "#D32828", "#000000"),
    ("Padrão Branco","#1B7A2E", "#D32828", "#FFFFFF"),
    ("Palmeiras",  "#006437", "#FFFFFF", "#F5C400"),
    ("Grêmio",     "#0D80BF", "#000000", "#FFFFFF"),
    ("Flamengo",   "#D6001C", "#000000", "#FFFFFF"),
    ("São Paulo",  "#FFFFFF", "#D6001C", "#000000"),
    ("Corinthians","#000000", "#FFFFFF", "#000000"),
    ("Fluminense", "#7A1E2C", "#007A3D", "#FFFFFF"),
    ("Internacional","#D6001C", "#FFFFFF", "#D6001C"),
    ("Cruzeiro",   "#0033A0", "#FFFFFF", "#0033A0"),
    ("Brasil",     "#009739", "#FFD700", "#FFFFFF"),
    ("Boca Jr",    "#0033A0", "#FFD700", "#0033A0"),
]

SEG_COLORS = [("#e8b93e", "Dourado"),
              ("#39d98a", "Verde"), ("#ff5c68", "Vermelho"),
              ("#3ec6ff", "Azul"), ("#101010", "Preto"),
              ("#ffffff", "Branco"), ("#f5d90a", "Amarelo"),
              ("#ff5cae", "Pink"), ("#a05cff", "Roxo")]


class Tooltip:
    def __init__(self, widget, text, delay=450):
        self.widget = widget
        self.text = text
        self.delay = delay
        self._after_id = None
        self._tip = None
        widget.bind("<Enter>", self._agendar, add="+")
        widget.bind("<Leave>", self._cancelar, add="+")
        widget.bind("<ButtonPress>", self._cancelar, add="+")

    def _agendar(self, event=None):
        self._cancelar()
        self._after_id = self.widget.after(self.delay, self._mostrar)

    def _cancelar(self, event=None):
        if self._after_id is not None:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        self._esconder()

    def _mostrar(self):
        if self._tip is not None:
            return
        try:
            x = self.widget.winfo_rootx() + self.widget.winfo_width() // 2
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        except Exception:
            return
        self._tip = tk.Toplevel(self.widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{x}+{y}")
        self._tip.attributes("-topmost", True)
        lbl = tk.Label(self._tip, text=self.text,
                       bg="#20272f", fg="#e7edf4",
                       font=("Segoe UI", 9),
                       padx=8, pady=3,
                       relief="solid", bd=1,
                       highlightbackground="#3d6082",
                       highlightthickness=1)
        lbl.pack()

    def _esconder(self):
        if self._tip is not None:
            try:
                self._tip.destroy()
            except Exception:
                pass
            self._tip = None


class SoundPlayer:
    def __init__(self):
        self.tocando = False

    def tocar_loop(self, caminho):
        if not HAS_WINSOUND:
            return False
        try:
            if caminho and os.path.exists(caminho):
                winsound.PlaySound(
                    caminho,
                    winsound.SND_FILENAME | winsound.SND_ASYNC |
                    winsound.SND_LOOP)
                self.tocando = True
                return True
        except Exception:
            pass
        try:
            winsound.MessageBeep(winsound.MB_ICONHAND)
            self.tocando = True
            return True
        except Exception:
            return False

    def parar(self):
        if not HAS_WINSOUND:
            return
        try:
            winsound.PlaySound(None, winsound.SND_PURGE)
        except Exception:
            pass
        self.tocando = False


class Panel:
    def __init__(self, app, frame, timeframe, source=False, overlay=False):
        self.app = app
        self.frame = frame
        self.timeframe = timeframe
        self.source = source
        self.overlay = overlay
        tema = app.tema if hasattr(app, "tema") else TEMA_ESCURO
        bg = OVERLAY_TRANSPARENT if overlay else tema["painel_fundo"]
        self.canvas = tk.Canvas(
            frame, bg=bg, highlightthickness=0 if overlay else 1,
            highlightbackground="#39414a"
        )
        self.canvas.pack(fill="both", expand=True)
        self.candles: List[Candle] = []

        # ===== Estado de zoom/pan (v31) =====
        self.zoom = 1.0
        self.zoom_y = 1.0
        self.offset_x = 0.0
        self.offset_y = 0.0

        # ===== Faixa da escala de preço (lateral direita) =====
        self.largura_escala = 30
        self.escala_canvas = None
        self._escala_drag_origin = None
        self._pan_origin = None
        if not overlay:
            self.escala_canvas = tk.Canvas(
                frame, width=self.largura_escala, bg="#0d1117",
                highlightthickness=0, bd=0, cursor="sb_v_double_arrow")
            self.escala_canvas.place(relx=1.0, rely=0.0, anchor="ne",
                                     relheight=1.0)
            self.escala_canvas.bind("<ButtonPress-1>", self._escala_drag_start)
            self.escala_canvas.bind("<B1-Motion>", self._escala_drag_move)
            self.escala_canvas.bind("<ButtonRelease-1>",
                                    self._escala_drag_end)

    def clear(self):
        self.candles.clear()
        self.redraw()

    # ===== Transformações mundo <-> tela =====
    def mundo_para_tela(self, x, y):
        tx = (x - self.offset_x) * self.zoom
        ty = (y - self.offset_y) * self.zoom_y
        return tx, ty

    def tela_para_mundo(self, x, y):
        mx = (x / self.zoom) + self.offset_x
        my = (y / self.zoom_y) + self.offset_y
        return mx, my

    def redraw(self):
        c = self.canvas
        c.delete("all")
        w = max(c.winfo_width(), 10)
        h = max(c.winfo_height(), 10)

        # Modo Lousa: fundo sólido, sem grade nem label
        if self.app.lousa_estado != 0 and self is self.app.lousa_panel:
            if self.app.lousa_estado == 1:
                c.config(bg="#ffffff")
            else:
                c.config(bg="#000000")
            for candle in self.candles:
                self.draw_candle(candle)                
            for line in self.app.lines:
                self.draw_hline(line)
            for seg in self.app.segments:
                self.draw_segment(seg)
            for arrow in self.app.arrows:
                self.draw_arrow(arrow)
            for circle in self.app.circles:
                self.draw_circle(circle)
            for rect in self.app.rects:
                self.draw_rect(rect)
            for txt in self.app.texts:
                self.draw_text(txt)
            for hl in self.app.highlights:
                self.draw_highlight(hl)
            if self.app.fib:
                self.draw_fib(self.app.fib)
            self.draw_selection()
            return        

        tema = self.app.tema if hasattr(self.app, "tema") else TEMA_ESCURO

        if self.overlay and self.app.overlay_photo is not None:
            c.create_image(0, 0, anchor="nw",
                           image=self.app.overlay_photo, tags="bg")

        if not self.overlay:
            c.config(bg=tema["painel_fundo"])

            # Grade — passo X usa zoom, passo Y usa zoom_y
            grade_passo_x = 50 * self.zoom
            grade_passo_y = 50 * self.zoom_y
            if grade_passo_x > 5:
                x0_tela = -((self.offset_x * self.zoom) % grade_passo_x)
                x = x0_tela
                while x < w:
                    c.create_line(x, 0, x, h, fill=tema["painel_grade"])
                    x += grade_passo_x
            if grade_passo_y > 5:
                y0_tela = -((self.offset_y * self.zoom_y) % grade_passo_y)
                y = y0_tela
                while y < h:
                    c.create_line(0, y, w, y, fill=tema["painel_grade"])
                    y += grade_passo_y

            role = self.app.panel_role(self)
            if role == "source":
                label = f"{self.timeframe} MIN · FONTE"
                color = tema["painel_label_fonte"]
            elif role == "projected":
                label = f"{self.timeframe} MIN · projetado"
                color = tema["painel_label_projetado"]
            elif role == "excluded":
                label = f"{self.timeframe} MIN · fora da contagem"
                color = tema["painel_label_excluido"]
            else:
                label = f"{self.timeframe} MIN"
                color = tema["painel_label_projetado"]
            c.create_text(8, 8, anchor="nw", text=label,
                          fill=color, font=("Segoe UI", 10, "bold"))

        for candle in self.candles:
            self.draw_candle(candle)

        for line in self.app.lines:
            self.draw_hline(line)
        for seg in self.app.segments:
            self.draw_segment(seg)
        for arrow in self.app.arrows:
            self.draw_arrow(arrow)
        for circle in self.app.circles:
            self.draw_circle(circle)
        for rect in self.app.rects:
            self.draw_rect(rect)
        for txt in self.app.texts:
            self.draw_text(txt)
        for hl in self.app.highlights:
            self.draw_highlight(hl)            

        if self.app.fib:
            self.draw_fib(self.app.fib)

        self.draw_selection()

        if not self.overlay and self.app.panel_role(self) == "excluded":
            src = self.app.source_panel()
            if src:
                c.create_text(
                    w / 2, h / 2,
                    text=f"Timeframe menor que a fonte ({src.timeframe} MIN)\n"
                         "Desenho desativado neste painel",
                    fill=tema["painel_label_excluido"],
                    font=("Segoe UI", 11),
                    justify="center"
                )

        # ===== Faixa da escala de preço (lateral direita) =====
        if self.escala_canvas is not None:
            e = self.escala_canvas
            e.delete("all")
            escuro = tema["painel_fundo"] != "#f0f0f0"
            fundo = "#1e2a38" if escuro else "#c8d4e0"
            grade = "#3a4d63" if escuro else "#a0b0c0"
            borda = "#5a7a9a" if escuro else "#8098b0"
            cor_txt = "#e7edf4" if escuro else "#1a2a3a"
            e.config(bg=fundo)
            altura = max(e.winfo_height(), 10)
            passo = 50
            y = 0
            while y < altura:
                e.create_line(0, y, self.largura_escala, y, fill=grade)
                y += passo
            e.create_line(0, 0, 0, altura, fill=borda, width=1)
            texto_zoom = f"{int(self.zoom_y * 100)}%"
            e.create_text(self.largura_escala // 2, 12,
                          text=texto_zoom, fill=cor_txt,
                          font=("Segoe UI", 8, "bold"))

    def _borda_atual(self):
        return self.app.tema["borda_candle"]

    def draw_candle(self, candle):
        # Só desenha se for do painel atual
        if getattr(candle, "painel", None) is not None and \
                getattr(candle, "painel", None) is not self:
            return        
        c = self.canvas

        # Pega as cores do template atual
        try:
            _, cor_alta, cor_baixa, cor_borda = CANDLE_TEMPLATES[
                self.app.candle_template_idx]
        except Exception:
            cor_alta = COR_CANDLE_ALTA
            cor_baixa = COR_CANDLE_BAIXA
            cor_borda = self._borda_atual()

        color = cor_alta if candle.bullish else cor_baixa

        # Na lousa, sobrepõe a borda pra contraste
        if self.app.lousa_estado == 1:
            borda = "#000000"
        elif self.app.lousa_estado == 2:
            borda = "#FFFFFF"
        else:
            borda = cor_borda            

        # Converte coordenadas de mundo pra tela
        x_tela, _ = self.mundo_para_tela(candle.x, 0)
        _, topo_tela = self.mundo_para_tela(0, candle.body_top)
        _, base_tela = self.mundo_para_tela(0, candle.body_bottom)
        _, high_tela = self.mundo_para_tela(0, candle.high_y)
        _, low_tela = self.mundo_para_tela(0, candle.low_y)

        w = candle.width * self.zoom

        c.create_line(x_tela, high_tela, x_tela, low_tela,
                      fill=borda, width=1)
        c.create_rectangle(
            x_tela - w / 2, topo_tela, x_tela + w / 2, base_tela,
            fill=color, outline=borda, width=1
        )

        r = max(2, 4 * self.zoom)
        c.create_oval(x_tela - r, high_tela - r, x_tela + r, high_tela + r,
                      outline=borda, width=1)
        c.create_oval(x_tela - r, low_tela - r, x_tela + r, low_tela + r,
                      outline=borda, width=1)

    def draw_hline(self, line):
        # Só desenha se for do painel atual (ou sem painel definido)
        if getattr(line, "painel", None) is not None and \
                getattr(line, "painel", None) is not self:
            return        
        c = self.canvas
        w = max(c.winfo_width(), 10)
        sel = (self.app.selected
               and self.app.selected[0] == "line"
               and self.app.selected[1] is line)
        esp = getattr(line, "espessura", 2)
        width = max(1, (esp + 1) * self.zoom) if sel else max(1, esp * self.zoom)
        color = getattr(line, "color", DEFAULT_LINE_COLOR) or DEFAULT_LINE_COLOR
        _, y_tela = self.mundo_para_tela(0, line.y)
        c.create_line(0, y_tela, w, y_tela, fill=color, width=int(width))
        if sel:
            c.create_oval(6, y_tela - 5, 16, y_tela + 5,
                          outline="#3ec6ff", width=2)

    def draw_segment(self, seg):
        # Só desenha se for do painel atual
        if getattr(seg, "painel", None) is not None and \
                getattr(seg, "painel", None) is not self:
            return        
        c = self.canvas
        sel = (self.app.selected
               and self.app.selected[0] == "segment"
               and self.app.selected[1] is seg)
        esp = getattr(seg, "espessura", 2)
        width = max(1, (esp + 1) * self.zoom) if sel else max(1, esp * self.zoom)
        x0, y0 = self.mundo_para_tela(seg.x0, seg.y0)
        x1, y1 = self.mundo_para_tela(seg.x1, seg.y1)
        c.create_line(x0, y0, x1, y1, fill=seg.color, width=int(width))
        if sel:
            r = 4
            c.create_oval(x0 - r, y0 - r, x0 + r, y0 + r,
                          outline="#3ec6ff", width=2)
            c.create_oval(x1 - r, y1 - r, x1 + r, y1 + r,
                          outline="#3ec6ff", width=2)

    def draw_arrow(self, arrow):
        # Só desenha se for do painel atual
        if getattr(arrow, "painel", None) is not None and \
                getattr(arrow, "painel", None) is not self:
            return        
        c = self.canvas
        sel = (self.app.selected
               and self.app.selected[0] == "arrow"
               and self.app.selected[1] is arrow)
        esp = getattr(arrow, "espessura", 2)
        width = max(1, (esp + 1) * self.zoom) if sel else max(1, esp * self.zoom)
        x0, y0 = self.mundo_para_tela(arrow.x0, arrow.y0)
        x1, y1 = self.mundo_para_tela(arrow.x1, arrow.y1)
        c.create_line(x0, y0, x1, y1,
                      fill=arrow.color, width=int(width), arrow=tk.LAST,
                      arrowshape=(10, 12, 5))

    def draw_circle(self, circle):
        # Só desenha se for do painel atual
        if getattr(circle, "painel", None) is not None and \
                getattr(circle, "painel", None) is not self:
            return        
        c = self.canvas
        sel = (self.app.selected
               and self.app.selected[0] == "circle"
               and self.app.selected[1] is circle)
        esp = getattr(circle, "espessura", 2)
        width = max(1, (esp + 1) * self.zoom) if sel else max(1, esp * self.zoom)
        x0, y0 = self.mundo_para_tela(circle.x0, circle.y0)
        x1, y1 = self.mundo_para_tela(circle.x1, circle.y1)
        c.create_oval(x0, y0, x1, y1, outline=circle.color, width=int(width))

    def draw_rect(self, rect):
        # Só desenha se for do painel atual
        if getattr(rect, "painel", None) is not None and \
                getattr(rect, "painel", None) is not self:
            return        
        c = self.canvas
        sel = (self.app.selected
               and self.app.selected[0] == "rect"
               and self.app.selected[1] is rect)
        esp = getattr(rect, "espessura", 2)
        width = max(1, (esp + 1) * self.zoom) if sel else max(1, esp * self.zoom)
        x0, y0 = self.mundo_para_tela(rect.x0, rect.y0)
        x1, y1 = self.mundo_para_tela(rect.x1, rect.y1)
        c.create_rectangle(x0, y0, x1, y1, outline=rect.color, width=int(width))

    def draw_highlight(self, hl):
        # Só desenha se for do painel atual
        if getattr(hl, "painel", None) is not None and \
                getattr(hl, "painel", None) is not self:
            return        
        c = self.canvas
        sel = (self.app.selected
               and self.app.selected[0] == "highlight"
               and self.app.selected[1] is hl)
        x0, y0 = self.mundo_para_tela(hl.x0, hl.y0)
        x1, y1 = self.mundo_para_tela(hl.x1, hl.y1)
        c.create_rectangle(x0, y0, x1, y1,
                           fill=hl.color, outline="",
                           stipple="gray50")
        if sel:
            c.create_rectangle(x0, y0, x1, y1,
                               outline="#3ec6ff", width=1.5,
                               dash=(4, 3))

    def draw_text(self, txt):
        # Só desenha se for do painel atual
        if getattr(txt, "painel", None) is not None and \
                getattr(txt, "painel", None) is not self:
            return        
        c = self.canvas
        sel = (self.app.selected
               and self.app.selected[0] == "text"
               and self.app.selected[1] is txt)
        x_tela, y_tela = self.mundo_para_tela(txt.x, txt.y)
        fonte_base = max(9, min(72, int(getattr(txt, "tamanho_fonte", 14))))
        fonte_tam = max(8, int(fonte_base * self.zoom))
        cor = txt.color
        if self.app.lousa_estado == 1 and cor == "#FFFFFF":
            cor = "#000000"
        elif self.app.lousa_estado == 2 and cor == "#000000":
            cor = "#FFFFFF"
        c.create_text(x_tela, y_tela, anchor="nw",
                      text=txt.texto, fill=cor,
                      font=("Segoe UI", fonte_tam, "bold"))
        if sel:
            bbox = c.bbox(c.find_all()[-1])
            if bbox:
                c.create_rectangle(bbox[0] - 3, bbox[1] - 3,
                                   bbox[2] + 3, bbox[3] + 3,
                                   outline="#3ec6ff", width=1.5,
                                   dash=(4, 3))    

    def draw_fib(self, fib):
        # Só desenha se for do painel atual
        if getattr(fib, "painel", None) is not None and \
                getattr(fib, "painel", None) is not self:
            return        
        c = self.canvas
        sel = self.app.selected and self.app.selected[0] == "fib"
        esp = getattr(fib, "espessura", 2)
        width = max(1, (esp + 1) * self.zoom) if sel else max(1, esp * self.zoom)

        x0, y0 = self.mundo_para_tela(fib.x0, fib.y0)
        x1, y1 = self.mundo_para_tela(fib.x1, fib.y1)
        left, right = min(x0, x1), max(x0, x1)
        top, bottom = min(y0, y1), max(y0, y1)

        cor_fib = getattr(fib, "color", None) or self.app.segment_color

        niveis = [
            (0.0, "0%", None, 1.0),
            (50.0, "50%", (7, 5), 1.0),
            (100.0, "100%", None, 1.0),
        ]
        if getattr(fib, "estilo", "padrao") == "completa":
            niveis = [
                (0.0, "0%", None, 1.0),
                (38.2, "38,2%", (4, 4), 0.65),
                (50.0, "50%", (7, 5), 1.0),
                (61.8, "61,8%", (4, 4), 0.65),
                (100.0, "100%", None, 1.0),
            ]
        for nivel, rotulo, tracejado, fator in niveis:
            y = bottom - (bottom - top) * nivel / 100.0
            largura = max(1, int(round(width * fator)))
            opcoes = {"fill": cor_fib, "width": largura}
            if tracejado:
                opcoes["dash"] = tracejado
            c.create_line(left, y, right, y, **opcoes)
            c.create_text(left + 4, y - 4, anchor="sw", text=rotulo,
                          fill=cor_fib, font=("Segoe UI", 9))

        if sel:
            r = 6
            for x, y in ((fib.x0, fib.y0), (fib.x1, fib.y0),
                         (fib.x0, fib.y1), (fib.x1, fib.y1)):
                tx, ty = self.mundo_para_tela(x, y)
                c.create_oval(tx - r, ty - r, tx + r, ty + r,
                              fill="#3ec6ff", outline="white", width=1)

    def draw_selection(self):
        sel = self.app.selected
        if not sel or sel[0] != "candle" or sel[1] is not self:
            return
        cd = sel[2]
        x_tela, _ = self.mundo_para_tela(cd.x, 0)
        _, high_tela = self.mundo_para_tela(0, cd.high_y)
        _, low_tela = self.mundo_para_tela(0, cd.low_y)
        w = cd.width * self.zoom
        self.canvas.create_rectangle(
            x_tela - w / 2 - 5, high_tela - 5,
            x_tela + w / 2 + 5, low_tela + 5,
            outline="#3ec6ff", width=1.5, dash=(4, 3)
        )

    # ===== Escala de preço (zoom vertical) =====
    def _escala_drag_start(self, event):
        """Clique na faixa da escala — começa o drag vertical."""
        self._escala_drag_origin = (event.y_root, self.zoom_y)

    def _escala_drag_move(self, event):
        """Arrasta a escala — aplica zoom vertical."""
        if not self._escala_drag_origin:
            return
        y0, zoom_y0 = self._escala_drag_origin
        dy = event.y_root - y0

        fator = 1.0 + (dy / 200.0)
        if fator <= 0.1:
            fator = 0.1
        novo_zoom_y = zoom_y0 * fator
        novo_zoom_y = max(ZOOM_MIN, min(ZOOM_MAX, novo_zoom_y))

        if abs(novo_zoom_y - self.zoom_y) > 1e-6:
            # Ponto do "mundo" que fica no centro da tela atual
            _, centro_y_tela = 0, self.canvas.winfo_height() / 2
            centro_y_mundo = (centro_y_tela / self.zoom_y) + self.offset_y

            self.zoom_y = novo_zoom_y

            # Recalcula offset pra manter o mesmo ponto do mundo no centro
            self.offset_y = centro_y_mundo - (centro_y_tela / self.zoom_y)
            self.redraw()

    def _escala_drag_end(self, event):
        self._escala_drag_origin = None

    # ===== Zoom =====
    def aplicar_zoom(self, event, direcao):
        """Aplica zoom mantendo o ponto sob o cursor fixo."""
        if self.overlay:
            return False

        if direcao > 0:
            novo_zoom = self.zoom * ZOOM_PASSO
        else:
            novo_zoom = self.zoom / ZOOM_PASSO

        novo_zoom = max(ZOOM_MIN, min(ZOOM_MAX, novo_zoom))
        if abs(novo_zoom - self.zoom) < 1e-6:
            return False

        # Pega coordenadas de mundo do ponto sob o cursor
        mx, my = self.tela_para_mundo(event.x, event.y)

        # Aplica o novo zoom
        self.zoom = novo_zoom

        # Recalcula offset pra manter (mx, my) na mesma posição de tela
        self.offset_x = mx - (event.x / self.zoom)
        self.offset_y = my - (event.y / self.zoom_y)
        return True
        
    # ===== Pan (arrastar canvas) =====
    def pan_iniciar(self, event):
        if self.overlay:
            return
        self._pan_origin = (event.x, event.y,
                            self.offset_x, self.offset_y)

    def pan_mover(self, event):
        if self._pan_origin is None or self.overlay:
            return
        x0, y0, ox, oy = self._pan_origin
        dx = event.x - x0
        dy = event.y - y0
        self.offset_x = ox - (dx / self.zoom)
        self.offset_y = oy - (dy / self.zoom_y)
        self.redraw()

    def pan_terminar(self, event):
        self._pan_origin = None


class CandleProjector:
    

    FERRAMENTAS = [
        ("move",    "➤", "Mover"),
        ("fib",     "≡",  "Fibo"),
        ("line",    "─",  "Linha"),
        ("segment", "╱",  "Traço"),
        ("arrow",   "↗",  "Seta"),
        ("circle",  "○",  "Círculo"),
        ("rect",    "□",  "Quadrado"),
        ("eraser",  "◩",  "Borracha"),
        ("text",    "A",  "Texto"),
        ("highlight", "◉", "Highlight"),
    ]

    def __init__(self, root):
        self.root = root
        self.root.title("DrawCandles Live")
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.geometry("+20+20")
        self.root.configure(bg="#111315")
        self._borda_arredondada_pendente = False
        self.root.bind("<Configure>", self._agendar_borda_arredondada, add="+")
        # A janela se ajusta ao conteúdo (só a toolbar)        

        self._icone_photo_principal = None
        self._icone_pequeno_collapse = None
        self._carregar_icones()

        cfg = carregar_config()
        self.tema_nome = cfg.get("tema", "escuro")
        self.tema = TEMA_ESCURO if self.tema_nome == "escuro" else TEMA_CLARO

        self.mode = "move"
        self.mode_names = {"move": "Mover", "candle": "Candles Encadeados",
                           "candle_free": "Candle Livre",
                           "fib": "Fibonacci", "line": "Linha",
                           "segment": "Traço", "eraser": "Borracha",
                           "arrow": "Seta", "circle": "Círculo",
                           "rect": "Quadrado", "text": "Texto",
                           "highlight": "Highlight"}        
        self.panels = []
        self.active_panel: Optional[Panel] = None

        self.fib: Optional[Fib] = None
        self.lines: List[HLine] = []
        self.segments: List[SegLine] = []
        self.arrows: List[ArrowShape] = []
        self.circles: List[CircleShape] = []
        self.rects: List[RectShape] = []
        self.texts: List[TextShape] = []
        self.highlights: List[HighlightShape] = []
        self.lousa_estado = 0  # 0=off, 1=branco, 2=preto
        self.lousa_frame = None
        self.lousa_panel = None
        self.lousa_workspace_backup = None
        self.lousa_desenho = {
            "candles": [],
            "lines": [],
            "segments": [],
            "arrows": [],
            "circles": [],
            "rects": [],
            "texts": [],
            "highlights": [],
            "fib": None,
            "zoom": 1.0,
            "zoom_y": 1.0,
            "offset_x": 0.0,
            "offset_y": 0.0,            
        }        

        self.segment_color = SEG_COLORS[0][0]
        self.candle_template_idx = 0  # índice do template atual
        self.candle_tamanho = 18       # largura base do candle (padrão)        
        self.espessura = 5
        self.tamanho_fonte_texto = 14
        self.fib_estilo = "padrao"

        self.historico = []

        self.seg_start = None
        self.seg_drag = None
        self.shape_start = None

        self.selected = None
        self.is_erasing = False

        self.drag = None
        self._provisional_candle = None
        self.wick_drag = None
        self.fib_drag = None
        self.line_drag = None
        self.fib_start = None
        self.line_start = None
        self.text_drag = None
        self.highlight_drag = None        

        self.overlay_active = False
        self._mouse_hook = None
        self._keyboard_hook = None
        self._mouse_hook_proc = None
        self._keyboard_hook_proc = None
        self._overlay_mouse_captured = False
        self._overlay_mouse_panel = None
        self._overlay_pending_motion = None
        self._overlay_motion_after = None
        self._overlay_event_queue = []
        self._overlay_event_timer = None
        self._toolbar_rect_cache = None
        self._ui_popup_rects_cache = []
        self._ui_popup_windows = []
        self._overlay_right_captured = False
        self._overlay_ctrlz_key_down = False
        self.overlay_holder = None
        self.overlay_panel: Optional[Panel] = None
        self.overlay_candles: List[Candle] = []
        self.overlay_photo = None
        self._normal_geometry = None
        self._overlay_monitor_rect = None
        self.collapsed = False
        self._collapsed_geometry = None
        self._pending_overlay_monitor_rect = None

        self.icon_window = None
        self.icon_canvas = None
        self._collapsed_clock_photo = None
        self._collapsed_clock_template = None
        self._collapsed_clock_image_id = None
        self._collapsed_clock_after = None
        self._collapsed_clock_size = None
        self._icon_drag_origin = None
        self._icon_dragged = False

        self.print_select_win = None
        self.print_start_pos = None
        self.print_rect_id = None

        self.clock_window = None

        self.tool_buttons = {}
        self._toolbar_icon_photos = {}

        self._shift_pressionado = False
        self.root.bind_all("<KeyPress-Shift_L>", self._shift_on, add="+")
        self.root.bind_all("<KeyRelease-Shift_L>", self._shift_off, add="+")
        self.root.bind_all("<KeyPress-Shift_R>", self._shift_on, add="+")
        self.root.bind_all("<KeyRelease-Shift_R>", self._shift_off, add="+")

        self.build_toolbar()

        self.workspace = tk.Frame(self.root, bg="#0b0e11")
        # Workspace não é mais usado no Live — a janela é só a toolbar
        self._criar_canvas_livre()        

        self.root.bind("<Delete>", lambda e: self.delete_selected())
        self.root.bind("<Escape>", self._on_escape)
        self.root.protocol("WM_DELETE_WINDOW", self._fechar_aplicativo)


        self.root.after(150, self.collapse_app)

    def _fechar_aplicativo(self):
        self.overlay_active = False
        self._remover_hooks_globais()
        self.root.destroy()

    # ============================================================
    # Tema
    # ============================================================

    # ============================================================
    # Histórico
    # ============================================================
    def _snapshot(self):
        import copy

        # Os objetos de desenho guardam uma referência ao painel Tkinter em
        # ``painel``. deepcopy tenta copiar também o Tcl/Tk interno e falha
        # com "cannot pickle '_tkinter.tkapp' object". As classes de desenho
        # são dataclasses com campos escalares; uma cópia rasa separa os
        # valores editáveis e mantém a referência do painel intacta.
        def copiar_objeto(objeto):
            return copy.copy(objeto) if objeto is not None else None

        def copiar_lista(itens):
            return [copiar_objeto(item) for item in itens]

        todos_paineis = list(self.panels)
        if self.lousa_panel is not None:
            todos_paineis.append(self.lousa_panel)
        for _, p in getattr(self, "_overlay_paineis", []):
            todos_paineis.append(p)
        estado = {
            "candles": {id(p): copiar_lista(p.candles) for p in todos_paineis},
            "lines": copiar_lista(self.lines),
            "segments": copiar_lista(self.segments),
            "arrows": copiar_lista(self.arrows),
            "circles": copiar_lista(self.circles),
            "rects": copiar_lista(self.rects),
            "texts": copiar_lista(self.texts),
            "highlights": copiar_lista(self.highlights),
            "fib": copiar_objeto(self.fib),
            "lousa_desenho": {
                **{chave: copiar_lista(self.lousa_desenho.get(chave, []))
                   for chave in ("candles", "lines", "segments", "arrows",
                                 "circles", "rects", "texts", "highlights")},
                "fib": copiar_objeto(self.lousa_desenho.get("fib")),
                **{chave: self.lousa_desenho.get(chave, padrao)
                   for chave, padrao in (("zoom", 1.0), ("zoom_y", 1.0),
                                         ("offset_x", 0.0), ("offset_y", 0.0))},
            },
        }
        self.historico.append(estado)
        if len(self.historico) > MAX_HISTORICO:
            self.historico.pop(0)    

    def _restaurar_snapshot(self, estado):
        todos_paineis = list(self.panels)
        if self.lousa_panel is not None:
            todos_paineis.append(self.lousa_panel)
        for _, p in getattr(self, "_overlay_paineis", []):
            todos_paineis.append(p)
        for p in todos_paineis:
            if id(p) in estado["candles"]:
                p.candles = list(estado["candles"][id(p)])
        self.lines = list(estado["lines"])
        self.segments = list(estado["segments"])
        self.arrows = list(estado["arrows"])
        self.circles = list(estado["circles"])
        self.rects = list(estado["rects"])
        self.texts = list(estado.get("texts", []))
        self.highlights = list(estado.get("highlights", []))
        self.fib = estado["fib"]
        if "lousa_desenho" in estado:
            self.lousa_desenho = estado["lousa_desenho"]
        # O painel ativo é a fonte da verdade. A lista overlay_candles é
        # apenas um espelho usado ao sair da sobreposição.
        if self.overlay_panel is not None:
            self.overlay_candles = self.overlay_panel.candles
        else:
            self.overlay_candles = list(estado.get("overlay_candles", []))
        self.selected = None
        if self.lousa_estado == 0:
            self.recalculate()
        self.redraw_all()            

    # ============================================================
    # Atalho Shift
    # ============================================================
    def _shift_on(self, event=None):
        self._shift_pressionado = True

    def _shift_off(self, event=None):
        self._shift_pressionado = False

    def _atalho_candle_traco(self):
        if self.mode == "candle":
            self.set_mode("segment")
            self.hint("Atalho → Traço")
        elif self.mode == "segment":
            self.set_mode("candle")
            self.hint("Atalho → Candle")

    # ============================================================
    # Ícones
    # ============================================================
    def _carregar_icones(self):
        if not HAS_PIL:
            return
        caminho = caminho_recurso(ICONE_ARQUIVO)
        if not os.path.exists(caminho):
            print(f"[aviso] Ícone não encontrado: {caminho}")
            return
        try:
            img = Image.open(caminho).convert("RGBA")
            img_janela = img.resize((48, 48), Image.LANCZOS)
            self._icone_photo_principal = ImageTk.PhotoImage(img_janela)
            try:
                self.root.iconphoto(True, self._icone_photo_principal)
            except Exception:
                pass
            img_peq = img.resize((ICONE_COLLAPSE_SIZE, ICONE_COLLAPSE_SIZE),
                                 Image.LANCZOS)
            self._icone_pequeno_collapse = ImageTk.PhotoImage(img_peq)
        except Exception as exc:
            print(f"[aviso] Falha ao carregar ícone: {exc}")

    # ============================================================
    # Toolbar
    # ============================================================
    def _barra_arrastar_iniciar(self, event):
        """Começa o arraste da janela pela toolbar."""
        self._barra_drag_origin = (event.x_root, event.y_root,
                                    self.root.winfo_x(),
                                    self.root.winfo_y())

    def _barra_arrastar_mover(self, event):
        """Move a janela junto com o mouse."""
        if not hasattr(self, "_barra_drag_origin") or \
                self._barra_drag_origin is None:
            return
        x0, y0, wx, wy = self._barra_drag_origin
        dx = event.x_root - x0
        dy = event.y_root - y0
        self.root.geometry(f"+{wx + dx}+{wy + dy}")

    def _barra_arrastar_terminar(self, event):
        self._barra_drag_origin = None

    def _agendar_borda_arredondada(self, event=None):
        if event is not None and event.widget is not self.root:
            return
        if self._borda_arredondada_pendente:
            return
        self._borda_arredondada_pendente = True
        try:
            self.root.after_idle(self._aplicar_borda_arredondada)
        except tk.TclError:
            self._borda_arredondada_pendente = False

    def _aplicar_borda_arredondada(self):
        """Recorta a janela da barra em cantos arredondados no Windows."""
        self._borda_arredondada_pendente = False
        try:
            import ctypes
            from ctypes import wintypes
            largura = self.root.winfo_width()
            altura = self.root.winfo_height()
            if largura <= 1 or altura <= 1:
                return
            gdi32 = ctypes.windll.gdi32
            user32 = ctypes.windll.user32
            gdi32.CreateRoundRectRgn.argtypes = [
                ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                ctypes.c_int, ctypes.c_int]
            gdi32.CreateRoundRectRgn.restype = ctypes.c_void_p
            user32.SetWindowRgn.argtypes = [
                wintypes.HWND, ctypes.c_void_p, wintypes.BOOL]
            user32.SetWindowRgn.restype = ctypes.c_int
            regiao = gdi32.CreateRoundRectRgn(
                0, 0, largura + 1, altura + 1,
                min(72, altura), min(72, altura))
            if regiao:
                sucesso = user32.SetWindowRgn(
                    wintypes.HWND(self.root.winfo_id()), regiao, True)
                if not sucesso:
                    gdi32.DeleteObject(regiao)
        except Exception:
            pass

    def _toolbar_x_extended(self, x_ref):
        """Mapeia a arte original para a versão com espaços extras."""
        return x_ref + (120 if x_ref >= 318 else 0) + (120 if x_ref >= 1050 else 0)

    def _toolbar_xy(self, x_ref, y_ref, slot=None):
        """Converte coordenadas da arte para a barra estendida."""
        if slot == "candle_size":
            x_ref = 318 + 60
        elif slot == "candle_size_dropdown":
            x_ref = 428
        elif slot == "candle_dropdown":
            x_ref = 302
        elif slot == "text":
            x_ref = 1050 + 120 + 35
        elif slot == "text_dropdown":
            x_ref = 1050 + 120 + 95
        else:
            x_ref = self._toolbar_x_extended(x_ref)
        return (x_ref * self._toolbar_w / self._toolbar_ref_w,
                y_ref * self._toolbar_h / 144)

    def _toolbar_popup_xy(self, x_ref, slot=None):
        self.root.update_idletasks()
        x, _ = self._toolbar_xy(x_ref, 72, slot=slot)
        return (int(self.toolbar.winfo_rootx() + x),
                int(self.toolbar.winfo_rooty() +
                    self.toolbar.winfo_height() + 2))

    def _desenhar_miniatura_candle(self, canvas, template_idx,
                                    tamanho=24, offset_x=0, offset_y=0):
        """Desenha um par de candles nos menus de seleção existentes."""
        _, cor_alta, cor_baixa, cor_borda = CANDLE_TEMPLATES[template_idx]
        w = tamanho // 3
        x1 = offset_x + w // 2 + 2
        y1_topo = offset_y + 4
        y1_base = offset_y + tamanho - 4
        canvas.create_line(x1, y1_topo, x1, y1_base,
                           fill=cor_borda, width=1)
        canvas.create_rectangle(
            x1 - w // 2, y1_topo + 4, x1 + w // 2, y1_base - 4,
            fill=cor_alta, outline=cor_borda, width=2)
        x2 = offset_x + w + w + w // 2 + 2
        y2_topo = offset_y + 7
        y2_base = offset_y + tamanho - 7
        canvas.create_line(x2, y2_topo, x2, y2_base,
                           fill=cor_borda, width=1)
        canvas.create_rectangle(
            x2 - w // 2, y2_topo + 4, x2 + w // 2, y2_base - 4,
            fill=cor_baixa, outline=cor_borda, width=2)

    def _icone_barra(self, nome):
        """Carrega um ícone simples como fallback para a barra."""
        if nome in self._toolbar_icon_photos:
            return self._toolbar_icon_photos[nome]
        try:
            caminho = caminho_recurso(
                os.path.join("assets", "toolbar_icons", f"{nome}.png"))
            imagem = tk.PhotoImage(file=caminho)
            self._toolbar_icon_photos[nome] = imagem
            return imagem
        except (tk.TclError, OSError):
            return None

    def _toolbar_icon(self, nome, tamanho_ref=54):
        tamanho = max(18, round(tamanho_ref * self._toolbar_w / self._toolbar_ref_w))
        chave = (nome, tamanho)
        if chave in self._toolbar_scaled_photos:
            return self._toolbar_scaled_photos[chave]
        try:
            caminho = caminho_recurso(
                os.path.join("assets", "toolbar_icons", f"{nome}.png"))
            imagem = Image.open(caminho).convert("RGBA")
            imagem = imagem.resize((tamanho, tamanho), Image.LANCZOS)
            foto = ImageTk.PhotoImage(imagem)
            self._toolbar_scaled_photos[chave] = foto
            return foto
        except Exception:
            return self._icone_barra(nome)

    def _atualizar_botao_candle(self):
        if not hasattr(self, "toolbar") or not self.toolbar.winfo_exists():
            return
        c = self.toolbar
        c.delete("toolbar_extra_icons")
        x, y = self._toolbar_xy(0, 72, slot="candle_size")
        foto = self._toolbar_icon("candle_free", 60)
        if foto:
            c.create_image(x, y, image=foto, tags="toolbar_extra_icons")
        # O próprio botão de tamanho também seleciona Candles Livres;
        # a seta ao lado continua abrindo o seletor de tamanho.
        escala = self._toolbar_w / self._toolbar_ref_w
        x, y = self._toolbar_xy(0, 72, slot="candle_size_dropdown")
        c.create_polygon(x-10*escala, y-7*escala,
                         x+10*escala, y-7*escala,
                         x, y+7*escala, fill="#c9cdd0",
                         outline="#7b8288", tags="toolbar_extra_icons")
        x, y = self._toolbar_xy(0, 72, slot="text")
        foto = self._toolbar_icon("text", 54)
        if foto:
            c.create_image(x, y, image=foto, tags="toolbar_extra_icons")
        for slot in ("text_dropdown",):
            x, y = self._toolbar_xy(0, 72, slot=slot)
            c.create_polygon(x-4, y-2, x+4, y-2, x, y+3,
                             fill="#c9cdd0", outline="#7b8288",
                             tags="toolbar_extra_icons")
        x, y = self._toolbar_xy(477, 72)
        c.create_polygon(x-4, y-2, x+4, y-2, x, y+3,
                         fill="#c9cdd0", outline="#7b8288",
                         tags="toolbar_extra_icons")

    def _atualizar_icone_lousa(self):
        if not hasattr(self, "toolbar") or not self.toolbar.winfo_exists():
            return
        c = self.toolbar
        c.delete("lousa_dynamic")
        # O X existente nesta posição vira o botão de lousa, como no app atual.
        x1, y1 = self._toolbar_xy(1684, 37)
        x2, y2 = self._toolbar_xy(1742, 108)
        c.create_rectangle(x1, y1, x2, y2, fill="#14171d", outline="",
                           tags="lousa_dynamic")
        nome = "board_black" if self.lousa_estado == 2 else "board_white"
        foto = self._toolbar_icon(nome, 48)
        if foto:
            x, y = self._toolbar_xy(1713, 72)
            c.create_image(x, y, image=foto, tags="lousa_dynamic")

    def _toolbar_round_rect(self, x1, y1, x2, y2, r, **kwargs):
        pontos = [x1+r, y1, x2-r, y1, x2, y1, x2, y1+r,
                  x2, y2-r, x2, y2, x2-r, y2, x1+r, y2,
                  x1, y2, x1, y2-r, x1, y1+r, x1, y1]
        return self.toolbar.create_polygon(
            pontos, smooth=True, splinesteps=12, **kwargs)

    def _make_icon_button(self, parent, emoji, command, tooltip_text,
                          color_bg=BTN_BG, hover_bg=BTN_BG_HOVER,
                          font_size=12, icon_key=None):
        # Mantido como compatibilidade com os popups internos; a barra principal
        # usa agora regiões clicáveis sobre a arte metálica de referência.
        return tk.Button(parent, text=emoji, command=command, bg=color_bg,
                         fg=BTN_FG, relief="flat", bd=0, cursor="hand2")

    def _atualizar_label_modo(self, texto=None):
        try:
            if not hasattr(self, "toolbar") or not self.toolbar.winfo_exists():
                return
            if texto is None:
                texto = self.mode_names[self.mode]
            elif texto.upper().startswith("MODO: "):
                texto = texto[6:]
            elif texto.upper().startswith("MODO:"):
                texto = texto[5:].strip()
            if texto.strip().upper() == "MOVER":
                self.toolbar.delete("mode_dynamic")
                return
            c = self.toolbar
            c.delete("mode_dynamic")
            x1, y1 = self._toolbar_xy(1765, 38)
            x2, y2 = self._toolbar_xy(1934, 108)
            c.create_rectangle(x1, y1, x2, y2, fill="#c8cdd1",
                               outline="#e5e6e7", width=1,
                               tags="mode_dynamic")
            # Sem o prefixo "MODO:", sobra espaço para nomes como
            # "Desenhar Candle". A fonte menor evita quebra ou corte.
            escala = self._toolbar_w / self._toolbar_ref_w
            x, y = self._toolbar_xy(1778, 73)
            largura_texto = (1934 - 1778) * escala
            c.create_text(x, y, text=texto.upper(), fill="#171a1d",
                          anchor="w", width=largura_texto,
                          font=("Segoe UI", max(8, round(15*escala)), "bold"),
                          tags="mode_dynamic")
        except (tk.TclError, AttributeError):
            pass

    def _atualizar_destaque(self):
        if not hasattr(self, "toolbar") or not self.toolbar.winfo_exists():
            return
        c = self.toolbar
        c.delete("tool_active")
        posicoes = {
            "candle": (246, "candle"), "move": (370, "move"),
            "candle_free": (378, "candle_free"),
            "fib": (448, "fib"),
            "line": (520, "line"), "segment": (576, None),
            "arrow": (656, "arrow"), "circle": (739, "circle"),
            "rect": (822, "rect"), "eraser": (906, "eraser"),
            "text": (0, "text"), "highlight": (1024, "highlight"),
        }
        if self.mode not in posicoes:
            self._atualizar_label_modo()
            return
        x_ref, nome_icone = posicoes[self.mode]
        slot = ("text" if self.mode == "text" else
                "candle_size" if self.mode == "candle_free" else None)
        x, y = self._toolbar_xy(x_ref, 72, slot=slot)
        escala = self._toolbar_w / self._toolbar_ref_w
        # O realce encadeado centraliza no ícone fixo e termina antes da seta.
        meia_largura = 38
        self._toolbar_round_rect(
            x - meia_largura*escala, 35*escala,
            x + meia_largura*escala, 110*escala,
            7*escala,
            fill="#55472b",
            outline="#c8aa62",
            width=max(1, round(1.4*escala)), tags="tool_active")
        if nome_icone:
            tamanho_icone = 60 if self.mode == "candle_free" else 54
            foto = self._toolbar_icon(nome_icone, tamanho_icone)
            if foto:
                c.create_image(x, y, image=foto, tags="tool_active")
        else:
            x1, y1 = self._toolbar_xy(x_ref - 15, 86)
            x2, y2 = self._toolbar_xy(x_ref + 15, 58)
            c.create_line(x1, y1, x2, y2, fill="#f1e3be",
                          width=max(2, round(2.5*escala)),
                          capstyle="round", tags="tool_active")
        self._atualizar_label_modo()

    def _atualizar_botao_cores(self):
        if not hasattr(self, "toolbar") or not self.toolbar.winfo_exists():
            return
        self.toolbar.delete("color_indicator")
        x1, y1 = self._toolbar_xy(1275, 106)
        x2, y2 = self._toolbar_xy(1317, 109)
        self.toolbar.create_line(x1, y1, x2, y2, fill=self.segment_color,
                                 width=max(2, round(3*self._toolbar_w/self._toolbar_ref_w)),
                                 tags="color_indicator")

    def _atualizar_botao_espessura(self):
        # Mantém a amostra metálica da referência; o valor segue disponível no menu.
        return

    def build_toolbar(self):
        self.root.bind("<Control-z>", lambda e: self.undo())
        self.root.bind("<Control-Z>", lambda e: self.undo())

        caminho = caminho_recurso(
            os.path.join("assets", "toolbar_reference_extended.png"))
        self._toolbar_ref_w = 2288
        self._toolbar_w = max(1100, min(1900, self.root.winfo_screenwidth()-40))
        self._toolbar_h = round(self._toolbar_w * 144 / self._toolbar_ref_w)
        fundo = Image.open(caminho).convert("RGB")
        fundo = fundo.resize((self._toolbar_w, self._toolbar_h), Image.LANCZOS)
        self._toolbar_background_photo = ImageTk.PhotoImage(fundo)
        self._toolbar_scaled_photos = {}

        bar = tk.Canvas(self.root, width=self._toolbar_w,
                        height=self._toolbar_h, bg="#111319",
                        highlightthickness=0, bd=0, cursor="arrow")
        bar.pack()
        self.toolbar = bar
        bar.create_image(0, 0, image=self._toolbar_background_photo,
                         anchor="nw", tags="background")

        # Referências compatíveis usadas pelos menus já existentes.
        self.btn_candle_canvas = self.toolbar
        self.btn_candle_dropdown = self.toolbar
        self.btn_tamanho_canvas = self.toolbar
        self.btn_cores = self.toolbar
        self.btn_espessura_canvas = self.toolbar
        self.btn_espessura_dropdown = self.toolbar
        self.lousa_btn = self.toolbar
        self.mode_box = self.toolbar
        self.mode_caption = self.toolbar
        self.mode_label = self.toolbar
        self._toolbar_dragging = False
        self._toolbar_regions = [
            (126, 215, self.collapse_app),
            (218, 284, lambda: self.set_mode("candle")),
            (348, 410, lambda: self.set_mode("move")),
            (412, 462, lambda: self.set_mode("fib")),
            (462, 482, self._abrir_dropdown_fibonacci),
            (484, 548, lambda: self.set_mode("line")),
            (548, 615, lambda: self.set_mode("segment")),
            (615, 694, lambda: self.set_mode("arrow")),
            (696, 777, lambda: self.set_mode("circle")),
            (778, 859, lambda: self.set_mode("rect")),
            (860, 958, lambda: self.set_mode("eraser")),
            (958, 1050, lambda: self.set_mode("highlight")),
            (1052, 1138, self.undo),
            (1145, 1237, self.set_mode_clear),
            (1245, 1340, self._abrir_paleta_cores),
            (1342, 1426, self._abrir_paleta_espessura),
            (1428, 1508, self.take_screenshot),
            (1508, 1591, self.take_screenshot_app),
            (1592, 1677, self.toggle_overlay),
            (1680, 1745, self.ciclar_lousa),
            (1938, 2048, self._fechar_aplicativo),
        ]
        self.tool_buttons = {m: p for m, p in {
            "candle": (246, "candle"), "move": (370, "move"),
            "candle_free": (378, "candle_free"),
            "fib": (448, "fib"),
            "line": (520, "line"), "segment": (576, "line"),
            "arrow": (656, "arrow"), "circle": (739, "circle"),
            "rect": (822, "rect"), "eraser": (906, "eraser"),
            "text": (0, "text"), "highlight": (1024, "highlight"),
        }.items()}
        bar.bind("<ButtonPress-1>", self._toolbar_press)
        bar.bind("<B1-Motion>", self._toolbar_motion)
        bar.bind("<ButtonRelease-1>", self._toolbar_release)
        bar.bind("<Motion>", self._toolbar_hover)
        self._atualizar_botao_candle()
        self._atualizar_icone_lousa()
        self._atualizar_botao_cores()
        self._atualizar_destaque()

    def _toolbar_press(self, event):
        x_ext = event.x * self._toolbar_ref_w / self._toolbar_w
        y_ref = event.y * 144 / self._toolbar_h
        if 25 <= y_ref <= 120 and 418 <= x_ext < 438:
            self._abrir_paleta_tamanho_candle()
            return
        if 25 <= y_ref <= 120 and 292 <= x_ext < 318:
            self._abrir_dropdown_candle()
            return
        if 25 <= y_ref <= 120 and 340 <= x_ext < 410:
            self.set_mode("candle_free")
            return
        if 25 <= y_ref <= 120 and 1240 <= x_ext < 1290:
            self._abrir_dropdown_fonte_texto()
            return
        if 25 <= y_ref <= 120 and 1170 <= x_ext < 1240:
            self.set_mode("text")
            return
        if x_ext >= 1290:
            x_ref = x_ext - 240
        elif x_ext >= 438:
            x_ref = x_ext - 120
        else:
            x_ref = x_ext
        if 25 <= y_ref <= 120 and 25 <= x_ref <= 120:
            self._toolbar_dragging = True
            self._barra_arrastar_iniciar(event)
            return
        for x1, x2, acao in self._toolbar_regions:
            if x1 <= x_ref < x2 and 25 <= y_ref <= 120:
                acao()
                return

    def _toolbar_motion(self, event):
        if self._toolbar_dragging:
            self._barra_arrastar_mover(event)

    def _toolbar_release(self, event):
        if self._toolbar_dragging:
            self._toolbar_dragging = False
            self._barra_arrastar_terminar(event)

    def _toolbar_hover(self, event):
        x_ext = event.x * self._toolbar_ref_w / self._toolbar_w
        y_ref = event.y * 144 / self._toolbar_h
        if 25 <= y_ref <= 120 and 292 <= x_ext < 318:
            self.toolbar.config(cursor="hand2")
            return
        if 25 <= y_ref <= 120 and 318 <= x_ext < 438:
            self.toolbar.config(cursor="hand2")
            return
        if 25 <= y_ref <= 120 and 1170 <= x_ext < 1290:
            self.toolbar.config(cursor="hand2")
            return
        if x_ext >= 1290:
            x_ref = x_ext - 240
        elif x_ext >= 438:
            x_ref = x_ext - 120
        else:
            x_ref = x_ext
        if 25 <= y_ref <= 120 and 25 <= x_ref <= 120:
            self.toolbar.config(cursor="fleur")
        elif any(x1 <= x_ref < x2 for x1, x2, _ in self._toolbar_regions):
            self.toolbar.config(cursor="hand2")
        else:
            self.toolbar.config(cursor="arrow")

    # ============================================================
    # Menus
    # ============================================================

    def _abrir_dropdown_fibonacci(self):
        pop = tk.Toplevel(self.root)
        self._ui_popup_windows.append(pop)
        pop.overrideredirect(True)
        pop.attributes("-topmost", True)
        pop.configure(bg="#0b0e11")
        try:
            x, y = self._toolbar_popup_xy(477)
            pop.geometry(f"+{x}+{y}")
        except Exception:
            pop.geometry("+300+100")

        frame = tk.Frame(pop, bg="#0b0e11", highlightthickness=1,
                         highlightbackground="#3d6082")
        frame.pack(padx=1, pady=1)
        tk.Label(frame, text="Modelo Fibonacci", bg="#0b0e11",
                 fg="#9eabb7", font=("Segoe UI", 9, "bold")) \
            .pack(fill="x", padx=8, pady=(6, 3))

        opcoes = [
            ("Padrão  ·  0 / 50 / 100", "padrao"),
            ("Completa  ·  0 / 38,2 / 50 / 61,8 / 100", "completa"),
        ]
        for rotulo, estilo in opcoes:
            ativo = estilo == self.fib_estilo
            botao = tk.Button(
                frame, text=rotulo, anchor="w", cursor="hand2",
                bg="#2f4a63" if ativo else "#161c23", fg="#e7edf4",
                activebackground="#385a78", activeforeground="#ffffff",
                relief="flat", bd=0, padx=8, pady=6,
                command=lambda e=estilo, p=pop: self._escolher_estilo_fibonacci(e, p))
            botao.pack(fill="x", padx=3, pady=1)
        pop.bind("<FocusOut>", lambda e: pop.destroy())
        pop.focus_set()

    def _escolher_estilo_fibonacci(self, estilo, popup):
        if self.selected and self.selected[0] == "fib" and self.fib is not None:
            self._snapshot()
            self.fib.estilo = estilo
        self.fib_estilo = estilo
        try:
            popup.destroy()
        except tk.TclError:
            pass
        self._atualizar_apos_opcao_ferramenta()
        nome = "completa" if estilo == "completa" else "padrão"
        self.hint(f"Fibonacci {nome} selecionada")

    def _abrir_dropdown_fonte_texto(self):
        pop = tk.Toplevel(self.root)
        self._ui_popup_windows.append(pop)
        pop.overrideredirect(True)
        pop.attributes("-topmost", True)
        pop.configure(bg="#0b0e11")
        try:
            x, y = self._toolbar_popup_xy(0, slot="text_dropdown")
            pop.geometry(f"+{x}+{y}")
        except Exception:
            pop.geometry("+300+100")

        frame = tk.Frame(pop, bg="#0b0e11", highlightthickness=1,
                         highlightbackground="#3d6082")
        frame.pack(padx=1, pady=1)
        tk.Label(frame, text="Tamanho da fonte", bg="#0b0e11",
                 fg="#9eabb7", font=("Segoe UI", 9, "bold")) \
            .grid(row=0, column=0, columnspan=8, sticky="w", padx=7, pady=5)
        tamanhos = range(9, 73)
        for indice, tamanho in enumerate(tamanhos):
            linha, coluna = divmod(indice, 8)
            ativo = tamanho == self.tamanho_fonte_texto
            botao = tk.Button(
                frame, text=str(tamanho), cursor="hand2", width=3,
                bg="#2f4a63" if ativo else "#161c23", fg="#e7edf4",
                activebackground="#385a78", activeforeground="#ffffff",
                relief="flat", bd=0, font=("Segoe UI", 9),
                command=lambda t=tamanho, p=pop: self._escolher_tamanho_fonte(t, p))
            botao.grid(row=linha+1, column=coluna, padx=2, pady=2)
        pop.bind("<FocusOut>", lambda e: pop.destroy())
        pop.focus_set()

    def _escolher_tamanho_fonte(self, tamanho, popup):
        tamanho = max(9, min(72, int(tamanho)))
        if self.selected and self.selected[0] == "text":
            self._snapshot()
            self.selected[1].tamanho_fonte = tamanho
        self.tamanho_fonte_texto = tamanho
        try:
            popup.destroy()
        except tk.TclError:
            pass
        self._atualizar_apos_opcao_ferramenta()
        self.hint(f"Tamanho da fonte: {tamanho}")

    def _abrir_dropdown_candle(self):
        """Abre dropdown com os templates de candle."""
        pop = tk.Toplevel(self.root)
        self._ui_popup_windows.append(pop)
        pop.overrideredirect(True)
        pop.attributes("-topmost", True)
        pop.configure(bg="#0b0e11")

        try:
            x, y = self._toolbar_popup_xy(0, slot="candle_dropdown")
            pop.geometry(f"+{x}+{y}")
        except Exception:
            pop.geometry("+300+100")

        frame = tk.Frame(pop, bg="#0b0e11",
                         highlightthickness=1,
                         highlightbackground="#3d6082")
        frame.pack(padx=1, pady=1)

        # Grid 3 colunas
        for i, (nome, _, _, _) in enumerate(CANDLE_TEMPLATES):
            r, c = divmod(i, 3)
            item = tk.Canvas(
                frame, width=60, height=50,
                bg="#161c23", highlightthickness=0, bd=0,
                cursor="hand2")
            item.grid(row=r, column=c, padx=2, pady=2)

            # Miniaturas
            self._desenhar_miniatura_candle(
                item, i, tamanho=44, offset_x=4, offset_y=3)

            item.bind("<ButtonPress-1>",
                      lambda e, idx=i, p=pop:
                      self._escolher_candle_template(idx, p))

        pop.bind("<FocusOut>", lambda e: pop.destroy())
        pop.focus_set()
    
    def _abrir_paleta_cores(self):
        pop = tk.Toplevel(self.root)
        self._ui_popup_windows.append(pop)
        pop.overrideredirect(True)
        pop.attributes("-topmost", True)
        pop.configure(bg="#0b0e11")

        try:
            x, y = self._toolbar_popup_xy(1294)
            pop.geometry(f"+{x}+{y}")
        except Exception:
            pop.geometry("+200+100")

        frame = tk.Frame(pop, bg="#0b0e11",
                         highlightthickness=1,
                         highlightbackground="#3d6082")
        frame.pack(padx=1, pady=1)

        for i, (col, nome) in enumerate(SEG_COLORS):
            r, c = divmod(i, 4)
            b = tk.Button(
                frame, text="  ", bg=col, activebackground=col,
                relief="flat", bd=1, width=3, height=1,
                cursor="hand2",
                command=lambda cc=col: self._escolher_cor(cc, pop))
            b.grid(row=r, column=c, padx=2, pady=2)

        pop.bind("<FocusOut>", lambda e: pop.destroy())
        pop.focus_set()

    def _escolher_candle_template(self, idx, popup):
        """Muda o template do candle e atualiza o botão + redesenha tudo."""
        self.candle_template_idx = idx
        self._atualizar_botao_candle()
        try:
            popup.destroy()
        except Exception:
            pass
        # O template é consultado durante o desenho, então redesenhe já sem
        # precisar trocar de ferramenta para atualizar os candles existentes.
        self._atualizar_apos_opcao_ferramenta()
        self.hint(f"Candle: {CANDLE_TEMPLATES[idx][0]}")

    def _forcar_redraw_global(self):
        """Redesenha TUDO: painéis normais, lousa e TODOS os painéis do Sobrepor.
        Também força update visual imediato."""
        try:
            # Redesenha painéis normais
            for p in self.panels:
                p.redraw()
            # Redesenha lousa (se ativa)
            if self.lousa_panel is not None:
                self.lousa_panel.redraw()
            # Redesenha TODOS os painéis do Sobrepor
            for _, p in getattr(self, "_overlay_paineis", []):
                p.redraw()
            # Força o Tkinter a atualizar visualmente AGORA
            self.root.update_idletasks()
        except Exception as e:
            print(f"Erro no redraw forçado: {e}")

    def _atualizar_apos_opcao_ferramenta(self):
        """Atualiza o overlay sem trocar a ferramenta que está selecionada."""
        if self.overlay_active:
            self._apply_overlay_mode()
        self._forcar_redraw_global()
    
    
    
    def _escolher_cor(self, col, popup):
        self.segment_color = col
        self._atualizar_botao_cores()
        try:
            popup.destroy()
        except Exception:
            pass
        self._atualizar_apos_opcao_ferramenta()
        self.hint("Cor definida para os próximos desenhos")

    def _abrir_paleta_tamanho_candle(self):
        """Abre popup com os tamanhos de candle."""
        pop = tk.Toplevel(self.root)
        self._ui_popup_windows.append(pop)
        pop.overrideredirect(True)
        pop.attributes("-topmost", True)
        pop.configure(bg="#0b0e11")

        try:
            x, y = self._toolbar_popup_xy(0, slot="candle_size_dropdown")
            pop.geometry(f"+{x}+{y}")
        except Exception:
            pop.geometry("+300+100")

        frame = tk.Frame(pop, bg="#0b0e11",
                         highlightthickness=1,
                         highlightbackground="#3d6082")
        frame.pack(padx=1, pady=1)

        # 5 tamanhos
        tamanhos = [
            ("Fino", 12, 2),
            ("Padrão", 18, 5),
            ("Médio", 24, 10),
            ("Grande", 32, 20),
            ("Gigante", 60, 30),
        ]

        for nome, largura, espessura_traco in tamanhos:
            ativo = (largura == self.candle_tamanho)
            row = tk.Frame(frame, bg="#161c23" if not ativo else "#2f4a63",
                           cursor="hand2")
            row.pack(fill="x", padx=2, pady=1)

            preview = tk.Canvas(row, width=70, height=30,
                                bg=row["bg"], highlightthickness=0, bd=0)
            preview.pack(side="left", padx=(6, 8), pady=4)
            # Desenha um candle de exemplo no preview
            x_c = 35
            w_preview = max(4, int(largura / 2))
            preview.create_line(x_c, 3, x_c, 27,
                                fill="#e7edf4", width=1)
            preview.create_rectangle(
                x_c - w_preview // 2, 8,
                x_c + w_preview // 2, 22,
                fill="#1B7A2E", outline="#e7edf4", width=1)

            lbl = tk.Label(row, text=f"{nome} ({largura}px)",
                           bg=row["bg"], fg="#e7edf4",
                           font=("Segoe UI", 10), anchor="w")
            lbl.pack(side="left", fill="x", expand=True, padx=(0, 12))

            def _escolher(l=largura, e=espessura_traco, p=pop):
                self.candle_tamanho = l
                try:
                    p.destroy()
                except Exception:
                    pass            
                self._atualizar_apos_opcao_ferramenta()
                self.hint(f"Tamanho do candle: {l}px")

            for w in (row, preview, lbl):
                w.bind("<ButtonPress-1>", lambda ev, f=_escolher: f())

        pop.bind("<FocusOut>", lambda e: pop.destroy())
        pop.focus_set()
    
    def _abrir_paleta_espessura(self):
        pop = tk.Toplevel(self.root)
        self._ui_popup_windows.append(pop)
        pop.overrideredirect(True)
        pop.attributes("-topmost", True)
        pop.configure(bg="#0b0e11")

        try:
            x, y = self._toolbar_popup_xy(1380)
            pop.geometry(f"+{x}+{y}")
        except Exception:
            pop.geometry("+300+100")        

        frame = tk.Frame(pop, bg="#0b0e11",
                         highlightthickness=1,
                         highlightbackground="#3d6082")
        frame.pack(padx=1, pady=1)

        for nome, valor in ESPESSURAS:
            ativo = (valor == self.espessura)
            row = tk.Frame(frame, bg="#161c23" if not ativo else "#2f4a63",
                           cursor="hand2")
            row.pack(fill="x", padx=2, pady=1)

            preview = tk.Canvas(row, width=50, height=20,
                                bg=row["bg"], highlightthickness=0, bd=0)
            preview.pack(side="left", padx=(6, 8), pady=4)
            preview.create_line(4, 10, 46, 10,
                                fill="#e7edf4", width=valor)

            lbl = tk.Label(row, text=nome, bg=row["bg"], fg="#e7edf4",
                           font=("Segoe UI", 10), anchor="w")
            lbl.pack(side="left", fill="x", expand=True, padx=(0, 12))

            def _escolher(v=valor, p=pop):
                self.espessura = v
                self._atualizar_botao_espessura()
                try:
                    p.destroy()
                except Exception:
                    pass                        
                self._atualizar_apos_opcao_ferramenta()
                self.hint(f"Espessura definida: {v}px")

            for w in (row, preview, lbl):
                w.bind("<ButtonPress-1>", lambda e, f=_escolher: f())

        pop.bind("<FocusOut>", lambda e: pop.destroy())
        pop.focus_set()

    def set_mode_clear(self):
        self.set_mode("clear")

    def hint(self, text):
        self._atualizar_label_modo(text)
        try:
            self.root.after(2500, self._atualizar_label_modo)
        except tk.TclError:
            pass

    def ciclar_lousa(self):
        """3 cliques: branco → preto → off."""
        if self.overlay_active:
            self.hint("Saia do modo Sobrepor antes de usar a lousa.")
            return
        self.lousa_estado = (self.lousa_estado + 1) % 3
        if self.lousa_estado == 1:
            cor_fundo = "#ffffff"
            nome = "Lousa branca"
        elif self.lousa_estado == 2:
            cor_fundo = "#000000"
            nome = "Lousa preta"
        else:
            cor_fundo = None
            nome = "Lousa desativada"
        self._atualizar_icone_lousa()
        self._aplicar_lousa(cor_fundo)
        self.hint(nome)

    def _aplicar_lousa(self, cor_fundo):
        """Aplica ou remove a lousa. cor_fundo=None remove.
        A lousa é uma janela separada em tela cheia no monitor atual."""
        if cor_fundo is None:
            if self.lousa_frame is not None:
                self._guardar_desenho_lousa()
                self._restaurar_desenho_app()
                try:
                    self.lousa_frame.destroy()
                except Exception:
                    pass
                self.lousa_frame = None
                self.lousa_panel = None
            self.redraw_all()
            return

        if self.lousa_frame is None:
            self._guardar_desenho_app()
            self._limpar_desenho_atual()
            # Cria janela em tela cheia no monitor atual
            self.lousa_frame = tk.Toplevel(self.root)
            self.lousa_frame.overrideredirect(True)
            self.lousa_frame.attributes("-topmost", True)
            self.lousa_frame.configure(bg=cor_fundo)
            rect = self._monitor_rect_for_window()
            if rect:
                l, t, w, h = rect
                self.lousa_frame.geometry(f"{w}x{h}+{l}+{t}")
            else:
                self.lousa_frame.attributes("-fullscreen", True)

            p = Panel(self, self.lousa_frame, 0, source=True)
            p.canvas.config(bg=cor_fundo)
            self.lousa_panel = p
            # Rebaixa a lousa e sobe a toolbar por cima
            def _reforcar_toolbar():
                try:
                    self.lousa_frame.attributes("-topmost", False)
                    self.root.lift()
                    self.root.attributes("-topmost", True)
                except Exception:
                    pass
            self.root.after(150, _reforcar_toolbar)            
            c = p.canvas
            c.bind("<ButtonPress-1>", lambda e, p=p: self.press(e, p))
            c.bind("<B1-Motion>", lambda e, p=p: self.motion(e, p))
            c.bind("<ButtonRelease-1>", lambda e, p=p: self.release(e, p))
            c.bind("<Motion>", lambda e, p=p: self.hover(e, p))
            c.bind("<ButtonPress-3>",
                   lambda e, p=p: self.release_mode_atalho(p))
            self._restaurar_desenho_lousa()
        else:
            self.lousa_frame.configure(bg=cor_fundo)
            if self.lousa_panel is not None:
                self.lousa_panel.canvas.config(bg=cor_fundo)
        self.lousa_panel.redraw()

    def _guardar_desenho_app(self):
        """Salva os desenhos atuais do app antes de entrar na lousa."""
        self.lousa_workspace_backup = {
            "candles_por_painel": {
                id(p): list(p.candles) for p in self.panels},
            "lines": list(self.lines),
            "segments": list(self.segments),
            "arrows": list(self.arrows),
            "circles": list(self.circles),
            "rects": list(self.rects),
            "texts": list(self.texts),
            "highlights": list(self.highlights),
            "fib": self.fib,
        }

    def _restaurar_desenho_app(self):
        """Restaura os desenhos do app ao sair da lousa."""
        if not self.lousa_workspace_backup:
            return
        dados = self.lousa_workspace_backup
        for p in self.panels:
            if id(p) in dados["candles_por_painel"]:
                p.candles = list(dados["candles_por_painel"][id(p)])
        self.lines = list(dados["lines"])
        self.segments = list(dados["segments"])
        self.arrows = list(dados["arrows"])
        self.circles = list(dados["circles"])
        self.rects = list(dados["rects"])
        self.texts = list(dados["texts"])
        self.highlights = list(dados["highlights"])
        self.fib = dados["fib"]
        self.lousa_workspace_backup = None

    def _guardar_desenho_lousa(self):
        """Salva os desenhos atuais da lousa antes de sair."""
        dados = {
            "candles": list(self.lousa_panel.candles)
            if self.lousa_panel else [],
            "lines": list(self.lines),
            "segments": list(self.segments),
            "arrows": list(self.arrows),
            "circles": list(self.circles),
            "rects": list(self.rects),
            "texts": list(self.texts),
            "highlights": list(self.highlights),
            "fib": self.fib,
            "zoom": 1.0,
            "zoom_y": 1.0,
            "offset_x": 0.0,
            "offset_y": 0.0,
        }
        if self.lousa_panel is not None:
            dados["zoom"] = self.lousa_panel.zoom
            dados["zoom_y"] = self.lousa_panel.zoom_y
            dados["offset_x"] = self.lousa_panel.offset_x
            dados["offset_y"] = self.lousa_panel.offset_y
        self.lousa_desenho = dados

    def _restaurar_desenho_lousa(self):
        """Restaura os desenhos da lousa ao entrar nela."""
        d = self.lousa_desenho
        if self.lousa_panel is not None:
            self.lousa_panel.candles = list(d.get("candles", []))
            self.lousa_panel.zoom = d.get("zoom", 1.0)
            self.lousa_panel.zoom_y = d.get("zoom_y", 1.0)
            self.lousa_panel.offset_x = d.get("offset_x", 0.0)
            self.lousa_panel.offset_y = d.get("offset_y", 0.0)
        self.lines = list(d["lines"])
        self.segments = list(d["segments"])
        self.arrows = list(d["arrows"])
        self.circles = list(d["circles"])
        self.rects = list(d["rects"])
        self.texts = list(d["texts"])
        self.highlights = list(d["highlights"])
        self.fib = d["fib"]    

        # Ao sair da lousa a janela/painel é destruída. Na próxima entrada,
        # os rascunhos continuam, mas precisam apontar para o novo painel.
        if self.lousa_panel is not None:
            objetos = list(self.lousa_panel.candles)
            for nome in ("lines", "segments", "arrows", "circles", "rects",
                         "texts", "highlights"):
                objetos.extend(getattr(self, nome))
            if self.fib is not None:
                objetos.append(self.fib)
            for objeto in objetos:
                if hasattr(objeto, "painel"):
                    objeto.painel = self.lousa_panel

    def _limpar_desenho_atual(self):
        """Limpa os desenhos atuais (sem tocar no histórico)."""
        self.lines = []
        self.segments = []
        self.arrows = []
        self.circles = []
        self.rects = []
        self.texts = []
        self.highlights = []
        self.fib = None
        for p in self.panels:
            p.candles = []
    


    # ============================================================
    # Zoom do scroll
    # ============================================================
    def _on_mousewheel(self, event, painel):
        """Chamado quando o usuário rola o scroll em cima de um painel.
        Windows: event.delta é positivo pra cima, negativo pra baixo."""
        if painel.overlay:
            return
        direcao = 1 if event.delta > 0 else -1
        if painel.aplicar_zoom(event, direcao):
            painel.redraw()

    # ============================================================
    # Print
    # ============================================================
    def _grab_monitor_snapshot(self, hide_app=True):
        if not HAS_PIL:
            return None, None

        rect = self._monitor_rect_for_window()

        if hide_app:
            try:
                self.root.attributes("-alpha", 0.0)
                self.root.update_idletasks()
                self.root.update()
                time.sleep(0.10)
                self.root.update()
            except tk.TclError:
                pass

        img = None
        try:
            try:
                full = ImageGrab.grab(all_screens=True)
            except TypeError:
                full = ImageGrab.grab()

            if rect:
                try:
                    import ctypes
                    vx = ctypes.windll.user32.GetSystemMetrics(76)
                    vy = ctypes.windll.user32.GetSystemMetrics(77)
                    l, t, w, h = rect
                    img = full.crop((l - vx, t - vy,
                                     l - vx + w, t - vy + h))
                except Exception:
                    img = full
            else:
                img = full
                try:
                    rect = (self.root.winfo_rootx(),
                            self.root.winfo_rooty(),
                            self.root.winfo_width(),
                            self.root.winfo_height())
                except Exception:
                    rect = (0, 0, img.size[0], img.size[1])
        except Exception:
            img = None

        if hide_app:
            try:
                self.root.attributes("-alpha", 1.0)
                self.root.update_idletasks()
            except tk.TclError:
                pass

        return img, rect

    def _open_selection_window(self, snap, rect, filename_prefix):
        if rect is None:
            rect = (0, 0, snap.size[0], snap.size[1])

        mon_left, mon_top, mon_w, mon_h = rect

        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        filename = f"{filename_prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        filepath = os.path.join(desktop_path, filename)

        self.print_select_win = tk.Toplevel(self.root)
        self.print_select_win.title("Selecione a área do print")
        self.print_select_win.attributes("-topmost", True)
        self.print_select_win.geometry(f"{mon_w}x{mon_h}+{mon_left}+{mon_top}")
        self.print_select_win.overrideredirect(True)

        canvas = tk.Canvas(self.print_select_win, cursor="cross",
                           highlightthickness=0, bd=0)
        canvas.pack(fill="both", expand=True)

        self._print_snap_photo = ImageTk.PhotoImage(snap)
        canvas.create_image(0, 0, anchor="nw", image=self._print_snap_photo)

        def on_btn_down(e):
            self.print_start_pos = (e.x, e.y)
            self.print_rect_id = canvas.create_rectangle(
                e.x, e.y, e.x, e.y, outline="red", width=2)

        def on_drag(e):
            if self.print_start_pos:
                cx0, cy0 = self.print_start_pos
                canvas.coords(self.print_rect_id, cx0, cy0, e.x, e.y)

        def on_btn_up(e):
            if not self.print_start_pos:
                return
            cx0, cy0 = self.print_start_pos
            cx1, cy1 = e.x, e.y
            self.print_start_pos = None

            try:
                self.print_select_win.destroy()
            except Exception:
                pass
            self.print_select_win = None
            self._print_snap_photo = None

            left = max(0, min(cx0, cx1))
            right = min(mon_w, max(cx0, cx1))
            top = max(0, min(cy0, cy1))
            bottom = min(mon_h, max(cy0, cy1))

            if right - left < 5 or bottom - top < 5:
                out = snap
            else:
                out = snap.crop((left, top, right, bottom))

            try:
                out.save(filepath)
                self.hint("Print salvo no Desktop!")
            except Exception as exc:
                self.hint(f"Erro ao salvar print: {exc}")

        def on_right(e):
            if self.print_select_win is not None:
                try:
                    self.print_select_win.destroy()
                except Exception:
                    pass
                self.print_select_win = None
                self._print_snap_photo = None
            self.hint("Print cancelado.")

        canvas.bind("<ButtonPress-1>", on_btn_down)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_btn_up)
        canvas.bind("<ButtonPress-3>", on_right)

    def take_screenshot(self):
        if not HAS_PIL:
            self.hint("Instale o Pillow para tirar print: pip install pillow")
            return
        snap, rect = self._grab_monitor_snapshot(hide_app=True)
        if snap is None:
            self.hint("Falha ao capturar a tela.")
            return
        self._open_selection_window(snap, rect, "Print_CandleProjector")

    def take_screenshot_app(self):
        if not HAS_PIL:
            self.hint("Instale o Pillow para tirar print: pip install pillow")
            return
        try:
            self.root.update_idletasks()
            self.root.update()
            time.sleep(0.05)
            self.root.update()
        except tk.TclError:
            pass
        snap, rect = self._grab_monitor_snapshot(hide_app=False)
        if snap is None:
            self.hint("Falha ao capturar a tela.")
            return
        self._open_selection_window(snap, rect, "Print_CandleProjector_App")

    # ============================================================
    # Recolher
    # ============================================================
    def _prepare_relogio_recolhido(self):
        """Carrega a faixa do relógio e limpa os dígitos de demonstração."""
        if not HAS_PIL:
            return
        try:
            caminho = caminho_recurso(
                os.path.join("assets", "collapsed_clock_banner.png"))
            imagem = Image.open(caminho).convert("RGBA")
            ImageDraw.Draw(imagem).rounded_rectangle(
                (1450, 97, 1875, 224), radius=10, fill=(5, 6, 9, 255))
            self._collapsed_clock_template = imagem
        except Exception as exc:
            self._collapsed_clock_template = None
            print(f"[aviso] Falha ao carregar a arte do relógio: {exc}")

    def _render_relogio_recolhido(self):
        """Desenha a hora local em segmentos digitais sobre a nova arte."""
        if not HAS_PIL or self._collapsed_clock_template is None:
            return None

        imagem = self._collapsed_clock_template.copy()
        hora = datetime.now().strftime("%H:%M:%S")
        largura_digito, altura_digito, esp_segmento = 48, 80, 8
        espacamento, largura_dois_pontos = 7, 12
        largura_total = 6 * (largura_digito + espacamento) + 2 * (
            largura_dois_pontos + espacamento)
        x = (1450 + 1875 - largura_total) // 2
        y = 111
        acesos = {
            "0": "abcdef", "1": "bc", "2": "abged",
            "3": "abgcd", "4": "fgbc", "5": "afgcd",
            "6": "afgecd", "7": "abc", "8": "abcdefg",
            "9": "abfgcd",
        }

        def horizontal(px, py):
            t, w = esp_segmento, largura_digito
            return [(px + t, py), (px + w - t, py),
                    (px + w, py + t // 2), (px + w - t, py + t),
                    (px + t, py + t), (px, py + t // 2)]

        def vertical(px, py, meio):
            t = esp_segmento
            return [(px + t // 2, py), (px + t, py + t // 2),
                    (px + t, meio - t // 2), (px + t // 2, meio),
                    (px, meio - t // 2), (px, py + t // 2)]

        meio = y + altura_digito // 2

        def poligono_segmento(nome, px):
            if nome == "a":
                return horizontal(px, y)
            if nome == "g":
                return horizontal(px, meio - esp_segmento // 2)
            if nome == "d":
                return horizontal(px, y + altura_digito - esp_segmento)
            if nome == "f":
                return vertical(px, y + esp_segmento // 2, meio)
            if nome == "b":
                return vertical(px + largura_digito - esp_segmento,
                                y + esp_segmento // 2, meio)
            if nome == "e":
                return vertical(px, meio, y + altura_digito - esp_segmento // 2)
            return vertical(px + largura_digito - esp_segmento,
                            meio, y + altura_digito - esp_segmento // 2)

        glow = Image.new("RGBA", imagem.size, (0, 0, 0, 0))
        glow_draw = ImageDraw.Draw(glow)
        desenho = []
        cursor_x = x
        for caractere in hora:
            if caractere == ":":
                desenho.append(("colon", cursor_x))
                cursor_x += largura_dois_pontos + espacamento
                continue
            for nome in "abcdefg":
                pontos = poligono_segmento(nome, cursor_x)
                ligado = nome in acesos[caractere]
                desenho.append(("on" if ligado else "off", pontos))
                if ligado:
                    glow_draw.polygon(pontos, fill=(190, 210, 255, 100))
            cursor_x += largura_digito + espacamento

        imagem = Image.alpha_composite(
            imagem, glow.filter(ImageFilter.GaussianBlur(4)))
        draw = ImageDraw.Draw(imagem)
        for estado, dados in desenho:
            if estado == "colon":
                for cy in (y + 25, y + altura_digito - 25):
                    draw.ellipse((dados + 2, cy - 4, dados + 10, cy + 4),
                                 fill=(238, 243, 255, 255))
            else:
                cor = ((238, 243, 255, 255) if estado == "on"
                       else (30, 33, 40, 180))
                draw.polygon(dados, fill=cor)

        if self._collapsed_clock_size:
            imagem = imagem.resize(self._collapsed_clock_size, Image.LANCZOS)
        return ImageTk.PhotoImage(imagem)

    def collapse_app(self):
        if self.collapsed:
            return

        # Destrói TODOS os Sobrepor antes de recolher
        self._monitor_cliques_ativo = False
        self._remover_hooks_globais()
        for _, p in getattr(self, "_overlay_paineis", []):
            try:
                p.canvas.master.destroy()
            except Exception:
                pass
        self._overlay_paineis = []
        self.overlay_panel = None
        self.overlay_holder = None
        self.overlay_active = False            

        if self.overlay_active:
            self.exit_overlay()

        if self.clock_window is not None:
            self.clock_window.close()
            self.clock_window = None

        self.root.update_idletasks()
        monitor = self._monitor_rect_for_toolbar()
        if monitor:
            left, top, monitor_w, _monitor_h = monitor
        else:
            left, top = self.root.winfo_rootx(), self.root.winfo_rooty()
            monitor_w = max(self.root.winfo_width(), COLLAPSED_CLOCK_MIN_WIDTH)
        max_width = max(COLLAPSED_CLOCK_MIN_WIDTH, monitor_w - 24)
        banner_w = max(COLLAPSED_CLOCK_MIN_WIDTH, min(
            COLLAPSED_CLOCK_MAX_WIDTH,
            int(monitor_w * COLLAPSED_CLOCK_MONITOR_RATIO), max_width))
        banner_h = round(banner_w * COLLAPSED_CLOCK_ASPECT)
        self._collapsed_clock_size = (banner_w, banner_h)
        self._collapsed_clock_target = (
            left + (monitor_w - banner_w) // 2, top)
        if self._collapsed_clock_template is None:
            self._prepare_relogio_recolhido()
        self.root.withdraw()

        self.icon_window = tk.Toplevel(self.root)
        self.icon_window.overrideredirect(True)
        self.icon_window.attributes("-topmost", True)
        self.icon_window.configure(bg="#08090a")
        try:
            self.icon_window.attributes("-transparentcolor", "#08090a")
        except tk.TclError:
            pass
        self.icon_window.geometry(f"{banner_w}x{banner_h}+0+0")

        try:
            if getattr(self, "_icone_photo_principal", None):
                self.icon_window.iconphoto(False, self._icone_photo_principal)
        except Exception:
            pass

        self.icon_canvas = tk.Canvas(
            self.icon_window, bg="#08090a", highlightthickness=0, bd=0)
        self.icon_canvas.pack(fill="both", expand=True)

        self._collapsed_clock_photo = self._render_relogio_recolhido()
        if self._collapsed_clock_photo is not None:
            self._collapsed_clock_image_id = self.icon_canvas.create_image(
                0, 0, image=self._collapsed_clock_photo, anchor="nw")
        else:
            self._collapsed_clock_image_id = self.icon_canvas.create_text(
                banner_w // 2, banner_h // 2,
                text=datetime.now().strftime("%H:%M:%S"),
                fill="#eef3ff", font=("Consolas", max(12, banner_h // 3), "bold"))

        self.icon_canvas.bind("<ButtonPress-1>", self._icon_press)
        self.icon_canvas.bind("<B1-Motion>", self._icon_motion)
        self.icon_canvas.bind("<ButtonRelease-1>", self._icon_release)

        self.collapsed = True
        self._icon_drag_origin = None
        self._icon_dragged = False
        self.icon_window.update_idletasks()
        self._posicionar_janela_recolhida(*self._collapsed_clock_target)
        self.icon_window.after(0, self._atualizar_relogio_recolhido)

    def _atualizar_relogio_recolhido(self):
        """Atualiza o mostrador digital com o horário local do Windows."""
        if (not self.collapsed or self.icon_window is None
                or self.icon_canvas is None):
            return
        try:
            if self._collapsed_clock_template is not None:
                self._collapsed_clock_photo = self._render_relogio_recolhido()
                self.icon_canvas.itemconfigure(
                    self._collapsed_clock_image_id,
                    image=self._collapsed_clock_photo)
            else:
                self.icon_canvas.itemconfigure(
                    self._collapsed_clock_image_id,
                    text=datetime.now().strftime("%H:%M:%S"))
            self._collapsed_clock_after = self.icon_window.after(
                1000, self._atualizar_relogio_recolhido)
        except tk.TclError:
            pass

    def _janela_recolhida_hwnd(self):
        """Retorna o HWND externo da janela, não o canvas interno do Tk."""
        try:
            handle = self.icon_window.wm_frame()
            hwnd = int(str(handle), 0)
            if win32gui.IsWindow(hwnd):
                return hwnd
        except Exception:
            pass
        return int(self.icon_window.winfo_id())

    def _posicionar_janela_recolhida(self, x, y):
        """Move o relógio com coordenadas absolutas do desktop do Windows."""
        largura, altura = (self._collapsed_clock_size or (
            COLLAPSED_CLOCK_MAX_WIDTH,
            round(COLLAPSED_CLOCK_MAX_WIDTH * COLLAPSED_CLOCK_ASPECT)))
        try:
            hwnd = self._janela_recolhida_hwnd()
            flags = win32con.SWP_NOACTIVATE | win32con.SWP_SHOWWINDOW
            win32gui.SetWindowPos(
                hwnd, win32con.HWND_TOPMOST, int(x), int(y),
                largura, altura, flags)
            return True
        except Exception:
            # Fallback para instalações onde pywin32 não move o HWND externo.
            if x >= 0 and y >= 0:
                try:
                    self.icon_window.geometry(
                        f"{largura}x{altura}+{int(x)}+{int(y)}")
                    return True
                except tk.TclError:
                    pass
            return False

    def _icon_press(self, event):
        try:
            hwnd = self._janela_recolhida_hwnd()
            cursor_x, cursor_y = win32gui.GetCursorPos()
            left, top, _right, _bottom = win32gui.GetWindowRect(hwnd)
            self._icon_drag_origin = (cursor_x, cursor_y, left, top)
        except Exception:
            self._icon_drag_origin = (
                event.x_root, event.y_root,
                self.icon_window.winfo_rootx(),
                self.icon_window.winfo_rooty())
        self._icon_dragged = False

    def _icon_motion(self, event):
        if not self._icon_drag_origin:
            return
        x0, y0, wx, wy = self._icon_drag_origin
        try:
            cursor_x, cursor_y = win32gui.GetCursorPos()
        except Exception:
            cursor_x, cursor_y = event.x_root, event.y_root
        dx = cursor_x - x0
        dy = cursor_y - y0
        if abs(dx) > 2 or abs(dy) > 2:
            self._icon_dragged = True
        self._posicionar_janela_recolhida(wx + dx, wy + dy)

    def _icon_release(self, event):
        if not self._icon_dragged:
            self.expand_from_collapse()
        self._icon_drag_origin = None
        self._icon_dragged = False

    def expand_from_collapse(self):
        if not self.collapsed:
            return

        if self._collapsed_clock_after and self.icon_window is not None:
            try:
                self.icon_window.after_cancel(self._collapsed_clock_after)
            except tk.TclError:
                pass
        self._collapsed_clock_after = None

        try:
            if self.icon_window is not None:
                self.icon_window.destroy()
        except Exception:
            pass
        self.icon_window = None
        self.icon_canvas = None
        self._collapsed_clock_photo = None
        self._collapsed_clock_image_id = None

        self.root.deiconify()
        self.root.lift()
        self.collapsed = False
        self.redraw_all()

    # ============================================================
    # Multi-monitor
    # ============================================================
    def _monitor_rect_for_toolbar(self):
        """Encontra o monitor da toolbar pelo HWND, sem usar o centro da tela."""
        try:
            handle = self.root.wm_frame()
            hwnd = int(str(handle), 0)
            if not win32gui.IsWindow(hwnd):
                hwnd = int(self.root.winfo_id())
            monitor = win32gui.MonitorFromWindow(
                hwnd, win32con.MONITOR_DEFAULTTONEAREST)
            info = win32gui.GetMonitorInfo(monitor)
            left, top, right, bottom = info["Monitor"]
            return (left, top, right-left, bottom-top)
        except Exception:
            # Fallback para a detecção existente, se a API do monitor falhar.
            return self._monitor_rect_for_window()

    def _monitor_rect_for_window(self):
        try:
            import ctypes
            from ctypes import wintypes
            user32 = ctypes.windll.user32
            cx = self.root.winfo_rootx() + self.root.winfo_width() // 2
            cy = self.root.winfo_rooty() + self.root.winfo_height() // 2
            rects = []

            def _cb(hMon, hdc, lprc, data):
                rc = lprc.contents
                rects.append((rc.left, rc.top, rc.right, rc.bottom))
                return True

            cb = ctypes.WINFUNCTYPE(
                wintypes.BOOL, wintypes.HMONITOR, wintypes.HDC,
                ctypes.POINTER(wintypes.RECT), wintypes.LPARAM)(_cb)
            user32.EnumDisplayMonitors(None, None, cb, 0)
            for (l, t, r, b) in rects:
                if l <= cx < r and t <= cy < b:
                    return (l, t, r - l, b - t)
        except Exception:
            pass
        return None

    # ============================================================
    # Sobrepor
    # ============================================================
    def toggle_overlay(self):
        if self.overlay_active:
            self.exit_overlay()
        else:
            self.open_overlay()

    def _iniciar_monitor_cliques(self):
        """Intercepta mouse/teclado enquanto uma ferramenta de desenho está ativa."""
        if self._mouse_hook and self._keyboard_hook:
            return

        import ctypes
        from ctypes import wintypes

        user32 = ctypes.windll.user32
        WH_MOUSE_LL, WH_KEYBOARD_LL = 14, 13
        HC_ACTION = 0
        WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP = 0x0200, 0x0201, 0x0202
        WM_RBUTTONDOWN, WM_RBUTTONUP = 0x0204, 0x0205
        WM_KEYDOWN, WM_KEYUP = 0x0100, 0x0101
        WM_SYSKEYDOWN, WM_SYSKEYUP = 0x0104, 0x0105
        VK_CONTROL, VK_Z = 0x11, 0x5A
        LRESULT = ctypes.c_ssize_t
        ULONG_PTR = ctypes.c_size_t

        class POINT(ctypes.Structure):
            _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]

        class MSLLHOOKSTRUCT(ctypes.Structure):
            _fields_ = [("pt", POINT), ("mouseData", wintypes.DWORD),
                        ("flags", wintypes.DWORD), ("time", wintypes.DWORD),
                        ("dwExtraInfo", ULONG_PTR)]

        class KBDLLHOOKSTRUCT(ctypes.Structure):
            _fields_ = [("vkCode", wintypes.DWORD), ("scanCode", wintypes.DWORD),
                        ("flags", wintypes.DWORD), ("time", wintypes.DWORD),
                        ("dwExtraInfo", ULONG_PTR)]

        HOOKPROC = ctypes.WINFUNCTYPE(LRESULT, ctypes.c_int,
                                      wintypes.WPARAM, wintypes.LPARAM)
        user32.CallNextHookEx.argtypes = [wintypes.HHOOK, ctypes.c_int,
                                          wintypes.WPARAM, wintypes.LPARAM]
        user32.CallNextHookEx.restype = LRESULT
        user32.SetWindowsHookExW.argtypes = [ctypes.c_int, HOOKPROC,
                                             wintypes.HINSTANCE, wintypes.DWORD]
        user32.SetWindowsHookExW.restype = wintypes.HHOOK

        def mouse_proc(n_code, w_param, l_param):
            try:
                if n_code == HC_ACTION:
                    msg = int(w_param)
                    if msg == WM_RBUTTONUP and self._overlay_right_captured:
                        self._overlay_right_captured = False
                        self._enfileirar_evento_overlay("right", 0, 0)
                        return 1
                    if self.overlay_active:                    
                        info = ctypes.cast(
                            l_param, ctypes.POINTER(MSLLHOOKSTRUCT)).contents
                        x, y = info.pt.x, info.pt.y
                        if msg == WM_LBUTTONDOWN:
                            if not self._ponto_em_cima_da_ui_hook(x, y):
                                panel, px, py = self._encontrar_painel_do_monitor(x, y)
                                if panel is not None:
                                    # No modo Mover, só captura o clique sobre
                                    # um desenho; áreas vazias continuam no Windows.
                                    tem_alvo = (self.mode != "move" or
                                                self.hit_test(panel, px, py) is not None)
                                    if tem_alvo:
                                        self._overlay_mouse_captured = True
                                        self._overlay_mouse_panel = panel
                                        self._enfileirar_evento_overlay("down", x, y)
                                        return 1
                        elif msg == WM_MOUSEMOVE and self._overlay_mouse_captured:
                            # Deixa WM_MOUSEMOVE seguir ao Windows para manter
                            # o cursor físico em movimento; o down/up continuam
                            # bloqueados e o desenho usa as coordenadas do hook.
                            self._enfileirar_evento_overlay("motion", x, y)
                        elif msg == WM_LBUTTONUP and self._overlay_mouse_captured:
                            self._overlay_mouse_captured = False
                            self._enfileirar_evento_overlay("up", x, y)
                            return 1
                        elif msg == WM_RBUTTONDOWN and self.mode != "move" \
                                and not self._ponto_em_cima_da_ui_hook(x, y):
                            self._overlay_right_captured = True
                            self._overlay_mouse_captured = False
                            return 1
            except Exception:
                pass
            return user32.CallNextHookEx(self._mouse_hook, n_code,
                                         w_param, l_param)

        def keyboard_proc(n_code, w_param, l_param):
            try:
                if n_code == HC_ACTION and self.overlay_active:
                    # Atualiza o estado do Shift direto do Windows
                    VK_SHIFT = 0x10
                    self._shift_pressionado = bool(
                        user32.GetAsyncKeyState(VK_SHIFT) & 0x8000)                    
                    msg = int(w_param)
                    info = ctypes.cast(
                        l_param, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
                    ctrl_down = bool(
                        user32.GetAsyncKeyState(VK_CONTROL) & 0x8000)
                    if info.vkCode == VK_Z and ctrl_down:
                        if msg in (WM_KEYDOWN, WM_SYSKEYDOWN):
                            if not self._overlay_ctrlz_key_down:
                                self._overlay_ctrlz_key_down = True
                                self._enfileirar_evento_overlay("undo", 0, 0)
                            return 1
                        if msg in (WM_KEYUP, WM_SYSKEYUP) \
                                and self._overlay_ctrlz_key_down:
                            self._overlay_ctrlz_key_down = False
                            return 1
                    elif info.vkCode == VK_Z \
                            and msg in (WM_KEYUP, WM_SYSKEYUP):
                        if self._overlay_ctrlz_key_down:
                            self._overlay_ctrlz_key_down = False
                            return 1
            except Exception:
                pass
            return user32.CallNextHookEx(self._keyboard_hook, n_code,
                                         w_param, l_param)

        self._mouse_hook_proc = HOOKPROC(mouse_proc)
        self._keyboard_hook_proc = HOOKPROC(keyboard_proc)
        # Hooks de baixo nível são entregues à thread que instalou o hook;
        # por isso lpModule pode ser NULL (o loop do Tk processa as mensagens).
        self._mouse_hook = user32.SetWindowsHookExW(
            WH_MOUSE_LL, self._mouse_hook_proc, None, 0)
        self._keyboard_hook = user32.SetWindowsHookExW(
            WH_KEYBOARD_LL, self._keyboard_hook_proc, None, 0)
        if not self._mouse_hook or not self._keyboard_hook:
            self._remover_hooks_globais()
            raise RuntimeError("O Windows não permitiu capturar mouse/teclado.")

    def _ponto_em_cima_da_toolbar_hook(self, x, y):
        """Versão sem Tkinter para uso dentro do callback do Windows."""
        r = self._toolbar_rect_cache
        if not r:
            return False
        left, top, right, bottom = r
        return left <= x <= right and top <= y <= bottom

    def _ponto_em_cima_da_ui_hook(self, x, y):
        """Ignora cliques na barra e nos menus do próprio programa."""
        if self._ponto_em_cima_da_toolbar_hook(x, y):
            return True
        for left, top, right, bottom in self._ui_popup_rects_cache:
            if left <= x <= right and top <= y <= bottom:
                return True
        return False

    def _atualizar_ui_popup_rects_cache(self):
        """Atualiza fora do callback Win32 os retângulos dos menus Tk abertos."""
        rects = []
        popups_validos = []
        for popup in self._ui_popup_windows:
            try:
                if not popup.winfo_exists():
                    continue
                popups_validos.append(popup)
                if not popup.winfo_viewable():
                    continue
                left = popup.winfo_rootx()
                top = popup.winfo_rooty()
                right = left + popup.winfo_width()
                bottom = top + popup.winfo_height()
                rects.append((left, top, right, bottom))
            except tk.TclError:
                continue
        self._ui_popup_windows = popups_validos
        self._ui_popup_rects_cache = rects

    def _enfileirar_evento_overlay(self, tipo, x, y):
        fila = self._overlay_event_queue
        evento = (tipo, x, y)
        if tipo == "motion" and fila and fila[-1][0] == "motion":
            fila[-1] = evento
        else:
            fila.append(evento)

    def _processar_eventos_overlay(self):
        self._overlay_event_timer = None
        self._toolbar_rect_cache = self._toolbar_rect()
        self._atualizar_ui_popup_rects_cache()
        fila = self._overlay_event_queue
        eventos = fila[:]
        fila.clear()

        for tipo, x, y in eventos:
            if tipo == "down" and self.overlay_active:
                self._on_clique_global_down(x, y)
            elif tipo == "motion" and self.overlay_active:
                self._on_motion_global(x, y)
            elif tipo == "up" and self.overlay_active:
                self._on_clique_global_up(x, y)
                self._overlay_mouse_panel = None            
            elif tipo == "undo" and self.overlay_active:
                self.undo()
            elif tipo == "right" and self.overlay_active:
                self.after_cancel_monitor_e_voltar_move()

        if self.overlay_active and self._mouse_hook:
            try:
                self._overlay_event_timer = self.root.after(
                    8, self._processar_eventos_overlay)
            except tk.TclError:
                self._overlay_event_timer = None

    def _remover_hooks_globais(self):
        import ctypes
        user32 = ctypes.windll.user32
        if self._mouse_hook:
            user32.UnhookWindowsHookEx(self._mouse_hook)
            self._mouse_hook = None
        if self._keyboard_hook:
            user32.UnhookWindowsHookEx(self._keyboard_hook)
            self._keyboard_hook = None
        self._mouse_hook_proc = None
        self._keyboard_hook_proc = None
        self._overlay_mouse_captured = False
        self._overlay_mouse_panel = None
        self._overlay_pending_motion = None
        self._overlay_motion_after = None
        self._overlay_event_queue.clear()
        if self._overlay_event_timer is not None:
            try:
                self.root.after_cancel(self._overlay_event_timer)
            except tk.TclError:
                pass
            self._overlay_event_timer = None
        self._overlay_right_captured = False
        self._overlay_ctrlz_key_down = False

    def _loop_monitor_cliques(self):
        """Compatibilidade com a versão anterior; a captura usa hooks."""
        return

    def after_cancel_monitor_e_voltar_move(self):
        """Botão direito global: se Shift pressionado, troca candle ↔ traço;
        senão, volta ao modo Mover."""
        if getattr(self, "_shift_pressionado", False):
            self._atalho_candle_traco()
            return
        self.release_mode()    

    def _monitor_rects(self):
        """Retorna lista de todos os monitores: [(l, t, w, h), ...]"""
        try:
            import ctypes
            from ctypes import wintypes
            user32 = ctypes.windll.user32
            rects = []

            def _cb(hMon, hdc, lprc, data):
                rc = lprc.contents
                rects.append((rc.left, rc.top,
                              rc.right - rc.left, rc.bottom - rc.top))
                return True

            cb = ctypes.WINFUNCTYPE(
                wintypes.BOOL, wintypes.HMONITOR, wintypes.HDC,
                ctypes.POINTER(wintypes.RECT), wintypes.LPARAM)(_cb)
            user32.EnumDisplayMonitors(None, None, cb, 0)
            return rects
        except Exception:
            return []

    def _toolbar_rect(self):
        """Retorna (l, t, r, b) da janela da toolbar."""
        try:
            l = self.root.winfo_x()
            t = self.root.winfo_y()
            r = l + self.root.winfo_width()
            b = t + self.root.winfo_height()
            return (l, t, r, b)
        except Exception:
            return None

    def _ponto_em_cima_da_toolbar(self, x, y):
        """Verifica se o ponto (x, y) tá em cima da toolbar ou de algum
        diálogo do app (login, etc)."""
        try:
            # Toolbar
            r = self._toolbar_rect()
            if r:
                l, t, rr, bb = r
                if l <= x <= rr and t <= y <= bb:
                    return True
            # Janela da lousa
            if self.lousa_frame is not None:
                try:
                    l = self.lousa_frame.winfo_x()
                    t = self.lousa_frame.winfo_y()
                    rr = l + self.lousa_frame.winfo_width()
                    bb = t + self.lousa_frame.winfo_height()
                    # Lousa cobre tudo — mas não bloqueia clique
                except Exception:
                    pass
            return False
        except Exception:
            return False

    def _encontrar_painel_do_monitor(self, x, y):
        """Acha qual Toplevel (por monitor) contém o ponto (x, y)."""
        for (l, t, w, h), painel in getattr(self, "_overlay_paineis", []):
            if l <= x < l + w and t <= y < t + h:
                # Coordenadas relativas ao monitor
                rel_x = x - l
                rel_y = y - t
                return painel, rel_x, rel_y
        return None, None, None

    def _on_clique_global_down(self, x=None, y=None):
        """Clique esquerdo global: se não tá em cima da toolbar,
        inicia um desenho no Sobrepor do monitor certo."""
        if not self.overlay_active:
            return    
        if x is None or y is None:
            x, y = win32api.GetCursorPos()
        if self._ponto_em_cima_da_toolbar(x, y):
            return

        painel, rx, ry = self._encontrar_painel_do_monitor(x, y)
        # Força foco na Toplevel do painel clicado
        if painel is not None:
            try:
                hwnd = painel.canvas.winfo_id()
                win32gui.SetFocus(hwnd)
            except Exception:
                pass        
        if painel is None:
            return

        # Cria um evento fake e chama o press normal
        class _E:
            pass
        e = _E()
        e.x = rx
        e.y = ry
        self.press(e, painel)

    def _on_motion_global(self, x=None, y=None):
        """Mouse se moveu com botão esquerdo pressionado:
        passa pro motion do painel certo."""
        if not self.overlay_active:
            return    
        if x is None or y is None:
            x, y = win32api.GetCursorPos()
        painel, rx, ry = self._encontrar_painel_do_monitor(x, y)
        if painel is None:
            return
        class _E:
            pass
        e = _E()
        e.x = rx
        e.y = ry
        self.motion(e, painel)
    
    def _on_clique_global_up(self, x=None, y=None):
        """Clique esquerdo global: solta o desenho atual."""
        if not self.overlay_active:
            return
        if x is None or y is None:
            x, y = win32api.GetCursorPos()
        painel, rx, ry = self._encontrar_painel_do_monitor(x, y)
        if painel is None and self._overlay_mouse_panel is not None:
            painel = self._overlay_mouse_panel
            for (left, top, width, height), candidate in self._overlay_paineis:
                if candidate is painel:
                    rx = min(max(x - left, 0), width - 1)
                    ry = min(max(y - top, 0), height - 1)
                    break
        if painel is None:
            return
        class _E:
            pass
        e = _E()
        e.x = rx
        e.y = ry
        self.release(e, painel)
        
    def open_overlay(self):
        """Abre o Sobrepor em TODOS os monitores, em tempo real.
        Cria 1 Toplevel invisível por monitor."""
        if self.overlay_active:
            return
        self.overlay_active = True

        # Se já existem os painéis, só mostra de novo
        if getattr(self, "_overlay_paineis", []):
            try:
                for rect, p in self._overlay_paineis:
                    try:
                        # Reconfigura a cor se mudou
                        pass
                    except Exception:
                        pass
                for painel in self._overlay_paineis:
                    pass
            except Exception:
                pass

        # Cria painéis pra cada monitor (SÓ SE AINDA NÃO TEM)
        if not getattr(self, "_overlay_paineis", []):
            self._overlay_paineis = []
            monitores = self._monitor_rects()

            for (l, t, w, h) in monitores:
                tl = tk.Toplevel(self.root)
                tl.overrideredirect(True)
                tl.configure(bg=OVERLAY_TRANSPARENT)
                tl.attributes("-topmost", True)
                tl.attributes("-transparentcolor", OVERLAY_TRANSPARENT)
                tl.geometry(f"{w}x{h}+{l}+{t}")

                p = Panel(self, tl, 0, overlay=True)
                p.canvas.configure(bg=OVERLAY_TRANSPARENT)

                self._overlay_paineis.append(((l, t, w, h), p))

                c = p.canvas
                c.bind("<Control-z>", lambda e: self.undo())
                c.bind("<Control-Z>", lambda e: self.undo())
                c.bind("<Delete>", lambda e: self.delete_selected())
                c.bind("<Escape>", self._on_escape)
                tl.focus_set()                

                try:
                    tl.deiconify()
                    tl.lift()
                except Exception:
                    pass

        # Painel ativo = primeiro (só pra compatibilidade)
        if self._overlay_paineis:
            self.overlay_panel = self._overlay_paineis[0][1]
            self.overlay_holder = self.overlay_panel.canvas.master
            self.active_panel = self.overlay_panel

        # Toolbar por cima
        self.root.lift()
        self.root.attributes("-topmost", True)

        # Redesenha todos
        for _, p in self._overlay_paineis:
            p.redraw()

        # Intercepta eventos globais para não chegarem ao aplicativo atrás.
        try:
            self._toolbar_rect_cache = self._toolbar_rect()
            self._iniciar_monitor_cliques()
            self._overlay_event_timer = self.root.after(
                8, self._processar_eventos_overlay)
        except Exception as exc:
            self.overlay_active = False
            self.root.bell()
            messagebox.showerror(
                "Sobreposição",
                "Não foi possível ativar a captura do mouse e teclado.\n\n"
                f"Detalhe: {exc}", parent=self.root)
            return
    



    

    def _set_clickthrough(self, on):
        """Com pywin32, a captura de clique é externa.
        Aqui só ajustamos a transparência das Toplevels."""
        pass    
    def exit_overlay(self):
        """Sai do Sobrepor: para o monitor, esconde as Toplevels.
        Os desenhos ficam salvos."""
        if not self.overlay_active:
            return
        self.overlay_active = False
        self._monitor_cliques_ativo = False
        self._remover_hooks_globais()
        # Libera qualquer captura Win32 legada.
        try:
            win32gui.ReleaseCapture()
        except Exception:
            pass

        if self.overlay_panel is not None:
            self.overlay_candles = self.overlay_panel.candles

        # NÃO esconde os painéis — deixa visíveis pra os desenhos
        # continuarem aparecendo. Só não capta clique (pywin32 cuida disso).

        self.root.lift()
        self.root.attributes("-topmost", True)
        self.redraw_all()            

    def _apply_overlay_mode(self):
        if not self.overlay_active:
            return
        if self.mode in ("candle", "candle_free", "fib", "line", "segment", "arrow",
                         "circle", "rect", "eraser"):
            self._set_clickthrough(False)
            if self.overlay_panel:
                self.overlay_panel.redraw()
        else:
            self.overlay_photo = None
            self._set_clickthrough(True)
            if self.overlay_panel:
                self.overlay_panel.redraw()

    def _hide_self_for_grab(self):
        try:
            self.root.attributes("-alpha", 0.0)
            self.root.update_idletasks()
            self.root.update()
            time.sleep(0.06)
            self.root.update()
        except tk.TclError:
            pass

    def _show_self_after_grab(self):
        try:
            self.root.attributes("-alpha", 1.0)
            self.root.update_idletasks()
        except tk.TclError:
            pass

    def _capture_bg(self):
        p = self.overlay_panel
        if p is None:
            return
        if not HAS_PIL:
            self.hint("Instale o Pillow para desenhar: pip install pillow")
            return

        self._hide_self_for_grab()

        try:
            rect = self._overlay_monitor_rect
            try:
                img = ImageGrab.grab(all_screens=True)
            except TypeError:
                img = ImageGrab.grab()
            if rect:
                try:
                    import ctypes
                    vx = ctypes.windll.user32.GetSystemMetrics(76)
                    vy = ctypes.windll.user32.GetSystemMetrics(77)
                    l, t, w, h = rect
                    img = img.crop((l - vx, t - vy,
                                    l - vx + w, t - vy + h))
                except Exception:
                    pass
            th = self.toolbar.winfo_height()
            w, h = img.size
            if 0 < th < h:
                img = img.crop((0, th, w, h))
            self.overlay_photo = ImageTk.PhotoImage(img)
        except Exception:
            self.overlay_photo = None
        finally:
            self._show_self_after_grab()

    def _on_escape(self, _e):
        if self.lousa_estado != 0:
            self.lousa_estado = 0
            self._atualizar_icone_lousa()
            self._aplicar_lousa(None)
            self.hint("Lousa desativada.")
            return
        if self.overlay_active:
            self.exit_overlay()
        else:
            self.deselect()

    # ============================================================
    # Layout
    # ============================================================
    
    def _criar_canvas_livre(self):
        """Cria o painel interno usado pelo Sobrepor. Não aparece
        no modo janelinha — só dentro do Sobrepor."""
        for child in self.workspace.winfo_children():
            child.destroy()
        self.panels = []

        frame = tk.Frame(self.workspace, bg="#101419")
        # Não empacota — fica invisível até o Sobrepor usar

        p = Panel(self, frame, 0, source=True)
        self.panels.append(p)

        c = p.canvas
        c.bind("<ButtonPress-1>", lambda e, p=p: self.press(e, p))
        c.bind("<B1-Motion>", lambda e, p=p: self.motion(e, p))
        c.bind("<ButtonRelease-1>", lambda e, p=p: self.release(e, p))
        c.bind("<Motion>", lambda e, p=p: self.hover(e, p))
        c.bind("<ButtonPress-3>",
               lambda e, p=p: self.release_mode_atalho(p))

        self.active_panel = p
        self.root.update_idletasks()
        p.redraw()

    def _after_layout_change(self):
        self.root.update_idletasks()
        self.redraw_all()

    def release_mode_atalho(self, p):
        if self._shift_pressionado:
            self._atalho_candle_traco()
            return
        # Botão direito: volta ao modo Mover/Selecionar.
        self.release_mode()

    def set_mode(self, mode):
        self._forcar_redraw_global()  # Atualiza cores caso template mudou                
        if mode == "undo":
            self.undo()
            return
        if mode == "clear":
            self._snapshot()
            # Limpa os desenhos visíveis, inclusive candles da lousa.
            paineis = list(self.panels)
            if self.lousa_panel is not None:
                paineis.append(self.lousa_panel)
            paineis.extend(p for _, p in getattr(self, "_overlay_paineis", []))
            vistos = set()
            for p in paineis:
                if id(p) in vistos:
                    continue
                vistos.add(id(p))
                p.candles.clear()
                p.redraw()
            self.fib = None
            self.lines.clear()
            self.segments.clear()
            self.arrows.clear()
            self.circles.clear()
            self.rects.clear()
            self.texts.clear()
            self.highlights.clear()
            # Apaga também o rascunho guardado quando a lousa está fechada.
            self.lousa_desenho = {
                "candles": [], "lines": [], "segments": [], "arrows": [],
                "circles": [], "rects": [], "texts": [], "highlights": [],
                "fib": None, "zoom": 1.0, "zoom_y": 1.0,
                "offset_x": 0.0, "offset_y": 0.0,
            }
            self.selected = None
            self.redraw_all()
            return

        self.mode = mode
        self._atualizar_label_modo(f"Modo: {self.mode_names[mode]}")
        self._atualizar_destaque()

        # Fase 3c: ferramentas ativam o Sobrepor automaticamente
        if self.lousa_estado != 0:
            # Na lousa não ativa Sobrepor
            return
        # Mover mantém a sobreposição e os hooks ativos: cliques sobre
        # desenhos são capturados para seleção; áreas vazias passam ao Windows.
        if not self.overlay_active:
            self.open_overlay()
        self._apply_overlay_mode()
        if mode == "clear":
            self._snapshot()
            for p in self.panels:
                p.clear()
            if self.overlay_panel is not None:
                self.overlay_panel.clear()
            self.fib = None
            self.lines.clear()
            self.segments.clear()
            self.arrows.clear()
            self.circles.clear()
            self.rects.clear()
            self.selected = None
            self.redraw_all()
            return
        self.mode = mode
        self._atualizar_label_modo(f"Modo: {self.mode_names[mode]}")
        self._atualizar_destaque()
        self._apply_overlay_mode()

    def _abrir_input_texto(self, painel, mx, my):
        """Abre uma caixinha de digitar no ponto clicado. Enter confirma, Esc cancela."""
        win = tk.Toplevel(self.root)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.configure(bg="#0b0e11")

        try:
            px = painel.canvas.winfo_rootx() + int(
                (mx - painel.offset_x) * painel.zoom)
            py = painel.canvas.winfo_rooty() + int(
                (my - painel.offset_y) * painel.zoom_y)
            win.geometry(f"+{px}+{py}")
        except Exception:
            win.geometry("+300+300")

        frame = tk.Frame(win, bg="#0b0e11",
                         highlightthickness=1,
                         highlightbackground="#3d6082")
        frame.pack(padx=1, pady=1)

        entry = tk.Entry(frame, bg="#20272f", fg="#e7edf4",
                         insertbackground="#e7edf4",
                         relief="flat", font=("Segoe UI", 11),
                         width=24)
        entry.pack(padx=6, pady=6)
        entry.focus_set()

        def confirmar(event=None):
            texto = entry.get().strip()
            win.destroy()
            if texto:
                self._snapshot()
                if self.segment_color == SEG_COLORS[0][0]:
                    cor_texto = (COR_TEXTO_ESCURO
                                 if self.tema_nome == "escuro"
                                 else COR_TEXTO_CLARO)
                else:
                    cor_texto = self.segment_color
                novo = TextShape(x=mx, y=my, texto=texto,
                                 color=cor_texto,
                                 espessura=self.espessura,
                                 tamanho_fonte=self.tamanho_fonte_texto)
                novo.painel = painel
                self.texts.append(novo)
                self.selected = ("text", novo)
                self.redraw_all()

        def cancelar(event=None):
            win.destroy()

        entry.bind("<Return>", confirmar)
        entry.bind("<Escape>", cancelar)
        win.bind("<FocusOut>", lambda e: cancelar())   
    
    def release_mode(self):
        self.drag = None
        self._provisional_candle = None
        self.wick_drag = None
        self.fib_drag = None
        self.line_drag = None
        self.text_drag = None
        self.highlight_drag = None
        self.seg_drag = None
        self.fib_start = None
        self.line_start = None
        self.seg_start = None
        self.shape_start = None
        self.is_erasing = False

        if self._provisional_candle is not None:
            p, candle = self._provisional_candle
            if candle in p.candles:
                p.candles.remove(candle)
            self._provisional_candle = None
            self.recalculate()
        if self.mode != "move":
            self.set_mode("move")

    def panel_role(self, p):
        """No Live só existe 1 painel — sempre 'free'."""
        return "free"

    def erase_at(self, p, x, y):
        mx, my = p.tela_para_mundo(x, y)
        erased = False

        for c in list(p.candles):
            if abs(mx - c.x) <= c.width and c.high_y - 5 <= my <= c.low_y + 5:
                p.candles.remove(c)
                erased = True

        for seg in list(self.segments):
            if self._dist_point_segment(mx, my, seg.x0, seg.y0,
                                        seg.x1, seg.y1) <= 10:
                self.segments.remove(seg)
                erased = True

        for arrow in list(self.arrows):
            if self._dist_point_segment(mx, my, arrow.x0, arrow.y0,
                                        arrow.x1, arrow.y1) <= 10:
                self.arrows.remove(arrow)
                erased = True

        for circle in list(self.circles):
            cx0, cx1 = min(circle.x0, circle.x1), max(circle.x0, circle.x1)
            cy0, cy1 = min(circle.y0, circle.y1), max(circle.y0, circle.y1)
            if cx0 - 5 <= mx <= cx1 + 5 and cy0 - 5 <= my <= cy1 + 5:
                self.circles.remove(circle)
                erased = True

        for rect in list(self.rects):
            rx0, rx1 = min(rect.x0, rect.x1), max(rect.x0, rect.x1)
            ry0, ry1 = min(rect.y0, rect.y1), max(rect.y0, rect.y1)
            if rx0 - 5 <= mx <= rx1 + 5 and ry0 - 5 <= my <= ry1 + 5:
                self.rects.remove(rect)
                erased = True

        for txt in list(self.texts):
            base = {1: 12, 2: 14, 3: 18, 5: 26}.get(txt.espessura, 14)
            fonte = max(8, int(base))
            largura_aprox = len(txt.texto) * fonte * 0.55
            altura_aprox = fonte * 1.3
            if (txt.x - 5 <= mx <= txt.x + largura_aprox + 5
                    and txt.y - 5 <= my <= txt.y + altura_aprox + 5):
                self.texts.remove(txt)
                erased = True

        for hl in list(self.highlights):
            hx0, hx1 = min(hl.x0, hl.x1), max(hl.x0, hl.x1)
            hy0, hy1 = min(hl.y0, hl.y1), max(hl.y0, hl.y1)
            if hx0 - 5 <= mx <= hx1 + 5 and hy0 - 5 <= my <= hy1 + 5:
                self.highlights.remove(hl)
                erased = True                

        for line in list(self.lines):
            if abs(my - line.y) <= 8:
                self.lines.remove(line)
                erased = True

        if self.fib:
            fx0, fx1 = min(self.fib.x0, self.fib.x1), max(self.fib.x0,
                                                          self.fib.x1)
            fy0, fy1 = min(self.fib.y0, self.fib.y1), max(self.fib.y0,
                                                          self.fib.y1)
            if fx0 <= mx <= fx1 and fy0 <= my <= fy1:
                self.fib = None
                erased = True

        if erased:
            self.recalculate()
            self.redraw_all()

    def deselect(self):
        if self.selected:
            self.selected = None
            self.redraw_all()

    def delete_selected(self):
        sel = self.selected
        if not sel:
            return
        self._snapshot()
        if sel[0] == "candle":
            p, cd = sel[1], sel[2]
            if cd in p.candles:
                p.candles.remove(cd)
            self.recalculate()
        elif sel[0] == "fib":
            self.fib = None
        elif sel[0] == "line":
            if sel[1] in self.lines:
                self.lines.remove(sel[1])
        elif sel[0] == "segment":
            if sel[1] in self.segments:
                self.segments.remove(sel[1])
        elif sel[0] == "arrow":
            if sel[1] in self.arrows:
                self.arrows.remove(sel[1])
        elif sel[0] == "circle":
            if sel[1] in self.circles:
                self.circles.remove(sel[1])
        elif sel[0] == "rect":
            if sel[1] in self.rects:
                self.rects.remove(sel[1])
        elif sel[0] == "text":
            if sel[1] in self.texts:
                self.texts.remove(sel[1])
        elif sel[0] == "highlight":
            if sel[1] in self.highlights:
                self.highlights.remove(sel[1])                
        self.selected = None
        self.redraw_all()

    @staticmethod
    def _dist_point_segment(px, py, x0, y0, x1, y1):
        dx, dy = x1 - x0, y1 - y0
        if dx == 0 and dy == 0:
            return ((px - x0) ** 2 + (py - y0) ** 2) ** 0.5
        t = ((px - x0) * dx + (py - y0) * dy) / (dx * dx + dy * dy)
        t = max(0.0, min(1.0, t))
        return ((px - (x0 + t * dx)) ** 2 + (py - (y0 + t * dy)) ** 2) ** 0.5

    def hit_test(self, p, x, y):
        mx, my = p.tela_para_mundo(x, y)
        fib = self.fib
        if fib:
            left, right = min(fib.x0, fib.x1), max(fib.x0, fib.x1)
            top, bottom = min(fib.y0, fib.y1), max(fib.y0, fib.y1)
            r = 9
            corners = [
                ("fib_tl", fib.x0, fib.y0),
                ("fib_tr", fib.x1, fib.y0),
                ("fib_bl", fib.x0, fib.y1),
                ("fib_br", fib.x1, fib.y1),
            ]
            for kind, cx, cy in corners:
                if abs(mx - cx) <= r and abs(my - cy) <= r:
                    return (kind,)
            if left <= mx <= right and abs(my - fib.y0) <= 6:
                return ("fib_top",)
            if left <= mx <= right and abs(my - fib.y1) <= 6:
                return ("fib_bottom",)
            if left <= mx <= right and top <= my <= bottom:
                return ("fib",)
        for line in self.lines:
            if abs(my - line.y) <= 5:
                return ("line", line)
        for seg in self.segments:
            if self._dist_point_segment(mx, my, seg.x0, seg.y0,
                                        seg.x1, seg.y1) <= 5:
                return ("segment", seg)
        for arrow in self.arrows:
            if self._dist_point_segment(mx, my, arrow.x0, arrow.y0,
                                        arrow.x1, arrow.y1) <= 5:
                return ("arrow", arrow)
        for circle in self.circles:
            cx0, cx1 = min(circle.x0, circle.x1), max(circle.x0, circle.x1)
            cy0, cy1 = min(circle.y0, circle.y1), max(circle.y0, circle.y1)
            if cx0 <= mx <= cx1 and cy0 <= my <= cy1:
                return ("circle", circle)
        for rect in self.rects:
            rx0, rx1 = min(rect.x0, rect.x1), max(rect.x0, rect.x1)
            ry0, ry1 = min(rect.y0, rect.y1), max(rect.y0, rect.y1)
            if rx0 <= mx <= rx1 and ry0 <= my <= ry1:
                return ("rect", rect)

        for hl in reversed(self.highlights):
            hx0, hx1 = min(hl.x0, hl.x1), max(hl.x0, hl.x1)
            hy0, hy1 = min(hl.y0, hl.y1), max(hl.y0, hl.y1)
            if hx0 <= mx <= hx1 and hy0 <= my <= hy1:
                return ("highlight", hl)                

        for txt in reversed(self.texts):
            base = {1: 12, 2: 14, 3: 18, 5: 26}.get(txt.espessura, 14)
            fonte = max(8, int(base))
            largura_aprox = len(txt.texto) * fonte * 0.55
            altura_aprox = fonte * 1.3
            if (txt.x <= mx <= txt.x + largura_aprox
                    and txt.y <= my <= txt.y + altura_aprox):
                return ("text", txt)                

        if self.panel_role(p) in ("free", "source"):
            candle = self.near_candle(p, mx, my)
            if candle:
                if abs(my - candle.high_y) <= 9:
                    return ("wick", p, candle, "high")
                if abs(my - candle.low_y) <= 9:
                    return ("wick", p, candle, "low")
                return ("body", p, candle)
        return None

    def select_from_hit(self, hit):
        kind = hit[0]
        if kind in ("body", "wick"):
            self.selected = ("candle", hit[1], hit[2])
        elif kind in ("fib", "fib_top", "fib_bottom"):
            self.selected = ("fib",)
        elif kind == "line":
            self.selected = ("line", hit[1])
        elif kind == "segment":
            self.selected = ("segment", hit[1])
        elif kind == "arrow":
            self.selected = ("arrow", hit[1])
        elif kind == "circle":
            self.selected = ("circle", hit[1])
        elif kind == "rect":
            self.selected = ("rect", hit[1])
        elif kind == "highlight":
            self.selected = ("highlight", hit[1])            
        elif kind == "text":
            self.selected = ("text", hit[1])            

    def press(self, e, p):
        self.active_panel = p
        role = self.panel_role(p)
        mx, my = p.tela_para_mundo(e.x, e.y)

        # Modo Mover: só seleciona/deseleciona objetos existentes
        if self.mode == "move":
            hit = self.hit_test(p, e.x, e.y)
            if hit is None:
                self.deselect()
                return
            self.select_from_hit(hit)
            kind = hit[0]
            if kind == "text":
                self.text_drag = (hit[1], mx - hit[1].x, my - hit[1].y)
                self.redraw_all()
                return
            if kind == "highlight":
                hl = hit[1]
                self.highlight_drag = (hl,
                                       mx - hl.x0, my - hl.y0,
                                       mx - hl.x1, my - hl.y1)
                self.redraw_all()
                return
            if kind in ("fib_tl", "fib_tr", "fib_bl", "fib_br"):
                self.fib_drag = (kind,)
            elif kind == "fib_top":
                self.fib_drag = ("top",)
            elif kind == "fib_bottom":
                self.fib_drag = ("bottom",)
            elif kind == "fib":
                fib = self.fib
                self.fib_drag = ("move", mx - fib.x0, my - fib.y0,
                                 fib.x1 - fib.x0, fib.y1 - fib.y0)
            elif kind == "line":
                self.line_drag = (hit[1], my - hit[1].y)
            elif kind == "segment":
                seg = hit[1]
                self.seg_drag = (seg, mx - seg.x0, my - seg.y0,
                                 mx - seg.x1, my - seg.y1)
            elif kind == "wick":
                self.wick_drag = (hit[1], hit[2], hit[3])
            elif kind == "body":
                c = hit[2]
                self.drag = {"panel": p, "candle": c,
                             "dx": mx - c.x, "dy": my - c.open_y}
            self.redraw_all()
            return
        

        if self.mode == "eraser":
            self._snapshot()
            self.is_erasing = True
            self.erase_at(p, e.x, e.y)
            return

        if self.mode in ("candle", "candle_free"):
            if role in ("projected", "excluded"):
                src = self.source_panel()
                if src:
                    self.hint(f"Desenhe apenas no painel FONTE "
                              f"({src.timeframe} MIN)")
                else:
                    self.hint("Desenhe no painel FONTE para criar candles.")
                return
            self._snapshot()
            if self.lousa_estado != 0 and p is self.lousa_panel:
                novo_x = mx
                novo_open = my
            elif self.mode == "candle" and p.candles:
                ultimo_criado = p.candles[-1]
                mais_a_direita = max(p.candles, key=lambda c: c.x)
                # Mantém 7 px de espaço livre entre os corpos. A distância
                # entre centros acompanha a largura de ambos os candles.
                espaco_corpos = 7.0
                distancia_centros = (
                    mais_a_direita.width + float(self.candle_tamanho)
                ) / 2.0 + espaco_corpos
                novo_x = mais_a_direita.x + distancia_centros
                novo_open = ultimo_criado.close_y
            else:
                novo_x = mx
                novo_open = my
            candle = Candle(
                open_y=novo_open, close_y=novo_open,
                high_y=novo_open - 12, low_y=novo_open + 12, x=novo_x,
                width=float(self.candle_tamanho)
            )
            candle.painel = p            
            p.candles.append(candle)
            self._provisional_candle = (p, candle)
            self.drag = {"panel": p, "candle": candle,
                         "start": novo_open, "now": novo_open,
                         "stage": "body"}
            self.selected = ("candle", p, candle)
            self.recalculate()
            return        
        if self.mode == "fib":
            self._snapshot()
            self.fib_start = (mx, my, p)
            return
        if self.mode == "line":
            self._snapshot()
            self.line_start = (my, p)
            return
        if self.mode == "segment":
            self._snapshot()
            self.seg_start = (mx, my, p)
            return
        if self.mode in ("arrow", "circle", "rect", "highlight"):
            self._snapshot()
            self.shape_start = (mx, my, p)
            return
            
        if self.mode == "text":
            self._abrir_input_texto(p, mx, my)
            return

        hit = self.hit_test(p, e.x, e.y)
        if hit is None:
            self.deselect()
            return

        self.select_from_hit(hit)
        kind = hit[0]

        if kind in ("fib_tl", "fib_tr", "fib_bl", "fib_br", "fib_top",
                    "fib_bottom", "fib", "line", "segment", "wick", "body"):
            self._snapshot()

        if kind == "text":
            self.text_drag = (hit[1], mx - hit[1].x, my - hit[1].y)
            self.redraw_all()
            return

        if kind == "highlight":
            hl = hit[1]
            self.highlight_drag = (hl,
                                   mx - hl.x0, my - hl.y0,
                                   mx - hl.x1, my - hl.y1)
            self.redraw_all()
            return            

        if kind in ("fib_tl", "fib_tr", "fib_bl", "fib_br"):
            self.fib_drag = (kind,)
        elif kind == "fib_top":
            self.fib_drag = ("top",)
        elif kind == "fib_bottom":
            self.fib_drag = ("bottom",)
        elif kind == "fib":
            fib = self.fib
            self.fib_drag = ("move", mx - fib.x0, my - fib.y0,
                             fib.x1 - fib.x0, fib.y1 - fib.y0)
        elif kind == "line":
            self.line_drag = (hit[1], my - hit[1].y)
        elif kind == "segment":
            seg = hit[1]
            self.seg_drag = (seg, mx - seg.x0, my - seg.y0,
                             mx - seg.x1, my - seg.y1)
        elif kind == "wick":
            self.wick_drag = (hit[1], hit[2], hit[3])
        elif kind == "body":
            c = hit[2]
            self.drag = {"panel": p, "candle": c,
                         "dx": mx - c.x, "dy": my - c.open_y}
        self.redraw_all()

    def motion(self, e, p):
        if self.mode == "eraser" and self.is_erasing:
            self.erase_at(p, e.x, e.y)
            return

        mx, my = p.tela_para_mundo(e.x, e.y)

        if self.drag and self.drag.get("stage") == "body":
            c = self.drag["candle"]
            c.close_y = my
            c.high_y = min(c.high_y, c.body_top - 2)
            c.low_y = max(c.low_y, c.body_bottom + 2)
            self.drag["now"] = my
            self.recalculate()
            self.redraw_all()
            return

        if self.seg_start is not None:
            x0, y0, _painel_orig = self.seg_start
            p.redraw()
            x0_t, y0_t = p.mundo_para_tela(x0, y0)
            p.canvas.create_line(x0_t, y0_t, e.x, e.y,
                                 fill=self.segment_color,
                                 width=max(1, int(self.espessura * p.zoom)),
                                 dash=(5, 3))
            return

        if self.shape_start is not None:
            x0, y0, _painel_orig = self.shape_start
            p.redraw()
            x0_t, y0_t = p.mundo_para_tela(x0, y0)
            wzoom = max(1, int(self.espessura * p.zoom))
            if self.mode == "arrow":
                p.canvas.create_line(x0_t, y0_t, e.x, e.y,
                                     fill=self.segment_color,
                                     width=wzoom,
                                     arrow=tk.LAST)
            elif self.mode == "circle":
                p.canvas.create_oval(x0_t, y0_t, e.x, e.y,
                                     outline=self.segment_color,
                                     width=wzoom)
            elif self.mode == "rect":
                p.canvas.create_rectangle(x0_t, y0_t, e.x, e.y,
                                          outline=self.segment_color,
                                          width=wzoom)
            elif self.mode == "highlight":
                p.canvas.create_rectangle(x0_t, y0_t, e.x, e.y,
                                          fill=self.segment_color,
                                          outline="",
                                          stipple="gray50")                                          
            return

        if self.drag and "candle" in self.drag:
            c = self.drag["candle"]
            dx = mx - (c.x + self.drag["dx"])
            dy = my - (c.open_y + self.drag["dy"])
            c.x += dx
            c.open_y += dy
            c.close_y += dy
            c.high_y += dy
            c.low_y += dy
            p.redraw()
            self.recalculate()
            return

        if self.seg_drag:
            seg, ox0, oy0, ox1, oy1 = self.seg_drag
            seg.x0 = mx - ox0
            seg.y0 = my - oy0
            seg.x1 = mx - ox1
            seg.y1 = my - oy1
            self.redraw_all()
            return

        if self.fib_drag:
            kind = self.fib_drag[0]
            fib = self.fib
            if fib is None:
                self.fib_drag = None
                return
            if kind == "fib_tl":
                fib.x0, fib.y0 = mx, my
            elif kind == "fib_tr":
                fib.x1, fib.y0 = mx, my
            elif kind == "fib_bl":
                fib.x0, fib.y1 = mx, my
            elif kind == "fib_br":
                fib.x1, fib.y1 = mx, my
            elif kind == "top":
                fib.y0 = my
            elif kind == "bottom":
                fib.y1 = my
            else:
                _, offx, offy, width, height = self.fib_drag
                fib.x0 = mx - offx
                fib.y0 = my - offy
                fib.x1 = fib.x0 + width
                fib.y1 = fib.y0 + height
            self.redraw_all()
            return

        if self.line_drag:
            line, off = self.line_drag
            line.y = my - off
            self.redraw_all()
            return

        if self.text_drag:
            txt, offx, offy = self.text_drag
            txt.x = mx - offx
            txt.y = my - offy
            self.redraw_all()
            return

        if self.highlight_drag:
            hl, ox0, oy0, ox1, oy1 = self.highlight_drag
            hl.x0 = mx - ox0
            hl.y0 = my - oy0
            hl.x1 = mx - ox1
            hl.y1 = my - oy1
            self.redraw_all()
            return            

        if self.wick_drag:
            p, c, which = self.wick_drag
            if which == "high":
                c.high_y = min(my, c.body_top - 2)
            else:
                c.low_y = max(my, c.body_bottom + 2)
            p.redraw()
            self.recalculate()

    def release(self, e, p):
        if self.mode == "eraser":
            self.is_erasing = False
            return

        mx, my = p.tela_para_mundo(e.x, e.y)

        if self.drag and self.drag.get("stage") == "body":
            candle = self.drag.get("candle")
            if candle is not None:
                corpo_pequeno = abs(candle.close_y - candle.open_y) < 4
                altura_total = abs(candle.low_y - candle.high_y)
                eh_doji_intencional = corpo_pequeno and altura_total > 20

                if corpo_pequeno and not eh_doji_intencional:
                    if candle in p.candles:
                        p.candles.remove(candle)
                    if self.selected and self.selected[0] == "candle" \
                            and self.selected[2] is candle:
                        self.selected = None
                else:
                    self.selected = ("candle", p, candle)
            self._provisional_candle = None
            self.recalculate()
            self.redraw_all()

        if self.seg_start is not None:
            x0, y0, painel_orig = self.seg_start
            if ((mx - x0) ** 2 + (my - y0) ** 2) ** 0.5 >= 4:
                seg = SegLine(x0, y0, mx, my, self.segment_color,
                              espessura=self.espessura)
                seg.painel = painel_orig
                self.segments.append(seg)
                self.selected = ("segment", seg)
                self.redraw_all()
            self.seg_start = None

        if self.shape_start is not None:
            x0, y0, painel_orig = self.shape_start
            if ((mx - x0) ** 2 + (my - y0) ** 2) ** 0.5 >= 4:
                if self.mode == "arrow":
                    shape = ArrowShape(x0, y0, mx, my, self.segment_color,
                                       espessura=self.espessura)
                    shape.painel = painel_orig
                    self.arrows.append(shape)
                    self.selected = ("arrow", shape)
                elif self.mode == "circle":
                    shape = CircleShape(x0, y0, mx, my, self.segment_color,
                                        espessura=self.espessura)
                    shape.painel = painel_orig
                    self.circles.append(shape)
                    self.selected = ("circle", shape)
                elif self.mode == "rect":
                    shape = RectShape(x0, y0, mx, my, self.segment_color,
                                      espessura=self.espessura)
                    shape.painel = painel_orig
                    self.rects.append(shape)
                    self.selected = ("rect", shape)
                elif self.mode == "highlight":
                    shape = HighlightShape(x0, y0, mx, my,
                                           self.segment_color)
                    shape.painel = painel_orig
                    self.highlights.append(shape)
                    self.selected = ("highlight", shape)
                self.redraw_all()
            self.shape_start = None

        if self.fib_start is not None:
            x0, y0, painel_orig = self.fib_start
            if ((mx - x0) ** 2 + (my - y0) ** 2) ** 0.5 >= 4:
                self.fib = Fib(x0, y0, mx, my, espessura=self.espessura,
                               color=self.segment_color,
                               estilo=self.fib_estilo)
                self.fib.painel = painel_orig
                self.selected = ("fib",)
                self.redraw_all()
            self.fib_start = None

        if self.line_start is not None:
            _y_orig, painel_orig = self.line_start
            nova_linha = HLine(y=my, color=self.segment_color,
                               espessura=self.espessura)
            nova_linha.painel = painel_orig
            self.lines.append(nova_linha)
            self.selected = ("line", nova_linha)
            self.redraw_all()
            self.line_start = None

        self.drag = None
        self._provisional_candle = None
        self.wick_drag = None
        self.fib_drag = None
        self.line_drag = None
        self.seg_drag = None
        self.text_drag = None
        self.highlight_drag = None        

    def near_candle(self, p, mx, my):
        for c in reversed(p.candles):
            if abs(mx - c.x) <= c.width + 8 and c.high_y - 10 <= my <= c.low_y + 10:
                return c
        return None

    def hover(self, e, p):
        if self.mode == "eraser":
            p.canvas.config(cursor="circle")
            return
        if self.mode != "move":
            p.canvas.config(cursor="crosshair")
            return
        mx, my = p.tela_para_mundo(e.x, e.y)
        if self.panel_role(p) in ("projected", "excluded"):
            hit = None
            fib = self.fib
            if fib:
                left, right = min(fib.x0, fib.x1), max(fib.x0, fib.x1)
                top, bottom = min(fib.y0, fib.y1), max(fib.y0, fib.y1)
                corners = ((fib.x0, fib.y0), (fib.x1, fib.y0),
                           (fib.x0, fib.y1), (fib.x1, fib.y1))
                if any(abs(mx-cx) <= 9 and abs(my-cy) <= 9
                       for cx, cy in corners):
                    hit = ("fib_corner",)
                elif left <= mx <= right and (abs(my-fib.y0) <= 6
                                              or abs(my-fib.y1) <= 6):
                    hit = ("fib_edge",)
                elif left <= mx <= right and top <= my <= bottom:
                    hit = ("fib",)
            else:
                for line in self.lines:
                    if abs(my - line.y) <= 5:
                        hit = ("line",)
                        break
        else:
            hit = self.hit_test(p, e.x, e.y)

        if hit is None:
            p.canvas.config(cursor="arrow")
        elif hit[0] in ("fib_tl", "fib_tr", "fib_bl", "fib_br", "fib_corner"):
            p.canvas.config(cursor="sizing")
        elif hit[0] in ("wick", "fib_top", "fib_bottom", "fib_edge"):
            p.canvas.config(cursor="sb_v_double_arrow")
        else:
            p.canvas.config(cursor="fleur")

    def redraw_all(self):
        if self.lousa_estado != 0 and self.lousa_panel is not None:
            self.lousa_panel.redraw()
            return
        for p in self.panels:
            p.redraw()
        for _, p in getattr(self, "_overlay_paineis", []):
            p.redraw()

    def recalculate(self):
        """No Live não há projeção entre painéis — não faz nada."""
        return

    def undo(self):
        if not self.historico:
            return
        estado = self.historico.pop()
        self._restaurar_snapshot(estado)

# ============================================================
# Sessão persistente (token salvo no PC do usuário)
# ============================================================
SESSION_FILE = os.path.join(os.path.expanduser("~"),
                            ".projetor_candles_session.json")


def salvar_sessao(email, access_token, refresh_token):
    try:
        dados = {
            "email": email,
            "access_token": access_token,
            "refresh_token": refresh_token,
            "salvo_em": datetime.now().isoformat(),
        }
        with open(SESSION_FILE, "w", encoding="utf-8") as f:
            json.dump(dados, f)
        return True
    except Exception as e:
        print(f"Erro ao salvar sessão: {e}")
        return False


def ler_sessao():
    try:
        if not os.path.exists(SESSION_FILE):
            return None
        with open(SESSION_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def apagar_sessao():
    try:
        if os.path.exists(SESSION_FILE):
            os.remove(SESSION_FILE)
    except Exception:
        pass


def validar_sessao_salva():
    dados = ler_sessao()
    if not dados:
        return (False, None)

    email = dados.get("email")
    access_token = dados.get("access_token")
    refresh_token = dados.get("refresh_token")

    if not email or not access_token:
        return (False, None)

    try:
        cliente = create_client(SUPABASE_URL, SUPABASE_ANON_KEY)
        cliente.auth.set_session(access_token, refresh_token)

        user = cliente.auth.get_user()
        if not user or not user.user:
            return (False, None)

        perfil = cliente.table("profiles").select("*").eq(
            "id", user.user.id).execute()

        if not perfil.data:
            return (False, None)

        info = perfil.data[0]
        if not info.get("ativo", True):
            return (False, None)

        expira_em = info.get("data_expiracao")
        if expira_em:
            from datetime import timezone
            expira_dt = datetime.fromisoformat(
                expira_em.replace("Z", "+00:00"))
            agora = datetime.now(timezone.utc)
            if agora > expira_dt:
                return (False, None)

        return (True, email)
    except Exception as e:
        print(f"Sessão inválida: {e}")
        return (False, None)


class LoginWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("DrawCandles — Entrar")
        self.root.geometry("400x480")
        self.root.resizable(False, False)
        self.root.configure(bg="#0b0e11")

        try:
            if os.path.exists(caminho_recurso(ICONE_ARQUIVO)):
                img = Image.open(caminho_recurso(ICONE_ARQUIVO)).convert("RGBA")
                img = img.resize((80, 80), Image.LANCZOS)
                self._icone = ImageTk.PhotoImage(img)
                self.root.iconphoto(True, self._icone)
        except Exception:
            pass

        self.usuario_logado = False
        self.email_usuario = None

        frame = tk.Frame(self.root, bg="#0b0e11")
        frame.pack(expand=True, fill="both", padx=40, pady=40)

        tk.Label(frame, text="🕯", bg="#0b0e11", fg="#e8b93e",
                 font=("Segoe UI Emoji", 48)).pack(pady=(0, 8))

        tk.Label(frame, text="DrawCandles", bg="#0b0e11",
                 fg="#e7edf4", font=("Segoe UI", 18, "bold")).pack()

        tk.Label(frame, text="Faça login para continuar",
                 bg="#0b0e11", fg="#91a4b7",
                 font=("Segoe UI", 10)).pack(pady=(4, 24))

        tk.Label(frame, text="E-mail", bg="#0b0e11", fg="#dfe7ef",
                 font=("Segoe UI", 10), anchor="w").pack(fill="x")
        self.entry_email = tk.Entry(frame, bg="#20272f", fg="#e7edf4",
                                    insertbackground="#e7edf4",
                                    relief="flat", font=("Segoe UI", 11))
        self.entry_email.pack(fill="x", ipady=6, pady=(4, 12))

        tk.Label(frame, text="Senha", bg="#0b0e11", fg="#dfe7ef",
                 font=("Segoe UI", 10), anchor="w").pack(fill="x")
        self.entry_senha = tk.Entry(frame, bg="#20272f", fg="#e7edf4",
                                    insertbackground="#e7edf4",
                                    relief="flat", font=("Segoe UI", 11),
                                    show="●")
        self.entry_senha.pack(fill="x", ipady=6, pady=(4, 20))

        self.btn_entrar = tk.Button(
            frame, text="Entrar", command=self.fazer_login,
            bg="#3d6082", fg="#e7edf4", activebackground="#4a7ba8",
            activeforeground="white", relief="flat",
            font=("Segoe UI", 11, "bold"), cursor="hand2")
        self.btn_entrar.pack(fill="x", ipady=8)

        self.lbl_status = tk.Label(frame, text="", bg="#0b0e11",
                                   fg="#ff5c68", font=("Segoe UI", 9),
                                   wraplength=320, justify="center")
        self.lbl_status.pack(pady=(12, 0))

        tk.Label(frame, text="Não tem conta? Crie em drawcandles.com",
                 bg="#0b0e11", fg="#91a4b7",
                 font=("Segoe UI", 8)).pack(side="bottom", pady=(20, 0))

        self.root.bind("<Return>", lambda e: self.fazer_login())

    def fazer_login(self):
        email = self.entry_email.get().strip()
        senha = self.entry_senha.get()

        if not email or not senha:
            self.lbl_status.config(text="Preencha e-mail e senha.")
            self.root.update_idletasks()
            return

        self.btn_entrar.config(text="Entrando...", state="disabled")
        self.lbl_status.config(text="")
        self.root.update_idletasks()

        try:
            cliente = create_client(SUPABASE_URL, SUPABASE_ANON_KEY)
            resposta = cliente.auth.sign_in_with_password({
                "email": email,
                "password": senha
            })

            perfil = cliente.table("profiles").select("*").eq(
                "id", resposta.user.id).execute()

            if not perfil.data:
                self.lbl_status.config(text="Perfil não encontrado.")
                self.root.update_idletasks()
                self.btn_entrar.config(text="Entrar", state="normal")
                return

            dados = perfil.data[0]

            if not dados.get("ativo", True):
                self.lbl_status.config(text="Sua conta está desativada.")
                self.root.update_idletasks()
                self.btn_entrar.config(text="Entrar", state="normal")
                return

            expira_em = dados.get("data_expiracao")
            if expira_em:
                from datetime import timezone
                expira_dt = datetime.fromisoformat(
                    expira_em.replace("Z", "+00:00"))
                agora = datetime.now(timezone.utc)
                if agora > expira_dt:
                    self.lbl_status.config(
                        text="Seu período de teste expirou.\n"
                             "Visite drawcandles.com para continuar.")
                    self.root.update_idletasks()
                    self.btn_entrar.config(text="Entrar", state="normal")
                    return

            try:
                session = cliente.auth.get_session()
                if session:
                    salvar_sessao(email,
                                  session.access_token,
                                  session.refresh_token)
            except Exception as e:
                print(f"Aviso: não consegui salvar sessão: {e}")

            self.usuario_logado = True
            self.email_usuario = email
            self.root.destroy()

        except Exception as e:
            mensagem = str(e)
            if "Invalid login" in mensagem or "invalid" in mensagem.lower():
                self.lbl_status.config(text="E-mail ou senha incorretos.")
            else:
                self.lbl_status.config(text=f"Erro: {mensagem[:80]}")
            self.root.update_idletasks()
            self.btn_entrar.config(text="Entrar", state="normal")

    def run(self):
        self.root.mainloop()
        return self.usuario_logado


if __name__ == "__main__":
    sessao_valida, email_sessao = validar_sessao_salva()

    if sessao_valida:
        print(f"Sessão válida para {email_sessao}. Pulando login.")
        root = tk.Tk()
        app = CandleProjector(root)
        root.mainloop()
    else:
        apagar_sessao()
        login = LoginWindow()
        if login.run():
            root = tk.Tk()
            app = CandleProjector(root)
            root.mainloop()
