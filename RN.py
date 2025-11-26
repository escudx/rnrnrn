import sys, traceback, os, json
from typing import Optional
import tkinter as tk
from tkinter import messagebox, filedialog
import customtkinter as ctk


def _enable_dpi_awareness():
    try:
        import ctypes
        # Usa System Aware (1) para evitar glitches em monitores mistos.
        # O nível 2 (Per Monitor) causava transparência/artefatos ao arrastar
        # entre telas de DPI diferentes.
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


_enable_dpi_awareness()


def _apply_dark_title_bar(win):
    """Força a barra de título do Windows a ficar escura."""
    try:
        import ctypes
        win.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(win.winfo_id())
        # 20 = DWMWA_USE_IMMERSIVE_DARK_MODE
        value = ctypes.c_int(1)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(value), 4)
    except Exception:
        pass


class SafeCTkTextbox(ctk.CTkTextbox):
    _FORBIDDEN_TAG_OPTIONS = {"font", "ctk_font"}

    def tag_config(self, tagName, **kwargs):  # noqa: N802 (Tkinter camelCase)
        cleaned = {k: v for k, v in kwargs.items() if k not in self._FORBIDDEN_TAG_OPTIONS}
        return super().tag_config(tagName, **cleaned)

    def tag_configure(self, tagName, **kwargs):  # noqa: N802 (Tkinter camelCase)
        cleaned = {k: v for k, v in kwargs.items() if k not in self._FORBIDDEN_TAG_OPTIONS}
        return super().tag_configure(tagName, **cleaned)


def _theme_color(widget: str, option: str, fallback):
    try:
        value = ctk.ThemeManager.theme[widget][option]
        if isinstance(value, str):
            return (value, value)
        return value
    except Exception:
        return fallback


def _mix_hex(color_a: str, color_b: str, factor: float) -> str:
    try:
        ca = color_a.lstrip("#")
        cb = color_b.lstrip("#")
        if len(ca) != 6 or len(cb) != 6:
            raise ValueError
        factor = max(0.0, min(1.0, float(factor)))
        ra, ga, ba = (int(ca[i : i + 2], 16) for i in (0, 2, 4))
        rb, gb, bb = (int(cb[i : i + 2], 16) for i in (0, 2, 4))
        blend = lambda va, vb: int(round(va + (vb - va) * factor))
        r = blend(ra, rb)
        g = blend(ga, gb)
        b = blend(ba, bb)
        return f"#{r:02x}{g:02x}{b:02x}"
    except Exception:
        return color_a


# --- PALETA ESTILO CÁPSULA (V6) ---
# Fundo cinza escuro para os blocos (Cápsulas)
CAPSULE_BG = ("#f0f0f0", "#2b2b2b")
# Borda sutil para definir os blocos
CAPSULE_BORDER = ("#dce4ee", "#454545")

# Manter cores de ação
DANGER_BG = ("#ff6b6b", "#2e2e2e")
DANGER_HOVER = ("#ff8080", "#c9585c")
PRIMARY_BTN = _theme_color("CTkButton", "fg_color", ("#3B8ED0", "#1F6AA5"))
PRIMARY_HOVER = _theme_color("CTkButton", "hover_color", ("#36719F", "#144870"))
BADGE_BG = PRIMARY_BTN
BADGE_TEXT = ("#ffffff", "#ffffff")
HINT_TEXT = ("gray60", "gray40")
ACTION_BAR_TEXT = _theme_color("CTkLabel", "text_color", ("#1a1a1a", "#f4f4f4"))
PANEL_BG = _theme_color("CTkFrame", "fg_color", ("#f4f4f5", "#18181b"))
ACTION_BAR_BG = (_mix_hex(PANEL_BG[0], "#000000", 0.02), _mix_hex(PANEL_BG[1], "#000000", 0.05))
ACTION_BAR_BORDER = (_mix_hex(ACTION_BAR_BG[0], "#000000", 0.1), _mix_hex(ACTION_BAR_BG[1], "#000000", 0.14))
COMBO_BORDER = _theme_color("CTkEntry", "border_color", ("#b5b5c8", "gray45"))
INPUT_BG = (_mix_hex(PANEL_BG[0], "#000000", 0.08), _mix_hex(PANEL_BG[1], "#ffffff", 0.4))


def _center_window(win, width=None, height=None, parent=None):
    try:
        win.update_idletasks()
        w = width or win.winfo_reqwidth()
        h = height or win.winfo_reqheight()

        if parent:
            x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (w // 2)
            y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (h // 2)
        else:
            ws = win.winfo_screenwidth()
            hs = win.winfo_screenheight()
            x = (ws // 2) - (w // 2)
            y = (hs // 2) - (h // 2)

        x = max(0, x)
        y = max(0, y)

        win.geometry(f"{w}x{h}+{x}+{y}")

        # --- CORREÇÃO APLICADA AQUI ---
        # Força a barra escura após o posicionamento final
        _apply_dark_title_bar(win)
        # ------------------------------
    except Exception:
        pass

def _show_fatal_error(exc: BaseException):
    tb = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    try:
        log_path = os.path.join(os.path.dirname(__file__), "rn_error.log")
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(tb)
    except Exception:
        pass
    try:
        _r = tk.Tk(); _r.withdraw()
        messagebox.showerror("Erro inesperado", tb)
        _r.destroy()
    except Exception:
        print(tb, file=sys.stderr)

sys.excepthook = lambda et, ev, tb: _show_fatal_error(ev)

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

CUR_L, CUR_R = "“", "”"

RESP_DEFAULTS = ["Escritório Externo", "Jurídico Interno", "Solicitante"]
RESP_TEXT_FREE = "Texto livre…"
SLA_TIPOS   = ["Dias úteis (fixo)", "Dias corridos (fixo)", "D- (antes do Marco)", "D+ (Apos o Marco)"]
GATILHOS    = [
    "Sempre que inserido novo OBJETO",
    "Em TAREFA se CAMPO for RESPOSTA",
    "Concluída TAREFA",
    "Após EVENTO",
]
OPERADORES  = ["for", "for respondido como"]

_LANG = "pt"

TR = {
    "pt": {
        "when_new_object": "Sempre que inserido novo {obj}",
        "when_task_field_answer": "Na conclusão da tarefa {task}, se o campo {field} for respondido com {answer}",
        "when_task_done": "Concluída a tarefa {task}",
        "when_after_event": "Após {event}",
        "action_task": "é acionada a tarefa {task}",
        "action_resp": ", de responsabilidade do {resp}",
        "action_flow": "deverá ser acionado o fluxo de {flow}",
        "action_status": "o status é atualizado para {status}",
        "action_return": "o fluxo retornará a fase {task}{restart}",
        "restart_on": ", reiniciando seu SLA",
        "restart_off": "",
        "closing_partial": "o fluxo é encerrado parcialmente",
        "closing_total": "o fluxo é encerrado totalmente",
        "unit_working_singular": "dia útil",
        "unit_working_plural": "dias úteis",
        "unit_calendar_singular": "dia corrido",
        "unit_calendar_plural": "dias corridos",
        "holidays_suffix": ", e considerando feriados",
        "sla_fixed_working": "com SLA de {n} {unit}, contados a partir de hoje{hol}",
        "sla_fixed_calendar": "com SLA de {n} {unit}, contados a partir de hoje",
        "sla_d_minus": "com SLA D-{n} em relação ao campo de data {marco}{hol}",
        "sla_d_plus": "com SLA de {n} {unit} após a {marco}{hol}",
        "and": " e ",
        "or": " ou ",
        "preview_placeholder": "Use os campos ao lado para montar a RN. Os textos serão exibidos aqui.",
        "mem_hint": "Cole itens (um por linha) e clique em Importar para adicioná-los.",
        "mem_saved_label": "Itens salvos",
        "mem_import_label": "Importar vários itens",
        "mem_none": "Nenhum item salvo ainda",
        "mem_one": "1 item salvo",
        "mem_many": "{n} itens salvos",
        "mem_new_placeholder": "Novo item...",
        "mem_add_button": "Adicionar",
        "mem_resp_label": "Responsáveis padrão",
        "mem_resp_import_label": "Importar responsáveis padrão",
        "mem_resp_hint": "Cole responsáveis (um por linha) e clique em Importar para adicioná-los.",
        "mem_resp_new_placeholder": "Novo responsável...",
        "mem_resp_add_button": "Adicionar responsável",
        "mem_resp_none": "Nenhum responsável salvo ainda",
        "mem_resp_one": "1 responsável salvo",
        "mem_resp_many": "{n} responsáveis salvos",
    },
    "es": {
        "when_new_object": "Siempre que se inserte un nuevo {obj}",
        "when_task_field_answer": "En la conclusión de la tarea {task}, si el campo {field} se responde con {answer}",
        "when_task_done": "Concluida la tarea {task}",
        "when_after_event": "Después de {event}",
        "action_task": "se activa la tarea {task}",
        "action_resp": ", a cargo de {resp}",
        "action_flow": "deberá activarse el flujo de {flow}",
        "action_status": "el estado se actualiza a {status}",
        "action_return": "el flujo volverá a la fase {task}{restart}",
        "restart_on": ", reiniciando su SLA",
        "restart_off": "",
        "closing_partial": "el flujo se cierra parcialmente",
        "closing_total": "el flujo se cierra totalmente",
        "unit_working_singular": "día hábil",
        "unit_working_plural": "días hábiles",
        "unit_calendar_singular": "día corrido",
        "unit_calendar_plural": "días corridos",
        "holidays_suffix": ", y considerando feriados",
        "sla_fixed_working": "con SLA de {n} {unit}, contados a partir de hoy{hol}",
        "sla_fixed_calendar": "con SLA de {n} {unit}, contados a partir de hoy",
        "sla_d_minus": "con SLA D-{n} con respecto al campo de fecha {marco}{hol}",
        "sla_d_plus": "con SLA de {n} {unit} después de {marco}{hol}",
        "and": " y ",
        "or": " o ",
        "preview_placeholder": "Utiliza los campos de la izquierda para construir la RN. El texto aparecerá aquí.",
        "mem_hint": "Pega los ítems (uno por línea) y haz clic en Importar para agregarlos.",
        "mem_saved_label": "Ítems guardados",
        "mem_import_label": "Importar varios ítems",
        "mem_none": "Ningún ítem guardado todavía",
        "mem_one": "1 ítem guardado",
        "mem_many": "{n} ítems guardados",
        "mem_new_placeholder": "Nuevo ítem...",
        "mem_add_button": "Agregar",
        "mem_resp_label": "Responsables predeterminados",
        "mem_resp_import_label": "Importar responsables predeterminados",
        "mem_resp_hint": "Pega responsables (uno por línea) y haz clic en Importar para agregarlos.",
        "mem_resp_new_placeholder": "Nuevo responsable...",
        "mem_resp_add_button": "Agregar responsable",
        "mem_resp_none": "Ningún responsable guardado todavía",
        "mem_resp_one": "1 responsable guardado",
        "mem_resp_many": "{n} responsables guardados",
    },
}

OPS_ES = {
    "for": "es",
    "for respondido como": "fue respondido con",
}

def set_lang(lang: str):
    global _LANG
    _LANG = "es" if str(lang).lower().startswith("es") else "pt"

def get_lang() -> str:
    return _LANG

def _t(key: str) -> str:
    pack = TR.get(get_lang(), TR["pt"])
    return pack.get(key, TR["pt"].get(key, key))

def _plural_unit(kind: str, n: int) -> str:
    if kind == "working":
        return _t("unit_working_singular") if n == 1 else _t("unit_working_plural")
    else:
        return _t("unit_calendar_singular") if n == 1 else _t("unit_calendar_plural")

def _render_sla(tipo: str, dias: int, marco: str, feriados: bool) -> str:
    try:
        dias = int(dias)
    except (ValueError, TypeError):
        dias = 0
        
    hol = _t("holidays_suffix") if feriados else ""
    if tipo.startswith("Dias úteis"):
        unit = _plural_unit("working", dias)
        return _t("sla_fixed_working").format(n=dias, unit=unit, hol=hol)
    if tipo.startswith("Dias corridos"):
        unit = _plural_unit("calendar", dias)
        return _t("sla_fixed_calendar").format(n=dias, unit=unit)
    if tipo.startswith("D- "):
        return _t("sla_d_minus").format(n=dias, marco=f"{CUR_L}{marco}{CUR_R}", hol=hol)
    if tipo.startswith("D+ "):
        unit = _plural_unit("working", dias)
        return _t("sla_d_plus").format(n=dias, unit=unit, marco=f"{CUR_L}{marco}{CUR_R}", hol=hol)
    return ""

def _join_conditions(conds, conj_flag: str) -> str:
    if not conds:
        return ""
    connector = _t("and") if conj_flag == "E" else _t("or")
    return connector.join(conds)

def _cond_to_text(campo: str, op: str, valor: str) -> str:
    lang = get_lang()
    cq = f"{CUR_L}{campo}{CUR_R}" if campo else ""

    if lang == "es":
        verbo = OPS_ES.get(op, op)
        return f"el campo {cq} {verbo} {CUR_L}{valor}{CUR_R}"
    
    return f"o campo {cq} {op} {CUR_L}{valor}{CUR_R}"

def _acao_tarefa_texto(tarefa: str, responsavel: str, sla_txt: str) -> str:
    base = _t("action_task").format(task=f"{CUR_L}{tarefa}{CUR_R}")
    if responsavel:
        base += _t("action_resp").format(resp=responsavel)
    if sla_txt:
        base += f", {sla_txt}"
    return base

def _acao_status_texto(status: str) -> str:
    return _t("action_status").format(status=f"{CUR_L}{status}{CUR_R}")

def _acao_fluxo_texto(fluxo: str) -> str:
    return _t("action_flow").format(flow=f"{CUR_L}{fluxo}{CUR_R}")

def _acao_retornar_texto(tarefa: str, reiniciar: bool) -> str:
    restart = _t("restart_on") if reiniciar else _t("restart_off")
    return _t("action_return").format(task=f"{CUR_L}{tarefa}{CUR_R}", restart=restart)

def _acao_encerramento(parcial: bool) -> str:
    return _t("closing_partial") if parcial else _t("closing_total")

def _compose_rn(idx: int, when: str, cond: str, acoes: str) -> str:
    linha = f"RN{idx}: {when}"
    if cond:
        link = "e " if (" se " in when or " Se " in when) else "se "
        linha += f", {link}{cond}"
    linha += f", {acoes}." if acoes else "."
    return linha

class CollapsibleGroup(ctk.CTkFrame):
    def __init__(self, master, text="Grupo", start_expanded=True, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.grid_columnconfigure(0, weight=1)
        self._expanded = start_expanded
        self._text = text

        self.title_frame = ctk.CTkFrame(self)
        self.title_frame.grid(row=0, column=0, sticky="ew")
        self.title_frame.grid_columnconfigure(0, weight=1)
        
        self.title_label = ctk.CTkLabel(self.title_frame, text=self._get_title_text(), font=ctk.CTkFont(weight="bold"))
        self.title_label.grid(row=0, column=0, sticky="w", padx=8, pady=8)

        self.inner = ctk.CTkFrame(self, fg_color="transparent")
        self.inner.grid_columnconfigure(0, weight=1)

        self.title_frame.bind("<Button-1>", self._toggle)
        self.title_label.bind("<Button-1>", self._toggle)
        
        self._update_visibility()

    def _get_title_text(self):
        return f"v  {self._text}" if self._expanded else f">  {self._text}"

    def _toggle(self, event=None):
        self._expanded = not self._expanded
        self._update_visibility()

    def _update_visibility(self):
        self.title_label.configure(text=self._get_title_text())
        if self._expanded:
            self.inner.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
        else:
            self.inner.grid_forget()

    def get_inner_frame(self):
        return self.inner

def scrollable_body(widget: ctk.CTkScrollableFrame):
    return getattr(widget, "scrollable_frame", getattr(widget, "_scrollable_frame", widget))

class IntSpin(ctk.CTkFrame):
    def __init__(self, master, from_=0, to=365, width=90, variable=None, on_change=None):
        super().__init__(master)
        self.min = from_; self.max = to
        self.on_change = on_change
        self.var = variable or tk.IntVar(value=0)
        ctk.CTkButton(self, text="−", width=28, command=self._dec).grid(row=0, column=0, padx=(0,2))
        self.entry = ctk.CTkEntry(self, width=width-56)
        self.entry.grid(row=0, column=1)
        ctk.CTkButton(self, text="+", width=28, command=self._inc).grid(row=0, column=2, padx=(2,0))
        self.entry.insert(0, str(self.var.get()))
        self.entry.bind("<FocusOut>", self._sync_from_entry)
        self.entry.bind("<Return>", self._sync_from_entry)

    def _clamp(self, v):
        return max(self.min, min(self.max, v))

    def _notify(self):
        if callable(self.on_change):
            self.on_change()

    def _sync_from_entry(self, *_):
        try:
            v = int(self.entry.get().strip())
        except Exception:
            v = self.var.get()
        v = self._clamp(v)
        if v != self.var.get():
            self.var.set(v)
        self.entry.delete(0, "end"); self.entry.insert(0, str(v))
        self._notify()

    def _dec(self):
        self.var.set(self._clamp(int(self.var.get()) - 1))
        self.entry.delete(0, "end"); self.entry.insert(0, str(self.var.get()))
        self._notify()

    def _inc(self):
        self.var.set(self._clamp(int(self.var.get()) + 1))
        self.entry.delete(0, "end"); self.entry.insert(0, str(self.var.get()))
        self._notify()

class MemManagerTab(ctk.CTkFrame):
    def __init__(self, master, title, mem_list, refresh_cb, **kwargs):
        super().__init__(master, fg_color="transparent")
        self.app = master.winfo_toplevel()
        self.mem_list = mem_list
        self.refresh_cb = refresh_cb

        hint_text = kwargs.get("hint_text", _t("mem_hint"))
        import_label = kwargs.get("import_label", _t("mem_import_label"))
        list_label = kwargs.get("list_label", _t("mem_saved_label"))
        placeholder = kwargs.get("placeholder", _t("mem_new_placeholder"))
        add_button_text = kwargs.get("add_button_text", _t("mem_add_button"))
        self.forbidden = {self._norm(v).lower() for v in kwargs.get("forbidden_values", []) if self._norm(v)}
        layout = kwargs.get("layout", "vertical").lower()

        side_by_side = layout in {"horizontal", "side", "two-column", "grid"}

        self.grid_columnconfigure(0, weight=1, uniform="mem")
        if side_by_side:
            self.grid_columnconfigure(1, weight=1, uniform="mem")
            self.grid_rowconfigure(1, weight=1)
        else:
            self.grid_rowconfigure(2, weight=1)

        # --- 1. HEADER CÁPSULA ---
        header_cols = 2 if side_by_side else 1
        header = ctk.CTkFrame(self, corner_radius=8, fg_color=CAPSULE_BG, border_width=1, border_color=CAPSULE_BORDER)
        header.grid(row=0, column=0, columnspan=header_cols, sticky="ew", pady=(0, 12))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text=title,
            font=ctk.CTkFont(size=16, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=16, pady=(16, 12))

        self.count_var = tk.StringVar(value="")
        ctk.CTkLabel(header, textvariable=self.count_var, font=ctk.CTkFont(size=12), text_color=("gray50", "gray70")).grid(row=0, column=1, sticky="e", padx=12, pady=(10, 0))

        if hint_text and not side_by_side:
            ctk.CTkLabel(header, text=hint_text, justify="left", wraplength=480, text_color=HINT_TEXT).grid(row=1, column=0, columnspan=2, sticky="w", padx=12, pady=(6, 10))

        # --- 2. IMPORT CARD CÁPSULA ---
        import_card = ctk.CTkFrame(self, corner_radius=8, fg_color=CAPSULE_BG, border_width=1, border_color=CAPSULE_BORDER)
        if side_by_side:
            import_card.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(0, 12))
        else:
            import_card.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        import_card.grid_columnconfigure(0, weight=1)
        import_card.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(import_card, text=import_label, font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        box_h = 140 if side_by_side else 100
        self.import_box = ctk.CTkTextbox(import_card, height=box_h)
        self.import_box.grid(row=1, column=0, sticky="nsew", padx=12)

        if hint_text and side_by_side:
            ctk.CTkLabel(import_card, text=hint_text, justify="left", wraplength=300, text_color=HINT_TEXT, font=ctk.CTkFont(size=11)).grid(row=2, column=0, sticky="w", padx=12, pady=(6, 0))

        ctk.CTkButton(import_card, text="Importar", width=120, command=self._import_items, fg_color=PRIMARY_BTN, hover_color=PRIMARY_HOVER).grid(row=3, column=0, sticky="e", padx=12, pady=(10, 12))

        # --- 3. LISTA CÁPSULA ---
        list_card = ctk.CTkFrame(self, corner_radius=8, fg_color=CAPSULE_BG, border_width=1, border_color=CAPSULE_BORDER)
        if side_by_side:
            list_card.grid(row=1, column=1, sticky="nsew", padx=(8, 0), pady=(0, 12))
        else:
            list_card.grid(row=2, column=0, sticky="nsew")
        list_card.grid_columnconfigure(0, weight=1)
        list_card.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(list_card, text=list_label, font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        self.scroll_frame = ctk.CTkScrollableFrame(list_card, fg_color="transparent")
        self.scroll_frame.grid(row=1, column=0, sticky="nsew", padx=6, pady=0)
        self.scroll_frame.grid_columnconfigure(0, weight=1)
        self.body = getattr(self.scroll_frame, "scrollable_frame", self.scroll_frame)
        self.body.grid_columnconfigure(0, weight=1)

        try:
            self.app._enable_scrollwheel(self.scroll_frame)
        except Exception:
            pass

        self.rows = []

        footer = ctk.CTkFrame(list_card, fg_color="transparent")
        footer.grid(row=2, column=0, sticky="ew", padx=12, pady=(10, 12))
        footer.grid_columnconfigure(0, weight=1)

        self.new_entry = ctk.CTkEntry(footer, placeholder_text=placeholder)
        self.new_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.new_entry.bind("<Return>", self._add_item)
        self.add_btn = ctk.CTkButton(footer, text=add_button_text, width=100, command=self._add_item, fg_color=PRIMARY_BTN, hover_color=PRIMARY_HOVER)
        self.add_btn.grid(row=0, column=1, sticky="e")

        self._rebuild_list()

    def _norm(self, s):
        return " ".join((s or "").split()).strip()

    def _is_allowed(self, v):
        norm = self._norm(v)
        return norm.lower() not in self.forbidden if norm else False

    def _add_item(self, e=None):
        v = self._norm(self.new_entry.get())
        if not v or len(v) < 3:
            return
        if not self._is_allowed(v):
            return
        if not any(v.lower() == x.lower() for x in self.mem_list):
            self.mem_list.append(v)
            self._rebuild_list()
        self.new_entry.delete(0, "end")

    def _import_items(self):
        txt = self.import_box.get("1.0", "end").strip()
        if not txt:
            return
        added = 0
        for x in txt.split("\n"):
            val = self._norm(x)
            if len(val) >= 3 and self._is_allowed(val) and not any(val.lower() == i.lower() for i in self.mem_list):
                self.mem_list.append(val)
                added += 1
        if added:
            self._rebuild_list()
            messagebox.showinfo("Importar", f"{added} itens.")
        self.import_box.delete("1.0", "end")

    def _remove_item(self, val):
        if messagebox.askyesno("Remover", f"Excluir '{val}'?"):
            if val in self.mem_list:
                self.mem_list.remove(val)
                self._rebuild_list()

    def _rename_item(self, old, new_val):
        nv = self._norm(new_val)
        if len(nv) < 3 or not self._is_allowed(nv):
            return
        if any(nv.lower() == x.lower() for x in self.mem_list if x != old):
            return
        if old in self.mem_list:
            self.mem_list[self.mem_list.index(old)] = nv
            self._rebuild_list()

    def _rebuild_list(self):
        for r in self.rows:
            r.destroy()
        self.rows.clear()
        self.mem_list.sort(key=str.lower)
        total = len(self.mem_list)
        self.count_var.set(f"{total} itens salvos" if total > 1 else f"{total} item salvo")

        for i, item in enumerate(self.mem_list):
            row = ctk.CTkFrame(self.body, corner_radius=6, fg_color=CAPSULE_BG, border_width=1, border_color=CAPSULE_BORDER)
            row.grid(row=i, column=0, sticky="ew", pady=(0, 8))
            row.grid_columnconfigure(0, weight=1)

            entry = ctk.CTkEntry(row, fg_color="transparent", border_width=0)
            entry.insert(0, item)
            entry.grid(row=0, column=0, sticky="ew", padx=(8, 4), pady=6)

            btn_box = ctk.CTkFrame(row, fg_color="transparent")
            btn_box.grid(row=0, column=1, sticky="e", padx=(0, 8), pady=6)

            ctk.CTkButton(btn_box, text="Renomear", width=70, height=24, font=ctk.CTkFont(size=11), command=lambda e=entry, o=item: self._rename_item(o, e.get()), fg_color=PRIMARY_BTN, hover_color=PRIMARY_HOVER).pack(side="left", padx=(0, 6))

            ctk.CTkButton(btn_box, text="Remover", width=70, height=24, font=ctk.CTkFont(size=11), command=lambda o=item: self._remove_item(o), fg_color=DANGER_BG, hover_color=DANGER_HOVER).pack(side="left")

            self.rows.append(row)
        self.refresh_cb()

class LinhaCondicao(ctk.CTkFrame):
    def __init__(self, master, on_change, on_remove):
        super().__init__(master)
        self.on_change = on_change
        self.on_remove = on_remove
        self.var_campo = tk.StringVar()
        self.var_op = tk.StringVar(value=OPERADORES[0])
        self.var_valor = tk.StringVar()

        self.grid_columnconfigure(0, weight=0, minsize=120)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0)

        ctk.CTkLabel(self, text="Campo:").grid(row=0, column=0, sticky="e", padx=(4, 6), pady=(4, 2))
        ctk.CTkEntry(self, textvariable=self.var_campo).grid(row=0, column=1, sticky="ew", padx=(0, 4), pady=(4, 2))

        ctk.CTkLabel(self, text="Operador:").grid(row=1, column=0, sticky="e", padx=(4, 6), pady=(0, 2))
        ctk.CTkComboBox(
            self,
            values=OPERADORES,
            variable=self.var_op,
            command=lambda *_: self.on_change(),
        ).grid(row=1, column=1, sticky="ew", padx=(0, 4), pady=(0, 2))

        ctk.CTkLabel(self, text="Valor / Resposta:").grid(row=2, column=0, sticky="e", padx=(4, 6), pady=(0, 2))
        ctk.CTkEntry(self, textvariable=self.var_valor).grid(row=2, column=1, sticky="ew", padx=(0, 4), pady=(0, 4))

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=0, column=2, rowspan=3, sticky="ne", padx=(0, 4), pady=(4, 4))

        app = self.winfo_toplevel()
        ctk.CTkButton(
            btn_frame,
            text="↑",
            width=28,
            command=lambda: app._move_row(self, -1, app.cond_rows, app.frm_conds_body),
        ).pack(pady=(0, 2))
        ctk.CTkButton(
            btn_frame,
            text="↓",
            width=28,
            command=lambda: app._move_row(self, 1, app.cond_rows, app.frm_conds_body),
        ).pack(pady=(0, 4))
        ctk.CTkButton(btn_frame, text="Remover", command=self._remove, width=80).pack()

        for v in (self.var_campo, self.var_op, self.var_valor):
            v.trace_add("write", lambda *a: self.on_change())

    def _remove(self):
        try:
            self.destroy()
        finally:
            if callable(self.on_remove):
                self.on_remove(self)
            self.on_change()

    def request_remove(self):
        self._remove()

    def to_dict(self):
        return {"campo": self.var_campo.get(), "op": self.var_op.get(), "valor": self.var_valor.get()}

    def from_dict(self, d: dict):
        self.var_campo.set(d.get("campo", ""))
        self.var_op.set(d.get("op", OPERADORES[0]))
        self.var_valor.set(d.get("valor", ""))

    def to_text(self):
        c, o, v = self.var_campo.get().strip(), self.var_op.get(), self.var_valor.get().strip()
        if not c:
            return ""
        if not v:
            return ""
        return _cond_to_text(c, o, v)

class LinhaAcao(ctk.CTkFrame):
    def __init__(self, master, on_change, on_remove, default_resp="Escritório Externo", default_resp_free="", row_list_key="acao_rows"):
        super().__init__(master)
        self.on_change = on_change
        self.on_remove = on_remove
        self.row_list_key = row_list_key

        self.var_tipo = tk.StringVar(value="Acionar Tarefa")

        self.var_tarefa = tk.StringVar()

        resp_options = self._resp_options()
        preset = (default_resp or "").strip()
        if preset and preset not in resp_options and preset != RESP_TEXT_FREE:
            try:
                self._app()._resp_add(preset)
                resp_options = self._resp_options()
            except Exception:
                pass

        if preset and preset in resp_options and preset != RESP_TEXT_FREE:
            resp_initial = preset
            resp_free_initial = (default_resp_free or "")
        elif preset == RESP_TEXT_FREE:
            resp_initial = RESP_TEXT_FREE
            resp_free_initial = default_resp_free or ""
        else:
            resp_initial = RESP_TEXT_FREE
            resp_free_initial = default_resp_free or (preset if preset else "")

        self.var_resp = tk.StringVar(value=resp_initial)
        self.var_resp_livre = tk.StringVar(value=resp_free_initial)
        self.var_sla_tipo = tk.StringVar(value=SLA_TIPOS[0])
        self.var_sla_dias = tk.IntVar(value=2)
        self.var_sla_marco = tk.StringVar(value="Prazo Fatal da Peça")
        self.var_sla_fer = tk.BooleanVar(value=True)
        self.var_status = tk.StringVar()
        self.var_fluxo = tk.StringVar()
        self.var_texto = tk.StringVar()

        self.var_ret_tarefa = tk.StringVar()
        self.var_ret_restart = tk.BooleanVar(value=True)

        self.grid_columnconfigure(0, weight=0, minsize=120)
        self.grid_columnconfigure(1, weight=1)

        self.tipo_bar = ctk.CTkFrame(
            self,
            fg_color=ACTION_BAR_BG,
            border_color=ACTION_BAR_BORDER,
            border_width=1,
            corner_radius=10,
        )
        self.tipo_bar.grid(
            row=0,
            column=0,
            columnspan=2,
            sticky="ew",
            padx=(6, 6),
            pady=(6, 10),
        )
        self.tipo_bar.grid_columnconfigure(0, weight=0)
        self.tipo_bar.grid_columnconfigure(1, weight=1)
        self.tipo_bar.grid_columnconfigure(2, weight=0)
        self.tipo_bar.grid_columnconfigure(3, weight=0)

        label_font = ctk.CTkFont(size=12, weight="bold")
        ctk.CTkLabel(
            self.tipo_bar,
            text="Tipo de ação",
            font=label_font,
            text_color=ACTION_BAR_TEXT,
        ).grid(row=0, column=0, sticky="w", padx=(14, 10), pady=10)

        self.cbo_tipo = ctk.CTkComboBox(
            self.tipo_bar,
            values=[
                "Acionar Tarefa",
                "Atualizar Status",
                "Texto Livre",
            ],
            variable=self.var_tipo,
            command=lambda *_: self._refresh(),
            border_width=1,
            border_color=COMBO_BORDER,
        )
        self.cbo_tipo.grid(row=0, column=1, sticky="ew", padx=(0, 12), pady=10)

        arrow_row = ctk.CTkFrame(self.tipo_bar, fg_color="transparent")
        arrow_row.grid(row=0, column=2, sticky="e", padx=(0, 4), pady=8)

        app = self.winfo_toplevel()
        row_list = getattr(app, self.row_list_key, [])
        frame_parent = self.master

        up_btn = ctk.CTkButton(
            arrow_row,
            text="↑",
            width=28,
            command=lambda: app._move_row(self, -1, row_list, frame_parent),
        )
        up_btn.grid(row=0, column=0, padx=(0, 6))

        down_btn = ctk.CTkButton(
            arrow_row,
            text="↓",
            width=28,
            command=lambda: app._move_row(self, 1, row_list, frame_parent),
        )
        down_btn.grid(row=0, column=1)

        ctk.CTkButton(
            self.tipo_bar,
            text="Remover",
            command=self._remove,
            width=90,
        ).grid(row=0, column=3, sticky="e", padx=(6, 12), pady=8)

        self.frm_dyn = ctk.CTkFrame(self, fg_color="transparent")
        self.frm_dyn.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(4, 4), pady=(0, 4))
        self.frm_dyn.grid_columnconfigure(0, weight=0, minsize=110)
        self.frm_dyn.grid_columnconfigure(1, weight=1)

        self._refresh()

    def _app(self):
        return self.winfo_toplevel()

    def _register_task_combo(self, combo, var_getter):
        try:
            self._app()._mem_register_task_combo(combo)
            self._app()._mem_bind_combo_capture(combo, var_getter, bucket="task")
        except Exception:
            pass

    def _register_field_combo(self, combo, var_getter):
        try:
            self._app()._mem_register_field_combo(combo)
            self._app()._mem_bind_combo_capture(combo, var_getter, bucket="field")
        except Exception:
            pass

    def _resp_options(self):
        try:
            return self._app()._resp_get_options()
        except Exception:
            base = sorted(RESP_DEFAULTS, key=str.lower)
            if RESP_TEXT_FREE not in base:
                base.append(RESP_TEXT_FREE)
            return base

    def _register_resp_combo(self, combo, var_getter, on_select=None):
        try:
            self._app()._resp_register_combo(combo)
            self._app()._resp_bind_combo_capture(combo, var_getter, on_select=on_select)
        except Exception:
            pass

    def _remove(self):
        try:
            self.destroy()
        finally:
            if callable(self.on_remove):
                self.on_remove(self)
            self.on_change()

    def request_remove(self):
        self._remove()

    def _refresh(self):
        for w in self.frm_dyn.winfo_children():
            w.destroy()
        t = self.var_tipo.get()

        if t.startswith("Encerrar Fluxo"):
            self.frm_dyn.grid_forget()
        else:
            self.frm_dyn.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(4, 4), pady=(0, 4))

        if t == "Acionar Tarefa":
            row = 0
            ctk.CTkLabel(self.frm_dyn, text="Tarefa:").grid(row=row, column=0, sticky="e", pady=(0, 2), padx=(0, 6))
            cb_tarefa = ctk.CTkComboBox(self.frm_dyn, values=self._app()._mem_get_tasks(), variable=self.var_tarefa)
            cb_tarefa.grid(row=row, column=1, sticky="ew", pady=(0, 2))
            self._register_task_combo(cb_tarefa, lambda: self.var_tarefa.get())
            row += 1

            ctk.CTkLabel(self.frm_dyn, text="Responsável:").grid(row=row, column=0, sticky="e", pady=(2, 2), padx=(0, 6))
            cb_resp = ctk.CTkComboBox(
                self.frm_dyn,
                values=self._resp_options(),
                variable=self.var_resp,
                border_width=1,
                border_color=COMBO_BORDER,
            )
            cb_resp.grid(row=row, column=1, sticky="ew", pady=(2, 2))
            row += 1

            resp_row = row
            self.entry_resp_livre = ctk.CTkEntry(
                self.frm_dyn,
                textvariable=self.var_resp_livre,
                placeholder_text="Responsável (texto livre)",
            )

            def _toggle_resp(*_):
                if self.var_resp.get() == RESP_TEXT_FREE:
                    self.entry_resp_livre.grid(row=resp_row, column=1, sticky="ew", pady=(0, 2))
                else:
                    self.var_resp_livre.set("")
                    self.entry_resp_livre.grid_forget()
                self.on_change()

            self._register_resp_combo(cb_resp, lambda: self.var_resp.get(), on_select=_toggle_resp)
            _toggle_resp()
            row = resp_row + 1 if self.entry_resp_livre.winfo_ismapped() else resp_row

            ctk.CTkLabel(self.frm_dyn, text="Tipo de SLA:").grid(row=row, column=0, sticky="e", pady=(2, 2), padx=(0, 6))
            cb_sla = ctk.CTkComboBox(self.frm_dyn, values=SLA_TIPOS, variable=self.var_sla_tipo)
            cb_sla.grid(row=row, column=1, sticky="ew", pady=(2, 2))
            row += 1

            ctk.CTkLabel(self.frm_dyn, text="SLA:").grid(row=row, column=0, sticky="e", pady=(2, 2), padx=(0, 6))
            
            sla_frame = ctk.CTkFrame(self.frm_dyn, fg_color="transparent")
            sla_frame.grid(row=row, column=1, sticky="w")

            IntSpin(
                sla_frame,
                from_=0,
                to=365,
                variable=self.var_sla_dias,
                width=120,
                on_change=self.on_change,
            ).grid(row=0, column=0, sticky="w", pady=(2, 2))

            self.chk_feriados = ctk.CTkCheckBox(
                sla_frame,
                text="Considerar feriados",
                variable=self.var_sla_fer,
                onvalue=True,
                offvalue=False,
                command=self.on_change,
            )
            self.chk_feriados.grid(row=0, column=1, sticky="w", pady=(2, 2), padx=(10, 0))
            row += 1

            marco_row = row
            self.lbl_marco = ctk.CTkLabel(self.frm_dyn, text="Campo de Data:")
            self.lbl_marco.grid_configure(sticky="e", padx=(0, 6))
            
            cb_marco = ctk.CTkComboBox(self.frm_dyn, values=self._app()._mem_get_fields(), variable=self.var_sla_marco)

            def _apply_sla_visibility(*_):
                tt = self.var_sla_tipo.get()
                is_corridos = (tt == "Dias corridos (fixo)")
                needs_marco = tt in ("D- (antes do Marco)", "D+ (Apos o Marco)")
                try:
                    self.chk_feriados.configure(state=("disabled" if is_corridos else "normal"))
                except Exception:
                    pass
                if needs_marco:
                    self.lbl_marco.grid(row=marco_row, column=0, sticky="e", pady=(2, 2), padx=(0, 6))
                    cb_marco.grid(row=marco_row, column=1, sticky="ew", pady=(2, 2))
                    self._register_field_combo(cb_marco, lambda: self.var_sla_marco.get())
                else:
                    self.lbl_marco.grid_forget()
                    cb_marco.grid_forget()
                self.on_change()

            cb_sla.configure(command=_apply_sla_visibility)
            _apply_sla_visibility()
            row = marco_row + 1 if cb_marco.winfo_ismapped() else marco_row

        elif t == "Atualizar Status":
            ctk.CTkLabel(self.frm_dyn, text="Status:").grid(row=0, column=0, sticky="e", pady=(0, 2), padx=(0, 6))
            ctk.CTkEntry(self.frm_dyn, textvariable=self.var_status).grid(row=0, column=1, sticky="ew", pady=(0, 2))

        elif t == "Acionar Fluxo":
            ctk.CTkLabel(self.frm_dyn, text="Fluxo:").grid(row=0, column=0, sticky="e", pady=(0, 2), padx=(0, 6))
            ctk.CTkEntry(self.frm_dyn, textvariable=self.var_fluxo).grid(row=0, column=1, sticky="ew", pady=(0, 2))

        elif t == "Retornar a Tarefa":
            ctk.CTkLabel(self.frm_dyn, text="Retornar a tarefa:").grid(row=0, column=0, sticky="e", pady=(0, 2), padx=(0, 6))
            cb_ret = ctk.CTkComboBox(self.frm_dyn, values=self._app()._mem_get_tasks(), variable=self.var_ret_tarefa)
            cb_ret.grid(row=0, column=1, sticky="ew", pady=(0, 2))
            self._register_task_combo(cb_ret, lambda: self.var_ret_tarefa.get())
            ctk.CTkCheckBox(
                self.frm_dyn,
                text="Reiniciar SLA",
                variable=self.var_ret_restart,
                onvalue=True,
                offvalue=False,
                command=self.on_change,
            ).grid(row=1, column=1, sticky="w", pady=(2, 2))

        elif t.startswith("Encerrar Fluxo"):
            pass 

        else: 
            ctk.CTkLabel(self.frm_dyn, text="Texto:").grid(row=0, column=0, sticky="e", pady=(0, 2), padx=(0, 6))
            ctk.CTkEntry(self.frm_dyn, textvariable=self.var_texto).grid(row=0, column=1, sticky="ew", pady=(0, 2))

        for v in (
            self.var_tarefa, self.var_resp, self.var_resp_livre, self.var_sla_tipo, self.var_sla_dias,
            self.var_sla_marco, self.var_sla_fer, self.var_status, self.var_fluxo, self.var_texto,
            self.var_ret_tarefa, self.var_ret_restart,
        ):
            try:
                v.trace_add("write", lambda *a: self.on_change())
            except Exception:
                pass

        self.on_change()

    def to_dict(self) -> dict:
        return {
            "tipo": self.var_tipo.get(),
            "tarefa": self.var_tarefa.get(),
            "resp": self.var_resp.get(),
            "resp_livre": self.var_resp_livre.get(),
            "sla_tipo": self.var_sla_tipo.get(),
            "sla_dias": int(self.var_sla_dias.get()),
            "sla_marco": self.var_sla_marco.get(),
            "sla_fer": bool(self.var_sla_fer.get()),
            "status": self.var_status.get(),
            "fluxo": self.var_fluxo.get(),
            "texto": self.var_texto.get(),
            "ret_tarefa": self.var_ret_tarefa.get(),
            "ret_restart": bool(self.var_ret_restart.get()),
        }

    def from_dict(self, d: dict):
        tipo = d.get("tipo", "Acionar Tarefa")
        texto = d.get("texto", "")

        if tipo == "Texto Livre":
            if texto == _t("closing_partial"):
                tipo = "Encerrar Fluxo (Parcial)"
            elif texto == _t("closing_total"):
                tipo = "Encerrar Fluxo (Total)"
        
        self.var_tipo.set(tipo); self._refresh()

        self.var_tarefa.set(d.get("tarefa", ""))

        resp_value = d.get("resp", RESP_DEFAULTS[0] if RESP_DEFAULTS else RESP_TEXT_FREE)
        resp_free = d.get("resp_livre", "")
        try:
            if resp_value and resp_value not in ("", RESP_TEXT_FREE):
                self._app()._resp_add(resp_value)
        except Exception:
            pass

        if resp_value in self._resp_options():
            self.var_resp.set(resp_value)
            self.var_resp_livre.set(resp_free)
        else:
            self.var_resp.set(RESP_TEXT_FREE)
            self.var_resp_livre.set(resp_free or (resp_value if resp_value not in (None, "", RESP_TEXT_FREE) else ""))
        self.var_sla_tipo.set(d.get("sla_tipo", SLA_TIPOS[0])); self._refresh()
        
        try:
            self.var_sla_dias.set(int(d.get("dias", d.get("sla_dias", 2))))
        except (ValueError, TypeError):
            self.var_sla_dias.set(2)
            
        self.var_sla_marco.set(d.get("sla_marco", "Prazo Fatal da Peça"))
        
        try:
            self.var_sla_fer.set(bool(d.get("sla_fer", True)))
        except (ValueError, TypeError):
             self.var_sla_fer.set(True)
             
        self.var_status.set(d.get("status", ""))
        self.var_fluxo.set(d.get("fluxo", ""))
        self.var_texto.set(d.get("texto", ""))
        self.var_ret_tarefa.set(d.get("ret_tarefa", ""))
        
        try:
            self.var_ret_restart.set(bool(d.get("ret_restart", True)))
        except (ValueError, TypeError):
            self.var_ret_restart.set(True)
            
        self._refresh()

    def to_text(self) -> str:
        t = self.var_tipo.get()
        if t == "Acionar Tarefa":
            tarefa = self.var_tarefa.get().strip()
            if not tarefa:
                return ""
            resp = self.var_resp.get()
            if resp == RESP_TEXT_FREE:
                resp = self.var_resp_livre.get().strip()
            sla = _render_sla(self.var_sla_tipo.get(), int(self.var_sla_dias.get()),
                              self.var_sla_marco.get().strip(), bool(self.var_sla_fer.get()))
            return _acao_tarefa_texto(tarefa, resp, sla)
        if t == "Atualizar Status":
            st = self.var_status.get().strip()
            return _acao_status_texto(st) if st else ""
        if t == "Acionar Fluxo":
            fx = self.var_fluxo.get().strip()
            return _acao_fluxo_texto(fx) if fx else ""
        if t == "Retornar a Tarefa":
            return _acao_retornar_texto(self.var_ret_tarefa.get().strip(), bool(self.var_ret_restart.get()))
        
        if t == "Encerrar Fluxo (Parcial)":
            return _acao_encerramento(parcial=True)
        if t == "Encerrar Fluxo (Total)":
            return _acao_encerramento(parcial=False)

        tx = self.var_texto.get().strip()
        return tx if tx else ""

class RNBuilder(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("RN Builder — Montador de Regras")
        self._apply_initial_geometry()
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.flows: dict[str, list[str]] = {"Fluxo Padrão": []}
        self.current_flow: str = "Fluxo Padrão"
        self.start_idx = tk.IntVar(value=1)
        self.flow_var = tk.StringVar(value=self.current_flow)
        self.flow_combo = None

        self._mem_tasks: list[str] = []
        self._mem_fields: list[str] = []
        self._resp_defaults: list[str] = list(RESP_DEFAULTS)
        self._task_combos: list = []
        self._field_combos: list = []
        self._resp_combos: list = []
        self._mem_guard = False

        self._mem_manager_window = None

        initial_resp = (self._resp_defaults[0] if self._resp_defaults else RESP_TEXT_FREE)
        self.var_resp_preset = tk.StringVar(value=initial_resp)
        self.var_resp_preset_free = tk.StringVar()

        self._lang_var = tk.StringVar(value=("Português" if get_lang() == "pt" else "Español"))

        self._build_header()

        self.content = ctk.CTkFrame(self, fg_color="transparent")
        self.content.grid(row=1, column=0, sticky="nsew")
        self.content.grid_rowconfigure(0, weight=1)

        # Substitui o grid fixo por um painel redimensionável (splitter)
        self.paned = tk.PanedWindow(
            self.content,
            orient="horizontal",
            sashwidth=6,
            bg="#2b2b2b",
            bd=0,
        )
        self.paned.pack(fill="both", expand=True, padx=6, pady=6)

        self.left_pane = ctk.CTkFrame(self.paned, fg_color="transparent")
        self.right_pane = ctk.CTkFrame(self.paned, fg_color="transparent")

        self.paned.add(self.left_pane, minsize=450, stretch="always")
        self.paned.add(self.right_pane, minsize=450, stretch="always")

        self._bind_shortcuts()

    def _norm(self, s: str) -> str:
        return " ".join((s or "").split()).strip()

    def _add_unique(self, bucket_list: list, value: str):
        v = self._norm(value)
        if len(v) < 3:
            return
        if any(v.lower() == x.lower() for x in bucket_list):
            return
        bucket_list.append(v)

    def _mem_add_task(self, s: str):
        before = len(self._mem_tasks)
        self._add_unique(self._mem_tasks, s)
        if len(self._mem_tasks) != before:
            self._refresh_task_combos()

    def _mem_add_field(self, s: str):
        before = len(self._mem_fields)
        self._add_unique(self._mem_fields, s)
        if len(self._mem_fields) != before:
            self._refresh_field_combos()

    def _mem_get_tasks(self):
        return list(self._mem_tasks)

    def _mem_get_fields(self):
        return list(self._mem_fields)

    def _mem_register_task_combo(self, combo):
        try:
            if combo not in self._task_combos:
                self._task_combos.append(combo)
            combo.configure(values=self._mem_get_tasks())
        except Exception:
            pass

    def _mem_register_field_combo(self, combo):
        try:
            if combo not in self._field_combos:
                self._field_combos.append(combo)
            combo.configure(values=self._mem_get_fields())
        except Exception:
            pass

    def _get_flow_names(self):
        return list(self.flows.keys()) if self.flows else ["Fluxo Padrão"]

    def _ensure_flow(self, name: str):
        if not name:
            name = "Fluxo Padrão"
        if name not in self.flows:
            self.flows[name] = []
        return name

    def _current_rns(self):
        self.current_flow = self._ensure_flow(self.current_flow)
        return self.flows[self.current_flow]

    def _set_current_flow(self, name: str):
        name = self._ensure_flow(self._norm(name))
        self.current_flow = name
        self.flow_var.set(name)
        self._refresh_flow_controls()
        self._refresh_textbox()

    def _refresh_flow_controls(self):
        names = self._get_flow_names()
        if not names:
            names = ["Fluxo Padrão"]
        if self.current_flow not in names:
            self.current_flow = names[0]
        try:
            self.flow_combo.configure(values=names)
        except Exception:
            pass
        try:
            self.flow_var.set(self.current_flow)
        except Exception:
            pass

    def _on_flow_selected(self, *_):
        name = self.flow_var.get()
        if name not in self.flows:
            name = self._get_flow_names()[0]
        self.current_flow = name
        self.flow_var.set(name)
        self._refresh_textbox()

    def _ask_flow_name(self, title: str, initial: str = "") -> str:
        try:
            dialog = ctk.CTkInputDialog(text=title, title=title)
            try:
                _apply_dark_title_bar(dialog)
            except Exception:
                pass

            # Tenta centralizar usando nossa função global segura
            try:
                _center_window(dialog, width=300, height=200, parent=self)
            except Exception:
                pass

            # Tenta preencher o valor inicial (para Renomear) com segurança
            if initial:
                try:
                    # Tenta acessar _entry (privado) ou entry (publico)
                    entry_widget = getattr(dialog, "_entry", getattr(dialog, "entry", None))
                    if entry_widget:
                        entry_widget.delete(0, "end")
                        entry_widget.insert(0, initial)
                except Exception:
                    pass

            result = dialog.get_input()
        except Exception:
            result = None
        return self._norm(result or "")

    def _new_flow(self):
        name = self._ask_flow_name("Nome do novo fluxo")
        if not name:
            return
        if name in self.flows:
            messagebox.showinfo("Fluxo existente", "Já existe um fluxo com esse nome.", parent=self)
            return
        self.flows[name] = []
        self._set_current_flow(name)

    def _rename_flow(self):
        old = self.current_flow
        name = self._ask_flow_name("Renomear fluxo", initial=old)
        if not name or name == old:
            return
        if name in self.flows:
            messagebox.showinfo("Fluxo existente", "Já existe um fluxo com esse nome.", parent=self)
            return
        self.flows[name] = self.flows.pop(old)
        self._set_current_flow(name)

    def _delete_flow(self):
        if len(self.flows) <= 1:
            messagebox.showinfo("Fluxos", "É necessário ter pelo menos um fluxo.", parent=self)
            return
        if not messagebox.askyesno("Remover fluxo", f"Remover o fluxo '{self.current_flow}' e suas RNs?", parent=self):
            return
        try:
            del self.flows[self.current_flow]
        except Exception:
            pass
        remaining = self._get_flow_names()
        self.current_flow = remaining[0] if remaining else "Fluxo Padrão"
        self._refresh_flow_controls()
        self._refresh_textbox()

    def _resp_get_options(self):
        seen = set()
        options: list[str] = []
        for item in self._resp_defaults:
            norm = self._norm(item)
            if not norm:
                continue
            key = norm.lower()
            if key == RESP_TEXT_FREE.lower():
                continue
            if key not in seen:
                seen.add(key)
                options.append(norm)
        options.sort(key=str.lower)
        if RESP_TEXT_FREE not in options:
            options.append(RESP_TEXT_FREE)
        return options

    def _resp_register_combo(self, combo):
        try:
            if combo not in self._resp_combos:
                self._resp_combos.append(combo)
            combo.configure(values=self._resp_get_options())
        except Exception:
            pass

    def _refresh_task_combos(self):
        if self._mem_guard:
            return
        self._mem_guard = True
        try:
            vals = self._mem_get_tasks()
            for cb in list(self._task_combos):
                try:
                    cb.configure(values=vals)
                except Exception:
                    self._task_combos.remove(cb)
        finally:
            self._mem_guard = False

    def _refresh_field_combos(self):
        if self._mem_guard:
            return
        self._mem_guard = True
        try:
            vals = self._mem_get_fields()
            for cb in list(self._field_combos):
                try:
                    cb.configure(values=vals)
                except Exception:
                    self._field_combos.remove(cb)
        finally:
            self._mem_guard = False

    def _refresh_resp_combos(self):
        if self._mem_guard:
            return
        self._mem_guard = True
        try:
            vals = self._resp_get_options()
            for cb in list(self._resp_combos):
                try:
                    cb.configure(values=vals)
                except Exception:
                    self._resp_combos.remove(cb)
        finally:
            self._mem_guard = False

    def _mem_bind_combo_capture(self, combo, getter, bucket: str):
        try:
            entry = combo._entry
            if bucket == "task":
                entry.bind("<FocusOut>", lambda e: self._mem_add_task(getter()))
                entry.bind("<Return>",   lambda e: self._mem_add_task(getter()))
            else:
                entry.bind("<FocusOut>", lambda e: self._mem_add_field(getter()))
                entry.bind("<Return>",   lambda e: self._mem_add_field(getter()))
        except Exception:
            pass

        def _on_select(_=None):
            if self._mem_guard:
                return
            try:
                self._mem_guard = True
                val = getter()
                if bucket == "task":
                    self._mem_add_task(val)
                else:
                    self._mem_add_field(val)
            finally:
                self._mem_guard = False
        try:
            combo.configure(command=_on_select)
        except Exception:
            pass

    def _resp_add(self, s: str):
        v = self._norm(s)
        if not v or v == RESP_TEXT_FREE:
            return
        if any(v.lower() == x.lower() for x in self._resp_defaults):
            return
        self._resp_defaults.append(v)
        self._resp_defaults.sort(key=str.lower)
        self._refresh_resp_combos()

    def _resp_bind_combo_capture(self, combo, getter, on_select=None):
        try:
            entry = combo._entry
            entry.bind("<FocusOut>", lambda e: self._resp_add(getter()))
            entry.bind("<Return>",   lambda e: self._resp_add(getter()))
        except Exception:
            pass

        def _on_select(_=None):
            prev_guard = self._mem_guard
            if not prev_guard:
                self._mem_guard = True
            try:
                if not prev_guard:
                    self._resp_add(getter())
            finally:
                if not prev_guard:
                    self._mem_guard = False
            if callable(on_select):
                on_select()

        try:
            combo.configure(command=_on_select)
        except Exception:
            pass

    def _clear_memories(self):
        if messagebox.askyesno("Limpar memórias", "Limpar listas de TAREFAS, CAMPOS e RESPONSÁVEIS deste projeto?"):
            self._mem_tasks.clear()
            self._mem_fields.clear()
            self._resp_defaults = list(RESP_DEFAULTS)
            self._refresh_task_combos()
            self._refresh_field_combos()
            self._refresh_resp_combos()
    
    def _open_mem_manager(self):
        try:
            if self._mem_manager_window is not None and self._mem_manager_window.winfo_exists():
                self._mem_manager_window.focus()
                return
        except Exception:
            pass

        top = ctk.CTkToplevel(self)
        top.title("Gerenciador de Listas")
        try:
            _apply_dark_title_bar(top)
        except Exception:
            pass
        top.transient(self)
        top.grab_set()
        self._mem_manager_window = top
        
        def _on_close():
            self._mem_manager_window = None
            top.destroy()
        top.protocol("WM_DELETE_WINDOW", _on_close)

        tab_view = ctk.CTkTabview(top)
        tab_view.pack(expand=True, fill="both", padx=16, pady=16)

        tab_tasks = tab_view.add("Tarefas")
        tab_fields = tab_view.add("Campos")
        tab_resps = tab_view.add("Responsáveis")

        mgr_tasks = MemManagerTab(
            tab_tasks,
            title="Gerenciar Tarefas",
            mem_list=self._mem_tasks,
            refresh_cb=self._refresh_task_combos,
            layout="horizontal",
        )
        mgr_tasks.pack(expand=True, fill="both")

        mgr_fields = MemManagerTab(
            tab_fields,
            title="Gerenciar Campos",
            mem_list=self._mem_fields,
            refresh_cb=self._refresh_field_combos,
            layout="horizontal",
        )
        mgr_fields.pack(expand=True, fill="both")

        mgr_resps = MemManagerTab(
            tab_resps,
            title="Gerenciar Responsáveis padrão",
            mem_list=self._resp_defaults,
            refresh_cb=self._refresh_resp_combos,
            hint_text=_t("mem_resp_hint"),
            import_label=_t("mem_resp_import_label"),
            list_label=_t("mem_resp_label"),
            placeholder=_t("mem_resp_new_placeholder"),
            add_button_text=_t("mem_resp_add_button"),
            count_labels={
                "none": _t("mem_resp_none"),
                "one": _t("mem_resp_one"),
                "many": _t("mem_resp_many"),
            },
            forbidden_values=[RESP_TEXT_FREE],
            layout="horizontal",
        )
        mgr_resps.pack(expand=True, fill="both")

        _center_window(top, width=960, height=540, parent=self)

    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="we", padx=10, pady=(10, 6))
        header.grid_columnconfigure(0, weight=1)

        button_bar = ctk.CTkFrame(header, fg_color="transparent")
        button_bar.grid(row=0, column=0, sticky="we")

        actions_bar = ctk.CTkFrame(button_bar, fg_color="transparent")
        actions_bar.pack(side="left", fill="x", expand=True)
        for text, cmd, width in (
            ("Limpar construtor", self._clear_builder, 150),
            ("Gerenciar Listas", self._open_mem_manager, 160),
            ("Limpar memórias", self._clear_memories, 150),
            ("Reset geral", self._reset_all, 140),
        ):
            ctk.CTkButton(actions_bar, text=text, command=cmd, width=width).pack(side="left", padx=(0, 8), pady=4)

        more_bar = ctk.CTkFrame(button_bar, fg_color="transparent")
        more_bar.pack(side="right")
        ctk.CTkLabel(more_bar, text="Idioma:").pack(side="left", padx=(0, 6))
        ctk.CTkOptionMenu(
            more_bar,
            values=["Português", "Español"],
            variable=self._lang_var,
            width=120,
            command=self._on_change_lang,
        ).pack(side="left", padx=(0, 12))

        self._more_btn = ctk.CTkButton(more_bar, text="Mais...", command=self._open_more_menu, width=140)
        self._more_btn.pack(side="left")

        config_row = ctk.CTkFrame(header, fg_color="transparent")
        config_row.grid(row=1, column=0, sticky="we", pady=(6, 0))
        config_row.grid_columnconfigure(2, weight=1)

        start_frame = ctk.CTkFrame(config_row, fg_color="transparent")
        start_frame.grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(start_frame, text="RN inicial:").pack(side="left", padx=(0, 6))
        self.spin_start = IntSpin(start_frame, from_=1, to=999, width=110, variable=self.start_idx)
        self.spin_start.pack(side="left")
        def _sync_start(*_):
            try:
                value = int(self.start_idx.get())
            except Exception:
                value = 1
            self.spin_start.entry.delete(0, "end")
            self.spin_start.entry.insert(0, str(value))
        self.start_idx.trace_add("write", lambda *a: _sync_start())
        _sync_start()

        resp_frame = ctk.CTkFrame(config_row, fg_color="transparent")
        resp_frame.grid(row=0, column=1, sticky="w", padx=(20, 0))
        ctk.CTkLabel(resp_frame, text="Responsável padrão (novas ações):").pack(side="left")

        cb = ctk.CTkComboBox(
            resp_frame,
            values=self._resp_get_options(),
            variable=self.var_resp_preset,
            width=220,
            border_width=1,
            border_color=COMBO_BORDER,
        )
        cb.pack(side="left", padx=(6, 0))

        self.entry_resp_preset_free = ctk.CTkEntry(
            resp_frame, textvariable=self.var_resp_preset_free, width=220
        )

        def _toggle_preset_free(*_):
            if self.var_resp_preset.get() == RESP_TEXT_FREE:
                self.entry_resp_preset_free.pack(side="left", padx=(6, 0))
            else:
                self.var_resp_preset_free.set("")
                try:
                    self.entry_resp_preset_free.pack_forget()
                except Exception:
                    pass

        self._resp_register_combo(cb)
        self._resp_bind_combo_capture(cb, lambda: self.var_resp_preset.get(), on_select=_toggle_preset_free)
        _toggle_preset_free()

    def _on_change_lang(self, _choice: str):
        val = self._lang_var.get()
        set_lang("es" if val.startswith("Espa") else "pt")
        try:
            if hasattr(self, "_update_preview"):
                self._update_preview()
        except Exception:
            pass

    def _open_more_menu(self):
        try:
            menu = tk.Menu(self, tearoff=False)
            menu.add_command(label="Salvar projeto…", command=self._save_project)
            menu.add_command(label="Abrir projeto…", command=self._open_project)
            menu.add_separator()
            menu.add_command(label="Sobre", command=self._show_about)
            bx = self._more_btn.winfo_rootx()
            by = self._more_btn.winfo_rooty() + self._more_btn.winfo_height()
            try:
                menu.tk_popup(bx, by)
            finally:
                menu.grab_release()
        except Exception as e:
            messagebox.showerror("Menu", str(e))

    def _bind_shortcuts(self):
        try:
            self.bind_all("<Control-s>", lambda e: self._save_project())
            self.bind_all("<Control-o>", lambda e: self._open_project())
            self.bind_all("<Control-L>", lambda e: self._clear_builder())
            self.bind_all("<Control-l>", lambda e: self._clear_builder())
            self.bind_all("<Control-M>", lambda e: self._clear_memories())
            self.bind_all("<Control-m>", lambda e: self._clear_memories())
            self.bind_all("<Control-Shift-R>", lambda e: self._reset_all())
            self.bind_all("<F1>", lambda e: self._show_about())

            self.bind_all("<Control-Return>", lambda e: self._add_rn())
            self.bind_all("<Alt-c>", lambda e: self._add_cond())
            self.bind_all("<Alt-a>", lambda e: self._add_acao())

        except Exception:
            pass

    def _apply_initial_geometry(self):
        def _clamp(value: int, minimum: int, maximum: int) -> int:
            return max(minimum, min(value, maximum))

        self.update_idletasks()
        screen_w = max(self.winfo_screenwidth(), 1)
        screen_h = max(self.winfo_screenheight(), 1)

        margin_w = 40 if screen_w > 40 else 0
        margin_h = 80 if screen_h > 80 else 0

        width = _clamp(int(screen_w * 0.85), int(screen_w * 0.6), screen_w - margin_w)
        height = _clamp(int(screen_h * 0.85), int(screen_h * 0.6), screen_h - margin_h)

        x = max((screen_w - width) // 2, 0)
        y = max((screen_h - height) // 2, 0)

        self.geometry(f"{width}x{height}+{x}+{y}")

        min_w = min(width, max(int(screen_w * 0.55), 880))
        min_h = min(height, max(int(screen_h * 0.55), 640))
        self.minsize(min_w, min_h)

        try:
            self.attributes("-alpha", 1.0)
        except Exception:
            pass

    def _enable_scrollwheel(self, scrollable: ctk.CTkScrollableFrame):
        canvas = getattr(scrollable, "_parent_canvas", None)
        if canvas is None:
            return

        def _on_mousewheel(event):
            delta = getattr(event, "delta", 0)
            if delta == 0 and getattr(event, "num", None) in (4, 5):
                delta = 120 if event.num == 4 else -120
            step = -1 if delta > 0 else 1
            canvas.yview_scroll(step, "units")
            return "break"

        canvas.bind("<MouseWheel>", _on_mousewheel, add=True)
        canvas.bind("<Button-4>", _on_mousewheel, add=True)
        canvas.bind("<Button-5>", _on_mousewheel, add=True)

    def _show_about(self):
        messagebox.showinfo("Sobre", "Criado por Rodrigo Salvador Escudero")

    def _clear_header(self):
        self.start_idx.set(1)
        options = self._resp_get_options()
        default_resp = next((opt for opt in options if opt != RESP_TEXT_FREE), RESP_TEXT_FREE)
        self.var_resp_preset.set(default_resp)
        self.var_resp_preset_free.set("")
        try:
            self.entry_resp_preset_free.pack_forget()
        except Exception:
            pass

    def _clear_builder(self):
        try:
            if hasattr(self, "_destroy_rows"):
                self._destroy_rows(getattr(self, "cond_rows", []))
                self._destroy_rows(getattr(self, "acao_rows", []))
                if hasattr(self, "_ensure_min_builder_rows"):
                    self._ensure_min_builder_rows()
            if hasattr(self, "_update_preview"):
                self._update_preview()
        except Exception:
            pass

    def _reset_all(self):
        self._resp_defaults = list(RESP_DEFAULTS)
        self._refresh_resp_combos()
        self._clear_header()
        self._mem_tasks.clear(); self._mem_fields.clear()
        self._refresh_task_combos(); self._refresh_field_combos()

        self._clear_builder()

        self.flows = {"Fluxo Padrão": []}
        self.current_flow = "Fluxo Padrão"
        self.flow_var.set(self.current_flow)
        self._refresh_flow_controls()

        try:
            if hasattr(self, "_reset_builder_defaults"):
                self._reset_builder_defaults()
            if hasattr(self, "_clear_preview"):
                self._clear_preview()
            if hasattr(self, "_update_preview"):
                self._update_preview()
            if hasattr(self, "_ensure_min_builder_rows"):
                self._ensure_min_builder_rows()
            if hasattr(self, "_refresh_textbox"):
                self._refresh_textbox()
        except Exception:
            pass

    def _collect_project(self) -> dict:
        proj = {
            "version": 5,
            "lang": get_lang(),
            "header": {
                "start_idx": int(self.start_idx.get()),
                "resp_preset": self.var_resp_preset.get(),
                "resp_preset_free": self.var_resp_preset_free.get(),
            },
            "builder": {},
            "memory": {
                "tarefas": list(self._mem_tasks),
                "campos": list(self._mem_fields),
                "responsaveis": list(self._resp_defaults),
            },
            "flows": {k: list(v) for k, v in self.flows.items()},
            "rns": list(self.flows.get("Fluxo Padrão", [])),
        }
        try:
            if hasattr(self, "_collect_builder_into"):
                self._collect_builder_into(proj)
        except Exception:
            pass
        return proj

    def _apply_project(self, proj: dict):
        try:
            lang = proj.get("lang", "pt")
            set_lang(lang)
            self._lang_var.set("Español" if lang == "es" else "Português")

            header = proj.get("header", {})
            self.start_idx.set(int(header.get("start_idx", 1)))
            self.var_resp_preset.set(header.get("resp_preset", RESP_DEFAULTS[0] if RESP_DEFAULTS else RESP_TEXT_FREE))
            self.var_resp_preset_free.set(header.get("resp_preset_free", ""))
            if self.var_resp_preset.get() == RESP_TEXT_FREE:
                self.entry_resp_preset_free.pack(side="left", padx=(6, 0))
            else:
                try:
                    self.entry_resp_preset_free.pack_forget()
                except Exception:
                    pass

            mem = proj.get("memory", {})
            self._mem_tasks = list(mem.get("tarefas", []))
            self._mem_fields = list(mem.get("campos", []))
            resp_list = mem.get("responsaveis", RESP_DEFAULTS)
            self._resp_defaults = list(resp_list) if resp_list else list(RESP_DEFAULTS)
            self._refresh_task_combos(); self._refresh_field_combos(); self._refresh_resp_combos()

            preset_value = self.var_resp_preset.get()
            if preset_value and preset_value not in self._resp_get_options() and preset_value != RESP_TEXT_FREE:
                self._resp_add(preset_value)
            elif preset_value == RESP_TEXT_FREE and not self.var_resp_preset_free.get():
                self.var_resp_preset_free.set("")

            try:
                if hasattr(self, "_apply_builder_from"):
                    self._apply_builder_from(proj)
            except Exception:
                pass

            flows_data = proj.get("flows")
            if isinstance(flows_data, dict) and flows_data:
                self.flows = {k: list(v) for k, v in flows_data.items()}
            else:
                legacy = proj.get("rns", [])
                self.flows = {"Fluxo Padrão": list(legacy) if isinstance(legacy, list) else []}
            self.current_flow = next(iter(self.flows.keys()), "Fluxo Padrão")
            self.flow_var.set(self.current_flow)
            self._refresh_flow_controls()
            try:
                if hasattr(self, "_refresh_textbox"):
                    self._refresh_textbox()
            except Exception:
                pass

            try:
                if hasattr(self, "_update_preview"):
                    self._update_preview()
            except Exception:
                pass

        except Exception as e:
            messagebox.showerror("Erro ao carregar projeto", str(e))

    def _save_project(self):
        proj = self._collect_project()
        path = filedialog.asksaveasfilename(
            title="Salvar projeto",
            defaultextension=".rnproj",
            filetypes=[("Projeto de RNs", ".rnproj"), ("JSON", ".json"), ("Todos", "*.*")],
            initialfile="meu_projeto.rnproj",
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(proj, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Projeto salvo", f"Projeto salvo em:\n{path}")
        except Exception as e:
            messagebox.showerror("Erro ao salvar projeto", str(e))

    def _open_project(self):
        path = filedialog.askopenfilename(
            title="Abrir projeto",
            filetypes=[("Projeto de RNs", ".rnproj"), ("JSON", ".json"), ("Todos", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                proj = json.load(f)
            self._apply_project(proj)
        except Exception as e:
            messagebox.showerror("Erro ao abrir projeto", str(e))

def _attach_builder_to_RNBuilder():
    def _build_rule(self: 'RNBuilder'):
        parent = self.left_pane
        self.builder_scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        self.builder_scroll.pack(fill="both", expand=True, padx=(10, 6), pady=6)
        self.builder_scroll.grid_columnconfigure(0, weight=1)
        self._enable_scrollwheel(self.builder_scroll)

        self.builder_container = scrollable_body(self.builder_scroll)
        self.builder_container.grid_columnconfigure(0, weight=1)
        self.builder_container.grid_rowconfigure(0, weight=4)
        self.builder_container.grid_rowconfigure(1, weight=3)
        self.builder_container.grid_rowconfigure(2, weight=5)

        self.frm_gatilho_collapsible = CollapsibleGroup(
            self.builder_container,
            text="Gatilho",
            start_expanded=True
        )
        self.frm_gatilho_collapsible.grid(row=0, column=0, padx=(0,0), pady=(0,10), sticky="nsew")
        frm_gatilho_group = self.frm_gatilho_collapsible.get_inner_frame()
        frm_gatilho_group.grid_columnconfigure(0, weight=1)

        self.var_gatilho_tipo = tk.StringVar(value=GATILHOS[1])
        gtbar = ctk.CTkFrame(frm_gatilho_group, fg_color="transparent")
        gtbar.grid(row=0, column=0, padx=6, pady=6, sticky="ew")
        gtbar.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(gtbar, text="Gatilho:").grid(row=0, column=0, sticky="w")
        ctk.CTkComboBox(
            gtbar,
            values=GATILHOS,
            variable=self.var_gatilho_tipo,
            command=lambda *_: self._refresh_gatilho_fields(),
        ).grid(row=0, column=1, sticky="ew", padx=(6, 0))

        self.var_obj = tk.StringVar(value="")
        self.var_tarefa_ctx = tk.StringVar(value="")
        self.var_campo = tk.StringVar(value="")
        self.var_resposta = tk.StringVar(value="")
        self.var_tarefa_done = tk.StringVar(value="")
        self.var_evento = tk.StringVar(value="")

        self.frm_gatilho = ctk.CTkFrame(frm_gatilho_group, fg_color="transparent")
        self.frm_gatilho.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 6))
        self.frm_gatilho.grid_columnconfigure(0, weight=0, minsize=120)
        self.frm_gatilho.grid_columnconfigure(1, weight=1)
        self._refresh_gatilho_fields()

        self.frm_cond_collapsible = CollapsibleGroup(
            self.builder_container,
            text="Condições adicionais",
            start_expanded=True
        )
        self.frm_cond_collapsible.grid(row=1, column=0, padx=(0,0), pady=(0,10), sticky="nsew")
        frm_cond_group = self.frm_cond_collapsible.get_inner_frame()
        frm_cond_group.grid_columnconfigure(0, weight=1)

        cond_hdr = ctk.CTkFrame(frm_cond_group, fg_color="transparent")
        cond_hdr.grid(row=0, column=0, padx=6, pady=(6, 2), sticky="ew")
        cond_hdr.grid_columnconfigure(0, weight=1)
        cond_hdr.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(cond_hdr, text="Condições adicionais:").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        self.var_conj = tk.StringVar(value="E")
        ctk.CTkRadioButton(
            cond_hdr,
            text="Todas (E)",
            variable=self.var_conj,
            value="E",
            command=self._update_preview,
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        ctk.CTkRadioButton(
            cond_hdr,
            text="Qualquer (OU)",
            variable=self.var_conj,
            value="OU",
            command=self._update_preview,
        ).grid(row=1, column=1, sticky="w", pady=(4, 0), padx=(12, 0))

        self.frm_conds = ctk.CTkFrame(frm_cond_group, fg_color="transparent")
        self.frm_conds.grid(row=1, column=0, sticky="ew", padx=6)
        self.frm_conds.grid_columnconfigure(0, weight=1)
        self.frm_conds_body = self.frm_conds
        self.frm_conds_body.grid_columnconfigure(0, weight=1)
        self.cond_rows = []

        cond_btnbar = ctk.CTkFrame(frm_cond_group, fg_color="transparent")
        cond_btnbar.grid(row=2, column=0, padx=6, pady=(4, 6), sticky="ew")
        cond_btnbar.grid_columnconfigure(0, weight=1)
        cond_btnbar.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(
            cond_btnbar,
            text="+ Adicionar condição",
            command=self._add_cond,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ctk.CTkButton(
            cond_btnbar,
            text="Limpar condições",
            command=self._clear_conditions,
        ).grid(row=0, column=1, sticky="ew")

        self.frm_acao_collapsible = CollapsibleGroup(
            self.builder_container,
            text="Ações",
            start_expanded=True
        )
        self.frm_acao_collapsible.grid(row=2, column=0, padx=(0,0), pady=(0,0), sticky="nsew")
        frm_acao_group = self.frm_acao_collapsible.get_inner_frame()
        frm_acao_group.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(frm_acao_group, text="Ações:").grid(row=0, column=0, padx=6, pady=(6,2), sticky="w")
        self.frm_acoes = ctk.CTkFrame(frm_acao_group, fg_color="transparent")
        self.frm_acoes.grid(row=1, column=0, sticky="ew", padx=6)
        self.frm_acoes.grid_columnconfigure(0, weight=1)
        self.frm_acoes_body = self.frm_acoes
        self.frm_acoes_body.grid_columnconfigure(0, weight=1)
        self.acao_rows = []

        acoes_btnbar = ctk.CTkFrame(frm_acao_group, fg_color="transparent")
        acoes_btnbar.grid(row=2, column=0, padx=6, pady=(4, 2), sticky="ew")
        acoes_btnbar.grid_columnconfigure(0, weight=1)
        acoes_btnbar.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(
            acoes_btnbar,
            text="+ Adicionar ação",
            command=self._add_acao,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ctk.CTkButton(
            acoes_btnbar,
            text="Limpar ações",
            command=self._clear_actions,
        ).grid(row=0, column=1, sticky="ew")

        freq_bar = ctk.CTkFrame(frm_acao_group, fg_color="transparent")
        freq_bar.grid(row=3, column=0, padx=6, pady=(8, 4), sticky="ew")
        freq_bar.grid_columnconfigure(0, weight=1)
        freq_bar.grid_columnconfigure(1, weight=0)
        ctk.CTkLabel(freq_bar, text="Ações frequentes — Acionar fluxo:").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        self.var_freq = tk.StringVar(value="Cadastro")
        flow_combo = ctk.CTkComboBox(
            freq_bar,
            values=[
                "Cadastro", "Acordo", "Decisão", "Réplica", "Garantias", "Bloqueio",
                "Penhora", "Encerramento", "Reativação", "Perícias", "Obrigações",
            ],
            variable=self.var_freq,
        )
        flow_combo.grid(row=1, column=0, sticky="ew", pady=(4, 0), padx=(0, 6))
        ctk.CTkButton(
            freq_bar,
            text="Inserir",
            command=self._insert_frequent_flow,
        ).grid(row=1, column=1, sticky="ew", pady=(4, 0))

        freq_ret = ctk.CTkFrame(frm_acao_group, fg_color="transparent")
        freq_ret.grid(row=4, column=0, padx=6, pady=(8, 4), sticky="ew")
        freq_ret.grid_columnconfigure(0, weight=2)
        freq_ret.grid_columnconfigure(1, weight=1)
        freq_ret.grid_columnconfigure(2, weight=1)
        
        ctk.CTkLabel(freq_ret, text="Retornar a tarefa:").grid(row=0, column=0, columnspan=3, sticky="w")
        
        self.var_freq_ret = tk.StringVar(value="")
        cb_ret = ctk.CTkComboBox(freq_ret, values=self._mem_get_tasks(), variable=self.var_freq_ret)
        cb_ret.grid(row=1, column=0, sticky="ew", pady=(4, 0), padx=(0, 6))
        self._mem_register_task_combo(cb_ret)
        self._mem_bind_combo_capture(cb_ret, lambda: self.var_freq_ret.get(), bucket="task")
        
        self.var_freq_ret_restart = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            freq_ret,
            text="Reiniciar SLA",
            variable=self.var_freq_ret_restart,
            onvalue=True,
            offvalue=False,
            command=self._update_preview,
        ).grid(row=1, column=1, sticky="w", pady=(4, 0), padx=(6,6))
        
        btn_ret = ctk.CTkButton(freq_ret, text="Inserir", command=self._insert_frequent_return)
        btn_ret.grid(row=1, column=2, sticky="ew", pady=(4, 0))

        freq_close = ctk.CTkFrame(frm_acao_group, fg_color="transparent")
        freq_close.grid(row=5, column=0, padx=6, pady=(8, 4), sticky="ew")
        freq_close.grid_columnconfigure(0, weight=0)
        freq_close.grid_columnconfigure(1, weight=1)
        freq_close.grid_columnconfigure(2, weight=1)
        
        ctk.CTkLabel(freq_close, text="Encerramento:").grid(row=0, column=0, sticky="w", padx=(0, 6))
        
        ctk.CTkButton(
            freq_close, text="Parcial",
            command=lambda: self._insert_frequent_close(parcial=True)
        ).grid(row=0, column=1, sticky="ew", padx=(0, 6))
        
        ctk.CTkButton(
            freq_close, text="Total",
            command=lambda: self._insert_frequent_close(parcial=False)
        ).grid(row=0, column=2, sticky="ew")

        freq_cond = ctk.CTkFrame(frm_acao_group, fg_color="transparent")
        freq_cond.grid(row=6, column=0, padx=6, pady=(8, 6), sticky="ew")
        
        freq_cond.grid_columnconfigure(0, weight=0)
        freq_cond.grid_columnconfigure(1, weight=3)
        freq_cond.grid_columnconfigure(2, weight=0)
        freq_cond.grid_columnconfigure(3, weight=2)

        ctk.CTkLabel(freq_cond, text="Condição rápida (Campo / Resposta):").grid(
            row=0, column=0, columnspan=4, sticky="w"
        )
        
        ctk.CTkLabel(freq_cond, text="Campo:").grid(row=1, column=0, sticky="w", pady=(4, 0), padx=(0, 6))
        self.var_freq_cond_field = tk.StringVar(value="")
        cb_field = ctk.CTkComboBox(
            freq_cond,
            values=self._mem_get_fields(),
            variable=self.var_freq_cond_field,
        )
        cb_field.grid(row=1, column=1, sticky="ew", pady=(4, 0))
        self._mem_register_field_combo(cb_field)
        self._mem_bind_combo_capture(cb_field, lambda: self.var_freq_cond_field.get(), bucket="field")

        ctk.CTkLabel(freq_cond, text="Resposta:").grid(row=1, column=2, sticky="w", pady=(4, 0), padx=(10, 6))
        self.var_freq_cond_resp = tk.StringVar(value="")
        ent_resp = ctk.CTkEntry(freq_cond, textvariable=self.var_freq_cond_resp)
        ent_resp.grid(row=1, column=3, sticky="ew", pady=(4, 0))
        
        ent_resp.bind("<Return>", lambda e: self._insert_frequent_condition())
        ent_resp.bind("<FocusOut>", lambda e: self._update_preview())

        ctk.CTkButton(
            freq_cond,
            text="Adicionar",
            command=self._insert_frequent_condition,
        ).grid(row=2, column=0, columnspan=4, sticky="ew", pady=(8, 0))


        if not self.cond_rows:
            self._add_cond()
        if not self.acao_rows:
            self._add_acao()

        for v in (
            self.var_gatilho_tipo, self.var_obj, self.var_tarefa_ctx, self.var_campo,
            self.var_resposta, self.var_tarefa_done, self.var_evento, self.var_conj
        ):
            v.trace_add("write", lambda *a: self._update_preview())

    def _reset_builder_defaults(self: 'RNBuilder'):
        self.var_gatilho_tipo.set(GATILHOS[1])
        self.var_obj.set("")
        self.var_tarefa_ctx.set("")
        self.var_campo.set("")
        self.var_resposta.set("")
        self.var_tarefa_done.set("")
        self.var_evento.set("")
        try:
            self.var_conj.set("E")
        except Exception:
            pass
        self._refresh_gatilho_fields()

    def _reset_for_next_rule(self: 'RNBuilder'):
        self._reset_builder_defaults()
        self._destroy_rows(getattr(self, 'cond_rows', []))
        self._destroy_rows(getattr(self, 'acao_rows', []))
        self._ensure_min_builder_rows()
        self._update_preview()
        try:
            self.builder_scroll._parent_canvas.yview_moveto(0)
        except Exception:
            pass

    def _toggle_yes_no(self: 'RNBuilder', value: str) -> str:
        lang = get_lang()
        v = (value or "").strip()
        base = v.lower()
        if lang == 'es':
            norm = base.replace('í', 'i')
            if norm == 'si':
                return 'No'
            if norm == 'no':
                return 'Sí'
        else:
            if base == 'sim':
                return 'Não'
            if base in ('não', 'nao'):
                return 'Sim'
        return value

    def _invert_yes_no_conditions(self: 'RNBuilder'):
        if self.var_gatilho_tipo.get() != GATILHOS[1]:
            return
        try:
            new_val = self._toggle_yes_no(self.var_resposta.get())
            self.var_resposta.set(new_val)
        except Exception:
            pass

        for row in getattr(self, 'cond_rows', []):
            try:
                cur = row.var_valor.get()
                toggled = self._toggle_yes_no(cur)
                if toggled != cur:
                    row.var_valor.set(toggled)
            except Exception:
                continue

    def _destroy_rows(self: 'RNBuilder', rows):
        for r in list(rows):
            try:
                r.destroy()
            except Exception:
                pass
        rows.clear()

    def _relayout_cond_rows(self: 'RNBuilder'):
        for idx, widget in enumerate(getattr(self, 'cond_rows', [])):
            try:
                widget.grid(row=idx, column=0, sticky="ew", padx=0, pady=(0, 6))
            except Exception:
                pass

    def _relayout_acao_rows(self: 'RNBuilder'):
        for idx, widget in enumerate(getattr(self, 'acao_rows', [])):
            try:
                widget.grid(row=idx, column=0, sticky="ew", padx=0, pady=(0, 6))
            except Exception:
                pass

    def _ensure_min_builder_rows(self: 'RNBuilder'):
        try:
            if hasattr(self, 'frm_conds_body') and not getattr(self, 'cond_rows', []):
                self._add_cond()
            if hasattr(self, 'frm_acoes_body') and not getattr(self, 'acao_rows', []):
                self._add_acao()
        except Exception:
            pass

    def _clear_conditions(self: 'RNBuilder', *, confirm=True):
        if not getattr(self, 'cond_rows', []):
            if hasattr(self, 'frm_conds_body'):
                self._add_cond()
            return
        if (not confirm) or messagebox.askyesno("Limpar condições", "Remover todas as condições?"):
            self._destroy_rows(self.cond_rows)
            if hasattr(self, 'frm_conds_body'):
                self._add_cond()

    def _clear_actions(self: 'RNBuilder', *, confirm=True):
        if not getattr(self, 'acao_rows', []):
            if hasattr(self, 'frm_acoes_body'):
                self._add_acao()
            return
        if (not confirm) or messagebox.askyesno("Limpar ações", "Remover todas as ações?"):
            self._destroy_rows(self.acao_rows)
            if hasattr(self, 'frm_acoes_body'):
                self._add_acao()

    def _add_cond(self: 'RNBuilder'):
        def _on_remove(row):
            if row in self.cond_rows:
                self.cond_rows.remove(row)
                self._relayout_cond_rows()
            self.after_idle(self._ensure_min_builder_rows)

        row = LinhaCondicao(
            self.frm_conds_body,
            on_change=self._update_preview,
            on_remove=_on_remove,
        )
        self.cond_rows.append(row)
        self._relayout_cond_rows()
        self._update_preview()
        self.after(100, self.builder_scroll._parent_canvas.yview_moveto, 1.0)

    def _add_acao(self: 'RNBuilder'):
        preset_value = (self.var_resp_preset_free.get().strip() if self.var_resp_preset.get() == RESP_TEXT_FREE
                        else self.var_resp_preset.get())
        def _on_remove(row):
            if row in self.acao_rows:
                self.acao_rows.remove(row)
                self._relayout_acao_rows()
            self.after_idle(self._ensure_min_builder_rows)

        row = LinhaAcao(self.frm_acoes_body, on_change=self._update_preview,
                        on_remove=_on_remove,
                        default_resp=preset_value,
                        default_resp_free=(self.var_resp_preset_free.get() if self.var_resp_preset.get() == RESP_TEXT_FREE else ""),
                        row_list_key="acao_rows")
        self.acao_rows.append(row)
        self._relayout_acao_rows()
        self._update_preview()
        self.after(100, self.builder_scroll._parent_canvas.yview_moveto, 1.0)

    def _insert_frequent_flow(self: 'RNBuilder'):
        try:
            nome = (self.var_freq.get() or '').strip()
            if not nome:
                return
            preset_value = (self.var_resp_preset_free.get().strip() if self.var_resp_preset.get() == RESP_TEXT_FREE
                            else self.var_resp_preset.get())
            def _remove(row):
                if row in self.acao_rows:
                    self.acao_rows.remove(row)
                    self._relayout_acao_rows()
                    self.after_idle(self._ensure_min_builder_rows)

            row = LinhaAcao(self.frm_acoes_body, on_change=self._update_preview,
                            on_remove=_remove,
                            default_resp=preset_value,
                            default_resp_free=(self.var_resp_preset_free.get() if self.var_resp_preset.get() == RESP_TEXT_FREE else ""),
                            row_list_key="acao_rows")
            row.var_tipo.set("Acionar Fluxo"); row._refresh(); row.var_fluxo.set(nome)
            self.acao_rows.append(row)
            self._relayout_acao_rows()
            self._update_preview()
            self.var_freq.set("Cadastro")
            self.after(100, self.builder_scroll._parent_canvas.yview_moveto, 1.0)
        except Exception as e:
            messagebox.showerror("Falha ao inserir ação frequente", str(e))

    def _insert_frequent_return(self: 'RNBuilder'):
        try:
            tarefa = (self.var_freq_ret.get() or '').strip()
            if not tarefa:
                return
            self._mem_add_task(tarefa)
            preset_value = (self.var_resp_preset_free.get().strip() if self.var_resp_preset.get() == RESP_TEXT_FREE
                            else self.var_resp_preset.get())
            def _remove(row):
                if row in self.acao_rows:
                    self.acao_rows.remove(row)
                    self._relayout_acao_rows()
                    self.after_idle(self._ensure_min_builder_rows)

            row = LinhaAcao(self.frm_acoes_body, on_change=self._update_preview,
                            on_remove=_remove,
                            default_resp=preset_value,
                            default_resp_free=(self.var_resp_preset_free.get() if self.var_resp_preset.get() == RESP_TEXT_FREE else ""),
                            row_list_key="acao_rows")
            row.var_tipo.set("Retornar a Tarefa"); row._refresh()
            row.var_ret_tarefa.set(tarefa)
            row.var_ret_restart.set(bool(self.var_freq_ret_restart.get()))
            self.acao_rows.append(row)
            self._relayout_acao_rows()
            self._update_preview()
            self.var_freq_ret.set("")
            self.after(100, self.builder_scroll._parent_canvas.yview_moveto, 1.0)
        except Exception as e:
            messagebox.showerror("Falha ao inserir 'Retornar a tarefa'", str(e))

    def _insert_frequent_condition(self: 'RNBuilder'):
        try:
            campo = (self.var_freq_cond_field.get() or '').strip()
            resp  = (self.var_freq_cond_resp.get() or '').strip()
            if not campo or not resp:
                messagebox.showwarning("Condição incompleta", "Informe o Campo e a Resposta.")
                return
            self._mem_add_field(campo)
            def _remove(row):
                if row in self.cond_rows:
                    self.cond_rows.remove(row)
                    self._relayout_cond_rows()
                    self.after_idle(self._ensure_min_builder_rows)

            row = LinhaCondicao(self.frm_conds_body, on_change=self._update_preview,
                                on_remove=_remove)
            row.var_campo.set(campo)
            row.var_op.set("for respondido como")
            row.var_valor.set(resp)
            self.cond_rows.append(row)
            self._relayout_cond_rows()
            self._update_preview()
            self.var_freq_cond_field.set("")
            self.var_freq_cond_resp.set("")
            self.after(100, self.builder_scroll._parent_canvas.yview_moveto, 1.0)
        except Exception as e:
            messagebox.showerror("Falha ao inserir condição rápida", str(e))

    def _insert_frequent_close(self: 'RNBuilder', *, parcial: bool):
        try:
            preset_value = (self.var_resp_preset_free.get().strip() if self.var_resp_preset.get() == RESP_TEXT_FREE
                            else self.var_resp_preset.get())
            def _remove(row):
                if row in self.acao_rows:
                    self.acao_rows.remove(row)
                    self._relayout_acao_rows()
                    self.after_idle(self._ensure_min_builder_rows)

            row = LinhaAcao(self.frm_acoes_body, on_change=self._update_preview,
                            on_remove=_remove,
                            default_resp=preset_value,
                            default_resp_free=(self.var_resp_preset_free.get() if self.var_resp_preset.get() == RESP_TEXT_FREE else ""),
                            row_list_key="acao_rows")
            
            tipo = "Encerrar Fluxo (Parcial)" if parcial else "Encerrar Fluxo (Total)"
            row.var_tipo.set(tipo); row._refresh()
            
            self.acao_rows.append(row)
            self._relayout_acao_rows()
            self._update_preview()
            self.after(100, self.builder_scroll._parent_canvas.yview_moveto, 1.0)
        except Exception as e:
            messagebox.showerror("Falha ao inserir ação de encerramento", str(e))

    def _refresh_gatilho_fields(self: 'RNBuilder'):
        for w in self.frm_gatilho.winfo_children():
            w.destroy()
        
        t = self.var_gatilho_tipo.get()
        if t == "Sempre que inserido novo OBJETO":
            ctk.CTkLabel(self.frm_gatilho, text="Objeto:").grid(
                row=0, column=0, padx=(0, 6), pady=3, sticky="e"
            )
            ctk.CTkEntry(self.frm_gatilho, textvariable=self.var_obj).grid(
                row=0, column=1, padx=(0, 0), pady=3, sticky="ew"
            )
        elif t == "Em TAREFA se CAMPO for RESPOSTA":
            ctk.CTkLabel(self.frm_gatilho, text="Tarefa:").grid(
                row=0, column=0, padx=(0, 6), pady=3, sticky="e"
            )
            cb_t = ctk.CTkComboBox(
                self.frm_gatilho,
                values=self._mem_get_tasks(),
                variable=self.var_tarefa_ctx,
            )
            cb_t.grid(row=0, column=1, padx=0, pady=3, sticky="ew")
            self._mem_register_task_combo(cb_t)
            self._mem_bind_combo_capture(cb_t, lambda: self.var_tarefa_ctx.get(), "task")

            ctk.CTkLabel(self.frm_gatilho, text="Campo:").grid(
                row=1, column=0, padx=(0, 6), pady=3, sticky="e"
            )
            cb_c = ctk.CTkComboBox(
                self.frm_gatilho,
                values=self._mem_get_fields(),
                variable=self.var_campo,
            )
            cb_c.grid(row=1, column=1, padx=0, pady=3, sticky="ew")
            self._mem_register_field_combo(cb_c)
            self._mem_bind_combo_capture(cb_c, lambda: self.var_campo.get(), "field")

            ctk.CTkLabel(self.frm_gatilho, text="Resposta:").grid(
                row=2, column=0, padx=(0, 6), pady=3, sticky="e"
            )
            ent_r = ctk.CTkEntry(self.frm_gatilho, textvariable=self.var_resposta)
            ent_r.grid(row=2, column=1, padx=0, pady=3, sticky="ew")
            ent_r.bind("<Return>", lambda e: self._update_preview())
            ent_r.bind("<FocusOut>", lambda e: self._update_preview())
        elif t == "Concluída TAREFA":
            ctk.CTkLabel(self.frm_gatilho, text="Tarefa concluída:").grid(
                row=0, column=0, padx=(0, 6), pady=3, sticky="e"
            )
            cb_done = ctk.CTkComboBox(
                self.frm_gatilho,
                values=self._mem_get_tasks(),
                variable=self.var_tarefa_done,
            )
            cb_done.grid(row=0, column=1, padx=0, pady=3, sticky="ew")
            self._mem_register_task_combo(cb_done)
            self._mem_bind_combo_capture(cb_done, lambda: self.var_tarefa_done.get(), "task")
        else:
            ctk.CTkLabel(self.frm_gatilho, text="Evento:").grid(
                row=0, column=0, padx=(0, 6), pady=3, sticky="e"
            )
            ctk.CTkEntry(self.frm_gatilho, textvariable=self.var_evento).grid(
                row=0, column=1, padx=0, pady=3, sticky="ew"
            )
        self.after(0, self._update_preview)

    def _when_text(self: 'RNBuilder') -> str:
        lang = get_lang()
        t = self.var_gatilho_tipo.get()

        if t == GATILHOS[0]:
            obj = self.var_obj.get().strip()
            if not obj:
                return ""
            return _t("when_new_object").format(obj=f"{CUR_L}{obj}{CUR_R}")

        if t == GATILHOS[1]:
            task = self.var_tarefa_ctx.get().strip()
            field = self.var_campo.get().strip()
            answer = self.var_resposta.get().strip()
            if not (task and field and answer):
                return ""
            return _t("when_task_field_answer").format(
                task=f"{CUR_L}{task}{CUR_R}",
                field=f"{CUR_L}{field}{CUR_R}",
                answer=f"{CUR_L}{answer}{CUR_R}"
            )

        if t == GATILHOS[2]:
            done = self.var_tarefa_done.get().strip()
            if not done:
                return ""
            return _t("when_task_done").format(task=f"{CUR_L}{done}{CUR_R}")

        event = self.var_evento.get().strip()
        if not event:
            return ""
        return _t("when_after_event").format(event=f"{CUR_L}{event}{CUR_R}")

    def _cond_text(self: 'RNBuilder') -> str:
        texts = [r.to_text() for r in getattr(self, 'cond_rows', [])]
        texts = [t for t in texts if t]
        conj = getattr(self, 'var_conj', tk.StringVar(value='E')).get()
        return _join_conditions(texts, conj) if texts else ""

    def _acoes_text(self: 'RNBuilder', rows) -> str:
        parts = [r.to_text() for r in rows]
        parts = [p for p in parts if p]
        if not parts:
            return ""
        out = parts[0]
        for p in parts[1:]:
            out += f"; e {p}"
        return out

    def _move_row(self: 'RNBuilder', widget, delta, row_list, frame_parent):
        try:
            idx = row_list.index(widget)
        except ValueError:
            return

        new_idx = idx + delta
        if not (0 <= new_idx < len(row_list)):
            return

        row_list[idx], row_list[new_idx] = row_list[new_idx], row_list[idx]

        if row_list and isinstance(row_list[0], LinhaCondicao):
            self._relayout_cond_rows()
        else:
            self._relayout_acao_rows()

        self._update_preview()

    RNBuilder._build_rule = _build_rule
    RNBuilder._reset_builder_defaults = _reset_builder_defaults
    RNBuilder._destroy_rows = _destroy_rows
    RNBuilder._relayout_cond_rows = _relayout_cond_rows
    RNBuilder._relayout_acao_rows = _relayout_acao_rows
    RNBuilder._ensure_min_builder_rows = _ensure_min_builder_rows
    RNBuilder._clear_conditions = _clear_conditions
    RNBuilder._clear_actions = _clear_actions
    RNBuilder._add_cond = _add_cond
    RNBuilder._add_acao = _add_acao
    RNBuilder._insert_frequent_flow = _insert_frequent_flow
    RNBuilder._insert_frequent_return = _insert_frequent_return
    RNBuilder._insert_frequent_condition = _insert_frequent_condition
    RNBuilder._insert_frequent_close = _insert_frequent_close
    RNBuilder._refresh_gatilho_fields = _refresh_gatilho_fields
    RNBuilder._when_text = _when_text
    RNBuilder._cond_text = _cond_text
    RNBuilder._acoes_text = _acoes_text
    RNBuilder._move_row = _move_row

_attach_builder_to_RNBuilder()


def _attach_panels_to_RNBuilder():
    def _build_panels(self: 'RNBuilder'):
        parent = self.right_pane
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        panel_scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        panel_scroll.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=6)
        panel_scroll.grid_columnconfigure(0, weight=1)
        panel_body = scrollable_body(panel_scroll)
        panel_body.grid_columnconfigure(0, weight=1)
        panel_body.grid_rowconfigure(0, weight=4)
        panel_body.grid_rowconfigure(1, weight=5)

        self.preview_collapsible = CollapsibleGroup(
            panel_body,
            text="Pré-visualização da RN atual",
            start_expanded=True
        )
        self.preview_collapsible.grid(row=0, column=0, padx=(0,0), pady=(0,8), sticky="nsew")
        preview_group = self.preview_collapsible.get_inner_frame()
        preview_group.grid_rowconfigure(0, weight=1)
        preview_group.grid_columnconfigure(0, weight=1)

        self.prev_box = SafeCTkTextbox(preview_group)
        self.prev_box.grid(row=0, column=0, sticky="nsew", padx=0, pady=(0, 8))

        try:
            self.prev_box.configure(wrap="word")
        except Exception:
            pass

        left_btnbar = ctk.CTkFrame(preview_group, fg_color="transparent")
        left_btnbar.grid(row=1, column=0, sticky="w", padx=0, pady=(0, 2))
        ctk.CTkButton(left_btnbar, text="Adicionar RN", command=self._add_rn, width=160).pack(side="left", padx=(0, 6))
        ctk.CTkButton(left_btnbar, text="Adicionar RN e preparar oposto (Sim/Não)",
                      command=self._add_rn_and_prepare_opposite, width=300).pack(side="left", padx=(0, 6))
        ctk.CTkButton(left_btnbar, text="Limpar pré-visualização", command=self._clear_preview, width=200).pack(side="left")

        self.rn_collapsible = CollapsibleGroup(
            panel_body,
            text="RNs (lista final)",
            start_expanded=True
        )
        self.rn_collapsible.grid(row=1, column=0, padx=(0,0), pady=(0,0), sticky="nsew")
        rn_group = self.rn_collapsible.get_inner_frame()
        rn_group.grid_rowconfigure(2, weight=2)
        rn_group.grid_rowconfigure(3, weight=3)
        rn_group.grid_columnconfigure(0, weight=1)

        flow_bar = ctk.CTkFrame(rn_group, fg_color="transparent")
        flow_bar.grid(row=0, column=0, sticky="ew", padx=0, pady=(0, 6))
        flow_bar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(flow_bar, text="Fluxo:", font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, sticky="w", padx=(0, 8))

        self.flow_combo = ctk.CTkComboBox(
            flow_bar,
            variable=self.flow_var,
            values=self._get_flow_names(),
            command=self._on_flow_selected,
            width=220,
        )
        self.flow_combo.grid(row=0, column=1, sticky="w", pady=2)

        flow_btns = ctk.CTkFrame(flow_bar, fg_color="transparent")
        flow_btns.grid(row=0, column=2, sticky="e", padx=(8, 0))
        ctk.CTkButton(flow_btns, text="Novo", width=70, command=self._new_flow).pack(side="left", padx=(0, 6))
        ctk.CTkButton(flow_btns, text="Renomear", width=90, command=self._rename_flow).pack(side="left", padx=(0, 6))
        ctk.CTkButton(flow_btns, text="Excluir", width=80, command=self._delete_flow, fg_color=DANGER_BG, hover_color=DANGER_HOVER).pack(side="left")

        right_btnbar = ctk.CTkFrame(rn_group, fg_color="transparent")
        right_btnbar.grid(row=1, column=0, sticky="w", padx=0, pady=(0, 6))
        ctk.CTkButton(right_btnbar, text="Copiar RN", command=self._copy_single_rn, width=120).pack(side="left", padx=(0, 6))
        ctk.CTkButton(right_btnbar, text="Copiar tudo", command=self._copy_all, width=110).pack(side="left", padx=(0, 6))
        ctk.CTkButton(right_btnbar, text="Salvar .txt", command=self._save_txt, width=110).pack(side="left", padx=(0, 6))
        ctk.CTkButton(right_btnbar, text="Limpar RNs", command=lambda: self._clear_rns(confirm=True), width=110).pack(side="left")

        self.rn_mgr = ctk.CTkScrollableFrame(rn_group)
        self.rn_mgr.grid(row=2, column=0, sticky="nsew", padx=0, pady=(0, 6))
        self.rn_mgr.grid_columnconfigure(0, weight=1)
        self._enable_scrollwheel(self.rn_mgr)

        self.txt = ctk.CTkTextbox(rn_group)
        self.txt.grid(row=3, column=0, sticky="nsew", padx=0, pady=0)
        try:
            self.txt.configure(wrap="word")
        except Exception:
            pass

        self._refresh_flow_controls()

    def _set_preview_text(self: 'RNBuilder', txt: str):
        try:
            self.prev_box.configure(state="normal")
            self.prev_box.delete("1.0", "end")
            self.prev_box.insert("1.0", txt)
            self.prev_box.configure(state="disabled")
        except Exception:
            pass

    def _clear_preview(self: 'RNBuilder'):
        try:
            self.prev_box.configure(state="normal")
            self.prev_box.delete("1.0", "end")
            self.prev_box.configure(state="disabled")
        except Exception:
            pass

    def _truncate(self: 'RNBuilder', s: str, n: int = 100) -> str:
        s = (s or "").replace("\n", " ").strip()
        return s if len(s) <= n else s[: n - 1] + "…"

    def _rebuild_rn_manager(self: 'RNBuilder'):
        mgr_body = scrollable_body(self.rn_mgr)
        for w in mgr_body.winfo_children():
            w.destroy()
        
        current = self._current_rns()
        for i, rn in enumerate(current):
            row = ctk.CTkFrame(mgr_body)
            row.grid(row=i, column=0, sticky="we", padx=4, pady=2)
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(row, text=self._truncate(rn), anchor="w").grid(row=0, column=0, sticky="w", padx=(4, 6))
            btns = ctk.CTkFrame(row, fg_color="transparent")
            btns.grid(row=0, column=1, sticky="e")
            up = ctk.CTkButton(btns, text="↑", width=28, command=lambda idx=i: self._move_rn(idx, -1))
            down = ctk.CTkButton(btns, text="↓", width=28, command=lambda idx=i: self._move_rn(idx, +1))
            edit = ctk.CTkButton(btns, text="Editar", width=70, command=lambda idx=i: self._edit_rn(idx))
            delete = ctk.CTkButton(btns, text="Excluir", width=70, command=lambda idx=i: self._delete_rn(idx))
            up.pack(side="left", padx=(0, 4))
            down.pack(side="left", padx=(0, 8))
            edit.pack(side="left", padx=(0, 6))
            delete.pack(side="left")
            if i == 0:
                up.configure(state="disabled")
            if i == len(current) - 1:
                down.configure(state="disabled")

    def _refresh_textbox(self: 'RNBuilder'):
        content = "\n\n".join(self._current_rns())
        
        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", content)
        self.txt.configure(state="disabled")
        self._rebuild_rn_manager()

    def _copy_all(self: 'RNBuilder'):
        txt = self.txt.get("1.0", "end").strip()
        if not txt:
            messagebox.showinfo("Nada para copiar", "Adicione ao menos uma RN.")
            return
        self.clipboard_clear()
        self.clipboard_append(txt)
        messagebox.showinfo("Copiado", "Todas as RNs foram copiadas.")

    def _copy_single_rn(self: 'RNBuilder'):
        try:
            sel = self.txt.get("sel.first", "sel.last").strip()
        except Exception:
            sel = ""
        if sel:
            payload = sel
        elif self._current_rns():
            payload = self._current_rns()[-1]
        else:
            payload = self.prev_box.get("1.0", "end").strip()
        if payload:
            self.clipboard_clear()
            self.clipboard_append(payload)
            messagebox.showinfo("Copiado", "RN copiada para a área de transferência.")
        else:
            messagebox.showinfo("Nada para copiar", "Não há RN para copiar.")

    def _save_txt(self: 'RNBuilder'):
        txt = self.txt.get("1.0", "end").strip()
        if not txt:
            messagebox.showinfo("Nada para salvar", "Adicione ao menos uma RN.")
            return
        path = filedialog.asksaveasfilename(
            title="Salvar RNs",
            defaultextension=".txt",
            filetypes=[("Texto", "*.txt"), ("Todos", "*.*")],
            initialfile="rns_geradas.txt",
        )
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(txt)
                messagebox.showinfo("Salvo", f"Arquivo salvo em:\n{path}")
            except Exception as e:
                messagebox.showerror("Erro ao salvar", str(e))

    def _move_rn(self: 'RNBuilder', idx: int, delta: int):
        rns = self._current_rns()
        j = idx + delta
        if 0 <= idx < len(rns) and 0 <= j < len(rns):
            rns[idx], rns[j] = rns[j], rns[idx]
            self._refresh_textbox()

    def _edit_rn(self: 'RNBuilder', idx: int):
        rns = self._current_rns()
        if not (0 <= idx < len(rns)):
            return
        top = ctk.CTkToplevel(self)
        top.title(f"Editar RN #{idx + 1}")
        try:
            _apply_dark_title_bar(top)
        except Exception:
            pass
        top.transient(self)
        top.lift()
        try:
            top.attributes("-topmost", True)
            top.after(300, lambda: top.attributes("-topmost", False))
        except Exception:
            pass
        top.grab_set()
        top.focus_force()
        box = ctk.CTkTextbox(top)
        box.pack(expand=True, fill="both", padx=10, pady=10)
        box.insert("1.0", rns[idx])
        btnbar = ctk.CTkFrame(top, fg_color="transparent")
        btnbar.pack(fill="x", padx=10, pady=(0, 10))

        def _save():
            txt = box.get("1.0", "end").strip()
            if txt:
                rns[idx] = txt
                self._refresh_textbox()
            top.destroy()

        ctk.CTkButton(btnbar, text="Salvar", command=_save, width=120).pack(side="right")
        ctk.CTkButton(btnbar, text="Cancelar", command=top.destroy, width=120).pack(side="right", padx=(0, 8))

        _center_window(top, width=760, height=440, parent=self)

    def _delete_rn(self: 'RNBuilder', idx: int):
        rns = self._current_rns()
        if not (0 <= idx < len(rns)):
            return
        if messagebox.askyesno("Excluir RN", f"Remover a RN #{idx + 1}?"):
            del rns[idx]
            self._refresh_textbox()

    def _update_preview(self: 'RNBuilder'):
        try:
            when = self._when_text()
            cond  = self._cond_text()
            acoes = self._acoes_text(getattr(self, 'acao_rows', []))

            self.prev_box.configure(state="normal")
            self.prev_box.delete("1.0", "end")

            preview = when or ""

            if cond:
                link = " e " if (" se " in when or " Se " in when) else ", se "
                preview += (link if preview else "") + cond

            if acoes:
                preview += (", " if preview else "") + acoes

            if preview:
                preview += "."
                final_txt = preview
            else:
                final_txt = _t("preview_placeholder")

            self.prev_box.insert("end", final_txt)

        except Exception as e:
            try:
                self.prev_box.delete("1.0", "end")
                self.prev_box.insert("end", f"Erro no preview: {e}")
            except Exception:
                pass
        finally:
            try:
                self.prev_box.configure(state="disabled")
            except Exception:
                pass

    def _add_rn(self: 'RNBuilder', *, reset_after: bool = True):
        when = self._when_text()
        cond = self._cond_text()
        acoes = self._acoes_text(getattr(self, 'acao_rows', []))
        if not acoes:
            messagebox.showwarning("Faltam ações", "Adicione pelo menos uma ação.")
            return
        current = self._current_rns()
        idx = int(self.start_idx.get()) + len(current)
        rn = _compose_rn(idx, when, cond, acoes)
        current.append(rn)
        self._refresh_textbox()
        if reset_after:
            self._reset_for_next_rule()

    def _add_rn_and_prepare_opposite(self: 'RNBuilder'):
        when_before = self._when_text()
        self._add_rn(reset_after=False)
        if not when_before:
            return
        self._invert_yes_no_conditions()
        self._destroy_rows(getattr(self, 'acao_rows', []))
        if hasattr(self, '_ensure_min_builder_rows'):
            self._ensure_min_builder_rows()
        self._update_preview()

    def _clear_rns(self: 'RNBuilder', *, confirm=True):
        if (not confirm) or messagebox.askyesno("Limpar", "Remover todas as RNs?"):
            self._current_rns().clear()
            self._refresh_textbox()

    RNBuilder._build_panels = _build_panels
    RNBuilder._set_preview_text = _set_preview_text
    RNBuilder._clear_preview = _clear_preview
    RNBuilder._truncate = _truncate
    RNBuilder._rebuild_rn_manager = _rebuild_rn_manager
    RNBuilder._refresh_textbox = _refresh_textbox
    RNBuilder._copy_all = _copy_all
    RNBuilder._copy_single_rn = _copy_single_rn
    RNBuilder._save_txt = _save_txt
    RNBuilder._move_rn = _move_rn
    RNBuilder._edit_rn = _edit_rn
    RNBuilder._delete_rn = _delete_rn
    RNBuilder._update_preview = _update_preview
    RNBuilder._add_rn = _add_rn
    RNBuilder._add_rn_and_prepare_opposite = _add_rn_and_prepare_opposite
    RNBuilder._clear_rns = _clear_rns

_attach_panels_to_RNBuilder()


if __name__ == "__main__":
    try:
        app = RNBuilder()
        
        # Constrói a interface inicial
        app._build_rule()
        app._build_panels()
        
        app._refresh_gatilho_fields()
        app._ensure_min_builder_rows()
        app._update_preview()
        app._refresh_textbox()
        
        # --- SPLASH SCREEN CLOSE ---
        # Fecha a imagem de carregamento apenas se estiver rodando como .EXE
        try:
            import pyi_splash
            if pyi_splash.is_alive():
                pyi_splash.close()
        except ImportError:
            pass
        # ---------------------------

        app.mainloop()
    except Exception as e:
        try:
            _show_fatal_error(e)
        finally:
            try:
                input("Pressione ENTER para sair...")
            except Exception:
                pass