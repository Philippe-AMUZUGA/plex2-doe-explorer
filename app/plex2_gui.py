#!/usr/bin/env python3
from __future__ import annotations

import ctypes
import json
import queue
import string
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from plex2_core import build_doe_explorer

APP_NAME = "PLEX²"
APP_SUBTITLE = "PLEX² • Explorateur de Plans d'Expériences"
WINDOW_TITLE = APP_NAME
FOOTER_CREDIT = "PLEX² v1.0 update link → https://github.com/Philippe-AMUZUGA • AGPLv3"
BG = "#EEF2F7"
CARD = "#FFFFFF"
ACCENT = "#335C81"
ACCENT_DARK = "#264A68"
TEXT = "#18212F"
MUTED = "#5C6778"
BORDER = "#D6DDE6"


class ScrollableFrame(ttk.Frame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.inner = ttk.Frame(self.canvas)
        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)

    def _on_inner_configure(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.window_id, width=event.width)

    def _bind_mousewheel(self, _event=None):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, _event=None):
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event):
        try:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except tk.TclError:
            pass


class PLEX2App(tk.Tk):
    def __init__(self) -> None:
        self._set_windows_dpi_awareness()
        super().__init__()
        self.title(WINDOW_TITLE)
        self.configure(bg=BG)
        self.resizable(True, True)
        self._set_initial_geometry()
        self.minsize(960, 620)

        self.factor_map: dict[str, list] = {}
        self.selected_factor_name: str | None = None
        self.log_queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self._action_buttons: list[ttk.Button] = []
        self._configure_job = None

        self.output_var = tk.StringVar(value=str(Path.cwd() / "PLEX2_plans.xlsx"))
        self.max_full_var = tk.StringVar(value="30000")
        self.uniform_k_var = tk.StringVar(value="5")
        self.uniform_l_var = tk.StringVar(value="3")
        self.status_var = tk.StringVar(value="Prêt")
        self.credit_var = tk.StringVar(value=FOOTER_CREDIT)

        self._set_icon()
        self._build_style()
        self._build_ui()
        self._load_example()
        self.after(150, self._poll_logs)
        self.bind("<Configure>", self._debounced_layout_guard)

    @staticmethod
    def _set_windows_dpi_awareness() -> None:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

    def _set_initial_geometry(self) -> None:
        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        ratio_w = min(int(screen_w * 0.88), 1366)
        ratio_h = int(ratio_w * 9 / 16)
        if ratio_h > int(screen_h * 0.86):
            ratio_h = min(int(screen_h * 0.86), 820)
            ratio_w = int(ratio_h * 16 / 9)
        x = max(0, (screen_w - ratio_w) // 2)
        y = max(0, (screen_h - ratio_h) // 2)
        self.geometry(f"{ratio_w}x{ratio_h}+{x}+{y}")

    def _guard_geometry(self) -> None:
        try:
            min_w = min(960, max(760, self.winfo_screenwidth() - 80))
            min_h = min(620, max(560, self.winfo_screenheight() - 80))
            w = max(self.winfo_width(), min_w)
            h = max(self.winfo_height(), min_h)
            screen_w = max(self.winfo_screenwidth(), 1024)
            screen_h = max(self.winfo_screenheight(), 700)
            max_w = max(min_w, int(screen_w * 0.98))
            max_h = max(min_h, int(screen_h * 0.94))
            if w > max_w or h > max_h:
                w = min(w, max_w)
                h = min(h, max_h)
                x = max(0, min(self.winfo_x(), screen_w - w))
                y = max(0, min(self.winfo_y(), screen_h - h))
                self.geometry(f"{w}x{h}+{x}+{y}")
            self._relayout_columns()
        except tk.TclError:
            pass

    def _debounced_layout_guard(self, _event=None) -> None:
        if self._configure_job is not None:
            self.after_cancel(self._configure_job)
        self._configure_job = self.after(120, self._guard_geometry)

    def _set_icon(self) -> None:
        assets_dir = Path(__file__).resolve().parent / "assets"
        png_path = assets_dir / "plex2_icon.png"
        ico_path = assets_dir / "plex2_icon.ico"
        try:
            self.iconbitmap(default=str(ico_path))
        except Exception:
            pass
        try:
            self._icon_photo = tk.PhotoImage(file=str(png_path))
            self.iconphoto(True, self._icon_photo)
        except Exception:
            pass

    def _build_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        self.option_add("*Font", "{Segoe UI} 11")
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 11))
        style.configure("Muted.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 11))
        style.configure("Title.TLabel", background=BG, foreground=TEXT, font=("Segoe UI Semibold", 24))
        style.configure("Subtitle.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 12))
        style.configure("Card.TLabelframe", background=CARD, bordercolor=BORDER, relief="solid")
        style.configure("Card.TLabelframe.Label", background=CARD, foreground=TEXT, font=("Segoe UI Semibold", 12))
        style.configure("TButton", font=("Segoe UI Semibold", 11), padding=(12, 8))
        style.configure("Accent.TButton", background=ACCENT, foreground="#FFFFFF", borderwidth=0)
        style.map("Accent.TButton", background=[("pressed", ACCENT_DARK), ("active", ACCENT_DARK)], foreground=[("disabled", "#EAF1F7")])
        style.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT, rowheight=34, font=("Segoe UI", 11), bordercolor=BORDER)
        style.configure("Treeview.Heading", background="#E8EEF5", foreground=TEXT, font=("Segoe UI Semibold", 11), relief="flat")
        style.map("Treeview", background=[("selected", "#D9E7F5")], foreground=[("selected", TEXT)])
        style.configure("TCheckbutton", background=CARD, foreground=TEXT, font=("Segoe UI", 11))
        style.configure("Status.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 11))
        style.configure("Vertical.TScrollbar", arrowsize=18)
        style.configure("Horizontal.TScrollbar", arrowsize=18)

    def _build_ui(self) -> None:
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        outer = ttk.Frame(self, padding=(18, 16, 18, 12))
        outer.grid(row=0, column=0, sticky="nsew")
        outer.grid_rowconfigure(1, weight=1)
        outer.grid_columnconfigure(0, weight=1)

        header = ttk.Frame(outer)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.grid_columnconfigure(0, weight=1)
        title_line = ttk.Frame(header)
        title_line.grid(row=0, column=0, sticky="w")
        # ttk.Label(title_line, text=APP_NAME, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(title_line, text=APP_SUBTITLE, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        # ttk.Label(title_line, text=APP_SUBTITLE, style="Subtitle.TLabel").grid(row=0, column=1, sticky="w", padx=(14, 0), pady=(6, 0))

        main = ttk.Frame(outer)
        main.grid(row=1, column=0, sticky="nsew")
        main.grid_rowconfigure(0, weight=1)
        main.grid_columnconfigure(0, weight=1)

        self.scrollable = ScrollableFrame(main)
        self.scrollable.grid(row=0, column=0, sticky="nsew")

        body = self.scrollable.inner
        body.configure(style="TFrame")
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)
        body.grid_rowconfigure(1, weight=1)

        self.left = ttk.Frame(body)
        self.right = ttk.Frame(body)
        self.left.grid_columnconfigure(0, weight=1)
        self.right.grid_columnconfigure(0, weight=1)

        self._build_left_panel(self.left)
        self._build_right_panel(self.right)
        self.scrollable.canvas.bind("<Configure>", self._on_scroll_canvas_configure, add="+")
        self.after(50, self._relayout_columns)

        footer = ttk.Frame(outer)
        footer.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        footer.grid_columnconfigure(0, weight=1)
        footer.grid_columnconfigure(1, weight=0)
        ttk.Separator(footer).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        ttk.Label(footer, textvariable=self.status_var, style="Status.TLabel").grid(row=1, column=0, sticky="w")
        ttk.Label(footer, textvariable=self.credit_var, style="Status.TLabel").grid(row=1, column=1, sticky="e")

    def _on_scroll_canvas_configure(self, _event=None) -> None:
        self._relayout_columns()

    def _relayout_columns(self) -> None:
        try:
            width = max(self.scrollable.canvas.winfo_width(), self.winfo_width() - 80)
        except Exception:
            width = self.winfo_width()

        if width < 1220:
            self.left.grid_forget()
            self.right.grid_forget()
            self.left.grid(row=0, column=0, sticky="nsew", padx=0, pady=(0, 10))
            self.right.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        else:
            self.left.grid_forget()
            self.right.grid_forget()
            self.left.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=0)
            self.right.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=0)

    def _build_left_panel(self, parent: ttk.Frame) -> None:
        factors_box = ttk.LabelFrame(parent, text="Table des Facteurs", style="Card.TLabelframe", padding=(14, 14, 14, 14))
        factors_box.grid(row=0, column=0, sticky="nsew")
        factors_box.grid_columnconfigure(0, weight=1)
        factors_box.grid_rowconfigure(0, weight=1)

        tree_wrap = ttk.Frame(factors_box)
        tree_wrap.grid(row=0, column=0, sticky="nsew")
        tree_wrap.grid_columnconfigure(0, weight=1)
        tree_wrap.grid_rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_wrap, columns=("name", "levels", "count", "type"), show="headings", selectmode="browse")
        self.tree.heading("name", text="Facteur")
        self.tree.heading("levels", text="Niveaux")
        self.tree.heading("count", text="Nb")
        self.tree.heading("type", text="Type")
        self.tree.column("name", width=160, minwidth=120, stretch=False)
        self.tree.column("levels", width=460, minwidth=220, stretch=True)
        self.tree.column("count", width=65, minwidth=55, anchor="center", stretch=False)
        self.tree.column("type", width=130, minwidth=100, anchor="center", stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", self._on_tree_select)

        tree_scroll = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        tree_scroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=tree_scroll.set)

        editor = ttk.LabelFrame(parent, text="Édition du facteur", style="Card.TLabelframe", padding=(14, 14, 14, 14))
        editor.grid(row=1, column=0, sticky="ew", pady=(10, 10))
        editor.grid_columnconfigure(0, weight=0, minsize=205)
        editor.grid_columnconfigure(1, weight=1)
        editor.grid_rowconfigure(2, weight=1)

        ttk.Label(editor, text="Nom du facteur", style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(editor, text="Niveaux (une valeur par ligne, ou séparés par ; , |)", style="Muted.TLabel").grid(row=0, column=1, sticky="w", padx=(12, 0))

        self.factor_name_var = tk.StringVar()
        name_entry = ttk.Entry(editor, textvariable=self.factor_name_var)
        name_entry.grid(row=1, column=0, sticky="ew", padx=(0, 12), pady=(4, 0))
        name_entry.bind("<Return>", lambda _e: self._add_or_update_factor())

        self.levels_text = ScrolledText(editor, height=6, width=28, wrap="word", font=("Consolas", 11), relief="solid", borderwidth=1)
        self.levels_text.grid(row=1, column=1, rowspan=2, sticky="nsew", pady=(4, 0))
        self.levels_text.configure(highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT, padx=10, pady=10)

        button_row = ttk.Frame(editor)
        button_row.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        save_btn = ttk.Button(button_row, text="Ajouter", style="Accent.TButton", command=self._add_or_update_factor)
        delete_btn = ttk.Button(button_row, text="Supprimer", command=self._delete_selected_factor)
        clear_btn = ttk.Button(button_row, text="Vider", command=self._clear_editor)
        save_btn.grid(row=0, column=0, sticky="w")
        delete_btn.grid(row=0, column=1, sticky="w", padx=(8, 0))
        clear_btn.grid(row=0, column=2, sticky="w", padx=(8, 0))
        self._action_buttons.extend([save_btn, delete_btn, clear_btn])

        quick = ttk.LabelFrame(parent, text="Création rapide", style="Card.TLabelframe", padding=(14, 14, 14, 14))
        quick.grid(row=2, column=0, sticky="ew")
        for idx in range(4):
            quick.grid_columnconfigure(idx, weight=1 if idx == 3 else 0)

        ttk.Label(quick, text="Nombre de facteurs", style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(quick, text="Niveaux par facteur", style="Muted.TLabel").grid(row=0, column=1, sticky="w", padx=(12, 0))
        ttk.Entry(quick, textvariable=self.uniform_k_var, width=10).grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Entry(quick, textvariable=self.uniform_l_var, width=10).grid(row=1, column=1, sticky="w", padx=(12, 0), pady=(4, 0))
        uniform_btn = ttk.Button(quick, text="Générer", command=self._generate_uniform_factors)
        uniform_btn.grid(row=1, column=2, sticky="w", padx=(12, 0), pady=(4, 0))
        self._action_buttons.extend([uniform_btn])

    def _build_right_panel(self, parent: ttk.Frame) -> None:
        settings = ttk.LabelFrame(parent, text="Génération", style="Card.TLabelframe", padding=(14, 14, 14, 14))
        settings.grid(row=0, column=0, sticky="ew")
        settings.grid_columnconfigure(0, weight=1)

        ttk.Label(settings, text="Fichier Excel de sortie", style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        output_row = ttk.Frame(settings)
        output_row.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        output_row.grid_columnconfigure(0, weight=1)
        ttk.Entry(output_row, textvariable=self.output_var).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        browse_btn = ttk.Button(output_row, text="Parcourir", command=self._choose_output)
        browse_btn.grid(row=0, column=1, sticky="e")
        self._action_buttons.append(browse_btn)

        ttk.Label(settings, text="Seuil max factoriel complet", style="Muted.TLabel").grid(row=2, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(settings, textvariable=self.max_full_var, width=14).grid(row=3, column=0, sticky="w", pady=(4, 0))

        actions = ttk.LabelFrame(parent, text="Actions", style="Card.TLabelframe", padding=(14, 14, 14, 14))
        actions.grid(row=1, column=0, sticky="ew", pady=(10, 10))
        actions.grid_columnconfigure(0, weight=1)

        generate_btn = ttk.Button(actions, text="Générer le fichier Excel", style="Accent.TButton", command=self._start_generation)
        export_btn = ttk.Button(actions, text="Exporter les facteurs en JSON", command=self._export_json)
        import_btn = ttk.Button(actions, text="Importer des facteurs JSON", command=self._import_json)
        reset_btn = ttk.Button(actions, text="Réinitialiser", command=self._reset_all)
        generate_btn.grid(row=0, column=0, sticky="ew")
        export_btn.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        import_btn.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        reset_btn.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        self._action_buttons.extend([generate_btn, export_btn, import_btn, reset_btn])

        summary = ttk.LabelFrame(parent, text="Résumé", style="Card.TLabelframe", padding=(14, 14, 14, 14))
        summary.grid(row=2, column=0, sticky="ew")
        summary.grid_columnconfigure(0, weight=1)
        self.summary_text = ScrolledText(summary, wrap="word", height=10, font=("Consolas", 11), relief="solid", borderwidth=1)
        self.summary_text.grid(row=0, column=0, sticky="ew")
        self.summary_text.configure(state="disabled", highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT, padx=10, pady=10)

        logbox = ttk.LabelFrame(parent, text="Journal", style="Card.TLabelframe", padding=(14, 14, 14, 14))
        logbox.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        logbox.grid_columnconfigure(0, weight=1)
        self.log_text = ScrolledText(logbox, wrap="word", height=8, font=("Consolas", 11), relief="solid", borderwidth=1)
        self.log_text.grid(row=0, column=0, sticky="ew")
        self.log_text.configure(state="disabled", highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT, padx=10, pady=10)

    def _parse_levels(self, raw: str) -> list:
        raw = raw.replace(";", "\n").replace(",", "\n").replace("|", "\n")
        items = [line.strip() for line in raw.splitlines() if line.strip()]
        if len(items) < 2:
            raise ValueError("Au moins 2 niveaux sont requis.")
        return [self._coerce_value(item) for item in items]

    @staticmethod
    def _coerce_value(value: str):
        lowered = value.strip().lower()
        if lowered in {"true", "vrai"}:
            return True
        if lowered in {"false", "faux"}:
            return False
        try:
            if value.strip().isdigit() or (value.strip().startswith("-") and value.strip()[1:].isdigit()):
                return int(value)
            return float(value)
        except ValueError:
            return value.strip()

    @staticmethod
    def _factor_type(levels: list) -> str:
        return "Numérique" if all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in levels) else "Texte / mixte"

    def _refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        for name, levels in self.factor_map.items():
            levels_str = ", ".join(str(v) for v in levels)
            self.tree.insert("", "end", iid=name, values=(name, levels_str, len(levels), self._factor_type(levels)))
        self._refresh_summary()

    def _refresh_summary(self) -> None:
        total_factors = len(self.factor_map)
        profile = [len(v) for v in self.factor_map.values()]
        full_runs = 1
        for p in profile:
            full_runs *= p
        lines = [f"Application : {APP_NAME}", f"Facteurs : {total_factors}"]
        if profile:
            lines.append(f"Profil de niveaux : {' x '.join(str(v) for v in profile)}")
            lines.append(f"Factoriel complet potentiel : {full_runs}")
            lines.append(f"Sortie : {Path(self.output_var.get()).name}")
            # lines.append("LHS : calcul automatique")
        else:
            lines.append("Aucun facteur saisi.")
        lines.append("")
        for name, levels in self.factor_map.items():
            lines.append(f"• {name} : {levels}")
        self.summary_text.configure(state="normal")
        self.summary_text.delete("1.0", "end")
        self.summary_text.insert("1.0", "\n".join(lines))
        self.summary_text.configure(state="disabled")

    def _add_or_update_factor(self) -> None:
        name = self.factor_name_var.get().strip()
        if not name:
            messagebox.showerror(APP_NAME, "Le nom du facteur est obligatoire.")
            return
        try:
            levels = self._parse_levels(self.levels_text.get("1.0", "end"))
        except Exception as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return

        if self.selected_factor_name and self.selected_factor_name != name and self.selected_factor_name in self.factor_map:
            self.factor_map.pop(self.selected_factor_name, None)

        self.factor_map[name] = levels
        self.selected_factor_name = name
        self._refresh_tree()
        self.tree.selection_set(name)
        self.tree.focus(name)
        self.status_var.set(f"Facteur « {name} » enregistré")

    def _on_tree_select(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        name = selection[0]
        self.selected_factor_name = name
        levels = self.factor_map[name]
        self.factor_name_var.set(name)
        self.levels_text.delete("1.0", "end")
        self.levels_text.insert("1.0", "\n".join(str(v) for v in levels))

    def _delete_selected_factor(self) -> None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showerror(APP_NAME, "Sélectionnez un facteur à supprimer.")
            return
        name = selection[0]
        self.factor_map.pop(name, None)
        self.selected_factor_name = None
        self._refresh_tree()
        self._clear_editor()
        self.status_var.set(f"Facteur « {name} » supprimé")

    def _clear_editor(self) -> None:
        self.selected_factor_name = None
        self.factor_name_var.set("")
        self.levels_text.delete("1.0", "end")
        self.tree.selection_remove(self.tree.selection())

    def _choose_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Choisir le fichier Excel de sortie",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=Path(self.output_var.get()).name,
        )
        if path:
            self.output_var.set(path)
            self._refresh_summary()

    def _generate_uniform_factors(self) -> None:
        try:
            k = int(self.uniform_k_var.get())
            l = int(self.uniform_l_var.get())
            if k < 1 or l < 2:
                raise ValueError
        except ValueError:
            messagebox.showerror(APP_NAME, "Valeurs invalides pour la création rapide.")
            return
        names = (list(string.ascii_uppercase) + [f"X{i}" for i in range(1, 200)])[:k]
        self.factor_map = {name: list(range(1, l + 1)) for name in names}
        self.selected_factor_name = None
        self._refresh_tree()
        self.status_var.set(f"{k} facteurs à {l} niveaux générés")

    def _load_example(self) -> None:
        self.factor_map = {
            "Matériau": ["Acier", "Alu"],
            "Température": [20.0, 120.0],
            "Pression": [1.0, 2.0, 3.0],
            "Lubrification": ["Oui", "Non"],
            "Epaisseur": [0.4, 0.6],
            "Dureté": [40.0, 60.0, 70.0],
            "Ra": [0.12, 0.8, 1.6],
        }
        self.selected_factor_name = None
        self._refresh_tree()
        self.status_var.set("Configuration par défaut chargée")

    def _export_json(self) -> None:
        if not self.factor_map:
            messagebox.showerror(APP_NAME, "Aucun facteur à exporter.")
            return
        path = filedialog.asksaveasfilename(
            title="Exporter les facteurs en JSON",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialfile="facteurs_plex2.json",
        )
        if not path:
            return
        Path(path).write_text(json.dumps(self.factor_map, ensure_ascii=False, indent=2), encoding="utf-8")
        self.status_var.set(f"Facteurs exportés : {path}")

    def _import_json(self) -> None:
        path = filedialog.askopenfilename(title="Importer des facteurs JSON", filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                raise ValueError("Le JSON doit contenir un objet {facteur: [niveaux]}.")
            clean: dict[str, list] = {}
            for key, values in data.items():
                if not isinstance(values, list) or len(values) < 2:
                    raise ValueError(f"Facteur invalide : {key}")
                clean[str(key)] = values
            self.factor_map = clean
            self.selected_factor_name = None
            self._refresh_tree()
            self.status_var.set(f"Facteurs importés : {path}")
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Import impossible :\n{exc}")

    def _reset_all(self) -> None:
        self.factor_map = {}
        self.selected_factor_name = None
        self.output_var.set(str(Path.cwd() / "PLEX2_plans.xlsx"))
        self.max_full_var.set("30000")
        self.uniform_k_var.set("5")
        self.uniform_l_var.set("3")
        self._refresh_tree()
        self._clear_editor()
        self._clear_log()
        self.status_var.set("Réinitialisé")

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _validate_inputs(self) -> tuple[dict, str, int]:
        if len(self.factor_map) < 1:
            raise ValueError("Ajoutez au moins un facteur.")
        output = self.output_var.get().strip()
        if not output:
            raise ValueError("Choisissez un fichier de sortie .xlsx.")
        if not output.lower().endswith(".xlsx"):
            output += ".xlsx"
            self.output_var.set(output)
        try:
            max_full = int(self.max_full_var.get())
            if max_full < 1:
                raise ValueError
        except ValueError:
            raise ValueError("Le seuil du factoriel complet doit être un entier positif.")
        return self.factor_map.copy(), output, max_full

    def _set_busy(self, busy: bool) -> None:
        cursor = "watch" if busy else ""
        self.configure(cursor=cursor)
        state = tk.DISABLED if busy else tk.NORMAL
        for button in self._action_buttons:
            try:
                button.configure(state=state)
            except tk.TclError:
                pass
        self.update_idletasks()

    def _start_generation(self) -> None:
        try:
            factors, output, max_full = self._validate_inputs()
        except Exception as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return

        self._append_log("Lancement de la génération...")
        self.status_var.set("Génération en cours...")
        self._set_busy(True)

        worker = threading.Thread(target=self._generate_worker, args=(factors, output, max_full), daemon=True)
        worker.start()

    def _generate_worker(self, factors: dict, output: str, max_full: int) -> None:
        try:
            Path(output).parent.mkdir(parents=True, exist_ok=True)
            summary_rows = build_doe_explorer(factors, outfile=output, max_full_factorial_runs=max_full, include_lhs=True)
            self.log_queue.put(("success", json.dumps({"output": output, "rows": summary_rows}, ensure_ascii=False)))
        except Exception as exc:
            self.log_queue.put(("error", str(exc)))

    def _poll_logs(self) -> None:
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()
                self._set_busy(False)
                if kind == "success":
                    data = json.loads(payload)
                    rows = data["rows"] or []
                    kept = [r for r in rows if not str(r.get("Notes", "")).startswith("⚠ OMIS")]
                    skipped = [r for r in rows if str(r.get("Notes", "")).startswith("⚠ OMIS")]
                    self._append_log(f"Fichier généré : {data['output']}")
                    self._append_log(f"Plans disponibles : {len(kept)}")
                    self._append_log(f"Plans omis : {len(skipped)}")
                    self.status_var.set(f"Terminé - {Path(data['output']).name}")
                    messagebox.showinfo(APP_NAME, f"Fichier Excel généré avec succès :\n{data['output']}")
                else:
                    self._append_log(f"Erreur : {payload}")
                    self.status_var.set("Erreur")
                    messagebox.showerror(APP_NAME, payload)
        except queue.Empty:
            pass
        self.after(150, self._poll_logs)


def main() -> None:
    app = PLEX2App()
    app.mainloop()


if __name__ == "__main__":
    main()
