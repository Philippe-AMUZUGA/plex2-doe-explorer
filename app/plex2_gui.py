#!/usr/bin/env python3
from __future__ import annotations

import ctypes
import json
import queue
import string
import sys
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
        self.inner.bind("<Configure>", lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.window_id, width=e.width))
        self.canvas.bind("<Enter>", lambda _e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.canvas.bind("<Leave>", lambda _e: self.canvas.unbind_all("<MouseWheel>"))

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
        self._set_app_icon()
        self.configure(bg=BG)
        self.resizable(True, True)
        self.geometry("1180x760")
        self.minsize(960, 620)

        self.factor_map: dict[str, list] = {}
        self.selected_factor_name: str | None = None
        self.log_queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self._action_buttons: list[ttk.Button] = []

        self.output_var = tk.StringVar(value=str(Path.cwd() / "examples" / "PLEX2_plans.xlsx"))
        self.max_plan_runs_var = tk.StringVar(value="99")
        self.enable_spatial_visualization_var = tk.BooleanVar(value=False)
        self.spatial_visualization_max_runs_var = tk.StringVar(value="99")
        self.uniform_k_var = tk.StringVar(value="5")
        self.uniform_l_var = tk.StringVar(value="3")
        self.status_var = tk.StringVar(value="Prêt")
        self.progress_var = tk.IntVar(value=0)
        self.credit_var = tk.StringVar(value=FOOTER_CREDIT)

        self._build_style()
        self._build_ui()
        self._load_example()
        self.after(150, self._poll_logs)

    @staticmethod
    def _set_windows_dpi_awareness() -> None:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

    def _asset_path(self, filename: str) -> Path:
        app_dir = Path(__file__).resolve().parent
        bundled_root = Path(getattr(sys, "_MEIPASS", app_dir))
        candidates = [
            bundled_root / "assets" / filename,
            app_dir / "assets" / filename,
        ]
        for candidate in candidates:
            if candidate.exists():
                return candidate
        return candidates[-1]

    def _set_app_icon(self) -> None:
        try:
            icon_ico = self._asset_path("plex2_icon.ico")
            if icon_ico.exists():
                self.iconbitmap(str(icon_ico))
                return
        except tk.TclError:
            pass

        try:
            icon_png = self._asset_path("plex2_icon.png")
            if icon_png.exists():
                self._app_icon_image = tk.PhotoImage(file=str(icon_png))
                self.iconphoto(True, self._app_icon_image)
        except tk.TclError:
            pass

    def _build_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        self.option_add("*Font", "{Segoe UI} 11")
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=TEXT)
        style.configure("Muted.TLabel", background=BG, foreground=MUTED)
        style.configure("Title.TLabel", background=BG, foreground=TEXT, font=("Segoe UI Semibold", 22))
        style.configure("Card.TLabelframe", background=CARD, bordercolor=BORDER, relief="solid")
        style.configure("Card.TLabelframe.Label", background=CARD, foreground=TEXT, font=("Segoe UI Semibold", 12))
        style.configure("TButton", font=("Segoe UI Semibold", 11), padding=(10, 7))
        style.configure("Accent.TButton", background=ACCENT, foreground="#FFFFFF")
        style.map("Accent.TButton", background=[("pressed", ACCENT_DARK), ("active", ACCENT_DARK)])
        style.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT, rowheight=32)
        style.configure("Treeview.Heading", background="#E8EEF5", foreground=TEXT, font=("Segoe UI Semibold", 10))

    def _build_ui(self) -> None:
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        outer = ttk.Frame(self, padding=(18, 16, 18, 12))
        outer.grid(row=0, column=0, sticky="nsew")
        outer.grid_rowconfigure(1, weight=1)
        outer.grid_columnconfigure(0, weight=1)

        ttk.Label(outer, text=APP_SUBTITLE, style="Title.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 12))

        scroll_area = ScrollableFrame(outer)
        scroll_area.grid(row=1, column=0, sticky="nsew")
        main = ttk.Frame(scroll_area.inner)
        main.grid(row=0, column=0, sticky="nsew")
        scroll_area.inner.grid_rowconfigure(0, weight=1)
        scroll_area.inner.grid_columnconfigure(0, weight=1)
        main.grid_columnconfigure(0, weight=2)
        main.grid_columnconfigure(1, weight=1)
        main.grid_rowconfigure(0, weight=1)

        left = ttk.Frame(main)
        right = ttk.Frame(main)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        left.grid_columnconfigure(0, weight=1)
        right.grid_columnconfigure(0, weight=1)

        self._build_left_panel(left)
        self._build_right_panel(right)

        footer = ttk.Frame(outer)
        footer.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        footer.grid_columnconfigure(0, weight=1)
        footer.grid_columnconfigure(1, weight=0)
        ttk.Separator(footer).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Label(footer, textvariable=self.status_var).grid(row=1, column=0, sticky="w")
        ttk.Label(footer, textvariable=self.credit_var).grid(row=1, column=1, sticky="e")

    def _build_left_panel(self, parent: ttk.Frame) -> None:
        factors_box = ttk.LabelFrame(parent, text="Table des Facteurs", style="Card.TLabelframe", padding=12)
        factors_box.grid(row=0, column=0, sticky="ew")
        factors_box.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=0)

        self.tree = ttk.Treeview(factors_box, columns=("name", "levels", "count", "type"), show="headings", selectmode="browse", height=8)
        for col, label, width in [("name", "Facteur", 150), ("levels", "Niveaux", 360), ("count", "Nb", 55), ("type", "Type", 110)]:
            self.tree.heading(col, text=label)
            self.tree.column(col, width=width, anchor="center" if col in {"count", "type"} else "w")
        self.tree.grid(row=0, column=0, sticky="ew")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        ttk.Scrollbar(factors_box, orient="vertical", command=self.tree.yview).grid(row=0, column=1, sticky="ns")

        editor = ttk.LabelFrame(parent, text="Édition du facteur", style="Card.TLabelframe", padding=12)
        editor.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        editor.grid_columnconfigure(1, weight=1)
        ttk.Label(editor, text="Nom", style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(editor, text="Niveaux", style="Muted.TLabel").grid(row=0, column=1, sticky="w", padx=(10, 0))
        self.factor_name_var = tk.StringVar()
        ttk.Entry(editor, textvariable=self.factor_name_var, width=20).grid(row=1, column=0, sticky="ew", padx=(0, 10))
        self.levels_text = ScrolledText(editor, height=5, font=("Consolas", 10), wrap="word")
        self.levels_text.grid(row=1, column=1, sticky="ew")
        row = ttk.Frame(editor)
        row.grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 0))
        for text, cmd, style in [("Ajouter / Modifier", self._add_or_update_factor, "Accent.TButton"), ("Supprimer", self._delete_selected_factor, "TButton"), ("Vider", self._clear_editor, "TButton")]:
            btn = ttk.Button(row, text=text, command=cmd, style=style)
            btn.pack(side="left", padx=(0, 8))
            self._action_buttons.append(btn)

        quick = ttk.LabelFrame(parent, text="Création rapide", style="Card.TLabelframe", padding=12)
        quick.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(quick, text="Facteurs", style="Muted.TLabel").grid(row=0, column=0)
        ttk.Label(quick, text="Niveaux", style="Muted.TLabel").grid(row=0, column=1, padx=(10, 0))
        ttk.Entry(quick, textvariable=self.uniform_k_var, width=8).grid(row=1, column=0)
        ttk.Entry(quick, textvariable=self.uniform_l_var, width=8).grid(row=1, column=1, padx=(10, 0))
        btn = ttk.Button(quick, text="Générer", command=self._generate_uniform_factors)
        btn.grid(row=1, column=2, padx=(10, 0))
        self._action_buttons.append(btn)

        summary = ttk.LabelFrame(parent, text="Résumé", style="Card.TLabelframe", padding=12)
        summary.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        self.summary_text = ScrolledText(summary, wrap="word", height=8, font=("Consolas", 10))
        self.summary_text.pack(fill="both", expand=True)
        self.summary_text.configure(state="disabled")

    def _build_right_panel(self, parent: ttk.Frame) -> None:
        settings = ttk.LabelFrame(parent, text="Génération", style="Card.TLabelframe", padding=12)
        settings.grid(row=0, column=0, sticky="ew")
        settings.grid_columnconfigure(0, weight=1)
        ttk.Label(settings, text="Fichier Excel de sortie", style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        row = ttk.Frame(settings)
        row.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        row.grid_columnconfigure(0, weight=1)
        ttk.Entry(row, textvariable=self.output_var).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        btn = ttk.Button(row, text="Parcourir", command=self._choose_output)
        btn.grid(row=0, column=1)
        self._action_buttons.append(btn)
        ttk.Label(settings, text="Limiter le nombre d'expériences", style="Muted.TLabel").grid(row=2, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(settings, textvariable=self.max_plan_runs_var, width=14).grid(row=3, column=0, sticky="w", pady=(4, 0))

        visu = ttk.LabelFrame(parent, text="Visualisation", style="Card.TLabelframe", padding=12)
        visu.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        visu.grid_columnconfigure(0, weight=1)
        ttk.Checkbutton(
            visu,
            text="Activer la visualisation dans l'Excel",
            variable=self.enable_spatial_visualization_var,
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            visu,
            text="Nombre maximal d'essais autorisé pour la visualisation",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(visu, textvariable=self.spatial_visualization_max_runs_var, width=14).grid(row=2, column=0, sticky="w", pady=(4, 0))

        actions = ttk.LabelFrame(parent, text="Actions", style="Card.TLabelframe", padding=12)
        actions.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        actions.grid_columnconfigure(0, weight=1)
        for text, cmd, style in [("Générer le fichier Excel", self._start_generation, "Accent.TButton"), ("Exporter les facteurs en JSON", self._export_json, "TButton"), ("Importer des facteurs JSON", self._import_json, "TButton"), ("Réinitialiser", self._reset_all, "TButton")]:
            btn = ttk.Button(actions, text=text, command=cmd, style=style)
            btn.pack(fill="x", pady=(0, 8))
            self._action_buttons.append(btn)

        progress_box = ttk.LabelFrame(parent, text="Progression", style="Card.TLabelframe", padding=12)
        progress_box.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        progress_box.grid_columnconfigure(0, weight=1)
        self.progress_bar = ttk.Progressbar(progress_box, variable=self.progress_var, maximum=100, mode="determinate")
        self.progress_bar.grid(row=0, column=0, sticky="ew")

        logbox = ttk.LabelFrame(parent, text="Journal", style="Card.TLabelframe", padding=12)
        logbox.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        self.log_text = ScrolledText(logbox, wrap="word", height=8, font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True)
        self.log_text.configure(state="disabled")

    def _parse_levels(self, raw: str) -> list:
        raw = raw.replace(";", "\n").replace(",", "\n").replace("\t", "\n")
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
            self.tree.insert("", "end", iid=name, values=(name, ", ".join(str(v) for v in levels), len(levels), self._factor_type(levels)))
        self._refresh_summary()

    def _refresh_summary(self) -> None:
        profile = [len(v) for v in self.factor_map.values()]
        full_runs = 1
        for p in profile:
            full_runs *= p
        lines = [f"Application : {APP_NAME}", f"Facteurs : {len(self.factor_map)}"]
        if profile:
            lines += [f"Profil de niveaux : {' x '.join(str(v) for v in profile)}", f"Factoriel complet potentiel : {full_runs}", f"Sortie : {Path(self.output_var.get()).name}", ""]
            for name, levels in self.factor_map.items():
                lines.append(f"• {name} : {levels}")
        else:
            lines.append("Aucun facteur saisi.")
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
        if self.selected_factor_name and self.selected_factor_name != name:
            self.factor_map.pop(self.selected_factor_name, None)
        self.factor_map[name] = levels
        self.selected_factor_name = name
        self._refresh_tree()
        self.tree.selection_set(name)
        self.status_var.set(f"Facteur « {name} » enregistré")

    def _on_tree_select(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        name = selection[0]
        self.selected_factor_name = name
        self.factor_name_var.set(name)
        self.levels_text.delete("1.0", "end")
        self.levels_text.insert("1.0", "\n".join(str(v) for v in self.factor_map[name]))

    def _delete_selected_factor(self) -> None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showerror(APP_NAME, "Sélectionnez un facteur à supprimer.")
            return
        name = selection[0]
        self.factor_map.pop(name, None)
        self.selected_factor_name = None
        self._clear_editor()
        self._refresh_tree()

    def _clear_editor(self) -> None:
        self.selected_factor_name = None
        self.factor_name_var.set("")
        self.levels_text.delete("1.0", "end")

    def _choose_output(self) -> None:
        output_path = Path(self.output_var.get())
        path = filedialog.asksaveasfilename(
            title="Choisir le fichier Excel de sortie",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialdir=str(output_path.parent),
            initialfile=output_path.name,
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
        self._refresh_tree()

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
        self._refresh_tree()

    def _export_json(self) -> None:
        if not self.factor_map:
            messagebox.showerror(APP_NAME, "Aucun facteur à exporter.")
            return
        path = filedialog.asksaveasfilename(title="Exporter les facteurs en JSON", defaultextension=".json", filetypes=[("JSON", "*.json")], initialfile="facteurs_plex2.json")
        if path:
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
            self.factor_map = {str(k): v for k, v in data.items() if isinstance(v, list) and len(v) >= 2}
            self._refresh_tree()
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Import impossible :\n{exc}")

    def _reset_all(self) -> None:
        self.factor_map = {}
        self.progress_var.set(0)
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

    def _validate_inputs(self) -> tuple[dict, str, int, bool, int]:
        if not self.factor_map:
            raise ValueError("Ajoutez au moins un facteur.")
        output = self.output_var.get().strip()
        if not output:
            raise ValueError("Choisissez un fichier de sortie .xlsx.")
        if not output.lower().endswith(".xlsx"):
            output += ".xlsx"
        self.output_var.set(output)
        try:
            max_plan_runs = int(self.max_plan_runs_var.get())
            if max_plan_runs < 1:
                raise ValueError
        except ValueError:
            raise ValueError("Le seuil max de configurations par plan doit être un entier positif.")
        try:
            spatial_visualization_max_runs = int(self.spatial_visualization_max_runs_var.get())
            if spatial_visualization_max_runs < 1:
                raise ValueError
        except ValueError:
            raise ValueError("Le nombre maximal d'essais pour la visualisation doit être un entier positif.")
        return (
            self.factor_map.copy(),
            output,
            max_plan_runs,
            bool(self.enable_spatial_visualization_var.get()),
            spatial_visualization_max_runs,
        )

    def _set_busy(self, busy: bool) -> None:
        self.configure(cursor="watch" if busy else "")
        state = tk.DISABLED if busy else tk.NORMAL
        for button in self._action_buttons:
            try:
                button.configure(state=state)
            except tk.TclError:
                pass
        self.update_idletasks()

    def _start_generation(self) -> None:
        try:
            (
                factors,
                output,
                max_plan_runs,
                enable_spatial_visualization,
                spatial_visualization_max_runs,
            ) = self._validate_inputs()
        except Exception as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        self.progress_var.set(0)
        self.status_var.set("Génération en cours... 0%")
        self._append_log("Lancement de la génération...")
        self._set_busy(True)
        worker = threading.Thread(
            target=self._generate_worker,
            args=(
                factors,
                output,
                max_plan_runs,
                enable_spatial_visualization,
                spatial_visualization_max_runs,
            ),
            daemon=True,
        )
        worker.start()

    def _progress_callback(self, percent: int, message: str = "Génération en cours") -> None:
        self.log_queue.put(("progress", json.dumps({"percent": int(percent), "message": message}, ensure_ascii=False)))

    def _generate_worker(
        self,
        factors: dict,
        output: str,
        max_plan_runs: int,
        enable_spatial_visualization: bool,
        spatial_visualization_max_runs: int,
    ) -> None:
        try:
            output_path = Path(output)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            summary_rows = build_doe_explorer(
                factors,
                outfile=output,
                max_plan_runs=max_plan_runs,
                include_lhs=True,
                enable_spatial_visualization=enable_spatial_visualization,
                spatial_visualization_max_runs=spatial_visualization_max_runs,
                progress_callback=self._progress_callback,
            )
            json_output = str(output_path.with_suffix(".json"))
            Path(json_output).write_text(json.dumps(factors, ensure_ascii=False, indent=2), encoding="utf-8")
            self.log_queue.put(("success", json.dumps({"output": output, "json_output": json_output, "rows": summary_rows}, ensure_ascii=False)))
        except Exception as exc:
            self.log_queue.put(("error", str(exc)))

    def _poll_logs(self) -> None:
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()
                if kind == "progress":
                    data = json.loads(payload)
                    pct = max(0, min(100, int(data.get("percent", 0))))
                    msg = data.get("message", "Génération en cours")
                    self.progress_var.set(pct)
                    self.status_var.set(f"{msg}... {pct}%")
                elif kind == "success":
                    self._set_busy(False)
                    self.progress_var.set(100)
                    data = json.loads(payload)
                    rows = data.get("rows") or []
                    kept = [r for r in rows if not str(r.get("Notes", "")).startswith("⚠ OMIS")]
                    skipped = [r for r in rows if str(r.get("Notes", "")).startswith("⚠ OMIS")]
                    self._append_log(f"Fichier généré : {data['output']}")
                    self._append_log(f"Facteurs JSON exportés : {data['json_output']}")
                    self._append_log(f"Plans disponibles : {len(kept)}")
                    self._append_log(f"Plans omis : {len(skipped)}")
                    self.status_var.set(f"Terminé - {Path(data['output']).name} - 100%")
                    messagebox.showinfo(APP_NAME, f"Fichier Excel généré avec succès :\n{data['output']}\n\nFacteurs JSON exportés :\n{data['json_output']}")
                else:
                    self._set_busy(False)
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
