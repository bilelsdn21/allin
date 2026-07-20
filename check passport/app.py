"""
Passport Check - desktop app (bilingual English / French)
=========================================================
A window with buttons for the whole job. No console needed.

  1. Turn PDFs into photos   -> PDF passports become images you can rename
  2. Add the ID numbers      -> matches photos to the Excel list and writes
                                the number on each photo

Everything free, built on Python's Tkinter (no extra install).
Run by double-clicking "Passport Check.bat" or the desktop shortcut.
"""

import importlib.util
import json
import os
import queue
import shutil
import sys
import threading
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import passport_check as pc
import split_pdfs as sp

# drag-and-drop is optional; fall back to click-to-browse if not installed
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    BaseTk = TkinterDnD.Tk
    HAS_DND = True
except Exception:
    BaseTk = tk.Tk
    HAS_DND = False

HERE = Path(__file__).parent
IMG_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp", ".pdf", ""}
XLS_EXTS = {".xls", ".xlsx", ".xlsm"}

# ---- colours -------------------------------------------------------------- #
BG = "#f4f6f8"
CARD = "#ffffff"
PRIMARY = "#2563eb"
GREEN = "#16a34a"
RED = "#dc2626"
AMBER = "#d97706"
GREY = "#6b7280"
DARK = "#1f2937"
PURPLE = "#6b21a8"

# ---- all texts: (English, French) ----------------------------------------- #
T = {
    "hint_dnd": ("Step 1: put your files on the two boxes below.   "
                 "Step 2: click the green button.",
                 "Étape 1 : déposez vos fichiers sur les deux cases.   "
                 "Étape 2 : cliquez sur le bouton vert."),
    "hint_click": ("Step 1: click each box and choose your files.   "
                   "Step 2: click the green button.",
                   "Étape 1 : cliquez sur chaque case et choisissez vos "
                   "fichiers.   Étape 2 : cliquez sur le bouton vert."),
    "zone_img": ("\U0001F4F7  Passport photos", "Photos des passeports"),
    "zone_img_sub": ("Put the photos (and PDF files) here",
                     "Déposez ici les photos (et les PDF)"),
    "zone_xls": ("\U0001F4CB  Name list (Excel)", "Liste des noms (Excel)"),
    "zone_xls_sub": ("Put the Excel name list here",
                     "Déposez ici la liste Excel"),
    "btn_split": ("First: turn PDFs into photos", "D'abord : PDF → photos"),
    "btn_run": ("✔  Add the ID numbers", "Ajouter les numéros"),
    "btn_output": ("Finished photos", "Photos finies"),
    "btn_report": ("The report", "Le rapport"),
    "btn_review": ("PDF photos to sort", "PDF à trier"),
    "btn_clear": ("Delete all photos", "Supprimer les photos"),
    "card_done": ("Number added", "Numéro ajouté"),
    "card_index": ("List number only", "Numéro de liste seul"),
    "card_no_photo": ("No photo yet", "Pas encore de photo"),
    "card_no_match": ("Photo has no match", "Photo sans nom trouvé"),
    "card_broken": ("Problem files", "Fichiers à problème"),
    "card_pdf": ("PDF files to sort", "PDF à trier"),
    "tab_summary": ("Summary", "Résumé"),
    "tab_photos": ("All photos", "Toutes les photos"),
    "tab_messages": ("Messages", "Messages"),
    "tab_settings": ("Settings", "Réglages"),
    "headline_empty": ("Click the green button to see the results.",
                       "Cliquez sur le bouton vert pour voir les résultats."),
    "banner_manual": ("  ⚠  Please check these by hand - double-click a "
                      "line to fix the name automatically",
                      "À vérifier à la main - double-cliquez une ligne pour "
                      "corriger le nom automatiquement"),
    "col_item": ("Photo or name", "Photo ou nom"),
    "col_todo": ("What to do", "Quoi faire"),
    "col_status": ("What happened", "Résultat"),
    "col_who": ("Name / photo", "Nom / photo"),
    "col_detail": ("Details", "Détails"),
    "photos_ready": ("photos ready", "photos prêtes"),
    "no_photos": ("no photos yet", "pas encore de photos"),
    "no_list": ("no list yet", "pas encore de liste"),
    "st_done": ("Done", "OK"),
    "st_index": ("List number only", "Numéro de liste"),
    "st_no_match": ("Photo has no match", "Sans nom trouvé"),
    "st_no_photo": ("No photo yet", "Pas de photo"),
    "st_broken": ("Problem file", "Fichier à problème"),
    "st_pdf": ("PDF file", "Fichier PDF"),
    "st_manual": ("Check by hand", "À vérifier"),
    "st_no_number": ("No number in list", "Pas de numéro"),
    "hint_photos_tab": ("Double-click a line to open the photo.",
                        "Double-cliquez une ligne pour ouvrir la photo."),
}


def tr(key, sep="  /  "):
    en, fr = T[key]
    return en + sep + fr


def tr2(key):
    """Two-line version (English on top, French under)."""
    en, fr = T[key]
    return en + "\n" + fr


class App(BaseTk):
    def __init__(self):
        super().__init__()
        self.title(f"Passport Check  v{pc.__version__}")
        self.configure(bg=BG)
        self.minsize(900, 700)

        self.msg_q = queue.Queue()
        self.running = False
        self.excel_path = self._find_existing_excel()
        self.manual_fix = {}   # tree row id -> (Path to rename, new name)
        self.row_file = {}     # all-photos row id -> Path to open

        self._restore_window()
        self._build_styles()
        self._build_header()
        self._build_dropzones()
        self._build_actions()
        self._build_summary_cards()
        self._build_tabs()
        self._load_settings_into_form()
        self._refresh_status()
        self._startup_selfcheck()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(80, self._drain_queue)

    # ------------------------------------------------------------- window mem
    def _restore_window(self):
        try:
            geo = pc.load_config().get("window", "")
            self.geometry(geo if geo else "1000x780")
        except Exception:
            self.geometry("1000x780")

    def _on_close(self):
        try:
            cfg = self._cfg()
            cfg["window"] = self.geometry()
            (HERE / "config.json").write_text(json.dumps(cfg, indent=2),
                                              encoding="utf-8")
        except Exception:
            pass
        self.destroy()

    def _find_existing_excel(self):
        for p in HERE.iterdir():
            if (p.suffix.lower() in XLS_EXTS and not p.name.startswith("~$")
                    and not p.name.lower().startswith("report")):
                return p
        return None

    # ------------------------------------------------------------------ style
    def _build_styles(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("TFrame", background=BG)
        st.configure("Card.TFrame", background=CARD, relief="flat")
        st.configure("TLabel", background=BG, foreground=DARK, font=("Segoe UI", 10))
        st.configure("Card.TLabel", background=CARD, foreground=DARK)
        st.configure("Title.TLabel", background=BG, foreground=DARK,
                     font=("Segoe UI Semibold", 20))
        st.configure("Sub.TLabel", background=BG, foreground=GREY,
                     font=("Segoe UI", 10))
        st.configure("Big.TButton", font=("Segoe UI Semibold", 10), padding=(12, 10))
        st.configure("Go.TButton", font=("Segoe UI Semibold", 11), padding=(14, 10),
                     foreground="white", background=GREEN)
        st.map("Go.TButton", background=[("active", "#15803d"),
                                         ("disabled", "#9ca3af")])
        st.configure("TButton", padding=(8, 6))
        st.configure("TNotebook", background=BG, borderwidth=0)
        st.configure("TNotebook.Tab", font=("Segoe UI", 10), padding=(12, 6))

    # ----------------------------------------------------------------- header
    def _build_header(self):
        top = ttk.Frame(self, style="TFrame")
        top.pack(fill="x", padx=20, pady=(14, 2))
        ttk.Label(top, text="Passport Check", style="Title.TLabel").pack(anchor="w")
        hint = tr("hint_dnd") if HAS_DND else tr("hint_click")
        ttk.Label(top, text=hint, style="Sub.TLabel",
                  wraplength=940, justify="left").pack(anchor="w")
        # amber banner filled by the startup self-check when something is off
        self.banner = tk.Label(top, text="", bg="#fef3c7", fg="#92400e",
                               anchor="w", font=("Segoe UI", 9), justify="left")

    def _startup_selfcheck(self):
        """Warn early (in both languages) if an optional piece is missing."""
        problems = []
        if pc.fitz is None:
            problems.append("PDF support is off (pymupdf is not installed). "
                            "/ Le support PDF est désactivé.")
        if not HAS_DND:
            problems.append("Drag-and-drop is off - click the boxes instead. "
                            "/ Le glisser-déposer est désactivé - cliquez "
                            "sur les cases.")
        if importlib.util.find_spec("xlrd") is None:
            problems.append("Old .xls lists will not open (xlrd missing). "
                            "/ Les anciens fichiers .xls ne s'ouvriront pas.")
        try:
            test = HERE / ".write_test"
            test.write_text("x")
            test.unlink()
        except OSError:
            problems.append("Cannot write in this folder - results will fail. "
                            "/ Impossible d'écrire dans ce dossier.")
        if problems:
            self.banner.config(text="  " + "\n  ".join(problems))
            self.banner.pack(fill="x", pady=(6, 0), ipady=4)

    # -------------------------------------------------------------- drop zones
    def _build_dropzones(self):
        wrap = ttk.Frame(self, style="TFrame")
        wrap.pack(fill="x", padx=20, pady=(8, 4))

        self.zone_img = self._make_zone(
            wrap, tr2("zone_img"), tr2("zone_img_sub"),
            self.pick_images, side="left")
        self.zone_xls = self._make_zone(
            wrap, tr2("zone_xls"), tr2("zone_xls_sub"),
            self.pick_excel, side="right")

    def _make_zone(self, parent, title, subtitle, on_click, side):
        outer = tk.Frame(parent, bg=CARD, highlightbackground="#c7d2fe",
                         highlightthickness=2, cursor="hand2")
        outer.pack(side=side, expand=True, fill="both",
                   padx=(0, 6) if side == "left" else (6, 0))
        tk.Label(outer, text=title, bg=CARD, fg=DARK, justify="center",
                 font=("Segoe UI Semibold", 11)).pack(pady=(12, 2))
        sub = tk.Label(outer, text=subtitle, bg=CARD, fg=GREY,
                       justify="center", font=("Segoe UI", 9))
        sub.pack()
        status = tk.Label(outer, text="", bg=CARD, fg=PRIMARY,
                          font=("Segoe UI Semibold", 10))
        status.pack(pady=(4, 12))

        for w in (outer, sub, status):
            w.bind("<Button-1>", lambda e: on_click())

        if HAS_DND:
            for w in (outer, sub, status):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>",
                           lambda e, fn=on_click: self._on_drop(e, fn))
        return {"outer": outer, "status": status}

    def _on_drop(self, event, kind_fn):
        paths = [Path(p) for p in self.tk.splitlist(event.data)]
        if kind_fn == self.pick_images:
            self._add_images(paths)
        else:
            self._set_excel(paths)

    # ---------------------------------------------------------------- actions
    def _build_actions(self):
        bar = ttk.Frame(self, style="TFrame")
        bar.pack(fill="x", padx=20, pady=6)

        self.btn_split = ttk.Button(bar, text=tr("btn_split", " / "),
                                    style="Big.TButton", command=self.on_split)
        self.btn_run = ttk.Button(bar, text=tr("btn_run", " / "),
                                  style="Go.TButton", command=self.on_run)
        self.btn_split.pack(side="left")
        self.btn_run.pack(side="left", padx=(10, 0))

        ttk.Button(bar, text=tr("btn_output", " / "),
                   command=lambda: self.open_path("output")
                   ).pack(side="left", padx=(18, 0))
        ttk.Button(bar, text=tr("btn_report", " / "),
                   command=self.open_report).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text=tr("btn_review", " / "),
                   command=lambda: self.open_path("to_review")
                   ).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text=tr("btn_clear", " / "),
                   command=self.clear_images).pack(side="left", padx=(6, 0))

        # progress: a real bar (0-100) + "photo 45 / 119"
        prog_wrap = ttk.Frame(bar, style="TFrame")
        prog_wrap.pack(side="right")
        self.prog_label = ttk.Label(prog_wrap, text="", style="Sub.TLabel")
        self.prog_label.pack(anchor="e")
        self.progress = ttk.Progressbar(prog_wrap, mode="determinate",
                                        length=170, maximum=100)
        self.progress.pack(anchor="e")

    # ---------------------------------------------------------------- summary
    def _build_summary_cards(self):
        wrap = ttk.Frame(self, style="TFrame")
        wrap.pack(fill="x", padx=20, pady=(4, 8))
        self.cards = {}
        specs = [("done", "card_done", GREEN),
                 ("index_only", "card_index", PRIMARY),
                 ("no_photo", "card_no_photo", RED),
                 ("no_match", "card_no_match", AMBER),
                 ("broken", "card_broken", RED),
                 ("pdf", "card_pdf", DARK)]
        for key, tkey, color in specs:
            card = tk.Frame(wrap, bg=CARD, highlightbackground="#e5e7eb",
                            highlightthickness=1)
            card.pack(side="left", expand=True, fill="both", padx=3)
            num = tk.Label(card, text="-", bg=CARD, fg=color,
                           font=("Segoe UI Semibold", 20))
            num.pack(pady=(8, 0))
            tk.Label(card, text=tr2(tkey), bg=CARD, fg=GREY, justify="center",
                     font=("Segoe UI", 8), wraplength=130).pack(pady=(0, 8))
            self.cards[key] = num

    # ------------------------------------------------------------------- tabs
    def _build_tabs(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=20, pady=(0, 14))
        self.nb = nb

        self._build_summary_tab(nb)

        # all-photos table
        res = ttk.Frame(nb, style="TFrame")
        nb.add(res, text=tr("tab_photos", " / "))
        self.photos_hint = ttk.Label(res, text=tr("hint_photos_tab"),
                                     style="Sub.TLabel")
        self.photos_hint.pack(anchor="w", padx=4, pady=(4, 2))
        cols = ("status", "who", "detail")
        self.tree = ttk.Treeview(res, columns=cols, show="headings", height=12)
        for c, w, tkey in (("status", 180, "col_status"),
                           ("who", 280, "col_who"),
                           ("detail", 430, "col_detail")):
            self.tree.heading(c, text=tr(tkey, " / "))
            self.tree.column(c, width=w, anchor="w")
        vs = ttk.Scrollbar(res, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        self.tree.tag_configure("ok", foreground=GREEN)
        self.tree.tag_configure("blue", foreground=PRIMARY)
        self.tree.tag_configure("red", foreground=RED)
        self.tree.tag_configure("amber", foreground=AMBER)
        self.tree.tag_configure("grey", foreground=GREY)
        self.tree.bind("<Double-1>", self._open_row_photo)

        # log
        logf = ttk.Frame(nb, style="TFrame")
        nb.add(logf, text=tr("tab_messages", " / "))
        self.log = tk.Text(logf, wrap="word", bg="#0b1020", fg="#d1e0ff",
                           font=("Consolas", 9), relief="flat")
        self.log.pack(fill="both", expand=True)

        self._build_settings_tab(nb)

    def _build_summary_tab(self, nb):
        s = tk.Frame(nb, bg=BG)
        nb.add(s, text=tr("tab_summary", " / "))

        self.headline = tk.Label(s, text=tr2("headline_empty"), bg=BG, fg=DARK,
                                 justify="left", font=("Segoe UI Semibold", 15))
        self.headline.pack(anchor="w", padx=16, pady=(14, 4))

        self.stat_lines = tk.Label(s, text="", bg=BG, fg=GREY, justify="left",
                                   font=("Segoe UI", 10))
        self.stat_lines.pack(anchor="w", padx=16, pady=(0, 8))

        banner = tk.Label(s, text=tr2("banner_manual"),
                          bg="#ede9fe", fg=PURPLE, anchor="w", justify="left",
                          font=("Segoe UI Semibold", 10))
        banner.pack(fill="x", padx=16, pady=(4, 0), ipady=5)

        wrap = tk.Frame(s, bg=BG)
        wrap.pack(fill="both", expand=True, padx=16, pady=(0, 14))
        self.manual_tree = ttk.Treeview(wrap, columns=("item", "why"),
                                        show="headings", height=8)
        self.manual_tree.heading("item", text=tr("col_item", " / "))
        self.manual_tree.heading("why", text=tr("col_todo", " / "))
        self.manual_tree.column("item", width=250)
        self.manual_tree.column("why", width=640)
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.manual_tree.yview)
        self.manual_tree.configure(yscrollcommand=vs.set)
        self.manual_tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        self.manual_tree.tag_configure("purple", background="#f5f3ff",
                                       foreground=PURPLE)
        self.manual_tree.bind("<Double-1>", self._fix_selected_name)

    def _build_settings_tab(self, nb):
        s = ttk.Frame(nb, style="TFrame")
        nb.add(s, text=tr("tab_settings", " / "))
        self.vars = {}

        def row(parent, label, r):
            ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w",
                                               padx=8, pady=6)

        grid = ttk.Frame(s, style="TFrame")
        grid.pack(anchor="w", padx=10, pady=10)

        row(grid, "Word before the ID number / Mot avant le numéro", 0)
        self.vars["prefix"] = tk.StringVar()
        ttk.Entry(grid, textvariable=self.vars["prefix"], width=18).grid(row=0, column=1)

        row(grid, "Also write the list number / Écrire aussi le n° de liste", 1)
        self.vars["show_index"] = tk.BooleanVar()
        ttk.Checkbutton(grid, variable=self.vars["show_index"]).grid(row=1, column=1, sticky="w")

        row(grid, "Word before the list number / Mot avant le n° de liste", 2)
        self.vars["index_prefix"] = tk.StringVar()
        ttk.Entry(grid, textvariable=self.vars["index_prefix"], width=18).grid(row=2, column=1)

        row(grid, "Where to write on the photo / Position sur la photo", 3)
        self.vars["position"] = tk.StringVar()
        ttk.Combobox(grid, textvariable=self.vars["position"], width=16, state="readonly",
                     values=["bottom-right", "bottom-left", "top-right", "top-left"]
                     ).grid(row=3, column=1, sticky="w")

        row(grid, "Text size / Taille du texte", 4)
        self.vars["font_size"] = tk.IntVar()
        ttk.Spinbox(grid, from_=12, to=200, textvariable=self.vars["font_size"],
                    width=8).grid(row=4, column=1, sticky="w")

        row(grid, "Number in front of file names / N° devant les fichiers", 5)
        self.vars["number_output_files"] = tk.BooleanVar()
        ttk.Checkbutton(grid, variable=self.vars["number_output_files"]
                        ).grid(row=5, column=1, sticky="w")

        ttk.Button(s, text="Save / Enregistrer", command=self.save_settings
                   ).pack(anchor="w", padx=18, pady=8)

    # --------------------------------------------------------------- settings io
    def _cfg(self):
        return json.loads((HERE / "config.json").read_text(encoding="utf-8"))

    def _load_settings_into_form(self):
        try:
            cfg = self._cfg()
        except Exception as e:
            messagebox.showerror("config.json", str(e))
            return
        t = cfg["id_text"]
        self.vars["prefix"].set(t.get("prefix", "ID: "))
        self.vars["show_index"].set(t.get("show_index", True))
        self.vars["index_prefix"].set(t.get("index_prefix", "No "))
        self.vars["position"].set(t.get("position", "bottom-right"))
        self.vars["font_size"].set(t.get("font_size", 48))
        self.vars["number_output_files"].set(cfg.get("number_output_files", True))

    def save_settings(self):
        cfg = self._cfg()
        t = cfg["id_text"]
        t["prefix"] = self.vars["prefix"].get()
        t["show_index"] = self.vars["show_index"].get()
        t["index_prefix"] = self.vars["index_prefix"].get()
        t["position"] = self.vars["position"].get()
        t["font_size"] = int(self.vars["font_size"].get())
        cfg["number_output_files"] = self.vars["number_output_files"].get()
        (HERE / "config.json").write_text(json.dumps(cfg, indent=2), encoding="utf-8")
        messagebox.showinfo("Saved / Enregistré",
                            "Saved. Click the green button to use the new look.\n"
                            "Enregistré. Cliquez sur le bouton vert pour "
                            "appliquer.")

    # --------------------------------------------------------------- drop logic
    def _images_dir(self):
        d = HERE / "images"
        d.mkdir(exist_ok=True)
        return d

    def pick_images(self):
        files = filedialog.askopenfilenames(
            title="Choose passport photos / Choisir les photos",
            filetypes=[("Images / PDF", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff "
                                        "*.webp *.pdf"), ("All files", "*.*")])
        if files:
            self._add_images([Path(f) for f in files])

    def pick_excel(self):
        f = filedialog.askopenfilename(
            title="Choose the name list / Choisir la liste",
            filetypes=[("Excel", "*.xls *.xlsx *.xlsm"), ("All files", "*.*")])
        if f:
            self._set_excel([Path(f)])

    def _add_images(self, paths):
        dest = self._images_dir()
        added = 0
        for p in paths:
            if p.is_dir():
                for sub in p.iterdir():
                    if sub.is_file() and sub.suffix.lower() in IMG_EXTS:
                        shutil.copy2(sub, dest / sub.name)
                        added += 1
            elif p.is_file() and p.suffix.lower() in IMG_EXTS:
                shutil.copy2(p, dest / p.name)
                added += 1
        self._refresh_status()
        if added:
            self._flash(self.zone_img["outer"])
        else:
            messagebox.showwarning(
                "No photos / Pas de photos",
                "No photo or PDF files were found in what you dropped.\n"
                "Aucune photo ou PDF trouvé dans ce que vous avez déposé.")

    def _set_excel(self, paths):
        xls = [p for p in paths if p.suffix.lower() in XLS_EXTS]
        if not xls:
            messagebox.showwarning(
                "Not an Excel file / Pas un fichier Excel",
                "Please drop an .xls or .xlsx file.\n"
                "Merci de déposer un fichier .xls ou .xlsx.")
            return
        self.excel_path = xls[0]
        self._refresh_status()
        self._flash(self.zone_xls["outer"])

    def _count_images(self):
        d = HERE / "images"
        if not d.exists():
            return 0
        return sum(1 for p in d.iterdir()
                   if p.is_file() and p.suffix.lower() in IMG_EXTS)

    def _refresh_status(self):
        n = self._count_images()
        self.zone_img["status"].config(
            text=f"{n} {tr('photos_ready', ' / ')}" if n else tr("no_photos", " / "))
        self.zone_xls["status"].config(
            text=self.excel_path.name if self.excel_path else tr("no_list", " / "))

    def _flash(self, widget, color="#16a34a"):
        old = widget.cget("highlightbackground")
        widget.config(highlightbackground=color)
        self.after(700, lambda: widget.config(highlightbackground=old))

    def clear_images(self):
        d = HERE / "images"
        n = self._count_images()
        if not n:
            messagebox.showinfo("Nothing to delete / Rien à supprimer",
                                "There are no photos here yet.\n"
                                "Il n'y a pas encore de photos ici.")
            return
        if messagebox.askyesno(
                "Delete all photos / Supprimer les photos",
                f"Delete all {n} photo(s) from this app?\n"
                f"(This does NOT touch your own copies elsewhere.)\n\n"
                f"Supprimer les {n} photo(s) de cette application ?\n"
                f"(Vos copies ailleurs ne sont PAS touchées.)"):
            for p in list(d.iterdir()):
                if p.is_file():
                    p.unlink()
            self._refresh_status()

    # -------------------------------------------------------------- preflight
    def _preflight(self):
        """Friendly checks before a run. Returns True if OK to start."""
        if self._count_images() == 0:
            messagebox.showwarning(
                "No photos / Pas de photos",
                "Put the photos in first (left box).\n"
                "Ajoutez d'abord les photos (case de gauche).")
            return False
        if not self.excel_path and not self._find_existing_excel():
            messagebox.showwarning(
                "No name list / Pas de liste",
                "Put the name list (Excel) in first (right box).\n"
                "Ajoutez d'abord la liste des noms (case de droite).")
            return False
        # report open in Excel? writing would silently go to a fallback file
        report = HERE / "report.xlsx"
        if report.exists():
            try:
                with open(report, "a"):
                    pass
            except PermissionError:
                messagebox.showwarning(
                    "Close the report / Fermez le rapport",
                    "The report is open in Excel. Close it, then click the "
                    "green button again.\n"
                    "Le rapport est ouvert dans Excel. Fermez-le, puis "
                    "cliquez à nouveau sur le bouton vert.")
                return False
        return True

    # ------------------------------------------------------------------ runners
    def _set_running(self, on):
        self.running = on
        state = "disabled" if on else "normal"
        self.btn_split.config(state=state)
        self.btn_run.config(state=state)
        if not on:
            self.progress.config(value=0)
            self.prog_label.config(text="")

    def _log(self, text):
        self.msg_q.put(("log", text))

    def on_split(self):
        if self.running:
            return
        pdfs = [p for p in (HERE / "images").glob("*")
                if p.is_file() and p.suffix.lower() == ".pdf"]
        if not pdfs:
            messagebox.showinfo(
                "No PDF files / Pas de PDF",
                "There are no PDF files to turn into photos.\n"
                "Il n'y a pas de fichiers PDF à transformer.")
            return
        self._set_running(True)
        self.log.delete("1.0", "end")
        threading.Thread(target=self._worker_split, daemon=True).start()

    def on_run(self):
        if self.running or not self._preflight():
            return
        self._set_running(True)
        self.log.delete("1.0", "end")
        threading.Thread(target=self._worker_run, daemon=True).start()

    def _redirect(self):
        app = self
        class W:
            def write(self, s):
                if s.strip():
                    app.msg_q.put(("log", s.rstrip("\n")))
            def flush(self):
                pass
        return W()

    def _worker_split(self):
        old = sys.stdout
        sys.stdout = self._redirect()
        try:
            n = sp.split_all(log=self._log)
            self.msg_q.put(("done_split", n))
        except Exception as e:
            pc.log_error("split_pdfs", e)
            self.msg_q.put(("error", str(e)))
        finally:
            sys.stdout = old

    def _worker_run(self):
        old = sys.stdout
        sys.stdout = self._redirect()
        try:
            cfg = pc.load_config()
            if self.excel_path:                       # use the chosen list
                cfg["excel_file"] = str(self.excel_path)
            results = pc.run_check(
                cfg=cfg, log=self._log,
                progress=lambda i, t, f: self.msg_q.put(("progress", (i, t, f))))
            self.msg_q.put(("done_run", results))
        except pc.CheckError as e:
            self.msg_q.put(("error_check", (str(e), e.message_fr)))
        except Exception as e:
            pc.log_error("run_check", e)
            self.msg_q.put(("error", str(e)))
        finally:
            sys.stdout = old

    # ------------------------------------------------------------------ queue
    def _drain_queue(self):
        try:
            while True:
                kind, payload = self.msg_q.get_nowait()
                if kind == "log":
                    self.log.insert("end", payload + "\n")
                    self.log.see("end")
                elif kind == "progress":
                    i, t, fname = payload
                    self.progress.config(value=(i / t * 100) if t else 0)
                    self.prog_label.config(text=f"Photo {i} / {t}")
                elif kind == "error_check":
                    self._set_running(False)
                    en, fr = payload
                    messagebox.showwarning("Please fix this / À corriger",
                                           en + ("\n\n" + fr if fr else ""))
                elif kind == "error":
                    self._set_running(False)
                    messagebox.showerror(
                        "Error / Erreur",
                        f"Something went wrong:\n{payload}\n\n"
                        f"Une erreur s'est produite.\n"
                        f"Details saved in / Détails dans : app_errors.log")
                elif kind == "done_split":
                    self._set_running(False)
                    messagebox.showinfo(
                        "PDF files done / PDF terminés",
                        f"{payload} photo(s) are in the 'PDF photos to sort' "
                        "folder.\n"
                        "1. Open that folder.  2. Give each passport the "
                        "client's name.  3. Put it with the other photos.  "
                        "4. Click the green button.\n\n"
                        f"{payload} photo(s) sont dans le dossier 'PDF à "
                        "trier'.\n"
                        "1. Ouvrez ce dossier.  2. Donnez à chaque passeport "
                        "le nom du client.  3. Mettez-le avec les autres "
                        "photos.  4. Cliquez sur le bouton vert.")
                elif kind == "done_run":
                    self._set_running(False)
                    self._show_results(payload)
                    self.nb.select(0)
                    s = payload.get("stats", {})
                    ready = s.get("stamped", 0) + s.get("index_only", 0)
                    todo = s.get("needs_manual", 0)
                    messagebox.showinfo(
                        "Finished / Terminé",
                        f"Done. {ready} photos are ready.\n"
                        + (f"{todo} need your attention (purple box).\n"
                           if todo else "")
                        + f"\nTerminé. {ready} photos sont prêtes."
                        + (f"\n{todo} demandent votre attention (case "
                           f"violette)." if todo else ""))
        except queue.Empty:
            pass
        self.after(80, self._drain_queue)

    # -------------------------------------------------------------- results ui
    def _show_results(self, r):
        s = r.get("stats", {})
        self.cards["done"].config(text=len(r["stamped"]))
        self.cards["index_only"].config(text=len(r.get("index_only", [])))
        self.cards["no_photo"].config(text=len(r["missing_image"]))
        self.cards["no_match"].config(text=len(r["not_found"]))
        self.cards["broken"].config(text=len(r.get("broken", [])))
        self.cards["pdf"].config(text=len(r["pdf_skipped"]))

        # ---- summary tab ----
        self.headline.config(
            text=f"We have photos for {s.get('received', 0)} of "
                 f"{s.get('total', 0)} people  ({s.get('percent', 0)}%)\n"
                 f"Nous avons les photos de {s.get('received', 0)} personnes "
                 f"sur {s.get('total', 0)}")
        self.stat_lines.config(text=(
            f"ID number added / Numéro ajouté :  {s.get('stamped', 0)}      "
            f"List number only / Numéro de liste seul :  "
            f"{s.get('index_only', 0)}\n"
            f"No photo yet / Pas encore de photo :  {s.get('missing', 0)}      "
            f"Photo has no match / Photo sans nom :  "
            f"{s.get('photos_not_found', 0)}\n"
            f"Problem files / Fichiers à problème :  {s.get('broken', 0)}      "
            f"PDF to sort / PDF à trier :  {s.get('pdf', 0)}"))

        # purple "check by hand" list, with one-click rename where possible
        self.manual_tree.delete(*self.manual_tree.get_children())
        self.manual_fix.clear()
        img_dir = HERE / "images"
        for item, why, why_fr, sugg in r.get("needs_manual", []):
            row = self.manual_tree.insert(
                "", "end", tags=("purple",),
                values=(item, why + "  |  " + why_fr))
            if sugg:
                src = img_dir / item
                if src.exists():
                    self.manual_fix[row] = (src, sugg)

        # ---- all photos tab ----
        self.tree.delete(*self.tree.get_children())
        self.row_file.clear()
        out_dir = HERE / "output"
        out_files = list(out_dir.iterdir()) if out_dir.exists() else []

        def find_output(name):
            key = pc.normalize(name)
            for p in out_files:
                if key in pc.clean_stem(p.stem):
                    return p
            return None

        for idx, name, cid in r["stamped"]:
            row = self.tree.insert(
                "", "end", tags=("ok",),
                values=(tr("st_done", " / "), name,
                        f"number {cid} added  (No {idx})"))
            p = find_output(name)
            if p:
                self.row_file[row] = p
        for idx, name in r.get("index_only", []):
            row = self.tree.insert(
                "", "end", tags=("blue",),
                values=(tr("st_index", " / "), name,
                        f"No {idx} added - waiting for the ID number / "
                        f"en attente du numéro"))
            p = find_output(name)
            if p:
                self.row_file[row] = p
        for fname, diag in r["not_found"]:
            row = self.tree.insert(
                "", "end", tags=("amber",),
                values=(tr("st_no_match", " / "), fname, diag["why"]))
            self.row_file[row] = HERE / "images" / fname
        for fname, why, _fr in r.get("broken", []):
            row = self.tree.insert(
                "", "end", tags=("red",),
                values=(tr("st_broken", " / "), fname, why))
            self.row_file[row] = HERE / "images" / fname
        for name in r["missing_image"]:
            self.tree.insert(
                "", "end", tags=("red",),
                values=(tr("st_no_photo", " / "), name,
                        "in the list, but no photo sent / dans la liste, "
                        "mais pas de photo"))
        for name in r["no_id"]:
            self.tree.insert(
                "", "end", tags=("grey",),
                values=(tr("st_no_number", " / "), name,
                        "no number and no list position / pas de numéro"))
        for fname in r["pdf_skipped"]:
            row = self.tree.insert(
                "", "end", tags=("grey",),
                values=(tr("st_pdf", " / "), fname,
                        "click 'turn PDFs into photos' / cliquez 'PDF → "
                        "photos'"))
            self.row_file[row] = HERE / "images" / fname
        for fname, ids in r["manual"]:
            row = self.tree.insert(
                "", "end", tags=("amber",),
                values=(tr("st_manual", " / "), fname,
                        "same name twice - choose the number / même nom deux "
                        "fois : " + ", ".join(ids)))
            self.row_file[row] = HERE / "images" / fname

    # ------------------------------------------------------- one-click rename
    def _fix_selected_name(self, _event=None):
        sel = self.manual_tree.selection()
        if not sel or sel[0] not in self.manual_fix:
            return
        src, new_name = self.manual_fix[sel[0]]
        if not src.exists():
            messagebox.showwarning("File moved / Fichier déplacé",
                                   "This photo is no longer in the folder.\n"
                                   "Cette photo n'est plus dans le dossier.")
            return
        target = src.with_name(new_name)
        if target.exists():
            messagebox.showwarning(
                "Name taken / Nom déjà pris",
                f"A file called '{new_name}' already exists.\n"
                f"Un fichier nommé '{new_name}' existe déjà.")
            return
        if messagebox.askyesno(
                "Fix the name / Corriger le nom",
                f"Rename this photo? / Renommer cette photo ?\n\n"
                f"{src.name}\n↓\n{new_name}"):
            try:
                src.rename(target)
            except OSError as e:
                pc.log_error(f"rename {src.name}", e)
                messagebox.showerror(
                    "Could not rename / Impossible de renommer",
                    "The file may be open in another program.\n"
                    "Le fichier est peut-être ouvert dans un autre programme.")
                return
            del self.manual_fix[sel[0]]
            self.manual_tree.delete(sel[0])
            self._refresh_status()
            messagebox.showinfo(
                "Renamed / Renommé",
                "Done. Click the green button again to update everything.\n"
                "C'est fait. Cliquez à nouveau sur le bouton vert pour tout "
                "mettre à jour.")

    # ---------------------------------------------------- open photo on dblclick
    def _open_row_photo(self, _event=None):
        sel = self.tree.selection()
        if sel and sel[0] in self.row_file:
            p = self.row_file[sel[0]]
            if p.exists():
                os.startfile(p)

    # ------------------------------------------------------------------- opens
    def open_path(self, sub):
        p = HERE / sub
        p.mkdir(exist_ok=True)
        os.startfile(p)

    def open_report(self):
        for name in ("report.xlsx", "report (new).xlsx", "report.txt"):
            p = HERE / name
            if p.exists():
                os.startfile(p)
                return
        messagebox.showinfo(
            "No report yet / Pas encore de rapport",
            "Click the green button first.\n"
            "Cliquez d'abord sur le bouton vert.")


if __name__ == "__main__":
    App().mainloop()
