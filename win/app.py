import os
import sys
import math
import io
import queue
import tempfile
import platform
import sqlite3
import subprocess
import threading
from datetime import datetime
from PIL import Image
import customtkinter as ctk
from tkinter import filedialog, messagebox
from pypdf import PdfReader
import requests

# Moteur de rendu PDF
try:
    import pymupdf
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

SYSTEM_OS = platform.system()

# URL Webhook Google Apps Script validée
GOOGLE_SHEET_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbz4e6qYq47cUV_xHfFA12rJ44uqFxhlgpfbmBjO3AcgUWKxYTm7YVsZVqlhumJAdJy9/exec"

# ----------------- 1. BASE DE DONNÉES LOCALE -----------------
def init_db():
    conn = sqlite3.connect("printpay.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS teachers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nom TEXT UNIQUE NOT NULL
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS print_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_name TEXT,
            filename TEXT,
            filepath TEXT,
            page_range TEXT,
            pages_count INTEGER,
            copies INTEGER,
            total_feuilles INTEGER,
            mode TEXT,
            type_impression TEXT,
            cout_total INTEGER,
            date_time TEXT,
            is_synced INTEGER DEFAULT 0
        )
    """)
    
    # N'insère les profs par défaut QUE si la table est totalement vide
    cursor.execute("SELECT COUNT(*) FROM teachers")
    if cursor.fetchone()[0] == 0:
        for p in ['papier gache', 'photocopie']:
            cursor.execute("INSERT INTO teachers (nom) VALUES (?)", (p,))
            
    conn.commit()
    conn.close()

# ----------------- 2. UTILITAIRES SYSTÈME & DÉTECTION SUMATRA -----------------
def find_sumatra_executable():
    """Détecte n'importe quel exécutable SumatraPDF (SumatraPDF.exe, SumatraPDF-3.6-64.exe, etc.)"""
    # 1. Dans le bundle temporaire PyInstaller
    if hasattr(sys, '_MEIPASS'):
        for f in os.listdir(sys._MEIPASS):
            if "sumatra" in f.lower() and f.lower().endswith(".exe"):
                return os.path.join(sys._MEIPASS, f)
        
    # 2. Dossiers locaux et système Windows
    base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    search_dirs = [
        base_dir,
        os.path.join(base_dir, "win"),
        os.getcwd(),
        os.path.join(os.getcwd(), "win"),
        r"C:\Program Files\SumatraPDF",
        r"C:\Program Files (x86)\SumatraPDF",
        os.path.expandvars(r"%LOCALAPPDATA%\SumatraPDF")
    ]
    for d in search_dirs:
        if os.path.exists(d):
            for f in os.listdir(d):
                if "sumatra" in f.lower() and f.lower().endswith(".exe"):
                    return os.path.join(d, f)
    return None

def parse_page_range(range_str, max_pages=None):
    pages = set()
    cleaned = range_str.strip().replace(" ", "")
    if not cleaned:
        return [], 0
    try:
        for part in cleaned.split(","):
            if "-" in part:
                sub = part.split("-")
                start, end = int(sub[0]), int(sub[1])
                if start <= end:
                    for p in range(start, end + 1):
                        if max_pages is None or p <= max_pages:
                            pages.add(p)
            elif part.isdigit():
                p = int(part)
                if max_pages is None or p <= max_pages:
                    pages.add(p)
        sorted_pages = sorted(list(pages))
        return sorted_pages, len(sorted_pages)
    except Exception:
        return [], 0

# FONCTION GLOBALE DE DÉCOUPE, ROTATION ET ESPACE DE COULEUR
def create_prepared_print_pdf(source_pdf, pages_to_keep, is_color=False):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    temp_pdf = os.path.join(tempfile.gettempdir(), f"printpay_job_{timestamp}.pdf")
    
    doc = pymupdf.open(source_pdf)
    output_doc = pymupdf.open()

    for p_num in pages_to_keep:
        if 1 <= p_num <= len(doc):
            page = doc[p_num - 1]
            rect = page.rect
            
            # Détection : la page est-elle en mode Paysage ?
            is_landscape = rect.width > rect.height

            # Si paysage : rotation de 90° pour s'aligner sur la feuille A4
            if is_landscape:
                page.set_rotation((page.rotation + 90) % 360)

            # Rendu haute définition 300 DPI : RGB réel si couleur, GRAY pur si N&B
            c_space = pymupdf.csRGB if is_color else pymupdf.csGRAY
            pix = page.get_pixmap(dpi=300, colorspace=c_space)
            img_bytes = pix.tobytes("png")
            
            new_page = output_doc.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(new_page.rect, stream=img_bytes)

    output_doc.save(temp_pdf)
    output_doc.close()
    doc.close()
    return temp_pdf

def get_printers():
    printers = ["🧪 Mode Test (Simulation sans imprimante)"]
    try:
        if SYSTEM_OS == "Windows":
            import win32print
            printer_data = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
            for p in printer_data:
                printers.append(p[2])
        elif SYSTEM_OS == "Darwin":
            res = subprocess.run(['lpstat', '-e'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            if res.returncode == 0 and res.stdout.strip():
                printers.extend(res.stdout.strip().splitlines())
    except Exception as e:
        print(f"Erreur détection imprimantes: {e}")
    return printers

def get_app_recent_files():
    conn = sqlite3.connect("printpay.db")
    c = conn.cursor()
    c.execute("SELECT DISTINCT filename, filepath FROM print_logs WHERE filepath IS NOT NULL AND filepath != '' ORDER BY id DESC LIMIT 8")
    rows = c.fetchall()
    conn.close()
    return rows

def count_pages(filepath):
    if not filepath or not os.path.exists(filepath):
        return 1
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".pdf":
        try:
            reader = PdfReader(filepath)
            return len(reader.pages)
        except Exception:
            return 1
    return 1

# ----------------- 3. INTERFACE GRAPHIQUE -----------------
class PrintPayApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("PrintPay - Station d'Impression")
        self.geometry("1100x740")
        self.resizable(False, False)

        init_db()
        self.selected_filepath = ""
        self.total_doc_pages = 1
        self.selected_pages_list = [1]
        self.current_selection_index = 0

        self.sync_queue = queue.Queue()

        self.setup_ui()
        self.refresh_recent_dropdown()

        self.poll_sync_queue()
        self.after(2000, self.schedule_next_sync)

    def poll_sync_queue(self):
        try:
            while True:
                text, color = self.sync_queue.get_nowait()
                self.lbl_sync_status.configure(text=text, text_color=color)
        except queue.Empty:
            pass
        self.after(400, self.poll_sync_queue)

    def schedule_next_sync(self):
        threading.Thread(target=self.push_pending_batch, daemon=True).start()
        self.after(15000, self.schedule_next_sync)

    def setup_ui(self):
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(15, 10))
        
        self.header = ctk.CTkLabel(header_frame, text="🖨️ PrintPay - Impression Scolaire", font=ctk.CTkFont(size=22, weight="bold"))
        self.header.pack(side="left")

        self.lbl_sync_status = ctk.CTkLabel(header_frame, text="☁️ Prêt", text_color="#10B981", font=ctk.CTkFont(size=12))
        self.lbl_sync_status.pack(side="right", padx=10)

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=15, pady=0)

        # ---------------- GAUCHE : FORMULAIRE ----------------
        left_frame = ctk.CTkFrame(body, width=580)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

        # 1. Enseignant
        ctk.CTkLabel(left_frame, text="1. Sélectionner l'Enseignant :", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(15, 2))

        teacher_box = ctk.CTkFrame(left_frame, fg_color="transparent")
        teacher_box.pack(anchor="w", padx=15, pady=2, fill="x")

        self.teacher_combo = ctk.CTkComboBox(teacher_box, values=self.get_teachers(), width=320, height=35)
        self.teacher_combo.pack(side="left", padx=(0, 8))

        self.btn_add_teacher = ctk.CTkButton(teacher_box, text="➕", width=45, height=35, fg_color="#475569", command=self.add_new_teacher)
        self.btn_add_teacher.pack(side="left", padx=(0, 5))

        self.btn_del_teacher = ctk.CTkButton(teacher_box, text="🗑️", width=45, height=35, fg_color="#DC2626", hover_color="#B91C1C", command=self.delete_current_teacher)
        self.btn_del_teacher.pack(side="left")

        # 2. Document
        ctk.CTkLabel(left_frame, text="2. Document à imprimer :", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(12, 2))

        doc_box = ctk.CTkFrame(left_frame, fg_color="transparent")
        doc_box.pack(anchor="w", padx=15, pady=2, fill="x")

        self.btn_browse = ctk.CTkButton(doc_box, text="📂 Parcourir...", command=self.browse_file, width=120, height=35)
        self.btn_browse.pack(side="left", padx=(0, 10))

        self.recent_dict = {}
        self.recent_combo = ctk.CTkComboBox(doc_box, values=["-- Historique vide --"], width=320, height=35, command=self.on_recent_selected)
        self.recent_combo.pack(side="left")

        self.lbl_filename = ctk.CTkLabel(left_frame, text="Aucun fichier sélectionné", text_color="#94A3B8")
        self.lbl_filename.pack(anchor="w", padx=15, pady=(2, 8))

        # 3. Options d'impression
        ctk.CTkLabel(left_frame, text="3. Options d'impression & Tarifs :", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(5, 2))

        opts_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        opts_frame.pack(anchor="w", padx=15, pady=2)

        self.recto_verso_var = ctk.StringVar(value="recto")
        self.rb_recto = ctk.CTkRadioButton(opts_frame, text="Recto (15 Fr)", variable=self.recto_verso_var, value="recto", command=self.update_total)
        self.rb_recto.pack(side="left", padx=(0, 15))
        
        self.rb_verso = ctk.CTkRadioButton(opts_frame, text="Recto-Verso (25 Fr)", variable=self.recto_verso_var, value="duplex", command=self.update_total)
        self.rb_verso.pack(side="left", padx=15)

        # Case Couleur reliée au basculement en direct de l'aperçu
        self.color_var = ctk.BooleanVar(value=False)
        self.chk_color = ctk.CTkCheckBox(
            opts_frame, 
            text="Couleur (+25 Fr)", 
            variable=self.color_var, 
            command=self.on_color_toggle
        )
        self.chk_color.pack(side="left", padx=15)

        # Plage & Copies
        params_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        params_frame.pack(anchor="w", padx=15, pady=8)

        ctk.CTkLabel(params_frame, text="Plage de pages :").pack(side="left", padx=(0, 5))
        self.entry_pages = ctk.CTkEntry(params_frame, width=90)
        self.entry_pages.insert(0, "1")
        self.entry_pages.pack(side="left", padx=5)
        self.entry_pages.bind("<KeyRelease>", lambda e: self.on_range_change())

        ctk.CTkLabel(params_frame, text="Copies :").pack(side="left", padx=(20, 5))
        self.btn_minus = ctk.CTkButton(params_frame, text="-", width=30, height=28, command=self.decrement_copies)
        self.btn_minus.pack(side="left", padx=2)

        self.entry_copies = ctk.CTkEntry(params_frame, width=45, justify="center")
        self.entry_copies.insert(0, "1")
        self.entry_copies.pack(side="left", padx=2)
        self.entry_copies.bind("<KeyRelease>", lambda e: self.update_total())

        self.btn_plus = ctk.CTkButton(params_frame, text="+", width=30, height=28, command=self.increment_copies)
        self.btn_plus.pack(side="left", padx=2)

        # 4. Imprimante
        ctk.CTkLabel(left_frame, text="4. Imprimante disponible :", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(10, 2))

        imp_box = ctk.CTkFrame(left_frame, fg_color="transparent")
        imp_box.pack(anchor="w", padx=15, pady=2, fill="x")

        printers_list = get_printers()
        self.printer_combo = ctk.CTkComboBox(imp_box, values=printers_list, width=360, height=35)
        self.printer_combo.set(printers_list[1] if len(printers_list) > 1 else printers_list[0])
        self.printer_combo.pack(side="left", padx=(0, 10))

        self.btn_refresh_printers = ctk.CTkButton(imp_box, text="🔄 Détecter", width=80, height=35, fg_color="#334155", command=self.refresh_printers)
        self.btn_refresh_printers.pack(side="left")

        # Résumé
        self.total_card = ctk.CTkFrame(left_frame, fg_color="#0F172A", corner_radius=10)
        self.total_card.pack(fill="x", padx=15, pady=12)

        self.lbl_resume = ctk.CTkLabel(self.total_card, text="Pages: 1 | Feuilles: 1 | Coût Total: 15 FCFA", font=ctk.CTkFont(size=15, weight="bold"), text_color="#38BDF8")
        self.lbl_resume.pack(pady=10)

        # Bouton IMPRIMER
        self.btn_print = ctk.CTkButton(
            left_frame, 
            text="🖨️ IMPRIMER", 
            font=ctk.CTkFont(size=17, weight="bold"), 
            fg_color="#10B981", hover_color="#059669", 
            height=48, 
            command=self.execute_print
        )
        self.btn_print.pack(fill="x", padx=15, pady=(5, 15))

        # ---------------- DROITE : APERÇU EN DIRECT ----------------
        right_frame = ctk.CTkFrame(body, width=480)
        right_frame.pack(side="right", fill="both", expand=True)

        ctk.CTkLabel(right_frame, text="👁️ Aperçu de la Sélection", font=ctk.CTkFont(size=15, weight="bold")).pack(pady=(12, 5))

        self.preview_canvas = ctk.CTkLabel(right_frame, text="Chargez un document pour voir l'aperçu", width=380, height=480, fg_color="#1E293B", corner_radius=8)
        self.preview_canvas.pack(pady=10, padx=15)

        nav_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        nav_frame.pack(pady=5)

        self.btn_prev_page = ctk.CTkButton(nav_frame, text="◀ Précédente", width=110, command=self.prev_preview_page)
        self.btn_prev_page.pack(side="left", padx=5)

        self.lbl_preview_page = ctk.CTkLabel(nav_frame, text="Page 0 / 0", font=ctk.CTkFont(weight="bold"))
        self.lbl_preview_page.pack(side="left", padx=10)

        self.btn_next_page = ctk.CTkButton(nav_frame, text="Suivante ▶", width=110, command=self.next_preview_page)
        self.btn_next_page.pack(side="left", padx=5)

    # ---------------- BASULEMENT DYNAMIQUE COULEUR & APERÇU ----------------
    def on_color_toggle(self):
        """Met à jour le total ET bascule l'aperçu immédiatement en N&B ou Couleur"""
        self.update_total()
        self.render_preview()

    def on_range_change(self):
        self.update_total()
        range_str = self.entry_pages.get().strip()
        pages, _ = parse_page_range(range_str, max_pages=self.total_doc_pages)
        if pages:
            self.selected_pages_list = pages
            self.current_selection_index = 0
            self.render_preview()

    def render_preview(self):
        if not HAS_PYMUPDF or not self.selected_filepath or not self.selected_filepath.lower().endswith(".pdf"):
            self.preview_canvas.configure(image=None, text="Aperçu disponible pour PDF")
            return

        if not self.selected_pages_list:
            return

        physical_page = self.selected_pages_list[self.current_selection_index]

        try:
            with pymupdf.open(self.selected_filepath) as doc:
                if physical_page > len(doc):
                    return
                page = doc.load_page(physical_page - 1)
                
                # Espace de couleur dynamique selon la case cochée
                is_color = self.color_var.get()
                c_space = pymupdf.csRGB if is_color else pymupdf.csGRAY
                
                pix = page.get_pixmap(dpi=85, colorspace=c_space)
                img_data = pix.tobytes("png")
                pil_img = Image.open(io.BytesIO(img_data))
                pil_img.thumbnail((360, 480))
                
                ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(pil_img.width, pil_img.height))
                self.preview_canvas.configure(image=ctk_img, text="")
                
                total_in_range = len(self.selected_pages_list)
                curr_in_range = self.current_selection_index + 1
                mode_tag = " [Couleur]" if is_color else " [N&B]"
                self.lbl_preview_page.configure(
                    text=f"Page {curr_in_range} / {total_in_range} (Feuille {physical_page}){mode_tag}"
                )
        except Exception as e:
            self.preview_canvas.configure(image=None, text=f"Erreur d'aperçu : {e}")

    def next_preview_page(self):
        if self.current_selection_index < len(self.selected_pages_list) - 1:
            self.current_selection_index += 1
            self.render_preview()

    def prev_preview_page(self):
        if self.current_selection_index > 0:
            self.current_selection_index -= 1
            self.render_preview()

    # ---------------- CONTRÔLES FORMULAIRE ----------------
    def refresh_recent_dropdown(self):
        recent_rows = get_app_recent_files()
        if not recent_rows:
            self.recent_combo.configure(values=["-- Historique vide --"])
            self.recent_combo.set("-- Historique vide --")
            self.recent_dict = {}
        else:
            names = [r[0] for r in recent_rows]
            self.recent_dict = {r[0]: r[1] for r in recent_rows}
            self.recent_combo.configure(values=["-- Fichiers récents --"] + names)
            self.recent_combo.set("-- Fichiers récents --")

    def on_recent_selected(self, choice):
        if choice in self.recent_dict and os.path.exists(self.recent_dict[choice]):
            self.load_file(self.recent_dict[choice])

    def refresh_printers(self):
        printers = get_printers()
        self.printer_combo.configure(values=printers)
        self.printer_combo.set(printers[1] if len(printers) > 1 else printers[0])
        messagebox.showinfo("Imprimantes", f"{len(printers) - 1} imprimante(s) trouvée(s).")

    def get_teachers(self):
        conn = sqlite3.connect("printpay.db")
        c = conn.cursor()
        c.execute("SELECT nom FROM teachers ORDER BY nom ASC")
        profs = [r[0] for r in c.fetchall()]
        conn.close()
        return profs

    def add_new_teacher(self):
        dialog = ctk.CTkInputDialog(text="Nom et matière de l'enseignant :", title="Ajout Enseignant")
        nom = dialog.get_input()
        if nom and nom.strip():
            conn = sqlite3.connect("printpay.db")
            c = conn.cursor()
            try:
                c.execute("INSERT INTO teachers (nom) VALUES (?)", (nom.strip(),))
                conn.commit()
                self.teacher_combo.configure(values=self.get_teachers())
                self.teacher_combo.set(nom.strip())
            except sqlite3.IntegrityError:
                messagebox.showwarning("Doublon", "Cet enseignant existe déjà.")
            conn.close()

    def delete_current_teacher(self):
        teacher = self.teacher_combo.get()
        if not teacher:
            return

        conf = messagebox.askyesno(
            "Confirmer la suppression", 
            f"Voulez-vous vraiment retirer '{teacher}' de la liste des enseignants ?"
        )
        if conf:
            conn = sqlite3.connect("printpay.db")
            c = conn.cursor()
            c.execute("DELETE FROM teachers WHERE nom = ?", (teacher,))
            conn.commit()
            conn.close()

            profs = self.get_teachers()
            self.teacher_combo.configure(values=profs)
            if profs:
                self.teacher_combo.set(profs[0])
            messagebox.showinfo("Supprimé", f"'{teacher}' a été retiré de la liste.")

    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Documents", "*.pdf *.docx *.doc *.xlsx")])
        if path:
            self.load_file(path)

    def load_file(self, path):
        self.selected_filepath = path
        name = os.path.basename(path)
        self.lbl_filename.configure(text=f"📄 {name}", text_color="#38BDF8")
        self.total_doc_pages = count_pages(path)
        
        self.entry_pages.delete(0, 'end')
        range_val = f"1-{self.total_doc_pages}" if self.total_doc_pages > 1 else "1"
        self.entry_pages.insert(0, range_val)
        
        pages, _ = parse_page_range(range_val, self.total_doc_pages)
        self.selected_pages_list = pages
        self.current_selection_index = 0
        
        self.update_total()
        self.render_preview()

    def increment_copies(self):
        try:
            val = int(self.entry_copies.get()) + 1
        except ValueError:
            val = 1
        self.entry_copies.delete(0, 'end')
        self.entry_copies.insert(0, str(val))
        self.update_total()

    def decrement_copies(self):
        try:
            val = max(1, int(self.entry_copies.get()) - 1)
        except ValueError:
            val = 1
        self.entry_copies.delete(0, 'end')
        self.entry_copies.insert(0, str(val))
        self.update_total()

    def update_total(self):
        range_str = self.entry_pages.get()
        _, nb_pages = parse_page_range(range_str, self.total_doc_pages)
        if nb_pages == 0:
            nb_pages = 1

        try:
            copies = max(1, int(self.entry_copies.get()))
        except ValueError:
            copies = 1

        is_duplex = (self.recto_verso_var.get() == "duplex")
        is_color = self.color_var.get()

        if is_duplex:
            feuilles = math.ceil(nb_pages / 2) * copies
            unit_price = 50 if is_color else 25
        else:
            feuilles = nb_pages * copies
            unit_price = 25 if is_color else 15

        cout = feuilles * unit_price
        self.calculated_pages = nb_pages
        self.calculated_feuilles = feuilles
        self.calculated_cout = cout
        
        mode_txt = "Recto-Verso" if is_duplex else "Recto"
        color_txt = " + Couleur" if is_color else ""
        self.lbl_resume.configure(
            text=f"Pages: {nb_pages}  |  Feuilles: {feuilles} ({mode_txt}{color_txt})  |  Coût Total: {cout:,} FCFA".replace(",", " ")
        )

    # ---------------- IMPRESSION PHYSIQUE ----------------
    def execute_print(self):
        teacher = self.teacher_combo.get()
        if not teacher:
            messagebox.showerror("Erreur", "Sélectionnez un enseignant !")
            return

        printer = self.printer_combo.get()
        is_simulation = "Mode Test" in printer

        range_str = self.entry_pages.get().strip()
        doc_name = os.path.basename(self.selected_filepath) if self.selected_filepath else "Document direct"
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        is_duplex = (self.recto_verso_var.get() == "duplex")
        is_color = self.color_var.get()
        copies = int(self.entry_copies.get())
        mode_str = "Recto-Verso" if is_duplex else "Recto Simple"
        type_str = "Couleur" if is_color else "Noir & Blanc"

        # 1. Enregistrement local SQLite
        conn = sqlite3.connect("printpay.db")
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO print_logs (teacher_name, filename, filepath, page_range, pages_count, copies, total_feuilles, mode, type_impression, cout_total, date_time, is_synced)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0)
        """, (
            teacher, doc_name, self.selected_filepath, range_str, self.calculated_pages,
            copies, self.calculated_feuilles,
            mode_str, type_str, self.calculated_cout, now
        ))
        conn.commit()
        conn.close()

        self.refresh_recent_dropdown()

        # 2. Impression physique
        if not is_simulation and self.selected_filepath and os.path.exists(self.selected_filepath):
            target_pages, _ = parse_page_range(range_str, self.total_doc_pages)
            if not target_pages:
                target_pages = [1]
            
            # APPEL DE LA FONCTION GLOBALE AVEC COULEUR EXACTE
            ready_to_print_pdf = create_prepared_print_pdf(self.selected_filepath, target_pages, is_color=is_color)

            try:
                if SYSTEM_OS == "Darwin":  # macOS (CUPS lp)
                    options = ["-d", printer, "-n", str(copies), "-o", "fit-to-page"]
                    if is_duplex:
                        options.extend(["-o", "sides=two-sided-long-edge"])
                    else:
                        options.extend(["-o", "sides=one-sided"])
                    
                    if is_color:
                        options.extend(["-o", "print-color-mode=color", "-o", "ColorModel=RGB", "-o", "CNColorMode=color"])
                    else:
                        options.extend(["-o", "print-color-mode=monochrome", "-o", "ColorModel=Gray", "-o", "CNColorMode=mono"])

                    subprocess.run(["lp"] + options + [ready_to_print_pdf])

                elif SYSTEM_OS == "Windows":
                    sumatra = find_sumatra_executable()
                    if sumatra:
                        duplex_opt = "duplex" if is_duplex else "simplex"
                        copies_opt = f"{copies}x"
                        # Ordre strict : color ou monochrome
                        color_opt = "color" if is_color else "monochrome"
                        settings = f"{duplex_opt},{copies_opt},{color_opt}"
                        cmd = f'"{sumatra}" -print-to "{printer}" -print-settings "{settings}" "{ready_to_print_pdf}"'
                        print(f"[PrintPay] Exécution : {cmd}")
                        subprocess.run(cmd, shell=True)
                    else:
                        print("[PrintPay] Sumatra introuvable, fallback vers printto")
                        import win32api
                        win32api.ShellExecute(0, "printto", ready_to_print_pdf, f'"{printer}"', ".", 0)

            except Exception as e:
                print(f"Erreur d'impression physique: {e}")

            success_msg = f"✅ Impression envoyée à '{printer}' !\nMode : {type_str}\n{self.calculated_cout} FCFA débités pour {teacher}."
        else:
            success_msg = f"🧪 [TEST RÉUSSI] Impression simulée !\n{self.calculated_cout} FCFA enregistrés pour {teacher}."

        messagebox.showinfo("Succès", success_msg)

        # 3. Synchronisation par lot en arrière-plan
        threading.Thread(target=self.push_pending_batch, daemon=True).start()

    # ---------------- SYNCHRONISATION MULTITHREAD SÉCURISÉE ----------------
    def push_pending_batch(self):
        try:
            conn = sqlite3.connect("printpay.db")
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, date_time, teacher_name, filename, page_range, total_feuilles, mode, type_impression, cout_total 
                FROM print_logs 
                WHERE is_synced = 0 
                ORDER BY id ASC
            """)
            rows = cursor.fetchall()

            if not rows:
                self.sync_queue.put(("☁️ Google Sheets à jour", "#10B981"))
                conn.close()
                return

            batch_payload = {"rows": [list(r) for r in rows]}

            response = requests.post(
                GOOGLE_SHEET_WEBHOOK_URL, 
                json=batch_payload, 
                timeout=25, 
                allow_redirects=True
            )
            
            if response.status_code in [200, 201]:
                res_json = response.json()
                synced_ids = res_json.get("synced_ids", [])
                if synced_ids:
                    cursor.executemany("UPDATE print_logs SET is_synced = 1 WHERE id = ?", [(i,) for i in synced_ids])
                    conn.commit()
                self.sync_queue.put(("☁️ Synchronisé", "#10B981"))
            else:
                self.sync_queue.put((f"⚠️ Erreur Google ({response.status_code})", "#F59E0B"))
            conn.close()
        except Exception:
            self.sync_queue.put(("⚠️ En attente connexion", "#F59E0B"))

# ----------------- 4. POINT D'ENTRÉE PRINCIPAL -----------------
if __name__ == "__main__":
    app = PrintPayApp()
    app.mainloop()