import os
import sys
import subprocess
import tempfile
import re
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageTk, ImageDraw, ImageFont
from ttkbootstrap import Style
from ttkbootstrap.widgets import Button, Label, Entry, Frame
from docx import Document
import fitz  # PyMuPDF
from datetime import datetime
import json
import pytesseract  # OCR avec Tesseract

# ---------- Config Tesseract ----------
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\ofriha\Desktop\projet\tesseract-5.5.1\tesseract.exe"

# ---------- Gestion historique des mots-clés ----------
HISTORIQUE_FICHIER = "historique_recherches.json"

def charger_historique():
    if os.path.exists(HISTORIQUE_FICHIER):
        try:
            with open(HISTORIQUE_FICHIER, "r", encoding="utf-8") as f:
                mots = json.load(f)
                if isinstance(mots, list):
                    return mots
        except Exception as e:
            print(f"Erreur lecture historique : {e}")
    return []

def sauvegarder_historique(mots):
    try:
        with open(HISTORIQUE_FICHIER, "w", encoding="utf-8") as f:
            json.dump(mots, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Erreur sauvegarde historique : {e}")

historique_mots = charger_historique()

# ---------- Gestion du chemin des ressources pour l'exe ----------
def resource_path(relative_path):
    """Retourne le chemin absolu correct de la ressource, que l’on soit dans l’exe ou dans le code Python."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ----------- Lecture des fichiers -----------
def lire_docx(path):
    try:
        doc = Document(path)
        return "\n".join([para.text for para in doc.paragraphs])
    except:
        return ""

def lire_pdf(path):
    try:
        with fitz.open(path) as doc:
            return "\n".join([page.get_text() for page in doc])
    except:
        return ""

def lire_image_ocr(path):
    try:
        image = Image.open(path)
        texte = pytesseract.image_to_string(image, lang='fra')
        return texte
    except:
        return ""

def texte_en_image(texte, largeur=200, hauteur=240, max_chars_per_line=40, max_lines=10, font_path="arial.ttf"):
    image = Image.new("RGB", (largeur, hauteur), color="white")
    draw = ImageDraw.Draw(image)

    lignes = []
    for ligne in texte.strip().split("\n"):
        while len(ligne) > max_chars_per_line:
            lignes.append(ligne[:max_chars_per_line])
            ligne = ligne[max_chars_per_line:]
        lignes.append(ligne)
    lignes = lignes[:max_lines]

    texte_reduit = "\n".join(lignes)
    taille_police = 14
    try:
        font = ImageFont.truetype(font_path, taille_police)
    except:
        font = ImageFont.load_default()

    while True:
        bbox = draw.multiline_textbbox((0, 0), texte_reduit, font=font, spacing=4)
        largeur_texte = bbox[2] - bbox[0]
        hauteur_texte = bbox[3] - bbox[1]

        if largeur_texte <= (largeur - 20) and hauteur_texte <= (hauteur - 20):
            break
        taille_police -= 1
        if taille_police < 8:
            break
        try:
            font = ImageFont.truetype(font_path, taille_police)
        except:
            font = ImageFont.load_default()

    x = (largeur - largeur_texte) // 2
    y = (hauteur - hauteur_texte) // 2
    draw.multiline_text((x, y), texte_reduit, fill="black", font=font, spacing=4)
    return image

def creer_aperçu_image(path):
    try:
        if path.lower().endswith(".pdf"):
            doc = fitz.open(path)
            if len(doc) > 0:
                pix = doc[0].get_pixmap()
                return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        elif path.lower().endswith(".docx"):
            return texte_en_image(lire_docx(path), largeur=200, hauteur=240)
        elif path.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff")):
            return texte_en_image(lire_image_ocr(path), largeur=200, hauteur=240)
    except:
        return None

# ----------- Recherche -----------
def rechercher_documents(dossier, mot_cle):
    resultats = []
    for root, _, files in os.walk(dossier):
        for file in files:
            chemin_complet = os.path.join(root, file)
            if file.endswith(".docx"):
                texte = lire_docx(chemin_complet)
            elif file.endswith(".pdf"):
                texte = lire_pdf(chemin_complet)
            elif file.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff")):
                texte = lire_image_ocr(chemin_complet)
            else:
                continue
            if mot_cle.lower() in texte.lower():
                date_modif = os.path.getmtime(chemin_complet)
                resultats.append((root, file, date_modif))
    return resultats

def parcourir_dossier():
    dossier = filedialog.askdirectory()
    champ_dossier.delete(0, tk.END)
    champ_dossier.insert(0, dossier)

# ----------- Highlight DOCX & PDF -----------
def highlight_word_docx(input_path, output_path, mot_cle):
    doc = Document(input_path)
    pattern = re.compile(re.escape(mot_cle), re.IGNORECASE)
    for para in doc.paragraphs:
        inline = para.runs
        new_runs = []
        for run in inline:
            text = run.text
            matches = list(pattern.finditer(text))
            if not matches:
                new_runs.append(run)
                continue
            last_end = 0
            for m in matches:
                if m.start() > last_end:
                    new_run = para.add_run(text[last_end:m.start()])
                    new_runs.append(new_run)
                new_run = para.add_run(text[m.start():m.end()])
                new_run.font.highlight_color = 7
                new_runs.append(new_run)
                last_end = m.end()
            if last_end < len(text):
                new_run = para.add_run(text[last_end:])
                new_runs.append(new_run)
            run.text = ""
        for r in inline:
            p = r._element
            p.getparent().remove(p)
        for r in new_runs:
            para._p.append(r._element)
    doc.save(output_path)

def highlight_word_pdf(input_path, output_path, mot_cle):
    doc = fitz.open(input_path)
    for page in doc:
        text_instances = page.search_for(mot_cle, flags=fitz.TEXT_DEHYPHENATE)
        for inst in text_instances:
            page.add_highlight_annot(inst)
    doc.save(output_path)

# ----------- Ouverture fichiers -----------
def ouvrir_fichier(path, mot_cle=None):
    try:
        suffix = os.path.splitext(path)[1].lower()
        if mot_cle:
            tmp_dir = tempfile.gettempdir()
            basename = os.path.basename(path)
            tmp_path = os.path.join(tmp_dir, f"tmp_highlight_{basename}")
            if suffix == ".docx":
                highlight_word_docx(path, tmp_path, mot_cle)
            elif suffix == ".pdf":
                highlight_word_pdf(path, tmp_path, mot_cle)
            else:
                tmp_path = path
        else:
            tmp_path = path

        if sys.platform.startswith('darwin'):
            subprocess.call(('open', tmp_path))
        elif os.name == 'nt':
            os.startfile(tmp_path)
        elif os.name == 'posix':
            subprocess.call(('xdg-open', tmp_path))
    except Exception as e:
        print(f"Erreur ouverture fichier : {e}")

# ----------- Affichage résultats -----------
def lancer_recherche():
    global resultats_recherche, historique_mots
    dossier = champ_dossier.get()
    mot_cle = champ_mot_cle.get().strip()
    if not os.path.isdir(dossier):
        messagebox.showerror("Erreur", "Le dossier spécifié n'existe pas.")
        return
    if not mot_cle:
        messagebox.showerror("Erreur", "Le mot-clé ne peut pas être vide.")
        return
    if mot_cle and mot_cle not in historique_mots:
        historique_mots.append(mot_cle)
        sauvegarder_historique(historique_mots)
        champ_mot_cle['values'] = historique_mots
    resultats_recherche = rechercher_documents(dossier, mot_cle)
    trier_afficher_resultats()

def trier_afficher_resultats(*args):
    ordre = choix_tri.get()
    inverser = (ordre == "Plus récent → Plus ancien")
    resultats_tries = sorted(resultats_recherche, key=lambda x: x[2], reverse=inverser)
    afficher_resultats(resultats_tries)

def afficher_resultats(resultats):
    for widget in cadre_resultats.winfo_children():
        widget.destroy()
    for dossier, fichier, date_modif in resultats:
        chemin = os.path.join(dossier, fichier)
        date_str = datetime.fromtimestamp(date_modif).strftime("%d/%m/%Y %H:%M")
        largeur_cadre = 230 if fichier.lower().endswith(".pdf") else 220
        cadre = Frame(
            cadre_resultats,
            width=largeur_cadre,
            height=400,
            padding=15,
            style="Result.TFrame",
            borderwidth=2,
            relief="solid"
        )
        cadre.pack(side="left", padx=12, pady=15)
        cadre.pack_propagate(False)
        img = creer_aperçu_image(chemin)
        if img:
            img = img.resize((200, 240), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            lbl_img = Label(cadre, image=photo, style="Result.TLabel")
            lbl_img.image = photo
            lbl_img.pack()
        Label(
            cadre,
            text=fichier,
            font=("Poppins", 11, "bold"),
            wraplength=largeur_cadre - 40,
            justify="center",
            style="Result.TLabel"
        ).pack(pady=5)
        Label(
            cadre,
            text=f"📅 {date_str}",
            font=("Roboto", 9),
            style="Result.TLabel"
        ).pack(side="bottom", pady=5)
        cadre.bind("<Double-Button-1>", lambda e, p=chemin, m=champ_mot_cle.get(): ouvrir_fichier(p, m))
        if img:
            lbl_img.bind("<Double-Button-1>", lambda e, p=chemin, m=champ_mot_cle.get(): ouvrir_fichier(p, m))

# ----------- Défilement horizontal -----------
def _on_mousewheel(event):
    delta = int(-1*(event.delta/120))
    canvas.xview_scroll(delta, "units")

def _on_mousewheel_linux(event):
    if event.num == 4:
        canvas.xview_scroll(-1, "units")
    elif event.num == 5:
        canvas.xview_scroll(1, "units")

# ----------- Fenêtre principale -----------
style = Style("darkly")
fenetre = style.master
fenetre.title("🔍 Recherche de mots-clés")
fenetre.geometry("1280x750")

# -- Icône --
try:
    fenetre.iconbitmap(resource_path("icone.ico"))
except Exception:
    print("Icône introuvable ou incompatible, utilisation de l'icône par défaut.")

# --- Titre INVENIO ---
img_title_path = resource_path("invenioo.png")
try:
    image_title = Image.open(img_title_path)
    photo_title = ImageTk.PhotoImage(image_title)
    Label(fenetre, image=photo_title, background=fenetre.cget("background")).pack(pady=(30, 20))
except Exception as e:
    print(f"⚠️ Impossible de charger l'image INVENIO : {e}")

# --- Formulaire ---
cadre_form = Frame(fenetre, style="Form.TFrame")
cadre_form.pack(pady=5, padx=30, fill="x")

Label(cadre_form, text="📁 Dossier :", font=("Roboto", 10, "bold"), style="TLabel").grid(row=0, column=0, sticky="e", padx=8, pady=6)
champ_dossier = Entry(cadre_form, width=70, style="TEntry", font=("Roboto", 10))
champ_dossier.grid(row=0, column=1, padx=8, pady=6, sticky="ew")
Button(cadre_form, text="Parcourir", command=parcourir_dossier, style="TButton", bootstyle="info").grid(row=0, column=2, padx=8, pady=6)

Label(cadre_form, text="🔑 Mot-clé :", font=("Roboto", 10, "bold"), style="TLabel").grid(row=1, column=0, sticky="e", padx=8, pady=6)
champ_mot_cle = ttk.Combobox(cadre_form, width=70, style="TCombobox", font=("Roboto", 10))
champ_mot_cle['values'] = historique_mots
champ_mot_cle.grid(row=1, column=1, padx=8, pady=6, sticky="ew")

choix_tri = ttk.Combobox(cadre_form, values=["Plus récent → Plus ancien", "Plus ancien → Plus récent"], state="readonly", style="TCombobox", font=("Roboto", 10))
choix_tri.grid(row=1, column=2, padx=8, pady=6)
choix_tri.current(0)
choix_tri.bind("<<ComboboxSelected>>", trier_afficher_resultats)

cadre_form.columnconfigure(1, weight=1)
Button(fenetre, text="🔍 Lancer la recherche", command=lancer_recherche, width=38, style="TButton", bootstyle="success").pack(pady=(15, 20))

# --- Canvas pour résultats ---
canvas = tk.Canvas(fenetre, height=400, highlightthickness=0)
scrollbar_x = tk.Scrollbar(fenetre, orient="horizontal", command=canvas.xview)
scrollable_frame = Frame(canvas, style="TFrame")
scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(xscrollcommand=scrollbar_x.set)

canvas.pack(fill="both", expand=True, padx=20)
scrollbar_x.pack(fill="x")

cadre_resultats = scrollable_frame
resultats_recherche = []

if fenetre.tk.call('tk', 'windowingsystem') == 'win32':
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
elif fenetre.tk.call('tk', 'windowingsystem') == 'x11':
    canvas.bind_all("<Button-4>", _on_mousewheel_linux)
    canvas.bind_all("<Button-5>", _on_mousewheel_linux)
else:
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
fenetre.mainloop()