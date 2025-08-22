# -*- coding: utf-8 -*-
"""
ScrapCommando V1.0 (patch auto) — Pilotage de Chrome pour déclencher WebCommando
- Conserve votre UX (Target + Fire / Fire / Extract)
- Quand vous cliquez Target+Fire ou Fire, l'app :
    1) bascule la fenêtre active vers Chrome,
    2) envoie le raccourci Ctrl+Maj+U (commande de l'extension),
    3) attend l'HTML dans le presse-papier,
    4) rebascule la fenêtre vers ScrapCommando,
    5) poursuit la détection/lecture de la LOOP.
⚠️ Windows uniquement (utilise pywin32).

Dépendances nouvelles: pip install pywin32
"""

__version__ = "1.0-auto"
__appname__ = "ScrapCommando"

import tkinter as tk
from tkinter import messagebox, ttk
import pyperclip, time, threading, re, platform
from bs4 import BeautifulSoup
import os, sys, ctypes


IS_WINDOWS = (platform.system() == "Windows")
if IS_WINDOWS:
    import win32gui, win32con, win32api, win32process

# Variables globales (reprennent V1.0)
derniere_classe_loop = None
indices_utiles = []
headers = []
tree = None


def resource_path(rel_path: str) -> str:
    # Permet de retrouver les ressources avec PyInstaller --onefile
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, rel_path)
    return os.path.join(os.path.abspath("."), rel_path)


# ------------------------ UTILITAIRES WINDOWS ------------------------
def _activate_window(hwnd):
    """Force une fenêtre au premier plan (Windows)."""
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        fg = win32gui.GetForegroundWindow()
        if fg == hwnd:
            return True
        # Attache les threads pour autoriser SetForegroundWindow
        curr_tid = win32api.GetCurrentThreadId()
        target_tid, _ = win32process.GetWindowThreadProcessId(hwnd)
        fg_tid, _ = win32process.GetWindowThreadProcessId(fg) if fg else (0,0)
        win32process.AttachThreadInput(curr_tid, target_tid, True)
        if fg:
            win32process.AttachThreadInput(curr_tid, fg_tid, True)
        try:
            win32gui.SetForegroundWindow(hwnd)
            win32gui.BringWindowToTop(hwnd)
            return True
        finally:
            if fg:
                win32process.AttachThreadInput(curr_tid, fg_tid, False)
            win32process.AttachThreadInput(curr_tid, target_tid, False)
    except Exception:
        return False

def _find_chrome_window():
    """Retourne le handle d'une fenêtre Chrome (meilleur effort)."""
    targets = []
    def enum_handler(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        cls = win32gui.GetClassName(hwnd)
        title = win32gui.GetWindowText(hwnd) or ""
        if cls == "Chrome_WidgetWin_1" and ("Chrome" in title or title):
            targets.append(hwnd)
    win32gui.EnumWindows(enum_handler, None)
    # Retourne la plus récente (haut de Z-order)
    return targets[0] if targets else None

def _send_hotkey_ctrl_shift_u():
    """Envoie Ctrl+Shift+U via WinAPI (keybd_event)."""
    VK_CONTROL = 0x11
    VK_SHIFT = 0x10
    VK_U = 0x55
    KEYEVENTF_KEYUP = 0x0002

    # Down: Ctrl, Shift, U
    win32api.keybd_event(VK_CONTROL, 0, 0, 0)
    win32api.keybd_event(VK_SHIFT, 0, 0, 0)
    win32api.keybd_event(VK_U, 0, 0, 0)
    time.sleep(0.035)
    # Up: U, Shift, Ctrl
    win32api.keybd_event(VK_U, 0, KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)

def _trigger_webcommando_then_get_clipboard(timeout=8.0):
    """
    Bascule sur Chrome, déclenche Ctrl+Maj+U, attend l'HTML dans le presse-papier, revient à l'app.
    Retourne le HTML (str) ou None.
    """
    if not IS_WINDOWS:
        messagebox.showerror("OS non supporté", "L'automatisation 100% nécessite Windows (pywin32).")
        return None

    # 1) Trouver et activer Chrome
    chrome_hwnd = _find_chrome_window()
    if not chrome_hwnd:
        messagebox.showerror("Chrome introuvable", "Impossible de localiser une fenêtre Chrome. Ouvrez l'onglet cible et réessayez.")
        return None

    # Mémoriser l'app pour revenir dessus
    app_hwnd = int(root.winfo_id())

    # Nettoyer le presse-papier pour détecter un nouveau contenu
    pyperclip.copy('')

    if not _activate_window(chrome_hwnd):
        messagebox.showerror("Focus", "Échec de l'activation de la fenêtre Chrome.")
        return None
    time.sleep(0.08)

    # 2) Envoyer le raccourci
    _send_hotkey_ctrl_shift_u()

    # 3) Attendre l'HTML dans le presse-papier
    start = time.time()
    html = None
    while time.time() - start < timeout:
        txt = pyperclip.paste() or ""
        if "<html" in txt.lower():
            html = txt
            break
        time.sleep(0.08)

    # 4) Revenir sur ScrapCommando
    _activate_window(app_hwnd)
    time.sleep(0.05)

    return html

# ------------------------ LOGIQUE V1.0 (avec patch auto) ------------------------
def afficher_table(data):
    global tree
    for widget in table_frame.winfo_children():
        widget.destroy()

    headers_upper = [h.upper() for h in headers]

    style = ttk.Style()
    style.configure("Custom.Treeview.Heading", font=('Segoe UI Emoji', 11, 'bold'))
    style.map("Custom.Treeview", background=[('selected', '#ccccff')])

    tree_scroll_y = tk.Scrollbar(table_frame, orient="vertical")
    tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

    tree_scroll_x = tk.Scrollbar(table_frame, orient="horizontal")
    tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

    tree = ttk.Treeview(table_frame, columns=list(range(len(headers_upper))), show='headings',
                        style="Custom.Treeview", yscrollcommand=tree_scroll_y.set,
                        xscrollcommand=tree_scroll_x.set)

    tree_scroll_y.config(command=tree.yview)
    tree_scroll_x.config(command=tree.xview)

    for i, h in enumerate(headers_upper):
        tree.heading(i, text=h)
        max_width = max(len(str(row[i])) if i < len(row) else 0 for row in data)
        tree.column(i, width=max(100, min(300, max_width * 10)))

    for idx, row in enumerate(data):
        tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
        tree.insert('', 'end', values=row, tags=(tag,))

    tree.tag_configure('evenrow', background='#c0c0c0')
    tree.tag_configure('oddrow', background='#ffffff')

    tree.pack(expand=True, fill=tk.BOTH)

def extract_table_to_clipboard():
    if not tree or not tree.get_children():
        return
    tree.selection_set(tree.get_children())
    lignes = []
    for row_id in tree.get_children():
        row = tree.item(row_id)['values']
        lignes.append("\t".join(str(cell) for cell in row))
    texte = "\n".join(lignes)
    root.clipboard_clear()
    root.clipboard_append(texte)
    root.update()

def lancer_detection_auto():
    """Target + Fire (auto) : pilotage Chrome -> copie HTML -> détection loop + aperçu."""
    global derniere_classe_loop, indices_utiles, headers
    btn_detect.config(state=tk.DISABLED)
    btn_scrapp.config(state=tk.DISABLED)
    btn_extract.config(state=tk.DISABLED)
    affichage.delete('1.0', tk.END)

    val1 = entry_val1.get().strip()
    val2 = entry_val2.get().strip()
    val3 = entry_val3.get().strip()
    if not all([val1, val2, val3]):
        messagebox.showwarning("Champs requis", "Veuillez remplir les 3 champs.")
        btn_detect.config(state=tk.NORMAL)
        btn_scrapp.config(state=tk.NORMAL)
        btn_extract.config(state=tk.NORMAL)
        return

    affichage.insert(tk.END, f"""📥 Données saisies :
  - data 1 : {val1}
  - data 2 : {val2}
  - data 3 : {val3}
───────
""")
    affichage.insert(tk.END, "🤖 Pilotage Chrome → copie HTML via WebCommando…\n")
    html = _trigger_webcommando_then_get_clipboard()
    if not html:
        affichage.insert(tk.END, "❌ Échec de récupération HTML depuis Chrome.\n")
        btn_detect.config(state=tk.NORMAL)
        btn_scrapp.config(state=tk.NORMAL)
        btn_extract.config(state=tk.NORMAL)
        return

    affichage.insert(tk.END, "✅ HTML reçu.\n")
    soup = BeautifulSoup(html, "html.parser")
    matching_divs = [
        div for div in soup.find_all("div")
        if val1 in div.get_text(" ", strip=True) and
           val2 in div.get_text(" ", strip=True) and
           val3 in div.get_text(" ", strip=True)
    ]
    if not matching_divs:
        affichage.insert(tk.END, "❌ Aucun bloc <div> contenant les 3 valeurs n'a été trouvé.\n")
        btn_detect.config(state=tk.NORMAL)
        btn_scrapp.config(state=tk.NORMAL)
        btn_extract.config(state=tk.NORMAL)
        return

    bloc = sorted(matching_divs, key=lambda el: len(list(el.parents)), reverse=True)[0]
    derniere_classe_loop = bloc.get("class", [])
    nom_loop = " ".join(derniere_classe_loop) or "(sans classe explicite)"
    affichage.insert(tk.END, f"🧩 LOOP détectée : <div class=\"{nom_loop}\">\n")

    found = soup.find_all("div", class_=lambda x: x and all(c in x for c in derniere_classe_loop))
    affichage.insert(tk.END, f"📊 {len(found)} ligne(s) détectée(s)\n\n")

    tableau = []
    for ligne in found:
        valeurs = [val.strip() for val in ligne.get_text("|", strip=True).split("|") if val.strip()]
        tableau.append(valeurs)

    # Déduction des colonnes (identique V1.0)
    for row in tableau:
        if val1 in row and val2 in row and val3 in row:
            headers[:], indices_utiles[:] = [], []
            for i, val in enumerate(row):
                if val1 in val:
                    headers.append("data 1"); indices_utiles.append(i)
                elif val2 in val:
                    headers.append("data 2"); indices_utiles.append(i)
                elif val3 in val:
                    headers.append("data 3"); indices_utiles.append(i)
            break

    if not indices_utiles:
        affichage.insert(tk.END, "⚠️ Impossible de déterminer les colonnes de référence.\n")
        btn_detect.config(state=tk.NORMAL)
        btn_scrapp.config(state=tk.NORMAL)
        btn_extract.config(state=tk.NORMAL)
        return

    tableau_filtre = [[row[i] for i in indices_utiles] for row in tableau if all(i < len(row) for i in indices_utiles)]
    afficher_table(tableau_filtre)

    btn_detect.config(state=tk.NORMAL)
    btn_scrapp.config(state=tk.NORMAL)
    btn_extract.config(state=tk.NORMAL)

def lancer_scrapp_depuis_loop_auto():
    """Fire (auto) : pilotage Chrome -> copie HTML -> extraction via loop existante."""
    global derniere_classe_loop, indices_utiles, headers
    val1 = entry_val1.get().strip()
    val2 = entry_val2.get().strip()
    val3 = entry_val3.get().strip()
    if not all([val1, val2, val3]) or not derniere_classe_loop or not indices_utiles:
        messagebox.showwarning("Erreur", "Vous devez d'abord exécuter 'Target 🎯 + Fire 🔥'.")
        return

    btn_detect.config(state=tk.DISABLED)
    btn_scrapp.config(state=tk.DISABLED)
    btn_extract.config(state=tk.DISABLED)
    affichage.delete('1.0', tk.END)
    affichage.insert(tk.END, "🤖 Pilotage Chrome → copie HTML via WebCommando…\n")

    def thread_scrap():
        html = _trigger_webcommando_then_get_clipboard()
        if not html:
            root.after(0, lambda: [
                affichage.insert(tk.END, "❌ Échec de récupération HTML depuis Chrome.\n"),
                btn_detect.config(state=tk.NORMAL),
                btn_scrapp.config(state=tk.NORMAL),
                btn_extract.config(state=tk.NORMAL)
            ])
            return

        soup = BeautifulSoup(html, "html.parser")

        def maj_gui():
            affichage.insert(tk.END, f"""📥 SCRAPP direct avec loop existante : {' '.join(derniere_classe_loop)}
✅ HTML reçu.
""")
            found = soup.find_all("div", class_=lambda x: x and all(c in x for c in derniere_classe_loop))
            affichage.insert(tk.END, f"📊 {len(found)} ligne(s) détectée(s) avec LOOP\n\n")

            tableau = []
            for ligne in found:
                valeurs = [val.strip() for val in ligne.get_text("|", strip=True).split("|") if val.strip()]
                tableau.append(valeurs)

            tableau_filtre = [[row[i] for i in indices_utiles] for row in tableau if all(i < len(row) for i in indices_utiles)]
            afficher_table(tableau_filtre)

            btn_detect.config(state=tk.NORMAL)
            btn_scrapp.config(state=tk.NORMAL)
            btn_extract.config(state=tk.NORMAL)

        root.after(0, maj_gui)

    threading.Thread(target=thread_scrap, daemon=True).start()

# ------------------------ INTERFACE ------------------------
root = tk.Tk()

# Avant root.title(...)
app_id = "ASWO.ScrapCommando.V1_0_auto"
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
except Exception:
    pass

def resource_path(rel_path: str) -> str:
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, rel_path)
    return os.path.join(os.path.abspath("."), rel_path)

icon_file = resource_path("scrapcommando.ico")

# 1) Contrôle d'existence
if not os.path.exists(icon_file):
    messagebox.showwarning("Icône manquante",
                           f"Fichier introuvable : {icon_file}\n"
                           f"Vérifiez --add-data et le nom du fichier.")
else:
    try:
        # 2) Icône fenêtre (barre des tâches reprend cette icône)
        root.iconbitmap(icon_file)
    except Exception as e:
        # 3) Fallback : si .ico non accepté par Tk, tenter iconphoto avec PNG intégré
        # (si tu as 'scrapcommando.png' embarqué, sinon ignore ce fallback)
        # from PIL import Image, ImageTk  # si Pillow dispo
        # img = Image.open(resource_path("scrapcommando.png"))
        # root.iconphoto(True, ImageTk.PhotoImage(img))
        messagebox.showwarning("Icône non chargée",
                               f"Échec root.iconbitmap()\n{e}\n"
                               f"Vérifie le format .ico.")


root.title(f"{__appname__} V{__version__}")
root.geometry("1150x750+10+10")

main_frame = tk.Frame(root)
main_frame.pack(fill=tk.BOTH, expand=True)

left_frame = tk.Frame(main_frame)
left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

frame_data = tk.LabelFrame(left_frame, text="HTML", font=("Arial", 11, "bold"), labelanchor='n')
frame_data.pack(padx=10, pady=10, fill=tk.BOTH, expand=False)

input_frame = tk.Frame(frame_data)
input_frame.pack(fill=tk.X, padx=10, pady=10)
tk.Label(input_frame, text="data 1").grid(row=0, column=0, sticky="e", padx=5)
entry_val1 = tk.Entry(input_frame)
entry_val1.grid(row=0, column=1, columnspan=3, sticky="we", pady=2)
tk.Label(input_frame, text="data 2").grid(row=1, column=0, sticky="e", padx=5)
entry_val2 = tk.Entry(input_frame)
entry_val2.grid(row=1, column=1, columnspan=3, sticky="we", pady=2)
tk.Label(input_frame, text="data 3").grid(row=2, column=0, sticky="e", padx=5)
entry_val3 = tk.Entry(input_frame)
entry_val3.grid(row=2, column=1, columnspan=3, sticky="we", pady=2)
input_frame.columnconfigure(1, weight=1)
input_frame.columnconfigure(2, weight=1)
input_frame.columnconfigure(3, weight=1)

scrollbar_affichage_y = tk.Scrollbar(frame_data, orient=tk.VERTICAL)
scrollbar_affichage_y.pack(side=tk.RIGHT, fill=tk.Y)
affichage = tk.Text(frame_data, height=6, wrap=tk.WORD, font=("Arial", 9), yscrollcommand=scrollbar_affichage_y.set)
affichage.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
scrollbar_affichage_y.config(command=affichage.yview)

frame_scrap = tk.LabelFrame(left_frame, text="SCRAPPING", font=("Arial", 11, "bold"), labelanchor='n')
frame_scrap.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0,10))
table_frame = tk.Frame(frame_scrap)
table_frame.pack(fill=tk.BOTH, expand=True)

emoji_font = ("Segoe UI Emoji", 11, "bold")
BTN_WIDTH = 22
BTN_HEIGHT = 2

right_frame = tk.LabelFrame(main_frame, text="", font=("Arial", 11), labelanchor='n', width=240, height=400)
right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10)
right_frame.pack_propagate(False)

btn_detect = tk.Button(right_frame, text="Target 🎯 + Fire 🔥", font=emoji_font,
                       command=lambda: threading.Thread(target=lancer_detection_auto).start(),
                       height=BTN_HEIGHT, width=BTN_WIDTH)
btn_detect.pack(pady=10, padx=5)

btn_scrapp = tk.Button(right_frame, text="Fire 🔥", font=emoji_font,
                       command=lancer_scrapp_depuis_loop_auto,
                       height=BTN_HEIGHT, width=BTN_WIDTH)
btn_scrapp.pack(pady=10, padx=5)

btn_extract = tk.Button(right_frame, text="Extract ✅", font=emoji_font,
                        command=extract_table_to_clipboard,
                        height=BTN_HEIGHT, width=BTN_WIDTH)
btn_extract.pack(pady=10, padx=5)

btn_quit = tk.Button(right_frame, text="Exit ❌", font=emoji_font,
                     command=root.destroy,
                     height=BTN_HEIGHT, width=BTN_WIDTH)
btn_quit.pack(side=tk.BOTTOM, pady=20, padx=5)

root.mainloop()
