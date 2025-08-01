__version__ = "1.0"
__appname__ = "ScrapCommando"



import tkinter as tk
from tkinter import messagebox, ttk
import pyperclip, time, threading
from bs4 import BeautifulSoup

# Variables globales
derniere_classe_loop = None
indices_utiles = []
headers = []
tree = None

def attendre_html(timeout=600):
    pyperclip.copy('')
    debut = time.time()
    while time.time() - debut < timeout:
        contenu = pyperclip.paste()
        if '<html' in contenu.lower():
            return contenu
        time.sleep(0.5)
    messagebox.showerror("Timeout", "HTML non détecté.")
    return None

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

def lancer_detection():
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
    affichage.insert(tk.END, "⏳ En attente du HTML depuis le presse-papier...\n")
    html = attendre_html()
    if not html:
        btn_detect.config(state=tk.NORMAL)
        btn_scrapp.config(state=tk.NORMAL)
        btn_extract.config(state=tk.NORMAL)
        return

    affichage.insert(tk.END, "✅ HTML détecté.\n")
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

    for row in tableau:
        if val1 in row and val2 in row and val3 in row:
            headers, indices_utiles = [], []
            for i, val in enumerate(row):
                if val1 in val:
                    headers.append("data 1")
                    indices_utiles.append(i)
                elif val2 in val:
                    headers.append("data 2")
                    indices_utiles.append(i)
                elif val3 in val:
                    headers.append("data 3")
                    indices_utiles.append(i)
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

def lancer_scrapp_depuis_loop():
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
    affichage.insert(tk.END, "⏳ En attente du HTML depuis le presse-papier...\n")

    def thread_scrap():
        html = attendre_html()
        if not html:
            root.after(0, lambda: [btn_detect.config(state=tk.NORMAL), btn_scrapp.config(state=tk.NORMAL), btn_extract.config(state=tk.NORMAL)])
            return

        soup = BeautifulSoup(html, "html.parser")

        def maj_gui():
            affichage.insert(tk.END, f"""📥 SCRAPP direct avec loop existante : {' '.join(derniere_classe_loop)}
✅ HTML détecté.
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

    threading.Thread(target=thread_scrap).start()

# Interface
root = tk.Tk()
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
                       command=lambda: threading.Thread(target=lancer_detection).start(),
                       height=BTN_HEIGHT, width=BTN_WIDTH)
btn_detect.pack(pady=10, padx=5)

btn_scrapp = tk.Button(right_frame, text="Fire 🔥", font=emoji_font,
                       command=lancer_scrapp_depuis_loop,
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
