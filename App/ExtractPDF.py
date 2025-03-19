import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd

def browse_and_list_pdfs():
    source_dir = filedialog.askdirectory(title="Sélectionnez le dossier source des PDF")

    if not source_dir:
        messagebox.showwarning("Aucun dossier sélectionné", "Veuillez sélectionner un dossier source.")
        root.destroy()
        return

    dest_dir = filedialog.askdirectory(title="Sélectionnez le dossier de destination pour le fichier Excel")

    if not dest_dir:
        messagebox.showwarning("Aucun dossier sélectionné", "Veuillez sélectionner un dossier de destination.")
        root.destroy()
        return

    pdf_files = []

    for current_root, dirs, files in os.walk(source_dir):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(file)

    if pdf_files:
        df = pd.DataFrame(pdf_files, columns=["Noms des fichiers PDF"])
        output_path = os.path.join(dest_dir, "Noms_fichiers.xlsx")
        df.to_excel(output_path, index=False)

        messagebox.showinfo("Succès", f"Fichier Excel créé :\n{output_path}")
    else:
        messagebox.showinfo("Aucun PDF trouvé", "Aucun fichier PDF trouvé dans ce dossier ou ses sous-dossiers.")

    root.destroy()

# GUI Tkinter
root = tk.Tk()
root.withdraw()  # Cache la fenêtre principale immédiatement
browse_and_list_pdfs()