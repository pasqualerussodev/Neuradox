import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Funzione per selezionare il file Excel
def seleziona_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xls *.xlsx")],
        title="Seleziona il file Excel"
    )
    if file_path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)

# Funzione per selezionare la cartella dei file da elaborare
def seleziona_cartella():
    cartella = filedialog.askdirectory(title="Seleziona la cartella con i file da elaborare")
    if cartella:
        entry_cartella.delete(0, tk.END)
        entry_cartella.insert(0, cartella)

# Funzione per selezionare la cartella dei file elaborati
def seleziona_cartella_elaborati():
    cartella = filedialog.askdirectory(title="Seleziona la cartella dei file elaborati")
    if cartella:
        entry_cartella_elaborati.delete(0, tk.END)
        entry_cartella_elaborati.insert(0, cartella)

# Funzione per creare e scaricare il file ZIP
def scarica_zip():
    cartella_file_elaborati = filedialog.askdirectory(title="Seleziona la cartella dei file elaborati")
    if not cartella_file_elaborati or not os.path.exists(cartella_file_elaborati) or not os.listdir(cartella_file_elaborati):
        messagebox.showerror("Errore", "La cartella dei file elaborati è vuota o non esiste!")
        return

    zip_filename = filedialog.asksaveasfilename(
        defaultextension=".zip",
        filetypes=[("Archivio ZIP", "*.zip")],
        title="Salva il file ZIP"
    )
    if zip_filename:
        shutil.make_archive(zip_filename.replace(".zip", ""), 'zip', cartella_file_elaborati)
        messagebox.showinfo("Download", f"Archivio '{zip_filename}' creato con successo!")

def mostra_guida():
    guida_testo = (
        "🔹 NeuraDox 1.0 - Guida Rapida 🔹\n\n"
        "📌 Procedura d'uso:\n"
        "1. Seleziona il file Excel contenente i nomi dei file da elaborare.\n"
        "2. Clicca su 'Sfoglia' per scegliere la cartella con i file da elaborare.\n"
        "3. Seleziona la cartella di destinazione per i file rinominati.\n"
        "4. Clicca su 'Avvia elaborazione'.\n"
        "5. Monitora la barra di avanzamento.\n"
        "6. In caso di file mancanti, verifica manualmente nome e posizione.\n"
        "7. Al completamento, utilizza 'Scarica ZIP' per creare un archivio dei file elaborati.\n\n"
        "📂 Gestione del file Excel:\n"
        "- Il file Excel deve contenere almeno 5 colonne, incluse quelle per il nome di origine del file "
        "e per il nuovo nome del file.\n"
        "- In caso di incongruenze, il sistema notificherà l'errore.\n\n"
        "📌 Licenza:\n"
        "Software rilasciato sotto la GNU GPLv3. È possibile utilizzarlo, modificarlo e ridistribuirlo, "
        "a condizione che il codice sorgente sia sempre disponibile.\n\n"
        "⚠️ Esclusione di responsabilità:\n"
        "Il software viene fornito 'così com'è', senza garanzie. L'autore non si assume responsabilità per eventuali danni.\n\n"
        "📜 Testo ufficiale della licenza:\n"
        "Visita: https://www.gnu.org/licenses/gpl-3.0.it.html\n\n"
        "👨‍💻 Sviluppato da: Pasquale Russo\n"
        "📧 Contatti: pasquale.russo@azinfocollection.it"
    )
    messagebox.showinfo("Guida - NeuraDox 1.0", guida_testo)

# Funzione per copiare e rinominare i documenti
def copia_documenti():
    file_excel = entry_file.get()
    if not os.path.exists(file_excel):
        messagebox.showerror("Errore", "Seleziona un file Excel valido!")
        return

    cartella_destinazione = entry_cartella_elaborati.get()
    if not os.path.exists(cartella_destinazione):
        messagebox.showerror("Errore", "Seleziona una cartella di destinazione valida!")
        return

    try:
        df = pd.read_excel(file_excel, dtype=str)
        cartella_file_da_elaborare = entry_cartella.get()

        if df.shape[1] < 5:
            messagebox.showerror("Errore", "Il file Excel non ha abbastanza colonne!")
            return

        # Si assume che la seconda colonna contenga il nome originale del file
        # e la quinta il nuovo nome
        nome_colonna_origine = df.columns[1]
        nome_colonna_nuovo_nome = df.columns[4]

        total_files = len(df)
        progress_bar["maximum"] = total_files
        progress_bar["value"] = 0
        status_label.config(text="Inizio elaborazione...")
        root.update_idletasks()

        for idx, row in df.iterrows():
            nome_file_origine = f"{row[nome_colonna_origine]}.pdf"
            nuovo_nome = f"{row[nome_colonna_nuovo_nome]}.pdf"
            origine = os.path.join(cartella_file_da_elaborare, nome_file_origine)
            destinazione = os.path.join(cartella_destinazione, nuovo_nome)

            if os.path.exists(origine):
                shutil.copy(origine, destinazione)
                progress_bar["value"] += 1
                perc = (progress_bar["value"] / total_files) * 100
                # Visualizza solo il nuovo nome del file e la percentuale
                status_label.config(text=f"{nuovo_nome} ({perc:.1f}%)")
            else:
                # Nel caso il file non venga trovato si mostra un messaggio di errore
                status_label.config(text=f"File non trovato: {nome_file_origine}")
                print(f"⚠️ File non trovato: {origine}")

            root.update_idletasks()

        messagebox.showinfo("Operazione completata", "L'elaborazione è terminata con successo!")
        status_label.config(text="Elaborazione completata.")

    except Exception as e:
        messagebox.showerror("Errore", f"Si è verificato un problema: {e}")
        status_label.config(text="Errore durante l'elaborazione.")

# Creazione della finestra principale
root = tk.Tk()
root.title("NeuraDox 1.0")
root.geometry("400x600")
root.resizable(False, False)
root.configure(bg="#f0f0f0")

# Configurazione dello stile per ttk
style = ttk.Style(root)
style.configure("TButton", font=("Helvetica", 10), padding=5)
style.configure("TLabel", font=("Helvetica", 10), background="#f0f0f0")
style.configure("TEntry", font=("Helvetica", 10))
style.configure("TFrame", background="#f0f0f0")

# Frame principale
main_frame = ttk.Frame(root, padding="10")
main_frame.pack(fill="both", expand=True)

# --- Sezione File Excel ---
frame_excel = ttk.LabelFrame(main_frame, text="File Excel", padding="10")
frame_excel.pack(fill="x", pady=5)

ttk.Label(frame_excel, text="Seleziona il file Excel:").grid(row=0, column=0, sticky="w")
entry_file = ttk.Entry(frame_excel, width=40)
entry_file.grid(row=1, column=0, padx=5, pady=5)
ttk.Button(frame_excel, text="Sfoglia...", command=seleziona_file).grid(row=1, column=1, padx=5, pady=5)

# --- Sezione File da elaborare ---
frame_elaborare = ttk.LabelFrame(main_frame, text="File da elaborare", padding="10")
frame_elaborare.pack(fill="x", pady=5)

ttk.Label(frame_elaborare, text="Seleziona la cartella dei file:").grid(row=0, column=0, sticky="w")
entry_cartella = ttk.Entry(frame_elaborare, width=40)
entry_cartella.grid(row=1, column=0, padx=5, pady=5)
ttk.Button(frame_elaborare, text="Sfoglia...", command=seleziona_cartella).grid(row=1, column=1, padx=5, pady=5)

# --- Sezione File elaborati ---
frame_destinazione = ttk.LabelFrame(main_frame, text="File elaborati", padding="10")
frame_destinazione.pack(fill="x", pady=5)

ttk.Label(frame_destinazione, text="Seleziona la cartella di destinazione:").grid(row=0, column=0, sticky="w")
entry_cartella_elaborati = ttk.Entry(frame_destinazione, width=40)
entry_cartella_elaborati.grid(row=1, column=0, padx=5, pady=5)
ttk.Button(frame_destinazione, text="Sfoglia...", command=seleziona_cartella_elaborati).grid(row=1, column=1, padx=5, pady=5)

# --- Sezione Operazioni ---
frame_operations = ttk.Frame(main_frame, padding="10")
frame_operations.pack(fill="x", pady=5)

ttk.Button(frame_operations, text="Avvia elaborazione", command=copia_documenti).pack(fill="x", padx=5, pady=5)
ttk.Button(frame_operations, text="Scarica ZIP", command=scarica_zip).pack(fill="x", padx=5, pady=5)
ttk.Button(frame_operations, text="Guida", command=mostra_guida).pack(fill="x", padx=5, pady=5)

# --- Barra di Avanzamento e Stato ---
progress_bar = ttk.Progressbar(main_frame, orient="horizontal", length=350, mode="determinate")
progress_bar.pack(pady=10)

status_label = ttk.Label(main_frame, text="Stato: in attesa di avvio...", anchor="center")
status_label.pack(fill="x", padx=5, pady=5)

root.mainloop()
