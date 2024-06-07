import os
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import subprocess
import matplotlib
matplotlib.use('Agg')  # Use the 'Agg' backend (for PNG output)

LOCK_FILE = "app.lock"

'''
def check_already_running():
    if os.path.exists(LOCK_FILE):
        messagebox.showerror("Fehler", "Das Programm wird bereits ausgeführt.")
        return False
    else:
        with open(LOCK_FILE, "w") as lock_file:
            lock_file.write(str(os.getpid()))
        return True
'''

def cleanup():
    if os.path.exists(LOCK_FILE):
        os.remove(LOCK_FILE)

def run_main_script():
    exe_path = "main.exe"  # Angenommen, die ausführbare Datei befindet sich im selben Verzeichnis wie dieses Skript
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW  # Minimiere das Konsolenfenster
    subprocess.Popen([exe_path], startupinfo=startupinfo)
    cleanup()
    root.destroy()  # Schließe das Hauptfenster, bevor das Programm endet

'''
if not check_already_running():
    exit()
'''

root = tk.Tk()
root.title("Kassenbestand Übersicht")
root.geometry("700x720")
root.configure(bg='#333333')

# Globale Variablen
data = []
entry_coin_qty = {}


def save_daily_data():
    try:
        datum = datetime.now().strftime("%Y-%m-%d")
        sum_cash = sum_coin + sum_paper
        start_balance = float(entry_start_balance.get().replace(',', '.'))  # Ersetze das Komma durch einen Punkt und konvertiere in eine Dezimalzahl
        daily_closing = float(entry_daily_closing.get().replace(',', '.'))
        other_income_expenses = entry_other_income_expenses.get().replace(',', '.')
        other_income_expenses = float(other_income_expenses) if other_income_expenses else 0.0
        total_cash_expected = start_balance + daily_closing + other_income_expenses
        deviation = sum_cash - total_cash_expected

        # Verwende die manuelle Eingabe für die Bargeldentnahme
        cash_removal = float(entry_manual_cash_removal.get().replace(',', '.'))  # Manuelle Eingabe für Bargeldentnahme
        end_balance = sum_cash + cash_removal  # Endbestand ist Summe Bargeld plus Entnahme Bargeld

        new_row = {
            "Datum": datum,
            "**Summe Münzen**": sum_coin,
            "**Summe Scheine**": sum_paper,
            "**Anfangsbestand Kassa**": start_balance,
            "**Summe Kassenabschluss**": daily_closing,
            "**Sonstige Einnahmen/Ausgaben**": other_income_expenses,
            "**Summe Bargeld**": sum_cash,
            "**Somme Soll-Bargeld**": total_cash_expected,
            "**Abweichung**": deviation,
            "**Entnahme Bargeld**": cash_removal,
            "**Endbestand Kassa**": end_balance
        }

        df = None
        if os.path.exists('daily_data.xlsx'):
            df = pd.read_excel('daily_data.xlsx')

        # Überprüfen, ob die neue Zeile bereits für den heutigen Tag existiert
        if df is not None and datum in df['Datum'].values:
            confirm = messagebox.askokcancel("Warnung", "Die Daten für heute existieren bereits. Möchten Sie sie überschreiben?")
            if confirm:
                # Zeile für den heutigen Tag aktualisieren
                index_to_update = df.index[df['Datum'] == datum][0]
                df.loc[index_to_update] = new_row
            else:
                return
        else:
            if df is None:
                df = pd.DataFrame(columns=new_row.keys())

            # Hinzufügen der neuen Zeile zu den vorhandenen Daten
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        
        # Sortiere die Daten nach Datum, aufsteigend
        df = df.sort_values(by='Datum', ascending=True)

        # Speichern der Daten in der Excel-Datei
        df.to_excel('daily_data.xlsx', index=False)

        # Speichern der Münzwerte
        coin_data = []
        for value, entry in entry_coin_qty.items():
            try:
                quantity = int(entry.get().replace(',', '.'))  # Ersetze das Komma durch einen Punkt und konvertiere in eine Dezimalzahl
            except ValueError:
                quantity = 0
            coin_data.append((value, quantity))

        coin_df = pd.DataFrame(coin_data, columns=['Value', 'Quantity'])
        coin_df.to_excel('coin_values.xlsx', index=False)

        messagebox.showinfo("Gespeichert", "Tagesstatistik und Münzwerte wurden gespeichert.")
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Speichern der Daten: {e}")


# Funktion zum Laden des Endbestands vom letzten Werktag
def load_last_day_balance():
    try:
        df = pd.read_excel('daily_data.xlsx')
        last_balance = df.iloc[-1]['**Endbestand Kassa**']
        entry_start_balance.delete(0, tk.END)
        entry_start_balance.insert(0, str(last_balance))
    except Exception as e:
        messagebox.showerror("Fehler", f"Endbestand konnte nicht geladen werden: {e}")

def load_last_coin_values(entry_coin_qty):
    try:
        if os.path.exists('coin_values.xlsx'):
            df = pd.read_excel('coin_values.xlsx')
            print("Münzwerte aus der Excel-Tabelle:")
            print(df)  # Debugging-Ausgabe der Münzwerte aus der Excel-Tabelle
            for index, row in df.iterrows():
                value = row['Value']
                quantity = int(row['Quantity'])  # Umwandlung in eine Ganzzahl
                if value in entry_coin_qty:
                    entry_coin_qty[value].delete(0, tk.END)  # Lösche vorhandenen Wert
                    entry_coin_qty[value].insert(0, str(quantity))  # Setze den ganzzahligen Wert ein
        else:
            messagebox.showinfo("Info", "Keine gespeicherten Münzwerte gefunden.")
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Laden der Münzwerte: {e}")


# Funktion zur Berechnung und Anzeige der täglichen Bargeldentnahme

def plot_cash_removal_monthly():
    try:
        df = pd.read_excel("daily_data.xlsx")
    except FileNotFoundError:
        messagebox.showerror("Fehler", "Die Datei 'daily_data.xlsx' wurde nicht gefunden.")
        return

    if "**Entnahme Bargeld**" not in df.columns:
        messagebox.showerror("Fehler", "Die Spalte '**Entnahme Bargeld**' wurde nicht gefunden.")
        return

    df['Datum'] = pd.to_datetime(df['Datum'])

    # Convert 'Entnahme Bargeld' column to numeric
    df['**Entnahme Bargeld**'] = pd.to_numeric(df['**Entnahme Bargeld**'], errors='coerce')  # Coerce errors to NaN

    # Drop rows with NaN values in the 'Entnahme Bargeld' column
    df = df.dropna(subset=['**Entnahme Bargeld**'])

    # Berechne den gleitenden Durchschnitt
    window_size = 5  # Anpassen der Fenstergröße nach Bedarf
    rolling_average = df['**Entnahme Bargeld**'].rolling(window=window_size).mean()

    # Grafische Darstellung der Monatsstatistik für die Entnahme von Bargeld mit einem gleitenden Durchschnitt
    plt.figure(figsize=(10, 5))
    plt.plot(df['Datum'], df['**Entnahme Bargeld**'].abs(), marker='o', linestyle='-', color='red', label='Entnahme Bargeld')
    plt.plot(df['Datum'], rolling_average.abs(), linestyle='--', color='blue', label=f'Gleitender Durchschnitt ({window_size} Tage)')
    plt.title("Tägliche Entnahme von Bargeld mit gleitendem Durchschnitt")
    plt.xlabel("Datum")
    plt.ylabel("Entnahme Bargeld")
    plt.legend()
    plt.grid(True)

    # Datum für den Dateinamen des PNG-Exports
    export_date = datetime.now().strftime("%Y-%m-%d")
    export_filename = f"cash_removal_plot_{export_date}.png"

    plt.savefig(export_filename)  # Save the plot to a file
    messagebox.showinfo("Exportiert", f"Die Plotdatei wurde als '{export_filename}' exportiert.")


def clear_fields():
    # Lösche den Inhalt der Eingabefelder für den Startbestand, den Kassenabschluss und andere Einnahmen/Ausgaben
    entry_start_balance.delete(0, tk.END)
    entry_daily_closing.delete(0, tk.END)
    entry_other_income_expenses.delete(0, tk.END)
    entry_manual_cash_removal.delete(0, tk.END)
    
    # Lösche den Inhalt der Eingabefelder für Münzen
    for entry in entry_coin_qty.values():
        entry.delete(0, tk.END)
    
    # Lösche den Inhalt der Eingabefelder für Scheine
    for entry in entry_paper_qty.values():
        entry.delete(0, tk.END)
    
    # Aktualisiere die Summen-Labels
    label_sum_coin.config(text="Summe Münzen: 0.00 €")
    label_sum_paper.config(text="Summe Scheine: 0.00 €")
    label_total_cash.config(text="Summe Bargeld: 0.00 €")
    label_total_cash_expected.config(text="Summe Soll-Bargeld: 0.00 €")
    label_deviation.config(text="Abweichung: 0.00 €")
    label_cash_removal.config(text="Entnahme Bargeld: 0.00 €", fg='white')
    label_end_balance.config(text="Endbestand Kassa: 0.00 €")


# Funktion zum Exportieren der Monatsstatistik
def export_monthly_statistics():
    try:
        df = pd.read_excel("daily_data.xlsx")
    except FileNotFoundError:
        messagebox.showerror("Fehler", "Die Datei 'daily_data.xlsx' wurde nicht gefunden.")
        return

    current_month = datetime.now().month
    current_year = datetime.now().year

    df['Datum'] = pd.to_datetime(df['Datum'])
    monthly_data = df[(df['Datum'].dt.month == current_month) & (df['Datum'].dt.year == current_year)]

    if monthly_data.empty:
        messagebox.showinfo("Keine Daten", "Für den aktuellen Monat sind keine Daten vorhanden.")
        return
    
    # Exportieren der Daten in eine Excel-Datei
    output_filename = f"Monatsstatistik_{current_year}_{current_month}.xlsx"
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        monthly_data.to_excel(writer, index=False, sheet_name='Daten')

        workbook = writer.book
        worksheet = writer.sheets['Daten']

        # Grafische Darstellung hinzufügen
        chart = workbook.add_chart({'type': 'line'})

        # Konvertiere das Datum in Excel-freundliches Format
        monthly_data['Datum'] = monthly_data['Datum'].dt.strftime('%Y-%m-%d')

        # Daten zum Diagramm hinzufügen
        chart.add_series({
            'name': 'Entnahme Bargeld',
            'categories': ['Daten', 1, 0, len(monthly_data), 0],
            'values': ['Daten', 1, 9, len(monthly_data), 9],  # Spalte 9 = "**Entnahme Bargeld**"
        })

        # Konfiguration des Diagramms
        chart.set_title({'name': 'Monatsstatistik der Entnahme von Bargeld'})
        chart.set_x_axis({'name': 'Datum'})
        chart.set_y_axis({'name': 'Entnahme Bargeld'})
        worksheet.insert_chart('M2', chart)


def export_monthly_statistics():
    try:
        df = pd.read_excel("daily_data.xlsx")
    except FileNotFoundError:
        messagebox.showerror("Fehler", "Die Datei 'daily_data.xlsx' wurde nicht gefunden.")
        return

    current_month = datetime.now().month
    current_year = datetime.now().year

    df['Datum'] = pd.to_datetime(df['Datum'])
    unique_months = df['Datum'].dt.to_period('M').unique()

    if len(unique_months) != 1 or unique_months[0].month != current_month:
        messagebox.showinfo("Export nicht möglich", "Der Export der Monatsstatistik ist nur möglich, wenn alle Datensätze aus dem aktuellen Monat stammen.")
        return

    if len(df) < 12:
        messagebox.showinfo("Export nicht möglich", "Der Export der Monatsstatistik ist erst möglich, wenn mindestens 12 Datensätze vorhanden sind.")
        return
    
    # Exportieren der Daten in eine Excel-Datei
    output_filename = f"Monatsstatistik_{current_year}_{current_month}.xlsx"
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Daten')

        workbook = writer.book
        worksheet = writer.sheets['Daten']

        # Konvertiere das Datum in Excel-freundliches Format
        df['Datum'] = df['Datum'].dt.strftime('%Y-%m-%d')

    messagebox.showinfo("Exportiert", f"Monatsstatistik wurde als '{output_filename}' exportiert.")


def bind_enter_to_next_field(entry_widgets):
    for i in range(len(entry_widgets) - 1):
        entry_widgets[i].bind('<Return>', lambda event, next_entry=entry_widgets[i+1]: next_entry.focus())



# Button zum Leeren der Felder
clear_button = tk.Button(root, text="Felder leeren", command=clear_fields, bg='#86BD3B', fg='white')
clear_button.pack(pady=5)

# Eingabefelder für die Münzen und Scheine
coin_values = [0.10, 0.20, 0.50, 1, 2]
paper_values = [5, 10, 20, 50, 100, 200, 500]

entry_coin_qty = {}
entry_coin_sum = {}
entry_paper_qty = {}
entry_paper_sum = {}

# Münzen Eingabe
coin_frame = tk.Frame(root, bg='#333333')
coin_frame.pack(pady=10)

tk.Label(coin_frame, text="Münzen", bg='#333333', fg='#86BD3B', font=('Arial', 14)).grid(row=0, column=0, columnspan=3)

for i, value in enumerate(coin_values):
    tk.Label(coin_frame, text=f"{value} €", bg='#333333', fg='white').grid(row=i+1, column=0)
    entry_coin_qty[value] = tk.Entry(coin_frame, width=10)
    entry_coin_qty[value].grid(row=i+1, column=1)
    entry_coin_sum[value] = tk.Label(coin_frame, text="0.00 €", bg='#333333', fg='white')
    entry_coin_sum[value].grid(row=i+1, column=2)

# Scheine Eingabe
paper_frame = tk.Frame(root, bg='#333333')
paper_frame.pack(pady=10)

tk.Label(paper_frame, text="Scheine", bg='#333333', fg='#86BD3B', font=('Arial', 14)).grid(row=0, column=0, columnspan=3)

for i, value in enumerate(paper_values):
    tk.Label(paper_frame, text=f"{value} €", bg='#333333', fg='white').grid(row=i+1, column=0)
    entry_paper_qty[value] = tk.Entry(paper_frame, width=10)
    entry_paper_qty[value].grid(row=i+1, column=1)
    entry_paper_sum[value] = tk.Label(paper_frame, text="0.00 €", bg='#333333', fg='white')
    entry_paper_sum[value].grid(row=i+1, column=2)

def calculate_sums():
    global sum_coin, sum_paper
    sum_coin = 0
    sum_paper = 0

    for value, entry in entry_coin_qty.items():
        try:
            quantity = float(entry.get().replace(',', '.'))  # Ersetze das Komma durch einen Punkt und konvertiere in eine Dezimalzahl
        except ValueError:
            quantity = 0
        total = value * quantity
        sum_coin += total
        entry_coin_sum[value].config(text=f"{total:.2f} €")

    for value, entry in entry_paper_qty.items():
        try:
            quantity = float(entry.get().replace(',', '.'))  # Ersetze das Komma durch einen Punkt und konvertiere in eine Dezimalzahl
        except ValueError:
            quantity = 0
        total = value * quantity
        sum_paper += total
        entry_paper_sum[value].config(text=f"{total:.2f} €")

    label_sum_coin.config(text=f"Summe Münzen: {sum_coin:.2f} €")
    label_sum_paper.config(text=f"Summe Scheine: {sum_paper:.2f} €")
    label_total_cash.config(text=f"Summe Bargeld: {sum_coin + sum_paper:.2f} €")

    # Berechnung der weiteren Felder
    try:
        start_balance = float(entry_start_balance.get().replace(',', '.'))  # Ersetze das Komma durch einen Punkt und konvertiere in eine Dezimalzahl
    except ValueError:
        start_balance = 0.0

    try:
        daily_closing = float(entry_daily_closing.get().replace(',', '.'))  # Ersetze das Komma durch einen Punkt und konvertiere in eine Dezimalzahl
    except ValueError:
        daily_closing = 0.0

    try:
        other_income_expenses = float(entry_other_income_expenses.get().replace(',', '.'))  # Ersetze das Komma durch einen Punkt und konvertiere in eine Dezimalzahl
    except ValueError:
        other_income_expenses = 0.0

    try:
        cash_removal = float(entry_manual_cash_removal.get().replace(',', '.'))  # Manuelle Eingabe für Bargeldentnahme
    except ValueError:
        cash_removal = 0.0

    total_cash_expected = start_balance + daily_closing + other_income_expenses
    deviation = (sum_coin + sum_paper) - total_cash_expected

    end_balance = sum_coin + sum_paper + cash_removal  # Adjusted line

    # Aktualisiere die Label-Texte
    label_total_cash_expected.config(text=f"Summe Soll-Bargeld: {total_cash_expected:.2f} €")
    label_deviation.config(text=f"Abweichung: {deviation:.2f} €", fg="green" if deviation >= 0 else "red")  # Ändere die Farbe je nach Wert
    label_cash_removal.config(text=f"Entnahme Bargeld: {cash_removal:.2f} €", fg="red")  # Behalte die rote Farbe bei
    label_end_balance.config(text=f"Endbestand Kassa: {end_balance:.2f} €")




# Labels und Eingabefelder für die restlichen Werte
control_frame = tk.Frame(root, bg='#333333')
control_frame.pack(pady=10)

tk.Label(control_frame, text="Anfangsbestand Kassa", bg='#333333', fg='white').grid(row=0, column=0)
entry_start_balance = tk.Entry(control_frame, width=10)
entry_start_balance.grid(row=0, column=1)

# Hinzufügen der padx-Eigenschaft, um den Button nach links zu verschieben
tk.Button(control_frame, text="Endbestand vom letzten Werktag laden", command=load_last_day_balance, bg='#86BD3B', fg='white').grid(row=0, column=2, padx=(10, 5), sticky='w')

tk.Label(control_frame, text="Summe Kassenabschluss", bg='#333333', fg='white').grid(row=1, column=0)
entry_daily_closing = tk.Entry(control_frame, width=10)
entry_daily_closing.grid(row=1, column=1)

# Hier den Hinweistext für Summe Kassenabschluss hinzufügen
tk.Label(control_frame, text="*Hier Summe aus Radix eingeben", bg='#333333', fg='white', font=("Arial", 8)).grid(row=1, column=2, padx=5, sticky='w')

tk.Label(control_frame, text="Sonstige Einnahmen/Ausgaben", bg='#333333', fg='white').grid(row=2, column=0)
entry_other_income_expenses = tk.Entry(control_frame, width=10)
entry_other_income_expenses.grid(row=2, column=1)

tk.Label(control_frame, text="Manuelle Eingabe Bargeldentnahme", bg='#333333', fg='white').grid(row=3, column=0)
entry_manual_cash_removal = tk.Entry(control_frame, width=10)
entry_manual_cash_removal.grid(row=3, column=1)


# Hier den Hinweistext für Sonstige Einnahmen/Ausgaben hinzufügen
tk.Label(control_frame, text="*Hier Sonstige Einnahmen oder Ausgaben wie z.B. Begleitrechungen eingeben", bg='#333333', fg='white', font=("Arial", 8)).grid(row=2, column=2, padx=5, sticky='w')

# Labels zur Anzeige der berechneten Summen
result_frame = tk.Frame(root, bg='#333333')
result_frame.pack(pady=10)

label_sum_coin = tk.Label(result_frame, text="Summe Münzen: 0.00 €", bg='#333333', fg='white')
label_sum_coin.pack()

label_sum_paper = tk.Label(result_frame, text="Summe Scheine: 0.00 €", bg='#333333', fg='white')
label_sum_paper.pack()

label_total_cash = tk.Label(result_frame, text="Summe Bargeld: 0.00 €", bg='#333333', fg='white')
label_total_cash.pack()

label_total_cash_expected = tk.Label(result_frame, text="Summe Soll-Bargeld: 0.00 €", bg='#333333', fg='white')
label_total_cash_expected.pack()

label_deviation = tk.Label(result_frame, text="Abweichung: 0.00 €", bg='#333333', fg='white')
label_deviation.pack()

label_cash_removal = tk.Label(result_frame, text="Entnahme Bargeld: 0.00 €", bg='#333333', fg='white')
label_cash_removal.pack()

label_end_balance = tk.Label(result_frame, text="Endbestand Kassa: 0.00 €", bg='#333333', fg='white')
label_end_balance.pack()

# Buttons für Aktionen
button_frame = tk.Frame(root, bg='#333333')
button_frame.pack(pady=10)

style = ttk.Style()
style.configure("TButton", font=('Arial', 10), borderwidth=1, focusthickness=3, focuscolor='#86BD3B')
style.map("TButton", foreground=[('active', '#333333')], background=[('active', '#86BD3B')])


tk.Button(button_frame, text="Berechnen", command=calculate_sums, bg='#86BD3B', fg='white').pack(side='left', padx=10)
tk.Button(button_frame, text="Speichern", command=save_daily_data, bg='#86BD3B', fg='white').pack(side='left', padx=10)
tk.Button(button_frame, text="Entnahme Bargeld Statistik", command=plot_cash_removal_monthly, bg='#86BD3B', fg='white').pack(side='left', padx=10)
tk.Button(button_frame, text="Export Monatsstatistik", command=export_monthly_statistics, bg='#86BD3B', fg='white').pack(side='left', padx=10)


load_last_coin_values(entry_coin_qty)



# Liste der Entry-Widgets für Münzen
coin_entry_widgets = list(entry_coin_qty.values())

# Liste der Entry-Widgets für Scheine
paper_entry_widgets = list(entry_paper_qty.values())

# Liste aller Entry-Widgets im GUI
all_entry_widgets = coin_entry_widgets + paper_entry_widgets + [entry_start_balance, entry_daily_closing, entry_other_income_expenses, entry_manual_cash_removal]

# Funktion aufrufen, um Enter-Taste zum Springen zum nächsten Feld zu binden
bind_enter_to_next_field(all_entry_widgets)



def on_closing():
    cleanup()
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

# Start main loop
root.mainloop()



