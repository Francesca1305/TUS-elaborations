import pandas as pd
import matplotlib.pyplot as plt

# Definizione del colore per ciascuna macro-regione
color_map = {
    1: 'blue',
    2: 'cyan',
    3: 'gray',
    4: 'yellow',
    5: 'green'
}

# Caricamento del file Excel
file_path = "Database_grafico_activities.xlsx"  # Sostituisci con il percorso del tuo file
df = pd.read_excel(file_path, sheet_name="Sheet1")

# Creazione della figura con 3 sottografici (uno sotto l'altro)
fig, axes = plt.subplots(3, 1, figsize=(10, 18), sharex=True)

for i, day in enumerate(range(1, 4)):  # Loop attraverso i 3 giorni (1, 2, 3)
    # Filtrare i dati dove "activity_class" è uguale a 1 e "day_type" corrisponde
    filtered_df = df[(df['activity_class'] == 1) & (df['day_type'] == day)]

    # Creazione dello scatter plot per ogni giorno
    for rip_value, color in color_map.items():
        subset = filtered_df[filtered_df['rip'] == rip_value]
        axes[i].scatter(
            subset['meseri_y'],
            subset['R2_prob'],
            label=f"macro-region {rip_value}",
            color=color,
            alpha=0.7,
            edgecolor='black',  # Contorno dei punti
            s=100  # Dimensione dei punti per renderli più visibili
        )

    # Configurazione di ciascun sottografico
    axes[i].set_ylabel("Correlation (RSQ)", fontsize=12)
    axes[i].set_ylim(0, 1)
    axes[i].grid(alpha=0.3)
    axes[i].spines['right'].set_visible(True)
    axes[i].spines['top'].set_visible(True)
    axes[i].legend(title="Macro-regions", fontsize=10)
    axes[i].set_title(f"Scatter plot for day type {day}", fontsize=14)

# Configurazione dell'asse X comune
axes[2].set_xlabel("Quarters", fontsize=12)
axes[2].set_xticks([1, 2, 3, 4])
axes[2].set_xticklabels(["Q1", "Q2", "Q3", "Q4"])  # Etichette per i trimestri

# Ottimizzazione del layout
plt.tight_layout()

# Salvataggio e visualizzazione del grafico
plt.savefig("Scatter_plots_three_days.png")
plt.show()
