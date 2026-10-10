import pandas as pd
import matplotlib.pyplot as plt

# Definizione dei colori e delle etichette per i due tipi di grafico
color_map_1 = {
    1: 'blue',
    2: 'cyan',
    3: 'gray',
    4: 'yellow',
    5: 'green'
}

color_map_2 = {
    2: 'mediumblue',
    3: 'gray',
    111: 'mediumvioletred',
    112: 'orange',
    113: 'darkgoldenrod',
    114: 'gold',
    115: 'salmon',
    116: 'lime',
    117: 'firebrick'
}

activity_labels = {
    2: 'Sleeping',
    3: 'Not at home',
    111: 'TV watching',
    112: 'Food preparation',
    113: 'Laundering',
    114: 'Dish washing',
    115: 'ICT related',
    116: 'Cleaning',
    117: 'Generic active occupancy'
}

# Mappatura day_type
day_type_labels = {
    1: "Weekday",
    2: "Saturday",
    3: "Sunday"
}

rip_class = {
    1: "North-East",
    2: "North-West",
    3: "Centre",
    4: "South",
    5: "Islands"
}

# Caricamento dei file Excel
file_path_1 = "Database_grafico_activities.xlsx"
file_path_2 = "Database_grafico_subactivities.xlsx"

df_1 = pd.read_excel(file_path_1, sheet_name="Sheet1")
df_2 = pd.read_excel(file_path_2, sheet_name="Sheet1")

# Creazione della figura con 3 righe e 2 colonne
fig, axes = plt.subplots(3, 2, figsize=(20, 18), sharey='row', sharex=False)  # Altezza ridotta a 12

for i, day in enumerate(range(1, 4)):  # Loop per day_type (1, 2, 3)
    day_label = day_type_labels[day]  # Etichetta del giorno corrente

    # Primo scatter plot (grafico a sinistra)
    filtered_df_1 = df_1[(df_1['activity_class'] == 1) & (df_1['day_type'] == day)]
    for rip_value, color in color_map_1.items():
        subset = filtered_df_1[filtered_df_1['rip'] == rip_value]
        axes[i, 0].scatter(
            subset['meseri_y'],
            subset['R2_prob'],
            label=f"{rip_class.get(rip_value, f'Macro-region {rip_value}')}",  # usa rip_class
            color=color,
            alpha=0.7,
            edgecolor='black',
            s=50
        )
    axes[i, 0].set_ylabel(f"RSQ between hourly load\nand activity probability\n- {day_label}",
                          fontsize=14)  # Aggiunta di \n per due righe
    axes[i, 0].set_ylim(0, 1)  # Imposta l'asse Y da 0 a 1
    axes[i, 0].set_yticks([0, 0.2, 0.4, 0.6, 0.8, 1])  # Intervalli di 0.2
    axes[i, 0].grid(alpha=0.3)
    axes[i, 0].set_xlabel("Quarters", fontsize=15)
    axes[i, 0].set_xticks([1, 2, 3, 4])
    axes[i, 0].set_xticklabels(["Q1", "Q2", "Q3", "Q4"], fontsize = 14)

    # Secondo scatter plot (grafico a destra)
    filtered_df_2 = df_2[df_2['day_type'] == day]
    for activity_class, color in color_map_2.items():
        subset = filtered_df_2[filtered_df_2['activity_class'] == activity_class]
        axes[i, 1].scatter(
            subset['rip'],
            subset['R2_prob'],
            label=activity_labels.get(activity_class, f"Activity {activity_class}"),
            color=color,
            alpha=0.7,
            edgecolor='black',
            s=50
        )
    axes[i, 1].set_ylim(0, 1)  # Imposta l'asse Y da 0 a 1
    axes[i, 1].set_yticks([0, 0.2, 0.4, 0.6, 0.8, 1])  # Intervalli di 0.2
    axes[i, 1].grid(alpha=0.3)
    axes[i, 1].set_xlabel("Macro-regions", fontsize=15)
    axes[i, 1].set_xticks([1, 2, 3, 4, 5])
    axes[i, 1].set_xticklabels(["North-East", "North-West", "Centre", "South", "Islands"], fontsize = 14)

# Configurazione delle legende
handles_1, labels_1 = axes[0, 0].get_legend_handles_labels()
handles_2, labels_2 = axes[0, 1].get_legend_handles_labels()

fig.legend(handles_1, labels_1, loc='upper center', ncol=3,
           title="Macro-regions (Left plots)", fontsize=13, title_fontsize=14,
           bbox_to_anchor=(0.25, 1.00))  # Posizionamento sopra i grafici

fig.legend(handles_2, labels_2, loc='upper center', ncol=3,
           title="Activity Categories (Right plots)", fontsize=13, title_fontsize=14,
           bbox_to_anchor=(0.75, 1.00))  # Posizionamento sopra i grafici

# Miglioramento del layout
plt.tight_layout(rect=[0, 0, 1, 0.94])  # Spazio extra sopra per le legende

# Salvataggio e visualizzazione del grafico
plt.savefig("Scatter_plots_six_panels.png")
plt.show()

