#!/usr/bin/env python

import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
from matplotlib import cm

# === CONFIGURAZIONE ===
file_path = "Analisi Sociometrica novembre.xlsx"   # percorso del file Excel
sheet_name = "Reciproci_positivi"         # nome esatto del foglio

# === LETTURA DATI ===
df = pd.read_excel(file_path, sheet_name=sheet_name, index_col=0)

# Pulizia dei nomi
df.columns = df.columns.astype(str).str.strip()
df.index = df.index.astype(str).str.strip()

# Rimozione di eventuali totali o celle spurie
df = df.loc[~df.index.str.contains("Totale", case=False, na=False)]
df = df.loc[:, ~df.columns.str.contains("Totale", case=False, na=False)]

# Forza valori numerici; celle non numeriche diventano NaN
df = df.apply(pd.to_numeric, errors="coerce")

# Collassa eventuali duplicati di nomi (es. righe/colonne ripetute)
if df.index.has_duplicates:
    df = df.groupby(level=0).max()
if df.columns.has_duplicates:
    df = df.T.groupby(level=0).max().T

# Mantiene solo righe/colonne comuni
common_names = df.index.intersection(df.columns)
df = df.loc[common_names, common_names]

# === COSTRUZIONE DEL GRAFO ===
G = nx.Graph()
for i in common_names:
    for j in common_names:
        if i != j and df.at[i, j] == 1 and df.at[j, i] == 1:
            G.add_edge(i, j)

# === ESTRAZIONE DELLE CLIQUE ===
cliques = list(nx.find_cliques(G))
print("Clique trovate:")
for c in cliques:
    print(" -", ", ".join(c))

# === CENTRALITÀ (GRADO) ===
degree_dict = dict(G.degree())
max_deg = max(degree_dict.values()) if degree_dict else 1

# Dimensione e colore dei nodi scalati
node_sizes = [400 + (degree_dict[n] / max_deg) * 1200 for n in G.nodes()]
node_colors = [degree_dict[n] for n in G.nodes()]

# === VISUALIZZAZIONE ===
plt.figure(figsize=(13, 9))
pos = nx.spring_layout(G, seed=42, k=0.9, iterations=300)

# Archi neri e molto trasparenti
nx.draw_networkx_edges(G, pos, edge_color="black", alpha=0.2, width=1)

# Nodi con contorno e scala di colore rosso→verde
nodes = nx.draw_networkx_nodes(
    G, pos,
    node_size=node_sizes,
    node_color=node_colors,
    cmap=cm.RdYlGn,
    alpha=0.95,
    linewidths=1.2,
    edgecolors="black"
)
nx.draw_networkx_labels(G, pos, font_size=9, font_weight="bold")

# Barra dei colori
cbar = plt.colorbar(nodes)
cbar.set_label("Numero di legami reciproci positivi", rotation=270, labelpad=25)

plt.title("Rete sociometrica – Clique di reciprocità positiva", fontsize=15)
plt.axis("off")
plt.tight_layout()

# Salva la figura su file quando il backend non è interattivo
output_path = "reciproci_pos.png"
plt.savefig(output_path, dpi=300, bbox_inches="tight")
print(f"Grafico salvato in: {output_path}")
