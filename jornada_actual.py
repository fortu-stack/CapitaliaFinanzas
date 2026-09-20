import json
import math
import numpy as np
import pandas as pd
from fpdf import FPDF

# Compatibilidad de FPDF con versiones v1 y v2
try:
    from fpdf.enums import XPos, YPos
    USE_FPDF2 = True
except ImportError:
    USE_FPDF2 = False

# ---------------------------------------------------------
# 1. CARGA DE DATOS (CON SOPORTE UTF-8-SIG PARA WINDOWS)
# ---------------------------------------------------------
print("📂 Cargando datos desde 'datos_jornada.json'...")
try:
    with open("datos_jornada.json", "r", encoding="utf-8-sig") as f:
        datos_partidos = json.load(f)
except FileNotFoundError:
    raise FileNotFoundError("❌ No se encontró 'datos_jornada.json'. Ejecuta primero el comando de generación de datos.")

partidos_titulos = [f"{p['id']}. {p['local']} - {p['visitante']}" for p in datos_partidos]

# ---------------------------------------------------------
# 2. CÁLCULO DE PROBABILIDADES % (POISSON + AJUSTE ELO)
# ---------------------------------------------------------
def poisson_pmf(k_range, lmbda):
    return np.array([(lmbda**k) * np.exp(-lmbda) / math.factorial(k) for k in k_range])

probs_puras = []
for o in datos_partidos:
    elo_diff = o["elo_loc"] - o["elo_vis"]
    l_loc = max(0.2, o["xg_loc"] * (1 + 0.00025 * elo_diff))
    l_vis = max(0.2, o["xg_vis"] * (1 - 0.00025 * elo_diff))
    
    g_grid = np.arange(10)
    p_mat = np.outer(poisson_pmf(g_grid, l_loc), poisson_pmf(g_grid, l_vis))
    p_mod = np.array([np.sum(np.tril(p_mat, -1)), np.sum(np.diag(p_mat)), np.sum(np.triu(p_mat, 1))])
    p_mod /= p_mod.sum()
    probs_puras.append(p_mod)

probs_puras = np.array(probs_puras) # Matriz (14, 3) -> [%1, %X, %2]

# CÁLCULO ESPECÍFICO DEL PLENO AL 15 (Partido estrella: Sevilla - Barcelona o similar)
# Matriz de probabilidad para resultados exactos (0, 1, 2, M)
pleno_match = {"local": "Sevilla", "visitante": "Barcelona", "xg_loc": 1.10, "xg_vis": 2.10, "elo_loc": 1650, "elo_vis": 1900}
l_p15_loc = max(0.2, pleno_match["xg_loc"] * (1 + 0.00025 * (pleno_match["elo_loc"] - pleno_match["elo_vis"])))
l_p15_vis = max(0.2, pleno_match["xg_vis"] * (1 - 0.00025 * (pleno_match["elo_loc"] - pleno_match["elo_vis"])))
p_mat_15 = np.outer(poisson_pmf(np.arange(10), l_p15_loc), poisson_pmf(np.arange(10), l_p15_vis))

# Opciones con mayor probabilidad lógica para el Pleno al 15
p15_options = ["1-M", "1-2", "0-M", "1-1", "0-2", "M-M"]

# ---------------------------------------------------------
# 3. 1.000.000 DE SIMULACIONES MONTE CARLO VECTORIZADAS
# ---------------------------------------------------------
NUM_SIM = 1_000_000
print(f"🚀 Ejecutando {NUM_SIM:,} simulaciones de Monte Carlo...")

np.random.seed(42)
rand_matrix = np.random.rand(NUM_SIM, 14)
cum_probs = np.cumsum(probs_puras, axis=1)

sim_indices = np.zeros((NUM_SIM, 14), dtype=int)
for m in range(14):
    sim_indices[:, m] = (rand_matrix[:, m] > cum_probs[m, 0]).astype(int) + (rand_matrix[:, m] > cum_probs[m, 1]).astype(int)

# Filtros de rentabilidad (5-8 victorias locales, 3-5 empates, 2-4 visitantes)
c1 = np.sum(sim_indices == 0, axis=1)
cX = np.sum(sim_indices == 1, axis=1)
c2 = np.sum(sim_indices == 2, axis=1)
mask = (c1 >= 5) & (c1 <= 8) & (cX >= 3) & (cX <= 5) & (c2 >= 2) & (c2 <= 4)
valid_indices = np.where(mask)[0]

sign_map = np.array(['1', 'X', '2'])
valid_cols_signs = sign_map[sim_indices[valid_indices]]

prob_mat = probs_puras[np.arange(14), sim_indices[valid_indices]]
probs_boletos = np.prod(prob_mat, axis=1)

sort_idx = np.argsort(-probs_boletos)
valid_cols_signs = valid_cols_signs[sort_idx]
probs_boletos = probs_boletos[sort_idx]

# ---------------------------------------------------------
# 4. REDUCCIÓN A 300 COLUMNAS Y SELECCIÓN DE TOP 3
# ---------------------------------------------------------
def hamming_dist(c1, c2):
    return np.sum(c1 != c2)

selected_cols = [valid_cols_signs[0]]
for i in range(1, len(valid_cols_signs)):
    if len(selected_cols) >= 300:
        break
    if min(hamming_dist(valid_cols_signs[i], s) for s in selected_cols) >= 3:
        selected_cols.append(valid_cols_signs[i])

# Selección de las 3 mejores columnas diversificadas
top3_cols = [selected_cols[0]]
for cand in selected_cols[1:]:
    if len(top3_cols) >= 3:
        break
    if min(hamming_dist(cand, s) for s in top3_cols) >= 4:
        top3_cols.append(cand)

# Asignar resultados recomendados del Pleno al 15
p15_top3 = ["1-M", "1-2", "1-1"]
p15_300 = [p15_options[i % len(p15_options)] for i in range(len(selected_cols))]

# ---------------------------------------------------------
# 5. GENERACIÓN DEL EXCEL CON PLENO AL 15 Y PORCENTAJES
# ---------------------------------------------------------
print("📊 Generando '300_Columnas_Millon_Sim.xlsx'...")

# Construcción de la matriz principal
excel_data = {"#": list(range(1, 15)) + [15], "Partido": partidos_titulos + ["15. Sevilla - Barcelona (Pleno 15)"]}

# Añadir columnas de porcentajes calculados
excel_data["% 1"] = [f"{probs_puras[i,0]*100:.1f}%" for i in range(14)] + ["-"]
excel_data["% X"] = [f"{probs_puras[i,1]*100:.1f}%" for i in range(14)] + ["-"]
excel_data["% 2"] = [f"{probs_puras[i,2]*100:.1f}%" for i in range(14)] + ["-"]

# Añadir las 300 columnas con el Pleno al 15 en la Fila 15
for col_idx in range(len(selected_cols)):
    tag = f"Columna #{col_idx+1}"
    if col_idx < 3:
        tag += " ⭐ TOP"
    
    col_full = list(selected_cols[col_idx]) + [p15_300[col_idx]]
    excel_data[tag] = col_full

df_excel = pd.DataFrame(excel_data)
df_excel.to_excel("300_Columnas_Millon_Sim.xlsx", index=False)

# ---------------------------------------------------------
# 6. GENERACIÓN DE ARCHIVOS TXT (16 CARACTERES)
# ---------------------------------------------------------
print("📄 Generando archivos TXT para validación oficial...")

# 3 Columnas a jugar (14 signos + 2 caracteres del Pleno al 15)
with open("3_Columnas_A_Jugar.txt", "w", encoding="utf-8") as f:
    for idx, col in enumerate(top3_cols):
        p15_clean = p15_top3[idx].replace("-", "")
        f.write("".join(col) + p15_clean + "\n")

# 300 Columnas
with open("300_Columnas_Millon_Sim.txt", "w", encoding="utf-8") as f:
    for idx, col in enumerate(selected_cols):
        p15_clean = p15_300[idx].replace("-", "")
        f.write("".join(col) + p15_clean + "\n")

# ---------------------------------------------------------
# 7. GENERACIÓN DEL PDF CON PORCENTAJES Y TOP 3
# ---------------------------------------------------------
print("📕 Generando 'Informe_Millon_Simulaciones.pdf'...")

class PDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 13)
        title = "LA QUINIELA - MATRIZ Y PORCENTAJES (1.000.000 SIMULACIONES)"
        if USE_FPDF2:
            self.cell(0, 8, title, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
        else:
            self.cell(0, 8, title, 0, 1, "C")
        self.ln(3)

pdf = PDF()
pdf.add_page()

# Sección 1: Porcentajes calculados por partido
pdf.set_font("Helvetica", "B", 10)
hdr_title = "1. PORCENTAJES Y PROBABILIDADES CALCULADAS DE LA JORNADA"
if USE_FPDF2:
    pdf.cell(0, 6, hdr_title, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
else:
    pdf.cell(0, 6, hdr_title, 0, 1)

pdf.set_font("Helvetica", "B", 8)
pdf.set_fill_color(220, 230, 242)

if USE_FPDF2:
    pdf.cell(8, 5, "#", 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(100, 5, "Partido", 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="L", fill=True)
    pdf.cell(22, 5, "% Local (1)", 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(22, 5, "% Empate (X)", 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(22, 5, "% Visitante (2)", 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", fill=True)
else:
    pdf.cell(8, 5, "#", 1, 0, "C", True)
    pdf.cell(100, 5, "Partido", 1, 0, "L", True)
    pdf.cell(22, 5, "% Local (1)", 1, 0, "C", True)
    pdf.cell(22, 5, "% Empate (X)", 1, 0, "C", True)
    pdf.cell(22, 5, "% Visitante (2)", 1, 1, "C", True)

pdf.set_font("Helvetica", "", 8)
for i, name in enumerate(partidos_titulos):
    p1 = f"{probs_puras[i,0]*100:.1f}%"
    px = f"{probs_puras[i,1]*100:.1f}%"
    p2 = f"{probs_puras[i,2]*100:.1f}%"
    if USE_FPDF2:
        pdf.cell(8, 4.2, str(i+1), 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C")
        pdf.cell(100, 4.2, name, 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="L")
        pdf.cell(22, 4.2, p1, 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C")
        pdf.cell(22, 4.2, px, 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C")
        pdf.cell(22, 4.2, p2, 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    else:
        pdf.cell(8, 4.2, str(i+1), 1, 0, "C")
        pdf.cell(100, 4.2, name, 1, 0, "L")
        pdf.cell(22, 4.2, p1, 1, 0, "C")
        pdf.cell(22, 4.2, px, 1, 0, "C")
        pdf.cell(22, 4.2, p2, 1, 1, "C")

pdf.ln(5)

# Sección 2: Las 3 columnas seleccionadas para jugar
pdf.set_font("Helvetica", "B", 10)
top_title = "2. LAS 3 COLUMNAS OPTIMIZADAS PARA JUGAR (COSTE: 2,25 EUR)"
if USE_FPDF2:
    pdf.cell(0, 6, top_title, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
else:
    pdf.cell(0, 6, top_title, 0, 1)

pdf.set_font("Helvetica", "B", 8)
pdf.set_fill_color(200, 220, 240)

if USE_FPDF2:
    pdf.cell(8, 5, "#", 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(100, 5, "Partido", 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="L", fill=True)
    pdf.cell(22, 5, "Apuesta 1", 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(22, 5, "Apuesta 2", 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(22, 5, "Apuesta 3", 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", fill=True)
else:
    pdf.cell(8, 5, "#", 1, 0, "C", True)
    pdf.cell(100, 5, "Partido", 1, 0, "L", True)
    pdf.cell(22, 5, "Apuesta 1", 1, 0, "C", True)
    pdf.cell(22, 5, "Apuesta 2", 1, 0, "C", True)
    pdf.cell(22, 5, "Apuesta 3", 1, 1, "C", True)

pdf.set_font("Helvetica", "", 8)
for i, name in enumerate(partidos_titulos):
    s1, s2, s3 = top3_cols[0][i], top3_cols[1][i], top3_cols[2][i]
    if USE_FPDF2:
        pdf.cell(8, 4.2, str(i+1), 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C")
        pdf.cell(100, 4.2, name, 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="L")
        pdf.cell(22, 4.2, s1, 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C")
        pdf.cell(22, 4.2, s2, 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C")
        pdf.cell(22, 4.2, s3, 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    else:
        pdf.cell(8, 4.2, str(i+1), 1, 0, "C")
        pdf.cell(100, 4.2, name, 1, 0, "L")
        pdf.cell(22, 4.2, s1, 1, 0, "C")
        pdf.cell(22, 4.2, s2, 1, 0, "C")
        pdf.cell(22, 4.2, s3, 1, 1, "C")

# Fila para el Pleno al 15
pdf.set_font("Helvetica", "B", 8)
if USE_FPDF2:
    pdf.cell(8, 5, "15", 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(100, 5, "15. Sevilla - Barcelona (Pleno al 15)", 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="L", fill=True)
    pdf.cell(22, 5, p15_top3[0], 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(22, 5, p15_top3[1], 1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(22, 5, p15_top3[2], 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", fill=True)
else:
    pdf.cell(8, 5, "15", 1, 0, "C", True)
    pdf.cell(100, 5, "15. Sevilla - Barcelona (Pleno al 15)", 1, 0, "L", True)
    pdf.cell(22, 5, p15_top3[0], 1, 0, "C", True)
    pdf.cell(22, 5, p15_top3[1], 1, 0, "C", True)
    pdf.cell(22, 5, p15_top3[2], 1, 1, "C", True)

pdf.output("Informe_Millon_Simulaciones.pdf")

print("\n🎉 ¡PROCESO COMPLETADO CON ÉXITO!")
print("Archivos actualizados en la carpeta:")
print(" 📄 3_Columnas_A_Jugar.txt       -> Con 16 caracteres (signos + Pleno al 15 listo para apostar)")
print(" 📊 300_Columnas_Millon_Sim.xlsx -> Excel con Fila 15 de Pleno al 15 + Porcentajes %1 %X %2")
print(" 📕 Informe_Millon_Simulaciones.pdf -> PDF con la tabla de % de probabilidad y las 3 apuestas")