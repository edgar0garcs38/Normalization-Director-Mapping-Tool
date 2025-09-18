import pandas as pd
from rapidfuzz import process, fuzz
import unicodedata
import re

def normalizar_alias(alias):
    if pd.isna(alias):
        return ""
    # Quitar tildes, pasar a mayúsculas, limpiar espacios
    alias = str(alias).upper()
    alias = unicodedata.normalize("NFKD", alias)
    alias = alias.encode("ASCII", "ignore").decode("utf-8")  # remove accents
    alias = re.sub(r"\s+", " ", alias)  # reemplaza múltiples espacios por uno
    alias = alias.strip()
    return alias

# ---------- Archivos y hoja ----------
archivo_jerarquia = "REPORTE DE JERARQUIAS AGENTES A 1 SEP.xlsx"
archivo_renovaciones = "BASE DE RENOVACIONES 2022 A 2024 PARA INCLUIR JERARQUIA.xlsx"
hoja_renovaciones = "Hoja1"
umbral_similitud = 65

# ---------- Leer archivo jerarquía ----------
df_jerarquia = pd.read_excel(archivo_jerarquia)
df_jerarquia = df_jerarquia.rename(columns=lambda x: str(x).strip())

# Diccionario: Alias → Nivel 2
map_alias_director = df_jerarquia.set_index("Alias")["Nivel 2"].to_dict()
alias_jerarquia = list(map_alias_director.keys())

# ---------- Leer archivo de renovaciones ----------
df_renovaciones = pd.read_excel(archivo_renovaciones, sheet_name=hoja_renovaciones)
df_renovaciones = df_renovaciones.rename(columns=lambda x: str(x).strip())

# Normalizar claves del diccionario
df_jerarquia["Alias_normalizado"] = df_jerarquia["Alias"].apply(normalizar_alias)
map_alias_director = dict(zip(df_jerarquia["Alias_normalizado"], df_jerarquia["Nivel 2"]))
alias_jerarquia = list(map_alias_director.keys())

# Normalizar los alias de renovaciones
df_renovaciones["Alias_normalizado"] = df_renovaciones["Alias"].apply(normalizar_alias)

# ---------- Primera pasada: coincidencia exacta ----------
df_renovaciones.insert(
    0,
    "Codigo Director",
    df_renovaciones["Alias_normalizado"].map(map_alias_director)
)

# ---------- Función para buscar coincidencia difusa ----------
def buscar_director_fuzzy(alias_normalizado):
    if pd.isna(alias_normalizado) or alias_normalizado == "":
        return None
    resultado = process.extractOne(
        alias_normalizado,
        alias_jerarquia,
        scorer=fuzz.token_sort_ratio
    )
    if resultado and resultado[1] >= umbral_similitud:
        return map_alias_director[resultado[0]]
    return None

# ---------- Segunda y tercera pasadas ----------
for intento in range(2):  # segunda y tercera pasada
    sin_director = df_renovaciones["Codigo Director"].isna()
    df_renovaciones.loc[sin_director, "Codigo Director"] = df_renovaciones.loc[sin_director, "Alias_normalizado"]\
        .apply(buscar_director_fuzzy)

# ---------- Exportar resultado ----------
df_renovaciones.to_excel("BASE_RENOVACIONES_con_directores.xlsx", index=False)
print("✅ Archivo generado: BASE_RENOVACIONES_con_directores.xlsx")

# ---------- Exportar alias sin director para revisión ----------
sin_director = df_renovaciones[df_renovaciones["Codigo Director"].isna()]
sin_director[["Alias", "Alias_normalizado"]].drop_duplicates()\
    .to_csv("ALIAS_SIN_DIRECTOR.csv", index=False)

print(f"📄 Log generado con {len(sin_director)} alias sin director: ALIAS_SIN_DIRECTOR.csv")