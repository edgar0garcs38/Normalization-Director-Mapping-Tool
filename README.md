# Alias Normalization & Director Mapping Tool

This Python script enriches a base dataset of agent renewals by assigning each agent their corresponding director code using a reference hierarchy file. It performs both exact and fuzzy matching on agent aliases to ensure robust mapping, even when names are inconsistently formatted.

## Files Required

Place the following Excel files in the same directory as the script:

- `REPORTE DE JERARQUIAS AGENTES A 1 SEP.xlsx`  
  Contains agent aliases and their corresponding "Nivel 2" director codes.
- `BASE DE RENOVACIONES 2022 A 2024 PARA INCLUIR JERARQUIA.xlsx`  
  Contains renewal records with agent aliases to be matched.

## How It Works

1. **Normalization**  
   All aliases are normalized by:
   - Removing accents
   - Converting to uppercase
   - Trimming and collapsing whitespace

2. **Exact Matching**  
   Attempts to match normalized aliases directly to the hierarchy dictionary.

3. **Fuzzy Matching**  
   For unmatched aliases, applies fuzzy matching using `rapidfuzz` with a similarity threshold of `65`.

4. **Multi-Pass Matching**  
   Runs two additional fuzzy matching passes to improve coverage.

5. **Export Results**  
   - `BASE_RENOVACIONES_con_directores.xlsx`: Enriched dataset with director codes.
   - `ALIAS_SIN_DIRECTOR.csv`: List of unmatched aliases for manual review.

## Dependencies

Install required packages via pip:

```bash
pip install pandas rapidfuzz openpyxl
