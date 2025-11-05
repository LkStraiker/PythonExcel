import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from pathlib import Path

# Caminhos dos arquivos
SRC_PATH = Path("Fabrica_Roupas_PT.xlsx")
CLEAN_XLSX_PATH = Path("Fabrica_Roupas_PT_CLEAN.xlsx")
PDF_PATH = Path("Relatorio_Fabrica_Roupas.pdf")

# -----------------------------
# Funções auxiliares
# -----------------------------
def strip_strings(df: pd.DataFrame) -> pd.DataFrame:
    """Remove espaços e substitui strings vazias por NaN"""
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].astype(str).str.strip()
        df[col].replace({"": np.nan}, inplace=True)
    return df

def auto_widths(writer, df: pd.DataFrame, sheet_name: str, minw=10, maxw=35):
    """Ajusta largura das colunas no Excel"""
