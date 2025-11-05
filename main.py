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
    ws = writer.sheets[sheet_name]
    for i, col in enumerate(df.columns):
        try:
            series = df[col].astype(str)
            width = max(series.map(len).max(), len(str(col))) + 2
            width = max(minw, min(maxw, width))
            ws.set_column(i, i, width)
        except Exception:
            ws.set_column(i, i, minw)

def add_table(pdf,df, title):
    "Criar a tabela no PDF"
    fig, ax = plt.subplots(figsize=(11.69,8.27))
    ax.axis("Off")
    ax.set_title(title, fontsize=14, fontweight='bold', pad=10)
    show = df.head(20).copy()
    tbl = ax.table(cellText=show.values, colLabels=show.columns, loc='center')
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(8)
    tbl.scale(1, 1.2)
    pdf.savefig(bbox_inches='tight')
    plt.close(fig)









