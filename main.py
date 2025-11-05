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

def add_chart(pdf, df, xcol, ycol, title, kind="line", xlabel=None, ylabel=None):
    "Criar grafico no Pdf"
    fig, ax = plt.subplots(figsize=(11.69,8.27))
    if kind == "bar":
        ax.bar(df[xcol], df[ycol], align='center')
    else:
        ax.plot(df[xcol], df[ycol])
        ax.set_title(title, fontsize=14, fontweight='bold',)
        if xlabel:
            ax.set_xlabel(xlabel)
            if ylabel:
                ax.set_ylabel(ylabel)
            fig.autofmt_xdate(rotation=45)
            pdf.savefig(fig, bbox_inches='tight')
            plt.close(fig)

#Leitura da planilha

xls = pd.ExcelFile(SRC_PATH)
df_vendas = pd.read_excel(xls, "Fato_vendas")
df_produto = pd.read_excel(xls, "Dim_Produto")
df_cliente = pd.read_excel(xls, "Dim_Cliente")
df_canal = pd.read_excel(xls, "Dim_Canal")
df_regiao = pd.read_excel(xls, "Dim_Regiao")
df_data = pd.read_excel(xls, "Dim_Data")

#Limpeza dos dados

for df in [df_vendas, df_produto, df_cliente, df_canal, df_regiao, df_data]:
    strip_strings(df)


df_vendas["DataPedido"] = pd.to_datatime(df_vendas["DataPedido"], errors="coerce")
df_vendas["DataEnvio"] = pd.to_datetime(df_vendas["DataEnvio"], errors="coerce")

num_cols = [
    "Quantidade","PreçoUnitario", "taxaDesconto", "Valordesconto", "ValorBruto",
    "Receita"," CustoTotal", "DespesasVariaveis", "LucroBruto", "Lucro"
]

for c in num_cols:
    df_vendas[c] = pd.to_numeric(df_vendas[c], errors="coerce")

#Remover duplicados

before = len(df_vendas)
df_vendas.drop_duplicates(subset= ["ID_Pedido", "ID_Produto"], inplace=True)
dup_removed = before - len(df_vendas)


#Corrigir chaves invalidas


mask_prod = df_vendas["ID_Produto"].isin(df_produto["ID_Produto"])
mask_cli = df_vendas["ID_Cliente"].isin(df_cliente["ID_Cliente"])
mask_canal = df_vendas["ID_Canal"].isin(df_canal["ID_Canal"])
df_vendas = df_vendas[mask_prod & mask_cli & mask_canal]


#3 KPIs e Ànalises


df_vendas["AnoMes"] = df_vendas["DataPedido"].dt.to.period("M").astype(str)

kpi = {
    "Receita Total": df_vendas["Receita"].sum(),
    "Lucro Total": df_vendas["Lucro"].sum(),
    "Margem(%)": df_vendas["LucroBruto"].sum() / df_vendas["Receita"].sum() * 100,
    "Ticket Médio (R$)": df_vendas["Receita"].mean(),
    "Quantidade": df_vendas["Quantidade"].sum(),
    "Linhas de Venda": len(df_vendas)
}

#Agrupamento


receita_mes = df_vendas.groupby("AnoMes", as_index=False)["Receita"].sum().sort_values("AnoMes")
vendas_cat = df_vendas.merge(df_produto[["ID_Produto","Categoria"]], on="ID_Produto") \
    .groupby("Categoria", as_index=False)[["Receita","Lucro"]].sum().sort_values("Receita", ascending=False)
vendas_canal = df_vendas.merge(df_canal, on="ID_Canal") \
    .groupby("Canal", as_index=False)[["Receita","Lucro"]].sum().sort_values("Receita", ascending=False)
vendas_regiao = df_vendas.groupby("Região", as_index=False)[["Receita","Lucro"]].sum().sort_values("Receita", ascending=False)

top_produtos = df_vendas.merge(df_produto[["ID_Produto","NomeProduto","Categoria"]], on="ID_Produto") \
    .groupby(["NomeProduto","Categoria"], as_index=False)[["Receita","Lucro","Quantidade"]].sum() \
    .sort_values("Receita", ascending=False).head(20)
















