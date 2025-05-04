import pandas as pd
from fpdf import FPDF
import unicodedata
import matplotlib.pyplot as plt
import os

# Função que usei para normalizar nome das colunas,lembrete :(tira acento, espaço e coloca minúsculo)
def normalizar_coluna(col):
    col = col.strip().lower()
    col = ''.join(
        c for c in unicodedata.normalize('NFD', col)
        if unicodedata.category(c) != 'Mn'
    )
    return col

# Entrada do arquivo
arquivo = input("Digite o nome do arquivo Excel (exemplo: data/teste.xlsx): ")
df = pd.read_excel(arquivo)

# Normalizei todas as colunas
colunas_original = df.columns
colunas_normalizadas = [normalizar_coluna(c) for c in colunas_original]
mapa_colunas = dict(zip(colunas_normalizadas, colunas_original))

# Agora usei o nome que está escrito no excel 
col_quantidade = mapa_colunas['quantidade']
col_valorunitario = mapa_colunas['valorunitario']

# Calculei Valor Total
df['Valor Total'] = df[col_quantidade] * df[col_valorunitario]

# Ordenei por Valor Total (menor para maior para o gráfico)
df_ordenado = df.sort_values(by='Valor Total', ascending=True)

# Calculei o total gerado
total_geral = df['Valor Total'].sum()

# Criei um  gráfico de barras (menor para maior)
plt.figure(figsize=(8, 4))
plt.barh(df_ordenado[mapa_colunas['produto']], df_ordenado['Valor Total'], color='#3498db')
plt.xlabel('Valor Total (R$)')
plt.ylabel('Produto')
plt.title('Vendas por Produto (Menor → Maior)')
plt.tight_layout()

# Criei uma pasta para gráficos 
os.makedirs('reports', exist_ok=True)
grafico_path = 'reports/grafico_vendas.png'
plt.savefig(grafico_path, dpi=150)
plt.close()

# Criei PDF moderno
pdf = FPDF(orientation='P', unit='mm', format='A4')
pdf.add_page()
pdf.set_font('Arial', 'B', 16)
pdf.set_text_color(33, 37, 41)
pdf.cell(0, 10, "Relatório de Vendas", ln=True, align='C')

# Adicionei gráfico no PDF
pdf.ln(5)
pdf.image(grafico_path, x=25, w=160)  # Centralizado e redimensionado
pdf.ln(5)

# Tabela de dados
pdf.set_font('Arial', 'B', 12)
pdf.set_fill_color(52, 152, 219)  # Azul moderno
pdf.set_text_color(255, 255, 255)
pdf.cell(60, 8, 'Produto', border=1, align='C', fill=True)
pdf.cell(40, 8, 'Quantidade', border=1, align='C', fill=True)
pdf.cell(40, 8, 'Valor Unitário', border=1, align='C', fill=True)
pdf.cell(40, 8, 'Valor Total', border=1, align='C', fill=True)
pdf.ln()

pdf.set_font('Arial', '', 11)
pdf.set_text_color(33, 37, 41)
for _, row in df_ordenado.iterrows():
    pdf.cell(60, 8, str(row[mapa_colunas['produto']]), border=1)
    pdf.cell(40, 8, str(row[col_quantidade]), border=1, align='C')
    pdf.cell(40, 8, f"R$ {row[col_valorunitario]:.2f}", border=1, align='R')
    pdf.cell(40, 8, f"R$ {row['Valor Total']:.2f}", border=1, align='R')
    pdf.ln()

# Total que gerei
pdf.set_font('Arial', 'B', 12)
pdf.cell(140, 8, "TOTAL GERAL:", border=1, align='R')
pdf.cell(40, 8, f"R$ {total_geral:.2f}", border=1, align='R')

# Gerei o PDF
pdf.output("reports/relatorio_vendas.pdf")

# Opcional: remover imagem depois de usar (se não quiser guardar)
# os.remove(grafico_path)

print("✅ Relatório gerado com sucesso em reports/relatorio_vendas.pdf")
