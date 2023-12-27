import pandas as pd
import matplotlib.pyplot as plt

# Carregar as planilhas
planilha1 = pd.read_excel("Analise_reincidencia_10_2023.xlsx")
planilha2 = pd.read_excel("Analise_reincidencia_11_2023.xlsx")

# Selecionar as colunas relevantes
colunas_interesse = ["Host", "Number of status changes"]
dados1 = planilha1[colunas_interesse]
dados2 = planilha2[colunas_interesse]

# Agrupar os dados por 'Host' e somar as mudanças de status
agrupado1 = dados1.groupby("Host").sum().reset_index()
agrupado2 = dados2.groupby("Host").sum().reset_index()

# Mesclar os dados das duas planilhas
merged_data = pd.merge(
    agrupado1, agrupado2, on="Host", suffixes=("_planilha1", "_planilha2")
)

# Preparar dados para o gráfico
hosts = merged_data["Host"]
status_changes1 = merged_data["Number of status changes_planilha1"]
status_changes2 = merged_data["Number of status changes_planilha2"]

# Gerar o gráfico de linhas horizontais com rotação nos rótulos
plt.figure(figsize=(10, 6))
plt.plot(hosts, status_changes1, marker="o", label="Outubro")
plt.plot(hosts, status_changes2, marker="o", label="Novembro")

plt.xlabel("Host")
plt.ylabel("Número de Mudanças de Status")
plt.title("Comparação de Mudanças de Status por Host")
plt.legend()
plt.grid(axis="y")

# Adicionar rotação nos rótulos do eixo x
plt.xticks(rotation=45, ha="right")

plt.tight_layout()

# Exibir o gráfico
plt.show()
