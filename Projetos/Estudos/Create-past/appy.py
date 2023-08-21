import os

num_pastas = int(input("Quantas pastas deseja criar: "))
nome_pasta_master = input("Digite o nome da pasta master: ")
print("Criando pastas...")

# Define o caminho para a pasta "master"
caminho_pasta = os.path.join("C:\\Users\\x\\Desktop", nome_pasta_master)

# Cria a pasta "master" se ela ainda não existir
if not os.path.exists(caminho_pasta):
    os.mkdir(caminho_pasta)

# Cria o número correspondente de pastas filhas
for i in range(1, num_pastas + 1):
    nome_pasta = str(i)
    caminho_pasta_filho = os.path.join(caminho_pasta, nome_pasta)
    os.mkdir(caminho_pasta_filho)

print("Pastas criadas com sucesso!")
