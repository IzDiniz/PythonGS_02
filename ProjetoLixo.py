# pip install openpyxl - Necessário instalar a biblioteca antes de tudo
import openpyxl # Importando a biblioteca após instalar

# Função criada para adicionar itens novos na categoria desejada
def adicionar_item(categorias):
    while True:
        escolha = input("Digite a categoria: ")
        categoria = escolha.capitalize()
        # Método de verificação onde vê se a categoria existe ou não na tabela
        if categoria in categorias:
            adicionando_item = input("Digite o novo item: ")
            novo_item = adicionando_item.capitalize()
            categorias[categoria].append(novo_item)
            # Opção criada caso o usuário deseje adicionar mais de um item na tabela
            opcao = input("Deseja adicionar mais itens? (1 - Sim, 2 - Não): ")
            if opcao == '2':
                break
        else:
            print("Categoria não existe.")

# Função criada para excluir itens da categoria desejada
def excluir_item(categorias):
    while True:
        escolha = input("Digite a categoria para excluir um item: ")
        categoria = escolha.capitalize()
        # Método de verificação onde vê se a categoria existe ou não na tabela
        if categoria in categorias:
            print(f"Lista de itens na categoria {categoria}:")
            for i, item in enumerate(categorias[categoria], start=1):
                print(f"{i}. {item}")
            # Opção criada onde selecionar qual item ele deseja excluir de acordo com o número gerado pelo código
            try:
                indice = int(input("Digite o número do item que deseja excluir: ")) - 1
                if 0 <= indice < len(categorias[categoria]):
                    item_excluido = categorias[categoria].pop(indice)
                    print(f"Item '{item_excluido}' excluído com sucesso!")
                else:
                    print("Número inválido. Tente novamente.")
            except ValueError:
                print("Entrada inválida. Digite um número válido.")
        else:
            print("Categoria não existe.")
        # Opção criada caso o usuário deseje excluir mais de um item da tabela
        opcao = input("Deseja excluir mais itens? (1 - Sim, 2 - Não): ")
        if opcao == '2':
            break

# Função criada para criar um arquivo em excell caso o usuário deseje
def criar_planilha_excel(categorias):
    wb = openpyxl.Workbook()  # Cria uma nova planilha em branco
    planilha = wb.active  # Seleciona a planilha ativa
    # Preenche a primeira linha com os nomes das colunas
    planilha.append(["Categoria", "Itens"])
    # Preenche as demais linhas com os dados das categorias
    for categoria, itens in categorias.items():
        planilha.append([categoria] + itens)
    # Salva a planilha com o nome "Tabela_Categorias.xlsx"
    wb.save("Tabela_Categorias.xlsx")

    print("Planilha criada com sucesso!")

# Função feita para contar cada item da categoria de forma individual
def contar_itens_por_categoria(categorias):
    for categoria, itens in categorias.items():
        quantidade = len(itens)
        print(f"A categoria {categoria} tem {quantidade} itens.")

# Função criada para contar todos os itens de todas as categorias juntas
def somar_itens_por_categoria(categorias):
    total = 0
    for itens in categorias.values():
        total += len(itens)
    return total

# Função de menu, usuário escolhe o que deseja fazer
def menu():
    print("Escolha uma opção:")
    print("1. Adicionar item")
    print("2. Excluir item")
    print("3. Criar arquivo Excel")
    print("4. Sair")

# Tabela de categorias e itens
categorias = {
    'Papel': ['Caderno', 'Sulfite'],
    'Plástico': ['Garrafa Pet', 'Vasilha'],
    'Metal': ['Frigideira', 'Latas de alumínio'],
    'Vidro': ['Copo de Vidro', 'Potes de vidro'],
    'Outros': ['Alimentos', 'Maquiagens']
}

# Verifica qual opção o usuário escolheu no menu, e se é verdadeira ou não
while True:
    menu()
    opcao_menu = input("Digite o número da opção desejada: ")

    if opcao_menu == '1':
        adicionar_item(categorias)
    elif opcao_menu == '2':
        excluir_item(categorias)
    elif opcao_menu == '3':
        criar_planilha_excel(categorias)
    elif opcao_menu == '4':
        break
    else:
        print("Opção inválida. Tente novamente.")

# Prints finais
print('\n' "Tabela de categorias atualizada:")
print(categorias, '\n')
print(contar_itens_por_categoria(categorias))
print('A soma de todos os itens de cada categoria é de:', somar_itens_por_categoria(categorias), 'itens')
