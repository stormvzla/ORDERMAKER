import pandas as pd
import random
import tkinter as tk
from tkinter import messagebox, filedialog
from openpyxl import Workbook

# Função para gerar o orçamento e salvar em uma planilha Excel
def gerar_orcamento():
    try:
        # Obtém o valor total desejado da interface gráfica
        valor_total = float(entry_valor.get())
        
        # Converte a tabela para um DataFrame do pandas
        # O arquivo CSV usa ";" como delimitador e tem um cabeçalho na segunda linha
        df = pd.read_csv("c:/pasta/produtos.csv", delimiter=";", skiprows=1, encoding='latin1')
        
        # Renomeia as colunas para facilitar o acesso
        df.columns = ['Descrição', 'Código', 'Unidade', 'Preço']
        
        # Remove possíveis espaços em branco nos nomes das colunas
        df.columns = df.columns.str.strip()
        
        # Converte a coluna "Preço" para numérico, tratando vírgulas como separadores decimais
        df['Preço'] = df['Preço'].str.replace('.', '').str.replace(',', '.').astype(float)
        
        # Verifica se há produtos no DataFrame
        if df.empty:
            messagebox.showerror("Erro", "Nenhum produto encontrado no arquivo CSV.")
            return
        
        # Lista para armazenar os produtos selecionados e suas quantidades
        produtos_selecionados = []
        valor_atual = 0
        
        # Define o valor máximo permitido (valor_total + margem de tolerância)
        margem_tolerancia = 0.1
        valor_maximo = valor_total * (1 + margem_tolerancia)
        
        # Enquanto o valor atual for menor que o valor máximo permitido
        while valor_atual < valor_maximo:
            # Escolhe um produto aleatoriamente
            produto = df.sample(1).iloc[0]
            
            # Escolhe uma quantidade aleatória para o produto (entre 1 e quantidade_maxima)
            quantidade_maxima = 10
            quantidade = random.randint(1, quantidade_maxima)
            
            # Calcula o valor total para esse produto
            valor_produto = produto['Preço'] * quantidade
            
            # Verifica se adicionar esse produto ultrapassa o valor máximo permitido
            if valor_atual + valor_produto > valor_maximo:
                # Ajusta a quantidade para não ultrapassar o valor máximo
                quantidade = int((valor_maximo - valor_atual) / produto['Preço'])
                if quantidade < 1:
                    break  # Se não for possível adicionar pelo menos 1 unidade, para o loop
            
            # Adiciona o produto e a quantidade à lista de selecionados
            produtos_selecionados.append({
                'Descrição': produto['Descrição'],
                'Código': produto['Código'],
                'Preço Unitário': produto['Preço'],
                'Quantidade': quantidade,
                'Valor Total': produto['Preço'] * quantidade
            })
            
            # Atualiza o valor atual
            valor_atual += produto['Preço'] * quantidade
        
        # Calcula o valor total dos produtos selecionados
        valor_total_produtos = sum(item['Valor Total'] for item in produtos_selecionados)
        
        # Calcula o desconto necessário (em valor)
        desconto = valor_total_produtos - valor_total
        if desconto < 0:
            desconto = 0  # Se o valor total dos produtos for menor que o desejado, não há desconto
        
        # Aplica o desconto
        valor_final = valor_total_produtos - desconto
        
        # Cria um arquivo Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "Orçamento"
        
        # Adiciona cabeçalhos à planilha
        ws.append(["Descrição", "Código", "Quantidade", "Preço Unitário", "Valor Total"])
        
        # Adiciona os produtos selecionados à planilha
        for item in produtos_selecionados:
            ws.append([item['Descrição'], item['Código'], item['Quantidade'], item['Preço Unitário'], item['Valor Total']])
        
        # Adiciona o valor total, desconto e valor final à planilha
        ws.append([])  # Linha em branco
        ws.append(["Valor total dos produtos:", valor_total_produtos])
        ws.append(["Desconto aplicado:", desconto])
        ws.append(["Valor final do orçamento:", valor_final])
        
        # Salva o arquivo Excel
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
        if file_path:
            wb.save(file_path)
            messagebox.showinfo("Sucesso", f"Orçamento salvo com sucesso em {file_path}")
    
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

# Cria a interface gráfica
root = tk.Tk()
root.title("Gerador de Orçamento")

# Label e Entry para o valor total
label_valor = tk.Label(root, text="Digite o valor total desejado para o orçamento: R$")
label_valor.pack(pady=5)

entry_valor = tk.Entry(root)
entry_valor.pack(pady=5)

# Botão para gerar o orçamento
button_gerar = tk.Button(root, text="Gerar Orçamento", command=gerar_orcamento)
button_gerar.pack(pady=10)

# Inicia o loop da interface gráfica
root.mainloop()