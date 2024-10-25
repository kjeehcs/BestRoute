import pandas as pd
import unidecode
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# Função para padronizar os textos, removendo acentos, aspas simples e duplas
def padronizar_texto(texto):
    texto = unidecode.unidecode(str(texto)).lower()
    return texto.replace("'", "").replace('"', '')

# Função para carregar e padronizar os dados da planilha
def carregar_planilha(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo)
    df['Origem'] = df['Origem'].apply(padronizar_texto)
    df['Destino'] = df['Destino'].apply(padronizar_texto)
    df['UF'] = df['UF'].apply(padronizar_texto)
    return df

# Função para comparar os prazos das planilhas
def comparar_planilhas(planilhas_dfs, planilhas_nomes):
    resultado = []
    percursos_unicos = pd.concat(planilhas_dfs)[['Origem', 'Destino', 'UF']].drop_duplicates()

    # Percorre todos os percursos únicos de todas as planilhas
    for _, percurso in percursos_unicos.iterrows():
        origem1, destino1, uf1 = percurso['Origem'], percurso['Destino'], percurso['UF']
        prazos = []
        mais_rapida = None
        prazo_mais_rapido = float('inf')
        empate = False

        # Verifica cada planilha
        for i in range(len(planilhas_dfs)):
            match_df = planilhas_dfs[i][(planilhas_dfs[i]['Origem'] == origem1) & 
                                        (planilhas_dfs[i]['Destino'] == destino1) & 
                                        (planilhas_dfs[i]['UF'] == uf1)]
            if not match_df.empty:
                prazo_atual = match_df.iloc[0]['Quantidade de Dias Máximo para entrega']
                prazos.append(prazo_atual)
                
                # Comparar prazos
                if prazo_atual < prazo_mais_rapido:
                    mais_rapida = planilhas_nomes[i]
                    prazo_mais_rapido = prazo_atual
                    empate = False
                elif prazo_atual == prazo_mais_rapido:
                    empate = True
            else:
                prazos.append("transportadora não faz percurso")

        resultado.append({
            'Origem': origem1,
            'Destino': destino1,
            'UF': uf1,
            **{f'Prazo {planilhas_nomes[i]}': prazos[i] for i in range(len(prazos))},
            'Mais Rápida': 'Empate' if empate else (mais_rapida if mais_rapida else "Nenhuma")
        })

    return pd.DataFrame(resultado)

# Função para escolher e carregar arquivos (até 5 planilhas)
def selecionar_arquivo(n):
    global arquivos, labels_arquivos
    arquivo = filedialog.askopenfilename(title=f"Selecione a planilha {n+1}", filetypes=[("Excel files", "*.xlsx")])
    if arquivo:
        arquivos[n] = arquivo
        labels_arquivos[n].config(text=arquivo)

# Função para executar a comparação
def executar_comparacao():
    try:
        planilhas_dfs = []
        planilhas_nomes = []

        # Carrega os dados das planilhas selecionadas
        for i, arquivo in enumerate(arquivos):
            if arquivo:
                planilhas_dfs.append(carregar_planilha(arquivo))
                planilhas_nomes.append(os.path.basename(arquivo).replace(".xlsx", ""))

        # Verifica se ao menos 2 planilhas foram selecionadas
        if len(planilhas_dfs) < 2:
            messagebox.showwarning("Aviso", "Selecione pelo menos 2 planilhas para comparação!")
            return

        # Compara as planilhas
        resultado = comparar_planilhas(planilhas_dfs, planilhas_nomes)

        # Exibe o resultado
        messagebox.showinfo("Resultado", "Comparação realizada com sucesso!")
        print(resultado)

        # Opção para salvar o resultado em um arquivo Excel
        opcao = messagebox.askyesno("Salvar", "Deseja salvar o resultado em um arquivo Excel?")
        if opcao:
            resultado.to_excel('resultado_comparacao.xlsx', index=False)
            messagebox.showinfo("Salvo", "Resultado salvo como 'resultado_comparacao.xlsx'.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

# Criar a janela principal
root = tk.Tk()
root.title("BestRoute")
root.geometry("500x400")

# Texto de instruções
label_instrucoes = tk.Label(root, text="Selecione os arquivos das planilhas para comparar", font=("Arial", 12))
label_instrucoes.pack(pady=10)

# Lista para armazenar caminhos de arquivos
arquivos = [None] * 5

# Labels e botões para seleção de arquivos
labels_arquivos = []
for i in range(5):
    btn_arquivo = tk.Button(root, text=f"Selecionar Planilha {i+1}", command=lambda i=i: selecionar_arquivo(i))
    btn_arquivo.pack(pady=5)

    label_arquivo = tk.Label(root, text="Nenhuma planilha selecionada", fg="gray")
    label_arquivo.pack()
    labels_arquivos.append(label_arquivo)

# Botão para executar a comparação
btn_comparar = tk.Button(root, text="Executar Comparação", command=executar_comparacao)
btn_comparar.pack(pady=20)

# Iniciar a interface gráfica
root.mainloop()
