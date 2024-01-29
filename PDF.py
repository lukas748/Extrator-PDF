import PyPDF2
import pandas as pd
import re
import os
from tkinter import filedialog, messagebox

def extrair_informacoes_pdf(caminho_pdf):

    with open(caminho_pdf, 'rb') as arquivo_pdf:
        leitor_pdf = PyPDF2.PdfReader(arquivo_pdf)

        dados = []

        lista_palavras_chave = [
            'ARMA DE FOGO', 'MACONHA', 'COCAÍNA', 'COCAINA', 'CRACK', 'PROJÉTIL', 'PROJETIL', 'MUNICAO', 'MUNIÇÃO', 'MUNICÃO', 'MUNIÇAO', 'SWAB', 'FACA', 'CELULAR'
        ]

        for numero_pagina in range(0, len(leitor_pdf.pages), 1):
            # A cada duas páginas, extrai as informações
            paginas = leitor_pdf.pages[numero_pagina]


            texto = paginas.extract_text()

            # Use expressões regulares para encontrar os valores desejados
            pcnet_match = re.search(r'Nº PCnet:\s*(\S+)', texto)
            reds_match = re.search(r'Nº REDS:\s*(\S+)', texto)
            fav_match = re.search(r'FICHA DE ACOMPANHAMENTO DE VESTÍGIO N.º:\s*(\S+)', texto)
            unidade_match = re.search(r'Unidade Origem:\s*.*?/\s*(.+)', texto)
            material_match = re.search(r'Material:\s*(.*?)\s*Número/Descrição', texto, re.DOTALL)
            req_match = re.search(r'Exame pericial \(nº requisição\):\s*(\d{4}-\d{9}(?:,\s*\d{4}-\d{9})*)', texto)

            pcnet = None
            if pcnet_match:
                # Extrair o penúltimo conjunto de números (9 dígitos)
                pcnet_valor_match = re.search(r'-(\d{9})-', pcnet_match.group(1))
                pcnet = pcnet_valor_match.group(1) if pcnet_valor_match else None
            reds = reds_match.group(1) if reds_match else None
            fav = fav_match.group(1) if fav_match else None
            unidade = unidade_match.group(1) if unidade_match else None
            req = req_match.group(1) if req_match else None
            if req_match:
                # Tenta extrair os últimos 9 dígitos
                # req_ultimos_match = re.search(r'-(\d{9})', req_match.group(1))
                # req = req_ultimos_match.group(1) if req_ultimos_match else None
                numeros = [int(num.split('-')[1]) for num in req_match.group(1).split(',')]
                req = max(numeros)
            material = None

            if material_match:
                # Verificar se alguma palavra da lista está presente no texto
                for palavra_chave in lista_palavras_chave:
                    if palavra_chave.upper() in material_match.group(1).upper():
                        material = palavra_chave
                        break
            
            if fav is not None:
                fav = int(fav)
            
            if req is not None:
                req = int(req)

            if pcnet is not None:
                pcnet = int(pcnet)

            dados.append({
                'FAV': fav,
                'Requisição': req,
                'PCnet': pcnet,
                'REDS': reds,
                'Unidade': unidade,
                'Material': material
            })

    nova_lista = [valores for valores in dados if valores.get('REDS') is not None]
    return nova_lista

def salvar_em_excel(dados, caminho_excel):
    try:
        df = pd.DataFrame(dados)
        df.to_excel(caminho_excel, index=False, engine='openpyxl')
        messagebox.showinfo("Sucesso","Os materiais foram cadastrados com sucesso!")
    except Exception as e:
        # Exibe uma caixa de diálogo de erro
        messagebox.showerror("Erro", f"Ocorreu um erro ao salvar os dados no arquivo Excel:\n{str(e)}")

if __name__ == "__main__":

    x = 1
    while x == 1:
        caminho_do_pdf = filedialog.askopenfilename()
        if caminho_do_pdf.lower().endswith(".pdf"):
            informacoes_extraidas = extrair_informacoes_pdf(caminho_do_pdf)
            x = 0
        elif not caminho_do_pdf:
            messagebox.showwarning("Alerta","Operação cancelada pelo usuário.")
            x = 0
        else:
            messagebox.showwarning("Alerta","O arquivo selecionado não é um PDF!")


    caminho_do_excel = os.path.abspath(os.path.join(os.getcwd(), "Dados Etiqueta.xlsx"))

    if caminho_do_pdf:
        salvar_em_excel(informacoes_extraidas, caminho_do_excel)
