import xmltodict
import os
import pandas as pd
from tkinter import Tk, filedialog

Tk().withdraw()  # Oculta a janela principal do Tkinter
pasta = filedialog.askdirectory(title="Selecione a pasta onde estão as NFs")

# Se o usuário cancelar, encerra o programa
if not pasta:
    print("Nenhuma pasta selecionada. Encerrando...")
    exit()

def pegar_infos(arquivo, valores):
    
    with open(f'nfs/{arquivo}', "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)

        if "NFe" in dic_arquivo:
            infos_nfs = dic_arquivo["NFe"]["infNFe"]
        else:
            infos_nfs = dic_arquivo["nfeProc"]["NFe"]["infNFe"]

        numero_nota = infos_nfs["@Id"]
        empresa_emissora = infos_nfs["emit"]["xNome"]
        nome_cliente = infos_nfs["dest"]["xNome"]
        endereco = infos_nfs["dest"]["enderDest"]

        if "vol" in infos_nfs["transp"]:
            peso = infos_nfs["transp"]["vol"]["pesoB"]
        else:
            peso = "Não informado"

        valores.append([numero_nota, empresa_emissora, nome_cliente, endereco, peso])
    
arquivos = os.listdir(pasta)

colunas = ["numero_nota", "empresa_emissora","nome_cliente","endereco","peso"]
valores = []

for arquivo in arquivos:
    pegar_infos(arquivo, valores)

tabela = pd.DataFrame(columns=colunas, data=valores)

tabela.to_excel("NotasFiscais.xlsx", index=False)