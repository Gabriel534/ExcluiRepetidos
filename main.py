import openpyxl
import os
from pprint import pprint
import re
import shutil

PASTA_FOTOS = "Arquivos"
PLANILHA_SAIDA = "CONTROLE_AT.xlsx"
PLANILHA_ENTRADA = 'CONTROLE.xlsx'

if __name__ == "__main__":
    # Carregar o arquivo XLSX
    workbook = openpyxl.load_workbook('CONTROLE.xlsx')

    # Abre as planilhas
    ControleNotebooks = workbook["Controle notebooks"]
    Componentes = workbook["Componentes"]

    cont = 1
    info = "AS"
    info2 = "AS"
    componentesLista:dict[str, list[str]] = {}
    modelosAvaliados = []
  
   # Extrai  componentes da planilha
    while True:
        info = Componentes[f"A{cont}"].value
        info2 = Componentes[f"B{cont}"].value
        try:
            componentesLista[info].append(info2)
        except KeyError:
            componentesLista[info] = []
            componentesLista[info].append(info2)
        
        if info == None and info2 == None: break
        cont += 1

    
    if not os.path.exists(PASTA_FOTOS): os.mkdir(PASTA_FOTOS)

    exists = False
    cont = 0

#-------------------------------------------Inconsistências-----------------------------------------------------
    for root, dirs, files in os.walk(PASTA_FOTOS):
            for dir in dirs:
                print(f"Verificando pasta {dir}")
                for root, dirs, files2 in os.walk(f"{PASTA_FOTOS}\\{dir}"):
                    
                    # Remove componentes da planilha que não tenham arquivo do modelo em questão
                    try:
                        for componente in componentesLista.get(dir, []):
                            exists = False
                            for file in files2:
                                if re.search(r'(.+?)\..+?', file).group(1).strip() == componente.strip():
                                    print(f"{componente} existe na planilha")
                                    exists = True
                            if not exists:
                                componentesLista[dir].remove(componente)
                                print(f"Componente {componente} removido  da planilha (arquivo não encontrado nas fotos)")
                        modelosAvaliados.append(dir)

                    except KeyError:
                        print(f"Skipando {dir} (Não encontrado)")

                    # Indica todos os arquivos que não estejam na planilha
                    for file in files2:
                        exists = False
                        try:
                            for componente in componentesLista.get(dir, []):
                                if re.search(r'(.+?)\..+?', file).group(1).strip() == componente.strip():
                                    print(f"Arquivo {componente} existe na planilha")
                                    exists = True
                            if not exists:
                                print(f"Arquivo {file} não existe na planilha, favor verificar")
                        except KeyError:
                            print(f"Skipando {dir} (Não encontrado)")
                print()
#-------------------------------------------Inconsistências-----------------------------------------------------
    input("Continuar? \n:> ")
#-----------------------------------------Eliminar-Repetidos-------------------------------------------------------
    todosComponentes: list[str] = []
    for modelo, componentes in componentesLista.items():
        for componente in componentes:
            todosComponentes.append(componente)

    print("Grande lista feita")

    for modelo in modelosAvaliados:
        print("Modelo: "+modelo)
        for root, dirs, files in os.walk(PASTA_FOTOS+"\\"+modelo):
            for componente in files:
                repeticoes = 0

                for a in todosComponentes:
                    if a == re.search(r'(.+?)\..+?', componente).group(1).strip():
                        repeticoes += 1

                print(componente+" "+str(repeticoes))
                if repeticoes > 1:
                    
                    componentesLista[modelo].remove(re.match(r'^(.+?)\..+$', componente).group(1))
                    print(f"Componente {componente} removido da planilha")
                    for file in os.listdir(PASTA_FOTOS+"\\"+modelo):
                        # Verifica se o nome do arquivo (sem extensão) corresponde ao nome que você deseja apagar
                        if file == componente:
                            caminho_completo = os.path.join(PASTA_FOTOS+"\\"+modelo, file)

                    os.remove(caminho_completo)
                    print(f"Componente {caminho_completo} apagado das fotos")
        print("Total: " + str(len(componentesLista[modelo])) + " componentes\n")
    #-----------------------------------------Eliminar-Repetidos-------------------------------------------------------
    
    #----------------------------------------Salvar planilha ---------------------------------------------------------
    print("Salvando planilha\n")
    
    workbook.remove(Componentes)
    workbook.create_sheet(title="Componentes")
    Componentes = workbook["Componentes"]
    
    cont = 1
    for modelo, componentes in componentesLista.items():
        for componente in componentes:
            Componentes[f"A{cont}"] = modelo
            Componentes[f"B{cont}"] = componente
            cont += 1
        

    # Salvar a planilha com a nova planilha adicionada
    workbook.save(PLANILHA_SAIDA)

    print("Processo concluído com sucesso!!!")

    workbook.close()
