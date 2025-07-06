import tkinter as tk
from tkinter import messagebox
import json
import os
from datetime import datetime
import xml.etree.ElementTree as ET
import carregarDados as CDD
import utils
import re

data = datetime.now()
DATA = data.strftime("%d/%m/%Y") #data atual DD/MM/AAAA
IDATRASO = None # constante para armazenar o ID do atraso, para o bip digitar o codigo bipado
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")

#Função pra salvar o arquivo(recebe como parametro o nome do arquivo a ser salvo, e o arquivo alterado para salvar (dict)). Abre o arquivo com esse nome, e salva ele com o arquivo alterado.
def salvarArquivo(nomeArquivo=str, arquivoAlterado=dict):
    with open(nomeArquivo, 'w') as arquivo:
        json.dump(arquivoAlterado, arquivo, indent=4, ensure_ascii=False)

#Função pra salvar os clientes no banco de dados expecifico(clientes.txt). Se o DB não existe, um novo é criado
def salvarArquivoCliente(nomeArquivo=str, cliente={}):
    try:
        conteudoArquivo= CDD.carregarArquivo(nomeArquivo)
        pessoas= conteudoArquivo["clientes"]
        achou= False
        for pessoa in pessoas:
            if cliente["nome"] == pessoa["nome"]:
                achou=True
        if achou == False:
            conteudoArquivo["clientes"].append(cliente)
            salvarArquivo("clientes.txt", conteudoArquivo)
            
    except:
        conteudoArquivo= {"clientes":[]}
        conteudoArquivo['clientes'].append(cliente)
        salvarArquivo("clientes.txt", conteudoArquivo)
        messagebox.showinfo('DB clientes','Arquivo de dados inexistente! Novo arquivo de dados para clientes criado com sucesso!')

#Função pra salvar a transportadora no banco de dados(transportadoras.txt), recebe a str transportadora, ve se ja existe no banco de dados, e salva se não existir
def salvarArquivoTrans(nomeArquivo=str, arquivo=str):
    try:
        conteudoArquivo=CDD.carregarArquivo(nomeArquivo)

        if conteudoArquivo["transportadoras"] == [] or arquivo not in conteudoArquivo["transportadoras"]:
            conteudoArquivo["transportadoras"].append(arquivo)
            salvarArquivo(nomeArquivo, conteudoArquivo)

    except:
        conteudoArquivo={"transportadoras": []}
        conteudoArquivo["transportadoras"].append(arquivo)
        salvarArquivo(nomeArquivo, conteudoArquivo)
        messagebox.showinfo('DB transportadoras','Arquivo de dados inexistente! Novo arquivo de dados para transportadoras criado com sucesso!')

#função novo pedido separado
#abre uma janela pra preenchimento de informações do pedido separado e salva no DB.txt
def addpedido(janela):
    utils.janelasKill(janela=janela)
    novoDb = False
    #Tenta abrir o arquivo, se nao existir cria um novo arquivo, salva ele e depois abre.
    try:
        DB=CDD.carregarArquivo('DB.txt')
        if DATA not in DB:
            DB[DATA] = []
    except:
        DB={DATA: []}
        salvarArquivo('DB.txt', DB)
        messagebox.showinfo('Arquivo DB','Arquivo de dados inexistente! Novo arquivo de dados de separação criado com sucesso!')
        novoDb = True
        DB=CDD.carregarArquivo('DB.txt')

    #função que adiciona o pedido no DB
    def add():
        PEDIDO={"nota": inputNota.get().strip(), 
                "nome": inputNome.get().strip().lower().title(),
                "cidade/ uf": "", 
                "vol.": inputVol.get().strip(), 
                "transportadora": inputTrans.get().strip().lower().title(), 
                "separador": inputSeparador.get().strip().lower().title(),
                "refrigerado": "sim" if refrigerado.get() else "nao",
                "enviado": "nao"}
        falseEnter = True
        repitida= False

        for chave, valor in PEDIDO.items():
            if valor != "" and chave != "refrigerado" and chave != "enviado":
                falseEnter = False
        if falseEnter :
            return

        for chave, valor in PEDIDO.items():
            if chave == "refrigerado":
                break

            if valor == "" and chave != "cidade/ uf":
                messagebox.showerror('ERRO: Algum dado nao preenchido', 'Preencha todos os campos corretamente!')
                utils.focus(janelaCadastro)
                if chave == "nota":
                    inputNota.focus_set()
                if chave == "nome":
                    inputNome.focus_set()
                if chave == "vol.":
                    inputVol.focus_set()
                if chave == "transportadora":
                    inputTrans.focus_set()
                if chave == "separador":
                    inputSeparador.focus_set()
                return

            if chave == "nota" or chave == "vol.":
                try:
                    valor = int(valor)
                    valor = str(valor)
                except:
                    if chave == "nota":
                        messagebox.showerror('ERRO: Nota', 'Digite apenas numeros inteiros no campo de Nota!')
                        inputNota.delete(0, tk.END)
                        utils.focus(janelaCadastro)
                        inputNota.focus_set()
                        return
                    else:
                        messagebox.showerror('ERRO: Volume', 'Digite apenas numeros inteiros no campo de Volumes!')
                        inputVol.delete(0, tk.END)
                        utils.focus(janelaCadastro)
                        inputVol.focus_set()
                        return
        
        if novoDb == False:
            for date, listaNotas in DB.items():
                for nota in listaNotas:
                    if nota["nota"] == PEDIDO["nota"]:
                        repitida= True
                        notaRepetida= nota
                        dia= date
                        break
                if repitida == True:
                    break
                
        if repitida == True:
            messagebox.showerror("ERRO NF duplicada", f"Numero de nota ja inclusa no sistema! \nNota: {notaRepetida["nota"]} \nNome: {notaRepetida["nome"]} \nVol: {notaRepetida["vol."]} \nTransportadora: {notaRepetida["transportadora"]} \nSeparado em {dia} por {notaRepetida["separador"]}\nEnviado: {notaRepetida["enviado"]}")
            inputNota.delete(0, tk.END)
            inputNome.delete(0, tk.END)
            inputVol.delete(0, tk.END)
            inputTrans.delete(0, tk.END)
            inputSeparador.delete(0, tk.END)
            refrigerado.set(False)
            inputNota.focus_set()
            utils.focus(janelaCadastro)
            return

        DB[DATA].append(PEDIDO)
        cliente = {"nome":PEDIDO['nome']}
        salvarArquivoCliente("clientes.txt", cliente)
        salvarArquivoTrans("transportadoras.txt", PEDIDO['transportadora'])
        salvarArquivo('DB.txt', DB)  
        messagebox.showinfo('DB', f'Nota {PEDIDO["nota"]} incluida com sucesso!')
        utils.cancelar(janelaCadastro)
        addpedido(janela)
        return

    #janela de preenchimento das informações do pedido        
    janelaCadastro=tk.Toplevel(janela)
    janelaCadastro.title('Novo Pedido')

    #numero da NF
    tk.Label(janelaCadastro,text='Nota').grid(row=0,column=0,padx=5,pady=(5,0), sticky="e")
    inputNota=tk.Entry(janelaCadastro, width=30)
    inputNota.grid(row=0,column=1,padx=(0,5),pady=(5,0))
    inputNota.focus_set()

    #nome do cliente
    tk.Label(janelaCadastro,text='1º e 2º Nome').grid(row=1,column=0,padx=5, sticky="e")
    inputNome=tk.Entry(janelaCadastro, width=30)
    inputNome.grid(row=1,column=1,padx=(0,5))

    #volume de caixas
    tk.Label(janelaCadastro,text='Volumes').grid(row=2,column=0,padx=5, sticky="e")
    inputVol=tk.Entry(janelaCadastro, width=30)
    inputVol.grid(row=2,column=1,padx=(0,5))

    #transportadora
    tk.Label(janelaCadastro,text='Transportadora').grid(row=3,column=0,padx=5, sticky="e")
    inputTrans=tk.Entry(janelaCadastro, width=30)
    inputTrans.grid(row=3,column=1,padx=(0,5))

    #quem separou o pedido
    tk.Label(janelaCadastro, text="Separador").grid(row=4,column=0,padx=5, sticky="e")
    inputSeparador = tk.Entry(janelaCadastro, width=30)
    inputSeparador.grid(row=4,column=1,padx=(0,5))

    #se a mercadoria é refrigerada
    refrigerado = tk.BooleanVar()
    check_refrigerado = tk.Checkbutton(janelaCadastro, text="Refrigerado", variable=refrigerado)
    check_refrigerado.grid(row=5, column=1, padx=5, pady=5, sticky="w")

    linhaVazia=tk.Label(janelaCadastro, text='', width=40)
    linhaVazia.grid(row=6, column=0, columnspan=2)

    botaoCadastrar= tk.Button(janelaCadastro, text='Incluir Nota', padx=8, command=add, default="active")
    botaoCadastrar.grid(row=7,column=0, pady=(0,5),padx=(40,0), sticky="w")

    botaoCancelar= tk.Button(janelaCadastro, text='Cancelar', command= lambda: utils.cancelar(janelaCadastro), padx=15)
    botaoCancelar.grid(row=7,column=1, pady=(0,5),padx=(0,40), sticky="e") 
    
    janelaCadastro.bind('<Return>', lambda event=None: botaoCadastrar.invoke())

#função para adicionar pedidos no DB com o bip atravez de xmls presentes no pc pasta XMLs area de trabalho
def addPedidoBip(janela):
    utils.janelasKill(janela=janela)
    pastaXml = os.path.join(DESKTOP, "XMLs")
    if not os.path.exists(pastaXml):
        messagebox.showerror("ERRO, Pasta XMLs", "A pasta contendo os XMLs nao existe ou nao esta na Area de Trabalho!")
        return
    
    # Função para procurar um XML pela chave de acesso e salvar as informações do pedido
    def buscarXml(chave, separador):
        try:
            DB=CDD.carregarArquivo('DB.txt')
            if DATA not in DB:
                DB[DATA] = []
        except:
            DB={DATA: []}
            salvarArquivo('DB.txt', DB)
            messagebox.showinfo('Arquivo DB','Arquivo de dados inexistente! Novo arquivo de dados de separação criado com sucesso!')
            DB=CDD.carregarArquivo('DB.txt')

        print(f"\nBuscando pela chave: {chave}")
    
        for arquivo in os.listdir(pastaXml): # Percorrer todos os arquivos na pasta
            if arquivo.endswith(".xml"):  # Verifica se o arquivo é um XML
                caminhoArquivo = os.path.join(pastaXml, arquivo)
                try:
                    # Abre o arquivo XML
                    tree = ET.parse(caminhoArquivo)
                    root = tree.getroot()
                    namespaces = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
                    
                    # Busque pela chave de acesso (nfe:chNFe), considerando o namespace
                    elementoChave = root.find(".//nfe:chNFe", namespaces)
                    
                    if elementoChave is not None and elementoChave.text == chave:
                        elementoNumNF = root.find(".//nfe:nNF", namespaces) # Buscar o número da nota fiscal (nfe:nNF)
                        if elementoNumNF is not None:
                            numeroNF = elementoNumNF.text
                            numeroNF = str(numeroNF)
                        else:
                            numeroNF = ""
                            messagebox.showerror("ERRO NF", f"Arquivo ncontrado: {arquivo}, mas não foi possível localizar o número da nota.")

                        elementoNomeCliente = root.find(".//nfe:dest/nfe:xNome", namespaces)
                        if elementoNomeCliente is not None:
                            nomeCliente = elementoNomeCliente.text
                            nomeCliente = str(nomeCliente)
                        else:
                            nomeCliente = ""
                            messagebox.showerror("ERRO", f"Arquivo encontrado: {arquivo}, mas não foi possível localizar o nome do cliente.")

                        elementoCidadeCliente = root.find(".//nfe:dest/nfe:enderDest/nfe:xMun", namespaces)
                        if elementoCidadeCliente is not None:
                            cidadeCliente = elementoCidadeCliente.text
                            cidadeCliente = str(cidadeCliente)
                        else:
                            cidadeCliente = ""
                            messagebox.showerror("ERRO", f"Arquivo encontrado: {arquivo}, mas não foi possível localizar a cidade do cliente.")

                        elementoUFCliente = root.find(".//nfe:dest/nfe:enderDest/nfe:UF", namespaces)
                        if elementoUFCliente is not None:
                            UFCliente = elementoUFCliente.text
                            UFCliente = str(UFCliente)
                        else:
                            UFCliente = ""
                            messagebox.showerror("ERRO", f"Arquivo encontrado: {arquivo}, mas não foi possível localizar a UF do cliente.")

                        if cidadeCliente != "" and UFCliente != "":
                            cidadeUF= cidadeCliente + "/ " + UFCliente

                        volumeElemento = root.find(".//nfe:transp/nfe:vol/nfe:qVol", namespaces)
                        achouVol = tk.BooleanVar()
                        volumeVar = tk.StringVar()
                        if volumeElemento is not None:
                            volume = volumeElemento.text
                            volume = str(volume)
                            achouVol.set(True)
                        else:
                            achouVol.set(False)
                            #função para adicionar o volume de caixas manualmente no sistema caso ele nao esteja na nf
                            def addVol():
                                entryVol = entradaVol.get().strip()
                                if entryVol == "":
                                    return
                                else:
                                    try:
                                        entryVol = int(entryVol)
                                        volumeVar.set(str(entryVol))
                                        utils.focus(janela=Janela)
                                        Entrada.focus_set()
                                        utils.cancelar(janelaVol)
                                    except:
                                        messagebox.showerror("ERRO", "Digite apenas numeros inteiros no campo de Volumes")
                                        entradaVol.delete(0, tk.END)
                                        entradaVol.focus_set()
                                        return

                            volume = "0"
                            messagebox.showerror("ERRO", f"Arquivo encontrado: {arquivo}, mas nao foi possível localizar o volume de caixas.")

                            #janela para adicionar o volume de caixas manualmente no sistema caso ele nao esteja na NF
                            janelaVol = tk.Toplevel()
                            janelaVol.title("Adicionar Volumes")

                            tk.Label(janelaVol, text="Digite a Quantidade de Caixas").grid(column=0, row=0, padx=55, pady=(10,0))
                            entradaVol = tk.Entry(janelaVol, width=13)
                            entradaVol.grid(column=0, row=1, padx=15, pady=(0,10))

                            botaoOk = tk.Button(janelaVol, text="Salvar", padx=23, command=addVol)
                            botaoOk.grid(column=0, row=2, padx=15, pady=(0,10))

                            janelaVol.bind('<Return>', lambda event=None: botaoOk.invoke())
                            janelaVol.grab_set()
                            janelaVol.focus_set()
                            entradaVol.focus_set()
                            janelaVol.wait_window()

                        if achouVol.get() == False:
                            volume = volumeVar.get()

                        elementoTransportadora = root.find(".//nfe:transp/nfe:transporta/nfe:xNome", namespaces)
                        if elementoTransportadora is not None:
                            transportadora = elementoTransportadora.text
                            transportadora = str(transportadora).lower().title()
                        else:
                            transportadora = ""
                            messagebox.showerror("ERRO", f"Arquivo encontrado: {arquivo}, mas nao foi possivel localizar a transportadora.")

                        padrao = re.compile(r"\b(0?1\s*L|1000\s*ml|400\s*ml|30\s*L|20\s*L)\b", re.IGNORECASE)
                        refrigerado = "nao"
                        for xprod in root.findall(".//nfe:prod/nfe:xProd", namespaces):
                            texto = xprod.text
                            if texto and padrao.search(texto):
                                refrigerado = "sim"
                                break

                        enviado = "nao"

                        PEDIDO = {
                                "chave NF": chave,
                                "nota": numeroNF.strip(),
                                "nome": nomeCliente.strip(),
                                "cidade/ uf": cidadeUF.strip(),
                                "vol.": volume.strip(),
                                "transportadora":transportadora.strip(),
                                "separador": separador,
                                "refrigerado": refrigerado,
                                "enviado":enviado
                                }
                        
                        for chave, info in PEDIDO.items():
                            if info == "" and chave != "cidade/ uf":
                                messagebox.showerror("ERRO Dados inexistentes", "Dados do arquivo XML incompletos!")
                                return None
                            
                        repitida=False
                        for date, listaNotas in DB.items():
                            for nota in listaNotas:
                                if nota["nota"] == numeroNF:
                                    repitida=True
                                    notaRepetida= nota
                                    dia=date

                        if repitida:
                            messagebox.showerror("ERRO", f"Numero de nota já inclusa no sistema!\nNota: {notaRepetida["nota"]}\nNome: {notaRepetida["nome"]}\nVol: {notaRepetida["vol."]}\nTransportadora: {notaRepetida["transportadora"]}\nSeparado em {dia} por {notaRepetida["separador"]}\nEnviado: {notaRepetida["enviado"]}")
                            return None

                        else:
                            for key, value in PEDIDO.items():
                                if key != "chave NF" and key !="refrigerado" and key != "enviado":
                                    print(f"{key}: {value}")

                            if DATA not in DB:
                                DB[DATA] = []
                            DB[DATA].append(PEDIDO)

                            cliente = {"nome": PEDIDO['nome'], "cidade/ uf": PEDIDO["cidade/ uf"]}
                            salvarArquivoCliente("clientes.txt", cliente)

                            salvarArquivoTrans("transportadoras.txt", PEDIDO['transportadora'])
                            salvarArquivo('DB.txt', DB)  

                            print(f'Nota {PEDIDO["nota"]} Incluida Com Sucesso!')
                            return caminhoArquivo
                        
                except ET.ParseError as e:
                    messagebox.showerror("ERRO", f"Erro ao processar o arquivo {arquivo}! ERROR: {e}")

        messagebox.showerror("ERRO", "Nenhum XML correspondente encontrado.")            
        return None

    # Função para buscar com atraso de 1 segundo
    def buscarComAtraso(event=None):
        global IDATRASO
        if IDATRASO is not None:
            Janela.after_cancel(IDATRASO)  # Cancela o atraso anterior, se ainda estiver pendente
        IDATRASO = Janela.after(1000, buscar)  # Define o atraso de 1 segundo antes de chamar buscar

    #função para verificar os dados antes de iniciar a busca do XML
    def buscar():
        chaveAcesso = Entrada.get().strip()
        nomeSeparador = separador.get().strip()

        if chaveAcesso:
            if nomeSeparador:
                caminhoXml = buscarXml(chaveAcesso, nomeSeparador)
                if caminhoXml:
                    os.remove(caminhoXml)

                Entrada.delete(0, tk.END)
                Entrada.focus()

            else:
                messagebox.showerror("ERRO Separador", "Informe o nome do separador")
                Entrada.delete(0, tk.END)
                separador.focus()

    #janela para inicio da bipagem
    Janela = tk.Toplevel(janela)
    Janela.geometry("400x150")
    Janela.title('Novo Pedido')

    # frame para manter os widgets no centro vertical e horizontal
    frame = tk.Frame(Janela)
    frame.pack(expand=True)

    #nome do separador
    tk.Label(frame, text="Separador").pack()
    separador = tk.Entry(frame, width=30)
    separador.pack(pady=(5,0))

    #area de bipagem
    tk.Label(frame, text="Chave de acesso da nota fiscal").pack()
    Entrada = tk.Entry(frame, width=50)
    Entrada.pack()

    botaoCancelar = tk.Button(frame, text="Cancelar", padx=15, command= lambda: utils.cancelar(janela= Janela))
    botaoCancelar.pack(pady=(20,5))

    separador.focus()
    Entrada.bind("<KeyRelease>", buscarComAtraso)

#Função para alterar informações de algum pedido no banco de dados
def alteraInfo(janela):
    utils.janelasKill(janela=janela)
    try:
        DB = CDD.carregarArquivo('DB.txt')
    except:
        messagebox.showerror('ERRO','Arquivo de dados inexistente!')
        return

    def procurar(numNF):
        if numNF == "":
            return
        else:
            try:
                numNF = int(numNF)
                NUMNF = str(numNF)
            except:
                NUMNF = None
                messagebox.showerror("Erro", "Digite apenas numeros inteiros para Notas Fiscais")
                utils.focus(janelaNumNF)
                inputNota.delete(0, tk.END)
                inputNota.focus_set()
                return
        
        if NUMNF:
            DB = CDD.carregarArquivo('DB.txt')
            #função que cria uma nova NF com base nos dados alterados para fazer a alteração com a NF original no DB
            def criarNovaNF():
                NOVANF = {"data separacao": novaDataSeparacaoEntry.get().strip(),
                          "nota": novaNotaEntry.get().strip(),
                          "nome": novoNomeEntry.get().strip().lower().title(),
                          "cidade/ uf": novaCidadeEntry.get().strip() if "cidade/ uf" in NF else "",
                          "vol.": novoVolumeEntry.get().strip(),
                          "transportadora": novaTransEntry.get().strip().lower().title(),
                          "separador": novoSeparadorEntry.get().strip().lower().title(),
                          "refrigerado": "sim" if novoRefrigeradoCheck.get() else "nao",
                          "enviado": novoEnviadoEntry.get().strip() if naoEnviadoVar.get() == False else "nao"}
                
                salvar(novaNF= NOVANF, numNFOriginal=NF["nota"], dataSeparacaoOriginal= DATASEPARACAO)

            #função para excluir a NF do DB
            def excluirNF(numNota):
                confirmacao = messagebox.askyesno("Excluir", f"Tem certeza que deseja excluir a NF: {NF["nota"]}?")
                if not confirmacao:
                    utils.focus(janelaAlteracao)
                    return
                
                for data, listaPedidos in DB.items():
                    for pedido in listaPedidos:
                        if pedido["nota"] == NF["nota"]:
                            listaPedidos.remove(pedido)
                            salvarArquivo("DB.txt", DB)
                            messagebox.showinfo("NF Removida", f"NF {NF["nota"]} removida com sucesso!")
                            utils.cancelar(janelaAlteracao)
                            utils.focus(janela=janela)
                            break

            #função que salva a nova NF criada com base nos dados alterados no banco de dados
            def salvar(novaNF:dict, numNFOriginal:str, dataSeparacaoOriginal:str):
                DB = CDD.carregarArquivo('DB.txt')
                if novaNF["data separacao"] != "":
                    data = novaNF.get("data separacao", "")
                    try:
                        datetime.strptime(data, "%d/%m/%Y")
                    except:
                        messagebox.showerror("Data", "Data de separação inválida ou fora do padrão DD/MM/AAAA")
                        novaDataSeparacaoEntry.delete(0, tk.END)
                        novaDataSeparacaoEntry.focus_set()
                        return

                enterFalso= True
                for chave, conteudo in novaNF.items():
                    if not enterFalso:
                        break
                    if conteudo != "" and chave != "refrigerado" and chave != "enviado":
                        enterFalso = False
                    if chave == "refrigerado" and conteudo != NF["refrigerado"]:
                        enterFalso = False
                    if chave == "enviado" and conteudo != NF["enviado"] and conteudo != "":
                        enterFalso = False
                    if chave == "enviado" and naoEnviadoVar.get() == True and NF["enviado"] != "nao":
                        enterFalso=False
                if enterFalso:
                    return

                if novaNF["nota"] != "":
                    try:
                        nota = int(novaNF["nota"])
                        nota = str(nota)
                        for dataDB, listaPedidos in DB.items():
                            for pedido in listaPedidos:
                                if pedido["nota"] == nota:
                                    messagebox.showerror("ERRO", "Numero de NF ja incluso no sistema!")
                                    novaNotaEntry.delete(0, tk.END)
                                    novaNotaEntry.focus_set()
                                    return

                    except Exception as e:
                        messagebox.showerror("ERRO", "Digite apenas numeros inteiros no campo de nota.")
                        print(e)
                        novaNotaEntry.delete(0, tk.END)
                        novaNotaEntry.focus_set()
                        return
                    
                if novaNF["cidade/ uf"] != "":
                    cidadeOk = False
                    cidade = novaNF["cidade/ uf"]
                    if '/' in cidade:
                        partes = cidade.split('/')
                        if len(partes) == 2:
                            nomeCidade = partes[0].strip().title()
                            uf = partes[1].strip().upper()
                            if len(uf) == 2:
                                cidadeOk = True
                                cidade = f"{nomeCidade}/ {uf}"

                    if cidadeOk == False:
                        messagebox.showerror("ERRO", "Formato inválido. Formato esperado 'Cidade/ UF'")
                        novaCidadeEntry.delete(0, tk.END)
                        novaCidadeEntry.focus_set()
                        return
                    
                if novaNF["vol."] != "":
                    try:
                        vol = int(novaNF["vol."])
                    except:
                        messagebox.showerror("ERRO", "Digite apenas numeros inteiros no campo de Volumes.")
                        novoVolumeEntry.delete(0, tk.END)
                        novoVolumeEntry.focus_set()
                        return

                if novaNF["enviado"] != "nao" and novaNF["enviado"] != "":
                    data = novaNF.get("enviado", "")
                    try:
                        datetime.strptime(data, "%d/%m/%Y")
                    except:
                        messagebox.showerror("Data", "Data de envio inválida ou fora do padrão DD/MM/AAAA")
                        naoEnviadoVar.set(False)
                        novoEnviadoEntry.delete(0, tk.END)
                        novoEnviadoEntry.focus_set()
                        return
                
                for chave, valor in novaNF.items():
                    if chave != "data separacao" and valor != "":
                        NF[chave] = valor

                if novaNF["data separacao"] != "":
                    novaData = novaNF["data separacao"]

                else:
                    novaData = dataSeparacaoOriginal

                if novaData not in DB:
                    DB[novaData] = []

                for dataDB, listaPedidos in DB.items():
                    if dataDB == dataSeparacaoOriginal:
                        for pedido in listaPedidos:
                            if pedido["nota"] == numNFOriginal:
                                listaPedidos.remove(pedido)
                                break
                
                DB[novaData].append(NF)
                salvarArquivo("DB.txt", DB)
                messagebox.showinfo("DB", "NF alterada com sucesso")
                utils.cancelar(janelaAlteracao)
                procurar(NF["nota"])
                return

            for data, listaPedidos in DB.items():
                for pedido in listaPedidos:
                    if pedido["nota"] == NUMNF:
                        DATASEPARACAO = data
                        NF = pedido
                        break
                    else:
                        NF = None
                if NF:
                    break
        
            if NF == None:
                messagebox.showinfo("DB", "Número de Nota Fiscal não encontrado no DB")
                utils.focus(janelaNumNF)
                inputNota.delete(0, tk.END)
                inputNota.focus_set()
                return

            #janela onde vai ser feita as alterações na nota fiscal
            janelaAlteracao = tk.Toplevel(janela)
            janelaAlteracao.title("Alteração de NF")

            tk.Label(janelaAlteracao, text="DB").grid(column=0, row=0, padx=10, pady=(5,10))

            tk.Label(janelaAlteracao, text=" "*10).grid(column=1, row=0)

            tk.Label(janelaAlteracao, text="Alterações").grid(column=3, row=0, pady=10)
            
            LINHA = 1

            tk.Label(janelaAlteracao, text= f"Nota: {NF["nota"]}").grid(column=0, row=LINHA, padx=10)

            tk.Label(janelaAlteracao, text= "Nota:").grid(column=2, row=LINHA, padx=(10,5), sticky="e")
            novaNotaEntry = tk.Entry(janelaAlteracao, width=30)
            novaNotaEntry.grid(column=3, row=LINHA, padx=(0,10))

            LINHA += 1

            tk.Label(janelaAlteracao, text= f"Nome: {NF["nome"]}").grid(column=0, row=LINHA, padx=10)

            tk.Label(janelaAlteracao, text= "Nome:").grid(column=2, row=LINHA, padx=(10,5), sticky="e")
            novoNomeEntry = tk.Entry(janelaAlteracao, width=30)
            novoNomeEntry.grid(column=3, row=LINHA, padx=(0,10))

            LINHA += 1

            if "cidade/ uf" in NF:
                tk.Label(janelaAlteracao, text= f"Cidade/ UF: {NF["cidade/ uf"]}").grid(column=0, row=LINHA, padx=10)
                
                tk.Label(janelaAlteracao, text= "Cidade/ UF:").grid(column=2, row=LINHA, padx=(10,5), sticky="e")
                novaCidadeEntry = tk.Entry(janelaAlteracao, width=30)
                novaCidadeEntry.grid(column=3, row=LINHA, padx=(0,10))

                LINHA += 1
            
            tk.Label(janelaAlteracao, text= f"Volumes: {NF["vol."]}").grid(column=0, row=LINHA, padx=10)

            tk.Label(janelaAlteracao, text= "Volumes:").grid(column=2, row=LINHA, padx=(10,5), sticky="e")
            novoVolumeEntry = tk.Entry(janelaAlteracao, width=30)
            novoVolumeEntry.grid(column=3, row=LINHA, padx=(0,10))

            LINHA += 1

            tk.Label(janelaAlteracao, text= f"Transportadora: {NF["transportadora"]}").grid(column=0, row=LINHA, padx=10)

            tk.Label(janelaAlteracao, text= "Transportadora:").grid(column=2, row=LINHA, padx=(10,5), sticky="e")
            novaTransEntry = tk.Entry(janelaAlteracao, width=30)
            novaTransEntry.grid(column=3, row=LINHA, padx=(0,10))

            LINHA += 1

            tk.Label(janelaAlteracao, text= f"Separador: {NF["separador"]}").grid(column=0, row=LINHA, padx=10)

            tk.Label(janelaAlteracao, text= "Separador:").grid(column=2, row=LINHA, padx=(10,5), sticky="e")
            novoSeparadorEntry = tk.Entry(janelaAlteracao, width=30)
            novoSeparadorEntry.grid(column=3, row=LINHA, padx=(0,10))

            LINHA += 1

            tk.Label(janelaAlteracao, text= f"Refrigerado: {NF["refrigerado"]}").grid(column=0, row=LINHA, padx=10)

            tk.Label(janelaAlteracao, text= "Refrigerado:").grid(column=2, row=LINHA, padx=(10,5), sticky="e")
            novoRefrigeradoCheck = tk.BooleanVar()
            checkRefrigeradas = tk.Checkbutton(janelaAlteracao, variable=novoRefrigeradoCheck)
            checkRefrigeradas.grid(column=3, row=LINHA, padx=(0,10), sticky="w")
            if NF["refrigerado"] == "sim":
                novoRefrigeradoCheck.set(True)

            LINHA += 1

            tk.Label(janelaAlteracao, text= f"Separado: {DATASEPARACAO}").grid(column=0, row=LINHA, padx=10)

            tk.Label(janelaAlteracao, text= "Separado:").grid(column=2, row=LINHA, padx=(10,5), sticky="e")
            novaDataSeparacaoEntry = tk.Entry(janelaAlteracao, width=30)
            novaDataSeparacaoEntry.grid(column=3, row=LINHA, padx=(0,10))

            LINHA += 1

            tk.Label(janelaAlteracao, text= f"Enviado: {NF["enviado"]}").grid(column=0, row=LINHA, padx=10)

            tk.Label(janelaAlteracao, text= "Enviado:").grid(column=2, row=LINHA, padx=(10,5), sticky="e")
            novoEnviadoEntry = tk.Entry(janelaAlteracao, width=20)
            novoEnviadoEntry.grid(column=3, row=LINHA, padx=(0,10), sticky="w")

            naoEnviadoVar = tk.BooleanVar()
            naoEnviadoCheck = tk.Checkbutton(janelaAlteracao, text="Não", variable=naoEnviadoVar)
            naoEnviadoCheck.grid(column=3, row=LINHA, padx=(0,10), sticky="e")
            if NF["enviado"] == "nao":
                naoEnviadoVar.set(True)

            LINHA += 1
            
            botaoSalvar=tk.Button(janelaAlteracao, text="Salvar", command= lambda: criarNovaNF(), padx= 23)
            botaoSalvar.grid(column=0, columnspan=4, row=LINHA, padx=10, pady=(20,10), sticky="e")

            botaoExcluir = tk.Button(janelaAlteracao, text="Excluir NF", command= lambda: excluirNF(NF["nota"]), padx=13)
            botaoExcluir.grid(column=0, columnspan=4, row=LINHA, pady=(20,10))

            botaoCancelar = tk.Button(janelaAlteracao, text="Cancelar", command= lambda: utils.cancelar(janelaAlteracao), padx=15) 
            botaoCancelar.grid(column=0, columnspan=4, row=LINHA, padx=10, pady=(20,10), sticky="w")

            janelaAlteracao.bind('<Return>', lambda event=None: botaoSalvar.invoke())

            utils.cancelar(janelaNumNF)

    #janela de busca da nota fiscal        
    janelaNumNF = tk.Toplevel(janela)
    janelaNumNF.title("Alteração de Pedidos")

    tk.Label(janelaNumNF, text="Digite o número da Nota Fiscal: ").grid(column=0, row=0, padx=(10,5), pady=(10,5))
    inputNota = tk.Entry(janelaNumNF, width=20)
    inputNota.grid(column=1, row=0, padx=(5,10), pady=(10,5))
    inputNota.focus()

    botaoProcurar = tk.Button(janelaNumNF, text="Procurar", command=lambda: procurar(numNF=str(inputNota.get().strip())), padx=15)
    botaoProcurar.grid(column= 0, row=1, rowspan=2, padx=10, pady=(5,10))

    botaoCancelar = tk.Button(janelaNumNF, text="Cancelar", command= lambda: utils.cancelar(janelaNumNF), padx=15)
    botaoCancelar.grid(column=1, row=1, rowspan=2, padx=10, pady=(5,10))

    janelaNumNF.bind('<Return>', lambda event=None: botaoProcurar.invoke())
