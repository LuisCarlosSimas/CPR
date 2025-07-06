#importações
import tkinter as tk
import salvarDados as SDD
import carregarDados as CDD
import utils

# Criar janela principal
janelaPrincipal = tk.Tk()
janelaPrincipal.title("CPR (Controle de pedidos e romaneios)")
janelaPrincipal.geometry('440x270')

#parte principal do programa! motivo de sua criação
tk.Label(janelaPrincipal, text="Pedidos e Romaneios", font=("TkDefaultFont", 10, "bold")).grid(column=0, columnspan=3, row=0, padx=5, pady=5, sticky="w")

#criar Botões
#botao controle de separação de pedidos
#salva os pedidos separados e suas informações no DB separando por dia
botaoAddPedido=tk.Button(janelaPrincipal, text='Novo pedido\nSeparado',padx=29, pady=10, anchor='center', command=lambda: SDD.addpedido(janelaPrincipal))
botaoAddPedido.grid(column=0, row=1, padx=5, pady=5)

#botao pra add pedidos no sistema com o bip
botaoAddPedidoBip=tk.Button(janelaPrincipal, text='Novo pedido\nSeparado\nBip',padx=29, pady=2, anchor='center', command=lambda: SDD.addPedidoBip(janelaPrincipal))
botaoAddPedidoBip.grid(column=1, row=1, padx=5, pady=5)

#botao tirar romaneio
#abre excel com os pedidos sendo expedidos no dia atual, e na transportadora selecionada
botaoRomaneio=tk.Button(janelaPrincipal, text='Romaneio\nDe\nEXPEDIÇÃO', command=lambda: CDD.romaneio(janelaPrincipal), padx=33, pady=2, anchor='center')
botaoRomaneio.grid(column=2, row=1, padx=5, pady=5)

#botao alterar informações do pedido
#faz qualquer alterações em qualquer pedido no DB
botaoAlteraçao = tk.Button(janelaPrincipal, text="Alterar Informações\nDe Pedidos", command=lambda: SDD.alteraInfo(janelaPrincipal), padx=11, pady=10)
botaoAlteraçao.grid(column= 0, row= 2, padx=5, pady=5)

#botao localizar pedido
#localizar se o pedido foi separado e outras informações atravez de uma busca
botaoLocalizarPedido=tk.Button(janelaPrincipal, text='Localizar\nPedido', command=lambda: CDD.acharPedido(janelaPrincipal), padx=41, pady=10, anchor='center')
botaoLocalizarPedido.grid(column=1, row=2, padx=5, pady=5)

#abrir excel com todos os pedidos separados do dia ou do mes escolhido
botaoRomaneioSeparacao = tk.Button(janelaPrincipal, text="Romaneio\nDe\nSEPARAÇÃO", command=lambda: CDD.romaneioSeparacao(janelaPrincipal), padx=31, pady=2, anchor='center')
botaoRomaneioSeparacao.grid(column=2, row=2, padx=5, pady=5)

#parte secundaria do programa! outra funções
tk.Label(janelaPrincipal, text="Entrada de Materiais e Insumos", font=("TkDefaultFont", 10, "bold")).grid(column=0, columnspan=3, row=3, padx=5, pady=5, sticky="w")

#salva materiais recebidos pela empresa em um banco de dados secundario. podendo salvar nfs apenas para registro, ou cupons fiscais e suas informaçoes.
botaoEntradaNF = tk.Button(janelaPrincipal, text="Entrada\nDe\nInsumos", command=lambda: utils.NFInsumo(janelaPrincipal), pady=2, padx=40)
botaoEntradaNF.grid(column=0, row=4, padx=5, pady=5)

#abre um excel com uma relação de gastos mensais da empresa com base em cupons fiscais, tambem mostra a lista de nfs recebidas pela empresa
botaoRelatorioGastos = tk.Button(janelaPrincipal, text= "Relatorio De\nGastos Mensais", padx=22, pady=9, command= lambda: utils.relatorioGastosMensais(janelaPrincipal))
botaoRelatorioGastos.grid(column=1, row=4, padx=5, pady=5)

utils.focus(janelaPrincipal)

# Iniciar loop da interface gráfica
janelaPrincipal.mainloop()

#aline linda amor da minha vida. ass: luis
