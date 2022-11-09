import win32com.client as win32
from datetime import datetime
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import os

# Definindo horario atual
hora = datetime.now()
# Somente o horario será coletado
horan = int(hora.hour)
saudacao = str()

# Definindo saudação conforme o horario do dia para início do e-mail.
if horan >= int('1') and horan < int('12'):
    saudacao = 'bom dia!'

elif horan >= int('12') and horan < int('18'):
    saudacao = 'boa tarde!'

else:
    saudacao = 'boa noite!'


def btNF():
    NF_envio.geometry('350x470')
    BTNF.destroy()
    BTCAIXINHA.destroy()
    BTCARTAO.destroy()
    tcartao.destroy()
    cartao.destroy()
    textselect.destroy()
    bt_envioNF = Button(NF_envio, text='Enviar e-mail', command=envio_NF)
    bt_envioNF.grid(column=0, row=16, columnspan=3, pady=20)


def btCAIXINHA():
    NF_envio.geometry('350x420')
    BTNF.destroy()
    BTCAIXINHA.destroy()
    BTCARTAO.destroy()
    tvencimento.destroy()
    vencimento.destroy()
    tcartao.destroy()
    cartao.destroy()
    textselect.destroy()
    tboleto.destroy()
    menuBoleto.destroy()
    bt_envioCAIXINHA = Button(NF_envio, text='Enviar e-mail', command=envio_Caixinha)
    bt_envioCAIXINHA.grid(column=0, row=16, columnspan=3, pady=20)


def btCARTAO():
    NF_envio.geometry('350x450')
    BTCARTAO.destroy()
    BTNF.destroy()
    BTCARTAO.destroy()
    BTCAIXINHA.destroy()
    tvencimento.destroy()
    vencimento.destroy()
    textselect.destroy()
    tboleto.destroy()
    menuBoleto.destroy()
    bt_envioCARTAO = Button(NF_envio, text='Enviar e-mail', command=envio_Cartao)
    bt_envioCARTAO.grid(column=0, row=16, columnspan=3, pady=20)


def envio_Cartao():
    if pedido.get() == '' or ano.get() == '' or mes.get() == '' or finalidade.get() == '' or projeto.get() == '' or setor.get() == '' or subsetor.get() == '' or nome.get() == '':
        messagebox.showwarning(title='Envio de pedido de compras.', message='Todos os campos com * são obrigatórios!')

    else:

        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        email.To = 'financas@porschegt3cup.com.br'
        email.CC = f'enzo@porschegt3cup.com.br;silvana@porschegt3cup.com.br;daniel@porschegt3cup.com.br;{emailcc.get()}@porschegt3cup.com.br'
        email.Subject = f'COMPRA NO CARTÃO DE CRÉDITO - FINAL {cartao.get()}'
        email.Body = f'''Prezados, {saudacao}

Segue abaixo informações da nota fiscal anexa: 
Nome comprador: MARCEL GATTI
Finalidade: {finalidade.get()}
Projeto: {projeto.get().upper()}
Setor: {setor.get().upper()}
Subsetor: {subsetor.get().upper()}
Pedido No: PEDIDO {pedido.get()}

{observacao.get()}

Att.
{nome.get().title()}
{('_' * 20)} 

        '''

        anexo1 = 'W:\\Compras de peças - REGISTROS_NFs_CONTROLE\\NOTAS FISCAIS\\{}\\{}\\NF pedido {}.pdf'.format(
            ano.get(), mes.get().upper(), pedido.get())
        if not os.path.exists(anexo1):
            messagebox.showwarning(title='Envio de pedido de compras.', message='''Nota fiscal não encotrada.
Verifiquei o n° do pedido, mês, ano ou se o arquivo está salvo na pasta.''')

        else:
            email.Attachments.Add(anexo1)

            if (finalidade.get().upper() == 'ANEXO') or (projeto.get().upper() == 'ANEXO') or (
                    setor.get().upper() == 'ANEXO') or (subsetor.get().upper() == 'ANEXO'):
                anexo2 = 'W:\\Compras de peças - REGISTROS_NFs_CONTROLE\\NOTAS FISCAIS\\{}\\{}\\pedido {}.pdf'.format(
                    ano.get(), mes.get().upper(), pedido.get())
                if not os.path.exists(anexo2):
                    messagebox.showwarning(title='Envio de pedido de compras.', message='''Anexo não encontrado.
Verifiquei o n° do pedido, mês, ano ou se o arquivo está salvo na pasta.''')

                else:
                    email.Attachments.Add(anexo2)

                    email.Send()

                    enviado = Label(NF_envio, text='E-mail enviado!', fg='blue')
                    enviado.grid(column=1, row=13, columnspan=3)

                    botao2 = Button(NF_envio, text='Sair', command=fechar, padx=30)
                    botao2.grid(column=1, row=16, columnspan=3, pady=20)

                    pedido.destroy(), ano.destroy(), menuMes.destroy(), finalidade.destroy(), projeto.destroy(), setor.destroy(), subsetor.destroy(), vencimento.destroy(),
                    tpedido.destroy(), tano.destroy(), tmes.destroy(), tfinalidade.destroy(), tprojeto.destroy(), tsetor.destroy(), tsubsetor.destroy(),
                    tvencimento.destroy(), tcartao.destroy(), cartao.destroy(), textinicial.destroy(), tnome.destroy(), nome.destroy(), temailcc.destroy(), emailcc.destroy(),
                    tobservacao.destroy(), observacao.destroy()

            else:
                email.Send()

                enviado = Label(NF_envio, text='E-mail enviado!', fg='blue')
                enviado.grid(column=1, row=13, columnspan=3)

                botao2 = Button(NF_envio, text='Sair', command=fechar, padx=30)
                botao2.grid(column=1, row=16, columnspan=3, pady=20)

                pedido.destroy(), ano.destroy(), menuMes.destroy(), finalidade.destroy(), projeto.destroy(), setor.destroy(), subsetor.destroy(), vencimento.destroy(),
                tpedido.destroy(), tano.destroy(), tmes.destroy(), tfinalidade.destroy(), tprojeto.destroy(), tsetor.destroy(), tsubsetor.destroy(),
                tvencimento.destroy(), tcartao.destroy(), cartao.destroy(), textinicial.destroy(), tnome.destroy(), nome.destroy(), temailcc.destroy(), emailcc.destroy(),
                tobservacao.destroy(), observacao.destroy()


def envio_Caixinha():  # Função para envio de e-mail com nota fiscal a ser paga.

    if pedido.get() == '' or ano.get() == '' or mes.get() == '' or finalidade.get() == '' or projeto.get() == '' or setor.get() == '' or subsetor.get() == '' or nome.get() == '':
        messagebox.showwarning(title='Envio de pedido de compras.', message='Todos os campos com * são obrigatórios!')

    else:

        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        email.To = 'financas@porschegt3cup.com.br'
        email.CC = f'enzo@porschegt3cup.com.br;silvana@porschegt3cup.com.br;{emailcc.get()}@porschegt3cup.com.br'
        email.Subject = 'PEDIDO {}'.format(pedido.get())
        email.Body = f'''Prezados, {saudacao}

Segue abaixo informações da nota fiscal anexa: PAGAMENTOS REALIZADOS PELO MOTOBOY OU MOTORISTAS, PEDIDO FEITO NA CAIXINHA.
Nome comprador: MARCEL GATTI
Finalidade: {finalidade.get()}
Projeto: {projeto.get().upper()}
Setor: {setor.get().upper()}
Subsetor: {subsetor.get().upper()}
Pedido No: PEDIDO {pedido.get()}

{observacao.get()}

Att.
{nome.get().title()}
{('_' * 20)} 

    '''

        anexo1 = 'W:\\Compras de peças - REGISTROS_NFs_CONTROLE\\NOTAS FISCAIS\\{}\\{}\\NF pedido {}.pdf'.format(
            ano.get(), mes.get().upper(), pedido.get())
        if not os.path.exists(anexo1):
            messagebox.showwarning(title='Envio de pedido de compras.', message='''Nota fiscal não encontrada.
Verifiquei o n° do pedido, mês, ano ou se o arquivo está salvo na pasta.''')

        else:
            email.Attachments.Add(anexo1)

            if (finalidade.get().upper() == 'ANEXO') or (projeto.get().upper() == 'ANEXO') or (
                    setor.get().upper() == 'ANEXO') or (subsetor.get().upper() == 'ANEXO'):
                anexo2 = 'W:\\Compras de peças - REGISTROS_NFs_CONTROLE\\NOTAS FISCAIS\\{}\\{}\\pedido {}.pdf'.format(
                    ano.get(), mes.get().upper(), pedido.get())
                if not os.path.exists(anexo2):
                    messagebox.showwarning(title='Envio de pedido de compras', message='''Anexo não encontrado.
Verifiquei o n° do pedido, mês, ano ou se o arquivo está salvo na pasta.''')

                else:
                    email.Attachments.Add(anexo2)

                    email.Send()

                    enviado = Label(NF_envio, text='E-mail enviado!', fg='blue')
                    enviado.grid(column=0, row=11, columnspan=3)

                    botao2 = Button(NF_envio, text='Sair', command=fechar, padx=30)
                    botao2.grid(column=0, row=16, columnspan=3, pady=20)

                    pedido.destroy(), ano.destroy(), menuMes.destroy(), finalidade.destroy(), projeto.destroy(), setor.destroy(), subsetor.destroy(), vencimento.destroy(),
                    tpedido.destroy(), tano.destroy(), tmes.destroy(), tfinalidade.destroy(), tprojeto.destroy(), tsetor.destroy(), tsubsetor.destroy(),
                    tvencimento.destroy(), textinicial.destroy(), tnome.destroy(), nome.destroy(), temailcc.destroy(), emailcc.destroy(),tobservacao.destroy(), observacao.destroy()
            else:
                email.Send()

                enviado = Label(NF_envio, text='E-mail enviado!', fg='blue')
                enviado.grid(column=0, row=11, columnspan=3)

                botao2 = Button(NF_envio, text='Sair', command=fechar, padx=30)
                botao2.grid(column=0, row=16, columnspan=3, pady=20)

                pedido.destroy(), ano.destroy(), menuMes.destroy(), finalidade.destroy(), projeto.destroy(), setor.destroy(), subsetor.destroy(), vencimento.destroy(),
                tpedido.destroy(), tano.destroy(), tmes.destroy(), tfinalidade.destroy(), tprojeto.destroy(), tsetor.destroy(), tsubsetor.destroy(),
                tvencimento.destroy(), textinicial.destroy(), tnome.destroy(), nome.destroy(), temailcc.destroy(), emailcc.destroy(),tobservacao.destroy(), observacao.destroy()


def envio_NF():  # Função para envio de e-mail com nota fiscal a ser paga.

    if pedido.get() == '' or ano.get() == '' or mes.get() == '' or finalidade.get() == '' or projeto.get() == '' or setor.get() == '' or subsetor.get() == '' or vencimento.get() == '' or vlista.get() == '' or nome.get() == '':
        messagebox.showwarning(title='Envio de pedido de compras.', message='Todos os campos com * são obrigatórios!')

    else:

        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        email.To = 'financas@porschegt3cup.com.br'
        email.CC = f'enzo@porschegt3cup.com.br;silvana@porschegt3cup.com.br;{emailcc.get()}@porschegt3cup.com.br'
        email.Subject = 'PEDIDO {}'.format(pedido.get())
        email.Body = f'''Prezados, {saudacao}

Segue abaixo informações da nota fiscal anexa: 
Nome comprador: MARCEL GATTI
Finalidade: {finalidade.get()}
Projeto: {projeto.get().upper()}
Setor: {setor.get().upper()}
Subsetor: {subsetor.get().upper()}
Pedido No: PEDIDO {pedido.get()}
Vencimento: {vencimento.get()}

{observacao.get()}

Att.
{nome.get().title()}
{('_' * 20)} 

    '''

        anexo1 = 'W:\\Compras de peças - REGISTROS_NFs_CONTROLE\\NOTAS FISCAIS\\{}\\{}\\NF pedido {}.pdf'.format(
            ano.get(), mes.get().upper(), pedido.get())
        if not os.path.exists(anexo1):
            messagebox.showwarning(title='Envio de pedido de compras.', message='''Nota fiscal não encontrada.
Verifiquei o n° do pedido, mês, ano ou se o arquivo está salvo na pasta.''')

        else:
            email.Attachments.Add(anexo1)

            if (finalidade.get().upper() == 'ANEXO') or (projeto.get().upper() == 'ANEXO') or (
                    setor.get().upper() == 'ANEXO') or (subsetor.get().upper() == 'ANEXO'):
                anexo2 = 'W:\\Compras de peças - REGISTROS_NFs_CONTROLE\\NOTAS FISCAIS\\{}\\{}\\pedido {}.pdf'.format(
                    ano.get(), mes.get().upper(), pedido.get())
                if not os.path.exists(anexo2):
                    messagebox.showwarning(title='Envio de pedido de compras.', message='''Anexo não encontrado.
Verifiquei o n° do pedido, mês, ano ou se o arquivo está salvo na pasta.''')

                else:
                    email.Attachments.Add(anexo2)

                    if vlista.get() == 'SIM':
                        anexo3 = 'W:\\Compras de peças - REGISTROS_NFs_CONTROLE\\NOTAS FISCAIS\\{}\\{}\\BOLETO pedido {}.pdf'.format(
                            ano.get(), mes.get().upper(), pedido.get())
                        if not os.path.exists(anexo3):
                            messagebox.showwarning(title='Envio de pedido de compras.', message='''Boleto não encontrado.
Verifiquei o n° do pedido, mês, ano ou se o arquivo está salvo na pasta.''')

                        else:
                            email.Attachments.Add(anexo3)

                            email.Send()

                            enviado = Label(NF_envio, text='E-mail enviado!', fg='blue')
                            enviado.grid(column=0, row=12, columnspan=3)

                            botao2 = Button(NF_envio, text='Sair', command=fechar, padx=30)
                            botao2.grid(column=0, row=16, columnspan=3, pady=20)

                            pedido.destroy(), ano.destroy(), menuMes.destroy(), finalidade.destroy(), projeto.destroy(), setor.destroy(), subsetor.destroy(), vencimento.destroy(),
                            tpedido.destroy(), tano.destroy(), tmes.destroy(), tfinalidade.destroy(), tprojeto.destroy(), tsetor.destroy(), tsubsetor.destroy(),
                            tvencimento.destroy(), textinicial.destroy(), tboleto.destroy(), menuBoleto.destroy(), tnome.destroy(), nome.destroy(), temailcc.destroy(), emailcc.destroy(),
                            tobservacao.destroy(), observacao.destroy()

                    else:
                        email.Send()

                        enviado = Label(NF_envio, text='E-mail enviado!', fg='blue')
                        enviado.grid(column=0, row=12, columnspan=3)

                        botao2 = Button(NF_envio, text='Sair', command=fechar, padx=30)
                        botao2.grid(column=0, row=16, columnspan=3, pady=20)

                        pedido.destroy(), ano.destroy(), menuMes.destroy(), finalidade.destroy(), projeto.destroy(), setor.destroy(), subsetor.destroy(), vencimento.destroy(),
                        tpedido.destroy(), tano.destroy(), tmes.destroy(), tfinalidade.destroy(), tprojeto.destroy(), tsetor.destroy(), tsubsetor.destroy(),
                        tvencimento.destroy(), textinicial.destroy(), tboleto.destroy(), menuBoleto.destroy(), tnome.destroy(), nome.destroy(), temailcc.destroy(), emailcc.destroy(),
                        tobservacao.destroy(), observacao.destroy()

            else:
                if vlista.get() == 'SIM':
                    anexo3 = 'W:\\Compras de peças - REGISTROS_NFs_CONTROLE\\NOTAS FISCAIS\\{}\\{}\\BOLETO pedido {}.pdf'.format(
                        ano.get(), mes.get().upper(), pedido.get())
                    if not os.path.exists(anexo3):
                        messagebox.showwarning(title='Envio de pedido de compras.', message='''Boleto não encontrado.
Verifique n° do pedido, mês, ano ou se o arquivo está salvo na pasta.''')

                    else:
                        email.Attachments.Add(anexo3)

                        email.Send()

                        enviado = Label(NF_envio, text='E-mail enviado!', fg='blue')
                        enviado.grid(column=0, row=12, columnspan=3)

                        botao2 = Button(NF_envio, text='Sair', command=fechar, padx=30)
                        botao2.grid(column=0, row=16, columnspan=3, pady=20)

                        pedido.destroy(), ano.destroy(), menuMes.destroy(), finalidade.destroy(), projeto.destroy(), setor.destroy(), subsetor.destroy(), vencimento.destroy(),
                        tpedido.destroy(), tano.destroy(), tmes.destroy(), tfinalidade.destroy(), tprojeto.destroy(), tsetor.destroy(), tsubsetor.destroy(),
                        tvencimento.destroy(), textinicial.destroy(), tboleto.destroy(), menuBoleto.destroy(), tnome.destroy(), nome.destroy(), temailcc.destroy(), emailcc.destroy(),
                        tobservacao.destroy(), observacao.destroy()

                else:
                    email.Send()

                    enviado = Label(NF_envio, text='E-mail enviado!', fg='blue')
                    enviado.grid(column=0, row=12, columnspan=3)

                    botao2 = Button(NF_envio, text='Sair', command=fechar, padx=30)
                    botao2.grid(column=0, row=16, columnspan=3, pady=20)

                    pedido.destroy(), ano.destroy(), menuMes.destroy(), finalidade.destroy(), projeto.destroy(), setor.destroy(), subsetor.destroy(), vencimento.destroy(),
                    tpedido.destroy(), tano.destroy(), tmes.destroy(), tfinalidade.destroy(), tprojeto.destroy(), tsetor.destroy(), tsubsetor.destroy(),
                    tvencimento.destroy(), textinicial.destroy(), tboleto.destroy(), menuBoleto.destroy(), tnome.destroy(), nome.destroy(), temailcc.destroy(), emailcc.destroy(),
                    tobservacao.destroy(), observacao.destroy()


def fechar():
    NF_envio.quit()


NF_envio = Tk()
NF_envio.title('Envio de pedido de compras.')
NF_envio.geometry('450x110')

textinicial = Label(NF_envio, text='Insira os dados para envio do e-mail.')
textinicial.grid(column=0, row=0, columnspan=3, padx=30, pady=10)

textselect = Label(NF_envio, text='Selecione o formato de e-mail.', fg='red')
textselect.grid(column=0, row=1, columnspan=3)

BTNF = Button(NF_envio, text='Pedido NF', command=btNF, padx=20)
BTNF.grid(column=0, row=2, pady=10)

BTCAIXINHA = Button(NF_envio, text='Pedido Caixinha', command=btCAIXINHA)
BTCAIXINHA.grid(column=1, row=2, pady=10)

BTCARTAO = Button(NF_envio, text='Pedido Cartão', command=btCARTAO, padx=10)
BTCARTAO.grid(column=2, row=2, pady=10, padx=35, sticky=E)

temailcc = Label(NF_envio, text='E-mail adicional para copia:')
temailcc.grid(column=0, row=3, padx=10, pady=5, sticky=E)
emailcc = Entry(NF_envio)
emailcc.grid(column=1, row=3, columnspan=2, sticky=EW)

tpedido = Label(NF_envio, text='N° do pedido:*')
tpedido.grid(column=0, row=4, padx=10, pady=5, sticky=E)
pedido = Entry(NF_envio)
pedido.grid(column=1, row=4, columnspan=2, sticky=EW)

tano = Label(NF_envio, text='Ano do pedido:*')
tano.grid(column=0, row=5, padx=10, pady=5, sticky=E)
ano = Entry(NF_envio)
ano.grid(column=1, row=5, columnspan=2, sticky=EW)

tmes = Label(NF_envio, text='Mês do pedido:*')
tmes.grid(column=0, row=6, padx=10, pady=5, sticky=E)
listames = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']
mes = StringVar()
mes.set('')
menuMes = OptionMenu(NF_envio, mes, *listames)
menuMes.grid(column=1, row=6, columnspan=2, sticky=EW)

'''mes = Entry(NF_envio)
mes.grid(column=1, row=5, columnspan=2, sticky=EW)'''

tfinalidade = Label(NF_envio, text='Finalidade:*')
tfinalidade.grid(column=0, row=7, padx=10, pady=5, sticky=E)
listafinalidade = ['ANEXO', '123010001 TERRENOS',	'123010002 EDIFICIOS',	'123010003 CONSTRUCOES',	'123020001 MOVEIS E UTENSILIOS',	'123020002 FERRAMENTAS',	'123020003 MAQUINAS E EQUIPAMENTOS',	'123020004 COMPUTADORES E PERIFERICOS',	'123020005 APARELHOS ELETRO-ELETRONICOS',	'123020006 INSTALACOES',	'123020007 AUTOMOVEIS',	'123020010 DIREITO DE USO DE LINHA TELEFONICA',	'123020011 BENFEITORIA EM IMOVEIS',	'123020012 CONTEINERS',	'123020016 QUOTA DE CONSORCIO',	'123020030 SOFTWARE',	'123070001 (-) DEPRECIACOES DE MOVEIS E UTENSILIOS',	'123070002 (-) DEPRECIACOES DE FERRAMENTAS',	'123070003 (-) DEPRECIACOES DE MAQUINAS, EQUIP. FER',	'123070004 (-) DEPRECIACOES DE COMPUTADORES E PERIF',	'123070005 (-) DEPRECIACOES DE APARELHOS ELETRO-ELE',	'123070006 (-) DEPRECIACOES DE INSTALACOES',	'123070007 (-) DEPRECIACOES DE AUTOMOVEIS',	'123070010 (-) DEPRECIACOES DE DIREITO DE USO DE LI',	'123070011 (-) DEPRECIACOES DE BENFEITORIA EM IMOVE',	'123070012 (-) DEPRECIACOES DE CONTEINERS',	'123070030 (-) AMORTIZACOES DE SOFTAWARE',	'241010002 DENER JORGE PIRES',	'332200001 COMBUSTIVEIS E LUBRIFICANTES CARROS DE C',	'332200002 MANUT CARROS CORRIDA',	'332200003 PNEUS CARROS DE CORRIDA',	'332200004 NITROGENIO/GASES CARROS DE CORRIDA',	'332200005 MATERIAIS CARROS CORRIDA',	'332200006 ADESIVAGEM CARROS DE CORRIDA',	'332200090 OUTROS GASTOS COM CARROS DE CORRIDA',	'332210001 BUFFET VIP',	'332210002 MONTAGEM',	'332210003 LOCACAO DE AUTODROMO',	'332210004 MODELOS E RECEPCIONISTA',	'332210005 CREDENCIAIS/INGRESSO',	'332210090 OUTROS GASTOS DE INFRA-ESTRUTURA DE EVEN',	'332220001 FILMAGEM/FOTOGRAFIA',	'332220002 LOCUTOR ESPORTISTA/COMENTARISTAS',	'332220003 SERVICOS DE COMUNICACAO - ETAPAS',	'332220004 MIDIAS SOCIAIS',	'332220005 PUBLICACOES PUBLICIDADES E PROPAGANDAS',	'332220006 TROFEUS',	'332220007 BRINDES - ETAPAS',	'332220008 COMUNICACAO VISUAL',	'332220090 OUTROS GASTOS DE MARKETING E MIDIA',	'332230001 AUTONOMOS - ETAPAS',	'332230002 SERVICOS TECNICOS ESPECIALIZADOS - ETAPA',	'332230003 RESGATES - ETAPAS',	'332230004 COMISSARIO/CBA/FEDERACAO',	'332230005 ASSESSORIA/CONSULTOR - ETAPAS',	'332230006 MATERIAL E SERVICOS DE LIMPEZA - ETAPAS',	'332230007 SEGURANCA E VIGILANCIA - ETAPAS',	'332230008 EQUIP E FERRAMENTAS - ETAPAS',	'332230009 LAVANDERIA - ETAPAS',	'332230010 INFORMATICA - ETAPAS',	'332230011 SERVICOS MEDICOS -ETAPAS',	'332230090 OUTROS SERVICOS TECNICOS E DE APOIO - ET',	'332240001 HOSPEDAGEM - ETAPAS',	'332240003 ALIMENTACAO - STAFF - ETAPAS',	'332240004 FRETE/ONIBUS - ETAPAS',	'332240005 PASSAGENS AEREAS',	'332240006 TRANSPORTE - ETAPAS',	'332240007 TAXI - ETAPAS',	'332240090 OUTROS GASTOS DE TRANSPORTE, ALIMENTACAO',	'332250001 DESPACH. ADUANEIRO - ETAPAS',	'332250002 FRETES E CARRETOS - ETAPAS',	'332250003 ESTACIONAMENTO - ETAPAS',	'332250090 OUTROS GASTOS DE LOGISTICA - ETAPAS',	'332260001 MATERIAL DE CONSUMO - ETAPAS',	'332260002 SERVICOS DE ENTREGA - ETAPAS',	'332260003 SEGUROS - ETAPAS',	'332260004 UNIFORMES - ETAPAS',	'332260090 OUTROS GASTOS GERAIS - ETAPAS',	'382010001 SALARIOS',	'382010002 PRO-LABORE',	'382010004 FERIAS',	'382010005 13.SALARIO',	'382010010 VALE TRANSPORTE',	'382010011 VALE ALIMENTACAO/REFEICAO',	'382010012 VALE COMBUSTIVEL',	'382010015 ASSISTENCIA MEDICA / ODONTOLOGICA',	'382010016 ACADEMIA',	'382010017 AUTONOMOS - SEDE',	'382010018 ESTAGIARIOS',	'382010019 CONFRATERNIZACOES',	'382010020 TREINAMENTOS',	'382010021 FARMACIA',	'382010022 RESCISOES E INDENIZACOES',	'382019999 OUTROS GASTOS COM PESSOAL',	'382020001 I.N.S.S.',	'382020002 F.G.T.S.',	'382020003 MULTA RESCISORIA FGTS',	'382030013 DEPRECIACOES E AMORTIZACOES',	'382030034 SEGURO - SEDE',	'382030035 REFEICOES, MANTIMENTOS E BEBIDAS - SEDE',	'382030040 MATERIAL DE CONSUMO - SEDE',	'382030041 UNIFORMES - SEDE',	'382030069 EQUIP. E FERRAMENTAS - SEDE',	'382030074 LAVANDERIA - SEDE',	'382030081 BRINDES - SEDE',	'382030082 DESPESAS CERTIFICADOS, TAXAS E CARTORIO',	'382030083 TAXAS E LICENCIAMENTO DE VEICULOS',	'382039999 OUTRAS DESPESAS GERAIS',	'382040001 SERVICOS DE TERCEIROS - PESSOA FISICA',	'382040004 HONORARIOS CONTABEIS',	'382040005 HONORARIOS ADVOCATICIOS',	'382040006 ASSESSORIA E CONSULTORIA',	'382040007 DESPESAS MEDICAS - SEDE',	'382040011 CONSULTORIA DE INFORMATICA',	'382040015 MATERIAL E SERVICOS DE LIMPEZA - SEDE',	'382040017 SEGURANCA E VIGILANCIA - SEDE',	'382040018 SOFTWARE - SEDE',	'382049999 OUTROS SERVICOS TECNICOS - SEDE',	'382050001 FRETES E CARRETOS - SEDE',	'382050004 ESTACIONAMENTO - SEDE',	'382050006 TAXI - SEDE',	'382050010 PASSAGENS AEREAS - SEDE',	'382050011 TRANSPORTE - SEDE',	'382050012 OUTROS SERVICOS DE TRANSPORTE -SEDE',	'382050013 SERVICOS DE ENTREGA-SEDE',	'382070001 IMOVEIS',	'382070002 AGUA/SABESP',	'382070003 TELEFONIA',	'382070004 ENERGIA ELETRICA',	'382070005 INTERNET',	'382079999 OUTROS GASTOS DE INFRAESTRUTURA - SEDE',	'382080001 MANUT. MAQ. E EQUIP.',	'382080002 MANUTENCAO FROTA',	'382080003 BENS MOVEIS',	'382080004 REFORMAS E REPAROS',	'382089999 OUTROS GASTOS COM MANUTENCAO - SEDE',	'383010002 IPTU',	'383010003 IRRF',	'383010004 ICMS SUBSTITUICAO TRIBUTARIA',	'383010005 TAXAS E EMOLUMENTOS',	'383010006 IPVA',	'383010007 ISS RETENCOES',	'383010010 IMPOSTO IMPORTACAO',	'383010011 IMPOSTOS E TAXAS',	'383020001 MULTAS E JUROS S/ TRIBUTOS',	'383020002 MULTAS DE TRANSITO',	'384010001 MANUTENCAO DE CARROS CLASSICOS',	'384010002 DESPESA SEM COMPROVANTE',	'384010003 DESPESA REF. PESSOA FISICA',	'386080099 COMISSOES DE VENDAS',	'387010001 JUROS PAGOS OU INCORRIDOS',	'387010003 COMISSOES E DESPESAS BANCARIAS',	'387010005 IOF',	'387010006 TAXA DE COMISSAO DE CARTAO',	'387020001 DESCONTOS OBTIDOS',	'387020002 JUROS RECEBIDOS OU AUFERIDOS',	'387020003 RENDIMENTOS DE APLICACOES FINANCEIRAS',	'392010001 IRPJ',	'392020001 CSLL',	'431010001 SERVICOS PRESTADOS',	'431020001 SEVICOS PRESTADOS AO EXTERIOR',	'441020001 LOCACOES',	'483010001 (-) CONTRIBUICAO PREVIDENCIARIA',	'483010002 (-) COFINS',	'483010003 (-) PIS',	'483010004 (-) IPI',	'483010005 (-) SIMPLES NACIONAL',	'483030001 (-) ISS',	'511010001 RESULTADO DO EXERCICIO']
finalidade = ttk.Combobox(NF_envio, values=listafinalidade)
finalidade.grid(column=1, row=7, columnspan=2, sticky=EW)

# finalidade = StringVar
# menufinalidade = OptionMenu(NF_envio, finalidade, *listafinalidade)
# menufinalidade.grid(column=1, row=6, columnspan=2, sticky=EW)
'''finalidade = Entry(NF_envio)
finalidade.grid(column=1, row=6, columnspan=2, sticky=EW)'''

tprojeto = Label(NF_envio, text='Projeto:*')
tprojeto.grid(column=0, row=8, padx=10, pady=5, sticky=E)
listaprojeto = ['ANEXO', 'TEMP2022', '22DENERPF', '22ESP', '22ET1', '22ET2', '22ET3', '22ET4', '22ET5', '22ET6', '22ET7', '22ET8', '22ET9', '22PJE']
projeto = ttk.Combobox(NF_envio, values=listaprojeto)
projeto.grid(column=1, row=8, columnspan=2, sticky=EW)

tsetor = Label(NF_envio, text='Setor:*')
tsetor.grid(column=0, row=9, padx=10, pady=5, sticky=E)
listasetor = ['ANEXO', 'ADM',	'ENG',	'EST',	'EVT',	'LOG',	'MKT',	'OFC',	'PER',	'RH']
setor = ttk.Combobox(NF_envio, values=listasetor)
setor.grid(column=1, row=9, columnspan=2, sticky=EW)

tsubsetor = Label(NF_envio, text='Subsetor:*')
tsubsetor.grid(column=0, row=10, padx=10, pady=5, sticky=E)
listasubsetor = ['ANEXO', 'ADE',	'ADM',	'ALMOX',	'CLÁSSICOS',	'COVID19',	'ENG',	'EVENTOS',	'FIN',	'FUNILARIA',	'IMPRENSA',	'LOG',	'MARKETING',	'MÍDIA',	'PER',	'PEÇAS',	'PNEU/RODA',	'PWT',	'REC E DES',	'REVISÃO',	'RH',	'TI']
subsetor = ttk.Combobox(NF_envio, values=listasubsetor)
subsetor.grid(column=1, row=10, columnspan=2, sticky=EW)

tvencimento = Label(NF_envio, text='Vencimento da NF:*')
tvencimento.grid(column=0, row=11, padx=10, pady=5, sticky=E)
vencimento = Entry(NF_envio)
vencimento.grid(column=1, row=11, columnspan=2, sticky=EW)

tcartao = Label(NF_envio, text='4 ultimos digitos do cartão:*')
tcartao.grid(column=0, row=12, padx=10, pady=5, sticky=E)
cartao = Entry(NF_envio)
cartao.grid(column=1, row=12, columnspan=2, sticky=EW)

tnome = Label(NF_envio, text='Nome do autor do e-mail:*')
tnome.grid(column=0, row=13, pady=5)
nome = Entry(NF_envio)
nome.grid(column=1, row=13, columnspan=2, sticky=EW)

tobservacao = Label(NF_envio, text='Observação:  ')
tobservacao.grid(column=0, row=14, padx=10, pady=5, sticky=E)
observacao = Entry(NF_envio)
observacao.grid(column=1, row=14, columnspan=2, sticky=EW)

tboleto = Label(NF_envio, text='Possui boleto?*')
tboleto.grid(column=0, row=15, padx=10, pady=5, sticky=E)

listaBoleto = ['NÃO', 'SIM']
vlista = StringVar()
vlista.set('')  # Definir valor padrão para o option menu

menuBoleto = OptionMenu(NF_envio, vlista, *listaBoleto)
menuBoleto.grid(column=1, row=15, columnspan=2, sticky=EW)

NF_envio.mainloop()
