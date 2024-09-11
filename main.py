#atualizado em 23/06/2023

from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import streamlit as st
import pandas as pd
from io import BytesIO

scope = ['https://spreadsheets.google.com/feeds']
k = st.secrets["senha"]
json = {
    "auth_provider_x509_cert_url":st.secrets["auth_provider_x509_cert_url"],
    "auth_uri":st.secrets["auth_uri"],
    "client_email":st.secrets["client_email"],
    "client_id":st.secrets["client_id"],
    "client_x509_cert_url":st.secrets["client_x509_cert_url"],
    "private_key":st.secrets["private_key"],
    "private_key_id":st.secrets["private_key_id"],
    "project_id":st.secrets["project_id"],
    "token_uri":st.secrets["token_uri"],
    "type":st.secrets["type"]
}

creds = ServiceAccountCredentials.from_json_keyfile_dict(json, scope)

cliente = gspread.authorize(creds)

sheet = ''  # cliente.open_by_key('1lWuFWU8lnw-WoLhfrd9GO4sFjATddgEM5-7XeOc00HM').get_worksheet("Respostas - Editável")
dados = ''  # sheet.get_all_records()  # Get a list of all records
df = ''  # pd.DataFrame(dados)


def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'})
    worksheet.set_column('A:A', None, format1)
    writer.save()
    processed_data = output.getvalue()
    return processed_data


# conexão planilha
def conexao(aba="Respostas - Editável",
            chave='1lWuFWU8lnw-WoLhfrd9GO4sFjATddgEM5-7XeOc00HM'):  # pasta="Transporte e Limpeza de Geladeira (respostas)",
    """
    carrega os dados da planilha do google sheets

    """
    sheet = cliente.open_by_key(chave).worksheet(aba)  # Open the spreadhseet
    dados = sheet.get_all_records()
    df = pd.DataFrame(dados)
    return sheet, dados, df


datando = []
origem_predio = []
origem_sala = []
destino_predio = []
destino_sala = []
nome = []
tipos = []
descricao = []
patrimonio = []
qtd = []
status = []
data_agendamento = []
obsemail = []
obsinterna = []
telefone = []
codigo = []
fotos = []

todos_status = ['', 'Procedente', 'Data não disponível', 'Atendida', 'Ignorar', 'Solicitação Repetida',
                'Restrição de pessoal', 'Não procedente', 'Cancelado']

# padroes
padrao = '<p style="font-family:Courier; color:Blue; font-size: 15px;"'
infor = '<p style="font-family:Courier; color:Green; font-size: 15px;"'
alerta = '<p style="font-family:Courier; color:Red; font-size: 15px;"'
titulo = '<p style="font-family:Courier; color:Blue; font-size: 20px;"'
cabecalho = '<div id="logo" class="span8 small"><a title="Universidade Federal do Tocantins"><img src="https://ww2.uft.edu.br/images/template/brasao.png" alt="Universidade Federal do Tocantins"><span class="portal-title-1"></span><h1 class="portal-title corto">Universidade Federal do Tocantins</h1><span class="portal-description">COINFRA - Transporte Predial e Limpeza de Geladeira/Bebedouro</span></a></div>'

st.sidebar.title('Gestão Transporte e Limpeza de Geladeira/Bebedouro')
a = k
# pg=st.sidebar.selectbox('Selecione a Página',['Solicitações em Aberto','Solicitações a Finalizar','Consulta'])
pg = st.sidebar.radio('', ['Transporte', 'Galão de Água', 'Controle de Galões','Limpeza de Geladeira/Bebedouro', 'Consulta Transporte'])

if (pg == 'Transporte'):
    st.markdown(cabecalho, unsafe_allow_html=True)
    st.subheader(pg)
    status_selecionar = st.selectbox('Filtrar por Status:', todos_status)
    ch_posterior = st.checkbox('Somente solicitações a executar ou posteriores a hoje', value=True)

    sheet, dados, df = conexao()
    df = df.astype(str)
    for dic in dados:

        if ch_posterior:
            if dic['Carimbo de data/hora'] != '' and dic['Status'] == status_selecionar and dic['Código'] != '' and str(
                    dic['data_posterior']) == '1':
                origem_predio.append(dic['Origem - Prédio'])
                origem_sala.append(dic['Origem - Sala/Local'])
                destino_predio.append(dic['Destino - Prédio'])
                destino_sala.append(dic['Destino - Sala/Local'])
                nome.append(dic['Nome do solicitante'])
                tipos.append(dic['Tipos de Materiais'])
                descricao.append(dic['Descrição'])
                patrimonio.append(dic[
                                      'Os materiais/equipamentos a serem transportados são patrimoniados (possuem número de patrimônio)?'])
                qtd.append(dic['Quantidade total aproximada de materiais'])
                status.append(dic['Status'])
                data_agendamento.append(dic['data agendamento'])
                obsemail.append(dic['Obs e-mail'])
                obsinterna.append(dic['Obs para os Carregadores'])
                telefone.append(dic['Telefone'])
                codigo.append(dic['Código'])
                fotos.append(dic['Fotos'])
        else:
            if dic['Carimbo de data/hora'] != '' and dic['Status'] == status_selecionar and dic['Código'] != '':
                origem_predio.append(dic['Origem - Prédio'])
                origem_sala.append(dic['Origem - Sala/Local'])
                destino_predio.append(dic['Destino - Prédio'])
                destino_sala.append(dic['Destino - Sala/Local'])
                nome.append(dic['Nome do solicitante'])
                tipos.append(dic['Tipos de Materiais'])
                descricao.append(dic['Descrição'])
                patrimonio.append(dic[
                                      'Os materiais/equipamentos a serem transportados são patrimoniados (possuem número de patrimônio)?'])
                qtd.append(dic['Quantidade total aproximada de materiais'])
                status.append(dic['Status'])
                data_agendamento.append(dic['data agendamento'])
                obsemail.append(dic['Obs e-mail'])
                obsinterna.append(dic['Obs para os Carregadores'])
                telefone.append(dic['Telefone'])
                codigo.append(dic['Código'])
                fotos.append(dic['Fotos'])

    if status_selecionar == '' and len(codigo)>0:
        if len(codigo)==1:
            st.markdown(alerta + '<p><b>Existe 1 solicitação nova pendente de análise!</b></p>', unsafe_allow_html=True)
        else:
            try:
                st.markdown(alerta + '<p><b>Existem ' + str(len(codigo)) +' solicitações novas pendentes de análise!</b></p>', unsafe_allow_html=True)
            except:
                st.markdown(alerta + '<p><b>Existem solicitações novas pendentes de análise!</b></p>', unsafe_allow_html=True)


    selecionado = st.selectbox('Nº da solicitação:', codigo)

    if (len(codigo) > 0):
        with st.form(key='my_form'):
            n = codigo.index(selecionado)

            links = fotos[n]
            links = links.split(',')
            midia = ""
            for i in range(len(links)):
                if (links[i]!=''):
                    midia = midia + "<p><a href='" + links[i] + "'>Mídia " + str(i + 1) + "</a></p>"
            # (midia)
            # apresentar dados da solicitação
            st.markdown(titulo + '<p><b>Dados da Solicitação</b></p>', unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Nome</b>: ' + str(nome[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Telefone</b>: ' + str(telefone[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Prédio de Origem</b>: ' + str(origem_predio[n]) + '</p>',
                        unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Sala de Origem</b>: ' + str(origem_sala[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Prédio de Destino</b>: ' + str(destino_predio[n]) + '</p>',
                        unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Sala de Destino</b>: ' + str(destino_sala[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Tipos de Materiais</b>: ' + str(tipos[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Bens com patrimônio?</b>: ' + str(patrimonio[n]) + '</p>',
                        unsafe_allow_html=True)
            st.markdown(padrao + "<p><b>Fotos/Vídeos</b>: " + midia + "</p>", unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Quantidade</b>: ' + str(qtd[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Descrição</b>: ' + str(descricao[n]) + '</p>', unsafe_allow_html=True)
            #st.markdown(padrao + '<b>Telefone</b>:<p> ' + str(telefone[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Data</b>: ' + str(data_agendamento[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Status Atual</b>: ' + str(status[n]) + '</p>', unsafe_allow_html=True)

            # Data
            d = '01/01/2023'
            # print('Data Agendamento registrada: ' + d_agend[n])
            if (data_agendamento[n] != ''):
                d = data_agendamento[n]
            else:
                # st.text('OS sem agendamento registrado ou com data de agendamento anterior a hoje!')
                print('Sem data registrada')
            d = d.replace('/', '-')

            data_ag = datetime.strptime(d, '%d-%m-%Y')

            if (data_ag == ''):
                data_ag = datetime.strptime("01-01-2023", '%d-%m-%Y')

            data_agendamento = st.date_input('Data de Transporte (ANO/MÊS/DIA)', value=data_ag)
            celula = sheet.find(codigo[n])
            status_alterado = st.selectbox('Selecione o status:', todos_status, index=todos_status.index(status[n]))
            obsemail_texto = st.text_area('Observação para o usuário: ', value=obsemail[n])
            obsinterna_texto = st.text_area('Observação interna: ', value=obsinterna[n])

            s = st.text_input("Senha:", value="", type="password")
            botao = st.form_submit_button('Registrar')
            if (botao == True and s == a):

                with st.spinner('Registrando dados...'):
                    try:
                        data = data_agendamento
                        data_formatada = str(data.day) + '/' + str(data.month) + '/' + str(data.year)
                        sheet.update_acell('W' + str(celula.row), data_formatada)
                        sheet.update_acell('P' + str(celula.row), status_alterado)
                        sheet.update_acell('AG' + str(celula.row), obsemail_texto)
                        sheet.update_acell('AH' + str(celula.row), obsinterna_texto)
                        sheet.update_acell('T' + str(celula.row), '')
                        st.success('Dados registrados!')
                    except Exception as e:
                        st.error('Ocorreu um erro ao tentar registrar os dados! ' + str(e))

            elif (botao == True and s != a):
                st.markdown(alerta + '<b>Senha incorreta!</b></p>', unsafe_allow_html=True)
    else:
        st.markdown(infor + '<b>Não há itens na condição ' + pg + '</b></p>', unsafe_allow_html=True)









# limpeza de geladeira / bebedouro

elif (pg == 'Limpeza de Geladeira/Bebedouro'):
    st.markdown(cabecalho, unsafe_allow_html=True)
    st.subheader(pg)


    #Datas disponíveis
    with st.expander("Datas Disponíveis"):
        chave = '1JAz12fD-1-zk0Iraa4dbNC_K8ygb-xHLYQwv5xjf3nM'
        aba = 'Datas'
        sheet2, dados2, df2 = conexao(aba=aba, chave=chave)

        datas_disponiveis = []
        dia_semana = []
        agendamentos = []
        disponivel = []

        for dic in dados2:
            if dic['verificacao'] == 1 and dic['Data'] != '':
                datas_disponiveis.append(dic['Data'])
                dia_semana.append(dic['Dia'])
                agendamentos.append(dic['Quantidade de Agendamentos'])
                disponivel.append(dic['Data Disponível'])

        data_disponivel = st.selectbox('Selecione a data', datas_disponiveis)
        n = datas_disponiveis.index(data_disponivel)

        with st.form(key='my_form2'):
            print(n)
            print(dia_semana[n])
            st.markdown(padrao + '<p><b>Dia da semana</b>: ' + str(dia_semana[n]) + '</p>',
                        unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Quantidade de agendamentos no dia</b>: ' + str(agendamentos[n]) + '</p>', unsafe_allow_html=True)
            disponibilidade = st.selectbox('Data disponível?',['Sim','Não'],index=['Sim','Não'].index(disponivel[n]))
            s1 = st.text_input("Senha:", value="", type="password")
            botao1 = st.form_submit_button('Registrar')
            celula1 = sheet2.find(data_disponivel)
            if (botao1 == True and s1 == a):
                with st.spinner('Registrando dados...'):
                    try:
                        sheet2.update_acell('D' + str(celula1.row), disponibilidade)
                        st.success('Dados registrados!')
                    except Exception as e:
                        st.error('Ocorreu um erro ao tentar registrar os dados! ' + str(e))

    #Respostas Editáveis
    chave = '1JAz12fD-1-zk0Iraa4dbNC_K8ygb-xHLYQwv5xjf3nM'
    aba = 'Respostas Editável'
    sheet1, dados1, df1 = conexao(aba=aba, chave=chave)
    todos_status1 = ['', 'Procedente', 'Finalizada', 'Cancelada', 'Data Não Disponível', 'Não Procedente', 'Manutenção']
    predio = []
    sala = []
    nome = []
    telefone = []
    descricao = []
    status = []
    data_limpeza = []
    obsemail = []
    obsinterna = []
    codigo = []
    fotos = []
    hora = []


    status_selecionar = st.selectbox('Filtrar por Status:', todos_status1)
    ch_posterior = st.checkbox('Somente solicitações a executar ou posteriores a hoje', value=True)

    for dic in dados1:

        if ch_posterior:
            if dic['Carimbo de data/hora'] != '' and dic['Status'] == status_selecionar and dic[
                'Nº da Solicitação'] != '' and str(dic['data_posterior']) == '1':
                predio.append(dic['Prédio'])
                sala.append(dic['Número da Sala/Local'])
                nome.append(dic['Nome do Solicitante'])
                telefone.append(dic['Telefone'])
                descricao.append(dic['Observações'])
                status.append(dic['Status'])
                data_limpeza.append(dic['Data Texto'])
                obsemail.append(dic['Obs E-mail'])
                obsinterna.append(dic['Obs Interna'])
                codigo.append(dic['Nº da Solicitação'])
                fotos.append(dic['Foto / Vídeo (Opcional)'])
                hora.append(dic['Horário definitivo'])
        else:
            if dic['Carimbo de data/hora'] != '' and dic['Status'] == status_selecionar and dic[
                'Nº da Solicitação'] != '':
                predio.append(dic['Prédio'])
                sala.append(dic['Número da Sala/Local'])
                nome.append(dic['Nome do Solicitante'])
                telefone.append(dic['Telefone'])
                descricao.append(dic['Observações'])
                status.append(dic['Status'])
                data_limpeza.append(dic['Data Texto'])
                obsemail.append(dic['Obs E-mail'])
                obsinterna.append(dic['Obs Interna'])
                codigo.append(dic['Nº da Solicitação'])
                fotos.append(dic['Foto / Vídeo (Opcional)'])
                hora.append(dic['Horário definitivo'])

    selecionado = st.selectbox('Nº da solicitação:', codigo)

    if (len(codigo) > 0):
        with st.form(key='my_form'):
            n = codigo.index(selecionado)

            links = fotos[n]
            links = links.split(',')
            midia = ""
            for i in range(len(links)):
                if (links[i] != ''):
                    midia = midia + "<a href='" + links[i] + "'>Mídia " + str(i + 1) + "</a> | "
            # (midia)
            # apresentar dados da solicitação
            st.markdown(titulo + '<p><b>Dados da Solicitação</b></p>', unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Nome</b>: ' + str(nome[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Telefone</b>: ' + str(telefone[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Prédio</b>: ' + str(predio[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Sala</b>: ' + str(sala[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + "<p><b>Fotos/Vídeos</b>: " + midia + "</p>", unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Descrição</b>: ' + descricao[n] + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Data</b>: ' + str(data_limpeza[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Hora</b>: ' + str(hora[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<p><b>Status Atual</b>: ' + str(status[n]) + '</p>', unsafe_allow_html=True)

            # Data
            horarios = ['','09:00 h','10:00 h','11:00 h','12:00 h','13:00 h','14:00 h','15:00 h','16:00 h']
            celula = sheet1.find(codigo[n])
            status_alterado = st.selectbox('Selecione o status:', todos_status1, index=todos_status1.index(status[n]))
            horarios_agendamento = st.selectbox('Selecione o horario:', horarios, index=horarios.index(hora[n]))
            # Data
            d = '01/01/2023'
            # print('Data Agendamento registrada: ' + d_agend[n])
            if (data_limpeza[n] != ''):
                d = data_limpeza[n]
            else:
                # st.text('OS sem agendamento registrado ou com data de agendamento anterior a hoje!')
                print('Sem data registrada')
            d = d.replace('/', '-')

            data_ag = datetime.strptime(d, '%d-%m-%Y')

            if (data_ag == ''):
                data_ag = datetime.strptime("01-01-2023", '%d-%m-%Y')

            data_agendamento = st.date_input('Data da Limpeza de Geladeira (ANO/MÊS/DIA)', value=data_ag)
            obsemail_texto = st.text_area('Observação para o usuário: ', value=obsemail[n])
            obsinterna_texto = st.text_area('Observação interna: ', value=obsinterna[n])

            s = st.text_input("Senha:", value="", type="password")

            botao = st.form_submit_button('Registrar')
            if (botao == True and s == a):

                with st.spinner('Registrando dados...'):
                    try:
                        data = data_agendamento
                        data_formatada = str(data.day) + '/' + str(data.month) + '/' + str(data.year)
                        sheet1.update_acell('W' + str(celula.row), data_formatada)
                        sheet1.update_acell('AG' + str(celula.row), horarios_agendamento)
                        sheet1.update_acell('K' + str(celula.row), status_alterado)
                        sheet1.update_acell('O' + str(celula.row), obsemail_texto)
                        sheet1.update_acell('P' + str(celula.row), obsinterna_texto)
                        sheet1.update_acell('M' + str(celula.row), '')
                        st.success('Dados registrados!')
                    except Exception as e:
                        st.error('Ocorreu um erro ao tentar registrar os dados! ' + str(e))

            elif (botao == True and s != a):
                st.markdown(alerta + '<b>Senha incorreta!</b></p>', unsafe_allow_html=True)
    else:
        st.markdown(infor + '<b>Não há itens na condição ' + pg + '</b></p>', unsafe_allow_html=True)
elif pg == 'Galão de Água':

    chave = '1TykbuQopkNBMzZ77_aSxCkcEvp3sCu62kXXYSVXzzjA'
    aba = 'Água(Editável)'
    sheet1, dados1, df1 = conexao(aba=aba, chave=chave)
    # print(f'Dados: {df1}')

    todos_status1 = ['', 'Procedente', 'Entregue', 'Reitoria', 'Ninguém no local','Excesso de estoque', 'Não procedente',
                     'Usuário retirou no local', 'Falta de água no estoque', 'Cancelado', 'Solicitação Repetida',
                     'Problemas Operacionais-Reagendamento']
    email = []
    predio = []
    sala = []
    nome = []
    quantidade = []
    status = []
    data_agendada = []
    obsemail = []
    codigo = []
    cod_confirmacao = []
    entregue = []

    st.markdown(cabecalho, unsafe_allow_html=True)
    st.subheader(pg)
    status_selecionar = st.selectbox('Filtrar por Status:', todos_status1)

    ch_posterior = st.checkbox('Somente solicitações posteriores a hoje', value=True)

    for dic in dados1:

        if ch_posterior:
            # if str(dic['data_posterior'])!='0':
            #     print(f"Data_posterior: {dic['data_posterior']} | Data {dic['Data Pré-Agendada']} | verificação {dic['data_posterior']=='1'}")
            if dic['Carimbo de data/hora'] != '' and dic['Status'] == status_selecionar and dic['CÓDIGO'] != '' and str(
                    dic['data_posterior']) == '1':
                data_agendada.append(dic['Data Pré-Agendada'])
        else:
            if dic['Carimbo de data/hora'] != '' and dic['Status'] == status_selecionar and dic['CÓDIGO'] != '':
                data_agendada.append(dic['Data Pré-Agendada'])

    data_solicitacao = st.selectbox('Data de entrega', list(set(data_agendada)))
    data_agendada = []
    for dic in dados1:

        if ch_posterior:
            if dic['Carimbo de data/hora'] != '' and dic['Status'] == status_selecionar and dic['CÓDIGO'] != '' and str(
                    dic['data_posterior']) == '1' and dic['Data Pré-Agendada'] == data_solicitacao:
                email.append(dic['Endereço de e-mail'])
                predio.append(dic['Prédio'])
                sala.append(dic['Sala/Local'])
                nome.append(dic['Nome do solicitante'])
                quantidade.append(dic['Insira a quantidade de Galões de 20 L'])
                status.append(dic['Status'])
                data_agendada.append(dic['Data Pré-Agendada'])
                obsemail.append(dic['Obs'])
                codigo.append(dic['CÓDIGO'])
                cod_confirmacao.append(dic['código de confirmação'])
                entregue.append(dic['confirmação'])

        else:
            if dic['Carimbo de data/hora'] != '' and dic['Status'] == status_selecionar and dic['CÓDIGO'] != '' and dic[
                'Data Pré-Agendada'] == data_solicitacao:
                email.append(dic['Endereço de e-mail'])
                predio.append(dic['Prédio'])
                sala.append(dic['Sala/Local'])
                nome.append(dic['Nome do solicitante'])
                quantidade.append(dic['Insira a quantidade de Galões de 20 L'])
                status.append(dic['Status'])
                data_agendada.append(dic['Data Pré-Agendada'])
                obsemail.append(dic['Obs'])
                codigo.append(dic['CÓDIGO'])
                cod_confirmacao.append(dic['código de confirmação'])
                entregue.append(dic['confirmação'])

    # calculando quantidade por bloco
    t = "<br><label><strong>Resumo de Quantidades a Entregar por Bloco</strong></label>"
    t = t + """<table class="table table-striped">
      <thead>
          <tr>
            <th scope="col">Prédio</th>
            <th scope="col">Quantidade a Entregar</th>
          </tr>
    </thead>
    <tbody>
      """
    list_aux = sorted(list(set(predio)))
    q_aux = []
    for pr in list_aux:
        n = 0
        k = 0
        entr = 0
        for j in predio:
            if j == pr and entregue[k] == '':
                if quantidade[k]!='':
                    n += int(quantidade[k])
                o = 0
                if entregue[k]!='':
                    if quantidade[k] != '':
                        o=int(quantidade[k])
                entr += o

            k += 1
        t = t + """<tr>
                      <th scope="row">""" + str(pr) + """</th>
                      <th scope="row">""" + str(n) + """</th>                      
                   </tr>"""
        q_aux.append(n)


    t = t + """
          </tbody>
      </table>
      """

    if len(q_aux)==0:
        t=""
    # with st.expander('Solicitações'):
    # print(f'data_agendada: {len(data_agendada)} | nome: {len(nome)} | prédio: {len(predio)} | sala: {len(sala)} | prédio: {len(quantidade)}')
    # dfs = pd.DataFrame({'Nome': nome, 'Prédio': predio, 'Sala': sala, 'Qtd': quantidade, 'Data': data_agendada})

    # print(dfs)
    #
    # dfs = df1[['CÓDIGO','Prédio','Sala/Local','Nome do solicitante','Insira a quantidade de Galões de 20 L','Data Pré-Agendada']].isin(codigo)
    # #dfs = df1.astype(df1)
    # dad = dados1[dfs]
    # st.dataframe(dad)
    cod, pred, sal, nom, qt, dat = st.columns(6)
    pr = predio
    n = 0
    texto = """
    <table class="table table-striped">
    <thead>
        <tr>
          <th scope="col">Código</th>
          <th scope="col">Prédio</th>
          <th scope="col">Sala</th>
          <th scope="col">Nome</th>
          <th scope="col">Quantidade</th>
          <th scope="col">Entregue</th>
        </tr>
  </thead>
  <tbody>

    """
    for i in range(len(codigo)):
        if (i == 0):
            # cod.text('Código')
            # pred.text('Prédio')
            # sal.text('Sala')
            # nom.text('Nome')
            # qt.text('Qtd')
            print('i')
            # dat.text('Data')

        # l = sorted(predio)
        k = i  # predio.index(l[i])
        if data_agendada[k] != '':
            # cod.text(codigo[k])
            # pred.text(predio[k])
            # sal.text(sala[k])
            # nom.text(nome[k])
            # qt.text(quantidade[k])
            entrega = ""
            if entregue[k] != '':
                entrega = "X"

            j = """<tr>
              <th scope="row">""" + str(codigo[k]) + """</th>
              <th scope="row">""" + str(predio[k]) + """</th>
              <th scope="row">""" + str(sala[k]) + """</th>
              <th scope="row">""" + str(nome[k]) + """</th>
              <th scope="row">""" + str(quantidade[k]) + """</th>  
              <th scope="row">""" + str(entrega) + """</th>  
             </tr>"""
            # dat.text(data_agendada[k])
            texto = texto + j
        n += 1
    j = """
          </tbody>
      </table>
      <br>
      """
    texto = texto + j
    st.markdown(texto, unsafe_allow_html=True)
    if len(q_aux)>0:
        with st.expander('Resumo de Quantidade a Entregar por Bloco'):
            st.markdown(t, unsafe_allow_html=True)
    # barra = pd.DataFrame((zip(predio,quantidade)),columns=["predio", "quantidade"])
    # st.bar_chart(barra)
    cod_alt = []
    for i in range(len(codigo)):
        if entregue[i] == '':
            cod_alt.append(codigo[i])

    selecionado = st.selectbox('Nº da solicitação:', cod_alt)

    if (len(codigo) > 0 and len(cod_alt)):
        with st.form(key='my_form'):
            n = codigo.index(selecionado)

            # apresentar dados da solicitação
            st.markdown(titulo + '<p><b>Dados da Solicitação</b></p>', unsafe_allow_html=True)
            st.markdown(padrao + '<b>Nome</b>:<p> ' + str(nome[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<b>Prédio</b>:<p> ' + str(predio[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<b>Sala</b>:<p> ' + str(sala[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + "<b>Quantidade</b>:<p> " + str(quantidade[n]) + "</p>", unsafe_allow_html=True)
            st.markdown(padrao + '<b>Data</b>: <p>' + str(data_agendada[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<b>E-mail</b>: <p>' + str(email[n]) + '</p>', unsafe_allow_html=True)
            st.markdown(padrao + '<b>Status Atual</b>:<p> ' + str(status[n]) + '</p>', unsafe_allow_html=True)
            status_alterado = st.selectbox('Selecione o status:', todos_status1, index=todos_status1.index(status[n]))
            obsemail_texto = st.text_area('Observação para o usuário: ', value=obsemail[n])

            if status_selecionar=='Falta de água no estoque':
                codigo_confirmacao = cod_confirmacao[n]
            else:
                codigo_confirmacao = st.text_input('Código de confirmação de entrega: ', value="")
                with st.expander('Código de confirmação...'):
                    st.text(cod_confirmacao[n])

            celula = sheet1.find(codigo[n])
            s = st.text_input("Senha:", value="", type="password")
            st.markdown('<label style="color:blue"><i>Obs.: Lembrar de <strong>verificar se está selecionado o Status correto</strong> antes de Registrar!</i></label>', unsafe_allow_html=True)
            botao = st.form_submit_button('Registrar')
            if (botao == True and s == a and str(codigo_confirmacao) == str(cod_confirmacao[n])):

                with st.spinner('Registrando dados...'):
                    try:
                        sheet1.update_acell('K' + str(celula.row), status_alterado)
                        sheet1.update_acell('AJ' + str(celula.row), obsemail_texto)
                        sheet1.update_acell('AQ' + str(celula.row), codigo_confirmacao)
                        st.success('Dados registrados!')
                    except Exception as e:
                        st.error('Ocorreu um erro ao tentar registrar os dados! ' + str(e))

            elif (botao == True and s != a):
                st.markdown(alerta + '<b>Senha incorreta!</b></p>', unsafe_allow_html=True)
            elif (botao == True and codigo_confirmacao != cod_confirmacao[n]):
                st.markdown(alerta + '<b>Código de confirmação incorreto!</b></p>', unsafe_allow_html=True)
    else:
        st.markdown(infor + '<b>Não há mais águas para entrega</b></p>', unsafe_allow_html=True)



elif pg == 'Controle de Galões':

    chave = '1TykbuQopkNBMzZ77_aSxCkcEvp3sCu62kXXYSVXzzjA'
    aba = 'Água(Editável)'
    sheet1, dados1, df1 = conexao(aba=aba, chave=chave)
    # print(f'Dados: {df1}')

    todos_status1 = ['', 'Procedente', 'Entregue', 'Reitoria', 'Ninguém no local','Excesso de estoque', 'Não procedente',
                     'Usuário retirou no local', 'Falta de água no estoque', 'Cancelado', 'Solicitação Repetida',
                     'Problemas Operacionais-Reagendamento']
    email = []
    predio = []
    sala = []
    nome = []
    quantidade = []
    status = []
    data_agendada = []
    obsemail = []
    codigo = []
    cod_confirmacao = []
    entregue = []

    st.markdown(cabecalho, unsafe_allow_html=True)
    st.subheader(pg)


    #status_selecionar = st.selectbox('Filtrar por Status:', todos_status1)



    st.markdown(titulo + '<p><b>Solicitações com status de Falta de Galão no Estoque:</b></p>', unsafe_allow_html=True)
    for dic in dados1:
        if dic['Carimbo de data/hora'] != '' and dic['Status'] == 'Falta de água no estoque' and dic['CÓDIGO'] != '':
            data_agendada.append(dic['Data Pré-Agendada'])
            email.append(dic['Endereço de e-mail'])
            predio.append(dic['Prédio'])
            sala.append(dic['Sala/Local'])
            nome.append(dic['Nome do solicitante'])
            quantidade.append(dic['Insira a quantidade de Galões de 20 L'])
            status.append(dic['Status'])
            obsemail.append(dic['Obs'])
            codigo.append(dic['CÓDIGO'])
            cod_confirmacao.append(dic['código de confirmação'])
            entregue.append(dic['confirmação'])

    #data_solicitacao = st.selectbox('Data de entrega', list(set(data_agendada)))
    #data_agendada = []


    cod, pred, sal, nom, qt, dat = st.columns(6)
    pr = predio
    n = 0
    texto = """
    <table class="table table-striped">
    <thead>
        <tr>
          <th scope="col">Código</th>
          <th scope="col">Prédio</th>
          <th scope="col">Sala</th>
          <th scope="col">Nome</th>
          <th scope="col">Quantidade</th>
          <th scope="col">Data Inicial</th>
        </tr>
  </thead>
  <tbody>

    """
    for i in range(len(codigo)):
        if (i == 0):
            # cod.text('Código')
            # pred.text('Prédio')
            # sal.text('Sala')
            # nom.text('Nome')
            # qt.text('Qtd')
            print('i')
            # dat.text('Data')

        # l = sorted(predio)
        k = i  # predio.index(l[i])
        if data_agendada[k] != '':
            # cod.text(codigo[k])
            # pred.text(predio[k])
            # sal.text(sala[k])
            # nom.text(nome[k])
            # qt.text(quantidade[k])
            entrega = ""
            if entregue[k] != '':
                entrega = "X"

            j = """<tr>
              <th scope="row">""" + str(codigo[k]) + """</th>
              <th scope="row">""" + str(predio[k]) + """</th>
              <th scope="row">""" + str(sala[k]) + """</th>
              <th scope="row">""" + str(nome[k]) + """</th>
              <th scope="row">""" + str(quantidade[k]) + """</th>  
              <th scope="row">""" + str(data_agendada[k]) + """</th>  
             </tr>"""
            # dat.text(data_agendada[k])
            texto = texto + j
        n += 1
    j = """
          </tbody>
      </table>
      <br>
      """
    texto = texto + j
    st.markdown(texto, unsafe_allow_html=True)
    with st.form(key='my_form'):
        falta_de_agua = st.checkbox('Falta de água no estoque?', value=sheet1.get('A10'))
        data = st.date_input('Selecione a data para reagendamento')
        data_formatada = str(data.day) + '/' + str(data.month) + '/' + str(data.year)
        qt = st.selectbox('Selecione a quantidade para reagendar na data indicada acima',[0].extend([(i+1) for i in range(len(codigo))]))

        botao = st.form_submit_button('Registrar')

        if botao:
            with st.spinner('Registrando dados...'):
                try:
                    sheet1.update_acell('A10', falta_de_agua)
                    n=1
                    while n<=qt:
                        print('Registrando!')
                        celula = sheet1.find(codigo[n-1])
                        sheet1.update_acell('K' + str(celula.row), 'Procedente')
                        sheet1.update_acell('AJ' + str(celula.row), 'Reagendado devido a falta de galões no estoque')
                        sheet1.update_acell('AI' + str(celula.row), data_formatada)
                        n += 1
                except Exception as e:
                    st.error('Ocorreu um erro ao tentar registrar os dados! ' + str(e))









elif pg == 'Consulta Transporte':
    sheet, dados, df = conexao()
    df = df.astype(str)

    data_reg = []
    for dic in dados:
        if dic['Carimbo de data/hora'] != '':
            data_reg.append(dic['Carimbo de data/hora'])
            origem_predio.append(dic['Origem - Prédio'])
            origem_sala.append(dic['Origem - Sala/Local'])
            destino_predio.append(dic['Destino - Prédio'])
            destino_sala.append(dic['Destino - Sala/Local'])
            nome.append(dic['Nome do solicitante'])
            tipos.append(dic['Tipos de Materiais'])
            descricao.append(dic['Descrição'])
            patrimonio.append(dic[
                                  'Os materiais/equipamentos a serem transportados são patrimoniados (possuem número de patrimônio)?'])
            qtd.append(dic['Quantidade total aproximada de materiais'])
            status.append(dic['Status'])
            data_agendamento.append(dic['data agendamento'])
            obsemail.append(dic['Obs e-mail'])
            obsinterna.append(dic['Obs para os Carregadores'])
            telefone.append(dic['Telefone'])
            codigo.append(dic['Código'])

    st.markdown(cabecalho, unsafe_allow_html=True)
    st.subheader(pg)

    lista_dados = ['Carimbo de data/hora', 'Código', 'Origem - Prédio', 'Origem - Sala/Local', 'Destino - Prédio',
                   'Destino - Sala/Local', 'Nome do solicitante', 'Tipos de Materiais', 'Descrição',
                   'Os materiais/equipamentos a serem transportados são patrimoniados (possuem número de patrimônio)?',
                   'Quantidade total aproximada de materiais', 'Telefone', 'Status', 'data agendamento', 'Obs e-mail',
                   'Obs para os Carregadores']

    filtra_por = st.selectbox('Filtrar por:', lista_dados)
    aux = []
    for dic in dados:
        if dic['Carimbo de data/hora'] != '':
            aux.append(dic[filtra_por])
    with st.form(key='form1'):
        aux = set(aux)
        aux = list(aux)
        aux = sorted(aux)
        filtro = st.selectbox('Filtro', aux)
        btn1 = st.form_submit_button('Filtrar')

    if (btn1 == True):
        dados = df[lista_dados]
        filtrar = dados[filtra_por].isin([filtro])
        st.dataframe(dados[filtrar].head())
        df_xlsx = to_excel(dados[filtrar])
        st.download_button(label='📥 Baixar Resultado do Filtro em Excel', data=df_xlsx,
                           file_name='filtro_planilha.xlsx')
    else:
        st.dataframe(df[lista_dados])
        df_xlsx = to_excel(df[lista_dados])
        st.download_button(label='📥 Baixar Resultado do Filtro em Excel', data=df_xlsx,
                           file_name='filtro_planilha.xlsx')

