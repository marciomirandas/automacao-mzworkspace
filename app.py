"""
pip install selenium
pip install webdriver-manager
pip install reportlab
pip install pandas
pip install openpyxl
pip install Unidecode
"""

# Importa as bibliotecas
import time
import os
import logging
import pandas as pd
import sys

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph
from reportlab.lib import colors

from unidecode import unidecode


def simulador1(tipo_simulacao, cidade, valor, renda, fgts, nascimento, dependente, amortizacao,
               prazo, valor_financiamento):
    
    # Teste de variáveis iniciais
    try:
        valor_inicial = int(str(valor).upper().replace('R$', '').replace('.', '').replace(',00', '').strip())
        renda_inicial = int(str(renda).upper().replace('R$', '').replace('.', '').replace(',00', '').strip())
        
        if valor_financiamento != "":
            valor_financiamento_inicial = int(str(valor_financiamento).upper().replace('R$', '').replace('.', '').replace(',00', '').strip())
    except:
        raise ValueError("Erro ao manipular as variáveis iniciais")
    
    if str(tipo_simulacao).strip() == '1':
            
            if valor_inicial > 264000:
                    raise ValueError("Valor do financiamento maior que o limite")

            elif renda_inicial > 8000:
                    raise ValueError("Valor da renda maior que o limite")
            
    if valor_financiamento != "":
        if valor_financiamento_inicial > valor_inicial:
                raise ValueError("Valor do financiamento maior que o valor do imóvel")
    
            
    # Iniciando a segunda janela 
    try:
        site_simulador = 'https://www.portaldeempreendimentos.caixa.gov.br/simulador/'

        driver2 = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        driver2.maximize_window()

        driver2.get(site_simulador)
        time.sleep(5)
    except:
        driver2.quit()
        raise ValueError("Erro ao abrir o simulador 1")
    

    # Simulador
    try:
        origem_recurso = Select(driver2.find_element(By.NAME, "origemRecurso"))

        if str(tipo_simulacao).strip() == '1':
            origem_recurso.select_by_visible_text("FGTS - FUNDO DE GARANTIA POR TEMPO DE SERVICO")
            time.sleep(2)

        elif str(tipo_simulacao).strip() == '2':
            origem_recurso.select_by_visible_text("SBPE")
            time.sleep(2)

        driver2.find_element(By.XPATH, '//*[@id="bottom_bar"]/fieldset/ul/li[2]/a').click()
        time.sleep(5)
    except:
        driver2.quit()
        raise ValueError("Erro na página simulador") 
    

    # Categoria do imóvel
    try:
        categoria_imovel = Select(driver2.find_element(By.NAME, "categoriaImovel"))
        categoria_imovel.select_by_visible_text("CONSTRUCAO/AQ TER CONST - IM. PLANTA E COLETIVAS")
        time.sleep(2)

        cidade_editada = str(cidade).strip()
        if not ' - ' in cidade_editada:
            cidade_editada = str(cidade_editada).replace('-', ' - ')
        driver2.find_element(By.ID, 'cidade').send_keys(cidade_editada)
        time.sleep(3)

        cidade_selecionada = driver2.find_element(By.ID, 'cidade:menu')
        cidade_selecionada.find_element(By.TAG_NAME, 'li').click()
        time.sleep(1)

        valor_editado = str(valor).upper().replace('R$', '').strip()
        if not ',' in valor_editado:
            valor_editado = valor_editado + ',00'
        driver2.find_element(By.ID, 'valorImovel').send_keys(valor_editado)
        time.sleep(1)

        renda_editada = str(renda).upper().replace('R$', '').strip()
        if not ',' in renda_editada:
            renda_editada = renda_editada + ',00'
        driver2.find_element(By.ID, 'renda').send_keys(renda_editada)
        time.sleep(1)

        participantes = Select(driver2.find_element(By.NAME, "quantidadeParticipantes"))
        participantes.select_by_visible_text("1")
        time.sleep(1)

        if str(tipo_simulacao).strip() == '1':
            if str(fgts).strip() == '1':
                tres_anos_fgts = driver2.find_element(By.NAME, "checkbox")
                tres_anos_fgts.click()
                time.sleep(1)

        driver2.find_element(By.XPATH, '//*[@id="bottom_bar"]/fieldset/ul/li[2]/a').click()
        time.sleep(5)
    except:
        driver2.quit()
        raise ValueError("Erro na categoria do imóvel") 


    # Valor da obra e legislação
    try:
        if str(tipo_simulacao).strip() == '1':
            driver2.find_element(By.XPATH, '//*[@id="bottom_bar"]/fieldset/ul/li[2]/a').click()
            time.sleep(5)
    except:
        driver2.quit()
        raise ValueError("Erro no valor da obra e legislação") 

    
    # Participantes
    try:
        nascimento_editado = str(nascimento).replace(' ', '')
        driver2.find_element(By.ID, 'dataNascimento').send_keys(nascimento_editado)
        time.sleep(1)

        driver2.find_element(By.XPATH, '//*[@id="bottom_bar"]/fieldset/ul/li[2]/a').click()
        time.sleep(5)
    except:
        driver2.quit()
        raise ValueError("Erro na página participantes")
    

    # Clica no link
    try:
        modalidade = driver2.find_element(By.ID, 'modalidade')
        links_modalidade = modalidade.find_elements(By.TAG_NAME, 'a')
        for link_modalidade in links_modalidade:

            if str(tipo_simulacao).strip() == '1':
                span_modalidade = link_modalidade.find_element(By.TAG_NAME, 'span')
            
                if span_modalidade.text == 'NPMCMV UNIDADE VINCULADA PF EMPREENDIMENTO COM FINANC PJ - (3280)':
                    wait = WebDriverWait(driver2, 10)
                    wait.until(EC.visibility_of(link_modalidade))
                    
                    link_modalidade.click()
                    break

            elif str(tipo_simulacao).strip() == '2':
                span_modalidade = link_modalidade.find_element(By.TAG_NAME, 'span')

                if span_modalidade.text == 'PRODUCAO IMOVEL - REC SBPE/SFH - AQ TER E CONSTRUCAO - POS FIXADA - PF - HH 122 - (1976)':
                    wait = WebDriverWait(driver2, 10)
                    wait.until(EC.visibility_of(link_modalidade))

                    link_modalidade.click()
                    break
            
        time.sleep(5)
    except:
        driver2.quit()
        raise ValueError("Erro ao clicar no link")


    # Sistema de amortização e apólice
    try:
        if str(tipo_simulacao).strip() == '1':

            if str(dependente).strip() == '2':
                possui_dependente = driver2.find_element(By.NAME, "possuiMaisUmParticipante")
                possui_dependente.click()
                time.sleep(1)

            driver2.find_element(By.ID, 'areaUtil').send_keys('0')
            time.sleep(1)

        sistema_amortizacao = Select(driver2.find_element(By.NAME, "rcrRge"))

        if str(tipo_simulacao).strip() == '1':
            if str(amortizacao).strip() == '1':
                sistema_amortizacao.select_by_visible_text("894 - PRICE FGTS")
                time.sleep(1)
            elif str(amortizacao).strip() == '2':
                sistema_amortizacao.select_by_visible_text("893 - SAC FGTS")
                time.sleep(1)

        elif str(tipo_simulacao).strip() == '2':
            if str(amortizacao).strip() == '1':
                sistema_amortizacao.select_by_visible_text("794 - PRICE TR")
                time.sleep(1)
            elif str(amortizacao).strip() == '2':
                sistema_amortizacao.select_by_visible_text("793 - SAC TR")
                time.sleep(1)

        driver2.find_element(By.XPATH, '//*[@id="bottom_bar"]/fieldset/ul/li[2]/a').click()
        time.sleep(5)
    except:
        driver2.quit()
        raise ValueError("Erro no sistema de amortização e apólices")


    # Cronograma
    try:
        driver2.find_element(By.ID, 'prazoObra').send_keys('36')

        driver2.find_element(By.XPATH, '//*[@id="cronogramaForm"]/div[2]/fieldset/ul/li/a').click()
        time.sleep(3)

        driver2.find_element(By.XPATH, '//*[@id="bottom_bar"]/fieldset/ul/li[2]/a').click()
        time.sleep(5)
    except:
        driver2.quit()
        raise ValueError("Erro na página Cronograma")


    # Detalhamento
    try:
        if str(valor_financiamento).strip() != '' and str(prazo).strip() != '':
            driver2.find_element(By.XPATH, '//*[@id="detalhamentoForm"]/div[2]/fieldset/ul/li[1]/a').click()
            time.sleep(5)

            if str(valor_financiamento).strip() != '':
                valor_editado = str(valor).upper().replace('R$', '').replace('.', '').replace(',00', '').strip()
                valor_financiamento_editado = str(valor_financiamento).upper().replace('R$', '').replace('.', '').replace(',00', '').strip()
                
                valor_financiamento_minimo = driver2.find_element(By.XPATH, '//*[@id="form"]/fieldset/p').text
                valor_financiamento_minimo = valor_financiamento_minimo.split('Financiamento Mínimo:')
                if len(valor_financiamento_minimo) > 1:
                    valor_financiamento_minimo = valor_financiamento_minimo[1].strip()

                    valor_editado = str(valor).upper().replace('R$', '').replace('.', '').replace(',00', '').strip()
                    valor_financiamento_editado = str(valor_financiamento).upper().replace('R$', '').replace('.', '').replace(',00', '').strip()

                    valor_financiamento_minimo_editado = str(valor_financiamento_minimo).upper().replace('R$', '').replace('.', '').replace(',00', '').strip()

                    if int(valor_financiamento_editado) > int(valor_financiamento_minimo_editado) and int(valor_financiamento_editado) < int(valor_editado):
                        entrada = str(int(valor_editado) - int(valor_financiamento_editado)) + ',00' 
                    else:
                        entrada = str(int(valor_editado) - int(valor_financiamento_minimo_editado)) + ',00'
            
                driver2.find_element(By.NAME, "valorEntrada").clear()
                driver2.find_element(By.NAME, "valorEntrada").send_keys(entrada)
                time.sleep(1)

            if str(prazo).strip() != '':
                if str(prazo).strip() == '1':
                    prazo_cliente = '120'
                elif str(prazo).strip() == '2':
                    prazo_cliente = '150'
                elif str(prazo).strip() == '3':
                    prazo_cliente = '180'
                elif str(prazo).strip() == '4':
                    prazo_cliente = '240'
                elif str(prazo).strip() == '5':
                    prazo_cliente = '300'
                elif str(prazo).strip() == '6':
                    prazo_cliente = '360'
                elif str(prazo).strip() == '7':
                    prazo_cliente = '420'

                prazo = driver2.find_element(By.NAME, "prazo").get_attribute('value')
                if int(prazo) > int(prazo_cliente):
                    driver2.find_element(By.NAME, "prazo").clear()
                    driver2.find_element(By.NAME, "prazo").send_keys(prazo_cliente)
                    time.sleep(1)

            driver2.find_element(By.XPATH, '//*[@id="bottom_bar"]/fieldset/ul/li/a').click()
    except:
        driver2.quit()
        raise ValueError("Erro na página detalhamento")


    # Cria o pdf
    try:
        doc = SimpleDocTemplate("resultados/Simulação.pdf", pagesize=A4)
        elements = []
    except:
        driver2.quit()
        raise ValueError("Erro ao criar o pdf 1")

    # Cria o título da tabela
    try:
        dados = [['Simulação']]
        t = Table(dados)
        t.setStyle(TableStyle([('FONTSIZE', (0, 0), (0, 0), 20)]))
        elements.append(t)
        elements.append(Spacer(width=0, height=50))
    except:
        driver2.quit()
        raise ValueError("Erro ao criar título da tabela")
    

    # Itera as tabelas 
    try:
        tabelas = driver2.find_elements(By.TAG_NAME, 'table')
        for i in range(0,len(tabelas)-1):
            lista_tabela = []
            
            rows = tabelas[i].find_elements(By.TAG_NAME, "tr")
            for row in rows:

                lista_celula = []
                cols = row.find_elements(By.TAG_NAME, "th")

                for col in cols:
                    if col.text.replace(' ', '') == '':
                        continue
                
                    try:
                        lista_celula.append(Paragraph(col.text))
                    except:
                        continue

                cols = row.find_elements(By.TAG_NAME, "td")
                
                for col in cols:
                    if col.text.replace(' ', '') == '':
                        continue

                    try:
                        lista_celula.append(Paragraph(col.text))
                    except:
                        continue

                try:
                    if len(lista_celula) == 0:
                        continue
                    
                    lista_tabela.append(lista_celula)
                except:
                    continue
                
            
            t = Table(lista_tabela)
        
            t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 1), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 1), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, 0), 'LEFT'),
                            ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 10),
                            ('FONTSIZE', (0, 1), (-1, -1), 8),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#FDC463')),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
            
            elements.append(t)
            elements.append(Spacer(width=0, height=30))
            
    except:
        driver2.quit()
        raise ValueError("Erro ao iterar as tabelas")

    try:
        doc.build(elements)
        time.sleep(1)
    except:
        driver2.quit()
        raise ValueError("Erro ao salvar o pdf 1")
    
    

    # Inicia o segundo pdf
    try:                   
        driver2.find_element(By.XPATH, '/html/body/div[4]/form/div[2]/fieldset/div/table[1]/tbody/tr[2]/th[4]/div/ul/li/a').click()
        time.sleep(5)
    except:
        driver2.quit()
        raise ValueError("Erro ao clicar no segundo pdf 2")
    

    # Cria o seugundo pdf
    try:
        doc2 = SimpleDocTemplate("resultados/Simulação Planilha CET.pdf", pagesize=A4)
        elements2 = []
        lista_tabela = []
    except:
        driver2.quit()
        raise ValueError("Erro ao criar o segundo pdf")

    try:
        dados2 = [['Simulação Planilha CET']]
        t2 = Table(dados2)
        t2.setStyle(TableStyle([('FONTSIZE', (0, 0), (0, 0), 20)]))
        elements2.append(t2)
        elements2.append(Spacer(width=0, height=50))
    except:
        driver2.quit()
        raise ValueError("Erro ao criar o título do segundo pdf")

    try:
        lista_celula = []
        limite = 0
        titulo = True
        resumos = driver2.find_elements(By.TAG_NAME, 'p')
        for resumo in resumos:
            tags = resumo.find_elements(By.TAG_NAME, '*')

            for tag in tags:

                if tag.get_attribute("tagName") == 'LABEL' or tag.get_attribute("tagName") == 'SPAN':
                    if titulo == True:
                        lista_celula.append(Paragraph(tag.text))
                        lista_tabela.append(lista_celula)
                        lista_celula = []
                        titulo = False

                    else:
                        if limite <= 1:
                            lista_celula.append(Paragraph(tag.text))
                            limite += 1

                            if limite == 2:
                                lista_tabela.append(lista_celula)
                                lista_celula = []
                                limite = 0
    except:
        driver2.quit()
        raise ValueError("Erro ao criar a tabela 1 pdf 2")
    
    try:
        t2 = Table(lista_tabela)
        t2.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 1), colors.grey),
                        ('TEXTCOLOR', (0, 0), (-1, 1), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, 0), 'LEFT'),
                        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 10),
                        ('FONTSIZE', (0, 1), (-1, -1), 8),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#FDC463')),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
        
        elements2.append(t2)
        elements2.append(Spacer(width=0, height=30))
    except:
        driver2.quit()
        raise ValueError("Erro ao estilizar as tabelas do pdf 2")


    # Itera as outras tabelas
    try:
        tabelas2 = driver2.find_elements(By.TAG_NAME, 'table')
        for i in range(0,len(tabelas2)-1):
            lista_tabela = []
            
            rows = tabelas2[i].find_elements(By.TAG_NAME, "tr")
            for row in rows:

                # Para o restante da tabela 
                lista_celula = []

                cols = row.find_elements(By.TAG_NAME, "th")
                for col in cols:
                    
                    if col.text.replace(' ', '') == '':
                        continue
                
                    try:
                        lista_celula.append(Paragraph(col.text))
                    except:
                        continue

                cols = row.find_elements(By.TAG_NAME, "td")
                
                for col in cols:
                    if col.text.replace(' ', '') == '':
                        continue

                    try:
                        lista_celula.append(Paragraph(col.text))
                    except:
                        continue

                if len(lista_celula) == 0:
                    continue
                
                lista_tabela.append(lista_celula)
                
            t2 = Table(lista_tabela)
        
            t2.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 1), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 1), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, 0), 'LEFT'),
                            ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 10),
                            ('FONTSIZE', (0, 1), (-1, -1), 8),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#FDC463')),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
            
            elements2.append(t2)
            elements2.append(Spacer(width=0, height=30))       
    except:
        driver2.quit()
        raise ValueError("Erro ao iterar as tabelas do pdf 2")

    try:
        doc2.build(elements2)
        driver2.quit()
    except:
        driver2.quit()
        raise ValueError("Erro ao salvar o pdf 2")
    

    try:
        # Envia os arquivos
        driver.find_element(By.CLASS_NAME, 'i-iniciar').click()
        time.sleep(5)

        # Captura o caminho da pasta
        diretorio_atual = os.getcwd()
        arquivo1 = diretorio_atual + '/resultados/Simulação.pdf'
        arquivo2 = diretorio_atual + '/resultados/Simulação Planilha CET.pdf'
        time.sleep(1)
    except:
        raise ValueError("Erro ao iniciar os arquivos")

    try:
        anexar = driver.find_element(By.ID, 'fileInput')
        anexar.send_keys(arquivo1)
        time.sleep(5)
        driver.find_element(By.XPATH, '//*[@id="image-preview"]/div[2]/button[2]').click()
        time.sleep(10)
    
        anexar.send_keys(arquivo2)
        time.sleep(5)
        driver.find_element(By.XPATH, '//*[@id="image-preview"]/div[2]/button[2]').click()
        time.sleep(10)

        driver.find_element(By.CLASS_NAME, 'i-finalizar').click()
        time.sleep(10)
    except:
        raise ValueError("Erro ao enviar os anexos")
    
    # Remove o arquivo
    try:
        os.remove(arquivo1)
        os.remove(arquivo2)
    except:
        raise ValueError("Erro ao apagar os arquivos")

    

# Segundo Simulador
def simulador2(tipo_imovel, cidade, valor, fgts, dependente, tipo_simulacao, nascimento, renda, valor_financiamento, prazo, valor_parcela, amortizacao):

    # Teste de variáveis iniciais
    try:
        valor_inicial = int(str(valor).upper().replace('R$', '').replace('.', '').replace(',00', '').strip())
        renda_inicial = int(str(renda).upper().replace('R$', '').replace('.', '').replace(',00', '').strip())
        
        if valor_financiamento != "":
            valor_financiamento_inicial = int(str(valor_financiamento).upper().replace('R$', '').replace('.', '').replace(',00', '').strip())
    except:
        raise ValueError("Erro ao manipular as variáveis iniciais")
    
    if str(tipo_simulacao).strip() == '1':
            
            if valor_inicial > 264000:
                    raise ValueError("Valor do financiamento maior que o limite")

            elif renda_inicial > 8000:
                    raise ValueError("Valor da renda maior que o limite")
            
    if valor_financiamento != "":
        if valor_financiamento_inicial > valor_inicial:
                raise ValueError("Valor do financiamento maior que o valor do imóvel")
     

    # Abre a planilha login na aba caixa
    try:
        df_caixa = pd.read_excel('arquivos/login.xlsx', sheet_name='caixa') 
    except:
        print('Erro ao abrir a planilha login na aba caixa')
        time.sleep(30)
        sys.exit()


    # Pega os dados de login e inicia a segunda janela
    try:
        site_caixa = 'https://habitacao.caixa.gov.br/siopiweb-web'
        usuario_caixa = df_caixa['usuario'][0]
        senha_caixa = df_caixa['senha'][0]

        options2 = Options()
        options2.add_argument("--profile-directory=Default")
        options2.add_argument(f"--user-data-dir={caminho_pasta_atual}/cookies2")

        driver2 = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options2)
        driver2.maximize_window()

        driver2.get(site_caixa)
        time.sleep(5)
    except:
        driver2.quit()
        raise ValueError("Simulador 2 - Erro ao abrir a segunda janela")


    # Faz o login
    try:
        driver2.find_element(By.ID, 'username').send_keys(usuario_caixa)
        time.sleep(1)
        driver2.find_element(By.ID, 'password').send_keys(senha_caixa)
        time.sleep(1)
        driver2.find_element(By.ID, 'btn_login').click()
        time.sleep(5)
    except:
        pass


    # Primeira página
    try:
        iframe = driver2.find_element(By.XPATH, '/html/body/iframe')
        driver2.switch_to.frame(iframe)
        time.sleep(1)

        driver2.find_element(By.ID, 'btn_menu').click()
        time.sleep(2)

        driver2.find_element(By.ID, 'id_menu_img_servicos').click()
        time.sleep(2)

        driver2.find_element(By.ID, 'HM_Item1_2').click()
        time.sleep(3)
    except:
        driver2.quit()
        raise ValueError("Simulador 2 - Erro ao acessar primeira página")


    # Identificação
    try:
        driver2.find_element(By.ID, "nuTipoFinanciamento_input").clear()
        driver2.find_element(By.ID, "nuTipoFinanciamento_input").send_keys('Residencial')
        action = ActionChains(driver2)
        action.send_keys("\ue004").perform()
        time.sleep(1)

        driver2.find_element(By.ID, "nuCategoria_input").clear()
        if str(tipo_imovel).strip() == '2':
            driver2.find_element(By.ID, "nuCategoria_input").send_keys('Aquisição de Imóvel Novo')
        elif str(tipo_imovel).strip() == '3':
            driver2.find_element(By.ID, "nuCategoria_input").send_keys('Aquisição de Imóvel Usado')
        action = ActionChains(driver2)
        action.send_keys("\ue004").perform()
        time.sleep(1)

        cidade_editada = str(cidade).strip().upper()
        if not ' - ' in cidade_editada:
            cidade_editada = str(cidade_editada).replace('-', ' - ')
        cidade_sem_acento = unidecode(cidade_editada)
        driver2.find_element(By.ID, "txtNomeAutocomplete").send_keys(cidade_sem_acento[:-1])
        time.sleep(2)
        action = ActionChains(driver2)
        action.send_keys("\ue004").perform()
        time.sleep(1)

        valor_editado = str(valor).upper().replace('R$', '').strip()
        if not ',' in valor_editado:
            valor_editado = valor_editado + ',00'
        driver2.find_element(By.ID, "vlrCompraVendaOrcamento").send_keys(valor_editado)
        time.sleep(1)
        driver2.find_element(By.ID, "vlrDaAvaliacao").send_keys(valor_editado)
        time.sleep(1)

        if str(tipo_simulacao).strip() == '1':
            if str(fgts).strip() == '1':
                driver2.find_element(By.NAME, "icPossuiContaFGTS").click()
                time.sleep(1)

            if str(dependente).strip() == '2':
                driver2.find_element(By.NAME, "icFatorSocial").click()
                time.sleep(1)


        nascimento_editado = str(nascimento).replace(' ', '')
        driver2.find_element(By.ID, "dtNascParticipante1").send_keys(nascimento_editado)
        time.sleep(1)

        renda_editada = str(renda).upper().replace('R$', '').strip()
        if not ',' in renda_editada:
            renda_editada = renda_editada + ',00'
        driver2.find_element(By.ID, "rendaParticipante1").send_keys(renda_editada)
        time.sleep(1)

        driver2.find_element(By.XPATH, '//*[@id="acoes_nav_inferior"]/a[2]').click()
        time.sleep(5)
    except:
        driver2.quit()
        raise ValueError("Simulador 2 - Erro na página identificação")


    # Produtos
    try:
        if str(tipo_simulacao).strip() == '1':
            driver2.find_element(By.XPATH, '//*[@id="tab_identificacao"]/table/tbody/tr[1]/th[2]/a').click()
            time.sleep(5)
        elif str(tipo_simulacao).strip() == '2':
            driver2.find_element(By.XPATH, '//*[@id="tab_identificacao"]/table/tbody/tr[5]/th[2]/a').click()
            time.sleep(5)
    except:
        driver2.quit()
        raise ValueError("Simulador 2 - Erro na página produtos")


    # Condições
    try:
        if valor_financiamento != '':
            valor_editado = str(valor).upper().replace('R$', '').replace('.', '').replace(',00', '').strip()
            valor_financiamento_editado = str(valor_financiamento).upper().replace('R$', '').replace('.', '').replace(',00', '').strip()

            entrada = str(int(valor_editado) - int(valor_financiamento_editado)) + ',00'

            driver2.find_element(By.ID, "vlrEntrada").send_keys(entrada)
            time.sleep(1)

        if prazo != '':
            if prazo == '1':
                prazo_cliente = '120'
            elif prazo == '2':
                prazo_cliente = '150'
            elif prazo == '3':
                prazo_cliente = '180'
            elif prazo == '4':
                prazo_cliente = '240'
            elif prazo == '5':
                prazo_cliente = '300'
            elif prazo == '6':
                prazo_cliente = '360'
            elif prazo == '7':
                prazo_cliente = '420'

            driver2.find_element(By.ID, "prazo").send_keys(prazo_cliente)
            time.sleep(1)

        if valor_parcela != '':
            valor_parcela_editada = str(valor_parcela).upper().replace('R$', '').strip()
            if not ',' in valor_parcela:
                valor_parcela_editada = valor_parcela_editada + ',00'
            driver2.find_element(By.ID, "vlrPrestacaoMaximaSIRIC").send_keys(valor_parcela_editada)
            time.sleep(1)
        
        driver2.find_element(By.ID, "nuSistemaAmortizacao_input").clear()
        if str(amortizacao).strip() == '1':
            driver2.find_element(By.ID, "nuSistemaAmortizacao_input").send_keys("PRICE")
        elif str(amortizacao).strip() == '2':
            driver2.find_element(By.ID, "nuSistemaAmortizacao_input").send_keys("SAC")
        time.sleep(2)
        action = ActionChains(driver2)
        action.send_keys("\ue004").perform()
        time.sleep(1)

        driver2.find_element(By.XPATH, '//*[@id="acoes_nav_inferior"]/a[3]').click()
        time.sleep(5)
    except:
        driver2.quit()
        raise ValueError("Simulador 2 - Erro na página condições")


    # Resultado
    try:
        erro = driver2.find_element(By.CLASS_NAME, 'msg_erro').text
        raise ValueError(erro)
    except:
        pass
    

    # Cria o pdf
    try:
        doc = SimpleDocTemplate("resultados/Simulação.pdf", pagesize=A4)
        elements = []

        tabelas = driver2.find_elements(By.TAG_NAME, 'table')

        dados = [['Simulação']]
        t = Table(dados)
        t.setStyle(TableStyle([('FONTSIZE', (0, 0), (0, 0), 20)]))
        elements.append(t)
        elements.append(Spacer(width=0, height=50))
    except:
        driver2.quit()
        raise ValueError("Simulador 2 - Erro ao criar o pdf")
    

    # Itera as outras tabelas da segunda até a penúltima
    try:
        for i in range(0,len(tabelas)):
            lista_tabela = []
            j = 0
            cabecalho = 0
            
            rows = tabelas[i].find_elements(By.TAG_NAME, "tr")
            for row in rows:

                lista_celula = []

                cols = row.find_elements(By.TAG_NAME, "th")

                for col in cols:
                    
                    if col.text.replace(' ', '') == '':
                        continue
                
                    try:
                        lista_celula.append(Paragraph(col.text))
                    except:
                        logging.info('Erro ao iterar os dados das tabelas')
                        continue

                cols = row.find_elements(By.TAG_NAME, "td")
                
                for col in cols:
                    if col.text.replace(' ', '') == '':
                        continue

                    try:
                        lista_celula.append(Paragraph(col.text))
                    except:
                        logging.info('Erro ao iterar os dados das tabelas')
                        continue

                if len(lista_celula) == 0:
                    continue
                
                lista_tabela.append(lista_celula)
                
            try:
                t = Table(lista_tabela)
            
                t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 1), colors.grey),
                                ('TEXTCOLOR', (0, 0), (-1, 1), colors.whitesmoke),
                                ('ALIGN', (0, 0), (-1, 0), 'LEFT'),
                                ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 10),
                                ('FONTSIZE', (0, 1), (-1, -1), 8),
                                ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#FDC463')),
                                ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
                
                elements.append(t)
                elements.append(Spacer(width=0, height=30))
            except:
                continue
    except:
        driver2.quit()
        raise ValueError("Simulador 2 - Erro na iteração das tabelas")


    # Salva o pdf
    try:
        doc.build(elements)
        driver2.quit()
    except:
        driver2.quit()
        raise ValueError("Simulador 2 - Erro na criação do pdf")
    

    # Inicia o envio do arquivo
    try:
        driver.find_element(By.CLASS_NAME, 'i-iniciar').click()
        time.sleep(5)

        # Captura o caminho da pasta
        diretorio_atual = os.getcwd()
        arquivo = diretorio_atual + '/resultados/Simulação.pdf'
        time.sleep(1)
    except:
        raise ValueError("Simulador 2 - Erro ao iniciar o arquivo")


    # Envia o arquivo
    try:
        anexar = driver.find_element(By.ID, 'fileInput')
        anexar.send_keys(arquivo)
        time.sleep(5)
        driver.find_element(By.XPATH, '//*[@id="image-preview"]/div[2]/button[2]').click()
        time.sleep(10)
    
        driver.find_element(By.CLASS_NAME, 'i-finalizar').click()
        time.sleep(10)
    except:
        raise ValueError("Simulador 2 - Erro ao enviar o anexo")
    

    # Remove o arquivo
    try:
        os.remove(arquivo)
    except:
        raise ValueError("Simulador 2 - Erro ao apagar arquivo")



# Abre a planilha login na aba mzworkspace
try:
    df_mzworkspace = pd.read_excel('arquivos/login.xlsx', sheet_name='mzworkspace') 
except:
    print('Erro ao abrir a planilha login na aba mzworkspace')
    time.sleep(30)
    sys.exit()



# Início do programa
try:
    site_login = 'https://app.mzworkspace.com/'
    usuario = df_mzworkspace['usuario'][0]
    senha = df_mzworkspace['senha'][0]
    caminho_pasta_atual = os.getcwd()
except:
    print("Erro ao pegar dados do login")
    time.sleep(30)
    sys.exit()


try:
    options = Options()
    options.add_argument("--profile-directory=Default")
    options.add_argument(f"--user-data-dir={caminho_pasta_atual}/cookies")

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    driver.maximize_window()

    driver.get(site_login)
    time.sleep(10)
except:
    print("Erro ao abrir navegador")
    time.sleep(30)
    sys.exit()

try:
    # Faz o login
    driver.find_element(By.ID, 'email').send_keys(usuario)
    time.sleep(1)
    driver.find_element(By.ID, 'password').send_keys(senha)
    time.sleep(1)
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/form/div/div[5]/button').click()
    time.sleep(20)
except:
    pass

try:
    driver.find_element(By.CLASS_NAME, 'btn-link').click()
    time.sleep(2)
except:
    pass


# Verifica se o usuário está online
try:
    driver.find_element(By.CLASS_NAME, 'mz-header-user-image').click()
    time.sleep(1)
    status = driver.find_element(By.XPATH, '/html/body/data/header/div/div[2]/div/ul/li[3]/a/span').text
    if str(status).strip() != 'Ficar Offline':
        driver.find_element(By.XPATH, '/html/body/data/header/div/div[2]/div/ul/li[3]/a/span').click()
    time.sleep(3)
except:
    print("Erro ao verificar se o usuário está online")
    time.sleep(30)
    sys.exit()


# Looping infinito
while True:
    try:
        driver.find_element(By.CLASS_NAME, 'tab-chat-pendentes').click()
        time.sleep(2)

        pendentes = driver.find_elements(By.CLASS_NAME, 'atendimento-item')
    except:
        print("Erro ao buscar chats pendentes")
        time.sleep(10)
        continue

    for pendente in pendentes:

        try:  
            pendente
        except:
            print('Erro na conversa pendente')
            continue 

        ultima_conversa = ''
        opcao = ''

        tipo_imovel = cidade = valor = renda = nascimento = prazo_financiamento = prazo = ''
        informar_valor_financiamento = valor_financiamento = valor_parcela = amortizacao = ''
        informar_valor_parcela = tipo_simulacao = dependente = fgts = ''

        valor_deseja = False

        try:
            pendente.click()
            time.sleep(5)
        except:
            continue

        try:
            tela = driver.find_element(By.XPATH, '//*[@id="atendimento-mensagens"]')
            time.sleep(1)

            # execute o script JavaScript para rolar a div
            driver.execute_script("arguments[0].scrollTop = -50;", tela)
        except:
            continue
        time.sleep(5)

        try:
            # execute o script JavaScript para rolar a div
            driver.execute_script("arguments[0].scrollTop = 5;", tela)
        except:
            continue
        time.sleep(3)

        try:
            # execute o script JavaScript para rolar a div
            driver.execute_script("arguments[0].scrollTop = -50;", tela)
        except:
            continue
        time.sleep(5)

        try:
            mensagens = driver.find_elements(By.CLASS_NAME, 'base')
        except:
            print("Erro ao buscar as mensagens")
            time.sleep(5)
            continue

        

        # Verifica onde começa a última conversa
        for i in reversed(range(len(mensagens))):

            try:
                divs = mensagens[i].find_elements(By.TAG_NAME, 'div')
                divs[3].text
            except:
                continue

            if "Eu sou a Kredy" in divs[3].text:             
                ultima_conversa = i
                break
        
        
        # Percorre a última conversa
        try:
            for i in range(ultima_conversa, len(mensagens)-1):
                resposta = ''

                divs = mensagens[i].find_elements(By.TAG_NAME, 'div')

                # Primeira Mensagem
                if "Eu sou a Kredy" in divs[3].text:
                    resposta = mensagens[i+1].find_elements(By.TAG_NAME, 'div')         
                    
                    if resposta[3].text == '2':
                        opcao = '2'
                        
                    else:
                        break

                # Tipo de imóvel
                elif divs[3].text.strip() == "Para qual tipo de imóvel? Envie uma das seguintes opções":
                    j = 2
                    while tipo_imovel not in ["1", "2", "3"]:
                        resposta = mensagens[i+j].find_elements(By.TAG_NAME, 'div')

                        tipo_imovel = resposta[3].text.strip()

                        j += 2

                
                # Cidade
                elif divs[3].text.strip() == "Qual cidade está localizado o imóvel? ex: São Paulo – SP":
                    resposta = mensagens[i+1].find_elements(By.TAG_NAME, 'div')

                    cidade = resposta[3].text


                # Valor
                elif divs[3].text.strip() == "Qual valor de compra e venda do imóvel?":
                    resposta = mensagens[i+1].find_elements(By.TAG_NAME, 'div')

                    valor = resposta[3].text


                # Renda Familiar
                elif divs[3].text.strip() == "Qual a renda Bruta familiar?":
                    resposta = mensagens[i+1].find_elements(By.TAG_NAME, 'div')

                    renda = resposta[3].text


                # Nascimento
                elif divs[3].text.strip() == "Informe a data de nascimento do participante mais velho?":
                    resposta = mensagens[i+1].find_elements(By.TAG_NAME, 'div')

                    nascimento = resposta[3].text


                # Valor do financiamento
                elif divs[3].text.strip() == "Deseja informar o prazo de financiamento?":
                    resposta = mensagens[i+2].find_elements(By.TAG_NAME, 'div')

                    prazo_financiamento = resposta[3].text


                # Prazo
                elif divs[3].text.strip() == "Qual prazo deseja?":
                    resposta = mensagens[i+2].find_elements(By.TAG_NAME, 'div')

                    prazo = resposta[3].text


                # Informar Valor do Financiamento
                elif divs[3].text.strip() == "Deseja informar o valor de financiamento?":
                    resposta = mensagens[i+2].find_elements(By.TAG_NAME, 'div')

                    informar_valor_financiamento = resposta[3].text


                # Valor do Financiamento ou Valor da parcela
                elif divs[3].text.strip() == "Qual Valor deseja?":
                    resposta = mensagens[i+1].find_elements(By.TAG_NAME, 'div')

                    if valor_deseja == False:
                        valor_financiamento = resposta[3].text
                        valor_deseja = True
                    else:
                        valor_parcela = resposta[3].text


                # Amortização
                elif divs[3].text.strip() == "Qual sistema de amortização você quer?":
                    resposta = mensagens[i+2].find_elements(By.TAG_NAME, 'div')

                    amortizacao = resposta[3].text


                # Informar Valor da parcela
                elif divs[3].text.strip() == "Deseja informar o valor da parcela?":
                    resposta = mensagens[i+2].find_elements(By.TAG_NAME, 'div')

                    informar_valor_parcela = resposta[3].text


                # Valor da parcela está mais acima junto com valor do financiamento


                # Tipo de simulação
                elif divs[3].text.strip() == "Qual tipo de simulação você quer?":
                    resposta = mensagens[i+2].find_elements(By.TAG_NAME, 'div')

                    tipo_simulacao = resposta[3].text


                # Dependente
                elif divs[3].text.strip() == "Possui mais de um comprador ou dependentes?":
                    resposta = mensagens[i+2].find_elements(By.TAG_NAME, 'div')

                    dependente = resposta[3].text
                    
                
                # FGTS
                elif divs[3].text.strip() == "Qualquer um dos participantes possuem 36 meses de trabalho com regime de FGTS?":
                    resposta = mensagens[i+2].find_elements(By.TAG_NAME, 'div')

                    fgts = resposta[3].text
        except:
            continue

        print(f'tipo_imovel: {tipo_imovel}')
        print(f'cidade: {cidade}')
        print(f'valor: {valor}')
        print(f'renda: {renda}')
        print(f'nascimento: {nascimento}')
        print(f'prazo_financiamento: {prazo_financiamento}')
        print(f'prazo: {prazo}')

        print(f'informar_valor_financiamento: {informar_valor_financiamento}')
        print(f'valor_financiamento: {valor_financiamento}')
        print(f'valor_parcela: {valor_parcela}')
        print(f'amortizacao: {amortizacao}')

        print(f'informar_valor_parcela: {informar_valor_parcela}')
        print(f'tipo_simulacao: {tipo_simulacao}')
        print(f'dependente: {dependente}')
        print(f'fgts: {fgts}')

        print(f'valor_deseja: {valor_deseja}')


        if str(tipo_imovel).strip() == '1':
            try:
                simulador1(tipo_simulacao, cidade, valor, renda, fgts, nascimento, dependente, amortizacao, prazo, valor_financiamento)
            except ValueError as error:
                print(error)
                time.sleep(30)
                continue

        elif str(tipo_imovel).strip() == '2' or str(tipo_imovel).strip() == '3':
            try:
                simulador2(tipo_imovel, cidade, valor, fgts, dependente, tipo_simulacao, nascimento, renda, valor_financiamento, prazo, valor_parcela, amortizacao)
            except ValueError as error:
                print(error)
                time.sleep(30)
                continue

