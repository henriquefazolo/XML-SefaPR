# coding: utf-8
# version : 0.2
# author : Henrique UM Fazolo
# e-mail : henriquefazolo@gmail.com

# 01/01/2021    alterado para firefox
# 08/02/2021    inicio da construção da interface grafica
# 09/02/2021    inicio da construção: comunicação entre a parte grafica e a logica, testes para realizar login no site
#               abre até a tela de login
# 11/02/2021    Alterado para IE
#               Termino da contrução da parte grafica
#               construir logica para tratamento de eventuais erro                              
# 18/02/2021    FUNCIONANDO        
#               ERRO: Não é possivel abrir um segundo arquivo txt.
#               CORRIGIDO ERRO 
#               INICIO DO TRATAMENTO DE ERROS
# 22/02/2021    O sistema trata espaços e linhas em branco do arquivo txt
# 23/02/2021    Criado listbox das empresas do txt
#               Função para selecionar empresa para iniciar o trabalho do robô: concluido

# imports

## parte grafica
import PySimpleGUI as sg

## parte logica
from selenium import webdriver

# Firefox = Talvez nao sera utilizado, esta sendo usado IE
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary

# classes

## classe grafica
class Interface_Grafica():

    # Criação do Layout
    def __init__(self):

        self.layout = [
            
            [sg.Text('Arquivo', size=(5,0)), sg.Input(key='-ARQUIVO-', enable_events=True),
            sg.FileBrowse('Selecione', initial_folder='*\\', file_types=(('Arquivos de Texto', '*.txt'),))], #FileBrowser aberto inicialmente na pasta raiz
            
            [sg.Text('Login', size=(5,0)), sg.Input(key='-LOGIN-')],
            
            [sg.Text('Senha', size=(5,0)), sg.Input(password_char='*', key='-SENHA-')],
            
            [sg.CalendarButton('Data inicial', size=(15,0), target=('-INPUT_DATA_INICIAL-'), format=('%d/%m/%Y')),
             sg.Input('',key='-INPUT_DATA_INICIAL-', enable_events=True, visible=False, ),
             sg.Txt('               ',key='-TXT_DATA_INICIAL-')
             ],

            [sg.CalendarButton('Data final', size=(15,0), target=('-INPUT_DATA_FINAL-'), format=('%d/%m/%Y')),
             sg.Input('',key='-INPUT_DATA_FINAL-', enable_events=True, visible=False, ),
             sg.Txt('', size=(30,1),key='-TXT_DATA_FINAL-')
             ],
            
            [sg.Listbox(values=(''), size=(50,5), key='-LISTBOX-', enable_events=True)],  
            [sg.Text('          ', key='-TEXTO_INICIO-') ,sg.Text('                                ', key='-EMPRESA_INICIAL-')],         
            [sg.Checkbox('Confirmar', key='-CHECKBOX-'), sg.Button('Solicitar')],

            # Criar caixa de texto output
            [sg.Output((50,20), key='-OUTPUT-')],
            
            [sg.Button('Sair')],
            ]

        # Criação da Janela

        self.window = sg.Window('XML SefaPR', self.layout)
        self.posicao_inicial_empresa = 0

    def iniciar_robo_sefapr(self):
        logica = Robo_SefaPR()
        logica.arquivo = str(self.values['-ARQUIVO-'])
        logica.login = str(self.values['-LOGIN-'])
        logica.senha = str(self.values['-SENHA-'])
        logica.data_inicial = str(self.values['-INPUT_DATA_INICIAL-'])
        logica.data_final = (self.values['-INPUT_DATA_FINAL-'])
        logica.posicao_cnpj_empresa_inicial = self.posicao_inicial_empresa
        logica.executar()    
        logica = ''
    
    def gerar_listbox_empresas(self, arquivo):
        listbox_empresas = []
        with open(arquivo) as arquivo:
            dados_empresa = arquivo.readlines()
            dados_tratados = []
            for dados in dados_empresa:    
                
                if len(dados.replace(' ', '')) >= 18:
                    dados_tratados.append(dados.replace(' ', ''))  
                    
            for dados_empresa in dados_tratados:
                nome_empresa = dados_empresa.split(':')[0]
                cnpj_empresa = dados_empresa.split(':')[1]
                listbox_empresas.append((nome_empresa, cnpj_empresa))
            
            return(listbox_empresas)
    
    def posicao_inicial(self, listbox_empresas, cnpj_inicial):
        nova_lista = []        
        for n in range(len(listbox_empresas)):
            cnpj_empresa_list_box = listbox_empresas[n][1][:-1]
            if cnpj_empresa_list_box == cnpj_inicial:
                posicao_inicial = n    
        return(posicao_inicial)
          
    def executar(self):

        while True:       
            try:     
                self.event, self.values = self.window.read()
                # Atualizações da interface

                self.window['-TXT_DATA_INICIAL-'].update(self.values['-INPUT_DATA_INICIAL-'])
                self.window['-TXT_DATA_FINAL-'].update(self.values['-INPUT_DATA_FINAL-'])

                self.window.Finalize()
            except:
                continue
                
            # Eventos dos botões
            if self.event == 'Sair' or self.event == sg.WINDOW_CLOSED:
                self.window.close()
                break

            if self.event == 'Solicitar':
                if self.values['-ARQUIVO-'] == '' or self.values['-LOGIN-'] == '' or self.values['-SENHA-'] == '' or \
                    self.values['-INPUT_DATA_INICIAL-'] == '' or self.values['-INPUT_DATA_FINAL-'] == '':

                    print('Há campos vazios.')
                else:                  
                    self.iniciar_robo_sefapr()
            
            if self.window.FindElement('-ARQUIVO-'):
                try:
                    lista_empresas = self.gerar_listbox_empresas(arquivo=self.values['-ARQUIVO-'])
                    self.window.FindElement('-LISTBOX-').update(values=(lista_empresas))
                    #self.window['-EMPRESA_INICIAL-'].update()
                    #print(self.values['-LISTBOX-'].values())
                except:
                    continue       
            
            if self.event == '-LISTBOX-' and len(self.values['-LISTBOX-']):
                empresa_inicial = self.values['-LISTBOX-']
                self.cnpj_empresa_inicial = empresa_inicial[0][1][:-1] 
                self.posicao_inicial_empresa = self.posicao_inicial(lista_empresas, self.cnpj_empresa_inicial)
                self.window['-EMPRESA_INICIAL-'].update(empresa_inicial[0][0])
                self.window['-TEXTO_INICIO-'].update('Inicio :')
      

## classe logica

class Robo_SefaPR():
    
    def __init__(self):
        self.__arquivo = ''
        self.__login = ''
        self.__senha = ''
        self.__data_inicial = ''
        self.__data_final = ''
        self.__site_receita_pr_login = 'https://receita.pr.gov.br/login'
        self.__site_receita_pr_solicitar_xml = 'https://www.dfeportal.fazenda.pr.gov.br/dfe-portal/manterDownloadDFe.do?action=iniciar'
        self.__tempo_espera = 10     
        self.__posicao_cnpj_empresa_inicial = 0         

    @property
    def arquivo(self):
        return self.__arquivo        
    @arquivo.setter
    def arquivo(self, arquivo):
        self.__arquivo = arquivo

    @property
    def login(self):
        return self.__login
    @login.setter
    def login(self, login):
        self.__login = login

    @property
    def senha(self):
        return self.__senha
    @senha.setter
    def senha(self, senha):
        self.__senha = senha

    @property
    def data_inicial(self):
        return self.__data_inicial
    @data_inicial.setter
    def data_inicial(self, data_inicial):
        self.__data_inicial =  data_inicial

    @property
    def data_final(self):
        return self.__data_final
    @data_final.setter
    def data_final(self, data_final):
        self.__data_final = data_final   
    
    @property
    def posicao_cnpj_empresa_inicial(self):
        return self.__posicao_cnpj_empresa_inicial
    @posicao_cnpj_empresa_inicial.setter
    def posicao_cnpj_empresa_inicial(self, posicao_cnpj_empresa_inicial):
        self.__posicao_cnpj_empresa_inicial = posicao_cnpj_empresa_inicial

    def navegador(self):
        self.navegador = webdriver.Ie("IEDriverServer.exe") #WebDriver na pasta raiz do programa
        self.navegador.set_window_size(800,600)
        return self.navegador

    def localizar_elemento(self, xpath):
        self.elemento = self.navegador.find_element_by_xpath(xpath)
        return self.elemento
     
    def tela_login(self, navegador):
        navegador.get(self.__site_receita_pr_login)

        campo_cpf = self.localizar_elemento('//*[@id="cpfusuario"]')
        campo_cpf.send_keys(self.login)

        campo_senha = self.localizar_elemento('/html/body/div[2]/form[1]/div[3]/div/input')
        campo_senha.send_keys(self.senha)

        botao_login = self.localizar_elemento('/html/body/div[2]/form[1]/div[4]/button')
        botao_login.click()


    def tela_solicitar_xml(self, navegador):
        navegador.get(self.__site_receita_pr_solicitar_xml)

    def solicitar_notas(self, nome_empresa, cnpj_empresa, data_inicial, data_final):
        def agendar_notas_emitidas(self):
            sleep(2)
            botao_tipo_notas_emitidas = self.localizar_elemento('//*[@id="ext-gen1112"]')
            botao_tipo_notas_emitidas.click() 

            campo_input_cnpj = self.localizar_elemento('//*[@id="ext-gen1081"]')
            campo_input_cnpj.clear()
            campo_input_cnpj.send_keys(cnpj_empresa.replace(' ', ''))

            campo_data_periodo_inicial = self.localizar_elemento('//*[@id="ext-gen1030"]')
            campo_data_periodo_inicial.clear()
            campo_data_periodo_inicial.send_keys(data_inicial)

            campo_data_periodo_final = self.localizar_elemento('//*[@id="ext-gen1032"]')
            campo_data_periodo_final.clear()
            campo_data_periodo_final.send_keys(data_final)

            campo_hora_periodo_inicial = self.localizar_elemento('//*[@id="ext-gen1022"]')
            campo_hora_periodo_inicial.clear()
            campo_hora_periodo_inicial.send_keys('00:00:00')

            campo_hora_periodo_final = self.localizar_elemento('//*[@id="ext-gen1023"]')
            campo_hora_periodo_final.clear()
            campo_hora_periodo_final.send_keys('23:59:59')

            sleep(1)        

            if interface.values['-CHECKBOX-'] == True:
                botao_agendar = self.localizar_elemento('//*[@id="ucs20_BtnAgendar-btnEl"]')
                botao_agendar.click()
                print(nome_empresa, 'Notas Emitidas - OK')
            else:
                print(nome_empresa, 'Notas Emitidas - Teste com sucesso')
            sleep(1)
        
        def agendar_notas_destinadas(self):
            sleep(2)
            botao_tipo_notas_destinadas = self.localizar_elemento('//*[@id="ext-gen1116"]')
            botao_tipo_notas_destinadas.click()

            campo_input_cnpj = self.localizar_elemento('//*[@id="ext-gen1081"]')
            campo_input_cnpj.clear()
            campo_input_cnpj.send_keys(cnpj_empresa.replace(' ', ''))

            campo_data_periodo_inicial = self.localizar_elemento('//*[@id="ext-gen1030"]')
            campo_data_periodo_inicial.clear()
            campo_data_periodo_inicial.send_keys(data_inicial)

            campo_data_periodo_final = self.localizar_elemento('//*[@id="ext-gen1032"]')
            campo_data_periodo_final.clear()
            campo_data_periodo_final.send_keys(data_final)

            campo_hora_periodo_inicial = self.localizar_elemento('//*[@id="ext-gen1022"]')
            campo_hora_periodo_inicial.clear()
            campo_hora_periodo_inicial.send_keys('00:00:00')

            campo_hora_periodo_final = self.localizar_elemento('//*[@id="ext-gen1023"]')
            campo_hora_periodo_final.clear()
            campo_hora_periodo_final.send_keys('23:59:59')
      
            if interface.values['-CHECKBOX-'] == True:
                botao_agendar = self.localizar_elemento('//*[@id="ucs20_BtnAgendar-btnEl"]')
                botao_agendar.click()
                print(nome_empresa, 'Notas Destinadas - OK')
            else:
                print(nome_empresa, 'Notas Destinadas - Teste com sucesso')
            sleep(1)
        
        agendar_notas_emitidas(self)
        agendar_notas_destinadas(self)

    def percorrer_lista_empresas(self):
        with open(self.arquivo) as arquivo:
            dados_empresa = arquivo.readlines()
            dados_tratados = []
            for dados in dados_empresa:    
                
                if len(dados.replace(' ', '')) >= 18:
                    dados_tratados.append(dados.replace(' ', ''))  

            del dados_tratados[:self.posicao_cnpj_empresa_inicial]
            for dados_empresa in dados_tratados:
                nome_empresa = dados_empresa.split(':')[0]
                cnpj_empresa = dados_empresa.split(':')[1]
                self.solicitar_notas(nome_empresa, cnpj_empresa,self.data_inicial, self.data_final)        

    def executar(self):
        try:
            self.navegador() #abre o navegador
            self.tela_login(self.navegador) # realiza o login
            self.tela_solicitar_xml(self.navegador) # vai até a pagina onde solicita o xml
            self.percorrer_lista_empresas()
            self.navegador.quit()
            print('\nFim da execução.')
        except Exception as erro:
            self.navegador.quit()            
            print('---------------------------')
            print(erro)
            print('---------------------------')

            if str(erro).find('[@id="ext-gen1081"]'):
                print('Erro de login.')
            if str(erro).find('file or directory.'):
                print('Erro no arquivo txt')
# execução

logica = ''
interface = Interface_Grafica()
interface.executar()
