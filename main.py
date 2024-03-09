import re # procurar e variações de REg exs
import os
import smtplib
import openpyxl
import pywhatkit
import email
from email.message import EmailMessage
from time import sleep
import pandas as pd
import pywhatkit as kit





class toDO:

# INICIAR 
    def iniciar (self):
          self.lista_tarefas = []
          self.email_destino()
          self.menu()
          self.Criar_Planilha()
          sleep(3)
          self.Enviar_Email()
          self.Enviar_whats()

# VALIDAÇÃO DE EMAIL:
    def email_destino(self):
        while True:
            self.email = str(input('Email de destino : ')). lower()  #reg exr ; validador de email  (site em favoritos :  stack overflow) entrar no site e colocar como favorito

            padrão_email = re.search(
            '^[a-z0-9._]+@[a-z0-9]+.[a-z]+(.[a-z]+)?$', self.email)

            if padrão_email:
                print('Email válido!!!')
                break
            else:
                print( 'Email inválido, tente outro ...')
# MENU:
    def menu (self): 

        while True:
            menu_principal  = int(input("""
            MENU PRINCIPAL
            [1] CADASTRAR
            [2] VISUALIZAR
            [3] SAIR
            Opção: """))
                 
            match menu_principal:
                case 1: self.cadastrar()
                case 2: self.Visualizar()
                case 3: break
                case _: print('\n Opção Inválida!')
# CADASTRO
    def cadastrar(self):
          while True:
                self.tarefa = str(input('Digite uma tarefa ou [S] para sair:')). capitalize()
                if self.tarefa == 'S':
                      break
                else: 
                    self.lista_tarefas.append(self.tarefa)
                    try:
                       with open('./SRC/Tarefas/Historico_Tarefas.txt', 'a', encoding='utf8') as arquivo:
                            arquivo.write(f'{self.tarefa}\n')

                    except FileNotFoundError as e:
                         print(f' \n ERRO: {e}')
#VISUALIZAÇÃO
    def Visualizar(self):
        try:
                with open('./SRC/Tarefas/Historico_Tarefas.txt', 'r', encoding= 'utf8') as arquivo:
                   print(arquivo.read())

        except FileNotFoundError as e:
            print(f'Erro: {e}')
#Criar _ pLanilha
    def Criar_Planilha(self):
        if len(self.lista_tarefas) > 0  :
            try :
                df = pd.DataFrame({"Tarefas": self.lista_tarefas})

                self.nome_arquivo = str(input('Nome do arquivo:')).lower()

                if self.nome_arquivo[-5: ] == '.xlsx':  # aqui esta criando o arquivo excel
                     df.to_excel(f'./SRC/Tarefas/{self.nome_arquivo}', index = False)

                else:
                     df.to_excel(f'./SRC/Tarefas/{self.nome_arquivo}.xlsx', index = False)

                print('\nPlanilha Criada com Sucesso! Pequeno Gafanhoto kkkk....')

            except Exception as e:
                 print(f'Erro: {e}')
#Enviar Email:                    
    def Enviar_Email(self):
            endereco = 'xxxxx@gmail.com' #colocar email
            
            with open( './SRC/Senha/Senha.txt', 'r', encoding='utf8') as arquivo:
                 s = arquivo.readlines()
            senha = s[0]

            msg = EmailMessage()
            msg['From'] = endereco
            msg[ 'To']  = self.email
            msg[ 'Subject'] = ' Ooooo Zé, chegou a planilha'
            msg.set_content(' Planilha em anexo.'
            )
            arquivos = [f'./SRC/Tarefas/{self.nome_arquivo}.xlsx']

            for arquivo in arquivos:
                with open(arquivo, 'rb') as arq: # rb para linguagem de maquina leitura binaria
                    dados = arq.read()
                    nome_arquivo = arq.name
                msg.add_attachment(
                     dados,
                     maintype = 'application',
                     subtype = 'octet-stream',
                     filename = nome_arquivo
                )
            server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            server.login(endereco, senha, initial_response_ok= True)
            server.send_message(msg)
            print('Email enviado com sucesso!!!')
 #Enviar pra whats up   
    def Enviar_whats(self):
        try:
            numero_destino = '+55119.......'# colocar celular
            mensagem =  'Oooou mandei a planilha lá!!!'

            kit.sendwhatmsg_instantly(numero_destino, mensagem,wait_time= 60)
            print('\nwhats Enviado')

        except Exception as e:
             print(f'Erro: {e}')
    



start = toDO()
start. iniciar()
