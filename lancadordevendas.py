import smtplib
import email.message
import pandas as pd
import win32com.client as win32
tabela_vendas = pd.read_excel('Vendas.xlsx')

#mostrar maximo de colunas
pd.set_option('display.max_columns', None)
print(tabela_vendas)

print('-'*50)
#informar faturamento: agrupando
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)

print('-'*50)
#quantidade de produtos vendidos
quantidade = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-'*50)
#ticket medio por produto em cada loja
ticket_medio = (faturamento['Valor Final']/quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

#enviar por email

def enviar_email():
    corpo_email = f'''

    <p>segue o relatório de vendas por cada loja.</p>
    
    <p>Faturamento:</p>
    {faturamento.to_html()}
    
    <p>Quantidade vendida:</p>
    {quantidade.to_html()}
    
    <p>Ticket médio dos produtos em cada loja:</p>
    {ticket_medio.to_html()}
    
    <p>Qualquer dúvida  estou à disposição.</p>
    
    <p>att.,</p>
    <p>Sidney</p>
    '''

    msg = email.message.Message()
    msg['Subject'] = "Assunto"
    msg['From'] = 'seuemail@gmail.com'
    msg['To'] = 'destinatario@gmail.com'
    password = 'suasenha'
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email )

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
print('Email enviado!')
enviar_email()