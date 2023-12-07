import pandas as pd

tabela_vendas = pd.read_excel('Vendas.xlsx')
#Faturamento por Loja
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
#Quantidade de produtos vendidos por loja
produtos_lojas = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
ticket_medio = (faturamento['Valor Final'] / produtos_lojas['Quantidade']).to_frame()
ticket_medio=ticket_medio.rename(columns={0:'Ticket_Médio'})
#Relatorio por Email
import smtplib
import email.message

def enviar_email():
    corpo_email = f'''
    
    <p>Faturamento</p>
    {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}
    <p>Quantidade Vendida:</p>
{produtos_lojas.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}
    '''

    msg = email.message.Message()
    msg['Subject'] = "Assunto"
    msg['From'] = 'Remetente'
    msg['To'] = 'Destinatario'
    password = 'sua senha' 
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email )

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')


# In[ ]:


enviar_email()

