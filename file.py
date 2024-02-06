import pandas as pd
import smtplib
import email.message

# Importar a base de dados 
tabela_vendas = pd.read_excel("Vendas.xlsx")

# Faturamento por loja 
faturamento = tabela_vendas.groupby("ID Loja")["Valor Final"].sum()

# Quantidade de produtos vendidos por loja 
quantidade = tabela_vendas.groupby("ID Loja")["Quantidade"].sum()

# Tiket médio por produto em cada loja 
ticket_medio = faturamento / quantidade

# Enviar e-mail com o relatório 
def enviar_email():  
    corpo_email = f"""
    <p>Prezados,</p>
    <p>Segue o relatório de vendas por cada loja</p>
    
    <p>Faturamento:</p>
    {faturamento.to_frame().to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

    <p>Quantidade:</p>
    {quantidade.to_frame().to_html()}
    
    <p>Ticket médio por loja:</p>
    {ticket_medio.to_frame().to_html(formatters={0: 'R${:,.2f}'.format})}
    
    <p>Qualquer dúvida estou à disposição!</p>
    <p>Att,</p>
    <p>Pedro Queiroz</p>
    """
    
    msg = email.message.Message()
    msg['Subject'] = "Relatório de vendas por loja"
    msg['From'] = 'douglas.333.goncalves@gmail.com'
    msg['To'] = 'douglas.333.goncalves@gmail.com'
    password = 'srik ijnh ctoo zaxb' 
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')

# Enviar e-mail
enviar_email()
