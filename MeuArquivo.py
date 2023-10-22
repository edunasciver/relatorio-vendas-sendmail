import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# importar a base de dados da tabela de vendas.
tabelas_vendas = pd.read_excel('Vendas.xlsx')

# faturamento por loja
faturamento = tabelas_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# quantidade por loja
quantidade = tabelas_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

# ticket medio por loja
print('_' * 50)
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

username = "edu.nasciver@gmail.com"
password = "&Du180997"
mail_from = "edu.nasciver@gmail.com"
mail_to = "edu.nasciver@gmail.com"
mail_subject = "Relatorio de Vendas por Loja"
mail_body = (f'''
<h1>Prezados,</h1>
                        
<p>Segue o Relatório de Vendas por cada Loja.</p>

Faturamento:
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

Quantidade Vendida:
{quantidade.to_html(formatters={'Quantidade Vendida': 'R${:,.2f}'.format})}

Tikect Médio dos Produtos em cada Loja:
{ticket_medio.to_html(formatters={'Tikect Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.<p/>
<p>Att.,<p/>
<p>Eduardo Nascimento<p/>                  
''')

mimemsg = MIMEMultipart()
mimemsg['From']=mail_from
mimemsg['To']=mail_to
mimemsg['Subject']=mail_subject
mimemsg.attach(MIMEText(mail_body, 'html'))

connection = smtplib.SMTP(host='smtp.office365.com', port=587)
connection.starttls()
connection.login(username,password)
connection.send_message(mimemsg)
connection.quit()

print('Email Enviado')


# visualizar a base de dados
# pd.set_option('display.max_columns', None)
# print(tabelas_vendas)
# quantidade de produtos vendidos por loja
# ticket médio por produto em cada loja
