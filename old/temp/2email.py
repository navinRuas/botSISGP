import win32com.client as win32 
import pandas as pd 

data = {'name': ['John', 'Smith', 'Paul'], 'age': [25, 30, 50]} 
df = pd.DataFrame(data) 

outlook = win32.Dispatch('outlook.application') 
email = outlook.CreateItem(0)
email.To = 'jamil.monteiro@inep.gov.br' 
email.Subject = 'Dataframe' 

html = df.to_html() 

email.HTMLBody = html
email.HTMLBody = f""" <p>Caros Chefe e Claudio</p>
<p>Cordialmente,</p>
<p>Email automático</p>
"""

anexo = f'C://Users\jamil.monteiro\Documents\PGD.xlsx'
email.Attachments.Add(anexo)

email.Send()

'''import win32com.client as win32

# criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')

# criar um email
email = outlook.CreateItem(0)

# configurar as informações do seu e-mail
email.To = "jamil.monteiro@inep.gov.br;luiz.senna@inep.gov.br"
email.Subject = "Primeiro e-mail automático do Python"
email.HTMLBody = f"""
<p>Olá Jamil, agora é a hora da verdadae</p>
<p>Abs,</p>
<p>Código Python</p>
"""

email.Send()
print("Email Enviado")

email = outlook.CreateItem(0)

# configurar as informações do seu e-mail
email.To = "jamil.monteiro@inep.gov.br"
email.Subject = "Segundo e-mail automático do Python"
email.HTMLBody = f"""
<p>Olá Jamil, agora é a hora da verdadae</p>
<p>Abs,</p>
<p>Código Python</p>
"""

email.Send()
print("Email Enviado")'''