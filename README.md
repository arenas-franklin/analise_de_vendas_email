# Analise de vendas 



## Objetivos

Realizar uma análise de vedas de uma tabela do excel e enviar um relatório por email

## Etapas 


- importar a base de dados
- visualizar a base de dados
- faturamento por loja

- quantidade de produtos vendidos por loja

- ticket médio por produto por cada loja

- enviar e-mail com relatório 



 ## Bibliotecas utilizadas 
 - pandas 
 - pywin32
 
 para instalar a biblioteca pode digitar no terminal 
~~~
    pip install nome_da_bibioteca 
~~~

ou 

~~~
pip install -r requirements.txt
~~~
 
 
## Exemplo do codigo envio de email:
 
~~~ python


import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outllok.CreateItem(0)
mail.To = 'To addres'
mail.Subject = 'Message Subject'
mail.HTMLBody = '<h2>HTML Message body</h2>'

mail.Send()

~~~


