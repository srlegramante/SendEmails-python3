import win32com.client as win32

faturamento_exe = 300.000
quantidade_pcs = 1324

#criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')

#criar um email
email = outlook.CreateItem(0)

#configurar as informações do seu e-mail
email.To = "swithvegas200@gmail.com"
email.Subject = "Faturamento do setor de roupas!"
email.HTMLBody = f"""
<p>Melhor taxa de faturamento desse ano</p>

<p>O faturamento desse mês foi de {faturamento_exe}</p>
<p>A quantidade de peças doi o total de {quantidade_pcs}</p>
<p>Tivemos o rendimento de 30% a mais do mês passado,é um aumento consideravél.</p>

<p>Abs,</p>
<p>Loja americana</p>
"""

"""
anexo = "C://local do arquivo"
email.Attachments.Add(anexo)
"""

email.Send()
print("Email enviado")