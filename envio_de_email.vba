Sub Enviando_email()
   
    Sheets("Restrições de Entrega").Activate 'ativando a aba para o robô não se perder.
   linha = 2 'definindo a Variavel como 2 = Linha inicial da aba Restrições de entrega
   Linha_Fim = Range("J1").End(xlDown).Row 'definindo a linha final
   


While linha <= Linha_Fim 'Loop para fazer enquanto a linha inicial for menor do que a inicial
Set Objeto_Outlook = CreateObject("outlook.Application") 'declarando que o nome(variavel) Outlook = A Aplicação Outlook
Set Email = Objeto_Outlook.CreateItem(0) 'Declarando que um novo item se chama Email
    Email.Display 'Mostrar Esse novo Email
    
    Email.To = Range("J" & linha).Value 'Escrever o valor da Coluna J Linha atual
    Email.cc = "torredecontrole_b2b@wine.com.br"
    Email.Subject = "B2Bl Restrição de entrega" & " " & "Pedido: " & Range("F" & linha).Value & " " & Range("L" & linha).Value & ", " & "NF " & Range("G" & linha).Value 'assunto do email
    Email.HTMLBody = Range("Texto!A1").Value & Worksheets("Restrições de entrega").Range("K" & linha).Value & "<br><br><br><img src=C:\Temp\image.png>" 'Escrever o valor Da Coluna A Linha 1 da aba TEXTO
    'para adicionar imagem lembre que precisa colocar a imagem na pasta C:\Temp com o nome de imagem.png
    'para trocar a imagem é só substituir esse arquivo mas o nome tem que ser o mesmo.
    
    Email.send ' Envia o email


linha = linha + 1 'A variavel Linha soma mais 1
Wend

End Sub
