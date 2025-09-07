' Criado por: André Athaydes  Data:01/04/25
' Atualizado por:
'Função: Gerar e copiar a mensagem para Área de Transferência

Sub mensagem_Cliente() 'Primeira Versão do Script - Não Utilizado Atualmente

Dim ovNumero As String
Dim horaAtual As Integer
Dim saudacao As String
Dim mensagem As String
Dim DataObj As New MSForms.DataObject

horaAtual = Hour(Now)

ovNumero = Sheets("RUA CADASTRADA").Range("A2").Value

If horaAtual < 12 Then
    saudacao = "Bom dia!"
    
Else
    saudacao = "Boa tarde!"
End If

mensagem = saudacao & vbCrLf & vbCrLf & "Somos da EMPRESA X e estamos entrando em contato referente à sua solicitação de ligação nova." & _
           vbCrLf & vbCrLf & "Para seguirmos com sua solicitação, precisamos analisar a rede que atende o local. Para agilizar o atendimento, pedimos que nos envie as coordenadas do local." & _
           vbCrLf & vbCrLf & "Para enviá-las, basta estar no local da ligação e, via WhatsApp, clicar em Anexos -> Localização -> Enviar localização fixa." & _
           vbCrLf & vbCrLf & "Não é necessário informar o endereço novamente, pois já o possuímos. Precisamos apenas das coordenadas para identificação no sistema." & _
           vbCrLf & vbCrLf & "Número da sua solicitação: " & ovNumero & _
           vbCrLf & vbCrLf & "Atenciosamente," & _
           vbCrLf & "*EMPRESA X*"
           

DataObj.SetText mensagem
DataObj.PutInClipboard


End Sub

Sub mensagem_Cliente_v2() ' Segunda Versão do Script - Utilizado Atualmente

Dim ovNumero As String
Dim horaAtual As Integer
Dim saudacao As String
Dim mensagem As String
Dim DataObj As New MSForms.DataObject

horaAtual = Hour(Now)


ovNumero = ActiveCell.EntireRow.Cells(1, 1).Value

If horaAtual < 12 Then
    saudacao = "Bom dia!"
    
Else
    saudacao = "Boa tarde!"
End If

mensagem = saudacao & vbCrLf & vbCrLf & "Somos da EMPRESA X e estamos entrando em contato referente à sua solicitação de ligação nova." & _
           vbCrLf & vbCrLf & "Para seguirmos com sua solicitação, precisamos analisar a rede que atende o local. Para agilizar o atendimento, pedimos que nos envie as coordenadas do local." & _
           vbCrLf & vbCrLf & "Para enviá-las, basta estar no local da ligação e, via WhatsApp, clicar em Anexos -> Localização -> Enviar localização fixa." & _
           vbCrLf & vbCrLf & "Não é necessário informar o endereço novamente, pois já o possuímos. Precisamos apenas das coordenadas para identificação no sistema." & _
           vbCrLf & vbCrLf & "Número da sua solicitação: " & ovNumero & _
           vbCrLf & vbCrLf & "Atenciosamente," & _
           vbCrLf & "*EMPRESA X*"
           

DataObj.SetText mensagem
DataObj.PutInClipboard

'MsgBox "Mensagem com OV " & ovNumero & " copiada com sucesso!", vbInformation
Application.StatusBar = "Mensagem com OV " & ovNumero & " copiada com sucesso!" 'Aviso sutil na barra de status
Application.Wait Now + TimeValue("0:00:03") ' Espera 3 segundos
Application.StatusBar = False               ' Restaura o padrão

End Sub
