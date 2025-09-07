' Criado por: André Athaydes  Data:01/04/25
' Atualizado por: André Athaydes  Data:05/06/25
' Função: Verifica se os logradouros foram processados no SAP através da transação SR22

Sub Logradouro_SAP_SR22() ' Verifica se o logradouro subiu no SAP e preenche o logradouro completo na planilha

Application.ScreenUpdating = False

usu = Sheets("Login").Range("B2").Value
Senha = Sheets("Login").Range("C2").Value

'Verifica se já existe conexão com SAP, se não cria uma nova.
If Not IsObject(Application1) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set Application1 = SapGuiAuto.GetScriptingEngine
End If

If Not IsObject(Connection1) Then
   Set Connection1 = Application1.OpenConnection("PRODUÇÃO CCS ( EP2 ) - EDP ES")
End If
If Not IsObject(session1) Then
   Set session1 = Connection1.Children(0) 'seleciona a primeira sessão aberta (janela) do sap
End If
If IsObject(WScript) Then
   WScript.ConnectObject session1, "on"
   WScript.ConnectObject Application, "on"
End If

'script de login SAP
session1.findById("wnd[0]").resizeWorkingPane 160, 32, False 'Redimensiona a janela principal do SAP
session1.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usu 'LOGIN
session1.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Senha 'SENHA
session1.findById("wnd[0]").sendVKey 0 'ENTER NO LOGIN
session1.findById("wnd[0]").maximize ' MAXIMIZA A JANELA
session1.findById("wnd[0]").sendVKey 0 'ENTER PARA CASO DE MENSAGENS OU AVISO QUE POSSA APARECER APÓS O LOGIN

'Seleciona uma célula da planilha
Sheets("SR22").Select
Range("A2").Select

'Iniciará um loop enquanto houver dados a partir da célula selecionada anteriormente
Do While ActiveCell <> "" 'Faça enquanto a célula não conter vazio

codigo = ActiveCell

'Abre a transação sr22
session1.findById("wnd[0]").maximize
session1.findById("wnd[0]/tbar[0]/okcd").Text = "sr22"
session1.findById("wnd[0]").sendVKey 0
session1.findById("wnd[0]").maximize
session1.findById("wnd[0]/usr/ctxtADRSTREETD-STRT_CODE").Text = codigo
session1.findById("wnd[0]/usr/ctxtADRSTREETD-COUNTRY").Text = "br"
session1.findById("wnd[0]").sendVKey 0

erroSAP = session1.findById("wnd[0]/sbar").Text
If erroSAP <> "" Then
    ActiveCell.Offset(0, 1).Value = "NÃO ENCONTRADO"
    ActiveCell.Offset(0, 2).Value = "-"
    ActiveCell.Offset(0, 3).Value = "-"
Else
    cidade = session1.findById("wnd[0]/usr/subCITY:SAPLSZRC:0220/ctxtADRCITYD-CITY_NAME").Text
    rua = session1.findById("wnd[0]/usr/ctxtADRSTREETD-STREET").Text
    bairro = session1.findById("wnd[0]/usr/txtADRSTREETD-CITY_PART").Text
    ActiveCell.Offset(0, 1).Value = rua
    ActiveCell.Offset(0, 2).Value = bairro
    ActiveCell.Offset(0, 3).Value = cidade
End If

ActiveCell.Offset(1, 0).Select
session1.findById("wnd[0]").sendVKey 3 'F3 - Retornar página anterior
Loop

Application.ScreenUpdating = True
End Sub

Sub SAP_SR22_versao2() 'Somente Verifica se o Logradouro subiu no SAP ou não

Application.ScreenUpdating = False

usu = Sheets("Login").Range("B2").Value
Senha = Sheets("Login").Range("C2").Value

'Verifica se já existe conexão com SAP, se não cria uma nova.
If Not IsObject(Application1) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set Application1 = SapGuiAuto.GetScriptingEngine
End If

If Not IsObject(Connection1) Then
   Set Connection1 = Application1.OpenConnection("PRODUÇÃO CCS ( EP2 ) - EDP ES")
End If
If Not IsObject(session1) Then
   Set session1 = Connection1.Children(0) 'seleciona a primeira sessão aberta (janela) do sap
End If
If IsObject(WScript) Then
   WScript.ConnectObject session1, "on"
   WScript.ConnectObject Application, "on"
End If

'script de login SAP
session1.findById("wnd[0]").resizeWorkingPane 160, 32, False 'Redimensiona a janela principal do SAP
session1.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usu 'LOGIN
session1.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Senha 'SENHA
session1.findById("wnd[0]").sendVKey 0 'ENTER NO LOGIN
session1.findById("wnd[0]").maximize ' MAXIMIZA A JANELA
session1.findById("wnd[0]").sendVKey 0 'ENTER PARA CASO DE MENSAGENS OU AVISO QUE POSSA APARECER APÓS O LOGIN

'Seleciona a célula E2 da planilha RUA CADASTRADA
Sheets("RUA CADASTRADA").Select
Range("E2").Select

'Iniciará um loop enquanto houver dados a partir da célula selecionada anteriormente
Do While ActiveCell <> "" 'Faça enquanto a célula não conter vazio

codigo = ActiveCell

'Abre a transação sr22
session1.findById("wnd[0]").maximize
session1.findById("wnd[0]/tbar[0]/okcd").Text = "sr22"
session1.findById("wnd[0]").sendVKey 0
session1.findById("wnd[0]").maximize
session1.findById("wnd[0]/usr/ctxtADRSTREETD-STRT_CODE").Text = codigo
session1.findById("wnd[0]/usr/ctxtADRSTREETD-COUNTRY").Text = "br"
session1.findById("wnd[0]").sendVKey 0

erroSAP = session1.findById("wnd[0]/sbar").Text
If erroSAP <> "" Then
    ActiveCell.Offset(0, 1).Value = "NÃO ENCONTRADO"
Else
    codigo_validacao = session1.findById("wnd[0]/usr/ctxtADRSTREETD-STRT_CODE").Text
    ActiveCell.Offset(0, 1).Value = codigo_validacao
End If

session1.findById("wnd[0]").sendVKey 3 'F3 - Retornar página anterior

ActiveCell.Offset(1, 0).Select

Loop

session1.findById("wnd[0]").sendVKey 3 'F3 - Retornar página anterior

Application.ScreenUpdating = True
End Sub
