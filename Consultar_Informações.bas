' Criado por: André Athaydes  Data:01/04/25
' Atualizado por:
' Função: Extrair os dados do logradouro da Transação ZDSS02

Sub Consultar_informacoes()

Application.ScreenUpdating = False

usu = Sheets("Login").Range("B2").Value
Senha = Sheets("Login").Range("C2").Value

If Not IsObject(Application1) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set Application1 = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection1) Then
   Set Connection1 = Application1.OpenConnection("PRODUÇÃO CCS ( EP2 ) - EDP ES")
End If
If Not IsObject(session1) Then
   Set session1 = Connection1.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session1, "on"
   WScript.ConnectObject Application, "on"
End If

'script de logon e acesso à transação no SAP
session1.findById("wnd[0]").resizeWorkingPane 160, 32, False
session1.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usu
session1.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Senha
session1.findById("wnd[0]").sendVKey 0
session1.findById("wnd[0]").maximize
session1.findById("wnd[0]").sendVKey 0


Sheets("INFORMAÇÕES").Select
Range("A2").Select

Do While ActiveCell <> ""

OV = ActiveCell

session1.findById("wnd[0]/tbar[0]/okcd").Text = "/nzsd02"
session1.findById("wnd[0]").sendVKey 0
session1.findById("wnd[0]/usr/ctxtS_VBELN-LOW").Text = OV
session1.findById("wnd[0]/usr/ctxtS_VBELN-LOW").caretPosition = 10
session1.findById("wnd[0]/tbar[1]/btn[8]").press
session1.findById("wnd[0]/usr/lbl[66,8]").SetFocus

rua = session1.findById("wnd[0]/usr/lbl[66,8]").Text
session1.findById("wnd[0]").sendVKey 0
session1.findById("wnd[0]/usr/lbl[9,9]").SetFocus

bairro = session1.findById("wnd[0]/usr/lbl[9,9]").Text
session1.findById("wnd[0]").sendVKey 0
session1.findById("wnd[0]/usr/lbl[83,9]").SetFocus

municipio = session1.findById("wnd[0]/usr/lbl[83,9]").Text
session1.findById("wnd[0]").sendVKey 0


ActiveCell.Offset(0, 1).Value = rua
ActiveCell.Offset(0, 2).Value = bairro
ActiveCell.Offset(0, 3).Value = municipio

ActiveCell.Offset(1, 0).Select

Loop
session1.findById("wnd[0]").sendVKey 3 'F3 - Retornar página anterior
session1.findById("wnd[0]").sendVKey 3 'F3 - Retornar página anterior
Application.ScreenUpdating = True

End Sub

