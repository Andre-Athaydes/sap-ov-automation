'Criado por André Athaydes Martins  Data:05/06/25
'Funcionamento: O Script verifica, primeiramente, se os logradouros de cada OV foram
'processados pelo SAP. Em seguida, o scritp altera o status da OV para 21 somente aquelas que cujo logradouros subiram.


Sub Validaçao_e_Alteracao_21()

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

'O código abaixo executa a alteração do Status da OV para 21 caso o Campo Validação seja igual ao Código do Logradouro
Sheets("RUA CADASTRADA").Select
Range("A9").Select

Do While ActiveCell <> ""

OV = ActiveCell
rua = ActiveCell.Offset(0, 1).Value
bairro = ActiveCell.Offset(0, 2).Value
municipio = ActiveCell.Offset(0, 3).Value
codigo = ActiveCell.Offset(0, 4).Value
codigo_validacao = ActiveCell.Offset(0, 5).Value
DataAtual = Format(Date, "dd/MM/yyyy")

'------------- Validação -------------------
If codigo_validacao <> codigo Then
    GoTo Proximo
End If

Dim textoNovo As String
Dim textoAntigo As String

textoNovo = "Rua cadastrada com sucesso!" & vbCrLf & _
        codigo & vbCrLf & _
        rua & vbCrLf & _
        bairro & vbCrLf & _
        municipio & vbCrLf & _
        Nome & vbCrLf & _
        DataAtual & vbCrLf & _
        "___________________________" & vbCrLf

'------------------Execução SAP -------------------------------
session1.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
session1.findById("wnd[0]").sendVKey 0

session1.findById("wnd[0]").maximize
session1.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = OV
session1.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 10
session1.findById("wnd[0]").sendVKey 0

session1.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").Text = "RUA CADASTRADA"

'session1.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").Text = "RUA CADASTRADA"session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBKD-BSTKD").Text = "RUA CADASTRADA"
session1.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press

session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09").Select
session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBKD-IHREZ").Text = codigo
session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBKD-IHREZ_E").Text = codigo
session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBKD-IHREZ_E").SetFocus
session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBKD-IHREZ_E").caretPosition = 12

session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select
session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "Z013", "Column1"
session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "Z013", "Column1"
session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").topNode = "Z002"
session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "Z013", "Column1"
textoAntigo = session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text
session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = textoNovo & vbCrLf & textoAntigo
session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes 42, 42
session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10").Select
session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4305/btnBT_KSTC").press
session1.findById("wnd[1]").Close
session1.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/sub:SAPLBSVA:0302[1]/txtJEST_BUF_E-ETX30[1,11]").SetFocus
session1.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/sub:SAPLBSVA:0302[1]/txtJEST_BUF_E-ETX30[1,11]").caretPosition = 21
session1.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/btnADOWN").press
session1.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/btnADOWN").press
session1.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/sub:SAPLBSVA:0302[1]/radJ_STMAINT-ANWS[2,0]").Select
session1.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/sub:SAPLBSVA:0302[1]/radJ_STMAINT-ANWS[2,0]").SetFocus
session1.findById("wnd[0]/tbar[0]/btn[3]").press
session1.findById("wnd[0]/tbar[0]/btn[11]").press
session1.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
GoTo Proximo

Proximo:
    ActiveCell.Offset(1, 0).Select

Loop
' Término do Código Alterar status para 21

session1.findById("wnd[0]").sendVKey 3 'F3 - Retornar página anterior

Application.ScreenUpdating = True

MsgBox "Script Finalizado!", vbInformation, "Concluído!"

End Sub



