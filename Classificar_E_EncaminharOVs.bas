'Criado por: André Athaydes Martins Data:
'Atualizado por:
'Função: Altera o status da OV p/ 21 em casos de logradouros Existente;
'       Transfere as OVs de novos logradouros para a tabela "RUA CADASTRADA";
'       Envia carta para os clientes para os pedidos não elaborados

Sub SepararOVs_CadastrosExistentes()

Application.ScreenUpdating = False

Dim textoNovo As String
Dim textoAntigo As String

Dim linhaAtual As Long
Dim linhaDestino As Long
Dim wsOrigem As Worksheet
Dim wsDestino As Worksheet
Dim tblOrigem As ListObject
Dim tblDestino As ListObject
Dim i As Long
Dim linhaVazia As Boolean
Dim LinhaNaTabela As Long

DataAtual = Format(Date, "dd/MM/yyyy")

usu = Sheets("Login").Range("B2").Value
Senha = Sheets("Login").Range("C2").Value
Nome = Sheets("Login").Range("A2").Value

If Not IsObject(Application1) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set Application1 = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection1) Then
   Set Connection1 = Application1.OpenConnection("PRODUÇÃO CCS ( EP2 )")
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

DataAtual = Format(Date, "dd/MM/yyyy")

Do While ActiveCell <> ""

    OV = ActiveCell
    rua = ActiveCell.Offset(0, 1).Value
    bairro = ActiveCell.Offset(0, 2).Value
    municipio = ActiveCell.Offset(0, 3).Value
    codigo = ActiveCell.Offset(0, 4).Value
    codigoLogradouro = ActiveCell.Offset(0, 5).Value
    cadastroNovo = ActiveCell.Offset(0, 6).Value
    
    
    If cadastroNovo = "SIM" Then
    
    'Define as abas
        Set wsOrigem = ThisWorkbook.Sheets("INFORMAÇÕES")
        Set wsDestino = ThisWorkbook.Sheets("RUA CADASTRADA")
        
        Set tblOrigem = wsOrigem.ListObjects("TabelaInformacoes")
        Set tblDestino = wsDestino.ListObjects("TabelaRuaCadastrada")
        
        
        'Copia a linha da célula ativa
        linhaAtual = ActiveCell.Row
        linhaVazia = False
        
        For i = 1 To tblDestino.ListRows.Count
            If Application.WorksheetFunction.CountA(tblDestino.ListRows(i).Range) = 0 Then
                'Linha completamente vazia encontrada
                tblDestino.ListRows(i).Range.Resize(1, 5).Value = wsOrigem.Range("A" & linhaAtual & ":E" & linhaAtual).Value
                linhaVazia = True
                Exit For
            End If
        Next i
        
        'Se não achou linha vazia, adiciona nova linha ao final Observar Orgem
        If Not linhaVazia Then
            tblDestino.ListRows.Add.Range.Resize(1, 5).Value = wsOrigem.Range("A" & linhaAtual & ":E" & linhaAtual).Value
        End If
        
        
        For i = 1 To tblOrigem.ListRows.Count
            If tblOrigem.ListRows(i).Range.Row = ActiveCell.Row Then
            LinhaNaTabela = i
            Exit For
            End If
        Next i
        tblOrigem.ListRows(LinhaNaTabela).Delete
    
    ElseIf cadastroNovo = "NE" Then
        'Call Enviar_Carta
                
        Dim Carta_Cliente As String
        
        Carta_Cliente = "Em atenção à solicitação de Vossa Senhoria, informamos que não foi possível dar sequência na análise necessária, pelo seguinte motivo:" & vbCrLf & vbCrLf & _
                        "Não foi apresentado documento, com data, que comprove a propriedade ou posse do imóvel, tais como: escritura, documento formal de partilha homologado, contrato de compra e venda, todos devidamente registrado no cartório de imóveis ou CCIR (Certificado de Cadastro de Imóvel Rural), acompanhado de contrato de compra e venda, para os casos de parcelamento do solo." & vbCrLf & vbCrLf & _
                        "De posse do referido documento, V. Sa deverá comparecer em uma de nossas agências de atendimento, para abertura de nova solicitação." & vbCrLf & vbCrLf & _
                        "Caso Vsª esteja localizado em zona urbana, pedimos que apresente o IPTU junto a Prefeitura Municipal."
        
        textoNovo = "PEDIDO NÃO ELABORADO" & vbCrLf & _
                "ENVIADA A CARTA PARA O CLIENTE" & vbCrLf & vbCrLf & _
                "O cliente não apresentou documento oficial que comprove a titularidade ou a existência do logradouro informado, conforme exigido para prosseguir com o cadastro. Exemplos de documentos aceitos incluem:" & vbCrLf & vbCrLf & _
                "- Escritura do imóvel;" & vbCrLf & _
                "- Contrato de compra e venda registrado em cartório de IMÓVEIS;" & vbCrLf & _
                "-IPTU emitido pela Prefeitura;" & vbCrLf & _
                "- CCIR (Certificado de Cadastro de Imóvel Rural), acompanhado de contrato de compra e venda (em caso de parcelamento do solo)." & vbCrLf & vbCrLf & _
                Nome & vbCrLf & _
                DataAtual & vbCrLf & _
                "___________________________" & vbCrLf
        
        'Acessar a Transação VA02
        session1.findById("wnd[0]").maximize
        session1.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
        session1.findById("wnd[0]").sendVKey 0                     'Imita a tecla Enter
        
        'Acessa a OV na Transação VA02
        session1.findById("wnd[0]").maximize                        'Maximiza a Janela
        session1.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = OV
        session1.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 10
        session1.findById("wnd[0]").sendVKey 0
        
        'Escreve NE CLIENTE no campo texto
        session1.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").Text = "NE CLIENTE"
        session1.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
        
        'Altera a equipe de vendas para CEE
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-VKGRP").Text = "CCE"
        
        'Acessa ao menu Dados do Pedido
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09").Select
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBKD-BSTKD").Text = "NE CLIENTE"  'Escreve NE CIENTE no campo
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBKD-IHREZ").Text = codigo        'Escreve o código no campo Referências
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBKD-IHREZ_E").Text = codigo      'Escreve o código no campo Referências
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBKD-IHREZ_E").SetFocus
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBKD-IHREZ_E").caretPosition = 12
        
        'Acessa ao menu Textos
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select
        
        'Acessa o campo do Observ.Pedido Não Elaborado do menu Textos
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "Z002", "Column1"
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "Z002", "Column1"
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "Z002", "Column1"
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = Carta_Cliente
        
        
        'Acessa o campo do Histórico de Solicitação do menu Textos
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "Z013", "Column1"
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "Z013", "Column1"
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").topNode = "Z002"
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "Z013", "Column1"
        textoAntigo = session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = textoNovo & vbCrLf & textoAntigo
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes 42, 42
        
        'Acessa ao Menu Status e Altera para status 1
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10").Select
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4305/btnBT_KSTC").press
        session1.findById("wnd[1]").Close
        session1.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/sub:SAPLBSVA:0302[1]/radJ_STMAINT-ANWS[0,0]").Select
        session1.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/sub:SAPLBSVA:0302[1]/radJ_STMAINT-ANWS[0,0]").SetFocus
        session1.findById("wnd[0]").sendVKey 0
        session1.findById("wnd[0]/tbar[0]/btn[3]").press
        
        'Acessa aos Grupos e seleciona "NE Resposta Cliente" no campo Grupo 4 e Grupo 5
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11").Select
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR4").Key = "011"
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR5").Key = "054"
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR5").SetFocus
        
        'Salva as alterações feitas na OV e volta para a página inicial da transação VA02
        session1.findById("wnd[0]/tbar[0]/btn[3]").press
        session1.findById("wnd[0]/tbar[0]/btn[11]").press
        session1.findById("wnd[0]").sendVKey 0
        
        
        ActiveCell.Offset(1, 0).Select
          
    ElseIf cadastroNovo = "NÃO" And codigoLogradouro <> "" Then
        'Call Rua_Existente
        textoNovo = "Rua Existente" & vbCrLf & _
        codigo & vbCrLf & _
        rua & vbCrLf & _
        bairro & vbCrLf & _
        municipio & vbCrLf & _
        Nome & vbCrLf & _
        DataAtual & vbCrLf & _
        "___________________________" & vbCrLf


        session1.findById("wnd[0]/tbar[0]/okcd").Text = "/nva02"
        session1.findById("wnd[0]").sendVKey 0
        
        session1.findById("wnd[0]").maximize
        session1.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = OV
        session1.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 10
        session1.findById("wnd[0]").sendVKey 0
        
        session1.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").Text = "RUA EXISTENTE"
        session1.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
        
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09").Select
        session1.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBKD-BSTKD").Text = "RUA EXISTENTE"
        
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
        
        ActiveCell.Offset(1, 0).Select
        
    End If
    
Loop

Application.ScreenUpdating = True

MsgBox "Script Finalizado!", vbInformation, "Concluído!"

End Sub
