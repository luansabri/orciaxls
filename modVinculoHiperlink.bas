Attribute VB_Name = "modVinculoHiperlink"
Option Explicit

' ==============================================================================
' 1. GATILHO DA RIBBON
' ==============================================================================
Public Sub RibbonVincularTudo(control As IRibbonControl)
    ExecutarVinculoUnificado
End Sub

' ==============================================================================
' 2. CHAMADA DIRETA / TESTE
' ==============================================================================
Public Sub ChamarVincularTudo()
    ExecutarVinculoUnificado
End Sub

' ==============================================================================
' 3. LÓGICA REAL
' ==============================================================================
Public Sub ExecutarVinculoUnificado()
    Dim wbDestino As Workbook
    Dim wsOrigem As Worksheet
    Dim wsItem As Worksheet
    Dim wsE As Worksheet
    Dim celulaItem As Range
    Dim celulaQuant As Range
    Dim codigoItem As String
    Dim nomeAbaE As String
    Dim ultimaLinha As Long
    Dim colQuant As Long
    Dim celUltimaE As Range
    Dim enderecoTotalE As String
    Dim subAddressDesejado As String
    Dim r As Long, rCab As Long
    Dim celCab As Range
    Dim countItens As Long, countQuant As Long
    
    Set wbDestino = ActiveWorkbook
    If wbDestino Is Nothing Then Exit Sub
    
    If wbDestino.Name Like "*.xlam" Then
        MsgBox "Abra o arquivo de orçamento antes de executar.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    On Error Resume Next
    Set wsOrigem = wbDestino.Worksheets("Memorial de Cálculo")
    If wsOrigem Is Nothing Then Set wsOrigem = wbDestino.Worksheets("Orçamento")
    If wsOrigem Is Nothing Then Set wsOrigem = wbDestino.Worksheets("MC")
    If wsOrigem Is Nothing Then Set wsOrigem = wbDestino.Worksheets("Orcamento")
    If wsOrigem Is Nothing Then Set wsOrigem = wbDestino.Worksheets("Memorial de Calculo")
    On Error GoTo 0
    
    If wsOrigem Is Nothing Then
        MsgBox "Nenhuma planilha base foi encontrada." & vbCrLf & _
               "O arquivo deve conter uma aba chamada:" & vbCrLf & _
               "• Memorial de Cálculo" & vbCrLf & _
               "• Orçamento" & vbCrLf & _
               "• MC", vbCritical, "Erro de Localização"
        Exit Sub
    End If
    
    colQuant = 0
    For rCab = 1 To 15
        Set celCab = wsOrigem.Rows(rCab).Find("Quant.", LookIn:=xlValues, LookAt:=xlPart)
        If Not celCab Is Nothing Then
            colQuant = celCab.Column
            Exit For
        End If
    Next rCab
    If colQuant = 0 Then colQuant = 6
    
    ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.count, "A").End(xlUp).Row
    If ultimaLinha < 2 Then
        MsgBox "A planilha base não contém dados suficientes.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Criando vínculos de navegação..."
    
    ' SEM Hyperlinks.Delete global — protege a formatação!
    
    For r = 2 To ultimaLinha
        Set celulaItem = wsOrigem.Cells(r, "A")
        codigoItem = Trim(CStr(celulaItem.Value))
        
        If codigoItem <> "" Then
            
            ' A. VÍNCULO DO ITEM (Coluna A)
            Set wsItem = Nothing
            On Error Resume Next
            Set wsItem = wbDestino.Worksheets(codigoItem)
            On Error GoTo 0
            
            If Not wsItem Is Nothing Then
                subAddressDesejado = "'" & codigoItem & "'!A1"
                
                If celulaItem.Hyperlinks.count = 0 Then
                    ' Não tem link — cria
                    wsOrigem.Hyperlinks.Add Anchor:=celulaItem, Address:="", SubAddress:=subAddressDesejado
                    celulaItem.Font.Color = RGB(0, 102, 204)
                    celulaItem.Font.Underline = xlUnderlineStyleSingle
                    countItens = countItens + 1
                    
                ElseIf celulaItem.Hyperlinks(1).SubAddress <> subAddressDesejado Then
                    ' Link desatualizado — atualiza preservando formatação
                    AtualizarHiperlink wsOrigem, celulaItem, subAddressDesejado
                    celulaItem.Font.Color = RGB(0, 102, 204)
                    celulaItem.Font.Underline = xlUnderlineStyleSingle
                    countItens = countItens + 1
                End If
                
                Set wsItem = Nothing
            End If
            
            ' B. VÍNCULO DA QUANTIDADE (Coluna Quant.)
            nomeAbaE = codigoItem & "E"
            Set wsE = Nothing
            On Error Resume Next
            Set wsE = wbDestino.Worksheets(nomeAbaE)
            On Error GoTo 0
            
            If Not wsE Is Nothing Then
                Set celulaQuant = wsOrigem.Cells(r, colQuant)
                
                Set celUltimaE = wsE.Cells.Find("*", _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    LookIn:=xlValues)
                    
                enderecoTotalE = IIf(Not celUltimaE Is Nothing, _
                    celUltimaE.Address(False, False), "A1")
                
                subAddressDesejado = "'" & nomeAbaE & "'!" & enderecoTotalE
                
                If celulaQuant.Hyperlinks.count = 0 Then
                    ' Não tem link — cria
                    wsOrigem.Hyperlinks.Add Anchor:=celulaQuant, Address:="", SubAddress:=subAddressDesejado
                    celulaQuant.Font.Color = RGB(0, 102, 204)
                    celulaQuant.Font.Bold = True
                    celulaQuant.NumberFormat = "#,##0.00"
                    countQuant = countQuant + 1
                    
                ElseIf celulaQuant.Hyperlinks(1).SubAddress <> subAddressDesejado Then
                    ' Link desatualizado — atualiza preservando formatação
                    AtualizarHiperlink wsOrigem, celulaQuant, subAddressDesejado
                    celulaQuant.Font.Color = RGB(0, 102, 204)
                    celulaQuant.Font.Bold = True
                    celulaQuant.NumberFormat = "#,##0.00"
                    countQuant = countQuant + 1
                End If
                
                Set wsE = Nothing
            End If
            
        End If
        
        If r Mod 50 = 0 Then
            Application.StatusBar = "Processando linha " & r & " de " & ultimaLinha & "..."
        End If
        
    Next r
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    MsgBox "Vínculos de navegação criados com sucesso!" & vbCrLf & vbCrLf & _
           "• Itens vinculados: " & countItens & vbCrLf & _
           "• Quantidades (E) vinculadas: " & countQuant, vbInformation, "Sucesso"
End Sub

' ==============================================================================
' 4. ATUALIZA HIPERLINK PRESERVANDO FORMATAÇÃO DA CÉLULA
' ==============================================================================
Private Sub AtualizarHiperlink(ws As Worksheet, cel As Range, novoSubAddress As String)
    Dim corFundo As Long
    Dim temFundo As Boolean
    Dim b As Long
    Dim indicesBorda(1 To 4) As Long
    Dim estiloBorda(1 To 4) As Long
    Dim corBorda(1 To 4) As Long
    Dim pesoBorda(1 To 4) As Long
    Dim temBorda(1 To 4) As Boolean
    
    indicesBorda(1) = xlEdgeLeft
    indicesBorda(2) = xlEdgeRight
    indicesBorda(3) = xlEdgeTop
    indicesBorda(4) = xlEdgeBottom
    
    ' Salva fundo
    temFundo = (cel.Interior.Pattern <> xlNone)
    If temFundo Then corFundo = cel.Interior.Color
    
    ' Salva bordas
    For b = 1 To 4
        With cel.Borders(indicesBorda(b))
            temBorda(b) = (.LineStyle <> xlNone)
            If temBorda(b) Then
                estiloBorda(b) = .LineStyle
                corBorda(b) = .Color
                pesoBorda(b) = .Weight
            End If
        End With
    Next b
    
    ' Remove e recria o hiperlink
    cel.Hyperlinks.Delete
    ws.Hyperlinks.Add Anchor:=cel, Address:="", SubAddress:=novoSubAddress
    
    ' Restaura fundo
    If temFundo Then
        cel.Interior.Color = corFundo
    Else
        cel.Interior.Pattern = xlNone
    End If
    
    ' Restaura bordas
    For b = 1 To 4
        With cel.Borders(indicesBorda(b))
            If temBorda(b) Then
                .LineStyle = estiloBorda(b)
                .Color = corBorda(b)
                .Weight = pesoBorda(b)
            Else
                .LineStyle = xlNone
            End If
        End With
    Next b
End Sub

