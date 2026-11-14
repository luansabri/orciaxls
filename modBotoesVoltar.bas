Attribute VB_Name = "modBotoesVoltar"
Option Explicit

' ==============================================================================
' 1. GATILHO DA RIBBON
' ==============================================================================
Public Sub RibbonAdicionarBotoesVoltar(control As IRibbonControl)
    ExecutarAdicionarBotoesVoltar
End Sub

' ==============================================================================
' 2. CHAMADA DIRETA / TESTE VIA ALT+F8
' ==============================================================================
Public Sub ChamarAdicionarBotoesVoltar()
    ExecutarAdicionarBotoesVoltar
End Sub

' ==============================================================================
' 3. LÓGICA REAL ISOLADA
' ==============================================================================
Public Sub ExecutarAdicionarBotoesVoltar()
    Dim wbDestino As Workbook
    Dim ws As Worksheet
    Dim wsBase As Worksheet
    Dim nomeBase As String
    Dim countBotoes As Long
    
    Set wbDestino = ActiveWorkbook
    If wbDestino Is Nothing Then Exit Sub
    
    ' Proteção contra execução no próprio suplemento
    If wbDestino.Name Like "*.xlam" Then
        MsgBox "Abra o arquivo de orçamento antes de executar.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' 1. Seleciona a aba base automaticamente usando o seletor modular
    nomeBase = SelecionarAbaBaseVoltar(wbDestino)
    If nomeBase = "" Then Exit Sub
    
    Set wsBase = wbDestino.Worksheets(nomeBase)
    
    On Error GoTo TratarErro
    Application.ScreenUpdating = False
    
    ' 2. Processa cada aba do arquivo
    For Each ws In wbDestino.Worksheets
        ' Ignora a aba base mantida e abas ocultas
        If ws.Name <> wsBase.Name And ws.Visible = xlSheetVisible Then
            
            ' Remove botões "VOLTAR" antigos existentes na aba
            RemoverBotoesVoltarAntigos ws
            
            ' Cria e formata o novo botão "VOLTAR" na célula A3
            CriarBotaoVoltarNaAba ws, wsBase
            
            countBotoes = countBotoes + 1
        End If
    Next ws

    Application.ScreenUpdating = True
    MsgBox "Botões 'VOLTAR' criados com sucesso em " & countBotoes & " aba(s)!" & vbCrLf & _
           "Aba de destino: '" & wsBase.Name & "'", vbInformation, "Concluído"
    Exit Sub

TratarErro:
    Application.ScreenUpdating = True
    MsgBox "Erro na aba '" & ws.Name & "': " & Err.Description, vbCritical, "Erro"
End Sub

' ==============================================================================
' 4. FUNÇÕES AUXILIARES E UTILITÁRIAS
' ==============================================================================

' Localiza e seleciona a aba base principal
Private Function SelecionarAbaBaseVoltar(wb As Workbook) As String
    Dim nomes(4) As String
    Dim encontrados(4) As String
    Dim count As Long
    Dim i As Long
    Dim msg As String
    Dim escolha As String
    Dim idx As Long
    Dim ws As Worksheet
    
    nomes(0) = "Memorial de Cálculo"
    nomes(1) = "Orçamento"
    nomes(2) = "MC"
    nomes(3) = "Orcamento"
    nomes(4) = "Memorial de Calculo"
    
    count = 0
    For i = 0 To 4
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(nomes(i))
        On Error GoTo 0
        If Not ws Is Nothing Then
            encontrados(count) = nomes(i)
            count = count + 1
        End If
    Next i
    
    If count = 0 Then
        MsgBox "Nenhuma aba base reconhecida foi encontrada." & vbCrLf & _
               "O arquivo deve conter uma aba chamada:" & vbCrLf & _
               "• Memorial de Cálculo" & vbCrLf & _
               "• Orçamento" & vbCrLf & _
               "• MC", vbCritical, "Erro"
        SelecionarAbaBaseVoltar = ""
        Exit Function
    End If
    
    If count = 1 Then
        SelecionarAbaBaseVoltar = encontrados(0)
        Exit Function
    End If
    
    msg = "Mais de uma aba base foi encontrada no arquivo." & vbCrLf & vbCrLf
    For i = 0 To count - 1
        msg = msg & i + 1 & " - " & encontrados(i) & vbCrLf
    Next i
    msg = msg & vbCrLf & "Digite o número da aba principal para onde o botão 'VOLTAR' deve retornar:"
    
    escolha = InputBox(msg, "Selecionar Aba Base Destino", "1")
    
    If escolha = "" Or Not IsNumeric(escolha) Then
        SelecionarAbaBaseVoltar = ""
        Exit Function
    End If
    
    idx = CLng(escolha) - 1
    If idx < 0 Or idx >= count Then
        MsgBox "Opção inválida.", vbCritical, "Erro"
        SelecionarAbaBaseVoltar = ""
        Exit Function
    End If
    
    SelecionarAbaBaseVoltar = encontrados(idx)
End Function

' Remove formas antigas com o nome BotaoVoltar
Private Sub RemoverBotoesVoltarAntigos(ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        If ws.Shapes(i).Name Like "BotaoVoltar*" Then
            ws.Shapes(i).Delete
        End If
    Next i
End Sub


' Cria a forma geométrica e associa o hiperlink
Private Sub CriarBotaoVoltarNaAba(ws As Worksheet, wsBase As Worksheet)
    Dim celDestino As Range
    Dim shp As Shape
    Dim margem As Double
    Dim enderecoLink As String
    Dim subEndereco As String
    Dim origem As String
    
    Set celDestino = ws.Range("A3")
    margem = 2
    
    ' 1ª prioridade: célula da aba base que tem o link para esta aba
    origem = LocalizarCelulaOrigem(wsBase, ws.Name)
    
    ' 2ª prioridade: hiperlink pré-existente na célula A2
    If origem = "" Then
        If ws.Range("A2").Hyperlinks.count > 0 Then
            enderecoLink = ws.Range("A2").Hyperlinks(1).Address
            subEndereco = ws.Range("A2").Hyperlinks(1).SubAddress
        End If
    End If
    
    ' Adiciona o retângulo arredondado em A3
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                celDestino.Left + margem, _
                                celDestino.Top + margem, _
                                (celDestino.Width * 2) - (2 * margem), _
                                (celDestino.Height * 2) - (2 * margem))
    
    ' Configuração visual do botão
    shp.Name = "BotaoVoltar_" & ws.Name
    shp.Fill.ForeColor.RGB = RGB(255, 153, 51)   ' Laranja
    shp.Line.Weight = 1.5
    
    With shp.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = "VOLTAR"
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Size = 11
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)   ' Vermelho
    End With
    
    ' Associação da hiperligação de retorno
    If origem <> "" Then
        ws.Hyperlinks.Add Anchor:=shp, Address:="", SubAddress:=origem
    ElseIf subEndereco <> "" Then
        ws.Hyperlinks.Add Anchor:=shp, Address:=enderecoLink, SubAddress:=subEndereco
    ElseIf enderecoLink <> "" Then
        ws.Hyperlinks.Add Anchor:=shp, Address:=enderecoLink
    Else
        ws.Hyperlinks.Add Anchor:=shp, Address:="", SubAddress:="'" & wsBase.Name & "'!A1"
    End If
End Sub

' Procura na aba base a célula cujo hiperlink aponta para a aba informada
Private Function LocalizarCelulaOrigem(wsBase As Worksheet, nomeAba As String) As String
    Dim hl As Hyperlink
    Dim destino As String
    
    For Each hl In wsBase.Hyperlinks
        If hl.Type = msoHyperlinkRange Then   ' ignora links de formas
            destino = Replace(hl.SubAddress, "'", "")
            If Left(destino, Len(nomeAba) + 1) = nomeAba & "!" Then
                LocalizarCelulaOrigem = "'" & wsBase.Name & "'!" & hl.Range.Address(False, False)
                Exit Function
            End If
        End If
    Next hl
End Function
