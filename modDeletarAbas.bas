Attribute VB_Name = "modDeletarAbas"
Option Explicit

' ==============================================================================
' 1. GATILHO DA RIBBON
' ==============================================================================
Public Sub RibbonDeletarTodasExcetoBase(control As IRibbonControl)
    ExecutarDeletarTodasExcetoBase
End Sub

' ==============================================================================
' 2. CHAMADA DIRETA / TESTE VIA ALT+F8
' ==============================================================================
Public Sub ChamarDeletarTodasExcetoBase()
    ExecutarDeletarTodasExcetoBase
End Sub

' ==============================================================================
' 3. LÓGICA REAL ISOLADA
' ==============================================================================
Public Sub ExecutarDeletarTodasExcetoBase()
    Dim wbDestino As Workbook
    Dim ws As Worksheet
    Dim wsBase As Worksheet
    Dim i As Long
    Dim contagem As Long
    Dim resposta As VbMsgBoxResult
    Dim colQuant As Long
    Dim c As Range
    Dim nomeSelecionado As String
    
    Set wbDestino = ActiveWorkbook
    If wbDestino Is Nothing Then Exit Sub
    
    If wbDestino.Name Like "*.xlam" Then
        MsgBox "Abra o arquivo de orçamento antes de executar.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' 1. Seleciona a aba base automaticamente
    nomeSelecionado = SelecionarAbaBase(wbDestino)
    If nomeSelecionado = "" Then Exit Sub
    
    Set wsBase = wbDestino.Worksheets(nomeSelecionado)
    
    ' 2. Confirmação final
    resposta = MsgBox("ATENÇÃO: Todas as abas serão APAGADAS, exceto '" & wsBase.Name & "'." & vbCrLf & vbCrLf & _
                      "Deseja continuar?", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Exclusão de Abas")
    If resposta = vbNo Then Exit Sub
    
    On Error GoTo TratarErro
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ' 3. Torna a aba base visível
    wsBase.Visible = xlSheetVisible
    
    ' 4. Localiza coluna Quant.
    colQuant = 6
    For Each c In wsBase.Range("A1:Z15")
        If InStr(1, CStr(c.Value), "Quant.", vbTextCompare) > 0 Then
            colQuant = c.Column
            Exit For
        End If
    Next c
    
    ' 5. Limpa hiperlinks preservando formatação
    LimparHiperlinksPreservandoFormatacao wsBase, "A"
    LimparHiperlinksPreservandoFormatacao wsBase, colQuant
    
    ' 6. Exclui todas as outras abas
    For i = wbDestino.Worksheets.count To 1 Step -1
        Set ws = wbDestino.Worksheets(i)
        If ws.Name <> wsBase.Name Then
            ws.Delete
            contagem = contagem + 1
        End If
    Next i

SairLimpo:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox contagem & " aba(s) removidas com sucesso!" & vbCrLf & _
           "Apenas '" & wsBase.Name & "' foi preservada.", _
           vbInformation, "Limpeza Concluída"
    Exit Sub

TratarErro:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Erro durante execução: " & Err.Description, vbCritical, "Erro"
    Resume SairLimpo
End Sub

' ==============================================================================
' 4. SELETOR DE ABA BASE
' ==============================================================================
Private Function SelecionarAbaBase(wb As Workbook) As String
    Dim nomes(4) As String
    Dim encontrados(4) As String
    Dim count As Long
    Dim i As Long
    Dim msg As String
    Dim escolha As String
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
    
    ' Nenhuma encontrada
    If count = 0 Then
        MsgBox "Nenhuma aba base reconhecida foi encontrada." & vbCrLf & _
               "O arquivo deve conter uma aba chamada:" & vbCrLf & _
               "• Memorial de Cálculo" & vbCrLf & _
               "• Orçamento" & vbCrLf & _
               "• MC", vbCritical, "Erro"
        SelecionarAbaBase = ""
        Exit Function
    End If
    
    ' Apenas uma — seleciona automaticamente
    If count = 1 Then
        SelecionarAbaBase = encontrados(0)
        Exit Function
    End If
    
    ' Mais de uma — lista as opções
    msg = "Mais de uma aba base foi encontrada." & vbCrLf & vbCrLf
    For i = 0 To count - 1
        msg = msg & i + 1 & " - " & encontrados(i) & vbCrLf
    Next i
    msg = msg & vbCrLf & "Digite o número da aba que deseja manter:"
    
    escolha = InputBox(msg, "Selecionar Aba Base", "1")
    
    If escolha = "" Or Not IsNumeric(escolha) Then
        SelecionarAbaBase = ""
        Exit Function
    End If
    
    Dim idx As Long
    idx = CLng(escolha) - 1
    If idx < 0 Or idx >= count Then
        MsgBox "Opção inválida.", vbCritical, "Erro"
        SelecionarAbaBase = ""
        Exit Function
    End If
    
    SelecionarAbaBase = encontrados(idx)
End Function

' ==============================================================================
' 5. LIMPA HIPERLINKS PRESERVANDO FORMATAÇÃO
' ==============================================================================
Private Sub LimparHiperlinksPreservandoFormatacao(ws As Worksheet, col As Variant)
    Dim hl As Hyperlink
    Dim cel As Range
    Dim i As Long, b As Long
    Dim corFundo As Long
    Dim temFundo As Boolean
    Dim indicesBorda(1 To 4) As Long
    Dim estiloBorda(1 To 4) As Long
    Dim corBorda(1 To 4) As Long
    Dim pesoBorda(1 To 4) As Long
    Dim temBorda(1 To 4) As Boolean
    
    indicesBorda(1) = xlEdgeLeft
    indicesBorda(2) = xlEdgeRight
    indicesBorda(3) = xlEdgeTop
    indicesBorda(4) = xlEdgeBottom
    
    For i = ws.Columns(col).Hyperlinks.count To 1 Step -1
        Set hl = ws.Columns(col).Hyperlinks(i)
        Set cel = hl.Range
        
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
        
        ' Remove hiperlink
        hl.Delete
        
        ' Reseta fonte
        cel.Font.Color = RGB(0, 0, 0)
        cel.Font.Underline = xlUnderlineStyleNone
        cel.Font.Bold = False
        
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
        
    Next i
End Sub
