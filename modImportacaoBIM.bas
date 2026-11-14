Attribute VB_Name = "modImportacaoBIM"
' Esta é a macro chamada pelo botão da Ribbon (onAction="ChamarMacro1")
Sub ChamarMacro1(control As IRibbonControl)
    ' Chama a rotina principal
    ImportarAbasBIM
End Sub

' Rotina principal de importação
Sub ImportarAbasBIM()
    Dim wbDestino As Workbook
    Dim wbOrigem As Workbook
    Dim caminhoArquivoOrigem As Variant
    Dim ws As Worksheet
    Dim totalAbas As Long
    Dim resposta As VbMsgBoxResult
    
    ' ATENÇÃO: Em suplementos (.xlam), usamos ActiveWorkbook para se referir
    ' à planilha do usuário que está aberta e visível na tela.
    Set wbDestino = ActiveWorkbook
    
    ' Valida se existe alguma planilha aberta no Excel para receber os dados
    If wbDestino Is Nothing Then
        MsgBox "Abra uma planilha antes de executar a importação.", vbCritical, "Nenhuma pasta de trabalho ativa"
        Exit Sub
    End If
    
    ' Abre a janela para selecionar o arquivo origem
    caminhoArquivoOrigem = Application.GetOpenFilename("Arquivos Excel (*.xls*), *.xls*", , "Selecione o arquivo origem")
    
    If caminhoArquivoOrigem = False Then
        MsgBox "Operação cancelada.", vbExclamation, "Cancelado"
        Exit Sub
    End If
    
    ' Otimização: desativa atualização de tela para rodar mais rápido
    Application.ScreenUpdating = False
    
    On Error GoTo TratarErro
    
    ' Abre o arquivo origem selecionado
    Set wbOrigem = Workbooks.Open(caminhoArquivoOrigem)
    
    ' Conta o número total de abas no arquivo origem
    totalAbas = wbOrigem.Worksheets.count
    
    ' Pergunta confirmação para o usuário
    resposta = MsgBox("O arquivo de origem contém " & totalAbas & " aba(s)." & vbCrLf & _
                     "Deseja copiar todas elas para a planilha ativa?", vbQuestion + vbYesNo, "Confirmação de Importação")
    
    If resposta = vbNo Then
        wbOrigem.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "Operação cancelada pelo usuário.", vbInformation, "Cancelado"
        Exit Sub
    End If
    
    ' Copia todas as abas da origem para o destino ativo
    For Each ws In wbOrigem.Worksheets
        ws.Copy After:=wbDestino.Sheets(wbDestino.Sheets.count)
    Next ws
    
    ' Fecha o arquivo origem sem salvar alterações
    wbOrigem.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    MsgBox "Todas as " & totalAbas & " abas foram copiadas com sucesso!", vbInformation, "Sucesso"
    Exit Sub

TratarErro:
    Application.ScreenUpdating = True
    If Not wbOrigem Is Nothing Then wbOrigem.Close SaveChanges:=False
    MsgBox "Ocorreu um erro durante a importação: " & Err.Description, vbCritical, "Erro"
End Sub
