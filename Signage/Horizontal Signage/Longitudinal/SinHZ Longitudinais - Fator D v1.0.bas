Sub SinHZ()


'SinHZ Longitudinais - Fator D.xlsm

'Extracting and organizing data of longitudinal horizontal signage for creation
'of spreadsheet adapted to calculate Factor D

'Extração o organização de dados de sinalização horizontal longitudinal para criação
'de planilha adapatada para cálculo de Fator D

'Version: 1.0

'Created by Matheus Nunes Reis on 27/09/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-FactorD/a9683ff1e55f7a31808dee2140e0e29ed33423de/LICENSE.md
'MIT License. Copyright © 2024 MatheusNReis


    Dim works As Worksheet
    Dim ColunaReferencia As String
    Dim NomePlanilha As String
    Dim LastRowPlanWorks As Long
    Dim linhaPlanCompilado As Long
    
    NomePlanilha = ThisWorkbook.Sheets("Informações").Cells(2, "C").Value
    PalavraChave = ThisWorkbook.Sheets("Informações").Cells(3, "C").Value 'Ex: Trecho
    TituloColunaChave = ThisWorkbook.Sheets("Informações").Cells(4, "C").Value 'Ex: Segmento
    
    linhaPlanCompilado = ThisWorkbook.Sheets("Compilado").Cells(Rows.Count, "A").End(xlUp).Row + 1 'Para iniciar na 1ª linha em branco
    
    
    If NomePlanilha = "" Then
        MsgBox "Informação 'Nome Planilha' não está preenchida."
        Exit Sub
    ElseIf PalavraChave = "" Then
        MsgBox "Informação 'Palavra-Chave' não está preenchida."
        Exit Sub
    ElseIf TituloColunaChave = "" Then
        MsgBox "Informação 'Titulo Coluna Chave' não está preenchida."
        Exit Sub
    End If
    
    
    Dim found As Boolean
    found = False
    For Each wb In Workbooks
        For Each ws In wb.Worksheets
            If ws.Name = NomePlanilha Then
            
                Dim resposta As VbMsgBoxResult
                resposta = MsgBox("'" & NomePlanilha & "' encontrado na planilha '" & wb.Name & "'", vbOKCancel + vbQuestion, "Confirmação de Planilha")
                If resposta = vbCancel Then
                    Exit Sub
                End If
                
                Set workb = wb
                Set works = wb.Sheets(NomePlanilha) 'works é a planilha origem dos dados
                found = True
                Exit For
                
            End If
        Next ws
        If found Then Exit For
    Next wb
        
    If Not found Then
        MsgBox "Planilha '" & NomePlanilha & "' não encontrada nas planilhas abertas."
        Exit Sub
    End If
    
    Dim Segmento As String, Rodovia As String, FaixaSinalizacao As String, MediaSegmento As String, Conc_Sup As String
    Dim MinimaRetrorrefletancia As Double
    Dim Ano As Integer
    Segmento = ThisWorkbook.Sheets("Informações").Cells(7, "B").Value
    Rodovia = ThisWorkbook.Sheets("Informações").Cells(7, "C").Value
    FaixaSinalizacao = ThisWorkbook.Sheets("Informações").Cells(7, "D").Value
    MediaSegmento = ThisWorkbook.Sheets("Informações").Cells(7, "E").Value
    MinimaRetrorrefletancia = ThisWorkbook.Sheets("Informações").Cells(7, "F").Value
    Conc_Sup = ThisWorkbook.Sheets("Informações").Cells(7, "G").Value
    Ano = Format(ThisWorkbook.Sheets("Informações").Cells(7, "H").Value, 0#)
    
    If Segmento = "" Then
        MsgBox "Informação da coluna 'Segmento' não está preenchida."
        Exit Sub
    ElseIf Rodovia = "" Then
        MsgBox "Informação da coluna 'Rodovia' não está preenchida."
        Exit Sub
    ElseIf FaixaSinalizacao = "" Then
        MsgBox "Informação da coluna 'Faixa de Sinalização' não está preenchida."
        Exit Sub
    ElseIf MediaSegmento = "" Then
        MsgBox "Informação da coluna 'Valor Média Segmento' não está preenchida."
        Exit Sub
    ElseIf MinimaRetrorrefletancia = 0 Then
        MsgBox "Informação da coluna 'Mínima Retrorrefletância' não está preenchida."
        Exit Sub
    ElseIf Conc_Sup = "" Then
        MsgBox "Informação da coluna 'Concessionária/Supervisora' não está preenchida."
        Exit Sub
    ElseIf Ano = 0 Then
        MsgBox "Informação da coluna 'Ano' não está preenchida."
        Exit Sub
    End If
    
    
    'Dim NProcessamento As Long
    'Dim LastRowPlanCompilado As Long
    'LastRowPlanCompilado = ThisWorkbook.Sheets("Compilado").Cells(Rows.Count, "A").End(xlUp).Row
    'If LastRowPlanCompilado = 1 Then
    '    NProcessamento = 1
    'Else
    '    NProcessamento = ThisWorkbook.Sheets("Compilado").Cells(LastRowPlanCompilado, "A").Value + 1
    'End If
    
    
    'Inicialização
    Dim i As Long
    i = 1 'i é linha na planilha works
    Dim FirstLineTrecho As Long, LastLineTrecho As Long, FirstLineReferencia As Long, LastLineReferencia As Long
      
    Do While (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, PalavraChave, vbTextCompare) = 0) And _
                (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) = 0)
        i = i + 1
    Loop
    
    Do While (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, PalavraChave, vbTextCompare) = 0) And _
                (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) > 0)
        i = i + 1
    Loop
    
    FirstLineTrecho = i
    
    LastRowPlanWorks = works.Cells(Rows.Count, MediaSegmento).End(xlUp).Row
    
    For i = i To LastRowPlanWorks
    
        Do While (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, PalavraChave, vbTextCompare) > 0) And _
                    (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) = 0)
            i = i + 1
        Loop
        LastLineTrecho = i - 1
        FirstLineReferencia = i
        
        Do While (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, PalavraChave, vbTextCompare) = 0) And _
                    (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) = 0)
            i = i + 1
            If i > LastRowPlanWorks Then
                Exit Do
            End If
        Loop
        
        LastLineReferencia = i - 1
        
        
        'Verificação referente à coluna referência
        For j = FirstLineReferencia To LastLineReferencia
            If works.Cells(j, MediaSegmento).Value < MinimaRetrorrefletancia Then
                ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "A").Value = workb.Name
                ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "B").Value = works.Cells(FirstLineTrecho, Segmento).MergeArea.Cells(1, 1).Value
                ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "C").Value = Rodovia
                ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "D").Value = works.Cells(j, FaixaSinalizacao).MergeArea.Cells(1, 1).Value
                ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "E").Value = works.Cells(j, MediaSegmento).MergeArea.Cells(1, 1).Value
                ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "F").Value = MinimaRetrorrefletancia
                ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "G").Value = "Não atende"
                ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "H").Value = Conc_Sup
                ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "I").Value = Ano
                linhaPlanCompilado = linhaPlanCompilado + 1
            End If
        Next j
        
        Do While (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, PalavraChave, vbTextCompare) = 0) And _
                (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) > 0)
        i = i + 1
        Loop
    
        FirstLineTrecho = i
        
    Next i
    
    MsgBox "Fim do Processo."

End Sub
