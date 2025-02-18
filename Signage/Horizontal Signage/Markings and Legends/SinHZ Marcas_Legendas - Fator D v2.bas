Sub SinHZ_Marcas_Legendas()


'SinHZ Marcas_Legendas - Fator D (1).xlsm

'Pre-evaluates markings and legends data of horizontal signage and
'reorganize them to calculate Factor D

'Pré-avalia dados de marcas e legendas de sinalização horizontal e os
'reorganiza para cálculo de Fator D

'Version: 2.0

'Created by Matheus Nunes Reis on 27/10/2024
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
    
    Dim Segmento As String, Rodovia As String, FaixaSinalizacao As String, MediaRetrorrefletancia As String, Conc_Sup As String
    Dim MinimaRetrorrefletancia As Double, ExtSegmento As Double
    Dim Ano As Integer
    Segmento = ThisWorkbook.Sheets("Informações").Cells(7, "B").Value
    km = ThisWorkbook.Sheets("Informações").Cells(7, "C").Value
    Rodovia = ThisWorkbook.Sheets("Informações").Cells(7, "D").Value
    FaixaSinalizacao = ThisWorkbook.Sheets("Informações").Cells(7, "E").Value
    MediaRetrorrefletancia = ThisWorkbook.Sheets("Informações").Cells(7, "F").Value
    MinimaRetrorrefletancia = ThisWorkbook.Sheets("Informações").Cells(7, "G").Value
    Conc_Sup = ThisWorkbook.Sheets("Informações").Cells(7, "H").Value
    Ano = Format(ThisWorkbook.Sheets("Informações").Cells(7, "I").Value, 0#)
    ExtSegmento = ThisWorkbook.Sheets("Informações").Cells(7, "J").Value 'Tamanho do segmento em km
    
    
    If Segmento = "" Then
        MsgBox "Informação da coluna 'Segmento' não está preenchida."
        Exit Sub
    ElseIf km = "" Then
        MsgBox "Informação da coluna 'km' não está preenchida."
        Exit Sub
    ElseIf Rodovia = "" Then
        MsgBox "Informação da coluna 'Rodovia' não está preenchida."
        Exit Sub
    ElseIf FaixaSinalizacao = "" Then
        MsgBox "Informação da coluna 'Faixa de Sinalização' não está preenchida."
        Exit Sub
    ElseIf MediaRetrorrefletancia = "" Then
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
    ElseIf ExtSegmento = 0 Then
        MsgBox "Informação da coluna 'Extensão Segmento' não está preenchida."
        Exit Sub
    End If
    
    
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
    
   
    LastRowPlanWorks = works.Cells(Rows.Count, MediaRetrorrefletancia).End(xlUp).Row
    
    For i = i To LastRowPlanWorks
    
        'Verificação referente à coluna referência
        If works.Cells(i, MediaRetrorrefletancia).Value < MinimaRetrorrefletancia And _
            works.Cells(i, FaixaSinalizacao).Value <> "" Then
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "A").Value = workb.Name
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "B").Value = works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "C").Value = works.Cells(i, km).MergeArea.Cells(1, 1).Value
            
            If InStr(1, works.Cells(i, km).MergeArea.Cells(1, 1).Value, "+", vbTextCompare) > 0 Then
                kmValue = CDbl(Replace(works.Cells(i, km).MergeArea.Cells(1, 1).Value, "+", ","))
            Else
                kmValue = CDbl(works.Cells(i, km).MergeArea.Cells(1, 1).Value)
            End If
            
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "D").Value = Application.WorksheetFunction.Floor(kmValue, ExtSegmento)
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "E").Value = Application.WorksheetFunction.Floor(kmValue, ExtSegmento) + ExtSegmento
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "F").Value = Rodovia
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "G").Value = works.Cells(i, FaixaSinalizacao).MergeArea.Cells(1, 1).Value
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "H").Value = works.Cells(i, MediaRetrorrefletancia).MergeArea.Cells(1, 1).Value
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "I").Value = MinimaRetrorrefletancia
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "J").Value = "Não atende"
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "K").Value = Conc_Sup
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "L").Value = Ano
            linhaPlanCompilado = linhaPlanCompilado + 1
        End If
        
    Next i
    
    MsgBox "Fim do Processo."

End Sub
