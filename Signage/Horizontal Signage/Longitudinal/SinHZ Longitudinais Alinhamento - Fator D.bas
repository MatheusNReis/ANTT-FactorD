Sub SinHZ_Longitudinal_SubstituiComplementar()


'SinHZ Longitudinais Alinhamento - Fator D.xlsm

'Align data of longitudinal horizontal signage by replacing and/or
'supplementing data from the complementary monitoring spreadsheet.

'Alinha dados de sinalização horizontal longitudinal substituindo
'e/ou acrescentando dados provindos de planilha complementar de monitoração

'Created by Matheus Nunes Reis on 27/10/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-FactorD/a9683ff1e55f7a31808dee2140e0e29ed33423de/LICENSE.md
'MIT License. Copyright © 2024 MatheusNReis


    Dim works As Worksheet
    Dim NomePlanilha As String
    Dim LastRowPlanWorks As Long
    Dim linhaPlanCompilado As Long
    
    NomePlanilha = ThisWorkbook.Sheets("Informações").Cells(2, "C").Value
    PalavraChave = ThisWorkbook.Sheets("Informações").Cells(3, "C").Value 'Ex: Trecho
    TituloColunaChave = ThisWorkbook.Sheets("Informações").Cells(4, "C").Value 'Ex: Segmento
    
    
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
                Set works = wb.Sheets(NomePlanilha) 'works é a planilha COMPLEMENTAR de origem dos dados
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
    
    Dim Segmento As String, EstacaoMedicao As String, ColunaInicial_Intervalo As String, ColunaFinal_Intervalo As String, FaixaSinalizacao As String
    Segmento = ThisWorkbook.Sheets("Informações").Cells(7, "B").Value
    EstacaoMedicao = ThisWorkbook.Sheets("Informações").Cells(7, "C").Value
    ColunaInicial_Intervalo = ThisWorkbook.Sheets("Informações").Cells(7, "D").Value
    ColunaFinal_Intervalo = ThisWorkbook.Sheets("Informações").Cells(7, "E").Value
    FaixaSinalizacao = ThisWorkbook.Sheets("Informações").Cells(7, "F").Value 'Necessário para definir número total de linhas da planilha works
    
    
    If Segmento = "" Then
        MsgBox "Informação da coluna 'Segmento' não está preenchida."
        Exit Sub
    ElseIf EstacaoMedicao = "" Then
        MsgBox "Informação da coluna 'Estação Medição' não está preenchida."
        Exit Sub
    ElseIf ColunaInicial_Intervalo = "" Then
        MsgBox "Informação da coluna 'Coluna Inicial Intervalo' não está preenchida."
        Exit Sub
    ElseIf ColunaFinal_Intervalo = "" Then
        MsgBox "Informação da coluna 'Coluna Final Intervalo' não está preenchida."
        Exit Sub
    End If
    
    
    
    'Inicialização
    Dim i As Long
    i = 1 'i é linha na planilha works (Complementar)
      
    Do While (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, PalavraChave, vbTextCompare) = 0) And _
                (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) = 0)
        i = i + 1
    Loop
    
    Do While (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, PalavraChave, vbTextCompare) = 0) And _
                (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) > 0)
        i = i + 1
    Loop
    
    
    LastRowPlanWorks = works.Cells(Rows.Count, FaixaSinalizacao).End(xlUp).Row
    
    
    Dim FirstLineBloco As Long, LastLineBloco As Long 'primeira e última linhas do bloco contendo todas as informações de um segmento, na planilha complementar
    Dim LastRowPlanComplementada As Long
    LastRowPlanComplementada = ThisWorkbook.Sheets("Complementada").Cells(Rows.Count, FaixaSinalizacao).End(xlUp).Row
    
    For i = i To LastRowPlanWorks 'Works é a planilha complementar
        
        'Define 1ª linha do bloco da planilha 'Complementar'
        FirstLineBloco = i
        'captar km da estação de medição (kmEstacao)
        If InStr(1, works.Cells(FirstLineBloco, EstacaoMedicao).MergeArea.Cells(1, 1).Value, "+", vbTextCompare) > 0 Then
            kmEstacao = CDbl(Replace(works.Cells(FirstLineBloco, EstacaoMedicao).MergeArea.Cells(1, 1).Value, "+", ","))
        Else
            kmEstacao = CDbl(works.Cells(FirstLineBloco, EstacaoMedicao).MergeArea.Cells(1, 1).Value)
        End If
        
        'Percorrer o bloco para definir sua linha final, na planilha complementar
        Do While (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, PalavraChave, vbTextCompare) > 0) And _
                    (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) = 0)
            i = i + 1
        Loop
        
        Do While (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, PalavraChave, vbTextCompare) = 0) And _
                    (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) = 0)
            i = i + 1
            If i > LastRowPlanWorks Then
                Exit Do
            End If
        Loop
        'Linha final do bloco da planilha 'Complementar'
        LastLineBloco = i - 1
        
        'Do while caso existam títulos coluna chaves repetidas vezes ao longo das linhas
        Do While (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, PalavraChave, vbTextCompare) = 0) And _
                (InStr(1, works.Cells(i, Segmento).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) > 0)
        i = i + 1
        Loop
        
        'Verifica se o trecho da planilha complementar se encontra na planilha1
        found = False
        For j = 1 To LastRowPlanComplementada 'j é linha na planilha 'Complementada'
            
             'A seguir, On Error Resume Next foi aplicado para contornar o problema de haver non-numeric string no input de CDbl
            If InStr(1, ThisWorkbook.Sheets("Complementada").Cells(j, EstacaoMedicao).MergeArea.Cells(1, 1).Value, "+", vbTextCompare) > 0 Then
                On Error Resume Next 'error handling: Ignore runtime errors and allow the code to continue
                kmEstacaoCompl = CDbl(Replace(ThisWorkbook.Sheets("Complementada").Cells(j, EstacaoMedicao).MergeArea.Cells(1, 1).Value, "+", ",")) 'kmEstacaoCompl é o km da estação de medição da planilha 'Complementada'
                If Err.Number <> 0 Then
                    Err.Clear 'clears the error so that it doesn’t affect subsequent code
                    On Error GoTo 0 'Resets error handling to the default behavior
                    GoTo NextIteration
                End If
                On Error GoTo 0
            Else
                On Error Resume Next
                kmEstacaoCompl = CDbl(ThisWorkbook.Sheets("Complementada").Cells(j, EstacaoMedicao).MergeArea.Cells(1, 1).Value)
                If Err.Number <> 0 Then
                    Err.Clear
                    On Error GoTo 0
                    GoTo NextIteration
                End If
            End If
            
            
            If kmEstacaoCompl = kmEstacao Then
                'O trecho na planilha complementar foi encontrado na planilha 'Complementada'
                'bloco da planilha complementar é então substituído na 'Complementada'
                IntervaloPlanComplementada = (ColunaInicial_Intervalo & j) & ":" & (ColunaFinal_Intervalo & (j + LastLineBloco - FirstLineBloco)) 'j é linha na planilha 'Complementada'
                IntervaloPlanComplementar = (ColunaInicial_Intervalo & FirstLineBloco) & ":" & (ColunaFinal_Intervalo & LastLineBloco)
                ThisWorkbook.Sheets("Complementada").Range(IntervaloPlanComplementada).Value = works.Range(IntervaloPlanComplementar).Value
                found = True 'O trecho foi encontrado
                Exit For
            End If
            
NextIteration: 'Faz parte do loop for j Next j

        Next j
          
        
        If Not found Then
            'Se o trecho não foi encontrado, adiciona o bloco do trecho na planilha 'Complementada'
            LastRowPlanComplementada_2 = ThisWorkbook.Sheets("Complementada").Cells(Rows.Count, FaixaSinalizacao).End(xlUp).Row 'Última linha contando com blocos adicionados, para não interferir na LastRowPlanComplementada, somado +1 para identificar 1ª linha em branco
            IntervaloPlanComplementada = (ColunaInicial_Intervalo & LastRowPlanComplementada_2) & ":" & (ColunaFinal_Intervalo & (LastRowPlanComplementada_2 + LastLineBloco - FirstLineBloco))
            IntervaloPlanComplementar = (ColunaInicial_Intervalo & FirstLineBloco) & ":" & (ColunaFinal_Intervalo & LastLineBloco)
            ThisWorkbook.Sheets("Complementada").Range(IntervaloPlanComplementada).Value = works.Range(IntervaloPlanComplementar).Value
        End If
               
               
        i = i - 1 'Correção devido ao 'next i' que será aplicado em seguida
        
    Next i
    
    MsgBox "Fim do Processo."

End Sub
