Sub ResumoSegmentos()
    
    
'Fator D - Resumo Segmentos.xlsm

'Organize and filter the highway segments whitch Factor D should be applied,
'whether due to horizontal, vertical signage or both

'Organiza e filtra os segmentos rodoviários os quais o Fator D deverá ser aplicado,
'seja por conta da sinalização horizontal, vertical ou ambas

'Created by Matheus Nunes Reis on 27/10/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-FactorD/a9683ff1e55f7a31808dee2140e0e29ed33423de/LICENSE.md
'MIT License. Copyright © 2024 MatheusNReis

    
    'Verifica planilha de dados e transforma-a numa outra com linhas de kmincial não repetidas (exclusivos)
    Dim works As Worksheet
    Dim ColunaReferencia As String
    Dim NomePlanilha As String
    Dim LastRowPlanWorks As Long
    Dim linhaPlanCompilado As Long
    
    NomePlanilha = ThisWorkbook.Sheets("Informações").Cells(2, "C").Value
    
    
    If NomePlanilha = "" Then
        MsgBox "Informação 'Nome Planilha' não está preenchida."
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
    
    Dim Rodovia As String, kmInicial As String, kmFinal As String, Conc_Sup As String, Ano As String
    Rodovia = ThisWorkbook.Sheets("Informações").Cells(7, "B").Value
    kmInicial = ThisWorkbook.Sheets("Informações").Cells(7, "C").Value
    kmFinal = ThisWorkbook.Sheets("Informações").Cells(7, "D").Value
    Conc_Sup = ThisWorkbook.Sheets("Informações").Cells(7, "E").Value
    Ano = ThisWorkbook.Sheets("Informações").Cells(7, "F").Value
    
    If Rodovia = "" Then
        MsgBox "Informação da coluna 'Rodovia' não está preenchida."
        Exit Sub
    ElseIf kmInicial = "" Then
        MsgBox "Informação da coluna 'km Inicial' não está preenchida."
        Exit Sub
    ElseIf kmFinal = "" Then
        MsgBox "Informação da coluna 'km final' não está preenchida."
        Exit Sub
    ElseIf Conc_Sup = "" Then
        MsgBox "Informação da coluna 'Concessionária/Supervisora' não está preenchida."
        Exit Sub
    ElseIf Ano = "" Then
        MsgBox "Informação da coluna 'Ano' não está preenchida."
        Exit Sub
    End If
    
    
    LastRowPlanWorks = works.Cells(Rows.Count, kmInicial).End(xlUp).Row
    
    
    'Classifica planilha 'Resumo Segmentos' pela coluna 'km inicial'
    With works.Sort
        .SortFields.Clear
        .SortFields.Add Key:=works.Range(kmInicial & "1:" & kmInicial & LastRowPlanWorks), Order:=xlAscending 'Coluna 'km inicial'
        .SetRange works.Range("A1:Z" & LastRowPlanWorks)
        .Header = xlYes
        .Apply
    End With

    
    'Inicialização
    Dim i As Long, LinhaResSeg As Long
    i = 2 'i é linha na planilha works
    LinhaResSeg = 2 'Linha da planilha Resumo Segmentos
       
    Do While i <= LastRowPlanWorks
        
        ThisWorkbook.Sheets("Resumo Segmentos").Cells(LinhaResSeg, "A").Value = workb.Name
        ThisWorkbook.Sheets("Resumo Segmentos").Cells(LinhaResSeg, "B").Value = works.Cells(i, Rodovia).Value
        ThisWorkbook.Sheets("Resumo Segmentos").Cells(LinhaResSeg, "C").Value = works.Cells(i, kmInicial).Value
        ThisWorkbook.Sheets("Resumo Segmentos").Cells(LinhaResSeg, "D").Value = works.Cells(i, kmFinal).Value
        ThisWorkbook.Sheets("Resumo Segmentos").Cells(LinhaResSeg, "E").Value = works.Cells(i, Conc_Sup).Value
        ThisWorkbook.Sheets("Resumo Segmentos").Cells(LinhaResSeg, "F").Value = works.Cells(i, Ano).Value
        LinhaResSeg = LinhaResSeg + 1
        
        'Definição da próxima Linha i com kmInicial não repetido
        Do While works.Cells(i, kmInicial).MergeArea.Cells(1, 1).Value = works.Cells(i + 1, kmInicial).MergeArea.Cells(1, 1).Value
            i = i + 1
        Loop

        i = i + 1
    Loop
    
    MsgBox "Fim do Processo."
  
End Sub
