Sub CompararPlanilha2_Lista()


'SinVT Alinhamento - Fator D - Modo Lista.xlsm

'Alinha dados de sinalização vertical provindos de planilha complementar, substituindo
'e/ou adicionado os dados correspondentes à pasta de trabalho atual

'Align data of vertical signage from a complementary worksheet, replacing
'and/or adding the corresponding data to the current workbook

'Version: 2.0

'Created by Matheus Nunes Reis on 27/10/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-FactorD/a9683ff1e55f7a31808dee2140e0e29ed33423de/LICENSE.md
'MIT License. Copyright © 2024 MatheusNReis


    Dim works As Worksheet
    Dim NomePlanilha As String
    Dim LastRowPlanWorks As Long
    Dim linhaPlanCompilado As Long
    
    NomePlanilha = ThisWorkbook.Sheets("Informações").Cells(15, "C").Value
    TituloColunaChave = ThisWorkbook.Sheets("Informações").Cells(16, "C").Value 'Ex: Identificação
        
    If NomePlanilha = "" Then
        MsgBox "Informação 'Nome Planilha' não está preenchida."
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
    
    Dim Identificacao As String, Latitude As String, Longitude As String, PeliculaTipo As String
    Dim Cor As String, MediaRetrorrefletancia As String, MinimaRetrorrefletancia As String, Conc_Sup As String
    Dim Ano As Integer
    Identificacao = ThisWorkbook.Sheets("Informações").Cells(19, "B").Value
    Latitude = ThisWorkbook.Sheets("Informações").Cells(19, "C").Value
    Longitude = ThisWorkbook.Sheets("Informações").Cells(19, "D").Value
    PeliculaTipo = ThisWorkbook.Sheets("Informações").Cells(19, "E").Value
    Cor = ThisWorkbook.Sheets("Informações").Cells(19, "F").Value
    MediaRetrorrefletancia = ThisWorkbook.Sheets("Informações").Cells(19, "G").Value
    MinimaRetrorrefletancia = ThisWorkbook.Sheets("Informações").Cells(19, "H").Value
    Conc_Sup = ThisWorkbook.Sheets("Informações").Cells(19, "I").Value
    Ano = Format(ThisWorkbook.Sheets("Informações").Cells(19, "J").Value, 0#)
    
    If Identificacao = "" Then
        MsgBox "Informação da coluna 'Identificação' não está preenchida."
        Exit Sub
    ElseIf Latitude = "" Then
        MsgBox "Informação da coluna 'Latitude' não está preenchida."
        Exit Sub
    ElseIf Longitude = "" Then
        MsgBox "Informação da coluna 'Longitude' não está preenchida."
        Exit Sub
    ElseIf PeliculaTipo = "" Then
        MsgBox "Informação da coluna 'Pelicula Tipo' não está preenchida."
        Exit Sub
    ElseIf Cor = "" Then
        MsgBox "Informação da coluna 'Cor' não está preenchida."
        Exit Sub
    ElseIf MediaRetrorrefletancia = "" Then
        MsgBox "Informação da coluna 'Valor Média Retrorrefletância' não está preenchida."
        Exit Sub
    ElseIf MinimaRetrorrefletancia = "" Then
        MsgBox "Informação da coluna 'Mínima Retrorrefletância' não está preenchida."
        Exit Sub
    ElseIf Conc_Sup = "" Then
        MsgBox "Informação da coluna 'Concessionária/Supervisora' não está preenchida."
        Exit Sub
    ElseIf Ano = 0 Then
        MsgBox "Informação da coluna 'Ano' não está preenchida."
        Exit Sub
    End If
    
    'Inicialização
    Dim i As Long
    i = 1 'i é linha na planilha works
      
    Do While (InStr(1, works.Cells(i, Identificacao).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) = 0)
        i = i + 1
    Loop
    
    Do While (InStr(1, works.Cells(i, Identificacao).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) > 0)
        i = i + 1
    Loop
    
    LastRowPlanWorks = works.Cells(Rows.Count, Identificacao).End(xlUp).Row + 2 ' +2 para garantir que todas as linha sejam consideradas - compensar a mescla
    linhaPlanCompilado = ThisWorkbook.Sheets("Compilado").Cells(Rows.Count, "A").End(xlUp).Row
    LinhaAdicional = linhaPlanCompilado + 1

    For i = i To LastRowPlanWorks
    
        For j = 2 To linhaPlanCompilado
            If works.Cells(i, Identificacao).MergeArea.Cells(1, 1).Value = ThisWorkbook.Sheets("Compilado").Cells(j, "B") And _
                works.Cells(i, PeliculaTipo).MergeArea.Cells(1, 1).Value = ThisWorkbook.Sheets("Compilado").Cells(j, "E") And _
                works.Cells(i, Cor).MergeArea.Cells(1, 1).Value = ThisWorkbook.Sheets("Compilado").Cells(j, "F") Then
                
                    ThisWorkbook.Sheets("Compilado").Cells(j, "A").Value = workb.Name
                    ThisWorkbook.Sheets("Compilado").Cells(j, "B").Value = works.Cells(i, Identificacao).MergeArea.Cells(1, 1).Value
                    ThisWorkbook.Sheets("Compilado").Cells(j, "C").Value = CDbl(works.Cells(i, Latitude).MergeArea.Cells(1, 1).Value)
                    ThisWorkbook.Sheets("Compilado").Cells(j, "D").Value = CDbl(works.Cells(i, Longitude).MergeArea.Cells(1, 1).Value)
                    ThisWorkbook.Sheets("Compilado").Cells(j, "E").Value = works.Cells(i, PeliculaTipo).MergeArea.Cells(1, 1).Value
                    ThisWorkbook.Sheets("Compilado").Cells(j, "F").Value = works.Cells(i, Cor).MergeArea.Cells(1, 1).Value
                    ThisWorkbook.Sheets("Compilado").Cells(j, "G").Value = CDbl(works.Cells(i, MediaRetrorrefletancia).MergeArea.Cells(1, 1).Value)
                    ThisWorkbook.Sheets("Compilado").Cells(j, "H").Value = CDbl(works.Cells(i, MinimaRetrorrefletancia).MergeArea.Cells(1, 1).Value)
                    ThisWorkbook.Sheets("Compilado").Cells(j, "I").Value = Conc_Sup
                    ThisWorkbook.Sheets("Compilado").Cells(j, "J").Value = Ano
                    Exit For
            
            End If
        Next j
        
        If j > linhaPlanCompilado Then
            ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "A").Value = workb.Name
            ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "B").Value = works.Cells(i, Identificacao).MergeArea.Cells(1, 1).Value
            ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "C").Value = CDbl(works.Cells(i, Latitude).MergeArea.Cells(1, 1).Value)
            ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "D").Value = CDbl(works.Cells(i, Longitude).MergeArea.Cells(1, 1).Value)
            ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "E").Value = works.Cells(i, PeliculaTipo).MergeArea.Cells(1, 1).Value
            ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "F").Value = works.Cells(i, Cor).MergeArea.Cells(1, 1).Value
            ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "G").Value = CDbl(works.Cells(i, MediaRetrorrefletancia).MergeArea.Cells(1, 1).Value)
            ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "H").Value = CDbl(works.Cells(i, MinimaRetrorrefletancia).MergeArea.Cells(1, 1).Value)
            ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "I").Value = Conc_Sup
            ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "J").Value = Ano
            LinhaAdicional = LinhaAdicional + 1
        End If
          
    Next i
    
    MsgBox "Fim do Processo."

End Sub
