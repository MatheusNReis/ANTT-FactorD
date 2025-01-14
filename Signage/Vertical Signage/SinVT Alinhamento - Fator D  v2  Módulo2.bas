Sub CompararPlanilha2()


'SinVT Alinhamento.xlsm

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


    Dim workb As Workbook
    Dim works As Worksheet
    Dim NomePlanilha As String
    Dim LastRowPlanWorks As Long
    
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
    
    Dim Identificacao As String, km As String, Latitude As String, Longitude As String, PeliculaTipo As String
    Dim Cor As String, MediaRetrorrefletancia As String, MinimaRetrorrefletancia As String, Conc_Sup As String
    Dim Ano As Integer
    Identificacao = ThisWorkbook.Sheets("Informações").Cells(19, "B").Value
    km = ThisWorkbook.Sheets("Informações").Cells(19, "C").Value
    Latitude = ThisWorkbook.Sheets("Informações").Cells(19, "D").Value
    Longitude = ThisWorkbook.Sheets("Informações").Cells(19, "E").Value
    PeliculaTipo = ThisWorkbook.Sheets("Informações").Cells(19, "F").Value
    Cor = ThisWorkbook.Sheets("Informações").Cells(19, "G").Value
    MediaRetrorrefletancia = ThisWorkbook.Sheets("Informações").Cells(19, "H").Value
    MinimaRetrorrefletancia = ThisWorkbook.Sheets("Informações").Cells(19, "I").Value
    Conc_Sup = ThisWorkbook.Sheets("Informações").Cells(19, "J").Value
    Ano = Format(ThisWorkbook.Sheets("Informações").Cells(19, "K").Value, 0#)
    
    If Identificacao = "" Then
        MsgBox "Informação da coluna 'Identificação' não está preenchida."
        Exit Sub
    ElseIf km = "" Then
        MsgBox "Informação da coluna 'km' não está preenchida."
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
    
    
    Dim LastRowCompilado As Long
    LastRowCompilado = ThisWorkbook.Sheets("Compilado").Cells(Rows.Count, "A").End(xlUp).Row
    
    'Converte dados da coluna 'km' da planilha 'Compilado' em número
    Dim k As Long
    For k = 2 To LastRowCompilado
        If InStr(1, ThisWorkbook.Sheets("Compilado").Cells(k, "C").MergeArea.Cells(1, 1).Value, "+", vbTextCompare) > 0 Then
                ThisWorkbook.Sheets("Compilado").Cells(k, "C").Value = CDbl(Replace(ThisWorkbook.Sheets("Compilado").Cells(k, "C").MergeArea.Cells(1, 1).Value, "+", ","))
            Else
                ThisWorkbook.Sheets("Compilado").Cells(k, "C").Value = CDbl(ThisWorkbook.Sheets("Compilado").Cells(k, "C").MergeArea.Cells(1, 1).Value)
        End If
    Next k
    
    'Classifica planilha 'Compilado' pela coluna km
    With ThisWorkbook.Sheets("Compilado").Sort
        .SortFields.Clear
        .SortFields.Add Key:=ThisWorkbook.Sheets("Compilado").Range("C1:C" & LastRowCompilado), Order:=xlAscending 'Coluna C (km)
        .SetRange ThisWorkbook.Sheets("Compilado").Range("A1:Z" & LastRowCompilado)
        .Header = xlYes
        .Apply
    End With
    
    
    'Inicialização
    Dim i As Long
    i = 1 'i é linha na planilha works
      
    Do While (InStr(1, works.Cells(i, Identificacao).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) = 0) 'Percorre linhas até encontrar linha com TituloColunaChave
        i = i + 1
    Loop
    
    Do While (InStr(1, works.Cells(i, Identificacao).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) > 0) 'Percorre linhas até encontrar linha sem TituloColunaChave
        i = i + 1
    Loop
    
    
    LastRowPlanWorks = works.Cells(Rows.Count, Identificacao).End(xlUp).Row + 2 ' +2 para garantir que todas as linha sejam consideradas - compensar a mescla


    Dim StartRowCompilado As Long, midRowCompilado As Long, LastRowCompiladoFixa As Long
    Dim LinhaAdicional As Long, j As Long
    LastRowCompiladoFixa = LastRowCompilado
    LinhaAdicional = LastRowCompilado + 1
    
    For i = i To LastRowPlanWorks
        
        'Reinicialização
        StartRowCompilado = 2
        LastRowCompilado = LastRowCompiladoFixa
        
        Do While LastRowCompilado - StartRowCompilado > 0 'Até que sobre apenas uma linha

            ' converte em número o 'km' da planilha works
            If InStr(1, works.Cells(i, km).MergeArea.Cells(1, 1).Value, "+", vbTextCompare) > 0 Then
                kmValueWorks = CDbl(Replace(works.Cells(i, km).MergeArea.Cells(1, 1).Value, "+", ","))
            Else
                kmValueWorks = CDbl(works.Cells(i, km).MergeArea.Cells(1, 1).Value)
            End If
           
            ' Verifica se 'km' está na primeira ou última linha no intervalo atual
            If ThisWorkbook.Sheets("Compilado").Cells(StartRowCompilado, "C").Value = kmValueWorks Then 'Se km for encontrado na primeira linha do intervalo atual avaliado
            '--
                'Neste ponto, 'km' correspondente foi encontrado na linha inicial do intervalo atual
                'LinhaInicial deve ser >= 2
                If (StartRowCompilado - 30) < 2 Then
                    LinhaInicial = 2
                Else
                    LinhaInicial = (StartRowCompilado - 30)
                End If
                
                For j = LinhaInicial To (StartRowCompilado + 30)
                    If ThisWorkbook.Sheets("Compilado").Cells(j, "B").Value = works.Cells(i, Identificacao).MergeArea.Cells(1, 1).Value Then
                    'Ao chegar neste ponto é porque foi encontrada a 1ª linha na qual 'km' e 'identificação' são correspondentes
                        For k = j To (j + 2)
                            ThisWorkbook.Sheets("Compilado").Cells(k, "F").Value = works.Cells(i, PeliculaTipo).MergeArea.Cells(1, 1).Value 'Película Tipo
                            ThisWorkbook.Sheets("Compilado").Cells(k, "G").Value = works.Cells(i, Cor).MergeArea.Cells(1, 1).Value 'Cor
                            ThisWorkbook.Sheets("Compilado").Cells(k, "H").Value = works.Cells(i, MediaRetrorrefletancia).MergeArea.Cells(1, 1).Value 'média retrorrefletância
                            ThisWorkbook.Sheets("Compilado").Cells(k, "I").Value = works.Cells(i, MinimaRetrorrefletancia).MergeArea.Cells(1, 1).Value 'minima retrorrefletância
                            i = i + 1
                        Next k
                        i = i - 1 'Correção devido ao 'next i' que será aplicado em seguida
                        Exit Do
                    End If
                Next j
            '--
            ElseIf ThisWorkbook.Sheets("Compilado").Cells(LastRowCompilado, "C").Value = kmValueWorks Then
            '--
                'Neste ponto, 'km' correspondente foi encontrado na última linha do intervalo atual
                'LinhaInicial deve ser >= 2
                If (LastRowCompilado - 30) < 2 Then
                    LinhaInicial = 2
                Else
                    LinhaInicial = (LastRowCompilado - 30)
                End If
                
                For j = LinhaInicial To (LastRowCompilado + 30)
                    If ThisWorkbook.Sheets("Compilado").Cells(j, "B").Value = works.Cells(i, Identificacao).MergeArea.Cells(1, 1).Value Then
                    'Ao chegar neste ponto é porque foi encontrada a 1ª linha na qual 'km' e 'identificação' são correspondentes
                        For k = j To (j + 2)
                            ThisWorkbook.Sheets("Compilado").Cells(k, "F").Value = works.Cells(i, PeliculaTipo).MergeArea.Cells(1, 1).Value 'Película Tipo
                            ThisWorkbook.Sheets("Compilado").Cells(k, "G").Value = works.Cells(i, Cor).MergeArea.Cells(1, 1).Value 'Cor
                            ThisWorkbook.Sheets("Compilado").Cells(k, "H").Value = works.Cells(i, MediaRetrorrefletancia).MergeArea.Cells(1, 1).Value 'média retrorrefletância
                            ThisWorkbook.Sheets("Compilado").Cells(k, "I").Value = works.Cells(i, MinimaRetrorrefletancia).MergeArea.Cells(1, 1).Value 'minima retrorrefletância
                            i = i + 1
                        Next k
                        i = i - 1 'Correção devido ao 'next i' que será aplicado em seguida
                        Exit Do
                    End If
                Next j
            '--
            End If
            
        
            ' Divide o intervalo e determina linha central do intervalo
            midRowCompilado = (StartRowCompilado + LastRowCompilado) \ 2
        
            ' Verifica em qual parte está "km" procurado
            If ThisWorkbook.Sheets("Compilado").Cells(midRowCompilado, "C").Value = kmValueWorks Then
            '--
                'Neste ponto, 'km' correspondente foi encontrado na linha central do intervalo atual
                'LinhaInicial deve ser >= 2
                If (midRowCompilado - 30) < 2 Then
                    LinhaInicial = 2
                Else
                    LinhaInicial = (midRowCompilado - 30)
                End If
                
                For j = LinhaInicial To (midRowCompilado + 30)
                    If ThisWorkbook.Sheets("Compilado").Cells(j, "B").Value = works.Cells(i, Identificacao).MergeArea.Cells(1, 1).Value Then
                    'Ao chegar neste ponto é porque foi encontrada a 1ª linha na qual 'km' e 'identificação' são correspondentes
                        For k = j To (j + 2)
                            ThisWorkbook.Sheets("Compilado").Cells(k, "F").Value = works.Cells(i, PeliculaTipo).MergeArea.Cells(1, 1).Value 'Película Tipo
                            ThisWorkbook.Sheets("Compilado").Cells(k, "G").Value = works.Cells(i, Cor).MergeArea.Cells(1, 1).Value 'Cor
                            ThisWorkbook.Sheets("Compilado").Cells(k, "H").Value = works.Cells(i, MediaRetrorrefletancia).MergeArea.Cells(1, 1).Value 'média retrorrefletância
                            ThisWorkbook.Sheets("Compilado").Cells(k, "I").Value = works.Cells(i, MinimaRetrorrefletancia).MergeArea.Cells(1, 1).Value 'minima retrorrefletância
                            i = i + 1
                        Next k
                        i = i - 1 'Correção devido ao 'next i' que será aplicado em seguida
                        Exit Do
                    End If
                Next j
            '--
            ElseIf ThisWorkbook.Sheets("Compilado").Cells(midRowCompilado, "C").Value > kmValueWorks Then
                'Caso nem a linha inicial, final ou central tenham o 'km' procurado
                LastRowCompilado = midRowCompilado - 1
            Else
                StartRowCompilado = midRowCompilado + 1
            End If
            
        Loop
    
        If LastRowCompilado - StartRowCompilado <= 0 Then 'Se todas as linhas foram verificadas mas 'km' não foi encontrado
            'Adiciona as 3 linhas de informações da placa complementar
            For k = 1 To 3
                ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "A").Value = workb.Name
                ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "B").Value = works.Cells(i, Identificacao).MergeArea.Cells(1, 1).Value
                ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "C").Value = works.Cells(i, km).MergeArea.Cells(1, 1).Value
                ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "D").Value = CDbl(works.Cells(i, Latitude).MergeArea.Cells(1, 1).Value)
                ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "E").Value = CDbl(works.Cells(i, Longitude).MergeArea.Cells(1, 1).Value)
                ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "F").Value = works.Cells(i, PeliculaTipo).MergeArea.Cells(1, 1).Value
                ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "G").Value = works.Cells(i, Cor).MergeArea.Cells(1, 1).Value
                If works.Cells(i, MediaRetrorrefletancia).MergeArea.Cells(1, 1).Value = "" Then
                    ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "H").Value = 0
                Else
                    ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "H").Value = CDbl(works.Cells(i, MediaRetrorrefletancia).MergeArea.Cells(1, 1).Value)
                End If
                If works.Cells(i, MinimaRetrorrefletancia).MergeArea.Cells(1, 1).Value = "" Then
                    ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "I").Value = 0
                Else
                    ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "I").Value = CDbl(works.Cells(i, MinimaRetrorrefletancia).MergeArea.Cells(1, 1).Value)
                End If
                ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "J").Value = Conc_Sup
                ThisWorkbook.Sheets("Compilado").Cells(LinhaAdicional, "K").Value = Ano
                LinhaAdicional = LinhaAdicional + 1
                i = i + 1
            Next k
            i = i - 1 'Correção devido ao 'next i' que será aplicado em seguida
        End If
        
    Next i
    
    MsgBox "Fim do Processo."
    
End Sub
