Sub SinVT()


'SinVT - Fator D.xlsm

'Pre-evaluation and organization of vertical-signage data to calculate Factor D

'Pré-avaliação e organização dos dados de sinalização vertical para cálculo de Fator D

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
    TituloColunaChave = ThisWorkbook.Sheets("Informações").Cells(3, "C").Value 'Ex: Identificação
    
    linhaPlanCompilado = ThisWorkbook.Sheets("Compilado").Cells(Rows.Count, "A").End(xlUp).Row + 1 'Para iniciar na 1ª linha em branco
    
    
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
    Dim kmInicial As Double, kmFinal As Double, QtdeIntervalo As Integer
    Identificacao = ThisWorkbook.Sheets("Informações").Cells(6, "B").Value
    km = ThisWorkbook.Sheets("Informações").Cells(6, "C").Value
    Latitude = ThisWorkbook.Sheets("Informações").Cells(6, "D").Value
    Longitude = ThisWorkbook.Sheets("Informações").Cells(6, "E").Value
    PeliculaTipo = ThisWorkbook.Sheets("Informações").Cells(6, "F").Value
    Cor = ThisWorkbook.Sheets("Informações").Cells(6, "G").Value
    MediaRetrorrefletancia = ThisWorkbook.Sheets("Informações").Cells(6, "H").Value
    MinimaRetrorrefletancia = ThisWorkbook.Sheets("Informações").Cells(6, "I").Value
    Conc_Sup = ThisWorkbook.Sheets("Informações").Cells(6, "J").Value
    Ano = Format(ThisWorkbook.Sheets("Informações").Cells(6, "K").Value, 0#)
    Rodovia = ThisWorkbook.Sheets("Informações").Cells(6, "L").Value
    kmInicial = ThisWorkbook.Sheets("Informações").Cells(6, "M").Value
    kmFinal = ThisWorkbook.Sheets("Informações").Cells(6, "N").Value
    Segmento = ThisWorkbook.Sheets("Informações").Cells(6, "O").Value 'Tamanho do segmento em km
    
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
    ElseIf Rodovia = "" Then
        MsgBox "Informação da coluna 'Rodovia' não está preenchida."
        Exit Sub
    ElseIf kmInicial = 0 Then
        response = MsgBox("km inicial é 0. Continuar?", vbOKCancel + vbQuestion, "Confirme ação")
        If response = vbCancel Then
            Exit Sub
        End If
    ElseIf kmFinal = 0 Then
        response = MsgBox("km final é 0. Continuar?", vbOKCancel + vbQuestion, "Confirme ação")
        If response = vbCancel Then
            Exit Sub
        End If
    ElseIf Segmento = 0 Then
        MsgBox "Informação da coluna 'Extensão Segmento' não está preenchida."
        Exit Sub
    End If
    
    
    QtdeIntervalo = WorksheetFunction.RoundUp((kmFinal - kmInicial) / Segmento, 0)
    
    'Inicialização Intervalo
    Dim Intervalo() As Long  ''Depende da quantidade de intervalos
    ReDim Intervalo(1 To QtdeIntervalo)
    Dim j As Integer
    For j = 1 To QtdeIntervalo
        Intervalo(j) = 1 'Considerando inicialmente a existência das placas e atendimento ao parâmetro
    Next j
    
    'Inicialização
    Dim i As Long
    i = 1 'i é linha na planilha works
      
    Do While (InStr(1, works.Cells(i, Identificacao).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) = 0)
        i = i + 1
    Loop
    
    Do While (InStr(1, works.Cells(i, Identificacao).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) > 0)
        i = i + 1
    Loop
    
    
    LastRowPlanWorks = works.Cells(Rows.Count, Identificacao).End(xlUp).Row + 4 ' +4 para garantir que todas as linha sejam consideradas - compensar a mescla
    
    Dim LinhaInicial As Long, LinhaFinal As Long
    LinhaInicial = i
    
    Do While i <= LastRowPlanWorks
    
        ContLinhaVazia = 0
        NaoAtende = 0
    
        'Verificação necessária caso lastRowPlanWorks não tenha sido atingido mas não tenham mais placas
        If works.Cells(i, Identificacao).MergeArea.Cells(1, 1).Value = "" Then
            Exit Do
        End If
    
        'Definição da LinhaFinal
        Do While works.Cells(i, Identificacao).MergeArea.Cells(1, 1).Value = works.Cells(i + 1, Identificacao).MergeArea.Cells(1, 1).Value
            i = i + 1
            LinhaFinal = i
        Loop
        
        'Verificação de ausência de placa
        For k = LinhaInicial To LinhaFinal
            If works.Cells(k, PeliculaTipo).Value = "" Then
                ContLinhaVazia = ContLinhaVazia + 1
            End If
        Next k
        If ContLinhaVazia = (LinhaFinal - LinhaInicial + 1) Then
                NaoAtende = 1 'Todas as linhas entre LinhaInicial e LinhaFinal estão sem informação - placa ausente ou removida
        End If
        
        'Verificação de atendimento à mínima retrorrefletância caso exista placa
        If NaoAtende = 0 Then '0 indica existência de placa
            For k = LinhaInicial To LinhaFinal
                If works.Cells(k, PeliculaTipo).MergeArea.Cells(1, 1).Value <> "" Then 'Verifica o não atendimento somente se 'PeliulaTipo <> ""'
                    If CDbl(works.Cells(k, MediaRetrorrefletancia).MergeArea.Cells(1, 1).Value) < CDbl(works.Cells(k, MinimaRetrorrefletancia).MergeArea.Cells(1, 1).Value) Then
                        NaoAtende = 1 'Existe placa e não atende ao parâmetro
                    End If
                End If
            Next k
        End If
        
        'Ocorrendo o não atendimento 'NaoAtende = 1'
        If NaoAtende = 1 Then
            'Verifica segmento com placa reprovada
            If InStr(1, works.Cells(i, km).MergeArea.Cells(1, 1).Value, "+", vbTextCompare) > 0 Then
                kmetro = CDbl(Replace(works.Cells(i, km).MergeArea.Cells(1, 1).Value, "+", ","))
            Else
                kmetro = CDbl(works.Cells(i, km).MergeArea.Cells(1, 1).Value)
            End If
            
            For j = 1 To QtdeIntervalo
                If kmetro >= kmInicial + (j - 1) * Segmento And kmetro < kmInicial + j * Segmento Then
                    Intervalo(j) = 0 'Segmento com placa ausente ou placa que não atende ao parâmetro
                End If
            Next j
        End If
        
        i = i + 1
        LinhaInicial = i
        
    Loop
        
    For j = 1 To QtdeIntervalo
        If Intervalo(j) = 0 Then 'Segmento com placa ausente ou placa que não atende ao parâmetro
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "A").Value = workb.Name
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "B").Value = "Placa ausente/Não atende"
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "C").Value = Rodovia
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "D") = kmInicial + (j - 1) * Segmento 'km Inicial
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "E") = kmInicial + j * Segmento 'km Final
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "F").Value = Conc_Sup
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "G").Value = Ano
            linhaPlanCompilado = linhaPlanCompilado + 1
        End If
    Next j
        
    MsgBox "Fim do Processo"
      
End Sub
