Sub SinHZ_TachaTachao()


'SinHZ Tachas_Tachões - Fator D.xlsm

'Verifica a ausência de tachas e tachões na sinalização horizontal da rodovia

'Check for the absence of raised pavement markers on horizontal highway signage

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
    TituloColunaChave = ThisWorkbook.Sheets("Informações").Cells(3, "C").Value 'Ex: km
    
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
    
    Dim Rodovia As String, FaixaSinalizacao As String, Conc_Sup As String
    Dim Segmento As Double, Ano As Integer
    Dim kmInicial As Double, kmFinal As Double, QtdeIntervalo As Integer
    
    km = ThisWorkbook.Sheets("Informações").Cells(6, "B").Value
    Rodovia = ThisWorkbook.Sheets("Informações").Cells(6, "C").Value
    kmInicial = ThisWorkbook.Sheets("Informações").Cells(6, "D").Value
    kmFinal = ThisWorkbook.Sheets("Informações").Cells(6, "E").Value
    Segmento = ThisWorkbook.Sheets("Informações").Cells(6, "F").Value 'Tamanho do segmento em km
    FaixaSinalizacao = ThisWorkbook.Sheets("Informações").Cells(6, "G").Value
    Conc_Sup = ThisWorkbook.Sheets("Informações").Cells(6, "H").Value
    Ano = Format(ThisWorkbook.Sheets("Informações").Cells(6, "I").Value, 0#)
    
    QtdeIntervalo = WorksheetFunction.RoundUp((kmFinal - kmInicial) / Segmento, 0)
    
    
    If km = "" Then
        MsgBox "Informação da coluna 'km' não está preenchida."
        Exit Sub
    ElseIf Rodovia = "" Then
        MsgBox "Informação da coluna 'Rodovia' não está preenchida."
        Exit Sub
    ElseIf kmInicial = 0 Then
        MsgBox "Informação da coluna 'km Inicial' não está preenchida."
        Exit Sub
    ElseIf kmFinal = 0 Then
        MsgBox "Informação da coluna 'km Final' não está preenchida."
        Exit Sub
    ElseIf Segmento = 0 Then
        MsgBox "Informação da coluna 'Segmento' não está preenchida."
        Exit Sub
    ElseIf FaixaSinalizacao = "" Then
        MsgBox "Informação da coluna 'Faixa de Sinalização' não está preenchida."
        Exit Sub
    ElseIf Conc_Sup = "" Then
        MsgBox "Informação da coluna 'Concessionária/Supervisora' não está preenchida."
        Exit Sub
    ElseIf Ano = 0 Then
        MsgBox "Informação da coluna 'Ano' não está preenchida."
        Exit Sub
    End If
    

    
    'Inicialização
    Dim Intervalo() As Long  ''Depende da quantidade de intervalos
    ReDim Intervalo(1 To QtdeIntervalo)
    Dim j As Integer
    For j = 1 To QtdeIntervalo
        Intervalo(j) = 0
    Next j
    
    'Encontra posição inicial dos dados
    Dim i As Long
    i = 1 'i é linha na planilha works
      
    Do While (InStr(1, works.Cells(i, km).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) = 0)
        i = i + 1
    Loop
    
    Do While (InStr(1, works.Cells(i, km).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) > 0)
        i = i + 1
    Loop
    
    
    LastRowPlanWorks = works.Cells(Rows.Count, km).End(xlUp).Row
    
    
    For i = i To LastRowPlanWorks
    
        If InStr(1, works.Cells(i, km).MergeArea.Cells(1, 1).Value, "+", vbTextCompare) > 0 Then
                kmetro = CDbl(Replace(works.Cells(i, km).MergeArea.Cells(1, 1).Value, "+", ","))
            Else
                kmetro = CDbl(works.Cells(i, km).MergeArea.Cells(1, 1).Value)
            End If
            
        For j = 1 To QtdeIntervalo
            If kmetro >= kmInicial + (j - 1) * Segmento And kmetro < kmInicial + j * Segmento Then
                Intervalo(j) = 1 'Há tachas/tachões na planilha works
            End If
        Next j
        
    Next i
    
    For j = 1 To QtdeIntervalo
        If Intervalo(j) = 0 Then 'Se não há tachas/tachões na planilha works para o intervalo
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "A").Value = workb.Name
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "B").Value = "Ausência de Tachas/tachões"
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "C").Value = Rodovia
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "D") = kmInicial + (j - 1) * Segmento 'km Inicial
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "E") = kmInicial + j * Segmento 'km Final
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "F").Value = Conc_Sup
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "G").Value = Ano
            linhaPlanCompilado = linhaPlanCompilado + 1
        End If
    Next j
    
            
    MsgBox "Fim do Processo."

End Sub
