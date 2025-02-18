Sub AreaTrincadaSegmentoPorkm()

    'Segment cracked area = FC1 + FC2 + FC3 in each km
    'Disapproved segment if cracked area is above a limit percentage of its area

    'Área trincada do segmento = FC1+FC2+FC3 a cada km
    'Segmento reprovado se a área trincada do segmento for acima de uma porcentagem limite de sua área
    
    'Created by Matheus Nunes Reis on 20/01/2025
    'Copyright © 2025 Matheus Nunes Reis. All rights reserved.
    
    
    'Dados Iniciais
    KmInicialFicha = "C13" 'Ficha sentido crescente
    KmFinalFicha = "E13" 'Ficha sentido decrescente
    LarguraPista = "A125"
    AreaTotalTrincadaSegmento = "M118"
    LimiteTrinca = 0.15 'Porcentagem limite
    
    
    Dim ws As Worksheet
    Set wsresult = ThisWorkbook.Sheets("Planilha1")
    
    Dim i As Long
    i = 2
    
    ' Percorre cada planilha na Pasta de Trabalho e verifica se o km é reprovado
    For Each ws In ThisWorkbook.Worksheets
        
        If InStr(ws.Name, "PDC") > 0 Or InStr(ws.Name, "PS") > 0 Then    'Sentido crescente
            
            AreaSegmento = Abs(ws.Range(KmFinalFicha).MergeArea.Cells(1, 1).Value - ws.Range(KmInicialFicha).MergeArea.Cells(1, 1).Value) * 1000 * _
                               ws.Range(LarguraPista).MergeArea.Cells(1, 1).Value
            
            TaxaAreaTrincadaSegmento = ws.Range(AreaTotalTrincadaSegmento).MergeArea.Cells(1, 1).Value / AreaSegmento
        
            If TaxaAreaTrincadaSegmento > LimiteTrinca Then
                wsresult.Cells(i, "A").Value = ws.Range(KmInicialFicha).MergeArea.Cells(1, 1).Value
                i = i + 1
            End If
        
        End If
        
        If InStr(ws.Name, "PDD") > 0 Then 'Sentido decrescente
        
            AreaSegmento = Abs(ws.Range(KmFinalFicha).MergeArea.Cells(1, 1).Value - ws.Range(KmInicialFicha).MergeArea.Cells(1, 1).Value) * 1000 * _
                               ws.Range(LarguraPista).MergeArea.Cells(1, 1).Value
            
            TaxaAreaTrincadaSegmento = ws.Range(AreaTotalTrincadaSegmento).MergeArea.Cells(1, 1).Value / AreaSegmento
        
            If TaxaAreaTrincadaSegmento > LimiteTrinca Then
                wsresult.Cells(i, "A").Value = ws.Range(KmFinalFicha).MergeArea.Cells(1, 1).Value
                i = i + 1
            End If
            
        End If
        
        
    Next ws

    wsresult.Range("A1").Value = "Todos (km)"
    wsresult.Range("B1").Value = "Exclusivos (km)"

    LastRowResults = wsresult.Cells(wsresult.Rows.Count, "A").End(xlUp).Row 'Coluna A: Km's reprovados e repetidos
    

    'Exibição dos resultados finais de km reprovados sem repetição
    '--
    Dim rng As Range
    Dim cell As Range
    Dim dict As Object
    Dim uniqueValues As Variant
    Dim outputRow As Long
    
    'Intervalo de km's reprovados sem
    Set rng = Range("A2:A" & LastRowResults)
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    'Loop em cada célula do intervalo
    For Each cell In rng
        If Not dict.exists(cell.Value) Then
            dict.Add cell.Value, Nothing
        End If
    Next cell
    
    'Retorna valores únicos na coluna B
    uniqueValues = dict.keys
    outputRow = 2
    For i = LBound(uniqueValues) To UBound(uniqueValues)
        wsresult.Cells(outputRow, 2).Value = uniqueValues(i) ' Column B é a coluna 2
        outputRow = outputRow + 1
    Next i
    '--


    LastRowResults2 = wsresult.Cells(wsresult.Rows.Count, "B").End(xlUp).Row
    Set rng = Range("B2:B" & LastRowResults2)
    rng.Sort Key1:=Range("B2"), Order1:=xlAscending, Header:=xlNo


    MsgBox "Fim da análise de trincas."

End Sub


