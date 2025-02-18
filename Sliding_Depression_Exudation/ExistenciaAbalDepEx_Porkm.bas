Sub ExistenciaAbalDepEx_Porkm()


'List the monitoring records for highway segments that have Sliding/Heaving, Settling/Depression
'or exudation, also showing unique km's in ascending order.
'Códigos conforme DNIT 007 PRO.

'Lista as fichas de monitoração dos trechos de rodovia que possuem escorregamento/abaulamento,
'afundamento/depressão ou exsudação, também apresentando km's exclusivos em ordem crescente.
'Codes according to DNIT 007 PRO.

'Created by Matheus Nunes Reis on 11/02/2025
'Copyright © 2025 Matheus Nunes Reis. All rights reserved.


    'Dados iniciais
    IntervaloTrincas = "F38:F116"
    KmInicialFicha = "C13" 'Ficha sentido crecente
    KmFinalFicha = "E13" 'Ficha sentido descrescente
    

    Dim ws As Worksheet
    Dim listaWs As Worksheet
    Dim valor As Variant
    Dim i As Long
    
    Set listaWs = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    'Definir a planilha "LISTA" onde os nomes das planilhas serão salvos
    listaWs.Name = "Existência_Abaul_Dep_Ex"
    'Inserir cabeçalho na planilha LISTA
    listaWs.Range("A1").Value = "Existência Abaulamento/Depressão/Exsudação(km)"

    'Inicializar o número da linha para a coluna A da planilha "LISTA" desconsiderando cabeçalho
    i = 2
    
    'Percorrer todas as planilhas na pasta de trabalho atual
    For Each ws In ThisWorkbook.Sheets
        
        If InStr(ws.Name, "PDC") > 0 Or InStr(ws.Name, "PS") > 0 Then 'Sentido crescente
            
            For Each valor In ws.Range(IntervaloTrincas).Value
                If valor = "ALC" Or valor = "ATC" Or valor = "ATP" Or valor = "ALP" Or valor = "E" Or valor = "EX" Then
                    'Se encontrar valor para abaulamento ou depressão, salvar o km na coluna A da planilha "Existência_Abaul_Dep_Ex"
                    listaWs.Cells(i, "A").Value = ws.Range(KmInicialFicha).MergeArea.Cells(1, 1).Value
                    i = i + 1
                    Exit For
                End If
            Next valor
        End If
        
        If InStr(ws.Name, "PDD") > 0 Then 'Sentido decrescente
            
            For Each valor In ws.Range(IntervaloTrincas).Value
                If valor = "ALC" Or valor = "ATC" Or valor = "ATP" Or valor = "ALP" Or valor = "E" Or valor = "EX" Then
                    'Se encontrar valor para abaulamento ou depressão, salvar o km na coluna A da planilha "Existência_Abaul_Dep_Ex"
                    listaWs.Cells(i, "A").Value = ws.Range(KmFinalFicha).MergeArea.Cells(1, 1).Value
                    i = i + 1
                    Exit For
                End If
            Next valor
        End If
        
    Next ws
    
    
    listaWs.Range("A1").Value = "Todos (km)"
    listaWs.Range("B1").Value = "Exclusivos (km)"

    LastRowResults = listaWs.Cells(listaWs.Rows.Count, "A").End(xlUp).Row 'Coluna A: Km's reprovados e repetidos
    
    
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
        listaWs.Cells(outputRow, 2).Value = uniqueValues(i) ' Column B é a coluna 2
        outputRow = outputRow + 1
    Next i
    '--
    
    
    LastRowResults2 = listaWs.Cells(listaWs.Rows.Count, "B").End(xlUp).Row
    Set rng = Range("B2:B" & LastRowResults2)
    rng.Sort Key1:=Range("B2"), Order1:=xlAscending, Header:=xlNo
    
    
    MsgBox "Processo concluído. km's com códigos de abaulamento, depressão ou exsudação foram registrados na planilha ""Existência_Abaul_Dep_Ex""."
    
    
End Sub


