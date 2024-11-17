Sub SomaValoresAreaSegmento()

    
'Calculates area of ​​segments every 20 km, allowing for variation of highway width,
'separated into ascending and descending directions.

'Calcula área de segmentos a cada 20 km, podendo considerar variação de largura da rodovia,
'separado em sentido crescente e decrescente.

'Created by Matheus Nunes Reis on 18/06/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-FactorD/f94aa23bda5eecc9f69e057f0860204d0f84459a/LICENSE.md
'MIT License. Copyright © 2024 MatheusNReis

    
''''Área trincada do segmento =  Trecho*(Largura da pista) a cada 20km'''

    ''''Sentido Crescente
    
    Dim ws As Worksheet
    Dim Intervalo(1 To 12) As Double  ''Depende da quantidade de intervalos
    Dim i As Integer
    Set wsResult = ThisWorkbook.Sheets("Planilha1")

    ' Inicializa os intervalos
    For i = 1 To 12 ''Depende da quantidade de intervalos
        Intervalo(i) = 0
    Next i
    
    ' Percorre cada planilha na Pasta de Trabalho
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "PDC") > 0 Then   ''Depende da planilha "PDC"
            For i = 1 To 12   ''Depende da quantidade de intervalos
                If ws.Range("C13").Value >= 0 + (i - 1) * 20 And ws.Range("C13").Value < 0 + i * 20 Then ''Depende da célula e do kmInicial
                    Intervalo(i) = Intervalo(i) + (Abs(ws.Range("C13").Value - ws.Range("E13").Value)) * 1000 * ws.Range("A125").Value ''Depende das células
                End If
            Next i
        End If
    Next ws

    ' Exibe os resultados
    For i = 1 To 12   ''Depende da quantidade de intervalos
        wsResult.Range("F" & (7 + i)).Value = Intervalo(i)
    Next i




    ''''Sentido Decrescente
    

    ' Inicializa os intervalos
    For i = 1 To 12 ''Depende da quantidade de intervalos
        Intervalo(i) = 0
    Next i
    
    ' Percorre cada planilha na Pasta de Trabalho
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "PDD") > 0 Then   ''Depende da planilha "PDD"
            For i = 1 To 12   ''Depende da quantidade de intervalos
                If ws.Range("E13").Value >= 0 + (i - 1) * 20 And ws.Range("E13").Value < 0 + i * 20 Then ''Depende da célula e do kmInicial
                    Intervalo(i) = Intervalo(i) + (Abs(ws.Range("C13").Value - ws.Range("E13").Value)) * 1000 * ws.Range("A125").Value ''Depende das células
                End If
            Next i
        End If
    Next ws

    ' Exibe os resultados
    For i = 1 To 12   ''Depende da quantidade de intervalos
        wsResult.Range("F" & (19 + i)).Value = Intervalo(i)   'Depende do número de intervalos + 7
    Next i

End Sub

