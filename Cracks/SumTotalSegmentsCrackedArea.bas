Sub SomaValores_AreaTrincadaTotalSegmento()


'Calculates total cracked areas of 20 km highway segments, considering FC1, FC2 and FC3 cracks,
'separated by ascending and descending direction.

'Calcula área trincada total de segmentos de rodovia de 20 km cada, considerando trincas FC1, FC2 e FC3,
'separada em sentido crescente e decrescente.

'Created by Matheus Nunes Reis on 22/06/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-FactorD/f94aa23bda5eecc9f69e057f0860204d0f84459a/LICENSE.md
'MIT License. Copyright © 2024 MatheusNReis


''''Área trincada total do segmento = FC1+FC2+FC3 a cada 20km'''
    
    
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
                    Intervalo(i) = Intervalo(i) + ws.Range("M118").Value   ''Depende da célula
                End If
            Next i
        End If
    Next ws

    ' Exibe os resultados
    For i = 1 To 12   ''Depende da quantidade de intervalos
        wsResult.Range("D" & (7 + i)).Value = Intervalo(i)
    Next i




    ''''Sentido Decrescente
    

    ' Inicializa os intervalos
    For i = 1 To 12 ''Depende da quantidade de intervalos
        Intervalo(i) = 0
    Next i
    
    ' Percorre cada planilha na Pasta de Trabalho
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "PDD") > 0 Then   ''Depende da planilha "PDC"
            For i = 1 To 12   ''Depende da quantidade de intervalos
                If ws.Range("E13").Value >= 0 + (i - 1) * 20 And ws.Range("E13").Value < 0 + i * 20 Then ''Depende da célula
                    Intervalo(i) = Intervalo(i) + ws.Range("M118").Value   ''Depende da célula
                End If
            Next i
        End If
    Next ws

    ' Exibe os resultados
    For i = 1 To 12   ''Depende da quantidade de intervalos
        wsResult.Range("D" & (19 + i)).Value = Intervalo(i)   'Depende do número de intervalos + 7
    Next i

End Sub

