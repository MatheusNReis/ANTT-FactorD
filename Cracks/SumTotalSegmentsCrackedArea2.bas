Sub SomaValores_AreaTrincadaTotalSegmento2()


'Calculates total cracked areas of ​20 km highway segments, considering FC1, FC2 and FC3 cracks,
'for each type of road (Single or dual Carriageway).

'Calcula área total trincada para segmentos de rodovia de 20 km cada, considerando trincas FC1, FC2 e FC3,
'para cada tipo de rodovia (pista simples ou dupla).

'Created by Matheus Nunes Reis on 23/06/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-FactorD/f94aa23bda5eecc9f69e057f0860204d0f84459a/LICENSE.md
'MIT License. Copyright © 2024 MatheusNReis


''''Área trincada total do segmento = FC1+FC2+FC3 a cada 20km'''
    
    
    Dim ws As Worksheet
    Dim Intervalo(1 To 11) As Double  ''Depende da quantidade de intervalos
    Dim i As Integer
    Set wsResult = ThisWorkbook.Sheets("Planilha1")

    ' Inicializa os intervalos
    For i = 1 To 11 ''Depende da quantidade de intervalos
        Intervalo(i) = 0
    Next i
    
    
 ''''Neste algoritmo somente as planilhas que iniciam com "PDD" tem a verificação pela célula E13 (If) enquanto os demais casos são verificados pela C13 (Else)
    ' Percorre cada planilha na Pasta de Trabalho
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "PDD") > 0 Then   ''Depende da planilha "PDD"
            For i = 1 To 11   ''Depende da quantidade de intervalos
                If ws.Range("E13").Value >= 380 + (i - 1) * 20 And ws.Range("E13").Value < 380 + i * 20 Then ''Depende da célula e do kmInicial
                    Intervalo(i) = Intervalo(i) + ws.Range("M118").Value   ''Depende da célula
                End If
            Next i
        Else
            For i = 1 To 11   ''Depende da quantidade de intervalos
                If ws.Range("C13").Value >= 380 + (i - 1) * 20 And ws.Range("C13").Value < 380 + i * 20 Then ''Depende da célula e do kmInicial
                    Intervalo(i) = Intervalo(i) + ws.Range("M118").Value   ''Depende da célula
                End If
            Next i
        End If
    Next ws

    ' Exibe os resultados
    For i = 1 To 11   ''Depende da quantidade de intervalos
        wsResult.Range("D" & (7 + i)).Value = Intervalo(i)
    Next i
    
End Sub
