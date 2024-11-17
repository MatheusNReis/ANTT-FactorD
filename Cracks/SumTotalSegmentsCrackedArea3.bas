Sub SomaValores_AreaTotalTrincadaSegmento3()


'Calculates the cracked area of highway segments considering FC1, FC2 and FC3 cracks.
'The segment sizes are defined by the user.

'Calcula área trincada de segmentos de rodovia, considerando trincas FC1, FC2 e FC3.
'Os tamanhos dos segmentos são definidos pelo usuário.

'Created by Matheus Nunes Reis on 21/07/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-FactorD/f94aa23bda5eecc9f69e057f0860204d0f84459a/LICENSE.md
'MIT License. Copyright © 2024 MatheusNReis


    'Área trincada do segmento = FC1+FC2+FC3 a cada 20km'''
    
    Dim Segmento As Integer
    Dim kmInicial As Integer
    Dim kmFinal As Integer
    Dim QtdeIntervalo As Integer
    'Preencher os dados
    Segmento = 20 'Tamanho do segmento em km
    kmInicial = 495
    kmFinal = 524
    
    QtdeIntervalo = WorksheetFunction.RoundUp((kmFinal - kmInicial) / Segmento, 0)
    
    Dim ws As Worksheet
    Dim Intervalo() As Double  ''Depende da quantidade de intervalos
    ReDim Intervalo(1 To QtdeIntervalo)
    Dim i As Integer
    Set wsResult = ThisWorkbook.Sheets("Planilha1")
    
    
    ' Inicializa os intervalos
    For i = 1 To QtdeIntervalo ''Depende da quantidade de intervalos
        Intervalo(i) = 0
    Next i
    
    'Neste algoritmo somente as planilhas que iniciam com "PDD" tem a verificação pela célula E13 (If) enquanto os demais casos são verificados pela C13 (Else)
    ' Percorre cada planilha na Pasta de Trabalho
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "Adicional") > 0 Then   ''Depende da planilha "PDC"
            For i = 1 To QtdeIntervalo   ''Depende da quantidade de intervalos
                If ws.Range("D18").Value >= kmInicial + (i - 1) * Segmento And ws.Range("D18").Value < kmInicial + i * Segmento Then ''Depende da célula
                Intervalo(i) = Intervalo(i) + ws.Range("K102").Value    ''Depende da célula
                End If
            Next i
        Else   ''Depende da planilha "PDC"
            For i = 1 To QtdeIntervalo   ''Depende da quantidade de intervalos
                If ws.Range("D18").Value >= kmInicial + (i - 1) * Segmento And ws.Range("D18").Value < kmInicial + i * Segmento Then ''Depende da célula e do kminicial
                    Intervalo(i) = Intervalo(i) + ws.Range("K98").Value + ws.Range("K100").Value  ''Depende da célula
                End If
            Next i
        End If
    Next ws

    ' Exibe os resultados
    For i = 1 To QtdeIntervalo   ''Depende da quantidade de intervalos
        wsResult.Range("C" & (5 + i)).Value = Intervalo(i)   'Depende do número de intervalos + 7
    Next i

End Sub

