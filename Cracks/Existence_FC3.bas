Sub ExistenciaFC3()


'List the monitoring records for highway segments that have FC3 type cracks

'Lista as fichas de monitoração dos trechos de rodovia que possuem trincas do tipo FC3

'Created by Matheus Nunes Reis on 18/07/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-FactorD/f94aa23bda5eecc9f69e057f0860204d0f84459a/LICENSE.md
'MIT License. Copyright © 2024 MatheusNReis


    Dim ws As Worksheet, w1 As Worksheet
    Dim listaWs As Worksheet
    Dim valor As Variant
    Dim i As Long
    
    Set listaWs = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ' Definir a planilha "LISTA" onde os nomes das planilhas serão salvos
    listaWs.Name = "ExistênciaFC3"
    ' Inserir cabeçalho na planilha LISTA
    listaWs.Range("G1").Value = "Existência de FC3(km)"

    ' Inicializar o número da linha para a coluna A da planilha "LISTA" desconsiderando cabeçalho
    i = 2
    
    ' Percorrer todas as planilhas na pasta de trabalho atual
    For Each ws In ThisWorkbook.Sheets
        ' Verificar se é a planilha "LISTA" ou outra planilha que você não quer verificar
        If ws.Name <> listaWs.Name Then
            ' Verificar se há algum valor > 10 no intervalo C12:E27 '''''''''''''''''Intervalo avaliado e limite de 10 mm de flecha
            For Each valor In ws.Range("F42:F96").Value
                If valor = "FC-3" Then
                    ' Se encontrar valor "FC-3", salvar o nome da planilha na coluna G da planilha "ExistênciaFC3"
                    listaWs.Cells(i, "G").Value = ws.Name
                    i = i + 1 ' Avançar para a próxima linha na coluna A da planilha "ExistênciaFC3"
                    Exit For
                End If
            Next valor
        End If
    Next ws
    
    ' Limpar conteúdo das células não utilizadas na coluna A da planilha "LISTA"
    listaWs.Range("G" & i & ":G" & listaWs.Rows.Count).ClearContents
    
    MsgBox "Processo concluído. Nomes das planilhas com valores ""FC-3"" foram registrados na planilha ""ExistênciaFC3""."
End Sub
