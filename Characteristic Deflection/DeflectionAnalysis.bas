Sub AnaliseFLECHAS()
    
    
'List segments of the highway that have a value above the pre-established characteristic deflection limit

'Lista trechos da rodovia que possuem valor acima do limite pré-estabelecido de deflexão característica

'Created by Matheus Nunes Reis on 17/07/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-FactorD/9066fa5ce0ea2d9c0dd456a4dc22dadb14aa15e9/LICENSE.md
'MIT License. Copyright © 2024 MatheusNReis
    
    
    Dim ws As Worksheet, w1 As Worksheet
    Dim listaWs As Worksheet
    Dim valor As Variant
    Dim i As Long
    
    Set listaWs = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ' Definir a planilha "LISTA" onde os nomes das planilhas serão salvos
    listaWs.Name = "FLE"
    ' Inserir cabeçalho na planilha LISTA
    listaWs.Range("A1").Value = "Lista"

    ' Inicializar o número da linha para a coluna A da planilha "LISTA" desconsiderando cabeçalho
    i = 2
    
    ' Percorrer todas as planilhas na pasta de trabalho atual
    For Each ws In ThisWorkbook.Sheets
        ' Verificar se é a planilha "LISTA" ou outra planilha que não deverá ser verificada
        If ws.Name <> listaWs.Name Then
            ' Verificar se há algum valor > 10 no intervalo C12:E27 '''''''''''''''''Intervalo avaliado e limite de 10 mm de flecha
            For Each valor In ws.Range("D117:K127").Value
                If IsNumeric(valor) And valor > 10 Then
                    ' Se encontrar valor > 10, salvar o nome da planilha na coluna A da planilha "FLE"
                    listaWs.Cells(i, 1).Value = ws.Name
                    i = i + 1 ' Avançar para a próxima linha na coluna A da planilha "FLE"
                    Exit For ' Sair do loop ao encontrar o primeiro valor > 10
                End If
            Next valor
        End If
    Next ws
    
    ' Limpar conteúdo das células não utilizadas na coluna A da planilha "LISTA"
    listaWs.Range("A" & i & ":A" & listaWs.Rows.Count).ClearContents
    
    MsgBox "Processo concluído. Nomes das planilhas com valores > 10 foram registrados na planilha 'FLE'."
End Sub



