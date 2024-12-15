Sub ProcurarValor_e_Armazena_em_CálculoIDPAV_AmbosSentidos()


'IRICalculus.xlsm

'Prepare a spreadsheet for IRI calculation, finding the corresponding value for each kilometer and storing it.
'It works for both directions of the highway.

'Prepara planilha para cálculo do IRI, encontrando valor correspondente para cada quilometragem e o armazenando.
'Funciona para ambas as direções de rodovia.

'Created by Matheus Nunes Reis on 22/07/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-FactorD/677e3b845d12e558a681e9e7b02176cb840fa511/LICENSE.md
'MIT License. Copyright © 2024 MatheusNReis


'Ambas as planilhas devem estar abertas - a de CálculoIDPAV e a de dados bruto IRI

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim cell2 As Range

    ' Define as planilhas
    Set ws1 = Workbooks("Cálculo IDPAV MSVIA").Sheets("Planilha1") 'Planilha para CálculoIDPAV
    Set ws2 = Workbooks("MSV-163MS-104-830-MON-OUT-RM-Z9-013-R00.xlsx").Sheets("IRI SF2") '''Defina workbook e sheet da Planilha com dados brutos de IRI

    ' Define o intervalo como a coluna A na Planilha1
    Set rng = ws1.Range("A3:A" & ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row)

    ' Loop através de cada célula na coluna A da Planilha1
    For Each cell In rng
        ' Loop através de cada célula na coluna B da IRI NF1
        For Each cell2 In ws2.Range("A1:A" & ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row)
            offsetCell = 5   'cell.offset(0, 1) -> 1 que representa offset a partir da coluna A até coluna de atende/não atende (CálculoIDPAV)
            offsetCell2 = 14 'cell2.offset(0, 14) -> 14 que representa offset a partir da coluna A até coluna de atende/não atende (Dados IRI)
            ' Se o valor for encontrado
            If cell.Value = cell2.Value Then
                ' Captura o valor correspondente de atende/não atende da planilha de dados brutos IRI
                ' Se a célula estiver mesclada, captura o valor da primeira célula da mescla
                If cell2.Offset(0, x).MergeCells Then
                    cell.Offset(0, offsetCell).Value = cell2.Offset(0, offsetCell2).MergeArea(1, 1).Value
                Else
                    cell.Offset(0, offsetCell).Value = cell2.Offset(0, offsetCell2).Value
                End If
                ' Sai do loop For Each uma vez que a correspondência é encontrada
                Exit For
            Else
                
            End If
        Next cell2
    Next cell
    MsgBox "Processo finalizado!"

End Sub
