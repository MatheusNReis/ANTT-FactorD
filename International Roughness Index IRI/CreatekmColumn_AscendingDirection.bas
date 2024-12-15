Sub CriaColunakm_SentidoCrescente()

'IRICalculus.xlsm

'Create a kilometer column for the ascending direction of the highway

'Cria coluna de quilometragem considerando a rodovia no sentido crescente

'Created by Matheus Nunes Reis on 18/07/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License:
'MIT License. Copyright © 2024 MatheusNReis


Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    ' Define a planilha ativa
    Set ws = Worksheets("IRI NF3") '''''Inserir nome da planilha contida na pasta de trabalho, que deve estar aberta já

    ' Insere uma nova coluna antes Da coluna A
    ws.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    ' Define o intervalo como a coluna A
    Set rng = ws.Range("B5:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row) ''''LINHA E (COLUNA +1) da Célula que inicia os dados, excluindo cabeçalhos
                                                                                            'o +1 no comentário acima é por conta da nova coluna que será criada em na coluna A
    ' Loop através de cada célula na coluna B
    For Each cell In rng
        ' Trunca o valor e coloca na coluna A
        cell.Offset(0, -1).Value = Int(cell.Value)
    Next cell

End Sub
