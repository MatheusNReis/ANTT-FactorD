Sub CriaColunakm_SentidoDecrescente()


'IRICalculus.xlsm

'Create a kilometer column for the descending direction of the highway

'Cria coluna de quilometragem considerando a rodovia no sentido decrescente


'Created by Matheus Nunes Reis on 18/07/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-FactorD/677e3b845d12e558a681e9e7b02176cb840fa511/LICENSE.md
'MIT License. Copyright © 2024 MatheusNReis


Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    ' Define a planilha ativa
    Set ws = Worksheets("IRI SF3") ''''Inserir nome da planilha contida na pasta de trabalho, que deve estar aberta já

    ' Insere uma nova coluna antes da coluna A
    ws.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    ' Define o intervalo como a coluna C
    Set rng = ws.Range("C5:C" & ws.Cells(ws.Rows.Count, "C").End(xlUp).Row) '''LINHA E (COLUNA +1) da Célula que inicia os dados, excluindo cabeçalhos
                                                                                            'o +1 no comentário acima é por conta da nova coluna que será criada em na coluna A
    ' Loop através de cada célula na coluna C
    For Each cell In rng
        ' Trunca o valor e coloca na coluna A
        cell.Offset(0, -2).Value = Int(cell.Value)
    Next cell

End Sub
