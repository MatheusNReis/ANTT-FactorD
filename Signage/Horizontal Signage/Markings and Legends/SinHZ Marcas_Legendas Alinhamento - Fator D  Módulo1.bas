Sub SinHZ_CopiarPlanilha1()


'SinHZ Marcas_Legendas Alinhamento - Fator D.xlsm

'Extract and organize markings and legends data of horizontal signage into a new spreadsheet

'Extrair e organizar dados de marcas e legendas de sinalização horizontal em nova planilha

'Created by Matheus Nunes Reis on 27/10/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-FactorD/a9683ff1e55f7a31808dee2140e0e29ed33423de/LICENSE.md
'MIT License. Copyright © 2024 MatheusNReis


    Dim works As Worksheet
    Dim NomePlanilha As String
    Dim LastRowPlanWorks As Long
    Dim linhaPlanCompilado As Long
    
    NomePlanilha = ThisWorkbook.Sheets("Informações").Cells(2, "C").Value
    TituloColunaChave = ThisWorkbook.Sheets("Informações").Cells(3, "C").Value 'Ex: Latitude
    
    linhaPlanCompilado = ThisWorkbook.Sheets("Compilado").Cells(Rows.Count, "A").End(xlUp).Row + 1 'Para iniciar na 1ª linha em branco
    
    
    If NomePlanilha = "" Then
        MsgBox "Informação 'Nome Planilha' não está preenchida."
        Exit Sub
    ElseIf TituloColunaChave = "" Then
        MsgBox "Informação 'Titulo Coluna Chave' não está preenchida."
        Exit Sub
    End If
    
    
    Dim found As Boolean
    found = False
    For Each wb In Workbooks
        For Each ws In wb.Worksheets
            If ws.Name = NomePlanilha Then
            
                Dim resposta As VbMsgBoxResult
                resposta = MsgBox("'" & NomePlanilha & "' encontrado na planilha '" & wb.Name & "'", vbOKCancel + vbQuestion, "Confirmação de Planilha")
                If resposta = vbCancel Then
                    Exit Sub
                End If
                
                Set workb = wb
                Set works = wb.Sheets(NomePlanilha) 'works é a planilha origem dos dados
                found = True
                Exit For
                
            End If
        Next ws
        If found Then Exit For
    Next wb
        
    If Not found Then
        MsgBox "Planilha '" & NomePlanilha & "' não encontrada nas planilhas abertas."
        Exit Sub
    End If
    
    
    Dim Identificacao As String, km As String, Latitude As String, Longitude As String, PeliculaTipo As String
    Dim Cor As String, MediaRetrorrefletancia As String, Conc_Sup As String, Ano As String
    Dim MinimaRetrorrefletancia As Integer
    km = ThisWorkbook.Sheets("Informações").Cells(6, "B").Value
    Latitude = ThisWorkbook.Sheets("Informações").Cells(6, "C").Value
    Longitude = ThisWorkbook.Sheets("Informações").Cells(6, "D").Value
    TipoSinalizacao = ThisWorkbook.Sheets("Informações").Cells(6, "E").Value
    Cor = ThisWorkbook.Sheets("Informações").Cells(6, "F").Value
    MediaRetrorrefletancia = ThisWorkbook.Sheets("Informações").Cells(6, "G").Value
    MinimaRetrorrefletancia = ThisWorkbook.Sheets("Informações").Cells(6, "H").Value
    Conc_Sup = ThisWorkbook.Sheets("Informações").Cells(6, "I").Value
    Ano = Format(ThisWorkbook.Sheets("Informações").Cells(6, "J").Value, 0#)
    
    
    If km = "" Then
        MsgBox "Informação da coluna 'km' não está preenchida."
        Exit Sub
    ElseIf Latitude = "" Then
        MsgBox "Informação da coluna 'Latitude' não está preenchida."
        Exit Sub
    ElseIf Longitude = "" Then
        MsgBox "Informação da coluna 'Longitude' não está preenchida."
        Exit Sub
    ElseIf TipoSinalizacao = "" Then
        MsgBox "Informação da coluna 'Tipo Sinalização' não está preenchida."
        Exit Sub
    ElseIf Cor = "" Then
        MsgBox "Informação da coluna 'Cor' não está preenchida."
        Exit Sub
    ElseIf MediaRetrorrefletancia = "" Then
        MsgBox "Informação da coluna 'Média Retrorrefletância' não está preenchida."
        Exit Sub
    ElseIf MinimaRetrorrefletancia = 0 Then
        MsgBox "Informação da coluna 'Mínima Retrorrefletância' não está preenchida."
        Exit Sub
    ElseIf Conc_Sup = "" Then
        MsgBox "Informação da coluna 'Concessionária/Supervisora' não está preenchida."
        Exit Sub
    ElseIf Ano = "" Then
        MsgBox "Informação da coluna 'Ano' não está preenchida."
        Exit Sub
    End If
    
    
    
    'Inicialização
    Dim i As Long
    i = 1 'i é linha na planilha works
      
    Do While (InStr(1, works.Cells(i, km).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) = 0)
        i = i + 1
    Loop
    
    Do While (InStr(1, works.Cells(i, km).MergeArea.Cells(1, 1).Value, TituloColunaChave, vbTextCompare) > 0)
        i = i + 1
    Loop
    
    LastRowPlanWorks = works.Cells(Rows.Count, km).End(xlUp).Row + 4 ' +4 para garantir que todas as linha sejam consideradas - compensar a mescla
    
    For i = i To LastRowPlanWorks
    
        ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "A").Value = workb.Name
        ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "B").Value = works.Cells(i, km).MergeArea.Cells(1, 1).Value
        ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "C").Value = CDbl(works.Cells(i, Latitude).MergeArea.Cells(1, 1).Value)
        ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "D").Value = CDbl(works.Cells(i, Longitude).MergeArea.Cells(1, 1).Value)
        ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "E").Value = works.Cells(i, TipoSinalizacao).MergeArea.Cells(1, 1).Value
        ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "F").Value = works.Cells(i, Cor).MergeArea.Cells(1, 1).Value
        ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "G").Value = CDbl(works.Cells(i, MediaRetrorrefletancia).MergeArea.Cells(1, 1).Value)
        If CDbl(works.Cells(i, MediaRetrorrefletancia).MergeArea.Cells(1, 1).Value) = 0 Then 'Condição aplicada por conta das linhas adicionais p/ compensar mescla que resultam numa MediaRetrorrefletancia = 0 que deve corresponder a uma MinimaRetrorrefletancia = 0 também para que não entre na lista de não-atendimentos
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "H").Value = 0
        Else
            ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "H").Value = MinimaRetrorrefletancia
        End If
        ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "I").Value = Conc_Sup
        ThisWorkbook.Sheets("Compilado").Cells(linhaPlanCompilado, "J").Value = Ano
        linhaPlanCompilado = linhaPlanCompilado + 1
          
    Next i
    
    MsgBox "Fim do Processo."

End Sub
