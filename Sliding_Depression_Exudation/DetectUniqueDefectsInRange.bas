Sub DetectUniqueDefectsInRange()


'Copy data from pavement defect range to column A of worksheet Planilha1 to determine unique values from
'monitoring sheets

'Copia dados contidos no intervalo de defeitos do pavimento para a coluna A da Planilha1 para definição de
'valores únicos das fichas de monitoração


    Dim wss As Worksheet
    Dim ws As Worksheet
    Dim destSheet As Worksheet
    Dim destRow As Long
    Dim sourceRange As Range
    Dim cell As Range
    
    'Range of pavement defect in each visual monitoring sheet
    PavementDefectsRange = "F38:F116"
    
    
    'Add a new worksheet called Planilha1
    Set destSheet = ThisWorkbook.Worksheets.Add
    destSheet.Name = "Planilha1"
    
    'Initialize the destination row
    destRow = 1
    
    'Loop through all worksheets except Planilha1
    For Each ws In ThisWorkbook.Worksheets
    
        If ws.Name <> "Planilha1" Then
            
            Set sourceRange = ws.Range(PavementDefectsRange)
            
            'Copy data from source range to Planilha1
            For Each cell In sourceRange
            
                destSheet.Cells(destRow, 1).Value = cell.Value
                destRow = destRow + 1
                
            Next cell
            
        End If
        
    Next ws
    
    
    
    Set wss = ThisWorkbook.Worksheets("Planilha1")
    
    LastRowColumnA = wss.Cells(wss.Rows.Count, "A").End(xlUp).Row
    
    wss.Range("B1").FormulaLocal = "=ÚNICO(A1:A" & LastRowColumnA & ")"
   
    wss.Range("C1").FormulaLocal = "Apague o '@' da fórmula de B1, caso haja."
   
   
    
    MsgBox "Fim do processo."
    
    
End Sub

