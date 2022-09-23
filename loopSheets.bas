Attribute VB_Name = "Module2"
Sub loopSheets():
    Dim ws As Worksheet
    Dim ws_count As Integer
    
    'set worksheet count to 1
    ws_count = 1
        
    'iterate through sheets
    For Each ws In Worksheets
    
        'call function to summarize stocks
        summarizeStocks
        'call function to find maximum percentage changes and volumes
        findMaxAmts
        
        'https://www.automateexcel.com/vba/activate-select-sheet/
        'if activesheet index number does not equal total number of sheets
        If ActiveSheet.Index <> Worksheets.Count Then
            
            'proceed to next sheet
            ActiveSheet.Next.Activate
                
        Else
            'return to first sheet
            Worksheets(1).Activate
        End If
        
    Next ws
End Sub

