Attribute VB_Name = "Module2"
Sub loop_all_worksheets()

Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    'loop module 1
    
    Call yearlystocks
    
Next ws
End Sub
