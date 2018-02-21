Attribute VB_Name = "Module3"

Sub Macro10()
Attribute Macro10.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro10 מאקרו
'

'
    Sheets(1).Copy
   ' ActiveSheet.Move
   
   ActiveSheet.Shapes.Range(Array("Button 2")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("Button 1")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("Button 3")).Select
    Selection.Delete
    Dim a As String
    
    a = Range("b2").Value
    
    Range("c3").Value = "נכון לתאריך"
    Range("d3").Value = Date
    
    
    ActiveWorkbook.SaveAs Filename:="תוקף " & a & " " & Day(Date) & "." & Month(Date) & "." & Year(Date)
    
    
    
    
End Sub

