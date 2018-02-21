Attribute VB_Name = "Module1"

Sub MakeReport()
Attribute MakeReport.VB_Description = "הפונ' יוצרת דוח מהשורה הלחוצה"
Attribute MakeReport.VB_ProcData.VB_Invoke_Func = "D\n14"
'
' Report
'

'
    sht1 = ActiveSheet.Name
    sht2 = "sheet2"
    
    
    
    exists = False
    For i = 1 To Worksheets.Count
    If (Worksheets(i).Name = sht2) Then exists = True
    
    Next i
    
    If (exists = False) Then Sheets.Add.Name = sht2
    
    
    
    
    
    Selection.Copy
    Set c = ActiveCell
    Sheets(sht2).Activate
    Range("E7").Activate
    ActiveSheet.Paste
    
    Sheets(sht1).Activate
    ActiveCell.Offset(0, 1).Range("A1").Select
    parasha = ActiveCell.Value
    
    Selection.Copy
    Sheets(sht2).Activate
    Range("D11").Activate
    ActiveSheet.Paste
    
    
    Sheets(sht1).Activate
    ActiveCell.Offset(0, 1).Range("A1").Activate
  '  Application.CutCopyMode = False
    Dim d As Date
    d = CDate(Selection.Value)
    Sheets(sht2).Activate
    Range("F11").Activate
    Range("F11").Value = d
   ' ActiveSheet.Paste
    
    Sheets(sht1).Activate
    ActiveCell.Offset(0, 1).Range("A1").Activate
  '  Application.CutCopyMode = False
  lastName = ActiveCell.Value
    Selection.Copy
    Sheets(sht2).Activate
    ActiveCell.Offset(-2, -2).Range("A1").Activate
    ActiveSheet.Paste
    
    
    Sheets(sht1).Select
    ActiveCell.Offset(0, 2).Range("A1:C1").Select
   ' Application.CutCopyMode = False
    Selection.Copy
    Sheets(sht2).Select
    Range("D16:F16").Select
    ActiveSheet.Paste
    
    For i = 1 To 9
    Sheets(sht1).Select
    ActiveCell.Offset(0, 3).Range("A1:C1").Select
  '  Application.CutCopyMode = False
    Selection.Copy
    Sheets(sht2).Select
    ActiveCell.Offset(1, 0).Range("A1:C1").Select
    ActiveSheet.Paste
    
    Next
   
   
    Sheets(sht1).Activate
    Range("A1").Select
    Selection.Copy
     ActiveSheet.Paste
    
   
    Sheets(sht2).Select
    Range("A1").Select
    
    
    
    
    
     Range("G16").Activate
    For i = 1 To 11
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 4.99893185216834E-02
    End With
    ActiveCell.Offset(1, 0).Activate
    Next
    
    
    
    
    
    
    
    
    
    
    
    
    
    Range("C6").Select
    Selection.Font.Bold = True
    Range("C6").Select
    ActiveCell.FormulaR1C1 = "ב""ה"
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "מס-"
    Range("E7").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("D7").Select
    Selection.Font.Bold = True
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "מס' חשבון"
    Range("C9").Select
    ActiveCell.FormulaR1C1 = "שם:"
    Range("D9:G9").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("C9").Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("E7").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("C11").Select
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("C11").Select
    ActiveCell.FormulaR1C1 = "פרשה:"
    Range("D11").Select
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("E11").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Font.Bold = True
    Range("E11").Select
    ActiveCell.FormulaR1C1 = "תאריך"
    Range("C11").Select
    ActiveCell.FormulaR1C1 = "פרשה"
    Range("F11:G11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("C13:G13").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Font.Bold = True
    Range("C13:G13").Select
    ActiveCell.FormulaR1C1 = "חשבון סופי"
    Range("C6:G13").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 16
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Columns("C:C").ColumnWidth = 7.5
    Columns("B:B").ColumnWidth = 5.63
    Range("E4").Select
    Columns("E:E").ColumnWidth = 11.88
    Columns("D:D").ColumnWidth = 9.88
    Columns("D:D").ColumnWidth = 11.25
    Columns("E:E").ColumnWidth = 12.25
   
    Range("C13:G13").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("C15").Select
    ActiveCell.FormulaR1C1 = "מס""ד"
    Range("D15").Select
    ActiveCell.FormulaR1C1 = "פריט"
    Range("E15").Select
    ActiveCell.FormulaR1C1 = "יח'"
    Range("F15").Select
    ActiveCell.FormulaR1C1 = "מחיר ליח'"
    Range("G15").Select
    ActiveCell.FormulaR1C1 = "סה""כ"
    Range("C16").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("C17").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("C18").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("C19").Select
    ActiveCell.FormulaR1C1 = "4"
    Range("C20").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("C21").Select
    ActiveCell.FormulaR1C1 = "6"
    Range("C22").Select
    ActiveCell.FormulaR1C1 = "7"
    Range("C23").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("C24").Select
    ActiveCell.FormulaR1C1 = "9"
    Range("C25").Select
    ActiveCell.FormulaR1C1 = "10"
    Range("C15:G15").Select
    Selection.Font.Bold = True
    Range("C15:G25").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("G26").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    
    
    
    
   Range("G16").Activate
    
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-2]"
    
    For i = 1 To 9
     ActiveCell.Offset(1, 0).Activate
     ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-2]"
     
     Next
   
   
    Range("G26").Activate
    ActiveCell.FormulaR1C1 = "=SUM(R[-10]C:R[-1]C)"
     
    Columns("E:E").ColumnWidth = 11.25
    
     Range("G26").Select
    With Selection
        .HorizontalAlignment = xlCenter
    
    End With
    
    

    
    
    
    
    Rows("15:25").RowHeight = 17
    
    
   
    Range("G16").Activate
    For i = 1 To 9
    
    If (ActiveCell.Value < 0) Then
     With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    End If
    ActiveCell.Offset(1, 0).Activate
    Next
    
     Range("G26").Activate
      If (ActiveCell.Value < 0) Then 'minos = False
     With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    End If
    
    
    
    Range("E7").Activate
    Selection.Font.Bold = True
    Range("E9").Activate
    Selection.Font.Bold = True
    
    
    
    
    Range("A1").Select
    Selection.Copy
     ActiveSheet.Paste
     
    
    Sheets(sht2).Select
  
  
  Range("C11").Select
    Selection.Font.Bold = False
    Range("E11").Select
    Selection.Font.Bold = False
    Range("D11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
  
  a = Range("G26").Value
  Sheets(sht1).Select
  c.Activate
  ActiveCell.Offset(0, 4).Activate
  
  ActiveCell.Value = a
  
  
  
  
  
  
  
    
    
        ActiveWorkbook.Sheets(sht2).Copy _
           After:=ActiveWorkbook.Sheets(sht2)
    ActiveSheet.Move
  
  
  
 ' FileOnly = ThisWorkbook.Name
  
 ' Range("A1").Value = FileOnly
 
 ActiveWorkbook.SaveAs Filename:=parasha + " " + lastName
  
  
  
  
  
  
  
  
  
  
  
    
    
    
   
 
End Sub

