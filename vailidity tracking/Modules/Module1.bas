Attribute VB_Name = "Module1"
Sub Tokef()
'
' Macro1 מאקרו
'1
'

Sht1 = Sheets(1).Name
sht2 = "sheet222"



 Sheets("template2").Select
    Sheets("template2").Copy After:=Sheets(3)
    Sheets("template2 (2)").Select
    Sheets("template2 (2)").Name = "sheet222"
    
    
    
Sheets(Sht1).Activate

Call Macro1

Sheets(Sht1).Activate




Dim Takin(1 To 2000) As Integer

Dim arr(1 To 2000) As Check

Dim mis(1 To 2000) As Check

reshet = Range("b2").Value


Dim a As New Check
k = 1
L = 1

For i = 3 To 23
For j = 4 To 90


If (IsDate(Cells(j, i).Value)) Then
Set arr(k) = New Check
With arr(k)
.cd = Cells(j, i).Value
.val = Cells(4, i).Value
.mosad = Cells(j, 2).Value
.Seif = Cells(5, i).Value
End With

k = k + 1
End If

If ((Cells(j, i).Value = "חסר") Or (Cells(j, i).Value = "לא תקין")) Then

If (Cells(j, i).Value = "לא תקין") Then Takin(L) = 1

Set mis(L) = New Check
With mis(L)
.val = Cells(4, i).Value
.mosad = Cells(j, 2).Value
.Seif = Cells(5, i).Value
End With
L = L + 1
End If




Next
Next


'Dim a As New Check
'With a
'.cd = Range("B12").Value
'.val = 0
'End With
'
'Dim b As New Check
'With b
''.cd = Range("C12").Value
'.val = 0
'End With



ok = False

For i = 1 To k - 1

For j = 2 To k - 1

If (Early(arr(j), arr(j - 1))) Then

Set tmp = arr(j - 1)
Set arr(j - 1) = arr(j)
Set arr(j) = tmp
End If
Next
Next



Sheets(sht2).Activate





Cells(5 + L, 3).Activate

For i = 1 To k - 1
arr(i).toString


If Not (soon(Date, arr(i).dl)) Then
    With ActiveCell.Font
      .Color = -16776961
        .TintAndShade = 0
    End With
     With ActiveCell.Offset(0, 1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
     With ActiveCell.Offset(0, 2).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
     With ActiveCell.Offset(0, 3).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    
    
    End If
   




ActiveCell.Offset(1, 0).Activate
Next



Range("C6").Activate

For i = 1 To L - 1



    
With ActiveCell.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    With ActiveCell.Offset(0, 1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
     With ActiveCell.Offset(0, 2).Font
        .Color = -16776961
        .TintAndShade = 0
    End With

ActiveCell.Value = CStr(mis(i).mosad)
ActiveCell.Offset(0, 1).Value = CStr(mis(i).Seif)

If (Takin(i) = 1) Then
ActiveCell.Offset(0, 2).Value = "לא תקין"
Else: ActiveCell.Offset(0, 2).Value = "חסר"
End If


ActiveCell.Offset(1, 0).Activate
Next





Range("D4").Value = Date

 Sheets(Array(sht2, Sheets(2).Name)).Move
 
 
   
ActiveWorkbook.SaveAs Filename:=" תוקף " & reshet & " " & Day(Date) & "." & Month(Date) & "." & Year(Date)


Sheets(2).Select
Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$5"
        .PrintTitleColumns = ""
    End With
    
    
Sheets(1).Move
Range("b3").Value = "סיום תוקף"
Range("c3").Value = "נכון לתאריך"
Range("d3").Value = Date

ActiveWorkbook.SaveAs Filename:="סיום תוקף " & reshet & " " & Day(Date) & "." & Month(Date) & "." & Year(Date)



' ActiveWorkbook.SaveAs Filename:=mosad + " " + Date




















 
          
End Sub


Function Early(c1 As Check, c2 As Check) As Boolean
'TRUE if c1<c2. FALSE otherwise


If (Year(c1.dl) < Year(c2.dl)) Then
Early = True
Exit Function
End If

If (Year(c1.dl) > Year(c2.dl)) Then
Early = False
Exit Function
End If



If (Month(c1.dl) < Month(c2.dl)) Then
Early = True
Exit Function
End If
If (Month(c1.dl) > Month(c2.dl)) Then
Early = False
Exit Function
End If

If (Day(c1.dl) < Day(c2.dl)) Then
Early = True
Exit Function
End If

If (Day(c1.dl) > Day(c2.dl)) Then
Early = False
Exit Function
End If

Early = False




End Function

Function soon(c1 As Date, c2 As Date) As Boolean
'TRUE if c1<c2. FALSE otherwise


If (Year(c1) < Year(c2)) Then
soon = True
Exit Function
End If

If (Year(c1) > Year(c2)) Then
soon = False
Exit Function
End If



If (Month(c1) < Month(c2)) Then
soon = True
Exit Function
End If
If (Month(c1) > Month(c2)) Then
soon = False
Exit Function
End If

If (Day(c1) < Day(c2)) Then
soon = True
Exit Function
End If

If (Day(c1) > Day(c2)) Then
soon = False
Exit Function
End If

soon = False

End Function








Sub Macro1()
'
' Macro1 מאקרו
'

'
    Sheets(1).Select
    
    If (Sheets(2).Name = "סיום תוקף") Then Sheets(2).Delete
    
    Sheets(1).Copy Before:=Sheets(2)
    ActiveSheet.Shapes.Range(Array("Button 1")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("Button 2")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("Button 3")).Select
    Selection.Delete
    
    
   
    Sheets(2).Name = "סיום תוקף"
    
    
    
    For i = 3 To 23
    For j = 4 To 90
    
    If (IsDate(Cells(j, i).Value)) Then
    
    Set a = New Check
    With a
    .cd = Cells(j, i).Value
    .val = Cells(4, i).Value
    End With

    Cells(j, i).Value = a.dl
    
    End If
    
    
    If (IsDate(Cells(j, i).Value)) Then
    If (soon(a.dl, Date)) Then
    With Cells(j, i).Font
      .Color = -16776961
        .TintAndShade = 0
    End With
    End If
    End If
    
    
    
    If Not (IsDate(Cells(j, i).Value)) Then
    
    
    If ((Cells(j, i).Value = "חסר") Or (Cells(j, i).Value = "לא תקין")) Then
    With Cells(j, i).Font
      .Color = -16776961
        .TintAndShade = 0
    End With
    
    End If
    End If
    
    Range("b3").Value = "נוצר בתאריך:"
    Range("c3").Value = Date
    
    Next
    Next

    
    
    
    
    
    
End Sub
