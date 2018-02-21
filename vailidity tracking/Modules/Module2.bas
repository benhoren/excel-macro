Attribute VB_Name = "Module2"
Sub TokefMosad()
'
' Macro1 מאקרו
'1
'

Sht1 = Sheets(1).Name
sht2 = "sheet2"
sht3 = "sheet3"

mosad = Selection.Value

Dim arr(1 To 100) As Check

Dim mis(1 To 100) As Check


'Dim v As String
k = 1
L = 1
r = ActiveCell.Row

For i = 3 To 23

If (IsDate(Cells(r, i).Value)) Then
Set arr(k) = New Check
With arr(k)
.cd = Cells(r, i).Value
.val = Cells(4, i).Value
.Seif = Cells(5, i).Value
End With

k = k + 1
End If

Dim Takin(1 To 100) As Integer

If ((Cells(r, i).Value = "חסר") Or (Cells(r, i).Value = "לא תקין")) Then
If (Cells(r, i).Value = "לא תקין") Then Takin(L) = 1
Set mis(L) = New Check
With mis(L)
.val = Cells(4, i).Value
.Seif = Cells(5, i).Value
End With
L = L + 1
End If


Next

For i = 1 To k - 1

For j = 2 To k - 1

If (Early(arr(j), arr(j - 1))) Then

Set tmp = arr(j - 1)
Set arr(j - 1) = arr(j)
Set arr(j) = tmp
End If

Next
Next



'Sheets(sht3).Activate


'Range("E8:H31").Select
'    With Selection.Font
'        .ThemeColor = xlThemeColorLight1
'        .TintAndShade = 0
'    End With
'    Selection.ClearContents
    
    Sheets("template1").Select
    Sheets("template1").Copy After:=Sheets(3)
    Sheets("template1 (2)").Select
    Sheets("template1 (2)").Name = "sheet3"



Cells(7 + L, 5).Activate
For i = 1 To k - 1


ActiveCell.Value = CStr(arr(i).Seif)
ActiveCell.Offset(0, 1).Value = arr(i).val
ActiveCell.Offset(0, 2).Value = arr(i).cd
ActiveCell.Offset(0, 3).Value = arr(i).dl


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


Range("E6").Value = mosad

Range("E8").Activate





For i = 1 To L - 1


With ActiveCell.Font
        .Color = -16776961
        .TintAndShade = 0
    End With

ActiveCell.Value = CStr(mis(i).Seif)
ActiveCell.Offset(0, 1).Value = mis(i).val
ActiveCell.Offset(0, 1).Font.Color = -16776961
If (Takin(i) = 1) Then
ActiveCell.Offset(0, 2).Value = "לא תקין"
Else: ActiveCell.Offset(0, 2).Value = "חסר"
End If
ActiveCell.Offset(0, 2).Font.Color = -16776961
ActiveCell.Offset(1, 0).Activate
Next




Range("E4").Value = Date

 
    ActiveSheet.Move


Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$7"
        .PrintTitleColumns = ""
    End With

'Range("a1").Value = cda
 ActiveWorkbook.SaveAs Filename:="תוקף" & " " & mosad & " " & Day(Date) & "." & Month(Date) & "." & Year(Date)



 
          
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




