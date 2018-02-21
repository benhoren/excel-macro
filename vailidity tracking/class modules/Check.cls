VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CLASS MODULE CODE

' Member variable
Public cd As Date   'current date
Public val As Double   'validity
Public mosad As String
Public Seif As String

Function Check()

End Function
Function toString()
'cd = Format(cd, "dd/mm/yyyy")
ActiveCell.Value = CStr(mosad)
ActiveCell.Offset(0, 1).Value = CStr(Seif)
ActiveCell.Offset(0, 2).Value = cd
ActiveCell.Offset(0, 3).Value = dl

End Function


Function dl() As Date
Dim s As Date
If (val = CInt(val)) Then s = DateAdd("yyyy", val, cd)

If Not (val = CInt(val)) Then s = DateAdd("m", val * 12, cd)
dl = s

End Function





