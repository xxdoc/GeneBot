Attribute VB_Name = "ModGeneral"
Public Const uname As String = "billyX0807"
Public Const user As String = "billyX0807"
Public Const address As String = "chat.freenode.net"
Public Const port = "6667"
Public Const channel = "#jasonfamily"


Public Sub push(payload As String)
    frmMain.List.Text = frmMain.List.Text & vbCrLf & Time & ":" & payload & "."
    frmMain.List1.AddItem Time & ":" & payload & "."
End Sub

Public Function Rndz(min As Integer, max As Integer)
Rndz = Int(Rnd * (max - min + 1)) + min
End Function

