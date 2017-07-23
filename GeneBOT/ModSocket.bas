Attribute VB_Name = "ModSocket"
Public Sub send(str As String)
If frmMain.sck.State = 7 Then
    frmMain.sck.SendData UTF8_Encode(str & vbCrLf)
    push "COMMAND SENT:" & str
End If
End Sub

Public Sub say(ByVal str As String)
send "PRIVMSG #jasonfamily :" & str
End Sub

Public Function getWebContent(browser As WebBrowser) As String
    Dim doc As Object
    Dim i As Object
    Dim strHtml As String
    
    Set doc = browser.Document
    For Each i In doc.All
        strHtml = strHtml & Chr(13) & i.innerText
    Next
    getWebContent = strHtml
End Function

