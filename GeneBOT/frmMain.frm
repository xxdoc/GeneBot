VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMain 
   Caption         =   "BillyBot V2.1 Beta"
   ClientHeight    =   7500
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   14715
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "��̨ģʽ"
      Height          =   615
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6840
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "frmMain.frx":0000
      Top             =   6720
      Width           =   615
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   6720
      Width           =   1575
      ExtentX         =   2778
      ExtentY         =   450
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   6480
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton Command4 
      Caption         =   "send command"
      Height          =   255
      Left            =   6480
      TabIndex        =   10
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "send message"
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   6000
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   120
      TabIndex        =   8
      Top             =   6000
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   3360
      TabIndex        =   7
      Top             =   6720
      Width           =   180
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   10560
      Top             =   4200
   End
   Begin VB.ListBox List1 
      Height          =   6540
      Left            =   11400
      TabIndex        =   6
      Top             =   240
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   10560
      Top             =   4680
   End
   Begin MSWinsockLib.Winsock sck 
      Left            =   10560
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "irc.freenode.net"
      RemotePort      =   6667
      LocalPort       =   6667
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ֹͣ"
      Height          =   495
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   7920
      TabIndex        =   1
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "��־"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin VB.TextBox List 
         Height          =   5415
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   240
         Width           =   10815
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TIME"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   6720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6720
      Width           =   735
   End
   Begin VB.Menu about 
      Caption         =   "����(&A).."
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Integer
Dim bt As Integer
Dim mas As String
Dim h(1 To 1000) As Integer
Dim m(1 To 1000) As Integer
Dim note(1 To 1000) As String
Dim a() As String
Dim pointer As Integer
Dim str As String
Dim fore, nex As String
Dim path As String
Dim textline, datas As String
Dim got As Boolean
Dim word As String
Const ip138 = "http://qq.ip138.com/weather/zhejiang/HangZhou.html"
Dim weatherw
Dim wsh
Dim res As VbMsgBoxResult
Dim ass
Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub Command1_Click()
'sck.Close
sck.Connect
push "Started"
End Sub

Private Sub Command2_Click()
sck.Close
push "Stopped"
End Sub

Sub shutdown()
Unload Me
End

End Sub
Private Sub Command3_Click()
say Text2.Text
Text2.Text = Empty

End Sub

Private Sub Command4_Click()
send Text2.Text
Text2.Text = Empty
End Sub

Private Sub Command5_Click()
Me.Visible = False
End Sub

Private Sub Form_Load()
Set wsh = CreateObject("WScript.shell")
n = 0
pointer = 1
Call Command1_Click

End Sub

Private Sub List_Change()
List.SelStart = Len(List.Text)
End Sub

Private Sub sck_Connect()
push "Connected@" & sck.RemoteHostIP & ":" & sck.RemotePort
send "NICK " & uname
send "USER " & user & " " & user & " " & user & " :" & user
send "JOIN " & channel

End Sub
Private Function min(a As Integer, b As Integer) As Integer
If a >= b Then
min = b
Else
min = a
End If
End Function
Private Sub sck_DataArrival(ByVal bytesTotal As Long)
Dim byData() As Byte
sck.GetData byData(), vbByte
Dim datas As String
Dim pt0 As Integer

datas = UTF8_Decode(byData())
push datas
Dim datal, diffmas As String
datal = LCase(Trim(datas))
If InStr(1, datas, " PRIVMSG #jasonfamily :") > 0 Then 'somebody said sth.
    pt0 = InStr(1, datas, " PRIVMSG #jasonfamily :")
    push "MESSAGE : " & Mid(datas, pt0 + Len("#jasonfamily :") + 9, Len(datas) - pt0 - Len("#jasonfamily :"))
    'mmas = Mid(datas, pt0 + Len("#jasonfamily :") + 9, Len(datas) - pt0 - Len("#jasonfamily :") - 10) '  what??
    'mas = Left(Trim(mmas), Min(4, Len(mmas)))
    
    diffmas = Mid(datas, pt0 + Len("#jasonfamily :") + 9, Len(datas) - pt0 - Len("#jasonfamily :"))
    diffmas = Left(diffmas, Len(diffmas) - 2)
    Text1.Text = diffmas
    diffmas = Text1.Text
    mas = Left(Trim(diffmas), min(4, Len(diffmas)))
    
    push "TRYing to intercept :" & diffmas

    If Left(diffmas, 2) = "ʱ��" Or LCase(mas) = "time" Then
    say "��ǰʱ���� " & Date & " " & Time & " ���� " & Weekday(Now, vbMonday)
    End If
    
    If LCase(diffmas) = "ping" Then say "PONG!"
        If Left(diffmas, 3) = "66" Then say "��������"
    If Left(diffmas, 3) = "666" Then say "��������"
    If Left(diffmas, 5) = "66666" Then say "��������"
    If Left(diffmas, 4) = "6666" Then say "��������"
    If Left(diffmas, 2) = "69" Then say "69���á���"
    
    If diffmas = "0807" Then
    say "Permission Granted."
    Sleep2 (0.5)
     say "The system will now halt."
    sck.Close
        shutdown
    End If

    If LCase(mas) = "reme" Then
    On Error GoTo freedom
    If Len(diffmas) < 7 Then
    say "reme�÷���reme Сʱ ���� ��Ϣ,ϵͳ������ʱ�䵽��ʱ������ϢŶ��"
    say "����: reme 12 30 �Է���!"
    say "PS:�¼����벻Ҫ�пո�Сʱ��24Сʱ��Ŷ��"
    Exit Sub
    Else
    a = Split(diffmas, " ")
    h(pointer) = a(1)
    m(pointer) = a(2)
    note(pointer) = a(3)
    say "�Ѽ�ס " & a(3) & ",���� " & a(1) & "ʱ" & a(2) & "�ִ�����ָ���Ϊ:" & pointer
    pointer = pointer + 1
    End If
    End If
freedom:
    DoEvents
    
    
    If LCase(mas) = "echo" Then
        say "�����ѽ��ܡ�"
        a = Split(diffmas, " ")
        Select Case LCase(a(1))
        Case "information"
            res = MsgBox(a(2), vbInformation)
            If res = vbYes Then say "��ѡ��Yes��"
            If res = vbNo Then say "��ѡ��No."
            If res = vbIgnore Then say "��ѡ��Ignore."
            If res = vbOK Then say "��ѡ��ok��"
            If res = vbCancel Then say "��ѡ��cancel."
            If res = vbAbort Then say "��ѡ��abort."
            If res = vbRetry Then say "��ѡ��retry."
        Case "yesno"
            res = MsgBox(a(2), vbYesNo)
            If res = vbYes Then say "��ѡ��Yes��"
            If res = vbNo Then say "��ѡ��No."
            If res = vbIgnore Then say "��ѡ��Ignore."
            If res = vbOK Then say "��ѡ��ok��"
            If res = vbCancel Then say "��ѡ��cancel."
            If res = vbAbort Then say "��ѡ��abort."
            If res = vbRetry Then say "��ѡ��retry."
        Case "yesnocancel"
            res = MsgBox(a(2), vbYesNoCancel)
            If res = vbYes Then say "��ѡ��Yes��"
            If res = vbNo Then say "��ѡ��No."
            If res = vbIgnore Then say "��ѡ��Ignore."
            If res = vbOK Then say "��ѡ��ok��"
            If res = vbCancel Then say "��ѡ��cancel."
            If res = vbAbort Then say "��ѡ��abort."
            If res = vbRetry Then say "��ѡ��retry."
        Case "warn"
            res = MsgBox(a(2), vbCritical)
            If res = vbYes Then say "��ѡ��Yes��"
            If res = vbNo Then say "��ѡ��No."
            If res = vbIgnore Then say "��ѡ��Ignore."
            If res = vbOK Then say "��ѡ��ok��"
            If res = vbCancel Then say "��ѡ��cancel."
            If res = vbAbort Then say "��ѡ��abort."
            If res = vbRetry Then say "��ѡ��retry."
        End Select
    End If
    
    If LCase(mas) = "dict" Then
        push "Dict query detected."
        a = Split(diffmas, " ")
        word = a(1)
        say "���ڲ�ѯ�У����Ժ�.."
        ass = query_dict(word)
        If ass <> Empty Then
            say word & "  :  " & ass
        Else
            say "sorry��ţ��Ӣ���ʵ�δ�鵽�ôʡ���"
        End If
    End If
    
    If diffmas = "." Then say "."
    If diffmas = ".." Then say ".."
        
    If diffmas = "..." Then say "...."
    If diffmas = "...." Then say "...."
    If diffmas = "��" Then say "��"
    If diffmas = "����" Then say "����"
    If diffmas = "������" Then say "������"
    
    If Left(diffmas, 3) = "pia" Then
    If Len(diffmas) <= 4 Then
        Randomize
        Select Case Rndz(1, 5)
        Case 1
        say "���s���ߣ����s�k�k "
        Case 2
        say "���s�F���䣩�s��ة���"
        Case 3
        say "(�s' - ')�s�� �ߩ��� "
        Case 4
        say "�Щ��� ��( ' - '��) {�ںðں�) "
        Sleep2 (0.2)
        say "(���������һ��} (�s�㧥��)�s�� �ߩ��� "
        Case 5
        say "�ߩ��ߦ�t(�F����)�s��ߩ��� �국"
        End Select
    Else
        a = Split(diffmas, " ")
        Randomize
        Select Case Rndz(1, 5)
        Case 1
        say "���s���ߣ����s�k�k " & a(1)
        Case 2
        say "���s�F���䣩�s��ة���" & a(1)
        Case 3
        say "(�s' - ')�s�� �ߩ��� " & a(1)
        Case 4
        say "�Щ��� ��( ' - '��) {�ںðں�) " & a(1)
        Sleep2 (0.2)
        say "(���������һ��} (�s�㧥��)�s�� �ߩ��� " & a(1)
        Case 5
        say "�ߩ��ߦ�t(�F����)�s��ߩ��� �국" & a(1)
        End Select
    End If
    End If
    
    If LCase(diffmas) = "weather" Or LCase(mas) = "����" Then
        say "���ڲ�ѯ����.."
        WebBrowser1.Navigate ip138
        Sleep2 (1.2)
        Text3.Text = getWebContent(WebBrowser1)
        If Len(Text3.Text) >= 10 Then
        a = Split(Text3.Text, vbCrLf)
        weatherw = Empty
        Dim i As Integer
            For i = 2 To 5
                weatherw = weatherw & a(i) & " "
            Next i
        say weatherw
            For i = 7 To 10
                weatherw = weatherw & a(i) & " "
            Next i
        say weatherw
        Else
        say "sorry��������������˵����⣡��"
        End If
        
    End If
    
    If LCase(mas) = "sudo" Then
        say "Ȩ�����."
        Dim ommand As String
        ommand = Mid(diffmas, InStr(1, diffmas, " "))
        Shell "cmd /c " & ommand, vbHide
        say "������ִ�С�"
    End If
    
    
    If LCase(mas) = "calc" Or mas = "����" Then
    If Len(diffmas) <= 5 Then
    say "calc��Ҫ��һ��ʽ��Ŷ��(�Ӽ��˳��˷�abs sin cos tan�ȣ�"
    Else
    a = Split(diffmas, " ")
    say a(1) & " = " & ScriptControl1.Eval(a(1))
    End If
    End If
    
    If LCase(mas) = "����" Or LCase(mas) = "help" Then
    say "billy��һ��13����Ȼ�����XD  _(:3 ����)_ "
    say "����:PITYHERO233 gavin"
    say "**�������ܣ�help time ping pong (gz����) gzmh reme dict pia calc weather **"
    say "��Ӧ���ģ����� ʱ�� ƹ �� ���� ���� �ʵ� pia ���� ����(��һ��ÿ��������������Ŷ)"
    say "**�������ܣ����㱨ʱ ������ **"
    say "�һ��кܶ�С�ʵ�Ŷ!!"
    say "."
    say "���������뷢�͡�����������"
    End If
    If LCase(mas) = "����" Or LCase(mas) = "gz" Or LCase(mas) = "gzmh" Or LCase(mas) = "��������" Then

    say "��ֱ�������������֣����Ѽ�¼���IP��ַ��"
    say "[��ŭ][��ŭ][��ŭ]"
    End If
    If LCase(mas) = "pong" Then
    say "PING!"
    End If
    
    If diffmas = "�������" Then say uname & " �����˹�������������ʱ��:" & Date & " " & Time
    
    If LCase(diffmas) = "hello" Or LCase(diffmas) = "hi" Then say "hello!!"
    If diffmas = "���" Then say "��á�����"
    
    If InStr(1, diffmas, "billy") > 0 Then
        If InStr(1, diffmas, "gay") > 0 Or InStr(1, diffmas, "����") > 0 Then say "�Ҳ���GAY!!"
        ElseIf InStr(1, diffmas, "gz") > 0 Or InStr(1, diffmas, "����") > 0 Or InStr(1, diffmas, "��������") > 0 Then say "gz??shut up����"

    End If
End If

    If mas = "show" Then Me.Visible = True
    If mas = "hide" Then Me.Visible = False
    
If Left(diffmas, 4) = "��������" Then
    say "BillyBot V2.1 Beta"
    say "һ������ IRC Bot ������"
    say "��Ӧ���ϸ�����GPLv3(General Public License v3),�κ�ʹ�ñ����Դ���ϵͳ����������ɡ�"
    say "������ϸ���Դ������Ȩ���������������ö�δ��Դȫ�������ߣ�������߽�ί��׷����"
    say "Copyleft 2017@PITYHERO233 Inc."
End If
    
If InStr(1, datas, "PING") > 0 Then
    send "PONG"
End If

End Sub

Private Sub sck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If n <= 10 Then
n = n + 1
push "Error occured.num=" & Number & ",desc=" & Description
push "Reconnect in 3 sec..."
sck.Close
Sleep2 (3)

Call Command2_Click
Call Command1_Click
Else
push "time elapsed.Stopped."
sck.Close
n = 0
End If
End Sub

Private Sub Timer1_Timer()
Label1.Caption = sck.State
If sck.State = 7 Then
Label1.BackColor = vbGreen

ElseIf sck.State = 8 Or 9 Then
Label1.BackColor = vbRed

Else
Label1.BackColor = vbWhite
Command1.Enabled = True
End If
Label2.Caption = Date & " " & Time
End Sub

Private Sub Timer2_Timer()
If pointer <> 1 Then
    Dim i As Integer
    For i = 1 To pointer
    If Hour(Now) = h(i) And Minute(Now) = m(i) Then
        say "��Ԥ��������ʱ�䵽�������¼�Ϊ " & note(i)
        h(i) = -1
        m(i) = -1
    End If
    Next

End If
If Minute(Now) = 0 And bt <> Hour(Now) + Minute(Now) Then
    bt = Hour(Now) + Minute(Now)
    Select Case Hour(Now)
    Case 1
    'say
    End Select
End If
End Sub



Private Function query_dict(str As String)
fore = UCase(Left(str, 1))
nex = LCase(Mid(str, 1, 2))
path = App.path & "\ţ����Ӵʵ�\" & fore & "\" & fore & ".txt"
If Dir(path) <> "" Then
Open path For Input As #1

got = False

Do While Not EOF(1)   ' ѭ�����ļ�β��

   Line Input #1, textline   ' ����һ�����ݲ����丳��ĳ������
      If got = True Then
      datas = textline
      Exit Do
      End If
   If textline = str Then got = True
Loop
Close #1
Else
got = False
End If

If got = True Then
query_dict = datas
Else
query_dict = Empty
End If
End Function
