Attribute VB_Name = "ModSleep"
Private Declare Function timeGetTime Lib "winmm.dll" () As Long '该声明得到系统开机到现在的时间(单位：毫秒)

Public Function Sleep2(T As Long)
    Dim Savetime As Long
    Savetime = timeGetTime '记下开始时的时间
    While timeGetTime < Savetime + T * 1000 '循环等待
        DoEvents '转让控制权
    Wend
End Function

