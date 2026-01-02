Attribute VB_Name = "mdbCDEDel005"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

'---------------------------------------------------------

Public Const recCDEDel005Len = 412 ' 34 + 378

Type typeCDEDel005
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    LOGCOSOC                As Long
    LOGAGENCE               As Long
    LOGCPTAMJ               As String * 8
    LOGCPTEUR               As Long
    LOGSERVIC               As String * 5
    LOGPROGR                As String * 20
    LOGPROFIL               As String * 20
    LOGREFCON               As String * 16
    LOGDEVISE               As String * 3
    LOGCOMPTE               As String * 11
    LOGCODERR               As String * 12
    LOGTEXTE1               As String * 128
    LOGTEXTE2               As String * 128
    LOGSYSAMJ               As String * 8
    LOGSYSHMS               As String * 6
   
End Type
'---------------------------------------------------------
Public Function CDEDel005_GetBuffer(lTxt As String, recCDEDel005 As typeCDEDel005)
'---------------------------------------------------------
Dim K As Integer, I As Integer
CDEDel005_GetBuffer = Null
    
    recCDEDel005.LOGCOSOC = CLng(Val(mId$(lTxt, 1, 3)))
    recCDEDel005.LOGAGENCE = CLng(Val(mId$(lTxt, 4, 3)))
    recCDEDel005.LOGCPTAMJ = mId$(lTxt, 7, 8)
    recCDEDel005.LOGCPTEUR = CLng(Val(mId$(lTxt, 15, 7)))
    recCDEDel005.LOGSERVIC = mId$(lTxt, 22, 5)
    recCDEDel005.LOGPROGR = mId$(lTxt, 27, 20)
    recCDEDel005.LOGPROFIL = mId$(lTxt, 47, 20)
    recCDEDel005.LOGREFCON = mId$(lTxt, 67, 16)
    recCDEDel005.LOGDEVISE = mId$(lTxt, 83, 3)
    recCDEDel005.LOGCOMPTE = mId$(lTxt, 86, 11)
    recCDEDel005.LOGCODERR = mId$(lTxt, 97, 12)
    recCDEDel005.LOGTEXTE1 = mId$(lTxt, 109, 128)
    recCDEDel005.LOGTEXTE2 = mId$(lTxt, 237, 128)
    recCDEDel005.LOGSYSAMJ = mId$(lTxt, 365, 8)
    recCDEDel005.LOGSYSHMS = mId$(lTxt, 373, 6)

End Function

Public Sub recCDEDel005_Init(recCDEDel005 As typeCDEDel005)
recCDEDel005.Method = ""
recCDEDel005.obj = "CDEDel005"
recCDEDel005.Err = ""
End Sub

