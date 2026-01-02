Attribute VB_Name = "ElpX"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Private As400Dtaq     As String * 8192
Private As400DtaqIn   As String
Private As400DtaqOut  As String
Private As400DtaqLen As Long
Private As400SndOk As Boolean
Private As400RcvOk As Boolean

Type typeXcom
   SrvObj       As String * 12
   SrvMethod    As String * 12
   SrvErr       As String * 10
   usrId       As String * 10
   pcId        As String * 10
   SrvType     As String * 10
   SrvId       As String * 10
   SrvDtaqLib  As String * 10
   SrvDtaqIn   As String * 10
   SrvDTaqOut  As String * 10
   SrvDTaqLen  As String * 5
   jplFree        As String * 5
End Type

Private Xcom As typeXcom
Private blnXCom As Boolean
Public XComlen  As Integer
Public elpSrvXcom  As String
Public elpSrvTxtin As Boolean
Public elpSrvTxtOut As Boolean

Public MsgTxt As String * 8078
Public MsgTxtLen As Integer
Public MsgTxtIndex  As Integer
Private Kerr        As Integer
Public Sub mainSoc_Close()


End Sub


Public Sub mainsoc()
End Sub

Public Function SndRcv_Init()
SndRcv_Init = Null
End Function

Public Sub mainSoc_Environment()

End Sub

Public Sub prtSoc()

End Sub

Public Sub Xcom_UsrId(lX As String)

End Sub


Public Sub Msg_Monitor(Msg As String)

End Sub
