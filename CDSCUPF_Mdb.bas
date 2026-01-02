Attribute VB_Name = "mdbCDSCUPF"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

'---------------------------------------------------------

Public Const recCDSCUPFLen = 400 ' 34 + 366

Type typeCDSCUPF
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    SCCENR                  As String * 1
    SCPERD                  As Long
    SCCNAL                  As String * 2
    SCLPAY                  As String * 25
    SCCPNC                  As String * 6
    SCNOM                   As String * 35
    SCACTY                  As String * 2
    SCCCY                   As String * 3
    SCDTCS                  As Long
    SCCOUR                  As Double
    SCMOUV                  As Currency
    SCMAUM                  As Currency
    SCMAUH                  As Currency
    SCMDIM                  As Currency
    SCMDIH                  As Currency
    SCEOUV                  As Currency
    SCEAUM                  As Currency
    SCEAUH                  As Currency
    SCEDIM                  As Currency
    SCEDIH                  As Currency
    SCCTRC                  As Long
    SCCTRN                  As Long
    SCCTRP                  As Long
    SCRAUG                  As Currency
    SCRDIM                  As Currency
    SCREAU                  As Currency
    SCREDI                  As Currency
    SCRCTC                  As Long
    SCRCTN                  As Long
    SCRCTP                  As Long
   
End Type
'---------------------------------------------------------
Public Function CDSCUPF_GetBuffer(lTxt As String, recCDSCUPF As typeCDSCUPF)
'---------------------------------------------------------
Dim K As Integer, I As Integer
CDSCUPF_GetBuffer = Null
    
    recCDSCUPF.SCCENR = mId$(lTxt, 1, 1)
    recCDSCUPF.SCPERD = CLng(Val(mId$(lTxt, 2, 6)))
    recCDSCUPF.SCCNAL = mId$(lTxt, 8, 2)
    recCDSCUPF.SCLPAY = mId$(lTxt, 10, 25)
    recCDSCUPF.SCCPNC = mId$(lTxt, 35, 6)
    recCDSCUPF.SCNOM = mId$(lTxt, 41, 35)
    recCDSCUPF.SCACTY = mId$(lTxt, 76, 2)
    recCDSCUPF.SCCCY = mId$(lTxt, 78, 3)
    recCDSCUPF.SCDTCS = CLng(Val(mId$(lTxt, 81, 8)))
    recCDSCUPF.SCCOUR = CDbl(Val(mId$(lTxt, 89, 10)) / 100000)
    recCDSCUPF.SCMOUV = CCur(Val(mId$(lTxt, 99, 17)) / 100)
    recCDSCUPF.SCMAUM = CCur(Val(mId$(lTxt, 116, 17)) / 100)
    recCDSCUPF.SCMAUH = CCur(Val(mId$(lTxt, 133, 17)) / 100)
    recCDSCUPF.SCMDIM = CCur(Val(mId$(lTxt, 150, 17)) / 100)
    recCDSCUPF.SCMDIH = CCur(Val(mId$(lTxt, 167, 17)) / 100)

    recCDSCUPF.SCEOUV = CCur(Val(mId$(lTxt, 184, 17)) / 100)
    recCDSCUPF.SCEAUM = CCur(Val(mId$(lTxt, 201, 17)) / 100)
    recCDSCUPF.SCEAUH = CCur(Val(mId$(lTxt, 218, 17)) / 100)
    recCDSCUPF.SCEDIM = CCur(Val(mId$(lTxt, 235, 17)) / 100)
    recCDSCUPF.SCEDIH = CCur(Val(mId$(lTxt, 252, 17)) / 100)
    recCDSCUPF.SCCTRC = CLng(Val(mId$(lTxt, 269, 5)))
    recCDSCUPF.SCCTRN = CLng(Val(mId$(lTxt, 274, 5)))
    recCDSCUPF.SCCTRP = CLng(Val(mId$(lTxt, 279, 5)))

    recCDSCUPF.SCRAUG = CCur(Val(mId$(lTxt, 284, 17)) / 100)
    recCDSCUPF.SCRDIM = CCur(Val(mId$(lTxt, 301, 17)) / 100)
    recCDSCUPF.SCREAU = CCur(Val(mId$(lTxt, 318, 17)) / 100)
    recCDSCUPF.SCREDI = CCur(Val(mId$(lTxt, 335, 17)) / 100)
    recCDSCUPF.SCRCTC = CLng(Val(mId$(lTxt, 352, 5)))
    recCDSCUPF.SCRCTN = CLng(Val(mId$(lTxt, 357, 5)))
    recCDSCUPF.SCRCTP = CLng(Val(mId$(lTxt, 362, 5)))

End Function

Public Sub recCDSCUPF_Init(recCDSCUPF As typeCDSCUPF)
recCDSCUPF.Method = ""
recCDSCUPF.obj = "CDSCUPF"
recCDSCUPF.Err = ""
End Sub




