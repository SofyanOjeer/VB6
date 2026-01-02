Attribute VB_Name = "srvYCRITAB0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCRITAB0Len = 138 ' 34 + 104
Public Const recYCRITAB0_Block = 50
Public Const memoYCRITAB0Len = 104
Public Const constYCRITAB0 = "YCRITAB0  "

Type typeYCRITAB0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    CRITABETA       As Integer                        ' ETABLISSEMENT
    CRITABNUM       As Long                           ' NUMERO TABLE
    CRITABARG       As String * 15                    ' ARGUMENT
    CRITABDON       As String * 80                  ' DONNEES
End Type
    
    
'---------------------------------------------------------
Public Sub srvYCRITAB0_PutBuffer(recYCRITAB0 As typeYCRITAB0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCRITAB0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCRITAB0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYCRITAB0.CRITABETA, "0000 ")
    Mid$(MsgTxt, K + 6, 4) = Format$(recYCRITAB0.CRITABNUM, "000 ")
    Mid$(MsgTxt, K + 10, 15) = recYCRITAB0.CRITABARG
    Mid$(MsgTxt, K + 25, 80) = recYCRITAB0.CRITABDON
    
MsgTxtLen = MsgTxtLen + recYCRITAB0Len
End Sub


