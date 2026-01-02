Attribute VB_Name = "mdbCDDOSPC"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableCDDOSPX As Recordset
Dim tableCDDOSPXOpen As Boolean
Public mCDDOSPX_Id As Long

Type typeCDDOSPX
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Id                  As String * 16
    Text                As String

End Type

Public recCDDOSPX As typeCDDOSPX

'---------------------------------------------------------

Public Const recCDDOSPCLen = 672 ' 34 + 638

Type typeCDDOSPC
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    DPDPFX                  As String * 3
    DPDNUM                  As Long
    DPREF                   As String * 20
    
    DPCNAL                  As String * 2
    DPLPAY                  As String * 25
    DPCPNC                  As String * 6
    DPNOM                   As String * 35
    
    DPDCTR                  As String * 8
    DPDEXP                  As String * 8
    DPSTAT                  As String * 4
    DPREAC                  As String * 1
    DPREV                   As String * 1
    DPNAT                   As String * 1
    DPUSC1                  As String * 2
    DPUSC2                  As String * 3
    DPUSC3                  As String * 3
    DPGPER                  As Double
    DPGAMT                  As Currency
    DPGCCY                  As String * 3
    DPGTIB                  As String * 4
    DPGTIN                  As String * 6
    DPGTIS                  As String * 3
    DPGCPT                  As String * 25
    DPAMT                   As Currency
    DPCCY                   As String * 3
    DPCPER                  As Double
    DPPLUS                  As Double
    DPMINS                  As Double
    DPQUA                   As String * 1
    DPOUTS                  As Currency
    DPLIAB                  As Currency
    DPLCCY                  As String * 3
    
    DPBNNC                  As String * 6
    DPBNNM                  As String * 35
    DPAPNC                  As String * 6
    DPAPNM                  As String * 35
    DPRCNC                  As String * 6
    DPRCNM                  As String * 35
    DPISNC                  As String * 6
    DPISNM                  As String * 35
    DPCCPT                  As String * 11
    DPCAMD                  As Currency
    DPCAMC                  As Currency
    DPNCPT                  As String * 11
    DPNAMD                  As Currency
    DPNAMC                  As Currency
    DPIBRC                  As String * 4
    DPBBRC                  As String * 4
    RBUSER                  As String * 20
    RBDCRT                  As String * 8
    RBDLUP                  As String * 8
    DPEVAL                  As String * 8
    DPRLRF                  As String * 20
    DPARCH                  As String * 20
    DPDCAC                  As String * 20
   
End Type
'---------------------------------------------------------
Public Function CDDOSPC_GetBuffer(lTxt As String, recCDDOSPC As typeCDDOSPC)
'---------------------------------------------------------
Dim K As Integer, I As Integer
CDDOSPC_GetBuffer = Null
    
    
    recCDDOSPC.DPDPFX = mId$(lTxt, 1, 3)
    recCDDOSPC.DPDNUM = CLng(Val(mId$(lTxt, 4, 6)))
    recCDDOSPC.DPREF = mId$(lTxt, 10, 20)
    recCDDOSPC.DPCNAL = mId$(lTxt, 30, 2)
    recCDDOSPC.DPLPAY = mId$(lTxt, 32, 25)
    recCDDOSPC.DPCPNC = mId$(lTxt, 57, 6)
    recCDDOSPC.DPNOM = mId$(lTxt, 63, 35)
    
    recCDDOSPC.DPDCTR = mId$(lTxt, 98, 8)
    recCDDOSPC.DPDEXP = mId$(lTxt, 106, 8)
    recCDDOSPC.DPSTAT = mId$(lTxt, 114, 4)
    recCDDOSPC.DPREAC = mId$(lTxt, 118, 1)
    recCDDOSPC.DPREV = mId$(lTxt, 119, 1)
    recCDDOSPC.DPNAT = mId$(lTxt, 120, 1)
    recCDDOSPC.DPUSC1 = mId$(lTxt, 121, 2)
    recCDDOSPC.DPUSC2 = mId$(lTxt, 123, 3)
    recCDDOSPC.DPUSC3 = mId$(lTxt, 126, 3)
    recCDDOSPC.DPGPER = CDbl(Val(mId$(lTxt, 129, 7)) / 100)
    recCDDOSPC.DPGAMT = CCur(Val(mId$(lTxt, 136, 17)) / 100)
    recCDDOSPC.DPGCCY = mId$(lTxt, 153, 3)
    recCDDOSPC.DPGTIB = mId$(lTxt, 156, 4)
    recCDDOSPC.DPGTIN = mId$(lTxt, 160, 6)
    recCDDOSPC.DPGTIS = mId$(lTxt, 166, 3)
    recCDDOSPC.DPGCPT = mId$(lTxt, 169, 25)
    recCDDOSPC.DPAMT = CCur(Val(mId$(lTxt, 194, 17)) / 100)
    recCDDOSPC.DPCCY = mId$(lTxt, 211, 3)
    recCDDOSPC.DPCPER = CDbl(Val(mId$(lTxt, 214, 7)) / 100)
    recCDDOSPC.DPPLUS = CDbl(Val(mId$(lTxt, 221, 7)) / 100)
    recCDDOSPC.DPMINS = CDbl(Val(mId$(lTxt, 228, 7)) / 100)
    recCDDOSPC.DPQUA = mId$(lTxt, 235, 1)
    recCDDOSPC.DPOUTS = CCur(Val(mId$(lTxt, 236, 17)) / 100)
    recCDDOSPC.DPLIAB = CCur(Val(mId$(lTxt, 253, 17)) / 100)
    recCDDOSPC.DPLCCY = mId$(lTxt, 270, 3)
    
    recCDDOSPC.DPBNNC = mId$(lTxt, 273, 6)
    recCDDOSPC.DPBNNM = mId$(lTxt, 279, 35)
    recCDDOSPC.DPAPNC = mId$(lTxt, 314, 6)
    recCDDOSPC.DPAPNM = mId$(lTxt, 320, 35)
    recCDDOSPC.DPRCNC = mId$(lTxt, 355, 6)
    recCDDOSPC.DPRCNM = mId$(lTxt, 361, 35)
    recCDDOSPC.DPISNC = mId$(lTxt, 396, 6)
    recCDDOSPC.DPISNM = mId$(lTxt, 402, 35)
    recCDDOSPC.DPCCPT = mId$(lTxt, 437, 11)
    recCDDOSPC.DPCAMD = CCur(Val(mId$(lTxt, 448, 17)) / 100)
    recCDDOSPC.DPCAMC = CCur(Val(mId$(lTxt, 465, 17)) / 100)
    recCDDOSPC.DPNCPT = mId$(lTxt, 482, 11)
    recCDDOSPC.DPNAMD = CCur(Val(mId$(lTxt, 493, 17)) / 100)
    recCDDOSPC.DPNAMC = CCur(Val(mId$(lTxt, 510, 17)) / 100)
    
    recCDDOSPC.DPIBRC = mId$(lTxt, 527, 4)
    recCDDOSPC.DPBBRC = mId$(lTxt, 531, 4)
    recCDDOSPC.RBUSER = mId$(lTxt, 535, 20)
    recCDDOSPC.RBDCRT = mId$(lTxt, 555, 8)
    recCDDOSPC.RBDLUP = mId$(lTxt, 563, 8)
    recCDDOSPC.DPEVAL = mId$(lTxt, 571, 8)
    recCDDOSPC.DPRLRF = mId$(lTxt, 579, 20)
    recCDDOSPC.DPARCH = mId$(lTxt, 599, 20)
    recCDDOSPC.DPDCAC = mId$(lTxt, 619, 20)

End Function


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableCDDOSPX_Close()
'-----------------------------------------------------
If tableCDDOSPXOpen Then
    tableCDDOSPX.Close
    tableCDDOSPXOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableCDDOSPX_GetBuffer(recCDDOSPX As typeCDDOSPX)
'---------------------------------------------------------
recCDDOSPX.Id = tableCDDOSPX("Id")
recCDDOSPX.Text = tableCDDOSPX("Text")

End Sub


'-----------------------------------------------------
Sub tableCDDOSPX_Open()
'-----------------------------------------------------

If Not tableCDDOSPXOpen Then
    Set tableCDDOSPX = MDB.OpenRecordset("CDDOSPX")
    tableCDDOSPX.Index = "PrimaryKey"
    tableCDDOSPXOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableCDDOSPX_PutBuffer(recCDDOSPX As typeCDDOSPX)
'---------------------------------------------------------

tableCDDOSPX("Id") = recCDDOSPX.Id
tableCDDOSPX("Text") = recCDDOSPX.Text
End Sub


'---------------------------------------------------------
Public Function tableCDDOSPX_Read(recCDDOSPX As typeCDDOSPX) As Integer
'---------------------------------------------------------

On Error GoTo tableCDDOSPX_Read_Error
tableCDDOSPX_Read = 0


Select Case Trim(recCDDOSPX.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableCDDOSPX.Seek "=", recCDDOSPX.Id
                        If tableCDDOSPX.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableCDDOSPX.Seek "<=", recCDDOSPX.Id
                        If tableCDDOSPX.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableCDDOSPX.Seek ">=", recCDDOSPX.Id
                        If tableCDDOSPX.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableCDDOSPX.Seek ">", recCDDOSPX.Id
                        If tableCDDOSPX.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableCDDOSPX.MoveNext
                        If tableCDDOSPX.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableCDDOSPX.MovePrevious
                        If tableCDDOSPX.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableCDDOSPX.MoveFirst
                        If tableCDDOSPX.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableCDDOSPX.MoveLast
                        If tableCDDOSPX.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recCDDOSPX.Method <> "AddNew      " Then
    Call tableCDDOSPX_GetBuffer(recCDDOSPX)
End If

Exit Function

'---------------------------------------------------------
tableCDDOSPX_Read_Error:
'---------------------------------------------------------

    tableCDDOSPX_Read = Err
    Resume tableCDDOSPX_Read_End

tableCDDOSPX_Read_End:

End Function

'---------------------------------------------------------
Public Function tableCDDOSPX_Update(recCDDOSPX As typeCDDOSPX) As Integer
'---------------------------------------------------------

On Error GoTo tableCDDOSPXUpdate_Error
tableCDDOSPX_Update = 0

Select Case Trim(recCDDOSPX.Method)

    Case "AddNew"
                        tableCDDOSPX.AddNew
                        Call tableCDDOSPX_PutBuffer(recCDDOSPX)
                        tableCDDOSPX.Update
    Case "Update"
                        tableCDDOSPX.Edit
                        Call tableCDDOSPX_PutBuffer(recCDDOSPX)
                        tableCDDOSPX.Update
    Case "Delete"
                        tableCDDOSPX.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableCDDOSPXUpdate_Error:
'---------------------------------------------------------
    tableCDDOSPX_Update = Err
    Resume tableCDDOSPXUpdate_End

tableCDDOSPXUpdate_End:

End Function

Public Function dbCDDOSPX_Import(lFileName As String, lNb As Long)
Dim X As String, xInput As String
On Error GoTo Error_Handler

Dim I As Integer, blnOk As Boolean
dbCDDOSPX_Import = Null
lNb = 0: I = 0
X = Dir(lFileName)
If X = "" Then dbCDDOSPX_Import = "? dbCDDOSPX_Import : Le fichier des mouvments n'existe pas": Exit Function


MDB.Execute "delete * from CDDOSPX"
tableCDDOSPX_Open
recCDDOSPX_Init recCDDOSPX
recCDDOSPX.Method = "AddNew"

Open lFileName For Input As #1

blnOk = False
Do Until EOF(1)
    Line Input #1, xInput
    
    If mId$(xInput, 1, 3) = "$$$" Then
        blnOk = True
        'SrvCptP0_Amj = mId$(xInput, 4, 8)
        I = Val(mId$(xInput, 12, 9))
        If I <> lNb Then
            X = "? dbCDDOSPX_Import : nombre enregistrements lus"
            Call MsgBox(X, vbCritical, "dbCDDOSPX_Import")
            dbCDDOSPX_Import = X: Exit Function
            Exit Do
        End If
    End If

    lNb = lNb + 1
    recCDDOSPX.Id = mId$(xInput, 1, 9)
    recCDDOSPX.Method = "AddNew"
    recCDDOSPX.Text = xInput
    dbCDDOSPX_Update recCDDOSPX
 
Loop

Close
tableCDDOSPX_Close


'If Not blnOk Then
'    X = "? dbCDDOSPX_Import : manque fin de fichier "
'    Call MsgBox(X, vbCritical, "dbCDDOSPX_Import")
'    dbCDDOSPX_Import = X: Exit Function
'End If

Exit Function
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------
    X = "? dbCDDOSPX_Import : " & Err & " : " & Error(Err)
    Call MsgBox(X, vbCritical, "dbCDDOSPX_Import")
    dbCDDOSPX_Import = X: Exit Function

End Function








'-----------------------------------------------------
Sub dbCDDOSPX_Error(recCDDOSPX As typeCDDOSPX)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recCDDOSPX.Id & ": " & Chr$(13)

Select Case mId$(recCDDOSPX.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recCDDOSPX.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbCDDOSPX.bas :  ( " & Trim(recCDDOSPX.obj) & " : " & Trim(recCDDOSPX.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbCDDOSPX_ReadE(recCDDOSPX As typeCDDOSPX)
'-----------------------------------------------------

dbCDDOSPX_ReadE = Null

recCDDOSPX.Err = tableCDDOSPX_Read(recCDDOSPX)
If recCDDOSPX.Err > 0 Then

'    If recCDDOSPX.Err < 9990 Or recCDDOSPX.Err >= 9999 Then
        Call dbCDDOSPX_Error(recCDDOSPX)
        dbCDDOSPX_ReadE = recCDDOSPX.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbCDDOSPX_Update(recCDDOSPX As typeCDDOSPX)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbCDDOSPX_Update = Null


recCDDOSPX.Err = tableCDDOSPX_Update(recCDDOSPX)

If recCDDOSPX.Err <> 0 Then
    Call dbCDDOSPX_Error(recCDDOSPX)
    dbCDDOSPX_Update = recCDDOSPX.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recCDDOSPX_Init(recCDDOSPX As typeCDDOSPX)
recCDDOSPX.Method = ""
recCDDOSPX.obj = "CDDOSPX"
recCDDOSPX.Err = ""
recCDDOSPX.Id = ""
recCDDOSPX.Text = ""
End Sub




