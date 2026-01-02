Attribute VB_Name = "mdbCDPOSPC"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableCDPOSPX As Recordset
Dim tableCDPOSPXOpen As Boolean
Public mCDPOSPX_Id As Long

Type typeCDPOSPX
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Id                  As String * 16
    Text                As String

End Type

Public recCDPOSPX As typeCDPOSPX

'---------------------------------------------------------

Public Const recCDPOSPCLen = 336 ' 34 + 302

Type typeCDPOSPC
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    PPDPFX                  As String * 3
    PPDNUM                  As Long
    PPPKEY                  As Long
    
    PPLOT                   As Long
    PPPIE                   As Long
    PPLIG                   As Long
    
    PPEPFX                  As String * 3
    PPENUM                  As Long
    
    PPSRDF                  As String * 20
    PPETYP                  As String * 10
    PPSTAT                  As String * 1
    PPSTAL                  As String * 20
    PPSTEP                  As String * 2
    PPREAC                  As String * 1
    PPDCRT                  As String * 8
    PPDLUP                  As String * 8
    PPUSER                  As String * 20
    
    PPBRC                   As String * 4
    PPDVAL                  As String * 8
    PPATIB                  As String * 4
    PPATIN                  As String * 6
    PPATIS                  As String * 3
    PPCPT                   As String * 25
    PPTRCD                  As String * 3
    PPDBCR                  As String * 1
    PPAMT                   As Currency
    PPCCY                   As String * 3
    PPACTY                  As String * 2
    PPSPCD                  As String * 6
    PPSKCD                  As String * 2
    
    PPPART                  As Long
    PPDTRT                  As String * 8
    PPDTOP                  As String * 8
    PPDTVL                  As String * 8
    PPMNT                   As Currency
    PPLIB                   As String * 30
   
End Type
'---------------------------------------------------------
Public Function CDPOSPC_GetBuffer(lTxt As String, recCDPOSPC As typeCDPOSPC)
'---------------------------------------------------------
Dim K As Integer, I As Integer
CDPOSPC_GetBuffer = Null
    
    recCDPOSPC.PPDPFX = mId$(lTxt, 1, 3)
    recCDPOSPC.PPDNUM = CLng(Val(mId$(lTxt, 4, 6)))
    recCDPOSPC.PPPKEY = CLng(Val(mId$(lTxt, 10, 12)))
    recCDPOSPC.PPLOT = CLng(Val(mId$(lTxt, 22, 4)))
    recCDPOSPC.PPPIE = CLng(Val(mId$(lTxt, 26, 7)))
    recCDPOSPC.PPLIG = CLng(Val(mId$(lTxt, 33, 4)))
    recCDPOSPC.PPEPFX = mId$(lTxt, 37, 3)
    recCDPOSPC.PPENUM = CLng(Val(mId$(lTxt, 40, 6)))
    recCDPOSPC.PPSRDF = mId$(lTxt, 46, 20)
    recCDPOSPC.PPETYP = mId$(lTxt, 66, 10)
    recCDPOSPC.PPSTAT = mId$(lTxt, 76, 1)
    recCDPOSPC.PPSTAL = mId$(lTxt, 77, 20)
    recCDPOSPC.PPSTEP = mId$(lTxt, 97, 2)
    recCDPOSPC.PPREAC = mId$(lTxt, 99, 1)
    recCDPOSPC.PPDCRT = mId$(lTxt, 100, 8)
    recCDPOSPC.PPDLUP = mId$(lTxt, 108, 8)
    recCDPOSPC.PPUSER = mId$(lTxt, 116, 20)
    recCDPOSPC.PPBRC = mId$(lTxt, 136, 4)
    recCDPOSPC.PPDVAL = mId$(lTxt, 140, 8)
    recCDPOSPC.PPATIB = mId$(lTxt, 148, 4)
    recCDPOSPC.PPATIN = mId$(lTxt, 152, 6)
    recCDPOSPC.PPATIS = mId$(lTxt, 158, 3)
    recCDPOSPC.PPCPT = mId$(lTxt, 161, 25)
    recCDPOSPC.PPTRCD = mId$(lTxt, 186, 3)
    recCDPOSPC.PPDBCR = mId$(lTxt, 189, 1)
    recCDPOSPC.PPAMT = CCur(Val(mId$(lTxt, 190, 17)) / 100)
    recCDPOSPC.PPCCY = mId$(lTxt, 207, 3)
    recCDPOSPC.PPACTY = mId$(lTxt, 210, 2)
    recCDPOSPC.PPSPCD = mId$(lTxt, 212, 6)
    recCDPOSPC.PPSKCD = mId$(lTxt, 218, 2)

    recCDPOSPC.PPPART = CLng(Val(mId$(lTxt, 220, 12)))
    recCDPOSPC.PPDTRT = mId$(lTxt, 232, 8)
    recCDPOSPC.PPDTOP = mId$(lTxt, 240, 8)
    recCDPOSPC.PPDTVL = mId$(lTxt, 248, 8)
    recCDPOSPC.PPMNT = CCur(Val(mId$(lTxt, 256, 17)) / 100)
    recCDPOSPC.PPLIB = mId$(lTxt, 273, 30)
End Function


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableCDPOSPX_Close()
'-----------------------------------------------------
If tableCDPOSPXOpen Then
    tableCDPOSPX.Close
    tableCDPOSPXOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableCDPOSPX_GetBuffer(recCDPOSPX As typeCDPOSPX)
'---------------------------------------------------------
recCDPOSPX.Id = tableCDPOSPX("Id")
recCDPOSPX.Text = tableCDPOSPX("Text")

End Sub


'-----------------------------------------------------
Sub tableCDPOSPX_Open()
'-----------------------------------------------------

If Not tableCDPOSPXOpen Then
    Set tableCDPOSPX = MDB.OpenRecordset("CDPOSPX")
    tableCDPOSPX.Index = "PrimaryKey"
    tableCDPOSPXOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableCDPOSPX_PutBuffer(recCDPOSPX As typeCDPOSPX)
'---------------------------------------------------------

tableCDPOSPX("Id") = recCDPOSPX.Id
tableCDPOSPX("Text") = recCDPOSPX.Text
End Sub


'---------------------------------------------------------
Public Function tableCDPOSPX_Read(recCDPOSPX As typeCDPOSPX) As Integer
'---------------------------------------------------------

On Error GoTo tableCDPOSPX_Read_Error
tableCDPOSPX_Read = 0


Select Case Trim(recCDPOSPX.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableCDPOSPX.Seek "=", recCDPOSPX.Id
                        If tableCDPOSPX.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableCDPOSPX.Seek "<=", recCDPOSPX.Id
                        If tableCDPOSPX.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableCDPOSPX.Seek ">=", recCDPOSPX.Id
                        If tableCDPOSPX.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableCDPOSPX.Seek ">", recCDPOSPX.Id
                        If tableCDPOSPX.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableCDPOSPX.MoveNext
                        If tableCDPOSPX.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableCDPOSPX.MovePrevious
                        If tableCDPOSPX.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableCDPOSPX.MoveFirst
                        If tableCDPOSPX.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableCDPOSPX.MoveLast
                        If tableCDPOSPX.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recCDPOSPX.Method <> "AddNew      " Then
    Call tableCDPOSPX_GetBuffer(recCDPOSPX)
End If

Exit Function

'---------------------------------------------------------
tableCDPOSPX_Read_Error:
'---------------------------------------------------------

    tableCDPOSPX_Read = Err
    Resume tableCDPOSPX_Read_End

tableCDPOSPX_Read_End:

End Function

'---------------------------------------------------------
Public Function tableCDPOSPX_Update(recCDPOSPX As typeCDPOSPX) As Integer
'---------------------------------------------------------

On Error GoTo tableCDPOSPXUpdate_Error
tableCDPOSPX_Update = 0

Select Case Trim(recCDPOSPX.Method)

    Case "AddNew"
                        tableCDPOSPX.AddNew
                        Call tableCDPOSPX_PutBuffer(recCDPOSPX)
                        tableCDPOSPX.Update
    Case "Update"
                        tableCDPOSPX.Edit
                        Call tableCDPOSPX_PutBuffer(recCDPOSPX)
                        tableCDPOSPX.Update
    Case "Delete"
                        tableCDPOSPX.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableCDPOSPXUpdate_Error:
'---------------------------------------------------------
    tableCDPOSPX_Update = Err
    Resume tableCDPOSPXUpdate_End

tableCDPOSPXUpdate_End:

End Function

Public Function dbCDPOSPX_Import(lFileName As String, lNb As Long)
Dim X As String, xInput As String
On Error GoTo Error_Handler

Dim I As Integer, blnOk As Boolean
dbCDPOSPX_Import = Null

lNb = 0: I = 0
X = Dir(lFileName)
If X = "" Then dbCDPOSPX_Import = "? dbCDPOSPX_Import : Le fichier des mouvments n'existe pas": Exit Function


MDB.Execute "delete * from CDPOSPX"
tableCDPOSPX_Open
recCDPOSPX_Init recCDPOSPX
recCDPOSPX.Method = "AddNew"

Open lFileName For Input As #1

blnOk = False
Do Until EOF(1)
    Line Input #1, xInput
    
    If mId$(xInput, 1, 3) = "$$$" Then
        blnOk = True
        'SrvCptP0_Amj = mId$(xInput, 4, 8)
        I = Val(mId$(xInput, 12, 9))
        If I <> lNb Then
            X = "? dbCDPOSPX_Import : nombre enregistrements lus"
            Call MsgBox(X, vbCritical, "dbCDPOSPX_Import")
            dbCDPOSPX_Import = X: Exit Function
            Exit Do
        End If
    End If

    lNb = lNb + 1
    recCDPOSPX.Id = mId$(xInput, 1, 9) & Format$(lNb, "0000000")
    recCDPOSPX.Method = "AddNew"
    recCDPOSPX.Text = xInput
    dbCDPOSPX_Update recCDPOSPX
 
Loop

Close
tableCDPOSPX_Close


'If Not blnOk Then
'    X = "? dbCDPOSPX_Import : manque fin de fichier "
'    Call MsgBox(X, vbCritical, "dbCDPOSPX_Import")
'    dbCDPOSPX_Import = X: Exit Function
'End If

Exit Function
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------
    X = "? dbCDPOSPX_Import : " & Err & " : " & Error(Err)
    Call MsgBox(X, vbCritical, "dbCDPOSPX_Import")
    dbCDPOSPX_Import = X: Exit Function

End Function








'-----------------------------------------------------
Sub dbCDPOSPX_Error(recCDPOSPX As typeCDPOSPX)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recCDPOSPX.Id & ": " & Chr$(13)

Select Case mId$(recCDPOSPX.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recCDPOSPX.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbCDPOSPX.bas :  ( " & Trim(recCDPOSPX.obj) & " : " & Trim(recCDPOSPX.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbCDPOSPX_ReadE(recCDPOSPX As typeCDPOSPX)
'-----------------------------------------------------

dbCDPOSPX_ReadE = Null

recCDPOSPX.Err = tableCDPOSPX_Read(recCDPOSPX)
If recCDPOSPX.Err > 0 Then

'    If recCDPOSPX.Err < 9990 Or recCDPOSPX.Err >= 9999 Then
        Call dbCDPOSPX_Error(recCDPOSPX)
        dbCDPOSPX_ReadE = recCDPOSPX.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbCDPOSPX_Update(recCDPOSPX As typeCDPOSPX)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbCDPOSPX_Update = Null


recCDPOSPX.Err = tableCDPOSPX_Update(recCDPOSPX)

If recCDPOSPX.Err <> 0 Then
    Call dbCDPOSPX_Error(recCDPOSPX)
    dbCDPOSPX_Update = recCDPOSPX.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recCDPOSPX_Init(recCDPOSPX As typeCDPOSPX)
recCDPOSPX.Method = ""
recCDPOSPX.obj = "CDPOSPX"
recCDPOSPX.Err = ""
recCDPOSPX.Id = ""
recCDPOSPX.Text = ""
End Sub



