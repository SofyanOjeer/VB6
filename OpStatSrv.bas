Attribute VB_Name = "srvOpStat"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableOpStat As Recordset

Public Const recOpStatLen = 146 '34 + 112

Type typeOpStat
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Référence              As String * 6
    CodeOpération          As String * 10
    Société                As String * 3
    Agence                 As String * 3
    Devise                 As String * 3
    Brut                   As Currency
    Sens                   As String * 1
    ComMontantFRF          As Currency
    BicIdSender            As String * 11
    BicIdCorrespondant     As String * 11
    BicIdReceiver          As String * 11
    xNature                As String * 10
    xAMJ                   As String * 8
    xHMS                   As String * 6
    
  End Type
    
Public arrOpStat() As typeOpStat
Public arrOpStatNb As Integer
Public arrOpStatNbMax As Integer
Public arrOpStatIndex As Integer
Public arrOpStatSuite As Boolean

Public recOpStat   As typeOpStat
Public paramOpStat As typeParamSnap
'-----------------------------------------------------
Sub tableOpStat_Close()
'-----------------------------------------------------

tableOpStat.Close

End Sub


'---------------------------------------------------------
Public Sub tableOpStat_GetBuffer(recOpStat As typeOpStat)
'---------------------------------------------------------

recOpStat.Référence = tableOpStat("Référence")
recOpStat.CodeOpération = tableOpStat("CodeOpération")
recOpStat.Société = tableOpStat("Société")
recOpStat.Agence = tableOpStat("Agence")
recOpStat.Devise = tableOpStat("Devise")
recOpStat.Brut = tableOpStat("Brut")
recOpStat.Sens = tableOpStat("Sens")
recOpStat.ComMontantFRF = tableOpStat("ComMontantFRF")
recOpStat.BicIdSender = tableOpStat("BicIdSender")
recOpStat.BicIdCorrespondant = tableOpStat("BicIdCorrespondant")
recOpStat.BicIdReceiver = tableOpStat("BicIdReceiver")
recOpStat.xNature = tableOpStat("xNature")
recOpStat.xAMJ = tableOpStat("xAMJ")
recOpStat.xHMS = tableOpStat("xHMS")

End Sub


'-----------------------------------------------------
Sub tableOpStat_Open()
'-----------------------------------------------------

Set tableOpStat = MDB.OpenRecordset("OpStat")
tableOpStat.Index = "PrimaryKey"

End Sub

'---------------------------------------------------------
Public Sub tableOpStat_PutBuffer(recOpStat As typeOpStat)
'---------------------------------------------------------
Dim X As String
paramOpStat.Nb = paramOpStat.Nb + 1
X = ""
Select Case Trim(paramOpStat.sortK1)
    Case "Devise": X = recOpStat.Devise
    Case "Sender": X = recOpStat.BicIdSender
    Case "Correspondant": X = recOpStat.BicIdCorrespondant
    Case "Receiver": X = recOpStat.BicIdReceiver
    Case "Référence": X = recOpStat.Référence
End Select

Select Case Trim(paramOpStat.sortK2)
    Case "Devise": X = X & recOpStat.Devise
    Case "Sender": X = X & recOpStat.BicIdSender
    Case "Correspondant": X = X & recOpStat.BicIdCorrespondant
    Case "Receiver": X = X & recOpStat.BicIdReceiver
    Case "Référence": X = X & recOpStat.Référence
End Select

tableOpStat("K") = X & recOpStat.Référence
tableOpStat("Référence") = recOpStat.Référence
tableOpStat("CodeOpération") = recOpStat.CodeOpération
tableOpStat("Société") = recOpStat.Société
tableOpStat("Agence") = recOpStat.Agence
tableOpStat("Devise") = recOpStat.Devise
tableOpStat("Brut") = recOpStat.Brut
tableOpStat("Sens") = recOpStat.Sens
tableOpStat("ComMontantFRF") = recOpStat.ComMontantFRF
tableOpStat("BicIdSender") = recOpStat.BicIdSender
tableOpStat("BicIdCorrespondant") = recOpStat.BicIdCorrespondant
tableOpStat("BicIdReceiver") = recOpStat.BicIdReceiver
tableOpStat("xNature") = recOpStat.xNature
tableOpStat("xAMJ") = recOpStat.xAMJ
tableOpStat("xHMS") = recOpStat.xHMS

End Sub


'---------------------------------------------------------
Public Function tableOpStat_Read(recOpStat As typeOpStat) As Integer
'---------------------------------------------------------

On Error GoTo tableOpStat_Read_Error
tableOpStat_Read = 0


Select Case recOpStat.Method
    Case "MoveNext    "
                        tableOpStat.MoveNext
                        If tableOpStat.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableOpStat.MovePrevious
                        If tableOpStat.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst   "
                        tableOpStat.MoveFirst
                        If tableOpStat.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast    "
                        tableOpStat.MoveLast
                        If tableOpStat.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recOpStat.Method <> "AddNew      " Then
    Call tableOpStat_GetBuffer(recOpStat)
End If

Exit Function

'---------------------------------------------------------
tableOpStat_Read_Error:
'---------------------------------------------------------

    tableOpStat_Read = Err
    Resume tableOpStat_Read_End

tableOpStat_Read_End:

End Function

'---------------------------------------------------------
Public Function tableOpStat_Update(recOpStat As typeOpStat) As Integer
'---------------------------------------------------------

On Error GoTo tableOpStatUpdate_Error
tableOpStat_Update = 0

Select Case recOpStat.Method

    Case "AddNew      "
                        tableOpStat.AddNew
                        Call tableOpStat_PutBuffer(recOpStat)
                        tableOpStat.Update
    Case "Update      "
                        tableOpStat.Edit
                        Call tableOpStat_PutBuffer(recOpStat)
                        tableOpStat.Update
    Case "Delete      "
                        tableOpStat.Delete
    Case Else
                        Error 9999
End Select


Exit Function

tableOpStatUpdate_Error:
'---------------------------------------------------------
    tableOpStat_Update = Err
    Resume tableOpStatUpdate_End

tableOpStatUpdate_End:

End Function








'-----------------------------------------------------
Public Function Monitor(recOpStat As typeOpStat)
'-----------------------------------------------------

arrOpStatSuite = False
Select Case mId$(Trim(recOpStat.Method), 1, 4)
    Case "Snap"
              Monitor = Snap(recOpStat)
    Case Else
                recOpStat.Err = recOpStat.Method
                Call ErrorX(recOpStat)
                Monitor = recOpStat.Err
End Select

End Function

'-----------------------------------------------------
Sub ErrorX(recOpStat As typeOpStat)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Lr Attributs: "

Select Case mId$(recOpStat.Err, 9, 2)
    Case "22"
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recOpStat.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : .bas  ( " _
                & Trim(recOpStat.obj) & " : " & Trim(recOpStat.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function GetBuffer(recOpStat As typeOpStat)
'---------------------------------------------------------
Dim K As Integer, I As Integer
GetBuffer = Null
recOpStat.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recOpStat.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recOpStat.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recOpStat.Err = Space$(10) Then
    recOpStat.Référence = mId$(MsgTxt, K + 1, 6)
    recOpStat.CodeOpération = mId$(MsgTxt, K + 7, 10)
    recOpStat.Société = mId$(MsgTxt, K + 17, 3)
    recOpStat.Agence = mId$(MsgTxt, K + 20, 3)
    recOpStat.Devise = mId$(MsgTxt, K + 23, 3)
    recOpStat.Brut = CCur(Val(mId$(MsgTxt, K + 26, 17)) / 100)
    recOpStat.Sens = mId$(MsgTxt, K + 43, 1)
    recOpStat.ComMontantFRF = CCur(Val(mId$(MsgTxt, K + 44, 12)) / 100)
    recOpStat.BicIdSender = mId$(MsgTxt, K + 56, 11)
    recOpStat.BicIdCorrespondant = mId$(MsgTxt, K + 67, 11)
    recOpStat.BicIdReceiver = mId$(MsgTxt, K + 78, 11)
    recOpStat.xNature = mId$(MsgTxt, K + 89, 10)
    recOpStat.xAMJ = mId$(MsgTxt, K + 99, 8)
    recOpStat.xHMS = mId$(MsgTxt, K + 107, 6)
Else
    GetBuffer = recOpStat.Err
End If

MsgTxtIndex = MsgTxtIndex + recOpStatLen

End Function

'---------------------------------------------------------
Private Sub PutBuffer(recOpStat As typeOpStat)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recOpStat.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recOpStat.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 6) = recOpStat.Référence
Mid$(MsgTxt, K + 7, 10) = recOpStat.CodeOpération
Mid$(MsgTxt, K + 17, 3) = recOpStat.Société
Mid$(MsgTxt, K + 20, 3) = recOpStat.Agence
Mid$(MsgTxt, K + 23, 3) = recOpStat.Devise
Mid$(MsgTxt, K + 26, 17) = Format$(recOpStat.Brut * 100, "00000000000000000")
Mid$(MsgTxt, K + 43, 1) = recOpStat.Sens
Mid$(MsgTxt, K + 44, 12) = Format$(recOpStat.ComMontantFRF * 100, "000000000000")
Mid$(MsgTxt, K + 56, 11) = recOpStat.BicIdSender
Mid$(MsgTxt, K + 67, 11) = recOpStat.BicIdCorrespondant
Mid$(MsgTxt, K + 78, 11) = recOpStat.BicIdReceiver
Mid$(MsgTxt, K + 89, 10) = recOpStat.xNature
Mid$(MsgTxt, K + 99, 8) = recOpStat.xAMJ
Mid$(MsgTxt, K + 107, 6) = recOpStat.xHMS

MsgTxtLen = MsgTxtLen + recOpStatLen
End Sub



'---------------------------------------------------------
Private Function Snap(recOpStat As typeOpStat)
'---------------------------------------------------------
Dim I As Integer
Snap = "?"
MsgTxtLen = 0
Call PutBuffer(recOpStat)
Call PutBuffer(arrOpStat(0))
If IsNull(SndRcv()) Then
    Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(GetBuffer(recOpStat)) Then
            recOpStat.Method = "AddNew      "
            recOpStat.Devise = DevX(recOpStat.Devise)
            Call tableOpStat_Update(recOpStat)
'           Call srvOpStat.AddItem(recOpStat)
            arrOpStatSuite = True
        Else
            arrOpStatSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub Init(recOpStat As typeOpStat)
'---------------------------------------------------------
MsgTxt = Space$(recOpStatLen)
MsgTxtIndex = 0
Call GetBuffer(recOpStat)
recOpStat.obj = "SRVOPSTAT"
End Sub

'---------------------------------------------------------
Public Sub AddItem(recOpStat As typeOpStat)
'---------------------------------------------------------
          
arrOpStatNb = arrOpStatNb + 1
    
If arrOpStatNb > arrOpStatNbMax Then
    arrOpStatNbMax = arrOpStatNbMax + 50
    ReDim Preserve arrOpStat(arrOpStatNbMax)
End If
recOpStat.Method = ""
arrOpStatIndex = arrOpStatNb
arrOpStat(arrOpStatIndex) = recOpStat
End Sub


Public Sub Filtre(xMethod As String, xNature As String)

recOpStat.Method = xMethod
recOpStat.xNature = xNature
recOpStat.xAMJ = paramOpStat.AmjMin
recOpStat.xHMS = paramOpStat.HMSMin
arrOpStat(0) = recOpStat
arrOpStat(0).Référence = "99999999999"
arrOpStat(0).xAMJ = paramOpStat.AmjMax
arrOpStat(0).xHMS = paramOpStat.HMSMax

arrOpStatSuite = True
Do Until Not arrOpStatSuite
    srvOpStat.Monitor recOpStat
'    recOpStat = arrOpStat(arrOpStatNb)
    recOpStat.Method = xMethod & "+"
    frmOpTrf.lblOpStatNb = Format$(paramOpStat.Nb, "### ### ##0") & " dossiers"
Loop

End Sub
