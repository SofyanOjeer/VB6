Attribute VB_Name = "srvYLIBEL0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYLIBEL0Len = 89 ' 34 +55
Public Const recYLIBEL0_Block = 100
Public Const memoYLIBEL0Len = 55
Public Const constYLIBEL0 = "YLIBEL0  "

Type typeYLIBEL0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    LIBELETA        As Integer                        ' ETABLISSEMENT
    LIBELPIE        As Long                           ' NUMERO DE PIECE
    LIBELECR        As Long                           ' NUMERO D'ECRITURE
    LIBELNUM        As Long                           ' NUMERO DE LIBELLE
    LIBELLIB        As String * 30                    ' LIBELLE
End Type
    
    
Public arrYLIBEL0() As typeYLIBEL0
Public arrYLIBEL0_NB As Integer
Public arrYLIBEL0_NBMax As Integer
Public arrYLIBEL0_Index As Integer
Public arrYLIBEL0_Suite As Boolean

'-----------------------------------------------------
Function srvYLIBEL0_Update(recYLIBEL0 As typeYLIBEL0)
'-----------------------------------------------------

srvYLIBEL0_Update = "?"

MsgTxtLen = 0
Call srvYLIBEL0_PutBuffer(recYLIBEL0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYLIBEL0_GetBuffer(recYLIBEL0)) Then
        Call srvYLIBEL0_Error(recYLIBEL0)
        srvYLIBEL0_Update = recYLIBEL0.Err
        Exit Function
    Else
        srvYLIBEL0_Update = Null
    End If
Else
    recYLIBEL0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYLIBEL0_Error(recYLIBEL0 As typeYLIBEL0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YLIBEL0" & Chr$(10) & Chr$(13)

Select Case mId$(recYLIBEL0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYLIBEL0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YLIBEL0s.bas  ( " & Trim(recYLIBEL0.obj) & " : " & Trim(recYLIBEL0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYLIBEL0_Monitor(recYLIBEL0 As typeYLIBEL0)
'-----------------------------------------------------

arrYLIBEL0_Suite = False
Select Case mId$(Trim(recYLIBEL0.Method), 1, 4)
    Case "Snap"
              srvYLIBEL0_Monitor = srvYLIBEL0_Snap(recYLIBEL0)
    Case Else
            srvYLIBEL0_Monitor = srvYLIBEL0_Seek(recYLIBEL0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYLIBEL0_GetBuffer(recYLIBEL0 As typeYLIBEL0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYLIBEL0_GetBuffer = Null
recYLIBEL0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYLIBEL0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYLIBEL0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYLIBEL0.Err = Space$(10) Then
    recYLIBEL0.LIBELETA = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYLIBEL0.LIBELPIE = CLng(Val(mId$(MsgTxt, K + 6, 10)))
    recYLIBEL0.LIBELECR = CLng(Val(mId$(MsgTxt, K + 16, 8)))
    recYLIBEL0.LIBELNUM = CLng(Val(mId$(MsgTxt, K + 24, 2)))
    recYLIBEL0.LIBELLIB = mId$(MsgTxt, K + 26, 30)
Else
    srvYLIBEL0_GetBuffer = recYLIBEL0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYLIBEL0Len

End Function

'---------------------------------------------------------
Public Sub srvYLIBEL0_PutBuffer(recYLIBEL0 As typeYLIBEL0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYLIBEL0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYLIBEL0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYLIBEL0.LIBELETA, "0000 ")
    Mid$(MsgTxt, K + 6, 10) = Format$(recYLIBEL0.LIBELPIE, "000000000 ")
    Mid$(MsgTxt, K + 16, 8) = Format$(recYLIBEL0.LIBELECR, "0000000 ")
    Mid$(MsgTxt, K + 24, 2) = Format$(recYLIBEL0.LIBELNUM, "0 ")
    Mid$(MsgTxt, K + 26, 30) = recYLIBEL0.LIBELLIB
MsgTxtLen = MsgTxtLen + recYLIBEL0Len
End Sub



'---------------------------------------------------------
Private Function srvYLIBEL0_Seek(recYLIBEL0 As typeYLIBEL0)
'---------------------------------------------------------

srvYLIBEL0_Seek = "?"
MsgTxtLen = 0
Call srvYLIBEL0_PutBuffer(recYLIBEL0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYLIBEL0_GetBuffer(recYLIBEL0)) Then
        srvYLIBEL0_Seek = Null
    Else
        Call srvYLIBEL0_Error(recYLIBEL0)
    End If
End If

End Function

Public Sub srvYLIBEL0_ElpDisplay(recYLIBEL0 As typeYLIBEL0)
frmElpDisplay.fgData.Rows = 6
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "LIBELETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYLIBEL0.LIBELETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "LIBELPIE    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DE PIECE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYLIBEL0.LIBELPIE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "LIBELECR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO D'ECRITURE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYLIBEL0.LIBELECR
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "LIBELNUM    1S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DE LIBELLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYLIBEL0.LIBELNUM
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "LIBELLIB   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBELLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYLIBEL0.LIBELLIB
frmElpDisplay.Show vbModal
End Sub
Public Sub srvYLIBEL0_Export_CSV()
Dim xIn As String
Open "D:\Temp\FTP\YLIBEL0.txt" For Input As #1
Open "D:\Temp\FTP\YLIBEL0.csv" For Output As #2
Print #2, "LIBELETA;LIBELPIE;LIBELECR;LIBELNUM;LIBELLIB;"
Print #2, "ETABLISSEMENT;NUMERO DE PIECE;NUMERO D'ECRITURE;NUMERO DE LIBELLE;LIBELLE;"
Print #2, ";;;;;"
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 10) & ";" _
      & mId$(xIn, 16, 8) & ";" _
      & mId$(xIn, 24, 2) & ";" _
      & mId$(xIn, 26, 30) & ";"
Loop
Close
End Sub

'---------------------------------------------------------
Private Function srvYLIBEL0_Snap(recYLIBEL0 As typeYLIBEL0)
'---------------------------------------------------------
srvYLIBEL0_Snap = "?"
MsgTxtLen = 0
Call srvYLIBEL0_PutBuffer(recYLIBEL0)
Call srvYLIBEL0_PutBuffer(arrYLIBEL0(0))
If IsNull(SndRcv()) Then
    srvYLIBEL0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYLIBEL0_GetBuffer(recYLIBEL0)) Then
            Call arrYLIBEL0_AddItem(recYLIBEL0)
            arrYLIBEL0_Suite = True
        Else
            arrYLIBEL0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYLIBEL0_AddItem(recYLIBEL0 As typeYLIBEL0)
'---------------------------------------------------------
          
arrYLIBEL0_NB = arrYLIBEL0_NB + 1
    
If arrYLIBEL0_NB > arrYLIBEL0_NBMax Then
    arrYLIBEL0_NBMax = arrYLIBEL0_NBMax + recYLIBEL0_Block
    ReDim Preserve arrYLIBEL0(arrYLIBEL0_NBMax)
End If
            
arrYLIBEL0(arrYLIBEL0_NB) = recYLIBEL0
End Sub



'---------------------------------------------------------
Public Sub recYLIBEL0_Init(recYLIBEL0 As typeYLIBEL0)
'---------------------------------------------------------
recYLIBEL0.obj = "ZLIBEL0_S"
recYLIBEL0.Method = ""
recYLIBEL0.Err = ""
'recYLIBEL0.MOUVEMETA = 1

End Sub








