Attribute VB_Name = "srvYRELEVE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYRELEVE0Len = 103 ' 34 +69
Public Const recYRELEVE0_Block = 200
Public Const memoYRELEVE0Len = 69
Public Const constYRELEVE0 = "YRELEVE0  "
Public paramYRELEVE0_Import As String

Type typeYRELEVE0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    RELEVEETA       As Integer                        ' ETABLISSEMENT
    RELEVEPLA       As Long                           ' NUMERO PLAN
    RELEVECOM       As String * 20                    ' NUMERO COMPTE
    RELEVEREL       As String * 1                     ' TABLES BASE 019
    RELEVETYP       As String * 1                     ' 1 client , 2 compte
    RELEVENUM       As String * 20                    ' N° Client ou Compte
    RELEVEADR       As String * 2                     ' CODE ADRESSE
    RELEVEGES       As String * 1                     ' RELEVE GESTIONNAIRE
    RELEVEDER       As Long                           ' DATE DERNIER RELEVE
    RELEVEEXT       As Long                           ' NUMERO D'EXTRAIT
End Type
    
    
Public arrYRELEVE0() As typeYRELEVE0
Public arrYRELEVE0_NB As Integer
Public arrYRELEVE0_NBMax As Integer
Public arrYRELEVE0_Index As Integer
Public arrYRELEVE0_Suite As Boolean

Dim meMVTP0 As typeMvtP0

'-----------------------------------------------------
Function srvYRELEVE0_Update(recYRELEVE0 As typeYRELEVE0)
'-----------------------------------------------------

srvYRELEVE0_Update = "?"

MsgTxtLen = 0
Call srvYRELEVE0_PutBuffer(recYRELEVE0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYRELEVE0_GetBuffer(recYRELEVE0)) Then
        Call srvYRELEVE0_Error(recYRELEVE0)
        srvYRELEVE0_Update = recYRELEVE0.Err
        Exit Function
    Else
        srvYRELEVE0_Update = Null
    End If
Else
    recYRELEVE0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYRELEVE0_Error(recYRELEVE0 As typeYRELEVE0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YRELEVE0" & Chr$(10) & Chr$(13)

Select Case mId$(recYRELEVE0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYRELEVE0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YRELEVE0s.bas  ( " & Trim(recYRELEVE0.Obj) & " : " & Trim(recYRELEVE0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYRELEVE0_Monitor(recYRELEVE0 As typeYRELEVE0)
'-----------------------------------------------------

arrYRELEVE0_Suite = False
Select Case mId$(Trim(recYRELEVE0.Method), 1, 4)
    Case "Snap"
              srvYRELEVE0_Monitor = srvYRELEVE0_Snap(recYRELEVE0)
    Case Else
            srvYRELEVE0_Monitor = srvYRELEVE0_Seek(recYRELEVE0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYRELEVE0_GetBuffer(recYRELEVE0 As typeYRELEVE0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYRELEVE0_GetBuffer = Null
recYRELEVE0.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYRELEVE0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYRELEVE0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYRELEVE0.Err = Space$(10) Then
    recYRELEVE0.RELEVEETA = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYRELEVE0.RELEVEPLA = CLng(Val(mId$(MsgTxt, K + 6, 4)))
    recYRELEVE0.RELEVECOM = mId$(MsgTxt, K + 10, 20)
    recYRELEVE0.RELEVEREL = mId$(MsgTxt, K + 30, 1)
    recYRELEVE0.RELEVETYP = mId$(MsgTxt, K + 31, 1)
    recYRELEVE0.RELEVENUM = mId$(MsgTxt, K + 32, 20)
    recYRELEVE0.RELEVEADR = mId$(MsgTxt, K + 52, 2)
    recYRELEVE0.RELEVEGES = mId$(MsgTxt, K + 54, 1)
    recYRELEVE0.RELEVEDER = CLng(Val(mId$(MsgTxt, K + 55, 8)))
    recYRELEVE0.RELEVEEXT = CLng(Val(mId$(MsgTxt, K + 63, 7)))
Else
    srvYRELEVE0_GetBuffer = recYRELEVE0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYRELEVE0Len

End Function

'---------------------------------------------------------
Public Sub srvYRELEVE0_PutBuffer(recYRELEVE0 As typeYRELEVE0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYRELEVE0.Obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYRELEVE0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
    Mid$(MsgTxt, K + 1, 5) = Format$(recYRELEVE0.RELEVEETA, "0000 ")
    Mid$(MsgTxt, K + 6, 4) = Format$(recYRELEVE0.RELEVEPLA, "000 ")
    Mid$(MsgTxt, K + 10, 20) = recYRELEVE0.RELEVECOM
    Mid$(MsgTxt, K + 30, 1) = recYRELEVE0.RELEVEREL
    Mid$(MsgTxt, K + 31, 1) = recYRELEVE0.RELEVETYP
    Mid$(MsgTxt, K + 32, 20) = recYRELEVE0.RELEVENUM
    Mid$(MsgTxt, K + 52, 2) = recYRELEVE0.RELEVEADR
    Mid$(MsgTxt, K + 54, 1) = recYRELEVE0.RELEVEGES
    Mid$(MsgTxt, K + 55, 8) = Format$(recYRELEVE0.RELEVEDER, "0000000 ")
    Mid$(MsgTxt, K + 63, 7) = Format$(recYRELEVE0.RELEVEEXT, "000000 ")
MsgTxtLen = MsgTxtLen + recYRELEVE0Len
End Sub


'---------------------------------------------------------
Private Function srvYRELEVE0_Seek(recYRELEVE0 As typeYRELEVE0)
'---------------------------------------------------------

srvYRELEVE0_Seek = "?"
MsgTxtLen = 0
Call srvYRELEVE0_PutBuffer(recYRELEVE0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYRELEVE0_GetBuffer(recYRELEVE0)) Then
        srvYRELEVE0_Seek = Null
    Else
        Call srvYRELEVE0_Error(recYRELEVE0)
    End If
End If

End Function

Public Sub srvYRELEVE0_ElpDisplay(recYRELEVE0 As typeYRELEVE0)
frmElpDisplay.fgData.Rows = 11
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "RELEVEETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYRELEVE0.RELEVEETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "RELEVEPLA    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PLAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYRELEVE0.RELEVEPLA
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "RELEVECOM   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYRELEVE0.RELEVECOM
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "RELEVEREL    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TABLES BASE 019"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYRELEVE0.RELEVEREL
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "RELEVETYP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "1 client , 2 compte            blanc pas adresse"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYRELEVE0.RELEVETYP
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "RELEVENUM   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° Client ou Compte"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYRELEVE0.RELEVENUM
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "RELEVEADR    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ADRESSE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYRELEVE0.RELEVEADR
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "RELEVEGES    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RELEVE GESTIONNAIRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYRELEVE0.RELEVEGES
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "RELEVEDER    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DERNIER RELEVE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYRELEVE0.RELEVEDER
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "RELEVEEXT    6P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO D'EXTRAIT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYRELEVE0.RELEVEEXT
frmElpDisplay.Show vbModal
End Sub
Public Sub srvYRELEVE0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YRELEVE0.txt" For Input As #1
Open "C:\Temp\YRELEVE0.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "RELEVEETA;RELEVEPLA;RELEVECOM;RELEVEREL;RELEVETYP;RELEVENUM;RELEVEADR;RELEVEGES;RELEVEDER;RELEVEEXT;"
    Print #2, "ETABLISSEMENT;NUMERO PLAN;NUMERO COMPTE;CODE RELEVE;Type de numéro;N° Client ou Compte;CODE ADRESSE;RELEVE GESTIONNAIRE;DATE DERNIER RELEVE;NUMERO D'EXTRAIT;"
    Print #2, ";;;;;;;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 4) & ";" _
      & mId$(xIn, 10, 20) & ";" _
      & mId$(xIn, 30, 1) & ";" _
      & mId$(xIn, 31, 1) & ";" _
      & mId$(xIn, 32, 20) & ";" _
      & mId$(xIn, 52, 2) & ";" _
      & mId$(xIn, 54, 1) & ";" _
      & mId$(xIn, 55, 8) & ";" _
      & mId$(xIn, 63, 7) & ";"
Loop
Close
End Sub

Public Function srvYRELEVE0_Import(lnb As Long)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYRELEVE0_Import = "?"

paramYRELEVE0_Import = paramYBase_DataF & Trim(constYRELEVE0) & paramYBase_Data_ExtensionP

Open Trim(paramYRELEVE0_Import) For Input As #1

lnb = 0

recMvtP0_Init meMVTP0
meMVTP0.Method = constAddNew

mdbMvtP0.tableMvtP0_Open

Do Until EOF(1)
    lnb = lnb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meMVTP0.ID = constYRELEVE0 & mId$(xIn, 10, 21) ' compte  & type
            meMVTP0.Text = xIn
            dbMvtP0_Update meMVTP0
            
    End If
        
Loop


Close
srvYRELEVE0_Import = Null
Exit Function

Error_Handle:
 MsgBox "erreur : srvYRELEVE0_Import" & xIn, vbCritical, Error
Close

srvYRELEVE0_Import = Error
End Function
Public Function srvYRELEVE0_Import_Read(lId As String, lYRELEVE0 As typeYRELEVE0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYRELEVE0_Import_Read = "?"

meMVTP0.Method = "Seek>="

meMVTP0.ID = lId                         ' compte  & type
If tableMvtP0_Read(meMVTP0) = 0 Then
    If mId$(lId, 1, 30) = mId$(meMVTP0.ID, 1, 30) Then
    
        MsgTxt = Space$(34) & meMVTP0.Text
        MsgTxtIndex = 0
        srvYRELEVE0_GetBuffer lYRELEVE0
        srvYRELEVE0_Import_Read = Null
    End If
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYRELEVE0_Import_Read" & xIn, vbCritical, Error
Close
srvYRELEVE0_Import_Read = Error
End Function

'---------------------------------------------------------
Private Function srvYRELEVE0_Snap(recYRELEVE0 As typeYRELEVE0)
'---------------------------------------------------------
srvYRELEVE0_Snap = "?"
MsgTxtLen = 0
Call srvYRELEVE0_PutBuffer(recYRELEVE0)
Call srvYRELEVE0_PutBuffer(arrYRELEVE0(0))
If IsNull(SndRcv()) Then
    srvYRELEVE0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYRELEVE0_GetBuffer(recYRELEVE0)) Then
            Call arrYRELEVE0_AddItem(recYRELEVE0)
            arrYRELEVE0_Suite = True
        Else
            arrYRELEVE0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYRELEVE0_AddItem(recYRELEVE0 As typeYRELEVE0)
'---------------------------------------------------------
          
arrYRELEVE0_NB = arrYRELEVE0_NB + 1
    
If arrYRELEVE0_NB > arrYRELEVE0_NBMax Then
    arrYRELEVE0_NBMax = arrYRELEVE0_NBMax + recYRELEVE0_Block
    ReDim Preserve arrYRELEVE0(arrYRELEVE0_NBMax)
End If
            
arrYRELEVE0(arrYRELEVE0_NB) = recYRELEVE0
End Sub



'---------------------------------------------------------
Public Sub recYRELEVE0_Init(recYRELEVE0 As typeYRELEVE0)
'---------------------------------------------------------
recYRELEVE0.Obj = "ZCOMEXP0_S"
recYRELEVE0.Method = ""
recYRELEVE0.Err = ""
recYRELEVE0.RELEVEETA = 1
recYRELEVE0.RELEVEPLA = 0
recYRELEVE0.RELEVECOM = ""
recYRELEVE0.RELEVEREL = ""
recYRELEVE0.RELEVETYP = ""
recYRELEVE0.RELEVENUM = ""
recYRELEVE0.RELEVEADR = ""
recYRELEVE0.RELEVEGES = ""
recYRELEVE0.RELEVEDER = 0
recYRELEVE0.RELEVEEXT = 0

End Sub






