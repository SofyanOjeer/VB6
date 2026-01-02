Attribute VB_Name = "srvYTITULA0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYTITULA0Len = 72 ' 34 +38
Public Const recYTITULA0_Block = 100
Public Const memoYTITULA0Len = 38
Public Const constYTITULA0 = "YTITULA0"
Public paramYTITULA0_Import As String
Dim meYbase As typeYBase

Type typeYTITULA0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    TITULAETA       As Integer                        ' ETABLISSEMENT
    TITULAPLA       As Long                           ' NUMERO PLAN
    TITULACOM       As String * 20                    ' NUMERO COMPTE
    TITULACLI       As String * 7                     ' NUMERO CLIENT
    TITULAPRI       As String * 1                     ' 0:PRINCIPAL, 1:AUTRE
    TITULATPR       As String * 1                     ' 0:PRINCIPAL, 1:AUTRE
End Type
    
    
Public arrYTITULA0() As typeYTITULA0
Public arrYTITULA0_NB As Integer
Public arrYTITULA0_NBMax As Integer
Public arrYTITULA0_Index As Integer
Public arrYTITULA0_Suite As Boolean

'---------------------------------------------------------
Public Function srvYTITULA0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYTITULA0 As typeYTITULA0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYTITULA0_GetBuffer_ODBC = Null

    recYTITULA0.TITULAETA = rsADO("TITULAETA")    '
    recYTITULA0.TITULAPLA = rsADO("TITULAPLA")
    recYTITULA0.TITULACOM = rsADO("TITULACOM")
    recYTITULA0.TITULACLI = rsADO("TITULACLI")
    recYTITULA0.TITULAPRI = rsADO("TITULAPRI")
    recYTITULA0.TITULATPR = rsADO("TITULATPR")


Exit Function

Error_Handler:
srvYTITULA0_GetBuffer_ODBC = Error

End Function


Public Sub srvYTITULA0_ElpDisplay(recYTITULA0 As typeYTITULA0)
frmElpDisplay.fgData.Rows = 7
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TITULAETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTITULA0.TITULAETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TITULAPLA    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PLAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTITULA0.TITULAPLA
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TITULACOM   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTITULA0.TITULACOM
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TITULACLI    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO CLIENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTITULA0.TITULACLI
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TITULAPRI    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "0:PRINCIPAL, 1:AUTRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTITULA0.TITULAPRI
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TITULATPR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "0:PRINCIPAL, 1:AUTRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTITULA0.TITULATPR
frmElpDisplay.Show vbModal
End Sub
Public Sub srvYTITULA0_Export_CSV(lIdFile_Source As Integer, lIdFile_Destination As Integer, loptSelect_CSV_Header As Boolean, lNb As Long)
Dim xIn As String
If loptSelect_CSV_Header Then
    Print #2, "TITULAETA;TITULAPLA;TITULACOM;TITULACLI;TITULAPRI;TITULATPR;"
    Print #2, "ETABLISSEMENT;NUMERO PLAN;NUMERO COMPTE;NUMERO CLIENT;COMPTE PRINCIPAL;TITULAIRE PRINCIPAL;"
    Print #2, ";;;;;;"
End If
Do Until EOF(lIdFile_Source)
      Line Input #lIdFile_Source, xIn
      lNb = lNb + 1
      Print #lIdFile_Destination, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 4) & ";" _
      & mId$(xIn, 10, 20) & ";" _
      & mId$(xIn, 30, 7) & ";" _
      & mId$(xIn, 37, 1) & ";" _
      & mId$(xIn, 38, 1) & ";"
Loop
End Sub

Public Function srvYTITULA0_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle


recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = constYTITULA0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    lX = meYbase.Text
    If mId$(lX, 1, 8) >= YBIATAB0_DATE_CPT_J Then
        srvYTITULA0_Import = Null
        Exit Function
    Else
        meYbase.Method = constDelete
        Call tableYBase_Update(meYbase)
    End If
End If




srvYTITULA0_Import = "?"

paramYTITULA0_Import = paramYBase_DataF & Trim(constYTITULA0) & paramYBase_Data_ExtensionP

Open Trim(paramYTITULA0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYTITULA0) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYTITULA0
            meYbase.K1 = mId$(xIn, 10, 20) & mId$(xIn, 37, 1) & mId$(xIn, 30, 7)  ' .TITULACOM .TITULATPR .TITULACLI
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYTITULA0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = constYTITULA0
meYbase.Text = YBIATAB0_DATE_CPT_J & "_" & DSys & "_" & time_Hms & "_" & Format$(Nb, "000000000")
lX = meYbase.Text
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYTITULA0_Import" & xIn, vbCritical, Error
Close

srvYTITULA0_Import = Error
End Function

Public Function srvYTITULA0_Import_Read(lId As String, lYTITULA0 As typeYTITULA0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYTITULA0_Import_Read = "?"

meYbase.Method = "Seek>="
meYbase.ID = constYTITULA0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    If Trim(mId$(meYbase.K1, 1, 20)) = Trim(lId) Then
        MsgTxt = Space$(34) & meYbase.Text
        MsgTxtIndex = 0
        srvYTITULA0_GetBuffer lYTITULA0
        srvYTITULA0_Import_Read = Null
    End If
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYTITULA0_Import_Read" & xIn, vbCritical, Error
srvYTITULA0_Import_Read = Error
End Function


'-----------------------------------------------------
Function srvYTITULA0_Update(recYTITULA0 As typeYTITULA0)
'-----------------------------------------------------

srvYTITULA0_Update = "?"

MsgTxtLen = 0
Call srvYTITULA0_PutBuffer(recYTITULA0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYTITULA0_GetBuffer(recYTITULA0)) Then
        Call srvYTITULA0_Error(recYTITULA0)
        srvYTITULA0_Update = recYTITULA0.Err
        Exit Function
    Else
        srvYTITULA0_Update = Null
    End If
Else
    recYTITULA0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYTITULA0_Error(recYTITULA0 As typeYTITULA0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YTITULA0" & Chr$(10) & Chr$(13)

Select Case mId$(recYTITULA0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYTITULA0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YTITULA0s.bas  ( " & Trim(recYTITULA0.Obj) & " : " & Trim(recYTITULA0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYTITULA0_Monitor(recYTITULA0 As typeYTITULA0)
'-----------------------------------------------------

arrYTITULA0_Suite = False
Select Case mId$(Trim(recYTITULA0.Method), 1, 4)
    Case "Snap"
              srvYTITULA0_Monitor = srvYTITULA0_Snap(recYTITULA0)
    Case Else
            srvYTITULA0_Monitor = srvYTITULA0_Seek(recYTITULA0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYTITULA0_GetBuffer(recYTITULA0 As typeYTITULA0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYTITULA0_GetBuffer = Null
recYTITULA0.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYTITULA0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYTITULA0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYTITULA0.Err = Space$(10) Then
    recYTITULA0.TITULAETA = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYTITULA0.TITULAPLA = CLng(Val(mId$(MsgTxt, K + 6, 4)))
    recYTITULA0.TITULACOM = mId$(MsgTxt, K + 10, 20)
    recYTITULA0.TITULACLI = mId$(MsgTxt, K + 30, 7)
    recYTITULA0.TITULAPRI = mId$(MsgTxt, K + 37, 1)
    recYTITULA0.TITULATPR = mId$(MsgTxt, K + 38, 1)
Else
    srvYTITULA0_GetBuffer = recYTITULA0.Err
    recYTITULA0.TITULACLI = "?"
End If
MsgTxtIndex = MsgTxtIndex + recYTITULA0Len

End Function

'---------------------------------------------------------
Public Sub srvYTITULA0_PutBuffer(recYTITULA0 As typeYTITULA0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYTITULA0.Obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYTITULA0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYTITULA0.TITULAETA, "0000 ")
    Mid$(MsgTxt, K + 6, 4) = Format$(recYTITULA0.TITULAPLA, "000 ")
    Mid$(MsgTxt, K + 10, 20) = recYTITULA0.TITULACOM
    Mid$(MsgTxt, K + 30, 7) = recYTITULA0.TITULACLI
    Mid$(MsgTxt, K + 37, 1) = recYTITULA0.TITULAPRI
    Mid$(MsgTxt, K + 38, 1) = recYTITULA0.TITULATPR
MsgTxtLen = MsgTxtLen + recYTITULA0Len
End Sub



'---------------------------------------------------------
Private Function srvYTITULA0_Seek(recYTITULA0 As typeYTITULA0)
'---------------------------------------------------------

srvYTITULA0_Seek = "?"
MsgTxtLen = 0
Call srvYTITULA0_PutBuffer(recYTITULA0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYTITULA0_GetBuffer(recYTITULA0)) Then
        srvYTITULA0_Seek = Null
    Else
        Call srvYTITULA0_Error(recYTITULA0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYTITULA0_Snap(recYTITULA0 As typeYTITULA0)
'---------------------------------------------------------
srvYTITULA0_Snap = "?"
MsgTxtLen = 0
Call srvYTITULA0_PutBuffer(recYTITULA0)
Call srvYTITULA0_PutBuffer(arrYTITULA0(0))
If IsNull(SndRcv()) Then
    srvYTITULA0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYTITULA0_GetBuffer(recYTITULA0)) Then
            Call arrYTITULA0_AddItem(recYTITULA0)
            arrYTITULA0_Suite = True
        Else
            arrYTITULA0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYTITULA0_AddItem(recYTITULA0 As typeYTITULA0)
'---------------------------------------------------------
          
arrYTITULA0_NB = arrYTITULA0_NB + 1
    
If arrYTITULA0_NB > arrYTITULA0_NBMax Then
    arrYTITULA0_NBMax = arrYTITULA0_NBMax + recYTITULA0_Block
    ReDim Preserve arrYTITULA0(arrYTITULA0_NBMax)
End If
            
arrYTITULA0(arrYTITULA0_NB) = recYTITULA0
End Sub



'---------------------------------------------------------
Public Sub recYTITULA0_Init(recYTITULA0 As typeYTITULA0)
'---------------------------------------------------------
recYTITULA0.Obj = "ZCLIREF0_S"
recYTITULA0.Method = ""
recYTITULA0.Err = ""

End Sub









