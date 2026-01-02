Attribute VB_Name = "srvYBASFUT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYBASFUT0Len = 140 ' 34 + 106
Public Const recYBASFUT0_Block = 200
Public Const constYBASFUT0 = "YBASFUT0"
Dim meYbase As typeYBase
Dim paramYBASFUT0_Import As String

Type typeYBASFUT0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    BASFUTETA       As Integer                        ' ETABLISSEMENT
    BASFUTOPE       As String * 3                     ' OPERATION
    BASFUTAGE       As Integer                        ' AGENCE
    BASFUTSER       As String * 2                     ' SERVICE
    BASFUTSSE       As String * 2                     ' SOUS SERVICE
    BASFUTDOS       As Long                           ' DOSSIER
    BASFUTDTE       As Long                           ' DATE EVENEMENT
    BASFUTEVE       As String * 3                     ' EVENEMENT
    BASFUTNUM       As Long                           ' NUMERO EVENEMENT
    BASFUTTYP       As String * 1                     ' TYPE EVENEMENT
    BASFUTNAT       As String * 3                     ' NATURE OPERATION
    BASFUTDVA       As Long                           ' DATE DE VALEUR
    BASFUTMON       As Currency                       ' MONTANT
    BASFUTSEN       As String * 1                     ' SENS OPERATION
    BASFUTDEV       As String * 3                     ' DEVISE
    BASFUTCPT       As String * 20                    ' COMPTE
    BASFUTTCL       As String * 1                     ' CLIENT TIERS
    BASFUTCLI       As String * 7                     ' CONTREPARTIE
    BASFUTTAU       As String * 1                     ' TAUX VARIABLE
    BASFUTNAG       As Integer                        ' AGENCE NETTING
    BASFUTNSE       As String * 2                     ' SERVICE NETTING
    BASFUTNSS       As String * 2                     ' S SERVICE NETTING
    BASFUTNDO       As Long                           ' DOSSIER NETTING
    BASFUTLIB       As String * 30                    ' LIBELLE

End Type
    
Type typeYBASFUT0_Total
    COMPTEDEV       As String * 3                     ' DEVISE
    COMPTECOM       As String * 20                    ' COMPTE
    COMPTEINT       As String * 32
    SOLDEVEN         As Currency
    BASFUTMON(6)     As Currency
    BASFUTDVA_Err    As Boolean
End Type

'---------------------------------------------------------
Public Function srvYBASFUT0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYBASFUT0 As typeYBASFUT0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYBASFUT0_GetBuffer_ODBC = Null

    recYBASFUT0.BASFUTETA = rsADO("BASFUTETA")       'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYBASFUT0.BASFUTOPE = rsADO("BASFUTOPE")       'mId$(MsgTxt, K + 6, 3)
    recYBASFUT0.BASFUTAGE = rsADO("BASFUTAGE")       'CInt(Val(mId$(MsgTxt, K + 9, 5)))
    recYBASFUT0.BASFUTSER = rsADO("BASFUTSER")       'mId$(MsgTxt, K + 14, 2)
    recYBASFUT0.BASFUTSSE = rsADO("BASFUTSSE")       'mId$(MsgTxt, K + 16, 2)
    recYBASFUT0.BASFUTDOS = rsADO("BASFUTDOS")       'CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYBASFUT0.BASFUTDTE = rsADO("BASFUTDTE")       'CLng(Val(mId$(MsgTxt, K + 28, 8)))
    recYBASFUT0.BASFUTEVE = rsADO("BASFUTEVE")       'mId$(MsgTxt, K + 36, 3)
    recYBASFUT0.BASFUTNUM = rsADO("BASFUTNUM")       'CLng(Val(mId$(MsgTxt, K + 39, 4)))
    recYBASFUT0.BASFUTTYP = rsADO("BASFUTTYP")       'mId$(MsgTxt, K + 43, 1)
    recYBASFUT0.BASFUTNAT = rsADO("BASFUTNAT")       'mId$(MsgTxt, K + 44, 3)
    recYBASFUT0.BASFUTDVA = rsADO("BASFUTDVA")       'CLng(Val(mId$(MsgTxt, K + 47, 8)))
    recYBASFUT0.BASFUTMON = rsADO("BASFUTMON")       'CCur(Val(mId$(MsgTxt, K + 55, 16))) / 100
    recYBASFUT0.BASFUTSEN = rsADO("BASFUTSEN")       'mId$(MsgTxt, K + 71, 1)
    recYBASFUT0.BASFUTDEV = rsADO("BASFUTDEV")       'mId$(MsgTxt, K + 72, 3)
    recYBASFUT0.BASFUTCPT = rsADO("BASFUTCPT")       'mId$(MsgTxt, K + 75, 20)
    recYBASFUT0.BASFUTTCL = rsADO("BASFUTTCL")       'mId$(MsgTxt, K + 95, 1)
    recYBASFUT0.BASFUTCLI = rsADO("BASFUTCLI")       'mId$(MsgTxt, K + 96, 7)
    recYBASFUT0.BASFUTTAU = rsADO("BASFUTTAU")       'mId$(MsgTxt, K + 103, 1)
    recYBASFUT0.BASFUTNAG = rsADO("BASFUTNAG")       'CInt(Val(mId$(MsgTxt, K + 104, 5)))
    recYBASFUT0.BASFUTNSE = rsADO("BASFUTNSE")       'mId$(MsgTxt, K + 109, 2)
    recYBASFUT0.BASFUTNSS = rsADO("BASFUTNSS")       'mId$(MsgTxt, K + 111, 2)
    recYBASFUT0.BASFUTNDO = rsADO("BASFUTNDO")       'CLng(Val(mId$(MsgTxt, K + 113, 10)))
    recYBASFUT0.BASFUTLIB = rsADO("BASFUTLIB")       'mId$(MsgTxt, K + 123, 30)

Exit Function

Error_Handler:
srvYBASFUT0_GetBuffer_ODBC = Error

End Function


Public Function srvYBASFUT0_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYBASFUT0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    srvYBASFUT0_Import = Null
    lX = CStr(meYbase.Text)
    Exit Function
End If


srvYBASFUT0_Import = "?"

paramYBASFUT0_Import = paramYBase_DataF & Trim(constYBASFUT0) & paramYBase_Data_ExtensionP

Open Trim(paramYBASFUT0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYBASFUT0) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYBASFUT0
            meYbase.K1 = mId$(xIn, 15, 27) 'recYBASFUT0.CDODOSCOP & recYBASFUT0.CDODOSDOS .........
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYBASFUT0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYBASFUT0
lX = DSys & "_" & time_Hms & "_" & Nb
meYbase.Text = lX
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYBASFUT0_Import" & xIn, vbCritical, Error
Close

srvYBASFUT0_Import = Error
End Function

Public Function srvYBASFUT0_Import_Read(lId As String, lYBASFUT0 As typeYBASFUT0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYBASFUT0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYBASFUT0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYBASFUT0_GetBuffer lYBASFUT0
    srvYBASFUT0_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYBASFUT0_Import_Read" & xIn, vbCritical, Error
srvYBASFUT0_Import_Read = Error
End Function





'-----------------------------------------------------
Public Function srvYBASFUT0_Monitor(recYBASFUT0 As typeYBASFUT0)
'-----------------------------------------------------

Select Case mId$(Trim(recYBASFUT0.Method), 1, 4)
    Case "Seek"
                srvYBASFUT0_Monitor = srvYBASFUT0_Seek(recYBASFUT0)
    Case Else
                recYBASFUT0.Err = recYBASFUT0.Method
                Call srvYBASFUT0_Error(recYBASFUT0)
                srvYBASFUT0_Monitor = recYBASFUT0.Err
End Select

End Function

'-----------------------------------------------------
Sub srvYBASFUT0_Error(recYBASFUT0 As typeYBASFUT0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YBASFUT0" & Chr$(10) & Chr$(13)

Select Case mId$(recYBASFUT0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYBASFUT0.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : YBASFUT0s.bas  ( " _
                & Trim(recYBASFUT0.obj) & " : " & Trim(recYBASFUT0.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvYBASFUT0_GetBuffer(recYBASFUT0 As typeYBASFUT0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYBASFUT0_GetBuffer = Null
recYBASFUT0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYBASFUT0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYBASFUT0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYBASFUT0.Err = Space$(10) Then

    recYBASFUT0.BASFUTETA = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYBASFUT0.BASFUTOPE = mId$(MsgTxt, K + 6, 3)
    recYBASFUT0.BASFUTAGE = CInt(Val(mId$(MsgTxt, K + 9, 5)))
    recYBASFUT0.BASFUTSER = mId$(MsgTxt, K + 14, 2)
    recYBASFUT0.BASFUTSSE = mId$(MsgTxt, K + 16, 2)
    recYBASFUT0.BASFUTDOS = CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYBASFUT0.BASFUTDTE = CLng(Val(mId$(MsgTxt, K + 28, 8)))
    recYBASFUT0.BASFUTEVE = mId$(MsgTxt, K + 36, 3)
    recYBASFUT0.BASFUTNUM = CLng(Val(mId$(MsgTxt, K + 39, 4)))
    recYBASFUT0.BASFUTTYP = mId$(MsgTxt, K + 43, 1)
    recYBASFUT0.BASFUTNAT = mId$(MsgTxt, K + 44, 3)
    recYBASFUT0.BASFUTDVA = CLng(Val(mId$(MsgTxt, K + 47, 8)))
    recYBASFUT0.BASFUTMON = CCur(Val(mId$(MsgTxt, K + 55, 16))) / 100
    recYBASFUT0.BASFUTSEN = mId$(MsgTxt, K + 71, 1)
    recYBASFUT0.BASFUTDEV = mId$(MsgTxt, K + 72, 3)
    recYBASFUT0.BASFUTCPT = mId$(MsgTxt, K + 75, 20)
    recYBASFUT0.BASFUTTCL = mId$(MsgTxt, K + 95, 1)
    recYBASFUT0.BASFUTCLI = mId$(MsgTxt, K + 96, 7)
    recYBASFUT0.BASFUTTAU = mId$(MsgTxt, K + 103, 1)
    recYBASFUT0.BASFUTNAG = CInt(Val(mId$(MsgTxt, K + 104, 5)))
    recYBASFUT0.BASFUTNSE = mId$(MsgTxt, K + 109, 2)
    recYBASFUT0.BASFUTNSS = mId$(MsgTxt, K + 111, 2)
    recYBASFUT0.BASFUTNDO = CLng(Val(mId$(MsgTxt, K + 113, 10)))
    recYBASFUT0.BASFUTLIB = mId$(MsgTxt, K + 123, 30)

Else
    srvYBASFUT0_GetBuffer = recYBASFUT0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYBASFUT0Len

End Function

'---------------------------------------------------------
Private Sub srvYBASFUT0_PutBuffer(recYBASFUT0 As typeYBASFUT0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recYBASFUT0Len) = Space$(recYBASFUT0Len)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYBASFUT0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYBASFUT0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYBASFUT0.BASFUTETA, "0000 ")
    Mid$(MsgTxt, K + 6, 3) = recYBASFUT0.BASFUTOPE
    Mid$(MsgTxt, K + 9, 5) = Format$(recYBASFUT0.BASFUTAGE, "0000 ")
    Mid$(MsgTxt, K + 14, 2) = recYBASFUT0.BASFUTSER
    Mid$(MsgTxt, K + 16, 2) = recYBASFUT0.BASFUTSSE
    Mid$(MsgTxt, K + 18, 10) = Format$(recYBASFUT0.BASFUTDOS, "000000000 ")
    Mid$(MsgTxt, K + 28, 8) = Format$(recYBASFUT0.BASFUTDTE, "0000000 ")
    Mid$(MsgTxt, K + 36, 3) = recYBASFUT0.BASFUTEVE
    Mid$(MsgTxt, K + 39, 4) = Format$(recYBASFUT0.BASFUTNUM, "000 ")
    Mid$(MsgTxt, K + 43, 1) = recYBASFUT0.BASFUTTYP
    Mid$(MsgTxt, K + 44, 3) = recYBASFUT0.BASFUTNAT
    Mid$(MsgTxt, K + 47, 8) = Format$(recYBASFUT0.BASFUTDVA, "0000000 ")
    Mid$(MsgTxt, K + 55, 16) = Format$(recYBASFUT0.BASFUTMON * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 71, 1) = recYBASFUT0.BASFUTSEN
    Mid$(MsgTxt, K + 72, 3) = recYBASFUT0.BASFUTDEV
    Mid$(MsgTxt, K + 75, 20) = recYBASFUT0.BASFUTCPT
    Mid$(MsgTxt, K + 95, 1) = recYBASFUT0.BASFUTTCL
    Mid$(MsgTxt, K + 96, 7) = recYBASFUT0.BASFUTCLI
    Mid$(MsgTxt, K + 103, 1) = recYBASFUT0.BASFUTTAU
    Mid$(MsgTxt, K + 104, 5) = Format$(recYBASFUT0.BASFUTNAG, "0000 ")
    Mid$(MsgTxt, K + 109, 2) = recYBASFUT0.BASFUTNSE
    Mid$(MsgTxt, K + 111, 2) = recYBASFUT0.BASFUTNSS
    Mid$(MsgTxt, K + 113, 10) = Format$(recYBASFUT0.BASFUTNDO, "000000000 ")
    Mid$(MsgTxt, K + 123, 30) = recYBASFUT0.BASFUTLIB

End Sub


'---------------------------------------------------------
Private Function srvYBASFUT0_Seek(recYBASFUT0 As typeYBASFUT0)
'---------------------------------------------------------

srvYBASFUT0_Seek = "?"
MsgTxtLen = 0
Call srvYBASFUT0_PutBuffer(recYBASFUT0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvYBASFUT0_GetBuffer(recYBASFUT0)) Then
            srvYBASFUT0_Seek = Null
        Else
            Call srvYBASFUT0_Error(recYBASFUT0)
        End If
    End If
End If

End Function
Public Sub srvYBASFUT0_ElpDisplay(recYBASFUT0 As typeYBASFUT0)
frmElpDisplay.fgData.Rows = 25
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTOPE    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTOPE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTAGE
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTSER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTSER
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTSSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTSSE
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTDOS    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTDOS
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTDTE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE EVENEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTDTE
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTEVE    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EVENEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTEVE
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTNUM    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO EVENEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTNUM
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTTYP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE EVENEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTTYP
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTNAT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NATURE OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTNAT
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTDVA    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DE VALEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTDVA
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTMON 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTMON
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTSEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SENS OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTSEN
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTDEV    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTDEV
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTCPT   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTCPT
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTTCL    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CLIENT TIERS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTTCL
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTCLI    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CONTREPARTIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTCLI
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTTAU    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX VARIABLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTTAU
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTNAG    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE NETTING"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTNAG
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTNSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE NETTING"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTNSE
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTNSS    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "S SERVICE NETTING"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTNSS
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTNDO    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DOSSIER NETTING"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTNDO
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BASFUTLIB   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBELLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBASFUT0.BASFUTLIB
frmElpDisplay.Show vbModal
End Sub

'-----------------------------------------------------
Function srvYBASFUT0_Update(recYBASFUT0 As typeYBASFUT0)
'-----------------------------------------------------

srvYBASFUT0_Update = "?"

MsgTxtLen = 0
Call srvYBASFUT0_PutBuffer(recYBASFUT0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYBASFUT0_GetBuffer(recYBASFUT0)) Then
        Call srvYBASFUT0_Error(recYBASFUT0)
        srvYBASFUT0_Update = recYBASFUT0.Err
        Exit Function
    Else
        srvYBASFUT0_Update = Null
    End If
Else
    recYBASFUT0.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recYBASFUT0_Init(recYBASFUT0 As typeYBASFUT0)
'---------------------------------------------------------
MsgTxt = Space$(recYBASFUT0Len)
MsgTxtIndex = 0
Call srvYBASFUT0_GetBuffer(recYBASFUT0)
recYBASFUT0.obj = "ZCDODOS0_S"
recYBASFUT0.BASFUTETA = 0      'AsInteger                        ' ETABLISSEMENT
recYBASFUT0.BASFUTOPE = ""      'As String * 3                     ' OPERATION
recYBASFUT0.BASFUTAGE = 0      'AsInteger                        ' AGENCE
recYBASFUT0.BASFUTSER = ""      'As String * 2                     ' SERVICE
recYBASFUT0.BASFUTSSE = ""      'As String * 2                     ' SOUS SERVICE
recYBASFUT0.BASFUTDOS = 0      'AsLong                           ' DOSSIER
recYBASFUT0.BASFUTDTE = 0      'AsLong                           ' DATE EVENEMENT
recYBASFUT0.BASFUTEVE = ""      'As String * 3                     ' EVENEMENT
recYBASFUT0.BASFUTNUM = 0      'AsLong                           ' NUMERO EVENEMENT
recYBASFUT0.BASFUTTYP = ""      'As String * 1                     ' TYPE EVENEMENT
recYBASFUT0.BASFUTNAT = ""      'As String * 3                     ' NATURE OPERATION
recYBASFUT0.BASFUTDVA = 0      'AsLong                           ' DATE DE VALEUR
recYBASFUT0.BASFUTMON = 0      'AsCurrency                       ' MONTANT
recYBASFUT0.BASFUTSEN = ""      'As String * 1                     ' SENS OPERATION
recYBASFUT0.BASFUTDEV = ""      'As String * 3                     ' DEVISE
recYBASFUT0.BASFUTCPT = ""      'As String * 20                    ' COMPTE
recYBASFUT0.BASFUTTCL = ""      'As String * 1                     ' CLIENT TIERS
recYBASFUT0.BASFUTCLI = ""      'As String * 7                     ' CONTREPARTIE
recYBASFUT0.BASFUTTAU = ""      'As String * 1                     ' TAUX VARIABLE
recYBASFUT0.BASFUTNAG = 0      'AsInteger                        ' AGENCE NETTING
recYBASFUT0.BASFUTNSE = ""      'As String * 2                     ' SERVICE NETTING
recYBASFUT0.BASFUTNSS = ""      'As String * 2                     ' S SERVICE NETTING
recYBASFUT0.BASFUTNDO = 0      'AsLong                           ' DOSSIER NETTING
recYBASFUT0.BASFUTLIB = ""      'As String * 30                    ' LIBELLE
End Sub
'---------------------------------------------------------
Public Sub recYBASFUT0_Total_Init(recYBASFUT0_Total As typeYBASFUT0_Total)
'---------------------------------------------------------
Dim K As Integer
recYBASFUT0_Total.COMPTEDEV = ""      'As String * 3                     ' DEVISE
recYBASFUT0_Total.COMPTECOM = ""      'As String * 20                    ' COMPTE
recYBASFUT0_Total.SOLDEVEN = 0
For K = 0 To 6
    recYBASFUT0_Total.BASFUTMON(K) = 0
Next K
recYBASFUT0_Total.BASFUTDVA_Err = False
End Sub








