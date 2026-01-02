Attribute VB_Name = "srvYCDOCOM0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCDOCOM0Len = 221 ' 34 + 187
Public Const recYCDOCOM0_Block = 200
Public Const constYCDOCOM0 = "YCDOCOM0"
Dim meYbase As typeYBase
Dim paramYCDOCOM0_Import As String

Type typeYCDOCOM0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    CDOCOMETB       As Integer                        ' CODE ETABLISSEMENT
    CDOCOMAGE       As Integer                        ' AGENCE
    CDOCOMSER       As String * 2                     ' SERVICE
    CDOCOMSSE       As String * 2                     ' SOUS-SERVICE
    CDOCOMCOP       As String * 3                     ' CODE OPERATION
    CDOCOMDOS       As Long                           ' NUMERO DOSSIER
    CDOCOMNUR       As Long                           ' N° RENOUVELLEMENT
    CDOCOMUTI       As Long                           ' N° UTILILSAT°./MODIF
    CDOCOMEVE       As String * 2                     ' EVENEMENT
    CDOCOMSEQ       As Long                           ' N° SEQUENCE
    CDOCOMCOM       As String * 6                     ' CODE COMMISSION
    CDOCOMDEM       As Long                           ' DT DEMANDE
    CDOCOMREG       As Long                           ' DT REGLEMENT
    CDOCOMCPT       As String * 20                    ' NUMERO DU COMPTE
    CDOCOMDEV       As String * 3                     ' DEVISE COMMISSION
    CDOCOMVAL       As Long                           ' DATE VALEUR
    CDOCOMCOU       As Double                         ' COURS DEVCOM/DEVCPT
    CDOCOMMRE       As String * 3                     ' MODE DE REGLEMENT
    CDOCOMBEN       As String * 1                     ' BENEFICIAIRE O/N
    CDOCOMMON       As Currency                       ' MONTANT COMMISSION
    CDOCOMMTV       As Currency                       ' MONTANT TVA
    CDOCOMAVI       As String * 1                     ' 1 NON,2 A EDIT,3 EDI
    CDOCOMPRO       As String * 1                     ' A PROVISIONNER (O/N)
    CDOCOMUTR       As Long                           ' UTILISATION DU REGLE
    CDOCOMNRE       As Long                           ' N° REGLEMENT
    CDOCOMETA       As String * 2                     ' ETAT
    CDOCOMSPE       As Long                           ' N°SEQUENCE PERIODIQ
    CDOCOMDBP       As Long                           ' DATE DEBUT PERIODE
    CDOCOMFNP       As Long                           ' DATE FIN PERIODE
    CDOCOMCUT       As Integer                        ' UTILISATEUR SAISIE
    CDOCOMCER       As String * 1                     ' COTAT°(O=CERTAIN/N)
    
End Type
    
'---------------------------------------------------------
Public Function srvYCDOCOM0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCDOCOM0 As typeYCDOCOM0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCDOCOM0_GetBuffer_ODBC = Null

    recYCDOCOM0.CDOCOMETB = rsADO("CDOCOMETB")    'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOCOM0.CDOCOMAGE = rsADO("CDOCOMAGE")    'CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDOCOM0.CDOCOMSER = rsADO("CDOCOMSER")    'mId$(MsgTxt, K + 11, 2)
    recYCDOCOM0.CDOCOMSSE = rsADO("CDOCOMSSE")    'mId$(MsgTxt, K + 13, 2)
    recYCDOCOM0.CDOCOMCOP = rsADO("CDOCOMCOP")    'mId$(MsgTxt, K + 15, 3)
    recYCDOCOM0.CDOCOMDOS = rsADO("CDOCOMDOS")    'CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDOCOM0.CDOCOMNUR = rsADO("CDOCOMNUR")    'CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDOCOM0.CDOCOMUTI = rsADO("CDOCOMUTI")    'CLng(Val(mId$(MsgTxt, K + 32, 6)))
    recYCDOCOM0.CDOCOMEVE = rsADO("CDOCOMEVE")    'mId$(MsgTxt, K + 38, 2)
    recYCDOCOM0.CDOCOMSEQ = rsADO("CDOCOMSEQ")    'CLng(Val(mId$(MsgTxt, K + 40, 4)))
    recYCDOCOM0.CDOCOMCOM = rsADO("CDOCOMCOM")    'mId$(MsgTxt, K + 44, 6)
    recYCDOCOM0.CDOCOMDEM = rsADO("CDOCOMDEM")    'CLng(Val(mId$(MsgTxt, K + 50, 8)))
    recYCDOCOM0.CDOCOMREG = rsADO("CDOCOMREG")    'CLng(Val(mId$(MsgTxt, K + 58, 8)))
    recYCDOCOM0.CDOCOMCPT = rsADO("CDOCOMCPT")    'mId$(MsgTxt, K + 66, 20)
    recYCDOCOM0.CDOCOMDEV = rsADO("CDOCOMDEV")    'mId$(MsgTxt, K + 86, 3)
    recYCDOCOM0.CDOCOMVAL = rsADO("CDOCOMVAL")    'CLng(Val(mId$(MsgTxt, K + 89, 8)))
    recYCDOCOM0.CDOCOMCOU = rsADO("CDOCOMCOU")    'CDbl(Val(mId$(MsgTxt, K + 97, 15))) / 1000000000
    recYCDOCOM0.CDOCOMMRE = rsADO("CDOCOMMRE")    'mId$(MsgTxt, K + 112, 3)
    recYCDOCOM0.CDOCOMBEN = rsADO("CDOCOMBEN")    'mId$(MsgTxt, K + 115, 1)
    recYCDOCOM0.CDOCOMMON = rsADO("CDOCOMMON")    'CCur(Val(mId$(MsgTxt, K + 116, 16))) / 100
    recYCDOCOM0.CDOCOMMTV = rsADO("CDOCOMMTV")    'CCur(Val(mId$(MsgTxt, K + 132, 16))) / 100
    recYCDOCOM0.CDOCOMAVI = rsADO("CDOCOMAVI")    'mId$(MsgTxt, K + 148, 1)
    recYCDOCOM0.CDOCOMPRO = rsADO("CDOCOMPRO")    'mId$(MsgTxt, K + 149, 1)
    recYCDOCOM0.CDOCOMUTR = rsADO("CDOCOMUTR")    'CLng(Val(mId$(MsgTxt, K + 150, 6)))
    recYCDOCOM0.CDOCOMNRE = rsADO("CDOCOMNRE")    'CLng(Val(mId$(MsgTxt, K + 156, 4)))
    recYCDOCOM0.CDOCOMETA = rsADO("CDOCOMETA")    'mId$(MsgTxt, K + 160, 2)
    recYCDOCOM0.CDOCOMSPE = rsADO("CDOCOMSPE")    'CLng(Val(mId$(MsgTxt, K + 162, 4)))
    recYCDOCOM0.CDOCOMDBP = rsADO("CDOCOMDBP")    'CLng(Val(mId$(MsgTxt, K + 166, 8)))
    recYCDOCOM0.CDOCOMFNP = rsADO("CDOCOMFNP")    'CLng(Val(mId$(MsgTxt, K + 174, 8)))
    recYCDOCOM0.CDOCOMCUT = rsADO("CDOCOMCUT")    'CInt(Val(mId$(MsgTxt, K + 182, 5)))
    recYCDOCOM0.CDOCOMCER = rsADO("CDOCOMCER")    'mId$(MsgTxt, K + 187, 1)

Exit Function

Error_Handler:
srvYCDOCOM0_GetBuffer_ODBC = Error

End Function



Public Function srvYCDOCOM0_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOCOM0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    srvYCDOCOM0_Import = Null
    lX = CStr(meYbase.Text)
    Exit Function
End If


srvYCDOCOM0_Import = "?"

paramYCDOCOM0_Import = paramYBase_DataF & Trim(constYCDOCOM0) & paramYBase_Data_ExtensionP

Open Trim(paramYCDOCOM0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYCDOCOM0) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYCDOCOM0
            meYbase.K1 = mId$(xIn, 15, 29) & mId$(xIn, 162, 4) 'recYCDOCOM0.CDODOSCOP & recYCDOCOM0.CDODOSDOS .........
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYCDOCOM0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOCOM0
lX = DSys & "_" & time_Hms & "_" & Nb
meYbase.Text = lX
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOCOM0_Import" & xIn, vbCritical, Error
Close

srvYCDOCOM0_Import = Error
End Function

Public Function srvYCDOCOM0_Import_Read(lId As String, lYCDOCOM0 As typeYCDOCOM0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYCDOCOM0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYCDOCOM0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYCDOCOM0_GetBuffer lYCDOCOM0
    srvYCDOCOM0_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOCOM0_Import_Read" & xIn, vbCritical, Error
srvYCDOCOM0_Import_Read = Error
End Function





'-----------------------------------------------------
Public Function srvYCDOCOM0_Monitor(recYCDOCOM0 As typeYCDOCOM0)
'-----------------------------------------------------

Select Case mId$(Trim(recYCDOCOM0.Method), 1, 4)
    Case "Seek"
                srvYCDOCOM0_Monitor = srvYCDOCOM0_Seek(recYCDOCOM0)
    Case Else
                recYCDOCOM0.Err = recYCDOCOM0.Method
                Call srvYCDOCOM0_Error(recYCDOCOM0)
                srvYCDOCOM0_Monitor = recYCDOCOM0.Err
End Select

End Function

'-----------------------------------------------------
Sub srvYCDOCOM0_Error(recYCDOCOM0 As typeYCDOCOM0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YCDOCOM0" & Chr$(10) & Chr$(13)

Select Case mId$(recYCDOCOM0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYCDOCOM0.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : YCDOCOM0s.bas  ( " _
                & Trim(recYCDOCOM0.Obj) & " : " & Trim(recYCDOCOM0.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvYCDOCOM0_GetBuffer(recYCDOCOM0 As typeYCDOCOM0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYCDOCOM0_GetBuffer = Null
recYCDOCOM0.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYCDOCOM0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYCDOCOM0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYCDOCOM0.Err = Space$(10) Then
    recYCDOCOM0.CDOCOMETB = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOCOM0.CDOCOMAGE = CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDOCOM0.CDOCOMSER = mId$(MsgTxt, K + 11, 2)
    recYCDOCOM0.CDOCOMSSE = mId$(MsgTxt, K + 13, 2)
    recYCDOCOM0.CDOCOMCOP = mId$(MsgTxt, K + 15, 3)
    recYCDOCOM0.CDOCOMDOS = CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDOCOM0.CDOCOMNUR = CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDOCOM0.CDOCOMUTI = CLng(Val(mId$(MsgTxt, K + 32, 6)))
    recYCDOCOM0.CDOCOMEVE = mId$(MsgTxt, K + 38, 2)
    recYCDOCOM0.CDOCOMSEQ = CLng(Val(mId$(MsgTxt, K + 40, 4)))
    recYCDOCOM0.CDOCOMCOM = mId$(MsgTxt, K + 44, 6)
    recYCDOCOM0.CDOCOMDEM = CLng(Val(mId$(MsgTxt, K + 50, 8)))
    recYCDOCOM0.CDOCOMREG = CLng(Val(mId$(MsgTxt, K + 58, 8)))
    recYCDOCOM0.CDOCOMCPT = mId$(MsgTxt, K + 66, 20)
    recYCDOCOM0.CDOCOMDEV = mId$(MsgTxt, K + 86, 3)
    recYCDOCOM0.CDOCOMVAL = CLng(Val(mId$(MsgTxt, K + 89, 8)))
    recYCDOCOM0.CDOCOMCOU = CDbl(Val(mId$(MsgTxt, K + 97, 15))) / 1000000000
    recYCDOCOM0.CDOCOMMRE = mId$(MsgTxt, K + 112, 3)
    recYCDOCOM0.CDOCOMBEN = mId$(MsgTxt, K + 115, 1)
    recYCDOCOM0.CDOCOMMON = CCur(Val(mId$(MsgTxt, K + 116, 16))) / 100
    recYCDOCOM0.CDOCOMMTV = CCur(Val(mId$(MsgTxt, K + 132, 16))) / 100
    recYCDOCOM0.CDOCOMAVI = mId$(MsgTxt, K + 148, 1)
    recYCDOCOM0.CDOCOMPRO = mId$(MsgTxt, K + 149, 1)
    recYCDOCOM0.CDOCOMUTR = CLng(Val(mId$(MsgTxt, K + 150, 6)))
    recYCDOCOM0.CDOCOMNRE = CLng(Val(mId$(MsgTxt, K + 156, 4)))
    recYCDOCOM0.CDOCOMETA = mId$(MsgTxt, K + 160, 2)
    recYCDOCOM0.CDOCOMSPE = CLng(Val(mId$(MsgTxt, K + 162, 4)))
    recYCDOCOM0.CDOCOMDBP = CLng(Val(mId$(MsgTxt, K + 166, 8)))
    recYCDOCOM0.CDOCOMFNP = CLng(Val(mId$(MsgTxt, K + 174, 8)))
    recYCDOCOM0.CDOCOMCUT = CInt(Val(mId$(MsgTxt, K + 182, 5)))
    recYCDOCOM0.CDOCOMCER = mId$(MsgTxt, K + 187, 1)
Else
    srvYCDOCOM0_GetBuffer = recYCDOCOM0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYCDOCOM0Len

End Function

'---------------------------------------------------------
Private Sub srvYCDOCOM0_PutBuffer(recYCDOCOM0 As typeYCDOCOM0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recYCDOCOM0Len) = Space$(recYCDOCOM0Len)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCDOCOM0.Obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCDOCOM0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
    Mid$(MsgTxt, K + 1, 5) = Format$(recYCDOCOM0.CDOCOMETB, "0000 ")
    Mid$(MsgTxt, K + 6, 5) = Format$(recYCDOCOM0.CDOCOMAGE, "0000 ")
    Mid$(MsgTxt, K + 11, 2) = recYCDOCOM0.CDOCOMSER
    Mid$(MsgTxt, K + 13, 2) = recYCDOCOM0.CDOCOMSSE
    Mid$(MsgTxt, K + 15, 3) = recYCDOCOM0.CDOCOMCOP
    Mid$(MsgTxt, K + 18, 10) = Format$(recYCDOCOM0.CDOCOMDOS, "000000000 ")
    Mid$(MsgTxt, K + 28, 4) = Format$(recYCDOCOM0.CDOCOMNUR, "000 ")
    Mid$(MsgTxt, K + 32, 6) = Format$(recYCDOCOM0.CDOCOMUTI, "00000 ")
    Mid$(MsgTxt, K + 38, 2) = recYCDOCOM0.CDOCOMEVE
    Mid$(MsgTxt, K + 40, 4) = Format$(recYCDOCOM0.CDOCOMSEQ, "000 ")
    Mid$(MsgTxt, K + 44, 6) = recYCDOCOM0.CDOCOMCOM
    Mid$(MsgTxt, K + 50, 8) = Format$(recYCDOCOM0.CDOCOMDEM, "0000000 ")
    Mid$(MsgTxt, K + 58, 8) = Format$(recYCDOCOM0.CDOCOMREG, "0000000 ")
    Mid$(MsgTxt, K + 66, 20) = recYCDOCOM0.CDOCOMCPT
    Mid$(MsgTxt, K + 86, 3) = recYCDOCOM0.CDOCOMDEV
    Mid$(MsgTxt, K + 89, 8) = Format$(recYCDOCOM0.CDOCOMVAL, "0000000 ")
    Mid$(MsgTxt, K + 97, 15) = Format$(recYCDOCOM0.CDOCOMCOU * 1000000000, "00000000000000 ")
    Mid$(MsgTxt, K + 112, 3) = recYCDOCOM0.CDOCOMMRE
    Mid$(MsgTxt, K + 115, 1) = recYCDOCOM0.CDOCOMBEN
    Mid$(MsgTxt, K + 116, 16) = Format$(recYCDOCOM0.CDOCOMMON * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 132, 16) = Format$(recYCDOCOM0.CDOCOMMTV * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 148, 1) = recYCDOCOM0.CDOCOMAVI
    Mid$(MsgTxt, K + 149, 1) = recYCDOCOM0.CDOCOMPRO
    Mid$(MsgTxt, K + 150, 6) = Format$(recYCDOCOM0.CDOCOMUTR, "00000 ")
    Mid$(MsgTxt, K + 156, 4) = Format$(recYCDOCOM0.CDOCOMNRE, "000 ")
    Mid$(MsgTxt, K + 160, 2) = recYCDOCOM0.CDOCOMETA
    Mid$(MsgTxt, K + 162, 4) = Format$(recYCDOCOM0.CDOCOMSPE, "000 ")
    Mid$(MsgTxt, K + 166, 8) = Format$(recYCDOCOM0.CDOCOMDBP, "0000000 ")
    Mid$(MsgTxt, K + 174, 8) = Format$(recYCDOCOM0.CDOCOMFNP, "0000000 ")
    Mid$(MsgTxt, K + 182, 5) = Format$(recYCDOCOM0.CDOCOMCUT, "0000 ")
    Mid$(MsgTxt, K + 187, 1) = recYCDOCOM0.CDOCOMCER
    MsgTxtLen = MsgTxtLen + recYCDOCOM0Len
End Sub



'---------------------------------------------------------
Private Function srvYCDOCOM0_Seek(recYCDOCOM0 As typeYCDOCOM0)
'---------------------------------------------------------

srvYCDOCOM0_Seek = "?"
MsgTxtLen = 0
Call srvYCDOCOM0_PutBuffer(recYCDOCOM0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvYCDOCOM0_GetBuffer(recYCDOCOM0)) Then
            srvYCDOCOM0_Seek = Null
        Else
            Call srvYCDOCOM0_Error(recYCDOCOM0)
        End If
    End If
End If

End Function
Public Sub srvYCDOCOM0_ElpDisplay(recYCDOCOM0 As typeYCDOCOM0)
frmElpDisplay.fgData.Rows = 32
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMETB
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMSER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMSER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMSSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMSSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMCOP    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMCOP
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMDOS    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMDOS
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMNUR    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° RENOUVELLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMNUR
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMUTI    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° UTILILSAT°./MODIF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMUTI
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMEVE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EVENEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMEVE
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMSEQ    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° SEQUENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMSEQ
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMCOM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE COMMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMCOM
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMDEM    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DT DEMANDE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMDEM
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMREG    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DT REGLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMREG
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMCPT   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DU COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMCPT
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMDEV    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE COMMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMDEV
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMVAL    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE VALEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMVAL
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMCOU 14.9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COURS DEVCOM/DEVCPT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMCOU
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMMRE    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MODE DE REGLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMMRE
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMBEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BENEFICIAIRE O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMBEN
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMMON 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT COMMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMMON
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMMTV 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMMTV
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMAVI    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "1 NON,2 A EDIT,3 EDI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMAVI
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMPRO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "A PROVISIONNER (O/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMPRO
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMUTR    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILISATION DU REGLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMUTR
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMNRE    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° REGLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMNRE
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMETA    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETAT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMETA
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMSPE    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N°SEQUENCE PERIODIQ"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMSPE
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMDBP    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DEBUT PERIODE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMDBP
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMFNP    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE FIN PERIODE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMFNP
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMCUT    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILISATEUR SAISIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMCUT
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOCOMCER    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COTAT°(O=CERTAIN/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOCOM0.CDOCOMCER
frmElpDisplay.Show vbModal
End Sub
Public Sub srvYCDOCOM0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YCDOCOM0.txt" For Input As #1
Open "C:\Temp\YCDOCOM0.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "CDOCOMETB;CDOCOMAGE;CDOCOMSER;CDOCOMSSE;CDOCOMCOP;CDOCOMDOS;CDOCOMNUR;CDOCOMUTI;CDOCOMEVE;CDOCOMSEQ;CDOCOMCOM;CDOCOMDEM;CDOCOMREG;CDOCOMCPT;CDOCOMDEV;CDOCOMVAL;CDOCOMCOU;CDOCOMMRE;CDOCOMBEN;CDOCOMMON;CDOCOMMTV;CDOCOMAVI;CDOCOMPRO;CDOCOMUTR;CDOCOMNRE;CDOCOMETA;CDOCOMSPE;CDOCOMDBP;CDOCOMFNP;CDOCOMCUT;CDOCOMCER;"
    Print #2, "CODE ETABLISSEMENT;AGENCE;SERVICE;SOUS-SERVICE;CODE OPERATION;NUMERO DOSSIER;N° RENOUVELLEMENT;N° UTILILSAT°./MODIF;EVENEMENT;N° SEQUENCE;CODE COMMISSION;DT DEMANDE;DT REGLEMENT;NUMERO DU COMPTE;DEVISE COMMISSION;DATE VALEUR;COURS DEVCOM/DEVCPT;MODE DE REGLEMENT;BENEFICIAIRE O/N;MONTANT COMMISSION;MONTANT TVA;1 NON,2 A EDIT,3 EDI;A PROVISIONNER (O/N);UTILISATION DU REGLE;N° REGLEMENT;ETAT;N°SEQUENCE PERIODIQ;DATE DEBUT PERIODE;DATE FIN PERIODE;UTILISATEUR SAISIE;COTAT°(O=CERTAIN/N);"
    Print #2, ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 5) & ";" & mId$(xIn, 11, 2) & ";" & mId$(xIn, 13, 2) & ";" & mId$(xIn, 15, 3) & ";" & mId$(xIn, 18, 10) & ";" _
      & mId$(xIn, 28, 4) & ";" & mId$(xIn, 32, 6) & ";" & mId$(xIn, 38, 2) & ";" & mId$(xIn, 40, 4) & ";" & mId$(xIn, 44, 6) & ";" _
      & mId$(xIn, 50, 8) & ";" & mId$(xIn, 58, 8) & ";" & mId$(xIn, 66, 20) & ";" _
      & mId$(xIn, 86, 3) & ";" & mId$(xIn, 89, 8) & ";" & mId$(xIn, 97, 15) & ";" _
      & mId$(xIn, 112, 3) & ";" & mId$(xIn, 115, 1) & ";" & mId$(xIn, 116, 16) & ";" & mId$(xIn, 132, 16) & ";" & mId$(xIn, 148, 1) & ";" _
      & mId$(xIn, 149, 1) & ";" & mId$(xIn, 150, 6) & ";" & mId$(xIn, 156, 4) & ";" & mId$(xIn, 160, 2) & ";" _
      & mId$(xIn, 162, 4) & ";" & mId$(xIn, 166, 8) & ";" & mId$(xIn, 174, 8) & ";" & mId$(xIn, 182, 5) & ";" _
      & mId$(xIn, 187, 1) & ";"
Loop
Close
End Sub

'-----------------------------------------------------
Function srvYCDOCOM0_Update(recYCDOCOM0 As typeYCDOCOM0)
'-----------------------------------------------------

srvYCDOCOM0_Update = "?"

MsgTxtLen = 0
Call srvYCDOCOM0_PutBuffer(recYCDOCOM0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYCDOCOM0_GetBuffer(recYCDOCOM0)) Then
        Call srvYCDOCOM0_Error(recYCDOCOM0)
        srvYCDOCOM0_Update = recYCDOCOM0.Err
        Exit Function
    Else
        srvYCDOCOM0_Update = Null
    End If
Else
    recYCDOCOM0.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recYCDOCOM0_Init(recYCDOCOM0 As typeYCDOCOM0)
'---------------------------------------------------------
MsgTxt = Space$(recYCDOCOM0Len)
MsgTxtIndex = 0
Call srvYCDOCOM0_GetBuffer(recYCDOCOM0)
recYCDOCOM0.Obj = "ZCDODOS0_S"

recYCDOCOM0.CDOCOMETB = 1 'As Integer                        ' CODE ETABLISSEMENT
recYCDOCOM0.CDOCOMAGE = 1 'As Integer                        ' AGENCE
recYCDOCOM0.CDOCOMSER = "" 'As String * 2                     ' SERVICE
recYCDOCOM0.CDOCOMSSE = "" 'As String * 2                     ' SOUS-SERVICE
recYCDOCOM0.CDOCOMCOP = "" 'As String * 3                     ' CODE OPERATION
recYCDOCOM0.CDOCOMDOS = 0 'As Long                            ' NUMERO DOSSIER
recYCDOCOM0.CDOCOMNUR = 0 'As Long                            ' N° RENOUVELLEMENT
recYCDOCOM0.CDOCOMUTI = 0 'As Long                            ' N° UTILILSAT°./MODIF
recYCDOCOM0.CDOCOMEVE = "" 'As String * 2                     ' EVENEMENT
recYCDOCOM0.CDOCOMSEQ = 0 'As Long                            ' N° SEQUENCE
recYCDOCOM0.CDOCOMCOM = "" 'As String * 6                     ' CODE COMMISSION
recYCDOCOM0.CDOCOMDEM = 0 'As Long                            ' DT DEMANDE
recYCDOCOM0.CDOCOMREG = 0 'As Long                            ' DT REGLEMENT
recYCDOCOM0.CDOCOMCPT = "" 'As String * 20                    ' NUMERO DU COMPTE
recYCDOCOM0.CDOCOMDEV = "" 'As String * 3                     ' DEVISE COMMISSION
recYCDOCOM0.CDOCOMVAL = 0 'As Long                            ' DATE VALEUR
recYCDOCOM0.CDOCOMCOU = 0 'As Double                         ' COURS DEVCOM/DEVCPT
recYCDOCOM0.CDOCOMMRE = "" 'As String * 3                     ' MODE DE REGLEMENT
recYCDOCOM0.CDOCOMBEN = "" 'As String * 1                     ' BENEFICIAIRE O/N
recYCDOCOM0.CDOCOMMON = 0 'As Currency                       ' MONTANT COMMISSION
recYCDOCOM0.CDOCOMMTV = 0 'As Currency                       ' MONTANT TVA
recYCDOCOM0.CDOCOMAVI = "" 'As String * 1                     ' 1 NON,2 A EDIT,3 EDI
recYCDOCOM0.CDOCOMPRO = "" 'As String * 1                     ' A PROVISIONNER (O/N)
recYCDOCOM0.CDOCOMUTR = 0 'As Long                            ' UTILISATION DU REGLE
recYCDOCOM0.CDOCOMNRE = 0 'As Long                            ' N° REGLEMENT
recYCDOCOM0.CDOCOMETA = "" 'As String * 2                     ' ETAT
recYCDOCOM0.CDOCOMSPE = 0 'As Long                            ' N°SEQUENCE PERIODIQ
recYCDOCOM0.CDOCOMDBP = 0 'As Long                            ' DATE DEBUT PERIODE
recYCDOCOM0.CDOCOMFNP = 0 'As Long                            ' DATE FIN PERIODE
recYCDOCOM0.CDOCOMCUT = 0 'As Integer                        ' UTILISATEUR SAISIE
recYCDOCOM0.CDOCOMCER = "" 'As String * 1                     ' COTAT°(O=CERTAIN/N)

End Sub




