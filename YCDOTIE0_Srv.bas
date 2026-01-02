Attribute VB_Name = "srvYCDOTIE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCDOTIE0Len = 255 ' 34 +221
Public Const recYCDOTIE0_Block = 50
Public Const memoYCDOTIE0Len = 221
Public Const constYCDOTIE0 = "YCDOTIE0"
Public paramYCDOTIE0_Import As String
Dim meYbase As typeYBase

Type typeYCDOTIE0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    CDOTIEETB       As Integer                        ' CODE ETABLISSEMENT
    CDOTIETIE       As String * 7                     ' NUMERO TIERS
    CDOTIECLI       As String * 7                     ' CLIENT ASSOCIE
    CDOTIERA1       As String * 32                    ' NOM OU DESIGNATION
    CDOTIERA2       As String * 32                    ' PRENOM/DESIGNATION
    CDOTIESIG       As String * 12                    ' SIGLE USUEL
    CDOTIEPAR       As String * 3                     ' CDE PAYS DE RESIDENC
    CDOTIEECO       As String * 3                     ' QUALITE/AG ECONOMIQU
    CDOTIECAT       As String * 3                     ' CATEGORIE CLIENT
    CDOTIEMES       As String * 1                     ' LANGUE MESSAGERIE
    CDOTIEBIC       As String * 16                    ' BIC (SWIFT)
    CDOTIEBAN       As String * 5                     ' CODE BANQUE
    CDOTIEGUI       As String * 5                     ' CODE GUICHET
    CDOTIECOM       As String * 20                    ' COMPTE
    CDOTIEAD1       As String * 32                    ' ADRESSE 1
    CDOTIEAD2       As String * 32                    ' ADRESSE 2
    CDOTIEAD3       As String * 32                    ' COMMUNE
    CDOTIECOP       As String * 6                     ' CODE POSTAL
    CDOTIEVIL       As String * 25                    ' BUREAU DISTRIBUTEUR
    CDOTIEPAY       As String * 32                    ' PAYS
    CDOTIETEL       As String * 20                    ' TELEPHONE
    CDOTIEFAX       As String * 20                    ' No FAX
    CDOTIETEX       As String * 20                    ' No TELEX
    CDOTIESRN       As String * 9                     ' NUMERO SIREN
    CDOTIECOT       As String * 1                     ' CORRESPOND. CLI/TIE
    CDOTIECOR       As String * 7                     ' CORRESPONDANT
    
End Type
    
    
Public arrYCDOTIE0() As typeYCDOTIE0
Public arrYCDOTIE0_NB As Integer
Public arrYCDOTIE0_NBMax As Integer
Public arrYCDOTIE0_Index As Integer
Public arrYCDOTIE0_Suite As Boolean

Dim meYCLIENA0 As typeYCLIENA0
'-----------------------------------------------------
Function srvYCDOTIE0_Update(recYCDOTIE0 As typeYCDOTIE0)
'-----------------------------------------------------

srvYCDOTIE0_Update = "?"

MsgTxtLen = 0
Call srvYCDOTIE0_PutBuffer(recYCDOTIE0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYCDOTIE0_GetBuffer(recYCDOTIE0)) Then
        Call srvYCDOTIE0_Error(recYCDOTIE0)
        srvYCDOTIE0_Update = recYCDOTIE0.Err
        Exit Function
    Else
        srvYCDOTIE0_Update = Null
    End If
Else
    recYCDOTIE0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYCDOTIE0_Error(recYCDOTIE0 As typeYCDOTIE0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YCDOTIE0" & Chr$(10) & Chr$(13)

Select Case mId$(recYCDOTIE0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYCDOTIE0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YCDOTIE0s.bas  ( " & Trim(recYCDOTIE0.obj) & " : " & Trim(recYCDOTIE0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYCDOTIE0_Monitor(recYCDOTIE0 As typeYCDOTIE0)
'-----------------------------------------------------

arrYCDOTIE0_Suite = False
Select Case mId$(Trim(recYCDOTIE0.Method), 1, 4)
    Case "Snap"
              srvYCDOTIE0_Monitor = srvYCDOTIE0_Snap(recYCDOTIE0)
    Case Else
            srvYCDOTIE0_Monitor = srvYCDOTIE0_Seek(recYCDOTIE0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYCDOTIE0_GetBuffer(recYCDOTIE0 As typeYCDOTIE0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYCDOTIE0_GetBuffer = Null
recYCDOTIE0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYCDOTIE0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYCDOTIE0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYCDOTIE0.Err = Space$(10) Then
    recYCDOTIE0.CDOTIEETB = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOTIE0.CDOTIETIE = mId$(MsgTxt, K + 6, 7)
    recYCDOTIE0.CDOTIECLI = mId$(MsgTxt, K + 13, 7)
    recYCDOTIE0.CDOTIERA1 = mId$(MsgTxt, K + 20, 32)
    recYCDOTIE0.CDOTIERA2 = mId$(MsgTxt, K + 52, 32)
    recYCDOTIE0.CDOTIESIG = mId$(MsgTxt, K + 84, 12)
    recYCDOTIE0.CDOTIEPAR = mId$(MsgTxt, K + 96, 3)
    recYCDOTIE0.CDOTIEECO = mId$(MsgTxt, K + 99, 3)
    recYCDOTIE0.CDOTIECAT = mId$(MsgTxt, K + 102, 3)
    recYCDOTIE0.CDOTIEMES = mId$(MsgTxt, K + 105, 1)
    recYCDOTIE0.CDOTIEBIC = mId$(MsgTxt, K + 106, 16)
    recYCDOTIE0.CDOTIEBAN = mId$(MsgTxt, K + 122, 5)
    recYCDOTIE0.CDOTIEGUI = mId$(MsgTxt, K + 127, 5)
    recYCDOTIE0.CDOTIECOM = mId$(MsgTxt, K + 132, 20)
    recYCDOTIE0.CDOTIEAD1 = mId$(MsgTxt, K + 152, 32)
    recYCDOTIE0.CDOTIEAD2 = mId$(MsgTxt, K + 184, 32)
    recYCDOTIE0.CDOTIEAD3 = mId$(MsgTxt, K + 216, 32)
    recYCDOTIE0.CDOTIECOP = mId$(MsgTxt, K + 248, 6)
    recYCDOTIE0.CDOTIEVIL = mId$(MsgTxt, K + 254, 25)
    recYCDOTIE0.CDOTIEPAY = mId$(MsgTxt, K + 279, 32)
    recYCDOTIE0.CDOTIETEL = mId$(MsgTxt, K + 311, 20)
    recYCDOTIE0.CDOTIEFAX = mId$(MsgTxt, K + 331, 20)
    recYCDOTIE0.CDOTIETEX = mId$(MsgTxt, K + 351, 20)
    recYCDOTIE0.CDOTIESRN = mId$(MsgTxt, K + 371, 9)
    recYCDOTIE0.CDOTIECOT = mId$(MsgTxt, K + 380, 1)
    recYCDOTIE0.CDOTIECOR = mId$(MsgTxt, K + 381, 7)
Else
    srvYCDOTIE0_GetBuffer = recYCDOTIE0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYCDOTIE0Len

End Function

'---------------------------------------------------------
Public Function srvYCDOTIE0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCDOTIE0 As typeYCDOTIE0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCDOTIE0_GetBuffer_ODBC = Null

    recYCDOTIE0.CDOTIEETB = rsADO("CDOTIEETB")    'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOTIE0.CDOTIETIE = rsADO("CDOTIETIE")    'mId$(MsgTxt, K + 6, 7)
    recYCDOTIE0.CDOTIECLI = rsADO("CDOTIECLI")    'mId$(MsgTxt, K + 13, 7)
    recYCDOTIE0.CDOTIERA1 = rsADO("CDOTIERA1")    'mId$(MsgTxt, K + 20, 32)
    recYCDOTIE0.CDOTIERA2 = rsADO("CDOTIERA2")    'mId$(MsgTxt, K + 52, 32)
    recYCDOTIE0.CDOTIESIG = rsADO("CDOTIESIG")    'mId$(MsgTxt, K + 84, 12)
    recYCDOTIE0.CDOTIEPAR = rsADO("CDOTIEPAR")    'mId$(MsgTxt, K + 96, 3)
    recYCDOTIE0.CDOTIEECO = rsADO("CDOTIEECO")    'mId$(MsgTxt, K + 99, 3)
    recYCDOTIE0.CDOTIECAT = rsADO("CDOTIECAT")    'mId$(MsgTxt, K + 102, 3)
    recYCDOTIE0.CDOTIEMES = rsADO("CDOTIEMES")    'mId$(MsgTxt, K + 105, 1)
    recYCDOTIE0.CDOTIEBIC = rsADO("CDOTIEBIC")    'mId$(MsgTxt, K + 106, 16)
    recYCDOTIE0.CDOTIEBAN = rsADO("CDOTIEBAN")    'mId$(MsgTxt, K + 122, 5)
    recYCDOTIE0.CDOTIEGUI = rsADO("CDOTIEGUI")    'mId$(MsgTxt, K + 127, 5)
    recYCDOTIE0.CDOTIECOM = rsADO("CDOTIECOM")    'mId$(MsgTxt, K + 132, 20)
    recYCDOTIE0.CDOTIEAD1 = rsADO("CDOTIEAD1")    'mId$(MsgTxt, K + 152, 32)
    recYCDOTIE0.CDOTIEAD2 = rsADO("CDOTIEAD2")    'mId$(MsgTxt, K + 184, 32)
    recYCDOTIE0.CDOTIEAD3 = rsADO("CDOTIEAD3")    'mId$(MsgTxt, K + 216, 32)
    recYCDOTIE0.CDOTIECOP = rsADO("CDOTIECOP")    'mId$(MsgTxt, K + 248, 6)
    recYCDOTIE0.CDOTIEVIL = rsADO("CDOTIEVIL")    'mId$(MsgTxt, K + 254, 25)
    recYCDOTIE0.CDOTIEPAY = rsADO("CDOTIEPAY")    'mId$(MsgTxt, K + 279, 32)
    recYCDOTIE0.CDOTIETEL = rsADO("CDOTIETEL")    'mId$(MsgTxt, K + 311, 20)
    recYCDOTIE0.CDOTIEFAX = rsADO("CDOTIEFAX")    'mId$(MsgTxt, K + 331, 20)
    recYCDOTIE0.CDOTIETEX = rsADO("CDOTIETEX")    'mId$(MsgTxt, K + 351, 20)
    recYCDOTIE0.CDOTIESRN = rsADO("CDOTIESRN")    'mId$(MsgTxt, K + 371, 9)
    recYCDOTIE0.CDOTIECOT = rsADO("CDOTIECOT")    'mId$(MsgTxt, K + 380, 1)
    recYCDOTIE0.CDOTIECOR = rsADO("CDOTIECOR")    'mId$(MsgTxt, K + 381, 7)

Exit Function

Error_Handler:
srvYCDOTIE0_GetBuffer_ODBC = Error

End Function



Public Function srvYCDOTIE0_Import(lX As String)
Dim xIn As String, X As String, Nb As String
On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOTIE0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    srvYCDOTIE0_Import = Null
    lX = meYbase.Text
    Exit Function
End If


srvYCDOTIE0_Import = "?"

paramYCDOTIE0_Import = paramYBase_DataF & Trim(constYCDOTIE0) & paramYBase_Data_ExtensionP

Open Trim(paramYCDOTIE0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYCDOTIE0) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYCDOTIE0
            meYbase.K1 = mId$(xIn, 6, 7)
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYCDOTIE0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOTIE0
lX = DSys & "_" & time_Hms & "_" & Nb
meYbase.Text = lX

dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOTIE0_Import" & xIn, vbCritical, Error
Close

srvYCDOTIE0_Import = Error
End Function


Public Function srvYCDOTIE0_Import_Read(lId As String, lYCDOTIE0 As typeYCDOTIE0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYCDOTIE0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYCDOTIE0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYCDOTIE0_GetBuffer lYCDOTIE0
    srvYCDOTIE0_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOTIE0_Import_Read" & xIn, vbCritical, Error
srvYCDOTIE0_Import_Read = Error
End Function


Public Sub srvYCDOTIE0_ElpDisplay(recYCDOTIE0 As typeYCDOTIE0)
frmElpDisplay.fgData.Rows = 27
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIEETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIEETB
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIETIE    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO TIERS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIETIE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIECLI    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CLIENT ASSOCIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIECLI
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIERA1   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NOM OU DESIGNATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIERA1
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIERA2   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PRENOM/DESIGNATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIERA2
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIESIG   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SIGLE USUEL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIESIG
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIEPAR    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CDE PAYS DE RESIDENC"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIEPAR
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIEECO    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "QUALITE/AG ECONOMIQU"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIEECO
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIECAT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CATEGORIE CLIENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIECAT
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIEMES    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LANGUE MESSAGERIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIEMES
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIEBIC   16A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BIC (SWIFT)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIEBIC
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIEBAN    5A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE BANQUE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIEBAN
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIEGUI    5A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE GUICHET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIEGUI
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIECOM   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIECOM
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIEAD1   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ADRESSE 1"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIEAD1
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIEAD2   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ADRESSE 2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIEAD2
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIEAD3   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMMUNE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIEAD3
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIECOP    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE POSTAL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIECOP
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIEVIL   25A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BUREAU DISTRIBUTEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIEVIL
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIEPAY   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PAYS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIEPAY
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIETEL   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TELEPHONE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIETEL
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIEFAX   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "No FAX"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIEFAX
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIETEX   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "No TELEX"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIETEX
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIESRN    9A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO SIREN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIESRN
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIECOT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CORRESPOND. CLI/TIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIECOT
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOTIECOR    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CORRESPONDANT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOTIE0.CDOTIECOR
frmElpDisplay.Show vbModal
End Sub

'---------------------------------------------------------
Public Sub srvYCDOTIE0_PutBuffer(recYCDOTIE0 As typeYCDOTIE0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCDOTIE0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCDOTIE0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYCDOTIE0.CDOTIEETB, "0000 ")
    Mid$(MsgTxt, K + 6, 7) = recYCDOTIE0.CDOTIETIE
    Mid$(MsgTxt, K + 13, 7) = recYCDOTIE0.CDOTIECLI
    Mid$(MsgTxt, K + 20, 32) = recYCDOTIE0.CDOTIERA1
    Mid$(MsgTxt, K + 52, 32) = recYCDOTIE0.CDOTIERA2
    Mid$(MsgTxt, K + 84, 12) = recYCDOTIE0.CDOTIESIG
    Mid$(MsgTxt, K + 96, 3) = recYCDOTIE0.CDOTIEPAR
    Mid$(MsgTxt, K + 99, 3) = recYCDOTIE0.CDOTIEECO
    Mid$(MsgTxt, K + 102, 3) = recYCDOTIE0.CDOTIECAT
    Mid$(MsgTxt, K + 105, 1) = recYCDOTIE0.CDOTIEMES
    Mid$(MsgTxt, K + 106, 16) = recYCDOTIE0.CDOTIEBIC
    Mid$(MsgTxt, K + 122, 5) = recYCDOTIE0.CDOTIEBAN
    Mid$(MsgTxt, K + 127, 5) = recYCDOTIE0.CDOTIEGUI
    Mid$(MsgTxt, K + 132, 20) = recYCDOTIE0.CDOTIECOM
    Mid$(MsgTxt, K + 152, 32) = recYCDOTIE0.CDOTIEAD1
    Mid$(MsgTxt, K + 184, 32) = recYCDOTIE0.CDOTIEAD2
    Mid$(MsgTxt, K + 216, 32) = recYCDOTIE0.CDOTIEAD3
    Mid$(MsgTxt, K + 248, 6) = recYCDOTIE0.CDOTIECOP
    Mid$(MsgTxt, K + 254, 25) = recYCDOTIE0.CDOTIEVIL
    Mid$(MsgTxt, K + 279, 32) = recYCDOTIE0.CDOTIEPAY
    Mid$(MsgTxt, K + 311, 20) = recYCDOTIE0.CDOTIETEL
    Mid$(MsgTxt, K + 331, 20) = recYCDOTIE0.CDOTIEFAX
    Mid$(MsgTxt, K + 351, 20) = recYCDOTIE0.CDOTIETEX
    Mid$(MsgTxt, K + 371, 9) = recYCDOTIE0.CDOTIESRN
    Mid$(MsgTxt, K + 380, 1) = recYCDOTIE0.CDOTIECOT
    Mid$(MsgTxt, K + 381, 7) = recYCDOTIE0.CDOTIECOR

MsgTxtLen = MsgTxtLen + recYCDOTIE0Len
End Sub



'---------------------------------------------------------
Private Function srvYCDOTIE0_Seek(recYCDOTIE0 As typeYCDOTIE0)
'---------------------------------------------------------

srvYCDOTIE0_Seek = "?"
MsgTxtLen = 0
Call srvYCDOTIE0_PutBuffer(recYCDOTIE0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYCDOTIE0_GetBuffer(recYCDOTIE0)) Then
        srvYCDOTIE0_Seek = Null
    Else
        Call srvYCDOTIE0_Error(recYCDOTIE0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYCDOTIE0_Snap(recYCDOTIE0 As typeYCDOTIE0)
'---------------------------------------------------------
srvYCDOTIE0_Snap = "?"
MsgTxtLen = 0
Call srvYCDOTIE0_PutBuffer(recYCDOTIE0)
Call srvYCDOTIE0_PutBuffer(arrYCDOTIE0(0))
If IsNull(SndRcv()) Then
    srvYCDOTIE0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYCDOTIE0_GetBuffer(recYCDOTIE0)) Then
            Call arrYCDOTIE0_AddItem(recYCDOTIE0)
            arrYCDOTIE0_Suite = True
        Else
            arrYCDOTIE0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYCDOTIE0_AddItem(recYCDOTIE0 As typeYCDOTIE0)
'---------------------------------------------------------
          
arrYCDOTIE0_NB = arrYCDOTIE0_NB + 1
    
If arrYCDOTIE0_NB > arrYCDOTIE0_NBMax Then
    arrYCDOTIE0_NBMax = arrYCDOTIE0_NBMax + recYCDOTIE0_Block
    ReDim Preserve arrYCDOTIE0(arrYCDOTIE0_NBMax)
End If
            
arrYCDOTIE0(arrYCDOTIE0_NB) = recYCDOTIE0
End Sub



'---------------------------------------------------------
Public Sub recYCDOTIE0_Init(recYCDOTIE0 As typeYCDOTIE0)
'---------------------------------------------------------
recYCDOTIE0.obj = "ZCDOTIE0_S"
recYCDOTIE0.Method = ""
recYCDOTIE0.Err = ""
recYCDOTIE0.CDOTIEETB = 0    'Integer                        ' CODE ETABLISSEMENT
recYCDOTIE0.CDOTIETIE = ""   'String * 7                     ' NUMERO TIERS
recYCDOTIE0.CDOTIECLI = ""   'String * 7                     ' CLIENT ASSOCIE
recYCDOTIE0.CDOTIERA1 = ""   'String * 32                    ' NOM OU DESIGNATION
recYCDOTIE0.CDOTIERA2 = ""   'String * 32                    ' PRENOM/DESIGNATION
recYCDOTIE0.CDOTIESIG = ""   'String * 12                    ' SIGLE USUEL
recYCDOTIE0.CDOTIEPAR = ""   'String * 3                     ' CDE PAYS DE RESIDENC
recYCDOTIE0.CDOTIEECO = ""   'String * 3                     ' QUALITE/AG ECONOMIQU
recYCDOTIE0.CDOTIECAT = ""   'String * 3                     ' CATEGORIE CLIENT
recYCDOTIE0.CDOTIEMES = ""   'String * 1                     ' LANGUE MESSAGERIE
recYCDOTIE0.CDOTIEBIC = ""   'String * 16                    ' BIC (SWIFT)
recYCDOTIE0.CDOTIEBAN = ""   'String * 5                     ' CODE BANQUE
recYCDOTIE0.CDOTIEGUI = ""   'String * 5                     ' CODE GUICHET
recYCDOTIE0.CDOTIECOM = ""   'String * 20                    ' COMPTE
recYCDOTIE0.CDOTIEAD1 = ""   'String * 32                    ' ADRESSE 1
recYCDOTIE0.CDOTIEAD2 = ""   'String * 32                    ' ADRESSE 2
recYCDOTIE0.CDOTIEAD3 = ""   'String * 32                    ' COMMUNE
recYCDOTIE0.CDOTIECOP = ""   'String * 6                     ' CODE POSTAL
recYCDOTIE0.CDOTIEVIL = ""   'String * 25                    ' BUREAU DISTRIBUTEUR
recYCDOTIE0.CDOTIEPAY = ""   'String * 32                    ' PAYS
recYCDOTIE0.CDOTIETEL = ""   'String * 20                    ' TELEPHONE
recYCDOTIE0.CDOTIEFAX = ""   'String * 20                    ' No FAX
recYCDOTIE0.CDOTIETEX = ""   'String * 20                    ' No TELEX
recYCDOTIE0.CDOTIESRN = ""   'String * 9                     ' NUMERO SIREN
recYCDOTIE0.CDOTIECOT = ""   'String * 1                     ' CORRESPOND. CLI/TIE
recYCDOTIE0.CDOTIECOR = ""   'String * 7                     ' CORRESPONDANT

End Sub









Public Sub srcYCDOTIE_Adresse(lCDODOSxxT As String, lCDODOSxxR As String, lCDODOSxxX As String, lYCDOTIE0 As typeYCDOTIE0, lYADRESS0 As typeYADRESS0, lConcat As String, lCodeAdresse As String, blnODBC As Boolean)
Dim wId As String
Dim X As String, X1 As String
Dim I As Integer, K As Integer
Dim V
Dim blnCDODOSxxX As Boolean

blnCDODOSxxX = False
recYADRESS0_Init lYADRESS0
If lCDODOSxxT = "T" Then
    lYCDOTIE0.CDOTIETIE = lCDODOSxxR
    If IsNull(srvYCDOTIE0_Read(lYCDOTIE0, blnODBC)) Then
        lYADRESS0.ADRESSETA = lYCDOTIE0.CDOTIEETB                      ' Etablissement
        lYADRESS0.ADRESSTYP = "T"      ' String * 1                     ' 1 client , 2 compte
        lYADRESS0.ADRESSPLA = 0       ' Long                           ' Numéro de plan
        lYADRESS0.ADRESSNUM = lYCDOTIE0.CDOTIETIE      ' String * 20                    ' ou numéro de client
        lYADRESS0.ADRESSCOA = ""      ' String * 2                     ' Code adresse
        lYADRESS0.ADRESSDLI = 0       ' Long                           ' Date limite validité
        lYADRESS0.ADRESSDDE = 0       ' Long                           ' Date début validité
        lYADRESS0.ADRESSRA1 = lYCDOTIE0.CDOTIERA1      ' String * 32                    ' ou raison sociale 1
        lYADRESS0.ADRESSRA2 = lYCDOTIE0.CDOTIERA2      ' String * 32                    ' ou raison sociale 2
        lYADRESS0.ADRESSAD1 = lYCDOTIE0.CDOTIEAD1     ' String * 32                    ' Adresse 1
        lYADRESS0.ADRESSAD2 = lYCDOTIE0.CDOTIEAD2     ' String * 32                    ' Adresse 2
        lYADRESS0.ADRESSAD3 = lYCDOTIE0.CDOTIEAD3      ' String * 32                    ' Adresse 3
        lYADRESS0.ADRESSCOP = lYCDOTIE0.CDOTIECOP    ' String * 6                     ' Code postal
        lYADRESS0.ADRESSVIL = lYCDOTIE0.CDOTIEVIL      ' String * 25                    ' Ville
        lYADRESS0.ADRESSPAY = lYCDOTIE0.CDOTIEPAY      ' String * 25                    ' Pays
        lYADRESS0.ADRESSTEL = lYCDOTIE0.CDOTIETEL     ' String * 20                    ' No Tel.
        lYADRESS0.ADRESSFAX = lYCDOTIE0.CDOTIEFAX       ' String * 20                    ' No Fax.
        lYADRESS0.ADRESSTEX = lYCDOTIE0.CDOTIETEX        ' String * 20                    ' No Télex
    End If
Else
    If Trim(lCDODOSxxR) <> "" Then
'Recherche adresse spécifique CREDOC dans le fichier ZADRESS0
        lYADRESS0.ADRESSTYP = "1"
        lYADRESS0.ADRESSNUM = " " & lCDODOSxxR
        lYADRESS0.ADRESSCOA = lCodeAdresse
        V = srvYADRESS0_Read(lYADRESS0, blnODBC)
        If Not IsNull(V) And lCodeAdresse <> "  " Then
            wId = "1 " & lCDODOSxxR
            lYADRESS0.ADRESSTYP = "1 "
            lYADRESS0.ADRESSNUM = " " & lCDODOSxxR
            lYADRESS0.ADRESSCOA = ""
            V = srvYADRESS0_Read(lYADRESS0, blnODBC)
        End If
        If Trim(lYADRESS0.ADRESSRA1) = "" Then
            meYCLIENA0.CLIENACLI = lCDODOSxxR
            If IsNull(srvYCLIENA0_Read(meYCLIENA0, blnODBC)) Then
                lYADRESS0.ADRESSRA1 = meYCLIENA0.CLIENARA1
                lYADRESS0.ADRESSRA2 = meYCLIENA0.CLIENARA2
            End If
        End If
    Else
        blnCDODOSxxX = True
        lYADRESS0.ADRESSRA1 = mId$(lCDODOSxxX, 1, 32)
        lYADRESS0.ADRESSVIL = mId$(lCDODOSxxX, 33, 32)
    End If
End If
If blnCDODOSxxX Then
    X = lCDODOSxxX
Else
    X = Trim(lYADRESS0.ADRESSRA1) & " - " & Trim(lYADRESS0.ADRESSCOP) & " " & Trim(lYADRESS0.ADRESSVIL) & " - " & Trim(lYADRESS0.ADRESSPAY)
End If
K = 1
lConcat = mId$(X, 1, 1)
For I = 2 To Len(X)
    X1 = mId$(X, I, 1)
    If X1 <> " " Or mId$(lConcat, I, 1) <> " " Then lConcat = lConcat & X1
Next I
End Sub
Public Function srvYCDOTIE0_Read(lYCDOTIE0 As typeYCDOTIE0, blnODBC As Boolean)

Dim xSQL As String
Dim V
Dim rsADO As New ADODB.Recordset

On Error GoTo Error_Handler

srvYCDOTIE0_Read = Null

If Not blnODBC Then
 'Lecture YBASE
'===============
   xSQL = lYCDOTIE0.CDOTIETIE
    srvYCDOTIE0_Read = srvYCDOTIE0_Import_Read(xSQL, lYCDOTIE0)
Else
'Lecture ODBC
'===============
    Set rsADO = Nothing
    xSQL = "select * from ZCDOTIE0 where CDOTIETIE = '" & lYCDOTIE0.CDOTIETIE & "'"
    rsADO.Open xSQL, paramODBC_DSN_SAB
    If rsADO.EOF Then
        V = "Tiers_CDO inconnu"
    Else
        V = srvYCDOTIE0_GetBuffer_ODBC(rsADO, lYCDOTIE0)
    End If
    
    If Not IsNull(V) Then
        srvYCDOTIE0_Read = "Lecture ZCDOTIE0 : " & V
        Exit Function
    End If
    rsADO.Close
End If

Exit Function

Error_Handler:
srvYCDOTIE0_Read = Error
     
End Function


