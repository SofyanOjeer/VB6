Attribute VB_Name = "srvYADRESS0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYADRESS0Len = 255 ' 34 +221
Public Const recYADRESS0_Block = 50
Public Const memoYADRESS0Len = 221
Public Const constYADRESS0 = "YADRESS0  "
Public paramYADRESS0_Import As String
Dim meYbase As typeYBase

Type typeYADRESS0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    ADRESSETA       As Integer                        ' Etablissement
    ADRESSTYP       As String * 1                     ' 1 client , 2 compte
    ADRESSPLA       As Long                           ' Numéro de plan
    ADRESSNUM       As String * 20                    ' ou numéro de client
    ADRESSCOA       As String * 2                     ' Code adresse
    ADRESSDLI       As Long                           ' Date limite validité
    ADRESSDDE       As Long                           ' Date début validité
    ADRESSRA1       As String * 32                    ' ou raison sociale 1
    ADRESSRA2       As String * 32                    ' ou raison sociale 2
    ADRESSAD1       As String * 32                    ' Adresse 1
    ADRESSAD2       As String * 32                    ' Adresse 2
    ADRESSAD3       As String * 32                    ' Adresse 3
    ADRESSCOP       As String * 6                     ' Code postal
    ADRESSVIL       As String * 25                    ' Ville
    ADRESSPAY       As String * 25                    ' Pays
    ADRESSTEL       As String * 20                    ' No Tel.
    ADRESSFAX       As String * 20                    ' No Fax.
    ADRESSTEX       As String * 20                    ' No Télex
End Type
    
    
Public arrYADRESS0() As typeYADRESS0
Public arrYADRESS0_NB As Integer
Public arrYADRESS0_NBMax As Integer
Public arrYADRESS0_Index As Integer
Public arrYADRESS0_Suite As Boolean

'---------------------------------------------------------
Public Function srvYADRESS0_Compte(lYADRESS0 As typeYADRESS0, cnADO As ADODB.Connection)
'---------------------------------------------------------
'Initialiser .ADRESSNUM= 'numéro de compte
'            .ADRESSCOA = '  ','CO','CH' ......
'=================================
Dim wADRESSNUM As String, wADRESSCOA As String
Dim blnOk As Boolean
Dim xSql As String
Dim rsADO As ADODB.Recordset
Dim V
Dim wYTITULA0 As typeYTITULA0

On Error GoTo Error_Handler
srvYADRESS0_Compte = Null
blnOk = False

wADRESSNUM = lYADRESS0.ADRESSNUM
wADRESSCOA = lYADRESS0.ADRESSCOA
recYADRESS0_Init lYADRESS0

'Lecture Adresse avec Code Adresse
'=================================
Set rsADO = Nothing
xSql = "select * from ZADRESS0 where ADRESSNUM = '" & wADRESSNUM & "' AND  ADRESSCOA = '" & wADRESSCOA _
        & "' AND ADRESSTYP = '2'"
Set rsADO = cnADO.Execute(xSql)
If Not rsADO.EOF Then
    V = srvZADRESS0_GetBuffer_ODBC(rsADO, lYADRESS0)
    If Not IsNull(V) Then
        srvYADRESS0_Compte = "srvYADRESS0_Compte_1 : Lecture ZADRESS0 : " & V
        Exit Function
    Else
        blnOk = True
    End If
End If
If blnOk Then Exit Function

'SINON Lecture Adresse avec type= '  '
'=================================
If wADRESSCOA <> "  " Then
    xSql = "select * from ZADRESS0 where ADRESSNUM = '" & wADRESSNUM & "' AND  ADRESSCOA = '  '" _
            & " AND ADRESSTYP = '2'"
    Set rsADO = cnADO.Execute(xSql)
    If Not rsADO.EOF Then
        V = srvZADRESS0_GetBuffer_ODBC(rsADO, lYADRESS0)
        If Not IsNull(V) Then
            srvYADRESS0_Compte = "srvYADRESS0_Compte_2 : Lecture ZADRESS0 : " & V
            Exit Function
        Else
            blnOk = True
        End If
    End If
End If

If blnOk Then Exit Function

'SINON Lecture TITULAIRE principal ==> ADRESSE CLIENT
'=================================
xSql = "select * from ZTITULA0 where TITULACOM = '" & wADRESSNUM & "' AND  TITULATPR = '0'"
Set rsADO = cnADO.Execute(xSql)
If Not rsADO.EOF Then
    V = srvYTITULA0_GetBuffer_ODBC(rsADO, wYTITULA0)
    If Not IsNull(V) Then
        srvYADRESS0_Compte = "srvYADRESS0_Compte_3 : Lecture ZTITULA0 : " & V
        Exit Function
    Else
        blnOk = True
    End If
End If

lYADRESS0.ADRESSNUM = wYTITULA0.TITULACLI
lYADRESS0.ADRESSCOA = wADRESSCOA

Call srvYADRESS0_Client(lYADRESS0, cnADO)

If blnOk Then Exit Function
srvYADRESS0_Compte = "Adresse non trouvée  : " & wADRESSNUM & " & " & wYTITULA0.TITULACLI

Exit Function

Error_Handler:
srvYADRESS0_Compte = Error

End Function
'---------------------------------------------------------
Public Function srvYADRESS0_Compte_BIC(lCOMPTECOM As String, lBIC As String, cnADO As ADODB.Connection)
'---------------------------------------------------------
Dim xSql As String
Dim rsADO As ADODB.Recordset
Dim wCLIENACLI As String
Dim blnOk As Boolean

On Error GoTo Error_Handler

lBIC = ""
'Lecture Compte => Racine
'==========================
srvYADRESS0_Compte_BIC = "? YBIACPT0 : srvYADRESS0_Compte_BIC"
blnOk = False
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '" & lCOMPTECOM & "'"
Set rsADO = cnADO.Execute(xSql)
If Not rsADO.EOF Then
    wCLIENACLI = rsADO("CLIENACLI")
    srvYADRESS0_Compte_BIC = "? ZADRESS0 : srvYADRESS0_Compte_BIC"

'Lecture Compte => Racine
    '==========================
    Set rsADO = Nothing
    xSql = "select * from " & paramIBM_Library_SAB & ".ZADRESS0 where ADRESSTYP = '4' and ADRESSNUM = ' " & wCLIENACLI & "'"
    Set rsADO = cnADO.Execute(xSql)
    If Not rsADO.EOF Then
        lBIC = mId$(rsADO("ADRESSRA1"), 11, 11)
        srvYADRESS0_Compte_BIC = Null
    End If
End If

Exit Function

Error_Handler:
srvYADRESS0_Compte_BIC = Error

End Function

'---------------------------------------------------------
Public Function srvYADRESS0_Client(lYADRESS0 As typeYADRESS0, cnADO As ADODB.Connection)
'---------------------------------------------------------
Dim wADRESSNUM As String, wADRESSCOA As String
Dim blnOk As Boolean
Dim xSql As String
Dim rsADO As ADODB.Recordset
Dim V
Dim wYCLIENA0 As typeYCLIENA0

On Error GoTo Error_Handler
srvYADRESS0_Client = Null
blnOk = False

wADRESSNUM = " " & Trim(lYADRESS0.ADRESSNUM)
wADRESSCOA = lYADRESS0.ADRESSCOA
recYADRESS0_Init lYADRESS0

'Lecture Adresse avec Code Adresse
'=================================
Set rsADO = Nothing
xSql = "select * from ZADRESS0 where ADRESSNUM = '" & wADRESSNUM & "' AND  ADRESSCOA = '" & wADRESSCOA _
        & "' AND ADRESSTYP = '1'"
Set rsADO = cnADO.Execute(xSql)

If Not rsADO.EOF Then
    V = srvZADRESS0_GetBuffer_ODBC(rsADO, lYADRESS0)
    If Not IsNull(V) Then
        srvYADRESS0_Client = "srvYADRESS0_Client_1 : Lecture ZADRESS0 : " & V
        Exit Function
    Else
        blnOk = True
    End If
End If

If Not blnOk Then
    'SINON Lecture Adresse avec type= '  '
    '=================================
    If wADRESSCOA <> "  " Then
        xSql = "select * from ZADRESS0 where ADRESSNUM = '" & wADRESSNUM & "' AND  ADRESSCOA = '  '" _
                & " AND ADRESSTYP = '1'"
        Set rsADO = cnADO.Execute(xSql)
        If Not rsADO.EOF Then
            V = srvZADRESS0_GetBuffer_ODBC(rsADO, lYADRESS0)
            If Not IsNull(V) Then
                srvYADRESS0_Client = "srvYADRESS0_Client_2 : Lecture ZADRESS0 : " & V
                Exit Function
            Else
                blnOk = True
            End If
        End If
    End If
End If

If Not blnOk Then
    srvYADRESS0_Client = "Adresse non trouvée  : " & wADRESSNUM
Else
    If Trim(lYADRESS0.ADRESSRA1) = "" Then
        'Lecture CLIENT principal ==> ADRESSRA1
        '=================================
        xSql = "select * from ZCLIENA0 where CLIENACLI = '" & Trim(wADRESSNUM) & "'"
        Set rsADO = cnADO.Execute(xSql)
        If Not rsADO.EOF Then
            V = srvYCLIENA0_GetBuffer_ODBC(rsADO, wYCLIENA0)
            If Not IsNull(V) Then
                srvYADRESS0_Client = "srvYADRESS0_Client_3 : Lecture ZCLIENT0 : " & V
                Exit Function
            Else
                lYADRESS0.ADRESSRA1 = wYCLIENA0.CLIENARA1
                lYADRESS0.ADRESSRA2 = wYCLIENA0.CLIENARA2
            End If
        End If
    End If
End If

Exit Function

Error_Handler:
srvYADRESS0_Client = Error

End Function


Public Function srvYADRESS0_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle


recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = constYADRESS0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    lX = meYbase.Text
    If mId$(lX, 1, 8) >= YBIATAB0_DATE_CPT_J Then
        srvYADRESS0_Import = Null
        Exit Function
    Else
        meYbase.Method = constDelete
        Call tableYBase_Update(meYbase)
    End If
End If




srvYADRESS0_Import = "?"

paramYADRESS0_Import = paramYBase_DataF & Trim(constYADRESS0) & paramYBase_Data_ExtensionP

Open Trim(paramYADRESS0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYADRESS0) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYADRESS0
            meYbase.K1 = mId$(xIn, 6, 1) & mId$(xIn, 11, 22)
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYADRESS0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = constYADRESS0
meYbase.Text = YBIATAB0_DATE_CPT_J & "_" & DSys & "_" & time_Hms & "_" & Format$(Nb, "000000000")
lX = meYbase.Text
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYADRESS0_Import" & xIn, vbCritical, Error
Close

srvYADRESS0_Import = Error
End Function

Public Function srvYADRESS0_Import_Read(lId As String, lYADRESS0 As typeYADRESS0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYADRESS0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYADRESS0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYADRESS0_GetBuffer lYADRESS0
    srvYADRESS0_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYADRESS0_Import_Read" & xIn, vbCritical, Error
srvYADRESS0_Import_Read = Error
End Function


Public Function srvYADRESS0_Read(lYADRESS0 As typeYADRESS0, blnODBC As Boolean)

Dim xSql As String
Dim V
Dim rsADO As New ADODB.Recordset

On Error GoTo Error_Handler

srvYADRESS0_Read = Null

If Not blnODBC Then
 'Lecture YBASE
'===============
   xSql = lYADRESS0.ADRESSTYP & lYADRESS0.ADRESSNUM & lYADRESS0.ADRESSCOA
    srvYADRESS0_Read = srvYADRESS0_Import_Read(xSql, lYADRESS0)
Else
'Lecture ODBC
'===============
    Set rsADO = Nothing
    xSql = "select * from ZADRESS0 where ADRESSNUM = '" & lYADRESS0.ADRESSNUM & "' AND ADRESSCOA = '" & lYADRESS0.ADRESSCOA & "'"
    rsADO.Open xSql, paramODBC_DSN_SAB
    If rsADO.EOF Then
        V = "Adresse inconnue"
    Else
        V = srvZADRESS0_GetBuffer_ODBC(rsADO, lYADRESS0)
    End If
    
    If Not IsNull(V) Then
        srvYADRESS0_Read = "Lecture ZADRESS0 : " & V
        Exit Function
    End If
    rsADO.Close
End If

Exit Function

Error_Handler:
srvYADRESS0_Read = Error
     
End Function






'-----------------------------------------------------
Function srvYADRESS0_Update(recYADRESS0 As typeYADRESS0)
'-----------------------------------------------------

srvYADRESS0_Update = "?"

MsgTxtLen = 0
Call srvYADRESS0_PutBuffer(recYADRESS0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYADRESS0_GetBuffer(recYADRESS0)) Then
        Call srvYADRESS0_Error(recYADRESS0)
        srvYADRESS0_Update = recYADRESS0.Err
        Exit Function
    Else
        srvYADRESS0_Update = Null
    End If
Else
    recYADRESS0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYADRESS0_Error(recYADRESS0 As typeYADRESS0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YADRESS0" & Chr$(10) & Chr$(13)

Select Case mId$(recYADRESS0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYADRESS0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " & recYADRESS0.ADRESSNUM _
        , I, "module : YADRESS0s.bas  ( " & Trim(recYADRESS0.Obj) & " : " & Trim(recYADRESS0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYADRESS0_Monitor(recYADRESS0 As typeYADRESS0)
'-----------------------------------------------------

arrYADRESS0_Suite = False
Select Case mId$(Trim(recYADRESS0.Method), 1, 4)
    Case "Snap"
              srvYADRESS0_Monitor = srvYADRESS0_Snap(recYADRESS0)
    Case Else
            srvYADRESS0_Monitor = srvYADRESS0_Seek(recYADRESS0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYADRESS0_GetBuffer(recYADRESS0 As typeYADRESS0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYADRESS0_GetBuffer = Null
recYADRESS0.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYADRESS0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYADRESS0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYADRESS0.Err = Space$(10) Then
    recYADRESS0.ADRESSETA = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYADRESS0.ADRESSTYP = mId$(MsgTxt, K + 6, 1)
    recYADRESS0.ADRESSPLA = CLng(Val(mId$(MsgTxt, K + 7, 4)))
    recYADRESS0.ADRESSNUM = mId$(MsgTxt, K + 11, 20)
    recYADRESS0.ADRESSCOA = mId$(MsgTxt, K + 31, 2)
    recYADRESS0.ADRESSDLI = CLng(Val(mId$(MsgTxt, K + 33, 8)))
    recYADRESS0.ADRESSDDE = CLng(Val(mId$(MsgTxt, K + 41, 8)))
    recYADRESS0.ADRESSRA1 = mId$(MsgTxt, K + 49, 32)
    recYADRESS0.ADRESSRA2 = mId$(MsgTxt, K + 81, 32)
    recYADRESS0.ADRESSAD1 = mId$(MsgTxt, K + 113, 32)
    recYADRESS0.ADRESSAD2 = mId$(MsgTxt, K + 145, 32)
    recYADRESS0.ADRESSAD3 = mId$(MsgTxt, K + 177, 32)
    recYADRESS0.ADRESSCOP = mId$(MsgTxt, K + 209, 6)
    recYADRESS0.ADRESSVIL = mId$(MsgTxt, K + 215, 25)
    recYADRESS0.ADRESSPAY = mId$(MsgTxt, K + 240, 25)
    recYADRESS0.ADRESSTEL = mId$(MsgTxt, K + 265, 20)
    recYADRESS0.ADRESSFAX = mId$(MsgTxt, K + 285, 20)
    recYADRESS0.ADRESSTEX = mId$(MsgTxt, K + 305, 20)

Else
    srvYADRESS0_GetBuffer = recYADRESS0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYADRESS0Len

End Function

Public Function srvYADRESS0_Import_Old(lNb As Long)

Dim xIn As String, X As String
Dim meMVTP0 As typeMvtP0

On Error GoTo Error_Handle

srvYADRESS0_Import_Old = "?"

paramYADRESS0_Import = paramYBase_DataF & Trim(constYADRESS0) & paramYBase_Data_ExtensionP

Open Trim(paramYADRESS0_Import) For Input As #1

lNb = 0

recMvtP0_Init meMVTP0
meMVTP0.Method = constAddNew

mdbMvtP0.tableMvtP0_Open

Do Until EOF(1)
    lNb = lNb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meMVTP0.ID = constYADRESS0 & mId$(xIn, 6, 1) & mId$(xIn, 11, 22)
            meMVTP0.Text = xIn
            dbMvtP0_Update meMVTP0
            
    End If
        
Loop


Close
srvYADRESS0_Import_Old = Null
Exit Function

Error_Handle:
 MsgBox "erreur : srvYADRESS0_Import_Old" & xIn, vbCritical, Error
Close

srvYADRESS0_Import_Old = Error
End Function

Public Sub srvYADRESS0_ElpDisplay(recYADRESS0 As typeYADRESS0)
frmElpDisplay.fgData.Rows = 19
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Etablissement"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSTYP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "1 client , 2 compte            3 adresse electroni"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSTYP
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSPLA    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Numéro de plan"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSPLA
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSNUM   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ou numéro de client"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSNUM
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSCOA    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Code adresse"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSCOA
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSDLI    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Date limite validité"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSDLI
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSDDE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Date début validité"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSDDE
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSRA1   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ou raison sociale 1            ou numéro de telex/"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSRA1
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSRA2   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ou raison sociale 2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSRA2
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSAD1   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Adresse 1"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSAD1
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSAD2   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Adresse 2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSAD2
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSAD3   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Adresse 3"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSAD3
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSCOP    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Code postal"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSCOP
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSVIL   25A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Ville"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSVIL
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSPAY   25A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Pays"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSPAY
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSTEL   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "No Tel."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSTEL
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSFAX   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "No Fax."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSFAX
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ADRESSTEX   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "No Télex"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYADRESS0.ADRESSTEX
frmElpDisplay.Show vbModal
End Sub

'---------------------------------------------------------
Public Sub srvYADRESS0_PutBuffer(recYADRESS0 As typeYADRESS0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYADRESS0.Obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYADRESS0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYADRESS0.ADRESSETA, "0000 ")
    Mid$(MsgTxt, K + 6, 1) = recYADRESS0.ADRESSTYP
    Mid$(MsgTxt, K + 7, 4) = Format$(recYADRESS0.ADRESSPLA, "000 ")
    Mid$(MsgTxt, K + 11, 20) = recYADRESS0.ADRESSNUM
    Mid$(MsgTxt, K + 31, 2) = recYADRESS0.ADRESSCOA
    Mid$(MsgTxt, K + 33, 8) = Format$(recYADRESS0.ADRESSDLI, "0000000 ")
    Mid$(MsgTxt, K + 41, 8) = Format$(recYADRESS0.ADRESSDDE, "0000000 ")
    Mid$(MsgTxt, K + 49, 32) = recYADRESS0.ADRESSRA1
    Mid$(MsgTxt, K + 81, 32) = recYADRESS0.ADRESSRA2
    Mid$(MsgTxt, K + 113, 32) = recYADRESS0.ADRESSAD1
    Mid$(MsgTxt, K + 145, 32) = recYADRESS0.ADRESSAD2
    Mid$(MsgTxt, K + 177, 32) = recYADRESS0.ADRESSAD3
    Mid$(MsgTxt, K + 209, 6) = recYADRESS0.ADRESSCOP
    Mid$(MsgTxt, K + 215, 25) = recYADRESS0.ADRESSVIL
    Mid$(MsgTxt, K + 240, 25) = recYADRESS0.ADRESSPAY
    Mid$(MsgTxt, K + 265, 20) = recYADRESS0.ADRESSTEL
    Mid$(MsgTxt, K + 285, 20) = recYADRESS0.ADRESSFAX
    Mid$(MsgTxt, K + 305, 20) = recYADRESS0.ADRESSTEX

MsgTxtLen = MsgTxtLen + recYADRESS0Len
End Sub



'---------------------------------------------------------
Private Function srvYADRESS0_Seek(recYADRESS0 As typeYADRESS0)
'---------------------------------------------------------

srvYADRESS0_Seek = "?"
MsgTxtLen = 0
Call srvYADRESS0_PutBuffer(recYADRESS0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYADRESS0_GetBuffer(recYADRESS0)) Then
        srvYADRESS0_Seek = Null
    Else
        Call srvYADRESS0_Error(recYADRESS0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYADRESS0_Snap(recYADRESS0 As typeYADRESS0)
'---------------------------------------------------------
srvYADRESS0_Snap = "?"
MsgTxtLen = 0
Call srvYADRESS0_PutBuffer(recYADRESS0)
Call srvYADRESS0_PutBuffer(arrYADRESS0(0))
If IsNull(SndRcv()) Then
    srvYADRESS0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYADRESS0_GetBuffer(recYADRESS0)) Then
            Call arrYADRESS0_AddItem(recYADRESS0)
            arrYADRESS0_Suite = True
        Else
            arrYADRESS0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYADRESS0_AddItem(recYADRESS0 As typeYADRESS0)
'---------------------------------------------------------
          
arrYADRESS0_NB = arrYADRESS0_NB + 1
    
If arrYADRESS0_NB > arrYADRESS0_NBMax Then
    arrYADRESS0_NBMax = arrYADRESS0_NBMax + recYADRESS0_Block
    ReDim Preserve arrYADRESS0(arrYADRESS0_NBMax)
End If
            
arrYADRESS0(arrYADRESS0_NB) = recYADRESS0
End Sub



'---------------------------------------------------------
Public Sub recYADRESS0_Init(recYADRESS0 As typeYADRESS0)
'---------------------------------------------------------
recYADRESS0.Obj = "ZADRESS0_S"
recYADRESS0.Method = ""
recYADRESS0.Err = ""
recYADRESS0.ADRESSETA = 0       ' Integer                        ' Etablissement
recYADRESS0.ADRESSTYP = ""      ' String * 1                     ' 1 client , 2 compte
recYADRESS0.ADRESSPLA = 0       ' Long                           ' Numéro de plan
recYADRESS0.ADRESSNUM = ""      ' String * 20                    ' ou numéro de client
recYADRESS0.ADRESSCOA = ""      ' String * 2                     ' Code adresse
recYADRESS0.ADRESSDLI = 0       ' Long                           ' Date limite validité
recYADRESS0.ADRESSDDE = 0       ' Long                           ' Date début validité
recYADRESS0.ADRESSRA1 = ""      ' String * 32                    ' ou raison sociale 1
recYADRESS0.ADRESSRA2 = ""      ' String * 32                    ' ou raison sociale 2
recYADRESS0.ADRESSAD1 = ""      ' String * 32                    ' Adresse 1
recYADRESS0.ADRESSAD2 = ""      ' String * 32                    ' Adresse 2
recYADRESS0.ADRESSAD3 = ""      ' String * 32                    ' Adresse 3
recYADRESS0.ADRESSCOP = ""      ' String * 6                     ' Code postal
recYADRESS0.ADRESSVIL = ""      ' String * 25                    ' Ville
recYADRESS0.ADRESSPAY = ""      ' String * 25                    ' Pays
recYADRESS0.ADRESSTEL = ""      ' String * 20                    ' No Tel.
recYADRESS0.ADRESSFAX = ""      ' String * 20                    ' No Fax.
recYADRESS0.ADRESSTEX = ""      ' String * 20                    ' No Télex

End Sub









Public Function srvYADRESS0_BIC(lCLIENACLI As String)
Dim xId As String, xYADRESS0 As typeYADRESS0
xId = "4 " & lCLIENACLI
If IsNull(srvYADRESS0_Import_Read(xId, xYADRESS0)) Then
    srvYADRESS0_BIC = mId$(xYADRESS0.ADRESSRA1, 11, 11)
Else
    srvYADRESS0_BIC = ""
End If

End Function
