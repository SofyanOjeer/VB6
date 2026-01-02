Attribute VB_Name = "srvYCLIENA0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCLIENA0Len = 255 ' 34 +221
Public Const recYCLIENA0_Block = 50
Public Const memoYCLIENA0Len = 221
Public Const constYCLIENA0 = "YCLIENA0  "
Public paramYCLIENA0_Import As String
Dim meYbase As typeYBase

Type typeYCLIENA0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    CLIENAETB       As Integer                        ' CODE ETABLISSEMENT
    CLIENACLI       As String * 7                     ' NUMERO CLIENT
    CLIENAAGE       As Integer                        ' CODE AGENCE
    CLIENAETA       As String * 4                     ' CODE ETAT
    CLIENARA1       As String * 32                    ' NOM OU DESIGNATION
    CLIENARA2       As String * 32                    ' PRENOM/DESIGNATION
    CLIENASIG       As String * 12                    ' SIGLE USUEL
    CLIENASRN       As String * 9                     ' NUMERO SIREN
    CLIENASRT       As Long                           ' NUMERO SIRET
    CLIENADNA       As Long                           ' DATE DE NAISSANCE
    CLIENAREG       As String * 6                     ' SECT ACTIVITE REGLEM
    CLIENANAT       As String * 3                     ' CDE PAYS NATIONALITE
    CLIENARSD       As String * 3                     ' CDE PAYS DE RESIDENC
    CLIENARES       As String * 3                     ' RESPONS/EXPLOITATION
    CLIENAECO       As String * 3                     ' QUALITE/AG ECONOMIQU
    CLIENAACT       As String * 1                     ' COTE ACTIVITE
    CLIENAPAI       As String * 1                     ' COTE PAIEMENT
    CLIENACRD       As String * 1                     ' COTE CREDIT
    CLIENAADM       As String * 1                     ' COTE ADMISSION
    CLIENAATR       As Long                           ' DAT ATRIB/COTAT BDF
    CLIENABIL       As Long                           ' AN DERN BIL COMM BDF
    CLIENACAT       As String * 3                     ' CATEGORIE CLIENT
    CLIENACOT       As String * 3                     ' COTATION INTERNE
    CLIENACHQ       As String * 1                     ' INTERDICTION CHEQUIE
    CLIENADAT       As Long                           ' INTERDIT CHEQUIER
    CLIENASAC       As String * 6                     ' SECTEUR D ACTIVITE
    CLIENAGEO       As String * 3                     ' SECTEUR GEOGRAPHIQUE
    CLIENAENT       As String * 3                     ' ENTREPRISE LIEE
    CLIENAMES       As String * 1                     ' LANGUE MESSAGERIE
    CLIENAPAY       As Long                           ' DATE ENTREE AU PAYS
    CLIENAFIL       As String * 32                    ' NOM DE JEUNE FILLE
    CLIENABIM       As Long                           ' BILAN DE MOIS
    CLIENADOU       As String * 1                     ' CLIENT DOUTEUX O/N
    CLIENALI1       As String * 3                     ' ZONE LIBRE DE 3 CAR.
    CLIENALI2       As String * 2                     ' ZONE LIBRE DE 2 CAR.
    CLIENAEXT       As String * 32                    ' EXTENTION DU NOM
    CLIENACOL       As String * 1                     ' 0=CLI/COLL=1/AUTRE=2
    CLIENATIE       As String * 7                     ' TIERS DE REFERENCE
    CLIENASEL       As String * 3                     ' CODE SELECTION
    CLIENAPCS       As String * 4                     ' CODE PCS
    CLIENACRE       As Long                           ' DATE CREATION
End Type
    
    
Public arrYCLIENA0() As typeYCLIENA0
Public arrYCLIENA0_NB As Integer
Public arrYCLIENA0_NBMax As Integer
Public arrYCLIENA0_Index As Integer
Public arrYCLIENA0_Suite As Boolean

Dim meMVTP0 As typeMvtP0

Public Function srvYCLIENA0_Read(lYCLIENA0 As typeYCLIENA0, blnODBC As Boolean)

Dim xSQL As String
Dim V
Dim rsADO As New ADODB.Recordset

On Error GoTo Error_Handler

srvYCLIENA0_Read = Null

If Not blnODBC Then
 'Lecture YBASE
'===============
   xSQL = lYCLIENA0.CLIENACLI
    srvYCLIENA0_Read = srvYCLIENA0_Import_Read_YBase(xSQL, lYCLIENA0)
Else
'Lecture ODBC
'===============
    Set rsADO = Nothing
    xSQL = "select * from ZCLIENA0 where CLIENACLI = '" & lYCLIENA0.CLIENACLI & "'"
    rsADO.Open xSQL, paramODBC_DSN_SAB
    If rsADO.EOF Then
        V = "ClientA  inconnu"
    Else
        V = srvYCLIENA0_GetBuffer_ODBC(rsADO, lYCLIENA0)
    End If
    
    If Not IsNull(V) Then
        srvYCLIENA0_Read = "Lecture ZCLIENA0 : " & V
        Exit Function
    End If
    rsADO.Close
End If

Exit Function

Error_Handler:
srvYCLIENA0_Read = Error
     
End Function

'---------------------------------------------------------
Public Function srvYCLIENA0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCLIENA0 As typeYCLIENA0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCLIENA0_GetBuffer_ODBC = Null

    recYCLIENA0.CLIENAETB = rsADO("CLIENAETB")    'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCLIENA0.CLIENACLI = rsADO("CLIENACLI")    'mId$(MsgTxt, K + 6, 7)
    recYCLIENA0.CLIENAAGE = rsADO("CLIENAAGE")    'CInt(Val(mId$(MsgTxt, K + 13, 5)))
    recYCLIENA0.CLIENAETA = rsADO("CLIENAETA")    'mId$(MsgTxt, K + 18, 4)
    recYCLIENA0.CLIENARA1 = rsADO("CLIENARA1")    'mId$(MsgTxt, K + 22, 32)
    recYCLIENA0.CLIENARA2 = rsADO("CLIENARA2")    'mId$(MsgTxt, K + 54, 32)
    recYCLIENA0.CLIENASIG = rsADO("CLIENASIG")    'mId$(MsgTxt, K + 86, 12)
    recYCLIENA0.CLIENASRN = rsADO("CLIENASRN")    'mId$(MsgTxt, K + 98, 9)
    recYCLIENA0.CLIENASRT = rsADO("CLIENASRT")    'CLng(Val(mId$(MsgTxt, K + 107, 6)))
    recYCLIENA0.CLIENADNA = rsADO("CLIENADNA")    'CLng(Val(mId$(MsgTxt, K + 113, 8)))
    recYCLIENA0.CLIENAREG = rsADO("CLIENAREG")    'mId$(MsgTxt, K + 121, 6)
    recYCLIENA0.CLIENANAT = rsADO("CLIENANAT")    'mId$(MsgTxt, K + 127, 3)
    recYCLIENA0.CLIENARSD = rsADO("CLIENARSD")    'mId$(MsgTxt, K + 130, 3)
    recYCLIENA0.CLIENARES = rsADO("CLIENARES")    'mId$(MsgTxt, K + 133, 3)
    recYCLIENA0.CLIENAECO = rsADO("CLIENAECO")    'mId$(MsgTxt, K + 136, 3)
    recYCLIENA0.CLIENAACT = rsADO("CLIENAACT")    'mId$(MsgTxt, K + 139, 1)
    recYCLIENA0.CLIENAPAI = rsADO("CLIENAPAI")    'mId$(MsgTxt, K + 140, 1)
    recYCLIENA0.CLIENACRD = rsADO("CLIENACRD")    'mId$(MsgTxt, K + 141, 1)
    recYCLIENA0.CLIENAADM = rsADO("CLIENAADM")    'mId$(MsgTxt, K + 142, 1)
    recYCLIENA0.CLIENAATR = rsADO("CLIENAATR")    'CLng(Val(mId$(MsgTxt, K + 143, 8)))
    recYCLIENA0.CLIENABIL = rsADO("CLIENABIL")    'CLng(Val(mId$(MsgTxt, K + 151, 4)))
    recYCLIENA0.CLIENACAT = rsADO("CLIENACAT")    'mId$(MsgTxt, K + 155, 3)
    recYCLIENA0.CLIENACOT = rsADO("CLIENACOT")    'mId$(MsgTxt, K + 158, 3)
    recYCLIENA0.CLIENACHQ = rsADO("CLIENACHQ")    'mId$(MsgTxt, K + 161, 1)
    recYCLIENA0.CLIENADAT = rsADO("CLIENADAT")    'CLng(Val(mId$(MsgTxt, K + 162, 8)))
    recYCLIENA0.CLIENASAC = rsADO("CLIENASAC")    'mId$(MsgTxt, K + 170, 6)
    recYCLIENA0.CLIENAGEO = rsADO("CLIENAGEO")    'mId$(MsgTxt, K + 176, 3)
    recYCLIENA0.CLIENAENT = rsADO("CLIENAENT")    'mId$(MsgTxt, K + 179, 3)
    recYCLIENA0.CLIENAMES = rsADO("CLIENAMES")    'mId$(MsgTxt, K + 182, 1)
    recYCLIENA0.CLIENAPAY = rsADO("CLIENAPAY")    'CLng(Val(mId$(MsgTxt, K + 183, 8)))
    recYCLIENA0.CLIENAFIL = rsADO("CLIENAFIL")    'mId$(MsgTxt, K + 191, 32)
    recYCLIENA0.CLIENABIM = rsADO("CLIENABIM")    'CLng(Val(mId$(MsgTxt, K + 223, 3)))
    recYCLIENA0.CLIENADOU = rsADO("CLIENADOU")    'mId$(MsgTxt, K + 226, 1)
    recYCLIENA0.CLIENALI1 = rsADO("CLIENALI1")    'mId$(MsgTxt, K + 227, 3)
    recYCLIENA0.CLIENALI2 = rsADO("CLIENALI2")    'mId$(MsgTxt, K + 230, 2)
    recYCLIENA0.CLIENAEXT = rsADO("CLIENAEXT")    'mId$(MsgTxt, K + 232, 32)
    recYCLIENA0.CLIENACOL = rsADO("CLIENACOL")    'mId$(MsgTxt, K + 264, 1)
    recYCLIENA0.CLIENATIE = rsADO("CLIENATIE")    'mId$(MsgTxt, K + 265, 7)
    recYCLIENA0.CLIENASEL = rsADO("CLIENASEL")    'mId$(MsgTxt, K + 272, 3)
    recYCLIENA0.CLIENAPCS = rsADO("CLIENAPCS")    'mId$(MsgTxt, K + 275, 4)
    recYCLIENA0.CLIENACRE = rsADO("CLIENACRE")    'CLng(Val(mId$(MsgTxt, K + 279, 8)))

Exit Function

Error_Handler:
srvYCLIENA0_GetBuffer_ODBC = Error

End Function

'-----------------------------------------------------
Function srvYCLIENA0_Update(recYCLIENA0 As typeYCLIENA0)
'-----------------------------------------------------

srvYCLIENA0_Update = "?"

MsgTxtLen = 0
Call srvYCLIENA0_PutBuffer(recYCLIENA0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYCLIENA0_GetBuffer(recYCLIENA0)) Then
        Call srvYCLIENA0_Error(recYCLIENA0)
        srvYCLIENA0_Update = recYCLIENA0.Err
        Exit Function
    Else
        srvYCLIENA0_Update = Null
    End If
Else
    recYCLIENA0.Err = "srv"
End If
End Function

Public Sub srvYCLIENA0_Export_CSV(lIdFile_Source As Integer, lIdFile_Destination As Integer, loptSelect_CSV_Header As Boolean, lnb As Long)
Dim xIn As String
If loptSelect_CSV_Header Then
    Print #2, "CLIENAETB;CLIENACLI;CLIENAAGE;CLIENAETA;CLIENARA1;CLIENARA2;CLIENASIG;CLIENASRN;CLIENASRT;CLIENADNA;CLIENAREG;CLIENANAT;CLIENARSD;CLIENARES;CLIENAECO;CLIENAACT;CLIENAPAI;CLIENACRD;CLIENAADM;CLIENAATR;CLIENABIL;CLIENACAT;CLIENACOT;CLIENACHQ;CLIENADAT;CLIENASAC;CLIENAGEO;CLIENAENT;CLIENAMES;CLIENAPAY;CLIENAFIL;CLIENABIM;CLIENADOU;CLIENALI1;CLIENALI2;CLIENAEXT;CLIENACOL;CLIENATIE;CLIENASEL;CLIENAPCS;CLIENACRE;"
    Print #2, "CODE ETABLISSEMENT;NUMERO CLIENT;CODE AGENCE;CODE ETAT;NOM OU DESIGNATION;PRENOM/DESIGNATION;SIGLE USUEL;NUMERO SIREN;NUMERO SIRET;DATE DE NAISSANCE;SECT ACTIVITE REGLEM;CDE PAYS NATIONALITE;CDE PAYS DE RESIDENC;RESPONS/EXPLOITATION;QUALITE/AG ECONOMIQU;COTE ACTIVITE;COTE PAIEMENT;COTE CREDIT;COTE ADMISSION;DAT ATRIB/COTAT BDF;AN DERN BIL COMM BDF;CATEGORIE CLIENT;COTATION INTERNE;INTERDICTION CHEQUIE;DATE LIMITE        N;SECTEUR D ACTIVITE;SECTEUR GEOGRAPHIQUE;ENTREPRISE LIEE;LANGUE MESSAGERIE;DATE ENTREE AU PAYS;NOM DE JEUNE FILLE;BILAN DE MOIS;CLIENT DOUTEUX O/N;ZONE LIBRE DE 3 CAR.;ZONE LIBRE DE 2 CAR.;EXTENTION DU NOM;0=CLI/COLL=1/AUTRE=2;TIERS DE REFERENCE;CODE SELECTION;CODE PCS;DATE CREATION;"
    Print #2, ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(lIdFile_Source)
      Line Input #lIdFile_Source, xIn
      lnb = lnb + 1
      Print #lIdFile_Destination, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 7) & ";" & mId$(xIn, 13, 5) & ";" & mId$(xIn, 18, 4) & ";" & mId$(xIn, 22, 32) & ";" & mId$(xIn, 54, 32) & ";" _
      & mId$(xIn, 86, 12) & ";" & mId$(xIn, 98, 9) & ";" & mId$(xIn, 107, 6) & ";" _
      & mId$(xIn, 113, 8) & ";" & mId$(xIn, 121, 6) & ";" & mId$(xIn, 127, 3) & ";" & mId$(xIn, 130, 3) & ";" _
      & mId$(xIn, 133, 3) & ";" & mId$(xIn, 136, 3) & ";" & mId$(xIn, 139, 1) & ";" & mId$(xIn, 140, 1) & ";" & mId$(xIn, 141, 1) & ";" _
      & mId$(xIn, 142, 1) & ";" & mId$(xIn, 143, 8) & ";" & mId$(xIn, 151, 4) & ";" & mId$(xIn, 155, 3) & ";" & mId$(xIn, 158, 3) & ";" _
      & mId$(xIn, 161, 1) & ";" & mId$(xIn, 162, 8) & ";" & mId$(xIn, 170, 6) & ";" & mId$(xIn, 176, 3) & ";" & mId$(xIn, 179, 3) & ";" _
      & mId$(xIn, 182, 1) & ";" & mId$(xIn, 183, 8) & ";" & mId$(xIn, 191, 32) & ";" & mId$(xIn, 223, 3) & ";" _
      & mId$(xIn, 226, 1) & ";" & mId$(xIn, 227, 3) & ";" & mId$(xIn, 230, 2) & ";" & mId$(xIn, 232, 32) & ";" _
      & mId$(xIn, 264, 1) & ";" & mId$(xIn, 265, 7) & ";" & mId$(xIn, 272, 3) & ";" & mId$(xIn, 275, 4) & ";" & mId$(xIn, 279, 8) & ";"
Loop
End Sub

'-----------------------------------------------------
Sub srvYCLIENA0_Error(recYCLIENA0 As typeYCLIENA0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YCLIENA0" & Chr$(10) & Chr$(13)

Select Case mId$(recYCLIENA0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYCLIENA0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " & recYCLIENA0.CLIENACLI _
        , I, "module : YCLIENA0s.bas  ( " & Trim(recYCLIENA0.obj) & " : " & Trim(recYCLIENA0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYCLIENA0_Monitor(recYCLIENA0 As typeYCLIENA0)
'-----------------------------------------------------

arrYCLIENA0_Suite = False
Select Case mId$(Trim(recYCLIENA0.Method), 1, 4)
    Case "Snap"
              srvYCLIENA0_Monitor = srvYCLIENA0_Snap(recYCLIENA0)
    Case Else
            srvYCLIENA0_Monitor = srvYCLIENA0_Seek(recYCLIENA0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYCLIENA0_GetBuffer(recYCLIENA0 As typeYCLIENA0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYCLIENA0_GetBuffer = Null
recYCLIENA0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYCLIENA0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYCLIENA0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYCLIENA0.Err = Space$(10) Then
    recYCLIENA0.CLIENAETB = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCLIENA0.CLIENACLI = mId$(MsgTxt, K + 6, 7)
    recYCLIENA0.CLIENAAGE = CInt(Val(mId$(MsgTxt, K + 13, 5)))
    recYCLIENA0.CLIENAETA = mId$(MsgTxt, K + 18, 4)
    recYCLIENA0.CLIENARA1 = mId$(MsgTxt, K + 22, 32)
    recYCLIENA0.CLIENARA2 = mId$(MsgTxt, K + 54, 32)
    recYCLIENA0.CLIENASIG = mId$(MsgTxt, K + 86, 12)
    recYCLIENA0.CLIENASRN = mId$(MsgTxt, K + 98, 9)
    recYCLIENA0.CLIENASRT = CLng(Val(mId$(MsgTxt, K + 107, 6)))
    recYCLIENA0.CLIENADNA = CLng(Val(mId$(MsgTxt, K + 113, 8)))
    recYCLIENA0.CLIENAREG = mId$(MsgTxt, K + 121, 6)
    recYCLIENA0.CLIENANAT = mId$(MsgTxt, K + 127, 3)
    recYCLIENA0.CLIENARSD = mId$(MsgTxt, K + 130, 3)
    recYCLIENA0.CLIENARES = mId$(MsgTxt, K + 133, 3)
    recYCLIENA0.CLIENAECO = mId$(MsgTxt, K + 136, 3)
    recYCLIENA0.CLIENAACT = mId$(MsgTxt, K + 139, 1)
    recYCLIENA0.CLIENAPAI = mId$(MsgTxt, K + 140, 1)
    recYCLIENA0.CLIENACRD = mId$(MsgTxt, K + 141, 1)
    recYCLIENA0.CLIENAADM = mId$(MsgTxt, K + 142, 1)
    recYCLIENA0.CLIENAATR = CLng(Val(mId$(MsgTxt, K + 143, 8)))
    recYCLIENA0.CLIENABIL = CLng(Val(mId$(MsgTxt, K + 151, 4)))
    recYCLIENA0.CLIENACAT = mId$(MsgTxt, K + 155, 3)
    recYCLIENA0.CLIENACOT = mId$(MsgTxt, K + 158, 3)
    recYCLIENA0.CLIENACHQ = mId$(MsgTxt, K + 161, 1)
    recYCLIENA0.CLIENADAT = CLng(Val(mId$(MsgTxt, K + 162, 8)))
    recYCLIENA0.CLIENASAC = mId$(MsgTxt, K + 170, 6)
    recYCLIENA0.CLIENAGEO = mId$(MsgTxt, K + 176, 3)
    recYCLIENA0.CLIENAENT = mId$(MsgTxt, K + 179, 3)
    recYCLIENA0.CLIENAMES = mId$(MsgTxt, K + 182, 1)
    recYCLIENA0.CLIENAPAY = CLng(Val(mId$(MsgTxt, K + 183, 8)))
    recYCLIENA0.CLIENAFIL = mId$(MsgTxt, K + 191, 32)
    recYCLIENA0.CLIENABIM = CLng(Val(mId$(MsgTxt, K + 223, 3)))
    recYCLIENA0.CLIENADOU = mId$(MsgTxt, K + 226, 1)
    recYCLIENA0.CLIENALI1 = mId$(MsgTxt, K + 227, 3)
    recYCLIENA0.CLIENALI2 = mId$(MsgTxt, K + 230, 2)
    recYCLIENA0.CLIENAEXT = mId$(MsgTxt, K + 232, 32)
    recYCLIENA0.CLIENACOL = mId$(MsgTxt, K + 264, 1)
    recYCLIENA0.CLIENATIE = mId$(MsgTxt, K + 265, 7)
    recYCLIENA0.CLIENASEL = mId$(MsgTxt, K + 272, 3)
    recYCLIENA0.CLIENAPCS = mId$(MsgTxt, K + 275, 4)
    recYCLIENA0.CLIENACRE = CLng(Val(mId$(MsgTxt, K + 279, 8)))

Else
    srvYCLIENA0_GetBuffer = recYCLIENA0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYCLIENA0Len

End Function

Public Function srvYCLIENA0_Import(lnb As Long)
Dim xIn As String, X As String

On Error GoTo Error_Handle


srvYCLIENA0_Import = "?"

paramYCLIENA0_Import = paramYBase_DataF & Trim(constYCLIENA0) & paramYBase_Data_ExtensionP
Open Trim(paramYCLIENA0_Import) For Input As #1

lnb = 0

recMvtP0_Init meMVTP0
meMVTP0.Method = constAddNew

mdbMvtP0.tableMvtP0_Open

Do Until EOF(1)
    lnb = lnb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meMVTP0.ID = constYCLIENA0 & mId$(xIn, 6, 7)
            meMVTP0.Text = xIn
            dbMvtP0_Update meMVTP0
            
    End If
        
Loop


Close
srvYCLIENA0_Import = Null
Exit Function

Error_Handle:
 MsgBox "erreur : srvYCLIENA0_Import" & xIn, vbCritical, Error
Close

srvYCLIENA0_Import = Error
End Function

Public Function srvYCLIENA0_Import_Ybase(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle


recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = constYCLIENA0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    lX = meYbase.Text
    If mId$(lX, 1, 8) >= YBIATAB0_DATE_CPT_J Then
        srvYCLIENA0_Import_Ybase = Null
        Exit Function
    Else
        meYbase.Method = constDelete
        Call tableYBase_Update(meYbase)
    End If
End If




srvYCLIENA0_Import_Ybase = "?"

paramYCLIENA0_Import = paramYBase_DataF & Trim(constYCLIENA0) & paramYBase_Data_ExtensionP

Open Trim(paramYCLIENA0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYCLIENA0) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYCLIENA0
            meYbase.K1 = mId$(xIn, 6, 7)  ' .CLIENACLI
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYCLIENA0_Import_Ybase = Null
meYbase.ID = constYBase
meYbase.K1 = constYCLIENA0
meYbase.Text = YBIATAB0_DATE_CPT_J & "_" & DSys & "_" & time_Hms & "_" & Format$(Nb, "000000000")
lX = meYbase.Text
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYCLIENA0_Import" & xIn, vbCritical, Error
Close

srvYCLIENA0_Import_Ybase = Error
End Function


Public Function srvYCLIENA0_Import_Read_YBase(lId As String, lYCLIENA0 As typeYCLIENA0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYCLIENA0_Import_Read_YBase = "?"

meYbase.Method = "Seek="
meYbase.ID = constYCLIENA0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
        MsgTxt = Space$(34) & meYbase.Text
        MsgTxtIndex = 0
        srvYCLIENA0_GetBuffer lYCLIENA0
        srvYCLIENA0_Import_Read_YBase = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYCLIENA0_Import_Read" & xIn, vbCritical, Error
srvYCLIENA0_Import_Read_YBase = Error
End Function





Public Function srvYCLIENA0_Import_Read(lId As String, lYCLIENA0 As typeYCLIENA0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYCLIENA0_Import_Read = "?"

meMVTP0.Method = "Seek="
meMVTP0.ID = lId
If tableMvtP0_Read(meMVTP0) = 0 Then
    MsgTxt = Space$(34) & meMVTP0.Text
    MsgTxtIndex = 0
    srvYCLIENA0_GetBuffer lYCLIENA0
    srvYCLIENA0_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYCLIENA0_Import_Read" & xIn, vbCritical, Error
Close
srvYCLIENA0_Import_Read = Error
End Function



Public Sub srvYCLIENA0_ElpDisplay(recYCLIENA0 As typeYCLIENA0)
frmElpDisplay.fgData.Rows = 42
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENAETB
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENACLI    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO CLIENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENACLI
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENAAGE
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAETA    4A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETAT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENAETA
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENARA1   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NOM OU DESIGNATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENARA1
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENARA2   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PRENOM/DESIGNATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENARA2
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENASIG   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SIGLE USUEL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENASIG
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENASRN    9A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO SIREN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENASRN
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENASRT    5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO SIRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENASRT
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENADNA    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DE NAISSANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENADNA
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAREG    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SECT ACTIVITE REGLEM"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENAREG
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENANAT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CDE PAYS NATIONALITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENANAT
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENARSD    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CDE PAYS DE RESIDENC"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENARSD
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENARES    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RESPONS/EXPLOITATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENARES
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAECO    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "QUALITE/AG ECONOMIQU"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENAECO
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAACT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COTE ACTIVITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENAACT
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAPAI    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COTE PAIEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENAPAI
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENACRD    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COTE CREDIT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENACRD
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAADM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COTE ADMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENAADM
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAATR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DAT ATRIB/COTAT BDF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENAATR
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENABIL    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AN DERN BIL COMM BDF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENABIL
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENACAT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CATEGORIE CLIENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENACAT
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENACOT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COTATION INTERNE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENACOT
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENACHQ    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTERDICTION CHEQUIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENACHQ
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENADAT    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTERDIT CHEQUIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENADAT
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENASAC    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SECTEUR D ACTIVITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENASAC
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAGEO    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SECTEUR GEOGRAPHIQUE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENAGEO
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAENT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ENTREPRISE LIEE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENAENT
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAMES    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LANGUE MESSAGERIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENAMES
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAPAY    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE ENTREE AU PAYS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENAPAY
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAFIL   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NOM DE JEUNE FILLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENAFIL
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENABIM    2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BILAN DE MOIS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENABIM
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENADOU    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CLIENT DOUTEUX O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENADOU
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENALI1    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ZONE LIBRE DE 3 CAR."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENALI1
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENALI2    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ZONE LIBRE DE 2 CAR."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENALI2
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAEXT   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EXTENTION DU NOM"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENAEXT
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENACOL    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "0=CLI/COLL=1/AUTRE=2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENACOL
frmElpDisplay.fgData.Row = 38
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENATIE    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TIERS DE REFERENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENATIE
frmElpDisplay.fgData.Row = 39
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENASEL    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE SELECTION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENASEL
frmElpDisplay.fgData.Row = 40
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAPCS    4A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE PCS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENAPCS
frmElpDisplay.fgData.Row = 41
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENACRE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE CREATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCLIENA0.CLIENACRE
frmElpDisplay.Show vbModal
End Sub

'---------------------------------------------------------
Public Sub srvYCLIENA0_PutBuffer(recYCLIENA0 As typeYCLIENA0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCLIENA0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCLIENA0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYCLIENA0.CLIENAETB, "0000 ")
    Mid$(MsgTxt, K + 6, 7) = recYCLIENA0.CLIENACLI
    Mid$(MsgTxt, K + 13, 5) = Format$(recYCLIENA0.CLIENAAGE, "0000 ")
    Mid$(MsgTxt, K + 18, 4) = recYCLIENA0.CLIENAETA
    Mid$(MsgTxt, K + 22, 32) = recYCLIENA0.CLIENARA1
    Mid$(MsgTxt, K + 54, 32) = recYCLIENA0.CLIENARA2
    Mid$(MsgTxt, K + 86, 12) = recYCLIENA0.CLIENASIG
    Mid$(MsgTxt, K + 98, 9) = recYCLIENA0.CLIENASRN
    Mid$(MsgTxt, K + 107, 6) = Format$(recYCLIENA0.CLIENASRT, "00000 ")
    Mid$(MsgTxt, K + 113, 8) = Format$(recYCLIENA0.CLIENADNA, "0000000 ")
    Mid$(MsgTxt, K + 121, 6) = recYCLIENA0.CLIENAREG
    Mid$(MsgTxt, K + 127, 3) = recYCLIENA0.CLIENANAT
    Mid$(MsgTxt, K + 130, 3) = recYCLIENA0.CLIENARSD
    Mid$(MsgTxt, K + 133, 3) = recYCLIENA0.CLIENARES
    Mid$(MsgTxt, K + 136, 3) = recYCLIENA0.CLIENAECO
    Mid$(MsgTxt, K + 139, 1) = recYCLIENA0.CLIENAACT
    Mid$(MsgTxt, K + 140, 1) = recYCLIENA0.CLIENAPAI
    Mid$(MsgTxt, K + 141, 1) = recYCLIENA0.CLIENACRD
    Mid$(MsgTxt, K + 142, 1) = recYCLIENA0.CLIENAADM
    Mid$(MsgTxt, K + 143, 8) = Format$(recYCLIENA0.CLIENAATR, "0000000 ")
    Mid$(MsgTxt, K + 151, 4) = Format$(recYCLIENA0.CLIENABIL, "000 ")
    Mid$(MsgTxt, K + 155, 3) = recYCLIENA0.CLIENACAT
    Mid$(MsgTxt, K + 158, 3) = recYCLIENA0.CLIENACOT
    Mid$(MsgTxt, K + 161, 1) = recYCLIENA0.CLIENACHQ
    Mid$(MsgTxt, K + 162, 8) = Format$(recYCLIENA0.CLIENADAT, "0000000 ")
    Mid$(MsgTxt, K + 170, 6) = recYCLIENA0.CLIENASAC
    Mid$(MsgTxt, K + 176, 3) = recYCLIENA0.CLIENAGEO
    Mid$(MsgTxt, K + 179, 3) = recYCLIENA0.CLIENAENT
    Mid$(MsgTxt, K + 182, 1) = recYCLIENA0.CLIENAMES
    Mid$(MsgTxt, K + 183, 8) = Format$(recYCLIENA0.CLIENAPAY, "0000000 ")
    Mid$(MsgTxt, K + 191, 32) = recYCLIENA0.CLIENAFIL
    Mid$(MsgTxt, K + 223, 3) = Format$(recYCLIENA0.CLIENABIM, "00 ")
    Mid$(MsgTxt, K + 226, 1) = recYCLIENA0.CLIENADOU
    Mid$(MsgTxt, K + 227, 3) = recYCLIENA0.CLIENALI1
    Mid$(MsgTxt, K + 230, 2) = recYCLIENA0.CLIENALI2
    Mid$(MsgTxt, K + 232, 32) = recYCLIENA0.CLIENAEXT
    Mid$(MsgTxt, K + 264, 1) = recYCLIENA0.CLIENACOL
    Mid$(MsgTxt, K + 265, 7) = recYCLIENA0.CLIENATIE
    Mid$(MsgTxt, K + 272, 3) = recYCLIENA0.CLIENASEL
    Mid$(MsgTxt, K + 275, 4) = recYCLIENA0.CLIENAPCS
    Mid$(MsgTxt, K + 279, 8) = Format$(recYCLIENA0.CLIENACRE, "0000000 ")

MsgTxtLen = MsgTxtLen + recYCLIENA0Len
End Sub



'---------------------------------------------------------
Private Function srvYCLIENA0_Seek(recYCLIENA0 As typeYCLIENA0)
'---------------------------------------------------------

srvYCLIENA0_Seek = "?"
MsgTxtLen = 0
Call srvYCLIENA0_PutBuffer(recYCLIENA0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYCLIENA0_GetBuffer(recYCLIENA0)) Then
        srvYCLIENA0_Seek = Null
    Else
        Call srvYCLIENA0_Error(recYCLIENA0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYCLIENA0_Snap(recYCLIENA0 As typeYCLIENA0)
'---------------------------------------------------------
srvYCLIENA0_Snap = "?"
MsgTxtLen = 0
Call srvYCLIENA0_PutBuffer(recYCLIENA0)
Call srvYCLIENA0_PutBuffer(arrYCLIENA0(0))
If IsNull(SndRcv()) Then
    srvYCLIENA0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYCLIENA0_GetBuffer(recYCLIENA0)) Then
            Call arrYCLIENA0_AddItem(recYCLIENA0)
            arrYCLIENA0_Suite = True
        Else
            arrYCLIENA0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYCLIENA0_AddItem(recYCLIENA0 As typeYCLIENA0)
'---------------------------------------------------------
          
arrYCLIENA0_NB = arrYCLIENA0_NB + 1
    
If arrYCLIENA0_NB > arrYCLIENA0_NBMax Then
    arrYCLIENA0_NBMax = arrYCLIENA0_NBMax + recYCLIENA0_Block
    ReDim Preserve arrYCLIENA0(arrYCLIENA0_NBMax)
End If
            
arrYCLIENA0(arrYCLIENA0_NB) = recYCLIENA0
End Sub



'---------------------------------------------------------
Public Sub recYCLIENA0_Init(recYCLIENA0 As typeYCLIENA0)
'---------------------------------------------------------
recYCLIENA0.obj = "ZCLIENA0_S"
recYCLIENA0.Method = ""
recYCLIENA0.Err = ""

End Sub







