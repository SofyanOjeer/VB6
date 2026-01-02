Attribute VB_Name = "srvYCDOUTI0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCDOUTI0Len = 285 ' 34 + 251
Public Const recYCDOUTI0_Block = 50
Public Const constYCDOUTI0 = "YCDOUTI0"
Dim meYbase As typeYBase
Dim paramYCDOUTI0_Import As String

Type typeYCDOUTI0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    CDOUTIETB       As Integer                        ' CODE ETABLISSEMENT
    CDOUTIAGE       As Integer                        ' AGENCE
    CDOUTISER       As String * 2                     ' SERVICE
    CDOUTISSE       As String * 2                     ' SOUS-SERVICE
    CDOUTICOP       As String * 3                     ' CODE OPERATION
    CDOUTIDOS       As Long                           ' NUMERO DOSSIER
    CDOUTINUR       As Long                           ' N° RENOUVELLEMENT
    CDOUTIUTI       As Long                           ' N° UTILISATION
    CDOUTITMO       As String * 1                     ' C/N/D
    CDOUTIMON       As Currency                       ' MONTANT UTILISATION
    CDOUTIMAD       As Currency                       ' MONTANT ADDITIONNEL
    CDOUTIMTO       As Currency                       ' MONTANT TOTAL
    CDOUTIMDO       As Currency                       ' MONTANT DOCUMENTS
    CDOUTIMPA       As Currency                       ' MONTANT A PAYER
    CDOUTIPRE       As Long                           ' DATE PREVUE UTILIS.
    CDOUTIDAR       As Long                           ' DATE REFUS DOCUMEN.
    CDOUTIOBJ       As String * 6                     ' OBJET UTILISATION
    CDOUTIMVU       As Currency                       ' MONTANT A VUE
    CDOUTIMCA       As Currency                       ' MONTANT ACCEPTATION
    CDOUTIMDI       As Currency                       ' MONTANT DIFFERE
    CDOUTICTR       As String * 1                     ' REMETTANT (C/T)
    CDOUTIREM       As String * 7                     ' REMETTANT
    CDOUTIRER       As String * 16                    ' REFE.REMETTANT
    CDOUTIDRE       As Long                           ' DATE REMISE (EXP)
    CDOUTIDCO       As String * 1                     ' DOC.CONFORMES (O/N)
    CDOUTIDCE       As String * 1                     ' DOCUMENTS ENVOYES
    CDOUTIRET       As String * 1                     ' RESERVES TRANSMISES
    CDOUTIDAC       As String * 1                     ' DEMANDE ACCORD
    CDOUTIPAR       As String * 1                     ' PAY.SOUS RESERVES
    CDOUTIIRR       As String * 6                     ' IRREGULARITES
    CDOUTIPOR       As String * 1                     ' PORTEFEUILLE
    CDOUTIREF       As String * 1                     ' REFINANCEMENT
    CDOUTIESC       As String * 1                     ' ESCOMPTE
    CDOUTIBEC       As String * 1                     ' BENEF PAY.COMMIS°
    CDOUTIVA1       As Integer                        ' 1ER VALIDEUR
    CDOUTIVA2       As Integer                        ' 2EME VALIDEUR
    CDOUTIEVE       As String * 2                     ' EVENEMENT
    CDOUTIATT       As String * 2                     ' ATTENTE
    CDOUTIETA       As String * 2                     ' ETAT UTILISATION
    
End Type
    
'---------------------------------------------------------
Public Function srvYCDOUTI0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCDOUTI0 As typeYCDOUTI0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCDOUTI0_GetBuffer_ODBC = Null

    recYCDOUTI0.CDOUTIETB = rsADO("CDOUTIETB")    'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOUTI0.CDOUTIAGE = rsADO("CDOUTIAGE")    'CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDOUTI0.CDOUTISER = rsADO("CDOUTISER")    'mId$(MsgTxt, K + 11, 2)
    recYCDOUTI0.CDOUTISSE = rsADO("CDOUTISSE")    'mId$(MsgTxt, K + 13, 2)
    recYCDOUTI0.CDOUTICOP = rsADO("CDOUTICOP")    'mId$(MsgTxt, K + 15, 3)
    recYCDOUTI0.CDOUTIDOS = rsADO("CDOUTIDOS")    'CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDOUTI0.CDOUTINUR = rsADO("CDOUTINUR")    'CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDOUTI0.CDOUTIUTI = rsADO("CDOUTIUTI")    'CLng(Val(mId$(MsgTxt, K + 32, 6)))
    recYCDOUTI0.CDOUTITMO = rsADO("CDOUTITMO")    'mId$(MsgTxt, K + 38, 1)
    recYCDOUTI0.CDOUTIMON = rsADO("CDOUTIMON")    'CCur(Val(mId$(MsgTxt, K + 39, 16))) / 100
    recYCDOUTI0.CDOUTIMAD = rsADO("CDOUTIMAD")    'CCur(Val(mId$(MsgTxt, K + 55, 16))) / 100
    recYCDOUTI0.CDOUTIMTO = rsADO("CDOUTIMTO")    'CCur(Val(mId$(MsgTxt, K + 71, 16))) / 100
    recYCDOUTI0.CDOUTIMDO = rsADO("CDOUTIMDO")    'CCur(Val(mId$(MsgTxt, K + 87, 16))) / 100
    recYCDOUTI0.CDOUTIMPA = rsADO("CDOUTIMPA")    'CCur(Val(mId$(MsgTxt, K + 103, 16))) / 100
    recYCDOUTI0.CDOUTIPRE = rsADO("CDOUTIPRE")    'CLng(Val(mId$(MsgTxt, K + 119, 8)))
    recYCDOUTI0.CDOUTIDAR = rsADO("CDOUTIDAR")    'CLng(Val(mId$(MsgTxt, K + 127, 8)))
    recYCDOUTI0.CDOUTIOBJ = rsADO("CDOUTIOBJ")    'mId$(MsgTxt, K + 135, 6)
    recYCDOUTI0.CDOUTIMVU = rsADO("CDOUTIMVU")    'CCur(Val(mId$(MsgTxt, K + 141, 16))) / 100
    recYCDOUTI0.CDOUTIMCA = rsADO("CDOUTIMCA")    'CCur(Val(mId$(MsgTxt, K + 157, 16))) / 100
    recYCDOUTI0.CDOUTIMDI = rsADO("CDOUTIMDI")    'CCur(Val(mId$(MsgTxt, K + 173, 16))) / 100
    recYCDOUTI0.CDOUTICTR = rsADO("CDOUTICTR")    'mId$(MsgTxt, K + 189, 1)
    recYCDOUTI0.CDOUTIREM = rsADO("CDOUTIREM")    'mId$(MsgTxt, K + 190, 7)
    recYCDOUTI0.CDOUTIRER = rsADO("CDOUTIRER")    'mId$(MsgTxt, K + 197, 16)
    recYCDOUTI0.CDOUTIDRE = rsADO("CDOUTIDRE")    'CLng(Val(mId$(MsgTxt, K + 213, 8)))
    recYCDOUTI0.CDOUTIDCO = rsADO("CDOUTIDCO")    'mId$(MsgTxt, K + 221, 1)
    recYCDOUTI0.CDOUTIDCE = rsADO("CDOUTIDCE")    'mId$(MsgTxt, K + 222, 1)
    recYCDOUTI0.CDOUTIRET = rsADO("CDOUTIRET")    'mId$(MsgTxt, K + 223, 1)
    recYCDOUTI0.CDOUTIDAC = rsADO("CDOUTIDAC")    'mId$(MsgTxt, K + 224, 1)
    recYCDOUTI0.CDOUTIPAR = rsADO("CDOUTIPAR")    'mId$(MsgTxt, K + 225, 1)
    recYCDOUTI0.CDOUTIIRR = rsADO("CDOUTIIRR")    'mId$(MsgTxt, K + 226, 6)
    recYCDOUTI0.CDOUTIPOR = rsADO("CDOUTIPOR")    'mId$(MsgTxt, K + 232, 1)
    recYCDOUTI0.CDOUTIREF = rsADO("CDOUTIREF")    'mId$(MsgTxt, K + 233, 1)
    recYCDOUTI0.CDOUTIESC = rsADO("CDOUTIESC")    'mId$(MsgTxt, K + 234, 1)
    recYCDOUTI0.CDOUTIBEC = rsADO("CDOUTIBEC")    'mId$(MsgTxt, K + 235, 1)
    recYCDOUTI0.CDOUTIVA1 = rsADO("CDOUTIVA1")    'CInt(Val(mId$(MsgTxt, K + 236, 5)))
    recYCDOUTI0.CDOUTIVA2 = rsADO("CDOUTIVA2")    'CInt(Val(mId$(MsgTxt, K + 241, 5)))
    recYCDOUTI0.CDOUTIEVE = rsADO("CDOUTIEVE")    'mId$(MsgTxt, K + 246, 2)
    recYCDOUTI0.CDOUTIATT = rsADO("CDOUTIATT")    'mId$(MsgTxt, K + 248, 2)
    recYCDOUTI0.CDOUTIETA = rsADO("CDOUTIETA")    'mId$(MsgTxt, K + 250, 2)

Exit Function

Error_Handler:
srvYCDOUTI0_GetBuffer_ODBC = Error

End Function

Public Sub srvYCDOUTI0_Export_CSV(lIdFile_Source As Integer, lIdFile_Destination As Integer, loptSelect_CSV_Header As Boolean, lnb As Long)
Dim xIn As String
If loptSelect_CSV_Header Then
    Print #lIdFile_Destination, "CDOUTIETB;CDOUTIAGE;CDOUTISER;CDOUTISSE;CDOUTICOP;CDOUTIDOS;CDOUTINUR;CDOUTIUTI;CDOUTITMO;CDOUTIMON;CDOUTIMAD;CDOUTIMTO;CDOUTIMDO;CDOUTIMPA;CDOUTIPRE;CDOUTIDAR;CDOUTIOBJ;CDOUTIMVU;CDOUTIMCA;CDOUTIMDI;CDOUTICTR;CDOUTIREM;CDOUTIRER;CDOUTIDRE;CDOUTIDCO;CDOUTIDCE;CDOUTIRET;CDOUTIDAC;CDOUTIPAR;CDOUTIIRR;CDOUTIPOR;CDOUTIREF;CDOUTIESC;CDOUTIBEC;CDOUTIVA1;CDOUTIVA2;CDOUTIEVE;CDOUTIATT;CDOUTIETA;"
    Print #lIdFile_Destination, "CODE ETABLISSEMENT;AGENCE;SERVICE;SOUS-SERVICE;CODE OPERATION;NUMERO DOSSIER;N° RENOUVELLEMENT;N° UTILISATION;C/N/D;MONTANT UTILISATION;MONTANT ADDITIONNEL;MONTANT TOTAL;MONTANT DOCUMENTS;MONTANT A PAYER;DATE PREVUE UTILIS.;DATE REFUS DOCUMEN.;OBJET UTILISATION;MONTANT A VUE;MONTANT ACCEPTATION;MONTANT DIFFERE;REMETTANT (C/T);REMETTANT;REFE.REMETTANT;DATE REMISE (EXP);DOC.CONFORMES (O/N);DOCUMENTS ENVOYES;RESERVES TRANSMISES;DEMANDE ACCORD;PAY.SOUS RESERVES;IRREGULARITES;PORTEFEUILLE;REFINANCEMENT;ESCOMPTE;BENEF PAY.COMMIS°;1ER VALIDEUR;2EME VALIDEUR;EVENEMENT;ATTENTE;ETAT UTILISATION;"
    Print #lIdFile_Destination, ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(lIdFile_Source)
      Line Input #lIdFile_Source, xIn
      lnb = lnb + 1
      Print #lIdFile_Destination, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 5) & ";" & mId$(xIn, 11, 2) & ";" & mId$(xIn, 13, 2) & ";" _
      & mId$(xIn, 15, 3) & ";" & mId$(xIn, 18, 10) & ";" & mId$(xIn, 28, 4) & ";" & mId$(xIn, 32, 6) & ";" _
      & mId$(xIn, 38, 1) & ";" _
      & cur_19V(CCur(mId$(xIn, 39, 16)) / 100) & ";" _
      & cur_19V(CCur(mId$(xIn, 55, 16)) / 100) & ";" _
      & cur_19V(CCur(mId$(xIn, 71, 16)) / 100) & ";" _
      & cur_19V(CCur(mId$(xIn, 87, 16)) / 100) & ";" _
      & cur_19V(CCur(mId$(xIn, 103, 16)) / 100) & ";" & mId$(xIn, 119, 8) & ";" & mId$(xIn, 127, 8) & ";" _
      & mId$(xIn, 135, 6) & ";" _
      & cur_19V(CCur(mId$(xIn, 141, 16)) / 100) & ";" _
      & cur_19V(CCur(mId$(xIn, 157, 16)) / 100) & ";" _
      & cur_19V(CCur(mId$(xIn, 173, 16)) / 100) & ";" _
      & mId$(xIn, 189, 1) & ";" & mId$(xIn, 190, 7) & ";" & mId$(xIn, 197, 16) & ";" & mId$(xIn, 213, 8) & ";" _
      & mId$(xIn, 221, 1) & ";" & mId$(xIn, 222, 1) & ";" & mId$(xIn, 223, 1) & ";" _
      & mId$(xIn, 224, 1) & ";" & mId$(xIn, 225, 1) & ";" & mId$(xIn, 226, 6) & ";" & mId$(xIn, 232, 1) & ";" _
      & mId$(xIn, 233, 1) & ";" & mId$(xIn, 234, 1) & ";" & mId$(xIn, 235, 1) & ";" & mId$(xIn, 236, 5) & ";" _
      & mId$(xIn, 241, 5) & ";" & mId$(xIn, 246, 2) & ";" & mId$(xIn, 248, 2) & ";" & mId$(xIn, 250, 2) & ";"
Loop
End Sub


Public Function srvYCDOUTI0_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOUTI0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    srvYCDOUTI0_Import = Null
    lX = CStr(meYbase.Text)
    Exit Function
End If


srvYCDOUTI0_Import = "?"

paramYCDOUTI0_Import = paramYBase_DataF & Trim(constYCDOUTI0) & paramYBase_Data_ExtensionP

Open Trim(paramYCDOUTI0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYCDOUTI0) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYCDOUTI0
            meYbase.K1 = mId$(xIn, 15, 23) 'recYCDOUTI0.CDODOSCOP & recYCDOUTI0.CDODOSDOS .........
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYCDOUTI0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOUTI0
lX = DSys & "_" & time_Hms & "_" & Nb
meYbase.Text = lX
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOUTI0_Import" & xIn, vbCritical, Error
Close

srvYCDOUTI0_Import = Error
End Function

Public Function srvYCDOUTI0_Import_Read(lId As String, lYCDOUTI0 As typeYCDOUTI0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYCDOUTI0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYCDOUTI0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYCDOUTI0_GetBuffer lYCDOUTI0
    srvYCDOUTI0_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOUTI0_Import_Read" & xIn, vbCritical, Error
srvYCDOUTI0_Import_Read = Error
End Function





'-----------------------------------------------------
Public Function srvYCDOUTI0_Monitor(recYCDOUTI0 As typeYCDOUTI0)
'-----------------------------------------------------

Select Case mId$(Trim(recYCDOUTI0.Method), 1, 4)
    Case "Seek"
                srvYCDOUTI0_Monitor = srvYCDOUTI0_Seek(recYCDOUTI0)
    Case Else
                recYCDOUTI0.Err = recYCDOUTI0.Method
                Call srvYCDOUTI0_Error(recYCDOUTI0)
                srvYCDOUTI0_Monitor = recYCDOUTI0.Err
End Select

End Function

'-----------------------------------------------------
Sub srvYCDOUTI0_Error(recYCDOUTI0 As typeYCDOUTI0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YCDOUTI0" & Chr$(10) & Chr$(13)

Select Case mId$(recYCDOUTI0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYCDOUTI0.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : YCDOUTI0s.bas  ( " _
                & Trim(recYCDOUTI0.obj) & " : " & Trim(recYCDOUTI0.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvYCDOUTI0_GetBuffer(recYCDOUTI0 As typeYCDOUTI0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYCDOUTI0_GetBuffer = Null
recYCDOUTI0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYCDOUTI0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYCDOUTI0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYCDOUTI0.Err = Space$(10) Then

    recYCDOUTI0.CDOUTIETB = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOUTI0.CDOUTIAGE = CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDOUTI0.CDOUTISER = mId$(MsgTxt, K + 11, 2)
    recYCDOUTI0.CDOUTISSE = mId$(MsgTxt, K + 13, 2)
    recYCDOUTI0.CDOUTICOP = mId$(MsgTxt, K + 15, 3)
    recYCDOUTI0.CDOUTIDOS = CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDOUTI0.CDOUTINUR = CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDOUTI0.CDOUTIUTI = CLng(Val(mId$(MsgTxt, K + 32, 6)))
    recYCDOUTI0.CDOUTITMO = mId$(MsgTxt, K + 38, 1)
    recYCDOUTI0.CDOUTIMON = CCur(Val(mId$(MsgTxt, K + 39, 16))) / 100
    recYCDOUTI0.CDOUTIMAD = CCur(Val(mId$(MsgTxt, K + 55, 16))) / 100
    recYCDOUTI0.CDOUTIMTO = CCur(Val(mId$(MsgTxt, K + 71, 16))) / 100
    recYCDOUTI0.CDOUTIMDO = CCur(Val(mId$(MsgTxt, K + 87, 16))) / 100
    recYCDOUTI0.CDOUTIMPA = CCur(Val(mId$(MsgTxt, K + 103, 16))) / 100
    recYCDOUTI0.CDOUTIPRE = CLng(Val(mId$(MsgTxt, K + 119, 8)))
    recYCDOUTI0.CDOUTIDAR = CLng(Val(mId$(MsgTxt, K + 127, 8)))
    recYCDOUTI0.CDOUTIOBJ = mId$(MsgTxt, K + 135, 6)
    recYCDOUTI0.CDOUTIMVU = CCur(Val(mId$(MsgTxt, K + 141, 16))) / 100
    recYCDOUTI0.CDOUTIMCA = CCur(Val(mId$(MsgTxt, K + 157, 16))) / 100
    recYCDOUTI0.CDOUTIMDI = CCur(Val(mId$(MsgTxt, K + 173, 16))) / 100
    recYCDOUTI0.CDOUTICTR = mId$(MsgTxt, K + 189, 1)
    recYCDOUTI0.CDOUTIREM = mId$(MsgTxt, K + 190, 7)
    recYCDOUTI0.CDOUTIRER = mId$(MsgTxt, K + 197, 16)
    recYCDOUTI0.CDOUTIDRE = CLng(Val(mId$(MsgTxt, K + 213, 8)))
    recYCDOUTI0.CDOUTIDCO = mId$(MsgTxt, K + 221, 1)
    recYCDOUTI0.CDOUTIDCE = mId$(MsgTxt, K + 222, 1)
    recYCDOUTI0.CDOUTIRET = mId$(MsgTxt, K + 223, 1)
    recYCDOUTI0.CDOUTIDAC = mId$(MsgTxt, K + 224, 1)
    recYCDOUTI0.CDOUTIPAR = mId$(MsgTxt, K + 225, 1)
    recYCDOUTI0.CDOUTIIRR = mId$(MsgTxt, K + 226, 6)
    recYCDOUTI0.CDOUTIPOR = mId$(MsgTxt, K + 232, 1)
    recYCDOUTI0.CDOUTIREF = mId$(MsgTxt, K + 233, 1)
    recYCDOUTI0.CDOUTIESC = mId$(MsgTxt, K + 234, 1)
    recYCDOUTI0.CDOUTIBEC = mId$(MsgTxt, K + 235, 1)
    recYCDOUTI0.CDOUTIVA1 = CInt(Val(mId$(MsgTxt, K + 236, 5)))
    recYCDOUTI0.CDOUTIVA2 = CInt(Val(mId$(MsgTxt, K + 241, 5)))
    recYCDOUTI0.CDOUTIEVE = mId$(MsgTxt, K + 246, 2)
    recYCDOUTI0.CDOUTIATT = mId$(MsgTxt, K + 248, 2)
    recYCDOUTI0.CDOUTIETA = mId$(MsgTxt, K + 250, 2)

Else
    srvYCDOUTI0_GetBuffer = recYCDOUTI0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYCDOUTI0Len

End Function

'---------------------------------------------------------
Private Sub srvYCDOUTI0_PutBuffer(recYCDOUTI0 As typeYCDOUTI0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recYCDOUTI0Len) = Space$(recYCDOUTI0Len)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCDOUTI0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCDOUTI0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYCDOUTI0.CDOUTIETB, "0000 ")
    Mid$(MsgTxt, K + 6, 5) = Format$(recYCDOUTI0.CDOUTIAGE, "0000 ")
    Mid$(MsgTxt, K + 11, 2) = recYCDOUTI0.CDOUTISER
    Mid$(MsgTxt, K + 13, 2) = recYCDOUTI0.CDOUTISSE
    Mid$(MsgTxt, K + 15, 3) = recYCDOUTI0.CDOUTICOP
    Mid$(MsgTxt, K + 18, 10) = Format$(recYCDOUTI0.CDOUTIDOS, "000000000 ")
    Mid$(MsgTxt, K + 28, 4) = Format$(recYCDOUTI0.CDOUTINUR, "000 ")
    Mid$(MsgTxt, K + 32, 6) = Format$(recYCDOUTI0.CDOUTIUTI, "00000 ")
    Mid$(MsgTxt, K + 38, 1) = recYCDOUTI0.CDOUTITMO
    Mid$(MsgTxt, K + 39, 16) = Format$(recYCDOUTI0.CDOUTIMON * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 55, 16) = Format$(recYCDOUTI0.CDOUTIMAD * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 71, 16) = Format$(recYCDOUTI0.CDOUTIMTO * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 87, 16) = Format$(recYCDOUTI0.CDOUTIMDO * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 103, 16) = Format$(recYCDOUTI0.CDOUTIMPA * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 119, 8) = Format$(recYCDOUTI0.CDOUTIPRE, "0000000 ")
    Mid$(MsgTxt, K + 127, 8) = Format$(recYCDOUTI0.CDOUTIDAR, "0000000 ")
    Mid$(MsgTxt, K + 135, 6) = recYCDOUTI0.CDOUTIOBJ
    Mid$(MsgTxt, K + 141, 16) = Format$(recYCDOUTI0.CDOUTIMVU * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 157, 16) = Format$(recYCDOUTI0.CDOUTIMCA * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 173, 16) = Format$(recYCDOUTI0.CDOUTIMDI * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 189, 1) = recYCDOUTI0.CDOUTICTR
    Mid$(MsgTxt, K + 190, 7) = recYCDOUTI0.CDOUTIREM
    Mid$(MsgTxt, K + 197, 16) = recYCDOUTI0.CDOUTIRER
    Mid$(MsgTxt, K + 213, 8) = Format$(recYCDOUTI0.CDOUTIDRE, "0000000 ")
    Mid$(MsgTxt, K + 221, 1) = recYCDOUTI0.CDOUTIDCO
    Mid$(MsgTxt, K + 222, 1) = recYCDOUTI0.CDOUTIDCE
    Mid$(MsgTxt, K + 223, 1) = recYCDOUTI0.CDOUTIRET
    Mid$(MsgTxt, K + 224, 1) = recYCDOUTI0.CDOUTIDAC
    Mid$(MsgTxt, K + 225, 1) = recYCDOUTI0.CDOUTIPAR
    Mid$(MsgTxt, K + 226, 6) = recYCDOUTI0.CDOUTIIRR
    Mid$(MsgTxt, K + 232, 1) = recYCDOUTI0.CDOUTIPOR
    Mid$(MsgTxt, K + 233, 1) = recYCDOUTI0.CDOUTIREF
    Mid$(MsgTxt, K + 234, 1) = recYCDOUTI0.CDOUTIESC
    Mid$(MsgTxt, K + 235, 1) = recYCDOUTI0.CDOUTIBEC
    Mid$(MsgTxt, K + 236, 5) = Format$(recYCDOUTI0.CDOUTIVA1, "0000 ")
    Mid$(MsgTxt, K + 241, 5) = Format$(recYCDOUTI0.CDOUTIVA2, "0000 ")
    Mid$(MsgTxt, K + 246, 2) = recYCDOUTI0.CDOUTIEVE
    Mid$(MsgTxt, K + 248, 2) = recYCDOUTI0.CDOUTIATT
    Mid$(MsgTxt, K + 250, 2) = recYCDOUTI0.CDOUTIETA

End Sub


Public Sub srvYCDOUTI0_ElpDisplay(recYCDOUTI0 As typeYCDOUTI0)
frmElpDisplay.fgData.Rows = 40
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIETB
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTISER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTISER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTISSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTISSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTICOP    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTICOP
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIDOS    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIDOS
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTINUR    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° RENOUVELLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTINUR
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIUTI    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° UTILISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIUTI
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTITMO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "C/N/D"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTITMO
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIMON 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT UTILISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIMON
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIMAD 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT ADDITIONNEL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIMAD
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIMTO 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT TOTAL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIMTO
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIMDO 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT DOCUMENTS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIMDO
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIMPA 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT A PAYER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIMPA
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIPRE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE PREVUE UTILIS."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIPRE
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIDAR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE REFUS DOCUMEN."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIDAR
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIOBJ    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "OBJET UTILISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIOBJ
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIMVU 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT A VUE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIMVU
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIMCA 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT ACCEPTATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIMCA
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIMDI 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT DIFFERE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIMDI
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTICTR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REMETTANT (C/T)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTICTR
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIREM    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REMETTANT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIREM
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIRER   16A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REFE.REMETTANT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIRER
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIDRE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE REMISE (EXP)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIDRE
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIDCO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DOC.CONFORMES (O/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIDCO
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIDCE    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DOCUMENTS ENVOYES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIDCE
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIRET    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RESERVES TRANSMISES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIRET
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIDAC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEMANDE ACCORD"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIDAC
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIPAR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PAY.SOUS RESERVES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIPAR
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIIRR    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "IRREGULARITES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIIRR
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIPOR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PORTEFEUILLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIPOR
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIREF    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REFINANCEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIREF
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIESC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ESCOMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIESC
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIBEC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BENEF PAY.COMMIS°"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIBEC
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIVA1    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "1ER VALIDEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIVA1
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIVA2    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "2EME VALIDEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIVA2
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIEVE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EVENEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIEVE
frmElpDisplay.fgData.Row = 38
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIATT    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ATTENTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIATT
frmElpDisplay.fgData.Row = 39
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOUTIETA    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETAT UTILISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOUTI0.CDOUTIETA
frmElpDisplay.Show vbModal
End Sub

'---------------------------------------------------------
Private Function srvYCDOUTI0_Seek(recYCDOUTI0 As typeYCDOUTI0)
'---------------------------------------------------------

srvYCDOUTI0_Seek = "?"
MsgTxtLen = 0
Call srvYCDOUTI0_PutBuffer(recYCDOUTI0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvYCDOUTI0_GetBuffer(recYCDOUTI0)) Then
            srvYCDOUTI0_Seek = Null
        Else
            Call srvYCDOUTI0_Error(recYCDOUTI0)
        End If
    End If
End If

End Function
'-----------------------------------------------------
Function srvYCDOUTI0_Update(recYCDOUTI0 As typeYCDOUTI0)
'-----------------------------------------------------

srvYCDOUTI0_Update = "?"

MsgTxtLen = 0
Call srvYCDOUTI0_PutBuffer(recYCDOUTI0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYCDOUTI0_GetBuffer(recYCDOUTI0)) Then
        Call srvYCDOUTI0_Error(recYCDOUTI0)
        srvYCDOUTI0_Update = recYCDOUTI0.Err
        Exit Function
    Else
        srvYCDOUTI0_Update = Null
    End If
Else
    recYCDOUTI0.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recYCDOUTI0_Init(recYCDOUTI0 As typeYCDOUTI0)
'---------------------------------------------------------
MsgTxt = Space$(recYCDOUTI0Len)
MsgTxtIndex = 0
Call srvYCDOUTI0_GetBuffer(recYCDOUTI0)
recYCDOUTI0.obj = "ZCDODOS0_S"

recYCDOUTI0.CDOUTIETB = 0     'As Integer                        ' CODE ETABLISSEMENT
recYCDOUTI0.CDOUTIAGE = 0     'As Integer                        ' AGENCE
recYCDOUTI0.CDOUTISER = ""     'As String * 2                     ' SERVICE
recYCDOUTI0.CDOUTISSE = ""     'As String * 2                     ' SOUS-SERVICE
recYCDOUTI0.CDOUTICOP = ""     'As String * 3                     ' CODE OPERATION
recYCDOUTI0.CDOUTIDOS = 0     'As Long                           ' NUMERO DOSSIER
recYCDOUTI0.CDOUTINUR = 0     'As Long                           ' N° RENOUVELLEMENT
recYCDOUTI0.CDOUTIUTI = 0     'As Long                           ' N° UTILISATION
recYCDOUTI0.CDOUTITMO = ""     'As String * 1                     ' C/N/D
recYCDOUTI0.CDOUTIMON = 0     'As Currency                       ' MONTANT UTILISATION
recYCDOUTI0.CDOUTIMAD = 0     'As Currency                       ' MONTANT ADDITIONNEL
recYCDOUTI0.CDOUTIMTO = 0     'As Currency                       ' MONTANT TOTAL
recYCDOUTI0.CDOUTIMDO = 0     'As Currency                       ' MONTANT DOCUMENTS
recYCDOUTI0.CDOUTIMPA = 0     'As Currency                       ' MONTANT A PAYER
recYCDOUTI0.CDOUTIPRE = 0     'As Long                           ' DATE PREVUE UTILIS.
recYCDOUTI0.CDOUTIDAR = 0     'As Long                           ' DATE REFUS DOCUMEN.
recYCDOUTI0.CDOUTIOBJ = ""     'As String * 6                     ' OBJET UTILISATION
recYCDOUTI0.CDOUTIMVU = 0     'As Currency                       ' MONTANT A VUE
recYCDOUTI0.CDOUTIMCA = 0     'As Currency                       ' MONTANT ACCEPTATION
recYCDOUTI0.CDOUTIMDI = 0     'As Currency                       ' MONTANT DIFFERE
recYCDOUTI0.CDOUTICTR = ""     'As String * 1                     ' REMETTANT (C/T)
recYCDOUTI0.CDOUTIREM = ""     'As String * 7                     ' REMETTANT
recYCDOUTI0.CDOUTIRER = ""     'As String * 16                    ' REFE.REMETTANT
recYCDOUTI0.CDOUTIDRE = 0     'As Long                           ' DATE REMISE (EXP)
recYCDOUTI0.CDOUTIDCO = ""     'As String * 1                     ' DOC.CONFORMES (O/N)
recYCDOUTI0.CDOUTIDCE = ""     'As String * 1                     ' DOCUMENTS ENVOYES
recYCDOUTI0.CDOUTIRET = ""     'As String * 1                     ' RESERVES TRANSMISES
recYCDOUTI0.CDOUTIDAC = ""     'As String * 1                     ' DEMANDE ACCORD
recYCDOUTI0.CDOUTIPAR = ""     'As String * 1                     ' PAY.SOUS RESERVES
recYCDOUTI0.CDOUTIIRR = ""     'As String * 6                     ' IRREGULARITES
recYCDOUTI0.CDOUTIPOR = ""     'As String * 1                     ' PORTEFEUILLE
recYCDOUTI0.CDOUTIREF = ""     'As String * 1                     ' REFINANCEMENT
recYCDOUTI0.CDOUTIESC = ""     'As String * 1                     ' ESCOMPTE
recYCDOUTI0.CDOUTIBEC = ""     'As String * 1                     ' BENEF PAY.COMMIS°
recYCDOUTI0.CDOUTIVA1 = 0     'As Integer                        ' 1ER VALIDEUR
recYCDOUTI0.CDOUTIVA2 = 0     'As Integer                        ' 2EME VALIDEUR
recYCDOUTI0.CDOUTIEVE = ""     'As String * 2                     ' EVENEMENT
recYCDOUTI0.CDOUTIATT = ""     'As String * 2                     ' ATTENTE
recYCDOUTI0.CDOUTIETA = ""     'As String * 2                     ' ETAT UTILISATION
End Sub








