Attribute VB_Name = "srvYCDOREG0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCDOREG0Len = 368 ' 34 + 334
Public Const recYCDOREG0_Block = 200
Public Const constYCDOREG0 = "YCDOREG0"
Dim meYbase As typeYBase
Dim paramYCDOREG0_Import As String

Type typeYCDOREG0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    CDOREGETB       As Integer                        ' CODE ETABLISSEMENT
    CDOREGAGE       As Integer                        ' AGENCE
    CDOREGSER       As String * 2                     ' SERVICE
    CDOREGSSE       As String * 2                     ' SOUS-SERVICE
    CDOREGCOP       As String * 3                     ' CODE OPERATION
    CDOREGDOS       As Long                           ' NUMERO DOSSIER
    CDOREGNUR       As Long                           ' N° RENOUVELLEMENT
    CDOREGUTI       As Long                           ' N° UTILISATION
    CDOREGPAI       As Long                           ' N° PAIEMENT
    CDOREGREG       As Long                           ' N° REGLEMENT/ENCAIS
    CDOREGCRD       As String * 1                     ' CREDIT /DEBIT C/D
    CDOREGMON       As Currency                       ' MONTANT DEV. UTILIS
    CDOREGMOR       As Currency                       ' MONTANT REGLE/ENCAI
    CDOREGDEV       As String * 3                     ' DEVISE REGLEM/ENCAI
    CDOREGDRE       As Long                           ' DATE REGLEM/ENCAIS.
    CDOREGDEM       As Long                           ' DATE EMISSION
    CDOREGDCR       As Long                           ' DATE COMPTA REG/ENC
    CDOREGDUT       As Long                           ' DATE COMPTA UTILISA
    CDOREGRES       As Long                           ' REFERENCE ESCOMPTE
    CDOREGRRE       As Long                           ' REFERENCE REFINANC.
    CDOREGDEC       As String * 1                     ' DESTINAT. CLI/TIERS
    CDOREGDES       As String * 7                     ' DESTINATAIRE
    CDOREGMOD       As String * 3                     ' MODE REGLEMENT/ENCA
    CDOREGINT       As String * 1                     ' CPT NOSTRO  (O/N)
    CDOREGCOM       As String * 20                    ' COMPTE
    CDOREGINC       As String * 1                     ' INTERMED. CLI/TIERS
    CDOREGINS       As String * 7                     ' INTERMEDIAIRE
    CDOREGPAC       As String * 1                     ' BANQ DEST CLI/TIERS
    CDOREGPAS       As String * 7                     ' BANQ. DEST.-PAYEUR
    CDOREGENV       As Long                           ' DATE ENVOI COURRIER
    CDOREGCOU       As Double                         ' COURS DEVREG/DEVDOS
    CDOREGDEN       As Long                           ' DATE ENGAGEMENT
    CDOREGDRP       As Long                           ' DATE RECEP.PREVUE
    CDOREGDRR       As Long                           ' DATE RECEP.REELLE
    CDOREGDAE       As Long                           ' DATE ECHEANCE
    CDOREGDVA       As Long                           ' DATE VALEUR
    CDOREGDIC       As Long                           ' DATE INIT CHANGE
    CDOREGBDF       As String * 3                     ' CODE BDF
    CDOREGPAY       As String * 3                     ' CODE PAYS
    CDOREGSIR       As String * 9                     ' N°SIREN
    CDOREGTRN       As String * 16                    ' TRN SAGITTAIRE
    CDOREGTCR       As String * 1                     ' TYPE CRP
    CDOREGCBA       As Long                           ' CODE BANQUE
    CDOREGCGU       As Long                           ' CODE GUICHET
    CDOREGATG       As String * 1                     ' ATTENTE GEST.
    CDOREGVA1       As Integer                        ' 1ER VALIDEUR
    CDOREGVA2       As Integer                        ' 2EME VALIDEUR
    CDOREGEVE       As String * 2                     ' EVENEMENT
    CDOREGATT       As String * 2                     ' ATTENTE
    CDOREGETA       As String * 2                     ' ETAT
    CDOREGNUA       As Long                           ' NUM OPE ATT
    CDOREGCAA       As String * 12                    ' CODE AUTOR AVAL
    CDOREGCER       As String * 1                     ' COTAT°(O=CERTAIN/N)
    
End Type
    
'---------------------------------------------------------
Public Function srvYCDOREG0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCDOREG0 As typeYCDOREG0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCDOREG0_GetBuffer_ODBC = Null

    recYCDOREG0.CDOREGETB = rsADO("CDOREGETB")    'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOREG0.CDOREGAGE = rsADO("CDOREGAGE")    'CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDOREG0.CDOREGSER = rsADO("CDOREGSER")    'mId$(MsgTxt, K + 11, 2)
    recYCDOREG0.CDOREGSSE = rsADO("CDOREGSSE")    'mId$(MsgTxt, K + 13, 2)
    recYCDOREG0.CDOREGCOP = rsADO("CDOREGCOP")    'mId$(MsgTxt, K + 15, 3)
    recYCDOREG0.CDOREGDOS = rsADO("CDOREGDOS")    'CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDOREG0.CDOREGNUR = rsADO("CDOREGNUR")    'CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDOREG0.CDOREGUTI = rsADO("CDOREGUTI")    'CLng(Val(mId$(MsgTxt, K + 32, 6)))
    recYCDOREG0.CDOREGPAI = rsADO("CDOREGPAI")    'CLng(Val(mId$(MsgTxt, K + 38, 2)))
    recYCDOREG0.CDOREGREG = rsADO("CDOREGREG")    'CLng(Val(mId$(MsgTxt, K + 40, 4)))
    recYCDOREG0.CDOREGCRD = rsADO("CDOREGCRD")    'mId$(MsgTxt, K + 44, 1)
    recYCDOREG0.CDOREGMON = rsADO("CDOREGMON")    'CCur(Val(mId$(MsgTxt, K + 45, 16))) / 100
    recYCDOREG0.CDOREGMOR = rsADO("CDOREGMOR")    'CCur(Val(mId$(MsgTxt, K + 61, 16))) / 100
    recYCDOREG0.CDOREGDEV = rsADO("CDOREGDEV")    'mId$(MsgTxt, K + 77, 3)
    recYCDOREG0.CDOREGDRE = rsADO("CDOREGDRE")    'CLng(Val(mId$(MsgTxt, K + 80, 8)))
    recYCDOREG0.CDOREGDEM = rsADO("CDOREGDEM")    'CLng(Val(mId$(MsgTxt, K + 88, 8)))
    recYCDOREG0.CDOREGDCR = rsADO("CDOREGDCR")    'CLng(Val(mId$(MsgTxt, K + 96, 8)))
    recYCDOREG0.CDOREGDUT = rsADO("CDOREGDUT")    'CLng(Val(mId$(MsgTxt, K + 104, 8)))
    recYCDOREG0.CDOREGRES = rsADO("CDOREGRES")    'CLng(Val(mId$(MsgTxt, K + 112, 10)))
    recYCDOREG0.CDOREGRRE = rsADO("CDOREGRRE")    'CLng(Val(mId$(MsgTxt, K + 122, 10)))
    recYCDOREG0.CDOREGDEC = rsADO("CDOREGDEC")    'mId$(MsgTxt, K + 132, 1)
    recYCDOREG0.CDOREGDES = rsADO("CDOREGDES")    'mId$(MsgTxt, K + 133, 7)
    recYCDOREG0.CDOREGMOD = rsADO("CDOREGMOD")    'mId$(MsgTxt, K + 140, 3)
    recYCDOREG0.CDOREGINT = rsADO("CDOREGINT")    'mId$(MsgTxt, K + 143, 1)
    recYCDOREG0.CDOREGCOM = rsADO("CDOREGCOM")    'mId$(MsgTxt, K + 144, 20)
    recYCDOREG0.CDOREGINC = rsADO("CDOREGINC")    'mId$(MsgTxt, K + 164, 1)
    recYCDOREG0.CDOREGINS = rsADO("CDOREGINS")    'mId$(MsgTxt, K + 165, 7)
    recYCDOREG0.CDOREGPAC = rsADO("CDOREGPAC")    'mId$(MsgTxt, K + 172, 1)
    recYCDOREG0.CDOREGPAS = rsADO("CDOREGPAS")    'mId$(MsgTxt, K + 173, 7)
    recYCDOREG0.CDOREGENV = rsADO("CDOREGENV")    'CLng(Val(mId$(MsgTxt, K + 180, 8)))
    recYCDOREG0.CDOREGCOU = rsADO("CDOREGCOU")    'CDbl(Val(mId$(MsgTxt, K + 188, 15))) / 1000000000
    recYCDOREG0.CDOREGDEN = rsADO("CDOREGDEN")    'CLng(Val(mId$(MsgTxt, K + 203, 8)))
    recYCDOREG0.CDOREGDRP = rsADO("CDOREGDRP")    'CLng(Val(mId$(MsgTxt, K + 211, 8)))
    recYCDOREG0.CDOREGDRR = rsADO("CDOREGDRR")    'CLng(Val(mId$(MsgTxt, K + 219, 8)))
    recYCDOREG0.CDOREGDAE = rsADO("CDOREGDAE")    'CLng(Val(mId$(MsgTxt, K + 227, 8)))
    recYCDOREG0.CDOREGDVA = rsADO("CDOREGDVA")    'CLng(Val(mId$(MsgTxt, K + 235, 8)))
    recYCDOREG0.CDOREGDIC = rsADO("CDOREGDIC")    'CLng(Val(mId$(MsgTxt, K + 243, 8)))
    recYCDOREG0.CDOREGBDF = rsADO("CDOREGBDF")    'mId$(MsgTxt, K + 251, 3)
    recYCDOREG0.CDOREGPAY = rsADO("CDOREGPAY")    'mId$(MsgTxt, K + 254, 3)
    recYCDOREG0.CDOREGSIR = rsADO("CDOREGSIR")    'mId$(MsgTxt, K + 257, 9)
    recYCDOREG0.CDOREGTRN = rsADO("CDOREGTRN")    'mId$(MsgTxt, K + 266, 16)
    recYCDOREG0.CDOREGTCR = rsADO("CDOREGTCR")    'mId$(MsgTxt, K + 282, 1)
    recYCDOREG0.CDOREGCBA = rsADO("CDOREGCBA")    'CLng(Val(mId$(MsgTxt, K + 283, 6)))
    recYCDOREG0.CDOREGCGU = rsADO("CDOREGCGU")    'CLng(Val(mId$(MsgTxt, K + 289, 6)))
    recYCDOREG0.CDOREGATG = rsADO("CDOREGATG")    'mId$(MsgTxt, K + 295, 1)
    recYCDOREG0.CDOREGVA1 = rsADO("CDOREGVA1")    'CInt(Val(mId$(MsgTxt, K + 296, 5)))
    recYCDOREG0.CDOREGVA2 = rsADO("CDOREGVA2")    'CInt(Val(mId$(MsgTxt, K + 301, 5)))
    recYCDOREG0.CDOREGEVE = rsADO("CDOREGEVE")    'mId$(MsgTxt, K + 306, 2)
    recYCDOREG0.CDOREGATT = rsADO("CDOREGATT")    'mId$(MsgTxt, K + 308, 2)
    recYCDOREG0.CDOREGETA = rsADO("CDOREGETA")    'mId$(MsgTxt, K + 310, 2)
    recYCDOREG0.CDOREGNUA = rsADO("CDOREGNUA")    'CLng(Val(mId$(MsgTxt, K + 312, 10)))
    recYCDOREG0.CDOREGCAA = rsADO("CDOREGCAA")    'mId$(MsgTxt, K + 322, 12)
    recYCDOREG0.CDOREGCER = rsADO("CDOREGCER")    'mId$(MsgTxt, K + 334, 1)

Exit Function

Error_Handler:
srvYCDOREG0_GetBuffer_ODBC = Error

End Function



Public Function srvYCDOREG0_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOREG0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    srvYCDOREG0_Import = Null
    lX = CStr(meYbase.Text)
    Exit Function
End If


srvYCDOREG0_Import = "?"

paramYCDOREG0_Import = paramYBase_DataF & Trim(constYCDOREG0) & paramYBase_Data_ExtensionP

Open Trim(paramYCDOREG0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYCDOREG0) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYCDOREG0
            meYbase.K1 = mId$(xIn, 15, 29) 'recYCDOREG0.CDODOSCOP & recYCDOREG0.CDODOSDOS .........
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYCDOREG0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOREG0
lX = DSys & "_" & time_Hms & "_" & Nb
meYbase.Text = lX
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOREG0_Import" & xIn, vbCritical, Error
Close

srvYCDOREG0_Import = Error
End Function

Public Function srvYCDOREG0_Import_Read(lId As String, lYCDOREG0 As typeYCDOREG0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYCDOREG0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYCDOREG0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYCDOREG0_GetBuffer lYCDOREG0
    srvYCDOREG0_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOREG0_Import_Read" & xIn, vbCritical, Error
srvYCDOREG0_Import_Read = Error
End Function





'-----------------------------------------------------
Public Function srvYCDOREG0_Monitor(recYCDOREG0 As typeYCDOREG0)
'-----------------------------------------------------

Select Case mId$(Trim(recYCDOREG0.Method), 1, 4)
    Case "Seek"
                srvYCDOREG0_Monitor = srvYCDOREG0_Seek(recYCDOREG0)
    Case Else
                recYCDOREG0.Err = recYCDOREG0.Method
                Call srvYCDOREG0_Error(recYCDOREG0)
                srvYCDOREG0_Monitor = recYCDOREG0.Err
End Select

End Function

'-----------------------------------------------------
Sub srvYCDOREG0_Error(recYCDOREG0 As typeYCDOREG0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YCDOREG0" & Chr$(10) & Chr$(13)

Select Case mId$(recYCDOREG0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYCDOREG0.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : YCDOREG0s.bas  ( " _
                & Trim(recYCDOREG0.obj) & " : " & Trim(recYCDOREG0.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvYCDOREG0_GetBuffer(recYCDOREG0 As typeYCDOREG0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYCDOREG0_GetBuffer = Null
recYCDOREG0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYCDOREG0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYCDOREG0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYCDOREG0.Err = Space$(10) Then

    recYCDOREG0.CDOREGETB = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOREG0.CDOREGAGE = CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDOREG0.CDOREGSER = mId$(MsgTxt, K + 11, 2)
    recYCDOREG0.CDOREGSSE = mId$(MsgTxt, K + 13, 2)
    recYCDOREG0.CDOREGCOP = mId$(MsgTxt, K + 15, 3)
    recYCDOREG0.CDOREGDOS = CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDOREG0.CDOREGNUR = CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDOREG0.CDOREGUTI = CLng(Val(mId$(MsgTxt, K + 32, 6)))
    recYCDOREG0.CDOREGPAI = CLng(Val(mId$(MsgTxt, K + 38, 2)))
    recYCDOREG0.CDOREGREG = CLng(Val(mId$(MsgTxt, K + 40, 4)))
    recYCDOREG0.CDOREGCRD = mId$(MsgTxt, K + 44, 1)
    recYCDOREG0.CDOREGMON = CCur(Val(mId$(MsgTxt, K + 45, 16))) / 100
    recYCDOREG0.CDOREGMOR = CCur(Val(mId$(MsgTxt, K + 61, 16))) / 100
    recYCDOREG0.CDOREGDEV = mId$(MsgTxt, K + 77, 3)
    recYCDOREG0.CDOREGDRE = CLng(Val(mId$(MsgTxt, K + 80, 8)))
    recYCDOREG0.CDOREGDEM = CLng(Val(mId$(MsgTxt, K + 88, 8)))
    recYCDOREG0.CDOREGDCR = CLng(Val(mId$(MsgTxt, K + 96, 8)))
    recYCDOREG0.CDOREGDUT = CLng(Val(mId$(MsgTxt, K + 104, 8)))
    recYCDOREG0.CDOREGRES = CLng(Val(mId$(MsgTxt, K + 112, 10)))
    recYCDOREG0.CDOREGRRE = CLng(Val(mId$(MsgTxt, K + 122, 10)))
    recYCDOREG0.CDOREGDEC = mId$(MsgTxt, K + 132, 1)
    recYCDOREG0.CDOREGDES = mId$(MsgTxt, K + 133, 7)
    recYCDOREG0.CDOREGMOD = mId$(MsgTxt, K + 140, 3)
    recYCDOREG0.CDOREGINT = mId$(MsgTxt, K + 143, 1)
    recYCDOREG0.CDOREGCOM = mId$(MsgTxt, K + 144, 20)
    recYCDOREG0.CDOREGINC = mId$(MsgTxt, K + 164, 1)
    recYCDOREG0.CDOREGINS = mId$(MsgTxt, K + 165, 7)
    recYCDOREG0.CDOREGPAC = mId$(MsgTxt, K + 172, 1)
    recYCDOREG0.CDOREGPAS = mId$(MsgTxt, K + 173, 7)
    recYCDOREG0.CDOREGENV = CLng(Val(mId$(MsgTxt, K + 180, 8)))
    recYCDOREG0.CDOREGCOU = CDbl(Val(mId$(MsgTxt, K + 188, 15))) / 1000000000
    recYCDOREG0.CDOREGDEN = CLng(Val(mId$(MsgTxt, K + 203, 8)))
    recYCDOREG0.CDOREGDRP = CLng(Val(mId$(MsgTxt, K + 211, 8)))
    recYCDOREG0.CDOREGDRR = CLng(Val(mId$(MsgTxt, K + 219, 8)))
    recYCDOREG0.CDOREGDAE = CLng(Val(mId$(MsgTxt, K + 227, 8)))
    recYCDOREG0.CDOREGDVA = CLng(Val(mId$(MsgTxt, K + 235, 8)))
    recYCDOREG0.CDOREGDIC = CLng(Val(mId$(MsgTxt, K + 243, 8)))
    recYCDOREG0.CDOREGBDF = mId$(MsgTxt, K + 251, 3)
    recYCDOREG0.CDOREGPAY = mId$(MsgTxt, K + 254, 3)
    recYCDOREG0.CDOREGSIR = mId$(MsgTxt, K + 257, 9)
    recYCDOREG0.CDOREGTRN = mId$(MsgTxt, K + 266, 16)
    recYCDOREG0.CDOREGTCR = mId$(MsgTxt, K + 282, 1)
    recYCDOREG0.CDOREGCBA = CLng(Val(mId$(MsgTxt, K + 283, 6)))
    recYCDOREG0.CDOREGCGU = CLng(Val(mId$(MsgTxt, K + 289, 6)))
    recYCDOREG0.CDOREGATG = mId$(MsgTxt, K + 295, 1)
    recYCDOREG0.CDOREGVA1 = CInt(Val(mId$(MsgTxt, K + 296, 5)))
    recYCDOREG0.CDOREGVA2 = CInt(Val(mId$(MsgTxt, K + 301, 5)))
    recYCDOREG0.CDOREGEVE = mId$(MsgTxt, K + 306, 2)
    recYCDOREG0.CDOREGATT = mId$(MsgTxt, K + 308, 2)
    recYCDOREG0.CDOREGETA = mId$(MsgTxt, K + 310, 2)
    recYCDOREG0.CDOREGNUA = CLng(Val(mId$(MsgTxt, K + 312, 10)))
    recYCDOREG0.CDOREGCAA = mId$(MsgTxt, K + 322, 12)
    recYCDOREG0.CDOREGCER = mId$(MsgTxt, K + 334, 1)

Else
    srvYCDOREG0_GetBuffer = recYCDOREG0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYCDOREG0Len

End Function

'---------------------------------------------------------
Private Sub srvYCDOREG0_PutBuffer(recYCDOREG0 As typeYCDOREG0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recYCDOREG0Len) = Space$(recYCDOREG0Len)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCDOREG0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCDOREG0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
    Mid$(MsgTxt, K + 1, 5) = Format$(recYCDOREG0.CDOREGETB, "0000 ")
    Mid$(MsgTxt, K + 6, 5) = Format$(recYCDOREG0.CDOREGAGE, "0000 ")
    Mid$(MsgTxt, K + 11, 2) = recYCDOREG0.CDOREGSER
    Mid$(MsgTxt, K + 13, 2) = recYCDOREG0.CDOREGSSE
    Mid$(MsgTxt, K + 15, 3) = recYCDOREG0.CDOREGCOP
    Mid$(MsgTxt, K + 18, 10) = Format$(recYCDOREG0.CDOREGDOS, "000000000 ")
    Mid$(MsgTxt, K + 28, 4) = Format$(recYCDOREG0.CDOREGNUR, "000 ")
    Mid$(MsgTxt, K + 32, 6) = Format$(recYCDOREG0.CDOREGUTI, "00000 ")
    Mid$(MsgTxt, K + 38, 2) = Format$(recYCDOREG0.CDOREGPAI, "0 ")
    Mid$(MsgTxt, K + 40, 4) = Format$(recYCDOREG0.CDOREGREG, "000 ")
    Mid$(MsgTxt, K + 44, 1) = recYCDOREG0.CDOREGCRD
    Mid$(MsgTxt, K + 45, 16) = Format$(recYCDOREG0.CDOREGMON * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 61, 16) = Format$(recYCDOREG0.CDOREGMOR * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 77, 3) = recYCDOREG0.CDOREGDEV
    Mid$(MsgTxt, K + 80, 8) = Format$(recYCDOREG0.CDOREGDRE, "0000000 ")
    Mid$(MsgTxt, K + 88, 8) = Format$(recYCDOREG0.CDOREGDEM, "0000000 ")
    Mid$(MsgTxt, K + 96, 8) = Format$(recYCDOREG0.CDOREGDCR, "0000000 ")
    Mid$(MsgTxt, K + 104, 8) = Format$(recYCDOREG0.CDOREGDUT, "0000000 ")
    Mid$(MsgTxt, K + 112, 10) = Format$(recYCDOREG0.CDOREGRES, "000000000 ")
    Mid$(MsgTxt, K + 122, 10) = Format$(recYCDOREG0.CDOREGRRE, "000000000 ")
    Mid$(MsgTxt, K + 132, 1) = recYCDOREG0.CDOREGDEC
    Mid$(MsgTxt, K + 133, 7) = recYCDOREG0.CDOREGDES
    Mid$(MsgTxt, K + 140, 3) = recYCDOREG0.CDOREGMOD
    Mid$(MsgTxt, K + 143, 1) = recYCDOREG0.CDOREGINT
    Mid$(MsgTxt, K + 144, 20) = recYCDOREG0.CDOREGCOM
    Mid$(MsgTxt, K + 164, 1) = recYCDOREG0.CDOREGINC
    Mid$(MsgTxt, K + 165, 7) = recYCDOREG0.CDOREGINS
    Mid$(MsgTxt, K + 172, 1) = recYCDOREG0.CDOREGPAC
    Mid$(MsgTxt, K + 173, 7) = recYCDOREG0.CDOREGPAS
    Mid$(MsgTxt, K + 180, 8) = Format$(recYCDOREG0.CDOREGENV, "0000000 ")
    Mid$(MsgTxt, K + 188, 15) = Format$(recYCDOREG0.CDOREGCOU * 1000000000, "00000000000000 ")
    Mid$(MsgTxt, K + 203, 8) = Format$(recYCDOREG0.CDOREGDEN, "0000000 ")
    Mid$(MsgTxt, K + 211, 8) = Format$(recYCDOREG0.CDOREGDRP, "0000000 ")
    Mid$(MsgTxt, K + 219, 8) = Format$(recYCDOREG0.CDOREGDRR, "0000000 ")
    Mid$(MsgTxt, K + 227, 8) = Format$(recYCDOREG0.CDOREGDAE, "0000000 ")
    Mid$(MsgTxt, K + 235, 8) = Format$(recYCDOREG0.CDOREGDVA, "0000000 ")
    Mid$(MsgTxt, K + 243, 8) = Format$(recYCDOREG0.CDOREGDIC, "0000000 ")
    Mid$(MsgTxt, K + 251, 3) = recYCDOREG0.CDOREGBDF
    Mid$(MsgTxt, K + 254, 3) = recYCDOREG0.CDOREGPAY
    Mid$(MsgTxt, K + 257, 9) = recYCDOREG0.CDOREGSIR
    Mid$(MsgTxt, K + 266, 16) = recYCDOREG0.CDOREGTRN
    Mid$(MsgTxt, K + 282, 1) = recYCDOREG0.CDOREGTCR
    Mid$(MsgTxt, K + 283, 6) = Format$(recYCDOREG0.CDOREGCBA, "00000 ")
    Mid$(MsgTxt, K + 289, 6) = Format$(recYCDOREG0.CDOREGCGU, "00000 ")
    Mid$(MsgTxt, K + 295, 1) = recYCDOREG0.CDOREGATG
    Mid$(MsgTxt, K + 296, 5) = Format$(recYCDOREG0.CDOREGVA1, "0000 ")
    Mid$(MsgTxt, K + 301, 5) = Format$(recYCDOREG0.CDOREGVA2, "0000 ")
    Mid$(MsgTxt, K + 306, 2) = recYCDOREG0.CDOREGEVE
    Mid$(MsgTxt, K + 308, 2) = recYCDOREG0.CDOREGATT
    Mid$(MsgTxt, K + 310, 2) = recYCDOREG0.CDOREGETA
    Mid$(MsgTxt, K + 312, 10) = Format$(recYCDOREG0.CDOREGNUA, "000000000 ")
    Mid$(MsgTxt, K + 322, 12) = recYCDOREG0.CDOREGCAA
    Mid$(MsgTxt, K + 334, 1) = recYCDOREG0.CDOREGCER


End Sub


Public Sub srvYCDOREG0_ElpDisplay(recYCDOREG0 As typeYCDOREG0)
frmElpDisplay.fgData.Rows = 54
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGETB
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGSER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGSER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGSSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGSSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGCOP    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGCOP
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGDOS    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGDOS
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGNUR    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° RENOUVELLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGNUR
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGUTI    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° UTILISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGUTI
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGPAI    1P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° PAIEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGPAI
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGREG    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° REGLEMENT/ENCAIS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGREG
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGCRD    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CREDIT /DEBIT C/D"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGCRD
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGMON 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT DEV. UTILIS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGMON
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGMOR 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT REGLE/ENCAI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGMOR
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGDEV    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE REGLEM/ENCAI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGDEV
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGDRE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE REGLEM/ENCAIS."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGDRE
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGDEM    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE EMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGDEM
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGDCR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE COMPTA REG/ENC"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGDCR
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGDUT    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE COMPTA UTILISA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGDUT
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGRES    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REFERENCE ESCOMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGRES
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGRRE    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REFERENCE REFINANC."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGRRE
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGDEC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DESTINAT. CLI/TIERS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGDEC
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGDES    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DESTINATAIRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGDES
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGMOD    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MODE REGLEMENT/ENCA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGMOD
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGINT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CPT NOSTRO  (O/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGINT
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGCOM   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGCOM
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGINC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTERMED. CLI/TIERS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGINC
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGINS    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTERMEDIAIRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGINS
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGPAC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BANQ DEST CLI/TIERS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGPAC
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGPAS    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BANQ. DEST.-PAYEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGPAS
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGENV    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE ENVOI COURRIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGENV
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGCOU 14.9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COURS DEVREG/DEVDOS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGCOU
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGDEN    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE ENGAGEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGDEN
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGDRP    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE RECEP.PREVUE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGDRP
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGDRR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE RECEP.REELLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGDRR
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGDAE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE ECHEANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGDAE
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGDVA    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE VALEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGDVA
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGDIC    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE INIT CHANGE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGDIC
frmElpDisplay.fgData.Row = 38
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGBDF    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE BDF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGBDF
frmElpDisplay.fgData.Row = 39
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGPAY    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE PAYS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGPAY
frmElpDisplay.fgData.Row = 40
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGSIR    9A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N°SIREN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGSIR
frmElpDisplay.fgData.Row = 41
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGTRN   16A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TRN SAGITTAIRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGTRN
frmElpDisplay.fgData.Row = 42
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGTCR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE CRP"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGTCR
frmElpDisplay.fgData.Row = 43
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGCBA    5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE BANQUE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGCBA
frmElpDisplay.fgData.Row = 44
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGCGU    5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE GUICHET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGCGU
frmElpDisplay.fgData.Row = 45
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGATG    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ATTENTE GEST."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGATG
frmElpDisplay.fgData.Row = 46
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGVA1    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "1ER VALIDEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGVA1
frmElpDisplay.fgData.Row = 47
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGVA2    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "2EME VALIDEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGVA2
frmElpDisplay.fgData.Row = 48
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGEVE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EVENEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGEVE
frmElpDisplay.fgData.Row = 49
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGATT    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ATTENTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGATT
frmElpDisplay.fgData.Row = 50
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGETA    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETAT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGETA
frmElpDisplay.fgData.Row = 51
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGNUA    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUM OPE ATT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGNUA
frmElpDisplay.fgData.Row = 52
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGCAA   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE AUTOR AVAL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGCAA
frmElpDisplay.fgData.Row = 53
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOREGCER    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COTAT°(O=CERTAIN/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOREG0.CDOREGCER
frmElpDisplay.Show vbModal
End Sub

'---------------------------------------------------------
Private Function srvYCDOREG0_Seek(recYCDOREG0 As typeYCDOREG0)
'---------------------------------------------------------

srvYCDOREG0_Seek = "?"
MsgTxtLen = 0
Call srvYCDOREG0_PutBuffer(recYCDOREG0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvYCDOREG0_GetBuffer(recYCDOREG0)) Then
            srvYCDOREG0_Seek = Null
        Else
            Call srvYCDOREG0_Error(recYCDOREG0)
        End If
    End If
End If

End Function
'-----------------------------------------------------
Function srvYCDOREG0_Update(recYCDOREG0 As typeYCDOREG0)
'-----------------------------------------------------

srvYCDOREG0_Update = "?"

MsgTxtLen = 0
Call srvYCDOREG0_PutBuffer(recYCDOREG0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYCDOREG0_GetBuffer(recYCDOREG0)) Then
        Call srvYCDOREG0_Error(recYCDOREG0)
        srvYCDOREG0_Update = recYCDOREG0.Err
        Exit Function
    Else
        srvYCDOREG0_Update = Null
    End If
Else
    recYCDOREG0.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recYCDOREG0_Init(recYCDOREG0 As typeYCDOREG0)
'---------------------------------------------------------
MsgTxt = Space$(recYCDOREG0Len)
MsgTxtIndex = 0
Call srvYCDOREG0_GetBuffer(recYCDOREG0)
recYCDOREG0.obj = "ZCDODOS0_S"

recYCDOREG0.CDOREGETB = 0     'As Integer                        ' CODE ETABLISSEMENT
recYCDOREG0.CDOREGAGE = 0     'As Integer                        ' AGENCE
recYCDOREG0.CDOREGSER = ""     'As String * 2                     ' SERVICE
recYCDOREG0.CDOREGSSE = ""     'As String * 2                     ' SOUS-SERVICE
recYCDOREG0.CDOREGCOP = ""     'As String * 3                     ' CODE OPERATION
recYCDOREG0.CDOREGDOS = 0     'As Long                           ' NUMERO DOSSIER
recYCDOREG0.CDOREGNUR = 0     'As Long                           ' N° RENOUVELLEMENT
recYCDOREG0.CDOREGUTI = 0     'As Long                           ' N° UTILISATION
recYCDOREG0.CDOREGPAI = 0     'As Long                           ' N° PAIEMENT
recYCDOREG0.CDOREGREG = 0     'As Long                           ' N° REGLEMENT/ENCAIS
recYCDOREG0.CDOREGCRD = ""     'As String * 1                     ' CREDIT /DEBIT C/D
recYCDOREG0.CDOREGMON = 0     'As Currency                       ' MONTANT DEV. UTILIS
recYCDOREG0.CDOREGMOR = 0     'As Currency                       ' MONTANT REGLE/ENCAI
recYCDOREG0.CDOREGDEV = ""     'As String * 3                     ' DEVISE REGLEM/ENCAI
recYCDOREG0.CDOREGDRE = 0     'As Long                           ' DATE REGLEM/ENCAIS.
recYCDOREG0.CDOREGDEM = 0     'As Long                           ' DATE EMISSION
recYCDOREG0.CDOREGDCR = 0     'As Long                           ' DATE COMPTA REG/ENC
recYCDOREG0.CDOREGDUT = 0     'As Long                           ' DATE COMPTA UTILISA
recYCDOREG0.CDOREGRES = 0     'As Long                           ' REFERENCE ESCOMPTE
recYCDOREG0.CDOREGRRE = 0     'As Long                           ' REFERENCE REFINANC.
recYCDOREG0.CDOREGDEC = ""     'As String * 1                     ' DESTINAT. CLI/TIERS
recYCDOREG0.CDOREGDES = ""     'As String * 7                     ' DESTINATAIRE
recYCDOREG0.CDOREGMOD = ""     'As String * 3                     ' MODE REGLEMENT/ENCA
recYCDOREG0.CDOREGINT = ""     'As String * 1                     ' CPT NOSTRO  (O/N)
recYCDOREG0.CDOREGCOM = ""     'As String * 20                    ' COMPTE
recYCDOREG0.CDOREGINC = ""     'As String * 1                     ' INTERMED. CLI/TIERS
recYCDOREG0.CDOREGINS = ""     'As String * 7                     ' INTERMEDIAIRE
recYCDOREG0.CDOREGPAC = ""     'As String * 1                     ' BANQ DEST CLI/TIERS
recYCDOREG0.CDOREGPAS = ""     'As String * 7                     ' BANQ. DEST.-PAYEUR
recYCDOREG0.CDOREGENV = 0     'As Long                           ' DATE ENVOI COURRIER
recYCDOREG0.CDOREGCOU = 0     'As Double                         ' COURS DEVREG/DEVDOS
recYCDOREG0.CDOREGDEN = 0     'As Long                           ' DATE ENGAGEMENT
recYCDOREG0.CDOREGDRP = 0     'As Long                           ' DATE RECEP.PREVUE
recYCDOREG0.CDOREGDRR = 0     'As Long                           ' DATE RECEP.REELLE
recYCDOREG0.CDOREGDAE = 0     'As Long                           ' DATE ECHEANCE
recYCDOREG0.CDOREGDVA = 0     'As Long                           ' DATE VALEUR
recYCDOREG0.CDOREGDIC = 0     'As Long                           ' DATE INIT CHANGE
recYCDOREG0.CDOREGBDF = ""     'As String * 3                     ' CODE BDF
recYCDOREG0.CDOREGPAY = ""     'As String * 3                     ' CODE PAYS
recYCDOREG0.CDOREGSIR = ""     'As String * 9                     ' N°SIREN
recYCDOREG0.CDOREGTRN = ""     'As String * 16                    ' TRN SAGITTAIRE
recYCDOREG0.CDOREGTCR = ""     'As String * 1                     ' TYPE CRP
recYCDOREG0.CDOREGCBA = 0     'As Long                           ' CODE BANQUE
recYCDOREG0.CDOREGCGU = 0     'As Long                           ' CODE GUICHET
recYCDOREG0.CDOREGATG = ""     'As String * 1                     ' ATTENTE GEST.
recYCDOREG0.CDOREGVA1 = 0     'As Integer                        ' 1ER VALIDEUR
recYCDOREG0.CDOREGVA2 = 0     'As Integer                        ' 2EME VALIDEUR
recYCDOREG0.CDOREGEVE = ""     'As String * 2                     ' EVENEMENT
recYCDOREG0.CDOREGATT = ""     'As String * 2                     ' ATTENTE
recYCDOREG0.CDOREGETA = ""     'As String * 2                     ' ETAT
recYCDOREG0.CDOREGNUA = 0     'As Long                           ' NUM OPE ATT
recYCDOREG0.CDOREGCAA = ""     'As String * 12                    ' CODE AUTOR AVAL
recYCDOREG0.CDOREGCER = ""     'As String * 1                     ' COTAT°(O=CERTAIN/N)
End Sub






