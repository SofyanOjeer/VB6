Attribute VB_Name = "srvYAUTE1I0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYAUTE1I0Len = 72 ' 34 +38
Public Const recYAUTE1I0_Block = 100
Public Const memoYAUTE1I0Len = 38
Public Const constYAUTE1I0 = "YAUTE1I0"
Public Const constYAUTE1I0_GRP = "YAUTE1I0_GRP"
Public paramYAUTE1I0_Import As String
Dim meYbase As typeYBase
Dim xYbase As typeYBase

Type typeYAUTE1I0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    AUTE1IETA       As Integer                        ' ETABLISSEMENT
    AUTE1IGRP       As String * 7                     ' GROUPE
    AUTE1ICLI       As String * 7                     ' CLIENT
    AUTE1ITYP       As String * 1                     ' TYPE 1,2,3
    AUTE1IAUT       As String * 20                    ' CODE AUTORISATION
    AUTE1IDEV       As String * 3                     ' DEVISE
    AUTE1IAGE       As Integer                        ' AGENCE
    AUTE1ISER       As String * 2                     ' SERVICE
    AUTE1ISRV       As String * 2                     ' SOUS SERVICE
    AUTE1ICOP       As String * 3                     ' CODE OPERATION
    AUTE1INOP       As Long                           ' NUMERO OPERATION
    AUTE1IOR1       As Long                           ' ORDRE 1
    AUTE1IOR2       As Long                           ' ORDRE 2
    AUTE1IOR3       As Long                           ' ORDRE 3
    AUTE1IOR4       As Long                           ' ORDRE 4
    AUTE1IDBA       As String * 3                     ' DEVISE DE BASE
    AUTE1IMDB       As Long                           ' MONTANT DEBIT
    AUTE1IMCR       As Long                           ' MONTANT CREDIT
    AUTE1IBDB       As Long                           ' MONTANT DEVBAS DB
    AUTE1IBCR       As Long                           ' MONTANT DEVBAS CR
    AUTE1IRDB       As Long                           ' MONTANT REPOR. DB
    AUTE1IRCR       As Long                           ' MONTANT REPOR. CR
    AUTE1IMAU       As Long                           ' MONTANT AUTO
    AUTE1IDAD       As Long                           ' DATE DEBUT AUTO.
    AUTE1IDAF       As Long                           ' DATE FIN AUTO.
    AUTE1IINT       As String * 1                     ' INTITULITE
    AUTE1IDMO       As Long                           ' DATE DERN.MOUV.
    AUTE1IRA1       As String * 32                    ' RAISON SOCIALE
    AUTE1IRA2       As String * 32                    ' RAISON SOCIALE 2
    AUTE1ISAC       As String * 6                     ' SECTEUR ACTIVITE
    AUTE1IREG       As String * 6                     ' SECTEUR ACT. REG.
    AUTE1ISRN       As String * 9                     ' NUMERO SIREN
    AUTE1IRES       As String * 3                     ' RESPON/EXPLOIT
    AUTE1IECO       As String * 3                     ' QUALITE/AG.ECONO
    AUTE1ICOT       As String * 3                     ' COTATION INTERNE
    AUTE1IBDF       As String * 4                     ' CODE BDF
    AUTE1IDOU       As String * 1                     ' DOUTEUX  O/N
    AUTE1IICH       As String * 1                     ' INTERDIT CHQ  O/N
    AUTE1ICET       As String * 4                     ' CODE ETAT
    AUTE1ISIG       As String * 12                    ' SIGLE
    AUTE1IRAG       As String * 32                    ' RAISON SOC GROUPE
    AUTE1IELM       As String * 1                     ' CODE ELEM. O/N
    AUTE1INIV       As Long                           ' NIVEAU
    AUTE1IBLO       As String * 1                     ' CODE BLOCAGE1,2,3
    AUTE1ICOM       As String * 1                     ' COMPENSATION O/N
    AUTE1ILAU       As String * 30                    ' LIBEL AUTO OU GAR
    AUTE1IECI       As Long                           ' ECHEANCE INTERNE
    AUTE1IDEP       As String * 1                     ' CODE DEPASSEMENT
    AUTE1IMTD       As Long                           ' MONTANT DEPASSEM
    AUTE1IDPD       As Long                           ' DATE 1ER DEPAS.
    AUTE1IDTD       As Long                           ' DEPASSEM  DEPUIS
    AUTE1IC1A       As String * 1                     ' C1AUT POUR DEPASS
    AUTE1IDEB       As Long                           ' DATE DEBUT OPERA
    AUTE1IFIN       As Long                           ' DATE FIN OPERA
    AUTE1ILIB       As String * 32                    ' LIBELLE OPERATION
    AUTE1IRAT       As String * 1                     ' RATTAC GROUPE O/N
    AUTE1IATR       As String * 1                     ' AUTO GROUPE(1à9)
    AUTE1IREL       As String * 3                     ' RELATION CLI-GRP
    AUTE1IRUB       As String * 10                    ' RUBRIQUE COMPT.
    AUTE1IAGC       As Integer                        ' AGENCE CLIENT
    AUTE1ISEG       As String * 3                     ' SEGEMENT DE RESULTAT
    AUTE1ISEP       As String * 3                     ' SEGEMENT POTENTIEL
    AUTE1IFUT       As String * 150                   ' ZONE FUTURE
    

End Type
    
Public arrYAUTE1I0() As typeYAUTE1I0
Public arrYAUTE1I0_NB As Integer
Public arrYAUTE1I0_NBMax As Integer
Public arrYAUTE1I0_Index As Integer
Public arrYAUTE1I0_Suite As Boolean

Public Function srvYAUTE1I0_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle


recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = constYAUTE1I0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    lX = meYbase.Text
    If mId$(lX, 1, 8) >= YBIATAB0_DATE_CPT_J Then
        srvYAUTE1I0_Import = Null
        Exit Function
    Else
        meYbase.Method = constDelete
        Call tableYBase_Update(meYbase)
    End If
End If




srvYAUTE1I0_Import = "?"

paramYAUTE1I0_Import = paramYBase_DataF & Trim(constYAUTE1I0) & paramYBase_Data_ExtensionP

Open Trim(paramYAUTE1I0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYAUTE1I0) & Chr$(34)
MDB.Execute X
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYAUTE1I0_GRP) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYAUTE1I0
            meYbase.K1 = mId$(xIn, 6, 7) & mId$(xIn, 13, 7) & mId$(xIn, 66, 18) & Format$(Nb, "000000000")
                        ' .AUTE1IGRP.AUTE1ICLI.AUTE1IOR1.AUTE1IOR.AUTE1IOR3.AUTE1ITYP

            meYbase.Text = xIn
            dbYBase_Update meYbase
            If mId$(xIn, 6, 7) <> "       " Then
            
                xYbase.ID = constYAUTE1I0_GRP
                xYbase.K1 = mId$(xIn, 13, 7) & mId$(xIn, 6, 7)
                            ' .AUTE1ICLI.AUTE1Igrp
                xYbase.Method = "Seek="
                If tableYBase_Read(xYbase) <> 0 Then
                    xYbase.Method = constAddNew
                    
                    xYbase.Text = xIn
                    dbYBase_Update xYbase
                End If
        End If
        
    End If
Loop


Close
srvYAUTE1I0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = constYAUTE1I0
meYbase.Text = YBIATAB0_DATE_CPT_J & "_" & DSys & "_" & time_Hms & "_" & Format$(Nb, "000000000")
lX = meYbase.Text
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYAUTE1I0_Import" & xIn, vbCritical, Error
Close

srvYAUTE1I0_Import = Error
End Function


Public Function srvYAUTE1I0_Import_Read(lId As String, lYAUTE1I0 As typeYAUTE1I0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYAUTE1I0_Import_Read = "?"

meYbase.Method = "Seek>="
meYbase.ID = constYAUTE1I0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    If Trim(mId$(meYbase.K1, 1, 32)) = Trim(lId) Then
        MsgTxt = Space$(34) & meYbase.Text
        MsgTxtIndex = 0
        srvYAUTE1I0_GetBuffer lYAUTE1I0
        srvYAUTE1I0_Import_Read = Null
    End If
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYAUTE1I0_Import_Read" & xIn, vbCritical, Error
srvYAUTE1I0_Import_Read = Error
End Function


'-----------------------------------------------------
Function srvYAUTE1I0_Update(recYAUTE1I0 As typeYAUTE1I0)
'-----------------------------------------------------

srvYAUTE1I0_Update = "?"

MsgTxtLen = 0
Call srvYAUTE1I0_PutBuffer(recYAUTE1I0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYAUTE1I0_GetBuffer(recYAUTE1I0)) Then
        Call srvYAUTE1I0_Error(recYAUTE1I0)
        srvYAUTE1I0_Update = recYAUTE1I0.Err
        Exit Function
    Else
        srvYAUTE1I0_Update = Null
    End If
Else
    recYAUTE1I0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYAUTE1I0_Error(recYAUTE1I0 As typeYAUTE1I0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YAUTE1I0" & Chr$(10) & Chr$(13)

Select Case mId$(recYAUTE1I0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYAUTE1I0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YAUTE1I0s.bas  ( " & Trim(recYAUTE1I0.Obj) & " : " & Trim(recYAUTE1I0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYAUTE1I0_Monitor(recYAUTE1I0 As typeYAUTE1I0)
'-----------------------------------------------------

arrYAUTE1I0_Suite = False
Select Case mId$(Trim(recYAUTE1I0.Method), 1, 4)
    Case "Snap"
              srvYAUTE1I0_Monitor = srvYAUTE1I0_Snap(recYAUTE1I0)
    Case Else
            srvYAUTE1I0_Monitor = srvYAUTE1I0_Seek(recYAUTE1I0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYAUTE1I0_GetBuffer(recYAUTE1I0 As typeYAUTE1I0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYAUTE1I0_GetBuffer = Null
recYAUTE1I0.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYAUTE1I0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYAUTE1I0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYAUTE1I0.Err = Space$(10) Then
    recYAUTE1I0.AUTE1IETA = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYAUTE1I0.AUTE1IGRP = mId$(MsgTxt, K + 6, 7)
    recYAUTE1I0.AUTE1ICLI = mId$(MsgTxt, K + 13, 7)
    recYAUTE1I0.AUTE1ITYP = mId$(MsgTxt, K + 20, 1)
    recYAUTE1I0.AUTE1IAUT = mId$(MsgTxt, K + 21, 20)
    recYAUTE1I0.AUTE1IDEV = mId$(MsgTxt, K + 41, 3)
    recYAUTE1I0.AUTE1IAGE = CInt(Val(mId$(MsgTxt, K + 44, 5)))
    recYAUTE1I0.AUTE1ISER = mId$(MsgTxt, K + 49, 2)
    recYAUTE1I0.AUTE1ISRV = mId$(MsgTxt, K + 51, 2)
    recYAUTE1I0.AUTE1ICOP = mId$(MsgTxt, K + 53, 3)
    recYAUTE1I0.AUTE1INOP = CLng(Val(mId$(MsgTxt, K + 56, 10)))
    recYAUTE1I0.AUTE1IOR1 = CLng(Val(mId$(MsgTxt, K + 66, 6)))
    recYAUTE1I0.AUTE1IOR2 = CLng(Val(mId$(MsgTxt, K + 72, 6)))
    recYAUTE1I0.AUTE1IOR3 = CLng(Val(mId$(MsgTxt, K + 78, 6)))
    recYAUTE1I0.AUTE1IOR4 = CLng(Val(mId$(MsgTxt, K + 84, 6)))
    recYAUTE1I0.AUTE1IDBA = mId$(MsgTxt, K + 90, 3)
    recYAUTE1I0.AUTE1IMDB = CLng(Val(mId$(MsgTxt, K + 93, 16)))
    recYAUTE1I0.AUTE1IMCR = CLng(Val(mId$(MsgTxt, K + 109, 16)))
    recYAUTE1I0.AUTE1IBDB = CLng(Val(mId$(MsgTxt, K + 125, 16)))
    recYAUTE1I0.AUTE1IBCR = CLng(Val(mId$(MsgTxt, K + 141, 16)))
    recYAUTE1I0.AUTE1IRDB = CLng(Val(mId$(MsgTxt, K + 157, 16)))
    recYAUTE1I0.AUTE1IRCR = CLng(Val(mId$(MsgTxt, K + 173, 16)))
    recYAUTE1I0.AUTE1IMAU = CLng(Val(mId$(MsgTxt, K + 189, 16)))
    recYAUTE1I0.AUTE1IDAD = CLng(Val(mId$(MsgTxt, K + 205, 8)))
    recYAUTE1I0.AUTE1IDAF = CLng(Val(mId$(MsgTxt, K + 213, 8)))
    recYAUTE1I0.AUTE1IINT = mId$(MsgTxt, K + 221, 1)
    recYAUTE1I0.AUTE1IDMO = CLng(Val(mId$(MsgTxt, K + 222, 8)))
    recYAUTE1I0.AUTE1IRA1 = mId$(MsgTxt, K + 230, 32)
    recYAUTE1I0.AUTE1IRA2 = mId$(MsgTxt, K + 262, 32)
    recYAUTE1I0.AUTE1ISAC = mId$(MsgTxt, K + 294, 6)
    recYAUTE1I0.AUTE1IREG = mId$(MsgTxt, K + 300, 6)
    recYAUTE1I0.AUTE1ISRN = mId$(MsgTxt, K + 306, 9)
    recYAUTE1I0.AUTE1IRES = mId$(MsgTxt, K + 315, 3)
    recYAUTE1I0.AUTE1IECO = mId$(MsgTxt, K + 318, 3)
    recYAUTE1I0.AUTE1ICOT = mId$(MsgTxt, K + 321, 3)
    recYAUTE1I0.AUTE1IBDF = mId$(MsgTxt, K + 324, 4)
    recYAUTE1I0.AUTE1IDOU = mId$(MsgTxt, K + 328, 1)
    recYAUTE1I0.AUTE1IICH = mId$(MsgTxt, K + 329, 1)
    recYAUTE1I0.AUTE1ICET = mId$(MsgTxt, K + 330, 4)
    recYAUTE1I0.AUTE1ISIG = mId$(MsgTxt, K + 334, 12)
    recYAUTE1I0.AUTE1IRAG = mId$(MsgTxt, K + 346, 32)
    recYAUTE1I0.AUTE1IELM = mId$(MsgTxt, K + 378, 1)
    recYAUTE1I0.AUTE1INIV = CLng(Val(mId$(MsgTxt, K + 379, 4)))
    recYAUTE1I0.AUTE1IBLO = mId$(MsgTxt, K + 383, 1)
    recYAUTE1I0.AUTE1ICOM = mId$(MsgTxt, K + 384, 1)
    recYAUTE1I0.AUTE1ILAU = mId$(MsgTxt, K + 385, 30)
    recYAUTE1I0.AUTE1IECI = CLng(Val(mId$(MsgTxt, K + 415, 8)))
    recYAUTE1I0.AUTE1IDEP = mId$(MsgTxt, K + 423, 1)
    recYAUTE1I0.AUTE1IMTD = CLng(Val(mId$(MsgTxt, K + 424, 16)))
    recYAUTE1I0.AUTE1IDPD = CLng(Val(mId$(MsgTxt, K + 440, 8)))
    recYAUTE1I0.AUTE1IDTD = CLng(Val(mId$(MsgTxt, K + 448, 8)))
    recYAUTE1I0.AUTE1IC1A = mId$(MsgTxt, K + 456, 1)
    recYAUTE1I0.AUTE1IDEB = CLng(Val(mId$(MsgTxt, K + 457, 8)))
    recYAUTE1I0.AUTE1IFIN = CLng(Val(mId$(MsgTxt, K + 465, 8)))
    recYAUTE1I0.AUTE1ILIB = mId$(MsgTxt, K + 473, 32)
    recYAUTE1I0.AUTE1IRAT = mId$(MsgTxt, K + 505, 1)
    recYAUTE1I0.AUTE1IATR = mId$(MsgTxt, K + 506, 1)
    recYAUTE1I0.AUTE1IREL = mId$(MsgTxt, K + 507, 3)
    recYAUTE1I0.AUTE1IRUB = mId$(MsgTxt, K + 510, 10)
    recYAUTE1I0.AUTE1IAGC = CInt(Val(mId$(MsgTxt, K + 520, 5)))
    recYAUTE1I0.AUTE1ISEG = mId$(MsgTxt, K + 525, 3)
    recYAUTE1I0.AUTE1ISEP = mId$(MsgTxt, K + 528, 3)
    recYAUTE1I0.AUTE1IFUT = mId$(MsgTxt, K + 531, 150)
Else
    srvYAUTE1I0_GetBuffer = recYAUTE1I0.Err
    recYAUTE1I0.AUTE1ICLI = "?"
End If
MsgTxtIndex = MsgTxtIndex + recYAUTE1I0Len

End Function

'---------------------------------------------------------
Public Sub srvYAUTE1I0_PutBuffer(recYAUTE1I0 As typeYAUTE1I0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYAUTE1I0.Obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYAUTE1I0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
    Mid$(MsgTxt, K + 1, 5) = Format$(recYAUTE1I0.AUTE1IETA, "0000 ")
    Mid$(MsgTxt, K + 6, 7) = recYAUTE1I0.AUTE1IGRP
    Mid$(MsgTxt, K + 13, 7) = recYAUTE1I0.AUTE1ICLI
    Mid$(MsgTxt, K + 20, 1) = recYAUTE1I0.AUTE1ITYP
    Mid$(MsgTxt, K + 21, 20) = recYAUTE1I0.AUTE1IAUT
    Mid$(MsgTxt, K + 41, 3) = recYAUTE1I0.AUTE1IDEV
    Mid$(MsgTxt, K + 44, 5) = Format$(recYAUTE1I0.AUTE1IAGE, "0000 ")
    Mid$(MsgTxt, K + 49, 2) = recYAUTE1I0.AUTE1ISER
    Mid$(MsgTxt, K + 51, 2) = recYAUTE1I0.AUTE1ISRV
    Mid$(MsgTxt, K + 53, 3) = recYAUTE1I0.AUTE1ICOP
    Mid$(MsgTxt, K + 56, 10) = Format$(recYAUTE1I0.AUTE1INOP, "000000000 ")
    Mid$(MsgTxt, K + 66, 6) = Format$(recYAUTE1I0.AUTE1IOR1, "00000 ")
    Mid$(MsgTxt, K + 72, 6) = Format$(recYAUTE1I0.AUTE1IOR2, "00000 ")
    Mid$(MsgTxt, K + 78, 6) = Format$(recYAUTE1I0.AUTE1IOR3, "00000 ")
    Mid$(MsgTxt, K + 84, 6) = Format$(recYAUTE1I0.AUTE1IOR4, "00000 ")
    Mid$(MsgTxt, K + 90, 3) = recYAUTE1I0.AUTE1IDBA
    Mid$(MsgTxt, K + 93, 16) = Format$(recYAUTE1I0.AUTE1IMDB, "000000000000000 ")
    Mid$(MsgTxt, K + 109, 16) = Format$(recYAUTE1I0.AUTE1IMCR, "000000000000000 ")
    Mid$(MsgTxt, K + 125, 16) = Format$(recYAUTE1I0.AUTE1IBDB, "000000000000000 ")
    Mid$(MsgTxt, K + 141, 16) = Format$(recYAUTE1I0.AUTE1IBCR, "000000000000000 ")
    Mid$(MsgTxt, K + 157, 16) = Format$(recYAUTE1I0.AUTE1IRDB, "000000000000000 ")
    Mid$(MsgTxt, K + 173, 16) = Format$(recYAUTE1I0.AUTE1IRCR, "000000000000000 ")
    Mid$(MsgTxt, K + 189, 16) = Format$(recYAUTE1I0.AUTE1IMAU, "000000000000000 ")
    Mid$(MsgTxt, K + 205, 8) = Format$(recYAUTE1I0.AUTE1IDAD, "0000000 ")
    Mid$(MsgTxt, K + 213, 8) = Format$(recYAUTE1I0.AUTE1IDAF, "0000000 ")
    Mid$(MsgTxt, K + 221, 1) = recYAUTE1I0.AUTE1IINT
    Mid$(MsgTxt, K + 222, 8) = Format$(recYAUTE1I0.AUTE1IDMO, "0000000 ")
    Mid$(MsgTxt, K + 230, 32) = recYAUTE1I0.AUTE1IRA1
    Mid$(MsgTxt, K + 262, 32) = recYAUTE1I0.AUTE1IRA2
    Mid$(MsgTxt, K + 294, 6) = recYAUTE1I0.AUTE1ISAC
    Mid$(MsgTxt, K + 300, 6) = recYAUTE1I0.AUTE1IREG
    Mid$(MsgTxt, K + 306, 9) = recYAUTE1I0.AUTE1ISRN
    Mid$(MsgTxt, K + 315, 3) = recYAUTE1I0.AUTE1IRES
    Mid$(MsgTxt, K + 318, 3) = recYAUTE1I0.AUTE1IECO
    Mid$(MsgTxt, K + 321, 3) = recYAUTE1I0.AUTE1ICOT
    Mid$(MsgTxt, K + 324, 4) = recYAUTE1I0.AUTE1IBDF
    Mid$(MsgTxt, K + 328, 1) = recYAUTE1I0.AUTE1IDOU
    Mid$(MsgTxt, K + 329, 1) = recYAUTE1I0.AUTE1IICH
    Mid$(MsgTxt, K + 330, 4) = recYAUTE1I0.AUTE1ICET
    Mid$(MsgTxt, K + 334, 12) = recYAUTE1I0.AUTE1ISIG
    Mid$(MsgTxt, K + 346, 32) = recYAUTE1I0.AUTE1IRAG
    Mid$(MsgTxt, K + 378, 1) = recYAUTE1I0.AUTE1IELM
    Mid$(MsgTxt, K + 379, 4) = Format$(recYAUTE1I0.AUTE1INIV, "000 ")
    Mid$(MsgTxt, K + 383, 1) = recYAUTE1I0.AUTE1IBLO
    Mid$(MsgTxt, K + 384, 1) = recYAUTE1I0.AUTE1ICOM
    Mid$(MsgTxt, K + 385, 30) = recYAUTE1I0.AUTE1ILAU
    Mid$(MsgTxt, K + 415, 8) = Format$(recYAUTE1I0.AUTE1IECI, "0000000 ")
    Mid$(MsgTxt, K + 423, 1) = recYAUTE1I0.AUTE1IDEP
    Mid$(MsgTxt, K + 424, 16) = Format$(recYAUTE1I0.AUTE1IMTD, "000000000000000 ")
    Mid$(MsgTxt, K + 440, 8) = Format$(recYAUTE1I0.AUTE1IDPD, "0000000 ")
    Mid$(MsgTxt, K + 448, 8) = Format$(recYAUTE1I0.AUTE1IDTD, "0000000 ")
    Mid$(MsgTxt, K + 456, 1) = recYAUTE1I0.AUTE1IC1A
    Mid$(MsgTxt, K + 457, 8) = Format$(recYAUTE1I0.AUTE1IDEB, "0000000 ")
    Mid$(MsgTxt, K + 465, 8) = Format$(recYAUTE1I0.AUTE1IFIN, "0000000 ")
    Mid$(MsgTxt, K + 473, 32) = recYAUTE1I0.AUTE1ILIB
    Mid$(MsgTxt, K + 505, 1) = recYAUTE1I0.AUTE1IRAT
    Mid$(MsgTxt, K + 506, 1) = recYAUTE1I0.AUTE1IATR
    Mid$(MsgTxt, K + 507, 3) = recYAUTE1I0.AUTE1IREL
    Mid$(MsgTxt, K + 510, 10) = recYAUTE1I0.AUTE1IRUB
    Mid$(MsgTxt, K + 520, 5) = Format$(recYAUTE1I0.AUTE1IAGC, "0000 ")
    Mid$(MsgTxt, K + 525, 3) = recYAUTE1I0.AUTE1ISEG
    Mid$(MsgTxt, K + 528, 3) = recYAUTE1I0.AUTE1ISEP
    Mid$(MsgTxt, K + 531, 150) = recYAUTE1I0.AUTE1IFUT

MsgTxtLen = MsgTxtLen + recYAUTE1I0Len
End Sub


Public Sub srvYAUTE1I0_ElpDisplay(recYAUTE1I0 As typeYAUTE1I0)
frmElpDisplay.fgData.Rows = 64
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IGRP    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "GROUPE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IGRP
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1ICLI    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CLIENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1ICLI
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1ITYP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE 1,2,3"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1ITYP
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IAUT   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE AUTORISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IAUT
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IDEV    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IDEV
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IAGE
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1ISER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1ISER
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1ISRV    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1ISRV
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1ICOP    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1ICOP
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1INOP    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1INOP
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IOR1    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ORDRE 1"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IOR1
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IOR2    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ORDRE 2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IOR2
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IOR3    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ORDRE 3"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IOR3
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IOR4    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ORDRE 4"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IOR4
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IDBA    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE DE BASE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IDBA
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IMDB   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT DEBIT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IMDB
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IMCR   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT CREDIT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IMCR
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IBDB   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT DEVBAS DB"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IBDB
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IBCR   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT DEVBAS CR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IBCR
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IRDB   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT REPOR. DB"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IRDB
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IRCR   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT REPOR. CR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IRCR
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IMAU   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT AUTO"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IMAU
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IDAD    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DEBUT AUTO."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IDAD
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IDAF    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE FIN AUTO."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IDAF
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IINT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTITULITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IINT
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IDMO    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DERN.MOUV."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IDMO
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IRA1   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RAISON SOCIALE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IRA1
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IRA2   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RAISON SOCIALE 2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IRA2
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1ISAC    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SECTEUR ACTIVITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1ISAC
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IREG    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SECTEUR ACT. REG."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IREG
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1ISRN    9A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO SIREN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1ISRN
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IRES    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RESPON/EXPLOIT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IRES
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IECO    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "QUALITE/AG.ECONO"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IECO
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1ICOT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COTATION INTERNE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1ICOT
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IBDF    4A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE BDF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IBDF
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IDOU    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DOUTEUX  O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IDOU
frmElpDisplay.fgData.Row = 38
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IICH    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTERDIT CHQ  O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IICH
frmElpDisplay.fgData.Row = 39
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1ICET    4A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETAT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1ICET
frmElpDisplay.fgData.Row = 40
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1ISIG   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SIGLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1ISIG
frmElpDisplay.fgData.Row = 41
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IRAG   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RAISON SOC GROUPE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IRAG
frmElpDisplay.fgData.Row = 42
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IELM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ELEM. O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IELM
frmElpDisplay.fgData.Row = 43
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1INIV    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NIVEAU"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1INIV
frmElpDisplay.fgData.Row = 44
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IBLO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE BLOCAGE1,2,3"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IBLO
frmElpDisplay.fgData.Row = 45
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1ICOM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPENSATION O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1ICOM
frmElpDisplay.fgData.Row = 46
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1ILAU   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBEL AUTO OU GAR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1ILAU
frmElpDisplay.fgData.Row = 47
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IECI    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ECHEANCE INTERNE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IECI
frmElpDisplay.fgData.Row = 48
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IDEP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE DEPASSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IDEP
frmElpDisplay.fgData.Row = 49
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IMTD   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT DEPASSEM"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IMTD
frmElpDisplay.fgData.Row = 50
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IDPD    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE 1ER DEPAS."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IDPD
frmElpDisplay.fgData.Row = 51
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IDTD    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEPASSEM  DEPUIS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IDTD
frmElpDisplay.fgData.Row = 52
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IC1A    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "C1AUT POUR DEPASS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IC1A
frmElpDisplay.fgData.Row = 53
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IDEB    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DEBUT OPERA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IDEB
frmElpDisplay.fgData.Row = 54
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IFIN    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE FIN OPERA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IFIN
frmElpDisplay.fgData.Row = 55
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1ILIB   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBELLE OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1ILIB
frmElpDisplay.fgData.Row = 56
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IRAT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RATTAC GROUPE O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IRAT
frmElpDisplay.fgData.Row = 57
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IATR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AUTO GROUPE(1à9)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IATR
frmElpDisplay.fgData.Row = 58
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IREL    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RELATION CLI-GRP"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IREL
frmElpDisplay.fgData.Row = 59
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IRUB   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RUBRIQUE COMPT."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IRUB
frmElpDisplay.fgData.Row = 60
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IAGC    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE CLIENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IAGC
frmElpDisplay.fgData.Row = 61
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1ISEG    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SEGEMENT DE RESULTAT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1ISEG
frmElpDisplay.fgData.Row = 62
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1ISEP    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SEGEMENT POTENTIEL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1ISEP
frmElpDisplay.fgData.Row = 63
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTE1IFUT  150A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ZONE FUTURE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTE1I0.AUTE1IFUT
frmElpDisplay.Show vbModal
End Sub
Public Sub srvYAUTE1I0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YAUTE1I0.txt" For Input As #1
Open "C:\Temp\YAUTE1I0.csv" For Output As #2
Print #2, "AUTE1IETA;AUTE1IGRP;AUTE1ICLI;AUTE1ITYP;AUTE1IAUT;AUTE1IDEV;AUTE1IAGE;AUTE1ISER;AUTE1ISRV;AUTE1ICOP;AUTE1INOP;AUTE1IOR1;AUTE1IOR2;AUTE1IOR3;AUTE1IOR4;AUTE1IDBA;AUTE1IMDB;AUTE1IMCR;AUTE1IBDB;AUTE1IBCR;AUTE1IRDB;AUTE1IRCR;AUTE1IMAU;AUTE1IDAD;AUTE1IDAF;AUTE1IINT;AUTE1IDMO;AUTE1IRA1;AUTE1IRA2;AUTE1ISAC;AUTE1IREG;AUTE1ISRN;AUTE1IRES;AUTE1IECO;AUTE1ICOT;AUTE1IBDF;AUTE1IDOU;AUTE1IICH;AUTE1ICET;AUTE1ISIG;AUTE1IRAG;AUTE1IELM;AUTE1INIV;AUTE1IBLO;AUTE1ICOM;AUTE1ILAU;AUTE1IECI;AUTE1IDEP;AUTE1IMTD;AUTE1IDPD;AUTE1IDTD;AUTE1IC1A;AUTE1IDEB;AUTE1IFIN;AUTE1ILIB;AUTE1IRAT;AUTE1IATR;AUTE1IREL;AUTE1IRUB;AUTE1IAGC;AUTE1ISEG;AUTE1ISEP;AUTE1IFUT;"
Print #2, "ETABLISSEMENT;GROUPE;CLIENT;TYPE 1,2,3;CODE AUTORISATION;DEVISE;AGENCE;SERVICE;SOUS SERVICE;CODE OPERATION;NUMERO OPERATION;ORDRE 1;ORDRE 2;ORDRE 3;ORDRE 4;DEVISE DE BASE;MONTANT DEBIT;MONTANT CREDIT;MONTANT DEVBAS DB;MONTANT DEVBAS CR;MONTANT REPOR. DB;MONTANT REPOR. CR;MONTANT AUTO;DATE DEBUT AUTO.;DATE FIN AUTO.;INTITULITE;DATE DERN.MOUV.;RAISON SOCIALE;RAISON SOCIALE 2;SECTEUR ACTIVITE;SECTEUR ACT. REG.;NUMERO SIREN;RESPON/EXPLOIT;QUALITE/AG.ECONO;COTATION INTERNE;CODE BDF;DOUTEUX  O/N;INTERDIT CHQ  O/N;CODE ETAT;SIGLE;RAISON SOC GROUPE;CODE ELEM. O/N;NIVEAU;CODE BLOCAGE1,2,3;COMPENSATION O/N;LIBEL AUTO OU GAR;ECHEANCE INTERNE;CODE DEPASSEMENT;MONTANT DEPASSEM;DATE 1ER DEPAS.;DEPASSEM  DEPUIS;C1AUT POUR DEPASS;DATE DEBUT OPERA;DATE FIN OPERA;LIBELLE OPERATION;RATTAC GROUPE O/N;AUTO GROUPE(1à9);RELATION CLI-GRP;RUBRIQUE COMPT.;AGENCE CLIENT;SEGEMENT DE RESULTAT;SEGEMENT POTENTIEL;ZONE FUTURE;"
Print #2, ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 7) & ";" & mId$(xIn, 13, 7) & ";" _
      & mId$(xIn, 20, 1) & ";" & mId$(xIn, 21, 20) & ";" & mId$(xIn, 41, 3) & ";" _
      & mId$(xIn, 44, 5) & ";" & mId$(xIn, 49, 2) & ";" & mId$(xIn, 51, 2) & ";" _
      & mId$(xIn, 53, 3) & ";" & mId$(xIn, 56, 10) & ";" & mId$(xIn, 66, 6) & ";" _
      & mId$(xIn, 72, 6) & ";" & mId$(xIn, 78, 6) & ";" & mId$(xIn, 84, 6) & ";" _
      & mId$(xIn, 90, 3) & ";" & mId$(xIn, 93, 16) & ";" & mId$(xIn, 109, 16) & ";" _
      & mId$(xIn, 125, 16) & ";" & mId$(xIn, 141, 16) & ";" & mId$(xIn, 157, 16) & ";" _
      & mId$(xIn, 173, 16) & ";" & mId$(xIn, 189, 16) & ";" & mId$(xIn, 205, 8) & ";" _
      & mId$(xIn, 213, 8) & ";" & mId$(xIn, 221, 1) & ";" & mId$(xIn, 222, 8) & ";" _
      & mId$(xIn, 230, 32) & ";" & mId$(xIn, 262, 32) & ";" & mId$(xIn, 294, 6) & ";" _
      & mId$(xIn, 300, 6) & ";" & mId$(xIn, 306, 9) & ";" & mId$(xIn, 315, 3) & ";" _
      & mId$(xIn, 318, 3) & ";" & mId$(xIn, 321, 3) & ";" & mId$(xIn, 324, 4) & ";" _
      & mId$(xIn, 328, 1) & ";" & mId$(xIn, 329, 1) & ";" & mId$(xIn, 330, 4) & ";" _
      & mId$(xIn, 334, 12) & ";" & mId$(xIn, 346, 32) & ";" & mId$(xIn, 378, 1) & ";" _
      & mId$(xIn, 379, 4) & ";" & mId$(xIn, 383, 1) & ";" & mId$(xIn, 384, 1) & ";" _
      & mId$(xIn, 385, 30) & ";" & mId$(xIn, 415, 8) & ";" & mId$(xIn, 423, 1) & ";" _
      & mId$(xIn, 424, 16) & ";" & mId$(xIn, 440, 8) & ";" & mId$(xIn, 448, 8) & ";" _
      & mId$(xIn, 456, 1) & ";" & mId$(xIn, 457, 8) & ";" & mId$(xIn, 465, 8) & ";" _
      & mId$(xIn, 473, 32) & ";" & mId$(xIn, 505, 1) & ";" & mId$(xIn, 506, 1) & ";" _
      & mId$(xIn, 507, 3) & ";" & mId$(xIn, 510, 10) & ";" & mId$(xIn, 520, 5) & ";" _
      & mId$(xIn, 525, 3) & ";" & mId$(xIn, 528, 3) & ";" & mId$(xIn, 531, 150) & ";"
Loop
Close
End Sub
'---------------------------------------------------------
Private Function srvYAUTE1I0_Seek(recYAUTE1I0 As typeYAUTE1I0)
'---------------------------------------------------------

srvYAUTE1I0_Seek = "?"
MsgTxtLen = 0
Call srvYAUTE1I0_PutBuffer(recYAUTE1I0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYAUTE1I0_GetBuffer(recYAUTE1I0)) Then
        srvYAUTE1I0_Seek = Null
    Else
        Call srvYAUTE1I0_Error(recYAUTE1I0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYAUTE1I0_Snap(recYAUTE1I0 As typeYAUTE1I0)
'---------------------------------------------------------
srvYAUTE1I0_Snap = "?"
MsgTxtLen = 0
Call srvYAUTE1I0_PutBuffer(recYAUTE1I0)
Call srvYAUTE1I0_PutBuffer(arrYAUTE1I0(0))
If IsNull(SndRcv()) Then
    srvYAUTE1I0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYAUTE1I0_GetBuffer(recYAUTE1I0)) Then
            Call arrYAUTE1I0_AddItem(recYAUTE1I0)
            arrYAUTE1I0_Suite = True
        Else
            arrYAUTE1I0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYAUTE1I0_AddItem(recYAUTE1I0 As typeYAUTE1I0)
'---------------------------------------------------------
          
arrYAUTE1I0_NB = arrYAUTE1I0_NB + 1
    
If arrYAUTE1I0_NB > arrYAUTE1I0_NBMax Then
    arrYAUTE1I0_NBMax = arrYAUTE1I0_NBMax + recYAUTE1I0_Block
    ReDim Preserve arrYAUTE1I0(arrYAUTE1I0_NBMax)
End If
            
arrYAUTE1I0(arrYAUTE1I0_NB) = recYAUTE1I0
End Sub



'---------------------------------------------------------
Public Sub recYAUTE1I0_Init(recYAUTE1I0 As typeYAUTE1I0)
'---------------------------------------------------------
recYAUTE1I0.Obj = "ZAUTE1I0_S"
recYAUTE1I0.Method = ""
recYAUTE1I0.Err = ""

End Sub










