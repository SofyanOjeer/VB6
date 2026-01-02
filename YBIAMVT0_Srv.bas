Attribute VB_Name = "srvYBIAMVT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYBIAMVT0Len = 463 ' 34 +429
Public Const recYBIAMVT0_Block = 100
Public Const memoYBIAMVT0Len = 429
Public Const constYBIAMVT0 = "YBIAMVT0"
Public paramYBIAMVT0_Import As String

Dim meYbase As typeYBase

Type typeYBIAMVT0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    MOUVEMETA       As Integer                        ' ETABLISSEMENT
    MOUVEMPLA       As Long                           ' NUMERO PLAN
    MOUVEMCOM       As String * 20                    ' NUMERO COMPTE
    MOUVEMMON       As Currency                       ' MONTANT
    MOUVEMDOP       As Long                           ' DATE D'OPERATION
    MOUVEMDVA       As Long                           ' DATE DE VALEUR
    MOUVEMDCO       As Long                           ' DATE COMPTABLE
    MOUVEMDTR       As Long                           ' DATE DE TRAITEMENT
    MOUVEMPIE       As Long                           ' NUMERO DE PIECE
    MOUVEMECR       As Long                           ' NUMERO D'ECRITURE
    MOUVEMOPE       As String * 3                     ' CODE OPERATION
    MOUVEMNUM       As Long                           ' NUMERO OPERATION
    MOUVEMSCH       As Integer                        ' CODE SCHEMA
    MOUVEMUTI       As Integer                        ' UTILISATEUR
    MOUVEMAGE       As Integer                        ' AGENCE OPERATRICE
    MOUVEMSER       As String * 2                     ' SERVICE OPERATEUR
    MOUVEMSSE       As String * 2                     ' S/SERVICE OPERATEUR
    MOUVEMEXO       As String * 1                     ' CODE EXONERATION
    MOUVEMANA       As String * 6                     ' CODE ANALYTIQUE
    MOUVEMBDF       As String * 3                     ' CODE BANQUE DE FR.
    MOUVEMANU       As String * 1                     ' CODE ANNULATION
    MOUVEMRET       As String * 1                     ' MOUVEMENT RETRO
    MOUVEMEVE       As String * 3                     ' EVENEMENT
    MOUVEMSAN       As String * 6                     ' STRUCT ANALY-CODE
    MOUVEMSAD       As String * 80                    ' STRUCT ANALY-DONNEES
    
    LIBELLIB1       As String * 30                    ' Libellé 1
    LIBELLIB2       As String * 30                    ' Libellé 2
    LIBELLIB3       As String * 30                    ' Libellé 3
    LIBELLIB4       As String * 30                    ' Libellé 4
    
    COMPTEOBL       As String * 10                    ' COMPTE OBLIGATOIRE
    COMPTEINT       As String * 32                    ' INTITULE
    COMPTEDEV       As String * 3                     ' TABLES BASE 013
    COMPTELOR       As String * 1                     ' Lori/Nostri/AUTRE
    COMPTECLA       As Long                           ' CLASSE SECURITE
    
    BIAMVTSD0       As Currency                       ' solde
    BIAMVTID        As Long                           ' référence

End Type
    
    
Public arrYBIAMVT0() As typeYBIAMVT0
Public arrYBIAMVT0_NB As Integer
Public arrYBIAMVT0_NBMax As Integer
Public arrYBIAMVT0_Index As Integer
Public arrYBIAMVT0_Suite As Boolean

Dim meMVTP0 As typeMvtP0

Public Sub srvYBIAMVT0_Export_CSV(lIdFile_Source As Integer, lIdFile_Destination As Integer, loptSelect_CSV_Header As Boolean, lnb As Long)
Dim xIn As String
Dim V

If loptSelect_CSV_Header Then
    Print #lIdFile_Destination, "?"
    Print #lIdFile_Destination, "?"
    Print #lIdFile_Destination, "?"
End If
Do Until EOF(lIdFile_Source)
      Line Input #lIdFile_Source, xIn
      lnb = lnb + 1
          Print #lIdFile_Destination, mId$(xIn, 1, 5) & ";" & mId$(xIn, 6, 4) & ";" & mId$(xIn, 10, 20) & ";" & _
                    cur_19V(CCur(mId$(xIn, 30, 18)) / 1000) & ";" _
                    ; mId$(xIn, 48, 8) & ";" & mId$(xIn, 56, 8) & ";" & _
                    mId$(xIn, 64, 8) & ";" & mId$(xIn, 72, 8) & ";" & _
                    mId$(xIn, 80, 10) & ";" & mId$(xIn, 90, 8) & ";" & _
                    mId$(xIn, 98, 3) & ";" & mId$(xIn, 101, 10) & ";" & _
                    mId$(xIn, 111, 5) & ";" & mId$(xIn, 116, 5) & ";" & _
                    mId$(xIn, 121, 5) & ";" & mId$(xIn, 126, 2) & ";" & _
                    mId$(xIn, 128, 2) & ";" & mId$(xIn, 130, 1) & ";" & _
                    mId$(xIn, 131, 6) & ";" & mId$(xIn, 137, 3) & ";" & _
                    mId$(xIn, 140, 1) & ";" & mId$(xIn, 141, 1) & ";" & _
                    mId$(xIn, 142, 3) & ";" & mId$(xIn, 145, 6) & ";" & _
                    mId$(xIn, 151, 80) & ";" & mId$(xIn, 231, 30) & ";" & _
                    mId$(xIn, 261, 30) & ";" & mId$(xIn, 291, 30) & ";" & _
                    mId$(xIn, 321, 30) & ";" & mId$(xIn, 351, 10) & ";" & mId$(xIn, 361, 32) & ";" & _
                    mId$(xIn, 393, 3) & ";" & mId$(xIn, 396, 1) & ";" & _
                    mId$(xIn, 397, 3) & ";" & _
                    cur_19V(CCur(mId$(xIn, 400, 19)) / 100) & ";" & _
                    mId$(xIn, 419, 11)
Loop

Exit Sub

End Sub


Public Function oldYBIAMVT0_Import(lnb As Long)
Dim xIn As String, X As String

On Error GoTo Error_Handle

oldYBIAMVT0_Import = "?"

paramYBIAMVT0_Import = paramYBase_DataF & Trim(constYBIAMVT0) & paramYBase_Data_ExtensionP

Open Trim(paramYBIAMVT0_Import) For Input As #1

lnb = 0

recMvtP0_Init meMVTP0
meMVTP0.Method = constAddNew

mdbMvtP0.tableMvtP0_Open

Do Until EOF(1)
    lnb = lnb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meMVTP0.ID = constYBIAMVT0 & mId$(xIn, 419, 11)   'Format(seq, "0000000000")
            meMVTP0.Text = xIn
            dbMvtP0_Update meMVTP0
            
    End If
Loop


Close
oldYBIAMVT0_Import = Null
Exit Function

Error_Handle:
 MsgBox "erreur : oldYBIAMVT0_Import" & xIn, vbCritical, Error
Close

oldYBIAMVT0_Import = Error
End Function


Public Function srvYBIAMVT0_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYBIAMVT0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    srvYBIAMVT0_Import = Null
    lX = meYbase.Text
    Exit Function
End If


srvYBIAMVT0_Import = "?"


Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYBIAMVT0) & Chr$(34)
MDB.Execute X

X = mId$(YBIATAB0_DATE_CPT_MP1, 1, 6) & "_" & constYBIAMVT0    ' mois précédent
srvYBIAMVT0_Import_File X, Nb

srvYBIAMVT0_Import_File constYBIAMVT0, Nb                       ' mois en cours

srvYBIAMVT0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYBIAMVT0
meYbase.Text = DSys & "_" & time_Hms & "_" & Format$(Nb, "000000000")
lX = meYbase.Text
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYBIAMVT0_Import" & xIn, vbCritical, Error
Close

srvYBIAMVT0_Import = Error
End Function



Public Function srvYBIAMVT0_Import_Read(lId As String, lYBIAMVT0 As typeYBIAMVT0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYBIAMVT0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYBIAMVT0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 And Trim(meYbase.ID) = constYBIAMVT0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYBIAMVT0_GetBuffer lYBIAMVT0
    srvYBIAMVT0_Import_Read = Null
Else
    recYBIAMVT0_Init lYBIAMVT0
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYBIAMVT0_Import_Read" & xIn, vbCritical, Error
srvYBIAMVT0_Import_Read = Error
End Function


Public Function oldYBIAMVT0_Import_Read(lId As String, lYBIAMVT0 As typeYBIAMVT0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

oldYBIAMVT0_Import_Read = "?"

meMVTP0.Method = "Seek="
meMVTP0.ID = lId
If tableMvtP0_Read(meMVTP0) = 0 Then
    MsgTxt = Space$(34) & meMVTP0.Text
    MsgTxtIndex = 0
    srvYBIAMVT0_GetBuffer lYBIAMVT0
    oldYBIAMVT0_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : oldYBIAMVT0_Import_Read" & xIn, vbCritical, Error
Close
oldYBIAMVT0_Import_Read = Error
End Function


'-----------------------------------------------------
Function srvYBIAMVT0_Update(recYBIAMVT0 As typeYBIAMVT0)
'-----------------------------------------------------

srvYBIAMVT0_Update = "?"

MsgTxtLen = 0
Call srvYBIAMVT0_PutBuffer(recYBIAMVT0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYBIAMVT0_GetBuffer(recYBIAMVT0)) Then
        Call srvYBIAMVT0_Error(recYBIAMVT0)
        srvYBIAMVT0_Update = recYBIAMVT0.Err
        Exit Function
    Else
        srvYBIAMVT0_Update = Null
    End If
Else
    recYBIAMVT0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYBIAMVT0_Error(recYBIAMVT0 As typeYBIAMVT0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YBIAMVT0" & Chr$(10) & Chr$(13)

Select Case mId$(recYBIAMVT0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYBIAMVT0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YBIAMVT0s.bas  ( " & Trim(recYBIAMVT0.Obj) & " : " & Trim(recYBIAMVT0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYBIAMVT0_Monitor(recYBIAMVT0 As typeYBIAMVT0)
'-----------------------------------------------------

arrYBIAMVT0_Suite = False
Select Case mId$(Trim(recYBIAMVT0.Method), 1, 4)
    Case "Snap"
              srvYBIAMVT0_Monitor = srvYBIAMVT0_Snap(recYBIAMVT0)
    Case Else
            srvYBIAMVT0_Monitor = srvYBIAMVT0_Seek(recYBIAMVT0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYBIAMVT0_GetBuffer(recYBIAMVT0 As typeYBIAMVT0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYBIAMVT0_GetBuffer = Null
recYBIAMVT0.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYBIAMVT0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYBIAMVT0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYBIAMVT0.Err = Space$(10) Then
    recYBIAMVT0.MOUVEMETA = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYBIAMVT0.MOUVEMPLA = CLng(Val(mId$(MsgTxt, K + 6, 4)))
    recYBIAMVT0.MOUVEMCOM = mId$(MsgTxt, K + 10, 20)
    recYBIAMVT0.MOUVEMMON = CCur(mId$(MsgTxt, K + 30, 18)) / 1000
    recYBIAMVT0.MOUVEMDOP = CLng(Val(mId$(MsgTxt, K + 48, 8)))
    recYBIAMVT0.MOUVEMDVA = CLng(Val(mId$(MsgTxt, K + 56, 8)))
    recYBIAMVT0.MOUVEMDCO = CLng(Val(mId$(MsgTxt, K + 64, 8)))
    recYBIAMVT0.MOUVEMDTR = CLng(Val(mId$(MsgTxt, K + 72, 8)))
    recYBIAMVT0.MOUVEMPIE = CLng(Val(mId$(MsgTxt, K + 80, 10)))
    recYBIAMVT0.MOUVEMECR = CLng(Val(mId$(MsgTxt, K + 90, 8)))
    recYBIAMVT0.MOUVEMOPE = mId$(MsgTxt, K + 98, 3)
    recYBIAMVT0.MOUVEMNUM = CLng(Val(mId$(MsgTxt, K + 101, 10)))
    recYBIAMVT0.MOUVEMSCH = CInt(Val(mId$(MsgTxt, K + 111, 5)))
    recYBIAMVT0.MOUVEMUTI = CInt(Val(mId$(MsgTxt, K + 116, 5)))
    recYBIAMVT0.MOUVEMAGE = CInt(Val(mId$(MsgTxt, K + 121, 5)))
    recYBIAMVT0.MOUVEMSER = mId$(MsgTxt, K + 126, 2)
    recYBIAMVT0.MOUVEMSSE = mId$(MsgTxt, K + 128, 2)
    recYBIAMVT0.MOUVEMEXO = mId$(MsgTxt, K + 130, 1)
    recYBIAMVT0.MOUVEMANA = mId$(MsgTxt, K + 131, 6)
    recYBIAMVT0.MOUVEMBDF = mId$(MsgTxt, K + 137, 3)
    recYBIAMVT0.MOUVEMANU = mId$(MsgTxt, K + 140, 1)
    recYBIAMVT0.MOUVEMRET = mId$(MsgTxt, K + 141, 1)
    recYBIAMVT0.MOUVEMEVE = mId$(MsgTxt, K + 142, 3)
    recYBIAMVT0.MOUVEMSAN = mId$(MsgTxt, K + 145, 6)
    recYBIAMVT0.MOUVEMSAD = mId$(MsgTxt, K + 151, 80)
    
    recYBIAMVT0.LIBELLIB1 = mId$(MsgTxt, K + 231, 30)
    recYBIAMVT0.LIBELLIB2 = mId$(MsgTxt, K + 261, 30)
    recYBIAMVT0.LIBELLIB3 = mId$(MsgTxt, K + 291, 30)
    recYBIAMVT0.LIBELLIB4 = mId$(MsgTxt, K + 321, 30)
        
    recYBIAMVT0.COMPTEOBL = mId$(MsgTxt, K + 351, 10)
    recYBIAMVT0.COMPTEINT = mId$(MsgTxt, K + 361, 32)
    recYBIAMVT0.COMPTEDEV = mId$(MsgTxt, K + 393, 3)
    recYBIAMVT0.COMPTELOR = mId$(MsgTxt, K + 396, 1)
    recYBIAMVT0.COMPTECLA = CLng(Val(mId$(MsgTxt, K + 397, 3)))
    If IsNumeric(mId$(MsgTxt, K + 400, 19)) Then
        recYBIAMVT0.BIAMVTSD0 = CCur(mId$(MsgTxt, K + 400, 19)) / 100
        recYBIAMVT0.BIAMVTID = CLng(Val(mId$(MsgTxt, K + 419, 11)))
    Else
        recYBIAMVT0.BIAMVTSD0 = 0
        recYBIAMVT0.BIAMVTID = 0
        MsgBox "YBIAMVT0 ???? " & MsgTxt, vbCritical
    End If
Else
    srvYBIAMVT0_GetBuffer = recYBIAMVT0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYBIAMVT0Len

End Function

'---------------------------------------------------------
Public Sub srvYBIAMVT0_PutBuffer(recYBIAMVT0 As typeYBIAMVT0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYBIAMVT0.Obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYBIAMVT0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYBIAMVT0.MOUVEMETA, "0000 ")
    Mid$(MsgTxt, K + 6, 4) = Format$(recYBIAMVT0.MOUVEMPLA, "000 ")
    Mid$(MsgTxt, K + 10, 20) = recYBIAMVT0.MOUVEMCOM
    Mid$(MsgTxt, K + 30, 18) = Format$(recYBIAMVT0.MOUVEMMON * 1000, "00000000000000000 ")
    Mid$(MsgTxt, K + 48, 8) = Format$(recYBIAMVT0.MOUVEMDOP, "0000000 ")
    Mid$(MsgTxt, K + 56, 8) = Format$(recYBIAMVT0.MOUVEMDVA, "0000000 ")
    Mid$(MsgTxt, K + 64, 8) = Format$(recYBIAMVT0.MOUVEMDCO, "0000000 ")
    Mid$(MsgTxt, K + 72, 8) = Format$(recYBIAMVT0.MOUVEMDTR, "0000000 ")
    Mid$(MsgTxt, K + 80, 10) = Format$(recYBIAMVT0.MOUVEMPIE, "000000000 ")
    Mid$(MsgTxt, K + 90, 8) = Format$(recYBIAMVT0.MOUVEMECR, "0000000 ")
    Mid$(MsgTxt, K + 98, 3) = recYBIAMVT0.MOUVEMOPE
    Mid$(MsgTxt, K + 101, 10) = Format$(recYBIAMVT0.MOUVEMNUM, "000000000 ")
    Mid$(MsgTxt, K + 111, 5) = Format$(recYBIAMVT0.MOUVEMSCH, "0000 ")
    Mid$(MsgTxt, K + 116, 5) = Format$(recYBIAMVT0.MOUVEMUTI, "0000 ")
    Mid$(MsgTxt, K + 121, 5) = Format$(recYBIAMVT0.MOUVEMAGE, "0000 ")
    Mid$(MsgTxt, K + 126, 2) = recYBIAMVT0.MOUVEMSER
    Mid$(MsgTxt, K + 128, 2) = recYBIAMVT0.MOUVEMSSE
    Mid$(MsgTxt, K + 130, 1) = recYBIAMVT0.MOUVEMEXO
    Mid$(MsgTxt, K + 131, 6) = recYBIAMVT0.MOUVEMANA
    Mid$(MsgTxt, K + 137, 3) = recYBIAMVT0.MOUVEMBDF
    Mid$(MsgTxt, K + 140, 1) = recYBIAMVT0.MOUVEMANU
    Mid$(MsgTxt, K + 141, 1) = recYBIAMVT0.MOUVEMRET
    Mid$(MsgTxt, K + 142, 3) = recYBIAMVT0.MOUVEMEVE
    Mid$(MsgTxt, K + 145, 6) = recYBIAMVT0.MOUVEMSAN
    Mid$(MsgTxt, K + 151, 80) = recYBIAMVT0.MOUVEMSAD

MsgTxtLen = MsgTxtLen + recYBIAMVT0Len
End Sub



Public Sub srvYBIAMVT0_ElpDisplay(recYBIAMVT0 As typeYBIAMVT0)
frmElpDisplay.fgData.Rows = 37
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMPLA    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PLAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMPLA
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMCOM   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMCOM
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMMON 17.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMMON
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMDOP    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE D'OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMDOP
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMDVA    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DE VALEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMDVA
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMDCO    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE COMPTABLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMDCO
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMDTR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DE TRAITEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMDTR
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMPIE    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DE PIECE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMPIE
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMECR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO D'ECRITURE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMECR
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMOPE    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMOPE
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMNUM    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMNUM
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMSCH    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE SCHEMA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMSCH
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMUTI    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILISATEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMUTI
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE OPERATRICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMAGE
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMSER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE OPERATEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMSER
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMSSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "S/SERVICE OPERATEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMSSE
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMEXO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE EXONERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMEXO
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMANA    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ANALYTIQUE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMANA
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMBDF    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE BANQUE DE FR."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMBDF
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMANU    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ANNULATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMANU
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMRET    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MOUVEMENT RETRO"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMRET
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMEVE    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EVENEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMEVE
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMSAN    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "STRUCT ANALY-CODE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMSAN
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMSAD   80A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "STRUCT ANALY-DONNEES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.MOUVEMSAD

frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "LIBELLIB1   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Libellé"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.LIBELLIB1
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "LIBELLIB2   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Libellé"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.LIBELLIB2
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "LIBELLIB3   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Libellé"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.LIBELLIB3
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "LIBELLIB4   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Libellé"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.LIBELLIB4

frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEOBL   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPTE OBLIGATOIRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.COMPTEOBL
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEINT   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTITULE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.COMPTEINT
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEDEV    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TABLES BASE 013"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.COMPTEDEV
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTELOR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Lori/Nostri/AUTRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.COMPTELOR
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTECLA    2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CLASSE SECURITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.COMPTECLA
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BIAMVTSD0 18.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDE Précédent"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.BIAMVTSD0
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "BIAMVTID 11s"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Référence"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIAMVT0.BIAMVTID

        
    

frmElpDisplay.Show vbModal
End Sub


Private Function srvYBIAMVT0_Seek(recYBIAMVT0 As typeYBIAMVT0)
'---------------------------------------------------------

srvYBIAMVT0_Seek = "?"
MsgTxtLen = 0
Call srvYBIAMVT0_PutBuffer(recYBIAMVT0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYBIAMVT0_GetBuffer(recYBIAMVT0)) Then
        srvYBIAMVT0_Seek = Null
    Else
        Call srvYBIAMVT0_Error(recYBIAMVT0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYBIAMVT0_Snap(recYBIAMVT0 As typeYBIAMVT0)
'---------------------------------------------------------
srvYBIAMVT0_Snap = "?"
MsgTxtLen = 0
Call srvYBIAMVT0_PutBuffer(recYBIAMVT0)
Call srvYBIAMVT0_PutBuffer(arrYBIAMVT0(0))
If IsNull(SndRcv()) Then
    srvYBIAMVT0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYBIAMVT0_GetBuffer(recYBIAMVT0)) Then
            Call arrYBIAMVT0_AddItem(recYBIAMVT0)
            arrYBIAMVT0_Suite = True
        Else
            arrYBIAMVT0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYBIAMVT0_AddItem(recYBIAMVT0 As typeYBIAMVT0)
'---------------------------------------------------------
          
arrYBIAMVT0_NB = arrYBIAMVT0_NB + 1
    
If arrYBIAMVT0_NB > arrYBIAMVT0_NBMax Then
    arrYBIAMVT0_NBMax = arrYBIAMVT0_NBMax + recYBIAMVT0_Block
    ReDim Preserve arrYBIAMVT0(arrYBIAMVT0_NBMax)
End If
            
arrYBIAMVT0(arrYBIAMVT0_NB) = recYBIAMVT0
End Sub



'---------------------------------------------------------
Public Sub recYBIAMVT0_Init(recYBIAMVT0 As typeYBIAMVT0)
'---------------------------------------------------------
recYBIAMVT0.Obj = "ZMOUVEA0_S"
recYBIAMVT0.Method = ""
recYBIAMVT0.Err = ""
recYBIAMVT0.MOUVEMETA = 1
recYBIAMVT0.MOUVEMETA = 1 '     As Integer                        ' ETABLISSEMENT
recYBIAMVT0.MOUVEMPLA = 0 '      As Long                           ' NUMERO PLAN
recYBIAMVT0.MOUVEMCOM = ""  '   As String * 20                    ' NUMERO COMPTE
recYBIAMVT0.MOUVEMMON = 0 '     As Currency                       ' MONTANT
recYBIAMVT0.MOUVEMDOP = 0 '      As Long                           ' DATE D'OPERATION
recYBIAMVT0.MOUVEMDVA = 0 '      As Long                           ' DATE DE VALEUR
recYBIAMVT0.MOUVEMDCO = 0 '      As Long                           ' DATE COMPTABLE
recYBIAMVT0.MOUVEMDTR = 0 '      As Long                           ' DATE DE TRAITEMENT
recYBIAMVT0.MOUVEMPIE = 0 '     As Long                           ' NUMERO DE PIECE
recYBIAMVT0.MOUVEMECR = 0 '     As Long                           ' NUMERO D'ECRITURE
recYBIAMVT0.MOUVEMOPE = ""   '     As String * 3                     ' CODE OPERATION
recYBIAMVT0.MOUVEMNUM = 0 '    As Long                           ' NUMERO OPERATION
recYBIAMVT0.MOUVEMSCH = 0 '    As Integer                        ' CODE SCHEMA
recYBIAMVT0.MOUVEMUTI = 0 '     As Integer                        ' UTILISATEUR
recYBIAMVT0.MOUVEMAGE = 0 '     As Integer                        ' AGENCE OPERATRICE
recYBIAMVT0.MOUVEMSER = ""   '     As String * 2                     ' SERVICE OPERATEUR
recYBIAMVT0.MOUVEMSSE = ""    '    As String * 2                     ' S/SERVICE OPERATEUR
recYBIAMVT0.MOUVEMEXO = ""   '     As String * 1                     ' CODE EXONERATION
recYBIAMVT0.MOUVEMANA = ""  '      As String * 6                     ' CODE ANALYTIQUE
recYBIAMVT0.MOUVEMBDF = ""  '      As String * 3                     ' CODE BANQUE DE FR.
recYBIAMVT0.MOUVEMANU = ""  '      As String * 1                     ' CODE ANNULATION
recYBIAMVT0.MOUVEMRET = ""  '      As String * 1                     ' MOUVEMENT RETRO
recYBIAMVT0.MOUVEMEVE = ""  '      As String * 3                     ' EVENEMENT
recYBIAMVT0.MOUVEMSAN = ""  '      As String * 6                     ' STRUCT ANALY-CODE
recYBIAMVT0.MOUVEMSAD = ""  '      As String * 80                    ' STRUCT ANALY-DONNEES

End Sub








Public Sub srvYBIAMVT0_Import_File(lFile As String, Nb As Long)
Dim xIn As String, X As String
On Error GoTo Error_Handle



Nb = 0
paramYBIAMVT0_Import = paramYBase_DataF & Trim(lFile) & paramYBase_Data_ExtensionP

Open Trim(paramYBIAMVT0_Import) For Input As #1
meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYBIAMVT0
            meYbase.K1 = mId$(xIn, 10, 20) & mId$(xIn, 72, 8) & mId$(xIn, 419, 11) ' compte DTRT Seq
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close

Exit Sub

Error_Handle:
 MsgBox "erreur : srvYBIAMVT0_Import_File" & xIn, vbCritical, Error
Close


End Sub




