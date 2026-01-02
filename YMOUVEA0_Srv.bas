Attribute VB_Name = "srvYMOUVEA0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYMOUVEA0Len = 264 ' 34 +230
Public Const recYMOUVEA0_Block = 100
Public Const memoYMOUVEA0Len = 230
Public Const constYMOUVEA0 = "YMOUVEA0  "

Type typeYMOUVEA0
    obj                     As String * 12
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
    
End Type
    
    
Public arrYMOUVEA0() As typeYMOUVEA0
Public arrYMOUVEA0_NB As Integer
Public arrYMOUVEA0_NBMax As Integer
Public arrYMOUVEA0_Index As Integer
Public arrYMOUVEA0_Suite As Boolean

'-----------------------------------------------------
Function srvYMOUVEA0_Update(recYMOUVEA0 As typeYMOUVEA0)
'-----------------------------------------------------

srvYMOUVEA0_Update = "?"

MsgTxtLen = 0
Call srvYMOUVEA0_PutBuffer(recYMOUVEA0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYMOUVEA0_GetBuffer(recYMOUVEA0)) Then
        Call srvYMOUVEA0_Error(recYMOUVEA0)
        srvYMOUVEA0_Update = recYMOUVEA0.Err
        Exit Function
    Else
        srvYMOUVEA0_Update = Null
    End If
Else
    recYMOUVEA0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYMOUVEA0_Error(recYMOUVEA0 As typeYMOUVEA0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YMOUVEA0" & Chr$(10) & Chr$(13)

Select Case mId$(recYMOUVEA0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYMOUVEA0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YMOUVEA0s.bas  ( " & Trim(recYMOUVEA0.obj) & " : " & Trim(recYMOUVEA0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYMOUVEA0_Monitor(recYMOUVEA0 As typeYMOUVEA0)
'-----------------------------------------------------

arrYMOUVEA0_Suite = False
Select Case mId$(Trim(recYMOUVEA0.Method), 1, 4)
    Case "Snap"
              srvYMOUVEA0_Monitor = srvYMOUVEA0_Snap(recYMOUVEA0)
    Case Else
            srvYMOUVEA0_Monitor = srvYMOUVEA0_Seek(recYMOUVEA0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYMOUVEA0_GetBuffer(recYMOUVEA0 As typeYMOUVEA0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYMOUVEA0_GetBuffer = Null
recYMOUVEA0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYMOUVEA0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYMOUVEA0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYMOUVEA0.Err = Space$(10) Then
    recYMOUVEA0.MOUVEMETA = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYMOUVEA0.MOUVEMPLA = CLng(Val(mId$(MsgTxt, K + 6, 4)))
    recYMOUVEA0.MOUVEMCOM = mId$(MsgTxt, K + 10, 20)
    recYMOUVEA0.MOUVEMMON = CCur(mId$(MsgTxt, K + 30, 18)) / 1000
    recYMOUVEA0.MOUVEMDOP = CLng(Val(mId$(MsgTxt, K + 48, 8)))
    recYMOUVEA0.MOUVEMDVA = CLng(Val(mId$(MsgTxt, K + 56, 8)))
    recYMOUVEA0.MOUVEMDCO = CLng(Val(mId$(MsgTxt, K + 64, 8)))
    recYMOUVEA0.MOUVEMDTR = CLng(Val(mId$(MsgTxt, K + 72, 8)))
    recYMOUVEA0.MOUVEMPIE = CLng(Val(mId$(MsgTxt, K + 80, 10)))
    recYMOUVEA0.MOUVEMECR = CLng(Val(mId$(MsgTxt, K + 90, 8)))
    recYMOUVEA0.MOUVEMOPE = mId$(MsgTxt, K + 98, 3)
    recYMOUVEA0.MOUVEMNUM = CLng(Val(mId$(MsgTxt, K + 101, 10)))
    recYMOUVEA0.MOUVEMSCH = CInt(Val(mId$(MsgTxt, K + 111, 5)))
    recYMOUVEA0.MOUVEMUTI = CInt(Val(mId$(MsgTxt, K + 116, 5)))
    recYMOUVEA0.MOUVEMAGE = CInt(Val(mId$(MsgTxt, K + 121, 5)))
    recYMOUVEA0.MOUVEMSER = mId$(MsgTxt, K + 126, 2)
    recYMOUVEA0.MOUVEMSSE = mId$(MsgTxt, K + 128, 2)
    recYMOUVEA0.MOUVEMEXO = mId$(MsgTxt, K + 130, 1)
    recYMOUVEA0.MOUVEMANA = mId$(MsgTxt, K + 131, 6)
    recYMOUVEA0.MOUVEMBDF = mId$(MsgTxt, K + 137, 3)
    recYMOUVEA0.MOUVEMANU = mId$(MsgTxt, K + 140, 1)
    recYMOUVEA0.MOUVEMRET = mId$(MsgTxt, K + 141, 1)
    recYMOUVEA0.MOUVEMEVE = mId$(MsgTxt, K + 142, 3)
    recYMOUVEA0.MOUVEMSAN = mId$(MsgTxt, K + 145, 6)
    recYMOUVEA0.MOUVEMSAD = mId$(MsgTxt, K + 151, 80)
Else
    srvYMOUVEA0_GetBuffer = recYMOUVEA0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYMOUVEA0Len

End Function

'---------------------------------------------------------
Public Sub srvYMOUVEA0_PutBuffer(recYMOUVEA0 As typeYMOUVEA0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYMOUVEA0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYMOUVEA0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYMOUVEA0.MOUVEMETA, "0000 ")
    Mid$(MsgTxt, K + 6, 4) = Format$(recYMOUVEA0.MOUVEMPLA, "000 ")
    Mid$(MsgTxt, K + 10, 20) = recYMOUVEA0.MOUVEMCOM
    Mid$(MsgTxt, K + 30, 18) = Format$(recYMOUVEA0.MOUVEMMON * 1000, "00000000000000000 ")
    Mid$(MsgTxt, K + 48, 8) = Format$(recYMOUVEA0.MOUVEMDOP, "0000000 ")
    Mid$(MsgTxt, K + 56, 8) = Format$(recYMOUVEA0.MOUVEMDVA, "0000000 ")
    Mid$(MsgTxt, K + 64, 8) = Format$(recYMOUVEA0.MOUVEMDCO, "0000000 ")
    Mid$(MsgTxt, K + 72, 8) = Format$(recYMOUVEA0.MOUVEMDTR, "0000000 ")
    Mid$(MsgTxt, K + 80, 10) = Format$(recYMOUVEA0.MOUVEMPIE, "000000000 ")
    Mid$(MsgTxt, K + 90, 8) = Format$(recYMOUVEA0.MOUVEMECR, "0000000 ")
    Mid$(MsgTxt, K + 98, 3) = recYMOUVEA0.MOUVEMOPE
    Mid$(MsgTxt, K + 101, 10) = Format$(recYMOUVEA0.MOUVEMNUM, "000000000 ")
    Mid$(MsgTxt, K + 111, 5) = Format$(recYMOUVEA0.MOUVEMSCH, "0000 ")
    Mid$(MsgTxt, K + 116, 5) = Format$(recYMOUVEA0.MOUVEMUTI, "0000 ")
    Mid$(MsgTxt, K + 121, 5) = Format$(recYMOUVEA0.MOUVEMAGE, "0000 ")
    Mid$(MsgTxt, K + 126, 2) = recYMOUVEA0.MOUVEMSER
    Mid$(MsgTxt, K + 128, 2) = recYMOUVEA0.MOUVEMSSE
    Mid$(MsgTxt, K + 130, 1) = recYMOUVEA0.MOUVEMEXO
    Mid$(MsgTxt, K + 131, 6) = recYMOUVEA0.MOUVEMANA
    Mid$(MsgTxt, K + 137, 3) = recYMOUVEA0.MOUVEMBDF
    Mid$(MsgTxt, K + 140, 1) = recYMOUVEA0.MOUVEMANU
    Mid$(MsgTxt, K + 141, 1) = recYMOUVEA0.MOUVEMRET
    Mid$(MsgTxt, K + 142, 3) = recYMOUVEA0.MOUVEMEVE
    Mid$(MsgTxt, K + 145, 6) = recYMOUVEA0.MOUVEMSAN
    Mid$(MsgTxt, K + 151, 80) = recYMOUVEA0.MOUVEMSAD

MsgTxtLen = MsgTxtLen + recYMOUVEA0Len
End Sub



Public Sub srvYMOUVEA0_ElpDisplay(recYMOUVEA0 As typeYMOUVEA0)
frmElpDisplay.fgData.Rows = 26
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMPLA    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PLAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMPLA
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMCOM   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMCOM
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMMON 17.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMMON
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMDOP    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE D'OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMDOP
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMDVA    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DE VALEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMDVA
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMDCO    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE COMPTABLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMDCO
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMDTR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DE TRAITEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMDTR
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMPIE    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DE PIECE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMPIE
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMECR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO D'ECRITURE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMECR
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMOPE    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMOPE
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMNUM    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMNUM
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMSCH    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE SCHEMA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMSCH
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMUTI    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILISATEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMUTI
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE OPERATRICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMAGE
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMSER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE OPERATEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMSER
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMSSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "S/SERVICE OPERATEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMSSE
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMEXO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE EXONERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMEXO
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMANA    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ANALYTIQUE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMANA
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMBDF    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE BANQUE DE FR."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMBDF
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMANU    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ANNULATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMANU
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMRET    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MOUVEMENT RETRO"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMRET
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMEVE    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EVENEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMEVE
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMSAN    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "STRUCT ANALY-CODE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMSAN
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MOUVEMSAD   80A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "STRUCT ANALY-DONNEES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMOUVEA0.MOUVEMSAD
frmElpDisplay.Show vbModal
End Sub

Public Sub srvYMOUVEA0_Export_CSV()
Dim xIn As String
Open "D:\Temp\FTP\YMOUVEA0.txt" For Input As #1
Open "D:\Temp\FTP\YMOUVEA0.csv" For Output As #2
Print #2, "MOUVEMETA;MOUVEMPLA;MOUVEMCOM;MOUVEMMON;MOUVEMDOP;MOUVEMDVA;MOUVEMDCO;MOUVEMDTR;MOUVEMPIE;MOUVEMECR;MOUVEMOPE;MOUVEMNUM;MOUVEMSCH;MOUVEMUTI;MOUVEMAGE;MOUVEMSER;MOUVEMSSE;MOUVEMEXO;MOUVEMANA;MOUVEMBDF;MOUVEMANU;MOUVEMRET;MOUVEMEVE;MOUVEMSAN;MOUVEMSAD;"
Print #2, "ETABLISSEMENT;NUMERO PLAN;NUMERO COMPTE;MONTANT;DATE D'OPERATION;DATE DE VALEUR;DATE COMPTABLE;DATE DE TRAITEMENT;NUMERO DE PIECE;NUMERO D'ECRITURE;CODE OPERATION;NUMERO OPERATION;CODE SCHEMA;UTILISATEUR;AGENCE OPERATRICE;SERVICE OPERATEUR;S/SERVICE OPERATEUR;CODE EXONERATION;CODE ANALYTIQUE;CODE BANQUE DE FR.;CODE ANNULATION;MOUVEMENT RETRO;EVENEMENT;STRUCT ANALY-CODE;STRUCT ANALY-DONNEES;"
Print #2, ";;;;;;;;;;;;;;;;;;;;;;;;;"
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 4) & ";" _
      & mId$(xIn, 10, 20) & ";" _
      & mId$(xIn, 30, 18) & ";" _
      & mId$(xIn, 48, 8) & ";" _
      & mId$(xIn, 56, 8) & ";" _
      & mId$(xIn, 64, 8) & ";" _
      & mId$(xIn, 72, 8) & ";" _
      & mId$(xIn, 80, 10) & ";" _
      & mId$(xIn, 90, 8) & ";" _
      & mId$(xIn, 98, 3) & ";" _
      & mId$(xIn, 101, 10) & ";" _
      & mId$(xIn, 111, 5) & ";" _
      & mId$(xIn, 116, 5) & ";" _
      & mId$(xIn, 121, 5) & ";" _
      & mId$(xIn, 126, 2) & ";" _
      & mId$(xIn, 128, 2) & ";" _
      & mId$(xIn, 130, 1) & ";" _
      & mId$(xIn, 131, 6) & ";" _
      & mId$(xIn, 137, 3) & ";" _
      & mId$(xIn, 140, 1) & ";" _
      & mId$(xIn, 141, 1) & ";" _
      & mId$(xIn, 142, 3) & ";" _
      & mId$(xIn, 145, 6) & ";" _
      & mId$(xIn, 151, 80) & ";"
Loop
Close
End Sub

'---------------------------------------------------------
Private Function srvYMOUVEA0_Seek(recYMOUVEA0 As typeYMOUVEA0)
'---------------------------------------------------------

srvYMOUVEA0_Seek = "?"
MsgTxtLen = 0
Call srvYMOUVEA0_PutBuffer(recYMOUVEA0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYMOUVEA0_GetBuffer(recYMOUVEA0)) Then
        srvYMOUVEA0_Seek = Null
    Else
        Call srvYMOUVEA0_Error(recYMOUVEA0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYMOUVEA0_Snap(recYMOUVEA0 As typeYMOUVEA0)
'---------------------------------------------------------
srvYMOUVEA0_Snap = "?"
MsgTxtLen = 0
Call srvYMOUVEA0_PutBuffer(recYMOUVEA0)
Call srvYMOUVEA0_PutBuffer(arrYMOUVEA0(0))
If IsNull(SndRcv()) Then
    srvYMOUVEA0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYMOUVEA0_GetBuffer(recYMOUVEA0)) Then
            Call arrYMOUVEA0_AddItem(recYMOUVEA0)
            arrYMOUVEA0_Suite = True
        Else
            arrYMOUVEA0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYMOUVEA0_AddItem(recYMOUVEA0 As typeYMOUVEA0)
'---------------------------------------------------------
          
arrYMOUVEA0_NB = arrYMOUVEA0_NB + 1
    
If arrYMOUVEA0_NB > arrYMOUVEA0_NBMax Then
    arrYMOUVEA0_NBMax = arrYMOUVEA0_NBMax + recYMOUVEA0_Block
    ReDim Preserve arrYMOUVEA0(arrYMOUVEA0_NBMax)
End If
            
arrYMOUVEA0(arrYMOUVEA0_NB) = recYMOUVEA0
End Sub



'---------------------------------------------------------
Public Sub recYMOUVEA0_Init(recYMOUVEA0 As typeYMOUVEA0)
'---------------------------------------------------------
recYMOUVEA0.obj = "ZMOUVEA0_S"
recYMOUVEA0.Method = ""
recYMOUVEA0.Err = ""
recYMOUVEA0.MOUVEMETA = 1
recYMOUVEA0.MOUVEMETA = 1 '     As Integer                        ' ETABLISSEMENT
recYMOUVEA0.MOUVEMPLA = 0 '      As Long                           ' NUMERO PLAN
recYMOUVEA0.MOUVEMCOM = ""  '   As String * 20                    ' NUMERO COMPTE
recYMOUVEA0.MOUVEMMON = 0 '     As Currency                       ' MONTANT
recYMOUVEA0.MOUVEMDOP = 0 '      As Long                           ' DATE D'OPERATION
recYMOUVEA0.MOUVEMDVA = 0 '      As Long                           ' DATE DE VALEUR
recYMOUVEA0.MOUVEMDCO = 0 '      As Long                           ' DATE COMPTABLE
recYMOUVEA0.MOUVEMDTR = 0 '      As Long                           ' DATE DE TRAITEMENT
recYMOUVEA0.MOUVEMPIE = 0 '     As Long                           ' NUMERO DE PIECE
recYMOUVEA0.MOUVEMECR = 0 '     As Long                           ' NUMERO D'ECRITURE
recYMOUVEA0.MOUVEMOPE = ""   '     As String * 3                     ' CODE OPERATION
recYMOUVEA0.MOUVEMNUM = 0 '    As Long                           ' NUMERO OPERATION
recYMOUVEA0.MOUVEMSCH = 0 '    As Integer                        ' CODE SCHEMA
recYMOUVEA0.MOUVEMUTI = 0 '     As Integer                        ' UTILISATEUR
recYMOUVEA0.MOUVEMAGE = 0 '     As Integer                        ' AGENCE OPERATRICE
recYMOUVEA0.MOUVEMSER = ""   '     As String * 2                     ' SERVICE OPERATEUR
recYMOUVEA0.MOUVEMSSE = ""    '    As String * 2                     ' S/SERVICE OPERATEUR
recYMOUVEA0.MOUVEMEXO = ""   '     As String * 1                     ' CODE EXONERATION
recYMOUVEA0.MOUVEMANA = ""  '      As String * 6                     ' CODE ANALYTIQUE
recYMOUVEA0.MOUVEMBDF = ""  '      As String * 3                     ' CODE BANQUE DE FR.
recYMOUVEA0.MOUVEMANU = ""  '      As String * 1                     ' CODE ANNULATION
recYMOUVEA0.MOUVEMRET = ""  '      As String * 1                     ' MOUVEMENT RETRO
recYMOUVEA0.MOUVEMEVE = ""  '      As String * 3                     ' EVENEMENT
recYMOUVEA0.MOUVEMSAN = ""  '      As String * 6                     ' STRUCT ANALY-CODE
recYMOUVEA0.MOUVEMSAD = ""  '      As String * 80                    ' STRUCT ANALY-DONNEES

End Sub







Public Function srvYMOUVEA0_YCOMPTE0(lYMOUVEA0 As typeYMOUVEA0, lYCOMPTE0 As typeYCOMPTE0)
Dim xMVTP0 As typeMvtP0
Dim intReturn As Integer

srvYMOUVEA0_YCOMPTE0 = Null
xMVTP0.Id = constYCOMPTE0 & Format$(lYMOUVEA0.MOUVEMETA, "0000 ") _
                          & Format$(lYMOUVEA0.MOUVEMPLA, "000 ") _
                          & lYMOUVEA0.MOUVEMCOM
xMVTP0.Method = "Seek="
intReturn = tableMvtP0_Read(xMVTP0)
If intReturn = 0 Then
    MsgTxt = Space$(34) & xMVTP0.Text
    MsgTxtIndex = 0
              
    srvYCOMPTE0_GetBuffer lYCOMPTE0
Else
    lYCOMPTE0.COMPTEINT = "?????????????????"
    srvYMOUVEA0_YCOMPTE0 = xMVTP0.Id & " => Erreur : " & xMVTP0.Err
End If

End Function
Public Function srvYMOUVEA0_YLIBEL0(lYMOUVEA0 As typeYMOUVEA0, lLIBELNUM As Integer, lYLIBEL0 As typeYLIBEL0)
Dim xMVTP0 As typeMvtP0
Dim intReturn As Integer

srvYMOUVEA0_YLIBEL0 = Null
xMVTP0.Id = constYLIBEL0 & Format$(lYMOUVEA0.MOUVEMETA, "0000 ") _
                          & Format$(lYMOUVEA0.MOUVEMPIE, "000000000 ") _
                          & Format$(lYMOUVEA0.MOUVEMECR, "0000000 ") _
                          & Format$(lLIBELNUM, "0 ")
xMVTP0.Method = "Seek="
intReturn = tableMvtP0_Read(xMVTP0)
If intReturn = 0 Then
    MsgTxt = Space$(34) & xMVTP0.Text
    MsgTxtIndex = 0
              
    srvYLIBEL0_GetBuffer lYLIBEL0
Else
    lYLIBEL0.LIBELLIB = "?????????????????"
    srvYMOUVEA0_YLIBEL0 = xMVTP0.Id & " => Erreur : " & xMVTP0.Err
End If

End Function


