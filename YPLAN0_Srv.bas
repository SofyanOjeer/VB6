Attribute VB_Name = "srvYPLAN0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYPLAN0Len = 149 ' 34 +115
Public Const recYPLAN0_Block = 50
Public Const memoYPLAN0Len = 115
Public Const constYPLAN0 = "YPLAN0"
Public Const constYPLAN0_PRO = "YPLAN0_PRO"
Public paramYPLAN0_Import As String
Dim meYbase As typeYBase

Type typeYPLAN0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    PLANETABL       As Integer                        ' ETABLISSEMENT
    PLANPLAN        As Long                           ' NUMERO PLAN
    PLANCOOBL       As String * 10                    ' COMPTE OBLIGATOIRE
    PLANINTIT       As String * 32                    ' INTITULE
    PLANCOPRO       As String * 3                     ' TABLES BASE 014
    PLANCLASS       As Long                           ' CLASSE SECURITE
    PLANFONCT       As String * 1                     ' TABLES BASE 015
    PLANSESOL       As String * 1                     ' CODE SENS SOLDE D/C
    PLANGEDEP       As String * 1                     ' O/N
    PLANTIERS       As String * 1                     ' COMPTE TIERS O/N
    PLANFICOB       As String * 1                     ' O/N
    PLANCARAC       As Long                           ' 3 à 20
    PLANPESTO       As String * 1                     ' Mois, Trimestre, Année
    PLANNBPER       As Long                           ' 1 à 24
    PLANNBMOU       As Long                           ' NB MVT A CONSERVER
    PLANINEXT       As String * 32                    ' INTITUL EXTRAIT CPT
    PLANPROGR       As String * 8                     ' PROGRAMME DE CONTROL
End Type
    
    
Public arrYPLAN0() As typeYPLAN0
Public arrYPLAN0_NB As Integer
Public arrYPLAN0_NBMax As Integer
Public arrYPLAN0_Index As Integer
Public arrYPLAN0_Suite As Boolean
Public Sub srvYPLAN0_Import_CSV(lFileName As String, lnb As Long)
Dim X As String, K As Integer
Dim xIn As String
Dim blnSelect As Boolean
Dim xYPLAN0 As typeYPLAN0

On Error GoTo Error_Handler

recYPLAN0_Init xYPLAN0

X = lFileName
K = InStr(1, X, ".txt")
Mid$(X, K, 4) = ".csv"
Open X For Input As #1
Open lFileName For Output As #2

lnb = 0


Do Until EOF(1)
    Line Input #1, xIn
    'If Nb Mod 1000 = 0 Then Call lstErr_ChangeLastItem(lstErr, cmdPrint, xIn)
    K = 0
    If mId$(xIn, 1, 2) = "1;" Then
        xYPLAN0.PLANETABL = CInt(CSV_Scan(xIn, K)) '       As Integer                        ' ETABLISSEMENT
        xYPLAN0.PLANPLAN = CLng(CSV_Scan(xIn, K))  '    As Long                           ' NUMERO PLAN
        xYPLAN0.PLANCOOBL = CSV_Scan(xIn, K) '       As String * 10                    ' COMPTE OBLIGATOIRE
        xYPLAN0.PLANINTIT = CSV_Scan(xIn, K) '       As String * 32                    ' INTITULE
        xYPLAN0.PLANCOPRO = CSV_Scan(xIn, K) '       As String * 3                     ' TABLES BASE 014
        xYPLAN0.PLANCLASS = CLng(CSV_Scan(xIn, K)) '     As Long                           ' CLASSE SECURITE
        xYPLAN0.PLANFONCT = CSV_Scan(xIn, K) '       As String * 1                     ' TABLES BASE 015
        xYPLAN0.PLANSESOL = CSV_Scan(xIn, K) '       As String * 1                     ' CODE SENS SOLDE D/C
        xYPLAN0.PLANGEDEP = CSV_Scan(xIn, K) '       As String * 1                     ' O/N
        xYPLAN0.PLANTIERS = CSV_Scan(xIn, K) '       As String * 1                     ' COMPTE TIERS O/N
        xYPLAN0.PLANFICOB = CSV_Scan(xIn, K) '       As String * 1                     ' O/N
        xYPLAN0.PLANCARAC = CLng(CSV_Scan(xIn, K))   '  As Long                           ' 3 à 20
        xYPLAN0.PLANPESTO = CSV_Scan(xIn, K) ''       As String * 1                     ' Mois, Trimestre, Année
        xYPLAN0.PLANNBPER = CLng(CSV_Scan(xIn, K)) '     As Long                           ' 1 à 24
        xYPLAN0.PLANNBMOU = CLng(CSV_Scan(xIn, K))  '   As Long                           ' NB MVT A CONSERVER
        xYPLAN0.PLANINEXT = CSV_Scan(xIn, K) '       As String * 32                    ' INTITUL EXTRAIT CPT
        xYPLAN0.PLANPROGR = CSV_Scan(xIn, K) '       As String * 8                     ' PROGRAMME DE CONTROL
    
        MsgTxtLen = 0
        srvYPLAN0_PutBuffer xYPLAN0
        lnb = lnb + 1
        
        Print #2, mId$(MsgTxt, 35, memoYPLAN0Len)
    End If
    DoEvents

Loop

Close
'Call lstErr_ChangeLastItem(lstErr, cmdPrint, "cmdInfo_YPLAN0_025 : " & NbOk & "/" & Nb): DoEvents
ReDim arrYBASTAU0_Key(1)

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Close
Call MsgBox(Err & " : " & Error(Err), vbCritical, "cmdInfo_YPLAN0_025")

End Sub



Public Sub srvYPLAN0_ElpDisplay(recYPLAN0 As typeYPLAN0)
frmElpDisplay.fgData.Rows = 18
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANETABL    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANETABL
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANPLAN    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PLAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANPLAN
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANCOOBL   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPTE OBLIGATOIRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANCOOBL
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANINTIT   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTITULE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANINTIT
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANCOPRO    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TABLES BASE 014"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANCOPRO
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANCLASS    2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CLASSE SECURITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANCLASS
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANFONCT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TABLES BASE 015"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANFONCT
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANSESOL    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE SENS SOLDE D/C"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANSESOL
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANGEDEP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANGEDEP
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANTIERS    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPTE TIERS O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANTIERS
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANFICOB    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANFICOB
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANCARAC    2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "3 à 20"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANCARAC
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANPESTO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Mois, Trimestre, Année"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANPESTO
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANNBPER    2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "1 à 24"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANNBPER
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANNBMOU    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NB MVT A CONSERVER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANNBMOU
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANINEXT   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTITUL EXTRAIT CPT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANINEXT
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANPROGR    8A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PROGRAMME DE CONTROL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYPLAN0.PLANPROGR
frmElpDisplay.Show vbModal
End Sub
Public Sub srvYPLAN0_Export_CSV(lIdFile_Source As Integer, lIdFile_Destination As Integer, loptSelect_CSV_Header As Boolean, lnb As Long)
Dim xIn As String
'Open "C:\Temp\YPLAN0.txt" For Input As #1
'Open "C:\Temp\YPLAN0.csv" For Output As #2
If loptSelect_CSV_Header Then
    Print #lIdFile_Destination, "PLANETABL;PLANPLAN;PLANCOOBL;PLANINTIT;PLANCOPRO;PLANCLASS;PLANFONCT;PLANSESOL;PLANGEDEP;PLANTIERS;PLANFICOB;PLANCARAC;PLANPESTO;PLANNBPER;PLANNBMOU;PLANINEXT;PLANPROGR;"
    Print #lIdFile_Destination, "ETABLISSEMENT;NUMERO PLAN;COMPTE OBLIGATOIRE;INTITULE;CODE PRODUIT;CLASSE SECURITE;CODE FONCTIONNEMENT;CODE SENS SOLDE D/C;GESTION DEPASSEMENT;COMPTE TIERS O/N;COMPTE DE CLIENTELE;NB CARACTERE COMPTE;PERIOD STOCKAGE MVT;NB PERIODE STOCKAGE;NB MVT A CONSERVER;INTITUL EXTRAIT CPT;PROGRAMME DE CONTROL;"
    Print #lIdFile_Destination, ";;;;;;;;;;;;;;;;;"
End If
Do Until EOF(lIdFile_Source)
      Line Input #lIdFile_Source, xIn
      lnb = lnb + 1
      Print #lIdFile_Destination, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 4) & ";" _
      & mId$(xIn, 10, 10) & ";" _
      & mId$(xIn, 20, 32) & ";" _
      & mId$(xIn, 52, 3) & ";" _
      & mId$(xIn, 55, 3) & ";" _
      & mId$(xIn, 58, 1) & ";" _
      & mId$(xIn, 59, 1) & ";" _
      & mId$(xIn, 60, 1) & ";" _
      & mId$(xIn, 61, 1) & ";" _
      & mId$(xIn, 62, 1) & ";" _
      & mId$(xIn, 63, 3) & ";" _
      & mId$(xIn, 66, 1) & ";" _
      & mId$(xIn, 67, 3) & ";" _
      & mId$(xIn, 70, 6) & ";" _
      & mId$(xIn, 76, 32) & ";" _
      & mId$(xIn, 108, 8) & ";"
Loop
End Sub

'-----------------------------------------------------
Function srvYPLAN0_Update(recYPLAN0 As typeYPLAN0)
'-----------------------------------------------------

srvYPLAN0_Update = "?"

MsgTxtLen = 0
Call srvYPLAN0_PutBuffer(recYPLAN0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYPLAN0_GetBuffer(recYPLAN0)) Then
        Call srvYPLAN0_Error(recYPLAN0)
        srvYPLAN0_Update = recYPLAN0.Err
        Exit Function
    Else
        srvYPLAN0_Update = Null
    End If
Else
    recYPLAN0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYPLAN0_Error(recYPLAN0 As typeYPLAN0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YPLAN0" & Chr$(10) & Chr$(13)

Select Case mId$(recYPLAN0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYPLAN0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YPLAN0s.bas  ( " & Trim(recYPLAN0.Obj) & " : " & Trim(recYPLAN0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYPLAN0_Monitor(recYPLAN0 As typeYPLAN0)
'-----------------------------------------------------

arrYPLAN0_Suite = False
Select Case mId$(Trim(recYPLAN0.Method), 1, 4)
    Case "Snap"
              srvYPLAN0_Monitor = srvYPLAN0_Snap(recYPLAN0)
    Case Else
            srvYPLAN0_Monitor = srvYPLAN0_Seek(recYPLAN0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYPLAN0_GetBuffer(recYPLAN0 As typeYPLAN0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYPLAN0_GetBuffer = Null
recYPLAN0.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYPLAN0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYPLAN0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYPLAN0.Err = Space$(10) Then
    recYPLAN0.PLANETABL = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYPLAN0.PLANPLAN = CLng(Val(mId$(MsgTxt, K + 6, 4)))
    recYPLAN0.PLANCOOBL = mId$(MsgTxt, K + 10, 10)
    recYPLAN0.PLANINTIT = mId$(MsgTxt, K + 20, 32)
    recYPLAN0.PLANCOPRO = mId$(MsgTxt, K + 52, 3)
    recYPLAN0.PLANCLASS = CLng(Val(mId$(MsgTxt, K + 55, 3)))
    recYPLAN0.PLANFONCT = mId$(MsgTxt, K + 58, 1)
    recYPLAN0.PLANSESOL = mId$(MsgTxt, K + 59, 1)
    recYPLAN0.PLANGEDEP = mId$(MsgTxt, K + 60, 1)
    recYPLAN0.PLANTIERS = mId$(MsgTxt, K + 61, 1)
    recYPLAN0.PLANFICOB = mId$(MsgTxt, K + 62, 1)
    recYPLAN0.PLANCARAC = CLng(Val(mId$(MsgTxt, K + 63, 3)))
    recYPLAN0.PLANPESTO = mId$(MsgTxt, K + 66, 1)
    recYPLAN0.PLANNBPER = CLng(Val(mId$(MsgTxt, K + 67, 3)))
    recYPLAN0.PLANNBMOU = CLng(Val(mId$(MsgTxt, K + 70, 6)))
    recYPLAN0.PLANINEXT = mId$(MsgTxt, K + 76, 32)
    recYPLAN0.PLANPROGR = mId$(MsgTxt, K + 108, 8)
Else
    srvYPLAN0_GetBuffer = recYPLAN0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYPLAN0Len

End Function

Public Function srvYPLAN0_Import(lX As String)
Dim xIn As String, X As String, Nb As String
On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = constYPLAN0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    lX = meYbase.Text
    If mId$(lX, 1, 8) >= YBIATAB0_DATE_CPT_J Then
        srvYPLAN0_Import = Null
        Exit Function
    Else
        meYbase.Method = constDelete
        Call tableYBase_Update(meYbase)
    End If
End If


srvYPLAN0_Import = "?"

paramYPLAN0_Import = paramYBase_DataF & Trim(constYPLAN0) & paramYBase_Data_ExtensionP

Open Trim(paramYPLAN0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYPLAN0) & Chr$(34)
MDB.Execute X
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYPLAN0_PRO) & Chr$(34)
MDB.Execute X
meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
            meYbase.ID = constYPLAN0
            meYbase.K1 = mId$(xIn, 10, 10)
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
            meYbase.ID = constYPLAN0_PRO
            meYbase.K1 = mId$(xIn, 52, 3) & mId$(xIn, 10, 10)
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYPLAN0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = constYPLAN0
meYbase.Text = YBIATAB0_DATE_CPT_J & "_" & DSys & "_" & time_Hms & "_" & Format$(Nb, "000000000")
lX = meYbase.Text
meYbase.Text = lX

dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYPLAN0_Import" & xIn, vbCritical, Error
Close

srvYPLAN0_Import = Error
End Function


Public Function srvYPLAN0_Import_Read(lId As String, lYPLAN0 As typeYPLAN0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYPLAN0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYPLAN0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYPLAN0_GetBuffer lYPLAN0
    srvYPLAN0_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYPLAN0_Import_Read" & xIn, vbCritical, Error
srvYPLAN0_Import_Read = Error
End Function


'---------------------------------------------------------
Public Sub srvYPLAN0_PutBuffer(recYPLAN0 As typeYPLAN0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYPLAN0.Obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYPLAN0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34


    Mid$(MsgTxt, K + 1, 5) = Format$(recYPLAN0.PLANETABL, "0000 ")
    Mid$(MsgTxt, K + 6, 4) = Format$(recYPLAN0.PLANPLAN, "000 ")
    Mid$(MsgTxt, K + 10, 10) = recYPLAN0.PLANCOOBL
    Mid$(MsgTxt, K + 20, 32) = recYPLAN0.PLANINTIT
    Mid$(MsgTxt, K + 52, 3) = recYPLAN0.PLANCOPRO
    Mid$(MsgTxt, K + 55, 3) = Format$(recYPLAN0.PLANCLASS, "00 ")
    Mid$(MsgTxt, K + 58, 1) = recYPLAN0.PLANFONCT
    Mid$(MsgTxt, K + 59, 1) = recYPLAN0.PLANSESOL
    Mid$(MsgTxt, K + 60, 1) = recYPLAN0.PLANGEDEP
    Mid$(MsgTxt, K + 61, 1) = recYPLAN0.PLANTIERS
    Mid$(MsgTxt, K + 62, 1) = recYPLAN0.PLANFICOB
    Mid$(MsgTxt, K + 63, 3) = Format$(recYPLAN0.PLANCARAC, "00 ")
    Mid$(MsgTxt, K + 66, 1) = recYPLAN0.PLANPESTO
    Mid$(MsgTxt, K + 67, 3) = Format$(recYPLAN0.PLANNBPER, "00 ")
    Mid$(MsgTxt, K + 70, 6) = Format$(recYPLAN0.PLANNBMOU, "00000 ")
    Mid$(MsgTxt, K + 76, 32) = recYPLAN0.PLANINEXT
    Mid$(MsgTxt, K + 108, 8) = recYPLAN0.PLANPROGR
MsgTxtLen = MsgTxtLen + recYPLAN0Len
End Sub



'---------------------------------------------------------
Private Function srvYPLAN0_Seek(recYPLAN0 As typeYPLAN0)
'---------------------------------------------------------

srvYPLAN0_Seek = "?"
MsgTxtLen = 0
Call srvYPLAN0_PutBuffer(recYPLAN0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYPLAN0_GetBuffer(recYPLAN0)) Then
        srvYPLAN0_Seek = Null
    Else
        Call srvYPLAN0_Error(recYPLAN0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYPLAN0_Snap(recYPLAN0 As typeYPLAN0)
'---------------------------------------------------------
srvYPLAN0_Snap = "?"
MsgTxtLen = 0
Call srvYPLAN0_PutBuffer(recYPLAN0)
Call srvYPLAN0_PutBuffer(arrYPLAN0(0))
If IsNull(SndRcv()) Then
    srvYPLAN0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYPLAN0_GetBuffer(recYPLAN0)) Then
            Call arrYPLAN0_AddItem(recYPLAN0)
            arrYPLAN0_Suite = True
        Else
            arrYPLAN0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYPLAN0_AddItem(recYPLAN0 As typeYPLAN0)
'---------------------------------------------------------
          
arrYPLAN0_NB = arrYPLAN0_NB + 1
    
If arrYPLAN0_NB > arrYPLAN0_NBMax Then
    arrYPLAN0_NBMax = arrYPLAN0_NBMax + recYPLAN0_Block
    ReDim Preserve arrYPLAN0(arrYPLAN0_NBMax)
End If
            
arrYPLAN0(arrYPLAN0_NB) = recYPLAN0
End Sub



'---------------------------------------------------------
Public Sub recYPLAN0_Init(recYPLAN0 As typeYPLAN0)
'---------------------------------------------------------
recYPLAN0.Obj = "ZCDOTIE0_S"
recYPLAN0.Method = ""
recYPLAN0.Err = ""
recYPLAN0.PLANETABL = 0 '       As Integer                        ' ETABLISSEMENT
recYPLAN0.PLANPLAN = 0  '    As Long                           ' NUMERO PLAN
recYPLAN0.PLANCOOBL = "" '       As String * 10                    ' COMPTE OBLIGATOIRE
recYPLAN0.PLANINTIT = "" '       As String * 32                    ' INTITULE
recYPLAN0.PLANCOPRO = "" '       As String * 3                     ' TABLES BASE 014
recYPLAN0.PLANCLASS = 0 '     As Long                           ' CLASSE SECURITE
recYPLAN0.PLANFONCT = "" '       As String * 1                     ' TABLES BASE 015
recYPLAN0.PLANSESOL = "" '       As String * 1                     ' CODE SENS SOLDE D/C
recYPLAN0.PLANGEDEP = "" '       As String * 1                     ' O/N
recYPLAN0.PLANTIERS = "" '       As String * 1                     ' COMPTE TIERS O/N
recYPLAN0.PLANFICOB = "" '       As String * 1                     ' O/N
recYPLAN0.PLANCARAC = 0   '  As Long                           ' 3 à 20
recYPLAN0.PLANPESTO = "" ''       As String * 1                     ' Mois, Trimestre, Année
recYPLAN0.PLANNBPER = 0 '     As Long                           ' 1 à 24
recYPLAN0.PLANNBMOU = 0  '   As Long                           ' NB MVT A CONSERVER
recYPLAN0.PLANINEXT = "" '       As String * 32                    ' INTITUL EXTRAIT CPT
recYPLAN0.PLANPROGR = "" '       As String * 8                     ' PROGRAMME DE CONTROL

End Sub










Public Sub srvYPLAN0_Import_cboPCEC(lCbo As ComboBox)

lCbo.Clear
lCbo.AddItem " "

meYbase.ID = constYPLAN0
meYbase.K1 = ""
meYbase.Method = "Seek>"
Do
    intReturn = tableYBase_Read(meYbase)
    If Trim(meYbase.ID) <> constYPLAN0 Then intReturn = -1
    If intReturn = 0 Then
        lCbo.AddItem Trim(meYbase.K1)
    End If
        
Loop Until intReturn <> 0


End Sub
Public Sub srvYPLAN0_Import_cboPLANCOPRO(lCbo As ComboBox)
Dim X3 As String * 3
lCbo.Clear
lCbo.AddItem " "
X3 = " "
meYbase.ID = constYPLAN0_PRO
meYbase.K1 = ""
meYbase.Method = "Seek>"
Do
    intReturn = tableYBase_Read(meYbase)
    If Trim(meYbase.ID) <> constYPLAN0_PRO Then intReturn = -1
    If intReturn = 0 Then
        If X3 <> mId$(meYbase.Text, 52, 3) Then
            X3 = mId$(meYbase.Text, 52, 3)
            lCbo.AddItem X3
        End If
    End If
        
Loop Until intReturn <> 0


End Sub

