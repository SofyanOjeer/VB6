Attribute VB_Name = "srvYCOMPTE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCOMPTE0Len = 184 + 50 ' 34 +150 +50
Public Const recYCOMPTE0_Block = 100
Public Const memoYCOMPTE0Len = 150 + 50
Public Const constYCOMPTE0 = "YCOMPTE0"
Public paramYCOMPTE0_Import As String
Dim meYbase As typeYBase
Public paramYCOMPTE0_Nb As Long

Dim xYBase_YSOLDE0 As typeYBase
Dim xYBase_YPLAN0 As typeYBase

Type typeYCOMPTE0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    COMPTEETA       As Integer                        ' ETABLISSEMENT
    COMPTEPLA       As Long                           ' NUMERO PLAN
    COMPTECOM       As String * 20                    ' NUMERO COMPTE
    COMPTEOBL       As String * 10                    ' COMPTE OBLIGATOIRE
    COMPTEINT       As String * 32                    ' INTITULE
    COMPTEAGE       As Integer                        ' AGENCE
    COMPTEDEV       As String * 3                     ' TABLES BASE 013
    COMPTEOUV       As Long                           ' DATE OUVERTURE
    COMPTECLO       As Long                           ' DATE CLOTURE
    COMPTELOR       As String * 1                     ' Lori/Nostri/AUTRE
    COMPTESUC       As String * 1                     ' O/N
    COMPTECLA       As Long                           ' CLASSE SECURITE
    COMPTEFON       As String * 1                     ' TABLES BASE 015
    COMPTEBLO       As Long                           ' DATE LIMITE BLOCAGE
    COMPTEMOT       As String * 32                    ' MOTIF BLOCAGE
    COMPTESEN       As String * 1                     ' CODE SENS SOLDE D/C
    COMPTEMOD       As Long                           ' DATE MODIFICATION
    
    SOLDEDMO        As Long                           ' DATE DERNIER MVT
    SOLDECEN        As Currency                         ' SOLDE ENCOURS
    SOLDEC01        As Currency                         ' SOLDE M

    PLANCOPRO       As String * 3                     ' TABLES BASE 014
    PLANTIERS       As String * 1                     ' COMPTE TIERS O/N


End Type
    
    
Public arrYCOMPTE0() As typeYCOMPTE0
Public arrYCOMPTE0_NB As Integer
Public arrYCOMPTE0_NBMax As Integer
Public arrYCOMPTE0_Index As Integer
Public arrYCOMPTE0_Suite As Boolean

Dim meMVTP0 As typeMvtP0

Public Function srvYCOMPTE0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCOMPTE0 As typeYCOMPTE0)
On Error GoTo Error_Handler
srvYCOMPTE0_GetBuffer_ODBC = Null
recYCOMPTE0.COMPTEETA = rsADO("COMPTEETA")
recYCOMPTE0.COMPTEPLA = rsADO("COMPTEPLA")
recYCOMPTE0.COMPTECOM = rsADO("COMPTECOM")
recYCOMPTE0.COMPTEOBL = rsADO("COMPTEOBL")
recYCOMPTE0.COMPTEINT = rsADO("COMPTEINT")
recYCOMPTE0.COMPTEAGE = rsADO("COMPTEAGE")
recYCOMPTE0.COMPTEDEV = rsADO("COMPTEDEV")
recYCOMPTE0.COMPTEOUV = rsADO("COMPTEOUV")
recYCOMPTE0.COMPTECLO = rsADO("COMPTECLO")
recYCOMPTE0.COMPTELOR = rsADO("COMPTELOR")
recYCOMPTE0.COMPTESUC = rsADO("COMPTESUC")
recYCOMPTE0.COMPTECLA = rsADO("COMPTECLA")
recYCOMPTE0.COMPTEFON = rsADO("COMPTEFON")
recYCOMPTE0.COMPTEBLO = rsADO("COMPTEBLO")
recYCOMPTE0.COMPTEMOT = rsADO("COMPTEMOT")
recYCOMPTE0.COMPTESEN = rsADO("COMPTESEN")
recYCOMPTE0.COMPTEMOD = rsADO("COMPTEMOD")
Exit Function
Error_Handler:
srvYCOMPTE0_GetBuffer_ODBC = Error
End Function
Public Function srvYCOMPTE0_Import(lX As String)
Dim x As String, Nb As Long
Dim wSolde As String, wPlan As String
Dim X200 As String * 200


On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = constYCOMPTE0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    lX = meYbase.Text
    If mId$(lX, 1, 8) >= YBIATAB0_DATE_CPT_J Then
        srvYCOMPTE0_Import = Null
        paramYCOMPTE0_Nb = CLng(mId$(lX, 26, 9))
        Exit Function
    Else
        meYbase.Method = constDelete
        Call tableYBase_Update(meYbase)
    End If
End If


xYBase_YSOLDE0.Method = "Seek="
xYBase_YSOLDE0.ID = constYSOLDE0

xYBase_YPLAN0 = xYBase_YSOLDE0
xYBase_YPLAN0.Method = "Seek="
xYBase_YPLAN0.ID = constYPLAN0

srvYCOMPTE0_Import = "?"

paramYCOMPTE0_Import = paramYBase_DataF & Trim(constYCOMPTE0) & paramYBase_Data_ExtensionP

Open Trim(paramYCOMPTE0_Import) For Input As #1

Nb = 0
x = "delete * from YBase where Id = " & Chr$(34) & Trim(constYCOMPTE0) & Chr$(34)
MDB.Execute x

meYbase.Method = constAddNew
X200 = ""


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, X200
    If Trim(X200) <> "" Then
            meYbase.ID = constYCOMPTE0
            meYbase.K1 = mId$(X200, 10, 20)
            
            xYBase_YSOLDE0.K1 = meYbase.K1
            If tableYBase_Read(xYBase_YSOLDE0) = 0 Then
                wSolde = mId$(xYBase_YSOLDE0.Text, 30, 8) & mId$(xYBase_YSOLDE0.Text, 46, 19) & mId$(xYBase_YSOLDE0.Text, 46, 19)
            Else
                wSolde = String(46, "0")
           End If
            If mId$(X200, 30, 10) <> mId$(xYBase_YPLAN0.K1, 1, 10) Then                      'recYCOMPTE0.COMPTEOBL
                 xYBase_YPLAN0.K1 = mId$(X200, 30, 10)
                 If tableYBase_Read(xYBase_YPLAN0) = 0 Then
                     wPlan = mId$(xYBase_YPLAN0.Text, 52, 3) & mId$(xYBase_YPLAN0.Text, 61, 1)
                 Else
                     wPlan = "???N"
                End If
            End If
            
            meYbase.ID = constYCOMPTE0
            meYbase.K1 = mId$(X200, 10, 20)
            Mid$(X200, 151, 46) = wSolde
            Mid$(X200, 197, 4) = wPlan
            meYbase.Text = X200
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYCOMPTE0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = constYCOMPTE0
paramYCOMPTE0_Nb = Nb
meYbase.Text = YBIATAB0_DATE_CPT_J & "_" & DSys & "_" & time_Hms & "_" & Format$(Nb, "000000000")
lX = meYbase.Text
dbYBase_Update meYbase
paramYCOMPTE0_Nb = Nb

Exit Function

Error_Handle:
 MsgBox "erreur : srvYCOMPTE0_Import" & X200, vbCritical, Error
Close

srvYCOMPTE0_Import = Error
End Function


Public Function srvYCOMPTE0_Import_Array(lnb As Long, marrYCOMPTE0() As typeYCOMPTE0)
Dim xIn As String, x As String
Dim intReturn As Integer
On Error GoTo Error_Handle

srvYCOMPTE0_Import_Array = "?"
lnb = 0
recYCOMPTE0_Init marrYCOMPTE0(0)

meYbase.ID = constYCOMPTE0
meYbase.K1 = ""
meYbase.Method = "Seek>"
intReturn = tableYBase_Read(meYbase)
'meYBase.Method = "MoveNext"
Do
    If Trim(meYbase.ID) <> constYCOMPTE0 Then intReturn = -1
    If intReturn = 0 Then
        lnb = lnb + 1
        MsgTxt = Space$(34) & meYbase.Text
        MsgTxtIndex = 0
        srvYCOMPTE0_GetBuffer marrYCOMPTE0(lnb)
    End If
    intReturn = tableYBase_Read(meYbase)
   '  If blnJPL And lnb > 500 Then intReturn = -1
Loop Until intReturn <> 0
srvYCOMPTE0_Import_Array = Null
Exit Function

Error_Handle:
    MsgBox "erreur : srvYCOMPTE0_Import" & xIn, vbCritical, Error
    srvYCOMPTE0_Import_Array = Error
End Function

Public Function srvYCOMPTE0_Import_Read(lId As String, lYCOMPTE0 As typeYCOMPTE0)

Dim xIn As String, x As String

On Error GoTo Error_Handle

srvYCOMPTE0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYCOMPTE0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYCOMPTE0_GetBuffer lYCOMPTE0
    srvYCOMPTE0_Import_Read = Null
Else
    recYCOMPTE0_Init lYCOMPTE0
    lYCOMPTE0.COMPTECOM = lId
    lYCOMPTE0.COMPTEINT = lId
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYCOMPTE0_Import_Read" & xIn, vbCritical, Error
srvYCOMPTE0_Import_Read = Error
End Function






'-----------------------------------------------------
Function srvYCOMPTE0_Update(recYCOMPTE0 As typeYCOMPTE0)
'-----------------------------------------------------

srvYCOMPTE0_Update = "?"

MsgTxtLen = 0
Call srvYCOMPTE0_PutBuffer(recYCOMPTE0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYCOMPTE0_GetBuffer(recYCOMPTE0)) Then
        Call srvYCOMPTE0_Error(recYCOMPTE0)
        srvYCOMPTE0_Update = recYCOMPTE0.Err
        Exit Function
    Else
        srvYCOMPTE0_Update = Null
    End If
Else
    recYCOMPTE0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYCOMPTE0_Error(recYCOMPTE0 As typeYCOMPTE0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YCOMPTE0" & Chr$(10) & Chr$(13)

Select Case mId$(recYCOMPTE0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYCOMPTE0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YCOMPTE0s.bas  ( " & Trim(recYCOMPTE0.Obj) & " : " & Trim(recYCOMPTE0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYCOMPTE0_Monitor(recYCOMPTE0 As typeYCOMPTE0)
'-----------------------------------------------------

arrYCOMPTE0_Suite = False
Select Case mId$(Trim(recYCOMPTE0.Method), 1, 4)
    Case "Snap"
              srvYCOMPTE0_Monitor = srvYCOMPTE0_Snap(recYCOMPTE0)
    Case Else
            srvYCOMPTE0_Monitor = srvYCOMPTE0_Seek(recYCOMPTE0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYCOMPTE0_GetBuffer(recYCOMPTE0 As typeYCOMPTE0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYCOMPTE0_GetBuffer = Null
recYCOMPTE0.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYCOMPTE0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYCOMPTE0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYCOMPTE0.Err = Space$(10) Then
    recYCOMPTE0.COMPTEETA = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCOMPTE0.COMPTEPLA = CLng(Val(mId$(MsgTxt, K + 6, 4)))
    recYCOMPTE0.COMPTECOM = mId$(MsgTxt, K + 10, 20)
    recYCOMPTE0.COMPTEOBL = mId$(MsgTxt, K + 30, 10)
    recYCOMPTE0.COMPTEINT = mId$(MsgTxt, K + 40, 32)
    recYCOMPTE0.COMPTEAGE = CInt(Val(mId$(MsgTxt, K + 72, 5)))
    recYCOMPTE0.COMPTEDEV = mId$(MsgTxt, K + 77, 3)
    recYCOMPTE0.COMPTEOUV = CLng(Val(mId$(MsgTxt, K + 80, 8)))
    recYCOMPTE0.COMPTECLO = CLng(Val(mId$(MsgTxt, K + 88, 8)))
    recYCOMPTE0.COMPTELOR = mId$(MsgTxt, K + 96, 1)
    recYCOMPTE0.COMPTESUC = mId$(MsgTxt, K + 97, 1)
    recYCOMPTE0.COMPTECLA = CLng(Val(mId$(MsgTxt, K + 98, 3)))
    recYCOMPTE0.COMPTEFON = mId$(MsgTxt, K + 101, 1)
    recYCOMPTE0.COMPTEBLO = CLng(Val(mId$(MsgTxt, K + 102, 8)))
    recYCOMPTE0.COMPTEMOT = mId$(MsgTxt, K + 110, 32)
    recYCOMPTE0.COMPTESEN = mId$(MsgTxt, K + 142, 1)
    recYCOMPTE0.COMPTEMOD = CLng(Val(mId$(MsgTxt, K + 143, 8)))
    
    
    recYCOMPTE0.SOLDEDMO = CLng(Val(mId$(MsgTxt, K + 151, 8)))
    recYCOMPTE0.SOLDECEN = CCur(mId$(MsgTxt, K + 159, 19)) / 1000
    recYCOMPTE0.SOLDEC01 = CCur(mId$(MsgTxt, K + 178, 19)) / 1000
    recYCOMPTE0.PLANCOPRO = mId$(MsgTxt, K + 197, 3)
    recYCOMPTE0.PLANTIERS = mId$(MsgTxt, K + 200, 1)

Else
    srvYCOMPTE0_GetBuffer = recYCOMPTE0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYCOMPTE0Len

End Function

'---------------------------------------------------------
Public Sub srvYCOMPTE0_PutBuffer(recYCOMPTE0 As typeYCOMPTE0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCOMPTE0.Obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCOMPTE0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYCOMPTE0.COMPTEETA, "0000 ")
    Mid$(MsgTxt, K + 6, 4) = Format$(recYCOMPTE0.COMPTEPLA, "000 ")
    Mid$(MsgTxt, K + 10, 20) = recYCOMPTE0.COMPTECOM
    Mid$(MsgTxt, K + 30, 10) = recYCOMPTE0.COMPTEOBL
    Mid$(MsgTxt, K + 40, 32) = recYCOMPTE0.COMPTEINT
    Mid$(MsgTxt, K + 72, 5) = Format$(recYCOMPTE0.COMPTEAGE, "0000 ")
    Mid$(MsgTxt, K + 77, 3) = recYCOMPTE0.COMPTEDEV
    Mid$(MsgTxt, K + 80, 8) = Format$(recYCOMPTE0.COMPTEOUV, "0000000 ")
    Mid$(MsgTxt, K + 88, 8) = Format$(recYCOMPTE0.COMPTECLO, "0000000 ")
    Mid$(MsgTxt, K + 96, 1) = recYCOMPTE0.COMPTELOR
    Mid$(MsgTxt, K + 97, 1) = recYCOMPTE0.COMPTESUC
    Mid$(MsgTxt, K + 98, 3) = Format$(recYCOMPTE0.COMPTECLA, "00 ")
    Mid$(MsgTxt, K + 101, 1) = recYCOMPTE0.COMPTEFON
    Mid$(MsgTxt, K + 102, 8) = Format$(recYCOMPTE0.COMPTEBLO, "0000000 ")
    Mid$(MsgTxt, K + 110, 32) = recYCOMPTE0.COMPTEMOT
    Mid$(MsgTxt, K + 142, 1) = recYCOMPTE0.COMPTESEN
    Mid$(MsgTxt, K + 143, 8) = Format$(recYCOMPTE0.COMPTEMOD, "0000000 ")

MsgTxtLen = MsgTxtLen + recYCOMPTE0Len
End Sub



Public Sub srvYCOMPTE0_ElpDisplay(recYCOMPTE0 As typeYCOMPTE0)
frmElpDisplay.fgData.Rows = 18
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTEETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEPLA    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PLAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTEPLA
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTECOM   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTECOM
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEOBL   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPTE OBLIGATOIRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTEOBL
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEINT   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTITULE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTEINT
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTEAGE
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEDEV    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TABLES BASE 013"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTEDEV
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEOUV    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE OUVERTURE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTEOUV
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTECLO    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE CLOTURE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTECLO
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTELOR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Lori/Nostri/AUTRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTELOR
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTESUC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTESUC
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTECLA    2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CLASSE SECURITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTECLA
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEFON    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TABLES BASE 015"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTEFON
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEBLO    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE LIMITE BLOCAGE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTEBLO
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEMOT   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MOTIF BLOCAGE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTEMOT
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTESEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE SENS SOLDE D/C"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTESEN
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEMOD    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE MODIFICATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMPTE0.COMPTEMOD
frmElpDisplay.Show vbModal
End Sub
Public Sub srvYCOMPTE0_Export_CSV(lIdFile_Source As Integer, lIdFile_Destination As Integer, loptSelect_CSV_Header As Boolean, lnb As Long)
Dim xIn As String
If loptSelect_CSV_Header Then
    Print #2, "COMPTEETA;COMPTEPLA;COMPTECOM;COMPTEOBL;COMPTEINT;COMPTEAGE;COMPTEDEV;COMPTEOUV;COMPTECLO;COMPTELOR;COMPTESUC;COMPTECLA;COMPTEFON;COMPTEBLO;COMPTEMOT;COMPTESEN;COMPTEMOD;"
    Print #2, "ETABLISSEMENT;NUMERO PLAN;NUMERO COMPTE;COMPTE OBLIGATOIRE;INTITULE;AGENCE;DEVISE;DATE OUVERTURE;DATE CLOTURE;Lori/Nostri/AUTRE;COMPTE SUCCESSION;CLASSE SECURITE;CODE FONCTIONNEMENT;DATE LIMITE BLOCAGE;MOTIF BLOCAGE;CODE SENS SOLDE D/C;DATE MODIFICATION;"
    Print #2, ";;;;;;;;;;;;;;;;;"
End If
Do Until EOF(lIdFile_Source)
      Line Input #lIdFile_Source, xIn
      lnb = lnb + 1
      Print #lIdFile_Destination, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 4) & ";" _
      & mId$(xIn, 10, 20) & ";" _
      & mId$(xIn, 30, 10) & ";" _
      & mId$(xIn, 40, 32) & ";" _
      & mId$(xIn, 72, 5) & ";" _
      & mId$(xIn, 77, 3) & ";" _
      & mId$(xIn, 80, 8) & ";" _
      & mId$(xIn, 88, 8) & ";" _
      & mId$(xIn, 96, 1) & ";" _
      & mId$(xIn, 97, 1) & ";" _
      & mId$(xIn, 98, 3) & ";" _
      & mId$(xIn, 101, 1) & ";" _
      & mId$(xIn, 102, 8) & ";" _
      & mId$(xIn, 110, 32) & ";" _
      & mId$(xIn, 142, 1) & ";" _
      & mId$(xIn, 143, 8) & ";"
Loop
End Sub


'---------------------------------------------------------
Private Function srvYCOMPTE0_Seek(recYCOMPTE0 As typeYCOMPTE0)
'---------------------------------------------------------

srvYCOMPTE0_Seek = "?"
MsgTxtLen = 0
Call srvYCOMPTE0_PutBuffer(recYCOMPTE0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYCOMPTE0_GetBuffer(recYCOMPTE0)) Then
        srvYCOMPTE0_Seek = Null
    Else
        Call srvYCOMPTE0_Error(recYCOMPTE0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYCOMPTE0_Snap(recYCOMPTE0 As typeYCOMPTE0)
'---------------------------------------------------------
srvYCOMPTE0_Snap = "?"
MsgTxtLen = 0
Call srvYCOMPTE0_PutBuffer(recYCOMPTE0)
Call srvYCOMPTE0_PutBuffer(arrYCOMPTE0(0))
If IsNull(SndRcv()) Then
    srvYCOMPTE0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYCOMPTE0_GetBuffer(recYCOMPTE0)) Then
            Call arrYCOMPTE0_AddItem(recYCOMPTE0)
            arrYCOMPTE0_Suite = True
        Else
            arrYCOMPTE0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYCOMPTE0_AddItem(recYCOMPTE0 As typeYCOMPTE0)
'---------------------------------------------------------
          
arrYCOMPTE0_NB = arrYCOMPTE0_NB + 1
    
If arrYCOMPTE0_NB > arrYCOMPTE0_NBMax Then
    arrYCOMPTE0_NBMax = arrYCOMPTE0_NBMax + recYCOMPTE0_Block
    ReDim Preserve arrYCOMPTE0(arrYCOMPTE0_NBMax)
End If
            
arrYCOMPTE0(arrYCOMPTE0_NB) = recYCOMPTE0
End Sub



'---------------------------------------------------------
Public Sub recYCOMPTE0_Init(recYCOMPTE0 As typeYCOMPTE0)
'---------------------------------------------------------
recYCOMPTE0.Obj = "ZCOMPTE0_S"
recYCOMPTE0.Method = ""
recYCOMPTE0.Err = ""
recYCOMPTE0.COMPTEETA = 1

recYCOMPTE0.COMPTEETA = 0    ' As Integer                        ' ETABLISSEMENT
recYCOMPTE0.COMPTEPLA = 0    ' As Long                           ' NUMERO PLAN
recYCOMPTE0.COMPTECOM = ""    ' As String * 20                    ' NUMERO COMPTE
recYCOMPTE0.COMPTEOBL = ""    ' As String * 10                    ' COMPTE OBLIGATOIRE
recYCOMPTE0.COMPTEINT = ""    ' As String * 32                    ' INTITULE
recYCOMPTE0.COMPTEAGE = 0    ' As Integer                        ' AGENCE
recYCOMPTE0.COMPTEDEV = ""    ' As String * 3                     ' TABLES BASE 013
recYCOMPTE0.COMPTEOUV = 0    ' As Long                           ' DATE OUVERTURE
recYCOMPTE0.COMPTECLO = 0    ' As Long                           ' DATE CLOTURE
recYCOMPTE0.COMPTELOR = ""    ' As String * 1                     ' Lori/Nostri/AUTRE
recYCOMPTE0.COMPTESUC = ""    ' As String * 1                     ' O/N
recYCOMPTE0.COMPTECLA = 0    ' As Long                           ' CLASSE SECURITE
recYCOMPTE0.COMPTEFON = ""    ' As String * 1                     ' TABLES BASE 015
recYCOMPTE0.COMPTEBLO = 0    ' As Long                           ' DATE LIMITE BLOCAGE
recYCOMPTE0.COMPTEMOT = ""    ' As String * 32                    ' MOTIF BLOCAGE
recYCOMPTE0.COMPTESEN = ""    ' As String * 1                     ' CODE SENS SOLDE D/C
recYCOMPTE0.COMPTEMOD = 0    ' As Long                           ' DATE MODIFICATION

recYCOMPTE0.SOLDEDMO = 0
recYCOMPTE0.SOLDECEN = 0
recYCOMPTE0.SOLDEC01 = 0
recYCOMPTE0.PLANCOPRO = ""
recYCOMPTE0.PLANTIERS = "N"

End Sub






