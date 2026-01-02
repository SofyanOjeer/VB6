Attribute VB_Name = "srvYBIATAB0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYBIATAB0Len = 198 ' 34 +164
Public Const recYBIATAB0_Block = 50
Public Const memoYBIATAB0Len = 164
Public Const constYBIATAB0 = "YBIATAB0"
Dim paramYBIATAB0_Import As String
Dim meYbase As typeYBase
Dim xYBIATAB0 As typeYBIATAB0

Type typeYBIATAB0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    BIATABID        As String * 12
    BIATABK1        As String * 12
    BIATABK2        As String * 12
    BIATABTEXT      As String * 128
End Type

Type typeCV
    DeviseIso     As String * 3
    DeviseN       As String * 3
    DeviseLibellé As String * 20
    Cours         As Double
    CoursAmj      As String * 8
    maxD          As String * 1
    EuroIn        As Boolean
    CotationCertain As Boolean
    
    Montant       As Currency
    OpéAmj        As String * 8
    CoursAmjMin   As String * 8
    AchatVente    As String * 1
    Normal        As String * 1
    CoursCompta   As String * 1

End Type

Type typeYBIAUSR0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    MNURUTUTI       As String * 10                    ' UTILISATEUR
    MNURUTNOM       As String * 30                    ' NOM
    MNURUTETB       As Integer                        ' ETAB. PAR DEFAUT
    MNURUTCUT       As Integer                        ' CODE INTERNE
    MNURUTLOG       As String * 1                     ' ENTREE LOGICIEL
      
    MNUUTICGR       As Integer                        ' CODE GROUPE
    MNUUTIDRG       As String * 1                     ' DROITS GROUPE
    MNUUTIOUT       As String * 10                    ' FILE ATTENTE
    MNUUTILAN       As String * 1                     ' LANGUE
    MNUUTIMSE       As String * 1                     ' MENU SERVICE
    MNUUTIAGE       As Integer                        ' AGENCE DEFAUT
    MNUUTISER       As String * 2                     ' SERVICE DEFAUT
    MNUUTISRV       As String * 2                     ' SOUS-SERV. DEFAUT
      
    MNUUTPAGE       As Integer                        ' Agence
    MNUUTPOIA       As String * 1                     ' Inter Agence
    MNUUTPCLA       As String * 99                    ' Classe
    
    CLIENASIG       As String * 12                    ' sigle usuel

End Type

Public Function YBIATAB0_Sql_Responsable(lK1 As String, cnADO As ADODB.Connection, rsADO As ADODB.Recordset) As String
Dim xSql As String
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'RESPONSABLE' AND BIATABK1 = '" & lK1 & "'"
Set rsADO = cnADO.Execute(xSql)

If Not rsADO.EOF Then
    
    YBIATAB0_Sql_Responsable = lK1 & " : " & mId$(rsADO("BIATABTXT"), 33, 13)
Else
    YBIATAB0_Sql_Responsable = lK1 & " : " & "????"
End If

End Function

Public Sub CV_Calc(lK2 As String, CV1 As typeCV, CV2 As typeCV)
Dim X36 As String * 36
Dim dblMontant As Double

If CV1.Montant = 0 Then
    CV1.Cours = 0
    CV2.Montant = 0
    Exit Sub
End If

X36 = "FIXING"
Mid$(X36, 13, 3) = CV1.DeviseIso
Mid$(X36, 25, Len(lK2)) = lK2  '"J" 'CV1.CoursCompta

If IsNull(srvYBIATAB0_Import_Read(X36, xYBIATAB0)) Then
    CV1.Cours = CDbl(mId$(xYBIATAB0.BIATABTEXT, 36 + 9, 15)) / 1000000000
    dblMontant = Abs(CV1.Montant) / CV1.Cours
    CV2.Montant = Fix((dblMontant + 0.00500001) * 100) / 100
    If CV1.Montant < 0 Then CV2.Montant = -CV2.Montant
Else
    CV1.Cours = 0
    CV2.Montant = 999999999999.99
    'MsgBox "Manque cours :" & CV1.DeviseIso, vbCritical, "BIA.CV_CALC"
    
End If


End Sub


Public Sub srvYBIATAB0_Import_cboDevise(lCbo As ComboBox)
Dim X3 As String * 3
lCbo.Clear
lCbo.AddItem " "
lCbo.Clear
lCbo.AddItem " "

X3 = " "
meYbase.ID = constYBIATAB0
meYbase.K1 = "DEVISE"
meYbase.Method = "Seek>"
Do
    intReturn = tableYBase_Read(meYbase)
    If Trim(meYbase.ID) <> constYBIATAB0 Then intReturn = -1
    If mId$(meYbase.K1, 1, 6) <> "DEVISE" Then intReturn = -1
    If intReturn = 0 Then
        If X3 <> mId$(meYbase.Text, 25, 3) Then
            X3 = mId$(meYbase.Text, 25, 3)
            lCbo.AddItem X3
        End If
    End If
        
Loop Until intReturn <> 0


End Sub

Public Sub srvYBIATAB0_Import_lstX(lstX As ListBox, lYBIATAB_ID As String, lYBIATAB_K1 As String)
Dim selK1 As String, selLen As Integer

lstX.Clear
selK1 = Space$(24)
Mid$(selK1, 1, 12) = lYBIATAB_ID
Mid$(selK1, 13, 12) = lYBIATAB_K1
selLen = 24
If Trim(lYBIATAB_K1) = "" Then selLen = 12


meYbase.ID = constYBIATAB0
meYbase.K1 = selK1
meYbase.Method = "Seek>"
Do
    intReturn = tableYBase_Read(meYbase)
    If Trim(meYbase.ID) <> constYBIATAB0 Then intReturn = -1
    If mId$(meYbase.K1, 1, selLen) <> mId$(selK1, 1, selLen) Then intReturn = -1
    If intReturn = 0 Then
        lstX.AddItem mId$(meYbase.Text, 36 + 25, 30)
    End If
        
Loop Until intReturn <> 0


End Sub


Public Function srvYBIATAB0_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = constYBIATAB0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    lX = meYbase.Text
    If mId$(lX, 1, 8) >= YBIATAB0_DATE_CPT_J Then
        srvYBIATAB0_Import = Null
        Exit Function
    Else
        meYbase.Method = constDelete
        Call tableYBase_Update(meYbase)
    End If
End If


srvYBIATAB0_Import = "?"

paramYBIATAB0_Import = paramYBase_DataF & Trim(constYBIATAB0) & paramYBase_Data_ExtensionP

Open Trim(paramYBIATAB0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYBIATAB0) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYBIATAB0
            meYbase.K1 = mId$(xIn, 1, 36)
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYBIATAB0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = constYBIATAB0
meYbase.Text = YBIATAB0_DATE_CPT_J & "_" & DSys & "_" & time_Hms & "_" & Format$(Nb, "000000000")
lX = meYbase.Text
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYBIATAB0_Import" & xIn, vbCritical, Error
Close

srvYBIATAB0_Import = Error
End Function


Public Function srvYBIATAB0_Import_Read(lId As String, lYBIATAB0 As typeYBIATAB0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYBIATAB0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYBIATAB0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYBIATAB0_GetBuffer lYBIATAB0
    srvYBIATAB0_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYBIATAB0_Import_Read" & xIn, vbCritical, Error
srvYBIATAB0_Import_Read = Error
End Function

Public Function srvYBIAUSR0_Import_Read(lYBIAUSR0 As typeYBIAUSR0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYBIAUSR0_Import_Read = "?"
recYBIAUSR0_Init lYBIAUSR0

meYbase.Method = "Seek="
meYbase.ID = constYBIATAB0
meYbase.K1 = "USER"
Mid$(meYbase.K1, 25, 12) = lYBIAUSR0.MNURUTUTI

If tableYBase_Read(meYbase) = 0 Then
    lYBIAUSR0.MNURUTNOM = mId(meYbase.Text, 36 + 1, 30) ' as String * 30                    ' NOM
    lYBIAUSR0.MNURUTETB = CInt(mId(meYbase.Text, 36 + 31, 4)) ' as Integer                        ' ETAB. PAR DEFAUT
    lYBIAUSR0.MNURUTCUT = CInt(mId(meYbase.Text, 36 + 35, 4)) ' as Integer                        ' CODE INTERNE
    lYBIAUSR0.MNURUTLOG = mId(meYbase.Text, 36 + 39, 1) ' as String * 1                     ' ENTREE LOGICIEL
          
    lYBIAUSR0.MNUUTICGR = CInt(mId(meYbase.Text, 36 + 40, 4)) ' as Integer                        ' CODE GROUPE
    lYBIAUSR0.MNUUTIDRG = mId(meYbase.Text, 36 + 44, 1) ' as String * 1                     ' DROITS GROUPE
    lYBIAUSR0.MNUUTIOUT = mId(meYbase.Text, 36 + 45, 10) ' as String * 10                    ' FILE ATTENTE
    lYBIAUSR0.MNUUTILAN = mId(meYbase.Text, 36 + 55, 1) ' as String * 1                     ' LANGUE
    lYBIAUSR0.MNUUTIMSE = mId(meYbase.Text, 36 + 56, 1) ' as String * 1                     ' MENU SERVICE
    lYBIAUSR0.MNUUTIAGE = CInt(mId(meYbase.Text, 36 + 57, 4)) ' as Integer                        ' AGENCE DEFAUT
    lYBIAUSR0.MNUUTISER = mId(meYbase.Text, 36 + 61, 2) ' as String * 2                     ' SERVICE DEFAUT
    lYBIAUSR0.MNUUTISRV = mId(meYbase.Text, 36 + 63, 2) ' as String * 2    ' SOUS-SERV. DEFAUT
    
    Mid$(meYbase.K1, 13, 12) = "CLASSE"
    If tableYBase_Read(meYbase) = 0 Then
          
        lYBIAUSR0.MNUUTPAGE = CInt(mId(meYbase.Text, 36 + 1, 4)) ' as Integer                        ' Agence
        lYBIAUSR0.MNUUTPOIA = mId(meYbase.Text, 36 + 5, 1) ' as String * 1                     ' Inter Agence
        lYBIAUSR0.MNUUTPCLA = mId(meYbase.Text, 36 + 6, 99) ' as String * 99                    ' Classe
        
        recElpTable_Init xElpTable
        xElpTable.Method = "Seek="
        xElpTable.ID = "User"
        xElpTable.K1 = usrId
        xElpTable.K2 = "CLIENASIG"
        If tableElpTable_Read(xElpTable) = 0 Then
            lYBIAUSR0.CLIENASIG = Trim(xElpTable.Memo)
        Else
            lYBIAUSR0.CLIENASIG = usrId
        End If
        
    End If
End If
        
Exit Function

Error_Handle:

Shell_MsgBox "srvYBIAUSR0_Import_Read " & Error, vbInformation, lYBIAUSR0.MNURUTUTI, True

srvYBIAUSR0_Import_Read = Error
End Function

'---------------------------------------------------------
Public Function srvYBIATAB0_GetBuffer(recYBIATAB0 As typeYBIATAB0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYBIATAB0_GetBuffer = Null
recYBIATAB0.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYBIATAB0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYBIATAB0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYBIATAB0.Err = Space$(10) Then
    recYBIATAB0.BIATABID = mId$(MsgTxt, K + 1, 12)
    recYBIATAB0.BIATABK1 = mId$(MsgTxt, K + 1, 12)
    recYBIATAB0.BIATABK2 = mId$(MsgTxt, K + 1, 12)
    recYBIATAB0.BIATABTEXT = mId$(MsgTxt, K + 1, 128)
Else
    srvYBIATAB0_GetBuffer = recYBIATAB0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYBIATAB0Len

End Function





Public Sub recYBIAUSR0_Init(lYBIAUSR0 As typeYBIAUSR0)
lYBIAUSR0.MNURUTNOM = "" ' as String * 30                    ' NOM
lYBIAUSR0.MNURUTETB = 0 ' as Integer                        ' ETAB. PAR DEFAUT
lYBIAUSR0.MNURUTCUT = 0 ' as Integer                        ' CODE INTERNE
lYBIAUSR0.MNURUTLOG = "" ' as String * 1                     ' ENTREE LOGICIEL

lYBIAUSR0.MNUUTICGR = 0 ' as Integer                        ' CODE GROUPE
lYBIAUSR0.MNUUTIDRG = "" ' as String * 1                     ' DROITS GROUPE
lYBIAUSR0.MNUUTIOUT = "" ' as String * 10                    ' FILE ATTENTE
lYBIAUSR0.MNUUTILAN = "" ' as String * 1                     ' LANGUE
lYBIAUSR0.MNUUTIMSE = "" ' as String * 1                     ' MENU SERVICE
lYBIAUSR0.MNUUTIAGE = 0 ' as Integer                        ' AGENCE DEFAUT
lYBIAUSR0.MNUUTISER = "" ' as String * 2                     ' SERVICE DEFAUT
lYBIAUSR0.MNUUTISRV = "" ' as String * 2    ' SOUS-SERV. DEFAUT


lYBIAUSR0.MNUUTPAGE = 0 ' as Integer                        ' Agence
lYBIAUSR0.MNUUTPOIA = "" ' as String * 1                     ' Inter Agence
lYBIAUSR0.MNUUTPCLA = "" ' as String * 99                    ' Classe

End Sub

Public Function srvYBIATAB0_Pays(lPays As String) As String
meYbase.ID = constYBIATAB0
meYbase.K1 = "SAB         CLIENAPAY   CLI" & lPays
meYbase.Method = "Seek="
intReturn = tableYBase_Read(meYbase)
If intReturn = 0 Then
    srvYBIATAB0_Pays = Trim(mId$(meYbase.Text, 49 + 3, 30))
Else
    srvYBIATAB0_Pays = lPays
End If

End Function
