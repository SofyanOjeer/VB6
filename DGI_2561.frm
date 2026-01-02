VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmDGI_2561 
   Caption         =   "Déclaration DGI 2561"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   9900
   Begin VB.CommandButton cmdImport_IFUTR141P1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Import IFUTR141P1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   0
      Width           =   1080
   End
   Begin VB.CommandButton cmdPrint_All 
      BackColor       =   &H000000FF&
      Caption         =   "PRINT *ALL"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   0
      Width           =   615
   End
   Begin VB.Frame fraId 
      Height          =   615
      Left            =   3240
      TabIndex        =   41
      Top             =   -120
      Width           =   1695
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   600
         TabIndex        =   42
         Top             =   240
         Width           =   735
      End
      Begin ComCtl2.UpDown UpDownDGI_2561 
         Height          =   495
         Left            =   1440
         TabIndex        =   43
         Top             =   120
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   327681
         Value           =   1
         AutoBuddy       =   -1  'True
         OrigLeft        =   3000
         OrigTop         =   3120
         OrigRight       =   3240
         OrigBottom      =   3405
         Max             =   9999
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblRFBENF 
         Caption         =   "Code "
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraDGI_2561 
      Height          =   6855
      Left            =   0
      TabIndex        =   22
      Top             =   600
      Width           =   9855
      Begin VB.CommandButton cmdImportZ 
         BackColor       =   &H00C0C0FF&
         Caption         =   "ImportZ 2002"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   120
         Width           =   1200
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "imprimer la déclaration après enregistrement"
         Height          =   255
         Left            =   1920
         TabIndex        =   45
         Top             =   240
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BENEFICIAIRE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   3135
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   9615
         Begin VB.TextBox txtZD 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1800
            TabIndex        =   1
            Top             =   600
            Width           =   7335
         End
         Begin VB.TextBox txtZG 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1800
            TabIndex        =   2
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtAF 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1800
            TabIndex        =   8
            Top             =   2400
            Width           =   1815
         End
         Begin VB.TextBox txtZJ 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7920
            TabIndex        =   5
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtZI 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1800
            TabIndex        =   4
            Top             =   1320
            Width           =   4935
         End
         Begin VB.TextBox txtAC 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1800
            TabIndex        =   6
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox txtAO 
            Height          =   285
            Left            =   5280
            TabIndex        =   9
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox txtCT 
            Height          =   285
            Left            =   1800
            TabIndex        =   10
            Top             =   2760
            Width           =   4935
         End
         Begin VB.TextBox txtZH 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2880
            TabIndex        =   3
            Top             =   960
            Width           =   6255
         End
         Begin VB.TextBox txtAE 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1800
            TabIndex        =   7
            Top             =   2040
            Width           =   7335
         End
         Begin VB.TextBox txtZC 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1800
            TabIndex        =   0
            Top             =   240
            Width           =   7335
         End
         Begin VB.Label lblCT 
            Caption         =   "Nom Marital"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label lblAF 
            Caption         =   "Département"
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label lblAO 
            Caption         =   "Sexe"
            Height          =   375
            Left            =   4560
            TabIndex        =   38
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label lblAE 
            Caption         =   "Lieu "
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label lblAC 
            Caption         =   "Date de Naissance"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label lblZJ 
            Caption         =   "Code Postal"
            Height          =   375
            Left            =   6960
            TabIndex        =   35
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label lblZI 
            Caption         =   "Commune"
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblZC 
            Caption         =   "Nom Patronymique"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblZD 
            Caption         =   "Prénoms"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblZG 
            Caption         =   "N° ...... Voie"
            Height          =   375
            Left            =   240
            TabIndex        =   31
            Top             =   1080
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "INFORMATIONS GENERALES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   2775
         Left            =   240
         TabIndex        =   23
         Top             =   3960
         Width           =   9615
         Begin VB.TextBox txtBP 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            TabIndex        =   16
            Top             =   2160
            Width           =   1575
         End
         Begin VB.TextBox txtBN 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            TabIndex        =   15
            Top             =   1800
            Width           =   1575
         End
         Begin VB.TextBox txtBR 
            Height          =   285
            Left            =   2160
            TabIndex        =   13
            Top             =   1080
            Width           =   2775
         End
         Begin VB.TextBox txtAH 
            Height          =   285
            Left            =   2160
            TabIndex        =   12
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox txtAR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            TabIndex        =   14
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox txtAI 
            Height          =   285
            Left            =   2160
            TabIndex        =   11
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label lblBP 
            Caption         =   "Montant du Prélèvement"
            Height          =   375
            Left            =   240
            TabIndex        =   29
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label lblBN 
            Caption         =   "Base de Prélèvement"
            Height          =   375
            Left            =   240
            TabIndex        =   28
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label lblAR 
            Caption         =   "Produits ou Gains"
            Height          =   375
            Left            =   240
            TabIndex        =   27
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "Type du Compte"
            Height          =   375
            Left            =   240
            TabIndex        =   26
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label lblAH 
            Caption         =   "Nature du Compte"
            Height          =   375
            Left            =   240
            TabIndex        =   25
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label lblAI 
            Caption         =   "Référence du Compte"
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   480
            Width           =   2055
         End
      End
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H80000004&
      Caption         =   "&Abandonner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   -120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   -120
      Width           =   1305
   End
   Begin VB.CommandButton cmdSuppress 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Supprimer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   0
      Width           =   945
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   6600
      TabIndex        =   19
      Top             =   0
      Width           =   2895
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   9480
      Picture         =   "DGI_2561.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Enregistrer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   0
      Width           =   1065
   End
   Begin VB.Menu cmdPrint_mnu 
      Caption         =   "Print"
      Visible         =   0   'False
      Begin VB.Menu cmdPrint_mnuList 
         Caption         =   "Liste téléphonique"
      End
      Begin VB.Menu cmdPrint_mnuDétail 
         Caption         =   "Liste détaillée"
      End
   End
End
Attribute VB_Name = "frmDGI_2561"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim Msg As String
Dim DGI_2561Aut As typeAuthorization
Dim currentMethod As String, currentAMJ As String

Dim updDGI_2561 As Boolean
Dim blnImport_Init As Boolean

Private Sub cmdImport_DGI_2561()
Dim X As String, xOut As String * 453
On Error GoTo Error_Monitor

Dim I As Integer

MDB.Execute "delete * from DGI_2561"
mdbDGI_2561.tableDGI_2561_Open

X = paramDGI_2561_Filename
''libImport_FileName = X

I = 0
Open X For Input As #1

Do Until EOF(1)
    Line Input #1, X
    I = I + 1
    xOut = X
'    libImport_Dta = xOut: DoEvents
    Import_DGI_2561 xOut, recDGI_2561
    If blnImport_Init Then
        recDGI_2561.AR = 0
        recDGI_2561.BN = 0
        recDGI_2561.BP = 0
    End If
    recDGI_2561.Method = constAddNew
    dbDGI_2561_Update recDGI_2561
Loop
'libImport_Dta = "Déclarations lus : " & I
Close #1
mdbDGI_2561.tableDGI_2561_Close
Exit Sub
'---------------------------------------------------------
Error_Monitor:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Erreur Import")

End Sub

Private Sub cmdImport_IFUTR141P1_Load()
Dim X As String, xVoie As String, iVoie As Integer, blnVoie As Boolean
Dim Nb As Integer, iSeq As Integer
Dim blnAddNew As Boolean
On Error GoTo Error_Monitor

Dim I As Integer

MDB.Execute "delete * from DGI_2561"
mdbDGI_2561.tableDGI_2561_Open

X = paramDGI_IFUTR141P1
Nb = 0: iSeq = 0
blnAddNew = False

I = 0
Open X For Input As #1
Line Input #1, X

Do Until EOF(1)
    Line Input #1, X
    I = I + 1
    
    If mId$(X, 1, 4) = "011 " Then
        If blnAddNew Then dbDGI_2561_Update recDGI_2561
        recDGI_25611_Init recDGI_2561
        recDGI_2561.Method = constAddNew
        Nb = Nb + 1
        recDGI_2561.Id = Format$(Nb, "00000")
        iSeq = 0
        blnAddNew = True
    End If
    
    iSeq = iSeq + Val(mId$(X, 4, 1))
    ''If Nb = 6 Then Debug.Print iSeq, mId$(X, 38, 100)
    Select Case iSeq
        Case 5: recDGI_2561.AI = Trim(mId$(X, 110, 10))
        Case 7: recDGI_2561.AH = Trim(mId$(X, 110, 10))
        Case 9: recDGI_2561.BR = Trim(mId$(X, 110, 10))
        Case 11: recDGI_2561.ZC = Trim(mId$(X, 38, 40))
                recDGI_2561.AC = Trim(mId$(X, 110, 20))
        Case 12: recDGI_2561.ZD = Trim(mId$(X, 38, 40))
        Case 14: xVoie = Trim(mId$(X, 38, 40))
                blnVoie = False
                For iVoie = 1 To Len(xVoie)
                    If IsNumeric(mId$(xVoie, iVoie, 1)) Then
                        blnVoie = True
                    Else
                        If mId$(xVoie, iVoie, 1) = "," Then Mid$(xVoie, iVoie, 1) = " "
                        Exit For
                    End If
                Next iVoie
                If blnVoie Then
                    recDGI_2561.ZG = mId$(xVoie, 1, iVoie - 1)
                    
                    recDGI_2561.ZH = Trim(mId$(xVoie, iVoie, Len(xVoie) - iVoie + 1))
                Else
                     recDGI_2561.ZH = xVoie
               End If
                
                 recDGI_2561.AE = Trim(mId$(X, 110, 40))
        Case 15: recDGI_2561.AF = Trim(mId$(X, 110, 40))
        Case 16: recDGI_2561.AO = IIf(Trim(mId$(X, 113, 1)) = "X", "1", "2")
        Case 17: recDGI_2561.CT = Trim(mId$(X, 110, 40))
        Case 18: recDGI_2561.ZI = Trim(mId$(X, 38, 40))
        Case 19: recDGI_2561.ZJ = Trim(mId$(X, 38, 40))
        Case 26: recDGI_2561.AR = CCur(Val(Trim(mId$(X, 38, 20))))
        Case 41: recDGI_2561.BN = CCur(Val(Trim(mId$(X, 110, 20))))
        Case 42: recDGI_2561.BP = CCur(Val(Trim(mId$(X, 110, 20))))
End Select

Loop

If blnAddNew Then dbDGI_2561_Update recDGI_2561

'libImport_Dta = "Déclarations lus : " & I
Close #1
mdbDGI_2561.tableDGI_2561_Close
Exit Sub
'---------------------------------------------------------
Error_Monitor:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Erreur Import")

End Sub

Public Function param_Init()
Dim V

MsgBox "Fichiers travail => C:\Temp\IFU\ ... à archiver dans Serveur0\_Compta\DGI_2561\2002\", vbInformation, "DGI_2561"


blnImport_Init = False:  cmdImportZ.Visible = False
' initialisation 2000 => 2001 => 2002  blnImport_Init = True:  cmdImportZ.Visible = True

param_Init = Null
recElpTable_Init recElpTable
recElpTable.Id = "Param"
recElpTable.K1 = "Compta"
recElpTable.Method = "Seek="

recElpTable.K2 = "DGI_2561"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramDGI_2561_Filename = paramServer(Trim(recElpTable.Memo))
paramDGI_2561_Filename = paramTemp_Folder & "\IFU\DGI_2561.txt"
Call lstErr_Clear(lstErr, cmdContext, "Fichier :" & paramDGI_2561_Filename)
paramDGI_IFUTR141P1 = paramTemp_Folder & "\IFU\IFUTR141P1.txt"
Exit Function

Table_Error:
param_Init = V
Exit Function

Memo_Error:
param_Init = "Memo"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "DGI_2561"
Exit Function

End Function

'---------------------------------------------------------
Function cmdImport_DGI_2002()
'---------------------------------------------------------
Dim X As String, KText As Integer, xReturn As String, iReturn As Integer
On Error GoTo Error_Monitor

Dim I As Integer
MsgBox "interdit", vbCritical
Exit Function

X = paramTemp_Folder & "\Ftp\IFU_2002.csv"

Open X For Input As #1

Do Until EOF(1)
    Line Input #1, X
    KText = 0
    recDGI_2561.Id = CSV_Scan(X, KText)
    recDGI_2561.Method = "Seek="
    iReturn = tableDGI_2561_Read(recDGI_2561)
    If iReturn <> 0 Then
''        Call MsgBox(X, vbCritical, "cmdImport_DGI_20013")
        recDGI_2561.ZC = ""
         recDGI_2561.ZD = ""
         recDGI_2561.ZG = ""
         recDGI_2561.ZH = ""
         recDGI_2561.ZI = ""
         recDGI_2561.ZJ = ""
         recDGI_2561.AI = ""
         recDGI_2561.AH = ""
         recDGI_2561.BR = ""
         recDGI_2561.AC = ""
         recDGI_2561.AE = ""
         recDGI_2561.AF = ""
         recDGI_2561.AO = ""
         recDGI_2561.CT = ""
        recDGI_2561.AR = 0
        recDGI_2561.BN = 0
        recDGI_2561.BP = 0
        recDGI_2561.Method = constAddNew

    Else
        recDGI_2561.Method = constUpdate
    End If

        recDGI_2561.AH = CSV_Scan(X, KText)
        xReturn = CSV_Scan(X, KText)
        xReturn = CSV_Scan(X, KText)
        xReturn = CSV_Scan(X, KText)
        xReturn = CSV_Scan(X, KText)
        recDGI_2561.AR = CCur(Val(CSV_Scan(X, KText)))
        recDGI_2561.BN = CCur(Val(CSV_Scan(X, KText)))
        recDGI_2561.BP = CCur(Val(CSV_Scan(X, KText)))
        
   ' recDGI_2561.AH = mId$(MsgTxt, 266, 20)
   ' recDGI_2561.AR = CCur(Val(mId$(MsgTxt, 406, 16)) / 100)
   ' recDGI_2561.BN = CCur(Val(mId$(MsgTxt, 422, 16)) / 100)
   ' recDGI_2561.BP = CCur(Val(mId$(MsgTxt, 438, 16)) / 100)

        dbDGI_2561_Update recDGI_2561
Loop
'libImport_Dta = "Déclarations lus : " & I
Close #1
Exit Function
'---------------------------------------------------------
Error_Monitor:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Erreur Import")


End Function




Private Sub cmdExport_DGI_2561()
Dim X As String, xOut As String * 453, intReturn As Integer
Dim X2 As String

On Error Resume Next

Dim I As Integer
fraDGI_2561.Enabled = False

mdbDGI_2561.tableDGI_2561_Open

X = paramDGI_2561_Filename
'libImport_FileName = X
X2 = X & "_Copie"
If Dir(X2, vbReadOnly + vbHidden) <> "" Then Kill X2

Name X As X2

I = 0
Open X For Output As #1

recDGI_2561.Method = "MoveFirst"
xOut = Space(453)
Do
    intReturn = tableDGI_2561_Read(recDGI_2561)
    If intReturn = 0 Then
        recDGI_2561.Method = "MoveNext"
        Export_DGI_2561 xOut, recDGI_2561
        Print #1, Trim(xOut)
        I = I + 1
    End If
Loop While intReturn = 0

Call lstErr_Clear(lstErr, cmdPrint, "Déclarations exportées : " & I)
Close #1
mdbDGI_2561.tableDGI_2561_Close
'fraLrBia.Enabled = True

End Sub

Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: cmdRechercher
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select
End Sub

Private Sub Rec_Update()

'--------------------------------------------------------------
Dim valX, X As String
recDGI_2561.Id = Trim(txtID)
recDGI_2561.ZC = Trim(txtZC)
recDGI_2561.ZD = Trim(txtZD)
recDGI_2561.ZG = Trim(txtZG)
recDGI_2561.ZH = Trim(txtZH)
recDGI_2561.ZI = Trim(txtZI)
recDGI_2561.ZJ = Trim(txtZJ)
recDGI_2561.AI = Trim(txtAI)
recDGI_2561.AH = Trim(txtAH)
recDGI_2561.BR = Trim(txtBR)
recDGI_2561.AC = Trim(txtAC)
recDGI_2561.AE = Trim(txtAE)
recDGI_2561.AF = Trim(txtAF)
recDGI_2561.AO = Trim(txtAO)
recDGI_2561.CT = Trim(txtCT)
X = num_Control(txtAR, valX, 13, 0): recDGI_2561.AR = CCur(valX)
X = num_Control(txtBN, valX, 13, 0): recDGI_2561.BN = CCur(valX)
X = num_Control(txtBP, valX, 13, 0): recDGI_2561.BP = CCur(valX)


End Sub

Private Sub cmdImport_IFUTR141P1_Click()
Dim I As Integer
''MsgBox "Interdit : annule et remplace le fichier existant", vbCritical
''Exit Function

blnImport_Init = True
tableDGI_2561_Close
cmdImport_IFUTR141P1_Load
tableDGI_2561_Open

blnImport_Init = False

End Sub

Private Sub cmdImportZ_Click()
blnImport_Init = True
tableDGI_2561_Close
cmdImport_DGI_2561

tableDGI_2561_Open
cmdImport_DGI_2002

blnImport_Init = False

End Sub

Private Sub cmdPrint_All_Click()
Dim X As String, xOut As String * 453, intReturn As Integer
Dim X2 As String

On Error Resume Next

Dim I As Integer
fraDGI_2561.Enabled = False

mdbDGI_2561.tableDGI_2561_Open


I = 0

recDGI_2561.Method = "MoveFirst"
xOut = Space(453)
Do
    intReturn = tableDGI_2561_Read(recDGI_2561)
    If intReturn = 0 Then
        If recDGI_2561.AR <> 0 Or recDGI_2561.BN <> 0 Or recDGI_2561.BP <> 0 Then
            prtrecDGI_2561 = recDGI_2561
            prtDGI_2561_Monitor " "
        End If
       recDGI_2561.Method = "MoveNext"

        I = I + 1
    End If
Loop While intReturn = 0

Call lstErr_Clear(lstErr, cmdPrint, "= cmdPrint_All : " & I)

End Sub

Private Sub cmdPrint_Click()
'Me.PopupMenu cmdPrint_mnu
Rec_Update
cmdPrint_Call
End Sub
Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim X As String

X = Trim(mId$(Msg, 13, Len(Msg)))
txtID.Visible = True 'DGI_2561Aut.Valider

    tableDGI_2561_Open
    updDGI_2561 = False
recDGI_2561.Method = "MoveFirst"
'cmddbDGI_2561
cmdClear

End Sub

Public Sub cmdContext_Quit()
If fraId.Enabled Then
    Unload Me
Else
    cmdClear
End If
End Sub

Public Sub cmdContext_Return()
If fraDGI_2561.Enabled Then
    If ActiveControl.Name = lastActiveControl_Name Then
 '       cmdOk_Click
    Else
        SendKeys "{TAB}"
    End If
Else
If cmdContext.Caption = constcmdRechercher Then cmdRechercher
End If

End Sub


Public Sub Rec_Display()
'----------------------------------------------------
txtID = Trim(recDGI_2561.Id)
txtZC = Trim(recDGI_2561.ZC)
txtZD = Trim(recDGI_2561.ZD)
txtZG = Trim(recDGI_2561.ZG)
txtZH = Trim(recDGI_2561.ZH)
txtZI = Trim(recDGI_2561.ZI)
txtZJ = Trim(recDGI_2561.ZJ)
txtAI = Trim(recDGI_2561.AI)
txtAH = Trim(recDGI_2561.AH)
txtBR = Trim(recDGI_2561.BR)
txtAC = Trim(recDGI_2561.AC)
txtAE = Trim(recDGI_2561.AE)
txtAF = Trim(recDGI_2561.AF)
txtAO = Trim(recDGI_2561.AO)
txtCT = Trim(recDGI_2561.CT)
txtAR = Format$(recDGI_2561.AR, "### ### ###")
txtBN = Format$(recDGI_2561.BN, "### ### ###")
txtBP = Format$(recDGI_2561.BP, "### ### ###")
End Sub

Private Sub cmdPrint_mnuDétail_Click()
Dim Msg As String
'Msg = "000001" & Format$(arrDGI_2561Nb, "000000") & "D"

'prtDGI_2561X Msg

End Sub

Private Sub cmdPrint_mnuList_Click()
Dim Msg As String
'Msg = "000001" & Format$(arrDGI_2561Nb, "000000") & "L"

'prtDGI_2561X Msg

End Sub

Private Sub cmdQuit_Click()
cmdContext_Quit
End Sub

Private Sub cmdSave_Click()
Rec_Update
dbDGI_2561_Update recDGI_2561
cmdClear
If chkPrint = "1" Then cmdPrint_Call
End Sub

Private Sub cmdSuppress_Click()
recDGI_2561.Method = constDelete
dbDGI_2561_Update recDGI_2561
cmdClear

End Sub

Public Sub cmddbDGI_2561()
recDGI_2561.Id = Trim(txtID)
dbDGI_2561_Read recDGI_2561
cmdContext.Caption = constcmdAbandonner
cmdSave.Visible = True
If recDGI_2561.Err = 0 Then
    recDGI_2561.Method = constUpdate
    cmdSuppress.Visible = True
    cmdPrint.Visible = True
    Rec_Display
Else
    Call lstErr_Clear(lstErr, cmdContext, "Création ")
    recDGI_2561.Method = constAddNew
End If
fraId.Enabled = False
fraDGI_2561.Enabled = True
txtZC.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub


Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
'frmDGI_2561.img.Picture = LoadPicture("C:\biadoc\dgi_2561.bmp")

param_Init
cmdImport_DGI_2561

tableDGI_2561_Open
End Sub


'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
End Sub
'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
cmdExport_DGI_2561

tableDGI_2561_Close
End Sub


Public Sub cmdClear()
txtID = ""
txtZC = ""
txtZD = ""
txtZG = ""
txtZH = ""
txtZI = ""
txtZJ = ""
txtAI = ""
txtAH = ""
txtBR = ""
txtAC = ""
txtAE = ""
txtAF = ""
txtAO = ""
txtCT = ""
txtAR = ""
txtBN = ""
txtBP = ""
cmdContext.Visible = True 'DGI_2561Aut.Valider
cmdContext.Caption = constcmdRechercher
cmdSave.Visible = False
cmdSuppress.Visible = False
cmdPrint.Visible = False
cmdPrint_mnuDétail.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "préciser le code ")
fraId.Enabled = True ': txtid.SetFocus
fraDGI_2561.Enabled = False
End Sub

Public Sub cmdRechercher()
If Trim(txtID) = "" Then
    Call lstErr_Clear(lstErr, cmdContext, "préciser le code ")
    Exit Sub
End If

recDGI_2561.Method = "Seek="
cmddbDGI_2561
End Sub


Private Sub txtAC_GotFocus()
txt_GotFocus txtAC
End Sub


Private Sub txtAC_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtAC_LostFocus()
txt_LostFocus txtAC
End Sub


Private Sub txtAE_GotFocus()
txt_GotFocus txtAE
End Sub


Private Sub txtAE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtAE_LostFocus()
txt_LostFocus txtAE
End Sub


Private Sub txtAF_GotFocus()
txt_GotFocus txtAF
End Sub


Private Sub txtAF_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtAF_LostFocus()
txt_LostFocus txtAF
End Sub


Private Sub txtAH_GotFocus()
txt_GotFocus txtAH
End Sub


Private Sub txtAH_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtAH_LostFocus()
txt_LostFocus txtAH
End Sub


Private Sub txtAI_GotFocus()
txt_GotFocus txtAI
End Sub


Private Sub txtAI_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtAI_LostFocus()
txt_LostFocus txtAI
End Sub


Private Sub txtAO_GotFocus()
txt_GotFocus txtAO
End Sub


Private Sub txtAO_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtAO_LostFocus()
txt_LostFocus txtAO
End Sub


Private Sub txtAR_GotFocus()
txt_GotFocus txtAR
End Sub


Private Sub txtAR_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtAR_LostFocus()
txt_LostFocus txtAR
End Sub


Private Sub txtBN_GotFocus()
txt_GotFocus txtBN
End Sub


Private Sub txtBN_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtBN_LostFocus()
txt_LostFocus txtBN
End Sub


Private Sub txtBP_GotFocus()
txt_GotFocus txtBP
End Sub


Private Sub txtBP_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtBP_LostFocus()
txt_LostFocus txtBP
End Sub


Private Sub txtBR_GotFocus()
txt_GotFocus txtBR
End Sub


Private Sub txtBR_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtBR_LostFocus()
txt_LostFocus txtBR
End Sub


Private Sub txtCT_GotFocus()
txt_GotFocus txtCT
End Sub


Private Sub txtCT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtCT_LostFocus()
txt_LostFocus txtCT
End Sub


Private Sub txtId_GotFocus()
txt_GotFocus txtID

End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtId_LostFocus()
txt_LostFocus txtID
cmdRechercher

End Sub


Private Sub txtZC_GotFocus()
txt_GotFocus txtZC
End Sub


Private Sub txtZC_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtZC_LostFocus()
txt_LostFocus txtZC
End Sub


Private Sub txtZD_GotFocus()
txt_GotFocus txtZD
End Sub


Private Sub txtZD_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtZD_LostFocus()
txt_LostFocus txtZD
End Sub


Private Sub txtZG_GotFocus()
txt_GotFocus txtZG
End Sub


Private Sub txtZG_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtZG_LostFocus()
txt_LostFocus txtZG
End Sub


Private Sub txtZH_GotFocus()
txt_GotFocus txtZH
End Sub


Private Sub txtZH_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtZH_LostFocus()
txt_LostFocus txtZH
End Sub


Private Sub txtZI_GotFocus()
txt_GotFocus txtZI
End Sub


Private Sub txtZI_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtZI_LostFocus()
txt_LostFocus txtZI
End Sub


Private Sub txtZJ_GotFocus()
txt_GotFocus txtZJ
End Sub


Private Sub txtZJ_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtZJ_LostFocus()
txt_LostFocus txtZJ
End Sub


Private Sub UpDownDGI_2561_DownClick()
recDGI_2561.Method = "MoveNext"
cmddbDGI_2561
End Sub

Private Sub UpDownDGI_2561_UpClick()
recDGI_2561.Method = "MovePrevious"
cmddbDGI_2561
End Sub





Public Sub cmdPrint_Call()
prtrecDGI_2561 = recDGI_2561
prtDGI_2561_Monitor " "

End Sub
