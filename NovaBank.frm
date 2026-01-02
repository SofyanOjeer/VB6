VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNovaBank 
   Caption         =   "NovaBank : interface"
   ClientHeight    =   7140
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   10080
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   0
      Top             =   0
      Width           =   3585
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5700
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   10054
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Interface NovaBank"
      TabPicture(0)   =   "NovaBank.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraFolder"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraNovaBank"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "NovaBank.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdPrintSIT"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cmdPrintSIT 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Impression du fichier SIT"
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
         Left            =   -74280
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1920
         Width           =   2880
      End
      Begin VB.Frame fraNovaBank 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   8895
         Begin VB.CommandButton cmdOK_Compta 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Interface Comptable"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   2760
            Width           =   3120
         End
         Begin VB.FileListBox filDoc 
            Height          =   480
            Left            =   360
            TabIndex        =   6
            Top             =   720
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CommandButton cmdOK_Virement 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Interface Virements SIT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   2760
            Width           =   3120
         End
         Begin MSFlexGridLib.MSFlexGrid fgNovaBank_Compta 
            Height          =   2250
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   3969
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   14737632
            ForeColor       =   12582912
            BackColorFixed  =   12632256
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   0
            GridLinesFixed  =   1
            FormatString    =   "<fichiers       Comptables           |<Date dernière modif    "
         End
         Begin MSFlexGridLib.MSFlexGrid fgNovaBank_Virement 
            Height          =   2250
            Left            =   4560
            TabIndex        =   9
            Top             =   240
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   3969
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   14737632
            ForeColor       =   12582912
            BackColorFixed  =   12632256
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   0
            GridLinesFixed  =   1
            FormatString    =   "<fichiers   Virements                      |<Date dernière modif   "
         End
      End
      Begin VB.Frame fraFolder 
         Height          =   1215
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   8895
         Begin VB.CommandButton cmdSelect 
            BackColor       =   &H00C0FFC0&
            Caption         =   "OK"
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
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   360
            Width           =   2880
         End
         Begin VB.OptionButton optTest 
            Caption         =   "Test"
            Height          =   255
            Left            =   720
            TabIndex        =   11
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton optProd 
            Caption         =   "Production"
            Height          =   375
            Left            =   720
            TabIndex        =   10
            Top             =   240
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "frmNovaBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim X As String, X1 As String, I As Integer
Dim Msg As String, valX As String
Dim currentMethod As String, lastMethod As String
Dim NovaBankAut As typeAuthorization

Dim IdShell

Dim blncmdOk_Run As Boolean, blnAuto_NovaBank As Boolean

Dim blnImportMsgFile As Boolean

Dim paramNovabank_Environnement As String
Dim paramNovaBank_Archive As String, paramNovaBank_Folder As String
Dim paramNovaBank_Pattern_Compta As String, paramNovaBank_Pattern_Virement As String
Dim paramNovabank_Virement_SIT As String
Dim paramNovabank_AS400_IN_Compta As String, paramNovabank_AS400_IN_SIT As String

Dim xIn As String, xOut As String

Dim meNumero As typeNumeroP0
Public Sub Msg_Rcv(Msg As String)

    Dim X As String
    Call BiaPgmAut_Init(mId$(Msg, 1, 12), NovaBankAut)
    cmdReset
    
'JPLTST  MsgBox "!!! @AUTO_NOVABK", vbCritical, "frmNovaBnak.Msg_Rcv"
'JPLTST  Mid$(Msg, 1, 12) = "@AUTO_NOVABK"


If Not NovaBankAut.Xspécial Then
    optProd = True
    cmdSelect_Click
End If

        Select Case UCase$(Trim(mId$(Msg, 1, 12)))
            Case "@AUTO_NOVABK":     blnAuto_NovaBank = True: Auto_NovaBank
            Case Else: blnAuto_NovaBank = False
        End Select
    
End Sub


Public Sub cmdContext_Quit()
If fraNovaBank.Enabled Then
    cmdReset
Else
    If blnMsgBox_Quit Then
       X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
    Else
       X = vbYes
    End If
    If X = vbYes Then Unload Me
End If
End Sub


Public Sub cmdContext_Return()

SendKeys "{TAB}"

End Sub

'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
currentActiveControl_Name = C.Name
End Sub

'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
lstErr.Clear
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub

Private Sub cmdContext_Click()
cmdContext_Quit

End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext
End Sub


Private Sub cmdOK_Compta_Click()
'****************************************************************************
' Procédure de lancement du traitement de génération du fichier SIT         *
'****************************************************************************

    Dim I As Integer
    blncmdOk_Run = True
    Me.Enabled = False
    lstErr.Clear
    
    ' Lecture de chacune des lignes de la FlexGrid fgNovaBank_Compta
    ' Avec chaque nom de fichier lu en colonne 0, appel de la procédure CmdOK_Compta_SIT
    '-------------------------------------------------------------------------------------
    For I = 1 To fgNovaBank_Compta.Rows - 1
        fgNovaBank_Compta.Row = I
        fgNovaBank_Compta.Col = 0:
        Call lstErr_AddItem(lstErr, cmdContext, "TRT : " & fgNovaBank_Compta.Text)
        
        cmdOK_Compta_SIT fgNovaBank_Compta.Text
        
    Next I

    ' Rechargement des FlexGrid
    '--------------------------
    fgNovaBank_Compta_Load
    fgNovaBank_Virement_Load

    Me.Enabled = True
    'AppActivate Me.Caption
    blncmdOk_Run = False

End Sub

Private Sub cmdOK_Compta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdOK_Compta
End Sub


Private Sub cmdOK_Virement_Click()
'****************************************************************************
' Procédure de lancement du traitement de génération du fichier SIT         *
'****************************************************************************

    Dim I As Integer
    blncmdOk_Run = True
    Me.Enabled = False
    lstErr.Clear
    
    ' Lecture de chacune des lignes de la FlexGrid fgNovaBank_Virement
    ' Avec chaque nom de fichier lu en colonne 0, appel de la procédure CmdOK_Virement_SIT
    '-------------------------------------------------------------------------------------
    For I = 1 To fgNovaBank_Virement.Rows - 1
        fgNovaBank_Virement.Row = I
        fgNovaBank_Virement.Col = 0:
        Call lstErr_AddItem(lstErr, cmdContext, "TRT : " & fgNovaBank_Virement.Text)
        cmdOK_Virement_SIT fgNovaBank_Virement.Text
    Next I

    ' Rechargement des FlexGrid
    '--------------------------
    fgNovaBank_Compta_Load
    fgNovaBank_Virement_Load

    Me.Enabled = True
    'AppActivate Me.Caption
    blncmdOk_Run = False
    
End Sub

Private Sub cmdOK_Virement_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdOK_Virement
End Sub


'******************************************************************************************
'* Click sur le bouron Impression du fichier SIT                                          *
'******************************************************************************************
Private Sub cmdPrintSIT_Click()
    
    'Appel de la procédure prtNovabank_SIT
    '-------------------------------------
    prtNovaBank_SIT paramNovabank_Virement_SIT
    
End Sub

'******************************************************************************************
'* Click sur le bouron OK de la Frame Folder pour choix de l'environnement Prod/Test      *
'******************************************************************************************
Private Sub cmdSelect_Click()
    
    Dim V
    
    ' Désactication de la Frame Folder
    '---------------------------------
    fraFolder.Enabled = False
    
    ' Interception des boutons d'option Optprod ou OptTest
    '-----------------------------------------------------
    If optProd Then
        paramNovabank_Environnement = "Prod"
    Else
        paramNovabank_Environnement = "Test"
    End If
           
    ' Appel de la fonction Param_Init pour récupération des paramètres d'environnement
    '---------------------------------------------------------------------------------
    ' Si l'appel de la fonction aboutit correctement :
    '       - Cacher la liste des Erreurs
    '       - Activer la Frame NovaBank
    '       - Charger la FlexGrid avec les nom de fichiers d'inteface comptable
    '       - Charger la FlexGrid avec les nom de fichiers d'inteface Virements
    '---------------------------------------------------------------------------------
    
    V = param_Init(paramNovabank_Environnement)
    
    If IsNull(V) Then
        lstErr.Clear: lstErr.Visible = False
        fraNovaBank.Enabled = True
        fgNovaBank_Compta_Load
        fgNovaBank_Virement_Load
    Else
        Call MsgBox(V, vbCritical, "frmNovaBank")
    End If

End Sub

Private Sub cmdSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdSelect
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case Is = 13: KeyCode = 0:  cmdContext_Return
        Case Is = 27:  cmdContext_Quit                      'Touche Echap
        Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
    End Select

End Sub

Private Sub Form_Load()
    Set XForm = Me
    Call MeInit(arrTagNb)
    ReDim arrTag(arrTagNb + 1)
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub fraFolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub optProd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optProd
End Sub


Private Sub optTest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optTest
End Sub


Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set SSTab1
End Sub


Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
         If TypeOf xobj Is CommandButton Or TypeOf xobj Is ListBox Or TypeOf xobj Is MSFlexGrid Then
           xobj.BackColor = MouseMoveActiveControl.BackColor
        Else
            xobj.ForeColor = MouseMoveActiveControl.ForeColor
        End If
        Exit For
    End If
Next xobj

End Sub

Public Sub MouseMoveActiveControl_Set(C As Control)
If MouseMoveActiveControl_Name <> C.Name Then
    MouseMoveActiveControl_Reset
    If Not C.Enabled Then
        MouseMoveActiveControl_Name = ""
    Else
        MouseMoveActiveControl_Name = C.Name
        If TypeOf C Is CommandButton Or TypeOf C Is ListBox Or TypeOf C Is MSFlexGrid Then
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            MouseMoveActiveControl.ForeColor = C.ForeColor
             C.ForeColor = MouseMoveUsr.ForeColor
        End If
    End If
End If

End Sub

'********************************************************************
' Fonction de recherche des attributs d'environnement               *
'********************************************************************
Public Function param_Init(lEnvironnement As String)
Dim V
param_Init = Null

recElpTable_Init recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "NovaBank"
Call lstErr_Clear(frmNovaBank.lstErr, frmNovaBank.cmdContext, "BIA.mdb : table : " & recElpTable.Id)

recElpTable.K1 = "Folder"
recElpTable.K2 = lEnvironnement
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramNovaBank_Folder = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmNovaBank.lstErr, frmNovaBank.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "AS400_IN"
recElpTable.K2 = "Compta"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramNovabank_AS400_IN_Compta = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmNovaBank.lstErr, frmNovaBank.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "SIT"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramNovabank_AS400_IN_SIT = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmNovaBank.lstErr, frmNovaBank.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Pattern"
recElpTable.K2 = "Compta"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramNovaBank_Pattern_Compta = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmNovaBank.lstErr, frmNovaBank.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Pattern"
recElpTable.K2 = "Virement"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramNovaBank_Pattern_Virement = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmNovaBank.lstErr, frmNovaBank.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Archive"
recElpTable.K2 = ""
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramNovaBank_Archive = paramNovaBank_Folder & paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmNovaBank.lstErr, frmNovaBank.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "SIT"
recElpTable.K2 = ""
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramNovabank_Virement_SIT = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmNovaBank.lstErr, frmNovaBank.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

Exit Function

Table_Error:
param_Init = V
Exit Function

Memo_Error:
param_Init = "Memo"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "srvTI.Param_Init"
Exit Function

Num_Error:
param_Init = "Num"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : " & recElpTable.Memo & " :Mémo non numérique", vbCritical, "srvTI.Param_Init"
End Function



'***************************************************************************************
' Procédure de chargement des noms de fichiers d'interface Virements                   *
'***************************************************************************************

Public Sub fgNovaBank_Virement_Load()

    Dim I As Integer, K As Integer, X As String, L As Integer, iSession As Integer

    'Chargement de la File Liste Boxe
    '--------------------------------
    filDoc.Path = paramNovaBank_Folder
    filDoc.Pattern = paramNovaBank_Pattern_Virement

    'Désactiver le bouton d'envoi du traitement tant que la liste n'est pas constituée.
    '----------------------------------------------------------------------------------
    cmdOK_Virement.Enabled = False

    ' Chargement de la FlexGrid Compta avec le contenu de la File Liste Box
    '----------------------------------------------------------------------
    fgNovaBank_Virement.Redraw = False
    fgNovaBank_Virement.Rows = 1
    fgNovaBank_Virement.Enabled = True
    
    For I = 0 To filDoc.ListCount - 1
        filDoc.ListIndex = I
        Set msFile = msFileSystem.GetFile(filDoc.Path & "\" & filDoc.Filename)
        fgNovaBank_Virement.Rows = fgNovaBank_Virement.Rows + 1
        fgNovaBank_Virement.Row = fgNovaBank_Virement.Rows - 1
        fgNovaBank_Virement.Col = 0: fgNovaBank_Virement.Text = Trim(filDoc.Filename)
        fgNovaBank_Virement.Col = 1: fgNovaBank_Virement.Text = msFile.DateLastModified
        cmdOK_Virement.Enabled = NovaBankAut.Virement    '  True
    Next I
    
fgNovaBank_Virement.Redraw = True

End Sub


Public Sub cmdOK_Virement_SIT(lFileName As String)
'**********************************************************************************
' Procédure de constitution d'un fichier Virement SIT à partir d'un fichier OC160 *
'**********************************************************************************
Dim xFileNameSit As String

Dim xFileName As String
Dim I As Integer
Dim x320 As String * 320

Dim CodeEnregistrement As String
Dim NbreOpérations As Integer
Dim MontantTotal As Currency
Dim MontantOpération As String
Dim CodeEtablissementDestinataire As String
Dim CodeGuichetDestinataire As String
Dim CompteDestinataire As String
Dim NomDestinataire As String
Dim CodeEtablissementDonneurOrdre As String
Dim CodeGuichetDonneurOrdre As String
Dim NuméroCompteDonneurOrdre As String
Dim NuméroEmetteur As String
Dim Domiciliation As String
Dim LibelléDestinataire As String
Dim DateJournée As Date
Dim DateRèglement As Date
Dim DateSit As String, X8 As String * 8
Dim Monnaie As String
Dim LibelléDevise As String
Dim NuméroRemise As Integer

Dim blnError As Boolean

On Error GoTo Error_Handle

' Test de présence d'un fichier SIT déjà existant.
'---------------------------------------------------------------------------------------
X = Dir(paramNovabank_Virement_SIT)
If X <> "" Then
    Call lstErr_AddItem(lstErr, cmdContext, "! SIT en cours : " & paramNovabank_Virement_SIT)
    Exit Sub
End If

X = Dir(paramNovabank_AS400_IN_SIT)
If X <> "" Then
    Call lstErr_AddItem(lstErr, cmdContext, "! SIT en cours : " & paramNovabank_AS400_IN_SIT)
    Exit Sub
End If


xFileName = paramNovaBank_Folder & lFileName
xFileNameSit = paramNovaBank_Folder & lFileName & "_W"
'Date SIT en accès dans la table FICDATP1 de BIA.MDB +  Numéro de remise SIT
'--------------------------------------------------------------------------------------
DateJournée = mId$(DSys, 3, 6)
If time_Hms < "133000" Then
    X8 = DSys
Else
    X8 = dateBIA("Ouvré", 1, DSys)
End If

DateSit = mId$(X8, 3, 6)                  ' A revoir avec appel fonction date (FICDATP1)
DateRèglement = DateSit

NuméroRemise = 0
recNumeroP0_Init meNumero
meNumero.NOCTR = 134
meNumero.Method = "Add"
Call srvNumeroP0_Update(meNumero)
NuméroRemise = meNumero.CTEUR               ' A revoir (Compteur AS/400 : 134)

Open xFileName For Input As #1
'Open paramNovabank_Virement_SIT For Output As #2
Open xFileNameSit For Output As #2

xIn = ""
xOut = ""
blnError = False

'Lecture du fichier Trilog : OC160
'---------------------------------
Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
         
    CodeEnregistrement = mId$(xIn, 1, 2)
    
    ' Lecture de l'enregistrement Entête du fichier Trilog
    '-----------------------------------------------------
    If CodeEnregistrement = "03" Then
    
        CodeGuichetDonneurOrdre = mId$(xIn, 87, 5)
        NuméroCompteDonneurOrdre = mId$(xIn, 92, 11)
        CodeEtablissementDonneurOrdre = mId$(xIn, 150, 5)
        
        Monnaie = mId$(xIn, 81, 1)
        If Monnaie = "F" Then
            LibelléDevise = "FRF2"
        Else
            LibelléDevise = "EUR2"
        End If
            
        ' Génération Entête de remise SIT ALLER
        '--------------------------------------
        x320 = Space$(320)
        Mid$(x320, 1, 2) = "<>"
        Mid$(x320, 3, 4) = "0320"
        Mid$(x320, 7, 2) = "01"
        Mid$(x320, 9, 5) = "12179"
        Mid$(x320, 14, 2) = String(2, "0")
        Mid$(x320, 16, 4) = "0320"
        Mid$(x320, 20, 2) = String(2, "0")
        Mid$(x320, 22, 1) = String(1, "0")
        Mid$(x320, 23, 6) = Format(NuméroRemise, "000000")
        Mid$(x320, 29, 6) = Format(DateJournée, "000000")
        Mid$(x320, 35, 6) = Format(DateSit, "000000")
        Mid$(x320, 41, 5) = String(5, "0")
        Mid$(x320, 46, 1) = "1"
        Mid$(x320, 47, 3) = "001"
        Mid$(x320, 50, 6) = Format(DateRèglement, "000000")
        Mid$(x320, 56, 4) = LibelléDevise
        Mid$(x320, 60, 60) = String(60, "0")
        Mid$(x320, 120, 1) = "0"
        Mid$(x320, 121, 5) = "04970"
        Mid$(x320, 126, 11) = "00001080042"
        Mid$(x320, 137, 184) = String(184, " ")
      
        xOut = x320
        Print #2, xOut
        
    End If
    
    
    ' Lecture de l'enregistrement Détail du fichier Trilog
    '-----------------------------------------------------
    If CodeEnregistrement = "06" Then
         
        MontantOpération = mId$(xIn, 103, 16)
        If Not IsNumeric(MontantOpération) Then
            MsgBox "Montant Opération : " & MontantOpération, vbInformation, "frmNovaBank.cmOK_Virement_Sit"
            blnError = True
        Else
            If MontantOpération <= 0 Then
                MsgBox "Montant Opération : " & MontantOpération, vbInformation, "frmNovaBank.cmOK_Virement_Sit"
                blnError = True
            End If

        End If

        NbreOpérations = NbreOpérations + 1
        MontantTotal = MontantTotal + CCur(MontantOpération)
        CodeEtablissementDestinataire = mId$(xIn, 150, 5)
        If Not IsNumeric(CodeEtablissementDestinataire) Then
            MsgBox "Code Etablissement Destinataire : " & CodeEtablissementDestinataire, vbInformation, "frmNovaBank.cmOK_Virement_Sit"
            blnError = True
        End If
        
        CodeGuichetDestinataire = mId$(xIn, 87, 5)
         If Not IsNumeric(CodeGuichetDestinataire) Then
            MsgBox "Code Guichet Destinataire : " & CodeGuichetDestinataire, vbInformation, "frmNovaBank.cmOK_Virement_Sit"
            blnError = True
        End If
       CompteDestinataire = mId$(xIn, 92, 11)
        NomDestinataire = mId$(xIn, 31, 24)
        Domiciliation = mId$(xIn, 55, 20)
        LibelléDestinataire = mId$(xIn, 119, 29)
        
        ' Génération Opération Détail de remise SIT ALLER
        '------------------------------------------------
        x320 = Space$(320)
        Mid$(x320, 1, 2) = "<>"
        Mid$(x320, 3, 4) = "0320"
        Mid$(x320, 7, 2) = "02"
        Mid$(x320, 9, 3) = "120"
        Mid$(x320, 12, 1) = "1"
        Mid$(x320, 13, 7) = "12179  "
        Mid$(x320, 20, 1) = "1"
        Mid$(x320, 21, 7) = CodeEtablissementDestinataire & "  "
        Mid$(x320, 28, 5) = CodeGuichetDestinataire
        Mid$(x320, 33, 4) = LibelléDevise
        Mid$(x320, 37, 16) = Format(MontantOpération, "0000000000000000")
        Mid$(x320, 53, 6) = Format(DateRèglement, "000000")
        Mid$(x320, 59, 1) = "0"
        Mid$(x320, 60, 1) = "0"
        Mid$(x320, 61, 6) = "000000"
        Mid$(x320, 67, 1) = "0"
        Mid$(x320, 68, 2) = "00"
        Mid$(x320, 70, 10) = String(10, " ")
        Mid$(x320, 80, 1) = Monnaie
        Mid$(x320, 81, 4) = "0001"
        Mid$(x320, 85, 16) = String(16, " ")
        Mid$(x320, 101, 10) = "COMPTA" & String(4, " ")
        Mid$(x320, 111, 16) = Format(CodeGuichetDonneurOrdre, "00000") & Format(NuméroCompteDonneurOrdre, "00000000000")
        Mid$(x320, 127, 14) = CompteDestinataire & "   "
        Mid$(x320, 141, 24) = "Banque Interc. Arabe    ' A"
        Mid$(x320, 165, 6) = NuméroEmetteur
        Mid$(x320, 171, 4) = String(4, " ")
        Mid$(x320, 175, 18) = String(18, " ")
        Mid$(x320, 193, 24) = NomDestinataire
        Mid$(x320, 217, 24) = Domiciliation
        Mid$(x320, 241, 1) = "0"
        Mid$(x320, 242, 1) = " "                                    'Indicateur balance des paiements
        Mid$(x320, 243, 1) = " "
        Mid$(x320, 244, 1) = " "
        Mid$(x320, 245, 1) = " "
        Mid$(x320, 246, 11) = String(11, " ")
        Mid$(x320, 257, 32) = LibelléDestinataire & "   "           'Libellé origine 29 + 3 blancs
        Mid$(x320, 289, 32) = String(32, " ")
        xOut = x320
        Print #2, xOut
        
    End If
    
    
    ' Lecture de l'enregistrement Total du fichier Trilog
    '-----------------------------------------------------
    If CodeEnregistrement = "08" Then
    
        ' Génération Récapitulatif de remise SIT ALLER
        '---------------------------------------------
        x320 = Space$(320)
        Mid$(x320, 1, 2) = "<>"
        Mid$(x320, 3, 4) = "0320"
        Mid$(x320, 7, 2) = "09"
        Mid$(x320, 9, 8) = Format(NbreOpérations, "00000000")
        Mid$(x320, 17, 18) = Format(MontantTotal, "000000000000000000")
        Mid$(x320, 35, 60) = String(60, " ")
        Mid$(x320, 95, 184) = String(226, " ")
        
        xOut = x320
        Print #2, xOut
        
    End If
    
Loop

Close

If blnError Then
    MsgBox "ni transmission SIT, ni Compta", vbCritical, "frmNovaBank.cmOK_Virement_Sit"
Else
    prtNovaBank_SIT xFileNameSit


    msFileSystem.MoveFile xFileName, paramNovaBank_Archive & DSys & "_" & time_Hms & "_" & lFileName

    If optProd Then
        msFileSystem.CopyFile xFileNameSit, paramNovabank_AS400_IN_SIT
    
        srvAs400Cmd.SBMJOB "NOVASITCL"
    
        msFileSystem.MoveFile xFileNameSit, paramNovabank_Virement_SIT
    Else
    
        MsgBox "TEST :ni transmission SIT, ni Compta", vbCritical, "frmNovaBank.cmOK_Virement_Sit"
   End If
End If

Exit Sub

Error_Handle:
Close
MsgBox lFileName & ":" & Error, vbCritical, "cmdOK_Virement_SIT"
End Sub

Public Sub cmdOK_Compta_SIT(lFileName As String)
'**********************************************************************************
' Procédure de constitution d'un fichier Compta à partir d'un fichier OC160 *
'**********************************************************************************
On Error GoTo Error_Handler

X = Dir(paramNovabank_AS400_IN_Compta)
If X <> "" Then Call lstErr_AddItem(lstErr, cmdContext, "! \\AS400_IN\NovaBank.dat en cours"): Exit Sub

msFileSystem.CopyFile paramNovaBank_Folder & lFileName, paramNovabank_AS400_IN_Compta
msFileSystem.MoveFile paramNovaBank_Folder & lFileName, paramNovaBank_Archive & DSys & "_" & time_Hms & "_" & lFileName

srvAs400Cmd.SBMJOB "NOVABANKCL"




Exit Sub

Error_Handler:
MsgBox lFileName & ":" & Error, vbCritical, "cmdOK_Virement_SIT"
End Sub


'***************************************************************************************
' Procédure de chargement des noms de fichiers d'interface comptable                   *
'***************************************************************************************

Public Sub fgNovaBank_Compta_Load()

Dim I As Integer, K As Integer, X As String, L As Integer, iSession As Integer

' Chargement de File Liste Box
'-----------------------------
filDoc.Path = paramNovaBank_Folder
filDoc.Pattern = paramNovaBank_Pattern_Compta

'Désactiver le bouton d'envoi du traitement tant que la liste n'est pas constituée.
'----------------------------------------------------------------------------------
cmdOK_Compta.Enabled = False

' Chargement de la FlexGrid Compta avec le contenu de la File Liste Box
'----------------------------------------------------------------------
fgNovaBank_Compta.Redraw = False
fgNovaBank_Compta.Rows = 1
fgNovaBank_Compta.Enabled = True

For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.Path & "\" & filDoc.Filename)
    fgNovaBank_Compta.Rows = fgNovaBank_Compta.Rows + 1
    fgNovaBank_Compta.Row = fgNovaBank_Compta.Rows - 1
    fgNovaBank_Compta.Col = 0: fgNovaBank_Compta.Text = Trim(filDoc.Filename)
    fgNovaBank_Compta.Col = 1: fgNovaBank_Compta.Text = msFile.DateLastModified
    cmdOK_Compta.Enabled = NovaBankAut.Xspécial  '''' True
Next I

fgNovaBank_Compta.Redraw = True

End Sub

Public Sub Auto_NovaBank()
Dim blnOk As Boolean
optProd = True
cmdSelect_Click


        If cmdOK_Compta.Enabled Then blnOk = False: cmdOK_Compta_Click: cmdOK_Compta.Enabled = False
    '    If cmdOK_Virement.Enabled Then blnOk = False: cmdOK_Virement_Click:cmdOK_Virement.Enabled=false
Unload Me

End Sub

Public Sub cmdOK_Run(C As CommandButton)
blncmdOk_Run = True
Me.Enabled = False

'Select Case Trim(C.Name)
'    Case "cmdOK_Compta":        cmdOK_Compta.Enabled = False
'                                Compta_Put
'    Case "cmdOk_SAA_Corona":    cmdOk_SAA_Corona.Enabled = False
'                                SAA_Corona_Put
'    Case "cmdOk_Loro":          cmdOk_Loro.Enabled = False
'                                Loro_Put
'    Case "cmdOK_Nostro":        cmdOK_Nostro.Enabled = False
'                                Nostro_Put
'    Case "cmdOK_Virement":        cmdOK_Virement.Enabled = False
'                                Virement_Put
'End Select

Me.Enabled = True
'AppActivate Me.Caption
blncmdOk_Run = False


End Sub


Public Sub cmdReset()
    lstErr.Clear
    cmdContext.Caption = constcmdAbandonner
    SSTab1.Tab = 0
    fraFolder.Enabled = True
    fraNovaBank.Enabled = False
    optTest = True
    cmdSelect.Enabled = True
    fgNovaBank_Compta.Clear
    fgNovaBank_Virement.Clear
End Sub
