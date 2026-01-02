VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmElpKM 
   AutoRedraw      =   -1  'True
   Caption         =   "Base d'informations"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14235
   Icon            =   "ElpKM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14235
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8040
      TabIndex        =   5
      Top             =   45
      Width           =   5175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   -15
      TabIndex        =   3
      Top             =   570
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Sélection"
      TabPicture(0)   =   "ElpKM.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lstId"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Informations sélectionnées"
      TabPicture(1)   =   "ElpKM.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgSelect"
      Tab(1).Control(1)=   "lstSelectUsr"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Documents associés"
      TabPicture(2)   =   "ElpKM.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Utilisateurs"
      TabPicture(3)   =   "ElpKM.frx":035E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "fraUsr"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.ListBox lstSelectUsr 
         BackColor       =   &H00FFFFFF&
         Height          =   7275
         Left            =   -66960
         TabIndex        =   17
         Top             =   720
         Width           =   5640
      End
      Begin VB.Frame fraUsr 
         Height          =   8430
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   13560
         Begin VB.Frame fraSelect_Niveau 
            Height          =   495
            Left            =   9960
            TabIndex        =   26
            Top             =   7800
            Width           =   3495
            Begin VB.OptionButton optSAB_MNU_L2 
               Alignment       =   1  'Right Justify
               Caption         =   " 2"
               Height          =   240
               Left            =   1680
               TabIndex        =   30
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton optSAB_MNU_L3 
               Alignment       =   1  'Right Justify
               Caption         =   " 3"
               Height          =   240
               Left            =   2280
               TabIndex        =   29
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton optSAB_MNU_L4 
               Alignment       =   1  'Right Justify
               Caption         =   "4"
               Height          =   240
               Left            =   2880
               TabIndex        =   28
               Top             =   240
               Value           =   -1  'True
               Width           =   435
            End
            Begin VB.OptionButton optSAB_MNU_L1 
               Alignment       =   1  'Right Justify
               Caption         =   "afficher niveau 1"
               Height          =   240
               Left            =   30
               TabIndex        =   27
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.CommandButton cmdZMNUMEN0_Scan 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Rechercher texte >"
            Height          =   405
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   7920
            Width           =   1515
         End
         Begin VB.TextBox txtZMNUMEN0_Scan 
            Height          =   315
            Left            =   8280
            TabIndex        =   18
            Top             =   7920
            Width           =   1650
         End
         Begin VB.PictureBox picUsr 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   1650
            Left            =   120
            ScaleHeight     =   1590
            ScaleWidth      =   6420
            TabIndex        =   16
            Top             =   120
            Width           =   6480
         End
         Begin VB.ListBox lstUsr 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5940
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   1920
            Width           =   6585
         End
         Begin MSFlexGridLib.MSFlexGrid fgZMNUMEN0 
            Height          =   7665
            Left            =   6720
            TabIndex        =   25
            Top             =   120
            Width           =   6720
            _ExtentX        =   11853
            _ExtentY        =   13520
            _Version        =   393216
            Rows            =   1
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   250
            BackColor       =   15007711
            ForeColor       =   12582912
            BackColorFixed  =   12648384
            ForeColorFixed  =   32768
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   2
            AllowUserResizing=   3
            FormatString    =   $"ElpKM.frx":037A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Sans Unicode"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame lstId 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8460
         Left            =   -74880
         TabIndex        =   6
         Top             =   360
         Width           =   13575
         Begin VB.Frame fraSelect1100 
            Height          =   1530
            Left            =   2640
            TabIndex        =   21
            Top             =   1755
            Width           =   2625
            Begin VB.OptionButton optSelect11000_LF 
               Caption         =   "uniquement LF"
               Height          =   225
               Left            =   120
               TabIndex        =   24
               Top             =   1035
               Width           =   2340
            End
            Begin VB.OptionButton optSelect11000_PF 
               Caption         =   "uniquement PF"
               Height          =   210
               Left            =   120
               TabIndex        =   23
               Top             =   705
               Width           =   2160
            End
            Begin VB.OptionButton optSelect11000_All 
               Caption         =   "Tous"
               Height          =   195
               Left            =   120
               TabIndex        =   22
               Top             =   345
               Value           =   -1  'True
               Width           =   2115
            End
         End
         Begin VB.ListBox lstW 
            Height          =   2010
            Left            =   9120
            Sorted          =   -1  'True
            TabIndex        =   20
            Top             =   3360
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.TextBox txtElpKMInfo_Id 
            Height          =   285
            Left            =   2295
            TabIndex        =   11
            Top             =   345
            Width           =   1575
         End
         Begin VB.CommandButton cmdElpKMInfo_AddNew 
            BackColor       =   &H0080FF80&
            Caption         =   "Ajouter un dossier"
            Height          =   975
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   4560
            Width           =   2415
         End
         Begin VB.ListBox lstElpKM_Classe 
            Height          =   3375
            ItemData        =   "ElpKM.frx":0499
            Left            =   225
            List            =   "ElpKM.frx":04A0
            TabIndex        =   9
            Top             =   1800
            Width           =   2235
         End
         Begin VB.TextBox txtSelect 
            Height          =   285
            Left            =   2340
            TabIndex        =   8
            Top             =   930
            Width           =   6360
         End
         Begin VB.CommandButton cmdSelect 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Rechercher"
            Height          =   975
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1920
            Width           =   2415
         End
         Begin VB.Label lblElpKMInfo_Id 
            Caption         =   "Référence"
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblSelect 
            Caption         =   "mots recherchés ==>                                       dans la documentation"
            Height          =   705
            Left            =   360
            TabIndex        =   12
            Top             =   960
            Width           =   1740
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSelect 
         Height          =   8325
         Left            =   -74865
         TabIndex        =   4
         Top             =   420
         Width           =   13560
         _ExtentX        =   23918
         _ExtentY        =   14684
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         RowHeightMin    =   250
         BackColor       =   14737632
         ForeColor       =   12582912
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12648384
         BackColorBkg    =   14737632
         AllowBigSelection=   0   'False
         TextStyleFixed  =   4
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"ElpKM.frx":04B5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "ElpKM.frx":05F8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label libRéférenceInterne 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   6375
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuExport_Options 
         Caption         =   "Export récap <Options / Groupes >"
      End
      Begin VB.Menu mnuContextOption 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuSelect1000 
      Caption         =   "Select1000"
      Visible         =   0   'False
      Begin VB.Menu mnuSelectUsr 
         Caption         =   "Arboresrence de l'option"
      End
      Begin VB.Menu mnuSelectUsrAll 
         Caption         =   "Utilisateurs habilités"
      End
   End
   Begin VB.Menu mnuSelect11000 
      Caption         =   "Select11000"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect11000_Display 
         Caption         =   "Afficher description de l'objet"
      End
      Begin VB.Menu mnuSelect11000_DisplayFields 
         Caption         =   "Afficher objet & champs"
      End
      Begin VB.Menu mnuSelect11000_PrintFields 
         Caption         =   "Imprimer objet séléctionné (& champs)"
      End
      Begin VB.Menu mnuSelect11000_X1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelect11000_PrintFields_All 
         Caption         =   "Imprimer TOUS les objets  (& champs)"
      End
   End
   Begin VB.Menu mnuSelect12000 
      Caption         =   "Select12000"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect12000_Display 
         Caption         =   "Afficher description de l'objet"
      End
      Begin VB.Menu mnuSelect12000_DisplayFields 
         Caption         =   "Afficher objet & champs"
      End
      Begin VB.Menu mnuSelect12000_PrintFields 
         Caption         =   "Imprimer objet & champs"
      End
   End
   Begin VB.Menu mnuPrint1 
      Caption         =   "Print1"
      Visible         =   0   'False
      Begin VB.Menu mnufgSelect_Print 
         Caption         =   "Imprimer la liste des documents sélectionnés"
      End
   End
   Begin VB.Menu mnuPrint3 
      Caption         =   "Print3"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint3_GroupUser 
         Caption         =   "Imprimer la liste ""Groupe / Utilisateur"""
      End
      Begin VB.Menu mnuPrint3_UserGroup 
         Caption         =   "Imprimer la liste ""Utilisateur / Groupe"""
      End
      Begin VB.Menu mnuPrint3_X1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint3All 
         Caption         =   "Imprimer toutes les options de l'utilisateur"
      End
      Begin VB.Menu mnuPrint3Menu 
         Caption         =   "Imprimer le menu sélectionné"
      End
      Begin VB.Menu mnuPrint3_X2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint3_xls 
         Caption         =   "exporter.xls  toutes les options de l'utilisateur"
      End
   End
   Begin VB.Menu mnufgZMNUMEN0 
      Caption         =   "mnufgZMNUMEN0"
      Visible         =   0   'False
      Begin VB.Menu mnufgZMNUMEN0_Doc 
         Caption         =   "Documents associés"
      End
      Begin VB.Menu mnufgZMNUMEN0_Print 
         Caption         =   "Imprimer le menu sélectionné"
      End
   End
   Begin VB.Menu mnulstUsr 
      Caption         =   "mnulstUsr"
      Visible         =   0   'False
      Begin VB.Menu mnulstUsr_Aut 
         Caption         =   "menu (options autorisées)"
      End
      Begin VB.Menu mnulstUsr_NonAut 
         Caption         =   "options non autorisées"
      End
      Begin VB.Menu mnuX4 
         Caption         =   "-"
      End
      Begin VB.Menu mnulstUsr_Aut_Export 
         Caption         =   "menu (options autorisées) Export "
      End
      Begin VB.Menu mnulstUsr_NonAut_Export 
         Caption         =   "options non autorisées Export"
      End
   End
End
Attribute VB_Name = "frmElpKM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim ElpKMAut As typeAuthorization

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean


Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean

Dim blnSetfocus As Boolean

Dim meElpKM_Classe As typeElpTable, xElpKM_Classe As typeElpTable
Dim mElpKM_Classe_Pass As Long, mElpKM_Classe_Abrégé As String, mElpKM_Classe_Folder As String
Dim lstElpKM_Classe_listIndex  As Integer

Dim meElpKMInfo As typeElpKmInfo, xElpKMInfo As typeElpKmInfo
Dim meElpKMIndex As typeElpKmIndex, xElpKMIndex As typeElpKmIndex

Dim arrElpKM_Classe() As Long, arrElpKMSrc_Id() As Long, arrElpKMInfo_Id() As String * 20, arrElpKM_Occurs() As Integer
Dim arrElpKM_Nb As Long, arrElpKM_Index As Long
Dim mElpKM_Classe As Long, mElpKMSrc_Id As Long, mElpKMInfo_Id As String * 20, mElpKMInfo_Description As String
Dim mElpKMInfo_ElpKmSrc_Id As Long

Dim mDocument_FileName As String, wDocument_FileName As String


Dim meZMNURUT0 As typeZMNURUT0, xZMNURUT0 As typeZMNURUT0
Dim meZMNUUTI0 As typeZMNUUTI0, xZMNUUTI0 As typeZMNUUTI0
Dim meZMNUHLB0 As typeZMNUHLB0, xZMNUHLB0 As typeZMNUHLB0
Dim arrZMNUHLB0() As typeZMNUHLB0, arrZMNUHLB0_Nb As Integer

Dim arrZMNUMEN0() As typeZMNUMEN0, arrZMNUMEN0_Nb As Long, meZMNUMEN0 As typeZMNUMEN0, xZMNUMEN0 As typeZMNUMEN0
Dim arrZMNUOPT0() As typeZMNUOPT0, arrZMNUOPT0_Nb As Long, meZMNUOPT0 As typeZMNUOPT0, xZMNUOPT0 As typeZMNUOPT0
Dim arrZMNUMEN0_Index As Long


Dim fgZMNUMEN0_Lib As String
'Dim meDSPFDY0 As typeDSPFDY0
'Dim meDSPFFDY0 As typeDSPFFDY0


Dim libMNURUT As String
'----------------------------------------------------------------------------
Dim blnZMNUMEN0_Export As Boolean
'----------------------------------------------------------------------------

Dim fgZMNUMEN0_FormatString As String, fgZMNUMEN0_K As Integer
Dim fgZMNUMEN0_RowDisplay As Integer, fgZMNUMEN0_RowClick As Integer
Dim fgZMNUMEN0_ColorClick As Long, fgZMNUMEN0_ColorDisplay As Long
Dim fgZMNUMEN0_Sort1 As Integer, fgZMNUMEN0_Sort2 As Integer
Dim fgZMNUMEN0_SortAD As Integer, fgZMNUMEN0_Sort1_Old As Integer
Dim fgZMNUMEN0_arrIndex As Integer
Dim blnfgZMNUMEN0_DisplayLine As Boolean
Public Sub fgZMNUMEN0_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgZMNUMEN0.Row

If lRow > 0 And lRow < fgZMNUMEN0.Rows Then
    fgZMNUMEN0.Row = lRow
    For I = 0 To fgZMNUMEN0_arrIndex
        fgZMNUMEN0.Col = I: fgZMNUMEN0.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgZMNUMEN0.Row = mRow
    If fgZMNUMEN0.Row > 0 Then
        lRow = fgZMNUMEN0.Row
        lColor_Old = fgZMNUMEN0.CellBackColor
        For I = 0 To fgZMNUMEN0_arrIndex
          fgZMNUMEN0.Col = I: fgZMNUMEN0.CellBackColor = lColor
        Next I
        fgZMNUMEN0.Col = 0
    End If
End If

End Sub

Private Sub fgZMNUMEN0_Display()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim kNiveau As Integer
On Error GoTo Error_Handler
SSTab1.Tab = 3
fgZMNUMEN0.Visible = False
fgZMNUMEN0_Reset
Call lstErr_Clear(lstErr, cmdContext, "Options : " & arrZMNUMEN0_Nb): DoEvents

fgZMNUMEN0.Rows = 1
fgZMNUMEN0.FormatString = fgZMNUMEN0_FormatString
currentAction = "fgZMNUMEN0_Display"
kNiveau = 4
If optSAB_MNU_L3 Then kNiveau = 3
If optSAB_MNU_L2 Then kNiveau = 2
If optSAB_MNU_L1 Then kNiveau = 1
For I = 1 To arrZMNUMEN0_Nb
         
    xZMNUMEN0 = arrZMNUMEN0(I)
    xZMNUOPT0 = arrZMNUOPT0(I)
    If xZMNUMEN0.Niveau <= kNiveau Then
        fgZMNUMEN0.Rows = fgZMNUMEN0.Rows + 1
        fgZMNUMEN0.Row = fgZMNUMEN0.Rows - 1
        fgZMNUMEN0_DisplayLine I
    End If
Next I

fgZMNUMEN0.Visible = True
If fgZMNUMEN0.Rows > 1 Then
    fgZMNUMEN0_Sort1 = 2: fgZMNUMEN0_Sort2 = 3: fgZMNUMEN0_Sort
    fgZMNUMEN0.Row = 1
End If
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Public Sub fgZMNUMEN0_DisplayLine(lIndex As Long)
Dim X As String
On Error Resume Next
fgZMNUMEN0.Col = 0: fgZMNUMEN0.Text = xZMNUMEN0.MNUMENCOD
Select Case xZMNUMEN0.Niveau
    Case "1": X = ""
    Case "2": X = "_  "
    Case "3": X = "_  _  "
    Case Else:  X = "_  _  _  "
End Select

fgZMNUMEN0.Col = 1: fgZMNUMEN0.Text = X & xZMNUOPT0.MNUOPTLIB
Select Case xZMNUMEN0.Niveau
    Case 1: fgZMNUMEN0.CellBackColor = &H80FF80
    Case 2: fgZMNUMEN0.CellBackColor = &HC0FFC0
End Select
fgZMNUMEN0.Col = 2: fgZMNUMEN0.Text = xZMNUMEN0.Hierarchie
 
fgZMNUMEN0.Col = fgZMNUMEN0_arrIndex: fgZMNUMEN0.Text = lIndex

End Sub

Public Sub fgZMNUMEN0_Load()
arrZMNUMEN0_Load meZMNUHLB0, arrZMNUMEN0, arrZMNUOPT0
arrZMNUMEN0_Nb = UBound(arrZMNUMEN0) - 1
fgZMNUMEN0_Display

End Sub
Public Sub fgZMNUMEN0_Reset()
fgZMNUMEN0.Clear
fgZMNUMEN0.FormatString = fgZMNUMEN0_FormatString
fgZMNUMEN0_Sort1 = 0: fgZMNUMEN0_Sort2 = 0
fgZMNUMEN0_Sort1_Old = -1
fgZMNUMEN0_RowDisplay = 0: fgZMNUMEN0_RowClick = 0
fgZMNUMEN0_arrIndex = fgZMNUMEN0.Cols - 1
blnfgZMNUMEN0_DisplayLine = False
fgZMNUMEN0_SortAD = 6
fgZMNUMEN0.LeftCol = 0

End Sub

Public Sub fgZMNUMEN0_Sort()
If fgZMNUMEN0.Rows > 1 Then
    fgZMNUMEN0.Row = 1
    fgZMNUMEN0.RowSel = fgZMNUMEN0.Rows - 1
    
    If fgZMNUMEN0_Sort1_Old = fgZMNUMEN0_Sort1 Then
        If fgZMNUMEN0_SortAD = 5 Then
            fgZMNUMEN0_SortAD = 6
        Else
            fgZMNUMEN0_SortAD = 5
        End If
    Else
        fgZMNUMEN0_SortAD = 5
    End If
    fgZMNUMEN0_Sort1_Old = fgZMNUMEN0_Sort1
    
    fgZMNUMEN0.Col = fgZMNUMEN0_Sort1
    fgZMNUMEN0.ColSel = fgZMNUMEN0_Sort2
    fgZMNUMEN0.Sort = fgZMNUMEN0_SortAD
End If

End Sub

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
blnControl = False
If currentAction <> "" Then
    X = MsgBox("Voulez-vous réellement abandonner la mise à jour?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
    If X = vbYes Then
        currentAction = ""
    Else
        Exit Sub
    End If
End If

If SSTab1.Tab = 2 Then SSTab1.Tab = 1: Exit Sub

If lstSelectUsr.Visible Then lstSelectUsr.Visible = False: Exit Sub

lstErr.Clear
If SSTab1.Tab > 0 Then
    SSTab1.Tab = SSTab1.Tab - 1
Else
End If
End Sub
Public Sub cmdSelect_Control()
Dim lMin As Double, lMax As Double
If Not Me.Enabled Then Exit Sub
Me.Enabled = False

'cmdOk.Visible = False
'cmdSave.Visible = False
blnControl = False
'blnSetfocus = False

lstSelectUsr.Visible = False
lstErr.Clear
lstErr.Height = 200
If Trim(txtSelect) = "" Then Call lstErr_AddItem(lstErr, cmdContext, "?  Précisez la recherche ")
If lstErr.ListCount > 0 Then
    lstErr.Visible = True
Else
    'cmdOk.Visible = blncmdOk_Visible
    'blnSetfocus = True: currentActiveControl_Name = "cmdOk"
End If

ExitSub:

Me.Enabled = True
    
blnControl = True

End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdElpKMInfo_AddNew_Click()
Dim X As String

cmdSelect_Control
If meElpKM_Classe.K2 = 1000 Then Call lstErr_AddItem(lstErr, cmdContext, "?SAB_Menu : interdit ")
X = Trim(txtElpKMInfo_Id)
If X = "" Then Call lstErr_AddItem(lstErr, cmdContext, "?  Précisez la référence ")

meElpKMInfo.ElpKMSrc_Id = meElpKM_Classe.K2
meElpKMInfo.Id = X
meElpKMInfo.Description = Trim(txtSelect)
meElpKMInfo.Pass = mElpKM_Classe_Pass
'meElpKMInfo.Memo=""
''meElpKMInfo.Method = "Seek="
''If tableElpKMInfo_Read(meElpKMInfo) = 0 Then Call lstErr_AddItem(lstErr, cmdContext, "?la référence existe déjà ")

''If lstErr.ListCount = 0 Then cmdElpKMInfo_AddNew_Ok

End Sub



Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

Private Sub cmdSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdSelect

End Sub

Private Sub cmdZMNUMEN0_Scan_Click()
Dim X As String, I As Integer, I0 As Integer, K As Integer
Dim wX As String
Dim xText As String

On Error Resume Next
lstErr.Clear
X = Text_LCase(txtZMNUMEN0_Scan)
If X = "" Then
    Call lstErr_AddItem(lstErr, cmdContext, "?  Précisez la recherche ")
    txtZMNUMEN0_Scan.SetFocus
Else
    fgZMNUMEN0.Visible = False
    I0 = fgZMNUMEN0.Row + 1
    For I = I0 To fgZMNUMEN0.Rows - 1
        fgZMNUMEN0.Row = I:
        fgZMNUMEN0.Col = 0: xText = fgZMNUMEN0.Text
        fgZMNUMEN0.Col = 1
        wX = Text_LCase(xText & " " & fgZMNUMEN0.Text)
        K = InStr(wX, X)
        If K > 0 Then
            Call fgZMNUMEN0_Color(fgZMNUMEN0_RowClick, MouseMoveUsr.BackColor, fgZMNUMEN0_ColorClick)
            fgZMNUMEN0.TopRow = I
            Exit For
        End If
    Next I
    fgZMNUMEN0.Visible = True
End If


End Sub

Private Sub lstElpKM_Classe_Click()
Dim X As String, I As Integer
I = Len(lstElpKM_Classe.Text)
X = Mid$(lstElpKM_Classe.Text, I - 11, 12)
meElpKM_Classe.K2 = Val(X)
If meElpKM_Classe.K2 = 11000 Then
    fraSelect1100.Visible = True
Else
    fraSelect1100.Visible = False
End If
optSelect11000_All = True
fctElpKM_Classe_Load

End Sub

Private Sub lstElpKM_Classe_GotFocus()
lstElpKM_Classe.BackColor = vbCyan
lstElpKM_Classe.BackColor = txtUsr.BackColor

End Sub


Private Sub lstUsr_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
lstUsr_Select
'Me.PopupMenu mnulstUsr, vbPopupMenuLeftButton
blnZMNUMEN0_Export = False
fgZMNUMEN0_Load_Aut

End Sub


Private Sub fgZMNUMEN0_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If fgZMNUMEN0.Rows > 1 Then
    Call fgZMNUMEN0_Color(fgZMNUMEN0_RowClick, MouseMoveUsr.BackColor, fgZMNUMEN0_ColorClick)
    fgZMNUMEN0.Col = fgZMNUMEN0_arrIndex:  arrZMNUMEN0_Index = CLng(fgZMNUMEN0.Text)
    meZMNUMEN0 = arrZMNUMEN0(arrZMNUMEN0_Index)
    meZMNUOPT0 = arrZMNUOPT0(arrZMNUMEN0_Index)
    fgZMNUMEN0.Col = 0: fgZMNUMEN0.LeftCol = 0
    
    Me.PopupMenu mnufgZMNUMEN0, vbPopupMenuLeftButton

End If

End Sub

Private Sub mnuExport_Options_Click()
Dim wFileName As String, wFreeFile As Integer
Dim V, wText As String, I As Integer, X As String, K As Integer, mX As String
Dim xSQL As String, xWhere As String
Dim wOption As String
Dim wlong As Long
Dim wUser As String


Me.Enabled = False: Me.MousePointer = vbHourglass
wUser = "G_ADMIN"
Call lstErr_AddItem(lstErr, cmdContext, "mnuExport_Options :  load " & wUser)
Call lst_Scan(wUser, lstUsr)
lstUsr_Select
blnZMNUMEN0_Export = False
fgZMNUMEN0_Load_Aut
Call lstErr_AddItem(lstErr, cmdContext, "mnuExport_Options :  NB " & fgZMNUMEN0.Rows)
'____________________________________________________________________ autres options
Call lstErr_AddItem(lstErr, cmdContext, "mnuExport_Options :  autres options")
fgZMNUMEN0.Visible = False
xSQL = "select count(*)  as Tally   from " & paramIBM_Library_SAB & ".ZMNUOPT0"
Set rsSab = cnsab.Execute(xSQL)

ReDim arrZMNUOPT0(rsSab("Tally") + 1)
xSQL = "select MNUOPTCOD, MNUOPTLIB from " & paramIBM_Library_SAB & ".ZMNUOPT0 order by MNUOPTCOD"
Set rsSab = cnsab.Execute(xSQL)
arrZMNUOPT0_Nb = 0
Do While Not rsSab.EOF
    arrZMNUOPT0_Nb = arrZMNUOPT0_Nb + 1
    arrZMNUOPT0(arrZMNUOPT0_Nb).MNUOPTCOD = rsSab("MNUOPTCOD")
    arrZMNUOPT0(arrZMNUOPT0_Nb).MNUOPTLIB = rsSab("MNUOPTLIB")
    rsSab.MoveNext
Loop

xSQL = "select MNUMENCOD from " & paramIBM_Library_SAB & ".ZMNUMEN0 " _
     & " where MNUMENGRP = '" & wUser & "'" _
     & " and   MNUMENREF =" & currentZMNUHLB0.MNUHLBREF _
     & " and   MNUMENETB =" & currentZMNUHLB0.MNUHLBETB _
     & " order by MNUMENCOD"
Set rsSab = cnsab.Execute(xSQL)

K = 1
Do While Not rsSab.EOF
    wlong = rsSab("MNUMENCOD")
    Do
        If arrZMNUOPT0(K).MNUOPTCOD = wlong Then
            arrZMNUOPT0(K).MNUOPTCOD = 0
            K = K + 1
            Exit Do
        Else
            If arrZMNUOPT0(K).MNUOPTCOD < wlong Then
                K = K + 1
            Else
                Exit Do
            End If
        End If
    Loop
    rsSab.MoveNext
Loop
For K = 1 To arrZMNUOPT0_Nb
    If arrZMNUOPT0(K).MNUOPTCOD <> 0 Then
        fgZMNUMEN0.Rows = fgZMNUMEN0.Rows + 1
        fgZMNUMEN0.Row = fgZMNUMEN0.Rows - 1
        fgZMNUMEN0.Col = 0: fgZMNUMEN0.Text = arrZMNUOPT0(K).MNUOPTCOD
        fgZMNUMEN0.Col = 1: fgZMNUMEN0.Text = "* " & arrZMNUOPT0(K).MNUOPTLIB
    End If
Next K
fgZMNUMEN0.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "mnuExport_Options :  NB " & fgZMNUMEN0.Rows)
'____________________________________________________________________


ZMNUHLB0_Load
MsgBox "c:\Temp\SAB_MNU.csv", vbInformation, "Export Options / Groupes"
wFileName = "c:\Temp\SAB_MNU.csv"
V = File_Export_Monitor("Output", wFreeFile, wFileName)
If Not IsNull(V) Then GoTo Exit_sub
wText = "Option;Libellé"
For K = 1 To arrZMNUHLB0_Nb
    wText = wText & ";" & arrZMNUHLB0(K).MNUHLBNOM
Next K
V = File_Export_Monitor("Print", wFreeFile, wText)
Call lstErr_AddItem(lstErr, cmdContext, "mnuExport_Options :  Groupes : " & arrZMNUHLB0_Nb)

For I = 1 To fgZMNUMEN0.Rows - 1
    fgZMNUMEN0.Row = I
    fgZMNUMEN0.Col = 0: wOption = fgZMNUMEN0.Text
    fgZMNUMEN0.Col = 1: wText = wOption & ";" & fgZMNUMEN0.Text
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "mnuExport_Options :  " & wOption)
    For K = 1 To arrZMNUHLB0_Nb
        arrZMNUHLB0(K).MNUHLBVAL = " "
    Next K
    xWhere = " where MNUMENCOD = " & wOption _
         & " and   MNUMENREF =" & currentZMNUHLB0.MNUHLBREF _
         & " and   MNUMENETB =" & currentZMNUHLB0.MNUHLBETB

    xSQL = "select MNUMENGRP from " & paramIBM_Library_SAB & ".ZMNUMEN0" & xWhere & " order by MNUMENGRP"
    Set rsSab = cnsab.Execute(xSQL)
    K = 0: mX = ""
    Do While Not rsSab.EOF
        X = rsSab("MNUMENGRP")
        If mX <> X Then
            mX = X
            Do
                K = K + 1
                If arrZMNUHLB0(K).MNUHLBNOM = X Then arrZMNUHLB0(K).MNUHLBVAL = "*": Exit Do
                If arrZMNUHLB0(K).MNUHLBNOM > X Then K = K - 1: Exit Do
           Loop Until K = arrZMNUHLB0_Nb
        End If
        rsSab.MoveNext
    Loop
    For K = 1 To arrZMNUHLB0_Nb
        wText = wText & ";" & arrZMNUHLB0(K).MNUHLBVAL
    Next K
    V = File_Export_Monitor("Print", wFreeFile, wText)

Next I
V = File_Export_Monitor("Close", wFreeFile, wText)

Exit_sub:
Call lstErr_AddItem(lstErr, cmdContext, "mnuExport_Options :  fin")
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnufgSelect_Print_Click()
Dim colX(5) As Integer
Dim wText As String, I0 As Integer, nbTab As Integer

I0 = InStr(1, lstElpKM_Classe, Chr$(9))
prtTitleText = Trim(Mid$(lstElpKM_Classe, 1, I0 - 1)) & " : " & Trim(txtSelect) & "*"
prtFontName = prtFontName_CourierNew
prtOrientation = vbPRORLandscape
prtPgmName = "prtElpKM"
prtTitleUsr = usrName

prtElpKM.prtStd_Open

colX(0) = prtMinX
colX(1) = prtMinX
colX(2) = prtMinX
colX(3) = prtMinX + 2000

prtElpKM_fgSelect_Form colX()
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
XPrt.FontSize = 7
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    If XPrt.CurrentY + 100 > prtMaxY Then
        frmElpPrt.prtNewPage
        prtElpKM_fgSelect_Form colX()
    End If
    
    fgSelect.Col = 2: XPrt.CurrentX = colX(2): XPrt.Print fgSelect.Text;
    XPrt.CurrentX = colX(3)
    
    If meElpKM_Classe.K2 = 30000 Then
 
    Else
        If meElpKM_Classe.K2 = 12000 Then
            fgSelect.Col = 1:  XPrt.Print Mid$(fgSelect.Text, 30, 10) & " ";
        End If
        fgSelect.Col = 3:  XPrt.Print fgSelect.Text;
    End If
Next I

prtElpKM.prtStd_Close

End Sub

Private Sub mnulstUsr_Aut_Click()
blnZMNUMEN0_Export = False
fgZMNUMEN0_Load_Aut
End Sub

Private Sub mnulstUsr_Aut_Export_Click()
Dim I As Integer, K As Integer, kCUT As Integer
Dim lenX As Integer
blnZMNUMEN0_Export = True
fgZMNUMEN0_Load_Aut
End Sub


Private Sub mnufgZMNUMEN0_Print_Click()
cmdPrint_fgZMNUMEN0 "M"

End Sub

Private Sub mnuPrint3_GroupUser_Click()
Dim xUTI As String * 10, xGR2 As String * 10
Dim xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass

lstW.Clear
xSQL = "select MNURUTUTI, MNUUTIGR3 from " & paramIBM_Library_SAB & ".ZMNURUT0," & paramIBM_Library_SAB & ". ZMNUUTI0" _
     & " where MNURUTLOG = 'O' and MNURUTCUT = MNUUTICUT"
     
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    xUTI = rsSab("MNURUTUTI")
    xGR2 = rsSab("MNUUTIGR3")
    lstW.AddItem xGR2 & xUTI
    rsSab.MoveNext
Loop
prtElpKm_UserGroup lstW, 2
Me.Enabled = True: Me.MousePointer = 0
End Sub


Private Sub mnuPrint3_UserGroup_Click()
Dim xUTI As String * 10, xGR2 As String * 10
Dim xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass

lstW.Clear
xSQL = "select MNURUTUTI, MNUUTIGR2 from " & paramIBM_Library_SAB & ".ZMNURUT0," & paramIBM_Library_SAB & ". ZMNUUTI0" _
     & " where MNURUTLOG = 'O' and MNURUTCUT = MNUUTICUT"
     
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    xUTI = rsSab("MNURUTUTI")
    xGR2 = rsSab("MNUUTIGR2")
    lstW.AddItem xUTI & xGR2
    rsSab.MoveNext
Loop
prtElpKm_UserGroup lstW, 1
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuPrint3_xls_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String

X = "SAB_MNU : " & Mid$(lstUsr.Text, 1, 10)
Call MSflexGrid_Excel("", "SAB_MNU", X, fgZMNUMEN0, 2)
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuPrint3All_Click()
cmdPrint_fgZMNUMEN0 ""
End Sub

Private Sub mnuPrint3Menu_Click()
cmdPrint_fgZMNUMEN0 "M"

End Sub



Private Sub mnuSelect11000_DisplayFields_Click()
mnuSelect11000_DisplayFields_Exe False, False ' blnAlpha,blnPrint
End Sub

Private Sub mnuSelect11000_DisplayFields_Exe(blnAlpha As Boolean, blnPrint As Boolean)
Dim wElpKMSrc_Id As Long
Me.Enabled = False
lstErr.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "DSPFFD : " & Time)

'wElpKMSrc_Id = 12000 + srvDSPFFDY0_Library(meDSPFDY0.ATLIB)

'Call srvDSPFFDY0_lstAddItem(lstW, wElpKMSrc_Id, meDSPFDY0.ATFILE, False)

'Call srvDSPFFDY0_frmRTF(lstW, wElpKMSrc_Id, meDSPFDY0.ATLIB, meDSPFDY0.ATFILE, blnPrint)

Call lstErr_AddItem(lstErr, cmdContext, "DSPFFD :  Fin")
Me.Enabled = True

End Sub

Private Sub mnuSelect11000_PrintFields_All_Click()
Dim I As Integer
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex: arrElpKM_Index = Val(fgSelect.Text)
    xElpKMInfo.Id = arrElpKMInfo_Id(arrElpKM_Index)
    xElpKMInfo.ElpKMSrc_Id = arrElpKMSrc_Id(arrElpKM_Index)
  '  intReturn = tableElpKMInfo_Read(xElpKMInfo)

   ' MsgTxt = Space$(34) & xElpKMInfo.Memo
   ' MsgTxtIndex = 0
    'srvDSPFDY0_GetBuffer meDSPFDY0

    mnuSelect11000_DisplayFields_Exe False, True ' blnAlpha,blnPrint
    
Next I
End Sub

Private Sub mnuSelect11000_PrintFields_Click()
mnuSelect11000_DisplayFields_Exe False, True ' blnAlpha,blnPrint

End Sub

Private Sub mnuSelect12000_DisplayFields_Click()
Dim wElpKMSrc_Id As Long
Me.Enabled = False
lstErr.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "DSPFFD : " & Time)

'wElpKMSrc_Id = 12000 + srvDSPFFDY0_Library(meDSPFFDY0.WHLIB)

'Call srvDSPFFDY0_lstAddItem(lstW, wElpKMSrc_Id, meDSPFFDY0.WHFILE, False)

'Call srvDSPFFDY0_frmRTF(lstW, wElpKMSrc_Id, meDSPFFDY0.WHLIB, meDSPFFDY0.WHFILE, False)


Call lstErr_AddItem(lstErr, cmdContext, "DSPFFD :  Fin")
Me.Enabled = True

End Sub

Private Sub mnuSelectUsr_Click()
Dim I As Integer, lenX As Integer, K As Integer
Dim blnOk As Boolean
Dim mTab As String

' meZMNUMEN0.MNUMENCOD = code option recherchée
'-------------------------------------------------
lstSelectUsr.Clear
lstSelectUsr.Visible = True
blnOk = False
For I = 1 To arrZMNUMEN0_Nb
         
    If meZMNUMEN0.MNUMENCOD = arrZMNUMEN0(I).MNUMENCOD Then
        meZMNUMEN0 = arrZMNUMEN0(I)
        blnOk = True
        Exit For
    End If
Next I
mTab = ""
If Not blnOk Then
    lstSelectUsr.AddItem "Option non autorisée pour l'utilisateur"
Else
    For I = 1 To Len(meZMNUMEN0.Hierarchie) Step 5
        K = Val(Mid$(meZMNUMEN0.Hierarchie, I, 5))
        xZMNUMEN0 = arrZMNUMEN0(K)
        mTab = mTab & vbTab
        lstSelectUsr.AddItem arrZMNUMEN0(K).MNUMENCOD & mTab & arrZMNUOPT0(K).MNUOPTLIB
    Next I
End If
End Sub

Private Sub mnuSelectUsrAll_Click()
Dim xSQL As String
lstSelectUsr.Visible = True
lstSelectUsr.Clear

xSQL = "select * from " & paramIBM_Library_SAB & ".ZMNUMEN0," & paramIBM_Library_SAB & ". ZMNUHLB0" _
     & " where MNUMENCOD = " & Val(mElpKMInfo_Id) _
     & " and MNUMENREF = MNUHLBREF " _
     & " and MNUMENGRP = MNUHLBNOM" _
     & " and MNUHLBCLA = '2' and MNUHLBVAL = '1' and MNUHLBFID = 0"

Set rsSab = cnsab.Execute(xSQL)
Do Until rsSab.EOF
    lstSelectUsr.AddItem rsSab("MNUMENREF") & " " & rsSab("MNUMENGRP")
    rsSab.MoveNext
Loop

End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
lstUsr.BackColor = &HE0E0E0
lstSelectUsr = &HE0E0E0
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
blncmdOk_Visible = False: blncmdSave_Visible = False
currentAction = ""


libRéférenceInterne = libMNURUT
'libMNURUT = ""
fgSelect_Reset

cmdElpKMInfo_AddNew.Visible = ElpKMAut.Saisir

'========================RTF

SSTab1.Tab = 0

blnControl = True
End Sub


Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
On Error Resume Next
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
    fgSelect.Row = lRow
    For I = 0 To fgSelect_arrIndex
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = 0 To fgSelect_arrIndex
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.Col = 0
    End If
End If

End Sub

Private Sub fgSelect_Display()
Dim xSQL As String
SSTab1.Tab = 1

fgSelect.Visible = True
fgSelect.Clear: fgSelect.Rows = 1: fgSelect_RowDisplay = 0

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Enabled = True


For arrElpKM_Index = 1 To arrElpKM_Nb
    If arrElpKM_Occurs(0) <= arrElpKM_Occurs(arrElpKM_Index) Then
        xSQL = "select * from ElpKmInfo" _
            & " where ElpKmSrc_Id = " & arrElpKMSrc_Id(arrElpKM_Index) _
            & " and ID ='" & arrElpKMInfo_Id(arrElpKM_Index) & "'"
    
        Set rsMDB = cnMDB.Execute(xSQL)
        If Not rsMDB.EOF Then
            Call rsElpKmInfo_GetBuffer(rsMDB, xElpKMInfo)
   
            If xElpKMInfo.ElpKMSrc_Id = 11000 Then
                If optSelect11000_PF Then
                 '   MsgTxt = Space$(34) & xElpKMInfo.Memo
                 '   MsgTxtIndex = 0
                 '   srvDSPFDY0_GetBuffer meDSPFDY0
                 '   If meDSPFDY0.ATFATR <> "PF    " Then intReturn = -1
                    
                End If
                If optSelect11000_LF Then
                  '  MsgTxt = Space$(34) & xElpKMInfo.Memo
                  '  MsgTxtIndex = 0
                  '  srvDSPFDY0_GetBuffer meDSPFDY0
                  '  If meDSPFDY0.ATFATR <> "LF    " Then intReturn = -1
                    
                End If
                    
            End If
        
                fgSelect.Rows = fgSelect.Rows + 1
                fgSelect.Row = fgSelect.Rows - 1
                fgSelect_DisplayLine
        End If
    End If
Next arrElpKM_Index

fgSelect_SortAD = 5
If fgSelect.Rows > 1 Then fgSelect_Sort



End Sub
Public Sub fgSelect_DisplayLine()
On Error Resume Next
fgSelect.Col = 0:
If xElpKMInfo.ElpKMSrc_Id = 1000 Then
    fgSelect.Text = Mid$(xElpKMInfo.Memo, 51, 8) & "  " & xElpKMInfo.Description
Else
    fgSelect.Text = xElpKMInfo.Description
End If

fgSelect.Col = 1: fgSelect.Text = mElpKM_Classe_Abrégé  'xElpKMInfo.ElpKMSrc_Id
fgSelect.Col = 2: fgSelect.Text = xElpKMInfo.Id
fgSelect.Col = 3:
If Not IsNull(xElpKMInfo.Memo) Then
    Select Case xElpKMInfo.ElpKMSrc_Id
        Case 11000: fgSelect.Text = Mid$(xElpKMInfo.Memo, 79, 50) & " " & Mid$(xElpKMInfo.Memo, 1, 78) & " " & Mid$(xElpKMInfo.Memo, 129, 84)
        Case Is >= 12000 <= 12999: fgSelect.Text = Mid$(xElpKMInfo.Memo, 177, 50) & " " & Mid$(xElpKMInfo.Memo, 1, 176) & " " & Mid$(xElpKMInfo.Memo, 231, 599)

        Case Else: fgSelect.Text = xElpKMInfo.Memo
    End Select
End If

fgSelect.Col = 4: fgSelect.Text = xElpKMInfo.ElpKMSrc_Id

fgSelect.Col = fgSelect_arrIndex - 1: fgSelect.Text = ""
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = arrElpKM_Index

End Sub
Public Sub fgSelect_Load()
Dim X As String, lX As Integer, intReturn As Integer, I As Integer
Dim wInfo_Id As String * 20, wInfo_K As Integer, wInfo_X As String
Dim blnOk As Boolean, blnInit As Boolean
Dim kIn As Integer, xIn As String
Dim xSQL As String

ReDim arrElpKM_Classe(10000)
ReDim arrElpKMSrc_Id(10000)
ReDim arrElpKMInfo_Id(10000)
ReDim arrElpKM_Occurs(10000)

arrElpKM_Nb = 0
libRéférenceInterne = libMNURUT
blnInit = True
arrElpKM_Occurs(0) = 0
xIn = Text_LCase(txtSelect)
kIn = 0
Do
    X = Text_KeyWord(xIn, kIn, True)
    If X <> "" Then
        arrElpKM_Occurs(0) = arrElpKM_Occurs(0) + 1
        xSQL = "select * from ElpKmIndex" _
            & " where Id like '%" & X & "%'" _
            & " and Classe = " & meElpKM_Classe.K2
    
        Set rsMDB = cnMDB.Execute(xSQL)
        Do Until rsMDB.EOF

            Call rsElpKmIndex_GetBuffer(rsMDB, xElpKMIndex)
            wInfo_X = xElpKMIndex.Memo
            For wInfo_K = 1 To Len(wInfo_X) Step 20
                wInfo_Id = Mid$(wInfo_X, wInfo_K, 20)
                blnOk = False
                For I = 1 To arrElpKM_Nb
                    If arrElpKM_Classe(I) = xElpKMIndex.Classe And arrElpKMSrc_Id(I) = xElpKMIndex.ElpKMSrc_Id And arrElpKMInfo_Id(I) = wInfo_Id Then
                        blnOk = True
                        If Not blnInit Then arrElpKM_Occurs(I) = arrElpKM_Occurs(I) + 1
                     End If
                Next I
                If blnInit And Not blnOk Then
                    arrElpKM_Nb = arrElpKM_Nb + 1
                    arrElpKM_Classe(arrElpKM_Nb) = xElpKMIndex.Classe
                    arrElpKMSrc_Id(arrElpKM_Nb) = xElpKMIndex.ElpKMSrc_Id
                    arrElpKMInfo_Id(arrElpKM_Nb) = wInfo_Id
                    arrElpKM_Occurs(arrElpKM_Nb) = 1
                End If
            Next wInfo_K
            rsMDB.MoveNext
        Loop
    End If
    
    blnInit = False
Loop Until X = ""

fgSelect_Display
End Sub
Public Sub fgSelect_Sort()
If fgSelect.Rows > 1 Then
    fgSelect.Row = 1
    fgSelect.RowSel = fgSelect.Rows - 1
    
    If fgSelect_Sort1_Old = fgSelect_Sort1 Then
        If fgSelect_SortAD = 5 Then
            fgSelect_SortAD = 6
        Else
            fgSelect_SortAD = 5
        End If
    Else
        fgSelect_SortAD = 5
    End If
    fgSelect_Sort1_Old = fgSelect_Sort1
    
    fgSelect.Col = fgSelect_Sort1
    fgSelect.ColSel = fgSelect_Sort2
    fgSelect.Sort = fgSelect_SortAD
End If

End Sub
Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
 '   meElpKMInfo_Index = Val(fgSelect.Text)
 '   fgSelect.Col = fgSelect_arrIndex - 1
 '  x = meElpKMInfo(meElpKMInfo_Index).SCCOMPTE & meElpKMInfo(meElpKMInfo_Index).SCDEVISE
 '   Select Case lK
 '       Case 0: fgSelect.Text = meElpKMInfo(meElpKMInfo_Index).SCSTATUS & X
 '       Case fgSelect_arrIndex: fgSelect.Text = Format$(meElpKMInfo_Index, "0000000000")
 '   End Select
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub



Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

SSTab1.Tab = 0
'ReDim meElpKMInfo(10)
param_Init

blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fgZMNUMEN0_FormatString = fgZMNUMEN0.FormatString

ReDim arrElpKM_Classe(1000)
ReDim arrElpKMSrc_Id(1000)
ReDim arrElpKMInfo_Id(1000)
ReDim arrElpKM_Occurs(1000)


cmdReset

lstUsr_Load

End Sub

Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
         If TypeOf xobj Is CommandButton Or TypeOf xobj Is ListBox Then
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
        If TypeOf C Is CommandButton Or TypeOf C Is ListBox Then
            
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            MouseMoveActiveControl.ForeColor = C.ForeColor
            C.ForeColor = MouseMoveUsr.ForeColor
        End If
    End If
End If

End Sub

'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
currentActiveControl_Name = C.Name
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
End Sub


'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Dim Msg As String
Dim I As Integer

Me.Enabled = False

Msg = Space$(50)
Select Case SSTab1.Tab
    Case 1: Me.PopupMenu mnuPrint1, vbPopupMenuLeftButton
    Case 3: Me.PopupMenu mnuPrint3, vbPopupMenuLeftButton
End Select

Me.Enabled = True

End Sub

Private Sub cmdSelect_Click()
cmdSelect_Control
If lstErr.ListCount = 0 Then fgSelect_Load
End Sub

'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False
fgSelect.Clear: fgSelect.Row = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wRow As Long

lstSelectUsr.Visible = False
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1: fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If Button = vbRightButton Then
        wRow = Fix(y / fgSelect.RowHeightMin)
        If wRow <= fgSelect.Rows Then fgSelect.Row = wRow
    End If
    fgSelect_MouseDown_Ok
    fgSelect.Col = fgSelect_arrIndex:  arrElpKM_Index = Val(fgSelect.Text)
        Select Case mElpKM_Classe
            Case 1000:
                    'arrElpKM_Classe (arrElpKM_Index)
                    If arrElpKMSrc_Id(arrElpKM_Index) = 1000 Then
                            meZMNUMEN0.MNUMENCOD = Val(arrElpKMInfo_Id(arrElpKM_Index))
                            Me.PopupMenu mnuSelect1000, vbPopupMenuLeftButton
                    End If
            Case 11000:
                            xElpKMInfo.Id = arrElpKMInfo_Id(arrElpKM_Index)
                            xElpKMInfo.ElpKMSrc_Id = arrElpKMSrc_Id(arrElpKM_Index)
                           ' intReturn = tableElpKMInfo_Read(xElpKMInfo)
        
                           ' MsgTxt = Space$(34) & xElpKMInfo.Memo
                            'MsgTxtIndex = 0
                            'srvDSPFDY0_GetBuffer meDSPFDY0
                            'Me.PopupMenu mnuSelect11000, vbPopupMenuLeftButton
            Case 12000:
                            xElpKMInfo.Id = arrElpKMInfo_Id(arrElpKM_Index)
                            xElpKMInfo.ElpKMSrc_Id = arrElpKMSrc_Id(arrElpKM_Index)
                            'intReturn = tableElpKMInfo_Read(xElpKMInfo)
        
                            'MsgTxt = Space$(34) & xElpKMInfo.Memo
                            'MsgTxtIndex = 0
                            'srvDSPFFDY0_GetBuffer meDSPFFDY0
                            'Me.PopupMenu mnuSelect12000, vbPopupMenuLeftButton
                            

    End Select
   ' End If
    
    
End If

End Sub
Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub

Private Sub mnuContextQuitter_Click()
Unload Me
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), ElpKMAut)

If Mid$(Elp.SrvDtaqIn, 1, 2) = "PC" Then
    If Mid$(Msg, 1, 12) = "X_DOC$      " And ElpKMAut.Saisir Then
            If DataBase_Open <> DataBase_Master Then MDB_Open DataBase_Master, paramDataBase_Password
    End If
End If

blnSetfocus = True
Form_Init


End Sub


Public Sub cmdContext_Return()
If SSTab1.Tab = 3 Then
    cmdZMNUMEN0_Scan_Click
Else
    If currentAction = "" Then
        If SSTab1.Tab > 0 Then
            SSTab1.Tab = 0
        Else
           'SendKeys "{TAB}"
            cmdSelect_Click
        End If
    End If
End If
End Sub


Public Sub fgSelect_Reset()
fgSelect_Sort1 = 1: fgSelect_Sort2 = 2
fgSelect_Sort1_Old = 0
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 6
blnfgSelect_DisplayLine = False
lstSelectUsr.Visible = False
End Sub


Private Sub optSAB_MNU_L1_Click()
fgZMNUMEN0_Display

End Sub

Private Sub optSAB_MNU_L2_Click()
fgZMNUMEN0_Display

End Sub

Private Sub optSAB_MNU_L3_Click()
fgZMNUMEN0_Display

End Sub

Private Sub optSAB_MNU_L4_Click()
fgZMNUMEN0_Display

End Sub





Private Sub txtSelect_GotFocus()
txt_GotFocus txtSelect

End Sub


Private Sub txtSelect_LostFocus()
txt_LostFocus txtSelect

End Sub




Public Sub param_Init()
Dim mK2 As String
Dim X As String

lstElpKM_Classe.Clear
lstElpKM_Classe_listIndex = -1
X = "select * from ElpTable where SNN = 0" _
    & " and id = 'ElpKM' and K1 = 'Info'"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    mK2 = rsMDB("K2")
    lstElpKM_Classe.AddItem rsMDB("Name") & vbTab & mK2
    If Trim(mK2) = "1000" Then lstElpKM_Classe_listIndex = lstElpKM_Classe.ListCount - 1 ' option menu SAB
    rsMDB.MoveNext
Loop


If lstElpKM_Classe_listIndex <> -1 Then
    lstElpKM_Classe.ListIndex = lstElpKM_Classe_listIndex
    lstElpKM_Classe_Click
End If

End Sub

Public Sub fctElpKM_Classe_Load()
Dim blnOk As Boolean
Dim I1 As Integer, I2 As Integer, I3 As Integer
Dim X As String, xName As String, xMemo As String

blnOk = True
cmdSelect.Enabled = False
mElpKM_Classe_Pass = -1
mElpKM_Classe_Abrégé = "?"
mElpKM_Classe_Folder = "?"
V = rsElpTable_Read("ElpKM", "Info", meElpKM_Classe.K2, xName, xMemo)
If Not IsNull(V) Then
    blnOk = False: Exit Sub
Else
    If IsNull(xMemo) Then
        MsgBox meElpKM_Classe.Name, vbCritical, "manque Memo":
        blnOk = False: Exit Sub
    Else
        I3 = Len(xMemo)
        I1 = 1
        I2 = InStr(I1, xMemo, Asc34)
        If I2 <= 0 Then MsgBox xMemo, vbCritical, "manque abrégé début": blnOk = False: Exit Sub
        If I2 <= I3 Then mElpKM_Classe_Pass = CLng(Mid$(xMemo, I1, I2 - 1))
        
        I1 = I2 + 1
        I2 = InStr(I1, xMemo, Asc34)
        If I2 <= 0 Then MsgBox xMemo, vbCritical, "manque abrégé fin": blnOk = False: Exit Sub
        If I2 <= I3 Then mElpKM_Classe_Abrégé = Mid$(xMemo, I1, I2 - I1)
    
        I1 = I2 + 1
        I2 = InStr(I1, xMemo, Asc34)
        If I2 <= 0 Then MsgBox xMemo, vbCritical, "manque répertoire début": blnOk = False: Exit Sub
        I1 = I2 + 1
        I2 = InStr(I1, xMemo, Asc34)
        If I2 <= 0 Then MsgBox xMemo, vbCritical, "manque répertoire fin": blnOk = False: Exit Sub
        If I2 <= I3 Then X = Mid$(xMemo, I1, I2 - I1)
    End If
End If


mElpKM_Classe_Folder = paramServer(X) & mElpKM_Classe_Abrégé & "\"
If blnOff_Line Then mElpKM_Classe_Folder = "C:\Temp\Sab\" & mElpKM_Classe_Abrégé & "\"
cmdSelect.Enabled = blnOk
End Sub

Public Sub lstUsr_Load()
Dim xMNURUTUTI As String
Dim wIndex As Integer, K As Integer
Dim xSQL As String
Dim X10 As String * 10

On Error Resume Next

wIndex = 0
lstUsr.Clear

xSQL = "select MNURUTUTI, MNURUTNOM from " & paramIBM_Library_SAB & ".ZMNURUT0"
     
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    xMNURUTUTI = rsSab("MNURUTUTI")
    X10 = xMNURUTUTI
    lstUsr.AddItem X10 & vbTab & rsSab("MNURUTNOM")
    ''If xMNURUTUTI = usrId Then wIndex = lstUsr.ListCount
    rsSab.MoveNext
Loop
Call lst_Scan(usrId, lstUsr)
lstUsr_Select
blnZMNUMEN0_Export = False
fgZMNUMEN0_Load_Aut

End Sub

Public Sub fgZMNUMEN0_Load_Aut()
Dim xSQL As String
On Error Resume Next
fgZMNUMEN0_Reset

xSQL = "select * from " & paramIBM_Library_SAB & ".ZMNUHLB0" _
     & " where MNUHLBNOM = '" & meZMNUUTI0.MNUUTIGR2 & "'" _
     & " and MNUHLBCLA = '2' and MNUHLBVAL = '1' and MNUHLBFID = 0" _
     & " and MNUHLBETB = " & meZMNURUT0.MNURUTETB



Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    Call rsZMNUHLB0_GetBuffer(rsSab, meZMNUHLB0)
     '   fraSelect.Caption = meZMNUHLB0.MNUHLBCLA & " : " & meZMNUHLB0.MNUHLBNOM
    fgZMNUMEN0_Load

End If

End Sub

Public Sub picUsr_Display()
picUsr.Cls

picUsr.ForeColor = lblUsr.ForeColor
picUsr.FontBold = False
picUsr.FontSize = 6
picUsr.CurrentX = 50: picUsr.CurrentY = 300: picUsr.Print "Nom";: picUsr.CurrentX = 1200: picUsr.Print ":";
picUsr.CurrentX = 50: picUsr.CurrentY = 50: picUsr.Print "Code";: picUsr.CurrentX = 1200: picUsr.Print ":";
picUsr.CurrentX = 50: picUsr.CurrentY = 550: picUsr.Print "Ag-Srv";: picUsr.CurrentX = 1200: picUsr.Print ":";
picUsr.CurrentX = 50: picUsr.CurrentY = 800: picUsr.Print "Groupe Menu";: picUsr.CurrentX = 1200: picUsr.Print ":";
picUsr.CurrentX = 50: picUsr.CurrentY = 1050: picUsr.Print "Groupe Droits";: picUsr.CurrentX = 1200: picUsr.Print ":";
picUsr.CurrentX = 50: picUsr.CurrentY = 1300: picUsr.Print "Groupe Métier";: picUsr.CurrentX = 1200: picUsr.Print ":";


picUsr.FontBold = True
picUsr.ForeColor = warnUsrColor

picUsr.CurrentX = 500: picUsr.CurrentY = 50: picUsr.Print Val(meZMNURUT0.MNURUTCUT);
picUsr.ForeColor = libUsr.ForeColor
picUsr.CurrentX = 1300: picUsr.Print meZMNURUT0.MNURUTUTI;
picUsr.CurrentX = 1300: picUsr.CurrentY = 300: picUsr.Print meZMNURUT0.MNURUTNOM;
picUsr.CurrentX = 1300: picUsr.CurrentY = 550: picUsr.Print meZMNUUTI0.MNUUTIAGE & " - " & meZMNUUTI0.MNUUTISER & " - " & meZMNUUTI0.MNUUTISRV;
picUsr.ForeColor = warnUsrColor
picUsr.CurrentX = 1300: picUsr.CurrentY = 800: picUsr.Print meZMNUUTI0.MNUUTIGR2;
picUsr.CurrentX = 1300: picUsr.CurrentY = 1050: picUsr.Print meZMNUUTI0.MNUUTIGR3;
picUsr.CurrentX = 1300: picUsr.CurrentY = 1300: picUsr.Print meZMNUUTI0.MNUUTIGR4;
picUsr.ForeColor = libUsr.ForeColor
'picUsr.Print " " & Mid$(xSAB_ZMNU.Memo, 1, 10);
picUsr.ForeColor = lblUsr.ForeColor
picUsr.FontBold = False
picUsr.CurrentX = 3500: picUsr.CurrentY = 550: picUsr.Print "entrée logiciel";: picUsr.CurrentX = 5200: picUsr.Print ":";
picUsr.CurrentX = 3500: picUsr.CurrentY = 800: picUsr.Print "file d'attente";: picUsr.CurrentX = 5200: picUsr.Print ":";
picUsr.CurrentX = 3500: picUsr.CurrentY = 1050: picUsr.Print "menu service";: picUsr.CurrentX = 5200: picUsr.Print ":";
picUsr.CurrentX = 3500: picUsr.CurrentY = 1300: picUsr.Print "groupe menu service";: picUsr.CurrentX = 5200: picUsr.Print ":";

picUsr.FontBold = True
picUsr.ForeColor = libUsr.ForeColor

picUsr.CurrentX = 5300: picUsr.CurrentY = 800: picUsr.Print meZMNUUTI0.MNUUTIOUT;

picUsr.CurrentX = 5300: picUsr.CurrentY = 550:
If meZMNURUT0.MNURUTLOG = "N" Then
    picUsr.ForeColor = vbRed: picUsr.Print "Non";
Else
    picUsr.ForeColor = libUsr.ForeColor: picUsr.Print "Oui";
End If

picUsr.CurrentX = 5300: picUsr.CurrentY = 1050:
If meZMNUUTI0.MNUUTIMSE = "N" Then
    picUsr.ForeColor = vbRed: picUsr.Print "Non";
Else
    picUsr.ForeColor = libUsr.ForeColor: picUsr.Print "Oui";
End If

picUsr.CurrentX = 5300: picUsr.CurrentY = 1300:
picUsr.ForeColor = libUsr.ForeColor: picUsr.Print meZMNUUTI0.MNUUTIGRS;
libMNURUT = meZMNURUT0.MNURUTCUT & " : " & meZMNURUT0.MNURUTUTI & "    " & meZMNURUT0.MNURUTNOM

libRéférenceInterne = libMNURUT

End Sub

Public Sub cmdPrint_fgZMNUMEN0(lFct As String)
Dim colX As Integer, col1 As Integer, col2 As Integer, col3 As Integer
Dim wText As String, I0 As Integer, mNiveau As Integer
Dim blnPrintAll As Boolean
Dim K As Integer

mNiveau = 1
If lFct = "" Then
    blnPrintAll = True
    fgZMNUMEN0.Row = 1
Else
    blnPrintAll = False
End If
prtTitleText = fgZMNUMEN0_Lib & meZMNURUT0.MNURUTUTI & " - " & meZMNURUT0.MNURUTNOM
I0 = fgZMNUMEN0.Row

prtOrientation = vbPRORLandscape
prtPgmName = "prtElpKM"
prtTitleUsr = usrName

prtFontName = "Century Gothic"
prtElpKM.prtStd_Open

XPrt.FontSize = 7
col1 = prtMinX
col2 = prtMinX + (prtMaxX - prtMinX) / 3 - 100
col3 = prtMinX + 2 * (prtMaxX - prtMinX) / 3 - 100
XPrt.Line (col3 - 50, prtMinY)-(col3 - 50, prtMaxY), prtLineColor
XPrt.Line (col2 - 50, prtMinY)-(col2 - 50, prtMaxY), prtLineColor

colX = col1
XPrt.CurrentY = prtMinY + prtHeaderHeight - prtlineHeight
For I = I0 To fgZMNUMEN0.Rows - 1
    fgZMNUMEN0.Row = I
    
    If Not blnPrintAll Then
        fgZMNUMEN0.Col = fgZMNUMEN0_arrIndex: K = fgZMNUMEN0.Text
        xZMNUMEN0 = arrZMNUMEN0(K)
        If I = I0 Then
            mNiveau = arrZMNUMEN0(K).Niveau
        Else
            If mNiveau >= arrZMNUMEN0(K).Niveau Then Exit For
        End If
    End If
    fgZMNUMEN0.Col = 0: wText = Trim(fgZMNUMEN0.Text)
    fgZMNUMEN0.Col = 1: wText = wText & " " & Trim(fgZMNUMEN0.Text)
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    If XPrt.CurrentY + prtlineHeight >= prtMaxY Then
        Select Case colX
            Case col1: colX = col2
                    'XPrt.Line (colX - 50, prtMinY)-(colX - 50, prtMaxY), prtLineColor

            Case col2: colX = col3
               ' XPrt.Line (colX - 50, prtMinY)-(colX - 50, prtMaxY), prtLineColor
            Case Else:
                    frmElpPrt.prtNewPage
                    XPrt.Line (col3 - 50, prtMinY)-(col3 - 50, prtMaxY), prtLineColor
                    XPrt.Line (col2 - 50, prtMinY)-(col2 - 50, prtMaxY), prtLineColor
                    XPrt.FontSize = 7
                    colX = col1
       End Select
        XPrt.CurrentY = prtMinY + prtHeaderHeight
    End If
    XPrt.CurrentX = colX
    If wText = UCase$(wText) Then
        XPrt.FontBold = True
        XPrt.Print wText;
        XPrt.FontBold = False
    Else
        XPrt.Print wText;
    End If
Next I

prtElpKM.prtStd_Close

End Sub


Public Sub fgSelect_MouseDown_Ok()
On Error Resume Next
   If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex
        arrElpKM_Index = Val(fgSelect.Text)
        mElpKMSrc_Id = arrElpKMSrc_Id(arrElpKM_Index)
        mElpKM_Classe = arrElpKM_Classe(arrElpKM_Index)
       mElpKMInfo_Id = arrElpKMInfo_Id(arrElpKM_Index)
        fgSelect.Col = 1
        mElpKM_Classe_Abrégé = Trim(fgSelect.Text)
        fgSelect.Col = 0
        mElpKMInfo_Description = Trim(fgSelect.Text)
        libRéférenceInterne = libMNURUT & Chr$(13) & mElpKM_Classe_Abrégé & "          " & mElpKMInfo_Id & Chr$(13) & mElpKMInfo_Description
      '  mnuSelectUsr_Click
        fgSelect.Col = 4
        mElpKMInfo_ElpKmSrc_Id = CLng(Trim(fgSelect.Text))
    End If

End Sub

Public Sub lstUsr_Select()
Dim xSQL As String
On Error GoTo Error_Handler

App_Debug = "> lstUsr_Select"
'--------------------------------------------------------------------------------------

xSQL = "select * from " & paramIBM_Library_SAB & ".ZMNURUT0" _
     & " where MNURUTUTI = '" & Trim(Mid$(lstUsr.Text, 1, 10)) & "'"
     
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    Call rsZMNURUT0_GetBuffer(rsSab, meZMNURUT0)
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZMNUUTI0" _
         & " where MNUUTICUT = " & meZMNURUT0.MNURUTCUT _
         & " and   MNUUTIETB = " & meZMNURUT0.MNURUTETB
         
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        Call rsZMNUUTI0_GetBuffer(rsSab, meZMNUUTI0)
        picUsr_Display
    End If
End If
Exit Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug

End Sub


Public Sub ZMNUHLB0_Load()
Dim xSQL As String
Dim xWhere As String

arrZMNUHLB0_Nb = 0
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "ZMNUHLB0_Load")
xWhere = " where   MNUHLBREF = " & currentZMNUHLB0.MNUHLBREF _
     & " and   MNUHLBETB = " & currentZMNURUT0.MNURUTETB _
     & " and   MNUHLBCLA = 2"
xSQL = "select count(*)  as Tally   from " & paramIBM_Library_SAB & ".ZMNUHLB0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

ReDim arrZMNUHLB0(rsSab("Tally") + 1)
xSQL = "select MNUHLBNOM from " & paramIBM_Library_SAB & ".ZMNUHLB0" & xWhere & " order by MNUHLBNOM"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrZMNUHLB0_Nb = arrZMNUHLB0_Nb + 1
    arrZMNUHLB0(arrZMNUHLB0_Nb).MNUHLBNOM = rsSab("MNUHLBNOM")
    rsSab.MoveNext
Loop

End Sub
