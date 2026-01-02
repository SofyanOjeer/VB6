VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCPT_SCHEMA 
   AutoRedraw      =   -1  'True
   Caption         =   "CPT_SCHEMA : validation des modifications des schémas comptables"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "CPT_SCHEMA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   8280
      TabIndex        =   4
      Top             =   45
      Width           =   5055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Sélection"
      TabPicture(0)   =   "CPT_SCHEMA.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "."
      TabPicture(1)   =   "CPT_SCHEMA.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstW"
      Tab(1).ControlCount=   1
      Begin VB.ListBox lstW 
         Height          =   255
         Left            =   -67800
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   90
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame fraSelect 
         Height          =   8445
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13560
         Begin MSComCtl2.DTPicker txtSelect_AmjMax 
            Height          =   300
            Left            =   12240
            TabIndex        =   21
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "dd  MM yyy"
            Format          =   119275523
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin VB.ListBox lstSelect 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5520
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   20
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Frame fraUpdate 
            BackColor       =   &H00F0F0F0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   7215
            Left            =   7830
            TabIndex        =   11
            Top             =   1125
            Visible         =   0   'False
            Width           =   5175
            Begin VB.ListBox lstUpdate_A 
               ForeColor       =   &H00FF00FF&
               Height          =   1230
               Left            =   135
               TabIndex        =   43
               Top             =   3645
               Width           =   4935
            End
            Begin VB.Frame fraUpdate_A 
               BackColor       =   &H00F0F0F0&
               Height          =   2535
               Left            =   150
               TabIndex        =   26
               Top             =   210
               Width           =   4935
               Begin VB.TextBox txtUpdate_CPTSCHAMJ2 
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   42
                  Top             =   1560
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_CPTSCHUSR2 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   41
                  Top             =   1560
                  Width           =   1455
               End
               Begin VB.TextBox txtUpdate_CPTSCHAMJ1 
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   39
                  Top             =   1200
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_CPTSCHUSR1 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   34
                  Top             =   1200
                  Width           =   1455
               End
               Begin VB.TextBox txtUpdate_CPTSCHUSR 
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   33
                  Text            =   "usr"
                  Top             =   2040
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_CPTSCHSTA 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   32
                  Top             =   2040
                  Width           =   615
               End
               Begin VB.TextBox txtUpdate_SCHEMAFDT 
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   31
                  Top             =   720
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_SCHEMAFUT 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   30
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.TextBox txtUpdate_SCHEMAARG 
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   29
                  Top             =   240
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_SCHEMAOPE 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   28
                  Top             =   240
                  Width           =   615
               End
               Begin VB.TextBox txtUpdate_SCHEMAEVE 
                  Height          =   285
                  Left            =   2160
                  TabIndex        =   27
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label lblUpdate_CPTSCHUSR2 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "Visa RSSI"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   40
                  Top             =   1680
                  Width           =   975
               End
               Begin VB.Label lblUpdate_BDFCMPUSR 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "User"
                  Height          =   255
                  Left            =   2040
                  TabIndex        =   38
                  Top             =   2040
                  Width           =   735
               End
               Begin VB.Label lblUpdate_CPTSCHUSR1 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "Visa Compta"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   37
                  Top             =   1200
                  Width           =   975
               End
               Begin VB.Label lblUpdate_BDFCMPSTA 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "Statut"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   36
                  Top             =   2040
                  Width           =   975
               End
               Begin VB.Label lblUpdate_SCHEMAOPE 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "Dossier"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF80FF&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   35
                  Top             =   240
                  Width           =   1095
               End
            End
            Begin VB.Frame fraUpdate_B 
               BackColor       =   &H00D0D0D0&
               Caption         =   "Motif / Visa"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2205
               Left            =   120
               TabIndex        =   22
               Top             =   4965
               Width           =   4935
               Begin VB.CommandButton cmdVPjointes 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "Voir les pièces jointes"
                  Height          =   525
                  Left            =   3000
                  Style           =   1  'Graphical
                  TabIndex        =   49
                  Top             =   1635
                  Width           =   1700
               End
               Begin MSComDlg.CommonDialog Dial1 
                  Left            =   4440
                  Top             =   1200
                  _ExtentX        =   847
                  _ExtentY        =   847
                  _Version        =   393216
               End
               Begin VB.CommandButton cmdPjointes 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "Ajouter une pièce jointe"
                  Height          =   1125
                  Left            =   2040
                  Style           =   1  'Graphical
                  TabIndex        =   47
                  Top             =   1035
                  Width           =   850
               End
               Begin VB.TextBox txtUpdate_CPTSCHTEXT 
                  Height          =   615
                  Left            =   240
                  MaxLength       =   64
                  MultiLine       =   -1  'True
                  TabIndex        =   44
                  Top             =   360
                  Width           =   4455
               End
               Begin VB.CommandButton cmdUpdate_Annuler 
                  BackColor       =   &H000000FF&
                  Caption         =   "Annuler le Visa"
                  Height          =   525
                  Left            =   240
                  Style           =   1  'Graphical
                  TabIndex        =   25
                  Top             =   1035
                  Width           =   1700
               End
               Begin VB.CommandButton cmdUpdate_Quit 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Abandonner"
                  Height          =   525
                  Left            =   240
                  Style           =   1  'Graphical
                  TabIndex        =   24
                  Top             =   1635
                  Width           =   1700
               End
               Begin VB.CommandButton cmdUpdate_Ok 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Enregistrer"
                  Height          =   525
                  Left            =   3000
                  Style           =   1  'Graphical
                  TabIndex        =   23
                  Top             =   1035
                  Width           =   1700
               End
               Begin VB.FileListBox File1 
                  Height          =   870
                  Left            =   180
                  TabIndex        =   48
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   555
               End
               Begin VB.Label labNomRep 
                  Caption         =   "Label1"
                  Height          =   255
                  Left            =   1560
                  TabIndex        =   50
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   2715
               End
            End
            Begin VB.Label libUpdate_EVE 
               BackColor       =   &H00C0C000&
               Caption         =   "libUpdate_EVE"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   165
               TabIndex        =   46
               Top             =   3195
               Width           =   4875
            End
            Begin VB.Label libUpdate_OPE 
               BackColor       =   &H00808000&
               Caption         =   "libUpdate_OPE"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   150
               TabIndex        =   45
               Top             =   2835
               Width           =   4920
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7185
            Left            =   120
            TabIndex        =   8
            Top             =   1080
            Visible         =   0   'False
            Width           =   13440
            _ExtentX        =   23707
            _ExtentY        =   12674
            _Version        =   393216
            Rows            =   1
            Cols            =   10
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   16777210
            ForeColor       =   8388608
            BackColorFixed  =   16776921
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   16777210
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"CPT_SCHEMA.frx":0044
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.ComboBox cboSelect_SQL 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   9
            Text            =   "cboSelect_SQL"
            Top             =   260
            Width           =   4215
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Rechercher"
            Height          =   525
            Left            =   11160
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   1815
         End
         Begin VB.Frame fraSelect_Options_1 
            Height          =   915
            Left            =   4440
            TabIndex        =   6
            Top             =   120
            Width           =   5955
            Begin VB.CheckBox chkSelect_CPTSCHDCRE 
               Caption         =   "Période de création"
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox txtSelect_SCHEMAEVE 
               Height          =   285
               Left            =   4920
               TabIndex        =   14
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox txtSelect_SCHEMAOPE 
               Height          =   285
               Left            =   4920
               TabIndex        =   13
               Top             =   240
               Width           =   615
            End
            Begin MSComCtl2.DTPicker txtSelect_CPTSCHDCRE 
               Height          =   300
               Left            =   2040
               TabIndex        =   12
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   118685699
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_CPTSCHDCRE_Max 
               Height          =   300
               Left            =   2040
               TabIndex        =   19
               Top             =   600
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   118685699
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_SCHEMAEVE 
               Caption         =   "Evénement"
               Height          =   255
               Left            =   3600
               TabIndex        =   16
               Top             =   600
               Width           =   855
            End
            Begin VB.Label lblSelect_SCHEMAOPE 
               Caption         =   "Code opération"
               Height          =   255
               Left            =   3600
               TabIndex        =   15
               Top             =   240
               Width           =   1215
            End
         End
         Begin MSComCtl2.DTPicker txtSelect_AmjMin 
            Height          =   300
            Left            =   10680
            TabIndex        =   10
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "dd  MM yyy"
            Format          =   118685699
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "CPT_SCHEMA.frx":00D4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Width           =   5535
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuContext_x1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuselect 
      Caption         =   "mnuSelect"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Quit 
         Caption         =   "Abandonner"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint0_All 
         Caption         =   "Imprimer TOUS les courriers"
      End
   End
End
Attribute VB_Name = "frmCPT_SCHEMA"
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
Dim BIA_CPTSCH_Aut As typeAuthorization
Dim blnTransaction As Boolean
Dim blnAuto As Boolean, blnAuto_Ok As Boolean
Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
Dim wAmjMin7 As Long, wAmjMax7 As Long


Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnSetfocus As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim mCLIENARA1 As String

'______________________________________________________________________

Dim xYCPTSCH0 As typeYCPTSCH0, meYCPTSCH0 As typeYCPTSCH0
Dim newYCPTSCH0 As typeYCPTSCH0, oldYCPTSCH0 As typeYCPTSCH0
Dim arrYCPTSCH0() As typeYCPTSCH0, arrYCPTSCH0_Nb As Long, arrYCPTSCH0_Max As Long, arrYCPTSCH0_Index As Long
Dim selYCPTSCH0() As typeYCPTSCH0, selYCPTSCH0_Nb As Long, selYCPTSCH0_Max As Long, selYCPTSCH0_Index As Long
Dim cmdSelect_Ok_Caption As String
Dim cmdSelect_SQL_K As String



Dim xZSCHEMAH0 As typeZSCHEMAH0, meZSCHEMAH0 As typeZSCHEMAH0

Dim arrMNURUTUTI() As String
Dim blnSendMail As Boolean
Private Function cadrage32(z As String, il As Long) As String
Dim iz As Long
Dim newil As Long

    iz = CLng(z)
    
    
End Function

Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
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
        For I = fgSelect_arrIndex To 0 Step -1
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.LeftCol = 0
    End If
End If

End Sub
Private Sub fgSelect_Display()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
cmdPrint.Enabled = False
currentAction = "fgselect_Display"

For I = 1 To arrYCPTSCH0_Nb

        xYCPTSCH0 = arrYCPTSCH0(I)
    
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
Next I

Call lstErr_Clear(lstErr, cmdContext, "Nb enregistrements : " & fgSelect.Rows - 1): DoEvents
If fgSelect.Rows > 1 Then
'    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
    cmdPrint.Enabled = True
End If
fgSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Private Sub lstSelect_Load_1()
Dim I As Long, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_1"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = True
fraSelect_Options_1.Enabled = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub lstSelect_Load_2()
Dim I As Long, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_2"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Private Sub lstSelect_Load_3()
Dim I As Long, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_3"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
Dim X As String, wColor As Long

On Error Resume Next
Select Case xYCPTSCH0.CPTSCHSTA
    Case Is = " ": wColor = vbBlue
    Case Else: wColor = vbGrayText
End Select

fgSelect.Col = 0: fgSelect.Text = dateIBM10(xYCPTSCH0.SCHEMAFDT, True)
fgSelect.CellForeColor = wColor
fgSelect.Col = 1: fgSelect.Text = arrMNURUTUTI(xYCPTSCH0.SCHEMAFUT)
fgSelect.CellForeColor = wColor
fgSelect.Col = 2: fgSelect.Text = xYCPTSCH0.SCHEMAETA & "_" & xYCPTSCH0.SCHEMAPLA
fgSelect.CellForeColor = wColor
fgSelect.Col = 3: fgSelect.Text = xYCPTSCH0.SCHEMAOPE & "  " & xYCPTSCH0.SCHEMAEVE & "  " & Trim(xYCPTSCH0.SCHEMAARG)
fgSelect.CellForeColor = wColor
fgSelect.Col = 4: fgSelect.Text = xYCPTSCH0.CPTSCHUSR1
fgSelect.CellForeColor = wColor
If xYCPTSCH0.CPTSCHAMJ1 > 0 Then
    fgSelect.Col = 5: fgSelect.Text = dateImp10(xYCPTSCH0.CPTSCHAMJ1) & " " & timeNImp8(xYCPTSCH0.CPTSCHHMS1)
    fgSelect.CellForeColor = wColor
Else
    fgSelect.Col = 5: fgSelect.Text = ""
End If
fgSelect.Col = 6: fgSelect.Text = xYCPTSCH0.CPTSCHUSR2
fgSelect.CellForeColor = wColor
If xYCPTSCH0.CPTSCHAMJ2 > 0 Then
    fgSelect.Col = 7: fgSelect.Text = dateImp10(xYCPTSCH0.CPTSCHAMJ2) & " " & timeNImp8(xYCPTSCH0.CPTSCHHMS2)
    fgSelect.CellForeColor = wColor
Else
    fgSelect.Col = 7: fgSelect.Text = ""
End If

fgSelect.Col = 8: fgSelect.Text = retourne_nom_repertoire(xYCPTSCH0)

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

End Sub

Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = fgSelect.Cols - 1
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = 0

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
Dim wIndex As Integer
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    wIndex = Val(fgSelect.Text)
    Select Case lK
        Case 0: X = Format$(arrYCPTSCH0(wIndex).SCHEMAFDT, "00000000")
        Case 5: X = Format$(arrYCPTSCH0(wIndex).CPTSCHAMJ1, "00000000") & Format$(arrYCPTSCH0(wIndex).CPTSCHHMS1, "000000")
        Case 7: X = Format$(arrYCPTSCH0(wIndex).CPTSCHAMJ2, "00000000") & Format$(arrYCPTSCH0(wIndex).CPTSCHHMS2, "000000")
    End Select
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
If fraUpdate.Visible Then fraUpdate.Visible = False: Exit Sub
If fgSelect.Visible Then fgSelect.Visible = False: cmdSelect_Ok.Caption = "Extraire les mouvements": Exit Sub
Unload Me
End Sub




Private Function nettoyage_nom_repertoire(z As String) As String
Dim ii As Long
Dim zz As String
Dim asci As Long

    'suppression des caractères interdits dans un nom de répertoire
    'code ASCII compris entre A-Z, a-z et 0-9
    zz = ""
    For ii = 1 To Len(z)
        asci = Asc(Mid(z, ii, 1))
        If (asci > 47 And asci < 58) Or (asci > 64 And asci < 91) Or (asci > 96 And asci < 123) Then
            'on garde ce caractère sinon on le supprime
            zz = zz & Chr(asci)
        End If
    Next ii
    nettoyage_nom_repertoire = zz
    
End Function

Private Function retourne_extension(z As String) As String
Dim s() As String
Dim zz As String

    zz = ""
    s = Split(z, ".")
    If UBound(s) > 0 Then
        zz = s(UBound(s))
    End If
    retourne_extension = zz
    
End Function

Private Function retourne_nom_repertoire(dYCPTSCH0 As typeYCPTSCH0) As String
Dim nrepertoire As String

    nrepertoire = Trim(dYCPTSCH0.SCHEMAFDT)
    nrepertoire = nrepertoire & Mid(CStr(CLng(dYCPTSCH0.SCHEMAFUT) + 10000), 2)
    nrepertoire = nrepertoire & CStr(dYCPTSCH0.SCHEMAETA)
    nrepertoire = nrepertoire & Trim(dYCPTSCH0.SCHEMAOPE)
    nrepertoire = nrepertoire & Trim(dYCPTSCH0.SCHEMAEVE)
    nrepertoire = nrepertoire & Mid(CStr(CLng(dYCPTSCH0.SCHEMAPLA) + 100), 2)
    nrepertoire = nrepertoire & Trim(dYCPTSCH0.SCHEMAARG)
    'nrepertoire = nrepertoire & Mid(CStr(CLng(dYCPTSCH0.CPTSCHUPDS) + 100), 2)
    nrepertoire = nettoyage_nom_repertoire(nrepertoire)
    retourne_nom_repertoire = nrepertoire
    
End Function

Private Sub cboSelect_SQL_Click()
cmdSelect_SQL_K = Mid$(cboSelect_SQL, 1, 1)
If blnControl Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    lstSelect.Visible = False
    txtSelect_AmjMin.Visible = False
    txtSelect_AmjMax.Visible = False
    fraSelect_Options_1.Visible = False
    fraUpdate.Visible = False
    Select Case cmdSelect_SQL_K
        Case "1": lstSelect_Load_1
        Case "2": lstSelect_Load_2
        Case "3": lstSelect_Load_3
        Case "7": lstSelect_Load_7
    End Select
    Me.Enabled = True: Me.MousePointer = 0
End If
End Sub


Private Sub chkSelect_CPTSCHDCRE_Click()
If chkSelect_CPTSCHDCRE = "1" Then
    If cmdSelect_SQL_K = "1" Then txtSelect_CPTSCHDCRE.Visible = True
    txtSelect_CPTSCHDCRE_Max.Visible = True
Else
    txtSelect_CPTSCHDCRE.Visible = False
    txtSelect_CPTSCHDCRE_Max.Visible = False
End If


End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdPjointes_Click()
Dim extension As String
Dim nomRep As String
Dim ind As Long

    Dial1.DefaultExt = "*.*"
    Dial1.FileName = ""
    Dial1.DialogTitle = "Ajouter une pièce jointe."
    Dial1.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    Dial1.ShowOpen
    If Dial1.FileName <> "" Then
        Screen.MousePointer = vbHourglass
        nomRep = labNomRep.Caption
        If Dir(paramCPT_SCHEMA_Dossier_Path & nomRep, vbDirectory) = "" Then
            MkDir paramCPT_SCHEMA_Dossier_Path & nomRep
            ind = 0
        Else
            File1.Path = paramCPT_SCHEMA_Dossier_Path & nomRep
            File1.Refresh
            ind = File1.ListCount
        End If
        ind = ind + 1
        extension = retourne_extension(Dial1.FileName)
        FileCopy Dial1.FileName, paramCPT_SCHEMA_Dossier_Path & nomRep & "\piece_" & Mid(CStr((ind + 100)), 2) & "." & extension
        Call SetAttr(paramCPT_SCHEMA_Dossier_Path & nomRep & "\piece_" & Mid(CStr((ind + 100)), 2) & "." & extension, vbArchive + vbReadOnly)
        Call Sleep(2000)
        If Dir(paramCPT_SCHEMA_Dossier_Path & nomRep & "\piece_" & Mid(CStr((ind + 100)), 2) & "." & extension) <> "" Then
            Call MsgBox("La pièce jointe a été ajoutée...")
            cmdVPjointes.Enabled = True
            Screen.MousePointer = vbDefault
        End If
    End If
    
End Sub
Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
Dim I As Integer

blnControl = False
usrColor_Set

cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
blncmdOk_Visible = False: blncmdSave_Visible = False
currentAction = ""

blnAuto = False
blnAuto_Ok = False
lstSelect.Visible = False
cmdSelect_Ok.Caption = "Extraire les mouvements"

libRéférenceInterne = ""
cboSelect_SQL.ListIndex = 0
fgSelect.Visible = False
fraUpdate.Visible = False
blnControl = True
cboSelect_SQL.ListIndex = 0
End Sub
Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

SSTab1.Tab = 0

blnControl = False

fgSelect_FormatString = fgSelect.FormatString
cmdSelect_Ok.Visible = False
fraSelect_Options_1.Visible = False
txtSelect_CPTSCHDCRE.Visible = False
txtSelect_CPTSCHDCRE_Max.Visible = False
cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1 - Consultation des modifications"
'If BIA_CPTSCH_Aut.Comptabiliser Then
cboSelect_SQL.AddItem "2 - modifications à valider par le service comptable"
'If BIA_CPTSCH_Aut.Valider Then
cboSelect_SQL.AddItem "3 - modifications à valider par le Contrôleur comptable"
If BIA_CPTSCH_Aut.Rapprocher Then cboSelect_SQL.AddItem "7 - Importation ds modifications (ZSCHEMAH0)"

Call DTPicker_Set(txtSelect_AmjMin, YBIATAB0_DATE_CPT_JP0)
Call DTPicker_Set(txtSelect_AmjMax, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtSelect_CPTSCHDCRE, YBIATAB0_DATE_CPT_JP0)
Call DTPicker_Set(txtSelect_CPTSCHDCRE_Max, YBIATAB0_DATE_CPT_J)

arrMNURUUTI_Load arrMNURUTUTI
cmdReset
libUpdate_OPE.ForeColor = vbWhite
libUpdate_EVE.ForeColor = vbWhite


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

Me.Enabled = False: Me.MousePointer = vbHourglass
    Select Case cmdSelect_SQL_K
    '    Case "2": cmdPrint_Ok_2
    '    Case "3": cmdPrint_Ok_3
    End Select

Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

Me.Enabled = False: Me.MousePointer = vbHourglass
blnOk = Not fgSelect.Visible
Call lstErr_Clear(lstErr, cmdContext, "> SAB_CDR_cmdSelect_Ok ........"): DoEvents
cmdSelect_Ok.Visible = False
fraUpdate.Visible = False
fraSelect_Options_1.Enabled = False
txtSelect_AmjMin.Enabled = False
txtSelect_AmjMax.Enabled = False
fgSelect.Clear: fgSelect.Rows = 1
DoEvents
If blnOk Then
    cmdSelect_Ok.Caption = "Modifier les options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    lstSelect.BackColor = &H8000000F
    Call usrColor_Container(lstSelect, lstSelect.BackColor)
    Select Case cmdSelect_SQL_K
        Case "1": cmdSelect_SQL
        Case "2": cmdSelect_SQL_2
        Case "3": cmdSelect_SQL_3
        Case "7": cmdSelect_SQL_7
    End Select

    fgSelect.Visible = True: fgSelect.Enabled = True
Else
    cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
    cmdSelect_Ok.BackColor = &HC0FFC0
    lstSelect.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(lstSelect, lstSelect.BackColor)
    fgSelect.Visible = False
    fgSelect.Enabled = False
    fraSelect_Options_1.Enabled = True
    txtSelect_AmjMin.Enabled = True
    txtSelect_AmjMax.Enabled = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_CDR_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
cmdSelect_Ok.Visible = True

End Sub


Private Sub cmdSelect_SQL()
Dim V
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL"): DoEvents

currentAction = "cmdSelect_SQL"

Set rsSab = Nothing
Call DTPicker_Control(txtSelect_CPTSCHDCRE, wAmjMin)
Call DTPicker_Control(txtSelect_CPTSCHDCRE_Max, wAmjMax)

If chkSelect_CPTSCHDCRE = "1" Then
    xWhere = xWhere & " and SCHEMAFDT >= " & wAmjMin - 19000000 _
                    & " and SCHEMAFDT <= " & wAmjMax - 19000000
End If
X = Trim(txtSelect_SCHEMAOPE)
If X <> "" Then xWhere = xWhere & " and SCHEMAOPE like '%" & X & "%'"
X = Trim(txtSelect_SCHEMAEVE)
If X <> "" Then xWhere = xWhere & " and SCHEMAEVE like '%" & X & "%'"

xWhere = Replace(xWhere, "and", "where", , 1)
arrYCPTSCH0_SQL xWhere & " order by SCHEMAFDT,SCHEMAFUT,SCHEMAOPE,SCHEMAEVE"
    
fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_2()
Dim V
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String
Dim wNb As Long, wCPTSCHMONE As Currency, wCPTSCHSTAT As String, xCPTSCHSTAT As String

On Error GoTo Error_Handler

xWhere = " where CPTSCHUSR1 = '' "
arrYCPTSCH0_SQL xWhere & " order by SCHEMAFDT,SCHEMAFUT,SCHEMAOPE,SCHEMAEVE"
    
fgSelect_Display
    
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_3()
Dim V
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler
xWhere = " where CPTSCHUSR2 = '' AND CPTSCHUSR1 <> ''"
arrYCPTSCH0_SQL xWhere & " order by SCHEMAFDT,SCHEMAFUT,SCHEMAOPE,SCHEMAEVE"
    
fgSelect_Display
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_7()
Dim V
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_7"): DoEvents

currentAction = "cmdSelect_SQL_7"
Call DTPicker_Control(txtSelect_AmjMin, wAmjMin)
Call DTPicker_Control(txtSelect_AmjMax, wAmjMax)
If wAmjMin = "00000000" Then
    MsgBox "Préciser la date", vbInformation, "Import des modifications à une date"
    Exit Sub
End If
    
cmdSelect_SQL_7_ZSCHEMAH0

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_7_ZSCHEMAH0()
Dim V
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String, X As String
Dim blnOk As Boolean
On Error GoTo Error_Handler

blnSendMail = False
blnOk = False
rsYCPTSCH0_Init xYCPTSCH0
Set rsSab = Nothing

xWhere = " where SCHEMAFDT >= " & wAmjMin - 19000000 & " and  SCHEMAFDT <= " & wAmjMax - 19000000
xSQL = "select * from " & paramIBM_Library_SAB & ".ZSCHEMAH0 " & xWhere & " order by SCHEMAFDT,SCHEMAFUT,SCHEMAETA,SCHEMAOPE,SCHEMAEVE,SCHEMAPLA,SCHEMAARG"
Set rsSab = cnsab.Execute(xSQL)

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

Do While Not rsSab.EOF
    V = rsZSCHEMAH0_GetBuffer(rsSab, xZSCHEMAH0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "cmdSelect_SQL_7_ZSCHEMAH0"
        '' Exit Sub
     Else
        If xYCPTSCH0.SCHEMAFDT <> xZSCHEMAH0.SCHEMAFDT _
        Or xYCPTSCH0.SCHEMAFUT <> xZSCHEMAH0.SCHEMAFUT _
        Or xYCPTSCH0.SCHEMAETA <> xZSCHEMAH0.SCHEMAETA _
        Or xYCPTSCH0.SCHEMAOPE <> xZSCHEMAH0.SCHEMAOPE _
        Or xYCPTSCH0.SCHEMAEVE <> xZSCHEMAH0.SCHEMAEVE _
        Or xYCPTSCH0.SCHEMAPLA <> xZSCHEMAH0.SCHEMAPLA _
        Or xYCPTSCH0.SCHEMAARG <> xZSCHEMAH0.SCHEMAARG Then
            blnSendMail = True
            If blnOk Then
                V = sqlYCPTSCH0_Insert(xYCPTSCH0)
                'If Not IsNull(V) Then
                '    If InStr(V, "SQL0803") = 0 Then GoTo Error_MsgBox ' ignorer clé en double
                'End If
            End If
            blnOk = True
            xYCPTSCH0.SCHEMAFDT = xZSCHEMAH0.SCHEMAFDT
            xYCPTSCH0.SCHEMAFUT = xZSCHEMAH0.SCHEMAFUT
            xYCPTSCH0.SCHEMAETA = xZSCHEMAH0.SCHEMAETA
            xYCPTSCH0.SCHEMAOPE = xZSCHEMAH0.SCHEMAOPE
            xYCPTSCH0.SCHEMAEVE = xZSCHEMAH0.SCHEMAEVE
            xYCPTSCH0.SCHEMAPLA = xZSCHEMAH0.SCHEMAPLA
            xYCPTSCH0.SCHEMAARG = xZSCHEMAH0.SCHEMAARG
        End If

    End If
'______________________________________________________
    

    rsSab.MoveNext

Loop
If blnOk Then
    V = sqlYCPTSCH0_Insert(xYCPTSCH0)
    V = Null
               ' If Not IsNull(V) Then
               '     If InStr(V, "SQL0803") = 0 Then GoTo Error_MsgBox ' ignorer clé en double
               ' End If
    
End If
    
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
    '$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<



End Sub
Private Sub fraUpdate_Display_ZSCHEMAH0()
Dim V
Dim xSQL As String, K As Long, K0 As Long
Dim xWhere As String, xAnd As String, X As String
On Error GoTo Error_Handler

lstUpdate_A.Clear
lstUpdate_A.ForeColor = warnUsrColor
Set rsSab = Nothing

xWhere = " where SCHEMAFDT = " & xYCPTSCH0.SCHEMAFDT _
       & " and   SCHEMAFUT = " & xYCPTSCH0.SCHEMAFUT _
       & " and   SCHEMAETA = " & xYCPTSCH0.SCHEMAETA _
       & " and   SCHEMAOPE = '" & xYCPTSCH0.SCHEMAOPE & "'" _
       & " and   SCHEMAEVE = '" & xYCPTSCH0.SCHEMAEVE & "'" _
       & " and   SCHEMAPLA = " & xYCPTSCH0.SCHEMAPLA _
       & " and   SCHEMAARG = '" & xYCPTSCH0.SCHEMAARG & "'" _

xSQL = "select SCHEMANUM,SCHEMAFHE,SCHEMALIB from " & paramIBM_Library_SAB & ".ZSCHEMAH0 " & xWhere & " order by SCHEMANUM,SCHEMAFHE"
Set rsSab = cnsab.Execute(xSQL)

K0 = 0
Do While Not rsSab.EOF
'______________________________________________________
    K = Val(rsSab("SCHEMANUM"))
    If K <> K0 Then K0 = K: lstUpdate_A.AddItem ""
    lstUpdate_A.AddItem Format$(K, "000") & "   " & dateIBM10(xYCPTSCH0.SCHEMAFDT, True) & "   " & timeImp(rsSab("SCHEMAFHE")) & "  " & Trim(rsSab("SCHEMALIB"))
    rsSab.MoveNext

Loop
Exit Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug

End Sub


Public Sub cmdSendMail()
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim X As String

X = "Créations / Modifications de schémas comptables au " & dateImp(wAmjMin)
    
wSendMail.FromDisplayName = "CPT_SCHEMA"
wSendMail.RecipientDisplayName = "AUDIT"

bgColor = "CYAN"
wSendMail.Subject = X
wSendMail.Attachment = ""
wSendMail.Message = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">" _
                    & "<FONT face=" & Asc34 & prtFontName_Arial & Asc34 & ">" _
                    & htmlFontColor("BLUE") & "<B></CENTER><U>" _
                    & "Il y a des schémas comptables à valider par le service comptable puis le Contrôle Comptable."

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub

Private Sub arrYCPTSCH0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrYCPTSCH0(501)
arrYCPTSCH0_Max = 500: arrYCPTSCH0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCPTSCH0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYCPTSCH0_GetBuffer(rsSab, xYCPTSCH0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgselect_Display"
        '' Exit Sub
     Else
         arrYCPTSCH0_Nb = arrYCPTSCH0_Nb + 1
         If arrYCPTSCH0_Nb > arrYCPTSCH0_Max Then
             arrYCPTSCH0_Max = arrYCPTSCH0_Max + 50
             ReDim Preserve arrYCPTSCH0(arrYCPTSCH0_Max)
         End If
         
         arrYCPTSCH0(arrYCPTSCH0_Nb) = xYCPTSCH0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdUpdate_Annuler_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents

newYCPTSCH0 = oldYCPTSCH0
If Trim(oldYCPTSCH0.CPTSCHUSR2) <> "" Then
    newYCPTSCH0.CPTSCHUSR2 = "": newYCPTSCH0.CPTSCHHMS2 = 0: newYCPTSCH0.CPTSCHAMJ2 = 0
Else
    newYCPTSCH0.CPTSCHUSR1 = "": newYCPTSCH0.CPTSCHHMS1 = 0: newYCPTSCH0.CPTSCHAMJ1 = 0
End If

    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents
    V = cmdUpdate_Ok_Transaction
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        arrYCPTSCH0(arrYCPTSCH0_Index) = newYCPTSCH0
        xYCPTSCH0 = newYCPTSCH0
        fgSelect_DisplayLine arrYCPTSCH0_Index
        fraUpdate.Visible = False

    Else
        MsgBox V, vbCritical, Me.Name & " : cmdUpdate_Ok"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdUpdate_Ok_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents

If IsNull(fraUpdate_Control) Then
    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents
    V = cmdUpdate_Ok_Transaction
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        arrYCPTSCH0(arrYCPTSCH0_Index) = newYCPTSCH0
        xYCPTSCH0 = newYCPTSCH0
        fgSelect_DisplayLine arrYCPTSCH0_Index
        fraUpdate.Visible = False
    Else
        MsgBox V, vbCritical, Me.Name & " : cmdUpdate_Ok"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdUpdate_Quit_Click()
fraUpdate.Visible = False

End Sub

Private Sub cmdVPjointes_Click()
Dim I As Long

    Load frmPieces
    fgSelect.Col = 0
    frmPieces.Label1.Caption = fgSelect.Text
    fgSelect.Col = 1
    frmPieces.Label1.Caption = frmPieces.Label1.Caption & " " & fgSelect.Text
    fgSelect.Col = 2
    frmPieces.Label1.Caption = frmPieces.Label1.Caption & " " & fgSelect.Text
    fgSelect.Col = 3
    frmPieces.Label1.Caption = frmPieces.Label1.Caption & " " & fgSelect.Text
    frmPieces.List1.Clear
    For I = 0 To File1.ListCount - 1
        frmPieces.List1.AddItem File1.List(I)
    Next I
    frmPieces.Show vbModal
    
End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
Me.Enabled = False
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
        Select Case fgSelect.Col
            Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_SortX 0
            Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
            Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
            Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
            Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
            Case 5: fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_SortX 5
            Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
            Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_SortX 7
            Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
           Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
        End Select
Else
    If fgSelect.Rows > 1 Then
        fgSelect.Col = fgSelect_arrIndex:  arrYCPTSCH0_Index = CLng(fgSelect.Text)
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        xYCPTSCH0 = arrYCPTSCH0(arrYCPTSCH0_Index)
        oldYCPTSCH0 = xYCPTSCH0
        '                                       '
        fgSelect.Col = 8
        labNomRep.Caption = fgSelect.Text
        If Dir(paramCPT_SCHEMA_Dossier_Path & labNomRep.Caption, vbDirectory) = "" Then
            cmdVPjointes.Enabled = False
        Else
            File1.Path = paramCPT_SCHEMA_Dossier_Path & labNomRep.Caption
            File1.Refresh
            If File1.ListCount > 0 Then
                cmdVPjointes.Enabled = True
            Else
                cmdVPjointes.Enabled = False
            End If
        End If
        '                                       '
        fraUpdate_Display
   End If
End If
Me.Enabled = True
End Sub


Public Function fraUpdate_Control()
Dim blnUpdate_Control As Boolean
Dim X As String
blnUpdate_Control = True
Call lstErr_AddItem(lstErr, cmdContext, ">_________Contrôle des données "): DoEvents
newYCPTSCH0 = oldYCPTSCH0

If cmdUpdate_Ok.Caption = "Visa Compta" Then
    X = Trim(txtUpdate_CPTSCHTEXT)
    If X = "" Then
        blnUpdate_Control = False
        txtUpdate_CPTSCHTEXT.BackColor = errUsr.BackColor
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le motif")
    Else
        txtUpdate_CPTSCHTEXT.BackColor = txtUsr.BackColor
        newYCPTSCH0.CPTSCHTEXT = txtUpdate_CPTSCHTEXT
    End If
End If
If Trim(oldYCPTSCH0.CPTSCHUSR1) = "" Then
    newYCPTSCH0.CPTSCHUSR1 = usrName_UCase
    newYCPTSCH0.CPTSCHAMJ1 = DSys
    newYCPTSCH0.CPTSCHHMS1 = time_Hms
Else
    newYCPTSCH0.CPTSCHUSR2 = usrName_UCase
    newYCPTSCH0.CPTSCHAMJ2 = DSys
    newYCPTSCH0.CPTSCHHMS2 = time_Hms
End If


If blnUpdate_Control Then
    fraUpdate_Control = Null
Else
    fraUpdate_Control = "?_________fraUpdate_Control"
End If
End Function

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
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset
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
Dim meUnit As typeUnit, X As String
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), BIA_CPTSCH_Aut)

blnSetfocus = True
Form_Init
blnAuto = False

Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case "@CPT_SCHEMA":    blnAuto = True
                        'Call cbo_Scan("7 ", cboSelect_SQL)
                        cmdSelect_SQL_K = "7"
                        lstSelect_Load_7
                        Call DTPicker_Set(txtSelect_AmjMin, YBIATAB0_DATE_CPT_JP0)

                        cmdSelect_SQL_7
                        If blnSendMail Then cmdSendMail
                        Unload Me

    Case Else: blnAuto = False
                If BIA_CPTSCH_Aut.Comptabiliser Then
                    Call cbo_Scan("2 ", cboSelect_SQL)
                    cmdSelect_Ok_Click
                Else
                    If BIA_CPTSCH_Aut.Valider Then
                        Call cbo_Scan("3 ", cboSelect_SQL)
                        cmdSelect_Ok_Click
                    Else
                        Call cbo_Scan("2 ", cboSelect_SQL)
                        cmdSelect_Ok_Click

                    End If
                End If
                
                    
End Select


End Sub


Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
'    If fraUpdate.Visible _
'   And fraUpdate_B.Enabled _
'    And cmdUpdate_Ok.Enabled Then cmdUpdate_Ok_Click: Exit Sub
Else
    If currentAction = "" Then
        If SSTab1.Tab > 0 Then
            SSTab1.Tab = 0
        Else
           'SendKeys "{TAB}"
           ' cmdSelect_Click
        End If
    End If
End If
End Sub









Private Sub mnuPrint0_All_Click()
Dim I As Long, K As Long
Me.Enabled = False: Me.MousePointer = vbHourglass
    
For I = 1 To arrYCPTSCH0_Nb
    fgSelect.Row = I
    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
    xYCPTSCH0 = arrYCPTSCH0(K)
    'prtSAB_CDR_Monitor xYCPTSCH0
Next I

Me.Show

Me.Enabled = True: Me.MousePointer = 0



End Sub




Public Sub cmdPrint_Ok()
Dim K As Long, X As String, xSQL As String
Dim wMOUVEMCOM As String
lstSelect.Visible = False


End Sub
Public Sub cmdPrint_Ok_1()
Dim K As Long, X As String
Dim wIndex As Integer

fgSelect.Visible = False

fgSelect.Visible = True


End Sub


Public Function cmdUpdate_Ok_Transaction()
Dim V, X As String, xSQL As String
Dim Nb As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdUpdate_Ok_Transaction"
'-------------------------------------------------------
cmdUpdate_Ok_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
V = sqlYCPTSCH0_Update(newYCPTSCH0, oldYCPTSCH0)
If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
    cmdUpdate_Ok_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function



Public Sub lstSelect_Load_7()
Dim xSQL As String
cmdSelect_Ok_Caption = "Importer les modifications"
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
txtSelect_AmjMin.Visible = True
txtSelect_AmjMin.Enabled = True
txtSelect_AmjMax.Visible = True
txtSelect_AmjMax.Enabled = True
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True


xSQL = "select SCHEMAFDT from " & paramIBM_Library_SABSPE & ".YCPTSCH0 order by SCHEMAFDT desc"
Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    wAmjMin = "20110101"
Else
    wAmjMin = rsSab("SCHEMAFDT") + 19000000
End If

Call DTPicker_Set(txtSelect_AmjMin, wAmjMin)
Call DTPicker_Set(txtSelect_AmjMax, DSys)

End Sub

Public Sub fraUpdate_Display()
Dim X As String, xWhere As String, xSQL As String

Call lstErr_Clear(lstErr, cmdContext, ">Affichage du détail dossier"): DoEvents

xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
     & " where BASTABETA = 1 and BASTABNUM = 23 and BASTABARG = '" & xYCPTSCH0.SCHEMAOPE & "'"
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    libUpdate_OPE = Trim(rsSab("BASTABLO1")) & " : " & Mid$(rsSab("BASTABDON"), 1, 32)
Else
    libUpdate_OPE = xYCPTSCH0.SCHEMAOPE
End If


xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
     & " where BASTABETA = 1 and BASTABNUM = 24 and BASTABARG = '" & xYCPTSCH0.SCHEMAOPE & xYCPTSCH0.SCHEMAEVE & "'"
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    libUpdate_EVE = Mid$(rsSab("BASTABDON"), 13, 32)
Else
    libUpdate_EVE = xYCPTSCH0.SCHEMAEVE
End If

txtUpdate_SCHEMAOPE = xYCPTSCH0.SCHEMAOPE
txtUpdate_SCHEMAEVE = xYCPTSCH0.SCHEMAEVE
txtUpdate_SCHEMAARG = xYCPTSCH0.SCHEMAARG
txtUpdate_SCHEMAFDT = dateIBM10(xYCPTSCH0.SCHEMAFDT, True)
txtUpdate_SCHEMAFUT = arrMNURUTUTI(xYCPTSCH0.SCHEMAFUT)
txtUpdate_CPTSCHUSR1 = xYCPTSCH0.CPTSCHUSR1
If xYCPTSCH0.CPTSCHAMJ1 > 0 Then
    txtUpdate_CPTSCHAMJ1 = dateImp10(xYCPTSCH0.CPTSCHAMJ1) & " " & timeNImp8(xYCPTSCH0.CPTSCHHMS1)
Else
    txtUpdate_CPTSCHAMJ1 = ""
End If
txtUpdate_CPTSCHUSR2 = xYCPTSCH0.CPTSCHUSR2
If xYCPTSCH0.CPTSCHAMJ2 > 0 Then
    txtUpdate_CPTSCHAMJ2 = dateImp10(xYCPTSCH0.CPTSCHAMJ2) & " " & timeNImp8(xYCPTSCH0.CPTSCHHMS2)
Else
    txtUpdate_CPTSCHAMJ2 = ""
End If
txtUpdate_CPTSCHTEXT.BackColor = txtUsr.BackColor
txtUpdate_CPTSCHTEXT = Trim(xYCPTSCH0.CPTSCHTEXT)
txtUpdate_CPTSCHSTA = xYCPTSCH0.CPTSCHSTA
txtUpdate_CPTSCHUSR = xYCPTSCH0.CPTSCHUSR
fraUpdate.Visible = True
fraUpdate_A.Enabled = False
fraUpdate_B.Enabled = True
cmdUpdate_Ok.Enabled = False
cmdUpdate_Annuler.Enabled = False
txtUpdate_CPTSCHTEXT.Enabled = False
cmdPjointes.Enabled = False
If currentZMNURUT0.MNURUTCUT <> xYCPTSCH0.SCHEMAFUT Then
    If Trim(xYCPTSCH0.CPTSCHUSR2) = "" And Trim(xYCPTSCH0.CPTSCHUSR1) <> "" And BIA_CPTSCH_Aut.Valider Then
        cmdUpdate_Ok.Enabled = True
        cmdUpdate_Ok.Caption = "Visa Contrôle Comptable"
        cmdPjointes.Enabled = True
    End If
    If Trim(xYCPTSCH0.CPTSCHUSR1) = "" And BIA_CPTSCH_Aut.Comptabiliser Then
        cmdUpdate_Ok.Enabled = True
        cmdUpdate_Ok.Caption = "Visa Compta"
        txtUpdate_CPTSCHTEXT.Enabled = True
        If Left(Trim(cboSelect_SQL.Text), 1) = "2" Or Left(Trim(cboSelect_SQL.Text), 1) = "3" Then
            cmdPjointes.Enabled = True
        End If
    End If
    If Trim(xYCPTSCH0.CPTSCHUSR1) <> "" And BIA_CPTSCH_Aut.Comptabiliser Then cmdUpdate_Annuler.Enabled = True: cmdUpdate_Annuler.Caption = "Annuler Visa Compta"
    If Trim(xYCPTSCH0.CPTSCHUSR2) <> "" And BIA_CPTSCH_Aut.Valider Then cmdUpdate_Annuler.Enabled = True: cmdUpdate_Annuler.Caption = "Annuler Visa Contrôle Comptable"
End If
fraUpdate_Display_ZSCHEMAH0
End Sub







