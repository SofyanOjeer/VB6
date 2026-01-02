VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAutomate 
   Caption         =   "Automate"
   ClientHeight    =   9495
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   10920
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "Automate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   500
   End
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
      Left            =   6840
      TabIndex        =   0
      Top             =   0
      Width           =   6225
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   60
      TabIndex        =   1
      Top             =   525
      Width           =   13725
      _ExtentX        =   24209
      _ExtentY        =   15690
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Flux"
      TabPicture(0)   =   "Automate.frx":0102
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraFolder"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Demande"
      TabPicture(1)   =   "Automate.frx":011E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraMonitor"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Détail"
      TabPicture(2)   =   "Automate.frx":013A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdExe_Abort"
      Tab(2).Control(1)=   "cmdExe"
      Tab(2).Control(2)=   "fgText"
      Tab(2).Control(3)=   "libExe_Dir"
      Tab(2).Control(4)=   "libExe_File"
      Tab(2).Control(5)=   "libExe_Action"
      Tab(2).ControlCount=   6
      Begin VB.CommandButton cmdExe_Abort 
         BackColor       =   &H000000FF&
         Caption         =   "Abandonner"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   -66720
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   7920
         Width           =   2490
      End
      Begin VB.Frame fraMonitor 
         Height          =   8355
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   13530
         Begin VB.CommandButton cmdBDP_Aller_Trimestriel 
            BackColor       =   &H00C0C0FF&
            Caption         =   "BDP : Aller Trimestriel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   6480
            Width           =   1800
         End
         Begin VB.CommandButton cmdBDP_Aller_Mensuel 
            BackColor       =   &H00C0C0FF&
            Caption         =   "BDP : Aller Mensuel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   6480
            Width           =   1800
         End
         Begin VB.CommandButton cmdEIC_Retour 
            BackColor       =   &H00C0FFFF&
            Caption         =   "EIC : Retour"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   1320
            Width           =   1800
         End
         Begin VB.CommandButton cmdCDR_Retour 
            BackColor       =   &H00C0FFFF&
            Caption         =   "CDR : Retour"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   7320
            Width           =   1800
         End
         Begin VB.CommandButton cmdSIT_Aller2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "SIT : Aller 2 (SG)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   240
            Width           =   1800
         End
         Begin VB.CommandButton cmdCHQ_Retour 
            BackColor       =   &H00C0FFFF&
            Caption         =   "CHQ : Retour"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   2280
            Width           =   1800
         End
         Begin VB.CommandButton cmdCDR_Aller 
            BackColor       =   &H00C0C0FF&
            Caption         =   "CDR : Aller"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   7440
            Width           =   1800
         End
         Begin VB.CommandButton cmdCHQ_Aller 
            BackColor       =   &H00C0C0FF&
            Caption         =   "CHQ : Aller"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   2280
            Width           =   1800
         End
         Begin VB.CommandButton cmdMonitor_OK 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Demande de connexion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   10080
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   7080
            Width           =   2940
         End
         Begin VB.Frame fraMonitor_Param 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6375
            Left            =   8640
            TabIndex        =   19
            Top             =   360
            Width           =   4605
            Begin VB.TextBox txtMonitor_Hms 
               Height          =   285
               Left            =   2760
               TabIndex        =   27
               Top             =   720
               Width           =   1080
            End
            Begin VB.ComboBox cboMonitor_Appli 
               Height          =   315
               Left            =   1320
               Sorted          =   -1  'True
               TabIndex        =   26
               Text            =   "Appli"
               Top             =   1800
               Width           =   930
            End
            Begin VB.TextBox txtMonitor_IBM_Library 
               Height          =   285
               Left            =   1440
               TabIndex        =   25
               Top             =   3120
               Width           =   1005
            End
            Begin VB.TextBox txtMonitor_IBM_File 
               Height          =   285
               Left            =   2640
               TabIndex        =   24
               Top             =   3120
               Width           =   1155
            End
            Begin VB.TextBox txtMonitor_File 
               Height          =   285
               Left            =   1440
               TabIndex        =   23
               Top             =   4320
               Width           =   2430
            End
            Begin VB.TextBox txtMonitor_Exe 
               Height          =   300
               Left            =   1410
               TabIndex        =   22
               Top             =   5520
               Width           =   2430
            End
            Begin VB.OptionButton optMonitor_Aller 
               Caption         =   "Aller"
               Height          =   225
               Left            =   2400
               TabIndex        =   21
               Top             =   1800
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton optMonitor_Retour 
               Caption         =   "Retour"
               Height          =   225
               Left            =   3240
               TabIndex        =   20
               Top             =   1800
               Width           =   780
            End
            Begin MSComCtl2.DTPicker txtMonitor_Amj 
               Height          =   300
               Left            =   1320
               TabIndex        =   28
               Top             =   720
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
               Format          =   120455171
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblMonitor_Planification 
               Caption         =   "Date Heure"
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   720
               Width           =   1080
            End
            Begin VB.Label lblMonitor_Appli 
               Caption         =   "Appli Flux"
               Height          =   180
               Left            =   120
               TabIndex        =   32
               Top             =   1800
               Width           =   780
            End
            Begin VB.Label lblMonitor_IBM 
               Caption         =   "AS400"
               Height          =   255
               Left            =   135
               TabIndex        =   31
               Top             =   3120
               Width           =   690
            End
            Begin VB.Label lblMonitor_File 
               Caption         =   "Fichier Xcom"
               Height          =   240
               Left            =   240
               TabIndex        =   30
               Top             =   4320
               Width           =   1005
            End
            Begin VB.Label lblMonitor_Exe 
               Caption         =   "Script"
               Height          =   195
               Left            =   480
               TabIndex        =   29
               Top             =   5640
               Width           =   690
            End
         End
         Begin VB.CommandButton cmdSIT_Aller1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "SIT : Aller 1 (SAB)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   240
            Width           =   1800
         End
         Begin VB.CommandButton cmdSIT_Retour 
            BackColor       =   &H00C0FFFF&
            Caption         =   "SIT : Retour"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   240
            Width           =   1800
         End
         Begin VB.CommandButton cmdCB_Aller 
            BackColor       =   &H00C0C0FF&
            Caption         =   "CB : Aller"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   5520
            Width           =   1800
         End
      End
      Begin VB.CommandButton cmdExe 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Lancer le traitement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   -73440
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7920
         Width           =   2490
      End
      Begin VB.Frame fraFolder 
         Height          =   8520
         Left            =   -74940
         TabIndex        =   2
         Top             =   315
         Width           =   13560
         Begin VB.Frame fraSelect 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8400
            Left            =   60
            TabIndex        =   4
            Top             =   0
            Width           =   13400
            Begin VB.FileListBox filDoc 
               Height          =   285
               Left            =   960
               TabIndex        =   6
               Top             =   840
               Visible         =   0   'False
               Width           =   2535
            End
            Begin MSFlexGridLib.MSFlexGrid fgMonitor 
               Height          =   3405
               Left            =   360
               TabIndex        =   5
               Top             =   120
               Width           =   6450
               _ExtentX        =   11377
               _ExtentY        =   6006
               _Version        =   393216
               Rows            =   1
               Cols            =   3
               FixedCols       =   0
               RowHeightMin    =   300
               BackColor       =   14737632
               ForeColor       =   12582912
               BackColorFixed  =   12632256
               ForeColorFixed  =   -2147483641
               BackColorSel    =   12648384
               BackColorBkg    =   14737632
               AllowBigSelection=   0   'False
               TextStyleFixed  =   4
               FocusRect       =   2
               HighLight       =   0
               GridLines       =   0
               GridLinesFixed  =   1
               AllowUserResizing=   3
               FormatString    =   "<Connexions planifiées         |< Date demande                      |"
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
            Begin MSFlexGridLib.MSFlexGrid fgRetour 
               Height          =   2205
               Left            =   360
               TabIndex        =   7
               Top             =   3600
               Width           =   6405
               _ExtentX        =   11298
               _ExtentY        =   3889
               _Version        =   393216
               Rows            =   1
               Cols            =   3
               FixedCols       =   0
               RowHeightMin    =   300
               BackColor       =   14737632
               ForeColor       =   12582912
               BackColorFixed  =   12632256
               ForeColorFixed  =   -2147483641
               BackColorSel    =   12648384
               BackColorBkg    =   14737632
               AllowBigSelection=   0   'False
               TextStyleFixed  =   4
               FocusRect       =   2
               HighLight       =   0
               GridLines       =   0
               GridLinesFixed  =   1
               AllowUserResizing=   3
               FormatString    =   "<Fichiers 'RETOUR'           |< Date réception                |"
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
            Begin MSFlexGridLib.MSFlexGrid fgAller 
               Height          =   2130
               Left            =   360
               TabIndex        =   8
               Top             =   6000
               Width           =   6405
               _ExtentX        =   11298
               _ExtentY        =   3757
               _Version        =   393216
               Rows            =   1
               Cols            =   3
               FixedCols       =   0
               RowHeightMin    =   300
               BackColor       =   14737632
               ForeColor       =   12582912
               BackColorFixed  =   12632256
               ForeColorFixed  =   -2147483641
               BackColorSel    =   12648384
               BackColorBkg    =   14737632
               AllowBigSelection=   0   'False
               TextStyleFixed  =   4
               FocusRect       =   2
               HighLight       =   0
               GridLines       =   0
               GridLinesFixed  =   1
               AllowUserResizing=   3
               FormatString    =   "<Fichiers 'ALLER'           |<Date export              |"
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
            Begin MSFlexGridLib.MSFlexGrid fgYBIAMON0 
               Height          =   8070
               Left            =   7920
               TabIndex        =   37
               Top             =   210
               Width           =   5295
               _ExtentX        =   9340
               _ExtentY        =   14235
               _Version        =   393216
               Rows            =   1
               Cols            =   4
               FixedCols       =   0
               RowHeightMin    =   300
               BackColor       =   14737632
               ForeColor       =   12582912
               BackColorFixed  =   12632256
               ForeColorFixed  =   -2147483641
               BackColorSel    =   12648384
               BackColorBkg    =   14737632
               AllowBigSelection=   0   'False
               TextStyleFixed  =   4
               FocusRect       =   2
               HighLight       =   0
               GridLines       =   0
               GridLinesFixed  =   1
               AllowUserResizing=   3
               FormatString    =   "<Appli           |<Flux            |<Statut        |     "
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
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgText 
         Height          =   7200
         Left            =   -74910
         TabIndex        =   10
         Top             =   450
         Width           =   13290
         _ExtentX        =   23442
         _ExtentY        =   12700
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedCols       =   0
         RowHeightMin    =   250
         BackColor       =   14737632
         ForeColor       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12648384
         BackColorBkg    =   14737632
         AllowBigSelection=   0   'False
         TextStyleFixed  =   4
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   0
         GridLinesFixed  =   0
         AllowUserResizing=   3
         FormatString    =   $"Automate.frx":0156
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
      Begin VB.Label libExe_Dir 
         Caption         =   "-"
         Height          =   300
         Left            =   -71310
         TabIndex        =   14
         Top             =   5200
         Width           =   4770
      End
      Begin VB.Label libExe_File 
         Caption         =   "-"
         Height          =   300
         Left            =   -71370
         TabIndex        =   13
         Top             =   4800
         Width           =   4770
      End
      Begin VB.Label libExe_Action 
         Caption         =   "-"
         Height          =   315
         Left            =   -72255
         TabIndex        =   12
         Top             =   4800
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmAutomate"
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
Dim Automate_Aut As typeAuthorization

Dim IdShell

Dim blncmdOk_Run As Boolean, blnAuto_Exe As Boolean

Dim blnImportMsgFile As Boolean
Dim meYBIAMON0 As typeYBIAMON0, oldYBIAMON0 As typeYBIAMON0


Dim xIn As String
Dim xNT_Folder As String, localNT_Folder As String
Dim xFile_sab As String
Dim xFile_TXT As String
Dim xFile_Archive As String
Dim xFile_XCom As String, localFile_XCom As String
Dim wFile_Log As String, localFile_Log As String
Dim wFile_Monitor As String
Dim blnFile_FTP_Open As Boolean
Dim xAppli As String, xFlux As String
Dim xIBM_File As String, xNT_File As String, xIBM_Library As String

Dim Nb As Long
Dim wAMJHMS As String

Dim paramCB_Aller_Path As String, constCB_Aller As String
Dim paramBDF_Archive_Path As String
Dim paramBDP_Aller_Path As String, constBDP_Aller As String

Dim appliFonction As String
Private Sub cmdMonitor_Control()
Dim V, X As String

Call lstErr_Clear(lstErr, cmdContext, xAppli & " : Demande de connexion")

cmdMonitor_Init

blncmdOk_Run = True

If xAppli = "" Then
     V = "? Préciser l'application "
    GoTo Exit_sub
End If

If xIBM_Library = "" Then
     V = "? Préciser la Librairie IBM  "
    GoTo Exit_sub
End If

If xIBM_File = "" Then
     V = "? Préciser le fichier IBM  "
    GoTo Exit_sub
End If

If xNT_File = "" Then
    V = "? Préciser le fichier NT  "
    GoTo Exit_sub
End If


If optMonitor_Aller Then
    V = cmdMonitor_Aller_FTP
    If IsNull(V) Then
        Select Case xAppli
            'Case "CB": V = cmdMonitor_Aller_CB
            Case "CHQ": V = cmdMonitor_Aller_CHQ
            Case "CDR": V = cmdMonitor_Aller_CDR
            Case "BDP": V = cmdMonitor_Aller_BDP
           'Case "ECH": V = cmdMonitor_Aller_ECH
            Case "SIT": V = cmdMonitor_Aller_SIT
        End Select
    End If
Else
    V = cmdMonitor_Retour_Gen_Bat
End If

Exit_sub:
    Call lstErr_AddItem(lstErr, cmdContext, xAppli & " " & V)
    
    fgMonitor_Load
    fgRetour_Load
    fgAller_Load
    cmdMonitor_OK.Enabled = False
    blncmdOk_Run = False

End Sub


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim X As String
'20040830 jpl :DTAQ à supprimer $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'=====================================
'If elpSrvXcom = "XXXX" Then
'    elpSrvXcom = "CAV4"
'    If Not IsNull(SndRcv_Init) Then End
'End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), Automate_Aut)    '
SSTab1.Tab = 0
fraFolder.Enabled = False

blnImportMsgFile = False

'_________________

'_________________________________________________________________________________
constCB_Aller = "CB_Aller"
paramCB_Aller_Path = paramServer("\\BDF_Aller\BAFI\")
paramBDF_Archive_Path = Replace(paramCB_Aller_Path, "BAFI", "Archive")
appliFonction = Trim(Mid$(Msg, 1, 12))
If appliFonction = "*=>BDF_CB" Then
    fraSelect.Enabled = False
    fraMonitor.Enabled = False
    SSTab1.Tab = 2: cmdMonitor_Aller_CB_Init
End If
paramBDP_Aller_Path = Replace(paramCB_Aller_Path, "BAFI", "BDP")
constBDP_Aller = "BDP_Aller"
If appliFonction = "*=>BDF_BDP" Then
    fraSelect.Enabled = False
    fraMonitor.Enabled = False
    SSTab1.Tab = 2: cmdMonitor_Aller_BDP_Init
End If
'_________________________________________________________________________________
If Automate_Aut.Valider Then
    fraFolder.Enabled = True
    fraMonitor_Reset
    'paramSAA_Init
    fgMonitor.Enabled = Automate_Aut.Xspécial
    fraMonitor_Param.Enabled = Automate_Aut.Xspécial
            
   ''' If blnOff_Line Then paramSAA_Init_Test
    
    
    fgRetour_Load
    fgMonitor_Load
    fgAller_Load
    
     Call lstErr_Clear(lstErr, cmdContext, "Paramètres chargés")
   
    Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
        Case "$Auto_Exe":     blnAuto_Exe = True:    Auto_Exe
        Case Else: blnAuto_Exe = False
    End Select
End If

End Sub


Public Sub cmdContext_Quit()
    If blnMsgBox_Quit Then
       X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
    Else
       X = vbYes
    End If
    If X = vbYes Then Unload Me

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

Private Sub cmdBDP_Aller_Mensuel_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call DTPicker_Set(txtMonitor_Amj, DSys)
txtMonitor_Hms = Time
cboMonitor_Appli.Text = "BDP"
optMonitor_Aller = True
txtMonitor_IBM_Library = paramIBM_Library_SAB
txtMonitor_IBM_File = "ZREPBPM0"
txtMonitor_File = "12179.bdp"
txtMonitor_Exe = ""

cmdMonitor_OK.Enabled = True
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdBDP_Aller_Trimestriel_Click()
cmdMonitor_Aller_BDP_Init

End Sub

Private Sub cmdCB_Aller_Click()
cmdMonitor_Aller_CB_Init
End Sub

Private Sub cmdMonitor_Aller_CB_Init()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call DTPicker_Set(txtMonitor_Amj, DSys)
txtMonitor_Hms = Time
cboMonitor_Appli.Text = "CB"
optMonitor_Aller = True
txtMonitor_IBM_Library = paramIBM_Library_SAB
txtMonitor_IBM_File = "ZREPSIT0"
txtMonitor_File = "Situ.cb"
txtMonitor_Exe = ""
blnFile_FTP_Open = False
cmdMonitor_Aller_CB
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdCDR_Aller_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call DTPicker_Set(txtMonitor_Amj, DSys)
txtMonitor_Hms = Time
cboMonitor_Appli.Text = "CDR"
optMonitor_Aller = True
txtMonitor_IBM_Library = paramIBM_Library_SAB
txtMonitor_IBM_File = "ZCRIDEC0"
txtMonitor_File = "CDR_Aller.txt"
txtMonitor_Exe = "CDR_Aller.cmd"

cmdMonitor_OK.Enabled = True
'MsgBox "JPL bricolage", vbCritical
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Function cmdMonitor_Aller_CDR()
Dim x240 As String * 240

On Error GoTo Error_Handler
cmdMonitor_Aller_CDR = "? cmdMonitor_Aller_CDR"

Call lstErr_AddItem(lstErr, cmdContext, "cmdCDR_Aller : Format 240")
Open xFile_sab For Input As #1

Nb = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Len(xIn) > 0 Then
        If Not blnFile_FTP_Open Then
            Call FEU_ROUGE
            Open xFile_TXT For Output As #2
            blnFile_FTP_Open = True
        End If

        Nb = Nb + 1
        x240 = xIn
        Print #2, x240
    End If
Loop
Close
Call FEU_VERT

Call lstErr_AddItem(lstErr, cmdContext, "cmdCDR_Aller : nb = " & Nb)

X = MsgBox("Confirmez-vous l'envoi de " & Nb & " enregistrements ?" & Asc10_13 & Trim(x240), vbYesNo + vbQuestion + vbDefaultButton2, "Déclaration CDR  ")

If X = vbYes Then
    cmdMonitor_Aller_CDR = cmdMonitor_Aller_Gen_Bat
Else
    cmdMonitor_Aller_CDR = "Abandon"

End If
Exit Function

Error_Handler:
cmdMonitor_Aller_CDR = Error
Shell_MsgBox "cmdCDR_Aller " & Error, vbCritical, "frmAutomate", False


End Function
Private Function cmdMonitor_Aller_CB()
Dim x118 As String * 118
Dim xSQL As String

On Error GoTo Error_Handler
cmdMonitor_Aller_CB = "? cmdMonitor_Aller_CB"

Call lstErr_AddItem(lstErr, cmdContext, "cmdCB_Aller : Format 118")
Set rsSab = Nothing
Nb = 0

xSQL = "select * from " & paramIBM_Library_SAB & ".ZREPSIT0 "
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    If Not blnFile_FTP_Open Then
        Open paramCB_Aller_Path & "situ.sab" For Output As #2
        Call FEU_ROUGE
        blnFile_FTP_Open = True
    End If
    Nb = Nb + 1
    x118 = rsSab("REPSITMV") & rsSab("REPSITDAT") & rsSab("REPSITCIB") _
            & rsSab("REPSITLC") & rsSab("REPSITDOC") & rsSab("REPSITFE") _
            & rsSab("REPSITZA") & rsSab("REPSITMO") & rsSab("REPSITFIL")
        Print #2, x118
'REPSITMV
'REPSITDAT
'REPSITCIB
'REPSITLC
'REPSITDOC
'REPSITFE
'REPSITZA
'REPSITMO
'REPSITFIL
'_____________________________________________________________________________________________

   rsSab.MoveNext

Loop

Close
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "cmdCB_Aller : nb = " & Nb)
fgText_Display constCB_Aller, paramCB_Aller_Path, "situ.sab"

Exit Function

Error_Handler:
cmdMonitor_Aller_CB = Error
Shell_MsgBox "cmdCB_Aller " & Error, vbCritical, "frmAutomate", False


End Function

Private Function cmdMonitor_Aller_BDP()
Dim x200 As String * 200
Dim xSQL As String

On Error GoTo Error_Handler
cmdMonitor_Aller_BDP = "? cmdMonitor_Aller_BDP"

Call lstErr_AddItem(lstErr, cmdContext, "cmdBDP_Aller : Format 200")
Set rsSab = Nothing
Nb = 0

xSQL = "select * from " & paramIBM_Library_SAB & ".ZREPBPT0 "
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    If Not blnFile_FTP_Open Then
        Open paramBDP_Aller_Path & "12179.sab" For Output As #2
        Call FEU_ROUGE

        blnFile_FTP_Open = True
    End If
    Nb = Nb + 1
    x200 = rsSab("REPBPTCOD") & rsSab("REPBPTCDM") & rsSab("REPBPTDDA") _
         & rsSab("REPBPTRES") & rsSab("REPBPTCDI") & rsSab("REPBPTCMD") _
         & rsSab("REPBPTNCL") & rsSab("REPBPTSES") & rsSab("REPBPTMTO") _
         & rsSab("REPBPTCDP") & rsSab("REPBPTDIS") & rsSab("REPBPTCDC")
        Print #2, x200
'_____________________________________________________________________________________________

   rsSab.MoveNext

Loop

Close
Call FEU_VERT

Call lstErr_AddItem(lstErr, cmdContext, "cmdBDP_Aller : nb = " & Nb)
fgText_Display constBDP_Aller, paramBDP_Aller_Path, "12179.sab"

Exit Function

Error_Handler:
cmdMonitor_Aller_BDP = Error
Shell_MsgBox "cmdBDP_Aller " & Error, vbCritical, "frmAutomate", False


End Function


Private Function cmdMonitor_Aller_SIT()
Dim wFile As String
Dim V
On Error GoTo Error_Handler
cmdMonitor_Aller_SIT = "? cmdMonitor_Aller_SIT"



Call lstErr_AddItem(lstErr, cmdContext, "cmdSIT_Aller : Format 240")
Open xFile_sab For Input As #1

Nb = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Len(xIn) > 0 Then
        If Not blnFile_FTP_Open Then
            Open xFile_TXT For Output As #2
            Call FEU_ROUGE

            blnFile_FTP_Open = True
        End If

        Nb = Nb + 1
        Print #2, xIn
    End If
Loop
Close
Call FEU_VERT

Call lstErr_AddItem(lstErr, cmdContext, "cmdSIT_Aller : nb = " & Nb)

X = MsgBox("Confirmez-vous l'envoi de " & Nb & " enregistrements ?" & Asc10_13 & Trim(xIn), vbYesNo + vbQuestion + vbDefaultButton2, "Déclaration SIT  ")

If X = vbYes Then
    cmdMonitor_Aller_SIT = cmdMonitor_Aller_Gen_Bat
'------------------------------------------------------------------------------------
    meYBIAMON0.MONAPP = xAppli
    meYBIAMON0.MONFLUX = UCase$(xFlux)
    V = rsYBIAMON0_Read(meYBIAMON0)
    If Not IsNull(V) Then GoTo Error_MsgBox
    If Trim(meYBIAMON0.MONSTATUS) <> "" Then
        V = "Action précédente en cours : " & meYBIAMON0.MONAPP & "_" & meYBIAMON0.MONFLUX & " > " & meYBIAMON0.MONSTATUS
        GoTo Error_MsgBox
    End If
'------------------------------------------------------------------------------------
    oldYBIAMON0 = meYBIAMON0
    meYBIAMON0.MONSTATUS = ""
    V = sqlYBIAMON0_Update(meYBIAMON0, oldYBIAMON0, True)
    If Not IsNull(V) Then GoTo Error_MsgBox
    
Else
    cmdMonitor_Aller_SIT = "Abandon"

End If
Exit Function

Error_Handler:
V = Error
Error_MsgBox:
cmdMonitor_Aller_SIT = V
Shell_MsgBox "cmdMonitor_Retour_SAB " & V, vbCritical, "frmAutomate", False


End Function







Private Sub cmdCDR_Retour_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call DTPicker_Set(txtMonitor_Amj, DSys)
txtMonitor_Hms = Time
cboMonitor_Appli.Text = "CDR"
optMonitor_Retour = True
txtMonitor_IBM_Library = paramIBM_Library_SAB
txtMonitor_IBM_File = "ZCRICEN0"
txtMonitor_File = "CDR_Retour.txt"
txtMonitor_Exe = "CDR_Retour.cmd"

cmdMonitor_OK.Enabled = True

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdCHQ_Aller_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
Call DTPicker_Set(txtMonitor_Amj, DSys)
txtMonitor_Hms = Time
cboMonitor_Appli.Text = "CHQ"
optMonitor_Aller = True
txtMonitor_IBM_Library = paramIBM_Library_SAB
txtMonitor_IBM_File = "SABCHQA"
txtMonitor_File = "CHQ_Aller.txt"
txtMonitor_Exe = "CHQ_Aller.cmd"

cmdMonitor_OK.Enabled = True
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Function cmdMonitor_Aller_CHQ()
Dim x513 As String * 513

On Error GoTo Error_Handler
cmdMonitor_Aller_CHQ = "?cmdMonitor_Aller_CHQ"

Call lstErr_AddItem(lstErr, cmdContext, "cmdCHQ_Aller : Format 513")
Open xFile_sab For Input As #1

Nb = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Len(xIn) > 0 Then
        If Not blnFile_FTP_Open Then
            Open xFile_TXT For Output As #2
            Call FEU_ROUGE

            blnFile_FTP_Open = True
        End If

        Nb = Nb + 1
        x513 = xIn
        Print #2, x513
    End If
Loop
Close
Call FEU_VERT

Call lstErr_AddItem(lstErr, cmdContext, "cmdCHQ_Aller : nb = " & Nb)

X = MsgBox("Confirmez-vous l'envoi de " & Nb & " enregistrements ?" & Asc10_13 & Trim(x513), vbYesNo + vbQuestion + vbDefaultButton2, "Commande Chèquiers ")

If X = vbYes Then
    cmdMonitor_Aller_CHQ = cmdMonitor_Aller_Gen_Bat
Else
    cmdMonitor_Aller_CHQ = "Abandon"

End If

Exit Function

Error_Handler:
cmdMonitor_Aller_CHQ = Error
Shell_MsgBox "cmdMonitor_Aller_CHQ " & Error, vbCritical, "frmAutomate", False


End Function
Private Function cmdMonitor_Aller_FTP()

On Error GoTo Error_Handler

cmdMonitor_Aller_FTP = "? cmdMonitor_Aller_FTP"
blnFile_FTP_Open = False

If Dir(xFile_XCom) <> "" Then
    cmdMonitor_Aller_FTP = "? le fichier existe déjà : " & xNT_File
    Exit Function
End If

Call lstErr_Clear(lstErr, cmdContext, "Aller : Ftp .....")
'If Trim(xAppli) <> "CB" Then
    If Not blnOff_Line Then Call Shell_FTP(xFile_sab, xIBM_Library, xIBM_File, True, False)
'End If

Call lstErr_AddItem(lstErr, cmdContext, "Aller : Ftp terminé")

cmdMonitor_Aller_FTP = Null
Exit Function

Error_Handler:
Shell_MsgBox "cmdMonitor_Aller_FTP " & Error, vbCritical, "frmAutomate", False
cmdMonitor_Aller_FTP = "? " & Error
End Function






Private Sub cmdCHQ_Retour_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call DTPicker_Set(txtMonitor_Amj, DSys)
txtMonitor_Hms = Time
cboMonitor_Appli.Text = "CHQ"
optMonitor_Retour = True
txtMonitor_IBM_Library = paramIBM_Library_SABSPE
txtMonitor_IBM_File = "BIACHQR"
txtMonitor_File = "CHQ_Retour.txt"
txtMonitor_Exe = "CHQ_Retour.cmd"

cmdMonitor_OK.Enabled = True

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdContext_Click()
cmdContext_Quit

End Sub

Private Sub cmdEIC_Retour_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call DTPicker_Set(txtMonitor_Amj, DSys)
txtMonitor_Hms = Time
cboMonitor_Appli.Text = "EIC"
optMonitor_Retour = True
txtMonitor_IBM_Library = paramIBM_Library_SABSPE
txtMonitor_IBM_File = "BIAEICRW"
txtMonitor_File = "EIC_Retour.txt"
txtMonitor_Exe = "EIC_Retour.cmd"

cmdMonitor_OK.Enabled = True
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdExe_Abort_Click()
If appliFonction = "*=>BDF_CB" Then Unload Me
If appliFonction = "*=>BDF_BDP" Then Unload Me
fgText.Clear
SSTab1.Tab = 1
End Sub

Private Sub cmdExe_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

If libExe_Action.Caption = constCB_Aller Then cmdMonitor_Aller_CB_Ok: GoTo Exit_sub
If libExe_Action.Caption = constBDP_Aller Then cmdMonitor_Aller_BDP_Ok: GoTo Exit_sub

If libExe_Action = constXCom Then cmdMonitor_XCom_Exe libExe_File
If libExe_Action = constRetour Then cmdMonitor_Retour_SAB libExe_File
fgMonitor_Load
fgRetour_Load
fgAller_Load

Exit_sub:

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdMonitor_OK_Click()

Call lstErr_Clear(lstErr, cmdContext, xAppli & " : Demande de connexion")

If Not Me.Enabled Then Exit Sub
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdMonitor_Control


Me.Show
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdSIT_Aller1_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call DTPicker_Set(txtMonitor_Amj, DSys)
txtMonitor_Hms = Time
cboMonitor_Appli.Text = "SIT"
optMonitor_Aller = True
txtMonitor_IBM_Library = paramIBM_Library_SAB
txtMonitor_IBM_File = "ZSIT0110"
txtMonitor_File = "SIT_Aller.txt"
txtMonitor_Exe = "SIT_Aller.cmd"

cmdMonitor_Init
cmdMonitor_Aller_SAB

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSIT_Aller2_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call DTPicker_Set(txtMonitor_Amj, DSys)
txtMonitor_Hms = Time
cboMonitor_Appli.Text = "SIT"
optMonitor_Aller = True
txtMonitor_IBM_Library = paramIBM_Library_SABSPE
txtMonitor_IBM_File = "ZSIT0110"
txtMonitor_File = "SIT_Aller.txt"
txtMonitor_Exe = "SIT_Aller.cmd"

cmdMonitor_Init
MsgBox "à revoir JPL :cmdSIT_Aller2_Click ", vbCritical, Me.Caption
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'recYBIAMON0_Init meYBIAMON0
'meYBIAMON0.MONAPP = xAppli
'meYBIAMON0.MONFLUX = UCase$(xFlux)

'meYBIAMON0.MONUSR = usrId
'meYBIAMON0.MONAMJ = DSys
'meYBIAMON0.MONHMS = time_Hms
'If IsNull(srvYBIAMON0_Monitor(meYBIAMON0)) Then


'    If Trim(meYBIAMON0.MONSTATUS) = "OUT_FTP" Then
'            txtMonitor_IBM_File = Trim(meYBIAMON0.MONFILE)
'            txtMonitor_File = "SIT_" & Trim(meYBIAMON0.MONFILE) & ".txt"
'            cmdMonitor_Control
'    End If
'End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Me.Enabled = True: Me.MousePointer = 0


End Sub


Private Sub cmdSIT_Retour_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call DTPicker_Set(txtMonitor_Amj, DSys)
txtMonitor_Hms = Time
cboMonitor_Appli.Text = "SIT"
optMonitor_Retour = True
txtMonitor_IBM_Library = paramIBM_Library_SABSPE
txtMonitor_IBM_File = "BIASITRW"
txtMonitor_File = "SIT_Retour.txt"
txtMonitor_Exe = "SIT_Retour.cmd"

cmdMonitor_OK.Enabled = True
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub fgMonitor_LostFocus()
fgMonitor.LeftCol = 0

End Sub

Private Sub fgMonitor_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim xId As String
Dim V

On Error Resume Next
 If fgMonitor.Rows > 1 Then
     'Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
     fgMonitor.Col = 0
        fgText_Display constXCom, paramPeliNT_MonitorF, fgMonitor.Text
End If

End Sub

Private Sub fgRetour_LostFocus()
fgRetour.LeftCol = 0

End Sub


Private Sub fgRetour_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim xId As String
Dim V
Dim K As Integer

On Error Resume Next
If fgRetour.Rows > 1 Then
    '' Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    fgRetour.Col = 0
    xFlux = UCase$(constRetour)
    K = InStr(1, fgRetour.Text, "_")
    xAppli = Mid$(fgRetour.Text, 1, K - 1)
    fgText_Display constRetour, paramPeliNT_Retour_XcomF, fgRetour.Text
End If

End Sub


Private Sub fgAller_LostFocus()
fgAller.LeftCol = 0

End Sub

Private Sub fgAller_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim xId As String
Dim V
Dim K As Integer

On Error Resume Next
If fgAller.Rows > 1 Then
    '' Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    fgAller.Col = 0
    K = InStr(1, fgAller.Text, "_")
    xAppli = Mid$(fgAller.Text, 1, K - 1)
    xFlux = UCase$(constAller)
    fgText_Display constAller, paramPeliNT_Aller_XcomF, fgAller.Text
End If

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0:  cmdContext_Return
    Case Is = 27:  cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub

Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)

''fraMT950.Enabled = False

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub fraFolder_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
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



Public Sub fgMonitor_Load()
Dim I As Integer, K As Integer, X As String, L As Integer, iSession As Integer
On Error GoTo Error_Handler


filDoc.Pattern = "x.xxx"
filDoc.PATH = paramPeliNT_Monitor
filDoc.Pattern = "*.*"
fgMonitor.Redraw = False
fgMonitor.Rows = 1
fgMonitor.Enabled = True
For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.PATH & "\" & filDoc.FileName)
    fgMonitor.Rows = fgMonitor.Rows + 1
    fgMonitor.Row = fgMonitor.Rows - 1
    fgMonitor.Col = 0: fgMonitor.Text = Trim(filDoc.FileName)
    fgMonitor.Col = 1: fgMonitor.Text = msFile.DateLastModified
Next I
fgMonitor.Redraw = True

Exit Sub

Error_Handler:
Shell_MsgBox "fgMonitor_Load " & paramPeliNT_Monitor & " : " & Error, vbCritical, Me.Caption, False


End Sub
Public Sub fgRetour_Load()
Dim I As Integer, K As Integer, X As String, L As Integer, iSession As Integer
On Error GoTo Error_Handler

filDoc.Pattern = "x.xxx"
filDoc.PATH = paramPeliNT_DataF & constRetour & "\" & constXCom
filDoc.Pattern = "*.*"
fgRetour.Redraw = False
fgRetour.Rows = 1
fgRetour.Enabled = True
For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.PATH & "\" & filDoc.FileName)
    fgRetour.Rows = fgRetour.Rows + 1
    fgRetour.Row = fgRetour.Rows - 1
    fgRetour.Col = 0: fgRetour.Text = Trim(filDoc.FileName)
    fgRetour.Col = 1: fgRetour.Text = msFile.DateLastModified
Next I
fgRetour.Redraw = True
Exit Sub

Error_Handler:
Shell_MsgBox "fgRetour_Load " & filDoc.PATH & " : " & Error, vbCritical, Me.Caption, False

End Sub

Public Sub fgAller_Load()
Dim I As Integer, K As Integer, X As String, L As Integer, iSession As Integer
On Error GoTo Error_Handler


filDoc.Pattern = "x.xxx"
filDoc.PATH = paramPeliNT_DataF & constAller & "\" & constXCom
filDoc.Pattern = "*.*"

fgAller.Redraw = False
fgAller.Rows = 1
fgAller.Enabled = True
For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.PATH & "\" & filDoc.FileName)
    fgAller.Rows = fgAller.Rows + 1
    fgAller.Row = fgAller.Rows - 1
    fgAller.Col = 0: fgAller.Text = Trim(filDoc.FileName)
    fgAller.Col = 1: fgAller.Text = msFile.DateLastModified
Next I
fgAller.Redraw = True

Exit Sub

Error_Handler:
Shell_MsgBox "fgAller_Load " & paramFTP_Out_XCom & " : " & Error, vbCritical, Me.Caption, False

End Sub



Public Sub Auto_Exe()
Dim blnOk As Boolean

blncmdOk_Run = False
Do
    If blncmdOk_Run = False Then
        blnOk = True
    
   '''     If cmdOk_SAA_to_Corona.Enabled Then blnOk = False: cmdOk_SAA_to_Corona_Click
   '''     If cmdOK.Enabled Then blnOk = False: cmdOK_Click
   '''     If filDoc.ListCount > 0 Then blnOk = False: cmdOK_SAA_from_SAB_Click
        DoEvents
    End If
Loop Until blnOk = True
Unload Me

End Sub


Public Sub fraMonitor_Reset()

cmdMonitor_OK.Enabled = False
fraMonitor.BackColor = &HF0FFFF '
usrColor_Container fraMonitor, fraMonitor.BackColor
fraMonitor_Param.BackColor = &HF0FFFF '
usrColor_Container fraMonitor_Param, fraMonitor_Param.BackColor
fraMonitor.Enabled = Automate_Aut.Valider

Call DTPicker_Set(txtMonitor_Amj, DSys)
txtMonitor_Hms = Time
cboMonitor_Appli.Clear
cboMonitor_Appli.AddItem "CDR"
cboMonitor_Appli.AddItem "CHQ"
optMonitor_Aller = True
txtMonitor_IBM_Library = paramIBM_Library_SAB
txtMonitor_IBM_File = ""
txtMonitor_File = ""
txtMonitor_Exe = ""
End Sub

Public Function cmdMonitor_Aller_Gen_Bat()
Dim X As String, xBat As String, xLog As String
Dim K As Integer
On Error GoTo Error_Handler

cmdMonitor_Aller_Gen_Bat = "? cmdMonitor_Aller_Gen_Bat"

Open paramPeliNT_MonitorF & wFile_Monitor For Output As #3
Call FEU_ROUGE
Print #3, "REM demande    " & wAMJHMS & Trim(usrId) & "_" & paramEnvironnement
Print #3, "REM planifiée  " & wFile_Monitor
Print #3, "REM répertoire " & paramPeliNT_MonitorF
Print #3, "REM ....................................................................."
X = "CMD/C " & localPeliNT_ExeF & Trim(txtMonitor_Exe) & "  " & localFile_XCom & " >> " & localFile_Log
Print #3, "cd /d " & localPeliNT_ExeF
Print #3, "Date/T"
Print #3, "Time/T"

If paramEnvironnement = constProduction Then
    Print #3, X
Else
    Print #3, "REM " & X
    Print #3, "Del " & localFile_XCom
End If
Print #3, "REM ....................................................................."
'$jpl 20091013 Print #3, "del \\pelisrv\PELINT.DAT\Logs\BIA_Transfert_En_Cours.log"
Print #3, "del " & paramServer("\\PELINT\") & "Logs\BIA_Transfert_En_Cours.log"

Close

msFileSystem.MoveFile xFile_sab, xFile_Archive
msFileSystem.MoveFile xFile_TXT, xFile_XCom
Call FEU_VERT
cmdMonitor_Aller_Gen_Bat = Null

Exit Function

Error_Handler:
Shell_MsgBox "cmdMonitor_Aller_Gen_Bat " & Error, vbCritical, "frmAutomate", False
cmdMonitor_Aller_Gen_Bat = "? " & Error

End Function
Public Function cmdMonitor_Retour_Gen_Bat()
Dim X As String, xBat As String, xLog As String
Dim K As Integer
On Error GoTo Error_Handler

cmdMonitor_Retour_Gen_Bat = "? cmdMonitor_Retour_Gen_Bat"
Call FEU_ROUGE

Open paramPeliNT_MonitorF & wFile_Monitor For Output As #3
Print #3, "REM demande    " & wAMJHMS & Trim(usrId) & "_" & paramEnvironnement
Print #3, "REM planifiée  " & wFile_Monitor
Print #3, "REM répertoire " & paramPeliNT_MonitorF
Print #3, "REM ....................................................................."
X = "CMD/C " & localPeliNT_ExeF & Trim(txtMonitor_Exe) & "  " & localFile_XCom & " >> " & localFile_Log
Print #3, "cd /d " & localPeliNT_ExeF
Print #3, "Date/T"
Print #3, "Time/T"

If paramEnvironnement = constProduction Then
    Print #3, X
Else
    Print #3, "REM " & X
End If
Print #3, "REM ....................................................................."
'$jpl 20091013 Print #3, "del \\pelisrv\PELINT.DAT\Logs\BIA_Transfert_En_Cours.log"
Print #3, "del " & paramServer("\\PELINT\") & "Logs\BIA_Transfert_En_Cours.log"
Close
Call FEU_VERT

cmdMonitor_Retour_Gen_Bat = Null

Exit Function

Error_Handler:
Shell_MsgBox "cmdMonitor_Retour_Gen_Bat " & Error, vbCritical, "frmAutomate", False
cmdMonitor_Retour_Gen_Bat = "? " & Error

End Function


Private Sub txtMonitor_Exe_GotFocus()
txt_GotFocus txtMonitor_Exe
End Sub


Private Sub txtMonitor_Exe_LostFocus()
txt_LostFocus txtMonitor_Exe
End Sub


Private Sub txtMonitor_File_GotFocus()
txt_GotFocus txtMonitor_File
End Sub


Private Sub txtMonitor_File_LostFocus()
txt_LostFocus txtMonitor_File
End Sub


Private Sub txtMonitor_Hms_GotFocus()
txt_GotFocus txtMonitor_Hms
End Sub

Private Sub txtMonitor_Hms_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)
End Sub


Private Sub txtMonitor_Hms_LostFocus()
txt_LostFocus txtMonitor_Hms
End Sub

Private Sub txtMonitor_IBM_File_GotFocus()
txt_GotFocus txtMonitor_Exe
End Sub

Private Sub txtMonitor_IBM_File_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtMonitor_IBM_File_LostFocus()
txt_LostFocus txtMonitor_IBM_File
End Sub

Private Sub txtMonitor_IBM_Library_GotFocus()
txt_GotFocus txtMonitor_IBM_Library
End Sub

Private Sub txtMonitor_IBM_Library_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtMonitor_IBM_Library_LostFocus()
txt_LostFocus txtMonitor_IBM_Library

End Sub



Public Sub cmdMonitor_XCom_Exe(lFile As String)
Dim IdShell
Dim xBat As String, xLog As String, X As String
Dim K As Integer
On Error GoTo Error_Handler

X = localPeliNT_MonitorF & lFile
xBat = localPeliNT_MonitorF & "Temp\" & lFile
msFileSystem.MoveFile X, xBat

xLog = xBat
K = InStr(1, xLog, ".bat")
Mid$(xLog, K, 4) = ".log"
X = MsgBox("Voulez lancer le traitement ?", vbYesNo + vbQuestion + vbDefaultButton2, paramPeliNT_MonitorF & lFile)

If X = vbYes Then
    X = xBat & " >> " & xLog
    IdShell = Shell(X, 1)
    DoEvents
    If IdShell > 0 Then
        ''AppActivate IdShell, True
        ''Kill xBat
    End If
End If
Exit Sub

Error_Handler:
Shell_MsgBox "cmdMonitor_Bat " & Error, vbCritical, "frmAutomate", False

End Sub

Public Sub cmdMonitor_Retour_SAB(lFile As String)
Dim IdShell, V
Dim X As String, wFTP_File As String
Dim K As Integer
Dim wIBM_Library As String
On Error GoTo Error_Handler
V = ""

X = MsgBox("Voulez lancer le traitement ?", vbYesNo + vbQuestion + vbDefaultButton2, xAppli & " " & lFile)

If X = vbYes Then
    If xAppli = "CDR" Then
        wIBM_Library = paramIBM_Library_SAB
'20050727 à revoir Journalisation YBIAMON*
'========================================
        wFTP_File = paramPeliNT_Retour_FTPF & lFile
        msFileSystem.MoveFile paramPeliNT_Retour_XcomF & lFile, wFTP_File
        Call Shell_FTP(paramPeliNT_Retour_FTPF & lFile, wIBM_Library, "ZCRICEN0", False, False)
        If Dir(wFTP_File) <> "" Then msFileSystem.DeleteFile wFTP_File
        Exit Sub
    Else
        wIBM_Library = paramIBM_Library_SABSPE
    End If
'------------------------------------------------------------------------------------
    meYBIAMON0.MONAPP = xAppli
    meYBIAMON0.MONFLUX = UCase$(xFlux)
    V = rsYBIAMON0_Read(meYBIAMON0)
    If Not IsNull(V) Then GoTo Error_MsgBox
    If Trim(meYBIAMON0.MONSTATUS) <> "" Then
        V = "Action précédente en cours : " & meYBIAMON0.MONAPP & "_" & meYBIAMON0.MONFLUX & " > " & meYBIAMON0.MONSTATUS
        GoTo Error_MsgBox
    End If
'------------------------------------------------------------------------------------
    oldYBIAMON0 = meYBIAMON0
    meYBIAMON0.MONSTATUS = "FTP"
    V = sqlYBIAMON0_Update(meYBIAMON0, oldYBIAMON0, True)
    If Not IsNull(V) Then GoTo Error_MsgBox
    
    wFTP_File = paramPeliNT_Retour_FTPF & lFile
    msFileSystem.MoveFile paramPeliNT_Retour_XcomF & lFile, wFTP_File
    
    Call Shell_FTP(paramPeliNT_Retour_FTPF & lFile, wIBM_Library, meYBIAMON0.MONFILE, False, False)
'------------------------------------------------------------------------------------
    
    oldYBIAMON0 = meYBIAMON0
    If Trim(meYBIAMON0.MONPGM) <> "" Then
        meYBIAMON0.MONSTATUS = "IN_OK"
    Else
        meYBIAMON0.MONSTATUS = ""
    End If
    V = sqlYBIAMON0_Update(meYBIAMON0, oldYBIAMON0, True)
    If Not IsNull(V) Then GoTo Error_MsgBox
    
    If Dir(wFTP_File) <> "" Then msFileSystem.DeleteFile wFTP_File
End If
Exit Sub

Error_Handler:
V = Error
Error_MsgBox:
Shell_MsgBox "cmdMonitor_Retour_SAB " & V, vbCritical, "frmAutomate", False

End Sub


Public Sub cmdMonitor_Aller_SAB()
Dim IdShell
Dim X As String, wFTP_File As String
Dim K As Integer
On Error GoTo Error_Handler
Call lstErr_Clear(lstErr, cmdContext, "> Aller_SAB : " & xAppli)
X = MsgBox("Voulez lancer le traitement ALLER ?", vbYesNo + vbQuestion + vbDefaultButton2, xAppli)

If X = vbYes Then
    Call lstErr_AddItem(lstErr, cmdContext, "> Aller_SAB : contrôle")
MsgBox "à revoir JPL :cmdMonitor_Aller_SAB ", vbCritical, Me.Caption
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
'    recYBIAMON0_Init meYBIAMON0
'    meYBIAMON0.MONAPP = xAppli
'    meYBIAMON0.MONFLUX = UCase$(xFlux)

'    meYBIAMON0.Method = "SBMJOB"
'    meYBIAMON0.MONUSR = paramIBM_QSYSOPR
'    meYBIAMON0.MONAMJ = DSys
'    meYBIAMON0.MONHMS = time_Hms
'    If IsNull(srvYBIAMON0_Monitor(meYBIAMON0)) Then
'       If Trim(meYBIAMON0.MONSTATUS) <> "SBMJOB" Then
'            Call lstErr_AddItem(lstErr, cmdContext, "? Aller_SAB : " & meYBIAMON0.MONSTATUS)
'        Else
'            Call lstErr_AddItem(lstErr, cmdContext, "+ Aller_SAB : " & meYBIAMON0.MONSTATUS)
'        End If
        
'    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, "= Aller_SAB : fin")
Exit Sub

Error_Handler:
Shell_MsgBox "cmdMonitor_Retour_SAB " & Error, vbCritical, "frmAutomate", False

End Sub



Public Sub fgText_Display(lAction As String, lDir As String, lFile As String)
Dim X As String

On Error GoTo Error_Handler

Me.Enabled = False: Me.MousePointer = vbHourglass
fgText.Visible = False
cmdExe.Enabled = True 'False
cmdExe.Caption = lAction
libExe_Action.Caption = lAction
libExe_Dir.Caption = lDir
libExe_File.Caption = lFile
Call lstErr_Clear(lstErr, cmdExe, lFile)
SSTab1.Tab = 2
fgText.Rows = 0

Open lDir & lFile For Input As #1
X = ""
Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
        fgText.Rows = fgText.Rows + 1
    fgText.Row = fgText.Rows - 1
    fgText.Col = 0: fgText.Text = Trim(xIn)

    
Loop
Close
cmdExe_Abort.Enabled = False
fgText.Visible = True
Select Case lAction
    Case constXCom: cmdExe.Enabled = Automate_Aut.Xspécial
    Case constCB_Aller, constBDP_Aller: cmdExe.Enabled = Automate_Aut.Saisir: cmdExe_Abort.Enabled = True
End Select

Me.Enabled = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
Shell_MsgBox "fgText_Display : " & Error, vbCritical, Me.Caption, False

Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdMonitor_Init()
Dim X As String, xHms As String
Dim I As Integer
If optMonitor_Aller Then
    xFlux = constAller
Else
    xFlux = constRetour
End If
xAppli = Trim(cboMonitor_Appli)

wAMJHMS = DSys & "_" & time_Hms & "_"
Call DTPicker_Control(txtMonitor_Amj, X)
I = 0: xHms = TimeHMS_Scan(txtMonitor_Hms & " ", I)
X = X & "_" & Mid$(xHms, 1, 6) & "_" & xAppli & "_" & xFlux
wFile_Monitor = X & ".bat"
wFile_Log = X & ".log"


xIBM_Library = Trim(txtMonitor_IBM_Library)

xIBM_File = Trim(txtMonitor_IBM_File)

xNT_File = Trim(txtMonitor_File)

xNT_Folder = paramPeliNT_DataF & xFlux & "\"
localNT_Folder = localPeliNT_DataF & xFlux & "\"
xFile_sab = xNT_Folder & constFTP & "\" & xIBM_File & ".sab"
xFile_TXT = xNT_Folder & constFTP & "\" & xNT_File
xFile_Archive = xNT_Folder & constArchive & "\" & wAMJHMS & xIBM_File & ".sab"
xFile_XCom = xNT_Folder & constXCom & "\" & xNT_File
localFile_XCom = localNT_Folder & constXCom & "\" & xNT_File
localFile_Log = localNT_Folder & constLog & "\" & wFile_Log

End Sub

Public Sub cmdMonitor_Aller_CB_Ok()
Dim xFile_sab As String, xFile_cb As String
wAMJHMS = DSys & "_" & time_Hms & "_"
xFile_sab = paramCB_Aller_Path & "situ.sab"
xFile_cb = paramCB_Aller_Path & "situ.cb"
If Trim(Dir(xFile_cb)) <> "" Then Kill xFile_cb
msFileSystem.CopyFile xFile_sab, paramBDF_Archive_Path & wAMJHMS & "situ.cb"
msFileSystem.MoveFile xFile_sab, xFile_cb
fgText.Clear
SSTab1.Tab = 1
If appliFonction = "*=>BDF_CB" Then Unload Me
End Sub

Public Sub cmdMonitor_Aller_BDP_Ok()
Dim xFile_sab As String, xFile_BDP As String
wAMJHMS = DSys & "_" & time_Hms & "_"
xFile_sab = paramBDP_Aller_Path & "12179.sab"
xFile_BDP = paramBDP_Aller_Path & "12179.BDP"
If Trim(Dir(xFile_BDP)) <> "" Then Kill xFile_BDP
msFileSystem.CopyFile xFile_sab, paramBDF_Archive_Path & wAMJHMS & "12179.BDP"
msFileSystem.MoveFile xFile_sab, xFile_BDP
fgText.Clear
SSTab1.Tab = 1
If appliFonction = "*=>BDF_BDP" Then Unload Me
End Sub

Public Sub cmdMonitor_Aller_BDP_Init()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call DTPicker_Set(txtMonitor_Amj, DSys)
txtMonitor_Hms = Time
cboMonitor_Appli.Text = "BDP"
optMonitor_Aller = True
txtMonitor_IBM_Library = paramIBM_Library_SAB
txtMonitor_IBM_File = "ZREPBPT0"
txtMonitor_File = "12179.bdp"
txtMonitor_Exe = ""
blnFile_FTP_Open = False
cmdMonitor_Aller_BDP

Me.Enabled = True: Me.MousePointer = 0

End Sub
