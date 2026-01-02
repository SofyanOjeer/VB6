VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSAA 
   Caption         =   "SAA : interface"
   ClientHeight    =   9495
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13875
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
      Left            =   7440
      TabIndex        =   0
      Top             =   0
      Width           =   6345
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8940
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   13770
      _ExtentX        =   24289
      _ExtentY        =   15769
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Interface SwiftAlliance"
      TabPicture(0)   =   "SAA.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraFolder"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Impression Message File"
      TabPicture(1)   =   "SAA.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraImportMsgFile"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Import BIC, Histo"
      TabPicture(2)   =   "SAA.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraImportHisto"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraImportBIC"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraYSAAMSG"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.Frame fraYSAAMSG 
         Caption         =   "Import de messages => SAB073Y.YSAAMSG*"
         Height          =   1215
         Left            =   -74640
         TabIndex        =   31
         Top             =   6120
         Width           =   11655
         Begin VB.CommandButton cmdYSAAMSG_Ok 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Import YSAAMSG"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   9480
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   360
            Width           =   1440
         End
         Begin VB.TextBox txtYSSAMSG_File 
            Height          =   375
            Left            =   360
            TabIndex        =   32
            Text            =   "C:\Temp\SAA_KHALDZAL\o_200012.txt"
            Top             =   600
            Width           =   7335
         End
      End
      Begin VB.Frame fraImportBIC 
         Caption         =   "Import fichier BIC"
         Height          =   2775
         Left            =   -74760
         TabIndex        =   25
         Top             =   3120
         Width           =   11775
         Begin VB.CommandButton cmdImportBIC 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Import BIC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   4560
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1080
            Width           =   2040
         End
         Begin VB.TextBox txtImportBIC 
            Height          =   285
            Left            =   120
            TabIndex        =   26
            Text            =   "D:\Temp\SAA\FI.dat"
            Top             =   480
            Width           =   11055
         End
      End
      Begin VB.Frame fraImportHisto 
         Caption         =   "Import Histo"
         Height          =   2295
         Left            =   -74880
         TabIndex        =   19
         Top             =   600
         Width           =   8415
         Begin VB.TextBox txtImportHistoPath 
            Height          =   285
            Left            =   2160
            TabIndex        =   23
            Text            =   "C:\Temp\Swift_Histo\Emission\"
            Top             =   600
            Width           =   3015
         End
         Begin VB.TextBox txtImportHistoFile 
            Height          =   285
            Left            =   2160
            TabIndex        =   22
            Text            =   "00140041.prt"
            Top             =   1080
            Width           =   3015
         End
         Begin VB.CommandButton cmdImportHisto 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Import Histo"
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
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   600
            Width           =   2160
         End
         Begin VB.CommandButton cmdPrintHisto 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Print Test"
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
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1320
            Width           =   2160
         End
         Begin VB.Label lblImportHistoInput 
            Caption         =   "Input File"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   1575
         End
      End
      Begin VB.Frame fraImportMsgFile 
         Caption         =   "Impression 'Message File'"
         Height          =   7815
         Left            =   -74760
         TabIndex        =   15
         Top             =   480
         Width           =   8775
         Begin VB.ListBox lstImportMsgFileSelect 
            Height          =   5325
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   28
            Top             =   1080
            Width           =   3735
         End
         Begin VB.CommandButton cmdImportMsgFile 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Impression"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   2880
            Width           =   2280
         End
         Begin VB.TextBox txtImportMsgFile 
            Height          =   285
            Left            =   240
            TabIndex        =   16
            Text            =   "C:\Temp\Swift\"
            Top             =   240
            Width           =   6135
         End
         Begin VB.Label lblImportMsgFileSelect 
            Caption         =   "Printer à sélectionner"
            Height          =   255
            Left            =   840
            TabIndex        =   18
            Top             =   720
            Width           =   1815
         End
      End
      Begin VB.Frame fraFolder 
         Height          =   8535
         Left            =   120
         TabIndex        =   2
         Top             =   330
         Width           =   13575
         Begin VB.Frame fraMT950 
            Caption         =   "MT950 "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3405
            Left            =   120
            TabIndex        =   10
            Top             =   5040
            Width           =   13335
            Begin VB.CommandButton cmdOk_SAA_to_Corona 
               BackColor       =   &H00C0FFC0&
               Caption         =   "SAA => CORONA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   8280
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   1200
               Width           =   2295
            End
            Begin VB.CommandButton cmdOk_Loro 
               BackColor       =   &H00808080&
               Caption         =   "Loro => SAA"
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
               Left            =   11520
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   2280
               Width           =   1860
            End
            Begin VB.CommandButton cmdOK_Nostro 
               BackColor       =   &H00808080&
               Caption         =   "Nostro => Corona"
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
               Left            =   11400
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   360
               Width           =   1740
            End
            Begin MSFlexGridLib.MSFlexGrid fgSAA_to_Corona 
               Height          =   3045
               Left            =   180
               TabIndex        =   13
               Top             =   270
               Width           =   7710
               _ExtentX        =   13600
               _ExtentY        =   5371
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
               FormatString    =   $"SAA.frx":0054
            End
         End
         Begin VB.Frame fraSAB 
            Caption         =   "SAB"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4740
            Left            =   135
            TabIndex        =   4
            Top             =   165
            Width           =   13320
            Begin VB.CheckBox chkPCC_Saa_from_SAB 
               Caption         =   "Format : *.sab => *.pcc "
               Height          =   270
               Left            =   10920
               TabIndex        =   30
               Top             =   1080
               Value           =   1  'Checked
               Width           =   1980
            End
            Begin VB.CheckBox chkFTP_Saa_from_SAB 
               Caption         =   "FTP : yswiall0=> *.sab"
               Height          =   270
               Left            =   10920
               TabIndex        =   29
               Top             =   480
               Value           =   1  'Checked
               Width           =   1980
            End
            Begin VB.CommandButton cmdOK_SAA_to_SAB 
               BackColor       =   &H00C0FFC0&
               Caption         =   "SAA   =>   SAB"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   8040
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   2640
               Width           =   2300
            End
            Begin VB.FileListBox filDoc 
               Height          =   285
               Left            =   360
               TabIndex        =   7
               Top             =   720
               Visible         =   0   'False
               Width           =   2535
            End
            Begin VB.CommandButton cmdOK_SAA_from_SAB 
               BackColor       =   &H00C0FFC0&
               Caption         =   "SAB  =>  SAA "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   8000
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   360
               Width           =   2300
            End
            Begin MSFlexGridLib.MSFlexGrid fgSAA_from_SAB 
               Height          =   1530
               Left            =   135
               TabIndex        =   6
               Top             =   240
               Width           =   7680
               _ExtentX        =   13547
               _ExtentY        =   2699
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
               FormatString    =   $"SAA.frx":00DB
            End
            Begin MSFlexGridLib.MSFlexGrid fgSAA_to_SAb 
               Height          =   2565
               Left            =   120
               TabIndex        =   8
               Top             =   1920
               Width           =   7785
               _ExtentX        =   13732
               _ExtentY        =   4524
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
               FormatString    =   $"SAA.frx":0169
            End
         End
      End
   End
End
Attribute VB_Name = "frmSAA"
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
Dim SwiftAut As typeAuthorization

Dim IdShell

Dim blncmdOk_Run As Boolean, blnAuto_Swift As Boolean

Dim blnImportMsgFile As Boolean
Dim blnTransaction As Boolean

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim X As String
Call BiaPgmAut_Init(Mid$(Msg, 1, 12), SwiftAut)    '

lstImportMsgFileSelect.Clear
cmdImportMsgFile.Caption = "Import"
fraFolder.Enabled = False
fraImportMsgFile.Enabled = False
fraImportHisto.Enabled = True 'False
fraImportBIC.Enabled = False

blnImportMsgFile = False

chkFTP_Saa_from_SAB = "1": chkFTP_Saa_from_SAB.Enabled = SwiftAut.Xspécial
chkPCC_Saa_from_SAB = "1": chkFTP_Saa_from_SAB.Enabled = SwiftAut.Xspécial

If SwiftAut.Consulter = True Then
    fraImportMsgFile.Enabled = True
End If
''SwiftAut.Swift= False
If SwiftAut.Swift = True Then
    fraImportBIC.Enabled = True
    fraFolder.Enabled = True
    cmdOK_SAA_to_SAB.Enabled = False
    cmdOK_SAA_from_SAB.Enabled = False
    cmdOK_Nostro.Enabled = False
    cmdOk_Loro.Enabled = False
    cmdOk_SAA_to_Corona.Enabled = False
    
    paramSAA_Init
    
            
   ''' If blnJPL Then paramSAA_Init_Test
    
    txtImportBIC = paramSwift_BIC_Input
    Call lstErr_Clear(lstErr, cmdContext, "Paramètres chargés")
    
    fgSAA_to_SAB_Load
    fgSAA_from_SAB_Load
    fgSAA_to_Corona_Load
    
    ''''X = Dir(paramSwiftNostro_MT950_File): If X <> "" Then cmdOK_Nostro.Enabled = True
    '''''''''X = Dir(paramSwiftLoro_MT950_File): If X <> "" Then cmdOk_Loro.Enabled = True
    
    Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
'        Case "$AUTO_SWIFT":     blnAuto_Swift = True:    Auto_Swift
        Case "@SAA_ENTRANT":     blnAuto_Swift = True: Auto_SAA_ENTRANT
        Case Else: blnAuto_Swift = False
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

Private Sub cmdContext_Click()
cmdContext_Quit

End Sub

Private Sub cmdImportBIC_Click()
Dim X As String, xFile As String
xFile = Trim(txtImportBIC)
X = Dir(xFile)
If X = "" Then
    Call lstErr_Clear(lstErr, cmdContext, "? fichier import BIC non trouvé")
Else
    srvSAA.ImportBIC_Load xFile
End If
End Sub

Private Sub cmdImportHisto_Click()
paramSwiftHisto_Input = Trim(txtImportHistoPath) & Trim(txtImportHistoFile)
MsgBox " à faire srvSAA.ImportHisto_Load, voir srvSAA_20050531", vbCritical
'ImportMsgFile_Load paramSwiftHisto_Input
End Sub

Private Sub cmdImportMsgFile_Click()
Dim X As String, xFile As String
xFile = Trim(txtImportMsgFile)
X = Dir(xFile)
If X = "" Then
    Call lstErr_Clear(lstErr, cmdContext, "? fichier import MsgFile non trouvé")
Else
    If blnImportMsgFile Then
            srvSAA.ImportMsgFile_Print xFile, Mid$(lstImportMsgFileSelect, 1, 4)
    Else
            srvSAA.ImportMsgFile_Load xFile
            blnImportMsgFile = True
            cmdImportMsgFile.Caption = "Impression"
            For I = 1 To arrMsgFile_Printer_Nb
                lstImportMsgFileSelect.AddItem arrMsgFile_Printer(I) & Chr$(9) & arrMsgFile_Seq(I, 0)
            Next I
        End If
End If

End Sub

Private Sub cmdOk_Loro_Click()
''cmdOK_Run cmdOk_Loro

End Sub

Private Sub cmdOK_Nostro_Click()
''cmdOK_Run cmdOK_Nostro

End Sub

Private Sub cmdOk_SAA_to_Corona_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdOK_Run cmdOk_SAA_to_Corona
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdOK_SAA_to_SAB_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdOK_Run cmdOK_SAA_to_SAB
'blncmdOk_Run = True
'cmdOK_SAA_to_SAB.Enabled = False
'Me.Enabled = False
'SAA_to_SAB
'Me.Enabled = True
'AppActivate Me.Caption
'blncmdOk_Run = False
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdOK_SAA_from_SAB_Click()
Call MsgBox("Cette fonction est remplacée par une autre procèdure ? ", vbCritical, "transfert SAB vers SAA")
'X = MsgBox("Cette fonction est remplacée par une autre procèdure, voulez-vous continuer malgré cet avis ? ", vbCritical + vbYesNo + vbDefaultButton2, "transfert SAB vers SAA")
'If X <> vbYes Then Exit Sub

'Me.Enabled = False: Me.MousePointer = vbHourglass

'Call lstErr_Clear(lstErr, cmdContext, "SAA_from_SAB : début")
'If chkFTP_Saa_from_SAB = "1" Then srvSAA.SAA_from_SAB_FTP

'fgSAA_from_SAB_Load

'If chkPCC_Saa_from_SAB = "1" Then
'    If filDoc.ListCount > 0 Then
'        cmdOK_Run cmdOK_SAA_from_SAB
'    End If
'End If
'Call lstErr_AddItem(lstErr, cmdContext, "SAA_from_SAB : " & filDoc.ListCount & " fichier")
'Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdPrintHisto_Click()
Dim Msg As String

Msg = Space$(100)
MsgBox "à faire :prtSwiftHisto_Monitor Msg", vbCritical
End Sub

Private Sub cmdYSAAMSG_Ok_Click()
Dim X As String, xFile As String
Me.Enabled = False: Me.MousePointer = vbHourglass
xFile = Trim(txtYSSAMSG_File)
X = Dir(xFile)
If X = "" Then
    Call lstErr_Clear(lstErr, cmdContext, "? fichier import  non trouvé")
Else
    YSAAMSG_Import xFile
End If
Me.Enabled = True: Me.MousePointer = 0

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
cmdOK_SAA_from_SAB.Enabled = False
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



Public Sub fgSAA_from_SAB_Load()
Dim I As Integer, K As Integer, X As String, L As Integer, iSession As Integer


filDoc.path = paramSAA_Data_from_SAB
filDoc.Pattern = "x.xxx"
filDoc.Pattern = "*" & paramSAA_Data_from_SAB_ExtensionP_sab
fgSAA_from_SAB.Redraw = False
fgSAA_from_SAB.Rows = 1
fgSAA_from_SAB.Enabled = True
For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.path & "\" & filDoc.FileName)
    fgSAA_from_SAB.Rows = fgSAA_from_SAB.Rows + 1
    fgSAA_from_SAB.Row = fgSAA_from_SAB.Rows - 1
    fgSAA_from_SAB.Col = 0: fgSAA_from_SAB.Text = Trim(filDoc.FileName)
    fgSAA_from_SAB.Col = 1: fgSAA_from_SAB.Text = msFile.DateLastModified
   ' cmdOK_SAA_from_SAB.Enabled = True
Next I
fgSAA_from_SAB.Redraw = True
'20050602 jpl cmdOK_SAA_from_SAB.Enabled = True


End Sub
Public Sub fgSAA_to_SAB_Load()
Dim I As Integer, K As Integer, X As String, L As Integer, iSession As Integer


filDoc.path = paramSAA_Data_to_SAB
filDoc.Pattern = "*" & paramSAA_Data_to_SAB_ExtensionP_out

fgSAA_to_SAb.Redraw = False
fgSAA_to_SAb.Rows = 1
fgSAA_to_SAb.Enabled = True
For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.path & "\" & filDoc.FileName)
    fgSAA_to_SAb.Rows = fgSAA_to_SAb.Rows + 1
    fgSAA_to_SAb.Row = fgSAA_to_SAb.Rows - 1
    fgSAA_to_SAb.Col = 0: fgSAA_to_SAb.Text = Trim(filDoc.FileName)
    fgSAA_to_SAb.Col = 1: fgSAA_to_SAb.Text = msFile.DateLastModified
    cmdOK_SAA_to_SAB.Enabled = True
Next I
fgSAA_to_SAb.Redraw = True

End Sub

Public Sub fgSAA_to_Corona_Load()
Dim I As Integer, K As Integer, X As String, L As Integer, iSession As Integer


filDoc.path = paramSAA_Data_to_Corona
filDoc.Pattern = "*.*"

fgSAA_to_Corona.Redraw = False
fgSAA_to_Corona.Rows = 1
fgSAA_to_Corona.Enabled = True
For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.path & "\" & filDoc.FileName)
    fgSAA_to_Corona.Rows = fgSAA_to_Corona.Rows + 1
    fgSAA_to_Corona.Row = fgSAA_to_Corona.Rows - 1
    fgSAA_to_Corona.Col = 0: fgSAA_to_Corona.Text = Trim(filDoc.FileName)
    fgSAA_to_Corona.Col = 1: fgSAA_to_Corona.Text = msFile.DateLastModified
    cmdOk_SAA_to_Corona.Enabled = True
Next I
fgSAA_to_Corona.Redraw = True

End Sub



Public Sub Auto_SAA_ENTRANT()
Dim blnOk As Boolean

blncmdOk_Run = False
Do
    If blncmdOk_Run = False Then
        blnOk = True
    
        If cmdOk_SAA_to_Corona.Enabled Then blnOk = False: cmdOk_SAA_to_Corona_Click
        If cmdOK_SAA_to_SAB.Enabled Then blnOk = False: cmdOK_SAA_to_SAB_Click
''''        If filDoc.ListCount > 0 Then blnOk = False: cmdOK_SAA_from_SAB_Click
        DoEvents
    End If
Loop Until blnOk = True
Unload Me

End Sub

Public Sub cmdOK_Run(C As CommandButton)

blncmdOk_Run = True

Select Case Trim(C.Name)
    Case "cmdOK_SAA_to_SAB":       cmdOK_SAA_to_SAB.Enabled = False
                                SAA_to_SAB blnAuto_Swift
    Case "cmdOk_SAA_to_Corona":    cmdOk_SAA_to_Corona.Enabled = False
                                SAA_to_Corona_Put
    ''Case "cmdOk_Loro":          cmdOk_Loro.Enabled = False
     ''                           Loro_Put
    ''Case "cmdOK_Nostro":        cmdOK_Nostro.Enabled = False
     '''                           Nostro_Put
    Case "cmdOK_SAA_from_SAB":       ''cmdOK_SAA_from_SAB.Enabled = False
                                    SAA_from_SAB
End Select

Me.Show
blncmdOk_Run = False


End Sub

