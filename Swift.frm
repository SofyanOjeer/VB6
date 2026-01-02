VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSwift 
   Caption         =   "Swift : interface"
   ClientHeight    =   6345
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9330
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
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Interface SwiftAlliance"
      TabPicture(0)   =   "Swift.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraFolder"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Impression Message File"
      TabPicture(1)   =   "Swift.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraImportMsgFile"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Import BIC, Histo"
      TabPicture(2)   =   "Swift.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraImportHisto"
      Tab(2).Control(1)=   "fraImportBIC"
      Tab(2).ControlCount=   2
      Begin VB.Frame fraImportBIC 
         Caption         =   "Import fichier BIC"
         Height          =   1935
         Left            =   -74760
         TabIndex        =   25
         Top             =   3120
         Width           =   3735
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
            Height          =   615
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   960
            Width           =   2040
         End
         Begin VB.TextBox txtImportBIC 
            Height          =   285
            Left            =   120
            TabIndex        =   26
            Text            =   "C:\Temp\FI.dat"
            Top             =   480
            Width           =   3495
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
            Text            =   "D:\Temp\Swift_Histo\Emission\"
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
         Height          =   4935
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   5295
         Begin VB.ListBox lstImportMsgFileSelect 
            Height          =   3375
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   28
            Top             =   960
            Width           =   2295
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
            Height          =   855
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   960
            Width           =   1680
         End
         Begin VB.TextBox txtImportMsgFile 
            Height          =   285
            Left            =   240
            TabIndex        =   16
            Text            =   "D:\Temp\Swift_2001"
            Top             =   240
            Width           =   4695
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
         Height          =   5175
         Left            =   -74880
         TabIndex        =   2
         Top             =   330
         Width           =   8895
         Begin VB.Frame fraMT950 
            Caption         =   "MT950"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   120
            TabIndex        =   10
            Top             =   2640
            Width           =   8655
            Begin VB.CommandButton cmdOk_SAA_Corona 
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
               Height          =   495
               Left            =   720
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   1800
               Width           =   3120
            End
            Begin VB.CommandButton cmdOk_Loro 
               BackColor       =   &H00C0FFC0&
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
               Height          =   495
               Left            =   4800
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   1080
               Width           =   3120
            End
            Begin VB.CommandButton cmdOK_Nostro 
               BackColor       =   &H00C0FFC0&
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
               Height          =   495
               Left            =   4800
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   360
               Width           =   3120
            End
            Begin MSFlexGridLib.MSFlexGrid fgSAA_Corona 
               Height          =   1410
               Left            =   240
               TabIndex        =   13
               Top             =   240
               Width           =   4185
               _ExtentX        =   7382
               _ExtentY        =   2487
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
               FormatString    =   "<fichiers     SAA  => CORONA        <|<Date dernière modif"
            End
         End
         Begin VB.Frame fraTI 
            Caption         =   "Trade Innovation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   8655
            Begin VB.CommandButton cmdOK_SAA_TI 
               BackColor       =   &H00C0FFC0&
               Caption         =   "SAA => TI"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   4800
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   1800
               Width           =   3120
            End
            Begin VB.FileListBox filDoc 
               Height          =   480
               Left            =   360
               TabIndex        =   7
               Top             =   720
               Visible         =   0   'False
               Width           =   2535
            End
            Begin VB.CommandButton cmdOK_TI_SAA 
               BackColor       =   &H00C0FFC0&
               Caption         =   "TI => SAA (uniquement MT730)"
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
               Left            =   600
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   1800
               Width           =   3120
            End
            Begin MSFlexGridLib.MSFlexGrid fgTI_SAA 
               Height          =   1410
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Width           =   4065
               _ExtentX        =   7170
               _ExtentY        =   2487
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
               FormatString    =   "<fichiers       TI => SAA             <|<Date dernière modif    "
            End
            Begin MSFlexGridLib.MSFlexGrid fgSAA_TI 
               Height          =   1410
               Left            =   4320
               TabIndex        =   8
               Top             =   240
               Width           =   4185
               _ExtentX        =   7382
               _ExtentY        =   2487
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
               FormatString    =   "<fichiers          SAA  => TI            <|<Date dernière modif   "
            End
         End
      End
   End
End
Attribute VB_Name = "frmSwift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim x As String, X1 As String, I As Integer
Dim Msg As String, valX As String
Dim currentMethod As String, lastMethod As String
Dim SwiftAut As typeAuthorization

Dim IdShell

Dim blncmdOk_Run As Boolean, blnAuto_Swift As Boolean

Dim blnImportMsgFile As Boolean
Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim x As String
Call BiaPgmAut_Init(mId$(Msg, 1, 12), SwiftAut)    '

lstImportMsgFileSelect.Clear
cmdImportMsgFile.Caption = "Import"
fraFolder.Enabled = False
fraImportMsgFile.Enabled = False
fraImportHisto.Enabled = False
fraImportBIC.Enabled = False

blnImportMsgFile = False

If SwiftAut.Consulter = True Then
    fraImportMsgFile.Enabled = True
End If
''SwiftAut.Swift= False
If SwiftAut.Swift = True Then
    fraFolder.Enabled = True
    cmdOK_SAA_TI.Enabled = False
    cmdOK_TI_SAA.Enabled = False
    cmdOK_Nostro.Enabled = False
    cmdOk_Loro.Enabled = False
    cmdOk_SAA_Corona.Enabled = False
    
    srvSwift.param_Init
    'MsgBox "param_init_TEST", vbInformation, "frmSwift.Msg_Rcv"
    'srvSwift.Param_Init_Test
    
    fgTI_SAA_Load
    fgSAA_TI_Load
    fgSAA_Corona_Load
    
    x = Dir(paramSwiftNostro_MT950_File): If x <> "" Then cmdOK_Nostro.Enabled = True
    x = Dir(paramSwiftLoro_MT950_File): If x <> "" Then cmdOk_Loro.Enabled = True
    
    Select Case UCase$(Trim(mId$(Msg, 1, 12)))
        Case "$AUTO_SWIFT":     blnAuto_Swift = True:    Auto_Swift
        Case "@AUTO_SWIFT":     blnAuto_Swift = True: Auto_Swift
        Case Else: blnAuto_Swift = False
    End Select
End If

End Sub


Public Sub cmdContext_Quit()
    If blnMsgBox_Quit Then
       x = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
    Else
       x = vbYes
    End If
    If x = vbYes Then Unload Me

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
Dim x As String, xFile As String
xFile = Trim(txtImportBIC)
x = Dir(xFile)
If x = "" Then
    Call lstErr_Clear(lstErr, cmdContext, "? fichier import BIC non trouvé")
Else
    srvSwift.ImportBIC_Load xFile
End If
End Sub

Private Sub cmdImportHisto_Click()
paramSwiftHisto_Input = Trim(txtImportHistoPath) & Trim(txtImportHistoFile)
srvSwift.ImportHisto_Load
End Sub

Private Sub cmdImportMsgFile_Click()
Dim x As String, xFile As String
xFile = Trim(txtImportMsgFile)
x = Dir(xFile)
If x = "" Then
    Call lstErr_Clear(lstErr, cmdContext, "? fichier import MsgFile non trouvé")
Else
    If blnImportMsgFile Then
            srvSwift.ImportMsgFile_Print xFile, mId$(lstImportMsgFileSelect, 1, 4)
    Else
            srvSwift.ImportMsgFile_Load xFile
            blnImportMsgFile = True
            cmdImportMsgFile.Caption = "Impression"
            For I = 1 To arrMsgFile_Printer_Nb
                lstImportMsgFileSelect.AddItem arrMsgFile_Printer(I) & Chr$(9) & arrMsgFile_Seq(I, 0)
            Next I
        End If
End If

End Sub

Private Sub cmdOk_Loro_Click()
cmdOK_Run cmdOk_Loro

End Sub

Private Sub cmdOK_Nostro_Click()
cmdOK_Run cmdOK_Nostro

End Sub

Private Sub cmdOk_SAA_Corona_Click()
cmdOK_Run cmdOk_SAA_Corona

End Sub

Private Sub cmdOK_SAA_TI_Click()
cmdOK_Run cmdOK_SAA_TI
'blncmdOk_Run = True
'cmdOK_SAA_TI.Enabled = False
'Me.Enabled = False
'SAA_TI_Put
'Me.Enabled = True
'AppActivate Me.Caption
'blncmdOk_Run = False
End Sub

Private Sub cmdOK_TI_SAA_Click()
cmdOK_Run cmdOK_TI_SAA
End Sub

Private Sub cmdPrintHisto_Click()
Dim Msg As String

Msg = Space$(100)
prtSwiftHisto_Monitor Msg
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

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub fraFolder_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set sstab1
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



Public Sub fgTI_SAA_Load()
Dim I As Integer, K As Integer, x As String, L As Integer, iSession As Integer


filDoc.Path = paramSwiftTiSaa_TI_Out
filDoc.Pattern = paramSwiftTiSaa_TI_Pattern

fgTI_SAA.Redraw = False
fgTI_SAA.Rows = 1
fgTI_SAA.Enabled = True
For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.Path & "\" & filDoc.Filename)
    fgTI_SAA.Rows = fgTI_SAA.Rows + 1
    fgTI_SAA.Row = fgTI_SAA.Rows - 1
    fgTI_SAA.Col = 0: fgTI_SAA.Text = Trim(filDoc.Filename)
    fgTI_SAA.Col = 1: fgTI_SAA.Text = msFile.DateLastModified
    cmdOK_TI_SAA.Enabled = True
Next I
fgTI_SAA.Redraw = True



End Sub
Public Sub fgSAA_TI_Load()
Dim I As Integer, K As Integer, x As String, L As Integer, iSession As Integer


filDoc.Path = paramSwiftSaaTi_SAA_Out
filDoc.Pattern = "*.*"

fgSAA_TI.Redraw = False
fgSAA_TI.Rows = 1
fgSAA_TI.Enabled = True
For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.Path & "\" & filDoc.Filename)
    fgSAA_TI.Rows = fgSAA_TI.Rows + 1
    fgSAA_TI.Row = fgSAA_TI.Rows - 1
    fgSAA_TI.Col = 0: fgSAA_TI.Text = Trim(filDoc.Filename)
    fgSAA_TI.Col = 1: fgSAA_TI.Text = msFile.DateLastModified
    cmdOK_SAA_TI.Enabled = True
Next I
fgSAA_TI.Redraw = True

End Sub

Public Sub fgSAA_Corona_Load()
Dim I As Integer, K As Integer, x As String, L As Integer, iSession As Integer


filDoc.Path = paramSwiftSaaCorona_SAA_Out
filDoc.Pattern = "*.*"

fgSAA_Corona.Redraw = False
fgSAA_Corona.Rows = 1
fgSAA_Corona.Enabled = True
For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.Path & "\" & filDoc.Filename)
    fgSAA_Corona.Rows = fgSAA_Corona.Rows + 1
    fgSAA_Corona.Row = fgSAA_Corona.Rows - 1
    fgSAA_Corona.Col = 0: fgSAA_Corona.Text = Trim(filDoc.Filename)
    fgSAA_Corona.Col = 1: fgSAA_Corona.Text = msFile.DateLastModified
    cmdOk_SAA_Corona.Enabled = True
Next I
fgSAA_Corona.Redraw = True

End Sub



Public Sub Auto_Swift()
Dim blnOk As Boolean

blncmdOk_Run = False
Do
    If blncmdOk_Run = False Then
        blnOk = True
    
        If cmdOk_Loro.Enabled Then blnOk = False: cmdOk_Loro_Click
        If cmdOK_Nostro.Enabled Then blnOk = False: cmdOK_Nostro_Click
        If cmdOk_SAA_Corona.Enabled Then blnOk = False: cmdOk_SAA_Corona_Click
        If cmdOK_SAA_TI.Enabled Then blnOk = False: cmdOK_SAA_TI_Click
        If cmdOK_TI_SAA.Enabled Then blnOk = False: cmdOK_TI_SAA_Click
        DoEvents
    End If
Loop Until blnOk = True
Unload Me

End Sub

Public Sub cmdOK_Run(C As CommandButton)
blncmdOk_Run = True
Me.Enabled = False

Select Case Trim(C.Name)
    Case "cmdOK_SAA_TI":        cmdOK_SAA_TI.Enabled = False
                                SAA_TI_Put
    Case "cmdOk_SAA_Corona":    cmdOk_SAA_Corona.Enabled = False
                                SAA_Corona_Put
    Case "cmdOk_Loro":          cmdOk_Loro.Enabled = False
                                Loro_Put
    Case "cmdOK_Nostro":        cmdOK_Nostro.Enabled = False
                                Nostro_Put
    Case "cmdOK_TI_SAA":        cmdOK_TI_SAA.Enabled = False
                                TI_SAA_Put
End Select

Me.Enabled = True
'AppActivate Me.Caption
blncmdOk_Run = False


End Sub

