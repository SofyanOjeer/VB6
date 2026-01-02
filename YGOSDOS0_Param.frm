VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmYGOSDOS0_Param 
   AutoRedraw      =   -1  'True
   Caption         =   "YGOSDOS0_Param"
   ClientHeight    =   11100
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   16035
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "YGOSDOS0_Param.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11100
   ScaleWidth      =   16035
   Begin VB.ComboBox cboSelect_SQL 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   60
      Width           =   3705
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   6120
      TabIndex        =   2
      Top             =   0
      Width           =   6900
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   13080
      Picture         =   "YGOSDOS0_Param.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin TabDlg.SSTab tabParam 
      Height          =   10515
      Left            =   -45
      TabIndex        =   4
      Top             =   540
      Width           =   16035
      _ExtentX        =   28284
      _ExtentY        =   18547
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Messages"
      TabPicture(0)   =   "YGOSDOS0_Param.frx":040C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "."
      TabPicture(1)   =   "YGOSDOS0_Param.frx":0428
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "."
      TabPicture(2)   =   "YGOSDOS0_Param.frx":0444
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame fraTab1 
         Height          =   10020
         Left            =   165
         TabIndex        =   5
         Top             =   375
         Width           =   15720
         Begin VB.Frame fraSelect_Cmd 
            BackColor       =   &H00E0FFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8820
            Left            =   10830
            TabIndex        =   7
            Top             =   180
            Width           =   4725
            Begin VB.CommandButton cmdSelect_Link 
               BackColor       =   &H0080C0FF&
               Caption         =   "Rapprochement partiel : Associer ce msg entrant au msg sortant  qui reste en attente de confirmation de la contrepartie"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1890
               Left            =   780
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   6500
               Width           =   3195
            End
            Begin VB.CommandButton cmdSelect_GOS_New 
               BackColor       =   &H0000FFFF&
               Caption         =   "Créer un dossier GOS"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   210
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   3015
               Width           =   1935
            End
            Begin VB.CommandButton cmdSelect_Ignore_2 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Ignorer"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1155
               Left            =   2580
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   1635
               Width           =   1935
            End
            Begin VB.CommandButton cmdSelect_Match 
               BackColor       =   &H0080FF80&
               Caption         =   "Rapprocher ces deux messages"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1305
               Left            =   780
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   4500
               Width           =   3195
            End
            Begin VB.CommandButton cmdSelect_Ignore_1 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Ignorer"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1170
               Left            =   210
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   1590
               Width           =   1935
            End
            Begin VB.CommandButton cmdSelect_Quit 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Quitter et réinitialiser"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   735
               Style           =   1  'Graphical
               TabIndex        =   8
               Top             =   600
               Width           =   3195
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   9705
            Left            =   30
            TabIndex        =   6
            Top             =   90
            Width           =   15600
            _ExtentX        =   27517
            _ExtentY        =   17119
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   16777215
            ForeColor       =   16384
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483633
            BackColorBkg    =   -2147483633
            AllowUserResizing=   3
            FormatString    =   $"YGOSDOS0_Param.frx":0460
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "mnuPrint"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmYGOSDOS0_Param"
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
Dim arrHab(19) As Boolean
Dim blnAuto As Boolean, blnError As Boolean
Dim cmdSelect_SQL_K As String



Dim HeightOfLine As Long, LinesOfText As Long

Dim txtRTF_prtForeColor_Header As Long

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim mXls2_Cols As Integer, mXls2_Row As Integer

Public Sub Form_Init(lfgSwift, blnVisible As Boolean)
Dim V, xSQL As String, X As String
Dim K As Long
'___________________________________________________________
On Error Resume Next
Me.Enabled = False

If Not frmYGOSDOS0_Param.Visible Then frmYGOSDOS0_Param.Left = 3000: frmYGOSDOS0_Param.Icon = frmElp_Icon

frmYGOSDOS0_Param.Visible = True
frmYGOSDOS0_Param.Show vbModeless
cmdSelect_Match.Visible = False
cmdSelect_Link.Visible = False
If mYSWIRAM0_Col = 0 Then
    fgSelect.Clear
    fraSelect_Cmd.Caption = ""
    cmdSelect_Ignore_1.Visible = False
    cmdSelect_Ignore_2.Visible = False
    
    cmdSelect_Ignore_1.Caption = "Ignorer " & oldYSWISAB0_1.SWISABWMTK & " " & oldYSWISAB0_1.SWISABWES _
        & " " & oldYSWISAB0_1.SWISABWL20 & " " & oldYSWISAB0_1.SWISABWN20
    cmdSelect_Ignore_1.Visible = blnVisible
    If oldYSWISAB0_1.SWISABXGOS = " " Then
        cmdSelect_GOS_New.Caption = "Créer un dossier GOS " & oldYSWISAB0_1.SWISABWMTK & " " & oldYSWISAB0_1.SWISABWES _
            & " " & oldYSWISAB0_1.SWISABWL20 & " " & oldYSWISAB0_1.SWISABWN20
        cmdSelect_GOS_New.Visible = blnVisible
    Else
        cmdSelect_GOS_New.Visible = False
        X = "select * from " & paramIBM_Library_SABSPE & ".YSWILNK0 " _
             & " where SWILNKSWID = " & oldYSWISAB0_1.SWISABSWID _
             & " and  SWILNKAPPC = 'GOS'"
        Set rsSab = cnsab.Execute(X)
        
        If Not rsSab.EOF Then
            fraSelect_Cmd.Caption = "Dossier GOS : " & rsSab("SWILNKAPPN")
            fraSelect_Cmd.ForeColor = vbMagenta
            
        End If
    End If
Else
    'cmdSelect_GOS_New.Visible = False
    cmdSelect_Ignore_2.Caption = "Ignorer " & oldYSWISAB0_2.SWISABWMTK & " " & oldYSWISAB0_2.SWISABWES _
        & " " & oldYSWISAB0_2.SWISABWL20 & " " & oldYSWISAB0_2.SWISABWN20
    cmdSelect_Ignore_2.Visible = blnVisible
    
    If oldYSWISAB0_1.SWISABWBIC = oldYSWISAB0_2.SWISABWBIC And oldYSWISAB0_1.SWISABWES <> oldYSWISAB0_2.SWISABWES Then
        If oldYSWISAB0_1.SWISABWMTK = oldYSWISAB0_2.SWISABWMTK And oldYSWISAB0_1.SWISABWMTK <> "399" Then
            cmdSelect_Match.Visible = blnVisible
            cmdSelect_Link.Visible = blnVisible
        Else
            If oldYSWISAB0_1.SWISABWMTK = "399" Or oldYSWISAB0_2.SWISABWMTK = "399" Then cmdSelect_Match.Visible = blnVisible: cmdSelect_Link.Visible = blnVisible
        End If
    End If
    For K = lfgSwift.Rows To fgSelect.Rows
        fgSelect.Row = K
        fgSelect.Col = 2: fgSelect.Text = ""
        fgSelect.Col = 3: fgSelect.Text = ""
    Next K

End If
'wRows = fgSelect.Rows
If lfgSwift.Rows > fgSelect.Rows Then fgSelect.Rows = lfgSwift.Rows

    fgSelect.Row = 0
    lfgSwift.Row = 0
    lfgSwift.Col = 0: fgSelect.Col = mYSWIRAM0_Col: fgSelect.Text = lfgSwift.Text
    fgSelect.CellForeColor = vbWhite: fgSelect.CellBackColor = mColor_GB
    lfgSwift.Col = 1: fgSelect.Col = 1 + mYSWIRAM0_Col: fgSelect.Text = lfgSwift.Text
    fgSelect.CellForeColor = vbWhite: fgSelect.CellBackColor = mColor_GB

For K = 1 To fgSelect.Rows - 1
    fgSelect.Row = K
    lfgSwift.Row = K
    lfgSwift.Col = 0: fgSelect.Col = mYSWIRAM0_Col: fgSelect.Text = lfgSwift.Text: fgSelect.CellForeColor = lfgSwift.CellForeColor
    fgSelect.CellBackColor = mColor_Y0
    lfgSwift.Col = 1: fgSelect.Col = 1 + mYSWIRAM0_Col: fgSelect.Text = lfgSwift.Text: fgSelect.CellForeColor = lfgSwift.CellForeColor
    fgSelect.CellBackColor = lfgSwift.CellBackColor
    
    If K < fgSelect.Rows - 4 Then
        If mYSWIRAM0_Col = 2 Then
            Select Case arrYSWIRAM0_Fields_X2(K)
                Case "=": fgSelect.CellBackColor = mColor_G1
                Case "#": fgSelect.CellBackColor = mColor_W0
                Case "*": fgSelect.CellBackColor = mColor_Y2
                Case Else: fgSelect.CellBackColor = vbWhite
            End Select
            If K > 1 Then
                fgSelect.Col = 1: fgSelect.Row = K - 1
                Select Case arrYSWIRAM0_Fields_X1(K)
                    Case "=": fgSelect.CellBackColor = mColor_G1
                    Case "#": fgSelect.CellBackColor = mColor_W0
                    Case "*": fgSelect.CellBackColor = mColor_Y2
                    Case Else: fgSelect.CellBackColor = vbWhite
                End Select
            End If
        End If
    End If
Next K

X = Trim(frmYGOSDOS0_Param.Caption)
AppActivate X

'____________________________________________________________
Me.Enabled = True

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


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim wFct As String

mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

wFct = UCase$(Trim(Mid$(Msg, 1, 12)))
Call BIA_VB_HAB(wFct, arrHab(), cboSelect_SQL)


Select Case wFct
    'Case "@?????":
    Case Else: 'blnAuto = False: Form_Init

End Select
End Sub



Private Sub cmdPrint_Click()

Select Case cmdSelect_SQL_K
    'Case "3": cmdSelect_SQL_ZBASTAB0_23_Exportation
End Select
End Sub

Private Sub cmdSelect_GOS_New_Click()
Call frmYGOSDOS0.YSWIRAM0_Match_6E_Fct("GOS_New")
Call cmdSelect_Quit_Click

End Sub

Private Sub cmdSelect_Ignore_1_Click()
Call frmYGOSDOS0.YSWIRAM0_Match_6E_Fct("I1")
Call cmdSelect_Quit_Click
End Sub

Private Sub cmdSelect_Ignore_2_Click()
Call frmYGOSDOS0.YSWIRAM0_Match_6E_Fct("I2")
Call cmdSelect_Quit_Click

End Sub


Private Sub cmdSelect_Link_Click()
Call frmYGOSDOS0.YSWIRAM0_Match_6E_Fct("L")
Call cmdSelect_Quit_Click

End Sub

Private Sub cmdSelect_Match_Click()
Call frmYGOSDOS0.YSWIRAM0_Match_6E_Fct("M")
Call cmdSelect_Quit_Click

End Sub


Private Sub cmdSelect_Quit_Click()
mYSWIRAM0_Col = 0
frmYGOSDOS0_Param.Hide
End Sub

Private Sub Form_Activate()
Set XForm = Me

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------

blnControl = False
blnError = False
usrColor_Set

cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
currentAction = ""
blnControl = True

End Sub

Public Sub cmdSelect_Clear()

lstErr.Clear

End Sub

Public Sub cmdSelect_Reset()
End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Return()
End Sub


Public Sub cmdContext_Quit()
lstErr.Clear: lstErr.Height = 200

cmdSelect_Quit_Click

Unload Me

End Sub

Private Sub Form_Load()


mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False

End Sub


Private Sub Form_Resize()

If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
mYSWIRAM0_Col = 0

End Sub

Private Sub lstErr_Click()
If lstErr.Height > 500 Then
    lstErr.Height = 480
Else
    lstErr.Height = lstErr.ListCount * 200 + 300
End If

End Sub



