VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmDevise 
   Caption         =   "Devise.frm"
   ClientHeight    =   5265
   ClientLeft      =   3060
   ClientTop       =   1005
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5265
   ScaleWidth      =   5880
   Begin VB.Frame fraCours 
      Caption         =   "Impression EUR / ***"
      Height          =   2175
      Left            =   0
      TabIndex        =   15
      Top             =   3000
      Width           =   5850
      Begin VB.Frame fraCoursSort 
         Caption         =   "Tri"
         Height          =   975
         Left            =   3480
         TabIndex        =   17
         Top             =   960
         Width           =   1575
         Begin VB.OptionButton optCoursSortIso 
            Caption         =   "ISO"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   600
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optCoursSortN 
            Caption         =   "BdF"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame fraCoursOrigine 
         Caption         =   "Origine"
         Height          =   975
         Left            =   720
         TabIndex        =   16
         Top             =   960
         Width           =   1695
         Begin VB.OptionButton optCoursC 
            Caption         =   "Comptabilité"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton optCoursT 
            Caption         =   "Trésorerie"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin MSComCtl2.DTPicker txtCoursAmj 
         Height          =   375
         Left            =   1800
         TabIndex        =   23
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   12632256
         CustomFormat    =   "dddd dd  MMMM yyy"
         Format          =   28180483
         CurrentDate     =   36299
         MaxDate         =   401768
         MinDate         =   2
      End
      Begin VB.Label lblCoursAMJ 
         Caption         =   "date d'arrêté"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   5500
      Picture         =   "DEVISE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   400
   End
   Begin VB.Frame Frame2 
      Caption         =   "Devise"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   390
      Width           =   5850
      Begin VB.TextBox txtDevLib 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         TabIndex        =   13
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtDevCode 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtDevAmj 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         TabIndex        =   11
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtDevCours 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtDevX 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblDevAmj 
         Caption         =   "Date du Cours"
         Height          =   255
         Left            =   100
         TabIndex        =   8
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblCours 
         Caption         =   "Cours / FRF"
         Height          =   255
         Left            =   105
         TabIndex        =   7
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label lblDevLib 
         Caption         =   "Libellé"
         Height          =   255
         Left            =   100
         TabIndex        =   6
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblDevCode 
         Caption         =   "Code Devise"
         Height          =   255
         Left            =   100
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblDevX 
         Caption         =   "Devise"
         Height          =   255
         Left            =   100
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdPagePrior 
      Caption         =   "<<....."
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdPageNext 
      Caption         =   "... >>"
      Height          =   300
      Left            =   1000
      TabIndex        =   1
      Top             =   0
      Width           =   1000
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ErrMsg"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   2295
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3000
   End
End
Attribute VB_Name = "frmDevise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim valCoursAMJ As String * 8

'---------------------------------------------------------
Public Sub lblErrMsgStd()
'---------------------------------------------------------
lblErrMsg.Visible = True
lblErrMsg.FontBold = True
lblErrMsg.FontSize = 8
lblErrMsg.Top = 0
lblErrMsg.Left = 6000

End Sub


'---------------------------------------------------------
Private Sub cmdPageNext_Click()
'---------------------------------------------------------

If XListBox.ListIndex < XListBox.ListCount - 1 Then
    XListBox.ListIndex = XListBox.ListIndex + 1
    dspDevise
End If

End Sub

'---------------------------------------------------------
Private Sub cmdPagePrior_Click()
'---------------------------------------------------------

If XListBox.ListIndex > 0 Then
    XListBox.ListIndex = XListBox.ListIndex - 1
    dspDevise
End If

End Sub

'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
Dim Msg As String
Me.Hide
Msg = Space$(50)
Msg = Format$(1, "000000") & Format$(999, "000000") & valCoursAMJ & "TI"
If optCoursC Then Mid$(Msg, 21, 1) = "C"
If optCoursSortN Then Mid$(Msg, 22, 1) = "N"

prtDeviseX Msg

Me.Show
End Sub


'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me

End Sub


Private Sub Form_Activate()
Set XForm = Me

End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 27: cmdQuit_Click
    Case Is = 34: cmdPageNext_Click
    Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
    Case Is = 13: SendKeys "{TAB}"
End Select
lblErrMsg.Visible = False

End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------

Set XForm = Me
Call MeInit(arrTagNb, "Locked")
ReDim arrTag(arrTagNb + 1)
valCoursAMJ = DSys
txtCoursAmj = dateImp(valCoursAMJ)
txtCoursAmj.Enabled = True
dspDevise

End Sub






'---------------------------------------------------------
Public Sub dspDevise()
'---------------------------------------------------------
Dim X3 As String

If XListBox.ListIndex > 0 Then
    cmdPagePrior.Enabled = True
Else
    cmdPagePrior.Enabled = False
End If

If XListBox.ListIndex < XListBox.ListCount - 1 Then
    cmdPageNext.Enabled = True
Else
    cmdPageNext.Enabled = False
End If

X3 = mId$(XListBox.Text, 1, 3)
If IsNumeric(X3) Then
    Call DevX(X3)
Else
    Call DevCode(X3)
End If

txtDevX = XDevise.DevX
txtDevCode = Format$(XDevise.DevCode, "000")
txtDevLib = XDevise.DevLib
txtDevCours = XDevise.Cours
txtDevAmj = XDevise.DevAmj
txtDevAmj = dateImp(XDevise.DevAmj)
End Sub


Private Sub txtCoursAmj_Change()
txtCoursAMJ_Control
End Sub

Private Sub txtCoursAmj_GotFocus()
DTPicker_GotFocus txtCoursAmj
End Sub

Private Sub txtCoursAMJ_Control()
Dim X As String
X = Format$(txtCoursAmj.Year, "0000") & Format$(txtCoursAmj.Month, "00") & Format$(txtCoursAmj.Day, "00")
lblCoursAMJ.ForeColor = lblUsr.ForeColor
cmdPrint.Enabled = False
lblCoursAMJ = "date d'arrêté : "
If Not IsNumeric(X) Then
    lblCoursAMJ = lblCoursAMJ & X
    lblCoursAMJ.ForeColor = errUsr.ForeColor
    DTPicker_Now txtCoursAmj
Else
    If mId$(X, 1, 8) > DSys Then
        lblCoursAMJ = lblCoursAMJ & " > aujourd'hui"
        lblCoursAMJ.ForeColor = errUsr.ForeColor
        DTPicker_Now txtCoursAmj
    Else
        valCoursAMJ = mId$(X, 1, 8)
        cmdPrint.Enabled = True
    End If
End If

End Sub


Private Sub txtCoursAmj_LostFocus()
DTPicker_LostFocus txtCoursAmj
txtCoursAMJ_Control
End Sub


Private Sub txtCoursAmj_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DTPicker_GotFocus txtCoursAmj

End Sub

