VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmElp 
   AutoRedraw      =   -1  'True
   Caption         =   "Bonjour"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6345
   ScaleWidth      =   9330
   Begin VB.CheckBox chkDtaq 
      Caption         =   "Affichage Dtaq"
      Height          =   375
      Left            =   525
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ListBox lstMain 
      BackColor       =   &H00C0C0C0&
      Height          =   3375
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   6495
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      DTREnable       =   -1  'True
      InBufferSize    =   8192
      OutBufferSize   =   8192
   End
   Begin VB.Image imgBiaMouse 
      Height          =   480
      Left            =   480
      Picture         =   "biatst.frx":0000
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblOut 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
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
      Height          =   270
      Left            =   600
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.Label lblIN 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
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
      Height          =   270
      Left            =   600
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Caption         =   "Initialisation  liaison AS400 ........."
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
      Height          =   275
      Left            =   1320
      TabIndex        =   0
      Top             =   650
      Width           =   6495
   End
   Begin VB.Image imgSocSignon 
      Height          =   6750
      Left            =   0
      Top             =   -240
      Width           =   9450
   End
End
Attribute VB_Name = "frmElp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean

Dim frmCaption As String



'---------------------------------------------------------
Private Sub cmdPrintForm_Click()
'---------------------------------------------------------
'Me.PrintForm
Dim Msg As String
Msg = Space(20)
'prtVirx Msg

End Sub

'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
'Unload Me
mainEnd
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
    Case Is = 27: cmdQuit_Click
'    Case Is = 34: cmdPageNext_Click
'    Case Is = 33: cmdPagePrior_Click
    Case Is = 44: frmElpPrt.prtScreen
    Case Is = 13: SendKeys "{TAB}"
End Select


End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------

frmCaption = Me.Caption
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)

End Sub





'---------------------------------------------------------
Private Sub lstMain_Click()
'---------------------------------------------------------
Dim X As String

X = Space$(100)
X = Trim(lstMain.Text)
If lblMain.Caption = "Menu" Then
    Msg_Rcv X
Else
    Elp.usrId = Mid$(X, 1, 10)
    usrService = Mid$(X, 11, 3)
    usrGestionnaire = Mid$(X, 14, 2)
    usrName = Mid$(X, 17, 34)
    Me.Caption = frmCaption

    Set XForm = Me
    Call MeInit(arrTagNb)
    ReDim arrTag(arrTagNb + 1)

    srvUsrApp
End If

End Sub

'---------------------------------------------------------
Public Sub frmCompte_Show()
'---------------------------------------------------------
Dim X As String

frmCompte.Show vbModeless
frmCompte.Visible = True
X = frmCompte.Caption
AppActivate X

End Sub
'---------------------------------------------------------
Public Sub frmOpTrf_Show()
'---------------------------------------------------------
Dim X As String

frmOpTrf.Show vbModeless
frmOpTrf.Visible = True
X = frmOpTrf.Caption
AppActivate X

End Sub
'---------------------------------------------------------
Public Sub frmDictio_Show()
'---------------------------------------------------------
Dim X As String

frmDictio.Show vbModeless
frmDictio.Visible = True
X = frmDictio.Caption
AppActivate X

End Sub


'---------------------------------------------------------
Public Sub frmBic_Show()
'---------------------------------------------------------
Dim X As String

frmBic.Show vbModeless
frmBic.Visible = True
X = frmBic.Caption
AppActivate X

End Sub
'---------------------------------------------------------
Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------

Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
   Case Is = "FRMCOMPTE", "COMPTE": frmCompte_Show: frmCompte.Msg_Rcv Msg
   Case Is = "FRMOPTRF", "OPTRF": frmOpTrf_Show: frmOpTrf.Msg_Rcv Msg
   Case Is = "FRMOPTRFD", "OPTRFD": frmOptrfD_show: frmOpTrfD.Msg_Rcv Msg
'   Case Is = "FRMBDF", "BDF": frmBdf_Show: frmBdf.Msg_Rcv Msg
   Case Is = "FRMBIC", "BIC": frmBic_Show: frmBic.Msg_Rcv Msg
   Case Is = "FRMBICLIST": frmBicList_Show: frmBicList.Msg_Rcv Msg
   Case Is = "FRMDICTIO", "DICTIO": frmDictio_Show: frmDictio.Msg_Rcv Msg
 '  Case Is = "FRMDEVCOUP", "TEST": frmDeviseCoupures_Show: frmDeviseCoupures.Msg_Rcv Msg
End Select

End Sub

Public Sub frmBicList_Show()
'---------------------------------------------------------
Dim X As String

frmBicList.Show vbModeless
frmBicList.Visible = True
X = frmBicList.Caption
AppActivate X

End Sub



Public Sub frmOptrfD_show()
Dim X As String

frmOpTrfD.Show vbModeless
frmOpTrfD.Visible = True
X = frmOpTrfD.Caption
AppActivate X

End Sub
