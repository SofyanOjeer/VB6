VERSION 5.00
Begin VB.Form frmElp 
   Caption         =   "Risques"
   ClientHeight    =   6735
   ClientLeft      =   15
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6735
   ScaleWidth      =   9480
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   9000
      Picture         =   "Risques.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   400
   End
   Begin VB.CheckBox chkFrm 
      Caption         =   "Lire source Form"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   6480
      Width           =   1665
   End
   Begin VB.CommandButton cmdRecherche 
      Caption         =   "Recherche"
      Default         =   -1  'True
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1000
   End
   Begin VB.ListBox LstMain 
      Height          =   2400
      Left            =   840
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.TextBox txtNuméro 
      Height          =   300
      Left            =   1050
      MaxLength       =   3
      TabIndex        =   0
      Top             =   540
      Width           =   1035
   End
   Begin VB.Image imgSocSignon 
      Height          =   6750
      Left            =   0
      Top             =   0
      Width           =   9450
   End
   Begin VB.Label lblIn 
      Caption         =   "Numéro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   2
      Top             =   585
      Width           =   855
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ErrMsg"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4560
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmElp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim Msg As String
Dim L As Long


Private Sub cmdPrint_Click()

arrRisques_Test
prtRisquesX " "

End Sub

Private Sub cmdRecherche_Click()

If Not IsNumeric(txtNuméro) Then MsgBox " numèro non numèrique ", vbCritical, " Guichet convention : saisie ": Exit Sub
L = Val(txtNuméro)
If L = 0 Or L > 99 Then MsgBox "Le numèro doit être compris entre 1 à 99  ", vbCritical, " Guichet convention : saisie ": Exit Sub

If optCompteJoint Then frmConventionJoint_Show
If OptCompteIndivis Then frmConventionJoint_Show
If optCompteEntreprise Then frmConventionEntreprise_Show
If optCompteParticulier Then frmConventionParticulier_Show
End Sub


Private Sub cmdRib_Click()

End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 44: frmElpPrt.prtScreen
    Case Is = 13: SendKeys "{TAB}"
End Select
lblErrMsg.Visible = False
End Sub



Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
dateAAmin = 1850

End Sub


Private Sub Form_Unload(Cancel As Integer)
Close
End Sub



Public Sub frmConventionJoint_Show()
srvGuichetConventionJoint.Ouverture
recGuichetConventionJoint.Numéro = L
If Not IsNull(srvGuichetConventionJoint.Lire) Or recGuichetConventionJoint.Numéro = 0 Or recGuichetConventionJoint.Numéro > 99 Then
    srvGuichetConventionJoint.Rec_Init
    recGuichetConventionJoint.Numéro = Val(txtNuméro)
End If
If recGuichetConventionJoint.Nature = "I" Then
    OptCompteIndivis = True
Else
    optCompteJoint = True
End If
frmGuichetConventionJoint.Rec_Display
frmGuichetConventionJoint.Show vbModeless 'vbModal

End Sub

Public Sub frmConventionEntreprise_Show()
srvGuichetConventionEntreprise.Ouverture
recGuichetConventionEntreprise.Numéro = L
If Not IsNull(srvGuichetConventionEntreprise.Lire) Or recGuichetConventionEntreprise.Numéro = 0 Or recGuichetConventionEntreprise.Numéro > 99 Then
    srvGuichetConventionEntreprise.Rec_Init
    recGuichetConventionEntreprise.Numéro = Val(txtNuméro)
End If
frmGuichetConventionEntreprise.Rec_Display
frmGuichetConventionEntreprise.Show vbModeless 'vbModal

End Sub

Public Sub frmConventionParticulier_Show()
srvGuichetConventionParticulier.Ouverture
recGuichetConventionParticulier.Numéro = L
If Not IsNull(srvGuichetConventionParticulier.Lire) Or recGuichetConventionParticulier.Numéro = 0 Or recGuichetConventionParticulier.Numéro > 99 Then
    srvGuichetConventionParticulier.Rec_Init
    recGuichetConventionParticulier.Numéro = Val(txtNuméro)
End If
frmGuichetConventionParticulier.Rec_Display
frmGuichetConventionParticulier.Show vbModeless 'vbModal

End Sub


