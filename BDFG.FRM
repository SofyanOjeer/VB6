VERSION 5.00
Begin VB.Form frmBdfG 
   AutoRedraw      =   -1  'True
   Caption         =   "Banque de France : Guichet"
   ClientHeight    =   6300
   ClientLeft      =   1455
   ClientTop       =   345
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6300
   ScaleWidth      =   8085
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   7680
      Picture         =   "BDFG.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   0
      Width           =   400
   End
   Begin VB.TextBox txtBdfG14CP 
      Height          =   285
      Left            =   2325
      TabIndex        =   53
      Top             =   3780
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   300
      Left            =   6720
      TabIndex        =   52
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdSuivant 
      Caption         =   "&Suivant"
      Enabled         =   0   'False
      Height          =   300
      Left            =   5760
      TabIndex        =   51
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdPrécédent 
      Caption         =   "&Précédent"
      Enabled         =   0   'False
      Height          =   300
      Left            =   4800
      TabIndex        =   50
      Top             =   0
      Width           =   1000
   End
   Begin VB.TextBox txtbdfG3 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2325
      TabIndex        =   2
      Top             =   345
      Width           =   735
   End
   Begin VB.TextBox txtBdfG23 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2325
      TabIndex        =   25
      Top             =   5910
      Width           =   1095
   End
   Begin VB.TextBox txtBdfG22 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6345
      TabIndex        =   24
      Top             =   5565
      Width           =   615
   End
   Begin VB.TextBox txtBdfG21 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2340
      TabIndex        =   23
      Top             =   5565
      Width           =   255
   End
   Begin VB.TextBox txtBdfG20 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2325
      TabIndex        =   22
      Top             =   5175
      Width           =   3015
   End
   Begin VB.TextBox txtBdfG19 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6345
      TabIndex        =   21
      Top             =   4500
      Width           =   735
   End
   Begin VB.TextBox txtBdfG18 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2325
      TabIndex        =   20
      Top             =   4845
      Width           =   1455
   End
   Begin VB.TextBox txtBdfG4 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6360
      TabIndex        =   3
      Top             =   345
      Width           =   255
   End
   Begin VB.TextBox txtBdfG12 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6345
      TabIndex        =   12
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtBdfG16 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2325
      TabIndex        =   18
      Top             =   4500
      Width           =   1455
   End
   Begin VB.TextBox txtBdfG14D 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3330
      TabIndex        =   16
      Top             =   3780
      Width           =   3975
   End
   Begin VB.TextBox txtBdfG13 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2325
      TabIndex        =   11
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtBdfG10 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2325
      TabIndex        =   9
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtBdfG17 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6345
      TabIndex        =   19
      Top             =   4155
      Width           =   1455
   End
   Begin VB.TextBox txtBdfG15 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2325
      TabIndex        =   17
      Top             =   4155
      Width           =   1455
   End
   Begin VB.TextBox txtBdfG14C 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2325
      TabIndex        =   15
      Top             =   3450
      Width           =   3975
   End
   Begin VB.TextBox txtbdfG14B 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2325
      TabIndex        =   14
      Top             =   3090
      Width           =   3975
   End
   Begin VB.TextBox txtBdfG14A 
      ForeColor       =   &H00FF00FF&
      Height          =   285
      Left            =   2325
      TabIndex        =   13
      Top             =   2760
      Width           =   3975
   End
   Begin VB.TextBox txtBdfG11 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6345
      TabIndex        =   10
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtBdfG8 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6345
      TabIndex        =   7
      Top             =   1335
      Width           =   735
   End
   Begin VB.TextBox txtBdfG9 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2325
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtBdfG7 
      ForeColor       =   &H00FF00FF&
      Height          =   285
      Left            =   2325
      TabIndex        =   6
      Top             =   1365
      Width           =   735
   End
   Begin VB.TextBox txtBdfG6 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2325
      TabIndex        =   5
      Top             =   1050
      Width           =   2535
   End
   Begin VB.TextBox txtBdfG5 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2325
      TabIndex        =   4
      Top             =   705
      Width           =   4935
   End
   Begin VB.TextBox txtBdfG2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2325
      TabIndex        =   0
      Top             =   -15
      Width           =   735
   End
   Begin VB.Label lblBdfG23 
      Caption         =   "nom comptoir BDF"
      Height          =   285
      Left            =   50
      TabIndex        =   49
      Top             =   5940
      Width           =   2175
   End
   Begin VB.Label lblBdfG22 
      Caption         =   "code règlement des crédits"
      Height          =   225
      Left            =   3975
      TabIndex        =   48
      Top             =   5625
      Width           =   2415
   End
   Begin VB.Label lblBdfG21 
      Caption         =   "nature guichet"
      Height          =   285
      Left            =   45
      TabIndex        =   47
      Top             =   5625
      Width           =   2220
   End
   Begin VB.Label lblBdfG20 
      Caption         =   "lib abréviatif dom RIB"
      Height          =   285
      Left            =   50
      TabIndex        =   46
      Top             =   5250
      Width           =   2130
   End
   Begin VB.Label lblBdfG19 
      Caption         =   "code routage guichet domicil"
      Height          =   225
      Left            =   3975
      TabIndex        =   45
      Top             =   4560
      Width           =   2265
   End
   Begin VB.Label lblBdfG2 
      Caption         =   "code établissement"
      Height          =   285
      Left            =   50
      TabIndex        =   44
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label lblBdfG18 
      Caption         =   "date disparition du FGD"
      Height          =   285
      Left            =   50
      TabIndex        =   43
      Top             =   4905
      Width           =   2145
   End
   Begin VB.Label lblBdfG17 
      Caption         =   "date fin de validité"
      Height          =   225
      Left            =   3975
      TabIndex        =   42
      Top             =   4200
      Width           =   2040
   End
   Begin VB.Label lblBdfG16 
      Caption         =   "date début de validité"
      Height          =   285
      Left            =   50
      TabIndex        =   41
      Top             =   4560
      Width           =   2250
   End
   Begin VB.Label lblBdfG15 
      Caption         =   "adresse SWIFT"
      Height          =   285
      Left            =   50
      TabIndex        =   40
      Top             =   4200
      Width           =   2265
   End
   Begin VB.Label lblBdfG14D 
      Caption         =   "adresse postale ligne 6"
      Height          =   285
      Left            =   50
      TabIndex        =   39
      Top             =   3840
      Width           =   2220
   End
   Begin VB.Label lblBdfG14C 
      Caption         =   "adresse postale ligne 5"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   50
      TabIndex        =   38
      Top             =   3480
      Width           =   2205
   End
   Begin VB.Label lblBdfG14B 
      Caption         =   "adresse postale ligne 4"
      Height          =   285
      Left            =   50
      TabIndex        =   37
      Top             =   3090
      Width           =   2235
   End
   Begin VB.Label lblBdfG14A 
      Caption         =   "adresse postale ligne 3"
      Height          =   285
      Left            =   50
      TabIndex        =   36
      Top             =   2760
      Width           =   2325
   End
   Begin VB.Label lblBdfG13 
      Caption         =   "comptoir BDF"
      Height          =   285
      Left            =   50
      TabIndex        =   35
      Top             =   2400
      Width           =   2235
   End
   Begin VB.Label lblBdfG12 
      Caption         =   "code recouvreur"
      Height          =   225
      Left            =   3975
      TabIndex        =   34
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label lblBdfG11 
      Caption         =   "code localité de compens 2"
      Height          =   225
      Left            =   3975
      TabIndex        =   33
      Top             =   1680
      Width           =   2385
   End
   Begin VB.Label lblBdfG10 
      Caption         =   "code localité de compens 1"
      Height          =   285
      Left            =   50
      TabIndex        =   32
      Top             =   2040
      Width           =   2235
   End
   Begin VB.Label lblBdfG9 
      Caption         =   "code géographique"
      Height          =   285
      Left            =   50
      TabIndex        =   31
      Top             =   1680
      Width           =   2205
   End
   Begin VB.Label lblBdfG8 
      Caption         =   "code guichet nouveau"
      Height          =   225
      Left            =   3975
      TabIndex        =   30
      Top             =   1380
      Width           =   1995
   End
   Begin VB.Label lblBdfG7 
      Caption         =   "code établissement nouveau"
      Height          =   285
      Left            =   50
      TabIndex        =   29
      Top             =   1380
      Width           =   2265
   End
   Begin VB.Label lblBdfG6 
      Caption         =   "nom du guichet"
      Height          =   285
      Left            =   50
      TabIndex        =   28
      Top             =   1065
      Width           =   2235
   End
   Begin VB.Label lblBdfG5 
      Caption         =   "dénomination particuliculière"
      Height          =   285
      Left            =   45
      TabIndex        =   27
      Top             =   720
      Width           =   2205
   End
   Begin VB.Label lblBdfG4 
      Caption         =   "code situation"
      Height          =   225
      Left            =   3975
      TabIndex        =   26
      Top             =   330
      Width           =   2175
   End
   Begin VB.Label lblBdfG3 
      Caption         =   "code guichet"
      Height          =   285
      Left            =   50
      TabIndex        =   1
      Top             =   345
      Width           =   2205
   End
End
Attribute VB_Name = "frmBdfG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean

'---------------------------------------------------------
Private Sub cmdOk_Click()
'---------------------------------------------------------

If arrBdfGNb > 0 Then frmBdf.cmdBdfGOk_Click
Unload Me

End Sub

'---------------------------------------------------------
Private Sub cmdPrécédent_Click()
'---------------------------------------------------------

arrBdfGIndex = arrBdfGIndex - 1
frmBdf.lstBdfG.ListIndex = frmBdf.lstBdfG.ListIndex - 1
Call frmBdfGInit
End Sub

'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
Dim Msg As String

Msg = Format$(arrBdfGIndex, "000000") & Format$(arrBdfGIndex, "000000")
prtBdfG Msg

End Sub

'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me

End Sub

'---------------------------------------------------------
Private Sub cmdSuivant_Click()
'---------------------------------------------------------

arrBdfGIndex = arrBdfGIndex + 1
frmBdf.lstBdfG.ListIndex = frmBdf.lstBdfG.ListIndex + 1
Call frmBdfGInit
End Sub

'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
frmBdf.sstabBdf.Tab = 1
Call frmBdfGInit

End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 27: cmdQuit_Click
'    Case Is = 34: cmdPageNext_Click
'    Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
    Case Is = 13: SendKeys "{TAB}"
End Select
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
Set XForm = Me
Call MeInit(arrTagNb, "Locked")
ReDim arrTag(arrTagNb + 1)

End Sub





'---------------------------------------------------------
Private Sub frmBdfGInit()
'---------------------------------------------------------

If arrBdfGNb > 1 And arrBdfGIndex < arrBdfGNb Then
    cmdSuivant.Enabled = True
Else
    cmdSuivant.Enabled = False
End If

If arrBdfGNb > 1 And arrBdfGIndex > 1 Then
    cmdPrécédent.Enabled = True
Else
    cmdPrécédent.Enabled = False
End If

If arrBdfGNb > 0 Then
    txtBdfG2 = arrBdfG(arrBdfGIndex).BdfG2
    txtbdfG3 = arrBdfG(arrBdfGIndex).BdfG3
    txtBdfG4 = arrBdfG(arrBdfGIndex).BdfG4
    txtBdfG5 = arrBdfG(arrBdfGIndex).BdfG5
    txtBdfG6 = arrBdfG(arrBdfGIndex).BdfG6
    txtBdfG7 = arrBdfG(arrBdfGIndex).BdfG7
    txtBdfG8 = arrBdfG(arrBdfGIndex).BdfG8
    txtBdfG9 = arrBdfG(arrBdfGIndex).BdfG9
    txtBdfG10 = arrBdfG(arrBdfGIndex).BdfG10
    txtBdfG11 = arrBdfG(arrBdfGIndex).BdfG11
    txtBdfG12 = arrBdfG(arrBdfGIndex).BdfG12
    txtBdfG13 = arrBdfG(arrBdfGIndex).BdfG13
    txtBdfG14A = arrBdfG(arrBdfGIndex).BdfG14A
    txtbdfG14B = arrBdfG(arrBdfGIndex).BdfG14B
    txtBdfG14C = arrBdfG(arrBdfGIndex).BdfG14C
    txtBdfG14CP = arrBdfG(arrBdfGIndex).BdfG14CP
    txtBdfG14D = arrBdfG(arrBdfGIndex).BdfG14D
    txtBdfG15 = arrBdfG(arrBdfGIndex).BdfG15
    Call dateDsp(arrBdfG(arrBdfGIndex).BdfG16, txtBdfG16)
    Call dateDsp(arrBdfG(arrBdfGIndex).BdfG17, txtBdfG17)
    Call dateDsp(arrBdfG(arrBdfGIndex).BdfG18, txtBdfG18)
    txtBdfG19 = arrBdfG(arrBdfGIndex).BdfG19
    txtBdfG20 = arrBdfG(arrBdfGIndex).BdfG20
    txtBdfG21 = arrBdfG(arrBdfGIndex).BdfG21
    txtBdfG23 = arrBdfG(arrBdfGIndex).BdfG23
    txtBdfG22 = arrBdfG(arrBdfGIndex).BdfG22
    txtBdfG23 = arrBdfG(arrBdfGIndex).BdfG23
End If

End Sub






