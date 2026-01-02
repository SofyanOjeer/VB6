VERSION 5.00
Begin VB.Form frmBdfE 
   AutoRedraw      =   -1  'True
   Caption         =   "Fichier Banque de France Etablissement"
   ClientHeight    =   6345
   ClientLeft      =   1665
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6345
   ScaleWidth      =   7875
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   7560
      Picture         =   "BDFE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   0
      Width           =   400
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   300
      Left            =   6600
      TabIndex        =   21
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdPrécédent 
      Caption         =   "&Précédent"
      Enabled         =   0   'False
      Height          =   300
      Left            =   4680
      TabIndex        =   19
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdSuivant 
      Caption         =   "&Suivant"
      Enabled         =   0   'False
      Height          =   300
      Left            =   5640
      TabIndex        =   20
      Top             =   0
      Width           =   1000
   End
   Begin VB.TextBox txtBdfE10CP 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2250
      TabIndex        =   13
      Top             =   4605
      Width           =   615
   End
   Begin VB.TextBox txtBdfE12 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2235
      TabIndex        =   16
      Top             =   5310
      Width           =   735
   End
   Begin VB.TextBox txtBdfE10B 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2235
      TabIndex        =   11
      Top             =   3885
      Width           =   3975
   End
   Begin VB.TextBox txtBdfE10A 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2265
      TabIndex        =   10
      Top             =   3525
      Width           =   3975
   End
   Begin VB.TextBox txtbdfE14 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2265
      TabIndex        =   18
      Top             =   6010
      Width           =   1575
   End
   Begin VB.TextBox txtbdfE13 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2250
      TabIndex        =   17
      Top             =   5685
      Width           =   255
   End
   Begin VB.TextBox txtbdfE11 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2250
      TabIndex        =   15
      Top             =   4950
      Width           =   855
   End
   Begin VB.TextBox txtBdfE10D 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2970
      TabIndex        =   14
      Top             =   4620
      Width           =   3255
   End
   Begin VB.TextBox txtBdfE10C 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2235
      TabIndex        =   12
      Top             =   4260
      Width           =   3975
   End
   Begin VB.TextBox txtbdfE9 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2250
      TabIndex        =   9
      Top             =   3165
      Width           =   1575
   End
   Begin VB.TextBox txtbdfE8 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2250
      TabIndex        =   8
      Top             =   2790
      Width           =   1575
   End
   Begin VB.TextBox txtbdfE7 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2250
      TabIndex        =   7
      Top             =   2430
      Width           =   1575
   End
   Begin VB.TextBox txtbdfE6 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2235
      TabIndex        =   6
      Top             =   2070
      Width           =   495
   End
   Begin VB.TextBox txtbdfE5 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2250
      TabIndex        =   5
      Top             =   1710
      Width           =   255
   End
   Begin VB.TextBox txtBdfE4B 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2235
      TabIndex        =   4
      Top             =   1350
      Width           =   1815
   End
   Begin VB.TextBox txtbdfE4 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2250
      TabIndex        =   3
      Top             =   990
      Width           =   4935
   End
   Begin VB.TextBox txtbdfE3 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2250
      TabIndex        =   2
      Top             =   645
      Width           =   255
   End
   Begin VB.TextBox txtbdfE2 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2235
      TabIndex        =   1
      Top             =   315
      Width           =   975
   End
   Begin VB.TextBox txtbdfE1 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2220
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblBdfE2 
      Caption         =   "Code établissement"
      Height          =   255
      Left            =   50
      TabIndex        =   39
      Top             =   390
      Width           =   1815
   End
   Begin VB.Label lblBdfE1 
      Caption         =   "code enregistrement"
      Height          =   255
      Left            =   50
      TabIndex        =   38
      Top             =   75
      Width           =   2055
   End
   Begin VB.Label lblBdfE14 
      Caption         =   "date prise effet identifiant"
      Height          =   255
      Left            =   50
      TabIndex        =   37
      Top             =   6075
      Width           =   1965
   End
   Begin VB.Label lblBdfE13 
      Caption         =   "option identifiant invariant"
      Height          =   255
      Left            =   50
      TabIndex        =   36
      Top             =   5715
      Width           =   1905
   End
   Begin VB.Label lblBdfE12 
      Caption         =   "code établissement nouveau"
      Height          =   255
      Left            =   50
      TabIndex        =   35
      Top             =   5325
      Width           =   2055
   End
   Begin VB.Label lblBdfE11 
      Caption         =   "argument de tri"
      Height          =   255
      Left            =   50
      TabIndex        =   34
      Top             =   4995
      Width           =   1995
   End
   Begin VB.Label lblBdfE10D 
      Caption         =   "adresse postale ligne 6"
      Height          =   255
      Left            =   50
      TabIndex        =   33
      Top             =   4635
      Width           =   1995
   End
   Begin VB.Label lblBdfE10C 
      Caption         =   "adresse postale ligne 5"
      Height          =   255
      Left            =   50
      TabIndex        =   32
      Top             =   4305
      Width           =   1995
   End
   Begin VB.Label lblBdfE10B 
      Caption         =   "adresse postale ligne 4"
      Height          =   255
      Left            =   50
      TabIndex        =   31
      Top             =   3915
      Width           =   1995
   End
   Begin VB.Label lblBdfE10A 
      Caption         =   "adresse postale ligne 3"
      Height          =   255
      Left            =   50
      TabIndex        =   30
      Top             =   3555
      Width           =   1995
   End
   Begin VB.Label lblBdfE9 
      Caption         =   "date de disparition FGD"
      Height          =   255
      Left            =   50
      TabIndex        =   29
      Top             =   3165
      Width           =   1995
   End
   Begin VB.Label lblBdfE8 
      Caption         =   "date fin de validité"
      Height          =   255
      Left            =   50
      TabIndex        =   28
      Top             =   2805
      Width           =   1995
   End
   Begin VB.Label lblBdfE7 
      Caption         =   "date début de validité"
      Height          =   255
      Left            =   50
      TabIndex        =   27
      Top             =   2460
      Width           =   1995
   End
   Begin VB.Label lblBdfE6 
      Caption         =   "organe représentatif"
      Height          =   255
      Left            =   50
      TabIndex        =   26
      Top             =   2130
      Width           =   1995
   End
   Begin VB.Label lblBdfE5 
      Caption         =   "code nature activité"
      Height          =   255
      Left            =   50
      TabIndex        =   25
      Top             =   1740
      Width           =   1995
   End
   Begin VB.Label lblBdfE4B 
      Caption         =   "dénomination 10"
      Height          =   2550
      Left            =   50
      TabIndex        =   24
      Top             =   1395
      Width           =   1995
   End
   Begin VB.Label lblBdfE4 
      Caption         =   "dénomination  40"
      Height          =   255
      Left            =   50
      TabIndex        =   23
      Top             =   1035
      Width           =   1995
   End
   Begin VB.Label lblBdfE3 
      Caption         =   "code situation"
      Height          =   255
      Left            =   50
      TabIndex        =   22
      Top             =   720
      Width           =   1995
   End
End
Attribute VB_Name = "frmBdfE"
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

If arrBdfENb > 0 Then frmBdf.cmdBdfEOk_Click

Unload Me

End Sub

'---------------------------------------------------------
Private Sub cmdPrécédent_Click()
'---------------------------------------------------------

arrBdfEIndex = arrBdfEIndex - 1
frmBdf.lstBdfE.ListIndex = frmBdf.lstBdfE.ListIndex - 1
Call frmBdfEInit
End Sub

'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
Dim Msg As String

Msg = Format$(arrBdfEIndex, "000000") & Format$(arrBdfEIndex, "000000")
prtBdfE Msg
End Sub

'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me

End Sub

'---------------------------------------------------------
Private Sub cmdSuivant_Click()
'---------------------------------------------------------

arrBdfEIndex = arrBdfEIndex + 1
frmBdf.lstBdfE.ListIndex = frmBdf.lstBdfE.ListIndex + 1
Call frmBdfEInit
End Sub

'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
frmBdf.sstabBdf.Tab = 0
Call frmBdfEInit

End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 27: cmdQuit_Click
'    Case Is = 34: cmdPageNext_Click
'    Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
'    Case Is = 13: SendKeys "{TAB}"
End Select
'lblErrMsg.Visible = False

End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------

Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)

End Sub





'---------------------------------------------------------
Private Sub frmBdfEInit()
'---------------------------------------------------------

If arrBdfENb > 1 And arrBdfEIndex < arrBdfENb Then
    cmdSuivant.Enabled = True
Else
    cmdSuivant.Enabled = False
End If

If arrBdfENb > 1 And arrBdfEIndex > 1 Then
    cmdPrécédent.Enabled = True
Else
    cmdPrécédent.Enabled = False
End If

If arrBdfENb > 0 Then
    txtbdfE1 = arrBdfE(arrBdfEIndex).BdfE1
    txtbdfE2 = arrBdfE(arrBdfEIndex).BdfE2
    txtbdfE3 = arrBdfE(arrBdfEIndex).BdfE3
    txtbdfE4 = arrBdfE(arrBdfEIndex).BdfE4
    txtBdfE4B = arrBdfE(arrBdfEIndex).BdfE4B
    txtbdfE5 = arrBdfE(arrBdfEIndex).BdfE5
    txtbdfE6 = arrBdfE(arrBdfEIndex).BdfE6
    Call dateDsp(arrBdfE(arrBdfEIndex).BdfE7, txtbdfE7)
    Call dateDsp(arrBdfE(arrBdfEIndex).BdfE8, txtbdfE8)
    Call dateDsp(arrBdfE(arrBdfEIndex).BdfE9, txtbdfE9)
    txtBdfE10A = arrBdfE(arrBdfEIndex).BdfE10A
    txtBdfE10B = arrBdfE(arrBdfEIndex).BdfE10B
    txtBdfE10C = arrBdfE(arrBdfEIndex).BdfE10C
    txtBdfE10CP = arrBdfE(arrBdfEIndex).BdfE10CP
    txtBdfE10D = arrBdfE(arrBdfEIndex).BdfE10D
    txtbdfE11 = arrBdfE(arrBdfEIndex).BdfE11
    txtBdfE12 = arrBdfE(arrBdfEIndex).BdfE12
    txtbdfE13 = arrBdfE(arrBdfEIndex).BdfE13
    Call dateDsp(arrBdfE(arrBdfEIndex).BdfE14, txtbdfE14)

End If

End Sub



