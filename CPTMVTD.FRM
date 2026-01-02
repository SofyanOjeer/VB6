VERSION 5.00
Begin VB.Form frmCptMvtD 
   Caption         =   "Détail D'un Mouvement Comptable"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6315
   ScaleWidth      =   7380
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   0
      TabIndex        =   43
      Top             =   3240
      Width           =   7335
      Begin VB.TextBox txtService 
         Height          =   285
         Left            =   1440
         TabIndex        =   47
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtPièce 
         Height          =   285
         Left            =   2040
         TabIndex        =   46
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtLigne 
         Height          =   285
         Left            =   2880
         TabIndex        =   45
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtCodeOpération 
         Height          =   285
         Left            =   1440
         TabIndex        =   44
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblService 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Service"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   855
      End
      Begin VB.Label libService 
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3480
         TabIndex        =   50
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblCodeOpération 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Code Opération"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label libCodeOpération 
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3480
         TabIndex        =   48
         Top             =   720
         Width           =   2910
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   6840
      Picture         =   "CPTMVTD.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   120
      Width           =   400
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saisie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   28
      Top             =   5280
      Width           =   7300
      Begin VB.TextBox txtAutomatique 
         Height          =   285
         Left            =   6600
         TabIndex        =   39
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtProgrammeSaisie 
         Height          =   285
         Left            =   4440
         TabIndex        =   35
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtLot 
         Height          =   285
         Left            =   3600
         TabIndex        =   33
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtOpérateurSaisie 
         Height          =   285
         Left            =   1560
         TabIndex        =   31
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtAmjSaisie 
         Height          =   285
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   1100
      End
      Begin VB.Label lblAutomatique 
         Caption         =   "Automatique"
         Height          =   225
         Left            =   6240
         TabIndex        =   38
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblAmjSaisie 
         Caption         =   "Le :"
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblProgrammeSaisie 
         Caption         =   "Programme Saisie"
         Height          =   255
         Left            =   4680
         TabIndex        =   34
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label lblLot 
         Caption         =   "N° de Lot"
         Height          =   255
         Left            =   3480
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblOpérateurSaisie 
         Caption         =   "Par :"
         Height          =   255
         Left            =   1920
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   0
      TabIndex        =   13
      Top             =   4320
      Width           =   7300
      Begin VB.TextBox txtCptComplémentaire 
         Height          =   285
         Left            =   5160
         TabIndex        =   26
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtEditionAvis 
         Height          =   285
         Left            =   6600
         TabIndex        =   23
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtExonéré 
         Height          =   285
         Left            =   5760
         TabIndex        =   21
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtAmjOpération 
         Height          =   285
         Left            =   3600
         TabIndex        =   19
         Top             =   480
         Width           =   1100
      End
      Begin VB.TextBox txtAmjTraitement 
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1100
      End
      Begin VB.TextBox txtAmjValeur 
         Height          =   285
         Left            =   1900
         TabIndex        =   17
         Top             =   480
         Width           =   1100
      End
      Begin VB.Label lblCptComplémentaire 
         Caption         =   "CptSup"
         Height          =   255
         Left            =   5040
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblEditionAvis 
         Caption         =   "Avis"
         Height          =   255
         Left            =   6480
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblExonéré 
         Caption         =   "Exo"
         Height          =   255
         Left            =   5760
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblAmjOpération 
         Caption         =   "Date D'opération"
         Height          =   255
         Left            =   3600
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblAmjValeur 
         Caption         =   "Date de Valeur"
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblAmjTraitement 
         Caption         =   "Date de Traitement"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2055
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   7300
      Begin VB.TextBox txtSoldeVeille 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   42
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtMt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   37
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtLibellé 
         Height          =   495
         Left            =   1440
         TabIndex        =   12
         Top             =   1440
         Width           =   4815
      End
      Begin VB.Label lblEuro 
         Alignment       =   1  'Right Justify
         Caption         =   "Contre-valeur en euros"
         Height          =   255
         Left            =   4080
         TabIndex        =   54
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblMt_Euro 
         Alignment       =   1  'Right Justify
         Caption         =   "-"
         Height          =   255
         Left            =   4080
         TabIndex        =   53
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblSoldeVeille_Euro 
         Alignment       =   1  'Right Justify
         Caption         =   "-"
         Height          =   255
         Left            =   4080
         TabIndex        =   52
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblSoldeVeille 
         Caption         =   "Solde veille"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   975
      End
      Begin VB.Label libDevise 
         Alignment       =   1  'Right Justify
         Caption         =   "-"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   240
         Width           =   2280
      End
      Begin VB.Label lblLibellé 
         Caption         =   "Libellé"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblMt 
         Caption         =   "Montant"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7300
      Begin VB.CommandButton cmdLineNext 
         Caption         =   ">>"
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
         Left            =   6060
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   600
      End
      Begin VB.CommandButton cmdLinePrior 
         Caption         =   "<<"
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
         Left            =   5460
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   600
      End
      Begin VB.TextBox txtCompte 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   600
         Width           =   1290
      End
      Begin VB.TextBox txtAgence 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   550
      End
      Begin VB.TextBox txtSociété 
         Height          =   285
         Left            =   1455
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label libIntitulé 
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3015
         TabIndex        =   8
         Top             =   615
         Width           =   4095
      End
      Begin VB.Label lblCompte 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Compte"
         Height          =   255
         Left            =   105
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.Label libSociété 
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblSociété 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Société, Agence"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Menu mnucmdPrint 
      Caption         =   "mnucmdPrint"
      Visible         =   0   'False
      Begin VB.Menu mnucmdPrint_Avis 
         Caption         =   "imprimer Avis"
      End
   End
End
Attribute VB_Name = "frmCptMvtD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor

Dim recCptMvt As typeCptMvt

''Dim recOpCpt As typeOpCpt

Dim shapeCptMvt_Top As Integer
Public Sub MouseMoveActiveControl_Set(C As Control)
If MouseMoveActiveControl_Name <> C.Name Then
    MouseMoveActiveControl_Reset
    If Not C.Enabled Then
        MouseMoveActiveControl_Name = ""
    Else
        MouseMoveActiveControl_Name = C.Name
        If TypeOf C Is CommandButton Then
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            MouseMoveActiveControl.ForeColor = C.ForeColor
            C.ForeColor = MouseMoveUsr.ForeColor
        End If
    End If
End If

End Sub


Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
        If TypeOf xobj Is CommandButton Then
            xobj.BackColor = MouseMoveActiveControl.BackColor
        Else
            xobj.ForeColor = MouseMoveActiveControl.ForeColor
        End If
        Exit For
    End If
Next xobj

End Sub

'---------------------------------------------------------
Public Sub frmCptMvtDInit()
'---------------------------------------------------------
dspCptMvtD

End Sub


'---------------------------------------------------------
Public Sub dspCptMvtD()
'---------------------------------------------------------
Dim X As String

shapeCptMvt_Top = frmCptMvt.picCptMvt.Height - frmCptMvt.Me_picLineHeight

If frmCptMvt.Amont Then
    dspCptMvtDEnabledAmont
Else
    dspCptMvtDEnabled
End If
If arrCptMvtIndex > arrCptMvtNb Then arrCptMvtIndex = arrCptMvtNb
If arrCptMvtIndex < 1 Then arrCptMvtIndex = 1
recCptMvt = arrCptMvt(arrCptMvtIndex)

txtSociété = recCptMvt.Société
libSociété = socName

txtAgence = recCptMvt.Agence
libDevise = recCptMvt.Devise
txtCompte = Format$(recCptMvt.Compte, "@@@@@.@@@.@@.@")
libIntitulé = arrCompte(arrCompteIndex).Intitulé & arrCompte(arrCompteIndex).Intitulé2
txtSoldeVeille.ForeColor = IIf(recCptMvt.SoldeVeille < 0, vbRed, vbBlue)
txtSoldeVeille.Locked = True
txtSoldeVeille.Enabled = True
txtSoldeVeille = Format$(Abs(recCptMvt.SoldeVeille), "### ### ### ### ##0.00")
CV_X2.DeviseIso = "EUR"
CV_X1.DeviseN = Format$(recCptMvt.Devise, "000")
CV_X1.DeviseIso = ""
CV_X1.Montant = Abs(recCptMvt.SoldeVeille)
Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X)
lblSoldeVeille_Euro.ForeColor = txtSoldeVeille.ForeColor
lblSoldeVeille_Euro = Format$(CV_X2.Montant, "#### ### ### ### ##0.00")

txtCodeOpération = recCptMvt.CodeOpération
libCodeOpération = DicLib(27, recCptMvt.CodeOpération)

txtService = recCptMvt.Service
libService = DicLib(4, recCptMvt.Service)

txtMt.ForeColor = IIf(recCptMvt.Mt < 0, vbRed, vbBlue)
txtMt.Locked = True
txtMt.Enabled = True
txtMt = Format$(Abs(recCptMvt.Mt), "### ### ### ### ##0.00")
CV_X1.Montant = Abs(recCptMvt.Mt)
Call CV_Calc(CV_X1, CV_X2, CV_X3)
lblMt_Euro.ForeColor = txtMt.ForeColor
lblMt_Euro = Format$(CV_X2.Montant, "#### ### ### ### ##0.00")
libDevise = Trim(CV_X1.DeviseLibellé)
txtAmjSaisie = dateImp(recCptMvt.AmjSaisie)
txtAmjOpération = dateImp(recCptMvt.AmjOpération)
txtAmjValeur = dateImp(recCptMvt.AmjValeur)
txtPièce = recCptMvt.Pièce
txtLigne = recCptMvt.Ligne
txtLot = recCptMvt.Lot
txtLibellé = recCptMvt.Libellé
txtExonéré = recCptMvt.Exonéré
txtAutomatique = recCptMvt.Automatique

txtEditionAvis = recCptMvt.EditionAvis
txtCptComplémentaire = recCptMvt.CptComplémentaire
txtAmjTraitement = dateImp(recCptMvt.AmjTraitement)
txtProgrammeSaisie = recCptMvt.ProgrammeSaisie
txtOpérateurSaisie = recCptMvt.OpérateurSaisie

End Sub

Private Sub cmdLineNext_Click()
'---------------------------------------------------------

frmCptMvt.shapeCptMvt.Top = frmCptMvt.shapeCptMvt.Top + frmCptMvt.Me_picLineHeight
frmCptMvt.shapeCptMvt.Visible = IIf(frmCptMvt.shapeCptMvt.Top < shapeCptMvt_Top, True, False)
If frmCptMvt.Amont Then
    PagePrior
Else
    PageNext
End If

End Sub


Private Sub cmdLineNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdLineNext

End Sub


Private Sub cmdLinePrior_Click()
'---------------------------------------------------------

frmCptMvt.shapeCptMvt.Top = frmCptMvt.shapeCptMvt.Top - frmCptMvt.Me_picLineHeight
frmCptMvt.shapeCptMvt.Visible = IIf(frmCptMvt.shapeCptMvt.Top > -25, True, False)
If frmCptMvt.Amont Then
    PageNext
Else
    PagePrior
End If

End Sub


Private Sub cmdLinePrior_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdLinePrior

End Sub


'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
mnucmdPrint_Avis = CompteAut.Avis
Me.PopupMenu mnucmdPrint, vbPopupMenuRightButton
End Sub


Public Sub prtAvis()
'recOpCpt_Init recOpCpt
'recOpCpt.Référence = Format(Val(arrCptMvt(arrCptMvtIndex).Pièce), "000000")
'recOpCpt.CodeOpération = arrCptMvt(arrCptMvtIndex).CodeOpération
'recOpCpt.Société = arrCptMvt(arrCptMvtIndex).Société
'recOpCpt.Agence = arrCptMvt(arrCptMvtIndex).Agence
'recOpCpt.Devise = Format(Val(arrCptMvt(arrCptMvtIndex).Devise), "000")
'recOpCpt.Compte = arrCptMvt(arrCptMvtIndex).Compte
'recOpCpt.Brut = arrCptMvt(arrCptMvtIndex).Mt
'recOpCpt.Sens = IIf(arrCptMvt(arrCptMvtIndex).Mt < 0, "D", "C")
'recOpCpt.AmjOpération = arrCptMvt(arrCptMvtIndex).AmjOpération
'recOpCpt.AmjValeur = arrCptMvt(arrCptMvtIndex).AmjValeur
'recOpCpt.Libellé = arrCptMvt(arrCptMvtIndex).Libellé
'recOpCpt.optAvis = "1"
'prtAvisX recOpCpt
prtAvis_CptMvt arrCptMvt(arrCptMvtIndex), " "
End Sub


'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me

End Sub


Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

Private Sub Form_Activate()
Set XForm = Me

End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 27: cmdQuit_Click
    Case Is = 40: cmdLineNext_Click
    Case Is = 38: cmdLinePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
    Case Is = 13: SendKeys "{TAB}"
End Select
'lblErrMsg.Visible = False

End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------

Set XForm = Me
Call MeInit(arrTagNb, "Locked")
ReDim arrTag(arrTagNb + 1)

Exit Sub

Frame1.BackColor = vbGrayText
Frame2.BackColor = vbGrayText
Frame3.BackColor = vbGrayText
Frame4.BackColor = vbGrayText
lblSociété.BackColor = vbGrayText
lblCompte.BackColor = vbGrayText
lblService.BackColor = vbGrayText
lblCodeOpération.BackColor = vbGrayText
lblMt.BackColor = vbGrayText
lblAmjTraitement.BackColor = vbGrayText
lblAmjValeur.BackColor = vbGrayText
lblAmjOpération.BackColor = vbGrayText
lblCptComplémentaire.BackColor = vbGrayText
lblExonéré.BackColor = vbGrayText
lblEditionAvis.BackColor = vbGrayText
lblAmjSaisie.BackColor = vbGrayText
lblOpérateurSaisie.BackColor = vbGrayText
lblProgrammeSaisie.BackColor = vbGrayText
lblAutomatique.BackColor = vbGrayText
libSociété.BackColor = vbGrayText
libCodeOpération.BackColor = vbGrayText
libService.BackColor = vbGrayText
libIntitulé.BackColor = vbGrayText
lblLot.BackColor = vbGrayText
libDevise.BackColor = vbGrayText
lblLibellé.BackColor = vbGrayText
Me.BackColor = vbGrayText
End Sub


'-------------------------------------------------'
Public Sub PageNext()
'-------------------------------------------------'

If arrCptMvtIndex < arrCptMvtNb Then
    arrCptMvtIndex = arrCptMvtIndex + 1
    dspCptMvtD
End If

End Sub

'-------------------------------------------------'
Public Sub PagePrior()
'-------------------------------------------------'
If arrCptMvtIndex > 1 Then
    arrCptMvtIndex = arrCptMvtIndex - 1
    dspCptMvtD
End If

End Sub

'---------------------------------------------------------
Public Sub dspCptMvtDEnabled()
'---------------------------------------------------------

If arrCptMvtIndex > 1 Then
    cmdLinePrior.Enabled = True
Else
    cmdLinePrior.Enabled = False
End If

If arrCptMvtIndex < arrCptMvtNb Then
    cmdLineNext.Enabled = True
Else
    cmdLineNext.Enabled = False
End If

End Sub

'---------------------------------------------------------
Public Sub dspCptMvtDEnabledAmont()
'---------------------------------------------------------

If arrCptMvtIndex > 1 Then
    cmdLineNext.Enabled = True
Else
    cmdLineNext.Enabled = False
End If

If arrCptMvtIndex < arrCptMvtNb Then
    cmdLinePrior.Enabled = True
Else
    cmdLinePrior.Enabled = False
End If

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub mnucmdPrint_Avis_Click()
prtAvis
End Sub


