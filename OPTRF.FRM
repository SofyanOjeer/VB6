VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOpTrf 
   AutoRedraw      =   -1  'True
   Caption         =   "OpTrf : menu général"
   ClientHeight    =   6375
   ClientLeft      =   15
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6375
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   9000
      Picture         =   "OPTRF.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   400
   End
   Begin VB.CommandButton cmdSaisir 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Saisie"
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
      Left            =   3200
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1000
   End
   Begin VB.PictureBox picList 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   5895
      Left            =   0
      ScaleHeight     =   5835
      ScaleWidth      =   9315
      TabIndex        =   5
      Top             =   360
      Width           =   9375
      Begin VB.Frame fraOpStat 
         Height          =   5775
         Left            =   1320
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   7935
         Begin VB.Frame fraOpStatImpression 
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
            Height          =   2895
            Left            =   120
            TabIndex        =   12
            Top             =   2640
            Width           =   7575
            Begin VB.CommandButton cmdOpStatOk 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Ok"
               Height          =   400
               Left            =   6600
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   1920
               Width           =   600
            End
            Begin VB.CommandButton cmdOpStatQuit 
               BackColor       =   &H00C0C0FF&
               Caption         =   "X"
               Height          =   400
               Left            =   6600
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   1080
               Width           =   600
            End
            Begin VB.CheckBox chkOpStatDétail 
               Caption         =   "imprimer détail"
               Height          =   495
               Left            =   6120
               TabIndex        =   39
               Top             =   360
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.Frame fraOpStatSort2 
               Caption         =   "Tri 2/2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2535
               Left            =   3240
               TabIndex        =   32
               Top             =   240
               Width           =   2775
               Begin VB.OptionButton optSort2Devise 
                  Caption         =   "devise"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   38
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.OptionButton optSort2Sender 
                  Caption         =   "émetteur"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   37
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.OptionButton optSort2Correspondant 
                  Caption         =   "correspondant"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   36
                  Top             =   1080
                  Value           =   -1  'True
                  Width           =   1455
               End
               Begin VB.OptionButton optSort2Receiver 
                  Caption         =   "destinataire"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   35
                  Top             =   1440
                  Width           =   1455
               End
               Begin VB.OptionButton optSort2Référence 
                  Caption         =   "n° dossier"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   34
                  Top             =   1800
                  Width           =   1575
               End
               Begin VB.CheckBox chkOpStatSort2 
                  Caption         =   "imprimer sous-total"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   33
                  Top             =   2160
                  Value           =   1  'Checked
                  Width           =   2175
               End
            End
            Begin VB.Frame fraOpStatSort1 
               Caption         =   "Tri 1/2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2535
               Left            =   240
               TabIndex        =   15
               Top             =   240
               Width           =   2775
               Begin VB.CheckBox chkOpStatSort1 
                  Caption         =   "imprimer sous-total"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   31
                  Top             =   2160
                  Value           =   1  'Checked
                  Width           =   2175
               End
               Begin VB.OptionButton optSort1Référence 
                  Caption         =   "n° dossier"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   30
                  Top             =   1800
                  Width           =   1575
               End
               Begin VB.OptionButton optSort1Receiver 
                  Caption         =   "destinataire"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   29
                  Top             =   1440
                  Width           =   1455
               End
               Begin VB.OptionButton optSort1Correspondant 
                  Caption         =   "correspondant"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   28
                  Top             =   1080
                  Width           =   1455
               End
               Begin VB.OptionButton optSort1Sender 
                  Caption         =   "émetteur"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   27
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.OptionButton optSort1Devise 
                  Caption         =   "devise"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   26
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   1815
               End
            End
         End
         Begin VB.Frame fraOpStatFiltre 
            Caption         =   "Sélection"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   7575
            Begin VB.Frame fraOpStatSélection 
               Caption         =   "inclure les opérations"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1935
               Left            =   4320
               TabIndex        =   14
               Top             =   240
               Width           =   3135
               Begin VB.CheckBox chkOpStatCompta 
                  Caption         =   "déjà comptabilisé"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   22
                  Top             =   1080
                  Value           =   1  'Checked
                  Width           =   2175
               End
               Begin VB.CheckBox chkOpStatàCompta 
                  Caption         =   "en attente de comptabilisation"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   21
                  Top             =   720
                  Value           =   1  'Checked
                  Width           =   2655
               End
               Begin VB.CheckBox chkOpStatàValider 
                  Caption         =   "en attente de validation"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   20
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   2655
               End
               Begin VB.Label lblOpStatNb 
                  Caption         =   " -"
                  Height          =   255
                  Left            =   480
                  TabIndex        =   42
                  Top             =   1560
                  Width           =   2055
               End
            End
            Begin VB.Frame fraOpStatPériode 
               Caption         =   "Période"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1935
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   4095
               Begin VB.OptionButton optOpStatMois 
                  Caption         =   "mois"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   19
                  Top             =   1080
                  Width           =   1695
               End
               Begin VB.OptionButton optOpStatManuel 
                  Caption         =   "du"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   18
                  Top             =   1440
                  Width           =   615
               End
               Begin VB.OptionButton optOpStatVeille 
                  Caption         =   "veille"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   17
                  Top             =   720
                  Width           =   1695
               End
               Begin VB.OptionButton optOpStatJour 
                  Caption         =   "jour"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   16
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   1695
               End
               Begin MSMask.MaskEdBox txtAmjMin 
                  Height          =   300
                  Left            =   840
                  TabIndex        =   23
                  Top             =   1440
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  _Version        =   393216
                  MaxLength       =   14
                  Mask            =   "## - ## - ####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox txtAmjMax 
                  Height          =   300
                  Left            =   2640
                  TabIndex        =   24
                  Top             =   1440
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  _Version        =   393216
                  MaxLength       =   14
                  Mask            =   "## - ## - ####"
                  PromptChar      =   "_"
               End
               Begin VB.Label lblOpStatAu 
                  Caption         =   "au"
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   25
                  Top             =   1440
                  Width           =   255
               End
            End
         End
      End
      Begin VB.CommandButton cmdGlobal 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Imprimer tous les avis"
         Height          =   1575
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ListBox lstSort 
         Enabled         =   0   'False
         Height          =   5520
         Left            =   5880
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   120
         Width           =   3375
      End
      Begin VB.Shape shapeList 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   9285
      End
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   0
      Width           =   2500
   End
   Begin VB.TextBox txtRecherche 
      Height          =   300
      Left            =   4440
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton cmdPageNext 
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
      Left            =   2200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdPagePrior 
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H80000004&
      Caption         =   "&Recherche"
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.Menu mnuOpTrf 
      Caption         =   "Recherche"
      Visible         =   0   'False
      Begin VB.Menu mnuOpTrfSaisir 
         Caption         =   "dossiers en cours de saisie"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOpTrfàValider 
         Caption         =   "dossiers à valider"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOpTrfàComptabiliser 
         Caption         =   "dossiers à comptabliser"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOpTrfSwiftàémettre 
         Caption         =   "messages SWIFT à émettre"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOpTrfVirementàémettre 
         Caption         =   "virement à émettre"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOpTrfAvisàImprimer 
         Caption         =   "Avis à imprimer"
      End
      Begin VB.Menu mnuOpTrfàRapprocher 
         Caption         =   "dossiers à rapprocher"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOpTrfX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpCptRéférence 
         Caption         =   "numéro de dossier"
      End
      Begin VB.Menu mnuSwift00 
         Caption         =   "dossiers (Donneur d'ordre)"
      End
      Begin VB.Menu mnuSwift20 
         Caption         =   "Référence du donneur d'ordre"
      End
      Begin VB.Menu mnuSwift01 
         Caption         =   "dossiers (banque du bénéficaire)"
      End
      Begin VB.Menu mnuOpTrfx2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatistiques 
         Caption         =   "Statistiques"
      End
   End
   Begin VB.Menu mnuDossier 
      Caption         =   "Dossier"
      Visible         =   0   'False
      Begin VB.Menu mnuDossierDétail 
         Caption         =   "détail"
      End
      Begin VB.Menu mnuDossierValider 
         Caption         =   "valider"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDossierAnnulerValidation 
         Caption         =   "annuler validation"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDossierComptabiliser 
         Caption         =   "comptabiliser"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDossierSwift 
         Caption         =   "émettre SWIFT"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDossierVirement 
         Caption         =   "émettre virement"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDossierAvis 
         Caption         =   "imprimer Avis"
      End
      Begin VB.Menu mnuDossierRapprocher 
         Caption         =   "rapprocher"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmOpTrf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim X As String, I As Integer
Dim Msg As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor

Dim mFieldNature As String

Private picListNb As Integer
Private picListNbMin As Integer
Private picListNbMax As Integer

Dim wOpTrfnb As Integer, wOpTrfnbMax As Integer, wOpTrfIndex As Integer, wOpTrfIndex0 As Integer
Dim wOpRéférence() As String
Dim wOpEch() As typeOpEch, wOpEch_Method As String
Dim wOpSwift() As typeOpSwift
Dim wOpCpt() As typeOpCpt
Dim sqlCase As String
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

'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
'cmdControl
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor

End Sub
'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
lstErr.Clear
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor

End Sub


Private Sub chkOpStatàCompta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkOpStatàCompta

End Sub


Private Sub chkOpStatàValider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkOpStatàValider

End Sub


Private Sub chkOpStatCompta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkOpStatCompta

End Sub


Private Sub chkOpStatDétail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkOpStatDétail

End Sub


Private Sub chkOpStatSort1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkOpStatSort1

End Sub


Private Sub chkOpStatSort2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkOpStatSort2

End Sub


Private Sub cmdContext_Click()
cmdReset
mnuOpTrfSaisir.Enabled = True ' XopAut.Saisir
mnuOpTrfàValider.Enabled = True ' XopAut.Valider
mnuOpTrfàComptabiliser.Enabled = True ' XopAut.Comptabiliser
mnuOpTrfSwiftàémettre.Enabled = True ' XopAut.Swift
mnuOpTrfVirementàémettre.Enabled = True ' XopAut.Virement
mnuOpTrfAvisàImprimer.Enabled = True ' XopAut.Avis
mnuOpTrfàRapprocher.Enabled = XopAut.Rapprocher
Me.PopupMenu mnuOpTrf, vbPopupMenuRightButton
End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub


Private Sub cmdGlobal_Click()
wOpEch_lstSort
End Sub

Private Sub cmdGlobal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdGlobal

End Sub


Private Sub cmdOpStatOk_Click()
Dim Msg As String
ReDim arrOpStat(50): arrOpStatNbMax = 50
arrOpStatNb = 0: arrOpStatIndex = 0
srvOpStat.Init recOpStat
lstErr.Clear

txtAmjMin_Control
txtAmjMax_Control
If paramOpStat.AmjMin > paramOpStat.AmjMax Then Call lstErr_AddItem(lstErr, txtAmjMax, "? date du > date au")

paramOpStat.Nb = 0
paramOpStat.HMSMin = "000000"
paramOpStat.HMSMax = "999999"

paramOpStat.selK1 = IIf(chkOpStatàValider = "1", constàValider, "")
paramOpStat.selK2 = IIf(chkOpStatàCompta = "1", constàCompta, "")
paramOpStat.selK3 = IIf(chkOpStatCompta = "1", constCompta, "")

If optSort1Devise Then paramOpStat.sortK1 = "Devise"
If optSort1Sender Then paramOpStat.sortK1 = "Sender"
If optSort1Correspondant Then paramOpStat.sortK1 = "Correspondant"
If optSort1Receiver Then paramOpStat.sortK1 = "Receiver"
If optSort1Référence Then paramOpStat.sortK1 = "Référence"

If optSort2Devise Then paramOpStat.sortK2 = "Devise"
If optSort2Sender Then paramOpStat.sortK2 = "Sender"
If optSort2Correspondant Then paramOpStat.sortK2 = "Correspondant"
If optSort2Receiver Then paramOpStat.sortK2 = "Receiver"
If optSort2Référence Then paramOpStat.sortK2 = "Référence"

paramOpStat.sortK3 = "Référence"

paramOpStat.prtK1 = IIf(chkOpStatSort1 = "1", True, False)
paramOpStat.prtK2 = IIf(chkOpStatSort2 = "1", True, False)
paramOpStat.prtK3 = False
paramOpStat.prtDétail = IIf(chkOpStatDétail = "1", True, False)

If paramOpStat.sortK1 = paramOpStat.sortK2 Then Call lstErr_AddItem(lstErr, txtAmjMax, "? tri 1 = tri 2")

If lstErr.ListCount <> 0 Then Exit Sub

MDB.Execute "delete * from OpStat"
srvOpStat.tableOpStat_Open

If Trim(paramOpStat.selK1) <> "" Then srvOpStat.Filtre "SnapL1", paramOpStat.selK1
If Trim(paramOpStat.selK2) <> "" Then srvOpStat.Filtre "SnapL1", paramOpStat.selK2
If Trim(paramOpStat.selK3) <> "" Then srvOpStat.Filtre "SnapL3", paramOpStat.selK3

Msg = Format$(1, "000000") & Format$(arrOpStatNb, "000000")

prtOpStatX Msg
srvOpStat.tableOpStat_Close

MDB.Execute "delete * from OpStat"
fraOpStat.Visible = False

End Sub

Private Sub cmdOpStatOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdOpStatOk

End Sub


Private Sub cmdOpStatQuit_Click()
fraOpStat.Visible = False
End Sub

Private Sub txtAmjMin_Control()
X = dateCtlDsys(txtAmjMin.Text)
If Not IsNumeric(X) Or mId$(X, 1, 8) = "00000000" Then
    Call lstErr_AddItem(lstErr, txtAmjMin, "? date : " & X)
Else
    paramOpStat.AmjMin = mId$(X, 1, 8)
    Call dateDsp(paramOpStat.AmjMin, txtAmjMin)
End If
End Sub


Private Sub txtAmjMax_Control()
X = dateCtlDsys(txtAmjMax.Text)
If Not IsNumeric(X) Or mId$(X, 1, 8) = "00000000" Then
    Call lstErr_AddItem(lstErr, txtAmjMax, "? date : " & X)
Else
    paramOpStat.AmjMax = mId$(X, 1, 8)
    Call dateDsp(paramOpStat.AmjMax, txtAmjMax)
End If
End Sub

Private Sub cmdOpStatQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdOpStatQuit

End Sub

Private Sub cmdPageNext_Click()
Select Case sqlCase
    Case "OpEch": wOpEch_Display
    
End Select
End Sub

Private Sub cmdPageNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPageNext

End Sub


Private Sub cmdPagePrior_Click()
Select Case sqlCase
    Case "OpEch": wOpTrfIndex = wOpTrfIndex0 - picListNb: wOpEch_Display
End Select
End Sub

Private Sub cmdPagePrior_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPagePrior

End Sub


'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------

X = Format$(1, "000000") & Format$(wOpTrfnb, "000000")

'prtBicX X

End Sub


'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
If fraOpStat.Visible Then
    fraOpStat.Visible = False
Else
    Unload Me
End If

End Sub




Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

Private Sub cmdSaisir_Click()
'-------------------------------------------------'
X = Space$(100)
Mid$(X, 1, 12) = "frmOpTrfD"
Mid$(X, 13, 12) = Me.Name
Mid$(X, 25, 10) = Space$(10)
Mid$(X, 35, 10) = constSaisie
Msg_Monitor X

End Sub

Private Sub cmdSaisir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdSaisir

End Sub


'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
If Not blnOpTrf_Update_on Then
    Select Case currentAction
        Case constàValider:     mnuOpTrfàValider_Click
        Case constSaisie:     mnuOpTrfSaisir_Click
        Case constàCompta:     mnuOpTrfàComptabiliser_Click
    End Select
End If
End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: cmdContext_Return
    Case Is = 27: cmdQuit_Click
'    Case Is = 34: cmdPageNext_Click
 '   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub


'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------

Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
Call BiaPgmAut_Init("OPTRF", XopAut)
ReDim wOpEch(1)
ReDim wOpCpt(1)
ReDim wOpSwift(1): wOpTrfnbMax = 0
ReDim arrOpEch(1): ReDim XopEch(1): arrOpEchNbMax = 0
ReDim arrOpSwift(1): ReDim XopSwift(1): arrOpSwiftNbMax = 0
ReDim arrOpCpt(1): ReDim XopCpt(1): arrOpCptNbMax = 0

recBicInit XopBicSender: XopBicReceiver = XopBicSender
recBicIbanInit XopBicIbanSender: XopBicIbanReceiver = XopBicIbanSender
Call cmdClear
blnOpTrf_Update_on = False
End Sub



'---------------------------------------------------------
Public Sub cmdClear()
'---------------------------------------------------------

txtRecherche = ""
cmdReset

End Sub




Public Sub wOpEch_Sql()
sqlCase = "OpEch"
recOpEch_Init wOpEch(0)
wOpEch(0).Method = wOpEch_Method
wOpEch(0).EchéanceNature = currentAction
wOpEch(0).EchéanceAmj = "00000000"
wOpEch(0).EchéanceHms = "000000"
arrOpEch(0) = wOpEch(0)
arrOpEch(0).EchéanceNature = Trim(currentAction) & "99999999999"
arrOpEch(0).Référence = "99999999999"
arrOpEch(0).EchéanceAmj = "99999999"
arrOpEch(0).EchéanceHms = "999999"
arrOpEchNb = 0: arrOpEchIndex = 0
arrOpEchSuite = True
Do Until Not arrOpEchSuite
    srvOpEch_Monitor wOpEch(0)
    wOpEch(0) = arrOpEch(arrOpEchNb)
    wOpEch(0).Method = wOpEch_Method & "+"
Loop
wOpEch_Load
End Sub

Private Sub wOpEch_Load()
If arrOpEchNb > 0 Then ReDim wOpEch(arrOpEchNb)
ReDim wOpRéférence(arrOpEchNb)
For I = 1 To arrOpEchNb
    wOpEch(I) = arrOpEch(I)
    wOpRéférence(I) = arrOpEch(I).Référence
Next I
wOpTrfnb = arrOpEchNb
wOpTrfnbMax = arrOpEchNb
wOpTrfIndex = 0
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub

Private Sub fraOpStat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraOpStatFiltre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraOpStatImpression_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraOpStatPériode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraOpStatSélection_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraOpStatSort1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraOpStatSort2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub mnuDossierAnnulerValidation_Click()
srvOptrf.AnnulerValidation wOpRéférence(wOpTrfIndex)
mnuOpTrfàComptabiliser_Click
End Sub

Private Sub mnuDossierAvis_Click()
Select Case currentAction
    Case constàAvis: mnuDossierAvis_wOpEch
    Case "wOpCpt": mnuDossierAvis_wOpCpt
End Select
End Sub

Private Sub mnuDossierComptabiliser_Click()
'Msg_Snd constCompta
End Sub

Private Sub mnuDossierDétail_Click()
Msg_Snd ""
End Sub

Private Sub mnuDossierValider_Click()
Msg_Snd constValider
End Sub

Private Sub mnuOpCptRéférence_Click()
currentAction = "wOpCpt"
mnuDossier_Disabled
mnuDossierAvis.Enabled = XopAut.Avis
wOpCpt_Sql
wOpCpt_Display

End Sub

Private Sub mnuOpTrfàComptabiliser_Click()
currentAction = constàCompta
mnuDossier_Disabled
mnuDossierComptabiliser.Enabled = False 'XopAut.Comptabiliser
mnuDossierAnnulerValidation.Enabled = XopAut.Valider
wOpEch_Method = "SnapL2"
wOpEch_Sql
wOpEch_Display
End Sub

Private Sub mnuOpTrfàRapprocher_Click()
currentAction = constàRapprocher
mnuDossier_Disabled
mnuDossierRapprocher.Enabled = False 'XopAut.Rapprocher
wOpEch_Method = "SnapL2"
wOpEch_Sql
wOpEch_Display
End Sub


Private Sub mnuOpTrfàValider_Click()
currentAction = constàValider
mnuDossier_Disabled
mnuDossierValider.Enabled = XopAut.Valider
wOpEch_Method = "SnapL2"
wOpEch_Sql
wOpEch_Display
End Sub

Private Sub mnuOpTrfAvisàImprimer_Click()
currentAction = constàAvis
mnuDossier_Disabled
mnuDossierAvis.Enabled = XopAut.Avis
wOpEch_Method = "SnapL1"
wOpEch_Sql
wOpEch_Display
If wOpTrfnb > 0 Then cmdGlobal.Visible = True: cmdGlobal.Caption = "Imprimer les avis" & Chr(13) & "(échéance max : " & dateImp(DSys) & " )"

End Sub


Private Sub mnuOpTrfRéférence_Click()
End Sub

Private Sub mnuOpTrfSaisir_Click()
currentAction = constSaisie
mnuDossier_Disabled
wOpEch_Method = "SnapL2"
wOpEch_Sql
wOpEch_Display
End Sub


Private Sub mnuStatistiques_Click()
lblOpStatNb = ""
lblOpStatNb.ForeColor = warnUsrColor
fraOpStat.Visible = True
optOpStatJour = True
optOpStatJour_Click
End Sub

Private Sub mnuSwift00_Click()
mFieldNature = "00 "
mnuDossier_Disabled
wOpSwift_Sql
wOpSwift_Display

End Sub

Private Sub mnuSwift01_Click()
mFieldNature = "01 "
mnuDossier_Disabled
wOpSwift_Sql
wOpSwift_Display

End Sub


Private Sub mnuSwift20_Click()
mFieldNature = "20 "
mnuDossier_Disabled
wOpSwift_Sql
wOpSwift_Display

End Sub


Private Sub optOpStatJour_Click()
txtAmjMin.Enabled = False: Call dateDsp(DSys, txtAmjMin)
txtAmjMax.Enabled = False: txtAmjMax = txtAmjMin
End Sub

Private Sub optOpStatJour_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optOpStatJour

End Sub


Private Sub optOpStatManuel_Click()
txtAmjMax.Enabled = True
txtAmjMin.Enabled = True
txtAmjMin.SetFocus
End Sub

Private Sub optOpStatManuel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optOpStatManuel

End Sub


Private Sub optOpStatMois_Click()
txtAmjMax.Enabled = False: Call dateDsp(DSys, txtAmjMax)
txtAmjMin.Enabled = False: Call dateDsp(mId$(DSys, 1, 6) & "01", txtAmjMin)

End Sub

Private Sub optOpStatMois_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optOpStatMois

End Sub


Private Sub optOpStatVeille_Click()
Dim X As String
X = dateElp("Ouvré", -1, DSys)
txtAmjMin.Enabled = False: Call dateDsp(X, txtAmjMin)
txtAmjMax.Enabled = False: txtAmjMax = txtAmjMin

End Sub


Private Sub optOpStatVeille_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optOpStatVeille

End Sub


Private Sub optSort1Correspondant_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort1Correspondant

End Sub


Private Sub optSort1Devise_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort1Devise

End Sub


Private Sub optSort1Receiver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort1Receiver

End Sub


Private Sub optSort1Référence_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort1Référence

End Sub


Private Sub optSort1Sender_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort1Sender

End Sub


Private Sub optSort2Correspondant_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort2Correspondant

End Sub


Private Sub optSort2Devise_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort2Devise

End Sub


Private Sub optSort2Receiver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort2Receiver

End Sub


Private Sub optSort2Référence_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort2Référence

End Sub


Private Sub optSort2Sender_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSort2Sender

End Sub


Private Sub picList_Click()
If wOpTrfIndex >= 0 And wOpTrfIndex <= wOpTrfnbMax Then
    shapeList.Top = (wOpTrfIndex - picListNbMin) * lstLineHeight '- 20
    shapeList.Height = lstLineHeight
    shapeList.Visible = True
    Me.PopupMenu mnuDossier, vbPopupMenuRightButton
End If
End Sub

Private Sub picList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------
'wOpTrfIndex = picListNbMin + (y - 20) \ lstLineHeight
wOpTrfIndex = picListNbMin + Y \ lstLineHeight
End Sub








Private Sub txtAmjMax_GotFocus()
Call txt_GotFocus(txtAmjMax)

End Sub


Private Sub txtAmjMax_LostFocus()
Call txt_LostFocus(txtAmjMin)
txtAmjMin_Control

End Sub


Private Sub txtAmjMin_GotFocus()
Call txt_GotFocus(txtAmjMin)
End Sub


Private Sub txtAmjMin_LostFocus()
Call txt_LostFocus(txtAmjMin)
txtAmjMin_Control
End Sub


'-------------------------------------------------'
Private Sub txtRecherche_GotFocus()
'-------------------------------------------------'
Call txt_GotFocus(txtRecherche)
End Sub

'---------------------------------------------------------
Private Sub txtRecherche_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------
KeyAscii = convUCase(KeyAscii)
End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
picList.Cls
cmdPagePrior.Enabled = False
cmdPageNext.Enabled = False
cmdSaisir.Enabled = XopAut.Saisir
cmdGlobal.Visible = False
lstSort.Clear: lstSort.Visible = False

arrTag_Set False
lstErr.Visible = False
lstErrClear = True
Call lstErr_Clear(lstErr, txtRecherche, "sélection")
usrColor_Set
sqlCase = ""
End Sub

'---------------------------------------------------------
Public Sub cmdValidation()
'---------------------------------------------------------

lstErrClear = False
End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub

'---------------------------------------------------------
Public Sub cmdList()
'---------------------------------------------------------
picList.Cls
End Sub

'---------------------------------------------------------
Private Sub wOpEch_Display()
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K3 As Integer

picList.Cls
If wOpTrfIndex < 0 Then wOpTrfIndex = 0

If wOpTrfnb <= 0 Then
    X = "Néant ( " & Now & " )"
    Call lstErr_Clear(lstErr, txtRecherche, X)
Else
    picListNbMin = wOpTrfIndex + 1
End If
    
K1 = 0
picListNb = picList.Height \ lstLineHeight - 1
K3 = 1
picList.CurrentY = -lstLineHeight
wOpTrfIndex0 = wOpTrfIndex

For K = K1 To picListNb Step K3

    If wOpTrfIndex >= wOpTrfnb Then
        Exit For
    Else
        wOpTrfIndex = wOpTrfIndex + 1
        picList.CurrentY = picList.CurrentY + lstLineHeight
        Call recOpEch_Display(picList, wOpEch(wOpTrfIndex))
    End If
Next K

picListNbMax = wOpTrfIndex
cmdPagePrior.Enabled = IIf(picListNbMin > 1, True, False)
cmdPageNext.Enabled = IIf(picListNbMax < wOpTrfnb, True, False)

End Sub
Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------

End Sub

Public Sub Msg_Snd(ByVal X As String)
Msg = Space$(100)

Mid$(Msg, 1, 12) = "frmOpTrfD   "
Mid$(Msg, 13, 12) = Me.Name
Mid$(Msg, 25, 10) = Space$(10)
Mid$(Msg, 35, 10) = X
Mid$(Msg, 45, 6) = wOpRéférence(wOpTrfIndex)
Msg_Monitor Msg

End Sub

Public Sub mnuDossier_Disabled()
mnuDossierComptabiliser.Enabled = False
mnuDossierRapprocher.Enabled = False
mnuDossierSwift.Enabled = False
mnuDossierValider.Enabled = False
mnuDossierAnnulerValidation.Enabled = False
mnuDossierVirement.Enabled = False
mnuDossierAvis.Enabled = False
End Sub

Public Sub wOpEch_lstSort()
Dim K As Integer
lstSort.Clear
lstSort.Visible = True
cmdGlobal.Visible = False
ReDim arrOpCpt(wOpTrfnb): arrOpCptNbMax = wOpTrfnb
recOpCpt_Init arrOpCpt(0)
arrOpCpt(0).Method = "SeekP0"
For wOpTrfIndex = 1 To wOpTrfnb
    If wOpEch(wOpTrfIndex).EchéanceAmj <= DSys Then
        arrOpCpt(wOpTrfIndex) = arrOpCpt(0)
        arrOpCpt(wOpTrfIndex).Référence = wOpEch(wOpTrfIndex).Référence
        arrOpCpt(wOpTrfIndex).Séquence = IIf(Trim(wOpEch(wOpTrfIndex).EchéanceNature) = constàAvisCr, 2, 1)
        If IsNull(srvOpCpt_Monitor(arrOpCpt(wOpTrfIndex))) Then
            lstSort.AddItem arrOpCpt(wOpTrfIndex).Compte & "." _
                        & arrOpCpt(wOpTrfIndex).Devise & "." _
                        & arrOpCpt(wOpTrfIndex).Référence & "." _
                        & Format$(wOpTrfIndex, "00000")
        End If
    End If
Next wOpTrfIndex
If lstSort.ListCount <= 0 Then lstSort.Visible = False: Exit Sub

prtAvis_Global lstSort

X = MsgBox("Voulez-vous mettre à jour l'échéancier ?", vbYesNo + vbQuestion + vbDefaultButton2, "Impression des avis")
If X = vbYes Then
    For K = 0 To lstSort.ListCount - 1
        lstSort.ListIndex = K
        wOpTrfIndex = Val(mId$(lstSort.Text, 24, 5))
        srvOpTrf_dbUpdateOpEch_ActionOk wOpEch(wOpTrfIndex)
    Next K
End If
cmdClear
End Sub

Public Sub wOpSwift_Sql()
sqlCase = "OpSwift"
recOpSwift_Init wOpSwift(0)
wOpSwift(0).Method = "SnapL1"
wOpSwift(0).FieldNature = mFieldNature
wOpSwift(0).FieldText = Trim(txtRecherche)
arrOpSwift(0) = wOpSwift(0)
arrOpSwift(0).FieldText = Trim(txtRecherche) & "99999999999"
arrOpSwiftNb = 0: arrOpSwiftIndex = 0
srvOpSwift_Monitor wOpSwift(0)
wOpSwift_Load

End Sub

Public Sub wOpSwift_Display()
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer

picList.Cls
If wOpTrfIndex < 0 Then wOpTrfIndex = 0

If wOpTrfnb <= 0 Then
    X = "Néant ( " & Now & " )"
    Call lstErr_Clear(lstErr, txtRecherche, X)
Else
    picListNbMin = wOpTrfIndex + 1
End If
    
K1 = 0
K2 = picList.Height \ lstLineHeight
K3 = 1
picList.CurrentY = -lstLineHeight
For K = K1 To K2 Step K3

    If wOpTrfIndex >= wOpTrfnb Then
        Exit For
    Else
        wOpTrfIndex = wOpTrfIndex + 1
        picList.CurrentY = picList.CurrentY + lstLineHeight
        Call recOpSwift_Display(picList, wOpSwift(wOpTrfIndex))
    End If
Next K

picListNbMax = wOpTrfIndex
cmdPagePrior.Enabled = IIf(picListNbMin > 1, True, False)
cmdPageNext.Enabled = IIf(picListNbMax < wOpTrfnb, True, False)


End Sub

Public Sub wOpSwift_Load()
If arrOpSwiftNb > 0 Then ReDim wOpSwift(arrOpSwiftNb)
ReDim wOpRéférence(arrOpSwiftNb)
For I = 1 To arrOpSwiftNb
    wOpSwift(I) = arrOpSwift(I)
    wOpRéférence(I) = arrOpSwift(I).Référence
Next I
wOpTrfnb = arrOpSwiftNb
wOpTrfnbMax = arrOpSwiftNb
wOpTrfIndex = 0

End Sub

Public Sub wOpCpt_Display()
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer

picList.Cls
If wOpTrfIndex < 0 Then wOpTrfIndex = 0

If wOpTrfnb <= 0 Then
    X = "Néant ( " & Now & " )"
    Call lstErr_Clear(lstErr, txtRecherche, X)
Else
    picListNbMin = wOpTrfIndex + 1
End If
    
K1 = 0
K2 = picList.Height \ lstLineHeight
K3 = 1
picList.CurrentY = -lstLineHeight
For K = K1 To K2 Step K3

    If wOpTrfIndex >= wOpTrfnb Then
        Exit For
    Else
        wOpTrfIndex = wOpTrfIndex + 1
        picList.CurrentY = picList.CurrentY + lstLineHeight
        Call recOpCpt_Display(picList, wOpCpt(wOpTrfIndex))
    End If
Next K

picListNbMax = wOpTrfIndex
cmdPagePrior.Enabled = IIf(picListNbMin > 1, True, False)
cmdPageNext.Enabled = arrOpCptSuite
End Sub

Public Sub wOpCpt_Load()
If arrOpCptNb > 0 Then ReDim wOpCpt(arrOpCptNb)
ReDim wOpRéférence(arrOpCptNb)
For I = 1 To arrOpCptNb
    wOpCpt(I) = arrOpCpt(I)
    wOpRéférence(I) = arrOpCpt(I).Référence
Next I
wOpTrfnb = arrOpCptNb
wOpTrfnbMax = arrOpCptNb
wOpTrfIndex = 0

End Sub

Public Sub wOpCpt_Sql()
sqlCase = "OpCpt"
recOpCpt_Init wOpCpt(0)
wOpCpt(0).Method = "SnapP0"
wOpCpt(0).Référence = Trim(txtRecherche)
arrOpCpt(0) = wOpCpt(0)
arrOpCpt(0).Référence = "99999999999"
arrOpCptNb = 0: arrOpCptIndex = 0
srvOpCpt_Monitor wOpCpt(0)
wOpCpt_Load

End Sub

Public Sub mnuDossierAvis_wOpEch()
srvOpTrf_dbReadEch wOpEch(wOpTrfIndex).Référence: XopEch_Load
srvOpTrf_dbReadCpt wOpEch(wOpTrfIndex).Référence: XopCpt_Load
If Trim(wOpEch(wOpTrfIndex).EchéanceNature) = constàAvisDb Then
    Call prtAvisX(XopCpt(1), " ")
Else
    Call prtAvisX(XopCpt(2), " ")
End If
srvOpTrf_dbUpdateOpEch_ActionOk wOpEch(wOpTrfIndex)

End Sub

Public Sub mnuDossierAvis_wOpCpt()
Dim X As String
'If wOpCpt(wOpTrfIndex).optAvis <> "0" Then
    Call prtAvisX(wOpCpt(wOpTrfIndex), " ")
'Else
'    X = MsgBox("NEANT", vbCritical, "Impression Avis")
'End If
End Sub


Public Sub cmdContext_Return()
If fraOpStat.Visible Then
    cmdOpStatOk_Click
Else
    cmdContext_Click
End If

End Sub
