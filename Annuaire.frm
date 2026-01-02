VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmAnnuaire 
   Caption         =   "Annuaire : Détail"
   ClientHeight    =   5670
   ClientLeft      =   3120
   ClientTop       =   1005
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   6225
   Begin VB.Frame fraAnnuaire 
      Height          =   4935
      Left            =   0
      TabIndex        =   15
      Top             =   360
      Width           =   6135
      Begin VB.Frame Frame3 
         Caption         =   "Micro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   975
         Left            =   120
         TabIndex        =   22
         Top             =   3840
         Width           =   5895
         Begin VB.TextBox txtMicroSN 
            Height          =   285
            Left            =   1800
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtMicroIP 
            Height          =   285
            Left            =   4320
            TabIndex        =   9
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblMicroSN 
            Caption         =   "Identification Réseau"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblMicroIP 
            Caption         =   "Adresse I P"
            Height          =   255
            Left            =   3120
            TabIndex        =   23
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Téléphone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1095
         Left            =   120
         TabIndex        =   19
         Top             =   2640
         Width           =   5895
         Begin VB.TextBox txtTél1 
            Height          =   285
            Left            =   1800
            TabIndex        =   5
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtTél3 
            Height          =   285
            Left            =   4320
            TabIndex        =   7
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtTél2 
            Height          =   285
            Left            =   3480
            TabIndex        =   6
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblTél1 
            Caption         =   "Poste Principal"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblTél2 
            Caption         =   "Autres Postes"
            Height          =   255
            Left            =   2400
            TabIndex        =   20
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Intitulé"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   2535
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   5895
         Begin VB.Frame Frame4 
            Height          =   1695
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1575
            Begin VB.OptionButton optCivilitéAutres 
               Caption         =   "Autre"
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   1320
               Width           =   975
            End
            Begin VB.OptionButton optCivilitéMademoiselle 
               Caption         =   "Mademoiselle"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   960
               Width           =   1335
            End
            Begin VB.OptionButton optCivilitéMadame 
               Caption         =   "Madame"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   600
               Width           =   975
            End
            Begin VB.OptionButton optCivilitéMonsieur 
               Caption         =   "Monsieur"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   240
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.TextBox txtBureau 
            Height          =   285
            Left            =   2520
            TabIndex        =   4
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox txtService 
            Height          =   285
            Left            =   2520
            TabIndex        =   3
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox txtNom 
            Height          =   285
            Left            =   2520
            TabIndex        =   1
            Top             =   360
            Width           =   3255
         End
         Begin VB.TextBox txtPrénoms 
            Height          =   285
            Left            =   2520
            TabIndex        =   2
            Top             =   720
            Width           =   3255
         End
         Begin VB.Label lblBureau 
            Caption         =   "Bureau"
            Height          =   255
            Left            =   1800
            TabIndex        =   27
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label lblService 
            Caption         =   "Service"
            Height          =   255
            Left            =   1800
            TabIndex        =   26
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lblNom 
            Caption         =   "Nom"
            Height          =   255
            Left            =   1800
            TabIndex        =   18
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lblPrénoms 
            Caption         =   "Prénoms"
            Height          =   255
            Left            =   1800
            TabIndex        =   17
            Top             =   840
            Width           =   735
         End
      End
   End
   Begin VB.TextBox txtId 
      Height          =   285
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdSuppress 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Supprimer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   5760
      Picture         =   "Annuaire.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   0
      Width           =   400
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Enregistrer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   1200
   End
   Begin ComCtl2.UpDown UpDownAnnuaire 
      Height          =   345
      Index           =   18
      Left            =   5160
      TabIndex        =   25
      Top             =   0
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   609
      _Version        =   327681
      Value           =   1
      AutoBuddy       =   -1  'True
      OrigLeft        =   3000
      OrigTop         =   3120
      OrigRight       =   3240
      OrigBottom      =   3405
      Max             =   9999
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label lblNuméro 
      Caption         =   "Code "
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   120
      Width           =   495
   End
   Begin VB.Menu cmdPrint_mnu 
      Caption         =   "Print"
      Visible         =   0   'False
      Begin VB.Menu cmdPrint_mnuList 
         Caption         =   "Liste téléphonique"
      End
      Begin VB.Menu cmdPrint_mnuDétail 
         Caption         =   "Liste détaillée"
      End
   End
End
Attribute VB_Name = "frmAnnuaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim Msg As String
Dim AnnuaireAut As typeAuthorization
Dim currentMethod As String, currentAMJ As String

Dim updAnnuaire As Boolean

Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: cmdRechercher
    Case Is = constcmdEnregistrer: cmdEnregistrer
End Select
End Sub

Private Sub cmdEnregistrer()

'--------------------------------------------------------------
If optCivilitéMonsieur Then recAnnuaire.Civilité = "1"
If optCivilitéMadame Then recAnnuaire.Civilité = "2"
If optCivilitéMademoiselle Then recAnnuaire.Civilité = "3"
If optCivilitéAutres Then recAnnuaire.Civilité = "4"

recAnnuaire.Nom = txtNom
recAnnuaire.Prénoms = txtPrénoms

recAnnuaire.Tél1 = txtTél1
recAnnuaire.Tél2 = txtTél2
recAnnuaire.Tél3 = txtTél3
recAnnuaire.MicroSN = txtMicroSN
recAnnuaire.MicroIP = txtMicroIP
recAnnuaire.Service = txtService
recAnnuaire.Bureau = txtBureau

dbAnnuaire_Update recAnnuaire
cmdClear

End Sub
Private Sub cmdPrint_Click()
Me.PopupMenu cmdPrint_mnu

End Sub
Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim X As String

fraAnnuaire.Enabled = False
X = Trim(mId$(Msg, 13, Len(Msg)))
Call BiaPgmAut_Init("Annuaire", AnnuaireAut)
If AnnuaireAut.Saisir Then MDB_Master

txtId.Visible = AnnuaireAut.Saisir

If X = "" Then
    tableAnnuaire_Open
    updAnnuaire = False
    cmdClear
Else
    updAnnuaire = True
    txtId.Enabled = False
    cmdContext.Visible = False
    cmdSuppress.Visible = False
    cmdPrint_mnuDétail.Visible = False
    arrAnnuaire_Scan X
    Rec_Display
End If

End Sub

Public Sub cmdContext_Quit()
If fraAnnuaire.Enabled Then
    cmdClear
Else
    Unload Me
End If

End Sub

Public Sub cmdContext_Return()
If fraAnnuaire.Enabled Then
    If ActiveControl.Name = lastActiveControl_Name Then
 '       cmdOk_Click
    Else
        SendKeys "{TAB}"
    End If
Else
If cmdContext.Caption = constcmdRechercher Then cmdRechercher
End If

End Sub


Public Sub Rec_Display()
'----------------------------------------------------
Select Case recAnnuaire.Civilité
    Case Is = "1": optCivilitéMonsieur = True
    Case Is = "2": optCivilitéMadame = True
    Case Is = "3": optCivilitéMademoiselle = True
    Case Is = "4": optCivilitéAutres = True
     
End Select
txtId = Trim(recAnnuaire.Id)
txtNom = Trim(recAnnuaire.Nom)
txtPrénoms = Trim(recAnnuaire.Prénoms)
txtTél1 = Trim(recAnnuaire.Tél1)
txtTél2 = Trim(recAnnuaire.Tél2)
txtTél3 = Trim(recAnnuaire.Tél3)
txtMicroSN = Trim(recAnnuaire.MicroSN)
txtMicroIP = Trim(recAnnuaire.MicroIP)
txtService = Trim(recAnnuaire.Service)
txtBureau = Trim(recAnnuaire.Bureau)

End Sub

Private Sub cmdPrint_mnuDétail_Click()
Dim Msg As String
Msg = "000001" & Format$(arrAnnuaireNb, "000000") & "D"

prtAnnuaireX Msg

End Sub

Private Sub cmdPrint_mnuList_Click()
Dim Msg As String
Msg = "000001" & Format$(arrAnnuaireNb, "000000") & "L"

prtAnnuaireX Msg

End Sub

Private Sub cmdSuppress_Click()
recAnnuaire.Method = "Delete"
dbAnnuaire_Update recAnnuaire
cmdClear

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub


Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
End Sub


'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
End Sub
'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
tableAnnuaire_Close
frmElp.lstAnnuaire.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDB_Local

End Sub

Private Sub txtBureau_GotFocus()
Call txt_GotFocus(txtBureau)

End Sub


Private Sub txtBureau_LostFocus()
Call txt_LostFocus(txtBureau)

End Sub


Private Sub txtMicroIP_GotFocus()
Call txt_GotFocus(txtMicroIP)

End Sub


Private Sub txtMicroIP_LostFocus()
Call txt_LostFocus(txtMicroIP)

End Sub


Private Sub txtMicroSN_GotFocus()
Call txt_GotFocus(txtMicroSN)

End Sub


Private Sub txtMicroSN_LostFocus()
Call txt_LostFocus(txtMicroSN)

End Sub


Private Sub txtNom_GotFocus()
Call txt_GotFocus(txtNom)

End Sub


Private Sub txtNom_LostFocus()
Call txt_LostFocus(txtNom)

End Sub


Private Sub txtID_GotFocus()
Call txt_GotFocus(txtId)

End Sub


Private Sub txtId_LostFocus()
Call txt_LostFocus(txtId)

End Sub


Private Sub txtPrénoms_GotFocus()
Call txt_GotFocus(txtPrénoms)

End Sub


Private Sub txtPrénoms_LostFocus()
Call txt_LostFocus(txtPrénoms)

End Sub


Private Sub txtService_GotFocus()
Call txt_GotFocus(txtService)

End Sub


Private Sub txtService_LostFocus()
Call txt_LostFocus(txtService)

End Sub


Private Sub txtTél1_GotFocus()
Call txt_GotFocus(txtTél1)

End Sub


Private Sub txtTél1_LostFocus()
Call txt_LostFocus(txtTél1)

End Sub


Private Sub txtTél2_GotFocus()
Call txt_GotFocus(txtTél2)

End Sub


Private Sub txtTél2_LostFocus()
Call txt_LostFocus(txtTél2)

End Sub


Private Sub txtTél3_GotFocus()
Call txt_GotFocus(txtTél3)

End Sub


Private Sub txtTél3_LostFocus()
Call txt_LostFocus(txtTél3)

End Sub



Public Sub cmdClear()
txtId = ""
cmdContext.Visible = AnnuaireAut.Saisir
cmdContext.Caption = constcmdRechercher
cmdSuppress.Visible = False
cmdPrint_mnuDétail.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "préciser le code ")
fraAnnuaire.Enabled = False
txtId.Enabled = True: txtId.SetFocus

End Sub

Public Sub cmdRechercher()
If Trim(txtId) = "" Then
    Call lstErr_Clear(lstErr, cmdContext, "préciser le code ")
    Exit Sub
End If

recAnnuaire_Init recAnnuaire
recAnnuaire.Id = Format$(RTrim(txtId), "@@@@")
recAnnuaire.Method = "Seek="
dbAnnuaire_Read recAnnuaire

If recAnnuaire.Err = 0 Then
    recAnnuaire.Method = "Update"
    cmdContext.Caption = constcmdEnregistrer
    cmdSuppress.Visible = True
Else
    recAnnuaire.Method = "AddNew"
    cmdContext.Caption = constcmdEnregistrer
    cmdSuppress.Visible = False
End If
txtId.Enabled = False
fraAnnuaire.Enabled = True
Rec_Display
txtNom.SetFocus
End Sub

Private Sub UpDownAnnuaire_DownClick(Index As Integer)
Dim X As String
If updAnnuaire Then
    If frmElp.lstAnnuaire.ListIndex < frmElp.lstAnnuaire.ListCount - 1 Then
        frmElp.lstAnnuaire.ListIndex = frmElp.lstAnnuaire.ListIndex + 1
        X = frmElp.lstAnnuaire.Text
        arrAnnuaire_Scan X
        Rec_Display
    End If
End If
End Sub

Private Sub UpDownAnnuaire_UpClick(Index As Integer)
Dim X As String
If updAnnuaire Then
    If frmElp.lstAnnuaire.ListIndex > 0 Then
        frmElp.lstAnnuaire.ListIndex = frmElp.lstAnnuaire.ListIndex - 1
        X = frmElp.lstAnnuaire.Text
        arrAnnuaire_Scan X
        Rec_Display
    End If
End If
End Sub


