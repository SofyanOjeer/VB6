VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmLrTiers 
   Caption         =   "LucaRisques : Déail Tiers"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   8970
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000004&
      Caption         =   "&Abandonner"
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   0
      Width           =   1200
   End
   Begin VB.Frame fraLrTiers 
      Height          =   6135
      Left            =   0
      TabIndex        =   39
      Top             =   480
      Width           =   8895
      Begin VB.Frame Frame3 
         Caption         =   "Personne Morale"
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
         Height          =   855
         Left            =   120
         TabIndex        =   44
         Top             =   3360
         Width           =   8655
         Begin VB.TextBox txtCTJURI 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   4200
            TabIndex        =   11
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtCDACCO 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   1920
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblCTJURI 
            Caption         =   "Code Juridique"
            Height          =   255
            Left            =   3000
            TabIndex        =   46
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblCDACCO 
            Caption         =   "Code APE"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Adresse Fiscale"
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
         Height          =   1815
         Left            =   120
         TabIndex        =   43
         Top             =   4200
         Width           =   8655
         Begin VB.TextBox txtCDPAYS2 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   3120
            TabIndex        =   13
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtCDDEPT2 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   5040
            TabIndex        =   14
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtLBCOMM2 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   1920
            TabIndex        =   18
            Top             =   1320
            Width           =   6615
         End
         Begin VB.TextBox txtCDPOST 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   1920
            TabIndex        =   16
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtNOVOIE 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   1920
            TabIndex        =   15
            Top             =   600
            Width           =   6615
         End
         Begin VB.TextBox txtCDRESI 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   1920
            TabIndex        =   12
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblCDPAYS2 
            Caption         =   "Pays"
            Height          =   255
            Left            =   2520
            TabIndex        =   33
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblCDDEPT2 
            Caption         =   "Département"
            Height          =   255
            Left            =   3840
            TabIndex        =   32
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblLBCOMM2 
            Caption         =   "Commune"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblCDPOST 
            Caption         =   "Code Postal"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label lblNOVOIE 
            Caption         =   "Voie"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblCDRESI 
            Caption         =   "Résident"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Personne Physique"
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
         Height          =   1815
         Left            =   120
         TabIndex        =   42
         Top             =   1560
         Width           =   8655
         Begin MSComCtl2.DTPicker txtAmjDATNAIS 
            Height          =   255
            Left            =   2640
            TabIndex        =   47
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   0   'False
            DateIsNull      =   -1  'True
            Format          =   28180481
            CurrentDate     =   36353
         End
         Begin VB.TextBox txtNOMCJT 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   1920
            TabIndex        =   9
            Top             =   1320
            Width           =   6615
         End
         Begin VB.TextBox txtLBCOMM1 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   1920
            TabIndex        =   8
            Top             =   960
            Width           =   6615
         End
         Begin VB.TextBox txtCDCOMM1 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   6120
            TabIndex        =   7
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtCDDEPT1 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   4320
            TabIndex        =   6
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtCDPAYS1 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   1920
            TabIndex        =   5
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtCDSEXE 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   6120
            TabIndex        =   4
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblCDSEXE 
            Caption         =   "Sexe"
            Height          =   255
            Left            =   5040
            TabIndex        =   22
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblNOMCJT 
            Caption         =   "Nom Conjoint"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label lblLBCOMM1 
            Caption         =   "Libellé"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblCDCOMM1 
            Caption         =   "Commune"
            Height          =   255
            Left            =   5040
            TabIndex        =   25
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblCDDEPT1 
            Caption         =   "Département"
            Height          =   255
            Left            =   3120
            TabIndex        =   24
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblCDPAYS1 
            Caption         =   "Pays"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblAmjDATNAIS 
            Caption         =   "Date Naissance     (à faire)"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Identité"
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
         Height          =   1455
         Left            =   120
         TabIndex        =   40
         Top             =   120
         Width           =   8655
         Begin VB.TextBox txtNSIREN 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   1920
            TabIndex        =   1
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtNomBNF 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   1920
            TabIndex        =   2
            Top             =   600
            Width           =   6615
         End
         Begin VB.TextBox txtPrénom 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   1920
            TabIndex        =   3
            Top             =   960
            Width           =   6615
         End
         Begin VB.Label lblNSIREN 
            Caption         =   "SIREN"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblNomBNF 
            Caption         =   "Nom"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   615
         End
         Begin VB.Label lblPrénom 
            Caption         =   "Prénom"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   735
         End
      End
   End
   Begin VB.TextBox txtRFBENF 
      Height          =   285
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   1335
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
      Height          =   500
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   6120
      TabIndex        =   36
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8400
      Picture         =   "LucaRisquesTiers.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   0
      Width           =   500
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
      Height          =   500
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   0
      Width           =   1200
   End
   Begin ComCtl2.UpDown UpDownLrTiers 
      Height          =   495
      Index           =   18
      Left            =   5760
      TabIndex        =   41
      Top             =   0
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   873
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
   Begin VB.Label lblRFBENF 
      Caption         =   "Code "
      Height          =   255
      Left            =   3720
      TabIndex        =   38
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
Attribute VB_Name = "frmLrTiers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim Msg As String
Dim LrTiersAut As typeAuthorization
Dim currentMethod As String, currentAMJ As String

Dim updLrTiers As Boolean

Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: cmdRechercher
    Case Is = constcmdEnregistrer: cmdEnregistrer
End Select
End Sub

Private Sub cmdEnregistrer()

'--------------------------------------------------------------

recLrTiers.NOMBNF = Trim(txtNomBNF)
recLrTiers.PRENOM = Trim(txtPrénom)
recLrTiers.NSIREN = Trim(txtNSIREN)
recLrTiers.RFBENF = Trim(txtRFBENF)
recLrTiers.CDSEXE = Trim(txtCDSEXE)
'recLrTiers.AmjDATNAIS = txtAmjDATNAIS ''''!!!!!!!!!!! si personne physique
recLrTiers.CDPAYS1 = Trim(txtCDPAYS1)
recLrTiers.CDDEPT1 = Trim(txtCDDEPT1)
recLrTiers.CDCOMM1 = Trim(txtCDCOMM1)
recLrTiers.LBCOMM1 = Trim(txtLBCOMM1)
recLrTiers.NOMCJT = Trim(txtNOMCJT)
recLrTiers.CDACCO = Trim(txtCDACCO)
recLrTiers.CTJURI = Trim(txtCTJURI)
recLrTiers.CDRESI = Trim(txtCDRESI)
recLrTiers.NOVOIE = Trim(txtNOVOIE)
recLrTiers.CDPOST = Trim(txtCDPOST)
recLrTiers.LBCOMM2 = Trim(txtLBCOMM2)
recLrTiers.CDDEPT2 = Trim(txtCDDEPT2)
recLrTiers.CDPAYS2 = Trim(txtCDPAYS2)

dbLrTiers_Update recLrTiers
cmdClear

End Sub
Private Sub cmdPrint_Click()
Me.PopupMenu cmdPrint_mnu

End Sub
Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim X As String

fraLrTiers.Enabled = False
X = Trim(mId$(Msg, 13, Len(Msg)))
txtRFBENF.Visible = True 'LrTiersAut.Valider

    tableLrTiers_Open
    updLrTiers = False
recLrTiers.Method = "MoveFirst"
cmddbLrTiers
cmdClear

End Sub

Public Sub cmdContext_Quit()
If fraLrTiers.Enabled Then
    cmdClear
Else
    Unload Me
End If

End Sub

Public Sub cmdContext_Return()
If fraLrTiers.Enabled Then
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
txtRFBENF = Trim(recLrTiers.RFBENF)
txtNomBNF = Trim(recLrTiers.NOMBNF)
txtPrénom = Trim(recLrTiers.PRENOM)
txtNSIREN = Trim(recLrTiers.NSIREN)
txtCDSEXE = Trim(recLrTiers.CDSEXE)
'txtAmjDATNAIS = Trim(recLrTiers.AmjDATNAIS)
txtCDPAYS1 = Trim(recLrTiers.CDPAYS1)
txtCDDEPT1 = Trim(recLrTiers.CDDEPT1)
txtCDCOMM1 = Trim(recLrTiers.CDCOMM1)
txtLBCOMM1 = Trim(recLrTiers.LBCOMM1)
txtNOMCJT = Trim(recLrTiers.NOMCJT)
txtCDACCO = Trim(recLrTiers.CDACCO)
txtCTJURI = Trim(recLrTiers.CTJURI)
txtCDRESI = Trim(recLrTiers.CDRESI)
txtNOVOIE = Trim(recLrTiers.NOVOIE)
txtCDPOST = Trim(recLrTiers.CDPOST)
txtLBCOMM2 = Trim(recLrTiers.LBCOMM2)
txtCDDEPT2 = Trim(recLrTiers.CDDEPT2)
txtCDPAYS2 = Trim(recLrTiers.CDPAYS2)
End Sub

Private Sub cmdPrint_mnuDétail_Click()
Dim Msg As String
'Msg = "000001" & Format$(arrLrTiersNb, "000000") & "D"

'prtLrTiersX Msg

End Sub

Private Sub cmdPrint_mnuList_Click()
Dim Msg As String
'Msg = "000001" & Format$(arrLrTiersNb, "000000") & "L"

'prtLrTiersX Msg

End Sub

Private Sub cmdQuit_Click()
cmdContext_Quit
End Sub

Private Sub cmdSuppress_Click()
recLrTiers.Method = constDelete
dbLrTiers_Update recLrTiers
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
tableLrTiers_Open
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
tableLrTiers_Close
End Sub


Private Sub lblAmjDATNAIS_Change()
'txtAmjDATNAIS_Control
End Sub

Private Sub txtAmjDATNAIS_Change()
'txtAmjDATNAIS_Control
End Sub


Private Sub txtAmjDATNAIS_GotFocus()
DTPicker_GotFocus txtAmjDATNAIS
End Sub


Private Sub txtAmjDATNAIS_LostFocus()
DTPicker_LostFocus txtAmjDATNAIS
'txtAmjDATNAIS_Control
End Sub


Private Sub txtCDACCO_GotFocus()
Call txt_GotFocus(txtCDACCO)
End Sub


Private Sub txtCDACCO_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtCDACCO_LostFocus()
Call txt_LostFocus(txtCDACCO)
End Sub


Private Sub txtCDCOMM1_GotFocus()
Call txt_GotFocus(txtCDCOMM1)
End Sub


Private Sub txtCDCOMM1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtCDCOMM1_LostFocus()
Call txt_LostFocus(txtCDCOMM1)
End Sub


Private Sub txtCDDEPT1_GotFocus()
Call txt_GotFocus(txtCDDEPT1)
End Sub


Private Sub txtCDDEPT1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtCDDEPT1_LostFocus()
Call txt_LostFocus(txtCDDEPT1)
End Sub


Private Sub txtCDDEPT2_GotFocus()
Call txt_GotFocus(txtCDDEPT2)
End Sub


Private Sub txtCDDEPT2_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtCDDEPT2_LostFocus()
Call txt_LostFocus(txtCDDEPT2)
End Sub


Private Sub txtCDPAYS1_GotFocus()
Call txt_GotFocus(txtCDPAYS1)
End Sub


Private Sub txtCDPAYS1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtCDPAYS1_LostFocus()
Call txt_LostFocus(txtCDPAYS1)
End Sub


Private Sub txtCDPAYS2_GotFocus()
Call txt_GotFocus(txtCDPAYS2)
End Sub


Private Sub txtCDPAYS2_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtCDPAYS2_LostFocus()
Call txt_LostFocus(txtCDPAYS2)
End Sub


Private Sub txtCDPOST_GotFocus()
Call txt_GotFocus(txtCDPOST)
End Sub


Private Sub txtCDPOST_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtCDPOST_LostFocus()
Call txt_LostFocus(txtCDPOST)
End Sub


Private Sub txtCDRESI_GotFocus()
Call txt_GotFocus(txtCDRESI)
End Sub


Private Sub txtCDRESI_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtCDRESI_LostFocus()
Call txt_LostFocus(txtCDRESI)
End Sub


Private Sub txtCDSEXE_GotFocus()
Call txt_GotFocus(txtCDSEXE)
End Sub


Private Sub txtCDSEXE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtCDSEXE_LostFocus()
Call txt_LostFocus(txtCDSEXE)
End Sub


Private Sub txtCTJURI_GotFocus()
Call txt_GotFocus(txtCTJURI)
End Sub


Private Sub txtCTJURI_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtCTJURI_LostFocus()
Call txt_LostFocus(txtCTJURI)
End Sub


Private Sub txtLBCOMM1_GotFocus()
Call txt_GotFocus(txtLBCOMM1)
End Sub


Private Sub txtLBCOMM1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtLBCOMM1_LostFocus()
Call txt_LostFocus(txtLBCOMM1)
End Sub


Private Sub txtLBCOMM2_GotFocus()
Call txt_GotFocus(txtLBCOMM2)
End Sub


Private Sub txtLBCOMM2_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtLBCOMM2_LostFocus()
Call txt_LostFocus(txtLBCOMM2)
End Sub


Private Sub txtNomBNF_GotFocus()
Call txt_GotFocus(txtNomBNF)

End Sub


Private Sub txtNomBNF_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtNomBNF_LostFocus()
Call txt_LostFocus(txtNomBNF)

End Sub


Private Sub txtNOMCJT_GotFocus()
Call txt_GotFocus(txtNOMCJT)
End Sub


Private Sub txtNOMCJT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtNOMCJT_LostFocus()
Call txt_LostFocus(txtNOMCJT)
End Sub


Private Sub txtNOVOIE_GotFocus()
Call txt_GotFocus(txtNOVOIE)
End Sub


Private Sub txtNOVOIE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtNOVOIE_LostFocus()
Call txt_LostFocus(txtNOVOIE)
End Sub


Private Sub txtNSIREN_GotFocus()
Call txt_GotFocus(txtNSIREN)
End Sub


Private Sub txtNSIREN_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtNSIREN_LostFocus()
Call txt_LostFocus(txtNSIREN)
End Sub


Private Sub txtPrénom_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtRFBENF_GotFocus()
Call txt_GotFocus(txtRFBENF)

End Sub


Private Sub txtRFBENF_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtRFBENF_LostFocus()
Call txt_LostFocus(txtRFBENF)
cmdRechercher
End Sub


Private Sub txtPrénom_GotFocus()
Call txt_GotFocus(txtPrénom)

End Sub


Private Sub txtPrénom_LostFocus()
Call txt_LostFocus(txtPrénom)

End Sub


Public Sub cmdClear()
txtRFBENF = ""
cmdContext.Visible = True 'LrTiersAut.Valider
cmdContext.Caption = constcmdRechercher
cmdSuppress.Visible = False
cmdPrint_mnuDétail.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "préciser le code ")
fraLrTiers.Enabled = False
txtRFBENF.Enabled = True ': txtRFBENF.SetFocus

End Sub

Public Sub cmdRechercher()
If Trim(txtRFBENF) = "" Then
    Call lstErr_Clear(lstErr, cmdContext, "préciser le code ")
    Exit Sub
End If

recLrTiers.Method = "Seek="
cmddbLrTiers
End Sub

Public Sub cmddbLrTiers()
recLrTiers.FILL02 = DTCENTenCours
recLrTiers.RFBENF = Trim(txtRFBENF)
dbLrTiers_Read recLrTiers

If recLrTiers.Err = 0 Then
    recLrTiers.Method = constUpdate
    cmdContext.Caption = constcmdEnregistrer
    cmdSuppress.Visible = True
Else
    Call lstErr_Clear(lstErr, cmdContext, "Erreur ")
    Exit Sub
End If
txtRFBENF.Enabled = False
fraLrTiers.Enabled = True
Rec_Display
txtNomBNF.SetFocus
End Sub


Private Sub UpDownLrTiers_DownClick(Index As Integer)
recLrTiers.Method = "MoveNext"
cmddbLrTiers
End Sub

Private Sub UpDownLrTiers_UpClick(Index As Integer)
recLrTiers.Method = "MovePrevious"
cmddbLrTiers
End Sub


