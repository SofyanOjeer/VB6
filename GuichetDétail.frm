VERSION 5.00
Begin VB.Form frmGuichetDétail 
   AutoRedraw      =   -1  'True
   Caption         =   "Guichet Détail"
   ClientHeight    =   6300
   ClientLeft      =   2595
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6300
   ScaleWidth      =   6885
   Begin VB.TextBox txtComplément3 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1200
      TabIndex        =   57
      Top             =   2640
      Width           =   5055
   End
   Begin VB.TextBox txtMontantAjustement 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1440
      TabIndex        =   47
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtLibellé 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1200
      TabIndex        =   45
      Top             =   1200
      Width           =   5055
   End
   Begin VB.TextBox txtOptAvis 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6120
      TabIndex        =   44
      Top             =   5640
      Width           =   375
   End
   Begin VB.TextBox txtAmjValeur 
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   5400
      TabIndex        =   43
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox txtMontantEspèces 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1440
      TabIndex        =   42
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox txtMontant 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1440
      TabIndex        =   33
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtComplément2 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   2280
      Width           =   5055
   End
   Begin VB.TextBox txtComplément1 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Top             =   1920
      Width           =   5055
   End
   Begin VB.TextBox txtIdentité 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   1560
      Width           =   5055
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   6480
      Picture         =   "GuichetDétail.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   0
      Width           =   400
   End
   Begin VB.CommandButton cmdSuivant 
      Caption         =   "&Suivant"
      Enabled         =   0   'False
      Height          =   300
      Left            =   5400
      TabIndex        =   10
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdPrécédent 
      Caption         =   "&Précédent"
      Enabled         =   0   'False
      Height          =   300
      Left            =   4320
      TabIndex        =   9
      Top             =   0
      Width           =   1000
   End
   Begin VB.TextBox txtCompte 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label libIntitulé 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2760
      TabIndex        =   58
      Top             =   840
      Width           =   4035
   End
   Begin VB.Label Label9 
      Caption         =   "Complément  3"
      Height          =   285
      Left            =   50
      TabIndex        =   56
      Top             =   2640
      Width           =   1395
   End
   Begin VB.Label libComptaAMJ 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2880
      TabIndex        =   55
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label libValidationAMJ 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2880
      TabIndex        =   54
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label libSaisieAMJ 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2880
      TabIndex        =   53
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label libchkSolde 
      Height          =   285
      Left            =   5520
      TabIndex        =   52
      Top             =   840
      Width           =   1425
   End
   Begin VB.Label Label8 
      Caption         =   "/"
      Height          =   285
      Left            =   5760
      TabIndex        =   51
      Top             =   3960
      Width           =   195
   End
   Begin VB.Label Label2 
      Caption         =   "/"
      Height          =   285
      Left            =   5280
      TabIndex        =   50
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label1 
      Caption         =   "Séquence"
      Height          =   285
      Left            =   1560
      TabIndex        =   49
      Top             =   360
      Width           =   825
   End
   Begin VB.Label libCptMvtPièceEspèces 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5520
      TabIndex        =   48
      Top             =   360
      Width           =   765
   End
   Begin VB.Label libCoursChangeEspèces 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6000
      TabIndex        =   46
      Top             =   3960
      Width           =   795
   End
   Begin VB.Label lblOptAvis 
      Caption         =   "Option Avis"
      Height          =   285
      Left            =   5160
      TabIndex        =   41
      Top             =   5760
      Width           =   915
   End
   Begin VB.Label lblMontantAjustement 
      Caption         =   "Ecart :"
      Height          =   285
      Left            =   50
      TabIndex        =   40
      Top             =   4320
      Width           =   555
   End
   Begin VB.Label lblCodeOpération 
      Caption         =   "Opération :"
      Height          =   285
      Left            =   50
      TabIndex        =   39
      Top             =   360
      Width           =   705
   End
   Begin VB.Label lblCptMvtPièce 
      Caption         =   "N° de Pièces"
      Height          =   285
      Left            =   3240
      TabIndex        =   38
      Top             =   360
      Width           =   1155
   End
   Begin VB.Label libCodeOpération 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   960
      TabIndex        =   37
      Top             =   360
      Width           =   735
   End
   Begin VB.Label libRéférence 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2400
      TabIndex        =   36
      Top             =   360
      Width           =   555
   End
   Begin VB.Label libDeviseEspèces 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3000
      TabIndex        =   35
      Top             =   3960
      Width           =   555
   End
   Begin VB.Label libDevise 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3000
      TabIndex        =   34
      Top             =   3600
      Width           =   435
   End
   Begin VB.Label lblMontant 
      Caption         =   "Montant Opération"
      Height          =   285
      Left            =   50
      TabIndex        =   32
      Top             =   3600
      Width           =   1395
   End
   Begin VB.Label lblMontantEspèce 
      Caption         =   "Espèces :"
      Height          =   285
      Left            =   50
      TabIndex        =   31
      Top             =   3960
      Width           =   795
   End
   Begin VB.Label lblAmjValeur 
      Caption         =   "date valeur"
      Height          =   285
      Left            =   4320
      TabIndex        =   30
      Top             =   3600
      Width           =   915
   End
   Begin VB.Label libCoursChange 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5040
      TabIndex        =   29
      Top             =   3960
      Width           =   795
   End
   Begin VB.Label lblCourChange 
      Caption         =   "Cours :"
      Height          =   285
      Left            =   4320
      TabIndex        =   28
      Top             =   3960
      Width           =   555
   End
   Begin VB.Label libComptaHMS 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4080
      TabIndex        =   27
      Top             =   5760
      Width           =   675
   End
   Begin VB.Label libValidationHMS 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4080
      TabIndex        =   26
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label libSaisieHMS 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4080
      TabIndex        =   25
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label libComptaUsr 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1440
      TabIndex        =   24
      Top             =   5760
      Width           =   915
   End
   Begin VB.Label libValidationUsr 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1440
      TabIndex        =   23
      Top             =   5400
      Width           =   1035
   End
   Begin VB.Label Label7 
      Caption         =   "à"
      Height          =   285
      Left            =   3720
      TabIndex        =   22
      Top             =   5760
      Width           =   315
   End
   Begin VB.Label Label6 
      Caption         =   "à"
      Height          =   285
      Left            =   3720
      TabIndex        =   21
      Top             =   5400
      Width           =   315
   End
   Begin VB.Label Label3 
      Caption         =   "à"
      Height          =   285
      Left            =   3720
      TabIndex        =   20
      Top             =   5040
      Width           =   195
   End
   Begin VB.Label libSaisieUsr 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1440
      TabIndex        =   19
      Top             =   5040
      Width           =   1035
   End
   Begin VB.Label Label5 
      Caption         =   "le"
      Height          =   285
      Left            =   2520
      TabIndex        =   18
      Top             =   5760
      Width           =   435
   End
   Begin VB.Label Label4 
      Caption         =   "le"
      Height          =   285
      Left            =   2520
      TabIndex        =   17
      Top             =   5400
      Width           =   675
   End
   Begin VB.Label lblComplément1 
      Caption         =   "Complément  1"
      Height          =   285
      Left            =   50
      TabIndex        =   16
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label lblComplément2 
      Caption         =   "Complément  2"
      Height          =   285
      Left            =   50
      TabIndex        =   15
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label lblComptaAMJ 
      Caption         =   "Comptabilisé Par :"
      Height          =   285
      Left            =   50
      TabIndex        =   8
      Top             =   5760
      Width           =   1260
   End
   Begin VB.Label lblValidationAMJ 
      Caption         =   "Validé Par :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   50
      TabIndex        =   7
      Top             =   5400
      Width           =   1485
   End
   Begin VB.Label lblSaisieAMJ 
      Caption         =   "Saisie Par :"
      Height          =   285
      Left            =   50
      TabIndex        =   6
      Top             =   5040
      Width           =   1035
   End
   Begin VB.Label lblBdfG13 
      Caption         =   "le"
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   5040
      Width           =   555
   End
   Begin VB.Label lblIdentité 
      Caption         =   "Identité"
      Height          =   285
      Left            =   50
      TabIndex        =   4
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label lblLibellé 
      Caption         =   "Libellé"
      Height          =   285
      Left            =   50
      TabIndex        =   3
      Top             =   1200
      Width           =   1185
   End
   Begin VB.Label lblCompte 
      Caption         =   "Compte"
      Height          =   285
      Left            =   50
      TabIndex        =   2
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label libCptMvtPièce 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4560
      TabIndex        =   1
      Top             =   360
      Width           =   675
   End
End
Attribute VB_Name = "frmGuichetDétail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean

Dim recGuichet As typeGuichet
'---------------------------------------------------------
Private Sub cmdOk_Click()
'---------------------------------------------------------

'If arrGuichetNb > 0 Then frmBdf.cmdBdfGOk_Click
Unload Me

End Sub

'---------------------------------------------------------
Private Sub cmdPrécédent_Click()
'---------------------------------------------------------

arrGuichetIndex = arrGuichetIndex - 1
'frmGuichet.lstOpération.ListIndex = frmGuichet.lstOpération.ListIndex - 1
'Call frmopérationInit
End Sub

'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
Dim Msg As String

'Msg = Format$(arrGuichetIndex, "000000") & Format$(arrGuichetIndex, "000000")
'prtBdfG Msg

End Sub

'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me

'lblOpération = CodeOpération

End Sub

'---------------------------------------------------------
Private Sub cmdSuivant_Click()
'---------------------------------------------------------

arrGuichetIndex = arrGuichetIndex + 1
'frmGuichetDétail.lstOpération.ListIndex = frmGuichetDétail.lstOpération.ListIndex + 1
'Call frmopérationInit
End Sub

'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
'arrGuichetIndex = frmGuichet.lstOpération.ListIndex + 1
frmGuichetDétail_Init
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
Private Sub frmGuichetDétail_Init()
Dim X As String, recCompte As typeCompte
'---------------------------------------------------------

If arrGuichetNb > 1 And arrGuichetIndex < arrGuichetNb Then
    cmdSuivant.Enabled = True
Else
    cmdSuivant.Enabled = False
End If

If arrGuichetNb > 1 And arrGuichetIndex > 1 Then
    cmdPrécédent.Enabled = True
Else
    cmdPrécédent.Enabled = False
End If

If arrGuichetIndex > 0 Then
    recGuichet = arrGuichet(arrGuichetIndex)
    recCompteInit recCompte
    recCompte.Société = recGuichet.Société
    recCompte.Agence = recGuichet.Agence
    recCompte.Devise = recGuichet.Devise
    recCompte.Numéro = recGuichet.Compte
    recCompte.BiaTyp = "000"
    recCompte.BiaNum = "00"
    
    '''srvCompteFind recCompte
    recCompte.Method = "SeekL1"
    If Not IsNull(mdbCptP0_Find(recCompte)) Then libIntitulé = "??????????????????????"

    libIntitulé = recCompte.Intitulé
    
    txtLibellé = recGuichet.Libellé
    txtIdentité = recGuichet.Identité
    txtComplément1 = recGuichet.Complément1
    txtComplément2 = recGuichet.Complément2
    txtComplément3 = recGuichet.Complément3
    
    libRéférence = recGuichet.Référence
    libCodeOpération = recGuichet.CodeOpération
    
    txtCompte = Compte_Imp(recGuichet.Compte)
   
    txtMontant = num_Display(recGuichet.Montant, 15, 2, Lx, X, "0")
    txtMontantEspèces = num_Display(recGuichet.MontantEspèces, 15, 2, Lx, X, "0")
    txtMontantAjustement = num_Display(recGuichet.MontantAjustement, 15, 2, Lx, X, "#")
    txtAMJValeur = dateImpS(recGuichet.AmjValeur)
    
    libDeviseEspèces = DevX(recGuichet.DeviseEspèces)
    libDevise = DevX(recGuichet.Devise)

   
   libSaisieUsr = recGuichet.SaisieUsr
   libValidationUsr = recGuichet.ValidationUsr
   libComptaUsr = recGuichet.ComptaUsr
   
   libSaisieHMS = timeImpHM(recGuichet.SaisieHMS)
   libValidationHMS = timeImpHM(recGuichet.ValidationHMS)
   libComptaHMS = timeImpHM(recGuichet.ComptaHMS)
   
  libSaisieAMJ = dateImpS(recGuichet.SaisieAmj)
   libValidationAMJ = dateImpS(recGuichet.ValidationAMJ)
   libComptaAMJ = dateImpS(recGuichet.ComptaAMJ)

   libCoursChange = recGuichet.CoursChange
   libCoursChangeEspèces = recGuichet.CoursChangeEspèces
   

  libCptMvtPièce = Format$(recGuichet.CptMvtPièce, "####") & "." & Format$(recGuichet.CptMvtLigne, "0000")
  libCptMvtPièceEspèces = Format$(recGuichet.CptMvtPièceEspèces, "####") & "." & Format$(recGuichet.CptMvtLigneEspèces, "0000")
  libchkSolde.FontBold = True
  libchkSolde.ForeColor = errUsr.ForeColor
  libchkSolde = ""
  If recGuichet.chkSolde <> "0" Then libchkSolde = "débiteur "
  If recGuichet.chkCompte <> "0" Then libchkSolde = libchkSolde & "bloqué"

  If recGuichet.optAvis <> "0" Then
      txtOptAvis = "non"
  Else
      txtOptAvis = "oui"
       
   End If

End If




End Sub











