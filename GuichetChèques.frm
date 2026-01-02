VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "Comct232.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmGuichetChèques 
   AutoRedraw      =   -1  'True
   Caption         =   "Guichet : remise de chèques"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6375
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Ok"
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
      TabIndex        =   37
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   6120
      TabIndex        =   15
      Top             =   0
      Width           =   2745
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8880
      Picture         =   "GuichetChèques.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   500
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   0
      TabIndex        =   17
      Top             =   1440
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Remise de chèques"
      TabPicture(0)   =   "GuichetChèques.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraGuichetChèques"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Comptabilité"
      TabPicture(1)   =   "GuichetChèques.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picCompta"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox picCompta 
         AutoRedraw      =   -1  'True
         FillColor       =   &H00E0E0E0&
         Height          =   4035
         Left            =   -74880
         ScaleHeight     =   3975
         ScaleWidth      =   9000
         TabIndex        =   20
         Top             =   420
         Width           =   9060
      End
      Begin VB.Frame fraGuichetChèques 
         Height          =   4335
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   8895
         Begin VB.Frame fraRemise 
            Height          =   1815
            Left            =   3240
            TabIndex        =   32
            Top             =   2040
            Width           =   5415
            Begin VB.TextBox txtComplément2 
               Height          =   285
               Left            =   2280
               MaxLength       =   50
               TabIndex        =   3
               Top             =   1200
               Width           =   2655
            End
            Begin VB.TextBox txtComplément1 
               Height          =   285
               Left            =   2280
               MaxLength       =   50
               TabIndex        =   2
               Top             =   720
               Width           =   2415
            End
            Begin VB.TextBox txtNb 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2280
               MaxLength       =   20
               TabIndex        =   1
               Top             =   240
               Width           =   660
            End
            Begin ComCtl2.UpDown UpDown1 
               Height          =   300
               Left            =   3000
               TabIndex        =   33
               Top             =   240
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   327681
               Value           =   10
               BuddyControl    =   "txtNb"
               BuddyDispid     =   196617
               OrigLeft        =   3600
               OrigTop         =   240
               OrigRight       =   3840
               OrigBottom      =   540
               Max             =   9999
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label lblAgence 
               Caption         =   "Agence du correspondant"
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   1320
               Width           =   1935
            End
            Begin VB.Label lblRéférence 
               Caption         =   "Référence de la remise"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   840
               Width           =   1815
            End
            Begin VB.Label lblNb 
               Caption         =   "Nb de chèques"
               Height          =   255
               Left            =   120
               TabIndex        =   34
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.Frame fraCompteBia 
            Caption         =   "Compte BIA à débiter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   3240
            TabIndex        =   26
            Top             =   1800
            Width           =   5415
            Begin VB.TextBox txtCompteBia 
               Height          =   285
               Left            =   1440
               TabIndex        =   4
               Top             =   840
               Width           =   1455
            End
            Begin VB.TextBox txtChèqueNo 
               Height          =   285
               Left            =   1440
               MaxLength       =   7
               TabIndex        =   5
               Top             =   1320
               Width           =   1455
            End
            Begin VB.ListBox lstOppChq 
               Height          =   1035
               Left            =   3120
               TabIndex        =   27
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label lblCompteBia 
               Caption         =   "Compte"
               Height          =   255
               Left            =   240
               TabIndex        =   31
               Top             =   840
               Width           =   855
            End
            Begin VB.Label lblChèqueNo 
               Caption         =   "N° chèque"
               Height          =   255
               Left            =   240
               TabIndex        =   30
               Top             =   1320
               Width           =   855
            End
            Begin VB.Label lblOppChq 
               Caption         =   "OPPOSITION >>>"
               Height          =   255
               Left            =   3120
               TabIndex        =   29
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label libCompteBia 
               Caption         =   "-"
               Height          =   255
               Left            =   240
               TabIndex        =   28
               Top             =   360
               Width           =   4455
            End
         End
         Begin VB.ListBox lstCompte 
            Height          =   1230
            Left            =   360
            TabIndex        =   25
            Top             =   2880
            Width           =   2295
         End
         Begin VB.Frame fraDétail 
            Height          =   1335
            Left            =   3240
            TabIndex        =   21
            Top             =   240
            Width           =   5415
            Begin VB.CheckBox chkAmjValeur 
               Caption         =   "Date valeur manuelle"
               Height          =   255
               Left            =   240
               TabIndex        =   24
               Top             =   360
               Width           =   1935
            End
            Begin VB.TextBox txtMontant 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2520
               MaxLength       =   20
               TabIndex        =   0
               Top             =   840
               Width           =   1500
            End
            Begin MSComCtl2.DTPicker txtAmjValeur 
               Height          =   300
               Left            =   2520
               TabIndex        =   14
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   28246019
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label libDeviseIso 
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4320
               TabIndex        =   23
               Top             =   960
               Width           =   735
            End
            Begin VB.Label lblMonatnt 
               Caption         =   "montant de la remise"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   22
               Top             =   960
               Width           =   1935
            End
         End
         Begin VB.Frame fraPlace 
            Caption         =   "Place"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Left            =   360
            TabIndex        =   19
            Top             =   360
            Width           =   2295
            Begin VB.OptionButton optPlaceEtranger 
               Caption         =   "Etranger"
               Height          =   375
               Left            =   120
               TabIndex        =   10
               Top             =   1680
               Width           =   1215
            End
            Begin VB.OptionButton optPlaceHorsRayon 
               Caption         =   "hors rayon"
               Height          =   375
               Left            =   120
               TabIndex        =   7
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton optPlaceSurRayon 
               Caption         =   "sur rayon"
               Height          =   375
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Width           =   1575
            End
            Begin VB.OptionButton optPlaceBia 
               Caption         =   "BIA"
               Height          =   375
               Left            =   120
               TabIndex        =   8
               Top             =   960
               Width           =   1215
            End
            Begin VB.OptionButton optPlaceDomTom 
               Caption         =   "Dom Tom"
               Height          =   375
               Left            =   120
               TabIndex        =   9
               Top             =   1320
               Width           =   1215
            End
         End
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C0FF&
      Caption         =   "en &Attente"
      Default         =   -1  'True
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
      TabIndex        =   12
      Top             =   0
      Width           =   1200
   End
   Begin VB.PictureBox picCpt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   900
      Left            =   0
      ScaleHeight     =   840
      ScaleWidth      =   9315
      TabIndex        =   16
      Top             =   480
      Width           =   9375
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   11
      Top             =   0
      Width           =   1200
   End
End
Attribute VB_Name = "frmGuichetChèques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String

Dim ClientCompte As typeCompte, ContrepartieCompte As typeCompte, CompteMin As typeCompte, CompteMax As typeCompte
Dim ClientRacine As typeRacine
Dim chkLevel As String * 1


Dim CV As typeCV
Dim wG_CV1 As typeCV, wG_CV2 As typeCV, wG_CV3 As typeCV
Dim maxDevise1D As Integer, maxDevise2D As Integer
Dim xConversion As String * 1

Dim recTable As typeElpTable
Dim valAMJ As String, valAMJ1 As String, valAMJ2 As String

Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
         If TypeOf xobj Is CommandButton Or TypeOf xobj Is ListBox Then
           xobj.BackColor = MouseMoveActiveControl.BackColor
        Else
            xobj.ForeColor = MouseMoveActiveControl.ForeColor
        End If
        Exit For
    End If
Next xobj

End Sub

Public Sub MouseMoveActiveControl_Set(C As Control)
If MouseMoveActiveControl_Name <> C.Name Then
    MouseMoveActiveControl_Reset
    If Not C.Enabled Then
        MouseMoveActiveControl_Name = ""
    Else
        MouseMoveActiveControl_Name = C.Name
        If TypeOf C Is CommandButton Or TypeOf C Is ListBox Then
            
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            MouseMoveActiveControl.ForeColor = C.ForeColor
            C.ForeColor = MouseMoveUsr.ForeColor
        End If
    End If
End If

End Sub

Private Sub frmCompte_Show()
Exit Sub ' $$$$$$$$$$$$$$$$$$$$$$$$$$ pb vbmodal

X = Space$(100)
Mid$(X, 1, 12) = "frmCompte   "
Mid$(X, 13, 12) = "frmGuichetEs"
Mid$(X, 25, 10) = Space$(10)
Mid$(X, 35, 3) = ClientCompte.Devise
Mid$(X, 38, 11) = ClientCompte.Numéro
Msg_Monitor X

End Sub

Public Sub cmdContext_Quit()
blnControl = False
If blnMsgBox_Quit Then
    X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
 Else
    X = vbYes
 End If
 If X = vbYes Then Unload Me

End Sub


Public Sub cmdContext_Return()

If cmdOk.Visible Then 'And ActiveControl.Name = lastActiveControl_Name Then
    cmdContext.SetFocus
    X = MsgBox("Voulez-vous enregistrer cette opération?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirmation de saisie")
    If X = vbYes Then
        cmdOk_Click
    Else
        txtMontant.SetFocus
    End If
Else
    SendKeys "{TAB}"
End If

End Sub






'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
'lstErr.Clear
currentActiveControl_Name = C.Name
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
End Sub
'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
'lstErr.Clear
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub

Private Sub chkAmjValeur_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkAmjValeur

End Sub

Private Sub chkAmjValeur_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkAmjValeur = "1" Then
    txtAmjValeur.Enabled = True
Else
    txtAmjValeur.Enabled = False
End If
If blnControl Then cmdControl

End Sub

Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
Dim I As Integer, Msg As String, X As String

prtCompta.CV1 = G_CV1
prtCompta.CV2 = G_CV2
prtCompta.CV3 = G_CV3

For I = 1 To G_arrCV030Nb
    prtCompta.arrCV030(I) = G_arrCV030(I)
Next I

Msg = Format$(1, "000000") & Format$(G_arrCV030Nb, "000000")
Me.Hide
prtCompta_Monitor Msg, "", currentAction
Me.Show vbModal

End Sub


'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me

End Sub




Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext
End Sub

Private Sub cmdOk_Click()
G_recGuichet.ValidationUsr = ""
cmdSave_Db

End Sub

Private Sub cmdOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdOk
End Sub


Private Sub cmdSave_Click()
G_recGuichet.ValidationUsr = constEnAttente
cmdSave_Db

End Sub


Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdSave
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
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
'   Case Is = 34: cmdPageNext_Click
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
End Sub
'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
usrColor_Set

lblOppChq.Visible = False
lblOppChq.ForeColor = vbRed
lstOppChq.Visible = False
lstOppChq.ForeColor = vbRed

picCpt.Cls
SSTab1.Tab = 0
cmdContext.Caption = constcmdAbandonner: blnMsgBox_Quit = False
cmdOk.Visible = False: cmdSave.Visible = False
arrTag_Set False
lstErr.Visible = False
blnAddNew = True
txtMontant = ""
txtNb = ""
libDeviseIso = ""
chkAmjValeur = "0"
txtAmjValeur.Enabled = False
lastActiveControl_Name = "txtComplément2"
recCompteInit ContrepartieCompte

chkLevel = "1"
If G_CV1.DeviseIso <> "FRF" And G_CV1.DeviseIso <> "EUR" Then
    optPlaceSurRayon.Enabled = False
    optPlaceHorsRayon.Enabled = False
    optPlaceDomTom.Enabled = False
End If

End Sub



'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub



Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------
End Sub


Public Sub Msg_Snd(ByVal X As String)
End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub

Private Sub fraplace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fratypedecompte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraDétail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraguichetchèques_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub
Public Sub cmdControl()
Dim X As String, ChèqueNb As Integer
Dim valX As String
Dim curX As Currency
Dim dblX As Double
Dim C As Control, blnFocus As Boolean
Dim blnOk As Boolean

If Not frmGuichetChèques.Enabled Then Exit Sub
frmGuichetChèques.Enabled = False
cmdOk.Visible = False
cmdSave.Visible = False

blnControl = False
blnFocus = False
blnOk = False

'For Each C In Me.Controls
'    If TypeOf C Is TextBox Then
'        If C.Name = currentActiveControl_Name Then
'            blnFocus = True
'            Exit For
'        End If
'    End If
'Next C

lstErr.Clear
lstErr.Height = 200
picCompta.Cls

G_recGuichet.CodeOpération = "    "
If optPlaceBia Then
    G_recGuichet.CodeOpération = "G010"
    fraRemise.Visible = False
    fraCompteBia.Visible = True
    txtNb = "1"
Else
    fraRemise.Visible = True
    fraCompteBia.Visible = False
End If
If optPlaceSurRayon Then
    G_recGuichet.CodeOpération = "G011"
    G_recGuichet.ContrepartieCompte = paramGuichetCompensateur
End If
If optPlaceHorsRayon Then
    G_recGuichet.CodeOpération = "G012"
    G_recGuichet.ContrepartieCompte = paramGuichetCompensateur
End If
If optPlaceDomTom Then
    G_recGuichet.CodeOpération = "G013"
    G_recGuichet.ContrepartieCompte = paramGuichetCompensateur
End If
If optPlaceEtranger Then
    G_recGuichet.CodeOpération = "G014"
    G_recGuichet.ContrepartieCompte = paramGuichetCompensateur
    If ClientCompte.BiaTyp <> "010" Then Call lstErr_AddItem(lstErr, cmdContext, "? Le type de compte doit être '010'")
End If

If G_recGuichet.CodeOpération = "    " Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser la place")

mGuichetAmjValeurCompensateur = paramGuichetAmjValeurSRCompensateur
mGuichetAmjEchRecouvreur = paramGuichetAmjEchSRRecouvreur

G_recGuichet.chkAmjValeur = chkAmjValeur
If G_recGuichet.chkAmjValeur = "1" Then
    If G_recGuichet.AmjValeur < paramGuichetAmjValeurMin Then Call lstErr_AddItem(lstErr, cmdContext, "? Date valeur < " & dateImp(paramGuichetAmjValeurMin))
    If G_recGuichet.AmjValeur > paramGuichetAmjValeurMax Then Call lstErr_AddItem(lstErr, cmdContext, "? Date valeur > " & dateImp(paramGuichetAmjValeurMax))
Else
    G_recGuichet.AmjValeur = G_recGuichet.AmjOpération
    
    Select Case G_recGuichet.CodeOpération
        Case "G011": G_recGuichet.AmjValeur = paramGuichetAmjValeurSR
        Case "G012": G_recGuichet.AmjValeur = paramGuichetAmjValeurHR
                    mGuichetAmjValeurCompensateur = paramGuichetAmjValeurHRCompensateur
                    mGuichetAmjEchRecouvreur = paramGuichetAmjEchHRRecouvreur
        Case "G013": G_recGuichet.AmjValeur = paramGuichetAmjValeurDom
                    mGuichetAmjValeurCompensateur = paramGuichetAmjValeurDomCompensateur
                    mGuichetAmjEchRecouvreur = paramGuichetAmjEchDomRecouvreur
        Case "G014": G_recGuichet.AmjValeur = G_recGuichet.AmjOpération
                    mGuichetAmjValeurCompensateur = G_recGuichet.AmjOpération
                    mGuichetAmjEchRecouvreur = paramGuichetAmjEchDomRecouvreur
    End Select
    If ClientRacine.MembreduPersonnel = "1" And G_recGuichet.CodeOpération <> "G010" Then G_recGuichet.AmjValeur = paramGuichetAmjValeurPersonnel
    
End If


Call DTPicker_Set(txtAmjValeur, G_recGuichet.AmjValeur)



X = num_Control(txtNb, valX, 6, 0)
ChèqueNb = CInt(valX)
If ChèqueNb <= 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? nombre de chèques")


X = num_Control(txtMontant, valX, 13, maxDevise1D)
G_CV1.Montant = valX
If G_CV1.Montant <= 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? montant"): GoTo ExitSub

Call CV_Transitoire(G_CV1, G_CV2, G_CV3, xConversion)

G_recGuichet.MontantEspèces = G_CV1.Montant
G_recGuichet.Montant = G_CV2.Montant
G_recGuichet.MontantAjustement = 0

G_recGuichet.CoursChangeEspèces = G_CV1.Cours
G_recGuichet.CoursChange = G_CV2.Cours

txtMontant = Format$(G_recGuichet.MontantEspèces, "### ### ### ##0.00")
''txtNb = Format$(G_recGuichet.Nb, "###0")


G_recGuichet.MontantEuro = G_CV3.Montant
G_recGuichet.Conversion = xConversion

G_recGuichet.Complément1 = Trim(txtComplément1)
G_recGuichet.Complément2 = Trim(txtComplément2)

Compte_Control


If optPlaceBia Then
    G_recGuichet.ContrepartieCompte = txtCompteBia
    ContrepartieCompte_Load
    txtChèqueNo_Control
    mGuichetAmjValeurCompensateur = G_recGuichet.AmjOpération
    If G_recGuichet.Compte = G_recGuichet.ContrepartieCompte Then Call lstErr_AddItem(lstErr, cmdContext, "? Tireur = Bénéficiaire")
    If ContrepartieCompte.TypeGA = "A" Then
        If ContrepartieCompte.SoldeXXX - G_recGuichet.Montant < 0 Then
            If G_recGuichet.chkSolde < chkLevel Then
'                    Call lstErr_AddItem(lstErr, cmdContext, "Dépassement :" & num_Display(ContrepartieCompte.SoldeXXX - G_recGuichet.Montant, 15, 2, Lx, X, "0"))
                X = "Dépassement :" & num_Display(ContrepartieCompte.SoldeXXX - G_recGuichet.Montant, 15, 2, Lx, X, "0")
                X = MsgBox(X & ",Confirmez-vous cette opération ?", vbYesNo + vbQuestion + vbDefaultButton2, Trim(ContrepartieCompte.Intitulé) & " : " & currentAction)
                If X = vbYes Then
                    G_recGuichet.chkSolde = chkLevel
                Else
                    Call lstErr_AddItem(lstErr, txtMontant, "Le compte est en dépassement")
                End If
            End If
        End If
    End If

Else
    If ClientCompte.TypeGA = "A" And ClientCompte.BiaTyp = "010" Then
        G_recGuichet.ContrepartieCompte = paramGuichetRecouvreur
        Mid$(G_recGuichet.Complément3, 15, 8) = mGuichetAmjEchRecouvreur
        G_recGuichet.optVirement = "R"
        mGuichetAmjValeurCompensateur = G_recGuichet.AmjValeur
    Else
        G_recGuichet.ContrepartieCompte = paramGuichetCompensateur
    End If
End If

G_recGuichet.Complément3 = mGuichetAmjValeurCompensateur & " " & Format$(ChèqueNb, "0000")

Guichet_Compta.Init
Guichet_Compta.Libellé
Guichet_Compta.Gen
Guichet_Compta.Display picCompta, lstErr

If lstErr.ListCount = 0 Then
    blnOk = True 'cmdOk.Visible = True
    If G_recGuichet.chkCompte <> "0" Or G_recGuichet.chkSolde <> "0" Then cmdSave.Visible = True
End If

ExitSub:

If blnOk Then cmdOk.Visible = True
frmGuichetChèques.Enabled = True
If blnOk Then cmdOk.SetFocus  '99-12-16 Visible = False: cmdOk.Visible = True
   
'If cmdOk.Visible Then cmdOk.Visible = False: cmdOk.Visible = True
blnControl = True
End Sub

Private Sub lstCompte_Click()
G_recGuichet.Compte = Trim(lstCompte)
Compte_Load
If blnControl Then cmdControl

End Sub

Private Sub optPlaceBia_Click()
If blnControl Then cmdControl
End Sub

Private Sub optPlaceBia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optPlaceBia
End Sub


Private Sub optPlaceDomTom_Click()
If blnControl Then cmdControl

End Sub

Private Sub optPlaceDomTom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optPlaceDomTom
End Sub


Private Sub optPlaceEtranger_Click()
If blnControl Then cmdControl

End Sub

Private Sub optPlaceEtranger_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optPlaceEtranger
End Sub


Private Sub optPlaceHorsRayon_Click()
If blnControl Then cmdControl

End Sub

Private Sub optPlaceHorsRayon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optPlaceHorsRayon
End Sub


Private Sub optPlaceSurRayon_Click()
If blnControl Then cmdControl

End Sub

Private Sub optplacesurrayon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optPlaceSurRayon
End Sub


Private Sub picCpt_Click()
frmCompte_Show
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If PreviousTab = 0 And blnControl Then cmdControl
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set SSTab1
End Sub


Private Sub txtChèqueNo_GotFocus()
Call txt_GotFocus(txtChèqueNo)
End Sub



Private Sub txtChèqueNo_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub



Private Sub txtChèqueNo_LostFocus()
Call txt_LostFocus(txtChèqueNo)
If blnControl Then cmdControl
End Sub



Private Sub txtComplément1_GotFocus()
Call txt_GotFocus(txtComplément1)
End Sub


Private Sub txtComplément1_LostFocus()
Call txt_LostFocus(txtComplément1)
If blnControl Then cmdControl
End Sub


Private Sub txtComplément2_GotFocus()
Call txt_GotFocus(txtComplément2)
End Sub


Private Sub txtComplément2_LostFocus()
Call txt_LostFocus(txtComplément2)
If blnControl Then cmdControl

End Sub



Private Sub txtCompteBia_GotFocus()
Call txt_GotFocus(txtCompteBia)

End Sub


Private Sub txtCompteBia_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtCompteBia_LostFocus()
Call txt_LostFocus(txtCompteBia)
If blnControl Then cmdControl

End Sub


Private Sub txtMontant_GotFocus()
Call txt_GotFocus(txtMontant)
End Sub


Private Sub txtMontant_KeyPress(KeyAscii As Integer)
If G_CV1.maxD = 0 Then
    Call num_KeyAscii(KeyAscii)
Else
    Call num_KeyAsciiD(KeyAscii, txtMontant)
End If

End Sub


Private Sub txtMontant_LostFocus()
Call txt_LostFocus(txtMontant)
If blnControl Then cmdControl
End Sub


Private Sub txtNb_Change()
If blnControl Then cmdControl

End Sub

Private Sub txtnb_GotFocus()
Call txt_GotFocus(txtNb)
End Sub


Private Sub txtnb_KeyPress(KeyAscii As Integer)
If G_CV2.maxD = 0 Then
    Call num_KeyAscii(KeyAscii)
Else
    Call num_KeyAsciiD(KeyAscii, txtNb)
End If
End Sub

Private Sub txtnb_LostFocus()
Call txt_LostFocus(txtNb)
If blnControl Then cmdControl

End Sub





Public Sub Form_Init(Msg As String)

lstCompte.Visible = False
fraRemise.Visible = True
fraCompteBia.Visible = False

tableElpTable_Open
currentAction = Trim(mId$(Msg, 25, 12))

Guichet_Compta.CV_Reset currentAction

G_CV1.Normal = "C"
G_CV2.Normal = "C"

G_CV1.DeviseIso = mId$(Msg, 38, 3)
If Not IsNull(CV_Attribut(G_CV1)) Then Call MsgBox("Devise1 inconnue: " & mId$(Msg, 38, 3), vbCritical, "frmGuichetChèques.Form_Init")
G_CV2.DeviseIso = mId$(Msg, 41, 3)
If Not IsNull(CV_Attribut(G_CV2)) Then Call MsgBox("Devise2 inconnue: " & mId$(Msg, 41, 3), vbCritical, "frmGuichetChèques.Form_Init")

cmdReset

G_recGuichet_Init mId$(Msg, 44, 11)

G_CV1.OpéAmj = G_recGuichet.AmjOpération
G_CV1.OpéAmj = G_recGuichet.AmjOpération
valAMJ1 = G_CV1.OpéAmj: valAMJ2 = G_CV1.OpéAmj

maxDevise1D = G_CV1.maxD
maxDevise2D = G_CV2.maxD
libDeviseIso = G_CV1.DeviseIso

CompteRacine_Load

Call DTPicker_Set(txtAmjValeur, DSys)
paramGuichetAmjValeurSR = DateElp_X(paramGuichetJValeurSR, G_recGuichet.AmjOpération)
paramGuichetAmjValeurHR = DateElp_X(paramGuichetJValeurHR, G_recGuichet.AmjOpération)
paramGuichetAmjValeurDom = DateElp_X(paramGuichetJValeurDom, G_recGuichet.AmjOpération)
paramGuichetAmjValeurMin = DateElp_X(paramGuichetJValeurMin, G_recGuichet.AmjOpération)
paramGuichetAmjValeurMax = DateElp_X(paramGuichetJValeurMax, G_recGuichet.AmjOpération)
paramGuichetAmjValeurPersonnel = DateElp_X(paramGuichetJValeurPersonnel, G_recGuichet.AmjOpération)

paramGuichetAmjEchSRRecouvreur = DateElp_X(paramGuichetJEchSRRecouvreur, G_recGuichet.AmjOpération)
paramGuichetAmjEchHRRecouvreur = DateElp_X(paramGuichetJEchHRRecouvreur, G_recGuichet.AmjOpération)
paramGuichetAmjEchDomRecouvreur = DateElp_X(paramGuichetJEchDomRecouvreur, G_recGuichet.AmjOpération)

paramGuichetAmjValeurSRCompensateur = DateElp_X(paramGuichetJValeurSRCompensateur, G_recGuichet.AmjOpération)
paramGuichetAmjValeurHRCompensateur = DateElp_X(paramGuichetJValeurHRCompensateur, G_recGuichet.AmjOpération)
paramGuichetAmjValeurDomCompensateur = DateElp_X(paramGuichetJValeurDomCompensateur, G_recGuichet.AmjOpération)

cmdControl

frmGuichetChèques.Show vbModal
If txtMontant.Visible And txtMontant.Enabled Then txtMontant.SetFocus

End Sub
'-------------------------------------------------'
Private Sub txtAmjValeur_Control()
'-------------------------------------------------'

Dim X As String
X = Format$(txtAmjValeur.Year, "0000") & Format$(txtAmjValeur.Month, "00") & Format$(txtAmjValeur.Day, "00")
If Not IsNumeric(X) Then
    Call lstErr_AddItem(lstErr, cmdContext, "? erreur date")
    DTPicker_Now txtAmjValeur
Else
    G_recGuichet.AmjValeur = mId$(X, 1, 8)
End If

End Sub

Private Sub txtAmjValeur_Change()
txtAmjValeur_Control
If blnControl Then cmdControl
End Sub


Private Sub txtAmjValeur_GotFocus()
DTPicker_GotFocus txtAmjValeur

End Sub


Private Sub txtAmjValeur_LostFocus()
DTPicker_LostFocus txtAmjValeur
txtAmjValeur_Control

End Sub


Private Sub txtAmjValeur_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DTPicker_GotFocus txtAmjValeur

End Sub


Public Function Compte_Control() As String
Dim X As String
X = vbYes
If ClientCompte.Situation <> " " Then
    If G_recGuichet.chkCompte < chkLevel Then
        X = MsgBox("Le compte est bloqué, confirmez-vous cette opération ?", vbYesNo + vbQuestion + vbDefaultButton2, Trim(ClientCompte.Intitulé) & " : " & currentAction)
        If X = vbYes Then
            G_recGuichet.chkCompte = chkLevel
        Else
            Call lstErr_AddItem(lstErr, txtMontant, "Le compte est bloqué")
        End If
    End If
End If
Compte_Control = X

End Function

Public Sub G_recGuichet_Init(xNuméro As String)
recGuichet_Init G_recGuichet
G_recGuichet.Method = constAddNew
G_recGuichet.Séquence = 1
G_recGuichet.Société = SocId$
G_recGuichet.Agence = SocAgence$
G_recGuichet.Journal = constChèque
G_recGuichet.Devise = G_CV2.DeviseN
G_recGuichet.Compte = xNuméro
G_recGuichet.Sens = "C"
G_recGuichet.AmjOpération = paramGuichetAMJValeur
G_recGuichet.AmjValeur = paramGuichetAMJValeur
G_recGuichet.chkCompte = "0"
G_recGuichet.chkSolde = "0"
G_recGuichet.chkAmjOpération = "0"
G_recGuichet.chkAmjValeur = "0"
G_recGuichet.optAvis = "0"
G_recGuichet.optVirement = "0"
G_recGuichet.optSwift = "0"
G_recGuichet.optAvisLangue = "1"

G_recGuichet.CptMvtPièce = 0
G_recGuichet.CptMvtLigne = 0
G_recGuichet.CptMvtService = paramGuichetService
G_recGuichet.CptMvtExonéré = "0"
G_recGuichet.chkChèque = "0"
G_recGuichet.NoChèque = ""
G_recGuichet.CoursChange = 1
G_recGuichet.CoursChangeEspèces = 1

G_recGuichet.optCours = "0"
G_recGuichet.DeviseEspèces = G_CV1.DeviseN

G_recGuichet.chkCoupureEspèces = " "
G_recGuichet.chkCoupureChange = " "

G_recGuichet.SaisieAmj = DSys
G_recGuichet.SaisieHMS = time_Hms
G_recGuichet.SaisieUsr = usrId

End Sub

Public Sub cmdSave_Db()
cmdControl
If lstErr.ListCount = 0 Then
    blnControl = False
    cmdOk.Visible = False
    If IsNull(srvGuichet_Update(G_recGuichet)) Then
        lastActiveControl_Name = ""
        cmdQuit_Click
    End If
End If
End Sub


Public Sub Compte_Load()
Dim I As Integer

recCompteInit ClientCompte
ClientCompte.Société = G_recGuichet.Société
ClientCompte.Agence = G_recGuichet.Agence
ClientCompte.Devise = G_recGuichet.Devise
ClientCompte.Numéro = G_recGuichet.Compte
ClientCompte.BiaTyp = "000"
ClientCompte.BiaNum = "00"
ClientCompte.Method = "SeekL1"
If Not IsNull(srvCompteMon(ClientCompte)) Then Call lstErr_AddItem(lstErr, lstErr, "? compte en " & ClientCompte.Devise): Exit Sub
recCompte_Display ClientCompte, picCpt

End Sub

Public Sub ContrepartieCompte_Load()
Dim I As Integer
If Trim(G_recGuichet.ContrepartieCompte) = "" Then Call lstErr_AddItem(lstErr, lstErr, "? compte BIA à débiter " & ContrepartieCompte.Devise): Exit Sub

If ContrepartieCompte.Numéro <> G_recGuichet.ContrepartieCompte Then
    recCompteInit ContrepartieCompte
    ContrepartieCompte.Société = G_recGuichet.Société
    ContrepartieCompte.Agence = G_recGuichet.Agence
    ContrepartieCompte.Devise = G_recGuichet.Devise
    ContrepartieCompte.Numéro = G_recGuichet.ContrepartieCompte
    ContrepartieCompte.BiaTyp = "000"
    ContrepartieCompte.BiaNum = "00"
    ContrepartieCompte.Method = "SeekL1"
    If Not IsNull(srvCompteMon(ContrepartieCompte)) Then Call lstErr_AddItem(lstErr, lstErr, "? compte BIA à débiter " & ContrepartieCompte.Devise): Exit Sub
    libCompteBia = ContrepartieCompte.Intitulé
    Call OppChq_Load(G_recGuichet.ContrepartieCompte, lstOppChq)
    If G_arrOppChq_Numéro_Nb > 0 Then lstOppChq.Visible = True: lblOppChq.Visible = True

End If

End Sub

Public Sub CompteRacine_Load()
Dim I As Integer

Compte_Load
If ClientCompte.TypeGA <> "A" Then Exit Sub

recRacineInit ClientRacine

ClientRacine.Method = "SeekL0"
ClientRacine.Numéro = Val(mId$(ClientCompte.Numéro, 1, 5))
If Not IsNull(srvRacineMon(ClientRacine)) Then Call lstErr_AddItem(lstErr, lstErr, "? Racine " & ClientRacine.Numéro): Exit Sub

lstCompte.Clear
CompteMin = ClientCompte
CompteMin.BiaTyp = "001"
CompteMin.Method = "SnapL1"
CompteMin.chkAnnul = "0"
CompteMin.MvtAmj = "19000101"
Mid$(CompteMin.Numéro, 6, 6) = "010000"
CompteMax = CompteMin
Mid$(CompteMax.Numéro, 6, 6) = "010999"
If Not IsNull(selCompte_Load(CompteMin, CompteMax, "Init")) Then Exit Sub


CompteMin = ClientCompte
CompteMin.BiaTyp = "010"
CompteMin.Method = "SnapL1"
CompteMin.chkAnnul = "0"
CompteMin.MvtAmj = "19000101"
Mid$(CompteMin.Numéro, 6, 6) = "010000"
CompteMax = CompteMin
Mid$(CompteMax.Numéro, 6, 6) = "010999"
If Not IsNull(selCompte_Load(CompteMin, CompteMax, "Add")) Then Exit Sub

CompteMin = ClientCompte
CompteMin.BiaTyp = "851"
CompteMin.Method = "SnapL1"
CompteMin.chkAnnul = "0"
CompteMin.MvtAmj = "19000101"
Mid$(CompteMin.Numéro, 6, 6) = "010000"
CompteMax = CompteMin
Mid$(CompteMax.Numéro, 6, 6) = "010999"
If Not IsNull(selCompte_Load(CompteMin, CompteMax, "Add")) Then Exit Sub

If selCompte_Nb > 0 Then
    lstCompte.Visible = True
    For I = 1 To selCompte_Nb
        Select Case selCompte(I).BiaTyp
            Case "001": X = " ordinaire"
            Case "010": X = " exigible après encaissement"
            Case "851": X = "créances douteuses"
        End Select
        
        lstCompte.AddItem selCompte(I).Numéro & X
    Next I
End If
Call selCompte_Load(CompteMin, CompteMax, "End")
End Sub

Public Function txtChèqueNo_Control() As Boolean
txtChèqueNo_Control = False
If Trim(txtChèqueNo) = "" Then Call lstErr_AddItem(lstErr, txtChèqueNo, "? N° chèque"): Exit Function
strOppChq_Numéro = Format(Val(txtChèqueNo), "0000000")
G_recGuichet.NoChèque = strOppChq_Numéro
For I = 1 To G_arrOppChq_Numéro_Nb
    If G_arrOppChq_Numéro(I) = strOppChq_Numéro Then
        Call lstErr_AddItem(lstErr, txtChèqueNo, "? opposition sur chèque")
        Exit Function
    End If
Next I
txtChèqueNo_Control = True

End Function

Private Sub UpDown1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'lstErr.Clear
End Sub


