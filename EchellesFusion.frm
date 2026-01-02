VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmEchellesFusion 
   AutoRedraw      =   -1  'True
   Caption         =   "Echelles : fusion des mouvements "
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6375
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   400
      Left            =   8880
      Picture         =   "EchellesFusion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   500
   End
   Begin MSFlexGridLib.MSFlexGrid fgEchellesFusion 
      Height          =   2850
      Left            =   3360
      TabIndex        =   0
      Top             =   480
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   5027
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   300
      BackColor       =   14737632
      ForeColor       =   12582912
      ForeColorFixed  =   -2147483641
      BackColorSel    =   12648384
      BackColorBkg    =   14737632
      AllowBigSelection=   0   'False
      TextStyleFixed  =   4
      FocusRect       =   2
      HighLight       =   0
      GridLines       =   2
      FormatString    =   "Devise Compte Fusion      |Devise Compte à fusionner  |      Dernier arrêté | état              "
   End
   Begin VB.Frame fraEchellesFusion 
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
      TabIndex        =   10
      Top             =   3480
      Width           =   9255
      Begin VB.TextBox txtCompteFusion 
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton fraEchellesFusion_cmdQuit 
         BackColor       =   &H00C0C0FF&
         Caption         =   "X"
         Height          =   525
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Width           =   1500
      End
      Begin VB.CommandButton fraEchellesFusion_cmdOk 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ok"
         Height          =   525
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1560
         Width           =   1500
      End
      Begin VB.TextBox txtDeviseFusion 
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtDeviseOrigine 
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtCompteOrigine 
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtAmjDébut 
         Height          =   300
         Left            =   3000
         TabIndex        =   5
         Top             =   1680
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   14
         Mask            =   "## - ## - ####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblValidité 
         Caption         =   "Date du dernier arrêté"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1800
         Width           =   1755
      End
      Begin VB.Label lblIntitulé 
         Caption         =   "Intitulé"
         Height          =   255
         Left            =   4800
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label libCompteOrigine 
         Caption         =   "---"
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label libCompteFusion 
         Caption         =   "---"
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label lblCompteFusion 
         Caption         =   "Compte de fusion"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblCompte 
         Caption         =   "Compte"
         Height          =   255
         Left            =   3360
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblCompteOrigine 
         Caption         =   "Compte à fusionner"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblDevise 
         Caption         =   "Devise"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.ListBox lstDevise 
      Height          =   3180
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3015
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      Top             =   0
      Width           =   2500
   End
   Begin VB.Menu mnuDevise 
      Caption         =   "Devise"
      Visible         =   0   'False
      Begin VB.Menu mnuDeviseDisplay 
         Caption         =   "Détail Devise"
      End
   End
   Begin VB.Menu mnuEchellesFusion 
      Caption         =   "EchellesFusion"
      Visible         =   0   'False
      Begin VB.Menu mnuEchellesFusionAddNew 
         Caption         =   "Ajouter un Compte"
      End
      Begin VB.Menu mnuEchellesFusionDelete 
         Caption         =   "Supprimer un Compte"
      End
      Begin VB.Menu mnuEchellesFusionPrint 
         Caption         =   "Impression"
      End
   End
End
Attribute VB_Name = "frmEchellesFusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim EchellesFusionAut As typeAuthorization
Dim X As String, X1 As String, I As Integer
Dim Msg As String, valX As String, V As Variant

Dim recEchellesFusion As typeEchellesFusion
Dim currentMethod As String, currentAMJ As String
Dim fgEchellesFusion_FormatString As String, fgEchellesFusion_K As Integer
Dim fgEchellesFusion_BackColorFixed As Long, fgEchellesFusion_BackColor As Long
Dim blnAddNew As Boolean
Dim dblX As Double
Dim CV As typeCV
Dim recCompte As typeCompte
Dim AMJDéfaut As String
'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
cmdControl
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
currentActiveControl_Name = C.Name
End Sub
'---------------------------------------------------------
Public Sub lstDevise_Display()
'---------------------------------------------------------
lstDevise.Visible = True
Call LstDictio(889, lstDevise)
End Sub

'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub



Private Sub cmdContext_Click()
'Select Case cmdContext.Caption
'    Case Is = constcmdRechercher
'    Case Is = constcmdAbandonner: cmdContext_Quit
'End Select

End Sub

'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
cmdPrintX ""
End Sub

Private Sub txtDeviseOrigine_Control()
Dim X As String, V As Variant
X = Trim(txtDeviseOrigine)
If X = "" Then Call lstErr_AddItem(lstErr, txtDeviseOrigine, "? Devise 1"): Exit Sub
If Not IsNumeric(X) Then
    CV.DeviseIso = X: V = CV_Attribut(CV)
Else
    CV.DeviseN = Format$(Val(X), "000"): V = CV_AttributN(CV)
End If
If Not IsNull(V) Then Call lstErr_AddItem(lstErr, txtDeviseOrigine, V): Exit Sub
If Not CV.EuroIn Then Call lstErr_AddItem(lstErr, txtDeviseOrigine, "uniquement devise IN"): Exit Sub
recEchellesFusion.DeviseOrigine = CV.DeviseN
txtDeviseOrigine = CV.DeviseIso
If Trim(txtCompteOrigine) <> "" Then txtCompteOrigine_Control
End Sub





Private Sub fgEchellesFusion_Click()
lstErr.Clear
fgEchellesFusion_K = fgEchellesFusion.Row * fgEchellesFusion.Cols
If fgEchellesFusion.Row > 0 Then lstDevise_Scan mId$(Trim(fgEchellesFusion.TextArray(1 + fgEchellesFusion_K)), 1, 3)
'fgEchellesFusion.Col = 1: fgEchellesFusion.CellBackColor = focusUsr.BackColor
Me.PopupMenu mnuEchellesFusion, vbPopupMenuRightButton
'fgEchellesFusion.Col = 1: fgEchellesFusion.CellBackColor = fgEchellesFusion_BackColor
End Sub

Private Sub fgEchellesFusion_GotFocus()
fgEchellesFusion.BackColorFixed = focusUsr.BackColor
fgEchellesFusion.BackColor = fgEchellesFusion_BackColor
End Sub


Private Sub fgEchellesFusion_LostFocus()
fgEchellesFusion.BackColorFixed = fgEchellesFusion_BackColorFixed
'fgEchellesFusion.BackColor = vbWindowBackground
End Sub


'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub


'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
Dim Amj As String, X As String
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
Call BiaPgmAut_Init("ECH_FUSION", EchellesFusionAut)
currentAMJ = DSys: Amj = DSys
AMJDéfaut = dateElp("FinDeMoisP", -1, currentAMJ)
cmdClear
recEchellesFusion_Init recEchellesFusion
lstDevise_Display
fgEchellesFusion_FormatString = fgEchellesFusion.FormatString
fgEchellesFusion_BackColorFixed = fgEchellesFusion.BackColorFixed
fgEchellesFusion_BackColor = fgEchellesFusion.BackColor
If EchellesFusionAut.Consulter Then
    blnAddNew = True
    fgEchellesFusion_Load Amj, blnAddNew
End If


End Sub



'---------------------------------------------------------
Public Sub cmdClear()
'---------------------------------------------------------
cmdReset
fgEchellesFusion.Enabled = True: fgEchellesFusion.Clear: fgEchellesFusion.Rows = 1
fraEchellesFusion.Visible = False
Call lstErr_Clear(lstErr, fgEchellesFusion, "choisir un compte 'click'")
lstDevise.Enabled = False: lstDevise.BackColor = vbWindowBackground
fraEchellesFusion_Clear
End Sub




'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
arrTag_Set False
lstErrClear = True
blnMsgBox_Quit = False
usrColor_Set
CV_Init CV

End Sub

'---------------------------------------------------------
Public Sub cmdValidation()
'---------------------------------------------------------
cmdControl
lstErrClear = False
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
Msg = Space$(100)

Mid$(Msg, 1, 12) = "frm         "
Mid$(Msg, 13, 12) = Me.Name
Mid$(Msg, 25, 10) = Space$(10)
Mid$(Msg, 35, 10) = X
Mid$(Msg, 45, 6) = "      "
Msg_Monitor Msg
End Sub

Public Sub cmdControl()
Dim I As Integer, xobjControl As Boolean
If lstErrClear Then lstErr.Clear

For I = 0 To arrTagNb
    If arrTag(I) Then
        xobjControl = False
        arrTag(I) = False
        X = Format(I, "000")
        For Each xobj In Me.Controls
            If X = xobj.Tag Then
                Select Case xobj.Name
                    Case "txtDeviseOrigine": txtDeviseOrigine_Control: xobjControl = True
                    Case "txtDeviseFusion": txtDeviseFusion_Control: xobjControl = True
                    Case "txtCompteOrigine": txtCompteOrigine_Control: xobjControl = True
                    Case "txtCompteFusion": txtCompteFusion_Control: xobjControl = True
                    Case "txtAmjDébut": txtAmjdébut_Control: xobjControl = True
                End Select
            End If
            If xobjControl Then Exit For
        Next xobj
    End If
Next I

lstErrClear = True
End Sub


Private Sub fraEchellesFusion_cmdOk_Click()
fraEchellesFusion_Control
If lstErr.ListCount > 0 Then Exit Sub

blnMsgBox_Quit = True

Select Case currentMethod
    Case constAddNew: fgEchellesFusion_AddNew
'    Case constUpdate: fgEchellesFusion_Update
'    Case constDelete: fgEchellesFusion_Delete
End Select
If lstErr.ListCount > 0 Then Exit Sub
fgEchellesFusion_Sort
fraEchellesFusion_Exit
End Sub

Private Sub fraEchellesFusion_cmdOk_GotFocus()
cmdControl
End Sub

Private Sub fraEchellesFusion_cmdQuit_Click()
blnAddNew = False
fraEchellesFusion_Exit
End Sub



Private Sub lblCompteOrigine_Click()
'X = Space$(100)
'Mid$(X, 1, 12) = "frmCompte   "
'Mid$(X, 13, 12) = "FRMECH_FUSIO"
'Mid$(X, 25, 10) = Space$(10)
'Mid$(X, 38, 11) = Trim(txtCompteOrigine)
'Mid$(X, 35, 3) = recEchellesFusion.DeviseOrigine
'Msg_Monitor X

End Sub


Private Sub mnuEchellesFusionPrint_Click()
cmdPrintX ""

End Sub

Private Sub txtAmjDébut_GotFocus()
Call txt_GotFocus(txtAmjDébut)

End Sub


Private Sub txtAmjDébut_LostFocus()
Call txt_LostFocus(txtAmjDébut)
End Sub


Private Sub txtCompteOrigine_GotFocus()
Call txt_GotFocus(txtCompteOrigine)
End Sub


Private Sub txtCompteOrigine_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub txtCompteOrigine_LostFocus()
Call txt_LostFocus(txtCompteOrigine)
txtCompteFusion = Trim(txtCompteOrigine)
End Sub

Private Sub txtDeviseOrigine_GotFocus()
Call txt_GotFocus(txtDeviseOrigine)
lstDevise.Enabled = True: lstDevise.BackColor = lstUsr.BackColor
End Sub

Private Sub txtDeviseOrigine_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtDeviseOrigine_LostFocus()
Call txt_LostFocus(txtDeviseOrigine)
lstDevise.Enabled = False: lstDevise.BackColor = vbWindowBackground
End Sub

Private Sub txtDeviseFusion_GotFocus()
Call txt_GotFocus(txtDeviseFusion)
lstDevise.Enabled = True: lstDevise.BackColor = lstUsr.BackColor
End Sub

Private Sub txtDeviseFusion_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtDeviseFusion_LostFocus()
Call txt_LostFocus(txtDeviseFusion)
lstDevise.Enabled = False: lstDevise.BackColor = vbWindowBackground
End Sub

Private Sub txtCompteFusion_GotFocus()
Call txt_GotFocus(txtCompteFusion)
End Sub

Private Sub txtCompteFusion_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub txtCompteFusion_LostFocus()
Call txt_LostFocus(txtCompteFusion)
End Sub


Private Sub lstDevise_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case currentActiveControl_Name
    Case "txtDeviseOrigine": txtDeviseOrigine = mId$(lstDevise.Text, 1, 3): txtDeviseOrigine.SetFocus
    Case "txtDeviseFusion": txtDeviseFusion = mId$(lstDevise.Text, 1, 3): txtDeviseFusion.SetFocus
End Select
SendKeys "{TAB}"

End Sub

Private Sub mnuEchellesFusionAddNew_Click()
blnAddNew = True
fgEchellesFusion.Enabled = False
fraEchellesFusion_Clear
fraEchellesFusion.Enabled = True: fraEchellesFusion.Visible = True
currentMethod = constAddNew
Call lstErr_Clear(lstErr, txtCompteOrigine, "Nouveau Compte")
recEchellesFusion_Init recEchellesFusion
fraEchellesFusion_cmdOk.Visible = EchellesFusionAut.Valider
txtDeviseOrigine_GotFocus
txtDeviseOrigine.SetFocus
End Sub

Public Function txtCompte_Check(lDevise As String, lCompte As String)

srvCompte.recCompteInit recCompte
recCompte.Method = "SeekL1"
recCompte.Société = SocId$
recCompte.Agence = SocAgence$
recCompte.Devise = lDevise
recCompte.Numéro = lCompte
V = srvCompteFind(recCompte)
''''If Not IsNull(V) Then MsgBox "Compte inexistant : " & recCompte.Devise & "." & recCompte.Numéro, vbCritical, " LrAttribut"
txtCompte_Check = V
End Function

Private Sub mnuEchellesFusionDelete_Click()
fgEchellesFusion_Scan
If lstErr.ListCount > 0 Then Exit Sub
'If recEchellesFusion.Method = constAddNew Then
'    currentMethod = constIgnore
'    fgEchellesFusion_Delete
'Else
    fraEchellesFusion.Enabled = False: fraEchellesFusion.Visible = True
    currentMethod = constDelete
    Call lstErr_Clear(lstErr, txtCompteFusion, "Suppression ligne")
    X = MsgBox("Voulez-vous réellement supprimer cette ligne ?", vbYesNo + vbQuestion + vbDefaultButton2, "ancienne ligne")
    If X = vbYes Then fgEchellesFusion_Delete
    fraEchellesFusion.Visible = False
    fgEchellesFusion.Enabled = True
    fgEchellesFusion.SetFocus
'End If
End Sub

Private Sub fgEchellesFusion_Load(Amj As String, blnZ As Boolean)
Dim blnValidation As Boolean, blnSaisie As Boolean, X As String
lstDevise.Enabled = False: lstDevise.BackColor = vbWindowBackground
srvEchellesFusion_Load Amj
blnValidation = False: blnSaisie = False
For arrEchellesFusionIndex = 1 To arrEchellesFusionNb
    If blnZ Then
        arrEchellesFusion(arrEchellesFusionIndex).Method = ""
    Else
        arrEchellesFusion(arrEchellesFusionIndex).Method = constAddNew
     
    End If
    
Next arrEchellesFusionIndex
If blnValidation And Not blnSaisie Then
    mnuEchellesFusionAddNew.Enabled = False
    mnuEchellesFusionDelete.Enabled = False
Else
    mnuEchellesFusionAddNew.Enabled = True
    mnuEchellesFusionDelete.Enabled = True
End If

fgEchellesFusion_Display
Call lstErr_Clear(lstErr, fgEchellesFusion, X)

End Sub

Private Sub mnuEchellesFusionUpdate_Click()
fgEchellesFusion_Scan
If lstErr.ListCount > 0 Then Exit Sub
fgEchellesFusion.Enabled = False
txtDeviseOrigine.Enabled = False: txtDeviseFusion.Enabled = False
fraEchellesFusion.Enabled = True: fraEchellesFusion.Visible = True
currentMethod = constUpdate
Call lstErr_Clear(lstErr, txtCompteFusion, "Modification Compte")
lastActiveControl_Name = "fraEchellesFusion_txtVentePrivilégié"
txtDeviseOrigine.SetFocus
End Sub

Private Sub mnuDeviseDisplay_Click()
If lstDevise.ListIndex >= 0 Then
    Set XListBox = frmEchellesFusion.lstDevise
    frmDevise.Show vbModal
End If
End Sub

Private Sub mnuDevisePrint_Click()
Dim Msg As String

Msg = Format$(1, "000000") & Format$(999, "000000")

prtDeviseX Msg

End Sub



Public Sub fgEchellesFusion_Display()
fgEchellesFusion.Rows = 1
fgEchellesFusion.Clear
fgEchellesFusion.FormatString = fgEchellesFusion_FormatString
fgEchellesFusion.Enabled = True
For arrEchellesFusionIndex = 1 To arrEchellesFusionNb
    If arrEchellesFusion(arrEchellesFusionIndex).Method <> constDelete _
    And arrEchellesFusion(arrEchellesFusionIndex).Method <> constIgnore Then
        fgEchellesFusion.Rows = fgEchellesFusion.Rows + 1
        fgEchellesFusion.Row = fgEchellesFusion.Rows - 1
        fgEchellesFusion_DisplayItem
    End If
Next arrEchellesFusionIndex
If fgEchellesFusion.Rows > 1 Then fgEchellesFusion_Sort

End Sub

Public Sub fgEchellesFusion_DisplayItem()
fgEchellesFusion_K = (fgEchellesFusion.Row) * fgEchellesFusion.Cols
fgEchellesFusion.TextArray(0 + fgEchellesFusion_K) = Format$(arrEchellesFusion(arrEchellesFusionIndex).DeviseFusion, "@@@") & "_" & Compte_Imp(arrEchellesFusion(arrEchellesFusionIndex).CompteFusion)
fgEchellesFusion.TextArray(1 + fgEchellesFusion_K) = Format$(arrEchellesFusion(arrEchellesFusionIndex).DeviseOrigine, "@@@") & "_" & Compte_Imp(arrEchellesFusion(arrEchellesFusionIndex).CompteOrigine)
fgEchellesFusion.TextArray(2 + fgEchellesFusion_K) = dateImp(arrEchellesFusion(arrEchellesFusionIndex).AmjDébut)
Select Case arrEchellesFusion(arrEchellesFusionIndex).Method
    Case constAddNew: fgEchellesFusion.TextArray(3 + fgEchellesFusion_K) = "Créé"
    Case constUpdate: fgEchellesFusion.TextArray(3 + fgEchellesFusion_K) = "Modifié"
End Select
fgEchellesFusion.TextArray(4 + fgEchellesFusion_K) = Format$(arrEchellesFusionIndex, "###0")
End Sub

Public Sub fgEchellesFusion_AddNew()
recEchellesFusion.Method = constAddNew
srvEchellesFusion_Update recEchellesFusion
Call arrEchellesFusion_AddItem(recEchellesFusion)
arrEchellesFusionIndex = arrEchellesFusionNb
fgEchellesFusion.Rows = fgEchellesFusion.Rows + 1
fgEchellesFusion.Row = fgEchellesFusion.Rows - 1
fgEchellesFusion_DisplayItem
End Sub

Public Sub fgEchellesFusion_Update()
If recEchellesFusion.Method <> constAddNew Then recEchellesFusion.Method = constUpdate
arrEchellesFusion(arrEchellesFusionIndex) = recEchellesFusion
fgEchellesFusion_DisplayItem
End Sub

Public Sub fgEchellesFusion_Delete()
recEchellesFusion.Method = currentMethod
arrEchellesFusion(arrEchellesFusionIndex) = recEchellesFusion
srvEchellesFusion_Update recEchellesFusion
fgEchellesFusion_Display
End Sub

Public Sub fraEchellesFusion_Clear()
lstErr.Clear
usrColor_Set

txtDeviseOrigine = ""
txtCompteOrigine = ""
txtCompteFusion = ""
txtDeviseFusion = "EUR": txtDeviseFusion.Enabled = False
libCompteOrigine = "": libCompteFusion = ""
fraEchellesFusion.Enabled = True
txtDeviseOrigine.Enabled = True
lastActiveControl_Name = "txtAmjDébut"
txtAmjDébut = dateImp(AMJDéfaut)
End Sub

Public Sub fraEchellesFusion_Enabled(ByVal bln As Boolean)
fgEchellesFusion.Enabled = bln
fraEchellesFusion_cmdOk.Visible = bln
End Sub

Public Sub cmdContext_Quit()
If fraEchellesFusion.Visible Then
    fraEchellesFusion_cmdQuit_Click
Else
    If blnMsgBox_Quit Then
        X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
     Else
        X = vbYes
     End If
     If X = vbYes Then Unload Me
End If

End Sub

Public Sub cmdContext_Return()
If fraEchellesFusion.Enabled Then
    If ActiveControl.Name = lastActiveControl_Name Then
        fraEchellesFusion_cmdOk_Click
    Else
        SendKeys "{TAB}"
    End If
Else
    SendKeys "{TAB}"
End If

End Sub

Public Sub fgEchellesFusion_Scan()
Dim K As Integer
fgEchellesFusion_K = fgEchellesFusion.Row * fgEchellesFusion.Cols
K = Val(Trim(fgEchellesFusion.TextArray(4 + fgEchellesFusion_K)))
recEchellesFusion.DeviseOrigine = arrEchellesFusion(K).DeviseOrigine
recEchellesFusion.CompteOrigine = arrEchellesFusion(K).CompteOrigine
If arrEchellesFusion_ScanDeviseOrigineCompteOrigine(recEchellesFusion) > 0 Then
    recEchellesFusion = arrEchellesFusion(arrEchellesFusionIndex)
    fraEchellesFusion_DisplayItem
Else
    Call lstErr_AddItem(lstErr, fgEchellesFusion, "?Erreur fgEchellesFusion_Scan")
End If
End Sub

Public Sub fraEchellesFusion_DisplayItem()
usrColor_Set

txtDeviseOrigine = recEchellesFusion.DeviseOrigine
txtCompteOrigine = recEchellesFusion.CompteOrigine
txtDeviseFusion = recEchellesFusion.DeviseFusion
txtCompteFusion = recEchellesFusion.CompteFusion

txtDeviseOrigine_Control
txtCompteOrigine_Control
txtDeviseFusion_Control
txtCompteFusion_Control
'txtAmjDébut.Text = dateImp(recEchellesFusion.AmjDébut)
fraEchellesFusion_cmdOk.Visible = False 'EchellesFusionAut.Valider

End Sub

Public Sub fraEchellesFusion_Exit()
lstDevise.Enabled = False
If blnAddNew Then
    mnuEchellesFusionAddNew_Click
Else
    fraEchellesFusion.Visible = False
    fgEchellesFusion.Enabled = True
    Call lstErr_Clear(lstErr, fgEchellesFusion, "choisir une compte 'click'")
    fgEchellesFusion.SetFocus
End If
End Sub

Public Sub fgEchellesFusion_Sort()
fgEchellesFusion.Row = 1
fgEchellesFusion.RowSel = 1 'fgLRAttribut.Rows - 1

fgEchellesFusion.Col = 0
fgEchellesFusion.ColSel = 1
fgEchellesFusion.Sort = flexSortStringAscending

End Sub

Public Sub txtCompteFusion_Control()
X = num_Control(txtCompteFusion, valX, 11, 0)
If X <> "" Then
    Call lstErr_AddItem(lstErr, txtCompteFusion, X)
Else
    txtDeviseFusion_Control
    recEchellesFusion.CompteFusion = valX
    txtCompteFusion = Compte_Display(recEchellesFusion.CompteFusion)
    V = txtCompte_Check(recEchellesFusion.DeviseFusion, recEchellesFusion.CompteFusion)
    If IsNull(V) Then
         libCompteFusion = recCompte.Intitulé
   Else
        Call lstErr_AddItem(lstErr, txtCompteFusion, X)
        libCompteFusion = "? compte"
    End If
End If
End Sub

Public Sub txtDeviseFusion_Control()
Dim X As String, V As Variant
X = Trim(txtDeviseFusion)
If X = "" Then Call lstErr_AddItem(lstErr, txtDeviseFusion, "? Devise 1"): Exit Sub
If Not IsNumeric(X) Then
    CV.DeviseIso = X: V = CV_Attribut(CV)
Else
    CV.DeviseN = Format$(Val(X), "000"): V = CV_AttributN(CV)
End If
If Not IsNull(V) Then Call lstErr_AddItem(lstErr, txtDeviseFusion, V): Exit Sub
'''If Not CV.EuroIn Then Call lstErr_AddItem(lstErr, txtDeviseFusion, "uniquement devise IN"): Exit Sub
recEchellesFusion.DeviseFusion = CV.DeviseN
txtDeviseFusion = CV.DeviseIso
End Sub

Public Sub fraEchellesFusion_Control()
lstErr.Clear: arrTag_Set False

txtDeviseOrigine_Control
txtCompteOrigine_Control
txtDeviseFusion_Control
txtCompteFusion_Control
txtAmjdébut_Control
If recEchellesFusion.DeviseOrigine = recEchellesFusion.DeviseFusion Then Call lstErr_AddItem(lstErr, txtDeviseOrigine, "? Devise1 = Devise2")
If currentMethod = constAddNew Then
    If arrEchellesFusion_ScanDeviseOrigineCompteOrigine(recEchellesFusion) > 0 Then
        Call lstErr_AddItem(lstErr, txtDeviseOrigine, "? Existe déjà")
    End If
End If
End Sub

Public Sub lstDevise_Scan(strdev As String)
Dim K As Integer
For K = 0 To lstDevise.ListCount - 1
    lstDevise.ListIndex = K
    If mId$(lstDevise.Text, 1, 3) = strdev Then Exit For
Next K
End Sub

Public Sub txtCompteOrigine_Control()
X = num_Control(txtCompteOrigine, valX, 11, 0)
If X <> "" Then
    Call lstErr_AddItem(lstErr, txtCompteOrigine, X)
Else
    recEchellesFusion.CompteOrigine = valX
    txtCompteOrigine = Compte_Display(recEchellesFusion.CompteOrigine)
    V = txtCompte_Check(recEchellesFusion.DeviseOrigine, recEchellesFusion.CompteOrigine)
    If IsNull(V) Then
        libCompteOrigine = recCompte.Intitulé
        recEchellesFusion.CompteFusion = recEchellesFusion.CompteOrigine
        If Trim(txtCompteFusion) = "" Then txtCompteFusion = txtCompteOrigine
        txtCompteFusion_Control
   Else
        Call lstErr_AddItem(lstErr, txtCompteOrigine, V)
        libCompteOrigine = "? compte"
    End If
End If

End Sub

Public Sub cmdPrintX(Msg As String)
Dim I As Integer, K As Integer

ReDim sortEchellesfusion(arrEchellesFusionNbMax)
For I = 1 To fgEchellesFusion.Rows - 1
    fgEchellesFusion_K = I * fgEchellesFusion.Cols
    sortEchellesfusion(I) = Val(Trim(fgEchellesFusion.TextArray(4 + fgEchellesFusion_K)))
Next I
X = Format$(1, "000000") & Format$(I - 1, "000000")

prtEchellesFusionX X, Msg

End Sub

Public Sub txtAmjdébut_Control()
Dim X As String

X = dateCtl(txtAmjDébut.Text)
If Not IsNumeric(X) Then
    Call lstErr_AddItem(lstErr, txtAmjDébut, X)
    txtAmjDébut.ForeColor = errUsr.ForeColor
Else
    recEchellesFusion.AmjDébut = mId$(X, 1, 8)
    If recEchellesFusion.AmjDébut < AMJDéfaut Then
        Call lstErr_AddItem(lstErr, txtAmjDébut, "? la date doit être >= " & AMJDéfaut)
        txtAmjDébut.ForeColor = errUsr.ForeColor
    Else
        If recEchellesFusion.AmjDébut <> dateFinDeMois(recEchellesFusion.AmjDébut) Then
            Call lstErr_AddItem(lstErr, txtAmjDébut, "? la date n'est pas une fin de mois ")
            txtAmjDébut.ForeColor = errUsr.ForeColor
        Else
           If recEchellesFusion.AmjDébut <> AMJDéfaut Then Call MsgBox("Vérifiez la date", vbExclamation, "Date différente de la dernière fin de mois")

            If recEchellesFusion.AmjDébut <> "00000000" Then
                txtAmjDébut.Text = dateImp(recEchellesFusion.AmjDébut)
            End If
        End If
    End If
End If

End Sub
