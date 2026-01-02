VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCptRelevé 
   AutoRedraw      =   -1  'True
   Caption         =   "Compte :Relevé périodique"
   ClientHeight    =   5025
   ClientLeft      =   75
   ClientTop       =   345
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5025
   ScaleWidth      =   9330
   Begin VB.Frame fraSort 
      Caption         =   "Tri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   3240
      TabIndex        =   10
      Top             =   900
      Width           =   2700
      Begin VB.OptionButton optSortGestionnaire 
         Caption         =   "par gestionnaire/compte"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   2295
      End
      Begin VB.OptionButton optSortCourrier 
         Caption         =   "par code courrier/compte"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   2415
      End
      Begin VB.OptionButton optSortCompte 
         Caption         =   "par compte"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame fraPériodicité 
      Caption         =   "Relevé"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   0
      TabIndex        =   3
      Top             =   900
      Width           =   2700
      Begin VB.OptionButton optRelevéDécadaire 
         Caption         =   "Décadaire"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optRelevéQuotidien 
         Caption         =   "Quotidien"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Value           =   -1  'True
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtAmjMin 
         Height          =   300
         Left            =   1320
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
      Begin MSMask.MaskEdBox txtAmjMax 
         Height          =   300
         Left            =   1320
         TabIndex        =   7
         Top             =   2160
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   14
         Mask            =   "## - ## - ####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblAmjMax 
         Caption         =   "au"
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
         Left            =   480
         TabIndex        =   6
         Top             =   2160
         Width           =   435
      End
      Begin VB.Label lblAmjMin 
         Caption         =   "Mvt du "
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
         Left            =   360
         TabIndex        =   4
         Top             =   1680
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   8950
      Picture         =   "CptRelevé.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   400
   End
   Begin VB.ListBox lstCptrelevé 
      Height          =   2985
      Left            =   6360
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   900
      Width           =   2700
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ErrMsg"
      ForeColor       =   &H000000FF&
      Height          =   825
      Left            =   6360
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmCptRelevé"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean

Private recCptRelevé As typeCptRelevé
Private CptLstMethod As String * 12, CptMvtMethod As String
Dim valAmjMin As String * 8
Dim valAmjMax As String * 8
Private recCptMvt As typeCptMvt

Dim nbExtrait As Integer
Private Sub optRelevéDécadaire_Click()
valAmjMax = DSys
valAmjMin = DSys
If Val(mId$(valAmjMax, 7, 2)) > 20 Then
    Mid$(valAmjMax, 7, 2) = "20"
    Mid$(valAmjMin, 7, 2) = "11"
Else
    If Val(mId$(valAmjMax, 7, 2)) > 10 Then
        Mid$(valAmjMax, 7, 2) = "10"
        Mid$(valAmjMin, 7, 2) = "01"
    Else
        valAmjMax = dateElp("FinDeMoisP", 0, valAmjMax)
        valAmjMin = valAmjMax
        Mid$(valAmjMin, 7, 2) = "20"
    End If
End If

txtAmjMax = dateImp(valAmjMax)
txtAmjMin = dateImp(valAmjMin)

End Sub

Private Sub optRelevéQuotidien_Click()
valAmjMax = dateElp("Ouvré", -1, DSys)
txtAmjMax = dateImp(valAmjMax)
valAmjMin = valAmjMax
txtAmjMin = dateImp(valAmjMin)

End Sub

'-------------------------------------------------'
Private Sub txtAmjMax_GotFocus()
'-------------------------------------------------'
txtAmjMax.BackColor = focusUsr.BackColor

End Sub

'-------------------------------------------------'
Private Sub txtAmjMax_LostFocus()
'-------------------------------------------------'

Dim X As String
txtAmjMax.BackColor = txtUsr.BackColor
txtAmjMax.ForeColor = txtUsr.ForeColor

X = dateCtl(txtAmjMax.Text)
If Not IsNumeric(X) Then
    lblErrMsg = X
    txtAmjMax.ForeColor = errUsr.ForeColor
Else
    valAmjMax = mId$(X, 1, 8)
    If valAmjMax <> "00000000" Then
        txtAmjMax.Text = dateImp(valAmjMax)
    Else
        valAmjMax = DSys
        txtAmjMax = dateImp(DSys)
    End If
End If

End Sub

'-------------------------------------------------'
Private Sub txtAmjMin_GotFocus()
'-------------------------------------------------'
txtAmjMin.BackColor = focusUsr.BackColor

End Sub

'-------------------------------------------------'
Private Sub txtAmjMin_LostFocus()
'-------------------------------------------------'
Dim X As String

txtAmjMin.BackColor = txtUsr.BackColor
txtAmjMin.ForeColor = txtUsr.ForeColor

X = dateCtl(txtAmjMin.Text)
If Not IsNumeric(X) Then
    lblErrMsg = X
    txtAmjMin.ForeColor = errUsr.ForeColor
Else
    valAmjMin = mId$(X, 1, 8)
    If valAmjMin <> "00000000" Then
        txtAmjMin.Text = dateImp(valAmjMin)
    End If
End If
End Sub

'---------------------------------------------------------
Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------
End Sub



'---------------------------------------------------------
Private Sub CptRelevé_load()
'---------------------------------------------------------

srvCptRelevé.Init recCptRelevé

recCptRelevé.Method = "SnapKE"
recCptRelevé.Société = SocId$
recCptRelevé.Agence = SocAgence$
recCptRelevé.Devise = "000"
recCptRelevé.Numéro = "00000000000"
If optRelevéDécadaire Then
    recCptRelevé.ExtraitPériodicité = "3"
Else
    recCptRelevé.ExtraitPériodicité = "1"
End If

arrCptRelevé(0) = recCptRelevé
arrCptRelevé(0).Devise = "999"
arrCptRelevé(0).Numéro = "99999999999"

arrCptRelevéNb = 0
arrCptRelevéIndex = 0
arrCptRelevésuite = True

Do Until Not arrCptRelevésuite
    srvCptRelevé.Monitor recCptRelevé
    recCptRelevé = arrCptRelevé(arrCptRelevéNb)
    recCptRelevé.Method = "SnapKE+"

Loop

End Sub




'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
lblErrMsg = ""
If IsNull(valDate) Then
    lblErrMsg = "Recherche des comptes"
    CptRelevé_load
    CptRelevé_Display
    lblErrMsg = "Impression des extraits"
    CptRelevé_Print
End If
End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 27: cmdQuit_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
    Case Is = 13: cmdReturn_Click
End Select
'lblErrMsg.Visible = False
End Sub

'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me
End Sub

'---------------------------------------------------------
Private Sub cmdReturn_Click()
'---------------------------------------------------------
cmdPrint_Click
End Sub



'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------

Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
arrCptRelevéNbMax = 1: arrCptRelevéNb = 0: ReDim arrCptRelevé(1)
arrCptMvtNbMax = 1: arrCptMvtNb = 0: ReDim arrCptMvt(1)

lblErrMsg = "Cliquer sur le bouton 'Imprimante'" & Chr$(13) & "pour lancer le traitement"
lblErrMsg.ForeColor = errUsr.ForeColor
optRelevéQuotidien_Click
End Sub

'---------------------------------------------------------
Private Sub CptMvt_Load()
'---------------------------------------------------------
Dim Msg As String
arrCptMvt(0) = recCptMvt

    CptMvtMethod = "SnapJA"
    recCptMvt.Method = CptMvtMethod
    recCptMvt.AmjTraitement = valAmjMin
    recCptMvt.Pièce = 0
    recCptMvt.Ligne = 0
    arrCptMvt(0).Method = CptMvtMethod
    arrCptMvt(0).AmjTraitement = valAmjMax
    arrCptMvt(0).Pièce = "999999999"
    arrCptMvt(0).Ligne = "9999"
    
arrCptMvtSuite = True
arrCptMvtNb = 0

Do Until Not arrCptMvtSuite
    srvCptMvtMon recCptMvt
    recCptMvt = arrCptMvt(arrCptMvtNb)
    recCptMvt.Method = CptMvtMethod & "+"
Loop

If arrCptMvtNb > 0 Then
    arrCompteIndex = 1: arrCompteNb = 1
    arrCompte(1).Société = arrCptMvt(1).Société
    arrCompte(1).Agence = arrCptMvt(1).Agence
    arrCompte(1).Devise = arrCptMvt(1).Devise
    arrCompte(1).Numéro = arrCptMvt(1).Compte
    arrCompte(1).BiaTyp = "000"
    arrCompte(1).BiaNum = "00"
    arrCompte(1).NuméroAncien = "00000"
    Msg = Format$(1, "000000") & Format$(arrCptMvtNb, "000000") & " E " & valAmjMin & valAmjMax
    CV_Init CV_X1
    CV_X1.DeviseN = Format$(arrCptMvt(1).Devise, "000")
    Call CV_AttributN(CV_X1)
    If CV_X1.EuroIn Then Mid$(Msg, 15, 1) = "E"
    prtCptMvt_Monitor Msg
   lblErrMsg = arrCptInfo(0).Intitulé
    nbExtrait = nbExtrait + 1
End If

End Sub





'-------------------------------------------------'
Public Function valDate()
'-------------------------------------------------'

valDate = "?"

txtAmjMin_LostFocus
txtAmjMax_LostFocus

If valAmjMin > valAmjMax Then
    lblErrMsg = "date début > date fin"
    Exit Function
'!!!!!!!!!!!
End If
valDate = Null

End Function


Public Sub CptRelevé_Display()
Dim I As Integer, strSort As String * 3
lstCptrelevé.Clear

For I = 1 To arrCptRelevéNb
    If optSortGestionnaire Then
        strSort = Format$(arrCptRelevé(I).Gestionnaire, "@@@")
    Else
        If optSortCourrier Then
            strSort = Format$(arrCptRelevé(I).Courrier, "@@@")
        Else
            strSort = "   "
        End If
    End If
    lstCptrelevé.AddItem strSort & "." & arrCptRelevé(I).Numéro & "." & arrCptRelevé(I).Devise
    
Next I

End Sub

Public Sub CptRelevé_Print()
Dim I As Integer

nbExtrait = 0
For I = 0 To lstCptrelevé.ListCount - 1
    lstCptrelevé.ListIndex = I
    recCptMvtInit recCptMvt
         
    recCptMvt.Société = SocId$
    recCptMvt.Agence = SocAgence$
    recCptMvt.Devise = mId$(lstCptrelevé, 17, 3)
    recCptMvt.Compte = mId$(lstCptrelevé, 5, 11)
    CptMvt_Load
Next I
lblErrMsg = Format(nbExtrait, "#####") & " extraits imprimés"
End Sub
