VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCompteGafi 
   Caption         =   "Compte : surveillance"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9180
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5040
      TabIndex        =   16
      Top             =   0
      Width           =   3585
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
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   0
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Choix d'un état"
      TabPicture(0)   =   "CompteGafi.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBalance"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Sélection des comptes"
      TabPicture(1)   =   "CompteGafi.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraOptions"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraOptions 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   8775
         Begin VB.Frame fraSelect 
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
            Height          =   5175
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   8535
            Begin VB.TextBox txtCr 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3240
               TabIndex        =   33
               Top             =   4680
               Width           =   1695
            End
            Begin VB.TextBox txtDb 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3240
               TabIndex        =   32
               Top             =   4200
               Width           =   1695
            End
            Begin VB.CheckBox chkCr 
               Caption         =   "Crédit  >  ******* EUR"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   4680
               Width           =   1935
            End
            Begin VB.CheckBox chkDb 
               Caption         =   "Débit >  ******* EUR"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   4200
               Width           =   2055
            End
            Begin VB.CheckBox chkService 
               Caption         =   "sélectionner le service "
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   3360
               Width           =   2895
            End
            Begin VB.TextBox txtService 
               Height          =   285
               Left            =   3360
               MaxLength       =   3
               TabIndex        =   28
               Top             =   3240
               Width           =   495
            End
            Begin VB.ListBox lstDevise 
               Height          =   2010
               Left            =   5400
               TabIndex        =   19
               Top             =   240
               Width           =   3015
            End
            Begin VB.TextBox txtDeviseCV 
               Height          =   285
               Left            =   3360
               MaxLength       =   3
               TabIndex        =   17
               Text            =   "EUR"
               Top             =   400
               Width           =   495
            End
            Begin VB.CheckBox chkDeviseIn 
               Caption         =   "Uniquement devises In et Euro"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   800
               Width           =   2775
            End
            Begin VB.TextBox txtBiaTyp 
               Height          =   285
               Left            =   3360
               MaxLength       =   3
               TabIndex        =   11
               Top             =   2760
               Width           =   495
            End
            Begin VB.CheckBox chkBiaTyp 
               Caption         =   "sélectionner le type de compte"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   2880
               Width           =   2535
            End
            Begin VB.CheckBox chkCompteMinMax 
               Caption         =   "sélectionner les comptes de"
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   1800
               Width           =   2535
            End
            Begin VB.TextBox txtDevise 
               Height          =   285
               Left            =   3360
               MaxLength       =   3
               TabIndex        =   8
               Top             =   1200
               Width           =   495
            End
            Begin VB.CheckBox chkDevise 
               Caption         =   "Sélectionner la devise"
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   1200
               Width           =   2055
            End
            Begin VB.TextBox txtCompteMax 
               Height          =   285
               Left            =   3360
               MaxLength       =   11
               TabIndex        =   5
               Top             =   2280
               Width           =   1575
            End
            Begin VB.TextBox txtCompteMin 
               Height          =   285
               Left            =   3360
               MaxLength       =   11
               TabIndex        =   4
               Top             =   1800
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "devise de contre-valeur"
               Height          =   255
               Left            =   360
               TabIndex        =   18
               Top             =   400
               Width           =   2415
            End
            Begin VB.Label lblMax 
               Caption         =   "à"
               Height          =   255
               Left            =   2280
               TabIndex        =   6
               Top             =   2400
               Width           =   255
            End
         End
      End
      Begin VB.Frame fraBalance 
         Height          =   5655
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8775
         Begin VB.Frame fraPériode 
            Height          =   3375
            Left            =   240
            TabIndex        =   22
            Top             =   2040
            Width           =   8295
            Begin VB.CommandButton cmdOk 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Ok"
               Height          =   885
               Left            =   6120
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   2160
               Width           =   1455
            End
            Begin MSComCtl2.DTPicker txtAmjMin 
               Height          =   300
               Left            =   1920
               TabIndex        =   23
               Top             =   600
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
               Format          =   64946179
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtAmjMax 
               Height          =   300
               Left            =   3840
               TabIndex        =   26
               Top             =   600
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
               Format          =   64946179
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   2
            End
            Begin VB.Label libInfo 
               Caption         =   "sont exclus : les mvts DAT  et TC"
               Height          =   375
               Left            =   5640
               TabIndex        =   35
               Top             =   1560
               Width           =   2535
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
               Left            =   3360
               TabIndex        =   25
               Top             =   720
               Width           =   315
            End
            Begin VB.Label lblAmjMin 
               Caption         =   "du"
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
               Left            =   1320
               TabIndex        =   24
               Top             =   720
               Width           =   315
            End
         End
         Begin VB.Frame fraScript 
            Caption         =   "Script"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   3135
            Begin VB.OptionButton optEtatManuel 
               Caption         =   "Manuel"
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   360
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.Frame fraEtat 
            Caption         =   "Etat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1635
            Left            =   3720
            TabIndex        =   12
            Top             =   360
            Width           =   4815
            Begin VB.OptionButton optEtat02 
               Caption         =   "Etat des cumuls (SOBF + ORPA) > 7 000"
               Height          =   255
               Left            =   120
               TabIndex        =   34
               Top             =   720
               Value           =   -1  'True
               Width           =   4455
            End
            Begin VB.OptionButton optEtat01 
               Caption         =   "Etat des mouvements > 150 000"
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   300
               Width           =   4455
            End
         End
      End
   End
End
Attribute VB_Name = "frmCompteGafi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean, blnSetfocus As Boolean, blnControl As Boolean
Dim CompteGafiAut As typeAuthorization
Dim X As String, X1 As String, I As Long
'Dim Msg As String, valX As String, V As Variant
Dim reccptp0 As typeCptP0
Dim recCompte As typeCompte, recRacine As typeRacine

Dim optEtat As String * 1, optSolde As String * 1, optAmj As String * 8, SrvCptP0_Amj As String * 8
Dim blnCompteMinMax As Boolean, selCompteMin As String * 11, selCompteMax As String * 11
Dim blnDevise As Boolean, selDeviseN As String * 3, blnDeviseIn As Boolean, selDeviseCV As String * 3
Dim blnService As Boolean, selService As String * 3
Dim blnBiaTyp As Boolean, selBiaTyp As String * 3
Dim optSortK As String * 1
Dim optEtatSortK As String * 2
Dim mDestinataire As String, mEnTete As String
Dim PrintRupture_Len As Integer

Dim blnExport As Boolean, X137 As String * 137
Dim X1000 As String * 1000
Dim cmdImport_Select_Nb As Long, cmdImport_Nb As Long

Dim blnService_Enabled As Boolean
Dim wL As Long, wPAys As String * 4, wX As String
Dim recdictio As typeDictio

Dim wAmjMin As String * 8, wAmjMax As String * 8, wAmj As String * 8
Dim xAmjMin As String, xAmjMax As String, xAMJ As String
Dim vReturn As Variant
Dim mID14 As String * 14, wMt As Currency, wMtCV As Currency
Dim wDB1 As Currency, wCR1 As Currency, wDB2 As Currency, wCR2 As Currency, wVR4 As Currency
Dim sSD1 As Currency, sCR As Currency, sDB As Currency, sSD2 As Currency
Dim tSD1 As Currency, tCR As Currency, tDB As Currency, tSD2 As Currency
Dim sDev As String * 3, tCompte As String * 11, tIntitulé As String

Dim paramCompteGafi_Cpt_Import As String, paramCompteGafi_Cpt_Export As String, paramCompteGafi_Mvt_Import As String

Dim curX As Currency, curDB As Currency, curCR As Currency
Dim blnDb As Boolean, blnCr As Boolean

Dim meCV1 As typeCV, meCV2 As typeCV, meCV3  As typeCV

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


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
blnSetfocus = False

SrvCptP0_Amj = "00000000"
blnService_Enabled = True
Call BiaPgmAut_Init("Compte_Gafi", CompteGafiAut)

If Not IsNull(param_Init) Then cmdOk.Visible = False
cmdReset

blnSetfocus = True
End Sub


Public Function param_Init()
Dim V
param_Init = Null
recElpTable_Init recElpTable
recElpTable.Id = "Param"
recElpTable.K1 = "ComptaExt"
recElpTable.Method = "Seek="

recElpTable.K2 = "Cpt_Import"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramCompteGafi_Cpt_Import = paramServer(recElpTable.Memo)
'''Call lstErr_Clear(lstErr, cmdContext, "Fichier :" & paramCompteGafi_Cpt_Import)


recElpTable.K2 = "Mvt_Import"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramCompteGafi_Mvt_Import = paramServer(recElpTable.Memo)

Exit Function

Table_Error:
param_Init = V
Exit Function

Memo_Error:
param_Init = "Memo"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "CompteGafi_Param_Init"
Exit Function

End Function


Private Sub chkBiaTyp_Click()
If chkBiaTyp = "1" Then
    txtBiaTyp.Visible = True: If blnSetfocus Then txtBiaTyp.SetFocus
Else
    txtBiaTyp.Visible = False
End If
End Sub

Private Sub chkBiaTyp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkBiaTyp
End Sub


Private Sub chkCompteMinMax_Click()
If chkCompteMinMax = "1" Then
    txtCompteMin.Visible = True: txtCompteMax.Visible = True
    If blnSetfocus Then txtCompteMin.SetFocus
Else
    txtCompteMin.Visible = False: txtCompteMax.Visible = False
End If

End Sub

Private Sub chkCompteMinMax_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkCompteMinMax
End Sub


Private Sub chkDevise_Click()
If chkDevise = "1" Then
    txtDevise.Visible = True: If blnSetfocus Then txtDevise.SetFocus
Else
    txtDevise.Visible = False
End If

End Sub

Private Sub chkDevise_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkDevise
End Sub


Private Sub chkDeviseIn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkDeviseIn
End Sub


Private Sub chkService_Click()
On Error GoTo Exit_Sub
If chkService = "1" Then
    txtService.Visible = True: If blnSetfocus Then txtService.SetFocus
Else
    txtService.Visible = False
End If
Exit_Sub:
End Sub

Private Sub chkService_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkService
End Sub


Public Sub cmdControl()
If Not Me.Enabled Then Exit Sub
Me.Enabled = False

blnControl = False
lstErr.Clear
lstErr.Height = 200

vReturn = DTPicker_Control(txtAmjMin, wAmjMin)
If Not IsNull(vReturn) Then Call lstErr_AddItem(lstErr, txtAmjMin, vReturn): Exit Sub
vReturn = DTPicker_Control(txtAmjMax, wAmjMax)
If Not IsNull(vReturn) Then Call lstErr_AddItem(lstErr, txtAmjMax, vReturn): Exit Sub


blnCompteMinMax = IIf(chkCompteMinMax = "1", True, False)

selCompteMin = Format$(Val(Trim(txtCompteMin)), "00000000000")
selCompteMax = Format$(Val(Trim(txtCompteMax)), "00000000000")

If blnCompteMinMax Then
    If selCompteMin = "00000000000" Then
        Call lstErr_AddItem(lstErr, cmdContext, "? préciser le compte min")
    Else
        If selCompteMax = "00000000000" Then selCompteMax = selCompteMin
    End If
    If selCompteMin > selCompteMax Then Call lstErr_AddItem(lstErr, cmdContext, "? compte min > compte max")

End If

selDeviseCV = Trim(txtDeviseCV)
Call CV_AttributS(selDeviseCV, CV_X2)
selDeviseCV = CV_X2.DeviseIso

blnDeviseIn = IIf(chkDeviseIn = "1", True, False)
blnDevise = IIf(chkDevise = "1", True, False)
selDeviseN = Trim(txtDevise)
'If IsNumeric(selDeviseN) Then
'    CV_X1.DeviseN = Format$(selDeviseN, "000")
'    Call CV_AttributN(CV_X1)
    Call CV_AttributS(selDeviseN, CV_X1)
    selDeviseN = CV_X1.DeviseIso
'End If
If blnDeviseIn And blnDevise Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser 1 devise ou devise In")
If blnDevise Then
    If Trim(txtDevise) = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser la devise")
End If


blnService = IIf(chkService = "1", True, False)
selService = Format$(Trim(txtService), "000")
If blnService Then
    If Trim(txtService) = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le service")
End If


blnBiaTyp = IIf(chkBiaTyp = "1", True, False)
selBiaTyp = Format$(Trim(txtBiaTyp), "000")
If blnBiaTyp Then
    If Trim(txtBiaTyp) = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le Type")
End If

If chkDb = "1" Then
    curDB = -Abs(CCur(Val(txtDb)))
    blnDb = True
Else
    curDB = 0
    blnDb = False
End If

If chkCr = "1" Then
    curCR = Abs(CCur(Val(txtCr)))
    blnCr = True
Else
    curCR = 0
    blnCr = False
End If


ExitSub:

Me.Enabled = True
    
blnControl = True

End Sub
Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdOk_Click()
Call lstErr_Clear(lstErr, cmdOk, "Début du traitement")
If blnControl Then cmdControl
If lstErr.ListCount <> 0 Then Exit Sub

cmdFlux

End Sub

Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
Form_Init
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


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub



Public Sub cmdContext_Return()

End Sub

Public Sub cmdContext_Quit()
Unload Me
End Sub

Private Function cmdImport_MvtP0_Select(Msg As String) As String
Dim wNuméro As String * 11, wDeviseN As String * 3, wAmjOpération As String * 8, wService As String * 3, wCodeOpération As String * 4
Dim X1 As String * 1

cmdImport_MvtP0_Select = ""
wAmjOpération = mId$(Msg, 163, 8)
If wAmjOpération < wAmjMin Then Exit Function
If wAmjOpération > wAmjMax Then Exit Function

' comptes auxiliaires et bilan

wNuméro = mId$(Msg, 10, 11)
If wNuméro < "10000000000" Then Exit Function
If mId$(Msg, 15, 1) = "9" Then Exit Function

' ignorer mvts TC & DAT

wCodeOpération = mId$(Msg, 21, 4)
If wCodeOpération = "G051" Or wCodeOpération = "G052" Or wCodeOpération = "T550" Then Exit Function

If blnCompteMinMax Then
        If wNuméro < selCompteMin Or wNuméro > selCompteMax Then Exit Function
End If

wDeviseN = mId$(Msg, 7, 3)
If CV_X1.DeviseN <> wDeviseN Then
    CV_X1.DeviseN = wDeviseN
    Call CV_AttributN(CV_X1)
End If
 

If blnDeviseIn Then
    If Not CV_X1.EuroIn And CV_X1.DeviseIso <> "EUR" Then Exit Function
End If

 If blnDevise Then
     If CV_X1.DeviseIso <> selDeviseN Then Exit Function
 End If
 
If blnService Then
    wService = mId$(Msg, 25, 3)
    If optEtat02 Then
         If wService <> "001" And wService <> "011" Then Exit Function
         wCodeOpération = mId$(Msg, 21, 4)
         If wCodeOpération = "G051" Or wCodeOpération = "G052" Then Exit Function
    Else
        If wService <> selService Then Exit Function
    End If
End If


If blnBiaTyp Then
    If mId$(Msg, 15, 3) <> selBiaTyp Then Exit Function
End If

curX = CCur(Val(mId$(Msg, 28, 19)))

If wDeviseN <> "978" Then
    meCV1.DeviseIso = ""
    meCV1.DeviseN = wDeviseN
    meCV1.Montant = curX
    meCV1.OpéAmj = wAmjOpération
    meCV2.OpéAmj = meCV1.OpéAmj
       
    Call CV_Transitoire(meCV1, meCV2, meCV3, X1)
    curX = meCV2.Montant
Else
    meCV2.Montant = curX
    meCV1.DeviseIso = "EUR"
End If

If optEtat01 Then
    If curX > 0 Then
        If curX < curCR Then Exit Function
    Else
        If curX > curDB Then Exit Function
    End If
End If

cmdImport_MvtP0_Select = "OK"

End Function

Private Sub fraEtat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraScript_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
currentActiveControl_Name = C.Name
End Sub
'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
lstErr.Clear
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub


Private Sub lstDevise_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case currentActiveControl_Name
    Case "txtDevise": txtDevise = mId$(lstDevise.Text, 1, 3): If blnSetfocus Then txtDevise.SetFocus
    Case "txtDeviseCV": txtDeviseCV = mId$(lstDevise.Text, 1, 3): If txtDeviseCV.Enabled Then txtDeviseCV.SetFocus
End Select

End Sub


Private Sub optEtat01_Click()
cmdReset
'chkCr = "1": txtCr = "150000"
'chkDb = "1": txtDb = "150000"

End Sub

Private Sub optEtat01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEtat01
End Sub


Private Sub optEtat02_Click()
chkBiaTyp = "1": txtBiaTyp = "001"
chkCr = "1": txtCr = "7000": chkCr.Caption = "Crédit & |Débit|  >  ******* EUR"
chkDb = "1": txtDb = "2000":: chkDb.Caption = "impr MVT >  ******* EUR"
chkService = "1": txtService = "001"
chkCompteMinMax = "1": txtCompteMin = "30000000000": txtCompteMax = "999999999999"
Call DTPicker_Control(txtAmjMax, wAmjMin)
Mid$(wAmjMin, 7, 2) = "01"
Call DTPicker_Set(txtAmjMin, wAmjMin)
End Sub

Private Sub optEtatManuel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEtatManuel
End Sub


Private Sub txtBiaTyp_GotFocus()
txt_GotFocus txtBiaTyp

End Sub


Private Sub txtBiaTyp_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtBiaTyp_LostFocus()
txt_LostFocus txtBiaTyp

End Sub


Private Sub txtCompteMax_GotFocus()
txt_GotFocus txtCompteMax
End Sub


Private Sub txtCompteMax_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtCompteMax_LostFocus()
txt_LostFocus txtCompteMax
End Sub


Private Sub txtCompteMin_GotFocus()
txt_GotFocus txtCompteMin
End Sub


Private Sub txtCompteMin_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtCompteMin_LostFocus()
txt_LostFocus txtCompteMin
End Sub


Private Sub txtDevise_GotFocus()
txt_GotFocus txtDevise
lstDevise.Visible = True

End Sub


Private Sub txtDevise_KeyPress(KeyAscii As Integer)
'Call num_KeyAscii(KeyAscii)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtDevise_LostFocus()
txt_LostFocus txtDevise
lstDevise.Visible = False
End Sub


Private Sub txtDeviseCV_GotFocus()
txt_GotFocus txtDeviseCV
lstDevise.Visible = True

End Sub


Private Sub txtDeviseCV_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtDeviseCV_LostFocus()
txt_LostFocus txtDeviseCV
lstDevise.Visible = False

End Sub

Private Sub txtService_GotFocus()
txt_GotFocus txtService
End Sub


Private Sub txtService_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub

Private Sub txtService_LostFocus()
txt_LostFocus txtService
End Sub



Public Sub cmdReset()
blnControl = False

SSTab1.Enabled = True 'CompteGafiAut.Saisir
'SSTab1.Tabs = 0
DTPicker_Set txtAmjMin, AmjCptVeille
DTPicker_Set txtAmjMax, AmjCptVeille
chkDb = "1": chkDb.Caption = "Débit >  ******* EUR"
chkCr = "1": chkCr.Caption = "Crédit  >  ******* EUR"
txtDeviseCV = "EUR": txtDeviseCV.Enabled = False
meCV1 = CV_Euro
meCV1.CoursCompta = "C"
meCV1.OpéAmj = DSys
meCV1.Normal = "P"
meCV1.AchatVente = " "
meCV2 = meCV1: meCV3 = meCV1

chkCompteMinMax.Value = "0": txtCompteMin = "": txtCompteMax = ""
txtCompteMin.Visible = False:: txtCompteMax.Visible = False
lstDevise.Visible = False
Call LstDictio(889, lstDevise)
chkDeviseIn.Value = "0"
chkDevise = "0": txtDevise = "": txtDevise.Visible = False
chkService.Value = "0": txtService = "": txtService.Visible = False
chkBiaTyp.Value = "0": txtBiaTyp = "": txtBiaTyp.Visible = False
txtCr = "150000": txtDb = "150000"

CV_X1 = CV_Euro
X1000 = ""
    optEtat01.Value = True

If Not blnService_Enabled Then
    txtService = usrService: txtService.Visible = True
    chkService.Value = "1"
    chkService.Enabled = False: txtService.Enabled = False
End If

blnControl = True

End Sub

Public Sub Form_Init()

fraEtat.Enabled = False
fraScript.Enabled = False
fraEtat.Enabled = True 'False

wAmj = dateElp("FinDeMoisP", 0, DSys)
Call DTPicker_Set(txtAmjMax, wAmj)
Mid$(wAmj, 7, 2) = "01"
Call DTPicker_Set(txtAmjMin, wAmj)

End Sub

Public Sub cmdFlux()
Call lstErr_Clear(lstErr, cmdOk, "Flux : Début du traitement")

X = Dir(paramCompteGafi_Mvt_Import)
If X = "" Then Call lstErr_Clear(lstErr, cmdOk, "? Le fichier des mouvements n'existe pas"): Exit Sub

Me.MousePointer = vbHourglass
Me.Enabled = False

cmdImport_MvtP0
If optEtat01 Then prtCompteGafi_Monitor "01", curDB, curCR, wAmjMin, wAmjMax
If optEtat02 Then prtCompteGafi_Monitor "02", curDB, curCR, wAmjMin, wAmjMax

Me.MousePointer = 0
Me.Enabled = True



End Sub

Public Sub cmdImport_MvtP0()
Dim xInput As String, blnOk As Boolean, blnMvtAdd As Boolean, blnCptOk As Boolean
Dim wCodeOpération As String
Dim mRupture As String

On Error Resume Next

Dim I As Long
Call lstErr_AddItem(lstErr, cmdOk, "Chargement des mouvements ...")

blnOk = False: blnCptOk = False
cmdImport_Select_Nb = 0: cmdImport_Nb = 0: I = 0
mID14 = ""

MDB.Execute "delete * from MvtP0"
mdbMvtP0.tableMvtP0_Open

Open paramCompteGafi_Mvt_Import For Input As #1
recMvtP0_Init recMvtp0
recMvtp0.Method = "AddNew"


Open paramCompteGafi_Mvt_Import For Input As #1
I = 0

Do Until EOF(1)
    Line Input #1, xInput
      
    If mId$(xInput, 1, 3) = "$$$" Then
        blnOk = True
        ''SrvMvtP0_Amj = mId$(xInput, 86, 8)
        I = Val(mId$(xInput, 94, 9))
        If I <> cmdImport_Nb Then
            cmdImport_Select_Nb = 0
            Call MsgBox("erreur : nombre enregistrements lus", vbCritical, "frmCompteGafi : cmdflux_Cptp0 :SrvMvtP0 ")
        End If
        Exit Do
    End If

    cmdImport_Nb = cmdImport_Nb + 1
    I = I + 1
    If I = 1000 Then I = 0: Call lstErr_ChangeLastItem(lstErr, cmdOk, "Sélection des mouvements : " & cmdImport_Select_Nb & " / " & cmdImport_Nb): DoEvents
       
    vReturn = cmdImport_MvtP0_Select(xInput)
    If vReturn <> "" Then
        cmdImport_Select_Nb = cmdImport_Select_Nb + 1
        recMvtp0.Text = cur_19P(meCV2.Montant) & meCV1.DeviseIso & xInput
        recMvtp0.Id = mId$(xInput, 10, 11) & mId$(xInput, 7, 3) & mId$(xInput, 163, 8) & cmdImport_Select_Nb  ' COMPTE DEVISE AMJTRT séq
        dbMvtP0_Update recMvtp0
    End If
Loop



Close
tableMvtP0_Close

If Not blnOk Then
'    cmdImport_Select_Nb = 0
    Call MsgBox("erreur : manque fin de fichier ", vbCritical, "frmCompteGafi : cmdflux_Mptp0 :SrvMvtP0 ")
End If

Call lstErr_AddItem(lstErr, cmdOk, "fin des mouvements  : " & cmdImport_Select_Nb & " / " & cmdImport_Nb): DoEvents

End Sub

