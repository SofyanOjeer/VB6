VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCompteExtrait 
   Caption         =   "Extraits de compte"
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
      Left            =   4920
      TabIndex        =   25
      Top             =   0
      Width           =   4185
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
      TabIndex        =   20
      Top             =   0
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   21
      Top             =   360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Sélection des extraits"
      TabPicture(0)   =   "CompteExtrait.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBalance"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Impression"
      TabPicture(1)   =   "CompteExtrait.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraImpression"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraImpression 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   24
         Top             =   480
         Width           =   8775
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H00C0FFC0&
            Height          =   1965
            Left            =   5640
            Picture         =   "CompteExtrait.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   720
            Width           =   2535
         End
         Begin VB.Frame fraMsg 
            Caption         =   "Message"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   120
            TabIndex        =   29
            Top             =   3360
            Width           =   8415
            Begin VB.TextBox txtMsgPersonnel 
               Height          =   285
               Left            =   1080
               TabIndex        =   19
               Top             =   1560
               Width           =   7095
            End
            Begin VB.TextBox txtMsgBanque 
               Height          =   285
               Left            =   1080
               TabIndex        =   17
               Top             =   480
               Width           =   7095
            End
            Begin VB.TextBox txtMsgClientèle 
               Height          =   285
               Left            =   1080
               TabIndex        =   18
               Top             =   1080
               Width           =   7095
            End
            Begin VB.Label lblMsgPersonnel 
               Caption         =   "Personnel"
               Height          =   255
               Left            =   240
               TabIndex        =   32
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label lblMsgBanque 
               Caption         =   "Banques"
               Height          =   255
               Left            =   240
               TabIndex        =   31
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label lblMsgClientèle 
               Caption         =   "Clientèle 03-2001"
               Height          =   495
               Left            =   240
               TabIndex        =   30
               Top             =   1080
               Width           =   1575
            End
         End
         Begin VB.Frame fraOptions 
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
            Height          =   2775
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   4215
            Begin VB.CheckBox chkSoldeFinal 
               Caption         =   "contrôle solde final / solde fin de mois"
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   2160
               Value           =   1  'Checked
               Width           =   3735
            End
            Begin VB.CheckBox chkSoldeInitial 
               Caption         =   "contrôle solde initial / solde dernier extrait"
               Height          =   375
               Left            =   120
               TabIndex        =   15
               Top             =   1560
               Width           =   3255
            End
            Begin VB.CheckBox chkPrintList 
               Caption         =   "impression de la liste des comptes"
               Height          =   375
               Left            =   120
               TabIndex        =   13
               Top             =   360
               Width           =   3255
            End
            Begin VB.CheckBox chkUpdate 
               Caption         =   "mise à jour de la base compte........... (incrémentaion N° extrait, solde dernier extrait)"
               Height          =   615
               Left            =   120
               TabIndex        =   14
               Top             =   960
               Width           =   3855
            End
         End
      End
      Begin VB.Frame fraBalance 
         Height          =   5655
         Left            =   120
         TabIndex        =   22
         Top             =   360
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
            Height          =   3735
            Left            =   3240
            TabIndex        =   26
            Top             =   240
            Width           =   5415
            Begin VB.TextBox txtCompteMin 
               Height          =   285
               Left            =   3120
               MaxLength       =   11
               TabIndex        =   9
               Top             =   1680
               Width           =   1575
            End
            Begin VB.TextBox txtCompteMax 
               Height          =   285
               Left            =   3120
               MaxLength       =   11
               TabIndex        =   10
               Top             =   2400
               Width           =   1575
            End
            Begin VB.CheckBox chkCompteMinMax 
               Caption         =   "sélectionner les comptes de"
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   1800
               Width           =   2535
            End
            Begin VB.CheckBox chkBiaTyp 
               Caption         =   "sélectionner le type de compte"
               Height          =   255
               Left            =   120
               TabIndex        =   11
               Top             =   3120
               Width           =   2535
            End
            Begin VB.TextBox txtBiaTyp 
               Height          =   285
               Left            =   3240
               MaxLength       =   3
               TabIndex        =   12
               Top             =   3120
               Width           =   495
            End
            Begin VB.CheckBox chkCptAux 
               Caption         =   "comptes auxiliaires"
               Height          =   255
               Left            =   120
               TabIndex        =   6
               Top             =   360
               Value           =   1  'Checked
               Width           =   2655
            End
            Begin VB.CheckBox chkCptGen 
               Caption         =   "comptes généraux"
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   1080
               Value           =   1  'Checked
               Width           =   2175
            End
            Begin VB.Label lblMax 
               Caption         =   "à"
               Height          =   255
               Left            =   2280
               TabIndex        =   27
               Top             =   2400
               Width           =   255
            End
         End
         Begin VB.Frame fraScript 
            Caption         =   "Extraits"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3735
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Width           =   2775
            Begin VB.OptionButton optExtraitZ 
               Caption         =   "fin d'exercice                          (mensuels sans mouvements)"
               Height          =   615
               Left            =   120
               TabIndex        =   36
               Top             =   1320
               Width           =   2415
            End
            Begin VB.OptionButton optExtraitAutre 
               Caption         =   "autre période"
               Height          =   255
               Left            =   120
               TabIndex        =   3
               Top             =   2520
               Width           =   2295
            End
            Begin VB.OptionButton optExtraitMensuel 
               Caption         =   "Mensuels"
               Height          =   375
               Left            =   120
               TabIndex        =   0
               Top             =   360
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.OptionButton optExtraitAnnuel 
               Caption         =   "fin d'exercice                            (non mensuels)"
               Height          =   375
               Left            =   120
               TabIndex        =   1
               Top             =   840
               Width           =   2535
            End
            Begin VB.OptionButton optExtraitInventaire 
               Caption         =   "d'inventaire"
               Height          =   255
               Left            =   120
               TabIndex        =   2
               Top             =   2160
               Width           =   2055
            End
            Begin MSComCtl2.DTPicker txtAmjMax 
               Height          =   300
               Left            =   1080
               TabIndex        =   5
               Top             =   3240
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
               Format          =   60227587
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   2
            End
            Begin MSComCtl2.DTPicker txtAmjMin 
               Height          =   300
               Left            =   1080
               TabIndex        =   4
               Top             =   2880
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
               Format          =   60227587
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   2
            End
            Begin VB.Label lblAmjMax 
               Caption         =   "au"
               Height          =   255
               Left            =   480
               TabIndex        =   34
               Top             =   3240
               Width           =   315
            End
            Begin VB.Label lblAmjMin 
               Caption         =   "du"
               Height          =   255
               Left            =   480
               TabIndex        =   33
               Top             =   2880
               Width           =   315
            End
         End
      End
   End
End
Attribute VB_Name = "frmCompteExtrait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' M = Extraits mensuels : au moins 1 mvt dans le mois
' A = Extraits annuels  : non mensuel et solde dans l'exercice et date dernier mvt < 0101***
' Z = Extraits mensuels : comptes sans  mvt dans le mois en cours (décembre)


Option Explicit

Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean, blnControl As Boolean
Dim CptExtraitAut As typeAuthorization
Dim X As String, X1 As String, I As Long
Dim valX As String, V As Variant
Dim reccptp0 As typeCptP0
Dim recCptInfo As typeCptInfo
Dim recCptMvt As typeCptMvt

Dim optEtat As String * 1, SrvCptP0_Amj As String * 8, SrvMvtP0_Amj As String * 8, SrvCptP0_Amj_Ok As String
Dim blnCompteMinMax As Boolean, selCompteMin As String * 11, selCompteMax As String * 11
Dim blnBiaTyp As Boolean, selBiaTyp As String * 3
Dim cmdImport_Select_Nb As Long, cmdImport_Nb As Long

Dim valAmjMin As String * 8, valAmjMax As String * 8, valAmjDernierMouvement As String * 8
Dim prtCptMvt_Nb As Long
Dim cumulMvt As Currency, curSoldeFinal  As Currency, mExtraitNuméro As String * 3
Dim xErr As String * 12

Dim blnInventaire As Boolean
'''''Dim mAMJMin As String * 8, mAMJMax As String * 8
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
Call BiaPgmAut_Init("Compte_Ext", CptExtraitAut)
If Not IsNull(param_Init) Then cmdPrint.Visible = False
SrvCptP0_Amj_Ok = vbNo
cmdReset

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
paramComptaExt_Cpt_Import = paramServer(recElpTable.Memo)
Call lstErr_Clear(lstErr, cmdContext, "Cpt_Import:" & paramComptaExt_Cpt_Import)

recElpTable.K2 = "Mvt_Import"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramComptaExt_Mvt_Import = paramServer(recElpTable.Memo)
Call lstErr_AddItem(lstErr, cmdContext, "Cpt_Import:" & paramComptaExt_Mvt_Import)

recElpTable.K2 = "Cpt_Export"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramComptaExt_Cpt_Export = paramServer(recElpTable.Memo)
Call lstErr_AddItem(lstErr, cmdContext, "Cpt_Import:" & paramComptaExt_Cpt_Export)

recElpTable.K2 = "Msg_Banque"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then
'    GoTo Table_Error
Else
    If IsNull(recElpTable.Memo) Then
'        GoTo Memo_Error
    Else
        X = Trim(recElpTable.Memo)
        If X <> "" Then txtMsgBanque = X
    End If
End If

recElpTable.K2 = "Msg_Client"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then
'    GoTo Table_Error
Else
    If IsNull(recElpTable.Memo) Then
'        GoTo Memo_Error
    Else
        X = Trim(recElpTable.Memo)
        If X <> "" Then txtMsgClientèle = X
    End If
End If

recElpTable.K2 = "Msg_Perso"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then
'    GoTo Table_Error
Else
    If IsNull(recElpTable.Memo) Then
'        GoTo Memo_Error
    Else
        X = Trim(recElpTable.Memo)
        If X <> "" Then txtMsgPersonnel = X
    End If
End If

Exit Function

Table_Error:
param_Init = V
Exit Function

Memo_Error:
param_Init = "Memo"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "Balance_Param_Init"
Exit Function

End Function



Private Sub cmdImport_CptP0()
Dim xInput As String, blnOk As Boolean
Dim vReturn As Variant
On Error Resume Next

Dim I As Integer
blnOk = False
cmdImport_Select_Nb = 0: cmdImport_Nb = 0: I = 0
cmdControl
If lstErr.ListCount <> 0 Then Exit Sub
X = Dir(paramComptaExt_Cpt_Import)
If X = "" Then Call lstErr_Clear(lstErr, cmdPrint, "? Le fichier des comptes n'existe pas"): Exit Sub

Call lstErr_Clear(lstErr, cmdPrint, "Chargement des comptes, tri ..."): DoEvents
Me.MousePointer = vbHourglass
Me.Enabled = False

MDB.Execute "delete * from CptP0"
mdbCptP0.tableCptP0_Open

Open paramComptaExt_Cpt_Import For Input As #1
recCptP0_Init reccptp0
reccptp0.Method = "AddNew"


Do Until EOF(1)
    Line Input #1, xInput
    
    If mId$(xInput, 1, 3) = "$$$" Then
        blnOk = True
        SrvCptP0_Amj = mId$(xInput, 35, 8)
        I = Val(mId$(xInput, 43, 9))
        If I <> cmdImport_Nb Then
            cmdImport_Select_Nb = 0
            Call MsgBox("erreur : nombre enregistrements lus", vbCritical, "frmCompteEXtrait : cmdImport_Cptp0 :SrvCptP0 ")
            Exit Do
        End If
    End If

    cmdImport_Nb = cmdImport_Nb + 1
    vReturn = cmdImport_Select(xInput)
    If vReturn <> "" Then
        reccptp0.Id = vReturn & Format$(cmdImport_Nb, "000000")
        reccptp0.Text = xInput
            cmdImport_Select_Nb = cmdImport_Select_Nb + 1
            dbCptP0_Update reccptp0
    End If
    If I = 1000 Then I = 0: Call lstErr_ChangeLastItem(lstErr, cmdPrint, "Sélection des comptes : " & cmdImport_Select_Nb & " / " & cmdImport_Nb): DoEvents
 
Loop

Close
mdbCptP0.tableCptP0_Close
Me.MousePointer = 0
If Not blnOk Then
    cmdImport_Select_Nb = 0
    Call MsgBox("erreur : manque fin de fichier ", vbCritical, "frmCompteEXtrait : cmdImport_Cptp0 :SrvCptP0 ")
End If

End Sub

Private Sub cmdImport_MvtP0()
Dim xInput As String, blnOk As Boolean
Dim vReturn As Variant
On Error Resume Next

Dim I As Long
blnOk = False
cmdImport_Select_Nb = 0: cmdImport_Nb = 0: I = 0
'Exit Sub
X = Dir(paramComptaExt_Mvt_Import)
If X = "" Then Call lstErr_Clear(lstErr, cmdPrint, "? Le fichier des mouvements n'existe pas"): Exit Sub

Call lstErr_AddItem(lstErr, cmdPrint, "Chargement des mouvements, tri ..."): DoEvents

MDB.Execute "delete * from MvtP0"
mdbMvtP0.tableMvtP0_Open

Open paramComptaExt_Mvt_Import For Input As #1
recMvtP0_Init recMvtp0
recMvtp0.Method = "AddNew"


Do Until EOF(1)
    Line Input #1, xInput
    
        
    If mId$(xInput, 1, 3) = "$$$" Then
        blnOk = True
        SrvMvtP0_Amj = mId$(xInput, 86, 8)
        I = Val(mId$(xInput, 94, 9))
        If I <> cmdImport_Nb Then
            cmdImport_Select_Nb = 0
            Call MsgBox("erreur : nombre enregistrements lus", vbCritical, "frmCompteEXtrait : cmdImport_Cptp0 :SrvMvtP0 ")
            Exit Do
        End If
    End If

    cmdImport_Nb = cmdImport_Nb + 1
 '   vReturn = mId$(xInput, 7, 14) & mId$(xInput, 163, 8) & mId$(xInput, 71, 11) ' devise compte date trt no pièce no ligne
'    If vReturn <> "" Then
    If mId$(xInput, 163, 8) >= valAmjMin And mId$(xInput, 163, 8) <= valAmjMax Then
        recMvtp0.Id = mId$(xInput, 7, 14) & mId$(xInput, 163, 8) & mId$(xInput, 71, 11) & Format$(cmdImport_Nb, "0000000") ' devise compte date trt no pièce no ligne
        recMvtp0.Text = xInput
            cmdImport_Select_Nb = cmdImport_Select_Nb + 1
            dbMvtP0_Update recMvtp0
    End If
'    End If

    If I = 1000 Then I = 0: Call lstErr_ChangeLastItem(lstErr, cmdPrint, "Sélection des mouvements: " & cmdImport_Select_Nb & " / " & cmdImport_Nb): DoEvents
 
Loop

Close
mdbMvtP0.tableMvtP0_Close

If Not blnOk Then
    cmdImport_Select_Nb = 0
    Call MsgBox("erreur : manque fin de fichier ", vbCritical, "frmCompteEXtrait : cmdImport_Cptp0 :SrvMvtP0 ")
End If

End Sub



Private Sub chkBiaTyp_Click()
If chkBiaTyp = "1" Then
    txtBiaTyp.Visible = True: txtBiaTyp.SetFocus
Else
    txtBiaTyp.Visible = False
End If
If blnControl Then cmdControl

End Sub

Private Sub chkBiaTyp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkBiaTyp
End Sub


Private Sub chkCompteMinMax_Click()
If chkCompteMinMax = "1" Then
    txtCompteMin.Visible = True: txtCompteMax.Visible = True: txtCompteMin.SetFocus
Else
    txtCompteMin.Visible = False: txtCompteMax.Visible = False
End If
If blnControl Then cmdControl

End Sub

Private Sub chkCompteMinMax_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkCompteMinMax
End Sub


Private Sub chkCptAux_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkCptAux
End Sub


Private Sub chkCptGen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkCptGen
End Sub


Private Sub chksoldefinal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkSoldeFinal
End Sub


Private Sub chksoldeinitial_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkSoldeInitial
End Sub


Private Sub chkprintlist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkPrintList
End Sub


Private Sub chkUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkUpdate
End Sub


Private Sub cmdPrint_Click()
Dim X, Nb As Integer, curX As Currency, IdKey As String, mIdKey As String
Dim Msg As String

cmdImport_CptP0
If cmdImport_Select_Nb = 0 Then
    Call lstErr_AddItem(lstErr, cmdPrint, "Aucun compte sélectionné !")
    GoTo cmdPrint_End
End If


cmdImport_MvtP0
If SrvCptP0_Amj <> SrvMvtP0_Amj Then
    Call MsgBox("Date Cpt : " & dateImp(SrvCptP0_Amj) & " Date Mvt : " & dateImp(SrvMvtP0_Amj), vbCritical, "frmCompteExtrait : cmdPrint.Click")
Else
    If SrvCptP0_Amj_Ok <> vbYes Then SrvCptP0_Amj_Ok = MsgBox("Confirmez la date de dernière compta : " & dateImp(SrvCptP0_Amj), vbQuestion + vbYesNo, "frmCompteExtrait : cmdPrint.Click")
    If SrvCptP0_Amj_Ok = vbYes Then
        Call lstErr_AddItem(lstErr, cmdPrint, "Impression : début")
        cmdPrint_Cpt_Load
        Call lstErr_AddItem(lstErr, cmdPrint, "Impression terminé : " & cmdImport_Select_Nb)
        If chkPrintList = "1" Then prtCompteExtrait_Monitor paramComptaExt_Cpt_Export
    End If
End If
cmdPrint_End:
Me.Enabled = True
AppActivate Me.Caption

End Sub



Public Sub cmdControl()
Dim wX As String

If Not Me.Enabled Then Exit Sub
Me.Enabled = False
'frmCptExtrait.Enabled = False

'cmdPrint.Visible = False
blnControl = False

lstErr.Clear
lstErr.Height = 200


lstErr.Clear
optEtat = ""
If optExtraitMensuel Then optEtat = "M"
If optExtraitAnnuel Then optEtat = "A"
If optExtraitZ Then optEtat = "Z"

If optEtat = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? option extrait non programmée")

cmdControl_txtAmj
blnCompteMinMax = IIf(chkCompteMinMax = "1", True, False)

wX = Trim(txtCompteMin)
If Trim(txtCompteMax) = "" And Trim(txtCompteMin) <> "" Then
    If Val(wX) < 100000 Then
        txtCompteMin = wX & "000000"
        txtCompteMax = wX & "999999"
    Else
        txtCompteMax = wX
    End If
End If
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

blnBiaTyp = IIf(chkBiaTyp = "1", True, False)
selBiaTyp = Format$(Trim(txtBiaTyp), "000")
If blnBiaTyp Then
    If Trim(txtBiaTyp) = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le Type")
End If

'If lstErr.ListCount = 0 Then cmdPrint.Visible = True

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

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint
End Sub

Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
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

Private Function cmdImport_Select(Msg As String) As String
Dim wCompteGénéral As String * 11, wDevise As String * 3, wNuméro As String * 11, wExtraitPériodicité As String * 1, wTypeGa As String * 1
Dim xAMJ As String * 8, xDébitFindeMois As Currency, xCréditFindeMois As Currency

cmdImport_Select = ""
wDevise = mId$(Msg, 7, 3)
wNuméro = mId$(Msg, 13, 11)
wTypeGa = mId$(Msg, 115, 1)
wCompteGénéral = Format$(Val(mId$(Msg, 255, 11)), "00000000000")
wExtraitPériodicité = mId$(Msg, 499, 1)
' wNuméro = "60242001014" Then
'    xAMJ = ""
'End If

Select Case optEtat
    Case "M"
        If wExtraitPériodicité <> 1 _
        And wExtraitPériodicité <> 3 _
        And wExtraitPériodicité <> 5 Then Exit Function
    Case "A"
        If wExtraitPériodicité = 0 _
        Or wExtraitPériodicité = 1 _
        Or wExtraitPériodicité = 3 _
        Or wExtraitPériodicité = 5 Then Exit Function
        
        xAMJ = mId$(Msg, 562, 8)
        xDébitFindeMois = CCur(Val(mId$(Msg, 338, 19)))
        xCréditFindeMois = CCur(Val(mId$(Msg, 357, 19)))
        If xDébitFindeMois = 0 And xCréditFindeMois = 0 And xAMJ < valAmjDernierMouvement Then Exit Function
    Case "Z"
        If wExtraitPériodicité <> 1 _
        And wExtraitPériodicité <> 3 _
        And wExtraitPériodicité <> 5 Then Exit Function
         
        xAMJ = mId$(Msg, 562, 8)
        If xAMJ >= valAmjDernierMouvement Then Exit Function
        xDébitFindeMois = CCur(Val(mId$(Msg, 338, 19)))
        xCréditFindeMois = CCur(Val(mId$(Msg, 357, 19)))
        If (xDébitFindeMois + xCréditFindeMois) = 0 Then Exit Function
       
    Case Else: Exit Function
End Select

If blnCompteMinMax Then
    If wTypeGa = "A" Then
        If wNuméro < selCompteMin Or wNuméro > selCompteMax Then Exit Function
    Else
        If wCompteGénéral < selCompteMin Or wCompteGénéral > selCompteMax Then Exit Function
    End If
End If

Select Case wTypeGa
    Case "A":   If chkCptAux <> "1" Then Exit Function
                If blnBiaTyp Then
                    If mId$(Msg, 241, 3) <> selBiaTyp Then Exit Function
                End If
                cmdImport_Select = "A" & mId$(wNuméro, 6, 3) & mId$(wNuméro, 1, 5) & mId$(wNuméro, 9, 2) & wDevise
    Case "G":   If chkCptGen <> "1" Then Exit Function
                cmdImport_Select = "G" & wNuméro & wDevise

End Select




End Function

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


Private Sub optExtraitAnnuel_Click()
If blnControl Then cmdControl

End Sub

Private Sub optExtraitAnnuel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optExtraitAnnuel
End Sub


Private Sub optExtraitAutre_Click()
If blnControl Then cmdControl

End Sub

Private Sub optExtraitAutre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optExtraitAutre
End Sub


Private Sub optExtraitInventaire_Click()
If blnControl Then cmdControl

End Sub

Private Sub optExtraitInventaire_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optExtraitInventaire
End Sub


Private Sub optExtraitMensuel_Click()
If blnControl Then cmdControl

End Sub

Private Sub optExtraitMensuel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optExtraitMensuel
End Sub


Private Sub optExtraitZ_Click()
If blnControl Then cmdControl

End Sub

Private Sub optExtraitZ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optExtraitZ

End Sub


Private Sub txtAmjMax_Change()
If blnControl Then cmdControl

End Sub

Private Sub txtAmjMax_GotFocus()
DTPicker_GotFocus txtAmjMax
End Sub


Private Sub txtAmjMax_LostFocus()
DTPicker_LostFocus txtAmjMax
If blnControl Then cmdControl
End Sub


Private Sub txtAmjMin_Change()
If blnControl Then cmdControl

End Sub

Private Sub txtAmjMin_GotFocus()
DTPicker_GotFocus txtAmjMin

End Sub


Private Sub txtAmjMin_LostFocus()
DTPicker_LostFocus txtAmjMin
If blnControl Then cmdControl
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
If blnControl Then cmdControl
End Sub


Private Sub txtCompteMin_GotFocus()
txt_GotFocus txtCompteMin
End Sub


Private Sub txtCompteMin_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtCompteMin_LostFocus()
txt_LostFocus txtCompteMin
If blnControl Then cmdControl
End Sub


Private Sub txtmsgbanque_GotFocus()
txt_GotFocus txtMsgBanque
End Sub


Private Sub txtmsgbanque_LostFocus()
txt_LostFocus txtMsgBanque
End Sub


Private Sub txtmsgclientèle_GotFocus()
txt_GotFocus txtMsgClientèle
End Sub


Private Sub txtmsgclientèle_LostFocus()
txt_LostFocus txtMsgClientèle
End Sub


Private Sub txtmsgpersonnel_GotFocus()
txt_GotFocus txtMsgPersonnel
End Sub


Private Sub txtmsgpersonnel_LostFocus()
txt_LostFocus txtMsgPersonnel
End Sub



Public Sub cmdReset()
blnControl = False
SSTab1.Enabled = CptExtraitAut.Saisir
optExtraitMensuel = True: cmdControl_txtAmj
chkUpdate.Enabled = False

chkCompteMinMax.Value = "0": txtCompteMin = "": txtCompteMax = ""
txtCompteMin.Visible = False:: txtCompteMax.Visible = False

chkBiaTyp.Value = "0": txtBiaTyp = "": txtBiaTyp.Visible = False
blnControl = True
End Sub

Public Sub cmdControl_txtAmj()
txtAmjMin.Enabled = False: txtAmjMax.Enabled = False
valAmjMax = DSys
If mId$(valAmjMax, 7, 2) < "10" Then
    valAmjMax = dateElp("FinDeMoisP", 0, valAmjMax)
Else
    valAmjMax = dateFinDeMois(valAmjMax)
End If
valAmjMin = valAmjMax

If optExtraitMensuel Then
    Mid$(valAmjMin, 7, 2) = "01"
    valAmjDernierMouvement = valAmjMin
End If

If optExtraitAnnuel Then
    Mid$(valAmjMin, 5, 4) = "0101"
    valAmjDernierMouvement = valAmjMin
End If

If optExtraitZ Then
    Mid$(valAmjMin, 7, 2) = "01"
    valAmjDernierMouvement = valAmjMax
    Mid$(valAmjDernierMouvement, 7, 2) = "01"
End If

If optExtraitInventaire Then
    Mid$(valAmjMin, 5, 4) = "0101"
End If


If optExtraitAutre Then
    txtAmjMin.Enabled = True: txtAmjMax.Enabled = True
    V = DTPicker_Control(txtAmjMax, valAmjMax)
    If Not IsNull(V) Then Call lstErr_AddItem(lstErr, txtAmjMax, V): Exit Sub
    V = DTPicker_Control(txtAmjMin, valAmjMin)
    If Not IsNull(V) Then Call lstErr_AddItem(lstErr, txtAmjMin, V): Exit Sub
    If valAmjMin > valAmjMax Then Call lstErr_AddItem(lstErr, txtAmjMin, "? Date début > date fin")
    If valAmjMax >= DSys Then Call lstErr_AddItem(lstErr, txtAmjMin, "? Date fin >= jour")
    
End If

txtAmjMax = dateImp(valAmjMax)
txtAmjMin = dateImp(valAmjMin)

End Sub

Public Sub cmdPrint_Cpt_Load()
Dim Msg As String, mMsgInfo As String
Dim X14 As String * 14
Dim blnprtCptMvt_Extrait As Boolean

prtCptMvt_Nb = 0
recCptMvtInit recCptMvt
         
recCptMvt.Société = SocId$
recCptMvt.Agence = SocAgence$
ReDim arrCptMvt(1000): arrCptMvtNbMax = 1000

Msg = "000001000000***" & valAmjMin & valAmjMax
Mid$(Msg, 15, 1) = "E"
Open paramComptaExt_Cpt_Export For Output As #1

prtCptMvt_Open Msg
Call lstErr_AddItem(lstErr, cmdPrint, "Impression : début")
mdbCptP0.tableCptP0_Open
mdbMvtP0.tableMvtP0_Open

reccptp0.Method = "MoveFirst"
Mid$(MsgTxt, 1, 34) = Space$(34)

V = dbCptP0_ReadE(reccptp0)

Do While reccptp0.Err = 0
    
    
    MsgTxtIndex = 0
    MsgTxt = Space$(recCptInfoLen)
    Mid$(MsgTxt, 35, memoCptInfoLen) = mId$(reccptp0.Text, 1, memoCptInfoLen)
    If IsNull(srvCptInfoGetBuffer(recCptInfo)) Then
        X14 = recCptInfo.Devise & recCptInfo.Numéro
        
        Select Case optEtat
            Case "M": blnprtCptMvt_Extrait = cmdPrint_Mvt_Load(X14)
            Case "A", "Z": blnprtCptMvt_Extrait = cmdPrint_Mvt_AS400(X14)
        End Select
''test  Exit Do
        If blnprtCptMvt_Extrait Then
            xErr = Space$(10)
            cmdPrint_Solde_Control xErr
       
            prtCptMvt_Nb = prtCptMvt_Nb + 1
            Call lstErr_ChangeLastItem(lstErr, cmdPrint, "Impression " & prtCptMvt_Nb & " / " & recCptInfo.Numéro & " : " & arrCptMvtNb): DoEvents
            Mid$(Msg, 7, 6) = Format$(arrCptMvtNb, "000000")
            
            If recCptInfo.TypeGA <> "A" Or recCptInfo.BiaTyp <> "001" Then
                mMsgInfo = ""
            Else
                If recCptInfo.Numéro < "30000000000" Then
                    mMsgInfo = Trim(txtMsgBanque)
                Else
                
                   If recCptInfo.Gestionnaire = "60" Then
                        mMsgInfo = Trim(txtMsgPersonnel)
                    Else
'2001.04.02 JPL                   If recCptInfo.Numéro > "50000000000" And recCptInfo.Numéro < "80000000000" Then

                        mMsgInfo = Trim(txtMsgClientèle)
                    End If
                End If
            End If
            If chkUpdate.Enabled = "1" Then
                mExtraitNuméro = Format$(Val(recCptInfo.ExtraitNuméro) + 1, "000")
            Else
                mExtraitNuméro = Format$(Val(recCptInfo.ExtraitNuméro), "000")
                If mExtraitNuméro = "000" Then mExtraitNuméro = "001"
            End If
            
            If recCptInfo.Courrier <> "0" Then
                recCptInfo.Adresse2 = ""
                recCptInfo.Adresse3 = ""
                recCptInfo.Adresse4 = ""
                recCptInfo.Adresse5 = ""
                recCptInfo.AdresseCP = ""
                recCptInfo.AdresseBD = ""
                recCptInfo.AdressePays = ""
                recCptInfo.Adresse4 = DicLib(61, recCptInfo.Courrier)
                If recCptInfo.Courrier = "9" Then recCptInfo.Adresse4 = Trim(recCptInfo.Adresse4) & " : " & DicLib(60, recCptInfo.Gestionnaire)
                
            End If
            
            prtCptMvt_Extrait Msg, recCptInfo, mMsgInfo, mExtraitNuméro
            cmdPrint_Cpt_Export xErr

        End If
    End If
    
    reccptp0.Method = "MoveNext    "
    reccptp0.Err = tableCptP0_Read(reccptp0)
Loop
prtCptMvt_Close
mdbCptP0.tableCptP0_Close
mdbMvtP0.tableMvtP0_Close
Close #1

End Sub


Public Function cmdPrint_Mvt_Load(mMvt_Id As String) As Boolean
Dim vReturn As Integer
recMvtp0.Method = "Seek>="
arrCptMvtNb = 0
cumulMvt = 0
recMvtp0.Id = mMvt_Id
vReturn = 0
Do
    vReturn = tableMvtP0_Read(recMvtp0)
    If vReturn = 0 Then
        If mId$(recMvtp0.Id, 1, 14) <> mMvt_Id Then
            vReturn = 9999
        Else
            MsgTxtIndex = 0
            Mid$(MsgTxt, 35, memoCptMvtLen) = mId$(recMvtp0.Text, 1, memoCptMvtLen)
            If IsNull(srvCptMvtGetBuffer(recCptMvt)) Then
                
                Call arrCptMvtAddItem(recCptMvt): cumulMvt = cumulMvt + recCptMvt.MT
                
'$2002.02.01 jpl      Select Case recCptMvt.CptComplémentaire

'$2002.02.01 jpl        Case "3", "4": If blnInventaire Then Call arrCptMvtAddItem(recCptMvt): cumulMvt = cumulMvt + recCptMvt.MT

'$2002.02.01 jpl        Case Else: If Not blnInventaire Then Call arrCptMvtAddItem(recCptMvt): cumulMvt = cumulMvt + recCptMvt.MT

'$2002.02.01 jpl      End Select
                recMvtp0.Method = "MoveNext"
            End If
        End If
    End If
Loop While vReturn = 0
        
If arrCptMvtNb > 0 Then
    cmdPrint_Mvt_Load = True
Else
    cmdPrint_Mvt_Load = False
End If

End Function

Public Sub cmdPrint_Solde_Control(xErr As String)
Dim X As String

curSoldeFinal = arrCptMvt(1).SoldeVeille + cumulMvt

If chkSoldeInitial = "1" Then
    If recCptInfo.ExtraitSolde <> arrCptMvt(1).SoldeVeille Then
        xErr = "ErrSoldeI"
        X = MsgBox("Compte : " & recCptInfo.Devise & " " & recCptInfo.Numéro & Chr$(10) & Chr$(13) & "Erreur solde initial (Compte / Mvt) ?", , Me.Caption)
 'IIf(X = vbNo, True, False)
    End If
End If

If chkSoldeFinal = "1" Then
    If recCptInfo.SoldeVeille <> curSoldeFinal Then
        xErr = "ErrSoldeF"
        X = MsgBox("Compte : " & recCptInfo.Devise & " " & recCptInfo.Numéro & Chr$(10) & Chr$(13) & "Erreur solde final (Compte / Mvt) ?", , Me.Caption)
 'IIf(X = vbNo, True, False)
    End If
End If

End Sub

Public Sub cmdPrint_Cpt_Export(xErr As String)
Dim xOut As String * 256
xOut = Space$(256)
Mid$(xOut, 1, 12) = "SRVCPTUPD"
Mid$(xOut, 13, 12) = "ComptaExt"
Mid$(xOut, 25, 10) = xErr
Mid$(xOut, 34 + 1, 3) = recCptInfo.Société
Mid$(xOut, 34 + 4, 3) = recCptInfo.Agence
Mid$(xOut, 34 + 7, 3) = Format$(Val(recCptInfo.Devise), "000")
Mid$(xOut, 34 + 10, 11) = Format$(Val(recCptInfo.Numéro), "00000000000")
Mid$(xOut, 34 + 21, 15) = Format$(Abs(curSoldeFinal) * 100, "000000000000000")
Mid$(xOut, 34 + 36, 1) = IIf(curSoldeFinal < 0, "D", "C")
Mid$(xOut, 34 + 37, 3) = mExtraitNuméro
Mid$(xOut, 34 + 40, 8) = valAmjMax
Print #1, xOut
End Sub

Public Function cmdPrint_Mvt_AS400(lX14 As String) As Boolean
Dim Nb As Integer

recCptMvt.Devise = mId$(lX14, 1, 3)
recCptMvt.Compte = mId$(lX14, 4, 11)

recCptMvt.Method = "SnapLA"
recCptMvt.AmjTraitement = valAmjMin
recCptMvt.Pièce = 0
recCptMvt.Ligne = 0

arrCptMvt(0) = recCptMvt
arrCptMvt(0).AmjTraitement = valAmjMax
arrCptMvt(0).Pièce = "999999999"
arrCptMvt(0).Ligne = "9999"
    
arrCptMvtSuite = True
arrCptMvtNb = 0
cumulMvt = 0

Do Until Not arrCptMvtSuite
    srvCptMvtMon recCptMvt
    recCptMvt = arrCptMvt(arrCptMvtNb)
    recCptMvt.Method = "SnapLA+"
Loop
cmdPrint_Mvt_AS400 = True

Nb = arrCptMvtNb
arrCptMvtNb = 0

For I = 1 To Nb
'    Select Case arrCptMvt(I).CptComplémentaire
'        Case "3", "4":
'        Case Else:
            arrCptMvtNb = arrCptMvtNb + 1
            arrCptMvt(arrCptMvtNb) = arrCptMvt(I)
            cumulMvt = cumulMvt + arrCptMvt(arrCptMvtNb).MT
'    End Select
Next I

If arrCptMvtNb = 0 Then
    arrCptMvt(0).SoldeVeille = recCptInfo.SoldeFindeMois
    arrCptMvt(1).SoldeVeille = recCptInfo.SoldeFindeMois
End If

End Function
