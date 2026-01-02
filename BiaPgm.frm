VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBiaPgm 
   AutoRedraw      =   -1  'True
   Caption         =   "BiaPgm : mise à jour de la table des programmes Bia.vbp"
   ClientHeight    =   9492
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   13872
   Icon            =   "BiaPgm.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9492
   ScaleWidth      =   13872
   Begin TabDlg.SSTab SSTab1 
      Height          =   8850
      Left            =   0
      TabIndex        =   17
      Top             =   360
      Width           =   5685
      _ExtentX        =   10033
      _ExtentY        =   15600
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Paramètrage"
      TabPicture(0)   =   "BiaPgm.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBiaPgm"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Utilisateurs autorisés"
      TabPicture(1)   =   "BiaPgm.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstBiaPgmAut"
      Tab(1).ControlCount=   1
      Begin VB.ListBox lstBiaPgmAut 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8004
         Left            =   -74760
         TabIndex        =   24
         Top             =   600
         Width           =   4875
      End
      Begin VB.Frame fraBiaPgm 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8295
         Left            =   0
         TabIndex        =   18
         Top             =   360
         Width           =   5535
         Begin VB.Frame fraBiaPgmAut 
            Caption         =   "Avis"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4440
            Left            =   960
            TabIndex        =   19
            Top             =   2520
            Width           =   3555
            Begin VB.CheckBox chkXHost 
               Caption         =   "Migration : opt autorisée"
               Height          =   372
               Left            =   240
               TabIndex        =   26
               Top             =   3600
               Width           =   3012
            End
            Begin VB.CheckBox chkXspécial 
               Caption         =   "Spécial"
               Height          =   255
               Left            =   240
               TabIndex        =   25
               Top             =   3240
               Width           =   975
            End
            Begin VB.CheckBox chkConsulter 
               Caption         =   "Consulter"
               Height          =   255
               Left            =   240
               TabIndex        =   3
               Top             =   360
               Width           =   1335
            End
            Begin VB.CheckBox chkSaisir 
               Caption         =   "Saisir"
               Height          =   255
               Left            =   240
               TabIndex        =   4
               Top             =   720
               Width           =   1335
            End
            Begin VB.CheckBox chkValider 
               Caption         =   "Valider"
               Height          =   255
               Left            =   240
               TabIndex        =   5
               Top             =   1080
               Width           =   1335
            End
            Begin VB.CheckBox chkComptabiliser 
               Caption         =   "Comptabiliser"
               Height          =   255
               Left            =   240
               TabIndex        =   6
               Top             =   1440
               Width           =   1335
            End
            Begin VB.CheckBox chkRapprocher 
               Caption         =   "Rapprocher"
               Height          =   255
               Left            =   240
               TabIndex        =   7
               Top             =   1800
               Width           =   1335
            End
            Begin VB.CheckBox chkSwift 
               Caption         =   "Swift"
               Height          =   255
               Left            =   240
               TabIndex        =   8
               Top             =   2160
               Width           =   855
            End
            Begin VB.CheckBox chkVirement 
               Caption         =   "Virement"
               Height          =   255
               Left            =   240
               TabIndex        =   9
               Top             =   2520
               Width           =   975
            End
            Begin VB.CheckBox chkAvis 
               Caption         =   "Avis"
               Height          =   255
               Left            =   240
               TabIndex        =   10
               Top             =   2880
               Width           =   975
            End
         End
         Begin VB.CommandButton cmdSuppress 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Supprimer"
            Height          =   645
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   7440
            Width           =   1980
         End
         Begin VB.CommandButton cmdOk 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Ok"
            Height          =   1125
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   7080
            Visible         =   0   'False
            Width           =   1980
         End
         Begin VB.TextBox txtId 
            Height          =   285
            Left            =   960
            MaxLength       =   12
            TabIndex        =   0
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   960
            MaxLength       =   36
            TabIndex        =   1
            Top             =   1200
            Width           =   3615
         End
         Begin VB.TextBox txtProjet 
            Height          =   285
            Left            =   960
            MaxLength       =   20
            TabIndex        =   2
            Text            =   "Bia"
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label lblId 
            Caption         =   "Identifiant"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblName 
            Caption         =   "Intitulé"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lblProjet 
            Caption         =   "Projet.vbp"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1920
            Width           =   975
         End
      End
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   8520
      TabIndex        =   15
      Top             =   0
      Width           =   4785
   End
   Begin VB.CommandButton cmdAddNew 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ajouter un pgm"
      Height          =   405
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   1260
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Abandonner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   400
      Left            =   13320
      Picture         =   "BiaPgm.frx":047A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   0
      Width           =   500
   End
   Begin MSFlexGridLib.MSFlexGrid fgBiaPgm 
      Height          =   8955
      Left            =   5760
      TabIndex        =   23
      Top             =   360
      Width           =   8115
      _ExtentX        =   14309
      _ExtentY        =   15790
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   280
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
      AllowUserResizing=   3
      FormatString    =   $"BiaPgm.frx":057C
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "mnuPrint"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint0 
         Caption         =   "Imprimer la liste des programmes"
      End
      Begin VB.Menu mnuPrint1 
         Caption         =   "Imprimer la liste des programmes et des utilisateurs"
      End
   End
End
Attribute VB_Name = "frmBiaPgm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim BiaPgmAut As typeAuthorization
Dim X As String, X1 As String, I As Integer
Dim Msg As String, valX As String, V As Variant

Dim currentMethod As String, currentAMJ As String
Dim fgBiaPgm_FormatString As String, fgBiaPgm_K As Integer
Dim fgBiaPgm_BackColorFixed As Long, fgBiaPgm_BackColor As Long

Dim recBiaPgm As typeElpTable, recBiaPgmAut As typeElpTable
Dim lstI As Integer
Dim mBiaPgm As String * 12


'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
''cmdControl
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
currentActiveControl_Name = C.Name
End Sub
'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub



Private Sub chkAvis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkAvis
End Sub


Private Sub chkVirement_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkVirement
End Sub


Private Sub chkXHost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkXHost

End Sub


Private Sub chkXspécial_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkXspécial

End Sub


Private Sub cmdAddNew_Click()
rsElpTable_Init recBiaPgm
currentMethod = constAddNew
recBiaPgm.Id = paramBiaPgm

fgBiaPgm.Enabled = False
fraBiaPgm_Reset
fraBiaPgm.Enabled = True
fraBiaPgmAut.Enabled = True
txtId.Enabled = True
txtId.SetFocus
cmdOk.Visible = BiaPgmAut.Valider

End Sub

Private Sub cmdAddNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdAddNew
End Sub


Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: cmdContext_Return
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext
End Sub


Private Sub cmdOk_Update()
Dim V

Select Case currentMethod
    Case constAddNew: V = adoElpTable_AddNew(rsMDB, recBiaPgm)
    Case constDelete: V = adoElpTable_Delete(rsMDB, recBiaPgm)
    Case constUpdate: V = adoElpTable_Update(rsMDB, recBiaPgm)
End Select
fgBiaPgm_Display
cmdContext_Quit

End Sub

Private Sub cmdOk_Click()

recBiaPgm.Memo = Space$(40)
Mid$(recBiaPgm.Memo, 1, 1) = IIf(chkConsulter.Value = 1, "X", " ")
Mid$(recBiaPgm.Memo, 2, 1) = IIf(chkSaisir.Value = 1, "X", " ")
Mid$(recBiaPgm.Memo, 3, 1) = IIf(chkValider.Value = 1, "X", " ")
Mid$(recBiaPgm.Memo, 4, 1) = IIf(chkComptabiliser.Value = 1, "X", " ")
Mid$(recBiaPgm.Memo, 5, 1) = IIf(chkRapprocher.Value = 1, "X", " ")
Mid$(recBiaPgm.Memo, 6, 1) = IIf(chkSwift.Value = 1, "X", " ")
Mid$(recBiaPgm.Memo, 7, 1) = IIf(chkVirement.Value = 1, "X", " ")
Mid$(recBiaPgm.Memo, 8, 1) = IIf(chkAvis.Value = 1, "X", " ")
Mid$(recBiaPgm.Memo, 9, 1) = IIf(chkXspécial.Value = 1, "X", " ")
Mid$(recBiaPgm.Memo, 10, 1) = IIf(chkXHost.Value = 1, "X", " ")
recBiaPgm.K1 = Trim(txtId)
recBiaPgm.Name = Trim(txtName)
Mid$(recBiaPgm.Memo, 21, 20) = Trim(txtProjet)

cmdOk_Update
End Sub

Private Sub cmdOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdOk
End Sub


'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
PopupMenu mnuPrint
End Sub








Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint
End Sub


Private Sub cmdSuppress_Click()
lstBiaPgmAut_Display
If lstErr.ListCount = 0 Then
    currentMethod = constDelete
    cmdOk_Update
End If
End Sub

Private Sub fgBiaPgm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
fgBiaPgm.Enabled = False
fgBiaPgm_K = (fgBiaPgm.Row) * fgBiaPgm.Cols
mBiaPgm = fgBiaPgm.TextArray(0 + fgBiaPgm_K)
fraBiaPgm_Reset
fraBiaPgm_Display
fraBiaPgm.Enabled = True
fraBiaPgmAut.Enabled = True
'If recBiaPgm.Err = 0 Then
    lstBiaPgmAut_Display
    txtId.Enabled = False
    txtName.SetFocus
    cmdOk.Visible = BiaPgmAut.Valider
    cmdSuppress.Visible = BiaPgmAut.Valider
'Else
'    cmdContext_Quit
'End If

End Sub

Private Sub fgBiaPgm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set fgBiaPgm
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
Dim AMJ As String, X As String
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
Call BiaPgmAut_Init("BiaPgm", BiaPgmAut)

'MDB_Master

currentAMJ = DSys

paramBiaPgm = "BiaPgm"

cmdClear   ' initialisation

fgBiaPgm_FormatString = fgBiaPgm.FormatString
fgBiaPgm_BackColorFixed = fgBiaPgm.BackColorFixed
fgBiaPgm_BackColor = fgBiaPgm.BackColor
Call fgBiaPgm_Display

cmdClear ' début saisie

End Sub

Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub



'---------------------------------------------------------
Public Sub cmdClear()
'---------------------------------------------------------
cmdReset
fraBiaPgm_Reset

Call lstErr_Clear(lstErr, fgBiaPgm, "Sélectionner un programme 'click'")
'fraBiaUsr.Enabled = True
fraBiaPgm.Enabled = False
fraBiaPgmAut.Enabled = False

cmdContext.Caption = constcmdAbandonner
mBiaPgm = ""

End Sub




'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
arrTag_Set False
lstErrClear = True
blnMsgBox_Quit = False
usrColor_Set

rsElpTable_Init recBiaPgm
recBiaPgm.Id = paramBiaPgm
cmdAddNew.Visible = BiaPgmAut.Valider
cmdOk.Visible = BiaPgmAut.Valider
End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate


End Sub

Public Sub Msg_Snd(ByVal X As String)
End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub

Private Sub Form_Unload(Cancel As Integer)
'2003.02.03 MDB_Local

End Sub

Private Sub fraBiaPgm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraBiaPgmAut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub chksaisir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkSaisir
End Sub



Private Sub chkconsulter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkConsulter
End Sub


Private Sub chkvalider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkValider
End Sub

Private Sub chkcomptabiliser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkComptabiliser
End Sub


Private Sub chkrapprocher_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkRapprocher
End Sub


Private Sub chkswift_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkSwift
End Sub


Private Sub frabiausr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Public Sub fgBiaPgm_Display()
Dim X As String
fgBiaPgm.Rows = 1
fgBiaPgm.Clear
fgBiaPgm.FormatString = fgBiaPgm_FormatString
fgBiaPgm.Enabled = True

rsElpTable_Init recBiaPgm

X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & paramBiaPgm & "' order by K1"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    Call rsElpTable_GetBuffer(rsMDB, recBiaPgm)
    fgBiaPgm.Rows = fgBiaPgm.Rows + 1
    fgBiaPgm.Row = fgBiaPgm.Rows - 1
    fgBiaPgm_DisplayItem
    
    rsMDB.MoveNext
Loop

End Sub

Public Sub fgBiaPgm_DisplayItem()
Dim X As String

fgBiaPgm_K = (fgBiaPgm.Row) * fgBiaPgm.Cols
fgBiaPgm.TextArray(0 + fgBiaPgm_K) = recBiaPgm.K1
fgBiaPgm.TextArray(1 + fgBiaPgm_K) = recBiaPgm.Name
fgBiaPgm.TextArray(2 + fgBiaPgm_K) = prtBiaMsg_Aut(mId$(recBiaPgm.Memo, 1, 20))
fgBiaPgm.TextArray(3 + fgBiaPgm_K) = mId$(recBiaPgm.Memo, 21, 20)

End Sub


Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
         If TypeOf xobj Is CommandButton Or TypeOf xobj Is ListBox Or TypeOf xobj Is MSFlexGrid Then
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
        If TypeOf C Is CommandButton Or TypeOf C Is ListBox Or TypeOf C Is MSFlexGrid Then
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            MouseMoveActiveControl.ForeColor = C.ForeColor
             C.ForeColor = MouseMoveUsr.ForeColor
        End If
    End If
End If

End Sub


Public Sub cmdContext_Quit()
If fraBiaPgm.Enabled Then
    lstErr.Clear
    fgBiaPgm.Enabled = True
    fraBiaPgm_Reset
    fraBiaPgmAut.Enabled = False
    fraBiaPgm.Enabled = False
Else
    Unload Me
End If

End Sub

Public Sub cmdContext_Return()
If cmdOk.Visible Then
    cmdOk_Click
Else
    SendKeys "{TAB}"
End If

End Sub

Public Sub fgBiaPgm_Sort()
fgBiaPgm.Row = 1
fgBiaPgm.RowSel = 1 'fgLRAttribut.Rows - 1

fgBiaPgm.Col = 0
fgBiaPgm.ColSel = 1
fgBiaPgm.Sort = flexSortStringAscending

End Sub






Public Sub cmdPrintX(Msg As String)
Dim I As Integer, K As Integer

prtBiaPgm_Open "Liste des programmes de l'application Bia.vbp "
X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & paramBiaPgm & "'" _
    & " order by Id,K1, K2"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    Call rsElpTable_GetBuffer(rsMDB, recBiaPgm)

    prtBiaPgm_Line recBiaPgm, Msg
    If Msg = "1" Then cmdPrint_Usr
    rsMDB.MoveNext
Loop

prtBiaPgm_Close
End Sub


Public Sub fraBiaPgm_Reset()
cmdOk.Visible = False
cmdSuppress.Visible = False

txtId = ""
txtName = ""
txtProjet = ""

chkConsulter.Value = 0
chkSaisir.Value = 0
chkValider.Value = 0
chkComptabiliser.Value = 0
chkRapprocher.Value = 0
chkSwift.Value = 0
chkVirement.Value = 0
chkAvis.Value = 0
chkXspécial.Value = 0
chkXHost.Value = 0

End Sub

Public Sub fraBiaPgm_Display()
Dim X As String
X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & paramBiaPgm & "'" _
    & " and K1 = '" & mBiaPgm & "'"
    
Set rsMDB = cnMDB.Execute(X)
If Not rsMDB.EOF Then
    Call rsElpTable_GetBuffer(rsMDB, recBiaPgm)
    currentMethod = constUpdate
    txtId = Trim(recBiaPgm.K1)
    txtName = Trim(recBiaPgm.Name)
    txtProjet = Trim(mId$(recBiaPgm.Memo, 21, 20))
    chkConsulter.Value = IIf(mId$(recBiaPgm.Memo, 1, 1) = "X", 1, 0)
    chkSaisir.Value = IIf(mId$(recBiaPgm.Memo, 2, 1) = "X", 1, 0)
    chkValider.Value = IIf(mId$(recBiaPgm.Memo, 3, 1) = "X", 1, 0)
    chkComptabiliser.Value = IIf(mId$(recBiaPgm.Memo, 4, 1) = "X", 1, 0)
    chkRapprocher.Value = IIf(mId$(recBiaPgm.Memo, 5, 1) = "X", 1, 0)
    chkSwift.Value = IIf(mId$(recBiaPgm.Memo, 6, 1) = "X", 1, 0)
    chkVirement.Value = IIf(mId$(recBiaPgm.Memo, 7, 1) = "X", 1, 0)
    chkAvis.Value = IIf(mId$(recBiaPgm.Memo, 8, 1) = "X", 1, 0)
    chkXspécial.Value = IIf(mId$(recBiaPgm.Memo, 9, 1) = "X", 1, 0)
    chkXHost.Value = IIf(mId$(recBiaPgm.Memo, 10, 1) = "X", 1, 0)
End If
End Sub

Private Sub mnuPrint0_Click()
cmdPrintX "0"

End Sub

Private Sub mnuPrint1_Click()
cmdPrintX "1"

End Sub

Private Sub txtId_GotFocus()
Call txt_GotFocus(txtId)
End Sub


Private Sub txtId_LostFocus()
Call txt_LostFocus(txtId)
End Sub


Private Sub txtName_GotFocus()
Call txt_GotFocus(txtName)

End Sub


Private Sub txtName_LostFocus()
Call txt_LostFocus(txtName)

End Sub


Private Sub txtProjet_GotFocus()
Call txt_GotFocus(txtProjet)
End Sub

Private Sub txtProjet_LostFocus()
Call txt_LostFocus(txtProjet)
End Sub



Public Sub lstBiaPgmAut_Display()

Dim Nb As Integer, X As String

Nb = 0
lstErr.Clear
lstBiaPgmAut.Clear
rsElpTable_Init recBiaPgmAut
X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & paramBiaPgmAut & "'" _
    & " and K2 = '" & recBiaPgm.K1 & "' order by K1"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    Call rsElpTable_GetBuffer(rsMDB, recBiaPgmAut)
    
   ' If recBiaPgm.K1 = recBiaPgmAut.K2 Then
        lstBiaPgmAut.AddItem recBiaPgmAut.K2 & " / " & recBiaPgmAut.K1 & " autorisé "
        Nb = Nb + 1
   ' End If
    
    rsMDB.MoveNext
Loop
If Nb > 0 Then Call lstErr_AddItem(lstErr, cmdContext, recBiaPgmAut.K2 & " " & Nb & " utilisateurs autorisés ")
End Sub

Public Sub cmdPrint_Usr()
Dim X As String
Dim rsADO As New ADODB.Recordset

X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & paramBiaPgmAut & "'" _
    & " and K2 = '" & recBiaPgm.K1 & "'" _
    & " order by Id,K1"
    
Set rsADO = cnMDB.Execute(X)
Do While Not rsADO.EOF
    Call rsElpTable_GetBuffer(rsADO, recBiaPgmAut)

    prtBiaPgm_Line recBiaPgmAut, "2"
    rsADO.MoveNext
Loop


End Sub
