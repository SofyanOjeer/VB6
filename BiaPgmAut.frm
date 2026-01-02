VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBiaPgmAut 
   AutoRedraw      =   -1  'True
   Caption         =   "Autorisations d'accès aux programmes Bia.vbp"
   ClientHeight    =   9492
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   14316
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9492
   ScaleWidth      =   14316
   Begin VB.CommandButton cmdBiaPgmAut_Size 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Resize Autorisations"
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   0
      Width           =   1695
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   432
      Left            =   8400
      TabIndex        =   19
      Top             =   0
      Width           =   4905
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   0
      TabIndex        =   2
      Top             =   405
      Width           =   6015
      _ExtentX        =   10605
      _ExtentY        =   15896
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Utilisateurs habilités"
      TabPicture(0)   =   "BiaPgmAut.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBiaUsr"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Autorisations"
      TabPicture(1)   =   "BiaPgmAut.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraBiaPgm"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraBiaUsr 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8535
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   5775
         Begin VB.ListBox lstBiaUsr 
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
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   5355
         End
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
         Height          =   8535
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   5775
         Begin VB.CommandButton cmdOk 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Ok"
            Height          =   645
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   5880
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.CommandButton cmdSuppress 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Supprimer"
            Height          =   645
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   5160
            Width           =   1020
         End
         Begin VB.ListBox lstBiaPgm 
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
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   3915
         End
         Begin VB.Frame fraBiaPgmAut 
            Height          =   4320
            Left            =   4080
            TabIndex        =   4
            Top             =   120
            Width           =   1635
            Begin VB.CheckBox chkXHost 
               Caption         =   "Migration : opt autorisée"
               Height          =   492
               Left            =   240
               TabIndex        =   22
               Top             =   3720
               Width           =   1212
            End
            Begin VB.CheckBox chkXspécial 
               Caption         =   "Spécial"
               Height          =   255
               Left            =   240
               TabIndex        =   20
               Top             =   3240
               Width           =   975
            End
            Begin VB.CheckBox chkConsulter 
               Caption         =   "Consulter"
               Height          =   255
               Left            =   240
               TabIndex        =   12
               Top             =   360
               Width           =   1095
            End
            Begin VB.CheckBox chkSaisir 
               Caption         =   "Saisir"
               Height          =   255
               Left            =   240
               TabIndex        =   11
               Top             =   720
               Width           =   975
            End
            Begin VB.CheckBox chkValider 
               Caption         =   "Valider"
               Height          =   255
               Left            =   240
               TabIndex        =   10
               Top             =   1080
               Width           =   975
            End
            Begin VB.CheckBox chkComptabiliser 
               Caption         =   "Comptabiliser"
               Height          =   255
               Left            =   240
               TabIndex        =   9
               Top             =   1440
               Width           =   1335
            End
            Begin VB.CheckBox chkRapprocher 
               Caption         =   "Rapprocher"
               Height          =   255
               Left            =   240
               TabIndex        =   8
               Top             =   1800
               Width           =   1335
            End
            Begin VB.CheckBox chkSwift 
               Caption         =   "Swift"
               Height          =   255
               Left            =   240
               TabIndex        =   7
               Top             =   2160
               Width           =   855
            End
            Begin VB.CheckBox chkVirement 
               Caption         =   "Virement"
               Height          =   255
               Left            =   240
               TabIndex        =   6
               Top             =   2520
               Width           =   975
            End
            Begin VB.CheckBox chkAvis 
               Caption         =   "Avis"
               Height          =   255
               Left            =   240
               TabIndex        =   5
               Top             =   2880
               Width           =   855
            End
         End
      End
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
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   400
      Left            =   13320
      Picture         =   "BiaPgmAut.frx":0038
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   500
   End
   Begin MSFlexGridLib.MSFlexGrid fgBiaPgmAut 
      Height          =   8955
      Left            =   6000
      TabIndex        =   16
      Top             =   480
      Width           =   8235
      _ExtentX        =   14520
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
      FormatString    =   $"BiaPgmAut.frx":013A
   End
End
Attribute VB_Name = "frmBiaPgmAut"
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
Dim BiaPgmAutAut As typeAuthorization
Dim X As String, X1 As String, I As Integer
Dim Msg As String, valX As String, V As Variant

Dim currentMethod As String, currentAMJ As String
Dim fgBiaPgmAut_FormatString As String, fgBiaPgmAut_K As Integer
Dim fgBiaPgmAut_BackColorFixed As Long, fgBiaPgmAut_BackColor As Long

Dim recBiaPgm As typeElpTable, recBiaPgmAut As typeElpTable
Dim lstI As Integer
Dim mBiaPgm As String * 12, mBiaUsr As String * 12

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


Private Sub cmdBiaPgmAut_Size_Click()
If fgBiaPgmAut.Left < 6000 Then
    fgBiaPgmAut.Left = 6120
    SSTab1.Visible = True
Else
    fgBiaPgmAut.Left = 120
    SSTab1.Visible = False
End If
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
    Case constAddNew: V = adoElpTable_AddNew(rsMDB, recBiaPgmAut)
    Case constDelete: V = adoElpTable_Delete(rsMDB, recBiaPgmAut)
    Case constUpdate: V = adoElpTable_Update(rsMDB, recBiaPgmAut)
End Select

'dbElpTable_Update recBiaPgmAut
fgBiaPgmAut_Display
cmdContext_Quit

End Sub

Private Sub cmdOk_Click()
recBiaPgmAut.Memo = Space$(20) & DSys & time_Hms
Mid$(recBiaPgmAut.Memo, 1, 1) = IIf(chkConsulter.Value = 1, "X", " ")
Mid$(recBiaPgmAut.Memo, 2, 1) = IIf(chkSaisir.Value = 1, "X", " ")
Mid$(recBiaPgmAut.Memo, 3, 1) = IIf(chkValider.Value = 1, "X", " ")
Mid$(recBiaPgmAut.Memo, 4, 1) = IIf(chkComptabiliser.Value = 1, "X", " ")
Mid$(recBiaPgmAut.Memo, 5, 1) = IIf(chkRapprocher.Value = 1, "X", " ")
Mid$(recBiaPgmAut.Memo, 6, 1) = IIf(chkSwift.Value = 1, "X", " ")
Mid$(recBiaPgmAut.Memo, 7, 1) = IIf(chkVirement.Value = 1, "X", " ")
Mid$(recBiaPgmAut.Memo, 8, 1) = IIf(chkAvis.Value = 1, "X", " ")
Mid$(recBiaPgmAut.Memo, 9, 1) = IIf(chkXspécial.Value = 1, "X", " ")
Mid$(recBiaPgmAut.Memo, 10, 1) = IIf(chkXHost.Value = 1, "X", " ")

cmdOk_Update
End Sub

Private Sub cmdOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdOk
End Sub


'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
cmdPrintX ""
End Sub








Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint
End Sub


Private Sub cmdSuppress_Click()
currentMethod = constDelete
cmdOk_Update
End Sub

Private Sub fgBiaPgmAut_Click()
Dim X As String
fgBiaPgmAut.Col = 1: X = Trim(fgBiaPgmAut.Text)
Call lst_Scan(X, lstBiaPgm)
If lstBiaPgm.ListIndex > -1 Then lstBiaPgm_Select
End Sub

Private Sub fgBiaPgmAut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set fgBiaPgmAut
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
Call BiaPgmAut_Init("BiaPgm_Aut", BiaPgmAutAut)

'MDB_Master

currentAMJ = DSys

cmdClear   ' initialisation

Call lstZMNURUT0_Load(lstBiaUsr)
'Call lstZMNURUT0_Load_Actif(lstBiaUsr)
Call lst_LoadK1(recBiaPgm.Id, lstBiaPgm)
'Call lstElpTable_Load(lstBiaPgm, recBiaPgm, 0, 1)


fgBiaPgmAut_FormatString = fgBiaPgmAut.FormatString
fgBiaPgmAut_BackColorFixed = fgBiaPgmAut.BackColorFixed
fgBiaPgmAut_BackColor = fgBiaPgmAut.BackColor

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
fgBiaPgmAut.Enabled = False: fgBiaPgmAut.Clear: fgBiaPgmAut.Rows = 1
fraBiaPgmAut_Reset

Call lstErr_Clear(lstErr, fgBiaPgmAut, "Sélectionner un utilisateur 'click'")
fgBiaPgmAut.Enabled = True: lstBiaUsr.Enabled = True
fraBiaPgm.Enabled = False: lstBiaPgm.Enabled = False

fraBiaPgmAut.Enabled = False

cmdContext.Caption = constcmdAbandonner
mBiaUsr = ""
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
rsElpTable_Init recBiaPgmAut
recBiaPgmAut.Id = paramBiaPgmAut
cmdOk.Visible = BiaPgmAutAut.Valider
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


Private Sub lstBiaPgm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lstBiaPgm_Select
End Sub

Private Sub lstBiaUsr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
fraBiaPgm.Enabled = True: lstBiaPgm.Enabled = True
fgBiaPgmAut.Enabled = False: lstBiaUsr.Enabled = False

mBiaUsr = Trim(mId$(lstBiaUsr, 1, 10))
fgBiaPgmAut_Display
fraBiaPgmAut.Enabled = False
SSTab1.Tab = 1
lstBiaPgm.ListIndex = 0
lstBiaUsr.TopIndex = lstBiaUsr.ListIndex
End Sub


Private Sub lstBiaUsr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set lstBiaUsr
'lstI = lstBiaUsr.TopIndex + Fix(y / 195)
'If lstI >= 0 And lstI < lstBiaUsr.ListCount Then
'    lstBiaUsr.ListIndex = lstI
'    lstBiaUsr.ToolTipText = lstBiaUsr.Text
'    If y < 195 And lstBiaUsr.TopIndex > 0 Then lstBiaUsr.TopIndex = lstBiaUsr.TopIndex - 1
'    If y > lstBiaUsr.Height - 195 Then lstBiaUsr.TopIndex = lstBiaUsr.TopIndex + 1
'End If

End Sub


Private Sub lstBiaPgm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set lstBiaPgm
'lstI = lstBiaPgm.TopIndex + Fix(y / 195)
'If lstI >= 0 And lstI < lstBiaPgm.ListCount Then
'    lstBiaPgm.ListIndex = lstI
'    lstBiaPgm.ToolTipText = lstBiaPgm.Text
'    If y < 195 And lstBiaPgm.TopIndex > 0 Then lstBiaPgm.TopIndex = lstBiaPgm.TopIndex - 1
'    If y > lstBiaPgm.Height - 195 Then lstBiaPgm.TopIndex = lstBiaPgm.TopIndex + 1
'End If
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


Public Sub fgBiaPgmAut_Display()
fgBiaPgmAut.Rows = 1
fgBiaPgmAut.Clear
fgBiaPgmAut.FormatString = fgBiaPgmAut_FormatString
fgBiaPgmAut.Enabled = True

X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & paramBiaPgmAut & "'" _
    & " and k1 = '" & mBiaUsr & "' order by K2"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    Call rsElpTable_GetBuffer(rsMDB, recBiaPgmAut)

    fgBiaPgmAut.Rows = fgBiaPgmAut.Rows + 1
    fgBiaPgmAut.Row = fgBiaPgmAut.Rows - 1
    fgBiaPgmAut_DisplayItem
    rsMDB.MoveNext
Loop

End Sub

Public Sub fgBiaPgmAut_DisplayItem()
Dim X As String

fgBiaPgmAut_K = (fgBiaPgmAut.Row) * fgBiaPgmAut.Cols
fgBiaPgmAut.TextArray(0 + fgBiaPgmAut_K) = recBiaPgmAut.K1
fgBiaPgmAut.TextArray(1 + fgBiaPgmAut_K) = recBiaPgmAut.K2
fgBiaPgmAut.TextArray(2 + fgBiaPgmAut_K) = recBiaPgmAut.Name
fgBiaPgmAut.TextArray(3 + fgBiaPgmAut_K) = prtBiaMsg_Aut(mId$(recBiaPgmAut.Memo, 1, 20))
fgBiaPgmAut.TextArray(4 + fgBiaPgmAut_K) = dateImp(mId$(recBiaPgmAut.Memo, 21, 8)) & " " & timeImp(mId$(recBiaPgmAut.Memo, 29, 6))

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
If fraBiaPgmAut.Enabled Then
    fraBiaPgmAut_Reset
    fraBiaPgmAut.Enabled = False
    lstBiaPgm.Enabled = True
Else
    If fraBiaPgm.Enabled Then
        fraBiaPgm.Enabled = False: lstBiaPgm.Enabled = False
        fgBiaPgmAut.Enabled = True: lstBiaUsr.Enabled = True
        SSTab1.Tab = 0
    Else
        If Not lstBiaUsr.Enabled Then
            lstBiaUsr.Enabled = True
            SSTab1.Tab = 0
        Else
            Unload Me
        End If
    End If
End If

End Sub

Public Sub cmdContext_Return()
If cmdOk.Visible Then
    cmdOk_Click
Else
    SendKeys "{TAB}"
End If

End Sub

Public Sub fgBiaPgmAut_Sort()
fgBiaPgmAut.Row = 1
fgBiaPgmAut.RowSel = 1 'fgLRAttribut.Rows - 1

fgBiaPgmAut.Col = 0
fgBiaPgmAut.ColSel = 1
fgBiaPgmAut.Sort = flexSortStringAscending

End Sub






Public Sub cmdPrintX(Msg As String)
Dim I As Integer, K As Integer

prtBiaPgm_Open "Liste des Utilisateurs / Programmes de l'application Bia.vbp "

X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & paramBiaPgmAut & "'" _
    & " order by K1, K2"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    Call rsElpTable_GetBuffer(rsMDB, recBiaPgmAut)

    prtBiaPgm_LineAut recBiaPgmAut
    rsMDB.MoveNext
Loop
prtBiaPgm_Close
End Sub


Public Sub fraBiaPgmAut_Reset()
cmdOk.Visible = False
cmdSuppress.Visible = False
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

Public Sub fraBiaPgmAut_Display()
Dim X As String

X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & paramBiaPgmAut & "'" _
    & " and K1 = '" & mBiaUsr & "'" _
    & " and K2 = '" & mBiaPgm & "'"
    
Set rsMDB = cnMDB.Execute(X)
If Not rsMDB.EOF Then
    Call rsElpTable_GetBuffer(rsMDB, recBiaPgmAut)

    currentMethod = constUpdate
    cmdSuppress.Visible = BiaPgmAutAut.Valider
    
    chkConsulter.Value = IIf(mId$(recBiaPgmAut.Memo, 1, 1) = "X", 1, 0)
    chkSaisir.Value = IIf(mId$(recBiaPgmAut.Memo, 2, 1) = "X", 1, 0)
    chkValider.Value = IIf(mId$(recBiaPgmAut.Memo, 3, 1) = "X", 1, 0)
    chkComptabiliser.Value = IIf(mId$(recBiaPgmAut.Memo, 4, 1) = "X", 1, 0)
    chkRapprocher.Value = IIf(mId$(recBiaPgmAut.Memo, 5, 1) = "X", 1, 0)
    chkSwift.Value = IIf(mId$(recBiaPgmAut.Memo, 6, 1) = "X", 1, 0)
    chkVirement.Value = IIf(mId$(recBiaPgmAut.Memo, 7, 1) = "X", 1, 0)
    chkAvis.Value = IIf(mId$(recBiaPgmAut.Memo, 8, 1) = "X", 1, 0)
    chkXspécial.Value = IIf(mId$(recBiaPgmAut.Memo, 9, 1) = "X", 1, 0)
    chkXHost.Value = IIf(mId$(recBiaPgmAut.Memo, 10, 1) = "X", 1, 0)

Else
    currentMethod = constAddNew
    recBiaPgmAut.Id = paramBiaPgmAut
    recBiaPgmAut.K1 = mBiaUsr
    recBiaPgmAut.K2 = mBiaPgm

    recBiaPgmAut.Memo = recBiaPgm.Memo
    cmdSuppress.Visible = False
End If
recBiaPgmAut.Name = recBiaPgm.Name
cmdOk.Visible = BiaPgmAutAut.Valider
If Not usrSituationCompte_Forçage Then
    If Trim(mBiaUsr) = Trim(usrId) Then cmdOk.Visible = False: cmdSuppress.Visible = False
End If
End Sub

Public Sub fraBiaPgm_Display()
Dim X As String

X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & paramBiaPgm & "'" _
    & " and K1 = '" & mBiaPgm & "' order by K1"
    
Set rsMDB = cnMDB.Execute(X)
If Not rsMDB.EOF Then
    Call rsElpTable_GetBuffer(rsMDB, recBiaPgm)

    chkConsulter.Enabled = IIf(mId$(recBiaPgm.Memo, 1, 1) = " ", False, True)
    chkSaisir.Enabled = IIf(mId$(recBiaPgm.Memo, 2, 1) = " ", False, True)
    chkValider.Enabled = IIf(mId$(recBiaPgm.Memo, 3, 1) = " ", False, True)
    chkComptabiliser.Enabled = IIf(mId$(recBiaPgm.Memo, 4, 1) = " ", False, True)
    chkRapprocher.Enabled = IIf(mId$(recBiaPgm.Memo, 5, 1) = " ", False, True)
    chkSwift.Enabled = IIf(mId$(recBiaPgm.Memo, 6, 1) = " ", False, True)
    chkVirement.Enabled = IIf(mId$(recBiaPgm.Memo, 7, 1) = " ", False, True)
    chkAvis.Enabled = IIf(mId$(recBiaPgm.Memo, 8, 1) = " ", False, True)
    chkXspécial.Enabled = IIf(mId$(recBiaPgm.Memo, 9, 1) = " ", False, True)
    chkXHost.Enabled = IIf(mId$(recBiaPgm.Memo, 10, 1) = " ", False, True)
End If
End Sub


Public Sub lstBiaPgm_Select()
Dim K As Integer
lstBiaPgm.Enabled = False
fraBiaPgmAut_Reset
K = InStr(1, lstBiaPgm, vbTab)
mBiaPgm = mId$(lstBiaPgm, 1, K - 1)
fraBiaPgm_Display
fraBiaPgmAut.Enabled = True
'If recBiaPgm.Err = 0 Then
    fraBiaPgmAut_Display
    lstBiaPgm.TopIndex = lstBiaPgm.ListIndex
'Else
'    cmdContext_Quit
'End If

End Sub
