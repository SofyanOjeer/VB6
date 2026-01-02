VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCompteModif 
   AutoRedraw      =   -1  'True
   Caption         =   "Compte : modification référentiel"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6375
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00C0C0C0&
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
      Height          =   400
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2700
      Left            =   0
      TabIndex        =   9
      Top             =   400
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4763
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Sélection des comptes"
      TabPicture(0)   =   "CompteModif.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraUpdate"
      Tab(0).Control(1)=   "fraFilter"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Modification des attributs"
      TabPicture(1)   =   "CompteModif.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "libCompte"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "libIntitulé"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "libExtrait_P"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "libExtrait_C"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fraExtrait_P"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "fraExtrait_C"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdOk"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ok"
         Height          =   645
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1920
         Width           =   1020
      End
      Begin VB.Frame fraUpdate 
         Caption         =   "Mise à jour"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   -69360
         TabIndex        =   16
         Top             =   600
         Width           =   3495
         Begin VB.OptionButton optUpdateOne 
            Caption         =   "Compte sélectionné"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   700
            Width           =   1695
         End
         Begin VB.OptionButton optUpdateAll 
            Caption         =   "Tous les comptes"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.Frame fraExtrait_C 
         Caption         =   "Code retenue du courrier"
         Height          =   2200
         Left            =   6480
         TabIndex        =   14
         Top             =   360
         Width           =   2775
         Begin VB.ListBox lstExtrait_C 
            Height          =   1815
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame fraExtrait_P 
         Caption         =   "Code Extrait"
         Height          =   2200
         Left            =   3600
         TabIndex        =   12
         Top             =   360
         Width           =   2775
         Begin VB.ListBox lstExtrait_P 
            Height          =   1815
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame fraFilter 
         Caption         =   "Critères de sélection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   -74880
         TabIndex        =   10
         Top             =   540
         Width           =   4815
         Begin VB.OptionButton optExtrait 
            Caption         =   "code extrait"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox txtContext 
            Height          =   285
            Left            =   2520
            TabIndex        =   0
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton optBiaTyp 
            Caption         =   "Type de compte"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton optRacine 
            Caption         =   "Racine"
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   700
            Width           =   1695
         End
         Begin VB.OptionButton optCompte 
            Caption         =   "Compte"
            Height          =   255
            Left            =   240
            TabIndex        =   1
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Label libExtrait_C 
         Caption         =   "-"
         Height          =   255
         Left            =   300
         TabIndex        =   20
         Top             =   2295
         Width           =   2055
      End
      Begin VB.Label libExtrait_P 
         Caption         =   "-"
         Height          =   255
         Left            =   300
         TabIndex        =   19
         Top             =   1905
         Width           =   2055
      End
      Begin VB.Label libIntitulé 
         Caption         =   "-"
         Height          =   495
         Left            =   300
         TabIndex        =   18
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label libCompte 
         Caption         =   "-"
         Height          =   255
         Left            =   300
         TabIndex        =   17
         Top             =   700
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   400
      Left            =   8880
      Picture         =   "CompteModif.frx":0038
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   500
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   0
      Width           =   2500
   End
   Begin MSFlexGridLib.MSFlexGrid fgCompteModif 
      Height          =   3210
      Left            =   0
      TabIndex        =   11
      Top             =   3120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   5662
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      RowHeightMin    =   350
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
      FormatString    =   $"CompteModif.frx":013A
   End
End
Attribute VB_Name = "frmCompteModif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim CompteModifAut As typeAuthorization
Dim X As String, X1 As String, I As Integer
Dim Msg As String, valX As String, V As Variant

Dim recCompteModif As typeCompteModif
Dim currentMethod As String, currentAMJ As String
Dim fgCompteModif_FormatString As String, fgCompteModif_K As Integer
Dim fgCompteModif_BackColorFixed As Long, fgCompteModif_BackColor As Long

Dim recExtrait_P As typeElpTable
Dim recExtrait_C As typeElpTable

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
Dim X As String, oldK1 As String, newK1 As String

X = fgCompteModif.TextArray(6 + fgCompteModif_K)

If lstExtrait_P.ListIndex >= 0 And lstExtrait_P.ListIndex < lstExtrait_P.ListCount Then
    newK1 = mId$(lstExtrait_P, 1, 1)
    oldK1 = mId$(X, 15, 1)
    If oldK1 <> newK1 Then
        recExtrait_P.K1 = newK1
        recExtrait_P.Method = "Seek="
        dbElpTable_ReadE recExtrait_P
        fgCompteModif.TextArray(4 + fgCompteModif_K) = Trim(newK1 & " : " & recExtrait_P.Name)
        Mid$(X, 15, 1) = newK1
        fgCompteModif.Col = 4: fgCompteModif.CellBackColor = focusUsr.BackColor
    End If
End If

 If lstExtrait_C.ListIndex >= 0 And lstExtrait_C.ListIndex < lstExtrait_C.ListCount Then
    newK1 = mId$(lstExtrait_C, 1, 1)
    oldK1 = mId$(X, 16, 1)
    If oldK1 <> newK1 Then
        recExtrait_C.K1 = newK1
        recExtrait_C.Method = "Seek="
        dbElpTable_ReadE recExtrait_C
        fgCompteModif.TextArray(5 + fgCompteModif_K) = Trim(newK1 & " : " & recExtrait_C.Name)
        Mid$(X, 16, 1) = mId$(recExtrait_C.K1, 1, 1)
        fgCompteModif.Col = 5: fgCompteModif.CellBackColor = focusUsr.BackColor
    End If
End If

fgCompteModif.TextArray(6 + fgCompteModif_K) = X

srvCompteModif_Init recCompteModif
recCompteModif.Method = constUpdate
recCompteModif.Société = SocId$
recCompteModif.Agence = SocAgence$
recCompteModif.Devise = mId$(X, 1, 3)
recCompteModif.Numéro = mId$(X, 4, 11)
recCompteModif.Extrait = mId$(X, 15, 1)
recCompteModif.Courrier = mId$(X, 16, 1)
srvCompteModif_Update recCompteModif
End Sub

Private Sub cmdOk_Click()
Dim K As Integer

If optUpdateOne Then
    cmdOk_Update
Else
    For K = 1 To fgCompteModif.Rows - 1
        fgCompteModif.Row = K
        fgCompteModif_K = fgCompteModif.Row * fgCompteModif.Cols
        cmdOk_Update
    Next K
End If


cmdOK_Off

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


Private Sub fgCompteModif_Click()
lstErr.Clear
fgCompteModif_K = fgCompteModif.Row * fgCompteModif.Cols
If fgCompteModif.Row > 0 Then
    libCompte = fgCompteModif.TextArray(0 + fgCompteModif_K) & " " & fgCompteModif.TextArray(1 + fgCompteModif_K)
    libIntitulé = fgCompteModif.TextArray(2 + fgCompteModif_K)
    libExtrait_P = fgCompteModif.TextArray(4 + fgCompteModif_K)
    libExtrait_C = fgCompteModif.TextArray(5 + fgCompteModif_K)
    SSTab1.Tab = 1
    cmdOK_Off
End If

'fgCompteModif.Col = 1: fgCompteModif.CellBackColor = focusUsr.BackColor
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
Call BiaPgmAut_Init("COMPTE_MOD", CompteModifAut)
currentAMJ = DSys
ReDim arrCompteModif(10): arrCompteModifNbMax = 10
ReDim arrExtrait_P(10)
ReDim arrExtrait_C(10)

cmdClear   ' initialisation

srvCompteModif_Init recCompteModif

Call lstElpTable_Load(lstExtrait_P, recExtrait_P, 0, 1)

Call lstElpTable_Load(lstExtrait_C, recExtrait_C, 0, 1)

fgCompteModif_FormatString = fgCompteModif.FormatString
fgCompteModif_BackColorFixed = fgCompteModif.BackColorFixed
fgCompteModif_BackColor = fgCompteModif.BackColor

cmdClear ' début saisie

End Sub



'---------------------------------------------------------
Public Sub cmdClear()
'---------------------------------------------------------
cmdReset
cmdOK_Off
SSTab1.Tab = 0
fgCompteModif.Enabled = False: fgCompteModif.Clear: fgCompteModif.Rows = 1

Call lstErr_Clear(lstErr, fgCompteModif, "Sélection des comptes 'click'")
fraFilter.Enabled = True
fraUpdate.Enabled = False
fraExtrait_P.Enabled = False
fraExtrait_C.Enabled = False
''optCompte = True
optUpdateOne = True

cmdContext.Caption = constcmdRechercher


End Sub




'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
arrTag_Set False
lstErrClear = True
blnMsgBox_Quit = False
usrColor_Set

recElpTable_Init recExtrait_P
recExtrait_P.Id = "Extrait_P"
recElpTable_Init recExtrait_C
recExtrait_C.Id = "Extrait_C"

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




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub

Private Sub fraExtrait_C_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraExtrait_P_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub lstExtrait_C_Click()
cmdOk.Visible = True

End Sub

Private Sub lstExtrait_C_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set lstExtrait_C
End Sub


Private Sub lstExtrait_P_Click()
cmdOk.Visible = True

End Sub


Private Sub lstExtrait_P_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set lstExtrait_P
End Sub


Private Sub optBiaTyp_Click()
txtContext.SetFocus
End Sub

Private Sub optBiaTyp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optBiaTyp
End Sub


Private Sub optCompte_Click()
txtContext.SetFocus
End Sub

Private Sub optCompte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optCompte
End Sub


Private Sub optExtrait_Click()
txtContext.SetFocus

End Sub

Private Sub optExtrait_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optExtrait
End Sub

Private Sub optRacine_Click()
txtContext.SetFocus
End Sub

Private Sub optRacine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optRacine
End Sub


Private Sub optUpdateAll_Click()
fgCompteModif.Enabled = False
libIntitulé = "Tous les comptes sont sélectionnés"
libCompte = ""
libExtrait_P = ""
libExtrait_C = ""
cmdOK_Off
SSTab1.Tab = 1
End Sub

Private Sub optUpdateAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optUpdateAll
End Sub


Private Sub optUpdateOne_Click()
fgCompteModif.Enabled = CompteModifAut.Valider
fgCompteModif_Click
End Sub


Private Sub optUpdateOne_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optUpdateOne
End Sub


Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set SSTab1
End Sub


Private Sub txtContext_GotFocus()
Call txt_GotFocus(txtContext)
End Sub

Private Sub txtContext_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtContext_LostFocus()
Call txt_LostFocus(txtContext)
End Sub



Public Sub fgCompteModif_Display()
fgCompteModif.Rows = 1
fgCompteModif.Clear
fgCompteModif.FormatString = fgCompteModif_FormatString
fgCompteModif.Enabled = True
For arrCompteModifIndex = 1 To arrCompteModifNb
    If arrCompteModif(arrCompteModifIndex).Method <> constDelete _
    And arrCompteModif(arrCompteModifIndex).Method <> constIgnore Then
        recCompteModif = arrCompteModif(arrCompteModifIndex)
        fgCompteModif.Rows = fgCompteModif.Rows + 1
        fgCompteModif.Row = fgCompteModif.Rows - 1
        fgCompteModif_DisplayItem
    End If
Next arrCompteModifIndex
'''If fgCompteModif.Rows > 1 Then fgCompteModif_Sort

End Sub

Public Sub fgCompteModif_DisplayItem()
Dim X As String

fgCompteModif_K = (fgCompteModif.Row) * fgCompteModif.Cols
fgCompteModif.TextArray(0 + fgCompteModif_K) = Format$(recCompteModif.Devise, "@@@")
fgCompteModif.TextArray(1 + fgCompteModif_K) = Compte_Imp(recCompteModif.Numéro)
fgCompteModif.TextArray(2 + fgCompteModif_K) = Trim(recCompteModif.Intitulé) & Trim(recCompteModif.Intitulé2)

Select Case recCompteModif.Situation
    Case " ": X = ""
    Case "A": X = "Annulé"
    Case "B ": X = "Bloqué"
    Case Else: X = "? " & recCompteModif.Situation
End Select
fgCompteModif.TextArray(3 + fgCompteModif_K) = X

recExtrait_P.K1 = recCompteModif.Extrait
recExtrait_P.Method = "Seek="
dbElpTable_ReadE recExtrait_P
fgCompteModif.TextArray(4 + fgCompteModif_K) = Trim(recCompteModif.Extrait & " : " & recExtrait_P.Name)

recExtrait_C.K1 = recCompteModif.Courrier
recExtrait_C.Method = "Seek="
dbElpTable_ReadE recExtrait_C
fgCompteModif.TextArray(5 + fgCompteModif_K) = Trim(recCompteModif.Courrier & " : " & recExtrait_C.Name)

fgCompteModif.TextArray(6 + fgCompteModif_K) = recCompteModif.Devise & recCompteModif.Numéro & recCompteModif.Extrait & recCompteModif.Courrier

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
If fraFilter.Enabled Then
    Unload Me
Else
    If blnMsgBox_Quit Then
        X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
     Else
        X = vbYes
     End If
     If X = vbYes Then cmdClear
End If

End Sub

Public Sub cmdContext_Return()
If fraFilter.Enabled Then
    fgCompteModif_Load
Else
    SendKeys "{TAB}"
End If

End Sub

Public Sub fgCompteModif_Sort()
fgCompteModif.Row = 1
fgCompteModif.RowSel = 1 'fgLRAttribut.Rows - 1

fgCompteModif.Col = 0
fgCompteModif.ColSel = 1
fgCompteModif.Sort = flexSortStringAscending

End Sub






Public Sub cmdPrintX(Msg As String)
Dim I As Integer, K As Integer

prtCompteModif_Open

For I = 1 To fgCompteModif.Rows - 1
    fgCompteModif_K = I * fgCompteModif.Cols
    Call prtCompteModif_Line(fgCompteModif.TextArray(0 + fgCompteModif_K) _
                              , fgCompteModif.TextArray(1 + fgCompteModif_K) _
                              , fgCompteModif.TextArray(2 + fgCompteModif_K) _
                              , fgCompteModif.TextArray(3 + fgCompteModif_K) _
                              , fgCompteModif.TextArray(4 + fgCompteModif_K) _
                              , fgCompteModif.TextArray(5 + fgCompteModif_K))

Next I

prtCompteModif_Close
End Sub


Public Sub fgCompteModif_Load()
X = num_Control(txtContext, valX, 11, 0)
If X <> "" Then
    Call lstErr_Clear(lstErr, txtContext, "? " & X)
Else
    If valX = 0 And Not optExtrait Then
        Call lstErr_Clear(lstErr, txtContext, "? Préciser le compte")
    Else
        srvCompteModif_Init recCompteModif
        recCompteModif.Method = currentMethod
        recCompteModif.Société = SocId$
        recCompteModif.Agence = SocAgence$
        If optCompte Then
            currentMethod = "SnapL5"
            recCompteModif.Numéro = valX
            recCompteModif.Devise = "000"
            arrCompteModif(0) = recCompteModif
            arrCompteModif(0).Devise = "999"
        End If
         If optRacine Then
            currentMethod = "SnapL5"
            recCompteModif.Numéro = (valX Mod 100000) & "000000"
            recCompteModif.Devise = "000"
            arrCompteModif(0) = recCompteModif
            arrCompteModif(0).Numéro = (valX Mod 100000) & "999999"
            arrCompteModif(0).Devise = "999"
        End If
          If optBiaTyp Then
            currentMethod = "SnapLA"
            recCompteModif.Numéro = "00000" & (valX Mod 1000) & "000"
            arrCompteModif(0) = recCompteModif
            arrCompteModif(0).Numéro = "99999" & (valX Mod 1000) & "999"
        End If
          If optExtrait Then
            currentMethod = "SnapKE"
            recCompteModif.Extrait = (valX Mod 10)
            recCompteModif.Devise = "000"
            recCompteModif.Numéro = "00000000000"
            arrCompteModif(0) = recCompteModif
            arrCompteModif(0).Numéro = "99999999999"
            arrCompteModif(0).Devise = "999"
        End If
        recCompteModif.Method = currentMethod
        arrCompteModifNb = 0
        arrCompteModifsuite = True
        
        Do Until Not arrCompteModifsuite
           
            V = srvCompteModif_Mon(recCompteModif)
            recCompteModif = arrCompteModif(arrCompteModifNb)
            recCompteModif.Method = currentMethod & "+"
            
        Loop
                
        fgCompteModif_Display
        If arrCompteModifNb = 0 Then
            Call lstErr_Clear(lstErr, txtContext, "? Aucune sélection")
        Else
            cmdContext.Caption = constcmdAbandonner
            fraFilter.Enabled = False
            fraUpdate.Enabled = CompteModifAut.Valider
            fraExtrait_P.Enabled = CompteModifAut.Valider
            fraExtrait_C.Enabled = CompteModifAut.Valider
            fgCompteModif.Row = 1
            fgCompteModif_Click
            Call lstErr_Clear(lstErr, txtContext, arrCompteModifNb & " comptes")
        End If
End If
End If
End Sub

Public Sub cmdOK_Off()
lstExtrait_P = -1
lstExtrait_C = -1
cmdOk.Visible = False

End Sub
