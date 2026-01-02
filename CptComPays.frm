VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCptComPays 
   Caption         =   "ComPays : mise à jour du code pays (libellé comptable)"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   9420
   Begin VB.Frame fraComPays 
      Caption         =   "Modification du libellé coimptable "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Width           =   6495
      Begin VB.CommandButton cmdComPays 
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
         Height          =   645
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtComPays 
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.Label libComPays 
         Caption         =   "-"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblComPays 
         Caption         =   "code pays : "
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   6240
      TabIndex        =   2
      Top             =   0
      Width           =   2745
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   9000
      Picture         =   "CptComPays.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid fgSelect 
      Height          =   6210
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   10954
      _Version        =   393216
      Rows            =   1
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
      AllowUserResizing=   3
      FormatString    =   "<Dossier                      |< Code pays "
   End
   Begin MSFlexGridLib.MSFlexGrid fgMvt 
      Height          =   4890
      Left            =   2760
      TabIndex        =   4
      Top             =   1800
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   8625
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   350
      BackColor       =   14737632
      ForeColor       =   12582912
      ForeColorFixed  =   -2147483641
      BackColorSel    =   12648384
      BackColorBkg    =   14737632
      AllowBigSelection=   0   'False
      Enabled         =   0   'False
      TextStyleFixed  =   4
      FocusRect       =   2
      HighLight       =   0
      GridLines       =   2
      AllowUserResizing=   3
      FormatString    =   $"CptComPays.frx":0102
   End
   Begin VB.Menu mnuComPays 
      Caption         =   "mnuComPays"
      Visible         =   0   'False
      Begin VB.Menu mnuComPays001 
         Caption         =   "001_France "
      End
      Begin VB.Menu mnuComPays208 
         Caption         =   "208_Algérie "
      End
      Begin VB.Menu mnuComPays216 
         Caption         =   "216_Libye"
      End
      Begin VB.Menu mnuComPaysAutre 
         Caption         =   "autre"
      End
   End
End
Attribute VB_Name = "frmCptComPays"
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
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim CptComPaysAut As typeAuthorization

Dim cmdImport_Select_Nb As Long, cmdImport_Nb As Long

Dim optEtat As String * 1, SrvCptP0_Amj As String * 8, SrvMvtP0_Amj As String * 8, SrvCptP0_Amj_Ok As String
Dim recElpTable As typeElpTable

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer


Dim fgMvt_FormatString As String, fgMvt_K As Integer
Dim fgMvt_RowDisplay As Integer, fgMvt_RowClick As Integer
Dim fgMvt_ColorClick As Long, fgMvt_ColorDisplay As Long
Dim fgMvt_Sort1 As Integer, fgMvt_Sort2 As Integer

Dim mId15 As String, mComPays As String * 3

Private Sub cmdImport_MvtP0()
Dim xInput As String, blnOk As Boolean
Dim vReturn As Variant
On Error Resume Next

Dim I As Long
blnOk = False
cmdImport_Select_Nb = 0: cmdImport_Nb = 0: I = 0
MDB.Execute "delete * from MvtP0"

X = Dir(paramComptaExt_Mvt_Import)
If X = "" Then Call lstErr_Clear(lstErr, cmdPrint, "? Le fichier des mouvements n'existe pas"): Exit Sub

Call lstErr_AddItem(lstErr, cmdPrint, "Chargement des mouvements, tri ...")

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

    cmdImport_Nb = cmdImport_Nb + 1: I = I + 1
 '   vReturn = mId$(xInput, 7, 14) & mId$(xInput, 163, 8) & mId$(xInput, 71, 11) ' devise compte date trt no pièce no ligne
    If mId$(xInput, 25, 3) = "010" And mId$(xInput, 86, 3) = "000" Then
        recMvtp0.Id = mId$(xInput, 86, 15) & Format$(cmdImport_Nb, "0000000")  ' devise compte date trt no pièce no ligne
        recMvtp0.Text = xInput
            cmdImport_Select_Nb = cmdImport_Select_Nb + 1
            dbMvtP0_Update recMvtp0
    End If
    If I = 1000 Then I = 0: Call lstErr_ChangeLastItem(lstErr, cmdPrint, "Sélection des mouvements: " & cmdImport_Select_Nb & " / " & cmdImport_Nb): DoEvents
 
Loop

Close
''mdbMvtP0.tableMvtP0_Close

If Not blnOk Then
    cmdImport_Select_Nb = 0
    Call MsgBox("erreur : manque fin de fichier ", vbCritical, "frmCompteEXtrait : cmdImport_Cptp0 :SrvMvtP0 ")
End If

End Sub




'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False

fgMvt.Enabled = False
fraComPays.Enabled = False
txtComPays = ""

End Sub

Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSelect.Row

If lRow > 0 Then
    fgSelect.Row = lRow
    For I = 0 To 1
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = 0 To 1
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.Col = 0
    End If
End If

End Sub

Private Sub fgSelect_Display()
Dim K2 As Integer, I As Integer

fgSelect.Visible = True
fgSelect.Clear: fgSelect.Rows = 1: fgSelect_RowDisplay = 0

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Enabled = True

recMvtp0.Method = "MoveFirst"

V = dbMvtP0_ReadE(recMvtp0)
mId15 = ""
Do While recMvtp0.Err = 0
    If mId15 <> mId$(recMvtp0.Text, 86, 15) Then
        mId15 = mId$(recMvtp0.Text, 86, 15)
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine
    End If
    
    recMvtp0.Method = "MoveNext    "
    recMvtp0.Err = tableMvtP0_Read(recMvtp0)
Loop

'If fgSelect.Rows = 1 Then Exit Sub
'fgSelect_Sort

End Sub
Private Sub fgMvt_Display()
Dim K2 As Integer, I As Integer

fgMvt.Visible = True
fgMvt.Clear: fgMvt.Rows = 1: fgMvt_RowDisplay = 0

fgMvt.Rows = 1
fgMvt.FormatString = fgMvt_FormatString
fgMvt.Enabled = True

recMvtp0.Method = "Seek>="
recMvtp0.Id = mId15
V = dbMvtP0_ReadE(recMvtp0)
Do While recMvtp0.Err = 0
    If mId15 = mId$(recMvtp0.Id, 1, 15) Then
        fgMvt.Rows = fgMvt.Rows + 1
        fgMvt.Row = fgMvt.Rows - 1
        fgMvt_DisplayLine
    Else
        recMvtp0.Err = 9999
    End If
    
    recMvtp0.Method = "MoveNext    "
    recMvtp0.Err = tableMvtP0_Read(recMvtp0)
Loop

'If fgMvt.Rows = 1 Then Exit Sub
'fgMvt_Sort

End Sub

Public Sub fgSelect_DisplayLine()
fgSelect_K = (fgSelect.Row) * fgSelect.Cols
fgSelect.TextArray(0 + fgSelect_K) = mId15
fgSelect.TextArray(1 + fgSelect_K) = mId$(recMvtp0.Text, 86, 3)

End Sub

Public Sub fgMvt_DisplayLine()
fgMvt_K = (fgMvt.Row) * fgMvt.Cols
fgMvt.TextArray(0 + fgMvt_K) = dateImp(mId$(recMvtp0.Text, 163, 8))
fgMvt.TextArray(1 + fgMvt_K) = Val(mId$(recMvtp0.Text, 71, 7)) & "." & Val(mId$(recMvtp0.Text, 78, 4))
fgMvt.TextArray(2 + fgMvt_K) = Compte_Display(mId$(recMvtp0.Text, 10, 11))
fgMvt.TextArray(3 + fgMvt_K) = CCur(Val(mId$(recMvtp0.Text, 28, 19)))
fgMvt.TextArray(4 + fgMvt_K) = mId$(recMvtp0.Text, 86, 50)

End Sub

Public Sub fgSelect_Sort()
If fgSelect.Rows > 1 Then
    fgSelect.Row = 1
    fgSelect.RowSel = fgSelect.Rows - 1
    
    fgSelect.Col = fgSelect_Sort1
    fgSelect.ColSel = fgSelect_Sort2
    fgSelect.Sort = 1
End If

End Sub



Public Sub Form_Init()
cmdReset

fgMvt_FormatString = fgMvt.FormatString
fgMvt_Sort1 = 0: fgMvt_Sort2 = 0
fgSelect_FormatString = fgSelect.FormatString
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0

recElpTable_Init recElpTable
recElpTable.Id = "Param"
recElpTable.K1 = "ComptaExt"
recElpTable.Method = "Seek="

recElpTable.K2 = "Mvt_Import"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramComptaExt_Mvt_Import = paramServer(recElpTable.Memo)
Call lstErr_AddItem(lstErr, cmdContext, "Cpt_Import:" & paramComptaExt_Mvt_Import)
'test   paramComptaExt_Mvt_Import = "S:\FTP\srvmvtp0_200007"
'test   MsgBox paramComptaExt_Mvt_Import, vbInformation, "form_init : Test"
Exit Sub

Table_Error:
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Table", vbCritical, "frmCptComPays.Form_Init"
Exit Sub

Memo_Error:
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "frmCptComPays.Form_Init"
Exit Sub

Num_Error:
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : " & recElpTable.Memo & " :Mémo non numérique", vbCritical, "TfluxEspèces_Param_Init"
End Sub
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

End Sub




Public Sub Msg_Snd(ByVal X As String)
End Sub

'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
currentActiveControl_Name = C.Name
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
End Sub


'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub


Private Sub cmdComPays_Click()
lstErr.Clear
txtComPays_Control
If lstErr.ListCount = 0 Then
    cmdUpdate
    fgMvt_Display
End If
End Sub

Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Dim I As Integer, X1 As String, X2 As String
prtCptComPays_Open "   "


For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = 0: X1 = fgSelect.Text
    fgSelect.Col = 1: X2 = Trim(fgSelect.Text)
    prtCptComPays_Line X1, X2
Next I

prtCptComPays_Close
'Me.PopupMenu mnucmdPrint, vbPopupMenuLeftButton
End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xStatut As String
If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_Sort
        Case 1: fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2, 15: fgSelect_Sort1 = 15: fgSelect_Sort2 = 15: fgSelect_Sort
        Case 3, 13: fgSelect_Sort1 = 13: fgSelect_Sort2 = 13: fgSelect_Sort
        Case 4, 14: fgSelect_Sort1 = 14: fgSelect_Sort2 = 14: fgSelect_Sort
        Case 5: fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_Sort
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
        Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
        Case 9: fgSelect_Sort1 = 9: fgSelect_Sort2 = 9: fgSelect_Sort
        Case 10: fgSelect_Sort1 = 10: fgSelect_Sort2 = 12: fgSelect_Sort
        Case 11: fgSelect_Sort1 = 11: fgSelect_Sort2 = 12: fgSelect_Sort
        Case 12: fgSelect_Sort1 = 12: fgSelect_Sort2 = 12: fgSelect_Sort
        Case 16: fgSelect_Sort1 = 16: fgSelect_Sort2 = 12: fgSelect_Sort
    End Select
Else

    fgSelect_K = fgSelect.Row * fgSelect.Cols
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        mId15 = mId$(fgSelect.TextArray(0 + fgSelect_K), 1, 15)
        Call fgMvt_Display
    
        Me.PopupMenu mnuComPays, vbPopupMenuLeftButton
    End If
End If

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
blnControl = False
Form_Init

cmdImport_MvtP0
fgSelect_Display
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub

Private Sub mnuAbandonner_Click()
cmdContext_Quit
End Sub


Private Sub mnuQuitter_Click()
Unload Me
End Sub

Public Sub cmdContext_Quit()
cmdReset
blnControl = False
If fraComPays.Enabled Then
    cmdContext.Caption = constcmdRechercher
    cmdReset
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
SendKeys "{TAB}"

End Sub


Private Sub mnuComPays001_Click()
txtComPays = "001"
cmdComPays_Click
End Sub


Private Sub mnuComPays208_Click()
txtComPays = "208"
cmdComPays_Click
End Sub


Private Sub mnuComPays216_Click()
txtComPays = "216"
cmdComPays_Click
End Sub


Private Sub mnuComPaysAutre_Click()
fraComPays.Enabled = True
txtComPays.SetFocus
End Sub


Private Sub txtComPays_GotFocus()
txt_GotFocus txtComPays

End Sub

Public Sub txtComPays_Control()
Dim X As String
If Val(txtComPays) = 0 Then Call lstErr_AddItem(lstErr, txtComPays, "?Pays de la commission "): Exit Sub
X = DicLib("019", "0" & Trim(txtComPays))
If X = "-" Then libComPays = "?": Call lstErr_AddItem(lstErr, txtComPays, "?Pays de la commission "): Exit Sub
libComPays = X
mComPays = Format$(Val(txtComPays), "000")

End Sub


Private Sub txtComPays_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub


Private Sub txtComPays_LostFocus()
txt_LostFocus txtComPays

End Sub



Public Sub cmdUpdate()
Dim iReturn As Integer, wMsg As String
Dim recCptMvt As typeCptMvt
recMvtp0.Method = "Seek>="
recMvtp0.Id = mId15
V = dbMvtP0_ReadE(recMvtp0)
Do While recMvtp0.Err = 0
    If mId15 = mId$(recMvtp0.Text, 86, 15) Then
        Mid$(recMvtp0.Text, 86, 3) = mComPays
        recMvtp0.Method = "Update"
        iReturn = tableMvtP0_Update(recMvtp0)
        If iReturn <> 0 Then Call MsgBox("Erreur update recMvtP0 : " & iReturn, vbCritical, " frmCompays_cmdUpdate")
        wMsg = Space$(recCptMvtLen)
        Mid$(wMsg, 1, 12) = "SRVCPTMVT"
        Mid$(wMsg, 13, 12) = "Update_Lib"
        Mid$(wMsg, 35, memoCptMvtLen) = recMvtp0.Text 'mId$(recMvtp0.Text, 1, recMvtp0.Text)
        V = srvCptMvt_UpdateBuffer(wMsg)
        If Not IsNull(V) Then Call MsgBox("Erreur update AS400 / HisMvtLA : " & V, vbCritical, " frmCompays_cmdUpdate")
    Else
        recMvtp0.Err = 9999
    End If
    
    recMvtp0.Method = "MoveNext    "
    recMvtp0.Err = tableMvtP0_Read(recMvtp0)
Loop

End Sub
