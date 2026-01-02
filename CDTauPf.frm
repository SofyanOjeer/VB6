VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCDTauPf 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "CDTAUPF : mise à jour des taux de commission"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   Icon            =   "CDTauPf.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   9420
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   5400
      TabIndex        =   7
      Top             =   0
      Width           =   3500
   End
   Begin TabDlg.SSTab sstab1 
      Height          =   6495
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Sélection"
      TabPicture(0)   =   "CDTauPf.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Liste des taux de commission"
      TabPicture(1)   =   "CDTauPf.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fgSelect"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraUpdate"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fraUpdate 
         Caption         =   "Modification d'une période"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3060
         Left            =   240
         TabIndex        =   11
         Top             =   3360
         Width           =   9015
         Begin VB.Frame fraUpdatePériode 
            Height          =   2055
            Left            =   4680
            TabIndex        =   20
            Top             =   240
            Width           =   4215
            Begin VB.Frame fraUpdateK 
               Height          =   2055
               Left            =   0
               TabIndex        =   27
               Top             =   0
               Width           =   2775
               Begin VB.OptionButton optTACODC_ZZ 
                  Caption         =   "ZZ"
                  Height          =   255
                  Left            =   1800
                  TabIndex        =   42
                  Top             =   1680
                  Width           =   615
               End
               Begin VB.OptionButton optTACODC_CK 
                  Caption         =   "CK"
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   31
                  Top             =   960
                  Width           =   615
               End
               Begin VB.OptionButton optTACODC_CN 
                  Caption         =   "CN"
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   30
                  Top             =   1320
                  Width           =   615
               End
               Begin VB.OptionButton optTACODC_CH 
                  Caption         =   "CH"
                  Height          =   255
                  Left            =   1800
                  TabIndex        =   29
                  Top             =   960
                  Width           =   615
               End
               Begin VB.OptionButton optTACODC_CF 
                  Caption         =   "CF"
                  Height          =   255
                  Left            =   1800
                  TabIndex        =   28
                  Top             =   1200
                  Width           =   615
               End
               Begin MSComCtl2.DTPicker txtTADEFF 
                  Height          =   300
                  Left            =   1200
                  TabIndex        =   32
                  Top             =   360
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
                  Format          =   28180483
                  CurrentDate     =   36299
                  MaxDate         =   401768
                  MinDate         =   -328351
               End
               Begin VB.Label lblTAFEFF 
                  Caption         =   "période :"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   34
                  Top             =   480
                  Width           =   735
               End
               Begin VB.Label Label1 
                  Caption         =   "Code commission"
                  Height          =   615
                  Left            =   120
                  TabIndex        =   33
                  Top             =   1080
                  Width           =   855
               End
            End
            Begin MSComCtl2.DTPicker txtTAFEFF 
               Height          =   300
               Left            =   2880
               TabIndex        =   21
               Top             =   360
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
               Format          =   28180483
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
         End
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H000000FF&
            Caption         =   "Supprimer la période"
            Height          =   615
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Frame fraUpdateTaux 
            Height          =   2655
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   4455
            Begin VB.TextBox txtTACMIN 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2760
               TabIndex        =   24
               Top             =   2280
               Width           =   1095
            End
            Begin VB.TextBox txtTAMETH 
               Height          =   285
               Left            =   2760
               TabIndex        =   22
               Top             =   1800
               Width           =   495
            End
            Begin VB.OptionButton optTAFRQ_Q 
               Caption         =   "Taux trimestriel"
               Height          =   255
               Left            =   240
               TabIndex        =   18
               Top             =   600
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton optTAFRQ_M 
               Caption         =   "Taux mensuel"
               Height          =   255
               Left            =   240
               TabIndex        =   17
               Top             =   960
               Width           =   1575
            End
            Begin VB.OptionButton optTAFRQ_D 
               Caption         =   "Taux quotidien"
               Height          =   255
               Left            =   240
               TabIndex        =   16
               Top             =   1440
               Width           =   1815
            End
            Begin VB.TextBox txtTATAUX 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2760
               TabIndex        =   15
               Top             =   720
               Width           =   975
            End
            Begin VB.Label libTADEFF 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               Height          =   255
               Left            =   0
               TabIndex        =   26
               Top             =   120
               Width           =   4455
            End
            Begin VB.Label lblTACMIN 
               Caption         =   "montant minimun :"
               Height          =   255
               Left            =   240
               TabIndex        =   25
               Top             =   2280
               Width           =   1575
            End
            Begin VB.Label lblTAMETH 
               Caption         =   "méthode(01 02 03 04 05)"
               Height          =   255
               Left            =   240
               TabIndex        =   23
               Top             =   1920
               Width           =   2055
            End
         End
         Begin VB.CommandButton cmdUpdateNOK 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Abandonner"
            Height          =   615
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CommandButton cmdUpdateOk 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Valider les modifications"
            Height          =   615
            Left            =   7080
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   2400
            Width           =   1815
         End
      End
      Begin VB.Frame fraSelect 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   -74760
         TabIndex        =   5
         Top             =   600
         Width           =   8895
         Begin VB.Frame Frame1 
            Height          =   1815
            Left            =   4200
            TabIndex        =   35
            Top             =   240
            Width           =   2775
            Begin VB.OptionButton optSelectZZ 
               Caption         =   "ZZ"
               Height          =   255
               Left            =   1560
               TabIndex        =   41
               Top             =   1320
               Width           =   735
            End
            Begin VB.OptionButton optSelectC 
               Caption         =   "tous"
               Height          =   255
               Left            =   600
               TabIndex        =   40
               Top             =   1320
               Width           =   735
            End
            Begin VB.OptionButton optSelectCF 
               Caption         =   "CF"
               Height          =   255
               Left            =   1560
               TabIndex        =   39
               Top             =   840
               Width           =   615
            End
            Begin VB.OptionButton optSelectCH 
               Caption         =   "CH"
               Height          =   255
               Left            =   1560
               TabIndex        =   38
               Top             =   360
               Width           =   615
            End
            Begin VB.OptionButton optSelectCN 
               Caption         =   "CN"
               Height          =   255
               Left            =   600
               TabIndex        =   37
               Top             =   840
               Width           =   615
            End
            Begin VB.OptionButton optSelectCK 
               Caption         =   "CK"
               Height          =   255
               Left            =   600
               TabIndex        =   36
               Top             =   360
               Width           =   615
            End
         End
         Begin VB.OptionButton optSelectCDI 
            Caption         =   "CDI"
            Height          =   255
            Left            =   360
            TabIndex        =   10
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton optSelectCDE 
            Caption         =   "CDE"
            Height          =   255
            Left            =   360
            TabIndex        =   9
            Top             =   480
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.CommandButton cmdSelect 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Rechercher"
            Height          =   1455
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   3240
            Width           =   2415
         End
         Begin VB.TextBox txtSelect 
            Height          =   285
            Left            =   1920
            TabIndex        =   6
            Top             =   600
            Width           =   1215
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSelect 
         Height          =   2850
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   5027
         _Version        =   393216
         Rows            =   1
         Cols            =   14
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   14737632
         ForeColor       =   12582912
         ForeColorFixed  =   -2147483641
         BackColorSel    =   14737632
         BackColorBkg    =   14737632
         AllowBigSelection=   0   'False
         TextStyleFixed  =   4
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"CDTauPf.frx":047A
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8880
      Picture         =   "CDTauPf.frx":0540
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
   Begin VB.Label libRéférenceInterne 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextOption 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnufgSelect 
      Caption         =   "Opération"
      Visible         =   0   'False
      Begin VB.Menu mnufgSelect_Update 
         Caption         =   "Modifier cette ligne"
      End
      Begin VB.Menu mnufgSelect_AddNew 
         Caption         =   "Ajouter une ligne"
      End
      Begin VB.Menu mnufgSelect_Display 
         Caption         =   "Afficher cette ligne"
      End
   End
   Begin VB.Menu mnucmdPrint 
      Caption         =   "Print"
      Visible         =   0   'False
      Begin VB.Menu mnucmdPrint_fgSelect 
         Caption         =   "Imprimer la liste"
      End
   End
End
Attribute VB_Name = "frmCDTauPf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim CDTAUPFAut As typeAuthorization

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim recCDTauPf As typeCDTauPf, xCDTAUPF As typeCDTauPf, mCDTAUPF As typeCDTauPf

Dim meCDTAUPF() As typeCDTauPf
Dim meCDTAUPF_Nb As Integer, meCDTAUPF_Index As Integer, meCDTAUPF_NbMax As Integer

Dim blncmdUpdateOK_Visible As Boolean, blnErr As Boolean, blnUpdatePériode_Enabled As Boolean
Dim blnfgSelect_DisplayLine As Boolean

Dim blnSetfocus As Boolean

Private Sub cmdDelete_Click()

X = MsgBox("Voulez-vous réellement supprimer cette ligne?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption)
If X = vbNo Then Exit Sub

Me.Enabled = False

lstErr.Clear
xCDTAUPF = mCDTAUPF

currentAction = constDelete

cmdUpdate_Db

Exit_Sub:

currentAction = ""
Me.Enabled = True
AppActivate Me.Caption

End Sub

Private Sub cmdUpdateNOK_Click()

currentAction = ""
txtTATAUX = ""
cmdUpdate_Db

End Sub

Private Sub cmdUpdateOk_Click()

Me.Enabled = False

lstErr.Clear
xCDTAUPF = mCDTAUPF: cmdUpdate_Control

If lstErr.ListCount <> 0 Then GoTo Exit_Sub

cmdUpdate_Db

Exit_Sub:

currentAction = ""
Me.Enabled = True
AppActivate Me.Caption



End Sub

Private Sub cmdUpdateQuit()
fgSelect.Row = 0
Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
txtTATAUX = ""
fraUpdate.Enabled = False
libRéférenceInterne = ""
End Sub

Private Sub fraUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraUpdateTaux_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub mnufgSelect_AddNew_Click()
meCDTAUPF_Index = meCDTAUPF_Nb
mCDTAUPF = meCDTAUPF(meCDTAUPF_Nb)

mCDTAUPF.TADEFF = dateElp("Jour", 1, mCDTAUPF.TAFEFF)
mCDTAUPF.TAFEFF = mCDTAUPF.TADEFF
mCDTAUPF.TADCRT = DSys
xCDTAUPF = mCDTAUPF

fraOpération_Display
currentAction = constAddNew
fraUpdate_Init


End Sub

Private Sub mnufgSelect_Display_Click()
srvCDTauPf_ElpDisplay mCDTAUPF

End Sub

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
blnControl = False
lstErr.Clear
If fraUpdate.Enabled Then
    cmdUpdateQuit
Else
    If sstab1.Tab <> 0 Then
        sstab1.Tab = 0
    Else
        If currentAction = "" Then
            If blnMsgBox_Quit Then
                X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
            Else
               X = vbYes
            End If
            If X = vbYes Then Unload Me
        Else
            cmdReset
        End If
    End If
End If


End Sub
Public Sub cmdControl()

If Not Me.Enabled Then Exit Sub
Me.Enabled = False

'cmdOk.Visible = False
'cmdSave.Visible = False
blnControl = False
'blnSetfocus = False

lstErr.Clear
lstErr.Height = 200

cmdControl_X

If xCDTAUPF.TADNUM = 0 Then Call lstErr_AddItem(lstErr, txtSelect, "? préciser le dossier")

If lstErr.ListCount > 0 Then
    lstErr.Visible = True
End If

ExitSub:

Me.Enabled = True
    
blnControl = True

End Sub


Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

Private Sub cmdSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdSelect

End Sub

Private Sub fraOption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub mnufgSelect_Update_Click()

'$$$$$ NE PAS EFFACER LE PREMIER ENREGISTREMENT
'$$$$$ CONTINUITE DES PERIODES : ON NE PEUT MODIFIER QUE LA DERNIERE DATE

currentAction = constUpdate
fraUpdate_Init

End Sub

Private Sub optTACODC_CF_Click()
optUpdatePériode_Color
End Sub

Private Sub optTACODC_CH_Click()
optUpdatePériode_Color
End Sub


Private Sub optTACODC_CK_Click()
optUpdatePériode_Color
End Sub


Private Sub optTACODC_CN_Click()
optUpdatePériode_Color
End Sub


Private Sub optTAFRQ_D_Click()
optUpdate_Color
End Sub

Private Sub optTAFRQ_D_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optTAFRQ_D

End Sub

Private Sub optTAFRQ_M_Click()
optUpdate_Color
End Sub

Private Sub optTAFRQ_M_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optTAFRQ_M

End Sub

Private Sub optTAFRQ_Q_Click()
optUpdate_Color
End Sub

Private Sub optTAFRQ_Q_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optTAFRQ_Q
End Sub


Private Sub optSelectCDI_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSelectCDI
End Sub


Private Sub optSelectCDE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSelectCDE
End Sub


Private Sub txtTATAUX_GotFocus()

txt_GotFocus txtTATAUX

End Sub


Private Sub txtTATAUX_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtTATAUX)
End Sub


Private Sub txtTATAUX_LostFocus()
txt_LostFocus txtTATAUX
'txtTATAUX_Control
End Sub

Private Sub txtSelect_GotFocus()
txt_GotFocus txtSelect
End Sub

Private Sub txtSelect_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub

Private Sub txtSelect_LostFocus()
txt_LostFocus txtSelect
End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
currentAction = ""
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
fgSelect_Reset
blncmdUpdateOK_Visible = False: blnUpdatePériode_Enabled = False

optSelectCDE.Value = "1"
optSelectC.Value = "1"
recCDTauPf_Init xCDTAUPF

cmdUpdateOk.Visible = False
cmdDelete.Visible = False

cmdUpdateNOK.Visible = False
blnControl = True
txtSelect.SetFocus
End Sub


Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSelect.Row

If lRow > 0 Then
    fgSelect.Row = lRow
    For I = 0 To fgSelect_arrIndex
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = 0 To fgSelect_arrIndex
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.Col = 0
    End If
End If

End Sub

Private Sub fgSelect_Display()
Dim K2 As Integer, I As Integer
Dim curDB As Currency, curCR As Currency, curX As Currency

cmdUpdateQuit
sstab1.Tab = 1

fgSelect.Visible = True
fgSelect.Clear: fgSelect.Rows = 1: fgSelect_RowDisplay = 0

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Enabled = True
For meCDTAUPF_Index = 1 To meCDTAUPF_Nb
    If xCDTAUPF.Method <> constIgnore And xCDTAUPF.Method <> constDelete Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine
    End If
Next meCDTAUPF_Index

fgSelect_SortAD = 5
fgSelect_Sort1_Old = 1: fgSelect_Sort1 = 1
If fgSelect.Rows > 1 Then fgSelect_SortX 1

End Sub
Public Sub fgSelect_DisplayLine()

xCDTAUPF = meCDTAUPF(meCDTAUPF_Index)
fgSelect.Col = 0:
fgSelect.Text = xCDTAUPF.TADPFX & " " & xCDTAUPF.TADNUM
fgSelect.Col = 1: fgSelect.Text = xCDTAUPF.TACODC

fgSelect.Col = 2: fgSelect.Text = dateImp(xCDTAUPF.TADEFF)
fgSelect.Col = 3: fgSelect.Text = dateImp(xCDTAUPF.TAFEFF)
fgSelect.Col = 4:
If xCDTAUPF.TATAUX = 0 Then
    fgSelect.Text = ""
Else
    fgSelect.Text = Format$(xCDTAUPF.TATAUX, "### ##0.00#####")
End If
fgSelect.Col = 5:
Select Case xCDTAUPF.TAFRQ
    Case "Q":   fgSelect.Text = "trimestriel"
    Case "M":   fgSelect.Text = "mensuel"
    Case "D":   fgSelect.Text = "quotidien"
    Case Else: fgSelect.Text = xCDTAUPF.TAFRQ
End Select

fgSelect.Col = 6: fgSelect.Text = xCDTAUPF.TAMETH
If xCDTAUPF.TACMIN = 0 Then
    fgSelect.Text = ""
Else
    fgSelect.Col = 7: fgSelect.Text = Format$(xCDTAUPF.TACMIN, "### ### ##0.00") & "  " & xCDTAUPF.TACCCY
End If
fgSelect.Col = 8: fgSelect.Text = dateImp(xCDTAUPF.TADCRT)
fgSelect.Col = 9: fgSelect.Text = dateImp(xCDTAUPF.TADLUP)
fgSelect.Col = 10: fgSelect.Text = Trim(xCDTAUPF.TAUSER)

fgSelect.Col = fgSelect_arrIndex - 1: fgSelect.Text = ""
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = meCDTAUPF_Index

End Sub
Public Sub fgSelect_Load()
Dim X As String, mMethod As String


xCDTAUPF.Method = "SnapPF"
xCDTAUPF.TADEFF = "00000000"

meCDTAUPF(0) = xCDTAUPF
If xCDTAUPF.TACODC = "  " Then meCDTAUPF(0).TACODC = "99"
meCDTAUPF(0).TADEFF = "99999999"

Call srvCDTauPf_Load(xCDTAUPF, meCDTAUPF(0))

meCDTAUPF_Nb = srvCDTauPf.arrCDTauPf_Nb
meCDTAUPF_NbMax = meCDTAUPF_Nb + 1: ReDim meCDTAUPF(meCDTAUPF_NbMax)
For I = 1 To meCDTAUPF_Nb
    meCDTAUPF(I) = srvCDTauPf.arrCDTauPf(I)
Next I

If meCDTAUPF_Nb = 0 Then
    Call lstErr_Clear(lstErr, txtSelect, "? pas de dossier")
Else
    fgSelect_Display
End If
End Sub
Public Sub fgSelect_Sort()
If fgSelect.Rows > 1 Then
    fgSelect.Row = 1
    fgSelect.RowSel = fgSelect.Rows - 1
    
    If fgSelect_Sort1_Old = fgSelect_Sort1 Then
        If fgSelect_SortAD = 5 Then
            fgSelect_SortAD = 6
        Else
            fgSelect_SortAD = 5
        End If
    Else
        fgSelect_SortAD = 5
    End If
    fgSelect_Sort1_Old = fgSelect_Sort1
    
    fgSelect.Col = fgSelect_Sort1
    fgSelect.ColSel = fgSelect_Sort2
    fgSelect.Sort = fgSelect_SortAD
End If

End Sub
Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    meCDTAUPF_Index = Val(fgSelect.Text)
    fgSelect.Col = fgSelect_arrIndex - 1
    Select Case lK
 '       Case 1: fgSelect.Text = Format$(xCDTAUPF.EARIdRef, "00000000")
 '       Case 2: fgSelect.Text = Format$(xCDTAUPF.MONDEV, "000000000000000.00")
        Case fgSelect_arrIndex: fgSelect.Text = Format$(meCDTAUPF_Index, "0000000000")
    End Select
Next I

fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub


Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

sstab1.Tab = 0
ReDim meCDTAUPF(10)

blnControl = False
fgSelect_FormatString = fgSelect.FormatString

cmdReset

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


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Me.PopupMenu mnucmdPrint, vbPopupMenuLeftButton


End Sub

Private Sub cmdSelect_Click()
cmdControl
If lstErr.ListCount = 0 Then fgSelect_Load
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
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False
fgSelect.Clear: fgSelect.Row = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xStatut As String

If Y <= fgSelect.RowHeightMin Then
'    Select Case fgSelect.Col
'        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
'        Case 1: fgSelect_SortX 1
'        Case 2:  fgSelect_SortX 2
'        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
'        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
'    End Select
Else
    If fgSelect.Rows > 1 Then
        fgSelect.Col = fgSelect_arrIndex
        meCDTAUPF_Index = Val(fgSelect.Text)
        mCDTAUPF = meCDTAUPF(meCDTAUPF_Index)
        xCDTAUPF = mCDTAUPF
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        
        fraOpération_Display
   
       If Button = vbRightButton Then
            mnufgSelect_Display = CDTAUPFAut.Consulter
            Me.PopupMenu mnufgSelect, vbPopupMenuLeftButton
       Else
            If CDTAUPFAut.Consulter Then mnufgSelect_Update_Click
       
       End If
    End If
End If

End Sub
Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub

Private Sub mnuContextQuitter_Click()
Unload Me
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), CDTAUPFAut)    ' "EAR"

blnSetfocus = True
Form_Init

End Sub


Public Sub cmdContext_Return()
If currentAction = "" Then
    cmdSelect_Click
Else
    SendKeys "{TAB}"
End If

End Sub



Public Sub fgSelect_Reset()
fgSelect_Sort1 = 1: fgSelect_Sort2 = 1
fgSelect_Sort1_Old = 0
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 13
blnfgSelect_DisplayLine = False

End Sub

Public Sub cmdUpdate_Control()
Dim V, X As String

lstErr.Clear
lstErr.Height = 200
If optTAFRQ_M Then
    xCDTAUPF.TAFRQ = "M"
Else
    If optTAFRQ_D Then
        xCDTAUPF.TAFRQ = "D"
    Else
        xCDTAUPF.TAFRQ = "Q"
    End If
End If
X = num_Control(txtTATAUX, V, 2, 7)
xCDTAUPF.TATAUX = CDbl(V)

xCDTAUPF.TAMETH = Trim(txtTAMETH)
Select Case xCDTAUPF.TAMETH
    Case "01", "02", "03", "04", "05", "99"
    Case Else: Call lstErr_AddItem(lstErr, cmdContext, "? méthode : 01,02,03,04,99"): txtTAMETH.SetFocus
End Select

X = num_Control(txtTACMIN, V, 11, 2)
xCDTAUPF.TACMIN = CCur(V)

If blnUpdatePériode_Enabled Then cmdUpdate_Control_Période

End Sub

Public Sub cmdUpdate_Control_Période()
Dim I As Integer, blnTest As Boolean

If optTACODC_CK Then
    xCDTAUPF.TACODC = "CK"
Else
    If optTACODC_CN Then
        xCDTAUPF.TACODC = "CN"
    Else
        If optTACODC_CF Then
            xCDTAUPF.TACODC = "CF"
        Else
            If optTACODC_CH Then
                xCDTAUPF.TACODC = "CH"
            Else
                xCDTAUPF.TACODC = "ZZ"
            End If
        End If
    End If
End If

Call DTPicker_Control(txtTADEFF, xCDTAUPF.TADEFF)
Call DTPicker_Control(txtTAFEFF, xCDTAUPF.TAFEFF)
If xCDTAUPF.TAFEFF < xCDTAUPF.TADEFF Then Call lstErr_AddItem(lstErr, cmdContext, "? date fin < date début"): txtTAFEFF.SetFocus
    
For I = 1 To meCDTAUPF_Nb
    blnTest = True
    If I = meCDTAUPF_Index And currentAction = constUpdate Then blnTest = False
    If xCDTAUPF.TACODC <> meCDTAUPF(I).TACODC Then blnTest = False
    
    If blnTest Then
        If xCDTAUPF.TADEFF >= meCDTAUPF(I).TADEFF And xCDTAUPF.TADEFF <= meCDTAUPF(I).TAFEFF Then Call lstErr_AddItem(lstErr, cmdContext, "? date début : chevauchement") ': txtTADEFF.SetFocus
        If xCDTAUPF.TADEFF >= meCDTAUPF(I).TAFEFF And xCDTAUPF.TAFEFF <= meCDTAUPF(I).TAFEFF Then Call lstErr_AddItem(lstErr, cmdContext, "? date fin : chevauchement") ': txtTAFEFF.SetFocus
    End If
Next I
End Sub



Public Sub optUpdate_Color()
optTAFRQ_Q.ForeColor = fraUpdate.ForeColor
optTAFRQ_D.ForeColor = fraUpdate.ForeColor
optTAFRQ_M.ForeColor = fraUpdate.ForeColor

If optTAFRQ_M Then
    optTAFRQ_M.ForeColor = warnUsrColor
Else
    If optTAFRQ_D Then
        optTAFRQ_D.ForeColor = warnUsrColor
    Else
        optTAFRQ_Q.ForeColor = warnUsrColor
    End If
End If

End Sub
Public Sub optUpdatePériode_Color()
optTACODC_CK.ForeColor = fraUpdate.ForeColor
optTACODC_CN.ForeColor = fraUpdate.ForeColor
optTACODC_CF.ForeColor = fraUpdate.ForeColor
optTACODC_CH.ForeColor = fraUpdate.ForeColor
optTACODC_ZZ.ForeColor = fraUpdate.ForeColor

If optTACODC_CK Then
    optTACODC_CK.ForeColor = warnUsrColor
Else
    If optTACODC_CN Then
        optTACODC_CN.ForeColor = warnUsrColor
    Else
        If optTACODC_CF Then
            optTACODC_CF.ForeColor = warnUsrColor
        Else
            If optTACODC_CH Then
                optTACODC_CH.ForeColor = warnUsrColor
            Else
                optTACODC_ZZ.ForeColor = warnUsrColor
            End If
        End If
    End If
End If

End Sub

Public Sub cmdUpdate_Db()
Dim blnfgSelect_Load As Boolean

fraUpdate.Enabled = False
cmdUpdateOk.Visible = False
cmdUpdateNOK.Visible = False
cmdDelete.Visible = False

Select Case currentAction
    Case constUpdate: xCDTAUPF.Method = constUpdate: blnfgSelect_DisplayLine = True
    Case constDelete: xCDTAUPF.Method = constDelete: blnfgSelect_DisplayLine = False
    Case constAddNew: xCDTAUPF.Method = constAddNew: blnfgSelect_DisplayLine = False
    Case Else
        Call lstErr_AddItem(lstErr, cmdContext, "Abandon action en cours : " & currentAction)
End Select

If lstErr.ListCount = 0 Then
    xCDTAUPF.TADLUP = DSys
    xCDTAUPF.TAUSER = usrId
    V = srvCDTauPf_Update(xCDTAUPF)
    If IsNull(V) Then
        If blnfgSelect_DisplayLine Then
            meCDTAUPF(meCDTAUPF_Index) = xCDTAUPF
            mCDTAUPF = xCDTAUPF
            fgSelect_DisplayLine
        Else
        cmdControl_X
          fgSelect_Load
        End If
    End If
End If

End Sub


Public Sub fraOpération_Display()
currentAction = ""
fraUpdate.Enabled = False
cmdUpdateOk.Visible = False
cmdUpdateNOK.Visible = False
cmdDelete.Visible = False

libTADEFF = xCDTAUPF.TACODC & " : du   " & dateImp(xCDTAUPF.TADEFF) & "  au   " & dateImp(xCDTAUPF.TAFEFF)
Call DTPicker_Set(txtTADEFF, xCDTAUPF.TADEFF)
Call DTPicker_Set(txtTAFEFF, xCDTAUPF.TAFEFF)
txtTATAUX = Format$(xCDTAUPF.TATAUX, "### ##0.00#####")
txtTAMETH = xCDTAUPF.TAMETH
txtTACMIN = Format$(xCDTAUPF.TACMIN, "### ### ##0.00")

Select Case xCDTAUPF.TACODC
    Case "CK":   optTACODC_CK = True
    Case "CN":   optTACODC_CN = True
    Case "CF":   optTACODC_CF = True
    Case "CH":   optTACODC_CH = True
    Case "ZZ":   optTACODC_ZZ = True
   Case Else:   optTACODC_CK = True
End Select
optUpdatePériode_Color

Select Case xCDTAUPF.TAFRQ
    Case "Q":   optTAFRQ_Q = True
    Case "M":   optTAFRQ_M = True
    Case Else:   optTAFRQ_D = True
End Select

optUpdate_Color
End Sub

Public Sub fraUpdate_Init()
fraUpdate.Enabled = True
fraUpdateTaux.Enabled = True
cmdUpdateOk.Visible = True: cmdUpdateOk.Caption = currentAction
cmdUpdateNOK.Visible = True
cmdDelete.Visible = False
fraUpdatePériode.Enabled = False: blnUpdatePériode_Enabled = False: fraUpdateK.Enabled = False

If currentAction = constAddNew Then
    fraUpdatePériode.Enabled = True: blnUpdatePériode_Enabled = True: fraUpdateK.Enabled = True
Else
    If meCDTAUPF_Index = meCDTAUPF_Nb And CDTAUPFAut.Valider Then
        fraUpdatePériode.Enabled = True: blnUpdatePériode_Enabled = True
        'If meCDTAUPF_Index > 1 Then
        cmdDelete.Visible = True
    End If
End If
End Sub

Public Sub cmdControl_X()
If optSelectCDE Then
    xCDTAUPF.TADPFX = "CDE"
Else
    xCDTAUPF.TADPFX = "CDI"
End If
xCDTAUPF.TADNUM = Val(txtSelect)
If optSelectC Then
    xCDTAUPF.TACODC = "  "
Else

    If optSelectCK Then
        xCDTAUPF.TACODC = "CK"
    Else
        If optSelectCN Then
            xCDTAUPF.TACODC = "CN"
        Else
            If optSelectCF Then
                xCDTAUPF.TACODC = "CF"
            Else
                If optSelectCH Then
                    xCDTAUPF.TACODC = "CH"
                Else
                    xCDTAUPF.TACODC = "ZZ"
                End If
            End If
        End If
    End If
End If

End Sub
