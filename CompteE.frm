VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCompteE 
   AutoRedraw      =   -1  'True
   Caption         =   "Compte"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   9420
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   5400
      TabIndex        =   10
      Top             =   0
      Width           =   3500
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11033
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
      TabCaption(0)   =   "Sélection"
      TabPicture(0)   =   "CompteE.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fgSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "picCompte"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "CompteE.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgFlux"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox picCompte 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   4800
         Left            =   5640
         ScaleHeight     =   4740
         ScaleWidth      =   3585
         TabIndex        =   20
         Top             =   1250
         Width           =   3645
      End
      Begin VB.Frame Frame2 
         ForeColor       =   &H8000000F&
         Height          =   900
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   9375
         Begin VB.TextBox txtAlpha 
            Height          =   300
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   3
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtDeviseX 
            Height          =   300
            Left            =   360
            MaxLength       =   3
            TabIndex        =   5
            Top             =   480
            Width           =   405
         End
         Begin VB.TextBox txtCompte 
            Height          =   300
            Left            =   1440
            MaxLength       =   11
            TabIndex        =   0
            Top             =   480
            Width           =   1140
         End
         Begin VB.TextBox txtBiaTyp 
            Height          =   300
            Left            =   3400
            MaxLength       =   3
            TabIndex        =   1
            Top             =   480
            Width           =   400
         End
         Begin VB.TextBox txtBiaNum 
            Height          =   300
            Left            =   4515
            MaxLength       =   2
            TabIndex        =   2
            Top             =   480
            Width           =   405
         End
         Begin VB.TextBox txtNuméroAncien 
            Height          =   300
            Left            =   8200
            MaxLength       =   6
            TabIndex        =   4
            Top             =   480
            Width           =   850
         End
         Begin VB.Label lblDeviseX 
            Caption         =   "  Devise"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   150
            Width           =   780
         End
         Begin VB.Label lblNuméro 
            Caption         =   " Racine/Compte"
            Height          =   255
            Left            =   1320
            TabIndex        =   16
            Top             =   150
            Width           =   1395
         End
         Begin VB.Label lblNuméroAncien 
            Caption         =   "     Ancien"
            Height          =   250
            Left            =   8000
            TabIndex        =   15
            Top             =   150
            Width           =   1300
         End
         Begin VB.Label lblIntitulé 
            Caption         =   "                 Alpha"
            Height          =   255
            Left            =   5520
            TabIndex        =   14
            Top             =   120
            Width           =   2655
         End
         Begin VB.Label lblBiaTyp 
            Caption         =   "       Type"
            Height          =   250
            Left            =   2950
            TabIndex        =   13
            Top             =   150
            Width           =   1100
         End
         Begin VB.Label lblBiaNum 
            Caption         =   "   N° d'ordre"
            Height          =   255
            Left            =   4080
            TabIndex        =   12
            Top             =   120
            Width           =   1305
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSelect 
         Height          =   4890
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   8625
         _Version        =   393216
         Rows            =   1
         Cols            =   9
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   14737632
         ForeColor       =   12582912
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12648384
         BackColorBkg    =   14737632
         AllowBigSelection=   0   'False
         TextStyle       =   4
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   2
         AllowUserResizing=   3
         BorderStyle     =   0
         FormatString    =   "<Statut|< Compte              |>Solde                             |<Devise |<Info|<Mvt|> Séq||"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid fgFlux 
         Height          =   4770
         Left            =   -75000
         TabIndex        =   19
         Top             =   1320
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   8414
         _Version        =   393216
         Rows            =   1
         Cols            =   5
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
         FormatString    =   "<Type de compte                             |>Solde EUR                             |>Autorisation             |<Date      |"
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
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
      TabIndex        =   8
      Top             =   0
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuOpérationSaisir 
         Caption         =   "Saisir "
      End
      Begin VB.Menu mnuX1 
         Caption         =   "-"
      End
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
   Begin VB.Menu mnuCompte 
      Caption         =   "mnuCompte"
      Visible         =   0   'False
      Begin VB.Menu mnuEuroSelectOn 
         Caption         =   "Euro : sélectionner"
      End
      Begin VB.Menu mnuEuroSelectOff 
         Caption         =   "Euro : ignorer"
      End
      Begin VB.Menu mnuEuroUpdate 
         Caption         =   "Euro : PASSAGE Définitif du compte sélectionné"
      End
      Begin VB.Menu mnuEuroUpdateAll 
         Caption         =   "Euro : PASSAGE Définitif de tous les comptes sélectionnés"
      End
      Begin VB.Menu mnuEuroAddNew 
         Caption         =   "Euro : création d'un compte EURO"
      End
      Begin VB.Menu mnuEuroUpdateF 
         Caption         =   "Euro : rendre le  compte EURO inactif (==> FRF)"
      End
   End
End
Attribute VB_Name = "frmCompteE"
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
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean, blnSetfocus As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim CompteGAut As typeAuthorization

Dim recTable As typeElpTable
Dim wAmjEngagement As String, wAmjEchéance As String, blnAmjEchéance As Boolean
Dim wAmjDébut  As String, wAmjFin As String

Dim fgFlux_FormatString As String, fgFlux_K As Integer
Dim fgFlux_RowDisplay As Integer, fgFlux_RowClick As Integer
Dim fgFlux_ColorClick As Long, fgFlux_ColorDisplay As Long
Dim fgFlux_Sort1 As Integer, fgFlux_Sort2 As Integer
Dim fgFlux_SortAD As Integer, fgFlux_Sort1_Old As Integer

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer


Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnfgSelect_DisplayLine As Boolean, blnfgEchéance_DisplayLine As Boolean

Dim meCompteMin As typeCompte, meCompteMax As typeCompte, zCompteMin As typeCompte
Dim meRacine As typeRacine
Dim meCompte() As typeCompte, mCompte As typeCompte, mCptInfo As typeCptInfo
Dim meCompte_Nb As Integer, meCompte_Index As Integer, meCompte_NbMax As Integer
Dim meCV1 As typeCV, meCV2 As typeCV, meCV3 As typeCV
'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
blnControl = False

lstErr.Clear
If currentAction = "" Then
    If blnMsgBox_Quit Then
        X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
    Else
           X = vbYes
    End If
    If X = vbYes Then Unload Me
Else
    currentAction = ""
    cmdContext.Caption = constcmdRechercher
    fgSelect.Enabled = True
    fgFlux.Enabled = True
    If fgSelect.Rows > 1 Then
        SSTab1.Tab = 1
    Else
        cmdReset
    End If
End If

End Sub
Public Sub cmdControl()
Dim x5 As String
If Not Me.Enabled Then Exit Sub
Me.Enabled = False

blnControl = False
blnSetfocus = False

lstErr.Clear
lstErr.Height = 200
lastActiveControl_Name = currentActiveControl_Name
x5 = Format$(mId$(Trim(txtCompte), 1, 5), "00000")

meCompteMin = zCompteMin
meCompteMin.Method = "SnapL5"
meCompteMin.Numéro = x5 & "000000"
meCompteMin.Devise = "000"
meCompteMin.MvtceJour = " "
meCompteMin.chkAnnul = "0"

meCompteMax = meCompteMin
meCompteMax.Devise = "999"
meCompteMax.Numéro = x5 & "999999"





ExitSub:

Me.Enabled = True
blnControl = True


End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set

cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
blncmdOk_Visible = False: blncmdSave_Visible = False
blnfgSelect_DisplayLine = False: blnfgEchéance_DisplayLine = False
    
fgFlux.Clear: fgFlux.Rows = 1: fgFlux_RowDisplay = 0
lastActiveControl_Name = "": currentActiveControl_Name = ""
txtCompte.SetFocus
blnControl = True
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

SSTab1.Tab = 0

fgSelect.Visible = True
fgSelect.Clear: fgSelect.Rows = 1: fgSelect_RowDisplay = 0: fgSelect_RowClick = 0

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Enabled = True
For meCompte_Index = 1 To meCompte_Nb
    If meCompte(meCompte_Index).Method <> constIgnore And meCompte(meCompte_Index).Method <> constDelete Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine
    End If
Next meCompte_Index

fgSelect_SortAD = 5
If fgSelect.Rows > 1 Then
    fgSelect.Row = 1
    meCompte_Index = 1
    mCompte = meCompte(meCompte_Index)
    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    CptInfo_Load
        
    picCompte_Display
End If
End Sub
Public Sub fgSelect_DisplayLine()
Dim X As String
X = Compte_Imp(meCompte(meCompte_Index).Numéro)

fgSelect.Col = 0: fgSelect.Text = meCompte(meCompte_Index).Situation
If meCompte(meCompte_Index).TypeGA = "A" Then
    fgSelect.Col = 1: fgSelect.Text = X
    fgSelect.Col = 6: fgSelect.Text = mId$(X, 11, 2)
End If
fgSelect.Col = 2: fgSelect.Text = Format(meCompte(meCompte_Index).SoldeInstantané, "#### ### ###.00 ")
Call CV_AttributS(meCompte(meCompte_Index).Devise, meCV1)

fgSelect.Col = 3: fgSelect.Text = meCV1.DeviseIso
fgSelect.Col = fgSelect_arrIndex - 1: fgSelect.Text = ""
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = meCompte_Index
If meCompte(meCompte_Index).Situation <> " " Then
    For I = 0 To fgSelect_arrIndex
      fgSelect.Col = I: fgSelect.CellForeColor = vbRed
    Next I
Else
    If meCV1.EuroIn Then
        fgSelect.Col = 0
       ' If mId$(meCompte(meCompte_Index).Numéro, 6, 3) <> "001" _
       ' And mId$(meCompte(meCompte_Index).Numéro, 6, 3) <> "004" _
       ' And mId$(meCompte(meCompte_Index).Numéro, 6, 3) <> "050" _
       ' And mId$(meCompte(meCompte_Index).Numéro, 6, 3) <> "080" _
       ' And mId$(meCompte(meCompte_Index).Numéro, 6, 3) <> "081" _
       ' And mId$(meCompte(meCompte_Index).Numéro, 6, 3) <> "807" Then
       '     fgSelect.Text = "?euro?"
       '     fgSelect.CellBackColor = RGB(200, 255, 200)
       ' Else
            fgSelect.Text = "=euro"
            For I = 0 To fgSelect_arrIndex
              fgSelect.Col = I: fgSelect.CellBackColor = RGB(200, 255, 200) 'greenColor.BackColor
            Next I
        'End If
    End If
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
    meCompte_Index = Val(fgSelect.Text)
    Select Case lK
        Case 2: fgSelect.Col = 3: X = Format$(meCompte(meCompte_Index).SoldeXXX, "000000000000000.00") & fgSelect.Text
        Case 3: fgSelect.Col = 3: X = fgSelect.Text & Format$(meCompte(meCompte_Index).SoldeXXX, "000000000000000.00")
        Case 6: fgSelect.Col = 6: X = fgSelect.Text: fgSelect.Col = 1: X = X & fgSelect.Text
       Case fgSelect_arrIndex: X = Format$(meCompte_Index, "0000000000")
    End Select
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I

fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub


Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents
recCompteInit zCompteMin
zCompteMin.Société = SocId$
zCompteMin.Agence = SocAgence$

recCptInfoInit mCptInfo
mCptInfo.Société = SocId$
mCptInfo.Agence = SocAgence$


meCV1 = CV_Euro
meCV1.OpéAmj = DSys
meCV1.Normal = "P"
meCV1.AchatVente = " "
meCV2 = meCV1: meCV3 = meCV1
SSTab1.Tab = 0

cmdReset

blnControl = False

fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = 0
fgSelect_FormatString = fgSelect.FormatString
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 8

fgFlux_Sort1 = 11: fgFlux_Sort2 = 11
fgFlux_Sort1_Old = 11
fgFlux_FormatString = fgFlux.FormatString
fgFlux_RowDisplay = 0: fgFlux_RowClick = 0


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
    Case Is = constcmdRechercher: Compte_Load ' Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

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
fgFlux.Clear: fgFlux.Row = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xStatut As String, blnMnu As Boolean

If Y <= fgSelect.RowHeightMin Then
    fgSelect.Row = 0
    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 1: fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 2: fgSelect_SortX 2
        Case 3: fgSelect_SortX 3
        Case 6: fgSelect_SortX 6
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        fgSelect_ColClick = fgSelect.Col
        fgSelect.Col = fgSelect_arrIndex
        meCompte_Index = Val(fgSelect.Text)
        mCompte = meCompte(meCompte_Index)
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
 
        CptInfo_Load
        
        picCompte_Display
        If fgSelect_ColClick = 0 Then 'Button = vbRightButton Then
            blnMnu = False
            mnuEuroSelectOff = False
            mnuEuroSelectOn = False
            mnuEuroUpdate = False
            mnuEuroUpdateF = False
            mnuEuroUpdateAll = False
            mnuEuroAddNew = False
         
            fgSelect.Col = 0: xStatut = fgSelect.Text
            If xStatut = "=euro" Then
                blnMnu = True
                mnuEuroSelectOff = CompteGAut.Xspécial
                mnuEuroUpdate = CompteGAut.Xspécial
                mnuEuroUpdateAll = CompteGAut.Xspécial
                mnuEuroAddNew = CompteGAut.Xspécial
            End If
            If xStatut = "?euro?" Then
                blnMnu = True
                mnuEuroSelectOn = CompteGAut.Xspécial
            End If
            If mCompte.Devise = 978 And mCompte.Situation = " " Then
                blnMnu = True
                mnuEuroUpdateF = CompteGAut.Xspécial
            End If
            If blnMnu Then Me.PopupMenu mnuCompte, vbPopupMenuLeftButton
        End If
    End If
End If
End Sub
Private Sub txtXXX_GotFocus()

'KeyAscii = convUCase(KeyAscii)

'txt_GotFocus txtXXX

'txt_LostFocus txtXXX
'If blnControl Then cmdControl

'DTPicker_GotFocus txtXXX

'DTPicker_LostFocus txtXXX
'If blnControl Then cmdControl

' Change : txtAmjfin_control

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

Private Sub mnuEuroAddNew_Click()
mCompte.obj = "SRVEURO     "
mCompte.Method = "Compte"

CptInfo_Load

If Not IsNull(srvCompte_Update(mCompte)) Then Call MsgBox("Erreur Maj", vbCritical, "Compte_Euro")
Compte_Load

End Sub

Private Sub mnuEuroSelectOff_Click()
fgSelect.Col = 0
fgSelect.Text = "?euro?"

End Sub

Private Sub mnuEuroSelectOn_Click()
fgSelect.Col = 0
fgSelect.Text = "=euro"
End Sub

Private Sub mnuEuroUpdate_Click()
mnuEuroUpdate_DB

Compte_Load
End Sub

Private Sub mnuEuroUpdateAll_Click()
Dim I As Integer

For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = 0
    If fgSelect.Text = "=euro" Then
        fgSelect.Col = fgSelect_arrIndex
        meCompte_Index = Val(fgSelect.Text)
        mCompte = meCompte(meCompte_Index)
        mnuEuroUpdate_DB
    End If
Next I

Compte_Load

End Sub


Private Sub mnuEuroUpdateF_Click()
mCompte.obj = "SRVEURO     "
mCompte.Method = "CompteF"

CptInfo_Load

If Not IsNull(srvCompte_Update(mCompte)) Then Call MsgBox("Erreur Maj", vbCritical, "Compte_Euro")
Compte_Load

End Sub

Private Sub mnuOpérationSaisir_Click()
If CompteGAut.Saisir Then
End If

End Sub


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), CompteGAut)
If blnJPL Then CompteGAut.Xspécial = True '
Form_Init

End Sub


Public Sub cmdContext_Return()
    If SSTab1.Tab = 0 Then
        Compte_Load
    Else
        SendKeys "{TAB}"
        
    End If

End Sub





Public Sub Compte_Load()
cmdControl

If blnJPL Then
    If Not IsNull(mdbCptP0_Sel(meCompteMin, meCompteMax, "Init")) Then Exit Sub
Else
    If Not IsNull(selCompte_Load(meCompteMin, meCompteMax, "Init")) Then Exit Sub
End If

meCompte_Nb = selCompte_Nb
meCompte_NbMax = selCompte_Nb + 1
ReDim meCompte(meCompte_NbMax)

For I = 1 To selCompte_Nb
    meCompte(I) = selCompte(I)
Next I

If blnJPL Then
    Call mdbCptP0_Sel(meCompteMin, meCompteMax, "End")
Else
    Call selCompte_Load(meCompteMin, meCompteMax, "End")
End If

fgSelect_Display

End Sub

Public Sub picCompte_Display()
Dim Situation As String, X As String
Dim Col4 As Integer, Col5 As Integer

meCV1.DeviseN = mCompte.Devise
meCV1.DeviseIso = ""
meCV1.Montant = mCompte.SoldeInstantané
Call CV_Transitoire(meCV1, meCV2, meCV3, X)

picCompte.Cls
Col4 = 1800: Col5 = 3000

If mCompte.Situation = " " Then
    Situation = "": picCompte.ForeColor = libUsr.ForeColor
Else
    picCompte.ForeColor = vbRed
    Select Case mCompte.Situation
        Case "A"
            Situation = "Annulé"
        Case "B"
            Situation = "Bloqué"
        Case Else
            Situation = "? " & mCompte.Situation
    End Select
End If
picCompte.FontBold = True
    
picCompte.CurrentY = 0

picCompte.CurrentX = 50: picCompte.Print Format$(mCompte.Devise, "000") & "." & Compte_Imp(mCompte.Numéro) & Situation;
If mCompte.ChéquierInterdit <> "0" Then
    picCompte.CurrentY = picCompte.CurrentY + 300
    picCompte.CurrentX = 50
    picCompte.FontBold = True
    picCompte.ForeColor = vbRed
    picCompte.Print "Interdit de chéquier";
    picCompte.FontBold = False
End If

picCompte.FontBold = False
picCompte.ForeColor = vbMagenta
If mCompte.TypeGA = "A" Then
    picCompte.CurrentY = picCompte.CurrentY + 300
    picCompte.CurrentX = 50
    picCompte.Print Trim(DicLib(13, mCompte.BiaTyp));
End If
picCompte.CurrentY = picCompte.CurrentY + 300
picCompte.CurrentX = 50: picCompte.Print meCV1.DeviseLibellé;
picCompte.CurrentY = picCompte.CurrentY + 300

picCompte.ForeColor = libUsr.ForeColor
picCompte.CurrentX = 50: picCompte.Print mCompte.Intitulé;
picCompte.CurrentY = picCompte.CurrentY + 300
picCompte.CurrentX = 50: picCompte.Print mCompte.Intitulé2;

picCompte.CurrentY = picCompte.CurrentY + 300
picCompte.ForeColor = lblUsr.ForeColor
picCompte.CurrentX = 50: picCompte.Print "Solde :";
If meCV1.Montant >= 0 Then
    picCompte.ForeColor = libUsr.ForeColor
Else
    picCompte.ForeColor = vbRed
End If
picCompte.FontBold = True
X = Format$(meCV1.Montant, "#### ### ### ### ##0.00")
picCompte.CurrentX = Col5 - picCompte.TextWidth(X)
picCompte.Print X;
picCompte.CurrentX = Col5 + 200
picCompte.FontBold = False
picCompte.Print meCV1.DeviseIso;
picCompte.CurrentY = picCompte.CurrentY + 300
picCompte.ForeColor = vbMagenta
picCompte.FontBold = False
X = Format$(meCV2.Montant, "#### ### ### ### ##0.00")
picCompte.CurrentX = Col5 - picCompte.TextWidth(X)
picCompte.Print X;
picCompte.CurrentX = Col5 + 200
picCompte.FontBold = False
picCompte.Print meCV2.DeviseIso;

If mCompte.DécouvertMontant > 0 Then
    picCompte.CurrentY = picCompte.CurrentY + 300
    picCompte.ForeColor = txtUsr.ForeColor
    picCompte.CurrentX = 50:
    If Val(mCompte.DécouvertAmj) < DSys Then
        picCompte.ForeColor = warnUsrColor
    Else
        If mCompte.SoldeInstantané + mCompte.DécouvertMontant < 0 Then
            picCompte.ForeColor = errUsr.ForeColor
        End If
    End If
    X = "Découvert autorisé : " & Trim(Format$(mCompte.DécouvertMontant, "### ### ### ###")) _
    & " juqu'au : " & dateImp(mCompte.DécouvertAmj)
    picCompte.Print X;
End If


picCompte.ForeColor = lblUsr.ForeColor
picCompte.CurrentY = picCompte.CurrentY + 300
picCompte.CurrentX = 50: picCompte.Print "Dernier mvt :  ";
picCompte.ForeColor = libUsr.ForeColor
picCompte.CurrentX = Col4: picCompte.Print dateImp(mCompte.MvtAmj);
picCompte.CurrentY = picCompte.CurrentY + 300
picCompte.ForeColor = lblUsr.ForeColor
picCompte.CurrentX = 50: picCompte.Print "Service Responsable :  ";
picCompte.ForeColor = libUsr.ForeColor
picCompte.CurrentX = Col4: picCompte.Print mCptInfo.ServiceResponsable & "_" & Trim(DicLib(4, mCptInfo.ServiceResponsable));
picCompte.CurrentY = picCompte.CurrentY + 300
picCompte.ForeColor = lblUsr.ForeColor
picCompte.CurrentX = 50: picCompte.Print "Gestionnaire :  ";
picCompte.ForeColor = libUsr.ForeColor
picCompte.CurrentX = Col4: picCompte.Print mCompte.Gestionnaire & "_" & Trim(DicLib(60, mCompte.Gestionnaire));
picCompte.CurrentY = picCompte.CurrentY + 300
picCompte.ForeColor = lblUsr.ForeColor
picCompte.CurrentX = 50: picCompte.Print "Alpha :  ";
picCompte.ForeColor = libUsr.ForeColor
picCompte.CurrentX = Col4: picCompte.Print mCompte.Alpha;
picCompte.CurrentY = picCompte.CurrentY + 300
picCompte.ForeColor = lblUsr.ForeColor
picCompte.CurrentX = 50: picCompte.Print "Echelle :  ";
picCompte.ForeColor = libUsr.ForeColor
picCompte.CurrentX = Col4: picCompte.Print Trim(DicLib(7, mCptInfo.Echelle)) & " " & dateImp(mCptInfo.EchelleAmj);
If mCptInfo.PrélèvementLibératoire = "1" Then picCompte.Print "Prél libératoire";
picCompte.CurrentY = picCompte.CurrentY + 300
picCompte.ForeColor = lblUsr.ForeColor
picCompte.CurrentX = 50: picCompte.Print "SEV :";
If mCptInfo.EchelleSolde >= 0 Then
    picCompte.ForeColor = libUsr.ForeColor
Else
    picCompte.ForeColor = vbRed
End If
picCompte.FontBold = True
X = Format$(mCptInfo.EchelleSolde, "#### ### ### ### ##0.00")
picCompte.CurrentX = Col5 - picCompte.TextWidth(X)
picCompte.Print X;



End Sub

Private Sub txtCompte_GotFocus()

txt_GotFocus txtCompte

End Sub


Private Sub txtCompte_KeyPress(KeyAscii As Integer)

KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtCompte_LostFocus()
txt_LostFocus txtCompte

End Sub



Public Sub CptInfo_Load()
'$$$$ lecture CPTINFO Veille
mCptInfo.Devise = mCompte.Devise
mCptInfo.Numéro = mCompte.Numéro
If blnJPL Then
    mdbCptInfoP0_Find mCptInfo
Else
    srvCptInfoFind mCptInfo
End If

End Sub

Public Sub mnuEuroUpdate_DB()
mCompte.obj = "SRVEURO     "
mCompte.Method = "Bascule"


CptInfo_Load

meCV1.DeviseN = mCompte.Devise
meCV1.DeviseIso = ""
meCV1.Montant = mCompte.DécouvertMontant
Call CV_Transitoire(meCV1, meCV2, meCV3, X)
mCompte.DécouvertMontant = Round(meCV3.Montant, 0)

meCV1.Montant = mCptInfo.EchelleSolde
Call CV_Transitoire(meCV1, meCV2, meCV3, X)
mCompte.SoldeVeille = meCV3.Montant

If Not IsNull(srvCompte_Update(mCompte)) Then Call MsgBox("Erreur Maj", vbCritical, "Compte_Euro")

End Sub
