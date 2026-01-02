VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEdition_Gestion 
   AutoRedraw      =   -1  'True
   Caption         =   "Edition_Gestion"
   ClientHeight    =   11700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15165
   Icon            =   "Edition_Gestion.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11700
   ScaleWidth      =   15165
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8640
      TabIndex        =   21
      Top             =   0
      Width           =   4665
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   11220
      Left            =   15
      TabIndex        =   19
      Top             =   450
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   19791
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Documents : spécifications d'impression"
      TabPicture(0)   =   "Edition_Gestion.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Utilisateurs :  gestion des impressions"
      TabPicture(1)   =   "Edition_Gestion.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraUsr"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraUsr 
         Height          =   10710
         Left            =   0
         TabIndex        =   23
         Top             =   360
         Width           =   14925
         Begin VB.Frame fraUsr_Détail 
            Height          =   9750
            Left            =   8070
            TabIndex        =   37
            Top             =   585
            Width           =   5520
            Begin VB.TextBox txtUsr_AliasWin 
               Height          =   285
               Left            =   1800
               TabIndex        =   56
               Top             =   5040
               Width           =   2055
            End
            Begin VB.OptionButton optUsr_Informatique 
               Alignment       =   1  'Right Justify
               Caption         =   "Informatique"
               Height          =   200
               Left            =   3600
               TabIndex        =   53
               Top             =   1320
               Width           =   1200
            End
            Begin VB.TextBox txtUsr_ClasseAut 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1800
               TabIndex        =   51
               Text            =   "0"
               Top             =   3840
               Width           =   375
            End
            Begin VB.OptionButton optUsr_Production 
               Caption         =   "Production"
               Height          =   200
               Left            =   360
               TabIndex        =   49
               Top             =   1320
               Width           =   1110
            End
            Begin VB.OptionButton optUsr_Test 
               Alignment       =   1  'Right Justify
               Caption         =   "Test"
               Height          =   200
               Left            =   2280
               TabIndex        =   48
               Top             =   1320
               Width           =   690
            End
            Begin VB.CommandButton cmdUsr_Esc 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Quitter"
               Height          =   705
               HelpContextID   =   16777215
               Left            =   870
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   8745
               Width           =   1440
            End
            Begin VB.CommandButton cmdUsr_Ok 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Ok"
               Height          =   735
               Left            =   3570
               Style           =   1  'Graphical
               TabIndex        =   46
               Top             =   8715
               Width           =   1305
            End
            Begin VB.CheckBox chkUsr_QSYSOPR 
               Alignment       =   1  'Right Justify
               Caption         =   "QSYSOPR"
               Height          =   270
               Left            =   240
               TabIndex        =   45
               Top             =   3360
               Width           =   1800
            End
            Begin VB.ComboBox cboUsr_Printer 
               Height          =   315
               Left            =   1800
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   4320
               Width           =   2205
            End
            Begin VB.CheckBox chkUsr_SplAut 
               Alignment       =   1  'Right Justify
               Caption         =   "Droits affichage"
               Height          =   270
               Left            =   240
               TabIndex        =   42
               Top             =   3000
               Width           =   1800
            End
            Begin VB.CheckBox chkUsr_Hold 
               Alignment       =   1  'Right Justify
               Caption         =   "Hold"
               Height          =   270
               Left            =   240
               TabIndex        =   41
               Top             =   2520
               Width           =   1800
            End
            Begin VB.ComboBox cboUsr_Unit 
               Height          =   315
               Left            =   1785
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   38
               Top             =   1920
               Width           =   3030
            End
            Begin VB.Label lblUsr_AliasWin 
               Caption         =   "Alias Windows (ACL spoules)"
               Height          =   495
               Left            =   240
               TabIndex        =   55
               Top             =   5040
               Width           =   1215
            End
            Begin VB.Label fraUsr_Opt 
               BorderStyle     =   1  'Fixed Single
               Height          =   405
               Left            =   240
               TabIndex        =   52
               Top             =   1200
               Width           =   4935
            End
            Begin VB.Label lblUsr_ClasseAut 
               Caption         =   "libre"
               Height          =   255
               Left            =   240
               TabIndex        =   50
               Top             =   3840
               Width           =   1215
            End
            Begin VB.Label lblUsr_Printer 
               Caption         =   "impr personnelle"
               Height          =   225
               Left            =   240
               TabIndex        =   44
               Top             =   4320
               Width           =   1395
            End
            Begin VB.Label libUsr_Détail 
               BorderStyle     =   1  'Fixed Single
               Height          =   405
               Left            =   105
               TabIndex        =   40
               Top             =   240
               Width           =   5310
            End
            Begin VB.Label lblUsr_Unit 
               Caption         =   "Unité Opé"
               Height          =   210
               Left            =   240
               TabIndex        =   39
               Top             =   2040
               Width           =   810
            End
         End
         Begin VB.ListBox lstUsr 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   10140
            Left            =   360
            Sorted          =   -1  'True
            TabIndex        =   36
            Top             =   225
            Width           =   7545
         End
      End
      Begin VB.Frame fraTab0 
         Height          =   10890
         Left            =   -74895
         TabIndex        =   20
         Top             =   240
         Width           =   14850
         Begin VB.CommandButton cmdPaperBin 
            BackColor       =   &H000000FF&
            Caption         =   "spécial JPL  Bac : 2 => 7"
            Height          =   972
            Left            =   7080
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   360
            Width           =   1092
         End
         Begin VB.Frame fraEdition_Form 
            Height          =   8985
            Left            =   8400
            TabIndex        =   24
            Top             =   1560
            Width           =   4515
            Begin VB.CheckBox chkEdition_NoPaper_Prod 
               Alignment       =   1  'Right Justify
               Caption         =   "NoPaper=>Prod     ne pas archiver pdf"
               Height          =   720
               Left            =   135
               TabIndex        =   59
               Top             =   6105
               Width           =   1800
            End
            Begin VB.CheckBox chkEdition_PrinterUnit 
               Alignment       =   1  'Right Justify
               Caption         =   "Toujours Printer Unit"
               Height          =   270
               Left            =   100
               TabIndex        =   35
               Top             =   5505
               Width           =   1800
            End
            Begin VB.TextBox txtEdition_FontSize 
               Height          =   285
               Left            =   1725
               TabIndex        =   34
               Top             =   1530
               Width           =   500
            End
            Begin VB.TextBox txtEdition_Copies 
               Height          =   285
               Left            =   1680
               TabIndex        =   10
               Top             =   4560
               Width           =   500
            End
            Begin VB.TextBox txtEdition_LinePerPage 
               Height          =   285
               Left            =   1710
               TabIndex        =   6
               Top             =   1155
               Width           =   500
            End
            Begin VB.ComboBox cboEdition_Unit 
               Height          =   315
               Left            =   1680
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   5040
               Width           =   1230
            End
            Begin VB.ComboBox cboEdition_PaperBin 
               Height          =   315
               Left            =   1710
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   3360
               Width           =   555
            End
            Begin VB.ComboBox cboEdition_fontName 
               Height          =   315
               Left            =   1725
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   2040
               Width           =   1725
            End
            Begin VB.CheckBox chkEdition_Duplex 
               Alignment       =   1  'Right Justify
               Caption         =   "Recto/verso"
               Height          =   270
               Left            =   100
               TabIndex        =   8
               Top             =   2400
               Width           =   1800
            End
            Begin VB.CheckBox chkEdition_Save 
               Alignment       =   1  'Right Justify
               Caption         =   "Archive"
               Height          =   270
               Left            =   100
               TabIndex        =   14
               Top             =   4200
               Width           =   1800
            End
            Begin VB.CheckBox chkEdition_Hold 
               Alignment       =   1  'Right Justify
               Caption         =   "Hold"
               Height          =   270
               Left            =   100
               TabIndex        =   13
               Top             =   3840
               Width           =   1800
            End
            Begin VB.CheckBox chkEdition_Orientation 
               Alignment       =   1  'Right Justify
               Caption         =   "Orientation Paysage"
               Height          =   270
               Left            =   100
               TabIndex        =   5
               Top             =   840
               Width           =   1800
            End
            Begin VB.CheckBox chkEdition_Courrier 
               Alignment       =   1  'Right Justify
               Caption         =   "Courrier"
               Height          =   270
               Left            =   100
               TabIndex        =   4
               Top             =   500
               Width           =   1800
            End
            Begin VB.ComboBox cboEdition_Filigrane 
               Height          =   315
               Left            =   1680
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   2880
               Width           =   1800
            End
            Begin VB.CommandButton cmdEdition_Form_Esc 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Quitter"
               Height          =   705
               HelpContextID   =   16777215
               Left            =   390
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   8055
               Width           =   1170
            End
            Begin VB.CommandButton cmdEdition_Form_Ok 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Ok"
               Height          =   705
               Left            =   3120
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   8010
               Width           =   1125
            End
            Begin VB.Label libEdition_PaperBin 
               Caption         =   "7 = auto"
               Height          =   252
               Left            =   2520
               TabIndex        =   57
               Top             =   3360
               Width           =   732
            End
            Begin VB.Label lblEdition_Unit 
               Caption         =   "Unité Opé"
               Height          =   210
               Left            =   100
               TabIndex        =   33
               Top             =   5160
               Width           =   810
            End
            Begin VB.Label lblEdition_PaperBin 
               Caption         =   "Bac imprimante"
               Height          =   315
               Left            =   120
               TabIndex        =   32
               Top             =   3360
               Width           =   1200
            End
            Begin VB.Label lblEdition_Copies 
               Caption         =   "Copies"
               Height          =   285
               Left            =   105
               TabIndex        =   31
               Top             =   4680
               Width           =   1200
            End
            Begin VB.Label lblEdition_FontSize 
               Caption         =   "Taille police"
               Height          =   270
               Left            =   100
               TabIndex        =   30
               Top             =   1545
               Width           =   1200
            End
            Begin VB.Label lblEdition_fontName 
               Caption         =   "Police"
               Height          =   165
               Left            =   120
               TabIndex        =   29
               Top             =   2040
               Width           =   1200
            End
            Begin VB.Label lblEdition_LinePerPage 
               Caption         =   "Ligne/page"
               Height          =   285
               Left            =   100
               TabIndex        =   28
               Top             =   1200
               Width           =   1200
            End
            Begin VB.Label lblEdition_Filigrane 
               Caption         =   "Filigrane"
               Height          =   225
               Left            =   105
               TabIndex        =   27
               Top             =   2880
               Width           =   1200
            End
            Begin VB.Label libEdition_form 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   105
               TabIndex        =   26
               Top             =   180
               Width           =   4230
            End
         End
         Begin VB.ListBox lstSelect 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   9930
            Left            =   195
            Sorted          =   -1  'True
            TabIndex        =   0
            Top             =   570
            Width           =   6660
         End
         Begin VB.Frame fraSelect 
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   150
            TabIndex        =   22
            Top             =   135
            Width           =   6660
            Begin VB.ComboBox cboSelect_K1 
               Height          =   315
               Left            =   30
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   60
               Width           =   1980
            End
            Begin VB.CommandButton cmdSelect 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Rechercher  = *"
               Height          =   345
               Left            =   2280
               Style           =   1  'Graphical
               TabIndex        =   2
               Top             =   0
               Width           =   1515
            End
            Begin VB.TextBox txtSelect 
               Height          =   285
               Left            =   4080
               TabIndex        =   3
               Top             =   120
               Width           =   2265
            End
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"Edition_Gestion.frx":0342
            ForeColor       =   &H00000000&
            Height          =   1215
            Left            =   8400
            TabIndex        =   54
            Top             =   240
            Width           =   4455
         End
      End
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   420
      Left            =   13320
      Picture         =   "Edition_Gestion.frx":0482
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   0
      Width           =   500
   End
   Begin VB.Label libSelect 
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
      Height          =   495
      Left            =   1185
      TabIndex        =   25
      Top             =   -15
      Width           =   3585
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
End
Attribute VB_Name = "frmEdition_Gestion"
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
Dim intReturn As Integer
Dim Edition_Gestion_Aut As typeAuthorization
Dim currentMethod As String

Dim blnError As Boolean


Dim meEdition_Form As typeEdition_Form, xEdition_Form As typeEdition_Form

Dim meElpTable As typeElpTable, xElpTable As typeElpTable

Dim meUser As typeUser
Dim blnUsr_AddNew As Boolean
'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), Edition_Gestion_Aut)

'blnSetfocus = True
Form_Init

'cmdSelect_Click

End Sub


Public Sub Form_Init()
On Error Resume Next

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True

blnControl = False

cmdEdition_Form_Ok.Visible = Edition_Gestion_Aut.Xspécial
cmdEdition_Form_Esc.Visible = Edition_Gestion_Aut.Xspécial
cmdPaperBin.Visible = False ' special JPL Edition_Gestion_Aut.Xspécial

cboSelect_K1.Clear
cboSelect_K1.AddItem "BIA"
cboSelect_K1.AddItem "SAB"

cboEdition_fontName.Clear
cboEdition_fontName.AddItem prtFontName_Arial
cboEdition_fontName.AddItem prtFontName_Comic
cboEdition_fontName.AddItem prtFontName_CenturyGothic
cboEdition_fontName.AddItem prtFontName_CourierNew
cboEdition_fontName.AddItem prtFontName_Comic
cboEdition_fontName.AddItem prtFontName_TimesNewRoman

cboEdition_Filigrane.Clear
cboEdition_Filigrane.AddItem "(non)"
cboEdition_Filigrane.AddItem "Automatique"
cboEdition_Filigrane.AddItem "BIA"


cboEdition_PaperBin.Clear
cboEdition_PaperBin.AddItem 7
cboEdition_PaperBin.AddItem 1
cboEdition_PaperBin.AddItem 2
cboEdition_PaperBin.AddItem 3

'recElpTable_Init xElpTable
'xElpTable.Id = "Unit"
Call cbo_LoadId_K2("Unit", "", cboEdition_Unit)
cboEdition_Unit.AddItem " "

xElpTable.Id = "Unit"
'Call cbo_LoadId_K2("Unit", "", cboUsr_Unit)
Call cbo_Load_Unit(cboUsr_Unit)

Call lstZMNURUT0_Load(lstUsr)

'xElpTable.Id = "Server"
'xElpTable.K1 = "Printer"
Call cbo_LoadK2("Server", "Printer", cboUsr_Printer)
cboUsr_Printer.AddItem " "

cboUsr_Printer.ListIndex = 0

cmdReset
Me.Enabled = True

End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------

blnControl = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
currentAction = ""
SSTab1.Tab = 1
libSelect = ""
lstSelect.Clear

txtSelect = ""
fraSelect.Enabled = True
'libEdition_form.ForeColor = warnUsrColor
lstSelect.Enabled = False 'Edition_Gestion_Aut.Xspécial

fraEdition_Form.Enabled = False
fraEdition_Form_Reset
fraUsr_Détail.Enabled = False
fraUsr_Détail_Reset

'fraInfo.Enabled = Edition_Gestion_Aut.Xspécial
'txtInfoYBASTAU0 = paramTemp_Folder & "FTP\YBASTAU0.txt"

blnControl = True

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

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdEdition_Form_Esc_Click()
cmdContext_Quit
End Sub

Private Sub cmdEdition_Form_Ok_Click()
Dim X As String
Dim V

Me.Enabled = False
fraEdition_Form.Enabled = False

X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & constEdition_Form & "'" _
    & " and K1 = '" & cboSelect_K1.Text & "'" _
    & " and K2 = '" & mId$(lstSelect.Text, 1, 10) & "'"
    
Set rsMDB = cnMDB.Execute(X)
If Not rsMDB.EOF Then
    Call rsElpTable_GetBuffer(rsMDB, xElpTable)

    xEdition_Form.Courrier = chkEdition_Courrier
    xEdition_Form.Orientation = chkEdition_Orientation
    xEdition_Form.Duplex = Val(chkEdition_Duplex) + 1
    xEdition_Form.Hold = chkEdition_Hold
    xEdition_Form.Save = chkEdition_Save
    xEdition_Form.LinePerPage = Val(txtEdition_LinePerPage)
    xEdition_Form.FontSize = Val(txtEdition_FontSize)
    xEdition_Form.Copies = Val(txtEdition_Copies)
    xEdition_Form.FontName = cboEdition_fontName
    X = Trim(cboEdition_Filigrane)
    Select Case X
        Case "Automatique": xEdition_Form.Filigrane = "A"
        Case "BIA": xEdition_Form.Filigrane = "B"
        Case Else: xEdition_Form.Filigrane = " "
    End Select
    
    xEdition_Form.PrinterUnit = chkEdition_PrinterUnit
    xEdition_Form.Unit = cboEdition_Unit
    xEdition_Form.PaperBin = cboEdition_PaperBin
    
    xEdition_Form.NoPaper_Prod = chkEdition_NoPaper_Prod
    
    X = Space$(200)
    rsEdition_Form_PutBuffer X, xEdition_Form
    
    xElpTable.Memo = Trim(X)
     
    V = adoElpTable_Update(rsMDB, xElpTable)

Else
    Shell_MsgBox "cmdEdition_Form_Ok_Click#  " & xElpTable.Id & " : " & xElpTable.K1 & " : " & xElpTable.K2, vbCritical, Me.Caption, False

End If

cmdContext_Quit

Me.Enabled = True
End Sub


Private Sub cmdPaperBin_Click()
Dim K As Integer
Me.Enabled = False
fraEdition_Form.Enabled = False

For K = 0 To lstSelect.ListCount - 1
    lstSelect.ListIndex = K
    X = "select * from ElpTable where SNN = 0" _
        & " and id = '" & constEdition_Form & "'" _
        & " and K1 = '" & cboSelect_K1.Text & "'" _
        & " and K2 = '" & mId$(lstSelect.Text, 1, 10) & "'"
        
    Set rsMDB = cnMDB.Execute(X)
    If Not rsMDB.EOF Then
        Call rsElpTable_GetBuffer(rsMDB, xElpTable)
    
        If mId$(xElpTable.Memo, 12, 1) = "2" Then
            Mid$(xElpTable.Memo, 12, 1) = "7"
             
            V = adoElpTable_Update(rsMDB, xElpTable)
        End If
    End If
Next K
Me.Enabled = True

End Sub

Private Sub cmdSelect_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

fraSelect.Enabled = False

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''MsgBox "Annule et remplace Edition_Form SAB "
''cmdYMNUETA0_Import
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
lstSelect_Load
lstSelect.Enabled = True
If lstSelect.ListCount = 0 Then
    cmdContext_Quit
    Call lstErr_AddItem(lstErr, cmdPrint, Time & " - aucun formulaire sélectionné"): DoEvents
End If

Me.Enabled = True: Me.MousePointer = 0



End Sub


Public Sub lstSelect_Load()
Dim xlen As Integer, X As String, wFile As String
Dim blnOk As Boolean, K As Integer
Dim blnGenX As Boolean
Dim Nb As Long
Dim xSQL As String, xSql_K2 As String

On Error Resume Next

lstSelect.Visible = False
lstSelect.Clear
'recElpTable_Init xElpTable
'xElpTable.Method = "Seek>="
'xElpTable.Id = constEdition_Form
'xElpTable.K1 = cboSelect_K1.Text
'meElpTable = xElpTable

X = Trim(txtSelect)
If mId$(X, 1, 1) <> "*" Then
    blnGenX = False
    ''xElpTable.K2 = X
    xSql_K2 = " and K2 like '" & X & "%'"
Else
    blnGenX = True
    Mid$(X, 1, 1) = " "
    X = Trim(X)
    '''xElpTable.K2 = ""
    xSql_K2 = " and K2 like '%" & X & "%'"
End If
xlen = Len(X)

xSQL = "select * from ElpTable where SNN = 0" _
    & " and id = '" & constEdition_Form & "'" _
    & " and k1 = '" & cboSelect_K1.Text & "'" _
    & xSql_K2
    
Set rsMDB = cnMDB.Execute(xSQL)
Do While Not rsMDB.EOF
    Call rsElpTable_GetBuffer(rsMDB, xElpTable)

    blnOk = True
    If xlen > 0 Then
        If blnGenX Then
            K = InStr(1, xElpTable.K2, X)
            If K = 0 Then blnOk = False
        Else
            If X <> mId$(xElpTable.K2, 1, xlen) Then blnOk = False
        End If
        
    End If
    
    If blnOk Then Nb = Nb + 1: lstSelect.AddItem xElpTable.K2 & " " & xElpTable.Name
    rsMDB.MoveNext
Loop

lstSelect.Visible = True
Call lstErr_Clear(lstErr, cmdContext, Nb & " sélectionné(s): ")
lstSelect.ListIndex = 0

End Sub
Public Sub lstEdition_Form_Select()
X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & constEdition_Form & "'" _
    & " and K1 = '" & cboSelect_K1.Text & "'" _
    & " and K2 = '" & mId$(lstSelect.Text, 1, 10) & "'"
    
Set rsMDB = cnMDB.Execute(X)
If Not rsMDB.EOF Then
    Call rsElpTable_GetBuffer(rsMDB, meElpTable)

    lstSelect.Enabled = False
    fraEdition_Form_Display
Else
    Shell_MsgBox "lstEdition_Form_Select#  " & meElpTable.Id & " : " & meElpTable.K1 & " : " & meElpTable.K2, vbCritical, Me.Caption, False

End If

End Sub

Private Sub cmdUsr_Esc_Click()
cmdContext_Quit

End Sub

Private Sub cmdUsr_Ok_Click()
Dim X As String, K As Integer

Me.Enabled = False
If cboUsr_Unit.ListIndex < 0 Then
    Call lstErr_Clear(lstErr, cmdContext, "? préciser le service")
Else

    If blnUsr_AddNew Then
        currentMethod = constAddNew
    Else
        currentMethod = constUpdate
    End If
    xElpTable.Id = "User"
    xElpTable.K1 = meUser.Id
    xElpTable.K2 = ""
    xElpTable.Name = mId$(lstUsr.Text, 14, Len(lstUsr.Text) - 13)
    xElpTable.Memo = Space$(11)
    Mid$(xElpTable.Memo, 1, 4) = Trim(cboUsr_Unit.Text)
    If optUsr_Production Then
        Mid$(xElpTable.Memo, 6, 1) = "P"
    Else
        If optUsr_Test Then
            Mid$(xElpTable.Memo, 6, 1) = "T"
        Else
            Mid$(xElpTable.Memo, 6, 1) = "I"
        End If
    End If
    
    Mid$(xElpTable.Memo, 7, 1) = chkUsr_Hold
    Mid$(xElpTable.Memo, 8, 1) = chkUsr_SplAut
    Mid$(xElpTable.Memo, 9, 1) = chkUsr_QSYSOPR
    Mid$(xElpTable.Memo, 10, 1) = mId$(Trim(txtUsr_ClasseAut), 1, 1) & " "
    
    X = Trim(cboUsr_Printer.Text)
    If X <> "" Then xElpTable.Memo = xElpTable.Memo & " PRINTER:" & X
 
    X = Trim(txtUsr_AliasWin)
    If X <> "" Then xElpTable.Memo = xElpTable.Memo & " CACLS:" & X
    
    Select Case currentMethod
        Case constAddNew: V = adoElpTable_AddNew(rsMDB, xElpTable)
        Case constDelete: V = adoElpTable_Delete(rsMDB, xElpTable)
        Case constUpdate: V = adoElpTable_Update(rsMDB, xElpTable)
    End Select
    If Not IsNull(V) Then MsgBox V, vbCritical, "BIA_SYSTEM: Edition_Gestion"
    cmdContext_Quit
End If
Me.Enabled = True

End Sub

Private Sub lstSelect_Click()
lstEdition_Form_Select
End Sub


Private Sub lstUsr_Click()
Dim K As Integer
lstUsr.Enabled = False
K = InStr(lstUsr, vbTab)
meUser.Id = mId$(lstUsr, 1, K - 1)
If Not IsNull(Table_User(meUser)) Then
    blnUsr_AddNew = True
    cmdUsr_Ok.Caption = "Créer"
    If mId$(meUser.Id, 1, 2) = "T_" Then
        meUser.Edition_Hold = "1"
        meUser.ProdTest = "T"
    Else
        meUser.Edition_Hold = "0"
        meUser.ProdTest = "P"
   End If
    
Else
    blnUsr_AddNew = False
    cmdUsr_Ok.Caption = "Modifier"
End If
fraUsr_Détail_Display

End Sub

Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub


Private Sub mnuContextQuitter_Click()
Unload Me
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

Public Sub cmdContext_Quit()
Dim blnTest As Boolean

blnTest = False
blnControl = False
lstErr.Clear: lstErr.Height = 200
If SSTab1.Tab = 0 Then
    If Not Me.Enabled Then blnTest = True
 ''   If fraEdition_Form.Enabled Then blnTest = True
    If Not lstSelect.Enabled Then blnTest = True
    If blnTest Then
        fraEdition_Form.BackColor = &H8000000F
        usrColor_Container fraEdition_Form, fraEdition_Form.BackColor
        fraEdition_Form.Enabled = False
        lstSelect.Enabled = True
        Exit Sub
    End If
    If Not fraSelect.Enabled Then lstSelect.Clear: fraSelect.Enabled = True: Exit Sub
End If
If SSTab1.Tab = 1 Then
    If Not Me.Enabled Then blnTest = True
  ''  If fraUsr_Détail.Enabled Then blnTest = True
    If Not lstUsr.Enabled Then blnTest = True
    If blnTest Then
        fraUsr_Détail.BackColor = &H8000000F
        usrColor_Container fraUsr_Détail, fraUsr_Détail.BackColor
        fraUsr_Détail.Enabled = False
        lstUsr.Enabled = True
        Exit Sub
    End If
End If
X = ""
If blnMsgBox_Quit Then
    X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
Else
    X = vbYes
End If

If X = vbYes Then Unload Me


End Sub

Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
    cmdSelect_Click
Else
    SendKeys "{TAB}"
End If
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False
End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Public Sub MouseMoveActiveControl_Set(C As Control)
If MouseMoveActiveControl_Name <> C.Name Then
    MouseMoveActiveControl_Reset
    If Not C.Enabled Then
        MouseMoveActiveControl_Name = ""
    Else
        MouseMoveActiveControl_Name = C.Name
        If TypeOf C Is CommandButton Then
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            If TypeOf C Is ListBox Then
                Elp_ResizeControl C
            Else
                MouseMoveActiveControl.ForeColor = C.ForeColor
                C.ForeColor = MouseMoveUsr.ForeColor
            End If
        End If
    End If
End If

End Sub


Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
        If TypeOf xobj Is CommandButton Then
            xobj.BackColor = MouseMoveActiveControl.BackColor
        Else
            If TypeOf xobj Is ListBox Then
                xobj.Height = 200
            Else
                xobj.ForeColor = MouseMoveActiveControl.ForeColor
            End If
        End If
        Exit For
    End If
Next xobj
End Sub


Public Sub txt_X()
'Call txt_GotFocus(txt)
'KeyAscii = convUCase(KeyAscii)
'Call txt_LostFocus(txt)

'Call txt_GotFocus(txt)
'If XopDevise(2).maxD = 0 Then
'    Call num_KeyAscii(KeyAscii)
'Else
'    Call num_KeyAsciiD(KeyAscii, txt)
'End If
'Call txt_GotFocus(txt)
'Call txt_LostFocus(txt)

End Sub



Private Sub txtEdition_Copies_GotFocus()
Call txt_GotFocus(txtEdition_Copies)

End Sub


Private Sub txtEdition_Copies_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtEdition_Copies_LostFocus()
Call txt_LostFocus(txtEdition_Copies)

End Sub


Private Sub txtEdition_FontSize_GotFocus()
Call txt_GotFocus(txtEdition_FontSize)

End Sub


Private Sub txtEdition_FontSize_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtEdition_FontSize_LostFocus()
Call txt_LostFocus(txtEdition_FontSize)

End Sub




Private Sub txtEdition_LinePerPage_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtSelect_GotFocus()
Call txt_GotFocus(txtSelect)

End Sub


Private Sub txtSelect_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtSelect_LostFocus()
Call txt_LostFocus(txtSelect)

End Sub




Public Sub fraEdition_Form_Reset()
libEdition_form = ""
cboSelect_K1.ListIndex = 1
cboEdition_PaperBin.ListIndex = 0
cboEdition_Filigrane.ListIndex = 1
cboEdition_fontName.ListIndex = 4

chkEdition_Courrier = "0"
chkEdition_Orientation = "1"
chkEdition_Duplex = "1"
chkEdition_Hold = "0"
chkEdition_Save = "0"

txtEdition_LinePerPage = 66
txtEdition_FontSize = 8
txtEdition_Copies = 1

End Sub

Public Sub fraUsr_Détail_Reset()
On Error Resume Next
libEdition_form = ""
cboUsr_Unit.ListIndex = 0
optUsr_Production = True

chkUsr_Hold = "0"
chkUsr_SplAut = "0"
chkUsr_QSYSOPR = "0"
txtUsr_ClasseAut = "0"
cboUsr_Printer.ListIndex = 0
txtUsr_AliasWin = ""
End Sub

Public Sub fraEdition_Form_Display()
Dim X As String

If Edition_Gestion_Aut.Valider Then
    fraEdition_Form.Enabled = True
    fraEdition_Form.BackColor = &HF0FFFF '
    usrColor_Container fraEdition_Form, fraEdition_Form.BackColor
Else
    lstSelect.Enabled = True
End If
X = meElpTable.Memo
rsEdition_Form_GetBuffer X, meEdition_Form

libEdition_form = meElpTable.K2 & " " & meElpTable.Name
chkEdition_Courrier = meEdition_Form.Courrier
chkEdition_Orientation = meEdition_Form.Orientation
chkEdition_Duplex = meEdition_Form.Duplex - 1
chkEdition_Hold = meEdition_Form.Hold
chkEdition_Save = meEdition_Form.Save
chkEdition_PrinterUnit = meEdition_Form.PrinterUnit
txtEdition_LinePerPage = meEdition_Form.LinePerPage
txtEdition_FontSize = meEdition_Form.FontSize
txtEdition_Copies = meEdition_Form.Copies
chkEdition_NoPaper_Prod = meEdition_Form.NoPaper_Prod

cbo_Scan Trim(meEdition_Form.FontName), cboEdition_fontName
cbo_Scan Trim(meEdition_Form.Filigrane), cboEdition_Filigrane
cbo_Scan Trim(meEdition_Form.Unit), cboEdition_Unit
cbo_Scan Trim(meEdition_Form.PaperBin), cboEdition_PaperBin

Select Case meEdition_Form.Filigrane
    Case "A": X = "Automatique"
    Case "B": X = "BIA"
    Case Else:    X = "(non)"
End Select
cbo_Scan X, cboEdition_Filigrane


End Sub
Public Sub fraUsr_Détail_Display()
Dim X As String

If Edition_Gestion_Aut.Valider Then
    fraUsr_Détail.Enabled = True
    fraUsr_Détail.BackColor = &HF0FFFF '
    usrColor_Container fraUsr_Détail, fraUsr_Détail.BackColor
Else
    lstUsr.Enabled = True
End If
libUsr_Détail = Trim(lstUsr.Text)
cbo_Scan Trim(meUser.Unit), cboUsr_Unit

Select Case meUser.ProdTest
    Case "P": optUsr_Production = True
    Case "T": optUsr_Test = True
    Case Else: optUsr_Informatique = True
End Select

 chkUsr_Hold = meUser.Edition_Hold
 chkUsr_SplAut = meUser.Edition_Aut
 chkUsr_QSYSOPR = meUser.QSYSOPR
 txtUsr_ClasseAut = meUser.XXXXXX
 cbo_Scan Trim(meUser.Printer), cboUsr_Printer
 txtUsr_AliasWin = Trim(meUser.AliasWin)
End Sub


Private Sub txtUsr_AliasWin_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtUsr_ClasseAut_GotFocus()
Call txt_GotFocus(txtUsr_ClasseAut)

End Sub


Private Sub txtUsr_ClasseAut_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub

Private Sub txtUsr_ClasseAut_LostFocus()
Call txt_LostFocus(txtUsr_ClasseAut)

End Sub


