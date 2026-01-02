VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGAdresse 
   AutoRedraw      =   -1  'True
   Caption         =   "Gestion des adresses"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6765
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C0FF&
      Caption         =   "en &Attente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
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
      Height          =   500
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   6120
      TabIndex        =   15
      Top             =   0
      Width           =   2745
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8880
      Picture         =   "GAdresse.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   0
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   0
      TabIndex        =   16
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   706
      TabCaption(0)   =   "Liste des adresses"
      TabPicture(0)   =   "GAdresse.frx":0102
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Adesse"
      TabPicture(1)   =   "GAdresse.frx":011E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraGAdresse"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraGEntité"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fraGEntité 
         Height          =   1215
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   9135
         Begin VB.Frame fraGEntitéD 
            Height          =   975
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   6735
            Begin VB.TextBox txtSéquence 
               Height          =   285
               Left            =   5880
               MaxLength       =   2
               TabIndex        =   41
               Top             =   600
               Width           =   495
            End
            Begin VB.ComboBox cboNature 
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   3000
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   600
               Width           =   2295
            End
            Begin VB.TextBox txtCompte 
               Height          =   285
               Left            =   120
               TabIndex        =   37
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox txtDevise 
               Height          =   285
               Left            =   1800
               TabIndex        =   35
               Top             =   600
               Width           =   615
            End
            Begin VB.Label lblSéquence 
               Caption         =   "Séquence"
               Height          =   255
               Left            =   5760
               TabIndex        =   39
               Top             =   240
               Width           =   855
            End
            Begin VB.Label lblNature 
               Caption         =   "Nature"
               Height          =   255
               Left            =   3000
               TabIndex        =   38
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblCompte 
               Caption         =   "Compte"
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Width           =   975
            End
            Begin VB.Label lblDevide 
               Caption         =   "Devise"
               Height          =   255
               Left            =   1800
               TabIndex        =   34
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.ComboBox cboGAdresseId 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label lblAdresseId 
            Caption         =   "Réf adresse"
            Height          =   255
            Left            =   7440
            TabIndex        =   31
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame fraGAdresse 
         Height          =   4335
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   9135
         Begin VB.CheckBox chkRoutage 
            Alignment       =   1  'Right Justify
            Caption         =   "Envoi du courrier"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtL1 
            Height          =   285
            Left            =   1920
            TabIndex        =   5
            Top             =   840
            Width           =   6975
         End
         Begin VB.TextBox txtL0 
            Height          =   285
            Left            =   1920
            TabIndex        =   4
            Top             =   360
            Width           =   6975
         End
         Begin VB.CheckBox chkL0 
            Alignment       =   1  'Right Justify
            Caption         =   "Complément"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Width           =   1575
         End
         Begin VB.Frame fraGAdresseD 
            Height          =   2535
            Left            =   120
            TabIndex        =   20
            Top             =   1680
            Width           =   8895
            Begin VB.TextBox txtCP 
               Height          =   285
               Left            =   1800
               MaxLength       =   5
               TabIndex        =   8
               Top             =   1560
               Width           =   855
            End
            Begin VB.TextBox txtPays 
               Height          =   285
               Left            =   1800
               TabIndex        =   10
               Top             =   2040
               Width           =   495
            End
            Begin VB.TextBox txtL4 
               Height          =   285
               Left            =   3000
               TabIndex        =   9
               Top             =   1560
               Width           =   5775
            End
            Begin VB.TextBox txtL3 
               Height          =   285
               Left            =   1800
               TabIndex        =   7
               Top             =   960
               Width           =   6975
            End
            Begin VB.TextBox txtL2 
               Height          =   285
               Left            =   1800
               TabIndex        =   6
               Top             =   360
               Width           =   6975
            End
            Begin VB.Label libPays 
               Caption         =   "-"
               Height          =   255
               Left            =   3000
               TabIndex        =   26
               Top             =   2040
               Width           =   5655
            End
            Begin VB.Label lblPays 
               Caption         =   "Pays"
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label lblL4 
               Caption         =   "Code postal commune"
               Height          =   495
               Left            =   120
               TabIndex        =   24
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label lblL3 
               Caption         =   "N° voie"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lblL2 
               Caption         =   "Lieu-dit,résidence"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.Label lblL0 
            Caption         =   "Nom"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame fraSelect 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   9135
         Begin VB.Frame fraG 
            Height          =   2655
            Left            =   4560
            TabIndex        =   42
            Top             =   720
            Width           =   4335
            Begin VB.Label libGadressePaysLibellé 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1080
               TabIndex        =   50
               Top             =   2160
               Width           =   3045
            End
            Begin VB.Label libGadressePays 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   120
               TabIndex        =   49
               Top             =   2160
               Width           =   645
            End
            Begin VB.Label libGadresseL4 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1080
               TabIndex        =   48
               Top             =   1680
               Width           =   3015
            End
            Begin VB.Label libGadresseCodePostal 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   120
               TabIndex        =   47
               Top             =   1680
               Width           =   615
            End
            Begin VB.Label libGadresseL3 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   120
               TabIndex        =   46
               Top             =   1320
               Width           =   4005
            End
            Begin VB.Label libGadresseL2 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   120
               TabIndex        =   45
               Top             =   960
               Width           =   4005
            End
            Begin VB.Label libGadresseL1 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   120
               TabIndex        =   44
               Top             =   600
               Width           =   4005
            End
            Begin VB.Label libGadresseL0 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   4005
            End
         End
         Begin VB.TextBox txtSelect 
            Height          =   375
            Left            =   1080
            TabIndex        =   0
            Top             =   240
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid fgEntité 
            Height          =   2490
            Left            =   120
            TabIndex        =   1
            Top             =   840
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   4392
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
            FormatString    =   $"GAdresse.frx":013A
         End
         Begin MSFlexGridLib.MSFlexGrid fgAdresse 
            Height          =   1890
            Left            =   120
            TabIndex        =   2
            Top             =   3480
            Width           =   8835
            _ExtentX        =   15584
            _ExtentY        =   3334
            _Version        =   393216
            Rows            =   1
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
            AllowUserResizing=   3
            FormatString    =   $"GAdresse.frx":01D4
         End
         Begin VB.Label libSelect 
            Caption         =   "-"
            Height          =   255
            Left            =   2520
            TabIndex        =   28
            Top             =   360
            Width           =   4815
         End
         Begin VB.Label lblSelect 
            Caption         =   "Racine"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   735
         End
      End
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
      Height          =   405
      Left            =   3600
      TabIndex        =   17
      Top             =   0
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuGAdresse_Test 
         Caption         =   "Tester une adresse"
      End
      Begin VB.Menu mnuAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuGAdresse 
      Caption         =   "mnuGAdresse"
      Visible         =   0   'False
      Begin VB.Menu mnuGAdresse_Update 
         Caption         =   "modifier une adresse"
      End
      Begin VB.Menu mnuGAdresse_Delete 
         Caption         =   "Effacer une adresse"
      End
   End
   Begin VB.Menu mnuGEntité 
      Caption         =   "mnuGEntité"
      Visible         =   0   'False
      Begin VB.Menu mnuGEntité_AddNew 
         Caption         =   "Ajouter un lien vers une adresse"
      End
      Begin VB.Menu mnuGEntité_Delete 
         Caption         =   "Effacer un lien vers une adresse"
      End
      Begin VB.Menu mnuGEntité_Update 
         Caption         =   "Modifierun lien vers une adresse"
      End
      Begin VB.Menu mnuX4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGAdresse_AddNew 
         Caption         =   "Ajouter une adresse"
      End
   End
End
Attribute VB_Name = "frmGAdresse"
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
Dim GAdresseAut As typeAuthorization

Dim recTable As typeElpTable
Dim mClasseInfo As String

Dim recCompte As typeCompte
Dim recRacine As typeRacine, mRacineX5 As String * 5

Dim fgadresse_FormatString As String, fgadresse_K As Integer
Dim fgadresse_RowDisplay As Integer, fgadresse_RowClick As Integer
Dim fgadresse_ColorClick As Long, fgadresse_ColorDisplay As Long
Dim fgadresse_Sort1 As Integer, fgadresse_Sort2 As Integer
Dim fgadresse_SortAD As Integer, fgadresse_Sort1_Old As Integer
Dim fgadresse_BackColorFixed As Long, fgadresse_BackColor As Long

Dim fgEntité_FormatString As String, fgEntité_K As Integer
Dim fgEntité_RowDisplay As Integer, fgEntité_RowClick As Integer
Dim fgEntité_ColorClick As Long, fgEntité_ColorDisplay As Long
Dim fgEntité_Sort1 As Integer, fgEntité_Sort2 As Integer
Dim fgEntité_SortAD As Integer, fgEntité_Sort1_Old As Integer
Dim fgEntité_BackColorFixed As Long, fgEntité_BackColor As Long

Dim recGEntité As typeGEntité, xGEntité As typeGEntité, mGEntité As typeGEntité
Dim arrGEntité() As typeGEntité
Dim arrGEntité_NB As Integer, arrGEntité_Index As Integer, arrGEntité_NBMax As Integer

Dim recGAdresse As typeGAdresse, mGAdresse As typeGAdresse
Dim arrGAdresse() As typeGAdresse
Dim arrGAdresse_NB As Integer, arrGAdresse_Index As Integer, arrGAdresse_NBMax As Integer
Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnfgEntité_DisplayLine As Boolean, blnfgAdresse_DisplayLine As Boolean

Dim recGAdresseX As typeGAdresseX
Public Sub fgEntité_Load()

ReDim arrGEntité(1)

recGEntité_Init recGEntité
recGEntité.Method = "SnapLC"
recGEntité.ClasseInfo = mClasseInfo
recGEntité.Compte = mRacineX5

arrGEntité(0) = recGEntité
arrGEntité(0).Compte = mRacineX5 & "999999"
arrGEntité(0).Devise = "999"

Call srvGEntité_Load(recGEntité, arrGEntité(0))

arrGEntité_NB = srvGEntité.arrGEntité_NB
ReDim arrGEntité(arrGEntité_NB)
For I = 1 To arrGEntité_NB
    arrGEntité(I) = srvGEntité.arrGEntité(I)
Next I

fgEntité_SortAD = 0
fgEntité_Display
End Sub
Private Sub fgEntité_Display()
Dim K2 As Integer, I As Integer
Dim curDB As Currency, curCR As Currency, curX As Currency

SSTab1.Tab = 0

fgEntité_Reset
fgEntité.Visible = True
fgEntité.Enabled = True
For arrGEntité_Index = 1 To arrGEntité_NB
    If arrGEntité(arrGEntité_Index).Method <> constIgnore And arrGEntité(arrGEntité_Index).Method <> constDelete Then
        fgEntité.Rows = fgEntité.Rows + 1
        fgEntité.Row = fgEntité.Rows - 1
        recGEntité = arrGEntité(arrGEntité_Index)
        fgEntité_DisplayLine
    End If
Next arrGEntité_Index

If fgEntité.Rows > 1 Then fgEntité_Sort

End Sub
Public Sub fgEntité_Sort()
If fgEntité.Rows > 1 Then
    fgEntité.Row = 1
    fgEntité.RowSel = fgEntité.Rows - 1
    
    If fgEntité_Sort1_Old = fgEntité_Sort1 Then
        If fgEntité_SortAD = 5 Then
            fgEntité_SortAD = 6
        Else
            fgEntité_SortAD = 5
        End If
    Else
        fgEntité_SortAD = 5
    End If
    fgEntité_Sort1_Old = fgEntité_Sort1
    
    fgEntité.Col = fgEntité_Sort1
    fgEntité.ColSel = fgEntité_Sort2
    fgEntité.Sort = fgEntité_SortAD
End If

End Sub


Public Sub fgadresse_Sort()
If fgAdresse.Rows > 1 Then
    fgAdresse.Row = 1
    fgAdresse.RowSel = fgAdresse.Rows - 1
    If fgadresse_Sort1_Old = fgadresse_Sort1 Then
        If fgadresse_SortAD = 5 Then
            fgadresse_SortAD = 6
        Else
            fgadresse_SortAD = 5
        End If
    Else
        fgadresse_SortAD = 5
    End If
    fgadresse_Sort1_Old = fgadresse_Sort1
    
    fgAdresse.Col = fgadresse_Sort1
    fgAdresse.ColSel = fgadresse_Sort2
    fgAdresse.Sort = fgadresse_SortAD
End If
    

End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
blnControl = False
    If currentAction <> "" Then
        currentAction = ""
        cmdContext.Caption = constcmdRechercher
        fgEntité.Enabled = True
        fgAdresse.Enabled = True
        cmdReset
        SSTab1.Tab = 0
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


'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me

End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
libGadresseL0 = "": libGadresseL0.BackColor = MouseMoveUsr.BackColor
libGadresseL1 = "": libGadresseL1.BackColor = MouseMoveUsr.BackColor
libGadresseL2 = "": libGadresseL2.BackColor = MouseMoveUsr.BackColor
libGadresseL3 = "": libGadresseL3.BackColor = MouseMoveUsr.BackColor
libGadresseL4 = "": libGadresseL4.BackColor = MouseMoveUsr.BackColor
libGadresseCodePostal = "": libGadresseCodePostal.BackColor = MouseMoveUsr.BackColor
libGadressePays = "": libGadressePays.BackColor = MouseMoveUsr.BackColor
libGadressePaysLibellé = "": libGadressePaysLibellé.BackColor = MouseMoveUsr.BackColor

SSTab1.Tab = 0
fraSelect.Enabled = True
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
cmdOk.Caption = constValider: cmdOk.Visible = False
cmdSave.Caption = constEnAttente: cmdSave.Visible = False
arrTag_Set False
libSelect = ""
lstErr.Visible = False
blncmdOk_Visible = False: blncmdSave_Visible = False
blnfgEntité_DisplayLine = False: blnfgAdresse_DisplayLine = False
cmdReset_fraGAdresse

fraGAdresse.Enabled = False
blnControl = True

End Sub


Public Sub fgEntité_Reset()
fgEntité.Clear
fgEntité.Rows = 1
fgEntité.Row = 0
fgEntité_RowDisplay = 0
fgEntité_RowClick = 0
fgEntité.FormatString = fgEntité_FormatString
fgEntité_Sort1 = 0: fgEntité_Sort2 = 0

End Sub

Public Sub fgadresse_Reset()
fgAdresse.Clear
fgAdresse.Rows = 1
fgAdresse.Row = 0
fgadresse_RowDisplay = 0
fgadresse_RowClick = 0
fgAdresse.FormatString = fgadresse_FormatString
fgadresse_Sort1 = 0: fgadresse_Sort2 = 0


End Sub

Public Sub Form_Init()
Call BiaPgmAut_Init("CPT_ADRESSE", GAdresseAut)
fgadresse_FormatString = fgAdresse.FormatString
'fgadresse_BackColorFixed = fgAdresse.BackColorFixed
fgadresse_BackColor = fgAdresse.BackColor
fgadresse_BackColorFixed = fgAdresse.BackColorFixed
fgEntité_FormatString = fgEntité.FormatString
fgEntité_BackColorFixed = fgEntité.BackColorFixed
fgEntité_BackColor = fgEntité.BackColor

'''tableElpTable_Open
ReDim arrGEntité(1)
ReDim arrGAdresse(1)

fgEntité_Reset
fgadresse_Reset

txtSelect = "": mRacineX5 = ""
recRacineInit recRacine

mClasseInfo = "Adr"

recElpTable_Init recElpTable
recElpTable.Id = "GEntité"
recElpTable.K1 = "Nature_" & mClasseInfo
Call cbo_Load(recElpTable, cboNature, 3)


cmdReset

End Sub

Private Sub cboGAdresseId_Click()
If blnControl Then cmdControl

End Sub

Private Sub cboNature_Click()
If blnControl Then cmdControl

End Sub


Private Sub cboNature_GotFocus()
lblNature.ForeColor = warnUsrColor
End Sub

Private Sub cboNature_LostFocus()
lblNature.ForeColor = lblUsr.ForeColor
If blnControl Then cmdControl
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


Private Sub chkL0_Click()

If blnControl Then cmdControl

End Sub

Private Sub chkL0_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkL0

End Sub


Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub


Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext
End Sub

Private Sub cmdOk_Click()
If currentAction = "GAdresse_Test" Then recGadresseX_Load: Exit Sub
        
cmdControl
If lstErr.ListCount <> 0 Then Exit Sub
cmdOk.Visible = False
frmGAdresse.Enabled = False
Select Case currentAction
    Case "GAdresse_AddNew"
        V = srvGEntité_Update(recGEntité)
        If IsNull(V) Then Call srvGAdresse_Update(recGAdresse)
        fgEntité_Load
        fgadresse_Load
        libGAdresse_Display
    Case "GEntité_AddNew", "GEntité_Delete"
        V = srvGEntité_Update(recGEntité)
        fgEntité_Load
        libGAdresse_Display
    Case "GEntité_Update"
        V = srvGEntité_Update(recGEntité)
        If IsNull(V) Then arrGEntité(arrGEntité_Index) = recGEntité:  fgEntité_Display
        libGAdresse_Display
    Case "GAdresse_Update"
        V = srvGAdresse_Update(recGAdresse)
        If IsNull(V) Then arrGAdresse(arrGAdresse_Index) = recGAdresse:  fgAdresse_Display
    Case "GAdresse_Delete"
        Call srvGAdresse_Update(recGAdresse)
        fgadresse_Load
    Case "GAdresse_Test"
        recGadresseX_Load
    Case Else
        Call lstErr_AddItem(lstErr, cmdContext, "? cmdOk : " & currentAction)
End Select

cmdReset
frmGAdresse.Enabled = True
AppActivate frmGAdresse.Caption
End Sub

Private Sub cmdPrint_Click()
'Me.PopupMenu mnucmdPrint, vbPopupMenuLeftButton
End Sub

Private Sub cmdSave_Click()
cmdControl
lstErr.Clear
frmGAdresse.Enabled = False

'If lstErr.ListCount = 0 Then cmdSave_Db
frmGAdresse.Enabled = True
End Sub

Private Sub fgAdresse_GotFocus()

fgAdresse.BackColorFixed = focusUsr.BackColor
fgAdresse.BackColor = fgadresse_BackColor


End Sub

Private Sub fgAdresse_LostFocus()
fgAdresse.BackColorFixed = fgadresse_BackColorFixed

End Sub


Private Sub fgadresse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y <= fgAdresse.RowHeightMin Then
    Select Case fgAdresse.Col
        Case 0: fgadresse_Sort1 = 0: fgadresse_Sort2 = 0: fgadresse_Sort
        Case 1: fgadresse_Sort1 = 1: fgadresse_Sort2 = 1: fgadresse_Sort
        Case 2: fgadresse_Sort1 = 2: fgadresse_Sort2 = 2: fgadresse_Sort
        Case 3: fgadresse_Sort1 = 3: fgadresse_Sort2 = 3: fgadresse_Sort
        Case 4: fgadresse_Sort1 = 4: fgadresse_Sort2 = 4: fgadresse_Sort
        Case 5: fgadresse_Sort1 = 5: fgadresse_Sort2 = 5: fgadresse_Sort
        Case 6: fgadresse_Sort1 = 6: fgadresse_Sort2 = 6: fgadresse_Sort
        Case 7: fgadresse_Sort1 = 7: fgadresse_Sort2 = 7: fgadresse_Sort
   End Select
Else
    mnuGAdresse_Update = False
    mnuGAdresse_Delete = False
    If fgAdresse.Rows > 1 Then
        Call fgadresse_Color(fgadresse_RowClick, MouseMoveUsr.BackColor, fgadresse_ColorClick)
        fgAdresse.Col = fgAdresse.Cols - 1
        arrGAdresse_Index = Val(fgAdresse.Text)
        recGAdresse = arrGAdresse(arrGAdresse_Index)
        
        mnuGAdresse_Update = GAdresseAut.Saisir
        mnuGAdresse_Delete = GAdresseAut.Saisir
    End If
    Me.PopupMenu mnuGAdresse, vbPopupMenuLeftButton
End If

End Sub


Private Sub fgEntité_GotFocus()
fgEntité.BackColorFixed = focusUsr.BackColor
fgEntité.BackColor = fgEntité_BackColor
End Sub

Private Sub fgEntité_LostFocus()
fgEntité.BackColorFixed = fgEntité_BackColorFixed

End Sub


Private Sub fgEntité_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xStatut As String
If Y <= fgEntité.RowHeightMin Then
    Select Case fgEntité.Col
        Case 0: fgEntité_Sort1 = 0: fgEntité_Sort2 = 0: fgEntité_Sort
        Case 1: fgEntité_Sort1 = 1: fgEntité_Sort2 = 1: fgEntité_Sort
        Case 2: fgEntité_Sort1 = 2: fgEntité_Sort2 = 2: fgEntité_Sort
        Case 3: fgEntité_Sort1 = 3: fgEntité_Sort2 = 3: fgEntité_Sort
        Case 4: fgEntité_Sort1 = 4: fgEntité_Sort2 = 4: fgEntité_Sort
    End Select
Else
    mnuGEntité_AddNew = False
    mnuGEntité_Update = False
    mnuGEntité_Delete = False
    mnuGAdresse_AddNew = False

    If fgEntité.Rows > 1 Then
        fgEntité.Col = fgEntité.Cols - 1
        Call fgEntité_Color(fgEntité_RowClick, MouseMoveUsr.BackColor, fgEntité_ColorClick)
        fgEntité.Col = fgEntité.Cols - 1
        arrGEntité_Index = Val(fgEntité.Text)
        recGEntité = arrGEntité(arrGEntité_Index)
        
        arrGAdresse_Index = fgAdresse_Scan(recGEntité.AdresseId)
        If arrGAdresse_Index > 0 Then
            recGAdresse = arrGAdresse(arrGAdresse_Index)
            libGAdresse_Display
        End If
        
        mnuGEntité_AddNew = GAdresseAut.Saisir
        mnuGEntité_Update = GAdresseAut.Saisir
        If recGEntité.Nature <> "Fis" Then mnuGEntité_Delete = GAdresseAut.Saisir
        mnuGAdresse_AddNew = GAdresseAut.Saisir
    
    End If
    Me.PopupMenu mnuGEntité, vbPopupMenuLeftButton
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
cmdReset

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub

Private Sub fraGAdresse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraGAdresseD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub mnuAbandonner_Click()
cmdContext_Quit
End Sub



Private Sub mnuGAdresse_AddNew_Click()

blnControl = False
cmdReset_fraGAdresse

fraGAdresseD.Enabled = True
fraGAdresse.Enabled = True
fraGEntitéD.Enabled = True
cboGAdresseId.Enabled = False

recGEntité_Init recGEntité
recGEntité.Method = constAddNew
recGEntité.Compte = mRacineX5
recGEntité.ClasseInfo = "Adr"
recGEntité.AdresseAmjDébut = "00000000"
recGEntité.AdresseAmjfin = "00000000"
fraGEntité_Display

recGAdresse_Init mGAdresse
recGAdresse.Method = constAddNew
recGAdresse.IdRéférence = recGEntité.AdresseId
fraGAdresse_Display_cbo
currentAction = "GAdresse_AddNew"

cmdContext_Tab1

End Sub

Private Sub mnuGAdresse_Delete_Click()
blnControl = False
cmdReset_fraGAdresse
fraGAdresse_Display

lstErr.Clear
recGEntité.AdresseId = recGAdresse.IdRéférence
Call cbo_Scan(recGEntité.AdresseId, cboGAdresseId)

For I = 1 To arrGEntité_NB
    If recGAdresse.IdRéférence = arrGEntité(I).AdresseId Then Call lstErr_AddItem(lstErr, cmdContext, "? adresse reliée : " & arrGEntité(I).Nature)

Next I
If lstErr.ListCount = 0 Then
    currentAction = "GAdresse_Delete"
    recGAdresse.Method = constDelete
    cmdOk.Caption = "Effacer"
    fraGEntité.Visible = False
    cmdOk.Visible = True
    cmdContext_Tab1
End If

End Sub

Private Sub mnuGAdresse_Test_Click()
blnControl = False
cmdReset
cmdReset_fraGAdresse

currentAction = "GAdresse_Test"
recGAdresseX.Method = "SeekP0"
cmdOk.Caption = "OK"
fraGEntitéD.Enabled = True

fraGAdresse.Enabled = False

cmdContext_Tab1
blnControl = False
cmdOk.Visible = True

End Sub

Private Sub mnuGAdresse_Update_Click()
blnControl = False
cmdReset_fraGAdresse

currentAction = "GAdresse_Update"
recGAdresse.Method = constUpdate
cmdOk.Caption = "Mise à jour A"
fraGEntité.Visible = False

txtL1.Enabled = False
chkL0.Enabled = False
chkRoutage.Enabled = False
fraGAdresseD.Enabled = True
fraGAdresse.Enabled = True
fraGAdresse_Display_cbo

cmdContext_Tab1
End Sub

Private Sub mnuGEntité_AddNew_Click()
blnControl = False
cmdReset_fraGAdresse

fraGAdresse.Enabled = True
txtL0.Enabled = False
txtL1.Enabled = True
chkL0.Enabled = True

recGEntité_Init recGEntité
recGEntité.Method = constAddNew
recGEntité.Compte = mRacineX5
recGEntité.ClasseInfo = "Adr"
recGEntité.AdresseAmjDébut = "00000000"
recGEntité.AdresseAmjfin = "00000000"
fraGEntité_Display
fraGEntitéD.Enabled = True

currentAction = "GEntité_AddNew"
cboGAdresseId.Enabled = True
cmdContext_Tab1

End Sub


Private Sub mnuGEntité_Delete_Click()

blnControl = False
cmdReset_fraGAdresse
arrGAdresse_Index = fgAdresse_Scan(recGEntité.AdresseId)
If arrGAdresse_Index < 0 Then Exit Sub
recGAdresse = arrGAdresse(arrGAdresse_Index)
fraGAdresse_Display_cbo

currentAction = "GEntité_Delete"
recGEntité.Method = constDelete
fraGEntité_Display
cmdOk.Caption = "Effacer"
cmdOk.Visible = True
cmdContext_Tab1

End Sub


Private Sub mnuGEntité_Update_Click()

blnControl = False
cmdReset_fraGAdresse
arrGAdresse_Index = fgAdresse_Scan(recGEntité.AdresseId)
If arrGAdresse_Index < 0 Then Exit Sub
recGAdresse = arrGAdresse(arrGAdresse_Index)
fraGAdresse_Display_cbo

currentAction = "GEntité_Update"
recGEntité.Method = constUpdate
fraGEntité_Display

fraGAdresse.Enabled = True
cboGAdresseId.Enabled = True
txtL0.Enabled = False
cmdOk.Caption = "Mise à jour"

cmdContext_Tab1
End Sub

Private Sub mnuQuitter_Click()
Unload Me
End Sub







Public Sub cmdControl()
Dim X As String, I As Integer

If Not Me.Enabled Then Exit Sub
Me.Enabled = False

cmdOk.Visible = False
cmdSave.Visible = False
blnControl = False

lstErr.Clear
lstErr.Height = 200

If recGEntité.Method = constAddNew Then
    cbo_Value recGEntité.Nature, cboNature

    recGEntité.Compte = Trim(txtCompte)
    recGEntité.Devise = Trim(txtDevise)
    If recGEntité.Nature = "Tit" Then
        txtSéquence.Visible = True
        I = Val(Trim(txtSéquence))
    Else
        txtSéquence.Visible = False
        I = 0
    End If
    
        If I > 99 Then
            Call lstErr_AddItem(lstErr, cmdContext, "? Séquence 0 à 99 : " & I)
        Else
            recGEntité.Séquence = I
        End If
End If

If recGAdresse.Method = constAddNew Then
    recGEntité.AdresseId = ""
    Mid$(recGEntité.AdresseId, 1, 11) = recGEntité.Compte
    Mid$(recGEntité.AdresseId, 12, 3) = recGEntité.Devise
    Mid$(recGEntité.AdresseId, 15, 3) = recGEntité.Nature
    If recGEntité.Séquence = 0 Then
        Mid$(recGEntité.AdresseId, 18, 2) = "  "
    Else
        Mid$(recGEntité.AdresseId, 18, 2) = Format$(recGEntité.Séquence, "00")
    End If
    recGAdresse.IdRéférence = recGEntité.AdresseId
    cbo_Scan recGEntité.AdresseId, cboGAdresseId
    If cboGAdresseId.ListIndex <> -1 Then Call lstErr_AddItem(lstErr, cmdContext, "? Relation existe déjà ; " & recGEntité.AdresseId)

Else
    cbo_Value recGEntité.AdresseId, cboGAdresseId
    If recGEntité.AdresseId <> recGAdresse.IdRéférence Then
        I = fgAdresse_Scan(recGEntité.AdresseId)
        If I > -1 Then
            recGAdresse = arrGAdresse(I)
            fraGAdresse_Display_cbo
        Else
            Call lstErr_AddItem(lstErr, cmdContext, "? GAdresseId ; " & recGEntité.AdresseId)
        End If
    End If
End If

recGEntité.AdresseL1 = Trim(txtL1)
If chkRoutage = "0" Then
    recGEntité.AdresseRoutage = "0"
Else
    recGEntité.AdresseRoutage = "1"
End If

chkL0_Label
If chkL0 = "0" Then
    recGEntité.AdresseL0K = "0"
Else
    recGEntité.AdresseL0K = "1"
End If


X = Trim(txtL0): recGAdresse.L0 = X
If X = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? Préciser le nom")

recGAdresse.L2 = Trim(txtL2)
recGAdresse.L3 = Trim(txtL3)
X = Trim(txtL4): recGAdresse.L4 = X
If X = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? Préciser la commune")

recGAdresse.CodePostal = Trim(txtCP)
X = Trim(txtPays): recGAdresse.Pays = X
If X = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? Préciser le pays")

'Select Case currentAction
'    Case constValider
'            V = fctGEntité_Compare(recGEntité, mGEntité)
'            If Not IsNull(V) Then
'                Call MsgBox("L'enregistrement après contrôle est différent de l'enregistrement lu :" & Chr$(13) & V, vbCritical, "me : cmdControl")
'                Call lstErr_AddItem(lstErr, cmdContext, "? Erreur Contrôle validation")
'            End If
'End Select

If lstErr.ListCount = 0 Then
    cmdOk.Visible = blncmdOk_Visible
End If

ExitSub:

Me.Enabled = True
'If cmdOk.Visible Then cmdOk.SetFocus
    
blnControl = True


End Sub

Public Sub fgAdresse_Display()
Dim I As Integer
cboGAdresseId.Clear
fgAdresse.Visible = True
fgadresse_Reset
fgAdresse.Enabled = True
For arrGAdresse_Index = 1 To arrGAdresse_NB
    recGAdresse = arrGAdresse(arrGAdresse_Index)
    fgAdresse.Rows = fgAdresse.Rows + 1
    fgAdresse.Row = fgAdresse.Rows - 1
    fgAdresse_DisplayLine
    cboGAdresseId.AddItem recGAdresse.IdRéférence

Next arrGAdresse_Index

fgadresse_K = fgAdresse.Cols
 
If fgAdresse.Rows > 1 Then fgadresse_Sort
End Sub

Public Sub fgAdresse_DisplayLine()

fgAdresse.Col = 0: fgAdresse.Text = recGAdresse.IdRéférence
fgAdresse.Col = 1: fgAdresse.Text = recGAdresse.L0
fgAdresse.Col = 2: fgAdresse.Text = recGAdresse.L2
fgAdresse.Col = 3: fgAdresse.Text = recGAdresse.L3
fgAdresse.Col = 4: fgAdresse.Text = recGAdresse.CodePostal
fgAdresse.Col = 5: fgAdresse.Text = recGAdresse.L4
fgAdresse.Col = 6: fgAdresse.Text = recGAdresse.Pays
fgAdresse.Col = fgAdresse.Cols - 1: fgAdresse.Text = arrGAdresse_Index

End Sub

Public Sub fgEntité_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgEntité.Row

If lRow > 0 Then
    fgEntité.Row = lRow
    For I = 0 To fgEntité.Col - 1
        fgEntité.Col = I: fgEntité.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgEntité.Row = mRow
    If fgEntité.Row > 0 Then
        lRow = fgEntité.Row
        lColor_Old = fgEntité.CellBackColor
        For I = 0 To fgEntité.Col - 1
          fgEntité.Col = I: fgEntité.CellBackColor = lColor
        Next I
        fgEntité.Col = 0
    End If
End If

End Sub

Public Sub fgadresse_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgAdresse.Row


If lRow > 0 Then
    fgAdresse.Row = lRow
    For I = 0 To fgAdresse.Cols - 1
        fgAdresse.Col = I: fgAdresse.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgAdresse.Row = mRow
    If fgAdresse.Row > 0 Then
        lRow = fgAdresse.Row
        lColor_Old = fgAdresse.CellBackColor
        For I = 0 To fgAdresse.Cols - 1
          fgAdresse.Col = I: fgAdresse.CellBackColor = lColor
        Next I
        fgAdresse.Col = 0
    End If
End If

End Sub

Public Sub fgadresse_Load()
ReDim arrGAdresse(1)

recGAdresse_Init recGAdresse
recGAdresse.Method = "SnapP0"
recGAdresse.IdRéférence = mRacineX5

arrGAdresse(0) = recGAdresse
arrGAdresse(0).IdRéférence = mRacineX5 & "99999999999999"

Call srvGadresse_Load(recGAdresse, arrGAdresse(0))

arrGAdresse_NB = srvGAdresse.arrGAdresse_NB
ReDim arrGAdresse(arrGAdresse_NB)
For I = 1 To arrGAdresse_NB
    arrGAdresse(I) = srvGAdresse.arrGAdresse(I)
Next I

fgadresse_SortAD = 5
fgAdresse_Display

End Sub

Public Sub fgEntité_DisplayLine()
fgEntité.Col = 0: fgEntité.Text = recGEntité.Compte & recGEntité.Devise
fgEntité.Col = 1
If recGEntité.Séquence = 0 Then
    fgEntité.Text = recGEntité.Nature
Else
    fgEntité.Text = recGEntité.Nature & "_" & Format$(recGEntité.Séquence, "00")
End If

fgEntité.Col = 2: fgEntité.Text = recGEntité.AdresseId
fgEntité.Col = 3: fgEntité.Text = recGEntité.AdresseL1
fgEntité.Col = fgEntité.Cols - 1: fgEntité.Text = arrGEntité_Index

End Sub

Private Sub chkRoutage_Click()
If blnControl Then cmdControl

End Sub


Private Sub chkRoutage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkRoutage
End Sub


Private Sub txtCP_GotFocus()
txt_GotFocus txtCP
End Sub

Private Sub txtCP_LostFocus()
txt_LostFocus txtCP
If blnControl Then cmdControl


End Sub


Private Sub txtL0_GotFocus()
txt_GotFocus txtL0

End Sub

Private Sub txtL0_LostFocus()
txt_LostFocus txtL0
If blnControl Then cmdControl

End Sub


Private Sub txtL1_GotFocus()
txt_GotFocus txtL1

End Sub

Private Sub txtL1_LostFocus()
txt_LostFocus txtL1
If blnControl Then cmdControl

End Sub


Private Sub txtL2_GotFocus()
txt_GotFocus txtL2

End Sub

Private Sub txtL2_LostFocus()
txt_LostFocus txtL2
If blnControl Then cmdControl


End Sub

Private Sub txtL3_GotFocus()
txt_GotFocus txtL3

End Sub

Private Sub txtL3_LostFocus()
txt_LostFocus txtL3
If blnControl Then cmdControl


End Sub

Private Sub txtL4_GotFocus()
txt_GotFocus txtL4

End Sub

Private Sub txtL4_LostFocus()
txt_LostFocus txtL4
If blnControl Then cmdControl


End Sub

Private Sub txtPays_GotFocus()
txt_GotFocus txtPays

End Sub

Private Sub txtPays_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtPays_LostFocus()
txt_LostFocus txtPays
If blnControl Then cmdControl

End Sub

Private Sub txtSelect_Change()
If currentAction <> "txtSelect" Then
    currentAction = "txtSelect"
    cmdReset
End If
End Sub



Public Sub cmdControl_Select()
Dim V

If Not Me.Enabled Then Exit Sub
Me.Enabled = False

cmdOk.Visible = False
cmdSave.Visible = False
blnControl = False

lstErr.Clear
lstErr.Height = 200

If Trim(txtSelect) = "" Then
    Call lstErr_AddItem(lstErr, cmdContext, "? préciser la racine")
    GoTo ExitSub
End If

recRacine.Numéro = CLng(txtSelect)
If Not IsNull(srvRacineFind(recRacine)) Then
    Call lstErr_AddItem(lstErr, cmdContext, "? Racine inconnue")
Else
    currentAction = constSaisie
    mRacineX5 = Format$(recRacine.Numéro, "00000")
    libSelect = recRacine.Intitulé
    fgEntité_Load
    fgadresse_Load
    If arrGEntité_NB = 0 Then mnuGAdresse_AddNew_Fis
End If


ExitSub:

Me.Enabled = True
    
blnControl = True

End Sub

Private Sub txtSelect_GotFocus()
txt_GotFocus txtSelect
End Sub


Private Sub txtSelect_KeyPress(KeyAscii As Integer)
num_KeyAsciiD KeyAscii, txtSelect

End Sub

Private Sub txtSelect_LostFocus()
txt_LostFocus txtSelect
If blnControl Then cmdControl_Select

End Sub



Public Sub cmdReset_fraGAdresse()

chkL0 = "1": chkL0.Enabled = True
txtL0 = "": txtL0.Enabled = True
txtL1 = "": txtL1.Enabled = True
txtL2 = "": txtL3 = "": txtL4 = "": txtCP = "": txtPays = "FR"

chkRoutage = "1"
chkRoutage.Enabled = True

fraGEntité.Visible = True
fraGEntitéD.Enabled = False
cboGAdresseId.Enabled = False
fraGAdresse.Enabled = False
fraGAdresseD.Enabled = False

End Sub

Public Sub mnuGAdresse_AddNew_Fis()
mnuGAdresse_AddNew_Click
Call cbo_Scan("Fis", cboNature)
fraGEntitéD.Enabled = False

End Sub

Public Function fgAdresse_Scan(lIdRéférence As String) As Integer
Dim I As Integer
fgAdresse_Scan = -1
For I = 1 To arrGAdresse_NB
    If arrGAdresse(I).IdRéférence = lIdRéférence Then fgAdresse_Scan = I: Exit For
    
Next I
End Function

Public Function fgEntité_Scan(lAdresseId As String) As Integer
Dim I As Integer
fgEntité_Scan = -1
For I = 1 To arrGEntité_NB
    If arrGEntité(I).AdresseId = lAdresseId Then fgEntité_Scan = I: Exit For
    
Next I
End Function


Public Sub fraGAdresse_Display_cbo()
Call cbo_Scan(recGEntité.AdresseId, cboGAdresseId)
fraGAdresse_Display

End Sub
Public Sub fraGEntité_Display()

chkL0 = recGEntité.AdresseL0K
chkL0_Label
txtL1 = recGEntité.AdresseL1
chkRoutage = recGEntité.AdresseRoutage
txtCompte = recGEntité.Compte
txtDevise = recGEntité.Devise
Call cbo_Scan(recGEntité.Nature, cboNature)
txtSéquence = recGEntité.Séquence
txtSéquence.Visible = False
Call cbo_Scan(recGEntité.AdresseId, cboGAdresseId)

End Sub


Public Sub cmdContext_Tab1()
fraSelect.Enabled = False
cmdContext.Caption = constcmdAbandonner
SSTab1.Tab = 1
blncmdOk_Visible = True
blnControl = True

End Sub

Public Sub chkL0_Label()
If chkL0 = "0" Then
    lblL0 = ""
    chkL0.Caption = "Complément"
Else
    lblL0 = "Nom"
    chkL0.Caption = "Nom"
End If

End Sub

Public Sub fraGAdresse_Display()

txtL0 = recGAdresse.L0
txtL2 = recGAdresse.L2
txtL3 = recGAdresse.L3
txtL4 = recGAdresse.L4
txtCP = Trim(recGAdresse.CodePostal)
txtPays = Trim(recGAdresse.Pays)

End Sub

Public Sub libGAdresse_Display()
Dim X As String

If recGEntité.AdresseL0K = "0" Then
    libGadresseL0 = recGAdresse.L0
    libGadresseL1 = recGEntité.AdresseL1
Else
    libGadresseL0 = recGEntité.AdresseL1
    libGadresseL1 = ""
End If
libGadresseL2 = recGAdresse.L2
libGadresseL3 = recGAdresse.L3
libGadresseL4 = recGAdresse.L4
libGadresseCodePostal = recGAdresse.CodePostal
libGadressePays = recGAdresse.Pays
libGadressePaysLibellé = Trim(DicLib(919, recGAdresse.Pays))

End Sub

Public Sub recGadresseX_Load()

recGAdresseX_Init recGAdresseX

recGAdresseX.Method = "SeekP0"
recGAdresseX.Compte = Trim(txtCompte)
recGAdresseX.Devise = Trim(txtDevise)
cbo_Value recGAdresseX.Nature, cboNature
recGAdresseX.Séquence = CInt(Val(txtSéquence))

If IsNull(srvGAdresseX_Monitor(recGAdresseX)) Then

    chkRoutage = recGAdresseX.Routage
    cboGAdresseId.AddItem recGAdresseX.Compte & recGAdresseX.Devise & recGAdresseX.Nature & recGAdresseX.Séquence
    txtL0 = recGAdresseX.L0
    txtL2 = recGAdresseX.L2
    txtL3 = recGAdresseX.L3
    txtL4 = recGAdresseX.L4
    txtCP = Trim(recGAdresseX.CodePostal)
    txtPays = Trim(recGAdresseX.Pays)
    libPays = recGAdresseX.PaysLibellé
End If

End Sub
