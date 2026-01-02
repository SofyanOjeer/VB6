VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSAB_CRE 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_CRE : Crédits "
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13560
   Icon            =   "SAB_CRE.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9150
   ScaleWidth      =   13560
   Begin VB.Frame fraSelect_Options 
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
      Height          =   5970
      Left            =   480
      TabIndex        =   13
      Top             =   1440
      Width           =   7575
      Begin VB.CheckBox chkSelect_Print 
         Caption         =   "Rechercher et imprimer les avis "
         Height          =   195
         Left            =   360
         TabIndex        =   27
         Top             =   1320
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.OptionButton optSelect_MAD 
         Caption         =   "Mise à disposition, date émission avis SAB  ( du .... au ....)"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   2280
         Width           =   4935
      End
      Begin VB.OptionButton optSelect_Avis_Echéance 
         Caption         =   "Avis d'échéance, date fin ( du .... au ....)"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   2880
         Width           =   3375
      End
      Begin VB.OptionButton optSelect_Confirmation 
         Caption         =   "Confirmation des conditions,  date début  (du .... au ....)"
         Height          =   495
         Left            =   360
         TabIndex        =   20
         Top             =   3360
         Width           =   4695
      End
      Begin VB.CheckBox chkDossier_prtNb 
         Caption         =   "Courrier en 2 exemplaires"
         Height          =   375
         Left            =   4800
         TabIndex        =   19
         Top             =   1680
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ListBox lstDossier_Contact 
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.OptionButton optSelect_Avis 
         Caption         =   "Avis échus générés par SAB  du .... au ...."
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   5280
         Width           =   3375
      End
      Begin VB.OptionButton optSelect_Dossier 
         Caption         =   "Dossiers Crédit / Prêt"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Value           =   -1  'True
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker txtAmjMin 
         Height          =   300
         Left            =   4440
         TabIndex        =   16
         Top             =   1200
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
         Format          =   121176067
         CurrentDate     =   36299
         MaxDate         =   401768
         MinDate         =   -328351
      End
      Begin MSComCtl2.DTPicker txtAmjMax 
         Height          =   300
         Left            =   6120
         TabIndex        =   17
         Top             =   1200
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
         Format          =   121176067
         CurrentDate     =   36299
         MaxDate         =   401768
         MinDate         =   -328351
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   7320
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7320
         Y1              =   960
         Y2              =   960
      End
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   4
      Top             =   0
      Width           =   5175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   2
      Top             =   500
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Rechercher"
      TabPicture(0)   =   "SAB_CRE.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Dossier"
      TabPicture(1)   =   "SAB_CRE.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDossier"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "X"
      TabPicture(2)   =   "SAB_CRE.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Détail fichiers"
      TabPicture(3)   =   "SAB_CRE.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fgDossier"
      Tab(3).ControlCount=   1
      Begin VB.Frame fraDossier 
         Height          =   8145
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   13260
         Begin VB.ListBox lstDossier_ZCREBIS0 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2760
            ItemData        =   "SAB_CRE.frx":037A
            Left            =   7320
            List            =   "SAB_CRE.frx":0381
            Style           =   1  'Checkbox
            TabIndex        =   23
            Top             =   3240
            Width           =   5685
         End
         Begin VB.Frame fraDossier_Info 
            Enabled         =   0   'False
            Height          =   7695
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   7095
            Begin VB.ListBox lstDossier_Adresse 
               Height          =   2010
               Left            =   360
               TabIndex        =   24
               Top             =   480
               Width           =   3735
            End
         End
         Begin VB.Frame fraDossier_Saisie 
            Height          =   615
            Left            =   7320
            TabIndex        =   9
            Top             =   7440
            Width           =   5895
         End
         Begin VB.ListBox lstDossier_ZCREAVI0 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1680
            ItemData        =   "SAB_CRE.frx":0391
            Left            =   7320
            List            =   "SAB_CRE.frx":0398
            Style           =   1  'Checkbox
            TabIndex        =   6
            Top             =   600
            Width           =   5685
         End
         Begin VB.Label lblDossier_ZCREBIS0 
            Caption         =   "Echéancier"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7320
            TabIndex        =   22
            Top             =   2880
            Width           =   3255
         End
         Begin VB.Label lblDossier_Avis 
            Caption         =   "Avis émis par SAB"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7320
            TabIndex        =   21
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   135
         TabIndex        =   3
         Top             =   330
         Width           =   13290
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Exécuter la requête"
            Height          =   525
            Left            =   11760
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox txtSelect 
            Height          =   285
            Left            =   135
            TabIndex        =   7
            Top             =   240
            Width           =   1230
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7185
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   13080
            _ExtentX        =   23072
            _ExtentY        =   12674
            _Version        =   393216
            Rows            =   1
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   200
            BackColor       =   14737632
            ForeColor       =   4210688
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            TextStyle       =   4
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   ">Dossier       |>Prêt       |> Nature    |>Montant              |<Etat    |>Date             ||"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgDossier 
         Height          =   7755
         Left            =   -74880
         TabIndex        =   11
         Top             =   600
         Width           =   13275
         _ExtentX        =   23416
         _ExtentY        =   13679
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         RowHeightMin    =   200
         BackColor       =   14737632
         ForeColor       =   4210688
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12648384
         BackColorBkg    =   14737632
         AllowBigSelection=   0   'False
         TextStyle       =   4
         TextStyleFixed  =   4
         FocusRect       =   2
         HighLight       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   3
         FormatString    =   $"SAB_CRE.frx":03B1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13080
      Picture         =   "SAB_CRE.frx":046B
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContext_Auto 
         Caption         =   "Auto :impression MAD, Avis d'échéance"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint0_Avis 
         Caption         =   "Imprimer les avis"
      End
   End
End
Attribute VB_Name = "frmSAB_CRE"
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
Dim SAB_CRE_Aut As typeAuthorization
Dim xAmjMin As String, xAmjMax As String
Dim IbmAmjMin As String, IbmAmjMax As String

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim fgDossier_FormatString As String, fgDossier_K As Integer
Dim fgDossier_RowDisplay As Integer, fgDossier_RowClick As Integer, fgDossier_ColClick As Integer
Dim fgDossier_ColorClick As Long, fgDossier_ColorDisplay As Long
Dim fgDossier_Sort1 As Integer, fgDossier_Sort2 As Integer
Dim fgDossier_SortAD As Integer, fgDossier_Sort1_Old As Integer
Dim fgDossier_arrIndex As Integer
Dim blnfgDossier_DisplayLine As Boolean


Dim meYBIACRE As typeYBIACRE
Dim meZCREPRE0 As typeZCREPRE0, xZCREPRE0 As typeZCREPRE0

Dim xZCREDOS0 As typeZCREDOS0
Dim xZCREPLA0 As typeZCREPLA0
Dim xZCREEVE0 As typeZCREEVE0
Dim xZCREAVI0 As typeZCREAVI0
Dim xZCREBIS0 As typeZCREBIS0
Dim xZCREEMP0 As typeZCREEMP0


Dim fgSelect_ZCREAVI0()         As typeZCREAVI0
Dim fgSelect_ZCREAVI0_Nb     As Integer
Dim fgSelect_ZCREAVI0_Max  As Integer

Dim fgSelect_ZCREEVE0()         As typeZCREEVE0
Dim fgSelect_ZCREEVE0_Nb     As Integer
Dim fgSelect_ZCREEVE0_Max  As Integer

Dim fgSelect_ZCREBIS0()         As typeZCREBIS0
Dim fgSelect_ZCREBIS0_Nb     As Integer
Dim fgSelect_ZCREBIS0_Max  As Integer

Dim lstDossier_ZCREBIS_Index() As Integer
Dim blnSelect_Print As Boolean
Private Sub fgDossier_Display()
Dim I As Integer
On Error Resume Next
SSTab1.Tab = 1
fraDossier.Enabled = True
fgDossier_Reset

fgDossier.Rows = 1
fgDossier.FormatString = fgDossier_FormatString

For I = 0 To lstDossier_ZCREAVI0.ListCount - 1
    lstDossier_ZCREAVI0.Selected(I) = False
Next I


End Sub

Public Sub fgDossier_DisplayLine(lOrigine As String, lId As String, lText As String)
On Error Resume Next
fgDossier.Rows = fgDossier.Rows + 1
fgDossier.Row = fgDossier.Rows - 1
fgDossier.Col = 0: fgDossier.Text = lOrigine
fgDossier.Col = 1: fgDossier.Text = lId
fgDossier.Col = 2: fgDossier.Text = lText
End Sub

Public Sub fgDossier_Sort()
If fgDossier.Rows > 1 Then
    fgDossier.Row = 1
    fgDossier.RowSel = fgDossier.Rows - 1
    
    If fgDossier_Sort1_Old = fgDossier_Sort1 Then
        If fgDossier_SortAD = 5 Then
            fgDossier_SortAD = 6
        Else
            fgDossier_SortAD = 5
        End If
    Else
        fgDossier_SortAD = 5
    End If
    fgDossier_Sort1_Old = fgDossier_Sort1
    
    fgDossier.Col = fgDossier_Sort1
    fgDossier.ColSel = fgDossier_Sort2
    fgDossier.Sort = fgDossier_SortAD
End If

End Sub
Public Sub fgDossier_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgDossier.Rows - 1
    fgDossier.Row = I
    fgDossier.Col = lK
    X = Format$(Val(fgDossier.Text), "0000000")
    fgDossier.Col = fgDossier_arrIndex - 1
    Select Case lK
        Case 1, 2: fgDossier.Text = X
    End Select
Next I


fgDossier_Sort1 = fgDossier_arrIndex - 1: fgDossier_Sort2 = fgDossier_arrIndex - 1
fgDossier_Sort
End Sub





Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
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
Dim V
Dim Nb As Long
Dim xSQL As String
Dim xW As String
On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgSelect_Display"

    Set rsSab = Nothing
    
    xSQL = "select CREPREDOS , CREPREPRE ,CREPRENAT , CREPREDEV , CREPREMON  , CREPRECTA  , CREPREOUV from  " & paramIBM_Library_SAB & ".ZCREPRE0"
   Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04

    Do While Not rsSab.EOF
        'Call rszCREPRE0_GetBuffer(rsSAB, meZCREPRE0)
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            fgSelect.Col = 0: fgSelect.Text = rsSab("CREPREDOS")
           fgSelect.CellForeColor = vbBlue
            fgSelect.Col = 1: fgSelect.Text = rsSab("CREPREPRE")
            fgSelect.Col = 2: fgSelect.Text = rsSab("CREPRENAT")
   
            fgSelect.Col = 3: fgSelect.Text = Format$(rsSab("CREPREMON"), "### ### ### ##0.00") & rsSab("CREPREDEV")
            fgSelect.CellForeColor = vbBlue
            
            fgSelect.Col = 4: fgSelect.Text = rsSab("CREPRECTA")
            fgSelect.Col = 5: fgSelect.Text = dateIBM10(rsSab("CREPREOUV"), True)
  
        rsSab.MoveNext
    Loop
    '$2003.11.04  rsSAB.Close
fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
fgSelect.Visible = True
fgSelect.TopRow = fgSelect.Rows - 12
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub fgSelect_Avis()
Dim V
Dim Nb As Long
Dim xSQL As String
Dim xW As String
On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

txtAMJ_Control

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgSelect_Avis"

ReDim fgSelect_ZCREAVI0(101)
fgSelect_ZCREAVI0_Nb = 0
fgSelect_ZCREAVI0_Max = 100

    Set rsSab = Nothing
    
   ' xSQL = "select CREAVIDOS , CREAVIPRE ,CREAVINAT , CREAVIDEV , CREAVIMON  ,  CREAVITYP , CREAVIDTC from ZCREAVI0"
    xSQL = "select * from  " & paramIBM_Library_SAB & ".ZCREAVI0" _
           & " where CREAVIDTC >= " & IbmAmjMin & " AND CREAVIDTC <= " & IbmAmjMax
           
   If optSelect_MAD Then xSQL = xSQL & " AND CREAVITYP = '00'"
   
   Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04

    Do While Not rsSab.EOF
    
        Call rsZCREAVI0_GetBuffer(rsSab, xZCREAVI0)
        fgSelect_ZCREAVI0_Nb = fgSelect_ZCREAVI0_Nb + 1
        If fgSelect_ZCREAVI0_Nb > fgSelect_ZCREAVI0_Max Then
            fgSelect_ZCREAVI0_Max = fgSelect_ZCREAVI0_Max + 100
            ReDim Preserve fgSelect_ZCREAVI0(fgSelect_ZCREAVI0_Max)
        End If
        
        fgSelect_ZCREAVI0(fgSelect_ZCREAVI0_Nb) = xZCREAVI0

            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            fgSelect.Col = 0: fgSelect.Text = xZCREAVI0.CREAVIDOS
           fgSelect.CellForeColor = vbBlue
            fgSelect.Col = 1: fgSelect.Text = xZCREAVI0.CREAVIPRE
            fgSelect.Col = 2: fgSelect.Text = xZCREAVI0.CREAVINAT
   
            fgSelect.Col = 3: fgSelect.Text = Format$(xZCREAVI0.CREAVIMON, "### ### ### ##0.00") & xZCREAVI0.CREAVIDEV
            fgSelect.CellForeColor = vbBlue
            
            fgSelect.Col = 4: fgSelect.Text = xZCREAVI0.CREAVITYP
            fgSelect.Col = 5: fgSelect.Text = dateIBM10(xZCREAVI0.CREAVIDTC, True)
  
        rsSab.MoveNext
    Loop
fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
fgSelect.Visible = True
If fgSelect.Rows > 12 Then fgSelect.TopRow = fgSelect.Rows - 12
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub fgSelect_ZCREBIS0_SQL()
Dim V
Dim Nb As Long, I As Long
Dim xSQL As String
Dim xW As String
Dim wK1 As Long, wK2 As Long, wIndex As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

txtAMJ_Control

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgSelect_ZCREBIS0"

ReDim fgSelect_ZCREBIS0(101)
fgSelect_ZCREBIS0_Nb = 0: fgSelect_ZCREBIS0_Max = 100

    Set rsSab = Nothing
    If optSelect_MAD Then
         xSQL = "select * from  " & paramIBM_Library_SAB & ".ZCREBIS0_BIS0001" _
               & " where CREBISEMI >= " & IbmAmjMin & " AND CREBISEMI <= " & IbmAmjMax _
               & " and CREBISTYP = '00' "
    End If
    
    If optSelect_Avis_Echéance Then
        xSQL = "select * from  " & paramIBM_Library_SAB & ".ZCREBIS0_BIS0001" _
               & " where CREBISFIN >= " & IbmAmjMin & " AND CREBISFIN <= " & IbmAmjMax _
               & " and CREBISTYP > '00'"
    End If
    
    If optSelect_Confirmation Then
        xSQL = "select * from  " & paramIBM_Library_SAB & ".ZCREBIS0_BIS0001" _
               & " where CREBISDEB >= " & IbmAmjMin & " AND CREBISDEB <= " & IbmAmjMax _
               & " and CREBISTYP > '00'"
    End If
    
  Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04

    Do While Not rsSab.EOF
    
        Call rsZCREBIS0_GetBuffer(rsSab, xZCREBIS0)
        fgSelect_ZCREBIS0_Nb = fgSelect_ZCREBIS0_Nb + 1
        If fgSelect_ZCREBIS0_Nb > fgSelect_ZCREBIS0_Max Then
            fgSelect_ZCREBIS0_Max = fgSelect_ZCREBIS0_Max + 100
            ReDim Preserve fgSelect_ZCREBIS0(fgSelect_ZCREBIS0_Max)
        End If
        
        fgSelect_ZCREBIS0(fgSelect_ZCREBIS0_Nb) = xZCREBIS0

            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            fgSelect.Col = 0: fgSelect.Text = xZCREBIS0.CREBISDOS
            fgSelect.CellForeColor = vbBlue
            fgSelect.Col = 1: fgSelect.Text = xZCREBIS0.CREBISPRE
            fgSelect.CellForeColor = vbBlue
            fgSelect.Col = 2: fgSelect.Text = xZCREBIS0.CREBISPAY
            fgSelect.Col = 3: fgSelect.Text = Format$(xZCREBIS0.CREBISMRE, "### ### ### ##0.00") & xZCREBIS0.CREBISDRE

            fgSelect.Col = 4: fgSelect.Text = xZCREBIS0.CREBISTYP
            fgSelect.Col = 5: fgSelect.Text = dateIBM10(xZCREBIS0.CREBISEMI, True)
            fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = fgSelect_ZCREBIS0_Nb

        rsSab.MoveNext
    Loop
fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
fgSelect.Visible = True
If fgSelect.Rows > 12 Then fgSelect.TopRow = fgSelect.Rows - 12

If chkSelect_Print = "1" Then
    For I = 1 To fgSelect_ZCREBIS0_Nb
        fgSelect.Row = I
        fgSelect.Col = 0: wK1 = CLng(Val(fgSelect.Text))
        fgSelect.Col = 1: wK2 = CLng(Val(fgSelect.Text))
        cmdSelect wK1, wK2
        fgSelect.Col = fgSelect_arrIndex: wIndex = CLng(Val(fgSelect.Text))
        
        blnOk = True
        Select Case meYBIACRE.ZCREDOS0(1).CREDOSNCR
            Case "CRA", "CFO", "RFC": blnOk = False                             ' ni MAD ni ECH
            Case "CHH", "CPB": If optSelect_Avis_Echéance Then blnOk = False    ' pas de ECH
        End Select
        
        If blnOk Then prtSAB_CRE_ZCREBIS0 fgSelect_ZCREBIS0(wIndex), meYBIACRE, optSelect_Confirmation
    Next I
    SSTab1.Tab = 0
    cmdSelect_Ok_Click
    Me.Show
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub



Public Sub fgSelect_DisplayLine()
On Error Resume Next
End Sub


Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 6
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = 0

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
    fgSelect.Col = lK
    X = Format$(Val(fgSelect.Text), "000000000000000.00")
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), SAB_CRE_Aut)

'blnSetfocus = True
Form_Init


End Sub


Public Sub Form_Init()
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistant", vbCritical, "frmZCREPRE0.param_init"
    Unload Me
Else
    lstErr.Clear
End If

blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgDossier_FormatString = fgDossier.FormatString
fgDossier.Enabled = True
Call DTPicker_Set(txtAmjMax, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtAmjMin, YBIATAB0_DATE_CPT_J)


cmdReset
Me.Enabled = True
Me.MousePointer = 0
End Sub
Public Sub txtAMJ_Control()
Call DTPicker_Control(txtAmjMax, xAmjMax)
IbmAmjMax = dateIBM(xAmjMax)
Call DTPicker_Control(txtAmjMin, xAmjMin)
IbmAmjMin = dateIBM(xAmjMin)
End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
currentAction = ""

SSTab1.Tab = 0

rsZCREPRE0_Init meZCREPRE0
xZCREPRE0 = meZCREPRE0
ReDim meYBIACRE.ZCREPRE0(1): meYBIACRE.ZCREPRE0(1) = meZCREPRE0


fraDossier.Enabled = False
lstDossier_Contact.ForeColor = vbMagenta

fraDossier_Info.Enabled = False   ' La frame n'est que affichage d'informations
cmdPrint.Enabled = SAB_CRE_Aut.Saisir

fraSelect_Options.Visible = False
fraSelect_Options.Top = 1560
fraSelect_Options.Left = 5600

cmdSelect_Ok_Click
blnControl = True

End Sub


Public Function param_Init()
Dim K As Integer, K1 As Integer, X As String
Dim wText As String
Dim V
Dim xName As String, xMemo As String
Dim iContact As Integer
param_Init = Null

Call lst_LoadK2("DAFI", "Contact", lstDossier_Contact, False)
Call rsElpTable_Read("DAFI", "Contact", usrName_UCase, xName, xMemo)
Call lst_Scan(Trim(xName), lstDossier_Contact)


'If lstDossier_Contact.ListCount > 0 Then lstDossier_Contact.ListIndex = iContact


End Function






Public Sub fgDossier_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgDossier.Row

If lRow > 0 And lRow < fgDossier.Rows Then
    fgDossier.Row = lRow
    For I = 0 To fgDossier_arrIndex
        fgDossier.Col = I: fgDossier.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgDossier.Row = mRow
    If fgDossier.Row > 0 Then
        lRow = fgDossier.Row
        lColor_Old = fgDossier.CellBackColor
        For I = 0 To fgDossier_arrIndex
          fgDossier.Col = I: fgDossier.CellBackColor = lColor
        Next I
        fgDossier.Col = 0
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

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Dim prtI As Integer
Dim I As Long, iLen As Long


If fgSelect.Rows > 1 Then
        Select Case SSTab1.Tab
            Case 0:
                If optSelect_Avis Then
                    Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
                End If
            Case 1: cmdPrint_lstDossier_ZCREAVI0
                    cmdPrint_lstDossier_ZCREBIS0
            Case Else:
                    MsgBox "Impression non gérée pour cet onglet", vbCritical, "frmSAB_CRE.cmdPrint"
                
        End Select
End If

End Sub

Private Sub cmdSelect(lK1 As Long, lK2 As Long)
Dim wId As String, wId2 As String
Dim V
Dim X As String

On Error GoTo Error_Handler
currentAction = "CmdSelect"

fgDossier_Display

currentAction = "Recherche Dossier "
meYBIACRE.mCREDOS = lK1
meYBIACRE.mCREPRE = lK2
V = srvYBIACRE_GetBuffer(meYBIACRE)
If Not IsNull(V) Then GoTo Error_MsgBox
fgDossier_Display_YBIACRE

lstDossier_Load



Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction & " : " & lK1
End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean

blnOk = fraSelect_Options.Visible
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAb_Stock_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
lstDossier_ZCREAVI0.Clear
lstDossier_ZCREBIS0.Clear
If blnOk Then
    cmdSelect_Ok.Caption = "Options"
    cmdSelect_Ok.BackColor = &HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Visible = False
    cmsSelect_SQL
Else
    cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Visible = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAb_Stock_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub fgDossier_Click()
fgDossier.LeftCol = 0

End Sub

Private Sub fgDossier_LeaveCell()
On Error Resume Next
fgDossier.CellBackColor = &HE0E0E0
End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wK1 As Long, wK2 As Long
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_SortX fgSelect_Sort1
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
       ' fgSelect.Col = fgSelect_arrIndex: wK1 = fgSelect.Text
        fgSelect.Col = 0: wK1 = CLng(Val(fgSelect.Text))
        fgSelect.Col = 1: wK2 = CLng(Val(fgSelect.Text))
        cmdSelect wK1, wK2
   End If
End If
End Sub


Private Sub mnuContext_Auto_Click()
chkSelect_Print = "1"
optSelect_Avis_Echéance = True
fraSelect_Options.Visible = True
cmdSelect_Ok_Click
fraSelect_Options.Visible = True
optSelect_MAD = True
cmdSelect_Ok_Click

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
blnControl = False
lstErr.Clear: lstErr.Height = 200

If fraDossier.Enabled Then
    fraDossier.Enabled = False
    SSTab1.Tab = 0
    fgSelect.LeftCol = 0

    Exit Sub
End If


If SSTab1.Tab = 0 Then
        Unload Me
    Exit Sub
Else
    SSTab1.Tab = SSTab1.Tab - 1
End If

End Sub

Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
    If optSelect_Dossier Then
        fgSelect.Row = fgSelect.TopRow
        fgSelect.Col = fgSelect_arrIndex: ' wK1 = fgSelect.Text
        'cmdSelect txtSelect ''fgSelect.Text
    
       ' cmdSelect_Click
    End If
Else
    SendKeys "{TAB}"
End If
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
On Error GoTo Error_Handler

mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False
fgDossier.Clear: fgDossier.Row = 0

Exit Sub

Error_Handler:
blnControl = False

End Sub





Private Sub fgDossier_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wOrigine As String
Dim I As Integer
On Error Resume Next
If y <= fgDossier.RowHeightMin Then
Else
    If fgDossier.Rows > 1 Then
        Call fgDossier_Color(fgDossier_RowClick, MouseMoveUsr.BackColor, fgDossier_ColorClick)
        fgDossier.Col = 0: wOrigine = Trim(fgDossier.Text)
                            
            fgDossier.Col = 1
            I = Val(fgDossier.Text)
            Select Case wOrigine
                ' Case constZCREDOS0: xZCREDOS0 = meYBIACRE.ZCREDOS0(I): srvZCREDOS0_ElpDisplay xZCREDOS0
                ' Case constZCREPRE0: xZCREPRE0 = meYBIACRE.ZCREPRE0(I): srvZCREPRE0_ElpDisplay xZCREPRE0
                ' Case constZCREPLA0: xZCREPLA0 = meYBIACRE.ZCREPLA0(I): srvZCREPLA0_ElpDisplay xZCREPLA0
                ' Case constZCREEMP0: xZCREEMP0 = meYBIACRE.ZCREEMP0(I): srvZCREEMP0_ElpDisplay xZCREEMP0
                ' Case constZCREEVE0: xZCREEVE0 = meYBIACRE.ZCREEVE0(I): srvZCREEVE0_ElpDisplay xZCREEVE0
                ' Case constZCREAVI0: xZCREAVI0 = meYBIACRE.ZCREAVI0(I): srvZCREAVI0_ElpDisplay xZCREAVI0
                ' Case constZCREBIS0: xZCREBIS0 = meYBIACRE.ZCREBIS0(I): srvZCREBIS0_ElpDisplay xZCREBIS0
            End Select
   End If
End If
End Sub

Public Sub fgDossier_Reset()
fgDossier.Clear
fgDossier_Sort1 = 0: fgDossier_Sort2 = 0
fgDossier_Sort1_Old = -1
fgDossier_RowDisplay = 0: fgDossier_RowClick = 0
fgDossier_arrIndex = 3
blnfgDossier_DisplayLine = False
fgDossier_SortAD = 6
fgDossier.LeftCol = 0

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
'Call txt_LostFocus(txt)
'Call txt_GotFocus(txt)
'Call txt_LostFocus(txt)

End Sub





Private Sub mnuPrint0_Avis_Click()
Dim K As Integer, prtI As Integer
For K = 1 To fgSelect_ZCREAVI0_Nb
    For prtI = 1 To meYBIACRE.prtNb
            prtSAB_CRE_ZCREAVI0 fgSelect_ZCREAVI0(K), meYBIACRE
    Next prtI
Next K

End Sub

Private Sub optSelect_Avis_Click()
Call DTPicker_Set(txtAmjMax, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtAmjMin, YBIATAB0_DATE_CPT_J)

End Sub

Private Sub optSelect_Confirmation_Click()
xAmjMin = dateElp("Jour", 1, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtAmjMin, xAmjMin)
Call DTPicker_Set(txtAmjMax, YBIATAB0_DATE_CPT_JS1)

End Sub

Private Sub optSelect_MAD_Click()
Call DTPicker_Set(txtAmjMin, dateElp("Jour", 1, YBIATAB0_DATE_CPT_JP1))
Call DTPicker_Set(txtAmjMax, YBIATAB0_DATE_CPT_J)

End Sub

Private Sub optSelect_Avis_Echéance_Click()
xAmjMin = dateElp("Jour", 15, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtAmjMin, xAmjMin)
xAmjMax = dateElp("Jour", 14, YBIATAB0_DATE_CPT_JS1)
Call DTPicker_Set(txtAmjMax, xAmjMax)

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then txtSelect.SetFocus

End Sub

Private Sub SSTab1_GotFocus()
Select Case SSTab1.Tab
    Case 0: fgSelect.LeftCol = 0
End Select
End Sub

Private Sub txtSelect_Change()
Dim I As Long, X As String, lenX As Integer
On Error Resume Next
X = Trim(txtSelect)
lenX = Len(X)
fgSelect.Col = 0
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    
    If X <= Mid$(fgSelect.Text, 1, lenX) Then
        fgSelect.LeftCol = 0
        fgSelect.TopRow = I
        Exit Sub
    End If
Next I

End Sub

Private Sub txtSelect_GotFocus()
Call txt_GotFocus(txtSelect)

End Sub


Private Sub txtSelect_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)
End Sub


Private Sub txtSelect_LostFocus()
Call txt_LostFocus(txtSelect)

End Sub

Public Sub lstDossier_Load()
Dim I As Integer, K As Integer
Dim xTYP As String, xTYP_Lib As String
Dim blnOk As Boolean

lstDossier_ZCREAVI0.Clear
lstDossier_ZCREBIS0.Clear

For I = 1 To meYBIACRE.ZCREAVI0_Nb
    xTYP = meYBIACRE.ZCREAVI0(I).CREAVITYP
    xTYP_Lib = srvCREEVETYP_Lib(xTYP)

    lstDossier_ZCREAVI0.AddItem dateIBM10(meYBIACRE.ZCREAVI0(I).CREAVIDTC, True) & " : " & xTYP & " : " & xTYP_Lib & meYBIACRE.ZCREAVI0(I).CREAVIDOS & "_" & meYBIACRE.ZCREAVI0(I).CREAVIPRE & "_" & I
Next I

ReDim lstDossier_ZCREBIS_Index(meYBIACRE.ZCREBIS0_Nb + 1)

For I = 1 To meYBIACRE.ZCREBIS0_Nb
    blnOk = True
    'If optSelect_Dossier Then blnOk = True
    'If meYBIACRE.ZCREBIS0(I).CREBISAVI <= 1 Then blnOk = True ''> 0 Then
    If blnOk Then
        lstDossier_ZCREBIS_Index(lstDossier_ZCREBIS0.ListCount) = I
        xTYP = meYBIACRE.ZCREBIS0(I).CREBISTYP
        xTYP_Lib = srvCREEVETYP_Lib(xTYP)
        lstDossier_ZCREBIS0.AddItem dateIBM10(meYBIACRE.ZCREBIS0(I).CREBISEMI, True) & " : " & xTYP & " : " & xTYP_Lib & meYBIACRE.ZCREBIS0(I).CREBISDOS & "_" & meYBIACRE.ZCREBIS0(I).CREBISPRE & "_" & I
   End If
Next I


End Sub









Public Sub fgDossier_Display_YBIACRE()
Dim X As String
Dim I As Integer


For I = 1 To meYBIACRE.ZCREDOS0_Nb
    X = CStr(meYBIACRE.ZCREDOS0(I).CREDOSDOS)
    Call fgDossier_DisplayLine("ZCREDOS0", CStr(I), X)
Next I


For I = 1 To meYBIACRE.ZCREEMP0_Nb
    X = CStr(meYBIACRE.ZCREEMP0(I).CREEMPDOS) & "_" & CStr(meYBIACRE.ZCREEMP0(I).CREEMPSEQ) & "_" & CStr(meYBIACRE.ZCREEMP0(I).CREEMPNCL)
    Call fgDossier_DisplayLine("ZCREEMP0", CStr(I), X)
Next I

For I = 1 To meYBIACRE.ZCREPRE0_Nb
    X = CStr(meYBIACRE.ZCREPRE0(I).CREPREDOS) & "_" & CStr(meYBIACRE.ZCREPRE0(I).CREPREPRE)
    Call fgDossier_DisplayLine("ZCREPRE0", CStr(I), X)
Next I


For I = 1 To meYBIACRE.ZCREPLA0_Nb
    X = CStr(meYBIACRE.ZCREPLA0(I).CREPLADOS) & "_" & CStr(meYBIACRE.ZCREPLA0(I).CREPLAPRE) & "_" & CStr(meYBIACRE.ZCREPLA0(I).CREPLAPLA)
    Call fgDossier_DisplayLine("ZCREPLA0", CStr(I), X)
Next I

For I = 1 To meYBIACRE.ZCREEVE0_Nb
    X = CStr(meYBIACRE.ZCREEVE0(I).CREEVEDOS) & "_" & CStr(meYBIACRE.ZCREEVE0(I).CREEVEPRE) & "_" & CStr(meYBIACRE.ZCREEVE0(I).CREEVEPLA) & " : " & CStr(meYBIACRE.ZCREEVE0(I).CREEVETYP)
    Call fgDossier_DisplayLine("ZCREEVE0", CStr(I), X)
Next I

For I = 1 To meYBIACRE.ZCREAVI0_Nb
    X = CStr(meYBIACRE.ZCREAVI0(I).CREAVIDOS) & "_" & CStr(meYBIACRE.ZCREAVI0(I).CREAVIPRE) & "_" & CStr(meYBIACRE.ZCREAVI0(I).CREAVITYP)
    Call fgDossier_DisplayLine("ZCREAVI0", CStr(I), X)
Next I

For I = 1 To meYBIACRE.ZCREBIS0_Nb
    X = CStr(meYBIACRE.ZCREBIS0(I).CREBISDOS) & "_" & CStr(meYBIACRE.ZCREBIS0(I).CREBISPRE) & "_" & dateIBM10(CStr(meYBIACRE.ZCREBIS0(I).CREBISDTR), True)
    Call fgDossier_DisplayLine("ZCREBIS0", CStr(I), X)
Next I

lstDossier_Adresse.Clear
lstDossier_Adresse.AddItem "Emprunteur : " & meYBIACRE.mCREEMPNCL
lstDossier_Adresse.AddItem ""
lstDossier_Adresse.AddItem Trim(meYBIACRE.CRE_ZADRESS0.ADRESSRA1) & Trim(meYBIACRE.CRE_ZADRESS0.ADRESSRA2)
lstDossier_Adresse.AddItem Trim(meYBIACRE.CRE_ZADRESS0.ADRESSAD1)
lstDossier_Adresse.AddItem Trim(meYBIACRE.CRE_ZADRESS0.ADRESSAD2)
lstDossier_Adresse.AddItem Trim(meYBIACRE.CRE_ZADRESS0.ADRESSAD3)
lstDossier_Adresse.AddItem Trim(meYBIACRE.CRE_ZADRESS0.ADRESSCOP) & " " & Trim(meYBIACRE.CRE_ZADRESS0.ADRESSVIL)
lstDossier_Adresse.AddItem Trim(meYBIACRE.CRE_ZADRESS0.ADRESSPAY)
End Sub

Public Sub cmsSelect_SQL()
On Error Resume Next
meYBIACRE.prtNb = 2
If chkDossier_prtNb = "0" Then meYBIACRE.prtNb = 1
meYBIACRE.Contact = lstDossier_Contact

mnuPrint0_Avis.Enabled = optSelect_Avis
If optSelect_Dossier Then fgSelect_Display: txtSelect.SetFocus
If optSelect_MAD Then fgSelect_ZCREBIS0_SQL
If optSelect_Confirmation Then fgSelect_ZCREBIS0_SQL
If optSelect_Avis_Echéance Then fgSelect_ZCREBIS0_SQL
If optSelect_Avis Then fgSelect_Avis

End Sub

Public Sub cmdPrint_lstDossier_ZCREAVI0()
Dim K As Integer, prtI As Integer
For K = 0 To lstDossier_ZCREAVI0.ListCount - 1
    'lstDossier_ZCREAVI0.ListIndex = K
    If lstDossier_ZCREAVI0.Selected(K) Then
        For prtI = 1 To meYBIACRE.prtNb
                prtSAB_CRE_ZCREAVI0 meYBIACRE.ZCREAVI0(K + 1), meYBIACRE
        Next prtI
    End If
Next K
End Sub
Public Sub cmdPrint_lstDossier_ZCREBIS0()
Dim K As Integer, prtI As Integer, K2 As Integer
For K = 0 To lstDossier_ZCREBIS0.ListCount - 1
    'lstDossier_ZCREBIS0.ListIndex = K
    If lstDossier_ZCREBIS0.Selected(K) Then
        For prtI = 1 To meYBIACRE.prtNb
                K2 = lstDossier_ZCREBIS_Index(K)
                prtSAB_CRE_ZCREBIS0 meYBIACRE.ZCREBIS0(K2), meYBIACRE, optSelect_Confirmation
        Next prtI
    End If
Next K
End Sub


