VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDAUTLIB0 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA_DWH : DAUTLIB0"
   ClientHeight    =   9144
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   13560
   Icon            =   "DAUTLIB0.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9144
   ScaleWidth      =   13560
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Export .xls"
      Height          =   405
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   0
      Width           =   972
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   7800
      TabIndex        =   2
      Top             =   0
      Width           =   4815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   13530
      _ExtentX        =   23855
      _ExtentY        =   15261
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "maintenance du fichier DAUTLIB0"
      TabPicture(0)   =   "DAUTLIB0.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "....."
      TabPicture(1)   =   "DAUTLIB0.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   13290
         Begin VB.CommandButton cmdSelect_Ajouter 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Ajouter"
            Enabled         =   0   'False
            Height          =   405
            Left            =   11520
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   720
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Frame fraDAUTLIB0 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Gestion du fichier - DAUTLIB0 -"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   6135
            Left            =   6720
            TabIndex        =   8
            Top             =   1560
            Visible         =   0   'False
            Width           =   6015
            Begin VB.ComboBox cboDAUTLIBRGP 
               Height          =   315
               Left            =   1440
               Sorted          =   -1  'True
               TabIndex        =   12
               Text            =   "cboDAUTLIBRGP"
               Top             =   2640
               Width           =   4335
            End
            Begin VB.CommandButton cmdfra_Quit 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Abandonner"
               Enabled         =   0   'False
               Height          =   732
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   5280
               Width           =   1095
            End
            Begin VB.ComboBox cboDAUTLIBAMO 
               Height          =   288
               Left            =   1440
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   4560
               Width           =   1095
            End
            Begin VB.TextBox txtDAUTLIBTXT 
               Height          =   1260
               Left            =   1440
               MaxLength       =   64
               MultiLine       =   -1  'True
               TabIndex        =   11
               Top             =   1080
               Width           =   4335
            End
            Begin VB.ComboBox cboDAUTLIBELM 
               Height          =   315
               Left            =   1440
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   4080
               Width           =   1095
            End
            Begin VB.CommandButton cmdfra_Creer 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Créer"
               Height          =   735
               Left            =   2160
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   5280
               Width           =   1095
            End
            Begin VB.CommandButton cmdfra_Modifier 
               BackColor       =   &H00FF80FF&
               Caption         =   "Modifier"
               Height          =   735
               Left            =   4560
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   5280
               Width           =   1095
            End
            Begin VB.CommandButton cmdfra_Supprimer 
               BackColor       =   &H000000FF&
               Caption         =   "Supprimer"
               Height          =   735
               Left            =   3360
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   5280
               Width           =   1095
            End
            Begin VB.TextBox txtDAUTLIBCOD 
               Height          =   300
               Left            =   1440
               MaxLength       =   20
               TabIndex        =   10
               Top             =   600
               Width           =   4335
            End
            Begin VB.Label lblDAUTLIBAMO 
               BackColor       =   &H00C0E0FF&
               Caption         =   "AMO"
               Height          =   375
               Left            =   240
               TabIndex        =   22
               Top             =   4680
               Width           =   1095
            End
            Begin VB.Label lblDAUTLIBRGP 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Regroupement"
               Height          =   375
               Left            =   240
               TabIndex        =   21
               Top             =   2760
               Width           =   1095
            End
            Begin VB.Label lblDAUTLIBELM 
               BackColor       =   &H00C0E0FF&
               Caption         =   "ELM"
               Height          =   372
               Left            =   240
               TabIndex        =   20
               Top             =   4080
               Width           =   1092
            End
            Begin VB.Label lblDAUTLIBTXT 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Libellé"
               Height          =   300
               Left            =   240
               TabIndex        =   15
               Top             =   1080
               Width           =   1815
            End
            Begin VB.Label lblDAUTLIBCOD 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Code"
               Height          =   300
               Left            =   240
               TabIndex        =   9
               Top             =   600
               Width           =   1815
            End
         End
         Begin VB.Frame fraSelect_Options 
            Height          =   1005
            Left            =   240
            TabIndex        =   6
            Top             =   120
            Width           =   10995
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Enabled         =   0   'False
            Height          =   435
            Left            =   11520
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6825
            Left            =   120
            TabIndex        =   7
            Top             =   1200
            Width           =   12840
            _ExtentX        =   22648
            _ExtentY        =   12044
            _Version        =   393216
            Rows            =   1
            Cols            =   13
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   16777210
            ForeColor       =   8388608
            BackColorFixed  =   16776921
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   16777210
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   "Code                   |<Libellé                                              |<Regroupement                    |<ELM |<AMO ||"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
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
   Begin VB.Label libSelect 
      BackColor       =   &H00FFFED9&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4905
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
Attribute VB_Name = "frmDAUTLIB0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit

Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim BIA_DWH_Aut As typeAuthorization
Dim curX1 As Currency, curX2 As Currency
Dim blnAuto As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim cnAdo As New ADODB.Connection, rsADO As New ADODB.Recordset, errADO As ADODB.Error
Dim blnTransaction As Boolean

'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
Dim xDAUTLIB0 As typeDAUTLIB0, newDAUTLIB0 As typeDAUTLIB0, oldDAUTLIB0 As typeDAUTLIB0
Dim arrDAUTLIB0() As typeDAUTLIB0, arrDAUTLIB0_Nb As Long, arrDAUTLIB0_Max As Long, arrDAUTLIB0_Index As Long
Dim xDRTAGRP As typeDRTAGRP
Dim xWhere_CRTA As String, xWhere_CGRP As String

'______________________________________________________________________

Dim fraDAUTLIB0_FormatString As String, fraDAUTLIB0_K As Integer
Dim fraDAUTLIB0_RowDisplay As Integer, fraDAUTLIB0_RowClick As Integer, fraDAUTLIB0_ColClick As Integer
Dim fraDAUTLIB0_ColorClick As Long, fraDAUTLIB0_ColorDisplay As Long
Dim fraDAUTLIB0_Sort1 As Integer, fraDAUTLIB0_Sort2 As Integer
Dim fraDAUTLIB0_SortAD As Integer, fraDAUTLIB0_Sort1_Old As Integer
Dim fraDAUTLIB0_arrIndex As Integer
Dim blnfraDAUTLIB0_DisplayLine As Boolean

Dim meDAUTLIB0_Status As typeDAUTLIB0, oldDAUTLIB0_Status As typeDAUTLIB0

Dim rsADO_Update As New ADODB.Recordset
Dim mTitleText As String

Dim fgSelect_TopRow_Memo As Integer










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
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgselect_Display"
    
For I = 1 To arrDAUTLIB0_Nb
         
    xDAUTLIB0 = arrDAUTLIB0(I)
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
Next I

fgSelect.Visible = True
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True

Call lstErr_AddItem(lstErr, cmdContext, "Nb d'enregistrement : " & arrDAUTLIB0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
    
End Sub

Private Sub arrDAUTLIB0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrDAUTLIB0(101)
arrDAUTLIB0_Max = 100: arrDAUTLIB0_Nb = 0

Set rsADO = Nothing

xSql = "select * from " & paramIBM_Library_BODWH & ".DAUTLIB0 " & xWhere & " order by DAUTLIBCOD"
Set rsADO = cnAdo.Execute(xSql)

Do While Not rsADO.EOF
    V = srvDAUTLIB0_GetBuffer_ODBC(rsADO, xDAUTLIB0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDAUTLIB0.fgselect_Display"
        '' Exit Sub
     Else
         arrDAUTLIB0_Nb = arrDAUTLIB0_Nb + 1
         If arrDAUTLIB0_Nb > arrDAUTLIB0_Max Then
             arrDAUTLIB0_Max = arrDAUTLIB0_Max + 50
             ReDim Preserve arrDAUTLIB0(arrDAUTLIB0_Max)
         End If
         
         arrDAUTLIB0(arrDAUTLIB0_Nb) = xDAUTLIB0
    End If
    rsADO.MoveNext
Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)

On Error Resume Next

fgSelect.Col = 0: fgSelect.Text = xDAUTLIB0.DAUTLIBCOD
fgSelect.Col = 1: fgSelect.Text = xDAUTLIB0.DAUTLIBTXT
fgSelect.Col = 2: fgSelect.Text = xDAUTLIB0.DAUTLIBRGP
fgSelect.Col = 3: fgSelect.Text = xDAUTLIB0.DAUTLIBELM
fgSelect.Col = 4: fgSelect.Text = xDAUTLIB0.DAUTLIBAMO


fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

End Sub


Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = fgSelect.Cols - 1
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
fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub

Public Function cmdDAUTLIB0_Control()
Dim xControl As String
cmdDAUTLIB0_Control = Null
xControl = ""
srvDAUTLIB0_Init newDAUTLIB0

newDAUTLIB0.DAUTLIBCOD = Trim(txtDAUTLIBCOD)
newDAUTLIB0.DAUTLIBTXT = Trim(txtDAUTLIBTXT)
newDAUTLIB0.DAUTLIBRGP = Trim(cboDAUTLIBRGP)
newDAUTLIB0.DAUTLIBELM = Trim(cboDAUTLIBELM)
newDAUTLIB0.DAUTLIBAMO = Trim(cboDAUTLIBAMO)
If Trim(newDAUTLIB0.DAUTLIBCOD) = "" Then xControl = xControl & "- Précisez le code" & vbCrLf
If Trim(newDAUTLIB0.DAUTLIBTXT) = "" Then xControl = xControl & "- Précisez le libellé" & vbCrLf
If Trim(newDAUTLIB0.DAUTLIBRGP) = "" Then xControl = xControl & "- Précisez le regroupement" & vbCrLf
If Trim(newDAUTLIB0.DAUTLIBELM) <> "Oui" And Trim(newDAUTLIB0.DAUTLIBELM) <> "Non" Then xControl = xControl & "- ELM différent de 'Oui' ou 'Non'" & vbCrLf
If Trim(newDAUTLIB0.DAUTLIBAMO) <> "Oui" And Trim(newDAUTLIB0.DAUTLIBAMO) <> "Non" Then xControl = xControl & "- AMO différent de 'Oui' ou 'Non'" & vbCrLf

If xControl <> "" Then
    MsgBox xControl, vbCritical, "Liste des anomalies détectées :"
    cmdDAUTLIB0_Control = xControl
End If
End Function


Public Sub cmdDAUTLIB0_Update()

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

V = sqlDAUTLIB0_Update(newDAUTLIB0, xDAUTLIB0, cnAdo)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

If Not IsNull(V) Then MsgBox V, vbCritical

End Sub

Public Sub cmdDAUTLIB0_Delete()

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

V = sqlDAUTLIB0_Delete(xDAUTLIB0, cnAdo)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

If Not IsNull(V) Then MsgBox V, vbCritical

End Sub

Public Sub cmdDAUTLIB0_Insert()
Dim V
Dim xSql As String, Nb As Integer
Dim wSWICLAOPR As String, wSWICLANUM As Long
xDAUTLIB0 = newDAUTLIB0
V = sqlDAUTLIB0_Read(xDAUTLIB0, cnAdo)
If IsNull(V) Then
    MsgBox "Enregistrement existant dans DAUTLIB0 !" & vbCrLf & xDAUTLIB0.DAUTLIBTXT, vbQuestion, xDAUTLIB0.DAUTLIBCOD
    Exit Sub
Else
    If V <> "? inconnu" Then
        MsgBox "Erreur de lecture du fichier DAUTLIB0", vbCritical, xDAUTLIB0.DAUTLIBCOD
        Exit Sub
    End If
End If


'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

' wAmjMax = Période de saisie
V = sqlDAUTLIB0_Insert(newDAUTLIB0, cnAdo)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

        
If Not IsNull(V) Then

    MsgBox V, vbCritical
    cmdfra_Creer.Visible = True
    cmdfra_Modifier.Visible = False
    cmdfra_Supprimer.Visible = False
    txtDAUTLIBCOD.Enabled = False
    cboDAUTLIBELM.Enabled = True
    
    fraDAUTLIB0.Visible = True
    cmdSelect_Ok.Enabled = False
    cmdSelect_Ajouter.Enabled = False

Else
   
    ' Charger newDAUTLIB0 dans arrDAUTLIB0 puis
    ' Rafraichir l'affichage de la liste dans -fgSelect-
    ' Maintenir l'affichage de la ligne ajoutée sur la page en cours d'affichage
    
    arrDAUTLIB0_Nb = arrDAUTLIB0_Nb + 1
    arrDAUTLIB0(arrDAUTLIB0_Nb) = newDAUTLIB0
    xDAUTLIB0 = newDAUTLIB0
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine arrDAUTLIB0_Nb

    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    fgSelect.TopRow = fgSelect.Row
     
    ' Maintenir l'affichage de la frame de gestion pour le mode -CREATION-
    xDAUTLIB0.DAUTLIBCOD = ""
    Call fraDAUTLIB0_Display

End If

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

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), BIA_DWH_Aut)
Form_Init

blnAuto = False

End Sub


Public Sub Form_Init()
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    If Not blnAuto Then MsgBox "paramétrage inconsistant", vbCritical, "frmBIA_DWH.param_init"
    Unload Me
Else
    lstErr.Clear
End If


blnControl = False
fgSelect_FormatString = fgSelect.FormatString

fgSelect.Enabled = True
cmdReset

Me.Enabled = True
Me.MousePointer = 0
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
fraDAUTLIB0.Visible = False
fraSelect_Options.Enabled = True
cmdSelect_Ok.Enabled = True

cmdSelect_Ok_Click
'cmdSelect_SQL
'cmdSelect_Ajouter.Enabled = True

blnControl = True

End Sub


Public Function param_Init()
Dim xSql As String
param_Init = Null
Call lstErr_Clear(lstErr, cmdContext, "Param_Init"): DoEvents

fgSelect.Visible = False

cboDAUTLIBELM.Clear
cboDAUTLIBELM.AddItem "Oui"
cboDAUTLIBELM.AddItem "Non"

cboDAUTLIBAMO.Clear
cboDAUTLIBAMO.AddItem "Oui"
cboDAUTLIBAMO.AddItem "Non"

cboDAUTLIBRGP.Clear
xSql = "select distinct DAUTLIBRGP from " & paramIBM_Library_BODWH & ".DAUTLIB0 order by DAUTLIBRGP"
Set rsADO = cnAdo.Execute(xSql)

Do While Not rsADO.EOF
    cboDAUTLIBRGP.AddItem Trim(rsADO("DAUTLIBRGP"))
    rsADO.MoveNext
Loop

Me.Enabled = True: Me.MousePointer = 0

End Function





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

Private Sub cmdfra_Creer_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
If IsNull(cmdDAUTLIB0_Control) Then cmdDAUTLIB0_Insert
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdfra_Modifier_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
If IsNull(cmdDAUTLIB0_Control) Then
    cmdDAUTLIB0_Update
    fraDAUTLIB0.Visible = False
    cmdSelect_Ok.Enabled = True
    cmdSelect_Ajouter.Enabled = True
    
    ' Rafraichir la liste d'affichage -fgSelect-
    arrDAUTLIB0(arrDAUTLIB0_Index) = newDAUTLIB0
    xDAUTLIB0 = newDAUTLIB0
    fgSelect_DisplayLine arrDAUTLIB0_Index
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdfra_Quit_Click()
cmdContext_Quit
End Sub

Private Sub cmdfra_Supprimer_Click()
Dim X As String
X = MsgBox("Confirmer la suppression ?", vbQuestion + vbOKCancel, "BIA_DWH : DAUTLIB0")

If X = vbOK Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    'cmdDAUTLIB0_Control
    cmdDAUTLIB0_Delete
    fraDAUTLIB0.Visible = False
    cmdSelect_Ok.Enabled = True
    cmdSelect_Ajouter.Enabled = True
    Me.Enabled = True: Me.MousePointer = 0
    
    fgSelect_TopRow_Memo = fgSelect.TopRow
    ' Rafraichir la liste sur -fgSelect- et positionner à la page en cours d'affichage
    cmdSelect_SQL
    If fgSelect.Rows > 1 Then fgSelect.TopRow = fgSelect_TopRow_Memo
Else
    fraDAUTLIB0.Visible = True
    cmdSelect_Ok.Enabled = False
    cmdSelect_Ajouter.Enabled = False
End If

End Sub


Private Sub cmdPrint_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdDAUTLIB0_Export
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_SQL()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim wAmj7 As Long
On Error GoTo Error_Handler

currentAction = "cmdDAUTLIB0_SQL"
  

arrDAUTLIB0_SQL ""
fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_Ajouter_Click()

Dim blnOk As Boolean, Nb As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_DWH_cmdSelect_Ajouter........"): DoEvents

xDAUTLIB0.DAUTLIBCOD = ""
Call fraDAUTLIB0_Display

Me.Enabled = True: Me.MousePointer = 0

End Sub



Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_DWH_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
'fraDAUTLIB0.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True

If blnOk Then
    cmdSelect_Ok.Caption = "Rechercher"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = False
    cmdSelect_Ajouter.Visible = True
    cmdSelect_SQL
Else
    cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = True
    cmdSelect_Ajouter.Visible = False
End If
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_DWH_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim K As Long
On Error Resume Next
If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 6: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  arrDAUTLIB0_Index = CLng(fgSelect.Text)
        fgSelect.LeftCol = 0
        
        oldDAUTLIB0 = arrDAUTLIB0(arrDAUTLIB0_Index)
        If IsNull(sqlDAUTLIB0_Read(oldDAUTLIB0, cnAdo)) Then
            xDAUTLIB0 = oldDAUTLIB0
            Call fraDAUTLIB0_Display
        Else
            MsgBox "Erreur de lecture du fichier DAUTLIB0", vbCritical, "frmDAUTLIB0"
        End If
   End If
End If
fgSelect.LeftCol = 0
End Sub

Public Sub fraDAUTLIB0_DisplayLine()
Dim X As String

On Error Resume Next

 txtDAUTLIBCOD = Trim(xDAUTLIB0.DAUTLIBCOD)
 txtDAUTLIBTXT = Trim(xDAUTLIB0.DAUTLIBTXT)
 cboDAUTLIBRGP = Trim(xDAUTLIB0.DAUTLIBRGP)
 
 Call cbo_Scan(xDAUTLIB0.DAUTLIBELM, cboDAUTLIBELM)
 Call cbo_Scan(xDAUTLIB0.DAUTLIBAMO, cboDAUTLIBAMO)
 

End Sub

Private Sub fraDAUTLIB0_Display()
Dim V
Dim X As String, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim I As Long

On Error GoTo Error_Handler

fraDAUTLIB0.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True
cmdfra_Quit.Enabled = True
cmdfra_Creer.Visible = False
cmdfra_Modifier.Visible = False
cmdfra_Supprimer.Visible = False
txtDAUTLIBCOD.Enabled = False
txtDAUTLIBTXT.Enabled = True
cboDAUTLIBRGP.Enabled = True

cboDAUTLIBELM.Enabled = True
cboDAUTLIBAMO.Enabled = True

currentAction = "fraDAUTLIB0_Display"

If Trim(xDAUTLIB0.DAUTLIBCOD) <> "" Then   ' Autres que Création
   
    fraDAUTLIB0_DisplayLine
    
    cmdfra_Creer.Visible = False
    cmdfra_Modifier.Visible = True
    cmdfra_Supprimer.Visible = True
    txtDAUTLIBCOD.Enabled = False

Else  ' Création
   
    txtDAUTLIBTXT = ""
    txtDAUTLIBCOD = ""
    cboDAUTLIBRGP.ListIndex = 0
    cboDAUTLIBELM.ListIndex = 1
    cboDAUTLIBAMO.ListIndex = 0
    cmdfra_Creer.Visible = True
    cmdfra_Modifier.Visible = False
    cmdfra_Supprimer.Visible = False
    txtDAUTLIBCOD.Enabled = True

End If

fraDAUTLIB0.Visible = True
cmdSelect_Ok.Enabled = False
cmdSelect_Ajouter.Enabled = False

Call lstErr_AddItem(lstErr, cmdContext, "Lignes : "): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 1
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'If blnControl Then
    cnAdo.Close
    Set cnAdo = Nothing

'End If
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
If fraDAUTLIB0.Visible Then
    fraDAUTLIB0.Visible = False
    cmdSelect_Ok.Enabled = True
    cmdSelect_Ajouter.Enabled = True
    Exit Sub
Else
    
    If SSTab1.Tab = 0 Then
            Unload Me
        Exit Sub
    Else
        SSTab1.Tab = SSTab1.Tab - 1
    End If
End If
End Sub

Public Sub cmdContext_Return()
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

Set cnAdo = New ADODB.Connection
cnAdo.CursorLocation = adUseClient
cnAdo.Open paramODBC_DSN_SAB

Exit Sub

Error_Handler:

blnControl = False
If Not blnAuto Then MsgBox Error
End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

End Sub














Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then cmdSelect_Ok.SetFocus

End Sub














Private Sub SSTab1_GotFocus()
Select Case SSTab1.Tab
    Case 0: fgSelect.LeftCol = 0
End Select
End Sub






Public Sub blnTransaction_Set()
If Not blnTransaction Then
    blnTransaction = True
    Set rsADO_Update = cnAdo.Execute("SET TRANSACTION ISOLATION LEVEL READ COMMITTED")

End If

End Sub








































Private Sub txtDAUTLIBCOD_GotFocus()
txt_GotFocus txtDAUTLIBCOD

End Sub


Private Sub txtDAUTLIBCOD_LostFocus()
txt_LostFocus txtDAUTLIBCOD

End Sub


Private Sub cboDAUTLIBRGP_GotFocus()
txt_GotFocus cboDAUTLIBRGP
End Sub


Private Sub cboDAUTLIBRGP_LostFocus()
txt_LostFocus cboDAUTLIBRGP
End Sub


Private Sub txtDAUTLIBTXT_GotFocus()
txt_GotFocus txtDAUTLIBTXT
End Sub


Private Sub txtDAUTLIBTXT_LostFocus()
txt_LostFocus txtDAUTLIBTXT
End Sub



Public Sub cmdDAUTLIB0_Export()
On Error GoTo Error_Handler
Dim Nb As Long
Dim wFile As String, xSql As String

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim X As String

wFile = "C:\temp\" & DSys & "_DAUTLIB0.xlsx"
X = MsgBox("export du fichier : " & wFile & " ?", vbYesNo, "Maintenance de la table BODWH.DAUTLIB0")
If X <> vbYes Then Exit Sub
Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
Set wsExcel = wbExcel.ActiveSheet
'__________________________________________________________________________________

Nb = 1
wsExcel.Cells(Nb, 1) = "Code"
wsExcel.Cells(Nb, 2) = "Libellé"
wsExcel.Cells(Nb, 3) = "Regroupement"
wsExcel.Cells(Nb, 4) = "ELM"
wsExcel.Cells(Nb, 5) = "AMO"

xSql = "select * from " & paramIBM_Library_BODWH & ".DAUTLIB0  order by DAUTLIBCOD"
Set rsADO = cnAdo.Execute(xSql)

Do While Not rsADO.EOF
    V = srvDAUTLIB0_GetBuffer_ODBC(rsADO, xDAUTLIB0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDAUTLIB0.fgselect_Display"
        '' Exit Sub
     Else
        Nb = Nb + 1
        wsExcel.Cells(Nb, 1) = xDAUTLIB0.DAUTLIBCOD
        wsExcel.Cells(Nb, 2) = xDAUTLIB0.DAUTLIBTXT
        wsExcel.Cells(Nb, 3) = xDAUTLIB0.DAUTLIBRGP
        wsExcel.Cells(Nb, 4) = xDAUTLIB0.DAUTLIBELM
        wsExcel.Cells(Nb, 5) = xDAUTLIB0.DAUTLIBAMO
    End If
    rsADO.MoveNext
Loop
Set rsADO = Nothing

'____________________________________________________________________________________

wbExcel.SaveAs wFile

wbExcel.Close
appExcel.Quit

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing

'_____________________________
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
