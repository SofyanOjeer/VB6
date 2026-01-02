VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDRENTA 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA_DWH"
   ClientHeight    =   9144
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   13560
   Icon            =   "DRENTA.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9144
   ScaleWidth      =   13560
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   432
      Left            =   7800
      TabIndex        =   3
      Top             =   0
      Width           =   5175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   13530
      _ExtentX        =   23855
      _ExtentY        =   15261
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Gestion sur les mois traités"
      TabPicture(0)   =   "DRENTA.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "....."
      TabPicture(1)   =   "DRENTA.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13290
         Begin VB.CommandButton cmdSelect_Ajouter 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Ajouter"
            Enabled         =   0   'False
            Height          =   285
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   720
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Frame fraDRENTA 
            BackColor       =   &H80000013&
            Caption         =   "Gestion du fichier - DRENTA -"
            Height          =   5775
            Left            =   6720
            TabIndex        =   11
            Top             =   1800
            Visible         =   0   'False
            Width           =   6015
            Begin VB.CommandButton cmdfra_Renommer 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Renommer"
               Height          =   495
               Left            =   1560
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   4920
               Width           =   1095
            End
            Begin VB.TextBox txtfg_DRTATXM 
               Height          =   300
               Left            =   2160
               TabIndex        =   22
               Text            =   "Taux de marge"
               Top             =   3360
               Width           =   1695
            End
            Begin VB.TextBox txtfg_DRTAMOYB 
               Height          =   300
               Left            =   2160
               TabIndex        =   20
               Text            =   "Moyenne"
               Top             =   2640
               Width           =   1695
            End
            Begin VB.ComboBox cbo_CRTA 
               Height          =   315
               Left            =   2160
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   1680
               Width           =   2775
            End
            Begin VB.CommandButton cmdfra_Creer 
               BackColor       =   &H00FFC0FF&
               Caption         =   "Créer"
               Height          =   495
               Left            =   5040
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   4920
               Width           =   855
            End
            Begin VB.CommandButton cmdfra_Modifier 
               BackColor       =   &H00FF80FF&
               Caption         =   "Modifier"
               Height          =   495
               Left            =   3960
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   4920
               Width           =   855
            End
            Begin VB.CommandButton cmdfra_Supprimer 
               BackColor       =   &H000000FF&
               Caption         =   "Supprimer"
               Height          =   495
               Left            =   2880
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   4920
               Width           =   855
            End
            Begin VB.TextBox txtfg_DRTAMMRB 
               Height          =   300
               Left            =   2160
               TabIndex        =   21
               Text            =   "Montant marge"
               Top             =   3000
               Width           =   1695
            End
            Begin VB.TextBox txtfg_DRTACTR 
               Height          =   300
               Left            =   2160
               TabIndex        =   23
               Text            =   "Comptage"
               Top             =   3720
               Width           =   615
            End
            Begin VB.TextBox txtfg_DRTACGRP 
               Height          =   300
               Left            =   2160
               TabIndex        =   19
               Text            =   "Code GRP renta"
               Top             =   2280
               Width           =   1095
            End
            Begin VB.TextBox txtfg_DRTACLIB 
               Enabled         =   0   'False
               Height          =   300
               Left            =   3000
               TabIndex        =   17
               Text            =   "No matricule"
               Top             =   1320
               Width           =   855
            End
            Begin VB.TextBox txtfg_DRTACLIA 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2160
               TabIndex        =   16
               Text            =   "Blanc / T"
               Top             =   1320
               Width           =   615
            End
            Begin VB.TextBox txtfg_DRTAETA 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2160
               TabIndex        =   15
               Text            =   "Etablissement"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtfg_DRTAVER 
               Enabled         =   0   'False
               Height          =   300
               Left            =   3000
               TabIndex        =   14
               Text            =   "Version"
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox txtfg_DRTASTA 
               Height          =   300
               Left            =   2160
               TabIndex        =   13
               Text            =   "Statut"
               Top             =   600
               Width           =   615
            End
            Begin VB.Label lblfg_DRTATXM 
               BackColor       =   &H80000013&
               Caption         =   "Taux de marge"
               Height          =   375
               Left            =   240
               TabIndex        =   35
               Top             =   3480
               Width           =   1815
            End
            Begin VB.Label lblfg_DRTAMOYB 
               BackColor       =   &H80000013&
               Caption         =   "Moyenne devise - BASE"
               Height          =   375
               Left            =   240
               TabIndex        =   34
               Top             =   2640
               Width           =   1815
            End
            Begin VB.Label lblfg_DRTACTR 
               BackColor       =   &H80000013&
               Caption         =   "Comptage"
               Height          =   300
               Left            =   240
               TabIndex        =   29
               Top             =   3840
               Width           =   1815
            End
            Begin VB.Label lblfg_DRTAMMRB 
               BackColor       =   &H80000013&
               Caption         =   "Montant marge renta - BASE"
               Height          =   375
               Left            =   240
               TabIndex        =   28
               Top             =   3000
               Width           =   1815
            End
            Begin VB.Label lblfg_DRTACGRP 
               BackColor       =   &H80000013&
               Caption         =   "Code regroupement renta"
               Height          =   300
               Left            =   240
               TabIndex        =   27
               Top             =   2280
               Width           =   1815
            End
            Begin VB.Label lblfg_DRTACRTA 
               BackColor       =   &H80000013&
               Caption         =   "Code renta"
               Height          =   300
               Left            =   240
               TabIndex        =   26
               Top             =   1800
               Width           =   1815
            End
            Begin VB.Label lblfg_DRTACLIA_CLIB 
               BackColor       =   &H80000013&
               Caption         =   "Blanc / T / No matricule"
               Height          =   300
               Left            =   240
               TabIndex        =   25
               Top             =   1440
               Width           =   1815
            End
            Begin VB.Label lblfg_DRTAETA 
               BackColor       =   &H80000013&
               Caption         =   "Code établissement"
               Height          =   300
               Left            =   240
               TabIndex        =   24
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label lblfg_DRTASTA_VER 
               BackColor       =   &H80000013&
               Caption         =   "Statut / Version"
               Height          =   300
               Left            =   240
               TabIndex        =   12
               Top             =   600
               Width           =   1815
            End
         End
         Begin VB.Frame fraSelect_Options 
            Height          =   1005
            Left            =   240
            TabIndex        =   7
            Top             =   120
            Width           =   11355
            Begin MSComCtl2.DTPicker txtSelect_DRTAPER 
               Height          =   300
               Left            =   1920
               TabIndex        =   8
               Top             =   240
               Width           =   1455
               _ExtentX        =   2561
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   10223619
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblSelect_DRTAPER 
               Caption         =   "Période de traitement"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Enabled         =   0   'False
            Height          =   315
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6825
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   12840
            _ExtentX        =   22648
            _ExtentY        =   12044
            _Version        =   393216
            Rows            =   1
            Cols            =   14
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
            FormatString    =   $"DRENTA.frx":047A
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
      TabIndex        =   1
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13080
      Picture         =   "DRENTA.frx":0503
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
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
      TabIndex        =   4
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
Attribute VB_Name = "frmDRENTA"
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

Dim cnADO As New ADODB.Connection, rsado As New ADODB.Recordset, errADO As ADODB.Error
Dim blnTransaction As Boolean

'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
Dim xDRENTA As typeDRENTA, newDRENTA As typeDRENTA, oldDRENTA As typeDRENTA
Dim arrDRENTA() As typeDRENTA, arrDRENTA_Nb As Long, arrDRENTA_Max As Long, arrDRENTA_Index As Long
Dim xDRTAGRP As typeDRTAGRP
Dim xWhere_CRTA As String, xWhere_CGRP As String

'______________________________________________________________________

Dim fraDRENTA_FormatString As String, fraDRENTA_K As Integer
Dim fraDRENTA_RowDisplay As Integer, fraDRENTA_RowClick As Integer, fraDRENTA_ColClick As Integer
Dim fraDRENTA_ColorClick As Long, fraDRENTA_ColorDisplay As Long
Dim fraDRENTA_Sort1 As Integer, fraDRENTA_Sort2 As Integer
Dim fraDRENTA_SortAD As Integer, fraDRENTA_Sort1_Old As Integer
Dim fraDRENTA_arrIndex As Integer
Dim blnfraDRENTA_DisplayLine As Boolean

Dim meDRENTA_Status As typeDRENTA, oldDRENTA_Status As typeDRENTA

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
    
For I = 1 To arrDRENTA_Nb
         
    xDRENTA = arrDRENTA(I)
    If xDRENTA.DRTASTA <> "" Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
    End If
Next I

fgSelect.Visible = True
fraDRENTA.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True

Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrDRENTA_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
    
End Sub

Private Sub arrDRENTA_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrDRENTA(101)
arrDRENTA_Max = 100: arrDRENTA_Nb = 0

Set rsado = Nothing

xSql = "select * from " & paramIBM_Library_BODWH & ".DRENTA " & xWhere & " order by DRTACLIA, DRTACLIB, DRTACRTA"
Set rsado = cnADO.Execute(xSql)

Do While Not rsado.EOF
    V = srvDRENTA_GetBuffer_ODBC(rsado, xDRENTA)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDRENTA.fgselect_Display"
        '' Exit Sub
     Else
         arrDRENTA_Nb = arrDRENTA_Nb + 1
         If arrDRENTA_Nb > arrDRENTA_Max Then
             arrDRENTA_Max = arrDRENTA_Max + 50
             ReDim Preserve arrDRENTA(arrDRENTA_Max)
         End If
         
         arrDRENTA(arrDRENTA_Nb) = xDRENTA
    End If
    rsado.MoveNext
Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub DRTAGRP_CRTA_SQL(xWhere_CRTA As String)

Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler

cbo_CRTA.Clear

Set rsado = Nothing

xSql = "select * from " & paramIBM_Library_BODWH & ".DRTAGRP " & xWhere_CRTA & " order by DRGRCRTA"
Set rsado = cnADO.Execute(xSql)

Do While Not rsado.EOF
    V = srvDRTAGRP_GetBuffer_ODBC(rsado, xDRTAGRP)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDRENTA.DRTAGRP_CRTA_SQL"
        '' Exit Sub
     Else
         cbo_CRTA.AddItem xDRTAGRP.DRGRCRTA & "  " & xDRTAGRP.DRGRLIB
    End If
    rsado.MoveNext
Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub DRTAGRP_CGRP_SQL(xWhere_CGRP As String)
Dim V
Dim wLecture As Boolean
Dim X As String, xSql As String
On Error GoTo Error_Handler

wLecture = False
Set rsado = Nothing

xSql = "select *from " & paramIBM_Library_BODWH & ".DRTAGRP where DRGRCRTA = " & Val(cbo_CRTA.Text) & xWhere_CGRP
Set rsado = cnADO.Execute(xSql)

Do While Not rsado.EOF
    V = srvDRTAGRP_GetBuffer_ODBC(rsado, xDRTAGRP)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDRENTA.DRTAGRP_CGRP_SQL"
         Exit Sub
     Else
         txtfg_DRTACGRP = xDRTAGRP.DRGRCGRP
         wLecture = True
     End If
    rsado.MoveNext
Loop

If wLecture = False Then
    X = MsgBox("Ressaisir le code renta inconnu !", vbCritical, "BIA_DWH : DRENTA")
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Public Sub fgSelect_DisplayLine(lIndex As Long)

On Error Resume Next

fgSelect.Col = 0: fgSelect.Text = xDRENTA.DRTASTA
fgSelect.Col = 1: fgSelect.Text = xDRENTA.DRTAVER
fgSelect.Col = 2: fgSelect.Text = dateImp10(xDRENTA.DRTAPER)
fgSelect.Col = 3: fgSelect.Text = xDRENTA.DRTAETA
fgSelect.Col = 4: fgSelect.Text = xDRENTA.DRTACLIA & " " & xDRENTA.DRTACLIB

fgSelect.Col = 5: fgSelect.Text = xDRENTA.DRTACRTA
fgSelect.Col = 6: fgSelect.Text = xDRENTA.DRTACGRP
fgSelect.Col = 7
If xDRENTA.DRTAMOYB = 0 Then
    fgSelect.Text = ""
Else
    fgSelect.Text = Format$(xDRENTA.DRTAMOYB, "### ### ### ##0.00")
    ' Mettre en rouge les montants débiteurs
    If xDRENTA.DRTAMOYB >= 0 Then
        fgSelect.CellForeColor = vbBlue
    Else
        fgSelect.CellForeColor = vbRed
    End If
End If

fgSelect.Col = 8
If xDRENTA.DRTAMMRB = 0 Then
    fgSelect.Text = ""
Else
  fgSelect.Text = Format$(xDRENTA.DRTAMMRB, "### ### ### ##0.00")
    ' Mettre en rouge les montants débiteurs
    If xDRENTA.DRTAMMRB >= 0 Then
        fgSelect.CellForeColor = vbBlue
    Else
        fgSelect.CellForeColor = vbRed
    End If
End If
fgSelect.Col = 9
If xDRENTA.DRTATXM = 0 Then
    fgSelect.Text = ""
Else
    fgSelect.Text = Format$(xDRENTA.DRTATXM, "### ##0.000000")
    ' Mettre en rouge les taux débiteurs
    If xDRENTA.DRTATXM >= 0 Then
        fgSelect.CellForeColor = vbBlue
    Else
        fgSelect.CellForeColor = vbRed
    End If
End If

fgSelect.Col = 10: fgSelect.Text = xDRENTA.DRTACTR

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
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    If lK = 2 Then
        fgSelect.Col = 2
        X = fgSelect.Text
    Else
        X = ""
    End If
    
    fgSelect.Col = 3
    X = X & Format$(Val(fgSelect.Text), "000000000000000.00")
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub

Public Sub cmdDRENTA_Charger()

srvDRENTA_Init newDRENTA

newDRENTA.DRTASTA = Trim(txtfg_DRTASTA)
newDRENTA.DRTAVER = Val(txtfg_DRTAVER)

newDRENTA.DRTAPER = wAmjMax   '''xDRENTA.DRTAPER
newDRENTA.DRTAETA = Trim(txtfg_DRTAETA)

newDRENTA.DRTACLIA = Trim(txtfg_DRTACLIA)
newDRENTA.DRTACLIB = Val(Mid$(txtfg_DRTACLIB, 1, 7))
newDRENTA.DRTACRTA = Val(Mid$(cbo_CRTA.Text, 1, 5))
newDRENTA.DRTACGRP = Val(txtfg_DRTACGRP)

' Montant tel qu'il est saisi : négatif ou positif
newDRENTA.DRTAMOYB = CCur(Val(txtfg_DRTAMOYB))
newDRENTA.DRTAMMRB = CCur(Val(txtfg_DRTAMMRB))

newDRENTA.DRTATXM = CDbl(Val(txtfg_DRTATXM))
newDRENTA.DRTACTR = Val(txtfg_DRTACTR)

End Sub


Public Sub cmdDRENTA_Update()

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

V = sqlDRENTA_Update(newDRENTA, xDRENTA, cnADO)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

If Not IsNull(V) Then MsgBox V, vbCritical

End Sub

Public Sub cmdDRENTA_Delete()

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

V = sqlDRENTA_Delete(newDRENTA, xDRENTA, cnADO)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

If Not IsNull(V) Then MsgBox V, vbCritical

End Sub

Public Sub cmdDRENTA_Insert()

Dim xSql As String, Nb As Integer

xSql = "select * from " & paramIBM_Library_BODWH & ".DRENTA " _
    & "where DRTAVER = " & newDRENTA.DRTAVER _
    & " and DRTAPER = " & newDRENTA.DRTAPER _
    & " and DRTAETA = '" & newDRENTA.DRTAETA & "'" _
    & " and DRTACLIA = '" & newDRENTA.DRTACLIA & "'" _
    & " and DRTACLIB = " & newDRENTA.DRTACLIB _
    & " and DRTACRTA = " & newDRENTA.DRTACRTA

Set rsado = cnADO.Execute(xSql)
If Not rsado.EOF Then
    MsgBox "Enregistrement existant dans DRENTA !", vbCritical, "fgDRENTA.cmd_Creer_Insert"
    Exit Sub
End If

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

' wAmjMax = Période de saisie
newDRENTA.DRTAPER = wAmjMax
V = sqlDRENTA_Insert(newDRENTA, cnADO)

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
    cmdfra_Renommer.Visible = False
    txtfg_DRTAVER.Enabled = False
    txtfg_DRTAETA.Enabled = False
    txtfg_DRTACLIA.Enabled = True
    txtfg_DRTACLIB.Enabled = True
    cbo_CRTA.Enabled = True
    
    fraDRENTA.Visible = True
    cmdSelect_Ok.Enabled = False
    cmdSelect_Ajouter.Enabled = False

Else
   
    ' Charger newDRENTA dans arrDRENTA puis
    ' Rafraichir l'affichage de la liste dans -fgSelect-
    ' Maintenir l'affichage de la ligne ajoutée sur la page en cours d'affichage
    
    arrDRENTA_Nb = arrDRENTA_Nb + 1
    arrDRENTA(arrDRENTA_Nb) = newDRENTA
    xDRENTA = newDRENTA
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine arrDRENTA_Nb

    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    fgSelect.TopRow = fgSelect.Row
     
    ' Maintenir l'affichage de la frame de gestion pour le mode -CREATION-
    xDRENTA.DRTASTA = ""
    Call fraDRENTA_Display

End If

End Sub

Public Sub cmdDRENTA_Rename()

Dim xSql As String, Nb As Integer

' Vérification que création possible

xSql = "select * from " & paramIBM_Library_BODWH & ".DRENTA " _
    & "where DRTAVER = " & newDRENTA.DRTAVER _
    & " and DRTAPER = " & newDRENTA.DRTAPER _
    & " and DRTAETA = '" & newDRENTA.DRTAETA & "'" _
    & " and DRTACLIA = '" & newDRENTA.DRTACLIA & "'" _
    & " and DRTACLIB = " & newDRENTA.DRTACLIB _
    & " and DRTACRTA = " & newDRENTA.DRTACRTA

Set rsado = cnADO.Execute(xSql)
If Not rsado.EOF Then
    MsgBox "Copie imposible, enregistrement existant dans DRENTA !", vbCritical, "fgDRENTA.cmd_Rename"
    Exit Sub
End If

' Vérification que suppression cohérente

xSql = "select * from " & paramIBM_Library_BODWH & ".DRENTA " _
    & "where DRTAVER = " & oldDRENTA.DRTAVER _
    & " and DRTAPER = " & oldDRENTA.DRTAPER _
    & " and DRTAETA = '" & oldDRENTA.DRTAETA & "'" _
    & " and DRTACLIA = '" & oldDRENTA.DRTACLIA & "'" _
    & " and DRTACLIB = " & oldDRENTA.DRTACLIB _
    & " and DRTACRTA = " & oldDRENTA.DRTACRTA _
    & " and DRTAMAJ <> " & oldDRENTA.DRTAMAJ

Set rsado = cnADO.Execute(xSql)
If Not rsado.EOF Then
    MsgBox "Suppression en vue d'une copie imposible, code MAJ dans DRENTA différent !", vbCritical, "fgDRENTA.cmd_Rename"
    Exit Sub
End If

'$ TRANSACTION >>>>>>>>>> SUPPRESSION FOR RENAME >>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

V = sqlDRENTA_Delete_ForRename(oldDRENTA, cnADO)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)

'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

If Not IsNull(V) Then MsgBox V, vbCritical

If Not IsNull(V) Then
    
    MsgBox V, vbCritical
    cmdfra_Creer.Visible = False
    cmdfra_Modifier.Visible = False
    cmdfra_Supprimer.Visible = False
    cmdfra_Renommer.Visible = True
    txtfg_DRTAVER.Enabled = False
    txtfg_DRTAETA.Enabled = False
    txtfg_DRTACLIA.Enabled = True
    txtfg_DRTACLIB.Enabled = True
    cbo_CRTA.Enabled = True
    
    fraDRENTA.Visible = True
    cmdSelect_Ok.Enabled = False
    cmdSelect_Ajouter.Enabled = False
    Exit Sub
      
Else
 
    fgSelect_TopRow_Memo = fgSelect.TopRow
    ' Rafraichir la liste sur -fgSelect- et positionner à la page en cours d'affichage
    cmdSelect_SQL
    If fgSelect.Rows > 1 Then fgSelect.TopRow = fgSelect_TopRow_Memo
    
End If

'$ TRANSACTION >>>>>>>>>> CREATION FOR RENAME >>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

' wAmjMax = Période de saisie
newDRENTA.DRTAPER = wAmjMax
V = sqlDRENTA_Insert(newDRENTA, cnADO)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
If Not IsNull(V) Then

    MsgBox V, vbCritical
    cmdfra_Creer.Visible = False
    cmdfra_Modifier.Visible = False
    cmdfra_Supprimer.Visible = False
    cmdfra_Renommer.Visible = True
    txtfg_DRTAVER.Enabled = False
    txtfg_DRTAETA.Enabled = False
    txtfg_DRTACLIA.Enabled = True
    txtfg_DRTACLIB.Enabled = True
    cbo_CRTA.Enabled = True
    
    fraDRENTA.Visible = True
    cmdSelect_Ok.Enabled = False
    cmdSelect_Ajouter.Enabled = False

Else
   
    ' Charger newDRENTA dans arrDRENTA puis
    ' Rafraichir l'affichage de la liste dans -fgSelect-
    ' Maintenir l'affichage de la ligne ajoutée sur la page en cours d'affichage
    
    arrDRENTA_Nb = arrDRENTA_Nb + 1
    arrDRENTA(arrDRENTA_Nb) = newDRENTA
    xDRENTA = newDRENTA
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine arrDRENTA_Nb

    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    fgSelect.TopRow = fgSelect.Row
     
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

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    If Not blnAuto Then MsgBox "paramétrage inconsistant", vbCritical, "frmBIA_DWH.param_init"
    Unload Me
Else
    lstErr.Clear
End If
If Not IsNull(paramSAA_Init) Then
    If Not blnAuto Then MsgBox "paramétrage inconsistant", vbCritical, "frmBIA_DWH.paramSAA_Init"
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
fraSelect_Options.Enabled = True
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = False

blnControl = True

End Sub


Public Function param_Init()

param_Init = Null
Call lstErr_Clear(lstErr, cmdContext, "Param_Init"): DoEvents

fgSelect.Visible = False

Call DTPicker_Set(txtSelect_DRTAPER, DSys)

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




Private Sub cbo_CRTA_GotFocus()
txt_GotFocus cbo_CRTA
End Sub


Private Sub cbo_CRTA_LostFocus()
DRTAGRP_CGRP_SQL xWhere_CGRP
End Sub


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdfra_Renommer_Click()

Select Case cmdfra_Renommer.Caption
    Case "Renommer": fraDRENTA_Display_Rename
    Case "Valider nouvelle clé":
                Me.Enabled = False: Me.MousePointer = vbHourglass
                cmdDRENTA_Charger  'OldDRENTA a tjs anciennes valeurs
                cmdDRENTA_Rename
                   If rsado.EOF Then
                        fraDRENTA.Visible = False
                        cmdSelect_Ok.Enabled = True
                        cmdSelect_Ajouter.Enabled = True
                   End If
                Me.Enabled = True: Me.MousePointer = 0
End Select

End Sub

Private Sub cmdfra_Creer_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
cmdDRENTA_Charger
cmdDRENTA_Insert
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdfra_Modifier_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
cmdDRENTA_Charger
cmdDRENTA_Update
Me.Enabled = True: Me.MousePointer = 0
fraDRENTA.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True

' Rafraichir la liste d'affichage -fgSelect-
arrDRENTA(arrDRENTA_Index) = newDRENTA
xDRENTA = newDRENTA
fgSelect_DisplayLine arrDRENTA_Index

End Sub


Private Sub cmdfra_Supprimer_Click()
Dim X As String
X = MsgBox("Confirmer la suppression ?", vbQuestion + vbOKCancel, "BIA_DWH : DRENTA")

If X = vbOK Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    cmdDRENTA_Charger
    cmdDRENTA_Delete
    fraDRENTA.Visible = False
    cmdSelect_Ok.Enabled = True
    cmdSelect_Ajouter.Enabled = True
    Me.Enabled = True: Me.MousePointer = 0
    
    fgSelect_TopRow_Memo = fgSelect.TopRow
    ' Rafraichir la liste sur -fgSelect- et positionner à la page en cours d'affichage
    cmdSelect_SQL
    If fgSelect.Rows > 1 Then fgSelect.TopRow = fgSelect_TopRow_Memo
Else
    fraDRENTA.Visible = True
    cmdSelect_Ok.Enabled = False
    cmdSelect_Ajouter.Enabled = False
End If

End Sub


Private Sub cmdPrint_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case SSTab1.Tab
    Case 0:
           ' If fgSelect.Rows > 1 Then cmdPrint_Ok
                
    Case 1:
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_SQL()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim wAmj7 As Long
On Error GoTo Error_Handler

currentAction = "cmdDRENTA_SQL"
Call DTPicker_Control(txtSelect_DRTAPER, wAmjMax)
xWhere_CRTA = " where DRGRPER = " & wAmjMax
xWhere_CGRP = " and DRGRPER = " & wAmjMax
xWhere = " where DRTAPER = " & wAmjMax
  
DRTAGRP_CRTA_SQL xWhere_CRTA

arrDRENTA_SQL xWhere
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

xDRENTA.DRTASTA = ""
Call fraDRENTA_Display

Me.Enabled = True: Me.MousePointer = 0
End Sub



Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_DWH_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
fraDRENTA.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True

If blnOk Then
    cmdSelect_Ok.Caption = "Nouveau mois"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = False
    cmdSelect_Ajouter.Visible = True
    cmdSelect_SQL
Else
    Me.Enabled = False: Me.MousePointer = vbHourglass
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
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
        Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  arrDRENTA_Index = CLng(fgSelect.Text)
        fgSelect.LeftCol = 0
        
        oldDRENTA = arrDRENTA(arrDRENTA_Index)
        xDRENTA = oldDRENTA
        Call fraDRENTA_Display
   End If
End If
fgSelect.LeftCol = 0
End Sub

Public Sub fraDRENTA_DisplayLine()
Dim X As String

On Error Resume Next

 txtfg_DRTASTA = xDRENTA.DRTASTA
 txtfg_DRTAVER = xDRENTA.DRTAVER
 txtfg_DRTAETA = xDRENTA.DRTAETA
 txtfg_DRTACLIA = xDRENTA.DRTACLIA
 txtfg_DRTACLIB = xDRENTA.DRTACLIB
 
 ' Le code renta du DRENTA peut être en 5 ou 4 ou 3 numériques rempli
 X = Format$(xDRENTA.DRTACRTA, "00000")
 If Mid$(X, 1, 1) = "0" Then
    X = Format$(X, "0000")
    If Mid$(X, 1, 1) = "0" Then
        X = Format$(X, "000")
    End If
 End If
 Call cbo_Scan(X, cbo_CRTA)
 
 txtfg_DRTACGRP = xDRENTA.DRTACGRP
 txtfg_DRTAMOYB = cur_P(xDRENTA.DRTAMOYB)
 txtfg_DRTAMMRB = cur_P(xDRENTA.DRTAMMRB)
 txtfg_DRTATXM = Comma_Point(xDRENTA.DRTATXM)
 txtfg_DRTACTR = xDRENTA.DRTACTR

End Sub

Private Sub fraDRENTA_Display()
Dim V
Dim X As String, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim I As Long

On Error GoTo Error_Handler

fraDRENTA.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True
cmdfra_Creer.Visible = False
cmdfra_Modifier.Visible = False
cmdfra_Supprimer.Visible = False
cmdfra_Renommer.Visible = False

cmdfra_Renommer.Caption = ""
cmdfra_Renommer.Caption = "Renommer"

txtfg_DRTASTA.Enabled = True
txtfg_DRTACGRP.Enabled = True
txtfg_DRTAMMRB.Enabled = True
txtfg_DRTAMOYB.Enabled = True
txtfg_DRTATXM.Enabled = True
txtfg_DRTACTR.Enabled = True

currentAction = "fraDRENTA_Display"

If Trim(xDRENTA.DRTASTA) <> "" Then   ' Autres que Création
   
    fraDRENTA_DisplayLine
    
    cmdfra_Creer.Visible = False
    cmdfra_Modifier.Visible = True
    cmdfra_Supprimer.Visible = True
    cmdfra_Renommer.Visible = True
    txtfg_DRTAVER.Enabled = False
    txtfg_DRTAETA.Enabled = False
    txtfg_DRTACLIA.Enabled = False
    txtfg_DRTACLIB.Enabled = False
    cbo_CRTA.Enabled = False

Else  ' Création
   
    txtfg_DRTASTA = "W"
    txtfg_DRTAVER = 1
    txtfg_DRTAETA = "01"
    txtfg_DRTACLIA = ""
    txtfg_DRTACLIB = ""
    txtfg_DRTACGRP = ""
    txtfg_DRTAMMRB = ""
    txtfg_DRTAMOYB = ""
    txtfg_DRTATXM = ""
    txtfg_DRTACTR = 1

    cmdfra_Creer.Visible = True
    cmdfra_Modifier.Visible = False
    cmdfra_Supprimer.Visible = False
    cmdfra_Renommer.Visible = False

    txtfg_DRTAVER.Enabled = False
    txtfg_DRTAETA.Enabled = False
    txtfg_DRTACLIA.Enabled = True
    txtfg_DRTACLIB.Enabled = True
    cbo_CRTA.Enabled = True

End If

fraDRENTA.Visible = True
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


Private Sub fraDRENTA_Display_Rename()
Dim V
Dim X As String, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim I As Long

On Error GoTo Error_Handler

fraDRENTA.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True
cmdfra_Creer.Visible = False
cmdfra_Modifier.Visible = False
cmdfra_Supprimer.Visible = False
cmdfra_Renommer.Visible = False

cmdfra_Renommer.Caption = ""
cmdfra_Renommer.Caption = "Valider nouvelle clé"

currentAction = "fraDRENTA_Display_Rename"

fraDRENTA_DisplayLine   'Dans ce cas renseigner la clé à renommer par défaut
cmdfra_Renommer.Visible = True

txtfg_DRTAVER.Enabled = True
txtfg_DRTAETA.Enabled = True
txtfg_DRTACLIA.Enabled = True
txtfg_DRTACLIB.Enabled = True
cbo_CRTA.Enabled = True

txtfg_DRTASTA.Enabled = False
txtfg_DRTACGRP.Enabled = False
txtfg_DRTAMMRB.Enabled = False
txtfg_DRTAMOYB.Enabled = False
txtfg_DRTATXM.Enabled = False
txtfg_DRTACTR.Enabled = False

fraDRENTA.Visible = True
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
    cnADO.Close
    Set cnADO = Nothing

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
If fraDRENTA.Visible Then
    fraDRENTA.Visible = False
    cmdSelect_Ok.Enabled = True
    cmdSelect_Ajouter.Enabled = True
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

Set cnADO = New ADODB.Connection
cnADO.CursorLocation = adUseClient
cnADO.Open paramODBC_DSN_SAB

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
    Set rsADO_Update = cnADO.Execute("SET TRANSACTION ISOLATION LEVEL READ COMMITTED")

End If

End Sub




Private Sub txtfg_DRTACGRP_GotFocus()
txt_GotFocus txtfg_DRTACLIB
DRTAGRP_CGRP_SQL xWhere_CGRP
End Sub

Private Sub txtfg_DRTACGRP_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtfg_DRTACLIA_GotFocus()
txt_GotFocus txtfg_DRTACLIA
End Sub

Private Sub txtfg_DRTACLIB_GotFocus()
txt_GotFocus txtfg_DRTACLIB
End Sub




Private Sub txtfg_DRTACTR_GotFocus()
txt_GotFocus txtfg_DRTACTR
End Sub


Private Sub txtfg_DRTAETA_GotFocus()
txt_GotFocus txtfg_DRTAETA
End Sub

Private Sub txtfg_DRTAETA_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtfg_DRTAMMRB_KeyPress(KeyAscii As Integer)
num_KeyAsciiD KeyAscii, txtfg_DRTAMMRB
End Sub


Private Sub txtfg_DRTAMMRB_LostFocus()
txtfg_DRTAMMRB = cur_P(CCur(Fix(Val(txtfg_DRTAMMRB) * 100) / 100))
End Sub

Private Sub txtfg_DRTAMOYB_KeyPress(KeyAscii As Integer)
num_KeyAsciiD KeyAscii, txtfg_DRTAMOYB
End Sub


Private Sub txtfg_DRTAMOYB_LostFocus()
txtfg_DRTAMOYB = cur_P(CCur(Fix(Val(txtfg_DRTAMOYB) * 100) / 100))
End Sub


Private Sub txtfg_DRTASTA_GotFocus()
txt_GotFocus txtfg_DRTASTA
End Sub

Private Sub txtfg_DRTASTA_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DRTASTA_LostFocus()
txt_LostFocus txtfg_DRTASTA

End Sub


Private Sub txtfg_DRTATXM_KeyPress(KeyAscii As Integer)
num_KeyAsciiD KeyAscii, txtfg_DRTATXM
End Sub

Private Sub txtfg_DRTATXM_LostFocus()
txtfg_DRTATXM = Comma_Point(CDbl(Fix(Val(txtfg_DRTATXM) * 1000000) / 1000000))
End Sub


Private Sub txtfg_DRTAVER_GotFocus()
txt_GotFocus txtfg_DRTAVER
End Sub





