Attribute VB_Name = "srvYMNUETA0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recEdition_GestionLen = 109 ' 34 + 75
Public Const memoEdition_GestionLen = 75
Public Const constEdition_Form = "Edition_Form"
''Public Const constEdition_Usr = "Edition_Usr"
Type typeEdition_Form
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    K1              As String * 12
    K2              As String * 12
    Name            As String * 40
    Courrier        As String * 1
    Orientation     As String * 1
    LinePerPage      As Integer
    FontSize        As Integer
    Duplex          As String * 1
    Filigrane       As String * 1
    Copies          As Integer
    PaperBin        As String * 1
    Hold            As String * 1
    Save            As String * 1
    FontName        As String * 30
    PrinterUnit     As String * 1           ' impression poste utilisateur sinon imp réseau  du service
    Unit            As String * 10
    Unit2           As String * 10          ' service destintaire d'une copie
    Unit3           As String * 10          ' service destintaire d'une copie
End Type
    


Public Const recYMNUETA0Len = 77 ' 34 + 43
Public Const recYMNUETA0_Block = 200
Type typeYMNUETA0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    MNUETAETA       As String * 10                    ' ETAT
    MNUETACLI       As String * 7                     ' CLIENT
    MNUETAOUT       As String * 10                    ' FILE ATTENTE
    MNUETALIB       As String * 30                    ' LIBELLE
    MNUETATYP       As String * 10                    ' TYPE IMPRIME
    MNUETAPOL       As String * 10                    ' ID. POLICE
    MNUETALON       As Long                           ' LONGUEUR PAGE
    MNUETALAR       As Long                           ' LARGEUR PAGE
    MNUETAFIN       As Long                           ' LIGNE FIN PAGE
    MNUETALPO       As String * 1                     ' LIGNE POUCE
    MNUETACPO       As String * 4                     ' CARACTERE POUCE
    MNUETAROT       As String * 5                     ' ROTATION PAGE
    MNUETANEX       As Long                           ' NOMBRE EXEMPLAIRE
    MNUETASUS       As String * 4                     ' SUSPENDRE
    MNUETACON       As String * 4                     ' CONSERVER
    MNUETAPRI       As String * 4                     ' PRIORITE SORTIE
    MNUETAQUA       As String * 6                     ' QUALITE IMPRESS.
    MNUETAAVI       As String * 1                     ' AVIS CLIENT. (B,1,2)
    MNUETAFON       As String * 8                     ' FRONT PAGE
    MNUETAFOL       As String * 10                    ' BIBLIO FRONT PAGE
End Type
    
'---------------------------------------------------------
Public Function srvYMNUETA0_GetBuffer(recYMNUETA0 As typeYMNUETA0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYMNUETA0_GetBuffer = Null
recYMNUETA0.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYMNUETA0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYMNUETA0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYMNUETA0.Err = Space$(10) Then
    recYMNUETA0.MNUETAETA = mId$(MsgTxt, K + 1, 10)
    recYMNUETA0.MNUETACLI = mId$(MsgTxt, K + 11, 7)
    recYMNUETA0.MNUETAOUT = mId$(MsgTxt, K + 18, 10)
    recYMNUETA0.MNUETALIB = mId$(MsgTxt, K + 28, 30)
    recYMNUETA0.MNUETATYP = mId$(MsgTxt, K + 58, 10)
    recYMNUETA0.MNUETAPOL = mId$(MsgTxt, K + 68, 10)
    recYMNUETA0.MNUETALON = CLng(Val(mId$(MsgTxt, K + 78, 4)))
    recYMNUETA0.MNUETALAR = CLng(Val(mId$(MsgTxt, K + 82, 4)))
    recYMNUETA0.MNUETAFIN = CLng(Val(mId$(MsgTxt, K + 86, 4)))
    recYMNUETA0.MNUETALPO = mId$(MsgTxt, K + 90, 1)
    recYMNUETA0.MNUETACPO = mId$(MsgTxt, K + 91, 4)
    recYMNUETA0.MNUETAROT = mId$(MsgTxt, K + 95, 5)
    recYMNUETA0.MNUETANEX = CLng(Val(mId$(MsgTxt, K + 100, 4)))
    recYMNUETA0.MNUETASUS = mId$(MsgTxt, K + 104, 4)
    recYMNUETA0.MNUETACON = mId$(MsgTxt, K + 108, 4)
    recYMNUETA0.MNUETAPRI = mId$(MsgTxt, K + 112, 4)
    recYMNUETA0.MNUETAQUA = mId$(MsgTxt, K + 116, 6)
    recYMNUETA0.MNUETAAVI = mId$(MsgTxt, K + 122, 1)
    recYMNUETA0.MNUETAFON = mId$(MsgTxt, K + 123, 8)
    recYMNUETA0.MNUETAFOL = mId$(MsgTxt, K + 131, 10)
Else
    srvYMNUETA0_GetBuffer = recYMNUETA0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYMNUETA0Len

End Function

'---------------------------------------------------------
Public Function srvEdition_Form_GetBuffer(recEdition_Gestion As typeEdition_Form)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvEdition_Form_GetBuffer = Null
recEdition_Gestion.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recEdition_Gestion.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recEdition_Gestion.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recEdition_Gestion.Err = Space$(10) Then
    recEdition_Gestion.Courrier = mId$(MsgTxt, K + 1, 1)
    recEdition_Gestion.Orientation = mId$(MsgTxt, K + 2, 1)
    recEdition_Gestion.LinePerPage = CInt(mId$(MsgTxt, K + 3, 3))
    recEdition_Gestion.FontSize = CInt(mId$(MsgTxt, K + 6, 2))
    recEdition_Gestion.Duplex = mId$(MsgTxt, K + 8, 1)
    recEdition_Gestion.Filigrane = mId$(MsgTxt, K + 9, 1)
    recEdition_Gestion.Copies = CInt(mId$(MsgTxt, K + 10, 2))
    recEdition_Gestion.PaperBin = mId$(MsgTxt, K + 12, 1)
    recEdition_Gestion.Hold = mId$(MsgTxt, K + 13, 1)
    recEdition_Gestion.Save = mId$(MsgTxt, K + 14, 1)
    recEdition_Gestion.FontName = mId$(MsgTxt, K + 15, 30)
    recEdition_Gestion.PrinterUnit = mId$(MsgTxt, K + 45, 1)
    recEdition_Gestion.Unit = mId$(MsgTxt, K + 46, 10)
    recEdition_Gestion.Unit2 = mId$(MsgTxt, K + 56, 10)
    recEdition_Gestion.Unit3 = mId$(MsgTxt, K + 66, 10)
    
Else
    srvEdition_Form_GetBuffer = recEdition_Gestion.Err
End If

MsgTxtIndex = MsgTxtIndex + recEdition_GestionLen

End Function

'---------------------------------------------------------
Public Sub srvEdition_Form_PutBuffer(recEdition_Gestion As typeEdition_Form)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recEdition_GestionLen) = Space$(recEdition_GestionLen)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recEdition_Gestion.Obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recEdition_Gestion.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 1) = recEdition_Gestion.Courrier
    Mid$(MsgTxt, K + 2, 1) = recEdition_Gestion.Orientation
    Mid$(MsgTxt, K + 3, 3) = Format$(recEdition_Gestion.LinePerPage, "000")
    Mid$(MsgTxt, K + 6, 2) = Format$(recEdition_Gestion.FontSize, "00")
    Mid$(MsgTxt, K + 8, 1) = recEdition_Gestion.Duplex
    Mid$(MsgTxt, K + 9, 1) = recEdition_Gestion.Filigrane
    Mid$(MsgTxt, K + 10, 2) = Format$(recEdition_Gestion.Copies, "00")
    Mid$(MsgTxt, K + 12, 1) = recEdition_Gestion.PaperBin
    Mid$(MsgTxt, K + 13, 1) = recEdition_Gestion.Hold
    Mid$(MsgTxt, K + 14, 1) = recEdition_Gestion.Save
    Mid$(MsgTxt, K + 15, 30) = recEdition_Gestion.FontName
    Mid$(MsgTxt, K + 45, 1) = recEdition_Gestion.PrinterUnit
    Mid$(MsgTxt, K + 46, 10) = recEdition_Gestion.Unit
    Mid$(MsgTxt, K + 56, 10) = recEdition_Gestion.Unit2
    Mid$(MsgTxt, K + 66, 10) = recEdition_Gestion.Unit3

MsgTxtLen = MsgTxtLen + recEdition_GestionLen
End Sub

Public Sub Table_Edition_Form_Init(lEdition_Form As typeEdition_Form)
lEdition_Form.Obj = "ZMNUETA0_S"
lEdition_Form.Method = ""
lEdition_Form.Err = ""
lEdition_Form.K1 = ""
lEdition_Form.K2 = ""
lEdition_Form.Name = ""
lEdition_Form.Courrier = "0"
lEdition_Form.FontSize = 6
lEdition_Form.Orientation = "1"
lEdition_Form.Duplex = "1"
lEdition_Form.Filigrane = "A"
lEdition_Form.Copies = 1
lEdition_Form.PaperBin = 2
lEdition_Form.Hold = "0"
lEdition_Form.Save = "0"
lEdition_Form.FontName = prtFontName_CourierNew
lEdition_Form.PrinterUnit = "0"
lEdition_Form.Unit = ""
lEdition_Form.Unit2 = ""
lEdition_Form.Unit3 = ""


End Sub


Public Sub srvYMNUETA0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YMNUETA0.txt" For Input As #1
Open "C:\Temp\YMNUETA0.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "MNUETAETA;MNUETACLI;MNUETAOUT;MNUETALIB;MNUETATYP;MNUETAPOL;MNUETALON;MNUETALAR;MNUETAFIN;MNUETALPO;MNUETACPO;MNUETAROT;MNUETANEX;MNUETASUS;MNUETACON;MNUETAPRI;MNUETAQUA;MNUETAAVI;MNUETAFON;MNUETAFOL;"
    Print #2, "ETAT;CLIENT;FILE ATTENTE;LIBELLE;TYPE IMPRIME;ID. POLICE;LONGUEUR PAGE;LARGEUR PAGE;LIGNE FIN PAGE;LIGNE POUCE;CARACTERE POUCE;ROTATION PAGE;NOMBRE EXEMPLAIRE;SUSPENDRE;CONSERVER;PRIORITE SORTIE;QUALITE IMPRESS.;AVIS CLIENT. (B,1,2);FRONT PAGE;BIBLIO FRONT PAGE;"
    Print #2, ";;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 10) & ";" _
      & mId$(xIn, 11, 7) & ";" _
      & mId$(xIn, 18, 10) & ";" _
      & mId$(xIn, 28, 30) & ";" _
      & mId$(xIn, 58, 10) & ";" _
      & mId$(xIn, 68, 10) & ";" _
      & mId$(xIn, 78, 4) & ";" _
      & mId$(xIn, 82, 4) & ";" _
      & mId$(xIn, 86, 4) & ";" _
      & mId$(xIn, 90, 1) & ";" _
      & mId$(xIn, 91, 4) & ";" _
      & mId$(xIn, 95, 5) & ";" _
      & mId$(xIn, 100, 4) & ";" _
      & mId$(xIn, 104, 4) & ";" _
      & mId$(xIn, 108, 4) & ";" _
      & mId$(xIn, 112, 4) & ";" _
      & mId$(xIn, 116, 6) & ";" _
      & mId$(xIn, 122, 1) & ";" _
      & mId$(xIn, 123, 8) & ";" _
      & mId$(xIn, 131, 10) & ";"
Loop
Close
End Sub



