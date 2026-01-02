Attribute VB_Name = "rsEdition_Form"
Option Explicit


Public Const constEdition_Form = "Edition_Form"
Type typeEdition_Form
    
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

'---------------------------------------------------------
Public Sub rsEdition_Form_PutBuffer(lMemo As String, recEdition_Gestion As typeEdition_Form)
'---------------------------------------------------------

Mid$(lMemo, 1, 1) = recEdition_Gestion.Courrier
Mid$(lMemo, 2, 1) = recEdition_Gestion.Orientation
Mid$(lMemo, 3, 3) = Format$(recEdition_Gestion.LinePerPage, "000")
Mid$(lMemo, 6, 2) = Format$(recEdition_Gestion.FontSize, "00")
Mid$(lMemo, 8, 1) = recEdition_Gestion.Duplex
Mid$(lMemo, 9, 1) = recEdition_Gestion.Filigrane
Mid$(lMemo, 10, 2) = Format$(recEdition_Gestion.Copies, "00")
Mid$(lMemo, 12, 1) = recEdition_Gestion.PaperBin
Mid$(lMemo, 13, 1) = recEdition_Gestion.Hold
Mid$(lMemo, 14, 1) = recEdition_Gestion.Save
Mid$(lMemo, 15, 30) = recEdition_Gestion.FontName
Mid$(lMemo, 45, 1) = recEdition_Gestion.PrinterUnit
Mid$(lMemo, 46, 10) = recEdition_Gestion.Unit
Mid$(lMemo, 56, 10) = recEdition_Gestion.Unit2
Mid$(lMemo, 66, 10) = recEdition_Gestion.Unit3

End Sub

'---------------------------------------------------------
Public Function rsEdition_Form_GetBuffer(lMemo As String, recEdition_Gestion As typeEdition_Form)
'---------------------------------------------------------
rsEdition_Form_GetBuffer = Null

    recEdition_Gestion.Courrier = mId$(lMemo, 1, 1)
    recEdition_Gestion.Orientation = mId$(lMemo, 2, 1)
    recEdition_Gestion.LinePerPage = CInt(mId$(lMemo, 3, 3))
    recEdition_Gestion.FontSize = CInt(mId$(lMemo, 6, 2))
    recEdition_Gestion.Duplex = mId$(lMemo, 8, 1)
    recEdition_Gestion.Filigrane = mId$(lMemo, 9, 1)
    recEdition_Gestion.Copies = CInt(mId$(lMemo, 10, 2))
    recEdition_Gestion.PaperBin = mId$(lMemo, 12, 1)
    recEdition_Gestion.Hold = mId$(lMemo, 13, 1)
    recEdition_Gestion.Save = mId$(lMemo, 14, 1)
    recEdition_Gestion.FontName = mId$(lMemo, 15, 30)
    recEdition_Gestion.PrinterUnit = mId$(lMemo, 45, 1)
    recEdition_Gestion.Unit = mId$(lMemo, 46, 10)
    recEdition_Gestion.Unit2 = mId$(lMemo, 56, 10)
    recEdition_Gestion.Unit3 = mId$(lMemo, 66, 10)
    

End Function


