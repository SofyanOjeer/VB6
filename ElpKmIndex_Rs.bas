Attribute VB_Name = "rsElpKmIndex"
'---------------------------------------------------------
Option Explicit
Type typeElpKmIndex
    Method                  As String * 12

    Id                      As String * 16
    Classe                  As Long
    ElpKMSrc_Id             As Long             'As String * 20
    Memo                    As Variant
    
    
End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsElpKmIndex_GetBuffer(rsAdo As ADODB.Recordset, rsElpKmIndex As typeElpKmIndex)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsElpKmIndex_GetBuffer = Null

rsElpKmIndex.Method = ""

rsElpKmIndex.Id = rsAdo("ID")
rsElpKmIndex.Classe = rsAdo("Classe")
rsElpKmIndex.ElpKMSrc_Id = rsAdo("ElpKMSrc_Id")
rsElpKmIndex.Memo = rsAdo("Memo")

Exit Function

Error_Handler:

rsElpKmIndex_GetBuffer = Error

End Function

