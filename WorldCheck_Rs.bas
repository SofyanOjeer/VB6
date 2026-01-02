Attribute VB_Name = "rsWorldCheck"
Option Explicit

Public cnWC As New ADODB.Connection
Public rsWC As New ADODB.Recordset
Public WC_DataBase_Name As String

Type typeWC_Data
    WC_Id           As Long
    WC_UpdD         As String
    WC_UpdH         As String
    WC_Sta          As String
    WC_LastName     As String
    WC_FirstName    As String
    WC_Memo         As String

End Type


Public Sub rsWC_Data_Init(rsWC_Data As typeWC_Data)
rsWC_Data.WC_Id = 0
rsWC_Data.WC_UpdD = ""
rsWC_Data.WC_UpdH = ""
rsWC_Data.WC_Sta = ""
rsWC_Data.WC_LastName = ""
rsWC_Data.WC_FirstName = ""
rsWC_Data.WC_Memo = ""


End Sub

'---------------------------------------------------------
Public Function rsWC_Data_GetBuffer(rsADO As ADODB.Recordset, rsWC_Data As typeWC_Data)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsWC_Data_GetBuffer = Null

rsWC_Data.WC_Id = rsADO("WC_Id")
rsWC_Data.WC_UpdD = rsADO("WC_UpdD")
rsWC_Data.WC_UpdH = rsADO("WC_UpdH")
rsWC_Data.WC_Sta = rsADO("WC_Sta")
rsWC_Data.WC_LastName = rsADO("WC_LAstName")
rsWC_Data.WC_FirstName = rsADO("WC_FirstName")
rsWC_Data.WC_Memo = rsADO("WC_Memo")
Exit Function

Error_Handler:

rsWC_Data_GetBuffer = Error
End Function

Public Sub WC_Data_Open()
Set cnWC = New ADODB.Connection
cnWC.Provider = "Microsoft.Jet.OLEDB.4.0"
cnWC.Properties("JET OLEDB:Database Password") = ""
cnWC.Mode = adModeReadWrite
WC_DataBase_Name = "C:\Temp\World-Check\WorldCheck.mdb"
cnWC.Open WC_DataBase_Name

If UCase$(WC_DataBase_Name) <> UCase$(cnWC.Properties("Data Source Name")) Then
    MsgBox WC_DataBase_Name, vbCritical, " non conforme "
    cnAdo_Info cnWC
    End
End If

End Sub
