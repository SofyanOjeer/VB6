Attribute VB_Name = "adoYBIASTO0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYBIASTO0_PutBuffer(rsADO As ADODB.Recordset, rsYBIASTO0 As typeYBIASTO0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYBIASTO0_PutBuffer = Null

rsADO("YSTOETA") = rsYBIASTO0.YSTOETA
rsADO("YSTOAGE") = rsYBIASTO0.YSTOAGE
rsADO("YSTOSER") = rsYBIASTO0.YSTOSER
rsADO("YSTOSSE") = rsYBIASTO0.YSTOSSE
rsADO("YSTOOPE") = rsYBIASTO0.YSTOOPE
rsADO("YSTONUM") = rsYBIASTO0.YSTONUM
rsADO("YSTOSEQ") = rsYBIASTO0.YSTOSEQ
rsADO("YSTOPCI") = rsYBIASTO0.YSTOPCI
rsADO("YSTOCCL") = rsYBIASTO0.YSTOCCL
rsADO("YSTOCLI") = rsYBIASTO0.YSTOCLI
rsADO("YSTODEV") = rsYBIASTO0.YSTODEV
rsADO("YSTOMON") = rsYBIASTO0.YSTOMON
rsADO("YSTODEB") = rsYBIASTO0.YSTODEB
rsADO("YSTOFIN") = rsYBIASTO0.YSTOFIN
rsADO("YSTOAPP") = rsYBIASTO0.YSTOAPP
rsADO("YSTONAT") = rsYBIASTO0.YSTONAT
rsADO("YSTOCC1") = rsYBIASTO0.YSTOCC1
rsADO("YSTOCL1") = rsYBIASTO0.YSTOCL1
rsADO("YSTOCC2") = rsYBIASTO0.YSTOCC2
rsADO("YSTOCL2") = rsYBIASTO0.YSTOCL2
rsADO("YSTOCTX") = rsYBIASTO0.YSTOCTX
rsADO("YSTOTAU") = rsYBIASTO0.YSTOTAU

    
Exit Function

Error_Handler:

rsYBIASTO0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoYBIASTO0_AddNew(rsADO As ADODB.Recordset, rsYBIASTO0 As typeYBIASTO0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoYBIASTO0_AddNew = Null
rsADO.AddNew
adoYBIASTO0_AddNew = rsYBIASTO0_PutBuffer(rsADO, rsYBIASTO0)
rsADO.Update

Exit Function

Error_Handler:

adoYBIASTO0_AddNew = Error

End Function

