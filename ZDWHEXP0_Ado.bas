Attribute VB_Name = "adoZDWHEXP0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZDWHEXP0_PutBuffer(rsADO As ADODB.Recordset, rsZDWHEXP0 As typeZDWHEXP0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZDWHEXP0_PutBuffer = Null
rsADO("DWHEXPDTA") = rsZDWHEXP0.DWHEXPDTA
rsADO("DWHEXPETA") = rsZDWHEXP0.DWHEXPETA
rsADO("DWHEXPAGE") = rsZDWHEXP0.DWHEXPAGE
rsADO("DWHEXPSER") = rsZDWHEXP0.DWHEXPSER
rsADO("DWHEXPSSE") = rsZDWHEXP0.DWHEXPSSE
rsADO("DWHEXPPLA") = rsZDWHEXP0.DWHEXPPLA
rsADO("DWHEXPOPE") = rsZDWHEXP0.DWHEXPOPE
rsADO("DWHEXPNAT") = rsZDWHEXP0.DWHEXPNAT
rsADO("DWHEXPNUM") = rsZDWHEXP0.DWHEXPNUM
rsADO("DWHEXPTYP") = rsZDWHEXP0.DWHEXPTYP
rsADO("DWHEXPCOM") = rsZDWHEXP0.DWHEXPCOM
rsADO("DWHEXPDEV") = rsZDWHEXP0.DWHEXPDEV
rsADO("DWHEXPFIN") = rsZDWHEXP0.DWHEXPFIN
rsADO("DWHEXPDUI") = rsZDWHEXP0.DWHEXPDUI
rsADO("DWHEXPDUR") = rsZDWHEXP0.DWHEXPDUR
rsADO("DWHEXPTYO") = rsZDWHEXP0.DWHEXPTYO
rsADO("DWHEXPCLI") = rsZDWHEXP0.DWHEXPCLI
rsADO("DWHEXPTAU") = rsZDWHEXP0.DWHEXPTAU
rsADO("DWHEXPENC") = rsZDWHEXP0.DWHEXPENC
rsADO("DWHEXPINT") = rsZDWHEXP0.DWHEXPINT
rsADO("DWHEXPIMP") = rsZDWHEXP0.DWHEXPIMP
rsADO("DWHEXPEXB") = rsZDWHEXP0.DWHEXPEXB
rsADO("DWHEXPPRO") = rsZDWHEXP0.DWHEXPPRO
rsADO("DWHEXPEXN") = rsZDWHEXP0.DWHEXPEXN
rsADO("DWHEXPCAT") = rsZDWHEXP0.DWHEXPCAT
rsADO("DWHEXPREG") = rsZDWHEXP0.DWHEXPREG
rsADO("DWHEXPTXP") = rsZDWHEXP0.DWHEXPTXP
rsADO("DWHEXPEXA") = rsZDWHEXP0.DWHEXPEXA
rsADO("DWHEXPEAP") = rsZDWHEXP0.DWHEXPEAP
rsADO("DWHEXPEXS") = rsZDWHEXP0.DWHEXPEXS
rsADO("DWHEXPESP") = rsZDWHEXP0.DWHEXPESP
rsADO("DWHEXPEXR") = rsZDWHEXP0.DWHEXPEXR
rsADO("DWHEXPFIL") = rsZDWHEXP0.DWHEXPFIL
    
Exit Function

Error_Handler:

rsZDWHEXP0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZDWHEXP0_AddNew(rsADO As ADODB.Recordset, rsZDWHEXP0 As typeZDWHEXP0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZDWHEXP0_AddNew = Null
rsADO.AddNew
adoZDWHEXP0_AddNew = rsZDWHEXP0_PutBuffer(rsADO, rsZDWHEXP0)
rsADO.Update

Exit Function

Error_Handler:

adoZDWHEXP0_AddNew = Error

End Function


