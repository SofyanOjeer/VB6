Attribute VB_Name = "rsZDWHEHB0"
'---------------------------------------------------------
Option Explicit
Type typeZDWHEHB0

  
  DWHEHBDTX   As Long        'DATE ANALYSE                          1       8      0     S  *
  DWHEHBETA   As Long        'ETABLISSEMENT                         9       4      0     S  *
  DWHEHBAGE   As Long        'AGENCE                               13       4      0     S  *
  DWHEHBSER   As String * 2  'SERVICE                              17       2            A  *
  DWHEHBSSE   As String * 2  'SOUS-SERVICE                         19       2            A  *
  DWHEHBOPE   As String * 6  'CODE OPERATION                       21       6            A  *
  DWHEHBNAT   As String * 6  'CODE NATURE                          27       6            A  *
  DWHEHBNDO   As Long        'N° OPERATION                         33       9      0     S  *
  DWHEHBPOO   As String * 2  'TYPE DE POOL                         42       2            A  *
  DWHEHBPOU   As Double      ' % DE PARTICIPATION                   44      14      9     S  *
  DWHEHBMBE   As Currency    'MT BRUT ENG. DEV.BAS                 58      18      3     S  *
  DWHEHBMNE   As Currency    'MT NET ENG. DEV. BAS                 76      18      3     S  *
  DWHEHBNUM   As String * 20 'NUM.COMPTE/NUM.AUT                   94      20            A  *
  DWHEHBAUT   As String * 1  'TYPE AUTORISATION                   114       1            A  *
  DWHEHBRUB   As String * 10 'RUBRIQUE COMPTABLE                  115      10            A  *
  DWHEHBOBJ   As String * 6  'OBJET DE FINANCEMENT                125       6            A  *
  DWHEHBDSY   As Long        'DATE SYSTEME                        131       8      0     S  *


End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZDWHEHB0_GetBuffer(rsSab As ADODB.Recordset, rsZDWHEHB0 As typeZDWHEHB0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZDWHEHB0_GetBuffer = Null


rsZDWHEHB0.DWHEHBDTX = rsSab("DWHEHBDTX")
rsZDWHEHB0.DWHEHBETA = rsSab("DWHEHBETA")
rsZDWHEHB0.DWHEHBAGE = rsSab("DWHEHBAGE")
rsZDWHEHB0.DWHEHBSER = rsSab("DWHEHBSER")
rsZDWHEHB0.DWHEHBSSE = rsSab("DWHEHBSSE")
rsZDWHEHB0.DWHEHBOPE = rsSab("DWHEHBOPE")
rsZDWHEHB0.DWHEHBNAT = rsSab("DWHEHBNAT")
rsZDWHEHB0.DWHEHBNDO = rsSab("DWHEHBNDO")
rsZDWHEHB0.DWHEHBPOO = rsSab("DWHEHBPOO")
rsZDWHEHB0.DWHEHBPOU = rsSab("DWHEHBPOU")
rsZDWHEHB0.DWHEHBMBE = rsSab("DWHEHBMBE")
rsZDWHEHB0.DWHEHBMNE = rsSab("DWHEHBMNE")
rsZDWHEHB0.DWHEHBNUM = rsSab("DWHEHBNUM")
rsZDWHEHB0.DWHEHBAUT = rsSab("DWHEHBAUT")
rsZDWHEHB0.DWHEHBRUB = rsSab("DWHEHBRUB")
rsZDWHEHB0.DWHEHBOBJ = rsSab("DWHEHBOBJ")
rsZDWHEHB0.DWHEHBDSY = rsSab("DWHEHBDSY")

Exit Function

Error_Handler:

rsZDWHEHB0_GetBuffer = Error

End Function

