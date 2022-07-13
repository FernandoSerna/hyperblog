VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_SUBMENUS_MODIFICA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim cmd As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVcha_men_menu_id As String, clCHAR_SME_SUBMENU_ID As String, clINTE_SME_NIVEL1 As Integer, clINTE_SME_NIVEL2 As Integer, _
clINTE_SME_NIVEL3 As Integer, clINTE_SME_NIVEL4 As Integer, clINTE_SME_NIVEL5 As Integer, clVCHA_SME_DESCRIPCION As String, clINTE_SME_NIVEL As Integer, clVCHA_FOR_FORMA_ID As String, _
clinte_sme_permiso1 As Integer, clinte_sme_permiso2 As Integer, clinte_sme_permiso3 As Integer, clinte_sme_permiso4 As Integer, clinte_int_numero As Integer) As Boolean
Dim rs As adodb.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set cmd.ActiveConnection = cnn
   cmd.CommandType = adCmdStoredProc
   cmd.CommandText = "SUBMENUS_M"
   cmd("@VCHA_MEN_MENU_ID") = clVcha_men_menu_id
   cmd("@CHAR_SME_SUBMENU_ID") = clCHAR_SME_SUBMENU_ID
   cmd("@INTE_SME_NIVEL1") = clINTE_SME_NIVEL1
   cmd("@INTE_SME_NIVEL2") = clINTE_SME_NIVEL2
   cmd("@INTE_SME_NIVEL3") = clINTE_SME_NIVEL3
   cmd("@INTE_SME_NIVEL4") = clINTE_SME_NIVEL4
   cmd("@INTE_SME_NIVEL5") = clINTE_SME_NIVEL5
   cmd("@VCHA_SME_DESCRIPCION") = clVCHA_SME_DESCRIPCION
   cmd("@INTE_SME_NIVEL") = clINTE_SME_NIVEL
   cmd("@VCHA_FOR_FORMA_ID") = clVCHA_FOR_FORMA_ID
   cmd("@INTE_SME_PERMISO1") = clinte_sme_permiso1
   cmd("@INTE_SME_PERMISO2") = clinte_sme_permiso2
   cmd("@INTE_SME_PERMISO3") = clinte_sme_permiso3
   cmd("@INTE_SME_PERMISO4") = clinte_sme_permiso4
   cmd("@INTE_SME_NUMERO") = clinte_int_numero
   cmd.execute
Set cmd = Nothing
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Anadir = False
    GoTo SIGUE
End Function












