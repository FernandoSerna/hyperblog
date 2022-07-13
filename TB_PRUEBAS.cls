VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_PRUEBAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim cmd As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_ART_ARTICULO_ID As String, clVcha_art_nombre_español As String, clVcha_lin_linea_id As String, clVcha_dis_diseño_id As String, clMone_art_precio_base As String, clMone_art_costo_estandar As String, clDtim_art_fecha_alta As String, clDtim_art_fecha_baja As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set cmd.ActiveConnection = cnn
   cmd.CommandType = adCmdStoredProc
   cmd.CommandText = "pruebas_I"
   cmd("@Vcha_art_articulo_id") = clVCHA_ART_ARTICULO_ID
   cmd("@Vcha_art_nombre_español") = clVcha_art_nombre_español
   cmd("@vcha_lin_linea_id") = clVcha_lin_linea_id
   cmd("@Vcha_dis_diseño_id") = clVcha_dis_diseño_id
   cmd("@Mone_art_precio_base") = clMone_art_precio_base
   cmd("@Mone_art_costo_estandar") = clMone_art_costo_estandar
   cmd("@Dtim_art_fecha_alta") = clDtim_art_fecha_alta
   If clDtim_art_fecha_baja <> "" Then
       cmd("@Dtim_art_fecha_baja") = clDtim_art_fecha_baja
    Else
       cmd("@Dtim_art_fecha_baja") = Null
    End If
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









