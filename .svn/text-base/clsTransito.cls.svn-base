VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public cnnOra As New ADODB.Connection
Public rsTransito As New ADODB.recordSet
Public cmdTrasito As New ADODB.Command
Public cmdPoliza As New ADODB.Command

Public Function fun_GuardaTransito(strQueryTransto As String) As Boolean

On Error GoTo errorTransito:
    Dim StrMensage As String
    Dim numParametros As Integer
    Dim numTotParamet As Integer
    
    fun_GuardaTransito = True
    If conectaOracle("interface", "ap", "interface") Then
        cnnOra.BeginTrans
        rsTransito.Open strQueryTransto, _
            cnn, _
            adOpenDynamic, _
            adLockOptimistic
        If rsTransito.RecordCount <> 0 Then
                With cmdTrasito
                    .ActiveConnection = cnnOra
                    .CommandText = "SP_TRANSITO_INSUPD"
                    .CommandType = adCmdStoredProc
                    Set objParam = .CreateParameter("P_TIPO_MOVIMIENTO", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TIPO_MOVIMIENTO").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_CONSECUTIVO", adBigInt, adParamInputOutput, 30, Val(rsTransito("NUMB_TRA_CONSECUTIVO").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_NOTA_ENVIO", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_NOTA_ENVIO").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_PLANTA_DESTINO", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_PLANTA_DESTINO").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_PLANTA_ORIGEN", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_PLANTA_ORIGEN").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_ALMACEN_DESTINO", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_ALMACEN_DESTINO").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_ALMACEN_ORIGEN", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_ALMACEN_ORIGEN").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_EMP_EMPRESA_ORIGEN", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_EMP_EMPRESA_ORIGEN").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_EMP_EMPRESA_DESTINO", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_EMP_EMPRESA_DESTINO").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_UOR_UNIDAD_ORIGEN", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_UOR_UNIDAD_ORIGEN").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_UOR_UNIDAD_DESTINO", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_UOR_UNIDAD_DESTINO").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_MOV_MOVIMIENTO_ORIGEN", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_MOV_MOVIMIENTO_ORIGEN").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_MOV_MOVIMIENTO_DESTINO", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_MOV_MOVIMIENTO_DESTINO").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_SER_SERIE_ORIGEN", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_SER_SERIE_ORIGEN").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_SER_SERIE_DESTINO", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_SER_SERIE_DESTINO").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_ART_ARTICULO_ORIGEN", adVarChar, adParamInputOutput, 30, rsTransito("VCHA_ART_ARTICULO_ORIGEN").Value)
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_ART_ARTICULO_DESTINO", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_ART_ARTICULO_DESTINO").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_ART_DESCRIPCION", adVarChar, adParamInputOutput, 250, CStr(rsTransito("VCHA_ART_DESCRIPCION").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_CANTIDAD_ENVIADA", adBigInt, adParamInputOutput, 30, Val(rsTransito("NUMB_TRA_CANTIDAD_ENVIADA").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_CANTIDAD_RECIBIDA", adBigInt, adParamInputOutput, 30, Val(rsTransito("NUMB_TRA_CANTIDAD_RECIBIDA").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_COSTO", adBigInt, adParamInputOutput, 30, Val(Round(rsTransito("NUMB_TRA_COSTO").Value, 3)))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_CALIDAD", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_CALIDAD").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_SISTEMA_ORIGEN", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_SISTEMA_ORIGEN").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_SISTEMA_DESTINO", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_SISTEMA_DESTINO").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_STATUS_ID", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_STATUS_ID").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_USUARIO_ORIGEN", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_USUARIO_ORIGEN").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_USUARIO_DESTINO", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_USUARIO_DESTINO").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_MAQUINA_ORIGEN", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_MAQUINA_ORIGEN").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_MAQUINA_DESTINO", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_MAQUINA_DESTINO").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_REFERENCIA1", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_REFERENCIA1").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_REFERENCIA2", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_REFERENCIA2").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_COM_TIPO_PROVEEDOR", adVarChar, adParamInputOutput, 30, CStr(rsTransito("CHAR_COM_TIPO_PROVEEDOR").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_TRANSPORTO", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_TRANSPORTO").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_CONTENEDOR_ID", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_TRA_CONTENEDOR_ID").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_PESO", adBigInt, adParamInputOutput, 30, Val(rsTransito("NUMB_TRA_PESO").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_MON_MONEDA_ID", adVarChar, adParamInputOutput, 30, CStr(rsTransito("VCHA_MON_MONEDA_ID").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_TRA_PRECIO", adBigInt, adParamInputOutput, 30, Val(rsTransito("NUMB_TRA_PRECIO").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_DTR_FOLIO_RECEPCION", adVarChar, adParamInputOutput, 30, CStr(rsTransito("numb_DTR_FOLIO_RECEPCION").Value))
                    .Parameters.Append (objParam)
                    Set objParam = .CreateParameter("P_DTR_FOLIO_RECEPCION", adVarChar, adParamInputOutput, 30, CStr(rsTransito("numb_DTR_FOLIO_RECEPCION").Value))
                    .Parameters.Append (objParam)
                End With
            While Not rsTransito.EOF
                With cmdTrasito
                    .ActiveConnection = cnnOra
                    .CommandText = "SP_TRANSITO_INSUPD"
                    .CommandType = adCmdStoredProc
                    
                    .Parameters("P_TIPO_MOVIMIENTO").Value = CStr(rsTransito("VCHA_TIPO_MOVIMIENTO").Value)
                    .Parameters("P_TRA_CONSECUTIVO").Value = Val(rsTransito("NUMB_TRA_CONSECUTIVO").Value)
                    .Parameters("P_TRA_NOTA_ENVIO").Value = CStr(rsTransito("VCHA_TRA_NOTA_ENVIO").Value)
                    .Parameters("P_TRA_PLANTA_DESTINO").Value = CStr(rsTransito("VCHA_TRA_PLANTA_DESTINO").Value)
                    .Parameters("P_TRA_PLANTA_ORIGEN").Value = CStr(rsTransito("VCHA_TRA_PLANTA_ORIGEN").Value)
                    .Parameters("P_TRA_ALMACEN_DESTINO").Value = CStr(rsTransito("VCHA_TRA_ALMACEN_DESTINO").Value)
                    .Parameters("P_TRA_ALMACEN_ORIGEN").Value = CStr(rsTransito("VCHA_TRA_ALMACEN_ORIGEN").Value)
                    .Parameters("P_EMP_EMPRESA_ORIGEN").Value = CStr(rsTransito("VCHA_EMP_EMPRESA_ORIGEN").Value)
                    .Parameters("P_EMP_EMPRESA_DESTINO").Value = CStr(rsTransito("VCHA_EMP_EMPRESA_DESTINO").Value)
                    .Parameters("P_UOR_UNIDAD_ORIGEN").Value = CStr(rsTransito("VCHA_UOR_UNIDAD_ORIGEN").Value)
                    .Parameters("P_UOR_UNIDAD_DESTINO").Value = CStr(rsTransito("VCHA_UOR_UNIDAD_DESTINO").Value)
                    .Parameters("P_MOV_MOVIMIENTO_ORIGEN").Value = CStr(rsTransito("VCHA_MOV_MOVIMIENTO_ORIGEN").Value)
                    .Parameters("P_MOV_MOVIMIENTO_DESTINO").Value = CStr(rsTransito("VCHA_MOV_MOVIMIENTO_DESTINO").Value)
                    .Parameters("P_SER_SERIE_ORIGEN").Value = CStr(rsTransito("VCHA_SER_SERIE_ORIGEN").Value)
                    .Parameters("P_SER_SERIE_DESTINO").Value = CStr(rsTransito("VCHA_SER_SERIE_DESTINO").Value)
                    .Parameters("P_ART_ARTICULO_ORIGEN").Value = rsTransito("VCHA_ART_ARTICULO_ORIGEN").Value
                    .Parameters("P_ART_ARTICULO_DESTINO").Value = CStr(rsTransito("VCHA_ART_ARTICULO_DESTINO").Value)
                    .Parameters("P_ART_DESCRIPCION").Value = CStr(rsTransito("VCHA_ART_DESCRIPCION").Value)
                    .Parameters("P_TRA_CANTIDAD_ENVIADA").Value = Val(rsTransito("NUMB_TRA_CANTIDAD_ENVIADA").Value)
                    .Parameters("P_TRA_CANTIDAD_RECIBIDA").Value = Val(rsTransito("NUMB_TRA_CANTIDAD_RECIBIDA").Value)
                    .Parameters("P_TRA_COSTO").Value = Val(Round(rsTransito("NUMB_TRA_COSTO").Value, 3))
                    .Parameters("P_TRA_CALIDAD").Value = CStr(rsTransito("VCHA_TRA_CALIDAD").Value)
                    .Parameters("P_TRA_SISTEMA_ORIGEN").Value = CStr(rsTransito("VCHA_TRA_SISTEMA_ORIGEN").Value)
                    .Parameters("P_TRA_SISTEMA_DESTINO").Value = CStr(rsTransito("VCHA_TRA_SISTEMA_DESTINO").Value)
                    .Parameters("P_TRA_STATUS_ID").Value = CStr(rsTransito("VCHA_TRA_STATUS_ID").Value)
                    .Parameters("P_TRA_USUARIO_ORIGEN").Value = CStr(rsTransito("VCHA_TRA_USUARIO_ORIGEN").Value)
                    .Parameters("P_TRA_USUARIO_DESTINO").Value = CStr(rsTransito("VCHA_TRA_USUARIO_DESTINO").Value)
                    .Parameters("P_TRA_MAQUINA_ORIGEN").Value = CStr(rsTransito("VCHA_TRA_MAQUINA_ORIGEN").Value)
                    .Parameters("P_TRA_MAQUINA_DESTINO").Value = CStr(rsTransito("VCHA_TRA_MAQUINA_DESTINO").Value)
                    .Parameters("P_TRA_REFERENCIA1").Value = CStr(rsTransito("VCHA_TRA_REFERENCIA1").Value)
                    .Parameters("P_TRA_REFERENCIA2").Value = CStr(rsTransito("VCHA_TRA_REFERENCIA2").Value)
                    .Parameters("P_COM_TIPO_PROVEEDOR").Value = CStr(rsTransito("CHAR_COM_TIPO_PROVEEDOR").Value)
                    .Parameters("P_TRA_TRANSPORTO").Value = CStr(rsTransito("VCHA_TRA_TRANSPORTO").Value)
                    .Parameters("P_TRA_CONTENEDOR_ID").Value = CStr(rsTransito("VCHA_TRA_CONTENEDOR_ID").Value)
                    .Parameters("P_TRA_PESO").Value = Val(rsTransito("NUMB_TRA_PESO").Value)
                    .Parameters("P_MON_MONEDA_ID").Value = CStr(rsTransito("VCHA_MON_MONEDA_ID").Value)
                    .Parameters("P_TRA_PRECIO").Value = Val(rsTransito("NUMB_TRA_PRECIO").Value)
                    .Parameters("P_DTR_FOLIO_RECEPCION").Value = CStr(rsTransito("numb_DTR_FOLIO_RECEPCION").Value)
                    .execute
                    'StrMensage = .Parameters("P_TRA_MENSAJE").Value
                    rsTransito.MoveNext
                    
                End With
                
            Wend
            'cnnOra.CommitTrans
            With cmdPoliza
                rsTransito.MoveFirst
                Set .ActiveConnection = cnnOra
                .CommandType = adCmdStoredProc
                .CommandText = "SP_TRANSITO_POLIZA"
                
                Set objParam = .CreateParameter("P_TRA_TIPO_MOVIMIENTO", adVarChar, adParamInputOutput, 30, rsTransito("VCHA_TIPO_MOVIMIENTO").Value)
                .Parameters.Append (objParam)
                Set objParam = .CreateParameter("P_TRA_NOTA_ENVIO", adVarChar, adParamInputOutput, 30, rsTransito("vcha_tra_nota_envio").Value)
                .Parameters.Append (objParam)
                Set objParam = .CreateParameter("P_TRA_PLANTA_ID", adVarChar, adParamInputOutput, 30, rsTransito("VCHA_TRA_PLANTA_ORIGEN").Value)
                .Parameters.Append (objParam)
                Set objParam = .CreateParameter("P_TRA_ALMACEN_ORIGEN", adVarChar, adParamInputOutput, 30, rsTransito("VCHA_TRA_ALMACEN_ORIGEN").Value)
                .Parameters.Append (objParam)
                Set objParam = .CreateParameter("P_TRA_ALMACEN_DESTINO", adVarChar, adParamInputOutput, 30, rsTransito("VCHA_TRA_ALMACEN_DESTINO").Value)
                .Parameters.Append (objParam)
                Set objParam = .CreateParameter("P_EMP_EMPRESA_ID", adVarChar, adParamInputOutput, 30, IIf(rsTransito("VCHA_TIPO_MOVIMIENTO").Value = "RECIBO", rsTransito("VCHA_EMP_EMPRESA_DESTINO").Value, rsTransito("VCHA_EMP_EMPRESA_ORIGEN").Value))
                .Parameters.Append (objParam)
                Set objParam = .CreateParameter("P_UOR_UNIDAD_ID", adVarChar, adParamInputOutput, 30, IIf(rsTransito("VCHA_TIPO_MOVIMIENTO").Value = "RECIBO", rsTransito("VCHA_UOR_UNIDAD_DESTINO").Value, rsTransito("VCHA_UOR_UNIDAD_ORIGEN").Value))
                .Parameters.Append (objParam)
                Set objParam = .CreateParameter("P_MOV_MOVIMIENTO_ID", adVarChar, adParamInputOutput, 30, rsTransito("VCHA_MOV_MOVIMIENTO_ORIGEN").Value)
                .Parameters.Append (objParam)
                Set objParam = .CreateParameter("P_TRA_STATUS_ID", adVarChar, adParamInputOutput, 30, rsTransito("VCHA_TRA_STATUS_ID").Value)
                .Parameters.Append (objParam)
                Set objParam = .CreateParameter("P_TRA_USUARIO", adVarChar, adParamInputOutput, 30, IIf(rsTransito("VCHA_TIPO_MOVIMIENTO").Value = "RECIBO", rsTransito("VCHA_TRA_USUARIO_DESTINO").Value, rsTransito("VCHA_TRA_USUARIO_ORIGEN").Value))
                .Parameters.Append (objParam)
                Set objParam = .CreateParameter("P_TRA_MAQUINA", adVarChar, adParamInputOutput, 30, IIf(rsTransito("VCHA_TIPO_MOVIMIENTO").Value = "RECIBO", rsTransito("VCHA_TRA_MAQUINA_DESTINO").Value, rsTransito("VCHA_TRA_MAQUINA_ORIGEN").Value))
                .Parameters.Append (objParam)
                Set objParam = .CreateParameter("P_DTR_FOLIO_RECEPCION", adVarChar, adParamInputOutput, 30, rsTransito("numb_DTR_FOLIO_RECEPCION").Value)
                .Parameters.Append (objParam)
                Set objParam = .CreateParameter("P_MESEGE", adVarChar, adParamOutput, 200, "F")
                .Parameters.Append (objParam)
                
                .execute
                StrMensage = IIf(IsNull(.Parameters("P_MESEGE").Value), "", .Parameters("P_MESEGE").Value)
            End With
            If StrMensage = "La poliza se agrego correctamente" Then
                cnnOra.CommitTrans
                MsgBox StrMensage, vbInformation, "Transito"
            Else
                GoTo errorTransito:
            End If
        Else
            MsgBox "No se encontró informacion para agregar a transito", vbCritical, "Transito"
        End If
        
    Else
        MsgBox "Error al conectar a la base de datos de Oracle", vbCritical, "Transito"
    End If


Exit Function
errorTransito:
fun_GuardaTransito = False
If rsTransito.State = 1 Then
    rsTransito.Close
End If

If cnnOra.State = 1 Then
    cnnOra.RollbackTrans
    cnnOra.Close
End If
    MsgBox StrMensage, vbCritical, "Transito"
End Function


Private Function conectaOracle(userIDOracle As String, dataSourceOracle As String, passOracle As String) As Boolean
    'Variables de bloque.
    Dim strConnectionStringOracle  As String
    
On Error GoTo errorConectarOracle
    conectaOracle = True
    If cnnOra.State = 1 Then
        cnnOra.Close
    End If
    
    'Establecer connection string para realizar la conexión a la base de
    'datos.
    strConnectionStringOracle = "Provider=OraOLEDB.Oracle.1;User ID=" & userIDOracle & "; " & _
                                "Data Source=" & dataSourceOracle & ";Extended Properties=; " & _
                                "Persist Security Info=True;Password=" & passOracle
   
    'MsgBox "Se esta conectando a pruebas , si ve este mensaje avise de inmediato al administrador", vbOKOnly
    
     'strConnectionStringOracle = "Provider=OraOLEDB.Oracle.1;User ID=interface; " & _
     '                           "Data Source=dbtest;Extended Properties=; " & _
     '                           "Persist Security Info=True;Password=interface"
   
    
    'Configurar objeto Connection.
    Set cnnOra = New ADODB.Connection
    
    cnnOra.CursorLocation = adUseClient
    cnnOra.CommandTimeout = 6000
    cnnOra.ConnectionString = strConnectionStringOracle
    'Abrir conexión a las base de datos.
    cnnOra.Open
    Exit Function
errorConectarOracle:
    conectaOracle = False
    MsgBox Error, vbCritical, "SIP"
End Function

