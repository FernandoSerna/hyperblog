VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "aClsNodoArbolTrazabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Propiedades.
Public articuloAcatex As String
Public articuloDescripcion As String
Public articuloIdentificador As String
Public eventoTipo As String
Public eventoNumero As String
Public informacionNodoIdentificador As String
Public informacionNodoPadreIdentificador As String
Public loteCantidad As String
Public loteNumero As String
Public MensajeError As String
Public nodoPadreTipo As String
Public nodoTipo As String
Public organizacion As String
Public recepcionFacturas As String
Public recepcionNumero As String
Public recepcionPaisOrigen As String
Public recepcionPedimento As String
Public recepcionProveedor As String


'Función que genera el identificador de la información del nodo.
Public Function generarIdentificadorInformacionNodo(conexion As ADODB.Connection) As Boolean
    'Variables de bloque.
    Dim strSentenciaSql As String
    Dim rsConsultaSql As New ADODB.recordSet

On Error GoTo errorGenerarIdentificadorInformacionNodo
    'Generar el identificador de la información del nodo.
    strSentenciaSql = "SELECT proSqInformacionNodo.NEXTVAL identificador FROM dual"
    rsConsultaSql.Open strSentenciaSql, conexion, adOpenDynamic, adLockReadOnly
    informacionNodoIdentificador = rsConsultaSql("identificador").Value
    rsConsultaSql.Close
    'Liberar espacio de memoria ocupado por los objetos instanciados.
    Set rsConsultaSql = Nothing
    
    generarIdentificadorInformacionNodo = True
    Exit Function
errorGenerarIdentificadorInformacionNodo:
    generarIdentificadorInformacionNodo = False
    MensajeError = "Imposible generar el identificador de la información del nodo." & vbCrLf & _
                    Err.Source & " " & Err.Number & " " & Err.Description
End Function
'Función que registra la información del nodo del árbol de trazabilidad.
Public Function registrarInformacionNodo(conexion As ADODB.Connection) As Boolean
    'Variables de bloque.
    Dim cmmComandoSql As New ADODB.Command
    
On Error GoTo errorRegistrarInformacionNodo
    cmmComandoSql.ActiveConnection = conexion
    cmmComandoSql.CommandType = adCmdText
    cmmComandoSql.CommandTimeout = 4
    'Revisar el tipo de nodo.
    Select Case nodoTipo
        Case "L"
            cmmComandoSql.CommandText = "CALL proPkgTrazabilidad.spLoteInsertar(" & informacionNodoIdentificador & "," & _
                                        "'" & organizacion & "','" & loteNumero & "','" & articuloIdentificador & "'," & _
                                        "'" & articuloDescripcion & "'," & loteCantidad & ", '" & articuloAcatex & "')"
        Case "E"
            cmmComandoSql.CommandText = "CALL proPkgTrazabilidad.spEventoInsertar(" & informacionNodoIdentificador & "," & _
                                        "'" & organizacion & "','" & eventoTipo & "','" & eventoNumero & "')"
        Case "R"
            cmmComandoSql.CommandText = "CALL proPkgTrazabilidad.spRecepcionInsertar(" & informacionNodoIdentificador & "," & _
                                        "'" & organizacion & "','" & recepcionNumero & "','" & recepcionProveedor & "'" & _
                                        "'" & recepcionFacturas & "','" & recepcionPedimento & "', '" & recepcionPaisOrigen & "')"
        Case Else
            Err.Raise 1, "S.I.P.", "Tipo de nodo inválido"
    End Select
    'Registrar la información del nodo del árbol de trazabilidad.
    cmmComandoSql.execute
    
    'Liberar espacio de memoria ocupado por los objetos instanciados.
    Set cmmComandoSql = Nothing
    
    registrarInformacionNodo = True
    Exit Function
errorRegistrarInformacionNodo:
    registrarInformacionNodo = False
    MensajeError = "Imposible registrar la información del nodo." & vbCrLf & _
                    Err.Source & " " & Err.Number & " " & Err.Description
End Function
'Función que registra el nodo del árbol de trazabilidad.
Public Function registrarNodo(conexion As ADODB.Connection) As Boolean
    'Variables de bloque.
    Dim cmmComandoSql As New ADODB.Command
    
On Error GoTo errorRegistrarNodo
    cmmComandoSql.ActiveConnection = conexion
    cmmComandoSql.CommandType = adCmdText
    cmmComandoSql.CommandTimeout = 4
    'Registrar el nodo del árbol de trazabilidad.
    cmmComandoSql.CommandText = "CALL proPkgTrazabilidad.spNodoInsertar('" & nodoTipo & "'," & informacionNodoIdentificador & "," & _
                                "'" & nodoPadreTipo & "'," & informacionNodoPadreIdentificador & ")"
    cmmComandoSql.execute
    'Liberar espacio de memoria ocupado por los objetos instanciados.
    Set cmmComandoSql = Nothing
    
    registrarNodo = True
    Exit Function
errorRegistrarNodo:
    registrarNodo = False
    MensajeError = "Imposible registrar el nodo." & vbCrLf & _
                    Err.Source & " " & Err.Number & " " & Err.Description
End Function

'Función que cancela una factura.
Public Function cancelarFactura(conexion As ADODB.Connection) As Boolean
    'Variables de bloque.
    Dim cmmComandoSql As New ADODB.Command
    
On Error GoTo errorRegistrarNodo
    cmmComandoSql.ActiveConnection = conexion
    cmmComandoSql.CommandType = adCmdText
    cmmComandoSql.CommandTimeout = 4
    'Registrar el nodo del árbol de trazabilidad.
    cmmComandoSql.CommandText = "CALL proPkgTrazabilidad.spFacturaCancelar('" & organizacion & "','" & eventoNumero & "')"
    cmmComandoSql.execute
    'Liberar espacio de memoria ocupado por los objetos instanciados.
    Set cmmComandoSql = Nothing
    
    cancelarFactura = True
    Exit Function
errorRegistrarNodo:
    cancelarFactura = False
    MensajeError = "Imposible cancelar la factura." & vbCrLf & _
                    Err.Source & " " & Err.Number & " " & Err.Description
End Function

