VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_DEPARTAMENTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String

Public ArrTB_DEPARTAMENTOS As Variant

Public cBint_dep_departamento_id As String
Public cVcha_dep_descripcion As String
Public cVcha_dep_responsable As String
Public cVcha_dep_status As String
Public cDtim_aud_fecha As String
Public cVcha_aud_usuario As String
Public cVcha_aud_maquina As String
Public cBint_pla_planta_id As String
Public cVcha_pla_descripc As String

'ANADIR UN REGISTRO
Public Function Anadir(cn As adodb.Connection, clBint_dep_departamento_id As String,clVcha_dep_descripcion As String,clVcha_dep_responsable As String,clVcha_dep_status As String,clDtim_aud_fecha As String,clVcha_aud_usuario As String,clVcha_aud_maquina As String,clBint_pla_planta_id As String,clVcha_pla_descripc As String) As Boolean
Dim rs As adodb.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
Cadena = ""
Cadena = Cadena + "Insert into TB_DEPARTAMENTOS("
Cadena = Cadena + "Bint_dep_departamento_id,"
Cadena = Cadena + "Vcha_dep_descripcion,"
Cadena = Cadena + "Vcha_dep_responsable,"
Cadena = Cadena + "Vcha_dep_status,"
Cadena = Cadena + "Dtim_aud_fecha,"
Cadena = Cadena + "Vcha_aud_usuario,"
Cadena = Cadena + "Vcha_aud_maquina,"
Cadena = Cadena + "Bint_pla_planta_id,"
Cadena = Cadena + "Vcha_pla_descripc"
Cadena = Cadena + ") values ("
Cadena = Cadena + " 0" + clBint_dep_departamento_id + ","
Cadena = Cadena + " '" + clVcha_dep_descripcion + "',"
Cadena = Cadena + " '" + clVcha_dep_responsable + "',"
Cadena = Cadena + " '" + clVcha_dep_status + "',"
Cadena = Cadena + " 0" + clDtim_aud_fecha + ","
Cadena = Cadena + " '" + clVcha_aud_usuario + "',"
Cadena = Cadena + " '" + clVcha_aud_maquina + "',"
Cadena = Cadena + " 0" + clBint_pla_planta_id + ","
Cadena = Cadena + " '" + clVcha_pla_descripc + "')
Set rs = cn.Execute(Cadena)
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Anadir = False
    GoTo SIGUE
End Function
'EDITAR UN REGISTRO
Public Function Modificar(cn As adodb.Connection, clBint_dep_departamento_id As String,clVcha_dep_descripcion As String,clVcha_dep_responsable As String,clVcha_dep_status As String,clDtim_aud_fecha As String,clVcha_aud_usuario As String,clVcha_aud_maquina As String,clBint_pla_planta_id As String,clVcha_pla_descripc As String) As Boolean
Dim rs As adodb.Recordset, Cadena As String
Modificar = True
On Error GoTo HELL
Cadena = ""
Cadena = Cadena + "Update TB_DEPARTAMENTOS set "
Cadena = Cadena + " Vcha_dep_descripcion='" + clVcha_dep_descripcion + "', "
Cadena = Cadena + " Vcha_dep_responsable='" + clVcha_dep_responsable + "', "
Cadena = Cadena + " Vcha_dep_status='" + clVcha_dep_status + "', "
Cadena = Cadena + " Dtim_aud_fecha= 0" + clDtim_aud_fecha + " , "
Cadena = Cadena + " Vcha_aud_usuario='" + clVcha_aud_usuario + "', "
Cadena = Cadena + " Vcha_aud_maquina='" + clVcha_aud_maquina + "', "
Cadena = Cadena + " Bint_pla_planta_id= 0" + clBint_pla_planta_id + " , "
Cadena = Cadena + " Vcha_pla_descripc='" + clVcha_pla_descripc + "'"
Cadena = Cadena + " where "
Cadena = Cadena + " Bint_dep_departamento_id= " + clBint_dep_departamento_id + "  "
Set rs = cn.Execute(Cadena)
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Modificar = False
    GoTo SIGUE
End Function
'ELIMINAR
Public Function Eliminar(cn As adodb.Connection, clBint_dep_departamento_id As String) As Boolean
Dim rs As adodb.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Cadena = ""
Cadena = Cadena + "Delete from  TB_DEPARTAMENTOS where "
Cadena = Cadena + " Bint_dep_departamento_id= " + clBint_dep_departamento_id + "  "
Set rs = cn.Execute(Cadena)
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function

'CARGAR REGISTROS
Public Function Cargar(cn As adodb.Connection, Optional CantidadaDevolver As Integer, Optional CamposOrden As String) As Boolean
Dim rs As adodb.Recordset, Cadena As String, Cuantos As Double
Dim Ejex As Double, EjeY As Double, Columna As adodb.Field
Cargar = True
On Error GoTo HELL
Cadena = ""
Cadena = Cadena + "select  "
Cadena = Cadena + IIf(CantidadaDevolver > 0, " top " + CStr(CantidadaDevolver) + " ", "")
Cadena = Cadena + " Bint_dep_departamento_id, "
Cadena = Cadena + " Vcha_dep_descripcion, "
Cadena = Cadena + " Vcha_dep_responsable, "
Cadena = Cadena + " Vcha_dep_status, "
Cadena = Cadena + " Dtim_aud_fecha, "
Cadena = Cadena + " Vcha_aud_usuario, "
Cadena = Cadena + " Vcha_aud_maquina, "
Cadena = Cadena + " Bint_pla_planta_id, "
Cadena = Cadena + " Vcha_pla_descripc"
Cadena = Cadena + " from  TB_DEPARTAMENTOS "
Cadena = Cadena + IIf(Len(CamposOrden) > 0, " order by " + CamposOrden, "")
Set rs = cn.Execute(Cadena)
If Not rs.EOF Then
 Cuantos = 0
 Do While Not rs.EOF
  Cuantos = Cuantos + 1
  rs.MoveNext
 Loop
rs.MoveFirst
 ReDim ArrTB_DEPARTAMENTOS(Cuantos, rs.Fields.Count)
 Ejex = 0
 Do While Not rs.EOF
  EjeY = 0
  For Each Columna In rs.Fields
   ArrTB_DEPARTAMENTOS(Ejex, EjeY) = Columna.Value
   EjeY = EjeY + 1
  Next
  EjeY = 0: Ejex = Ejex + 1
  rs.MoveNext
 Loop
End If
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Cargar = False
    GoTo SIGUE
End Function

'UBICAR UN REGISTRO
Public Function Ubicar(cn As adodb.Connection, clBint_dep_departamento_id As String) As Boolean
Dim rs As adodb.Recordset, Cadena As String
Ubicar = True
On Error GoTo HELL
Cadena = ""
Cadena = Cadena + "select  "
Cadena = Cadena + " Bint_dep_departamento_id, "
Cadena = Cadena + " Vcha_dep_descripcion, "
Cadena = Cadena + " Vcha_dep_responsable, "
Cadena = Cadena + " Vcha_dep_status, "
Cadena = Cadena + " Dtim_aud_fecha, "
Cadena = Cadena + " Vcha_aud_usuario, "
Cadena = Cadena + " Vcha_aud_maquina, "
Cadena = Cadena + " Bint_pla_planta_id, "
Cadena = Cadena + " Vcha_pla_descripc"
Cadena = Cadena + " from  TB_DEPARTAMENTOS where  "
Cadena = Cadena + " Bint_dep_departamento_id='" + clBint_dep_departamento_id + "' "
Set rs = cn.Execute(Cadena)
If Not rs.EOF Then
 cBint_dep_departamento_id = rs("Bint_dep_departamento_id").Value & ""
 cVcha_dep_descripcion = rs("Vcha_dep_descripcion").Value & ""
 cVcha_dep_responsable = rs("Vcha_dep_responsable").Value & ""
 cVcha_dep_status = rs("Vcha_dep_status").Value & ""
 cDtim_aud_fecha = rs("Dtim_aud_fecha").Value & ""
 cVcha_aud_usuario = rs("Vcha_aud_usuario").Value & ""
 cVcha_aud_maquina = rs("Vcha_aud_maquina").Value & ""
 cBint_pla_planta_id = rs("Bint_pla_planta_id").Value & ""
 cVcha_pla_descripc = rs("Vcha_pla_descripc").Value & ""
End If
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Ubicar = False
    GoTo SIGUE
End Function

'SIGUIENTE REGISTRO LLAVE
Public Function Siguiente(cn As adodb.Connection) As String
Dim rs As adodb.Recordset, Cadena As String
Siguiente = ""
On Error GoTo HELL
Cadena = ""
Cadena = Cadena + "Select max ("
Cadena = Cadena + " TB_DEPARTAMENTOS"
Cadena = Cadena + ")  from TB_DEPARTAMENTOS "
Set rs = cn.Execute(Cadena)
If Not rs.EOF Then
 Siguiente = rs(0)
End If
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    GoTo SIGUE
End Function

