VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEquivalenciasCoppel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Equivalencias - Coppel"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ilstCoppel 
      Left            =   0
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquivalenciasCoppel.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquivalenciasCoppel.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquivalenciasCoppel.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquivalenciasCoppel.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquivalenciasCoppel.frx":0178
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7920
      TabIndex        =   13
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   6480
      TabIndex        =   12
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   9135
      Begin VB.CheckBox chkTodos 
         Caption         =   "Seleccionar todos"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   2295
      End
      Begin MSComctlLib.ListView lvLineas 
         Height          =   1575
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Anterior"
            Object.Width           =   3122
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nuevo"
            Object.Width           =   2858
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Artículo"
            Object.Width           =   2223
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripción"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cantidad"
            Object.Width           =   1508
         EndProperty
      End
      Begin VB.Label Label7 
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   8655
      End
      Begin VB.Label Label3 
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   8895
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Caption         =   " Paso 3: Revisar equivalencias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   15
         TabIndex        =   9
         Top             =   105
         Width           =   9105
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   5280
      TabIndex        =   4
      Top             =   120
      Width           =   3975
      Begin VB.FileListBox flstArchivos 
         Height          =   1455
         Left            =   120
         Pattern         =   "*.ZIP;*.DBF"
         TabIndex        =   6
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "Archivos para la impresión de etiquetas."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         Caption         =   " Paso 2: Seleccionar archivo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   15
         TabIndex        =   5
         Top             =   105
         Width           =   3945
      End
   End
   Begin VB.Frame fmeBuscarCarpeta 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.DriveListBox dveCoppel 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   4695
      End
      Begin VB.DirListBox dirCoppel 
         Height          =   990
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   4695
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccione la carpeta que contiene los archivos con la información para la impresión de etiquetas."
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   " Paso 1: Buscar carpeta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   15
         TabIndex        =   2
         Top             =   105
         Width           =   4905
      End
   End
End
Attribute VB_Name = "frmEquivalenciasCoppel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variables de bloque.
Private cnnSid As ADODB.Connection

Private Sub chkTodos_Click()
    If chkTodos.Value = 1 Then
        SeleccionarTodos (True)
    Else
        SeleccionarTodos (False)
    End If
End Sub

Private Sub cmdActualizar_Click()
    actualizarEquivalencias
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub dirCoppel_Change()
On Error GoTo errorDirCoppel_Change
    'Asignar a la lista de archivos el directorio donde se encuentran
    'los archivos.
    flstArchivos.Path = dirCoppel.Path
    Exit Sub
errorDirCoppel_Change:
    MsgBox "Error al asignar la lista de archivos el directorio donde se encuentran los archivos." & vbCrLf & _
            Err.Source & " " & Err.Number & " " & Err.Description, vbCritical, "S.I.D."
End Sub

Private Sub dveCoppel_Change()
On Error GoTo errorDveCoppel_Change
    'Asignar el drive seleccionado.
    Me.dirCoppel.Path = dveCoppel.Drive
    Exit Sub
errorDveCoppel_Change:
    MsgBox "Error al asignar el drive seleccionado. " & _
            Err.Source & " " & Err.Number & " " & Err.Description, vbCritical, "S.I.D."
End Sub

Private Sub flstArchivos_Click()
    'Cargar las líneas contenidas en el archivo.
    If Not cargarLineasArchivo Then
        lvLineas.ListItems.Clear
    End If
End Sub

Private Function cargarLineasArchivo() As Boolean
    'Variables de bloque.
    Dim strQuery As String
    Dim cnnArchivo As New ADODB.Connection
    Dim rsArchivo As New ADODB.recordSet
    Dim itmLinea As ListItem
    Dim intResultado As String
    Dim strEquivalencia As String
    Dim strCodigo As String
    Dim strDescripcion As String
    
    
On Error GoTo errorCargarLineasArchivo
    cargarLineasArchivo = True
    chkTodos.Value = 0
    cnnArchivo.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=DSN=dBASE Files;DBQ=" & dirCoppel.Path & ";DefaultDir=" & dirCoppel.Path & ";DriverId=533;MaxBufferSize=2048;PageTimeout=5;"
    strQuery = "SELECT * FROM " & flstArchivos.FileName
    rsArchivo.Open strQuery, cnnArchivo, adOpenDynamic, adLockOptimistic
    'Cargar lineas del archivo en la lista de las líneas de archivo.
    lvLineas.ListItems.Clear
    'Abrir conexión a la base de datos.
    Set cnnSid = New ADODB.Connection
    If Not conectarBD(cnnSid) Then
        GoTo finCargarLineasArchivo
    End If
    While Not rsArchivo.EOF
        'Consultar en el almacén de datos el código para la equivalencia
        'contenida en el archivo.
        strEquivalencia = rsArchivo("Dato_4").Value
        strCodigo = ""
        strDescripcion = ""
        intResultado = buscarEquivalencia(strEquivalencia, strCodigo, strDescripcion)
        'Revisar el resultado de buscar el código para la equivalencia
        'contenida en el archivo.
        Select Case intResultado
            Case -1
                lvLineas.ListItems.Clear
                GoTo finCargarLineasArchivo
            Case 0
                strEquivalencia = Mid(rsArchivo("Dato_4").Value, 1, 10) & "%" & Mid(rsArchivo("Dato_4").Value, 13, 16)
                strCodigo = ""
                strDescripcion = ""
                'Consultar en el almacén de datos el código para la equivalencia
                'contenida en el archivo sin importar el valor del caracter
                'en la posición 12.
                intResultado = buscarEquivalencia(strEquivalencia, strCodigo, strDescripcion)
                'Revisar el resultado de buscar el código para la equivalencia
                'contenida en el archivo.
                Select Case intResultado
                    Case -1
                        lvLineas.ListItems.Clear
                        GoTo finCargarLineasArchivo
                    Case 0
                        strEquivalencia = ""
                End Select
        End Select
        
        Set itmLinea = lvLineas.ListItems.Add(, , strEquivalencia) 'Código anterior.
        itmLinea.SubItems(1) = rsArchivo("Dato_4").Value 'Código contenido en el archivo.
        itmLinea.SubItems(2) = strCodigo 'Código del artículo.
        itmLinea.SubItems(3) = strDescripcion 'Descripción.
        itmLinea.SubItems(4) = rsArchivo("Cantidad") 'Cantidad en el archivo.
        

        rsArchivo.MoveNext
    Wend
    'Dar formato a la lista.
    lvLineas.Refresh
    If lvLineas.ListItems.Count < 5 Then
        lvLineas.ColumnHeaders(4).Width = 3300.095
    Else
        lvLineas.ColumnHeaders(4).Width = 3060.284
    End If
finCargarLineasArchivo:
    'Liberar espacio de memoria ocupado por el objeto Connection
    If cnnSid.State = 1 Then
        cnnSid.Close
    End If
    cnnArchivo.Close
    Set cnnSid = Nothing
    Set cnnArchivo = Nothing
    Set rsArchivo = Nothing
    Exit Function
errorCargarLineasArchivo:
    cargarLineasArchivo = False
    MsgBox "Error al abrir el archivo." & vbCrLf & _
            Err.Source & " " & Err.Number & " " & Err.Description, vbCritical, "S.I.D."
End Function

Private Sub Form_Load()
    Top = 700
    Left = 1300
    Me.dirCoppel.Path = "C:\"
End Sub

Private Sub SeleccionarTodos(blnValor As Boolean)
    'Variables de bloque.
    Dim intI As Integer
    
    For intI = 1 To lvLineas.ListItems.Count
        lvLineas.ListItems(intI).Checked = blnValor
    Next intI
End Sub

'Función para conectarse a la base de datos de DISTRIBUCION.
Private Function conectarBD(ByRef cnnCBD As ADODB.Connection) As Boolean
    'Variables de bloque
    Dim strConnectionString As String
    
On Error GoTo errorConectarBD
    conectarBD = True
    'Establecer connection string para realizar las conexiones a la base de
    'datos.
    strConnectionStringSID = "Provider=SQLOLEDB.1;Password=ELIA" & _
                               ";Persist Security Info=True;User ID=SA" & _
                               ";Initial Catalog=VIANNEY;Data Source=DISTRIBUCION"
    'Configurar objetos Connection
    cnnCBD.CursorLocation = adUseClient
    cnnCBD.CommandTimeout = 60
    cnnCBD.ConnectionString = strConnectionStringSID
    'Abrir conexiones a las bases de datos
    cnnCBD.Open
    Exit Function
errorConectarBD:
    conectarBD = False
    MsgBox "Error al abrir la conexión con el servidor de base de datos." & vbCrLf & _
            Err.Source & " " & Err.Number & " " & Err.Description, vbCritical, "S.I.D."
End Function

'Función para buscar en el almacen de datos una equivalencia para un código.
Public Function buscarEquivalencia(ByRef strEquivalencia As String, ByRef strCodigo As String, ByRef strDesc As String) As Integer
    'Variables de bloque.
    Dim rsEq As New ADODB.recordSet
    
On Error GoTo errorBuscarEquivalencia
    'Obtener la equivalencia de un código.
    strQuery = "SELECT e.vcha_equ_codigo_equivalente equivalencia, a.vcha_art_articulo_id articuloId, a.vcha_art_nombre_español descripcion " & _
                    "FROM tb_equivalencias e, tb_articulos a " & _
                    "WHERE a.vcha_art_articulo_id = e.vcha_art_articulo_id " & _
                    "AND e.vcha_equ_codigo_equivalente LIKE '" & strEquivalencia & "' " & _
                    "AND LEN(e.vcha_equ_codigo_equivalente) = 16"
    rsEq.Open strQuery, cnnSid, adOpenDynamic, adLockOptimistic
    'Revisar si se obtuvo alguna equivalencia.
    If rsEq.RecordCount > 0 Then
        'Establecer el código y su descripción.
        buscarEquivalencia = 1
        strEquivalencia = rsEq("equivalencia").Value
        strCodigo = rsEq("articuloId").Value
        strDesc = rsEq("descripcion").Value
    Else
        buscarEquivalencia = 0
    End If
finBuscarEquivalencia:
    'Liberar espacio de memoria ocupado por los objetos instanciados.
    Set rsEq = Nothing
    Exit Function
errorBuscarEquivalencia:
    buscarEquivalencia = -1
    MsgBox "Error al buscar el código correspondiente a la equivalencia indicada." & vbCrLf & _
            Err.Source & " " & Err.Number & " " & Err.Description, vbCritical, "S.I.D."
End Function

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lvLineas_DblClick()
   buscarArticulo
End Sub
'Procedimiento que muestra la forma para buscar un artículo.
Private Sub buscarArticulo()
    'Variables de bloque.
    Dim frmBuscar As New frmBuscarArticulo
On Error GoTo errorBuscarArticulo
    'Revisar si se seleccionó un elemento de la lista.
    If lvLineas.selectedItem Is Nothing Then GoTo finBuscarArticulo
    
    With frmBuscar
        'Mostrar forma de búsqueda.
        .blnSeleccionArticulo = False
        .Show vbModal, Me
        'Revisar si se seleccionó algún artículo.
        If .blnSeleccionArticulo Then
            lvLineas.selectedItem.SubItems(2) = .strId
            lvLineas.selectedItem.SubItems(3) = .strDescripcion
        End If
    End With
finBuscarArticulo:
    'Liberar espacio de memoria ocupado por los objetos instanciados.
    Set frmBuscar = Nothing
    Exit Sub
errorBuscarArticulo:
    MsgBox "Error al buscar el artículo." & vbCrLf & _
            Err.Source & " " & Err.Number & " " & Err.Description, vbCritical, "S.I.D."
End Sub
'Procedimiento que actualiza las equivalencias.
Private Sub actualizarEquivalencias()
    'Variables de bloque.
    Dim cmmComandoSql As New ADODB.Command
    Dim itmElemento As ListItem
    
On Error GoTo errorActualizarEquivalencias
    'Abrir conexión a la base de datos.
    Set cnnSid = New ADODB.Connection
    If Not conectarBD(cnnSid) Then GoTo finActualizarEquivalencias
    'Configurar comando.
    cmmComandoSql.ActiveConnection = cnnSid
    cmmComandoSql.CommandType = adCmdStoredProc
    cmmComandoSql.CommandText = "spGuardarEquivalenciaCoppel"
    'Iniciar transacción en la base de datos.
    cnnSid.BeginTrans
    'Recorrer cada elemento de la lista.
    For Each itmElemento In lvLineas.ListItems
        'Revisar si el elemento fue seleccionado.
        If itmElemento.Checked Then
            cmmComandoSql("@vcha_anterior").Value = itmElemento
            cmmComandoSql("@vcha_nuevo").Value = itmElemento.SubItems(1)
            cmmComandoSql("@vcha_articulo").Value = itmElemento.SubItems(2)
            cmmComandoSql.execute
        End If
    Next
    'Terminar transacción en la base de datos.
    cnnSid.CommitTrans
    MsgBox "Se actualizaron correctamente las equivalencias.", vbInformation, "S.I.D."
finActualizarEquivalencias:
    'Liberar espacio de memoria ocupado por los objetos instanciados.
    If cnnSid.State = 1 Then
        cnnSid.Close
    End If
    Set cnnSid = Nothing
    Set cmmComandoSql = Nothing
    Exit Sub
errorActualizarEquivalencias:
    MsgBox "Error al actualizar las equivalencias." & vbCrLf & _
            Err.Source & " " & Err.Number & " " & Err.Description, vbCritical, "S.I.D."
End Sub
