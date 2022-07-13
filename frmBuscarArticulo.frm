VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBuscarArticulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar artículo"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6150
   Begin MSComctlLib.StatusBar stbrArticulos 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   4095
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11007
            MinWidth        =   11007
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtBuscar 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   5655
      End
      Begin MSComctlLib.ListView lvArticulos 
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Id"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7761
         EndProperty
      End
      Begin VB.Label lblRegistros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         Height          =   195
         Left            =   4920
         TabIndex        =   7
         Top             =   3480
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Registros:"
         Height          =   195
         Left            =   4080
         TabIndex        =   6
         Top             =   3480
         Width           =   705
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Catálogo de artículos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   15
         TabIndex        =   4
         Top             =   105
         Width           =   5850
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buscar:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmBuscarArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variables de bloque
Public blnSeleccionArticulo As Boolean
Public strId As String
Public strDescripcion As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Revisar si se presionó la tecla Escape
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub lvArticulos_DblClick()
    seleccionarArticulo
End Sub

Private Sub lvArticulos_KeyPress(KeyAscii As Integer)
    'Revisar si se presionó la tecla Enter.
    If KeyAscii = 13 Then
        seleccionarArticulo
    End If
End Sub

Private Sub txtBuscar_GotFocus()
    stbrArticulos.Panels(1).Text = "Presione Enter para buscar."
End Sub

Private Sub txtBuscar_LostFocus()
    stbrArticulos.Panels(1).Text = ""
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    'Revisar si se presionó la tecla Enter.
    If KeyAscii = 13 And txtBuscar.Text <> "" Then
        If Not llenarListaArticulos(txtBuscar.Text) Then
            lvArticulos.ListItems.Clear
        End If
    End If
End Sub
'Función que llena la lista de artículos.
Private Function llenarListaArticulos(strCadena As String) As Boolean
    'Variables de bloque.
    Dim strQuery As String
    Dim cnnArticulos As New ADODB.Connection
    Dim rsArticulos As New ADODB.recordSet
    Dim itmArticulo As ListItem

On Error GoTo errorLlenarListaArticulos
    llenarListaArticulos = True
    'Abrir conección con la base de datos.
    If Not conectarBD(cnnArticulos) Then GoTo finLlenarListaArticulos
    'Recuperar los artículos.
    strQuery = "SELECT vcha_art_articulo_id, vcha_art_nombre_español " & _
                "FROM tb_articulos " & _
                "WHERE (vcha_art_articulo_id LIKE '" & strCadena & "%' OR vcha_art_nombre_español LIKE '%" & Replace(strCadena, " ", "%") & "%')"
    rsArticulos.Open strQuery, cnnArticulos, adOpenDynamic, adLockOptimistic
    'Llenar lista con los artículos recuperados.
    lvArticulos.ListItems.Clear
    lblRegistros.Caption = rsArticulos.RecordCount
    While Not rsArticulos.EOF
        Set itmArticulo = lvArticulos.ListItems.Add(, , rsArticulos("vcha_art_articulo_id").Value)
        itmArticulo.SubItems(1) = rsArticulos("vcha_art_nombre_español").Value
        rsArticulos.MoveNext
    Wend
    rsArticulos.Close
finLlenarListaArticulos:
    'Liberar memoria ocupada por los objetos instanciados.
    If cnnArticulos.State = 1 Then
        cnnArticulos.Close
    End If
    Set cnnArticulos = Nothing
    Set rsArticulos = Nothing
    Exit Function
errorLlenarListaArticulos:
    llenarListaArticulos = False
    MsgBox "Error al llenar la lista de artículos." & vbCrLf & _
            Err.Source & " " & Err.Number & " " & Err.Description, vbCritical, "S.I.D."
End Function
'Procedimiento que carga el artículo seleccionado.
Private Sub seleccionarArticulo()
    'Revisar si se seleccionó algún artículo.
    If lvArticulos.selectedItem Is Nothing Then
    Else
        blnSeleccionArticulo = True
        strId = lvArticulos.selectedItem
        strDescripcion = lvArticulos.selectedItem.SubItems(1)
        Unload Me
    End If
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

