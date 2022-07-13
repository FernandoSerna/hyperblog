VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBucarArticulos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtTexto 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   6135
      End
      Begin MSComctlLib.ListView lvwCatalogo 
         Height          =   2055
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Id"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lblCatalogo 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Catálogo ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   15
         TabIndex        =   6
         Top             =   105
         Width           =   6330
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1080
         TabIndex        =   5
         Top             =   3360
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Registros:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   3360
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Texto a buscar:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1110
      End
   End
End
Attribute VB_Name = "frmBucarArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variables de módulo.
Public sentenciaSql As String
Public comodin As String
Public valorSeleccionado1 As String
Public valorSeleccionado2 As String
Public valorSeleccionado3 As String
Public catalogo As String
Public cnnSID As New ADODB.Connection
Public servidor As String
Public bd As String
Public famAgrupador As String
Dim str_valorParaServidor As String
Dim str_valorParaBaseDatos As String
Dim str_valorParaTabla As String


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 'Revisar si se presionó la tecla Escape.
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    If Not Conectar_BDCompucaja(cnnSID, "compucaja", "srvtdacantia") Then
        MsgBox "Error al conectar a la base de datos Cantia." & Chr(13) & _
                "No se pudieron recuperar los datos Cantia", vbCritical, "S.I.P."
    Else
        ValorSeleccionado = ""
        lblCatalogo.Caption = " Catálogo " & catalogo
    End If

End Sub
Private Function Conectar_BDCompucaja(ByRef cnnCBD As ADODB.Connection, ByVal bd As String, ByVal servidor As String) As Boolean
    'Variables de bloque
    Dim strConnection_String As String
    
On Error GoTo Error_Conectar_BDS
    Conectar_BDCompucaja = True
    'Establecer connection strings para realizar las conexiones a las bases de
    'datos
    If cnnSID.State = 1 Then
        cnnSID.Close
    End If
    
    
    strConnection_String_SID = "Provider=SQLOLEDB.1;Password=compucaja" & _
                                ";Persist Security Info=True;User ID=sa" & _
                                ";Initial Catalog=" & UCase(bd) & ";Data Source=" & UCase(servidor)
    
    'Configurar objetos Connection
    'cnnCBD.CursorLocation = adUseClient
    If cnnCBD.State = 1 Then
        cnnCBD.Close
    End If
    cnnCBD.ConnectionString = strConnection_String_SID
    cnnCBD.CommandTimeout = 60
    cnnCBD.CursorLocation = adUseClient
    
    'Abrir conexiones a las bases de datos
    cnnCBD.Open
    Exit Function
Error_Conectar_BDS:
    Conectar_BDCompucaja = False
    MsgBox Err.Description, vbCritical, "SID"
End Function


'Procedimiento que busca en el catálogo el texto que
'se capturó.
Private Sub buscarTexto()
    'Variables de bloque.
    Dim strSentencia As String
    Dim rsBuscar As New ADODB.recordSet
    Dim itmCatalogo As ListItem
    Dim intI As Integer
    
On Error GoTo errorBuscarTexto
    strSentencia = Replace(sentenciaSql, comodin, txtTexto.Text)
    'Recuperar elementos del catálogo.
7    rsBuscar.Open strSentencia, cnnSID, adOpenDynamic, adLockOptimistic
    'Limpiar la lista de elementos del catálogo.
    lvwCatalogo.ListItems.Clear
    'Mostrar el número de elementos recuperados.
    lblRegistros.Caption = rsBuscar.RecordCount
    'Mostrar elementos recuperados.
    If rsBuscar.RecordCount > 0 Then
        lvwCatalogo.ColumnHeaders.Clear
        For intI = 0 To rsBuscar.Fields.Count - 1
            lvwCatalogo.ColumnHeaders.Add , , rsBuscar.Fields(intI).Name
        Next intI
    End If
    While Not rsBuscar.EOF
        Set itmCatalogo = lvwCatalogo.ListItems.Add(, , rsBuscar(0).Value)
        For intI = 1 To rsBuscar.Fields.Count - 1
            itmCatalogo.SubItems(intI) = IIf(IsNull(rsBuscar(intI).Value), "NULL", rsBuscar(intI).Value)
        Next intI
        rsBuscar.MoveNext
    Wend
    lvwCatalogo.ColumnHeaders(1).Width = 1000
    lvwCatalogo.ColumnHeaders(2).Width = 3500
    'Liberar espacio de memoria ocupado por objetos instanciados.
    If rsBuscar.RecordCount <> 0 Then
        lvwCatalogo.SetFocus
        lvwCatalogo.ListItems(1).Selected = True
    End If
    Set rsBuscar = Nothing
    
    
    
    Exit Sub
errorBuscarTexto:
    MsgBox "Error al filtrar catálogo." & vbCrLf & Err.Source & " " & Err.Number & " " & Err.Description, vbCritical, "S.I.P."
End Sub
'Procedimiento que asigna los valores del elemento de catálogo
'seleccionado.
Private Sub asignarValores()
    'Revisar si se seleccionó un elemento de la lista.
    If Not lvwCatalogo.selectedItem Is Nothing Then
        valorSeleccionado1 = lvwCatalogo.selectedItem.Text
        If lvwCatalogo.ColumnHeaders.Count > 1 Then
            valorSeleccionado2 = lvwCatalogo.selectedItem.SubItems(1)
            If lvwCatalogo.ColumnHeaders.Count > 2 Then
                valorSeleccionado3 = lvwCatalogo.selectedItem.SubItems(2)
            End If
        End If
        Unload Me
    End If
End Sub

Private Sub lvwCatalogo_DblClick()
    'Asignar valores del elemento de catálogo seleccionado.
    asignarValores
End Sub

Private Sub lvwCatalogo_KeyPress(KeyAscii As Integer)
    'Revisar si se presionó Enter.
    If KeyAscii = 13 Then
        asignarValores
    End If
End Sub


Private Sub txtTexto_KeyPress(KeyAscii As Integer)
    'Revisar si se presionó Enter.
    If KeyAscii = 13 Then
        'Buscar en el catálogo el texto capturado.
        buscarTexto
    End If
    
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub


