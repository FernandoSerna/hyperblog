VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEtiquetasPreciosCantia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etiquetas Precios"
   ClientHeight    =   5460
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmEtiquetasPreciosCantia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1500
      Picture         =   "frmEtiquetasPreciosCantia.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Configurar Impresora"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7395
      Picture         =   "frmEtiquetasPreciosCantia.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_ConfiguraEtiquetas 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1160
      Picture         =   "frmEtiquetasPreciosCantia.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Configura y Agrega Tama?o Etiqueta"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   470
      Picture         =   "frmEtiquetasPreciosCantia.frx":0B48
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_cargar_Excel 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmEtiquetasPreciosCantia.frx":0C4A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Cargar Desde Excel los Articulos"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   300
      Width           =   7620
   End
   Begin VB.Frame frmArticulo 
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   390
      Width           =   7650
      Begin VB.Frame Frame1 
         Height          =   120
         Index           =   0
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   7455
      End
      Begin VB.TextBox txt_codigo 
         Height          =   285
         Left            =   840
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   285
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   5565
      End
      Begin VB.TextBox txt_Cantidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cbo_TiposEtiquetas 
         Height          =   315
         ItemData        =   "frmEtiquetasPreciosCantia.frx":0D4C
         Left            =   840
         List            =   "frmEtiquetasPreciosCantia.frx":0D4E
         TabIndex        =   10
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lbl_Cantidad 
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lbl_enc_Tipo 
         Caption         =   "Tama?o:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lbl_Codigo 
         Caption         =   "Articulo:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame fra_catalago 
      Caption         =   "Carga Articulos"
      Height          =   1935
      Left            =   1320
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Height          =   255
         Left            =   5160
         TabIndex        =   5
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmd_cargar_batch 
         Caption         =   "&Ejecutar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmd_abrir 
         Height          =   300
         Left            =   4800
         Picture         =   "frmEtiquetasPreciosCantia.frx":0D50
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txt_archivo 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   4455
      End
      Begin VB.CommandButton cmd_cancelar_batch 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   1320
         Width           =   1335
      End
      Begin MSComctlLib.ProgressBar PB_carga 
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComDlg.CommonDialog dia_abrir_archivo 
      Left            =   7200
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lv_articulos 
      Height          =   3195
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   5636
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "sku"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   9701
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "color"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "precio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "via_medida"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cantidad"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "subLinea"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "talla"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "dise?o"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "largo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "alto"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "tamano"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pb_AvanceImpresion 
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   75
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Menu OpcionesEtiquetas 
      Caption         =   "OpcionesEtiquetas"
      Visible         =   0   'False
      Begin VB.Menu mnuEliminar 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "frmEtiquetasPreciosCantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conCompucaja As New ADODB.Connection
Dim rsConfiguracion As New ADODB.recordSet

Private Sub cmd_abrir_Click()
    With dia_abrir_archivo
        .DialogTitle = "Archivo Catalogo para Cargar"
        .Filter = "(*.XLS)|*.XLS"
        .ShowOpen
    End With
    If dia_abrir_archivo.FileName <> "" Then
        txt_archivo = dia_abrir_archivo.FileName
        'If txt_archivo.Text <> "" Then
        '    fra_opciones.Enabled = True
        'End If
        cmd_cargar_batch.Enabled = True
    End If

End Sub

Private Sub cmd_cancelar_batch_Click()
    fra_catalago.Visible = False
    txt_archivo.Text = ""
End Sub

Private Sub cmd_cargar_batch_Click()
    Dim conArchivoSID As String
    Dim i As Integer
    
    'Dim rsCuante As New ADODB.recordSet
    Dim strValida, strcompara As String
    
    strValida = ""
    conArchivoSID = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & txt_archivo.Text
    
    rs.Open "SELECT codigo, cantidad FROM [impresion$] where codigo is not null ", conArchivoSID
 
    If rs.RecordCount <> 0 Then
        
        PB_carga.Visible = True
        PB_carga.Value = 0
        PB_carga.Max = rs.RecordCount
    
        i = 1
        lv_articulos.ListItems.Clear
        If Conectar_BDCompucaja(conCompucaja, "compucaja", "srvtdacantia") Then
            While Not rs.EOF
                If Not IsNull(rs("Codigo").Value) And Not IsNull(rs("cantidad").Value) Then
                    strcompara = pro_AgregaFila(rs("Codigo").Value, rs("Cantidad").Value)
                    If strcompara <> "" Then
                        strValida = strValida & vbCrLf & strcompara
                    End If
                    rs.MoveNext
                    PB_carga.Value = i
                    i = i + 1
                End If
            Wend
        Else
            MsgBox "Error al conectar con el servidor de Cantia", vbCritical, "SID"
        End If
        conCompucaja.Close
    Else
        MsgBox "El Archivo no tiene Informacion", vbExclamation, "SIP"
    End If
    If rs.State = 1 Then
        rs.Close
    End If
    If strValida <> "" Then
        MsgBox "Los siguientes articulo no se encontraron" & vbCrLf & strValida, vbCritical, "SID"
    End If
    fra_catalago.Visible = False
    cmd_cargar_batch.Enabled = False
    txt_archivo.Text = ""

End Sub

Private Sub cmd_cargar_pedido_Click()
    If fra_catalago.Visible = False Then
        fra_catalago.Visible = True
    Else
        fra_catalago.Visible = False
    End If
End Sub

Private Sub cmd_cargar_Excel_Click()
    If fra_catalago.Visible = False Then
        fra_catalago.Visible = True
    Else
        fra_catalago.Visible = False
    End If
End Sub

Private Sub cmd_ConfiguraEtiquetas_Click()
    frmConfiguracionTama?os.Show 1
    Dim rsEtiquetas As New ADODB.recordSet
    If Conectar_BDCompucaja(conCompucaja, "compucaja", "srvtdacantia") Then
        rsEtiquetas.Open "Select inte_eti_etiqueta_id, vcha_eti_etiqueta " & _
                        "From tb_etiquetasConfiguracion with(nolock) ", _
                    conCompucaja, _
                    adOpenDynamic, _
                    adLockOptimistic
        Call RecsetToCombo(cbo_TiposEtiquetas.hwnd, rsEtiquetas, 1)
        rsEtiquetas.Close
    Else
        MsgBox "Error al conectar con el servidor", vbCritical, "SID"
    End If
    If conCompucaja.State = 1 Then conCompucaja.Close
End Sub

Private Sub cmd_imprimir_Click()
    Dim x As Double
    If lv_articulos.ListItems.Count > 0 Then
        If cbo_TiposEtiquetas.Text <> "" Then
            
                If MsgBox("El Tama?o de papel que debe ir en la impresora es: " & cbo_TiposEtiquetas.Text & vbCrLf & "?Es correcto?", vbYesNo, "SID") = vbYes Then
                    
                        If Conectar_BDCompucaja(conCompucaja, "compucaja", "srvtdaCantia") Then
                            Dim fila As Integer
                            rsConfiguracion.Open "select * " & _
                                    "from tb_etiquetasConfiguracion with(nolock)  " & _
                                    "where  vcha_eti_etiqueta = '" & cbo_TiposEtiquetas.Text & "'", _
                                conCompucaja, _
                                adOpenDynamic, _
                                adLockOptimistic
                            If rsConfiguracion.RecordCount > 0 Then
                                pb_AvanceImpresion.Visible = True
                                pb_AvanceImpresion.Value = 0
                                pb_AvanceImpresion.Max = lv_articulos.ListItems.Count
                                
                                For fila = 1 To lv_articulos.ListItems.Count
                                    Call pro_generaEtiqueta(lv_articulos.ListItems(fila), _
                                                            lv_articulos.ListItems(fila).SubItems(1), _
                                                            lv_articulos.ListItems(fila).SubItems(5), _
                                                            lv_articulos.ListItems(fila).SubItems(2), _
                                                            lv_articulos.ListItems(fila).SubItems(3), _
                                                            lv_articulos.ListItems(fila).SubItems(4), _
                                                            lv_articulos.ListItems(fila).SubItems(6), _
                                                            lv_articulos.ListItems(fila).SubItems(7), _
                                                            lv_articulos.ListItems(fila).SubItems(8))
                                    
                                    pb_AvanceImpresion.Value = fila
                                    x = Shell(App.Path & "\EtiPre.bat", vbHide)
                                    Sleep 1000
                                Next
                                MsgBox "Se mandaron correctamente las impresiones", vbInformation, "SID"
                                cbo_TiposEtiquetas.Enabled = True
                                cbo_TiposEtiquetas.Text = ""
                                pb_AvanceImpresion.Visible = False
                                lv_articulos.ListItems.Clear
                            Else
                                MsgBox "No se encontr? la configuracion de la etiqueta " & cbo_TiposEtiquetas.Text, vbCritical, "SID"
                            End If
                            
                            rsConfiguracion.Close
                        Else
                            MsgBox "Error al conectar al servidor ", vbCritical, "SID"
                        End If
                    
                Else
                    MsgBox "Favor de preparar la impresora con el papel correcto", vbExclamation, "SID"
                End If
            
        Else
            MsgBox "Favor de Seleccionar el Tama?o de la Etiqueta", vbCritical, "SID"
        End If
    Else
        MsgBox "No hay Informacion Para Imprimir", vbExclamation, "SID"
    End If

End Sub

Private Sub cmd_nuevo_Click()
    If lv_articulos.ListItems.Count > 0 Then
        If MsgBox("?Esta seguro de comenzar?, esta operacion borrar? lo capturado", vbYesNo, "SIP") = vbYes Then
            Call pro_Nuevo
        End If
    Else
        Call pro_Nuevo
    End If
End Sub
Private Sub pro_Nuevo()
    cbo_TiposEtiquetas.Text = ""
    txt_cantidad.Text = ""
    txt_codigo.Text = ""
    txt_descripcion.Text = ""
    lv_articulos.ListItems.Clear
End Sub
Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command2_Click()
    frmConImpEtiPre.Show vbModal
End Sub

Private Sub Form_Load()
    Dim rsEtiquetas As New ADODB.recordSet
    If Conectar_BDCompucaja(conCompucaja, "compucaja", "srvtdacantia") Then
        rsEtiquetas.Open "Select inte_eti_etiqueta_id, vcha_eti_etiqueta " & _
                        "From tb_etiquetasConfiguracion with(nolock) ", _
                    conCompucaja, _
                    adOpenDynamic, _
                    adLockOptimistic
            
        Call RecsetToCombo(cbo_TiposEtiquetas.hwnd, rsEtiquetas, 1)
        rsEtiquetas.Close
    Else
        MsgBox "Error al conectar con el servidor", vbCritical, "SID"
    End If
    If conCompucaja.State = 1 Then conCompucaja.Close
    Top = 1000
    Left = 2000
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub lv_articulos_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If lv_articulos.ListItems.Count > 0 Then
        If Button = 2 Then
            Me.PopupMenu OpcionesEtiquetas
        Else
            OpcionesEtiquetas.Enabled = False
            If Button = 1 Then
                txt_codigo.Text = lv_articulos.selectedItem
                txt_descripcion.Text = lv_articulos.selectedItem.SubItems(1)
                txt_cantidad.Text = lv_articulos.selectedItem.SubItems(5)
                txt_cantidad.Enabled = True
                txt_cantidad.SetFocus
            End If
        End If
    End If


End Sub

Private Sub mnuEliminar_Click()
    If lv_articulos.ListItems.Count > 0 Then
        lv_articulos.ListItems.Remove lv_articulos.selectedItem.Index
        txt_codigo.Text = ""
        txt_descripcion.Text = ""
        txt_cantidad.Text = ""
        txt_codigo.SetFocus
    End If
End Sub

Private Sub txt_Cantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim bln_existe  As Boolean
    Dim i As Integer
    If KeyCode = 13 Then
        If IsNumeric(txt_cantidad.Text) Then
            If CInt(txt_cantidad.Text) > 0 Then
                If cbo_TiposEtiquetas.Text <> "" Then
                    bln_existe = False
                    For i = 1 To lv_articulos.ListItems.Count
                        If lv_articulos.ListItems(i).Text = Trim(txt_codigo.Text) Then
                            bln_existe = True
                            lv_articulos.ListItems(i).SubItems(5) = txt_cantidad.Text
                            lv_articulos.SetFocus
                            txt_cantidad.Text = ""
                            txt_codigo.Text = ""
                            txt_descripcion.Text = ""
                            txt_cantidad.Enabled = False
                            
                            Exit For
                        End If
                    Next
                    If bln_existe = False Then
                        
                        If Conectar_BDCompucaja(conCompucaja, "compucaja", "srvtdacantia") Then
                        
                            Call pro_AgregaFila(txt_codigo.Text, txt_cantidad.Text)
                            txt_cantidad.Text = ""
                            txt_codigo.Text = ""
                            cbo_TiposEtiquetas.Enabled = False
                            txt_descripcion.Text = ""
                            txt_cantidad.Enabled = False
                            txt_codigo.SetFocus
                            conCompucaja.Close
                        Else
                            MsgBox "Error al conectar al servidor de Compucaja", vbCritical, "SID"
                        End If
                    End If
                Else
                    MsgBox "Favor de seleccionar el tipo de etiqueta", vbExclamation, "SIP"
                End If

            Else
                MsgBox "Solo se aceptan cantidades mayores a ''Cero''", vbCritical, "SID"
            End If
        Else
            MsgBox "Solo se aceptan numeros", vbCritical, "SID"
        End If
    End If
End Sub


Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Conectar_BDCompucaja(conCompucaja, "compucaja", "srvtdacantia") Then
            Call pro_buscarCodigo(txt_codigo.Text)
            If txt_descripcion.Text <> "" Then
                txt_cantidad.Enabled = True
                txt_cantidad.SetFocus
            Else
                MsgBox "No se encontr? el articulo", vbCritical, "SID"
            End If
            conCompucaja.Close
        Else
            MsgBox "Error al conectar al servidor de Compucaja", vbCritical, "SID"
        End If
    End If
    If KeyCode = 116 Then
        Call pro_BuscarArticulo
    End If
End Sub

Private Sub pro_BuscarArticulo()
    Dim frmBuscarArticulo As New frmBucarArticulos
    
    With frmBuscarArticulo
        
        .catalogo = " de Articulos de Cantia"
        .sentenciaSql = "SELECT sku Codigo, nombre Descripcion " & _
                    "FROM VW_ETIQUETASPROD WITH (NOLOCK) " & _
                    "WHERE (sku  LIKE '%XXXXXX%' " & _
                    "OR nombre LIKE '%XXXXXX%') "
        
        .comodin = "XXXXXX"
        
        .Show vbModal
        'Revisar si se seleccion? alguna clase de art?culo.
        If .valorSeleccionado1 <> "" Then
                txt_codigo.Text = .valorSeleccionado1
                txt_descripcion.Text = .valorSeleccionado2
                txt_cantidad.Enabled = True
                txt_cantidad.SetFocus
        End If
    End With
    
    'Liberar espacio de memoria ocupado por los objetos instanciados.
    Set frmBuscarArticulo = Nothing
    
End Sub



Private Function pro_AgregaFila(strCodigo As String, Cantidad As Integer) As String
On Error GoTo ErrorBusca:
    Dim rsEtiquetas As New ADODB.recordSet
    Dim strQry As String
    Dim strTitulo1 As String
    Dim strTitulo2 As String
    Dim item As ListItem
    
    pro_AgregaFila = ""
    strQry = "Select  vcha_eti_campoEncabezado1, " & _
                    " vcha_eti_campoEncabezado2 " & _
            "From tb_etiquetasConfiguracion with(nolock) " & _
            "where vcha_eti_etiqueta = '" & cbo_TiposEtiquetas.Text & "' "
    rsEtiquetas.Open strQry, _
                conCompucaja, _
                adOpenDynamic, _
                adLockOptimistic
                
    strTitulo1 = rsEtiquetas("vcha_eti_campoEncabezado1").Value
    strTitulo2 = IIf(rsEtiquetas("vcha_eti_campoEncabezado2").Value = "", "''", rsEtiquetas("vcha_eti_campoEncabezado2").Value)
    rsEtiquetas.Close
    strQry = "select art_codigo, " & _
                    "articulo, " & _
                    "col_nombre, " & _
                    "precio1, " & _
                    "isnull(via_medida,'') via_medida, " & _
                    "isnull(via_localizacion,'') loc, " & _
                    strTitulo1 & " categoria, " & _
                    strTitulo2 & " diseno " & _
            "from vw_info_etiquetas " & _
            "where art_codigo ='" & strCodigo & "'"
    rsEtiquetas.Open strQry, _
            conCompucaja, _
            adOpenDynamic, _
            adLockOptimistic
    If rsEtiquetas.RecordCount > 0 Then
        If Not IsNull(rsEtiquetas("categoria").Value) Then
            If Not IsNull(rsEtiquetas("diseno").Value) Then
                Set item = lv_articulos.ListItems.Add(, , rsEtiquetas("art_codigo").Value)
                item.SubItems(1) = IIf(IsNull(rsEtiquetas("articulo").Value), "", rsEtiquetas("articulo").Value)
                item.SubItems(2) = IIf(IsNull(rsEtiquetas("col_nombre").Value), "", rsEtiquetas("col_nombre").Value)
                item.SubItems(3) = IIf(IsNull(rsEtiquetas("precio1").Value), "0", rsEtiquetas("precio1").Value)
                item.SubItems(4) = IIf(IsNull(rsEtiquetas("via_medida").Value), "", rsEtiquetas("via_medida").Value)
                item.SubItems(5) = Cantidad
                item.SubItems(6) = IIf(IsNull(rsEtiquetas("loc").Value), "", UCase(Mid(rsEtiquetas("loc").Value, 1, 1)) & LCase(Mid(rsEtiquetas("loc").Value, 2, 200)))
                item.SubItems(7) = IIf(IsNull(rsEtiquetas("categoria").Value), "", Trim(Replace(rsEtiquetas("categoria").Value, rsEtiquetas("ART_CODIGO").Value, "")))
                item.SubItems(8) = IIf(IsNull(rsEtiquetas("diseno").Value), "", rsEtiquetas("diseno").Value)
            Else
                MsgBox "Falta El dise?o", vbCritical, "SID"
            End If
        Else
            MsgBox "Falta La Categoria", vbCritical, "SID"
        End If
    Else
        pro_AgregaFila = rsEtiquetas("sku").Value
    End If
    rsEtiquetas.Close
    Set rsEtiquetas = Nothing
    
Exit Function
ErrorBusca:
    If rsEtiquetas.State = 1 Then rsEtiquetas.Close
    pro_AgregaFila = Err.Description
    MsgBox Err.Description, vbCritical, "SIP"
End Function

Private Sub pro_buscarCodigo(strCodigo As String)

On Error GoTo ErrorBusca:
    Dim rsEtiquetas As New ADODB.recordSet
    Dim strQry As String
    
    strQry = "select nombre " & _
            "from VW_ETIQUETASPROD " & _
            "where sku ='" & strCodigo & "'"
    rsEtiquetas.Open strQry, _
            conCompucaja, _
            adOpenDynamic, _
            adLockOptimistic
    If rsEtiquetas.RecordCount > 0 Then
        txt_descripcion.Text = rsEtiquetas("nombre").Value
    Else
        txt_descripcion.Text = ""
    End If
    rsEtiquetas.Close
    Set rsEtiquetas = Nothing
Exit Sub
ErrorBusca:
    If rsEtiquetas.State = 1 Then rsEtiquetas.Close
    MsgBox Err.Description, vbCritical, "SID"
    
End Sub


Private Function Conectar_BDCompucaja(ByRef cnnCBD As ADODB.Connection, ByVal bd As String, ByVal servidor As String) As Boolean
    'Variables de bloque
    Dim strConnection_String As String
    
On Error GoTo Error_Conectar_BDS
    Conectar_BDCompucaja = True
    'Establecer connection strings para realizar las conexiones a las bases de
    'datos
    If conCompucaja.State = 1 Then
        conCompucaja.Close
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


Private Sub pro_generaEtiqueta(strCodigo As String, strNombre As String, strCantidad As String, _
                                strColor As String, strPrecio As String, StrViaMedida As String, _
                                strLoc As String, strCategoria As String, strDiseno As String)
On Error Resume Next
    Dim rsPuerto As New ADODB.recordSet
    Dim linea As String, var_comillas As String
    Dim nomArticulo As String
    Dim NumCaracteres As Integer
    Dim caracter As Chart
    Dim strPos As String
    Dim numbPosiEnc2 As Integer
    Dim numbPosiX As Integer
    Dim numbPosiY As Integer
    Dim numbTamLinea As Integer
    Dim numbTama?o As Integer
    Dim numbTama?oAgregado As Integer
    Dim strNomMaq As String
    Dim numPosIni As Integer
    Dim numPosFin As Integer
    Dim numbEntraTitulos As Integer
    strNomMaq = fun_NombrePc
    strPos = rsConfiguracion("bint_eti_posicion").Value
    rsPuerto.Open "Select * from tb_etiquetasConfiguracionImpresora where vcha_eci_maquina= '" & strNomMaq & "' ", cnn_compucaja, adOpenDynamic, adLockOptimistic
    nomArticulo = ""
    NumCaracteres = 0
    caracter = ""
    strNombre = Trim(Replace(strNombre, strColor, ""))
    Open (App.Path & "\EtiPre.bat") For Output As #1
    If rsPuerto("vcha_eci_puerto").Value = "- LPT1" Then
        Print #1, "COPY " & App.Path & "\EtiPre.txt LPT1"
    Else
        Print #1, "COPY " & App.Path & "\EtiPre.txt \\" & rsPuerto("vcha_eci_maquina").Value & "\" & rsPuerto("vcha_eci_maquinaImpresora").Value
    End If
    Close #1
    
    var_comillas = Chr(34)
    'Open (App.path & "EtiPre.txt") For Output As #1
    Open (App.Path & "\EtiPre.txt") For Output As #1
    Print #1, "US"
    Print #1, "N"
    Print #1, "UN"
    Print #1, "Q" & (rsConfiguracion("floa_eti_largo").Value * 80) & "," & rsConfiguracion("vcha_eti_lineaNegra").Value & rsConfiguracion("floa_eti_TamLineaNegra").Value * 80 & "+" & (rsConfiguracion("floa_eti_TamLineaNegra").Value * 80)
    Print #1, "q" & (rsConfiguracion("floa_eti_ancho").Value * 80)
    Print #1, "S2"
    Print #1, "D8"
    Print #1, "ZB"
    
    
    NumCaracteres = rsConfiguracion("bint_eti_CaracteresEncabezado1").Value
    numbPosiX = rsConfiguracion("bint_eti_posi_x_Encabezado1").Value
    numbPosiY = rsConfiguracion("bint_eti_posi_y_Encabezado1").Value
    numbTamLinea = rsConfiguracion("int_eti_TamLineaEncabezado1").Value
    numbTama?oAgregado = -1
    
    numbEntraTitulos = Len(strCategoria)
    If numbTama?oAgregado = -1 Then
        numbTama?oAgregado = 0
    End If
    While NumCaracteres >= Len(nomArticulo) And numbEntraTitulos > 0
        If InStr(strCategoria, " ") > 0 Then
            nomArticulo = nomArticulo & Mid(strCategoria, 1, InStr(strCategoria, " "))
        Else
            nomArticulo = nomArticulo & strCategoria
            strCategoria = ""
            NumCaracteres = -1
        End If
        strCategoria = Trim(Mid(strCategoria, InStr(strCategoria, " ") + 1, 200))
        If InStr(strCategoria, " ") > 0 Then
            numbTama?o = Len(Trim(nomArticulo & Mid(strCategoria, 1, InStr(Trim(strCategoria), " "))))
        Else
            numbTama?o = Len(nomArticulo & strCategoria)
        End If
        If numbTama?o <= NumCaracteres Then
            If InStr(strCategoria, " ") > 0 Then
                numbTama?oAgregado = Len(nomArticulo)
            End If
        Else
             Print #1, "A" & numbPosiX & "," & _
                    numbPosiY & "," & _
                    strPos & "," & _
                    rsConfiguracion("vcha_eti_fuenteEncabezado1").Value & "," & _
                    "1,1,N, " & var_comillas & Trim(nomArticulo) & var_comillas; ""
            nomArticulo = ""
            'strCategoria = Mid(strCategoria, InStr(strCategoria, " ") + 1, 200)
            If strPos = "0" Then
                numbPosiY = numbPosiY + numbTamLinea
            Else
                numbPosiX = numbPosiX - numbTamLinea
            End If
        End If
    Wend
    
        
    NumCaracteres = rsConfiguracion("int_eti_CaracteresEncabezado2").Value
    numbTamLinea = rsConfiguracion("int_eti_TamLineaEncabezado2").Value
    nomArticulo = ""
    numbEntraTitulos = Len(strDiseno)
    
    
    If numbTama?oAgregado = -1 Then
        numbTama?oAgregado = 0
    End If
    
    
    While NumCaracteres >= Len(nomArticulo) And numbEntraTitulos > 0
        If InStr(strDiseno, " ") > 0 Then
            nomArticulo = nomArticulo & Mid(strDiseno, 1, InStr(strDiseno, " "))
        Else
            nomArticulo = nomArticulo & strDiseno
            strDiseno = ""
            NumCaracteres = -1
        End If
        strDiseno = Trim(Mid(strDiseno, InStr(strDiseno, " ") + 1, 200))
        If InStr(strDiseno, " ") > 0 Then
            numbTama?o = Len(Trim(nomArticulo & Mid(strDiseno, 1, InStr(Trim(strDiseno), " "))))
        Else
            numbTama?o = Len(nomArticulo & strDiseno)
        End If
        If numbTama?o <= NumCaracteres Then
            If InStr(strDiseno, " ") > 0 Then
                numbTama?oAgregado = Len(nomArticulo)
            End If
        Else
             Print #1, "A" & numbPosiX & "," & _
                    numbPosiY & "," & _
                    strPos & "," & _
                    rsConfiguracion("vcha_eti_fuenteEncabezado2").Value & "," & _
                    "1,1,N, " & var_comillas & Trim(nomArticulo) & var_comillas; ""
            nomArticulo = ""
            'strdiseno = Mid(strdiseno, InStr(strdiseno, " ") + 1, 200)
            If strPos = "0" Then
                numbPosiY = numbPosiY + numbTamLinea
            Else
                numbPosiX = numbPosiX - numbTamLinea
            End If
        End If
    Wend
    
    If Len(strPrecio) > 3 Then
        strPrecio = Mid(strPrecio, 1, Len(strPrecio) - 3) & "," & Right(strPrecio, 3)
        
    End If
    
    If Len(strPrecio) >= rsConfiguracion("int_eti_CaracteresPresio").Value Then
        If strPos = 0 Then
            numbPosiX = 5
        Else
            numbPosiX = rsConfiguracion("bint_eti_posi_x_SignoPrecio").Value
        End If
    Else
        numbPosiX = rsConfiguracion("bint_eti_posi_x_SignoPrecio").Value
    End If
    Print #1, "A" & numbPosiX & "," & rsConfiguracion("bint_eti_posi_y_SignoPrecio").Value & "," & strPos & "," & rsConfiguracion("vcha_eti_fuenteSignoPrecio").Value & "," & "1,1,N, " & var_comillas & "$" & var_comillas
    If Len(strPrecio) >= rsConfiguracion("int_eti_CaracteresPresio").Value Then
        If strPos = 0 Then
            numbPosiX = 5
        Else
            numbPosiX = rsConfiguracion("bint_eti_posi_x_Presio").Value
        End If
    Else
        numbPosiX = rsConfiguracion("bint_eti_posi_x_Presio").Value
    End If
    Print #1, "A" & numbPosiX & "," & _
                rsConfiguracion("bint_eti_posi_y_Presio").Value & "," & _
                strPos & "," & _
                rsConfiguracion("vcha_eti_fuentePrecio").Value & "," & _
                "1,1,N, " & var_comillas & strPrecio & var_comillas; ""
                
    If rsConfiguracion("floa_eti_largo").Value > 10 Or strPos = "1" Then
        If strPos = "1" Then
            Print #1, "A" & rsConfiguracion("bint_eti_posi_x_Codigo1").Value & "," & _
                    rsConfiguracion("bint_eti_posi_y_Codigo1").Value & "," & _
                    strPos & "," & _
                    rsConfiguracion("vcha_eti_fuenteCodigo2").Value & "," & _
                    "1,1,N, " & var_comillas & "COD: "; var_comillas; ""
            Print #1, "A" & rsConfiguracion("bint_eti_posi_x_Codigo1").Value & "," & rsConfiguracion("bint_eti_posi_y_Codigo1").Value + 85 & "," & strPos & "," & rsConfiguracion("vcha_eti_fuenteCodigo1").Value & "," & "1,1,N, " & var_comillas & strCodigo & var_comillas
        Else
            Print #1, "A" & rsConfiguracion("bint_eti_posi_x_Codigo1").Value & "," & rsConfiguracion("bint_eti_posi_y_Codigo1").Value & "," & strPos & "," & rsConfiguracion("vcha_eti_fuenteCodigo2").Value & "," & "1,1,N, " & var_comillas & "CODIGO: " & var_comillas
            Print #1, "A" & rsConfiguracion("bint_eti_posi_x_Codigo1").Value + 180 & "," & _
                    rsConfiguracion("bint_eti_posi_y_Codigo1").Value - 5 & "," & _
                    strPos & "," & _
                    rsConfiguracion("vcha_eti_fuenteCodigo1").Value & "," & _
                    "1,1,N, " & var_comillas & strCodigo & var_comillas; ""
        End If
    Else
        If rsConfiguracion("bint_eti_posi_x_Codigo1").Value >= 0 Then
            Print #1, "A" & rsConfiguracion("bint_eti_posi_x_Codigo1").Value & "," & _
                    rsConfiguracion("bint_eti_posi_y_Codigo1").Value & "," & _
                    strPos & "," & _
                    rsConfiguracion("vcha_eti_fuenteCodigo2").Value & "," & _
                    "1,1,N, " & var_comillas & "CODIGO: " & var_comillas; ""
            If rsConfiguracion("int_eti_CaracteresCodigo1").Value > Len(strCodigo) Then
                Print #1, "A" & rsConfiguracion("bint_eti_posi_x_Codigo1").Value + 155 & "," & _
                    rsConfiguracion("bint_eti_posi_y_Codigo1").Value - 10 & "," & _
                    strPos & "," & _
                    rsConfiguracion("vcha_eti_fuenteCodigo1").Value & "," & _
                    "1,1,N, " & var_comillas & strCodigo & var_comillas; ""
            Else
                If rsConfiguracion("bint_eti_posi_y_Codigo2").Value = -1 Then
                    Print #1, "A" & rsConfiguracion("bint_eti_posi_x_Codigo1").Value & "," & rsConfiguracion("bint_eti_posi_y_Codigo1").Value + rsConfiguracion("int_eti_TamLineaCodigo1").Value & "," & strPos & "," & rsConfiguracion("vcha_eti_fuenteCodigo1").Value & "," & "1,1,N, " & var_comillas & strCodigo & var_comillas
                Else
                Print #1, "A" & rsConfiguracion("bint_eti_posi_x_Codigo1").Value & "," & _
                    rsConfiguracion("bint_eti_posi_y_Codigo1").Value + rsConfiguracion("int_eti_TamLineaCodigo1").Value & "," & _
                    strPos & "," & _
                    rsConfiguracion("vcha_eti_fuenteCodigo1").Value & "," & _
                    "1,1,N, " & var_comillas & strColor & var_comillas
                End If
            End If
        End If
    End If
    
    If rsConfiguracion("bint_eti_posi_x_Codigo2").Value >= 0 Then
        Print #1, "A" & rsConfiguracion("bint_eti_posi_x_Codigo2").Value & "," & rsConfiguracion("bint_eti_posi_y_Codigo2").Value & "," & strPos & "," & rsConfiguracion("vcha_eti_fuenteCodigo2").Value & ",1,1,N, " & var_comillas & "COLOR: " & strColor & var_comillas; ""
    End If
    
    If rsConfiguracion("bint_eti_posi_x_Medida").Value >= 0 Then
        Print #1, "A" & rsConfiguracion("bint_eti_posi_x_Medida").Value & "," & _
                rsConfiguracion("bint_eti_posi_y_Medida").Value & "," & _
                strPos & "," & _
                rsConfiguracion("vcha_eti_fuenteMedida").Value & "," & _
                "1,1,N, " & var_comillas & "MEDIDA: " & var_comillas; ""
        If strPos = 0 Then
            Print #1, "A" & rsConfiguracion("bint_eti_posi_x_Medida").Value & "," & _
                    rsConfiguracion("bint_eti_posi_y_Medida").Value + rsConfiguracion("bint_eti_TamLineaMedida").Value & "," & _
                    strPos & "," & _
                    rsConfiguracion("vcha_eti_fuenteMedida").Value & "," & _
                    "1,1,N, " & var_comillas & StrViaMedida & var_comillas; ""
        Else
            Print #1, "A" & rsConfiguracion("bint_eti_posi_x_Medida").Value + rsConfiguracion("bint_eti_TamLineaMedida").Value & "," & _
                    rsConfiguracion("bint_eti_posi_y_Medida").Value & "," & _
                    strPos & "," & _
                    rsConfiguracion("vcha_eti_fuenteMedida").Value & "," & _
                    "1,1,N, " & var_comillas & StrViaMedida & var_comillas; ""
        End If
    End If
    
    If rsConfiguracion("bint_eti_posi_x_ubicacion").Value >= 0 Then
        If strPos = "1" Then
            Print #1, "A" & rsConfiguracion("bint_eti_posi_x_ubicacion").Value + rsConfiguracion("int_eti_TamLineaUbicacion").Value & ",23," & strPos & ",i,1,1,N, " & var_comillas & "Encu?ntralo en: " & var_comillas; ""
        Else
            Print #1, "A23," & rsConfiguracion("bint_eti_posi_y_ubicacion").Value + rsConfiguracion("int_eti_TamLineaUbicacion").Value & "," & strPos & ",i,1,1,N, " & var_comillas & "Encu?ntralo en: " & var_comillas; ""
        End If
        Print #1, "A" & rsConfiguracion("bint_eti_posi_x_ubicacion").Value & "," & rsConfiguracion("bint_eti_posi_y_ubicacion").Value & "," & strPos & "," & rsConfiguracion("vcha_eti_fuenteUbicacion").Value & ",1,1,N, " & var_comillas & strLoc & var_comillas; ""
    End If
    
    Print #1, "P" & strCantidad
    Print #1, ""
    Close #1
End Sub




