VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frmmenu2 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   Icon            =   "Frmmenu2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11685
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   7440
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   176389
            MinWidth        =   176389
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":0E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":173E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":1A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":1D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":264C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList fserna 
      Left            =   4845
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":2966
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":2A78
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":2B8A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":2C9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":2FEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":3340
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":3692
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":39E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":3D36
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":4088
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":43DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":472C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":4A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmmenu2.frx":4DD0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7275
      Left            =   15
      TabIndex        =   0
      Top             =   105
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   12832
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "fserna"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuAlmacen 
      Caption         =   "Almacen"
      Visible         =   0   'False
      Begin VB.Menu mnuAlmacenes 
         Caption         =   "Almacenes"
      End
      Begin VB.Menu sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuartalm 
         Caption         =   "Articulos de Almacen"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnumovimientos 
         Caption         =   "Movimientos"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProveedores 
         Caption         =   "Proveedores"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnulineas 
         Caption         =   "Lineas"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuunidades2 
         Caption         =   "Unidades"
      End
   End
   Begin VB.Menu mnunotas 
      Caption         =   "notas"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuPrenomina 
      Caption         =   "prenomina"
      Visible         =   0   'False
      Begin VB.Menu mnuEmpleados 
         Caption         =   "Empleados"
      End
      Begin VB.Menu mnuModulos 
         Caption         =   "Modulos"
      End
      Begin VB.Menu mnuDestajos 
         Caption         =   "Destajos"
      End
      Begin VB.Menu mnuFS 
         Caption         =   "Fueras de Standar"
      End
      Begin VB.Menu mnuCategorias 
         Caption         =   "Categorias"
      End
   End
   Begin VB.Menu mnumovimientos2 
      Caption         =   "Movimientos"
      Visible         =   0   'False
      Begin VB.Menu mnumovimientos3 
         Caption         =   "Movimientos"
      End
      Begin VB.Menu mnuclasesdemovimientos 
         Caption         =   "Clases de movimientos"
      End
   End
   Begin VB.Menu mnuarticulos3 
      Caption         =   "Articulos"
      Visible         =   0   'False
      Begin VB.Menu mnuArticulos4 
         Caption         =   "Art�culos"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnucatalogos 
         Caption         =   "Catalogos"
      End
      Begin VB.Menu mnulicencias 
         Caption         =   "Licencias"
      End
      Begin VB.Menu mnudise�os 
         Caption         =   "Dise�os"
      End
      Begin VB.Menu mnulineas2 
         Caption         =   "Lineas"
      End
      Begin VB.Menu mnusublineas2 
         Caption         =   "Sublineas"
      End
      Begin VB.Menu mnuproductos 
         Caption         =   "Productos"
      End
      Begin VB.Menu mnusubtipodeproductos 
         Caption         =   "Subtipo de productos"
      End
      Begin VB.Menu mnuclases 
         Caption         =   "Clases"
      End
      Begin VB.Menu mnuestampados 
         Caption         =   "Estampados"
      End
      Begin VB.Menu mnucolores 
         Caption         =   "Colores"
      End
      Begin VB.Menu mnutonos 
         Caption         =   "Tonos"
      End
      Begin VB.Menu mnuusos 
         Caption         =   "Usos"
      End
      Begin VB.Menu mnusubtipodesusos 
         Caption         =   "Subtipo de usos"
      End
      Begin VB.Menu mnutallas 
         Caption         =   "Tallas"
      End
      Begin VB.Menu mnuunidades3 
         Caption         =   "Unidades"
      End
      Begin VB.Menu mnucajas 
         Caption         =   "Cajas"
      End
      Begin VB.Menu mnuubicaciones 
         Caption         =   "Ubicaciones"
      End
   End
   Begin VB.Menu mnucatalogoscxc 
      Caption         =   "Catalogos cxc"
      Visible         =   0   'False
      Begin VB.Menu mnuclientescxc 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mnugrupodeclientescxc 
         Caption         =   "Grupo de clientes"
      End
      Begin VB.Menu mnuagentescxc 
         Caption         =   "Agentes"
      End
      Begin VB.Menu mnutipodeagentescxc 
         Caption         =   "Tipo de agentes"
      End
      Begin VB.Menu mnurutascxc 
         Caption         =   "Rutas"
      End
      Begin VB.Menu mnudireccionescxc 
         Caption         =   "Direcciones"
      End
      Begin VB.Menu mnuzonascxc 
         Caption         =   "Zonas"
      End
      Begin VB.Menu mnucomisiones1 
         Caption         =   "Comosiones"
      End
   End
   Begin VB.Menu mnucatalogostesoreria 
      Caption         =   "catalogos tesoreria"
      Visible         =   0   'False
      Begin VB.Menu mnumonedastesoreria 
         Caption         =   "Monedas"
      End
      Begin VB.Menu mnutipodecambiotesoreria 
         Caption         =   "Tipo de cambio"
      End
   End
   Begin VB.Menu mnucatalogosfinanzas 
      Caption         =   "Catalogos"
      Visible         =   0   'False
      Begin VB.Menu mnupaises 
         Caption         =   "Cat�logo de paises"
      End
      Begin VB.Menu mnuestados 
         Caption         =   "Cat�logo de Estados"
      End
      Begin VB.Menu mnuciudades 
         Caption         =   "Cat�logo de Ciudades"
      End
      Begin VB.Menu mnurutas 
         Caption         =   "Cat�logo de rutas"
      End
      Begin VB.Menu mnutipodeagentes 
         Caption         =   "Cat�logo de tipo de agentes"
      End
      Begin VB.Menu mnuagentes 
         Caption         =   "Cat�logo de agentes"
      End
      Begin VB.Menu mnucomisiones 
         Caption         =   "Cat�logo de comisiones"
      End
      Begin VB.Menu mnuaseguradoras 
         Caption         =   "Cat�logo de aseguradoras"
      End
      Begin VB.Menu mnumonedas 
         Caption         =   "Cat�logo de monedas"
      End
      Begin VB.Menu mnutipocambio 
         Caption         =   "Tipo cambio"
      End
   End
   Begin VB.Menu mnuclientes 
      Caption         =   "clientes"
      Visible         =   0   'False
      Begin VB.Menu mnuclientesporveedores 
         Caption         =   "Clientes"
      End
      Begin VB.Menu x 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclpvendedores 
         Caption         =   "Vendedores"
      End
      Begin VB.Menu mnuclprutas 
         Caption         =   "Rutas"
      End
      Begin VB.Menu mnuclpmonedas 
         Caption         =   "Monedas"
      End
      Begin VB.Menu mnuclpplazos 
         Caption         =   "Plazos"
      End
      Begin VB.Menu mnuclptiposdeclientes 
         Caption         =   "Tipos de clientes"
      End
      Begin VB.Menu mnuclplistadeprecios 
         Caption         =   "Lista de precios"
      End
      Begin VB.Menu mnuclpcanalesdeventa 
         Caption         =   "Canales de venta"
      End
      Begin VB.Menu mnuclpagrupadores 
         Caption         =   "Agrupadores"
      End
      Begin VB.Menu mnuclptransportes 
         Caption         =   "Transportes"
      End
      Begin VB.Menu mnuclpgrupos 
         Caption         =   "Grupos de clientes"
      End
   End
   Begin VB.Menu mnumovmovimientosdiarios 
      Caption         =   "Movimientos diarios"
      Visible         =   0   'False
      Begin VB.Menu mnumoventradasdeproduccion 
         Caption         =   "Entradas de producci�n"
      End
   End
End
Attribute VB_Name = "Frmmenu2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
var_x = 1
If var_x = 1 Then
cnn.Close
cnn.Open var_conexion_string

Dim dl As Long                                 ' Valor devuelto por la funci�n API
Dim sAttributes As String                  ' Aributos
Dim sDriver As String                       ' Nombre del controlador
Dim sDescription As String                ' Descripci�n del DSN
Dim sDsnName As String                  ' Nombre del DSN

Const ODBC_ADD_SYS_DSN As Long = 4         ' Se crear� un DSN de sistema
Const vbAPINull As Long = 0&                         ' Puntero NULL

' se elimina
Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminar� un DSN de sistema
sDsnName = "DSN=sqlsistema"
sDriver = "SQL Server"
dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)

'se crea
sDsnName = "sqlsistema"
sDescription = "sqlsistema"
sDriver = "SQL Server"
sAttributes = "DSN=" & sDsnName & Chr(0)
sAttributes = sAttributes & "Server=" + parametros(0) & Chr$(0)
sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
sAttributes = sAttributes & "Database=" + var_bd_reportes & Chr(0)
strAttributes = strAttributes & "UID=sa" & Chr$(0)
strAttributes = strAttributes & "PWD=elia" & Chr$(0)
dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = 116 Then
      var_activa_forma_existencias_generales = "MENU"
      frmoracle_existencias_rapidas.Show
   End If
   If Shift = 1 And KeyCode = 117 Then
      var_activa_forma_existencias_generales = "MENU"
      frmoracle_ubicacion_articulos.Show 1
   End If
   If Shift = 1 And KeyCode = 118 Then
      var_activa_forma_existencias_generales = "MENU"
      frmexistencias_rapidas.Show 1
   End If
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub



Private Sub Form_Load()
   Dim nodX As Node
   Dim var_icono As Integer
   Dim fs, D, f, s
   var_x = 1
   If var_x = 1 Then
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set f = fs.GetFile(App.Path + "/sistema.exe")
   var_cadena_seguridad = ""
   var_archivo_local = App.Path + "\sistema.exe"
   var_cadena_puerto = ""
   If var_modo_texto_ip = 1 Then
      var_cadena_puerto = "DVR: " + var_dvr_texto + ", PUERTO: " + var_puerto_texto
   End If
   var_fecha_sistema_string = "Usted se encuentra conectado al servidor de " + parametros(0) + " y a la base de datos " + parametros(1) + " " + var_cadena_puerto
   Frmmenu2.Caption = var_fecha_sistema_string
   var_top = 0
   var_left = 0
   Frmmenu2.Top = var_top
   Frmmenu2.Left = var_left
   If var_clave_usuario_global <> "1" Then
      rsaux.Open "select * from tb_submenus where vcha_men_menu_id = '" + var_global_menu + "' order by vcha_men_menu_id,inte_sme_nivel,inte_sme_numero,char_sme_submenu_id", cnn, adOpenDynamic, adLockOptimistic
      TreeView1.Nodes.Clear
      If Not rsaux.EOF Then
         While Not rsaux.EOF
            var_n = rsaux(8).Value
            If IsNull(rsaux(9).Value) Then
               var_icono = 3
            Else
               If rsaux(9).Value = "00" Then
                  var_icono = 3
               Else
                  var_icono = 2
               End If
            End If
            If var_n = 1 Then
               var_c = Trim(Mid(rsaux(1).Value, 1, 4))
               Set nodX = TreeView1.Nodes.Add(, , """" + var_c + """", "" + rsaux(7).Value + "", var_icono)
            End If
            If var_n = 2 Then
               var_c2 = Trim(Mid(rsaux(1).Value, 1, 4))
               var_c3 = Trim(Mid(rsaux(1).Value, 1, 6))
               Set nodX = TreeView1.Nodes.Add("""" + var_c2 + """", tvwChild, """" + var_c3 + """", "" + rsaux(7).Value + "", var_icono)
            End If
            If var_n = 3 Then
               var_c3 = Trim(Mid(rsaux(1).Value, 1, 6))
               var_c4 = Trim(Mid(rsaux(1).Value, 1, 8))
               Set nodX = TreeView1.Nodes.Add("""" + var_c3 + """", tvwChild, """" + var_c4 + """", "" + rsaux(7).Value + "", var_icono)
            End If
            If var_n = 4 Then
               var_c4 = Trim(Mid(rsaux(1).Value, 1, 8))
               var_c5 = Trim(Mid(rsaux(1).Value, 1, 10))
               Set nodX = TreeView1.Nodes.Add("""" + var_c4 + """", tvwChild, """" + var_c5 + """", "" + rsaux(7).Value + "", var_icono)
            End If
            rsaux.MoveNext:
         Wend
      End If
      rsaux.Close
      TreeView1.Style = 7
   Else
      rsaux.Open "select * from tb_submenus where vcha_men_menu_id = '99' order by vcha_men_menu_id,inte_sme_nivel,inte_sme_numero,char_sme_submenu_id", cnn, adOpenDynamic, adLockOptimistic
      TreeView1.Nodes.Clear
      If Not rsaux.EOF Then
         While Not rsaux.EOF
            var_n = rsaux(8).Value
            If IsNull(rsaux(9).Value) Then
               var_icono = 3
            Else
               If rsaux(9).Value = "00" Then
                  var_icono = 3
               Else
                  var_icono = 2
               End If
            End If
            If var_n = 1 Then
               var_c = Trim(Mid(rsaux(1).Value, 1, 4))
               Set nodX = TreeView1.Nodes.Add(, , """" + var_c + """", "" + rsaux(7).Value + "", var_icono)
            End If
            If var_n = 2 Then
               var_c2 = Trim(Mid(rsaux(1).Value, 1, 4))
               var_c3 = Trim(Mid(rsaux(1).Value, 1, 6))
               Set nodX = TreeView1.Nodes.Add("""" + var_c2 + """", tvwChild, """" + var_c3 + """", "" + rsaux(7).Value + "", var_icono)
            End If
            If var_n = 3 Then
               var_c3 = Trim(Mid(rsaux(1).Value, 1, 6))
               var_c4 = Trim(Mid(rsaux(1).Value, 1, 8))
               Set nodX = TreeView1.Nodes.Add("""" + var_c3 + """", tvwChild, """" + var_c4 + """", "" + rsaux(7).Value + "", var_icono)
            End If
            If var_n = 4 Then
               var_c4 = Trim(Mid(rsaux(1).Value, 1, 8))
               var_c5 = Trim(Mid(rsaux(1).Value, 1, 10))
               Set nodX = TreeView1.Nodes.Add("""" + var_c4 + """", tvwChild, """" + var_c5 + """", "" + rsaux(7).Value + "", var_icono)
            End If
            rsaux.MoveNext:
         Wend
      End If
      rsaux.Close
      TreeView1.Style = 7
   End If
   If var_empresa = "15" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\ConectorEE\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\ConectorEE\envio\enviados"
   End If
   If var_empresa = "16" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\ConectorMYG\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\ConectorMYG\envio\enviados"
   End If
   If var_empresa = "30" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectortur\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectortur\envio\enviados"
   End If
   If var_empresa = "31" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectorcan\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectorcan\envio\enviados"
   End If
   If var_empresa = "31" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectorcan\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectorcan\envio\enviados"
   End If
   If var_empresa = "02" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectorvth\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectorvth\envio\enviados"
   End If
   If var_empresa = "03" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectorvth\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectorvth\envio\enviados"
   End If
   If var_empresa = "06" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectorvth\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectorvth\envio\enviados"
   End If
   If var_empresa = "18" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectorvth\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectorvth\envio\enviados"
   End If
   If var_empresa = "17" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectorvth\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectorvth\envio\enviados"
   End If
   If var_empresa = "29" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectorvth\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectorvth\envio\enviados"
   End If
   If var_empresa = "02" Or var_empresa = "03" Or var_empresa = "06" Or var_empresa = "18" Or var_empresa = "17" Or var_empresa = "29" Or var_empresa = "15" Or var_empresa = "16" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectorvth\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectorvth\envio\enviados"
   End If
   
   If var_empresa = "32" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectorare\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectorare\envio\enviados"
   End If
   If var_empresa = "33" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectormpu\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectormpu\envio\enviados"
   End If
   If var_empresa = "34" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectormul\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectormul\envio\enviados"
   End If
   If var_empresa = "35" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectorvth\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectorvth\envio\enviados"
   End If
   If var_empresa = "36" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectorsme\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectorsme\envio\enviados"
   End If
   If var_empresa = "37" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectorvth\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectorvth\envio\enviados"
   End If
   If var_empresa = "38" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectorvia\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectorvia\envio\enviados"
   End If
   If var_empresa = "39" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\ConectorCAN\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\ConectorCAN\envio\enviados"
   End If
   If var_empresa = "40" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\ConectorvIN\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\ConectorvIN\envio\enviados"
   End If
   If var_empresa = "41" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectorcop\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectorcop\envio\enviados"
   End If
   If var_empresa = "42" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\Conectorcma\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\Conectorcma\envio\enviados"
   End If
   If var_empresa = "43" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\ConectorVOP\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\ConectorvOP\envio\enviados"
   End If
   If var_empresa = "44" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\ConectorUTV\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\ConectorUTV\envio\enviados"
   End If
   If parametros(1) = "SIDALMACENBKP" Then
      var_ruta_documentos_electronicos = "\\FACELECTRONICA\fefiles\ConectorTST\envio\por_enviar"
      var_ruta_documentos_electronicos_pdf = "\\FACELECTRONICA\fefiles\ConectorTST\envio\enviados"
   End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_bloque_global = 4 Then
      Unload Me
      Frmacceso.Show
   Else
      End
   End If
End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
    i = TreeView1.selectedItem.Index
    TreeView1.Nodes(i).Image = 3
End Sub

Private Sub TreeView1_DblClick()
   var_pedido_internet = 0
   var_clave_nivel1 = 0
   var_clave_nivel2 = 0
   var_clave_nivel3 = 0
   var_clave_nivel4 = 0
   var_clave_nivel5 = 0
   If TreeView1.Nodes.Count > 0 Then
      var_c = TreeView1.selectedItem.Key
      var_longitud = Len(Trim(var_c))
      If var_longitud = 6 Then
         var_c = Trim(Mid(var_c, 2, 4))
      End If
      If var_longitud = 8 Then
         var_c = Trim(Mid(var_c, 2, 6))
      End If
      If var_longitud = 10 Then
         var_c = Trim(Mid(var_c, 2, 8))
      End If
      If var_longitud = 12 Then
         var_c = Trim(Mid(var_c, 2, 10))
      End If
      var_longitud = Len(Trim(var_c))
      If var_longitud = 4 Then
         var_c2 = var_c + "000000"
         If rs.State = 1 Then
            rs.Close
         End If
         rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
         var_nivel = rs(8).Value
         var_nivel1 = rs(2).Value
         var_clave_nivel1 = rs(2).Value
         var_nombre_submenu = rs(7).Value
         var_accion_submenu = rs(9).Value
         var_global_permiso1 = rs!inte_sme_permiso1
         var_global_permiso2 = rs!inte_sme_permiso2
         var_global_permiso3 = rs!inte_sme_permiso3
         var_global_permiso4 = rs!inte_sme_permiso4
         rs.Close
      End If
      If var_longitud = 6 Then
         var_c2 = var_c + "0000"
         rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
         var_nivel = rs(8).Value
         var_nivel1 = rs(2).Value
         var_nivel2 = rs(3).Value
         var_clave_nivel1 = rs(2).Value
         var_clave_nivel2 = rs(3).Value
         var_accion_submenu = rs(9).Value
         var_nombre_submenu = rs(7).Value
         var_global_permiso1 = rs!inte_sme_permiso1
         var_global_permiso2 = rs!inte_sme_permiso2
         var_global_permiso3 = rs!inte_sme_permiso3
         var_global_permiso4 = rs!inte_sme_permiso4
         rs.Close
      End If
      If var_longitud = 8 Then
         var_c2 = var_c + "00"
         rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
         var_nivel = rs(8).Value
         var_nivel1 = rs(2).Value
         var_nivel2 = rs(3).Value
         var_nivel3 = rs(4).Value
         var_accion_submenu = rs(9).Value
         var_nombre_submenu = rs(7).Value
         var_global_permiso1 = rs!inte_sme_permiso1
         var_global_permiso2 = rs!inte_sme_permiso2
         var_global_permiso3 = rs!inte_sme_permiso3
         var_global_permiso4 = rs!inte_sme_permiso4
         rs.Close
      End If
      If var_longitud = 10 Then
         var_c2 = var_c
         rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
         var_nivel = rs(8).Value
         var_nivel1 = rs(2).Value
         var_nivel2 = rs(3).Value
         var_nivel3 = rs(4).Value
         var_accion_submenu = rs(9).Value
         var_nombre_submenu = rs(7).Value
         var_global_permiso1 = rs!inte_sme_permiso1
         var_global_permiso2 = rs!inte_sme_permiso2
         var_global_permiso3 = rs!inte_sme_permiso3
         var_global_permiso4 = rs!inte_sme_permiso4
         rs.Close
      End If
      var_opcion_seguridad = 1
      If var_global_permiso1 = 1 Then
          If var_global_permiso2 = 1 Then
             frmpasswords2.Show 1
          Else
             frmpasswords.Show 1
          End If
      Else
          ejecuta_forma
      End If
   End If
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
    var_pedido_internet = 0
    i = TreeView1.selectedItem.Index
    TreeView1.Nodes(i).Image = 1
End Sub

Private Sub TreeView1_GotFocus()
   var_oracle_tipo_movimiento = ""
End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_pedido_internet = 0
      var_clave_nivel1 = 0
      var_clave_nivel2 = 0
      var_clave_nivel3 = 0
      var_clave_nivel4 = 0
      var_clave_nivel5 = 0
      If TreeView1.Nodes.Count > 0 Then
         var_c = TreeView1.selectedItem.Key
         var_longitud = Len(Trim(var_c))
         If var_longitud = 6 Then
            var_c = Trim(Mid(var_c, 2, 4))
         End If
         If var_longitud = 8 Then
            var_c = Trim(Mid(var_c, 2, 6))
         End If
         If var_longitud = 10 Then
            var_c = Trim(Mid(var_c, 2, 8))
         End If
         If var_longitud = 12 Then
            var_c = Trim(Mid(var_c, 2, 10))
         End If
         var_longitud = Len(Trim(var_c))
         If var_longitud = 4 Then
            var_c2 = var_c + "000000"
            rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
            var_nivel = rs(8).Value
            var_nivel1 = rs(2).Value
            var_clave_nivel1 = rs(2).Value
            var_nombre_submenu = rs(7).Value
            var_accion_submenu = rs(9).Value
            var_global_permiso1 = rs!inte_sme_permiso1
            var_global_permiso2 = rs!inte_sme_permiso2
            var_global_permiso3 = rs!inte_sme_permiso3
            var_global_permiso4 = rs!inte_sme_permiso4
            rs.Close
         End If
         If var_longitud = 6 Then
            var_c2 = var_c + "0000"
            If rs.State = 1 Then
               rs.Close
            End If
            cnn.CommandTimeout = 360
            rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
            var_nivel = rs(8).Value
            var_nivel1 = rs(2).Value
            var_nivel2 = rs(3).Value
            var_clave_nivel1 = rs(2).Value
            var_clave_nivel2 = rs(3).Value
            var_accion_submenu = rs(9).Value
            var_nombre_submenu = rs(7).Value
            var_global_permiso1 = rs!inte_sme_permiso1
            var_global_permiso2 = rs!inte_sme_permiso2
            var_global_permiso3 = rs!inte_sme_permiso3
            var_global_permiso4 = rs!inte_sme_permiso4
            rs.Close
         End If
         If var_longitud = 8 Then
            var_c2 = var_c + "00"
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
            var_nivel = rs(8).Value
            var_nivel1 = rs(2).Value
            var_nivel2 = rs(3).Value
            var_nivel3 = rs(4).Value
            var_accion_submenu = rs(9).Value
            var_nombre_submenu = rs(7).Value
            var_global_permiso1 = rs!inte_sme_permiso1
            var_global_permiso2 = rs!inte_sme_permiso2
            var_global_permiso3 = rs!inte_sme_permiso3
            var_global_permiso4 = rs!inte_sme_permiso4
            rs.Close
         End If
         If var_longitud = 10 Then
            var_c2 = var_c
            rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
            var_nivel = rs(8).Value
            var_nivel1 = rs(2).Value
            var_nivel2 = rs(3).Value
            var_nivel3 = rs(4).Value
            var_accion_submenu = rs(9).Value
            var_nombre_submenu = rs(7).Value
            var_global_permiso1 = rs!inte_sme_permiso1
            var_global_permiso2 = rs!inte_sme_permiso2
            var_global_permiso3 = rs!inte_sme_permiso3
            var_global_permiso4 = rs!inte_sme_permiso4
            rs.Close
         End If
         var_opcion_seguridad = 1
         If var_global_permiso1 = 1 Then
            If var_global_permiso2 = 1 Then
               frmpasswords2.Show 1
            Else
               frmpasswords.Show 1
            End If
         Else
            ejecuta_forma
         End If
      End If
   End If
End Sub

