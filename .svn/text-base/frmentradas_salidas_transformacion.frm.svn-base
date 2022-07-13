VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmentradas_salidas_transformacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   525
      TabIndex        =   25
      Top             =   1095
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         TabIndex        =   26
         Top             =   495
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   30
         TabIndex        =   27
         Top             =   120
         Width           =   3075
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1215
      TabIndex        =   22
      Top             =   585
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   23
         Top             =   480
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3228
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
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   24
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame3 
      Height          =   780
      Left            =   120
      TabIndex        =   12
      Top             =   1095
      Width           =   11370
      Begin VB.TextBox txt_folio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   8685
         TabIndex        =   18
         Top             =   150
         Width           =   2610
      End
      Begin VB.CommandButton cmd_cargar_archivo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6690
         Picture         =   "frmentradas_salidas_transformacion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Cargar archivo de excel"
         Top             =   270
         Width           =   330
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   375
         Left            =   2160
         TabIndex        =   15
         Top             =   240
         Width           =   4485
      End
      Begin VB.TextBox txt_almacen 
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         Height          =   195
         Left            =   8220
         TabIndex        =   17
         Top             =   330
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Left            =   225
         TabIndex        =   13
         Top             =   330
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5325
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   11370
      Begin VB.CommandButton cmd_cargar_movimientos 
         Caption         =   "Cargar movimiento"
         Height          =   405
         Left            =   105
         TabIndex        =   21
         Top             =   4815
         Width           =   1680
      End
      Begin VB.TextBox txt_cantidad 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8925
         TabIndex        =   20
         Top             =   4785
         Width           =   2355
      End
      Begin MSComctlLib.ListView lv_entradas 
         Height          =   4350
         Left            =   90
         TabIndex        =   10
         Top             =   405
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   7673
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6527
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Convierte"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   6527
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Cantidad"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   8205
         TabIndex        =   19
         Top             =   4890
         Width           =   675
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   11
         Top             =   120
         Width           =   11280
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   105
      TabIndex        =   6
      Top             =   585
      Width           =   11460
   End
   Begin VB.TextBox txt_clave_movimiento 
      Height          =   285
      Left            =   2250
      TabIndex        =   5
      Top             =   750
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txt_tipo_documento 
      Height          =   285
      Left            =   3135
      TabIndex        =   4
      Top             =   750
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   180
      Picture         =   "frmentradas_salidas_transformacion.frx":0312
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Nuevo Movimiento Alt + N"
      Top             =   705
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   510
      Picture         =   "frmentradas_salidas_transformacion.frx":0414
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Buscar Movimiento Alt + B"
      Top             =   705
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Picture         =   "frmentradas_salidas_transformacion.frx":0516
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Movimiento Alt + I"
      Top             =   705
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11160
      Picture         =   "frmentradas_salidas_transformacion.frx":0618
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   720
      Width           =   330
   End
   Begin MSComDlg.CommonDialog cmdentradas 
      Left            =   2625
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Busqueda de archivo"
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   60
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_salidas_transformacion.frx":0C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_salidas_transformacion.frx":152C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_salidas_transformacion.frx":1E06
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_salidas_transformacion.frx":23A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_salidas_transformacion.frx":2C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_salidas_transformacion.frx":3558
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_salidas_transformacion.frx":3E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_salidas_transformacion.frx":3F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_salidas_transformacion.frx":4056
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_salidas_transformacion.frx":4168
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_salidas_transformacion.frx":427A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_salidas_transformacion.frx":438C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   0
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   810
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   90
      TabIndex        =   7
      Top             =   960
      Width           =   11475
   End
   Begin VB.Label lblnombremovimiento 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   90
      TabIndex        =   8
      Top             =   75
      Width           =   11445
   End
End
Attribute VB_Name = "frmentradas_salidas_transformacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_estatus_movimiento As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report


Private Sub cmd_buscar_Click()
   Me.frm_busqueda.Visible = True
   Me.txt_busqueda_folio = ""
   Me.txt_busqueda_folio.SetFocus
End Sub

Private Sub cmd_cargar_archivo_Click()
'On Error GoTo salir:
   If Me.txt_folio = "" Then
      Me.lv_entradas.ListItems.Clear
      frmbusqueda_archivo.Show 1
      If Trim(var_archivo_buscar) <> "" Then
         strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls); DBQ=" & var_archivo_buscar
         If rsaux2.State = 1 Then
            rsaux2.Close
         End If
         
         rsaux2.Open "SELECT codigo1, codigo2, cantidad FROM [HOJA1$]", strConnectionString
         var_cantidad_total = 0
         While Not rsaux2.EOF
               var_ctex = CStr(IIf(IsNull(rsaux2!codigo2), "", rsaux2!codigo2))
               If var_ctex <> "" Then
                  rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + CStr(IIf(IsNull(rsaux2!codigo1), "", rsaux2!codigo1)) + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_descripcion_1 = ""
                  If Not rs.EOF Then
                     var_descripcion_1 = IIf(IsNull(rs!vcha_art_nombre_Español), "", rs!vcha_art_nombre_Español)
                  End If
                  rs.Close
                  rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + IIf(IsNull(CStr(IIf(IsNull(rsaux2!codigo2), "", rsaux2!codigo2))), "", CStr(IIf(IsNull(rsaux2!codigo2), "", rsaux2!codigo2))) + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_descripcion_2 = ""
                  If Not rs.EOF Then
                     var_descripcion_2 = IIf(IsNull(rs!vcha_art_nombre_Español), "", rs!vcha_art_nombre_Español)
                  End If
                  rs.Close
                  Set list_item = lv_entradas.ListItems.Add(, , IIf(IsNull(rsaux2!codigo1), "", rsaux2!codigo1))
                  list_item.SubItems(1) = var_descripcion_1
                  list_item.SubItems(2) = IIf(IsNull(rsaux2!codigo2), 0, rsaux2!codigo2)
                  list_item.SubItems(3) = var_descripcion_2
                  list_item.SubItems(4) = IIf(IsNull(rsaux2!Cantidad), 0, rsaux2!Cantidad)
                  var_cantidad_total = var_cantidad_total + IIf(IsNull(rsaux2!Cantidad), 0, rsaux2!Cantidad)
               End If
               rsaux2.MoveNext
         Wend
         Me.txt_cantidad = Format(var_cantidad_total, "###,###,##.00")
         rsaux2.Close
         Me.cmd_cargar_movimientos.Enabled = True
      End If
   Else
      MsgBox "El movimiento ya esta cargado", vbOKOnly, "ATENCION"
      Me.cmd_cargar_movimientos.Enabled = False
   End If
   Exit Sub
salir:
 MsgBox "A surgido un error al cargar el archivo, puede que no tenga el formato correcto. Debe de ser CODIGO1, CODIGO2, CANTIDAD y la hoja debe de llamarse Hoja1", vbOKOnly, "ATENCION"
 Me.cmd_cargar_movimientos.Enabled = False
End Sub

Private Sub cmd_cargar_movimientos_Click()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   Dim var_inserta As Boolean
   
   If Me.txt_almacen <> "" Then
      var_cadena_articulos = ""
      If Me.txt_folio = "" Then
         For var_j = 1 To Me.lv_entradas.ListItems.Count
             Me.lv_entradas.ListItems(var_j).Selected = True
             If Trim(Me.lv_entradas.selectedItem.SubItems(1)) = "" Then
                If Trim(var_cadena_articulos) = "" Then
                   var_cadena_articulos = var_cadena_articulos + Me.lv_entradas.selectedItem
                Else
                   var_cadena_articulos = var_cadena_articulos + ", " + Me.lv_entradas.selectedItem
                End If
             End If
             Me.lv_entradas.ListItems(var_j).Selected = True
             If Trim(Me.lv_entradas.selectedItem.SubItems(3)) = "" Then
                If Trim(var_cadena_articulos) = "" Then
                   var_cadena_articulos = var_cadena_articulos + Me.lv_entradas.selectedItem
                Else
                   var_cadena_articulos = var_cadena_articulos + ", " + Me.lv_entradas.selectedItem.SubItems(3)
                End If
             End If
         Next var_j
         If var_cadena_articulos = "" Then
            var_año = 2005
            var_primera_vez = True
            For var_j = 1 To Me.lv_entradas.ListItems.Count
               Me.lv_entradas.ListItems.Item(var_j).Selected = True
               txt_codigo = Trim(Me.lv_entradas.selectedItem)
               If Trim(txt_codigo) <> "" Then
                  cnn.CommandTimeout = 360
                  bandera_suma = False
                  If var_primera_vez = True Then
                     var_inserta = False
                     var_numero_folio = 0
                     var_insreta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, Me.txt_almacen, var_clave_movimiento, Now, CDbl(var_numero_folio), 0, "", "", Me.txt_almacen, Me.txt_almacen, "", var_clave_usuario_global, fun_NombrePc, "", "", "", "", "B", "", "", 0, 0, 0, "1", 1)
                     var_numero_folio = var_numero_folio_regreso
                     txt_folio = var_numero_folio
                     var_primera_vez = False
                  End If
                  rsaux.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + Trim(Me.lv_entradas.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_costo = IIf(IsNull(rsaux!FLOA_eXI_COSTO), 0, rsaux!FLOA_eXI_COSTO)
                     If var_costo = 0 Then
                        rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.lv_entradas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_costo = IIf(IsNull(rs!mone_Art_costo_estandar), 0, rs!mone_Art_costo_estandar)
                        var_precio = IIf(IsNull(rs!mone_Art_precio_base), 0, rs!mone_Art_precio_base)
                        rs.Close
                     Else
                        rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.lv_entradas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_precio = IIf(IsNull(rs!mone_Art_precio_base), 0, rs!mone_Art_precio_base)
                        rs.Close
                     End If
                  Else
                     rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.lv_entradas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_costo = IIf(IsNull(rs!mone_Art_costo_estandar), 0, rs!mone_Art_costo_estandar)
                     var_precio = IIf(IsNull(rs!mone_Art_precio_base), 0, rs!mone_Art_precio_base)
                     rs.Close
                  End If
                  rsaux.Close
                  
               
                  rsaux.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + Trim(Me.lv_entradas.selectedItem.SubItems(2)) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     VAR_COSTO_ENTRADA = IIf(IsNull(rsaux!FLOA_eXI_COSTO), 0, rsaux!FLOA_eXI_COSTO)
                     If VAR_COSTO_ENTRADA = 0 Then
                        rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.lv_entradas.selectedItem.SubItems(2) + "'", cnn, adOpenDynamic, adLockOptimistic
                        VAR_COSTO_ENTRADA = IIf(IsNull(rs!mone_Art_costo_estandar), 0, rs!mone_Art_costo_estandar)
                        VAR_PRECIO_ENTRADA = IIf(IsNull(rs!mone_Art_precio_base), 0, rs!mone_Art_precio_base)
                        rs.Close
                     Else
                        rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.lv_entradas.selectedItem.SubItems(2) + "'", cnn, adOpenDynamic, adLockOptimistic
                        VAR_PRECIO_ENTRADA = IIf(IsNull(rs!mone_Art_precio_base), 0, rs!mone_Art_precio_base)
                        rs.Close
                     End If
                  Else
                     rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.lv_entradas.selectedItem.SubItems(2) + "'", cnn, adOpenDynamic, adLockOptimistic
                     'MsgBox Me.lv_entradas.selectedItem.SubItems(2)
                     If Not rs.EOF Then
                        VAR_COSTO_ENTRADA = IIf(IsNull(rs!mone_Art_costo_estandar), 0, rs!mone_Art_costo_estandar)
                        VAR_PRECIO_ENTRADA = IIf(IsNull(rs!mone_Art_precio_base), 0, rs!mone_Art_precio_base)
                     Else
                        VAR_COSTO_ENTRADA = 0
                        VAR_PRECIO_ENTRADA = 0
                     End If
                     rs.Close
                  End If
                  rsaux.Close
                  
                  
                  var_cadena = "insert into tb_temporal_salidas (vcha_Emp_empresa_id, vcha_uor_unidad_id,                    vcha_Alm_almacen_id,      vcha_mov_movimiento_id,                  inte_Sal_numero,      vcha_Art_Articulo_id,                        floa_sal_costo,            floa_sal_precio, floa_sal_cantidad, VCHA_SAL_CODIGO_ENTRADA, FLOA_SAL_COSTO_eNTRADA, FLOA_SAL_PRECIO_ENTRADA)"
                  var_cadena = var_cadena + " values            ('" + var_empresa + "',   '" + var_unidad_organizacional + "', '" + Me.txt_almacen + "', '" + var_clave_movimiento + "'," + CStr(var_numero_folio) + ",'" + Trim(Me.lv_entradas.selectedItem) + "', " + CStr(var_costo) + "," + CStr(var_precio) + "," + CStr(CDbl(Me.lv_entradas.selectedItem.SubItems(4))) + ",'" + Trim(Me.lv_entradas.selectedItem.SubItems(2)) + "'," + CStr(var_costo) + ", " + CStr(VAR_PRECIO_ENTRADA) + ")"
                  rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               End If
            Next var_j
         Else
            MsgBox "Los siguientes códigos no estan dados de alta en el SID " + var_cadena_articulos, vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El movimiento ya fue cargado con anterioridad", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un almacén", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
   Dim var_posible_cerrar_movimiento As Integer
   var_posible_cerrar_movimiento = 1
   
   Dim dl As Long                                 ' Valor devuelto por la función API
   Dim sAttributes As String                  ' Aributos
   Dim sDriver As String                       ' Nombre del controlador
   Dim sDescription As String                ' Descripción del DSN
   Dim sDsnName As String                  ' Nombre del DSN

   Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
   Const vbAPINull As Long = 0&                         ' Puntero NULL

   ' se elimina
   Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema
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
   sAttributes = sAttributes & "Database=" + var_bd_movimientos & Chr(0)
   strAttributes = strAttributes & "UID=sa" & Chr$(0)
   strAttributes = strAttributes & "PWD=elia" & Chr$(0)
   dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_ENTRADAS_INSERTA = New TB_ENTRADAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Dim var_correo_electronico As String
   Dim var_acepta_traspaso As Integer
   If IsNumeric(Me.txt_folio) Then
      var_numero_folio = CDbl(Me.txt_folio)
   Else
      var_numero_folio = 0
   End If
   If var_numero_folio > 0 Then
      var_correo_estado_cuenta = 0
      var_almacen_origen = Me.txt_almacen
      var_almacen_Destino = Me.txt_almacen
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         Set reporte = appl.OpenReport(App.Path + "\REP_eNTRADAS_SALIDAS_TRANSFORMACION.rpt")
         reporte.RecordSelectionFormula = "{VW_ENTRADAS_SALIDAS_TRANSFORMACION.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and {VW_ENTRADAS_SALIDAS_TRANSFORMACION.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' and {VW_ENTRADAS_SALIDAS_TRANSFORMACION.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_ENTRADAS_SALIDAS_TRANSFORMACION.INTE_SAL_NUMERO} = " + Me.txt_folio
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Movimientos"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
      Else
         
         var_acepta_traspaso = 1
         If var_acepta_traspaso = 1 Then
            var_si = MsgBox("¿Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
            If var_si = 1 Then
               var_posible_traspaso = 0
               If var_posible_traspaso = 0 Then
                  Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
                  cnn.BeginTrans
                  rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  While Not rs.EOF
                        var_año = 2005
                        rsaux4.Open "insert into tb_salidas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_sal_numero, vcha_art_articulo_id, floa_sal_cantidad, floa_sal_costo, floa_sal_precio, inte_sal_año,VCHA_SAL_CODIGO_ENTRADA) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_origen + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_Articulo_id + "', " + CStr(rs!floa_Sal_Cantidad) + ", " + CStr(rs!floa_Sal_costo) + " , " + CStr(rs!floa_Sal_precio) + ", " + CStr(var_año) + ",'" + rs!VCHA_SAL_CODIGO_ENTRADA + "')", cnn, adOpenDynamic, adLockOptimistic
                        rsaux4.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO, VCHA_ENT_ALMACEN_ORIGEN) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!VCHA_SAL_CODIGO_ENTRADA + "', " + CStr(rs!floa_Sal_Cantidad) + ", " + CStr(IIf(IsNull(rs!FLOA_SAL_COSTO_eNTRADA), 0, rs!FLOA_SAL_COSTO_eNTRADA)) + " , " + CStr(IIf(IsNull(rs!FLOA_SAL_PRECIO_ENTRADA), 0, rs!FLOA_SAL_PRECIO_ENTRADA)) + ", " + CStr(var_año) + ", '" + var_almacen_origen + "')", cnn, adOpenDynamic, adLockOptimistic
                        rs.MoveNext
                  Wend
                  rs.Close
                  var_estatus_movimiento = "I"
                  var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, CStr(var_almacen_origen), var_clave_movimiento, CDbl(var_numero_folio), "I", Now, 1)
                  cnn.CommitTrans
                  Set reporte = appl.OpenReport(App.Path + "\REP_eNTRADAS_SALIDAS_TRANSFORMACION.rpt")
                  reporte.RecordSelectionFormula = "{VW_ENTRADAS_SALIDAS_TRANSFORMACION.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and {VW_ENTRADAS_SALIDAS_TRANSFORMACION.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' and {VW_ENTRADAS_SALIDAS_TRANSFORMACION.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_ENTRADAS_SALIDAS_TRANSFORMACION.INTE_SAL_NUMERO} = " + Me.txt_folio
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Movimientos"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  var_estatus_movimiento = "I"
               Else
                  MsgBox var_cadena_faltantes
               End If 'var_posible_traspaso
            End If
         Else
            MsgBox "No se puede cerrar el traspaso hasta que haya una autorización del almacén " + txt_nombre_almacen_destino, vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   var_estatus_movimiento = ""
   Me.txt_almacen.Enabled = True
   Me.txt_almacen = ""
   Me.txt_nombre_almacen = ""
   Me.txt_folio = ""
   Me.txt_cantidad = ""
   Me.lv_entradas.ListItems.Clear
   Me.txt_almacen.SetFocus
   Me.cmd_cargar_movimientos.Enabled = False
End Sub

Private Sub Command1_Click()
   frmbusqueda_archivo.Show 1
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_estatus_movimiento = ""
   Top = 0
   Left = 0
   Me.frm_lista.Visible = False
   Me.frm_busqueda.Visible = False
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_traspasos)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_entradas.ListItems.Count >= 0 Then
         Me.txt_almacen = Me.lv_lista.selectedItem
         Me.txt_nombre_almacen = Me.lv_lista.selectedItem.SubItems(1)
         Me.txt_almacen.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.txt_almacen.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      End If
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Almacenes"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre_almacen.Enabled = True
      Me.txt_nombre_almacen.SetFocus
   End If
End Sub

Private Sub txt_almacen_LostFocus()
   If Trim(Me.txt_almacen) <> "" Then
   If Trim(txt_almacen) <> "" Then
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_alm_almacen_id = '" + txt_almacen + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_Alm_almacen_id = '" + txt_almacen + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      End If
      If Not rs.EOF Then
         txt_nombre_almacen = rs!VCHA_ALM_NOMBRE
         txt_almacen.Enabled = False
         txt_nombre_almacen.Enabled = False
      Else
         Me.txt_nombre_almacen = ""
         Me.txt_almacen = ""
         MsgBox "Clave de almacen incorrecto", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
   Else
      Me.txt_nombre_almacen = ""
   End If
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_busqueda_folio) Then
         var_cadena = "SELECT     dbo.TB_TEMPORAL_SALIDAS.VCHA_EMP_EMPRESA_ID, dbo.TB_TEMPORAL_SALIDAS.VCHA_UOR_UNIDAD_ID, dbo.TB_TEMPORAL_SALIDAS.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_TEMPORAL_SALIDAS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_TEMPORAL_SALIDAS.INTE_SAL_NUMERO, dbo.TB_TEMPORAL_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_TEMPORAL_SALIDAS.FLOA_SAL_CANTIDAD, dbo.TB_TEMPORAL_SALIDAS.VCHA_SAL_CODIGO_ENTRADA, TB_ARTICULOS_1.VCHA_ART_NOMBRE_ESPAÑOL AS DESCRIPCION_CODIGO_eNTRADA FROM dbo.TB_TEMPORAL_SALIDAS INNER JOIN dbo.TB_ALMACENES ON dbo.TB_TEMPORAL_SALIDAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_TEMPORAL_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_ARTICULOS TB_ARTICULOS_1 ON "
         var_cadena = var_cadena + " dbo.TB_TEMPORAL_SALIDAS.VCHA_SAL_CODIGO_ENTRADA = TB_ARTICULOS_1.VCHA_ART_ARTICULO_ID WHERE     (dbo.TB_TEMPORAL_SALIDAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_TEMPORAL_SALIDAS.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') AND (dbo.TB_TEMPORAL_SALIDAS.INTE_SAL_NUMERO = " + Me.txt_busqueda_folio + ") AND (dbo.TB_TEMPORAL_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = 'EST') "
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            lv_entradas.ListItems.Clear
            Me.txt_folio = rs!INTE_SAL_NUMERO
            Me.txt_almacen = rs!VCHA_ALM_ALMACEN_ID
            Me.txt_nombre_almacen = rs!VCHA_ALM_NOMBRE
            var_cantidad_total = 0
            While Not rs.EOF
                  Set list_item = lv_entradas.ListItems.Add(, , IIf(IsNull(rs!vcha_Art_Articulo_id), "", rs!vcha_Art_Articulo_id))
                  list_item.SubItems(1) = rs!vcha_art_nombre_Español
                  list_item.SubItems(2) = IIf(IsNull(rs!VCHA_SAL_CODIGO_ENTRADA), 0, rs!VCHA_SAL_CODIGO_ENTRADA)
                  list_item.SubItems(3) = rs!DESCRIPCION_CODIGO_eNTRADA
                  list_item.SubItems(4) = IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad)
                  var_cantidad_total = var_cantidad_total + IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad)
                  rs.MoveNext
            Wend
            Me.txt_cantidad = Format(var_cantidad_total, "###,###,##.00")
            rsaux.Open "SELECT * FROM TB_ENCABEZADO_MOVIMIENTOS WHERE VCHA_EMP_eMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = 'EST' AND INTE_EMO_NUMERO = " + Me.txt_busqueda_folio, cnn, adOpenDynamic, adLockOptimistic
            var_estatus_movimiento = IIf(IsNull(rsaux!char_Emo_estatus), "", rsaux!char_Emo_estatus)
            rsaux.Close
            Me.cmd_cargar_movimientos.Enabled = False
         Else
            MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
         Me.frm_busqueda.Visible = False
      Else
         MsgBox "El número de movimiento es incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_busqueda_folio_LostFocus()
   Me.frm_busqueda.Visible = False
End Sub

Private Sub txt_nombre_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If Me.txt_almacen.Enabled = True Then
      If KeyCode = 116 Then
         lv_lista.ListItems.Clear
         If var_tipo_permiso = 1 Then
            rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
         Else
            rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
         End If
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "Almacenes"
         var_tipo_lista = 1
         Dim var_n As Integer
         var_n = lv_lista.ListItems.Count
         If var_n > 6 Then
            lv_lista.ColumnHeaders(2).Width = 4270.71
         Else
            lv_lista.ColumnHeaders(2).Width = 4499.71
         End If
         frm_lista.Visible = True
         lv_lista.SetFocus
      End If
   
   End If
End Sub

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_cargar_archivo.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub
