VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmprincipal_configuracion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración del sistema"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   2085
      TabIndex        =   39
      Top             =   750
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   40
         Top             =   495
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
         TabIndex        =   41
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   90
      TabIndex        =   38
      Top             =   390
      Width           =   8640
   End
   Begin VB.CommandButton cmd_aceptar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   165
      Picture         =   "frmprincipal_configuracion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   495
      Picture         =   "frmprincipal_configuracion.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancelar Esc"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos principales de configuración "
      Height          =   6345
      Left            =   135
      TabIndex        =   0
      Top             =   465
      Width           =   8595
      Begin VB.TextBox txt_ruta_archivos_reempaque 
         Height          =   315
         Left            =   3345
         TabIndex        =   20
         Top             =   5850
         Width           =   5145
      End
      Begin VB.TextBox txt_ruta_archivos_enviar 
         Height          =   315
         Left            =   3345
         TabIndex        =   19
         Top             =   5505
         Width           =   5145
      End
      Begin VB.TextBox txt_dias_tolerancia 
         Height          =   315
         Left            =   3345
         TabIndex        =   18
         Top             =   5160
         Width           =   1125
      End
      Begin VB.TextBox txt_tipo_calculo_factura_catalogos 
         Height          =   315
         Left            =   3345
         TabIndex        =   17
         Top             =   4815
         Width           =   1125
      End
      Begin VB.TextBox txt_ruta_envio_facturas 
         Height          =   315
         Left            =   3360
         TabIndex        =   16
         Top             =   4470
         Width           =   5145
      End
      Begin VB.TextBox txt_ruta_archivos_articulos 
         Height          =   315
         Left            =   3345
         TabIndex        =   15
         Top             =   4125
         Width           =   5145
      End
      Begin VB.TextBox txt_ruta_archivos_pedidos_sugerido 
         Height          =   315
         Left            =   3345
         TabIndex        =   14
         Top             =   3780
         Width           =   5145
      End
      Begin VB.TextBox txt_tipo_agrupamiento 
         Height          =   315
         Left            =   3345
         TabIndex        =   13
         Top             =   3435
         Width           =   1125
      End
      Begin VB.TextBox txt_tolerancia_saldos 
         Height          =   315
         Left            =   3345
         TabIndex        =   12
         Top             =   3090
         Width           =   1125
      End
      Begin VB.TextBox txt_renglones_nota_credito 
         Height          =   315
         Left            =   3345
         TabIndex        =   6
         Top             =   1020
         Width           =   1125
      End
      Begin VB.TextBox txt_ruta_cobranza 
         Height          =   315
         Left            =   3345
         TabIndex        =   11
         Top             =   2745
         Width           =   5145
      End
      Begin VB.TextBox txt_ruta_almacenes 
         Height          =   315
         Left            =   3345
         TabIndex        =   10
         Top             =   2400
         Width           =   5145
      End
      Begin VB.TextBox txt_ruta_pedidos 
         Height          =   315
         Left            =   3345
         TabIndex        =   9
         Top             =   2055
         Width           =   5145
      End
      Begin VB.TextBox txt_ruta_devoluciones_tiendas 
         Height          =   315
         Left            =   3345
         TabIndex        =   8
         Top             =   1710
         Width           =   5145
      End
      Begin VB.TextBox txt_ruta_nota_envios_plantas 
         Height          =   315
         Left            =   3345
         TabIndex        =   7
         Top             =   1365
         Width           =   5145
      End
      Begin VB.TextBox txt_renglones_factura 
         Height          =   315
         Left            =   3345
         TabIndex        =   5
         Top             =   675
         Width           =   1125
      End
      Begin VB.TextBox txt_nombre_empresa 
         Height          =   315
         Left            =   4500
         TabIndex        =   4
         Top             =   345
         Width           =   3960
      End
      Begin VB.TextBox txt_empresa 
         Height          =   315
         Left            =   3345
         TabIndex        =   3
         Top             =   345
         Width           =   1125
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Archivos Reempaque:"
         Height          =   195
         Left            =   390
         TabIndex        =   37
         Top             =   5910
         Width           =   1965
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Archivos Enviar:"
         Height          =   195
         Left            =   390
         TabIndex        =   36
         Top             =   5565
         Width           =   1545
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Dias Tolerancia Facturación Catálogos:"
         Height          =   195
         Left            =   390
         TabIndex        =   35
         Top             =   5220
         Width           =   2790
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Calculo Factura Catalogos:"
         Height          =   195
         Left            =   390
         TabIndex        =   34
         Top             =   4875
         Width           =   2265
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Envios Facturas:"
         Height          =   195
         Left            =   390
         TabIndex        =   33
         Top             =   4530
         Width           =   1575
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Archivos de Articulos:"
         Height          =   195
         Left            =   390
         TabIndex        =   32
         Top             =   4185
         Width           =   1920
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Archivos Pedido Sugerido:"
         Height          =   195
         Left            =   390
         TabIndex        =   31
         Top             =   3840
         Width           =   2265
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Agrupamiento:"
         Height          =   195
         Left            =   390
         TabIndex        =   30
         Top             =   3495
         Width           =   1380
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tolerancia Saldos:"
         Height          =   195
         Left            =   390
         TabIndex        =   29
         Top             =   3150
         Width           =   1320
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Renglones Nota Credito:"
         Height          =   195
         Left            =   390
         TabIndex        =   28
         Top             =   1080
         Width           =   1740
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Cobranza:"
         Height          =   195
         Left            =   390
         TabIndex        =   27
         Top             =   2805
         Width           =   1110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Almacenes:"
         Height          =   195
         Left            =   390
         TabIndex        =   26
         Top             =   2460
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Pedidos:"
         Height          =   195
         Left            =   390
         TabIndex        =   25
         Top             =   2115
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Devoluciones de Tiendas:"
         Height          =   195
         Left            =   390
         TabIndex        =   24
         Top             =   1770
         Width           =   2250
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Notas de Envio de plantas:"
         Height          =   195
         Left            =   390
         TabIndex        =   23
         Top             =   1425
         Width           =   2310
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Renglones Factura:"
         Height          =   195
         Left            =   390
         TabIndex        =   22
         Top             =   735
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         Height          =   195
         Left            =   390
         TabIndex        =   21
         Top             =   390
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmprincipal_configuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_Click()
   If Trim(txt_empresa) <> "" Then
      rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + txt_empresa + "'"
      If rs.EOF Then
         var_si = MsgBox("Desea insertar la información", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_cadena = "INSERT INTO TB_PRINCIPAL (VCHA_PRI_RUTA_NOTAS_ENVIO, VCHA_PRI_RUTA_DEVOLUCIONES_TIENDA, VCHA_PRI_RUTA_PEDIDOS, INTE_PRI_RENGLONES_FACTURA, INTE_PRI_FACTURA, INTE_PRI_NOTA_CREDITO, INTE_PRI_NOTA_CARGO, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_PRI_RUTA_ALMACENES, VCHA_PRI_RUTA_COBRANZA, INTE_PRI_RENGLONES_NOTA_CREDITO, FLOA_PRI_TOLERANCIA_SALDOS, CHAR_PRI_TIPO_AGRUPAMIENTO, VCHA_PRI_RUTA_PEDIDO_SUGERIDO, VCHA_PRI_RUTA_ARCHIVOS_ARTICULOS, VCHA_PRI_RUTA_ENVIOS_FACTURAS, CHAR_PRI_TIPO_CALCULO_FACTURA_CATALOGOS, INTE_PRI_DIAS_TOLERANCIA_FACTURACION_CATALOGOS, VCHA_PRI_RUTA_ARCHIVOS_ENVIAR, VCHA_PRI_RUTA_REEMPAQUE)"
            var_cadena = var_cadena + " Values ('" + Me.txt_ruta_nota_envios_plantas + "','" + Me.txt_ruta_devoluciones_tiendas + "', '" + Me.txt_ruta_pedidos + "', " + Me.txt_renglones_factura + ", 0,0,0, " + txt_empresa + ", '', '" + Me.txt_ruta_almacenes + "', '" + Me.txt_ruta_cobranza + "', " + Me.txt_renglones_nota_credito + ", " + Me.txt_tolerancia_saldos + ", '" + Me.txt_tipo_agrupamiento + "','" + Me.txt_ruta_archivos_pedidos_sugerido + "', '" + Me.txt_ruta_archivos_articulos + "', '" + Me.txt_ruta_envio_facturas + "', '" + Me.txt_tipo_calculo_factura_catalogos + "', " + Me.txt_dias_tolerancia + ", '" + Me.txt_ruta_archivos_enviar + "', '" + Me.txt_ruta_archivos_reempaque + "')"
            rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         End If
      Else
         var_si = MsgBox("Desea actualizar los datos de la configuracion", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rsaux.Open "DELETE FROM TB_PRINCIPAL WHERE VCHA_EMP_EMPRESA_ID = '" + txt_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            var_cadena = "INSERT INTO TB_PRINCIPAL (VCHA_PRI_RUTA_NOTAS_ENVIO, VCHA_PRI_RUTA_DEVOLUCIONES_TIENDA, VCHA_PRI_RUTA_PEDIDOS, INTE_PRI_RENGLONES_FACTURA, INTE_PRI_FACTURA, INTE_PRI_NOTA_CREDITO, INTE_PRI_NOTA_CARGO, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_PRI_RUTA_ALMACENES, VCHA_PRI_RUTA_COBRANZA, INTE_PRI_RENGLONES_NOTA_CREDITO, FLOA_PRI_TOLERANCIA_SALDOS, CHAR_PRI_TIPO_AGRUPAMIENTO, VCHA_PRI_RUTA_PEDIDO_SUGERIDO, VCHA_PRI_RUTA_ARCHIVOS_ARTICULOS, VCHA_PRI_RUTA_ENVIOS_FACTURAS, CHAR_PRI_TIPO_CALCULO_FACTURA_CATALOGOS, INTE_PRI_DIAS_TOLERANCIA_FACTURACION_CATALOGOS, VCHA_PRI_RUTA_ARCHIVOS_ENVIAR, VCHA_PRI_RUTA_REEMPAQUE)"
            var_cadena = var_cadena + " Values ('" + Me.txt_ruta_nota_envios_plantas + "','" + Me.txt_ruta_devoluciones_tiendas + "', '" + Me.txt_ruta_pedidos + "', " + Me.txt_renglones_factura + ", 0,0,0, '" + txt_empresa + "', '', '" + Me.txt_ruta_almacenes + "', '" + Me.txt_ruta_cobranza + "', " + Me.txt_renglones_nota_credito + ", " + Me.txt_tolerancia_saldos + ", '" + Me.txt_tipo_agrupamiento + "','" + Me.txt_ruta_archivos_pedidos_sugerido + "', '" + Me.txt_ruta_archivos_articulos + "', '" + Me.txt_ruta_envio_facturas + "', '" + Me.txt_tipo_calculo_factura_catalogos + "', " + Me.txt_dias_tolerancia + ", '" + Me.txt_ruta_archivos_enviar + "', '" + Me.txt_ruta_archivos_reempaque + "')"
            rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         End If
      End If
      rs.Close
   Else
      MsgBox "No se a seleccionado una empresa", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_cancelar_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 300
   Left = 1300
   frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_principal_configuracion)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_empresa = lv_lista.selectedItem
      txt_nombre_empresa = lv_lista.selectedItem.SubItems(1)
      txt_empresa.SetFocus
   End If
   If KeyAscii = 13 Then
      txt_empresa.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_dias_tolerancia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_empresa_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_EMPRESAS order by vcha_emp_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_emp_empresa_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_emp_nombre), "", rs!vcha_emp_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "EMPRESAS"
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

Private Sub txt_empresa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_empresa_LostFocus()
   If Trim(txt_empresa) <> "" Then
      rs.Open "SELECT * FROM TB_EMPRESAS WHERE VCHA_EMP_EMPRESA_ID = '" + txt_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_empresa = IIf(IsNull(rs!vcha_emp_nombre), "", rs!vcha_emp_nombre)
         rsaux.Open "select * from tb_principal where vcha_emp_empresa_id = '" + txt_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            Me.txt_renglones_factura = IIf(IsNull(rsaux!INTE_PRI_RENGLONES_FACTURA), 0, rsaux!INTE_PRI_RENGLONES_FACTURA)
            Me.txt_renglones_nota_credito = IIf(IsNull(rsaux!INTE_PRI_RENGLONES_NOTA_CREDITO), 0, rsaux!INTE_PRI_RENGLONES_NOTA_CREDITO)
            Me.txt_ruta_nota_envios_plantas = IIf(IsNull(rsaux!VCHA_PRI_RUTA_NOTAS_ENVIO), "", rsaux!VCHA_PRI_RUTA_NOTAS_ENVIO)
            Me.txt_ruta_devoluciones_tiendas = IIf(IsNull(rsaux!VCHA_PRI_RUTA_DEVOLUCIONES_TIENDA), "", rsaux!VCHA_PRI_RUTA_DEVOLUCIONES_TIENDA)
            Me.txt_ruta_pedidos = IIf(IsNull(rsaux!VCHA_PRI_RUTA_PEDIDOS), "", rsaux!VCHA_PRI_RUTA_PEDIDOS)
            Me.txt_ruta_almacenes = IIf(IsNull(rsaux!VCHA_PRI_RUTA_ALMACENES), "", rsaux!VCHA_PRI_RUTA_ALMACENES)
            Me.txt_ruta_cobranza = IIf(IsNull(rsaux!VCHA_PRI_RUTA_COBRANZA), "", rsaux!VCHA_PRI_RUTA_COBRANZA)
            Me.txt_tolerancia_saldos = IIf(IsNull(rsaux!FLOA_PRI_TOLERANCIA_SALDOS), 0, rsaux!FLOA_PRI_TOLERANCIA_SALDOS)
            Me.txt_tipo_agrupamiento = IIf(IsNull(rsaux!CHAR_PRI_TIPO_AGRUPAMIENTO), "", rsaux!CHAR_PRI_TIPO_AGRUPAMIENTO)
            Me.txt_ruta_archivos_pedidos_sugerido = IIf(IsNull(rsaux!VCHA_PRI_RUTA_PEDIDO_SUGERIDO), "", rsaux!VCHA_PRI_RUTA_PEDIDO_SUGERIDO)
            Me.txt_ruta_archivos_articulos = IIf(IsNull(rsaux!VCHA_PRI_RUTA_ARCHIVOS_ARTICULOS), "", rsaux!VCHA_PRI_RUTA_ARCHIVOS_ARTICULOS)
            Me.txt_ruta_envio_facturas = IIf(IsNull(rsaux!VCHA_PRI_RUTA_ENVIOS_FACTURAS), "", rsaux!VCHA_PRI_RUTA_ENVIOS_FACTURAS)
            Me.txt_tipo_calculo_factura_catalogos = IIf(IsNull(rsaux!CHAR_PRI_TIPO_CALCULO_FACTURA_CATALOGOS), "", rsaux!CHAR_PRI_TIPO_CALCULO_FACTURA_CATALOGOS)
            Me.txt_dias_tolerancia = IIf(IsNull(rsaux!INTE_PRI_DIAS_TOLERANCIA_FACTURACION_CATALOGOS), 0, rsaux!INTE_PRI_DIAS_TOLERANCIA_FACTURACION_CATALOGOS)
            Me.txt_ruta_archivos_enviar = IIf(IsNull(rsaux!VCHA_PRI_RUTA_ARCHIVOS_ENVIAR), "", rsaux!VCHA_PRI_RUTA_ARCHIVOS_ENVIAR)
            Me.txt_ruta_archivos_reempaque = IIf(IsNull(rsaux!VCHA_PRI_RUTA_REEMPAQUE), "", rsaux!VCHA_PRI_RUTA_REEMPAQUE)
         Else
            Me.txt_renglones_factura = ""
            Me.txt_renglones_nota_credito = ""
            Me.txt_ruta_nota_envios_plantas = ""
            Me.txt_ruta_devoluciones_tiendas = ""
            Me.txt_ruta_pedidos = ""
            Me.txt_ruta_almacenes = ""
            Me.txt_ruta_cobranza = ""
            Me.txt_tolerancia_saldos = ""
            Me.txt_tipo_agrupamiento = ""
            Me.txt_ruta_archivos_pedidos_sugerido = ""
            Me.txt_ruta_archivos_articulos = ""
            Me.txt_ruta_envio_facturas = ""
            Me.txt_tipo_calculo_factura_catalogos = ""
            Me.txt_dias_tolerancia = ""
            Me.txt_ruta_archivos_enviar = ""
            Me.txt_ruta_archivos_reempaque = ""
         End If
         rsaux.Close
      Else
         Me.txt_renglones_factura = ""
         Me.txt_renglones_nota_credito = ""
         Me.txt_ruta_nota_envios_plantas = ""
         Me.txt_ruta_devoluciones_tiendas = ""
         Me.txt_ruta_pedidos = ""
         Me.txt_ruta_almacenes = ""
         Me.txt_ruta_cobranza = ""
         Me.txt_tolerancia_saldos = ""
         Me.txt_tipo_agrupamiento = ""
         Me.txt_ruta_archivos_pedidos_sugerido = ""
         Me.txt_ruta_archivos_articulos = ""
         Me.txt_ruta_envio_facturas = ""
         Me.txt_tipo_calculo_factura_catalogos = ""
         Me.txt_dias_tolerancia = ""
         Me.txt_ruta_archivos_enviar = ""
         Me.txt_ruta_archivos_reempaque = ""
         MsgBox "Clave de empresa incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_nombre_empresa = ""
      Me.txt_renglones_factura = ""
      Me.txt_renglones_nota_credito = ""
      Me.txt_ruta_nota_envios_plantas = ""
      Me.txt_ruta_devoluciones_tiendas = ""
      Me.txt_ruta_pedidos = ""
      Me.txt_ruta_almacenes = ""
      Me.txt_ruta_cobranza = ""
      Me.txt_tolerancia_saldos = ""
      Me.txt_tipo_agrupamiento = ""
      Me.txt_ruta_archivos_pedidos_sugerido = ""
      Me.txt_ruta_archivos_articulos = ""
      Me.txt_ruta_envio_facturas = ""
      Me.txt_tipo_calculo_factura_catalogos = ""
      Me.txt_dias_tolerancia = ""
      Me.txt_ruta_archivos_enviar = ""
      Me.txt_ruta_archivos_reempaque = ""
   End If
End Sub

Private Sub txt_nombre_empresa_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_renglones_factura_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_renglones_nota_credito_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_almacenes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_archivos_articulos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_archivos_enviar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_archivos_pedidos_sugerido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_archivos_reempaque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_cobranza_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_devoluciones_tiendas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_envio_facturas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_nota_envios_plantas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_pedidos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_tipo_agrupamiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_tipo_calculo_factura_catalogos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_tolerancia_saldos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub
