VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmbonificaciones_externas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aplicación de Notas de Crédito a Cartera"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista2 
      Height          =   2400
      Left            =   1275
      TabIndex        =   0
      Top             =   240
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista2 
         Height          =   1875
         Left            =   30
         TabIndex        =   8
         Top             =   435
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3307
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
      Begin VB.Label lbl_lista2 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   9
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   1695
      Left            =   1245
      TabIndex        =   10
      Top             =   435
      Width           =   4050
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1230
         Left            =   30
         TabIndex        =   11
         Top             =   405
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   2170
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
            Object.Width           =   1464
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5380
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   12
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos Generales "
      Height          =   1725
      Left            =   105
      TabIndex        =   23
      Top             =   420
      Width           =   11355
      Begin VB.TextBox txt_numero 
         Height          =   330
         Left            =   3465
         TabIndex        =   7
         Top             =   1245
         Width           =   1965
      End
      Begin VB.TextBox txt_serie 
         Height          =   345
         Left            =   1755
         TabIndex        =   6
         Top             =   1245
         Width           =   810
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   315
         Left            =   1755
         TabIndex        =   3
         Top             =   585
         Width           =   1725
      End
      Begin VB.TextBox txt_saldo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1755
         TabIndex        =   5
         Top             =   915
         Width           =   1725
      End
      Begin VB.TextBox txt_clase 
         Height          =   315
         Left            =   1755
         TabIndex        =   1
         Top             =   255
         Width           =   1725
      End
      Begin VB.TextBox txt_nombre_clase 
         Height          =   315
         Left            =   3495
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   255
         Width           =   5190
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   3495
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   585
         Width           =   5190
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   2700
         TabIndex        =   29
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   165
         TabIndex        =   27
         Top             =   645
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe a Aplicar:"
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   975
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   165
         TabIndex        =   25
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento:"
         Height          =   195
         Left            =   150
         TabIndex        =   24
         Top             =   315
         Width           =   855
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11100
      Picture         =   "frmbonificaciones_externas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmbonificaciones_externas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmbonificaciones_externas.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Cargos "
      Height          =   4875
      Left            =   105
      TabIndex        =   13
      Top             =   2205
      Width           =   11355
      Begin VB.Frame frm_cantidad_aplicar 
         Height          =   885
         Left            =   4065
         TabIndex        =   15
         Top             =   1350
         Width           =   2955
         Begin VB.TextBox txt_cantidad_aplicar 
            Height          =   360
            Left            =   1005
            TabIndex        =   16
            Top             =   390
            Width           =   1890
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   165
            TabIndex        =   18
            Top             =   473
            Width           =   570
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000D&
            Caption         =   " Importe a Aplicar"
            ForeColor       =   &H8000000E&
            Height          =   270
            Left            =   0
            TabIndex        =   17
            Top             =   15
            Width           =   2940
         End
      End
      Begin VB.TextBox txt_total_aplicado 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9735
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   4455
         Width           =   1545
      End
      Begin MSComctlLib.ListView lv_facturas 
         Height          =   4185
         Left            =   90
         TabIndex        =   19
         Top             =   225
         Width           =   11190
         _ExtentX        =   19738
         _ExtentY        =   7382
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
         MousePointer    =   1
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Número"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Plazo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Moneda"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Abonos"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Saldo"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Aplicar    "
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Clave Moneda"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Serie"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Iva"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   90
      TabIndex        =   28
      Top             =   285
      Width           =   11430
   End
End
Attribute VB_Name = "frmbonificaciones_externas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_serie As String
Dim var_clave_moneda As String
Dim var_tipo_Cambio As Double
Dim var_agente As String
Dim var_grupo_actual As String
Dim var_grupo_real As String
Dim var_titular As String
Dim var_plazo As Integer
Dim var_iva As Double
Dim var_numero_renglones As Double






Private Sub cmb_series_Click()
   var_serie = cmb_series
End Sub

Private Sub cmd_aplicar_Click()
End Sub

Private Sub cmd_imprimir_Click()
Dim var_importe_neto_1 As Double
Dim var_importe_total_1 As Double
Dim var_subimporte_1 As Double
Dim var_importe_iva_1 As Double

Dim var_tipo_Cambio As Double
Dim var_importe_factura As Double
Dim var_importe_pago As Double
Dim var_importe_saldo_pago As Double
Dim var_importe_total As Double
Dim var_fecha_pago As Date
Dim var_fecha_factura As Date
Dim var_contador_pagos As Double
Dim var_contador_facturas As Double
Dim var_descuento_agente As Double
Dim var_descuento_sistema As Double
Dim var_saldo As Double
Dim si As Integer
Dim i, n As Integer
Dim var_importe As Double
Dim var_descuento As Double
Dim var_importe_descuento As Double
Dim var_moneda_local As Integer
Dim var_posible_tipo_cambio As Boolean
Dim var_numero_folio As Double
Dim var_serie_cargo As String
Dim var_importe_neto As Double
Dim var_subimporte As Double
Dim var_importe_iva As Double
Dim var_numero_nota_inicio As Double
Dim var_k As Integer
Dim var_l As Integer
Dim var_numero_nota As Double
Dim var_contador_notas As Double
Dim var_iva_pasado As Double
var_posible_iva = 1
var_iva_pasado = 0
For var_j = 1 To lv_facturas.ListItems.Count
    lv_facturas.ListItems.Item(var_j).Selected = True
    If (lv_facturas.selectedItem.SubItems(8) * 1) > 0 Then
       If var_iva_pasado = 0 Then
          var_iva_pasado = CDbl(Me.lv_facturas.selectedItem.SubItems(11))
       Else
          If var_iva_pasado <> CDbl(Me.lv_facturas.selectedItem.SubItems(11)) Then
             var_posible_iva = 0
          End If
       End If
    End If
Next var_j
If var_posible_iva = 1 Then
   var_iva = var_iva_pasado
If lv_facturas.ListItems.Count > 0 Then
   If Trim(txt_clase) <> "" Then
      If txt_serie <> "" Then
         If IsNumeric(txt_numero) Then
            If rsaux4.State = 1 Then
               rsaux4.Close
            End If
            var_importe = 0
            For var_j = 1 To Me.lv_facturas.ListItems.Count
                lv_facturas.ListItems.Item(var_j).Selected = True
                If (lv_facturas.selectedItem.SubItems(8) * 1) > 0 Then
                   var_importe = var_importe + CDbl(lv_facturas.selectedItem.SubItems(8))
                End If
            Next var_j
            If CDbl(var_importe) = CDbl(Me.txt_saldo) Then
               rsaux4.Open "select * from tb_encabezado_Cartera where vcha_car_documento = '" + txt_clase + "' and vcha_ser_serie_id = '" + txt_serie + "' and inte_car_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
               If rsaux4.EOF Then
                  rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_moneda_local = IIf(IsNull(rs!inte_mon_moneda_local), 0, rs!inte_mon_moneda_local)
                  End If
                  rs.Close
                  var_tipo_Cambio = 1
                  If var_moneda_local = 0 Then
                     rs.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_tipo_Cambio = IIf(IsNull(rs!mone_tca_importe), 1, rs!mone_tca_importe)
                        var_posible_tipo_cambio = True
                     Else
                        var_posible_tipo_cambio = False
                     End If
                     rs.Close
                  Else
                     var_posible_tipo_cambio = True
                  End If
                  'MsgBox CStr(var_tipo_Cambio)
                  If var_posible_tipo_cambio = True Then
                     var_contador_renglones = 0
                     var_contador_notas = 0
                     n = lv_facturas.ListItems.Count
                     For i = 1 To n
                         lv_facturas.ListItems.Item(i).Selected = True
                         If (lv_facturas.selectedItem.SubItems(8) * 1) > 0 Then
                            var_contador_renglones = var_contador_renglones + 1
                         End If
                         If var_contador_renglones = var_numero_renglones Then
                            var_contador_notas = var_contador_notas + 1
                            var_contador_renglones = 0
                         End If
                     Next i
                     If (var_contador_renglones > 0) And (var_contador_renglones < var_numero_renglones) Then
                        var_contador_notas = var_contador_notas + 1
                     End If
                     var_serie = txt_serie
                     var_numero_folio = txt_numero
            
                     var_numero_nota = var_numero_folio
                     var_numero_nota_anterior = var_numero_nota
                     var_numero_nota_inicio = var_numero_folio
                     If var_contador_notas > 0 Then
                        If var_contador_notas = 1 Then
                           si = MsgBox("Se va a imprimir la nota de crédito número " + Str(var_numero_folio) + ", ¿la impresora esta lista?", vbYesNo, "ATENCION")
                        End If
                        If var_contador_notas > 1 Then
                           si = MsgBox("Se van a imprimir de la nota " + Str(var_numero_folio) + " a la " + Str(var_numero_folio + (var_contador_notas - 1)) + ", ¿la impresora esta lista?", vbYesNo, "ATENCION")
                        End If
                        var_numero_nota_inicio = var_numero_folio
                        If si = 6 Then
                           si = MsgBox("Confirmar la impresión de la Nota de Crédito", vbYesNo, "ATENCION")
                           If si = 6 Then
                              Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
                              Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
                              n = lv_facturas.ListItems.Count
                      
                              cnn.BeginTrans
                              var_contador_notas = 0
                              var_j = 0
                      
                              For i = 1 To n
                                  lv_facturas.ListItems.Item(i).Selected = True
                                  If (lv_facturas.selectedItem.SubItems(8) * 1) > 0 Then
                                     var_j = var_j + 1
                                  End If
                              Next i
                              For i = 1 To n
                                  lv_facturas.ListItems.Item(i).Selected = True
                                  If i = 1 Then
                                     var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "NC", txt_clase, txt_clase, var_numero_folio, "-", "", "", 0, CStr(Date), var_agente, var_grupo_actual, var_grupo_real, var_titular, txt_clave_cliente, "", var_plazo, var_iva, 0, 0, 0, 0, 0, var_importe_total, var_importe_iva, 0, 0, 0, 0, 0, var_subimporte, var_importe_neto, "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, var_clave_moneda, var_tipo_Cambio, var_serie, "")
                                     rsaux9.Open "update tb_encabezado_cartera set inte_car_nota_credito_tienda = 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_tipo_documento = 'NC' and vcha_car_documento = '" + Me.txt_clase + "' and vcha_Ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                                     rsaux3.Open "insert into tb_secuencia_notas_credito (vcha_emp_empresa_id, vcha_Ser_serie_id, inte_snc_numero_anterior, inte_snc_numero_actual) values ('" + var_empresa + "', '" + var_serie + "', " + CStr(var_numero_folio) + ", " + CStr(var_numero_folio) + ")", cnn, adOpenDynamic, adLockOptimistic
                                     var_importe_neto = 0
                                     var_importe_total = 0
                                     var_subimporte = 0
                                     var_importe_iva = 0
                                     var_contador = 0
                                  End If
                                  
                                  If (lv_facturas.selectedItem.SubItems(8) * 1) > 0 Then
                                     var_importe_neto_1 = ((lv_facturas.selectedItem.SubItems(8) * 1) * var_tipo_Cambio)
                                     var_importe_total_1 = ((var_importe_neto_1 / (1 + (var_iva / 100))))
                                     var_subimporte_1 = var_importe_total_1
                                     var_importe_iva_1 = (var_importe_neto_1 - var_importe_total_1)
                            
                                     var_importe_neto = var_importe_neto + var_importe_neto_1
                                     var_importe_total = var_importe_total + var_importe_total_1
                                     var_subimporte = var_subimporte + var_subimporte_1
                                     var_importe_iva = var_importe_iva + var_importe_iva_1
                            
                                     var_contador = var_contador + 1
                           
                                     var_serie_cargo = lv_facturas.selectedItem.SubItems(10)
                                     var_importe = lv_facturas.selectedItem.SubItems(8) * 1
                                     var_descuento = 0
                                     var_importe_descuento = 0
                                     rs.Open "insert into tb_estado_cuenta (vcha_emp_empresa_id, vcha_ecu_serie_cargo, vcha_ecu_movimiento_cargo, inte_ecu_numero_cargo, vcha_ecu_serie_abono, vcha_ecu_movimiento_abono, inte_ecu_numero_abono, floa_ecu_importe_Cargo, floa_ecu_importe_abono) values ('" + var_empresa + "', '" + var_serie_cargo + "' ,'" + Trim(lv_facturas.selectedItem) + "', " + lv_facturas.selectedItem.SubItems(1) + ",'" + var_serie + "' ,'" + txt_clase + "'," + Str(var_numero_folio) + ", 0, " + Str(var_importe * var_tipo_Cambio) + ")", cnn, adOpenDynamic, adLockOptimistic
                                     rs.Open "insert into tb_detalle_bonificaciones (vcha_emp_empresa_id, vcha_car_documento,vcha_car_clase_id, vcha_ser_serie_id, inte_car_numero, inte_car_factura, floa_dbo_importe, floa_dbo_iva, char_dbo_estatus) values ('" + var_empresa + "', '" + txt_clase + "', '','" + var_serie + "', " + CStr(var_numero_folio) + "," + lv_facturas.selectedItem.SubItems(1) + ", " + Str(var_importe * var_tipo_Cambio) + "," + Str(var_iva) + ",'')"
                                  End If
                                  If (var_contador = var_numero_renglones) Or (i = n) Then
                                     var_contador = 0
                                     var_imprimir = True
                                  End If
                                  var_imprimir = False
                              Next i
                              rs.Open "update tb_series set inte_ser_nota_credito = inte_ser_nota_credito + " + CStr(var_contador_notas) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                              cnn.CommitTrans
'''''''''''''''''
'''''''''''''''     IMPRESION DE LA NOTA DE CARGO
                              If Trim(txt_clave_cliente) <> "" Then
                                 rs.Open "select * from tb_clientes where vcha_cli_clave_id ='" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rs.EOF Then
                                    txt_nombre_cliente = rs!VCHA_CLI_NOMBRE
                                    rs.Close
                                    rs.Open "select * from vw_saldos_facturas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + txt_clave_cliente + "' and floa_sal_importe > 0", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rs.EOF Then
                                       lv_facturas.ListItems.Clear
                                       var_contador_facturas = 0
                                       While Not rs.EOF
                                             var_saldo = (IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE) - IIf(IsNull(rs!importe_saldo), 0, rs!importe_saldo))
                                             If var_saldo > 0 Then
                                                Set list_item = lv_facturas.ListItems.Add(, , rs!vcha_Car_documento)
                                                var_tipo_Cambio = IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                                                var_importe_factura = IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto) / var_tipo_Cambio
                                                list_item.SubItems(1) = IIf(IsNull(rs!INTE_CAR_NUMERO), "", rs!INTE_CAR_NUMERO)
                                                list_item.SubItems(2) = IIf(IsNull(rs!dtim_Car_fecha), "", Format(rs!dtim_Car_fecha, "Short Date"))
                                                var_fecha_factura = Format(rs!dtim_Car_fecha, "Short Date")
                                                var_dias = var_fecha_pago - var_fecha_factura
                                                list_item.SubItems(3) = IIf(IsNull(rs!INTE_CAR_PLAZO), 0, rs!INTE_CAR_PLAZO)
                                                list_item.SubItems(4) = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
                                                list_item.SubItems(5) = Format(var_importe_factura, "###,##0.00")
                                                list_item.SubItems(6) = Format(var_importe_factura - IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE), "###,##0.00")
                                                list_item.SubItems(7) = Format(IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE), "###,##0.00")
                                                list_item.SubItems(8) = Format(0, "###,##0.00")
                                                list_item.SubItems(9) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                                                list_item.SubItems(10) = IIf(IsNull(rs!VCHA_SER_SERIE_ID), "", rs!VCHA_SER_SERIE_ID)
                                             End If
                                             rs.MoveNext:
                                       Wend
                                       rs.Close
                                       txt_total_aplicado = Format(0, "###,##0.00")
                                    Else
                                       rs.Close
                                       lv_facturas.ListItems.Clear
                                       txt_total_aplicado = Format(0, "###,##0.00")
                                    End If
                                 Else
                                    rs.Close
                                    txt_clave_cliente = ""
                                    txt_nombre_cliente = ""
                                    txt_saldo = ""
                                    txt_total_aplicado = Format(0, "###,##0.00")
                                    lv_facturas.ListItems.Clear
                                    MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
                                End If
                             End If
                             MsgBox "Se han terminado de aplicar los pagos", vbOKOnly, "ATENCION"
                          End If
                       End If
                    End If
                 Else
                    MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
                 End If
              Else
                 MsgBox "Ya existe el documento", vbOKOnly, "ATENCION"
              End If
           Else
              MsgBox "El importe asignado a las facturas es diferente al importe de la nota de credito", vbOKOnly, "ATENCION"
           End If
        Else
           MsgBox "Número de folio incorrecto", vbOKOnly, "ATENCION"
        End If
     Else
         MsgBox "Se debe de indicar una clase", vbOKOnly, "ATENCION"
     End If
   Else
      MsgBox "No se a indicado una clase del movimiento", vbOKOnly, "ATENCION"
   End If
Else
   MsgBox "El cliente seleccionado no tiene facturas vivas", vbOKOnly, "ATENCION"
End If
   MsgBox "No puede mezclar facturas con distintos tipos de IVA", vbOKOnly, "ATENCION"
Else
End If
   txt_clave_cliente.Enabled = False
   cmd_imprimir.Enabled = False
   txt_saldo.Enabled = False
End Sub

Private Sub cmd_nuevo_Click()
   txt_clave_cliente.Enabled = True
   cmd_imprimir.Enabled = True
   txt_clave_cliente = ""
   txt_nombre_cliente = ""
   txt_saldo = ""
   lv_facturas.ListItems.Clear
   txt_total_aplicado = ""
   txt_saldo.Enabled = True
   If Me.txt_clase.Enabled = True Then
      txt_clase.SetFocus
   Else
      txt_clave_cliente.SetFocus
   End If
   
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   frm_lista2.Visible = False
   txt_clave_cliente.Enabled = True
   frm_lista.Visible = False
   txt_saldo.Enabled = True
   Top = 0
   Left = 0
   frm_cantidad_aplicar.Visible = False
   rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'"
   var_numero_renglones = rs!INTE_PRI_RENGLONES_NOTA_CREDITO
   rs.Close
   rs.Open "select * from tb_clases_Cartera where vcha_car_documento = 'BO' order by vcha_car_nombre ", cnn, adOpenDynamic, adLockBatchOptimistic
   If Not rs.EOF Then
      var_contador_movimiento = 0
      While Not rs.EOF
         var_contador_movimiento = var_contador_movimiento + 1
         rs.MoveNext
      Wend
      
      If var_contador_movimiento > 1 Then
         txt_nombre_clase.Enabled = True
         txt_clase.Enabled = True
      Else
         txt_nombre_clase.Enabled = False
         txt_clase.Enabled = False
      End If
      rs.MoveFirst
      txt_nombre_clase = rs!vcha_Car_nombre
      txt_clase = rs!vcha_Car_clase_id
   Else
      MsgBox "No se a indicado una clase de Bonificación", vbOKOnly, "ATENCION"
      txt_clase.Enabled = False
      cmb_clases.Enabled = False
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_reporte_valuacion_devoluciones)
End Sub

Private Sub lv_facturas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_facturas, ColumnHeader)
End Sub

Private Sub lv_facturas_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F4 para indicar el importe a aplicar a la factura"
End Sub

Private Sub lv_facturas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 115 Then
      If Trim(txt_saldo) = "" Then
         txt_saldo = 0
      End If
      If txt_saldo > 0 Then
         frm_cantidad_aplicar.Visible = True
         txt_cantidad_aplicar = ""
         txt_cantidad_aplicar.SetFocus
      Else
         MsgBox "No se a indicado el importe de la bonificación", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub lv_facturas_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_clase = lv_lista.selectedItem
         txt_nombre_clase = lv_lista.selectedItem.SubItems(1)
      Else
         txt_clase = ""
         txt_nombre_clase = ""
      End If
      frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_lista2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista2, ColumnHeader)
End Sub

Private Sub lv_lista2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista2.ListItems.Count > 0 Then
         txt_clave_cliente = lv_lista2.selectedItem
         txt_nombre_cliente = lv_lista2.selectedItem.SubItems(1)
      Else
         txt_clave_cliente = ""
         txt_nombre_cliente = ""
      End If
      txt_clave_cliente.SetFocus
      frm_lista2.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista2.Visible = False
   End If
End Sub

Private Sub lv_lista2_LostFocus()
   frm_lista2.Visible = False
End Sub

Private Sub txt_cantidad_aplicar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If IsNumeric(txt_cantidad_aplicar) Then
         If (txt_cantidad_aplicar * 1) + (lv_facturas.selectedItem.SubItems(8) * 1) > (lv_facturas.selectedItem.SubItems(7) * 1) Then
            MsgBox "La cantidad a aplicar exede el importe del saldo de la factura", vbOKOnly, "ATENCIO"
         Else
            If (txt_total_aplicado * 1) + (txt_cantidad_aplicar * 1) <= (txt_saldo * 1) Then
               lv_facturas.selectedItem.SubItems(8) = Format(txt_cantidad_aplicar + (lv_facturas.selectedItem.SubItems(8) * 1), "###,##0.00")
               txt_total_aplicado = (txt_total_aplicado * 1) + (txt_cantidad_aplicar * 1)
            Else
               MsgBox "La cantidad a aplicar exede al importe del saldo del pago del cliente", vbOKOnly, "ATENCION"
            End If
         End If
      Else
         MsgBox "Importe Incorrecto", vbOKOnly, "ATENCION"
      End If
      frm_cantidad_aplicar.Visible = False
      lv_facturas.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_cantidad_aplicar.Visible = False
      If lv_facturas.ListItems.Count > 0 Then
         lv_facturas.SetFocus
      End If
   End If
End Sub

Private Sub txt_clase_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clase_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      rs.Open "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where vcha_car_documento= 'BO' order by vcha_Car_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            var_contador_lista = var_contador_lista + 1
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Car_clase_id)
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_Car_nombre), "", rs!vcha_Car_nombre))
            rs.MoveNext
         Wend
      End If
      rs.Close
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clase_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clase_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(Me.txt_clase) <> "" Then
      rs.Open "select * FROM tb_clases_cartera where vcha_car_documento= 'BO' and vcha_Car_clase_id = '" + Me.txt_clase + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_clase = IIf(IsNull(rs!vcha_Car_nombre), "", rs!vcha_Car_nombre)
      Else
         MsgBox "Clase de cartera invalida", vbOKOnly, "ATENCION"
         Me.txt_clase = ""
         Me.txt_nombre_clase = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_clase = ""
      Me.txt_clase = ""
   End If
End Sub

Private Sub txt_clave_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clave_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista2.ListItems.Clear
      rs.Open "select * from vw_clientes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_cli_nombre ", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista2.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista2 = "CLIENTES"
      VAR_TIPO_LISTA = 1
      frm_lista2.Visible = True
      lv_lista2.SetFocus
   End If
End Sub

Private Sub txt_clave_cliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_cliente = rs!VCHA_CLI_NOMBRE
         rs.Close
      Else
         rs.Close
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
         txt_clave_cliente = ""
         txt_nombre_cliente = ""
      End If
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_cliente_LostFocus()
   Dim var_importe_factura As Double
   Dim var_importe_pago As Double
   Dim var_importe_saldo_pago As Double
   Dim var_importe_total As Double
   Dim var_fecha_pago As Date
   Dim var_fecha_factura As Date
   Dim var_contador_pagos As Double
   Dim var_contador_facturas As Double
   Dim var_descuento_agente As Double
   Dim var_descuento_sistema As Double
   Dim var_saldo As Double
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_clave_cliente) <> "" Then
      rs.Open "select * from vw_clientes where vcha_cli_clave_id ='" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_cliente = rs!VCHA_CLI_NOMBRE
         var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
         var_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
         var_grupo_actual = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
         var_grupo_real = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
         var_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
         var_plazo = IIf(IsNull(rs!inte_pla_dias), 0, rs!inte_pla_dias)
         var_iva = IIf(IsNull(rs!FLOA_TPE_IVA), 0, rs!FLOA_TPE_IVA)
         rs.Close
         rs.Open "select * from vw_saldos_facturas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + txt_clave_cliente + "' and floa_sal_importe > 0", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            lv_facturas.ListItems.Clear
            var_contador_facturas = 0
            While Not rs.EOF
               var_saldo = (IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE) - IIf(IsNull(rs!importe_saldo), 0, rs!importe_saldo))
               If var_saldo > 0 Then
                  Set list_item = lv_facturas.ListItems.Add(, , rs!vcha_Car_documento)
                  var_tipo_Cambio = IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                  var_importe_factura = IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto) / var_tipo_Cambio
                  list_item.SubItems(1) = IIf(IsNull(rs!INTE_CAR_NUMERO), "", rs!INTE_CAR_NUMERO)
                  list_item.SubItems(2) = IIf(IsNull(rs!dtim_Car_fecha), "", Format(rs!dtim_Car_fecha, "Short Date"))
                  var_fecha_factura = Format(rs!dtim_Car_fecha, "Short Date")
                  var_dias = var_fecha_pago - var_fecha_factura
                  list_item.SubItems(3) = IIf(IsNull(rs!INTE_CAR_PLAZO), 0, rs!INTE_CAR_PLAZO)
                  list_item.SubItems(4) = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
                  list_item.SubItems(5) = Format(var_importe_factura, "###,##0.00")
                  list_item.SubItems(6) = Format(var_importe_factura - IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE), "###,##0.00")
                  list_item.SubItems(7) = Format(IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE), "###,##0.00")
                  list_item.SubItems(8) = Format(0, "###,##0.00")
                  list_item.SubItems(9) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                  list_item.SubItems(10) = IIf(IsNull(rs!VCHA_SER_SERIE_ID), "", rs!VCHA_SER_SERIE_ID)
                  list_item.SubItems(11) = IIf(IsNull(rs!floa_Car_porcentaje_iva), "", rs!floa_Car_porcentaje_iva)
               End If
               rs.MoveNext:
            Wend
            rs.Close
            txt_total_aplicado = Format(0, "###,##0.00")
         Else
            rs.Close
            lv_facturas.ListItems.Clear
            txt_total_aplicado = Format(0, "###,##0.00")
         End If
      Else
         rs.Close
         txt_clave_cliente = ""
         txt_nombre_cliente = ""
         txt_saldo = ""
         txt_total_aplicado = Format(0, "###,##0.00")
         lv_facturas.ListItems.Clear
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   If Me.lv_facturas.ListItems.Count > 20 Then
      lv_facturas.ColumnHeaders(2).Width = 1250
   Else
      lv_facturas.ColumnHeaders(2).Width = 1500.09
   End If
End Sub

Private Sub txt_descuento_Change()

End Sub

Private Sub txt_descuento_KeyPress(KeyAscii As Integer)
End Sub

Private Sub txt_nombre_clase_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      rs.Open "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where vcha_car_documento= 'DV' order by vcha_Car_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            var_contador_lista = var_contador_lista + 1
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Car_clase_id)
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_Car_nombre), "", rs!vcha_Car_nombre))
            rs.MoveNext
         Wend
      End If
      rs.Close
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_clase_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista2.ListItems.Clear
      rs.Open "select * from vw_clientes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_cli_nombre ", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista2.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista2 = "CLIENTES"
      VAR_TIPO_LISTA = 1
      frm_lista2.Visible = True
      lv_lista2.SetFocus
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Me.lv_facturas.ListItems.Count > 0 Then
         lv_facturas.SetFocus
      Else
         Call pro_enfoque(KeyAscii)
      End If
   End If
End Sub

Private Sub txt_saldo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub
