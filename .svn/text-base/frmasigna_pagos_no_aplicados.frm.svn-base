VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmasigna_pagos_no_aplicados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de pagos no aplicados"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   3795
      TabIndex        =   24
      Top             =   -75
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1875
         Left            =   30
         TabIndex        =   25
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
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   26
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.TextBox txt_serie 
      Height          =   285
      Left            =   12540
      TabIndex        =   23
      Top             =   2370
      Width           =   150
   End
   Begin VB.CommandButton cmd_notas_credito 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmasigna_pagos_no_aplicados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Nota de Crédito"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Pagos sin saldo"
      Height          =   2130
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   11370
      Begin MSComctlLib.ListView lv_pagos 
         Height          =   1815
         Left            =   75
         TabIndex        =   4
         Top             =   210
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   3201
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
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Folio Cobranza"
            Object.Width           =   2743
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Num. Pago"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Moneda"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Importe      "
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Aplicado     "
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Saldo        "
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Aplicar     "
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Tipo Cambio "
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Clave Moneda"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Serie"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Banco"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Cheque"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Partida"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos Generales "
      Height          =   1050
      Left            =   120
      TabIndex        =   11
      Top             =   510
      Width           =   11385
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   3510
         TabIndex        =   2
         Top             =   240
         Width           =   5160
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   315
         Left            =   1770
         TabIndex        =   0
         Top             =   240
         Width           =   1725
      End
      Begin VB.TextBox txt_saldo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   165
         TabIndex        =   13
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Saldo a Aplicar:"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   660
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11115
      Picture         =   "frmasigna_pagos_no_aplicados.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_aplicar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmasigna_pagos_no_aplicados.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Aplicar Pagos Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmasigna_pagos_no_aplicados.frx":0886
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Cargos "
      Height          =   3270
      Left            =   150
      TabIndex        =   1
      Top             =   3945
      Width           =   11355
      Begin VB.Frame frm_cantidad_aplicar 
         Height          =   1200
         Left            =   4080
         TabIndex        =   15
         Top             =   765
         Width           =   2955
         Begin VB.TextBox txt_descuento 
            Height          =   360
            Left            =   1005
            TabIndex        =   20
            Top             =   772
            Width           =   795
         End
         Begin VB.TextBox txt_cantidad_aplicar 
            Height          =   360
            Left            =   1005
            TabIndex        =   17
            Top             =   390
            Width           =   1890
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   1875
            TabIndex        =   22
            Top             =   855
            Width           =   120
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Descuento:"
            Height          =   195
            Left            =   165
            TabIndex        =   21
            Top             =   855
            Width           =   825
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   165
            TabIndex        =   19
            Top             =   473
            Width           =   570
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000D&
            Caption         =   " Importe a Aplicar"
            ForeColor       =   &H8000000E&
            Height          =   270
            Left            =   0
            TabIndex        =   16
            Top             =   15
            Width           =   2940
         End
      End
      Begin VB.TextBox txt_total_aplicado 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   10050
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2865
         Width           =   1200
      End
      Begin MSComctlLib.ListView lv_facturas 
         Height          =   2655
         Left            =   90
         TabIndex        =   5
         Top             =   195
         Width           =   11190
         _ExtentX        =   19738
         _ExtentY        =   4683
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
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   917
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Número"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Plazo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Moneda"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "% Desc. Sist."
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Abonos"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Saldo"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "% Desc."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Aplicar    "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Dias"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Clave Moneda"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Serie"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   0
      TabIndex        =   10
      Top             =   330
      Width           =   11520
   End
End
Attribute VB_Name = "frmasigna_pagos_no_aplicados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim list_item As ListItem
Dim var_pago_seleccionado As Integer
Dim var_tolerancia As Integer
Dim var_dias As Integer
Dim var_serie As String
Dim var_serie_cargo As String
Dim var_agente As String




Private Sub cmb_clientes_LostFocus()
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
   If Trim(txt_clave_cliente) <> "" Then
      rs.Open "select * from tb_clientes where vcha_cli_clave_id ='" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         cmb_clientes = rs!vcha_cli_nombre
         rs.Close
         rs.Open "select a.inte_rut_tolerancia from tb_rutas a, tb_clientes b where b.vcha_cli_clave_id = '" + txt_clave_cliente + "' and a.vcha_rut_ruta_id = b.vcha_rut_ruta_id", cnn, adOpenDynamic, adLockOptimistic
         var_tolerancia = IIf(IsNull(rs!inte_rut_tolerancia), 0, rs!inte_rut_tolerancia)
         rs.Close
         rs.Open "select * from vw_saldos_pagos_no_aplicados where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + txt_clave_cliente + "' and floa_sal_importe > 0 and char_sal_afectacion = '-'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_importe_total = 0
            lv_pagos.ListItems.Clear
            var_pago_seleccionado = 0
            var_contador_pagos = 0
            While Not rs.EOF
               Set list_item = lv_pagos.ListItems.Add(, , rs!VCHA_RCO_FOLIO)
               list_item.SubItems(1) = Format(IIf(IsNull(rs!DTIM_CAR_FECHA), "", rs!DTIM_CAR_FECHA), "Short Date")
               If var_contador_pagos = 0 Then
                  var_fecha_pago = Format(rs!DTIM_CAR_FECHA, "Short Date")
                  var_descuento_agente = 0
               End If
               list_item.SubItems(2) = IIf(IsNull(rs!inte_car_numero), 0, rs!inte_car_numero)
               list_item.SubItems(3) = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
               var_importe_pago = IIf(IsNull(rs!floa_Rco_importe), 0, rs!floa_Rco_importe)
               var_importe_saldo_pago = IIf(IsNull(rs!floa_sal_importe), 0, rs!floa_sal_importe)
               var_importe_total = var_importe_total + var_importe_saldo_pago
               list_item.SubItems(4) = Format(var_importe_pago, "###,##0.00")
               list_item.SubItems(5) = Format((var_importe_pago - var_importe_saldo_pago), "###,##0.00")
               list_item.SubItems(6) = Format((var_importe_saldo_pago), "###,##0.00")
               list_item.SubItems(7) = Format(0, "###,##0.00")
               list_item.SubItems(9) = IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
               list_item.SubItems(10) = IIf(IsNull(rs!VCHA_MON_MONEDA_ID), "", rs!VCHA_MON_MONEDA_ID)
               list_item.SubItems(11) = IIf(IsNull(rs!vcha_ser_serie_id), "", rs!vcha_ser_serie_id)
               list_item.SubItems(12) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
               list_item.SubItems(13) = IIf(IsNull(rs!VCHA_RCO_CHEQUE), "", rs!VCHA_RCO_CHEQUE)
               var_contador_pagos = var_contador_pagos + 1
               rs.MoveNext
            Wend
            rs.Close
            txt_saldo = Format(var_importe_total, "###,##0.00")
            rs.Open "select * from vw_saldos_facturas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + txt_clave_cliente + "' and floa_sal_importe > 0", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               lv_facturas.ListItems.Clear
               var_contador_facturas = 0
               While Not rs.EOF
                  var_saldo = (IIf(IsNull(rs!floa_sal_importe), 0, rs!floa_sal_importe) - IIf(IsNull(rs!importe_saldo), 0, rs!importe_saldo))
                  If var_saldo > 0 Then
                     Set list_item = lv_facturas.ListItems.Add(, , rs!vcha_car_documento)
                     var_tipo_Cambio = IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                     var_importe_factura = IIf(IsNull(rs!FLOA_CAR_IMPORTE_NETO), 0, rs!FLOA_CAR_IMPORTE_NETO) / var_tipo_Cambio
                     list_item.SubItems(1) = IIf(IsNull(rs!inte_car_numero), "", rs!inte_car_numero)
                     list_item.SubItems(2) = IIf(IsNull(rs!DTIM_CAR_FECHA), "", Format(rs!DTIM_CAR_FECHA, "Short Date"))
                     var_fecha_factura = Format(rs!DTIM_CAR_FECHA, "Short Date")
                     var_dias = var_fecha_pago - var_fecha_factura
                     rsaux2.Open "select max(floa_dpg_descuento) as descuento from tb_descuentos_pronto_pago where vcha_emp_empresa_id = '" + var_empresa + "' and inte_dpg_limite_inferior <= " + Str(var_dias) + " and inte_dpg_limite_superior + " + Str(var_tolerancia) + " >= " + Str(var_dias), cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_descuento_sistema = IIf(IsNull(rsaux2!descuento), 0, rsaux2!descuento)
                     Else
                        var_descuento_sistema = 0
                     End If
                     rsaux2.Close
                     list_item.SubItems(3) = IIf(IsNull(rs!inte_car_PLAZO), 0, rs!inte_car_PLAZO)
                     list_item.SubItems(4) = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
                     list_item.SubItems(5) = Format(var_importe_factura, "###,##0.00")
                     list_item.SubItems(6) = Format(var_descuento_sistema, "###,##0.00")
                     list_item.SubItems(7) = Format(var_importe_factura - IIf(IsNull(rs!floa_sal_importe), 0, rs!floa_sal_importe), "###,##0.00")
                     list_item.SubItems(8) = Format(IIf(IsNull(rs!floa_sal_importe), 0, rs!floa_sal_importe), "###,##0.00")
                     list_item.SubItems(9) = Format(0, "###,##0.00")
                     list_item.SubItems(10) = Format(0, "###,##0.00")
                     list_item.SubItems(11) = var_dias + var_tolerancia
                     list_item.SubItems(12) = IIf(IsNull(rs!VCHA_MON_MONEDA_ID), "", rs!VCHA_MON_MONEDA_ID)
                     list_item.SubItems(13) = IIf(IsNull(rs!vcha_ser_serie_id), "", rs!vcha_ser_serie_id)
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
            txt_saldo = ""
            txt_total_aplicado = Format(0, "###,##0.00")
            lv_facturas.ListItems.Clear
            rs.Close
            MsgBox "El cliente no tiene pagos por aplicar", vbOKOnly, "ATENCION"
         End If
      Else
         rs.Close
         txt_clave_cliente = ""
         cmb_clientes = ""
         txt_saldo = ""
         txt_total_aplicado = Format(0, "###,##0.00")
         lv_facturas.ListItems.Clear
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub cmd_aplicar_Click()
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
Dim var_numero_pago As Double
Dim var_clave_moneda As String
Dim var_partida As Double
Dim var_relacion As String
Dim var_agente As String
Dim var_
If rs.State = 1 Then
   rs.Close
End If
If lv_pagos.ListItems.Count > 0 Then
   si = MsgBox("¿Deseas aplicar los pagos?", vbYesNo, "ATENCION")
   If si = 6 Then
      si = MsgBox("Confirmar la aplicación de los pagos", vbYesNo, "ATENCION")
      If si = 6 Then
         Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
         var_tipo_Cambio = lv_pagos.selectedItem.SubItems(9)
         var_clave_moneda = lv_pagos.selectedItem.SubItems(10)
         n = lv_pagos.ListItems.Count
         For i = 1 To n
            lv_pagos.ListItems.Item(i).Selected = True
            If lv_pagos.selectedItem.SubItems(8) = "*" Then
               var_serie = lv_pagos.selectedItem.SubItems(11)
               var_numero_pago = lv_pagos.selectedItem.SubItems(2)
               var_partida = lv_pagos.selectedItem.SubItems(14) * 1
            End If
         Next i
         n = lv_facturas.ListItems.Count
         For i = 1 To n
             lv_facturas.ListItems.Item(i).Selected = True
             If (lv_facturas.selectedItem.SubItems(10) * 1) > 0 Then
                var_serie_cargo = lv_facturas.selectedItem.SubItems(13)
                var_importe = lv_facturas.selectedItem.SubItems(10) * 1
                var_descuento = lv_facturas.selectedItem.SubItems(9) * 1
                var_importe_descuento = (var_importe * (var_descuento / 100))
                
                rs.Open "update tb_saldos set floa_sal_importe = floa_Sal_importe - " + Str(var_importe - var_importe_descuento) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + txt_clave_cliente + "' and vcha_car_documento = 'NA' and inte_car_numero = " + CStr(var_numero_pago) + " and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                rs.Open "update tb_estado_cuenta set floa_ecu_importe_abono =  floa_ecu_importe_abono - " + Str((var_importe - var_importe_descuento) * var_tipo_Cambio) + " where vcha_ecu_movimiento_Cargo = 'NA' and inte_ecu_numero_cargo = " + CStr(var_numero_pago) + " and vcha_ecu_Serie_cargo = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                rs.Open "insert into tb_estado_cuenta (vcha_emp_empresa_id, vcha_ecu_serie_cargo, vcha_ecu_movimiento_cargo, inte_ecu_numero_cargo, vcha_ecu_serie_abono, vcha_ecu_movimiento_abono, inte_ecu_numero_abono, floa_ecu_importe_Cargo, floa_ecu_importe_abono) values ('" + var_empresa + "', '" + var_serie_cargo + "' ,'" + Trim(lv_facturas.selectedItem) + "', " + lv_facturas.selectedItem.SubItems(1) + ",'" + var_serie + "' ,'PA'," + CStr(var_numero_pago) + ", 0, " + Str((var_importe - var_importe_descuento) * var_tipo_Cambio) + ")", cnn, adOpenDynamic, adLockOptimistic
                
                'var_cadena = " [TB_RELACION_COBRANZA] ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_RCO_FOLIO], [DTIM_RCO_FECHA_RELACION], [VCHA_AGE_AGENTE_ID], [VCHA_CLI_CLAVE_ID], [VCHA_RCO_CHEQUE], [DTIM_RCO_FECHA_CHEQUE], [FLOA_RCO_IMPORTE], [FLOA_RCO_DESCUENTO_OTORGADO], [INTE_CAR_NUMERO], [FLOA_CAR_IMPORTE], [FLOA_CAR_DESCUENTO_APLICADO], [INTE_RCO_PARTIDA], [INTE_RCO_DESCUENTO_APLICADO],"
                'var_cadena = " [VCHA_SER_SERIE_ID],[VCHA_CAR_DOCUMENTO], [VCHA_BAN_BANCO_ID], [CHAR_RCO_APLICADA]) Values "
                'var_cadena = " ('" + var_empresa + "', '', '" + var_relacion + "', " + var_fecha_relacion + ",'" + var_agente + "',  '" + var_cliente + "', '" + var_cheque + "', " + VAR_FECHA_CHEQUE + ", " + CStr(var_importe_pago) + ", "+CSTR(VAR_dESCUENTO_OTORGADO+", "+VAR_FACTURA+", "+CSTR(VAR_IMPORTE_FACTURA)+", " +CSTR(VAR_dESCUENTO_APLICADO+", "+CSTR(VAR_PARTIDA_APLICADA)+", @INTE_RCO_DESCUENTO_APLICADO,
     '@VCHA_SER_SERIE_ID,
     '@VCHA_CAR_DOCUMENTO,
     '@VCHA_BAN_BANCO_ID,
      '   '')

                
                
                
                If Trim(lv_facturas.selectedItem) = "FA" Then
                   rs.Open "select * from VW_DETALLE_FACTURACION_LINEAS WHERE VCHA_EMP_EMPRESA_ID =  '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + var_serie + "' AND VCHA_CAR_TIPO_DOCUMENTO = 'FA' and inte_Car_numero = " + lv_facturas.selectedItem.SubItems(1), cnn, adOpenDynamic, adLockOptimistic
                   var_banco = lv_pagos.selectedItem.SubItems(12)
                   var_cheque = lv_pagos.selectedItem.SubItems(13)
                   var_folio = lv_pagos.selectedItem
                   While Not rs.EOF
                         var_fecha_factura = CDate(Format(CStr(rs!DTIM_CAR_FECHA), "short date"))
                         var_cadena = "INSERT INTO TB_COMISIONES_APLICADAS ([VCHA_EMP_EMPRESA_ID], [VCHA_AGE_AGENTE_ID], [VCHA_CAR_TIPO_DOCUMENTO], [VCHA_SER_SERIE_ID], [INTE_CAR_NUMERO], [DTIM_CAR_FECHA], [FLOA_CAP_IMPORTE_FACTURA], [VCHA_RCO_FOLIO], [DTIM_CAP_FECHA_PAGO], [VCHA_LIN_LINEA_ID], [FLOA_CAP_IMPORTE_PARTICIPACION], [FLOA_CAP_PORCENTAJE_PARTICIPACION], [FLOA_COM_PORCENTAJE], [FLOA_CAP_IMPORTE_COMISION], [VCHA_BAN_BANCO_ID] , [VCHA_RCO_CHEQUE], [FLOA_CAP_IMPORTE_PAGO], [VCHA_CLI_CLAVE_ID])"
                         var_cadena = var_cadena + "Values ('" + var_empresa + "', '" + rs!vcha_age_agente_id + "', 'FA', '" + var_serie + "', " + lv_facturas.selectedItem.SubItems(1) + ", '" + Str(Day(var_fecha_factura)) + "/" + Str(Month(var_fecha_factura)) + "/" + Str(Year(var_fecha_factura)) + "', " + CStr(rs!FLOA_CAR_IMPORTE_NETO / rs!floa_car_tipo_cambio) + ", '" + var_folio + "', '" + Str(Day(Date)) + "/" + Str(Month(Date)) + "/" + Str(Year(Date)) + "', '" + rs!VCHA_LIN_LINEA_ID + "', " + CStr(rs!importe / rs!floa_car_tipo_cambio) + ", 0, 0, 0,'" + var_banco + "' ,'" + var_cheque + "', " + CStr((var_importe - var_importe_descuento) / (1 + (rs!floa_car_porcentaje_iva / 100))) + ", '" + txt_clave_cliente + "')"
                         rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                         rs.MoveNext
                   Wend
                   rs.Close
                End If
                
                If var_descuento > 0 Then
                   rs.Open "INSERT INTO TB_SALDOS_APLICAR ([VCHA_EMP_EMPRESA_ID], [VCHA_CAR_DOCUMENTO], [VCHA_CLI_CLAVE_ID], [INTE_CAR_NUMERO],[VCHA_SAP_DOCUMENTO_CARGO] , [INTE_SAP_NUMERO_CARGO] ,[VCHA_MON_MONEDA_ID], [DTIM_SAP_FECHA], [FLOA_SAP_DESCUENTO_1], [FLOA_SAP_DESCUENTO_2], [FLOA_SAP_DESCUENTO_3], [FLOA_SAP_IMPORTE], [VCHA_SAP_TIPO_SALDO], [CHAR_SAP_ESTATUS], [FLOA_SAP_TIPO_CAMBIO],[VCHA_SER_SERIE_ID]) Values  ('" + var_empresa + "', 'PA', '" + txt_clave_cliente + "', " + Str(lv_pagos.selectedItem.SubItems(2) * 1) + ", '" + Trim(lv_facturas.selectedItem) + "', " + Str(lv_facturas.selectedItem.SubItems(1) * 1) + ", '" + var_clave_moneda + "', " + CStr(Date) + ", " + Str(var_descuento) + ", 0, 0, " + CStr(var_importe_descuento * var_tipo_Cambio) + ", 'DF','', " + Str(var_tipo_Cambio) + ",'" + var_serie_cargo + "')", cnn, adOpenDynamic, adLockOptimistic
                End If
             End If
         Next i
         '''''''''''
         If Trim(txt_clave_cliente) <> "" Then
            rs.Open "select * from tb_clientes where vcha_cli_clave_id ='" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               cmb_clientes = rs!vcha_cli_nombre
               rs.Close
               rs.Open "select a.inte_rut_tolerancia from tb_rutas a, tb_clientes b where b.vcha_cli_clave_id = '" + txt_clave_cliente + "' and a.vcha_rut_ruta_id = b.vcha_rut_ruta_id", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_tolerancia = IIf(IsNull(rs!inte_rut_tolerancia), 0, rs!inte_rut_tolerancia)
               Else
                  var_tolerancia = 0
               End If
               rs.Close
               rs.Open "select * from vw_saldos_pagos_no_aplicados where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + txt_clave_cliente + "' and floa_sal_importe > 0 and char_sal_afectacion = '-'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_importe_total = 0
                  lv_pagos.ListItems.Clear
                  var_pago_seleccionado = 0
                  var_contador_pagos = 0
                  While Not rs.EOF
                     Set list_item = lv_pagos.ListItems.Add(, , rs!VCHA_RCO_FOLIO)
                     list_item.SubItems(1) = Format(IIf(IsNull(rs!DTIM_CAR_FECHA), "", rs!DTIM_CAR_FECHA), "Short Date")
                     If var_contador_pagos = 0 Then
                        var_fecha_pago = Format(rs!DTIM_CAR_FECHA, "Short Date")
                        var_descuento_agente = 0
                     End If
                     list_item.SubItems(2) = IIf(IsNull(rs!inte_car_numero), 0, rs!inte_car_numero)
                     list_item.SubItems(3) = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
                     var_importe_pago = IIf(IsNull(rs!floa_Rco_importe), 0, rs!floa_Rco_importe)
                     var_importe_saldo_pago = IIf(IsNull(rs!floa_sal_importe), 0, rs!floa_sal_importe)
                     var_importe_total = var_importe_total + var_importe_saldo_pago
                     list_item.SubItems(4) = Format(var_importe_pago, "###,##0.00")
                     list_item.SubItems(5) = Format((var_importe_pago - var_importe_saldo_pago), "###,##0.00")
                     list_item.SubItems(6) = Format((var_importe_saldo_pago), "###,##0.00")
                     list_item.SubItems(7) = Format(0, "###,##0.00")
                     list_item.SubItems(9) = IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                     list_item.SubItems(10) = IIf(IsNull(rs!VCHA_MON_MONEDA_ID), "", rs!VCHA_MON_MONEDA_ID)
                     list_item.SubItems(11) = IIf(IsNull(rs!vcha_ser_serie_id), "", rs!vcha_ser_serie_id)
                     list_item.SubItems(12) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
                     list_item.SubItems(13) = IIf(IsNull(rs!VCHA_RCO_CHEQUE), "", rs!VCHA_RCO_CHEQUE)
                     var_contador_pagos = var_contador_pagos + 1
                     rs.MoveNext
                  Wend
                  rs.Close
                  txt_saldo = Format(var_importe_total, "###,##0.00")
                  rs.Open "select * from vw_saldos_facturas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + txt_clave_cliente + "' and floa_sal_importe > 0", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     lv_facturas.ListItems.Clear
                     var_contador_facturas = 0
                     While Not rs.EOF
                        var_saldo = (IIf(IsNull(rs!floa_sal_importe), 0, rs!floa_sal_importe) - IIf(IsNull(rs!importe_saldo), 0, rs!importe_saldo))
                        If var_saldo > 0 Then
                           Set list_item = lv_facturas.ListItems.Add(, , rs!vcha_car_documento)
                           var_tipo_Cambio = IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                           var_importe_factura = IIf(IsNull(rs!FLOA_CAR_IMPORTE_NETO), 0, rs!FLOA_CAR_IMPORTE_NETO) / var_tipo_Cambio
                           list_item.SubItems(1) = IIf(IsNull(rs!inte_car_numero), "", rs!inte_car_numero)
                           list_item.SubItems(2) = IIf(IsNull(rs!DTIM_CAR_FECHA), "", Format(rs!DTIM_CAR_FECHA, "Short Date"))
                           var_fecha_factura = Format(rs!DTIM_CAR_FECHA, "Short Date")
                           var_dias = var_fecha_pago - var_fecha_factura
                           If var_dias <= 0 Then
                              rsaux2.Open "select max(floa_dpg_descuento) as descuento from tb_descuentos_pronto_pago where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rs.EOF Then
                                 var_descuento_sistema = IIf(IsNull(rsaux2!descuento), 0, rsaux2!descuento)
                              Else
                                 var_descuento_sistema = 0
                              End If
                              rsaux2.Close
                           Else
                              rsaux2.Open "select max(floa_dpg_descuento) as descuento from tb_descuentos_pronto_pago where vcha_emp_empresa_id = '" + var_empresa + "' and inte_dpg_limite_inferior <= " + Str(var_dias) + " and inte_dpg_limite_superior + " + Str(var_tolerancia) + " >= " + Str(var_dias), cnn, adOpenDynamic, adLockOptimistic
                              If Not rs.EOF Then
                                 var_descuento_sistema = IIf(IsNull(rsaux2!descuento), 0, rsaux2!descuento)
                              Else
                                 var_descuento_sistema = 0
                              End If
                              rsaux2.Close
                           End If
                           list_item.SubItems(3) = IIf(IsNull(rs!inte_car_PLAZO), 0, rs!inte_car_PLAZO)
                           list_item.SubItems(4) = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
                           list_item.SubItems(5) = Format(var_importe_factura, "###,##0.00")
                           list_item.SubItems(6) = Format(var_descuento_sistema, "###,##0.00")
                           list_item.SubItems(7) = Format(var_importe_factura - IIf(IsNull(rs!floa_sal_importe), 0, rs!floa_sal_importe), "###,##0.00")
                           list_item.SubItems(8) = Format(IIf(IsNull(rs!floa_sal_importe), 0, rs!floa_sal_importe), "###,##0.00")
                           list_item.SubItems(9) = Format(0, "###,##0.00")
                           list_item.SubItems(10) = Format(0, "###,##0.00")
                           list_item.SubItems(11) = var_dias + var_tolerancia
                           list_item.SubItems(12) = IIf(IsNull(rs!VCHA_MON_MONEDA_ID), "", rs!VCHA_MON_MONEDA_ID)
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
                  txt_saldo = ""
                  txt_total_aplicado = Format(0, "###,##0.00")
                  lv_facturas.ListItems.Clear
                  rs.Close
                  MsgBox "El cliente no tiene pagos por aplicar", vbOKOnly, "ATENCION"
               End If
            Else
               rs.Close
               txt_clave_cliente = ""
               cmb_clientes = ""
               txt_saldo = ""
               txt_total_aplicado = Format(0, "###,##0.00")
               lv_facturas.ListItems.Clear
               MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
            End If
         End If

         '''''''''''
         MsgBox "Se han terminado de aplicar los pagos", vbOKOnly, "ATENCION"
      End If
   End If
Else
    MsgBox "No existen pagos por aplicar", vbOKOnly, "ATENCION"
End If
End Sub

Private Sub cmd_notas_credito_Click()
   If rs.State = 1 Then
      rs.Close
   End If
   If Trim(txt_clave_cliente) <> "" Then
      frmnota_credito_saldos_descuento_financiero.txt_clave_cliente = txt_clave_cliente
      rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         frmnota_credito_saldos_descuento_financiero.txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      Else
         frmnota_credito_saldos_descuento_financiero.txt_nombre_agente = ""
      End If
      frmnota_credito_saldos_descuento_financiero.Show
   Else
      MsgBox "Se debe de seleccionar un cliente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   If rs.State = 1 Then
      rs.Close
   End If
   txt_clave_cliente = ""
   cmb_clientes = ""
   txt_saldo = ""
   lv_pagos.ListItems.Clear
   lv_facturas.ListItems.Clear
   txt_total_aplicado = ""
   txt_clave_cliente.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If Me.frm_lista.Visible = True Then
         Me.frm_lista.Visible = False
      Else
         If Me.frm_cantidad_aplicar.Visible = True Then
            Me.frm_cantidad_aplicar.Visible = False
         Else
            Unload Me
         End If
      End If
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 0
   frm_cantidad_aplicar.Visible = False
   frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If rs.State = 1 Then
      rs.Close
   End If
   Call activa_forma(var_activa_forma_asigna_pagos_no_aplicados)
   
End Sub

Private Sub lv_facturas_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F4 para indicar el importe a aplicar"
End Sub

Private Sub lv_facturas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 115 Then
      If var_pago_seleccionado > 0 Then
         lv_pagos.ListItems.Item(var_pago_seleccionado).Selected = True
         frm_cantidad_aplicar.Visible = True
         If lv_facturas.selectedItem.SubItems(10) * 1 > 0 Then
            txt_descuento = lv_facturas.selectedItem.SubItems(11)
            txt_descuento.Enabled = False
         Else
            txt_descuento = ""
            txt_descuento.Enabled = True
         End If
         txt_cantidad_aplicar = ""
         txt_cantidad_aplicar.SetFocus
      Else
         MsgBox "No se a seleccionado ningún pago", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub lv_facturas_KeyPress(KeyAscii As Integer)
   Frmmenu2.StatusBar1.Panels(1) = "Marque los pagos a aplicar presionando enter"
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
         txt_clave_cliente = lv_lista.selectedItem
         txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
      Else
         txt_clave_cliente = ""
         txt_nombre_cliente = ""
      End If
      txt_clave_cliente.SetFocus
      frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_pagos_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Marque los pagos a aplicar presionando enter"
End Sub

Private Sub lv_pagos_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim i, n As Integer
Dim var_fecha_pago As Date
Dim var_fecha_factura As Date
Dim var_descuento_sistema As Double
n = lv_facturas.ListItems.Count
var_fecha_pago = lv_pagos.selectedItem.SubItems(1)
For i = 1 To n
   lv_facturas.ListItems.Item(i).Selected = True
   var_fecha_factura = lv_facturas.selectedItem.SubItems(2)
   var_dias = var_fecha_pago - var_fecha_factura
   If var_dias <= 0 Then
      rsaux2.Open "select max(floa_dpg_descuento) as descuento from tb_descuentos_pronto_pago where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux2.EOF Then
         var_descuento_sistema = IIf(IsNull(rsaux2!descuento), 0, rsaux2!descuento)
      Else
         var_descuento_sistema = 0
      End If
      rsaux2.Close
   Else
      rsaux2.Open "select max(floa_dpg_descuento) as descuento from tb_descuentos_pronto_pago where vcha_emp_empresa_id = '" + var_empresa + "' and inte_dpg_limite_inferior <= " + Str(var_dias) + " and inte_dpg_limite_superior + " + Str(var_tolerancia) + " >= " + Str(var_dias), cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux2.EOF Then
         var_descuento_sistema = IIf(IsNull(rsaux2!descuento), 0, rsaux2!descuento)
      Else
         var_descuento_sistema = 0
      End If
      rsaux2.Close
   End If
   lv_facturas.selectedItem.SubItems(6) = Format(var_descuento_sistema, "###,##0.00")
   
Next i
End Sub

Private Sub lv_pagos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_pagos.selectedItem.SubItems(8) = "*" Then
         lv_pagos.selectedItem.SubItems(8) = ""
         var_pago_seleccionado = 0
         lv_pagos.selectedItem.ForeColor = &H80000012
         lv_pagos.selectedItem.ListSubItems.Item(1).ForeColor = &H80000012
         lv_pagos.selectedItem.ListSubItems.Item(2).ForeColor = &H80000012
         lv_pagos.selectedItem.ListSubItems.Item(3).ForeColor = &H80000012
         lv_pagos.selectedItem.ListSubItems.Item(4).ForeColor = &H80000012
         lv_pagos.selectedItem.ListSubItems.Item(5).ForeColor = &H80000012
         lv_pagos.selectedItem.ListSubItems.Item(6).ForeColor = &H80000012
         lv_pagos.selectedItem.ListSubItems.Item(7).ForeColor = &H80000012
         lv_pagos.selectedItem.ListSubItems.Item(8).ForeColor = &H80000012
      Else
         If var_pago_seleccionado > 0 Then
            MsgBox "Ya se encuentra un pago seleccionado", vbOKOnly, "ATENCION"
         Else
            lv_pagos.selectedItem.SubItems(8) = "*"
            var_pago_seleccionado = lv_pagos.selectedItem.Index
            var_serie = lv_pagos.selectedItem.SubItems(11)
            lv_pagos.selectedItem.ForeColor = &HFF0000
            lv_pagos.selectedItem.ListSubItems.Item(1).ForeColor = &HFF0000
            lv_pagos.selectedItem.ListSubItems.Item(2).ForeColor = &HFF0000
            lv_pagos.selectedItem.ListSubItems.Item(3).ForeColor = &HFF0000
            lv_pagos.selectedItem.ListSubItems.Item(4).ForeColor = &HFF0000
            lv_pagos.selectedItem.ListSubItems.Item(5).ForeColor = &HFF0000
            lv_pagos.selectedItem.ListSubItems.Item(6).ForeColor = &HFF0000
            lv_pagos.selectedItem.ListSubItems.Item(7).ForeColor = &HFF0000
            lv_pagos.selectedItem.ListSubItems.Item(8).ForeColor = &HFF0000
            If lv_facturas.ListItems.Count > 0 Then
               lv_facturas.SetFocus
            End If
         End If
      End If
   End If
End Sub

Private Sub lv_pagos_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_cantidad_aplicar_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If IsNumeric(txt_cantidad_aplicar) Then
         If txt_descuento.Enabled = False Then
            If (txt_descuento * 1) <= lv_facturas.selectedItem.SubItems(6) Then
               If (txt_cantidad_aplicar * 1) + (lv_facturas.selectedItem.SubItems(10) * 1) > (lv_facturas.selectedItem.SubItems(8) * 1) Then
                  MsgBox "La cantidad a aplicar exede el importe del saldo de la factura", vbOKOnly, "ATENCIO"
               Else
                  If (txt_total_aplicado * 1) + (txt_cantidad_aplicar * 1) <= (lv_pagos.selectedItem.SubItems(6) * 1) Then
                     lv_facturas.selectedItem.SubItems(9) = Format(txt_descuento, "###,##0.00")
                     lv_facturas.selectedItem.SubItems(10) = Format(txt_cantidad_aplicar + (lv_facturas.selectedItem.SubItems(10) * 1), "###,##0.00")
                     txt_total_aplicado = (txt_total_aplicado * 1) + (txt_cantidad_aplicar * 1)
                     lv_pagos.selectedItem.SubItems(7) = (lv_pagos.selectedItem.SubItems(7) * 1) + (txt_cantidad_aplicar * 1)
                  Else
                     MsgBox "La cantidad a aplicar exede al importe del saldo del pago del cliente", vbOKOnly, "ATENCION"
                  End If
               End If
            Else
               MsgBox "El descuento no puede ser mayor al descuento otorgado por el sistema", vbOKOnly, "ATENCION"
            End If
            lv_facturas.SetFocus
            frm_cantidad_aplicar.Visible = False
         Else
            txt_descuento.SetFocus
         End If
      Else
         txt_cantidad_aplicar = 0
      End If
   End If
End Sub

Private Sub txt_cantidad_aplicar_LostFocus()
   If IsNumeric(txt_cantidad_aplicar) Then
   Else
      MsgBox "Cantidad Incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_clave_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clave_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_clientes order by vcha_cli_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
            rs.MoveNext
      Wend
      rs.Close
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      lbl_lista = "CLIENTES"
      var_tipo_lista = 1
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clave_cliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_clave_cliente_LostFocus()
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
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_clave_cliente) <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select * from tb_clientes where vcha_cli_clave_id ='" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_cliente = rs!vcha_cli_nombre
         rs.Close
         rs.Open "select a.inte_rut_tolerancia from tb_rutas a, tb_clientes b where b.vcha_cli_clave_id = '" + txt_clave_cliente + "' and a.vcha_rut_ruta_id = b.vcha_rut_ruta_id", cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            var_tolerancia = 0
         Else
            var_tolerancia = IIf(IsNull(rs!inte_rut_tolerancia), 0, rs!inte_rut_tolerancia)
         End If
         rs.Close
         rs.Open "select * from vw_saldos_pagos_no_aplicados where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + txt_clave_cliente + "' and floa_sal_importe > 0 and char_sal_afectacion = '-'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_importe_total = 0
            lv_pagos.ListItems.Clear
            var_pago_seleccionado = 0
            var_contador_pagos = 0
            While Not rs.EOF
               Set list_item = lv_pagos.ListItems.Add(, , rs!VCHA_RCO_FOLIO)
               list_item.SubItems(1) = Format(IIf(IsNull(rs!DTIM_CAR_FECHA), "", rs!DTIM_CAR_FECHA), "Short Date")
               If var_contador_pagos = 0 Then
                  var_fecha_pago = Format(rs!DTIM_CAR_FECHA, "Short Date")
                  var_descuento_agente = 0
               End If
               list_item.SubItems(2) = IIf(IsNull(rs!inte_car_numero), 0, rs!inte_car_numero)
               list_item.SubItems(3) = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
               var_importe_pago = IIf(IsNull(rs!floa_Rco_importe), 0, rs!floa_Rco_importe)
               var_importe_saldo_pago = IIf(IsNull(rs!floa_sal_importe), 0, rs!floa_sal_importe)
               var_importe_total = var_importe_total + var_importe_saldo_pago
               list_item.SubItems(4) = Format(var_importe_pago, "###,##0.00")
               list_item.SubItems(5) = Format((var_importe_pago - var_importe_saldo_pago), "###,##0.00")
               list_item.SubItems(6) = Format((var_importe_saldo_pago), "###,##0.00")
               list_item.SubItems(7) = Format(0, "###,##0.00")
               list_item.SubItems(9) = IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
               list_item.SubItems(10) = IIf(IsNull(rs!VCHA_MON_MONEDA_ID), "", rs!VCHA_MON_MONEDA_ID)
               list_item.SubItems(11) = IIf(IsNull(rs!vcha_ser_serie_id), "", rs!vcha_ser_serie_id)
               list_item.SubItems(14) = IIf(IsNull(rs!inte_rco_partida), "", rs!inte_rco_partida)
               var_contador_pagos = var_contador_pagos + 1
               rs.MoveNext
            Wend
            rs.Close
            
            If var_contador_pagos > 6 Then
               lv_pagos.ColumnHeaders(2).Width = 1000.18
            Else
               lv_pagos.ColumnHeaders(2).Width = 1200.18
            End If
            
            
            txt_saldo = Format(var_importe_total, "###,##0.00")
            rs.Open "select * from vw_saldos_facturas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + txt_clave_cliente + "' and floa_sal_importe > 0", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               lv_facturas.ListItems.Clear
               var_contador_facturas = 0
               While Not rs.EOF
                  var_saldo = (IIf(IsNull(rs!floa_sal_importe), 0, rs!floa_sal_importe) - IIf(IsNull(rs!importe_saldo), 0, rs!importe_saldo))
                  If var_saldo > 0 Then
                     Set list_item = lv_facturas.ListItems.Add(, , rs!vcha_car_documento)
                     var_tipo_Cambio = IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                     var_importe_factura = IIf(IsNull(rs!FLOA_CAR_IMPORTE_NETO), 0, rs!FLOA_CAR_IMPORTE_NETO) / var_tipo_Cambio
                     list_item.SubItems(1) = IIf(IsNull(rs!inte_car_numero), "", rs!inte_car_numero)
                     list_item.SubItems(2) = IIf(IsNull(rs!DTIM_CAR_FECHA), "", Format(rs!DTIM_CAR_FECHA, "Short Date"))
                     var_fecha_factura = Format(rs!DTIM_CAR_FECHA, "Short Date")
                     PLAZO = IIf(IsNull(rs!inte_car_PLAZO), 0, rs!inte_car_PLAZO)
                     var_dias = var_fecha_pago - var_fecha_factura
                     If var_dias <= 0 Then
                        rsaux2.Open "select max(floa_dpg_descuento) as descuento from tb_descuentos_pronto_pago where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           var_descuento_sistema = IIf(IsNull(rsaux2!descuento), 0, rsaux2!descuento)
                        Else
                           var_descuento_sistema = 0
                        End If
                        rsaux2.Close
                     Else
                        rsaux2.Open "select max(floa_dpg_descuento) as descuento from tb_descuentos_pronto_pago where vcha_emp_empresa_id = '" + var_empresa + "' and inte_dpg_limite_inferior <= " + Str(var_dias) + " and inte_dpg_limite_superior + " + Str(var_tolerancia) + " >= " + Str(var_dias), cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           var_descuento_sistema = IIf(IsNull(rsaux2!descuento), 0, rsaux2!descuento)
                        Else
                           var_descuento_sistema = 0
                        End If
                        rsaux2.Close
                     End If
                     list_item.SubItems(3) = IIf(IsNull(rs!inte_car_PLAZO), 0, rs!inte_car_PLAZO)
                     list_item.SubItems(4) = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
                     list_item.SubItems(5) = Format(var_importe_factura, "###,##0.00")
                     list_item.SubItems(6) = Format(var_descuento_sistema, "###,##0.00")
                     list_item.SubItems(7) = Format(var_importe_factura - IIf(IsNull(rs!floa_sal_importe), 0, rs!floa_sal_importe), "###,##0.00")
                     list_item.SubItems(8) = Format(IIf(IsNull(rs!floa_sal_importe), 0, rs!floa_sal_importe), "###,##0.00")
                     list_item.SubItems(9) = Format(0, "###,##0.00")
                     list_item.SubItems(10) = Format(0, "###,##0.00")
                     list_item.SubItems(11) = var_dias + var_tolerancia
                     list_item.SubItems(12) = IIf(IsNull(rs!VCHA_MON_MONEDA_ID), "", rs!VCHA_MON_MONEDA_ID)
                     list_item.SubItems(13) = IIf(IsNull(rs!vcha_ser_serie_id), "", rs!vcha_ser_serie_id)
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
            txt_saldo = ""
            txt_total_aplicado = Format(0, "###,##0.00")
            lv_facturas.ListItems.Clear
            rs.Close
            MsgBox "El cliente no tiene pagos por aplicar", vbOKOnly, "ATENCION"
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
   If Me.lv_facturas.ListItems.Count > 10 Then
      lv_facturas.ColumnHeaders(5).Width = 949.73
   Else
      lv_facturas.ColumnHeaders(5).Width = 1149.73
   End If

End Sub

Private Sub txt_descuento_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If IsNumeric(txt_cantidad_aplicar) Then
         If Trim(txt_descuento) = "" Then
            txt_descuento = "0"
         End If
         If IsNumeric(txt_descuento) Then
            If (txt_descuento * 1) <= lv_facturas.selectedItem.SubItems(6) Then
               Dim var_cantidad_1 As Double
               Dim var_cantidad_2 As Double
               var_cantidad_1 = Round((txt_cantidad_aplicar * 1 / ((100 - txt_descuento) / 100)), 2) + (lv_facturas.selectedItem.SubItems(10) * 1)
               var_cantidad_2 = (lv_facturas.selectedItem.SubItems(8) * 1)
               var_x = Round((txt_cantidad_aplicar * 1 / ((100 - txt_descuento) / 100)), 2) + (lv_facturas.selectedItem.SubItems(10) * 1)
               If (lv_facturas.selectedItem.SubItems(8) * 1) < var_x Then
                  var_centavos = var_x - (lv_facturas.selectedItem.SubItems(8) * 1)
                  If var_centavos <= 0.05 Then
                     If (lv_facturas.selectedItem.SubItems(8) * 1) + var_centavos = var_x Then
                        var_x = var_x - var_centavos
                     End If
                  End If
               End If
               If var_x > (lv_facturas.selectedItem.SubItems(8) * 1) Then
                  MsgBox "La cantidad a aplicar exede el importe del saldo de la factura", vbOKOnly, "ATENCION"
               Else
                  If (txt_total_aplicado * 1) + (txt_cantidad_aplicar * 1) <= (lv_pagos.selectedItem.SubItems(6) * 1) Then
                     lv_facturas.selectedItem.SubItems(9) = Format(txt_descuento, "###,##0.00")
                     lv_facturas.selectedItem.SubItems(10) = Format(Round((txt_cantidad_aplicar * 1 / ((100 - txt_descuento) / 100)), 2) + (lv_facturas.selectedItem.SubItems(10) * 1), "###,##0.00")
                     txt_total_aplicado = (txt_total_aplicado * 1) + (txt_cantidad_aplicar * 1)
                     lv_pagos.selectedItem.SubItems(7) = (lv_pagos.selectedItem.SubItems(7) * 1) + (txt_cantidad_aplicar * 1)
                  Else
                     MsgBox "La cantidad a aplicar exede al importe del saldo del pago del cliente", vbOKOnly, "ATENCION"
                  End If
               End If
            Else
               MsgBox "El descuento no puede ser mayor al descuento otorgado por el sistema", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Descuento Incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Importe Incorrecto", vbOKOnly, "ATENCION"
      End If
      lv_facturas.SetFocus
   End If
End Sub

Private Sub txt_descuento_LostFocus()
   frm_cantidad_aplicar.Visible = False
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
End Sub
