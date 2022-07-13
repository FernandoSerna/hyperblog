VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmbonificaciones_financieras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bonificación Financiera"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   9375
   Begin VB.Frame frm_lista2 
      Height          =   2400
      Left            =   2100
      TabIndex        =   25
      Top             =   225
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista2 
         Height          =   1875
         Left            =   30
         TabIndex        =   26
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
         TabIndex        =   27
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2895
      Left            =   3405
      TabIndex        =   22
      Top             =   195
      Width           =   4050
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2430
         Left            =   30
         TabIndex        =   23
         Top             =   405
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   4286
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
         TabIndex        =   24
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8880
      Picture         =   "frmbonificacion_financiera.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos Generales "
      Height          =   1365
      Left            =   120
      TabIndex        =   4
      Top             =   465
      Width           =   9165
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   960
         Width           =   5445
      End
      Begin VB.TextBox txt_nombre_clase 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   270
         Width           =   5415
      End
      Begin VB.TextBox txt_clase 
         Height          =   315
         Left            =   1260
         TabIndex        =   19
         Top             =   270
         Width           =   1020
      End
      Begin VB.TextBox txt_fecha_relacion 
         Height          =   315
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   615
         Width           =   1725
      End
      Begin VB.TextBox txt_folio 
         Height          =   315
         Left            =   1260
         TabIndex        =   12
         Top             =   615
         Width           =   1020
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   315
         Left            =   1260
         TabIndex        =   10
         Top             =   960
         Width           =   1020
      End
      Begin VB.ComboBox cmb_series 
         Height          =   315
         Left            =   8280
         TabIndex        =   8
         Top             =   270
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento:"
         Height          =   195
         Left            =   345
         TabIndex        =   20
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Relación:"
         Height          =   195
         Left            =   2400
         TabIndex        =   14
         Top             =   675
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   345
         TabIndex        =   11
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   7830
         TabIndex        =   9
         Top             =   330
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         Height          =   195
         Left            =   345
         TabIndex        =   5
         Top             =   675
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmbonificacion_financiera.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmbonificacion_financiera.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalle de la Relación de Cobranza "
      Height          =   5340
      Left            =   135
      TabIndex        =   3
      Top             =   1860
      Width           =   9150
      Begin VB.Frame frm_descuento_correcto 
         Height          =   765
         Left            =   6330
         TabIndex        =   15
         Top             =   1635
         Width           =   1530
         Begin VB.TextBox txt_descuento_correcto 
            Height          =   360
            Left            =   90
            TabIndex        =   17
            Top             =   345
            Width           =   1155
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   1305
            TabIndex        =   18
            Top             =   398
            Width           =   120
         End
         Begin VB.Label Label5 
            BackColor       =   &H8000000D&
            Caption         =   "Descuento Correcto"
            ForeColor       =   &H8000000E&
            Height          =   225
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   1515
         End
      End
      Begin MSComctlLib.ListView lv_detalle 
         Height          =   5040
         Left            =   90
         TabIndex        =   6
         Top             =   210
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   8890
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
         NumItems        =   17
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Número"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Dias"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Importe"
            Object.Width           =   2364
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Pago       "
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Saldo"
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "% Agente"
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "% Aplicado"
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "% Correcto"
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Serie"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Descuento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Importe"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Saldo con descuento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "establecimiento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "IVA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "cheque"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Banco"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   90
      TabIndex        =   2
      Top             =   330
      Width           =   9195
   End
End
Attribute VB_Name = "frmbonificaciones_financieras"
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
Dim var_tolerancia_saldo As Double
Dim var_almacen As String
Dim var_cliente As String
Dim var_establecimiento As String


   


Private Sub cmb_series_Change()
   var_serie = cmb_series
End Sub

Private Sub cmd_imprimir_Click()
   Dim var_porcentaje As Double
   Dim var_imprimir As Boolean
   Dim var_contador As Integer
   Dim var_contador_notas As Integer
   Dim var_tipo_Cambio As Double
   Dim var_iva As Double
   Dim var_importe_total As Double
   Dim var_importe_iva As Double
   Dim var_subimporte As Double
   Dim var_importe As Double
   Dim var_saldo As Double
   Dim si As Integer
   Dim i, n As Integer
   Dim var_contador_renglones As Integer
   Dim var_numero_folio As Double
   Dim var_numero_nota As Double
   Dim var_numero_nota_anterior As Double
   Dim var_moneda_local As Integer
   Dim var_posible_tipo_cambio As Boolean
   Dim var_serie_cargo As String
   Dim var_cheque As String
   Dim var_banco As String
   Dim var_j As Integer
   Dim var_k As Integer
   Dim var_numero_nota_inicio As Double
   Dim var_contador_renglones_nota As Integer
   Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
   Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
   If Trim(txt_clase) <> "" Then
      var_moneda_local = 0
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
      If var_posible_tipo_cambio = True Then
         n = lv_detalle.ListItems.Count
         var_contador_renglones = 0
         var_contador_notas = 0
         For i = 1 To n
            lv_detalle.ListItems.Item(i).Selected = True
            If (lv_detalle.selectedItem.SubItems(7) * 1) > 0 Then
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
         rs.Open "select * from tb_Series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Ser_Serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
         var_numero_folio = IIf(IsNull(rs!inte_ser_nota_credito), 0, rs!inte_ser_nota_credito)
         rs.Close
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
            If si = 6 Then
               n = lv_detalle.ListItems.Count
               For i = 1 To n
                  lv_detalle.ListItems.Item(i).Selected = True
                  If lv_detalle.selectedItem.SubItems(11) = "" Then
                     lv_detalle.selectedItem.SubItems(11) = "0"
                  End If
                  If (lv_detalle.selectedItem.SubItems(11) * 1) > 0 Then
                     var_contador_renglones = var_contador_renglones + 1
                  End If
                  If var_contador_renglones = var_numero_renglones Then
                     var_contador_notas = var_contador_notas + 1
                     var_contador_renglones = 0
                  End If
               Next i
               If (var_contador_renglones > 0) And (var_contador_renglones < var_numero_renglones) Then
               End If
          '''''
               cnn.BeginTrans
               var_serie_cargo = lv_detalle.selectedItem.SubItems(9)
               var_insertar = False
               var_imprimir = False
               var_tolerancia_saldo = var_tolerancia_saldo / var_tipo_Cambio
               For i = 1 To n
                  lv_detalle.ListItems.Item(i).Selected = True
                  If lv_detalle.selectedItem.SubItems(12) = "" Then
                     lv_detalle.selectedItem.SubItems(12) = "0"
                  End If
                  var_saldo = lv_detalle.selectedItem.SubItems(12) * 1
                  var_iva = lv_detalle.selectedItem.SubItems(14) * 1
                  var_establecimiento = lv_detalle.selectedItem.SubItems(13)
                  If var_saldo < var_tolerancia_saldo Then
                     var_importe = var_importe + (lv_detalle.selectedItem.SubItems(11) * 1) + var_saldo
                  Else
                     var_importe = var_importe + (lv_detalle.selectedItem.SubItems(11) * 1)
                  End If
                  var_contador = var_contador + 1
                  If (var_contador = var_numero_renglones) Or (i = n) Then
                     var_contador = 0
                     var_imprimir = True
                  End If
                  If var_imprimir = True Then
                     var_subimporte = var_importe / (1 + (var_iva / 100))
                     var_importe_iva = var_importe - var_subimporte
                     var_numero_nota_ = var_numero_folio
                     var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "NC", "BF", txt_clase, var_numero_nota, "-", var_almacen, "", 0, Date, var_agente, var_grupo_actual, var_grupo_real, var_titular, txt_clave_cliente, var_establecimiento, 0, var_iva, 0, 0, 0, 0, 0, var_importe * var_tipo_Cambio, var_importe_iva * var_tipo_Cambio, 0, 0, 0, 0, 0, var_subimporte * var_tipo_Cambio, var_importe * var_tipo_Cambio, "", var_clave_usuario_global, "", Date, 0, Date, Date, var_clave_moneda, var_tipo_Cambio, var_serie, "")
                     rsaux3.Open "insert into tb_secuencia_notas_credito (vcha_emp_empresa_id, vcha_Ser_serie_id, inte_snc_numero_anterior, inte_snc_numero_actual) values ('" + var_empresa + "', '" + var_serie + "', " + CStr(var_numero_nota) + ", " + CStr(var_numero_nota) + ")", cnn, adOpenDynamic, adLockOptimistic
                     var_numero_nota = var_numero_nota + 1
                     var_contador_notas = var_contador_notas + 1
                  End If
                  var_imprimir = False
               Next i
               var_numero_nota = var_numero_nota_anterior
               var_importe = 0
               For i = 1 To n
                  lv_detalle.ListItems.Item(i).Selected = True
                  var_saldo = lv_detalle.selectedItem.SubItems(12) * 1
                  var_iva = lv_detalle.selectedItem.SubItems(14)
                  
                  If var_saldo < var_tolerancia_saldo Then
                     var_importe = (lv_detalle.selectedItem.SubItems(11) * 1) + var_saldo
                  Else
                     var_importe = (lv_detalle.selectedItem.SubItems(11) * 1)
                  End If
                  var_contador = var_contador + 1
                  If (var_contador = var_numero_renglones) Or (i = n) Then
                     var_contador = 0
                     var_imprimir = True
                  End If
                  var_cheque = lv_detalle.selectedItem.SubItems(15)
                  var_banco = lv_detalle.selectedItem.SubItems(16)
                  var_porcentaje = IIf(Not IsNumeric(lv_detalle.selectedItem.SubItems(10)), 0, lv_detalle.selectedItem.SubItems(10))
                  var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, var_serie_cargo, "FA", lv_detalle.selectedItem, var_serie, "BF", var_numero_nota, 0, (var_importe * var_tipo_Cambio))
                  rsaux2.Open "insert into tb_detalle_bonificacion_financiera (vcha_emp_empresa_id, vcha_car_documento, vcha_ser_serie_id, inte_Car_numero, inte_dbf_factura, floa_dbf_porcentaje, floa_dbf_importe, floa_dbf_iva) values ('" + var_empresa + "', 'BF', '" + var_serie + "', " + CStr(var_numero_nota) + ", " + lv_detalle.selectedItem + ", " + CStr(var_porcentaje) + ", " + CStr((var_importe * var_tipo_Cambio)) + ", " + CStr(var_iva) + ")"
                  rsaux3.Open "update tb_relacion_cobranza set inte_rco_numero_bonificacion_financiera = " + Str(var_numero_nota) + ", dtim_rco_fecha_bonificacion_financiera = getdate(), FLOA_RCO_BONIFICACION_FINANCIERA = " + CStr(var_porcentaje) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + txt_clave_cliente + "' and vcha_rco_folio = '" + txt_folio + "' and vcha_rco_cheque = '" + var_cheque + "' and vcha_ban_banco_id = '" + var_banco + "' and inte_car_numero = " + lv_detalle.selectedItem + " and vcha_ser_Serie_id = '" + var_serie_cargo + "' and vcha_Car_documento = 'FA'", cnn, adOpenDynamic, adLockOptimistic
                  If var_imprimir = True Then
                     var_subimporte = var_importe / (1 + (var_iva / 100))
                     var_importe_iva = var_importe - var_subimporte
                     var_numero_nota = var_numero_nota + 1
                  End If
                  var_imprimir = False
               Next i
               rs.Open "update tb_series set inte_ser_nota_credito = inte_ser_nota_credito + " + Str(var_contador_notas) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
'''''''''''''
          
               For var_k = var_numero_nota_inicio To var_numero_nota
                   rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'BF' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_k), cnn, adOpenDynamic, adLockOptimistic
                   If Not rs.EOF Then
'''''''''''''''IMPRESION DE LA NOTA DE CARGO
                      Open (App.Path & "\nota_credito" + Trim(Str(rs!inte_car_numero)) + ".txt") For Output As #1
                      Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                      Print #1, Spc(92); Str(rs!inte_car_numero)
                      Print #1, ""
                      Print #1, Spc(93); "       "; Format(rs!dtim_Car_fecha, "Short Date")
                      var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                      For var_j = 1 + Len(Trim(var_cliente)) To 83
                          var_cliente = var_cliente + " "
                      Next var_j
                      var_cliente = var_cliente + "AGUASCALIENTES, AGS."
                      Print #1, ""
                      Print #1, Spc(10); var_cliente
                      var_domicilio = Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION)) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                      For var_j = 1 + Len(Trim(var_domicilio)) To 83
                          var_domicilio = var_domicilio + " "
                      Next var_j
                      var_agente = ""
                      var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                      For var_j = 1 + Len(Trim(var_agente)) To 8
                          var_agente = var_agente + " "
                      Next var_j
                      var_agente = var_agente + Mid(IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE), 1, 30)
                      var_domicilio = var_domicilio
                      Print #1, Spc(10); var_domicilio
                      var_ciudad = ""
                      var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                      For var_j = 1 + Len(Trim(var_ciudad)) To 37
                          var_ciudad = var_ciudad + " "
                      Next var_j
                      var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                      For var_j = 1 + Len(Trim(var_estado)) To 46
                          var_estado = var_estado + " "
                      Next var_j
                      var_ciudad = var_ciudad + var_estado
                             
                      For var_j = 1 + Len(Trim(var_ciudad)) To 14
                          var_ciudad = var_ciudad + " "
                      Next var_j
                          
                      var_ciudad = var_ciudad + var_agente
                           
                      Print #1, Spc(10); var_ciudad
                      var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                      var_rfc = "RFC:  " + var_rfc
                      For var_j = 1 + Len(Trim(var_rfc)) To 89
                          var_rfc = var_rfc + " "
                      Next var_j
                      var_rfc = var_rfc
                      Print #1, Spc(4); var_rfc
                      Print #1, ""
                      Print #1, ""
                      Print #1, ""
                      Print #1, ""
                      var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                      If rsaux.State = 1 Then
                         rsaux.Close
                      End If
                      
                      rsaux.Open "select * from tb_detalle_bonificacion_financiera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'BF' and vcha_ser_serie_id = '" + var_serie + "' and inte_car_numero = " + CStr(var_k), cnn, adOpenDynamic, adLockOptimistic
                      var_contador_renglones_nota = 0
                      While Not rsaux.EOF
                         var_linea = "BF" + Str(rs!inte_car_numero) + " " + rs!vcha_Car_nombre + " FACTURA " + Str(rsaux!inte_dbf_factura) + " " + Format(rsaux!floa_dbf_porcentaje, "###,###,##0.0000") + "%"
                         If Len(Trim(var_linea)) < 108 Then
                            For var_j = 1 + Len(Trim(var_linea)) To 114
                                var_linea = var_linea + " "
                            Next var_j
                         End If
                         If Len(Trim(var_rfc)) = 0 Then
                            var_importe_str = Format((IIf(IsNull(rsaux!FLOA_DBF_IMPORTE), 0, rsaux!FLOA_DBF_IMPORTE)), "###,###,##0.00")
                         Else
                            var_importe_str = Format(((IIf(IsNull(rsaux!FLOA_DBF_IMPORTE), 0, rsaux!FLOA_DBF_IMPORTE)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)) / (1 + ((IIf(IsNull(rsaux!floa_dbf_iva), 1, rsaux!floa_dbf_iva) / 100)))), "###,###,##0.00")
                         End If
                         If Len(Trim(var_importe_str)) < 14 Then
                            For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                var_importe_str = " " + var_importe_str
                            Next var_j
                         End If
                         var_linea = var_linea + var_importe_str
                         Print #1, var_linea
                         rsaux.MoveNext
                         var_contador_renglones_nota = var_contador_renglones_nota + 1
                      Wend
                      rsaux.Close
                      If var_contador_renglones_nota < var_numero_renglones Then
                         For var_l = var_contador_renglones_nota To var_numero_renglones
                             Print #1, ""
                         Next var_l
                      End If
                      var_cantidad_letra = rs!vcha_car_importe_letra
                      var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                      If Len(Trim(var_linea)) < 99 Then
                         For var_j = 1 + Len(Trim(var_linea)) To 99
                             var_linea = var_linea + " "
                         Next var_j
                      End If
                      
                      
                      If Len(Trim(var_rfc)) = 0 Then
                         var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                         If Len(Trim(var_subimporte_str)) < 14 Then
                            For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                var_subimporte_str = " " + var_subimporte_str
                            Next var_j
                         End If
                         var_iva = "-"
                         For var_j = 1 + Len(Trim(var_iva_str)) To 14
                             var_iva_str = " " + var_iva_str
                         Next var_j
                      Else
                         var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                         If Len(Trim(var_subimporte_str)) < 14 Then
                            For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                var_subimporte_str = " " + var_subimporte_str
                            Next var_j
                         End If
                         var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                         If Len(Trim(var_iva_str)) < 14 Then
                            For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                var_iva_str = " " + var_iva_str
                            Next var_j
                         End If
                      End If
                      var_linea = var_linea + "           " + var_subimporte_str
                      Print #1, Spc(4); var_linea
                      Print #1, Spc(114); var_iva_str
                      var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                      If Len(Trim(var_importe_str)) < 14 Then
                         For var_j = 1 + Len(Trim(var_importe_str)) To 14
                             var_importe_str = " " + var_importe_str
                         Next var_j
                      End If
                      Print #1, Spc(114); var_importe_str
                      Print #1, ""
                      Print #1, ""
                      Print #1, ""
                      Print #1, Spc(85); "SISTEMAS"
                      Print #1, ""
                      Print #1, ""
                      Print #1, ""
                      Print #1, ""
                      Print #1, ""
                      Print #1, ""
                      Close #1
                      
                      Open (App.Path & "\nota_credito" + Trim(Str(rs!inte_car_numero)) + ".bat") For Output As #2
                      var_Archivo = App.Path & "\nota_credito" + Trim(Str(rs!inte_car_numero)) + ".bat"
                      Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!inte_car_numero)) + ".txt lpt1"
                      Close #2
                      x = Shell(var_Archivo, vbHide)
'''''''''''''''
                     End If
                     rs.Close
                   Next var_k
'''''''''''''
                
                MsgBox "Se a terminado de generar la nota de crédito", vbOKOnly, "ATENCION"
         '''''
            Else
               MsgBox "La impresión a sido cancelada", vbOKOnly, "ATENCION"
            End If
         End If
      Else
         MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a indicado una clase de movimiento", vbOKOnly, "ATENCION"
   End If
   cmd_imprimir.Enabled = False
End Sub

Private Sub cmd_nuevo_Click()
   cmd_imprimir.Enabled = True
   txt_folio = ""
   txt_clave_cliente = ""
   txt_nombre_cliente = ""
   txt_fecha_relacion = ""
   lv_detalle.ListItems.Clear
   txt_folio.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   frm_lista2.Visible = False
   cmd_imprimir.Enabled = False
   Top = 0
   Left = 1000
   rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   var_numero_renglones = rs!INTE_PRI_RENGLONES_NOTA_CREDITO
   var_tolerancia_saldo = rs!FLOA_PRI_TOLERANCIA_SALDOS
   rs.Close
   frm_descuento_correcto.Visible = False
   frm_lista.Visible = False
   rs.Open "select * from tb_clases_Cartera where vcha_car_documento = 'BF' order by vcha_car_nombre ", cnn, adOpenDynamic, adLockBatchOptimistic
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
         'txt_nombre_clase.Enabled = False
         'txt_clase.Enabled = False
      End If
      rs.MoveFirst
      txt_nombre_clase = rs!vcha_Car_nombre
      txt_clase = rs!vcha_Car_clase_id
   Else
      MsgBox "No se a indicado una clase de Bonificación", vbOKOnly, "ATENCION"
      txt_clase.Enabled = False
      txt_nombre_clase.Enabled = False
   End If
   rs.Close
   
   rs.Open "select vcha_ser_serie_id from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_contador_serie = 0
      While Not rs.EOF
         var_contador_serie = var_contador_serie + 1
         rs.MoveNext
      Wend
      rs.MoveFirst
      txt_folio.Enabled = True
      Call RecsetToCombo(cmb_series.hwnd, rs, 0)
      If var_contador_serie > 1 Then
         cmb_series.Enabled = True
      Else
         cmb_series.Enabled = False
      End If
      rs.MoveFirst
      cmb_series = rs!vcha_ser_Serie_id
      var_serie = rs!vcha_ser_Serie_id
   Else
      MsgBox "No se a indicado una serie para esta Unidad organizacional", vbOKOnly, "ATENCION"
      txt_folio.Enabled = False
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call activa_forma(var_activa_forma_bonificaciones_financieras)
End Sub

Private Sub lv_detalle_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 115 Then
      frm_descuento_correcto.Visible = True
      txt_descuento_correcto.SetFocus
   End If
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
      txt_clave_cliente.SetFocus
   End If
End Sub

Private Sub lv_lista2_LostFocus()
   frm_lista2.Visible = False
End Sub

Private Sub txt_clase_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clase_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      rs.Open "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where vcha_car_documento= 'BF' order by vcha_Car_nombre", cnn, adOpenDynamic, adLockOptimistic
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
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_clase_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_clave_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clave_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista2.ListItems.Clear
      rs.Open "select distinct vcha_cli_nombre, vcha_cli_clave_id from vw_relacion_cobranza_para_BF where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + txt_folio + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista2.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista2 = "CLIENTES"
      var_tipo_lista = 1
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
         rs.Close
         lv_detalle.SetFocus
      Else
         rs.Close
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
         txt_clave_cliente = ""
      End If
   End If
End Sub

Private Sub txt_clave_cliente_LostFocus()
   Dim var_fecha_relacion As Date
   Dim var_fecha_factura As Date
   Dim var_dias As Integer
   Dim list_item As ListItem
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_clave_cliente) <> "" Then
      Me.lv_detalle.ListItems.Clear
      rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      txt_nombre_cliente = rs!VCHA_CLI_NOMBRE
      var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
      var_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
      var_grupo_actual = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
      var_grupo_real = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
      var_plazo = 0
      rs.Close
      rs.Open "select * from VW_RELACION_COBRANZA_PARA_BF where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + txt_folio + "' and vcha_cli_clave_id = '" + txt_clave_cliente + "' and floa_sal_importe > 0.01", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
      var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
      Set list_item = lv_detalle.ListItems.Add(, , rs!inte_car_numero)
      list_item.SubItems(1) = IIf(IsNull(rs!dtim_Car_fecha), "", Format(rs!dtim_Car_fecha, "Short Date"))
      var_fecha_factura = rs!dtim_Car_fecha
      var_fecha_relacion = txt_fecha_relacion
      var_dias = var_fecha_relacion - var_fecha_factura
      list_item.SubItems(2) = var_dias
      list_item.SubItems(3) = IIf(IsNull(rs!floa_Car_importe), 0, rs!floa_Car_importe)
      list_item.SubItems(4) = IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe)
      list_item.SubItems(5) = IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE)
      list_item.SubItems(6) = IIf(IsNull(rs!FLOA_RCO_DESCUENTO_OTORGADO), 0, rs!FLOA_RCO_DESCUENTO_OTORGADO)
      list_item.SubItems(7) = IIf(IsNull(rs!floa_Car_descuento_aplicado), 0, rs!floa_Car_descuento_aplicado)
      list_item.SubItems(8) = ""
      list_item.SubItems(9) = IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id)
      list_item.SubItems(13) = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
      list_item.SubItems(14) = IIf(IsNull(rs!floa_rco_iva), 0, rs!floa_rco_iva)
         rs.MoveNext
      Wend
      rs.Close
   End If
End Sub

Private Sub txt_descuento_correcto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(txt_descuento_correcto) Then
         Dim var_importe_factura As Double
         Dim var_importe_relacion As Double
         Dim var_descuento_aplicado As Double
         Dim var_descuento_correcto As Double
         Dim var_descuento_agente As Double
         Dim var_importe_total_aplicado As Double
         Dim var_importe_total_correcto As Double
         Dim var_saldo_factura As Double
         Dim var_saldo_correcto As Double
         Dim var_posible_BF As Boolean
         Dim var_aplica As String
         rs.Open "select * from tb_descuentos_pronto_pago where vcha_emp_empresa_id = '" + var_empresa + "' and floa_dpg_descuento = " + txt_descuento_correcto, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_saldo_factura = lv_detalle.selectedItem.SubItems(5)
            var_importe_relacion = lv_detalle.selectedItem.SubItems(4) * 1
            If lv_detalle.selectedItem.SubItems(6) * 1 < lv_detalle.selectedItem.SubItems(7) * 1 Then
               var_descuento_aplicado = lv_detalle.selectedItem.SubItems(6) * 1
            Else
               var_descuento_aplicado = lv_detalle.selectedItem.SubItems(7) * 1
            End If
            var_porcentaje = Round((var_saldo_factura / lv_detalle.selectedItem.SubItems(3)) * 100, 2)
            'var_importe_total_aplicado = (var_importe_relacion * 100) / (100 - var_descuento_aplicado)
            'var_descuento_correcto = txt_descuento_correcto
            'var_importe_total_correcto = (var_importe_relacion * 100) / (100 - var_descuento_correcto)
            'var_diferencia = var_importe_total_correcto - var_importe_total_aplicado
            'var_porcentaje = (var_diferencia * 100) / var_importe_total_correcto
            'var_importe_descuento = var_importe_total_correcto - (var_importe_total_correcto * ((100 - var_porcentaje) / 100))
            
            
            var_posible_BF = False
            rsaux2.Open "select * from TB_RANGOS_DESCUENTOS_FINANCIEROS where FLOA_RDF_RANGO_INFERIOR <= " + Str(var_porcentaje) + " and FLOA_RDF_RANGO_SUPERIOR >= " + Str(var_porcentaje), cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               var_aplica = IIf(IsNull(rsaux2!CHAR_RDF_APLICA), "", rsaux2!CHAR_RDF_APLICA)
               If Trim(var_aplica) = "*" Then
                  var_posible_BF = True
               Else
                  var_posible_BF = False
               End If
            Else
               var_posible_BF = False
            End If
            rsaux2.Close
            If var_posible_BF = True Then
               lv_detalle.selectedItem.SubItems(8) = txt_descuento_correcto
               lv_detalle.selectedItem.SubItems(10) = var_porcentaje
               lv_detalle.selectedItem.SubItems(11) = Round(var_importe_descuento, 2)
               lv_detalle.selectedItem.SubItems(12) = var_saldo_factura - Round(var_importe_descuento, 2)
               
            Else
               MsgBox "El porcentaje faltante no entra dentro de los rangos para bonificación financiera", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El descuento indicado no se encuentra dentro de las politicas de descuento"
         End If
         rs.Close
      Else
         MsgBox "Descuento Incorrecto", vbOKOnly, "ATENCION"
      End If
      txt_descuento_correcto = ""
      frm_descuento_correcto.Visible = False
   If Me.lv_detalle.ListItems.Count > 0 Then
      Me.lv_detalle.SetFocus
   End If
   End If
   If KeyAscii = 27 Then
      frm_descuento_correcto.Visible = False
   End If
End Sub

Private Sub txt_folio_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_clave_cliente.SetFocus
   End If
End Sub

Private Sub txt_folio_LostFocus()
   If Trim(txt_folio) <> "" Then
      rs.Open "select * from vw_relacion_cobranza_para_BF where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + txt_folio + "'order by vcha_cli_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      If Not rs.EOF Then
         txt_fecha_relacion = Format(rs!dtim_rco_fecha_relacion, "Short Date")
      Else
         MsgBox "La relación de cobranza no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub txt_nombre_clase_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      rs.Open "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where vcha_car_documento= 'BF' order by vcha_Car_nombre", cnn, adOpenDynamic, adLockOptimistic
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
