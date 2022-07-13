VERSION 5.00
Begin VB.Form frmaplicacion_pagos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aplicaci�n de pagos"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   " Importe a aplicar "
      Height          =   735
      Left            =   150
      TabIndex        =   24
      Top             =   2580
      Width           =   7260
      Begin VB.TextBox txt_importe_aplicar 
         Height          =   375
         Left            =   5115
         TabIndex        =   14
         Top             =   225
         Width           =   1965
      End
      Begin VB.TextBox txt_comision 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2700
         TabIndex        =   13
         Top             =   225
         Width           =   1035
      End
      Begin VB.TextBox txt_folio_externo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   765
         TabIndex        =   12
         Top             =   225
         Width           =   1035
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Importe a aplicar:"
         Height          =   195
         Left            =   3855
         TabIndex        =   27
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Comisi�n:"
         Height          =   195
         Left            =   1980
         TabIndex        =   26
         Top             =   315
         Width           =   675
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   315
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos del Documento "
      Height          =   2070
      Left            =   150
      TabIndex        =   16
      Top             =   435
      Width           =   7215
      Begin VB.TextBox txt_documento 
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   300
         Width           =   390
      End
      Begin VB.TextBox txt_serie 
         Height          =   375
         Left            =   2385
         TabIndex        =   4
         Top             =   300
         Width           =   1110
      End
      Begin VB.TextBox txt_numero 
         Height          =   375
         Left            =   5010
         TabIndex        =   5
         Top             =   300
         Width           =   1560
      End
      Begin VB.TextBox txt_agente 
         Height          =   375
         Left            =   780
         TabIndex        =   6
         Top             =   720
         Width           =   1035
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   375
         Left            =   1845
         TabIndex        =   7
         Top             =   720
         Width           =   5280
      End
      Begin VB.TextBox txt_cliente 
         Height          =   375
         Left            =   780
         TabIndex        =   8
         Top             =   1125
         Width           =   1035
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   375
         Left            =   1845
         TabIndex        =   9
         Top             =   1125
         Width           =   5280
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   780
         TabIndex        =   10
         Top             =   1530
         Width           =   1035
      End
      Begin VB.TextBox txt_saldo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2685
         TabIndex        =   11
         Top             =   1530
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
         Height          =   195
         Left            =   135
         TabIndex        =   23
         Top             =   390
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   1755
         TabIndex        =   22
         Top             =   390
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "N�mero:"
         Height          =   195
         Left            =   4230
         TabIndex        =   21
         Top             =   390
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   135
         TabIndex        =   20
         Top             =   810
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   135
         TabIndex        =   19
         Top             =   1215
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   135
         TabIndex        =   18
         Top             =   1620
         Width           =   570
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
         Height          =   195
         Left            =   2175
         TabIndex        =   17
         Top             =   1620
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   60
      TabIndex        =   15
      Top             =   330
      Width           =   7395
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmaplicacion_pagos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmaplicacion_pagos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7020
      Picture         =   "frmaplicacion_pagos.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
End
Attribute VB_Name = "frmaplicacion_pagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
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
Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
If IsNumeric(txt_saldo) Then
   If CDbl(Me.txt_importe_aplicar) <= CDbl(Me.txt_saldo) Then
      rs.Open "select * from vw_clientes where vcha_cli_clave_id ='" + txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
         var_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
         var_grupo_actual = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
         var_grupo_real = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
         var_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
         var_plazo = IIf(IsNull(rs!inte_pla_dias), 0, rs!inte_pla_dias)
         var_iva = IIf(IsNull(rs!FLOA_TPE_IVA), 0, rs!FLOA_TPE_IVA)
      End If
      rs.Close
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
         var_si = MsgBox("�Desea aplicar el pago?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar la aplicaci�n del pago", vbYesNo, "ATENCION")
            If var_si = 6 Then
               cnn.CommandTimeout = 360
               cnn.BeginTrans
               rsaux5.Open "select * from TB_MAXIMO_PAGO", cnn_sid_quezada, adOpenDynamic, adLockOptimistic
               If rsaux5.EOF Then
                  var_numero_folio = 0
               Else
                  var_numero_folio = IIf(IsNull(rsaux5!inte_max_maximo_pago), 0, rsaux5!inte_max_maximo_pago)
               End If
               rsaux5.Close
               rsaux5.Open "update TB_MAXIMO_PAGO set inte_max_maximo_pago = inte_max_maximo_pago + 1", cnn_sid_quezada, adOpenDynamic, adLockOptimistic
               
               var_importe_neto = CDbl(txt_importe_aplicar) * var_tipo_Cambio
            
               var_importe_total = var_importe_neto
               
               If var_iva > 0 Then
                  var_importe_iva = var_importe_neto - (var_importe_neto / (1 + (var_iva / 100)))
               Else
                  var_importe_iva = 0
               End If
               var_subimporte = var_importe_neto - var_importe_iva
               var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "PA", "PA", "PA", CDbl(var_numero_folio), "-", "", "", 0, CStr(Date), CStr(var_agente), CStr(var_grupo_actual), CStr(var_grupo_real), CStr(var_titular), CStr(txt_cliente), "", 0, CDbl(var_iva), 0, 0, 0, 0, 0, CDbl(var_importe_total), CDbl(var_importe_iva), 0, 0, 0, 0, 0, CDbl(var_subimporte), CDbl(var_importe_neto), "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, CStr(var_clave_moneda), CDbl(var_tipo_Cambio), CStr(txt_serie), "")
               rsaux.Open "update tb_encabezado_cartera set VCHA_CAR_FOLIO_EXTERNO = '" + Me.txt_folio_externo + "', FLOA_CAR_COMISION = " + Me.txt_comision + " where vcha_emp_Empresa_id = '" + var_empresa + "' and vcha_car_documento = 'PA' and inte_Car_numero = " + Str(var_numero_folio) + " and vcha_ser_serie_id = '" + txt_serie + "'", cnn, adOpenDynamic, adLockOptimistic
               rsaux.Open "insert into tb_estado_cuenta (vcha_emp_empresa_id, vcha_ecu_serie_cargo, vcha_ecu_movimiento_cargo, inte_ecu_numero_cargo, vcha_ecu_serie_abono, vcha_ecu_movimiento_abono, inte_ecu_numero_abono, floa_ecu_importe_Cargo, floa_ecu_importe_abono) values ('" + var_empresa + "', '" + txt_serie + "' ,'" + txt_documento + "', " + txt_numero + ",'" + txt_serie + "' ,'PA'," + Str(var_numero_folio) + ", 0, " + Str(var_importe_neto) + ")", cnn, adOpenDynamic, adLockOptimistic
               If rs.State = 1 Then
                  rs.Close
               End If
               cnn.CommitTrans
               txt_documento = ""
               Me.txt_agente = ""
               Me.txt_cliente = ""
               Me.txt_documento = ""
               Me.txt_importe = ""
               Me.txt_nombre_agente = ""
               Me.txt_nombre_cliente = ""
               Me.txt_saldo = ""
               Me.txt_numero = ""
               Me.txt_serie = ""
               Me.txt_importe_aplicar = ""
               Me.txt_folio_externo = ""
               Me.txt_comision = ""
            End If
         End If
      Else
        MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "El importe a pagar no puede ser mayor al saldo del documento", vbOKOnly, "ATENCION"
   End If
Else
   MsgBox "Importe Incorrecto", vbOKOnly, "ATENCION"
End If
End Sub

Private Sub cmd_cancelar_pedidos_Click()
   Unload Me
End Sub

Private Sub cmd_nuevo_Click()
   txt_documento = ""
   Me.txt_agente = ""
   Me.txt_cliente = ""
   Me.txt_documento = ""
   Me.txt_importe = ""
   Me.txt_nombre_agente = ""
   Me.txt_nombre_cliente = ""
   Me.txt_saldo = ""
   Me.txt_numero = ""
   Me.txt_serie = ""
   Me.txt_folio_externo = ""
   Me.txt_comision = ""
   
   Me.txt_importe_aplicar = ""
   Me.txt_documento.SetFocus
End Sub

Private Sub Form_Load()
   Top = 2200
   Left = 2000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_comision_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_documento_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_documento_LostFocus()
   If Trim(txt_documento) <> "" Then
      If Trim(txt_documento) = "FA" Or Trim(txt_documento) = "NC" Or Trim(txt_documento) = "CH" Or Trim(txt_documento) = "CR" Then
      Else
         MsgBox "Documento incorrecto", vbOKOnly, "ATENCION"
         txt_documento = ""
      End If
   End If
End Sub

Private Sub txt_folio_externo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_importe_aplicar_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Me.cmd_aceptar_pedidos.SetFocus
   End If
End Sub

Private Sub txt_importe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
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
   If KeyAscii = 13 Then
      Me.txt_folio_externo.SetFocus
   End If
End Sub

Private Sub txt_numero_LostFocus()
   If Trim(txt_documento) <> "" Then
      If IsNumeric(txt_numero) Then
         rs.Open "SELECT * FROM TB_ENCABEZADO_cARTERA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_cAR_DOCUMENTO = '" + txt_documento + "' AND VCHA_sER_SERIE_ID = '" + txt_serie + "' AND INTE_cAR_NUMERO = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_importe = Format(IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,##0.00")
            txt_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
            txt_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
            rsaux4.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
            txt_nombre_agente = IIf(IsNull(rsaux4!VCHA_AGE_NOMBRE), "", rsaux4!VCHA_AGE_NOMBRE)
            rsaux4.Close
            rsaux4.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
            txt_nombre_cliente = IIf(IsNull(rsaux4!VCHA_CLI_NOMBRE), "", rsaux4!VCHA_CLI_NOMBRE)
            rsaux4.Close
            rsaux4.Open "SELECT * FROM TB_SALDOS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_cAR_DOCUMENTO = '" + txt_documento + "' AND VCHA_sER_SERIE_ID = '" + txt_serie + "' AND INTE_cAR_NUMERO = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
            txt_saldo = Format(IIf(IsNull(rsaux4!FLOA_sAL_IMPORTE), 0, rsaux4!FLOA_sAL_IMPORTE), "###,##0.00")
            rsaux4.Close
            txt_fecha_factura = IIf(IsNull(rs!dtim_Car_fecha), "", rs!dtim_Car_fecha)
         Else
            MsgBox "El documento no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "N�mero de documento incorrecto", vbOKOnly, "ATENCION"
         txt_numero = ""
      End If
   Else
      MsgBox "Falta indicar el documento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_saldo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

