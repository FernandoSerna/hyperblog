VERSION 5.00
Begin VB.Form frmcorreccion_importes_ORACLE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Corrección "
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " Correccion de importes de ORACLE "
      Height          =   1050
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   3870
      Begin VB.CommandButton cmd_ejecutar 
         Caption         =   "Ejecutar "
         Height          =   420
         Left            =   1965
         TabIndex        =   2
         Top             =   390
         Width           =   1755
      End
      Begin VB.TextBox txt_fecha 
         Height          =   360
         Left            =   360
         TabIndex        =   1
         Top             =   420
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmcorreccion_importes_ORACLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ejecutar_Click()
   If IsDate(Me.txt_fecha) Then
      var_si = MsgBox("¿Desea ejecutar el proceso de corrección del ORACLE?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la ejecución del proceso de corrección del oracle", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_dia = CStr(Day(Me.txt_fecha))
            var_mes = CStr(Month(Me.txt_fecha))
            var_año = CStr(Year(Me.txt_fecha))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            x = 1
            If x = 0 Then
            rsaux9.Open "select vcha_cli_clave_id, vcha_Car_documento, inte_Car_numero,dtim_car_fecha, floa_Car_importe_neto, vcha_aud_usuario, vcha_cli_clave_id from tb_encabezado_Cartera where vcha_Car_documento = 'PA' and vcha_ser_serie_id = 'FT' and dtim_Car_fecha >= " + var_fecha + " order by dtim_Car_fecha desc", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux9.EOF
                  
                  rsaux7.Open "select * from vw_clientes where vcha_cli_clave_id = '" + rsaux9!vcha_Cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux7.EOF Then
                     var_referencia_cliente_tienda = IIf(IsNull(rsaux7!vcha_cli_referencia), "", rsaux7!vcha_cli_referencia)
                     var_clave_cliente_tienda = IIf(IsNull(rsaux7!vcha_Cli_clave_id), "", rsaux7!vcha_Cli_clave_id)
                     var_agente_cliente_tienda = IIf(IsNull(rsaux7!vcha_age_agente_id), "", rsaux7!vcha_age_agente_id)
                     var_canal_cliente_tienda = IIf(IsNull(rsaux7!vcha_can_canal_venta_id), "", rsaux7!vcha_can_canal_venta_id)
                     var_grupo_real_tienda = IIf(IsNull(rsaux7!vcha_gre_grupo_real_id), "", rsaux7!vcha_gre_grupo_real_id)
                     var_grupo_actual_tienda = IIf(IsNull(rsaux7!vcha_gac_grupo_Actual_id), "", rsaux7!vcha_gac_grupo_Actual_id)
                     var_titular_tienda = IIf(IsNull(rsaux7!vcha_tit_titular_id), "", rsaux7!vcha_tit_titular_id)
                     var_porcentaje_iva_tienda = IIf(IsNull(rsaux7!FLOA_TPE_IVA), "", rsaux7!FLOA_TPE_IVA)
                     var_clave_moneda_tienda = IIf(IsNull(rsaux7!vcha_mon_moneda_id), "1", rsaux7!vcha_mon_moneda_id)
                  End If
                  rsaux7.Close
                  If rsaux9!vcha_aud_usuario = "U0000000019" Then
                     rs.Open "select * from tb_relacion_cobranza where inte_rco_pago = " + CStr(rsaux9!inte_car_numero), cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        rsaux7.Open "call SP_AGREGA_CARGO ('" + var_canal_cliente_tienda + "','" + var_agente_cliente_tienda + "', " + CStr(CDbl(rs!vcha_Rco_folio)) + ",'" + Trim(var_referencia_cliente_tienda) + "'," + CStr(CDbl(rsaux9!floa_car_importe_neto)) + ", " + CStr(CDbl(rsaux9!floa_car_importe_neto)) + ",TO_DATE('" + CStr(CDate(Me.txt_fecha)) + "','DD/MM/YYYY'),'VA')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                     Else
                        rsaux7.Open "call SP_AGREGA_CARGO ('" + var_canal_cliente_tienda + "','" + var_agente_cliente_tienda + "', " + CStr(CDbl(rsaux9!inte_car_numero)) + ",'" + Trim(var_referencia_cliente_tienda) + "'," + CStr(CDbl(rsaux9!floa_car_importe_neto)) + ", " + CStr(CDbl(rsaux9!floa_car_importe_neto)) + ",TO_DATE('" + CStr(CDate(Me.txt_fecha)) + "','DD/MM/YYYY'),'VA')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                     End If
                     rs.Close
                  Else
                     rsaux7.Open "call SP_AGREGA_CARGO ('" + var_canal_cliente_tienda + "','" + var_agente_cliente_tienda + "', " + CStr(CDbl(rsaux9!inte_car_numero)) + ",'" + Trim(var_referencia_cliente_tienda) + "'," + CStr(CDbl(rsaux9!floa_car_importe_neto)) + ", 0,TO_DATE('" + CStr(CDate(Me.txt_fecha)) + "','DD/MM/YYYY'),'VA')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux9.MoveNext
            Wend
            rsaux9.Close
            End If
            
            rsaux9.Open "select * from vw_pedidos_tiendas where dtim_ors_fecha_liberacion  >= " + var_fecha + " AND DTIM_ORS_FECHA_LIBERACION <= " + var_fecha + "+ 1 -.00000001  and inte_ped_pedido_credito = 0 order by dtim_ors_fecha_liberacion desc", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux9.EOF
                  rsaux7.Open "select * from vw_clientes where vcha_cli_clave_id = '" + rsaux9!vcha_Cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux7.EOF Then
                     var_referencia_cliente_tienda = IIf(IsNull(rsaux7!vcha_cli_referencia), "", rsaux7!vcha_cli_referencia)
                     var_clave_cliente_tienda = IIf(IsNull(rsaux7!vcha_Cli_clave_id), "", rsaux7!vcha_Cli_clave_id)
                     var_agente_cliente_tienda = IIf(IsNull(rsaux7!vcha_age_agente_id), "", rsaux7!vcha_age_agente_id)
                     var_canal_cliente_tienda = IIf(IsNull(rsaux7!vcha_can_canal_venta_id), "", rsaux7!vcha_can_canal_venta_id)
                     var_grupo_real_tienda = IIf(IsNull(rsaux7!vcha_gre_grupo_real_id), "", rsaux7!vcha_gre_grupo_real_id)
                     var_grupo_actual_tienda = IIf(IsNull(rsaux7!vcha_gac_grupo_Actual_id), "", rsaux7!vcha_gac_grupo_Actual_id)
                     var_titular_tienda = IIf(IsNull(rsaux7!vcha_tit_titular_id), "", rsaux7!vcha_tit_titular_id)
                     var_porcentaje_iva_tienda = IIf(IsNull(rsaux7!FLOA_TPE_IVA), "", rsaux7!FLOA_TPE_IVA)
                     var_clave_moneda_tienda = IIf(IsNull(rsaux7!vcha_mon_moneda_id), "1", rsaux7!vcha_mon_moneda_id)
                  End If
                  rsaux7.Close
                  rsaux7.Open "call SP_AGREGA_CARGO ('" + var_canal_cliente_tienda + "','" + var_agente_cliente_tienda + "', " + CStr(rsaux9!inte_ped_numero) + ",'" + Trim(rsaux9!vcha_cli_referencia) + "',0," + CStr(rsaux9!importe_pedido) + ", TO_DATE('" + CStr(CDate(Me.txt_fecha)) + "','DD/MM/YYYY'),'VA')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                  rsaux9.MoveNext
            Wend
            rsaux9.Close
         End If
      End If
   End If
End Sub

Private Sub Form_Load()
   cnn_clientes_tiendas.Open var_conexion_pedidos_tiendas
   cnn_clientes_tiendas.CursorLocation = adUseClient
   Top = 3000
   Left = 3500
   Me.txt_fecha = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         frmcalendario.Mes = CDate(Me.txt_fecha)
      Else
         frmcalendario.Mes = Date
      End If
      frmcalendario.Show 1
      Me.txt_fecha = var_fecha_general
   End If
End Sub
