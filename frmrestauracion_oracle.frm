VERSION 5.00
Begin VB.Form frmrestauracion_oracle 
   Caption         =   "Restauracion de pedidos y cargos del ORACLE"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1290
   ScaleWidth      =   4470
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4050
      Picture         =   "frmrestauracion_oracle.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmrestauracion_oracle.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   15
      TabIndex        =   5
      Top             =   360
      Width           =   4485
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   75
      TabIndex        =   0
      Top             =   450
      Width           =   4335
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2700
         TabIndex        =   2
         Top             =   255
         Width           =   1095
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   690
         TabIndex        =   1
         Top             =   255
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2385
         TabIndex        =   4
         Top             =   315
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   315
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmrestauracion_oracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_imprimir_Click()
   Dim var_clave_movimiento As String
   Dim var_fecha_inicio As String, var_fecha_fin As String
   Dim var_consecutivo As Double
   
   var_clave_movimiento = ""
   rs.Open "select isnull(vcha_mov_movimiento_id,'') from tb_movimientos where inte_mov_reempaque = 2", cnn, adOpenDynamic, adLockOptimistic
   If rs.EOF Then
      var_clave_movimiento = ""
   Else
      var_clave_movimiento = rs(0).Value
   End If
   rs.Close
   If Trim(var_clave_movimiento) <> "" Then
      If IsDate(txt_inicio) Then
         If IsDate(txt_fin) Then
            If CDate(txt_inicio) <= CDate(txt_fin) Then
               var_fecha_fin_1 = CDate(txt_fin) + 1
               var_dia = CStr(Day(CDate(txt_inicio)))
               var_mes = CStr(Month(CDate(txt_inicio)))
               var_año = CStr(Year(CDate(txt_inicio)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
               var_dia = CStr(Day(var_fecha_fin_1))
               var_mes = CStr(Month(var_fecha_fin_1))
               var_año = CStr(Year(var_fecha_fin_1))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
               
               var_fecha_fin_1 = CDate(txt_fin)
               var_dia = CStr(Day(var_fecha_fin_1))
               var_mes = CStr(Month(var_fecha_fin_1))
               var_año = CStr(Year(var_fecha_fin_1))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               
               var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
               
               var_cadena = "SELECT     TOP 100 PERCENT dbo.TB_CLIENTES.VCHA_CLI_REFERENCIA, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_PEDIDO_CREDITO FROM         dbo.TB_ENCABEZADO_CARTERA INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND "
               var_cadena = var_cadena + "  dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND  dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO ON                       dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO INNER JOIN dbo.TB_ENCABEZADO_PEDIDOS ON                      dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO = dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO INNER JOIN"
               var_cadena = var_cadena + " dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID WHERE     (dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID = 'FT') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS = 'C') AND (dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_PEDIDO_CREDITO <> 1) and dtim_car_fecha >= " + var_fecha_inicio + " and dtim_Car_fecha <= " + var_fecha_fin + " ORDER BY dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO DESC, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA             "
               rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               
               While Not rs.EOF
                     var_fecha = CDate(Format(rs!DTIM_car_FECHA, "DD/MM/YYYY"))
                     rsaux.Open "CALL SP_AGREGA_ABONO('" + Trim(rs!vcha_cli_referencia) + "'," + CStr(rs!floa_car_importe_neto) + ", 0,TO_DATE('" + CStr(var_fecha) + "','DD/MM/YYYY'),TO_DATE('" + CStr(var_fecha) + "','DD/MM/YYYY'),'" + CStr(rs!inte_car_numero) + "','','CF','Cancelacion de facturación')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               
               rs.Close
               x = 0
               If x = 1 Then
               var_cadena = "EXEC SP_RESTRUCTURA_FACTURAS_ORACLE " + var_fecha_inicio + "," + var_fecha_fin
               rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               
               var_cadena = "SELECT     TOP 100 PERCENT dbo.VW_IMPORTES_SEGURO_PAQUETERIA.IMPORTE_PEDIDO, dbo.VW_IMPORTES_SEGURO_PAQUETERIA.IMPORTE_SEGURO, dbo.VW_IMPORTES_SEGURO_PAQUETERIA.IMPORTE_PAQUETERIA, dbo.VW_IMPORTES_SEGURO_PAQUETERIA.FLOA_PAQ_COSTO_REFERENCIA, dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_LIBERACION, dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO,"
               var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_PEDIDO_CREDITO , dbo.TB_CLIENTES.VCHA_CLI_REFERENCIA, dbo.tb_clientes.vcha_cli_clave_id, dbo.tb_clientes.vcha_age_agente_id FROM         dbo.VW_IMPORTES_SEGURO_PAQUETERIA INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO ON                    dbo.VW_IMPORTES_SEGURO_PAQUETERIA.INTE_PED_NUMERO = dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO INNER JOIN dbo.TB_ENCABEZADO_PEDIDOS ON"
               var_cadena = var_cadena + " dbo.VW_IMPORTES_SEGURO_PAQUETERIA.INTE_PED_NUMERO = dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENC_ORDEN_SURTIDO.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID WHERE     (dbo.VW_IMPORTES_SEGURO_PAQUETERIA.DTIM_ORS_FECHA_LIBERACION >= " + var_fecha_inicio + ") AND"
               var_cadena = var_cadena + " (dbo.VW_IMPORTES_SEGURO_PAQUETERIA.DTIM_ORS_FECHA_LIBERACION <= " + var_fecha_fin + ") AND (dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_PEDIDO_CREDITO = 0) ORDER BY dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_LIBERACION"
               rs.Open var_cadena
               While Not rs.EOF
                    
                     var_importe_neto = rs!importe_pedido + rs!importe_seguro + rs!importe_paqueteria + rs!floa_paq_costo_referencia
                     rsaux9.Open "select vcha_can_canal_venta_id from tb_agentes where vcha_age_agente_id = '" + rs!vcha_age_agente_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_canal = rsaux9!vcha_can_canal_venta_id
                     rsaux9.Close
                     var_fecha = CDate(Format(rs!dtim_ors_fecha_liberacion, "DD/MM/YYYY"))
                     VAR_FECHA_inicio_c = CDate(Format(CDate(Me.txt_inicio), "DD/MM/YYYY"))
                     
                     rsaux9.Open "SELECT * FROM TB_CARGO WHERE VCHA_CAR_NUM_DOCUM = '" + CStr(rs!inte_ped_numero) + "' AND Date_CAR_FECHA_CARGO < to_date('" + CStr(VAR_FECHA_inicio_c) + "','DD/MM/YYYY')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                     If rsaux9.EOF Then
                        rsaux7.Open "call SP_AGREGA_CARGO ('" + var_canal + "','" + rs!vcha_age_agente_id + "', " + CStr(rs!inte_ped_numero) + ",'" + Trim(rs!vcha_cli_referencia) + "',0," + CStr(var_importe_neto) + ", TO_DATE('" + CStr(var_fecha) + "','DD/MM/YYYY'),'VA')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux9.Close
                     rs.MoveNext
               Wend
               rs.Close
               
               ' para subir los pagos de los de credito
               
              var_cadena = "SELECT     TOP 100 PERCENT dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, dbo.TB_ESTADO_CUENTA.FLOA_ECU_IMPORTE_ABONO, dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO,  dbo.TB_ESTADO_CUENTA.FLOA_ECU_IMPORTE_CARGO, dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_AGE_AGENTE_ID, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_PEDIDO_CREDITO, TB_ENCABEZADO_CARTERA_1.DTIM_CAR_FECHA, dbo.TB_CLIENTES.VCHA_CLI_REFERENCIA AS REFERENCIA FROM         dbo.TB_ESTADO_CUENTA INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_ESTADO_CUENTA.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_ESTADO_CUENTA.VCHA_ECU_SERIE_CARGO = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_ESTADO_CUENTA.VCHA_ECU_MOVIMIENTO_CARGO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND  dbo.TB_ESTADO_CUENTA.INTE_ECU_NUMERO_CARGO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON "
                      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND  dbo.TB_ENCABEZADO_CARTERA.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN  dbo.TB_ENC_ORDEN_SURTIDO ON                       dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO INNER JOIN dbo.TB_ENCABEZADO_PEDIDOS ON                       dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO = dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO INNER JOIN dbo.TB_ENCABEZADO_CARTERA TB_ENCABEZADO_CARTERA_1 ON dbo.TB_ESTADO_CUENTA.VCHA_EMP_EMPRESA_ID = TB_ENCABEZADO_CARTERA_1.VCHA_EMP_EMPRESA_ID AND "
                      var_cadena = var_cadena + " dbo.TB_ESTADO_CUENTA.VCHA_ECU_SERIE_ABONO = TB_ENCABEZADO_CARTERA_1.VCHA_SER_SERIE_ID AND dbo.TB_ESTADO_CUENTA.VCHA_ECU_MOVIMIENTO_ABONO = TB_ENCABEZADO_CARTERA_1.VCHA_CAR_DOCUMENTO AND dbo.TB_ESTADO_CUENTA.INTE_ECU_NUMERO_ABONO = TB_ENCABEZADO_CARTERA_1.INTE_CAR_NUMERO INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID WHERE     (dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID = 'FT') AND (dbo.TB_ESTADO_CUENTA.FLOA_ECU_IMPORTE_CARGO = 0) AND (dbo.TB_ESTADO_CUENTA.VCHA_ECU_MOVIMIENTO_ABONO = 'PA') AND (dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_PEDIDO_CREDITO = 1) AND (TB_ENCABEZADO_CARTERA_1.DTIM_CAR_FECHA >= " + var_fecha_inicio + ") AND (TB_ENCABEZADO_CARTERA_1.DTIM_CAR_FECHA <= " + var_fecha_fin + ") ORDER BY dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO             "
                              
               rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     rsaux9.Open "select vcha_can_canal_venta_id from tb_agentes where vcha_age_agente_id = '" + rs!vcha_age_agente_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_canal = rsaux9!vcha_can_canal_venta_id
                     rsaux9.Close
                     var_fecha = CDate(Format(rs!DTIM_car_FECHA, "DD/MM/YYYY"))

                     rsaux7.Open "call SP_AGREGA_CARGO ('" + var_canal + "','" + rs!vcha_age_agente_id + "', " + CStr(CDbl(rs!inte_car_numero)) + ",'" + Trim(rs!Referencia) + "'," + CStr(CDbl(rs!floa_ecu_importe_abono)) + ", " + CStr(CDbl(rs!floa_ecu_importe_abono)) + ",TO_DATE('" + CStr(var_fecha) + "','DD/MM/YYYY'),'PC')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rs.Close
               
               
               
               
               
               'para subir los de contado
               
        var_cadena = "SELECT     TOP 100 PERCENT dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, dbo.TB_ESTADO_CUENTA.FLOA_ECU_IMPORTE_ABONO,dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO, dbo.TB_ESTADO_CUENTA.FLOA_ECU_IMPORTE_CARGO, dbo.RESTAURAR_MOVIMIENTOS.IMPORTE_MOVIMIENTO, dbo.RESTAURAR_MOVIMIENTOS.FECHA_MOVIMIENTO, dbo.RESTAURAR_MOVIMIENTOS.REFERENCIA, dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_AGE_AGENTE_ID, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_PEDIDO_CREDITO"
        var_cadena = var_cadena + " FROM         dbo.TB_ESTADO_CUENTA INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_ESTADO_CUENTA.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_ESTADO_CUENTA.VCHA_ECU_SERIE_CARGO = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_ESTADO_CUENTA.VCHA_ECU_MOVIMIENTO_CARGO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_ESTADO_CUENTA.INTE_ECU_NUMERO_CARGO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.RESTAURAR_MOVIMIENTOS ON                       dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO = dbo.RESTAURAR_MOVIMIENTOS.MOVIMIENTO INNER JOIN "
        var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS ON dbo.RESTAURAR_MOVIMIENTOS.PEDIDO = dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO WHERE     (dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID = 'FT') AND (dbo.TB_ESTADO_CUENTA.FLOA_ECU_IMPORTE_CARGO = 0) AND                       (dbo.TB_ESTADO_CUENTA.VCHA_ECU_MOVIMIENTO_ABONO = 'PA') AND (dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_PEDIDO_CREDITO = 0) ORDER BY dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO"
               
               
               
               rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     rsaux9.Open "select vcha_can_canal_venta_id from tb_agentes where vcha_age_agente_id = '" + rs!vcha_age_agente_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_canal = rsaux9!vcha_can_canal_venta_id
                     rsaux9.Close
                     var_fecha = CDate(Format(rs!FECHA_MOVIMIENTO, "DD/MM/YYYY"))
                     rsaux7.Open "call SP_AGREGA_CARGO ('" + var_canal + "','" + rs!vcha_age_agente_id + "', " + CStr(CDbl(rs!inte_car_numero)) + ",'" + Trim(rs!Referencia) + "'," + CStr(CDbl(rs!floa_ecu_importe_abono)) + ", 0,TO_DATE('" + CStr(var_fecha) + "','DD/MM/YYYY'),'VA')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rs.Close
               
               
               
               
               
               
               
               
               
               rs.Open "select * from RESTAURAR_MOVIMIENTOS where diferencia > 0", cnn, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     var_fecha = CDate(Format(rs!FECHA_MOVIMIENTO, "DD/MM/YYYY"))
                     rsaux8.Open "CALL SP_AGREGA_ABONO('" + Trim(rs!Referencia) + "',0.00," + CStr(rs!diferencia) + ",TO_DATE('" + CStr(var_fecha) + "','DD/MM/YYYY'),TO_DATE('" + CStr(var_fecha) + "','DD/MM/YYYY'),'" + CStr(rs!pedido) + "','','DF','')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rs.Close
               
               
               var_cadena = "SELECT     dbo.TB_SORTEO_BOLETOS_PREMIO.INTE_SOR_BOLETO, dbo.TB_SORTEO_BOLETOS_PREMIO.INTE_SOR_PREMIO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_CLIENTES.VCHA_AGE_AGENTE_ID, dbo.TB_CLIENTES.VCHA_CLI_REFERENCIA fROM         dbo.TB_SORTEO_BOLETOS_PREMIO INNER JOIN  dbo.TB_CLIENTES ON dbo.TB_SORTEO_BOLETOS_PREMIO.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID INNER JOIN"
               var_cadena = var_cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS ON Dbo.TB_SORTEO_BOLETOS_PREMIO.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND          dbo.TB_SORTEO_BOLETOS_PREMIO.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_SORTEO_BOLETOS_PREMIO.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_SORTEO_BOLETOS_PREMIO.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND"
               var_cadena = var_cadena + "  dbo.TB_SORTEO_BOLETOS_PREMIO.INTE_EMO_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO WHERE     (dbo.TB_SORTEO_BOLETOS_PREMIO.INTE_SOR_PREMIO = 1) AND     (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND  (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA < " + var_fecha_fin + ")"
               rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     var_fecha = CDate(Format(rs!dtim_emo_fecha, "DD/MM/YYYY"))
                     rsaux.Open "CALL SP_AGREGA_ABONO('" + Trim(rs!vcha_cli_referencia) + "',400, 400,TO_DATE('" + CStr(var_fecha) + "','DD/MM/YYYY'),TO_DATE('" + CStr(var_fecha) + "','DD/MM/YYYY'),'','','PS','Premio de sorteo')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rs.Close
               End If
          Else
               MsgBox "La fecha de inicio debe de ser menor a la fecha final", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha de Inicio Incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No existe un movimiento de reempaque", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   If cnn_clientes_tiendas.State = 0 Then
      cnn_clientes_tiendas.Open var_conexion_pedidos_tiendas
      cnn_clientes_tiendas.CursorLocation = adUseClient
   End If
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes.Value = CDate(Me.txt_fin)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes.Value = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub
