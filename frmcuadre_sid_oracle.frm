VERSION 5.00
Begin VB.Form frmcuadre_sid_oracle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuadrar SID ORACLE"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " Referencia "
      Height          =   1095
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   4410
      Begin VB.TextBox txt_referencia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   210
         TabIndex        =   1
         Top             =   360
         Width           =   4020
      End
   End
End
Attribute VB_Name = "frmcuadre_sid_oracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Top = 3000
   Left = 3500
   If cnn_clientes_tiendas.State = 0 Then
      cnn_clientes_tiendas.Open var_conexion_pedidos_tiendas
      cnn_clientes_tiendas.CursorLocation = adUseClient
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub txt_referencia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_referencia <> "" Then
         rsaux.Open "select vcha_Cli_clave_id from tb_clientes where vcha_cli_referencia = '" + Me.txt_referencia + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_clave_cliente = IIf(IsNull(rsaux(0).Value), "", rsaux(0).Value)
         Else
            var_clave_cliente = ""
         End If
         rsaux.Close
         If var_clave_cliente <> "" Then
            var_x = 0
            If var_x = 1 Then
            rs.Open "delete from TB_TEMP_CUADRE_ORACLE_SID", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "select  vcha_car_tipo_documento, vcha_Car_num_docum, date_car_fecha_cargo, numb_car_importe, numb_car_importe_disponible from tb_Cargo where vcha_car_referencia = '" + Me.txt_referencia + "' AND vcha_car_tipo_documento = 'VA' order by date_Car_Fecha_Cargo", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  var_dia = CStr(Day(rs!date_car_fecha_cargo))
                  var_mes = CStr(Month(rs!date_car_fecha_cargo))
                  var_año = CStr(Year(rs!date_car_fecha_cargo))
                  If Len(Trim(var_dia)) = 1 Then
                      var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"

                  rsaux.Open "insert into TB_TEMP_CUADRE_ORACLE_SID (inte_tem_numero, DTIM_TEM_FECHA, floa_Tem_importe_oracle, FLOA_tEM_IMPORTE_DISPONIBLE_ORACLE) values (" + CStr(rs!vcha_Car_num_docum) + ", " + var_fecha_inicio + "," + CStr(rs!numb_car_importe) + ", " + CStr(rs!numb_car_importe_disponible) + ") ", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            rs.Open "select * from tb_temp_cuadre_oracle_sid", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux1.Open "select * from tb_encabezado_cartera where inte_Car_numero =  " + CStr(rs!inte_tem_numero) + " and vcha_Car_documento = 'FA' and vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                     rsaux2.Open "update tb_temp_cuadre_oracle_sid set floa_tem_importe_sid = " + CStr(rsaux1!FLOA_CAR_IMPORTE_NETO) + " where inte_tem_numero = " + CStr(rs!inte_tem_numero), cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux1.Close
                  rsaux1.Open "select * from vw_pedidos_tiendas where inte_ped_numero = " + CStr(rs!inte_tem_numero), cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                     rsaux2.Open "update tb_temp_cuadre_oracle_sid set floa_tem_importe_disponible_sid = " + CStr(rsaux1!importe_seguro + rsaux1!importe_paqueteria + rsaux1!floa_paq_costo_referencia + rsaux1!importe_pedido) + " where inte_tem_numero = " + CStr(rs!inte_tem_numero), cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux1.Close
                  rs.MoveNext
            Wend
            rs.Close
            
            
            
            rs.Open "select inte_car_numero, floa_Car_importe_neto from tb_encabezado_cartera where vcha_Cli_clave_id = '" + var_clave_cliente + "' and vcha_Car_documento = 'fa' and (char_Car_estatus <> 'C' or char_car_Estatus is null)"
            While Not rs.EOF
                  rsaux.Open "select * from tb_Temp_cuadre_oracle_sid where inte_tem_numero = " + CStr(rs!inte_Car_numero), cnn, adOpenDynamic, adLockOptimistic
                  If rsaux.EOF Then
                     rsaux1.Open "insert into tb_temp_cuadre_oracle_sid (inte_tem_numero, floa_tem_importe_sid) values (" + CStr(rs!inte_Car_numero) + "," + CStr(rs!FLOA_CAR_IMPORTE_NETO) + ")", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux.Close
                  rs.MoveNext
            Wend
            rs.Close
            Else
               If rs.State = 1 Then
                  rs.Close
               End If
               rs.Open "DELETE FROM TB_TEMP_FACTURAS_SID_ORACLE", cnn, adOpenDynamic, adLockOptimistic
               rs.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_REFERENCIA = '" + Me.txt_referencia + "'", cnn, adOpenDynamic, adLockOptimistic
            
               var_cadena_clientes = ""
               While Not rs.EOF
                     If var_cadena_clientes = "" Then
                        var_cadena_clientes = "'" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'"
                     Else
                        var_cadena_clientes = var_cadena_clientes + ",'" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'"
                     End If
                     rs.MoveNext
               Wend
               var_cadena_clientes = "(" + var_cadena_clientes + ")"
               rs.Close
               'MsgBox "insert into TB_TEMP_FACTURAS_SID_ORACLE (inte_car_numero, floa_car_importe_neto, floa_tem_importe_oracle) select inte_car_numero, floa_car_importe_neto, 0 from tb_encabezado_Cartera where vcha_ser_Serie_id = 'FT' and vcha_Car_documento = 'FA' and vcha_Cli_clave_id in " + var_cadena_clientes + " and (char_Car_estatus <> 'C' or char_Car_Estatus is null)"
               rs.Open "insert into TB_TEMP_FACTURAS_SID_ORACLE (inte_car_numero, DTIM_cAR_FECHA, floa_car_importe_neto, floa_tem_importe_oracle, VCHA_cLI_REFERENCIA) select inte_car_numero, DTIM_cAR_FECHA, floa_car_importe_neto, 0, VCHA_CLI_REFERENCIA from tb_encabezado_Cartera, TB_CLIENTES where vcha_ser_Serie_id = 'FT' and vcha_Car_documento = 'FA' and TB_ENCABEZADO_CARTERA.vcha_Cli_clave_id in " + var_cadena_clientes + " and (char_Car_estatus <> 'C' or char_Car_Estatus is null) AND TB_ENCABEZADO_cARTERA.VCHA_CLI_CLAVE_ID = TB_CLIENTES.VCHA_CLI_CLAVE_ID", cnn, adOpenDynamic, adLockOptimistic
               rs.Open "select * from TB_TEMP_FACTURAS_SID_ORACLE", cnn, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     Me.txt_referencia = rs!VCHA_CLI_REFERENCIA
                     rsaux.Open "select  vcha_car_tipo_documento, vcha_Car_num_docum, date_car_fecha_cargo, numb_car_importe, numb_car_importe_disponible from tb_Cargo where vcha_car_referencia = '" + Me.txt_referencia + "' AND vcha_car_tipo_documento = 'VA' and vcha_Car_num_docum = " + CStr(rs!inte_Car_numero), cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        var_importe_factura = 0
                        While Not rsaux.EOF
                              var_importe_factura = var_importe_factura + rsaux!numb_car_importe
                              rsaux.MoveNext
                        Wend
                        rsaux1.Open "update  TB_TEMP_FACTURAS_SID_ORACLE set floa_tem_importe_oracle = " + CStr(var_importe_factura) + " where inte_Car_numero = " + CStr(rs!inte_Car_numero), cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux.Close
                     rs.MoveNext
               Wend
               rs.Close
            End If
         End If
      End If
   End If
End Sub
