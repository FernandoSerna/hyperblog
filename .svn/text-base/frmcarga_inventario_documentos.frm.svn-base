VERSION 5.00
Begin VB.Form frmcarga_inventario_documentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cargar inventario de documentos"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_cargar_inventario 
      Caption         =   "Cargar inventario de documentos"
      Height          =   810
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   3705
   End
End
Attribute VB_Name = "frmcarga_inventario_documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cargar_inventario_Click()


        rs.Open "SELECT * FROM TB_INVENTARIO_DOCUMENTOS WHERE VCHA_EMP_EMPRESA_ID = '18' AND DTIM_IDO_FECHA_ENTRAGA >= {d '2009-03-15'} ORDER BY DTIM_IDO_FECHA_ENTRAGA DESC", cnn_sqlquezada2, adOpenDynamic, adLockOptimistic
        While Not rs.EOF
              'rsaux1.Open "SELECT * FROM TB_INVENTARIO_DOCUMENTOS WHERE VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' and inte_emb_embarque = " + CStr(rs!inte_emb_embarque) + " and vcha_age_agente_id = '" + rs!vcha_age_Agente_id + "' and vcha_Car_tipo_documento = '" + rs!vcha_Car_tipo_documento + "' and vcha_car_documento = '" + rs!vcha_Car_documento + "' and vcha_Car_clase_id = '" + rs!vcha_Car_clase_id + "' and inte_car_numero = " + CStr(rs!INTE_cAR_NUMERO), cnn, adOpenDynamic, adLockOptimistic
              'MsgBox "SELECT * FROM TB_INVENTARIO_DOCUMENTOS WHERE VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' and vcha_car_documento = '" + rs!vcha_Car_documento + "' and inte_car_numero = " + CStr(rs!INTE_cAR_NUMERO) + " and vcha_Ser_Serie_id = '" + rs!vcha_Ser_Serie_id + "'"
              rsaux1.Open "SELECT * FROM TB_INVENTARIO_DOCUMENTOS WHERE VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' and vcha_car_documento = '" + rs!vcha_Car_documento + "' and inte_car_numero = " + CStr(rs!inte_car_numero) + " and vcha_Ser_Serie_id = '" + rs!vcha_ser_Serie_id + "'", cnn, adOpenDynamic, adLockOptimistic
              If rsaux1.EOF Then
                 var_dia = CStr(Day(rs!dtim_ido_fecha_entraga))
                 var_mes = CStr(Month(rs!dtim_ido_fecha_entraga))
                 var_año = CStr(Year(rs!dtim_ido_fecha_entraga))
                 If Len(Trim(var_dia)) = 1 Then
                    var_dia = "0" + var_dia
                 End If
                 If Len(Trim(var_mes)) = 1 Then
                    var_mes = "0" + var_mes
                 End If
                 dtim_ido_fecha_entraga = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"

                 var_dia = CStr(Day(rs!dtim_ido_fecha_insercion))
                 var_mes = CStr(Month(rs!dtim_ido_fecha_insercion))
                 var_año = CStr(Year(rs!dtim_ido_fecha_insercion))
                 If Len(Trim(var_dia)) = 1 Then
                    var_dia = "0" + var_dia
                 End If
                 If Len(Trim(var_mes)) = 1 Then
                    var_mes = "0" + var_mes
                 End If
                 dtim_ido_fecha_insercio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"


                 var_cadena = "insert into tb_inventario_documentos (vcha_emp_empresa_id, inte_emb_Embarque, vcha_age_Agente_id, vcha_Car_tipo_documento, vcha_Car_documento, vcha_Car_clase_id, inte_car_numero,char_car_afectacion, vcha_Ser_serie_id, char_ido_estatus, floa_ido_cantidad, floa_Car_importe_neto, floa_car_tipo_cambio, vcha_mon_moneda_id, dtim_ido_fecha_entraga, vcha_cli_clave_id, dtim_ido_fecha_insercion, vcha_ido_agente_real) values   ('" + rs!VCHA_EMP_EMPRESA_ID + "', " + CStr(rs!inte_emb_embarque) + ", '" + rs!VCHA_AGE_AGENTE_ID + "', '" + rs!vcha_Car_tipo_documento + "', '" + rs!vcha_Car_documento + "', '" + rs!vcha_Car_clase_id + "', " + CStr(rs!inte_car_numero) + ",'" + rs!char_car_afectacion + "', '" + rs!vcha_ser_Serie_id + "', '" + rs!char_ido_estatus + "', " + CStr(rs!floa_ido_cantidad) + ", " + CStr(rs!floa_Car_importe_neto) + ", " + CStr(rs!floa_car_tipo_cambio) + ", '" + rs!vcha_mon_moneda_id + "',"
                 var_cadena = var_cadena + " " + dtim_ido_fecha_entraga + ", '" + rs!vcha_cli_clave_id + "', " + dtim_ido_fecha_insercio + ", '" + rs!VCHA_AGE_AGENTE_ID + "')"
                    
                 'MsgBox var_cadena
                 
                 
                 rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
              End If
              rsaux1.Close
              rs.MoveNext
        Wend
        rs.Close
        MsgBox "Se a terminado de cargar el inventario de documentos", vbOKOnly, "ATENCION"




End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub
