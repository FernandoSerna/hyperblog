VERSION 5.00
Begin VB.Form frmoracle_comparacion_pedido_afectado 
   Caption         =   "Form1"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   675
      Left            =   150
      TabIndex        =   0
      Top             =   315
      Width           =   5220
   End
End
Attribute VB_Name = "frmoracle_comparacion_pedido_afectado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   rs.Open "delete from TB_ORACLE_COMPARACION_PEDIDO_AFECTACION", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "SELECT inte_emb_embarque, SOURCE_HEADER_NUMBER, SUM(FLOA_sal_cANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_sALIDAS  where source_header_number > 2500 GROUP BY inte_emb_embarque, SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux2.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(rs!inte_emb_embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux2.EOF Then
            rsaux.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha) VALUES (" + CStr(rs!source_header_number) + "," + CStr(rs!Cantidad) + ",0, '" + CStr(rsaux2!FECHA_INICIO) + "')", cnn, adOpenDynamic, adLockOptimistic
         Else
            rsaux.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha) VALUES (" + CStr(rs!source_header_number) + "," + CStr(rs!Cantidad) + ",0, '')", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux2.Close
         rs.MoveNext
   Wend
   rs.Close
   rs.Open "SELECT inte_emb_embarque, SOURCE_HEADER_NUMBER, SUM(FLOA_sal_cANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_sALIDAS_CAJAS  where source_header_number > 2500  GROUP BY inte_emb_embarque, SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux2.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(rs!inte_emb_embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
         rsaux.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha) VALUES (" + CStr(rs!source_header_number) + "," + CStr(rs!Cantidad) + ",0, '" + CStr(rsaux2!FECHA_INICIO) + "')", cnn, adOpenDynamic, adLockOptimistic
         rsaux2.Close
         rs.MoveNext
   Wend
   rs.Close
   'rs.Open "SELECT pedido FROM TB_ORACLE_COMPARACION_PEDIDO_AFECTACION where pedido is not null", cnn, adOpenDynamic, adLockOptimistic
   rsaux1.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   rs.Open "SELECT SOURCE_HEADER_NUMBER, SUM(SHIPPED_QUANTITY) AS CANTIDAD FROM WSH_DELIVERABLES_V where  source_header_number > 2500 and organization_id = 93 group by SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "select * from TB_ORACLE_COMPARACION_PEDIDO_AFECTACION where pedido = " + CStr(rs!source_header_number), cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux1.Open "UPDATE TB_ORACLE_COMPARACION_PEDIDO_AFECTACION SET CANTIDAD_AFECTADA = " + CStr(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) + " WHERE PEDIDO = " + CStr(rs!source_header_number), cnn, adOpenDynamic, adLockOptimistic
         Else
            rsaux1.Open "insert into TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (pedido, cantidad_leida, cantidad_afectada) values (" + CStr(rs!source_header_number) + ",0," + CStr(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) + ")", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux.Close
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub
