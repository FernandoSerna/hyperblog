VERSION 5.00
Begin VB.Form frmejecutar_ajustes_inventario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ejecutar ajustes de inventario"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_salida 
      Height          =   360
      Left            =   1740
      TabIndex        =   4
      Top             =   1590
      Width           =   2220
   End
   Begin VB.TextBox txt_entrada 
      Height          =   360
      Left            =   1740
      TabIndex        =   2
      Top             =   1170
      Width           =   2220
   End
   Begin VB.CommandButton cmd_ejecutar_ajustes 
      Caption         =   "Ejecutar ajustes de inventario"
      Height          =   855
      Left            =   210
      TabIndex        =   0
      Top             =   225
      Width           =   4080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Salida:"
      Height          =   195
      Left            =   1005
      TabIndex        =   3
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Entrada:"
      Height          =   195
      Left            =   1005
      TabIndex        =   1
      Top             =   1230
      Width           =   600
   End
End
Attribute VB_Name = "frmejecutar_ajustes_inventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ejecutar_ajustes_Click()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_ENCABEZADO_MOVIMIENTOS_I = New TB_ENCABEZADO_MOVIMIENTOS_I
   Dim var_numero_folio As Double
   Dim var_almacen_Destino As String
   Dim var_clave_movimiento As String
   Dim txt_referencia As String
   Dim txt_codigo As String
   Dim var_año As Integer
   var_si = MsgBox("¿Desea ejecutar los ajustes del inventario?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox("Confirmar el calculo de ajustes del inventario", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rs.Open "SP_EJECUTA_CALCULOS_AJUSTES_INVENTARIO", cnn, adOpenDynamic, adLockOptimistic
         rs.Open "select * from tb_inventario_final where floa_inf_diferencia > 0", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_almacen_Destino = rs!vcha_alm_almacen_id
            var_clave_movimiento = "EA"
            txt_referencia = "AJUSTE DE INVENTARIO " + CStr(Date)
            var_clave_moneda = 1
            var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, CDbl(var_numero_folio), 0, "", "", "", var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", txt_referencia, "", "B", "", "", 0, 0, 0, "1", 1)
            var_numero_folio = var_numero_folio_regreso
            While Not rs.EOF
                  txt_codigo = rs!VCHA_ART_ARTICULO_ID
                  var_cantidad_leida = IIf(IsNull(rs!FLOA_INF_DIFERENCIA), 0, rs!FLOA_INF_DIFERENCIA)
                  var_año = 2005
                  var_costo = IIf(IsNull(rs!floa_inf_costo), 0, rs!floa_inf_costo)
                  var_precio = IIf(IsNull(rs!floa_inf_precio), 0, rs!floa_inf_precio)
                  var_inserta = False
                  var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
                  rs.MoveNext
            Wend
            rs.Close
            cnn.BeginTrans
            Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  var_inserta = False
                  rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO) values ('" + rs!vcha_emp_empresa_id + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_ent_numero) + ", '" + rs!VCHA_ART_ARTICULO_ID + "', " + CStr(rs!floa_ent_cantidad) + ", " + CStr(rs!floa_ent_costo) + " , " + CStr(rs!floa_ent_precio) + ", " + CStr(2005) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            var_estatus_movimiento = "I"
            var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
            var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
            cnn.CommitTrans
            Me.txt_entrada = var_numero_folio
         Else
            rs.Close
         End If
         rs.Open "select * from tb_inventario_final where floa_inf_diferencia < 0", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_almacen_Destino = rs!vcha_alm_almacen_id
            var_clave_movimiento = "SA"
            txt_referencia = "AJUSTE DE INVENTARIO " + CStr(Date)
            var_clave_moneda = 1
            
            var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", "", var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", txt_referencia, "", "B", "", "", 0, 0, 0, "1", 0)
            var_numero_folio = var_numero_folio_regreso
            While Not rs.EOF
                  txt_codigo = rs!VCHA_ART_ARTICULO_ID
                  var_cantidad_leida = (IIf(IsNull(rs!FLOA_INF_DIFERENCIA), 0, rs!FLOA_INF_DIFERENCIA)) * -1
                  var_año = 2005
                  var_costo = IIf(IsNull(rs!floa_inf_costo), 0, rs!floa_inf_costo)
                  var_precio = IIf(IsNull(rs!floa_inf_precio), 0, rs!floa_inf_precio)
                  rsaux.Open "INSERT INTO TB_TEMPORAL_SALIDAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ")", cnn, adOpenDynamic, adLockOptimistic
                  'rsaux.Open "INSERT INTO SALIDAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            rs.Open "EXEC SP_INSERTA_MOVIMIENTOS_SALIDA '" + var_empresa + "','" + var_unidad_organizacional + "', '" + var_almacen_Destino + "','" + var_clave_movimiento + "'," + Str(var_numero_folio) + ",1", cnn, adOpenDynamic, adLockOptimistic
         Else
            rs.Close
         End If
         Me.txt_salida = var_numero_folio
         MsgBox "Se a terminado de ejecutar los ajustes", vbOKOnly, "ATENCION"
      End If
   End If
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_articulos2)
End Sub
