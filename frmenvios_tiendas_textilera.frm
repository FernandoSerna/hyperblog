VERSION 5.00
Begin VB.Form frmenvios_tiendas_textilera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envios tiendas desde la textilera"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5610
      Picture         =   "frmenvios_tiendas_textilera.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmenvios_tiendas_textilera.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   30
      TabIndex        =   6
      Top             =   270
      Width           =   5970
   End
   Begin VB.Frame Frame1 
      Caption         =   " Salida de la textilera "
      Height          =   1080
      Left            =   135
      TabIndex        =   0
      Top             =   465
      Width           =   5760
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2115
         TabIndex        =   3
         Top             =   630
         Width           =   3495
      End
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   855
         TabIndex        =   2
         Top             =   630
         Width           =   1230
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   855
         TabIndex        =   1
         Top             =   285
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   705
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   315
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmenvios_tiendas_textilera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()

Dim var_año As Integer
Dim var_almacen_Destino As String
Dim var_primera_vez As Boolean
Dim var_numero_folio As Double
Dim var_cantidad_leida As Double
Dim var_costo As Double
Dim var_precio As Double
Dim var_descripcion_articulo As String
Dim var_estatus_movimiento As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_numero_causa As Integer
Dim var_elimina As Boolean
Dim var_ventana As Integer
Dim var_clave_moneda As String
Dim var_renglon As Double
Dim txt_codigo As Variant


   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_I = New TB_ENTRADAS_I
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M


   If IsNumeric(Me.txt_numero) Then
      rs.Open "select * from tb_archivos_envios where inte_aco_numero = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
               var_equivalencia = (rs!vcha_aco_codigo_externo)
               rsaux1.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + var_equivalencia + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  rsaux3.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + IIf(IsNull(rsaux1!vcha_Art_articulo_id), "", rsaux1!vcha_Art_articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     rsaux2.Open "update tb_archivos_envios set vcha_Art_Articulo_id = '" + IIf(IsNull(rsaux1!vcha_Art_articulo_id), "", rsaux1!vcha_Art_articulo_id) + "' where inte_aco_consecutivo = " + CStr(rs!inte_aco_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux2.Open "update tb_archivos_envios set vcha_Art_Articulo_id = '' where inte_aco_consecutivo = " + CStr(rs!inte_aco_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux3.Close
               Else
                  rsaux2.Open "update tb_archivos_envios set vcha_Art_Articulo_id = '' where inte_aco_consecutivo = " + CStr(rs!inte_aco_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux1.Close
               rs.MoveNext
         Wend
         rsaux.Open "select * from tb_archivos_envios where inte_Aco_numero = " + Me.txt_numero + " and vcha_Art_Articulo_id = ''", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_posible = False
         Else
            var_posible = True
         End If
         rsaux.Close
         If var_posible = True Then
'''' SE HACE LA ENTRADA
            Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
            Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
            Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
            Dim var_inserta As Boolean
            bandera_suma = False
            var_inserta = False
            var_almacen_Destino = "8"
            var_clave_movimiento = "EA"

            var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, "8", "EA", Now, 0, 0, "", "", "", "8", "", var_clave_usuario_global, fun_NombrePc, 0, "", "Entrada por envio a tienda desde la textilera " + Me.txt_numero, "", "B", "", "", 0, 0, 0, "1", 0)
            var_numero_folio = var_numero_folio_regreso
            rs.Close
            rs.Open "select * from tb_archivos_envios where inte_aco_numero = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
            rs.MoveFirst
            While Not rs.EOF
                  txt_codigo = IIf(IsNull(rs!vcha_Art_articulo_id), "", rs!vcha_Art_articulo_id)
                  var_cantidad_leida = IIf(IsNull(rs!floa_Aco_Cantidad), 0, rs!floa_Aco_Cantidad)
                  Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
                  rsaux.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_inserta = False
                     MsgBox Cadena
                     var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, 2005)
                     rsaux.Close
                  Else
                     var_inserta = False
                     MsgBox Cadena
                     rsaux10.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux10.EOF Then
                        var_precio = IIf(IsNull(rsaux10!mone_Art_precio_base), 0, rsaux10!mone_Art_precio_base)
                     Else
                        var_precio = 0
                     End If
                     rsaux10.Close
                     var_costo = IIf(IsNull(rs!floa_aco_costo), 0, rs!floa_aco_costo)
                     var_cadena = "insert into tb_temporal_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, inte_ent_año) values ('" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_Destino + "','" + var_clave_movimiento + "'," + CStr(var_numero_folio) + ",'" + txt_codigo + "'," + CStr(var_cantidad_leida) + "," + CStr(var_costo) + "," + CStr(var_precio) + ",2005)"
                     MsgBox "insert into tb_temporal_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, inte_ent_año) values ('" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_Destino + "','" + var_clave_movimiento + "'," + CStr(var_numero_folio) + ",'" + txt_codigo + "'," + CStr(var_cantidad_leida) + "," + CStr(var_costo) + "," + CStr(var_precio) + ",2005)"
                     rsaux8.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     'var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", "", 2005)
                     rsaux.Close
                  End If
                  rs.MoveNext
            Wend
            
            
            Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  var_inserta = False
                  If rsaux.State = 1 Then
                     rsaux.Close
                  End If
                  rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO) values ('" + rs!vcha_emp_empresa_id + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_ent_numero) + ", '" + rs!vcha_Art_articulo_id + "', " + CStr(rs!floa_ent_cantidad) + ", " + CStr(rs!floa_ent_costo) + " , " + CStr(rs!floa_ent_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            var_estatus_movimiento = "I"
            var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
            var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)

'''' FIN DE LA ENTRADA
         Else
            MsgBox "Existen artículos sin equivalencia en el almacen general", vbOKOnly, "ATENCION"
         End If
      Else
         rs.Close
         MsgBox "Número de salida no existe", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de salida incorrecto", vbOKOnly, "ATENCION"
   End If
   If rs.State = 1 Then
      rs.Close
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    activa_forma (var_activa_forma_packing_list)
End Sub

