VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcancelacion_embarque_relectura_cajas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelacion de embarque y relectura de cajas"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Height          =   2700
      Left            =   135
      TabIndex        =   13
      Top             =   2670
      Width           =   7020
      Begin MSComctlLib.ListView lv_diferencias 
         Height          =   2190
         Left            =   30
         TabIndex        =   14
         Top             =   450
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   3863
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2295
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Caja     "
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Salida     "
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Unidad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Almacen"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Movimiento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Numero"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Descuento1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Descuento2"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Lista_precios"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "tipo_cambio"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Diferencias"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   15
         Top             =   120
         Width           =   6945
      End
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6855
      Picture         =   "frmcancelacion_embarque_relectura_cajas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cancelar Factura"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   165
      Picture         =   "frmcancelacion_embarque_relectura_cajas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   120
      TabIndex        =   10
      Top             =   255
      Width           =   7035
   End
   Begin VB.Frame Frame4 
      Height          =   1050
      Left            =   3675
      TabIndex        =   8
      Top             =   1545
      Width           =   3465
      Begin VB.TextBox txt_cantidad_salida 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   615
         TabIndex        =   2
         Top             =   495
         Width           =   2340
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad Salida"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   9
         Top             =   135
         Width           =   3390
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   135
      TabIndex        =   6
      Top             =   1545
      Width           =   3465
      Begin VB.TextBox txt_cantidad_cajas 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   585
         TabIndex        =   1
         Top             =   495
         Width           =   2340
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad Cajas"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   7
         Top             =   135
         Width           =   3390
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   135
      TabIndex        =   4
      Top             =   405
      Width           =   7020
      Begin VB.TextBox txt_embarque 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2490
         TabIndex        =   0
         Top             =   495
         Width           =   1860
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   " Embarque"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   5
         Top             =   135
         Width           =   6945
      End
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3630
      Width           =   165
   End
End
Attribute VB_Name = "frmcancelacion_embarque_relectura_cajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
   If var_empresa = "03" Then
      If IsNumeric(Me.txt_embarque) Then
         If CDbl(Me.txt_cantidad_cajas) <> CDbl(Me.txt_cantidad_salida) Then
            If CDbl(Me.txt_cantidad_cajas) > 0 Then
               If Me.lv_diferencias.ListItems.Count > 0 Then
                  var_si = MsgBox("¿Desea corregir las diferencias?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     For var_j = 1 To lv_diferencias.ListItems.Count
                         lv_diferencias.ListItems.Item(var_j).Selected = True
                         var_unidad_diferencia = lv_diferencias.selectedItem.SubItems(4)
                         var_almacen_diferencia = lv_diferencias.selectedItem.SubItems(5)
                         var_movimiento_diferencia = lv_diferencias.selectedItem.SubItems(6)
                         var_numero_diferencia = CDbl(lv_diferencias.selectedItem.SubItems(7))
                         var_descuento_1_diferencia = CDbl(lv_diferencias.selectedItem.SubItems(8))
                         var_descuento_2_diferencia = CDbl(lv_diferencias.selectedItem.SubItems(9))
                         var_lista_diferencia = lv_diferencias.selectedItem.SubItems(10)
                         var_tipo_cambio_diferencia = CDbl(lv_diferencias.selectedItem.SubItems(11))
                         var_diferencia = CDbl(lv_diferencias.selectedItem.SubItems(2)) - CDbl(Me.lv_diferencias.selectedItem.SubItems(3))
                         rs.Open "SELECT * FROM TB_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_diferencia + "' and vcha_alm_almacen_id = '" + var_almacen_diferencia + "' and vcha_mov_movimiento_id = '" + var_movimiento_diferencia + "' and inte_sal_numero = " + CStr(var_numero_diferencia) + " and vcha_art_articulo_id = '" + Trim(Me.lv_diferencias.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                         If Not rs.EOF Then
                            rsaux.Open "update tb_salidas set floa_sal_cantidad = isnull(floa_sal_Cantidad,0) + " + CStr(var_diferencia) + " where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_diferencia + "' and vcha_alm_almacen_id = '" + var_almacen_diferencia + "' and vcha_mov_movimiento_id = '" + var_movimiento_diferencia + "' and inte_sal_numero = " + CStr(var_numero_diferencia) + " and vcha_art_articulo_id = '" + Trim(Me.lv_diferencias.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                         Else
                            rsaux.Open "select floa_dli_precio from tb_detalle_lista_precios where vcha_lis_lista_precios_id = '" + var_lista_diferencia + "' and vcha_art_articulo_id = '" + lv_diferencias.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                            If Not rsaux.EOF Then
                               var_precio = IIf(IsNull(rsaux!floa_dli_precio), 0, rsaux!floa_dli_precio) * var_tipo_cambio_diferencia
                               rsaux2.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_diferencia + "' and vcha_Art_Articulo_id = '" + lv_diferencias.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                               If Not rsaux2.EOF Then
                                  var_costo = IIf(IsNull(rsaux2!floa_exi_costo_2005), 0, rsaux2!floa_exi_costo_2005)
                               Else
                                  rsaux3.Open "select * from tb_articulos where vcha_art_articulo_id = '" + lv_diferencias.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                                  If Not rsaux3.EOF Then
                                     var_costo = IIf(IsNull(rsaux3!mone_art_costo_estandar), 0, rsaux3!mone_art_costo_estandar)
                                  Else
                                     var_costo = 0
                                  End If
                                  rsaux3.Close
                               End If
                               rsaux2.Close
                               var_cadena = "insert into tb_salidas (vcha_emp_empresa_id,        vcha_uor_unidad_id,             vcha_alm_almacen_id,             vcha_mov_movimiento_id,        inte_sal_numero,                vcha_Art_Articulo_id,                     floa_sal_Cantidad, floa_sal_precio, floa_sal_costo, floa_Sal_descuento, char_ped_tipo, floa_sal_descuento_1, floa_sal_Descuento_2, inte_sal_año,floa_sal_promocion_1, floa_sal_promocion_2)"
                               var_cadena = var_cadena + " values    ('" + var_empresa + "','" + var_unidad_diferencia + "','" + var_almacen_diferencia + "','" + var_movimiento_diferencia + "', " + CStr(var_numero_diferencia) + ",'" + lv_diferencias.selectedItem + "'," + CStr(var_diferencia) + "," + CStr(var_precio) + "," + CStr(var_costo) + ",0,'E'," + CStr(var_descuento_1_diferencia) + "," + CStr(var_descuento_2_diferencia) + ",2005,0,0)"
                               rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                            Else
                               MsgBox "El articulo " + lv_diferencias.selectedItem.SubItems(1) + " no se encuentra en la lista de precios", vbOKOnly, "ATENCION"
                            End If
                            rsaux.Close
                         End If
                         rs.Close
                     Next var_j
                     MsgBox "Se a terminado de corregir el embarque", vbOKOnly, "ATENCION"
                     Me.txt_embarque = ""
                     Me.txt_cantidad_cajas = ""
                     Me.txt_cantidad_salida = ""
                     Me.lv_diferencias.ListItems.Clear
                  End If
               Else
                  MsgBox "No existen diferencias", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El embarque " + Me.txt_embarque + " no fue empaquetado", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No existen diferencias", vbOKOnly, "ATNECION"
         End If
      Else
         MsgBox "No se a indicado un embarque", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "MODULO NO VALIDO PARA ESTA EMPRESA", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 1000
   Left = 2000
   If var_empresa <> "03" Then
      MsgBox "MODULO NO VALIDO PARA ESTA EMPRESA", vbOKOnly, "ATENCION"
      Me.txt_cantidad_cajas.Enabled = False
      Me.txt_cantidad_salida.Enabled = False
      Me.txt_embarque.Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_cantidad_cajas_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_cantidad_salida_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_embarque_LostFocus()
   If Trim(Me.txt_embarque) <> "" Then
      Me.txt_cantidad_cajas = "0.00"
      Me.txt_cantidad_salida = "0.00"
      Me.lv_diferencias.ListItems.Clear
      If IsNumeric(Me.txt_embarque) Then
         Me.txt_cantidad_cajas = "0.00"
         Me.txt_cantidad_salida = "0.00"
         Me.lv_diferencias.ListItems.Clear
         rs.Open "SELECT * FROM TB_DETALLE_EMBARQUES WHERE INTE_eMB_EMBARQUE = " + Me.txt_embarque + " AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_cantidad_Salida = 0
            While Not rs.EOF
                  rsaux.Open "SELECT SUM(FLOA_sAL_cANTIDAD) FROM TB_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + rs!vcha_uor_unidad_id + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + rs!vcha_mov_movimiento_id + "' AND INTE_SAL_NUMERO = " + CStr(rs!inte_sal_numero), cnn, adOpenDynamic, adLockOptimistic
                  var_cantidad_Salida = var_cantidad_Salida + rsaux(0).Value
                  rsaux.Close
                  rs.MoveNext
            Wend
            Me.txt_cantidad_salida = Format(var_cantidad_Salida, "###,###,##0.00")
            rsaux.Open "select sum(floa_paq_cantidad) from tb_Detalle_cajas where vcha_emp_empresa_id = '" + var_empresa + "' and inte_Emb_embarque = " + Me.txt_embarque + " and char_paq_estatus = 'S'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_cantidad_cajas = Format(IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value), "###,###,##0.00")
            Else
               Me.txt_cantidad_cajas = "0.00"
            End If
            rsaux.Close
            rs.Close
            cnn.BeginTrans
            rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_CORRECCION_EMBARQUES_DIFERENCIAS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rs.Close
            rs.Open "INSERT INTO TB_TEMP_CORRECCION_EMBARQUES_DIFERENCIAS (INTE_TEM_CONSECUTIVO, INTE_EMB_EMBARQUE) VALUES (" + CStr(var_consecutivo) + "," + Me.txt_embarque + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_cadena = "SELECT VCHA_ART_ARTICULO_ID, SUM(FLOA_PAQ_CANTIDAD) AS FLOA_PAQ_CANTIDAD From dbo.TB_DETALLE_CAJAS WHERE (VCHA_EMP_EMPRESA_ID = '03') AND (INTE_EMB_EMBARQUE = " + Me.txt_embarque + ") AND (CHAR_PAQ_ESTATUS <> 'C') GROUP BY VCHA_ART_ARTICULO_ID"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux.Open "INSERT INTO TB_TEMP_CORRECCION_EMBARQUES_DIFERENCIAS (INTE_TEM_CONSECUTIVO,INTE_EMB_EMBARQUE, VCHA_ART_aRTICULO_ID, FLOA_PAQ_CANTIDAD,FLOA_SAL_CANTIDAD) VALUES (" + CStr(var_consecutivo) + "," + Me.txt_embarque + ",'" + rs!vcha_Art_articulo_id + "'," + CStr(rs!floa_paq_cantidad) + ",0)", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            rs.Open "select inte_sal_numero from tb_detalle_embarques where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
            var_cadena_numero = "("
            While Not rs.EOF
                  If var_cadena_numero = "(" Then
                     var_cadena_numero = var_cadena_numero + "INTE_SAL_NUMERO = " + CStr(rs!inte_sal_numero)
                  Else
                     var_cadena_numero = var_cadena_numero + " or INTE_SAL_NUMERO = " + CStr(rs!inte_sal_numero)
                  End If
                  rs.MoveNext
            Wend
            rs.Close
            var_cadena_numero = var_cadena_numero + ")"
            var_cadena = "SELECT VCHA_ART_ARTICULO_ID, sum(FLOA_SAL_CANTIDAD) as floa_sal_cantidad From dbo.TB_SALIDAS WHERE  VCHA_MOV_MOVIMIENTO_ID = 'ex' AND " + var_cadena_numero + " group by vcha_art_Articulo_id"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux2.Open "update TB_TEMP_CORRECCION_EMBARQUES_DIFERENCIAS set  floa_sal_Cantidad = floa_sal_Cantidad + " + CStr(rs!floa_sal_cantidad) + " where vcha_Art_Articulo_id = '" + rs!vcha_Art_articulo_id + "' and inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            rs.Open "select * from VW_TEMP_CORRECCION_EMBARQUES_DIFERENCIAS where floa_sal_cantidad <> floa_paq_cantidad and inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Dim list_item As ListItem
               numero_items_ALMACENES = 0
               While Not rs.EOF
                     Set list_item = Me.lv_diferencias.ListItems.Add(, , rs!vcha_Art_articulo_id)
                     list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
                     list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_paq_cantidad), 0, rs!floa_paq_cantidad), "###,###,##0.00")
                     list_item.SubItems(3) = Format(IIf(IsNull(rs!floa_sal_cantidad), 0, rs!floa_sal_cantidad), "###,###,##0.00")
                     list_item.SubItems(4) = IIf(IsNull(rs!vcha_uor_unidad_id), "", rs!vcha_uor_unidad_id)
                     list_item.SubItems(5) = IIf(IsNull(rs!vcha_alm_almacen_id), "", rs!vcha_alm_almacen_id)
                     list_item.SubItems(6) = IIf(IsNull(rs!vcha_mov_movimiento_id), "", rs!vcha_mov_movimiento_id)
                     list_item.SubItems(7) = IIf(IsNull(rs!inte_sal_numero), 0, rs!inte_sal_numero)
                     list_item.SubItems(8) = IIf(IsNull(rs!floa_emo_descuento_1), 0, rs!floa_emo_descuento_1)
                     list_item.SubItems(9) = IIf(IsNull(rs!floa_emo_descuento_2), 0, rs!floa_emo_descuento_2)
                     list_item.SubItems(10) = IIf(IsNull(rs!vcha_lis_lista_id), "", rs!vcha_lis_lista_id)
                     list_item.SubItems(11) = IIf(IsNull(rs!floa_emo_tipo_cambio), "", rs!floa_emo_tipo_cambio)
                     rs.MoveNext
               Wend
            Else
               MsgBox "No existen diferencias", vbOKOnly, "ATENCION"
            End If
            rs.Close
            rs.Open "DELETE FROM TB_TEMP_CORRECCION_EMBARQUES_DIFERENCIAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            rs.Close
            MsgBox "El embarque no existe o no a sido cerrado aun", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub
