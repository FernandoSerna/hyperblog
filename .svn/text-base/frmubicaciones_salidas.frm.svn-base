VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmubicaciones_salidas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ubicación salidas"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Datos del movimiento "
      Height          =   1275
      Left            =   30
      TabIndex        =   12
      Top             =   240
      Width           =   7965
      Begin VB.TextBox txt_movimiento 
         Height          =   360
         Left            =   1095
         TabIndex        =   16
         Top             =   300
         Width           =   795
      End
      Begin VB.TextBox txt_nombre_movimiento 
         Height          =   360
         Left            =   1920
         TabIndex        =   15
         Top             =   300
         Width           =   5565
      End
      Begin VB.TextBox txt_numero 
         Height          =   360
         Left            =   1095
         TabIndex        =   14
         Top             =   690
         Width           =   1305
      End
      Begin VB.TextBox txt_fecha 
         Height          =   360
         Left            =   3405
         TabIndex        =   13
         Top             =   690
         Width           =   2025
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento:"
         Height          =   195
         Left            =   210
         TabIndex        =   19
         Top             =   383
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   225
         TabIndex        =   18
         Top             =   773
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   2760
         TabIndex        =   17
         Top             =   780
         Width           =   495
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   4155
      TabIndex        =   7
      Top             =   1725
      Width           =   3720
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   8
         Top             =   480
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   3228
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
            Text            =   "UBICACION"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CANTIDAD"
            Object.Width           =   2469
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         Caption         =   " UBICACIONES"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   9
         Top             =   120
         Width           =   3645
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Ubicación "
      Height          =   1020
      Left            =   30
      TabIndex        =   5
      Top             =   5295
      Width           =   7965
      Begin VB.TextBox txt_ubicacion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2325
         TabIndex        =   6
         Top             =   300
         Width           =   3255
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Articulo "
      Height          =   825
      Left            =   75
      TabIndex        =   0
      Top             =   6390
      Width           =   7905
      Begin VB.TextBox txt_codigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   825
         TabIndex        =   2
         Top             =   195
         Width           =   2505
      End
      Begin VB.TextBox txt_cantidad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4770
         TabIndex        =   1
         Top             =   165
         Width           =   2505
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   345
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   4020
         TabIndex        =   3
         Top             =   315
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Detalle del movimiento "
      Height          =   3690
      Left            =   30
      TabIndex        =   10
      Top             =   1560
      Width           =   7965
      Begin MSComctlLib.ListView lv_detalle_movimiento 
         Height          =   3375
         Left            =   45
         TabIndex        =   11
         Top             =   240
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   5953
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODIGO"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPCION"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "CANTIDAD"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "UBICADOS"
            Object.Width           =   1764
         EndProperty
      End
   End
End
Attribute VB_Name = "frmubicaciones_salidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Top = 0
   Left = 2000
   Me.frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_articulos2)
End Sub

Private Sub lv_detalle_movimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      Me.lv_lista.ListItems.Clear
      rs.Open "SELECT * FROM TB_UBICACIONES_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_SAL_NUMERO = " + Me.txt_numero + " AND VCHA_ART_ARTICULO_ID = '" + Me.lv_detalle_movimiento.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , IIf(IsNull(rs!VCHA_UBI_UBICACION), "", rs!VCHA_UBI_UBICACION))
            list_item.SubItems(1) = IIf(IsNull(rs!FLOA_SAL_CANTIDAD), "", rs!FLOA_SAL_CANTIDAD)
            rs.MoveNext:
      Wend
      rs.Close
      Me.frm_lista.Visible = True
      Me.lv_lista.SetFocus
   End If
End Sub

Private Sub lv_detalle_movimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_ubicacion.SetFocus
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         var_si = MsgBox("Desea eliminar la ubicación", vbYesNo, "ATENCION")
         If var_si = 6 Then
            Me.lv_detalle_movimiento.selectedItem.SubItems(3) = CDbl(Me.lv_detalle_movimiento.selectedItem.SubItems(3)) - CDbl(Me.lv_lista.selectedItem.SubItems(1))
            rs.Open "DELETE FROM TB_UBICACIONES_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_SAL_NUMERO = " + Me.txt_numero + " AND VCHA_aRT_aRTICULO_ID = '" + Trim(Me.lv_detalle_movimiento.selectedItem) + "' AND VCHA_UBI_UBICACION = '" + Trim(Me.lv_lista.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
            MsgBox "Se a eliminado la ubicación", vbOKOnly, "ATENCION"
            Me.lv_detalle_movimiento.SetFocus
         Else
            Me.lv_detalle_movimiento.SetFocus
         End If
      End If
   End If
   If KeyAscii = 27 Then
      Me.lv_detalle_movimiento.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_cantidad_GotFocus()
   Me.txt_cantidad = ""
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_movimiento <> "" Then
         If IsNumeric(Me.txt_numero) Then
            If IsNumeric(Me.txt_cantidad) Then
               rs.Open "SELECT * FROM TB_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_SAL_NUMERO = " + Me.txt_numero + " AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_cantidad = IIf(IsNull(rs!FLOA_SAL_CANTIDAD), 0, rs!FLOA_SAL_CANTIDAD)
                  VAR_CANTIDAD_UBICADA = IIf(IsNull(rs!FLOA_SAL_CANTIDAD_UBICADA), 0, rs!FLOA_SAL_CANTIDAD_UBICADA)
                  var_cantidad_leida = CDbl(Me.txt_cantidad)
                  If var_cantidad >= VAR_CANTIDAD_UBICADA + var_cantidad_leida Then
                     valor = Trim(txt_codigo)
                     Set itmfound = Me.lv_detalle_movimiento.findItem(valor, lvwText, , lvwPartial)
                     itmfound.EnsureVisible
                     itmfound.Selected = True
                     Me.lv_detalle_movimiento.selectedItem.SubItems(3) = Me.lv_detalle_movimiento.selectedItem.SubItems(3) + var_cantidad_leida
                     rsaux.Open "SELECT * FROM TB_UBICACIONES_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_SAL_NUMERO = " + Me.txt_numero + " AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "' AND VCHA_UBI_UBICACION = '" + Me.txt_ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        rsaux1.Open "UPDATE TB_UBICACIONES_SALIDAS SET FLOA_SAL_CANTIDAD = ISNULL(FLOA_SAL_CANTIDAD,0) + " + Me.txt_cantidad + " WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_SAL_NUMERO = " + Me.txt_numero + " AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "' AND VCHA_UBI_UBICACION = '" + Me.txt_ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                     Else
                        rsaux1.Open "INSERT INTO TB_UBICACIONES_SALIDAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, VCHA_UBI_UBICACION, FLOA_SAL_CANTIDAD) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + Me.txt_movimiento + "', " + Me.txt_numero + ",'" + Me.txt_codigo + "','" + Me.txt_ubicacion + "'," + Me.txt_cantidad + ")", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux.Close
                     Me.txt_ubicacion.SetFocus
                  Else
                     MsgBox "La cantidad excede a la cantidad en el movimiento", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "El artículo no esta contenido en el movimiento", vbOKOnly, "ATENCION"
                  Me.txt_codigo.SetFocus
               End If
               rs.Close
            Else
               MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se a seleccionado un número de movimiento", vbOKOnly, "ATENCION"
            Me.txt_numero.SetFocus
         End If
      Else
         MsgBox "No se a seleccionado un movimiento", vbOKOnly, "ATENCION"
         Me.txt_movimiento.SetFocus
      End If
   End If
End Sub

Private Sub txt_codigo_GotFocus()
   Me.txt_codigo = ""
   Me.txt_cantidad = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_cantidad.SetFocus
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_ubicacion.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_movimiento_Change()
   Me.lv_detalle_movimiento.ListItems.Clear
   Me.txt_nombre_movimiento = ""
   Me.txt_fecha = ""
   Me.txt_ubicacion = ""
   Me.txt_codigo = ""
   Me.txt_numero = ""
   Me.txt_fecha = ""
   Me.txt_ubicacion = ""
   Me.txt_cantidad = ""
   Me.txt_codigo = ""
End Sub

Private Sub txt_movimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre_movimiento.SetFocus
   End If
End Sub

Private Sub txt_movimiento_LostFocus()
   If Me.txt_movimiento <> "" Then
      rs.Open "select * from tb_movimientos where vcha_mov_movimiento_id = '" + Me.txt_movimiento + "'"
      If Not rs.EOF Then
         Me.lv_detalle_movimiento.ListItems.Clear
         Me.txt_nombre_movimiento = ""
         Me.txt_numero = ""
         Me.txt_fecha = ""
         Me.txt_ubicacion = ""
         Me.txt_codigo = ""
         Me.txt_cantidad = ""
         Me.txt_nombre_movimiento = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
      Else
         Me.lv_detalle_movimiento.ListItems.Clear
         Me.txt_nombre_movimiento = ""
         Me.txt_numero = ""
         Me.txt_fecha = ""
         Me.txt_ubicacion = ""
         Me.txt_cantidad = ""
         Me.txt_codigo = ""
      End If
      rs.Close
   Else
      Me.lv_detalle_movimiento.ListItems.Clear
      Me.txt_nombre_movimiento = ""
      Me.txt_numero = ""
      Me.txt_fecha = ""
      Me.txt_ubicacion = ""
      Me.txt_cantidad = ""
      Me.txt_codigo = ""
   End If
End Sub

Private Sub txt_nombre_movimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_numero.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_numero_Change()
   Me.lv_detalle_movimiento.ListItems.Clear
   Me.txt_fecha = ""
   Me.txt_cantidad = ""
   Me.txt_ubicacion = ""
   Me.txt_codigo = ""
   Me.txt_fecha = ""
   Me.txt_ubicacion = ""
   Me.txt_codigo = ""
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_fecha.SetFocus
   End If
End Sub

Private Sub txt_numero_LostFocus()
   If Me.txt_movimiento <> "" Then
      If IsNumeric(Me.txt_numero) Then
         Me.lv_detalle_movimiento.ListItems.Clear
         Me.txt_fecha = ""
         Me.txt_ubicacion = ""
         Me.txt_codigo = ""
         Me.txt_fecha = ""
         Me.txt_ubicacion = ""
         Me.txt_codigo = ""
         Me.txt_cantidad = ""
         var_cadena = "SELECT dbo.tb_salidas.VCHA_EMP_EMPRESA_ID, dbo.tb_Salidas.VCHA_UOR_UNIDAD_ID, dbo.tb_Salidas.VCHA_MOV_MOVIMIENTO_ID, dbo.tb_salidas.DTIM_SAL_FECHA, dbo.tb_salidas.INTE_SAL_NUMERO, dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, ISNULL(dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD_UBICADA, 0) As FLOA_SAL_CANTIDAD_UBICADA FROM dbo.TB_SALIDAS INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_SALIDAS.INTE_SAL_NUMERO = " + Me.txt_numero + ") AND (dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_sALIDAS.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') AND (dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "')"
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_fecha = rs!dtim_SAL_Fecha
            While Not rs.EOF
                  Set list_item = lv_detalle_movimiento.ListItems.Add(, , IIf(IsNull(rs!VCHA_aRT_ARTICULO_ID), "", rs!VCHA_aRT_ARTICULO_ID))
                  list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ART_NOMBRE_ESPAÑOL), "", rs!VCHA_ART_NOMBRE_ESPAÑOL)
                  list_item.SubItems(2) = IIf(IsNull(rs!FLOA_SAL_CANTIDAD), 0, rs!FLOA_SAL_CANTIDAD)
                  list_item.SubItems(3) = IIf(IsNull(rs!FLOA_SAL_CANTIDAD_UBICADA), 0, rs!FLOA_SAL_CANTIDAD_UBICADA)
                  rs.MoveNext:
            Wend
         Else
            MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
            Me.lv_detalle_movimiento.ListItems.Clear
            Me.txt_fecha = ""
            Me.txt_ubicacion = ""
            Me.txt_codigo = ""
            Me.txt_fecha = ""
            Me.txt_ubicacion = ""
            Me.txt_codigo = ""
            Me.txt_cantidad = ""
         End If
         rs.Close
      Else
         MsgBox "Número de movimiento incorrecto", vbOKOnly, "ATENCION"
         Me.lv_detalle_movimiento.ListItems.Clear
         Me.txt_fecha = ""
         Me.txt_ubicacion = ""
         Me.txt_codigo = ""
         Me.txt_fecha = ""
         Me.txt_ubicacion = ""
         Me.txt_cantidad = ""
         Me.txt_codigo = ""
      End If
   Else
      MsgBox "No se a seleccionado un movimiento", vbOKOnly, "ATENCION"
      Me.lv_detalle_movimiento.ListItems.Clear
      Me.txt_nombre_movimiento = ""
      Me.txt_fecha = ""
      Me.txt_ubicacion = ""
      Me.txt_codigo = ""
      Me.txt_numero = ""
      Me.txt_fecha = ""
      Me.txt_ubicacion = ""
      Me.txt_cantidad = ""
      Me.txt_codigo = ""
   End If
End Sub

Private Sub txt_ubicacion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_codigo.SetFocus
   End If
End Sub

