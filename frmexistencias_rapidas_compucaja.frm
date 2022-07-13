VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmexistencias_rapidas_compucaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Existencias"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_disponibles 
      Height          =   3585
      Left            =   1380
      TabIndex        =   5
      Top             =   1470
      Width           =   7110
      Begin VB.TextBox txt_nombre_articulo 
         Height          =   345
         Left            =   90
         TabIndex        =   6
         Top             =   420
         Width           =   6915
      End
      Begin MSComctlLib.ListView lv_disponibles 
         Height          =   2700
         Left            =   75
         TabIndex        =   7
         Top             =   795
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   4763
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
            Text            =   "Código"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre del Artículo"
            Object.Width           =   9701
         EndProperty
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000D&
         Caption         =   " Artículos Disponibles"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   8
         Top             =   120
         Width           =   7035
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Existencias "
      Height          =   4290
      Left            =   180
      TabIndex        =   3
      Top             =   1485
      Width           =   8970
      Begin MSComctlLib.ListView lv_existencias 
         Height          =   3900
         Left            =   120
         TabIndex        =   4
         Top             =   225
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   6879
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sistema"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Almacén"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Artículo "
      Height          =   1290
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   8985
      Begin VB.CommandButton cmd_pedido 
         Height          =   330
         Left            =   8460
         Picture         =   "frmexistencias_rapidas_compucaja.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   810
         Width           =   360
      End
      Begin VB.TextBox txt_precio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1455
         TabIndex        =   10
         Top             =   690
         Width           =   2295
      End
      Begin VB.TextBox txt_descripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   1455
         TabIndex        =   2
         Top             =   240
         Width           =   7410
      End
      Begin VB.TextBox txt_codigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label lbl_estatus_2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3885
         TabIndex        =   13
         Top             =   960
         Width           =   3195
      End
      Begin VB.Label lbl_estatus_1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3885
         TabIndex        =   12
         Top             =   690
         Width           =   3000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Precio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   330
         TabIndex        =   9
         Top             =   750
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmexistencias_rapidas_compucaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   rs.Open "select art_codigo, art_descripcion from VIA_EXISTENCIA_ALMACEN", cnn_compucaja, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux1.Open "update tb_Articulos set vcha_Art_nombre_Español = '" + rs!art_descripcion + "' where vcha_Art_articulo_id = '" + rs!art_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub cmd_pedido_Click()
   frmpedido_tienda_cantia.Show
End Sub

Private Sub Form_Load()
   Top = 600
   Left = 1200
   Me.frm_disponibles.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub lv_disponibles_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 38 And Shift = 1 Then
      Me.txt_nombre_articulo.SetFocus
   End If
End Sub

Private Sub lv_disponibles_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_disponibles.ListItems.Count > 0 Then
         Me.txt_codigo = Me.lv_disponibles.selectedItem
         Me.txt_descripcion = Me.lv_disponibles.selectedItem.SubItems(1)
         Me.txt_codigo.SetFocus
         Me.frm_disponibles.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_disponibles.Visible = False
   End If
End Sub

Private Sub lv_existencias_GotFocus()
   Me.frm_disponibles.Visible = False
End Sub

Private Sub lv_existencias_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_codigo_Change()
   Me.txt_descripcion = ""
   Me.txt_precio = ""
   Me.lbl_estatus_1 = ""
   Me.lbl_estatus_2 = ""
   Me.lv_existencias.ListItems.Clear
End Sub

Private Sub txt_codigo_GotFocus()
   Me.frm_disponibles.Visible = False
End Sub

Private Sub txt_Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.lbl_estatus_1 = ""
      Me.lbl_estatus_2 = ""
      Me.txt_nombre_articulo = ""
      Me.lv_disponibles.ListItems.Clear
      Me.frm_disponibles.Visible = True
      Me.txt_nombre_articulo.SetFocus
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.txt_descripcion.SetFocus
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   If Trim(Me.txt_codigo) <> "" Then
      Me.lbl_estatus_1 = ""
      Me.lbl_estatus_2 = ""
      Me.txt_descripcion = ""
      Me.lv_existencias.ListItems.Clear
      var_codigo = ""
      var_nombre = ""
      rs.Open "select * from tb_Articulos where vcha_art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      'rs.Open "SELECT DISTINCT ART_CODIGO AS vcha_Art_articulo_id, ART_DESCRIPCION AS vcha_art_nombre_Español FROM VIA_EXISTENCIA_ALMACEN WHERE ART_CODIGO = '" + Me.txt_codigo + "'", cnn_compucaja, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_codigo = IIf(IsNull(rs!vcha_Art_Articulo_id), "", rs!vcha_Art_Articulo_id)
         var_nombre = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
         var_estatus = IIf(IsNull(rs!vcha_Car_clase_id), "0", rs!vcha_Car_clase_id)
         
      Else
         rsaux.Open "SELECT * FROM TB_EQUIVALENCIAS WHERE VCHA_EQU_CODIGO_EQUIVALENTE = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux1.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + IIf(IsNull(rsaux!vcha_Art_Articulo_id), "", rsaux!vcha_Art_Articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               var_codigo = IIf(IsNull(rsaux1!vcha_Art_Articulo_id), "", rsaux1!vcha_Art_Articulo_id)
               var_nombre = IIf(IsNull(rsaux1!vcha_Art_nombre_español), "", rsaux1!vcha_Art_nombre_español)
               var_estatus = IIf(IsNull(rsaux1!vcha_Car_clase_id), "", rsaux1!vcha_Car_clase_id)
            Else
               var_codigo = ""
               Me.lbl_estatus_1 = ""
               Me.lbl_estatus_2 = ""
            End If
            rsaux1.Close
         Else
            Me.lbl_estatus_1 = ""
            Me.lbl_estatus_2 = ""
            var_codigo = ""
         End If
         rsaux.Close
      End If
      rs.Close
      If var_codigo <> "" Then
         If var_estatus = "0" Then
            Me.lbl_estatus_1 = "VIGENTE"
            Me.lbl_estatus_2 = ""
         End If
         If var_estatus = "1" Then
            Me.lbl_estatus_1 = "DESCONTINUADO"
            Me.lbl_estatus_2 = ""
         End If
         If var_estatus = "2" Then
            Me.lbl_estatus_1 = "VIGENTE"
            Me.lbl_estatus_2 = "Temporalmente fuera de servicio"
         End If
         Me.txt_descripcion = var_nombre
         rsaux.Open "select * from VIA_EXISTENCIA_ALMACEN where art_codigo = '" + var_codigo + "' or art_gtin = '" + var_codigo + "'", cnn_compucaja, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            Me.txt_precio = Format(IIf(IsNull(rsaux!lpa_precioventaimp), 0, rsaux!lpa_precioventaimp), "###,###,##0.00")
         '   Set list_item = lv_existencias.ListItems.Add(, , "CCAJA")
         '   list_item.SubItems(1) = "TIENDA"
         '   list_item.SubItems(2) = Format(IIf(IsNull(rsaux!TIENDA), 0, rsaux!TIENDA), "###,###,##0.00")
         '   Set list_item = lv_existencias.ListItems.Add(, , "CCAJA")
         '   list_item.SubItems(1) = "EXHIBICION"
         '   list_item.SubItems(2) = Format(IIf(IsNull(rsaux!EXHIBICION), 0, rsaux!EXHIBICION), "###,###,##0.00")
         Else
            Me.txt_precio = ""
         End If
         rsaux.Close
         If var_clave_usuario_global <> "U0000000157" Then
            rsaux.Open "SELECT dbo.TB_ALMACENES.vcha_alm_almacen_id, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_EXISTENCIAS.FLOA_EXI_CANTIDAD, dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID fROM dbo.TB_EXISTENCIAS INNER JOIN dbo.TB_ALMACENES ON dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID where vcha_Art_Articulo_id = '" + var_codigo + "' ", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  If rsaux!VCHA_ALM_ALMACEN_ID <> "ACCAN" Then
                     If rsaux!VCHA_ALM_ALMACEN_ID <> "CA00231" Then
                        If rsaux!VCHA_ALM_ALMACEN_ID <> "CA00232" Then
                           If var_clave_usuario_global = "U0000000145" Or var_clave_usuario_global = "U0000000150" Or var_clave_usuario_global = "U0000000165" Or var_clave_usuario_global = "U0000000189" Or var_clave_usuario_global = "U0000000183" Or var_clave_usuario_global = "U0000000152" Or var_clave_usuario_global = "U0000000167" Or var_clave_usuario_global = "U0000000153" Then
                              Set list_item = lv_existencias.ListItems.Add(, , "SID")
                              list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_ALM_NOMBRE), "", rsaux!VCHA_ALM_NOMBRE)
                              list_item.SubItems(2) = Format(IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad), "###,###,##0.00")
                           Else
                              If rsaux!VCHA_ALM_ALMACEN_ID = "CC_1" Or rsaux!VCHA_ALM_ALMACEN_ID = "PTVH" Or rsaux!VCHA_ALM_ALMACEN_ID = "CC_5" Then
                                 var_cantidad = IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad)
                                 If var_cantidad > 0 Then
                                    Set list_item = lv_existencias.ListItems.Add(, , "SID")
                                    list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_ALM_NOMBRE), "", rsaux!VCHA_ALM_NOMBRE)
                                    list_item.SubItems(2) = Format(IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad), "###,###,##0.00")
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
                  rsaux.MoveNext
            Wend
            rsaux.Close
         Else
            rsaux.Open "SELECT dbo.TB_ALMACENES.vcha_alm_almacen_id, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_EXISTENCIAS.FLOA_EXI_CANTIDAD, dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID fROM dbo.TB_EXISTENCIAS INNER JOIN dbo.TB_ALMACENES ON dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID where vcha_Art_Articulo_id = '" + var_codigo + "' ", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  If IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad) <> 0 Then
                     Set list_item = lv_existencias.ListItems.Add(, , "SID")
                     list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_ALM_NOMBRE), "", rsaux!VCHA_ALM_NOMBRE)
                     list_item.SubItems(2) = Format(IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad), "###,###,##0.00")
                  End If
                  rsaux.MoveNext
            Wend
            rsaux.Close
         End If
      Else
         MsgBox "Código incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      Me.txt_descripcion = ""
      Me.lv_existencias.ListItems.Clear
   End If
End Sub

Private Sub txt_descripcion_GotFocus()
   Me.frm_disponibles.Visible = False
End Sub

Private Sub txt_descripcion_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.txt_nombre_articulo = ""
      Me.lv_disponibles.ListItems.Clear
      Me.frm_disponibles.Visible = True
      Me.txt_nombre_articulo.SetFocus
   End If
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_existencias.ListItems.Count > 0 Then
         Me.lv_existencias.SetFocus
      End If
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_nombre_articulo_Change()
   Me.lv_disponibles.ListItems.Clear
End Sub

Private Sub txt_nombre_articulo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 27 Then
      Me.txt_codigo.SetFocus
      Me.frm_disponibles.Visible = False
   End If
   If KeyAscii = 39 Or KeyAscii = 61 Or KeyAscii = 44 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(Me.txt_nombre_articulo) <> "" Then
         var_cadena = ""
         var_like_1 = ""
         var_like_2 = ""
         var_like_3 = ""
         var_like_4 = ""
         var_like_5 = ""
         var_like_6 = ""
         var_like_7 = ""
         var_j = 1
         For var_i = 1 To Len(Me.txt_nombre_articulo)
             If Mid(Me.txt_nombre_articulo, var_i, 1) <> " " Then
                If var_j = 1 Then
                   var_like_1 = var_like_1 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 2 Then
                   var_like_2 = var_like_2 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 3 Then
                   var_like_3 = var_like_3 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 4 Then
                   var_like_4 = var_like_4 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 5 Then
                   var_like_5 = var_like_5 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 6 Then
                   var_like_6 = var_like_6 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j >= 7 Then
                   var_like_7 = var_like_7 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
             Else
                var_j = var_j + 1
             End If
         Next var_i
      End If
      If Trim(var_like_1) <> "" Then
         var_cadena = " vcha_Art_nombre_Español  like '%" + var_like_1 + "%'"
      End If
      If Trim(var_like_2) <> "" Then
         var_cadena = var_cadena + " and  vcha_Art_nombre_Español like '%" + var_like_2 + "%'"
      End If
      If Trim(var_like_3) <> "" Then
         var_cadena = var_cadena + " and  vcha_Art_nombre_Español like '%" + var_like_3 + "%'"
      End If
      If Trim(var_like_4) <> "" Then
         var_cadena = var_cadena + " and  vcha_Art_nombre_Español like '%" + var_like_4 + "%'"
      End If
      If Trim(var_like_5) <> "" Then
         var_cadena = var_cadena + " and  vcha_Art_nombre_Español like '%" + var_like_5 + "%'"
      End If
      If Trim(var_like_6) <> "" Then
         var_cadena = var_cadena + " and  vcha_Art_nombre_Español like '%" + var_like_6 + "%'"
      End If
      If Trim(var_like_7) <> "" Then
         var_cadena = var_cadena + " and  vcha_Art_nombre_Español like '%" + var_like_7 + "%'"
      End If
      Me.lv_disponibles.ListItems.Clear
      If Trim(var_cadena) <> "" Then
      
         'var_cadena = "SELECT DISTINCT ART_CODIGO AS VCHA_ART_ARTICULO_ID, ART_DESCRIPCION AS VCHA_ART_NOMBRE_ESPAÑOL FROM VIA_EXISTENCIA_ALMACEN WHERE " + var_cadena
         var_cadena = "select * from tb_Articulos WHERE " + var_cadena
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
            Set list_item = lv_disponibles.ListItems.Add(, , rs!vcha_Art_Articulo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
            rs.MoveNext
         Wend
         rs.Close
         If Me.lv_disponibles.ListItems.Count > 0 Then
            Me.lv_disponibles.SetFocus
         End If
         If lv_disponibles.ListItems.Count > 11 Then
            lv_disponibles.ColumnHeaders(2).Width = 5300
         Else
            lv_disponibles.ColumnHeaders(2).Width = 5500
         End If
      End If
   End If
End Sub

Private Sub txt_precio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_codigo.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub
