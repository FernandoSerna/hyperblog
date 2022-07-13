VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmbusqueda_articulo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda de artículos"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_nombre_articulo 
      Height          =   345
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   6750
   End
   Begin MSComctlLib.ListView lv_disponibles 
      Height          =   2910
      Left            =   45
      TabIndex        =   1
      Top             =   360
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   5133
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre del Artículo"
         Object.Width           =   7057
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Disponible"
         Object.Width           =   2470
      EndProperty
   End
End
Attribute VB_Name = "frmbusqueda_articulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lv_disponibles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_disponibles, ColumnHeader)
End Sub

Private Sub lv_disponibles_DblClick()
   If Me.lv_disponibles.ListItems.Count > 0 Then
      var_codigo_seleccionado = Me.lv_disponibles.selectedItem
      Unload Me
   End If
End Sub

Private Sub lv_disponibles_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_disponibles.ListItems.Count > 0 Then
         If var_empresa = "30" Then
            var_codigo_seleccionado = Me.lv_disponibles.selectedItem
            var_descripcion_global = Me.lv_disponibles.selectedItem.SubItems(1)
            rs.Open "select * from tb_detalle_lista_precios where vcha_Art_articulo_id = '" + var_codigo_seleccionado + "' and vcha_lis_lista_precios_id = '" + var_lista_precios_global + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_precio_global = IIf(IsNull(rs!floa_dli_precio), 0, rs!floa_dli_precio)
            Else
               var_precio_global = 0
            End If
            rs.Close
         Else
            var_codigo_seleccionado = Me.lv_disponibles.selectedItem
         End If
         Unload Me
      End If
   End If
End Sub

Private Sub txt_nombre_articulo_Change()
   Me.lv_disponibles.ListItems.Clear
End Sub

Private Sub txt_nombre_articulo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
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
         var_cadena = " vcha_art_nombre_Español like '%" + var_like_1 + "%'"
      End If
      If Trim(var_like_2) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_2 + "%'"
      End If
      If Trim(var_like_3) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_3 + "%'"
      End If
      If Trim(var_like_4) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_4 + "%'"
      End If
      If Trim(var_like_5) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_5 + "%'"
      End If
      If Trim(var_like_6) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_6 + "%'"
      End If
      If Trim(var_like_7) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_7 + "%'"
      End If
      Me.lv_disponibles.ListItems.Clear
      If Trim(var_cadena) <> "" Then
         If var_unidad_organizacional = "60" Then
            var_cadena = "SELECT * FROM tb_Articulos WHERE " + var_cadena + " and substring(vcha_Art_articulo_id,1,2) = 'SM'"
         Else
            If var_empresa = "30" Then
               var_cadena = "SELECT * FROM tb_Articulos WHERE " + var_cadena + " and substring(vcha_Art_articulo_id,1,2) = 'TR'"
            Else
               var_cadena = "SELECT * FROM tb_Articulos WHERE " + var_cadena
            End If
         End If
         rsaux9.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux9.EOF
            Set list_item = lv_disponibles.ListItems.Add(, , rsaux9!vcha_Art_Articulo_id)
            list_item.SubItems(1) = IIf(IsNull(rsaux9!VCHA_aRT_NOMBRE_ESPAÑOL), "", rsaux9!VCHA_aRT_NOMBRE_ESPAÑOL)
            'list_item.SubItems(2) = Format(Round(IIf(IsNull(rs!FLOA_EXI_CANTIDAD_DISPONIBLE), 0, rs!FLOA_EXI_CANTIDAD_DISPONIBLE), 4), "###,###,##0.0000")
            'If Mid(rs!vcha_Art_articulo_id, 11, 1) Then
            '   list_item.ForeColor = &HFF&
            '   list_item.ListSubItems(1).ForeColor = &HFF&
            '   list_item.ListSubItems(2).ForeColor = &HFF&
            'End If
            rsaux9.MoveNext
         Wend
         rsaux9.Close
         If Me.lv_disponibles.ListItems.Count > 0 Then
            Me.lv_disponibles.SetFocus
         End If
         If lv_disponibles.ListItems.Count > 11 Then
            lv_disponibles.ColumnHeaders(3).Width = 1200.18
         Else
            lv_disponibles.ColumnHeaders(3).Width = 1400.18
         End If
      End If
   End If
End Sub
