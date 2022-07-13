VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmconcentrado_ordenes_tiendas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Concentrado de ordenes de surtido"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmconcentrado_ordenes_tiendas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Nuevo "
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aplicar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmconcentrado_ordenes_tiendas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9690
      Picture         =   "frmconcentrado_ordenes_tiendas.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   75
      TabIndex        =   2
      Top             =   315
      Width           =   9990
   End
   Begin VB.Frame frm_ordenes 
      Caption         =   " Ordenes de Surtido "
      Height          =   5385
      Left            =   105
      TabIndex        =   0
      Top             =   405
      Width           =   9930
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         Picture         =   "frmconcentrado_ordenes_tiendas.frx":083E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   750
         Picture         =   "frmconcentrado_ordenes_tiendas.frx":0A54
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Marcar (Enter)"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         Picture         =   "frmconcentrado_ordenes_tiendas.frx":0C9E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   90
         Picture         =   "frmconcentrado_ordenes_tiendas.frx":0D70
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   420
         Picture         =   "frmconcentrado_ordenes_tiendas.frx":0E72
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   240
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_ordenes 
         Height          =   4680
         Left            =   90
         TabIndex        =   1
         Top             =   615
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   8255
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Orden de S."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Tienda"
            Object.Width           =   6932
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cantidad"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmconcentrado_ordenes_tiendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_aplicar_Click()
   cnn.BeginTrans
   var_consecutivo = 0
   rs.Open "select max(INTE_TEM_CONSECUTIVO) from TB_TEMP_CONCENTRADO_ORDENES_SURTIDO_TIENDA", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
   Else
      var_consecutivo = 1
   End If
   rs.Close
   rs.Open "insert into TB_TEMP_CONCENTRADO_ORDENES_SURTIDO_TIENDA (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
   cnn.CommitTrans
   VAR_CADENA_FILTRO = ""
   var_ordenes = ""
   For var_i = 1 To Me.lv_ordenes.ListItems.Count
       lv_ordenes.ListItems.Item(var_i).Selected = True
       If lv_ordenes.selectedItem.SubItems(5) = "*" Then
          If VAR_CADENA_FILTRO = "" Then
             VAR_CADENA_FILTRO = " inte_ors_orden_surtido = " + Me.lv_ordenes.selectedItem.SubItems(1)
             var_ordenes = Me.lv_ordenes.selectedItem.SubItems(1)
          Else
             VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " or inte_ors_orden_surtido = " + Me.lv_ordenes.selectedItem.SubItems(1)
             var_ordenes = var_ordenes + ", " + Me.lv_ordenes.selectedItem.SubItems(1)
          End If
       End If
   Next var_i
   If var_ordenes <> "" Then
      var_cadena = "INSERT INTO TB_TEMP_CONCENTRADO_ORDENES_SURTIDO_TIENDA (INTE_TEM_CONSECUTIVO, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_ESPAÑOL, FLOA_TEM_CANTIDAD, VCHA_TEM_ORDENES) SELECT " + CStr(var_consecutivo) + ",dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID, dbo.TB_Articulos.VCHA_ART_NOMBRE_ESPAÑOL, SUM(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR) AS CANTIDAD,  '" + CStr(var_ordenes) + "'"
      var_cadena = var_cadena + " FROM dbo.TB_DET_ORDEN_SURTIDO INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID Where (" + VAR_CADENA_FILTRO + ") GROUP BY dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID"
      rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
      rs.Open "select * from TB_TEMP_CONCENTRADO_ORDENES_SURTIDO_TIENDA where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            rsaux.Open "select * from tb_ubicaciones_almacen where vcha_Art_articulo_id = '" + IIf(IsNull(rs!vcha_Art_Articulo_id), "", rs!vcha_Art_Articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               rsaux2.Open "update TB_TEMP_CONCENTRADO_ORDENES_SURTIDO_TIENDA set vcha_tem_ubicacion = '" + IIf(IsNull(rsaux!vcha_ubi_ubicacion_1), "", rsaux!vcha_ubi_ubicacion_1) + "' where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_art_articulo_id = '" + IIf(IsNull(rs!vcha_Art_Articulo_id), "", rs!vcha_Art_Articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            rsaux.Close
            rs.MoveNext
      Wend
      rs.Close
      Set reporte = appl.OpenReport(App.Path + "\rep_consentrado_ordenes_surtido_tiendas.rpt")
      reporte.RecordSelectionFormula = "{TB_TEMP_CONCENTRADO_ORDENES_SURTIDO_TIENDA.inte_tem_consecutivo}= " + CStr(var_consecutivo) + " and {TB_TEMP_CONCENTRADO_ORDENES_SURTIDO_TIENDA.FLOA_TEM_CANTIDAD} > 0"
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Concentrado de ordenes de surtido"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      rs.Open "DELETE FROM TB_TEMP_CONCENTRADO_ORDENES_SURTIDO_TIENDA WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   Else
      MsgBox "No se a seleccionado ninguna orden de surtido", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_invertir_Click()
   If var_todos_articulos = 1 Then
   Else
        var_todos_articulos = 0
   End If
   n = lv_ordenes.ListItems.Count
   For i = 1 To n
      lv_ordenes.ListItems.Item(i).Selected = True
      If lv_ordenes.selectedItem.SubItems(5) = "*" Then
         lv_ordenes.selectedItem.SubItems(5) = ""
         lv_ordenes.ListItems.Item(i).Bold = False
         lv_ordenes.ListItems.Item(i).ForeColor = &H80000012
         lv_ordenes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_ordenes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_ordenes.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_ordenes.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_ordenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_ordenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_ordenes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_ordenes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      Else
         lv_ordenes.selectedItem.SubItems(5) = "*"
         lv_ordenes.ListItems.Item(i).Bold = True
         lv_ordenes.ListItems.Item(i).ForeColor = &H8000&
         lv_ordenes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_ordenes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_ordenes.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_ordenes.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_ordenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
         lv_ordenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
         lv_ordenes.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
         lv_ordenes.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   var_todos_articulos = 0
   i = lv_ordenes.selectedItem.Index
   If lv_ordenes.selectedItem.SubItems(5) = "*" Then
      lv_ordenes.selectedItem.SubItems(5) = ""
      lv_ordenes.ListItems.Item(i).Bold = False
      lv_ordenes.ListItems.Item(i).ForeColor = &H80000012
      lv_ordenes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_ordenes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_ordenes.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_ordenes.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_ordenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_ordenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_ordenes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_ordenes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_ordenes.Refresh
   Else
      lv_ordenes.selectedItem.SubItems(5) = "*"
      lv_ordenes.ListItems.Item(i).Bold = True
      lv_ordenes.ListItems.Item(i).ForeColor = &H8000&
      lv_ordenes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_ordenes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_ordenes.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_ordenes.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_ordenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
      lv_ordenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
      lv_ordenes.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
      lv_ordenes.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
      lv_ordenes.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
    
   n = lv_ordenes.ListItems.Count
   For i = 1 To n
       lv_ordenes.ListItems.Item(i).SubItems(5) = " "
       lv_ordenes.ListItems.Item(i).Bold = False
       lv_ordenes.ListItems.Item(i).ForeColor = &H80000012
       lv_ordenes.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_ordenes.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_ordenes.ListItems.Item(i).ListSubItems(3).Bold = False
       lv_ordenes.ListItems.Item(i).ListSubItems(4).Bold = False
       lv_ordenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_ordenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
       lv_ordenes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
       lv_ordenes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
    Next
    lv_ordenes.Refresh
End Sub

Private Sub cmd_nuevo_Click()
   Me.lv_ordenes.ListItems.Clear
   If var_unidad_organizacional = "23" Then
      Me.lv_ordenes.ColumnHeaders.Item(4) = "Tienda"
      cnn.CommandTimeout = 6000
      rs.Open "select * from VW_CONCENTRADO_ORDENES_SURTIDO_TIENDAS where vcha_emp_empresa_id = '" + var_empresa + "' and CHAR_TPE_TIPO_PEDIDO_ID = 'FT'", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_ordenes.ListItems.Add(, , rs!inte_ped_numero)
            list_item.SubItems(1) = IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), "", rs!INTE_ORS_ORDEN_SURTIDO)
            list_item.SubItems(2) = IIf(IsNull(rs!dtim_ped_fecha), "", rs!dtim_ped_fecha)
            list_item.SubItems(3) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            list_item.SubItems(4) = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), "", rs!FLOA_ORS_CANTIDAD_SURTIR)
            rs.MoveNext
            
      Wend
      rs.Close
   Else
      Me.lv_ordenes.ColumnHeaders.Item(4) = "Agente"
      cnn.CommandTimeout = 6000
      rs.Open "select * from VW_CONCENTRADO_ORDENES_SURTIDO_TIENDAS where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            var_cantidad = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), "", rs!FLOA_ORS_CANTIDAD_SURTIR)
            If var_cantidad > 0 Then
               Set list_item = lv_ordenes.ListItems.Add(, , rs!inte_ped_numero)
               list_item.SubItems(1) = IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), "", rs!INTE_ORS_ORDEN_SURTIDO)
               list_item.SubItems(2) = IIf(IsNull(rs!dtim_ped_fecha), "", rs!dtim_ped_fecha)
               list_item.SubItems(3) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
               list_item.SubItems(4) = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), "", rs!FLOA_ORS_CANTIDAD_SURTIR)
            End If
            rs.MoveNext
      Wend
      rs.Close
   
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   If var_todos_articulos = 1 Then
   Else
         var_todos_articulos = 0
   End If
   n = lv_ordenes.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_ordenes.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_ordenes.selectedItem.SubItems(5) = "" And var_rellena = True Then
         lv_ordenes.selectedItem.SubItems(5) = "*"
         lv_ordenes.ListItems.Item(i).Bold = True
         lv_ordenes.ListItems.Item(i).ForeColor = &H8000&
         lv_ordenes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_ordenes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_ordenes.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_ordenes.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_ordenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
         lv_ordenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
         lv_ordenes.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
         lv_ordenes.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
      Else
         If var_encontro = True And lv_ordenes.selectedItem.SubItems(5) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_ordenes.selectedItem.SubItems(5) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i

End Sub

Private Sub cmd_todos_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_ordenes.ListItems.Count
   For i = 1 To n
       lv_ordenes.ListItems.Item(i).SubItems(5) = "*"
       lv_ordenes.ListItems.Item(i).Bold = True
       lv_ordenes.ListItems.Item(i).ForeColor = &H8000&
       lv_ordenes.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_ordenes.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_ordenes.ListItems.Item(i).ListSubItems(3).Bold = True
       lv_ordenes.ListItems.Item(i).ListSubItems(4).Bold = True
       lv_ordenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_ordenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
       lv_ordenes.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
       lv_ordenes.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
   Next
   lv_ordenes.Refresh
End Sub

Private Sub Form_Load()
    Top = 1000
    Left = 800
   If var_unidad_organizacional = "23" Then
      Me.lv_ordenes.ColumnHeaders.Item(4) = "Tienda"
   Else
      Me.lv_ordenes.ColumnHeaders.Item(4) = "Agente"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_ordenes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_ordenes, ColumnHeader)
End Sub

Private Sub lv_ordenes_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then
      Unload Me
   End If
End Sub

Private Sub lv_ordenes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_ordenes.selectedItem.Index
      If lv_ordenes.selectedItem.SubItems(5) = "*" Then
         lv_ordenes.selectedItem.SubItems(5) = ""
         lv_ordenes.ListItems.Item(i).Bold = False
         lv_ordenes.ListItems.Item(i).ForeColor = &H80000012
         lv_ordenes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_ordenes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_ordenes.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_ordenes.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_ordenes.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_ordenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_ordenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_ordenes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_ordenes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_ordenes.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_ordenes.Refresh
      Else
         lv_ordenes.selectedItem.SubItems(5) = "*"
         lv_ordenes.ListItems.Item(i).Bold = True
         lv_ordenes.ListItems.Item(i).ForeColor = &H8000&
         lv_ordenes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_ordenes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_ordenes.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_ordenes.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_ordenes.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_ordenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
         lv_ordenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
         lv_ordenes.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
         lv_ordenes.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
         lv_ordenes.ListItems.Item(i).ListSubItems(5).ForeColor = &H8000&
         lv_ordenes.Refresh
      End If
   End If
End Sub
