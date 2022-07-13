VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_articulos_vendidos_tipo_canal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipo de Canal de Venta"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_filtro 
      Caption         =   "&Filtrar por artículo"
      Height          =   435
      Left            =   1500
      TabIndex        =   10
      Top             =   4125
      Width           =   1410
   End
   Begin VB.CommandButton cmd_ejecutar 
      Caption         =   "&Ejecutar"
      Height          =   435
      Left            =   2910
      TabIndex        =   9
      Top             =   4125
      Width           =   1410
   End
   Begin VB.Frame Frame2 
      Caption         =   " Periodo "
      Height          =   705
      Left            =   60
      TabIndex        =   4
      Top             =   3315
      Width           =   5685
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   3345
         TabIndex        =   6
         Top             =   255
         Width           =   1095
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   1350
         TabIndex        =   5
         Top             =   255
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3045
         TabIndex        =   8
         Top             =   315
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   900
         TabIndex        =   7
         Top             =   315
         Width           =   420
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   90
      TabIndex        =   3
      Top             =   4125
      Width           =   1410
   End
   Begin VB.CommandButton cmd_siguiente 
      Caption         =   "&Siguiente >>>"
      Height          =   435
      Left            =   4320
      TabIndex        =   2
      Top             =   4125
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Caption         =   " Tipo de Canal de Venta "
      Height          =   3135
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   5700
      Begin MSComctlLib.ListView lv_tipos 
         Height          =   2775
         Left            =   60
         TabIndex        =   1
         Top             =   225
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   4895
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
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmreporte_articulos_vendidos_tipo_canal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_cancelar_Click()
   Unload Me
End Sub

Private Sub cmd_ejecutar_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         var_contador = 0
         For var_i = 1 To Me.lv_tipos.ListItems.Count
             lv_tipos.ListItems.Item(var_i).Selected = True
             If Me.lv_tipos.selectedItem.SubItems(2) = "*" Then
                var_contador = var_contador + 1
             End If
         Next var_i
         If var_contador > 0 Then
            Frmmenu2.StatusBar1.Panels(1).Text = "Procesando información espere un momento"
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_dia = CStr(Day(txt_fin))
            var_mes = CStr(Month(txt_fin))
            var_año = CStr(Year(txt_fin))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            cnn.CommandTimeout = 6000
            cnn.BeginTrans
            rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_TEMPORAL", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rs.Close
            rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_TEMPORAL (INTE_TEM_CONSECUTIVO, DTIM_TEM_fecha_INICIO, DTIM_TEM_fecha_FIN) VALUES (" + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            For var_i = 1 To lv_tipos.ListItems.Count
                lv_tipos.ListItems.Item(var_i).Selected = True
                If lv_tipos.selectedItem.SubItems(2) = "*" Then
                   rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_TIPOS_CANALES (INTE_TEM_CONSECUTIVO, CHAR_TPE_TIPO_PEDIDO_ID) VALUES (" + CStr(var_consecutivo) + ",'" + lv_tipos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                End If
            Next var_i
            rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_ARTICULOS (INTE_TEM_CONSECUTIVO, VCHA_ART_ARTICULO_ID) SELECT " + CStr(var_consecutivo) + ", VCHA_ART_ARTICULO_ID FROM TB_ARTICULOS", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "EXEC SP_REPORTE_ARTICULOS_VENDIDOS_PERIODO " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ",1", cnn, adOpenDynamic, adLockOptimistic
            Set reporte = appl.OpenReport(App.Path + "\rep_ventas_articulos_tipo_canal_concentrado.rpt")
            reporte.RecordSelectionFormula = "{VW_REPORTE_ARTICULOS_VENDIDOS_TIPO_CANAL_GENERAL.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\Reporte_ventas_articulos_tipo_canal_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
            rs.Open "delete from TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_TEMPORAL where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_TIPOS_CANALES where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "DELETE FROM TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_ARTICULOS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            Frmmenu2.StatusBar1.Panels(1).Text = ""

         Else
             MsgBox "No se a seleccionado un tipo de canal", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_filtro_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         var_cadena_reporte_articulos = ""
         For var_i = 1 To lv_tipos.ListItems.Count
             lv_tipos.ListItems.Item(var_i).Selected = True
             If lv_tipos.selectedItem.SubItems(2) = "*" Then
                If Trim(var_cadena_reporte_articulos) = "" Then
                   var_cadena_reporte_articulos = " char_tpe_tipo_pedido_id = '" + Me.lv_tipos.selectedItem + "'"
                Else
                   var_cadena_reporte_articulos = var_cadena_reporte_articulos + " or char_tpe_tipo_pedido_id = '" + Me.lv_tipos.selectedItem + "'"
                End If
             End If
         Next var_i
         If Trim(var_cadena_reporte_articulos) <> "" Then
            frmreporte_articulos_vendidos_articulos.txt_tipo_reporte = "Tipo canal de venta"
            frmreporte_articulos_vendidos_articulos.txt_inicio = Me.txt_inicio
            frmreporte_articulos_vendidos_articulos.txt_fin = Me.txt_fin
            frmreporte_articulos_vendidos_articulos.Show
         Else
            MsgBox "Debe de seleccionar algún tipo de canal", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_siguiente_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         var_cadena_reporte_articulos = ""
         For var_i = 1 To lv_tipos.ListItems.Count
             lv_tipos.ListItems.Item(var_i).Selected = True
             If lv_tipos.selectedItem.SubItems(2) = "*" Then
                If Trim(var_cadena_reporte_articulos) = "" Then
                   var_cadena_reporte_articulos = " char_tpe_tipo_pedido_id = '" + Me.lv_tipos.selectedItem + "'"
                Else
                   var_cadena_reporte_articulos = var_cadena_reporte_articulos + " or char_tpe_tipo_pedido_id = '" + Me.lv_tipos.selectedItem + "'"
                End If
             End If
         Next var_i
         If Trim(var_cadena_reporte_articulos) <> "" Then
            frmreporte_articulos_vendidos_canales.Show
         Else
            MsgBox "Debe de seleccionar algún tipo de canal", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 67 Then
      cmd_cancelar_Click
   End If
   If Shift = 4 And KeyCode = 69 Then
      cmd_ejecutar_Click
   End If
   If Shift = 4 And KeyCode = 83 Then
      cmd_siguiente_Click
   End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Top = 1500
   Left = 3200
   Me.txt_fin = Date
   Me.txt_inicio = Date
   rs.Open "select distinct char_tpe_tipo_pedido_id, vcha_tpe_nombre from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' AND INTE_TPE_REPORTE =  1 and vcha_rut_ruta_id is not null and vcha_tit_nombre <> '' and char_tpe_tipo_pedido_id is not null order by vcha_tpe_nombre ", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      Set list_item = lv_tipos.ListItems.Add(, , IIf(IsNull(rs(0).Value), "", rs(0).Value))
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_tpe_NOMBRE), "", rs!VCHA_tpe_NOMBRE)
      list_item.SubItems(2) = ""
      rs.MoveNext:
   Wend
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_tipos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_tipos, ColumnHeader)
End Sub

Private Sub lv_tipos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
   If KeyAscii = 13 Then
      i = lv_tipos.selectedItem.Index
      If lv_tipos.selectedItem.SubItems(2) = "*" Then
         lv_tipos.selectedItem.SubItems(2) = ""
         lv_tipos.ListItems.Item(i).Bold = False
         lv_tipos.ListItems.Item(i).ForeColor = &H80000012
         lv_tipos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_tipos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_tipos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_tipos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_tipos.Refresh
      Else
         lv_tipos.selectedItem.SubItems(2) = "*"
         lv_tipos.ListItems.Item(i).Bold = True
         lv_tipos.ListItems.Item(i).ForeColor = &HFF0000
         lv_tipos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_tipos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_tipos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_tipos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_tipos.Refresh
      End If
   End If
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes.Value = CDate(Me.txt_fin)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes.Value = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub
