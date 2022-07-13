VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_articulos_vendidos_rutas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rutas"
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
      TabIndex        =   19
      Top             =   4125
      Width           =   1410
   End
   Begin VB.CommandButton cmd_ejecutar 
      Caption         =   "&Ejecutar"
      Height          =   435
      Left            =   2910
      TabIndex        =   18
      Top             =   4125
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Caption         =   " Periodo "
      Height          =   705
      Left            =   60
      TabIndex        =   13
      Top             =   3315
      Width           =   5700
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   3345
         TabIndex        =   15
         Top             =   255
         Width           =   1095
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   1350
         TabIndex        =   14
         Top             =   255
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3045
         TabIndex        =   17
         Top             =   315
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   900
         TabIndex        =   16
         Top             =   315
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Rutas "
      Height          =   3135
      Left            =   60
      TabIndex        =   2
      Top             =   150
      Width           =   5700
      Begin VB.TextBox txt_busqueda 
         Height          =   315
         Left            =   930
         TabIndex        =   9
         Top             =   720
         Width           =   4620
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Picture         =   "frmreporte_articulos_vendidos_rutas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   780
         Picture         =   "frmreporte_articulos_vendidos_rutas.frx":0216
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar (Enter)"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         Picture         =   "frmreporte_articulos_vendidos_rutas.frx":0460
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   120
         Picture         =   "frmreporte_articulos_vendidos_rutas.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         Picture         =   "frmreporte_articulos_vendidos_rutas.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   225
         Width           =   330
      End
      Begin VB.Frame Frame6 
         Height          =   120
         Left            =   30
         TabIndex        =   3
         Top             =   540
         Width           =   5565
      End
      Begin MSComctlLib.ListView lv_rutas 
         Height          =   1935
         Left            =   60
         TabIndex        =   10
         Top             =   1125
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   3413
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
      Begin VB.Frame Frame2 
         Height          =   120
         Left            =   15
         TabIndex        =   11
         Top             =   990
         Width           =   5565
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda:"
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   780
         Width           =   765
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "<<< &Anterior"
      Height          =   435
      Left            =   90
      TabIndex        =   1
      Top             =   4125
      Width           =   1410
   End
   Begin VB.CommandButton cmd_siguiente 
      Caption         =   "&Siguiente >>>"
      Height          =   435
      Left            =   4320
      TabIndex        =   0
      Top             =   4125
      Width           =   1410
   End
End
Attribute VB_Name = "frmreporte_articulos_vendidos_rutas"
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
         For var_i = 1 To Me.lv_rutas.ListItems.Count
             lv_rutas.ListItems.Item(var_i).Selected = True
             If Me.lv_rutas.selectedItem.SubItems(2) = "*" Then
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
            For var_i = 1 To lv_rutas.ListItems.Count
                lv_rutas.ListItems.Item(var_i).Selected = True
                If lv_rutas.selectedItem.SubItems(2) = "*" Then
                   rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_RUTAS (INTE_TEM_CONSECUTIVO, VCHA_RUT_RUTA_ID) VALUES (" + CStr(var_consecutivo) + ",'" + lv_rutas.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                End If
            Next var_i
            rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_ARTICULOS (INTE_TEM_CONSECUTIVO, VCHA_ART_ARTICULO_ID) SELECT " + CStr(var_consecutivo) + ", VCHA_ART_ARTICULO_ID FROM TB_ARTICULOS", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "EXEC SP_REPORTE_ARTICULOS_VENDIDOS_PERIODO " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ",4", cnn, adOpenDynamic, adLockOptimistic
            Set reporte = appl.OpenReport(App.Path + "\rep_ventas_articulos_rutas_concentrado.rpt")
            reporte.RecordSelectionFormula = "{VW_REPORTE_ARTICULOS_VENDIDOS_RUTAS_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\Reporte_ventas_articulos_rutas_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
            rs.Open "delete from TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_TEMPORAL where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_RUTAS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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
   var_cadena_reporte_articulos = ""
   For var_i = 1 To lv_rutas.ListItems.Count
       lv_rutas.ListItems.Item(var_i).Selected = True
       If lv_rutas.selectedItem.SubItems(2) = "*" Then
          If Trim(var_cadena_reporte_articulos) = "" Then
             var_cadena_reporte_articulos = " vcha_rut_ruta_id = '" + lv_rutas.selectedItem + "'"
          Else
             var_cadena_reporte_articulos = var_cadena_reporte_articulos + " or vcha_rut_ruta_id = '" + lv_rutas.selectedItem + "'"
          End If
       End If
   Next var_i
   If var_cadena_reporte_articulos = "" Then
      MsgBox "Debe de seleccionar una ruta", vbOKOnly, "ATENCION"
   Else
      frmreporte_articulos_vendidos_articulos.txt_tipo_reporte = "Rutas"
      frmreporte_articulos_vendidos_articulos.txt_inicio = Me.txt_inicio
      frmreporte_articulos_vendidos_articulos.txt_fin = Me.txt_fin
      frmreporte_articulos_vendidos_articulos.Show
   End If
End Sub

Private Sub cmd_invertir_Click()
   n = lv_rutas.ListItems.Count
   For i = 1 To n
      lv_rutas.ListItems.Item(i).Selected = True
      If lv_rutas.selectedItem.SubItems(2) = "*" Then
         lv_rutas.selectedItem.SubItems(2) = ""
         lv_rutas.ListItems.Item(i).Bold = False
         lv_rutas.ListItems.Item(i).ForeColor = &H80000012
         lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_rutas.selectedItem.SubItems(2) = "*"
         lv_rutas.ListItems.Item(i).Bold = True
         lv_rutas.ListItems.Item(i).ForeColor = &HFF0000
         lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_rutas.selectedItem.Index
   If lv_rutas.selectedItem.SubItems(2) = "*" Then
      lv_rutas.selectedItem.SubItems(2) = ""
      lv_rutas.ListItems.Item(i).Bold = False
      lv_rutas.ListItems.Item(i).ForeColor = &H80000012
      lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_rutas.Refresh
   Else
      lv_rutas.selectedItem.SubItems(2) = "*"
      lv_rutas.ListItems.Item(i).Bold = True
      lv_rutas.ListItems.Item(i).ForeColor = &HFF0000
      lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_rutas.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_rutas.ListItems.Count
   For i = 1 To n
      lv_rutas.ListItems.Item(i).Selected = True
      lv_rutas.selectedItem.SubItems(2) = ""
      lv_rutas.ListItems.Item(i).Bold = False
      lv_rutas.ListItems.Item(i).ForeColor = &H80000012
      lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_rutas.Refresh
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_rutas.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_rutas.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_rutas.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_rutas.selectedItem.SubItems(2) = "*"
         lv_rutas.ListItems.Item(i).Bold = True
         lv_rutas.ListItems.Item(i).ForeColor = &HFF0000
         lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_rutas.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_rutas.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_siguiente_Click()
   var_cadena_reporte_articulos = ""
   For var_i = 1 To lv_rutas.ListItems.Count
       lv_rutas.ListItems.Item(var_i).Selected = True
       If lv_rutas.selectedItem.SubItems(2) = "*" Then
          If Trim(var_cadena_reporte_articulos) = "" Then
             var_cadena_reporte_articulos = " vcha_rut_ruta_id = '" + lv_rutas.selectedItem + "'"
          Else
             var_cadena_reporte_articulos = var_cadena_reporte_articulos + " or vcha_rut_ruta_id = '" + lv_rutas.selectedItem + "'"
          End If
       End If
   Next var_i
   If var_cadena_reporte_articulos = "" Then
      MsgBox "Debe de seleccionar una ruta", vbOKOnly, "ATENCION"
   Else
      frmreporte_articulos_vendidos_titulares.Show
   End If
End Sub

Private Sub cmd_todos_Click()
   n = lv_rutas.ListItems.Count
   For i = 1 To n
      lv_rutas.ListItems.Item(i).Selected = True
      lv_rutas.selectedItem.SubItems(2) = "*"
      lv_rutas.ListItems.Item(i).Bold = True
      lv_rutas.ListItems.Item(i).ForeColor = &HFF0000
      lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_rutas.Refresh
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 65 Then
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
   var_cadena_seguridad = ""
   Top = 1500
   Left = 3200
   txt_inicio = Date
   txt_fin = Date
   'opt_linea = True
   rs.Open "select distinct vcha_rut_ruta_id, vcha_rut_nombre from vw_clientes where  (" + var_cadena_reporte_articulos + ")  order by vcha_rut_nombre ", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      If IsNull(rs!VCHA_RUT_RUTA_ID) Then
      Else
         Set list_item = lv_rutas.ListItems.Add(, , rs!VCHA_RUT_RUTA_ID)
         list_item.SubItems(1) = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
         list_item.SubItems(2) = ""
      End If
      rs.MoveNext:
   Wend
   rs.Close
   If lv_rutas.ListItems.Count > 8 Then
      lv_rutas.ColumnHeaders(2).Width = 4220
   Else
      lv_rutas.ColumnHeaders(2).Width = 4400.71
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_salidas)
End Sub

Private Sub lv_rutas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_rutas, ColumnHeader)
End Sub

Private Sub lv_rutas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_rutas.selectedItem.Index
      If lv_rutas.selectedItem.SubItems(2) = "*" Then
         lv_rutas.selectedItem.SubItems(2) = ""
         lv_rutas.ListItems.Item(i).Bold = False
         lv_rutas.ListItems.Item(i).ForeColor = &H80000012
         lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_rutas.Refresh
      Else
         lv_rutas.selectedItem.SubItems(2) = "*"
         lv_rutas.ListItems.Item(i).Bold = True
         lv_rutas.ListItems.Item(i).ForeColor = &HFF0000
         lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_rutas.Refresh
      End If
   End If
End Sub

Private Sub txt_busqueda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "select top 1  vcha_can_nombre from vw_clientes where vcha_can_nombre like '%" + Me.txt_busqueda + "%' AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Call pro_busca_registro(lv_rutas, rs!vcha_can_nombre, False)
      Else
         MsgBox "No existen coincidencias", vbOKOnly, "ATENCION"
      End If
      rs.Close
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
