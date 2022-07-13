VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_catalogos_vendidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de catálogos vendidos"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Height          =   60
      Left            =   0
      TabIndex        =   18
      Top             =   315
      Width           =   7335
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmreporte_catalogos_vendidos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6900
      Picture         =   "frmreporte_catalogos_vendidos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Caption         =   " Catálogos"
      Height          =   3135
      Left            =   90
      TabIndex        =   5
      Top             =   405
      Width           =   7215
      Begin VB.TextBox txt_busqueda 
         Height          =   315
         Left            =   930
         TabIndex        =   12
         Top             =   720
         Width           =   6165
      End
      Begin VB.Frame Frame2 
         Height          =   120
         Left            =   15
         TabIndex        =   14
         Top             =   990
         Width           =   7125
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Picture         =   "frmreporte_catalogos_vendidos.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   780
         Picture         =   "frmreporte_catalogos_vendidos.frx":0952
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Marcar (Enter)"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         Picture         =   "frmreporte_catalogos_vendidos.frx":0B9C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   120
         Picture         =   "frmreporte_catalogos_vendidos.frx":0C6E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         Picture         =   "frmreporte_catalogos_vendidos.frx":0D70
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   225
         Width           =   330
      End
      Begin VB.Frame Frame6 
         Height          =   120
         Left            =   30
         TabIndex        =   6
         Top             =   540
         Width           =   7140
      End
      Begin MSComctlLib.ListView lv_canales 
         Height          =   1935
         Left            =   60
         TabIndex        =   13
         Top             =   1125
         Width           =   7095
         _ExtentX        =   12515
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
            Object.Width           =   2469
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda:"
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   780
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Periodo "
      Height          =   705
      Left            =   90
      TabIndex        =   0
      Top             =   3570
      Width           =   7230
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   4140
         TabIndex        =   2
         Top             =   255
         Width           =   1095
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   2145
         TabIndex        =   1
         Top             =   255
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3840
         TabIndex        =   4
         Top             =   315
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   1695
         TabIndex        =   3
         Top             =   315
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmreporte_catalogos_vendidos"
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
         For var_i = 1 To Me.lv_canales.ListItems.Count
             lv_canales.ListItems.Item(var_i).Selected = True
             If Me.lv_canales.selectedItem.SubItems(2) = "*" Then
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
            For var_i = 1 To lv_canales.ListItems.Count
                lv_canales.ListItems.Item(var_i).Selected = True
                If lv_canales.selectedItem.SubItems(2) = "*" Then
                   rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_CANALES (INTE_TEM_CONSECUTIVO, VCHA_CAN_cANAL_VENTA_ID) VALUES (" + CStr(var_consecutivo) + ",'" + lv_canales.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                End If
            Next var_i
            cnn.CommandTimeout = 360
            rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_ARTICULOS (INTE_TEM_CONSECUTIVO, VCHA_ART_ARTICULO_ID) SELECT " + CStr(var_consecutivo) + ", VCHA_ART_ARTICULO_ID FROM TB_ARTICULOS", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "EXEC SP_REPORTE_ARTICULOS_VENDIDOS_PERIODO " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ",2", cnn, adOpenDynamic, adLockOptimistic
            Set reporte = appl.OpenReport(App.Path + "\rep_ventas_articulos_canal_concentrado.rpt")
            reporte.RecordSelectionFormula = "{VW_REPORTE_ARTICULOS_VENDIDOS_CANALES_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\Reporte_ventas_articulos_canal_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
            rs.Open "delete from TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_TEMPORAL where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_CANALES where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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
   For var_i = 1 To lv_canales.ListItems.Count
       lv_canales.ListItems.Item(var_i).Selected = True
       If lv_canales.selectedItem.SubItems(2) = "*" Then
          If Trim(var_cadena_reporte_articulos) = "" Then
             var_cadena_reporte_articulos = " vcha_can_canal_venta_id = '" + lv_canales.selectedItem + "'"
          Else
             var_cadena_reporte_articulos = var_cadena_reporte_articulos + " or vcha_can_canal_venta_id = '" + lv_canales.selectedItem + "'"
          End If
       End If
   Next var_i
   If var_cadena_reporte_articulos = "" Then
      MsgBox "Debe de seleccionar un canal de venta", vbOKOnly, "ATENCION"
   Else
      frmreporte_articulos_vendidos_articulos.txt_tipo_reporte = "Canales de Venta"
      frmreporte_articulos_vendidos_articulos.txt_inicio = Me.txt_inicio
      frmreporte_articulos_vendidos_articulos.txt_fin = Me.txt_fin
      frmreporte_articulos_vendidos_articulos.Show
   End If
End Sub

Private Sub cmd_imprimir_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         var_contador = 0
         For var_i = 1 To Me.lv_canales.ListItems.Count
             lv_canales.ListItems.Item(var_i).Selected = True
             If Me.lv_canales.selectedItem.SubItems(2) = "*" Then
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
            
            var_fecha_fin_1 = CDate(Me.txt_fin) + 1
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            cnn.BeginTrans
            rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_REPORTE_ARTICULOS_CATALOGOS_VENDIDOS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rs.Close
            rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_CATALOGOS_VENDIDOS (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            For var_i = 1 To lv_canales.ListItems.Count
                lv_canales.ListItems.Item(var_i).Selected = True
                If lv_canales.selectedItem.SubItems(2) = "*" Then
                   rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_CATALOGOS_VENDIDOS (INTE_TEM_CONSECUTIVO, vcha_Art_Articulo_id) VALUES (" + CStr(var_consecutivo) + ",'" + lv_canales.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                End If
            Next var_i
            cnn.CommandTimeout = 360
            rs.Open "EXEC SP_REPORTE_CATALOGOS_VENDIDOS " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin, cnn, adOpenDynamic, adLockOptimistic
            Set reporte = appl.OpenReport(App.Path + "\rep_catalogos_vendidos.rpt")
            reporte.RecordSelectionFormula = "{VW_REPORTE_cATALOGOS_VENDIDOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {VW_REPORTE_CATALOGOS_VENDIDOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de ventas"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_catalogos_vendidos.rpt")
               reporte.RecordSelectionFormula = "{VW_REPORTE_cATALOGOS_VENDIDOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {VW_REPORTE_CATALOGOS_VENDIDOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_catalogos_vendidos_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
               rs.Open "delete from TB_TEMP_REPORTE_ARTICULOS_CATALOGOS_VENDIDOS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               rs.Open "delete from TB_TEMP_REPORTE_CATALOGOS_VENDIDOS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               Frmmenu2.StatusBar1.Panels(1).Text = ""
            End If
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

Private Sub cmd_invertir_Click()
   n = lv_canales.ListItems.Count
   For i = 1 To n
      lv_canales.ListItems.Item(i).Selected = True
      If lv_canales.selectedItem.SubItems(2) = "*" Then
         lv_canales.selectedItem.SubItems(2) = ""
         lv_canales.ListItems.Item(i).Bold = False
         lv_canales.ListItems.Item(i).ForeColor = &H80000012
         lv_canales.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_canales.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_canales.selectedItem.SubItems(2) = "*"
         lv_canales.ListItems.Item(i).Bold = True
         lv_canales.ListItems.Item(i).ForeColor = &HFF0000
         lv_canales.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_canales.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_canales.selectedItem.Index
   If lv_canales.selectedItem.SubItems(2) = "*" Then
      lv_canales.selectedItem.SubItems(2) = ""
      lv_canales.ListItems.Item(i).Bold = False
      lv_canales.ListItems.Item(i).ForeColor = &H80000012
      lv_canales.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_canales.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_canales.Refresh
   Else
      lv_canales.selectedItem.SubItems(2) = "*"
      lv_canales.ListItems.Item(i).Bold = True
      lv_canales.ListItems.Item(i).ForeColor = &HFF0000
      lv_canales.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_canales.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_canales.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_canales.ListItems.Count
   For i = 1 To n
      lv_canales.ListItems.Item(i).Selected = True
      lv_canales.selectedItem.SubItems(2) = ""
      lv_canales.ListItems.Item(i).Bold = False
      lv_canales.ListItems.Item(i).ForeColor = &H80000012
      lv_canales.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_canales.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_canales.Refresh
End Sub

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_canales.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_canales.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_canales.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_canales.selectedItem.SubItems(2) = "*"
         lv_canales.ListItems.Item(i).Bold = True
         lv_canales.ListItems.Item(i).ForeColor = &HFF0000
         lv_canales.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_canales.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_canales.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_canales.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_siguiente_Click()
   var_cadena_reporte_articulos = ""
   For var_i = 1 To lv_canales.ListItems.Count
       lv_canales.ListItems.Item(var_i).Selected = True
       If lv_canales.selectedItem.SubItems(2) = "*" Then
          If Trim(var_cadena_reporte_articulos) = "" Then
             var_cadena_reporte_articulos = " vcha_can_canal_venta_id = '" + lv_canales.selectedItem + "'"
          Else
             var_cadena_reporte_articulos = var_cadena_reporte_articulos + " or vcha_can_canal_venta_id = '" + lv_canales.selectedItem + "'"
          End If
       End If
   Next var_i
   If var_cadena_reporte_articulos = "" Then
      MsgBox "Debe de seleccionar un canal de venta", vbOKOnly, "ATENCION"
   Else
      frmreporte_articulos_vendidos_agentes.Show
   End If
End Sub

Private Sub cmd_todos_Click()
   n = lv_canales.ListItems.Count
   For i = 1 To n
      lv_canales.ListItems.Item(i).Selected = True
      lv_canales.selectedItem.SubItems(2) = "*"
      lv_canales.ListItems.Item(i).Bold = True
      lv_canales.ListItems.Item(i).ForeColor = &HFF0000
      lv_canales.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_canales.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_canales.Refresh
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
   Left = 2500
   txt_inicio = Date
   txt_fin = Date
   'opt_linea = True
   rs.Open "select distinct vcha_Art_articulo_id, vcha_art_nombre_español from tb_articulos where  vcha_lin_linea_id = '90' order by vcha_art_nombre_Español", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      If IsNull(rs!vcha_Art_articulo_id) Then
      Else
         Set list_item = lv_canales.ListItems.Add(, , rs!vcha_Art_articulo_id)
         list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
         list_item.SubItems(2) = ""
      End If
      rs.MoveNext:
   Wend
   rs.Close
   If lv_canales.ListItems.Count > 8 Then
      lv_canales.ColumnHeaders(2).Width = 5400
   Else
      lv_canales.ColumnHeaders(2).Width = 5600
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_canales_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_canales, ColumnHeader)
End Sub

Private Sub lv_canales_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_canales.selectedItem.Index
      If lv_canales.selectedItem.SubItems(2) = "*" Then
         lv_canales.selectedItem.SubItems(2) = ""
         lv_canales.ListItems.Item(i).Bold = False
         lv_canales.ListItems.Item(i).ForeColor = &H80000012
         lv_canales.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_canales.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_canales.Refresh
      Else
         lv_canales.selectedItem.SubItems(2) = "*"
         lv_canales.ListItems.Item(i).Bold = True
         lv_canales.ListItems.Item(i).ForeColor = &HFF0000
         lv_canales.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_canales.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_canales.Refresh
      End If
   End If
End Sub

Private Sub txt_busqueda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "select top 1  vcha_can_nombre from vw_clientes where vcha_can_nombre like '%" + Me.txt_busqueda + "%' AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Call pro_busca_registro(lv_canales, rs!vcha_can_nombre, False)
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

