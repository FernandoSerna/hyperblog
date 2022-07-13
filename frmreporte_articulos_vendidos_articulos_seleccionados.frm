VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_articulos_vendidos_articulos_seleccionados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Artículos seleccionados"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_ejecutar 
      Caption         =   "&Ejecutar"
      Height          =   435
      Left            =   4695
      TabIndex        =   10
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "<<< &Anterior"
      Height          =   435
      Left            =   3345
      TabIndex        =   9
      Top             =   6840
      Width           =   1350
   End
   Begin VB.Frame Frame1 
      Caption         =   " Artículos "
      Height          =   6780
      Left            =   60
      TabIndex        =   0
      Top             =   -15
      Width           =   5880
      Begin VB.TextBox txt_codigo 
         Height          =   315
         Left            =   1035
         TabIndex        =   11
         Top             =   975
         Width           =   1860
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   465
         Picture         =   "frmreporte_articulos_vendidos_articulos_seleccionados.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   255
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   135
         Picture         =   "frmreporte_articulos_vendidos_articulos_seleccionados.frx":0216
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   255
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1125
         Picture         =   "frmreporte_articulos_vendidos_articulos_seleccionados.frx":0318
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   255
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   795
         Picture         =   "frmreporte_articulos_vendidos_articulos_seleccionados.frx":03EA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Marcar (Enter)"
         Top             =   255
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1455
         Picture         =   "frmreporte_articulos_vendidos_articulos_seleccionados.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   255
         Width           =   330
      End
      Begin VB.TextBox txt_busqueda 
         Height          =   315
         Left            =   1035
         TabIndex        =   1
         Top             =   630
         Width           =   4770
      End
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   5385
         Left            =   60
         TabIndex        =   7
         Top             =   1320
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   9499
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
            Object.Width           =   7497
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Código:"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   1005
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Busqueda:"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   660
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmreporte_articulos_vendidos_articulos_seleccionados"
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
   var_cadena_articulos = ""
   For var_i = 1 To lv_articulos.ListItems.Count
       lv_articulos.ListItems.Item(var_i).Selected = True
       If lv_articulos.selectedItem.SubItems(2) = "*" Then
          If Trim(var_cadena_articulos) = "" Then
             var_cadena_articulos = "( VCHA_ART_ARTICULO_ID = '" + lv_articulos.selectedItem + "'"
          Else
             var_cadena_articulos = var_cadena_articulos + " OR VCHA_ART_ARTICULO_ID = '" + lv_articulos.selectedItem + "'"
          End If
       End If
   Next var_i
   If var_cadena_articulos <> "" Then
      var_cadena_articulos = var_cadena_articulos + ")"
      If frmreporte_articulos_vendidos_articulos.txt_tipo_reporte = "Tipo canal de venta" Then
         If IsDate(frmreporte_articulos_vendidos_articulos.txt_inicio) Then
            If IsDate(frmreporte_articulos_vendidos_articulos.txt_fin) Then
               var_contador = 0
               For var_i = 1 To frmreporte_articulos_vendidos_tipo_canal.lv_tipos.ListItems.Count
                   frmreporte_articulos_vendidos_tipo_canal.lv_tipos.ListItems.Item(var_i).Selected = True
                   If frmreporte_articulos_vendidos_tipo_canal.lv_tipos.selectedItem.SubItems(2) = "*" Then
                      var_contador = var_contador + 1
                   End If
               Next var_i
               If var_contador > 0 Then
                  Frmmenu2.StatusBar1.Panels(1).Text = "Procesando información espere un momento"
                  var_dia = CStr(Day(CDate(frmreporte_articulos_vendidos_articulos.txt_inicio)))
                  var_mes = CStr(Month(CDate(frmreporte_articulos_vendidos_articulos.txt_inicio)))
                  var_año = CStr(Year(CDate(frmreporte_articulos_vendidos_articulos.txt_inicio)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
               
                  
                  var_dia = CStr(Day(frmreporte_articulos_vendidos_articulos.txt_fin))
                  var_mes = CStr(Month(frmreporte_articulos_vendidos_articulos.txt_fin))
                  var_año = CStr(Year(frmreporte_articulos_vendidos_articulos.txt_fin))
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
                  For var_i = 1 To frmreporte_articulos_vendidos_tipo_canal.lv_tipos.ListItems.Count
                      frmreporte_articulos_vendidos_tipo_canal.lv_tipos.ListItems.Item(var_i).Selected = True
                      If frmreporte_articulos_vendidos_tipo_canal.lv_tipos.selectedItem.SubItems(2) = "*" Then
                         rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_TIPOS_CANALES (INTE_TEM_CONSECUTIVO, CHAR_TPE_TIPO_PEDIDO_ID) VALUES (" + CStr(var_consecutivo) + ",'" + frmreporte_articulos_vendidos_tipo_canal.lv_tipos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                      End If
                  Next var_i
                  For var_i = 1 To lv_articulos.ListItems.Count
                      lv_articulos.ListItems.Item(var_i).Selected = True
                      If lv_articulos.selectedItem.SubItems(2) = "*" Then
                         rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_ARTICULOS (INTE_TEM_CONSECUTIVO, VCHA_ART_ARTICULO_ID) VALUES (" + CStr(var_consecutivo) + ", '" + lv_articulos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
       
                      End If
                  Next var_i
                  
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
      End If
      If frmreporte_articulos_vendidos_articulos.txt_tipo_reporte = "Canales de Venta" Then
         If IsDate(frmreporte_articulos_vendidos_articulos.txt_inicio) Then
            If IsDate(frmreporte_articulos_vendidos_articulos.txt_fin) Then
               var_contador = 0
               For var_i = 1 To frmreporte_articulos_vendidos_canales.lv_canales.ListItems.Count
                   frmreporte_articulos_vendidos_canales.lv_canales.ListItems.Item(var_i).Selected = True
                   If frmreporte_articulos_vendidos_canales.lv_canales.selectedItem.SubItems(2) = "*" Then
                      var_contador = var_contador + 1
                   End If
               Next var_i
               If var_contador > 0 Then
                  Frmmenu2.StatusBar1.Panels(1).Text = "Procesando información espere un momento"
                  var_dia = CStr(Day(CDate(frmreporte_articulos_vendidos_articulos.txt_inicio)))
                  var_mes = CStr(Month(CDate(frmreporte_articulos_vendidos_articulos.txt_inicio)))
                  var_año = CStr(Year(CDate(frmreporte_articulos_vendidos_articulos.txt_inicio)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
              
              
                  var_dia = CStr(Day(frmreporte_articulos_vendidos_articulos.txt_fin))
                  var_mes = CStr(Month(frmreporte_articulos_vendidos_articulos.txt_fin))
                  var_año = CStr(Year(frmreporte_articulos_vendidos_articulos.txt_fin))
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
                  For var_i = 1 To frmreporte_articulos_vendidos_canales.lv_canales.ListItems.Count
                      frmreporte_articulos_vendidos_canales.lv_canales.ListItems.Item(var_i).Selected = True
                      If frmreporte_articulos_vendidos_canales.lv_canales.selectedItem.SubItems(2) = "*" Then
                         rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_CANALES (INTE_TEM_CONSECUTIVO, VCHA_CAN_cANAL_VENTA_ID) VALUES (" + CStr(var_consecutivo) + ",'" + frmreporte_articulos_vendidos_canales.lv_canales.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                      End If
                  Next var_i
                  cnn.CommandTimeout = 360
                  For var_i = 1 To lv_articulos.ListItems.Count
                      lv_articulos.ListItems.Item(var_i).Selected = True
                      If lv_articulos.selectedItem.SubItems(2) = "*" Then
                         rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_ARTICULOS (INTE_TEM_CONSECUTIVO, VCHA_ART_ARTICULO_ID) VALUES (" + CStr(var_consecutivo) + ", '" + lv_articulos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
       
                      End If
                  Next var_i
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
      End If
      If frmreporte_articulos_vendidos_articulos.txt_tipo_reporte = "Agentes" Then
         If IsDate(frmreporte_articulos_vendidos_articulos.txt_inicio) Then
            If IsDate(frmreporte_articulos_vendidos_articulos.txt_fin) Then
               var_contador = 0
               For var_i = 1 To frmreporte_articulos_vendidos_agentes.lv_agentes.ListItems.Count
                   frmreporte_articulos_vendidos_agentes.lv_agentes.ListItems.Item(var_i).Selected = True
                   If frmreporte_articulos_vendidos_agentes.lv_agentes.selectedItem.SubItems(2) = "*" Then
                      var_contador = var_contador + 1
                   End If
               Next var_i
               If var_contador > 0 Then
                  Frmmenu2.StatusBar1.Panels(1).Text = "Procesando información espere un momento"
                  var_dia = CStr(Day(CDate(frmreporte_articulos_vendidos_articulos.txt_inicio)))
                  var_mes = CStr(Month(CDate(frmreporte_articulos_vendidos_articulos.txt_inicio)))
                  var_año = CStr(Year(CDate(frmreporte_articulos_vendidos_articulos.txt_inicio)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                  
               
                  var_dia = CStr(Day(frmreporte_articulos_vendidos_articulos.txt_fin))
                  var_mes = CStr(Month(frmreporte_articulos_vendidos_articulos.txt_fin))
                  var_año = CStr(Year(frmreporte_articulos_vendidos_articulos.txt_fin))
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
                  For var_i = 1 To frmreporte_articulos_vendidos_agentes.lv_agentes.ListItems.Count
                      frmreporte_articulos_vendidos_agentes.lv_agentes.ListItems.Item(var_i).Selected = True
                      If frmreporte_articulos_vendidos_agentes.lv_agentes.selectedItem.SubItems(2) = "*" Then
                         rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_AGENTES (INTE_TEM_CONSECUTIVO, VCHA_AGE_AGENTE_ID) VALUES (" + CStr(var_consecutivo) + ",'" + frmreporte_articulos_vendidos_agentes.lv_agentes.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                      End If
                  Next var_i
                  cnn.CommandTimeout = 360
                  For var_i = 1 To lv_articulos.ListItems.Count
                      lv_articulos.ListItems.Item(var_i).Selected = True
                      If lv_articulos.selectedItem.SubItems(2) = "*" Then
                         rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_ARTICULOS (INTE_TEM_CONSECUTIVO, VCHA_ART_ARTICULO_ID) VALUES (" + CStr(var_consecutivo) + ", '" + lv_articulos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
       
                      End If
                  Next var_i
                  rs.Open "EXEC SP_REPORTE_ARTICULOS_VENDIDOS_PERIODO " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ",3", cnn, adOpenDynamic, adLockOptimistic
                  Set reporte = appl.OpenReport(App.Path + "\rep_ventas_articulos_agentes_concentrado.rpt")
                  reporte.RecordSelectionFormula = "{VW_REPORTE_ARTICULOS_VENDIDOS_AGENTES_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\Reporte_ventas_articulos_agentes_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
                  rs.Open "delete from TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_TEMPORAL where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  rs.Open "delete from TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_AGENTES where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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
      End If
      If frmreporte_articulos_vendidos_articulos.txt_tipo_reporte = "Rutas" Then
         If IsDate(frmreporte_articulos_vendidos_articulos.txt_inicio) Then
            If IsDate(frmreporte_articulos_vendidos_articulos.txt_fin) Then
               var_contador = 0
               For var_i = 1 To frmreporte_articulos_vendidos_rutas.lv_rutas.ListItems.Count
                   frmreporte_articulos_vendidos_rutas.lv_rutas.ListItems.Item(var_i).Selected = True
                   If frmreporte_articulos_vendidos_rutas.lv_rutas.selectedItem.SubItems(2) = "*" Then
                      var_contador = var_contador + 1
                   End If
               Next var_i
               If var_contador > 0 Then
                  Frmmenu2.StatusBar1.Panels(1).Text = "Procesando información espere un momento"
                  var_dia = CStr(Day(CDate(frmreporte_articulos_vendidos_articulos.txt_inicio)))
                  var_mes = CStr(Month(CDate(frmreporte_articulos_vendidos_articulos.txt_inicio)))
                  var_año = CStr(Year(CDate(frmreporte_articulos_vendidos_articulos.txt_inicio)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
               
                  
                  var_dia = CStr(Day(frmreporte_articulos_vendidos_articulos.txt_fin))
                  var_mes = CStr(Month(frmreporte_articulos_vendidos_articulos.txt_fin))
                  var_año = CStr(Year(frmreporte_articulos_vendidos_articulos.txt_fin))
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
                  For var_i = 1 To frmreporte_articulos_vendidos_rutas.lv_rutas.ListItems.Count
                      frmreporte_articulos_vendidos_rutas.lv_rutas.ListItems.Item(var_i).Selected = True
                      If frmreporte_articulos_vendidos_rutas.lv_rutas.selectedItem.SubItems(2) = "*" Then
                         rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_RUTAS (INTE_TEM_CONSECUTIVO, VCHA_RUT_RUTA_ID) VALUES (" + CStr(var_consecutivo) + ",'" + frmreporte_articulos_vendidos_rutas.lv_rutas.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                      End If
                  Next var_i
                  cnn.CommandTimeout = 360
                  For var_i = 1 To lv_articulos.ListItems.Count
                      lv_articulos.ListItems.Item(var_i).Selected = True
                      If lv_articulos.selectedItem.SubItems(2) = "*" Then
                         rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_ARTICULOS (INTE_TEM_CONSECUTIVO, VCHA_ART_ARTICULO_ID) VALUES (" + CStr(var_consecutivo) + ", '" + lv_articulos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
       
                      End If
                  Next var_i
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
      End If
      If frmreporte_articulos_vendidos_articulos.txt_tipo_reporte = "Titulares" Then
         If IsDate(frmreporte_articulos_vendidos_articulos.txt_inicio) Then
            If IsDate(frmreporte_articulos_vendidos_articulos.txt_fin) Then
               var_contador = 0
               For var_i = 1 To frmreporte_articulos_vendidos_titulares.lv_titulares.ListItems.Count
                   frmreporte_articulos_vendidos_titulares.lv_titulares.ListItems.Item(var_i).Selected = True
                   If frmreporte_articulos_vendidos_titulares.lv_titulares.selectedItem.SubItems(2) = "*" Then
                      var_contador = var_contador + 1
                   End If
               Next var_i
               If var_contador > 0 Then
                  Frmmenu2.StatusBar1.Panels(1).Text = "Procesando información espere un momento"
                  var_dia = CStr(Day(CDate(frmreporte_articulos_vendidos_articulos.txt_inicio)))
                  var_mes = CStr(Month(CDate(frmreporte_articulos_vendidos_articulos.txt_inicio)))
                  var_año = CStr(Year(CDate(frmreporte_articulos_vendidos_articulos.txt_inicio)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
               
                  
                  var_dia = CStr(Day(frmreporte_articulos_vendidos_articulos.txt_fin))
                  var_mes = CStr(Month(frmreporte_articulos_vendidos_articulos.txt_fin))
                  var_año = CStr(Year(frmreporte_articulos_vendidos_articulos.txt_fin))
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
                  For var_i = 1 To frmreporte_articulos_vendidos_titulares.lv_titulares.ListItems.Count
                      frmreporte_articulos_vendidos_titulares.lv_titulares.ListItems.Item(var_i).Selected = True
                      If frmreporte_articulos_vendidos_titulares.lv_titulares.selectedItem.SubItems(2) = "*" Then
                         rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_TITULARES (INTE_TEM_CONSECUTIVO, VCHA_TIT_TITULAR_ID) VALUES (" + CStr(var_consecutivo) + ",'" + frmreporte_articulos_vendidos_titulares.lv_titulares.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                      End If
                  Next var_i
                  cnn.CommandTimeout = 360
                  For var_i = 1 To lv_articulos.ListItems.Count
                      lv_articulos.ListItems.Item(var_i).Selected = True
                      If lv_articulos.selectedItem.SubItems(2) = "*" Then
                         rs.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_ARTICULOS (INTE_TEM_CONSECUTIVO, VCHA_ART_ARTICULO_ID) VALUES (" + CStr(var_consecutivo) + ", '" + lv_articulos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
       
                      End If
                  Next var_i
                  rs.Open "EXEC SP_REPORTE_ARTICULOS_VENDIDOS_PERIODO " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ",5", cnn, adOpenDynamic, adLockOptimistic
                  Set reporte = appl.OpenReport(App.Path + "\rep_ventas_articulos_titulares_concentrado.rpt")
                  reporte.RecordSelectionFormula = "{VW_REPORTE_ARTICULOS_VENDIDOS_TITULARES_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\Reporte_ventas_articulos_titulares_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
                  rs.Open "delete from TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_TEMPORAL where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  rs.Open "delete from TB_TEMP_REPORTE_ARTICULOS_VENDIDOS_TITULARES where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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
      End If
   Else
      MsgBox "Debe de seleccionar algun artículo", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_invertir_Click()
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      If lv_articulos.selectedItem.SubItems(2) = "*" Then
         lv_articulos.selectedItem.SubItems(2) = ""
         lv_articulos.ListItems.Item(i).Bold = False
         lv_articulos.ListItems.Item(i).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_articulos.selectedItem.SubItems(2) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_articulos.selectedItem.Index
   If lv_articulos.selectedItem.SubItems(2) = "*" Then
      lv_articulos.selectedItem.SubItems(2) = ""
      lv_articulos.ListItems.Item(i).Bold = False
      lv_articulos.ListItems.Item(i).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_articulos.Refresh
   Else
      lv_articulos.selectedItem.SubItems(2) = "*"
      lv_articulos.ListItems.Item(i).Bold = True
      lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_articulos.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      lv_articulos.selectedItem.SubItems(2) = ""
      lv_articulos.ListItems.Item(i).Bold = False
      lv_articulos.ListItems.Item(i).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_articulos.Refresh
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_articulos.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_articulos.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_articulos.selectedItem.SubItems(2) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_articulos.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_articulos.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_todos_Click()
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      lv_articulos.selectedItem.SubItems(2) = "*"
      lv_articulos.ListItems.Item(i).Bold = True
      lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_articulos.Refresh
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 3000
   var_cadena = ""
   If var_cadena_reporte_articulos_catalogos <> "" Then
      var_cadena = "(" + var_cadena_reporte_articulos_catalogos + ")"
   End If
   
   
   If var_cadena_reporte_articulos_familias <> "" Then
      If var_cadena = "" Then
         var_cadena = "(" + var_cadena_reporte_articulos_familias + ")"
      Else
         var_cadena = var_cadena + " AND (" + var_cadena_reporte_articulos_familias + ")"
      End If
   End If
   
   If var_cadena_reporte_articulos_lineas <> "" Then
      If var_cadena = "" Then
         var_cadena = "(" + var_cadena_reporte_articulos_lineas + ")"
      Else
         var_cadena = var_cadena + " AND (" + var_cadena_reporte_articulos_lineas + ")"
      End If
   End If
   
   If var_cadena_reporte_articulos_tallas <> "" Then
      If var_cadena = "" Then
         var_cadena = "(" + var_cadena_reporte_articulos_tallas + ")"
      Else
         var_cadena = var_cadena + " AND (" + var_cadena_reporte_articulos_tallas + ")"
      End If
   End If
   If var_cadena <> "" Then
      rs.Open "select distinct VCHA_ART_ARTICULO_ID, vcha_ART_nombre_ESPAÑOL from TB_ARTICULOS WHERE " + var_cadena + " order by vcha_ART_nombre_ESPAÑOL ", cnn, adOpenDynamic, adLockOptimistic
      numero_items_ALMACENES = 0
      While Not rs.EOF
         If IsNull(rs!vcha_Art_articulo_id) Then
         Else
            Set list_item = lv_articulos.ListItems.Add(, , rs!vcha_Art_articulo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
            list_item.SubItems(2) = ""
         End If
         rs.MoveNext:
      Wend
      rs.Close
      If lv_articulos.ListItems.Count > 24 Then
         lv_articulos.ColumnHeaders(2).Width = 4000
      Else
         lv_articulos.ColumnHeaders(2).Width = 4250.26
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmreporte_articulos_vendidos_articulos.Enabled = True
End Sub

Private Sub lv_articulos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_articulos, ColumnHeader)
End Sub

Private Sub lv_articulos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_articulos.selectedItem.Index
      If lv_articulos.selectedItem.SubItems(2) = "*" Then
         lv_articulos.selectedItem.SubItems(2) = ""
         lv_articulos.ListItems.Item(i).Bold = False
         lv_articulos.ListItems.Item(i).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_articulos.Refresh
      Else
         lv_articulos.selectedItem.SubItems(2) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_articulos.Refresh
      End If
   End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
End Sub

Private Sub txt_busqueda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "select top 1 vcha_art_nombre_español from tb_articulos where vcha_art_nombre_español like '%" + Me.txt_busqueda + "%'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Call pro_busca_registro(lv_articulos, rs!vcha_art_nombre_español, False)
      Else
         MsgBox "No existen coincidencias", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_encontro = 0
         For var_i = 1 To lv_articulos.ListItems.Count
             lv_articulos.ListItems.Item(var_i).Selected = True
             If lv_articulos.selectedItem = Me.txt_codigo Then
                var_encontro = 1
             End If
         Next var_i
         If var_encontro = 1 Then
            Call pro_busca_registro(lv_articulos, rs!vcha_Art_articulo_id, False)
         Else
            Set list_item = lv_articulos.ListItems.Add(, , rs!vcha_Art_articulo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
            list_item.SubItems(2) = ""
         End If
      Else
         rsaux.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            txt_codigo = rsaux!vcha_Art_articulo_id
            rsaux2.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            var_encontro = 0
            For var_i = 1 To lv_articulos.ListItems.Count
                lv_articulos.ListItems.Item(var_i).Selected = True
                If lv_articulos.selectedItem = Me.txt_codigo Then
                   var_encontro = 1
                End If
            Next var_i
            If var_encontro = 1 Then
               Call pro_busca_registro(lv_articulos, rsaux!vcha_Art_articulo_id, False)
            Else
               Set list_item = lv_articulos.ListItems.Add(, , rsaux!vcha_Art_articulo_id)
               list_item.SubItems(1) = IIf(IsNull(rsaux2!vcha_art_nombre_español), "", rsaux2!vcha_art_nombre_español)
               list_item.SubItems(2) = ""
            End If
            rsaux2.Close
         Else
            MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         End If
         rsaux.Close
      End If
      rs.Close
   End If
End Sub
