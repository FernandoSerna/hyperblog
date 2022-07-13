VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_ubicaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Ubicaciones"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Height          =   315
      Left            =   135
      Picture         =   "frmreporte_ubicaciones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Caption         =   " Ubicación "
      Height          =   1005
      Left            =   135
      TabIndex        =   19
      Top             =   6195
      Width           =   5880
      Begin VB.TextBox txt_lugar 
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
         Height          =   420
         Left            =   4785
         MaxLength       =   3
         TabIndex        =   15
         Text            =   "000"
         Top             =   345
         Width           =   615
      End
      Begin VB.TextBox txt_nivel 
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
         Height          =   420
         Left            =   3450
         MaxLength       =   3
         TabIndex        =   14
         Text            =   "000"
         Top             =   345
         Width           =   615
      End
      Begin VB.TextBox txt_modulo 
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
         Height          =   420
         Left            =   2235
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "000"
         Top             =   345
         Width           =   615
      End
      Begin VB.TextBox txt_pasillo 
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
         Height          =   420
         Left            =   765
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "000"
         Top             =   345
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Lugar:"
         Height          =   195
         Left            =   4275
         TabIndex        =   23
         Top             =   458
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nivel:"
         Height          =   195
         Left            =   3000
         TabIndex        =   22
         Top             =   458
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Modulo:"
         Height          =   195
         Left            =   1620
         TabIndex        =   21
         Top             =   458
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pasillo:"
         Height          =   195
         Left            =   210
         TabIndex        =   20
         Top             =   458
         Width           =   495
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmreporte_ubicaciones.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5655
      Picture         =   "frmreporte_ubicaciones.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   15
      TabIndex        =   18
      Top             =   345
      Width           =   6090
   End
   Begin VB.Frame Frame1 
      Caption         =   " Artículos "
      Height          =   5685
      Left            =   135
      TabIndex        =   16
      Top             =   450
      Width           =   5880
      Begin VB.TextBox txt_codigo 
         Height          =   315
         Left            =   1020
         TabIndex        =   10
         Top             =   975
         Width           =   2640
      End
      Begin VB.CheckBox chk_todos 
         Caption         =   "Todos los artículos"
         Height          =   255
         Left            =   3900
         TabIndex        =   8
         Top             =   270
         Width           =   1815
      End
      Begin VB.TextBox txt_busqueda 
         Height          =   315
         Left            =   1020
         TabIndex        =   9
         Top             =   630
         Width           =   4770
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1455
         Picture         =   "frmreporte_ubicaciones.frx":083E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   255
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   795
         Picture         =   "frmreporte_ubicaciones.frx":0A54
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar (Enter)"
         Top             =   255
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1125
         Picture         =   "frmreporte_ubicaciones.frx":0C9E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   255
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   135
         Picture         =   "frmreporte_ubicaciones.frx":0D70
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   255
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   465
         Picture         =   "frmreporte_ubicaciones.frx":0E72
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   255
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   4230
         Left            =   60
         TabIndex        =   11
         Top             =   1395
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   7461
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
         TabIndex        =   24
         Top             =   1005
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   660
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmreporte_ubicaciones"
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
                  cnn.CommandTimeout = 360
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

Private Sub cmd_imprimir_Click()
   Dim var_pasillo As String
   var_pasillo = ""
   var_consecutivo = 0
   If Trim(Me.txt_pasillo) <> "" Then
      var_pasillo = Trim(Me.txt_pasillo)
   End If
   If Trim(Me.txt_modulo) <> "" Then
      var_pasillo = var_pasillo + Trim(Me.txt_modulo)
   End If
   If Trim(Me.txt_nivel) <> "" Then
      var_pasillo = var_pasillo + Trim(Me.txt_nivel)
   End If
   If Trim(Me.txt_lugar) <> "" Then
      var_pasillo = var_pasillo + Trim(Me.txt_lugar)
   End If
   If Me.chk_todos = 1 Then
      cnn.BeginTrans
      rs.Open "select max(inte_tem_consecutivo) from tb_temp_reporte_ubicaciones", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
      Else
         var_consecutivo = 0
      End If
      var_consecutivo = var_consecutivo + 1
      rs.Close
      rs.Open "Insert into tb_temp_reporte_ubicaciones (INTE_Tem_CONSECUTIVO) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
      cnn.CommitTrans
      rs.Open "Insert into tb_temp_reporte_ubicaciones (INTE_Tem_CONSECUTIVO, vcha_Art_articulo_id) select " + CStr(var_consecutivo) + ",vcha_Art_articulo_id from tb_Articulos ", cnn, adOpenDynamic, adLockOptimistic
   Else
      If lv_articulos.ListItems.Count > 0 Then
         var_cadena = ""
         For var_j = 1 To lv_articulos.ListItems.Count
             lv_articulos.ListItems.Item(var_j).Selected = True
             If lv_articulos.selectedItem.SubItems(2) = "*" Then
                If var_cadena = "" Then
                   var_cadena = " VCHA_ART_ARTICULO_ID = '" + lv_articulos.selectedItem + "'"
                Else
                   var_cadena = var_cadena + " OR VCHA_ART_ARTICULO_ID = '" + lv_articulos.selectedItem + "'"
                End If
             End If
         Next var_j
         If var_cadena <> "" Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) from tb_temp_reporte_ubicaciones", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "Insert into tb_temp_reporte_ubicaciones (INTE_Tem_CONSECUTIVO) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            rs.Open "Insert into tb_temp_reporte_ubicaciones (INTE_Tem_CONSECUTIVO, vcha_Art_articulo_id) select " + CStr(var_consecutivo) + ",vcha_Art_articulo_id from tb_Articulos WHERE " + var_cadena, cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "No se a seleccionado algun artículos", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a seleccionado algun artículo", vbOKOnly, "ATENCION"
      End If
   End If
   
   rs.Open "SELECT * FROM TB_TEMP_REPORTE_UBICACIONES WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Set reporte = appl.OpenReport(App.Path + "\rep_ubicaciones.rpt")
      If var_pasillo <> "" Then
         If var_empresa <> "31" Then
            If Len(var_pasillo) = 1 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},1) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_2},1) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_3},1) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 2 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},2) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_2},2) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_3},2) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 3 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},3) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_2},3) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_3},3) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 4 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},4) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_2},4) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_3},4) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 5 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},5) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_2},5) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_3},5) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 6 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},6) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_2},6) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_3},6) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 7 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},7) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_2},7) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_3},7) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 8 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},8) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_2},8) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_3},8) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 9 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},9) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_2},9) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_3},9) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 10 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},10) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_2},10) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_3},10) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 11 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},11) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_2},11) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_3},11) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 12 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},12) = '" + var_pasillo + "')"
            End If
         Else
            If Len(var_pasillo) = 1 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},1) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 2 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},2) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 3 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},3) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 4 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},4) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 5 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},5) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 6 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},6) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 7 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},7) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 8 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},8) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 9 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},9) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 10 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},10) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 11 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},11) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 12 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},12) = '" + var_pasillo + "')"
            End If
         End If
      Else
         reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo)
      End If
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de ubicaciones"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         Set reporte = appl.OpenReport(App.Path + "\rep_ubicaciones.rpt")
         If var_pasillo <> "" Then
            If Len(var_pasillo) = 3 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},3) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_2},3) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_3},3) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 6 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},6) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_2},6) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_3},6) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 9 Then
               reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},9) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_2},9) = '" + var_pasillo + "' OR left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_3},9) = '" + var_pasillo + "')"
            End If
            If Len(var_pasillo) = 12 Then
              reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND (Left({VW_REPORTE_UBICACIONES.VCHA_UBI_UBICACION_1},12) = '" + var_pasillo + "')"
            End If
         Else
            reporte.RecordSelectionFormula = "{VW_reporte_ubicaciones.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo)
         End If
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reportessid\repote_ubicaciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
         MsgBox "Se a terminado de guardar el archivo " + archivo
      End If
   Else
      MsgBox "No existen respuesta para el reporte", vbOKOnly, "ATENCION"
   End If
   rs.Close
   rs.Open "delete from tb_temp_reporte_ubicaciones where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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

Private Sub cmd_salir_Click()
   Unload Me
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

Private Sub Command1_Click()
   Me.txt_pasillo = ""
   Me.txt_busqueda = ""
   Me.txt_lugar = ""
   Me.txt_modulo = ""
   Me.txt_nivel = ""
   Me.txt_codigo = ""
   Me.lv_articulos.ListItems.Clear
   Me.chk_todos = 0
   Me.txt_busqueda.SetFocus
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 3000
   var_cadena = ""
   Me.txt_pasillo = ""
   Me.txt_modulo = ""
   Me.txt_nivel = ""
   Me.txt_lugar = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
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
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(Me.txt_busqueda) <> "" Then
         var_cadena = ""
         var_like_1 = ""
         var_like_2 = ""
         var_like_3 = ""
         var_like_4 = ""
         var_like_5 = ""
         var_like_6 = ""
         var_like_7 = ""
         var_j = 1
         For var_i = 1 To Len(Me.txt_busqueda)
             If Mid(Me.txt_busqueda, var_i, 1) <> " " Then
                If var_j = 1 Then
                   var_like_1 = var_like_1 + Mid(Me.txt_busqueda, var_i, 1)
                End If
                If var_j = 2 Then
                   var_like_2 = var_like_2 + Mid(Me.txt_busqueda, var_i, 1)
                End If
                If var_j = 3 Then
                   var_like_3 = var_like_3 + Mid(Me.txt_busqueda, var_i, 1)
                End If
                If var_j = 4 Then
                   var_like_4 = var_like_4 + Mid(Me.txt_busqueda, var_i, 1)
                End If
                If var_j = 5 Then
                   var_like_5 = var_like_5 + Mid(Me.txt_busqueda, var_i, 1)
                End If
                If var_j = 6 Then
                   var_like_6 = var_like_6 + Mid(Me.txt_busqueda, var_i, 1)
                End If
                If var_j >= 7 Then
                   var_like_7 = var_like_7 + Mid(Me.txt_busqueda, var_i, 1)
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
      Me.lv_articulos.ListItems.Clear
      If Trim(var_cadena) <> "" Then
         var_cadena = "SELECT * FROM tb_Articulos WHERE " + var_cadena
         rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = lv_articulos.ListItems.Add(, , rsaux!vcha_Art_Articulo_id)
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_Art_nombre_español), "", rsaux!vcha_Art_nombre_español)
               list_item.SubItems(2) = ""
               rsaux.MoveNext
         Wend
         rsaux.Close
         If Me.lv_articulos.ListItems.Count > 0 Then
            Me.lv_articulos.SetFocus
         End If
         If lv_articulos.ListItems.Count > 11 Then
            lv_articulos.ColumnHeaders(2).Width = 4000
         Else
            lv_articulos.ColumnHeaders(2).Width = 4250.26
         End If
      Else
         MsgBox "No existen coincidencias", vbOKOnly, "ATENCION"
      End If
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(Me.txt_codigo) <> "" Then
         var_cadena = ""
         var_like_1 = ""
         var_like_2 = ""
         var_like_3 = ""
         var_like_4 = ""
         var_like_5 = ""
         var_like_6 = ""
         var_like_7 = ""
         var_j = 1
         For var_i = 1 To Len(Me.txt_codigo)
             If Mid(Me.txt_codigo, var_i, 1) <> " " Then
                If var_j = 1 Then
                   var_like_1 = var_like_1 + Mid(Me.txt_codigo, var_i, 1)
                End If
                If var_j = 2 Then
                   var_like_2 = var_like_2 + Mid(Me.txt_codigo, var_i, 1)
                End If
                If var_j = 3 Then
                   var_like_3 = var_like_3 + Mid(Me.txt_codigo, var_i, 1)
                End If
                If var_j = 4 Then
                   var_like_4 = var_like_4 + Mid(Me.txt_codigo, var_i, 1)
                End If
                If var_j = 5 Then
                   var_like_5 = var_like_5 + Mid(Me.txt_codigo, var_i, 1)
                End If
                If var_j = 6 Then
                   var_like_6 = var_like_6 + Mid(Me.txt_codigo, var_i, 1)
                End If
                If var_j >= 7 Then
                   var_like_7 = var_like_7 + Mid(Me.txt_codigo, var_i, 1)
                End If
             Else
                var_j = var_j + 1
             End If
         Next var_i
      End If
      If Trim(var_like_1) <> "" Then
         var_cadena = " vcha_Art_articulo_id like  '%" + var_like_1 + "%'"
      End If
      If Trim(var_like_2) <> "" Then
         var_cadena = var_cadena + " or  vcha_Art_articulo_id like  '%" + var_like_2 + "%'"
      End If
      If Trim(var_like_3) <> "" Then
         var_cadena = var_cadena + " or  vcha_Art_articulo_id like  '%" + var_like_3 + "%'"
      End If
      If Trim(var_like_4) <> "" Then
         var_cadena = var_cadena + " or  vcha_Art_articulo_id like  '%" + var_like_4 + "%'"
      End If
      If Trim(var_like_5) <> "" Then
         var_cadena = var_cadena + " or  vcha_Art_articulo_id like  '%" + var_like_5 + "%'"
      End If
      If Trim(var_like_6) <> "" Then
         var_cadena = var_cadena + " or  vcha_Art_articulo_id like  '%" + var_like_6 + "%'"
      End If
      If Trim(var_like_7) <> "" Then
         var_cadena = var_cadena + " or  vcha_Art_articulo_id like  '%" + var_like_7 + "%'"
      End If
      Me.lv_articulos.ListItems.Clear
      If Trim(var_cadena) <> "" Then
         var_cadena = "SELECT * FROM tb_Articulos WHERE " + var_cadena
         rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = lv_articulos.ListItems.Add(, , rsaux!vcha_Art_Articulo_id)
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_Art_nombre_español), "", rsaux!vcha_Art_nombre_español)
               list_item.SubItems(2) = ""
               rsaux.MoveNext
         Wend
         rsaux.Close
         If Me.lv_articulos.ListItems.Count > 0 Then
            Me.lv_articulos.SetFocus
         End If
         If lv_articulos.ListItems.Count > 11 Then
            lv_articulos.ColumnHeaders(2).Width = 4000
         Else
            lv_articulos.ColumnHeaders(2).Width = 4250.26
         End If
      Else
         MsgBox "No existen coincidencias", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_lugar_KeyPress(KeyAscii As Integer)
   'Select Case KeyAscii
   'Case 48 To 57, 52, 13, 8, 46
   'Case Else
   '    KeyAscii = 0
   'End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_lugar_LostFocus()
'   If Len(Trim(Me.txt_lugar)) < 3 Then
'      If Len(Trim(Me.txt_lugar)) = 1 Then
'         Me.txt_lugar = "00" + Trim(Me.txt_lugar)
'      End If
'      If Len(Trim(Me.txt_nivel)) = 2 Then
'         Me.txt_lugar = "0" + Trim(Me.txt_lugar)
'      End If
'   End If
End Sub

Private Sub txt_modulo_Change()
   Me.txt_nivel = ""
End Sub

Private Sub txt_modulo_KeyPress(KeyAscii As Integer)
   'Select Case KeyAscii
   'Case 48 To 57, 52, 13, 8, 46
   'Case Else
   '    KeyAscii = 0
   'End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_modulo_LostFocus()
'   If Len(Trim(Me.txt_modulo)) < 3 Then
'      If Len(Trim(Me.txt_modulo)) = 1 Then
'         Me.txt_modulo = "00" + Trim(Me.txt_modulo)
'      End If
'      If Len(Trim(Me.txt_modulo)) = 2 Then
'         Me.txt_modulo = "0" + Trim(Me.txt_modulo)
'      End If
'   End If
End Sub

Private Sub txt_nivel_Change()
   Me.txt_lugar = ""
End Sub

Private Sub txt_nivel_KeyPress(KeyAscii As Integer)
   'Select Case KeyAscii
   'Case 48 To 57, 52, 13, 8, 46
   'Case Else
   '    KeyAscii = 0
   'End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nivel_LostFocus()
'   If Len(Trim(Me.txt_nivel)) < 3 Then
'      If Len(Trim(Me.txt_nivel)) = 1 Then
'         Me.txt_nivel = "00" + Trim(Me.txt_nivel)
'      End If
'      If Len(Trim(Me.txt_nivel)) = 2 Then
'         Me.txt_nivel = "0" + Trim(Me.txt_nivel)
'      End If
'   End If
End Sub

Private Sub txt_pasillo_Change()
   Me.txt_modulo = ""
   
End Sub

Private Sub txt_pasillo_KeyPress(KeyAscii As Integer)
   'Select Case KeyAscii
   'Case 48 To 57, 52, 13, 8, 46
   'Case Else
   '    KeyAscii = 0
   'End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_pasillo_LostFocus()
   'If Len(Trim(Me.txt_pasillo)) < 3 Then
   '   If Len(Trim(Me.txt_pasillo)) = 1 Then
   '      Me.txt_pasillo = "00" + Trim(Me.txt_pasillo)
   '   End If
   '   If Len(Trim(Me.txt_pasillo)) = 2 Then
   '      Me.txt_pasillo = "0" + Trim(Me.txt_pasillo)
   '   End If
   'End If
End Sub
