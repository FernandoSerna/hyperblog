VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_ventas_lineas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de ventas por linea"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5325
      Picture         =   "frmreporte_ventas_lineas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmreporte_ventas_lineas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   45
      TabIndex        =   15
      Top             =   360
      Width           =   5685
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   60
      TabIndex        =   12
      Top             =   3345
      Width           =   5640
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   1095
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3375
         TabIndex        =   14
         Top             =   330
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   660
         TabIndex        =   13
         Top             =   330
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "  Lineas "
      Height          =   2880
      Left            =   60
      TabIndex        =   10
      Top             =   405
      Width           =   5625
      Begin VB.Frame Frame6 
         Height          =   120
         Left            =   30
         TabIndex        =   11
         Top             =   540
         Width           =   5565
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         Picture         =   "frmreporte_ventas_lineas.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   120
         Picture         =   "frmreporte_ventas_lineas.frx":0952
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         Picture         =   "frmreporte_ventas_lineas.frx":0A54
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   780
         Picture         =   "frmreporte_ventas_lineas.frx":0B26
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar (Enter)"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Picture         =   "frmreporte_ventas_lineas.frx":0D70
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   225
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_lineas 
         Height          =   2025
         Left            =   45
         TabIndex        =   2
         Top             =   690
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   3572
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
Attribute VB_Name = "frmreporte_ventas_lineas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_mes As Integer
Private Sub cmd_imprimir_Click()
   Dim pError As ADODB.Error
   'On Error GoTo salir:
   Dim var_consecutivo As Double
   Dim var_contador As Double
   Dim var_cadena As String
   Dim var_cadena_2 As String
   Dim var_contador_errores As Integer
   Dim var_dia As String
   Dim var_mes As String
   Dim var_año As String
   Dim var_año_anterior As String
   var_contador_errores = 0
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If var_empresa = "02" Or var_empresa = "03" Then
             
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            var_año_anterior = CStr(Year(CDate(txt_inicio)) - 1)
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + CStr(var_año) + "-" + CStr(var_mes) + "-" + CStr(var_dia) + "'}"
            var_fecha_inicio_anterior = "{d '" + CStr(var_año_anterior) + "-" + CStr(var_mes) + "-" + CStr(var_dia) + "'}"
            
            
            var_dia = CStr(Day(CDate(txt_fin)))
            var_mes = CStr(Month(CDate(txt_fin)))
            var_año = CStr(Year(CDate(txt_fin)))
            var_año_anterior = CStr(Year(CDate(txt_fin)) - 1)
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + CStr(var_año) + "-" + CStr(var_mes) + "-" + CStr(var_dia) + "'}"
            var_fecha_fin_anterior = "{d '" + CStr(var_año_anterior) + "-" + CStr(var_mes) + "-" + CStr(var_dia) + "'}"
            var_contador = 0
            var_cadena = ""
            var_cadena_2 = ""
            For var_i = 1 To lv_lineas.ListItems.Count
                lv_lineas.ListItems.Item(var_i).Selected = True
                If lv_lineas.selectedItem.SubItems(2) = "*" Then
                   var_contador = var_contador + 1
                End If
            Next var_i
            If var_contador > 0 Then
               cnn.CommandTimeout = 360
               cnn.BeginTrans
               rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_REPORTE_VENTAS_LINEA_COMPARATIVO", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rs.Close
               rs.Open "insert into TB_TEMP_REPORTE_VENTAS_LINEA_COMPARATIVO (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               var_n = Me.lv_lineas.ListItems.Count
               VAR_LINEAS = ""
               For var_i = 1 To var_n
                   lv_lineas.ListItems.Item(var_i).Selected = True
                   If lv_lineas.selectedItem.SubItems(2) = "*" Then
                      If VAR_LINEAS = "" Then
                         VAR_LINEAS = "'" + Trim(Me.lv_lineas.selectedItem) + "'"
                      Else
                         VAR_LINEAS = VAR_LINEAS + ",'" + Trim(Me.lv_lineas.selectedItem) + "'"
                      End If
                   End If
               Next var_i
               var_cadena = "insert into TB_TEMP_REPORTE_VENTAS_LINEA_COMPARATIVO (inte_tem_consecutivo, dtim_tem_fecha_inicio, dtim_Tem_fecha_fin, dtim_tem_fecha_inicio_anterior, dtim_tem_fecha_fin_anterior, vcha_Cli_clave_id, vcha_Art_articulo_id, floa_tem_cantidad_salida, floa_Tem_importe_salida, floa_tem_cantidad_devuelta, floa_tem_importe_devuelta, floa_Tem_cantidad_salida_anterior, floa_Tem_importe_salida_anterior, floa_tem_cantidad_devuelta_anterior, floa_tem_importe_devuelta_anterior )"
               var_cadena = var_cadena + " SELECT  " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + "," + var_fecha_inicio_anterior + "," + var_fecha_fin_anterior + ",dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_Articulos.VCHA_ART_ARTICULO_ID, SUM(dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD) As Cantidad,SUM((dbo.TB_SALIDAS.FLOA_SAL_PRECIO * (1 - dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1 / 100))* (1 - dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2 / 100) * dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD) AS IMPORTE,0,0,0,0,0,0 FROM dbo.TB_SALIDAS INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALIDAS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN "
               var_cadena = var_cadena + " dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_LINEAS ON dbo.TB_ARTICULOS.VCHA_LIN_LINEA_ID = dbo.TB_LINEAS.VCHA_LIN_LINEA_ID WHERE (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA < " + var_fecha_fin + "+1) AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_ARTICULOS.VCHA_LIN_LINEA_ID in (" + VAR_LINEAS + ")) AND (dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') GROUP BY dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID "
               rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               
               var_cadena = "insert into TB_TEMP_REPORTE_VENTAS_LINEA_COMPARATIVO (inte_tem_consecutivo, dtim_tem_fecha_inicio, dtim_Tem_fecha_fin, dtim_tem_fecha_inicio_anterior, dtim_tem_fecha_fin_anterior, vcha_Cli_clave_id, vcha_Art_articulo_id, floa_tem_cantidad_salida, floa_Tem_importe_salida, floa_tem_cantidad_devuelta, floa_tem_importe_devuelta, floa_Tem_cantidad_salida_anterior, floa_Tem_importe_salida_anterior, floa_tem_cantidad_devuelta_anterior, floa_tem_importe_devuelta_anterior )"
               var_cadena = var_cadena + " SELECT  " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + "," + var_fecha_inicio_anterior + "," + var_fecha_fin_anterior + ",dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_Articulos.VCHA_ART_ARTICULO_ID,0,0,0,0,SUM(dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD) As Cantidad,SUM((dbo.TB_SALIDAS.FLOA_SAL_PRECIO * (1 - dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1 / 100))* (1 - dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2 / 100) * dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD) AS IMPORTE,0,0 FROM dbo.TB_SALIDAS INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALIDAS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN "
               var_cadena = var_cadena + " dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_LINEAS ON dbo.TB_ARTICULOS.VCHA_LIN_LINEA_ID = dbo.TB_LINEAS.VCHA_LIN_LINEA_ID WHERE (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA >= " + var_fecha_inicio_anterior + ") AND (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA < " + var_fecha_fin_anterior + "+1) AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_ARTICULOS.VCHA_LIN_LINEA_ID in (" + VAR_LINEAS + ")) AND (dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') GROUP BY dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID "
               rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               
               
               
               var_cadena = "insert into TB_TEMP_REPORTE_VENTAS_LINEA_COMPARATIVO (inte_tem_consecutivo, dtim_tem_fecha_inicio, dtim_Tem_fecha_fin, dtim_tem_fecha_inicio_anterior, dtim_tem_fecha_fin_anterior, vcha_Cli_clave_id, vcha_Art_articulo_id, floa_tem_cantidad_salida, floa_Tem_importe_salida, floa_tem_cantidad_devuelta, floa_tem_importe_devuelta, floa_Tem_cantidad_salida_anterior, floa_Tem_importe_salida_anterior, floa_tem_cantidad_devuelta_anterior, floa_tem_importe_devuelta_anterior )"
               var_cadena = var_cadena + " SELECT     " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + "," + var_fecha_inicio_anterior + "," + var_fecha_fin_anterior + ", dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_CLI_CLAVE_ID, dbo.TB_DEVOLUCIONES.VCHA_ART_ARTICULO_ID, 0,0, SUM(dbo.TB_DEVOLUCIONES.FLOA_DEV_CANTIDAD) AS CANTIDAD,SUM(((dbo.TB_DEVOLUCIONES.FLOA_CDE_PRECIO * (1 - dbo.TB_DEVOLUCIONES.FLOA_CDE_DESCUENTO_1 / 100)) * (1 - dbo.TB_DEVOLUCIONES.FLOA_CDE_DESCUENTO_2 / 100)) * (1 - dbo.TB_DEVOLUCIONES.FLOA_CDE_DESCUENTO_3 / 100) * dbo.TB_DEVOLUCIONES.FLOA_DEV_CANTIDAD) AS IMPORTE,0,0,0,0 FROM dbo.TB_ARTICULOS INNER JOIN dbo.TB_DEVOLUCIONES ON dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID = dbo.TB_DEVOLUCIONES.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_DEVOLUCIONES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DEVOLUCIONES.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND "
               var_cadena = var_cadena + " dbo.TB_DEVOLUCIONES.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DEVOLUCIONES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND  dbo.TB_DEVOLUCIONES.INTE_EMO_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO WHERE (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA < " + var_fecha_fin + "+1) AND (dbo.TB_DEVOLUCIONES.VCHA_MOV_MOVIMIENTO_ID = 'CA') AND (dbo.TB_ARTICULOS.VCHA_LIN_LINEA_ID IN (" + VAR_LINEAS + ")) AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') GROUP BY dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_CLI_CLAVE_ID, dbo.TB_DEVOLUCIONES.VCHA_ART_ARTICULO_ID               "
               rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               
               
               var_cadena = "insert into TB_TEMP_REPORTE_VENTAS_LINEA_COMPARATIVO (inte_tem_consecutivo, dtim_tem_fecha_inicio, dtim_Tem_fecha_fin, dtim_tem_fecha_inicio_anterior, dtim_tem_fecha_fin_anterior, vcha_Cli_clave_id, vcha_Art_articulo_id, floa_tem_cantidad_salida, floa_Tem_importe_salida, floa_tem_cantidad_devuelta, floa_tem_importe_devuelta, floa_Tem_cantidad_salida_anterior, floa_Tem_importe_salida_anterior, floa_tem_cantidad_devuelta_anterior, floa_tem_importe_devuelta_anterior )"
               var_cadena = var_cadena + " SELECT     " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + "," + var_fecha_inicio_anterior + "," + var_fecha_fin_anterior + ", dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_CLI_CLAVE_ID, dbo.TB_DEVOLUCIONES.VCHA_ART_ARTICULO_ID, 0,0,0,0,0,0,SUM(dbo.TB_DEVOLUCIONES.FLOA_DEV_CANTIDAD) AS CANTIDAD,SUM(((dbo.TB_DEVOLUCIONES.FLOA_CDE_PRECIO * (1 - dbo.TB_DEVOLUCIONES.FLOA_CDE_DESCUENTO_1 / 100)) * (1 - dbo.TB_DEVOLUCIONES.FLOA_CDE_DESCUENTO_2 / 100)) * (1 - dbo.TB_DEVOLUCIONES.FLOA_CDE_DESCUENTO_3 / 100) * dbo.TB_DEVOLUCIONES.FLOA_DEV_CANTIDAD) AS IMPORTE FROM dbo.TB_ARTICULOS INNER JOIN dbo.TB_DEVOLUCIONES ON dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID = dbo.TB_DEVOLUCIONES.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_DEVOLUCIONES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DEVOLUCIONES.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND "
               var_cadena = var_cadena + " dbo.TB_DEVOLUCIONES.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DEVOLUCIONES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND  dbo.TB_DEVOLUCIONES.INTE_EMO_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO WHERE (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio_anterior + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA < " + var_fecha_fin_anterior + "+1) AND (dbo.TB_DEVOLUCIONES.VCHA_MOV_MOVIMIENTO_ID = 'CA') AND (dbo.TB_ARTICULOS.VCHA_LIN_LINEA_ID IN (" + VAR_LINEAS + ")) AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') GROUP BY dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_CLI_CLAVE_ID, dbo.TB_DEVOLUCIONES.VCHA_ART_ARTICULO_ID               "
               rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               
               rs.Open "delete from TB_TEMP_REPORTE_VENTAS_LINEA_COMPARATIVO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and dtim_Tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
               Set reporte = appl.OpenReport(App.Path + "\rep_ventas_lineas_comparativo.rpt")
               reporte.RecordSelectionFormula = "{VW_VENTAS_LINEAS_COMPARATIVO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_ventas_lineas_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
               rs.Open "delete from TB_TEMP_REPORTE_VENTAS_LINEA_COMPARATIVO where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            Else
               MsgBox "No se a seleccionado alguna linea", vbOKOnly, "ATENCION"
            End If
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
    Exit Sub
    
salir:
If Err.Number = -2147217871 Then
   var_si = MsgBox("El sistema a marcado tiempo de espera agotado, ¿Desea continuar?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      Resume
      var_contador_errores = var_contador_errores + 1
      If var_contador_errores = 4 Then
         MsgBox "A surgido un error al conectarce a la base de datos", vbOKOnly, "ATENCION"
         Exit Sub
      End If
   Else
      Exit Sub
   End If
  
Else
   MsgBox "A surgido un error", vbOKOnly, "ATENCION"
End If
    
End Sub

Private Sub cmd_invertir_Click()
   n = lv_lineas.ListItems.Count
   For i = 1 To n
      lv_lineas.ListItems.Item(i).Selected = True
      If lv_lineas.selectedItem.SubItems(2) = "*" Then
         lv_lineas.selectedItem.SubItems(2) = ""
         lv_lineas.ListItems.Item(i).Bold = False
         lv_lineas.ListItems.Item(i).ForeColor = &H80000012
         lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_lineas.selectedItem.SubItems(2) = "*"
         lv_lineas.ListItems.Item(i).Bold = True
         lv_lineas.ListItems.Item(i).ForeColor = &HFF0000
         lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i

End Sub

Private Sub cmd_marcar_Click()
   i = lv_lineas.selectedItem.Index
   If lv_lineas.selectedItem.SubItems(2) = "*" Then
      lv_lineas.selectedItem.SubItems(2) = ""
      lv_lineas.ListItems.Item(i).Bold = False
      lv_lineas.ListItems.Item(i).ForeColor = &H80000012
      lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_lineas.Refresh
   Else
      lv_lineas.selectedItem.SubItems(2) = "*"
      lv_lineas.ListItems.Item(i).Bold = True
      lv_lineas.ListItems.Item(i).ForeColor = &HFF0000
      lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_lineas.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_lineas.ListItems.Count
   For i = 1 To n
      lv_lineas.ListItems.Item(i).Selected = True
      lv_lineas.selectedItem.SubItems(2) = ""
      lv_lineas.ListItems.Item(i).Bold = False
      lv_lineas.ListItems.Item(i).ForeColor = &H80000012
      lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_lineas.Refresh
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_lineas.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_lineas.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_lineas.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_lineas.selectedItem.SubItems(2) = "*"
         lv_lineas.ListItems.Item(i).Bold = True
         lv_lineas.ListItems.Item(i).ForeColor = &HFF0000
         lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_lineas.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_lineas.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i

End Sub

Private Sub cmd_todos_Click()
   n = lv_lineas.ListItems.Count
   For i = 1 To n
      lv_lineas.ListItems.Item(i).Selected = True
      lv_lineas.selectedItem.SubItems(2) = "*"
      lv_lineas.ListItems.Item(i).Bold = True
      lv_lineas.ListItems.Item(i).ForeColor = &HFF0000
      lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_lineas.Refresh
End Sub



Private Sub Form_Load()
Dim dl As Long                                 ' Valor devuelto por la función API
Dim sAttributes As String                  ' Aributos
Dim sDriver As String                       ' Nombre del controlador
Dim sDescription As String                ' Descripción del DSN
Dim sDsnName As String                  ' Nombre del DSN

   cnn.Close
   cnn.Open var_conexion_string_distribucion

   Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
   Const vbAPINull As Long = 0&                         ' Puntero NULL

   ' se elimina
   Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema
   sDsnName = "DSN=sqlsistema"
   sDriver = "SQL Server"
   dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)

   'se crea
   sDsnName = "sqlsistema"
   sDescription = "sqlsistema"
   sDriver = "SQL Server"
   sAttributes = "DSN=" & sDsnName & Chr(0)
   sAttributes = sAttributes & "Server=" + parametros(0) & Chr$(0)
   sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
   sAttributes = sAttributes & "Database=" + var_bd_reportes & Chr(0)
   strAttributes = strAttributes & "UID=sa" & Chr$(0)
   strAttributes = strAttributes & "PWD=elia" & Chr$(0)
   dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
   
   
   
   
   
   var_cadena_seguridad = ""
   Top = 1500
   Left = 3200
   txt_inicio = Date
   txt_fin = Date
   'opt_linea = True
   rs.Open "SELECT DISTINCT TOP 100 PERCENT dbo.TB_LINEAS.VCHA_LIN_LINEA_ID, dbo.TB_LINEAS.VCHA_LIN_NOMBRE FROM dbo.TB_LINEAS INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_LINEAS.VCHA_LIN_LINEA_ID = dbo.TB_ARTICULOS.VCHA_LIN_LINEA_ID ORDER BY dbo.TB_LINEAS.VCHA_LIN_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      Set list_item = lv_lineas.ListItems.Add(, , rs!vcha_lin_linea_id)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_lin_NOMBRE), "", rs!VCHA_lin_NOMBRE)
      list_item.SubItems(2) = ""
      rs.MoveNext:
   Wend
   rs.Close
   If lv_lineas.ListItems.Count > 7 Then
      lv_lineas.ColumnHeaders(2).Width = 4220
   Else
      lv_lineas.ColumnHeaders(2).Width = 4499.71
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub lv_lineas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lineas, ColumnHeader)
End Sub

Private Sub lv_lineas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_lineas.selectedItem.Index
      If lv_lineas.selectedItem.SubItems(2) = "*" Then
         lv_lineas.selectedItem.SubItems(2) = ""
         lv_lineas.ListItems.Item(i).Bold = False
         lv_lineas.ListItems.Item(i).ForeColor = &H80000012
         lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_lineas.Refresh
      Else
         lv_lineas.selectedItem.SubItems(2) = "*"
         lv_lineas.ListItems.Item(i).Bold = True
         lv_lineas.ListItems.Item(i).ForeColor = &HFF0000
         lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_lineas.Refresh
      End If
   End If
End Sub

Private Sub mes_DateDblClick(ByVal DateDblClicked As Date)
   If var_mes = 1 Then
      txt_inicio = mes.Value
   End If
   If var_mes = 2 Then
      txt_fin = mes.Value
   End If
   mes.Visible = False
End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      mes.Visible = False
   End If
End Sub

Private Sub mes_LostFocus()
   mes.Visible = False
End Sub


Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes = CDate(Me.txt_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      Me.txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub
