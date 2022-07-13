VERSION 5.00
Begin VB.Form frmreporte_negado_produccion_resumen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen de negado de producción"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "SC"
      Height          =   315
      Left            =   405
      Picture         =   "frmreporte_negado_produccion_resumen.frx":0000
      TabIndex        =   8
      ToolTipText     =   "Reporte sin catálogos"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   90
      TabIndex        =   2
      Top             =   435
      Width           =   4245
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   315
         Width           =   1140
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   375
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2325
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3975
      Picture         =   "frmreporte_negado_produccion_resumen.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmreporte_negado_produccion_resumen.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   270
      Width           =   4395
   End
End
Attribute VB_Name = "frmreporte_negado_produccion_resumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_REPORTE_NEGADO_PRODUCCION_RESUMEN", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_REPORTE_NEGADO_PRODUCCION_RESUMEN (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin) + 1
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
            
            var_cadena = " INSERT INTO TB_TEMP_REPORTE_NEGADO_PRODUCCION_RESUMEN  ( [INTE_TEM_CONSECUTIVO], [DTIM_TEM_FECHA_INICIO], [DTIM_TEM_FECHA_FIN], [CHAR_TPE_TIPO_PEDIDO_ID], [VCHA_AGE_AGENTE_ID], [VCHA_TIT_TITULAR_ID],  [FLOA_ORS_CANTIDAD_PEDIDA], [FLOA_ORS_CANTIDAD_SURTIR], [FLOA_ORS_CANTIDAD_SURTIDA], [VCHA_AGE_NOMBRE], [VCHA_CAN_CANAL_VENTA_ID], [VCHA_CAN_NOMBRE], [VCHA_RUT_RUTA_ID], [VCHA_RUT_NOMBRE], [VCHA_TIT_NOMBRE], [VCHA_CLI_CLAVE_ID], [VCHA_CLI_NOMBRE])"
            var_cadena = var_cadena + " SELECT     " + CStr(var_consecutivo) + " ," + var_fecha_inicio + "," + var_fecha_fin + "- .000001, dbo.TB_ENC_ORDEN_SURTIDO.CHAR_TPE_TIPO_PEDIDO_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID, dbo.TB_ENC_ORDEN_SURTIDO.VCHA_TIT_TITULAR_ID, SUM(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_PEDIDA) AS FLOA_ORS_CANTIDAD_PEDIDA, SUM(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR) AS FLOA_ORS_CANTIDAD_SURTIR, SUM(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIDA) AS FLOA_ORS_CANTIDAD_SURTIDA, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CANALESVENTAS.VCHA_CAN_CANAL_VENTA_ID, dbo.TB_CANALESVENTAS.VCHA_CAN_NOMBRE, dbo.TB_RUTAS.VCHA_RUT_RUTA_ID, dbo.TB_RUTAS.VCHA_RUT_NOMBRE, dbo.TB_TITULARES.VCHA_TIT_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE FROM         dbo.TB_ENC_ORDEN_SURTIDO INNER JOIN dbo.TB_DET_ORDEN_SURTIDO ON dbo.TB_ENC_ORDEN_SURTIDO.VCHA_EMP_EMPRESA_ID = dbo.TB_DET_ORDEN_SURTIDO.VCHA_EMP_EMPRESA_ID AND"
            var_cadena = var_cadena + " dbo.TB_ENC_ORDEN_SURTIDO.VCHA_UOR_UNIDAD_ID = dbo.TB_DET_ORDEN_SURTIDO.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO INNER JOIN dbo.TB_ENCABEZADO_PEDIDOS ON DBo.TB_ENC_ORDEN_SURTIDO.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_PEDIDOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENC_ORDEN_SURTIDO.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO = dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO INNER JOIN dbo.TB_AGENTES ON dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID INNER JOIN dbo.TB_CANALESVENTAS ON dbo.TB_AGENTES.VCHA_CAN_CANAL_VENTA_ID = dbo.TB_CANALESVENTAS.VCHA_CAN_CANAL_VENTA_ID INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENC_ORDEN_SURTIDO.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID INNER JOIN DBO.TB_RUTAS ON dbo.TB_CLIENTES.VCHA_RUT_RUTA_ID = dbo.TB_RUTAS.VCHA_RUT_RUTA_ID INNER JOIN "
            var_cadena = var_cadena + " dbo.TB_TITULARES ON dbo.TB_ENC_ORDEN_SURTIDO.VCHA_TIT_TITULAR_ID = dbo.TB_TITULARES.VCHA_TIT_TITULAR_ID "
            var_cadena = var_cadena + " WHERE     (dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA BETWEEN " + var_fecha_inicio + " AND " + var_fecha_fin + " - .000001) GROUP BY dbo.TB_ENC_ORDEN_SURTIDO.CHAR_TPE_TIPO_PEDIDO_ID, dbo.TB_CANALESVENTAS.VCHA_CAN_NOMBRE, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_ENC_ORDEN_SURTIDO.VCHA_TIT_TITULAR_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CANALESVENTAS.VCHA_CAN_CANAL_VENTA_ID, dbo.TB_CANALESVENTAS.VCHA_CAN_NOMBRE, dbo.TB_RUTAS.VCHA_RUT_RUTA_ID, dbo.TB_RUTAS.VCHA_RUT_NOMBRE, dbo.TB_TITULARES.VCHA_TIT_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE ORDER BY dbo.TB_CANALESVENTAS.VCHA_CAN_CANAL_VENTA_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID,  dbo.TB_RUTAS.VCHA_RUT_RUTA_ID"
            cnn.CommandTimeout = 360
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            Text1 = var_cadena
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            rs.Open "delete from TB_TEMP_REPORTE_NEGADO_PRODUCCION_RESUMEN where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
            
            Set reporte = appl.OpenReport(App.Path + "\rep_negado_produccion_resumen.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_NEGADO_PRODUCCION_RESUMEN.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            For ntablas = 1 To reporte.Database.Tables.Count
               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\negado_produccion_resumen_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
            rs.Open "delete from TB_TEMP_REPORTE_DEVOLUCIONES_DETALLE_ARTICULO where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_REPORTE_NEGADO_PRODUCCION_RESUMEN", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_REPORTE_NEGADO_PRODUCCION_RESUMEN (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin) + 1
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
            
            var_cadena = " INSERT INTO TB_TEMP_REPORTE_NEGADO_PRODUCCION_RESUMEN  ( [INTE_TEM_CONSECUTIVO], [DTIM_TEM_FECHA_INICIO], [DTIM_TEM_FECHA_FIN], [CHAR_TPE_TIPO_PEDIDO_ID], [VCHA_AGE_AGENTE_ID], [VCHA_TIT_TITULAR_ID],  [FLOA_ORS_CANTIDAD_PEDIDA], [FLOA_ORS_CANTIDAD_SURTIR], [FLOA_ORS_CANTIDAD_SURTIDA], [VCHA_AGE_NOMBRE], [VCHA_CAN_CANAL_VENTA_ID], [VCHA_CAN_NOMBRE], [VCHA_RUT_RUTA_ID], [VCHA_RUT_NOMBRE], [VCHA_TIT_NOMBRE], [VCHA_CLI_CLAVE_ID], [VCHA_CLI_NOMBRE])"
            var_cadena = var_cadena + " SELECT TOP 100 PERCENT " + CStr(var_consecutivo) + " AS Expr1, " + var_fecha_inicio + " AS Expr2, " + var_fecha_fin + " - .000001 AS Expr3, dbo.TB_ENC_ORDEN_SURTIDO.CHAR_TPE_TIPO_PEDIDO_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID, dbo.TB_ENC_ORDEN_SURTIDO.VCHA_TIT_TITULAR_ID , SUM(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_PEDIDA) AS FLOA_ORS_CANTIDAD_PEDIDA, SUM(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR) AS FLOA_ORS_CANTIDAD_SURTIR, SUM(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIDA) AS FLOA_ORS_CANTIDAD_SURTIDA, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CANALESVENTAS.VCHA_CAN_CANAL_VENTA_ID, dbo.TB_CANALESVENTAS.VCHA_CAN_NOMBRE, dbo.TB_RUTAS.VCHA_RUT_RUTA_ID, dbo.TB_RUTAS.VCHA_RUT_NOMBRE, dbo.TB_TITULARES.VCHA_TIT_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE FROM  dbo.TB_ENC_ORDEN_SURTIDO INNER JOIN dbo.TB_DET_ORDEN_SURTIDO ON "
            var_cadena = var_cadena + " dbo.TB_ENC_ORDEN_SURTIDO.VCHA_EMP_EMPRESA_ID = dbo.TB_DET_ORDEN_SURTIDO.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENC_ORDEN_SURTIDO.VCHA_UOR_UNIDAD_ID = dbo.TB_DET_ORDEN_SURTIDO.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO INNER JOIN dbo.TB_ENCABEZADO_PEDIDOS ON dbo.TB_ENC_ORDEN_SURTIDO.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_PEDIDOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENC_ORDEN_SURTIDO.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO = dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO INNER JOIN dbo.TB_AGENTES ON dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID INNER JOIN dbo.TB_CANALESVENTAS ON dbo.TB_AGENTES.VCHA_CAN_CANAL_VENTA_ID = dbo.TB_CANALESVENTAS.VCHA_CAN_CANAL_VENTA_ID INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENC_ORDEN_SURTIDO.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID INNER JOIN "
            var_cadena = var_cadena + " dbo.TB_RUTAS ON dbo.TB_CLIENTES.VCHA_RUT_RUTA_ID = dbo.TB_RUTAS.VCHA_RUT_RUTA_ID INNER JOIN dbo.TB_TITULARES ON dbo.TB_ENC_ORDEN_SURTIDO.VCHA_TIT_TITULAR_ID = dbo.TB_TITULARES.VCHA_TIT_TITULAR_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA BETWEEN " + var_fecha_inicio + " AND " + var_fecha_fin + " - .000001) AND (dbo.TB_ARTICULOS.VCHA_LIN_LINEA_ID <> '90' or DBO.tb_articulos.vcha_lin_linea_id is null) "
            var_cadena = var_cadena + " GROUP BY dbo.TB_ENC_ORDEN_SURTIDO.CHAR_TPE_TIPO_PEDIDO_ID, dbo.TB_CANALESVENTAS.VCHA_CAN_NOMBRE, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_ENC_ORDEN_SURTIDO.VCHA_TIT_TITULAR_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CANALESVENTAS.VCHA_CAN_CANAL_VENTA_ID, dbo.TB_CANALESVENTAS.VCHA_CAN_NOMBRE, dbo.TB_RUTAS.VCHA_RUT_RUTA_ID, dbo.TB_RUTAS.VCHA_RUT_NOMBRE, dbo.TB_TITULARES.VCHA_TIT_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, "
            var_cadena = var_cadena + " dbo.TB_CLIENTES.VCHA_CLI_NOMBRE ORDER BY dbo.TB_CANALESVENTAS.VCHA_CAN_CANAL_VENTA_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID, dbo.TB_RUTAS.VCHA_RUT_RUTA_ID "
            cnn.CommandTimeout = 360
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            Text1 = var_cadena
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            rs.Open "delete from TB_TEMP_REPORTE_NEGADO_PRODUCCION_RESUMEN where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
            
            Set reporte = appl.OpenReport(App.Path + "\rep_negado_produccion_resumen.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_NEGADO_PRODUCCION_RESUMEN.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            For ntablas = 1 To reporte.Database.Tables.Count
               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\negado_produccion_resumen_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
            rs.Open "delete from TB_TEMP_REPORTE_DEVOLUCIONES_DETALLE_ARTICULO where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_fin_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar la fecha"
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes = CDate(Me.txt_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_fin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_fin_LostFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = ""
End Sub

Private Sub txt_inicio_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar la fecha"
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

Private Sub txt_inicio_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_inicio_LostFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = ""
End Sub



