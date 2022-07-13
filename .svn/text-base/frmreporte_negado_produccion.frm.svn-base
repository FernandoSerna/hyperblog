VERSION 5.00
Begin VB.Form frmreporte_negado_produccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Negado por producción"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Height          =   840
      Left            =   90
      TabIndex        =   2
      Top             =   435
      Width           =   2670
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmreporte_negado_produccion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2415
      Picture         =   "frmreporte_negado_produccion.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   285
      Width           =   2655
   End
End
Attribute VB_Name = "frmreporte_negado_produccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   If IsDate(Me.txt_inicio) Then
      cnn.CommandTimeout = 360
      cnn.BeginTrans
      rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_REPORTE_negado_PRODUCCION", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
      Else
         var_consecutivo = 0
      End If
      var_consecutivo = var_consecutivo + 1
      rsaux.Open "INSERT INTO TB_TEMP_REPORTE_negado_PRODUCCION (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
      rs.Close
      cnn.CommitTrans
      var_año = CStr(Year(CDate(Me.txt_inicio)))
      var_mes = CStr(Month(CDate(Me.txt_inicio)))
      var_dia = CStr(Day(CDate(Me.txt_inicio)))
      If Len(var_año) = 2 Then
         var_año_str = "20" + var_año
      Else
         If Len(var_año) = 3 Then
            var_año = "2" + var_año
         Else
            var_año_str = var_año
         End If
      End If
      If Len(var_mes) = 1 Then
         var_mes_str = "0" + var_mes
      Else
         var_mes_str = var_mes
      End If
      If Len(var_dia) = 1 Then
         var_dia_str = "0" + var_dia
      Else
         var_dia_str = var_dia
      End If
      VAR_FECHA_STR = var_año_str + "-" + var_mes_str + "-" + var_dia_str
      VAR_FECHA_STR = "{d '" + VAR_FECHA_STR + "'}"
            
            
      var_cadena = "insert into TB_TEMP_REPORTE_NEGADO_PRODUCCION (inte_tem_consecutivo, CHAR_TPE_TIPO_PEDIDO, VCHA_CAN_NOMBRE, VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE, INTE_PED_NUMERO, INTE_ORS_ORDEN_SURTIDO, DTIM_ORS_FECHA_CARGA, VCHA_TIT_TITULAR_ID, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_ESPAÑOL, VCHA_LIN_NOMBRE, VCHA_ART_CATALOGO_VIGENTE, FLOA_ORS_CANTIDAD_PEDIDA, FLOA_ORS_CANTIDAD_SURTIR, FLOA_ORS_CANTIDAD_SURTIDA, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, FLOA_SAL_PRECIO , FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, FLOA_ORS_PRECIO, FLOA_ORS_DESCUENTO_1, FLOA_ORS_DESCUENTO_2, VCHA_ZON_ZONA_ID, VCHA_ZON_DESCRIPCION, VCHA_RUT_RUTA_ID, VCHA_RUT_NOMBRE, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GAC_NOMBRE, VCHA_TIT_NOMBRE, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_ESB_NOMBRE)"
      var_cadena = var_cadena + " SELECT  TOP 100 PERCENT " + CStr(var_consecutivo) + ", dbo.TB_ENC_ORDEN_SURTIDO.CHAR_TPE_TIPO_PEDIDO_ID, dbo.TB_CANALESVENTAS.VCHA_CAN_NOMBRE, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO, dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO, dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA, dbo.TB_ENC_ORDEN_SURTIDO.VCHA_TIT_TITULAR_ID, dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_LINEAS.VCHA_LIN_NOMBRE, dbo.TB_ARTICULOS.VCHA_ART_CATALOGO_VIGENTE, dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_PEDIDA, dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR, dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIDA, dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO, dbo.TB_SALIDAS.VCHA_SER_SERIE_ID, dbo.TB_SALIDAS.INTE_CAR_NUMERO, "
      var_cadena = var_cadena + " dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2, dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_PRECIO, dbo.TB_ENC_ORDEN_SURTIDO.FLOA_ORS_DESCUENTO_1, dbo.TB_ENC_ORDEN_SURTIDO.FLOA_ORS_DESCUENTO_2, dbo.VW_CLIENTES.VCHA_ZON_ZONA_ID, dbo.VW_CLIENTES.VCHA_ZON_DESCRIPCION, dbo.VW_CLIENTES.VCHA_RUT_RUTA_ID, dbo.VW_CLIENTES.VCHA_RUT_NOMBRE, dbo.VW_CLIENTES.VCHA_GAC_GRUPO_ACTUAL_ID, dbo.VW_CLIENTES.VCHA_GAC_NOMBRE, dbo.VW_CLIENTES.VCHA_TIT_NOMBRE, dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.VW_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID, dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_NOMBRE FROM dbo.TB_ESTABLECIMIENTOS INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO INNER JOIN dbo.TB_DET_ORDEN_SURTIDO ON dbo.TB_ENC_ORDEN_SURTIDO.VCHA_EMP_EMPRESA_ID = dbo.TB_DET_ORDEN_SURTIDO.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENC_ORDEN_SURTIDO.VCHA_UOR_UNIDAD_ID = dbo.TB_DET_ORDEN_SURTIDO.VCHA_UOR_UNIDAD_ID AND "
      var_cadena = var_cadena + " dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_ENCABEZADO_PEDIDOS ON dbo.TB_ENC_ORDEN_SURTIDO.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_PEDIDOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENC_ORDEN_SURTIDO.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO = dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO INNER JOIN dbo.TB_AGENTES ON dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID INNER JOIN dbo.TB_CANALESVENTAS ON dbo.TB_AGENTES.VCHA_CAN_CANAL_VENTA_ID = dbo.TB_CANALESVENTAS.VCHA_CAN_CANAL_VENTA_ID INNER JOIN dbo.TB_LINEAS ON dbo.TB_ARTICULOS.VCHA_LIN_LINEA_ID = dbo.TB_LINEAS.VCHA_LIN_LINEA_ID ON "
      var_cadena = var_cadena + " dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID = dbo.TB_ENC_ORDEN_SURTIDO.VCHA_ESB_ESTABLECIMIENTO_ID INNER JOIN dbo.VW_CLIENTES ON dbo.TB_ENC_ORDEN_SURTIDO.VCHA_CLI_CLAVE_ID = dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID LEFT OUTER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_SALIDAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO ON dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN AND dbo.TB_Articulos.VCHA_ART_ARTICULO_ID = dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID"
      var_cadena = var_cadena + " WHERE     (dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA BETWEEN " + VAR_FECHA_STR + " AND " + VAR_FECHA_STR + " + 1 - .000001) ORDER BY dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO, dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID"
      'MsgBox var_cadena
      rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
      rs.Open "delete from TB_TEMP_REPORTE_NEGADO_PRODUCCION where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and CHAR_TPE_TIPO_PEDIDO is null", cnn, adOpenDynamic, adLockOptimistic
      Set reporte = appl.OpenReport(App.Path + "\REP_NEGADO_PRODUCCION.rpt")
      reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_NEGADO_PRODUCCION.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      reporte.ExportOptions.FormatType = crEFTExcel80
      reporte.ExportOptions.DestinationType = crEDTDiskFile
      archivo = "c:\reportessid\Reporte_negado_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
      reporte.ExportOptions.DiskFileName = archivo
      reporte.Export False
      Set reporte = Nothing
      MsgBox "Se a generado el archivo " + archivo, vbOKOnly, "ATENCION"
      rs.Open "delete from TB_TEMP_REPORTE_NEGADO_PRODUCCION where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 4000
   txt_inicio = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_packing_list)
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
