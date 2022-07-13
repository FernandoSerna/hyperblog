VERSION 5.00
Begin VB.Form frmreporte_articulos_nuevos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de artículos nuevos"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmreporte_articulos_nuevos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4020
      Picture         =   "frmreporte_articulos_nuevos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   105
      TabIndex        =   0
      Top             =   435
      Width           =   4245
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   1140
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2325
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   375
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   15
      TabIndex        =   7
      Top             =   270
      Width           =   4395
   End
End
Attribute VB_Name = "frmreporte_articulos_nuevos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_consecutivo As Double
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   
   'On Error GoTo salir:
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            
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
             
             var_dia = CStr(Day(CDate(txt_fin)))
             var_mes = CStr(Month(CDate(txt_fin)))
             var_año = CStr(Year(CDate(txt_fin)))
             If Len(Trim(var_dia)) = 1 Then
                var_dia = "0" + var_dia
             End If
             If Len(Trim(var_mes)) = 1 Then
                var_mes = "0" + var_mes
             End If
             var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
             
             
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
             
             
             
            cnn.BeginTrans
            rs.Open "select max(isnull(INTE_TEM_CONSECUTIVO,0)) as numero from TB_TEMP_REPORTE_ARTICULOS_NUEVOS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rsaux.Open "INSERT INTO TB_TEMP_REPORTE_ARTICULOS_NUEVOS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTiM_TEM_FECHA_FIN) VALUES (" + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            cnn.CommandTimeout = 360
            var_cadena = " INSERT INTO TB_TEMP_REPORTE_ARTICULOS_NUEVOS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, INTE_PED_PEDIDO, INTE_CAR_NUMERO, DTIM_CAR_FECHA, VCHA_SER_SERIE_ID, VCHA_CAR_DOCUMENTO,  VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, vcha_Esb_Establecimiento_id) "
            
            'var_cadena = var_cadena + " SELECT  " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + ", tb_encabezado_Cartera.vcha_gac_grupo_actual_id, tb_encabezado_cartera.vcha_tit_titular_id, tb_encabezado_Cartera.vcha_cli_clave_id, dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO,  dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO,dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID,  dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1 , dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2,dbo.TB_ENCABEZADO_CARTERA.vcha_Esb_establecimiento_iD "
            'var_cadena = var_cadena + " FROM dbo.TB_ENCABEZADO_PEDIDOS INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO ON dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO = dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN INNER JOIN "
            'var_cadena = var_cadena + " dbo.TB_SALIDAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID AND dbo.TB_SALIDAS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALIDAS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO "
            'var_cadena = var_cadena + " WHERE (dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_ARTICULOS_NUEVOS = 1) and dtim_Car_fecha >= " + var_fecha_inicio + " and dtim_Car_Fecha <= " + var_fecha_fin + " + 1 -.000001"
            
            
            var_cadena = var_cadena + " SELECT     " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + ", dbo.TB_ENCABEZADO_CARTERA.VCHA_GAC_GRUPO_ACTUAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID, dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO, dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID, dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2, dbo.TB_ENCABEZADO_CARTERA.VCHA_ESB_ESTABLECIMIENTO_ID FROM dbo.TB_ENCABEZADO_PEDIDOS INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO ON dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO = dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON "
            var_cadena = var_cadena + " dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN INNER JOIN dbo.TB_SALIDAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID AND dbo.TB_SALIDAS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND "
            var_si = MsgBox("Desea el reporte unicamente de artículos nuevos", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_cadena = var_cadena + " dbo.TB_SALIDAS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA <= " + var_fecha_fin + " + 1 - .000001) AND (dbo.TB_ARTICULOS.DTIM_ART_PROMOCION_NUEVOS IS NOT NULL)             "
            Else
               var_cadena = var_cadena + " dbo.TB_SALIDAS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA <= " + var_fecha_fin + " + 1 - .000001) AND (dbo.TB_ARTICULOS.vcha_art_catalogo_vigente in ('VNG2012','BBV2012'))"
            End If
            
            
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_REPORTE_ARTICULOS_NUEVOS where vcha_gac_grupo_Actual_id is null and inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            'rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         
            Set reporte = appl.OpenReport(App.Path + "\rep_Articulos_nuevos_detalle.rpt")
            reporte.RecordSelectionFormula = "{VW_REPORTE_ARTICULOS_NUEVOS_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\detalle_de_Articulos_nuevos_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a generado el reporte " + archivo, vbOKOnly, "ATENCION"
            rs.Open "delete from TB_TEMP_REPORTE_ARTICULOS_NUEVOS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser mayor", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha Final Incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio Incorrecta", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir:
   MsgBox "A surgido un error al generar el reporte", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
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
   Call activa_forma(var_activa_forma_articulos2)
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



