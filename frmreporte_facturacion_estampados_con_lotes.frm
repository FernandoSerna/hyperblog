VERSION 5.00
Begin VB.Form frmreporte_facturacion_estampados_con_lotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de facturación "
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmreporte_facturacion_estampados_con_lotes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3975
      Picture         =   "frmreporte_facturacion_estampados_con_lotes.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   90
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
      Left            =   0
      TabIndex        =   7
      Top             =   270
      Width           =   4395
   End
End
Attribute VB_Name = "frmreporte_facturacion_estampados_con_lotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim cnn_cantia As ADODB.Connection


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
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_REPORTE_FACTURACION_CON_LOTES", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_REPORTE_FACTURACION_CON_LOTES (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            
            
            var_fecha_fin_1 = CDate(txt_fin)
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
            
            cnn.CommandTimeout = 360
            
            
            var_cadena = "INSERT INTO TB_TEMP_REPORTE_FACTURACION_CON_LOTES (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_AGE_AGENTE_ID, VCHA_CLI_CLAVE_ID, VCHA_SER_SERIE_ID, VCHA_CAR_DOCUMENTO, INTE_CAR_NUMERO, DTIM_CAR_FECHA, INTE_EMB_EMBARQUE, VCHA_ART_ARTICULO_ID, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_CANTIDAD, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_MOV_MOVIMIENTO_ID, VCHA_EMO_REFERENCIA) "
            var_cadena = var_cadena + " SELECT " + CStr(var_consecutivo) + ", " + var_fecha_inicio + "," + var_fecha_fin + " -.000001,dbo.TB_ENCABEZADO_CARTERA.VCHA_AGE_AGENTE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_SALIDAS.FLOA_SAL_COSTO,  dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, dbo.TB_SALIDAS.FLOA_SAL_PROMOCION_1, dbo.TB_SALIDAS.FLOA_SAL_PROMOCION_2, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2, dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA FROM dbo.TB_ENCABEZADO_CARTERA INNER JOIN dbo.TB_SALIDAS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = "
            var_cadena = var_cadena + " dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND "
            var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO AND dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = dbo.TB_SALIDAS.VCHA_SER_SERIE_ID AND dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = dbo.TB_SALIDAS.INTE_CAR_NUMERO INNER JOIN dbo.TB_DETALLE_EMBARQUES ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID AND dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID AND dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_SALIDAS.INTE_SAL_NUMERO = dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND "
            var_cadena = var_cadena + " dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_SALIDAS.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO WHERE (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA <= " + var_fecha_fin + " + 1 - .000001) AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            rsaux.Open "SELECT VCHA_EMO_REFERENCIA, VCHA_ART_aRTICULO_ID FROM TB_TEMP_REPORTE_FACTURACION_CON_LOTES WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND VCHA_MOV_MOVIMIENTO_ID = 'VDIP'", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
               VAR_NOTAS = IIf(IsNull(rsaux(0).Value), "", rsaux(0).Value)
               If VAR_NOTAS <> "" Then
                  var_cadena = " SELECT d.VCHA_PRO_PRODUCTO_ID, p.FLOA_CTO_UNIT AS floa_com_costo, SUM(d.BINT_DNO_CANTIDAD) AS floa_com_cantidad, d.BINT_NOT_NOTA_ID, C.VCHA_NOT_CLASIFICACION, p.VCHA_LOT_LOTE_ID, dbo.TB_MODULOS.BINT_MOD_MODULO_ID, dbo.TB_MODULOS.VCHA_MOD_DESCRIPCION FROM dbo.TB_LOTES p INNER JOIN dbo.TB_DNOTAS d ON p.VCHA_PRO_PRODUCTO_ID = d.VCHA_PRO_PRODUCTO_ID AND p.VCHA_LOT_LOTE_ID = d.VCHA_LOT_LOTE_ID INNER JOIN dbo.TB_NOTAS C ON d.BINT_NOT_NOTA_ID = C.BINT_NOT_NOTA_ID INNER JOIN dbo.TB_MODULOS ON p.BINT_MOD_MODULO_ID = dbo.TB_MODULOS.BINT_MOD_MODULO_ID WHERE (d.BINT_NOT_NOTA_ID = " + VAR_NOTAS + " AND d.VCHA_PRO_PRODUCTO_ID = '" + rsaux!vcha_Art_articulo_id + "') GROUP BY d.VCHA_PRO_PRODUCTO_ID, p.FLOA_CTO_UNIT, d.BINT_NOT_NOTA_ID, C.VCHA_NOT_CLASIFICACION, p.VCHA_LOT_LOTE_ID, dbo.TB_MODULOS.BINT_MOD_MODULO_ID , dbo.TB_MODULOS.VCHA_MOD_DESCRIPCION             "
                  rs.Open var_cadena, cnn_cantia, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                        var_cadena = "UPDATE  TB_TEMP_REPORTE_FACTURACION_CON_LOTES SET VCHA_TEM_NOTA = '" + CStr(rs!BINT_NOT_NOTA_ID) + "', vcha_lot_lote_id = '" + rs!vcha_lot_lote_id + "', vcha_not_clasificacion = '" + rs!vcha_not_clasificacion + "', bint_mod_modulo_id = " + CStr(rs!bint_mod_modulo_id) + ", vcha_mod_descripcion = '" + rs!vcha_mod_descripcion + "'  WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND VCHA_EMO_REFERENCIA = '" + CStr(VAR_NOTAS) + "' AND VCHA_MOV_MOVIMIENTO_ID = 'VDIP' AND VCHA_aRT_aRTICULO_ID = '" + rsaux!vcha_Art_articulo_id + "'"
                        'MsgBox var_cadena
                        rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        rs.MoveNext
                  End If
                  rs.Close
               End If
               rsaux.MoveNext
            Wend
            rsaux.Close
            
            
            
            rs.Open "delete from TB_TEMP_REPORTE_FACTURACION_CON_LOTES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
            
            
           
            Set reporte = appl.OpenReport(App.Path + "\REP_FACTURACION_LOTES.rpt")
            reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_LOTES_FACTURADOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de lotes facturados"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\REP_FACTURACION_LOTES.rpt")
               reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_LOTES_FACTURADOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\Reporte_lotes_facturados_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "delete from TB_TEMP_REPORTE_LOTES_FACTURADOS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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



Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
   Set cnn_cantia = CreateObject("ADODB.connection")
   rs.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_conexion_cantia = IIf(IsNull(rs!vcha_uor_conexion), "", rs!vcha_uor_conexion)
   Else
      var_conexion_cantia = ""
   End If
   rs.Close
   If var_conexion_cantia <> "" Then
      cnn_cantia.Open var_conexion_cantia
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_articulos2)
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



