VERSION 5.00
Begin VB.Form frmreporte_lotes_facturados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de lotes facturados"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   90
      TabIndex        =   4
      Top             =   435
      Width           =   4245
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   2
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
      Picture         =   "frmreporte_lotes_facturados.frx":0000
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
      Picture         =   "frmreporte_lotes_facturados.frx":063A
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
Attribute VB_Name = "frmreporte_lotes_facturados"
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
   Dim a?o As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn_sid_estampados.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_reporte_lotes_facturados", cnn_sid_estampados, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_reporte_lotes_facturados (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn_sid_estampados, adOpenDynamic, adLockOptimistic
            cnn_sid_estampados.CommitTrans
            
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_a?o = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_a?o = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
            
            cnn_sid_estampados.CommandTimeout = 3600
            
            var_cadena = " SELECT DISTINCT TOP 100 PERCENT dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_SALIDAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID AND"
            var_cadena = var_cadena + " dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_ALM_ALMACEN_ID AND dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_SALIDAS.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO WHERE  (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = '15') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C' OR dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL) AND (dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = 'VDIP') AND (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA < " + var_fecha_fin + "+1) ORDER BY dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID"
            rs.Open var_cadena, cnn_sid_estampados, adOpenDynamic, adLockOptimistic
            VAR_NOTAS = ""
            While Not rs.EOF
                  If IsNumeric(rs!vcha_Emo_referencia) Then
                     'If VAR_NOTAS = "" Then
                     '   VAR_NOTAS = VAR_NOTAS + "'" + rs!VCHA_EMO_REFERENCIA + "'"
                     'Else
                     '   VAR_NOTAS = VAR_NOTAS + ",'" + rs!VCHA_EMO_REFERENCIA + "'"
                     'End If
                     If VAR_NOTAS = "" Then
                        VAR_NOTAS = VAR_NOTAS + "" + rs!vcha_Emo_referencia + ""
                     Else
                        VAR_NOTAS = VAR_NOTAS + "," + rs!vcha_Emo_referencia + ""
                     End If
                  
                  End If
                  rs.MoveNext
            Wend
            rs.Close
            'MsgBox var_notas
            VAR_NOTAS = "(" + VAR_NOTAS + ")"
            If VAR_NOTAS <> "()" Then
               var_cadena = " SELECT d.VCHA_PRO_PRODUCTO_ID, p.FLOA_CTO_UNIT AS floa_com_costo, SUM(d.BINT_DNO_CANTIDAD) AS floa_com_cantidad, d.BINT_NOT_NOTA_ID, C.VCHA_NOT_CLASIFICACION, p.VCHA_LOT_LOTE_ID, dbo.TB_MODULOS.BINT_MOD_MODULO_ID, dbo.TB_MODULOS.VCHA_MOD_DESCRIPCION FROM dbo.TB_LOTES p INNER JOIN dbo.TB_DNOTAS d ON p.VCHA_PRO_PRODUCTO_ID = d.VCHA_PRO_PRODUCTO_ID AND p.VCHA_LOT_LOTE_ID = d.VCHA_LOT_LOTE_ID INNER JOIN dbo.TB_NOTAS C ON d.BINT_NOT_NOTA_ID = C.BINT_NOT_NOTA_ID INNER JOIN dbo.TB_MODULOS ON p.BINT_MOD_MODULO_ID = dbo.TB_MODULOS.BINT_MOD_MODULO_ID WHERE (d.BINT_NOT_NOTA_ID IN " + VAR_NOTAS + ") GROUP BY d.VCHA_PRO_PRODUCTO_ID, p.FLOA_CTO_UNIT, d.BINT_NOT_NOTA_ID, C.VCHA_NOT_CLASIFICACION, p.VCHA_LOT_LOTE_ID, dbo.TB_MODULOS.BINT_MOD_MODULO_ID , dbo.TB_MODULOS.VCHA_MOD_DESCRIPCION             "
               'MsgBox var_cadena
               rs.Open var_cadena, cnn_cantia, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     'MsgBox CStr(var_consecutivo)
                     var_cadena = "insert into TB_TEMP_REPORTE_LOTES_FACTURADOS (inte_tem_consecutivo, dtim_tem_fecha_inicio, dtim_tem_fecha_fin, vcha_art_articulo_id, floa_com_cantidad, bint_not_nota_id, vcha_lot_lote_id, vcha_not_clasificacion, bint_mod_modulo_id, vcha_mod_descripcion) values "
                     var_cadena = var_cadena + " (" + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ",'" + rs!vcha_pro_producto_id + "'," + CStr(rs!floa_com_cantidad) + "," + CStr(rs!BINT_NOT_NOTA_ID) + ", '" + rs!vcha_lot_lote_id + "','" + rs!vcha_not_clasificacion + "'," + CStr(rs!bint_mod_modulo_id) + ", '" + rs!vcha_mod_descripcion + "' )"
                     'MsgBox var_cadena
                     rsaux.Open var_cadena, cnn_sid_estampados, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rs.Close
               rs.Open "delete from TB_TEMP_REPORTE_LOTES_FACTURADOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn_sid_estampados, adOpenDynamic, adLockOptimistic
               rs.Open "select * from TB_TEMP_REPORTE_LOTES_FACTURADOS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn_sid_estampados, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     var_cadena = " SELECT TOP 100 PERCENT dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID , dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, DBO.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, FLOA_SAL_COSTO, FLOA_SAL_PRECIO FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_SALIDAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO INNER JOIN "
                     var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID AND dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_ALM_ALMACEN_ID AND dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_SALIDAS.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO INNER JOIN dbo.TB_DETALLE_EMBARQUES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID AND "
                     var_cadena = var_cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO WHERE (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = '15') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'VDIP') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA = '" + CStr(rs!BINT_NOT_NOTA_ID) + "') and (char_car_estatus <> 'C' or char_car_Estatus is null)AND (dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = '" + rs!vcha_Art_Articulo_id + "') ORDER BY dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID"
                     
                     rsaux.Open var_cadena, cnn_sid_estampados, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        'MsgBox "update TB_TEMP_REPORTE_LOTES_FACTURADOS set inte_Car_numero = " + CStr(rsaux!inte_Car_numero) + ", DTIM_CAR_FECHA = " + CStr(Format(rsaux!DTIM_CAR_FECHA, "Short Date")) + ", INTE_EMB_EMBARQUE = " + CStr(rsaux!INTE_EMB_EMBARQUE) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND BINT_NOT_NOTA_ID = " + CStr(rs!BINT_NOT_NOTA_ID) + " AND VCHA_aRT_ARTICULO_ID = '" + rs!vcha_art_Articulo_id + "'"
                        var_dia = CStr(Day(rsaux!dtim_Car_fecha))
                        var_mes = CStr(Month(rsaux!dtim_Car_fecha))
                        var_a?o = CStr(Year(rsaux!dtim_Car_fecha))
                        If Len(Trim(var_dia)) = 1 Then
                           var_dia = "0" + var_dia
                        End If
                        If Len(Trim(var_mes)) = 1 Then
                           var_mes = "0" + var_mes
                        End If
                        var_fecha = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
                        
                        rsaux1.Open "update TB_TEMP_REPORTE_LOTES_FACTURADOS set inte_Car_numero = " + CStr(rsaux!inte_car_numero) + ", DTIM_CAR_FECHA = " + var_fecha + ", INTE_EMB_EMBARQUE = " + CStr(rsaux!inte_emb_embarque) + ", FLOA_TEM_COSTO_TELA_CRUDA = " + CStr(rsaux!floa_Sal_costo) + ", FLOA_TEM_COSTO_ACABADO = " + CStr(rsaux!floa_Sal_precio) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND BINT_NOT_NOTA_ID = " + CStr(rs!BINT_NOT_NOTA_ID) + " AND VCHA_aRT_ARTICULO_ID = '" + rs!vcha_Art_Articulo_id + "'", cnn_sid_estampados, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux.Close
                     rs.MoveNext
               Wend
               rs.Close
            
            
            
               Dim dl As Long                                 ' Valor devuelto por la funci?n API
               Dim sAttributes As String                  ' Aributos
               Dim sDriver As String                       ' Nombre del controlador
               Dim sDescription As String                ' Descripci?n del DSN
               Dim sDsnName As String                  ' Nombre del DSN

               Const ODBC_ADD_SYS_DSN As Long = 4         ' Se crear? un DSN de sistema
               Const vbAPINull As Long = 0&                         ' Puntero NULL

               ' se elimina
               Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminar? un DSN de sistema
               sDsnName = "DSN=sqlsistema"
               sDriver = "SQL Server"
               dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)

               'se crea
               sDsnName = "sqlsistema"
               sDescription = "sqlsistema"
               sDriver = "SQL Server"
               sAttributes = "DSN=" & sDsnName & Chr(0)
               sAttributes = sAttributes & "Server=" + "sqlquezada2" & Chr$(0)
               sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
               sAttributes = sAttributes & "Database=" + "sidtextilera" & Chr(0)
               strAttributes = strAttributes & "UID=sa" & Chr$(0)
               strAttributes = strAttributes & "PWD=elia" & Chr$(0)
               dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
            
            
            
            
            
            
           
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
               
               
               var_si = MsgBox("?Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
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
               
               
               

               ' se elimina
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
               
               
               
               
               
               
               If rs.State = 1 Then
                  rs.Close
               End If
            Else
               MsgBox "No existen notas facturas para las fecha seleccionadas", vbOKOnly, "ATENCION"
            End If
            rs.Open "delete from TB_TEMP_REPORTE_LOTES_FACTURADOS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn_sid_estampados, adOpenDynamic, adLockOptimistic
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
   'MsgBox cnn_sid_estampados.ConnectionString
   rs.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn_sid_estampados, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_conexion_cantia = IIf(IsNull(rs!vcha_uor_conexion), "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=sipestampados;Data Source=sqlquezada2", rs!vcha_uor_conexion)
   Else
      var_conexion_cantia = "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=sipestampados;Data Source=sqlquezada2"
   End If
   rs.Close
   
   If var_conexion_cantia <> "" Then
      cnn_cantia.Open var_conexion_cantia
   End If
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



