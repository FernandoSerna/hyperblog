VERSION 5.00
Begin VB.Form frmreporte_antiguedad_saldos_historico_empresas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen de antiguedad de saldos"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   3285
      Begin VB.TextBox txt_fecha 
         Height          =   375
         Left            =   990
         TabIndex        =   2
         Top             =   915
         Width           =   1335
      End
      Begin VB.CommandButton cmd_ejecuta 
         Caption         =   "Ejecuta Reportes"
         Height          =   630
         Left            =   75
         TabIndex        =   1
         Top             =   195
         Width           =   3150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   990
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmreporte_antiguedad_saldos_historico_empresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_servidor_Temporal As String
Dim var_base_Datos_Temporal As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_mes As Integer

Private Sub cmd_ejecuta_Click()
   
   
   
   Dim var_cadena As String
   Dim var_mes As String
   Dim var_dia As String
   Dim var_año As String
   If IsDate(Me.txt_fecha) Then
      var_cadena = ""
      'On Error GoTo salir:
      If IsDate(Me.txt_fecha) Then
         var_contador = 0
         var_fecha_fin_1 = CDate(txt_fecha)
         var_dia = CStr(Day(CDate(txt_fecha)))
         var_mes = CStr(Month(CDate(txt_fecha)))
         var_año = CStr(Year(CDate(txt_fecha)))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
         CommandTimeout = 3600
         cnn_reportes.BeginTrans
         'MsgBox cnn_reportes.ConnectionString
         rs.Open "select max(inte_tem_consecutivo) from tb_temp_antiguedad_saldos", cnn_reportes, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            var_consecutivo = 1
         Else
            var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
         End If
         rs.Close
         rs.Open "insert into tb_temp_antiguedad_saldos (inte_tem_consecutivo, dtim_tem_fecha) values (" + CStr(var_consecutivo) + ", " + var_fecha + ")", cnn_reportes, adOpenDynamic, adLockOptimistic
         cnn_reportes.CommitTrans
         'rs.Open "exec SP_ANTIGUEDAD_SALDOS " + var_fecha + "," + CStr(var_consecutivo), cnn_reportes, adOpenDynamic, adLockOptimistic
         var_cadena_2 = "INSERT INTO TB_TEMP_ANTIGUEDAD_SALDOS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_AGE_AGENTE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO, FLOA_TEM_IMPORTE, floa_tem_saldo, DTIM_CAR_FECHA, INTE_TEM_DIFERENCIA, FLOA_CAR_TIPO_CAMBIO, FLOA_CAR_IMPORTE_NETO, INTE_CAR_PLAZO)"
         var_cadena_2 = var_cadena_2 + "  select " + CStr(var_consecutivo) + ", " + var_fecha + ", a.vcha_emp_empresa_id, a.vcha_Car_documento, a.vcha_Ser_Serie_id, a.vcha_age_agente_id, a.vcha_cli_clave_id, a.inte_Car_numero, "
         var_cadena_2 = var_cadena_2 + " (round((a.floa_Car_importe_neto/a.floa_Car_tipo_cambio),2) - isnull((select round(sum(floa_car_importe_neto/floa_Car_tipo_cambio),2) from vw_abonos where vcha_emp_empresa_id = a.vcha_emp_empresa_id and vcha_ecu_movimiento_cargo = a.vcha_car_documento and vcha_ecu_serie_cargo = a.vcha_ser_Serie_id and inte_Ecu_numero_cargo = a.inte_car_numero and dtim_Car_fecha <= ((" + var_fecha + " + 1) - .000001)) ,0)) * a.floa_car_tipo_cambio, ((a.floa_Car_importe_neto/a.floa_Car_tipo_cambio) - isnull((select sum(floa_car_importe_neto/floa_Car_tipo_cambio) from vw_abonos where vcha_emp_empresa_id = a.vcha_emp_empresa_id and vcha_ecu_movimiento_cargo = a.vcha_car_documento and vcha_ecu_serie_cargo = a.vcha_ser_Serie_id and inte_Ecu_numero_cargo = a.inte_car_numero and dtim_car_fecha <= ((" + var_fecha + " + 1) - .000001)) ,0)),"
         var_cadena_2 = var_cadena_2 + " a.dtim_Car_fecha, datediff(day, a.dtim_Car_fecha+A.INTE_cAR_PLAZO, " + var_fecha + "), a.floa_car_tipo_Cambio, a.floa_car_importe_neto, a.inte_Car_plazo from tb_encabezado_cartera a where a.dtim_Car_fecha <= ((" + var_fecha + " + 1) - .0000001) and a.char_Car_afectacion = '+' and (a.char_car_Estatus <> 'C' or a.char_Car_estatus is null)"
         cnn_reportes.CommandTimeout = 36000
         'MsgBox cnn_reportes.ConnectionString
         rs.Open var_cadena_2, cnn_reportes, adOpenDynamic, adLockOptimistic
         rsaux5.Open "EXEC SP_ANTIGUEDAD_SALDOS_RESUMEN_2 " + CStr(var_consecutivo), cnn_reportes, adOpenDynamic, adLockOptimistic
         Dim var_no_termina As Boolean
         var_no_termina = False
         While var_no_termina = False
               'MsgBox cnn_reportes.ConnectionString
               rsaux5.Open "select * from TB_TEMP_ANTIGUEDAD_SALDOS_HISTORICO where inte_tem_tabla = 2 and inte_tem_consecutivo = " + CStr(var_consecutivo), cnn_reportes, adOpenDynamic, adLockOptimistic
               If rsaux5.EOF Then
                  var_no_termina = False
               Else
                  var_no_termina = True
               End If
               rsaux5.Close
         Wend
                 
         If rsaux11.State = 1 Then
            rsaux11.Close
         End If
         rsaux11.Open "DELETE FROM TB_TEMP_ANTIGUEDAD_SALDOS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND VCHA_AGE_AGENTE_ID IS NULL", cnn_reportes, adOpenDynamic, adLockOptimistic
         
         rsaux11.Open "select distinct tb_empresas.vcha_emp_empresa_id, tb_empresas.vcha_emp_nombre from tb_empresas, TB_TEMP_ANTIGUEDAD_SALDOS where tb_Empresas.vcha_emp_empresa_id = TB_TEMP_ANTIGUEDAD_SALDOS.vcha_emp_Empresa_id and TB_TEMP_ANTIGUEDAD_SALDOS.inte_tem_consecutivo = " + CStr(var_consecutivo), cnn_reportes, adOpenDynamic, adLockOptimistic
         var_empresa_anterioR = var_empresa
         
         Dim dl As Long                                 ' Valor devuelto por la función API
         Dim sAttributes As String                  ' Aributos
         Dim sDriver As String                       ' Nombre del controlador
         Dim sDescription As String                ' Descripción del DSN
         Dim sDsnName As String                  ' Nombre del DSN
         Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
         Const vbAPINull As Long = 0&                         ' Puntero NULL

         ' se elimina
         Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema
         rsaux10.Open "DELETE FROM TB_TEMP_ANTIGUEDAD_SALDOS_HISTORICO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND VCHA_AGE_AGENTE_ID IS NULL", cnn_reportes, adOpenDynamic, adLockOptimistic
         While Not rsaux11.EOF
              var_empresa = IIf(IsNull(rsaux11!VCHA_EMP_EMPRESA_ID), "", rsaux11!VCHA_EMP_EMPRESA_ID)
              If var_empresa <> "" Then

                 sDsnName = "DSN=sqlsistema"
                 sDriver = "SQL Server"
                 dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)

   'se crea
                 sDsnName = "sqlsistema"
                 sDescription = "sqlsistema"
                 sDriver = "SQL Server"
                 sAttributes = "DSN=" & sDsnName & Chr(0)
                 sAttributes = sAttributes & "Server=" + var_sr_reportes & Chr$(0)
                 sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
                 sAttributes = sAttributes & "Database=" + var_bd_reportes & Chr(0)
                 strAttributes = strAttributes & "UID=sa" & Chr$(0)
                 strAttributes = strAttributes & "PWD=elia" & Chr$(0)
                 dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
                 
                 
                 Set reporte = appl.OpenReport(App.Path + "\rep_antiguedad_saldos_AGENTE_HISTORICO.rpt")
                 reporte.RecordSelectionFormula = "{TB_TEMP_ANTIGUEDAD_SALDOS_HISTORICO.inte_tem_consecutivo}= " + CStr(var_consecutivo) + " and {TB_TEMP_ANTIGUEDAD_SALDOS_HISTORICO.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
                 For ntablas = 1 To reporte.Database.Tables.Count
                     reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                 Next ntablas
                 reporte.ExportOptions.FormatType = crEFTExcel80
                 reporte.ExportOptions.DestinationType = crEDTDiskFile
                 archivo = "c:\reportessid\rep_res_ant_saldos_" + rsaux11!VCHA_EMP_NOMBRE + "_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                 reporte.ExportOptions.DiskFileName = archivo
                 reporte.Export False
                 Set reporte = Nothing
                 
                 
                 
              End If
              rsaux11.MoveNext
         Wend
         var_empresa = var_empresa_anterioR
         rs.Open "delete from tb_temp_antiguedad_saldos where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn_reportes, adOpenDynamic, adLockOptimistic
         rs.Open "delete from tb_temp_antiguedad_saldos_historico where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn_reportes, adOpenDynamic, adLockOptimistic
      End If
   End If
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
Exit Sub
salir:
If Err.Number = -2147217871 Then
   var_contador = var_contador + 1
   If var_contador <= 3 Then
      Resume
   Else
      MsgBox "A surgido un error al generar el reporte", vbOKOnly, "ATENCION"
      rs.Open "delete from tb_temp_antiguedad_saldos where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn_reportes, adOpenDynamic, adLockOptimistic
      rs.Open "delete from tb_temp_antiguedad_saldos_historico where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn_reportes, adOpenDynamic, adLockOptimistic
   End If
End If

End Sub

Private Sub Form_Load()
   var_servidor_Temporal = var_sr_reportes
   var_base_Datos_Temporal = var_bd_reportes
   'var_sr_reportes = "SQLQUEZADA"
   'var_bd_reportes = "SIDQUEZADA"
   Top = 3000
   Left = 3500
   Me.txt_fecha = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_sr_reportes = var_servidor_Temporal
   var_bd_reportes = var_base_Datos_Temporal
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmcalendario.mes.Value = CDate(Me.txt_fecha)
      frmcalendario.Caption = CStr(CDate(Me.txt_fecha))
      frmcalendario.Show 1
      txt_fecha = var_fecha_general
   End If
End Sub
