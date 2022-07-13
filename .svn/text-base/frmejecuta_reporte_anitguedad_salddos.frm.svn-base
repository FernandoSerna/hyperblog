VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmejecuta_reporte_anitguedad_salddos 
   Caption         =   "Form1"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   12750
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1350
      Top             =   45
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   825
      Top             =   30
   End
   Begin VB.TextBox txt_fecha 
      Height          =   510
      Left            =   300
      TabIndex        =   2
      Text            =   "16/12/2008 10:56:03 a.m."
      Top             =   2475
      Width           =   3945
   End
   Begin VB.TextBox txt_reloj 
      Height          =   435
      Left            =   300
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1785
      Width           =   3870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1035
      Left            =   390
      TabIndex        =   0
      Top             =   435
      Width           =   3735
   End
   Begin CRVIEWERLibCtl.CRViewer cr 
      Height          =   4500
      Left            =   4380
      TabIndex        =   3
      Top             =   75
      Width           =   8190
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmejecuta_reporte_anitguedad_salddos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim cnn As ADODB.Connection
Dim cnn_SRVDISENO As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rsaux As ADODB.Recordset
Dim rsres As ADODB.Recordset
Dim rsdet As ADODB.Recordset
Dim rsaux1 As ADODB.Recordset
Dim rsaux2 As ADODB.Recordset
Dim rsaux3 As ADODB.Recordset
Dim rsaux4 As ADODB.Recordset
Dim rsaux5 As ADODB.Recordset
Dim rsaux6 As ADODB.Recordset
Dim rsaux7 As ADODB.Recordset
Dim rsaux8 As ADODB.Recordset
Dim rsaux9 As ADODB.Recordset
Dim rsaux10 As ADODB.Recordset
Dim rsaux11 As ADODB.Recordset

Private Sub Command1_Click()
   Me.Timer1.Enabled = False
End Sub

Private Sub Form_Load()
   Me.txt_fecha = Date + 1 - 0.99999
   Set cnn = CreateObject("ADODB.connection")
   Set cnn_SRVDISENO = CreateObject("ADODB.connection")
   Set rs = CreateObject("ADODB.recordset")
   Set rsaux = CreateObject("ADODB.recordset")
   Set rsaux1 = CreateObject("ADODB.recordset")
   Set rsaux2 = CreateObject("ADODB.recordset")
   Set rsaux3 = CreateObject("ADODB.recordset")
   Set rsaux4 = CreateObject("ADODB.recordset")
   Set rsaux5 = CreateObject("ADODB.recordset")
   Set rsaux6 = CreateObject("ADODB.recordset")
   Set rsaux7 = CreateObject("ADODB.recordset")
   Set rsaux8 = CreateObject("ADODB.recordset")
   Set rsaux9 = CreateObject("ADODB.recordset")
   Set rsaux10 = CreateObject("ADODB.recordset")
   Set rsaux11 = CreateObject("ADODB.recordset")
   cnn.Open "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=vianney;Data Source=DISTRIBUCION"
   Me.Timer1.Enabled = True
   Me.Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
  Me.txt_reloj = CStr(Now)
  
End Sub

Private Sub Timer2_Timer()
   Dim var_cadena As String
   Dim var_mes As String
   Dim var_dia As String
   Dim var_año As String
   If Me.txt_reloj = Me.txt_fecha Then
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
         cnn.CommandTimeout = 3600
         cnn.BeginTrans
         rs.Open "select max(inte_tem_consecutivo) from tb_temp_antiguedad_saldos", cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            var_consecutivo = 1
         Else
            var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
         End If
         rs.Close
         rs.Open "insert into tb_temp_antiguedad_saldos (inte_tem_consecutivo, dtim_tem_fecha) values (" + CStr(var_consecutivo) + ", " + var_fecha + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
         'rs.Open "exec SP_ANTIGUEDAD_SALDOS " + var_fecha + "," + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         var_cadena_2 = "INSERT INTO TB_TEMP_ANTIGUEDAD_SALDOS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_AGE_AGENTE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO, FLOA_TEM_IMPORTE, floa_tem_saldo, DTIM_CAR_FECHA, INTE_TEM_DIFERENCIA, FLOA_CAR_TIPO_CAMBIO, FLOA_CAR_IMPORTE_NETO, INTE_CAR_PLAZO)"
         var_cadena_2 = var_cadena_2 + "  select " + CStr(var_consecutivo) + ", " + var_fecha + ", a.vcha_emp_empresa_id, a.vcha_Car_documento, a.vcha_Ser_Serie_id, a.vcha_age_agente_id, a.vcha_cli_clave_id, a.inte_Car_numero, "
         var_cadena_2 = var_cadena_2 + " (round((a.floa_Car_importe_neto/a.floa_Car_tipo_cambio),2) - isnull((select round(sum(floa_car_importe_neto/floa_Car_tipo_cambio),2) from vw_abonos where vcha_emp_empresa_id = a.vcha_emp_empresa_id and vcha_ecu_movimiento_cargo = a.vcha_car_documento and vcha_ecu_serie_cargo = a.vcha_ser_Serie_id and inte_Ecu_numero_cargo = a.inte_car_numero and dtim_Car_fecha <= ((" + var_fecha + " + 1) - .000001)) ,0)) * a.floa_car_tipo_cambio, ((a.floa_Car_importe_neto/a.floa_Car_tipo_cambio) - isnull((select sum(floa_car_importe_neto/floa_Car_tipo_cambio) from vw_abonos where vcha_emp_empresa_id = a.vcha_emp_empresa_id and vcha_ecu_movimiento_cargo = a.vcha_car_documento and vcha_ecu_serie_cargo = a.vcha_ser_Serie_id and inte_Ecu_numero_cargo = a.inte_car_numero and dtim_car_fecha <= ((" + var_fecha + " + 1) - .000001)) ,0)),"
         var_cadena_2 = var_cadena_2 + " a.dtim_Car_fecha, datediff(day, a.dtim_Car_fecha+A.INTE_cAR_PLAZO, " + var_fecha + "), a.floa_car_tipo_Cambio, a.floa_car_importe_neto, a.inte_Car_plazo from tb_encabezado_cartera a where a.dtim_Car_fecha <= ((" + var_fecha + " + 1) - .0000001) and a.char_Car_afectacion = '+' and (a.char_car_Estatus <> 'C' or a.char_Car_estatus is null)"
         rs.Open var_cadena_2, cnn, adOpenDynamic, adLockOptimistic
         rsaux5.Open "EXEC SP_ANTIGUEDAD_SALDOS_RESUMEN_2 " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         rsaux11.Open "select * from tb_empresas ", cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux11.EOF
              var_empresa = IIf(IsNull(rsaux11!vcha_emp_empresa_id), "", rsaux11!vcha_emp_empresa_id)
              If var_empresa <> "" Then
                 Set reporte = appl.OpenReport(App.Path + "\rep_antiguedad_saldos_AGENTE_HISTORICO.rpt")
                 reporte.RecordSelectionFormula = "{TB_TEMP_ANTIGUEDAD_SALDOS_HISTORICO.inte_tem_consecutivo}= " + CStr(var_consecutivo) + " and {TB_TEMP_ANTIGUEDAD_SALDOS_HISTORICO.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
                 For ntablas = 1 To reporte.Database.Tables.Count
                     reporte.Database.Tables(ntablas).SetLogOnInfo "sqlsistema", "vianney", "sa", "elia"
                 Next ntablas
                 reporte.ExportOptions.FormatType = crEFTExcel80
                 reporte.ExportOptions.DestinationType = crEDTDiskFile
                 archivo = "c:\rep_res_ant_saldos_" + rsaux11!vcha_emp_nombre + "_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                 reporte.ExportOptions.DiskFileName = archivo
                 reporte.Export False
                 Set reporte = Nothing
              End If
              rsaux11.MoveNext
         Wend
         rs.Open "delete from tb_temp_antiguedad_saldos where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         rs.Open "delete from tb_temp_antiguedad_saldos_historico where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      End If
   End If
Exit Sub
salir:
If Err.Number = -2147217871 Then
   var_contador = var_contador + 1
   If var_contador <= 3 Then
      Resume
   Else
      MsgBox "A surgido un error al generar el reporte", vbOKOnly, "ATENCION"
      rs.Open "delete from tb_temp_antiguedad_saldos where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      rs.Open "delete from tb_temp_antiguedad_saldos_historico where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   End If
End If
End Sub
            
            
            
            
            
            
            
            
            
            
            
            
            
            
