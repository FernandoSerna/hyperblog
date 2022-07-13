VERSION 5.00
Begin VB.Form frmreporte_relacion_cobranza_auditoria_nuevo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte relación de cobranza"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   45
      TabIndex        =   7
      Top             =   345
      Width           =   4305
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmreporte_relacion_cobranza_auditoria_nuevo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3945
      Picture         =   "frmreporte_relacion_cobranza_auditoria_nuevo.frx":0312
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   75
      TabIndex        =   4
      Top             =   405
      Width           =   4245
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   315
         Width           =   1140
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   1
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
End
Attribute VB_Name = "frmreporte_relacion_cobranza_auditoria_nuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
  ' On Error GoTo salir:
      If IsDate(Me.txt_fin) Then
         If IsDate(Me.txt_inicio) Then
            
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_REPORTE_RELACION_COBRANZA_AUDITORIA", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_REPORTE_RELACION_COBRANZA_AUDITORIA (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            var_fecha_inicio = var_dia + "/" + var_mes + "/" + var_año
            var_fecha_inicio_sql = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
      
             
             
            var_dia = CStr(Day(CDate(txt_fin) + 1))
            var_mes = CStr(Month(CDate(txt_fin) + 1))
            var_año = CStr(Year(CDate(txt_fin) + 1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            var_fecha_fin = var_dia + "/" + var_mes + "/" + var_año
            var_fecha_fin_sql = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
             
            'var_cadena = "select * from vw_relacion_cobranza_auditoria@msqlsiddist, vw_depositos_banc Where to_char(folio) = to_char(INTE_RCO_NUMERO_DEPOSITO) and to_date(to_char(dtim_rco_fecha_relacion, 'dd/mm/yyyy'), 'dd/mm/yyyy') > to_date('" + var_fecha_inicio + "', 'dd/mm/yyyy') and to_date(to_char(dtim_rco_fecha_relacion, 'dd/mm/yyyy'), 'dd/mm/yyyy') < to_date('" + var_fecha_fin + "', 'dd/mm/yyyy') "
            var_cadena = "select  ta.VCHA_RCO_FOLIO, ta.VCHA_AGE_AGENTE_ID, ta.VCHA_AGE_NOMBRE, ta.VCHA_BANCO_DEPOSITO, ta.VCHA_NOMBRE_BANCO_DEPOSITO, ta.DTIM_RCO_FECHA_DEPOSITO, ta.VCHA_RCO_CHEQUE, ta.VCHA_BAN_BANCO_ID, ta.VCHA_BAN_NOMBRE, ta.VCHA_CLI_CLAVE_ID, ta.VCHA_CLI_NOMBRE, ta.VCHA_CAR_DOCUMENTO, ta.INTE_CAR_NUMERO, ta.DTIM_CAR_FECHA, ta.FLOA_CAR_IMPORTE, ta.FLOA_CAR_TIPO_CAMBIO, ta.FLOA_RCO_DESCUENTO_OTORGADO, ta.VCHA_RCO_DEPOSITO, ta.FLOA_RCO_IMPORTE, ta.FLOA_RCO_TIPO_CAMBIO, ta.DTIM_RCO_FECHA_RELACION, ta.INTE_RCO_NUMERO_DEPOSITO, tb.REFERENCIA,  to_char(tb.FECHA_AUTORIZACION,'dd/mm/yy') as FECHA_AUTORIZACION, tb.IMPORTE, tb.CUENTA,"
            var_cadena = var_cadena + " tb.DIVISA, tb.ORIGEN, tb.FOLIO, tb.NO_AUTORIZACION, tb.DESCRIPCION, tb.TIPO, tb.RECIBO,  to_char(tb.fecha_deposito, 'dd/mm/yy') as fecha_Deposito  from vw_relacion_cobranza_auditoria@msqlsiddist ta, vw_depositos_banc tb Where to_char(folio) = to_char(ta.INTE_RCO_NUMERO_DEPOSITO) and to_date(to_char(ta.dtim_rco_fecha_relacion, 'dd/mm/yyyy'), 'dd/mm/yyyy') > to_date('" + var_fecha_inicio + "', 'dd/mm/yyyy') and to_date(to_char(ta.dtim_rco_fecha_relacion, 'dd/mm/yyyy'), 'dd/mm/yyyy') < to_date('" + var_fecha_fin + "', 'dd/mm/yyyy') order by tb.fecha_deposito"
            rs.Open var_cadena, cnnoracle_2, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  DTIM_RCO_FECHA_DEPOSITO_sql = CStr(rs!dtim_rco_fecha_deposito)
                  var_dia = CStr(Day(CDate(DTIM_RCO_FECHA_DEPOSITO_sql)))
                  var_mes = CStr(Month(CDate(DTIM_RCO_FECHA_DEPOSITO_sql)))
                  var_año = CStr(Year(CDate(DTIM_RCO_FECHA_DEPOSITO_sql)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  dtim_rco_fecha_deposito = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                  
                  DTIM_CAR_FECHA_sql = CStr(CDate(IIf(IsNull(rs!dtim_Car_fecha), "01/01/1900", rs!dtim_Car_fecha)))
                  var_dia = CStr(Day(CDate(DTIM_CAR_FECHA_sql)))
                  var_mes = CStr(Month(CDate(DTIM_CAR_FECHA_sql)))
                  var_año = CStr(Year(CDate(DTIM_CAR_FECHA_sql)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  dtim_Car_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                  
                  DTIM_RCO_FECHA_RELACION_sql = CStr(rs!dtim_rco_fecha_relacion)
                  var_dia = CStr(Day(CDate(DTIM_RCO_FECHA_RELACION_sql)))
                  var_mes = CStr(Month(CDate(DTIM_RCO_FECHA_RELACION_sql)))
                  var_año = CStr(Year(CDate(DTIM_RCO_FECHA_RELACION_sql)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  dtim_rco_fecha_relacion = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                  
                  
                  
                  FECHA_DEPOSITO_sql = Mid(CStr(rs!FECHA_DEPOSITO), 1, 6) + "20" + Mid(CStr(rs!FECHA_DEPOSITO), 7, 2)
                  'MsgBox FECHA_DEPOSITO_sql
                  var_dia = CStr(Day(CDate(FECHA_DEPOSITO_sql)))
                  var_mes = CStr(Month(CDate(FECHA_DEPOSITO_sql)))
                  var_año = CStr(Year(CDate(FECHA_DEPOSITO_sql)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  FECHA_DEPOSITO = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                  
                  
                  FECHA_AUTORIZACION_sql = Mid(CStr(rs!FECHA_AUTORIZACION), 1, 6) + "20" + Mid(CStr(rs!FECHA_AUTORIZACION), 7, 2)
                  
                  var_dia = CStr(Day(CDate(FECHA_AUTORIZACION_sql)))
                  var_mes = CStr(Month(CDate(FECHA_AUTORIZACION_sql)))
                  var_año = CStr(Year(CDate(FECHA_AUTORIZACION_sql)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  FECHA_AUTORIZACION = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                  
     
                  
                  
                  var_cadena = "INSERT INTO TB_TEMP_REPORTE_RELACION_COBRANZA_AUDITORIA (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_RCO_FOLIO, VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE,         VCHA_BANCO_DEPOSITO,                                                               VCHA_NOMBRE_BANCO_DEPOSITO,    DTIM_RCO_FECHA_DEPOSITO, VCHA_RCO_CHEQUE, VCHA_BAN_BANCO_ID, VCHA_BAN_NOMBRE, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_CAR_DOCUMENTO, INTE_CAR_NUMERO, DTIM_CAR_FECHA, FLOA_CAR_IMPORTE, FLOA_CAR_TIPO_CAMBIO, FLOA_RCO_DESCUENTO_OTORGADO, VCHA_RCO_DEPOSITO, FLOA_RCO_IMPORTE, FLOA_RCO_TIPO_CAMBIO, DTIM_RCO_FECHA_RELACION, INTE_RCO_NUMERO_DEPOSITO, REFERENCIA, FECHA_DEPOSITO, FECHA_AUTORIZACION, IMPORTE, CUENTA, DIVISA, ORIGEN, FOLIO, NO_AUTORIZACION, DESCRIPCION, TIPO, RECIBO)"
                  'MsgBox CStr(rs!floa_rco_tipo_cambio)
                  var_cadena = var_cadena + " Values ( " + CStr(var_consecutivo) + "," + var_fecha_inicio_sql + "," + var_fecha_fin_sql + ",'" + rs!vcha_Rco_folio + "', '" + rs!VCHA_AGE_AGENTE_ID + "', '" + rs!VCHA_AGE_NOMBRE + "', '" + IIf(IsNull(rs!vcha_banco_deposito), "", rs!vcha_banco_deposito) + "','" + IIf(IsNull(rs!VCHA_NOMBRE_BANCO_DEPOSITO), "", rs!VCHA_NOMBRE_BANCO_DEPOSITO) + "'," + dtim_rco_fecha_deposito + ", '" + rs!VCHA_rCO_CHEQUE + "', '" + rs!vcha_ban_banco_id + "','" + rs!VCHA_BAN_NOMBRE + "', '" + rs!vcha_cli_clave_id + "', '" + rs!VCHA_CLI_NOMBRE + "', '" + rs!vcha_Car_documento + "', " + CStr(rs!inte_Car_numero) + "," + dtim_Car_fecha + "," + CStr(IIf(IsNull(rs!floa_Car_importe), 0, rs!floa_Car_importe)) + ", " + CStr(IIf(IsNull(rs!floa_car_tipo_cambio), 0, rs!floa_car_tipo_cambio)) + ", "
                  var_cadena = var_cadena + CStr(rs!FLOA_RCO_DESCUENTO_OTORGADO) + ",'" + rs!vcha_rco_deposito + "', " + CStr(rs!floa_rco_importe) + ", " + CStr(IIf(IsNull(rs!floa_rco_tipo_cambio), 1, rs!floa_rco_tipo_cambio)) + ",  " + dtim_rco_fecha_relacion + ","
                  'MsgBox var_cadena
                  var_cadena = var_cadena + CStr(rs!inte_rco_numero_deposito) + ", '" + IIf(IsNull(rs!Referencia), "", rs!Referencia) + "', " + FECHA_DEPOSITO + "," + FECHA_AUTORIZACION + ", " + CStr(rs!Importe) + ", '" + rs!CUENTA + "', '" + rs!DIVISA + "', '" + rs!Origen + "', " + CStr(rs!FOLIO) + ", " + CStr(rs!NO_AUTORIZACION) + ", '" + rs!descripcion + "','" + rs!tipo + "', '" + IIf(IsNull(rs!recibo), "", rs!recibo) + "')"
                  'x = CStr(rs!inte_rco_numero_deposito) + ", '" + rs!Referencia + "', " + FECHA_DEPOSITO + "," + FECHA_AUTORIZACION + ", " + CStr(rs!Importe) + ", '" + rs!CUENTA + "', '" + rs!DIVISA + "', '" + rs!Origen + "', " + CStr(rs!FOLIO) + ", " + CStr(rs!NO_AUTORIZACION) + ", '" + rs!DESCRIPCION + "','" + rs!tipo + "', '" + IIf(IsNull(rs!recibo), "", rs!recibo) + "')"
                  'MsgBox x
                  rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            rs.Open "delete from  TB_TEMP_REPORTE_RELACION_COBRANZA_AUDITORIA where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "select * from TB_TEMP_REPORTE_RELACION_COBRANZA_AUDITORIA where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Set reporte = appl.OpenReport(App.Path + "\REP_RELACION_COBRANZA_AUDITORIA_nuevo.rpt")
               reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_RELACION_COBRANZA_AUDITORIA.inte_tem_consecutivo} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\relacion_cobranza_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + ".xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            Else
               MsgBox "No existe relaciones de cobranza para la fecha seleccionada", vbOKOnly, "ATENCION"
            End If
            rs.Close
            rs.Open "delete from TB_TEMP_REPORTE_RELACION_COBRANZA_AUDITORIA where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "Fecha inicio incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   
   Exit Sub
salir:
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   MsgBox "A surgido un error al generar el archivo, puede que este este abierto.", vbOKOnly, "ATENCION"
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
   Me.txt_fin = Date
   Me.txt_inicio = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_articulos2)
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

