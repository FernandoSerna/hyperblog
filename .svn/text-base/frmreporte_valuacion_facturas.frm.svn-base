VERSION 5.00
Begin VB.Form frmreporte_valuacion_facturas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valuación de Facturación"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4500
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   630
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   30
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   765
      Left            =   75
      TabIndex        =   5
      Top             =   420
      Width           =   4335
      Begin VB.OptionButton opt_almacen 
         Caption         =   "Almacen General"
         Height          =   315
         Left            =   315
         TabIndex        =   9
         Top             =   915
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.OptionButton opt_textilera 
         Caption         =   "Textilera"
         Height          =   300
         Left            =   2505
         TabIndex        =   8
         Top             =   900
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   690
         TabIndex        =   0
         Top             =   255
         Width           =   1080
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2700
         TabIndex        =   1
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2385
         TabIndex        =   6
         Top             =   315
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   15
      TabIndex        =   3
      Top             =   330
      Width           =   4485
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmreporte_valuacion_facturas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4050
      Picture         =   "frmreporte_valuacion_facturas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "frmreporte_valuacion_facturas"
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
   Dim dl As Long                                 ' Valor devuelto por la función API
   Dim sAttributes As String                  ' Aributos
   Dim sDriver As String                       ' Nombre del controlador
   Dim sDescription As String                ' Descripción del DSN
   Dim sDsnName As String                  ' Nombre del DSN
   Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
   Const vbAPINull As Long = 0&                         ' Puntero NULL
   ' se elimina
   Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema
   x = 1
   If x = 0 Then
   If Me.opt_almacen = True Then
      cnn.Close
      'MsgBox var_conexion_string_distribucion
      cnn.Open var_conexion_string_distribucion
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
   End If

   If Me.opt_textilera = True Then
      
      cnn.Close
      sDsnName = "DSN=sqlsistema"
      sDriver = "SQL Server"
      dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)
      'se crea
      sDsnName = "sqlsistema"
      sDescription = "sqlsistema"
      sDriver = "SQL Server"
      sAttributes = "DSN=" & sDsnName & Chr(0)
      sAttributes = sAttributes & "Server=sqlquezada2" & Chr$(0)
      sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
      sAttributes = sAttributes & "Database=sidtextilera & Chr(0)"
      strAttributes = strAttributes & "UID=sa" & Chr$(0)
      strAttributes = strAttributes & "PWD=elia" & Chr$(0)
      dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
   End If
   End If
   'On Error GoTo salir:
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
             
             cnn.BeginTrans
             rs.Open "select max(inte_tvf_consecutivo) from tb_temp_valuacion_facturacion", cnn, adOpenDynamic, adLockOptimistic
             If Not rs.EOF Then
                var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
             Else
                var_consecutivo = 0
             End If
             var_consecutivo = var_consecutivo + 1
             rs.Close
             rs.Open "Insert into tb_temp_valuacion_facturacion (INTE_TVF_CONSECUTIVO, vcha_aud_usuario, vcha_aud_maquina) values (" + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
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
             
             'rs.Open "select * from tb_encabezado_cartera where dtim_Car_fecha >= " + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + "-.00001 and vcha_car_documento = 'FA'", cnn, adOpenDynamic, adLockOptimistic
             Text1 = "select " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-.001, VCHA_EMP_EMPRESA_ID ,VCHA_CAR_TIPO_DOCUMENTO , vcha_ser_serie_id , inte_car_numero, '" + var_clave_usuario_global + "', '" + fun_NombrePc + "' from tb_encabezado_cartera where dtim_Car_fecha >= " + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + " - 0.0001"
             var_cadena = " INSERT INTO TB_TEMP_VALUACION_FACTURACION (INTE_TVF_CONSECUTIVO, DTIM_TVF_FECHA_INICIO, DTIM_TVF_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA) select " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-.00001, VCHA_EMP_EMPRESA_ID ,VCHA_CAR_TIPO_DOCUMENTO , vcha_ser_serie_id , inte_car_numero, '" + var_clave_usuario_global + "', '" + fun_NombrePc + "' from tb_encabezado_cartera where dtim_Car_fecha >= " + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + " - 0.0001 AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'"
             Text1 = var_cadena
             'MsgBox cnn.ConnectionString
             rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
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
             var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
             'If Not rs.EOF Then
             '   While Not rs.EOF
             '      var_cadena = " INSERT INTO TB_TEMP_VALUACION_FACTURACION (INTE_TVF_CONSECUTIVO, DTIM_TVF_FECHA_INICIO, DTIM_TVF_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA)"
             '      var_cadena = var_cadena + "Values ( " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + ", '" + rs!VCHA_EMP_EMPRESA_ID + "','" + rs!VCHA_CAR_TIPO_DOCUMENTO + "', '" + rs!vcha_ser_serie_id + "', " + CStr(rs!inte_car_numero) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "')"
             '      rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
             '      rs.MoveNext
             '   Wend
             'End If
             
             
             'rs.Close
         Set reporte = appl.OpenReport(App.Path + "\rep_valuacion_facturacion.rpt")
         'reporte.RecordSelectionFormula = " {VW_VALUACION_FACTURACION.INTE_TVF_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_VALUACION_FACTURACION.VCHA_AUD_USUARIO} = '" + var_clave_usuario_global + "' and{VW_VALUACION_FACTURACION.VCHA_AUD_MAQUINA} = '" + fun_NombrePc + "'"
         reporte.RecordSelectionFormula = " {VW_VALUACION_FACTURACION.INTE_TVF_CONSECUTIVO} = " + CStr(var_consecutivo)
         'MsgBox CStr(var_consecutivo)
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Valuación de Facturas"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            Set reporte = appl.OpenReport(App.Path + "\rep_valuacion_facturacion.rpt")
            'reporte.RecordSelectionFormula = " {VW_VALUACION_FACTURACION.INTE_TVF_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_VALUACION_FACTURACION.VCHA_AUD_USUARIO} = '" + var_clave_usuario_global + "' and{VW_VALUACION_FACTURACION.VCHA_AUD_MAQUINA} = '" + fun_NombrePc + "'"
            reporte.RecordSelectionFormula = " {VW_VALUACION_FACTURACION.INTE_TVF_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\reporte_valuacion_facturas" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
         End If
         
         rs.Open "delete from TB_TEMP_VALUACION_FACTURACION where INTE_TVF_CONSECUTIVO = " + CStr(var_consecutivo) + " and VCHA_AUD_USUARIO = '" + var_clave_usuario_global + "' and VCHA_AUD_MAQUINA = '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
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
   Dim dl As Long                                 ' Valor devuelto por la función API
   Dim sAttributes As String                  ' Aributos
   Dim sDriver As String                       ' Nombre del controlador
   Dim sDescription As String                ' Descripción del DSN
   Dim sDsnName As String                  ' Nombre del DSN
   Me.opt_almacen = True
   'cnn.Close
   'MsgBox var_conexion_string_distribucion
   'cnn.Open var_conexion_string_distribucion
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
   Top = 3000
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro = False
   Call activa_forma(var_activa_forma_reporte_valuacion_facturas)
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(txt_fin) Then
         frmcalendario.mes.Value = CDate(txt_fin)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_fin_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(txt_inicio) Then
         frmcalendario.mes.Value = CDate(txt_inicio)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub
