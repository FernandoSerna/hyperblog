VERSION 5.00
Begin VB.Form frmreporte_detalle_facturacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Facturación"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   135
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
      Left            =   4050
      Picture         =   "frmreporte_detalle_facturacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmreporte_detalle_facturacion.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   45
      TabIndex        =   7
      Top             =   270
      Width           =   4395
   End
End
Attribute VB_Name = "frmreporte_detalle_facturacion"
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
            rs.Open "select max(isnull(INTE_TEM_CONSECUTIVO,0)) as numero from TB_TEMP_REPORTE_DETALLE_FACTURACION", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rsaux.Open "INSERT INTO TB_TEMP_REPORTE_DETALLE_FACTURACION (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTiM_TEM_FECHA_FIN) VALUES (" + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            
            var_cadena = "INSERT INTO TB_TEMP_REPORTE_DETALLE_FACTURACION (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_SER_SERIE_ID, VCHA_CAR_DOCUMENTO, INTE_CAR_NUMERO, FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, VCHA_ART_ARTICULO_ID,"
            var_cadena = var_cadena + " FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_ART_NOMBRE_ESPAÑOL, DTIM_CAR_FECHA, VCHA_LIN_LINEA_ID, VCHA_LIN_MOMBRE, floa_car_tipo_cambio, VCHA_CAT_CATALOGO_ID, VCHA_CAT_NOMBRE, vcha_Emo_referencia) SELECT " + CStr(var_consecutivo) + ", " + var_fecha_inicio + "," + var_fecha_fin + ", VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_SER_SERIE_ID, VCHA_CAR_DOCUMENTO, INTE_CAR_NUMERO, FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, VCHA_ART_ARTICULO_ID,"
            var_cadena = var_cadena + " FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, case (1 - (isnull(FLOA_SAL_PROMOCION_1,0)/100)) when 0 then 0 else ((FLOA_SAL_PRECIO/(1 - (isnull(FLOA_SAL_PROMOCION_1,0)/100))))/ (1 - (isnull(floa_sal_promocion_2,0)/100)) end, FLOA_SAL_PROMOCION_1, isnull(FLOA_SAL_PROMOCION_2,0), isnull(FLOA_SAL_DESCUENTO_1,0), isnull(FLOA_SAL_DESCUENTO_2,0), VCHA_ART_NOMBRE_ESPAÑOL, DTIM_CAR_FECHA, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, isnull(floa_CAR_tipo_cambio,1), VCHA_CAT_CATALOGO_ID, VCHA_CAT_NOMBRE, vcha_emo_referencia FROM VW_DETALLE_FACTURACION WHERE DTIM_CAR_FECHA >= " + var_fecha_inicio + " AND DTIM_CAR_FECHA <= " + var_fecha_fin_2 + " and vcha_emp_empresa_id = '" + var_empresa + "' ORDER BY INTE_CAR_NUMERO, VCHA_LIN_LINEA_ID"
            cnn.CommandTimeout = 6000
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         
         
            Set reporte = appl.OpenReport(App.Path + "\rep_detalle_facturacion.rpt")
            reporte.RecordSelectionFormula = "{tb_temp_reporte_detalle_facturacion.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo)
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\detalle_facturacion_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a generado el reporte " + archivo, vbOKOnly, "ATENCION"
            rs.Open "delete from TB_TEMP_REPORTE_DETALLE_FACTURACION where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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


