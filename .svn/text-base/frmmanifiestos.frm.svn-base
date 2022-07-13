VERSION 5.00
Begin VB.Form frmmanifiesto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manifiesto"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   " Paqueteria "
      Height          =   840
      Left            =   60
      TabIndex        =   9
      Top             =   465
      Width           =   4335
      Begin VB.ComboBox cmb_paqueteria 
         Height          =   315
         ItemData        =   "frmmanifiestos.frx":0000
         Left            =   105
         List            =   "frmmanifiestos.frx":001F
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   330
         Width           =   4110
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   60
      TabIndex        =   6
      Top             =   1425
      Width           =   4335
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   690
         TabIndex        =   1
         Top             =   255
         Width           =   1080
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2700
         TabIndex        =   2
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2385
         TabIndex        =   7
         Top             =   315
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   4
      Top             =   330
      Width           =   4485
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmmanifiestos.frx":0099
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4035
      Picture         =   "frmmanifiestos.frx":019B
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "frmmanifiesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmb_paqueteria_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

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
            If Trim(Me.cmb_paqueteria) <> "" Then
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
               cnn.CommandTimeout = 360000000
               cnn.BeginTrans
               rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_MANIFIESTO", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
               Else
                  var_consecutivo = 0
               End If
               var_consecutivo = var_consecutivo + 1
               rs.Close
               rs.Open "Insert into TB_TEMP_MANIFIESTO (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               var_nombre_paqueteria = ""
               If Me.cmb_paqueteria = "ESTAFETA" Then
                  var_paqueteria = "(VCHA_PAQ_CLAVE_ID = '001' OR VCHA_PAQ_CLAVE_ID = '004')"
                  var_nombre_paqueteria = "ESTAFETA"
               End If
               If Me.cmb_paqueteria = "MULTIPACK" Then
                  var_paqueteria = "(VCHA_PAQ_CLAVE_ID = '002' OR VCHA_PAQ_CLAVE_ID = '005')"
                  var_nombre_paqueteria = "MULTIPACK"
               End If
               If Me.cmb_paqueteria = "OMNIBUS" Then
                  var_paqueteria = "(VCHA_PAQ_CLAVE_ID = '003' OR VCHA_PAQ_CLAVE_ID = '006')"
                  var_nombre_paqueteria = "OMNIBUS"
               End If
               If Me.cmb_paqueteria = "AEROFLASH" Then
                  var_paqueteria = "(VCHA_PAQ_CLAVE_ID = '007' OR VCHA_PAQ_CLAVE_ID = '008')"
                  var_nombre_paqueteria = "AEROFLASH"
               End If
               
               If Me.cmb_paqueteria = "ESTRELLA BLANCA" Then
                  var_paqueteria = "(VCHA_PAQ_CLAVE_ID = '013' OR VCHA_PAQ_CLAVE_ID = '016')"
                  var_nombre_paqueteria = "ESTRELLA BLANCA"
               End If
               
               If Me.cmb_paqueteria = "QUALITY POST" Then
                  var_paqueteria = "(VCHA_PAQ_CLAVE_ID = '010' OR VCHA_PAQ_CLAVE_ID = '011')"
                  var_nombre_paqueteria = "QUALITY POST"
               End If
               
               If Me.cmb_paqueteria = "MENSAJERIA EXPRESS" Then
                  var_paqueteria = "(VCHA_PAQ_CLAVE_ID = '017' OR VCHA_PAQ_CLAVE_ID = '015')"
                  var_nombre_paqueteria = "MENSAJERIA EXPRESS"
               End If
               
               If Me.cmb_paqueteria = "PROPIA SIN COSTO" Then
                  var_paqueteria = "(VCHA_PAQ_CLAVE_ID = '009')"
                  var_nombre_paqueteria = "PROPIA SIN COSTO"
               End If
               
               If Me.cmb_paqueteria = "CARSSA" Then
                  var_paqueteria = "(VCHA_PAQ_CLAVE_ID = '019')"
                  var_nombre_paqueteria = "CARSSA"
               End If
               
               
               'var_cadena = "INSERT INTO TB_TEMP_MANIFIESTO (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, TOTAL_CAJAS, INTE_ORS_ORDEN_SURTIDO, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_PAI_NOMBRE, VCHA_EST_NOMBRE, VCHA_CIU_NOMBRE, VCHA_CLI_DIRECCION, VCHA_COL_NOMBRE, FLOA_CAR_IMPORTE_NETO, IMPORTE_SEGURO, IMPORTE_PAQUETERRIA, FLOA_PAQ_COSTO_REFERENCIA, DTIM_EMB_FECHA_INICIO, CHAR_EMB_ESTATUS, VCHA_CLI_CP, VCHA_TEM_PAQUETERIA, VCHA_CLI_TELEFONO, INTE_PED_NUMERO, VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE) "
               'var_cadena = var_cadena + " SELECT  " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-.000001, TOTAL_CAJAS, INTE_ORS_ORDEN_SURTIDO, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_PAI_NOMBRE, VCHA_EST_NOMBRE, VCHA_CIU_NOMBRE, VCHA_CLI_DIRECCION, VCHA_COL_NOMBRE, sum(FLOA_CAR_IMPORTE_NETO) as FLOA_CAR_IMPORTE_NETO, IMPORTE_SEGURO, IMPORTE_PAQUETERIA, FLOA_PAQ_COSTO_REFERENCIA, DTIM_EMB_FECHA_INICIO, CHAR_EMB_ESTATUS, VCHA_CLI_CP, '" + var_nombre_paqueteria + "', vcha_cli_telefono, INTE_PED_NUMERO, VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE  FROM VW_MANIFIESTO WHERE DTIM_EMB_FECHA_INICIO >= " + var_fecha_inicio + " AND DTIM_EMB_FECHA_INICIO <= " + var_fecha_fin + " -.00001 AND " + var_paqueteria + " group by TOTAL_CAJAS, INTE_ORS_ORDEN_SURTIDO, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE,"
               'var_cadena = var_cadena + " VCHA_PAI_NOMBRE, VCHA_EST_NOMBRE, VCHA_CIU_NOMBRE, VCHA_CLI_DIRECCION, VCHA_COL_NOMBRE, IMPORTE_SEGURO, IMPORTE_PAQUETERIA, FLOA_PAQ_COSTO_REFERENCIA, DTIM_EMB_FECHA_INICIO, CHAR_EMB_ESTATUS, VCHA_CLI_CP, VCHA_CLI_TELEFONO, INTE_PED_NUMERO, VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE"
               
               var_cadena = "INSERT INTO TB_TEMP_MANIFIESTO (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, TOTAL_CAJAS, INTE_ORS_ORDEN_SURTIDO, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_PAI_NOMBRE, VCHA_EST_NOMBRE, VCHA_CIU_NOMBRE, VCHA_CLI_DIRECCION, VCHA_COL_NOMBRE, FLOA_CAR_IMPORTE_NETO, DTIM_EMB_FECHA_INICIO, CHAR_EMB_ESTATUS, VCHA_CLI_CP, VCHA_TEM_PAQUETERIA, VCHA_CLI_TELEFONO, INTE_PED_NUMERO, VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE) "
               var_cadena = var_cadena + " SELECT  " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-.000001, TOTAL_CAJAS, INTE_ORS_ORDEN_SURTIDO, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_PAI_NOMBRE, VCHA_EST_NOMBRE, VCHA_CIU_NOMBRE, VCHA_CLI_DIRECCION, VCHA_COL_NOMBRE, sum(FLOA_CAR_IMPORTE_NETO) as FLOA_CAR_IMPORTE_NETO,  DTIM_EMB_FECHA_INICIO, CHAR_EMB_ESTATUS, VCHA_CLI_CP, '" + var_nombre_paqueteria + "', vcha_cli_telefono, INTE_PED_NUMERO, VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE  FROM VW_MANIFIESTO WHERE DTIM_EMB_FECHA_INICIO >= " + var_fecha_inicio + " AND DTIM_EMB_FECHA_INICIO <= " + var_fecha_fin + " -.00001 AND " + var_paqueteria + " group by TOTAL_CAJAS, INTE_ORS_ORDEN_SURTIDO, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE,"
               var_cadena = var_cadena + " VCHA_PAI_NOMBRE, VCHA_EST_NOMBRE, VCHA_CIU_NOMBRE, VCHA_CLI_DIRECCION, VCHA_COL_NOMBRE, DTIM_EMB_FECHA_INICIO, CHAR_EMB_ESTATUS, VCHA_CLI_CP, VCHA_CLI_TELEFONO, INTE_PED_NUMERO, VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE"
               
               
               cnn.CommandTimeout = 360
               'AQUI FALLA
               rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               rs.Open "select inte_ped_numero from TB_TEMP_MANIFIESTO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and inte_ped_numero is not null", cnn, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     If rsaux1.State = 1 Then
                        rsaux1.Close
                     End If
                     rsaux1.Open "select * from vw_importes_seguro_paqueteria where inte_ped_numero = " + CStr(rs!inte_ped_numero), cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux1.EOF Then
                        rsaux2.Open "update TB_TEMP_MANIFIESTO set IMPORTE_SEGURO = " + CStr(IIf(IsNull(rsaux1!importe_seguro), 0, rsaux1!importe_seguro)) + ", IMPORTE_PAQUETERRIA = " + CStr(IIf(IsNull(rsaux1!importe_paqueteria), 0, rsaux1!importe_paqueteria)) + ", FLOA_PAQ_COSTO_REFERENCIA = " + CStr(IIf(IsNull(rsaux1!floa_paq_costo_referencia), 0, rsaux1!floa_paq_costo_referencia)) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " And inte_ped_numero = " + CStr(rs!inte_ped_numero), cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux1.Close
                     rs.MoveNext
               Wend
               rs.Close
               If rsaux9.State = 1 Then
                  rsaux9.Close
               End If
               rsaux9.Open "select  INTE_ORS_ORDEN_SURTIDO from tb_temp_manifiesto where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and inte_ors_orden_surtido is not null", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux9.EOF
                     var_cadena = "SELECT dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO,dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID FROM         dbo.TB_ENCABEZADO_CARTERA INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND"
                     var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO"
                     var_cadena = var_cadena + " WHERE     (dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID = 'FT') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = " + CStr(rsaux9!INTE_ORS_ORDEN_SURTIDO) + ")"
                     var_facturas = ""
                     rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     While Not rsaux10.EOF
                           If var_facturas = "" Then
                              var_facturas = CStr(IIf(IsNull(rsaux10!inte_Car_numero), "", rsaux10!inte_Car_numero))
                           Else
                              var_facturas = var_facturas + ", " + CStr(IIf(IsNull(rsaux10!inte_Car_numero), "", rsaux10!inte_Car_numero))
                           End If
                           rsaux10.MoveNext
                     Wend
                     rsaux10.Close
                     var_cadena = "SELECT     TOP 100 PERCENT dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN, dbo.TB_SELLOS.VCHA_SEL_SELLO, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID FROM dbo.TB_SELLOS INNER JOIN dbo.TB_DETALLE_EMBARQUES ON dbo.TB_SELLOS.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID AND dbo.TB_SELLOS.INTE_EMB_EMBARQUE = dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND"
                     var_cadena = var_cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND    dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID wHERE     (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'FT') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = " + CStr(rsaux9!INTE_ORS_ORDEN_SURTIDO) + ") ORDER BY dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN DESC"
                     rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     var_sellos = ""
                     While Not rsaux10.EOF
                           If var_sellos = "" Then
                              var_sellos = IIf(IsNull(rsaux10!vcha_sel_Sello), "", rsaux10!vcha_sel_Sello)
                           Else
                              var_sellos = var_sellos + ", " + IIf(IsNull(rsaux10!vcha_sel_Sello), "", rsaux10!vcha_sel_Sello)
                           End If
                           rsaux10.MoveNext
                     Wend
                     rsaux10.Close
                     rsaux10.Open "update tb_temp_manifiesto set vcha_tem_facturas = '" + var_facturas + "', vcha_tem_guias = '" + var_sellos + "' where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and inte_ors_orden_surtido = " + CStr(rsaux9!INTE_ORS_ORDEN_SURTIDO), cnn, adOpenDynamic, adLockOptimistic
                     rsaux9.MoveNext
                Wend
               rsaux9.Close
               
               
               Set reporte = appl.OpenReport(App.Path + "\REP_MANIFIESTO.rpt")
               reporte.RecordSelectionFormula = "{TB_TEMP_MANIFIESTO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {TB_TEMP_MANIFIESTO.INTE_ORS_ORDEN_SURTIDO} >0"
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte de existencias en cajas"
               frmvistasprevias.Show 1
               Set reporte = Nothing
           
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  Set reporte = appl.OpenReport(App.Path + "\REP_MANIFIESTO.rpt")
                  reporte.RecordSelectionFormula = "{TB_TEMP_MANIFIESTO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {TB_TEMP_MANIFIESTO.INTE_ORS_ORDEN_SURTIDO} >0"
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\manifiesto_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
               rs.Open "delete from TB_TEMP_MANIFIESTO where INTE_Tem_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            Else
               MsgBox "No se a seleccionado alguna paqueteria", vbOKOnly, "ATENCION"
            End If
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
   Top = 2700
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro = False
   Call activa_forma(var_activa_forma_existencias_generales)
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



