VERSION 5.00
Begin VB.Form frmrelacion_facturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relación de Facturas"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   3975
      Begin VB.TextBox txt_embarque_relacion 
         Height          =   315
         Left            =   1950
         TabIndex        =   1
         Top             =   345
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número de Embarque:"
         Height          =   195
         Left            =   165
         TabIndex        =   2
         Top             =   405
         Width           =   1590
      End
   End
End
Attribute VB_Name = "frmrelacion_facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub Form_Load()
   Top = 3000
   Left = 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_relacion_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      rs.Open "select * from vw_facturas_distintas where VCHA_EMP_EMPRESA_ID ='" + var_empresa + "' and inte_emb_embarque = " + txt_embarque_relacion, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open "select * from tb_inventario_documentos where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and VCHA_CAR_TIPO_DOCUMENTO = '" + rs!vcha_Car_tipo_documento + "' and VCHA_CAR_DOCUMENTO = '" + rs!vcha_Car_documento + "' and VCHA_AGE_AGENTE_ID = '" + rs!VCHA_AGE_AGENTE_ID + "' and VCHA_CAR_CLASE_ID = '" + rs!vcha_Car_clase_id + "' and INTE_CAR_NUMERO = " + CStr(rs!inte_Car_numero) + " and  VCHA_SER_SERIE_ID = '" + rs!VCHA_SER_SERIE_ID + "'", cnn, adOpenDynamic, adLockOptimistic
            If rsaux.EOF Then
               var_cadena = "INSERT INTO [TB_INVENTARIO_DOCUMENTOS] ([VCHA_EMP_EMPRESA_ID],[VCHA_AGE_AGENTE_ID],[VCHA_CAR_TIPO_DOCUMENTO], [VCHA_CAR_DOCUMENTO],  [VCHA_CAR_CLASE_ID], [INTE_CAR_NUMERO], [CHAR_CAR_AFECTACION], [VCHA_SER_SERIE_ID], [CHAR_IDO_ESTATUS], [FLOA_IDO_CANTIDAD], [FLOA_CAR_IMPORTE_NETO], [FLOA_CAR_TIPO_CAMBIO], [VCHA_MON_MONEDA_ID],[DTIM_IDO_FECHA_ENTRAGA],[VCHA_CLI_CLAVE_ID], [INTE_EMB_EMBARQUE])"
               var_cadena = var_cadena + " Values ( '" + var_empresa + "', '" + rs!VCHA_AGE_AGENTE_ID + "', '" + rs!vcha_Car_tipo_documento + "', '" + rs!vcha_Car_documento + "', '" + rs!vcha_Car_clase_id + "', " + CStr(rs!inte_Car_numero) + ", '+', '" + rs!VCHA_SER_SERIE_ID + "',  'A', " + CStr(rs!Cantidad) + ", " + CStr(rs!floa_Car_importe_neto) + ", " + CStr(rs!floa_car_tipo_cambio) + ", '" + rs!vcha_mon_moneda_id + "', '" + CStr(Date) + "', '" + rs!vcha_cli_clave_id + "', " + txt_embarque_relacion + ")"
               
               rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            End If
            rsaux.Close
            rs.MoveNext
         Wend
         rsaux4.Open "select distinct vcha_age_agente_id from tb_inventario_documentos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_embarque_relacion, cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux4.EOF
            Set reporte = appl.OpenReport(App.Path + "\rep_inventario_documentos_embarque.rpt")
            var_cadena = "{VW_INVENTARIO_DOCUMENTOS.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_INVENTARIO_DOCUMENTOS.inte_emb_embarque} = " + txt_embarque_relacion + " and {VW_INVENTARIO_DOCUMENTOS.vcha_age_agente_id} = '" + rsaux4!VCHA_AGE_AGENTE_ID + "'"
            reporte.RecordSelectionFormula = var_cadena
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
                  
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_inventario_documentos_embarque_2.rpt")
               var_cadena = "{VW_INVENTARIO_DOCUMENTOS.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_INVENTARIO_DOCUMENTOS.inte_emb_embarque} = " + txt_embarque_relacion + " and {VW_INVENTARIO_DOCUMENTOS.vcha_age_agente_id} = '" + rsaux4!VCHA_AGE_AGENTE_ID + "'"
               reporte.RecordSelectionFormula = var_cadena
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\relacion_facturas_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            
            
            rsaux4.MoveNext
         Wend
         rsaux4.Close
      Else
         MsgBox "No existen facturas en el embarque indicado", vbOKOnly, "ATENCION"
      End If
      rs.Close
      'frm_embarque_relacion.Visible = False
   End If
   If keyasii = 27 Then
      Unload Me
   End If
End Sub
