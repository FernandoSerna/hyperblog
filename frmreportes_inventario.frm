VERSION 5.00
Begin VB.Form frmreportes_inventario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes de inventario"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2490
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4410
      Begin VB.CommandButton cmd_diferencias 
         Caption         =   "Artículos con mayor número de diferencias"
         Height          =   645
         Left            =   105
         TabIndex        =   4
         Top             =   1680
         Width           =   4200
      End
      Begin VB.CommandButton cmd_contado 
         Caption         =   "Cuanto va contado"
         Height          =   645
         Left            =   105
         TabIndex        =   3
         Top             =   1680
         Visible         =   0   'False
         Width           =   4200
      End
      Begin VB.CommandButton cmd_ordenes 
         Caption         =   "Mercancia en ordenes de surtido"
         Height          =   645
         Left            =   105
         TabIndex        =   2
         Top             =   945
         Width           =   4200
      End
      Begin VB.CommandButton cmd_existencias 
         Caption         =   "Existencias"
         Height          =   645
         Left            =   105
         TabIndex        =   1
         Top             =   210
         Width           =   4200
      End
   End
End
Attribute VB_Name = "frmreportes_inventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_contado_Click()
   Dim var_cadena_agentes As String
   Dim var_n As Integer, var_i As Integer, var_contador_agentes As Integer, var_consecutivo As Integer
   cnn.CommandTimeout = 6000
        
   Set reporte = appl.OpenReport(App.Path + "\rep_inventario_contado.rpt")
   reporte.RecordSelectionFormula = "{VW_INVENTARIO_cONTADO.CANTIDAD}>0"
   frmvistasprevias.cr.ReportSource = reporte
   For ntablas = 1 To reporte.Database.Tables.Count
       reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   Next ntablas
   frmvistasprevias.cr.ViewReport
   frmvistasprevias.Caption = "Artículos contados en el inventario"
   frmvistasprevias.Show 1
   Set reporte = Nothing

   var_si = MsgBox("¿Desea importar el reporte?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      Set reporte = appl.OpenReport(App.Path + "\rep_inventario_contado.rpt")
      reporte.RecordSelectionFormula = "{VW_INVENTARIO_cONTADO.CANTIDAD}>0"
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      reporte.ExportOptions.FormatType = crEFTExcel80
      reporte.ExportOptions.DestinationType = crEDTDiskFile
      archivo = "c:\reportessid\INVENTARIO_CONTADO_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
      reporte.ExportOptions.DiskFileName = archivo
      reporte.Export False
      Set reporte = Nothing
      MsgBox "Se a terminado de guardar el archivo " + archivo
   End If
End Sub

Private Sub cmd_diferencias_Click()
   Dim var_cadena_agentes As String
   Dim var_n As Integer, var_i As Integer, var_contador_agentes As Integer, var_consecutivo As Integer
   cnn.CommandTimeout = 6000
        
   Set reporte = appl.OpenReport(App.Path + "\rep_inventario_diferencias.rpt")
   For ntablas = 1 To reporte.Database.Tables.Count
       reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   Next ntablas
   reporte.ExportOptions.FormatType = crEFTExcel80
   reporte.ExportOptions.DestinationType = crEDTDiskFile
   archivo = "c:\reportessid\INVENTARIO_DIFERENCIAS_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
   reporte.ExportOptions.DiskFileName = archivo
   reporte.Export False
   Set reporte = Nothing
   MsgBox "Se a terminado de guardar el archivo " + archivo
End Sub

Private Sub cmd_existencias_Click()
   Set reporte = appl.OpenReport(App.Path + "\rep_inventario_existencias.rpt")
   frmvistasprevias.cr.ReportSource = reporte
   For ntablas = 1 To reporte.Database.Tables.Count
       reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   Next ntablas
   frmvistasprevias.cr.ViewReport
   frmvistasprevias.Caption = "Reporte de existencias"
   frmvistasprevias.Show 1
   Set reporte = Nothing
   var_si = MsgBox("¿Desea importar el reporte?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      Set reporte = appl.OpenReport(App.Path + "\rep_inventario_existencias.rpt")
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      reporte.ExportOptions.FormatType = crEFTExcel80
      reporte.ExportOptions.DestinationType = crEDTDiskFile
      archivo = "c:\reportessid\REPORTE_EXISTENCIAS_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
      reporte.ExportOptions.DiskFileName = archivo
      reporte.Export False
      Set reporte = Nothing
      MsgBox "Se a terminado de guardar el archivo " + archivo
   End If
End Sub

Private Sub cmd_ordenes_Click()
   Dim var_cadena_agentes As String
   Dim var_n As Integer, var_i As Integer, var_contador_agentes As Integer, var_consecutivo As Integer
   cnn.CommandTimeout = 6000
   var_n = 0
   rs.Open "select * from VW_ORDENES_SURTIDO_PENDIENTES ", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      cnn.BeginTrans
      If rsaux.State = 1 Then
         rsaux.Close
      End If
      rsaux.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_ORDENES_SURTIDO_PENDIENTES", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         var_consecutivo = IIf(IsNull(rsaux!NUMERO), 0, rsaux!NUMERO)
      Else
         var_consecutivo = 0
      End If
      rsaux.Close
      var_consecutivo = var_consecutivo + 1
      rsaux.Open "insert into TB_TEMP_ORDENES_SURTIDO_PENDIENTES (INTE_TEM_CONSECUTIVO) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
      cnn.CommitTrans
      While Not rs.EOF
            rsaux.Open "INSERT INTO TB_TEMP_ORDENES_SURTIDO_PENDIENTES (INTE_TEM_CONSECUTIVO, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, INTE_ORS_ORDEN_SURTIDO, VCHA_AGE_AGENTE_ID, DTIM_TEM_FECHA) values  (" + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', '" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', " + CStr(rs!INTE_ORS_ORDEN_SURTIDO) + ", '" + rs!VCHA_AGE_AGENTE_ID + "',GETDATE())", cnn, adOpenDynamic, adLockOptimistic
            rs.MoveNext
      Wend
          
      Set reporte = appl.OpenReport(App.Path + "\rep_inventario_ordenes_pendientes.rpt")
      reporte.RecordSelectionFormula = "{VW_INVENTARIO_ORDENES_PENDIENTES.INTE_TEM_CONSECUTIVO} = " + Str(var_consecutivo) + " AND ({VW_INVENTARIO_ORDENES_PENDIENTES.VCHA_EMP_EMPRESA_ID} = '02' OR {VW_INVENTARIO_ORDENES_PENDIENTES.VCHA_EMP_EMPRESA_ID} = '03')"
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Ordenes de surtido pendientes de empacar o facturar"
      frmvistasprevias.Show 1
      Set reporte = Nothing
   
      var_si = MsgBox("¿Desea importar el reporte?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         Set reporte = appl.OpenReport(App.Path + "\rep_inventario_ordenes_pendientes.rpt")
         reporte.RecordSelectionFormula = "{VW_INVENTARIO_ORDENES_PENDIENTES.INTE_TEM_CONSECUTIVO} = " + Str(var_consecutivo) + " AND ({VW_INVENTARIO_ORDENES_PENDIENTES.VCHA_EMP_EMPRESA_ID} = '02' OR {VW_INVENTARIO_ORDENES_PENDIENTES.VCHA_EMP_EMPRESA_ID} = '03')"
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reportessid\ARTICULOS_PENDIENTES_SURTIR_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
         MsgBox "Se a terminado de guardar el archivo " + archivo
      End If
      rsaux.Open "DELETE FROM TB_TEMP_ORDENES_SURTIDO_PENDIENTES WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   
   Else
      MsgBox "No existen ordenes de surtido pendientes para los agentes seleccionados", vbOKOnly, "ATENCION"
   End If
   rs.Close
End Sub

Private Sub Form_Load()
   Top = 2000
   Left = 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub
