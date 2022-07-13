VERSION 5.00
Begin VB.Form frmpacking_list 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packing List"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   2760
   Begin VB.Frame Frame2 
      Caption         =   " Embarque "
      Height          =   750
      Left            =   60
      TabIndex        =   5
      Top             =   480
      Width           =   2640
      Begin VB.TextBox txt_embarque 
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   2370
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2355
      Picture         =   "frmpacking_list.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmpacking_list.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmpacking_list.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   15
      TabIndex        =   0
      Top             =   345
      Width           =   2715
   End
End
Attribute VB_Name = "frmpacking_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   If IsNumeric(txt_embarque) Then
      cnn.BeginTrans
      rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_DETALLE_CAJAS", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
      Else
         var_consecutivo = 1
      End If
      rs.Close
      rs.Open "insert into tb_temp_detalle_cajas (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
      cnn.CommitTrans
      cnn.CommandTimeout = 600
      var_cadena = "INSERT INTO TB_TEMP_DETALLE_CAJAS (INTE_TEM_CONSECUTIVO, INTE_ORS_ORDEN_SURTIDO, INTE_PAQ_CAJA, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_ART_ARTICULO_ID, FLOA_PAQ_CANTIDAD, CHAR_PAQ_ESTATUS, VCHA_PAQ_MOVIMIENTO_DESTINO, INTE_PAQ_NUMERO_DESTINO, FLOA_PAQ_COSTO, FLOA_PAQ_PRECIO, INTE_EMB_EMBARQUE, CHAR_PED_TIPO, FLOA_PAQ_DESCUENTO_1, FLOA_PAQ_DESCUENTO_2) "
      'var_cadena = var_cadena + " select " + CStr(var_consecutivo) + ", INTE_ORS_ORDEN_SURTIDO, INTE_PAQ_CAJA, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_ART_ARTICULO_ID, FLOA_PAQ_CANTIDAD, CHAR_PAQ_ESTATUS, VCHA_PAQ_MOVIMIENTO_DESTINO, INTE_PAQ_NUMERO_DESTINO, FLOA_PAQ_COSTO, FLOA_PAQ_PRECIO, INTE_EMB_EMBARQUE, CHAR_PED_TIPO from tb_Detalle_cajas where inte_emb_embarque = " + Me.txt_embarque + " and vcha_emp_empresa_id = '" + var_empresa + "'"
      var_cadena = var_cadena + " SELECT " + CStr(var_consecutivo) + ", dbo.TB_DETALLE_CAJAS.INTE_ORS_ORDEN_SURTIDO, dbo.TB_DETALLE_CAJAS.INTE_PAQ_CAJA,dbo.TB_DETALLE_CAJAS.VCHA_EMP_EMPRESA_ID, dbo.TB_DETALLE_CAJAS.VCHA_UOR_UNIDAD_ID,dbo.TB_DETALLE_CAJAS.VCHA_ALM_ALMACEN_ID, dbo.TB_DETALLE_CAJAS.VCHA_ART_ARTICULO_ID, dbo.TB_DETALLE_CAJAS.FLOA_PAQ_CANTIDAD, dbo.TB_DETALLE_CAJAS.CHAR_PAQ_ESTATUS, dbo.TB_DETALLE_CAJAS.VCHA_PAQ_MOVIMIENTO_DESTINO, dbo.TB_DETALLE_CAJAS.INTE_PAQ_NUMERO_DESTINO, dbo.TB_DETALLE_CAJAS.FLOA_PAQ_COSTO, dbo.TB_DETALLE_CAJAS.FLOA_PAQ_PRECIO, dbo.TB_DETALLE_CAJAS.INTE_EMB_EMBARQUE, dbo.TB_DETALLE_CAJAS.CHAR_PED_TIPO, dbo.TB_ENC_ORDEN_SURTIDO.FLOA_ORS_DESCUENTO_1, dbo.TB_ENC_ORDEN_SURTIDO.FLOA_ORS_DESCUENTO_2 FROM dbo.TB_DETALLE_CAJAS INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO ON dbo.TB_DETALLE_CAJAS.INTE_ORS_ORDEN_SURTIDO = dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO "
      var_cadena = var_cadena + " WHERE  (dbo.TB_DETALLE_CAJAS.INTE_EMB_EMBARQUE = " + Me.txt_embarque + ") AND (dbo.TB_DETALLE_CAJAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')"
      rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
      
      var_cadena = "INSERT INTO TB_TEMP_PACKING_LIST_SALIDAS (INTE_TEM_CONSECUTIVO, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID,  VCHA_ART_ARTICULO_ID, FLOA_SAL_PRECIO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, INTE_EMB_EMBARQUE, FLOA_AGR_FRACCION_ARANCELARIA) "
      var_cadena = var_cadena + " SELECT " + CStr(var_consecutivo) + ", VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID,  VCHA_ART_ARTICULO_ID, FLOA_SAL_PRECIO/floa_sal_cantidad, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, INTE_EMB_EMBARQUE, FLOA_AGR_FRACCION_ARANCELARIA FROM VW_PACKING_LIST_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque
      rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
      
      rs.Open "select * from vw_packing_list_TEMPORAL where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_emb_embarque = " + txt_embarque + " and floa_paq_cantidad > 0 AND INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         If var_empresa = "03" Then
            Set reporte = appl.OpenReport(App.Path + "\rep_PACKING_LIST_TEMPORAL_EXPORTACIONES.rpt")
            reporte.RecordSelectionFormula = "{VW_PACKING_LIST_temporal.inte_EMB_EMBARQUE} = " + txt_embarque + " and {VW_PACKING_LIST_temporal.VCHA_EMP_EMPRESA_ID} ='" + var_empresa + "' and {VW_PACKING_LIST_temporal.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' and {VW_PACKING_LIST_temporal.floa_paq_cantidad} > 0 and {VW_PACKING_LIST_temporal.inte_tem_consecutivo} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Packing List"
            frmvistasprevias.Show 1
            Set reporte = Nothing
      
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_PACKING_LIST_TEMPORAL_eXPORTACIONES.rpt")
               reporte.RecordSelectionFormula = "{VW_PACKING_LIST_temporal.inte_EMB_EMBARQUE} = " + txt_embarque + " and {VW_PACKING_LIST_temporal.VCHA_EMP_EMPRESA_ID} ='" + var_empresa + "' and {VW_PACKING_LIST_temporal.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' and {VW_PACKING_LIST_temporal.floa_paq_cantidad} > 0 and {VW_PACKING_LIST_temporal.inte_tem_consecutivo} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\packing_list" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
         Else
            rsaux2.Open "select * from tb_encabezado_embarques where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_agente_packing_list = rsaux2!VCHA_AGE_AGENTE_ID
            End If
            rsaux2.Close
            
            rsaux.Open "select * from tb_Agentes where vcha_age_agente_id = '" + var_agente_packing_list + "'", cnn, adOpenDynamic, adLockOptimistic
            var_tipo_agente = IIf(IsNull(rsaux!vcha_tag_tipoagente_id), "", rsaux!vcha_tag_tipoagente_id)
            If var_tipo_agente = "TDA" Then
               var_nombre_reporte = "packing_list"
               var_embarque_packing_list = CDbl(Me.txt_embarque)
               var_consecutivo_packing_list = var_consecutivo
            Else
               var_nombre_reporte = ""
            End If
            rsaux.Close
            
            
            Set reporte = appl.OpenReport(App.Path + "\rep_PACKING_LIST_TEMPORAL.rpt")
            reporte.RecordSelectionFormula = "{VW_PACKING_LIST_temporal.inte_EMB_EMBARQUE} = " + txt_embarque + " and {VW_PACKING_LIST_temporal.VCHA_EMP_EMPRESA_ID} ='" + var_empresa + "' and {VW_PACKING_LIST_temporal.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' and {VW_PACKING_LIST_temporal.floa_paq_cantidad} > 0 and {VW_PACKING_LIST_temporal.inte_tem_consecutivo} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Packing List"
            frmvistasprevias.Show 1
            Set reporte = Nothing
      
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_PACKING_LIST_TEMPORAL.rpt")
               reporte.RecordSelectionFormula = "{VW_PACKING_LIST_temporal.inte_EMB_EMBARQUE} = " + txt_embarque + " and {VW_PACKING_LIST_temporal.VCHA_EMP_EMPRESA_ID} ='" + var_empresa + "' and {VW_PACKING_LIST_temporal.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' and {VW_PACKING_LIST_temporal.floa_paq_cantidad} > 0 and {VW_PACKING_LIST_temporal.inte_tem_consecutivo} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\packing_list" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
         End If
      
      Else
         MsgBox "El embarque no existe o no a hay paquetes en el", vbOKOnly, "ATENCION"
      End If
      rs.Close
      rs.Open "delete from TB_TEMP_DETALLE_CAJAS where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      rs.Open "delete from TB_TEMP_PACKING_LIST_SALIDAS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   Else
      MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   txt_embarque = ""
   txt_embarque.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3850
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      cmd_imprimir.SetFocus
   End If
End Sub
