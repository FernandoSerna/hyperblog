VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmvistasprevias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   Icon            =   "frmvistasprevias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer cr 
      Height          =   5175
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   11655
      _cx             =   20558
      _cy             =   9128
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
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
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   2058
      EnableInteractiveParameterPrompting=   0   'False
   End
   Begin VB.Frame frm_lista 
      Height          =   2760
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   11565
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2235
         Left            =   45
         TabIndex        =   3
         Top             =   420
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   3942
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Correo"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   4
         Top             =   135
         Width           =   11475
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11310
      Picture         =   "frmvistasprevias.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   15
      TabIndex        =   1
      Top             =   270
      Width           =   11670
   End
End
Attribute VB_Name = "frmvistasprevias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_correo_Click()
   Dim var_n As Integer
   var_cadena_correo_2 = "fserna@vianney.com.mx"
   var_numero_tabla = var_consecutivo_estado_cuenta
   Me.lv_lista.ListItems.Clear
   If var_tipo_reporte_estado_cuenta = 1 Then
      var_si = MsgBox("Desea enviar por correo el estado de cuenta", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar el envio del estado de cuenta", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_posible = True
            rs.Open "select distinct isnull(vcha_cli_email,'') as vcha_cli_email, vcha_cli_nombre from tb_clientes where vcha_cli_clave_id = '" + var_clave_estado_cuenta + "'", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  var_cadena_correo = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                  If var_cadena_correo <> "" Then
                     rsaux10.Open "select * from VW_ESTADO_CUENTA where INTE_TEC_NUMERO_CONTROL = " + Str(var_numero_tabla) + " and VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "' and VCHA_MAQ_MAQUINA_ID = '" + fun_NombrePc + "' and VCHA_CLI_CLAVE_ID = '" + var_clave_estado_cuenta + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux10.EOF Then
                        Set reporte = appl.OpenReport(App.Path + "\rep_estado_cuenta.rpt")
                        reporte.RecordSelectionFormula = "{VW_ESTADO_CUENTA.INTE_TEC_NUMERO_CONTROL} = " + Str(var_numero_tabla) + " and {VW_ESTADO_CUENTA.VCHA_USU_USUARIO_ID} = '" + var_clave_usuario_global + "' and {VW_ESTADO_CUENTA.VCHA_MAQ_MAQUINA_ID} = '" + fun_NombrePc + "' and {VW_ESTADO_CUENTA.VCHA_CLI_CLAVE_ID} = '" + var_clave_estado_cuenta + "'"
                        For ntablas = 1 To reporte.Database.Tables.Count
                            reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                        Next ntablas
                        reporte.ExportOptions.FormatType = crEFTExcel80
                        reporte.ExportOptions.DestinationType = crEDTDiskFile
                        archivo = "c:\reportessid\estado_cuenta_" + var_clave_cliente + ".xls"
                        reporte.ExportOptions.DiskFileName = archivo
                        reporte.Export False
                                 
                        If MAPISession1.SessionID = 0 Then
                           MAPISession1.SignOn
                        End If
                        MAPIMessages1.SessionID = MAPISession1.SessionID
                        MAPIMessages1.Compose
                        MAPIMessages1.RecipDisplayName = var_cadena_correo
                        MAPIMessages1.RecipAddress = var_cadena_correo
                        MAPIMessages1.AddressResolveUI = True
                        MAPIMessages1.ResolveName
                        MAPIMessages1.MsgSubject = "Estado de cuenta del cliente " + IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
                        MAPIMessages1.MsgNoteText = "Se adjunta archivo con el estado de cuenta del cliente : " + var_clave_cliente + " " + IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
                        MAPIMessages1.AttachmentPathName = archivo
                        MAPIMessages1.send False
                        If MAPISession1.SessionID > 0 Then
                           MAPISession1.SignOff
                        End If
                     End If
                     rsaux10.Close
                  End If
                  rs.MoveNext
            Wend
            rs.Close
            var_clave_titular = ""
            var_clave_grupo = ""
            MsgBox "A terminado el proceso de envio de estado de cuentas"
         End If
      End If
   End If
                     
                     
   If var_tipo_reporte_estado_cuenta = 2 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from vw_CLIENTES  WHERE VCHA_TIT_TITULAR_ID = '" + var_clave_estado_cuenta + "' and vcha_cli_email is not null and vcha_cli_email <> '' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CLI_CLAVE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
            list_item.SubItems(2) = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      VAR_TIPO_LISTA = 100
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 7100
      Else
         lv_lista.ColumnHeaders(2).Width = 7300
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
                     
   If var_tipo_reporte_estado_cuenta = 3 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from vw_CLIENTES  WHERE vcha_gac_grupo_Actual_id = '" + var_clave_estado_cuenta + "' and vcha_cli_email is not null and vcha_cli_email <> '' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CLI_CLAVE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
            list_item.SubItems(2) = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      VAR_TIPO_LISTA = 100
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 7100
      Else
         lv_lista.ColumnHeaders(2).Width = 7300
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   
   If var_nombre_reporte = "packing_list" Then
      rsaux3.Open "SELECT * FROM TB_AGENTES WHERE VCHA_aGE_AGENTE_ID = '" + var_agente_packing_list + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         VAR_CORREO_ELECTRONICO = IIf(IsNull(rsaux3!VCHA_AGE_EMAIL), "", rsaux3!VCHA_AGE_EMAIL)
         If VAR_CORREO_ELECTRONICO <> "" Then
            Set reporte = appl.OpenReport(App.Path + "\rep_PACKING_LIST_TEMPORAL_eXPORTACIONES.rpt")
            reporte.RecordSelectionFormula = "{VW_PACKING_LIST_temporal.inte_EMB_EMBARQUE} = " + CStr(var_embarque_packing_list) + " and {VW_PACKING_LIST_temporal.VCHA_EMP_EMPRESA_ID} ='" + var_empresa + "' and {VW_PACKING_LIST_temporal.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "' and {VW_PACKING_LIST_temporal.floa_paq_cantidad} > 0 and {VW_PACKING_LIST_temporal.inte_tem_consecutivo} = " + CStr(var_consecutivo_packing_list)
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\packing_list_" + CStr(var_embarque_packing_list) + ".xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
      
            If MAPISession1.SessionID = 0 Then
               MAPISession1.SignOn
            End If
            MAPIMessages1.SessionID = MAPISession1.SessionID
            MAPIMessages1.Compose
            MAPIMessages1.RecipDisplayName = VAR_CORREO_ELECTRONICO
            MAPIMessages1.RecipAddress = VAR_CORREO_ELECTRONICO
            MAPIMessages1.AddressResolveUI = True
            MAPIMessages1.ResolveName
            MAPIMessages1.MsgSubject = "Packing list del embarque " + CStr(var_embarque_packing_list) + " del agente " + rsaux3!VCHA_AGE_NOMBRE
            MAPIMessages1.MsgNoteText = "Se adjunta Packing list del embarque " + CStr(var_embarque_packing_list) + " del agente " + rsaux3!VCHA_AGE_NOMBRE
            MAPIMessages1.AttachmentPathName = archivo
            MAPIMessages1.send True
            If MAPISession1.SessionID > 0 Then
               MAPISession1.SignOff
            End If
         Else
            MsgBox "El agente " + Trim(rsaux3!VCHA_AGE_NOMBRE) + " no tiene cuenta de correo electronico", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El agente " + Trim(rsaux3!VCHA_AGE_NOMBRE) + " no tiene cuenta de correo electronico", vbOKOnly, "ATENCION"
      End If
      rsaux3.Close
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Me.frm_lista.Visible = False
   If var_correo_estado_cuenta = 1 Then
      Me.cmd_salir.Visible = True
   Else
      Me.cmd_salir.Visible = False
   End If
   If Trim(var_nombre_reporte) <> "" Then
      If var_nombre_reporte = "packing_list" Then
         Me.cmd_salir.Visible = True
      End If
   Else
      Me.cmd_salir.Visible = False
   End If
End Sub

Private Sub Form_Resize()
   cr.Height = Me.Height - 50
   cr.Width = Me.Width - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_empresa = "31" Then
      On Error GoTo SALIR:
   End If
   var_correo_estado_cuenta = 0
   var_nombre_reporte = ""
   Call activa_forma(var_activa_forma_vistasprevias)
   Exit Sub
SALIR:
   Exit Sub
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   var_cadena_correo_2 = "fserna@vianney.com.mx"
   var_numero_tabla = var_consecutivo_estado_cuenta
   If KeyAscii = 13 Then
      If var_tipo_reporte_estado_cuenta = 2 Then
         var_si = MsgBox("Desea enviar por correo el estado de cuenta", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar el envio del estado de cuenta", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_posible = True
               var_cadena_correo = Me.lv_lista.selectedItem.SubItems(2)
               If var_cadena_correo <> "" Then
                  rsaux10.Open "select * from VW_ESTADO_CUENTA where INTE_TEC_NUMERO_CONTROL = " + Str(var_numero_tabla) + " and VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "' and VCHA_MAQ_MAQUINA_ID = '" + fun_NombrePc + "' and VCHA_TIT_TITULAR_ID = '" + var_clave_estado_cuenta + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux10.EOF Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_estado_cuenta.rpt")
                     reporte.RecordSelectionFormula = "{VW_ESTADO_CUENTA.INTE_TEC_NUMERO_CONTROL} = " + Str(var_numero_tabla) + " and {VW_ESTADO_CUENTA.VCHA_USU_USUARIO_ID} = '" + var_clave_usuario_global + "' and {VW_ESTADO_CUENTA.VCHA_MAQ_MAQUINA_ID} = '" + fun_NombrePc + "' and {VW_ESTADO_CUENTA.VCHA_TIT_TITULAR_ID} = '" + var_clave_estado_cuenta + "'"
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\estado_cuenta_" + rsaux10!vcha_tit_titular_id + ".xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                                   
                     If MAPISession1.SessionID = 0 Then
                        MAPISession1.SignOn
                     End If
                     MAPIMessages1.SessionID = MAPISession1.SessionID
                     MAPIMessages1.Compose
                     MAPIMessages1.RecipDisplayName = var_cadena_correo
                     MAPIMessages1.RecipAddress = var_cadena_correo
                     MAPIMessages1.AddressResolveUI = True
                     MAPIMessages1.ResolveName
                     MAPIMessages1.MsgSubject = "Estado de cuenta del titular " + IIf(IsNull(rsaux10!VCHA_tit_NOMBRE), "", rsaux10!VCHA_tit_NOMBRE)
                     MAPIMessages1.MsgNoteText = "Se adjunta archivo con el estado de cuenta del titular : " + IIf(IsNull(rsaux10!vcha_tit_titular_id), "", rsaux10!vcha_tit_titular_id) + " " + IIf(IsNull(rsaux10!VCHA_tit_NOMBRE), "", rsaux10!VCHA_tit_NOMBRE)
                     MAPIMessages1.AttachmentPathName = archivo
                     MAPIMessages1.send False
                     If MAPISession1.SessionID > 0 Then
                        MAPISession1.SignOff
                     End If
                  End If
                  rsaux10.Close
               End If
               var_clave_cliente = ""
               var_clave_grupo = ""
               MsgBox "A terminado el proceso de envio de estado de cuentas"
            End If
         End If
      End If
      
      If var_tipo_reporte_estado_cuenta = 3 Then
         var_si = MsgBox("Desea enviar por correo el estado de cuenta", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar el envio del estado de cuenta", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_posible = True
               var_clave_cliente = ""
               var_clave_titular = ""
               var_clave_grupo = txt_grupo_actual
               var_cadena_correo = Me.lv_lista.selectedItem
               If var_cadena_correo <> "" Then
                  If rsaux10.State = 1 Then
                     rsaux10.Close
                  End If
                  rsaux10.Open "select * from VW_ESTADO_CUENTA where INTE_TEC_NUMERO_CONTROL = " + Str(var_numero_tabla) + " and VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "' and VCHA_MAQ_MAQUINA_ID = '" + fun_NombrePc + "' and VCHA_GAC_GRUPO_ACTUAL_ID = '" + var_clave_estado_cuenta + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux10.EOF Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_estado_cuenta.rpt")
                     reporte.RecordSelectionFormula = "{VW_ESTADO_CUENTA.INTE_TEC_NUMERO_CONTROL} = " + Str(var_numero_tabla) + " and {VW_ESTADO_CUENTA.VCHA_USU_USUARIO_ID} = '" + var_clave_usuario_global + "' and {VW_ESTADO_CUENTA.VCHA_MAQ_MAQUINA_ID} = '" + fun_NombrePc + "'"
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\estado_cuenta_" + rsaux10!VCHA_GAC_GRUPO_aCTUAL_ID + ".xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                  
                     If MAPISession1.SessionID = 0 Then
                        MAPISession1.SignOn
                     End If
                     MAPIMessages1.SessionID = MAPISession1.SessionID
                     MAPIMessages1.Compose
                     MAPIMessages1.RecipDisplayName = var_cadena_correo
                     MAPIMessages1.RecipAddress = var_cadena_correo
                     MAPIMessages1.AddressResolveUI = True
                     MAPIMessages1.ResolveName
                     rsaux11.Open "select * from tb_gruposactuales where vcha_gac_grupo_Actual_id = '" + rsaux10!VCHA_GAC_GRUPO_aCTUAL_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                     MAPIMessages1.MsgSubject = "Estado de cuenta del grupo " + IIf(IsNull(rsaux11!vcha_gac_nombre), "", rsaux11!vcha_gac_nombre)
                     MAPIMessages1.MsgNoteText = "Se adjunta archivo con el estado de cuenta del grupo : " + IIf(IsNull(rsaux10!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux10!VCHA_GAC_GRUPO_aCTUAL_ID) + " " + IIf(IsNull(rsaux11!vcha_gac_nombre), "", rsaux11!vcha_gac_nombre)
                     rsaux11.Close
                     MAPIMessages1.AttachmentPathName = archivo
                     MAPIMessages1.send False
                     If MAPISession1.SessionID > 0 Then
                        MAPISession1.SignOff
                     End If
                  End If
                  rsaux10.Close
               End If
               MsgBox "A terminado el proceso de envio de estado de cuentas"
            End If
         End If
      End If
      Me.frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      Me.frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub
