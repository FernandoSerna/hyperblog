VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmestados_cuenta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado de Cuenta"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   2415
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   1485
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton cmd_correo 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   735
      Picture         =   "frmestados_cuenta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Enviar estado de cuenta por correo"
      Top             =   15
      Visible         =   0   'False
      Width           =   75
   End
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Left            =   4395
      TabIndex        =   25
      Top             =   4890
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   86179841
      CurrentDate     =   38303
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   120
      TabIndex        =   22
      Top             =   390
      Width           =   6090
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   23
         Top             =   495
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3228
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
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Gurpo"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   24
         Top             =   120
         Width           =   6015
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   " Tipo de Reporte "
      Height          =   1170
      Left            =   9090
      TabIndex        =   21
      Top             =   6075
      Width           =   2475
      Begin VB.OptionButton opt_tipo_documento_2 
         Caption         =   "A la fecha"
         Height          =   375
         Left            =   195
         TabIndex        =   14
         Top             =   660
         Width           =   2145
      End
      Begin VB.OptionButton opt_tipo_documento_1 
         Caption         =   "Por Movimientos"
         Height          =   330
         Left            =   195
         TabIndex        =   13
         Top             =   330
         Width           =   2205
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmestados_cuenta.frx":0282
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
      Picture         =   "frmestados_cuenta.frx":0384
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11205
      Picture         =   "frmestados_cuenta.frx":0486
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   45
      Width           =   330
   End
   Begin VB.Frame Frame5 
      Height          =   120
      Left            =   75
      TabIndex        =   20
      Top             =   270
      Width           =   11535
   End
   Begin VB.Frame Frame4 
      Caption         =   " Condiciones del Estado de Cuenta"
      Height          =   1170
      Left            =   5850
      TabIndex        =   17
      Top             =   6075
      Width           =   3195
      Begin VB.CommandButton cmd_mes_2 
         Height          =   330
         Left            =   2385
         Picture         =   "frmestados_cuenta.frx":0AC0
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   690
         Width           =   345
      End
      Begin VB.CommandButton cmd_mes_1 
         Height          =   330
         Left            =   2370
         Picture         =   "frmestados_cuenta.frx":1D32
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   300
         Width           =   345
      End
      Begin VB.TextBox txt_fecha_fin 
         Height          =   330
         Left            =   1005
         TabIndex        =   12
         Top             =   705
         Width           =   1320
      End
      Begin VB.TextBox txt_fecha_inicio 
         Height          =   315
         Left            =   1005
         TabIndex        =   11
         Top             =   300
         Width           =   1305
      End
      Begin VB.OptionButton opt_tipo_reporte_3 
         Caption         =   "Grupo"
         Height          =   360
         Left            =   195
         TabIndex        =   10
         Top             =   765
         Width           =   1785
      End
      Begin VB.OptionButton opt_tipo_reporte_2 
         Caption         =   "Titular"
         Height          =   300
         Left            =   195
         TabIndex        =   9
         Top             =   510
         Width           =   2010
      End
      Begin VB.OptionButton opt_tipo_reporte_1 
         Caption         =   "Cliente"
         Height          =   345
         Left            =   195
         TabIndex        =   8
         Top             =   195
         Width           =   1830
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Clientes  "
      Height          =   5520
      Left            =   5850
      TabIndex        =   16
      Top             =   465
      Width           =   5700
      Begin MSComctlLib.ListView lv_clientes 
         Height          =   5220
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   9208
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Titulares "
      Height          =   5970
      Left            =   120
      TabIndex        =   15
      Top             =   1290
      Width           =   5670
      Begin MSComctlLib.ListView lv_titulares 
         Height          =   5700
         Left            =   90
         TabIndex        =   6
         Top             =   195
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   10054
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Grupo Actual"
      Height          =   720
      Left            =   120
      TabIndex        =   0
      Top             =   465
      Width           =   5670
      Begin VB.TextBox txt_nombre_grupo_actual 
         Height          =   315
         Left            =   1500
         TabIndex        =   5
         Top             =   255
         Width           =   4110
      End
      Begin VB.TextBox txt_grupo_actual 
         Height          =   315
         Left            =   75
         TabIndex        =   4
         Top             =   255
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frmestados_cuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_servidor_Temporal As String
Dim var_base_Datos_Temporal As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_tipo_reporte As Integer
Dim var_tipo_mes As Integer
Dim var_tipo_lista As Integer
Private Sub cmb_grupos_actuales_Click()
   txt_grupo_actual = Obtener_llave(cnn_reportes, rs, "TB_GRUPOSACTUALES", "VCHA_GAC_NOMBRE", cmb_grupos_actuales, 0, "T")
   If Trim(txt_grupo_actual) <> "" Then
      rs.Open "select * from vw_clientes where vcha_gac_grupo_Actual_id = '" + txt_grupo_actual + "'", cnn_reportes, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         cmb_grupos_actuales = rs!vcha_gac_nombre
         Dim list_item As ListItem
         Dim var_titualar As String
         Dim n As Double
         lv_titulares.ListItems.Clear
         rsaux2.Open "select distinct vcha_tit_titular_id, vcha_tit_nombre from vw_clientes where vcha_gac_grupo_actual_id = '" + txt_grupo_actual + "' order by vcha_tit_nombre", cnn_reportes, adOpenDynamic, adLockOptimistic
         numero_items_estampados = 0
         While Not rsaux2.EOF
            Set list_item = lv_titulares.ListItems.Add(, , rsaux2(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
            rsaux2.MoveNext:
            numero_items_estampados = numero_items_estampados + 1
         Wend
         rsaux2.Close
         n = lv_titulares.ListItems.Count
         If n > 0 Then
            lv_titulares.ListItems.Item(1).Selected = True
            var_titular = lv_titulares.selectedItem
            lv_clientes.ListItems.Clear
            rsaux2.Open "select distinct vcha_cli_clave_id, vcha_cli_nombre from vw_clientes where vcha_tit_titular_id = '" + var_titular + "' order by vcha_cli_nombre", cnn_reportes, adOpenDynamic, adLockOptimistic
            numero_items_estampados = 0
            While Not rsaux2.EOF
               Set list_item = lv_clientes.ListItems.Add(, , rsaux2(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
               rsaux2.MoveNext:
               numero_items_estampados = numero_items_estampados + 1
            Wend
            rsaux2.Close
         End If
      Else
         txt_grupo_actual = ""
      End If
      rs.Close
   End If
End Sub

Private Sub cmd_correo_Click()
   Dim var_clave_cliente As String
   Dim var_clave_titualr As String
   Dim var_clave_grupo As String
   Dim var_posible As Boolean
   If IsDate(txt_fecha_inicio) Then
      If IsDate(txt_fecha_fin) Then
         var_cadena_correo = ""
         If opt_tipo_reporte_1 = True Then
            If lv_clientes.ListItems.Count > 0 Then
               var_posible = True
               var_clave_cliente = lv_clientes.selectedItem
               rs.Open "select isnull(vcha_cli_email,'') as vcha_cli_email from tb_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn_reportes, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     If var_cadena_correo = "" Then
                        var_cadena_correo = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                     Else
                        var_cadena_correo = ";" + IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                     End If
                     rs.MoveNext
               Wend
               rs.Close
               var_clave_titular = ""
               var_clave_grupo = ""
            Else
               var_posible = False
            End If
         End If
         
         If opt_tipo_reporte_2 = True Then
            If lv_titulares.ListItems.Count > 0 Then
               var_posible = True
               var_clave_titular = lv_titulares.selectedItem
               rs.Open "select isnull(vcha_cli_email,'') as vcha_cli_email from tb_clientes where vcha_tit_titular_id = '" + var_clave_titular + "'", cnn_reportes, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     If var_cadena_correo = "" Then
                        var_cadena_correo = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                     Else
                        var_cadena_correo = ";" + IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                     End If
                     rs.MoveNext
               Wend
               rs.Close
               var_clave_cliente = ""
               var_clave_grupo = ""
            Else
               var_posible = False
            End If
         End If
         
         If opt_tipo_reporte_3 = True Then
            If Trim(txt_grupo_actual) <> "" Then
               var_posible = True
               var_clave_cliente = ""
               var_clave_titular = ""
               var_clave_grupo = txt_grupo_actual
               rs.Open "select isnull(vcha_cli_email,'') as vcha_cli_email from vw_clientes where vcha_gac_grupo_actual_id = '" + var_clave_grupo + "'", cnn_reportes, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     If var_cadena_correo = "" Then
                        var_cadena_correo = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                     Else
                        var_cadena_correo = ";" + IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                     End If
                     rs.MoveNext
               Wend
               rs.Close
            Else
               var_posible = False
            End If
         End If
         
         
         If var_cadena_correo <> "" Then
            If opt_tipo_documento_1 = True Or opt_tipo_documento_2 = True Then
               If opt_tipo_reporte_1 = True Then
                  If lv_clientes.ListItems.Count > 0 Then
                     var_posible = True
                     var_clave_cliente = lv_clientes.selectedItem
                     var_clave_titular = ""
                     var_clave_grupo = ""
                  Else
                     var_posible = False
                  End If
               End If
               If opt_tipo_reporte_2 = True Then
                  If lv_titulares.ListItems.Count > 0 Then
                     var_posible = True
                     var_clave_titular = lv_titulares.selectedItem
                     var_clave_cliente = ""
                     var_clave_grupo = ""
                  Else
                     var_posible = False
                  End If
               End If
               If opt_tipo_reporte_3 = True Then
                  If Trim(txt_grupo_actual) <> "" Then
                     var_posible = True
                     var_clave_cliente = ""
                     var_clave_titular = ""
                     var_clave_grupo = txt_grupo_actual
                  Else
                     var_posible = False
                  End If
               End If
               If var_posible = True Then
                  cnn_reportes.CommandTimeout = 3000
                  Dim cmd As New Command
                  Dim var_numero_tabla As Double
                  Set cmd.ActiveConnection = cnn_reportes
                  cmd.CommandType = adCmdStoredProc
                  cmd.CommandText = "SALDO_INICIAL_ESTADO_CUENTA"
                  cmd("@maquina") = fun_NombrePc
                  cmd("@usuario") = var_clave_usuario_global
                  cmd("@EMPRESA") = var_empresa
                  cmd("@UNIDAD") = var_unidad_organizacional
                  cmd("@CLIENTE") = var_clave_cliente
                  cmd("@titular") = var_clave_titular
                  cmd("@grupo") = var_clave_grupo
                  cmd("@FECHA_INICIO") = txt_fecha_inicio
                  cmd("@fecha_fin") = txt_fecha_fin
                  cmd("@numero_tabla") = 0
                  cmd.execute
                  var_numero_tabla = cmd("@numero_tabla")
                  Set cmd = Nothing
                  
                  var_cadena_correo_2 = "fserna@vianney.com.mx"
                  If opt_tipo_reporte_1 = True Then
                     If lv_clientes.ListItems.Count > 0 Then
                        var_posible = True
                        var_clave_cliente = lv_clientes.selectedItem
                        rs.Open "select isnull(vcha_cli_email,'') as vcha_cli_email, vcha_cli_nombre from tb_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn_reportes, adOpenDynamic, adLockOptimistic
                        While Not rs.EOF
                              var_cadena_correo = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                              If var_cadena_correo <> "" Then
                                 rsaux10.Open "select * from VW_ESTADO_CUENTA where INTE_TEC_NUMERO_CONTROL = " + Str(var_numero_tabla) + " and VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "' and VCHA_MAQ_MAQUINA_ID = '" + fun_NombrePc + "' and VCHA_CLI_CLAVE_ID = '" + var_clave_cliente + "'", cnn_reportes, adOpenDynamic, adLockOptimistic
                                 If Not rsaux10.EOF Then
                                    Set reporte = appl.OpenReport(App.Path + "\rep_estado_cuenta.rpt")
                                    reporte.RecordSelectionFormula = "{VW_ESTADO_CUENTA.INTE_TEC_NUMERO_CONTROL} = " + Str(var_numero_tabla) + " and {VW_ESTADO_CUENTA.VCHA_USU_USUARIO_ID} = '" + var_clave_usuario_global + "' and {VW_ESTADO_CUENTA.VCHA_MAQ_MAQUINA_ID} = '" + fun_NombrePc + "' and {VW_ESTADO_CUENTA.VCHA_CLI_CLAVE_ID} = '" + var_clave_cliente + "'"
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
                                    MAPIMessages1.RecipDisplayName = var_cadena_correo_2
                                    MAPIMessages1.RecipAddress = var_cadena_correo_2
                                    MAPIMessages1.AddressResolveUI = True
                                    MAPIMessages1.ResolveName
                                    MAPIMessages1.MsgSubject = "Estado de cuenta del cliente " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                                    MAPIMessages1.MsgNoteText = "Se adjunta archivo con el estado de cuenta del cliente : " + var_clave_cliente + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                                    MAPIMessages1.AttachmentPathName = archivo
                                    MAPIMessages1.Send False
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
                     Else
                        var_posible = False
                     End If
                  End If
                  
                     
                     
                  If opt_tipo_reporte_2 = True Then
                     If lv_titulares.ListItems.Count > 0 Then
                        var_posible = True
                        var_clave_titular = lv_titulares.selectedItem
                        rs.Open "select isnull(vcha_cli_email,'') as vcha_cli_email, vcha_cli_clave_id, vcha_cli_nombre  from tb_clientes where vcha_tit_titular_id = '" + var_clave_titular + "'", cnn_reportes, adOpenDynamic, adLockOptimistic
                        While Not rs.EOF
                              var_cadena_correo = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                              var_clave_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                              If var_cadena_correo <> "" Then
                                 rsaux10.Open "select * from VW_ESTADO_CUENTA where INTE_TEC_NUMERO_CONTROL = " + Str(var_numero_tabla) + " and VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "' and VCHA_MAQ_MAQUINA_ID = '" + fun_NombrePc + "' and VCHA_CLI_CLAVE_ID = '" + var_clave_cliente + "'", cnn_reportes, adOpenDynamic, adLockOptimistic
                                 If Not rsaux10.EOF Then
                                    Set reporte = appl.OpenReport(App.Path + "\rep_estado_cuenta.rpt")
                                    reporte.RecordSelectionFormula = "{VW_ESTADO_CUENTA.INTE_TEC_NUMERO_CONTROL} = " + Str(var_numero_tabla) + " and {VW_ESTADO_CUENTA.VCHA_USU_USUARIO_ID} = '" + var_clave_usuario_global + "' and {VW_ESTADO_CUENTA.VCHA_MAQ_MAQUINA_ID} = '" + fun_NombrePc + "' and {VW_ESTADO_CUENTA.VCHA_CLI_CLAVE_ID} = '" + var_clave_cliente + "'"
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
                                    MAPIMessages1.RecipDisplayName = var_cadena_correo_2
                                    MAPIMessages1.RecipAddress = var_cadena_correo_2
                                    MAPIMessages1.AddressResolveUI = True
                                    MAPIMessages1.ResolveName
                                    MAPIMessages1.MsgSubject = "Estado de cuenta del cliente " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                                    MAPIMessages1.MsgNoteText = "Se adjunta archivo con el estado de cuenta del cliente : " + var_clave_cliente + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                                    MAPIMessages1.AttachmentPathName = archivo
                                    MAPIMessages1.Send False
                                    If MAPISession1.SessionID > 0 Then
                                       MAPISession1.SignOff
                                    End If
                                 End If
                                 rsaux10.Close
                              End If
                              rs.MoveNext
                        Wend
                        rs.Close
                        var_clave_cliente = ""
                        var_clave_grupo = ""
                     Else
                        var_posible = False
                     End If
                  End If
                     
                     
                     
                  If opt_tipo_reporte_3 = True Then
                     If Trim(txt_grupo_actual) <> "" Then
                        var_posible = True
                        var_clave_cliente = ""
                        var_clave_titular = ""
                        var_clave_grupo = txt_grupo_actual
                        rs.Open "select isnull(vcha_cli_email,'') as vcha_cli_email, vcha_cli_clave_id, vcha_cli_nombre  from vw_clientes where vcha_gac_grupo_actual_id = '" + var_clave_grupo + "'", cnn_reportes, adOpenDynamic, adLockOptimistic
                        While Not rs.EOF
                              var_cadena_correo = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                              var_clave_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                              If var_cadena_correo <> "" Then
                                 rsaux10.Open "select * from VW_ESTADO_CUENTA where INTE_TEC_NUMERO_CONTROL = " + Str(var_numero_tabla) + " and VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "' and VCHA_MAQ_MAQUINA_ID = '" + fun_NombrePc + "' and VCHA_CLI_CLAVE_ID = '" + var_clave_cliente + "'"
                                 If Not rsaux10.EOF Then
                                    Set reporte = appl.OpenReport(App.Path + "\rep_estado_cuenta.rpt")
                                    reporte.RecordSelectionFormula = "{VW_ESTADO_CUENTA.INTE_TEC_NUMERO_CONTROL} = " + Str(var_numero_tabla) + " and {VW_ESTADO_CUENTA.VCHA_USU_USUARIO_ID} = '" + var_clave_usuario_global + "' and {VW_ESTADO_CUENTA.VCHA_MAQ_MAQUINA_ID} = '" + fun_NombrePc + "' and {VW_ESTADO_CUENTA.VCHA_CLI_CLAVE_ID} = '" + var_clave_cliente + "'"
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
                                    MAPIMessages1.RecipDisplayName = var_cadena_correo_2
                                    MAPIMessages1.RecipAddress = var_cadena_correo_2
                                    MAPIMessages1.AddressResolveUI = True
                                    MAPIMessages1.ResolveName
                                    MAPIMessages1.MsgSubject = "Estado de cuenta del cliente " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                                    MAPIMessages1.MsgNoteText = "Se adjunta archivo con el estado de cuenta del cliente : " + var_clave_cliente + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                                    MAPIMessages1.AttachmentPathName = archivo
                                    MAPIMessages1.Send False
                                    If MAPISession1.SessionID > 0 Then
                                       MAPISession1.SignOff
                                    End If
                                 End If
                                 rsaux10.Close
                              
                              End If
                              rs.MoveNext
                        Wend
                        rs.Close
                     Else
                        var_posible = False
                     End If
                  End If
                     
                     
                     
                  If rs.State = 1 Then
                     rs.Close
                  End If
                  rs.Open "delete from tb_temp_estado_cuenta where INTE_TEC_NUMERO_CONTROL = " + Str(var_numero_tabla) + " and VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "' and VCHA_MAQ_MAQUINA_ID = '" + fun_NombrePc + "'", cnn_reportes, adOpenDynamic, adLockOptimistic
               Else
                  MsgBox "No es posible ejecutar el reporte", vbOKOnly, "ATENCION"
               End If
            End If
            x = 0
            If x = 1 Then
            If opt_tipo_documento_2 = True Then
               If opt_tipo_reporte_1 = True Then
                  If lv_clientes.ListItems.Count > 0 Then
                     var_posible = True
                     var_clave_cliente = lv_clientes.selectedItem
                     var_clave_titular = ""
                     var_clave_grupo = ""
                  Else
                     var_posible = False
                  End If
               End If
               If opt_tipo_reporte_2 = True Then
                  If lv_titulares.ListItems.Count > 0 Then
                     var_posible = True
                     var_clave_titular = lv_titulares.selectedItem
                     var_clave_cliente = ""
                     var_clave_grupo = ""
                  Else
                     var_posible = False
                  End If
               End If
               If opt_tipo_reporte_3 = True Then
                  If Trim(txt_grupo_actual) <> "" Then
                     var_posible = True
                     var_clave_cliente = ""
                     var_clave_titular = ""
                     var_clave_grupo = txt_grupo_actual
                  Else
                     var_posible = False
                  End If
               End If
               If var_posible = True Then
               
               
                  If opt_tipo_reporte_1 = True Then
                     If lv_clientes.ListItems.Count > 0 Then
                        var_posible = True
                        var_clave_cliente = lv_clientes.selectedItem
                        rs.Open "select isnull(vcha_cli_email,'') as vcha_cli_email from tb_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn_reportes, adOpenDynamic, adLockOptimistic
                        While Not rs.EOF
                              var_cadena_correo = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                              If var_cadena_correo <> "" Then
                                 Set reporte = appl.OpenReport(App.Path + "\rep_cartera.rpt")
                                 reporte.RecordSelectionFormula = "{VW_CARTERA.VCHA_CLI_CLAVE_ID} = '" + var_clave_cliente + "'"
                                 For ntablas = 1 To reporte.Database.Tables.Count
                                     reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                 Next ntablas
                                 reporte.ExportOptions.FormatType = crEFTExcel80
                                 reporte.ExportOptions.DestinationType = crEDTDiskFile
                                 archivo = "c:\reportessid\estado_cuenta_" + var_clave_cliente + ".xls"
                                 reporte.ExportOptions.DiskFileName = archivo
                                 reporte.Export False
                                 Set reporte = Nothing
                              End If
                              rs.MoveNext
                        Wend
                        rs.Close
                        var_clave_titular = ""
                        var_clave_grupo = ""
                     Else
                        var_posible = False
                     End If
                  End If
               
               
                  If opt_tipo_reporte_2 = True Then
                     If lv_titulares.ListItems.Count > 0 Then
                        var_posible = True
                        var_clave_titular = lv_titulares.selectedItem
                        rs.Open "select isnull(vcha_cli_email,'') as vcha_cli_email, vcha_cli_clave_id  from tb_clientes where vcha_tit_titular_id = '" + var_clave_titular + "'", cnn_reportes, adOpenDynamic, adLockOptimistic
                        While Not rs.EOF
                              var_cadena_correo = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                              var_clave_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                              If var_cadena_correo <> "" Then
                                 Set reporte = appl.OpenReport(App.Path + "\rep_cartera.rpt")
                                 reporte.RecordSelectionFormula = "{VW_CARTERA.VCHA_CLI_CLAVE_ID} = '" + var_clave_cliente + "'"
                                 frmvistasprevias.cr.ReportSource = reporte
                                 For ntablas = 1 To reporte.Database.Tables.Count
                                 reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                 Next ntablas
                                 reporte.ExportOptions.FormatType = crEFTExcel80
                                 reporte.ExportOptions.DestinationType = crEDTDiskFile
                                 archivo = "c:\reportessid\estado_cuenta_" + var_clave_cliente + ".xls"
                                 reporte.ExportOptions.DiskFileName = archivo
                                 reporte.Export False
                                 Set reporte = Nothing
                              End If
                              rs.MoveNext
                        Wend
                        rs.Close
                        var_clave_cliente = ""
                        var_clave_grupo = ""
                     Else
                        var_posible = False
                     End If
                  End If
                     
               
                  If opt_tipo_reporte_3 = True Then
                     If Trim(txt_grupo_actual) <> "" Then
                        var_posible = True
                        var_clave_cliente = ""
                        var_clave_titular = ""
                        var_clave_grupo = txt_grupo_actual
                        rs.Open "select isnull(vcha_cli_email,'') as vcha_cli_email, vcha_cli_clave_id from vw_clientes where vcha_gac_grupo_actual_id = '" + var_clave_grupo + "'", cnn_reportes, adOpenDynamic, adLockOptimistic
                        While Not rs.EOF
                              var_clave_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                              If var_cadena_correo <> "" Then
                                 Set reporte = appl.OpenReport(App.Path + "\rep_cartera.rpt")
                                 reporte.RecordSelectionFormula = "{VW_CARTERA.VCHA_CLI_CLAVE_ID} = '" + var_clave_cliente + "'"
                                 frmvistasprevias.cr.ReportSource = reporte
                                 For ntablas = 1 To reporte.Database.Tables.Count
                                 reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                 Next ntablas
                                 reporte.ExportOptions.FormatType = crEFTExcel80
                                 reporte.ExportOptions.DestinationType = crEDTDiskFile
                                 archivo = "c:\reportessid\estado_cuenta_" + var_clave_cliente + ".xls"
                                 reporte.ExportOptions.DiskFileName = archivo
                                 reporte.Export False
                                 Set reporte = Nothing
                              End If
                              rs.MoveNext
                        Wend
                        rs.Close
                     Else
                        var_posible = False
                     End If
                  End If
                     
               
               
               
               
               
               Else
                  MsgBox "No es posible ejecutar el reporte", vbOKOnly, "ATENCION"
               End If
            End If
            End If
         Else
            MsgBox "No existen clientes con correo electronico", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "La fecha final es incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "La fecha de inicio es incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
   Dim var_clave_cliente As String
   Dim var_clave_titualr As String
   Dim var_clave_grupo As String
   Dim var_posible As Boolean
   Dim var_posible_correo As Boolean
   
   If IsDate(txt_fecha_inicio) Then
      If IsDate(txt_fecha_fin) Then
         If opt_tipo_documento_1 = True Then
            ''' para enciar el correo
            var_correo_estado_cuenta = 1
            var_cadena_correo = ""
            
            '''' fin de envio del correo
            
            var_nombre_reporte = "estado_cuenta"
            
            If opt_tipo_reporte_1 = True Then
               If lv_clientes.ListItems.Count > 0 Then
                  var_tipo_reporte = 1
                  var_tipo_reporte_estado_cuenta = 1
                  var_posible = True
                  var_clave_cliente = lv_clientes.selectedItem
                  var_clave_estado_cuenta = lv_clientes.selectedItem
                  var_clave_titular = ""
                  var_clave_grupo = ""
               Else
                  var_posible = False
               End If
            End If
            If opt_tipo_reporte_2 = True Then
               If lv_titulares.ListItems.Count > 0 Then
                  var_tipo_reporte_estado_cuenta = 2
                  var_tipo_reporte = 2
                  var_posible = True
                  var_clave_titular = lv_titulares.selectedItem
                  var_clave_estado_cuenta = lv_titulares.selectedItem
                  var_clave_cliente = ""
                  var_clave_grupo = ""
               Else
                  var_posible = False
               End If
            End If
            If opt_tipo_reporte_3 = True Then
               If Trim(txt_grupo_actual) <> "" Then
                  var_tipo_reporte_estado_cuenta = 3
                  var_tipo_reporte = 3
                  var_posible = True
                  var_clave_cliente = ""
                  var_clave_titular = ""
                  var_clave_grupo = txt_grupo_actual
                  var_clave_estado_cuenta = txt_grupo_actual
               Else
                  var_posible = False
               End If
            End If
            If var_posible = True Then
               
               'cnn_reportes.CommandTimeout = 0
               'Dim CMD As New Command
               'Dim var_numero_tabla As Double
               'Set CMD.ActiveConnection = cnn_reportes
               'CMD.CommandType = adCmdStoredProc
               'CMD.CommandText = "SALDO_INICIAL_ESTADO_CUENTA"
               'CMD("@maquina") = fun_NombrePc
               'CMD("@usuario") = var_clave_usuario_global
               'CMD("@EMPRESA") = var_empresa
               'CMD("@UNIDAD") = var_unidad_organizacional
               'CMD("@CLIENTE") = var_clave_cliente
               'CMD("@titular") = var_clave_titular
               'CMD("@grupo") = var_clave_grupo
               'CMD("@FECHA_INICIO") = txt_fecha_inicio
               'CMD("@fecha_fin") = txt_fecha_fin
               'CMD("@numero_tabla") = 0
               'CMD.execute
               'var_numero_tabla = CMD("@numero_tabla")
               'Set CMD = Nothing
               
               cnn_reportes.BeginTrans
               rs.Open "select max(INTE_TEC_NUMERO_CONTROL) as numero_maximo from tb_temp_estado_cuenta where VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "' and VCHA_MAQ_MAQUINA_ID = '" + fun_NombrePc + "'", cnn_reportes, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
               Else
                  var_consecutivo = 0
               End If
               rs.Close
               var_consecutivo = var_consecutivo + 1
               rs.Open "insert into tb_temp_estado_cuenta (INTE_TEC_NUMERO_CONTROL, VCHA_USU_USUARIO_ID, VCHA_MAQ_MAQUINA_ID) values (" + CStr(var_consecutivo) + ",'" + var_usuario_global + "','" + fun_NombrePc + "')", cnn_reportes, adOpenDynamic, adLockOptimistic
               cnn_reportes.CommitTrans
               
             var_fecha_fin_1 = CDate(Me.txt_fecha_fin)
             
             var_dia = CStr(Day(CDate(Me.txt_fecha_inicio)))
             var_mes = CStr(Month(CDate(Me.txt_fecha_inicio)))
             var_año = CStr(Year(CDate(Me.txt_fecha_inicio)))
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
               
               
               
               cnn_reportes.CommandTimeout = 360
               'MsgBox "exec SALDO_INICIAL_ESTADO_CUENTA_2 '" + fun_NombrePc + "','" + var_clave_usuario_global + "','" + var_empresa + "','" + var_unidad_organizacional + "',  '" + var_clave_cliente + "','" + var_clave_titular + "','" + var_clave_grupo + "', " + var_fecha_inicio + ", " + var_fecha_fin + ", " + CStr(var_consecutivo)
               rs.Open "exec SALDO_INICIAL_ESTADO_CUENTA_2 '" + fun_NombrePc + "','" + var_clave_usuario_global + "','" + var_empresa + "','" + var_unidad_organizacional + "',  '" + var_clave_cliente + "','" + var_clave_titular + "','" + var_clave_grupo + "', " + var_fecha_inicio + ", " + var_fecha_fin + ", " + CStr(var_consecutivo), cnn_reportes, adOpenDynamic, adLockOptimistic
               
               
Dim dl As Long                                 ' Valor devuelto por la función API
Dim sAttributes As String                  ' Aributos
Dim sDriver As String                       ' Nombre del controlador
Dim sDescription As String                ' Descripción del DSN
Dim sDsnName As String                  ' Nombre del DSN


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
   sAttributes = sAttributes & "Server=" + var_sr_reportes & Chr$(0)
   sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
   sAttributes = sAttributes & "Database=" + var_bd_reportes & Chr(0)
   strAttributes = strAttributes & "UID=sa" & Chr$(0)
   strAttributes = strAttributes & "PWD=elia" & Chr$(0)
   dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
   
   
               
               
               Set reporte = appl.OpenReport(App.Path + "\rep_estado_cuenta.rpt")
               reporte.RecordSelectionFormula = "{VW_ESTADO_CUENTA.INTE_TEC_NUMERO_CONTROL} = " + Str(var_consecutivo) + " and {VW_ESTADO_CUENTA.VCHA_USU_USUARIO_ID} = '" + var_clave_usuario_global + "' and {VW_ESTADO_CUENTA.VCHA_MAQ_MAQUINA_ID} = '" + fun_NombrePc + "'"
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte de Movimientos"
               var_consecutivo_estado_cuenta = var_consecutivo
               frmvistasprevias.Show 1
               Set reporte = Nothing
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_estado_cuenta.rpt")
                  reporte.RecordSelectionFormula = "{VW_ESTADO_CUENTA.INTE_TEC_NUMERO_CONTROL} = " + Str(var_consecutivo) + " and {VW_ESTADO_CUENTA.VCHA_USU_USUARIO_ID} = '" + var_clave_usuario_global + "' and {VW_ESTADO_CUENTA.VCHA_MAQ_MAQUINA_ID} = '" + fun_NombrePc + "'"
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\reporte_historico_estado_cuenta_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
               End If
               If rs.State = 1 Then
                  rs.Close
               End If
               rs.Open "delete from tb_temp_estado_cuenta where INTE_TEC_NUMERO_CONTROL = " + Str(var_consecutivo) + " and VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "' and VCHA_MAQ_MAQUINA_ID = '" + fun_NombrePc + "'", cnn_reportes, adOpenDynamic, adLockOptimistic
            Else
               MsgBox "No es posible ejecutar el reporte", vbOKOnly, "ATENCION"
            End If
         End If
         If opt_tipo_documento_2 = True Then
            If opt_tipo_reporte_1 = True Then
               If lv_clientes.ListItems.Count > 0 Then
                  var_tipo_reporte_estado_cuenta = 1
                  var_posible = True
                  var_clave_cliente = lv_clientes.selectedItem
                  var_clave_titular = ""
                  var_clave_grupo = ""
               Else
                  var_posible = False
               End If
            End If
            If opt_tipo_reporte_2 = True Then
               If lv_titulares.ListItems.Count > 0 Then
                  var_tipo_reporte_estado_cuenta = 2
                  var_posible = True
                  var_clave_titular = lv_titulares.selectedItem
                  var_clave_cliente = ""
                  var_clave_grupo = ""
               Else
                  var_posible = False
               End If
            End If
            If opt_tipo_reporte_3 = True Then
               If Trim(txt_grupo_actual) <> "" Then
                  var_tipo_reporte_estado_cuenta = 3
                  var_posible = True
                  var_clave_cliente = ""
                  var_clave_titular = ""
                  var_clave_grupo = txt_grupo_actual
               Else
                  var_posible = False
               End If
            End If
            If var_posible = True Then
               Set reporte = appl.OpenReport(App.Path + "\rep_cartera.rpt")
               If Trim(var_clave_cliente) <> "" Then
                  reporte.RecordSelectionFormula = "{VW_CARTERA.VCHA_CLI_CLAVE_ID} = '" + var_clave_cliente + "'"
               End If
               If Trim(var_clave_titular) <> "" Then
                  reporte.RecordSelectionFormula = "{VW_CARTERA.VCHA_TIT_TITULAR_ID} = '" + var_clave_titular + "'"
               End If
               If Trim(var_clave_grupo) <> "" Then
                  reporte.RecordSelectionFormula = "{VW_CARTERA.VCHA_GAC_GRUPO_ACTUAL} = '" + var_clave_grupo + "'"
               End If
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte de Movimientos"
               var_consecutivo_estado_cuenta = var_numero_tabla
               frmvistasprevias.Show 1
               Set reporte = Nothing
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_cartera.rpt")
                  If Trim(var_clave_cliente) <> "" Then
                     reporte.RecordSelectionFormula = "{VW_CARTERA.VCHA_CLI_CLAVE_ID} = '" + var_clave_cliente + "'"
                  End If
                  If Trim(var_clave_titular) <> "" Then
                     reporte.RecordSelectionFormula = "{VW_CARTERA.VCHA_TIT_TITULAR_ID} = '" + var_clave_titular + "'"
                  End If
                  If Trim(var_clave_grupo) <> "" Then
                     reporte.RecordSelectionFormula = "{VW_CARTERA.VCHA_GAC_GRUPO_ACTUAL} = '" + var_clave_grupo + "'"
                  End If
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  For ntablas = 1 To reporte.Database.Tables.Count
                     reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\reporte_historico_estado_cuenta_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
               End If
            Else
               MsgBox "No es posible ejecutar el reporte", vbOKOnly, "ATENCION"
            End If
         End If
      Else
         MsgBox "La fecha final es incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "La fecha de inicio es incorrecta", vbOKOnly, "ATENCION"
   End If
   var_correo_estado_cuenta = 0
End Sub

Private Sub cmd_mes_1_Click()
   If IsDate(Me.txt_fecha_inicio) Then
      mes.Value = CDate(Me.txt_fecha_inicio)
      mes.Visible = True
      mes.SetFocus
      var_tipo_mes = 1
   Else
      mes.Value = Date
      mes.Visible = True
      mes.SetFocus
      var_tipo_mes = 1
   End If
End Sub

Private Sub cmd_mes_2_Click()
   If IsDate(Me.txt_fecha_fin) Then
      mes.Value = CDate(Me.txt_fecha_fin)
   Else
      mes.Value = Date
   End If
   var_tipo_mes = 2
   mes.Visible = True
   mes.SetFocus
End Sub

Private Sub cmd_nuevo_Click()
   var_tipo_reporte = 1
   opt_tipo_reporte_1.Value = True
   opt_tipo_documento_1.Value = True
   Me.txt_grupo_actual.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
var_servidor_Temporal = var_sr_reportes
var_base_Datos_Temporal = var_bd_reportes
'var_sr_reportes = "SQLQUEZADA"
'var_bd_reportes = "SIDQUEZADA"
Dim dl As Long                                 ' Valor devuelto por la función API
Dim sAttributes As String                  ' Aributos
Dim sDriver As String                       ' Nombre del controlador
Dim sDescription As String                ' Descripción del DSN
Dim sDsnName As String                  ' Nombre del DSN

   'cnn_reportes.Close
   'cnn_reportes.Open var_conexion_string_distribucion

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
   sAttributes = sAttributes & "Server=" + var_sr_reportes & Chr$(0)
   sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
   sAttributes = sAttributes & "Database=" + var_bd_reportes & Chr(0)
   strAttributes = strAttributes & "UID=sa" & Chr$(0)
   strAttributes = strAttributes & "PWD=elia" & Chr$(0)
   dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
   
   
   
   mes.Visible = False
   var_cadena_seguridad = ""
   Top = 0
   Left = 0
   txt_fecha_inicio = Date
   txt_fecha_fin = Date
   var_tipo_reporte = 1
   opt_tipo_reporte_1.Value = True
   opt_tipo_documento_1.Value = True
   frm_lista.Visible = False
   If cnn_reportes.State = 1 Then
      cnn_reportes.Close
   End If
   cnn_reportes.Open var_conexion_reportes
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_estados_cuenta)
   var_sr_reportes = var_servidor_Temporal
   var_bd_reportes = var_base_Datos_Temporal
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_grupo_actual = lv_lista.selectedItem.SubItems(2)
         txt_nombre_grupo_actual = lv_lista.selectedItem.SubItems(1)
      End If
      txt_grupo_actual.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
      Me.txt_grupo_actual.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_titulares_ItemClick(ByVal Item As MSComctlLib.ListItem)
   n = lv_titulares.ListItems.Count
   If n > 0 Then
      var_titular = lv_titulares.selectedItem
      lv_clientes.ListItems.Clear
      rsaux2.Open "select distinct vcha_cli_clave_id, vcha_cli_nombre from vw_clientes where vcha_tit_titular_id = '" + var_titular + "' order by vcha_cli_nombre", cnn_reportes, adOpenDynamic, adLockOptimistic
      numero_items_estampados = 0
      While Not rsaux2.EOF
            Set list_item = lv_clientes.ListItems.Add(, , rsaux2(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
            rsaux2.MoveNext:
            numero_items_estampados = numero_items_estampados + 1
      Wend
      rsaux2.Close
   End If
End Sub

Private Sub mes_DateDblClick(ByVal DateDblClicked As Date)
   If var_tipo_mes = 1 Then
      Me.txt_fecha_inicio = mes.Value
      txt_fecha_inicio.SetFocus
   End If
   If var_tipo_mes = 2 Then
      Me.txt_fecha_fin = mes.Value
      txt_fecha_fin.SetFocus
   End If
End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_mes = 1 Then
         Me.txt_fecha_inicio = mes.Value
         txt_fecha_inicio.SetFocus
      End If
      If var_tipo_mes = 2 Then
         Me.txt_fecha_fin = mes.Value
         txt_fecha_fin.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      If var_tipo_mes = 1 Then
         txt_fecha_inicio.SetFocus
      End If
      If var_tipo_mes = 2 Then
         txt_fecha_fin.SetFocus
      End If
   End If
End Sub

Private Sub mes_LostFocus()
   mes.Visible = False
End Sub

Private Sub txt_fecha_fin_LostFocus()
   If Not IsDate(txt_fecha_fin) Then
      MsgBox "Fecha Incorrecta", vbOKOnly, "ATENCION"
      txt_fecha_fin = Date
   End If
End Sub

Private Sub txt_fecha_inicio_LostFocus()
   If Not IsDate(txt_fecha_inicio) Then
      MsgBox "Fecha Incorrecta", vbOKOnly, "ATENCION"
      txt_fecha_inicio = Date
   End If
End Sub

Private Sub txt_grupo_actual_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione busqueda de cliente presione F5 para grupo, F6 Titular, F7 cliente"
End Sub

Private Sub txt_grupo_actual_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GAC_NOMBRE from  VW_CLIENTES WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_gac_nombre", cnn_reportes, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID))
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
            list_item.SubItems(2) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "GRUPOS ACTUALES"
      var_tipo_lista = 1
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4499.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4699.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_tit_titular_id, vcha_tit_nombre, vcha_gac_grupo_Actual_id from vw_clientes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_tit_nombre", cnn_reportes, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tit_titular_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
            list_item.SubItems(2) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TITULARES"
      var_tipo_lista = 2
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4499.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4699.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 118 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_cli_clave_id, vcha_cli_nombre, vcha_gac_grupo_Actual_id from vw_clientes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_cli_nombre", cnn_reportes, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            list_item.SubItems(2) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 3
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4499.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4699.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_grupo_actual_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_grupo_actual_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_grupo_actual) <> "" Then
      rs.Open "select * from vw_clientes where vcha_gac_grupo_Actual_id = '" + txt_grupo_actual + "'", cnn_reportes, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_grupo_actual = rs!vcha_gac_nombre
         Dim list_item As ListItem
         Dim var_titualar As String
         Dim n As Double
         lv_titulares.ListItems.Clear
         rsaux2.Open "select distinct vcha_tit_titular_id, vcha_tit_nombre from vw_clientes where vcha_gac_grupo_actual_id = '" + txt_grupo_actual + "' order by vcha_tit_nombre", cnn_reportes, adOpenDynamic, adLockOptimistic
         numero_items_estampados = 0
         While Not rsaux2.EOF
               Set list_item = lv_titulares.ListItems.Add(, , rsaux2(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
               rsaux2.MoveNext:
               numero_items_estampados = numero_items_estampados + 1
         Wend
         rsaux2.Close
         n = lv_titulares.ListItems.Count
         If n > 0 Then
            lv_titulares.ListItems.Item(1).Selected = True
            var_titular = lv_titulares.selectedItem
            lv_clientes.ListItems.Clear
            rsaux2.Open "select distinct vcha_cli_clave_id, vcha_cli_nombre from vw_clientes where vcha_tit_titular_id = '" + var_titular + "' order by vcha_cli_nombre", cnn_reportes, adOpenDynamic, adLockOptimistic
            numero_items_estampados = 0
            While Not rsaux2.EOF
                  Set list_item = lv_clientes.ListItems.Add(, , rsaux2(0).Value)
                  list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
                  rsaux2.MoveNext:
                  numero_items_estampados = numero_items_estampados + 1
            Wend
            rsaux2.Close
         End If
      Else
         txt_grupo_actual = ""
         txt_nombre_grupo_actual = ""
         lv_titulares.ListItems.Clear
         lv_clientes.ListItems.Clear
      End If
      rs.Close
   Else
      txt_nombre_grupo_actual = ""
      lv_titulares.ListItems.Clear
      lv_clientes.ListItems.Clear
   End If
End Sub

Private Sub txt_nombre_grupo_actual_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione busqueda de cliente presione F5 para grupo, F6 Titular, F7 cliente"
End Sub

Private Sub txt_nombre_grupo_actual_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select VCHA_GAC_GRUPO_ACTUAL_ID from VW_CLIENTES WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_gac_nombre", cnn_reportes, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID))
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
            list_item.SubItems(2) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "GRUPOS ACTUALES"
      var_tipo_lista = 1
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4499.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4699.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_tit_titular_id, vcha_tit_nombre, vcha_gac_grupo_Actual_id from vw_clientes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_tit_nombre", cnn_reportes, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tit_titular_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
            list_item.SubItems(2) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TITULARES"
      var_tipo_lista = 2
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4499.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4699.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 118 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_cli_clave_id, vcha_cli_nombre, vcha_gac_grupo_Actual_id from vw_clientes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_cli_nombre", cnn_reportes, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            list_item.SubItems(2) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 3
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4499.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4699.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_grupo_actual_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_grupo_actual_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub
