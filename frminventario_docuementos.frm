VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frminventario_documentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario de Documentos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11625
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   4605
      TabIndex        =   28
      Top             =   3705
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   29
         Top             =   480
         Width           =   5595
         _ExtentX        =   9869
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   30
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.Frame frm_cambio_agente 
      Height          =   2700
      Left            =   4515
      TabIndex        =   17
      Top             =   3690
      Width           =   6030
      Begin VB.TextBox txt_nombre_agente_nuevo 
         Height          =   330
         Left            =   1470
         TabIndex        =   26
         Top             =   2100
         Width           =   4485
      End
      Begin VB.TextBox txt_agente_nuevo 
         Height          =   330
         Left            =   165
         TabIndex        =   25
         Top             =   2100
         Width           =   1290
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         Picture         =   "frminventario_docuementos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Cancelar"
         Top             =   465
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   30
         Picture         =   "frminventario_docuementos.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cambiar el agente"
         Top             =   465
         Width           =   330
      End
      Begin VB.Frame Frame5 
         Height          =   30
         Left            =   0
         TabIndex        =   21
         Top             =   795
         Width           =   6015
      End
      Begin VB.TextBox txt_nombre_agente_actual 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1455
         TabIndex        =   20
         Top             =   1215
         Width           =   4485
      End
      Begin VB.TextBox txt_agente_actual 
         Enabled         =   0   'False
         Height          =   330
         Left            =   150
         TabIndex        =   19
         Top             =   1215
         Width           =   1290
      End
      Begin VB.Label Label3 
         Caption         =   "Agente nuevo"
         Height          =   255
         Left            =   165
         TabIndex        =   27
         Top             =   1830
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Agente actual"
         Height          =   255
         Left            =   150
         TabIndex        =   24
         Top             =   945
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   " Cambiar facturas de agente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Left            =   30
         TabIndex        =   18
         Top             =   120
         Width           =   5955
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11205
      Picture         =   "frminventario_docuementos.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frminventario_docuementos.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Actualizar Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   5850
      TabIndex        =   2
      Top             =   390
      Width           =   5670
      Begin VB.CommandButton cmd_imprimir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1035
         Picture         =   "frminventario_docuementos.frx":09D0
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Reporte"
         Top             =   165
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frminventario_docuementos.frx":0AD2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Cambiar facturas de agente"
         Top             =   165
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   405
         Picture         =   "frminventario_docuementos.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   570
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   75
         Picture         =   "frminventario_docuementos.frx":0DEA
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   570
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1065
         Picture         =   "frminventario_docuementos.frx":0EEC
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   570
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   735
         Picture         =   "frminventario_docuementos.frx":0FBE
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Marcar (Enter)"
         Top             =   570
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1395
         Picture         =   "frminventario_docuementos.frx":1208
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   570
         Width           =   330
      End
      Begin VB.CommandButton cmd_filtro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   705
         Picture         =   "frminventario_docuementos.frx":141E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Filtrar"
         Top             =   165
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         Picture         =   "frminventario_docuementos.frx":1518
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   165
         Width           =   330
      End
      Begin VB.Frame Frame4 
         Height          =   120
         Left            =   15
         TabIndex        =   10
         Top             =   420
         Width           =   5655
      End
      Begin MSComctlLib.ListView lv_facturas 
         Height          =   4935
         Left            =   45
         TabIndex        =   3
         Top             =   1800
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8705
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ser"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Doc"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Número"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nombre"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Estatus"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_agente 
         Alignment       =   2  'Center
         Caption         =   "Facturas Activas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   31
         Top             =   915
         Width           =   5505
      End
      Begin VB.Label lbl_facturas 
         Alignment       =   2  'Center
         Caption         =   "Facturas Activas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   9
         Top             =   1305
         Width           =   5505
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6855
      Left            =   90
      TabIndex        =   0
      Top             =   390
      Width           =   5670
      Begin MSComctlLib.ListView lv_agentes 
         Height          =   6600
         Left            =   45
         TabIndex        =   1
         Top             =   195
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   90
      Left            =   60
      TabIndex        =   6
      Top             =   285
      Width           =   11520
   End
End
Attribute VB_Name = "frminventario_documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_agente As String
Dim var_filtro As Integer

Private Sub cmd_aceptar_Click()
   Dim var_si As Integer
   Dim var_documento As String
   Dim var_serie As String
   Dim var_i As Integer
   Dim var_n As Integer
   Dim var_j As Integer
   Dim var_estatus As String
   Dim var_numero As Integer
   Dim var_primera_vez As Integer
   Dim var_mensage As String
   Me.frm_cambio_agente.Visible = False
   If lbl_facturas = "Facturas Activas" Then
      var_mensage = "¿Deseas eliminar los documentos del inventario del agente?"
      var_mensage2 = "Confirmar la eliminación de los documentos del inventario del agente"
   Else
      var_mensage = "¿Desea volver a agregar los documentos al inventario del agente?"
      var_mensage2 = "Confirmar el agregado de los documentos al inventario del agente"
   End If
   
   var_si = MsgBox(var_mensage, vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox(var_mensage2, vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_n = lv_facturas.ListItems.Count
         var_primera_vez = 1
         For var_i = 1 To var_n
             lv_facturas.ListItems.Item(var_i).Selected = True
             If lv_facturas.selectedItem.SubItems(4) = "*" Then
                If var_primera_vez = 1 Then
                   rs.Open "select max(INTE_TID_CONCECUTIVO) from TB_TEMP_INVENTARIO_DOCUMENTOS", cnn, adOpenDynamic, adLockOptimistic
                   If Not rs.EOF Then
                      var_numero = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
                   Else
                      var_numero = 0
                   End If
                   var_numero = var_numero + 1
                   var_primera_vez = 0
                   rs.Close
                End If
                var_serie = lv_facturas.selectedItem
                var_documento = lv_facturas.selectedItem.SubItems(1)
                var_estatus = lv_facturas.selectedItem.SubItems(5)
                
                If var_estatus = "A" Then
                   rs.Open "update tb_inventario_documentos set char_ido_estatus = 'O' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + Trim(var_documento) + "' and vcha_ser_serie_id = '" + var_serie + "' and inte_car_numero = " + lv_facturas.selectedItem.SubItems(2), cnn, adOpenDynamic, adLockOptimistic
                End If
                If var_estatus = "O" Then
                   rs.Open "update tb_inventario_documentos set char_ido_estatus = 'A' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + Trim(var_documento) + "' and vcha_ser_serie_id = '" + var_serie + "' and inte_car_numero = " + lv_facturas.selectedItem.SubItems(2), cnn, adOpenDynamic, adLockOptimistic
                End If
                rs.Open "insert into TB_TEMP_INVENTARIO_DOCUMENTOS (VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, INTE_CAR_NUMERO, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, INTE_TID_CONCECUTIVO, VCHA_SER_SERIE_ID,CHAR_TID_ESTATUS) values ('" + var_empresa + "', '" + Trim(var_documento) + "'," + lv_facturas.selectedItem.SubItems(2) + ", '" + var_clave_usuario_global + "','" + fun_NombrePc + "', " + CStr(var_numero) + ", '" + var_serie + "', '" + var_estatus + "')", cnn, adOpenDynamic, adLockOptimistic
             End If
         Next var_i
         
         'var_n = lv_facturas.ListItems.Count
         'For var_i = 1 To var_n
         '    var_j = lv_facturas.ListItems.Count
         '    If var_i <= var_j Then
         '       lv_facturas.ListItems.Item(var_i).Selected = True
         '       If lv_facturas.selectedItem.SubItems(4) = "*" Then
         '          lv_facturas.ListItems.Remove (lv_facturas.selectedItem.Index)
         '       End If
         '    Else
         '       lv_facturas.ListItems.Item(var_j).Selected = True
         '       If lv_facturas.selectedItem.SubItems(4) = "*" Then
         '          lv_facturas.ListItems.Remove (lv_facturas.selectedItem.Index)
         '       End If
         '    End If
         'Next var_i
         
   lv_facturas.ListItems.Clear
   var_agente = lv_agentes.selectedItem
   var_filtro = 0
   lbl_facturas = "Facturas Activas"
   If var_agente <> "" Then
      rs.Open "select vcha_ser_Serie_id,vcha_car_documento,inte_car_numero, vcha_cli_nombre, char_ido_estatus from vw_inventario_documentos where vcha_age_agente_id = '" + var_agente + "' and CHAR_IDO_ESTATUS  = 'A' order by inte_car_numero", cnn, adOpenDynamic, adLockOptimistic
      numero_items_lineas = 0
      While Not rs.EOF
            Set list_item = lv_facturas.ListItems.Add(, , rs!VCHA_SER_SERIE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
            list_item.SubItems(2) = IIf(IsNull(rs!inte_Car_numero), 0, rs!inte_Car_numero)
            list_item.SubItems(3) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            list_item.SubItems(4) = ""
            list_item.SubItems(5) = IIf(IsNull(rs!char_ido_estatus), "", rs!char_ido_estatus)
            rs.MoveNext:
            numero_items_lineas = numero_items_lineas + 1
       Wend
       rs.Close
       If numero_items_lineas > 28 Then
          lv_facturas.ColumnHeaders(4).Width = 3300
       Else
          lv_facturas.ColumnHeaders(4).Width = 3100
       End If
   End If
         
         
         If var_estatus = "O" Then
            Set reporte = appl.OpenReport(App.Path + "\rep_inventario_documentos_traspasos_agente.rpt")
            reporte.RecordSelectionFormula = "{VW_INVENTARIO_DOCUMENTOS_TRASPASOS.VCHA_AUD_USUARIO} = '" + var_clave_usuario_global + "' and {VW_INVENTARIO_DOCUMENTOS_TRASPASOS.VCHA_AUD_MAQUINA} = '" + fun_NombrePc + "' and {VW_INVENTARIO_DOCUMENTOS_TRASPASOS.INTE_TID_CONCECUTIVO} = " + CStr(var_numero)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Progreso de Surtido"
            frmvistasprevias.Show 1
            Set reporte = Nothing
         End If
         If var_estatus = "A" Then
            Set reporte = appl.OpenReport(App.Path + "\rep_inventario_documentos_traspasos_cobranza.rpt")
            reporte.RecordSelectionFormula = "{VW_INVENTARIO_DOCUMENTOS_TRASPASOS.VCHA_AUD_USUARIO} = '" + var_clave_usuario_global + "' and {VW_INVENTARIO_DOCUMENTOS_TRASPASOS.VCHA_AUD_MAQUINA} = '" + fun_NombrePc + "' and {VW_INVENTARIO_DOCUMENTOS_TRASPASOS.INTE_TID_CONCECUTIVO} = " + CStr(var_numero)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Progreso de Surtido"
            frmvistasprevias.Show 1
            Set reporte = Nothing
         End If
         rs.Open "delete from TB_TEMP_INVENTARIO_DOCUMENTOS  where vcha_aud_maquina = '" + fun_NombrePc + "' and vcha_aud_usuario = '" + var_clave_usuario_global + "' and INTE_TID_CONCECUTIVO = " + CStr(var_numero), cnn, adOpenDynamic, adLockOptimistic
      End If
   End If
End Sub

Private Sub cmd_cancelar_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_facturas.ListItems.Count
   For i = 1 To n
       lv_facturas.ListItems.Item(i).SubItems(4) = " "
       lv_facturas.ListItems.Item(i).Bold = False
       lv_facturas.ListItems.Item(i).ForeColor = &H80000012
       lv_facturas.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_facturas.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_facturas.ListItems.Item(i).ListSubItems(3).Bold = False
       lv_facturas.ListItems.Item(i).ListSubItems(4).Bold = False
       lv_facturas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_facturas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
       lv_facturas.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
       lv_facturas.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
   Next
   lv_facturas.Refresh
End Sub

Private Sub cmd_filtro_Click()
   Me.frm_cambio_agente.Visible = False
   If var_filtro = 0 Then
      var_filtro = 1
      lbl_facturas = "Facturas en crédito y cobranza"
   Else
      var_filtro = 0
      lbl_facturas = "Facturas Activas"
   End If
   If var_filtro = 1 Then
      lv_facturas.ListItems.Clear
      var_agente = lv_agentes.selectedItem
      If var_agente <> "" Then
         rs.Open "select vcha_ser_Serie_id,vcha_car_documento,inte_car_numero, vcha_cli_nombre, char_ido_estatus from vw_inventario_documentos where vcha_age_agente_id = '" + var_agente + "' and CHAR_IDO_ESTATUS  = 'O' and floa_sal_importe > 0.01 order by inte_car_numero", cnn, adOpenDynamic, adLockOptimistic
         numero_items_lineas = 0
         While Not rs.EOF
               Set list_item = lv_facturas.ListItems.Add(, , rs!VCHA_SER_SERIE_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
               list_item.SubItems(2) = IIf(IsNull(rs!inte_Car_numero), 0, rs!inte_Car_numero)
               list_item.SubItems(3) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
               list_item.SubItems(4) = ""
               list_item.SubItems(5) = IIf(IsNull(rs!char_ido_estatus), "", rs!char_ido_estatus)
               rs.MoveNext:
               numero_items_lineas = numero_items_lineas + 1
          Wend
          rs.Close
          If numero_items_lineas > 28 Then
             lv_facturas.ColumnHeaders(4).Width = 3300
          Else
             lv_facturas.ColumnHeaders(4).Width = 3100
          End If
      End If
   End If
   If var_filtro = 0 Then
      lv_facturas.ListItems.Clear
      var_agente = lv_agentes.selectedItem
      If var_agente <> "" Then
         rs.Open "select vcha_ser_Serie_id,vcha_car_documento,inte_car_numero, vcha_cli_nombre, char_ido_estatus from vw_inventario_documentos where vcha_age_agente_id = '" + var_agente + "' and CHAR_IDO_ESTATUS  = 'A' and floa_sal_importe > 0.01 order by inte_car_numero", cnn, adOpenDynamic, adLockOptimistic
         numero_items_lineas = 0
         While Not rs.EOF
               Set list_item = lv_facturas.ListItems.Add(, , rs!VCHA_SER_SERIE_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
               list_item.SubItems(2) = IIf(IsNull(rs!inte_Car_numero), 0, rs!inte_Car_numero)
               list_item.SubItems(3) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
               list_item.SubItems(4) = ""
               list_item.SubItems(5) = IIf(IsNull(rs!char_ido_estatus), "", rs!char_ido_estatus)
               rs.MoveNext:
               numero_items_lineas = numero_items_lineas + 1
          Wend
          rs.Close
          If numero_items_lineas > 28 Then
             lv_facturas.ColumnHeaders(4).Width = 3300
          Else
             lv_facturas.ColumnHeaders(4).Width = 3100
          End If
      End If
   End If
End Sub

Private Sub cmd_guardar_Click()
   Dim list_item As ListItem
   Me.frm_cambio_agente.Visible = False
   lbl_facturas = "Facturas Activas"
   var_agente = ""
   rs.Open "select DISTINCT vcha_age_agente_id, vcha_age_nombre from tb_agentes order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      numero_items_lineas = 0
      While Not rs.EOF
         Set list_item = lv_agentes.ListItems.Add(, , rs(0).Value)
         list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         rs.MoveNext:
         numero_items_lineas = numero_items_lineas + 1
       Wend
       rs.MoveFirst
       var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
       rs.Close
       If numero_items_lineas > 11 Then
          lv_agentes.ColumnHeaders(2).Width = 4200
       Else
          lv_agentes.ColumnHeaders(2).Width = 4499.71
       End If
    End If
   If var_agente <> "" Then
      rs.Open "select vcha_ser_Serie_id,vcha_car_documento,inte_car_numero, vcha_cli_nombre, char_ido_estatus from vw_inventario_documentos where vcha_age_agente_id = '" + var_agente + "' and CHAR_IDO_ESTATUS  = 'A' order by inte_car_numero", cnn, adOpenDynamic, adLockOptimistic
      numero_items_lineas = 0
      While Not rs.EOF
            Set list_item = lv_facturas.ListItems.Add(, , rs!VCHA_SER_SERIE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
            list_item.SubItems(2) = IIf(IsNull(rs!inte_Car_numero), 0, rs!inte_Car_numero)
            list_item.SubItems(3) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            list_item.SubItems(4) = ""
            list_item.SubItems(5) = IIf(IsNull(rs!char_ido_estatus), "", rs!char_ido_estatus)
            rs.MoveNext:
            numero_items_lineas = numero_items_lineas + 1
       Wend
       rs.Close
       If numero_items_lineas > 28 Then
          lv_facturas.ColumnHeaders(4).Width = 3300
       Else
          lv_facturas.ColumnHeaders(4).Width = 3100
       End If
   End If

End Sub

Private Sub cmd_imprimir_Click()
   If lv_facturas.ListItems.Count > 0 Then
      var_estatus = lv_facturas.selectedItem.SubItems(5)
         If var_estatus = "O" Then
            Set reporte = appl.OpenReport(App.Path + "\rep_inventario_documentos_reporte.rpt")
            reporte.RecordSelectionFormula = "{VW_INVENTARIO_DOCUMENTOS_REPORTE.VCHA_AGE_AGENTE_ID} = '" + Me.lv_agentes.selectedItem + "' and {VW_INVENTARIO_DOCUMENTOS_REPORTE.CHAR_IDO_ESTATUS} = 'O' AND {VW_INVENTARIO_DOCUMENTOS_REPORTE.FLOA_SAL_IMPORTE} > .01"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de inventario de documentos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("Desea exportar el reporte a excel", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_inventario_documentos_reporte.rpt")
               reporte.RecordSelectionFormula = "{VW_INVENTARIO_DOCUMENTOS_REPORTE.VCHA_AGE_AGENTE_ID} = '" + Me.lv_agentes.selectedItem + "' and {VW_INVENTARIO_DOCUMENTOS_REPORTE.CHAR_IDO_ESTATUS} = 'O' AND {VW_INVENTARIO_DOCUMENTOS_REPORTE.FLOA_SAL_IMPORTE} > .01"
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_inventario_documentos_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
            End If
            
            
         End If
         If var_estatus = "A" Then
            Set reporte = appl.OpenReport(App.Path + "\rep_inventario_documentos_reporte.rpt")
            reporte.RecordSelectionFormula = "{VW_INVENTARIO_DOCUMENTOS_REPORTE.VCHA_AGE_AGENTE_ID} = '" + Me.lv_agentes.selectedItem + "' and {VW_INVENTARIO_DOCUMENTOS_REPORTE.CHAR_IDO_ESTATUS} = 'A' AND {VW_INVENTARIO_DOCUMENTOS_REPORTE.FLOA_SAL_IMPORTE} > .009"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de inventario de documentos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
         
            var_si = MsgBox("Desea exportar el reporte a excel", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_inventario_documentos_reporte.rpt")
               reporte.RecordSelectionFormula = "{VW_INVENTARIO_DOCUMENTOS_REPORTE.VCHA_AGE_AGENTE_ID} = '" + Me.lv_agentes.selectedItem + "' and {VW_INVENTARIO_DOCUMENTOS_REPORTE.CHAR_IDO_ESTATUS} = 'A' AND {VW_INVENTARIO_DOCUMENTOS_REPORTE.FLOA_SAL_IMPORTE} > .009"
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_inventario_documentos_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
            End If
         
         
         
         End If
   Else
      MsgBox "No se encuentran facturas", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_invertir_Click()
   If var_todos_articulos = 1 Then
   Else
        var_todos_articulos = 0
   End If
   n = lv_facturas.ListItems.Count
   For i = 1 To n
      lv_facturas.ListItems.Item(i).Selected = True
      If lv_facturas.selectedItem.SubItems(4) = "*" Then
         lv_facturas.selectedItem.SubItems(4) = ""
         lv_facturas.ListItems.Item(i).Bold = False
         lv_facturas.ListItems.Item(i).ForeColor = &H80000012
         lv_facturas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_facturas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_facturas.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_facturas.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_facturas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_facturas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_facturas.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_facturas.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      Else
         lv_facturas.selectedItem.SubItems(4) = "*"
         lv_facturas.ListItems.Item(i).Bold = True
         lv_facturas.ListItems.Item(i).ForeColor = &H8000&
         lv_facturas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_facturas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_facturas.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_facturas.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_facturas.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
         lv_facturas.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
         lv_facturas.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
         lv_facturas.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   var_todos_articulos = 0
   i = lv_facturas.selectedItem.Index
   If lv_facturas.selectedItem.SubItems(4) = "*" Then
      lv_facturas.selectedItem.SubItems(4) = ""
      lv_facturas.ListItems.Item(i).Bold = False
      lv_facturas.ListItems.Item(i).ForeColor = &H80000012
      lv_facturas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_facturas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_facturas.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_facturas.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_facturas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_facturas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_facturas.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_facturas.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_facturas.Refresh
   Else
      lv_facturas.selectedItem.SubItems(4) = "*"
      lv_facturas.ListItems.Item(i).Bold = True
      lv_facturas.ListItems.Item(i).ForeColor = &H8000&
      lv_facturas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_facturas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_facturas.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_facturas.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_facturas.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
      lv_facturas.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
      lv_facturas.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
      lv_facturas.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
      lv_facturas.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_facturas.ListItems.Count
   For i = 1 To n
       lv_facturas.ListItems.Item(i).SubItems(4) = " "
       lv_facturas.ListItems.Item(i).Bold = False
       lv_facturas.ListItems.Item(i).ForeColor = &H80000012
       lv_facturas.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_facturas.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_facturas.ListItems.Item(i).ListSubItems(3).Bold = False
       lv_facturas.ListItems.Item(i).ListSubItems(4).Bold = False
       lv_facturas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_facturas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
       lv_facturas.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
       lv_facturas.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
    Next
    lv_facturas.Refresh
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   If var_todos_articulos = 1 Then
   Else
         var_todos_articulos = 0
   End If
   n = lv_facturas.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_facturas.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_facturas.selectedItem.SubItems(4) = "" And var_rellena = True Then
         lv_facturas.selectedItem.SubItems(4) = "*"
         lv_facturas.ListItems.Item(i).Bold = True
         lv_facturas.ListItems.Item(i).ForeColor = &H8000&
         lv_facturas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_facturas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_facturas.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_facturas.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_facturas.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
         lv_facturas.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
         lv_facturas.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
         lv_facturas.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
      Else
         If var_encontro = True And lv_facturas.selectedItem.SubItems(4) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_facturas.selectedItem.SubItems(4) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_todos_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_facturas.ListItems.Count
   For i = 1 To n
       lv_facturas.ListItems.Item(i).SubItems(4) = "*"
       lv_facturas.ListItems.Item(i).Bold = True
       lv_facturas.ListItems.Item(i).ForeColor = &H8000&
       lv_facturas.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_facturas.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_facturas.ListItems.Item(i).ListSubItems(3).Bold = True
       lv_facturas.ListItems.Item(i).ListSubItems(4).Bold = True
       lv_facturas.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_facturas.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
       lv_facturas.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
       lv_facturas.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
   Next
   lv_facturas.Refresh
End Sub

Private Sub Command1_Click()
   Me.frm_cambio_agente.Visible = False
   Me.frm_cambio_agente.Visible = True
   Me.txt_agente_actual = lv_agentes.selectedItem
   Me.txt_nombre_agente_actual = lv_agentes.selectedItem.SubItems(1)
   Me.txt_agente_nuevo = ""
   Me.txt_nombre_agente_nuevo = ""
   Me.txt_agente_nuevo.SetFocus
End Sub

Private Sub Command2_Click()
   If Me.txt_agente_nuevo <> "" Then
      var_si = MsgBox("¿Desea cambiar de agente las facturas seleccionadas?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar el cambio del agente", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_n = lv_facturas.ListItems.Count
            var_primera_vez = 1
            For var_i = 1 To var_n
                lv_facturas.ListItems.Item(var_i).Selected = True
                If lv_facturas.selectedItem.SubItems(4) = "*" Then
                   var_serie = lv_facturas.selectedItem
                   var_documento = lv_facturas.selectedItem.SubItems(1)
                   var_estatus = lv_facturas.selectedItem.SubItems(5)
                   rs.Open "update tb_inventario_documentos set VCHA_IDO_AGENTE_REAL = '" + Me.txt_agente_nuevo + "' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + Trim(var_documento) + "' and vcha_ser_serie_id = '" + var_serie + "' and inte_car_numero = " + lv_facturas.selectedItem.SubItems(2), cnn, adOpenDynamic, adLockOptimistic
                End If
            Next var_i
         
            
            var_agente = lv_agentes.selectedItem
            var_filtro = 0
            lbl_facturas = "Facturas Activas"
         
         
            lv_facturas.ListItems.Clear
            rs.Open "select vcha_ser_Serie_id,vcha_car_documento,inte_car_numero, vcha_cli_nombre, char_ido_estatus from vw_inventario_documentos where vcha_ido_agente_real = '" + var_agente + "' and CHAR_IDO_ESTATUS  = 'A' and vcha_emp_empresa_id = '" + var_empresa + "' order by inte_car_numero", cnn, adOpenDynamic, adLockOptimistic
            numero_items_lineas = 0
            While Not rs.EOF
                  Set list_item = lv_facturas.ListItems.Add(, , rs!VCHA_SER_SERIE_ID)
                  list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
                  list_item.SubItems(2) = IIf(IsNull(rs!inte_Car_numero), 0, rs!inte_Car_numero)
                  list_item.SubItems(3) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                  list_item.SubItems(4) = ""
                  list_item.SubItems(5) = IIf(IsNull(rs!char_ido_estatus), "", rs!char_ido_estatus)
                  rs.MoveNext:
                  numero_items_lineas = numero_items_lineas + 1
             Wend
             rs.Close
             If numero_items_lineas > 28 Then
                lv_facturas.ColumnHeaders(4).Width = 3300
             Else
                lv_facturas.ColumnHeaders(4).Width = 3100
             End If
         
         
         End If
      End If
   Else
      MsgBox "No se selecciono un agente", vbOKOnly, "ATENCION"
   End If
   Me.frm_cambio_agente.Visible = False
End Sub

Private Sub Command3_Click()
   Me.frm_cambio_agente.Visible = False
End Sub

Private Sub Form_Load()
   Me.frm_cambio_agente.Visible = False
   Me.frm_lista.Visible = False
   var_cadena_seguridad = ""
   var_filtro = 0
   Top = 0
   Left = 0
   Dim list_item As ListItem
   lbl_facturas = "Facturas Activas"
   var_agente = ""
   rs.Open "select DISTINCT vcha_age_agente_id, vcha_age_nombre from tb_agentes where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      numero_items_lineas = 0
      While Not rs.EOF
         Set list_item = lv_agentes.ListItems.Add(, , rs(0).Value)
         list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         rs.MoveNext:
         numero_items_lineas = numero_items_lineas + 1
       Wend
       rs.MoveFirst
       var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
       rs.Close
       If numero_items_lineas > 11 Then
          lv_agentes.ColumnHeaders(2).Width = 4200
       Else
          lv_agentes.ColumnHeaders(2).Width = 4499.71
       End If
    End If
   If var_agente <> "" Then
      rs.Open "select vcha_ser_Serie_id,vcha_car_documento,inte_car_numero, vcha_cli_nombre, char_ido_estatus from vw_inventario_documentos where vcha_age_agente_id = '" + var_agente + "' and CHAR_IDO_ESTATUS  = 'A' order by inte_car_numero", cnn, adOpenDynamic, adLockOptimistic
      numero_items_lineas = 0
      While Not rs.EOF
            Set list_item = lv_facturas.ListItems.Add(, , rs!VCHA_SER_SERIE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
            list_item.SubItems(2) = IIf(IsNull(rs!inte_Car_numero), 0, rs!inte_Car_numero)
            list_item.SubItems(3) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            list_item.SubItems(4) = ""
            list_item.SubItems(5) = IIf(IsNull(rs!char_ido_estatus), "", rs!char_ido_estatus)
            rs.MoveNext:
            numero_items_lineas = numero_items_lineas + 1
       Wend
       rs.Close
       If numero_items_lineas > 28 Then
          lv_facturas.ColumnHeaders(4).Width = 3300
       Else
          lv_facturas.ColumnHeaders(4).Width = 3100
       End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_inventario_documentos)
End Sub

Private Sub lv_agentes_GotFocus()
   Me.frm_cambio_agente.Visible = False
   Me.lbl_agente = lv_agentes.selectedItem.SubItems(1)
End Sub

Private Sub lv_agentes_ItemClick(ByVal Item As MSComctlLib.ListItem)
   lv_facturas.ListItems.Clear
   var_agente = lv_agentes.selectedItem
   var_filtro = 0
   lbl_facturas = "Facturas Activas"
   If var_agente <> "" Then
      Me.lbl_agente = lv_agentes.selectedItem.SubItems(1)
      rs.Open "select vcha_ser_Serie_id,vcha_car_documento,inte_car_numero, vcha_cli_nombre, char_ido_estatus from vw_inventario_documentos where vcha_ido_agente_real = '" + var_agente + "' and CHAR_IDO_ESTATUS  = 'A' and vcha_emp_empresa_id = '" + var_empresa + "' and floa_sal_importe > 0.01 order by inte_car_numero", cnn, adOpenDynamic, adLockOptimistic
      numero_items_lineas = 0
      While Not rs.EOF
            Set list_item = lv_facturas.ListItems.Add(, , rs!VCHA_SER_SERIE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
            list_item.SubItems(2) = IIf(IsNull(rs!inte_Car_numero), 0, rs!inte_Car_numero)
            list_item.SubItems(3) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            list_item.SubItems(4) = ""
            list_item.SubItems(5) = IIf(IsNull(rs!char_ido_estatus), "", rs!char_ido_estatus)
            rs.MoveNext:
            numero_items_lineas = numero_items_lineas + 1
       Wend
       rs.Close
       If numero_items_lineas > 28 Then
          lv_facturas.ColumnHeaders(4).Width = 3300
       Else
          lv_facturas.ColumnHeaders(4).Width = 3100
       End If
   End If
End Sub

Private Sub lv_facturas_GotFocus()
   Me.frm_cambio_agente.Visible = False
End Sub

Private Sub lv_facturas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim numero_lineas As Integer
      Dim numero_seleccionado1 As Integer
      Dim numero_seleccionado2 As Integer
      Dim primera_vez As Boolean
      Dim segunda_vez As Boolean
      Dim i As Integer
      Dim n As Integer
      Dim list_item As ListItem
      n = lv_facturas.ListItems.Count
      i = lv_facturas.selectedItem.Index
      If lv_facturas.ListItems.Item(i).SubItems(4) = "*" Then
      lv_facturas.ListItems.Item(i).SubItems(4) = " "
             lv_facturas.ListItems.Item(i).Bold = False
             lv_facturas.ListItems.Item(i).ForeColor = &H80000012
             lv_facturas.ListItems.Item(i).ListSubItems(1).Bold = False
             lv_facturas.ListItems.Item(i).ListSubItems(2).Bold = False
             lv_facturas.ListItems.Item(i).ListSubItems(3).Bold = False
             lv_facturas.ListItems.Item(i).ListSubItems(4).Bold = False
             lv_facturas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
             lv_facturas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
             lv_facturas.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
             lv_facturas.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
          Else
             lv_facturas.ListItems.Item(i).SubItems(4) = "*"
             lv_facturas.ListItems.Item(i).Bold = True
             lv_facturas.ListItems.Item(i).ForeColor = &H8000&
             lv_facturas.ListItems.Item(i).ListSubItems(1).Bold = True
             lv_facturas.ListItems.Item(i).ListSubItems(2).Bold = True
             lv_facturas.ListItems.Item(i).ListSubItems(3).Bold = True
             lv_facturas.ListItems.Item(i).ListSubItems(4).Bold = True
             lv_facturas.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
             lv_facturas.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
             lv_facturas.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
             lv_facturas.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
         End If
      lv_facturas.Refresh
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_agente_nuevo = lv_lista.selectedItem
      Me.txt_nombre_agente_nuevo = lv_lista.selectedItem.SubItems(1)
      Me.txt_agente_nuevo.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.txt_agente_nuevo.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_agente_actual_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_cambio_agente.Visible = False
   End If
End Sub

Private Sub txt_agente_nuevo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 5
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_agente_nuevo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre_agente_nuevo.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.frm_cambio_agente.Visible = False
   End If
End Sub

Private Sub txt_agente_nuevo_LostFocus()
   If Me.txt_agente_nuevo <> "" Then
      rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + Me.txt_agente_nuevo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         Me.txt_nombre_agente_nuevo = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
      Else
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
         Me.txt_agente_nuevo = ""
         Me.txt_nombre_agente_nuevo = ""
      End If
      rsaux.Close
   End If
End Sub

Private Sub txt_nombre_agente_actual_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_cambio_agente.Visible = False
   End If
End Sub

Private Sub txt_nombre_agente_nuevo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_cambio_agente.Visible = False
   Else
      If KeyAscii = 13 Then
         Me.Command2.SetFocus
      Else
         KeyAscii = 0
      End If
   End If
End Sub

