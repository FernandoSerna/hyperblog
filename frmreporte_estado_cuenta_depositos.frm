VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_estado_cuenta_depositos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado de cuenta de deposito en el ORACLE"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   870
      TabIndex        =   8
      Top             =   15
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   9
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
         TabIndex        =   10
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Clientes "
      Height          =   3405
      Left            =   135
      TabIndex        =   7
      Top             =   1395
      Width           =   7215
      Begin MSComctlLib.ListView lv_clientes 
         Height          =   3045
         Left            =   90
         TabIndex        =   2
         Top             =   270
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   5371
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
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Referencia"
            Object.Width           =   2645
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Agente "
      Height          =   900
      Left            =   105
      TabIndex        =   6
      Top             =   435
      Width           =   7245
      Begin VB.TextBox txt_nombre_agente 
         Height          =   345
         Left            =   1095
         TabIndex        =   1
         Top             =   375
         Width           =   6060
      End
      Begin VB.TextBox txt_agente 
         Height          =   345
         Left            =   165
         TabIndex        =   0
         Top             =   375
         Width           =   900
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6990
      Picture         =   "frmreporte_estado_cuenta_depositos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmreporte_estado_cuenta_depositos.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   75
      TabIndex        =   5
      Top             =   390
      Width           =   7320
   End
End
Attribute VB_Name = "frmreporte_estado_cuenta_depositos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_dia As String
   Dim var_mes As String
   Dim var_año As String
   Dim var_disponible As Double
   Dim var_real As Double
   Dim var_referencia As String
   If Trim(Me.txt_agente) <> "" Then
      If Me.lv_clientes.ListItems.Count > 0 Then
         var_referencia = lv_clientes.selectedItem.SubItems(2)
         rsaux.Open "select IDENTIFICADOR, CLAVE, CANAL, REFERENCIA, DESCRIPCION, IMPORTE_ABONO, IMPORTE_CARGO, IMPORTE_A_DISPONIBLE, DESCRIPCION_DETALLADA, FOLIO, ORIGEN, TO_CHAR(FECHA_DE_DEPOSITO,'DD/MM/YYYY') AS FECHA_DE_DEPOSITO, TO_CHAR(FECHA_DE_AUTORIZACION,'DD/MM/YYYY') AS FECHA_DE_AUTORIZACION, TIPO_DE_AFECTACION, FOLIO_DOCUMENTO, CIE_SUC from vw_edo_cuenta where referencia = '" + Trim(var_referencia) + "'", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rs.Open "SELECT * FROM TB_SALDO WHERE VCHA_SAL_REFERENCIA = '" + var_referencia + "'", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
            var_disponible = 0
            var_real = 0
            If Not rs.EOF Then
               var_disponible = IIf(IsNull(rs!NUMB_SAL_IMPORTE_DISPONIBLE), 0, rs!NUMB_SAL_IMPORTE_DISPONIBLE)
               var_real = IIf(IsNull(rs!NUMB_SAL_IMPORTE), 0, rs!NUMB_SAL_IMPORTE)
            Else
              var_disponible = 0
              var_real = 0
            End If
            rs.Close
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_ESTADO_CUENTA_DEPOSITOS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
   
            rs.Open "insert into TB_TEMP_ESTADO_CUENTA_DEPOSITOS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            While Not rsaux.EOF
                  var_cadena = "INSERT INTO TB_TEMP_ESTADO_CUENTA_DEPOSITOS (FLOA_TEM_SALDO_DISPONIBLE, FLOA_TEM_SALDO_REAL, VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, INTE_TEM_CONSECUTIVO, IDENTIFICADOR, CLAVE, CANAL, REFERENCIA, DESCRIPCION, IMPORTE_ABONO, IMPORTE_CARGO, IMPORTE_DISPONIBLE, DESCRIPCION_DETALLADA, FOLIO, ORIGEN, FECHA_DE_DEPOSITO, FECHA_DE_AUTORIZACION, TIPO_DE_AFECTACION, FOLIO_DOCUMENTO, CIE_SUC)"
                  var_cadena = var_cadena + " Values (" + CStr(var_disponible) + "," + CStr(var_real) + ",'" + Me.txt_agente + "','" + Me.txt_nombre_agente + "','" + Me.lv_clientes.selectedItem + "','" + Me.lv_clientes.selectedItem.SubItems(1) + "', " + CStr(var_consecutivo) + ", " + CStr(IIf(IsNull(rsaux!Identificador), 0, rsaux!Identificador)) + ",'" + IIf(IsNull(rsaux!CLAVE), "", rsaux!CLAVE) + "','" + IIf(IsNull(rsaux!canal), "", rsaux!canal) + "','" + IIf(IsNull(rsaux!Referencia), "", rsaux!Referencia) + "','" + IIf(IsNull(rsaux!descripcion), "", rsaux!descripcion) + "'," + CStr(IIf(IsNull(rsaux!IMPORTE_ABONO), 0, rsaux!IMPORTE_ABONO)) + ", " + CStr(IIf(IsNull(rsaux!importe_Cargo), 0, rsaux!importe_Cargo)) + ","
                  var_cadena = var_cadena + CStr(IIf(IsNull(rsaux!IMPORTE_A_DISPONIBLE), 0, rsaux!IMPORTE_A_DISPONIBLE)) + ",'" + IIf(IsNull(rsaux!DESCRIPCION_DETALLADA), "", rsaux!DESCRIPCION_DETALLADA) + "','" + IIf(IsNull(rsaux!FOLIO), "", rsaux!FOLIO) + "','" + IIf(IsNull(rsaux!Origen), "", rsaux!Origen) + "',"
                  var_dia = Mid(rsaux!FECHA_de_DEPOSITO, 1, 2)
                  var_mes = Mid(rsaux!FECHA_de_DEPOSITO, 4, 2)
                  var_año = "20" + Mid(rsaux!FECHA_de_DEPOSITO, 9, 2)
                  var_fecha_deposito = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                  var_cadena = var_cadena + var_fecha_deposito
         
                  var_dia = Mid(rsaux!FECHA_de_AUTORIZACION, 1, 2)
                  var_mes = Mid(rsaux!FECHA_de_AUTORIZACION, 4, 2)
                  var_año = "20" + Mid(rsaux!FECHA_de_AUTORIZACION, 9, 2)
                  var_fecha_autorizacion = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
         
         
                  var_cadena = var_cadena + "," + var_fecha_autorizacion + ",'" + IIf(IsNull(rsaux!TIPO_DE_AFECTACION), "", rsaux!TIPO_DE_AFECTACION) + "','" + IIf(IsNull(rsaux!FOLIO_DOCUMENTO), "", rsaux!FOLIO_DOCUMENTO) + "','" + IIf(IsNull(rsaux!CIE_SUC), "", rsaux!CIE_SUC) + "')"
                  rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rsaux.MoveNext
            Wend
            
            
            Set reporte = appl.OpenReport(App.Path + "\REP_ESTADO_CUENTA_DEPOSITOS.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_ESTADO_CUENTA_DEPOSITOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {TB_TEMP_ESTADO_CUENTA_DEPOSITOS.VCHA_AGE_AGENTE_ID} = '" + Me.txt_agente + "'"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte Consecutivo de Movimientos por Agente (cargos)"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("¿Desea importar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\REP_ESTADO_CUENTA_DEPOSITOS.rpt")
               reporte.RecordSelectionFormula = "{TB_TEMP_ESTADO_CUENTA_DEPOSITOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {TB_TEMP_ESTADO_CUENTA_DEPOSITOS.VCHA_AGE_AGENTE_ID} = '" + Me.txt_agente + "'"
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_consecutivo_movimientos_cargos" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            rs.Open "delete from TB_TEMP_ESTADO_CUENTA_DEPOSITOS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "No existe información del cliente seleccionado", vbOKOnly, "ATENCION"
         End If
         rsaux.Close
      Else
         MsgBox "Se debe de seleccionar un cliente", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Se debe de indicar un agente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.frm_lista.Visible = False
   Top = 1500
   Left = 2500
   If cnn_clientes_tiendas.State = 0 Then
      cnn_clientes_tiendas.Open var_conexion_pedidos_tiendas
      cnn_clientes_tiendas.CursorLocation = adUseClient
   End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_clientes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_lista.ListItems.Count > 0 Then
         Me.txt_agente = lv_lista.selectedItem
         Me.txt_nombre_agente = lv_lista.selectedItem.SubItems(1)
      End If
      Me.txt_agente.SetFocus
      Me.frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      Me.frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_agente_Change()
   Me.lv_clientes.ListItems.Clear
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes WHERE VCHA_TAG_TIPOAGENTE_ID = 'TDA' or vcha_Age_agente_id = '00028' ORDER by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 100
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

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   pro_enfoque (KeyAscii)
End Sub

Private Sub txt_agente_LostFocus()
   Dim list_item As ListItem
   If Trim(Me.txt_agente) <> "" Then
      rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + Me.txt_agente + "' AND  VCHA_TAG_TIPOAGENTE_ID = 'TDA' or vcha_age_Agente_id = '00028' ", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         Me.txt_nombre_agente = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
         rs.Open "select vcha_cli_clave_id, vcha_cli_nombre, vcha_cli_referencia from tb_clientes where vcha_age_agente_id = '" + Me.txt_agente + "' AND (VCHA_CLI_REFERENCIA IS NOT NULL OR VCHA_CLI_REFERENCIA <> '') order by vcha_CLI_nombre", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.lv_clientes.ListItems.Clear
            While Not rs.EOF
                  Set list_item = lv_clientes.ListItems.Add(, , rs(0).Value)
                  list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
                  list_item.SubItems(2) = Trim(IIf(IsNull(rs(2).Value), "", rs(2).Value))
                  rs.MoveNext:
            Wend
         Else
            Me.lv_clientes.ListItems.Clear
         End If
         rs.Close
      Else
         MsgBox "El agente no existe", vbOKOnly, "ATENCION"
         Me.lv_clientes.ListItems.Clear
      End If
      rsaux.Close
   Else
      lv_clientes.ListItems.Clear
   End If
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes WHERE VCHA_TAG_TIPOAGENTE_ID = 'TDA' order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 100
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

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   pro_enfoque (KeyAscii)
End Sub
