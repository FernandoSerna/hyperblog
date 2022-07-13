VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_estado_cuenta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado de Cuenta"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2055
      Left            =   0
      TabIndex        =   9
      Top             =   -45
      Width           =   4665
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1620
         Left            =   30
         TabIndex        =   10
         Top             =   390
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   2858
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
         TabIndex        =   11
         Top             =   120
         Width           =   4590
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   105
      TabIndex        =   8
      Top             =   1320
      Width           =   4470
      Begin VB.TextBox txt_nombre 
         Height          =   345
         Left            =   1260
         TabIndex        =   4
         Top             =   180
         Width           =   3135
      End
      Begin VB.TextBox txt_clave 
         Height          =   345
         Left            =   30
         TabIndex        =   3
         Top             =   180
         Width           =   1200
      End
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   90
      TabIndex        =   7
      Top             =   330
      Width           =   4485
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmreporte_estado_cuenta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4185
      Picture         =   "frmreporte_estado_cuenta.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Tipo de Reporte "
      Height          =   855
      Left            =   90
      TabIndex        =   0
      Top             =   465
      Width           =   4485
      Begin VB.OptionButton opt_familia 
         Caption         =   "Familia"
         Height          =   390
         Left            =   2595
         TabIndex        =   2
         Top             =   315
         Width           =   870
      End
      Begin VB.OptionButton opt_cliente 
         Caption         =   "Cliente"
         Height          =   390
         Left            =   885
         TabIndex        =   1
         Top             =   315
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmreporte_estado_cuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_cadena As String
   var_cadena = ""
   If Me.txt_clave <> "" Then
   If Me.opt_cliente = True Then
      Set reporte = appl.OpenReport(App.Path + "\REP_ESTADO_CUENTA_SALDOS.rpt")
      reporte.RecordSelectionFormula = "{VW_ESTADO_CUENTA_SALDOS.VCHA_CLI_CLAVE_ID} = '" + txt_clave + "'"
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte Estado de Cuenta"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         Set reporte = appl.OpenReport(App.Path + "\REP_ESTADO_CUENTA_SALDOS.rpt")
         reporte.RecordSelectionFormula = "{VW_ESTADO_CUENTA_SALDOS.VCHA_CLI_CLAVE_ID} = '" + txt_clave + "'"
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         'archivo = "c:\reportessid\Reporte_estado_cuenta" & Replace(Str(Date), "/", "") & "_" & Replace(Str(Time), ":", ".") & ".exls"
          archivo = "c:\reportessid\Reporte_estado_cuenta" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
         MsgBox "Se a terminado de guardar el archivo " + archivo
     End If
   Else
      Set reporte = appl.OpenReport(App.Path + "\REP_ESTADO_CUENTA_SALDOS.rpt")
      reporte.RecordSelectionFormula = "{VW_ESTADO_CUENTA_SALDOS.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + txt_clave + "'"
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte Estado de Cuenta"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         Set reporte = appl.OpenReport(App.Path + "\REP_ESTADO_CUENTA_SALDOS.rpt")
         reporte.RecordSelectionFormula = "{VW_ESTADO_CUENTA_SALDOS.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + txt_clave + "'"
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reportessid\Reporte_estado_cuenta" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
         MsgBox "Se a terminado de guardar el archivo " + archivo
     End If
   End If
   Else
      MsgBox "Clave Incorrecta", vbOKOnly, "ATENCION"
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

   cnn.Close
   cnn.Open var_conexion_string_distribucion

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
   Top = 2000
   Left = 3200
   Me.opt_cliente = True
   frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_reporte_comisiones)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         Me.txt_clave = lv_lista.selectedItem
         Me.txt_nombre = lv_lista.selectedItem.SubItems(1)
      End If
      Me.frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub opt_cliente_Click()
   Me.txt_clave = ""
   Me.txt_nombre = ""
End Sub

Private Sub opt_familia_Click()
   Me.txt_clave = ""
   Me.txt_nombre = ""
End Sub

Private Sub txt_clave_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Me.opt_cliente = True Then
         lv_lista.ListItems.Clear
         rs.Open "select * from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "CLIENTES"
         var_tipo_lista = 1
         Dim var_n As Integer
         var_n = lv_lista.ListItems.Count
         If var_n > 6 Then
            lv_lista.ColumnHeaders(2).Width = 3250.71
         Else
            lv_lista.ColumnHeaders(2).Width = 3470.71
         End If
         frm_lista.Visible = True
         lv_lista.SetFocus
      Else
         lv_lista.ListItems.Clear
         rs.Open "select * from TB_GRUPOSACTUALES order by vcha_GAC_nombre", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_GAC_GRUPO_aCTUAL_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "GRUPOS ACTUALES"
         var_tipo_lista = 1
         var_n = lv_lista.ListItems.Count
         If var_n > 6 Then
            lv_lista.ColumnHeaders(2).Width = 4270.71
         Else
            lv_lista.ColumnHeaders(2).Width = 4499.71
         End If
         frm_lista.Visible = True
         lv_lista.SetFocus
      End If
   End If
End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_LostFocus()
   If Me.txt_clave <> "" Then
      If Me.opt_cliente.Value = True Then
         rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_nombre = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         Else
            Me.txt_clave = ""
            Me.txt_nombre = ""
            MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         rs.Open "select * from tb_gruposactuales where vcha_gac_grupo_actual_id = '" + txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_nombre = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
         Else
            Me.txt_clave = ""
            Me.txt_nombre = ""
            MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
   End If
End Sub

Private Sub txt_nombre_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Me.opt_cliente = True Then
         lv_lista.ListItems.Clear
         rs.Open "select * from tb_clientes where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "CLIENTES"
         var_tipo_lista = 1
         Dim var_n As Integer
         var_n = lv_lista.ListItems.Count
         If var_n > 6 Then
            lv_lista.ColumnHeaders(2).Width = 3250.71
         Else
            lv_lista.ColumnHeaders(2).Width = 3470.71
         End If
         frm_lista.Visible = True
         lv_lista.SetFocus
      Else
         lv_lista.ListItems.Clear
         rs.Open "select * from TB_GRUPOSACTUALES order by vcha_GAC_nombre", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_GAC_GRUPO_aCTUAL_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "GRUPOS ACTUALES"
         var_tipo_lista = 1
         var_n = lv_lista.ListItems.Count
         If var_n > 6 Then
            lv_lista.ColumnHeaders(2).Width = 4270.71
         Else
            lv_lista.ColumnHeaders(2).Width = 4499.71
         End If
         frm_lista.Visible = True
         lv_lista.SetFocus
      End If
   End If
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmd_imprimir.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub
