VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_dias_promedio_pago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Dias Promedio de Pago"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_mensaje 
      Height          =   2535
      Left            =   60
      TabIndex        =   21
      Top             =   375
      Width           =   6510
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Procesando Información Espere un momento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1230
         Left            =   90
         TabIndex        =   22
         Top             =   630
         Width           =   6285
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   465
      TabIndex        =   18
      Top             =   255
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   19
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
         TabIndex        =   20
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Datos para el Reporte "
      Height          =   1830
      Left            =   120
      TabIndex        =   13
      Top             =   405
      Width           =   6375
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   855
         TabIndex        =   7
         Top             =   1350
         Width           =   1065
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   1950
         TabIndex        =   8
         Top             =   1350
         Width           =   4170
      End
      Begin VB.TextBox txt_titular 
         Height          =   315
         Left            =   855
         TabIndex        =   5
         Top             =   990
         Width           =   1065
      End
      Begin VB.TextBox txt_nombre_titular 
         Height          =   315
         Left            =   1950
         TabIndex        =   6
         Top             =   990
         Width           =   4170
      End
      Begin VB.TextBox txt_nombre_grupo 
         Height          =   315
         Left            =   1950
         TabIndex        =   4
         Top             =   630
         Width           =   4170
      End
      Begin VB.TextBox txt_grupo 
         Height          =   315
         Left            =   855
         TabIndex        =   3
         Top             =   630
         Width           =   1065
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   1950
         TabIndex        =   2
         Top             =   270
         Width           =   4170
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   855
         TabIndex        =   1
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   1410
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   1050
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   690
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   330
         Width           =   555
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6135
      Picture         =   "frmreporte_dias_promedio_pago.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmreporte_dias_promedio_pago.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   0
      TabIndex        =   12
      Top             =   330
      Width           =   6510
   End
   Begin VB.Frame Frame1 
      Caption         =   "Año"
      Height          =   645
      Left            =   120
      TabIndex        =   0
      Top             =   2250
      Width           =   1755
      Begin VB.TextBox txt_año 
         Height          =   360
         Left            =   75
         TabIndex        =   9
         Top             =   210
         Width           =   1605
      End
   End
End
Attribute VB_Name = "frmreporte_dias_promedio_pago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_tipo_lista As Integer

Private Sub cmd_imprimir_Click()
   If Trim(txt_agente) <> "" Or Trim(txt_grupo) <> "" Or Trim(txt_cliente) <> "" Or Trim(txt_titular) <> "" Then
      If IsNumeric(Me.txt_año) Then
         Me.frm_mensaje.Visible = True
         Me.Refresh
         cnn.CommandTimeout = 360
         cnn.BeginTrans
         rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_REPORTE_DIAS_PROMEDIO", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
         Else
            var_consecutivo = 1
         End If
         rs.Close
         rs.Open "insert into TB_TEMP_REPORTE_DIAS_PROMEDIO (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
         If Trim(Me.txt_titular) = "" Then
            rs.Open "exec SP_REPORTE_DIAS_PROMEDIO_PAGO " + CStr(var_consecutivo) + "," + txt_año
         Else
            rs.Open "exec SP_REPORTE_DIAS_PROMEDIO_PAGO " + CStr(var_consecutivo) + "," + txt_año
            'rs.Open "exec SP_REPORTE_DIAS_PROMEDIO_PAGO_TITULAR " + CStr(var_consecutivo) + "," + txt_año
         End If
         Me.frm_mensaje.Visible = False
         If Trim(txt_agente) <> "" Then
            Set reporte = appl.OpenReport(App.Path + "\rep_dias_pago_promedio.rpt")
            reporte.RecordSelectionFormula = "{VW_REPORTE_DIAS_PROMEDIO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {VW_REPORTE_DIAS_PROMEDIO.VCHA_AGE_AGENTE_ID} = '" + txt_agente + "'"
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\reporte_dias_pago_promedio" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
         End If
         If Trim(txt_grupo) <> "" Then
            Set reporte = appl.OpenReport(App.Path + "\rep_dias_pago_promedio.rpt")
            reporte.RecordSelectionFormula = "{VW_REPORTE_DIAS_PROMEDIO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {VW_REPORTE_DIAS_PROMEDIO.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + txt_grupo + "'"
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\reporte_dias_pago_promedio" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
         End If
         If Trim(txt_titular) <> "" Then
            Set reporte = appl.OpenReport(App.Path + "\rep_dias_pago_promedio.rpt")
            reporte.RecordSelectionFormula = "{VW_REPORTE_DIAS_PROMEDIO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {VW_REPORTE_DIAS_PROMEDIO.VCHA_TIT_TITULAR_ID} = '" + txt_titular + "'"
            'Set reporte = appl.OpenReport(App.Path + "\rep_dias_pago_promedio_titular.rpt")
            'reporte.RecordSelectionFormula = "{VW_REPORTE_DIAS_PROMEDIO_titular.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {VW_REPORTE_DIAS_PROMEDIO_titular.Expr1} = '01' and {VW_REPORTE_DIAS_PROMEDIO_titular.vcha_emp_empresa_id} = '" + var_empresa + "'"
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\reporte_dias_pago_promedio" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            
            'Set reporte = appl.OpenReport(App.Path + "\rep_dias_pago_promedio_titular.rpt")
            'reporte.RecordSelectionFormula = "{VW_REPORTE_DIAS_PROMEDIO_titular.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {VW_REPORTE_DIAS_PROMEDIO_titular.Expr1} = '02' and {VW_REPORTE_DIAS_PROMEDIO_titular.vcha_emp_empresa_id} = '" + var_empresa + "'"
            'For ntablas = 1 To reporte.Database.Tables.Count
            '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            'Next ntablas
            'reporte.ExportOptions.FormatType = crEFTExcel80
            'reporte.ExportOptions.DestinationType = crEDTDiskFile
            'archivo = "c:\reporte_dias_pago_promedio" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            'reporte.ExportOptions.DiskFileName = archivo
            'reporte.Export False
            'Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
         End If
         If Trim(txt_cliente) <> "" Then
            Set reporte = appl.OpenReport(App.Path + "\rep_dias_pago_promedio.rpt")
            reporte.RecordSelectionFormula = "{VW_REPORTE_DIAS_PROMEDIO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {VW_REPORTE_DIAS_PROMEDIO.VCHA_cLI_CLAVE_ID} = '" + txt_cliente + "'"
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\reporte_dias_pago_promedio" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
         End If
      Else
         MsgBox "Año incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado información", vbOKOnly, "ATENCION"
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
   
   
   
   
   frm_lista.Visible = False
   Top = 2000
   Left = 2500
   Me.frm_mensaje.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_reporte_valuacion_devoluciones)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_agente = lv_lista.selectedItem
            txt_nombre_agente = lv_lista.selectedItem.SubItems(1)
         Else
            txt_agente = ""
            txt_nombre_agente = ""
         End If
         txt_agente.SetFocus
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_grupo = lv_lista.selectedItem
            txt_nombre_grupo = lv_lista.selectedItem.SubItems(1)
         Else
            txt_grupo = ""
            txt_nombre_grupos = ""
         End If
         txt_grupo.SetFocus
      End If
      If var_tipo_lista = 3 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_titular = lv_lista.selectedItem
            txt_nombre_titular = lv_lista.selectedItem.SubItems(1)
         Else
            txt_titular = ""
            txt_nombre_titular = ""
         End If
         txt_titular.SetFocus
      End If
      If var_tipo_lista = 4 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_cliente = lv_lista.selectedItem
            txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
         Else
            txt_cliente = ""
            txt_nombre_cliente = ""
         End If
         txt_cliente.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      If var_tipo_lista = 1 Then
         txt_agente.SetFocus
      End If
      If var_tipo_lista = 2 Then
         txt_grupo.SetFocus
      End If
      If var_tipo_lista = 3 Then
         txt_titular.SetFocus
      End If
      If var_tipo_lista = 4 Then
         txt_cliente.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_AGENTES order by vcha_AGE_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 1
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
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agente_LostFocus()
   If Trim(txt_agente) <> "" Then
      txt_nombre_agente = ""
      txt_grupo = ""
      txt_nombre_grupo = ""
      txt_titular = ""
      txt_nombre_titular = ""
      txt_cliente = ""
      txt_nombre_cliente = ""
      rs.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      Else
         MsgBox "Clave de agente no existe", vbOKOnly, "ATENCION"
         txt_agente = ""
         txt_nombre_agente = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_año_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_CLIENTES order by vcha_CLI_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 4
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

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_cliente_LostFocus()
   If Trim(txt_cliente) <> "" Then
      txt_nombre_grupo = ""
      txt_agente = ""
      txt_nombre_agente = ""
      txt_titular = ""
      txt_nombre_titular = ""
      txt_nombre_cliente = ""
      txt_grupo = ""
      rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
      Else
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
         txt_cliente = ""
         txt_nombre_cliente = ""
      End If
      rs.Close
   Else
      txt_nombre_cliente = ""
   End If
End Sub

Private Sub txt_grupo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_GRUPOSACTUALES order by vcha_GAC_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_GAC_GRUPO_aCTUAL_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "GRUPOS"
      var_tipo_lista = 2
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

Private Sub txt_grupo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_grupo_LostFocus()
   If Trim(txt_grupo) <> "" Then
      txt_nombre_grupo = ""
      txt_agente = ""
      txt_nombre_agente = ""
      txt_titular = ""
      txt_nombre_titular = ""
      txt_cliente = ""
      txt_nombre_cliente = ""
      rs.Open "SELECT * FROM TB_GRUPOSACTUALES WHERE VCHA_GAC_GRUPO_ACTUAL_ID = '" + txt_grupo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_grupo = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
      Else
         MsgBox "Clave de grupo incorrecto", vbOKOnly, "ATENCION"
         txt_nombre_grupo = ""
         txt_grupo = ""
      End If
      rs.Close
   Else
      txt_nombre_grupo = ""
   End If
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_AGENTES order by vcha_AGE_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 1
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
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_CLIENTES order by vcha_CLI_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CLI_CLIENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 4
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

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_grupo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_GRUPOSACTUALES order by vcha_GAC_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_GAC_GRUPO_aCTUAL_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "GRUPOS"
      var_tipo_lista = 2
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

Private Sub txt_nombre_grupo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_titular_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TITULARES order by vcha_TIT_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tit_titular_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TITULARES"
      var_tipo_lista = 3
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

Private Sub txt_nombre_titular_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_titular_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TITULARES order by vcha_TIT_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tit_titular_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TITULARES"
      var_tipo_lista = 3
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

Private Sub txt_titular_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_titular_LostFocus()
   If Trim(txt_titular) <> "" Then
      txt_nombre_grupo = ""
      txt_agente = ""
      txt_nombre_agente = ""
      txt_nombre_titular = ""
      txt_cliente = ""
      txt_nombre_cliente = ""
      txt_grupo = ""
      rs.Open "select * from tb_titulares where vcha_tit_titular_id = '" + txt_titular + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_titular = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
      Else
         MsgBox "Clave de titular incorrecta", vbOKOnly, "ATENCION"
         txt_titualar = ""
         txt_nombre_titular = ""
      End If
      rs.Close
   Else
      txt_nombre_titular = ""
   End If
End Sub
