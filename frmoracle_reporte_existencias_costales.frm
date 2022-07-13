VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_reporte_existencias_costales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Existencias costales"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Height          =   840
      Left            =   90
      TabIndex        =   12
      Top             =   5250
      Width           =   9435
      Begin VB.TextBox txt_fecha 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3810
         TabIndex        =   14
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2805
         TabIndex        =   13
         Top             =   315
         Width           =   840
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmoracle_reporte_existencias_costales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9210
      Picture         =   "frmoracle_reporte_existencias_costales.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   120
      Left            =   60
      TabIndex        =   9
      Top             =   270
      Width           =   9495
   End
   Begin VB.Frame Frame1 
      Height          =   4800
      Left            =   75
      TabIndex        =   0
      Top             =   345
      Width           =   9450
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Picture         =   "frmoracle_reporte_existencias_costales.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   570
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   690
         Picture         =   "frmoracle_reporte_existencias_costales.frx":0952
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar (Enter)"
         Top             =   570
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         Picture         =   "frmoracle_reporte_existencias_costales.frx":0B9C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   570
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   30
         Picture         =   "frmoracle_reporte_existencias_costales.frx":0C6E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   570
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         Picture         =   "frmoracle_reporte_existencias_costales.frx":0D70
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   570
         Width           =   330
      End
      Begin VB.Frame Frame2 
         Height          =   60
         Left            =   15
         TabIndex        =   3
         Top             =   885
         Width           =   9375
      End
      Begin MSComctlLib.ListView lv_rutas 
         Height          =   3765
         Left            =   45
         TabIndex        =   1
         Top             =   975
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   6641
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   13229
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Tipo"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   " Rutas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   30
         TabIndex        =   2
         Top             =   135
         Width           =   9330
      End
   End
End
Attribute VB_Name = "frmoracle_reporte_existencias_costales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter

Private Sub cmd_imprimir_Click()
   If IsDate(Me.txt_fecha) Then
      var_posible = 0
      VAR_CADENA_RUTAS = ""
      For var_j = 1 To Me.lv_rutas.ListItems.Count
          Me.lv_rutas.ListItems.Item(var_j).Selected = True
          If Me.lv_rutas.selectedItem.SubItems(2) = "*" Then
             var_posible = 1
             If VAR_CADENA_RUTAS = "" Then
                VAR_CADENA_RUTAS = "'" + Me.lv_rutas.selectedItem + "'"
             Else
                VAR_CADENA_RUTAS = VAR_CADENA_RUTAS + ",'" + Me.lv_rutas.selectedItem + "'"
             End If
          End If
      Next var_j
      If var_posible = 1 Then
         cnn.BeginTrans
         rs.Open "select max(inte_tem_consecutivo) from tb_temp_oracle_existencias_costales", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
         Else
            var_consecutivo = 1
         End If
         rs.Close
         rs.Open "insert into tb_temp_oracle_existencias_costales (inte_tem_consecutivo) values ( " + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
         var_dia_str = CStr(Day(CDate(Me.txt_fecha) + 1))
         var_mes_str = CStr(Month(CDate(Me.txt_fecha) + 1))
         var_año_str = CStr(Year(CDate(Me.txt_fecha) + 1))
         If Len(var_dia_str) = 1 Then
            var_dia_str = "0" + var_dia_str
         End If
         If Len(var_mes_str) = 1 Then
            var_mes_str = "0" + var_mes_str
         End If
         If Len(var_año_str) = 2 Then
            var_año_str = "20" + CStr(var_año_str)
         End If
         var_fecha = var_dia_str + "/" + var_mes_str + "/" + var_año_str
         
         var_dia_str = CStr(Day(CDate(Me.txt_fecha)))
         var_mes_str = CStr(Month(CDate(Me.txt_fecha)))
         var_año_str = CStr(Year(CDate(Me.txt_fecha)))
         If Len(var_dia_str) = 1 Then
            var_dia_str = "0" + var_dia_str
         End If
         If Len(var_mes_str) = 1 Then
            var_mes_str = "0" + var_mes_str
         End If
         If Len(var_año_str) = 2 Then
            var_año_str = "20" + CStr(var_año_str)
         End If
         var_fecha_real = "{d '" + var_año_str + "-" + var_mes_str + "-" + var_dia_str + "'}"
         
         rs.Open "select vcha_almacen_origen, vcha_tipo_costal, sum(numb_cantidad) as cantidad from xxvia_tb_control_costales where date_fecha_creacion < to_date('" + var_fecha + "','DD/MM/YYYY') AND numb_org_id = 93 and vcha_almacen_origen in (" + VAR_CADENA_RUTAS + ") GROUP BY  vcha_almacen_origen, vcha_tipo_costal", cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux.Open "insert into tb_temp_oracle_existencias_costales (inte_tem_consecutivo, ruta, tipo_bulto, cantidad, FECHA) values (" + CStr(var_consecutivo) + ",'" + rs!vcha_almacen_origen + "','" + rs!vcha_tipo_costal + "'," + CStr(rs!Cantidad) + "," + var_fecha_real + ")", cnn, adOpenDynamic, adLockOptimistic
               rs.MoveNext
         Wend
         rs.Close
         rs.Open "delete from tb_temp_oracle_existencias_costales  where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ruta is null", cnn, adOpenDynamic, adLockOptimistic
         rs.Open "select distinct ruta from tb_temp_oracle_existencias_costales where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               strconsulta = "select secondary_inventory_name, description from MTL_SECONDARY_INVENTORIES where organization_id = ? and secondary_inventory_name = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(IIf(IsNull(rs!RUTA), "", rs!RUTA)))
                    .Parameters.Append parametro
               End With
               Set rsaux = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               If Not rsaux.EOF Then
                  var_nombre = rsaux!Description
                  rsaux.Close
               Else
                  strconsulta = "select NAME AS nombre_ruta, salesrep_id as clave_ruta from jtf_rs_salesreps where salesrep_id = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(IIf(IsNull(rs!RUTA), 0, rs!RUTA)))
                       .Parameters.Append parametro
                  End With
                  Set rsaux = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If Not rsaux.EOF Then
                     var_nombre = rsaux!nombre_ruta
                  Else
                     var_nombre = ""
                  End If
                  rsaux.Close
               End If
               
               rsaux.Open "update tb_temp_oracle_existencias_costales set nombre_ruta = '" + IIf(IsNull(var_nombre), "", var_nombre) + "' where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ruta = '" + IIf(IsNull(rs!RUTA), "", rs!RUTA) + "'", cnn, adOpenDynamic, adLockOptimistic
               
               rs.MoveNext
         Wend
         rs.Close
         Set reporte = appl.OpenReport(App.Path + "\rep_oracle_existencias_bultos.rpt")
         reporte.RecordSelectionFormula = "{VW_TEMP_ORACLE_EXISTENCIAS_BULTOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Pedidos cargados"
         frmvistasprevias.Show 1
         Set reporte = Nothing
    
         var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_existencias_bultos_excel.rpt")
            reporte.RecordSelectionFormula = "{VW_TEMP_ORACLE_EXISTENCIAS_BULTOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\existencias_bultos_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
         End If
         rs.Open "delete from tb_temp_oracle_existencias_costales  where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      End If
   End If
End Sub

Private Sub cmd_invertir_Click()
   n = lv_rutas.ListItems.Count
   For i = 1 To n
      lv_rutas.ListItems.Item(i).Selected = True
      If lv_rutas.selectedItem.SubItems(2) = "*" Then
         lv_rutas.selectedItem.SubItems(2) = ""
         lv_rutas.ListItems.Item(i).Bold = False
         lv_rutas.ListItems.Item(i).ForeColor = &H80000012
         lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_rutas.selectedItem.SubItems(2) = "*"
         lv_rutas.ListItems.Item(i).Bold = True
         lv_rutas.ListItems.Item(i).ForeColor = &HFF0000
         lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
   If Me.lv_rutas.ListItems.Count > 0 Then
      Me.lv_rutas.SetFocus
   End If
End Sub

Private Sub cmd_marcar_Click()
   i = lv_rutas.selectedItem.Index
   If lv_rutas.selectedItem.SubItems(2) = "*" Then
      lv_rutas.selectedItem.SubItems(2) = ""
      lv_rutas.ListItems.Item(i).Bold = False
      lv_rutas.ListItems.Item(i).ForeColor = &H80000012
      lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_rutas.Refresh
   Else
      lv_rutas.selectedItem.SubItems(2) = "*"
      lv_rutas.ListItems.Item(i).Bold = True
      lv_rutas.ListItems.Item(i).ForeColor = &HFF0000
      lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_rutas.Refresh
   End If
   If Me.lv_rutas.ListItems.Count > 0 Then
      Me.lv_rutas.SetFocus
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_rutas.ListItems.Count
   For i = 1 To n
      lv_rutas.ListItems.Item(i).Selected = True
      lv_rutas.selectedItem.SubItems(2) = ""
      lv_rutas.ListItems.Item(i).Bold = False
      lv_rutas.ListItems.Item(i).ForeColor = &H80000012
      lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_rutas.Refresh
   If Me.lv_rutas.ListItems.Count > 0 Then
      Me.lv_rutas.SetFocus
   End If

End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_rutas.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_rutas.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_rutas.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_rutas.selectedItem.SubItems(2) = "*"
         lv_rutas.ListItems.Item(i).Bold = True
         lv_rutas.ListItems.Item(i).ForeColor = &HFF0000
         lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_rutas.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_rutas.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
   If Me.lv_rutas.ListItems.Count > 0 Then
      Me.lv_rutas.SetFocus
   End If
End Sub

Private Sub cmd_todos_Click()
   n = lv_rutas.ListItems.Count
   For i = 1 To n
      lv_rutas.ListItems.Item(i).Selected = True
      lv_rutas.selectedItem.SubItems(2) = "*"
      lv_rutas.ListItems.Item(i).Bold = True
      lv_rutas.ListItems.Item(i).ForeColor = &HFF0000
      lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_rutas.Refresh
   If Me.lv_rutas.ListItems.Count > 0 Then
      Me.lv_rutas.SetFocus
   End If
End Sub

Private Sub Form_Load()
   Top = 400
   Left = 1100
   Me.txt_fecha = Date
   rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   'strconsulta = "select distinct numb_movimiento_id, numb_org_id, vcha_almacen_origen, vcha_origen_movimiento from xxvia_tb_control_costales where numb_org_id = ? and numb_movimiento_id in 89 and numb_cantidad >0"
   strconsulta = "select distinct numb_movimiento_id, numb_org_id, vcha_almacen_origen, vcha_origen_movimiento from xxvia_tb_control_costales where numb_org_id = ? "
   With comandoORA
       .ActiveConnection = cnnoracle_4
       .CommandType = adCmdText
       .CommandText = strconsulta
       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_unidad_organizacional))
       .Parameters.Append parametro
   End With
   Set rs = comandoORA.execute
   Set comandoORA = Nothing
   Set parametro = Nothing
   'rs.Open "select distinct numb_movimiento_id, numb_org_id, vcha_almacen_origen from xxvia_tb_control_costales where numb_org_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         'If rs!vcha_origen_movimiento = "PI" Then
         '    strconsulta = "select secondary_inventory_name, description from MTL_SECONDARY_INVENTORIES where organization_id = ? and secondary_inventory_name = ?"
         '    With comandoORA
         '         .ActiveConnection = cnnoracle_4
         '         .CommandType = adCmdText
         '         .CommandText = strconsulta
         '         Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
         '         .Parameters.Append parametro
         '         Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!vcha_almacen_origen))
         '         .Parameters.Append parametro
         '    End With
         '    Set rsaux = comandoORA.execute
         '    Set comandoORA = Nothing
         '    Set parametro = Nothing
         '    var_clave = CStr(rs!vcha_almacen_origen)
         '    var_nombre = rsaux!Description
         '    rsaux.Close
         'End If
         'If rs!vcha_origen_movimiento = "PM" Then
             strconsulta = "select * from xxvia_vw_rutas where clave = ?"
             With comandoORA
                  .ActiveConnection = cnnoracle_4
                  .CommandType = adCmdText
                  .CommandText = strconsulta
                  Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!vcha_almacen_origen))
                  .Parameters.Append parametro
             End With
             Set rsaux = comandoORA.execute
             Set comandoORA = Nothing
             Set parametro = Nothing
             var_clave = CStr(rs!vcha_almacen_origen)
             var_nombre = rsaux!descripcion
             rsaux.Close
         'End If
         Set list_item = lv_rutas.ListItems.Add(, , var_clave)
         list_item.SubItems(1) = IIf(IsNull(var_nombre), "", var_nombre)
         list_item.SubItems(2) = ""
         list_item.SubItems(3) = rs!vcha_origen_movimiento
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_rutas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_rutas, ColumnHeader)
End Sub

Private Sub lv_rutas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_rutas.selectedItem.Index
      If lv_rutas.selectedItem.SubItems(2) = "*" Then
         lv_rutas.selectedItem.SubItems(2) = ""
         lv_rutas.ListItems.Item(i).Bold = False
         lv_rutas.ListItems.Item(i).ForeColor = &H80000012
         lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_rutas.Refresh
      Else
         lv_rutas.selectedItem.SubItems(2) = "*"
         lv_rutas.ListItems.Item(i).Bold = True
         lv_rutas.ListItems.Item(i).ForeColor = &HFF0000
         lv_rutas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_rutas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_rutas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_rutas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_rutas.Refresh
      End If
      If Me.lv_rutas.ListItems.Count > 0 Then
         Me.lv_rutas.SetFocus
      End If
   End If
End Sub

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         frmcalendario.mes = CDate(Me.txt_fecha)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      Me.txt_fecha = var_fecha_general
   End If
End Sub
