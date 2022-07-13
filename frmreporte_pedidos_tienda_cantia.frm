VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_pedidos_tienda_cantia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte pedidos tienda Cantia"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmreporte_pedidos_tienda_cantia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      Picture         =   "frmreporte_pedidos_tienda_cantia.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   120
      TabIndex        =   8
      Top             =   3315
      Width           =   5640
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   1095
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3375
         TabIndex        =   12
         Top             =   330
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   660
         TabIndex        =   11
         Top             =   330
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   90
      TabIndex        =   15
      Top             =   285
      Width           =   5670
   End
   Begin VB.Frame Frame3 
      Caption         =   "  Agentes "
      Height          =   2880
      Left            =   120
      TabIndex        =   0
      Top             =   390
      Width           =   5625
      Begin VB.Frame Frame6 
         Height          =   120
         Left            =   30
         TabIndex        =   6
         Top             =   540
         Width           =   5565
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         Picture         =   "frmreporte_pedidos_tienda_cantia.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   120
         Picture         =   "frmreporte_pedidos_tienda_cantia.frx":0952
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         Picture         =   "frmreporte_pedidos_tienda_cantia.frx":0A54
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   780
         Picture         =   "frmreporte_pedidos_tienda_cantia.frx":0B26
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Marcar (Enter)"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Picture         =   "frmreporte_pedidos_tienda_cantia.frx":0D70
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   225
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_vendedore 
         Height          =   2025
         Left            =   45
         TabIndex        =   7
         Top             =   690
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   3572
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
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmreporte_pedidos_tienda_cantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
            var_cadena_agentes = ""
            For var_j = 1 To Me.lv_vendedore.ListItems.Count
                Me.lv_vendedore.ListItems.Item(var_j).Selected = True
                If Me.lv_vendedore.selectedItem.SubItems(2) = "*" Then
                   If var_cadena_agentes = "" Then
                      var_cadena_agentes = "'" + Me.lv_vendedore.selectedItem + "'"
                   Else
                      var_cadena_agentes = var_cadena_agentes + ",'" + Me.lv_vendedore.selectedItem + "'"
                   End If
                End If
            Next var_j
            If var_cadena_agentes <> "" Then
               cnn.BeginTrans
               rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_REPORTE_PEDIDOS_TIENDA_CANTIA", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
               Else
                  var_consecutivo = 0
               End If
               var_consecutivo = var_consecutivo + 1
               rs.Close
               rs.Open "insert into TB_TEMP_REPORTE_PEDIDOS_TIENDA_CANTIA (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               var_dia = CStr(Day(CDate(txt_inicio)))
               var_mes = CStr(Month(CDate(txt_inicio)))
               var_año = CStr(Year(CDate(txt_inicio)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
               var_fecha_fin_1 = CDate(txt_fin) + 1
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
               
               var_cadena = "INSERT INTO TB_TEMP_REPORTE_PEDIDOS_TIENDA_CANTIA (INTE_TEM_CONSECUTIVO, DTIM_TEM_fECHA_INICIO, DTIM_TEM_FECHA_FIN, INTE_PED_NUMERO, VCHA_PED_CLIENTE, VCHA_PED_TELEFONO, VCHA_USU_USUARIO_ID, VCHA_aRT_aRTICULO_ID, FLOA_PED_CANTIDAD, DTIM_PED_FECHA, VCHA_USU_NOMBRE, VCHA_ART_NOMBRE_ESPAÑOL) "
               var_cadena = var_cadena + "SELECT " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + "-.00001, dbo.TB_PEDIDO_TIENDA_CANTIA.INTE_PED_NUMERO, dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_PED_CLIENTE, dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_PED_TELEFONO, dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_USU_USUARIO_ID, dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_ART_ARTICULO_ID, dbo.TB_PEDIDO_TIENDA_CANTIA.FLOA_PED_CANTIDAD, dbo.TB_PEDIDO_TIENDA_CANTIA.DTIM_PED_FECHA, dbo.TB_USUARIOS_PEDIDOS_CANTIA.VCHA_USU_NOMBRE, dbo.TB_Articulos.VCHA_ART_NOMBRE_ESPAÑOL FROM dbo.TB_PEDIDO_TIENDA_CANTIA INNER JOIN dbo.TB_USUARIOS_PEDIDOS_CANTIA ON dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_USU_USUARIO_ID = dbo.TB_USUARIOS_PEDIDOS_CANTIA.VCHA_USU_USUARIO_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_PEDIDO_TIENDA_CANTIA.DTIM_PED_FECHA >= " + var_fecha_inicio + ") AND  (dbo.TB_PEDIDO_TIENDA_CANTIA.DTIM_PED_FECHA < " + var_fecha_fin + ") AND"
               var_cadena = var_cadena + "(dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_USU_USUARIO_ID IN (" + var_cadena_agentes + "))"
               rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               
               Set reporte = appl.OpenReport(App.Path + "\rep_pedidos_tiendas_cantia.rpt")
               reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_PEDIDOS_TIENDA_CANTIA.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\pedidos_tienda_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               rs.Open "delete from TB_TEMP_REPORTE_PEDIDOS_TIENDA_CANTIA where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            
            
            
            Else
               MsgBox "No se han seleccionado vendedores", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "La fecha final no debe de ser inferior a la de inicio", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_invertir_Click()
   n = lv_vendedore.ListItems.Count
   For i = 1 To n
      lv_vendedore.ListItems.Item(i).Selected = True
      If lv_vendedore.selectedItem.SubItems(2) = "*" Then
         lv_vendedore.selectedItem.SubItems(2) = ""
         lv_vendedore.ListItems.Item(i).Bold = False
         lv_vendedore.ListItems.Item(i).ForeColor = &H80000012
         lv_vendedore.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_vendedore.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_vendedore.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_vendedore.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_vendedore.selectedItem.SubItems(2) = "*"
         lv_vendedore.ListItems.Item(i).Bold = True
         lv_vendedore.ListItems.Item(i).ForeColor = &HFF0000
         lv_vendedore.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_vendedore.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_vendedore.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_vendedore.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_vendedore.selectedItem.Index
   If lv_vendedore.selectedItem.SubItems(2) = "*" Then
      lv_vendedore.selectedItem.SubItems(2) = ""
      lv_vendedore.ListItems.Item(i).Bold = False
      lv_vendedore.ListItems.Item(i).ForeColor = &H80000012
      lv_vendedore.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_vendedore.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_vendedore.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_vendedore.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_vendedore.Refresh
   Else
      lv_vendedore.selectedItem.SubItems(2) = "*"
      lv_vendedore.ListItems.Item(i).Bold = True
      lv_vendedore.ListItems.Item(i).ForeColor = &HFF0000
      lv_vendedore.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_vendedore.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_vendedore.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_vendedore.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_vendedore.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_vendedore.ListItems.Count
   For i = 1 To n
      lv_vendedore.ListItems.Item(i).Selected = True
      lv_vendedore.selectedItem.SubItems(2) = ""
      lv_vendedore.ListItems.Item(i).Bold = False
      lv_vendedore.ListItems.Item(i).ForeColor = &H80000012
      lv_vendedore.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_vendedore.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_vendedore.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_vendedore.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_vendedore.Refresh
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_vendedore.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_vendedore.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_vendedore.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_vendedore.selectedItem.SubItems(2) = "*"
         lv_vendedore.ListItems.Item(i).Bold = True
         lv_vendedore.ListItems.Item(i).ForeColor = &HFF0000
         lv_vendedore.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_vendedore.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_vendedore.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_vendedore.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_vendedore.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_vendedore.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_todos_Click()
   n = lv_vendedore.ListItems.Count
   For i = 1 To n
      lv_vendedore.ListItems.Item(i).Selected = True
      lv_vendedore.selectedItem.SubItems(2) = "*"
      lv_vendedore.ListItems.Item(i).Bold = True
      lv_vendedore.ListItems.Item(i).ForeColor = &HFF0000
      lv_vendedore.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_vendedore.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_vendedore.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_vendedore.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_vendedore.Refresh
End Sub

Private Sub Form_Load()
   Top = 1400
   Left = 3000
   rs.Open "select * from TB_USUARIOS_PEDIDOS_CANTIA order by vcha_usu_nombre", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      Set list_item = Me.lv_vendedore.ListItems.Add(, , rs!vcha_usu_usuario_ID)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_USU_NOMBRE), "", rs!VCHA_USU_NOMBRE)
      list_item.SubItems(2) = ""
      rs.MoveNext:
   Wend
   rs.Close
   If lv_vendedore.ListItems.Count > 7 Then
      lv_vendedore.ColumnHeaders(2).Width = 4220
   Else
      lv_vendedore.ColumnHeaders(2).Width = 4499.71
   End If
   Me.txt_fin = Date
   Me.txt_inicio = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub lv_vendedore_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_vendedore, ColumnHeader)
End Sub

Private Sub lv_vendedore_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_vendedore.selectedItem.Index
      If lv_vendedore.selectedItem.SubItems(2) = "*" Then
         lv_vendedore.selectedItem.SubItems(2) = ""
         lv_vendedore.ListItems.Item(i).Bold = False
         lv_vendedore.ListItems.Item(i).ForeColor = &H80000012
         lv_vendedore.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_vendedore.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_vendedore.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_vendedore.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_vendedore.Refresh
      Else
         lv_vendedore.selectedItem.SubItems(2) = "*"
         lv_vendedore.ListItems.Item(i).Bold = True
         lv_vendedore.ListItems.Item(i).ForeColor = &HFF0000
         lv_vendedore.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_vendedore.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_vendedore.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_vendedore.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_vendedore.Refresh
      End If
   End If
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes = CDate(Me.txt_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub
