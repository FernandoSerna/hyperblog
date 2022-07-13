VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_mercancia_empacada_transito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mercancia empacada"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   5745
   Begin VB.Frame Frame3 
      Caption         =   " Filtrado "
      Height          =   765
      Left            =   2895
      TabIndex        =   22
      Top             =   390
      Width           =   2775
      Begin VB.OptionButton opt_tipo_filtrado_2 
         Caption         =   "Canal de Venta"
         Height          =   255
         Left            =   135
         TabIndex        =   24
         Top             =   465
         Width           =   1935
      End
      Begin VB.OptionButton opt_tipo_filtrado_1 
         Caption         =   "Agente"
         Height          =   345
         Left            =   135
         TabIndex        =   23
         Top             =   165
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5295
      Picture         =   "frmreporte_mercancia_empacada_transito.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmreporte_mercancia_empacada_transito.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   15
      TabIndex        =   0
      Top             =   270
      Width           =   5700
   End
   Begin VB.Frame Frame2 
      Caption         =   " Tipo de Reporte "
      Height          =   765
      Left            =   75
      TabIndex        =   3
      Top             =   390
      Width           =   2775
      Begin VB.OptionButton opt_detalle 
         Caption         =   " Detallado por Artículo"
         Height          =   210
         Left            =   105
         TabIndex        =   5
         Top             =   465
         Width           =   1965
      End
      Begin VB.OptionButton opt_linea 
         Caption         =   " Agrupado por Lineas"
         Height          =   270
         Left            =   105
         TabIndex        =   4
         Top             =   165
         Width           =   1920
      End
   End
   Begin VB.Frame frm_agentes 
      Height          =   4305
      Left            =   60
      TabIndex        =   6
      Top             =   1170
      Width           =   5610
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Picture         =   "frmreporte_mercancia_empacada_transito.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   690
         Picture         =   "frmreporte_mercancia_empacada_transito.frx":0952
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Marcar (Enter)"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         Picture         =   "frmreporte_mercancia_empacada_transito.frx":0B9C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   30
         Picture         =   "frmreporte_mercancia_empacada_transito.frx":0C6E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         Picture         =   "frmreporte_mercancia_empacada_transito.frx":0D70
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   360
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_agentes 
         Height          =   3510
         Left            =   75
         TabIndex        =   12
         Top             =   720
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   6191
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
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6967
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   " Agentes "
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   13
         Top             =   120
         Width           =   5535
      End
   End
   Begin VB.Frame frm_canales 
      Height          =   4305
      Left            =   60
      TabIndex        =   14
      Top             =   1170
      Width           =   5610
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         Picture         =   "frmreporte_mercancia_empacada_transito.frx":0F86
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   30
         Picture         =   "frmreporte_mercancia_empacada_transito.frx":119C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         Picture         =   "frmreporte_mercancia_empacada_transito.frx":129E
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   690
         Picture         =   "frmreporte_mercancia_empacada_transito.frx":1370
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Marcar (Enter)"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Picture         =   "frmreporte_mercancia_empacada_transito.frx":15BA
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   360
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_canales 
         Height          =   3495
         Left            =   75
         TabIndex        =   20
         Top             =   720
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   6165
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
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6967
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Canales de Venta"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   21
         Top             =   120
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmreporte_mercancia_empacada_transito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_i As Integer, var_n As Integer, var_contador As Integer
   var_contador = 0
   var_primera_vez = 1
   If opt_tipo_filtrado_2.Value = True Then
      var_n = lv_canales.ListItems.Count
      For var_i = 1 To var_n
          lv_canales.ListItems(var_i).Selected = True
          If lv_canales.selectedItem.SubItems(2) = "*" Then
             rs.Open "select vcha_age_agente_id from vw_clientes_1 where vcha_can_canal_venta_id = '" + lv_canales.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
             While Not rs.EOF
                If Not IsNull(rs!VCHA_AGE_AGENTE_ID) Then
                   If var_primera_vez = 1 Then
                      var_primera_vez = 0
                      var_cadena = "vcha_age_agente_id = '" + rs!VCHA_AGE_AGENTE_ID + "'"
                   Else
                      var_cadena = var_cadena + " or vcha_age_agente_id = '" + rs!VCHA_AGE_AGENTE_ID + "'"
                   End If
                End If
                rs.MoveNext
             Wend
             rs.Close
          End If
      Next var_i
   End If
   var_primera_vez = 1
   If opt_tipo_filtrado_1.Value = True Then
      var_n = lv_agentes.ListItems.Count
      For var_i = 1 To var_n
          lv_agentes.ListItems.Item(var_i).Selected = True
          If lv_agentes.selectedItem.SubItems(2) = "*" Then
             If var_primera_vez = 1 Then
                var_primera_vez = 0
                var_cadena = "vcha_age_agente_id = '" + lv_agentes.selectedItem + "'"
             Else
                var_cadena = var_cadena + " or vcha_age_agente_id = '" + lv_agentes.selectedItem + "'"
             End If
          End If
      Next var_i
   End If
   
   If Trim(var_cadena) <> "" Then
      cnn.BeginTrans
      rs.Open "select max(INTE_TEM_CONSECTUTVO) as numero from TB_TEMP_MERCANCIA_TRANSITO_AGENTES", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
      Else
         var_consecutivo = 0
      End If
      rs.Close
      var_consecutivo = var_consecutivo + 1
      rs.Open "insert into TB_TEMP_MERCANCIA_TRANSITO_AGENTES (INTE_TEM_CONSECTUTVO) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
      cnn.CommitTrans
      rs.Open "select * from tb_agentes where " + var_cadena, cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            rsaux.Open "insert into TB_TEMP_MERCANCIA_TRANSITO_AGENTES (INTE_TEM_CONSECTUTVO, vcha_age_agente_id) values (" + CStr(var_consecutivo) + ",'" + rs!VCHA_AGE_AGENTE_ID + "')", cnn, adOpenDynamic, adLockOptimistic
            rs.MoveNext
      Wend
      rs.Close
      If Me.opt_tipo_filtrado_1 = True Then
         If opt_linea = True Then
            Set reporte = appl.OpenReport(App.Path + "\rep_mercancia_empacada_transito_linea_agente.rpt")
            reporte.RecordSelectionFormula = "{VW_MERCANCIA_EMPACADA_TRANSITO_LINEA.INTE_TEM_CONSECTUTVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
         End If
         If opt_detalle = True Then
            Set reporte = appl.OpenReport(App.Path + "\rep_mercancia_empacada_transito_detalle_agente.rpt")
            reporte.RecordSelectionFormula = "{VW_MERCANCIA_EMPACADA_TRANSITO_DETALLE.INTE_TEM_CONSECTUTVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
         End If
      End If
      If Me.opt_tipo_filtrado_2 = True Then
         If opt_linea = True Then
            Set reporte = appl.OpenReport(App.Path + "\rep_mercancia_empacada_transito_linea_canal.rpt")
            reporte.RecordSelectionFormula = "{VW_MERCANICA_EMPACADA_TRANSITO_LINEA_CANAL.INTE_TEM_CONSECTUTVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
         End If
         If opt_detalle = True Then
            Set reporte = appl.OpenReport(App.Path + "\rep_mercancia_empacada_transito_detalle_canal.rpt")
            reporte.RecordSelectionFormula = "{VW_MERCANCIA_EMPACADA_TRANSITO_DETALLE_CANAL.INTE_TEM_CONSECTUTVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
         End If
      End If
      rs.Open "delete from TB_TEMP_MERCANCIA_TRANSITO_AGENTES  where INTE_TEM_CONSECTUTVO = " + CStr(var_consecutivo) + "", cnn, adOpenDynamic, adLockOptimistic
   Else
      If Me.opt_tipo_filtrado_1 = True Then
         MsgBox "No se a seleccionado algun agente", vbOKOnly, "ATENCION"
      End If
      If Me.opt_tipo_filtrado_2 = True Then
         MsgBox "No se a seleccionado algun canal de venta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub cmd_invertir_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_agentes.ListItems.Count
   For i = 1 To n
       If lv_agentes.ListItems.Item(i).SubItems(2) = "*" Then
          lv_agentes.ListItems.Item(i).SubItems(2) = " "
          lv_agentes.ListItems.Item(i).Bold = False
          lv_agentes.ListItems.Item(i).ForeColor = &H80000012
          lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
          lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
          lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
          lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
       Else
          lv_agentes.ListItems.Item(i).SubItems(2) = "*"
          lv_agentes.ListItems.Item(i).Bold = True
          lv_agentes.ListItems.Item(i).ForeColor = &H8000&
          lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
          lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
          lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
          lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
      End If
   Next
   lv_agentes.Refresh
End Sub

Private Sub cmd_marcar_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   i = lv_agentes.selectedItem.Index
   If lv_agentes.selectedItem.SubItems(2) = "*" Then
       lv_agentes.ListItems.Item(i).SubItems(2) = " "
       lv_agentes.ListItems.Item(i).Bold = False
       lv_agentes.ListItems.Item(i).ForeColor = &H80000012
       lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Else
      lv_agentes.ListItems.Item(i).SubItems(2) = "*"
      lv_agentes.ListItems.Item(i).Bold = True
      lv_agentes.ListItems.Item(i).ForeColor = &H8000&
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
  End If
  lv_agentes.Refresh
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
   n = lv_agentes.ListItems.Count
   For i = 1 To n
       lv_agentes.ListItems.Item(i).SubItems(2) = " "
       lv_agentes.ListItems.Item(i).Bold = False
       lv_agentes.ListItems.Item(i).ForeColor = &H80000012
       lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
    Next
    lv_agentes.Refresh
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   primera_vez = False
   segunda_vez = False
   n = lv_agentes.ListItems.Count
   For i = 1 To n
       If lv_agentes.ListItems.Item(i).SubItems(2) = "*" And primera_vez = False Then
          numero_seleccionado1 = i
          primera_vez = True
       End If
       If lv_agentes.ListItems.Item(i).SubItems(2) = "*" And primera_vez = True Then
          numero_seleccionado2 = i
       End If
   Next
   For i = numero_seleccionado1 To numero_seleccionado2
       lv_agentes.ListItems.Item(i).SubItems(2) = "*"
       lv_agentes.ListItems.Item(i).Bold = True
       lv_agentes.ListItems.Item(i).ForeColor = &H8000&
       lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
       lv_agentes.Refresh
   Next
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
   n = lv_agentes.ListItems.Count
   For i = 1 To n
       lv_agentes.ListItems.Item(i).SubItems(2) = "*"
       lv_agentes.ListItems.Item(i).Bold = True
       lv_agentes.ListItems.Item(i).ForeColor = &H8000&
       lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
   Next
   lv_agentes.Refresh
End Sub

Private Sub Command1_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_canales.ListItems.Count
   For i = 1 To n
       lv_canales.ListItems.Item(i).SubItems(2) = "*"
       lv_canales.ListItems.Item(i).Bold = True
       lv_canales.ListItems.Item(i).ForeColor = &H8000&
       lv_canales.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_canales.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
   Next
   lv_canales.Refresh
End Sub

Private Sub Command2_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_canales.ListItems.Count
   For i = 1 To n
       lv_canales.ListItems.Item(i).SubItems(2) = " "
       lv_canales.ListItems.Item(i).Bold = False
       lv_canales.ListItems.Item(i).ForeColor = &H80000012
       lv_canales.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_canales.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
    Next
    lv_canales.Refresh
End Sub

Private Sub Command3_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_canales.ListItems.Count
   For i = 1 To n
       If lv_canales.ListItems.Item(i).SubItems(2) = "*" Then
          lv_canales.ListItems.Item(i).SubItems(2) = " "
          lv_canales.ListItems.Item(i).Bold = False
          lv_canales.ListItems.Item(i).ForeColor = &H80000012
          lv_canales.ListItems.Item(i).ListSubItems(1).Bold = False
          lv_canales.ListItems.Item(i).ListSubItems(2).Bold = False
          lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
          lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
       Else
          lv_canales.ListItems.Item(i).SubItems(2) = "*"
          lv_canales.ListItems.Item(i).Bold = True
          lv_canales.ListItems.Item(i).ForeColor = &H8000&
          lv_canales.ListItems.Item(i).ListSubItems(1).Bold = True
          lv_canales.ListItems.Item(i).ListSubItems(2).Bold = True
          lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
          lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
      End If
   Next
   lv_canales.Refresh
End Sub

Private Sub Command4_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   i = lv_canales.selectedItem.Index
   If lv_canales.selectedItem.SubItems(2) = "*" Then
      lv_canales.ListItems.Item(i).SubItems(2) = " "
      lv_canales.ListItems.Item(i).Bold = False
      lv_canales.ListItems.Item(i).ForeColor = &H80000012
      lv_canales.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_canales.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Else
      lv_canales.ListItems.Item(i).SubItems(2) = "*"
      lv_canales.ListItems.Item(i).Bold = True
      lv_canales.ListItems.Item(i).ForeColor = &H8000&
      lv_canales.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_canales.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
      lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
   End If
   lv_canales.Refresh
End Sub

Private Sub Command5_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   primera_vez = False
   segunda_vez = False
   n = lv_canales.ListItems.Count
   For i = 1 To n
       If lv_canales.ListItems.Item(i).SubItems(2) = "*" And primera_vez = False Then
          numero_seleccionado1 = i
          primera_vez = True
       End If
       If lv_canales.ListItems.Item(i).SubItems(2) = "*" And primera_vez = True Then
          numero_seleccionado2 = i
       End If
   Next
   For i = numero_seleccionado1 To numero_seleccionado2
       lv_canales.ListItems.Item(i).SubItems(2) = "*"
       lv_canales.ListItems.Item(i).Bold = True
       lv_canales.ListItems.Item(i).ForeColor = &H8000&
       lv_canales.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_canales.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
       lv_canales.Refresh
   Next
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Me.opt_linea = True
   Me.opt_tipo_filtrado_1 = True
   Top = 1000
   Left = 3000
   rs.Open "select * from TB_AGENTES WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'ORDER BY VCHA_AGE_NOMBRE ", cnn, adOpenDynamic, adLockOptimistic
   numero_items_agentes = 0
   While Not rs.EOF
      Set list_item = lv_agentes.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      list_item.SubItems(2) = " "
      rs.MoveNext:
      numero_items_agentes = numero_items_agentes + 1
    Wend
    rs.Close
   If numero_items_agentes > 12 Then
      lv_agentes.ColumnHeaders(1).Width = lv_agentes.ColumnHeaders(1).Width - 200
   End If

   rs.Open "select * from TB_CANALESVENTAS ORDER BY VCHA_CAN_NOMBRE ", cnn, adOpenDynamic, adLockOptimistic
   numero_items_canales = 0
   While Not rs.EOF
      Set list_item = lv_canales.ListItems.Add(, , rs!vcha_can_canal_venta_id)
      list_item.SubItems(1) = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
      list_item.SubItems(2) = " "
      rs.MoveNext:
      numero_items_canales = numero_items_canales + 1
    Wend
    rs.Close
   If numero_items_canales > 12 Then
      lv_canales.ColumnHeaders(1).Width = lv_canales.ColumnHeaders(1).Width - 200
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro = False
   Call activa_forma(var_activa_forma_reporte_mercancia_empacada_transito)
End Sub

Private Sub lv_agentes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_agentes, ColumnHeader)
End Sub

Private Sub opt_tipo_filtrado_1_Click()
   If opt_tipo_filtrado_1.Value = True Then
      frm_agentes.Visible = True
      frm_canales.Visible = False
   Else
      frm_agentes.Visible = False
      frm_canales.Visible = True
   End If
End Sub

Private Sub opt_tipo_filtrado_2_Click()
   If opt_tipo_filtrado_1.Value = True Then
      frm_agentes.Visible = True
      frm_canales.Visible = False
   Else
      frm_agentes.Visible = False
      frm_canales.Visible = True
   End If
End Sub
