VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_catalogo_articulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rerporte de Catálogo de Articulos"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5805
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5370
      Picture         =   "frmreporte_catalogo_articulos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   " Movimientos "
      Height          =   2445
      Left            =   60
      TabIndex        =   2
      Top             =   405
      Width           =   5685
      Begin VB.Frame Frame7 
         Height          =   120
         Left            =   15
         TabIndex        =   8
         Top             =   525
         Width           =   5640
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   435
         Picture         =   "frmreporte_catalogo_articulos.frx":063A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   105
         Picture         =   "frmreporte_catalogo_articulos.frx":0850
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1095
         Picture         =   "frmreporte_catalogo_articulos.frx":0952
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         Picture         =   "frmreporte_catalogo_articulos.frx":0A24
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Marcar (Enter)"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1425
         Picture         =   "frmreporte_catalogo_articulos.frx":0C6E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   210
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_lineas 
         Height          =   1620
         Left            =   45
         TabIndex        =   9
         Top             =   720
         Width           =   5595
         _ExtentX        =   9869
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
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   345
      Width           =   5760
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmreporte_catalogo_articulos.frx":0E84
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "frmreporte_catalogo_articulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_i As Integer, var_n As Integer, var_contado As Integer
   Dim var_cadena As String
   var_n = lv_lineas.ListItems.Count
   var_contador = 0
   var_cadena = ""
   For var_i = 1 To var_n
      lv_lineas.ListItems.Item(var_i).Selected = True
      If Trim(lv_lineas.selectedItem.SubItems(2)) = "*" Then
         If var_contador = 0 Then
            var_cadena = "{VW_CATALOGO_ARTICULOS.VCHA_LIN_LINEA_ID} = '" + Trim(lv_lineas.selectedItem) + "'"
            var_contador = 1
         Else
            var_cadena = var_cadena + " or {VW_CATALOGO_ARTICULOS.VCHA_LIN_LINEA_ID} = '" + Trim(lv_lineas.selectedItem) + "'"
         End If
      End If
   Next var_i
   If Trim(var_cadena) <> "" Then
      Set reporte = appl.OpenReport(App.Path + "\rep_catalogo_Articulos.rpt")
      reporte.RecordSelectionFormula = var_cadena
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Ordenes de surtido pendientes de empacar o facturar"
      frmvistasprevias.Show 1
      Set reporte = Nothing
   Else
      MsgBox "No se a seleccionado una linea", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
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
   primera_vez = False
   segunda_vez = False
   n = lv_lineas.ListItems.Count
   For i = 1 To n
       If lv_lineas.ListItems.Item(i).SubItems(2) = "*" And primera_vez = False Then
          numero_seleccionado1 = i
          primera_vez = True
       End If
       If lv_lineas.ListItems.Item(i).SubItems(2) = "*" And primera_vez = True Then
          numero_seleccionado2 = i
       End If
   Next
   For i = numero_seleccionado1 To numero_seleccionado2
       lv_lineas.ListItems.Item(i).SubItems(2) = "*"
       lv_lineas.ListItems.Item(i).Bold = True
       lv_lineas.ListItems.Item(i).ForeColor = &H8000&
       lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
       lv_lineas.Refresh
   Next
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
   i = lv_lineas.selectedItem.Index
   If lv_lineas.selectedItem.SubItems(2) = "*" Then
       lv_lineas.ListItems.Item(i).SubItems(2) = " "
       lv_lineas.ListItems.Item(i).Bold = False
       lv_lineas.ListItems.Item(i).ForeColor = &H80000012
       lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Else
      lv_lineas.ListItems.Item(i).SubItems(2) = "*"
      lv_lineas.ListItems.Item(i).Bold = True
      lv_lineas.ListItems.Item(i).ForeColor = &H8000&
      lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
      lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
  End If
  lv_lineas.Refresh
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
   n = lv_lineas.ListItems.Count
   For i = 1 To n
       If lv_lineas.ListItems.Item(i).SubItems(2) = "*" Then
          lv_lineas.ListItems.Item(i).SubItems(2) = " "
          lv_lineas.ListItems.Item(i).Bold = False
          lv_lineas.ListItems.Item(i).ForeColor = &H80000012
          lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = False
          lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = False
          lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
          lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
       Else
          lv_lineas.ListItems.Item(i).SubItems(2) = "*"
          lv_lineas.ListItems.Item(i).Bold = True
          lv_lineas.ListItems.Item(i).ForeColor = &H8000&
          lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
          lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
          lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
          lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
      End If
   Next
   lv_lineas.Refresh
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
   n = lv_lineas.ListItems.Count
   For i = 1 To n
       lv_lineas.ListItems.Item(i).SubItems(2) = " "
       lv_lineas.ListItems.Item(i).Bold = False
       lv_lineas.ListItems.Item(i).ForeColor = &H80000012
       lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
    Next
    lv_lineas.Refresh
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
   n = lv_lineas.ListItems.Count
   For i = 1 To n
       lv_lineas.ListItems.Item(i).SubItems(2) = "*"
       lv_lineas.ListItems.Item(i).Bold = True
       lv_lineas.ListItems.Item(i).ForeColor = &H8000&
       lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
   Next
   lv_lineas.Refresh
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 2000
   Left = 3000
   rs.Open "select * from tb_lineas order by vcha_lin_nombre", cnn, adOpenDynamic, adLockOptimistic
   numero_items_lineas = 0
   While Not rs.EOF
      Set list_item = lv_lineas.ListItems.Add(, , rs!vcha_lin_linea_id)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_lin_NOMBRE), "", rs!VCHA_lin_NOMBRE)
      list_item.SubItems(2) = ""
      rs.MoveNext:
      numero_items_lineas = numero_items_lineas + 1
    Wend
   rs.Close
   If numero_items_lineas > 12 Then
      lv_lineas.ColumnHeaders(2).Width = 4200.71
   Else
      lv_lineas.ColumnHeaders(2).Width = 4499.71
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro = False
   Call activa_forma(var_activa_forma_reporte_catalogo_articulos)
End Sub
