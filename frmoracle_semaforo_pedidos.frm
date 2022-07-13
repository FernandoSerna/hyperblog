VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_semaforo_pedidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Semaforo de pedidos"
   ClientHeight    =   10515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   23115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10515
   ScaleWidth      =   23115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   720
      Left            =   45
      TabIndex        =   2
      Top             =   -15
      Width           =   23025
      Begin VB.CommandButton Command7 
         Caption         =   "Command7"
         Height          =   420
         Left            =   10725
         TabIndex        =   31
         Top             =   225
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmd_exportar 
         Height          =   435
         Left            =   7500
         Picture         =   "frmoracle_semaforo_pedidos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Exporta a excel"
         Top             =   195
         Width           =   435
      End
      Begin VB.CommandButton Command6 
         Height          =   435
         Left            =   22440
         Picture         =   "frmoracle_semaforo_pedidos.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Ejecuta despacho de pedidos seleccionados"
         Top             =   180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton cmd_actualizar 
         Height          =   435
         Left            =   7065
         Picture         =   "frmoracle_semaforo_pedidos.frx":04A4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Actualiza Grafica"
         Top             =   188
         Width           =   435
      End
      Begin VB.TextBox txt_fin 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5430
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   195
         Width           =   1575
      End
      Begin VB.TextBox txt_inicio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   195
         Width           =   1560
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ejecutar despacho de pedidos seleccionados."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   15900
         TabIndex        =   16
         Top             =   195
         Visible         =   0   'False
         Width           =   6510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Periodo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   420
         TabIndex        =   14
         Top             =   225
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4950
         TabIndex        =   6
         Top             =   255
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1830
         TabIndex        =   5
         Top             =   255
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9795
      Left            =   45
      TabIndex        =   0
      Top             =   675
      Width           =   23025
      Begin VB.TextBox txt_leido 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   14625
         TabIndex        =   28
         Top             =   9225
         Width           =   1755
      End
      Begin VB.TextBox txt_cancelado 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   21195
         TabIndex        =   26
         Top             =   9225
         Width           =   1755
      End
      Begin VB.TextBox txt_bo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   17940
         TabIndex        =   24
         Top             =   9225
         Width           =   1755
      End
      Begin VB.TextBox txt_surtido 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   11820
         TabIndex        =   22
         Top             =   9225
         Width           =   1755
      End
      Begin VB.TextBox txt_surtir 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8850
         TabIndex        =   20
         Top             =   9225
         Width           =   1755
      End
      Begin VB.TextBox txt_pedido 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6060
         TabIndex        =   18
         Top             =   9225
         Width           =   1755
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   420
         Picture         =   "frmoracle_semaforo_pedidos.frx":05A6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   165
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   90
         Picture         =   "frmoracle_semaforo_pedidos.frx":07BC
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   165
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         Picture         =   "frmoracle_semaforo_pedidos.frx":08BE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   165
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   750
         Picture         =   "frmoracle_semaforo_pedidos.frx":0990
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Marcar (Enter)"
         Top             =   165
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         Picture         =   "frmoracle_semaforo_pedidos.frx":0BDA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   165
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Frame Frame7 
         Height          =   120
         Left            =   15
         TabIndex        =   8
         Top             =   435
         Width           =   22980
      End
      Begin MSComctlLib.ListView lv_pedidos 
         Height          =   8565
         Left            =   60
         TabIndex        =   1
         Top             =   570
         Width           =   22905
         _ExtentX        =   40402
         _ExtentY        =   15108
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
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   25
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha Pedido"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha Despacho"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Ruta"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Establecimiento"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Pedido"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Por Surtir"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Surtido"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Leido"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Back Order"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Cancelado"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Embarque"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Estatus"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Posible"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   16
            Text            =   "Negado distribucion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   17
            Text            =   "Volumen"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Header_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Type_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Type_name"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Customer_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "Rule_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Rule_name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Site_use_id_establecimiento"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Leido:"
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
         Left            =   13860
         TabIndex        =   29
         Top             =   9285
         Width           =   750
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Totales:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   3570
         TabIndex        =   27
         Top             =   9218
         Width           =   1440
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cancelado:"
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
         Left            =   19830
         TabIndex        =   25
         Top             =   9285
         Width           =   1365
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Back Order:"
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
         Left            =   16410
         TabIndex        =   23
         Top             =   9285
         Width           =   1440
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Surtido:"
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
         Left            =   10845
         TabIndex        =   21
         Top             =   9285
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Surtir:"
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
         Left            =   8055
         TabIndex        =   19
         Top             =   9285
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
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
         Left            =   5085
         TabIndex        =   17
         Top             =   9285
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmoracle_semaforo_pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim objConn As New adodb.Connection
Dim objCmd As New adodb.Command
Dim objParm As adodb.Parameter

Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_exportar_Click()
   If Me.lv_pedidos.ListItems.Count > 0 Then
      cnn.BeginTrans
      rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_ORACLE_SEMAFORO_PEDIDOS_REPORTE", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
      Else
         var_consecutivo = 0
      End If
      rs.Close
      var_consecutivo = var_consecutivo + 1
      rs.Open "insert into TB_TEMP_ORACLE_SEMAFORO_PEDIDOS_REPORTE (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
      cnn.CommitTrans
      For var_j = 1 To Me.lv_pedidos.ListItems.Count
          Me.lv_pedidos.ListItems.Item(var_j).Selected = True
          var_cadena = "insert into TB_TEMP_ORACLE_SEMAFORO_PEDIDOS_REPORTE (inte_tem_consecutivo, FECHA_INICIO, fecha_fin,pedido, FECHA_PEDIDO, fecha_despacho, ruta, cliente, establecimiento, cantidad_pedida, CANTIDAD_SURTIR, cantidad_surtida, cantidad_leida, cantidad_bo, cantidad_cancelada, embarque, estatus) "
          var_cadena = var_cadena + " values (" + CStr(var_consecutivo) + ",'" + Me.txt_inicio + "','" + Me.txt_fin + "', '" + Trim(Me.lv_pedidos.selectedItem) + "','" + Me.lv_pedidos.selectedItem.SubItems(1) + "','" + Me.lv_pedidos.selectedItem.SubItems(2) + "','" + Me.lv_pedidos.selectedItem.SubItems(3) + "','" + Me.lv_pedidos.selectedItem.SubItems(4) + "'"
          var_cadena = var_cadena + " ,'" + Me.lv_pedidos.selectedItem.SubItems(5) + "'," + Me.lv_pedidos.selectedItem.SubItems(6) + "," + Me.lv_pedidos.selectedItem.SubItems(7) + "," + Me.lv_pedidos.selectedItem.SubItems(8)
          var_cadena = var_cadena + "," + Me.lv_pedidos.selectedItem.SubItems(9) + "," + Me.lv_pedidos.selectedItem.SubItems(10) + "," + Me.lv_pedidos.selectedItem.SubItems(11)
          var_cadena = var_cadena + ",'" + Me.lv_pedidos.selectedItem.SubItems(12) + "','" + Me.lv_pedidos.selectedItem.SubItems(13) + "')"
          'MsgBox var_cadena
          rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
      Next var_j
      Set reporte = appl.OpenReport(App.Path + "\rep_oracle_semaforo_pedidos.rpt")
      reporte.RecordSelectionFormula = "{VW_ORACLE_SEMAFORO_PEDIDOS.INTE_TEM_CONSECUTIVO} = '" + CStr(var_consecutivo) + "'"
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      reporte.ExportOptions.FormatType = crEFTExcel80
      reporte.ExportOptions.DestinationType = crEDTDiskFile
      archivo = "c:\reportessid\SEMAFORO_PEDIDOS_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
      reporte.ExportOptions.DiskFileName = archivo
      reporte.Export False
      Set reporte = Nothing
      MsgBox "Se a terminado de guardar el archivo " + archivo
   Else
      MsgBox "La gráfica esta vacia", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command1_Click()
   n = lv_pedidos.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_pedidos.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_pedidos.selectedItem.SubItems(14) = "" And var_rellena = True Then
         If Me.lv_pedidos.selectedItem.SubItems(15) <> "*" Then
         lv_pedidos.selectedItem.SubItems(14) = "*"
         lv_pedidos.ListItems.Item(i).Bold = True
         lv_pedidos.ListItems.Item(i).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(9).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(10).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(11).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(12).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(13).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(14).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(9).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(10).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(11).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(12).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(13).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(14).ForeColor = &HFF0000
         End If
      Else
         If var_encontro = True And lv_pedidos.selectedItem.SubItems(14) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_pedidos.selectedItem.SubItems(14) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
   If Me.lv_pedidos.ListItems.Count > 0 Then
      Me.lv_pedidos.SetFocus
   End If
End Sub

Private Sub cmd_actualizar_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
            Me.lv_pedidos.ListItems.Clear
            Me.txt_bo = 0
            Me.txt_cancelado = 0
            Me.txt_leido = 0
            Me.txt_pedido = 0
            Me.txt_surtido = 0
            Me.txt_surtir = 0
            VAR_HORA_INICIO = CStr(Now)
            var_dia = Day(CDate(Me.txt_inicio))
            var_mes = Month(CDate(Me.txt_inicio))
            var_año = Year(CDate(Me.txt_inicio))
            If Len(CStr(var_dia)) = 1 Then
               var_dia_s = "0" + CStr(var_dia)
            Else
               var_dia_s = CStr(var_dia)
            End If
            If Len(CStr(var_mes)) = 1 Then
               var_mes_s = "0" + CStr(var_mes)
            Else
               var_mes_s = CStr(var_mes)
            End If
            
            If Len(CStr(var_año)) = 2 Then
               var_año_s = "20" + CStr(var_año)
            Else
               var_año_s = CStr(var_año)
            End If
            var_fecha_inicio = var_dia_s + "/" + var_mes_s + "/" + var_año_s
            
            var_dia = Day(CDate(Me.txt_fin) + 1)
            var_mes = Month(CDate(Me.txt_fin) + 1)
            var_año = Year(CDate(Me.txt_fin) + 1)
            If Len(CStr(var_dia)) = 1 Then
               var_dia_s = "0" + CStr(var_dia)
            Else
               var_dia_s = CStr(var_dia)
            End If
            If Len(CStr(var_mes)) = 1 Then
               var_mes_s = "0" + CStr(var_mes)
            Else
               var_mes_s = CStr(var_mes)
            End If
            
            If Len(CStr(var_año)) = 2 Then
               var_año_s = "20" + CStr(var_año)
            Else
               var_año_s = CStr(var_año)
            End If
            var_fecha_fin = var_dia_s + "/" + var_mes_s + "/" + var_año_s
            
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = AMERICAN", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rs.Open "alter session set nls_date_format='dd-mm-yyyy hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            
            var_cadena = "SELECT c.SOURCE_HEADER_ID,   c.SOURCE_HEADER_TYPE_ID,   c.SOURCE_HEADER_TYPE_NAME,   c.customer_id,   ORDERED_DATE,   ORDER_NUMBER,   A.ORDER_TYPE_ID,   RELEASED_STATUS,   a.source_document_id,   d.razon_social_cliente AS cliente,   e.razon_social_cliente AS establecimiento,   F.CREATION_DATE        AS FECHA_DESPACHO,   e.site_use_id,   (SELECT x.ATTRIBUTE1   FROM po_requisition_headers_ALL x,     MTL_SECONDARY_INVENTORIES Y   Where requisition_header_id = A.source_document_id   AND secondary_inventory_name = x.ATTRIBUTE1   ) AS almacen,   (SELECT y.description   FROM po_requisition_headers_ALL x,     MTL_SECONDARY_INVENTORIES Y   Where requisition_header_id = A.source_document_id   AND secondary_inventory_name = x.ATTRIBUTE1   )                                                                                                                                        AS nombre_almacen,"
            var_cadena = var_cadena + " SUM(DECODE(NVL(C.SHIPPED_QUANTITY,0),0,C.REQUESTED_QUANTITY+NVL(C.CANCELLED_QUANTITY,0),C.SHIPPED_QUANTITY+NVL(C.CANCELLED_QUANTITY,0))) AS CANTIDAD_PEDIDA,   SUM(NVL(C.SHIPPED_QUANTITY,0))                                                                                                           AS CANTIDAD_SURTIDA,   SUM(NVL(c.CANCELLED_QUANTITY, 0))                                                                                                        As CANTIDAD_CANCELADA FROM OE_ORDER_HEADERS_ALL A,   OE_ORDER_LINES_ALL B,   WSH_DELIVERABLES_V C,   XXVIA_VW_CLIENTES_BCP d,   XXVIA_VW_CLIENTES_BCP e,   WSH_DLVB_DLVY_V f WHERE A.ORDERED_DATE    >= TO_DATE(?,'DD/MM/YYYY') AND A.ORDERED_DATE       < TO_dATE(?,'DD/MM/YYYY') AND A.SHIP_FROM_ORG_ID   = 93 AND A.ORDER_TYPE_ID     IN (1106,1042,1002) AND A.HEADER_ID          = B.HEADER_ID AND A.HEADER_ID          = SOURCE_HEADER_ID AND B.LINE_ID            = SOURCE_LINE_ID "
            var_cadena = var_cadena + " AND a.invoice_to_org_id  = d.site_use_id AND a.ship_to_org_id     = e.site_use_id AND C.DELIVERY_DETAIL_ID = F.DELIVERY_DETAIL_ID(+)  GROUP BY c.SOURCE_HEADER_ID,   c.SOURCE_HEADER_TYPE_ID,   c.SOURCE_HEADER_TYPE_NAME,   c.customer_id,   ORDERED_DATE,   ORDER_NUMBER,   A.ORDER_TYPE_ID,   RELEASED_STATUS,   a.source_document_id,   d.razon_social_cliente,   e.razon_social_cliente,   e.site_use_id,   f.CREATION_DATE ORDER BY TO_NUMBER(ORDER_NUMBER) "

            
            
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = var_cadena
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_fecha_inicio)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_fecha_fin)
                 .Parameters.Append parametro
            End With
            Set rs = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If Not rs.EOF Then
               cnn.BeginTrans
               rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_SEMAFORO_PEDIDOS", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
               Else
                  var_consecutivo = 0
               End If
               rsaux.Close
               var_consecutivo = var_consecutivo + 1
               rsaux.Open "INSERT INTO TB_TEMP_ORACLE_SEMAFORO_PEDIDOS (INTE_tEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               While Not rs.EOF
                     rsaux.Open "SELECT * FROM TB_TEMP_ORACLE_SEMAFORO_PEDIDOS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rs!order_number), cnn, adOpenDynamic, adLockOptimistic
                     If rsaux.EOF Then

                        var_fecha_fin = CStr(rs!ORDERED_DATE)
                        VAR_FECHA_DESPACHO = CStr(IIf(IsNull(rs!FECHA_DESPACHO), "", rs!FECHA_DESPACHO))
                        
                        
                        'VAR_vendedor = rs!vendedor
                        VAR_CLIENTE = rs!Cliente
                        VAR_ESTABLECIMIENTO = rs!ESTABLECIMIENTO
                        If VAR_CLIENTE = "VIANNEY TEXTIL HOGAR SA DE CV" Then
                           VAR_ESTABLECIMIENTO = rs!nombre_almacen
                           VAR_vendedor = "INTERCOMPAÑIA"
                           VAR_SITE_ID = rs!ALMACEN
                        Else
                           VAR_SITE_ID = rs!site_use_id
                        End If
                        var_estatus = rs!released_status
                        If var_estatus = "R" Then
                           VAR_CANTIDAD_R = rs!CANTIDAD_PEDIDA
                           VAR_CANTIDAD_Y = 0
                           VAR_CANTIDAD_D = 0
                           VAR_CANTIDAD_C = 0
                           VAR_CANTIDAD_B = 0
                        End If
                        If var_estatus = "Y" Then
                           VAR_CANTIDAD_R = 0
                           VAR_CANTIDAD_Y = rs!CANTIDAD_PEDIDA
                           VAR_CANTIDAD_D = 0
                           VAR_CANTIDAD_C = 0
                           VAR_CANTIDAD_B = 0
                        End If
                        If var_estatus = "D" Then
                           VAR_CANTIDAD_R = 0
                           VAR_CANTIDAD_Y = 0
                           VAR_CANTIDAD_D = rs!CANTIDAD_CANCELADA
                           VAR_CANTIDAD_C = 0
                           VAR_CANTIDAD_B = 0
                        End If
                        If var_estatus = "C" Then
                           VAR_CANTIDAD_R = 0
                           VAR_CANTIDAD_Y = 0
                           VAR_CANTIDAD_D = 0
                           VAR_CANTIDAD_C = rs!CANTIDAD_SURTIDA
                           VAR_CANTIDAD_B = 0
                        End If
                        If var_estatus = "B" Then
                           VAR_CANTIDAD_R = 0
                           VAR_CANTIDAD_Y = 0
                           VAR_CANTIDAD_D = 0
                           VAR_CANTIDAD_C = 0
                           VAR_CANTIDAD_B = rs!CANTIDAD_PEDIDA
                        End If
                        'MsgBox "INSERT INTO TB_TEMP_ORACLE_SEMAFORO_PEDIDOS (INTE_TEM_CONSECUTIVO, PEDIDO, FECHA, VENDEDOR, CLIENTE, ESTABLECIMIENTO, R, Y, C, D, B) VALUES (" + CStr(var_consecutivo) + ",'" + CStr(rs!ORDER_NUMBER) + "'," + var_fecha_fin + ",'" + VAR_VENDEDOR + "','" + VAR_CLIENTE + "','" + VAR_ESTABLECIMIENTO + "'," + CStr(VAR_CANTIDAD_R) + ", " + CStr(VAR_CANTIDAD_Y) + ", " + CStr(VAR_CANTIDAD_C) + ", " + CStr(VAR_CANTIDAD_D) + ", " + CStr(VAR_CANTIDAD_B) + ")"
                        If rs!SOURCE_HEADER_TYPE_ID = 0 Then
                           VAR_REGLA = 2
                           VAR_NOMBRE_REGLA = "VTH_CEDI"
                        End If
                        If rs!SOURCE_HEADER_TYPE_ID = 0 Then
                           VAR_REGLA = 2
                           VAR_NOMBRE_REGLA = ""
                        End If
                        If rs!SOURCE_HEADER_TYPE_ID = 0 Then
                           VAR_REGLA = 2
                           VAR_NOMBRE_REGLA = ""
                        End If
                        VAR_REGLA = 2
                        VAR_NOMBRE_REGLA = "VTH_CEDI"
                        rsaux1.Open "INSERT INTO TB_TEMP_ORACLE_SEMAFORO_PEDIDOS (INTE_TEM_CONSECUTIVO, PEDIDO, fecha, VENDEDOR, CLIENTE, ESTABLECIMIENTO, R, Y, C, D, B, SITE_ID,FECHA_DESPACHO, HEADER_ID, TYPE_ID, TYPE_NAME, CUSTOMER_ID, RULE_ID, RULE_NAME) VALUES (" + CStr(var_consecutivo) + ",'" + CStr(rs!order_number) + "','" + var_fecha_fin + "','" + VAR_vendedor + "','" + VAR_CLIENTE + "','" + IIf(IsNull(VAR_ESTABLECIMIENTO), "", VAR_ESTABLECIMIENTO) + "'," + CStr(VAR_CANTIDAD_R) + ", " + CStr(VAR_CANTIDAD_Y) + ", " + CStr(VAR_CANTIDAD_C) + ", " + CStr(VAR_CANTIDAD_D) + ", " + CStr(VAR_CANTIDAD_B) + ",'" + CStr(VAR_SITE_ID) + "','" + VAR_FECHA_DESPACHO + "','" + CStr(rs!SOURCE_HEADER_ID) + "', '" + CStr(rs!SOURCE_HEADER_TYPE_ID) + "', '" + rs!source_header_type_name + "', '" + CStr(rs!CUSTOMER_ID) + "','" + CStr(VAR_REGLA) + "','" + VAR_NOMBRE_REGLA + "')", cnn, adOpenDynamic, adLockOptimistic
                     Else
                        VAR_FECHA_DESPACHO = CStr(IIf(IsNull(rs!FECHA_DESPACHO), "", rs!FECHA_DESPACHO))
                        var_estatus = rs!released_status
                        If var_estatus = "R" Then
                           VAR_CANTIDAD_R = rs!CANTIDAD_PEDIDA
                           rsaux1.Open "UPDATE TB_TEMP_ORACLE_SEMAFORO_PEDIDOS SET R = " + CStr(VAR_CANTIDAD_R) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rs!order_number), cnn, adOpenDynamic, adLockOptimistic
                        End If
                        If var_estatus = "Y" Then
                           VAR_CANTIDAD_Y = rs!CANTIDAD_PEDIDA
                           rsaux1.Open "UPDATE TB_TEMP_ORACLE_SEMAFORO_PEDIDOS SET Y = " + CStr(VAR_CANTIDAD_Y) + ", FECHA_DESPACHO = '" + VAR_FECHA_DESPACHO + "' WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rs!order_number), cnn, adOpenDynamic, adLockOptimistic
                        End If
                        If var_estatus = "D" Then
                           VAR_CANTIDAD_D = rs!CANTIDAD_CANCELADA
                           rsaux1.Open "UPDATE TB_TEMP_ORACLE_SEMAFORO_PEDIDOS SET D = " + CStr(VAR_CANTIDAD_D) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rs!order_number), cnn, adOpenDynamic, adLockOptimistic
                        End If
                        If var_estatus = "C" Then
                           VAR_CANTIDAD_C = rs!CANTIDAD_SURTIDA
                           rsaux1.Open "UPDATE TB_TEMP_ORACLE_SEMAFORO_PEDIDOS SET C = " + CStr(VAR_CANTIDAD_C) + ", FECHA_DESPACHO = '" + VAR_FECHA_DESPACHO + "' WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rs!order_number), cnn, adOpenDynamic, adLockOptimistic
                        End If
                        If var_estatus = "B" Then
                           VAR_CANTIDAD_B = rs!CANTIDAD_PEDIDA
                           rsaux1.Open "UPDATE TB_TEMP_ORACLE_SEMAFORO_PEDIDOS SET B = " + CStr(VAR_CANTIDAD_B) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rs!order_number), cnn, adOpenDynamic, adLockOptimistic
                        End If
                     End If
                     rsaux.Close
                     rs.MoveNext
               Wend
               rsaux.Open "SELECT * FROM TB_TEMP_ORACLE_SEMAFORO_pedidos where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido is not null order by pedido", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux.EOF
                     Set list_item = Me.lv_pedidos.ListItems.Add(, , rsaux!PEDIDO)
                     
                     list_item.SubItems(1) = IIf(IsNull(rsaux!Fecha), "", rsaux!Fecha)
                     list_item.SubItems(2) = IIf(IsNull(rsaux!FECHA_DESPACHO), "", rsaux!FECHA_DESPACHO)
                     rsaux1.Open "select NOMBRE_RUTA, ESTABLECIMIENTO from XXVIA_TB_RUTAS_DISTRIBUCION a, XXVIA_VW_CLIENTES_RUTAS_DISTR b where a.ruta = b.RUTA AND ESTABLECIMIENTO = '" + rsaux!site_id + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux1.EOF Then
                        VAR_vendedor = IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta)
                     Else
                        VAR_vendedor = ""
                     End If
                     rsaux1.Close
                     list_item.SubItems(3) = VAR_vendedor
                     list_item.SubItems(4) = rsaux!Cliente
                     list_item.SubItems(5) = rsaux!ESTABLECIMIENTO
                     list_item.SubItems(6) = rsaux!r
                     list_item.SubItems(7) = rsaux!Y
                     list_item.SubItems(8) = rsaux!c
                     list_item.SubItems(18) = rsaux!header_id
                     list_item.SubItems(19) = rsaux!type_id
                     list_item.SubItems(20) = rsaux!type_name
                     list_item.SubItems(21) = rsaux!CUSTOMER_ID
                     list_item.SubItems(22) = rsaux!rule_id
                     list_item.SubItems(23) = rsaux!rule_name
                     list_item.SubItems(24) = rsaux!site_id
                     x = 1
                     If x = 1 Then
                        var_cantidad_total = rsaux!Y + rsaux!c + rsaux!D + rsaux!B
                        If var_cantidad_total > rsaux!r Then
                           list_item.SubItems(6) = var_cantidad_total
                        End If
                        rsaux3.Open "select * from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where pedido = " + CStr(rsaux!PEDIDO), cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           list_item.SubItems(12) = rsaux3!Embarque
                           list_item.SubItems(13) = "A"
                           list_item.SubItems(12) = IIf(IsNull(rsaux3!Embarque), "", rsaux3!Embarque)
                           list_item.SubItems(17) = Format(IIf(IsNull(rsaux3!VOLUMEN), "0", rsaux3!VOLUMEN), "###,##0.00000")
                        End If
                        rsaux3.Close
                     
                        var_cadena = "select  embarque, source_header_number, char_emb_estatus estatus, sum(floa_sal_cantidad_leida) as cantidad_leida from xxvia_tb_salidas_cajas a, xxvia_Tb_encabezado_embarques b where  a.inte_emb_embarque = b.embarque  and source_header_number = ? group by embarque, source_header_number, char_emb_estatus"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = var_cadena
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux!PEDIDO))
                             .Parameters.Append parametro
                        End With
                        Set rsaux2 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If Not rsaux2.EOF Then
                           list_item.SubItems(9) = rsaux2!cantidad_leida
                           list_item.SubItems(12) = rsaux2!Embarque
                           list_item.SubItems(13) = IIf(IsNull(rsaux2!estatus), "", rsaux2!estatus)
                        Else
                           list_item.SubItems(9) = 0
                           list_item.SubItems(12) = ""
                           list_item.SubItems(13) = ""
                        End If
                        If rsaux!Y = 0 And rsaux!c > 0 Then
                           var_cadena = "select sum(cantidad) from xxvia_tb_negado_distribucion where source_header_number = ?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = var_cadena
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux!PEDIDO))
                                .Parameters.Append parametro
                           End With
                           Set rsaux3 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux3.EOF Then
                              list_item.SubItems(7) = list_item.SubItems(8) + IIf(IsNull(rsaux3(0).Value), 0, rsaux3(0))
                              list_item.SubItems(16) = IIf(IsNull(rsaux3(0).Value), 0, rsaux3(0))
                           Else
                              list_item.SubItems(7) = list_item.SubItems(8)
                              list_item.SubItems(16) = 0
                           End If
                           rsaux3.Close
                        End If
                        list_item.SubItems(10) = rsaux!B
                        list_item.SubItems(11) = rsaux!D
                     
                        rsaux2.Close
                     End If
                     Me.Refresh
                     rsaux.MoveNext
               Wend
               rsaux.Close
               rsaux.Open "delete from tb_temp_oracle_semaforo_pedidos where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               var_pedido = 0
               var_surtir = 0
               var_surtido = 0
               var_leido = 0
               var_bo = 0
               var_cancelado = 0
               For var_j = 1 To Me.lv_pedidos.ListItems.Count
                   Me.lv_pedidos.ListItems.Item(var_j).Selected = True
                   If Trim(Me.lv_pedidos.selectedItem.SubItems(13)) <> "" Then
                      Me.lv_pedidos.selectedItem.SubItems(15) = "*"
                   End If
                   'MsgBox Me.lv_pedidos.selectedItem
                   If CDbl(Trim(Me.lv_pedidos.selectedItem.SubItems(11))) <> 0 Then
                      Me.lv_pedidos.selectedItem.SubItems(15) = "*"
                   End If
                   If CDbl(Trim(Me.lv_pedidos.selectedItem.SubItems(8))) <> 0 Then
                      Me.lv_pedidos.selectedItem.SubItems(15) = "*"
                   End If
                   If CDbl(Trim(Me.lv_pedidos.selectedItem.SubItems(10))) <> 0 Then
                      Me.lv_pedidos.selectedItem.SubItems(15) = "*"
                   End If
                   var_pedido = var_pedido + CDbl(Me.lv_pedidos.selectedItem.SubItems(6))
                   var_surtir = var_surtir + CDbl(Me.lv_pedidos.selectedItem.SubItems(7))
                   var_surtido = var_surtido + CDbl(Me.lv_pedidos.selectedItem.SubItems(8))
                   var_leido = var_leido + CDbl(Me.lv_pedidos.selectedItem.SubItems(9))
                   var_bo = var_bo + CDbl(Me.lv_pedidos.selectedItem.SubItems(10))
                   var_cancelado = var_cancelado + CDbl(Me.lv_pedidos.selectedItem.SubItems(11))
                   If Me.lv_pedidos.selectedItem.SubItems(16) = "" Then
                      var_negado = 0
                   Else
                      var_negado = CDbl(Me.lv_pedidos.selectedItem.SubItems(16))
                   End If
                   If CDbl(Me.lv_pedidos.selectedItem.SubItems(7)) > 0 Then
                      If CDbl(Me.lv_pedidos.selectedItem.SubItems(7)) = (CDbl(Me.lv_pedidos.selectedItem.SubItems(9)) + var_negado) Then
                         lv_pedidos.ListItems.Item(var_j).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ForeColor = &HC000&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(1).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(2).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(3).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(4).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(5).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(6).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(7).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(8).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(9).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(10).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(11).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(12).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(13).Bold = False
                         'lv_pedidos.ListItems.Item(var_j).ListSubItems(14).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(1).ForeColor = &HC000&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(2).ForeColor = &HC000&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(3).ForeColor = &HC000&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(4).ForeColor = &HC000&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(5).ForeColor = &HC000&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(6).ForeColor = &HC000&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(7).ForeColor = &HC000&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(8).ForeColor = &HC000&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(9).ForeColor = &HC000&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(10).ForeColor = &HC000&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(11).ForeColor = &HC000&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(12).ForeColor = &HC000&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(13).ForeColor = &HC000&
                         'lv_pedidos.ListItems.Item(var_j).ListSubItems(14).ForeColor = &HC000&
                      End If
                   End If
                   If Me.lv_pedidos.selectedItem.SubItems(16) = "" Then
                      var_negado = 0
                   Else
                      var_negado = CDbl(Me.lv_pedidos.selectedItem.SubItems(16))
                   End If
                   If CDbl(Me.lv_pedidos.selectedItem.SubItems(9)) > 0 Then
                      If CDbl(Me.lv_pedidos.selectedItem.SubItems(7)) > CDbl(Me.lv_pedidos.selectedItem.SubItems(9)) + var_negado Then
                         lv_pedidos.ListItems.Item(var_j).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ForeColor = &HFF0000
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(1).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(2).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(3).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(4).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(5).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(6).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(7).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(8).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(9).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(10).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(11).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(12).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(13).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(1).ForeColor = &HFF0000
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(2).ForeColor = &HFF0000
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(3).ForeColor = &HFF0000
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(4).ForeColor = &HFF0000
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(5).ForeColor = &HFF0000
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(6).ForeColor = &HFF0000
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(7).ForeColor = &HFF0000
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(8).ForeColor = &HFF0000
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(9).ForeColor = &HFF0000
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(10).ForeColor = &HFF0000
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(11).ForeColor = &HFF0000
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(12).ForeColor = &HFF0000
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(13).ForeColor = &HFF0000
                      End If
                   End If
                   If CDbl(Me.lv_pedidos.selectedItem.SubItems(7)) > 0 Then
                      If CDbl(Me.lv_pedidos.selectedItem.SubItems(8)) = 0 And CDbl(Me.lv_pedidos.selectedItem.SubItems(9)) = 0 Then
                         lv_pedidos.ListItems.Item(var_j).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ForeColor = &HC0&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(1).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(2).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(3).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(4).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(5).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(6).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(7).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(8).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(9).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(10).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(11).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(12).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(13).Bold = False
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(1).ForeColor = &HC0&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(2).ForeColor = &HC0&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(3).ForeColor = &HC0&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(4).ForeColor = &HC0&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(5).ForeColor = &HC0&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(6).ForeColor = &HC0&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(7).ForeColor = &HC0&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(8).ForeColor = &HC0&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(9).ForeColor = &HC0&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(10).ForeColor = &HC0&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(11).ForeColor = &HC0&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(12).ForeColor = &HC0&
                         lv_pedidos.ListItems.Item(var_j).ListSubItems(13).ForeColor = &HC0&
                      End If
                   End If
                   If Me.lv_pedidos.selectedItem.SubItems(17) = "" Then
                      If CDbl(Me.lv_pedidos.selectedItem.SubItems(7)) > 0 Then
                         var_cadena = "SELECT SUM(SRC_REQUESTED_QUANTITY * UNIT_VOLUME) FROM WSH_DELIVERABLES_V WHERE source_header_number = ? AND RELEASED_STATUS = 'Y'"
                         With comandoORA
                              .ActiveConnection = cnnoracle_4
                              .CommandType = adCmdText
                              .CommandText = var_cadena
                              Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_pedidos.selectedItem))
                              .Parameters.Append parametro
                         End With
                         Set rsaux3 = comandoORA.execute
                         Set comandoORA = Nothing
                         Set parametro = Nothing
                         If Not rsaux3.EOF Then
                            Me.lv_pedidos.selectedItem.SubItems(17) = Format(IIf(IsNull(rsaux3(0).Value), 0, rsaux3(0)), "###,##0.00000")
                         Else
                            Me.lv_pedidos.selectedItem.SubItems(17) = "0.00000"
                         End If
                         rsaux3.Close
                      End If
                   End If
               
               Next var_j
               Me.txt_pedido = Format(var_pedido, "###,###,###.00")
               Me.txt_surtir = Format(var_surtir, "###,###,###.00")
               Me.txt_surtido = Format(var_surtido, "###,###,###.00")
               Me.txt_leido = Format(var_leido, "###,###,###.00")
               Me.txt_bo = Format(var_bo, "###,###,###.00")
               Me.txt_cancelado = Format(var_cancelado, "###,###,###.00")
            Else
               MsgBox "No existe respuesta para el periodo seleccionado", vbOKOnly, "ATENCION"
            End If
            rs.Close
            VAR_HORA_FIN = CStr(Now)
            'MsgBox VAR_HORA_INICIO + " " + VAR_HORA_FIN
            
            MsgBox "Termino la carga", vbOKOnly, "ATENCION"
         Else
            MsgBox "La fecha final no puede ser menor a la fecha de inicio", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command2_Click()
   i = lv_pedidos.selectedItem.Index
   If Me.lv_pedidos.selectedItem.SubItems(15) <> "*" Then
      If lv_pedidos.selectedItem.SubItems(14) = "*" Then
         lv_pedidos.selectedItem.SubItems(14) = ""
         lv_pedidos.ListItems.Item(i).Bold = False
         lv_pedidos.ListItems.Item(i).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(9).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(10).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(11).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(12).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(13).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(14).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(9).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(10).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(11).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(12).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(13).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(14).ForeColor = &H80000012
         lv_pedidos.Refresh
      Else
         lv_pedidos.selectedItem.SubItems(14) = "*"
         lv_pedidos.ListItems.Item(i).Bold = True
         lv_pedidos.ListItems.Item(i).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(9).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(10).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(11).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(12).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(13).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(14).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(9).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(10).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(11).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(12).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(13).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(14).ForeColor = &HFF0000
         lv_pedidos.Refresh
      End If
   Else
         MsgBox "El pedido no puede ser seleccionado para despacharlo", vbOKOnly, "ATENCION"
   End If
   If Me.lv_pedidos.ListItems.Count > 0 Then
      Me.lv_pedidos.SetFocus
   End If
End Sub

Private Sub Command3_Click()
   n = lv_pedidos.ListItems.Count
   For i = 1 To n
      lv_pedidos.ListItems.Item(i).Selected = True
      If lv_pedidos.selectedItem.SubItems(14) = "*" Then
         lv_pedidos.selectedItem.SubItems(14) = ""
         lv_pedidos.ListItems.Item(i).Bold = False
         lv_pedidos.ListItems.Item(i).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(9).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(10).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(11).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(12).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(13).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(14).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(9).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(10).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(11).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(12).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(13).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(14).ForeColor = &H80000012
      Else
         lv_pedidos.selectedItem.SubItems(14) = "*"
         lv_pedidos.ListItems.Item(i).Bold = True
         lv_pedidos.ListItems.Item(i).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(9).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(10).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(11).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(12).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(13).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(14).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(9).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(10).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(11).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(12).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(13).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(14).ForeColor = &HFF0000
      End If
   Next i
   If Me.lv_pedidos.ListItems.Count > 0 Then
      Me.lv_pedidos.SetFocus
   End If
End Sub

Private Sub Command4_Click()
   n = lv_pedidos.ListItems.Count
   For i = 1 To n
      lv_pedidos.ListItems.Item(i).Selected = True
      lv_pedidos.selectedItem.SubItems(14) = ""
      lv_pedidos.ListItems.Item(i).Bold = False
      lv_pedidos.ListItems.Item(i).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(9).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(10).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(11).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(12).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(13).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(14).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(9).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(10).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(11).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(12).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(13).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(14).ForeColor = &H80000012
   Next i
   lv_pedidos.Refresh
   If Me.lv_pedidos.ListItems.Count > 0 Then
      Me.lv_pedidos.SetFocus
   End If
End Sub

Private Sub Command5_Click()
   n = lv_pedidos.ListItems.Count
   For i = 1 To n
      lv_pedidos.ListItems.Item(i).Selected = True
      If Me.lv_pedidos.selectedItem.SubItems(15) <> "*" Then
         lv_pedidos.selectedItem.SubItems(14) = "*"
         lv_pedidos.ListItems.Item(i).Bold = True
         lv_pedidos.ListItems.Item(i).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(9).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(10).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(11).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(12).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(13).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(14).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(9).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(10).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(11).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(12).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(13).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(14).ForeColor = &HFF0000
      End If
   Next i
   lv_pedidos.Refresh
   If Me.lv_pedidos.ListItems.Count > 0 Then
      Me.lv_pedidos.SetFocus
   End If
End Sub

Private Sub Command6_Click()
   var_posible = 0
   For var_j = 1 To Me.lv_pedidos.ListItems.Count
       Me.lv_pedidos.ListItems.Item(var_j).Selected = True
       If Me.lv_pedidos.selectedItem.SubItems(14) = "*" Or Me.lv_pedidos.selectedItem.SubItems(14) = "" Then
          On Error GoTo salir2:
          rsaux1.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
          var_cadena = "call xxvia_sp_llam_despachar_ped (?, ?, ?, ?, ?, ?, ?, ?, ?)"
          With comandoORA
               .ActiveConnection = cnnoracle_4
               .CommandType = adCmdText
               .CommandText = var_cadena
               Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_pedidos.selectedItem.SubItems(21)))
               .Parameters.Append parametro
               Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_pedidos.selectedItem.SubItems(18)))
               .Parameters.Append parametro
               Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, 92)
               .Parameters.Append parametro
               Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, 93)
               .Parameters.Append parametro
               Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, 2)
               .Parameters.Append parametro
               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "VTH_CEDI")
               .Parameters.Append parametro
               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "E")
               .Parameters.Append parametro
               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Me.lv_pedidos.selectedItem.SubItems(20))
               .Parameters.Append parametro
               Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_pedidos.selectedItem.SubItems(19)))
               .Parameters.Append parametro
               End With
          Set rsaux7 = comandoORA.execute
          Set comandoORA = Nothing
          Set parametro = Nothing
          
          
          rsaux1.Open "CALL xxvia_sp_llam_despachar_ped(" + Me.lv_pedidos.selectedItem.SubItems(21) + "," + Me.lv_pedidos.selectedItem.SubItems(18) + ",92,93,2,'VTH_CEDI','E','" + Me.lv_pedidos.selectedItem.SubItems(20) + "'," + Me.lv_pedidos.selectedItem.SubItems(19) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
'          objConn.Open var_conexion_oracle
'          With objCmd
'               objConn.BeginTrans
'               .ActiveConnection = cnnoracle_4
'               .CommandText = "xxvia_sp_despachar_pedido"
'               .CommandType = adCmdStoredProc
'
'               Set objParm = .CreateParameter("p_custumer_id", adNumeric, adParamInput, 50, CDbl(Me.lv_pedidos.selectedItem.SubItems(21)))
'               .Parameters.Append objParm
'
'               Set objParm = .CreateParameter("p_header_id", adNumeric, adParamInput, 100, CDbl(Me.lv_pedidos.selectedItem.SubItems(18)))
'               .Parameters.Append objParm
'
'               Set objParm = .CreateParameter("p_uo_id", adNumeric, adParamInput, 100, 92)
'               .Parameters.Append objParm
'
'               Set objParm = .CreateParameter("p_org_id", adNumeric, adParamInput, 100, 93)
'               .Parameters.Append objParm
'
'               Set objParm = .CreateParameter("p_rule_id", adNumeric, adParamInput, 100, 2)
'               .Parameters.Append objParm
'
'               Set objParm = .CreateParameter("p_rule_name", adVarChar, adParamInput, 100, "VTH_CEDI")
'               .Parameters.Append objParm
'
'               Set objParm = .CreateParameter("P_PEDIDO_DESPACHADO", adVarChar, adParamInput, 100, "E")
'               .Parameters.Append objParm
'
'               Set objParm = .CreateParameter("p_tipo_pedido", adVarChar, adParamInput, 100, "Me.lv_pedidos.selectedItem.SubItems(20)")
'               .Parameters.Append objParm
'
'               Set objParm = .CreateParameter("p_tipo_pedido_id", adNumeric, adParamInput, 100, CDbl(Me.lv_pedidos.selectedItem.SubItems(19)))
'               .Parameters.Append objParm
'
'               Set objParm = .CreateParameter("p_msj", adVarChar, adParamOutput, 50, "")
'               .Parameters.Append objParm
'
'               rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
'               rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
'               On Error GoTo SALIR:
'               .execute
'
'               VAR_X_TRIP_ID = .Parameters("p_msj").Value
'               objConn.CommitTrans
'          End With
'          Set objConn = Nothing
'          Set objCmd = Nothing
       End If
   Next
   If var_posible = 1 Then
      MsgBox "PROXIMAMENTE", vbOKOnly, "ATENCION"
   Else
      MsgBox "No se han seleccionado pedidos a despachar", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir2:
   'MsgBox Err.Description
   If Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux12.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux12.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If var_contador_errores < 6 Then
         var_contador_errores = var_contador_errores + 1
         MsgBox Err.Description
         Resume
      Else
         var_contador_errores = 0
         Resume Next
      End If
   Else
      MsgBox Err.Description
   End If
   
   
   
SALIR:
   If Err.Number = -2147467259 Or Err.Number = -2147217900 Then
      
      MsgBox Err.Description
      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic

      Resume
   End If
End Sub

Private Sub Command7_Click()
   rs.Open "DELETE FROM TB_TEMP_ORACLE_CLIENTES_SIN_RUTA", cnn, adOpenDynamic, adLockOptimistic
   For var_j = 1 To Me.lv_pedidos.ListItems.Count
       Me.lv_pedidos.ListItems.Item(var_j).Selected = True
       rs.Open "INSERT INTO TB_TEMP_ORACLE_CLIENTES_SIN_RUTA (RUTA, CLIENTE, ESTABLECIMIENTO, CLAVE_CLIENTE, SITE_USE_ID) VALUES ('" + Me.lv_pedidos.selectedItem.SubItems(3) + "','" + Me.lv_pedidos.selectedItem.SubItems(4) + "','" + Me.lv_pedidos.selectedItem.SubItems(5) + "','" + Me.lv_pedidos.selectedItem.SubItems(21) + "','" + Me.lv_pedidos.selectedItem.SubItems(24) + "')", cnn, adOpenDynamic, adLockOptimistic
   Next var_j
End Sub

Private Sub Form_Load()
   Me.txt_inicio = Date
   Me.txt_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_pedidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_pedidos, ColumnHeader)
End Sub

Private Sub lv_pedidos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_pedidos.selectedItem.Index
      If lv_pedidos.selectedItem.SubItems(15) <> "*" Then
         If lv_pedidos.selectedItem.SubItems(14) = "*" Then
            lv_pedidos.selectedItem.SubItems(14) = ""
            lv_pedidos.ListItems.Item(i).Bold = False
            lv_pedidos.ListItems.Item(i).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(9).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(10).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(11).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(12).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(13).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(14).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(9).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(10).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(11).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(12).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(13).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(14).ForeColor = &H80000012
            lv_pedidos.Refresh
         Else
            lv_pedidos.selectedItem.SubItems(14) = "*"
            lv_pedidos.ListItems.Item(i).Bold = True
            lv_pedidos.ListItems.Item(i).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(9).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(10).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(11).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(12).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(13).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(14).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(9).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(10).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(11).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(12).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(13).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(14).ForeColor = &HFF0000
            lv_pedidos.Refresh
         End If
         If Me.lv_pedidos.ListItems.Count > 0 Then
            Me.lv_pedidos.SetFocus
         End If
      Else
         MsgBox "El pedido no puede ser seleccionado para despacharlo", vbOKOnly, "ATENCION"
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
