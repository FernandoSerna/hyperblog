VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_asignar_pedidos_embarque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de pedidos a embarques"
   ClientHeight    =   11115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Height          =   390
      Left            =   7545
      Picture         =   "frmoracle_asignar_pedidos_embarque.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5595
      Width           =   450
   End
   Begin VB.CommandButton cmd_pasar 
      Height          =   390
      Left            =   6975
      Picture         =   "frmoracle_asignar_pedidos_embarque.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5595
      Width           =   450
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   14760
      Picture         =   "frmoracle_asignar_pedidos_embarque.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmoracle_asignar_pedidos_embarque.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Asignar"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   120
      TabIndex        =   19
      Top             =   270
      Width           =   15015
   End
   Begin VB.TextBox txt_embarque 
      Height          =   510
      Left            =   8235
      TabIndex        =   18
      Top             =   5535
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Frame Frame2 
      Height          =   5085
      Left            =   135
      TabIndex        =   10
      Top             =   5955
      Width           =   14925
      Begin VB.TextBox txt_volumen_unidad 
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
         Left            =   4080
         TabIndex        =   29
         Top             =   4560
         Width           =   1740
      End
      Begin VB.TextBox txt_porcentaje 
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
         Left            =   8865
         TabIndex        =   27
         Top             =   4575
         Width           =   1545
      End
      Begin VB.TextBox txt_total_volumen 
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
         Left            =   13095
         TabIndex        =   26
         Top             =   4590
         Width           =   1740
      End
      Begin VB.CommandButton cmd_todos_2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmoracle_asignar_pedidos_embarque.frx":0940
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   510
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno_2 
         Height          =   315
         Left            =   60
         Picture         =   "frmoracle_asignar_pedidos_embarque.frx":0B56
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   510
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir_2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         Picture         =   "frmoracle_asignar_pedidos_embarque.frx":0C58
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   510
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar_2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         Picture         =   "frmoracle_asignar_pedidos_embarque.frx":0D2A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Marcar (Enter)"
         Top             =   510
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion_2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1380
         Picture         =   "frmoracle_asignar_pedidos_embarque.frx":0F74
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   510
         Visible         =   0   'False
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_asignados 
         Height          =   3510
         Left            =   75
         TabIndex        =   16
         Top             =   870
         Width           =   14760
         _ExtentX        =   26035
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Agente"
            Object.Width           =   8378
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   8378
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Piezas"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Orden"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Volumen"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Establecimiento"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Left            =   10500
         TabIndex        =   31
         Top             =   4635
         Width           =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Volumen unidad:"
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
         Left            =   1995
         TabIndex        =   30
         Top             =   4635
         Width           =   2025
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje ocupación:"
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
         Left            =   6135
         TabIndex        =   28
         Top             =   4650
         Width           =   2685
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total volumen:"
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
         Left            =   11280
         TabIndex        =   25
         Top             =   4665
         Width           =   1755
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   " Pedidos asignados"
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   45
         TabIndex        =   17
         Top             =   135
         Width           =   14835
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5100
      Left            =   135
      TabIndex        =   2
      Top             =   390
      Width           =   14910
      Begin VB.TextBox txt_clave 
         Height          =   360
         Left            =   2475
         TabIndex        =   33
         Top             =   495
         Width           =   735
      End
      Begin VB.TextBox txt_nombre 
         Height          =   360
         Left            =   3240
         TabIndex        =   32
         Top             =   495
         Width           =   5190
      End
      Begin VB.TextBox txt_total_volumen_seleccionado 
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
         Left            =   13065
         TabIndex        =   24
         Top             =   4575
         Width           =   1740
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1380
         Picture         =   "frmoracle_asignar_pedidos_embarque.frx":118A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   510
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         Picture         =   "frmoracle_asignar_pedidos_embarque.frx":13A0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar (Enter)"
         Top             =   510
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         Picture         =   "frmoracle_asignar_pedidos_embarque.frx":15EA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   510
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   60
         Picture         =   "frmoracle_asignar_pedidos_embarque.frx":16BC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   510
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmoracle_asignar_pedidos_embarque.frx":17BE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   510
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_pendientes 
         Height          =   3585
         Left            =   45
         TabIndex        =   3
         Top             =   930
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   6324
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Agente"
            Object.Width           =   8378
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   8467
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Piezas"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Orden"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Volumen"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Establecimiento"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Left            =   2010
         TabIndex        =   34
         Top             =   555
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total volumen:"
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
         Left            =   11250
         TabIndex        =   23
         Top             =   4650
         Width           =   1755
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   " Pedidos pendientes"
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   45
         TabIndex        =   9
         Top             =   135
         Width           =   14820
      End
   End
   Begin VB.Label lbl_anden 
      AutoSize        =   -1  'True
      Caption         =   "Estación:"
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
      Left            =   6555
      TabIndex        =   22
      Top             =   -15
      Width           =   1305
   End
End
Attribute VB_Name = "frmoracle_asignar_pedidos_embarque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter
Dim var_consecutivo_orden As Integer
Private Sub cmd_guardar_Click()
   If Me.lv_asignados.ListItems.Count > 0 Then
      For var_j = 1 To lv_asignados.ListItems.Count
          lv_asignados.ListItems(var_j).Selected = True
          If Me.lv_asignados.selectedItem.SubItems(5) = "" Then
             var_orden = 0
          Else
             var_orden = CDbl(Me.lv_asignados.selectedItem.SubItems(5))
          End If
          If CDbl(Me.lv_asignados.selectedItem) = -5 Then
             rsaux.Open "select * from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(var_embarque_asignar) + " and pedido = '" + CStr(Trim(Me.lv_asignados.selectedItem)) + "'", cnn, adOpenDynamic, adLockOptimistic
             If rsaux.EOF Then
                rsaux1.Open "insert into tb_oracle_pedidos_asignados_embarques (agente, nombre_agente, pedido, cliente, piezas, ORGANIZACION) values ('1016','" + Me.lv_asignados.selectedItem.SubItems(1) + "','" + Me.lv_asignados.selectedItem + "','" + Me.lv_asignados.selectedItem.SubItems(2) + "',0," + CStr(var_unidad_organizacional) + ")", cnn, adOpenDynamic, adLockOptimistic
                rsaux1.Open "insert into tb_oracle_cajas_aduana (embarque, pedido, numero_caja, caja, agente, cliente, establecimiento, piezas, estatus, tipo_empaque, caja_pedido, sello, lote) values  ('" + CStr(var_embarque_asignar) + "','" + Me.lv_asignados.selectedItem + "'," + CStr(Me.lv_asignados.selectedItem) + ",'C" + Trim(Me.lv_asignados.selectedItem) + "','" + Me.lv_asignados.selectedItem.SubItems(1) + "','" + Me.lv_asignados.selectedItem.SubItems(2) + "','',0,'L','TEXTILERA', " + Trim(Me.lv_asignados.selectedItem) + ",'',0) ", cnn, adOpenDynamic, adLockOptimistic
                var_cadena = "INSERT INTO XXVIA_TB_SALIDAS_CAJAS (INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, FLOA_SAL_CANTIDAD_LEIDA, INTE_PAQ_CAJA,  TIPO_cAJA, CAJA_PEDIDO,LOTE, ESTATUS_PEDIDO)"
                var_cadena = var_cadena + " values (" + CStr(var_embarque_asignar) + "," + CStr(CDbl(Me.lv_asignados.selectedItem)) + ",0," + CStr(Me.lv_asignados.selectedItem) + ",'TEXTILERA'," + CStr(Me.lv_asignados.selectedItem) + ", 0, 1) "
                rsaux1.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                
                
                
             End If
             rsaux.Close
          End If
          rsaux.Open "update tb_oracle_pedidos_asignados_embarques set embarque = " + CStr(var_embarque_asignar) + ", dia = " + CStr(Day(Date)) + ", mes = " + CStr(Month(Date)) + ", año = " + CStr(Year(Date)) + ", orden_pedido = " + CStr(var_orden) + ", ESTACION = " + CStr(var_anden_global) + ", VOLUMEN = " + CStr(CDbl(Me.lv_asignados.selectedItem.SubItems(6))) + " where pedido = '" + Me.lv_asignados.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
      Next var_j
      Unload Me
   Else
      MsgBox "No se asigno ningun pedido al embarque", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_invertir_2_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = Me.lv_asignados.ListItems.Count
   For i = 1 To n
       If Me.lv_asignados.ListItems.Item(i).SubItems(4) = "*" Then
          Me.lv_asignados.ListItems.Item(i).SubItems(4) = " "
          Me.lv_asignados.ListItems.Item(i).Bold = False
          Me.lv_asignados.ListItems.Item(i).ForeColor = &H80000012
          Me.lv_asignados.ListItems.Item(i).ListSubItems(1).Bold = False
          Me.lv_asignados.ListItems.Item(i).ListSubItems(2).Bold = False
          Me.lv_asignados.ListItems.Item(i).ListSubItems(3).Bold = False
          Me.lv_asignados.ListItems.Item(i).ListSubItems(4).Bold = False
          Me.lv_asignados.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
          Me.lv_asignados.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
          Me.lv_asignados.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
          Me.lv_asignados.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
       Else
          Me.lv_asignados.ListItems.Item(i).SubItems(4) = "*"
          Me.lv_asignados.ListItems.Item(i).Bold = True
          Me.lv_asignados.ListItems.Item(i).ForeColor = &H8000&
          Me.lv_asignados.ListItems.Item(i).ListSubItems(1).Bold = True
          Me.lv_asignados.ListItems.Item(i).ListSubItems(2).Bold = True
          Me.lv_asignados.ListItems.Item(i).ListSubItems(3).Bold = True
          Me.lv_asignados.ListItems.Item(i).ListSubItems(4).Bold = True
          Me.lv_asignados.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
          Me.lv_asignados.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
          Me.lv_asignados.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
          Me.lv_asignados.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
      End If
   Next
   Me.lv_asignados.Refresh

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
   n = lv_pendientes.ListItems.Count
   For i = 1 To n
       If lv_pendientes.ListItems.Item(i).SubItems(4) = "*" Then
          lv_pendientes.ListItems.Item(i).SubItems(4) = " "
          lv_pendientes.ListItems.Item(i).Bold = False
          lv_pendientes.ListItems.Item(i).ForeColor = &H80000012
          lv_pendientes.ListItems.Item(i).ListSubItems(1).Bold = False
          lv_pendientes.ListItems.Item(i).ListSubItems(2).Bold = False
          lv_pendientes.ListItems.Item(i).ListSubItems(3).Bold = False
          lv_pendientes.ListItems.Item(i).ListSubItems(4).Bold = False
          lv_pendientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
          lv_pendientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
          lv_pendientes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
          lv_pendientes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
       Else
          lv_pendientes.ListItems.Item(i).SubItems(4) = "*"
          lv_pendientes.ListItems.Item(i).Bold = True
          lv_pendientes.ListItems.Item(i).ForeColor = &H8000&
          lv_pendientes.ListItems.Item(i).ListSubItems(1).Bold = True
          lv_pendientes.ListItems.Item(i).ListSubItems(2).Bold = True
          lv_pendientes.ListItems.Item(i).ListSubItems(3).Bold = True
          lv_pendientes.ListItems.Item(i).ListSubItems(4).Bold = True
          lv_pendientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
          lv_pendientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
          lv_pendientes.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
          lv_pendientes.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
      End If
   Next
   lv_pendientes.Refresh
  Me.txt_total_volumen_seleccionado = Format(0, "###,##0.00000")
  For var_j = 1 To Me.lv_pendientes.ListItems.Count
      Me.lv_pendientes.ListItems.Item(var_j).Selected = True
      If Me.lv_pendientes.selectedItem.SubItems(4) = "*" Then
         Me.txt_total_volumen_seleccionado = Format(CDbl(Me.txt_total_volumen_seleccionado) + CDbl(Me.lv_pendientes.selectedItem.SubItems(6)), "###,##0.00000")
      End If
  Next var_j

End Sub

Private Sub cmd_marcar_2_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   i = Me.lv_asignados.selectedItem.Index
   If Me.lv_asignados.selectedItem.SubItems(4) = "*" Then
       Me.lv_asignados.ListItems.Item(i).SubItems(4) = " "
       Me.lv_asignados.ListItems.Item(i).Bold = False
       Me.lv_asignados.ListItems.Item(i).ForeColor = &H80000012
       Me.lv_asignados.ListItems.Item(i).ListSubItems(1).Bold = False
       Me.lv_asignados.ListItems.Item(i).ListSubItems(2).Bold = False
       Me.lv_asignados.ListItems.Item(i).ListSubItems(3).Bold = False
       Me.lv_asignados.ListItems.Item(i).ListSubItems(4).Bold = False
       Me.lv_asignados.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       Me.lv_asignados.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
       Me.lv_asignados.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
       Me.lv_asignados.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
   Else
      Me.lv_asignados.ListItems.Item(i).SubItems(4) = "*"
      Me.lv_asignados.ListItems.Item(i).Bold = True
      Me.lv_asignados.ListItems.Item(i).ForeColor = &H8000&
      Me.lv_asignados.ListItems.Item(i).ListSubItems(1).Bold = True
      Me.lv_asignados.ListItems.Item(i).ListSubItems(2).Bold = True
      Me.lv_asignados.ListItems.Item(i).ListSubItems(3).Bold = True
      Me.lv_asignados.ListItems.Item(i).ListSubItems(4).Bold = True
      Me.lv_asignados.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
      Me.lv_asignados.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
      Me.lv_asignados.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
      Me.lv_asignados.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
  End If
  Me.lv_asignados.Refresh


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
   i = lv_pendientes.selectedItem.Index
   If lv_pendientes.selectedItem.SubItems(4) = "*" Then
       lv_pendientes.ListItems.Item(i).SubItems(4) = " "
       lv_pendientes.ListItems.Item(i).Bold = False
       lv_pendientes.ListItems.Item(i).ForeColor = &H80000012
       lv_pendientes.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_pendientes.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_pendientes.ListItems.Item(i).ListSubItems(3).Bold = False
       lv_pendientes.ListItems.Item(i).ListSubItems(4).Bold = False
       lv_pendientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_pendientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
       lv_pendientes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
       lv_pendientes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
   Else
      lv_pendientes.ListItems.Item(i).SubItems(4) = "*"
      lv_pendientes.ListItems.Item(i).Bold = True
      lv_pendientes.ListItems.Item(i).ForeColor = &H8000&
      lv_pendientes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_pendientes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_pendientes.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_pendientes.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_pendientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
      lv_pendientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
      lv_pendientes.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
      lv_pendientes.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
  End If
  lv_pendientes.Refresh
  Me.txt_total_volumen_seleccionado = Format(0, "###,##0.00000")
  For var_j = 1 To Me.lv_pendientes.ListItems.Count
      Me.lv_pendientes.ListItems.Item(var_j).Selected = True
      If Me.lv_pendientes.selectedItem.SubItems(4) = "*" Then
         Me.txt_total_volumen_seleccionado = Format(CDbl(Me.txt_total_volumen_seleccionado) + CDbl(Me.lv_pendientes.selectedItem.SubItems(6)), "###,##0.00000")
      End If
  Next var_j

End Sub

Private Sub cmd_ninguno_2_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_asignados.ListItems.Count
   For i = 1 To n
       lv_asignados.ListItems.Item(i).SubItems(4) = " "
       lv_asignados.ListItems.Item(i).Bold = False
       lv_asignados.ListItems.Item(i).ForeColor = &H80000012
       lv_asignados.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_asignados.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_asignados.ListItems.Item(i).ListSubItems(3).Bold = False
       lv_asignados.ListItems.Item(i).ListSubItems(4).Bold = False
       lv_asignados.ListItems.Item(i).ListSubItems(6).Bold = False
       lv_asignados.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_asignados.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
       lv_asignados.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
       lv_asignados.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
       lv_asignados.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
    Next
    lv_asignados.Refresh

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
   n = lv_pendientes.ListItems.Count
   For i = 1 To n
       lv_pendientes.ListItems.Item(i).SubItems(4) = " "
       lv_pendientes.ListItems.Item(i).Bold = False
       lv_pendientes.ListItems.Item(i).ForeColor = &H80000012
       lv_pendientes.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_pendientes.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_pendientes.ListItems.Item(i).ListSubItems(3).Bold = False
       lv_pendientes.ListItems.Item(i).ListSubItems(4).Bold = False
       lv_pendientes.ListItems.Item(i).ListSubItems(6).Bold = False
       lv_pendientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_pendientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
       lv_pendientes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
       lv_pendientes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
       lv_pendientes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
       lv_pendientes.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
    Next
    lv_pendientes.Refresh
    Me.txt_total_volumen_seleccionado = "0.00000"
End Sub

Private Sub cmd_pasar_Click()
   Me.txt_total_volumen = "0.00000"
   For var_j = 1 To Me.lv_pendientes.ListItems.Count
       Me.lv_pendientes.ListItems.Item(var_j).Selected = True
       If Me.lv_pendientes.selectedItem.SubItems(4) = "*" Then
          Set list_item = lv_asignados.ListItems.Add(, , Me.lv_pendientes.selectedItem)
          list_item.SubItems(1) = Me.lv_pendientes.selectedItem.SubItems(1)
          list_item.SubItems(2) = Me.lv_pendientes.selectedItem.SubItems(2)
          list_item.SubItems(3) = Me.lv_pendientes.selectedItem.SubItems(3)
          list_item.SubItems(5) = Me.lv_pendientes.selectedItem.SubItems(5)
          list_item.SubItems(6) = Me.lv_pendientes.selectedItem.SubItems(6)
          list_item.SubItems(7) = Me.lv_pendientes.selectedItem.SubItems(7)
       End If
   Next var_j
   For var_j = Me.lv_pendientes.ListItems.Count To 1 Step -1
       Me.lv_pendientes.ListItems.Item(var_j).Selected = True
       If Me.lv_pendientes.selectedItem.SubItems(4) = "*" Then
          lv_pendientes.ListItems.Remove (lv_pendientes.selectedItem.Index)
       End If
   Next var_j
   Me.txt_total_volumen = Format(0, "###,##0.00000")
   For var_j = 1 To Me.lv_asignados.ListItems.Count
       Me.lv_asignados.ListItems.Item(var_j).Selected = True
       Me.txt_total_volumen = Format(CDbl(Me.txt_total_volumen) + CDbl(Me.lv_asignados.selectedItem.SubItems(6)), "###,##0.00000")
   Next var_j
   
   Me.txt_total_volumen_seleccionado.Text = "0.00000"
   Me.lv_pendientes.Refresh
   Me.lv_asignados.Refresh
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_2_Click()
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
   n = Me.lv_asignados.ListItems.Count
   For i = 1 To n
       If Me.lv_asignados.ListItems.Item(i).SubItems(4) = "*" And primera_vez = False Then
          numero_seleccionado1 = i
          primera_vez = True
       End If
       If Me.lv_asignados.ListItems.Item(i).SubItems(4) = "*" And primera_vez = True Then
          numero_seleccionado2 = i
       End If
   Next
   For i = numero_seleccionado1 To numero_seleccionado2
       Me.lv_asignados.ListItems.Item(i).SubItems(4) = "*"
       Me.lv_asignados.ListItems.Item(i).Bold = True
       Me.lv_asignados.ListItems.Item(i).ForeColor = &H8000&
       Me.lv_asignados.ListItems.Item(i).ListSubItems(1).Bold = True
       Me.lv_asignados.ListItems.Item(i).ListSubItems(2).Bold = True
       Me.lv_asignados.ListItems.Item(i).ListSubItems(3).Bold = True
       Me.lv_asignados.ListItems.Item(i).ListSubItems(4).Bold = True
       Me.lv_asignados.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       Me.lv_asignados.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
       Me.lv_asignados.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
       Me.lv_asignados.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
       Me.lv_asignados.ListItems.Item(i).ListSubItems(6).ForeColor = &H8000&
       Me.lv_asignados.Refresh
   Next

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
   n = lv_pendientes.ListItems.Count
   For i = 1 To n
       If lv_pendientes.ListItems.Item(i).SubItems(4) = "*" And primera_vez = False Then
          numero_seleccionado1 = i
          primera_vez = True
       End If
       If lv_pendientes.ListItems.Item(i).SubItems(4) = "*" And primera_vez = True Then
          numero_seleccionado2 = i
       End If
   Next
   For i = numero_seleccionado1 To numero_seleccionado2
       lv_pendientes.ListItems.Item(i).SubItems(4) = "*"
       lv_pendientes.ListItems.Item(i).Bold = True
       lv_pendientes.ListItems.Item(i).ForeColor = &H8000&
       lv_pendientes.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_pendientes.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_pendientes.ListItems.Item(i).ListSubItems(3).Bold = True
       lv_pendientes.ListItems.Item(i).ListSubItems(4).Bold = True
       lv_pendientes.ListItems.Item(i).ListSubItems(6).Bold = True
       lv_pendientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_pendientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
       lv_pendientes.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
       lv_pendientes.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
       lv_pendientes.ListItems.Item(i).ListSubItems(6).ForeColor = &H8000&
       lv_pendientes.Refresh
   Next
  Me.txt_total_volumen_seleccionado = Format(0, "###,##0.00000")
  For var_j = 1 To Me.lv_pendientes.ListItems.Count
      Me.lv_pendientes.ListItems.Item(var_j).Selected = True
      If Me.lv_pendientes.selectedItem.SubItems(4) = "*" Then
         Me.txt_total_volumen_seleccionado = Format(CDbl(Me.txt_total_volumen_seleccionado) + CDbl(Me.lv_pendientes.selectedItem.SubItems(6)), "###,##0.00000")
      End If
  Next var_j

End Sub

Private Sub cmd_todos_2_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = Me.lv_asignados.ListItems.Count
   For i = 1 To n
       Me.lv_asignados.ListItems.Item(i).SubItems(4) = "*"
       Me.lv_asignados.ListItems.Item(i).Bold = True
       Me.lv_asignados.ListItems.Item(i).ForeColor = &H8000&
       Me.lv_asignados.ListItems.Item(i).ListSubItems(1).Bold = True
       Me.lv_asignados.ListItems.Item(i).ListSubItems(2).Bold = True
       Me.lv_asignados.ListItems.Item(i).ListSubItems(3).Bold = True
       Me.lv_asignados.ListItems.Item(i).ListSubItems(4).Bold = True
       Me.lv_asignados.ListItems.Item(i).ListSubItems(6).Bold = True
       Me.lv_asignados.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       Me.lv_asignados.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
       Me.lv_asignados.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
       Me.lv_asignados.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
       Me.lv_asignados.ListItems.Item(i).ListSubItems(6).ForeColor = &H8000&
   Next
   lv_asignados.Refresh

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
   n = lv_pendientes.ListItems.Count
   Me.txt_total_volumen_seleccionado = "0.00000"
   For i = 1 To n
       lv_pendientes.ListItems.Item(i).SubItems(4) = "*"
       lv_pendientes.ListItems.Item(i).Bold = True
       lv_pendientes.ListItems.Item(i).ForeColor = &H8000&
       lv_pendientes.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_pendientes.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_pendientes.ListItems.Item(i).ListSubItems(3).Bold = True
       lv_pendientes.ListItems.Item(i).ListSubItems(4).Bold = True
       lv_pendientes.ListItems.Item(i).ListSubItems(6).Bold = True
       Me.txt_total_volumen_seleccionado = Format(CDbl(Me.txt_total_volumen_seleccionado) + CDbl(lv_pendientes.ListItems(i).ListSubItems(6).Text), "###,##0.00000")
       lv_pendientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_pendientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
       lv_pendientes.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
       lv_pendientes.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
       lv_pendientes.ListItems.Item(i).ListSubItems(6).ForeColor = &H8000&
   Next
   lv_pendientes.Refresh

End Sub

Private Sub Command2_Click()
   
   For var_j = 1 To Me.lv_asignados.ListItems.Count
       Me.lv_asignados.ListItems.Item(var_j).Selected = True
       If Me.lv_asignados.selectedItem.SubItems(4) = "*" Then
          Set list_item = lv_pendientes.ListItems.Add(, , Me.lv_asignados.selectedItem)
          list_item.SubItems(1) = Me.lv_asignados.selectedItem.SubItems(1)
          list_item.SubItems(2) = Me.lv_asignados.selectedItem.SubItems(2)
          list_item.SubItems(3) = Me.lv_asignados.selectedItem.SubItems(3)
          list_item.SubItems(3) = Me.lv_asignados.selectedItem.SubItems(4)
          list_item.SubItems(6) = Me.lv_asignados.selectedItem.SubItems(6)
          list_item.SubItems(7) = Me.lv_asignados.selectedItem.SubItems(7)
          Me.txt_total_volumen = CDbl(Me.txt_total_volumen) - CDbl(Me.lv_asignados.selectedItem.SubItems(6))
       End If
   Next var_j
   For var_j = Me.lv_asignados.ListItems.Count To 1 Step -1
       Me.lv_asignados.ListItems.Item(var_j).Selected = True
       If Me.lv_asignados.selectedItem.SubItems(4) = "*" Then
          lv_asignados.ListItems.Remove (lv_asignados.selectedItem.Index)
       End If
   Next var_j
   Me.lv_asignados.Refresh
   Me.lv_pendientes.Refresh
   

End Sub

Private Sub Form_Load()
   Me.txt_volumen_unidad = Format(var_volumen_transporte, "###,##0.00000")
   Me.txt_porcentaje = "0"
   Me.txt_total_volumen_seleccionado = "0.00000"
   rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   rs.Open "SELECT DISTINCT SOURCE_HEADER_NUMBER FROM WSH_DELIVERABLES_V WHERE RELEASED_STATUS = 'Y'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_cadena = ""
   While Not rs.EOF
         If var_cadena = "" Then
            var_cadena = CStr(rs!source_header_number)
         Else
            var_cadena = var_cadena + "," + CStr(rs!source_header_number)
         End If
         rs.MoveNext
   Wend
   'var_cadena = "324962,325363,325434,325541,325549"
   'MsgBox var_cadena
   rs.Close
   var_consecutivo_orden = 0
   lbl_anden = lbl_anden + " " + CStr(var_anden_global)
   rs.Open "select * from tb_oracle_pedidos_asignados_embarques where piezas > 0 and agente in (" + var_agente_asignar + "') and embarque = 0 and pedido in(" + var_cadena + ") AND ORGANIZACION = " + CStr(var_unidad_organizacional), cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_pendientes.ListItems.Add(, , rs!pedido)
         'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
         list_item.SubItems(1) = IIf(IsNull(rs!nombre_agente), "", rs!nombre_agente)
         'rsaux.Close
         list_item.SubItems(2) = rs!Cliente
         list_item.SubItems(3) = Format(rs!PIEZAS, "###,###,##0.00")
                  
         strconsulta = "select source_header_number, sum(src_requested_quantity * to_number(nvl(a.unit_volume,'0'))) as volumen from xxvia_system_items_b a, wsh_deliverables_v b where a.unit_volume is not null and a.inventory_item_id = b.inventory_item_id and released_status = 'Y' and source_header_number = ? and a.organization_id = b.organization_id group by source_header_number "
         'strconsulta = "select source_header_number, sum(src_requested_quantity * to_number(nvl(a.unit_volume,'0'))) as volumen from xxvia_system_items_b a, wsh_deliverables_v b where a.unit_volume is not null and a.inventory_item_id = b.inventory_item_id and released_status = 'Y' and source_header_number = ? group by source_header_number "
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rs!pedido))
              .Parameters.Append parametro
         End With
         Set rsaux6 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         var_volumen = 0
         If Not rsaux6.EOF Then
            var_volumen = IIf(IsNull(rsaux6!VOLUMEN), 0, rsaux6!VOLUMEN)
         End If
         rsaux6.Close
         list_item.SubItems(6) = Format(var_volumen, "###,###,##0.00000")
         
         strconsulta = "select * from oe_order_headers_all where order_number = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
              .Parameters.Append parametro
         End With
         Set rsaux6 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
            strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                 .Parameters.Append parametro
            End With
            Set rsaux8 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
            rsaux8.Close
        
            
            
            strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                  Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                 .Parameters.Append parametro
            End With
            Set rsaux7 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If rsaux7.EOF Then
               var_establecimiento = ""
            Else
               var_establecimiento = rsaux7!attribute1
            End If
         Else
            var_establecimiento = rsaux6!ship_to_org_id
         End If
         If rsaux7.State = 1 Then
            rsaux7.Close
         End If
         list_item.SubItems(7) = var_establecimiento
         rsaux6.Close
         rs.MoveNext
   Wend
   rs.Close
   
   x = 0
   If x = 1 Then
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "SELECT DISTINCT SOURCE_HEADER_NUMBER FROM WSH_DELIVERABLES_V WHERE RELEASED_STATUS= 'Y' AND CREATION_dATE >= TO_dATE('01-03-2012','DD-MM-YYYY') AND CREATION_dATE < TO_dATE('31-03-2012','DD-MM-YYYY') AND ORGANIZATION_ID =  93"
            var_Cadena_pedidos = ""
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  If var_Cadena_pedidos = "" Then
                     var_Cadena_pedidos = CStr(rs!source_header_number)
                  Else
                     var_Cadena_pedidos = var_Cadena_pedidos + "," + CStr(rs!source_header_number)
                  End If
                  rs.MoveNext
            Wend
            rs.Close
            If var_Cadena_pedidos <> "" Then
               var_Cadena_pedidos = Mid(var_Cadena_pedidos, 1, 1000)
               var_cadena = "SELECT distinct source_document_id, source_header_type_name, TRUNC(A.LAST_UPDATE_DATE) AS FECHA, source_header_number, D.COLLECTOR_ID, HL.ADDRESS1 AS CUSTOMER_NAME, D.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, xxvia_vw_agentes D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCAS.CUST_ACCOUNT_ID = D.CUST_ACCOUNT_ID AND to_number(source_header_number) IN (" + var_Cadena_pedidos + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y'"
   
               rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               numero_items_permisos = 0
               While Not rs.EOF
                     nombre_cliente = rs!customer_name
             
                     If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rs!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                        rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(IIf(IsNull(rs!source_document_id), "0", rs!source_document_id)) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           nombre_cliente = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                        End If
                        rsaux2.Close
                     End If
                     rsaux1.Open "select sum(requested_quantity) as cantidad from wsh_deliverables_v where source_header_number = " + CStr(rs!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rsaux.Open "INSERT INTO TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES (AGENTE, PEDIDO, CLIENTE, PIEZAS, EMBARQUE, DIA, MES, AÑO, ORGANIZACION) VALUES ('" + CStr(rs!collector_id) + "', '" + CStr(rs!source_header_number) + "','" + nombre_cliente + "'," + CStr(rsaux1!cantidad) + ", 1000," + CStr(Day(Date)) + "," + CStr(Month(Date)) + "," + CStr(Year(Date)) + "," + CStr(var_unidad_organizacional) + ")", cnn, adOpenDynamic, adLockOptimistic
                     rsaux1.Close
                     rs.MoveNext
               Wend
               rs.Close
            End If
   End If
End Sub

Private Sub lv_asignados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_asignados, ColumnHeader)
End Sub

Private Sub lv_asignados_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim numero_lineas As Integer
      Dim numero_seleccionado1 As Integer
      Dim numero_seleccionado2 As Integer
      Dim primera_vez As Boolean
      Dim segunda_vez As Boolean
      Dim i As Integer
      Dim n As Integer
      Dim list_item As ListItem
      i = Me.lv_asignados.selectedItem.Index
      If Me.lv_asignados.selectedItem.SubItems(4) = "*" Then
          Me.lv_asignados.ListItems.Item(i).SubItems(4) = " "
          Me.lv_asignados.ListItems.Item(i).Bold = False
          Me.lv_asignados.ListItems.Item(i).ForeColor = &H80000012
          Me.lv_asignados.ListItems.Item(i).ListSubItems(1).Bold = False
          Me.lv_asignados.ListItems.Item(i).ListSubItems(2).Bold = False
          Me.lv_asignados.ListItems.Item(i).ListSubItems(3).Bold = False
          Me.lv_asignados.ListItems.Item(i).ListSubItems(4).Bold = False
          Me.lv_asignados.ListItems.Item(i).ListSubItems(6).Bold = False
          Me.lv_asignados.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
          Me.lv_asignados.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
          Me.lv_asignados.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
          Me.lv_asignados.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
          Me.lv_asignados.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
      Else
         Me.lv_asignados.ListItems.Item(i).SubItems(4) = "*"
         Me.lv_asignados.ListItems.Item(i).Bold = True
         Me.lv_asignados.ListItems.Item(i).ForeColor = &H8000&
         Me.lv_asignados.ListItems.Item(i).ListSubItems(1).Bold = True
         Me.lv_asignados.ListItems.Item(i).ListSubItems(2).Bold = True
         Me.lv_asignados.ListItems.Item(i).ListSubItems(3).Bold = True
         Me.lv_asignados.ListItems.Item(i).ListSubItems(4).Bold = True
         Me.lv_asignados.ListItems.Item(i).ListSubItems(6).Bold = True
         Me.lv_asignados.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
         Me.lv_asignados.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
         Me.lv_asignados.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
         Me.lv_asignados.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
         Me.lv_asignados.ListItems.Item(i).ListSubItems(6).ForeColor = &H8000&
      End If
      Me.lv_asignados.Refresh
   End If
End Sub

Private Sub lv_pendientes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_pendientes, ColumnHeader)
End Sub

Private Sub lv_pendientes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim numero_lineas As Integer
      Dim numero_seleccionado1 As Integer
      Dim numero_seleccionado2 As Integer
      Dim primera_vez As Boolean
      Dim segunda_vez As Boolean
      Dim i As Integer
      Dim n As Integer
      Dim list_item As ListItem
      i = lv_pendientes.selectedItem.Index
      If lv_pendientes.selectedItem.SubItems(4) = "*" Then
          lv_pendientes.ListItems.Item(i).SubItems(4) = " "
          lv_pendientes.ListItems.Item(i).Bold = False
          lv_pendientes.ListItems.Item(i).ForeColor = &H80000012
          lv_pendientes.ListItems.Item(i).ListSubItems(1).Bold = False
          lv_pendientes.ListItems.Item(i).ListSubItems(2).Bold = False
          lv_pendientes.ListItems.Item(i).ListSubItems(3).Bold = False
          lv_pendientes.ListItems.Item(i).ListSubItems(4).Bold = False
          lv_pendientes.ListItems.Item(i).ListSubItems(6).Bold = False
          lv_pendientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
          lv_pendientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
          lv_pendientes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
          lv_pendientes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
          lv_pendientes.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
          Me.txt_total_volumen_seleccionado = Format(CDbl(Me.txt_total_volumen_seleccionado) - CDbl(Me.lv_pendientes.ListItems.Item(i).ListSubItems(6).Text), "###,##0.00000")
      Else
         lv_pendientes.ListItems.Item(i).SubItems(4) = "*"
         lv_pendientes.ListItems.Item(i).Bold = True
         lv_pendientes.ListItems.Item(i).ForeColor = &H8000&
         lv_pendientes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_pendientes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_pendientes.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_pendientes.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_pendientes.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_pendientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
         lv_pendientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
         lv_pendientes.ListItems.Item(i).ListSubItems(3).ForeColor = &H8000&
         lv_pendientes.ListItems.Item(i).ListSubItems(4).ForeColor = &H8000&
         lv_pendientes.ListItems.Item(i).ListSubItems(6).ForeColor = &H8000&
          Me.txt_total_volumen_seleccionado = Format(CDbl(Me.txt_total_volumen_seleccionado) + CDbl(Me.lv_pendientes.ListItems.Item(i).ListSubItems(6).Text), "###,##0.00000")
      End If
      lv_pendientes.Refresh
   End If
End Sub

Private Sub Text2_Change()

End Sub

Private Sub txt_clave_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_tipo_asigna_ruta = 2
      frmoracle_asignar_ruta.Show 1
      Me.txt_clave = var_ruta_distribucion
      rs.Open "select * from XXVIA_vw_RUTAS_DISTRIBUCION where ruta = '" + Me.txt_clave + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         rsaux1.Open "select * from XXVIA_VW_CLIENTES_RUTAS_DISTR  where ruta = '" + Me.txt_clave + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         VAR_CN_TEXTILERA = 0
         var_agente = ""
         var_cliente = ""
         VAR_PRIORIDAD_TEXT = 0
         While Not rsaux1.EOF
                
               If IIf(IsNull(rsaux1!cn_textilera), 0, rsaux1!cn_textilera) = 1 Then
                  VAR_CN_TEXTILERA = 1000
                  var_agente = IIf(IsNull(rsaux1!nombre_titular), "", rsaux1!nombre_titular)
                  var_cliente = IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento)
                  VAR_PRIORIDAD_TEXT = IIf(IsNull(rsaux1!prioridad), "0", rsaux1!prioridad)
               End If
               For var_j = 1 To Me.lv_pendientes.ListItems.Count
                   Me.lv_pendientes.ListItems.Item(var_j).Selected = True
                   If Me.lv_pendientes.selectedItem.SubItems(7) = IIf(IsNull(rsaux1!ESTABLECIMIENTO), "", rsaux1!ESTABLECIMIENTO) Then
                      Me.lv_pendientes.selectedItem.SubItems(5) = IIf(IsNull(rsaux1!prioridad), 0, rsaux1!prioridad)
                      Me.lv_pendientes.selectedItem.SubItems(4) = "*"
                      If lv_pendientes.selectedItem.SubItems(4) = "*" Then
                         lv_pendientes.ListItems.Item(var_j).Bold = True
                         lv_pendientes.ListItems.Item(var_j).ForeColor = &H8000&
                         lv_pendientes.ListItems.Item(var_j).ListSubItems(1).Bold = True
                         lv_pendientes.ListItems.Item(var_j).ListSubItems(2).Bold = True
                         lv_pendientes.ListItems.Item(var_j).ListSubItems(3).Bold = True
                         lv_pendientes.ListItems.Item(var_j).ListSubItems(4).Bold = True
                         lv_pendientes.ListItems.Item(var_j).ListSubItems(7).Bold = True
                         lv_pendientes.ListItems.Item(var_j).ListSubItems(6).Bold = True
                         lv_pendientes.ListItems.Item(var_j).ListSubItems(1).ForeColor = &H8000&
                         lv_pendientes.ListItems.Item(var_j).ListSubItems(2).ForeColor = &H8000&
                         lv_pendientes.ListItems.Item(var_j).ListSubItems(3).ForeColor = &H8000&
                         lv_pendientes.ListItems.Item(var_j).ListSubItems(4).ForeColor = &H8000&
                         lv_pendientes.ListItems.Item(var_j).ListSubItems(5).ForeColor = &H8000&
                         lv_pendientes.ListItems.Item(var_j).ListSubItems(6).ForeColor = &H8000&
                         Me.txt_total_volumen_seleccionado = Format(CDbl(Me.txt_total_volumen_seleccionado) + CDbl(Me.lv_pendientes.ListItems.Item(var_j).ListSubItems(6).Text), "###,##0.00000")
                      End If
                      lv_pendientes.Refresh
                   End If
               Next var_j
               rsaux1.MoveNext
         Wend
         rsaux1.Close
         var_encontro = 0
         If VAR_CN_TEXTILERA = 1 Then
            For var_j = 1 To Me.lv_pendientes.ListItems.Count
                Me.lv_pendientes.ListItems.Item(var_j).Selected = True
                cnn.BeginTrans
                rsaux1.Open "select max(pedido) from tb_oracle_consecutivo_pedidos_cn_textilera", cnn, adOpenDynamic, adLockOptimistic
                var_consecutivo = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value) + 1
                rsaux2.Open "update tb_oracle_consecutivo_pedidos_cn_textilera set pedido = pedido + 1", cnn, adOpenDynamic, adLockOptimistic
                rsaux1.Close
                cnn.CommitTrans
                If Me.lv_pendientes.selectedItem = var_consecutivo Then
                   var_encontro = 1
                   Me.lv_pendientes.selectedItem.SubItems(4) = "*"
                   Me.lv_pendientes.selectedItem.SubItems(5) = VAR_PRIORIDAD_TEXT
                End If
            Next var_j
            For var_j = 1 To Me.lv_pendientes.ListItems.Count
                Me.lv_pendientes.ListItems.Item(var_j).Selected = True
                If lv_pendientes.selectedItem.SubItems(4) = "*" Then
                   lv_pendientes.ListItems.Item(var_j).Bold = True
                   lv_pendientes.ListItems.Item(var_j).ForeColor = &H8000&
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(1).Bold = True
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(2).Bold = True
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(3).Bold = True
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(4).Bold = True
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(7).Bold = True
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(6).Bold = True
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(1).ForeColor = &H8000&
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(2).ForeColor = &H8000&
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(3).ForeColor = &H8000&
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(4).ForeColor = &H8000&
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(5).ForeColor = &H8000&
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(6).ForeColor = &H8000&
                End If
                lv_pendientes.Refresh
            Next var_j
         End If
         If var_encontro = 0 Then
            Set list_item = lv_pendientes.ListItems.Add(, , var_consecutivo)
            list_item.SubItems(1) = IIf(IsNull(var_agente), "", var_agente)
            list_item.SubItems(2) = var_cliente
            list_item.SubItems(3) = Format(0, "###,###,##0.00")
            list_item.SubItems(4) = "*"
            list_item.SubItems(5) = VAR_PRIORIDAD_TEXT
            list_item.SubItems(6) = Format(0, "###,###,##0.00")
            For var_j = 1 To Me.lv_pendientes.ListItems.Count
                Me.lv_pendientes.ListItems.Item(var_j).Selected = True
                If lv_pendientes.selectedItem.SubItems(4) = "*" Then
                   lv_pendientes.ListItems.Item(var_j).Bold = True
                   lv_pendientes.ListItems.Item(var_j).ForeColor = &H8000&
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(1).Bold = True
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(2).Bold = True
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(3).Bold = True
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(4).Bold = True
                   'lv_pendientes.ListItems.Item(var_j).ListSubItems(7).Bold = True
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(6).Bold = True
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(1).ForeColor = &H8000&
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(2).ForeColor = &H8000&
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(3).ForeColor = &H8000&
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(4).ForeColor = &H8000&
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(5).ForeColor = &H8000&
                   lv_pendientes.ListItems.Item(var_j).ListSubItems(6).ForeColor = &H8000&
                End If
                lv_pendientes.Refresh
            Next var_j
         
         End If
         
      End If
      rs.Close
   End If
End Sub

Private Sub txt_porcentaje_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_total_volumen_Change()
    If Me.txt_volumen_unidad > 0 Then
       Me.txt_porcentaje = Format((CDbl(Me.txt_total_volumen) / CDbl(Me.txt_volumen_unidad)) * 100, "###,##0.000")
    Else
       Me.txt_porcentaje = 0
    End If
End Sub

Private Sub txt_total_volumen_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_total_volumen_seleccionado_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_volumen_unidad_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
