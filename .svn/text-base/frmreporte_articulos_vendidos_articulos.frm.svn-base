VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_articulos_vendidos_articulos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Articulos Vendidos por Artículo"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   585
      Left            =   135
      TabIndex        =   38
      Top             =   -30
      Width           =   11385
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   10080
         TabIndex        =   41
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   8085
         TabIndex        =   40
         Top             =   180
         Width           =   1080
      End
      Begin VB.TextBox txt_tipo_reporte 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   180
         Width           =   6315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   9780
         TabIndex        =   43
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   7635
         TabIndex        =   42
         Top             =   240
         Width           =   420
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "<<< &Anterior"
      Height          =   435
      Left            =   8955
      TabIndex        =   36
      Top             =   6825
      Width           =   1350
   End
   Begin VB.CommandButton cmd_siguiente 
      Caption         =   "&Siguiente >>>"
      Height          =   435
      Left            =   10305
      TabIndex        =   35
      Top             =   6825
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   " Caracteristicas del artículo"
      Height          =   6180
      Left            =   150
      TabIndex        =   0
      Top             =   555
      Width           =   11370
      Begin VB.Frame Frame5 
         Height          =   90
         Left            =   5670
         TabIndex        =   37
         Top             =   3150
         Width           =   5685
      End
      Begin VB.Frame Frame4 
         Height          =   90
         Left            =   15
         TabIndex        =   34
         Top             =   3150
         Width           =   5625
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   6135
         Left            =   5640
         TabIndex        =   33
         Top             =   15
         Width           =   30
      End
      Begin VB.TextBox txt_busqueda_talla 
         Height          =   315
         Left            =   6630
         TabIndex        =   30
         Top             =   3645
         Width           =   4635
      End
      Begin VB.CommandButton Command15 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7005
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   3270
         Width           =   330
      End
      Begin VB.CommandButton Command14 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6345
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":0216
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Marcar (Enter)"
         Top             =   3270
         Width           =   330
      End
      Begin VB.CommandButton Command13 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6675
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":0460
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   3270
         Width           =   330
      End
      Begin VB.CommandButton Command12 
         Height          =   315
         Left            =   5685
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   3270
         Width           =   330
      End
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6015
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   3270
         Width           =   330
      End
      Begin VB.TextBox txt_busqueda_linea 
         Height          =   315
         Left            =   6630
         TabIndex        =   22
         Top             =   660
         Width           =   4635
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7005
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":084A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   285
         Width           =   330
      End
      Begin VB.CommandButton Command9 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6345
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":0A60
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Marcar (Enter)"
         Top             =   285
         Width           =   330
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6675
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":0CAA
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   285
         Width           =   330
      End
      Begin VB.CommandButton Command7 
         Height          =   315
         Left            =   5685
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":0D7C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   285
         Width           =   330
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6015
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":0E7E
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   285
         Width           =   330
      End
      Begin VB.TextBox txt_busqueda_familia 
         Height          =   315
         Left            =   900
         TabIndex        =   14
         Top             =   3645
         Width           =   4635
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1365
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":1094
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   3270
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   705
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":12AA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Marcar (Enter)"
         Top             =   3270
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1035
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":14F4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   3270
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   45
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":15C6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   3270
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":16C8
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   3270
         Width           =   330
      End
      Begin VB.TextBox txt_busqueda_catalogo 
         Height          =   315
         Left            =   900
         TabIndex        =   6
         Top             =   660
         Width           =   4635
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1365
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":18DE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   285
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   705
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":1AF4
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Marcar (Enter)"
         Top             =   285
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1035
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":1D3E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   285
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   45
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":1E10
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   285
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmreporte_articulos_vendidos_articulos.frx":1F12
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   285
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_catalogos 
         Height          =   2130
         Left            =   45
         TabIndex        =   7
         Top             =   1005
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   3757
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
      Begin MSComctlLib.ListView lv_familias 
         Height          =   2130
         Left            =   45
         TabIndex        =   15
         Top             =   3990
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   3757
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
      Begin MSComctlLib.ListView lv_lineas 
         Height          =   2130
         Left            =   5775
         TabIndex        =   23
         Top             =   1005
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   3757
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
      Begin MSComctlLib.ListView lv_tallas 
         Height          =   2130
         Left            =   5775
         TabIndex        =   31
         Top             =   3990
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   3757
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
      Begin VB.Label Label4 
         Caption         =   "Busqueda"
         Height          =   255
         Left            =   5820
         TabIndex        =   32
         Top             =   3675
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Busqueda"
         Height          =   255
         Left            =   5820
         TabIndex        =   24
         Top             =   690
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "Busqueda"
         Height          =   255
         Left            =   90
         TabIndex        =   16
         Top             =   3675
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Busqueda"
         Height          =   255
         Left            =   90
         TabIndex        =   8
         Top             =   690
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmreporte_articulos_vendidos_articulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancelar_Click()
   Unload Me
End Sub

Private Sub cmd_invertir_Click()
   n = lv_catalogos.ListItems.Count
   For i = 1 To n
      lv_catalogos.ListItems.Item(i).Selected = True
      If lv_catalogos.selectedItem.SubItems(2) = "*" Then
         lv_catalogos.selectedItem.SubItems(2) = ""
         lv_catalogos.ListItems.Item(i).Bold = False
         lv_catalogos.ListItems.Item(i).ForeColor = &H80000012
         lv_catalogos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_catalogos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_catalogos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_catalogos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_catalogos.selectedItem.SubItems(2) = "*"
         lv_catalogos.ListItems.Item(i).Bold = True
         lv_catalogos.ListItems.Item(i).ForeColor = &HFF0000
         lv_catalogos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_catalogos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_catalogos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_catalogos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_catalogos.selectedItem.Index
   If lv_catalogos.selectedItem.SubItems(2) = "*" Then
      lv_catalogos.selectedItem.SubItems(2) = ""
      lv_catalogos.ListItems.Item(i).Bold = False
      lv_catalogos.ListItems.Item(i).ForeColor = &H80000012
      lv_catalogos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_catalogos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_catalogos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_catalogos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_catalogos.Refresh
   Else
      lv_catalogos.selectedItem.SubItems(2) = "*"
      lv_catalogos.ListItems.Item(i).Bold = True
      lv_catalogos.ListItems.Item(i).ForeColor = &HFF0000
      lv_catalogos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_catalogos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_catalogos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_catalogos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_catalogos.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_catalogos.ListItems.Count
   For i = 1 To n
      lv_catalogos.ListItems.Item(i).Selected = True
      lv_catalogos.selectedItem.SubItems(2) = ""
      lv_catalogos.ListItems.Item(i).Bold = False
      lv_catalogos.ListItems.Item(i).ForeColor = &H80000012
      lv_catalogos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_catalogos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_catalogos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_catalogos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_catalogos.Refresh
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_catalogos.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_catalogos.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_catalogos.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_catalogos.selectedItem.SubItems(2) = "*"
         lv_catalogos.ListItems.Item(i).Bold = True
         lv_catalogos.ListItems.Item(i).ForeColor = &HFF0000
         lv_catalogos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_catalogos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_catalogos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_catalogos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_catalogos.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_catalogos.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_siguiente_Click()
   var_cadena_reporte_articulos_catalogos = ""
   For var_i = 1 To lv_catalogos.ListItems.Count
       lv_catalogos.ListItems.Item(var_i).Selected = True
       If lv_catalogos.selectedItem.SubItems(2) = "*" Then
          If var_cadena_reporte_articulos_catalogos = "" Then
             var_cadena_reporte_articulos_catalogos = " VCHA_ART_CATALOGO_VIGENTE = '" + Trim(lv_catalogos.selectedItem) + "' "
          Else
             var_cadena_reporte_articulos_catalogos = var_cadena_reporte_articulos_catalogos + " OR  VCHA_ART_CATALOGO_VIGENTE = '" + Trim(lv_catalogos.selectedItem) + "' "
          End If
       End If
   Next var_i
   var_cadena_reporte_articulos_familias = ""
   For var_i = 1 To lv_familias.ListItems.Count
       lv_familias.ListItems.Item(var_i).Selected = True
       If lv_familias.selectedItem.SubItems(2) = "*" Then
          If var_cadena_reporte_articulos_familias = "" Then
             var_cadena_reporte_articulos_familias = " VCHA_DIS_DISEÑO_ID = '" + Trim(lv_familias.selectedItem) + "' "
          Else
             var_cadena_reporte_articulos_familias = var_cadena_reporte_articulos_familias + " OR  VCHA_DIS_DISEÑO_ID = '" + Trim(lv_familias.selectedItem) + "' "
          End If
       End If
   Next var_i
   var_cadena_reporte_articulos_lineas = ""
   For var_i = 1 To lv_lineas.ListItems.Count
       lv_lineas.ListItems.Item(var_i).Selected = True
       If lv_lineas.selectedItem.SubItems(2) = "*" Then
          If var_cadena_reporte_articulos_lineas = "" Then
             var_cadena_reporte_articulos_lineas = " VCHA_LIN_LINEA_ID = '" + Trim(lv_lineas.selectedItem) + "' "
          Else
             var_cadena_reporte_articulos_lineas = var_cadena_reporte_articulos_lineas + " OR  VCHA_LIN_LINEA_ID = '" + Trim(lv_lineas.selectedItem) + "' "
          End If
       End If
   Next var_i
   var_cadena_reporte_articulos_tallas = ""
   For var_i = 1 To lv_tallas.ListItems.Count
       lv_tallas.ListItems.Item(var_i).Selected = True
       If lv_tallas.selectedItem.SubItems(2) = "*" Then
          If var_cadena_reporte_articulos_tallas = "" Then
             var_cadena_reporte_articulos_tallas = " vcha_tal_talla_id = '" + Trim(lv_tallas.selectedItem) + "' "
          Else
             var_cadena_reporte_articulos_tallas = var_cadena_reporte_articulos_tallas + " OR  vcha_tal_talla_id = '" + Trim(lv_tallas.selectedItem) + "' "
          End If
       End If
   Next var_i
   Me.Enabled = False
   var_activa_forma_ordensurtido = Me.Name
   frmreporte_articulos_vendidos_articulos_seleccionados.Show
   
End Sub

Private Sub cmd_todos_Click()
   n = lv_catalogos.ListItems.Count
   For i = 1 To n
      lv_catalogos.ListItems.Item(i).Selected = True
      lv_catalogos.selectedItem.SubItems(2) = "*"
      lv_catalogos.ListItems.Item(i).Bold = True
      lv_catalogos.ListItems.Item(i).ForeColor = &HFF0000
      lv_catalogos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_catalogos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_catalogos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_catalogos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_catalogos.Refresh
End Sub

Private Sub Command1_Click()
   n = lv_familias.ListItems.Count
   For i = 1 To n
      lv_familias.ListItems.Item(i).Selected = True
      lv_familias.selectedItem.SubItems(2) = "*"
      lv_familias.ListItems.Item(i).Bold = True
      lv_familias.ListItems.Item(i).ForeColor = &HFF0000
      lv_familias.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_familias.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_familias.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_familias.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_familias.Refresh
End Sub

Private Sub Command10_Click()
   n = lv_lineas.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_lineas.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_lineas.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_lineas.selectedItem.SubItems(2) = "*"
         lv_lineas.ListItems.Item(i).Bold = True
         lv_lineas.ListItems.Item(i).ForeColor = &HFF0000
         lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_lineas.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_lineas.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub Command11_Click()
   n = lv_tallas.ListItems.Count
   For i = 1 To n
      lv_tallas.ListItems.Item(i).Selected = True
      lv_tallas.selectedItem.SubItems(2) = "*"
      lv_tallas.ListItems.Item(i).Bold = True
      lv_tallas.ListItems.Item(i).ForeColor = &HFF0000
      lv_tallas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_tallas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_tallas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_tallas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_tallas.Refresh
End Sub

Private Sub Command12_Click()
   n = lv_tallas.ListItems.Count
   For i = 1 To n
      lv_tallas.ListItems.Item(i).Selected = True
      lv_tallas.selectedItem.SubItems(2) = ""
      lv_tallas.ListItems.Item(i).Bold = False
      lv_tallas.ListItems.Item(i).ForeColor = &H80000012
      lv_tallas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_tallas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_tallas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_tallas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_tallas.Refresh
End Sub

Private Sub Command13_Click()
   n = lv_tallas.ListItems.Count
   For i = 1 To n
      lv_tallas.ListItems.Item(i).Selected = True
      If lv_tallas.selectedItem.SubItems(2) = "*" Then
         lv_tallas.selectedItem.SubItems(2) = ""
         lv_tallas.ListItems.Item(i).Bold = False
         lv_tallas.ListItems.Item(i).ForeColor = &H80000012
         lv_tallas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_tallas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_tallas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_tallas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_tallas.selectedItem.SubItems(2) = "*"
         lv_tallas.ListItems.Item(i).Bold = True
         lv_tallas.ListItems.Item(i).ForeColor = &HFF0000
         lv_tallas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_tallas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_tallas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_tallas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub Command14_Click()
   i = lv_tallas.selectedItem.Index
   If lv_tallas.selectedItem.SubItems(2) = "*" Then
      lv_tallas.selectedItem.SubItems(2) = ""
      lv_tallas.ListItems.Item(i).Bold = False
      lv_tallas.ListItems.Item(i).ForeColor = &H80000012
      lv_tallas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_tallas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_tallas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_tallas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_tallas.Refresh
   Else
      lv_tallas.selectedItem.SubItems(2) = "*"
      lv_tallas.ListItems.Item(i).Bold = True
      lv_tallas.ListItems.Item(i).ForeColor = &HFF0000
      lv_tallas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_tallas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_tallas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_tallas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_tallas.Refresh
   End If
End Sub

Private Sub Command15_Click()
   n = lv_tallas.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_tallas.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_tallas.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_tallas.selectedItem.SubItems(2) = "*"
         lv_tallas.ListItems.Item(i).Bold = True
         lv_tallas.ListItems.Item(i).ForeColor = &HFF0000
         lv_tallas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_tallas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_tallas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_tallas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_tallas.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_tallas.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub Command2_Click()
   n = lv_familias.ListItems.Count
   For i = 1 To n
      lv_familias.ListItems.Item(i).Selected = True
      lv_familias.selectedItem.SubItems(2) = ""
      lv_familias.ListItems.Item(i).Bold = False
      lv_familias.ListItems.Item(i).ForeColor = &H80000012
      lv_familias.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_familias.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_familias.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_familias.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_familias.Refresh
End Sub

Private Sub Command3_Click()
   n = lv_familias.ListItems.Count
   For i = 1 To n
      lv_familias.ListItems.Item(i).Selected = True
      If lv_familias.selectedItem.SubItems(2) = "*" Then
         lv_familias.selectedItem.SubItems(2) = ""
         lv_familias.ListItems.Item(i).Bold = False
         lv_familias.ListItems.Item(i).ForeColor = &H80000012
         lv_familias.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_familias.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_familias.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_familias.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_familias.selectedItem.SubItems(2) = "*"
         lv_familias.ListItems.Item(i).Bold = True
         lv_familias.ListItems.Item(i).ForeColor = &HFF0000
         lv_familias.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_familias.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_familias.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_familias.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub Command4_Click()
   i = lv_familias.selectedItem.Index
   If lv_familias.selectedItem.SubItems(2) = "*" Then
      lv_familias.selectedItem.SubItems(2) = ""
      lv_familias.ListItems.Item(i).Bold = False
      lv_familias.ListItems.Item(i).ForeColor = &H80000012
      lv_familias.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_familias.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_familias.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_familias.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_familias.Refresh
   Else
      lv_familias.selectedItem.SubItems(2) = "*"
      lv_familias.ListItems.Item(i).Bold = True
      lv_familias.ListItems.Item(i).ForeColor = &HFF0000
      lv_familias.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_familias.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_familias.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_familias.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_familias.Refresh
   End If
End Sub

Private Sub Command5_Click()
   n = lv_familias.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_familias.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_familias.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_familias.selectedItem.SubItems(2) = "*"
         lv_familias.ListItems.Item(i).Bold = True
         lv_familias.ListItems.Item(i).ForeColor = &HFF0000
         lv_familias.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_familias.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_familias.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_familias.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_familias.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_familias.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub Command6_Click()
   n = lv_lineas.ListItems.Count
   For i = 1 To n
      lv_lineas.ListItems.Item(i).Selected = True
      lv_lineas.selectedItem.SubItems(2) = "*"
      lv_lineas.ListItems.Item(i).Bold = True
      lv_lineas.ListItems.Item(i).ForeColor = &HFF0000
      lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_lineas.Refresh
End Sub

Private Sub Command7_Click()
   n = lv_lineas.ListItems.Count
   For i = 1 To n
      lv_lineas.ListItems.Item(i).Selected = True
      lv_lineas.selectedItem.SubItems(2) = ""
      lv_lineas.ListItems.Item(i).Bold = False
      lv_lineas.ListItems.Item(i).ForeColor = &H80000012
      lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_lineas.Refresh
End Sub

Private Sub Command8_Click()
   n = lv_lineas.ListItems.Count
   For i = 1 To n
      lv_lineas.ListItems.Item(i).Selected = True
      If lv_lineas.selectedItem.SubItems(2) = "*" Then
         lv_lineas.selectedItem.SubItems(2) = ""
         lv_lineas.ListItems.Item(i).Bold = False
         lv_lineas.ListItems.Item(i).ForeColor = &H80000012
         lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_lineas.selectedItem.SubItems(2) = "*"
         lv_lineas.ListItems.Item(i).Bold = True
         lv_lineas.ListItems.Item(i).ForeColor = &HFF0000
         lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub Command9_Click()
   i = lv_lineas.selectedItem.Index
   If lv_lineas.selectedItem.SubItems(2) = "*" Then
      lv_lineas.selectedItem.SubItems(2) = ""
      lv_lineas.ListItems.Item(i).Bold = False
      lv_lineas.ListItems.Item(i).ForeColor = &H80000012
      lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_lineas.Refresh
   Else
      lv_lineas.selectedItem.SubItems(2) = "*"
      lv_lineas.ListItems.Item(i).Bold = True
      lv_lineas.ListItems.Item(i).ForeColor = &HFF0000
      lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_lineas.Refresh
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 65 Then
      cmd_cancelar_Click
   End If
   If Shift = 4 And KeyCode = 83 Then
      cmd_siguiente_Click
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   rs.Open "select distinct vcha_cat_catalogo_id, vcha_cat_nombre from tb_Catalogos order by vcha_cat_nombre ", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      If IsNull(rs!vcha_cat_catalogo_id) Then
      Else
         Set list_item = lv_catalogos.ListItems.Add(, , rs!vcha_cat_catalogo_id)
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_cat_NOMBRE), "", rs!VCHA_cat_NOMBRE)
         list_item.SubItems(2) = ""
      End If
      rs.MoveNext:
   Wend
   rs.Close
   If lv_catalogos.ListItems.Count > 8 Then
      lv_catalogos.ColumnHeaders(2).Width = 4220
   Else
      lv_catalogos.ColumnHeaders(2).Width = 4400.71
   End If

   rs.Open "select distinct vcha_LIN_LINEA_id, vcha_LIN_nombre from TB_LINEAS order by vcha_LIN_nombre ", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      If IsNull(rs!VCHA_LIN_LINEA_ID) Then
      Else
         Set list_item = lv_lineas.ListItems.Add(, , rs!VCHA_LIN_LINEA_ID)
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_lin_NOMBRE), "", rs!VCHA_lin_NOMBRE)
         list_item.SubItems(2) = ""
      End If
      rs.MoveNext:
   Wend
   rs.Close
   If lv_lineas.ListItems.Count > 8 Then
      lv_lineas.ColumnHeaders(2).Width = 4220
   Else
      lv_lineas.ColumnHeaders(2).Width = 4400.71
   End If

   rs.Open "select distinct vcha_DIS_DISEÑO_id, vcha_DIS_nombre from TB_DISEÑOS order by vcha_DIS_nombre ", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      If IsNull(rs!VCHA_DIS_DISEÑO_ID) Then
      Else
         Set list_item = lv_familias.ListItems.Add(, , rs!VCHA_DIS_DISEÑO_ID)
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_dis_NOMBRE), "", rs!VCHA_dis_NOMBRE)
         list_item.SubItems(2) = ""
      End If
      rs.MoveNext:
   Wend
   rs.Close
   If lv_familias.ListItems.Count > 8 Then
      lv_familias.ColumnHeaders(2).Width = 4220
   Else
      lv_familias.ColumnHeaders(2).Width = 4400.71
   End If

   rs.Open "select distinct vcha_TAL_TALLA_id, vcha_TAL_nombre from TB_TALLAs order by vcha_TAL_nombre ", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      If IsNull(rs!VCHA_TAL_TALLA_ID) Then
      Else
         Set list_item = lv_tallas.ListItems.Add(, , rs!VCHA_TAL_TALLA_ID)
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_tal_NOMBRE), "", rs!VCHA_tal_NOMBRE)
         list_item.SubItems(2) = ""
      End If
      rs.MoveNext:
   Wend
   rs.Close
   If lv_tallas.ListItems.Count > 8 Then
      lv_tallas.ColumnHeaders(2).Width = 4220
   Else
      lv_tallas.ColumnHeaders(2).Width = 4400.71
   End If

End Sub

Private Sub lv_catalogos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_catalogos, ColumnHeader)
End Sub

Private Sub lv_catalogos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_catalogos.selectedItem.Index
      If lv_catalogos.selectedItem.SubItems(2) = "*" Then
         lv_catalogos.selectedItem.SubItems(2) = ""
         lv_catalogos.ListItems.Item(i).Bold = False
         lv_catalogos.ListItems.Item(i).ForeColor = &H80000012
         lv_catalogos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_catalogos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_catalogos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_catalogos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_catalogos.Refresh
      Else
         lv_catalogos.selectedItem.SubItems(2) = "*"
         lv_catalogos.ListItems.Item(i).Bold = True
         lv_catalogos.ListItems.Item(i).ForeColor = &HFF0000
         lv_catalogos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_catalogos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_catalogos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_catalogos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_catalogos.Refresh
      End If
   End If
End Sub

Private Sub lv_familias_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_familias, ColumnHeader)
End Sub

Private Sub lv_familias_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_familias.selectedItem.Index
      If lv_familias.selectedItem.SubItems(2) = "*" Then
         lv_familias.selectedItem.SubItems(2) = ""
         lv_familias.ListItems.Item(i).Bold = False
         lv_familias.ListItems.Item(i).ForeColor = &H80000012
         lv_familias.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_familias.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_familias.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_familias.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_familias.Refresh
      Else
         lv_familias.selectedItem.SubItems(2) = "*"
         lv_familias.ListItems.Item(i).Bold = True
         lv_familias.ListItems.Item(i).ForeColor = &HFF0000
         lv_familias.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_familias.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_familias.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_familias.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_familias.Refresh
      End If
   End If
End Sub

Private Sub lv_lineas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lineas, ColumnHeader)
End Sub

Private Sub lv_lineas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_lineas.selectedItem.Index
      If lv_lineas.selectedItem.SubItems(2) = "*" Then
         lv_lineas.selectedItem.SubItems(2) = ""
         lv_lineas.ListItems.Item(i).Bold = False
         lv_lineas.ListItems.Item(i).ForeColor = &H80000012
         lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_lineas.Refresh
      Else
         lv_lineas.selectedItem.SubItems(2) = "*"
         lv_lineas.ListItems.Item(i).Bold = True
         lv_lineas.ListItems.Item(i).ForeColor = &HFF0000
         lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_lineas.Refresh
      End If
   End If
End Sub

Private Sub lv_tallas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_tallas, ColumnHeader)
End Sub

Private Sub lv_tallas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_tallas.selectedItem.Index
      If lv_tallas.selectedItem.SubItems(2) = "*" Then
         lv_tallas.selectedItem.SubItems(2) = ""
         lv_tallas.ListItems.Item(i).Bold = False
         lv_tallas.ListItems.Item(i).ForeColor = &H80000012
         lv_tallas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_tallas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_tallas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_tallas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_tallas.Refresh
      Else
         lv_tallas.selectedItem.SubItems(2) = "*"
         lv_tallas.ListItems.Item(i).Bold = True
         lv_tallas.ListItems.Item(i).ForeColor = &HFF0000
         lv_tallas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_tallas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_tallas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_tallas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_tallas.Refresh
      End If
   End If
End Sub

Private Sub txt_busqueda_Change()

End Sub


Private Sub Text1_Change()

End Sub

Private Sub txt_busqueda_catalogo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "select top 1 vcha_cat_nombre from tb_catalogos where vcha_cat_nombre like '%" + Me.txt_busqueda_catalogo + "%'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Call pro_busca_registro(lv_catalogos, rs!VCHA_cat_NOMBRE, False)
      Else
         MsgBox "No existen coincidencias", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub txt_busqueda_familia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "select top 1 vcha_dis_nombre from tb_diseños where vcha_dis_nombre like '%" + Me.txt_busqueda_familia + "%'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Call pro_busca_registro(lv_familias, rs!VCHA_dis_NOMBRE, False)
      Else
         MsgBox "No existen coincidencias", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub txt_busqueda_linea_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "select top 1 vcha_lin_nombre from tb_lineas where vcha_lin_nombre like '%" + Me.txt_busqueda_linea + "%'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Call pro_busca_registro(lv_lineas, rs!VCHA_lin_NOMBRE, False)
      Else
         MsgBox "No existen coincidencias", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub txt_busqueda_talla_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "select top 1 vcha_tal_nombre from tb_tallas where vcha_tal_nombre like '%" + Me.txt_busqueda_talla + "%'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Call pro_busca_registro(lv_tallas, rs!VCHA_tal_NOMBRE, False)
      Else
         MsgBox "No existen coincidencias", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes.Value = CDate(Me.txt_fin)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes.Value = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub
