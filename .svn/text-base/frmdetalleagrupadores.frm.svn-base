VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdetalleagrupadores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de agrupador"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmdetalleagrupadores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.Frame frm_lista 
      Height          =   3375
      Left            =   105
      TabIndex        =   37
      Top             =   570
      Width           =   5700
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2310
         Picture         =   "frmdetalleagrupadores.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   405
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1650
         Picture         =   "frmdetalleagrupadores.frx":0AE0
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Marcar (Enter)"
         Top             =   405
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1980
         Picture         =   "frmdetalleagrupadores.frx":0D2A
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   405
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   990
         Picture         =   "frmdetalleagrupadores.frx":0DFC
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   405
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1320
         Picture         =   "frmdetalleagrupadores.frx":0EFE
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   405
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Picture         =   "frmdetalleagrupadores.frx":1114
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   420
         Width           =   330
      End
      Begin VB.CommandButton cmd_cancelar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmdetalleagrupadores.frx":125E
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   420
         Width           =   330
      End
      Begin VB.Frame Frame6 
         Height          =   120
         Left            =   30
         TabIndex        =   40
         Top             =   705
         Width           =   5625
      End
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2430
         Left            =   45
         TabIndex        =   38
         Top             =   855
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   4286
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
            Object.Width           =   7232
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "MARCA"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   39
         Top             =   120
         Width           =   5625
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5430
      Picture         =   "frmdetalleagrupadores.frx":13A8
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      Picture         =   "frmdetalleagrupadores.frx":19E2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1110
      Picture         =   "frmdetalleagrupadores.frx":1AE4
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   780
      Picture         =   "frmdetalleagrupadores.frx":1BE6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   450
      Picture         =   "frmdetalleagrupadores.frx":1CB8
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmdetalleagrupadores.frx":1DBA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2490
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   26
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de agrupadores"
      Height          =   3465
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   5715
      Begin VB.TextBox txt_nombre_producto 
         Height          =   315
         Left            =   2595
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   21
         Top             =   2715
         Width           =   3015
      End
      Begin VB.TextBox txt_nombre_subtipo_articulo 
         Height          =   315
         Left            =   2595
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   23
         Top             =   3060
         Width           =   3015
      End
      Begin VB.TextBox txt_nombre_linea 
         Height          =   315
         Left            =   2595
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1380
         Width           =   3015
      End
      Begin VB.TextBox txt_nombre_sublinea 
         Height          =   315
         Left            =   2595
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1725
         Width           =   3015
      End
      Begin VB.TextBox txt_nombre_articulo 
         Height          =   315
         Left            =   2595
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         Top             =   495
         Width           =   3015
      End
      Begin VB.Frame Frame5 
         Height          =   75
         Left            =   15
         TabIndex        =   36
         Top             =   2070
         Width           =   5685
      End
      Begin VB.Frame Frame2 
         Height          =   75
         Left            =   15
         TabIndex        =   35
         Top             =   825
         Width           =   5670
      End
      Begin VB.TextBox txt_subtipo_producto 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   22
         Top             =   3060
         Width           =   1290
      End
      Begin VB.TextBox txt_producto 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   20
         Top             =   2715
         Width           =   1290
      End
      Begin VB.TextBox txt_sublinea 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1725
         Width           =   1290
      End
      Begin VB.TextBox txt_linea 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1380
         Width           =   1290
      End
      Begin VB.OptionButton opt_tipoagrupador 
         Caption         =   "Tipo Artículo"
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   19
         Top             =   2490
         Width           =   2175
      End
      Begin VB.OptionButton opt_tipoagrupador 
         Caption         =   "Producto"
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   18
         Top             =   2220
         Width           =   1065
      End
      Begin VB.OptionButton opt_tipoagrupador 
         Caption         =   "Sublinea"
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   13
         Top             =   1185
         Width           =   1035
      End
      Begin VB.OptionButton opt_tipoagrupador 
         Caption         =   "Linea"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   12
         Top             =   930
         Width           =   810
      End
      Begin VB.OptionButton opt_tipoagrupador 
         Caption         =   "Artículos individuales"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   255
         Width           =   2175
      End
      Begin VB.TextBox txt_articulo 
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   8
         Top             =   495
         Width           =   1290
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Artículo:"
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   30
         Top             =   3120
         Width           =   960
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   29
         Top             =   2775
         Width           =   690
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Sublinea:"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   28
         Top             =   1785
         Width           =   660
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Linea:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   27
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   10
         Top             =   555
         Width           =   540
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   105
      TabIndex        =   24
      Top             =   285
      Width           =   5730
   End
   Begin VB.Frame Frame3 
      Height          =   3315
      Left            =   105
      TabIndex        =   11
      Top             =   3900
      Width           =   5730
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -285
         Top             =   1380
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetalleagrupadores.frx":1EBC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetalleagrupadores.frx":2796
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetalleagrupadores.frx":3070
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetalleagrupadores.frx":360C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetalleagrupadores.frx":3EE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetalleagrupadores.frx":47C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetalleagrupadores.frx":509A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetalleagrupadores.frx":53B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetalleagrupadores.frx":56CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetalleagrupadores.frx":5C6A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList icono_encabezado 
         Left            =   795
         Top             =   1035
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   12
         ImageHeight     =   12
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetalleagrupadores.frx":5F84
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetalleagrupadores.frx":685E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_detalle_agrupadores 
         Height          =   3120
         Index           =   4
         Left            =   105
         TabIndex        =   34
         Top             =   645
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   5503
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre del Subtipo de Producto"
            Object.Width           =   6879
         EndProperty
      End
      Begin MSComctlLib.ListView lv_detalle_agrupadores 
         Height          =   3120
         Index           =   0
         Left            =   75
         TabIndex        =   25
         Top             =   165
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   5503
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre "
            Object.Width           =   6350
         EndProperty
      End
      Begin MSComctlLib.ListView lv_detalle_agrupadores 
         Height          =   3120
         Index           =   1
         Left            =   60
         TabIndex        =   31
         Top             =   165
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   5503
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre de la Linea"
            Object.Width           =   6879
         EndProperty
      End
      Begin MSComctlLib.ListView lv_detalle_agrupadores 
         Height          =   3120
         Index           =   2
         Left            =   45
         TabIndex        =   32
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   5503
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "Nombre de la Sublinea"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clave Linea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nombre de la Linea"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lv_detalle_agrupadores 
         Height          =   3120
         Index           =   3
         Left            =   105
         TabIndex        =   33
         Top             =   180
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   5503
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre del Tipo de Producto"
            Object.Width           =   6879
         EndProperty
      End
   End
End
Attribute VB_Name = "frmdetalleagrupadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim vartipoagrupador As Integer
Dim var_tipo_lista As Integer

Private Sub cmd_aceptar_Click()
      If var_tipo_lista = 1 Then
         Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
         n = lv_lista.ListItems.Count
         For i = 1 To n
             lv_lista.ListItems.Item(i).Selected = True
             If lv_lista.selectedItem.SubItems(2) = "*" Then
                rs.Open "SELECT * FROM TB_DETALLE_AGRUPADORES WHERE VCHA_AGR_AGRUPADOR_ID = '" + varagrupador + "' AND INTE_DEA_TIPO = " + CStr(vartipoagrupador) + " AND VCHA_ART_ARTICULO_ID = '" + Trim(lv_lista.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                If rs.EOF Then
                   var_modifica_registro = False
                   ok = TB_DETALLE_AGRUPADORES.Anadir(varagrupador, vartipoagrupador, lv_lista.selectedItem, " ", " ", " ", " ")
                   If ok Then
                      Set list_item = lv_detalle_agrupadores(0).ListItems.Add(, , lv_lista.selectedItem)
                      list_item.SubItems(1) = Trim(lv_lista.selectedItem.SubItems(1))
                      list_item.EnsureVisible
                      list_item.Selected = True
                      txt_Articulo.Enabled = False
                      txt_registros = lv_detalle_agrupadores(0).ListItems.Count
                      var_modifica_registro = True
                   Else
                      MsgBox "No se puede grabar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
                   End If
                End If
                rs.Close
             End If
         Next i
      End If
      If var_tipo_lista = 2 Then
         Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
         n = lv_lista.ListItems.Count
         For i = 1 To n
             lv_lista.ListItems.Item(i).Selected = True
             If lv_lista.selectedItem.SubItems(2) = "*" Then
                rs.Open "SELECT * FROM TB_DETALLE_AGRUPADORES WHERE VCHA_AGR_AGRUPADOR_ID = '" + varagrupador + "' AND INTE_DEA_TIPO = " + CStr(vartipoagrupador) + " AND VCHA_LIN_LINEA_ID = '" + Trim(lv_lista.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                If rs.EOF Then
                   var_modifica_registro = False
                   ok = TB_DETALLE_AGRUPADORES.Anadir(varagrupador, vartipoagrupador, " ", lv_lista.selectedItem, " ", " ", " ")
                   If ok Then
                      Set list_item = lv_detalle_agrupadores(1).ListItems.Add(, , lv_lista.selectedItem)
                      list_item.SubItems(1) = Trim(lv_lista.selectedItem.SubItems(1))
                      list_item.EnsureVisible
                      list_item.Selected = True
                      txt_Articulo.Enabled = False
                      txt_registros = lv_detalle_agrupadores(1).ListItems.Count
                      var_modifica_registro = True
                   Else
                      MsgBox "No se puede grabar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
                   End If
                End If
                rs.Close
             End If
         Next i
      End If
      If var_tipo_lista = 3 Then
         Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
         n = lv_lista.ListItems.Count
         For i = 1 To n
             lv_lista.ListItems.Item(i).Selected = True
             If lv_lista.selectedItem.SubItems(2) = "*" Then
                rs.Open "SELECT * FROM TB_DETALLE_AGRUPADORES WHERE VCHA_AGR_AGRUPADOR_ID = '" + varagrupador + "' AND INTE_DEA_TIPO = " + CStr(vartipoagrupador) + " AND VCHA_SLI_SUBLINEA_ID = '" + Trim(lv_lista.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                If rs.EOF Then
                   var_modifica_registro = False
                   ok = TB_DETALLE_AGRUPADORES.Anadir(varagrupador, vartipoagrupador, " ", " ", lv_lista.selectedItem, " ", " ")
                   If ok Then
                      Set list_item = lv_detalle_agrupadores(2).ListItems.Add(, , lv_lista.selectedItem)
                      list_item.SubItems(1) = Trim(lv_lista.selectedItem.SubItems(1))
                      list_item.EnsureVisible
                      list_item.Selected = True
                      txt_Articulo.Enabled = False
                      txt_registros = lv_detalle_agrupadores(2).ListItems.Count
                      var_modifica_registro = True
                   Else
                      MsgBox "No se puede grabar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
                   End If
                End If
                rs.Close
             End If
         Next i
      End If
      If var_tipo_lista = 4 Then
         Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
         n = lv_lista.ListItems.Count
         For i = 1 To n
             lv_lista.ListItems.Item(i).Selected = True
             If lv_lista.selectedItem.SubItems(2) = "*" Then
                rs.Open "SELECT * FROM TB_DETALLE_AGRUPADORES WHERE VCHA_AGR_AGRUPADOR_ID = '" + varagrupador + "' AND INTE_DEA_TIPO = " + CStr(vartipoagrupador) + " AND VCHA_PRO_PRODUCTO_ID= '" + Trim(lv_lista.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                If rs.EOF Then
                   var_modifica_registro = False
                   ok = TB_DETALLE_AGRUPADORES.Anadir(varagrupador, vartipoagrupador, " ", " ", " ", lv_lista.selectedItem, " ")
                   If ok Then
                      Set list_item = lv_detalle_agrupadores(3).ListItems.Add(, , lv_lista.selectedItem)
                      list_item.SubItems(1) = Trim(lv_lista.selectedItem.SubItems(1))
                      list_item.EnsureVisible
                      list_item.Selected = True
                      txt_Articulo.Enabled = False
                      txt_registros = lv_detalle_agrupadores(3).ListItems.Count
                      var_modifica_registro = True
                   Else
                      MsgBox "No se puede grabar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
                   End If
                End If
                rs.Close
             End If
         Next i
      End If
      If var_tipo_lista = 5 Then
         Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
         n = lv_lista.ListItems.Count
         For i = 1 To n
             lv_lista.ListItems.Item(i).Selected = True
             If lv_lista.selectedItem.SubItems(2) = "*" Then
                rs.Open "SELECT * FROM TB_DETALLE_AGRUPADORES WHERE VCHA_AGR_AGRUPADOR_ID = '" + varagrupador + "' AND INTE_DEA_TIPO = " + CStr(vartipoagrupador) + " AND VCHA_TAR_TIPO_ARTICULO_ID = '" + Trim(lv_lista.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                If rs.EOF Then
                   var_modifica_registro = False
                   ok = TB_DETALLE_AGRUPADORES.Anadir(varagrupador, vartipoagrupador, " ", " ", " ", " ", lv_lista.selectedItem)
                   If ok Then
                      Set list_item = lv_detalle_agrupadores(4).ListItems.Add(, , lv_lista.selectedItem)
                      list_item.SubItems(1) = Trim(lv_lista.selectedItem.SubItems(1))
                      list_item.EnsureVisible
                      list_item.Selected = True
                      txt_Articulo.Enabled = False
                      txt_registros = lv_detalle_agrupadores(4).ListItems.Count
                      var_modifica_registro = True
                   Else
                      MsgBox "No se puede grabar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
                   End If
                End If
                rs.Close
             End If
         Next i
      End If
      frm_lista.Visible = False
End Sub

Private Sub cmd_cancelar_Click()
   frm_lista.Visible = False
End Sub

Private Sub cmd_deshacer_Click()
   Call pro_textos
   cmd_guardar.Enabled = False
   cmd_deshacer.Enabled = False
End Sub

Private Sub cmd_deshacer_GotFocus()
   frm_lista.Visible = False
End Sub

Private Sub cmd_eliminar_Click()
   var_opcion_seguridad = 2
   var_acepta_seguridad = 1
   If var_global_permiso3 = 1 Then
      var_acepta_seguridad = 2
      If var_global_permiso4 = 1 Then
         frmpasswords2.Show 1
      Else
         frmpasswords.Show 1
      End If
   End If
   If var_acepta_seguridad = 1 Then
      Call pro_elimina_detalle_agrupadores
   End If
End Sub

Private Sub cmd_eliminar_GotFocus()
   frm_lista.Visible = False
End Sub

Private Sub cmd_guardar_Click()
   var_opcion_seguridad = 2
   var_acepta_seguridad = 1
   If var_global_permiso3 = 1 Then
      var_acepta_seguridad = 2
      If var_global_permiso4 = 1 Then
         frmpasswords2.Show 1
      Else
         frmpasswords.Show 1
      End If
   End If
   If var_acepta_seguridad = 1 Then
      Call pro_guardar_detalle_agrupadores
   End If
   cmd_guardar.Enabled = False
   cmd_deshacer.Enabled = False
End Sub

Private Sub cmd_guardar_GotFocus()
   frm_lista.Visible = False
End Sub

Private Sub cmd_imprimir_Click()
    If vector_valida_passwords(var_indice_menu) = "*" Then
       frmpasswords.Show
    Else
       Call gPrintListView(lv_detalle_agrupadores, "LISTADO DE detalle_agrupadores")
    End If
End Sub

Private Sub cmd_imprimir_GotFocus()
   frm_lista.Visible = False
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
   n = lv_lista.ListItems.Count
   For i = 1 To n
       If lv_lista.ListItems.Item(i).SubItems(2) = "*" Then
          lv_lista.ListItems.Item(i).SubItems(2) = " "
          lv_lista.ListItems.Item(i).Bold = False
          lv_lista.ListItems.Item(i).ForeColor = &H80000012
          lv_lista.ListItems.Item(i).ListSubItems(1).Bold = False
          lv_lista.ListItems.Item(i).ListSubItems(2).Bold = False
          lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
          lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
       Else
          lv_lista.ListItems.Item(i).SubItems(2) = "*"
          lv_lista.ListItems.Item(i).Bold = True
          lv_lista.ListItems.Item(i).ForeColor = &H8000&
          lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
          lv_lista.ListItems.Item(i).ListSubItems(2).Bold = True
          lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
          lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
      End If
   Next
   lv_lista.Refresh
End Sub

Private Sub cmd_marcar_Click()
   Dim i As Integer
   i = lv_lista.selectedItem.Index
   If lv_lista.selectedItem.SubItems(2) = "*" Then
      lv_lista.selectedItem.SubItems(2) = ""
      lv_lista.ListItems.Item(i).Bold = False
      lv_lista.ListItems.Item(i).ForeColor = &H80000012
      lv_lista.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_lista.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_lista.Refresh
   Else
      lv_lista.selectedItem.SubItems(2) = "*"
      lv_lista.ListItems.Item(i).Bold = True
      lv_lista.ListItems.Item(i).ForeColor = &HFF0000
      lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_lista.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_lista.Refresh
   End If
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
   n = lv_lista.ListItems.Count
   For i = 1 To n
       lv_lista.ListItems.Item(i).SubItems(2) = " "
       lv_lista.ListItems.Item(i).Bold = False
       lv_lista.ListItems.Item(i).ForeColor = &H80000012
       lv_lista.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_lista.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next
   lv_lista.Refresh
End Sub

Private Sub cmd_nuevo_Click()
   frm_lista.Visible = False
   If Me.opt_tipoagrupador(0).Value = False And Me.opt_tipoagrupador(1).Value = False And Me.opt_tipoagrupador(2).Value = False And Me.opt_tipoagrupador(3).Value = False And Me.opt_tipoagrupador(4).Value = False Then
      MsgBox "No se a seleccionado un tipo de agrupamiento", vbOKOnly, "ATENCION"
   Else
      Call pro_limpiatextos(Me)
      If Me.opt_tipoagrupador(0).Value = True Then
         txt_Articulo.Enabled = True
         txt_Articulo.SetFocus: var_modifica_registro = False
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
      End If
      If Me.opt_tipoagrupador(1).Value = True Then
         txt_linea.Enabled = True
         txt_linea.SetFocus: var_modifica_registro = False
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
      End If
      If Me.opt_tipoagrupador(2).Value = True Then
         txt_linea.Enabled = True
         txt_linea.SetFocus: var_modifica_registro = False
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
      End If
      If Me.opt_tipoagrupador(3).Value = True Then
         txt_producto.Enabled = True
         txt_producto.SetFocus: var_modifica_registro = False
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
      End If
      If Me.opt_tipoagrupador(4).Value = True Then
         txt_producto.Enabled = True
         txt_producto.SetFocus: var_modifica_registro = False
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
      End If
   End If
End Sub

Private Sub cmd_nuevo_GotFocus()
   frm_lista.Visible = False
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro = False Then
      var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
      If var_si <> 6 Then
         GoTo salir:
      End If
   Else
      If var_hubo_cambios = True Then
         var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
         If var_si <> 6 Then
            GoTo salir:
         End If
      End If
   End If
   Unload Me
   Exit Sub
salir:
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
   n = lv_lista.ListItems.Count
   For i = 1 To n
       If lv_lista.ListItems.Item(i).SubItems(2) = "*" And primera_vez = False Then
          numero_seleccionado1 = i
          primera_vez = True
       End If
       If lv_lista.ListItems.Item(i).SubItems(2) = "*" And primera_vez = True Then
          numero_seleccionado2 = i
       End If
   Next
   For i = numero_seleccionado1 To numero_seleccionado2
       lv_lista.ListItems.Item(i).SubItems(2) = "*"
       lv_lista.ListItems.Item(i).Bold = True
       lv_lista.ListItems.Item(i).ForeColor = &H8000&
       lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_lista.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
       lv_lista.Refresh
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
   n = lv_lista.ListItems.Count
   For i = 1 To n
       lv_lista.ListItems.Item(i).SubItems(2) = "*"
       lv_lista.ListItems.Item(i).Bold = True
       lv_lista.ListItems.Item(i).ForeColor = &H8000&
       lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_lista.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
    Next
    lv_lista.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 71 Then
      cmd_guardar_Click
   End If
   If Shift = 4 And KeyCode = 68 Then
      cmd_deshacer_Click
   End If
   If Shift = 4 And KeyCode = 69 Then
      cmd_eliminar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 2900
   frm_lista.Visible = False
   varagrupador = frmagrupadores.txt_agrupadores(0)
   var_modifica_registro = True
   lv_detalle_agrupadores(0).SmallIcons = ImageList1
   rs.Open "select * from TB_DETALLE_AGRUPADORES where vcha_agr_agrupador_id = '" & varagrupador & "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      vartipoagrupador = rs(1).Value
      If vartipoagrupador = 1 Then
         rs.Close
         Call pro_encabezadosView(Me, lv_detalle_agrupadores(0), False)
         Me.txt_Articulo.Enabled = False
         Me.txt_linea.Enabled = False
         Me.txt_producto.Enabled = False
         Me.txt_sublinea.Enabled = False
         Me.txt_subtipo_producto.Enabled = False
         txt_nombre_articulo.Enabled = False
         txt_nombre_linea.Enabled = False
         txt_nombre_sublinea.Enabled = False
         txt_nombre_producto.Enabled = False
         txt_nombre_subtipo_articulo.Enabled = False
         lv_detalle_agrupadores(0).Visible = True
         lv_detalle_agrupadores(1).Visible = False
         lv_detalle_agrupadores(2).Visible = False
         lv_detalle_agrupadores(3).Visible = False
         lv_detalle_agrupadores(4).Visible = False
         rs.Open "select a.vcha_art_articulo_id,b.vcha_art_nombre_español from TB_DETALLE_AGRUPADORES a, TB_ARTICULOS b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_art_articulo_id = b.vcha_art_articulo_id", cnn, adOpenDynamic, adLockOptimistic
         opt_tipoagrupador(0).Value = True
         opt_tipoagrupador(1).Value = False
         opt_tipoagrupador(2).Value = False
         opt_tipoagrupador(3).Value = False
         opt_tipoagrupador(4).Value = False
      End If
      If vartipoagrupador = 2 Then
         rs.Close
         Call pro_encabezadosView(Me, lv_detalle_agrupadores(1), False)
         Me.txt_Articulo.Enabled = False
         Me.txt_linea.Enabled = False
         Me.txt_producto.Enabled = False
         Me.txt_sublinea.Enabled = False
         Me.txt_subtipo_producto.Enabled = False
         txt_nombre_articulo.Enabled = False
         txt_nombre_linea.Enabled = False
         txt_nombre_sublinea.Enabled = False
         txt_nombre_producto.Enabled = False
         txt_nombre_subtipo_articulo.Enabled = False
         lv_detalle_agrupadores(0).Visible = False
         lv_detalle_agrupadores(1).Visible = True
         lv_detalle_agrupadores(2).Visible = False
         lv_detalle_agrupadores(3).Visible = False
         lv_detalle_agrupadores(4).Visible = False
         rs.Open "select a.vcha_lin_linea_id,b.vcha_lin_nombre from TB_DETALLE_AGRUPADORES a, TB_lineas b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_lin_linea_id = b.vcha_lin_linea_id", cnn, adOpenDynamic, adLockOptimistic
         opt_tipoagrupador(0).Value = False
         opt_tipoagrupador(1).Value = True
         opt_tipoagrupador(2).Value = False
         opt_tipoagrupador(3).Value = False
         opt_tipoagrupador(4).Value = False
      End If
      If vartipoagrupador = 3 Then
         rs.Close
         Call pro_encabezadosView(Me, lv_detalle_agrupadores(2), False)
         Me.txt_Articulo.Enabled = False
         Me.txt_linea.Enabled = False
         Me.txt_producto.Enabled = False
         Me.txt_sublinea.Enabled = False
         Me.txt_subtipo_producto.Enabled = False
         txt_nombre_articulo.Enabled = False
         txt_nombre_linea.Enabled = False
         txt_nombre_sublinea.Enabled = False
         txt_nombre_producto.Enabled = False
         txt_nombre_subtipo_articulo.Enabled = False
         lv_detalle_agrupadores(0).Visible = False
         lv_detalle_agrupadores(1).Visible = False
         lv_detalle_agrupadores(2).Visible = True
         lv_detalle_agrupadores(3).Visible = False
         lv_detalle_agrupadores(4).Visible = False
         rs.Open "select a.vcha_sli_sublinea_id,b.vcha_sli_nombre from TB_DETALLE_AGRUPADORES a, TB_sublineas b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_sli_sublinea_id = b.vcha_sli_sublinea_id", cnn, adOpenDynamic, adLockOptimistic
         opt_tipoagrupador(0).Value = False
         opt_tipoagrupador(1).Value = False
         opt_tipoagrupador(2).Value = True
         opt_tipoagrupador(3).Value = False
         opt_tipoagrupador(4).Value = False
      End If
      If vartipoagrupador = 4 Then
         rs.Close
         Call pro_encabezadosView(Me, lv_detalle_agrupadores(3), False)
         Me.txt_Articulo.Enabled = False
         Me.txt_linea.Enabled = False
         Me.txt_producto.Enabled = False
         Me.txt_sublinea.Enabled = False
         Me.txt_subtipo_producto.Enabled = False
         txt_nombre_articulo.Enabled = False
         txt_nombre_linea.Enabled = False
         txt_nombre_sublinea.Enabled = False
         txt_nombre_producto.Enabled = False
         txt_nombre_subtipo_articulo.Enabled = False
         lv_detalle_agrupadores(0).Visible = False
         lv_detalle_agrupadores(1).Visible = False
         lv_detalle_agrupadores(2).Visible = False
         lv_detalle_agrupadores(3).Visible = True
         lv_detalle_agrupadores(4).Visible = False
         rs.Open "select a.vcha_pro_producto_id,b.vcha_pro_nombre from TB_DETALLE_AGRUPADORES a, TB_productos b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_pro_producto_id = b.vcha_pro_producto_id", cnn, adOpenDynamic, adLockOptimistic
         opt_tipoagrupador(0).Value = False
         opt_tipoagrupador(1).Value = False
         opt_tipoagrupador(2).Value = False
         opt_tipoagrupador(3).Value = True
         opt_tipoagrupador(4).Value = False
      End If
      If vartipoagrupador = 5 Then
         rs.Close
         Call pro_encabezadosView(Me, lv_detalle_agrupadores(4), False)
         Me.txt_Articulo.Enabled = False
         Me.txt_linea.Enabled = False
         Me.txt_producto.Enabled = False
         Me.txt_sublinea.Enabled = False
         Me.txt_subtipo_producto.Enabled = False
         txt_nombre_articulo.Enabled = False
         txt_nombre_linea.Enabled = False
         txt_nombre_sublinea.Enabled = False
         txt_nombre_producto.Enabled = False
         txt_nombre_subtipo_articulo.Enabled = False
         lv_detalle_agrupadores(0).Visible = False
         lv_detalle_agrupadores(1).Visible = False
         lv_detalle_agrupadores(2).Visible = False
         lv_detalle_agrupadores(3).Visible = False
         lv_detalle_agrupadores(4).Visible = True
         rs.Open "select a.vcha_tar_tipo_articulo_id,b.vcha_tar_nombre from TB_DETALLE_AGRUPADORES a, TB_tipoarticulos b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_tar_tipo_articulo_id = b.vcha_tar_tipo_articulo_id", cnn, adOpenDynamic, adLockOptimistic
         opt_tipoagrupador(0).Value = False
         opt_tipoagrupador(1).Value = False
         opt_tipoagrupador(2).Value = False
         opt_tipoagrupador(3).Value = False
         opt_tipoagrupador(4).Value = True
      End If
   Else
      vartipoagrupador = 1
      If vartipoagrupador = 1 Then
         rs.Close
         Call pro_encabezadosView(Me, lv_detalle_agrupadores(0), False)
         Call pro_llena_listview1
         pro_textos
         Me.txt_Articulo.Enabled = False
         Me.txt_linea.Enabled = False
         Me.txt_producto.Enabled = False
         Me.txt_sublinea.Enabled = False
         Me.txt_subtipo_producto.Enabled = False
         txt_nombre_articulo.Enabled = False
         txt_nombre_linea.Enabled = False
         txt_nombre_sublinea.Enabled = False
         txt_nombre_producto.Enabled = False
         txt_nombre_subtipo_articulo.Enabled = False
         lv_detalle_agrupadores(0).Visible = True
         lv_detalle_agrupadores(1).Visible = False
         lv_detalle_agrupadores(2).Visible = False
         lv_detalle_agrupadores(3).Visible = False
         lv_detalle_agrupadores(4).Visible = False
         rs.Open "select a.vcha_art_articulo_id,b.vcha_art_nombre_español from TB_DETALLE_AGRUPADORES a, TB_ARTICULOS b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_art_articulo_id = b.vcha_art_articulo_id", cnn, adOpenDynamic, adLockOptimistic
         opt_tipoagrupador(0).Value = True
         opt_tipoagrupador(1).Value = False
         opt_tipoagrupador(2).Value = False
         opt_tipoagrupador(3).Value = False
         opt_tipoagrupador(4).Value = False
      End If
   End If
   cmd_guardar.Enabled = False
   cmd_deshacer.Enabled = False
   rs.Close
   Call pro_llena_listview1
   pro_textos
   Me.opt_tipoagrupador(0).Value = False
   Me.opt_tipoagrupador(1).Value = False
   Me.opt_tipoagrupador(2).Value = False
   Me.opt_tipoagrupador(3).Value = False
   Me.opt_tipoagrupador(4).Value = False
   Me.txt_Articulo.Enabled = False
   Me.txt_linea.Enabled = False
   Me.txt_nombre_articulo.Enabled = False
   Me.txt_nombre_linea.Enabled = False
   Me.txt_nombre_producto.Enabled = False
   Me.txt_nombre_sublinea.Enabled = False
   Me.txt_nombre_subtipo_articulo.Enabled = False
   Me.txt_producto.Enabled = False
   Me.txt_sublinea.Enabled = False
   Me.txt_subtipo_producto.Enabled = False
   Me.txt_Articulo = ""
   Me.txt_linea = ""
   Me.txt_nombre_articulo = ""
   Me.txt_nombre_linea = ""
   Me.txt_nombre_producto = ""
   Me.txt_nombre_sublinea = ""
   Me.txt_nombre_subtipo_articulo = ""
   Me.txt_producto = ""
   Me.txt_sublinea = ""
   Me.txt_subtipo_producto = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_detalleagrupadores)
End Sub

Private Sub lv_detalle_agrupadores_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_detalle_agrupadores, ColumnHeader)
End Sub

Private Sub lv_detalle_agrupadores_GotFocus(Index As Integer)
   frm_lista.Visible = False
   Me.txt_Articulo = Me.lv_detalle_agrupadores(0).selectedItem
   Me.txt_linea = ""
   Me.txt_nombre_articulo = Me.lv_detalle_agrupadores(0).selectedItem.SubItems(1)
   Me.txt_nombre_linea = ""
   Me.txt_nombre_producto = ""
   Me.txt_nombre_sublinea = ""
   Me.txt_nombre_subtipo_articulo = ""
   Me.txt_producto = ""
   Me.txt_sublinea = ""
   Me.txt_subtipo_producto = ""
End Sub

Private Sub lv_detalle_agrupadores_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
   If opt_tipoagrupador(0).Value = True Then
      Set lv_detalle_agrupadores(0).selectedItem = Item
      var_modifica_registro = True
      txt_Articulo.Enabled = False
      txt_Articulo = lv_detalle_agrupadores(0).selectedItem
      txt_nombre_articulo = lv_detalle_agrupadores(0).selectedItem.SubItems(1)
   End If
   If opt_tipoagrupador(1).Value = True Then
      Set lv_detalle_agrupadores(1).selectedItem = Item
      var_modifica_registro = True
      txt_linea.Enabled = False
      txt_linea = lv_detalle_agrupadores(1).selectedItem
      txt_nombre_linea = lv_detalle_agrupadores(1).selectedItem.SubItems(1)
   End If
   If opt_tipoagrupador(2).Value = True Then
      Set lv_detalle_agrupadores(2).selectedItem = Item
      var_modifica_registro = True
      txt_sublinea.Enabled = False
      Me.txt_sublinea = lv_detalle_agrupadores(2).selectedItem
      Me.txt_nombre_sublinea = lv_detalle_agrupadores(2).selectedItem.SubItems(1)
   End If
   If opt_tipoagrupador(3).Value = True Then
      Set lv_detalle_agrupadores(3).selectedItem = Item
      var_modifica_registro = True
      txt_producto.Enabled = False
      txt_producto = lv_detalle_agrupadores(1).selectedItem
      txt_nombre_producto = lv_detalle_agrupadores(1).selectedItem.SubItems(1)
   End If
   If opt_tipoagrupador(4).Value = True Then
      Set lv_detalle_agrupadores(4).selectedItem = Item
      var_modifica_registro = True
      txt_subtipo_producto.Enabled = False
      Me.txt_subtipo_producto = lv_detalle_agrupadores(1).selectedItem
      Me.txt_nombre_subtipo_articulo = lv_detalle_agrupadores(1).selectedItem.SubItems(1)
   End If
   pro_textos
End Sub


Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim i As Integer
      i = lv_lista.selectedItem.Index
      If lv_lista.selectedItem.SubItems(2) = "*" Then
         lv_lista.selectedItem.SubItems(2) = ""
         lv_lista.ListItems.Item(i).Bold = False
         lv_lista.ListItems.Item(i).ForeColor = &H80000012
         lv_lista.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_lista.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_lista.Refresh
      Else
         lv_lista.selectedItem.SubItems(2) = "*"
         lv_lista.ListItems.Item(i).Bold = True
         lv_lista.ListItems.Item(i).ForeColor = &HFF0000
         lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_lista.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_lista.Refresh
      End If
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub opt_tipoagrupador_Click(Index As Integer)
   If opt_tipoagrupador(0).Value = True Then
      lv_detalle_agrupadores(0).Visible = True
      lv_detalle_agrupadores(1).Visible = False
      lv_detalle_agrupadores(2).Visible = False
      lv_detalle_agrupadores(3).Visible = False
      lv_detalle_agrupadores(4).Visible = False
      Me.txt_Articulo.Enabled = True
      Me.txt_nombre_articulo.Enabled = True
      Me.txt_linea.Enabled = False
      Me.txt_nombre_linea.Enabled = False
      Me.txt_sublinea.Enabled = False
      Me.txt_nombre_sublinea.Enabled = False
      Me.txt_producto.Enabled = False
      Me.txt_nombre_producto.Enabled = False
      Me.txt_subtipo_producto.Enabled = False
      Me.txt_nombre_subtipo_articulo.Enabled = False
      pro_textos
      vartipoagrupador = 1
   End If
   If opt_tipoagrupador(1).Value = True Then
      lv_detalle_agrupadores(0).Visible = False
      lv_detalle_agrupadores(1).Visible = True
      lv_detalle_agrupadores(2).Visible = False
      lv_detalle_agrupadores(3).Visible = False
      lv_detalle_agrupadores(4).Visible = False
      Me.txt_Articulo.Enabled = False
      Me.txt_nombre_articulo.Enabled = False
      Me.txt_linea.Enabled = True
      Me.txt_nombre_linea.Enabled = True
      Me.txt_sublinea.Enabled = False
      Me.txt_nombre_sublinea.Enabled = False
      Me.txt_producto.Enabled = False
      Me.txt_nombre_producto.Enabled = False
      Me.txt_subtipo_producto.Enabled = False
      Me.txt_nombre_subtipo_articulo.Enabled = False
      vartipoagrupador = 2
      pro_textos
   End If
   If opt_tipoagrupador(2).Value = True Then
      lv_detalle_agrupadores(0).Visible = False
      lv_detalle_agrupadores(1).Visible = False
      lv_detalle_agrupadores(2).Visible = True
      lv_detalle_agrupadores(3).Visible = False
      lv_detalle_agrupadores(4).Visible = False
      Me.txt_Articulo.Enabled = False
      Me.txt_nombre_articulo.Enabled = False
      Me.txt_linea.Enabled = True
      Me.txt_nombre_linea.Enabled = True
      Me.txt_sublinea.Enabled = True
      Me.txt_nombre_sublinea.Enabled = True
      Me.txt_producto.Enabled = False
      Me.txt_nombre_producto.Enabled = False
      Me.txt_subtipo_producto.Enabled = False
      Me.txt_nombre_subtipo_articulo.Enabled = False
      vartipoagrupador = 3
      pro_textos
   End If
   If opt_tipoagrupador(3).Value = True Then
      lv_detalle_agrupadores(0).Visible = False
      lv_detalle_agrupadores(1).Visible = False
      lv_detalle_agrupadores(2).Visible = False
      lv_detalle_agrupadores(3).Visible = True
      lv_detalle_agrupadores(4).Visible = False
      Me.txt_Articulo.Enabled = False
      Me.txt_nombre_articulo.Enabled = False
      Me.txt_linea.Enabled = False
      Me.txt_nombre_linea.Enabled = False
      Me.txt_sublinea.Enabled = False
      Me.txt_nombre_sublinea.Enabled = False
      Me.txt_producto.Enabled = True
      Me.txt_nombre_producto.Enabled = True
      Me.txt_subtipo_producto.Enabled = False
      Me.txt_nombre_subtipo_articulo.Enabled = False
      vartipoagrupador = 4
      pro_textos
   End If
   If opt_tipoagrupador(4).Value = True Then
      lv_detalle_agrupadores(0).Visible = False
      lv_detalle_agrupadores(1).Visible = False
      lv_detalle_agrupadores(2).Visible = False
      lv_detalle_agrupadores(3).Visible = False
      lv_detalle_agrupadores(4).Visible = True
      Me.txt_Articulo.Enabled = False
      Me.txt_nombre_articulo.Enabled = False
      Me.txt_linea.Enabled = False
      Me.txt_nombre_linea.Enabled = False
      Me.txt_sublinea.Enabled = False
      Me.txt_nombre_sublinea.Enabled = False
      Me.txt_producto.Enabled = True
      Me.txt_nombre_producto.Enabled = True
      Me.txt_subtipo_producto.Enabled = True
      Me.txt_nombre_subtipo_articulo.Enabled = True
      vartipoagrupador = 5
      pro_textos
   End If
End Sub


Sub pro_guardar_detalle_agrupadores()
   Dim ok As Boolean
   If opt_tipoagrupador(0).Value = True Then
      vartipoagrupador = 1
   End If
   If opt_tipoagrupador(1).Value = True Then
      vartipoagrupador = 2
   End If
   If opt_tipoagrupador(2).Value = True Then
      vartipoagrupador = 3
   End If
   If opt_tipoagrupador(3).Value = True Then
      vartipoagrupador = 4
   End If
   If opt_tipoagrupador(4).Value = True Then
      vartipoagrupador = 5
   End If
   If vartipoagrupador = 1 Then
      Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
      ok = True
      If txt_Articulo <> "" Then
         If var_hubo_cambios Then
            ok = TB_DETALLE_AGRUPADORES.Anadir(varagrupador, vartipoagrupador, txt_Articulo, " ", " ", " ", " ")
            If ok Then
               pro_actualiza_ListView
               txt_Articulo.Enabled = False
               MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
               txt_registros = lv_detalle_agrupadores(0).ListItems.Count
               var_modifica_registro = True
            Else
               MsgBox "No se puede grabar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
         End If
      End If
   End If
   If vartipoagrupador = 2 Then
      Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
      ok = True
      If txt_linea <> "" Then
         If var_hubo_cambios Then
            ok = TB_DETALLE_AGRUPADORES.Anadir(varagrupador, vartipoagrupador, " ", txt_linea.Text, " ", " ", " ")
            If ok Then
               pro_actualiza_ListView
               txt_linea.Enabled = False
               MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
               txt_registros = lv_detalle_agrupadores(1).ListItems.Count
               var_modifica_registro = True
            Else
               MsgBox "No se puede grabar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
         End If
      End If
   End If
   If vartipoagrupador = 3 Then
      Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
      ok = True
      If txt_linea <> "" And txt_sublinea <> "" Then
         If var_hubo_cambios Then
            ok = TB_DETALLE_AGRUPADORES.Anadir(varagrupador, vartipoagrupador, " ", txt_linea, txt_sublinea, " ", " ")
            If ok Then
               pro_actualiza_ListView
               txt_sublinea.Enabled = False
               MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
               txt_registros = lv_detalle_agrupadores(2).ListItems.Count
               var_modifica_registro = True
            Else
               MsgBox "No se puede grabar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
         End If
      End If
   End If
   If vartipoagrupador = 4 Then
      Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
      ok = True
      If txt_producto <> "" Then
         If var_hubo_cambios Then
            ok = TB_DETALLE_AGRUPADORES.Anadir(varagrupador, vartipoagrupador, " ", " ", " ", txt_producto, " ")
            If ok Then
               pro_actualiza_ListView
               txt_producto.Enabled = False
               MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
               txt_registros = lv_detalle_agrupadores(3).ListItems.Count
               var_modifica_registro = True
            Else
               MsgBox "No se puede grabar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
         End If
      End If
   End If
   If vartipoagrupador = 5 Then
      Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
      ok = True
      If txt_producto <> "" And txt_subtipo_producto <> "" Then
         If var_hubo_cambios Then
            ok = TB_DETALLE_AGRUPADORES.Anadir(varagrupador, vartipoagrupador, " ", " ", " ", txt_producto, txt_subtipo_producto)
            If ok Then
               pro_actualiza_ListView
               txt_subtipo_producto.Enabled = False
               MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
               txt_registros = lv_detalle_agrupadores(4).ListItems.Count
               var_modifica_registro = True
            Else
               MsgBox "No se puede grabar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
         End If
      End If
   End If
Set TB_DETALLE_AGRUPADORES = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_detalle_agrupadores()
   Dim var_llave_usuarios As String
   Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
   Dim ok As Boolean
   If opt_tipoagrupador(0).Value = True Then
      vartipoagrupador = 1
   End If
   If opt_tipoagrupador(1).Value = True Then
      vartipoagrupador = 2
   End If
   If opt_tipoagrupador(2).Value = True Then
      vartipoagrupador = 3
   End If
   If opt_tipoagrupador(3).Value = True Then
      vartipoagrupador = 4
   End If
   If opt_tipoagrupador(4).Value = True Then
      vartipoagrupador = 5
   End If
   On Error GoTo salir:
   ok = True
   If vartipoagrupador = 1 Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         rs.Open "delete from TB_DETALLE_AGRUPADORES where VCHA_AGR_AGRUPADOR_ID = '" + varagrupador + "' and INTE_DEA_TIPO  = '1' and VCHA_ART_ARTICULO_ID  = '" + txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
      Else
         GoTo salir:
      End If
      If ok Then
        MsgBox "Se Elimino Correctamente el Registro", vbInformation
        lv_detalle_agrupadores(0).ListItems.Remove (lv_detalle_agrupadores(0).selectedItem.Index)
        Call pro_limpiatextos(Me)
        txt_registros = lv_detalle_agrupadores(0).ListItems.Count
        lv_detalle_agrupadores(0).selectedItem.Selected = True
        pro_textos
      Else
        MsgBox "No se puede eliminar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If

   If vartipoagrupador = 2 Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         rs.Open "delete from TB_DETALLE_AGRUPADORES where VCHA_AGR_AGRUPADOR_ID = '" + varagrupador + "' and INTE_DEA_TIPO  = '2' and VCHA_LIN_LINEA_ID  = '" + txt_linea + "'", cnn, adOpenDynamic, adLockOptimistic
      Else
         GoTo salir:
      End If
      If ok Then
        MsgBox "Se Elimino Correctamente el Registro", vbInformation
        lv_detalle_agrupadores(1).ListItems.Remove (lv_detalle_agrupadores(1).selectedItem.Index)
        Call pro_limpiatextos(Me)
        txt_registros = lv_detalle_agrupadores(1).ListItems.Count
        lv_detalle_agrupadores(1).selectedItem.Selected = True
        pro_textos
      Else
        MsgBox "No se puede eliminar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If


   If vartipoagrupador = 3 Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         rs.Open "delete from TB_DETALLE_AGRUPADORES where VCHA_AGR_AGRUPADOR_ID = '" + varagrupador + "' and INTE_DEA_TIPO  = '3' and VCHA_LIN_LINEA_ID  = '" + txt_linea + "' and VCHA_SLI_SUBLINEA_ID = '" + txt_sublinea + "'", cnn, adOpenDynamic, adLockOptimistic
      Else
         GoTo salir:
      End If
      If ok Then
        MsgBox "Se Elimino Correctamente el Registro", vbInformation
        lv_detalle_agrupadores(2).ListItems.Remove (lv_detalle_agrupadores(2).selectedItem.Index)
        Call pro_limpiatextos(Me)
        txt_registros = lv_detalle_agrupadores(2).ListItems.Count
        lv_detalle_agrupadores(2).selectedItem.Selected = True
        pro_textos
      Else
        MsgBox "No se puede eliminar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If

   If vartipoagrupador = 4 Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         rs.Open "delete from TB_DETALLE_AGRUPADORES where VCHA_AGR_AGRUPADOR_ID = '" + varagrupador + "' and INTE_DEA_TIPO  = '4' and VCHA_PRO_PRODUCTO_ID  = '" + txt_producto + "'", cnn, adOpenDynamic, adLockOptimistic
      Else
         GoTo salir:
      End If
      If ok Then
        MsgBox "Se Elimino Correctamente el Registro", vbInformation
        lv_detalle_agrupadores(3).ListItems.Remove (lv_detalle_agrupadores(3).selectedItem.Index)
        Call pro_limpiatextos(Me)
        txt_registros = lv_detalle_agrupadores(3).ListItems.Count
        lv_detalle_agrupadores(3).selectedItem.Selected = True
        pro_textos
      Else
        MsgBox "No se puede eliminar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If

   If vartipoagrupador = 5 Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         rs.Open "delete from TB_DETALLE_AGRUPADORES where VCHA_AGR_AGRUPADOR_ID = '" + varagrupador + "' and INTE_DEA_TIPO  = '5' and VCHA_PRO_PRODUCTO_ID  = '" + txt_producto + "' and VCHA_TAR_TIPO_ARTICULO_ID = '" + txt_subtipo_producto + "'", cnn, adOpenDynamic, adLockOptimistic
      Else
         GoTo salir:
      End If
      If ok Then
        MsgBox "Se Elimino Correctamente el Registro", vbInformation
        lv_detalle_agrupadores(4).ListItems.Remove (lv_detalle_agrupadores(4).selectedItem.Index)
        Call pro_limpiatextos(Me)
        txt_registros = lv_detalle_agrupadores(4).ListItems.Count
        lv_detalle_agrupadores(4).selectedItem.Selected = True
        pro_textos
      Else
        MsgBox "No se puede eliminar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If


salir:
   Set TB_DETALLE_AGRUPADORES = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rsaux2.Open "select distinct inte_dea_tipo from tb_detalle_agrupadores where VCHA_AGR_AGRUPADOR_ID = '" + varagrupador + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux2.EOF Then
      While Not rsaux2.EOF
         vartipoagrupador = rsaux2(0).Value
         If vartipoagrupador = 1 Then
            lv_detalle_agrupadores(0).Visible = True
            lv_detalle_agrupadores(1).Visible = False
            lv_detalle_agrupadores(2).Visible = False
            lv_detalle_agrupadores(3).Visible = False
            lv_detalle_agrupadores(4).Visible = False
            rs.Open "select a.vcha_art_articulo_id,b.vcha_art_nombre_español from TB_DETALLE_AGRUPADORES a, TB_ARTICULOS b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_art_articulo_id = b.vcha_art_articulo_id", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
               Set list_item = lv_detalle_agrupadores(0).ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               rs.MoveNext:
            Wend
            rs.Close
            lv_detalle_agrupadores(0).Visible = True
            lv_detalle_agrupadores(1).Visible = False
            lv_detalle_agrupadores(2).Visible = False
            lv_detalle_agrupadores(3).Visible = False
            lv_detalle_agrupadores(4).Visible = False
            opt_tipoagrupador(0).Value = True
            opt_tipoagrupador(1).Value = False
            opt_tipoagrupador(2).Value = False
            opt_tipoagrupador(3).Value = False
            opt_tipoagrupador(4).Value = False
            pro_textos
         End If
         If vartipoagrupador = 2 Then
            lv_detalle_agrupadores(0).Visible = False
            lv_detalle_agrupadores(1).Visible = True
            lv_detalle_agrupadores(2).Visible = False
            lv_detalle_agrupadores(3).Visible = False
            lv_detalle_agrupadores(4).Visible = False
            rs.Open "select distinct a.vcha_lin_linea_id,b.vcha_lin_nombre from TB_DETALLE_AGRUPADORES a, TB_lineas b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_lin_linea_id = b.vcha_lin_linea_id", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
               Set list_item = lv_detalle_agrupadores(1).ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               rs.MoveNext:
            Wend
            rs.Close
            lv_detalle_agrupadores(0).Visible = False
            lv_detalle_agrupadores(1).Visible = True
            lv_detalle_agrupadores(2).Visible = False
            lv_detalle_agrupadores(3).Visible = False
            lv_detalle_agrupadores(4).Visible = False
            opt_tipoagrupador(0).Value = False
            opt_tipoagrupador(1).Value = True
            opt_tipoagrupador(2).Value = False
            opt_tipoagrupador(3).Value = False
            opt_tipoagrupador(4).Value = False
            pro_textos
         End If
         If vartipoagrupador = 3 Then
            lv_detalle_agrupadores(0).Visible = False
            lv_detalle_agrupadores(1).Visible = False
            lv_detalle_agrupadores(2).Visible = True
            lv_detalle_agrupadores(3).Visible = False
            lv_detalle_agrupadores(4).Visible = False
            rs.Open "select distinct a.vcha_sli_sublinea_id,b.vcha_sli_nombre,a.vcha_lin_linea_id,c.vcha_lin_nombre from TB_DETALLE_AGRUPADORES a, TB_sublineas b, tb_lineas c where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_sli_sublinea_id = b.vcha_sli_sublinea_id and a.vcha_lin_linea_id = c.vcha_lin_linea_id", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
               Set list_item = lv_detalle_agrupadores(2).ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
               list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
               rs.MoveNext:
            Wend
            rs.Close
            lv_detalle_agrupadores(0).Visible = False
            lv_detalle_agrupadores(1).Visible = False
            lv_detalle_agrupadores(2).Visible = True
            lv_detalle_agrupadores(3).Visible = False
            lv_detalle_agrupadores(4).Visible = False
            opt_tipoagrupador(0).Value = False
            opt_tipoagrupador(1).Value = False
            opt_tipoagrupador(2).Value = True
            opt_tipoagrupador(3).Value = False
            opt_tipoagrupador(4).Value = False
            pro_textos
         End If
         If vartipoagrupador = 4 Then
            lv_detalle_agrupadores(0).Visible = False
            lv_detalle_agrupadores(1).Visible = False
            lv_detalle_agrupadores(2).Visible = False
            lv_detalle_agrupadores(3).Visible = True
            lv_detalle_agrupadores(4).Visible = False
            rs.Open "select distinct a.vcha_pro_producto_id,b.vcha_pro_nombre from TB_DETALLE_AGRUPADORES a, TB_productos b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_pro_producto_id = b.vcha_pro_producto_id", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
               Set list_item = lv_detalle_agrupadores(3).ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               rs.MoveNext:
            Wend
            rs.Close
            lv_detalle_agrupadores(0).Visible = False
            lv_detalle_agrupadores(1).Visible = False
            lv_detalle_agrupadores(2).Visible = False
            lv_detalle_agrupadores(3).Visible = True
            lv_detalle_agrupadores(4).Visible = False
            opt_tipoagrupador(0).Value = False
            opt_tipoagrupador(1).Value = False
            opt_tipoagrupador(2).Value = False
            opt_tipoagrupador(3).Value = True
            opt_tipoagrupador(4).Value = False
            pro_textos
         End If
         If vartipoagrupador = 5 Then
            lv_detalle_agrupadores(0).Visible = False
            lv_detalle_agrupadores(1).Visible = False
            lv_detalle_agrupadores(2).Visible = False
            lv_detalle_agrupadores(3).Visible = False
            lv_detalle_agrupadores(4).Visible = True
            rs.Open "select distinct a.vcha_tar_tipo_articulo_id,b.vcha_tar_nombre from TB_DETALLE_AGRUPADORES a, TB_tipoarticulos b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_tar_tipo_articulo_id = b.vcha_tar_tipo_articulo_id", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
               Set list_item = lv_detalle_agrupadores(4).ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               rs.MoveNext:
            Wend
            rs.Close
            lv_detalle_agrupadores(0).Visible = False
            lv_detalle_agrupadores(1).Visible = False
            lv_detalle_agrupadores(2).Visible = False
            lv_detalle_agrupadores(3).Visible = False
            lv_detalle_agrupadores(4).Visible = True
            opt_tipoagrupador(0).Value = False
            opt_tipoagrupador(1).Value = False
            opt_tipoagrupador(2).Value = False
            opt_tipoagrupador(3).Value = False
            opt_tipoagrupador(4).Value = True
            pro_textos
         End If
         rsaux2.MoveNext:
         Wend
      Else
         vartipoagrupador = 1
      End If
      rsaux2.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
   If opt_tipoagrupador(0).Value = True Then
      var_n = lv_detalle_agrupadores(0).ListItems.Count
      If var_n > 0 Then
         txt_Articulo = lv_detalle_agrupadores(0).selectedItem
         rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_nombre_articulo = IIf(IsNull(rs!VCHA_ART_NOMBRE_ESPAÑOL), "", rs!VCHA_ART_NOMBRE_ESPAÑOL)
         Else
            txt_nombre_articulo = ""
         End If
         rs.Close
         opt_tipoagrupador(0).Value = True
         opt_tipoagrupador(1).Value = False
         opt_tipoagrupador(2).Value = False
         opt_tipoagrupador(3).Value = False
         opt_tipoagrupador(4).Value = False
      End If
      var_numero_renglones = lv_detalle_agrupadores(0).Height / 312.5
      If var_n > var_numero_renglones Then
         lv_detalle_agrupadores(0).ColumnHeaders(2).Width = 3850
      Else
         lv_detalle_agrupadores(0).ColumnHeaders(2).Width = 4099.9
      End If
   End If
   If opt_tipoagrupador(1).Value = True Then
      var_n = lv_detalle_agrupadores(1).ListItems.Count
      If var_n > 0 Then
         txt_linea = lv_detalle_agrupadores(1).selectedItem
         txt_nombre_linea = Obtener_llave(cnn, rs, "TB_lineas", "VCHA_lin_linea_ID", txt_linea, 1, "T")
         opt_tipoagrupador(0).Value = False
         opt_tipoagrupador(1).Value = True
         opt_tipoagrupador(2).Value = False
         opt_tipoagrupador(3).Value = False
         opt_tipoagrupador(4).Value = False
      End If
      var_numero_renglones = lv_detalle_agrupadores(1).Height / 312.5
      If var_n > var_numero_renglones Then
         lv_detalle_agrupadores(1).ColumnHeaders(2).Width = 3850
      Else
         lv_detalle_agrupadores(1).ColumnHeaders(2).Width = 4099.9
      End If
   End If
   If opt_tipoagrupador(2).Value = True Then
      var_n = lv_detalle_agrupadores(2).ListItems.Count
      If var_n > 0 Then
         txt_linea = lv_detalle_agrupadores(2).selectedItem.SubItems(2)
         txt_nombre_linea = lv_detalle_agrupadores(2).selectedItem.SubItems(3)
         txt_sublinea = lv_detalle_agrupadores(2).selectedItem
         txt_nombre_sublinea = lv_detalle_agrupadores(2).selectedItem.SubItems(1)
         opt_tipoagrupador(0).Value = False
         opt_tipoagrupador(1).Value = False
         opt_tipoagrupador(2).Value = True
         opt_tipoagrupador(3).Value = False
         opt_tipoagrupador(4).Value = False
      End If
      var_numero_renglones = lv_detalle_agrupadores(2).Height / 312.5
      If var_n > var_numero_renglones Then
         lv_detalle_agrupadores(2).ColumnHeaders(2).Width = 3899.9
      Else
         lv_detalle_agrupadores(2).ColumnHeaders(2).Width = 4099.9
      End If
   End If
   If opt_tipoagrupador(3).Value = True Then
      var_n = lv_detalle_agrupadores(3).ListItems.Count
      If var_n > 0 Then
         txt_producto = lv_detalle_agrupadores(3).selectedItem
         cmb_productos = Obtener_llave(cnn, rs, "TB_productos", "VCHA_pro_producto_ID", txt_producto, 1, "T")
         opt_tipoagrupador(0).Value = False
         opt_tipoagrupador(1).Value = False
         opt_tipoagrupador(2).Value = False
         opt_tipoagrupador(3).Value = True
         opt_tipoagrupador(4).Value = False
      End If
      var_numero_renglones = lv_detalle_agrupadores(3).Height / 312.5
      If var_n > var_numero_renglones Then
         lv_detalle_agrupadores(3).ColumnHeaders(2).Width = 3850
      Else
         lv_detalle_agrupadores(3).ColumnHeaders(2).Width = 4099.9
      End If
  End If
   If opt_tipoagrupador(4).Value = True Then
      var_n = lv_detalle_agrupadores(3).ListItems.Count
      If var_n > 0 Then
         txt_producto = lv_detalle_agrupadores(3).selectedItem
         cmb_productos = Obtener_llave(cnn, rs, "TB_productos", "VCHA_pro_producto_ID", txt_producto, 1, "T")
         txt_subtipo_producto = lv_detalle_agrupadores(4).selectedItem
         cmb_subtipos_producto = Obtener_llave(cnn, rs, "TB_tipoarticulos", "VCHA_tar_tipo_articulo_ID", txt_subtipo_producto, 1, "T")
         opt_tipoagrupador(0).Value = False
         opt_tipoagrupador(1).Value = False
         opt_tipoagrupador(2).Value = False
         opt_tipoagrupador(3).Value = False
         opt_tipoagrupador(4).Value = True
      End If
      var_numero_renglones = lv_detalle_agrupadores(4).Height / 312.5
      If var_n > var_numero_renglones Then
         lv_detalle_agrupadores(4).ColumnHeaders(2).Width = 3850
      Else
         lv_detalle_agrupadores(4).ColumnHeaders(2).Width = 4099.9
      End If
   End If
   var_modifica_registro = True
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
   If vartipoagrupador = 1 Then
      If var_modifica_registro = False Then
         Set list_item = lv_detalle_agrupadores(0).ListItems.Add(, , txt_Articulo)
         list_item.SubItems(1) = txt_nombre_articulo
         list_item.EnsureVisible
         list_item.Selected = True
      Else
         lv_detalle_agrupadores(0).ListItems.Item(lv_detalle_agrupadores(0).selectedItem.Index).Checked = False
         lv_detalle_agrupadores(0).ListItems.Item(lv_detalle_agrupadores(0).selectedItem.Index) = txt_Articulo
         lv_detalle_agrupadores(0).ListItems.Item(lv_detalle_agrupadores(0).selectedItem.Index).ListSubItems(1) = txt_nombre_articulo
         lv_detalle_agrupadores(0).ListItems.Item(lv_detalle_agrupadores(0).selectedItem.Index).Selected = True
      End If
   End If
   If vartipoagrupador = 2 Then
      If var_modifica_registro = False Then
         Set list_item = lv_detalle_agrupadores(1).ListItems.Add(, , txt_linea)
         list_item.SubItems(1) = vardetallearticulo
         list_item.EnsureVisible
         list_item.Selected = True
      Else
         lv_detalle_agrupadores(1).ListItems.Item(lv_detalle_agrupadores(1).selectedItem.Index) = txt_linea
         lv_detalle_agrupadores(1).ListItems.Item(lv_detalle_agrupadores(1).selectedItem.Index).ListSubItems(1) = vardetallearticulo
         lv_detalle_agrupadores(1).ListItems.Item(lv_detalle_agrupadores(1).selectedItem.Index).Selected = True
      End If
   End If
   If vartipoagrupador = 3 Then
      If var_modifica_registro = False Then
         Set list_item = lv_detalle_agrupadores(2).ListItems.Add(, , txt_sublinea)
         list_item.SubItems(1) = vardetallearticulo
         list_item.SubItems(2) = vardetallearticulo
         list_item.SubItems(3) = vardetallearticulo
         list_item.EnsureVisible
         list_item.Selected = True
      Else
         lv_detalle_agrupadores(2).ListItems.Item(lv_detalle_agrupadores(2).selectedItem.Index) = txt_sublinea
         lv_detalle_agrupadores(2).ListItems.Item(lv_detalle_agrupadores(2).selectedItem.Index).ListSubItems(1) = vardetallearticulo
         lv_detalle_agrupadores(2).ListItems.Item(lv_detalle_agrupadores(2).selectedItem.Index).ListSubItems(2) = vardetallearticulo
         lv_detalle_agrupadores(2).ListItems.Item(lv_detalle_agrupadores(2).selectedItem.Index).ListSubItems(3) = vardetallearticulo
         lv_detalle_agrupadores(2).ListItems.Item(lv_detalle_agrupadores(2).selectedItem.Index).Selected = True
      End If
   End If
   If vartipoagrupador = 4 Then
      If var_modifica_registro = False Then
         Set list_item = lv_detalle_agrupadores(3).ListItems.Add(, , txt_producto)
         list_item.SubItems(1) = vardetallearticulo
         list_item.EnsureVisible
         list_item.Selected = True
      Else
         lv_detalle_agrupadores(3).ListItems.Item(lv_detalle_agrupadores(3).selectedItem.Index).Checked = False
         lv_detalle_agrupadores(3).ListItems.Item(lv_detalle_agrupadores(3).selectedItem.Index) = txt_producto
         lv_detalle_agrupadores(3).ListItems.Item(lv_detalle_agrupadores(3).selectedItem.Index).ListSubItems(1) = vardetallearticulo
         lv_detalle_agrupadores(3).ListItems.Item(lv_detalle_agrupadores(3).selectedItem.Index).Selected = True
      End If
   End If
   If vartipoagrupador = 5 Then
      If var_modifica_registro = False Then
         Set list_item = lv_detalle_agrupadores(4).ListItems.Add(, , txt_subtipo_producto)
         list_item.SubItems(1) = vardetallearticulo
         list_item.EnsureVisible
         list_item.Selected = True
      Else
         lv_detalle_agrupadores(4).ListItems.Item(lv_detalle_agrupadores(4).selectedItem.Index).Checked = False
         lv_detalle_agrupadores(4).ListItems.Item(lv_detalle_agrupadores(4).selectedItem.Index) = txt_subtipo_producto
         lv_detalle_agrupadores(4).ListItems.Item(lv_detalle_agrupadores(4).selectedItem.Index).ListSubItems(1) = vardetallearticulo
         lv_detalle_agrupadores(4).ListItems.Item(lv_detalle_agrupadores(4).selectedItem.Index).Selected = True
      End If
   End If
End Sub



Private Sub txt_articulo_Change()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_articulo_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_articulo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_articulos order by vcha_art_nombre_español", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_aRT_ARTICULO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ART_NOMBRE_ESPAÑOL), "", rs!VCHA_ART_NOMBRE_ESPAÑOL)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ARTICULOS"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3950
      Else
         lv_lista.ColumnHeaders(2).Width = 4150
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_articulo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_articulo_LostFocus()
   Dim var_posible As Boolean
   If Trim(txt_Articulo) <> "" Then
      var_posible = False
      rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = True
         txt_nombre_articulo = rs!VCHA_ART_NOMBRE_ESPAÑOL
         rs.Close
      Else
         rs.Close
         rs.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!VCHA_aRT_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_posible = True
               txt_Articulo = rs!VCHA_aRT_ARTICULO_ID
               txt_nombre_articulo = rsaux!VCHA_ART_NOMBRE_ESPAÑOL
               rsaux.Close
               rs.Close
            Else
               var_posible = False
               rsaux.Close
               rs.Close
            End If
         Else
            rs.Close
         End If
      End If
      If var_posible = True Then
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         txt_Articulo.SetFocus
      End If
   End If
End Sub
Private Sub txt_linea_Change()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_linea_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_linea_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_lineas order by vcha_lin_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_LIN_LINEA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_LIN_NOMBRE), "", rs!VCHA_LIN_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "LINEAS"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3950
      Else
         lv_lista.ColumnHeaders(2).Width = 4150
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_linea_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_linea_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_linea) <> "" Then
      rs.Open "SELECT * FROM TB_LINEAS WHERE VCHA_LIN_LINEA_ID = '" + txt_linea + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_linea = IIf(IsNull(rs!VCHA_LIN_NOMBRE), "", rs!VCHA_LIN_NOMBRE)
      Else
         txt_nombre_linea = ""
         txt_linea = ""
         MsgBox "Clave de linea incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_linea = ""
   End If
End Sub

Private Sub txt_nombre_articulo_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_articulo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_articulos order by vcha_art_nombre_español", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_aRT_ARTICULO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ART_NOMBRE_ESPAÑOL), "", rs!VCHA_ART_NOMBRE_ESPAÑOL)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ARTICULOS"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3950
      Else
         lv_lista.ColumnHeaders(2).Width = 4150
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_articulo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub

Private Sub txt_nombre_articulo_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_linea_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_linea_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_lineas order by vcha_lin_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_LIN_LINEA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_LIN_NOMBRE), "", rs!VCHA_LIN_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "LINEAS"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3950
      Else
         lv_lista.ColumnHeaders(2).Width = 4150
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_linea_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.opt_tipoagrupador(1).Value = True Then
         If cmd_guardar.Enabled = True Then
            cmd_guardar.SetFocus
         End If
      Else
         Me.txt_sublinea.SetFocus
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_linea_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_producto_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_producto_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_productos order by vcha_pro_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_PRO_PRODUCTO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_PRO_NOMBRE), "", rs!VCHA_PRO_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PRODUCTOS"
      var_tipo_lista = 4
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3950
      Else
         lv_lista.ColumnHeaders(2).Width = 4150
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_producto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.opt_tipoagrupador(3).Value = True Then
         If cmd_guardar.Enabled = True Then
            cmd_guardar.SetFocus
         End If
      Else
         Me.txt_subtipo_producto.SetFocus
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_producto_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_sublinea_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_sublinea_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_SUBLINEAS WHERE VCHA_LIN_LINEA_ID = '" + txt_linea + "' order by vcha_sli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_SLI_SUBLINEA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_SLI_NOMBRE), "", rs!VCHA_SLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "SUBLINEAS DE " + txt_nombre_linea
      var_tipo_lista = 3
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3950
      Else
         lv_lista.ColumnHeaders(2).Width = 4150
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_sublinea_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmd_guardar.Enabled = True Then
         cmd_guardar.SetFocus
      End If
   End If
End Sub

Private Sub txt_nombre_sublinea_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_subtipo_articulo_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_subtipo_articulo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_tipoarticulos order by vcha_tar_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TAR_TIPO_ARTICULO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TAR_NOMBRE), "", rs!VCHA_TAR_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPO PRODUCTOS"
      var_tipo_lista = 5
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3950
      Else
         lv_lista.ColumnHeaders(2).Width = 4150
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_subtipo_articulo_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_producto_Change()
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_producto_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_producto_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_productos order by vcha_pro_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_PRO_PRODUCTO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_PRO_NOMBRE), "", rs!VCHA_PRO_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PRODUCTOS"
      var_tipo_lista = 4
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3950
      Else
         lv_lista.ColumnHeaders(2).Width = 4150
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_producto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_producto_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_producto) <> "" Then
      rs.Open "SELECT * FROM TB_PRODUCTOS WHERE VCHA_RPO_PRODUCTO_ID = '" + txt_producto + "'"
      If Not rs.EOF Then
         txt_nombre_producto = IIf(IsNull(rs!VCHA_PRO_NOMBRE), "", rs!VCHA_PRO_NOMBRE)
      Else
         txt_nombre_producto = ""
         txt_producto = ""
         MsgBox "Clave de producto incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_producto = ""
   End If
End Sub

Private Sub txt_sublinea_Change()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_sublinea_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_sublinea_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_SUBLINEAS WHERE VCHA_LIN_LINEA_ID = '" + txt_linea + "' order by vcha_sli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_SLI_SUBLINEA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_SLI_NOMBRE), "", rs!VCHA_SLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "SUBLINEAS DE " + txt_nombre_linea
      var_tipo_lista = 3
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3950
      Else
         lv_lista.ColumnHeaders(2).Width = 4150
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_sublinea_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_sublinea_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_sublinea) <> "" Then
      rs.Open "SELECT * FROM TB_SUBLINEAS WHERE VCHA_SLI_SUBLINEA_ID = '" + txt_sublinea + "'"
      If Not rs.EOF Then
         txt_nombre_sublinea = IIf(IsNull(rs!VCHA_SLI_NOMBRE), "", rs!VCHA_SLI_NOMBRE)
      Else
         txt_sublinea = ""
         txt_nombre_sublinea = ""
         MsgBox "Clave de subliena incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_sublinea = ""
   End If
End Sub

Private Sub txt_subtipo_producto_Change()
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_subtipo_producto_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_subtipo_producto_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_tipoarticulos order by vcha_tar_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TAR_TIPO_ARTICULO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TAR_NOMBRE), "", rs!VCHA_TAR_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPO PRODUCTOS"
      var_tipo_lista = 5
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3950
      Else
         lv_lista.ColumnHeaders(2).Width = 4150
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_subtipo_producto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_subtipo_producto_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_subtipo_producto) <> "" Then
      rs.Open "select * from tb_tipoarticulos where vcha_tar_tipo_articulo_id = '" + txt_subtipo_producto + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_subtipo_articulo = IIf(IsNull(rs!VCHA_TAR_NOMBRE), "", rs!VCHA_TAR_NOMBRE)
      Else
         txt_nombre_subtipo_articulo = ""
         txt_subtipo_producto = ""
         MsgBox "Clave de tipo de artículo incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_subtipo_articulo = ""
   End If
End Sub
