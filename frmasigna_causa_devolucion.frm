VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmasigna_causa_devolucion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Causas de Devolución"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11685
   Begin VB.Frame frmlotes 
      Height          =   3840
      Left            =   3825
      TabIndex        =   43
      Top             =   1350
      Width           =   6300
      Begin VB.ComboBox cmb_estatus 
         Height          =   315
         ItemData        =   "frmasigna_causa_devolucion.frx":0000
         Left            =   1425
         List            =   "frmasigna_causa_devolucion.frx":000A
         TabIndex        =   52
         Top             =   3300
         Width           =   1860
      End
      Begin VB.ComboBox cmb_justifica 
         Height          =   315
         ItemData        =   "frmasigna_causa_devolucion.frx":0020
         Left            =   1425
         List            =   "frmasigna_causa_devolucion.frx":002A
         TabIndex        =   51
         Top             =   2955
         Width           =   1860
      End
      Begin VB.TextBox txt_proveedor 
         Height          =   315
         Left            =   1425
         TabIndex        =   50
         Top             =   2625
         Width           =   1665
      End
      Begin VB.ComboBox cmb_tipo_defecto 
         Height          =   315
         ItemData        =   "frmasigna_causa_devolucion.frx":0036
         Left            =   1425
         List            =   "frmasigna_causa_devolucion.frx":0040
         TabIndex        =   49
         Top             =   2280
         Width           =   2910
      End
      Begin VB.TextBox txt_nota_envio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         TabIndex        =   48
         Top             =   1935
         Width           =   1665
      End
      Begin VB.TextBox txt_fecha_lote 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         TabIndex        =   47
         Top             =   1590
         Width           =   1665
      End
      Begin VB.TextBox txt_supervisor 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         TabIndex        =   46
         Top             =   1245
         Width           =   4545
      End
      Begin VB.CommandButton cmd_cancelar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   420
         Picture         =   "frmasigna_causa_devolucion.frx":005D
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   405
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   90
         Picture         =   "frmasigna_causa_devolucion.frx":01A7
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   405
         Width           =   330
      End
      Begin VB.Frame Frame8 
         Height          =   45
         Left            =   15
         TabIndex        =   54
         Top             =   735
         Width           =   6240
      End
      Begin VB.TextBox txt_lote 
         Height          =   315
         Left            =   1425
         TabIndex        =   45
         Top             =   900
         Width           =   1680
      End
      Begin VB.Label Estatus 
         Caption         =   "Estatus:"
         Height          =   225
         Left            =   270
         TabIndex        =   63
         Top             =   3345
         Width           =   690
      End
      Begin VB.Label Label18 
         Caption         =   "Justifica:"
         Height          =   225
         Left            =   270
         TabIndex        =   62
         Top             =   3000
         Width           =   690
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor:"
         Height          =   195
         Left            =   270
         TabIndex        =   61
         Top             =   2685
         Width           =   780
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Tipo defecto:"
         Height          =   195
         Left            =   270
         TabIndex        =   60
         Top             =   2340
         Width           =   945
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Nota de envio:"
         Height          =   195
         Left            =   270
         TabIndex        =   59
         Top             =   1995
         Width           =   1050
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   270
         TabIndex        =   58
         Top             =   1650
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Supervisor:"
         Height          =   195
         Left            =   270
         TabIndex        =   57
         Top             =   1305
         Width           =   795
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
         Height          =   195
         Left            =   270
         TabIndex        =   53
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         Caption         =   "  Datos Calidad"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   44
         Top             =   135
         Width           =   6225
      End
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   765
      Picture         =   "frmasigna_causa_devolucion.frx":02F1
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Eliminar Movimiento "
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   2745
      TabIndex        =   31
      Top             =   1050
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1875
         Left            =   30
         TabIndex        =   32
         Top             =   435
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3307
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
         TabIndex        =   33
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11175
      Picture         =   "frmasigna_causa_devolucion.frx":043B
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_cerrar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmasigna_causa_devolucion.frx":0A75
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cerrar Alt + C"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmasigna_causa_devolucion.frx":0B77
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame frm_catalogo_correciones 
      Height          =   3165
      Left            =   5115
      TabIndex        =   27
      Top             =   1830
      Width           =   4125
      Begin MSComctlLib.ListView lv_catalogo_correcciones 
         Height          =   2865
         Left            =   45
         TabIndex        =   28
         Top             =   255
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   5054
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
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         Caption         =   "Posibles Correcciones"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   4110
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   " Corrección "
      Height          =   1155
      Left            =   5925
      TabIndex        =   25
      Top             =   6045
      Width           =   5670
      Begin MSComctlLib.ListView lv_detalle_correciones 
         Height          =   900
         Left            =   45
         TabIndex        =   26
         Top             =   180
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   1588
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2214
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "consecutivo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "destino"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame frm_causas_rechazo 
      Height          =   3165
      Left            =   4560
      TabIndex        =   20
      Top             =   1800
      Width           =   4125
      Begin MSComctlLib.ListView lv_causas_rechazo 
         Height          =   2865
         Left            =   45
         TabIndex        =   21
         Top             =   255
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   5054
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
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   "Causas de Rechazo (F8)"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   4125
      End
   End
   Begin VB.Frame frm_causas 
      Height          =   3165
      Left            =   4110
      TabIndex        =   8
      Top             =   1755
      Width           =   4110
      Begin MSComctlLib.ListView lv_causas 
         Height          =   2865
         Left            =   45
         TabIndex        =   9
         Top             =   255
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   5054
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Destino"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Causa de Devolución     Cliente(F5)  Real(F6)"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   4140
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " Causas de Devolución Real (F6) "
      Height          =   1170
      Left            =   90
      TabIndex        =   16
      Top             =   6030
      Width           =   5670
      Begin MSComctlLib.ListView lv_detalle_real 
         Height          =   870
         Left            =   45
         TabIndex        =   17
         Top             =   195
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   1535
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2214
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "consecutivo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "destino"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Causas de Devolución del Cliente (F5) "
      Height          =   1170
      Left            =   90
      TabIndex        =   14
      Top             =   4845
      Width           =   5670
      Begin MSComctlLib.ListView lv_detalle_cliente 
         Height          =   885
         Left            =   45
         TabIndex        =   15
         Top             =   195
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   1561
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2214
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "consecutivo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "destino"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Artículos (Asignación con F4) "
      Height          =   3465
      Left            =   90
      TabIndex        =   12
      Top             =   1350
      Width           =   11490
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   11070
         Picture         =   "frmasigna_causa_devolucion.frx":0C79
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   165
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   2880
         Left            =   60
         TabIndex        =   13
         Top             =   525
         Width           =   11370
         _ExtentX        =   20055
         _ExtentY        =   5080
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
         NumItems        =   16
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7232
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "consecutivo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "destino"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "sin factura"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "precio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "asignado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Lote"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Supervisor"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Fecha Lote"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Nota de Envio"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Tipo defecto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Proveedor"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Justificacion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Estatus"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_asignacion 
         Alignment       =   1  'Right Justify
         Caption         =   "999,999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3945
         TabIndex        =   41
         Top             =   210
         Width           =   990
      End
      Begin VB.Label lbl_movimiento 
         Alignment       =   1  'Right Justify
         Caption         =   "999,999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1590
         TabIndex        =   40
         Top             =   210
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Asignación:"
         Height          =   195
         Left            =   3000
         TabIndex        =   39
         Top             =   255
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento:"
         Height          =   195
         Left            =   615
         TabIndex        =   38
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Sin Facturar"
         Height          =   195
         Left            =   705
         TabIndex        =   24
         Top             =   6240
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H000000FF&
         Height          =   225
         Left            =   90
         TabIndex        =   23
         Top             =   6225
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Asignación de Causas a Devoluciones "
      Height          =   960
      Left            =   90
      TabIndex        =   0
      Top             =   390
      Width           =   11490
      Begin VB.TextBox txt_factura 
         Height          =   315
         Left            =   855
         TabIndex        =   36
         Top             =   540
         Width           =   1200
      End
      Begin VB.CheckBox chk_factura 
         Caption         =   "Devolución de una sola factura"
         Height          =   330
         Left            =   180
         TabIndex        =   35
         Top             =   210
         Width           =   2610
      End
      Begin VB.TextBox txt_nombre_movimiento 
         Height          =   315
         Left            =   4485
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   195
         Width           =   5010
      End
      Begin VB.TextBox txt_movimiento 
         Height          =   315
         Left            =   3855
         TabIndex        =   4
         Top             =   195
         Width           =   615
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   3855
         TabIndex        =   5
         Top             =   555
         Width           =   1755
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Factura:"
         Height          =   195
         Left            =   150
         TabIndex        =   37
         Top             =   615
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento:"
         Height          =   195
         Left            =   2970
         TabIndex        =   7
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   2970
         TabIndex        =   6
         Top             =   615
         Width           =   600
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   4485
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasigna_causa_devolucion.frx":0E8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasigna_causa_devolucion.frx":1769
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasigna_causa_devolucion.frx":2043
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasigna_causa_devolucion.frx":25DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasigna_causa_devolucion.frx":2EBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasigna_causa_devolucion.frx":3795
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasigna_causa_devolucion.frx":406F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasigna_causa_devolucion.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasigna_causa_devolucion.frx":4293
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasigna_causa_devolucion.frx":43A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasigna_causa_devolucion.frx":44B7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   90
      Left            =   60
      TabIndex        =   11
      Top             =   285
      Width           =   11550
   End
   Begin VB.Frame Frame6 
      Caption         =   " Causas de Rechazo "
      Height          =   1170
      Left            =   5910
      TabIndex        =   18
      Top             =   4830
      Width           =   5670
      Begin MSComctlLib.ListView lv_detalle_rechazo 
         Height          =   870
         Left            =   45
         TabIndex        =   19
         Top             =   225
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   1535
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2214
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "consecutivo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "destino"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmasigna_causa_devolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_devolucion As Integer
Dim var_almacen As String
Dim var_almacen_Destino As String
Dim var_estatus As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim item_seleccionado As Integer
Dim var_movimiento_factura As Integer
Dim var_almacen_costeo As Integer
Dim var_requiere_factura As Integer
Dim var_año As Integer
Dim var_ventana As Integer



Private Sub marca()

End Sub


Private Sub detalle_causas()
   Dim list_item As ListItem
   lv_detalle_cliente.ListItems.Clear
   lv_detalle_real.ListItems.Clear
   lv_detalle_rechazo.ListItems.Clear
   lv_detalle_correciones.ListItems.Clear
   rsaux3.Open "select * from vw_detalle_devolucion_cliente where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_art_articulo_id = '" + lv_articulos.selectedItem + "' and inte_cde_consecutivo  = " + lv_articulos.selectedItem.SubItems(2), cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux3.EOF Then
      While Not rsaux3.EOF
         Set list_item = lv_detalle_cliente.ListItems.Add(, , rsaux3!INTE_CDE_CAUSA_ID)
         list_item.SubItems(1) = IIf(IsNull(rsaux3!vcha_cde_nombre), "", rsaux3!vcha_cde_nombre)
         rsaux3.MoveNext
      Wend
   End If
   rsaux3.Close
   rsaux3.Open "select * from vw_detalle_devolucion_real where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_art_articulo_id = '" + lv_articulos.selectedItem + "' and inte_cde_consecutivo  = " + lv_articulos.selectedItem.SubItems(2), cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux3.EOF Then
      While Not rsaux3.EOF
         Set list_item = lv_detalle_real.ListItems.Add(, , rsaux3!INTE_CDE_CAUSA_ID)
         list_item.SubItems(1) = rsaux3!vcha_cde_nombre
         rsaux3.MoveNext
      Wend
   End If
   rsaux3.Close
   rsaux3.Open "select * from vw_detalle_devolucion_RECHAZO where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_art_articulo_id = '" + lv_articulos.selectedItem + "' and inte_cde_consecutivo  = " + lv_articulos.selectedItem.SubItems(2), cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux3.EOF Then
      While Not rsaux3.EOF
         Set list_item = lv_detalle_rechazo.ListItems.Add(, , rsaux3!inte_cRe_causa_id)
         list_item.SubItems(1) = rsaux3!vcha_cRe_nombre
         rsaux3.MoveNext
      Wend
   End If
   rsaux3.Close
   rsaux3.Open "select * from vw_detalle_devolucion_ajustes where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_art_articulo_id = '" + lv_articulos.selectedItem + "' and inte_cde_consecutivo  = " + lv_articulos.selectedItem.SubItems(2), cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux3.EOF Then
      While Not rsaux3.EOF
         Set list_item = lv_detalle_correciones.ListItems.Add(, , rsaux3!INTE_CDE_CAUSA_ID)
         list_item.SubItems(1) = rsaux3!VCHA_CAJ_NOMBRE
         rsaux3.MoveNext
      Wend
   End If
   rsaux3.Close
End Sub


Private Sub cmb_movimientos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txt_movimiento) <> "" Then
         txt_movimiento.Enabled = False
         cmb_movimientos.Enabled = False
         txt_numero.Enabled = True
         txt_numero.SetFocus
      End If
   End If
End Sub

Private Sub chk_factura_Click()
   Me.frmlotes.Visible = False
   If chk_factura.Value = 0 Then
      txt_factura = ""
      txt_factura.Enabled = False
   Else
      txt_factura = ""
      txt_factura.Enabled = True
      txt_factura.SetFocus
   End If
End Sub

Private Sub cmb_estatus_KeyPress(KeyAscii As Integer)
   Me.cmd_aceptar.SetFocus
End Sub

Private Sub cmb_estatus_LostFocus()
   If Me.cmb_estatus <> "Primera" And Me.cmb_estatus <> "Segunda" Then
      MsgBox "Estatus no valido", vbOKOnly, "ATENCION"
      Me.cmb_estatus = ""
   End If
End Sub

Private Sub cmb_justifica_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub cmb_justifica_LostFocus()
   If Me.cmb_justifica <> "Sí" And Me.cmb_justifica <> "No" Then
      MsgBox "Tipo de justificación no valida", vbOKOnly, "ATENCION"
      Me.cmb_justifica = ""
   End If
End Sub

Private Sub cmb_tipo_defecto_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub


Private Sub cmb_tipo_defecto_LostFocus()
   If Me.cmb_tipo_defecto <> "Imputable" And Me.cmb_tipo_defecto <> "No Imputable" Then
      MsgBox "Tipo de defecto no valido", vbOKOnly, "ATENCION"
      Me.cmb_tipo_defecto = ""
   End If
End Sub

Private Sub cmd_aceptar_Click()
   var_si = MsgBox("¿Desea actualizar los datos?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      If Trim(Me.txt_lote) <> "" Then
         If Trim(cmb_tipo_defecto) <> "" Then
            If Trim(Me.txt_proveedor) <> "" Then
               If Trim(Me.cmb_justifica) <> "" Then
                  If Trim(Me.cmb_estatus) <> "" Then
                     rsaux5.Open "update tb_devoluciones set inte_dev_lote = " + CStr(Me.txt_lote) + ", vcha_dev_supervisor = '" + Me.txt_supervisor + "', dtim_dev_fecha_lote = '" + Format(Me.txt_fecha_lote, "Short Date") + "', vcha_dev_nota_envio= '" + Me.txt_nota_envio + "', VCHA_DEV_TIPO_DEFECTO = '" + Me.cmb_tipo_defecto + "', VCHA_DEV_PROVEEDOR = '" + Me.txt_proveedor + "', VCHA_DEV_JUSTIFICA_DEVOLUCION = '" + Me.cmb_justifica + "', VCHA_DEV_ESTATUS = '" + Me.cmb_estatus + "' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id ='" + Me.txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_art_articulo_id = '" + Me.lv_articulos.selectedItem + "' and inte_cde_consecutivo = " + lv_articulos.selectedItem.SubItems(2), cnn, adOpenDynamic, adLockOptimistic
                     lv_articulos.selectedItem.SubItems(8) = Me.txt_lote
                     lv_articulos.selectedItem.SubItems(9) = Me.txt_supervisor
                     lv_articulos.selectedItem.SubItems(10) = Format(Me.txt_fecha_lote, "Short Date")
                     lv_articulos.selectedItem.SubItems(11) = Me.txt_nota_envio
                     lv_articulos.selectedItem.SubItems(12) = Me.cmb_tipo_defecto
                     lv_articulos.selectedItem.SubItems(13) = Me.txt_proveedor
                     lv_articulos.selectedItem.SubItems(14) = Me.cmb_justifica
                     lv_articulos.selectedItem.SubItems(15) = Me.cmb_estatus
                     Me.frmlotes.Visible = False
                  Else
                     MsgBox "Debe de indicar el estatus del artículo", vbOKOnly, "ATENCION"
                     Me.cmb_estatus.SetFocus
                  End If
               Else
                  MsgBox "Debe indicar si se justifica o no la devolución", vbOKOnly, "ATENCION"
                  Me.cmb_justifica.SetFocus
               End If
            Else
               MsgBox "Se debe de indicar un proveedor", vbOKOnly, "ATENCION"
               Me.txt_proveedor.SetFocus
            End If
         Else
            MsgBox "Se debe de indicar un tipo de defecto", vbOKOnly, "ATENCION"
            Me.cmb_tipo_defecto.SetFocus
         End If
      Else
         rsaux5.Open "update tb_devoluciones set inte_dev_lote = null, vcha_dev_supervisor = '', dtim_dev_fecha_lote = null, vcha_dev_nota_envio= '', VCHA_DEV_TIPO_DEFECTO = '', VCHA_DEV_PROVEEDOR = '', VCHA_DEV_JUSTIFICA_DEVOLUCION = '', VCHA_DEV_ESTATUS = '' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id ='" + Me.txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_art_articulo_id = '" + Me.lv_articulos.selectedItem + "' and inte_cde_consecutivo = " + lv_articulos.selectedItem.SubItems(2), cnn, adOpenDynamic, adLockOptimistic
         lv_articulos.selectedItem.SubItems(8) = ""
         lv_articulos.selectedItem.SubItems(9) = ""
         lv_articulos.selectedItem.SubItems(10) = ""
         lv_articulos.selectedItem.SubItems(11) = ""
         lv_articulos.selectedItem.SubItems(12) = ""
         lv_articulos.selectedItem.SubItems(13) = ""
         lv_articulos.selectedItem.SubItems(14) = ""
         lv_articulos.selectedItem.SubItems(15) = ""
         Me.frmlotes.Visible = False
      End If
   End If
End Sub

Private Sub cmd_cancelar_Click()
   Me.frmlotes.Visible = False
End Sub

Private Sub cmd_cerrar_Click()
        Me.frmlotes.Visible = False
         Dim var_posible_valuado As Boolean
         Dim var_posible_asignado As Boolean
         If Val(txt_numero) > 0 Then
            If Trim(var_estatus) = "" Then
               lv_articulos.ListItems.Clear
               rs.Open "select * from tb_devoluciones where vcha_emp_empresa_id = '" + var_empresa + "' and  vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero, cnn, adOpenDynamic
               var_almacen = rs!VCHA_ALM_ALMACEN_ID
               If Not rs.EOF Then
                  var_contador_articulos = 0
                  While Not rs.EOF
                     var_estatus = IIf(IsNull(rs!CHAR_CDE_ESTATUS), "", rs!CHAR_CDE_ESTATUS)
                     var_nombre_articulo = ""
                     rsaux2.Open "select * from tb_articulos where vcha_Art_articulo_id ='" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        var_nombre_articulo = rsaux2!vcha_Art_nombre_español
                     End If
                     rsaux2.Close
                     Set list_item = lv_articulos.ListItems.Add(, , rs!vcha_Art_Articulo_id)
                     list_item.SubItems(1) = Trim(var_nombre_articulo)
                     list_item.SubItems(2) = IIf(IsNull(rs!INTE_CDE_CONSECUTIVO), 0, rs!INTE_CDE_CONSECUTIVO)
                     If rs!inte_fac_factura = 0 Then
                        list_item.SubItems(5) = "*"
                     End If
                     list_item.SubItems(6) = IIf(IsNull(rs!floa_cde_precio), 0, rs!floa_cde_precio)
                     list_item.SubItems(7) = IIf(IsNull(rs!inte_cde_asignado), 0, rs!inte_cde_asignado)
                     var_contador_articulos = var_contador_articulos + 1
                     rs.MoveNext
                  Wend
                  If var_contador_articulos > 27 Then
                     lv_articulos.ColumnHeaders(2).Width = 3900
                  Else
                     lv_articulos.ColumnHeaders(2).Width = 4100
                  End If
                  txt_numero.Enabled = False
                  txt_movimiento.Enabled = False
                  j = lv_articulos.ListItems.Count
                  For i = 1 To j
                     lv_articulos.ListItems.item(i).Selected = True
                     If lv_articulos.selectedItem.SubItems(5) = "*" Then
                        lv_articulos.selectedItem.Bold = True
                        lv_articulos.selectedItem.ListSubItems.item(1).Bold = True
                        lv_articulos.selectedItem.ListSubItems.item(2).Bold = True
                        lv_articulos.selectedItem.ListSubItems.item(3).Bold = True
                        lv_articulos.selectedItem.ForeColor = &HFF&
                        lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &HFF&
                        lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &HFF&
                        lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &HFF&
                     End If
                  Next i
                  lv_articulos.ListItems.item(1).Selected = True
               End If
               rs.Close
               Call detalle_causas
               
               j = lv_articulos.ListItems.Count
               var_posible_valuado = True
               var_posible_asignado = True
               For i = 1 To j
                  lv_articulos.ListItems.item(i).Selected = True
                  If lv_articulos.selectedItem.SubItems(6) = 0 Then
                     var_posible_valuado = False
                  End If
                  If lv_articulos.selectedItem.SubItems(7) = 0 Then
                     var_posible_asignado = False
                  End If
               Next i
               If var_posible_valuado = True Then
                  If var_posible_asignado = True Then
                     si = MsgBox("¿Deseas cerrar el movimiento?", vbYesNo, "ATENCION")
                     If si = 6 Then
                        si = MsgBox("Confirmar el cerrado del movimiento", vbYesNo, "ATENCION")
                        If si = 6 Then
                           Set TB_DEVOLUCIONES_ESTATUS = New TB_DEVOLUCIONES_ESTATUS
                           var_estatus = "I"
                           var_modifica = False
                           var_modifica = TB_DEVOLUCIONES_ESTATUS.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_movimiento, txt_numero, "I")
                           MsgBox "El movimiento a sido cerrado", vbOKOnly, "ATENCION"
                           rs.Open "select * from tb_detalle_devolucion_real where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and inte_cde_causa_id = 19", cnn, adOpenDynamic, adLockOptimistic
                           If Not rs.EOF Then
                              While Not rs.EOF
                                 rs.MoveNext
                              Wend
                           End If
                           rs.Close
                        End If
                     End If
                  Else
                     MsgBox "Faltan artículos por asignarles causa real", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "Faltan artículos por valuar", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El movimiento ya habia sido cerrado con anterioridad", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se a seleccionado ningun movimiento", vbOKOnly, "ATENCION"
         End If
End Sub

Private Sub cmd_cerrar_GotFocus()
   Me.frmlotes.Visible = False
End Sub

Private Sub cmd_eliminar_Click()
   Me.frmlotes.Visible = False
   var_si = MsgBox("¿Deseas eliminar el movimiento?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox("Confirmar la eliminación del movimiento", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rs.Open "SELECT * FROM TB_DETALLE_DEVOLUCION_CLIENTE WHERE VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_EMO_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            rsaux.Open "SELECT * FROM TB_DETALLE_DEVOLUCION_REAL WHERE VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_EMO_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
            If rsaux.EOF Then
               rsaux1.Open "SELECT * FROM TB_DETALLE_DEVOLUCION_REAL WHERE VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_EMO_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
               If rsaux1.EOF Then
                  rsaux2.Open "SELECT * FROM TB_DEVOLUCIONES WHERE VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_EMO_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                  If Trim(IIf(IsNull(rsaux2!CHAR_CDE_ESTATUS), "", rsaux2!CHAR_CDE_ESTATUS)) = "" Then
                     rsaux3.Open "DELETE FROM TB_DEVOLUCIONES WHERE VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_EMO_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                  Else
                     MsgBox "El movimiento ya no puede ser eliminado", vbOKOnly, "ATENCION"
                  End If
                  rsaux2.Close
               Else
                  MsgBox "El movimiento ya no puede ser eliminado", vbOKOnly, "ATENCION"
               End If
               rsaux1.Close
            Else
               MsgBox "El movimiento ya no puede ser eliminado", vbOKOnly, "ATENCION"
            End If
            rsaux.Close
         Else
            MsgBox "El movimiento ya no puede ser eliminado", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
   End If
End Sub

Private Sub cmd_eliminar_GotFocus()
   Me.frmlotes.Visible = False
End Sub

Private Sub cmd_nuevo_Click()
   Me.frmlotes.Visible = False
   Me.lbl_asignacion = "0"
   Me.lbl_movimiento = "0"
   chk_factura.Value = 0
   txt_factura = ""
   lv_articulos.ListItems.Clear
   lv_detalle_cliente.ListItems.Clear
   lv_detalle_real.ListItems.Clear
   lv_detalle_correciones.ListItems.Clear
   lv_detalle_rechazo.ListItems.Clear
   txt_movimiento = ""
   txt_numero = ""
   cmb_movimientos = ""
   txt_movimiento.Enabled = True
   txt_movimiento.SetFocus
   txt_numero.Enabled = False
   txt_nombre_movimiento = ""
End Sub

Private Sub cmd_nuevo_GotFocus()
   Me.frmlotes.Visible = False
End Sub

Private Sub cmd_salir_Click()
   Me.frmlotes.Visible = False
   Unload Me
End Sub

Private Sub cmd_salir_GotFocus()
   Me.frmlotes.Visible = False
End Sub

Private Sub cmd_todos_Click()
   For var_i = 1 To lv_articulos.ListItems.Count
      lv_articulos.ListItems.item(var_i).Selected = True
      i = lv_articulos.selectedItem.Index
      If lv_articulos.selectedItem.SubItems(4) = "*" Then
         If lv_articulos.selectedItem.SubItems(5) = "*" Then
            lv_articulos.selectedItem.SubItems(4) = ""
            lv_articulos.selectedItem.Bold = True
            lv_articulos.selectedItem.ListSubItems.item(1).Bold = True
            lv_articulos.selectedItem.ListSubItems.item(2).Bold = True
            lv_articulos.selectedItem.ListSubItems.item(3).Bold = True
            lv_articulos.selectedItem.ForeColor = &HFF&
            lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &HFF&
            lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &HFF&
            lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &HFF&
         Else
            lv_articulos.selectedItem.SubItems(4) = ""
            lv_articulos.selectedItem.Bold = False
            lv_articulos.selectedItem.ListSubItems.item(1).Bold = False
            lv_articulos.selectedItem.ListSubItems.item(2).Bold = False
            lv_articulos.selectedItem.ListSubItems.item(3).Bold = False
            lv_articulos.selectedItem.ListSubItems.item(4).Bold = False
            lv_articulos.selectedItem.ForeColor = &H0&
            lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
            lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
            lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
            lv_articulos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
         End If
      Else
         lv_articulos.selectedItem.SubItems(4) = "*"
         lv_articulos.selectedItem.Bold = True
         lv_articulos.selectedItem.ListSubItems.item(1).Bold = True
         lv_articulos.selectedItem.ListSubItems.item(2).Bold = True
         lv_articulos.selectedItem.ListSubItems.item(3).Bold = True
         lv_articulos.selectedItem.ListSubItems.item(4).Bold = True
         lv_articulos.selectedItem.ForeColor = &HFF0000
         lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &HFF0000
         lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &HFF0000
         lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &HFF0000
         lv_articulos.selectedItem.ListSubItems.item(4).ForeColor = &HFF0000
      End If
    Next var_i
End Sub

Private Sub cmd_todos_GotFocus()
   Me.frmlotes.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 67 Then
      cmd_cerrar_Click
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
         
      'Unload Me
   End If
End Sub

Private Sub Form_Load()
   var_ventana = 0
   frmlotes.Visible = False
   var_cadena_seguridad = ""
   frm_lista.Visible = False
   Top = 0
   Left = 0
   txt_factura.Enabled = False
   chk_factura.Value = 0
   frm_causas.Visible = False
   frm_causas_rechazo.Visible = False
   frm_catalogo_correciones.Visible = False
   var_tipo_devolucion = 0
   Me.lbl_asignacion = 0
   Me.lbl_movimiento = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_asigna_causa_devolucion)
End Sub

Private Sub lv_articulos_GotFocus()
   Me.frmlotes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Seleccione los artículos presionando ENTER sobre ellos, presione F4 para seleccionar las causas de devolución F5 para asignar el Lote"
End Sub

Private Sub lv_articulos_ItemClick(ByVal item As MSComctlLib.ListItem)
   Call detalle_causas
   item_seleccionado = lv_articulos.selectedItem.Index
End Sub

Private Sub lv_articulos_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub lv_catalogo_correcciones_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_catalogo_correcciones, ColumnHeader)
End Sub

Private Sub lv_catalogo_correcciones_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Seleccione las correciones con ENTER y presione F9 para aplicarlas"
End Sub

Private Sub lv_catalogo_correcciones_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim l As Integer
   Set TB_DETALLE_DEVOLUCION_AJUSTE = New TB_DETALLE_DEVOLUCION_AJUSTE
   If KeyCode = 120 Then
      var_tipo_detalle_devolucion = 0
      j = lv_articulos.ListItems.Count
      l = lv_catalogo_correcciones.ListItems.Count
      For i = 1 To j
         lv_articulos.ListItems.item(i).Selected = True
         If lv_articulos.selectedItem.SubItems(4) = "*" Then
            lv_articulos.selectedItem.SubItems(4) = ""
            If lv_articulos.selectedItem.SubItems(5) = "*" Then
               lv_articulos.selectedItem.Bold = True
               lv_articulos.selectedItem.ListSubItems.item(1).Bold = True
               lv_articulos.selectedItem.ListSubItems.item(2).Bold = True
               lv_articulos.selectedItem.ListSubItems.item(3).Bold = True
               lv_articulos.selectedItem.ForeColor = &HFF&
               lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &HFF&
               lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &HFF&
               lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &HFF&
            Else
               lv_articulos.selectedItem.Bold = False
               lv_articulos.selectedItem.ListSubItems.item(1).Bold = False
               lv_articulos.selectedItem.ListSubItems.item(2).Bold = False
               lv_articulos.selectedItem.ListSubItems.item(3).Bold = False
               lv_articulos.selectedItem.ListSubItems.item(4).Bold = False
               lv_articulos.selectedItem.ForeColor = &H0&
               lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
               lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
               lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
               lv_articulos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
            End If
            For k = 1 To l
               lv_catalogo_correcciones.ListItems.item(k).Selected = True
               If lv_catalogo_correcciones.selectedItem.SubItems(2) = "*" Then
                  var_modifica = TB_DETALLE_DEVOLUCION_AJUSTE.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_movimiento, txt_numero, lv_articulos.selectedItem, lv_articulos.selectedItem.SubItems(2), lv_catalogo_correcciones.selectedItem)
               End If
            Next k
         End If
      Next i
      frm_catalogo_correciones.Visible = False
      lv_articulos.SetFocus
      Call detalle_causas
   End If
End Sub

Private Sub lv_catalogo_correcciones_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim i As Integer
      i = lv_catalogo_correcciones.selectedItem.Index
      If lv_catalogo_correcciones.selectedItem.SubItems(2) = "*" Then
         lv_catalogo_correcciones.selectedItem.SubItems(2) = ""
         lv_catalogo_correcciones.selectedItem.Bold = False
         lv_catalogo_correcciones.selectedItem.ListSubItems.item(1).Bold = False
         lv_catalogo_correcciones.selectedItem.ListSubItems.item(2).Bold = False
         lv_catalogo_correcciones.selectedItem.ForeColor = &H0&
         lv_catalogo_correcciones.selectedItem.ListSubItems.item(1).ForeColor = &H0&
         lv_catalogo_correcciones.selectedItem.ListSubItems.item(2).ForeColor = &H0&
      Else
         lv_catalogo_correcciones.selectedItem.SubItems(2) = "*"
         lv_catalogo_correcciones.selectedItem.Bold = True
         lv_catalogo_correcciones.selectedItem.ListSubItems.item(1).Bold = True
         lv_catalogo_correcciones.selectedItem.ListSubItems.item(2).Bold = True
         lv_catalogo_correcciones.selectedItem.ForeColor = &HC0&
         lv_catalogo_correcciones.selectedItem.ListSubItems.item(1).ForeColor = &HC0&
         lv_catalogo_correcciones.selectedItem.ListSubItems.item(2).ForeColor = &HC0&
      End If
   End If
   If KeyAscii = 27 Then
      frm_catalogo_correciones.Visible = False
   End If
End Sub

Private Sub lv_catalogo_correcciones_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub lv_causas_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Seleccione las causas de devolución con ENTER y presione F5 para aplicar la causa del cliente y F6 para aplicar la causa real"
End Sub

Private Sub lv_causas_KeyDown(KeyCode As Integer, Shift As Integer)
      Dim i As Integer
      Dim j As Integer
      Dim k As Integer
      Dim l As Integer
      Dim m As Integer
      Dim n As Integer
      Set TB_DETALLE_DEVOLUCION_CLIENTE_I = New TB_DETALLE_DEVOLUCION_CLIENTE_I
      Set TB_DETALLE_DEVOLUCION_REAL_I = New TB_DETALLE_DEVOLUCION_REAL_I
   If KeyCode = 116 Then
      var_modifica = False
      var_tipo_detalle_devolucion = 0
      j = lv_articulos.ListItems.Count
      For i = 1 To j
         lv_articulos.ListItems.item(i).Selected = True
         If lv_articulos.selectedItem.SubItems(4) = "*" Then
            l = lv_causas.ListItems.Count
            For k = 1 To l
               lv_causas.ListItems.item(k).Selected = True
               If lv_causas.selectedItem.SubItems(3) = "*" Then
                  var_almacen_Destino = Trim(lv_causas.selectedItem.SubItems(2))
                  If var_almacen_Destino = "REMPAQUE" Then
                     rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + lv_articulos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_almacen_Destino = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
                     rs.Close
                  End If
                  var_modifica = TB_DETALLE_DEVOLUCION_CLIENTE_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_movimiento, txt_numero, lv_articulos.selectedItem, lv_articulos.selectedItem.SubItems(2), lv_causas.selectedItem)
                  lv_articulos.selectedItem.SubItems(3) = var_almacen_Destino
                  lv_articulos.selectedItem.SubItems(4) = ""
                  If lv_articulos.selectedItem.SubItems(5) = "*" Then
                     lv_articulos.selectedItem.Bold = True
                     lv_articulos.selectedItem.ListSubItems.item(1).Bold = True
                     lv_articulos.selectedItem.ListSubItems.item(2).Bold = True
                     lv_articulos.selectedItem.ListSubItems.item(3).Bold = True
                     lv_articulos.selectedItem.ForeColor = &HFF&
                     lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &HFF&
                     lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &HFF&
                     lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &HFF&
                  Else
                     lv_articulos.selectedItem.Bold = False
                     lv_articulos.selectedItem.ListSubItems.item(1).Bold = False
                     lv_articulos.selectedItem.ListSubItems.item(2).Bold = False
                     lv_articulos.selectedItem.ListSubItems.item(3).Bold = False
                     lv_articulos.selectedItem.ListSubItems.item(4).Bold = False
                     lv_articulos.selectedItem.ForeColor = &H0&
                     lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
                     lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
                     lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
                     lv_articulos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
                  End If
               End If
            Next k
         End If
      Next i
      For k = 1 To l
         lv_causas.ListItems.item(k).Selected = True
         lv_causas.selectedItem.SubItems(3) = ""
         lv_causas.selectedItem.Bold = False
         lv_causas.selectedItem.ListSubItems.item(1).Bold = False
         lv_causas.selectedItem.ListSubItems.item(2).Bold = False
         lv_causas.selectedItem.ForeColor = &H0&
         lv_causas.selectedItem.ListSubItems.item(1).ForeColor = &H0&
         lv_causas.selectedItem.ListSubItems.item(2).ForeColor = &H0&
      Next k
      frm_causas.Visible = False
      lv_articulos.SetFocus
      Call detalle_causas
   End If
   
   If KeyCode = 117 Then
      var_modifica = False
      j = lv_articulos.ListItems.Count
      var_tipo_detalle_devolucion = 0
      For i = 1 To j
         lv_articulos.ListItems.item(i).Selected = True
         If lv_articulos.selectedItem.SubItems(4) = "*" Then
            l = lv_causas.ListItems.Count
            For k = 1 To l
               lv_causas.ListItems.item(k).Selected = True
               If lv_causas.selectedItem.SubItems(3) = "*" Then
                  var_almacen_Destino = Trim(lv_causas.selectedItem.SubItems(2))
                  If var_almacen_Destino = "REMPAQUE" Then
                     rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + lv_articulos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_almacen_Destino = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
                     rs.Close
                  End If
                  rsaux2.Open "select * from tb_detalle_devolucion_real where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_art_articulo_id = '" + lv_articulos.selectedItem + "' and inte_cde_consecutivo = " + Str(lv_articulos.selectedItem.SubItems(2)) + " and inte_cde_causa_id = " + Str(lv_causas.selectedItem), cnn, adOpenDynamic, adLockOptimistic
                  If rsaux2.EOF Then
                     lv_articulos.selectedItem.SubItems(7) = lv_articulos.selectedItem.SubItems(7) + 1
                  End If
                  rsaux2.Close
                  var_modifica = TB_DETALLE_DEVOLUCION_REAL_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_movimiento, txt_numero, lv_articulos.selectedItem, lv_articulos.selectedItem.SubItems(2), lv_causas.selectedItem, var_requiere_factura)
                  lv_articulos.selectedItem.SubItems(3) = var_almacen_Destino
                  lv_articulos.selectedItem.SubItems(4) = ""
                  If lv_articulos.selectedItem.SubItems(5) = "*" Then
                     lv_articulos.selectedItem.Bold = True
                     lv_articulos.selectedItem.ListSubItems.item(1).Bold = True
                     lv_articulos.selectedItem.ListSubItems.item(2).Bold = True
                     lv_articulos.selectedItem.ListSubItems.item(3).Bold = True
                     lv_articulos.selectedItem.ForeColor = &HFF&
                     lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &HFF&
                     lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &HFF&
                     lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &HFF&
                  Else
                     lv_articulos.selectedItem.Bold = False
                     lv_articulos.selectedItem.ListSubItems.item(1).Bold = False
                     lv_articulos.selectedItem.ListSubItems.item(2).Bold = False
                     lv_articulos.selectedItem.ListSubItems.item(3).Bold = False
                     lv_articulos.selectedItem.ListSubItems.item(4).Bold = False
                     lv_articulos.selectedItem.ForeColor = &H0&
                     lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
                     lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
                     lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
                     lv_articulos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
                  End If
               End If
            Next k
         End If
      Next i
      For k = 1 To l
         lv_causas.ListItems.item(k).Selected = True
         lv_causas.selectedItem.SubItems(3) = ""
         lv_causas.selectedItem.Bold = False
         lv_causas.selectedItem.ListSubItems.item(1).Bold = False
         lv_causas.selectedItem.ListSubItems.item(2).Bold = False
         lv_causas.selectedItem.ForeColor = &H0&
         lv_causas.selectedItem.ListSubItems.item(1).ForeColor = &H0&
         lv_causas.selectedItem.ListSubItems.item(2).ForeColor = &H0&
      Next k
      frm_causas.Visible = False
      If item_seleccionado > 0 Then
         If item_seleccionado <= lv_articulos.ListItems.Count Then
            lv_articulos.ListItems.item(item_seleccionado).Selected = True
         End If
      End If
      lv_articulos.SetFocus
      Call detalle_causas
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      rs.Open "select * from tb_movimientos where vcha_mov_movimiento_id = '" + txt_movimiento + "' and INTE_MOV_CAUSA_DEVOLUCION = 1", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         cmb_movimientos = rs!vcha_mov_nombre
         rs.Close
      Else
         rs.Close
         MsgBox "Clave de Movimiento Incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
End Sub

Private Sub lv_articulos_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim list_item As ListItem
   Dim i As Integer
   Dim j As Integer
   Dim var_posible As Boolean
   If KeyCode = 115 Then
      If Me.lbl_asignacion <> Me.lbl_movimiento Then
         MsgBox "No se cargo el total de los artículos, elimine el movimiento y vuelvalo a intentar", vbOKOnly, "ATENCION"
      Else
      
      If Trim(var_estatus) = "" Then
         var_posible = False
         j = lv_articulos.ListItems.Count
         For i = 1 To j
             lv_articulos.ListItems.item(i).Selected = True
             If lv_articulos.selectedItem.SubItems(4) = "*" Then
                var_posible = True
             End If
         Next i
         If var_posible = True Then
            var_tipo_devolucion = 2
            lv_causas.ListItems.Clear
            rs.Open "select * from tb_causas_devolucion order by vcha_cde_nombre", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
               Set list_item = lv_causas.ListItems.Add(, , rs!INTE_CDE_CAUSA_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_cde_nombre), "", rs!vcha_cde_nombre)
               list_item.SubItems(2) = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
               rs.MoveNext
            Wend
            rs.Close
            var_ventana = 1
            frm_causas.Visible = True
            lv_causas.SetFocus
         Else
            MsgBox "No se a seleccionado ningun artículo", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El movimiento ya no puede ser modificado", vbOKOnly, "ATENCION"
      End If
      End If
   End If
   If KeyCode = 120 Then
      If Me.lbl_asignacion <> Me.lbl_movimiento Then
         MsgBox "No se cargo el total de los artículos, elimine el movimiento y vuelvalo a intentar", vbOKOnly, "ATENCION"
      Else
      var_posible = False
      j = lv_articulos.ListItems.Count
      For i = 1 To j
          lv_articulos.ListItems.item(i).Selected = True
          If lv_articulos.selectedItem.SubItems(4) = "*" Then
             var_posible = True
          End If
      Next i
      If var_posible = True Then
         var_ventana = 2
         frm_catalogo_correciones.Visible = True
         rs.Open "select * from tb_causas_ajuste order by vcha_CAJ_nombre", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
            Set list_item = lv_catalogo_correcciones.ListItems.Add(, , rs!INTE_CAJ_CAUSA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CAJ_NOMBRE), "", rs!VCHA_CAJ_NOMBRE)
            rs.MoveNext
         Wend
         rs.Close
         lv_catalogo_correcciones.SetFocus
     Else
        MsgBox "No se a seleccionado ningún artículo", vbOKOnly, "ATENCION"
      End If
      End If
   End If
   If KeyCode = 116 Then
      If var_empresa <> "18" Then
         var_ventana = 3
         Me.txt_lote = Me.lv_articulos.selectedItem.SubItems(8)
         Me.txt_supervisor = Me.lv_articulos.selectedItem.SubItems(9)
         Me.txt_fecha_lote = Me.lv_articulos.selectedItem.SubItems(10)
         Me.txt_nota_envio = Me.lv_articulos.selectedItem.SubItems(11)
         Me.cmb_tipo_defecto = Me.lv_articulos.selectedItem.SubItems(12)
         Me.txt_proveedor = Me.lv_articulos.selectedItem.SubItems(13)
         Me.cmb_justifica = Me.lv_articulos.selectedItem.SubItems(14)
         Me.cmb_estatus = Me.lv_articulos.selectedItem.SubItems(15)
         frmlotes.Visible = True
         Me.txt_lote.SetFocus
      End If
   End If
End Sub

Private Sub lv_articulos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim i As Integer
      i = lv_articulos.selectedItem.Index
     
      If lv_articulos.selectedItem.SubItems(4) = "*" Then
         If lv_articulos.selectedItem.SubItems(5) = "*" Then
            lv_articulos.selectedItem.SubItems(4) = ""
            lv_articulos.selectedItem.Bold = True
            lv_articulos.selectedItem.ListSubItems.item(1).Bold = True
            lv_articulos.selectedItem.ListSubItems.item(2).Bold = True
            lv_articulos.selectedItem.ListSubItems.item(3).Bold = True
            lv_articulos.selectedItem.ForeColor = &HFF&
            lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &HFF&
            lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &HFF&
            lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &HFF&
         Else
            lv_articulos.selectedItem.SubItems(4) = ""
            lv_articulos.selectedItem.Bold = False
            lv_articulos.selectedItem.ListSubItems.item(1).Bold = False
            lv_articulos.selectedItem.ListSubItems.item(2).Bold = False
            lv_articulos.selectedItem.ListSubItems.item(3).Bold = False
            lv_articulos.selectedItem.ListSubItems.item(4).Bold = False
            lv_articulos.selectedItem.ForeColor = &H0&
            lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
            lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
            lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
            lv_articulos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
         End If
      Else
         lv_articulos.selectedItem.SubItems(4) = "*"
         lv_articulos.selectedItem.Bold = True
         lv_articulos.selectedItem.ListSubItems.item(1).Bold = True
         lv_articulos.selectedItem.ListSubItems.item(2).Bold = True
         lv_articulos.selectedItem.ListSubItems.item(3).Bold = True
         lv_articulos.selectedItem.ListSubItems.item(4).Bold = True
         lv_articulos.selectedItem.ForeColor = &HFF0000
         lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &HFF0000
         lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &HFF0000
         lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &HFF0000
         lv_articulos.selectedItem.ListSubItems.item(4).ForeColor = &HFF0000
      End If
   End If
End Sub

Private Sub lv_causas_KeyPress(KeyAscii As Integer)
   Dim i As Integer
   Dim j As Integer
   Dim list_item As ListItem
   If KeyAscii = 27 Then
      frm_causas.Visible = False
   End If
   If KeyAscii = 13 Then
      If lv_causas.selectedItem.SubItems(3) = "*" Then
         lv_causas.selectedItem.SubItems(3) = ""
         lv_causas.selectedItem.Bold = False
         lv_causas.selectedItem.ListSubItems.item(1).Bold = False
         lv_causas.selectedItem.ListSubItems.item(2).Bold = False
         lv_causas.selectedItem.ListSubItems.item(3).Bold = False
         lv_causas.selectedItem.ForeColor = &H0&
         lv_causas.selectedItem.ListSubItems.item(1).ForeColor = &H0&
         lv_causas.selectedItem.ListSubItems.item(2).ForeColor = &H0&
         lv_causas.selectedItem.ListSubItems.item(3).ForeColor = &H0&
      Else
         If lv_causas.selectedItem = "19" Then
            j = lv_causas.ListItems.Count
            For i = 1 To j
               lv_causas.ListItems.item(i).Selected = True
               If lv_causas.selectedItem <> "19" Then
                  lv_causas.selectedItem.Bold = False
                  lv_causas.selectedItem.ListSubItems.item(1).Bold = False
                  lv_causas.selectedItem.ListSubItems.item(2).Bold = False
                  lv_causas.selectedItem.ForeColor = &H0&
                  lv_causas.selectedItem.ListSubItems.item(1).ForeColor = &H0&
                  lv_causas.selectedItem.ListSubItems.item(2).ForeColor = &H0&
               Else
                  lv_causas.selectedItem.SubItems(3) = "*"
                  lv_causas.selectedItem.Bold = True
                  lv_causas.selectedItem.ListSubItems.item(1).Bold = True
                  lv_causas.selectedItem.ListSubItems.item(2).Bold = True
                  lv_causas.selectedItem.ListSubItems.item(3).Bold = True
                  lv_causas.selectedItem.ForeColor = &HFF00&
                  lv_causas.selectedItem.ListSubItems.item(1).ForeColor = &HFF00&
                  lv_causas.selectedItem.ListSubItems.item(2).ForeColor = &HFF00&
                  lv_causas.selectedItem.ListSubItems.item(3).ForeColor = &HFF00&
               End If
            Next i
            lv_causas_rechazo.ListItems.Clear
            rs.Open "select * from tb_causas_rechazo order by vcha_cre_nombre", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
               Set list_item = lv_causas_rechazo.ListItems.Add(, , rs!inte_cRe_causa_id)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_cRe_nombre), "", rs!vcha_cRe_nombre)
               list_item.SubItems(2) = ""
               rs.MoveNext
            Wend
            rs.Close
            frm_causas_rechazo.Visible = True
            lv_causas_rechazo.SetFocus
         Else
            lv_causas.selectedItem.SubItems(3) = "*"
            lv_causas.selectedItem.Bold = True
            lv_causas.selectedItem.ListSubItems.item(1).Bold = True
            lv_causas.selectedItem.ListSubItems.item(2).Bold = True
            lv_causas.selectedItem.ListSubItems.item(3).Bold = True
            lv_causas.selectedItem.ForeColor = &HFF00&
            lv_causas.selectedItem.ListSubItems.item(1).ForeColor = &HFF00&
            lv_causas.selectedItem.ListSubItems.item(2).ForeColor = &HFF00&
            lv_causas.selectedItem.ListSubItems.item(3).ForeColor = &HFF00&
         End If
      End If
   End If
End Sub

Private Sub lv_causas_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   frm_causas.Visible = False
End Sub

Private Sub lv_causas_rechazo_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Seleccione las causas de rechazo con ENTER y presione F8 para aplicarlas"
End Sub

Private Sub lv_causas_rechazo_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim l As Integer
   Set TB_DETALLE_DEVOLUCION_REAL_I = New TB_DETALLE_DEVOLUCION_REAL_I
   Set TB_DETALLE_DEVOLUCION_RECHAZO_I = New TB_DETALLE_DEVOLUCION_RECHAZO_I
   If KeyCode = 119 Then
      var_tipo_detalle_devolucion = 0
      j = lv_articulos.ListItems.Count
      l = lv_causas_rechazo.ListItems.Count
      For i = 1 To j
         lv_articulos.ListItems.item(i).Selected = True
         If lv_articulos.selectedItem.SubItems(4) = "*" Then
            var_almacen_Destino = Trim(lv_causas.selectedItem.SubItems(2))
            rs.Open "delete from tb_detalle_devolucion_real where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_art_articulo_id = '" + lv_articulos.selectedItem + "' and inte_cde_consecutivo  = " + lv_articulos.selectedItem.SubItems(2), cnn, adOpenDynamic, adLockOptimistic
            var_modifica = TB_DETALLE_DEVOLUCION_REAL_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_movimiento, txt_numero, lv_articulos.selectedItem, lv_articulos.selectedItem.SubItems(2), 19, var_requiere_factura)
            lv_articulos.selectedItem.SubItems(3) = var_almacen_Destino
            lv_articulos.selectedItem.SubItems(4) = ""
            If lv_articulos.selectedItem.SubItems(5) = "*" Then
               lv_articulos.selectedItem.Bold = True
               lv_articulos.selectedItem.ListSubItems.item(1).Bold = True
               lv_articulos.selectedItem.ListSubItems.item(2).Bold = True
               lv_articulos.selectedItem.ListSubItems.item(3).Bold = True
               lv_articulos.selectedItem.ForeColor = &HFF&
               lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &HFF&
               lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &HFF&
               lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &HFF&
            Else
               lv_articulos.selectedItem.Bold = False
               lv_articulos.selectedItem.ListSubItems.item(1).Bold = False
               lv_articulos.selectedItem.ListSubItems.item(2).Bold = False
               lv_articulos.selectedItem.ListSubItems.item(3).Bold = False
               lv_articulos.selectedItem.ListSubItems.item(4).Bold = False
               lv_articulos.selectedItem.ForeColor = &H0&
               lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
               lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
               lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
               lv_articulos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
            End If
            For k = 1 To l
               lv_causas_rechazo.ListItems.item(k).Selected = True
               If lv_causas_rechazo.selectedItem.SubItems(2) = "*" Then
                  var_modifica = TB_DETALLE_DEVOLUCION_RECHAZO_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_movimiento, txt_numero, lv_articulos.selectedItem, lv_articulos.selectedItem.SubItems(2), lv_causas_rechazo.selectedItem)
               End If
            Next k
         End If
      Next i
      frm_causas.Visible = False
      frm_causas_rechazo.Visible = False
      lv_articulos.SetFocus
      Call detalle_causas
   End If
End Sub

Private Sub lv_causas_rechazo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim i As Integer
      i = lv_causas_rechazo.selectedItem.Index
      If lv_causas_rechazo.selectedItem.SubItems(2) = "*" Then
         lv_causas_rechazo.selectedItem.SubItems(2) = ""
         lv_causas_rechazo.selectedItem.Bold = False
         lv_causas_rechazo.selectedItem.ListSubItems.item(1).Bold = False
         lv_causas_rechazo.selectedItem.ListSubItems.item(2).Bold = False
         lv_causas_rechazo.selectedItem.ForeColor = &H0&
         lv_causas_rechazo.selectedItem.ListSubItems.item(1).ForeColor = &H0&
         lv_causas_rechazo.selectedItem.ListSubItems.item(2).ForeColor = &H0&
      Else
         lv_causas_rechazo.selectedItem.SubItems(2) = "*"
         lv_causas_rechazo.selectedItem.Bold = True
         lv_causas_rechazo.selectedItem.ListSubItems.item(1).Bold = True
         lv_causas_rechazo.selectedItem.ListSubItems.item(2).Bold = True
         lv_causas_rechazo.selectedItem.ForeColor = &HC0&
         lv_causas_rechazo.selectedItem.ListSubItems.item(1).ForeColor = &HC0&
         lv_causas_rechazo.selectedItem.ListSubItems.item(2).ForeColor = &HC0&
      End If
   End If
   If KeyAscii = 27 Then
      frm_causas_rechazo.Visible = False
   End If
End Sub

Private Sub lv_causas_rechazo_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub lv_detalle_cliente_GotFocus()
   Me.frmlotes.Visible = False
End Sub

Private Sub lv_detalle_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      si = MsgBox("¿Deseas eliminar la causa de devolución del cliente?", vbYesNo, "ATENCION")
      If si = 6 Then
         j = lv_articulos.ListItems.Count
         For i = 1 To j
            lv_articulos.ListItems.item(i).Selected = True
            If lv_articulos.selectedItem.SubItems(4) = "*" Then
               var_tipo_detalle_devolucion = 1
               Set TB_DETALLE_DEVOLUCION_CLIENTE_I = New TB_DETALLE_DEVOLUCION_CLIENTE_I
               var_modifica = TB_DETALLE_DEVOLUCION_CLIENTE_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_movimiento, txt_numero, lv_articulos.selectedItem, lv_articulos.selectedItem.SubItems(2), lv_detalle_cliente.selectedItem)
               If lv_articulos.selectedItem.SubItems(5) = "*" Then
                  lv_articulos.selectedItem.Bold = True
                  lv_articulos.selectedItem.ListSubItems.item(1).Bold = True
                  lv_articulos.selectedItem.ListSubItems.item(2).Bold = True
                  lv_articulos.selectedItem.ListSubItems.item(3).Bold = True
                  lv_articulos.selectedItem.ForeColor = &HFF&
                  lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &HFF&
                  lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &HFF&
                  lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &HFF&
               Else
                  lv_articulos.selectedItem.Bold = False
                  lv_articulos.selectedItem.ListSubItems.item(1).Bold = False
                  lv_articulos.selectedItem.ListSubItems.item(2).Bold = False
                  lv_articulos.selectedItem.ForeColor = &H0&
                  lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
                  lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
               End If
            End If
         Next i
         Call detalle_causas
      End If
   End If
End Sub

Private Sub lv_detalle_correciones_GotFocus()
   Me.frmlotes.Visible = False
End Sub

Private Sub lv_detalle_correciones_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      si = MsgBox("¿Deseas eliminar la corrección que se a echo?", vbYesNo, "ATENCION")
      If si = 6 Then
         j = lv_articulos.ListItems.Count
         For i = 1 To j
            lv_articulos.ListItems.item(i).Selected = True
            If lv_articulos.selectedItem.SubItems(4) = "*" Then
               var_tipo_detalle_devolucion = 1
               Set TB_DETALLE_DEVOLUCION_AJUSTE = New TB_DETALLE_DEVOLUCION_AJUSTE
               var_modifica = TB_DETALLE_DEVOLUCION_AJUSTE.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_movimiento, txt_numero, lv_articulos.selectedItem, lv_articulos.selectedItem.SubItems(2), lv_detalle_correciones.selectedItem)
               If lv_articulos.selectedItem.SubItems(5) = "*" Then
                  lv_articulos.selectedItem.Bold = True
                  lv_articulos.selectedItem.ListSubItems.item(1).Bold = True
                  lv_articulos.selectedItem.ListSubItems.item(2).Bold = True
                  lv_articulos.selectedItem.ListSubItems.item(3).Bold = True
                  lv_articulos.selectedItem.ForeColor = &HFF&
                  lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &HFF&
                  lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &HFF&
                  lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &HFF&
               Else
                  lv_articulos.selectedItem.Bold = False
                  lv_articulos.selectedItem.ListSubItems.item(1).Bold = False
                  lv_articulos.selectedItem.ListSubItems.item(2).Bold = False
                  lv_articulos.selectedItem.ForeColor = &H0&
                  lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
                  lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
               End If
            End If
         Next i
         Call detalle_causas
      End If
   End If
End Sub

Private Sub lv_detalle_real_GotFocus()
   Me.frmlotes.Visible = False
End Sub

Private Sub lv_detalle_real_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      si = MsgBox("¿Deseas eliminar la causa de devolución del cliente?", vbYesNo, "ATENCION")
      If si = 6 Then
         j = lv_articulos.ListItems.Count
         For i = 1 To j
            lv_articulos.ListItems.item(i).Selected = True
            If lv_articulos.selectedItem.SubItems(4) = "*" Then
               var_tipo_detalle_devolucion = 1
               rs.Open "delete from tb_detalle_devolucion_rechazo where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_EMO_NUMERO = " + Me.txt_numero + " and vcha_art_articulo_id = '" + lv_articulos.selectedItem + "' and INTE_CDE_CONSECUTIVO = " + lv_articulos.selectedItem.SubItems(2), cnn, adOpenDynamic, adLockOptimistic
               Set TB_DETALLE_DEVOLUCION_REAL_I = New TB_DETALLE_DEVOLUCION_REAL_I
               var_modifica = TB_DETALLE_DEVOLUCION_REAL_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_movimiento, txt_numero, lv_articulos.selectedItem, lv_articulos.selectedItem.SubItems(2), lv_detalle_real.selectedItem, var_requiere_factura)
               If lv_articulos.selectedItem.SubItems(5) = "*" Then
                  lv_articulos.selectedItem.Bold = True
                  lv_articulos.selectedItem.ListSubItems.item(1).Bold = True
                  lv_articulos.selectedItem.ListSubItems.item(2).Bold = True
                  lv_articulos.selectedItem.ListSubItems.item(3).Bold = True
                  lv_articulos.selectedItem.ForeColor = &HFF&
                  lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &HFF&
                  lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &HFF&
                  lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &HFF&
               Else
                  lv_articulos.selectedItem.Bold = False
                  lv_articulos.selectedItem.ListSubItems.item(1).Bold = False
                  lv_articulos.selectedItem.ListSubItems.item(2).Bold = False
                  lv_articulos.selectedItem.ForeColor = &H0&
                  lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
                  lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
                  lv_articulos.selectedItem.SubItems(7) = lv_articulos.selectedItem.SubItems(7) - 1
               End If
            End If
         Next i
         Call detalle_causas
      End If
   End If
End Sub


Private Sub lv_detalle_rechazo_GotFocus()
   Me.frmlotes.Visible = False
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_movimiento = lv_lista.selectedItem
         txt_nombre_movimiento = lv_lista.selectedItem.SubItems(1)
      Else
         txt_movimiento = ""
         txt_nombre_movimiento = ""
      End If
      If Me.txt_movimiento.Enabled = True Then
         txt_movimiento.SetFocus
      End If
      frm_lista.Visible = False
      var_ventana = 0
   End If
   If KeyAscii = 27 Then
      var_ventana = 0
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_factura_GotFocus()
   Me.frmlotes.Visible = False
End Sub

Private Sub txt_factura_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_movimiento.Enabled = True Then
         Me.txt_movimiento.SetFocus
      End If
   End If
End Sub

Private Sub txt_lote_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(Me.txt_lote) = "" Then
         var_si = MsgBox("¿Desea eliminar el lote del producto?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            Me.txt_lote = ""
            Me.txt_supervisor = ""
            Me.txt_fecha_lote = ""
            Me.txt_nota_envio = ""
            Me.txt_nota_envio = ""
            Me.cmb_tipo_defecto = ""
            Me.txt_proveedor = ""
            Me.cmb_justifica = ""
            Me.cmb_estatus = ""
         End If
      Else
         If IsNumeric(Me.txt_lote) Then
            If Me.txt_lote = 0 Then
               Me.txt_supervisor = ""
               Me.txt_fecha_lote = ""
               Me.txt_nota_envio = ""
            Else
               If rs.State = 1 Then
                  rs.Close
               End If
               var_codigo = Me.lv_articulos.selectedItem
               If rs.State = 1 Then
                  rs.Close
               End If
               rs.Open "select * from vw_catalogo_articulos where vcha_art_articulo_id = '" + Me.lv_articulos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_conexion = IIf(IsNull(rs!vcha_uor_conexion), "", rs!vcha_uor_conexion)
                  'MsgBox var_conexion
                  If var_conexion <> "" Then
                     Dim cnn_lote As ADODB.Connection
                     Set cnn_lote = CreateObject("ADODB.connection")
                     'MsgBox var_conexion
                     cnn_lote.Open var_conexion
                     'MsgBox var_conexion
                     cnn_lote.CursorLocation = adUseClient
                     cnn_lote.CommandTimeout = 3000
                     var_cadena = "SELECT SUBSTRING(L.VCHA_LOT_LOTE_ID, 3, LEN(L.VCHA_LOT_LOTE_ID) - 2) LOTE, M.VCHA_MOD_RESPONSABLE SUPERVISOR,"
                     var_cadena = var_cadena + " L.VCHA_LOT_FECHATER FECHA_TERMINO, L.VCHA_LOT_NOTAENVIO NOTA FROM TB_LOTES L, TB_MODULOS M Where m.BINT_MOD_MODULO_ID = l.BINT_MOD_MODULO_ID and SUBSTRING(L.VCHA_LOT_LOTE_ID, 3, LEN(L.VCHA_LOT_LOTE_ID) - 2) = '" + Me.txt_lote + "'"
                     rsaux.Open var_cadena, cnn_lote, adOpenDynamic, adLockOptimistic
                     'MsgBox cnn_lote.ConnectionString
                     If Not rsaux.EOF Then
                        lote = rsaux!lote
                        supervisor = rsaux!supervisor
                        fecha_lote = rsaux!fecha_termino
                        Me.txt_lote = lote
                        Me.txt_supervisor = supervisor
                        Me.txt_fecha_lote = fecha_lote
                        Me.txt_nota_envio = IIf(IsNull(rsaux!nota), "", rsaux!nota)
                     Else
                        MsgBox "El lote no se encontro", vbOKOnly, "ATENCION"
                     End If
                     rsaux.Close
                     Set cnn_lote = Nothing
                  Else
                     MsgBox "No se puede conectar a la planta correspondiente", vbOKOnly, "ATENCION"
                   End If
               End If
               rs.Close
            End If
         Else
            MsgBox "Número de lote incorrecto", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   Call pro_enfoque(KeyAscii)
   If KeyAscii = 27 Then
      frmlotes.Visible = False
   End If
End Sub

Private Sub txt_movimiento_GotFocus()
   Me.frmlotes.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_movimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 1
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_movimientos where INTE_MOV_CAUSA_DEVOLUCION = 1 and (char_mov_afectacion <> 'T') or vcha_mov_movimiento_id = 'TEC'  or vcha_mov_movimiento_id = 'T'", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Movimientos"
      var_tipo_lista = 1
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_movimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      txt_nombre_movimiento.SetFocus
   End If
End Sub

Private Sub txt_movimiento_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_movimiento) <> "" Then
      rs.Open "select * from tb_movimientos where (INTE_MOV_CAUSA_DEVOLUCION = 1 and (char_mov_afectacion <> 'T') or vcha_mov_movimiento_id = 'TEC'  or vcha_mov_movimiento_id = 'T') and vcha_mov_movimiento_id ='" + txt_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_movimiento = rs!vcha_mov_nombre
         rs.Close
         txt_numero.Enabled = True
         txt_movimiento.Enabled = False
      Else
         txt_movimiento = ""
         txt_nombre_movimiento = ""
         rs.Close
         MsgBox "Clave de movimiento incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_nombre_movimiento_GotFocus()
   Me.frmlotes.Visible = False
End Sub

Private Sub txt_nombre_movimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If txt_movimiento.Enabled = True Then
      If KeyCode = 116 Then
         var_ventana = 1
         lv_lista.ListItems.Clear
         rs.Open "select * from tb_movimientos where INTE_MOV_CAUSA_DEVOLUCION = 1 and char_mov_afectacion <> 'T'", cnn, adOpenDynamic, adLockBatchOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "Movimientos"
         var_tipo_lista = 1
         frm_lista.Visible = True
         lv_lista.SetFocus
      End If
   End If
End Sub

Private Sub txt_nombre_movimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If txt_numero.Enabled = True Then
      txt_numero.SetFocus
   End If
End Sub

Private Sub txt_numero_GotFocus()
   Me.frmlotes.Visible = False
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      cnn.CommandTimeout = 10000
      Dim var_posible_factura As Boolean
      var_posible_factura = False
      If chk_factura.Value = 0 Then
         var_posible_factura = True
      Else
         If Not IsNumeric(txt_factura) Then
            MsgBox "Número de factura incorrecto", vbOKOnly, "ATENCION"
            var_posible_factura = False
         Else
            rs.Open "select * from tb_encabezado_Cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Car_documento = 'FA' and inte_car_numero =  " + txt_factura, cnn, adOpenDynamic, adLockOptimistic
            If rs.EOF Then
               MsgBox "La factura no existe", vbOKOnly, "ATENCION"
               var_posible_factura = False
            Else
               var_serie_FACTURA = rs!vcha_Ser_Serie_id
               var_cliente_factura = rs!vcha_cli_clave_id
            End If
            rs.Close
            rs.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and char_emo_estatus = 'I' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' ", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_referencia = rs!vcha_Emo_referencia
               var_cliente = rs!vcha_cli_clave_id
               var_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
               var_titular = rs!vcha_tit_titular_id
            End If
            rs.Close
            If var_cliente_factura = var_cliente Then
               var_posible_factura = True
            Else
               MsgBox "La factura seleccionada no corresponde al cliente de la devolución", vbOKOnly, "ATENCION"
               var_posible_factura = False
            End If
         End If
      End If
      If var_posible_factura = True Then
         If Trim(txt_numero) <> "" Then
            lv_articulos.ListItems.Clear
            Dim var_cantidad_pasar As Double
            Dim var_factura As Double
            Dim var_posible As Boolean
            Dim var_consecutivo As Integer
            Dim var_contador_articulos As Double
            Dim var_tipo_busqueda As Integer
            Dim var_grupo As String
            Dim list_item As ListItem
            Dim var_contador As Double
            Dim var_cantidad As Double
            Dim var_nombre_articulo As String
            Dim var_nombre_causa_1 As String
            Dim var_nombre_causa_2 As String
            Dim var_descuento_1 As Double
            Dim var_descuento_2 As Double
            Dim var_descuento_3 As Double
            Dim var_iva As Double
            Dim var_precio As Double
            Dim var_precio_anterior As Double
            Dim var_costo As Double
            Dim var_clave_moneda As String
            Dim var_tipo_cambio_anterior As Double
            Dim var_tipo_Cambio As Double
            Dim var_moneda_local As Integer
            Dim var_numero_factura As Double
            Dim var_lista_precios As String
            Dim var_descuento_volumen As Double
            Dim var_descuento_financiero As Double
            Dim var_descuento_pago As Double
            Dim var_canal_venta As String
            Dim var_iva_canal As Double
            Dim var_almacen_costeo As String
            'Dim var_serie_factura As String
            var_posible = False
            rs.Open "select * from tb_movimientos where vcha_mov_movimiento_id = '" + txt_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            var_requiere_factura = 0
            If Not rs.EOF Then
               var_requiere_factura = IIf(IsNull(rs!INTE_MOV_DEVOLUCION_FACTURA), 0, rs!INTE_MOV_DEVOLUCION_FACTURA)
            End If
            rs.Close
            rs.Open "select * from tb_almacenes where vcha_emp_empresa_id = '" + var_empresa + "' and inte_alm_costeo = 1", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_almacen_costeo = rs!VCHA_ALM_ALMACEN_ID
            End If
            rs.Close
            rs.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and char_emo_estatus = 'I' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux4.Open "select sum(floa_ent_cantidad) from tb_entradas where vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_ent_numero = " + txt_numero + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organziacional + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                  Me.lbl_movimiento = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
               Else
                  Me.lbl_movimiento = "0"
               End If
               rsaux4.Close
               var_referencia = rs!vcha_Emo_referencia
               var_cliente = rs!vcha_cli_clave_id
               var_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
               var_titular = rs!vcha_tit_titular_id
               var_tipo_busqueda = 1
               
               rsaux2.Open "select * from vw_clientes where vcha_cli_clave_id ='" + var_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  var_tipo_busqueda = IIf(IsNull(rsaux2!INTE_CAN_BUSQUEDA_FACTURA_GRUPO), 1, rsaux2!INTE_CAN_BUSQUEDA_FACTURA_GRUPO)
                  var_grupo = IIf(IsNull(rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID)
                  var_lista_precios = IIf(IsNull(rsaux2!vcha_LIS_LISTA_iD), "", rsaux2!vcha_LIS_LISTA_iD)
                  var_canal_venta = IIf(IsNull(rsaux2!vcha_can_canal_venta_id), "", rsaux2!vcha_can_canal_venta_id)
                  var_iva_canal = IIf(IsNull(rsaux2!FLOA_TPE_IVA), 0, rsaux2!FLOA_TPE_IVA)
               Else
                  var_grupo = ""
                  var_tipo_busqueda = 1
                  var_lista_precios = ""
                  var_canal_venta = ""
                  var_iva_canal = 0
               End If
               rsaux2.Close
               
               If var_canal_venta <> "" Then
                  rsaux2.Open "select floa_gac_descuento_1, floa_gac_descuento_2 from tb_gruposactuales where vcha_gac_grupo_Actual_id = '" + var_grupo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     var_descuento_volumen = IIf(IsNull(rsaux2!floa_gac_Descuento_1), 0, rsaux2!floa_gac_Descuento_1)
                     var_descuento_pago = IIf(IsNull(rsaux2!FLOA_GAC_DESCUENTO_2), 0, rsaux2!FLOA_GAC_DESCUENTO_2)
                  Else
                     var_descuento_volumen = 0
                     var_descuento_pago = 0
                  End If
                  rsaux2.Close
               Else
                  var_descuento_volumen = 0
                  var_descuento_financiero = 0
                  var_descuento_pago = 0
               End If
               rs.Close
               
               rs.Open "select * from tb_devoluciones where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
               If rs.EOF Then
                  Set TB_DEVOLUCIONES_INSERTA = New TB_DEVOLUCIONES_INSERTA
                  If rs.State = 1 Then
                     rs.Close
                  End If
                  rs.Open "select * from tb_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and  vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_ent_numero = " + txt_numero + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_almacen = rs!VCHA_ALM_ALMACEN_ID
                     var_consecutivo = 0
                      While Not rs.EOF
                           var_descuento_1 = 0
                           var_descuento_2 = 0
                           var_descuento_3 = 0
                           var_precio = 0
                           var_costo = 0
                           var_iva = 0
                           var_factura = 0
                           var_costo = rs!floa_ent_costo
                           var_año = rs!inte_ent_año
                           If var_requiere_factura = 1 Then
                              
                              'var_serie_factura = ""
                              If var_empresa = "02" Or var_empresa = "18" Or var_empresa = "31" Then
                                 If Me.chk_factura = 1 Then
                                    var_numero_factura = CDbl(txt_factura)
                                    If var_numero_factura > 0 Then
                                       rsaux2.Open "select * from VW_ORDEN_FECHAS_FACTURAS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_art_articulo_id = '" + rs!vcha_Art_Articulo_id + "' and inte_car_numero = " + Str(var_numero_factura) + " and vcha_Ser_Serie_id = '" + var_serie_FACTURA + "' order by dtim_car_fecha desc", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux2.EOF Then
                                          var_clave_moneda = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
                                          var_moneda_local = IIf(IsNull(rsaux2!inte_mon_moneda_local), 0, rsaux2!inte_mon_moneda_local)
                                          If var_moneda_local = 1 Then
                                             var_tipo_Cambio = 1
                                             var_precio = IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio)
                                             var_descuento_1 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_1), 0, rsaux2!FLOA_SAL_DESCUENTO_1)
                                             var_descuento_2 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_2), 0, rsaux2!FLOA_SAL_DESCUENTO_2)
                                             var_iva = IIf(IsNull(rsaux2!floa_car_porcentaje_iva), 0, rsaux2!floa_car_porcentaje_iva)
                                             var_factura = IIf(IsNull(rsaux2!inte_Car_numero), 0, rsaux2!inte_Car_numero)
                                             rsaux4.Open "SELECT max(FLOA_RCO_DESCUENTO_APLICAR) FROM TB_rELACION_COBRANZA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO =  " + CStr(var_factura) + " AND VCHA_CAR_DOCUMENTO= 'FA' and vcha_Ser_serie_id = '" + rsaux2!vcha_Ser_Serie_id + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                             If Not rsaux4.EOF Then
                                                var_descuento_3 = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                             Else
                                                var_descuento_3 = 0
                                             End If
                                             rsaux4.Close
                                             var_serie_FACTURA = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
                                          Else
                                             rsaux3.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                             If Not rsaux3.EOF Then
                                                var_tipo_Cambio = IIf(IsNull(rsaux3!mone_tca_importe), 1, rsaux3!mone_tca_importe)
                                                var_tipo_cambio_anterior = IIf(IsNull(rsaux2!floa_car_tipo_cambio), 1, rsaux2!floa_car_tipo_cambio)
                                                var_precio = IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio)
                                                var_precio_anterior = var_precio / var_tipo_cambio_anterior
                                                var_precio = var_precio_anterior * var_tipo_Cambio
                                                var_descuento_1 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_1), 0, rsaux2!FLOA_SAL_DESCUENTO_1)
                                                var_descuento_2 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_2), 0, rsaux2!FLOA_SAL_DESCUENTO_2)
                                                rsaux4.Open "SELECT max(FLOA_RCO_DESCUENTO_APLICAR) FROM TB_rELACION_COBRANZA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO =  " + CStr(var_factura) + " AND VCHA_CAR_DOCUMENTO= 'FA' and vcha_Ser_serie_id = '" + rsaux2!vcha_Ser_Serie_id + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                                If Not rsaux4.EOF Then
                                                   var_descuento_3 = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                                Else
                                                   var_descuento_3 = 0
                                                End If
                                                rsaux4.Close
                                                var_iva = IIf(IsNull(rsaux2!floa_car_porcentaje_iva), 0, rsaux2!floa_car_porcentaje_iva)
                                                var_factura = IIf(IsNull(rsaux2!inte_Car_numero), 0, rsaux2!inte_Car_numero)
                                                var_serie_FACTURA = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
                                             Else
                                                GoTo salir:
                                             End If
                                             rsaux3.Close
                                          End If
                                       End If
                                       rsaux2.Close
                                    End If
                                 Else
                                    ''' proceso normal
                                    var_serie_FACTURA = ""
                                    
                                    'If var_tipo_busqueda = 1 And Trim(var_grupo) <> "" Then
                                    '   rsaux3.Open "select * from VW_ORDEN_FACTURAS_FECHA_GRUPO where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_gac_grupo_actual_id ='" + var_grupo + "' and vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    '   If Not rsaux3.EOF Then
                                    '      var_numero_factura = rsaux3!inte_car_numero
                                    '      var_serie_factura = rsaux3!vcha_ser_serie_id
                                    '   End If
                                    '   rsaux3.Close
                                    'Else
                                    '   rsaux3.Open "select * from VW_ORDEN_FECHA_FACTURAS_TITULAR where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_tit_titular_id ='" + var_titular + "' and vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    '   If Not rsaux3.EOF Then
                                    '      var_numero_factura = rsuax3!inte_car_numero
                                    '      var_serie_factura = rsaux3!vcha_ser_serie_id
                                    '   End If
                                    '   rsaux3.Close
                                    'End If
                                    var_numero_factura = 1
                                    If var_numero_factura > 0 Then
                                       var_numero_factura = 0
                                       'rsaux2.Open "select * from VW_ORDEN_FECHAS_FACTURAS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "' and inte_car_numero = " + Str(var_numero_factura) + " and vcha_Ser_Serie_id = '" + var_serie_factura + "'", cnn, adOpenDynamic, adLockOptimistic
                                       '1
                                       If rsaux2.State = 1 Then
                                          rsaux2.Close
                                       End If
                                       var_cadena = "SELECT TOP 1 dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_PORCENTAJE_IVA, "
                                       var_cadena = var_cadena + " dbo.TB_SALIDAS.vcha_ser_serie_id,  dbo.TB_SALIDAS.INTE_CAR_NUMERO, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2, dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_MONEDAS.VCHA_MON_MONEDA_ID, dbo.TB_MONEDAS.INTE_MON_MONEDA_LOCAL FROM dbo.TB_SALIDAS INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALIDAS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO AND dbo.TB_SALIDAS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID INNER JOIN dbo.TB_MONEDAS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_MON_MONEDA_ID = dbo.TB_MONEDAS.VCHA_MON_MONEDA_ID WHERE (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND "
                                       var_cadena = var_cadena + " (dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID = '" + var_titular + "') AND (dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = '" + rs!vcha_Art_Articulo_id + "') ORDER BY dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA DESC"
                                       'MsgBox " (dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID = '" + var_titular + "') AND (dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = '" + rs!vcha_Art_Articulo_id + "') ORDER BY dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA DESC"
                                        rsaux2.Open var_cadena
                                       'rsaux2.Open "select * from VW_ORDEN_FECHAS_FACTURAS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_art_articulo_id = '" + rs!vcha_Art_Articulo_id + "'  AND VCHA_TIT_TITULAR_ID = '" + var_titular + "'  order by dtim_car_fecha desc", cnn, adOpenDynamic, adLockOptimistic
                                       
                                       
                                       
                                       If Not rsaux2.EOF Then
                                          
                                          var_numero_factura = IIf(IsNull(rsaux2!inte_Car_numero), 0, rsaux2!inte_Car_numero)
                                          var_serie_FACTURA = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
                                          var_clave_moneda = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
                                          var_moneda_local = IIf(IsNull(rsaux2!inte_mon_moneda_local), 0, rsaux2!inte_mon_moneda_local)
                                          If var_moneda_local = 1 Then
                                             var_tipo_Cambio = 1
                                             var_precio = IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio)
                                             var_descuento_1 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_1), 0, rsaux2!FLOA_SAL_DESCUENTO_1)
                                             var_descuento_2 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_2), 0, rsaux2!FLOA_SAL_DESCUENTO_2)
                                             var_iva = IIf(IsNull(rsaux2!floa_car_porcentaje_iva), 0, rsaux2!floa_car_porcentaje_iva)
                                             var_factura = IIf(IsNull(rsaux2!inte_Car_numero), 0, rsaux2!inte_Car_numero)
                                             cnn.CommandTimeout = 360
                                             If rsaux4.State = 1 Then
                                                rsaux4.Close
                                             End If
                                             rsaux4.Open "SELECT max(FLOA_RCO_DESCUENTO_APLICAR) FROM TB_rELACION_COBRANZA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO =  " + CStr(var_factura) + " AND VCHA_CAR_DOCUMENTO= 'FA' and vcha_Ser_serie_id = '" + rsaux2!vcha_Ser_Serie_id + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                             If Not rsaux4.EOF Then
                                                var_descuento_3 = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                             Else
                                                var_descuento_3 = 0
                                             End If
                                             rsaux4.Close
                                             var_serie_FACTURA = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
                                           Else
                                             rsaux3.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                             If Not rsaux3.EOF Then
                                                var_tipo_Cambio = IIf(IsNull(rsaux3!mone_tca_importe), 1, rsaux3!mone_tca_importe)
                                                var_tipo_cambio_anterior = IIf(IsNull(rsaux2!floa_car_tipo_cambio), 1, rsaux2!floa_car_tipo_cambio)
                                                var_precio = IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio)
                                                var_precio_anterior = var_precio / var_tipo_cambio_anterior
                                                var_precio = var_precio_anterior * var_tipo_Cambio
                                                var_descuento_1 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_1), 0, rsaux2!FLOA_SAL_DESCUENTO_1)
                                                var_descuento_2 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_2), 0, rsaux2!FLOA_SAL_DESCUENTO_2)
                                                rsaux4.Open "SELECT max(FLOA_RCO_DESCUENTO_APLICAR) FROM TB_rELACION_COBRANZA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO =  " + CStr(var_factura) + " AND VCHA_CAR_DOCUMENTO= 'FA' and vcha_Ser_serie_id = '" + rsaux2!vcha_Ser_Serie_id + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                                If Not rsaux4.EOF Then
                                                   var_descuento_3 = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                                Else
                                                   var_descuento_3 = 0
                                                End If
                                                rsaux4.Close
                                                var_iva = IIf(IsNull(rsaux2!floa_car_porcentaje_iva), 0, rsaux2!floa_car_porcentaje_iva)
                                                var_factura = IIf(IsNull(rsaux2!inte_Car_numero), 0, rsaux2!inte_Car_numero)
                                                var_serie_FACTURA = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
                                             Else
                                                GoTo salir:
                                             End If
                                             rsaux3.Close
                                          End If
                                       Else
                                          rsaux9.Open "select isnull(floa_Gac_descuento_1,0), isnull(floa_gac_descuento_2,0) from tb_gruposactuales where vcha_gac_grupo_actual_id = '" + var_grupo + "'", cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux9.EOF Then
                                             var_descuento_1 = rsaux9(0).Value
                                             var_descuento_2 = rsaux9(1).Value
                                             
                                          End If
                                          rsaux9.Close
                                       
                                       End If
                                       rsaux2.Close
                                    Else
                                       If Trim(var_lista_precios) <> "" Then
                                          rsaux3.Open "select * from vw_detalle_lista_precios where vcha_lis_lista_precios_id = '" + var_lista_precios + "' and vcha_Art_Articulo_id = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux3.EOF Then
                                             var_clave_moneda = IIf(IsNull(rsaux3!vcha_mon_moneda), "", rsaux3!vcha_mon_moneda)
                                             var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
                                             If Not rsaux3.EOF Then
                                                If var_moneda_local = 1 Then
                                                   var_tipo_Cambio = 1
                                                   var_precio = IIf(IsNull(rsaux3!floa_dli_precio), 0, rsaux3!floa_dli_precio)
                                                   var_descuento_1 = var_descuento_volumen
                                                   var_descuento_2 = var_descuento_pago
                                                   var_descuento_3 = var_descuento_financiero
                                                   var_iva = var_iva_canal
                                                Else
                                                   rsaux.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                                   If Not rsaux.EOF Then
                                                      var_tipo_Cambio = IIf(IsNull(rsaux!mone_tca_importe), 1, rsaux!mone_tca_importe)
                                                      var_precio = IIf(IsNull(rsaux3!floa_dli_precio), 0, rsaux3!floa_dli_precio * var_tipo_Cambio)
                                                      var_descuento_1 = var_descuento_volumen
                                                      var_descuento_2 = var_descuento_pago
                                                      var_descuento_3 = var_descuento_financiero
                                                      var_iva = var_iva_canal
                                                   Else
                                                      GoTo salir:
                                                   End If
                                                   rsaux.Close
                                                End If
                                             Else
                                                GoTo salir_3:
                                             End If
                                             rsaux3.Close
                                          Else
                                             rsaux3.Close
                                             Cadena = "SELECT dbo.TB_DETALLE_LISTA_PRECIOS.VCHA_LIS_LISTA_PRECIOS_ID, dbo.TB_DETALLE_LISTA_PRECIOS.VCHA_ART_ARTICULO_ID,   dbo.TB_DETALLE_LISTA_PRECIOS.FLOA_DLI_PRECIO, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, "
                                             Cadena = Cadena + " dbo.TB_LISTADEPRECIOS.DTIM_LIS_FECHA_INICIO , dbo.TB_LISTADEPRECIOS.DTIM_LIS_FECHA_FIN, dbo.TB_LISTADEPRECIOS.VCHA_MON_MONEDA, dbo.TB_MONEDAS.INTE_MON_MONEDA_LOCAL"
                                             Cadena = Cadena + " FROM dbo.TB_DETALLE_LISTA_PRECIOS INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_DETALLE_LISTA_PRECIOS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_LISTADEPRECIOS ON dbo.TB_DETALLE_LISTA_PRECIOS.VCHA_LIS_LISTA_PRECIOS_ID = dbo.TB_LISTADEPRECIOS.VCHA_LIS_LISTA_ID INNER JOIN dbo.TB_MONEDAS ON dbo.TB_LISTADEPRECIOS.VCHA_MON_MONEDA = dbo.TB_MONEDAS.VCHA_MON_MONEDA_ID  where dbo.TB_DETALLE_LISTA_PRECIOS.vcha_lis_lista_precios_id = '" + var_lista_precios + "' and dbo.TB_DETALLE_LISTA_PRECIOS.vcha_Art_Articulo_id = '" + rs!vcha_Art_Articulo_id + "'"
                                             rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                             If Not rsaux3.EOF Then
                                                var_clave_moneda = IIf(IsNull(rsaux3!vcha_mon_moneda), "", rsaux3!vcha_mon_moneda)
                                                var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
                                                If Not rsaux3.EOF Then
                                                   If var_moneda_local = 1 Then
                                                      var_tipo_Cambio = 1
                                                      var_precio = IIf(IsNull(rsaux3!floa_dli_precio), 0, rsaux3!floa_dli_precio)
                                                      var_descuento_1 = var_descuento_volumen
                                                      var_descuento_2 = var_descuento_pago
                                                      var_descuento_3 = var_descuento_financiero
                                                      var_iva = var_iva_canal
                                                   Else
                                                      rsaux.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                                      If Not rsaux.EOF Then
                                                         var_tipo_Cambio = IIf(IsNull(rsaux!mone_tca_importe), 1, rsaux!mone_tca_importe)
                                                         var_precio = IIf(IsNull(rsaux3!floa_dli_precio), 0, rsaux3!floa_dli_precio * var_tipo_Cambio)
                                                         var_descuento_1 = var_descuento_volumen
                                                         var_descuento_2 = var_descuento_pago
                                                         var_descuento_3 = var_descuento_financiero
                                                         var_iva = var_iva_canal
                                                      Else
                                                         GoTo salir:
                                                      End If
                                                      rsaux.Close
                                                   End If
                                                Else
                                                   GoTo salir_3:
                                                End If
                                             End If
                                             rsaux3.Close
                                          End If
                                       Else
                                          GoTo salir_2:
                                       End If
                                    End If
                                 End If
                              End If
                              If var_empresa = "03" Then
                                 If Me.chk_factura = 1 Then
                                    var_numero_factura = CDbl(txt_factura)
                                    If var_numero_factura > 0 Then
                                       rsaux2.Open "select * from VW_ORDEN_FECHAS_FACTURAS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_art_articulo_id = '" + rs!vcha_Art_Articulo_id + "' and inte_car_numero = " + Str(var_numero_factura) + " and vcha_Ser_Serie_id = '" + var_serie_FACTURA + "'  order by dtim_car_fecha desc", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux2.EOF Then
                                          var_clave_moneda = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
                                          var_moneda_local = IIf(IsNull(rsaux2!inte_mon_moneda_local), 0, rsaux2!inte_mon_moneda_local)
                                          If var_moneda_local = 1 Then
                                             var_tipo_Cambio = 1
                                             var_precio = IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio)
                                             var_descuento_1 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_1), 0, rsaux2!FLOA_SAL_DESCUENTO_1)
                                             var_descuento_2 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_2), 0, rsaux2!FLOA_SAL_DESCUENTO_2)
                                             var_iva = IIf(IsNull(rsaux2!floa_car_porcentaje_iva), 0, rsaux2!floa_car_porcentaje_iva)
                                             var_factura = IIf(IsNull(rsaux2!inte_Car_numero), 0, rsaux2!inte_Car_numero)
                                             rsaux4.Open "SELECT max(FLOA_RCO_DESCUENTO_APLICAR) FROM TB_rELACION_COBRANZA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO =  " + CStr(var_factura) + " AND VCHA_CAR_DOCUMENTO= 'FA' and vcha_Ser_serie_id = '" + rsaux2!vcha_Ser_Serie_id + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                             If Not rsaux4.EOF Then
                                                var_descuento_3 = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                             Else
                                                var_descuento_3 = 0
                                             End If
                                             rsaux4.Close
                                             var_serie_FACTURA = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
                                          Else
                                             rsaux3.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                             If Not rsaux3.EOF Then
                                                var_tipo_Cambio = IIf(IsNull(rsaux3!mone_tca_importe), 1, rsaux3!mone_tca_importe)
                                                var_tipo_cambio_anterior = IIf(IsNull(rsaux2!floa_car_tipo_cambio), 1, rsaux2!floa_car_tipo_cambio)
                                                var_precio = IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio)
                                                var_precio_anterior = var_precio / var_tipo_cambio_anterior
                                                var_precio = var_precio_anterior * var_tipo_Cambio
                                                var_descuento_1 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_1), 0, rsaux2!FLOA_SAL_DESCUENTO_1)
                                                var_descuento_2 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_2), 0, rsaux2!FLOA_SAL_DESCUENTO_2)
                                                rsaux4.Open "SELECT max(FLOA_RCO_DESCUENTO_APLICAR) FROM TB_rELACION_COBRANZA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO =  " + CStr(var_factura) + " AND VCHA_CAR_DOCUMENTO= 'FA' and vcha_Ser_serie_id = '" + rsaux2!vcha_Ser_Serie_id + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                                'empieza descuento 3
                                                If Not rsaux4.EOF Then
                                                   var_descuento_3 = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                                Else
                                                   var_descuento_3 = 0
                                                End If
                                                rsaux4.Close
                                                var_iva = IIf(IsNull(rsaux2!floa_car_porcentaje_iva), 0, rsaux2!floa_car_porcentaje_iva)
                                                var_factura = IIf(IsNull(rsaux2!inte_Car_numero), 0, rsaux2!inte_Car_numero)
                                                var_serie_FACTURA = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
                                             Else
                                                GoTo salir:
                                             End If
                                             rsaux3.Close
                                          End If
                                       End If
                                       rsaux2.Close
                                    End If
                                 Else
                                    If Trim(var_lista_precios) <> "" Then
                                        rsaux3.Open "select * from vw_detalle_lista_precios where vcha_lis_lista_precios_id = '" + var_lista_precios + "' and vcha_Art_Articulo_id = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                        If Not rsaux3.EOF Then
                                           var_clave_moneda = IIf(IsNull(rsaux3!vcha_mon_moneda), "", rsaux3!vcha_mon_moneda)
                                           var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
                                           If Not rsaux3.EOF Then
                                              If var_moneda_local = 1 Then
                                                 var_tipo_Cambio = 1
                                                 var_precio = IIf(IsNull(rsaux3!floa_dli_precio), 0, rsaux3!floa_dli_precio)
                                                 var_descuento_1 = var_descuento_volumen
                                                 var_descuento_2 = var_descuento_pago
                                                 var_descuento_3 = var_descuento_financiero
                                                 var_iva = var_iva_canal
                                              Else
                                                 rsaux.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                                 If Not rsaux.EOF Then
                                                    var_tipo_Cambio = IIf(IsNull(rsaux!mone_tca_importe), 1, rsaux!mone_tca_importe)
                                                    var_precio = IIf(IsNull(rsaux3!floa_dli_precio), 0, rsaux3!floa_dli_precio * var_tipo_Cambio)
                                                    var_descuento_1 = var_descuento_volumen
                                                    var_descuento_2 = var_descuento_pago
                                                    var_descuento_3 = var_descuento_financiero
                                                    var_iva = var_iva_canal
                                                 Else
                                                    GoTo salir:
                                                 End If
                                                 rsaux.Close
                                              End If
                                           Else
                                              GoTo salir_3:
                                           End If
                                           rsaux3.Close
                                        Else
                                           rsaux3.Close
                                           Cadena = "SELECT dbo.TB_DETALLE_LISTA_PRECIOS.VCHA_LIS_LISTA_PRECIOS_ID, dbo.TB_DETALLE_LISTA_PRECIOS.VCHA_ART_ARTICULO_ID,   dbo.TB_DETALLE_LISTA_PRECIOS.FLOA_DLI_PRECIO, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, "
                                           Cadena = Cadena + " dbo.TB_LISTADEPRECIOS.DTIM_LIS_FECHA_INICIO , dbo.TB_LISTADEPRECIOS.DTIM_LIS_FECHA_FIN, dbo.TB_LISTADEPRECIOS.VCHA_MON_MONEDA, dbo.TB_MONEDAS.INTE_MON_MONEDA_LOCAL"
                                           Cadena = Cadena + " FROM dbo.TB_DETALLE_LISTA_PRECIOS INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_DETALLE_LISTA_PRECIOS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_LISTADEPRECIOS ON dbo.TB_DETALLE_LISTA_PRECIOS.VCHA_LIS_LISTA_PRECIOS_ID = dbo.TB_LISTADEPRECIOS.VCHA_LIS_LISTA_ID INNER JOIN dbo.TB_MONEDAS ON dbo.TB_LISTADEPRECIOS.VCHA_MON_MONEDA = dbo.TB_MONEDAS.VCHA_MON_MONEDA_ID  where dbo.TB_DETALLE_LISTA_PRECIOS.vcha_lis_lista_precios_id = '" + var_lista_precios + "' and dbo.TB_DETALLE_LISTA_PRECIOS.vcha_Art_Articulo_id = '" + rs!vcha_Art_Articulo_id + "'"
                                           rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                           If Not rsaux3.EOF Then
                                              var_clave_moneda = IIf(IsNull(rsaux3!vcha_mon_moneda), "", rsaux3!vcha_mon_moneda)
                                              var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
                                              If Not rsaux3.EOF Then
                                                 If var_moneda_local = 1 Then
                                                    var_tipo_Cambio = 1
                                                    var_precio = IIf(IsNull(rsaux3!floa_dli_precio), 0, rsaux3!floa_dli_precio)
                                                    var_descuento_1 = var_descuento_volumen
                                                    var_descuento_2 = var_descuento_pago
                                                    var_descuento_3 = var_descuento_financiero
                                                    var_iva = var_iva_canal
                                                 Else
                                                    rsaux.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                                    If Not rsaux.EOF Then
                                                       var_tipo_Cambio = IIf(IsNull(rsaux!mone_tca_importe), 1, rsaux!mone_tca_importe)
                                                       var_precio = IIf(IsNull(rsaux3!floa_dli_precio), 0, rsaux3!floa_dli_precio * var_tipo_Cambio)
                                                       var_descuento_1 = var_descuento_volumen
                                                       var_descuento_2 = var_descuento_pago
                                                       var_descuento_3 = var_descuento_financiero
                                                       var_iva = var_iva_canal
                                                    Else
                                                       GoTo salir:
                                                    End If
                                                    rsaux.Close
                                                 End If
                                              Else
                                                 GoTo salir_3:
                                              End If
                                           End If
                                           rsaux3.Close
                                        End If
                                     Else
                                        GoTo salir_2:
                                     End If
                                  End If
                               End If
                            End If
                            If Trim(var_clave_moneda) = "" Then
                               rsaux2.Open "select * from tb_monedas where inte_mon_moneda_local = 1", cnn, adOpenDynamic, adLockOptimistic
                               var_clave_moneda = rsaux2!vcha_mon_moneda_id
                               var_tipo_Cambio = 1
                               rsaux2.Close
                            End If
                            If var_precio = 0 Then
                               If var_empresa = "06" Then
                                  rsaux2.Open "SELECT * FROM TB_dETALLE_LISTA_PRECIOS WHERE VCHA_LIS_LISTA_PRECIOS_ID = '" + var_lista_precios + "' AND VCHA_aRT_ARTICULO_ID = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                  If Not rsaux2.EOF Then
                                     var_precio = IIf(IsNull(rsaux2!floa_dli_precio), 0, rsaux2!floa_dli_precio)
                                  Else
                                     var_precio = 0
                                  End If
                                  rsaux2.Close
                                  If var_precio = 0 Then
                                     rsaux2.Open "select * from tb_Articulos where vcha_art_articulo_id = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                     If Not rsaux2.EOF Then
                                        var_precio = IIf(IsNull(rsaux2!mone_Art_precio_base), 0, rsaux2!mone_Art_precio_base)
                                     End If
                                     rsaux2.Close
                                  End If
                               Else
                                  rsaux2.Open "select * from tb_Articulos where vcha_art_articulo_id = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                  If Not rsaux2.EOF Then
                                     var_precio = IIf(IsNull(rsaux2!mone_Art_precio_base), 0, rsaux2!mone_Art_precio_base)
                                  End If
                                  rsaux2.Close
                               End If
                            End If
                            var_contador = 0
                            var_cantidad = rs!floa_ent_Cantidad
                            'For var_contador = 1 To var_cantidad
                            var_contador = var_cantidad
                            var_cantidad_pasar = 0
                             While var_contador > 0
                                   If var_empresa <> "16" Then
                                      If var_empresa <> "06" Then
                                         If var_contador >= 1 Then
                                            var_cantidad_pasar = 1
                                            var_contador = var_contador - 1
                                         Else
                                            var_cantidad_pasar = var_contador
                                            var_contador = 0
                                         End If
                                      End If
                                   End If
                                   var_consecutivo = var_consecutivo + 1
                                   If var_requiere_factura = 1 Then
                                      If var_iva = 0 Then
                                         rsaux5.Open "select * from vw_clientes where vcha_cli_clave_id = '" + var_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                                         If Not rsaux5.EOF Then
                                            var_iva = IIf(IsNull(rsaux5!FLOA_TPE_IVA), 0, rsaux5!FLOA_TPE_IVA)
                                         End If
                                         'If var_iva = 0 Then
                                         '   If var_empresa = "02" Or var_empresa = "18" Or var_empresa = "16" Or var_empresa = "06" Or var_empresa = "31" Or var_empresa = "15" Then
                                         '      var_iva = 15
                                         '   End If
                                         'End If
                                         rsaux6.Open "SELECT top 1 FLOA_IVA_PORCENTAJE FROM TB_IVA WHERE VCHA_eMP_EMPRESA_ID = " + var_empresa + " AND GETDATE() BETWEEN DTIM_IVA_LIMITE_INFERIOR AND DTIM_IVA_LIMITE_SUPERIOR", cnn, adOpenDynamic, adLockOptimistic
                                         If Not rsaux6.EOF Then
                                            var_iva = IIf(IsNull(rsaux6(0).Value), 16, rsaux6(0).Value)
                                         Else
                                            var_iva = 16
                                         End If
                                         rsaux6.Close
                                         rsaux5.Close
                                      End If
                                   End If
                                   
                                   
                                   If var_empresa = "16" Or var_empresa = "06" Then
                                      var_cantidad_pasar = var_contador
                                      var_contador = 0
                                      If var_iva = "15" Then
                                         var_iva = "16"
                                      End If
                                      var_inserta = TB_DEVOLUCIONES_INSERTA.Anadir(CStr(var_empresa), CStr(var_unidad_organizacional), CStr(var_almacen), CStr(txt_movimiento), CDbl(txt_numero), CStr(rs!vcha_Art_Articulo_id), 0, 0, "", CInt(var_consecutivo), "", CDbl(var_costo), CDbl(var_precio), CDbl(var_descuento_1), CDbl(var_descuento_2), CDbl(var_descuento_3), CDbl(var_iva), CDbl(var_factura), CStr(var_referencia), CStr(var_clave_moneda), CDbl(var_tipo_Cambio), CStr(var_serie_FACTURA), CInt(var_año))
                                      rsaux5.Open "update tb_Devoluciones set floa_dev_cantidad = " + CStr(var_cantidad_pasar) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + Me.txt_movimiento + "' and inte_emo_numero = " + Me.txt_numero + " and vcha_art_articulo_id = '" + rs!vcha_Art_Articulo_id + "' and inte_cde_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                                   Else
                                      If var_iva = "15" Then
                                         var_iva = "16"
                                      End If
                                      var_inserta = TB_DEVOLUCIONES_INSERTA.Anadir(CStr(var_empresa), CStr(var_unidad_organizacional), CStr(var_almacen), CStr(txt_movimiento), CDbl(txt_numero), CStr(rs!vcha_Art_Articulo_id), 0, 0, "", CInt(var_consecutivo), "", CDbl(var_costo), CDbl(var_precio), CDbl(var_descuento_1), CDbl(var_descuento_2), CDbl(var_descuento_3), CDbl(var_iva), CDbl(var_factura), CStr(var_referencia), CStr(var_clave_moneda), CDbl(var_tipo_Cambio), CStr(var_serie_FACTURA), CInt(var_año))
                                      rsaux5.Open "update tb_Devoluciones set floa_dev_cantidad = " + CStr(var_cantidad_pasar) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + Me.txt_movimiento + "' and inte_emo_numero = " + Me.txt_numero + " and vcha_art_articulo_id = '" + rs!vcha_Art_Articulo_id + "' and inte_cde_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                                   End If
                             Wend
                            'Next var_contador
                            rs.MoveNext
                      Wend
                      var_posible = True
                   Else
                      rs.Close
                      MsgBox "El movimiento no a sido cerrado", vbOKOnly, "ATENCION"
                      var_posible = False
                   End If
                Else
                   var_posible = True
                End If
                If rs.State = 1 Then
                   rs.Close
                End If
                If var_posible = True Then
                   rs.Open "select sum(floa_dev_cantidad) from tb_devoluciones where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                   If rs.EOF Then
                      Me.lbl_asignacion = "0"
                   Else
                      Me.lbl_asignacion = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
                   End If
                   rs.Close
                   rsaux4.Open "select sum(floa_ent_cantidad) from tb_entradas where vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_ent_numero = " + txt_numero + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                   If Not rsaux4.EOF Then
                      Me.lbl_movimiento = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                   Else
                      Me.lbl_movimiento = "0"
                   End If
                   rsaux4.Close
                   
                   rs.Open "select * from tb_devoluciones where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic
                   var_almacen = rs!VCHA_ALM_ALMACEN_ID
                   If Not rs.EOF Then
                      var_contador_articulos = 0
                      While Not rs.EOF
                            var_estatus = IIf(IsNull(rs!CHAR_CDE_ESTATUS), "", rs!CHAR_CDE_ESTATUS)
                            var_nombre_articulo = ""
                            rsaux2.Open "select * from tb_articulos where vcha_Art_articulo_id ='" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                            If Not rsaux2.EOF Then
                               var_nombre_articulo = rsaux2!vcha_Art_nombre_español
                            End If
                            rsaux2.Close
                            Set list_item = lv_articulos.ListItems.Add(, , rs!vcha_Art_Articulo_id)
                            list_item.SubItems(1) = Trim(var_nombre_articulo)
                            list_item.SubItems(2) = IIf(IsNull(rs!INTE_CDE_CONSECUTIVO), 0, rs!INTE_CDE_CONSECUTIVO)
                            If rs!inte_fac_factura = 0 Then
                               list_item.SubItems(5) = "*"
                            End If
                            list_item.SubItems(6) = IIf(IsNull(rs!floa_cde_precio), 0, rs!floa_cde_precio)
                            list_item.SubItems(7) = IIf(IsNull(rs!inte_cde_asignado), 0, rs!inte_cde_asignado)
                            list_item.SubItems(8) = IIf(IsNull(rs!inte_dev_lote), "", rs!inte_dev_lote)
                            list_item.SubItems(9) = IIf(IsNull(rs!vcha_dev_supervisor), "", rs!vcha_dev_supervisor)
                                  
                            list_item.SubItems(10) = IIf(IsNull(rs!dtim_dev_fecha_lote), "", rs!dtim_dev_fecha_lote)
                            If list_item.SubItems(10) = "01/01/1900" Then
                               list_item.SubItems(10) = ""
                            End If
                            list_item.SubItems(11) = IIf(IsNull(rs!vcha_dev_nota_envio), "", rs!vcha_dev_nota_envio)
                            
                            list_item.SubItems(12) = IIf(IsNull(rs!VCHA_DEV_TIPO_DEFECTO), "", rs!VCHA_DEV_TIPO_DEFECTO)
                            list_item.SubItems(13) = IIf(IsNull(rs!VCHA_DEV_PROVEEDOR), "", rs!VCHA_DEV_PROVEEDOR)
                            list_item.SubItems(14) = IIf(IsNull(rs!VCHA_DEV_JUSTIFICA_DEVOLUCION), "", rs!VCHA_DEV_JUSTIFICA_DEVOLUCION)
                            list_item.SubItems(15) = IIf(IsNull(rs!VCHA_DEV_ESTATUS), "", rs!VCHA_DEV_ESTATUS)
                            
                            var_contador_articulos = var_contador_articulos + 1
                            rs.MoveNext
                      Wend
                      If var_contador_articulos > 27 Then
                         lv_articulos.ColumnHeaders(2).Width = 3900
                      Else
                         lv_articulos.ColumnHeaders(2).Width = 4100
                      End If
                      txt_numero.Enabled = False
                      txt_movimiento.Enabled = False
                      j = lv_articulos.ListItems.Count
                      For i = 1 To j
                         lv_articulos.ListItems.item(i).Selected = True
                         If lv_articulos.selectedItem.SubItems(5) = "*" Then
                            lv_articulos.selectedItem.Bold = True
                            lv_articulos.selectedItem.ListSubItems.item(1).Bold = True
                            lv_articulos.selectedItem.ListSubItems.item(2).Bold = True
                            lv_articulos.selectedItem.ListSubItems.item(3).Bold = True
                            lv_articulos.selectedItem.ForeColor = &HFF&
                            lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &HFF&
                            lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &HFF&
                            lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &HFF&
                         End If
                      Next i
                      lv_articulos.ListItems.item(1).Selected = True
                   End If
                   rs.Close
                   Call detalle_causas
                End If
             Else
                MsgBox "El Movimiento no existe o no a sido terminado aun", vbOKOnly, "ATENCION"
                rs.Close
             End If
          Else
             MsgBox "Número de Movimiento Incorrecto", vbOKOnly, "ATENCION"
          End If
       End If
      If Me.lv_articulos.ListItems.Count > 20 Then
         lv_articulos.ColumnHeaders(2).Width = 4100.03 - 200
      Else
         lv_articulos.ColumnHeaders(2).Width = 4100.03
      End If
       
    End If
    
Exit Sub
salir:
   MsgBox "No es posible asignar este movimiento ya que no se a indicado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   Exit Sub
salir_2:
   MsgBox "El cliente no tiene una lista de precios asignada", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   Exit Sub
salir_3:
   MsgBox "El cliente no tiene asignado todos los articulos en la lista de precios", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   Exit Sub
End Sub

Private Sub txt_proveedor_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_supervisor_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub
