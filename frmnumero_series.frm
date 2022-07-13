VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmnumero_series 
   BorderStyle     =   0  'None
   Caption         =   "Números de Serie"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_Cantidad 
      Height          =   300
      Left            =   3030
      TabIndex        =   6
      Top             =   3855
      Width           =   900
   End
   Begin VB.TextBox txt_numero 
      Height          =   285
      Left            =   2055
      TabIndex        =   5
      Top             =   3855
      Width           =   885
   End
   Begin VB.TextBox txt_almacen_destino 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   3855
      Width           =   825
   End
   Begin VB.TextBox txt_movimiento 
      Height          =   300
      Left            =   90
      TabIndex        =   3
      Top             =   3855
      Width           =   1080
   End
   Begin VB.Frame frm_articulos_serie 
      Height          =   3345
      Left            =   15
      TabIndex        =   0
      Top             =   270
      Width           =   6270
      Begin MSComctlLib.ListView lv_lista 
         Height          =   3150
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5556
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Artículo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Número de Serie"
            Object.Width           =   10583
         EndProperty
      End
   End
   Begin VB.Label lbl_estampado 
      BackColor       =   &H8000000D&
      Caption         =   " Números de Serie"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6315
   End
End
Attribute VB_Name = "frmnumero_series"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         rsaux4.Open "delete from TB_EXISTENCIAS_SERIES where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + Me.txt_almacen_destino + "' and vcha_mov_movimiento_id = '" + Me.txt_movimiento + "' and inte_emo_numero = " + Me.txt_numero + " and vcha_Art_articulo_id = '" + Me.lv_lista.selectedItem + "' and vcha_art_numero_serie = '" + Me.lv_lista.selectedItem.SubItems(1) + "'", cnn, adOpenDynamic, adLockOptimistic
         var_si_elimino = 1
         frmentradas.frm_eliminar.Visible = False
         Unload Me
      End If
   End If
   If KeyAscii = 27 Then
      frmentradas.frm_eliminar.Visible = False
      Unload Me
   End If
End Sub
