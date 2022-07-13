VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmpedido_tienda_cantia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar pedido"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "  Existencias"
      Height          =   735
      Left            =   135
      TabIndex        =   37
      Top             =   2595
      Width           =   9120
      Begin VB.TextBox txt_existencia_tienda 
         Height          =   390
         Left            =   1455
         TabIndex        =   40
         Top             =   210
         Width           =   1110
      End
      Begin VB.TextBox txt_existencia_exhibicion 
         Height          =   390
         Left            =   4575
         TabIndex        =   39
         Top             =   210
         Width           =   1110
      End
      Begin VB.TextBox txt_existencia_almacen 
         Height          =   390
         Left            =   7920
         TabIndex        =   38
         Top             =   210
         Width           =   1110
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tienda:"
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
         Left            =   510
         TabIndex        =   43
         Top             =   255
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Exhibición:"
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
         Left            =   3225
         TabIndex        =   42
         Top             =   255
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Almacén:"
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
         Left            =   6720
         TabIndex        =   41
         Top             =   255
         Width           =   1125
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   2475
      TabIndex        =   34
      Top             =   255
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   35
         Top             =   480
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3228
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
         TabIndex        =   36
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.Frame frm_disponibles 
      Height          =   3585
      Left            =   1155
      TabIndex        =   27
      Top             =   3540
      Width           =   7110
      Begin VB.TextBox txt_nombre_articulo 
         Height          =   345
         Left            =   90
         TabIndex        =   28
         Top             =   420
         Width           =   6915
      End
      Begin MSComctlLib.ListView lv_disponibles 
         Height          =   2700
         Left            =   75
         TabIndex        =   29
         Top             =   795
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   4763
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
            Text            =   "Código"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre del Artículo"
            Object.Width           =   9701
         EndProperty
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000D&
         Caption         =   " Artículos Disponibles"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   30
         Top             =   120
         Width           =   7035
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8865
      Picture         =   "frmpedido_tienda_cantia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   1230
      TabIndex        =   23
      Top             =   -15
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         MaxLength       =   10
         TabIndex        =   24
         Top             =   510
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   2
         Left            =   30
         TabIndex        =   25
         Top             =   120
         Width           =   3075
      End
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Picture         =   "frmpedido_tienda_cantia.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Buscar Movimiento Alt + B"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   180
      Picture         =   "frmpedido_tienda_cantia.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   510
      Picture         =   "frmpedido_tienda_cantia.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   "Lista de artículos pedidos "
      Height          =   2835
      Left            =   135
      TabIndex        =   20
      Top             =   3330
      Width           =   9120
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   5955
         TabIndex        =   31
         Top             =   1020
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   75
            TabIndex        =   32
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   "Cantidad a eliminar"
            ForeColor       =   &H8000000E&
            Height          =   225
            Index           =   3
            Left            =   0
            TabIndex        =   33
            Top             =   15
            Width           =   2895
         End
      End
      Begin VB.TextBox txt_total 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7110
         TabIndex        =   13
         Top             =   2355
         Width           =   1920
      End
      Begin MSComctlLib.ListView lv_pedido 
         Height          =   2115
         Left            =   135
         TabIndex        =   12
         Top             =   240
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   3731
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Existen"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cantidad"
            Object.Width           =   2469
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Left            =   6405
         TabIndex        =   26
         Top             =   2385
         Width           =   690
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Artículo "
      Height          =   735
      Left            =   135
      TabIndex        =   19
      Top             =   1830
      Width           =   9120
      Begin VB.TextBox txt_cantidad 
         Height          =   390
         Left            =   7920
         TabIndex        =   11
         Top             =   210
         Width           =   1110
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   390
         Left            =   1665
         TabIndex        =   10
         Top             =   225
         Width           =   4830
      End
      Begin VB.TextBox txt_codigo 
         Height          =   390
         Left            =   240
         TabIndex        =   9
         Top             =   225
         Width           =   1395
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
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
         Left            =   6720
         TabIndex        =   21
         Top             =   255
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos del pedido "
      Height          =   1335
      Left            =   135
      TabIndex        =   15
      Top             =   480
      Width           =   9120
      Begin VB.TextBox txt_folio 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7395
         TabIndex        =   6
         Top             =   255
         Width           =   1620
      End
      Begin VB.TextBox txt_felefono 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7395
         TabIndex        =   8
         Top             =   810
         Width           =   1620
      End
      Begin VB.TextBox txt_cliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1410
         TabIndex        =   7
         Top             =   810
         Width           =   4740
      End
      Begin VB.TextBox txt_nombre_usuario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2145
         TabIndex        =   5
         Top             =   255
         Width           =   4005
      End
      Begin VB.TextBox txt_usuario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1410
         TabIndex        =   4
         Top             =   255
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
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
         Left            =   6300
         TabIndex        =   22
         Top             =   330
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Télefono:"
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
         Left            =   6255
         TabIndex        =   18
         Top             =   885
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         Index           =   0
         Left            =   150
         TabIndex        =   17
         Top             =   885
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
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
         Left            =   135
         TabIndex        =   16
         Top             =   330
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   75
      TabIndex        =   14
      Top             =   330
      Width           =   9180
   End
End
Attribute VB_Name = "frmpedido_tienda_cantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_renglon As Double
Dim var_primera_vez As Boolean
Sub ilumina_grid()
   var_n = lv_pedido.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_pedido.ListItems.Item(var_i).Bold = True
          lv_pedido.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_pedido.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_pedido.ListItems.Item(var_i).ForeColor = &H8000&
          lv_pedido.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_pedido.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
       Else
          lv_pedido.ListItems.Item(var_i).Bold = False
          lv_pedido.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_pedido.ListItems.Item(var_i).ListSubItems(2).Bold = False
          lv_pedido.ListItems.Item(var_i).ForeColor = &H80000012
          lv_pedido.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
          lv_pedido.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_pedido.ListItems.Item(var_renglon).Selected = True
      lv_pedido.selectedItem.EnsureVisible
   End If
   lv_pedido.Refresh
End Sub

Private Sub cmd_buscar_Click()
   Me.frm_busqueda.Visible = True
   Me.frm_disponibles.Visible = False
   Me.txt_busqueda_folio = ""
   Me.txt_busqueda_folio.SetFocus
End Sub

Private Sub cmd_imprimir_Click()
   Me.frm_busqueda.Visible = False
   Me.frm_disponibles.Visible = False
   If Me.txt_folio <> "" Then
      rs.Open "SELECT dbo.TB_PEDIDO_TIENDA_CANTIA.DTIM_PED_FECHA, dbo.TB_PEDIDO_TIENDA_CANTIA.INTE_PED_NUMERO, dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_PED_CLIENTE, dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_PED_TELEFONO, dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_USU_USUARIO_ID, dbo.TB_USUARIOS_PEDIDOS_CANTIA.VCHA_USU_NOMBRE, dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_ART_ARTICULO_ID, dbo.TB_PEDIDO_TIENDA_CANTIA.FLOA_PED_CANTIDAD , dbo.TB_Articulos.VCHA_ART_NOMBRE_ESPAÑOL FROM dbo.TB_ARTICULOS INNER JOIN dbo.TB_PEDIDO_TIENDA_CANTIA ON dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID = dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_USUARIOS_PEDIDOS_CANTIA ON dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_USU_USUARIO_ID = dbo.TB_USUARIOS_PEDIDOS_CANTIA.VCHA_USU_USUARIO_ID Where (dbo.TB_PEDIDO_TIENDA_CANTIA.INTE_PED_NUMERO = " + Me.txt_folio + ") AND FLOA_PED_CANTIDAD > 0", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         'Open (App.Path & "\pedido" + Trim(Me.txt_folio) + ".txt") For Output As #1
         Open ("c:\sistemas\puntodeventa\imprimir.txt") For Output As #1
         Print #1, ""
         Print #1, "             CANTIA S.A. DE C.V."
         Print #1, "PEDIDO: " + Me.txt_folio
         Print #1, "FECHA: " + CStr(rs!dtim_ped_fecha)
         Print #1, "CLIENTE: " + Me.txt_cliente
         Print #1, "VENDEDOR: " + Me.txt_usuario + " " + Me.txt_nombre_usuario
         Print #1, ""
         Print #1, "-----------------------------------------"
         While Not rs.EOF
               Print #1, "SKU: " + rs!VCHA_aRT_ARTICULO_ID
               Print #1, "DESCRIPCION: " + rs!VCHA_ART_NOMBRE_ESPAÑOL
               Print #1, "CANTIDAD: " + CStr(rs!FLOA_PED_CANTIDAD)
               Print #1, ""
               rs.MoveNext
         Wend
         Print #1, "-----------------------------------------"
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Print #1, Chr(27) + Chr(105)
         Close #1
         
         VAR_MAQUINA = fun_NombrePc
         If VAR_MAQUINA = "ESALAS" Then
            Open (App.Path & "\pedido" + Trim(Me.txt_folio) + ".bat") For Output As #2
            Print #2, "copy c:\sistemas\puntodeventa\imprimir.txt \\" + Trim(fun_NombrePc) + "\epson"
            var_Archivo = App.Path & "\pedido" + Trim(Me.txt_folio) + ".bat"
            Close #2
            x = Shell(var_Archivo, vbHide)
         Else
            If VAR_MAQUINA = "FSERNAPORT" Then
               Open (App.Path & "\pedido" + Trim(Me.txt_folio) + ".bat") For Output As #2
               Print #2, "copy c:\sistemas\puntodeventa\imprimir.txt \\" + Trim(fun_NombrePc) + "\epson"
               var_Archivo = App.Path & "\pedido" + Trim(Me.txt_folio) + ".bat"
               Close #2
               x = Shell(var_Archivo, vbHide)
            Else
               x = Shell("c:\sistemas\puntodeventa\imprimir.exe", vbHide)
            End If
         End If
      Else
      End If
      rs.Close
   Else
      MsgBox "No se a seleccionado un pedido", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub cmd_nuevo_Click()
   Me.frm_busqueda.Visible = False
   Me.frm_disponibles.Visible = False
   var_primera_vez = True
   Me.txt_usuario = ""
   Me.txt_nombre_usuario = ""
   Me.txt_folio = ""
   Me.txt_cliente = ""
   Me.txt_felefono = ""
   Me.txt_codigo = ""
   Me.txt_descripcion = ""
   Me.txt_cantidad = ""
   Me.txt_total = "0"
   Me.lv_pedido.ListItems.Clear
   Me.txt_usuario.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_salir_GotFocus()
   Me.frm_busqueda.Visible = False
   Me.frm_disponibles.Visible = False
End Sub

Private Sub Form_Load()
   Top = 600
   Left = 1200
   Me.frm_busqueda.Visible = False
   Me.frm_disponibles.Visible = False
   Me.frm_eliminar.Visible = False
   Me.frm_lista.Visible = False
   var_primera_vez = True
End Sub

Private Sub lv_disponibles_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 38 And Shift = 1 Then
      Me.txt_nombre_articulo.SetFocus
   End If
End Sub

Private Sub lv_disponibles_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_disponibles.ListItems.Count > 0 Then
         Me.txt_codigo = Me.lv_disponibles.selectedItem
         Me.txt_descripcion = Me.lv_disponibles.selectedItem.SubItems(1)
         Me.txt_codigo.SetFocus
         Me.frm_disponibles.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_disponibles.Visible = False
   End If
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_usuario = Me.lv_lista.selectedItem
      Me.txt_nombre_usuario = Me.lv_lista.selectedItem.SubItems(1)
      Me.txt_usuario.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub lv_pedido_GotFocus()
   Me.frm_busqueda.Visible = False
   Me.frm_disponibles.Visible = False
End Sub

Private Sub lv_pedido_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      Me.txt_cantidad_eliminar = ""
      Me.frm_eliminar.Visible = True
      Me.txt_cantidad_eliminar.SetFocus
   End If
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_busqueda_folio) Then
         var_cadena = "SELECT dbo.TB_PEDIDO_TIENDA_CANTIA.INTE_PED_NUMERO, dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_PED_CLIENTE, dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_PED_TELEFONO, dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_USU_USUARIO_ID, dbo.TB_USUARIOS_PEDIDOS_CANTIA.VCHA_USU_NOMBRE, dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_ART_ARTICULO_ID, dbo.TB_PEDIDO_TIENDA_CANTIA.FLOA_PED_CANTIDAD , dbo.TB_Articulos.VCHA_ART_NOMBRE_ESPAÑOL FROM dbo.TB_ARTICULOS INNER JOIN dbo.TB_PEDIDO_TIENDA_CANTIA ON dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID = dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_USUARIOS_PEDIDOS_CANTIA ON dbo.TB_PEDIDO_TIENDA_CANTIA.VCHA_USU_USUARIO_ID = dbo.TB_USUARIOS_PEDIDOS_CANTIA.VCHA_USU_USUARIO_ID Where (dbo.TB_PEDIDO_TIENDA_CANTIA.INTE_PED_NUMERO = " + Me.txt_busqueda_folio + ")"
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_usuario = IIf(IsNull(rs!vcha_usu_usuario_ID), "", rs!vcha_usu_usuario_ID)
            Me.txt_nombre_usuario = IIf(IsNull(rs!VCHA_USU_NOMBRE), "", rs!VCHA_USU_NOMBRE)
            Me.txt_folio = IIf(IsNull(rs!inte_ped_numero), "", rs!inte_ped_numero)
            Me.txt_cliente = IIf(IsNull(rs!VCHA_PED_CLIENTE), "", rs!VCHA_PED_CLIENTE)
            Me.txt_felefono = IIf(IsNull(rs!VCHA_PED_TELEFONO), "", rs!VCHA_PED_TELEFONO)
            Me.txt_codigo = ""
            Me.txt_descripcion = ""
            Me.txt_cantidad = ""
            Me.lv_pedido.ListItems.Clear
            Me.txt_total = 0
            While Not rs.EOF
                  Set list_item = lv_pedido.ListItems.Add(, , rs!VCHA_aRT_ARTICULO_ID)
                  list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ART_NOMBRE_ESPAÑOL), "", rs!VCHA_ART_NOMBRE_ESPAÑOL)
                  list_item.SubItems(3) = Format(IIf(IsNull(rs!FLOA_PED_CANTIDAD), 0, rs!FLOA_PED_CANTIDAD), "###,###,##0.00")
                  var_codigo = rs!VCHA_aRT_ARTICULO_ID
                  rsaux.Open "select * from VIA_EXISTENCIA_ALMACEN where art_codigo = '" + var_codigo + "' or art_gtin = '" + var_codigo + "'", cnn_compucaja, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                    var_existen = IIf(IsNull(rsaux!TIENDA), 0, rsaux!TIENDA) + IIf(IsNull(rsaux!EXHIBICION), 0, rsaux!EXHIBICION)
                  End If
                  rsaux.Close
                  rsaux.Open "SELECT dbo.TB_ALMACENES.vcha_alm_almacen_id, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_EXISTENCIAS.FLOA_EXI_CANTIDAD, dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID fROM dbo.TB_EXISTENCIAS INNER JOIN dbo.TB_ALMACENES ON dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID where vcha_Art_Articulo_id = '" + var_codigo + "' ", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux.EOF
                        If rsaux!vcha_alm_almacen_id = "PTVH" Then
                           var_existen = var_existen + IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad)
                        End If
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  list_item.SubItems(2) = Format(var_existen, "###,###,##0.00")
                  
                  
                  Me.txt_total = CDbl(Me.txt_total) + IIf(IsNull(rs!FLOA_PED_CANTIDAD), 0, rs!FLOA_PED_CANTIDAD)
                  rs.MoveNext
            Wend
            var_primera_vez = False
            Me.txt_codigo.SetFocus
         Else
            MsgBox "El pedido no existe", vbOKOnly, "ATENCION"
            var_primera_vez = True
            Me.txt_usuario = ""
            Me.txt_nombre_usuario = ""
            Me.txt_folio = ""
            Me.txt_cliente = ""
            Me.txt_felefono = ""
            Me.txt_codigo = ""
            Me.txt_descripcion = ""
            Me.txt_cantidad = ""
            Me.lv_pedido.ListItems.Clear
            Me.txt_total = ""
            Me.txt_usuario.SetFocus
         End If
         rs.Close
      Else
         MsgBox "Número de pedido incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_busqueda_folio_LostFocus()
   Me.frm_busqueda.Visible = False
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad_eliminar) Then
         If CDbl(Me.lv_pedido.selectedItem.SubItems(3)) >= CDbl(Me.txt_cantidad_eliminar) Then
            rs.Open "UPDATE TB_PEDIDO_TIENDA_CANTIA SET FLOA_PED_CANTIDAD = ISNULL(FLOA_PED_CANTIDAD,0) - " + Me.txt_cantidad_eliminar + " WHERE INTE_PED_NUMERO = " + Me.txt_folio + " AND VCHA_ART_ARTICULO_ID = '" + Me.lv_pedido.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            Me.lv_pedido.selectedItem.SubItems(3) = CDbl(Me.lv_pedido.selectedItem.SubItems(3)) - CDbl(Me.txt_cantidad_eliminar)
            Me.txt_total = CDbl(Me.txt_total) - CDbl(Me.txt_cantidad_eliminar)
            Me.frm_eliminar.Visible = False
         Else
            MsgBox "La cantidad a eliminar no debe de ser mayor a la cantidad pedida", vbOKOnly, "ATENCION"
            Me.frm_eliminar.Visible = False
         End If
      Else
         MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
         Me.frm_eliminar.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_eliminar.Visible = False
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   Me.frm_eliminar.Visible = False
End Sub

Private Sub txt_cantidad_GotFocus()
   Me.frm_busqueda.Visible = False
   Me.frm_disponibles.Visible = False
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad) Then
         If Me.txt_usuario <> "" Then
            If Me.txt_cliente <> "" Then
               If Me.txt_codigo <> "" Then
                  If Me.txt_total = "" Then
                     Me.txt_total = "0"
                  End If
                  If var_primera_vez = True Then
                     cnn.BeginTrans
                     rs.Open "select MAX(INTE_PED_NUMERO)  from tb_PEDIDO_TIENDA_CANTIA", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        Me.txt_folio = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                     Else
                        Me.txt_folio = 1
                     End If
                     rs.Close
                     rs.Open "INSERT INTO TB_PEDIDO_TIENDA_CANTIA (INTE_PED_NUMERO, VCHA_PED_CLIENTE, VCHA_PED_TELEFONO, VCHA_ART_ARTICULO_ID, FLOA_PED_CANTIDAD, VCHA_USU_USUARIO_ID) VALUES (" + CStr(Me.txt_folio) + ",'" + Me.txt_cliente + "','" + Me.txt_felefono + "','" + Me.txt_codigo + "'," + Me.txt_cantidad + ", '" + Me.txt_usuario + "')", cnn, adOpenDynamic, adLockOptimistic
                     cnn.CommitTrans
                     var_codigo = Trim(Me.txt_codigo)
                     Set list_item = lv_pedido.ListItems.Add(, , Trim(txt_codigo))
                     list_item.SubItems(1) = Me.txt_descripcion
                     list_item.SubItems(3) = Me.txt_cantidad
                     var_renglon = lv_pedido.ListItems.Count
                     Me.txt_total = CDbl(Me.txt_total) + CDbl(Me.txt_cantidad)
                     Call ilumina_grid
                     var_primera_vez = False
                     Me.txt_codigo.SetFocus
                  Else
                     rs.Open "SELECT * FROM TB_PEDIDO_TIENDA_CANTIA WHERE INTE_PED_NUMERO = " + Me.txt_folio + " AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        rsaux.Open "UPDATE TB_PEDIDO_TIENDA_CANTIA SET FLOA_PED_CANTIDAD = ISNULL(FLOA_PED_CANTIDAD,0) + " + Me.txt_cantidad + " WHERE INTE_PED_NUMERO = " + Me.txt_folio + " AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_codigo = Trim(Me.txt_codigo)
                        valor = Trim(txt_codigo)
                        var_j = 1
                        Me.lv_pedido.ListItems(var_j).Selected = True
                        While valor <> Me.lv_pedido.selectedItem
                              var_j = var_j + 1
                              Me.lv_pedido.ListItems(var_j).Selected = True
                        Wend
                        lv_pedido.selectedItem.SubItems(3) = lv_pedido.selectedItem.SubItems(3) + CDbl(Me.txt_cantidad)
                        var_renglon = lv_pedido.selectedItem.Index
                        Me.txt_total = CDbl(Me.txt_total) + CDbl(Me.txt_cantidad)
                        Call ilumina_grid
                     Else
                        var_codigo = Trim(Me.txt_codigo)
                        rsaux.Open "INSERT INTO TB_PEDIDO_TIENDA_CANTIA (INTE_PED_NUMERO, VCHA_PED_CLIENTE, VCHA_PED_TELEFONO, VCHA_ART_ARTICULO_ID, FLOA_PED_CANTIDAD, VCHA_USU_USUARIO_ID) VALUES (" + CStr(Me.txt_folio) + ",'" + Me.txt_cliente + "','" + Me.txt_felefono + "','" + Me.txt_codigo + "'," + Me.txt_cantidad + ", '" + Me.txt_usuario + "')", cnn, adOpenDynamic, adLockOptimistic
                        Set list_item = lv_pedido.ListItems.Add(, , Trim(txt_codigo))
                        list_item.SubItems(1) = Me.txt_descripcion
                        list_item.SubItems(3) = Me.txt_cantidad
                        var_renglon = lv_pedido.ListItems.Count
                        Me.txt_total = CDbl(Me.txt_total) + CDbl(Me.txt_cantidad)
                        Call ilumina_grid
                     End If
                     rs.Close
                     Me.txt_codigo = ""
                     Me.txt_descripcion = ""
                     Me.txt_cantidad = ""
                     Me.txt_codigo.SetFocus
                  End If
                  var_existen = 0
                  rsaux.Open "select * from VIA_EXISTENCIA_ALMACEN where art_codigo = '" + var_codigo + "' or art_gtin = '" + var_codigo + "'", cnn_compucaja, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                    'var_existen = IIf(IsNull(rsaux!TIENDA), 0, rsaux!TIENDA) + IIf(IsNull(rsaux!EXHIBICION), 0, rsaux!EXHIBICION)
                    var_existen = IIf(IsNull(rsaux!TIENDA), 0, rsaux!TIENDA)
                  End If
                  rsaux.Close
                  rsaux.Open "SELECT dbo.TB_ALMACENES.vcha_alm_almacen_id, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_EXISTENCIAS.FLOA_EXI_CANTIDAD, dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID fROM dbo.TB_EXISTENCIAS INNER JOIN dbo.TB_ALMACENES ON dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID where vcha_Art_Articulo_id = '" + var_codigo + "' ", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux.EOF
                        If rsaux!vcha_alm_almacen_id <> "ACCAN" Then
                           If rsaux!vcha_alm_almacen_id <> "INVH" Then
                              If rsaux!vcha_alm_almacen_id <> "CANCA" Then
                                 var_existen = var_existen + IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad)
                              End If
                           End If
                        End If
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  lv_pedido.selectedItem.SubItems(2) = Format(var_existen, "###,###,##0.00")

                  
                  
               Else
                  MsgBox "Código incorrecto", vbOKOnly, "ATENCION"
                  Me.txt_codigo.SetFocus
               End If
            Else
               MsgBox "Cliente incorrecto", vbOKOnly, "ATENCION"
               Me.txt_cliente.SetFocus
            End If
         Else
            MsgBox "Usuario incorrecto", vbOKOnly, "ATENCION"
            Me.txt_usuario.SetFocus
         End If
      Else
         MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      If KeyAscii = 27 Then
         Unload Me
      End If
   End If
End Sub

Private Sub txt_cantidad_LostFocus()
   Me.txt_codigo = ""
   Me.txt_descripcion = ""
   Me.txt_cantidad = ""
End Sub

Private Sub txt_cliente_GotFocus()
   Me.frm_busqueda.Visible = False
   Me.frm_disponibles.Visible = False
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_felefono.SetFocus
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_codigo_Change()
   Me.txt_existencia_almacen = ""
   Me.txt_existencia_exhibicion = ""
   Me.txt_existencia_tienda = ""
End Sub

Private Sub txt_codigo_GotFocus()
   Me.frm_busqueda.Visible = False
   Me.frm_disponibles.Visible = False
End Sub

Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.txt_nombre_articulo = ""
      Me.lv_disponibles.ListItems.Clear
      Me.frm_disponibles.Visible = True
      Me.txt_nombre_articulo.SetFocus
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_cantidad.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      End If
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   If Trim(Me.txt_codigo) <> "" Then
      Me.txt_descripcion = ""
      var_codigo = ""
      var_nombre = ""
      rs.Open "select * from tb_Articulos where vcha_art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_codigo = IIf(IsNull(rs!VCHA_aRT_ARTICULO_ID), "", rs!VCHA_aRT_ARTICULO_ID)
         var_nombre = IIf(IsNull(rs!VCHA_ART_NOMBRE_ESPAÑOL), "", rs!VCHA_ART_NOMBRE_ESPAÑOL)
      Else
         rsaux.Open "SELECT * FROM TB_EQUIVALENCIAS WHERE VCHA_EQU_CODIGO_EQUIVALENTE = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux1.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + IIf(IsNull(rsaux!VCHA_aRT_ARTICULO_ID), "", rsaux!VCHA_aRT_ARTICULO_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               var_codigo = IIf(IsNull(rsaux1!VCHA_aRT_ARTICULO_ID), "", rsaux1!VCHA_aRT_ARTICULO_ID)
               var_nombre = IIf(IsNull(rsaux1!VCHA_ART_NOMBRE_ESPAÑOL), "", rsaux1!VCHA_ART_NOMBRE_ESPAÑOL)
            Else
               var_codigo = ""
            End If
            rsaux1.Close
         Else
            var_codigo = ""
         End If
         rsaux.Close
      End If
      rs.Close
      If var_codigo <> "" Then
         Me.txt_descripcion = var_nombre
         Me.txt_codigo = var_codigo
      
      
         rsaux.Open "select * from VIA_EXISTENCIA_ALMACEN where art_codigo = '" + var_codigo + "' or art_gtin = '" + var_codigo + "'", cnn_compucaja, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            'Me.txt_existencia_tienda = IIf(IsNull(rsaux!TIENDA), 0, rsaux!TIENDA)
            'Me.txt_existencia_exhibicion = IIf(IsNull(rsaux!EXHIBICION), 0, rsaux!EXHIBICION)
         End If
         rsaux.Close
         rsaux.Open "SELECT dbo.TB_ALMACENES.vcha_alm_almacen_id, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_EXISTENCIAS.FLOA_EXI_CANTIDAD, dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID fROM dbo.TB_EXISTENCIAS INNER JOIN dbo.TB_ALMACENES ON dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID where vcha_Art_Articulo_id = '" + var_codigo + "' ", cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               If rsaux!vcha_alm_almacen_id = "PTVH" Then
                  Me.txt_existencia_almacen = IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad)
               End If
               If rsaux!vcha_alm_almacen_id = "CC_1" Then
                  Me.txt_existencia_tienda = IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad)
               End If
               If rsaux!vcha_alm_almacen_id = "CC_5" Then
                  Me.txt_existencia_exhibicion = IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad)
               End If
               rsaux.MoveNext
        Wend
        rsaux.Close
      
      
      
      
      
      Else
         Me.txt_descripcion = ""
         Me.txt_codigo = ""
         MsgBox "Código incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      Me.txt_codigo = ""
      Me.txt_descripcion = ""
   End If
End Sub

Private Sub txt_descripcion_GotFocus()
   Me.frm_busqueda.Visible = False
   Me.frm_disponibles.Visible = False
End Sub

Private Sub txt_descripcion_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.txt_nombre_articulo = ""
      Me.lv_disponibles.ListItems.Clear
      Me.frm_disponibles.Visible = True
      Me.txt_nombre_articulo.SetFocus
   End If
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_cantidad.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_felefono_GotFocus()
   Me.frm_busqueda.Visible = False
   Me.frm_disponibles.Visible = False
End Sub

Private Sub txt_felefono_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_codigo.SetFocus
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_folio_GotFocus()
   Me.frm_busqueda.Visible = False
   Me.frm_disponibles.Visible = False
End Sub

Private Sub txt_nombre_articulo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyAscii = 13 Then
      If Me.lv_disponibles.ListItems.Count > 0 Then
         Me.txt_codigo = Me.lv_disponibles.selectedItem
         Me.txt_descripcion = Me.lv_disponibles.selectedItem.SubItems(1)
         Me.txt_codigo.SetFocus
         Me.frm_disponibles.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_disponibles.Visible = False
   End If
End Sub

Private Sub txt_nombre_articulo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 27 Then
      Me.txt_codigo.SetFocus
      Me.frm_disponibles.Visible = False
   End If
   If KeyAscii = 39 Or KeyAscii = 61 Or KeyAscii = 44 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(Me.txt_nombre_articulo) <> "" Then
         var_cadena = ""
         var_like_1 = ""
         var_like_2 = ""
         var_like_3 = ""
         var_like_4 = ""
         var_like_5 = ""
         var_like_6 = ""
         var_like_7 = ""
         var_j = 1
         For var_i = 1 To Len(Me.txt_nombre_articulo)
             If Mid(Me.txt_nombre_articulo, var_i, 1) <> " " Then
                If var_j = 1 Then
                   var_like_1 = var_like_1 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 2 Then
                   var_like_2 = var_like_2 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 3 Then
                   var_like_3 = var_like_3 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 4 Then
                   var_like_4 = var_like_4 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 5 Then
                   var_like_5 = var_like_5 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 6 Then
                   var_like_6 = var_like_6 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j >= 7 Then
                   var_like_7 = var_like_7 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
             Else
                var_j = var_j + 1
             End If
         Next var_i
      End If
      If Trim(var_like_1) <> "" Then
         var_cadena = " vcha_art_nombre_Español like '%" + var_like_1 + "%'"
      End If
      If Trim(var_like_2) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_2 + "%'"
      End If
      If Trim(var_like_3) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_3 + "%'"
      End If
      If Trim(var_like_4) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_4 + "%'"
      End If
      If Trim(var_like_5) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_5 + "%'"
      End If
      If Trim(var_like_6) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_6 + "%'"
      End If
      If Trim(var_like_7) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_7 + "%'"
      End If
      Me.lv_disponibles.ListItems.Clear
      If Trim(var_cadena) <> "" Then
         var_cadena = "SELECT * FROM tb_articulos WHERE " + var_cadena
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
            Set list_item = lv_disponibles.ListItems.Add(, , rs!VCHA_aRT_ARTICULO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ART_NOMBRE_ESPAÑOL), "", rs!VCHA_ART_NOMBRE_ESPAÑOL)
            rs.MoveNext
         Wend
         rs.Close
         If Me.lv_disponibles.ListItems.Count > 0 Then
            Me.lv_disponibles.SetFocus
         End If
         If lv_disponibles.ListItems.Count > 11 Then
            lv_disponibles.ColumnHeaders(2).Width = 5300
         Else
            lv_disponibles.ColumnHeaders(2).Width = 5500
         End If
      End If
   End If
End Sub

Private Sub txt_nombre_usuario_GotFocus()
   Me.frm_busqueda.Visible = False
   Me.frm_disponibles.Visible = False
End Sub

Private Sub txt_nombre_usuario_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_usuarios_pedidos_cantia order by vcha_usu_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_usu_usuario_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_USU_NOMBRE), "", rs!VCHA_USU_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "VENDEDORES"
      VAR_TIPO_LISTA = 21
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_usuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_cliente.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_total_GotFocus()
   Me.frm_busqueda.Visible = False
   Me.frm_disponibles.Visible = False
End Sub

Private Sub txt_usuario_Change()
   Me.txt_cliente = ""
   Me.txt_nombre_usuario = ""
   Me.lv_pedido.ListItems.Clear
   Me.txt_total = ""
   Me.txt_busqueda_folio = ""
   Me.txt_folio = ""
   Me.txt_codigo = ""
   Me.txt_descripcion = ""
   Me.txt_cantidad = ""
End Sub

Private Sub txt_usuario_GotFocus()
   Me.frm_busqueda.Visible = False
   Me.frm_disponibles.Visible = False
End Sub

Private Sub txt_usuario_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_usuarios_pedidos_cantia order by vcha_usu_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_usu_usuario_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_USU_NOMBRE), "", rs!VCHA_USU_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "VENDEDORES"
      VAR_TIPO_LISTA = 21
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_usuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre_usuario.SetFocus
   End If
End Sub

Private Sub txt_usuario_LostFocus()
    If Me.txt_usuario <> "" Then
       rs.Open "SELECT * FROM TB_USUARIOS_PEDIDOS_CANTIA WHERE VCHA_USU_USUARIO_ID = '" + Me.txt_usuario + "'", cnn, adOpenDynamic, adLockOptimistic
       If Not rs.EOF Then
          Me.txt_nombre_usuario = IIf(IsNull(rs!VCHA_USU_NOMBRE), "", rs!VCHA_USU_NOMBRE)
       Else
          MsgBox "Clave de usuario incorrecta", vbOKOnly, "ATENCION"
       End If
       rs.Close
    End If
End Sub



