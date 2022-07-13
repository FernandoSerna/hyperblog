VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmasignacion_catalogo_lista_precios 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   5790
      TabIndex        =   28
      Top             =   3630
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   29
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
         TabIndex        =   30
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11145
      Picture         =   "frmassignacion_catalogo_lista_precios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar 
      Height          =   315
      Left            =   75
      Picture         =   "frmassignacion_catalogo_lista_precios.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Aceptar"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmassignacion_catalogo_lista_precios.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Cancelar"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   45
      Left            =   0
      TabIndex        =   18
      Top             =   390
      Width           =   11565
   End
   Begin VB.Frame Frame3 
      Caption         =   " Vigencia "
      Height          =   675
      Left            =   7365
      TabIndex        =   17
      Top             =   6495
      Width           =   4155
      Begin VB.TextBox txt_fecha_fin 
         Height          =   315
         Left            =   2865
         TabIndex        =   25
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox txt_fecha_inicio 
         Height          =   315
         Left            =   705
         TabIndex        =   24
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2565
         TabIndex        =   27
         Top             =   270
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Inicio: "
         Height          =   195
         Left            =   225
         TabIndex        =   26
         Top             =   300
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Catálogo "
      Height          =   645
      Left            =   7350
      TabIndex        =   16
      Top             =   5835
      Width           =   4155
      Begin VB.TextBox txt_nombre_catalogo 
         Height          =   315
         Left            =   990
         TabIndex        =   23
         Top             =   210
         Width           =   3090
      End
      Begin VB.TextBox txt_catalogo 
         Height          =   315
         Left            =   60
         TabIndex        =   22
         Top             =   210
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5355
      Left            =   7365
      TabIndex        =   8
      Top             =   420
      Width           =   4140
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmassignacion_catalogo_lista_precios.frx":08CE
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command9 
         Height          =   315
         Left            =   30
         Picture         =   "frmassignacion_catalogo_lista_precios.frx":0AE4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         Picture         =   "frmassignacion_catalogo_lista_precios.frx":0BE6
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   690
         Picture         =   "frmassignacion_catalogo_lista_precios.frx":0CB8
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Marcar (Enter)"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Picture         =   "frmassignacion_catalogo_lista_precios.frx":0F02
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   360
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_canalesventas 
         Height          =   4605
         Left            =   45
         TabIndex        =   14
         Top             =   690
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   8123
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
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5327
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   " Canales de Venta"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   15
         Top             =   120
         Width           =   4065
      End
   End
   Begin VB.Frame frm_canales 
      Height          =   6705
      Left            =   195
      TabIndex        =   0
      Top             =   420
      Width           =   7095
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Picture         =   "frmassignacion_catalogo_lista_precios.frx":1118
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   690
         Picture         =   "frmassignacion_catalogo_lista_precios.frx":132E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Marcar (Enter)"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         Picture         =   "frmassignacion_catalogo_lista_precios.frx":1578
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   30
         Picture         =   "frmassignacion_catalogo_lista_precios.frx":164A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         Picture         =   "frmassignacion_catalogo_lista_precios.frx":174C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   360
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   5940
         Left            =   45
         TabIndex        =   6
         Top             =   690
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   10478
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
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6526
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Catálogo"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Articulos"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   7
         Top             =   120
         Width           =   7005
      End
   End
End
Attribute VB_Name = "frmasignacion_catalogo_lista_precios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_Click()
   Dim var_canal_venta As String
   Dim var_articulo As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_dia As Integer
   Dim var_mes As Integer
   Dim var_año As Integer
   Dim var_fecha As Date
   If IsDate(txt_fecha_inicio) Then
      If IsDate(txt_fecha_fin) Then
         var_fecha = CDate(txt_fecha_fin)
         var_dia = CStr(Day(var_fecha))
         var_mes = CStr(Month(var_fecha))
         var_año = CStr(Year(var_fecha))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         var_fecha_fin = "{d '" + CStr(var_año) + "-" + CStr(var_mes) + "-" + CStr(var_dia) + "'}"
         
         var_fecha = CDate(txt_fecha_inicio)
         var_dia = CStr(Day(var_fecha))
         var_mes = CStr(Month(var_fecha))
         var_año = CStr(Year(var_fecha))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         var_fecha_inicio = "{d '" + CStr(var_año) + "-" + CStr(var_mes) + "-" + CStr(var_dia) + "'}"
         
         If Trim(txt_catalogo) <> "" Then
            var_n = lv_articulos.ListItems.Count
            var_articulos = 0
            If var_n > 0 Then
               For var_i = 1 To var_n
                   lv_articulos.ListItems.Item(var_i).Selected = True
                   If lv_articulos.selectedItem.SubItems(3) = "*" Then
                      var_articulos = var_articulos + 1
                   End If
               Next var_i
            End If
            
            var_n = Me.lv_canalesventas.ListItems.Count
            var_Canales = 0
            If var_n > 0 Then
               For var_i = 1 To var_n
                   Me.lv_canalesventas.ListItems.Item(var_i).Selected = True
                   If lv_canalesventas.selectedItem.SubItems(2) = "*" Then
                      var_Canales = var_Canales + 1
                   End If
               Next var_i
            End If
            If var_articulos > 0 Then
               If var_Canales > 0 Then
                  var_n_canales = Me.lv_canalesventas.ListItems.Count
                  VAR_N_ARTICULOS = Me.lv_articulos.ListItems.Count
                  VAR_I_CANALES = 0
                  var_i_Articulos = 0
                  For VAR_I_CANALES = 1 To var_n_canales
                  Me.lv_canalesventas.ListItems.Item(VAR_I_CANALES).Selected = True
                  If Me.lv_canalesventas.selectedItem.SubItems(2) = "*" Then
                     var_canal_venta = lv_canalesventas.selectedItem
                     rs.Open "delete from TB_CATALOGOS_CANAL_VENTA where vcha_cat_catalogo_id = '" + txt_catalogo + "' and vcha_can_canal_Venta_id = '" + var_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
                     For var_i_Articulos = 1 To VAR_N_ARTICULOS
                        lv_articulos.ListItems.Item(var_i_Articulos).Selected = True
                        If Me.lv_articulos.selectedItem.SubItems(3) = "*" Then
                           var_articulo = lv_articulos.selectedItem
                           rsaux.Open "insert into TB_CATALOGOS_CANAL_VENTA (vcha_cat_catalogo_id, vcha_can_canal_venta_id, vcha_art_articulo_id) values ('" + txt_catalogo + "', '" + Me.lv_canalesventas.selectedItem + "', '" + lv_articulos.selectedItem + "')"
                        End If
                     Next var_i_Articulos
                  End If
                  Next VAR_I_CANALES
                  For VAR_I_CANALES = 1 To var_n_canales
                     Me.lv_canalesventas.ListItems.Item(VAR_I_CANALES).Selected = True
                     If Me.lv_canalesventas.selectedItem.SubItems(2) = "*" Then
                        var_canal_venta = lv_canalesventas.selectedItem
                        rs.Open "select * from tb_catalogos_vigencias where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cat_catalogo_id = '" + txt_catalogo + "' and vcha_can_canal_venta_id = '" + var_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           rsaux2.Open "update tb_catalogos_vigencias set dtim_vig_fecha_inicio = '" + txt_fecha_inicio + "', dtim_vig_fecha_fin = '" + txt_vigencia_fin + "' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cat_catalogo_id = '" + txt_catalogo + "' and vcha_can_canal_venta_id = '" + var_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux2.Open "insert into tb_catalogos_vigencias (vcha_emp_empresa_id, vcha_cat_catalogo_id, vcha_can_canal_venta_id, dtim_vig_fecha_inicio, dtim_vig_fecha_fin) values ('" + var_empresa + "','" + txt_catalogo + "', '" + var_canal_venta + "', '" + txt_fecha_inicio + "','" + txt_fecha_fin + "')", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rs.Close
                     End If
                  Next VAR_I_CANALES
                  MsgBox "Se a terminado el proceso", vbOKOnly, "ATENCION"
               Else
                  MsgBox "No se a seleccionado ningún canal de venta", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No se a seleccionado ningún artículo", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se a seleccionado un catálogo", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   var_todos_lineas = 1
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_articulos.ListItems.Count
   For i = 1 To n
       lv_articulos.ListItems.Item(i).SubItems(3) = "*"
       lv_articulos.ListItems.Item(i).Bold = True
       lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
   Next
   lv_articulos.Refresh
End Sub

Private Sub Command10_Click()
   var_todos_lineas = 1
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_canalesventas.ListItems.Count
   For i = 1 To n
       lv_canalesventas.ListItems.Item(i).SubItems(2) = "*"
       lv_canalesventas.ListItems.Item(i).Bold = True
       lv_canalesventas.ListItems.Item(i).ForeColor = &HFF0000
       lv_canalesventas.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_canalesventas.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_canalesventas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
       lv_canalesventas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next
   lv_canalesventas.Refresh
End Sub

Private Sub Command2_Click()
   var_todos_lineas = 0
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      lv_articulos.selectedItem.SubItems(3) = ""
      lv_articulos.ListItems.Item(i).Bold = False
      lv_articulos.ListItems.Item(i).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
   Next i
   lv_articulos.Refresh
End Sub

Private Sub Command3_Click()
   If var_todos_lineas = 1 Then
   Else
        var_todos_lineas = 0
   End If
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      If lv_articulos.selectedItem.SubItems(3) = "*" Then
         lv_articulos.selectedItem.SubItems(3) = ""
         lv_articulos.ListItems.Item(i).Bold = False
         lv_articulos.ListItems.Item(i).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      Else
         lv_articulos.selectedItem.SubItems(3) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      End If
   Next i

End Sub

Private Sub Command4_Click()
   var_todos_lineas = 0
   i = lv_articulos.selectedItem.Index
   If lv_articulos.selectedItem.SubItems(3) = "*" Then
      lv_articulos.selectedItem.SubItems(3) = ""
      lv_articulos.ListItems.Item(i).Bold = False
      lv_articulos.ListItems.Item(i).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_articulos.Refresh
   Else
      lv_articulos.selectedItem.SubItems(3) = "*"
      lv_articulos.ListItems.Item(i).Bold = True
      lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_articulos.Refresh
   End If
End Sub

Private Sub Command5_Click()
   If var_todos_lineas = 1 Then
   Else
         var_todos_lineas = 0
   End If
   n = lv_articulos.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_articulos.selectedItem.SubItems(3) = "" And var_rellena = True Then
         lv_articulos.selectedItem.SubItems(3) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_articulos.selectedItem.SubItems(3) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_articulos.selectedItem.SubItems(3) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub Command6_Click()
   If var_todos_lineas = 1 Then
   Else
         var_todos_lineas = 0
   End If
   n = lv_canalesventas.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_canalesventas.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_canalesventas.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_canalesventas.selectedItem.SubItems(2) = "*"
         lv_canalesventas.ListItems.Item(i).Bold = True
         lv_canalesventas.ListItems.Item(i).ForeColor = &HFF0000
         lv_canalesventas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_canalesventas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_canalesventas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_canalesventas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_canalesventas.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_canalesventas.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub Command7_Click()
   var_todos_lineas = 0
   i = lv_canalesventas.selectedItem.Index
   If lv_canalesventas.selectedItem.SubItems(2) = "*" Then
      lv_canalesventas.selectedItem.SubItems(2) = ""
      lv_canalesventas.ListItems.Item(i).Bold = False
      lv_canalesventas.ListItems.Item(i).ForeColor = &H80000012
      lv_canalesventas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_canalesventas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_canalesventas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_canalesventas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_canalesventas.Refresh
   Else
      lv_canalesventas.selectedItem.SubItems(2) = "*"
      lv_canalesventas.ListItems.Item(i).Bold = True
      lv_canalesventas.ListItems.Item(i).ForeColor = &HFF0000
      lv_canalesventas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_canalesventas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_canalesventas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_canalesventas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_canalesventas.Refresh
   End If

End Sub

Private Sub Command8_Click()
   If var_todos_lineas = 1 Then
   Else
        var_todos_lineas = 0
   End If
   n = lv_canalesventas.ListItems.Count
   For i = 1 To n
      lv_canalesventas.ListItems.Item(i).Selected = True
      If lv_canalesventas.selectedItem.SubItems(2) = "*" Then
         lv_canalesventas.selectedItem.SubItems(2) = ""
         lv_canalesventas.ListItems.Item(i).Bold = False
         lv_canalesventas.ListItems.Item(i).ForeColor = &H80000012
         lv_canalesventas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_canalesventas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_canalesventas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_canalesventas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_canalesventas.selectedItem.SubItems(2) = "*"
         lv_canalesventas.ListItems.Item(i).Bold = True
         lv_canalesventas.ListItems.Item(i).ForeColor = &HFF0000
         lv_canalesventas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_canalesventas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_canalesventas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_canalesventas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub Command9_Click()
   var_todos_lineas = 0
   n = lv_canalesventas.ListItems.Count
   For i = 1 To n
      lv_canalesventas.ListItems.Item(i).Selected = True
      lv_canalesventas.selectedItem.SubItems(2) = ""
      lv_canalesventas.ListItems.Item(i).Bold = False
      lv_canalesventas.ListItems.Item(i).ForeColor = &H80000012
      lv_canalesventas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_canalesventas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_canalesventas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_canalesventas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_canalesventas.Refresh

End Sub

Private Sub Form_Load()
   frm_lista.Visible = False
   txt_fecha_inicio = Date
   txt_fecha_fin = Date
   Top = 0
   Left = 0
   rs.Open "select * from tb_articulos order by vcha_art_articulo_id", cnn, adOpenDynamic, adLockOptimistic
   numero_items_agentes = 0
   While Not rs.EOF
      Set list_item = lv_articulos.ListItems.Add(, , rs!vcha_art_articulo_id)
      list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español))
      list_item.SubItems(2) = Trim(IIf(IsNull(rs!vcha_art_catalogo_vigente), "", rs!vcha_art_catalogo_vigente))
      list_item.SubItems(3) = ""
      rs.MoveNext
   Wend
   rs.Close
   rs.Open "select * from tb_canalesventas order by vcha_can_canal_Venta_id", cnn, adOpenDynamic, adLockOptimistic
   numero_items_agentes = 0
   While Not rs.EOF
      Set list_item = Me.lv_canalesventas.ListItems.Add(, , rs!vcha_can_canal_venta_id)
      list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre))
      list_item.SubItems(2) = ""
      rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_asignacion_catalogo_lista_precios)
End Sub

Private Sub lv_articulos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_articulos, ColumnHeader)
End Sub

Private Sub lv_articulos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim numero_lineas As Integer
      Dim numero_seleccionado1 As Integer
      Dim numero_seleccionado2 As Integer
      Dim primera_vez As Boolean
      Dim segunda_vez As Boolean
      Dim i As Integer
      Dim n As Integer
      Dim list_item As ListItem
      n = lv_articulos.ListItems.Count
      i = lv_articulos.selectedItem.Index
      If lv_articulos.ListItems.Item(i).SubItems(3) = "*" Then
      lv_articulos.ListItems.Item(i).SubItems(3) = " "
             lv_articulos.ListItems.Item(i).Bold = False
             lv_articulos.ListItems.Item(i).ForeColor = &H80000012
             lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
             lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
             lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
             lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
          Else
             lv_articulos.ListItems.Item(i).SubItems(3) = "*"
             lv_articulos.ListItems.Item(i).Bold = True
             lv_articulos.ListItems.Item(i).ForeColor = &H8000&
             lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
             lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
             lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
             lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
         End If
      lv_articulos.Refresh
   End If
End Sub

Private Sub lv_canalesventas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_canalesventas, ColumnHeader)
End Sub

Private Sub lv_canalesventas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim numero_lineas As Integer
      Dim numero_seleccionado1 As Integer
      Dim numero_seleccionado2 As Integer
      Dim primera_vez As Boolean
      Dim segunda_vez As Boolean
      Dim i As Integer
      Dim n As Integer
      Dim list_item As ListItem
      n = lv_canalesventas.ListItems.Count
      i = lv_canalesventas.selectedItem.Index
      If lv_canalesventas.ListItems.Item(i).SubItems(2) = "*" Then
      lv_canalesventas.ListItems.Item(i).SubItems(2) = " "
             lv_canalesventas.ListItems.Item(i).Bold = False
             lv_canalesventas.ListItems.Item(i).ForeColor = &H80000012
             lv_canalesventas.ListItems.Item(i).ListSubItems(1).Bold = False
             lv_canalesventas.ListItems.Item(i).ListSubItems(2).Bold = False
             lv_canalesventas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
             lv_canalesventas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
          Else
             lv_canalesventas.ListItems.Item(i).SubItems(2) = "*"
             lv_canalesventas.ListItems.Item(i).Bold = True
             lv_canalesventas.ListItems.Item(i).ForeColor = &H8000&
             lv_canalesventas.ListItems.Item(i).ListSubItems(1).Bold = True
             lv_canalesventas.ListItems.Item(i).ListSubItems(2).Bold = True
             lv_canalesventas.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
             lv_canalesventas.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
         End If
      lv_canalesventas.Refresh
   End If

End Sub

Private Sub txt_canal_venta_Change()

End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_catalogo = lv_lista.selectedItem
         txt_nombre_catalogo = lv_lista.selectedItem.SubItems(1)
         txt_catalogo.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      txt_catalogo.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_catalogo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_catalogos order by vcha_cat_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Cat_catalogo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cat_nombre), "", rs!vcha_cat_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CATALOGOS"
      var_tipo_lista = 1
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

Private Sub txt_catalogo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_catalogo_LostFocus()
   If Trim(txt_catalogo) <> "" Then
      rs.Open "select * from tb_catalogos where vcha_cat_catalogo_id = '" + txt_catalogo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_catalogo = IIf(IsNull(rs!vcha_cat_nombre), "", rs!vcha_cat_nombre)
      Else
         MsgBox "Clave de catálogo incorrecto", vbOKOnly, "ATENCION"
         txt_catalogo = ""
         txt_nombre_catalogo = ""
      End If
      rs.Close
   Else
      txt_nombre_catalogo = ""
   End If
End Sub

Private Sub txt_fecha_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(txt_fecha_fin) Then
         frmcalendario.mes.Value = CDate(txt_fecha_fin)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fecha_fin = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_fin_KeyPress(KeyAscii As Integer)
   Me.cmd_aceptar.SetFocus
End Sub

Private Sub txt_fecha_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(txt_fecha_inicio) Then
         frmcalendario.mes.Value = CDate(txt_fecha_inicio)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fecha_inicio = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_inicio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_catalogo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
      If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub
