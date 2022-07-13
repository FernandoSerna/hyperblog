VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmprecios_paqueterias_almacen 
   Caption         =   "Precios de paqueteria del almacen"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   4815
      Left            =   90
      TabIndex        =   24
      Top             =   2340
      Width           =   5655
      Begin MSComctlLib.ListView lv_paqueterias 
         Height          =   4635
         Left            =   45
         TabIndex        =   25
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8176
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
            Text            =   "Caja"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre caja"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "clave"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "precio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "costo"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Precios de paqueteria "
      Height          =   1410
      Left            =   90
      TabIndex        =   11
      Top             =   360
      Width           =   5655
      Begin VB.TextBox txt_costo 
         Height          =   315
         Left            =   4755
         MaxLength       =   3
         TabIndex        =   18
         Top             =   915
         Width           =   780
      End
      Begin VB.TextBox txt_paqueteria 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   17
         Top             =   225
         Width           =   630
      End
      Begin VB.TextBox txt_nombre 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1935
         MaxLength       =   50
         TabIndex        =   16
         Top             =   225
         Width           =   3600
      End
      Begin VB.TextBox txt_clave 
         Height          =   315
         Left            =   1290
         TabIndex        =   15
         Top             =   915
         Width           =   915
      End
      Begin VB.TextBox txt_precio 
         Height          =   315
         Left            =   3150
         MaxLength       =   3
         TabIndex        =   14
         Top             =   915
         Width           =   825
      End
      Begin VB.TextBox txt_caja 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   13
         Top             =   570
         Width           =   615
      End
      Begin VB.TextBox txt_nombre_caja 
         Height          =   315
         Left            =   1935
         MaxLength       =   50
         TabIndex        =   12
         Top             =   570
         Width           =   3600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Costo:"
         Height          =   195
         Index           =   3
         Left            =   4275
         TabIndex        =   23
         Top             =   975
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Paqueteria:"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   22
         Top             =   285
         Width           =   810
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Caja:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   21
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   20
         Top             =   975
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Precio:"
         Height          =   195
         Index           =   4
         Left            =   2580
         TabIndex        =   19
         Top             =   975
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   90
      TabIndex        =   7
      Top             =   1785
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1785
         TabIndex        =   8
         Top             =   165
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3735
         TabIndex        =   9
         Top             =   150
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al primero"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un Registro Atras"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un registro adelante"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al ultimo"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de linea:"
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   195
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2550
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   75
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmprecios_paqueterias_almacen.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   -15
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmprecios_paqueterias_almacen.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guardar Alt + G"
      Top             =   -15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      Picture         =   "frmprecios_paqueterias_almacen.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   -15
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   30
      TabIndex        =   0
      Top             =   2535
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   1
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
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7057
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Width           =   5610
      End
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   3015
      Top             =   -15
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
            Picture         =   "frmprecios_paqueterias_almacen.frx":083E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprecios_paqueterias_almacen.frx":1118
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   2565
      Top             =   -330
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
            Picture         =   "frmprecios_paqueterias_almacen.frx":19F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprecios_paqueterias_almacen.frx":22CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprecios_paqueterias_almacen.frx":2BA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprecios_paqueterias_almacen.frx":3142
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprecios_paqueterias_almacen.frx":3A1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprecios_paqueterias_almacen.frx":42F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprecios_paqueterias_almacen.frx":4BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprecios_paqueterias_almacen.frx":4CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprecios_paqueterias_almacen.frx":4DF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprecios_paqueterias_almacen.frx":4F08
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprecios_paqueterias_almacen.frx":501A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   90
      TabIndex        =   26
      Top             =   240
      Visible         =   0   'False
      Width           =   5655
   End
End
Attribute VB_Name = "frmprecios_paqueterias_almacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_guardar_Click()
   Dim list_item As ListItem
   Me.txt_clave.Enabled = False
   If Me.txt_clave <> "" Then
      If Me.txt_caja <> "" Then
         If IsNumeric(Me.txt_precio) Then
            If IsNumeric(Me.txt_costo) Then
               rs.Open "SELECT * FROM TB_PRECIOS_CAJAS WHERE VCHA_PCA_CLAVE_ID = '" + Me.txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_si = MsgBox("¿Desea ejecutar los cambios", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     rsaux2.Open "UPDATE TB_PRECIOS_cAJAS SET FLOA_PCA_PRECIO = " + Me.txt_precio + ", FLOA_PCA_COSTO = " + Me.txt_costo + " WHERE VCHA_PCA_CLAVE_ID = '" + Me.txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
                     lv_paqueterias.ListItems.Item(lv_paqueterias.selectedItem.Index).Checked = False
                     lv_paqueterias.ListItems.Item(lv_paqueterias.selectedItem.Index) = Me.txt_caja
                     lv_paqueterias.ListItems.Item(lv_paqueterias.selectedItem.Index).ListSubItems(1) = Me.txt_nombre_caja
                     lv_paqueterias.ListItems.Item(lv_paqueterias.selectedItem.Index).ListSubItems(2) = Me.txt_clave
                     lv_paqueterias.ListItems.Item(lv_paqueterias.selectedItem.Index).ListSubItems(3) = Me.txt_precio
                     lv_paqueterias.ListItems.Item(lv_paqueterias.selectedItem.Index).ListSubItems(4) = Me.txt_costo
                     lv_paqueterias.ListItems.Item(lv_paqueterias.selectedItem.Index).Selected = True
                     MsgBox "Se han ejecutado los cambios correctamente", vbOKOnly, "ATENCION"
                  End If
               End If
               rs.Close
            Else
               MsgBox "Costo incorrecto", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Precio incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a seleccionado", vbOKOnly, "ATENCION"
      End If
   Else
      rsaux2.Open "select max(cast(vcha_pca_clave_id as integer)) from tb_precios_cajas", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux2.EOF Then
         VAR_CLAVE_NUMERO = IIf(IsNull(rsaux2(0).Value), 0, rsaux2(0).Value)
      Else
         VAR_CLAVE_NUMERO = 0
      End If
      rsaux2.Close
      VAR_CLAVE_NUMERO = VAR_CLAVE_NUMERO + 1
      If VAR_CLAVE_NUMERO < 10 Then
         VAR_CLAVE_PAQUETERIA = "0000" + Trim(CStr(VAR_CLAVE_NUMERO))
      Else
         If VAR_CLAVE_NUMERO < 100 Then
            VAR_CLAVE_PAQUETERIA = "000" + Trim(CStr(VAR_CLAVE_NUMERO))
         Else
            If VAR_CLAVE_NUMERO < 1000 Then
               VAR_CLAVE_PAQUETERIA = "00" + Trim(CStr(VAR_CLAVE_NUMERO))
            Else
              VAR_CLAVE_PAQUETERIA = "0" + Trim(CStr(VAR_CLAVE_NUMERO))
            End If
         End If
      End If
      rs.Open "INSERT INTO TB_PRECIOS_CAJAS (VCHA_PAQ_CLAVE_ID, VCHA_CAJ_CAJA_ID, VCHA_PCA_CLAVE_ID, FLOA_PCA_PRECIO, FLOA_PCA_COSTO) VALUES ('" + var_paqueteria + "','" + Me.txt_caja + "','" + VAR_CLAVE_PAQUETERIA + "'," + Me.txt_precio + "," + Me.txt_costo + ")", cnn, adOpenDynamic, adLockOptimistic
      Me.txt_clave = VAR_CLAVE_PAQUETERIA
      Set list_item = lv_paqueterias.ListItems.Add(, , Me.txt_caja)
      list_item.SubItems(1) = Me.txt_nombre_caja
      list_item.SubItems(2) = VAR_CLAVE_PAQUETERIA
      list_item.SubItems(3) = Me.txt_precio
      list_item.SubItems(4) = Me.txt_costo
      list_item.EnsureVisible
      list_item.Selected = True
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_caja = ""
   Me.txt_nombre_caja = ""
   Me.txt_clave = ""
   Me.txt_precio = ""
   Me.txt_costo = ""
   Me.txt_clave.Enabled = False
   Me.txt_caja.Enabled = True
   Me.txt_caja.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 2900
   Dim list_item As ListItem
   lv_paqueterias.ListItems.Clear
   rs.Open "select * from vw_precios_paqueteria_sid where vcha_paq_clave_id = '" + var_paqueteria + "'", cnn, adOpenDynamic, adLockOptimistic
   numero_items_lineas = 0
   While Not rs.EOF
      Set list_item = lv_paqueterias.ListItems.Add(, , rs!vcha_caj_caja_id)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CAJ_NOMBRE), "", rs!VCHA_CAJ_NOMBRE)
      list_item.SubItems(2) = IIf(IsNull(rs!vcha_pca_clave_id), "", rs!vcha_pca_clave_id)
      list_item.SubItems(3) = IIf(IsNull(rs!floa_pca_precio), "", rs!floa_pca_precio)
      list_item.SubItems(4) = IIf(IsNull(rs!floa_pca_costo), "", rs!floa_pca_costo)
      rs.MoveNext:
    Wend
   rs.Close
   Me.txt_paqueteria.Enabled = False
   Me.txt_caja.Enabled = False
   Me.frm_lista.Visible = False
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_caja = lv_lista.selectedItem
      Me.txt_nombre_caja = lv_lista.selectedItem.SubItems(1)
      Me.txt_caja.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub lv_paqueterias_BeforeLabelEdit(Cancel As Integer)
   Me.txt_clave.Enabled = False
   Me.txt_caja.Enabled = False
   Me.txt_caja = Me.lv_paqueterias.selectedItem
   Me.txt_nombre_caja = Me.lv_paqueterias.selectedItem.SubItems(1)
   Me.txt_clave = Me.lv_paqueterias.selectedItem.SubItems(2)
   Me.txt_precio = Me.lv_paqueterias.selectedItem.SubItems(3)
   Me.txt_costo = Me.lv_paqueterias.selectedItem.SubItems(4)
End Sub

Private Sub lv_paqueterias_GotFocus()
   Me.txt_clave.Enabled = False
   Me.txt_caja.Enabled = False
   Me.txt_caja = Me.lv_paqueterias.selectedItem
   Me.txt_nombre_caja = Me.lv_paqueterias.selectedItem.SubItems(1)
   Me.txt_clave = Me.lv_paqueterias.selectedItem.SubItems(2)
   Me.txt_precio = Me.lv_paqueterias.selectedItem.SubItems(3)
   Me.txt_costo = Me.lv_paqueterias.selectedItem.SubItems(4)
End Sub

Private Sub lv_paqueterias_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Me.txt_clave.Enabled = False
   Me.txt_caja.Enabled = False
   Me.txt_caja = Me.lv_paqueterias.selectedItem
   Me.txt_nombre_caja = Me.lv_paqueterias.selectedItem.SubItems(1)
   Me.txt_clave = Me.lv_paqueterias.selectedItem.SubItems(2)
   Me.txt_precio = Me.lv_paqueterias.selectedItem.SubItems(3)
   Me.txt_costo = Me.lv_paqueterias.selectedItem.SubItems(4)
End Sub

Private Sub txt_caja_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_caj_caja_id, vcha_caj_nombre from tb_cajas order by vcha_caj_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_caj_caja_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CAJ_NOMBRE), "", rs!VCHA_CAJ_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Cajas"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_caja_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_caja_LostFocus()
   If Trim(Me.txt_caja) <> "" Then
      rs.Open "select * from tb_cajas where vcha_caj_Caja_id = '" + Me.txt_caja + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_caja = IIf(IsNull(rs!VCHA_CAJ_NOMBRE), "", rs!VCHA_CAJ_NOMBRE)
      Else
         MsgBox "Clave de caja incorrecta", vbOKOnly, "ATENCION"
         Me.txt_caja = ""
         Me.txt_nombre_caja = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_caja = ""
   End If
End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_costo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_nombre_caja_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_precio_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

