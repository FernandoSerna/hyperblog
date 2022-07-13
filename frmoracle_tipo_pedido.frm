VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_tipo_pedido 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipo de pedido"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lv_lista 
      Height          =   1410
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   2487
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
         Text            =   "Titular"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmoracle_tipo_pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   If var_unidad_organizacional = "90" Then
      If var_clave_usuario_global = "U0000000241" Then
         Set list_item = lv_lista.ListItems.Add(, , "1701")
         list_item.SubItems(1) = "TEX_VTA_DIRECTA_INV"
      Else
         Set list_item = lv_lista.ListItems.Add(, , "1141")
         list_item.SubItems(1) = "TEX_VTA_DIRECTA"
      End If
   End If
   If var_unidad_organizacional = "89" Then
      Set list_item = lv_lista.ListItems.Add(, , "1402")
      list_item.SubItems(1) = "VIA_CANTIA_MUEBLES"
      Set list_item = lv_lista.ListItems.Add(, , "1078")
      list_item.SubItems(1) = "MUE_CANTIA"
      
   End If
   If var_unidad_organizacional = "85" Then
      Set list_item = lv_lista.ListItems.Add(, , "1050")
      list_item.SubItems(1) = "VIA_CANTIA"
   End If
   If var_unidad_organizacional = "94" Then
      Set list_item = lv_lista.ListItems.Add(, , "1562")
      list_item.SubItems(1) = "VIA_CANTIA_VERGEL"
   End If
   If var_unidad_organizacional = "93" Then
      If var_clave_usuario_global = "U0000000763" Then
         Set list_item = lv_lista.ListItems.Add(, , "1042")
         list_item.SubItems(1) = "VIA_MAYOREO_NACIONAL"
         Set list_item = lv_lista.ListItems.Add(, , "2061")
         list_item.SubItems(1) = "VIA_PRIVALIA"
      Else
         Set list_item = lv_lista.ListItems.Add(, , "1042")
         list_item.SubItems(1) = "VIA_MAYOREO_NACIONAL"
      End If
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_tipo_pedido_ventas_directas = CDbl(Me.lv_lista.selectedItem)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_tipo_pedido_ventas_directas = CDbl(Me.lv_lista.selectedItem)
      Unload Me
   End If
   If KeyAscii = 27 Then
      var_tipo_pedido_ventas_directas = CDbl(Me.lv_lista.selectedItem)
      Unload Me
   End If
End Sub
