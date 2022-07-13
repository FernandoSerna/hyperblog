VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_CN_frontera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Centros de Negocio en frontera"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   6720
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   8310
      Begin MSComctlLib.ListView lv_clientes 
         Height          =   6525
         Left            =   60
         TabIndex        =   1
         Top             =   135
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   11509
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
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre Ruta"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "CN"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmoracle_CN_frontera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter


Private Sub Form_Load()
   var_cadena = "select secondary_inventory_name, description from INV.MTL_SECONDARY_INVENTORIES where attribute3 = 'PTO_VTA' order by description"
   rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_clientes.ListItems.Add(, , rs!secondary_inventory_name)
         list_item.SubItems(1) = rs!Description
         list_item.SubItems(2) = ""
         rs.MoveNext
   Wend
   rs.Close
   
   rs.Open "SELECT * FROM TB_ORACLE_CN_FRONTERA", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         For var_j = 1 To Me.lv_clientes.ListItems.Count
             Me.lv_clientes.ListItems(var_j).Selected = True
             If Me.lv_clientes.selectedItem = rs!CLAVE Then
                Me.lv_clientes.selectedItem.SubItems(2) = rs!CLAVE
          lv_clientes.ListItems.Item(var_j).Bold = True
          lv_clientes.ListItems.Item(var_j).ListSubItems(1).Bold = True
          lv_clientes.ListItems.Item(var_j).ForeColor = &H8000&
          lv_clientes.ListItems.Item(var_j).ListSubItems(1).ForeColor = &H8000&
             
             End If
             
         Next var_j
         rs.MoveNext
   Wend
   rs.Close
   
End Sub

Private Sub lv_clientes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_clientes, ColumnHeader)
End Sub

Private Sub lv_clientes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_i = Me.lv_clientes.selectedItem.Index
      If Me.lv_clientes.selectedItem.SubItems(2) = "" Then
         rs.Open "INSERT INTO TB_ORACLE_CN_FRONTERA (CLAVE) VALUES ('" + Me.lv_clientes.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
         Me.lv_clientes.selectedItem.SubItems(2) = Me.lv_clientes.selectedItem
          lv_clientes.ListItems.Item(var_i).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_clientes.ListItems.Item(var_i).ForeColor = &H8000&
          lv_clientes.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
         
         
      Else
         rs.Open "DELETE FROM TB_ORACLE_CN_FRONTERA WHERE CLAVE = '" + Me.lv_clientes.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
         Me.lv_clientes.selectedItem.SubItems(2) = ""
         
          lv_clientes.ListItems.Item(var_i).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_clientes.ListItems.Item(var_i).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF&
         
      
      End If
   End If
End Sub
