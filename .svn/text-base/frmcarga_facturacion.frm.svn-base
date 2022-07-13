VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcarga_facturacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de Facturación"
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
   Begin MSComctlLib.ListView lv_facturas 
      Height          =   6885
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   12144
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cliente"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Anterior"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nombre"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Documento"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Serie"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Número"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Importe"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Comision"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Estatus"
         Object.Width           =   1411
      EndProperty
   End
End
Attribute VB_Name = "frmcarga_facturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Top = 0
   Left = 0
   rs.Open "select * from tb_temp_carga_facturas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Tem_Estatus = ''", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_facturas.ListItems.Add(, , IIf(IsNull(rs!vcha_Cli_clave_id), "", rs!vcha_Cli_clave_id))
         list_item.SubItems(1) = rs!vcha_cli_clave_anterior_id
         rsaux5.Open "select * from tb_clientes where vcha_Cli_clave_id = '" + IIf(IsNull(rs!vcha_Cli_clave_id), "", rs!vcha_Cli_clave_id) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux5.EOF Then
            list_item.SubItems(2) = IIf(IsNull(rsaux5!vcha_cli_nombre), "", rsaux5!vcha_cli_nombre)
         Else
            list_item.SubItems(2) = ""
         End If
         rsaux5.Close
         list_item.SubItems(3) = rs!vcha_Car_documento
         list_item.SubItems(4) = rs!vcha_ser_serie_id
         list_item.SubItems(5) = rs!inte_Car_numero
         list_item.SubItems(6) = rs!floa_tem_importe
         list_item.SubItems(7) = rs!vcha_tem_estatus
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_reporte_valuacion_devoluciones)
End Sub
