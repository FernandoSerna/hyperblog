VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcorreccion_cobranza_aplicada_no_reflejada 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Corrección de cobranza aplicada pero no reflejada en cartera"
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
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11190
      Picture         =   "frmcorreccion_cobranza_aplicada_no_reflejada.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmcorreccion_cobranza_aplicada_no_reflejada.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Actualizar"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   75
      TabIndex        =   1
      Top             =   300
      Width           =   11490
   End
   Begin MSComctlLib.ListView lv_movimientos 
      Height          =   6690
      Left            =   90
      TabIndex        =   0
      Top             =   465
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   11800
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Empresa"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Documento Cargo"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Número Cargo"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Documento Abono"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Serie Abono"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Número Abono"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Importe Abono"
         Object.Width           =   2822
      EndProperty
   End
End
Attribute VB_Name = "frmcorreccion_cobranza_aplicada_no_reflejada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
   Me.lv_movimientos.ListItems.Clear
   rs.Open "SELECT * FROm vw_comportamiento_pagos where inte_ecu_numero_abono is null", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_movimientos.ListItems.Add(, , rs!vcha_emp_empresa_id)
            list_item.SubItems(1) = IIf(IsNull(rs!documento_cargo), "", rs!documento_cargo)
            list_item.SubItems(2) = IIf(IsNull(rs!numero_Cargo), "", rs!numero_Cargo)
            list_item.SubItems(3) = "PA"
            list_item.SubItems(4) = IIf(IsNull(rs!vcha_ser_serie_id), "", rs!vcha_ser_serie_id)
            list_item.SubItems(5) = IIf(IsNull(rs!inte_Car_numero), "", rs!inte_Car_numero)
            list_item.SubItems(6) = IIf(IsNull(rs!floa_car_importe_neto), "", rs!floa_car_importe_neto)
            rs.MoveNext:
      Wend
   End If
   rs.Close
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   rs.Open "SELECT * FROm vw_comportamiento_pagos where inte_ecu_numero_abono is null", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_movimientos.ListItems.Add(, , rs!vcha_emp_empresa_id)
            list_item.SubItems(1) = IIf(IsNull(rs!documento_cargo), "", rs!documento_cargo)
            list_item.SubItems(2) = IIf(IsNull(rs!numero_Cargo), "", rs!numero_Cargo)
            list_item.SubItems(3) = "PA"
            list_item.SubItems(4) = IIf(IsNull(rs!vcha_ser_serie_id), "", rs!vcha_ser_serie_id)
            list_item.SubItems(5) = IIf(IsNull(rs!inte_Car_numero), "", rs!inte_Car_numero)
            list_item.SubItems(6) = IIf(IsNull(rs!floa_car_importe_neto), "", rs!floa_car_importe_neto)
            rs.MoveNext:
      Wend
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_movimientos_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_si = MsgBox("¿Desea corregir el pago?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_cadena = "insert into tb_estado_cuenta (vcha_emp_empresa_id, vcha_ecu_movimiento_cargo, vcha_ecu_serie_cargo, inte_ecu_numero_Cargo, vcha_ecu_movimiento_abono, vcha_ecu_serie_abono, inte_ecu_numero_abono, floa_ecu_importe_abono)"
         var_cadena = var_cadena + "values ('" + lv_movimientos.selectedItem + "','" + lv_movimientos.selectedItem.SubItems(1) + "', '" + lv_movimientos.selectedItem.SubItems(4) + "', " + CStr(CDbl(Me.lv_movimientos.selectedItem.SubItems(2))) + ",'PA','" + Me.lv_movimientos.selectedItem.SubItems(4) + "'," + CStr(CDbl(lv_movimientos.selectedItem.SubItems(5))) + "," + CStr(CDbl(Me.lv_movimientos.selectedItem.SubItems(6))) + ")"
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         Me.lv_movimientos.ListItems.Clear
         rs.Open "SELECT * FROm vw_comportamiento_pagos where inte_ecu_numero_abono is null", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Set list_item = lv_movimientos.ListItems.Add(, , rs!vcha_emp_empresa_id)
                  list_item.SubItems(1) = IIf(IsNull(rs!documento_cargo), "", rs!documento_cargo)
                  list_item.SubItems(2) = IIf(IsNull(rs!numero_Cargo), "", rs!numero_Cargo)
                  list_item.SubItems(3) = "PA"
                  list_item.SubItems(4) = IIf(IsNull(rs!vcha_ser_serie_id), "", rs!vcha_ser_serie_id)
                  list_item.SubItems(5) = IIf(IsNull(rs!inte_Car_numero), "", rs!inte_Car_numero)
                  list_item.SubItems(6) = IIf(IsNull(rs!floa_car_importe_neto), "", rs!floa_car_importe_neto)
                  rs.MoveNext:
            Wend
         End If
         rs.Close
      
      
      End If
   End If
End Sub
