VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_dividir_pedido 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dividir pedido"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frm_embarque 
      Height          =   1080
      Left            =   8280
      TabIndex        =   4
      Top             =   3000
      Width           =   2895
      Begin VB.TextBox txt_embarque 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   60
         MaxLength       =   10
         TabIndex        =   5
         Top             =   480
         Width           =   2745
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Embarque a cambiar"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   2850
      End
   End
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   15045
   End
   Begin VB.TextBox txt_pedido 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin MSComctlLib.ListView lv_cajas 
      Height          =   8220
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   14499
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
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "          Código"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Pedido"
         Object.Width           =   1605
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Agente"
         Object.Width           =   4198
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cliente"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Cantidad"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "estatus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Caja"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Tipo empaque"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Caja"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Sello"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Guia"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "C. Nueva"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Marca cambiar embarque"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Embarque"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Estatus"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Pedido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   1095
   End
End
Attribute VB_Name = "frmoracle_dividir_pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_seleccion_Click()

End Sub

Private Sub cmd_guardar_Click()

End Sub

Private Sub Form_Load()
   Me.frm_embarque.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_cajas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call pro_ordena_listas(Me.lv_cajas, ColumnHeader)
End Sub

Private Sub lv_cajas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.frm_embarque.Visible = True
      Me.txt_embarque = ""
      Me.txt_embarque.SetFocus
   End If
End Sub

Private Sub lv_cajas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_cajas.ListItems.Count > 0 Then
         i = lv_cajas.selectedItem.Index
         If lv_cajas.selectedItem.SubItems(12) = "*" Then
            lv_cajas.selectedItem.SubItems(12) = ""
            lv_cajas.ListItems.Item(i).Bold = False
            lv_cajas.ListItems.Item(i).ForeColor = &H80000012
            lv_cajas.ListItems.Item(i).ListSubItems(1).Bold = False
            lv_cajas.ListItems.Item(i).ListSubItems(2).Bold = False
            lv_cajas.ListItems.Item(i).ListSubItems(3).Bold = False
            lv_cajas.ListItems.Item(i).ListSubItems(4).Bold = False
            lv_cajas.ListItems.Item(i).ListSubItems(5).Bold = False
            lv_cajas.ListItems.Item(i).ListSubItems(6).Bold = False
            lv_cajas.ListItems.Item(i).ListSubItems(7).Bold = False
            lv_cajas.ListItems.Item(i).ListSubItems(8).Bold = False
            lv_cajas.ListItems.Item(i).ListSubItems(9).Bold = False
            lv_cajas.ListItems.Item(i).ListSubItems(10).Bold = False
            lv_cajas.ListItems.Item(i).ListSubItems(11).Bold = False
            lv_cajas.ListItems.Item(i).ListSubItems(12).Bold = False
            lv_cajas.ListItems.Item(i).ListSubItems(13).Bold = False
            lv_cajas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
            lv_cajas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
            lv_cajas.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
            lv_cajas.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
            lv_cajas.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
            lv_cajas.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
            lv_cajas.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
            lv_cajas.ListItems.Item(i).ListSubItems(8).ForeColor = &H80000012
            lv_cajas.ListItems.Item(i).ListSubItems(9).ForeColor = &H80000012
            lv_cajas.ListItems.Item(i).ListSubItems(10).ForeColor = &H80000012
            lv_cajas.ListItems.Item(i).ListSubItems(11).ForeColor = &H80000012
            lv_cajas.ListItems.Item(i).ListSubItems(12).ForeColor = &H80000012
            lv_cajas.ListItems.Item(i).ListSubItems(13).ForeColor = &H80000012
            lv_cajas.Refresh
            rs.Open "UPDATE TB_ORACLE_CAJAS_ADUANA SET MARCA_CAMBIO_EMBARQUE_2 = '' WHERE CAJA = '" + Me.lv_cajas.selectedItem + "' AND EMBARQUE = " + Me.lv_cajas.selectedItem.SubItems(13), cnn, adOpenDynamic, adLockOptimistic
         Else
            lv_cajas.selectedItem.SubItems(12) = "*"
            lv_cajas.ListItems.Item(i).Bold = True
            lv_cajas.ListItems.Item(i).ForeColor = &HFF0000
            lv_cajas.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_cajas.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_cajas.ListItems.Item(i).ListSubItems(3).Bold = True
            lv_cajas.ListItems.Item(i).ListSubItems(4).Bold = True
            lv_cajas.ListItems.Item(i).ListSubItems(5).Bold = True
            lv_cajas.ListItems.Item(i).ListSubItems(6).Bold = True
            lv_cajas.ListItems.Item(i).ListSubItems(7).Bold = True
            lv_cajas.ListItems.Item(i).ListSubItems(8).Bold = True
            lv_cajas.ListItems.Item(i).ListSubItems(9).Bold = True
            lv_cajas.ListItems.Item(i).ListSubItems(10).Bold = True
            lv_cajas.ListItems.Item(i).ListSubItems(11).Bold = True
            lv_cajas.ListItems.Item(i).ListSubItems(12).Bold = True
            lv_cajas.ListItems.Item(i).ListSubItems(13).Bold = True
            lv_cajas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_cajas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_cajas.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
            lv_cajas.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
            lv_cajas.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
            lv_cajas.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
            lv_cajas.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
            lv_cajas.ListItems.Item(i).ListSubItems(8).ForeColor = &HFF0000
            lv_cajas.ListItems.Item(i).ListSubItems(9).ForeColor = &HFF0000
            lv_cajas.ListItems.Item(i).ListSubItems(10).ForeColor = &HFF0000
            lv_cajas.ListItems.Item(i).ListSubItems(11).ForeColor = &HFF0000
            lv_cajas.ListItems.Item(i).ListSubItems(12).ForeColor = &HFF0000
            lv_cajas.ListItems.Item(i).ListSubItems(13).ForeColor = &HFF0000
            lv_cajas.Refresh
            rs.Open "UPDATE TB_ORACLE_CAJAS_ADUANA SET MARCA_CAMBIO_EMBARQUE_2 = '*' WHERE CAJA = '" + Me.lv_cajas.selectedItem + "' AND EMBARQUE = " + Me.lv_cajas.selectedItem.SubItems(13), cnn, adOpenDynamic, adLockOptimistic
         End If
      End If
   End If
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_embarque.Visible = False
   End If
   Dim var_i As Integer
   If KeyAscii = 13 Then
      var_si = MsgBox("¿Desea reasignar de embarque a las cajas seleccionadas?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar el cambio de embarque", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rs.Open "select MAX(EMBARQUE) AS MAXIMO_EMBARQUE from XXVIA_TB_ENCABEZADO_EMBARQUES", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If rs.EOF Then
               var_numero_embarque = 1
            Else
               var_numero_embarque = IIf(IsNull(rs!maximo_embarque), 0, rs!maximo_embarque) + 1
            End If
            rs.Close
            Dim txt_numero As String
            'txt_numero = CStr(var_numero_embarque)
            rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            'var_cadena = "insert into xxvia_tb_encabezado_embarques (EMBARQUE,               JAULA,              VEHICULO,               AGENTE,          FECHA_INICIO, CHAR_EMB_ESTATUS, CHOFER,  BLOQUEADO, BLOQUEADO_POR, TIPO_EMBARQUE, MAQUINA, USUARIO, ARREGLO_0, ARREGLO_1, ORGANIZACION, transporte, agentes) "
            'var_cadena = var_cadena + " select "+txt_numero +",               JAULA,              VEHICULO,               AGENTE,          FECHA_INICIO, CHAR_EMB_ESTATUS, CHOFER,  BLOQUEADO, BLOQUEADO_POR, TIPO_EMBARQUE, MAQUINA, USUARIO, ARREGLO_0, ARREGLO_1, ORGANIZACION, transporte, agentes from xxvia_Tb_encabezado_embarques where embarque = " + CStr(Me.lv_cajas.ListItems.Item(8).Text)
            'rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         
         
            For var_i = 1 To Me.lv_cajas.ListItems.Count
                Me.lv_cajas.ListItems.Item(var_i).Selected = True
                If Me.lv_cajas.selectedItem.SubItems(12) = "*" Then
                    'MsgBox Me.lv_cajas.selectedItem.SubItems(6)
                   rs.Open "UPDATE TB_ORACLE_CAJAS_ADUANA SET EMBARQUE_ANTERIOR = '" + Me.lv_cajas.selectedItem.SubItems(13) + "' WHERE CAJA = '" + Me.lv_cajas.selectedItem + "' AND EMBARQUE = " + Me.lv_cajas.selectedItem.SubItems(13), cnn, adOpenDynamic, adLockOptimistic
                   rs.Open "UPDATE TB_ORACLE_CAJAS_ADUANA SET EMBARQUE = " + CStr(Me.txt_embarque) + ", EMBARQUE_ACTUAL = '" + Me.txt_embarque + "' WHERE CAJA = '" + Me.lv_cajas.selectedItem + "' AND EMBARQUE = " + Me.lv_cajas.selectedItem.SubItems(13), cnn, adOpenDynamic, adLockOptimistic
                   var_cadena = "insert into TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES (agente, nombre_agente, pedido, cliente, piezas, embarque, dia, mes, año, ORDEN_PEDIDO, ESTATUS, ESTATUS_PEDIDO, ESTACION, volumen, CANTIDAD_SIN_CATALOGOS, CANTIDAD_CATALOGOS, ORGANIZACION, PAQUETERIA, EMBARQUE_ANTERIOR, EMBARQUE_ACTUAL) "
                   var_cadena = var_cadena + " select agente, nombre_agente, pedido, cliente, piezas, '" + Me.txt_embarque + "', dia, mes, año, ORDEN_PEDIDO, '', ESTATUS_PEDIDO, ESTACION, volumen, CANTIDAD_SIN_CATALOGOS, CANTIDAD_CATALOGOS, ORGANIZACION, PAQUETERIA, EMBARQUE_ANTERIOR, EMBARQUE_ACTUAL from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where pedido = " + Me.txt_pedido + " and embarque = " + Me.lv_cajas.selectedItem.SubItems(13)
                   rsaux.Open "select * from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where pedido = " + Me.txt_pedido + " and embarque = " + (CStr(Me.txt_embarque)), cnn, adOpenDynamic, adLockOptimistic
                   If rsaux.EOF Then
                      rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                   End If
                   rsaux.Close
                   rs.Open "update xxvia_tb_salidas_Cajas set inte_emb_embarque = " + CStr(Me.txt_embarque) + ", embarque_Actual = " + Me.lv_cajas.selectedItem.SubItems(13) + " where source_header_number = " + Me.txt_pedido + " and inte_paq_caja = " + Me.lv_cajas.selectedItem.SubItems(6), cnnoracle_4, adOpenDynamic, adLockOptimistic
                   'rs.Open "update xxvia_tb_encabezado_embarques set char_emb_estatus= 'E' where embarque = " + CStr(txt_embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
                End If
            Next var_i
            Me.lv_cajas.SetFocus
            Me.frm_embarque.Visible = False
         End If
      End If
   End If
End Sub

Private Sub txt_embarque_LostFocus()
   Me.frm_embarque.Visible = False
End Sub

Private Sub txt_pedido_Change()
   Me.lv_cajas.ListItems.Clear
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_pedido) Then
         If rs.State = 1 Then
            rs.Close
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
         If rsaux4.State = 1 Then
            rsaux4.Close
         End If
         If rsaux5.State = 1 Then
            rsaux5.Close
         End If
         If rsaux6.State = 1 Then
            rsaux6.Close
         End If
         If rsaux7.State = 1 Then
            rsaux7.Close
         End If
         If rsaux8.State = 1 Then
            rsaux8.Close
         End If
         If rsaux9.State = 1 Then
            rsaux9.Close
         End If
         If rsaux10.State = 1 Then
            rsaux10.Close
         End If
         If rsaux11.State = 1 Then
            rsaux11.Close
         End If
         If rsaux12.State = 1 Then
            rsaux12.Close
         End If
         If rsaux13.State = 1 Then
            rsaux13.Close
         End If
         If rsaux14.State = 1 Then
            rsaux14.Close
         End If
         If rsaux15.State = 1 Then
            rsaux15.Close
         End If
   
   
         rs.Open "select embarque, AGENTE, nombre_agente, pedido, cliente, orden_pedido, ESTATUS, estatus_pedido, isnull(paqueteria,0) paqueteria from tb_oracle_pedidos_asignados_embarques WHERE pedido= " + Me.txt_pedido + " order by orden_pedido, pedido", cnn, adOpenDynamic, adLockOptimistic
         var_Cadena_embarques = ""
         While Not rs.EOF
               If var_Cadena_embarques = "" Then
                  var_Cadena_embarques = CStr(rs!Embarque)
               Else
                  var_Cadena_embarques = var_Cadena_embarques + ", " + CStr(rs!Embarque)
               End If
               rs.MoveNext
         Wend
         If rs.RecordCount > 0 Then
            rs.MoveFirst
         End If
         If Not rs.EOF Then
            Me.lv_cajas.ListItems.Clear
            var_pedido = rs!pedido
            rsaux10.Open "select * from tb_oracle_cajas_aduana where pedido = " + Me.txt_pedido + " and embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux10.EOF
                  'rsaux12.Open "select AGENTE, nombre_agente, pedido, cliente, orden_pedido, ESTATUS, estatus_pedido, isnull(paqueteria,0) paqueteria from tb_oracle_pedidos_asignados_embarques WHERE pedido = " + Me.txt_pedido + " order by orden_pedido, pedido", cnn, adOpenDynamic, adLockOptimistic
                  Set list_item = Me.lv_cajas.ListItems.Add(, , rsaux10!Caja)
                  list_item.SubItems(1) = IIf(IsNull(rsaux10!pedido), "", rsaux10!pedido)
                  list_item.SubItems(2) = IIf(IsNull(rsaux10!Agente), "", rsaux10!Agente)
                  list_item.SubItems(3) = IIf(IsNull(rsaux10!Cliente), "", rsaux10!Cliente)
                  list_item.SubItems(4) = Format(rsaux10!PIEZAS, "###,###,##0.00")
                  list_item.SubItems(5) = IIf(IsNull(rsaux10!estatus), "", rsaux10!estatus)
                  list_item.SubItems(6) = IIf(IsNull(rsaux10!numero_caja), "", rsaux10!numero_caja)
                  list_item.SubItems(7) = IIf(IsNull(rsaux10!TIPO_EMPAQUE), "", rsaux10!TIPO_EMPAQUE)
                  list_item.SubItems(8) = IIf(IsNull(rsaux10!caja_pedido), "", rsaux10!caja_pedido)
                  list_item.SubItems(9) = IIf(IsNull(rsaux10!sello), "", rsaux10!sello)
                  list_item.SubItems(11) = IIf(IsNull(rsaux10!caja_actual), "", rsaux10!caja_actual)
                  list_item.SubItems(12) = IIf(IsNull(rsaux10!MARCA_CAMBIO_EMBARQUE_2), "", rsaux10!MARCA_CAMBIO_EMBARQUE_2)
                  list_item.SubItems(13) = IIf(IsNull(rsaux10!Embarque), "", rsaux10!Embarque)
                  list_item.SubItems(14) = IIf(IsNull(rsaux10!estatus), "", rsaux10!estatus)
                     
                  rsaux10.MoveNext
            Wend
            rsaux10.Close
            
            For i = 1 To Me.lv_cajas.ListItems.Count
                Me.lv_cajas.ListItems.Item(i).Selected = True
                If lv_cajas.selectedItem.SubItems(12) = "" Then
                   lv_cajas.ListItems.Item(i).Bold = False
                   lv_cajas.ListItems.Item(i).ForeColor = &H80000012
                   lv_cajas.ListItems.Item(i).ListSubItems(1).Bold = False
                   lv_cajas.ListItems.Item(i).ListSubItems(2).Bold = False
                   lv_cajas.ListItems.Item(i).ListSubItems(3).Bold = False
                   lv_cajas.ListItems.Item(i).ListSubItems(4).Bold = False
                   lv_cajas.ListItems.Item(i).ListSubItems(5).Bold = False
                   lv_cajas.ListItems.Item(i).ListSubItems(6).Bold = False
                   lv_cajas.ListItems.Item(i).ListSubItems(7).Bold = False
                   lv_cajas.ListItems.Item(i).ListSubItems(8).Bold = False
                   lv_cajas.ListItems.Item(i).ListSubItems(9).Bold = False
                   lv_cajas.ListItems.Item(i).ListSubItems(10).Bold = False
                   lv_cajas.ListItems.Item(i).ListSubItems(11).Bold = False
                   lv_cajas.ListItems.Item(i).ListSubItems(12).Bold = False
                   lv_cajas.ListItems.Item(i).ListSubItems(13).Bold = False
                   lv_cajas.ListItems.Item(i).ListSubItems(14).Bold = False
                   lv_cajas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
                   lv_cajas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
                   lv_cajas.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
                   lv_cajas.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
                   lv_cajas.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
                   lv_cajas.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
                   lv_cajas.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
                   lv_cajas.ListItems.Item(i).ListSubItems(8).ForeColor = &H80000012
                   lv_cajas.ListItems.Item(i).ListSubItems(9).ForeColor = &H80000012
                   lv_cajas.ListItems.Item(i).ListSubItems(10).ForeColor = &H80000012
                   lv_cajas.ListItems.Item(i).ListSubItems(11).ForeColor = &H80000012
                   lv_cajas.ListItems.Item(i).ListSubItems(12).ForeColor = &H80000012
                   lv_cajas.ListItems.Item(i).ListSubItems(13).ForeColor = &H80000012
                   lv_cajas.ListItems.Item(i).ListSubItems(14).ForeColor = &H80000012
                   lv_cajas.Refresh
                Else
                   lv_cajas.ListItems.Item(i).Bold = True
                   lv_cajas.ListItems.Item(i).ForeColor = &HFF0000
                   lv_cajas.ListItems.Item(i).ListSubItems(1).Bold = True
                   lv_cajas.ListItems.Item(i).ListSubItems(2).Bold = True
                   lv_cajas.ListItems.Item(i).ListSubItems(3).Bold = True
                   lv_cajas.ListItems.Item(i).ListSubItems(4).Bold = True
                   lv_cajas.ListItems.Item(i).ListSubItems(5).Bold = True
                   lv_cajas.ListItems.Item(i).ListSubItems(6).Bold = True
                   lv_cajas.ListItems.Item(i).ListSubItems(7).Bold = True
                   lv_cajas.ListItems.Item(i).ListSubItems(8).Bold = True
                   lv_cajas.ListItems.Item(i).ListSubItems(9).Bold = True
                   lv_cajas.ListItems.Item(i).ListSubItems(10).Bold = True
                   lv_cajas.ListItems.Item(i).ListSubItems(11).Bold = True
                   lv_cajas.ListItems.Item(i).ListSubItems(12).Bold = True
                   lv_cajas.ListItems.Item(i).ListSubItems(13).Bold = True
                   lv_cajas.ListItems.Item(i).ListSubItems(14).Bold = True
                   lv_cajas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
                   lv_cajas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
                   lv_cajas.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
                   lv_cajas.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
                   lv_cajas.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
                   lv_cajas.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
                   lv_cajas.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
                   lv_cajas.ListItems.Item(i).ListSubItems(8).ForeColor = &HFF0000
                   lv_cajas.ListItems.Item(i).ListSubItems(9).ForeColor = &HFF0000
                   lv_cajas.ListItems.Item(i).ListSubItems(10).ForeColor = &HFF0000
                   lv_cajas.ListItems.Item(i).ListSubItems(11).ForeColor = &HFF0000
                   lv_cajas.ListItems.Item(i).ListSubItems(12).ForeColor = &HFF0000
                   lv_cajas.ListItems.Item(i).ListSubItems(13).ForeColor = &HFF0000
                   lv_cajas.ListItems.Item(i).ListSubItems(14).ForeColor = &HFF0000
                   lv_cajas.Refresh
                End If
            Next i
         Else
            MsgBox "Pedido no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
       Else
          MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
       End If
   End If
End Sub
