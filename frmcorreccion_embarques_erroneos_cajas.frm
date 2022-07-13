VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmcorreccion_embarques_erroneos_cajas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Coreccion de embarques erroneas por cajas"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7170
      Picture         =   "frmcorreccion_embarques_erroneos_cajas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salir"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmcorreccion_embarques_erroneos_cajas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   45
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Movimientos "
      Height          =   2190
      Left            =   90
      TabIndex        =   8
      Top             =   4515
      Width           =   7485
      Begin MSComctlLib.ListView lv_movimientos 
         Height          =   1875
         Left            =   135
         TabIndex        =   9
         Top             =   225
         Width           =   7230
         _ExtentX        =   12753
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Movimiento"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Unidad"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Número"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha Movimiento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Orden Surtido"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Ordenes de Surtido "
      Height          =   2190
      Left            =   90
      TabIndex        =   6
      Top             =   2220
      Width           =   7485
      Begin MSComctlLib.ListView lv_ordenes 
         Height          =   1875
         Left            =   105
         TabIndex        =   7
         Top             =   195
         Width           =   7260
         _ExtentX        =   12806
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Orden"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clave"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   8467
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   60
      TabIndex        =   4
      Top             =   360
      Width           =   7485
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos del embarque "
      Height          =   1695
      Left            =   105
      TabIndex        =   3
      Top             =   450
      Width           =   7455
      Begin VB.TextBox txt_estatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1200
         TabIndex        =   2
         Top             =   1155
         Width           =   480
      End
      Begin VB.TextBox txt_agente 
         Height          =   420
         Left            =   1200
         TabIndex        =   1
         Top             =   705
         Width           =   6000
      End
      Begin VB.TextBox txt_embarque 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1200
         TabIndex        =   0
         Top             =   255
         Width           =   2085
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Estatus:"
         Height          =   195
         Left            =   225
         TabIndex        =   13
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   210
         TabIndex        =   10
         Top             =   750
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   330
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmcorreccion_embarques_erroneos_cajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
   If Trim(Me.txt_estatus) = "E" Then
      If Me.lv_movimientos.ListItems.Count >= 0 Then
         var_si = MsgBox("¿Deseas corregir el embarque?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar la corrección del embarque", vbYesNo, "ATENCION")
            If var_si = 6 Then
               For var_j = 1 To lv_movimientos.ListItems.Count
                   lv_movimientos.ListItems.Item(var_j).Selected = True
                   var_movimiento_str = lv_movimientos.selectedItem
                   Var_unidad = lv_movimientos.selectedItem.SubItems(1)
                   var_numero = lv_movimientos.selectedItem.SubItems(2)
                   rs.Open "update tb_encabezado_movimientos set inte_emo_numero_origen = inte_emo_numero_origen * 1000 where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + Var_unidad + "' and vcha_mov_movimiento_id = '" + var_movimiento_str + "' and inte_emo_numero = " + CStr(var_numero), cnn, adOpenDynamic, adLockOptimistic
               Next var_j
               rs.Open "delete from tb_detalle_embarques where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
               rs.Open "UPDATE TB_dETALLE_cAJAS SET CHAR_PAQ_ESTATUS = 'I' WHERE vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque + " AND CHAR_PAQ_ESTATUS = 'S'", cnn, adOpenDynamic, adLockOptimistic
               MsgBox "Se a corregido el embarque", vbOKOnly, "ATENCION"
               Me.txt_agente = ""
               Me.txt_embarque = ""
               Me.txt_estatus = ""
               Me.lv_movimientos.ListItems.Clear
               Me.lv_ordenes.ListItems.Clear
            Else
               MsgBox "Se a cancelado la correccion del embarque", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Se a cancelado la correccion del embarque", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El embarque no tiene movimientos aun", vbOKOnly, "ATENCION"
      End If
   Else
      If Trim(Me.txt_estatus) = "F" Then
         MsgBox "El embarque ya fue facturado", vbOKOnly, "ATENCION"
      End If
      If Trim(Me.txt_estatus) = "I" Then
         MsgBox "El embarque ya fue cerrado para su envio", vbOKOnly, "ATENCION"
      End If
      If Trim(Me.txt_estatus) = "" Then
         MsgBox "El embarque no a sido cerrado en el modulo de cajas", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 200
   Left = 1500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_embarque_LostFocus()
   If Trim(Me.txt_embarque) <> "" Then
      If IsNumeric(Me.txt_embarque) Then
         Me.lv_ordenes.ListItems.Clear
         Me.lv_movimientos.ListItems.Clear
         Me.txt_estatus = ""
         Me.txt_agente = ""
         rs.Open "select * from tb_encabezado_embarques where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_estatus = IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", rs!CHAR_EMB_ESTATUS)
            rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + rs!vcha_age_Agente_id + "'", cnn, adOpenDynamic, adLockOptimistic
            Me.txt_agente = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
            If Not rsaux.EOF Then
               rsaux1.Open "SELECT DISTINCT INTE_ORS_ORDEN_SURTIDO, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE FROM VW_ORDENES_SURTIDO_CAJAS WHERE vcha_emp_Empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  var_Cadena_ordenes = ""
                  While Not rsaux1.EOF
                        Set list_item = lv_ordenes.ListItems.Add(, , rsaux1!INTE_ORS_ORDEN_SURTIDO)
                        If var_Cadena_ordenes = "" Then
                           var_Cadena_ordenes = CStr(rsaux1!INTE_ORS_ORDEN_SURTIDO)
                        Else
                           var_Cadena_ordenes = var_Cadena_ordenes + "," + CStr(rsaux1!INTE_ORS_ORDEN_SURTIDO)
                        End If
                        list_item.SubItems(1) = IIf(IsNull(rsaux1!vcha_cli_clave_id), "", rsaux1!vcha_cli_clave_id)
                        list_item.SubItems(2) = IIf(IsNull(rsaux1!VCHA_CLI_NOMBRE), "", rsaux1!VCHA_CLI_NOMBRE)
                        rsaux1.MoveNext:
                  Wend
                  var_Cadena_ordenes = "(" + var_Cadena_ordenes + ")"
                  If rsaux2.State = 1 Then
                     rsaux2.Close
                  End If
                  rsaux2.Open "SELECT * FROM tb_encabezado_movimientos WHERE vcha_emp_Empresa_id = '" + var_empresa + "' and inte_emo_numero_origen in " + var_Cadena_ordenes + " and vcha_mov_movimiento_id in ('FA','ET','FT','EX')", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     While Not rsaux2.EOF
                           Set list_item = lv_movimientos.ListItems.Add(, , rsaux2!VCHA_MOV_MOVIMIENTO_ID)
                           list_item.SubItems(1) = IIf(IsNull(rsaux2!VCHA_UOR_UNIDAD_ID), "", rsaux2!VCHA_UOR_UNIDAD_ID)
                           list_item.SubItems(2) = IIf(IsNull(rsaux2!INTE_eMO_NUMERO), "", rsaux2!INTE_eMO_NUMERO)
                           'list_item.SubItems(3) = IIf(IsNull(rsaux2!dtim_emo_fecha), "", rsaux2!dtim_emo_fecha)
                           'list_item.SubItems(4) = IIf(IsNull(rsaux2!INTE_EMO_NUMERO_ORIGEN), "", rsaux2!INTE_EMO_NUMERO_ORIGEN)
                           rsaux2.MoveNext:
                     Wend
                  Else
                     MsgBox "El embarque no contiene movimientos", vbOKOnly, "ATENCION"
                  End If
                  rsaux2.Close
               
               
               
               
               Else
                  MsgBox "El embarque no se empaqueto", vbOKOnly, "ATENCION"
               End If
               rsaux1.Close
            Else
               MsgBox "El embarque tiene un agente incorrecto", vbOKOnly, "ATENCION"
            End If
            rsaux.Close
         Else
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Embarque incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_estatus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub
