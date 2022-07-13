VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmrelacion_cobranza_correccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relacion de Cobranza"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_consecutivo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8835
      TabIndex        =   59
      Top             =   1290
      Width           =   1020
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   2010
      TabIndex        =   53
      Top             =   1035
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   54
         Top             =   495
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
         TabIndex        =   55
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   " Cheque "
      Height          =   780
      Left            =   135
      TabIndex        =   49
      Top             =   2490
      Width           =   8265
      Begin VB.TextBox txt_fecha_cheque 
         Height          =   345
         Left            =   6750
         TabIndex        =   56
         Top             =   285
         Width           =   1170
      End
      Begin VB.TextBox txt_banco_cheque 
         Height          =   345
         Left            =   3375
         TabIndex        =   7
         Top             =   315
         Width           =   540
      End
      Begin VB.TextBox txt_nombre_banco_cheque 
         Height          =   345
         Left            =   3930
         TabIndex        =   8
         Top             =   315
         Width           =   1755
      End
      Begin VB.TextBox txt_cheque 
         Height          =   345
         Left            =   1410
         MaxLength       =   4
         TabIndex        =   6
         Top             =   315
         Width           =   990
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   2835
         TabIndex        =   52
         Top             =   375
         Width           =   510
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   6240
         TabIndex        =   51
         Top             =   375
         Width           =   495
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Cheque:"
         Height          =   195
         Left            =   720
         TabIndex        =   50
         Top             =   375
         Width           =   600
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   " Deposito "
      Height          =   780
      Left            =   120
      TabIndex        =   45
      Top             =   1635
      Width           =   8265
      Begin VB.TextBox txt_numero_deposito 
         Height          =   345
         Left            =   840
         TabIndex        =   1
         Top             =   315
         Width           =   915
      End
      Begin VB.TextBox txt_fecha_deposito 
         Height          =   345
         Left            =   6885
         TabIndex        =   5
         Top             =   300
         Width           =   1170
      End
      Begin VB.TextBox txt_deposito 
         Height          =   345
         Left            =   1770
         TabIndex        =   2
         Top             =   315
         Width           =   1830
      End
      Begin VB.TextBox txt_nombre_banco 
         Height          =   345
         Left            =   4785
         TabIndex        =   4
         Top             =   315
         Width           =   1500
      End
      Begin VB.TextBox txt_banco 
         Height          =   345
         Left            =   4230
         TabIndex        =   3
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Deposito"
         Height          =   195
         Left            =   150
         TabIndex        =   48
         Top             =   375
         Width           =   630
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   6360
         TabIndex        =   47
         Top             =   375
         Width           =   495
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   3690
         TabIndex        =   46
         Top             =   375
         Width           =   510
      End
   End
   Begin VB.CommandButton cmd_cancelar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmrelacion_cobranza correccion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Cancelar Alt + C"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmrelacion_cobranza correccion.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame5 
      Height          =   75
      Left            =   105
      TabIndex        =   43
      Top             =   300
      Width           =   8325
   End
   Begin VB.Frame Frame4 
      Caption         =   " Importe a aplicar "
      Height          =   720
      Left            =   150
      TabIndex        =   39
      Top             =   5640
      Width           =   8265
      Begin VB.TextBox txt_descuento 
         Height          =   315
         Left            =   5835
         MaxLength       =   3
         TabIndex        =   13
         Top             =   270
         Width           =   540
      End
      Begin VB.TextBox txt_importe_aplicar 
         Height          =   315
         Left            =   2070
         TabIndex        =   12
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   6600
         TabIndex        =   42
         Top             =   330
         Width           =   120
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Descuento:"
         Height          =   195
         Left            =   4725
         TabIndex        =   41
         Top             =   330
         Width           =   825
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   1380
         TabIndex        =   40
         Top             =   330
         Width           =   570
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Datos del Documento "
      Height          =   1455
      Left            =   150
      TabIndex        =   26
      Top             =   4140
      Width           =   8250
      Begin VB.TextBox txt_saldo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6150
         TabIndex        =   38
         Top             =   975
         Width           =   1710
      End
      Begin VB.TextBox txt_fecha_factura 
         Enabled         =   0   'False
         Height          =   315
         Left            =   930
         TabIndex        =   37
         Top             =   990
         Width           =   1620
      End
      Begin VB.TextBox txt_importe 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3420
         TabIndex        =   36
         Top             =   975
         Width           =   1605
      End
      Begin VB.TextBox txt_nombre_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2580
         TabIndex        =   35
         Top             =   645
         Width           =   5280
      End
      Begin VB.TextBox txt_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   945
         TabIndex        =   34
         Top             =   645
         Width           =   1620
      End
      Begin VB.TextBox txt_nombre_agente_factura 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2580
         TabIndex        =   33
         Top             =   300
         Width           =   5280
      End
      Begin VB.TextBox txt_agente_factura 
         Enabled         =   0   'False
         Height          =   315
         Left            =   945
         TabIndex        =   32
         Top             =   300
         Width           =   1620
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
         Height          =   195
         Left            =   5610
         TabIndex        =   31
         Top             =   1065
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   150
         TabIndex        =   30
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   2835
         TabIndex        =   29
         Top             =   1005
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   135
         TabIndex        =   28
         Top             =   705
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   375
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Documentos a aplicar "
      Height          =   780
      Left            =   135
      TabIndex        =   16
      Top             =   3300
      Width           =   8265
      Begin VB.TextBox txt_serie 
         Height          =   315
         Left            =   3690
         TabIndex        =   10
         Top             =   300
         Width           =   540
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   5160
         TabIndex        =   11
         Top             =   300
         Width           =   1320
      End
      Begin VB.TextBox txt_documento 
         Height          =   315
         Left            =   1995
         MaxLength       =   5
         TabIndex        =   9
         Top             =   300
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   3150
         TabIndex        =   44
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   4425
         TabIndex        =   25
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
         Height          =   195
         Left            =   960
         TabIndex        =   22
         Top             =   360
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos Relación "
      Height          =   1185
      Left            =   135
      TabIndex        =   0
      Top             =   435
      Width           =   8265
      Begin VB.TextBox txt_fecha_insercion 
         Height          =   345
         Left            =   6750
         TabIndex        =   57
         Top             =   285
         Width           =   1410
      End
      Begin VB.TextBox txt_fecha 
         Height          =   345
         Left            =   3810
         TabIndex        =   24
         Top             =   285
         Width           =   1410
      End
      Begin VB.TextBox txt_nombre_agente 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2685
         TabIndex        =   21
         Top             =   720
         Width           =   5475
      End
      Begin VB.TextBox txt_agente 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1410
         TabIndex        =   20
         Top             =   720
         Width           =   1245
      End
      Begin VB.TextBox txt_relacion 
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
         Height          =   435
         Left            =   1425
         TabIndex        =   18
         Text            =   "0000000000"
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Fecha inserción:"
         Height          =   195
         Left            =   5550
         TabIndex        =   58
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   3210
         TabIndex        =   23
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   735
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Relación:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   360
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmrelacion_cobranza_correccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_lista As Integer
Private Sub cmd_aceptar_pedidos_Click()
   If Trim(txt_documento) <> "" Then
      If Trim(txt_numero) <> "" Then
         If Trim(txt_cliente) <> "" Then
            If IsNumeric(Me.txt_importe_aplicar) Then
               If Trim(Me.txt_Descuento) = "" Then
                  Me.txt_Descuento = "0"
               End If
               If IsNumeric(txt_numero_deposito) Then
                  If IsNumeric(txt_Descuento) Then
                     If Trim(txt_deposito) <> "" Then
                        If Trim(Me.txt_banco) <> "" Then
                           If IsDate(Me.txt_fecha_deposito) Then
                              If Trim(Me.txt_cheque) <> "" Then
                                 If Trim(txt_banco_cheque) <> "" Then
                                    If IsDate(Me.txt_fecha_cheque) Then
                                       var_si = MsgBox("¿Desea aplicar la cobranza?", vbYesNo, "ATENCION")
                                       If var_si = 6 Then
                                          var_si = MsgBox("Confirmar la aplicación del pago", vbYesNo, "ATENCION")
                                          If var_si = 6 Then
                                             If txt_agente <> Me.txt_agente_factura Then
                                                var_si = 0
                                                var_si = MsgBox("la factura no corresponde al agente seleccionado, ¿Desea aplicar el pago?", vbYesNo, "ATENCION")
                                             End If
                                             If var_si = 6 Then
                                                rs.Open "SELECT MAX(INTE_RCO_PARTIDA) FROM TB_RELACION_COBRANZA WHERE VCHA_RCO_FOLIO = '" + txt_relacion + "'", cnn, adOpenDynamic, adLockOptimistic
                                                If Not rs.EOF Then
                                                   var_partida = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                                                Else
                                                   var_partida = 1
                                                End If
                                                rs.Close
                                               
           

                                                rs.Open "EXECUTE RELACION_COBRANZA_DEPOSITO '" + var_empresa + "', '" + var_unidad_organizacional + "', '" + Me.txt_relacion + "', '" + Me.txt_fecha + "', '" + Me.txt_agente + "', '" + Me.txt_cliente + "', '" + txt_cheque + "', '" + Format(Me.txt_fecha_cheque, "Short Date") + "', " + Me.txt_importe_aplicar + ", " + Me.txt_Descuento + ", " + txt_numero + ", 0, 0, " + CStr(var_partida) + ", 0, '" + txt_serie + "', '" + txt_documento + "', '" + txt_banco_cheque + "', '" + txt_deposito + "', '" + Me.txt_fecha_deposito + "', '" + Me.txt_banco + "'," + Me.txt_numero_deposito, cnn, adOpenDynamic, adLockOptimistic
                                                rsaux.Open "select * from vw_Clientes where vcha_cli_clave_id = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                                                If Trim(txt_fecha_insercion) = "" Then
                                                   rs.Open "update tb_relacion_cobranza set dtim_rco_fecha_insercion = null, vcha_gac_grupo_Actual_id = '" + rsaux!vcha_gac_grupo_Actual_id + "', vcha_gre_grupo_real_id = '" + rsaux!vcha_gre_grupo_real_id + "', vcha_tit_titular_id = '" + rsaux!VCHA_TIT_TITULAR_ID + "', vcha_mon_moneda_id = '" + rsaux!vcha_mon_moneda_id + "' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + txt_relacion + "' and inte_rco_partida = " + CStr(var_partida), cnn, adOpenDynamic, adLockOptimistic
                                                Else
                                                   rs.Open "update tb_relacion_cobranza set dtim_rco_fecha_insercion = '" + Format(txt_fecha_insercion, "SHORT DATE") + "', vcha_gac_grupo_Actual_id = '" + rsaux!vcha_gac_grupo_Actual_id + "', vcha_gre_grupo_real_id = '" + rsaux!vcha_gre_grupo_real_id + "', vcha_tit_titular_id = '" + rsaux!VCHA_TIT_TITULAR_ID + "', vcha_mon_moneda_id = '" + rsaux!vcha_mon_moneda_id + "'  where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + txt_relacion + "' and inte_rco_partida = " + CStr(var_partida), cnn, adOpenDynamic, adLockOptimistic
                                                End If
                                                rsaux.Close
                                                MsgBox "El pago se aplico correctamente", vbOKOnly, "ATENCION"
                                             End If
                                          End If
                                       End If
                                    Else
                                       MsgBox "Fecha de cheque invalida", vbOKOnly, "ATENCION"
                                    End If
                                 Else
                                    MsgBox "Debe de indicar un banco para el cheque", vbOKOnly, "ATENCION"
                                 End If
                              Else
                                 MsgBox "Debe de indicar un cheque", vbOKOnly, "ATENCION"
                              End If
                           Else
                              MsgBox "Se debe de indicar la fecha del deposito", vbOKOnly, "ATENCION"
                           End If
                        Else
                           MsgBox "Se debe de indicar el banco del deposito", vbOKOnly, "ATENCION"
                        End If
                     Else
                        MsgBox "Se debe de indicar un deposito", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "Descuento invalido", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "Número de deposito incorrecto", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Importe a aplicar incorrecto"
            End If
         Else
            MsgBox "Clave de cliente incorrecta", vbOKOnly, "TENCION"
         End If
      Else
         MsgBox "Numero de documento incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un tipo de documento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_cancelar_pedidos_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
    frm_lista.Visible = False
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         If var_tipo_lista = 3 Then
            txt_agente = lv_lista.selectedItem
            txt_nombre_agente = lv_lista.selectedItem.SubItems(1)
         End If
         If var_tipo_lista = 1 Then
            txt_banco = lv_lista.selectedItem
            txt_nombre_banco = lv_lista.selectedItem.SubItems(1)
         End If
         If var_tipo_lista = 2 Then
            txt_banco_cheque = lv_lista.selectedItem
            txt_nombre_banco_cheque = lv_lista.selectedItem.SubItems(1)
         End If
         If var_tipo_lista = 4 Then
            Me.txt_cliente = lv_lista.selectedItem
            Me.txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
            Me.txt_agente_factura = Me.txt_agente
            Me.txt_nombre_agente_factura = Me.txt_nombre_agente
         End If
      Else
         txt_agente = ""
         txt_nombre_agente = ""
      End If
      If var_tipo_lista = 3 Then
         txt_agente.SetFocus
      End If
      If var_tipo_lista = 1 Then
         txt_banco.SetFocus
      End If
      If var_tipo_lista = 2 Then
         txt_banco_cheque.SetFocus
      End If
      If var_tipo_lista = 4 Then
         Me.txt_importe_aplicar.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      If var_tipo_lista = 3 Then
         txt_agente.SetFocus
      End If
      If var_tipo_lista = 1 Then
         txt_banco.SetFocus
      End If
      If var_tipo_lista = 2 Then
         txt_banco_cheque.SetFocus
      End If
      If var_tipo_lista = 4 Then
         Me.txt_documento = ""
         Me.txt_documento.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_AGENTES where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_AGE_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_age_agente_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
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

Private Sub txt_banco_cheque_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_bancos order by vcha_ban_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ban_banco_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_BAN_NOMBRE), "", rs!VCHA_BAN_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "BANCO DEL CHEQUE"
      var_tipo_lista = 2
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

Private Sub txt_banco_cheque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_banco_cheque_LostFocus()
   If Trim(txt_banco_cheque) <> "" Then
      rs.Open "SELECT * FROM TB_BANCOS WHERE VCHA_BAN_BANCO_ID = '" + Me.txt_banco_cheque + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_banco_cheque = IIf(IsNull(rs!VCHA_BAN_NOMBRE), "", rs!VCHA_BAN_NOMBRE)
      Else
         MsgBox "Clave de banco incorrecto", vbOKOnly, "ATENCION"
         Me.txt_banco_cheque = ""
         Me.txt_nombre_banco_cheque = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_banco_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_bancos where vcha_ban_banco_id = '20' or vcha_ban_banco_id = '11' or vcha_ban_banco_id = '10' or vcha_ban_banco_id = '22' order by vcha_ban_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ban_banco_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_BAN_NOMBRE), "", rs!VCHA_BAN_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "BANCO DEL DEPOSITO"
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

Private Sub txt_banco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_banco_LostFocus()
   If Trim(txt_banco) <> "" Then
      rs.Open "SELECT * FROM TB_BANCOS WHERE VCHA_BAN_BANCO_ID = '" + Me.txt_banco + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         If Trim(rs!vcha_ban_banco_id) = "22" Or Trim(rs!vcha_ban_banco_id) = "20" Or Trim(rs!vcha_ban_banco_id) = "11" Or Trim(rs!vcha_ban_banco_id) = "10" Then
            txt_nombre_banco = IIf(IsNull(rs!VCHA_BAN_NOMBRE), "", rs!VCHA_BAN_NOMBRE)
         Else
            MsgBox "Banco incorrecto", vbOKOnly, "ATENCION"
            txt_banco = ""
            Me.txt_nombre_banco = ""
         End If
      Else
         MsgBox "Clave de banco incorrecto", vbOKOnly, "ATENCION"
         Me.txt_banco = ""
         Me.txt_nombre_banco = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_cheque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_deposito_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_descuento_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_documento_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_documento_LostFocus()
   If Trim(txt_documento) <> "" Then
      If Trim(txt_documento) = "FA" Or Trim(txt_documento) = "NC" Or Trim(txt_documento) = "CH" Or Trim(txt_documento) = "CR" Then
      Else
         If Trim(txt_documento) = "SALDO" Then
            Me.txt_serie = ""
            Me.txt_numero = 0
            lv_lista.ListItems.Clear
            rs.Open "select distinct vcha_cli_clave_id, vcha_cli_nombre from tb_clientes where vcha_age_agente_id = '" + txt_agente + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
            Dim list_item As ListItem
            Dim var_contador_lista As Integer
            If Not rs.EOF Then
               While Not rs.EOF
                     var_contador_lista = var_contador_lista + 1
                     Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
                     list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre))
                     rs.MoveNext
               Wend
            End If
            frm_lista.Visible = True
            lv_lista.SetFocus
            var_tipo_lista = 4
            rs.Close
         Else
            MsgBox "Documento incorrecto", vbOKOnly, "ATENCION"
            txt_documento = ""
         End If
      End If
   End If
End Sub

Private Sub txt_fecha_cheque_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha_cheque) Then
         frmcalendario.mes.Value = CDate(Me.txt_fecha_cheque)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fecha_cheque = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_cheque_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_fecha_deposito_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha_deposito) Then
         frmcalendario.mes.Value = CDate(Me.txt_fecha_deposito)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fecha_deposito = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_deposito_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_importe_aplicar_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_banco_cheque_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_bancos order by vcha_ban_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ban_banco_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_BAN_NOMBRE), "", rs!VCHA_BAN_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "BANCO DEL CHEQUE"
      var_tipo_lista = 2
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

Private Sub txt_nombre_banco_cheque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_banco_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_bancos where vcha_ban_banco_id = '20' or vcha_ban_banco_id = '11' or vcha_ban_banco_id = '10' or vcha_ban_banco_id = '22' order by vcha_ban_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ban_banco_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_BAN_NOMBRE), "", rs!VCHA_BAN_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "BANCO DEL DEPOSITO"
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

Private Sub txt_nombre_banco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_numero_deposito_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_LostFocus()
   If Trim(txt_documento) <> "" Then
      If IsNumeric(txt_numero) Then
         rs.Open "SELECT * FROM TB_ENCABEZADO_cARTERA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_cAR_DOCUMENTO = '" + txt_documento + "' AND VCHA_sER_SERIE_ID = '" + txt_serie + "' AND INTE_cAR_NUMERO = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_importe = Format(IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto) / IIf(IsNull(rs!FLOA_cAR_TIPO_cAMBIO), 1, rs!FLOA_cAR_TIPO_cAMBIO), "###,###,##0.00")
            txt_agente_factura = IIf(IsNull(rs!vcha_age_agente_id), "", rs!vcha_age_agente_id)
            txt_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
            If rsaux4.State = 1 Then
               rsaux4.Close
            End If
            rsaux4.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + txt_agente_factura + "'", cnn, adOpenDynamic, adLockOptimistic
            txt_nombre_agente_factura = IIf(IsNull(rsaux4!vcha_age_nombre), "", rsaux4!vcha_age_nombre)
            rsaux4.Close
            rsaux4.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
            txt_nombre_cliente = IIf(IsNull(rsaux4!vcha_cli_nombre), "", rsaux4!vcha_cli_nombre)
            rsaux4.Close
            rsaux4.Open "SELECT * FROM TB_SALDOS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_cAR_DOCUMENTO = '" + txt_documento + "' AND VCHA_sER_SERIE_ID = '" + txt_serie + "' AND INTE_cAR_NUMERO = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
            txt_saldo = Format(IIf(IsNull(rsaux4!floa_sal_importe), 0, rsaux4!floa_sal_importe), "###,###,##0.00")
            rsaux4.Close
            txt_fecha_factura = IIf(IsNull(rs!DTIM_car_FECHA), "", rs!DTIM_car_FECHA)
         Else
            MsgBox "El documento no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Número de documento incorrecto", vbOKOnly, "ATENCION"
         txt_numero = ""
      End If
   Else
      MsgBox "Falta indicar el documento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub
