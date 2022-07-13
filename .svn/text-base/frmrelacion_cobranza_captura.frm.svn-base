VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmrelacion_cobranza_captura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Captura de Cobranza"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   2505
      TabIndex        =   46
      Top             =   1020
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   47
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
         TabIndex        =   48
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   " Deposito "
      Height          =   780
      Left            =   105
      TabIndex        =   54
      Top             =   1740
      Width           =   8265
      Begin VB.TextBox txt_numero_deposito 
         Height          =   345
         Left            =   855
         TabIndex        =   4
         Top             =   300
         Width           =   915
      End
      Begin VB.TextBox txt_banco 
         Height          =   345
         Left            =   4395
         TabIndex        =   6
         Top             =   300
         Width           =   540
      End
      Begin VB.TextBox txt_nombre_banco 
         Height          =   345
         Left            =   4950
         TabIndex        =   7
         Top             =   300
         Width           =   1560
      End
      Begin VB.TextBox txt_deposito 
         Height          =   345
         Left            =   1785
         TabIndex        =   5
         Top             =   300
         Width           =   1995
      End
      Begin VB.TextBox txt_fecha_deposito 
         Height          =   345
         Left            =   7050
         TabIndex        =   8
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   3855
         TabIndex        =   57
         Top             =   375
         Width           =   510
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   6525
         TabIndex        =   56
         Top             =   375
         Width           =   495
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Deposito:"
         Height          =   195
         Left            =   150
         TabIndex        =   55
         Top             =   375
         Width           =   675
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   " Cheque "
      Height          =   780
      Left            =   120
      TabIndex        =   50
      Top             =   2595
      Width           =   8265
      Begin VB.TextBox txt_fecha_cheque 
         Height          =   345
         Left            =   6270
         TabIndex        =   12
         Top             =   300
         Width           =   1410
      End
      Begin VB.TextBox txt_cheque 
         Height          =   345
         Left            =   1410
         MaxLength       =   4
         TabIndex        =   9
         Top             =   315
         Width           =   990
      End
      Begin VB.TextBox txt_nombre_banco_cheque 
         Height          =   345
         Left            =   3690
         TabIndex        =   11
         Top             =   315
         Width           =   1755
      End
      Begin VB.TextBox txt_banco_cheque 
         Height          =   345
         Left            =   3135
         TabIndex        =   10
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Cheque:"
         Height          =   195
         Left            =   720
         TabIndex        =   53
         Top             =   375
         Width           =   600
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   5745
         TabIndex        =   52
         Top             =   375
         Width           =   495
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   2595
         TabIndex        =   51
         Top             =   375
         Width           =   510
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmrelacion_cobranza_captura.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos Relación "
      Height          =   1185
      Left            =   120
      TabIndex        =   42
      Top             =   435
      Width           =   8265
      Begin VB.TextBox txt_relacion 
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
         MaxLength       =   10
         TabIndex        =   0
         Top             =   270
         Width           =   1635
      End
      Begin VB.TextBox txt_agente 
         Height          =   330
         Left            =   1425
         TabIndex        =   2
         Top             =   720
         Width           =   1245
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   330
         Left            =   2685
         TabIndex        =   3
         Top             =   720
         Width           =   4935
      End
      Begin VB.TextBox txt_fecha 
         Height          =   345
         Left            =   3810
         TabIndex        =   1
         Top             =   315
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Relación:"
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Top             =   345
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Top             =   735
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   3225
         TabIndex        =   43
         Top             =   375
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Documentos a aplicar "
      Height          =   780
      Left            =   120
      TabIndex        =   38
      Top             =   3510
      Width           =   8265
      Begin VB.TextBox txt_documento 
         Height          =   315
         Left            =   1935
         MaxLength       =   5
         TabIndex        =   13
         Top             =   300
         Width           =   1020
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   5160
         TabIndex        =   15
         Top             =   300
         Width           =   1320
      End
      Begin VB.TextBox txt_serie 
         Height          =   315
         Left            =   3690
         TabIndex        =   14
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
         Height          =   195
         Left            =   885
         TabIndex        =   41
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   4425
         TabIndex        =   40
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   3150
         TabIndex        =   39
         Top             =   360
         Width           =   405
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Datos del Documento "
      Height          =   1455
      Left            =   135
      TabIndex        =   32
      Top             =   4365
      Width           =   8250
      Begin VB.TextBox txt_agente_factura 
         Enabled         =   0   'False
         Height          =   315
         Left            =   930
         TabIndex        =   16
         Top             =   300
         Width           =   1620
      End
      Begin VB.TextBox txt_nombre_agente_factura 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2580
         TabIndex        =   17
         Top             =   300
         Width           =   5280
      End
      Begin VB.TextBox txt_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   930
         TabIndex        =   18
         Top             =   645
         Width           =   1620
      End
      Begin VB.TextBox txt_nombre_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2580
         TabIndex        =   19
         Top             =   645
         Width           =   5280
      End
      Begin VB.TextBox txt_importe 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3420
         TabIndex        =   21
         Top             =   975
         Width           =   1605
      End
      Begin VB.TextBox txt_fecha_factura 
         Enabled         =   0   'False
         Height          =   315
         Left            =   930
         TabIndex        =   20
         Top             =   990
         Width           =   1620
      End
      Begin VB.TextBox txt_saldo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6150
         TabIndex        =   22
         Top             =   975
         Width           =   1710
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   375
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   135
         TabIndex        =   36
         Top             =   705
         Width           =   525
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   2835
         TabIndex        =   35
         Top             =   1005
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   150
         TabIndex        =   34
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
         Height          =   195
         Left            =   5610
         TabIndex        =   33
         Top             =   1065
         Width           =   450
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Importe a aplicar "
      Height          =   720
      Left            =   150
      TabIndex        =   28
      Top             =   5865
      Width           =   8265
      Begin VB.TextBox txt_importe_aplicar 
         Height          =   315
         Left            =   2070
         TabIndex        =   23
         Top             =   270
         Width           =   1635
      End
      Begin VB.TextBox txt_descuento 
         Height          =   315
         Left            =   5835
         MaxLength       =   1
         TabIndex        =   24
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   1380
         TabIndex        =   31
         Top             =   330
         Width           =   570
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Descuento:"
         Height          =   195
         Left            =   4725
         TabIndex        =   30
         Top             =   330
         Width           =   825
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   6600
         TabIndex        =   29
         Top             =   330
         Width           =   120
      End
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmrelacion_cobranza_captura.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   795
      Picture         =   "frmrelacion_cobranza_captura.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Cancelar Alt + C"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame5 
      Height          =   75
      Left            =   90
      TabIndex        =   27
      Top             =   300
      Width           =   8325
   End
End
Attribute VB_Name = "frmrelacion_cobranza_captura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_lista As Integer
Private Sub cmd_aceptar_pedidos_Click()
   If Trim(txt_agente) <> "" Then
       If IsDate(Me.txt_fecha) Then
          If Trim(txt_documento) <> "" Then
             If Trim(txt_numero) <> "" Then
                If Trim(txt_cliente) <> "" Then
                   If IsNumeric(Me.txt_importe_aplicar) Then
                      If Trim(Me.txt_descuento) = "" Then
                         Me.txt_descuento = "0"
                      End If
                      If IsNumeric(Me.txt_numero_deposito) Then
                         If IsNumeric(txt_descuento) Then
                            If Trim(txt_deposito) <> "" Then
                               If Trim(Me.txt_banco) <> "" Then
                                  If IsDate(Me.txt_fecha_deposito) Then
                                     If Trim(txt_cheque) <> "" Then
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
                                                       Cadena = "EXECUTE RELACION_COBRANZA_DEPOSITO '" + var_empresa + "', '" + var_unidad_organizacional + "', '" + Me.txt_relacion + "', '" + Me.txt_fecha + "', '" + Me.txt_agente + "', '" + Me.txt_cliente + "', '" + txt_cheque + "', '"
                                                       Cadena = Cadena + Me.txt_fecha_cheque + "', " + Me.txt_importe_aplicar + ", " + Me.txt_descuento + ", " + txt_numero + ", 0, 0, " + CStr(var_partida) + ", 0, '" + txt_serie + "', '" + txt_documento + "', '" + txt_banco_cheque + "', '" + txt_deposito + "', '" + Me.txt_fecha_deposito + "', '" + Me.txt_banco + "'," + Me.txt_numero_deposito
                                                       rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                                       rs.Open "update tb_relacion_cobranza set dtim_rco_fecha_insercion = getdate() where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + txt_relacion + "' and inte_rco_partida = " + CStr(var_partida), cnn, adOpenDynamic, adLockOptimistic
                                                       MsgBox "Se a cargado la relación correctamente", vbOKOnly, "ATENCION"
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
       Else
          MsgBox "Fecha de relación de cobranza incorrecta", vbOKOnly, "ATENCION"
       End If
    Else
       MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
    End If
End Sub

Private Sub cmd_cancelar_pedidos_Click()
   Unload Me
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_agente = ""
   Me.txt_agente_factura = ""
   Me.txt_cliente = ""
   Me.txt_descuento = ""
   Me.txt_documento = ""
   Me.txt_fecha = Date
   Me.txt_fecha_factura = Date
   Me.txt_importe = ""
   Me.txt_importe_aplicar = ""
   Me.txt_nombre_agente = ""
   Me.txt_nombre_agente_factura = ""
   Me.txt_nombre_cliente = ""
   Me.txt_numero = ""
   Me.txt_relacion = ""
   Me.txt_saldo = ""
   Me.txt_serie = ""
   Me.txt_banco = ""
   Me.txt_deposito = ""
   Me.txt_numero_deposito = ""
   Me.txt_fecha_deposito = ""
   Me.txt_nombre_banco = ""
   Me.txt_fecha_deposito = Date
   Me.txt_relacion.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Top = 500
   Left = 1500
   Me.txt_relacion = ""
   frm_lista.Visible = False
   Me.txt_fecha = Date
   Me.txt_fecha_cheque = Date
   Me.txt_fecha_deposito = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_reporte_valuacion_devoluciones)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
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
   frm_lista.Visible = False
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_AGENTES where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_AGE_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 3
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

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agente_LostFocus()
   If Trim(txt_agente) <> "" Then
      rs.Open "SELECT * FROM TB_aGENTES WHERE VCHA_AGE_AGENTE_ID = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      Else
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
         txt_agente = ""
         txt_nombre_agente = ""
      End If
      rs.Close
   Else
      txt_nombre_agente = ""
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
                     list_item.SubItems(1) = Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
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

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         frmcalendario.mes.Value = CDate(Me.txt_fecha)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fecha = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_importe_aplicar_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_AGENTES order by vcha_AGE_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 3
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

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
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
      rs.Open "select * from tb_bancos where vcha_ban_banco_id = '20' or vcha_ban_banco_id = '11' or vcha_ban_banco_id = '10'  or vcha_ban_banco_id = '22' order by vcha_ban_nombre", cnn, adOpenDynamic, adLockOptimistic
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
   If Me.txt_documento = "SALDO" Then
   Else
      If Trim(txt_documento) <> "" Then
         If IsNumeric(txt_numero) Then
            rs.Open "SELECT * FROM TB_ENCABEZADO_cARTERA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_cAR_DOCUMENTO = '" + txt_documento + "' AND VCHA_sER_SERIE_ID = '" + txt_serie + "' AND INTE_cAR_NUMERO = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               txt_importe = IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
               txt_agente_factura = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
               txt_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
               If rsaux4.State = 1 Then
                  rsaux4.Close
               End If
               rsaux4.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + txt_agente_factura + "'", cnn, adOpenDynamic, adLockOptimistic
               txt_nombre_agente_factura = IIf(IsNull(rsaux4!VCHA_AGE_NOMBRE), "", rsaux4!VCHA_AGE_NOMBRE)
               rsaux4.Close
               rsaux4.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
               txt_nombre_cliente = IIf(IsNull(rsaux4!VCHA_CLI_NOMBRE), "", rsaux4!VCHA_CLI_NOMBRE)
               rsaux4.Close
               rsaux4.Open "SELECT * FROM TB_SALDOS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_cAR_DOCUMENTO = '" + txt_documento + "' AND VCHA_sER_SERIE_ID = '" + txt_serie + "' AND INTE_cAR_NUMERO = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
               txt_saldo = IIf(IsNull(rsaux4!FLOA_sAL_IMPORTE), 0, rsaux4!FLOA_sAL_IMPORTE)
               rsaux4.Close
               txt_fecha_factura = IIf(IsNull(rs!dtim_Car_fecha), "", rs!dtim_Car_fecha)
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
   End If
End Sub

Private Sub txt_relacion_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_relacion_LostFocus()
    If Trim(txt_relacion) <> "" Then
       rs.Open "select * from tb_relacion_cobranza with (nolock) where vcha_Rco_folio = '" + txt_relacion + "'", cnn, adOpenDynamic, adLockOptimistic
       If Not rs.EOF Then
          MsgBox "La relación de cobranza ya existe", vbOKOnly, "ATENCION"
          txt_relacion = ""
       End If
       rs.Close
    End If
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

