VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcheques_devueltos_inicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura de Cheques Devuelto"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   8145
   Begin VB.Frame frm_lista 
      Height          =   2565
      Left            =   1215
      TabIndex        =   25
      Top             =   105
      Width           =   4050
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2115
         Left            =   45
         TabIndex        =   26
         Top             =   405
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   3731
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
            Object.Width           =   1464
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5380
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   27
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos Generales "
      Height          =   2070
      Left            =   90
      TabIndex        =   13
      Top             =   435
      Width           =   7920
      Begin VB.TextBox txt_clase_cartera 
         Height          =   315
         Left            =   885
         TabIndex        =   4
         Top             =   285
         Width           =   1725
      End
      Begin VB.TextBox txt_nombre_clase_cartera 
         Height          =   315
         Left            =   2610
         TabIndex        =   5
         Top             =   285
         Width           =   5175
      End
      Begin VB.ComboBox cmb_series 
         Height          =   315
         Left            =   7005
         TabIndex        =   22
         Top             =   1605
         Width           =   795
      End
      Begin VB.TextBox txt_comision 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   885
         TabIndex        =   12
         Top             =   1605
         Width           =   1725
      End
      Begin VB.TextBox txt_nombre_banco 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2625
         TabIndex        =   9
         Top             =   945
         Width           =   5175
      End
      Begin VB.TextBox txt_nombre_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2625
         TabIndex        =   7
         Top             =   615
         Width           =   5175
      End
      Begin VB.TextBox txt_clave_banco 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   885
         TabIndex        =   8
         Top             =   945
         Width           =   1725
      End
      Begin VB.TextBox txt_clave_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   885
         TabIndex        =   6
         Top             =   615
         Width           =   1725
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   3345
         TabIndex        =   11
         Top             =   1275
         Width           =   1725
      End
      Begin VB.TextBox txt_cheque 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   885
         TabIndex        =   10
         Top             =   1275
         Width           =   1725
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   135
         TabIndex        =   24
         Top             =   345
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   6525
         TabIndex        =   23
         Top             =   1665
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Comisión:"
         Height          =   195
         Left            =   135
         TabIndex        =   21
         Top             =   1665
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   135
         TabIndex        =   19
         Top             =   1005
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   135
         TabIndex        =   17
         Top             =   675
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto:"
         Height          =   195
         Left            =   2790
         TabIndex        =   16
         Top             =   1335
         Width           =   495
      End
      Begin VB.Label lbl_moneda 
         Height          =   285
         Left            =   5130
         TabIndex        =   15
         Top             =   915
         Width           =   2625
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cheque:"
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   1335
         Width           =   600
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7650
      Picture         =   "frmcheques_devueltos_inicio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmcheques_devueltos_inicio.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Nota de Cargo por Comisión "
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_actualizar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmcheques_devueltos_inicio.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Actualizar"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   75
      TabIndex        =   18
      Top             =   270
      Width           =   7920
   End
   Begin VB.Frame Frame3 
      Caption         =   " Cheques "
      Height          =   3765
      Left            =   105
      TabIndex        =   20
      Top             =   2535
      Width           =   7905
      Begin MSComctlLib.ListView lv_cheques 
         Height          =   3495
         Left            =   75
         TabIndex        =   3
         Top             =   195
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   6165
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
         NumItems        =   19
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Folio"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Banco"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cheque"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Fecha"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe   "
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Clave cliente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Clave banco"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Estatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Moneda"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Moneda local"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Moneda Plurar"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "comision"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "iva"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "grupo_actual"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "grupo_real"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "titular"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "agente"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmcheques_devueltos_inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_ruta As String
Dim var_tabla As ADODB.Connection
Dim var_serie As String




Private Sub cmd_guardar_Click()
End Sub


Private Sub cmb_series_Click()
   var_serie = cmb_series
End Sub

Private Sub cmd_actualizar_Click()
   txt_clave_cliente = ""
   cmb_clientes = ""
   txt_clave_banco = ""
   cmb_bancos = ""
   txt_importe = ""
   txt_cheque = ""
   Me.txt_nombre_banco = ""
   Me.txt_nombre_clase_cartera = "CHEQUE DEVUELTO"
   Me.txt_nombre_cliente = ""
   lv_cheques.ListItems.Clear
   rs.Open "select * from vw_cheques where vcha_Emp_empresa_id = '" + var_empresa + "' and char_che_estatus = 'A'", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
      Set list_item = lv_cheques.ListItems.Add(, , rs!vcha_rco_folio)
      list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre))
      list_item.SubItems(2) = Trim(IIf(IsNull(rs!vcha_ban_nombre), "", rs!vcha_ban_nombre))
      list_item.SubItems(3) = IIf(IsNull(rs!vcha_che_cheque), "", rs!vcha_che_cheque)
      list_item.SubItems(4) = IIf(IsNull(rs!dtim_che_fecha), "", Format(rs!dtim_che_fecha, "Short Date"))
      list_item.SubItems(5) = Format(IIf(IsNull(rs!floa_che_importe), 0, rs!floa_che_importe), "###,###,##0.00")
      list_item.SubItems(6) = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
      list_item.SubItems(7) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
      list_item.SubItems(8) = IIf(IsNull(rs!char_che_estatus), "", rs!char_che_estatus)
      list_item.SubItems(10) = IIf(IsNull(rs!VCHA_MON_MONEDA_ID), "", rs!VCHA_MON_MONEDA_ID)
      list_item.SubItems(11) = IIf(IsNull(rs!inte_mon_moneda_local), 1, rs!inte_mon_moneda_local)
      list_item.SubItems(12) = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
      list_item.SubItems(13) = Format(IIf(IsNull(rs!FLOA_CHE_COMISION_BOTADO), 0, rs!FLOA_CHE_COMISION_BOTADO), "###,###,##0.00")
      list_item.SubItems(14) = IIf(IsNull(rs!FLOA_TPE_IVA), 0, rs!FLOA_TPE_IVA)
      list_item.SubItems(15) = IIf(IsNull(rs!vcha_gac_grupo_Actual_id), "", rs!vcha_gac_grupo_Actual_id)
      list_item.SubItems(16) = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
      list_item.SubItems(17) = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
      list_item.SubItems(18) = IIf(IsNull(rs!vcha_age_agente_id), "", rs!vcha_age_agente_id)
      rs.MoveNext
   Wend
   rs.Close
   n = lv_cheques.ListItems.Count
   For i = 1 To n
      lv_cheques.ListItems.Item(i).Selected = True
      If lv_cheques.selectedItem.SubItems(8) = "B" Then
         lv_cheques.selectedItem.ForeColor = &HFF&
         lv_cheques.selectedItem.ListSubItems.Item(1).ForeColor = &HFF&
         lv_cheques.selectedItem.ListSubItems.Item(2).ForeColor = &HFF&
         lv_cheques.selectedItem.ListSubItems.Item(3).ForeColor = &HFF&
         lv_cheques.selectedItem.ListSubItems.Item(4).ForeColor = &HFF&
         lv_cheques.selectedItem.ListSubItems.Item(5).ForeColor = &HFF&
         lv_cheques.selectedItem.ListSubItems.Item(6).ForeColor = &HFF&
         lv_cheques.selectedItem.ListSubItems.Item(7).ForeColor = &HFF&
      End If
   Next i
   Me.txt_clase_cartera.Enabled = False
   Me.txt_nombre_clase_cartera.Enabled = False
End Sub

Private Sub cmd_imprimir_Click()
   Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
   Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
   Dim si As Integer
   Dim var_folio As String
   Dim var_agente As String
   Dim var_grupo_actual As String
   Dim var_grupo_real As String
   Dim var_titular As String
   Dim var_importe_iva As Double
   Dim var_importe_total As Double
   Dim var_importe_neto As Double
   Dim var_subimporte As Double
   Dim var_tipo_Cambio As Double
   Dim var_numero_folio As Double
   Dim var_moneda_local As Integer
   Dim var_posible_tipo_cambio As Boolean
   Dim var_clave_moneda As String
   Dim var_iva As Double
   var_moneda_local = 0
   If Not Trim(txt_comision) = "" Or Not txt_comision = 0 Then
      var_clave_moneda = lv_cheques.selectedItem.SubItems(10)
      rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_moneda_local = IIf(IsNull(rs!inte_mon_moneda_local), 0, rs!inte_mon_moneda_local)
      End If
      rs.Close
      var_tipo_Cambio = 1
      If var_moneda_local = 0 Then
         rs.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_tipo_Cambio = IIf(IsNull(rs!mone_tca_importe), 1, rs!mone_tca_importe)
            var_posible_tipo_cambio = True
         Else
            var_posible_tipo_cambio = False
         End If
         rs.Close
      Else
         var_posible_tipo_cambio = True
      End If
      
      If var_posible_tipo_cambio = True Then
         si = MsgBox("¿Deseas imprimir la Nota de Cargo?", vbYesNo, "ATENCION")
         If si = 6 Then
            si = MsgBox("Confirmar la impresíon de la Nota de Cargo", vbYesNo, "ATENCION")
            If si = 6 Then
               rs.Open "select inte_ser_nota_cargo from tb_series where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  rs.Close
                  rsaux2.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_numero_folio = rsaux2!INTE_SER_NOTA_CARGO
                  rsaux2.Close
                  MsgBox "Se va a imprimir la Nota de cargo Número " + Str(var_numero_folio), vbYesNo, "ATENCION"
                  si = MsgBox("¿La impresora esta lista?", vbYesNo, "ATENCION")
                  If si = 6 Then
                     cnn.BeginTrans
                     var_iva = lv_cheques.selectedItem.SubItems(14) * 1
                     var_grupo_actual = lv_cheques.selectedItem.SubItems(15)
                     var_grupo_real = lv_cheques.selectedItem.SubItems(16)
                     var_titular = lv_cheques.selectedItem.SubItems(17)
                     var_agente = lv_cheques.selectedItem.SubItems(18)
                     var_importe_total = (txt_comision / (1 + (var_iva / 100))) * var_tipo_Cambio
                     var_importe_neto = txt_comision * var_tipo_Cambio
                     var_importe_iva = var_importe_neto - var_importe_total
                     var_subimporte = var_importe_total
                     var_insertar = False
                     var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "NG", "NC", txt_clase_cartera, var_numero_folio, "+", "", "", 0, CStr(Date), var_agente, var_grupo_actual, var_grupo_real, var_titular, txt_clave_cliente, "", 0, var_iva, 0, 0, 0, 0, 0, var_importe_total, var_importe_iva, 0, 0, 0, 0, 0, var_subimporte, var_importe_neto, "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, var_clave_moneda, var_tipo_Cambio, var_serie, "")
                     var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, var_serie, "NC", var_numero_folio, "", "", 0, var_importe_neto, 0)
                     rsaux3.Open "update tb_series set inte_ser_nota_Cargo = inte_ser_nota_cargo + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_folio = lv_cheques.selectedItem
                     rsaux3.Open "update tb_cheques set CHAR_CHE_ESTATUS = 'C' where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_RCO_FOLIO = '" + var_folio + "' and vcha_cli_clave_id = '" + txt_clave_cliente + "' and vcha_ban_banco_id = '" + txt_clave_banco + "' and VCHA_CHE_CHEQUE = '" + txt_cheque + "'", cnn, adOpenDynamic, adLockOptimistic
                     cnn.CommitTrans
'''''''
                     If rs.State = 1 Then
                        rs.Close
                     End If
                     rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'NC' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
'''''''''''''''  IMPRESION DE LA NOTA DE CARGO
                       Open (App.Path & "\nota_cargo" + Trim(Str(rs!inte_car_numero)) + ".txt") For Output As #1
                       Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                       Print #1, ""
                       Print #1, ""
                       Print #1, ""
                       Print #1, ""
                       Print #1, Spc(92); Str(rs!inte_car_numero)
                       Print #1, ""
                       Print #1, ""
                       Print #1, Spc(93); "FECHA: "; Format(rs!DTIM_CAR_FECHA, "Short Date")
                       Print #1, ""
                       var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
                       For var_j = 1 + Len(Trim(var_cliente)) To 83
                           var_cliente = var_cliente + " "
                       Next var_j
                       var_cliente = var_cliente + "AGUASCALIENTES, AGS."
                       Print #1, ""
                       Print #1, Spc(10); var_cliente
                       var_domicilio = IIf(IsNull(rs!vcha_cli_direccion), "", rs!vcha_cli_direccion) + " C.P. " + IIf(IsNull(rs!vcha_cli_cp), "", rs!vcha_cli_cp)
                       For var_j = 1 + Len(Trim(var_domicilio)) To 83
                           var_domicilio = var_domicilio + " "
                       Next var_j
                       var_agente = ""
                       var_agente = IIf(IsNull(rs!vcha_age_agente_id), "", rs!vcha_age_agente_id)
                       For var_j = 1 + Len(Trim(var_agente)) To 8
                           var_agente = var_agente + " "
                       Next var_j
                       var_agente = var_agente + IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
                       var_domicilio = var_domicilio
                       Print #1, Spc(10); var_domicilio
                       var_ciudad = ""
                       var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                       For var_j = 1 + Len(Trim(var_ciudad)) To 37
                           var_ciudad = var_ciudad + " "
                       Next var_j
                       var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                       For var_j = 1 + Len(Trim(var_estado)) To 46
                           var_estado = var_estado + " "
                       Next var_j
                       var_ciudad = var_ciudad + var_estado
                              
                       For var_j = 1 + Len(Trim(var_ciudad)) To 14
                           var_ciudad = var_ciudad + " "
                       Next var_j
                           
                       var_ciudad = var_ciudad + var_agente
                            
                       Print #1, Spc(10); var_ciudad
                       var_rfc = IIf(IsNull(rs!vcha_cli_rfc), "", rs!vcha_cli_rfc)
                       var_rfc = "RFC:  " + var_rfc
                       For var_j = 1 + Len(Trim(var_rfc)) To 89
                           var_rfc = var_rfc + " "
                       Next var_j
                       var_rfc = var_rfc
                       Print #1, Spc(4); var_rfc
                       Print #1, ""
                       Print #1, ""
                       var_linea = "NC" + Str(rs!inte_car_numero) + " " + rs!vcha_car_nombre
                       If Len(Trim(var_linea)) < 108 Then
                          For var_j = 1 + Len(Trim(var_linea)) To 108
                              var_linea = var_linea + " "
                          Next var_j
                       End If
                       
                       var_importe_str = Format(((IIf(IsNull(rs!FLOA_CAR_IMPORTE_NETO), 0, rs!FLOA_CAR_IMPORTE_NETO)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                       If Len(Trim(var_importe_str)) < 14 Then
                          For var_j = 1 + Len(Trim(var_importe_str)) To 14
                              var_importe_str = " " + var_importe_str
                          Next var_j
                       End If
                       var_linea = var_linea + var_importe_str
                       Print #1, var_linea
                       Print #1, ""
                       Print #1, ""
                       Print #1, ""
                       Print #1, ""
                       Print #1, ""
                       Print #1, ""
                       Print #1, ""
                       Print #1, ""
                       Print #1, ""
                       var_cantidad_letra = rs!vcha_car_importe_letra
                       
                       var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                       If Len(Trim(var_linea)) < 93 Then
                          For var_j = 1 + Len(Trim(var_linea)) To 93
                              var_linea = var_linea + " "
                          Next var_j
                       End If
                       
                       var_rfc = IIf(IsNull(rs!vcha_cli_rfc), "", rs!vcha_cli_rfc)
                       
                       If Len(Trim(var_rfc)) = 0 Then
                          var_subimporte_str = Format((IIf(IsNull(rs!FLOA_CAR_IMPORTE_NETO), 0, rs!FLOA_CAR_IMPORTE_NETO)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                          If Len(Trim(var_subimporte_str)) < 14 Then
                             For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                 var_subimporte_str = " " + var_subimporte_str
                             Next var_j
                          End If
                          var_iva = "      -        "
                          For var_j = 1 + Len(Trim(var_iva_str)) To 14
                              var_iva_str = " " + var_iva_str
                          Next var_j
                       Else
                          var_subimporte_str = Format(((IIf(IsNull(rs!FLOA_CAR_IMPORTE_NETO), 0, rs!FLOA_CAR_IMPORTE_NETO)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                          If Len(Trim(var_subimporte_str)) < 14 Then
                             For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                 var_subimporte_str = " " + var_subimporte_str
                             Next var_j
                          End If
                          var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                          If Len(Trim(var_iva_str)) < 14 Then
                             For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                 var_iva_str = " " + var_iva_str
                             Next var_j
                          End If
                       End If
                       var_linea = var_linea + "           " + var_subimporte_str
                       Print #1, Spc(4); var_linea
                       Print #1, Spc(108); var_iva_str
                       var_importe_str = Format((IIf(IsNull(rs!FLOA_CAR_IMPORTE_NETO), 0, rs!FLOA_CAR_IMPORTE_NETO)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                       If Len(Trim(var_importe_str)) < 14 Then
                          For var_j = 1 + Len(Trim(var_importe_str)) To 14
                              var_importe_str = " " + var_importe_str
                          Next var_j
                       End If
                       Print #1, Spc(108); var_importe_str
                       Print #1, ""
                       Print #1, ""
                       Print #1, ""
                       Print #1, Spc(85); "SISTEMAS"
                       Close #1
                       
                       Open (App.Path & "\nota_cargo" + Trim(Str(rs!inte_car_numero)) + ".bat") For Output As #2
                       var_Archivo = App.Path & "\nota_cargo" + Trim(Str(rs!inte_car_numero)) + ".bat"
                       Print #2, "copy " + App.Path + "\nota_cargo" + Trim(Str(rs!inte_car_numero)) + ".txt lpt1"
                       Close #2
                       x = Shell(var_Archivo, vbHide)
'''''''''''''''
                     End If
                     rs.Close
'''''''
                     MsgBox "Se a terminado la Impresión de la Nota de Cargo", vbOKOnly, "ATENCION"
                  Else
                     MsgBox "La impresión de la Nota de Cargo a sido cancelada", vbOKOnly, "ATENCION"
                  End If
               Else
                  rs.Close
                  MsgBox "No se a definido la empresa", vbOKOnly, "ATENCION"
               End If
            End If
         End If
      Else
         MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No existe una comisión", vbOKOnly, "ATENCION"
   End If
End Sub


Private Sub cmd_nuevo_Click()
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   txt_clase_cartera.Enabled = False
   txt_nombre_clase_cartera.Enabled = False
   frm_lista.Visible = False
   rs.Open "select vcha_ser_serie_id from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_contador_serie = 0
      While Not rs.EOF
         var_contador_serie = var_contador_serie + 1
         rs.MoveNext
      Wend
      rs.MoveFirst
      cmd_actualizar.Enabled = True
      cmd_imprimir.Enabled = True
      Call RecsetToCombo(cmb_series.hwnd, rs, 0)
      If var_contador_serie > 1 Then
         cmb_series.Enabled = True
      Else
         cmb_series.Enabled = False
      End If
      rs.MoveFirst
      cmb_series = rs!vcha_ser_serie_id
      var_serie = rs!vcha_ser_serie_id
   Else
      MsgBox "No se a indicado una serie para esta Unidad organizacional", vbOKOnly, "ATENCION"
      cmd_actualizar.Enabled = False
      cmd_imprimir.Enabled = False
   End If
   rs.Close
   Dim i, n As Integer
   Top = 500
   Left = 1800
   Me.txt_clase_cartera = "CH"
   Me.txt_nombre_clase_cartera = "CHEQUE DEVUELTO"
   rs.Open "select * from vw_cheques where vcha_Emp_empresa_id = '" + var_empresa + "' and char_che_estatus = 'A'", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
      Set list_item = lv_cheques.ListItems.Add(, , rs!vcha_rco_folio)
      list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre))
      list_item.SubItems(2) = Trim(IIf(IsNull(rs!vcha_ban_nombre), "", rs!vcha_ban_nombre))
      list_item.SubItems(3) = IIf(IsNull(rs!vcha_che_cheque), "", rs!vcha_che_cheque)
      list_item.SubItems(4) = IIf(IsNull(rs!dtim_che_fecha), "", Format(rs!dtim_che_fecha, "Short Date"))
      list_item.SubItems(5) = Format(IIf(IsNull(rs!floa_che_importe), 0, rs!floa_che_importe), "###,###,##0.00")
      list_item.SubItems(6) = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
      list_item.SubItems(7) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
      list_item.SubItems(8) = IIf(IsNull(rs!char_che_estatus), "", rs!char_che_estatus)
      list_item.SubItems(10) = IIf(IsNull(rs!VCHA_MON_MONEDA_ID), "", rs!VCHA_MON_MONEDA_ID)
      list_item.SubItems(11) = IIf(IsNull(rs!inte_mon_moneda_local), 1, rs!inte_mon_moneda_local)
      list_item.SubItems(12) = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
      list_item.SubItems(13) = Format(IIf(IsNull(rs!FLOA_CHE_COMISION_BOTADO), 0, rs!FLOA_CHE_COMISION_BOTADO), "###,###,##0.00")
      list_item.SubItems(14) = IIf(IsNull(rs!FLOA_TPE_IVA), 0, rs!FLOA_TPE_IVA)
      list_item.SubItems(15) = IIf(IsNull(rs!vcha_gac_grupo_Actual_id), "", rs!vcha_gac_grupo_Actual_id)
      list_item.SubItems(16) = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
      list_item.SubItems(17) = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
      list_item.SubItems(18) = IIf(IsNull(rs!vcha_age_agente_id), "", rs!vcha_age_agente_id)
      rs.MoveNext
   Wend
   rs.Close
   n = lv_cheques.ListItems.Count
   For i = 1 To n
      lv_cheques.ListItems.Item(i).Selected = True
      If lv_cheques.selectedItem.SubItems(8) = "B" Then
         lv_cheques.selectedItem.ForeColor = &HFF&
         lv_cheques.selectedItem.ListSubItems.Item(1).ForeColor = &HFF&
         lv_cheques.selectedItem.ListSubItems.Item(2).ForeColor = &HFF&
         lv_cheques.selectedItem.ListSubItems.Item(3).ForeColor = &HFF&
         lv_cheques.selectedItem.ListSubItems.Item(4).ForeColor = &HFF&
         lv_cheques.selectedItem.ListSubItems.Item(5).ForeColor = &HFF&
         lv_cheques.selectedItem.ListSubItems.Item(6).ForeColor = &HFF&
         lv_cheques.selectedItem.ListSubItems.Item(7).ForeColor = &HFF&
      End If
   Next i
   Me.txt_clase_cartera.Enabled = False
   Me.txt_nombre_clase_cartera.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_cheques_devueltos_inicio)
End Sub

Private Sub lv_cheques_ItemClick(ByVal Item As MSComctlLib.ListItem)
   txt_clave_cliente = lv_cheques.selectedItem.SubItems(6)
   txt_nombre_cliente = lv_cheques.selectedItem.SubItems(1)
   txt_clave_banco = lv_cheques.selectedItem.SubItems(7)
   txt_nombre_banco = lv_cheques.selectedItem.SubItems(2)
   txt_cheque = lv_cheques.selectedItem.SubItems(3)
   txt_importe = Format(lv_cheques.selectedItem.SubItems(5), "###,###,##0.00")
   lbl_moneda = lv_cheques.selectedItem.SubItems(12)
   txt_comision = lv_cheques.selectedItem.SubItems(13)
   txt_clase_cartera.Enabled = False
   txt_nombre_clase_cartera.Enabled = False
End Sub

Private Sub lv_cheques_KeyPress(KeyAscii As Integer)
z = 0
If z = 1 Then
   If KeyAscii = 13 Then
      If lv_cheques.selectedItem.SubItems(8) = "A" Then
         If lv_cheques.selectedItem.SubItems(9) = "" Then
            lv_cheques.selectedItem.SubItems(9) = "*"
            lv_cheques.selectedItem.ForeColor = &HC000&
            lv_cheques.selectedItem.ListSubItems.Item(1).ForeColor = &HC000&
            lv_cheques.selectedItem.ListSubItems.Item(2).ForeColor = &HC000&
            lv_cheques.selectedItem.ListSubItems.Item(3).ForeColor = &HC000&
            lv_cheques.selectedItem.ListSubItems.Item(4).ForeColor = &HC000&
            lv_cheques.selectedItem.ListSubItems.Item(5).ForeColor = &HC000&
            lv_cheques.selectedItem.ListSubItems.Item(6).ForeColor = &HC000&
         Else
            lv_cheques.selectedItem.SubItems(9) = ""
            lv_cheques.selectedItem.ForeColor = &HFF&
            lv_cheques.selectedItem.ListSubItems.Item(1).ForeColor = &HFF&
            lv_cheques.selectedItem.ListSubItems.Item(2).ForeColor = &HFF&
            lv_cheques.selectedItem.ListSubItems.Item(3).ForeColor = &HFF&
            lv_cheques.selectedItem.ListSubItems.Item(4).ForeColor = &HFF&
            lv_cheques.selectedItem.ListSubItems.Item(5).ForeColor = &HFF&
            lv_cheques.selectedItem.ListSubItems.Item(6).ForeColor = &HFF&
            lv_cheques.selectedItem.ListSubItems.Item(7).ForeColor = &HFF&
         End If
      Else
         MsgBox "El cheque no puede ser cargado en cartera", vbOKOnly, "ATENCION"
      End If
   End If
   End If
End Sub


Private Sub Text2_Change()

End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_clase_cartera = lv_lista.selectedItem
         txt_nombre_clase_cartera = lv_lista.selectedItem.SubItems(1)
      Else
         txt_clase_cartera = ""
         txt_nombre_clase_cartera = ""
      End If
      frm_lista.Visible = False
      txt_clase_cartera.Enabled = True
      Me.txt_clase_cartera.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub txt_cheque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_clase_cartera_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clase_cartera_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      rs.Open "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where vcha_car_documento= 'NC' order by vcha_Car_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            var_contador_lista = var_contador_lista + 1
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_car_clase_id)
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre))
            rs.MoveNext
         Wend
      End If
      rs.Close
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clase_cartera_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clase_cartera_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_clave_banco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_clave_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_comision_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_banco_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_clase_cartera_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      rs.Open "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where vcha_car_documento= 'NC' order by vcha_Car_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            var_contador_lista = var_contador_lista + 1
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_car_clase_id)
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre))
            rs.MoveNext
         Wend
      End If
      rs.Close
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_clase_cartera_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub
