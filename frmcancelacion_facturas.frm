VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmcancelacion_facturas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelación y Reimpresión de Facturas"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11205
      Picture         =   "frmcancelacion_facturas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Salir"
      Top             =   60
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmcancelacion_facturas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Reimprimir Factura Alt + I"
      Top             =   60
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmcancelacion_facturas.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Cancelar Alt + C"
      Top             =   60
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Factura "
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   11460
      Begin VB.TextBox txt_estatus 
         Height          =   315
         Left            =   6675
         TabIndex        =   14
         Top             =   285
         Width           =   1980
      End
      Begin VB.TextBox txt_numero_factura 
         Height          =   315
         Left            =   885
         TabIndex        =   8
         Top             =   285
         Width           =   1215
      End
      Begin VB.TextBox txt_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   885
         TabIndex        =   7
         Top             =   615
         Width           =   4860
      End
      Begin VB.TextBox txt_fecha 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2760
         TabIndex        =   6
         Top             =   285
         Width           =   1500
      End
      Begin VB.TextBox txt_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6675
         TabIndex        =   5
         Top             =   615
         Width           =   4710
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estatus:"
         Height          =   195
         Left            =   6045
         TabIndex        =   13
         Top             =   345
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   345
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   195
         TabIndex        =   11
         Top             =   675
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   2235
         TabIndex        =   10
         Top             =   345
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   6090
         TabIndex        =   9
         Top             =   675
         Width           =   525
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalle de Factura"
      Height          =   5565
      Left            =   120
      TabIndex        =   1
      Top             =   1635
      Width           =   11475
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   10185
         TabIndex        =   2
         Top             =   6015
         Width           =   1350
      End
      Begin MSComctlLib.ListView lv_detalle_factura 
         Height          =   5265
         Left            =   45
         TabIndex        =   3
         Top             =   210
         Width           =   11370
         _ExtentX        =   20055
         _ExtentY        =   9287
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2364
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad  "
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Precio       "
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Importe     "
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Descuento      "
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "IVA       "
            Object.Width           =   1946
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Total         "
            Object.Width           =   2205
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   120
      TabIndex        =   0
      Top             =   330
      Width           =   11610
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   3060
      Top             =   60
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
            Picture         =   "frmcancelacion_facturas.frx":083E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcancelacion_facturas.frx":1118
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   2295
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcancelacion_facturas.frx":19F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcancelacion_facturas.frx":22CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcancelacion_facturas.frx":2BA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcancelacion_facturas.frx":3142
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcancelacion_facturas.frx":3A1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcancelacion_facturas.frx":42F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcancelacion_facturas.frx":4BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcancelacion_facturas.frx":4CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcancelacion_facturas.frx":4DF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcancelacion_facturas.frx":4F08
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcancelacion_facturas.frx":501A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcancelacion_facturas.frx":512C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcancelacion_facturas.frx":52AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcancelacion_facturas.frx":53C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3975
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcancelacion_facturas.frx":54D2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmcancelacion_facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim var_estatus_factura As String

Private Sub cmd_imprimir_Click()
   Dim var_numero_factura As Integer
        rs.Open "select * from tb_encabezado_facturas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_fac_numero = " + txt_numero_factura, cnn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
           var_estatus_factura = rs!char_fac_estatus
           If var_estatus_factura = "C" Then
               Set TB_ENC_FACTURAS_I = New TB_ENC_FACTURAS_I
               Set TB_DET_FACTURAS_I = New TB_DET_FACTURAS_I
               Set TB_INCREMENTA_FACTURA = New TB_INCREMENTA_FACTURA
               rsaux2.Open "select * from vw_maximo_factura WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  var_numero_factura = rsaux2!MAXIMO_FACTURA
               Else
                  var_numero_factura = 1
               End If
               rsaux2.Close
               si = MsgBox("Se va a imprimir la factura " + txt_numero_factura + " en la factura " + Str(var_numero_factura), vbYesNo, "ATENCION")
               If si = 6 Then
                  var_modifica = False
                  var_modifica = TB_INCREMENTA_FACTURA.Anadir(var_empresa, 1)
                  var_inserta = False
                  var_inserta = TB_ENC_FACTURAS_I.Anadir(rs(0).Value, rs(1).Value, rs(2).Value, rs(3).Value, var_numero_factura, rs(5).Value, rs(6).Value, rs(7).Value, rs(8).Value, rs(9).Value, rs(10).Value, rs(11).Value, rs(12).Value, rs(13).Value, rs(14).Value, rs(15).Value, rs(16).Value, rs(17).Value, rs(18).Value, rs(19).Value, rs(20).Value, rs(21).Value, rs(22).Value, rs(23).Value, rs(24).Value, "I", rs(26).Value)
                  rsaux2.Open "select * from tb_detalle_facturas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_fac_numero = " + Str(rs!inte_fac_numero), cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     While Not rsaux2.EOF
                        var_inserta = False
                        var_inserta = TB_DET_FACTURAS_I.Anadir(rsaux2(0).Value, rsaux2(1).Value, rsaux2(2).Value, var_numero_factura, rsaux2(4).Value, rsaux2(5).Value, rsaux2(6).Value, rsaux2(7).Value, rsaux2(8).Value)
                        rsaux2.MoveNext
                     Wend
                  End If
                  rsaux2.Close
               Else
                  MsgBox "La reimpresión de la factura a sido cancelada", vbOKOnly, "ATENCION"
               End If
            End If
            If var_estatus_factura = "I" Then
            End If
        
         End If
         rs.Close
End Sub

Private Sub cmd_cancelar_Click()
   Dim var_numero_factura As Integer
         If Trim(txt_numero_factura) <> "" Then
            rs.Open "select * from tb_encabezado_facturas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_fac_numero = " + txt_numero_factura, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If Trim(rs!char_fac_estatus) = "C" Then
                  MsgBox "La factura ya habia sido cancelada", vbOKOnly, "ATENCION"
               Else
                  si = MsgBox("¿Deseas cancelar la factura " + Trim(txt_numero_factura) + "?", vbYesNo, "ATENCION")
                  If si = 6 Then
                     si = MsgBox("Confirmar cancelar factura", vbYesNo, "ATENCION")
                     If si = 6 Then
                        Set TB_ENC_FACTURA_M = New TB_ENC_FACTURA_M
                        var_modifica = False
                        var_modifica = TB_ENC_FACTURA_M.Anadir(var_empresa, var_unidad_organizacional, txt_numero_factura, "C")
                        txt_estatus = "CANCELADA"
                        var_estatus_factura = "C"
                     End If
                  End If
               End If
            End If
            rs.Close
         End If
End Sub


Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 67 Then
      cmd_cancelar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If var_activa_menu = True Then
       Frmmenu2.Enabled = True
    End If
End Sub

Private Sub txt_numero_factura_KeyPress(KeyAscii As Integer)
   Dim list_item As ListItem
   Dim var_descuento_1 As Double
   Dim var_descuento_2 As Double
   Dim var_descuento_3 As Double
   Dim var_imp_desc_1 As Double
   Dim var_imp_desc_2 As Double
   Dim var_imp_desc_3 As Double
   Dim var_precio As Double
   Dim var_cantidad As Double
   Dim var_iva As Double
   Dim var_imp_iva As Double
   Dim var_imp_desc As Double
   Dim var_importe As Double
   If KeyAscii = 13 Then
      lv_detalle_factura.ListItems.Clear
      txt_estatus = ""
      txt_agente = ""
      txt_cliente = ""
      txt_fecha = ""
      If Trim(txt_numero_factura) <> "" Then
         rs.Open "select * from tb_encabezado_facturas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_fac_numero = " + txt_numero_factura, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_estatus_factura = rs!char_fac_estatus
            If Trim(rs!char_fac_estatus) = "I" Then
               txt_estatus = "IMPRESA"
            End If
            If Trim(rs!char_fac_estatus) = "C" Then
               txt_estatus = "CANCELADA"
            End If
            rsaux2.Open "select * from vw_detalle_facturas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_fac_numero = " + Str(rs!inte_fac_numero), cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               txt_agente = rsaux2!vcha_age_nombre
               txt_fecha = rsaux2!dtim_fac_fecha
               txt_cliente = rsaux2!vcha_cli_nombre
               txt_importe = Format(rsaux2!floa_fac_total, "###,###,##0.00")
               While Not rsaux2.EOF
                  var_descuento_1 = 0
                  var_descuento_2 = 0
                  var_descuento_3 = 0
                  var_imp_desc_1 = 0
                  var_imp_desc_2 = 0
                  var_imp_desc_3 = 0
                  var_precio = 0
                  var_iva = 0
                  var_imp_iva = 0
                  Set list_item = lv_detalle_factura.ListItems.Add(, , rsaux2!vcha_art_articulo_id)
                  list_item.SubItems(1) = IIf(IsNull(rsaux2!vcha_Art_nombre_español), "", Trim(rsaux2!vcha_Art_nombre_español))
                  list_item.SubItems(2) = Format(IIf(IsNull(rsaux2!floa_fac_cantidad), 0, rsaux2!floa_fac_cantidad), "###,###,##0.00")
                  list_item.SubItems(3) = Format(IIf(IsNull(rsaux2!floa_fac_precio), 0, rsaux2!floa_fac_precio), "###,###,##0.00")
                  var_cantidad = Format(IIf(IsNull(rsaux2!floa_fac_cantidad), 0, rsaux2!floa_fac_cantidad), "###,###,##0.00")
                  var_precio = Format(IIf(IsNull(rsaux2!floa_fac_precio), 0, rsaux2!floa_fac_precio), "###,###,##0.00")
                  var_descuento_1 = IIf(IsNull(rsaux2!floa_fac_descuento_1), 0, rsaux2!floa_fac_descuento_1)
                  var_descuento_2 = IIf(IsNull(rsaux2!floa_fac_descuento_2), 0, rsaux2!floa_fac_descuento_2)
                  var_descuento_3 = IIf(IsNull(rsaux2!floa_fac_descuento_3), 0, rsaux2!floa_fac_descuento_3)
                  list_item.SubItems(4) = Format(var_cantidad * var_precio, "###,###,##0.00")
                  If var_descuento_1 > 0 Then
                     var_imp_desc_1 = var_precio * (var_descuento_1 / 100)
                  End If
                  If var_descuento_2 > 0 Then
                     var_imp_desc_2 = (var_precio - var_imp_desc_1) * (var_descuento_2 / 100)
                  End If
                  list_item.SubItems(5) = Format((var_imp_desc_1 + var_imp_desc_2) * var_cantidad, "###,###,##0.00")
                  var_importe = var_precio - var_imp_desc_1 - var_imp_desc_2
                  var_iva = rsaux2!floa_Fac_iva
                  var_imp_iva = var_importe * (var_iva / 100)
                  list_item.SubItems(6) = Format(var_imp_iva, "###,###,##0.00")
                  list_item.SubItems(7) = Format(var_imp_iva + var_importe, "###,###,##0.00")
                  rsaux2.MoveNext
               Wend
            Else
               MsgBox "El detalle de la factura no existe", vbOKOnly, "ATENCION"
            End If
            rsaux2.Close
         Else
            MsgBox "La factura no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Debe de seleccionar una factura", vbOKOnly, "ATENCION"
      End If
   End If
End Sub
