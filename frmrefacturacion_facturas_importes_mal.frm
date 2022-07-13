VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmrefacturacion_facturas_importes_mal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Refacturación por importes incorrectos"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   30
      TabIndex        =   12
      Top             =   375
      Width           =   9660
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos de la factura "
      Height          =   840
      Left            =   90
      TabIndex        =   5
      Top             =   630
      Width           =   9435
      Begin VB.TextBox txt_Estatus 
         Height          =   330
         Left            =   8205
         TabIndex        =   11
         Top             =   337
         Width           =   1140
      End
      Begin VB.TextBox txt_cliente 
         Height          =   330
         Left            =   2865
         TabIndex        =   9
         Top             =   337
         Width           =   4215
      End
      Begin VB.TextBox txt_factura 
         Height          =   330
         Left            =   840
         TabIndex        =   7
         Top             =   337
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Estatus:"
         Height          =   195
         Left            =   7530
         TabIndex        =   10
         Top             =   405
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   2280
         TabIndex        =   8
         Top             =   405
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   405
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4905
      Left            =   90
      TabIndex        =   0
      Top             =   1560
      Width           =   9435
      Begin VB.TextBox txt_correccion 
         Height          =   345
         Left            =   7815
         TabIndex        =   4
         Top             =   4410
         Width           =   1500
      End
      Begin VB.TextBox txt_cantidad 
         Height          =   345
         Left            =   6225
         TabIndex        =   3
         Top             =   4410
         Width           =   1500
      End
      Begin MSComctlLib.ListView lv_factura 
         Height          =   4275
         Left            =   45
         TabIndex        =   1
         Top             =   135
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   7541
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   10230
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Correccion"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   5640
         TabIndex        =   2
         Top             =   4500
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmrefacturacion_facturas_importes_mal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim list_item As ListItem

Private Sub Form_Load()
   Top = 300
   Left = 1100
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_packing_list)
End Sub

Private Sub txt_factura_Change()
   Me.txt_cliente = ""
   Me.txt_Estatus = ""
   Me.txt_correccion = ""
   Me.txt_cantidad = ""
   Me.lv_factura.ListItems.Clear
End Sub

Private Sub txt_factura_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_cliente.SetFocus
   End If
End Sub

Private Sub txt_factura_LostFocus()
   If Me.txt_factura <> "" Then
      var_numero = ""
      var_serie = ""
      For var_j = 1 To Len(Me.txt_factura)
          If IsNumeric(Mid(Me.txt_factura, var_j, 1)) Then
             var_numero = var_numero + Mid(Me.txt_factura, var_j, 1)
          Else
             var_serie = var_serie + Mid(Me.txt_factura, var_j, 1)
          End If
      Next var_j
      If var_numero <> "" Then
         rs.Open "select * from tb_encabezado_Cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "' and inte_car_numero = " + var_numero + " and vcha_Car_documento = 'FA'", cnn, adOpenDynamic, adLockBatchOptimistic
         If Not rs.EOF Then
            var_cadena = " SELECT     dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, dbo.TB_SALIDAS.FLOA_SAL_COSTO, dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, ISNULL(dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS, 'I') AS CHAR_CAR_ESTATUS FROM dbo.TB_ENCABEZADO_CARTERA INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_SALIDAS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO AND "
            var_cadena = var_cadena + "  dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = dbo.TB_SALIDAS.VCHA_SER_SERIE_ID AND           dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = dbo.TB_SALIDAS.INTE_CAR_NUMERO INNER JOIN  dbo.TB_ARTICULOS ON dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE     (dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = '" + var_serie + "') AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = " + var_numero + ") "
            rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_cliente = IIf(IsNull(rsaux!VCHA_CLI_NOMBRE), "", rsaux!VCHA_CLI_NOMBRE)
               Me.txt_Estatus = IIf(IsNull(rsaux!CHAR_CAR_ESTATUS), "I", rsaux!CHAR_CAR_ESTATUS)
               var_cantidad = 0
               While Not rsaux.EOF
                     Set list_item = lv_factura.ListItems.Add(, , rsaux!VCHA_aRT_ARTICULO_ID)
                     list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_aRT_NOMBRE_eSPAÑOL), "", rsaux!VCHA_aRT_NOMBRE_eSPAÑOL)
                     list_item.SubItems(2) = IIf(IsNull(rsaux!FLOA_SAL_cANTIDAD), 0, rsaux!FLOA_SAL_cANTIDAD)
                     list_item.SubItems(3) = IIf(IsNull(rsaux!FLOA_SAL_cANTIDAD), 0, rsaux!FLOA_SAL_cANTIDAD)
                     var_cantidad = var_cantidad + IIf(IsNull(rsaux!FLOA_SAL_cANTIDAD), 0, rsaux!FLOA_SAL_cANTIDAD)
                     Me.txt_cantidad = Format(var_cantidad, "###,###,##0.0000")
                     Me.txt_correccion = Format(var_cantidad, "###,###,##0.0000")
                     rsaux.MoveNext:
               Wend
               
               
               
            Else
               MsgBox "La factura no existe", vbOKOnly, "ATENCION"
               Me.txt_cliente = ""
               Me.txt_Estatus = ""
               Me.txt_correccion = ""
               Me.txt_cantidad = ""
               Me.lv_factura.ListItems.Clear
            End If
            rsaux.Close
         Else
            MsgBox "La factura no existe", vbOKOnly, "ATENCION"
            Me.txt_cliente = ""
            Me.txt_Estatus = ""
            Me.txt_correccion = ""
            Me.txt_cantidad = ""
            Me.lv_factura.ListItems.Clear
         End If
         rs.Close
      Else
         MsgBox "Número de factura incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      Me.txt_cliente = ""
      Me.txt_Estatus = ""
      Me.txt_correccion = ""
      Me.txt_cantidad = ""
      Me.lv_factura.ListItems.Clear
   End If
End Sub
