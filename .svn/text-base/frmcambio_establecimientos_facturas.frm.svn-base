VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmcambio_establecimientos_facturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar establecimiento a facturas"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   240
      TabIndex        =   20
      Top             =   450
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   21
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
         TabIndex        =   22
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5745
      Picture         =   "frmcambio_establecimientos_facturas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmcambio_establecimientos_facturas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmcambio_establecimientos_facturas.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Caption         =   " Establecimiento nuevo "
      Height          =   750
      Left            =   75
      TabIndex        =   16
      Top             =   2565
      Width           =   5985
      Begin VB.TextBox txt_nombre_establecimiento_nuevo 
         Height          =   345
         Left            =   1590
         TabIndex        =   9
         Top             =   270
         Width           =   4245
      End
      Begin VB.TextBox txt_establecimiento_nuevo 
         Height          =   345
         Left            =   165
         TabIndex        =   8
         Top             =   270
         Width           =   1395
      End
   End
   Begin VB.Frame Frame2 
      Height          =   75
      Left            =   15
      TabIndex        =   11
      Top             =   315
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos de la factura "
      Height          =   1995
      Left            =   75
      TabIndex        =   10
      Top             =   480
      Width           =   5985
      Begin VB.TextBox txt_fecha 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4365
         TabIndex        =   18
         Top             =   1500
         Width           =   1500
      End
      Begin VB.TextBox txt_serie 
         Height          =   375
         Left            =   1425
         TabIndex        =   3
         Top             =   255
         Width           =   615
      End
      Begin VB.TextBox txt_importe 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1425
         TabIndex        =   7
         Top             =   1500
         Width           =   1500
      End
      Begin VB.TextBox txt_establecimiento 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1425
         TabIndex        =   6
         Top             =   1095
         Width           =   4440
      End
      Begin VB.TextBox txt_cliente 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1425
         TabIndex        =   5
         Top             =   675
         Width           =   4440
      End
      Begin VB.TextBox txt_factura 
         Height          =   375
         Left            =   2940
         TabIndex        =   4
         Top             =   255
         Width           =   1500
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   3675
         TabIndex        =   19
         Top             =   1590
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Factura:"
         Height          =   195
         Left            =   2235
         TabIndex        =   17
         Top             =   345
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   1590
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   1185
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   765
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   345
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmcambio_establecimientos_facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_guardar_Click()
   If IsNumeric(Me.txt_factura) Then
      If Me.txt_establecimiento_nuevo <> "" Then
         var_cadena = " SELECT     dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO AS IMPORTE, dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_ALM_ALMACEN_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO FROM dbo.TB_ENCABEZADO_CARTERA INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID WHERE (dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = " + Me.txt_factura + ") AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = '" + Me.txt_serie + "')"
         rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rs.Open "UPDATE TB_ENCABEZADO_MOVIMIENTOS SET VCHA_ESB_eSTABLECIMIENTO_ID = '" + Me.txt_establecimiento_nuevo + "' WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + IIf(IsNull(rsaux!VCHA_UOR_UNIDAD_ID), "", rsaux!VCHA_UOR_UNIDAD_ID) + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + IIf(IsNull(rsaux!VCHA_MOV_MOVIMIENTO_ID), "", rsaux!VCHA_MOV_MOVIMIENTO_ID) + "' AND INTE_EMO_NUMERO = " + CStr(IIf(IsNull(rsaux!INTE_EMO_NUMERO), 0, rsaux!INTE_EMO_NUMERO)), cnn, adOpenDynamic, adLockOptimistic
            MsgBox "Se a actualizado el movimiento", vbOKOnly, "ATENCION"
         End If
         rsaux.Close
      End If
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_cliente = ""
   Me.txt_establecimiento = ""
   Me.txt_establecimiento_nuevo = ""
   Me.txt_factura = ""
   Me.txt_importe = ""
   Me.txt_nombre_establecimiento_nuevo = ""
   Me.txt_fecha = ""
   Me.txt_serie.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Top = 1000
    Left = 2000
    Me.frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_lista.ListItems.Count > 0 Then
         Me.txt_establecimiento_nuevo = Me.lv_lista.selectedItem
         Me.txt_nombre_establecimiento_nuevo = Me.lv_lista.selectedItem.SubItems(1)
         Me.txt_establecimiento_nuevo.SetFocus
      Else
         Me.txt_establecimiento_nuevo.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.txt_establecimiento_nuevo.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_establecimiento_nuevo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsNumeric(Me.txt_factura) Then
         var_cadena = " SELECT     dbo.TB_ENCABEZADO_CARTERA.vcha_cli_clave_id, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_NOMBRE, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO AS IMPORTE, dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID fROM dbo.TB_ENCABEZADO_CARTERA INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND "
         var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN dbo.TB_ESTABLECIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID = dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID WHERE (dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = " + Me.txt_factura + ") AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = '" + Me.txt_serie + "') "
         rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_clave_cliente = IIf(IsNull(rsaux!VCHA_CLI_CLAVE_ID), "", rsaux!VCHA_CLI_CLAVE_ID)
         Else
            var_clave_cliente = ""
         End If
         rsaux.Close
         lv_lista.ListItems.Clear
         rs.Open "select * from VW_DETALLE_ESTABELCIMIENTOS WHERE VCHA_cLI_CLAVE_ID = '" + var_clave_cliente + "' ORDER by vcha_esb_nombre", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ESB_ESTABLECIMIENTO_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "Establecimientos"
         VAR_TIPO_LISTA = 21
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
   End If
End Sub

Private Sub txt_establecimiento_nuevo_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_establecimiento_nuevo_LostFocus()
   If Trim(Me.txt_establecimiento_nuevo) <> "" Then
      rs.Open "select * from tb_establecimientos where vCha_esb_establecimiento_id = '" + Me.txt_establecimiento_nuevo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_establecimiento_nuevo = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
      Else
         MsgBox "El establecimiento no existe", vbOKOnly, "ATENCION"
         Me.txt_establecimiento_nuevo = ""
         Me.txt_nombre_establecimiento_nuevo = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_factura_Change()
   Me.txt_cliente = ""
   Me.txt_establecimiento = ""
   Me.txt_establecimiento_nuevo = ""
   Me.txt_importe = ""
   Me.txt_fecha = ""
   Me.txt_nombre_establecimiento_nuevo = ""
End Sub

Private Sub txt_factura_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_factura_LostFocus()
   If IsNumeric(Me.txt_factura) Then
      var_cadena = " SELECT     dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_NOMBRE, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO AS IMPORTE, dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID fROM dbo.TB_ENCABEZADO_CARTERA INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND "
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN dbo.TB_ESTABLECIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID = dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID WHERE (dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = " + Me.txt_factura + ") AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = '" + Me.txt_serie + "') "
      rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         Me.txt_establecimiento = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
         Me.txt_importe = Format(IIf(IsNull(rs!Importe), 0, rs!Importe), "###,###,##0.00")
         Me.txt_fecha = IIf(IsNull(rs!dtim_Car_fecha), "", rs!dtim_Car_fecha)
         Me.txt_establecimiento_nuevo = ""
         Me.txt_nombre_establecimiento_nuevo = ""
      Else
         MsgBox "La factura no existe", vbOKOnly, "ATENCION"
         Me.txt_cliente = ""
         Me.txt_establecimiento = ""
         Me.txt_establecimiento_nuevo = ""
         Me.txt_importe = ""
         Me.txt_nombre_establecimiento_nuevo = ""
         Me.txt_fecha = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_nombre_establecimiento_nuevo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsNumeric(Me.txt_factura) Then
         var_cadena = " SELECT dbo.TB_ENCABEZADO_CARTERA.VCHA_cLI_CLAVE_ID, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_NOMBRE, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO AS IMPORTE fROM dbo.TB_ENCABEZADO_CARTERA INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_ESTABLECIMIENTOS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_ESB_ESTABLECIMIENTO_ID = dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID wHERE (dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = " + Me.txt_factura + ") AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = '" + Me.txt_serie + "') "
         rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_clave_cliente = IIf(IsNull(rsaux!VCHA_CLI_CLAVE_ID), "", rsaux!VCHA_CLI_CLAVE_ID)
         Else
            var_clave_cliente = ""
         End If
         rsaux.Close
         lv_lista.ListItems.Clear
         rs.Open "select * from VW_DETALLE_ESTABELCIMIENTOS WHERE VCHA_cLI_CLAVE_ID = '" + var_clave_cliente + "' ORDER by vcha_esb_nombre", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ESB_ESTABLECIMIENTO_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "Establecimientos"
         VAR_TIPO_LISTA = 21
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
   End If
End Sub

Private Sub txt_nombre_establecimiento_nuevo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Me.cmd_guardar.SetFocus
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub txt_serie_Change()
   Me.txt_cliente = ""
   Me.txt_establecimiento = ""
   Me.txt_establecimiento_nuevo = ""
   Me.txt_importe = ""
   Me.txt_nombre_establecimiento_nuevo = ""
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub
