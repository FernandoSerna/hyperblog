VERSION 5.00
Begin VB.Form frmcancela_facturas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelación de Facturas"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   7290
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmcancela_facturas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6885
      Picture         =   "frmcancela_facturas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   0
      TabIndex        =   20
      Top             =   330
      Width           =   7260
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmcancela_facturas.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Cancelar Alt + C"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Factura "
      Height          =   2355
      Left            =   75
      TabIndex        =   0
      Top             =   480
      Width           =   7170
      Begin VB.ComboBox cmb_series 
         Height          =   315
         Left            =   6135
         TabIndex        =   23
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox txt_fecha 
         Height          =   315
         Left            =   3855
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   225
         Width           =   1515
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1950
         Width           =   1665
      End
      Begin VB.TextBox txt_nombre_establecimiento 
         Height          =   315
         Left            =   2955
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1605
         Width           =   3975
      End
      Begin VB.TextBox txt_clave_establecimiento 
         Height          =   315
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1605
         Width           =   1665
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2955
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1260
         Width           =   3975
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   315
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1260
         Width           =   1665
      End
      Begin VB.TextBox txt_nombre_titular 
         Height          =   315
         Left            =   2955
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   915
         Width           =   3975
      End
      Begin VB.TextBox txt_clave_titular 
         Height          =   315
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   915
         Width           =   1665
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   2955
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   570
         Width           =   3975
      End
      Begin VB.TextBox txt_clave_agente 
         Height          =   315
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   570
         Width           =   1665
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   1275
         TabIndex        =   2
         Top             =   225
         Width           =   1665
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   5655
         TabIndex        =   24
         Top             =   300
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   3300
         TabIndex        =   18
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   105
         TabIndex        =   15
         Top             =   2010
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Left            =   105
         TabIndex        =   14
         Top             =   1665
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   105
         TabIndex        =   13
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Left            =   105
         TabIndex        =   12
         Top             =   975
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   105
         TabIndex        =   11
         Top             =   630
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   105
         TabIndex        =   1
         Top             =   285
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmcancela_facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_serie  As String

Private Sub cmb_series_Click()
   var_serie = cmb_series
End Sub

Private Sub cmd_cancelar_Click()
   Dim var_fecha As Date
   Dim var_fecha_factura As Date
   Dim si As Integer
   rs.Open "select getdate()", cnn, adOpenDynamic, adLockOptimistic
   var_fecha = Format(rs(0).Value, "Short Date")
   rs.Close
   var_fecha_factura = txt_fecha
   If var_fecha_factura <> var_fecha Then
      MsgBox "No es posible cancelar la factura ya que no corresponde a la fecha actual", vbOKOnly, "ATENCION"
   Else
      si = MsgBox("¿Deseas cancelar la factura " + txt_numero, vbYesNo, "ATENCION")
      If si = 6 Then
         si = MsgBox("Confirmar la cancelación de la factura", vbYesNo, "ATENCION")
         If si = 6 Then
            rs.Open "Update tb_encabezado_cartera set char_car_estatus = 'C' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "' and inte_Car_numero = " + txt_numero + " and vcha_car_documento = 'FA'", cnn, adOpenDynamic, adLockOptimistic
            MsgBox "Se a terminado la cancelación de la factura", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 1800
   Left = 2200
   rs.Open "select vcha_ser_serie_id from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_contador_serie = 0
      While Not rs.EOF
         var_contador_serie = var_contador_serie + 1
         rs.MoveNext
      Wend
      rs.MoveFirst
      txt_numero.Enabled = True
      Call RecsetToCombo(cmb_series.hwnd, rs, 0)
      If var_contador_serie > 1 Then
         cmb_series.Enabled = True
      Else
         cmb_series.Enabled = False
      End If
      rs.MoveFirst
      var_serie = rs!vcha_ser_serie_id
      cmb_series = var_serie
   Else
      MsgBox "No se a indicado una serie para esta Unidad organizacional", vbOKOnly, "ATENCION"
      txt_numero.Enabled = False
   End If
   rs.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_cancela_facturas)
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_fecha.SetFocus
   End If
End Sub

Private Sub txt_numero_LostFocus()
   If Trim(txt_numero) <> "" Then
      If IsNumeric(txt_numero) Then
         rs.Open "select * from vw_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "' and inte_Car_numero = " + txt_numero + " and vcha_car_documento = 'FA'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_fecha = IIf(IsNull(rs!DTIM_car_FECHA), "", Format(rs!DTIM_car_FECHA, "Short Date"))
            txt_clave_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
            txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            txt_clave_agente = IIf(IsNull(rs!vcha_age_agente_id), "", rs!vcha_age_agente_id)
            txt_nombre_agente = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
            txt_clave_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
            txt_nombre_titular = IIf(IsNull(rs!VCHA_tit_NOMBRE), "", rs!VCHA_tit_NOMBRE)
            txt_importe = Format(IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto), "###,###,##0.00")
            txt_clave_establecimiento = IIf(IsNull(rs!vcha_esb_establecimiento_id), 0, rs!vcha_esb_establecimiento_id)
            txt_nombre_establecimiento = IIf(IsNull(rs!vcha_esb_nombre), 0, rs!vcha_esb_nombre)
         Else
            MsgBox "La factura " + txt_numero + " no existe", vbOKOnly, "ATENCION"
            txt_fecha = ""
            txt_clave_cliente = ""
            txt_nombre_cliente = ""
            txt_clave_agente = ""
            txt_nombre_agente = ""
            txt_clave_titular = ""
            txt_nombre_titular = ""
            txt_importe = ""
            txt_clave_establecimiento = ""
            txt_nombre_establecimiento = ""
         End If
         rs.Close
      Else
         txt_fecha = ""
         txt_clave_cliente = ""
         txt_nombre_cliente = ""
         txt_clave_agente = ""
         txt_nombre_agente = ""
         txt_clave_titular = ""
         txt_nombre_titular = ""
         txt_importe = ""
         txt_clave_establecimiento = ""
         txt_nombre_establecimiento = ""
         MsgBox "Número de Factura Incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      txt_fecha = ""
      txt_clave_cliente = ""
      txt_nombre_cliente = ""
      txt_clave_agente = ""
      txt_nombre_agente = ""
      txt_clave_titular = ""
      txt_nombre_titular = ""
      txt_importe = ""
      txt_clave_establecimiento = ""
      txt_nombre_establecimiento = ""
   End If
End Sub
