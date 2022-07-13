VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmembarques_paquetes_2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empaques"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frmembarques_paquetes_2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5745
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmembarques_paquetes_2.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Guardar Alt + G"
      Top             =   60
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5340
      Picture         =   "frmembarques_paquetes_2.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   60
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos del Empaque"
      Height          =   1305
      Left            =   15
      TabIndex        =   0
      Top             =   510
      Width           =   5655
      Begin VB.ComboBox cmb_clientes 
         Height          =   315
         Left            =   2115
         TabIndex        =   3
         Top             =   750
         Width           =   3420
      End
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   1005
         TabIndex        =   2
         Top             =   750
         Width           =   1095
      End
      Begin VB.TextBox txt_paquete 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1005
         TabIndex        =   1
         Top             =   405
         Width           =   1095
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   525
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   465
         Width           =   600
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   4860
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques_paquetes_2.frx":1006
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques_paquetes_2.frx":18E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques_paquetes_2.frx":21BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques_paquetes_2.frx":2756
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques_paquetes_2.frx":3032
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques_paquetes_2.frx":390C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques_paquetes_2.frx":41E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques_paquetes_2.frx":42F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques_paquetes_2.frx":440A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques_paquetes_2.frx":451C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques_paquetes_2.frx":462E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   345
      Width           =   5655
   End
End
Attribute VB_Name = "frmembarques_paquetes_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_numero_paquete As Integer

Private Sub cmb_clientes_Click()
   txt_cliente = Obtener_llave(cnn, rsaux, "TB_clientes", "VCHA_cli_NOMBRE", cmb_clientes, 0, "T")
End Sub

Private Sub cmd_guardar_Click()
   Set TB_ENCABEZADO_PAQUETES_I = New TB_ENCABEZADO_PAQUETES_I
         If Trim(txt_cliente) <> "" Then
            rs.Open "select maximo_paquete from vw_maximo_paquete where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_numero_paquete = rs!maximo_paquete + 1
            Else
               var_numero_paquete = 1
            End If
            rs.Close
            ok = TB_ENCABEZADO_PAQUETES_I.Anadir(var_empresa, var_unidad_organizacional, var_numero_paquete, txt_cliente)
            Unload Me
         Else
            MsgBox "No se a seleccionado al cliente", vbOKOnly, "ATENCION"
         End If

End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 71 Then
      cmd_guardar_Click
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   var_cadena_seguridad = ""
   rs.Open "select maximo_paquete from vw_maximo_paquete where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_numero_paquete = rs!maximo_paquete + 1
   Else
      var_numero_paquete = 1
   End If
   rs.Close
   txt_paquete = var_numero_paquete
   rs.Open "select * from tb_clientes order by vcha_cli_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_clientes.hwnd, rs, 1)
   rs.Close
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_embarques_paquetes_2)
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(txt_cliente) <> "" Then
         rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rs.Close
            cmb_clientes.Text = rs!vcha_cli_nombre
            txt_cliente.Enabled = False
            cmb_clientes.Enabled = False
         Else
            rs.Close
            txt_cliente = ""
            cmb_clientes.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txt_paquete_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub
