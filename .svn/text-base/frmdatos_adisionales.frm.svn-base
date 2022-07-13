VERSION 5.00
Begin VB.Form frmdatos_adisionales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos adicionales"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_cancelar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmdatos_adisionales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Cancelar Alt + C"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmdatos_adisionales.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   45
      TabIndex        =   17
      Top             =   315
      Width           =   7260
   End
   Begin VB.Frame frm_datos_adicionales 
      Height          =   2850
      Left            =   45
      TabIndex        =   9
      Top             =   405
      Width           =   7185
      Begin VB.TextBox txt_numero_interno 
         Height          =   315
         Left            =   4110
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1665
         Width           =   1080
      End
      Begin VB.TextBox txt_nombre_2 
         Height          =   315
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   0
         Top             =   210
         Width           =   5385
      End
      Begin VB.TextBox txt_paterno 
         Height          =   315
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   1
         Top             =   570
         Width           =   5385
      End
      Begin VB.TextBox txt_materno 
         Height          =   315
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   2
         Top             =   945
         Width           =   5385
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1665
         Width           =   1080
      End
      Begin VB.TextBox txt_clave_tel_pais 
         Height          =   315
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2040
         Width           =   1080
      End
      Begin VB.TextBox txt_clave_tel_estado 
         Height          =   315
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2400
         Width           =   1080
      End
      Begin VB.TextBox txt_calle 
         Height          =   315
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1305
         Width           =   5385
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número interno:"
         Height          =   195
         Left            =   2820
         TabIndex        =   20
         Top             =   1740
         Width           =   1125
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   300
         TabIndex        =   16
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Paterno:"
         Height          =   195
         Left            =   300
         TabIndex        =   15
         Top             =   645
         Width           =   600
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Materno:"
         Height          =   195
         Left            =   300
         TabIndex        =   14
         Top             =   1005
         Width           =   630
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   1740
         Width           =   600
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Clave tel. Pais:"
         Height          =   195
         Left            =   300
         TabIndex        =   12
         Top             =   2100
         Width           =   1050
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Clave tel. Estado:"
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   2460
         Width           =   1245
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Calle:"
         Height          =   195
         Left            =   300
         TabIndex        =   10
         Top             =   1380
         Width           =   390
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número:"
      Height          =   195
      Left            =   2850
      TabIndex        =   19
      Top             =   2175
      Width           =   600
   End
End
Attribute VB_Name = "frmdatos_adisionales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
   var_nombre_cliente_ad = Me.txt_nombre_2
   var_paterno_cliente_ad = Me.txt_paterno
   var_materno_cliente_ad = Me.txt_materno
   var_numero_cliente_ad = Me.txt_numero
   var_calle_cliente_ad = Me.txt_calle
   var_clave_tel_pais_ad = Me.txt_clave_tel_pais
   var_clave_tel_estado_ad = Me.txt_clave_tel_estado
   var_numero_interno_cliente_ad = Me.txt_numero_interno
   If var_tipo_datos_adicionales = 2 Then
      rsaux.Open "update tb_establecimientos set vcha_Esb_nombre_2 = '" + Me.txt_nombre_2 + "', vcha_esb_paterno = '" + Me.txt_paterno + "', vcha_esb_materno = '" + Me.txt_materno + "', vcha_esb_numero = '" + Me.txt_numero + "', vcha_esb_calle = '" + Me.txt_calle + "', vcha_esb_clave_tel_pais = '" + Me.txt_clave_tel_pais + "', vcha_esb_clave_tel_estado = '" + Me.txt_clave_tel_estado + "', vcha_esb_numero_interno = '" + Me.txt_numero_interno + "' where vcha_esb_establecimiento_id = '" + var_clave_establecimiento_global + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   If var_tipo_datos_adicionales = 3 Then
      rsaux.Open "update tb_titulares set vcha_tit_nombre_2 = '" + Me.txt_nombre_2 + "', vcha_tit_paterno = '" + Me.txt_paterno + "', vcha_tit_materno = '" + Me.txt_materno + "', vcha_tit_numero = '" + Me.txt_numero + "', vcha_tit_calle = '" + Me.txt_calle + "', vcha_tit_clave_tel_pais = '" + Me.txt_clave_tel_pais + "', vcha_tit_clave_tel_estado = '" + Me.txt_clave_tel_estado + "', vcha_tit_numero_interno = '" + Me.txt_numero_interno + "' where vcha_tit_titular_id = '" + var_clave_titular_global + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   Unload Me
End Sub

Private Sub cmd_cancelar_pedidos_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.txt_nombre_2 = var_nombre_cliente_ad
   Me.txt_paterno = var_paterno_cliente_ad
   Me.txt_materno = var_materno_cliente_ad
   Me.txt_numero = var_numero_cliente_ad
   Me.txt_calle = var_calle_cliente_ad
   Me.txt_numero_interno = var_numero_interno_cliente_ad
   Me.txt_clave_tel_pais = var_clave_tel_pais_ad
   Me.txt_clave_tel_estado = var_clave_tel_estado_ad
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_calle_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_tel_estado_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_tel_pais_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_materno_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_2_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_interno_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_paterno_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub
