VERSION 5.00
Begin VB.Form frmclientes_multibondeados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Información de clientes"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9270
      Picture         =   "frmclientes_multibondeados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   495
      Picture         =   "frmclientes_multibondeados.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   165
      Picture         =   "frmclientes_multibondeados.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   75
      Left            =   165
      TabIndex        =   16
      Top             =   315
      Width           =   9480
   End
   Begin VB.Frame Frame1 
      Caption         =   " Cliente "
      Height          =   2295
      Left            =   180
      TabIndex        =   0
      Top             =   420
      Width           =   9465
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   360
         Left            =   3120
         TabIndex        =   14
         Top             =   240
         Width           =   6195
      End
      Begin VB.TextBox txt_cliente 
         Height          =   360
         Left            =   1635
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txt_nombre_establecimiento 
         Height          =   360
         Left            =   3120
         TabIndex        =   8
         Top             =   1800
         Width           =   6195
      End
      Begin VB.TextBox txt_establecimiento 
         Height          =   360
         Left            =   1635
         TabIndex        =   7
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txt_nombre_grupo_actual 
         Enabled         =   0   'False
         Height          =   360
         Left            =   3120
         TabIndex        =   6
         Top             =   1410
         Width           =   6195
      End
      Begin VB.TextBox txt_grupo_actual 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1635
         TabIndex        =   5
         Top             =   1410
         Width           =   1455
      End
      Begin VB.TextBox txt_nombre_grupo_real 
         Enabled         =   0   'False
         Height          =   360
         Left            =   3120
         TabIndex        =   4
         Top             =   1020
         Width           =   6195
      End
      Begin VB.TextBox txt_grupo_real 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1635
         TabIndex        =   3
         Top             =   1020
         Width           =   1455
      End
      Begin VB.TextBox txt_nombre_titular 
         Enabled         =   0   'False
         Height          =   360
         Left            =   3120
         TabIndex        =   2
         Top             =   630
         Width           =   6195
      End
      Begin VB.TextBox txt_titular 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1635
         TabIndex        =   1
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   210
         TabIndex        =   15
         Top             =   323
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   1890
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Grupo Actual:"
         Height          =   195
         Left            =   210
         TabIndex        =   11
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Grupo Real:"
         Height          =   195
         Left            =   210
         TabIndex        =   10
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   720
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmclientes_multibondeados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_nuevo_Click()
   Me.txt_cliente = ""
   Me.txt_nombre_cliente = ""
   Me.txt_establecimiento = ""
   Me.txt_nombre_establecimiento = ""
   Me.txt_grupo_real = ""
   Me.txt_nombre_grupo_real = ""
   Me.txt_grupo_actual = ""
   Me.txt_nombre_grupo_actual = ""
   Me.txt_cliente.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub txt_cliente_LostFocus()
   If Me.txt_cliente <> "" Then
      rs.Open "select vcha_cli_nombre, vcha_tit_titular_id, vcha_tit_nombre, vcha_gre_grupo_real_id, vcha_gre_nombre, vcha_gac_grupo_actual_id from vw_clientes where vcha_cli_clave_id = '" + Me.txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         Me.txt_titular = IIf(IsNull(rs!VCHA_TIT_TITULAR_ID), "", rs!VCHA_TIT_TITULAR_ID)
         Me.txt_nombre_titular = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
         Me.txt_grupo_real = IIf(IsNull(rs!VCHA_GRE_GRUPO_REAL_ID), "", rs!VCHA_GRE_GRUPO_REAL_ID)
         Me.txt_nombre_grupo_real = IIf(IsNull(rs!VCHA_GRE_NOMBRE), "", rs!VCHA_GRE_NOMBRE)
         Me.txt_grupo_actual = IIf(IsNull(VCHA_GAC_GRUPO_ACTUAL_ID), "", rs!VCHA_GAC_GRUPO_ACTUAL_ID)
         Me.txt_nombre_grupo_actual = IIf(IsNull(rs!VCHA_GAC_NOMBRE), "", rs!VCHA_GAC_NOMBRE)
         rsaux.Open "SELECT * FROM TB_DETALLE_eSTABLECIMIENTOS WHERE VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux1.Open "SELECT * FROM TB_eSTABLECIMIENTOS where vcha_esb_establecimiento_id ='" + Me.txt_establecimiento + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               Me.txt_establecimiento = IIf(IsNull(rsaux1!vcha_esb_establecimiento_id), "", rsaux1!vcha_esb_establecimiento_id)
               Me.txt_nombre_establecimiento = IIf(IsNull(rsuax1!vcha_esb_nombre), "", rsaux1!vcha_esb_nombre)
            End If
            rsaux1.Close
         Else
         End If
         rsaux.Close
      Else
         MsgBox "El cliente no existe", vbOKOnly, "ATENCION"
         Me.txt_nombre_cliente = ""
         Me.txt_cliente = ""
         Me.txt_titular = ""
         Me.txt_nombre_titular = ""
         Me.txt_nombre_grupo_real = ""
         Me.txt_grupo_real = ""
         Me.txt_grupo_actual = ""
         Me.txt_nombre_grupo_actual = ""
         Me.txt_establecimiento = ""
         Me.txt_nombre_establecimiento = ""
      End If
      rs.Close
   Else
      Me.txt_titular = ""
      Me.txt_grupo_actual = ""
      Me.txt_grupo_real = ""
      Me.txt_establecimiento = ""
      Me.txt_nombre_cliente = ""
      Me.txt_nombre_establecimiento = ""
      Me.txt_nombre_grupo_actual = ""
      Me.txt_nombre_grupo_real = ""
      Me.txt_nombre_titular = ""
   End If
End Sub
