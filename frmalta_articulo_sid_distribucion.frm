VERSION 5.00
Begin VB.Form frmalta_articulo_sid_distribucion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alta de artículos desde el almacén general"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6960
      Picture         =   "frmalta_articulo_sid_distribucion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmalta_articulo_sid_distribucion.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmalta_articulo_sid_distribucion.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   90
      Left            =   15
      TabIndex        =   3
      Top             =   330
      Width           =   7380
   End
   Begin VB.Frame Frame1 
      Caption         =   " Artículo "
      Height          =   900
      Left            =   75
      TabIndex        =   0
      Top             =   465
      Width           =   7215
      Begin VB.TextBox txt_descripcion 
         Height          =   390
         Left            =   1950
         TabIndex        =   2
         Top             =   315
         Width           =   5145
      End
      Begin VB.TextBox txt_codigo 
         Height          =   390
         Left            =   75
         TabIndex        =   1
         Top             =   315
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frmalta_articulo_sid_distribucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_nuevo_Click()
   Me.txt_codigo = ""
   Me.txt_descripcion = ""
   Me.txt_codigo.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_descripcion.SetFocus
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   If Me.txt_codigo <> "" Then
      rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_descripcion = IIf(IsNull(rsaux1!vcha_Art_nombre_Español), "", rsaux1!vcha_Art_nombre_Español)
      Else
         rsaux.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Me.txt_codigo + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux1.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + rsaux!VCHA_aRT_aRTICULO_ID + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               Me.txt_descripcion = IIf(IsNull(rsaux1!vcha_Art_nombre_Español), "", rsaux1!vcha_Art_nombre_Español)
            Else
               Me.txt_descripcion = ""
               MsgBox "El artículo no existe en el almacén general", vbOKOnly, "ATENCION"
            End If
            rsaux1.Close
         Else
            rsaux1.Open "select * from tb_Articulos where substring(vcha_Art_Articulo_id,7,5) = '" + Me.txt_codigo + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               Me.txt_descripcion = IIf(IsNull(rsaux1!vcha_Art_nombre_Español), "", rsaux1!vcha_Art_nombre_Español)
            Else
               Me.txt_descripcion = ""
               MsgBox "El artículo no existe en el almacén general", vbOKOnly, "ATENCION"
            End If
            rsaux1.Close
         End If
         rsaux.Close
      End If
      rs.Close
   Else
      Me.txt_descripcion = ""
   End If
End Sub
