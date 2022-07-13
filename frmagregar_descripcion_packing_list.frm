VERSION 5.00
Begin VB.Form frmagregar_descripcion_packing_list 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de contenido para packing list"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmagregar_descripcion_packing_list.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmagregar_descripcion_packing_list.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guardar Alt + G"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7710
      Picture         =   "frmagregar_descripcion_packing_list.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   105
      TabIndex        =   7
      Top             =   285
      Width           =   8010
   End
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   120
      TabIndex        =   6
      Top             =   420
      Width           =   7980
      Begin VB.TextBox txt_contenido 
         Height          =   345
         Left            =   1230
         TabIndex        =   2
         Top             =   630
         Width           =   6540
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   345
         Left            =   2685
         TabIndex        =   1
         Top             =   240
         Width           =   5085
      End
      Begin VB.TextBox txt_articulo 
         Height          =   345
         Left            =   1230
         TabIndex        =   0
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   705
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   315
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmagregar_descripcion_packing_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub cmd_guardar_Click()
   If Me.txt_articulo <> "" Then
      var_si = MsgBox("¿Desea actualizar el registro?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rs.Open "UPDATE TB_ARTICULOS SET VCHA_aRT_DESCRIPCION_PACKING_LIST = '" + Me.txt_contenido + "' WHERE VCHA_aRT_ARTICULO_ID = '" + Me.txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
         MsgBox "Se a actualizado el registro", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un artículo", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_articulo = ""
   Me.txt_contenido = ""
   Me.txt_descripcion = ""
   Me.txt_articulo.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 2500
   Left = 2000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_packing_list)
End Sub

Private Sub txt_articulo_Change()
   Me.txt_contenido = ""
   Me.txt_descripcion = ""
End Sub

Private Sub txt_articulo_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_articulo_LostFocus()
   If Me.txt_articulo <> "" Then
      rs.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + Me.txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_descripcion = IIf(IsNull(rs!VCHA_aRT_NOMBRE_ESPAÑOL), "", rs!VCHA_aRT_NOMBRE_ESPAÑOL)
         Me.txt_contenido = IIf(IsNull(rs!VCHA_aRT_DESCRIPCION_PACKING_LIST), "", rs!VCHA_aRT_DESCRIPCION_PACKING_LIST)
      Else
         rsaux.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Me.txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux1.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID= '" + IIf(IsNull(rsaux!vcha_Art_articulo_id), "", rsaux!vcha_Art_articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               Me.txt_articulo = IIf(IsNull(rsaux1!vcha_Art_articulo_id), "", rsaux1!vcha_Art_articulo_id)
               Me.txt_descripcion = IIf(IsNull(rsaux1!VCHA_aRT_NOMBRE_ESPAÑOL), "", rsaux1!VCHA_aRT_NOMBRE_ESPAÑOL)
               Me.txt_contenido = IIf(IsNull(rsaux1!VCHA_aRT_DESCRIPCION_PACKING_LIST), "", rsaux1!VCHA_aRT_DESCRIPCION_PACKING_LIST)
            Else
               MsgBox "El código no existe", vbOKOnly, "ATENCION"
               Me.txt_contenido = ""
               Me.txt_articulo = ""
               Me.txt_descripcion = ""
            End If
            rsaux1.Close
         Else
            MsgBox "El código no existe", vbOKOnly, ""
            Me.txt_contenido = ""
            Me.txt_articulo = ""
            Me.txt_descripcion = ""
         End If
         rsaux.Close
      End If
      rs.Close
   End If
End Sub

Private Sub txt_contenido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub
