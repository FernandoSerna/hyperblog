VERSION 5.00
Begin VB.Form frmliberar_apartado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liberación de apartado"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7335
      Picture         =   "frmliberar_apartado.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmliberar_apartado.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmliberar_apartado.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Nuevo Pedido Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   15
      TabIndex        =   15
      Top             =   270
      Width           =   7725
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos del artículo "
      Height          =   1635
      Left            =   45
      TabIndex        =   0
      Top             =   465
      Width           =   7590
      Begin VB.TextBox txt_eliminar 
         Height          =   350
         Left            =   2505
         TabIndex        =   7
         Top             =   1140
         Width           =   1395
      End
      Begin VB.TextBox txt_disponible 
         Height          =   350
         Left            =   6045
         TabIndex        =   6
         Top             =   735
         Width           =   1395
      End
      Begin VB.TextBox txt_apartado 
         Height          =   350
         Left            =   3345
         TabIndex        =   5
         Top             =   735
         Width           =   1395
      End
      Begin VB.TextBox txt_existen 
         Height          =   350
         Left            =   870
         TabIndex        =   4
         Top             =   735
         Width           =   1395
      End
      Begin VB.TextBox txt_nombre 
         Height          =   350
         Left            =   2370
         TabIndex        =   3
         Top             =   345
         Width           =   5070
      End
      Begin VB.TextBox txt_codigo 
         Height          =   350
         Left            =   870
         TabIndex        =   2
         Top             =   345
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad eliminar del apartado:"
         Height          =   195
         Left            =   165
         TabIndex        =   14
         Top             =   1215
         Width           =   2175
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Disponibles:"
         Height          =   195
         Left            =   4995
         TabIndex        =   13
         Top             =   810
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Apartadas:"
         Height          =   195
         Left            =   2505
         TabIndex        =   12
         Top             =   810
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Existen:"
         Height          =   195
         Left            =   165
         TabIndex        =   11
         Top             =   810
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   165
         TabIndex        =   1
         Top             =   420
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmliberar_apartado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_almacen As String
Private Sub cmd_aceptar_pedidos_Click()
   If IsNumeric(Me.txt_eliminar) Then
      If CDbl(Me.txt_eliminar) > 0 Then
         If CDbl(Me.txt_eliminar) <= CDbl(Me.txt_apartado) Then
            var_si = MsgBox("Desea eliminar la cantidad disponible", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_si = MsgBox("Confirmar la eliminacion de la cantidad disponible", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  rsaux.Open "UPDATE TB_EXISTENCIAS SET FLOA_eXI_CANTIDAD_apartada = ISNULL(FLOA_EXI_CANTIDAD_apartada,0) - " + Me.txt_eliminar + " WHERE VCHA_ALM_ALMACEN_ID = '" + var_almacen + "' AND VCHA_aRT_ARTICULo_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  Me.txt_apartado = CDbl(Me.txt_apartado) - CDbl(Me.txt_eliminar)
                  Me.txt_disponible = CDbl(Me.txt_disponible) + CDbl(Me.txt_eliminar)
                  MsgBox "Se a eliminado la cantidad correctamente", vbOKOnly, "ATENCION"
               End If
            End If
         Else
            MsgBox "La cantidad a eliminar no debe de ser mayo a la cantidad apartada", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "La cantidad a eliminar debe de ser mayo que 0", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Cantidad a eliminar incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
  Me.txt_codigo = ""
  Me.txt_nombre = ""
  Me.txt_apartado = ""
  Me.txt_disponible = ""
  Me.txt_existen = ""
  Me.txt_eliminar = ""
  Me.txt_codigo.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_almacen = "8"
   Top = 3000
   Left = 2000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_apartado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   If Trim(Me.txt_codigo) <> "" Then
      rs.Open "select * from tb_articulos where vcha_Art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre = IIf(IsNull(rs!VCHA_aRT_NOMBRE_ESPAÑOL), "", rs!VCHA_aRT_NOMBRE_ESPAÑOL)
         rsaux3.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ALM_ALMACEN_ID = '" + var_almacen + "' AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux3.EOF Then
            Me.txt_existen = IIf(IsNull(rsaux3!FLOA_EXI_CANTIDAD), 0, rsaux3!FLOA_EXI_CANTIDAD)
            Me.txt_apartado = IIf(IsNull(rsaux3!FLOA_EXI_cANTIDAD_APARTADA), 0, rsaux3!FLOA_EXI_cANTIDAD_APARTADA)
            Me.txt_disponible = IIf(IsNull(rsaux3!FLOA_EXI_CANTIDAD_DISPONIBLE), 0, rsaux3!FLOA_EXI_CANTIDAD_DISPONIBLE)
            Me.txt_eliminar = ""
            Me.txt_eliminar.SetFocus
         Else
            Me.txt_existen = 0
            Me.txt_apartado = 0
            Me.txt_disponible = 0
            Me.txt_eliminar = ""
            Me.txt_eliminar.SetFocus
         End If
         rsaux3.Close
      Else
         rsaux.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux1.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + IIf(IsNull(rsaux!VCHA_aRT_ARTICULO_ID), "", rsaux!VCHA_aRT_ARTICULO_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               Me.txt_codigo = IIf(IsNull(rsaux!VCHA_aRT_ARTICULO_ID), "", rsaux!VCHA_aRT_ARTICULO_ID)
               Me.txt_nombre = IIf(IsNull(rsaux1!VCHA_aRT_NOMBRE_ESPAÑOL), "", rsaux1!VCHA_aRT_NOMBRE_ESPAÑOL)
               rsaux3.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ALM_ALMACEN_ID = '" + var_almacen + "' AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux3.EOF Then
                  Me.txt_existen = IIf(IsNull(rsaux3!FLOA_EXI_CANTIDAD), 0, rsaux3!FLOA_EXI_CANTIDAD)
                  Me.txt_apartado = IIf(IsNull(rsaux3!FLOA_EXI_cANTIDAD_APARTADA), 0, rsaux3!FLOA_EXI_cANTIDAD_APARTADA)
                  Me.txt_disponible = IIf(IsNull(rsaux3!FLOA_EXI_CANTIDAD_DISPONIBLE), 0, rsaux3!FLOA_EXI_CANTIDAD_DISPONIBLE)
                  Me.txt_eliminar = ""
                  Me.txt_eliminar.SetFocus
              Else
                  Me.txt_existen = 0
                  Me.txt_apartado = 0
                  Me.txt_disponible = 0
                  Me.txt_eliminar = ""
                  Me.txt_eliminar.SetFocus
              End If
              rsaux3.Close
            
            Else
               MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
               Me.txt_apartado = ""
               Me.txt_codigo = ""
               Me.txt_disponible = ""
               Me.txt_eliminar = ""
               Me.txt_existen = ""
               Me.txt_nombre = ""
            End If
            rsaux1.Close
         Else
            MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
            Me.txt_apartado = ""
            Me.txt_codigo = ""
            Me.txt_disponible = ""
            Me.txt_eliminar = ""
            Me.txt_existen = ""
            Me.txt_nombre = ""
         End If
         rsaux.Close
      End If
      rs.Close
   Else
      Me.txt_apartado = ""
      Me.txt_codigo = ""
      Me.txt_disponible = ""
      Me.txt_eliminar = ""
      Me.txt_existen = ""
      Me.txt_nombre = ""
   End If
End Sub

Private Sub txt_disponible_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_aceptar_pedidos.SetFocus
   End If
End Sub

Private Sub txt_existen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub
