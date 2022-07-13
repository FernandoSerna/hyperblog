VERSION 5.00
Begin VB.Form frminventario_textilera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Poner digito verificador a la tabla de inventarios"
      Height          =   660
      Left            =   225
      TabIndex        =   0
      Top             =   255
      Width           =   4230
   End
End
Attribute VB_Name = "frminventario_textilera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Dim VERIFICADOR As Integer
   'para que funciones esto debio de haberse llenado la tabla de tb_inventario con la tabla invfinal
   rs.Open "select vcha_inv_codigo from tb_inventario where len(vcha_inv_codigo) = 11", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_codigo = rs!vcha_inv_codigo
         sum1 = 0
         sum2 = 0
         mcodigo = var_codigo
         longitud = Len(mcodigo)
         For icont = 1 To longitud
             If ((icont / 2) - Int((icont / 2))) = 0 Then
                sum2 = sum2 + Val(Mid(mcodigo, icont, 1))
             Else
                sum1 = sum1 + Val(Mid(mcodigo, icont, 1))
             End If
         Next icont
         msuma = sum1 * 13 + sum2
         VERIFICADOR = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
         If VERIFICADOR = 10 Then
            VERIFICADOR = 0
         End If
         var_codigo = var_codigo + Trim(CStr(VERIFICADOR))
         rsaux.Open "update tb_inventario set vcha_Art_articulo_id = '" + var_codigo + "' where vcha_inv_codigo = '" + rs!vcha_inv_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub
