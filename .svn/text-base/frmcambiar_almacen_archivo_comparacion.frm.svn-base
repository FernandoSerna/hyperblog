VERSION 5.00
Begin VB.Form frmcambiar_almacen_archivo_comparacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar almacén"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_folio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   525
      Left            =   345
      TabIndex        =   1
      Top             =   255
      Width           =   4080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Almacén temporal de entradas de producción"
      Height          =   780
      Left            =   300
      TabIndex        =   0
      Top             =   960
      Width           =   4125
   End
End
Attribute VB_Name = "frmcambiar_almacen_archivo_comparacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   
   rs.Open "select * from tb_Archivo_comparacion where substring(vcha_com_referencia,1,3) = 'EPU' and vcha_com_referencia = '" + Me.txt_folio + "' and dtim_com_fecha >= {d '2011-01-06'} and dtim_com_Fecha < {d '2011-01-09'}", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_posible = 1
      While Not rs.EOF
            var_cantidad = IIf(IsNull(rs!FLOA_com_cANTIDAD_RECIBIDA), 0, rs!FLOA_com_cANTIDAD_RECIBIDA)
            If var_cantidad > 0 Then
               var_posible = 0
            End If
            rs.MoveNext
      Wend
      If var_posible = 1 Then
         rsaux.Open "UPDATE TB_ARCHIVO_COMPARACION SET VCHA_ALM_ALMACEN_ID = 'ATEP' WHERE VCHA_COM_REFERENCIA = '" + Me.txt_folio + "'", cnn, adOpenDynamic, adLockOptimistic
      Else
         MsgBox "La nota ya no puede ser modificada", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Nota incorrecta", vbOKOnly, "ATENCION"
   End If
   rs.Close
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_articulos2)
End Sub
