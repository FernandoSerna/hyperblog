VERSION 5.00
Begin VB.Form frmcancela_cajas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelación de Cajas"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3345
   Begin VB.TextBox txt_codigo_caja 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   225
      TabIndex        =   0
      Top             =   210
      Width           =   2910
   End
End
Attribute VB_Name = "frmcancela_cajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   var_cadena_seguridad = ""
   var_activa_menu = True
   Top = 3000
   Left = 3850
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call activa_forma(var_activa_forma_cancela_cajas)
End Sub

Private Sub txt_codigo_caja_KeyPress(KeyAscii As Integer)
Dim var_empaque As Double
Dim var_caja As Integer
Dim var_posible_caja As Boolean
Dim var_numero_folio As Integer
Dim var_almacen_Destino As String
Dim var_orden_surtido As Integer
Dim list_item As ListItem
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(txt_codigo_caja) <> "" Then
         If Left(Trim(txt_codigo_caja), 1) = "C" Then
            x = Mid(txt_codigo_caja, 2, 6)
            If IsNumeric(x) Then
               var_empaque = x
               x = Mid(txt_codigo_caja, 8, 3)
               If IsNumeric(x) Then
                  var_caja = x
                  var_posible_caja = True
               Else
                  var_posible_caja = False
               End If
            Else
               var_posible_caja = False
            End If
         Else
            var_posible_caja = False
         End If
         If var_posible_caja = True Then
            frmdetalle_cajas.txt_empaque = var_empaque
            frmdetalle_cajas.txt_caja = var_caja
            frmdetalle_cajas.cmd_aceptar.Enabled = False
            rs.Open "SELECT vcha_art_articulo_id,floa_paq_cantidad,char_paq_estatus FROM TB_DETALLE_cajas WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND INTE_EMB_EMBARQUE = " + Str(var_empaque) + " AND INTE_PAQ_CAJA = " + Str(var_caja), cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If rs!char_paq_estatus = "S" Then
                  MsgBox "La caja ya fue surtida", vbOKOnly, "ATENCION"
                  frmdetalle_cajas.cmd_aceptar.Enabled = False
               Else
                  If rs!char_paq_estatus = "C" Then
                     MsgBox "La caja ya fue cancelada", vbOKOnly, "ATENCION"
                     frmdetalle_cajas.cmd_aceptar.Enabled = False
                  Else
                     frmdetalle_cajas.cmd_aceptar.Enabled = True
                  End If
               End If
            Else
               MsgBox "La caja no existe", vbOKOnly, "ATENCION"
               frmdetalle_cajas.cmd_aceptar.Enabled = False
            End If
            While Not rs.EOF
               Set list_item = frmdetalle_cajas.lv_detalle_cajas.ListItems.Add(, , rs!VCHA_aRT_ARTICULO_ID)
               rsaux3.Open "select vcha_art_nombre_Español from tb_articulos where vcha_Art_articulo_id = '" + rs!VCHA_aRT_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rsaux3(0).Value), "", rsaux3(0).Value)
               rsaux3.Close
               list_item.SubItems(2) = IIf(IsNull(rs!floa_paq_cantidad), 0, rs!floa_paq_cantidad)
               rs.MoveNext
            Wend
            rs.Close
            frmdetalle_cajas.Show
         Else
            MsgBox "El código no pertenece a una caja", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub
