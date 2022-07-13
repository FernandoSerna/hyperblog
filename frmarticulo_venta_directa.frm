VERSION 5.00
Begin VB.Form frmarticulo_venta_directa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Articulo"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_precio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   1815
      TabIndex        =   5
      Top             =   1185
      Width           =   1365
   End
   Begin VB.TextBox txt_descripcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1815
      TabIndex        =   3
      Top             =   660
      Width           =   5295
   End
   Begin VB.TextBox txt_codigo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1815
      TabIndex        =   0
      Top             =   135
      Width           =   2055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Precio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   165
      TabIndex        =   4
      Top             =   1245
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   165
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   165
      TabIndex        =   1
      Top             =   195
      Width           =   990
   End
End
Attribute VB_Name = "frmarticulo_venta_directa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmbusqueda_articulo.Show 1
      Me.txt_codigo = var_codigo_seleccionado
      Me.txt_descripcion = var_descripcion_global
      Me.txt_precio = var_precio_global
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_codigo_LostFocus()
   If var_empresa = "30" Then
      If Mid(Me.txt_codigo, 1, 2) = "TR" Then
         rs.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_descripcion = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
            rsaux.Open "select * from tb_Detalle_lista_precios where vcha_lis_lista_precios_id = '" + var_lista_precios_global + "' and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_precio = IIf(IsNull(rsaux!floa_dli_precio), 0, rsaux!floa_dli_precio)
            Else
               Me.txt_precio = ""
            End If
            rsaux.Close
         Else
            'MsgBox "El código no existe", vbOKOnly, "ATENCION"
            Me.txt_descripcion = ""
            Me.txt_precio = ""
         End If
         rs.Close
      End If
   End If
   If var_empresa = "16" Then
      If Mid(Me.txt_codigo, 1, 2) = "MG" Then
         rs.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_descripcion = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
            rsaux.Open "select * from tb_Detalle_lista_precios where vcha_lis_lista_precios_id = '" + var_lista_precios_global + "' and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_precio = IIf(IsNull(rsaux!floa_dli_precio), 0, rsaux!floa_dli_precio)
            Else
               Me.txt_precio = ""
            End If
            rsaux.Close
         Else
            'MsgBox "El código no existe", vbOKOnly, "ATENCION"
            Me.txt_descripcion = ""
            Me.txt_precio = ""
         End If
         rs.Close
      End If
   End If
   If var_empresa = "30" Or var_empresa = "32" Or var_empresa = "33" Or var_empresa = "34" Or var_empresa = "35" Or var_empresa = "36" Or var_empresa = "37" Or var_empresa = "38" Or var_empresa = "39" Or var_empresa = "40" Or var_empresa = "41" Or var_empresa = "42" Or var_empresa = "43" Or var_empresa = "44" Or var_empresa = "29" Then
      If Mid(Me.txt_codigo, 1, 2) = "EX" Then
         rs.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_descripcion = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
            rsaux.Open "select * from tb_Detalle_lista_precios where vcha_lis_lista_precios_id = '" + var_lista_precios_global + "' and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_precio = IIf(IsNull(rsaux!floa_dli_precio), 0, rsaux!floa_dli_precio)
            Else
               Me.txt_precio = ""
            End If
            rsaux.Close
         Else
            'MsgBox "El código no existe", vbOKOnly, "ATENCION"
            Me.txt_descripcion = ""
            Me.txt_precio = ""
         End If
         rs.Close
      End If
   End If
   
End Sub

Private Sub txt_descripcion_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmbusqueda_articulo.Show 1
      Me.txt_codigo = var_codigo_seleccionado
      Me.txt_descripcion = var_descripcion_global
      Me.txt_precio = var_precio_global
   End If
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   pro_enfoque (KeyAscii)
End Sub

Private Sub txt_precio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmbusqueda_articulo.Show 1
      Me.txt_codigo = var_codigo_seleccionado
      Me.txt_descripcion = var_descripcion_global
      Me.txt_precio = var_precio_global
   End If
End Sub

Private Sub txt_precio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Me.txt_codigo) <> "" Then
         If Trim(Me.txt_descripcion) <> "" Then
            If IsNumeric(Me.txt_precio) Then
               var_posible = 0
               If var_empresa = "30" Then
                  If Mid(Me.txt_codigo, 1, 2) = "TR" Then
                     var_posible = 1
                  End If
               End If
               If var_empresa = "16" Then
                  If Mid(Me.txt_codigo, 1, 2) = "MG" Then
                     var_posible = 1
                  End If
               End If
               If var_empresa = "32" Or var_empresa = "33" Or var_empresa = "34" Or var_empresa = "35" Or var_empresa = "36" Or var_empresa = "37" Or var_empresa = "38" Or var_empresa = "39" Or var_empresa = "40" Or var_empresa = "41" Or var_empresa = "42" Or var_empresa = "43" Or var_empresa = "44" Or var_empresa = "29" Then
                  If Mid(Me.txt_codigo, 1, 2) = "EX" Then
                     var_posible = 1
                  End If
               End If
               If var_posible = 1 Then
                  var_descripcion_global = Me.txt_descripcion
                  var_precio_global = Me.txt_precio
                  var_codigo_seleccionado = Me.txt_codigo
                  rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If rs.EOF Then
                     var_si = MsgBox("El artículo no existe ¿Desea darlo de alta?", vbYesNo, "ATENCION")
                     If var_si = 6 Then
                        rsaux.Open "INSERT INTO TB_ARTICULOS (VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_ESPAÑOL, MONE_ART_COSTO_eSTANDAR, MONE_ART_PRECIO_BASE, VCHA_TAL_TALLA_ID, VCHA_UNI_UNIDAD_ID, VCHA_EMP_EMPRESA_ID) VALUES ('" + Me.txt_codigo + "','" + Me.txt_descripcion + "',0," + Me.txt_precio + ",'UNI','01','" + var_empresa + "')", cnn, adOpenDynamic, adLockOptimistic
                        rsaux.Open "SELECT * FROM TB_dETALLE_LISTA_PRECIOS WHERE VCHA_LIS_LISTA_PRECIOS_ID = '" + var_lista_precios_global + "' AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        If rsaux.EOF Then
                           rsaux1.Open "INSERT INTO TB_DETALLE_LISTA_PRECIOS (VCHA_LIS_LISTA_PRECIOS_ID, VCHA_ART_ARTICULO_ID, FLOA_DLI_PRECIO) VALUES ('" + var_lista_precios_global + "', '" + Me.txt_codigo + "'," + Me.txt_precio + ")", cnn, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux1.Open "UPDATE TB_DETALLE_LISTA_PRECIOS  SET FLOA_DLI_PRECIO = " + Me.txt_precio + " WHERE VCHA_ART_aRTICULO_ID = '" + Me.txt_codigo + "' AND VCHA_LIS_LISTA_PRECIOS_ID = '" + var_lista_precios_global + "'", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux.Close
                     End If
                  Else
                     rsaux.Open "UPDATE TB_aRTICULOS SET MONE_aRT_PRECIO_BASE = " + Me.txt_precio + " WHERE VCHA_aRT_aRTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     rsaux.Open "SELECT * FROM TB_dETALLE_LISTA_PRECIOS WHERE VCHA_LIS_LISTA_PRECIOS_ID = '" + var_lista_precios_global + "' AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If rsaux.EOF Then
                        rsaux1.Open "INSERT INTO TB_DETALLE_LISTA_PRECIOS (VCHA_LIS_LISTA_PRECIOS_ID, VCHA_ART_ARTICULO_ID, FLOA_DLI_PRECIO) VALUES ('" + var_lista_precios_global + "', '" + Me.txt_codigo + "'," + Me.txt_precio + ")", cnn, adOpenDynamic, adLockOptimistic
                     Else
                        rsaux1.Open "UPDATE TB_DETALLE_LISTA_PRECIOS  SET FLOA_DLI_PRECIO = " + Me.txt_precio + " WHERE VCHA_ART_aRTICULO_ID = '" + Me.txt_codigo + "' AND VCHA_LIS_LISTA_PRECIOS_ID = '" + var_lista_precios_global + "'", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux.Close
                     var_descripcion_global = Me.txt_descripcion
                     var_precio_global = Me.txt_precio
                     var_codigo_seleccionado = Me.txt_codigo
                     Unload Me
                  End If
                  rs.Close
                  Unload Me
               Else
                  MsgBox "Código incorrecto para la empresa seleccionada", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Precio incorrecto", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Descripción incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a seleccionado un código", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      var_codigo_seleccionado = ""
      var_descripcion_global = ""
      var_precio_global = 0
      Unload Me
   End If
End Sub
