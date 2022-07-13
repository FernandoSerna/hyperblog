VERSION 5.00
Begin VB.Form frmimportar_codigos_compucaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar códigos del compucaja"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_importar 
      Caption         =   "Importar códigos del compucaja"
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   4245
   End
End
Attribute VB_Name = "frmimportar_codigos_compucaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_importar_Click()
   var_si = MsgBox("¿Desea validar los artículos nuevos?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      rs.Open "SELECT ART_CODIGO AS VCHA_ART_ARTICULO_ID, ART_GTIN AS VCHA_EQU_CODIGO_EQUIVALENTE, EXPR1 AS VCHA_ART_NOMBRE_ESPAÑOL, ART_ULTIMOCOSTO AS MONE_ART_COSTO_ESTANDAR, LPA_PRECIOVENTA /1.16 AS MONE_ART_PRECIO_BASE FROM AVL_PRECIOS", cnn_compucaja, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            rsaux.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If rsaux.EOF Then
               If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "D" Then
                  If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "P" Then
                     If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "R" Then
                        rsaux2.Open "select * from tb_Articulos where vcha_Art_Articulo_id = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If rsaux2.EOF Then
                           var_cadena = "insert into tb_Articulos (vcha_art_articulo_id, vcha_Art_nombre_Español, mone_art_costo_estandar, mone_art_precio_base, dtim_Art_fecha_alta, vcha_Art_catalogo_inicio, vcha_Art_catalogo_vigente, vcha_lic_licencia_id, vcha_Art_numero_lic, vcha_tal_talla_id, vcha_uni_unidad_id, inte_art_detenido, vcha_emp_empresa_id )"
                           var_cadena = var_cadena + " values ('" + rs!vcha_Art_Articulo_id + "','" + rs!vcha_Art_nombre_español + "'," + CStr(rs!mone_Art_costo_estandar) + "," + CStr(rs!mone_art_precio_base) + ",getdate(),'CANTIA','CANTIA','SIN LICENCIA','','UNI','01',0,'31')"
                           rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux1.Open "update tb_Articulos set vcha_art_nombre_Español = '" + rs!vcha_Art_nombre_español + "', inte_art_detenido = 0", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux2.Close
                     End If
                  End If
               End If
            End If
            rsaux.Close
            var_si_equivalencia = IIf(IsNull(rs!vcha_equ_codigo_equivalente), "", rs!vcha_equ_codigo_equivalente)
            If var_si_equivalencia <> "" Then
               rsaux.Open "SELECT * FROM tb_equivalencias WHERE vcha_equ_codigo_equivalente = '" + rs!vcha_equ_codigo_equivalente + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux.EOF Then
                  If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "D" Then
                     If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "P" Then
                        If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "R" Then
                           var_cadena = "insert into tb_equivalencias (vcha_art_articulo_id, vcha_equ_codigo_equivalente)"
                           var_cadena = var_cadena + " values ('" + rs!vcha_Art_Articulo_id + "','" + rs!vcha_equ_codigo_equivalente + "')"
                           rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        End If
                     End If
                  End If
               End If
               rsaux.Close
            End If
            
            rsaux.Open "SELECT * FROM tb_Detalle_lista_precios WHERE VCHA_aRT_ARTICULO_ID = '" + rs!vcha_Art_Articulo_id + "' and vcha_lis_lista_precios_id = '01'", cnn, adOpenDynamic, adLockOptimistic
            If rsaux.EOF Then
               If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "D" Then
                  If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "P" Then
                     If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "R" Then
                        var_cadena = "insert into tb_Detalle_lista_precios (vcha_art_articulo_id, vcha_lis_lista_precios_id, floa_dli_precio)"
                        'MsgBox rs!vcha_Art_articulo_id
                        var_cadena = var_cadena + " values ('" + rs!vcha_Art_Articulo_id + "','01'," + CStr(rs!mone_art_precio_base) + ") "
                        rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     End If
                  End If
               End If
            End If
            rsaux.Close
            
            rs.MoveNext
      Wend
      rs.Close
      MsgBox "Se termino la importación de los artículos", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub
