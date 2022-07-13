VERSION 5.00
Begin VB.Form frmalta_articulos_textilera_vianney 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar códigos a textilera"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   45
      TabIndex        =   5
      Top             =   360
      Width           =   5910
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   5475
      Picture         =   "frmalta_articulos_textilera_vianney.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   360
   End
   Begin VB.Frame Frame1 
      Caption         =   " Catálogo "
      Height          =   840
      Left            =   105
      TabIndex        =   4
      Top             =   480
      Width           =   5805
      Begin VB.TextBox txt_nombre_catalogo 
         Height          =   390
         Left            =   1290
         TabIndex        =   1
         Top             =   285
         Width           =   4380
      End
      Begin VB.TextBox txt_codigo 
         Height          =   390
         Left            =   105
         TabIndex        =   0
         Top             =   285
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmd_alta_codigos_textilera_vianney 
      Height          =   345
      Left            =   105
      Picture         =   "frmalta_articulos_textilera_vianney.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Importar artículos"
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmalta_articulos_textilera_vianney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_alta_codigos_textilera_vianney_Click()
   Dim var_bbb As Integer
   Dim var_codigo As String
   Dim var_codigo_anterior As String
   Dim VERIFICADOR As Integer
   If Me.txt_codigo <> "" Then
      'rsaux11.Open "select substring(vcha_Art_Articulo_id, 7,5) as vcha_Art_Articulo_id, mone_Art_costo_estandar from tb_Articulos where substring(vcha_Art_articulo_id,1,6) = '646244' and vcha_Art_catalogo_vigente = '" + Me.txt_codigo + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      rsaux11.Open "SELECT substring(vcha_Art_Articulo_id, 7,5) as vcha_Art_Articulo_id, dbo.tb_ARTICULOS.MONE_ART_COSTO_ESTANDAR FROM dbo.precios_textilera_130111 INNER JOIN dbo.tb_ARTICULOS ON dbo.precios_textilera_130111.codigo = dbo.tb_ARTICULOS.VCHA_ART_ARTICULO_ID ", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rsaux11.EOF
            var_codigo_anterior = rsaux11!vcha_Art_articulo_id
            var_codigo = rsaux11!vcha_Art_articulo_id
            var_costo = IIf(IsNull(rsaux11!mone_Art_costo_estandar), 0, rsaux11!mone_Art_costo_estandar)
            rsaux2.Open "select * from tb_Detalle_lista_precios where vcha_lis_lista_precios_id = '01' and vcha_Art_articulo_id like '646244" + CStr(rsaux11!vcha_Art_articulo_id) + "%'", cnn_distribucion, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               var_precio = IIf(IsNull(rsaux2!floa_dli_precio), 0, rsaux2!floa_dli_precio)
            Else
               var_precio = 0
            End If
            rsaux2.Close
            var_equivalencia = rsaux11!vcha_Art_articulo_id
            rsaux9.Open "select * from tb_Articulos where substring(vcha_art_articulo_id,7,5) = '" + Trim(var_equivalencia) + "' and substring(vcha_Art_articulo_id,1,6) = '646244'", cnn_distribucion, adOpenDynamic, adLockOptimistic
            If Not rsaux9.EOF Then
               var_xx = rsaux9!vcha_Art_articulo_id
               var_linea = IIf(IsNull(rsaux9!vcha_lin_linea_id), "", rsaux9!vcha_lin_linea_id)
               var_costo = IIf(IsNull(rsaux9!mone_Art_costo_estandar), 0, rsaux9!mone_Art_costo_estandar)
               var_catalogo = IIf(IsNull(rsaux9!vcha_Art_catalogo_vigente), "", rsaux9!vcha_Art_catalogo_vigente)
               var_DEscripcion = IIf(IsNull(rsaux9!vcha_Art_nombre_español), "", rsaux9!vcha_Art_nombre_español)
               var_proveedor = ""
               linea_textilera = ""
               If var_linea = "00" Then
                  linea_textilera = "13"
               End If
               If var_linea = "2" Then
                  linea_textilera = "30"
               End If
               If var_linea = "10" Then
                  linea_textilera = "12"
               End If
               If var_linea = "11" Then
                  linea_textilera = "75"
               End If
               If var_linea = "12" Then
                  linea_textilera = "10"
               End If
               If var_linea = "13" Then
                  linea_textilera = "30"
               End If
               If var_linea = "14" Then
                  linea_textilera = "40"
               End If
               If var_linea = "15" Then
                  linea_textilera = "50"
               End If
               If var_linea = "16" Then
                  linea_textilera = "20"
               End If
               If var_linea = "20" Then
                  linea_textilera = "16"
               End If
               If var_linea = "22" Then
                  linea_textilera = "16"
               End If
               If var_linea = "23" Then
                  linea_textilera = "16"
               End If
               If var_linea = "24" Then
                  linea_textilera = "16"
               End If
               If var_linea = "28" Then
                  linea_textilera = "13"
               End If
               If var_linea = "29" Then
                  linea_textilera = "12"
               End If
               If var_linea = "30" Then
                  linea_textilera = "13"
               End If
               If var_linea = "31" Then
                  linea_textilera = "13"
               End If
               If var_linea = "35" Then
                  linea_textilera = "16"
               End If
               If var_linea = "39" Then
                  linea_textilera = "13"
               End If
               If var_linea = "40" Then
                  linea_textilera = "14"
               End If
               If var_linea = "41" Then
                  linea_textilera = "14"
               End If
               If var_linea = "42" Then
                  linea_textilera = "15"
               End If
               If var_linea = "43" Then
                  linea_textilera = "15"
               End If
               If var_linea = "44" Then
                  linea_textilera = "25"
               End If
               If var_linea = "45" Then
                  linea_textilera = "24"
               End If
               If var_linea = "50" Then
                  linea_textilera = "15"
               End If
               If var_linea = "55" Then
                  linea_textilera = "13"
               End If
               If var_linea = "59" Then
                  linea_textilera = "13"
               End If
               If var_linea = "60" Then
                  linea_textilera = "14"
               End If
               If var_linea = "65" Then
                  linea_textilera = "13"
               End If
               If var_linea = "70" Then
                  linea_textilera = "16"
               End If
               If var_linea = "75" Then
                  linea_textilera = "13"
               End If
               If var_linea = "80" Then
                  linea_textilera = "16"
               End If
               If var_linea = "90" Then
                  linea_textilera = "16"
               End If
               If var_linea = "91" Then
                  linea_textilera = "16"
               End If
               If var_linea = "92" Then
                  linea_textilera = "16"
               End If
               If var_linea = "93" Then
                  linea_textilera = "16"
               End If
               If var_linea = "94" Then
                  linea_textilera = "13"
               End If
               If var_linea = "95" Then
                  linea_textilera = "13"
               End If
               Select Case linea_textilera
                      Case ""
                           var_division = "00"
                      Case "10"
                           var_division = "10"
                      Case "12"
                           var_division = "12"
                      Case "13"
                           var_division = "30"
                      Case "14"
                           var_division = "40"
                      Case "15"
                           var_division = "50"
                      Case "16"
                           var_division = "20"
                      Case "20"
                           var_division = "00"
                      Case "22"
                           var_division = "00"
                      Case "24"
                           var_division = "44"
                      Case "25"
                           var_division = "43"
               End Select
               
               For var_bbb = 0 To 9
                   var_codigo = rsaux11!vcha_Art_articulo_id
                   var_codigo_textilera = "6" + var_division + "00" + var_codigo + CStr(var_bbb)
                   txt_tipo = "6"
                   txt_division = var_division
                   txt_subdivision = "00"
                   txt_estampado = rsaux11!vcha_Art_articulo_id
                   txt_descuento = CStr(var_bbb)
                   var_codigo = var_codigo_textilera
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
                   var_codigo_textilera = var_codigo
                   If rs.State = 1 Then
                      rs.Close
                   End If
                   rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_aRTICULO_ID = '" + var_codigo + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                   If rs.EOF Then
                      VAR_dESCUENTO_ARTICULO = (100 - (CDbl(txt_descuento) * 10)) / 100
                      var_precio_text = var_precio * VAR_dESCUENTO_ARTICULO
                      
                      var_cadena = "INSERT INTO TB_ARTICULOS (VCHA_aRT_aRTICULO_ID, VCHA_aRT_nombre_español, MONE_ART_PRECIO_BASE, MONE_ART_COSTO_ESTANDAR, VCHA_LIN_LINEA_ID, VCHA_ART_CATALOGO_VIGENTE, VCHA_TPR_TIPO_PRODUCTO_ID,       VCHA_DIV_DIVISION_ID,            VCHA_SUB_SUBDIVISION_ID,          VCHA_EST_ESTAMPADO_ID, VCHA_eMP_eMPRESA_ID, INTE_ART_ALTA_MASIVA, VCHA_TAL_tALLA_ID, VCHA_UNI_UNIDAD_ID) VALUES"
                      var_cadena = var_cadena + "('" + var_codigo_textilera + "', '" + var_DEscripcion + "', " + CStr(var_precio_text) + ", " + CStr(var_costo) + ",    '" + var_linea + "',  '" + var_catalogo + "', '" + Mid(var_codigo_textilera, 1, 1) + "','" + Mid(var_codigo_textilera, 2, 2) + "', '" + Mid(var_codigo_textilera, 4, 2) + "', '" + Mid(var_codigo_textilera, 6, 5) + "','18',1,'UNI','01')"
                      rsaux.Open var_cadena, cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                      
                      rsaux.Open "select * from tb_detalle_lista_precios where vcha_art_articulo_id = '" + var_codigo_textilera + "' and vcha_lis_lista_precios_id = '01'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                      If rsaux.EOF Then
                         rsaux2.Open "insert into tb_detalle_lista_precios (vcha_art_articulo_id, vcha_lis_lista_precios_id, floa_dli_precio, inte_lis_alta_masiva) values ('" + var_codigo + "','01'," + CStr(var_precio_text) + ",1)", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                      Else
                         rsaux2.Open "update tb_detalle_lista_precios set floa_dli_precio = " + CStr(var_precio_text) + " where vcha_Art_articulo_id =  '" + var_codigo_textilera + "' and vcha_lis_lista_precios_id = '01'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                      End If
                      rsaux.Close
                    
                   
                      rsaux.Open "select * from tb_estampados where vcha_est_estampado_id = '" + Mid(var_codigo_textilera, 6, 5) + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                      If rsaux.EOF Then
                         rsaux2.Open "insert into tb_estampados (vcha_est_estampado_id, vcha_est_nombre, INTE_EST_ALTA_MASIVA) values ('" + Mid(var_codigo_textilera, 6, 5) + "', '" + var_DEscripcion + "',1)", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                      End If
                      rsaux.Close
                      If txt_descuento = "0" Then
                         rsaux.Open "select * from tb_equivalencias where vcha_art_articulo_id = '" + var_codigo_textilera + "' and vcha_equ_codigo_equivalente = '" + var_xx + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                         If rsaux.EOF Then
                            rsaux2.Open "insert into tb_equivalencias (vcha_equ_codigo_equivalente, vcha_art_articulo_id, INTE_EQU_ALTA_MASIVA) values ('" + var_xx + "', '" + var_codigo_textilera + "',1)", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                         End If
                         rsaux.Close
                      End If
                   Else
                      VAR_dESCUENTO_ARTICULO = (100 - (CDbl(txt_descuento) * 10)) / 100
                      var_precio_text = var_precio * VAR_dESCUENTO_ARTICULO
                      
                      rsaux.Open "select * from tb_detalle_lista_precios where vcha_art_articulo_id = '" + var_codigo_textilera + "' and vcha_lis_lista_precios_id = '01'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                      If rsaux.EOF Then
                         rsaux2.Open "insert into tb_detalle_lista_precios (vcha_art_articulo_id, vcha_lis_lista_precios_id, floa_dli_precio, inte_lis_alta_masiva) values ('" + var_codigo + "','01'," + CStr(var_precio_text) + ",1)", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                      Else
                         rsaux2.Open "update tb_detalle_lista_precios set floa_dli_precio = " + CStr(var_precio_text) + " where vcha_Art_articulo_id =  '" + var_codigo_textilera + "' and vcha_lis_lista_precios_id = '01'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                      End If
                      rsaux.Close
                      rsaux.Open "update tb_Articulos set mone_art_precio_base = " + CStr(var_precio_text) + " where vcha_Art_articulo_id =  '" + var_codigo_textilera + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                      If txt_descuento = "0" Then
                         rsaux.Open "select * from tb_equivalencias where vcha_art_articulo_id = '" + var_codigo_textilera + "' and vcha_equ_codigo_equivalente = '" + var_xx + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                         If rsaux.EOF Then
                            rsaux2.Open "insert into tb_equivalencias (vcha_equ_codigo_equivalente, vcha_art_articulo_id, INTE_EQU_ALTA_MASIVA) values ('" + var_xx + "', '" + var_codigo_textilera + "',1)", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                         End If
                         rsaux.Close
                      End If
                      rsaux.Open "select * from tb_estampados where vcha_est_estampado_id = '" + Mid(var_codigo_textilera, 6, 5) + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                      If rsaux.EOF Then
                         rsaux2.Open "insert into tb_estampados (vcha_est_estampado_id, vcha_est_nombre, INTE_EST_ALTA_MASIVA) values ('" + Mid(var_codigo_textilera, 6, 5) + "', '" + var_DEscripcion + "',1)", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                      End If
                      rsaux.Close
                   
                   End If
                   rs.Close
               Next var_bbb
            End If
            rsaux9.Close
            rsaux11.MoveNext
      Wend
      rsaux11.Close
      MsgBox "Se a terminado de importar la información", vbOKOnly, "ATENCION"
   Else
      MsgBox "No se a seleccionado un catálogo", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3400
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_articulos2)
End Sub

Private Sub txt_codigo_Change()
   Me.txt_nombre_catalogo = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   If Trim(Me.txt_codigo) <> "" Then
      rs.Open "select * from tb_Catalogos where vcha_cat_catalogo_id = '" + Me.txt_codigo + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_catalogo = IIf(IsNull(rs!vcha_Cat_nombre), "", rs!vcha_Cat_nombre)
      Else
         MsgBox "Catálogo no existe", vbOKOnly, "ATENCION"
         Me.txt_nombre_catalogo = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_catalogo = ""
   End If
End Sub

Private Sub txt_nombre_catalogo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_alta_codigos_textilera_vianney.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub
