VERSION 5.00
Begin VB.Form frmalta_codigos_textilera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alta de códigos vianney-textilera"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_ejecutar 
      Caption         =   "Ejecutar alta de artículos"
      Height          =   900
      Left            =   60
      TabIndex        =   0
      Top             =   195
      Width           =   4545
   End
End
Attribute VB_Name = "frmalta_codigos_textilera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ejecutar_Click()
   Dim VERIFICADOR As Integer
   rsaux11.Open "select codigo from codigos_descuento_20_150910", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux11.EOF
         If rsaux10.State = 1 Then
            rsaux10.Close
         End If
         rsaux10.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux11!codigo + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
         If Not rsaux10.EOF Then
            var_equivalencia = rsaux11!codigo
            var_linea = rsaux10!vcha_lin_linea_id
            var_costo = rsaux10!mone_Art_costo_estandar
            var_precio = rsaux10!mone_Art_precio_base
            var_catalogo = rsaux10!vcha_Art_catalogo_vigente
            var_DEscripcion = rsaux10!vcha_Art_nombre_español
            If var_linea = "00" Then
               linea_textilera = "13"
            End If
            If var_linea = "10" Then
               linea_textilera = "12"
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
            var_codigo = Mid(var_equivalencia, 7, 5)
            var_codigo_textilera = "6" + var_linea + "00" + var_codigo + "0"
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
             
            rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_aRTICULO_ID = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If rs.EOF Then
               var_cadena = "INSERT INTO TB_ARTICULOS (VCHA_aRT_aRTICULO_ID, VCHA_aRT_nombre_español, MONE_ART_PRECIO_BASE, MONE_ART_COSTO_ESTANDAR, VCHA_LIN_LINEA_ID, VCHA_ART_CATALOGO_VIGENTE, VCHA_TPR_TIPO_PRODUCTO_ID,       VCHA_DIV_DIVISION_ID,            VCHA_SUB_SUBDIVISION_ID,          VCHA_EST_ESTAMPADO_ID) VALUES"
               var_cadena = var_cadena + "('" + var_codigo_textilera + "', '" + var_DEscripcion + "', " + CStr(var_precio) + ", " + CStr(var_costo) + ",    '" + var_linea + "',  '" + var_catalogo + "', '" + Mid(var_codigo_textilera, 1, 1) + "','" + Mid(var_codigo_textilera, 2, 2) + "', '" + Mid(var_codigo_textilera, 4, 2) + "', '" + Mid(var_codigo_textilera, 6, 5) + "')"
               rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               rsaux.Open "INSERT INTO TB_DETALLE_LISTA_PRECIOS (VCHA_LIS_LISTA_PRECIOS_ID, VCHA_ART_ARTICULO_ID, FLOA_DLI_PRECIO) VALUES ('01','" + var_codigo_textilera + "', " + CStr(var_precio) + ")", cnn, adOpenDynamic, adLockOptimistic
               rsaux.Open "select * from tb_estampados where vcha_est_estampado_id = '" + Mid(var_codigo_textilera, 6, 5) + "'"
               If rsaux.EOF Then
                  rsaux2.Open "insert into tb_estampados (vcha_est_estampado_id, vcha_est_nombre) values ('" + Mid(var_codigo_textilera, 6, 5) + "', '" + var_DEscripcion + "')", cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux.Close
               rsaux.Open "select * from tb_equivalencias where vcha_art_articulo_id = '" + var_codigo_textilera + "' and vcha_equ_codigo_equivalente = '" + var_equivalencia + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux.EOF Then
                  rsaux2.Open "insert into tb_equivalencias (vcha_equ_codigo_equivalente, vcha_art_articulo_id) values ('" + var_equivalencia + "', '" + var_codigo_textilera + "')", cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux.Close
            End If
            rs.Close
            var_codigo_1 = Mid(var_codigo_textilera, 1, 10)
            var_codigo_general = var_codigo_textilera
            rsaux.Open "select * from tb_equivalencias where vcha_art_articulo_id = '" + var_codigo_textilera + "' and vcha_equ_codigo_equivalente = '" + var_equivalencia + "'", cnn, adOpenDynamic, adLockOptimistic
            If rsaux.EOF Then
               rsaux2.Open "insert into tb_equivalencias (vcha_equ_codigo_equivalente, vcha_art_articulo_id) values ('" + var_equivalencia + "', '" + var_codigo_textilera + "')", cnn, adOpenDynamic, adLockOptimistic
            End If
            rsaux.Close
         End If
         rsaux11.MoveNext
   Wend
End Sub

Private Sub Form_Load()
   Top = 3500
   Left = 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_articulos2)
End Sub
