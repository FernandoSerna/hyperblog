VERSION 5.00
Begin VB.Form frmcancela_documentos_blanco 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelación de documentos fiscales en blanco"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   7035
   Begin VB.Frame Frame2 
      Height          =   75
      Left            =   60
      TabIndex        =   10
      Top             =   330
      Width           =   6885
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmcancela_documentos_blanco.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cancelar Alt + C"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6540
      Picture         =   "frmcancela_documentos_blanco.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir Esc"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Documento a cancelar "
      Height          =   1455
      Left            =   90
      TabIndex        =   0
      Top             =   435
      Width           =   6825
      Begin VB.ComboBox cmb_series 
         Height          =   315
         Left            =   1500
         TabIndex        =   6
         Top             =   645
         Width           =   795
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   1500
         TabIndex        =   5
         Top             =   990
         Width           =   1035
      End
      Begin VB.ComboBox cmb_documentos 
         Height          =   315
         ItemData        =   "frmcancela_documentos_blanco.frx":0784
         Left            =   2550
         List            =   "frmcancela_documentos_blanco.frx":0791
         TabIndex        =   3
         Top             =   300
         Width           =   4155
      End
      Begin VB.TextBox txt_documento 
         Height          =   315
         Left            =   1500
         TabIndex        =   2
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   495
         TabIndex        =   7
         Top             =   705
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   495
         TabIndex        =   4
         Top             =   1050
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Documento: "
         Height          =   195
         Left            =   495
         TabIndex        =   1
         Top             =   360
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmcancela_documentos_blanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_serie As String
Private Sub cmb_documentos_Click()
   If cmb_documentos = "FACTURA" Then
      txt_documento = "FA"
   End If
   If cmb_documentos = "NOTA DE CREDITO" Then
      txt_documento = "NC"
   End If
   If cmb_documentos = "NOTA DE CARGO" Then
      txt_documento = "NG"
   End If
End Sub

Private Sub cmb_documentos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_series.Enabled = True Then
         cmb_series.SetFocus
      Else
         txt_numero.SetFocus
      End If
   End If
End Sub

Private Sub cmb_series_Click()
   var_clase = cmb_series
End Sub

Private Sub cmb_series_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_numero.SetFocus
   End If
End Sub

Private Sub cmd_deshacer_Click()
   
End Sub

Private Sub cmd_cancelar_Click()
Dim si As Integer
Dim var_documento As String
Dim var_clase_documento As String
Dim var_afectacion As String
Dim var_cadena As String
Dim var_tipo_cancelacion As String
Set TB_ENCABEZA_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
If Trim(txt_documento) <> "" Then
   If Trim(txt_numero) <> "" Then
      rs.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Car_tipo_documento = '" + txt_documento + "' and inte_Car_numero = " + txt_numero, cnn, adOpenDynamic, adLockBatchOptimistic
      If Not rs.EOF Then
         rs.Close
         rs.Open "SELECT * FROM VW_DOCUMENTOS_DEL_DIA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_CAR_TIPO_DOCUMENTO = '" + txt_documento + "' AND INTE_CAR_NUMERO = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If rs!vcha_car_documento = "CF" Or rs!vcha_car_documento = "CN" Or rs!vcha_car_documento = "CG" Then
               MsgBox "El documento a sido cancelado con anterioridad", vbOKOnly, "ATENCION"
            Else
               si = MsgBox("¿Deseas cancelar el documento " + Trim(cmb_documentos) + " serie " + Trim(cmb_series) + " número " + txt_numero, vbYesNo, "ATENCION")
               If si = 6 Then
                  si = MsgBox("Confirmar la cancelación del documento", vbYesNo, "ATENCION")
                  If si = 6 Then
                     If txt_documento = "FA" Then
                        var_tipo_cancelacion = "CF"
                     End If
                     If txt_documento = "NC" Then
                        var_tipo_cancelacion = "CN"
                     End If
                     If txt_documento = "NG" Then
                        var_tipo_cancelacion = "CG"
                     End If
                     var_documento = rs!vcha_car_documento
                     var_clase_documento = rs!vcha_car_clase_id
                     var_afectacion = rs!char_car_afectacion
                     If var_afectacion = "+" Then
                        rsaux.Open "select * from tb_estado_cuenta where vcha_Emp_empresa_id = '" + var_empresa + "'  and vcha_ecu_serie_cargo = '" + var_serie + "'  and vcha_Ecu_movimiento_cargo = '" + txt_documento + "' and inte_Ecu_numero_cargo = " + txt_numero + " and floa_ecu_importe_abono > 0", cnn, adOpenDynamic, adLockOptimistic
                        If rsaux.EOF Then
                           var_inserta = TB_ENCABEZA_CARTERA_I.Anadir(rs!vcha_emp_empresa_id, rs!vcha_uor_unidad_id, var_tipo_cancelacion, var_tipo_cancelacion, var_tipo_cancelacion, rs!INTE_CAR_NUMERO, "-", _
                           rs!vcha_alm_almacen_id, rs!vcha_mov_movimiento_id, rs!inte_emo_numero, rs!DTIM_CAR_FECHA, rs!vcha_age_agente_id, _
                           rs!vcha_gac_grupo_Actual_id, rs!vcha_gre_grupo_real_id, rs!vcha_tit_titular_id, rs!vcha_cli_clave_id, rs!vcha_esb_establecimiento_id, _
                           rs!inte_car_PLAZO, rs!floa_car_porcentaje_iva, rs!floa_Car_porcentaje_impuesto_1, rs!floa_car_porcentaje_impuesto_2, rs!floa_car_porcentaje_descuento_1, _
                           rs!floa_car_porcentaje_descuento_2, rs!floa_car_porcentaje_descuento_3, rs!floa_car_importe_total, rs!floa_car_importe_iva, rs!floa_car_importe_impuesto_1, _
                           rs!floa_car_importe_impuesto_2, rs!floa_car_importe_descuento_1, rs!floa_car_importe_descuento_2, rs!floa_car_importe_descuento_3, rs!floa_car_subimporte, _
                           rs!FLOA_CAR_IMPORTE_NETO, rs!vcha_car_importe_letra, rs!Vcha_aud_usuario, rs!Vcha_aud_maquina, rs!vcha_aud_fecha, _
                           0, Date, Date, rs!vcha_mon_moneda_id, rs!floa_car_tipo_cambio, rs!vcha_Ser_serie_id)
                           var_insertar = False
                           var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, cmb_series, txt_documento, txt_numero, rs!vcha_Ser_serie_id, var_tipo_cancelacion, rs!INTE_CAR_NUMERO, 0, rs!FLOA_CAR_IMPORTE_NETO / rs!floa_car_tipo_cambio)
                        Else
                           MsgBox "El documento ya no puede ser cancelado ya que tiene abonos", vbOKOnly, "ATENCION"
                        End If
                        rsaux.Close
                     End If
                     If var_afectacion = "-" Then
                     End If
                  Else
                     MsgBox "Se a cancelado la cancelación del documento", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "Se a cancelado la cancelación del documento", vbOKOnly, "ATENCION"
               End If
            End If
         Else
            MsgBox "El documento no existe o fue elaborado otro dia", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         rs.Close
         si = MsgBox("¿Deseas cancelar el documento " + Trim(cmb_documentos) + " serie " + Trim(cmb_series) + " número " + txt_numero, vbYesNo, "ATENCION")
         If si = 6 Then
            si = MsgBox("Confirmar la cancelación del documento", vbYesNo, "ATENCION")
            If si = 6 Then
               If txt_documento = "FA" Then
                  var_tipo_cancelacion = "CF"
                  var_afectacion = "+"
               End If
               If txt_documento = "NC" Then
                  var_tipo_cancelacion = "CN"
                  var_afectacion = "-"
               End If
               If txt_documento = "NG" Then
                  var_tipo_cancelacion = "CG"
                  var_afectacion = "+"
               End If
               var_documento = txt_documento
               var_clase_documento = var_clase
               If var_afectacion = "+" Then
                  var_inserta = TB_ENCABEZA_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, txt_documento, var_tipo_cancelacion, var_tipo_cancelacion, Val(txt_numero), var_afectacion, _
                  "", "", 0, Date, "", "", "", "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, "", 0, var_serie)
                  var_inserta = TB_ENCABEZA_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, var_tipo_cancelacion, var_tipo_cancelacion, var_tipo_cancelacion, Val(txt_numero), var_afectacion, _
                  "", "", 0, Date, "", "", "", "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, "", 0, var_serie)
               End If
               If var_afectacion = "-" Then
                  var_inserta = TB_ENCABEZA_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, txt_documento, var_tipo_cancelacion, var_tipo_cancelacion, Val(txt_numero), var_afectacion, _
                  "", "", 0, Date, "", "", "", "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, "", 0, var_serie)
                  var_inserta = TB_ENCABEZA_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, var_tipo_cancelacion, var_tipo_cancelacion, var_tipo_cancelacion, Val(txt_numero), var_afectacion, _
                  "", "", 0, Date, "", "", "", "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, "", 0, var_serie)
               End If
            Else
               MsgBox "Se a cancelado la cancelación del documento", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Se a cancelado la cancelación del documento", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "Número de documento incorrecto", vbOKOnly, "ATENCION"
   End If
Else
   MsgBox "Documento incorrecto", vbOKOnly, "ATENCION"
End If
End Sub


Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Top = 2500
   Left = 1900
   rs.Open "select vcha_ser_serie_id from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_contador_serie = 0
      While Not rs.EOF
         var_contador_serie = var_contador_serie + 1
         rs.MoveNext
      Wend
      rs.MoveFirst
      txt_documento.Enabled = True
      cmb_documentos.Enabled = True
      txt_numero.Enabled = True
      Call RecsetToCombo(cmb_series.hwnd, rs, 0)
      If var_contador_serie > 1 Then
         cmb_series.Enabled = True
      Else
         cmb_series.Enabled = False
      End If
      rs.MoveFirst
      cmb_series = rs!vcha_Ser_serie_id
      var_serie = rs!vcha_Ser_serie_id
   Else
      MsgBox "No se a indicado una serie para esta Unidad organizacional", vbOKOnly, "ATENCION"
      txt_documento.Enabled = False
      cmb_documentos.Enabled = False
      txt_numero.Enabled = False
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
    If var_activa_menu = True Then
       Frmmenu2.Enabled = True
    End If
End Sub

Private Sub txt_documento_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      cmb_documentos.SetFocus
   End If
End Sub

Private Sub txt_documento_LostFocus()
   If Trim(txt_documento) <> "" Then
      If txt_documento = "FA" Then
         cmb_documentos = "FACTURA"
      Else
         If txt_documento = "NC" Then
            cmb_documentos = "NOTA DE CREDITO"
         Else
            If txt_documento = "NG" Then
               cmb_documentos = "NOTA DE CARGO"
            Else
               MsgBox "Clave de documento incorrecta", vbOKOnly, "ATENCION"
               txt_documento = ""
               cmb_documentos = ""
            End If
         End If
      End If
   End If
End Sub

Private Sub txt_numero_LostFocus()
   If Not IsNumeric(txt_numero) Then
      MsgBox "Número de documento incorrecto", vbOKOnly, "ATENCION"
      txt_numero = ""
   End If
End Sub
