VERSION 5.00
Begin VB.Form frmcorreccion_documentos_fiscales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Correción de plazo y fechas a facturas"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7530
      Picture         =   "frmcorreccion_documentos_fiscales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_aplicar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmcorreccion_documentos_fiscales.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Aplicar"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmcorreccion_documentos_fiscales.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Nuevo "
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   75
      Left            =   60
      TabIndex        =   20
      Top             =   330
      Width           =   7890
   End
   Begin VB.Frame Frame1 
      Caption         =   " Documento "
      Height          =   3030
      Left            =   90
      TabIndex        =   0
      Top             =   435
      Width           =   7815
      Begin VB.TextBox txt_plazo 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1065
         TabIndex        =   9
         Top             =   2520
         Width           =   750
      End
      Begin VB.TextBox txt_fecha 
         Height          =   330
         Left            =   1065
         TabIndex        =   8
         Top             =   2145
         Width           =   1410
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   1065
         TabIndex        =   7
         Top             =   1770
         Width           =   1860
      End
      Begin VB.TextBox txt_nombre 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2310
         TabIndex        =   6
         Top             =   1395
         Width           =   5205
      End
      Begin VB.TextBox txt_clave 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1065
         TabIndex        =   5
         Top             =   1395
         Width           =   1215
      End
      Begin VB.TextBox txt_numero 
         Height          =   330
         Left            =   1065
         TabIndex        =   4
         Top             =   1020
         Width           =   1230
      End
      Begin VB.TextBox txt_serie 
         Height          =   330
         Left            =   1065
         TabIndex        =   3
         Top             =   645
         Width           =   870
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   330
         Left            =   1965
         TabIndex        =   2
         Top             =   270
         Width           =   5565
      End
      Begin VB.TextBox txt_tipo 
         Height          =   330
         Left            =   1065
         TabIndex        =   1
         Top             =   270
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Plazo"
         Height          =   195
         Left            =   210
         TabIndex        =   19
         Top             =   2588
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   2213
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   210
         TabIndex        =   17
         Top             =   1838
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   195
         TabIndex        =   16
         Top             =   1470
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   165
         TabIndex        =   15
         Top             =   1088
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   720
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   330
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmcorreccion_documentos_fiscales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command11_Click()

End Sub

Private Sub cmd_aplicar_Click()
   If Trim(Me.txt_tipo) <> "" Then
      If Trim(Me.txt_serie) <> "" Then
         If Trim(Me.txt_numero) <> "" Then
            If Trim(Me.txt_clave) <> "" Then
               If Trim(Me.txt_importe) <> "" Then
                  If IsDate(Me.txt_fecha) Then
                     If IsNumeric(Me.txt_plazo) Then
                        var_si = MsgBox("¿Deseas cambiar los datos al documento?", vbYesNo, "ATENCION")
                        If var_si = 6 Then
                           var_si = MsgBox("Confirmar el cambio de datos del documento", vbYesNo, "ATENCION")
                           If var_si = 6 Then
                              var_dia = CStr(Day(CDate(Me.txt_fecha)))
                              var_mes = CStr(Month(CDate(Me.txt_fecha)))
                              var_año = CStr(Year(CDate(Me.txt_fecha)))
                              If Len(Trim(var_dia)) = 1 Then
                                 var_dia = "0" + var_dia
                              End If
                              If Len(Trim(var_mes)) = 1 Then
                                 var_mes = "0" + var_mes
                              End If
                              var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                              
                              rs.Open "SELECT * FROM TB_ENCABEZADO_CARTERA WHERE vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and  VCHA_CAR_TIPO_DOCUMENTO = '" + Me.txt_tipo + "' AND VCHA_SER_sERIE_ID = '" + Me.txt_serie + "' AND INTE_cAR_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                              VAR_FECHA_ANTERIOR = Format(IIf(IsNull(rs!DTIM_car_FECHA), "", rs!DTIM_car_FECHA), "SHORT DATE")
                              var_dia = CStr(Day(CDate(VAR_FECHA_ANTERIOR)))
                              var_mes = CStr(Month(CDate(VAR_FECHA_ANTERIOR)))
                              var_año = CStr(Year(CDate(VAR_FECHA_ANTERIOR)))
                              If Len(Trim(var_dia)) = 1 Then
                                 var_dia = "0" + var_dia
                              End If
                              If Len(Trim(var_mes)) = 1 Then
                                 var_mes = "0" + var_mes
                              End If
                              VAR_FECHA_ANTERIOR = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                              rsaux.Open "UPDATE TB_ENCABEZADO_CARTERA SET INTE_CAR_PLAZO = " + Me.txt_plazo + ", DTIM_CAR_FECHA = " + var_fecha + " WHERE vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and  VCHA_CAR_TIPO_DOCUMENTO = '" + Me.txt_tipo + "' AND VCHA_SER_sERIE_ID = '" + Me.txt_serie + "' AND INTE_cAR_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                              Cadena = "INSERT INTO TB_BITACORA_CAMBIO_PLAZOS_FECHAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, DTIM_BIT_FECHA_ANTERIOR, DTIM_BIT_FECHA_ACTUAL, INTE_BIT_PLAZO_ANTERIOR, INTE_BIT_PLAZO_ACTUAL, DTIM_BIT_FECHA_MODIFICACION, VCHA_USU_USUARIO_ID, VCHA_BIT_MAQUINA)"
                              Cadena = Cadena + " VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + Me.txt_tipo + "','" + Me.txt_serie + "' , " + Me.txt_numero + "," + VAR_FECHA_ANTERIOR + ", " + var_fecha + ", " + CStr(IIf(IsNull(rs!INTE_CAR_PLAZO), 0, rs!INTE_CAR_PLAZO)) + ", " + Me.txt_plazo + ",GETDATE(),'" + var_clave_usuario_global + "','" + fun_NombrePc + "')"
                              rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                              rs.Close
                           End If
                        End If
                     Else
                        MsgBox "Plazo incorrecto", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "Fecha incorrecta", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "Importe incorrecto", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Número de documento incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Serie de documento incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Tipo de documento incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_clave = ""
   Me.txt_descripcion = ""
   Me.txt_fecha = ""
   Me.txt_importe = ""
   Me.txt_importe = ""
   Me.txt_nombre = ""
   Me.txt_numero = ""
   Me.txt_plazo = ""
   Me.txt_serie = ""
   Me.txt_tipo = ""
   Me.txt_tipo.Enabled = True
   Me.txt_descripcion.Enabled = True
   Me.txt_serie.Enabled = True
   Me.txt_numero.Enabled = True
   Me.txt_tipo.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_importe_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_Change()
   Me.txt_clave = ""
   Me.txt_fecha = ""
   Me.txt_importe = ""
   Me.txt_importe = ""
   Me.txt_nombre = ""
   Me.txt_plazo = ""
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_LostFocus()
   If IsNumeric(Me.txt_numero) Then
     rs.Open "SELECT * FROM TB_ENCABEZADO_cARTERA WHERE vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and  VCHA_CAR_TIPO_DOCUMENTO = '" + Me.txt_tipo + "' AND VCHA_SER_sERIE_ID = '" + Me.txt_serie + "' AND INTE_cAR_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
     If Not rs.EOF Then
        Me.txt_clave = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
        rsaux.Open "select * from tb_clientes where vcha_cli_clave_id = '" + Me.txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
        If Not rsaux.EOF Then
           Me.txt_nombre = IIf(IsNull(rsaux!VCHA_CLI_NOMBRE), "", rsaux!VCHA_CLI_NOMBRE)
        Else
           Me.txt_nombre = ""
        End If
        rsaux.Close
        Me.txt_fecha = Format(IIf(IsNull(rs!DTIM_car_FECHA), "", rs!DTIM_car_FECHA), "short Date")
        Me.txt_importe = Format(IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto) / IIf(IsNull(rs!FLOA_cAR_TIPO_cAMBIO), 1, rs!FLOA_cAR_TIPO_cAMBIO), "###,###,##0.00")
        Me.txt_plazo = IIf(IsNull(rs!INTE_CAR_PLAZO), 0, rs!INTE_CAR_PLAZO)
     Else
        MsgBox "El documento no existe", vbOKOnly, "ATENCION"
     End If
     rs.Close
   Else
      MsgBox "Número de docuemnto incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_plazo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_aplicar.SetFocus
   End If
End Sub

Private Sub txt_serie_Change()
   Me.txt_clave = ""
   Me.txt_fecha = ""
   Me.txt_importe = ""
   Me.txt_importe = ""
   Me.txt_nombre = ""
   Me.txt_numero = ""
   Me.txt_plazo = ""
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_tipo_Change()
   Me.txt_clave = ""
   Me.txt_fecha = ""
   Me.txt_importe = ""
   Me.txt_importe = ""
   Me.txt_nombre = ""
   Me.txt_numero = ""
   Me.txt_plazo = ""
   Me.txt_serie = ""
End Sub

Private Sub txt_tipo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_tipo_LostFocus()
   If Trim(Me.txt_tipo) = "FA" Or Trim(Me.txt_tipo) = "NG" Or Trim(Me.txt_tipo) = "NC" Or Trim(Me.txt_tipo) = "PA" Then
      If Me.txt_tipo = "FA" Then
         Me.txt_descripcion = "FACTURA"
      End If
      If Me.txt_tipo = "NG" Then
         Me.txt_descripcion = "NOTA DE CARGO"
      End If
      If Me.txt_tipo = "NC" Then
         Me.txt_descripcion = "NOTA DE CREDITO"
      End If
      If Me.txt_tipo = "PA" Then
         Me.txt_descripcion = "PAGO"
      End If
   Else
      Me.txt_tipo = ""
      Me.txt_descripcion = ""
      MsgBox "Documento incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub
