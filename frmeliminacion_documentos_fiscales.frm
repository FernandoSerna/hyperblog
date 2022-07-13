VERSION 5.00
Begin VB.Form frmeliminacion_documentos_fiscales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eliminación de documentos fiscales"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6240
      Picture         =   "frmeliminacion_documentos_fiscales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmeliminacion_documentos_fiscales.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Cancelar Alt + C"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   60
      Left            =   30
      TabIndex        =   14
      Top             =   330
      Width           =   6585
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos del documento "
      Height          =   1410
      Left            =   90
      TabIndex        =   7
      Top             =   1950
      Width           =   6495
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   225
         Width           =   4260
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   945
         Width           =   1350
      End
      Begin VB.TextBox txt_fecha 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   585
         Width           =   1350
      End
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   1350
      End
      Begin VB.Label lbl_moneda 
         Height          =   255
         Left            =   2175
         TabIndex        =   19
         Top             =   990
         Width           =   1305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   105
         TabIndex        =   10
         Top             =   1020
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   105
         TabIndex        =   9
         Top             =   630
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   270
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Documento "
      Height          =   1470
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   6510
      Begin VB.TextBox txt_nombre_documento 
         Height          =   315
         Left            =   2235
         TabIndex        =   17
         Top             =   330
         Width           =   4185
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   1650
         TabIndex        =   6
         Top             =   1020
         Width           =   1950
      End
      Begin VB.TextBox txt_serie 
         Height          =   315
         Left            =   1650
         TabIndex        =   5
         Top             =   675
         Width           =   765
      End
      Begin VB.TextBox txt_documento 
         Height          =   315
         Left            =   1650
         TabIndex        =   4
         Top             =   330
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   735
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de documento:"
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   390
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmeliminacion_documentos_fiscales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_cancelar_Click()
   If Trim(txt_documento) <> "" Then
      If txt_documento = "FA" Or txt_documento = "NC" Or txt_documento = "NG" Then
         If Trim(Me.txt_Serie) <> "" Then
            If Trim(Me.txt_numero) <> "" Then
               If IsNumeric(Me.txt_numero) Then
                  rs.Open "SELECT * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_tipo_documento = '" + Trim(txt_documento) + "' and vcha_Ser_serie_id = '" + Trim(Me.txt_Serie) + "' and inte_car_numero = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     Me.txt_fecha = Format(IIf(IsNull(rs!dtim_car_fecha), "", rs!dtim_car_fecha), "Short date")
                     Me.txt_importe = Format(IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto) / IIf(IsNull(rs!floa_Car_tipo_cambio), 1, rs!floa_Car_tipo_cambio), "###,###,##0.00")
                     Me.txt_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                     rsaux.Open "select * from tb_clientes where vcha_cli_clave_id = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        Me.txt_nombre_cliente = IIf(IsNull(rsaux!vcha_cli_nombre), "", rsaux!vcha_cli_nombre)
                     Else
                        Me.txt_nombre_cliente = ""
                     End If
                     rsaux.Close
                     var_estatus = IIf(IsNull(rs!CHAR_CAR_ESTATUS), "", rs!CHAR_CAR_ESTATUS)
                     var_documento = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
                     var_afectacion = IIf(IsNull(rs!char_car_afectacion), "", rs!char_car_afectacion)
                     If var_estatus = "C" Then
                        var_si = MsgBox("¿Desea eliminar el documento", vbYesNo, "ATENCION")
                        If var_si = 6 Then
                           var_si = MsgBox("Confirmar la eliminación del documento", vbYesNo, "ATENCION")
                           If var_si = 6 Then
                              If var_afectacion = "+" Then
                                 rsaux.Open "SELECT * FROM TB_ESTADO_CUENTA WHERE vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ecu_movimiento_cargo =  '" + Me.txt_documento + "' and vcha_ecu_serie_cargo = '" + Me.txt_Serie + "' and inte_ecu_numero_cargo = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                                 x = rsaux.RecordCount
                                 If x > 1 Then
                                    MsgBox "Aun no se han eliminado todos los abonos de este cargo", vbOKOnly, "ATENCION"
                                 Else
                                    cnn.BeginTrans
                                    rsaux1.Open "DELETE FROM TB_ENCABEZADO_CARTERA where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_CAR_DOCUMENTO = '" + var_documento + "' and vcha_Ser_serie_id = '" + Trim(Me.txt_Serie) + "' and inte_car_numero = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                                    rsaux1.Open "DELETE FROM TB_SALDOS WHERE vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_CAR_DOCUMENTO = '" + var_documento + "' and vcha_Ser_serie_id = '" + Trim(Me.txt_Serie) + "' and inte_car_numero = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                                    rsaux1.Open "DELETE FROM TB_ESTADO_CUENTA WHERE  vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_ECU_movimiento_CARGO = '" + var_documento + "' and vcha_ECU_serie_CARGO = '" + Trim(Me.txt_Serie) + "' and inte_ECU_numero_cARGO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                                    rsaux1.Open "INSERT INTO TB_BITACORA_ELIMINACION_DOCUMENTOS_FISCALES (VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, DTIM_CAR_FECHA, FLOA_CAR_IMPORTE_NETO, DTIM_CAR_FECHA_ELIMINACION, VCHA_AUD_USUARIO_ID, VCHA_AUD_MAQUINA) Values ( '" + var_empresa + "', '" + var_documento + "', '" + Me.txt_Serie + "', " + Me.txt_numero + ", '" + Me.txt_fecha + "', " + CStr(CDbl(Me.txt_importe)) + ",getdate(),'" + var_clave_usuario_global + "','" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
                                    cnn.CommitTrans
                                    MsgBox "Se a eliminado el documento", vbOKOnly, "ATENCION"
                                 End If
                                 rsaux.Close
                              End If
                              If var_afectacion = "-" Then
                                 cnn.BeginTrans
                                 rsaux1.Open "DELETE FROM TB_ENCABEZADO_CARTERA where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_CAR_DOCUMENTO = '" + var_documento + "' and vcha_Ser_serie_id = '" + Trim(Me.txt_Serie) + "' and inte_car_numero = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                                 rsaux1.Open "DELETE FROM TB_ESTADO_CUENTA WHERE  vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_ECU_movimiento_ABONO = '" + var_documento + "' and vcha_ECU_serie_ABONO = '" + Trim(Me.txt_Serie) + "' and inte_ECU_numero_ABONO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                                 rsaux1.Open "INSERT INTO TB_BITACORA_ELIMINACION_DOCUMENTOS_FISCALES (VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, DTIM_CAR_FECHA, FLOA_CAR_IMPORTE_NETO, DTIM_CAR_FECHA_ELIMINACION, VCHA_AUD_USUARIO_ID, VCHA_AUD_MAQUINA) Values ( '" + var_empresa + "', '" + var_documento + "', '" + Me.txt_Serie + "', " + Me.txt_numero + ", '" + Me.txt_fecha + "', " + CStr(CDbl(Me.txt_importe)) + ",getdate(),'" + var_clave_usuario_global + "','" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
                                 cnn.CommitTrans
                                 MsgBox "Se a eliminado el documento", vbOKOnly, "ATENCION"
                              End If
                           End If
                        End If
                     Else
                        MsgBox "El documento no a sido cancelado", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "El documento no existe", vbOKOnly, "ATENCION"
                  End If
                  rs.Close
               Else
                  MsgBox "Número de documento incorrecto", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No se a indicado un número de documento", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se a seleccionado una serie", vbOKOnly, "ATENCION"
         End If
      Else
         txt_documento = ""
         MsgBox "Documento incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un documento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 2000
   Left = 2000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub txt_documento_Change()
   Me.txt_cliente = ""
   Me.txt_fecha = ""
   Me.txt_importe = ""
   Me.txt_numero = ""
   Me.txt_Serie = ""
   Me.txt_nombre_cliente = ""
   Me.txt_nombre_documento = ""
   Me.lbl_moneda = ""
End Sub

Private Sub txt_documento_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_documento_LostFocus()
   If Trim(txt_documento) <> "" Then
      If txt_documento = "FA" Or txt_documento = "NC" Or txt_documento = "NG" Then
         If txt_documento = "FA" Then
            Me.txt_nombre_documento = "FACTURA"
         End If
         If Me.txt_documento = "NC" Then
            Me.txt_nombre_documento = "NOTA DE CREDITO"
         End If
         If Me.txt_documento = "NG" Then
            Me.txt_nombre_documento = "NOTA DE CARGO"
         End If
      Else
         txt_documento = ""
         MsgBox "Documento incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_nombre_documento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   
   If KeyAscii = 13 Then
      Me.cmd_cancelar.SetFocus
   End If
End Sub

Private Sub txt_numero_LostFocus()
   If Trim(txt_numero) <> "" Then
      If IsNumeric(Me.txt_numero) Then
         If Trim(txt_documento) <> "" Then
            If Trim(Me.txt_Serie) <> "" Then
               rs.Open "Select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_tipo_documento = '" + Trim(Me.txt_documento) + "' and vcha_ser_serie_id = '" + Trim(Me.txt_Serie) + "' and inte_car_numero = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  Me.txt_fecha = Format(IIf(IsNull(rs!dtim_car_fecha), "", rs!dtim_car_fecha), "Short date")
                  If rs!floa_Car_tipo_cambio = 0 Then
                     Me.txt_importe = Format(IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto) / 1, "###,###,##0.00")
                  Else
                     Me.txt_importe = Format(IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto) / IIf(IsNull(rs!floa_Car_tipo_cambio), 1, rs!floa_Car_tipo_cambio), "###,###,##0.00")
                  End If
                  Me.txt_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                  rsaux.Open "select * from tb_clientes where vcha_cli_clave_id = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     Me.txt_nombre_cliente = IIf(IsNull(rsaux!vcha_cli_nombre), "", rsaux!vcha_cli_nombre)
                  Else
                     Me.txt_nombre_cliente = ""
                  End If
                  rsaux.Close
               Else
                  MsgBox "El documento no existe", vbOKOnly, "ATENCION"
               End If
               rs.Close
            Else
               MsgBox "No se a seleccionado una serie", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se a seleccionado un documento", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Número de documento incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_serie_Change()
   Me.txt_numero = ""
   Me.txt_cliente = ""
   Me.txt_nombre_cliente = ""
   Me.txt_fecha = ""
   Me.txt_importe = ""
End Sub

Private Sub txt_Serie_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub
