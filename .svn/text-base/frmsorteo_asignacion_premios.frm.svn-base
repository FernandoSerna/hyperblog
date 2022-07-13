VERSION 5.00
Begin VB.Form frmsorteo_asignacion_premios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Boletos premiados"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   " Saldo del cliente "
      Height          =   615
      Left            =   90
      TabIndex        =   14
      Top             =   3855
      Width           =   6960
      Begin VB.Label lbl_despues 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4755
         TabIndex        =   18
         Top             =   165
         Width           =   1920
      End
      Begin VB.Label lbl_antes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1500
         TabIndex        =   17
         Top             =   165
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Actual:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3720
         TabIndex        =   16
         Top             =   165
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anterior:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         TabIndex        =   15
         Top             =   165
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmd_eliminar_boleto 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmsorteo_asignacion_premios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Eliminar"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6690
      Picture         =   "frmsorteo_asignacion_premios.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar_boleto 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmsorteo_asignacion_premios.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   120
      Left            =   75
      TabIndex        =   10
      Top             =   300
      Width           =   6945
   End
   Begin VB.Frame Frame2 
      Caption         =   " Premio "
      Height          =   1650
      Left            =   90
      TabIndex        =   8
      Top             =   2175
      Width           =   6930
      Begin VB.OptionButton opt_reintegro 
         Caption         =   "Reintegro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   255
         TabIndex        =   5
         Top             =   885
         Width           =   2925
      End
      Begin VB.OptionButton opt_400 
         Caption         =   "400 pesos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   255
         TabIndex        =   4
         Top             =   300
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos del Boleto "
      Height          =   1665
      Left            =   90
      TabIndex        =   6
      Top             =   465
      Width           =   6930
      Begin VB.TextBox txt_codigo 
         Height          =   360
         Left            =   1080
         TabIndex        =   3
         Top             =   1185
         Width           =   1875
      End
      Begin VB.TextBox txt_boleto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   0
         Top             =   270
         Width           =   1875
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   350
         Left            =   2415
         TabIndex        =   2
         Top             =   810
         Width           =   4410
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   350
         Left            =   1080
         TabIndex        =   1
         Top             =   810
         Width           =   1290
      End
      Begin VB.Label Label4 
         Caption         =   "Código:"
         Height          =   270
         Left            =   225
         TabIndex        =   19
         Top             =   1290
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Boleto:"
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clliente:"
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   885
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmsorteo_asignacion_premios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_almacen As String
Dim var_unidad_org As String
Dim var_cliente As String
Dim var_numero As Double
Dim var_movimiento As String


Private Sub cmd_guardar_boleto_Click()
   VAR_GLOBAL_ACCESO_SORTEO = 0
   Dim var_acceso_sorteo As Boolean
   Dim var_codigo_leido As Integer
   frmacceso_boletos.Show 1
   If Trim(Me.txt_codigo) <> "" Then
      If rsaux3.State = 1 Then
         rsaux3.Close
      End If
      rsaux3.Open "SELECT * FROM TB_SORTEO_CODIGOS_PREMIOS WHERE VCHA_SOR_CODIGO = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux3.EOF Then
         var_codigo_leido = IIf(IsNull(rsaux3!inte_sor_codigo_leido), 0, rsaux3!inte_sor_codigo_leido)
         If var_codigo_leido = 0 Then
            var_acceso_sorteo = False
            VAR_TIPO_PREMIO = Trim(rsaux3!VCHA_SOR_PREMIO)
            If VAR_TIPO_PREMIO = "01" And Me.opt_400.Value = True Then
               var_acceso_sorteo = True
            Else
               If VAR_TIPO_PREMIO = "02" And Me.opt_reintegro.Value = True Then
                  var_acceso_sorteo = True
               Else
                  var_acceso_sorteo = False
               End If
            End If
            If var_acceso_sorteo = True Then
               If VAR_GLOBAL_ACCESO_SORTEO = 1 Then
                  If Trim(Me.txt_boleto) <> "" Then
                     If Trim(Me.txt_clave_cliente) <> "" Then
                        rsaux2.Open "select * from TB_SORTEO_BOLETOS_PREMIO where inte_sor_boleto = " + Me.txt_boleto, cnn, adOpenDynamic, adLockOptimistic
                        If rsaux2.EOF Then
                           var_si = MsgBox("¿Desea asignar el boleto", vbYesNo, "ATENCION")
                           If var_si = 6 Then
                              var_si = MsgBox("Confirmar la asignación del boleto", vbYesNo, "ATENCION")
                              If var_si = 6 Then
                                 If Me.opt_reintegro = True Then
                                     rs.Open "select * from tb_sorteo_folios", cnn, adOpenDynamic, adLockOptimistic
                                     var_numero_Actual = IIf(IsNull(rs!inte_sor_folio_actual), 0, rs!inte_sor_folio_actual)
                                     rs.Close
                                     var_si = MsgBox("¿El boleto siguiente es el " + CStr(var_numero_Actual) + "?", vbYesNo, "ATENCION")
                                     If var_si = 6 Then
                                        rs.Open "UPDATE TB_SORTEO_FOLIOS SET INTE_SOR_FOLIO_aCTUAL = INTE_SOR_FOLIO_ACTUAL + 1"
                                        var_cadena = "INSERT INTO TB_SORTEO_BOLETOS_PREMIO (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, VCHA_CLI_CLAVE_ID, INTE_SOR_BOLETO, INTE_SOR_PREMIO, INTE_SOR_REINTEGRO) VALUES"
                                        var_cadena = var_cadena + "('" + var_empresa + "','" + var_unidad_org + "','" + var_almacen + "', '" + var_movimiento + "', " + CStr(var_numero) + ",'" + Me.txt_clave_cliente + "', " + Me.txt_boleto + ",0,1)"
                                        rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                        var_cadena = "insert into tb_sorteo_boletos_movimiento (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_emo_numero, floa_sor_importe, inte_sor_numero_boletos, inte_sor_boleto_inicio, inte_sor_boleto_final, inte_sor_reintegro) values"
                                        var_cadena = var_cadena + "('" + var_empresa + "', '" + var_unidad_org + "', '" + var_almacen + "', '" + var_movimiento + "', " + CStr(var_numero) + ",0,1," + CStr(var_numero_Actual) + "," + CStr(var_numero_Actual) + ",1)"
                                        rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                        var_cadena = "call PC_BOLETOS_UPD (" + Me.txt_boleto + ", 2, 3,'E000001666', " + CStr(var_numero) + ",'" + Me.txt_clave_cliente + "', '" + Me.txt_nombre_cliente + "',0,0,0," + Me.txt_codigo + ",SYSDATE)"
                                        rs.Open var_cadena, cnnsorteo, adOpenDynamic, adLockOptimistic
                                        rs.Open "update TB_SORTEO_CODIGOS_PREMIOS set inte_sor_codigo_leido = 1 where vcha_sor_codigo = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                     End If
                                 End If
                                 If Me.opt_400 = True Then
                                    var_cadena = "INSERT INTO TB_SORTEO_BOLETOS_PREMIO (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, VCHA_CLI_CLAVE_ID, INTE_SOR_BOLETO, INTE_SOR_PREMIO, INTE_SOR_REINTEGRO) VALUES"
                                    var_cadena = var_cadena + "('" + var_empresa + "','" + var_unidad_org + "','" + var_almacen + "', '" + var_movimiento + "', " + CStr(var_numero) + ",'" + Me.txt_clave_cliente + "', " + Me.txt_boleto + ",1,0)"
                                    rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    var_cadena = "call PC_BOLETOS_UPD (" + Me.txt_boleto + ", 2, 2,'E000001666', " + CStr(var_numero) + ",'" + Me.txt_clave_cliente + "', '" + Me.txt_nombre_cliente + "',0,0,0," + Me.txt_codigo + ",SYSDATE)"
                                    rs.Open var_cadena, cnnsorteo, adOpenDynamic, adLockOptimistic
                                    rs.Open "select vcha_cli_referencia from tb_clientes where vcha_cli_clave_id = '" + Me.txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                                    var_referencia = Trim(rs!vcha_cli_referencia)
                                    rs.Close
                                    rs.Open "CALL SP_AGREGA_ABONO('" + var_referencia + "',400, 400,SYSDATE,SYSDATE,'','','PS','Premio de sorteo')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                                    rsaux4.Open "select NUMB_SAL_IMPORTE_disponible from tb_saldo where vcha_sal_referencia = '" + var_referencia + "'", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                                    If Not rsaux4.EOF Then
                                       Me.lbl_despues = Format(rsaux4(0).Value, "###,###,##0.00")
                                    Else
                                      Me.lbl_despues = "0.00"
                                    End If
                                    rs.Open "update TB_SORTEO_CODIGOS_PREMIOS set inte_sor_codigo_leido = 1 where vcha_sor_codigo = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                    rsaux4.Close
                                 End If
                              End If
                           End If
                        Else
                           MsgBox "El boleto ya fue asignado", vbOKOnly, "ATENCION"
                        End If
                        rsaux2.Close
                     Else
                        MsgBox "El boleto no tiene un cliente asociado", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "Debe de indicar un número de boleto", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "Imposible asignar boletos", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El boleto no corresponde al premio seleccionado"
            End If
         Else
            MsgBox "El codigo ya fue asignado en otro boleto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El boleto no tiene premio", vbOKOnly, "ATENCION"
         rsaux3.Close
      End If
   Else
      MsgBox "Debe de indicar el código del boleto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 1800
   Left = 2200
   Me.opt_400.Value = True
   If cnnsorteo.State = 0 Then
      cnnsorteo.Open var_conexion_sorteo
      cnnsorteo.CursorLocation = adUseClient
   End If
   
   If cnn_clientes_tiendas.State = 0 Then
      cnn_clientes_tiendas.Open var_conexion_pedidos_tiendas
      cnn_clientes_tiendas.CursorLocation = adUseClient
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_boleto_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_boleto_LostFocus()
   If Trim(Me.txt_boleto) <> "" Then
      If IsNumeric(Me.txt_boleto) Then
         rs.Open "select * from tb_sorteo_boletos_movimiento where " + Me.txt_boleto + " >= inte_sor_boleto_inicio and " + Me.txt_boleto + "<= inte_sor_boleto_final", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_unidad_org = IIf(IsNull(rs!vcha_uor_unidad_id), "", rs!vcha_uor_unidad_id)
            var_almacen = IIf(IsNull(rs!vcha_alm_almacen_id), "", rs!vcha_alm_almacen_id)
            var_movimiento = IIf(IsNull(rs!VCHA_MOV_MOVIMIENTO_ID), 0, rs!VCHA_MOV_MOVIMIENTO_ID)
            var_numero = IIf(IsNull(rs!INTE_EMO_NUMERO), 0, rs!INTE_EMO_NUMERO)
            rsaux1.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_org + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + var_movimiento + "' and inte_emo_numero = " + CStr(var_numero), cnn, adOpenDynamic, adLockOptimistic
            Me.txt_clave_cliente = IIf(IsNull(rsaux1!vcha_cli_clave_id), "", rsaux1!vcha_cli_clave_id)
            rsaux1.Close
            rsaux2.Open "select vcha_cli_nombre, vcha_cli_referencia from tb_clientes where vcha_Cli_clave_id = '" + Me.txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
            Me.txt_nombre_cliente = IIf(IsNull(rsaux2!vcha_cli_nombre), "", rsaux2!vcha_cli_nombre)
            var_referencia = Trim(rsaux2!vcha_cli_referencia)
            rsaux2.Close
            rsaux2.Open "select NUMB_SAL_IMPORTE_disponible from tb_saldo where vcha_sal_referencia = '" + var_referencia + "'", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               Me.lbl_antes = Format(rsaux2(0).Value, "###,###,##0.00")
               Me.lbl_despues = Format(rsaux2(0).Value, "###,###,##0.00")
            Else
               Me.lbl_antes = "0.00"
               Me.lbl_despues = "0.00"
            End If
            rsaux2.Close
         Else
            var_unidad_org = ""
            var_almacen = ""
            var_movimiento = ""
            var_numero = 0
            Me.txt_clave_cliente = ""
            Me.txt_nombre_cliente = ""
            MsgBox "El boleto no a sido asignado", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Boleto incorrecto", vbOKOnly, "ATENCION"
         var_unidad_org = ""
         var_almacen = ""
         var_numero = 0
         var_movimiento = ""
         Me.txt_clave_cliente = ""
         Me.txt_nombre_cliente = ""
      End If
   Else
      var_unidad_org = ""
      var_almacen = ""
      var_movimiento = ""
      Me.txt_clave_cliente = ""
      Me.txt_nombre_cliente = ""
   End If
End Sub

Private Sub txt_clave_cliente_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub
