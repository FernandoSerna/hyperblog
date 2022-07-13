VERSION 5.00
Begin VB.Form frmcancelacion_documentos_electronicos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelación de documentos electrónicos"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6225
      Picture         =   "frmcancelacion_documentos_electronicos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmcancelacion_documentos_electronicos.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   105
      TabIndex        =   19
      Top             =   315
      Width           =   6435
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos del documento "
      Height          =   1155
      Left            =   135
      TabIndex        =   9
      Top             =   1320
      Width           =   6390
      Begin VB.TextBox txt_estatus 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5805
         TabIndex        =   18
         Top             =   645
         Width           =   480
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   390
         Left            =   2955
         TabIndex        =   16
         Top             =   645
         Width           =   2205
      End
      Begin VB.TextBox txt_fecha 
         Enabled         =   0   'False
         Height          =   390
         Left            =   945
         TabIndex        =   14
         Top             =   645
         Width           =   1260
      End
      Begin VB.TextBox txt_nombre_cliente 
         Enabled         =   0   'False
         Height          =   345
         Left            =   2220
         TabIndex        =   12
         Top             =   270
         Width           =   4065
      End
      Begin VB.TextBox txt_cliente 
         Enabled         =   0   'False
         Height          =   345
         Left            =   945
         TabIndex        =   11
         Top             =   270
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Estatus:"
         Height          =   195
         Left            =   5220
         TabIndex        =   17
         Top             =   750
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   2310
         TabIndex        =   15
         Top             =   750
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   210
         TabIndex        =   13
         Top             =   750
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   345
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Documento "
      Height          =   825
      Left            =   135
      TabIndex        =   5
      Top             =   435
      Width           =   6390
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   4890
         TabIndex        =   2
         Top             =   330
         Width           =   1410
      End
      Begin VB.TextBox txt_serie 
         Height          =   315
         Left            =   3600
         TabIndex        =   1
         Top             =   330
         Width           =   450
      End
      Begin VB.ComboBox cmb_tipo 
         Height          =   315
         ItemData        =   "frmcancelacion_documentos_electronicos.frx":0784
         Left            =   1080
         List            =   "frmcancelacion_documentos_electronicos.frx":078E
         TabIndex        =   0
         Text            =   "Factura"
         Top             =   330
         Width           =   1590
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   4215
         TabIndex        =   8
         Top             =   390
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   2820
         TabIndex        =   7
         Top             =   390
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   225
         TabIndex        =   6
         Top             =   390
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmcancelacion_documentos_electronicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_tipo_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub cmd_aceptar_pedidos_Click()
   var_documento = Me.cmb_tipo.Text
   If Trim(Me.cmb_tipo.Text) <> "" Then
      If UCase(var_documento) = "FACTURA" Then
         If IsNumeric(Me.txt_numero) Then
            rs.Open "SELECT * FROM TB_ENCABEZADO_CARTERA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_CAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + Me.txt_serie + "' AND INTE_CAR_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux1.Open "select * from tb_clientes where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  var_email = IIf(IsNull(rsaux1!vcha_cli_email), "", rsaux1!vcha_cli_email)
               End If
               rsaux1.Close
               var_estatus = IIf(IsNull(rs!CHAR_CAR_ESTATUS), "I", rs!CHAR_CAR_ESTATUS)
               If var_estatus = "I" Or var_estatus = "C" Then
                  var_si = 6
                  If var_si = 6 Then
                  'If Year(Date) = Year(rs!dtim_Car_fecha) And Month(Date) = Month(rs!dtim_Car_fecha) And Day(Date) = Day(rs!dtim_Car_fecha) Then
                     var_si = MsgBox("¿Desea cancelar la factura " + Trim(Me.txt_serie) + Trim(Me.txt_numero) + "?", vbYesNo, "ATENCION")
                     If var_si = 6 Then
                        var_si = MsgBox("Confirmar la cancelación de la factura " + Trim(Me.txt_serie) + Trim(Me.txt_numero) + "?", vbYesNo, "ATENCION")
                        If var_si = 6 Then
                            var_cadena = "SELECT     dbo.TB_ESTADO_CUENTA.VCHA_EMP_EMPRESA_ID, dbo.TB_ESTADO_CUENTA.VCHA_ECU_SERIE_CARGO, dbo.TB_ESTADO_CUENTA.VCHA_ECU_MOVIMIENTO_CARGO, dbo.TB_ESTADO_CUENTA.INTE_ECU_NUMERO_CARGO, dbo.TB_ESTADO_CUENTA.VCHA_ECU_SERIE_ABONO, dbo.TB_ESTADO_CUENTA.VCHA_ECU_MOVIMIENTO_ABONO, dbo.TB_ESTADO_CUENTA.INTE_ECU_NUMERO_ABONO, dbo.TB_ESTADO_CUENTA.FLOA_ECU_IMPORTE_CARGO, dbo.TB_ESTADO_CUENTA.FLOA_ECU_IMPORTE_ABONO , dbo.TB_ESTADO_CUENTA.CHAR_ECU_ESTATUS FROM dbo.TB_ENCABEZADO_CARTERA INNER JOIN dbo.TB_ESTADO_CUENTA ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_ESTADO_CUENTA.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = dbo.TB_ESTADO_CUENTA.VCHA_ECU_SERIE_ABONO AND dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = dbo.TB_ESTADO_CUENTA.VCHA_ECU_MOVIMIENTO_ABONO AND dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = dbo.TB_ESTADO_CUENTA.INTE_ECU_NUMERO_ABONO "
                            var_cadena = var_cadena + " WHERE (dbo.TB_ESTADO_CUENTA.INTE_ECU_NUMERO_CARGO = " + Me.txt_numero + ") AND (dbo.TB_ESTADO_CUENTA.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ESTADO_CUENTA.VCHA_ECU_SERIE_CARGO = '" + Me.txt_serie + "') AND (dbo.TB_ESTADO_CUENTA.VCHA_ECU_MOVIMIENTO_CARGO = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C') OR (dbo.TB_ESTADO_CUENTA.INTE_ECU_NUMERO_CARGO = " + Me.txt_numero + ") AND (dbo.TB_ESTADO_CUENTA.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ESTADO_CUENTA.VCHA_ECU_SERIE_CARGO = '" + Me.txt_serie + "') AND (dbo.TB_ESTADO_CUENTA.VCHA_ECU_MOVIMIENTO_CARGO = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL)                             "
                            If rsaux.State = 1 Then
                               rsaux.Close
                            End If
                            rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                            If rsaux.EOF Then
                               If var_estatus = "I" Then
                                  rsaux3.Open "UPDATE TB_ENCABEZADO_CARTERA SET CHAR_CAR_ESTATUS = 'C' WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_CAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + Me.txt_serie + "' AND INTE_CAR_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                               End If
                               Open (var_ruta_documentos_electronicos & "\" + Trim(Trim(Me.txt_serie)) + Trim(Str(Me.txt_numero)) + ".fci") For Output As #1
                               var_cadena = "<Factura>" + Chr(13) + "serie=" + Trim(Me.txt_serie) + Chr(13) + "folio=" + Me.txt_numero + Chr(13) + "email=" + var_email + Chr(13) + "</Factura>"
                               Print #1, var_cadena
                               Close #1
                               
                               Set fs = CreateObject("Scripting.FileSystemObject")
                               ArchivoOrigen = var_ruta_documentos_electronicos & "\" + Trim(Me.txt_serie) + Trim(Str(Me.txt_numero)) + ".fci"
                               ArchivoDestino = var_ruta_documentos_electronicos & "\" + Trim(Me.txt_serie) + Trim(Str(Me.txt_numero)) + ".ffc"
                               fs.CopyFile ArchivoOrigen, ArchivoDestino
                               fs.DeleteFile ArchivoOrigen
                               
                               
                               
                            Else
                               MsgBox "La factura ya no puede ser cancelada ya que contiene abonos", vbOKOnly, "ATENCION"
                            End If
                            rsaux.Close
                        End If
                     End If
                     
                  Else
                     MsgBox "La factura ya no puede ser cancelada ya que corresponde a otro dia", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "La factura ya fue cancelada con anterioridad", vbOKOnly, "ATENCION"
               End If
            End If
            rs.Close
         Else
            MsgBox "Número de factura incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         If IsNumeric(Me.txt_numero) Then
            rs.Open "SELECT * FROM TB_ENCABEZADO_CARTERA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND vcha_car_tipo_documento = 'NC' AND VCHA_SER_SERIE_ID = '" + Me.txt_serie + "' AND INTE_CAR_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux1.Open "select * from tb_clientes where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  var_email = IIf(IsNull(rsaux1!vcha_cli_email), "", rsaux1!vcha_cli_email)
               End If
               rsaux1.Close
               var_estatus = Trim(IIf(IsNull(rs!CHAR_CAR_ESTATUS), "I", rs!CHAR_CAR_ESTATUS))
               If Trim(var_estatus) = "" Then
                  var_estatus = "I"
               End If
               If var_estatus = "I" Then
                  If Year(Date) = Year(rs!dtim_Car_fecha) And Month(Date) = Month(rs!dtim_Car_fecha) And Day(Date) = Day(rs!dtim_Car_fecha) Then
                     var_si = MsgBox("¿Desea cancelar la nota de crédito " + Trim(Me.txt_serie) + Trim(Me.txt_numero) + "?", vbYesNo, "ATENCION")
                     If var_si = 6 Then
                        var_si = MsgBox("Confirmar la cancelación de la nota de crédito " + Trim(Me.txt_serie) + Trim(Me.txt_numero) + "?", vbYesNo, "ATENCION")
                        If var_si = 6 Then
                           rsaux3.Open "UPDATE TB_ENCABEZADO_CARTERA SET CHAR_CAR_ESTATUS = 'C' WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND vcha_car_tipo_documento = 'NC' AND VCHA_SER_SERIE_ID = '" + Me.txt_serie + "' AND INTE_CAR_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                           Open (var_ruta_documentos_electronicos & "\cancela_nc" + Trim(Trim(Me.txt_serie)) + Trim(Str(Me.txt_numero)) + ".fi") For Output As #1
                           var_cadena = "<Factura>" + Chr(13) + "serie=" + Trim(Me.txt_serie) + Chr(13) + "folio=" + Me.txt_numero + Chr(13) + "email=" + var_email + Chr(13) + "</Factura>"
                           Print #1, var_cadena
                           Close #1
                        End If
                     End If
                  Else
                     MsgBox "La nota de crédito ya no puede ser cancelada ya que corresponde a otro dia", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "La nota de crédito ya fue cancelada con anterioridad", vbOKOnly, "ATENCION"
               End If
            End If
            rs.Close
         Else
            MsgBox "Número de la nota de crédito incorrecto", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "No se a seleccionado un documento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 2500
   Left = 2500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub txt_numero_Change()
   Me.txt_cliente = ""
   Me.txt_nombre_cliente = ""
   Me.txt_importe = ""
   Me.txt_estatus = ""
   Me.txt_fecha = ""
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_LostFocus()
   Dim var_tipo_documento As String
   If Me.cmb_tipo = "Factura" Then
      var_tipo_documento = "FA"
   Else
      var_tipo_documento = "NC"
   End If
   If IsNumeric(Me.txt_numero) Then
      rs.Open "select * from tb_encabezado_Cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_tipo_documento = '" + var_tipo_documento + "' and vcha_Ser_serie_id = '" + Me.txt_serie + "' and inte_car_numero = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         rsaux1.Open "select * from tb_clientes where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            Me.txt_cliente = rs!vcha_cli_clave_id
            Me.txt_nombre_cliente = rsaux1!VCHA_CLI_NOMBRE
            Me.txt_importe = Format(IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,###,##0.00")
            Me.txt_estatus = IIf(IsNull(rs!CHAR_CAR_ESTATUS), "I", rs!CHAR_CAR_ESTATUS)
            Me.txt_fecha = rs!dtim_Car_fecha
         Else
            Me.txt_fecha = ""
            Me.txt_cliente = ""
            Me.txt_nombre_cliente = ""
            Me.txt_importe = ""
            Me.txt_estatus = ""
            MsgBox "El documento es incorrecto", vbOKOnly, "ATENCION"
         End If
         rsaux1.Close
      Else
         Me.txt_fecha = ""
         Me.txt_cliente = ""
         Me.txt_nombre_cliente = ""
         Me.txt_estatus = ""
         Me.txt_importe = ""
         MsgBox "El documento no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      MsgBox "Número de documento incorrecto", vbOKOnly
   End If
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub
