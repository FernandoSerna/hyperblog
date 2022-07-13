VERSION 5.00
Begin VB.Form frmcorreccion_saldos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Correccion de saldos"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_aceptar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmcorreccion_saldos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Aceptar"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmcorreccion_saldos.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4545
      Picture         =   "frmcorreccion_saldos.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   " Saldo "
      Height          =   705
      Left            =   90
      TabIndex        =   14
      Top             =   1500
      Width           =   4830
      Begin VB.TextBox txt_saldo_real 
         Height          =   345
         Left            =   3000
         TabIndex        =   8
         Top             =   270
         Width           =   1695
      End
      Begin VB.TextBox txt_saldo_actual 
         Height          =   345
         Left            =   720
         TabIndex        =   7
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Real:"
         Height          =   195
         Left            =   2565
         TabIndex        =   16
         Top             =   345
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Actual:"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   345
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   15
      TabIndex        =   10
      Top             =   345
      Width           =   4950
   End
   Begin VB.Frame nada 
      Caption         =   " Documento "
      Height          =   870
      Left            =   90
      TabIndex        =   9
      Top             =   570
      Width           =   4815
      Begin VB.CommandButton cmd_buscar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4365
         Picture         =   "frmcorreccion_saldos.frx":0886
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Buscar"
         Top             =   345
         Width           =   330
      End
      Begin VB.TextBox txt_numero 
         Height          =   360
         Left            =   3150
         TabIndex        =   5
         Top             =   322
         Width           =   1140
      End
      Begin VB.TextBox txt_serie 
         Height          =   360
         Left            =   1740
         TabIndex        =   4
         Top             =   322
         Width           =   615
      End
      Begin VB.TextBox txt_tipo 
         Height          =   360
         Left            =   630
         TabIndex        =   3
         Top             =   322
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   2445
         TabIndex        =   13
         Top             =   405
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   1290
         TabIndex        =   12
         Top             =   405
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   405
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmcorreccion_saldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_Click()
   If Me.txt_tipo <> "" Then
      If IsNumeric(Me.txt_numero) Then
         If Me.txt_serie <> "" Then
            rs.Open "SELECT * FROM TB_sALDOS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_cAR_DOCUMENTO = '" + Me.txt_tipo + "' AND VCHA_sER_SERIE_ID = '" + Me.txt_serie + "' AND INTE_cAR_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_saldo_Actual = Round(IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE), 2)
               If CDbl(Me.txt_saldo_actual) = CDbl(var_saldo_Actual) Then
                  'MsgBox "UPDATE TB_sALDOS SET FLOA_SAL_IMPORTE = " + CStr(CDbl(Me.txt_saldo_real)) + " WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_cAR_DOCUMENTO = '" + Me.txt_tipo + "' AND VCHA_sER_SERIE_ID = '" + Me.txt_serie + "' AND INTE_cAR_NUMERO = " + Me.txt_numero
                  rsaux.Open "UPDATE TB_sALDOS SET FLOA_SAL_IMPORTE = " + CStr(CDbl(Me.txt_saldo_real)) + " WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_cAR_DOCUMENTO = '" + Me.txt_tipo + "' AND VCHA_sER_SERIE_ID = '" + Me.txt_serie + "' AND INTE_cAR_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                  ' rsaux.Open "UPDATE TB_sALDOS SET FLOA_SAL_IMPORTE = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_cAR_DOCUMENTO = '" + Me.txt_tipo + "' AND VCHA_sER_SERIE_ID = '" + Me.txt_serie + "' AND INTE_cAR_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                  MsgBox "El saldo actual se cambio con exito", vbOKOnly, "ATENCION"
               Else
                  MsgBox "El importe del saldo actual ya no puede ser cambiado", vbOKOnly, "ATENCION"
               End If
               
            Else
               MsgBox "El documento no existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "Se debe de indicar una serie", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Se debe de indicar un numero", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Se debe de indicar un tipo de documento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_buscar_Click()
   If Trim(Me.txt_tipo) <> "" Then
      If Trim(Me.txt_serie) <> "" Then
         If IsNumeric(Me.txt_numero) Then
            var_cadena = "SELECT dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID, dbo.TB_SALDOS.VCHA_SER_SERIE_ID, dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO,  dbo.TB_SALDOS.INTE_CAR_NUMERO, dbo.TB_SALDOS.FLOA_SAL_IMPORTE, dbo.TB_ESTADO_CUENTA.FLOA_ECU_IMPORTE_ABONO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO AS IMPORTE_ABONOS, TB_ENCABEZADO_CARTERA_1.FLOA_CAR_IMPORTE_NETO / TB_ENCABEZADO_CARTERA_1.FLOA_CAR_TIPO_CAMBIO AS IMPORTE_CARGO FROM dbo.TB_SALDOS INNER JOIN dbo.TB_ESTADO_CUENTA ON dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ESTADO_CUENTA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALDOS.VCHA_SER_SERIE_ID = dbo.TB_ESTADO_CUENTA.VCHA_ECU_SERIE_CARGO AND dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = dbo.TB_ESTADO_CUENTA.VCHA_ECU_MOVIMIENTO_CARGO AND dbo.TB_SALDOS.INTE_CAR_NUMERO = dbo.TB_ESTADO_CUENTA.INTE_ECU_NUMERO_CARGO INNER JOIN dbo.TB_ENCABEZADO_CARTERA AS TB_ENCABEZADO_CARTERA_1 ON dbo.TB_ESTADO_CUENTA.VCHA_EMP_EMPRESA_ID = TB_ENCABEZADO_CARTERA_1.VCHA_EMP_EMPRESA_ID AND "
            var_cadena = var_cadena + " dbo.TB_ESTADO_CUENTA.VCHA_ECU_SERIE_CARGO = TB_ENCABEZADO_CARTERA_1.VCHA_SER_SERIE_ID AND dbo.TB_ESTADO_CUENTA.VCHA_ECU_MOVIMIENTO_CARGO = TB_ENCABEZADO_CARTERA_1.VCHA_CAR_DOCUMENTO AND dbo.TB_ESTADO_CUENTA.INTE_ECU_NUMERO_CARGO = TB_ENCABEZADO_CARTERA_1.INTE_CAR_NUMERO LEFT OUTER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_ESTADO_CUENTA.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_ESTADO_CUENTA.VCHA_ECU_SERIE_ABONO = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_ESTADO_CUENTA.VCHA_ECU_MOVIMIENTO_ABONO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_ESTADO_CUENTA.INTE_ECU_NUMERO_ABONO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO WHERE (dbo.TB_SALDOS.INTE_CAR_NUMERO = " + Me.txt_numero + ") AND (dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = '" + Me.txt_tipo + "') AND (dbo.TB_SALDOS.VCHA_SER_SERIE_ID = '" + Me.txt_serie + "') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C') OR "
            var_cadena = var_cadena + " (dbo.TB_SALDOS.INTE_CAR_NUMERO = " + Me.txt_numero + ") AND (dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = '" + Me.txt_tipo + "') AND (dbo.TB_SALDOS.VCHA_SER_SERIE_ID = '" + Me.txt_serie + "') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL) "
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               
               var_importe_Cargo = IIf(IsNull(rs!importe_Cargo), 0, rs!importe_Cargo)
               var_saldo_Actual = Round(IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE), 2)
               Me.txt_saldo_actual = Format(var_saldo_Actual, "###,###,##0.00")
               var_abonos = 0
               While Not rs.EOF
                     var_abonos = var_abonos + IIf(IsNull(rs!importe_abonos), 0, rs!importe_abonos)
                     rs.MoveNext
               Wend
               Me.txt_saldo_real = Format(var_importe_Cargo - var_abonos, "###,###,##0.00")
               
            Else
               MsgBox "El documento no existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "Número de documento incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Se debe de indicar una serie", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Se debe de indicar un tipo de documento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_numero = ""
   Me.txt_saldo_actual = ""
   Me.txt_saldo_real = ""
   Me.txt_serie = ""
   Me.txt_serie = ""
   Me.txt_tipo = ""
   Me.txt_tipo.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 2800
   Left = 3400
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub txt_numero_Change()
   Me.txt_saldo_actual = ""
   Me.txt_saldo_real = ""
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_serie_Change()
   Me.txt_saldo_actual = ""
   Me.txt_saldo_real = ""
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_tipo_Change()
   Me.txt_saldo_actual = ""
   Me.txt_saldo_real = ""
End Sub

Private Sub txt_tipo_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub
