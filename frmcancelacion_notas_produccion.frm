VERSION 5.00
Begin VB.Form frmcancelacion_notas_produccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelación de notas de producción"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3855
      Picture         =   "frmcancelacion_notas_produccion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancelar Factura"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmcancelacion_notas_produccion.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   15
      TabIndex        =   3
      Top             =   360
      Width           =   4230
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   105
      TabIndex        =   0
      Top             =   390
      Width           =   4080
      Begin VB.TextBox txt_nota 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1695
         TabIndex        =   1
         Top             =   285
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nota de producción:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   435
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmcancelacion_notas_produccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
    
End Sub

Private Sub cmd_aceptar_Click()
   Dim var_fecha_actual As String
   Dim var_fecha_movimiento As String
   If IsNumeric(Me.txt_nota) Then
      rs.Open "select * from tb_encabezado_movimientos where vcha_Emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'VDIP' and vcha_Emo_Referencia = '" + Me.txt_nota + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_dia = CStr(Day(rs!dtim_Emo_fecha))
         var_mes = CStr(Month(rs!dtim_Emo_fecha))
         var_año = CStr(Year(rs!dtim_Emo_fecha))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         var_fecha_movimiento = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
         var_dia = CStr(Day(Date))
         var_mes = CStr(Month(Date))
         var_año = CStr(Year(Date))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         var_fecha_movimiento = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
         var_cancelado = ""
         rsaux.Open "select distinct inte_Car_numero, vcha_Ser_Serie_id from tb_Salidas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'VDIP' and inte_sal_numero = " + CStr(rs!INTE_eMO_NUMERO), cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               rsaux1.Open "select isnull(char_car_estatus,'')  from tb_encabezado_Cartera where inte_car_numero = " + CStr(IIf(IsNull(rsaux!INTE_cAR_NUMERO), 0, rsaux!INTE_cAR_NUMERO)) + " and vcha_Ser_serie_id = '" + IIf(IsNull(rsaux!VCHA_sER_sERIE_ID), "", rsaux!VCHA_sER_sERIE_ID) + "' and vcha_Car_documento = 'FA' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux1(0).Value <> "C" Then
                  var_cancelado = "I"
               End If
               rsaux1.Close
               rsaux.MoveNext
         Wend
         rsaux.Close
         If var_cancelado = "I" Then
            MsgBox "Se deben de cancelar primero las facturas", vbOKOnly, "ATENCION"
         Else
            rsaux.Open "SELECT * FROM TB_eNCABEZADO_MOVIMIENTOS where vcha_Emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'EPVD' and vcha_Emo_Referencia = '" + Me.txt_nota + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               rsaux1.Open "SELECT * FROM TB_ENTRADAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = 'EPVD' AND INTE_ENT_NUMERO = " + CStr(rsaux!INTE_eMO_NUMERO), cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux1.EOF
                     'rsaux2.Open "DELETE FROM TB_ENTRADAS WHERE INTE_ENT_CONSECUTIVO_TABLA = " + CStr(rsaux1!INTE_ENT_CONSECUTIVO_TABLA), cnn, adOpenDynamic, adLockOptimistic
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
            End If
            rsaux.Close
            rsaux1.Open "SELECT * FROM TB_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = 'VDIP' AND INTE_SAL_NUMERO = " + CStr(rs!INTE_eMO_NUMERO), cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux1.EOF
                  'rsaux2.Open "DELETE FROM TB_sALIDAS WHERE INTE_SAL_CONSECUTIVO_TABLA = " + CStr(rsaux1!INTE_SAL_CONSECUTIVO_TABLA), cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.MoveNext
            Wend
            rsaux1.Close
            'MsgBox "UPDATE TB_eNCABEZADO_MOVIMIENTOS SET VCHA_EMO_REFERENCIA = 'CANCELADA NOTA '" + Me.txt_nota + " where vcha_Emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'EPVD' and vcha_Emo_Referencia = '" + Me.txt_nota + "'"
            rsaux1.Open "UPDATE TB_eNCABEZADO_MOVIMIENTOS SET VCHA_EMO_REFERENCIA = 'CANCELADA NOTA " + Me.txt_nota + "' where vcha_Emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'EPVD' and vcha_Emo_Referencia = '" + Me.txt_nota + "'", cnn, adOpenDynamic, adLockOptimistic
            rsaux1.Open "UPDATE TB_eNCABEZADO_MOVIMIENTOS SET VCHA_EMO_REFERENCIA = 'CANCELADA NOTA " + Me.txt_nota + "' where vcha_Emp_Empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'VDIP' and vcha_Emo_Referencia = '" + Me.txt_nota + "'", cnn, adOpenDynamic, adLockOptimistic
            MsgBox "Se a cancelado los movimientos de la nota de producción " + Me.txt_nota, vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se encontro ningún movimiento con la nota de producción " + Me.txt_nota, vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      MsgBox "Debe de selecciónar una nota de producción", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub
