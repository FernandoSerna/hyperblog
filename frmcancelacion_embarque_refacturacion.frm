VERSION 5.00
Begin VB.Form frmcancelacion_embarque_refacturacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelacion de embarque para re-facturación"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6900
      Picture         =   "frmcancelacion_embarque_refacturacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmcancelacion_embarque_refacturacion.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   45
      TabIndex        =   12
      Top             =   360
      Width           =   7230
   End
   Begin VB.Frame Frame1 
      Caption         =   " Embarque "
      Height          =   2340
      Left            =   135
      TabIndex        =   5
      Top             =   465
      Width           =   7095
      Begin VB.TextBox txt_estatus 
         Height          =   350
         Left            =   1425
         TabIndex        =   4
         Top             =   1845
         Width           =   315
      End
      Begin VB.TextBox txt_usuario 
         Height          =   350
         Left            =   1425
         TabIndex        =   3
         Top             =   1470
         Width           =   5445
      End
      Begin VB.TextBox txt_fecha 
         Height          =   350
         Left            =   1425
         TabIndex        =   2
         Top             =   1095
         Width           =   2070
      End
      Begin VB.TextBox txt_agente 
         Height          =   350
         Left            =   1425
         TabIndex        =   1
         Top             =   720
         Width           =   5445
      End
      Begin VB.TextBox txt_embarque 
         Height          =   350
         Left            =   1425
         TabIndex        =   0
         Top             =   345
         Width           =   1500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estatus:"
         Height          =   195
         Left            =   210
         TabIndex        =   11
         Top             =   1920
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   210
         TabIndex        =   10
         Top             =   1545
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha creación:"
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   1170
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   795
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   423
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmcancelacion_embarque_refacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
   Dim var_fecha_embarque As String
   Dim var_fecha_hoy As String
   If Trim(Me.txt_embarque) <> "" Then
      If Trim(Me.txt_estatus) = "F" Then
         var_si = MsgBox("Deseas cancelar el embarque para refacturarlo", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rsaux10.Open "select getdate() from tb_lineas", cnn, adOpenDynamic, adLockOptimistic
            var_dia = CStr(Day(rsaux10(0).Value))
            var_mes = CStr(Month(rsaux10(0).Value))
            var_año = CStr(Year(rsaux10(0).Value))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_hoy = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
   
            var_dia = CStr(Day(CDate(Me.txt_fecha)))
            var_mes = CStr(Month(CDate(Me.txt_fecha)))
            var_año = CStr(Year(CDate(Me.txt_fecha)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_embarque = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            rsaux10.Close
            
            
            If var_fecha_hoy = var_fecha_embarque Then
               var_si = MsgBox("Confirmar la cancelación del embarque", vbYesNo, "ATENCION")
               If var_si = 6 Then
               
                  rs.Open "SELECT * FROM VW_DETALLE_EMBARQUES_FACTURACION WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
                  If rs!VCHA_MOV_MOVIMIENTO_ID = "FT" Then
                     MsgBox "Los embarques de facturación de tiendas no se pueden cancelar con este modulo", vbOKOnly, "ATENCION"
                  Else
                     While Not rs.EOF
                           'MsgBox rs!vcha_ser_serie_id + CStr((rs!inte_Car_numero))
                           If var_empresa = "03" Then
                              rsaux.Open "UPDATE TB_sALDOS SET FLOA_SAL_IMPORTE = 100000000 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_cAR_NUMERO = " + CStr(rs!inte_car_numero) + " AND VCHA_SER_sERIE_ID = '" + rs!vcha_ser_Serie_id + "' AND VCHA_cAR_DOCUMENTO = 'FA'", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux.Open "UPDATE TB_ENCABEZADO_cARTERA SET CHAR_CAR_ESTATUS = 'C' WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_cAR_NUMERO = " + CStr(rs!inte_car_numero) + " AND VCHA_SER_sERIE_ID = '" + rs!vcha_ser_Serie_id + "' AND VCHA_cAR_DOCUMENTO = 'FA'", cnn, adOpenDynamic, adLockOptimistic
                           If var_trazabilidad = 1000 Then
                              If cnn_trazabilidad.State = 0 Then
                                 cnn_trazabilidad.Open
                              End If
                       
                              rsaux10.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                              var_nombre_unidad = ""
                              If Not rsaux10.EOF Then
                                 var_nombre_unidad = IIf(IsNull(rsaux10!VCHA_UOR_NOMBRE), "", rsaux10!VCHA_UOR_NOMBRE)
                              End If
                              rsaux10.Close
                                 
                              ndo.organizacion = var_nombre_unidad
                              ndo.eventoNumero = CStr(rs!inte_car_numero)
                              If ndo.cancelarFactura(cnn_trazabilidad) Then
                                   
                              Else
                                 MsgBox "No se pudo ejecutar la trazabilidad", vbOKOnly, "ATENCION"
                              End If
                              cnn_trazabilidad.Close
                           End If
                           rs.MoveNext
                     Wend
                     rsaux.Open "UPDATE TB_ENCABEZADO_EMBARQUES SET CHAR_EMB_ESTATUS = 'I' WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_eMB_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
                     Me.txt_agente = ""
                     Me.txt_embarque = ""
                     Me.txt_estatus = ""
                     Me.txt_fecha = ""
                     Me.txt_usuario = ""
                     Me.txt_embarque = ""
                     
                     MsgBox "Se a terminado de cancelar el embarque", vbOKOnly, "ATENCION"
                  End If
                  rs.Close
               End If
            Else
               MsgBox "El embarque ya no puede ser cancelado ya que fue hecho otro dia", vbOKOnly, "ATENCION"
            End If
         End If
      Else
         MsgBox "El embarque no puede ser cancelado ya que no se a facturado", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Debe de indicarse un número de embarque", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 1800
   Left = 2200
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_embarque_LostFocus()
   If Trim(Me.txt_embarque) <> "" Then
      If IsNumeric(Me.txt_embarque) Then
         rs.Open "SELECT * FROM VW_EMBARQUES WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            Me.txt_estatus = IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", rs!CHAR_EMB_ESTATUS)
            Me.txt_usuario = IIf(IsNull(rs!VCHA_USU_NOMBRE), "", rs!VCHA_USU_NOMBRE) + " " + IIf(IsNull(rs!vcha_usu_apellidos), "", rs!vcha_usu_apellidos)
            Me.txt_fecha = IIf(IsNull(rs!DTIM_EMB_FECHA_FINAL), "", rs!DTIM_EMB_FECHA_FINAL)
         Else
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
            Me.txt_agente = ""
            Me.txt_embarque = ""
            Me.txt_estatus = ""
            Me.txt_fecha = ""
            Me.txt_usuario = ""
         End If
         rs.Close
      Else
         MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
         Me.txt_agente = ""
         Me.txt_embarque = ""
         Me.txt_estatus = ""
         Me.txt_fecha = ""
         Me.txt_usuario = ""
      End If
   Else
      Me.txt_agente = ""
      Me.txt_embarque = ""
      Me.txt_estatus = ""
      Me.txt_fecha = ""
      Me.txt_usuario = ""
   End If
End Sub

Private Sub txt_estatus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_usuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub
