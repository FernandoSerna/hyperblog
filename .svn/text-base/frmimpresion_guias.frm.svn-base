VERSION 5.00
Begin VB.Form frmimpresion_guias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión de guias"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmimpresion_guias.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5685
      Picture         =   "frmimpresion_guias.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   75
      Left            =   0
      TabIndex        =   18
      Top             =   345
      Width           =   6105
   End
   Begin VB.Frame Frame1 
      Height          =   2625
      Left            =   120
      TabIndex        =   10
      Top             =   435
      Width           =   5895
      Begin VB.TextBox txt_clave_cliente 
         Height          =   360
         Left            =   1080
         TabIndex        =   2
         Top             =   600
         Width           =   1425
      End
      Begin VB.TextBox txt_valor 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1080
         TabIndex        =   7
         Top             =   2160
         Width           =   1785
      End
      Begin VB.TextBox txt_fecha 
         Height          =   360
         Left            =   3570
         TabIndex        =   1
         Top             =   195
         Width           =   2235
      End
      Begin VB.TextBox txt_cajas 
         Height          =   360
         Left            =   1080
         TabIndex        =   6
         Top             =   1770
         Width           =   4725
      End
      Begin VB.TextBox txt_guia 
         Height          =   360
         Left            =   1080
         TabIndex        =   5
         Top             =   1380
         Width           =   3015
      End
      Begin VB.TextBox txt_paqueteria 
         Height          =   360
         Left            =   1080
         TabIndex        =   4
         Top             =   990
         Width           =   3030
      End
      Begin VB.TextBox txt_cliente 
         Height          =   360
         Left            =   2535
         TabIndex        =   3
         Top             =   600
         Width           =   3270
      End
      Begin VB.TextBox txt_embarque 
         Height          =   360
         Left            =   1080
         TabIndex        =   0
         Top             =   210
         Width           =   1425
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Left            =   210
         TabIndex        =   17
         Top             =   2250
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   2985
         TabIndex        =   16
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cajas"
         Height          =   195
         Left            =   210
         TabIndex        =   15
         Top             =   1860
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Guia:"
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   1470
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Paqueteria:"
         Height          =   195
         Left            =   210
         TabIndex        =   13
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   690
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   210
         TabIndex        =   11
         Top             =   240
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmimpresion_guias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_imprimir_Click()
   If Me.txt_embarque <> "" Then
      If Me.txt_cliente <> "" Then
         If Me.txt_paqueteria <> "" Then
            If Me.txt_guia <> "" Then
               If Me.txt_cajas <> "" Then
                  var_si = MsgBox("SE VA A IMPRIMIR LA GUIA " + Me.txt_guia + " DE LA PAQUETERIA " + Me.txt_paqueteria, vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     var_si = MsgBox("¿La impresora esta lista?", vbYesNo, "ATENCION")
                     If var_si = 6 Then
                        If Me.txt_paqueteria = "MULTIPACK" Or Me.txt_paqueteria = "MULTIPACK SIN COSTO" Then
                           Open (App.Path & "\MUPA_GUIA_" + Trim(Me.txt_guia) + ".bat") For Output As #2
                           Open (App.Path & "\MUPA_GUIA_" + Trim(Me.txt_guia) + ".txt") For Output As #1
                           Print #1, Chr(15) + Chr(27) + Chr(64)
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, Spc(80); "BLANCOS"
                           Print #1, Spc(80); Me.txt_valor
                           If rs.State = 1 Then
                              rs.Close
                           End If
                           rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + Me.txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, Me.txt_cliente
                           var_colonia = Mid(IIf(IsNull(rs!VCHA_COL_NOMBRE), "", rs!VCHA_COL_NOMBRE), 1, 22)
                           Print #1, Spc(10); IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION)
                           For var_j = Len(Trim(var_colonia)) To 33
                               var_colonia = var_colonia + " "
                           Next var_j
                           var_colonia = var_colonia + Trim(Mid(IIf(IsNull(rs!VCHA_CIU_NOMBRE), "", rs!VCHA_CIU_NOMBRE), 1, 15))
                           Print #1, Spc(5); var_colonia
                           
                           var_estado = Mid(IIf(IsNull(rs!VCHA_EST_NOMBRE), "", rs!VCHA_EST_NOMBRE), 1, 15)
                           For var_j = Len(Trim(var_estado)) To 33
                               var_estado = var_estado + " "
                           Next var_j
                           var_estado = var_estado + Trim(Mid(IIf(IsNull(rs!VCHA_MUN_NOMBRE), "", rs!VCHA_MUN_NOMBRE), 1, 15))
                           Print #1, Spc(5); var_estado
                           var_pais = Mid(IIf(IsNull(rs!VCHA_PAI_NOMBRE), "", rs!VCHA_PAI_NOMBRE), 1, 30)
                           For var_j = Len(Trim(var_pais)) To 33
                               var_pais = var_pais + " "
                           Next var_j
                           var_pais = var_pais + Trim(IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP))
                           Print #1, Spc(5); var_pais
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Close #1
                           Print #2, "copy " + App.Path + "\MUPA_GUIA_" + Trim(Me.txt_guia) + ".txt lpt1"
                           Close #2
                           var_Archivo = App.Path & "\MUPA_GUIA_" + Trim(Me.txt_guia) + ".bat"
                           x = Shell(var_Archivo, vbHide)
                           rs.Close
                        End If
                     Else
                        MsgBox "Se a cancelado la impresión de las etiquetas", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "Se a cancelado la impresión de las etiquetas", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "El embarque no contiene cajas", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No existe la guia", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No existe la paqueteria", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No existe el cliente", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un embarque", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 2000
   Left = 3200
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_cajas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      pro_enfoque (KeyAscii)
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_clave_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      pro_enfoque (KeyAscii)
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      pro_enfoque (KeyAscii)
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_embarque_LostFocus()
   If Me.txt_embarque <> "" Then
      If IsNumeric(Me.txt_embarque) Then
         var_cadena = "SELECT TOP 100 PERCENT dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_EMBARQUES.DTIM_EMB_FECHA_FINAL, SUM(dbo.TB_DETALLE_CAJAS.FLOA_PAQ_CANTIDAD * (((dbo.TB_DETALLE_CAJAS.FLOA_PAQ_PRECIO * (1 + dbo.TB_TIPOPEDIDOS.FLOA_TPE_IVA / 100)) * (1 - dbo.TB_ENC_ORDEN_SURTIDO.FLOA_ORS_DESCUENTO_1 / 100)) * (1 - dbo.TB_ENC_ORDEN_SURTIDO.FLOA_ORS_DESCUENTO_2 / 100))) AS IMPORTE, dbo.TB_ENC_ORDEN_SURTIDO.VCHA_PAQ_CLAVE_ID, dbo.TB_ENC_ORDEN_SURTIDO.VCHA_PAQ_NOMBRE, dbo.TB_ENC_ORDEN_SURTIDO.VCHA_PAQ_GUIA, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID , dbo.TB_CLIENTES.VCHA_CLI_NOMBRE FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_ENCABEZADO_EMBARQUES ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE INNER JOIN dbo.TB_DETALLE_CAJAS ON"
         var_cadena = var_cadena + " dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_CAJAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE = dbo.TB_DETALLE_CAJAS.INTE_EMB_EMBARQUE INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO ON                       dbo.TB_DETALLE_CAJAS.INTE_ORS_ORDEN_SURTIDO = dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO INNER JOIN dbo.TB_TIPOPEDIDOS ON dbo.TB_ENC_ORDEN_SURTIDO.CHAR_TPE_TIPO_PEDIDO_ID = dbo.TB_TIPOPEDIDOS.CHAR_TPE_TIPO_PEDIDO_ID INNER JOIN dbo.TB_AGENTES ON dbo.TB_ENCABEZADO_EMBARQUES.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENC_ORDEN_SURTIDO.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID GROUP BY dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_EMBARQUES.DTIM_EMB_FECHA_FINAL, dbo.TB_ENC_ORDEN_SURTIDO.VCHA_PAQ_CLAVE_ID, dbo.TB_ENC_ORDEN_SURTIDO.VCHA_PAQ_NOMBRE, "
         var_cadena = var_cadena + " dbo.TB_ENC_ORDEN_SURTIDO.VCHA_PAQ_GUIA, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID , dbo.TB_CLIENTES.VCHA_CLI_NOMBRE,dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMP_EMPRESA_ID, dbo.TB_DETALLE_CAJAS.CHAR_PAQ_ESTATUS Having (dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque + ") AND (dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_DETALLE_CAJAS.CHAR_PAQ_ESTATUS = 'S') ORDER BY dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE       "
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_fecha = IIf(IsNull(rs!DTIM_EMB_FECHA_FINAL), "", rs!DTIM_EMB_FECHA_FINAL)
            Me.txt_clave_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
            Me.txt_cliente = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
            Me.txt_paqueteria = IIf(IsNull(rs!vcha_paq_nombre), "", rs!vcha_paq_nombre)
            Me.txt_guia = IIf(IsNull(rs!vcha_paq_guia), "", rs!vcha_paq_guia)
            Me.txt_cajas = ""
            Me.txt_valor = Format(IIf(IsNull(rs!Importe), 0, rs!Importe), "###,###,##0.00")
            var_cadena = "SELECT  COUNT(*) AS CANTIDAD, VCHA_EMP_EMPRESA_ID, INTE_EMB_EMBARQUE, VCHA_CAJ_NOMBRE From dbo.VW_PAQUETERIA_NUMERO_CAJAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque + " GROUP BY VCHA_EMP_EMPRESA_ID, INTE_EMB_EMBARQUE, VCHA_CAJ_NOMBRE "
            rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            Me.txt_cajas = ""
            If Not rsaux.EOF Then
               While Not rsaux.EOF
                     If Me.txt_cajas = "" Then
                        Me.txt_cajas = CStr(rsaux!Cantidad) + " " + rsaux!VCHA_CAJ_NOMBRE
                     Else
                        Me.txt_cajas = Me.txt_cajas + ", " + CStr(rsaux!Cantidad) + " " + rsaux!VCHA_CAJ_NOMBRE
                     End If
                     rsaux.MoveNext
               Wend
            Else
               Me.txt_cajas = ""
            End If
            rsaux.Close
         Else
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
            Me.txt_fecha = ""
            Me.txt_clave_cliente = ""
            Me.txt_cliente = ""
            Me.txt_paqueteria = ""
            Me.txt_guia = ""
            Me.txt_cajas = ""
            Me.txt_valor = ""
         End If
         rs.Close
      Else
         MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
         Me.txt_fecha = ""
         Me.txt_clave_cliente = ""
         Me.txt_cliente = ""
         Me.txt_paqueteria = ""
         Me.txt_guia = ""
         Me.txt_cajas = ""
         Me.txt_valor = ""
      End If
   Else
      Me.txt_fecha = ""
      Me.txt_clave_cliente = ""
      Me.txt_cliente = ""
      Me.txt_paqueteria = ""
      Me.txt_guia = ""
      Me.txt_cajas = ""
      Me.txt_valor = ""
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      pro_enfoque (KeyAscii)
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_guia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      pro_enfoque (KeyAscii)
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_paqueteria_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      pro_enfoque (KeyAscii)
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_valor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      pro_enfoque (KeyAscii)
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub
