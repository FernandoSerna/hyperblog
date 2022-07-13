VERSION 5.00
Begin VB.Form frmcompletar_orden_de_surtido_access 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Completar Orden Surtido Access"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Vaciar tabla"
      Height          =   315
      Left            =   780
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5730
      Picture         =   "frmcompletar_orden_de_surtido_access.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   450
      Picture         =   "frmcompletar_orden_de_surtido_access.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Guardar"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmcompletar_orden_de_surtido_access.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Nuevo"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   90
      Left            =   75
      TabIndex        =   5
      Top             =   315
      Width           =   6075
   End
   Begin VB.Frame Frame1 
      Height          =   2955
      Left            =   120
      TabIndex        =   0
      Top             =   450
      Width           =   5985
      Begin VB.TextBox txt_almacen 
         Height          =   315
         Left            =   1470
         TabIndex        =   19
         Top             =   2580
         Width           =   510
      End
      Begin VB.TextBox txt_unidad 
         Height          =   315
         Left            =   1470
         TabIndex        =   10
         Top             =   2250
         Width           =   510
      End
      Begin VB.TextBox txt_Empresa 
         Height          =   315
         Left            =   1470
         TabIndex        =   9
         Top             =   1910
         Width           =   510
      End
      Begin VB.TextBox txt_pedido 
         Height          =   315
         Left            =   1470
         TabIndex        =   8
         Top             =   1570
         Width           =   1065
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   1470
         TabIndex        =   7
         Top             =   1230
         Width           =   1065
      End
      Begin VB.TextBox txt_movimiento 
         Height          =   315
         Left            =   1470
         TabIndex        =   6
         Top             =   890
         Width           =   510
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   1470
         TabIndex        =   4
         Top             =   550
         Width           =   4350
      End
      Begin VB.TextBox txt_orden 
         Height          =   315
         Left            =   1470
         TabIndex        =   2
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Left            =   135
         TabIndex        =   20
         Top             =   2640
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Unidad:"
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   2310
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   1970
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   1630
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   1290
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento:"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   950
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   610
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orden de Surtido:"
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   270
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmcompletar_orden_de_surtido_access"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancelar_pedidos_Click()
   If Me.txt_orden <> "" Then
      If Me.txt_numero <> "" Then
         rs.Open "select * from vw_orden_surtido where inte_ors_orden_surtido = " + Me.txt_orden + " and floa_ors_cantidad_surtir > 0", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  rsaux1.Open "select * from tb_salidas where vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "' and vcha_sal_archivo = '" + Me.txt_Empresa + Me.txt_unidad + Me.txt_almacen + Me.txt_movimiento + Me.txt_numero + "'", cnnaccess, adOpenDynamic, adLockOptimistic
                  If rsaux1.EOF Then
                     Cadena = "insert into tb_salidas (VCHA_SAL_ARCHIVO, INTE_PED_NUMERO, INTE_ORS_ORDEN_SURTIDO, VCHA_EMP_EMPRESA_ID, INTE_SAL_NUMERO,VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ORS_CANTIDAD_SURTIR, FLOA_ORS_CANTIDAD_SURTIDA, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, "
                     Cadena = Cadena + " VCHA_SAL_TIPO, INTE_SAL_CONSECUTIVO) VALUES ('" + Trim(Me.txt_Empresa) + Trim(Me.txt_unidad) + Trim(Me.txt_almacen) + Trim(Me.txt_movimiento) + Trim(Me.txt_numero) + "'," + Trim(txt_pedido) + "," + Me.txt_orden + ",'" + Me.txt_Empresa + "'," + Trim(Me.txt_numero) + ",'" + rs!vcha_Art_articulo_id + "',''," + CStr(rs!FLOA_ORS_CANTIDAD_SURTIR) + ", " + CStr(rs!FLOA_ORS_CANTIDAD_SURTIDA) + ", " + CStr(rs!FLOA_ORS_CANTIDAD_SURTIDA) + "," + CStr(rs!floa_ors_costo) + "," + CStr(rs!floa_ors_precio) + "," + CStr(rs!floa_ors_promocion_1) + "," + CStr(rs!floa_ors_promocion_2) + ",'" + rs!char_ped_tipo + "',0)"
                     rsaux.Open Cadena, cnnaccess, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux1.Close
                  rs.MoveNext
            Wend
         End If
         rs.Close
      End If
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   'rsaux.Open "UPDATE TB_SALIDAS SET FLOA_SAL_cANTIDAD = FLOA_ORS_CANTIDAD_SURTIR", cnnaccess, adOpenDynamic, adLockOptimistic
   rsaux.Open "delete from TB_SALIDAS", cnnaccess, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Form_Load()
   Top = 2500
   Left = 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
       KeyAscii = 0
    Else
      Call pro_enfoque(KeyAscii)
    End If
End Sub

Private Sub txt_almacen_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
       KeyAscii = 0
    Else
      Call pro_enfoque(KeyAscii)
    End If
End Sub

Private Sub txt_empresa_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
       KeyAscii = 0
    Else
      Call pro_enfoque(KeyAscii)
    End If
End Sub

Private Sub txt_movimiento_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
       KeyAscii = 0
    Else
      Call pro_enfoque(KeyAscii)
    End If
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
       KeyAscii = 0
    Else
      Call pro_enfoque(KeyAscii)
    End If
End Sub

Private Sub txt_orden_Change()
   Me.txt_agente = ""
   Me.txt_Empresa = ""
   Me.txt_movimiento = ""
   Me.txt_numero = ""
   Me.txt_pedido = ""
   Me.txt_unidad = ""
End Sub

Private Sub txt_orden_KeyPress(KeyAscii As Integer)
      Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_orden_LostFocus()
   If IsNumeric(Me.txt_orden) Then
      rs.Open "SELECT     dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO, dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.vcha_Alm_almacen_id FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_AGENTES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO ON                      dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO Where (dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = " + Me.txt_orden + ") ", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
         Me.txt_Empresa = IIf(IsNull(rs!VCHA_EMP_EMPRESA_ID), "", rs!VCHA_EMP_EMPRESA_ID)
         Me.txt_unidad = IIf(IsNull(rs!VCHA_UOR_UNIDAD_ID), "", rs!VCHA_UOR_UNIDAD_ID)
         Me.txt_movimiento = IIf(IsNull(rs!VCHA_MOV_MOVIMIENTO_ID), "", rs!VCHA_MOV_MOVIMIENTO_ID)
         Me.txt_numero = IIf(IsNull(rs!inte_emo_numero), "", rs!inte_emo_numero)
         Me.txt_pedido = IIf(IsNull(rs!inte_ped_numero), "", rs!inte_ped_numero)
         Me.txt_almacen = IIf(IsNull(rs!vcha_Alm_almacen_id), "", rs!vcha_Alm_almacen_id)
      Else
         Me.txt_agente = ""
         Me.txt_Empresa = ""
         Me.txt_unidad = ""
         Me.txt_movimiento = ""
         Me.txt_numero = ""
         Me.txt_pedido = ""
         Me.txt_almacen = ""
         MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_agente = ""
      Me.txt_Empresa = ""
      Me.txt_unidad = ""
      Me.txt_movimiento = ""
      Me.txt_numero = ""
      Me.txt_pedido = ""
      Me.txt_almacen = ""
   End If
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
       KeyAscii = 0
    Else
      Call pro_enfoque(KeyAscii)
    End If
End Sub

Private Sub txt_unidad_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
       KeyAscii = 0
    Else
      Call pro_enfoque(KeyAscii)
    End If
End Sub
