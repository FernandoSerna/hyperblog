VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdetalle_cajas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de la Caja"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   6570
   Begin VB.Frame Frame1 
      Height          =   810
      Left            =   105
      TabIndex        =   6
      Top             =   5280
      Width           =   6300
      Begin VB.CommandButton cmd_aceptar 
         Caption         =   "&Cancelar Caja"
         Height          =   465
         Left            =   1245
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   225
         Width           =   1755
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Salir"
         Height          =   465
         Left            =   3150
         TabIndex        =   7
         Top             =   225
         Width           =   1755
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Detalle de la Caja "
      Height          =   5085
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6285
      Begin VB.TextBox txt_caja 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   870
         Width           =   1515
      End
      Begin VB.TextBox txt_empaque 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   1515
      End
      Begin MSComctlLib.ListView lv_detalle_cajas 
         Height          =   3540
         Left            =   90
         TabIndex        =   1
         Top             =   1395
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   6244
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6526
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Caja:"
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
         Left            =   1230
         TabIndex        =   5
         Top             =   915
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
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
         Left            =   1230
         TabIndex        =   4
         Top             =   375
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmdetalle_cajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_si_trazabilidad As Integer
Dim var_empaque As Double
Dim var_caja As Integer
Dim var_OS As Integer
Dim var_orden_surtido As Double
Dim var_almacen_Destino As String
Dim var_numero_folio As Integer
Dim var_precio As Variant
Private Sub cmd_aceptar_Click()
   Dim var_posible As Boolean
   var_caja = txt_caja
   var_empaque = txt_empaque
   si = MsgBox("¿Deseas cancelar la caja?", vbYesNo, "ATENCION")
   If si = 6 Then
      si = MsgBox("Confirmar la cancelación de la caja", vbYesNo, "ATENCION")
      If si = 6 Then
         If lv_detalle_cajas.ListItems.Count > 0 Then
            rsaux4.Open "select isnull(char_paq_estatus,''), VCHA_ALM_ALMACEN_ID from tb_detalle_cajas where  vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque =  " + Me.txt_empaque + " and inte_paq_caja = " + Me.txt_caja, cnn, adOpenDynamic, adLockOptimistic
            var_estatus = IIf(IsNull(rsaux4(0).Value), "", rsaux4(0).Value)
            VAR_ALMACEN_CAJA = IIf(IsNull(rsaux4!VCHA_ALM_ALMACEN_ID), "", rsaux4!VCHA_ALM_ALMACEN_ID)
            rsaux4.Close
            If var_estatus = "I" Then
               VAR_MOVIMIENTO_CAJA = "CAJA-" + var_empresa + "-" + Me.txt_caja
               var_posible = True
               If var_posible_kanban = 1 Then
                  Set TB_CANCELAR_RES_FUERA_DE_KANBAN = New TB_CANCELAR_RES_FUERA_DE_KANBAN
                  rsaux5.Open "SELECT * FROM tb_fuera_kanban_en_movimientos WHERE VCHA_ALMACEN_ID = '" + VAR_ALMACEN_CAJA + "' AND VCHA_TIPO_MOVIMIENTO_ID = '" + VAR_MOVIMIENTO_CAJA + "' AND BINT_NUMERO_MOVIMIENTO = " + Me.txt_empaque, cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux5.EOF
                        var_inserta = TB_CANCELAR_RES_FUERA_DE_KANBAN.Anadir(VAR_ALMACEN_CAJA, VAR_MOVIMIENTO_CAJA, CStr(Me.txt_empaque), rsaux5!VCHA_ARTICULO_ID, CDbl(rsaux5!FLOA_CANTIDAD), "", "")
                        var_kanban_es_un_kanban = var_kanban_es_un_kanban
                        var_kanban_almacen_id = var_kanban_almacen_id
                        var_kanban_articulo_id = var_kanban_articulo_id
                        var_kanban_exito = var_kanban_exito
                        var_kanban_mensaje = var_kanban_mensaje
                        If var_kanban_exito = "S" Then
                           var_posible = True
                        Else
                           var_posible = False
                        End If
                        rsaux5.MoveNext
                  Wend
                  rsaux5.Close
                  
                  Set TB_CANCELAR_RESERVACION_KANBAN = New TB_CANCELAR_RESERVACION_KANBAN
                  rsaux5.Open "SELECT * FROM TB_KANBANS_EN_MOVIMIENTO WHERE VCHA_ALMACEN_ID = '" + VAR_ALMACEN_CAJA + "' AND VCHA_TIPO_MOVIMIENTO_ID = '" + VAR_MOVIMIENTO_CAJA + "' AND BINT_NUMERO_MOVIMIENTO = " + CStr(Me.txt_empaque), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux5.EOF
                        var_inserta = TB_CANCELAR_RESERVACION_KANBAN.Anadir(VAR_ALMACEN_CAJA, VAR_MOVIMIENTO_CAJA, CDbl(Me.txt_empaque), rsaux5!VCHA_KANBAN_ID, "", "")
                        var_kanban_es_un_kanban = var_kanban_es_un_kanban
                        var_kanban_almacen_id = var_kanban_almacen_id
                        var_kanban_articulo_id = var_kanban_articulo_id
                        var_kanban_exito = var_kanban_exito
                        var_kanban_mensaje = var_kanban_mensaje
                        txt_cantidad_eliminar = 1
                        var_cantidad_eliminar = 1
                        If var_kanban_exito = "S" Then
                           var_posible = True
                        Else
                           var_posible = False
                        End If
                        rsaux5.MoveNext
                  Wend
                  rsaux5.Close
               End If
               If var_posible = True Then
                  rsaux3.Open "select * from tb_detalle_cajas with (nolock) where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque =  " + Me.txt_empaque + " and inte_paq_caja = " + Me.txt_caja, cnn, adOpenDynamic, adLockOptimistic
                  var_almacen_Destino = rsaux3!VCHA_ALM_ALMACEN_ID
                  var_orden_surtido = rsaux3!INTE_ORS_ORDEN_SURTIDO
                  rsaux3.Close
                  rsaux3.Open "select * from tb_det_orden_surtido where INTE_ORS_ORDEN_SURTIDO = " + CStr(var_orden_surtido), cnn, adOpenDynamic, adLockOptimistic
                  var_tipo_pedido = IIf(IsNull(rsaux3!char_ped_tipo), "", rsaux3!char_ped_tipo)
                  rsaux3.Close
                  Set TB_DETALLE_CAJAS_CANCELA = New TB_DETALLE_CAJAS_CANCELA
                  Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
                  rs.Open "SELECT * FROM TB_DETALLE_CAJAS with (nolock) WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND inte_emb_embarque = " + txt_empaque + " AND INTE_PAQ_CAJA = " + txt_caja, cnn, adOpenDynamic, adLockOptimistic
                  While Not rs.EOF
                        var_precio = rs!floa_paq_precio
                        If rsaux4.State = 1 Then
                           rsaux4.Close
                        End If
                        rsaux4.Open "update tb_det_orden_surtido set FLOA_ORS_CANTIDAD_EMPACADA = FLOA_ORS_CANTIDAD_EMPACADA - " + CStr(rs!floa_paq_cantidad) + " where inte_ors_orden_surtido = " + CStr(var_orden_surtido) + " and vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        'var_actualiza = TB_DET_ORDEN_SURTIDO_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_orden_surtido, rs!VCHA_ART_ARTICULO_ID, 0, 0 - rs!floa_paq_cantidad, var_precio, var_tipo_pedido)
                        rs.MoveNext
                  Wend
                  rs.Close
                  If rsaux4.State = 1 Then
                     rsaux4.Close
                  End If
                  rsaux4.Open "update tb_detalle_cajas set char_paq_estatus = 'C' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_emb_embarque = " + Me.txt_empaque + " and inte_paq_caja = " + Me.txt_caja, cnn, adOpenDynamic, adLockOptimistic
                  If var_trazabilidad = 1 Then
                     rsaux4.Open "update TB_TRAZABILIDAD_CODIGOS  set floa_tra_cantidad = 0 where inte_emb_embarque = " + txt_empaque + "  AND INTE_PAQ_CAJA = " + txt_caja, cnn, adOpenDynamic, adLockOptimistic
                  End If
                  'var_actualiza = TB_DETALLE_CAJAS_CANCELA.Anadir(var_empaque, var_caja, var_empresa, var_unidad_organizacional, var_almacen_destino, "C")
                  MsgBox "La caja a sido cancelada", vbOKOnly, "ATENCION"
                  cmd_aceptar.Enabled = False
               Else
                  MsgBox "No se pueden eliminar los kanban", vbOKOnly, "ATENCION"
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub


Private Sub Form_Load()
    Top = 800
    Left = 2500
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call activa_forma(var_activa_forma_detalle_cajas)
End Sub

