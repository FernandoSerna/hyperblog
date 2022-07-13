VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmactualizar_entradas_multibondeados 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actualizacion de entradas de multibondeados"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   765
      Picture         =   "frmactualizar_entradas_multibondeados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Imprimir Movimiento Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmactualizar_entradas_multibondeados.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8640
      Picture         =   "frmactualizar_entradas_multibondeados.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmactualizar_entradas_multibondeados.frx":0886
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Nuevo Pedido Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Caption         =   " Detalle Movimiento "
      Height          =   3990
      Left            =   75
      TabIndex        =   2
      Top             =   1395
      Width           =   8940
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   5475
         TabIndex        =   10
         Top             =   1215
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            TabIndex        =   11
            Top             =   390
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   "Cantidad a actual"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   12
            Top             =   15
            Width           =   2895
         End
      End
      Begin MSComctlLib.ListView lv_entradas 
         Height          =   3645
         Left            =   60
         TabIndex        =   9
         Top             =   240
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   6429
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7735
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Anterior"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Actual"
            Object.Width           =   2646
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Encabezado movimiento "
      Height          =   930
      Left            =   75
      TabIndex        =   1
      Top             =   435
      Width           =   8925
      Begin VB.TextBox txt_fecha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5595
         TabIndex        =   5
         Top             =   315
         Width           =   3270
      End
      Begin VB.TextBox txt_lote 
         Alignment       =   1  'Right Justify
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
         Left            =   2985
         TabIndex        =   4
         Top             =   300
         Width           =   1905
      End
      Begin VB.TextBox txt_numero 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1035
         TabIndex        =   3
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   4965
         TabIndex        =   8
         Top             =   405
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
         Height          =   195
         Index           =   0
         Left            =   2475
         TabIndex        =   7
         Top             =   405
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   405
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   60
      TabIndex        =   0
      Top             =   270
      Width           =   8955
   End
End
Attribute VB_Name = "frmactualizar_entradas_multibondeados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_aceptar_Click()
   If Trim(Me.txt_numero) <> "" Then
      If IsNumeric(Me.txt_numero) Then
         If lv_entradas.ListItems.Count > 0 Then
            var_si = MsgBox("¿Deseas ejecutar los cambios?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_si = MsgBox("Confirmar el cambio del movimiento", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  For var_j = 1 To lv_entradas.ListItems.Count
                      lv_entradas.ListItems.Item(var_j).Selected = True
                      If CDbl(lv_entradas.selectedItem.SubItems(2)) <> CDbl(lv_entradas.selectedItem.SubItems(3)) Then
                         lv_entradas.ListItems.Item(var_j).Selected = True
                         var_diferencia = CDbl(lv_entradas.selectedItem.SubItems(2)) - CDbl(lv_entradas.selectedItem.SubItems(3))
                         cnn.BeginTrans
                         rs.Open "UPDATE TB_ENTRADAS SET FLOA_ENT_CANTIDAD = " + CStr(CDbl(lv_entradas.selectedItem.SubItems(3))) + " WHERE VCHA_EMP_EMPRESA_ID = '16' AND VCHA_UOR_UNIDAD_ID = '09' AND VCHA_ALM_ALMACEN_ID = '28' AND VCHA_MOV_MOVIMIENTO_ID = 'EPTM' AND INTE_ENT_NUMERO = " + Me.txt_numero + " AND VCHA_aRT_ARTICULO_ID = '" + Trim(Me.lv_entradas.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                         rs.Open "UPDATE TB_TEMPORAL_ENTRADAS SET FLOA_ENT_CANTIDAD = " + CStr(lv_entradas.selectedItem.SubItems(3)) + " WHERE VCHA_EMP_EMPRESA_ID = '16' AND VCHA_UOR_UNIDAD_ID = '09' AND VCHA_ALM_ALMACEN_ID = '28' AND VCHA_MOV_MOVIMIENTO_ID = 'EPTM' AND INTE_ENT_NUMERO = " + Me.txt_numero + " AND VCHA_aRT_ARTICULO_ID = '" + Trim(Me.lv_entradas.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                         rs.Open "INSERT INTO TB_BITACORA_CAMBIOS_MOVIMIENTOS_MULTIBONDEADOS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_aRT_ARTICULO_ID,FLOA_BIT_CANTIDAD_ANTERIOR, FLOA_BIT_CANTIDAD_ACTUAL, VCHA_BIT_USUARIO, VCHA_BIT_MAQUINA) VALUES ('16','09','28','EPTM'," + Me.txt_numero + ",'" + Trim(Me.lv_entradas.selectedItem) + "', " + CStr(CDbl(Me.lv_entradas.selectedItem.SubItems(2))) + "," + CStr(CDbl(Me.lv_entradas.selectedItem.SubItems(3))) + ", '" + var_clave_usuario_global + "','" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
                         rs.Open "Update tb_existencias set floa_Exi_cantidad = floa_Exi_cantidad - " + CStr(var_diferencia) + ", floa_Exi_cantidad_2005 = floa_exi_Cantidad_2005 - " + CStr(var_diferencia) + " where vcha_alm_almacen_id = '28' and vcha_Art_articulo_id = '" + lv_entradas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                         cnn.CommitTrans
                         MsgBox "Se a aplicado el cambio", vbOKOnly, "ATENCION"
                         Me.txt_numero = ""
                      End If
                  Next var_j
               Else
                  MsgBox "Se a cancelado el cambio del movimiento", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Se a cancelado el cambio del movimiento", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No existe información por cambiar", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Número de movimiento incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Se debe de indicar un movimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
         Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_ENTRADAS.rpt")
         reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_ENTRADA.VCHA_EMP_EMPRESA_ID} = '16' AND {VW_MOVIMIENTOS_ENTRADA.VCHA_MOV_MOVIMIENTO_ID} = 'EPTM' AND {VW_MOVIMIENTOS_ENTRADA.INTE_EMO_NUMERO} = " + Me.txt_numero + " AND {VW_MOVIMIENTOS_ENTRADA.VCHA_ALM_ALMACEN_ID} = '28' and {VW_MOVIMIENTOS_ENTRADA.VCHA_eMP_EMPRESA_ID} = '16'"
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Movimientos"
         frmvistasprevias.Show 1
         Set reporte = Nothing
End Sub

Private Sub cmd_nuevo_Click()
  Me.txt_fecha = ""
  Me.txt_lote = ""
  Me.txt_numero = ""
  Me.lv_entradas.ListItems.Clear
  Me.txt_numero.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 1000
   Left = 1500
   Me.frm_eliminar.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_entradas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_entradas, ColumnHeader)
End Sub

Private Sub lv_entradas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      Me.txt_cantidad_eliminar = ""
      frm_eliminar.Visible = True
      txt_cantidad_eliminar.SetFocus
   End If
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad_eliminar) Then
         If Me.txt_cantidad_eliminar <= lv_entradas.selectedItem.SubItems(2) Then
            Me.lv_entradas.selectedItem.SubItems(3) = Me.txt_cantidad_eliminar
         Else
            MsgBox "La cantidad a corregir debe de ser menor o igual a la cantidad en el movimiento", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
      End If
      If Me.lv_entradas.ListItems.Count > 0 Then
         Me.lv_entradas.SetFocus
      Else
         Me.frm_eliminar.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      If Me.lv_entradas.ListItems.Count > 0 Then
         Me.lv_entradas.SetFocus
      Else
         Me.frm_eliminar.Visible = False
      End If
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   Me.frm_eliminar.Visible = False
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_lote_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_numero_Change()
   Me.txt_lote = ""
   Me.txt_fecha = ""
   Me.lv_entradas.ListItems.Clear
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_LostFocus()
   If Me.txt_numero <> "" Then
      If IsNumeric(Me.txt_numero) Then
         rs.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emo_numero = " + Me.txt_numero + " and vcha_mov_movimiento_id = 'EPTM'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_movimiento_bloqueado = 0
            If var_movimiento_bloqueado = 0 Then
               var_almacen_destino_tem = rs!VCHA_ALM_ALMACEN_ID
               var_posible = 0
               var_estatus = IIf(IsNull(rs!char_Emo_estatus), "", rs!char_Emo_estatus)
               If var_estatus = "I" Then
                  var_posible = 1
               End If
               If var_posible = 1 Then
                  var_estatus_movimiento = rs!char_Emo_estatus
                  var_almacen_Destino = rs!VCHA_ALM_ALMACEN_ID
                  Me.txt_lote = IIf(IsNull(rs!vcha_Emo_referencia), "", rs!vcha_Emo_referencia)
                  Me.txt_fecha = IIf(IsNull(rs!dtim_emo_fecha), "", rs!dtim_emo_fecha)
                  lv_entradas.ListItems.Clear
                  var_primera_vez = False
                  var_numero_folio = rs!INTE_EMO_NUMERO
                  txt_folio = var_numero_folio
                  rsaux.Open "select * from tb_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and inte_ent_numero = " + Me.txt_numero + " and vcha_mov_movimiento_id = 'EPTM'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     While Not rsaux.EOF
                        rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           Set list_item = lv_entradas.ListItems.Add(, , rsaux!vcha_Art_articulo_id)
                           list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
                           list_item.SubItems(2) = IIf(IsNull(rsaux!floa_ent_Cantidad), "", rsaux!floa_ent_Cantidad)
                           list_item.SubItems(3) = IIf(IsNull(rsaux!floa_ent_Cantidad), "", rsaux!floa_ent_Cantidad)
                           rsaux2.Close
                           rsaux.MoveNext:
                        End If
                     Wend
                  End If
                  rsaux.Close
               Else
                  MsgBox "El movimiento aun no a sido impreso", vbOKOnly, "ATENCION"
               End If
            End If
         Else
            MsgBox "El número de movimiento no existe ", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Número incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      Me.txt_lote = ""
      Me.txt_fecha = ""
      Me.lv_entradas.ListItems.Clear
   End If
End Sub
