VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmdescuentos_pronto_pago_cambios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de descuentos por pronto pago y puntual"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   8175
   Begin VB.Frame frm_lista 
      Height          =   2250
      Left            =   1995
      TabIndex        =   12
      Top             =   -15
      Width           =   5670
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   13
         Top             =   375
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3228
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   14
         Top             =   120
         Width           =   5595
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   150
      TabIndex        =   8
      Top             =   435
      Width           =   7875
      Begin VB.TextBox txt_clave 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   540
         Width           =   1395
      End
      Begin VB.TextBox txt_nombre 
         Height          =   315
         Left            =   2850
         TabIndex        =   4
         Top             =   540
         Width           =   4830
      End
      Begin VB.CheckBox chk_descuento 
         Caption         =   "Aplica descuento"
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   915
         Width           =   2625
      End
      Begin VB.TextBox txt_causa 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   1215
         Width           =   6240
      End
      Begin VB.TextBox txt_periodo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   195
         Width           =   6240
      End
      Begin VB.Label lbl_tipo 
         AutoSize        =   -1  'True
         Caption         =   "Grupo_actual:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Causa de cambio:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1275
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Periodo:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   255
         Width           =   585
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7695
      Picture         =   "frmdescuentos_pronto_pago_cambios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aplicar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmdescuentos_pronto_pago_cambios.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Aplicar Pagos Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   120
      TabIndex        =   7
      Top             =   300
      Width           =   7920
   End
End
Attribute VB_Name = "frmdescuentos_pronto_pago_cambios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_Asignacion As String

Private Sub cmd_aplicar_Click()
   If Me.chk_descuento.Value = 1 Then
      If Trim(Me.txt_clave) <> "" Then
         var_si = MsgBox("¿Desea aplicar el descuento al grupo actual?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar la aplicación del descuento", vbOKOnly, "ATENCION")
            If var_si = 6 Then
               If var_tipo_Asignacion = "A" Then
                  rs.Open "update TB_DESCUENTOS_PAGO_CORRECTO_AUXILIAR set FLOA_DPC_DESCUENTO_ASIGNADO =  2, VCHA_DPC_CAUSA = '" + txt_causa + "' where vcha_gac_grupo_actual_id = '" + txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
                  rs.Open "UPDATE TB_GRUPOSACTUALES SET FLOA_GAC_DESCUENTO_2 = 2  WHERE VCHA_GAC_GRUPO_ACTUAL_ID = '" + Me.txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
               End If
            End If
         End If
      Else
         MsgBox "No se a seleccionado un grupo actual", vbOKOnly, "ATENCION"
      End If
  Else
     MsgBox "No se a indicado el descuento", vbOKOnly, "ATENCION"
  End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Dim fecha_anterior As Date
   Dim dia_anterior As Integer
   Dim mes_anterior As Integer
   Dim año_anterior As Integer
   Dim dia As Integer
   Dim mes As Integer
   Dim año As Integer
   Dim periodo As String
   frm_lista.Visible = False
   var_cadena_seguridad = ""
   Top = 2500
   Left = 2000
   mes = Month(Date)
   año = Year(Date)
   rs.Open "select CHAR_PRI_TIPO_AGRUPAMIENTO from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      If Not IsNull(rs(0).Value) Then
         chk_descuento.Enabled = True
         txt_causa.Enabled = True
         txt_clave.Enabled = True
         txt_nombre.Enabled = True
         var_tipo_Asignacion = rs(0).Value
         var_tipo_Asignacion = "A"
         If var_tipo_Asignacion = "A" Then
            lbl_tipo = "Grupo Actual:"
         End If
         If var_tipo_Asignacion = "R" Then
            lbl_tipo = "Grupo Real:"
         End If
         If var_tipo_Asignacion = "T" Then
            lbl_tipo = "Titular:"
         End If
         If var_tipo_Asignacion = "C" Then
            lbl_tipo = "Cliente:"
         End If
      Else
         chk_descuento.Enabled = False
         txt_causa.Enabled = False
         txt_clave.Enabled = False
         txt_nombre.Enabled = False
      End If
   Else
      chk_descuento.Enabled = False
      txt_causa.Enabled = False
      txt_clave.Enabled = False
      txt_nombre.Enabled = False
   End If
   rs.Close
   If mes = 1 Then
      mes_anterior = 12
      año_anterior = año - 1
   Else
      mes_anterior = mes - 1
      año_anterior = año
   End If
   If mes_anterior = 1 Or mes_anterior = 3 Or mes_anterior = 5 Or mes_anterior = 7 Or mes_anterior = 8 Or mes_anterior = 10 Or mes_anterior = 12 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      fecha_fin = CDate("31/" + Str(mes_anterior) + "/" + Str(año_anterior))
      dia_anterior = 31
   End If
   If mes_anterior = 4 Or mes_anterior = 6 Or mes_anterior = 9 Or mes_anterior = 11 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      fecha_fin = CDate("30/" + Str(mes_anterior) + "/" + Str(año_anterior))
      dia_anterior = 30
   End If
   
   If mes_anterior = 2 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      If año_anterior = 2004 Or año_anterior = 2008 Or año_anterior = 2012 Or año_anterior = 2016 Or año_anterior = 2020 Or año_anterior = 2024 Then
         fecha_fin = CDate("29/" + Str(mes_anterior) + "/" + Str(año_anterior))
         dia_anterior = 29
      Else
         fecha_fin = CDate("28/" + Str(mes_anterior) + "/" + Str(año_anterior))
         dia_anterior = 28
      End If
   End If
   
   mes = mes_anterior
   año = año_anterior
  
   If mes = 1 Then
      periodo = "Enero"
   End If
   If mes = 2 Then
      periodo = "Febrero"
   End If
   If mes = 3 Then
      periodo = "Marzo"
   End If
   If mes = 4 Then
      periodo = "Abril"
   End If
   If mes = 5 Then
      periodo = "Mayo"
   End If
   If mes = 6 Then
      periodo = "Junio"
   End If
   If mes = 7 Then
      periodo = "Julio"
   End If
   If mes = 8 Then
      periodo = "Agosto"
   End If
   If mes = 9 Then
      periodo = "Septiembre"
   End If
   If mes = 10 Then
      periodo = "Octubre"
   End If
   If mes = 11 Then
      periodo = "Noviembre"
   End If
   If mes = 12 Then
      periodo = "Diciembre"
   End If
   txt_periodo = "1 de " + periodo + " al " + Str(dia_anterior) + " de " + periodo + " del " + Str(año)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_descuentos_pronto_pago_cambios)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         Me.txt_clave = lv_lista.selectedItem
         Me.txt_nombre = lv_lista.selectedItem.SubItems(1)
      End If
      Me.txt_clave.SetFocus
   End If
   If KeyAscii = 27 Then
      lv_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_causa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Me.cmd_aplicar.SetFocus
   End If
End Sub

Private Sub txt_clave_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_gruposactuales order by vcha_gac_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_gac_grupo_actual_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "GRUPOS ACTUALES"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_LostFocus()
   If Trim(txt_clave) <> "" Then
      If var_tipo_Asignacion = "A" Then
         rs.Open "select * from VW_DESCUENTOS_PAGO_CORRECTO_AUXILIAR where vcha_gac_grupo_actual_id = '" + txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_nombre = rs!vcha_gac_nombre
            txt_causa = rs!vcha_dpc_causa
            If Not IsNull(rs!FLOA_DPC_DESCUENTO_ASIGNADO) Then
               If rs!FLOA_DPC_DESCUENTO_ASIGNADO > 0 Then
                  chk_descuento.Enabled = False
                  chk_descuento = 1
               Else
                  chk_descuento = 0
                   chk_descuento.Enabled = True
               End If
            Else
               chk_descuento = 0
               chk_descuento.Enabled = True
            End If
         Else
            MsgBox "El grupo no se encuentra dentro de los descuentos por pronto pago", vbOKOnly, "ATENCION"
            chk_descuento = 0
            Me.txt_clave = ""
            txt_nombre = ""
            txt_causa = ""
         End If
         rs.Close
      End If

      If var_tipo_Asignacion = "R" Then
         rs.Open "select * from VW_DESCUENTOS_PAGO_CORRECTO_AUXILIAR where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_gre_grupo_real_id = '" + txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_nombre = rs!vcha_gre_nombre
            txt_causa = rs!vcha_dpc_causa
            If Not IsNull(rs!FLOA_DPC_DESCUENTO_ASIGNADO) Then
               If rs!FLOA_DPC_DESCUENTO_ASIGNADO > 0 Then
                  chk_descuento.Enabled = False
                  chk_descuento = 1
               Else
                  chk_descuento = 0
                  chk_descuento.Enabled = True
               End If
            Else
               chk_descuento = 0
               chk_descuento.Enabled = True
            End If
         Else
            chk_descuento = 0
            txt_nombre = ""
            txt_causa = ""
         End If
         rs.Close
      End If

      If var_tipo_Asignacion = "T" Then
         rs.Open "select * from VW_DESCUENTOS_PAGO_CORRECTO_AUXILIAR where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_gac_grupo_actual_id = '" + txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_nombre = rs!vcha_tit_nombre
            txt_causa = rs!vcha_dpc_causa
            If Not IsNull(rs!FLOA_DPC_DESCUENTO_ASIGNADO) Then
               If rs!FLOA_DPC_DESCUENTO_ASIGNADO > 0 Then
                  chk_descuento.Enabled = False
                  chk_descuento = 1
                Else
                  chk_descuento.Enabled = True
                  chk_descuento = 0
               End If
             Else
               chk_descuento.Enabled = True
               chk_descuento = 0
            End If
         Else
            chk_descuento = 0
            txt_nombre = ""
            txt_causa = ""
         End If
         rs.Close
      End If

      If var_tipo_Asignacion = "C" Then
         rs.Open "select * from VW_DESCUENTOS_PAGO_CORRECTO_AUXILIAR where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_nombre = rs!vcha_cli_nombre
            txt_causa = rs!vcha_dpc_causa
            If Not IsNull(rs!FLOA_DPC_DESCUENTO_ASIGNADO) Then
               If rs!FLOA_DPC_DESCUENTO_ASIGNADO > 0 Then
                  chk_descuento.Enabled = False
                  chk_descuento = 1
               Else
                  chk_descuento.Enabled = True
                  chk_descuento = 0
               End If
            Else
               chk_descuento.Enabled = True
               chk_descuento = 0
            End If
         Else
            chk_descuento = 0
            txt_nombre = ""
            txt_causa = ""
         End If
         rs.Close
      End If
  End If
End Sub

Private Sub txt_nombre_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_gruposactuales where order by vcha_gac_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_gac_grupo_actual_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "GRUPOS ACTUALES"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_periodo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub
