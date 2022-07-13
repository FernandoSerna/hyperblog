VERSION 5.00
Begin VB.Form frmeliminacion_relacion_cobranza 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eliminación de relación de cobranza"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4215
      Picture         =   "frmeliminacion_relacion_cobranza.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmeliminacion_relacion_cobranza.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancelar Alt + C"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   90
      TabIndex        =   10
      Top             =   300
      Width           =   4500
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos de la Relación "
      Height          =   1395
      Left            =   105
      TabIndex        =   7
      Top             =   1230
      Width           =   4470
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   675
         TabIndex        =   4
         Top             =   630
         Width           =   1575
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   1470
         TabIndex        =   3
         Top             =   285
         Width           =   2925
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   675
         TabIndex        =   2
         Top             =   285
         Width           =   780
      End
      Begin VB.Label lbl_estatus 
         Alignment       =   2  'Center
         Caption         =   "APLICADA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   11
         Top             =   1035
         Width           =   4245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   75
         TabIndex        =   9
         Top             =   690
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   60
         TabIndex        =   8
         Top             =   345
         Width           =   555
      End
   End
   Begin VB.Frame fr 
      Caption         =   " Folio de la relación "
      Height          =   705
      Left            =   105
      TabIndex        =   0
      Top             =   465
      Width           =   4470
      Begin VB.TextBox txt_relacion 
         Height          =   315
         Left            =   1290
         TabIndex        =   1
         Top             =   240
         Width           =   2040
      End
   End
End
Attribute VB_Name = "frmeliminacion_relacion_cobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancelar_Click()
   Dim var_posible_eliminar As Boolean
   If Trim(Me.txt_relacion) <> "" Then
      rs.Open "select * from tb_Relacion_cobranza where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + Me.txt_relacion + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux1.Open "select * from tb_Relacion_cobranza where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + Me.txt_relacion + "' AND CHAR_RCO_APLICADA = '*'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            var_posible_eliminar = False
         Else
            var_posible_eliminar = True
         End If
         rsaux1.Close
         rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + IIf(IsNull(rs!vcha_age_agente_id), "", rs!vcha_age_agente_id) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            Me.txt_agente = rs!vcha_age_agente_id
            Me.txt_nombre_agente = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
         Else
            Me.txt_agente = ""
            Me.txt_nombre_agente = ""
         End If
         rsaux.Close
         rsaux.Open "select sum(floa_rco_importe) from tb_relacion_cobranza where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + Me.txt_relacion + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            Me.txt_importe = Format(IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value), "###,###,##0.00")
         Else
            Me.txt_importe = Format(0, "###,###,##0.00")
         End If
         rsaux.Close
         If var_posible_eliminar = True Then
            var_si = MsgBox("¿Desea eliminar la relación de cobranza", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_si = MsgBox("Confirmar la eliminación de la relación de cobranza", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  rsaux1.Open "DELETE from tb_Relacion_cobranza where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + Me.txt_relacion + "' AND CHAR_RCO_APLICADA <> '*'", cnn, adOpenDynamic, adLockOptimistic
                  MsgBox "Se a eliminado la relación de cobranza", vbOKOnly, "ATENCION"
               End If
            End If
         Else
            MsgBox "La relación de cobranza ya no puede ser eliminada ya que fue aplicada", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "La relación de cobranza no existe", vbOKOnly, "ATENCION"
         Me.txt_agente = ""
         Me.txt_importe = ""
         Me.txt_nombre_agente = ""
         Me.lbl_estatus = ""
      End If
      rs.Close
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 2500
   Left = 3500
   Me.lbl_estatus = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_importe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_cancelar.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_relacion_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_relacion_LostFocus()
   If Trim(Me.txt_relacion) <> "" Then
      rs.Open "select * from tb_Relacion_cobranza where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + Me.txt_relacion + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux1.Open "select * from tb_Relacion_cobranza where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + Me.txt_relacion + "' AND CHAR_RCO_APLICADA = '*'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            Me.lbl_estatus = "APLICADA"
         Else
            Me.lbl_estatus = ""
         End If
         rsaux1.Close
         rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + IIf(IsNull(rs!vcha_age_agente_id), "", rs!vcha_age_agente_id) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            Me.txt_agente = rs!vcha_age_agente_id
            Me.txt_nombre_agente = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
         Else
            Me.txt_agente = ""
            Me.txt_nombre_agente = ""
         End If
         rsaux.Close
         rsaux.Open "select sum(floa_rco_importe) from tb_relacion_cobranza where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + Me.txt_relacion + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            Me.txt_importe = Format(IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value), "###,###,##0.00")
         Else
            Me.txt_importe = Format(0, "###,###,##0.00")
         End If
         rsaux.Close
      Else
         MsgBox "La relación de cobranza no existe", vbOKOnly, "ATENCION"
         Me.txt_agente = ""
         Me.txt_importe = ""
         Me.txt_nombre_agente = ""
         Me.lbl_estatus = ""
      End If
      rs.Close
   End If
End Sub
